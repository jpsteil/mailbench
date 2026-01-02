"""SQLite database for storing accounts, settings, and cached messages."""

import json
import os
import sqlite3
from pathlib import Path


def _is_installed():
    """Check if running as an installed package (pipx/pip) vs development."""
    return 'site-packages' in str(Path(__file__).resolve())


def _get_data_dir():
    """Get the appropriate data directory based on install type."""
    if _is_installed():
        if os.name == 'nt':  # Windows
            base = os.environ.get('APPDATA', Path.home())
            data_dir = Path(base) / 'mailbench'
        else:  # Linux/Mac
            xdg_data = os.environ.get('XDG_DATA_HOME', Path.home() / '.local' / 'share')
            data_dir = Path(xdg_data) / 'mailbench'

        try:
            data_dir.mkdir(parents=True, exist_ok=True)
            return data_dir
        except Exception:
            return Path(__file__).parent
    else:
        return Path(__file__).parent


class Database:
    def __init__(self, db_path=None):
        if db_path is None:
            data_dir = _get_data_dir()
            db_path = data_dir / "mailbench.db"

        self.db_path = db_path
        self._init_db()

    def _get_conn(self):
        return sqlite3.connect(self.db_path)

    def _init_db(self):
        with self._get_conn() as conn:
            # Settings table (key-value store)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS settings (
                    key TEXT PRIMARY KEY,
                    value TEXT
                )
            """)

            # Email accounts
            conn.execute("""
                CREATE TABLE IF NOT EXISTS accounts (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL UNIQUE,
                    email TEXT NOT NULL,
                    server TEXT NOT NULL,
                    username TEXT NOT NULL,
                    password TEXT NOT NULL,
                    auth_type TEXT DEFAULT 'basic',
                    ews_url TEXT,
                    autodiscover INTEGER DEFAULT 0,
                    sync_interval INTEGER DEFAULT 300,
                    is_default INTEGER DEFAULT 0,
                    display_order INTEGER DEFAULT 0,
                    last_sync TIMESTAMP,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)

            # Folder cache
            conn.execute("""
                CREATE TABLE IF NOT EXISTS folders (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    account_id INTEGER NOT NULL,
                    folder_id TEXT NOT NULL,
                    name TEXT NOT NULL,
                    parent_id TEXT,
                    folder_type TEXT,
                    unread_count INTEGER DEFAULT 0,
                    total_count INTEGER DEFAULT 0,
                    sync_state TEXT,
                    FOREIGN KEY (account_id) REFERENCES accounts(id) ON DELETE CASCADE,
                    UNIQUE(account_id, folder_id)
                )
            """)

            # Message cache
            conn.execute("""
                CREATE TABLE IF NOT EXISTS messages (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    account_id INTEGER NOT NULL,
                    folder_id TEXT NOT NULL,
                    item_id TEXT NOT NULL UNIQUE,
                    change_key TEXT,
                    conversation_id TEXT,
                    subject TEXT,
                    sender_name TEXT,
                    sender_email TEXT,
                    recipients TEXT,
                    cc TEXT,
                    date_received TIMESTAMP,
                    date_sent TIMESTAMP,
                    size INTEGER,
                    importance TEXT DEFAULT 'normal',
                    is_read INTEGER DEFAULT 0,
                    has_attachments INTEGER DEFAULT 0,
                    body_preview TEXT,
                    body_type TEXT,
                    body TEXT,
                    categories TEXT,
                    is_flagged INTEGER DEFAULT 0,
                    cached_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (account_id) REFERENCES accounts(id) ON DELETE CASCADE
                )
            """)

            conn.execute("""
                CREATE INDEX IF NOT EXISTS idx_messages_folder
                ON messages(account_id, folder_id, date_received DESC)
            """)

            conn.execute("""
                CREATE INDEX IF NOT EXISTS idx_messages_conversation
                ON messages(conversation_id)
            """)

            # Attachments metadata
            conn.execute("""
                CREATE TABLE IF NOT EXISTS attachments (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    message_id INTEGER NOT NULL,
                    attachment_id TEXT,
                    name TEXT NOT NULL,
                    content_type TEXT,
                    size INTEGER,
                    is_inline INTEGER DEFAULT 0,
                    content_id TEXT,
                    FOREIGN KEY (message_id) REFERENCES messages(id) ON DELETE CASCADE
                )
            """)

            # Contacts cache (Phase 2)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS contacts (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    account_id INTEGER NOT NULL,
                    item_id TEXT NOT NULL,
                    display_name TEXT,
                    email_addresses TEXT,
                    phone_numbers TEXT,
                    company TEXT,
                    job_title TEXT,
                    notes TEXT,
                    cached_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (account_id) REFERENCES accounts(id) ON DELETE CASCADE,
                    UNIQUE(account_id, item_id)
                )
            """)

            # Calendar events cache (Phase 2)
            conn.execute("""
                CREATE TABLE IF NOT EXISTS calendar_events (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    account_id INTEGER NOT NULL,
                    calendar_id TEXT,
                    item_id TEXT NOT NULL,
                    subject TEXT,
                    location TEXT,
                    start_time TIMESTAMP,
                    end_time TIMESTAMP,
                    is_all_day INTEGER DEFAULT 0,
                    recurrence_pattern TEXT,
                    attendees TEXT,
                    body TEXT,
                    reminder_minutes INTEGER,
                    status TEXT,
                    cached_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (account_id) REFERENCES accounts(id) ON DELETE CASCADE,
                    UNIQUE(account_id, item_id)
                )
            """)

            # Saved view state
            conn.execute("""
                CREATE TABLE IF NOT EXISTS saved_state (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    view_type TEXT NOT NULL,
                    account_id INTEGER,
                    folder_id TEXT,
                    selected_message_id TEXT,
                    scroll_position INTEGER,
                    view_order INTEGER
                )
            """)

            # Email address cache for autocomplete
            conn.execute("""
                CREATE TABLE IF NOT EXISTS email_cache (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    email TEXT NOT NULL UNIQUE COLLATE NOCASE,
                    name TEXT,
                    send_count INTEGER DEFAULT 0,
                    last_used TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            """)

            conn.execute("""
                CREATE INDEX IF NOT EXISTS idx_email_cache_send_count
                ON email_cache(send_count DESC)
            """)

            conn.commit()

    # ==================== Settings ====================

    def get_setting(self, key, default=None):
        with self._get_conn() as conn:
            cursor = conn.execute("SELECT value FROM settings WHERE key = ?", (key,))
            row = cursor.fetchone()
            return row[0] if row else default

    def set_setting(self, key, value):
        with self._get_conn() as conn:
            conn.execute(
                "INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?)",
                (key, value)
            )
            conn.commit()

    # ==================== Accounts ====================

    def get_accounts(self):
        """Get all accounts (without passwords)."""
        with self._get_conn() as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.execute("""
                SELECT id, name, email, server, username, auth_type, ews_url,
                       autodiscover, sync_interval, is_default, display_order, last_sync
                FROM accounts ORDER BY display_order, name
            """)
            return [dict(row) for row in cursor.fetchall()]

    def get_account(self, account_id):
        """Get a single account by ID (includes password)."""
        with self._get_conn() as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.execute("""
                SELECT id, name, email, server, username, password, auth_type, ews_url,
                       autodiscover, sync_interval, is_default, display_order, last_sync
                FROM accounts WHERE id = ?
            """, (account_id,))
            row = cursor.fetchone()
            return dict(row) if row else None

    def get_account_by_name(self, name):
        """Get a single account by name (includes password)."""
        with self._get_conn() as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.execute("""
                SELECT id, name, email, server, username, password, auth_type, ews_url,
                       autodiscover, sync_interval, is_default, display_order, last_sync
                FROM accounts WHERE name = ?
            """, (name,))
            row = cursor.fetchone()
            return dict(row) if row else None

    def save_account(self, name, email, server, username, password, auth_type='basic',
                     ews_url=None, autodiscover=False, sync_interval=300, is_default=False,
                     display_order=0, account_id=None):
        """Save or update an account."""
        with self._get_conn() as conn:
            if account_id:
                conn.execute("""
                    UPDATE accounts SET
                        name = ?, email = ?, server = ?, username = ?, password = ?,
                        auth_type = ?, ews_url = ?, autodiscover = ?, sync_interval = ?,
                        is_default = ?, display_order = ?
                    WHERE id = ?
                """, (name, email, server, username, password, auth_type, ews_url,
                      1 if autodiscover else 0, sync_interval, 1 if is_default else 0,
                      display_order, account_id))
            else:
                conn.execute("""
                    INSERT INTO accounts (name, email, server, username, password, auth_type,
                                         ews_url, autodiscover, sync_interval, is_default, display_order)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, (name, email, server, username, password, auth_type, ews_url,
                      1 if autodiscover else 0, sync_interval, 1 if is_default else 0, display_order))
            conn.commit()

    def delete_account(self, account_id):
        """Delete an account and all associated data."""
        with self._get_conn() as conn:
            conn.execute("DELETE FROM accounts WHERE id = ?", (account_id,))
            conn.commit()

    def update_last_sync(self, account_id):
        """Update the last_sync timestamp for an account."""
        with self._get_conn() as conn:
            conn.execute(
                "UPDATE accounts SET last_sync = CURRENT_TIMESTAMP WHERE id = ?",
                (account_id,)
            )
            conn.commit()

    # ==================== Folders ====================

    def get_folders(self, account_id):
        """Get all folders for an account."""
        with self._get_conn() as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.execute("""
                SELECT id, folder_id, name, parent_id, folder_type, unread_count, total_count
                FROM folders WHERE account_id = ? ORDER BY name
            """, (account_id,))
            return [dict(row) for row in cursor.fetchall()]

    def save_folder(self, account_id, folder_id, name, parent_id=None, folder_type=None,
                    unread_count=0, total_count=0, sync_state=None):
        """Save or update a folder."""
        with self._get_conn() as conn:
            conn.execute("""
                INSERT OR REPLACE INTO folders
                    (account_id, folder_id, name, parent_id, folder_type, unread_count, total_count, sync_state)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (account_id, folder_id, name, parent_id, folder_type, unread_count, total_count, sync_state))
            conn.commit()

    def update_folder_counts(self, account_id, folder_id, unread_count, total_count):
        """Update folder message counts."""
        with self._get_conn() as conn:
            conn.execute("""
                UPDATE folders SET unread_count = ?, total_count = ?
                WHERE account_id = ? AND folder_id = ?
            """, (unread_count, total_count, account_id, folder_id))
            conn.commit()

    def clear_folders(self, account_id):
        """Clear all folders for an account."""
        with self._get_conn() as conn:
            conn.execute("DELETE FROM folders WHERE account_id = ?", (account_id,))
            conn.commit()

    # ==================== Messages ====================

    def get_messages(self, account_id, folder_id, limit=100, offset=0):
        """Get messages from a folder."""
        with self._get_conn() as conn:
            conn.row_factory = sqlite3.Row
            if limit is None or limit < 0:
                cursor = conn.execute("""
                    SELECT id, item_id, subject, sender_name, sender_email, recipients,
                           date_received, is_read, has_attachments, importance, body_preview, is_flagged
                    FROM messages
                    WHERE account_id = ? AND folder_id = ?
                    ORDER BY date_received DESC
                """, (account_id, folder_id))
            else:
                cursor = conn.execute("""
                    SELECT id, item_id, subject, sender_name, sender_email, recipients,
                           date_received, is_read, has_attachments, importance, body_preview, is_flagged
                    FROM messages
                    WHERE account_id = ? AND folder_id = ?
                    ORDER BY date_received DESC
                    LIMIT ? OFFSET ?
                """, (account_id, folder_id, limit, offset))
            return [dict(row) for row in cursor.fetchall()]

    def get_message(self, message_id):
        """Get a single message by ID (includes body)."""
        with self._get_conn() as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.execute("""
                SELECT * FROM messages WHERE id = ?
            """, (message_id,))
            row = cursor.fetchone()
            return dict(row) if row else None

    def get_message_by_item_id(self, item_id):
        """Get a message by EWS item ID."""
        with self._get_conn() as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.execute("""
                SELECT * FROM messages WHERE item_id = ?
            """, (item_id,))
            row = cursor.fetchone()
            return dict(row) if row else None

    def save_message(self, account_id, folder_id, item_id, change_key=None,
                     conversation_id=None, subject=None, sender_name=None,
                     sender_email=None, recipients=None, cc=None,
                     date_received=None, date_sent=None, size=None,
                     importance='normal', is_read=False, has_attachments=False,
                     body_preview=None, body_type=None, body=None,
                     categories=None, is_flagged=False):
        """Save or update a message."""
        recipients_json = json.dumps(recipients) if recipients else None
        cc_json = json.dumps(cc) if cc else None
        categories_json = json.dumps(categories) if categories else None

        with self._get_conn() as conn:
            # Use ON CONFLICT to preserve body if it already exists
            conn.execute("""
                INSERT INTO messages
                    (account_id, folder_id, item_id, change_key, conversation_id,
                     subject, sender_name, sender_email, recipients, cc,
                     date_received, date_sent, size, importance, is_read,
                     has_attachments, body_preview, body_type, body, categories, is_flagged)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(item_id) DO UPDATE SET
                    folder_id = excluded.folder_id,
                    change_key = excluded.change_key,
                    subject = excluded.subject,
                    sender_name = excluded.sender_name,
                    sender_email = excluded.sender_email,
                    recipients = excluded.recipients,
                    cc = excluded.cc,
                    date_received = excluded.date_received,
                    is_read = excluded.is_read,
                    has_attachments = excluded.has_attachments,
                    body_preview = COALESCE(messages.body_preview, excluded.body_preview)
            """, (account_id, folder_id, item_id, change_key, conversation_id,
                  subject, sender_name, sender_email, recipients_json, cc_json,
                  date_received, date_sent, size, importance, 1 if is_read else 0,
                  1 if has_attachments else 0, body_preview, body_type, body,
                  categories_json, 1 if is_flagged else 0))
            conn.commit()

    def update_message_body(self, message_id, body, body_type='html'):
        """Update message body (for lazy loading)."""
        with self._get_conn() as conn:
            conn.execute(
                "UPDATE messages SET body = ?, body_type = ? WHERE id = ?",
                (body, body_type, message_id)
            )
            conn.commit()

    def update_message_read(self, message_id, is_read):
        """Update message read status."""
        with self._get_conn() as conn:
            conn.execute(
                "UPDATE messages SET is_read = ? WHERE id = ?",
                (1 if is_read else 0, message_id)
            )
            conn.commit()

    def update_message_flagged(self, message_id, is_flagged):
        """Update message flagged status."""
        with self._get_conn() as conn:
            conn.execute(
                "UPDATE messages SET is_flagged = ? WHERE id = ?",
                (1 if is_flagged else 0, message_id)
            )
            conn.commit()

    def delete_message(self, message_id):
        """Delete a message from cache."""
        with self._get_conn() as conn:
            conn.execute("DELETE FROM messages WHERE id = ?", (message_id,))
            conn.commit()

    def clear_messages(self, account_id, folder_id=None):
        """Clear cached messages for an account (optionally by folder)."""
        with self._get_conn() as conn:
            if folder_id:
                conn.execute(
                    "DELETE FROM messages WHERE account_id = ? AND folder_id = ?",
                    (account_id, folder_id)
                )
            else:
                conn.execute("DELETE FROM messages WHERE account_id = ?", (account_id,))
            conn.commit()

    def get_message_count(self, account_id, folder_id):
        """Get count of cached messages in a folder."""
        with self._get_conn() as conn:
            cursor = conn.execute(
                "SELECT COUNT(*) FROM messages WHERE account_id = ? AND folder_id = ?",
                (account_id, folder_id)
            )
            return cursor.fetchone()[0]

    # ==================== Attachments ====================

    def get_attachments(self, message_id):
        """Get attachments for a message."""
        with self._get_conn() as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.execute("""
                SELECT id, attachment_id, name, content_type, size, is_inline, content_id
                FROM attachments WHERE message_id = ?
            """, (message_id,))
            return [dict(row) for row in cursor.fetchall()]

    def save_attachment(self, message_id, attachment_id, name, content_type=None,
                        size=None, is_inline=False, content_id=None):
        """Save attachment metadata."""
        with self._get_conn() as conn:
            conn.execute("""
                INSERT INTO attachments
                    (message_id, attachment_id, name, content_type, size, is_inline, content_id)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (message_id, attachment_id, name, content_type, size,
                  1 if is_inline else 0, content_id))
            conn.commit()

    # ==================== View State ====================

    def save_view_state(self, view_type, account_id=None, folder_id=None,
                        selected_message_id=None, scroll_position=None):
        """Save current view state."""
        with self._get_conn() as conn:
            # Clear existing state for this view type
            conn.execute("DELETE FROM saved_state WHERE view_type = ?", (view_type,))
            conn.execute("""
                INSERT INTO saved_state
                    (view_type, account_id, folder_id, selected_message_id, scroll_position)
                VALUES (?, ?, ?, ?, ?)
            """, (view_type, account_id, folder_id, selected_message_id, scroll_position))
            conn.commit()

    def get_view_state(self, view_type):
        """Get saved view state."""
        with self._get_conn() as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.execute("""
                SELECT account_id, folder_id, selected_message_id, scroll_position
                FROM saved_state WHERE view_type = ?
            """, (view_type,))
            row = cursor.fetchone()
            return dict(row) if row else None

    # ==================== Email Cache ====================

    def add_email_to_cache(self, email, name=None, increment_send=False):
        """Add or update an email address in the cache.

        If increment_send is True, increment the send_count (for sent emails).
        Otherwise just add if new (for received emails).
        """
        if not email:
            return
        email = email.strip().lower()
        if not email:
            return

        with self._get_conn() as conn:
            if increment_send:
                # Insert or update with incremented send_count
                conn.execute("""
                    INSERT INTO email_cache (email, name, send_count, last_used)
                    VALUES (?, ?, 1, CURRENT_TIMESTAMP)
                    ON CONFLICT(email) DO UPDATE SET
                        name = COALESCE(excluded.name, email_cache.name),
                        send_count = email_cache.send_count + 1,
                        last_used = CURRENT_TIMESTAMP
                """, (email, name))
            else:
                # Just add if new, don't increment count
                conn.execute("""
                    INSERT INTO email_cache (email, name, send_count, last_used)
                    VALUES (?, ?, 0, CURRENT_TIMESTAMP)
                    ON CONFLICT(email) DO UPDATE SET
                        name = COALESCE(excluded.name, email_cache.name),
                        last_used = CURRENT_TIMESTAMP
                """, (email, name))
            conn.commit()

    def get_cached_emails(self):
        """Get all cached email addresses, sorted by send_count descending."""
        with self._get_conn() as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.execute("""
                SELECT email, name, send_count
                FROM email_cache
                ORDER BY send_count DESC, last_used DESC
            """)
            return [dict(row) for row in cursor.fetchall()]

    def get_unique_senders_from_messages(self):
        """Get unique sender emails from messages table for cache building."""
        with self._get_conn() as conn:
            conn.row_factory = sqlite3.Row
            cursor = conn.execute("""
                SELECT DISTINCT sender_email, sender_name
                FROM messages
                WHERE sender_email IS NOT NULL AND sender_email != ''
            """)
            return [dict(row) for row in cursor.fetchall()]

    def bulk_add_emails_to_cache(self, emails):
        """Bulk add emails to cache (for background building).

        emails: list of (email, name) tuples
        """
        with self._get_conn() as conn:
            for email, name in emails:
                if not email:
                    continue
                email = email.strip().lower()
                if not email:
                    continue
                # Only insert if not exists, don't update existing
                conn.execute("""
                    INSERT OR IGNORE INTO email_cache (email, name, send_count, last_used)
                    VALUES (?, ?, 0, CURRENT_TIMESTAMP)
                """, (email, name))
            conn.commit()
