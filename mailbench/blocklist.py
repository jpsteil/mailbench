"""Blocklist manager for syncing blocked senders between local DB and Kerio Notes."""

import json
from datetime import datetime
from typing import Callable, Optional

from PySide6.QtCore import QTimer


# Note subject for blocklist storage
BLOCKLIST_NOTE = "[Mailbench] Blocked Senders"

# Top email domains that should never be blocked at domain level
DEFAULT_ALLOWED_DOMAINS = [
    # Major providers
    "gmail.com", "googlemail.com", "yahoo.com", "hotmail.com", "outlook.com",
    "aol.com", "icloud.com", "me.com", "mac.com", "live.com", "msn.com",
    "protonmail.com", "proton.me", "zoho.com", "mail.com", "yandex.com",
    "gmx.com", "gmx.net",
    # US ISPs
    "comcast.net", "verizon.net", "att.net", "sbcglobal.net", "bellsouth.net",
    "cox.net", "charter.net", "earthlink.net", "spectrum.net",
    # International Yahoo/Hotmail
    "yahoo.co.uk", "yahoo.co.jp", "yahoo.com.br", "yahoo.de", "yahoo.fr",
    "hotmail.co.uk", "hotmail.de", "hotmail.fr", "outlook.de", "outlook.fr",
    # German providers
    "web.de", "gmx.de", "t-online.de", "freenet.de",
    # French providers
    "orange.fr", "free.fr", "wanadoo.fr", "sfr.fr",
    # Italian providers
    "libero.it", "virgilio.it", "tin.it",
    # Chinese providers
    "qq.com", "163.com", "126.com", "sina.com", "sohu.com",
    # Korean providers
    "naver.com", "daum.net", "hanmail.net",
    # Russian providers
    "mail.ru", "rambler.ru", "yandex.ru",
    # Other
    "rediffmail.com", "tutanota.com", "fastmail.com", "hey.com",
]


class BlocklistManager:
    """Manages blocked sender lists synced between local SQLite and Kerio Notes."""

    def __init__(self, db, sync_manager):
        self.db = db
        self.sync_manager = sync_manager
        self._notes_folder_id: Optional[str] = None
        self._note_id: Optional[str] = None
        self._sync_timer: Optional[QTimer] = None
        self._account_id: Optional[int] = None
        self._user_domain: Optional[str] = None  # User's email domain

    def initialize(self, account_id: int, user_email: str = None,
                   callback: Optional[Callable] = None):
        """Initialize blocklist manager for an account.

        Syncs blocklist from Kerio Notes to local DB.
        Initializes allowed domains list with defaults + user's domain.
        """
        self._account_id = account_id

        # Extract user's domain to add to allowed list
        if user_email and '@' in user_email:
            self._user_domain = user_email.split('@')[1].lower()

        # Initialize allowed domains with defaults if empty
        self._init_allowed_domains()

        def on_folder_found(success, error, folder_id):
            if not success or not folder_id:
                # Notes folder not found - use local-only mode
                print(f"Notes folder not found, using local-only blocklist: {error}")
                if callback:
                    callback(True, None)
                return

            self._notes_folder_id = folder_id
            # Load blocklist data from notes
            self._load_from_server(callback)

        self.sync_manager.get_notes_folder_id(account_id, on_folder_found)

    def _init_allowed_domains(self):
        """Initialize allowed domains list with defaults if empty."""
        existing = self.db.get_allowed_domains()
        if not existing:
            # Add default domains
            for domain in DEFAULT_ALLOWED_DOMAINS:
                self.db.add_allowed_domain(domain)
            # Add user's own domain
            if self._user_domain:
                self.db.add_allowed_domain(self._user_domain)

    def _load_from_server(self, callback: Optional[Callable] = None):
        """Load blocklist from server Note into local DB."""
        if not self._notes_folder_id or not self._account_id:
            if callback:
                callback(False, "Not initialized")
            return

        def on_note_found(success, error, note):
            if success and note:
                self._note_id = note["id"]
                self._parse_and_store(note["body"])
            if callback:
                callback(True, None)

        self.sync_manager.find_note_by_subject(
            self._account_id, self._notes_folder_id,
            BLOCKLIST_NOTE, on_note_found
        )

    def _parse_and_store(self, body: str):
        """Parse JSON from note body and store in local DB."""
        if not body:
            return

        try:
            data = json.loads(body)
            # Clear existing and load fresh from server
            self.db.clear_blocklist()
            self.db.clear_trusted_senders()

            domains = data.get("domains", [])
            if domains:
                self.db.bulk_add_blocked_domains(domains)

            emails = data.get("emails", [])
            if emails:
                self.db.bulk_add_blocked_emails(emails)

            # Load allowed_domains from server (merge with defaults)
            allowed = data.get("allowed_domains", [])
            if allowed:
                self.db.bulk_add_allowed_domains(allowed)

            # Load trusted senders for remote images
            trusted = data.get("trusted_senders", [])
            if trusted:
                self.db.bulk_add_trusted_senders(trusted)
        except json.JSONDecodeError:
            pass  # Invalid JSON, ignore

    def _serialize(self) -> str:
        """Serialize all blocklist entries from local DB to JSON."""
        domains = self.db.get_blocked_domains()
        emails = self.db.get_blocked_emails()
        allowed = self.db.get_allowed_domains()
        trusted = self.db.get_trusted_senders()

        data = {
            "version": 1,
            "modified_at": datetime.utcnow().isoformat() + "Z",
            "domains": [{
                "value": e["domain"],
                "blocked_count": e["blocked_count"],
                "last_blocked": e["last_blocked"]
            } for e in domains],
            "emails": [{
                "value": e["email"],
                "blocked_count": e["blocked_count"],
                "last_blocked": e["last_blocked"]
            } for e in emails],
            "allowed_domains": [e["domain"] for e in allowed],
            "trusted_senders": trusted
        }
        return json.dumps(data, indent=2)

    def _save_to_server(self, callback: Optional[Callable] = None):
        """Save blocklist to server Note."""
        if not self._notes_folder_id or not self._account_id:
            if callback:
                callback(False, "Not initialized")
            return

        body = self._serialize()

        def on_saved(success, error, *args):
            if callback:
                callback(success, error)

        def on_created(success, error, new_note_id):
            if success and new_note_id:
                self._note_id = new_note_id
            if callback:
                callback(success, error)

        if self._note_id:
            # Update existing note
            self.sync_manager.update_note(
                self._account_id, self._note_id, BLOCKLIST_NOTE, body, on_saved
            )
        else:
            # Create new note
            self.sync_manager.create_note(
                self._account_id, self._notes_folder_id, BLOCKLIST_NOTE, body, on_created
            )

    def add_domain(self, domain: str, callback: Optional[Callable] = None):
        """Add a domain to blocklist (local + server sync)."""
        domain = domain.lower().strip()
        # Check if domain is in allowed list
        if self.db.is_allowed_domain(domain):
            if callback:
                callback(False, f"Cannot block '{domain}' - it's in the allowed domains list")
            return
        if self.db.add_blocked_domain(domain):
            self._save_to_server(callback=callback)
        elif callback:
            callback(True, None)  # Already exists

    def add_email(self, email: str, callback: Optional[Callable] = None):
        """Add an email to blocklist (local + server sync)."""
        email = email.lower().strip()
        if self.db.add_blocked_email(email):
            self._save_to_server(callback=callback)
        elif callback:
            callback(True, None)  # Already exists

    def remove_domain(self, domain: str, callback: Optional[Callable] = None):
        """Remove a domain from blocklist (local + server sync)."""
        self.db.remove_blocked_domain(domain)
        self._save_to_server(callback=callback)

    def remove_email(self, email: str, callback: Optional[Callable] = None):
        """Remove an email from blocklist (local + server sync)."""
        self.db.remove_blocked_email(email)
        self._save_to_server(callback=callback)

    def is_blocked(self, sender_email: str) -> tuple[bool, str]:
        """Check if sender is blocked. Returns (is_blocked, match_type)."""
        return self.db.is_blocked(sender_email)

    def increment_blocked(self, value: str, is_domain: bool):
        """Increment blocked count for a blocklist entry."""
        self.db.increment_blocked_count(value, is_domain)
        # Don't sync immediately for every block - batch updates periodically

    def get_domains(self) -> list:
        """Get all blocked domains."""
        return self.db.get_blocked_domains()

    def get_emails(self) -> list:
        """Get all blocked emails."""
        return self.db.get_blocked_emails()

    def get_allowed_domains(self) -> list:
        """Get all allowed domains."""
        return self.db.get_allowed_domains()

    def add_allowed_domain(self, domain: str, callback: Optional[Callable] = None):
        """Add a domain to allowed list (local + server sync)."""
        domain = domain.lower().strip()
        if self.db.add_allowed_domain(domain):
            # If domain is currently blocked, remove it
            self.db.remove_blocked_domain(domain)
            self._save_to_server(callback=callback)
        elif callback:
            callback(True, None)  # Already exists

    def remove_allowed_domain(self, domain: str, callback: Optional[Callable] = None):
        """Remove a domain from allowed list (local + server sync)."""
        self.db.remove_allowed_domain(domain)
        self._save_to_server(callback=callback)

    def is_allowed(self, domain: str) -> bool:
        """Check if domain is in allowed list."""
        return self.db.is_allowed_domain(domain)

    # ==================== Trusted Senders (for remote images) ====================

    def add_trusted_sender(self, email: str, callback: Optional[Callable] = None):
        """Add an email to trusted senders list (local + server sync)."""
        email = email.lower().strip()
        self.db.add_trusted_sender(email)
        self._save_to_server(callback=callback)

    def remove_trusted_sender(self, email: str, callback: Optional[Callable] = None):
        """Remove an email from trusted senders list (local + server sync)."""
        self.db.remove_trusted_sender(email)
        self._save_to_server(callback=callback)

    def is_trusted_sender(self, email: str) -> bool:
        """Check if sender is in trusted list for remote images."""
        return self.db.is_trusted_sender(email)

    def get_trusted_senders(self) -> list:
        """Get all trusted sender emails."""
        return self.db.get_trusted_senders()

    def start_periodic_sync(self, interval_ms: int = 600000):
        """Start periodic sync from server (default: 10 minutes)."""
        if self._sync_timer:
            self._sync_timer.stop()

        self._sync_timer = QTimer()
        self._sync_timer.timeout.connect(self._periodic_sync)
        self._sync_timer.start(interval_ms)

    def stop_periodic_sync(self):
        """Stop periodic sync."""
        if self._sync_timer:
            self._sync_timer.stop()
            self._sync_timer = None

    def _periodic_sync(self):
        """Called by timer to refresh from server."""
        if self._account_id and self._notes_folder_id:
            self._load_from_server()

    def sync_now(self, callback: Optional[Callable] = None):
        """Force immediate sync from server."""
        self._load_from_server(callback)

    def save_all(self, callback: Optional[Callable] = None):
        """Save blocklist to server."""
        if not self._notes_folder_id:
            if callback:
                callback(True, None)
            return

        self._save_to_server(callback=callback)
