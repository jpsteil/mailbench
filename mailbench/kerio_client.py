"""Kerio Connect JSON-RPC API client."""

from __future__ import annotations

import json
import threading
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass
from typing import Any, Callable, Dict, List, Optional
from urllib.parse import urljoin
import re
import requests


def parse_email_address(addr: str) -> tuple[str, str]:
    """Parse an email address string into (name, email).

    Handles formats like:
    - "email@example.com" -> ("", "email@example.com")
    - "Name <email@example.com>" -> ("Name", "email@example.com")
    - "<email@example.com>" -> ("", "email@example.com")
    """
    addr = addr.strip()
    # Match "Name <email>" or "<email>" format
    match = re.match(r'^(?:([^<]*?)\s*)?<([^>]+)>$', addr)
    if match:
        name = (match.group(1) or "").strip()
        email = match.group(2).strip()
        return (name, email)
    # Plain email address
    return ("", addr)


def clean_error_message(msg: str) -> str:
    """Clean up Kerio error messages with unfilled placeholders.

    Kerio sometimes returns error messages with %1, %2, etc. placeholders
    that aren't filled in. This function removes them for readability.
    """
    # Handle specific known error patterns for better readability
    if "Attachment with ID" in msg and "%1" in msg:
        return "Attachment reference expired. Please re-attach the file and try again."

    # Remove unfilled placeholders like %1, %2, etc.
    cleaned = re.sub(r'%\d+', '', msg)
    # Clean up extra whitespace and punctuation artifacts
    cleaned = re.sub(r'\s+', ' ', cleaned)
    cleaned = re.sub(r'\s+([.,;:])', r'\1', cleaned)
    cleaned = re.sub(r'([.,;:])\s*\1', r'\1', cleaned)  # Remove duplicate punctuation
    return cleaned.strip()


@dataclass
class KerioConfig:
    """Configuration for a Kerio Connect connection."""
    email: str
    username: str
    password: str
    server: str  # e.g., "mail.company.com"


class KerioSession:
    """Manages a single Kerio Connect JSON-RPC session."""

    def __init__(self, config: KerioConfig):
        self.config = config
        self.base_url = f"https://{config.server}/webmail/api/jsonrpc/"
        self.token: Optional[str] = None
        self.session = requests.Session()
        self.session.verify = True  # SSL verification
        self._request_id = 0
        self._lock = threading.Lock()

    def _next_id(self) -> int:
        with self._lock:
            self._request_id += 1
            return self._request_id

    def call(self, method: str, params: Optional[Dict] = None) -> Dict:
        """Make a JSON-RPC call."""
        request_id = self._next_id()

        payload = {
            "jsonrpc": "2.0",
            "id": request_id,
            "method": method,
        }
        if params:
            payload["params"] = params
        if self.token:
            payload["token"] = self.token

        headers = {
            "Accept": "application/json-rpc",
            "Content-Type": "application/json-rpc; charset=UTF-8",
        }
        if self.token:
            headers["X-Token"] = self.token

        response = self.session.post(
            self.base_url,
            json=payload,
            headers=headers,
            timeout=30
        )
        response.raise_for_status()

        result = response.json()
        if "error" in result:
            error = result["error"]
            raise KerioError(clean_error_message(error.get("message", "Unknown error")), error.get("code", -1))

        return result.get("result", {})

    def login(self) -> bool:
        """Authenticate with the Kerio server."""
        params = {
            "userName": self.config.username,
            "password": self.config.password,
            "application": {
                "name": "Mailbench",
                "vendor": "Mailbench",
                "version": "1.0"
            }
        }

        result = self.call("Session.login", params)
        self.token = result.get("token")
        return self.token is not None

    def logout(self):
        """End the session."""
        if self.token:
            try:
                self.call("Session.logout")
            except Exception:
                pass
            self.token = None

    def whoami(self) -> Dict:
        """Get current user info."""
        return self.call("Session.whoAmI")

    def get_signature(self) -> str:
        """Get user's email signature from webmail settings."""
        import re
        try:
            # Signature is in the dynamically generated defaults JS file
            url = f"https://{self.config.server}/webmail/generatedDefaults.js"
            headers = {"X-Token": self.token} if self.token else {}
            resp = self.session.get(url, headers=headers, timeout=10)

            if resp.status_code != 200:
                return ""

            # Extract mailSignature field
            match = re.search(r'mailSignature:\s*("(?:[^"\\]|\\.)*")', resp.text)
            if match:
                # Use json.loads to decode the escaped string
                return json.loads(match.group(1))
            return ""
        except Exception:
            return ""

    def upload_attachment(self, filename: str, content: bytes, content_type: str = None) -> tuple[Optional[str], Optional[str]]:
        """Upload an attachment file and return the attachment ID.

        Kerio requires attachments to be uploaded first, then referenced by ID
        when creating the mail.

        Args:
            filename: Name of the file
            content: File content as bytes
            content_type: MIME type (defaults to application/octet-stream)

        Returns:
            Tuple of (attachment_id, error_message). On success, error is None.
        """
        if not content_type:
            content_type = "application/octet-stream"

        upload_url = f"https://{self.config.server}/webmail/api/jsonrpc/attachment-upload/"

        headers = {}
        if self.token:
            headers["X-Token"] = self.token

        # Build multipart body manually with proper headers
        import uuid
        boundary = f"----WebKitFormBoundary{uuid.uuid4().hex[:16]}"

        body = []
        body.append(f'--{boundary}'.encode())
        body.append(f'Content-Disposition: form-data; name="file"; filename="{filename}"'.encode())
        body.append(f'Content-Description: {filename}'.encode())
        body.append(f'Content-Type: {content_type}'.encode())
        body.append(b'')
        body.append(content)
        body.append(f'--{boundary}--'.encode())

        multipart_body = b'\r\n'.join(body)
        headers["Content-Type"] = f"multipart/form-data; boundary={boundary}"
        headers["Content-Description"] = filename

        try:
            response = self.session.post(
                upload_url,
                data=multipart_body,
                headers=headers,
                timeout=60  # Longer timeout for file uploads
            )
            response.raise_for_status()

            result = response.json()

            # Check for error in response
            if "error" in result:
                error = result["error"]
                return None, error.get("message", str(error))

            # The upload response format is: {result: {fileUpload: {id: "...", ...}}}
            res = result.get("result", {})
            if isinstance(res, dict):
                file_upload = res.get("fileUpload", {})
                if isinstance(file_upload, dict) and "id" in file_upload:
                    return str(file_upload["id"]), None

            # Fallback: look for ID in other locations
            if isinstance(res, dict):
                for key in ["id", "attachmentId", "fileId", "uploadId"]:
                    if key in res and res[key]:
                        return str(res[key]), None

            # Return full response for debugging
            return None, f"Could not find attachment ID in response: {result}"
        except requests.exceptions.HTTPError as e:
            return None, f"HTTP {e.response.status_code}: {e.response.text[:200] if e.response.text else str(e)}"
        except Exception as e:
            return None, str(e)



class KerioError(Exception):
    """Kerio API error."""
    def __init__(self, message: str, code: int = -1):
        self.message = message
        self.code = code
        super().__init__(f"Kerio Error {code}: {message}")


class KerioConnectionPool:
    """Manages Kerio sessions for multiple accounts."""

    def __init__(self):
        self._sessions: Dict[int, KerioSession] = {}
        self._lock = threading.Lock()

    def connect(self, account_id: int, config: KerioConfig) -> KerioSession:
        """Create or return existing session."""
        with self._lock:
            if account_id in self._sessions:
                return self._sessions[account_id]

            session = KerioSession(config)
            session.login()
            self._sessions[account_id] = session
            return session

    def get_session(self, account_id: int) -> Optional[KerioSession]:
        """Get existing session."""
        return self._sessions.get(account_id)

    def disconnect(self, account_id: int):
        """Disconnect and remove session."""
        with self._lock:
            session = self._sessions.pop(account_id, None)
            if session:
                session.logout()

    def disconnect_all(self):
        """Disconnect all sessions."""
        with self._lock:
            for session in self._sessions.values():
                try:
                    session.logout()
                except Exception:
                    pass
            self._sessions.clear()

    def close_all(self):
        """Close all sessions immediately without logout (for fast shutdown)."""
        with self._lock:
            for session in self._sessions.values():
                try:
                    session.session.close()  # Close requests.Session
                except Exception:
                    pass
            self._sessions.clear()


class SyncManager:
    """Manages sync operations with background execution."""

    def __init__(self, pool: KerioConnectionPool, db, root=None):
        self.pool = pool
        self.db = db
        self.root = root
        self.executor = ThreadPoolExecutor(max_workers=4)
        self._sync_in_progress: Dict[int, bool] = {}
        self._shutdown = False
        self._change_listeners: Dict[int, bool] = {}  # account_id -> listening
        self._sync_keys: Dict[int, Dict] = {}  # account_id -> syncKey

    def _ui_callback(self, callback: Callable, *args, **kwargs):
        """Execute callback on UI thread if root is available."""
        if self._shutdown:
            return
        if self.root and callback:
            self.root.after(0, lambda: callback(*args, **kwargs))
        elif callback:
            callback(*args, **kwargs)

    def sync_folders(self, account_id: int, callback: Optional[Callable] = None):
        """Sync folder list for an account in background."""
        def do_sync():
            try:
                session = self.pool.get_session(account_id)
                if not session:
                    self._ui_callback(callback, False, "Account not connected")
                    return

                # Get folders
                result = session.call("Folders.get")
                folders = result.get("list", [])

                # Clear existing folders
                self.db.clear_folders(account_id)

                # Save folders to DB
                for folder in folders:
                    folder_type = self._get_folder_type(folder)
                    self.db.save_folder(
                        account_id=account_id,
                        folder_id=folder.get("id", ""),
                        name=folder.get("name", ""),
                        parent_id=folder.get("parentId"),
                        folder_type=folder_type,
                        unread_count=folder.get("unreadCount", 0),
                        total_count=folder.get("messageCount", 0)
                    )

                self.db.update_last_sync(account_id)
                self._ui_callback(callback, True, None)

            except Exception as e:
                self._ui_callback(callback, False, str(e))

        self.executor.submit(do_sync)

    def _get_folder_type(self, folder: Dict) -> str:
        """Determine folder type from Kerio folder data."""
        # Kerio provides a 'type' or we can infer from name
        folder_type = folder.get("type", "").lower()
        name_lower = folder.get("name", "").lower()

        if folder_type == "inbox" or name_lower == "inbox":
            return "inbox"
        elif folder_type == "sent" or "sent" in name_lower:
            return "sent"
        elif folder_type == "drafts" or "draft" in name_lower:
            return "drafts"
        elif folder_type == "trash" or "deleted" in name_lower or "trash" in name_lower:
            return "trash"
        elif folder_type == "junk" or "junk" in name_lower or "spam" in name_lower:
            return "junk"
        elif folder_type == "outbox" or "outbox" in name_lower:
            return "outbox"
        elif name_lower == "quarantine":
            return "quarantine"
        return "custom"

    def sync_messages(self, account_id: int, folder_id: str, limit: int = -1,
                      callback: Optional[Callable] = None):
        """Sync messages from a folder in background."""
        if self._sync_in_progress.get(account_id):
            return

        self._sync_in_progress[account_id] = True

        def do_sync():
            try:
                session = self.pool.get_session(account_id)
                if not session:
                    self._sync_in_progress[account_id] = False
                    self._ui_callback(callback, False, "Account not connected", None)
                    return

                # Query messages - only fetch fields needed for list display (47x faster!)
                query = {
                    "fields": ["id", "subject", "from", "to", "receiveDate", "isSeen", "hasAttachment", "isFlagged", "isAnswered", "isForwarded"],
                    "start": 0,
                    "limit": limit,
                    "orderBy": [{"columnName": "receiveDate", "direction": "Desc", "caseSensitive": False}]
                }
                result = session.call("Mails.get", {
                    "folderIds": [folder_id],
                    "query": query
                })

                messages = result.get("list", [])
                total_items = result.get("totalItems", 0)

                # Build UI data directly from API response (skip slow DB caching)
                messages_data = []
                for msg in messages:
                    item_id = msg.get("id", "")
                    if not item_id:
                        continue

                    # Parse sender
                    sender = msg.get("from", {})
                    sender_name = sender.get("name", "")
                    sender_email = sender.get("address", "")

                    # Parse recipients for sent/drafts folder display
                    to_list = msg.get("to", [])
                    if to_list:
                        # Get first recipient for display
                        first_to = to_list[0] if isinstance(to_list, list) else {}
                        to_name = first_to.get("name", "") if isinstance(first_to, dict) else ""
                        to_email = first_to.get("address", "") if isinstance(first_to, dict) else str(first_to)
                        to_count = len(to_list) if isinstance(to_list, list) else 1
                    else:
                        to_name = ""
                        to_email = ""
                        to_count = 0

                    # Check for attachments - Kerio may use different field names
                    has_attachments = (msg.get("hasAttachment", False) or
                                      msg.get("hasAttachments", False) or
                                      bool(msg.get("attachments", [])))

                    messages_data.append({
                        "item_id": item_id,
                        "subject": msg.get("subject", ""),
                        "sender_name": sender_name,
                        "sender_email": sender_email,
                        "to_name": to_name,
                        "to_email": to_email,
                        "to_count": to_count,
                        "date_received": msg.get("receiveDate", ""),
                        "is_read": msg.get("isSeen", False),
                        "has_attachments": has_attachments,
                        "is_flagged": msg.get("isFlagged", False),
                        "is_answered": msg.get("isAnswered", False),
                        "is_forwarded": msg.get("isForwarded", False)
                    })

                self._sync_in_progress[account_id] = False
                self._ui_callback(callback, True, None, messages_data)

            except Exception as e:
                self._sync_in_progress[account_id] = False
                self._ui_callback(callback, False, str(e), None)

        self.executor.submit(do_sync)

    def fetch_message_body(self, account_id: int, item_id: str,
                           callback: Optional[Callable] = None):
        """Fetch full message body on demand."""
        def do_fetch():
            try:
                session = self.pool.get_session(account_id)
                if not session:
                    self._ui_callback(callback, False, "Account not connected", None)
                    return

                # Fetch full message by ID
                result = session.call("Mails.getById", {"ids": [item_id]})
                # Kerio returns {result: [messages]} for getById
                messages = result.get("result", []) if isinstance(result.get("result"), list) else []
                if not messages:
                    # Try direct list in case API changed
                    messages = result.get("list", []) if isinstance(result.get("list"), list) else []
                if not messages and isinstance(result, list):
                    messages = result

                if not messages:
                    self._ui_callback(callback, False, "Message not found", None)
                    return

                msg = messages[0]

                # Get body from displayableParts
                # Kerio returns parts with contentType: ctTextHtml or ctTextPlain
                body = ""
                body_type = "text"
                parts = msg.get("displayableParts", [])

                # Prefer HTML, fall back to plain text
                for part in parts:
                    content_type = part.get("contentType", "")
                    if content_type == "ctTextHtml":
                        body = part.get("content", "")
                        body_type = "html"
                        break
                    elif content_type == "ctTextPlain" and not body:
                        body = part.get("content", "")

                # Handle inline attachments (images with contentId)
                # Replace cid: references with actual URLs
                attachments = msg.get("attachments", [])
                session = self.pool.get_session(account_id)
                if session and attachments and body:
                    base_url = f"https://{session.config.server}"
                    for att in attachments:
                        content_id = att.get("contentId", "")
                        att_url = att.get("url", "")
                        if content_id and att_url:
                            # Replace cid:contentId with full URL
                            full_url = base_url + att_url
                            body = body.replace(f'cid:{content_id}', full_url)

                # Get recipients (to, cc)
                to_list = msg.get("to", [])
                cc_list = msg.get("cc", [])

                # Format recipients as "Name <email>" strings for reply support
                def format_recipient(r):
                    name = r.get('name', '')
                    addr = r.get('address', '')
                    if name and addr:
                        return f"{name} <{addr}>"
                    return addr or name

                to_str = ", ".join(format_recipient(r) for r in to_list)
                cc_str = ", ".join(format_recipient(r) for r in cc_list)

                # Get non-inline attachments for display
                display_attachments = []
                for att in attachments:
                    # Skip inline attachments (they're shown in HTML body)
                    if att.get("contentId"):
                        continue
                    display_attachments.append({
                        "id": att.get("id", ""),
                        "name": att.get("name", "attachment"),
                        "size": att.get("size", 0),
                        "url": att.get("url", ""),
                        "contentType": att.get("contentType", "")
                    })

                # Return as dict with all message data
                msg_data = {
                    "item_id": item_id,
                    "body": body,
                    "body_type": body_type,
                    "to": to_str,
                    "cc": cc_str,
                    "attachments": display_attachments
                }

                self._ui_callback(callback, True, msg_data, None)

            except Exception as e:
                self._ui_callback(callback, False, str(e), None)

        self.executor.submit(do_fetch)

    def fetch_message_raw(self, account_id: int, item_id: str,
                          callback: Optional[Callable] = None):
        """Fetch raw message content (RFC822 format) for attachment."""
        def do_fetch():
            try:
                session = self.pool.get_session(account_id)
                if not session:
                    self._ui_callback(callback, False, None)
                    return

                # Use Mails.getRaw to get RFC822 content
                result = session.call("Mails.getRaw", {"ids": [item_id]})
                messages = result.get("result", [])
                if messages and len(messages) > 0:
                    raw_content = messages[0].get("raw", "")
                    self._ui_callback(callback, True, raw_content)
                else:
                    self._ui_callback(callback, False, None)

            except Exception as e:
                # If getRaw doesn't work, try building a simple .eml from message data
                try:
                    result = session.call("Mails.getById", {"ids": [item_id]})
                    messages = result.get("result", [])
                    if messages:
                        msg = messages[0]
                        # Build basic .eml content
                        from_addr = msg.get("from", {}).get("address", "")
                        to_list = ", ".join(r.get("address", "") for r in msg.get("to", []))
                        subject = msg.get("subject", "")
                        body = ""
                        for part in msg.get("displayableParts", []):
                            if part.get("contentType") == "ctTextPlain":
                                body = part.get("content", "")
                                break
                            elif part.get("contentType") == "ctTextHtml":
                                body = part.get("content", "")

                        eml = f"From: {from_addr}\r\n"
                        eml += f"To: {to_list}\r\n"
                        eml += f"Subject: {subject}\r\n"
                        eml += "MIME-Version: 1.0\r\n"
                        eml += "Content-Type: text/plain; charset=utf-8\r\n"
                        eml += "\r\n"
                        eml += body
                        self._ui_callback(callback, True, eml)
                    else:
                        self._ui_callback(callback, False, None)
                except Exception:
                    self._ui_callback(callback, False, None)

        self.executor.submit(do_fetch)

    def mark_as_read(self, account_id: int, item_id: str, is_read: bool = True,
                     callback: Optional[Callable] = None):
        """Mark a message as read/unread on the server.

        Security: This only updates the 'isSeen' flag on the server.
        It does NOT send a read receipt (MDN) to the sender, even if the
        original message requested one. This protects user privacy.
        """
        def do_mark():
            try:
                session = self.pool.get_session(account_id)
                if not session:
                    self._ui_callback(callback, False, "Account not connected")
                    return

                # Update message seen status (Kerio uses isSeen, not flags.seen)
                result = session.call("Mails.set", {
                    "mails": [{
                        "id": item_id,
                        "isSeen": is_read
                    }]
                })

                # Update cache
                cached = self.db.get_message_by_item_id(item_id)
                if cached:
                    self.db.update_message_read(cached["id"], is_read)

                self._ui_callback(callback, True, None)

            except Exception as e:
                self._ui_callback(callback, False, str(e))

        self.executor.submit(do_mark)

    def set_flag(self, account_id: int, item_id: str, is_flagged: bool = True,
                 callback: Optional[Callable] = None):
        """Set or clear the flag on a message."""
        def do_set_flag():
            try:
                session = self.pool.get_session(account_id)
                if not session:
                    self._ui_callback(callback, False, "Account not connected")
                    return

                # Update message flag status
                result = session.call("Mails.set", {
                    "mails": [{
                        "id": item_id,
                        "isFlagged": is_flagged
                    }]
                })

                self._ui_callback(callback, True, None)

            except Exception as e:
                self._ui_callback(callback, False, str(e))

        self.executor.submit(do_set_flag)

    def delete_message(self, account_id: int, item_id: str, hard_delete: bool = False,
                       callback: Optional[Callable] = None):
        """Delete a message. If hard_delete=True, permanently delete; otherwise move to trash."""
        def do_delete():
            try:
                session = self.pool.get_session(account_id)
                if not session:
                    self._ui_callback(callback, False, "Account not connected")
                    return

                if hard_delete:
                    # Permanently delete
                    session.call("Mails.remove", {"ids": [item_id]})
                else:
                    # Move to trash - find trash folder first
                    folders = self.db.get_folders(account_id)
                    trash_folder = None
                    for f in folders:
                        if f.get("folder_type") == "trash":
                            trash_folder = f
                            break

                    if trash_folder:
                        session.call("Mails.move", {
                            "ids": [item_id],
                            "folder": trash_folder["folder_id"]
                        })
                    else:
                        # No trash folder found, just remove
                        session.call("Mails.remove", {"ids": [item_id]})

                # Remove from cache
                cached = self.db.get_message_by_item_id(item_id)
                if cached:
                    self.db.delete_message(cached["id"])

                self._ui_callback(callback, True, None)

            except Exception as e:
                self._ui_callback(callback, False, str(e))

        self.executor.submit(do_delete)

    def move_message(self, account_id: int, item_id: str, target_folder_id: str,
                     callback: Optional[Callable] = None):
        """Move a message to a different folder."""
        def do_move():
            try:
                session = self.pool.get_session(account_id)
                if not session:
                    self._ui_callback(callback, False, "Account not connected")
                    return

                # Move message to target folder
                session.call("Mails.move", {
                    "ids": [item_id],
                    "folder": target_folder_id
                })

                self._ui_callback(callback, True, None)

            except Exception as e:
                self._ui_callback(callback, False, str(e))

        self.executor.submit(do_move)

    def empty_trash(self, account_id: int, folder_id: str,
                    callback: Optional[Callable] = None):
        """Permanently delete all messages in a trash folder."""
        def do_empty():
            try:
                session = self.pool.get_session(account_id)
                if not session:
                    self._ui_callback(callback, False, "Account not connected", 0)
                    return

                # Get all message IDs in the folder
                query = {
                    "fields": ["id"],
                    "start": 0,
                    "limit": 10000  # Large limit to get all messages
                }
                result = session.call("Mails.get", {
                    "folderIds": [folder_id],
                    "query": query
                })

                messages = result.get("list", [])
                if not messages:
                    self._ui_callback(callback, True, None, 0)
                    return

                # Extract IDs and permanently delete
                ids = [msg.get("id") for msg in messages if msg.get("id")]
                if ids:
                    session.call("Mails.remove", {"ids": ids})

                # Clear local cache for this folder
                self.db.clear_messages(account_id, folder_id)

                self._ui_callback(callback, True, None, len(ids))

            except Exception as e:
                self._ui_callback(callback, False, str(e), 0)

        self.executor.submit(do_empty)

    def send_message(self, account_id: int, to: List[str], subject: str, body: str,
                     cc: List[str] = None, bcc: List[str] = None,
                     attachments: List[dict] = None,
                     original_id: str = None, is_reply: bool = False, is_forward: bool = False,
                     callback: Optional[Callable] = None):
        """Send a new email with optional attachments.

        Security: This function intentionally does NOT request read receipts
        (Disposition-Notification-To header) to protect recipient privacy.
        """
        def do_send():
            try:
                session = self.pool.get_session(account_id)
                if not session:
                    self._ui_callback(callback, False, "Account not connected")
                    return

                # Build recipients - parse "Name <email>" format to extract email
                def build_recipient(addr):
                    name, email = parse_email_address(addr)
                    if name:
                        return {"name": name, "address": email}
                    return {"address": email}

                to_list = [build_recipient(addr) for addr in to]
                cc_list = [build_recipient(addr) for addr in (cc or [])]
                bcc_list = [build_recipient(addr) for addr in (bcc or [])]

                mail = {
                    "from": {"address": session.config.email},
                    "to": to_list,
                    "subject": subject,
                    "displayableParts": [{
                        "contentType": "ctTextHtml",
                        "content": body
                    }],
                    "send": True  # Actually send the email, not just save as draft
                }
                if cc_list:
                    mail["cc"] = cc_list
                if bcc_list:
                    mail["bcc"] = bcc_list

                # Upload and add attachments
                if attachments:
                    attachment_parts = []
                    for att in attachments:
                        content = att.get('content', b'')
                        if isinstance(content, str):
                            content = content.encode('utf-8')
                        name = att.get('name', 'attachment')

                        # Upload attachment first to get an ID
                        att_id, upload_error = session.upload_attachment(name, content)
                        if att_id:
                            attachment_parts.append({
                                "id": att_id,
                                "name": name,
                                "contentType": "application/octet-stream"
                            })
                        else:
                            # Upload failed - report error with details
                            error_msg = upload_error or "Unknown error"
                            self._ui_callback(callback, False,
                                f"Failed to upload attachment: {name}\n{error_msg}")
                            return

                    if attachment_parts:
                        mail["attachments"] = attachment_parts

                result = session.call("Mails.create", {"mails": [mail]})
                # Check for errors in result
                errors = result.get("errors", [])
                if errors:
                    error_msgs = [clean_error_message(e.get("message", str(e))) for e in errors]
                    self._ui_callback(callback, False, "; ".join(error_msgs))
                    return

                # Mark original message as answered/forwarded
                if original_id and (is_reply or is_forward):
                    try:
                        update_data = {"id": original_id}
                        if is_reply:
                            update_data["isAnswered"] = True
                        if is_forward:
                            update_data["isForwarded"] = True
                        session.call("Mails.set", {"mails": [update_data]})
                    except Exception:
                        pass  # Don't fail the send if marking fails

                self._ui_callback(callback, True, None)

            except Exception as e:
                self._ui_callback(callback, False, str(e))

        self.executor.submit(do_send)

    def start_change_listener(self, account_id: int, callback: Optional[Callable] = None):
        """Start listening for mailbox changes using long-polling.

        Callback is called with (account_id, changes_list) when changes occur.
        """
        if self._change_listeners.get(account_id):
            return  # Already listening

        self._change_listeners[account_id] = True

        def do_listen():
            try:
                session = self.pool.get_session(account_id)
                if not session:
                    return

                # Get initial sync key
                result = session.call("Changes.getSyncKey")
                sync_key = result.get("syncKey")
                self._sync_keys[account_id] = sync_key

                while self._change_listeners.get(account_id) and not self._shutdown:
                    try:
                        # Long-poll for changes (30 second timeout for efficiency)
                        result = session.call("Changes.get", {
                            "lastSyncKey": sync_key,
                            "timeout": 30
                        })

                        changes = result.get("list", [])
                        sync_key = result.get("syncKey")
                        self._sync_keys[account_id] = sync_key

                        if changes and callback:
                            self._ui_callback(callback, account_id, changes)

                    except Exception as e:
                        if not self._shutdown:
                            # Brief pause before retry on error
                            import time
                            time.sleep(1)

            except Exception:
                pass
            finally:
                self._change_listeners[account_id] = False

        self.executor.submit(do_listen)

    def stop_change_listener(self, account_id: int):
        """Stop listening for changes."""
        self._change_listeners[account_id] = False

    def fetch_contacts(self, account_id: int, callback: Optional[Callable] = None):
        """Fetch contacts from Kerio address book."""
        def do_fetch():
            try:
                session = self.pool.get_session(account_id)
                if not session:
                    self._ui_callback(callback, False, "Account not connected", [])
                    return

                contacts = []
                seen_emails = set()  # Avoid duplicates

                # First, get all contact folders
                try:
                    folders_result = session.call("Folders.get")
                    contact_folder_ids = []
                    for f in folders_result.get("list", []):
                        if f.get("type", "").lower() == "fcontact":
                            contact_folder_ids.append(f.get("id"))
                except Exception:
                    contact_folder_ids = []

                # Fetch contacts from each contact folder
                for folder_id in contact_folder_ids:
                    try:
                        # Use minimal query - don't specify fields
                        result = session.call("Contacts.get", {
                            "folderIds": [folder_id],
                            "query": {
                                "start": 0,
                                "limit": 500
                            }
                        })

                        for contact in result.get("list", []):
                            # Build name from available fields
                            name = contact.get("commonName", "")
                            if not name:
                                first = contact.get("firstName", "")
                                last = contact.get("surName", "")
                                name = f"{first} {last}".strip()

                            # Get all email addresses
                            for email_obj in contact.get("emailAddresses", []):
                                email = email_obj.get("address", "") if isinstance(email_obj, dict) else str(email_obj)
                                if email and email.lower() not in seen_emails:
                                    seen_emails.add(email.lower())
                                    contacts.append({
                                        "name": name,
                                        "email": email,
                                        "type": "contact"
                                    })
                    except Exception:
                        pass

                self._ui_callback(callback, True, None, contacts)

            except Exception as e:
                self._ui_callback(callback, False, str(e), [])

        self.executor.submit(do_fetch)

    def fetch_users(self, account_id: int, callback: Optional[Callable] = None):
        """Fetch internal mail users by extracting from sent messages."""
        def do_fetch():
            try:
                session = self.pool.get_session(account_id)
                if not session:
                    self._ui_callback(callback, False, "Account not connected", [])
                    return

                users = []
                seen_emails = set()

                # Get the email domain from the account
                account_email = session.config.email
                domain = account_email.split("@")[1] if "@" in account_email else ""

                # Find the Sent folder
                try:
                    folders_result = session.call("Folders.get")
                    sent_folder_id = None
                    for f in folders_result.get("list", []):
                        ftype = f.get("type", "").lower()
                        fname = f.get("name", "").lower()
                        if ftype == "fsent" or "sent" in fname:
                            sent_folder_id = f.get("id")
                            break
                except Exception:
                    sent_folder_id = None

                # Extract recipients from sent messages
                if sent_folder_id and domain:
                    try:
                        result = session.call("Mails.get", {
                            "folderIds": [sent_folder_id],
                            "query": {
                                "fields": ["id", "to", "cc"],
                                "start": 0,
                                "limit": 200,
                                "orderBy": [{"columnName": "receiveDate", "direction": "Desc"}]
                            }
                        })

                        for msg in result.get("list", []):
                            for recipient_list in [msg.get("to", []), msg.get("cc", [])]:
                                for r in recipient_list:
                                    email = r.get("address", "")
                                    name = r.get("name", "")
                                    # Only include users from the same domain (internal users)
                                    if email and email.lower().endswith("@" + domain.lower()):
                                        if email.lower() not in seen_emails:
                                            seen_emails.add(email.lower())
                                            users.append({
                                                "name": name,
                                                "email": email,
                                                "type": "user"
                                            })
                    except Exception:
                        pass

                self._ui_callback(callback, True, None, users)

            except Exception as e:
                self._ui_callback(callback, False, str(e), [])

        self.executor.submit(do_fetch)

    def fetch_signature(self, account_id: int, callback: Optional[Callable] = None):
        """Fetch user's email signature."""
        def do_fetch():
            try:
                session = self.pool.get_session(account_id)
                if not session:
                    self._ui_callback(callback, False, "Not connected", "")
                    return

                signature = session.get_signature()
                self._ui_callback(callback, True, None, signature)

            except Exception as e:
                self._ui_callback(callback, False, str(e), "")

        self.executor.submit(do_fetch)

    def shutdown(self):
        """Shutdown the executor."""
        self._shutdown = True
        # Stop all change listeners
        for account_id in list(self._change_listeners.keys()):
            self._change_listeners[account_id] = False
        self.executor.shutdown(wait=False, cancel_futures=True)

    # ==================== Folder Management ====================

    def create_folder(self, account_id: int, name: str, parent_id: str = None,
                      callback: Optional[Callable] = None):
        """Create a new mail folder."""
        def do_create():
            try:
                session = self.pool.get_session(account_id)
                if not session:
                    self._ui_callback(callback, False, "Account not connected", None)
                    return

                folder_data = {
                    "name": name,
                    "type": "fmail"  # Mail folder type
                }
                if parent_id:
                    folder_data["parentId"] = parent_id

                result = session.call("Folders.create", {"folders": [folder_data]})

                # Get created folder ID
                created = result.get("result", [])
                if created and len(created) > 0:
                    folder_id = created[0].get("id", "")
                    self._ui_callback(callback, True, None, folder_id)
                else:
                    self._ui_callback(callback, True, None, None)

            except Exception as e:
                self._ui_callback(callback, False, str(e), None)

        self.executor.submit(do_create)

    def get_junk_folder(self, account_id: int,
                        callback: Optional[Callable] = None):
        """Get the Junk/Spam folder ID for blocked messages."""
        def do_get():
            try:
                session = self.pool.get_session(account_id)
                if not session:
                    self._ui_callback(callback, False, "Account not connected", None)
                    return

                result = session.call("Folders.get")
                folders = result.get("list", [])

                # Look for junk/spam folder
                for folder in folders:
                    name = folder.get("name", "").lower()
                    folder_type = folder.get("type", "").lower()
                    if folder_type == "fjunk" or "junk" in name or "spam" in name:
                        self._ui_callback(callback, True, None, folder.get("id"))
                        return

                self._ui_callback(callback, False, "Junk folder not found", None)

            except Exception as e:
                self._ui_callback(callback, False, str(e), None)

        self.executor.submit(do_get)

    # ==================== Notes API ====================

    def get_notes_folder_id(self, account_id: int, callback: Optional[Callable] = None):
        """Find the Notes folder ID for an account."""
        def do_fetch():
            try:
                session = self.pool.get_session(account_id)
                if not session:
                    self._ui_callback(callback, False, "Account not connected", None)
                    return

                result = session.call("Folders.get")
                folders = result.get("list", [])

                # Look for notes folder (type "fnotes" or name "Notes")
                for folder in folders:
                    folder_type = folder.get("type", "").lower()
                    folder_name = folder.get("name", "").lower()
                    if folder_type == "fnotes" or folder_name == "notes":
                        self._ui_callback(callback, True, None, folder.get("id"))
                        return

                self._ui_callback(callback, False, "Notes folder not found", None)

            except Exception as e:
                self._ui_callback(callback, False, str(e), None)

        self.executor.submit(do_fetch)

    def fetch_notes(self, account_id: int, folder_id: str,
                    callback: Optional[Callable] = None):
        """Fetch all notes from a notes folder."""
        def do_fetch():
            try:
                session = self.pool.get_session(account_id)
                if not session:
                    self._ui_callback(callback, False, "Account not connected", [])
                    return

                result = session.call("Notes.get", {
                    "folderIds": [folder_id],
                    "query": {
                        "start": 0,
                        "limit": 500
                    }
                })

                notes = result.get("list", [])
                # Extract relevant fields
                notes_data = []
                for note in notes:
                    text = note.get("text", "")
                    # Subject is first line of text
                    lines = text.split("\n", 1)
                    subject = lines[0] if lines else ""
                    body = lines[1].lstrip("\n") if len(lines) > 1 else ""
                    notes_data.append({
                        "id": note.get("id", ""),
                        "subject": subject,
                        "body": body,
                    })

                self._ui_callback(callback, True, None, notes_data)

            except Exception as e:
                self._ui_callback(callback, False, str(e), [])

        self.executor.submit(do_fetch)

    def create_note(self, account_id: int, folder_id: str, subject: str, text: str,
                    callback: Optional[Callable] = None):
        """Create a new note.

        Note: Kerio Notes use 'text' field, not 'body'. The subject is stored
        as the first line of text. API-created notes may crash Kerio's web UI
        when opened, but work fine via API.
        """
        def do_create():
            try:
                session = self.pool.get_session(account_id)
                if not session:
                    self._ui_callback(callback, False, "Account not connected", None)
                    return

                # Store subject as first line, then the content
                full_text = f"{subject}\n\n{text}"
                result = session.call("Notes.create", {
                    "notes": [{
                        "folderId": folder_id,
                        "text": full_text
                    }]
                })

                # Get created note ID
                created = result.get("result", [])
                if created and len(created) > 0:
                    note_id = created[0].get("id", "")
                    self._ui_callback(callback, True, None, note_id)
                else:
                    self._ui_callback(callback, True, None, None)

            except Exception as e:
                self._ui_callback(callback, False, str(e), None)

        self.executor.submit(do_create)

    def update_note(self, account_id: int, note_id: str, subject: str, text: str,
                    callback: Optional[Callable] = None):
        """Update an existing note."""
        def do_update():
            try:
                session = self.pool.get_session(account_id)
                if not session:
                    self._ui_callback(callback, False, "Account not connected")
                    return

                # Store subject as first line, then the content
                full_text = f"{subject}\n\n{text}"
                session.call("Notes.set", {
                    "notes": [{
                        "id": note_id,
                        "text": full_text
                    }]
                })

                self._ui_callback(callback, True, None)

            except Exception as e:
                self._ui_callback(callback, False, str(e))

        self.executor.submit(do_update)

    def delete_note(self, account_id: int, note_id: str,
                    callback: Optional[Callable] = None):
        """Delete a note."""
        def do_delete():
            try:
                session = self.pool.get_session(account_id)
                if not session:
                    self._ui_callback(callback, False, "Account not connected")
                    return

                session.call("Notes.remove", {"ids": [note_id]})
                self._ui_callback(callback, True, None)

            except Exception as e:
                self._ui_callback(callback, False, str(e))

        self.executor.submit(do_delete)

    def find_note_by_subject(self, account_id: int, folder_id: str, subject: str,
                              callback: Optional[Callable] = None):
        """Find a note by its subject (exact match)."""
        def do_find():
            try:
                session = self.pool.get_session(account_id)
                if not session:
                    self._ui_callback(callback, False, "Account not connected", None)
                    return

                result = session.call("Notes.get", {
                    "folderIds": [folder_id],
                    "query": {
                        "start": 0,
                        "limit": 500
                    }
                })

                notes = result.get("list", [])
                for note in notes:
                    text = note.get("text", "")
                    # Subject is first line of text
                    lines = text.split("\n", 1)
                    note_subject = lines[0] if lines else ""
                    body = lines[1].lstrip("\n") if len(lines) > 1 else ""

                    if note_subject == subject:
                        self._ui_callback(callback, True, None, {
                            "id": note.get("id", ""),
                            "subject": note_subject,
                            "body": body
                        })
                        return

                # Not found
                self._ui_callback(callback, True, None, None)

            except Exception as e:
                self._ui_callback(callback, False, str(e), None)

        self.executor.submit(do_find)
