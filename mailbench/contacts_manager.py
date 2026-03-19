"""Contacts manager for syncing contacts between local DB and Kerio server."""

import json
from typing import Callable, Optional, List

from PySide6.QtCore import QTimer


class ContactsManager:
    """Manages contact sync between local SQLite and Kerio server."""

    def __init__(self, db, sync_manager):
        self.db = db
        self.sync_manager = sync_manager
        self._account_id: Optional[int] = None
        self._sync_timer: Optional[QTimer] = None
        self._folders_loaded = False

    def initialize(self, account_id: int, callback: Optional[Callable] = None):
        """Initialize contacts manager for an account.

        Syncs contact folders and contacts from server.
        """
        self._account_id = account_id
        self._folders_loaded = False

        # First sync folders
        self.sync_folders(callback)

    def sync_folders(self, callback: Optional[Callable] = None):
        """Sync contact folders from server."""
        if not self._account_id:
            if callback:
                callback(False, "Not initialized", [])
            return

        def on_folders_synced(success, error, folders):
            if not success:
                if callback:
                    callback(False, error, [])
                return

            # Clear existing folders
            self.db.clear_contact_folders(self._account_id)

            # Save new folders
            for folder in folders:
                self.db.save_contact_folder(
                    account_id=self._account_id,
                    folder_id=folder.get("folder_id", ""),
                    name=folder.get("name", ""),
                    parent_id=folder.get("parent_id"),
                    is_default=folder.get("is_default", False)
                )

            self._folders_loaded = True

            if callback:
                callback(True, None, folders)

        self.sync_manager.sync_contact_folders(self._account_id, on_folders_synced)

    def sync_contacts(self, folder_id: str, callback: Optional[Callable] = None):
        """Sync contacts from a specific folder."""
        if not self._account_id:
            if callback:
                callback(False, "Not initialized", [])
            return

        def on_contacts_fetched(success, error, contacts):
            if not success:
                if callback:
                    callback(False, error, [])
                return

            # Clear existing contacts for this folder
            self.db.clear_contacts(self._account_id, folder_id)

            # Save new contacts
            for contact in contacts:
                # Convert lists to JSON strings for storage
                emails_json = json.dumps(contact.get("email_addresses", []))
                phones_json = json.dumps(contact.get("phone_numbers", []))
                home_addr = contact.get("home_address")
                work_addr = contact.get("work_address")
                home_addr_json = json.dumps(home_addr) if home_addr else None
                work_addr_json = json.dumps(work_addr) if work_addr else None

                self.db.save_contact(
                    account_id=self._account_id,
                    folder_id=folder_id,
                    item_id=contact.get("item_id", ""),
                    common_name=contact.get("common_name"),
                    first_name=contact.get("first_name"),
                    last_name=contact.get("last_name"),
                    nickname=contact.get("nickname"),
                    title=contact.get("title"),
                    company=contact.get("company"),
                    job_title=contact.get("job_title"),
                    department=contact.get("department"),
                    email_addresses=emails_json,
                    phone_numbers=phones_json,
                    home_address=home_addr_json,
                    work_address=work_addr_json,
                    website=contact.get("website"),
                    birthday=contact.get("birthday"),
                    anniversary=contact.get("anniversary"),
                    notes=contact.get("notes"),
                    photo_url=contact.get("photo_url")
                )

            if callback:
                callback(True, None, contacts)

        self.sync_manager.fetch_contacts_full(self._account_id, folder_id, on_contacts_fetched)

    def sync_all_contacts(self, callback: Optional[Callable] = None):
        """Sync contacts from all folders."""
        if not self._account_id:
            if callback:
                callback(False, "Not initialized", [])
            return

        folders = self.db.get_contact_folders(self._account_id)
        if not folders:
            if callback:
                callback(True, None, [])
            return

        all_contacts = []
        pending = len(folders)

        def on_folder_done(success, error, contacts):
            nonlocal pending, all_contacts
            if success and contacts:
                all_contacts.extend(contacts)
            pending -= 1
            if pending == 0 and callback:
                callback(True, None, all_contacts)

        for folder in folders:
            self.sync_contacts(folder.get("folder_id"), on_folder_done)

    def get_folders(self) -> List[dict]:
        """Get contact folders from local DB."""
        if not self._account_id:
            return []
        return self.db.get_contact_folders(self._account_id)

    def get_contacts(self, folder_id: str = None) -> List[dict]:
        """Get contacts from local DB, optionally filtered by folder."""
        if not self._account_id:
            return []
        return self.db.get_contacts(self._account_id, folder_id)

    def get_all_contacts(self) -> List[dict]:
        """Get all contacts from local DB."""
        if not self._account_id:
            return []
        return self.db.get_contacts(self._account_id)

    def get_contact(self, item_id: str) -> Optional[dict]:
        """Get a single contact by item_id."""
        if not self._account_id:
            return None
        return self.db.get_contact(self._account_id, item_id)

    def search(self, query: str) -> List[dict]:
        """Search contacts by name, email, or company."""
        if not self._account_id:
            return []
        return self.db.search_contacts(self._account_id, query)

    def create_contact(self, folder_id: str, contact_data: dict,
                       callback: Optional[Callable] = None):
        """Create a new contact on server and in local DB."""
        if not self._account_id:
            if callback:
                callback(False, "Not initialized", None)
            return

        def on_created(success, error, item_id):
            if not success:
                if callback:
                    callback(False, error, None)
                return

            # Save to local DB
            emails_json = json.dumps(contact_data.get("email_addresses", []))
            phones_json = json.dumps(contact_data.get("phone_numbers", []))

            self.db.save_contact(
                account_id=self._account_id,
                folder_id=folder_id,
                item_id=item_id,
                common_name=contact_data.get("common_name"),
                first_name=contact_data.get("first_name"),
                last_name=contact_data.get("last_name"),
                company=contact_data.get("company"),
                job_title=contact_data.get("job_title"),
                department=contact_data.get("department"),
                email_addresses=emails_json,
                phone_numbers=phones_json,
                website=contact_data.get("website"),
                birthday=contact_data.get("birthday"),
                notes=contact_data.get("notes")
            )

            if callback:
                callback(True, None, item_id)

        self.sync_manager.create_contact(self._account_id, folder_id, contact_data, on_created)

    def update_contact(self, item_id: str, contact_data: dict,
                       callback: Optional[Callable] = None):
        """Update a contact on server and in local DB."""
        if not self._account_id:
            if callback:
                callback(False, "Not initialized")
            return

        def on_updated(success, error):
            if not success:
                if callback:
                    callback(False, error)
                return

            # Update local DB
            existing = self.db.get_contact(self._account_id, item_id)
            if existing:
                emails_json = json.dumps(contact_data.get("email_addresses", []))
                phones_json = json.dumps(contact_data.get("phone_numbers", []))

                self.db.save_contact(
                    account_id=self._account_id,
                    folder_id=existing.get("folder_id", ""),
                    item_id=item_id,
                    common_name=contact_data.get("common_name"),
                    first_name=contact_data.get("first_name"),
                    last_name=contact_data.get("last_name"),
                    company=contact_data.get("company"),
                    job_title=contact_data.get("job_title"),
                    department=contact_data.get("department"),
                    email_addresses=emails_json,
                    phone_numbers=phones_json,
                    website=contact_data.get("website"),
                    birthday=contact_data.get("birthday"),
                    notes=contact_data.get("notes")
                )

            if callback:
                callback(True, None)

        self.sync_manager.update_contact(self._account_id, item_id, contact_data, on_updated)

    def delete_contact(self, item_id: str, callback: Optional[Callable] = None):
        """Delete a contact from server and local DB."""
        if not self._account_id:
            if callback:
                callback(False, "Not initialized")
            return

        def on_deleted(success, error):
            if not success:
                if callback:
                    callback(False, error)
                return

            # Remove from local DB
            self.db.delete_contact(self._account_id, item_id)

            if callback:
                callback(True, None)

        self.sync_manager.delete_contact(self._account_id, item_id, on_deleted)

    def start_periodic_sync(self, interval_ms: int = 300000):
        """Start periodic sync (default: 5 minutes)."""
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
        if self._account_id:
            self.sync_all_contacts()

    def clear(self):
        """Clear all cached data."""
        if self._account_id:
            self.db.clear_contacts(self._account_id)
            self.db.clear_contact_folders(self._account_id)
