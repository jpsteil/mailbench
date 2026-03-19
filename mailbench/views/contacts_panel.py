"""Contacts panel with folder tree and contact list."""

import json
from typing import Optional

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QSplitter, QTreeView, QListView,
    QLineEdit, QPushButton, QFrame, QAbstractItemView, QLabel, QMenu
)
from PySide6.QtGui import QStandardItemModel, QStandardItem, QIcon, QFont, QPalette, QAction
from PySide6.QtWidgets import QStyle, QApplication

from mailbench.models.contact_model import ContactListModel, ContactDelegate


class ContactsPanel(QWidget):
    """Main contacts panel with folders and contact list."""

    folderSelected = Signal(int, str, str)  # account_id, folder_id, folder_name
    contactSelected = Signal(int, str)  # account_id, item_id
    contactDoubleClicked = Signal(int, str)  # account_id, item_id
    newContactRequested = Signal()
    newMessageToContact = Signal(str, str)  # email, name

    def __init__(self, parent=None):
        super().__init__(parent)
        self._current_account_id: Optional[int] = None
        self._current_folder_id: Optional[str] = None
        self._folder_items: dict[str, QStandardItem] = {}
        self._user_domain: str = ""
        self._setup_ui()

    def get_user_domain(self) -> str:
        """Get the user's email domain for filtering."""
        return self._user_domain

    def _setup_ui(self):
        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Left pane - Folder tree (match FolderPanel styling)
        self._folder_model = QStandardItemModel()
        self._folder_tree = QTreeView()
        self._folder_tree.setModel(self._folder_model)
        self._folder_tree.setHeaderHidden(True)
        self._folder_tree.setRootIsDecorated(True)
        self._folder_tree.setIndentation(16)
        self._folder_tree.setFrameShape(QFrame.Shape.NoFrame)
        self._folder_tree.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self._folder_tree.clicked.connect(self._on_folder_clicked)
        self._folder_tree.setMinimumWidth(150)
        self._folder_tree.setMaximumWidth(250)

        # Match FolderPanel palette (white background, black text)
        palette = self._folder_tree.palette()
        palette.setColor(QPalette.ColorRole.Base, Qt.GlobalColor.white)
        palette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.black)
        self._folder_tree.setPalette(palette)

        # Use app font
        self._folder_tree.setFont(QApplication.font())

        # Match FolderPanel item spacing
        self._folder_tree.setStyleSheet("""
            QTreeView::item {
                padding: 4px 0px;
            }
        """)
        layout.addWidget(self._folder_tree)

        # Right pane - Search + Contact list
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(0, 0, 0, 0)
        right_layout.setSpacing(0)

        # Toolbar with search and new contact button
        toolbar = QWidget()
        toolbar.setStyleSheet("background-color: #f5f5f5; border-bottom: 1px solid #d0d0d0;")
        toolbar_layout = QHBoxLayout(toolbar)
        toolbar_layout.setContentsMargins(8, 4, 8, 4)
        toolbar_layout.setSpacing(8)

        # Search box
        self._search_entry = QLineEdit()
        self._search_entry.setPlaceholderText("Search contacts...")
        self._search_entry.textChanged.connect(self._on_search_changed)
        self._search_entry.setClearButtonEnabled(True)
        toolbar_layout.addWidget(self._search_entry, 1)

        # New Contact button
        new_btn = QPushButton("+ New Contact")
        new_btn.clicked.connect(self.newContactRequested.emit)
        toolbar_layout.addWidget(new_btn)

        right_layout.addWidget(toolbar)

        # Contact list
        self._contact_model = ContactListModel()
        self._contact_delegate = ContactDelegate()
        self._contact_list = QListView()
        self._contact_list.setModel(self._contact_model)
        self._contact_list.setItemDelegate(self._contact_delegate)
        self._contact_list.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self._contact_list.setSpacing(1)
        self._contact_list.setFrameShape(QFrame.Shape.NoFrame)
        self._contact_list.clicked.connect(self._on_contact_clicked)
        self._contact_list.doubleClicked.connect(self._on_contact_double_clicked)
        self._contact_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._contact_list.customContextMenuRequested.connect(self._show_contact_context_menu)
        right_layout.addWidget(self._contact_list, 1)

        # Empty state label (hidden by default)
        self._empty_label = QLabel("No contacts")
        self._empty_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._empty_label.setStyleSheet("color: gray; font-size: 14pt;")
        self._empty_label.hide()
        right_layout.addWidget(self._empty_label)

        layout.addWidget(right_widget, 1)

    def set_account(self, account_id: int, email: str = ""):
        """Set the current account and display header."""
        self._current_account_id = account_id
        self._user_domain = email.split('@')[1] if '@' in email else ""

        # Create account header in folder tree
        self._folder_model.clear()
        self._folder_items.clear()

        account_item = QStandardItem(email or "Contacts")
        account_item.setData(("account", account_id), Qt.ItemDataRole.UserRole)
        account_item.setSelectable(False)
        font = account_item.font()
        font.setBold(True)
        account_item.setFont(font)
        self._folder_model.invisibleRootItem().appendRow(account_item)
        self._folder_tree.expand(self._folder_model.indexFromItem(account_item))

        # Add "All Contacts" virtual folder
        all_item = QStandardItem("All Contacts")
        all_item.setData(("folder", account_id, "__all__", "All Contacts"), Qt.ItemDataRole.UserRole)
        all_item.setIcon(QIcon.fromTheme("contact-new",
            QApplication.style().standardIcon(QStyle.StandardPixmap.SP_FileIcon)))
        account_item.appendRow(all_item)
        self._folder_items["__all__"] = all_item

        # Add "Global Address List" virtual folder (server users)
        gal_item = QStandardItem("Global Address List")
        gal_item.setData(("folder", account_id, "__company__", "Global Address List"), Qt.ItemDataRole.UserRole)
        gal_item.setIcon(QIcon.fromTheme("system-users",
            QApplication.style().standardIcon(QStyle.StandardPixmap.SP_DirIcon)))
        account_item.appendRow(gal_item)
        self._folder_items["__company__"] = gal_item

    def add_folder(self, folder_id: str, name: str, is_default: bool = False):
        """Add a contact folder to the tree."""
        if not self._current_account_id:
            return

        key = folder_id
        if key in self._folder_items:
            # Update existing
            self._folder_items[key].setText(name)
            return

        # Find account item
        account_item = self._folder_model.item(0)
        if not account_item:
            return

        item = QStandardItem(name)
        item.setData(("folder", self._current_account_id, folder_id, name), Qt.ItemDataRole.UserRole)
        item.setIcon(QIcon.fromTheme("x-office-address-book",
            QApplication.style().standardIcon(QStyle.StandardPixmap.SP_DirIcon)))
        account_item.appendRow(item)
        self._folder_items[key] = item

        # Auto-select default folder
        if is_default:
            self._folder_tree.setCurrentIndex(self._folder_model.indexFromItem(item))

    def clear_folders(self):
        """Clear all folders."""
        self._folder_model.clear()
        self._folder_items.clear()
        self._current_folder_id = None

    def set_contacts(self, contacts: list):
        """Set the contact list data."""
        self._contact_model.set_contacts(contacts)

        # Show/hide empty label
        if contacts:
            self._empty_label.hide()
            self._contact_list.show()
        else:
            self._contact_list.hide()
            self._empty_label.show()

    def clear_contacts(self):
        """Clear the contact list."""
        self._contact_model.clear()
        self._contact_list.hide()
        self._empty_label.show()

    def select_contact(self, item_id: str):
        """Select a contact by item_id."""
        index = self._contact_model.find_index(item_id)
        if index.isValid():
            self._contact_list.setCurrentIndex(index)

    def _on_folder_clicked(self, index):
        """Handle folder selection."""
        item = self._folder_model.itemFromIndex(index)
        if item:
            data = item.data(Qt.ItemDataRole.UserRole)
            if data and data[0] == "folder":
                self._current_folder_id = data[2]
                self.folderSelected.emit(data[1], data[2], data[3])

    def _on_contact_clicked(self, index):
        """Handle contact selection."""
        if not index.isValid():
            return
        contact = self._contact_model.get_contact(index.row())
        if contact and self._current_account_id:
            self.contactSelected.emit(self._current_account_id, contact.get("item_id", ""))

    def _on_contact_double_clicked(self, index):
        """Handle contact double-click (edit)."""
        if not index.isValid():
            return
        contact = self._contact_model.get_contact(index.row())
        if contact and self._current_account_id:
            self.contactDoubleClicked.emit(self._current_account_id, contact.get("item_id", ""))

    def _on_search_changed(self, text: str):
        """Handle search text changes."""
        self._contact_model.set_filter(text)

    def get_selected_contact(self) -> Optional[dict]:
        """Get the currently selected contact data."""
        index = self._contact_list.currentIndex()
        if index.isValid():
            return self._contact_model.get_contact(index.row())
        return None

    def current_folder_id(self) -> Optional[str]:
        """Get the currently selected folder ID."""
        return self._current_folder_id

    def update_font(self):
        """Update font to match application font."""
        self._folder_tree.setFont(QApplication.font())
        self._contact_list.setFont(QApplication.font())

    def _show_contact_context_menu(self, pos):
        """Show context menu for contact list."""
        index = self._contact_list.indexAt(pos)
        if not index.isValid():
            return

        contact = self._contact_model.get_contact(index.row())
        if not contact:
            return

        # Get primary email
        emails = contact.get("email_addresses", [])
        if isinstance(emails, str):
            try:
                emails = json.loads(emails)
            except (json.JSONDecodeError, TypeError):
                emails = []

        menu = QMenu(self)

        # New Message action (only if contact has email)
        if emails:
            first_email = emails[0]
            email = first_email.get("address", "") if isinstance(first_email, dict) else str(first_email)

            name = contact.get("common_name", "")
            if not name:
                first = contact.get("first_name", "") or ""
                last = contact.get("last_name", "") or ""
                name = f"{first} {last}".strip()

            new_msg_action = menu.addAction("New Message")
            new_msg_action.triggered.connect(lambda: self.newMessageToContact.emit(email, name))

        if menu.actions():
            menu.exec(self._contact_list.mapToGlobal(pos))
