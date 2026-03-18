"""Folder panel using QTreeView with clean styling."""

from typing import Optional
from PySide6.QtCore import Qt, Signal, QMimeData
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QTreeView, QFrame, QAbstractItemView
)
from PySide6.QtGui import QStandardItemModel, QStandardItem, QIcon
from PySide6.QtWidgets import QStyle, QApplication


class FolderPanel(QWidget):
    """Folder panel using standard QTreeView."""

    folderSelected = Signal(int, str, str)
    messagesDropped = Signal(QMimeData, int, str)
    contextMenuRequested = Signal(object, int, str, str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._current_account_id: Optional[int] = None
        self._folder_items: dict[str, QStandardItem] = {}
        self._folders_group: Optional[QStandardItem] = None
        self._account_item: Optional[QStandardItem] = None

        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        self._model = QStandardItemModel()
        self._tree = QTreeView()
        self._tree.setModel(self._model)
        self._tree.setHeaderHidden(True)
        self._tree.setRootIsDecorated(True)
        self._tree.setIndentation(16)
        self._tree.setFrameShape(QFrame.Shape.NoFrame)
        self._tree.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self._tree.setDragDropMode(QAbstractItemView.DragDropMode.DropOnly)
        self._tree.setAcceptDrops(True)
        self._tree.clicked.connect(self._on_clicked)
        self._tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._tree.customContextMenuRequested.connect(self._on_context_menu)

        # Use the same palette as the rest of the app
        from PySide6.QtGui import QPalette
        palette = self._tree.palette()
        palette.setColor(QPalette.ColorRole.Base, Qt.GlobalColor.white)
        palette.setColor(QPalette.ColorRole.Text, Qt.GlobalColor.black)
        self._tree.setPalette(palette)

        # Use app font
        self._tree.setFont(QApplication.font())

        # Add vertical spacing
        self._tree.setStyleSheet("""
            QTreeView::item {
                padding: 4px 0px;
            }
        """)

        layout.addWidget(self._tree)

    def set_account(self, account_id: int, name: str, email: str = "", connected: bool = False):
        self._current_account_id = account_id

        # Create account header item
        display = f"{email} {'(connected)' if connected else ''}"
        if self._account_item:
            self._account_item.setText(display)
        else:
            self._account_item = QStandardItem(display)
            self._account_item.setData(("account", account_id), Qt.ItemDataRole.UserRole)
            self._account_item.setSelectable(False)
            self._account_item.setIcon(QIcon.fromTheme("mail-message",
                QApplication.style().standardIcon(QStyle.StandardPixmap.SP_ComputerIcon)))
            # Make it bold
            font = self._account_item.font()
            font.setBold(True)
            self._account_item.setFont(font)
            self._model.invisibleRootItem().appendRow(self._account_item)
            # Expand by default
            self._tree.expand(self._model.indexFromItem(self._account_item))

    def add_folder(self, account_id: int, folder_id: str, name: str,
                   unread_count: int = 0, is_base_folder: bool = True):
        key = f"{account_id}:{folder_id}"

        # Update existing
        if key in self._folder_items:
            item = self._folder_items[key]
            display = f"{name} ({unread_count})" if unread_count > 0 else name
            item.setText(display)
            return

        # Determine parent - all folders go under account item
        if is_base_folder:
            # Base folders go directly under account
            parent = self._account_item if self._account_item else self._model.invisibleRootItem()
        else:
            # Other folders go under "Folders" group
            if not self._folders_group:
                self._folders_group = QStandardItem("Folders")
                self._folders_group.setData(("section",), Qt.ItemDataRole.UserRole)
                self._folders_group.setSelectable(False)
                self._folders_group.setIcon(QIcon.fromTheme("folder-open",
                    QApplication.style().standardIcon(QStyle.StandardPixmap.SP_DirOpenIcon)))
                group_parent = self._account_item if self._account_item else self._model.invisibleRootItem()
                group_parent.appendRow(self._folders_group)
            parent = self._folders_group

        display = f"{name} ({unread_count})" if unread_count > 0 else name
        item = QStandardItem(display)
        item.setData(("folder", account_id, folder_id, name), Qt.ItemDataRole.UserRole)
        item.setIcon(self._get_folder_icon(name.lower()))
        parent.appendRow(item)
        self._folder_items[key] = item

    def _get_folder_icon(self, folder_name: str):
        """Get icon for folder based on name using system theme."""
        if 'inbox' in folder_name:
            icon = QIcon.fromTheme("mail-inbox")
        elif 'sent' in folder_name:
            icon = QIcon.fromTheme("mail-sent")
        elif 'draft' in folder_name:
            icon = QIcon.fromTheme("mail-drafts")
        elif 'spam' in folder_name or 'junk' in folder_name:
            icon = QIcon.fromTheme("mail-mark-junk")
        elif 'trash' in folder_name or 'deleted' in folder_name:
            icon = QIcon.fromTheme("user-trash")
        else:
            icon = QIcon.fromTheme("folder")

        # Fallback to Qt standard icon if theme icon not found
        if icon.isNull():
            icon = QApplication.style().standardIcon(QStyle.StandardPixmap.SP_DirIcon)
        return icon

    def clear_folders(self, account_id: int = None):
        # Clear folder items but preserve account item
        self._folder_items.clear()
        self._folders_group = None

        if self._account_item:
            # Remove all children of account item (the folders)
            self._account_item.removeRows(0, self._account_item.rowCount())
        else:
            # No account item, just clear the model
            self._model.clear()

    def clear_all(self):
        """Clear everything including account item."""
        self._model.clear()
        self._folder_items.clear()
        self._folders_group = None
        self._account_item = None

    def select_folder(self, account_id: int, folder_id: str):
        key = f"{account_id}:{folder_id}"
        if key in self._folder_items:
            item = self._folder_items[key]
            index = self._model.indexFromItem(item)
            self._tree.setCurrentIndex(index)

    def update_unread_count(self, account_id: int, folder_id: str, count: int):
        key = f"{account_id}:{folder_id}"
        if key in self._folder_items:
            item = self._folder_items[key]
            data = item.data(Qt.ItemDataRole.UserRole)
            if data:
                name = data[3]
                display = f"{name} ({count})" if count > 0 else name
                item.setText(display)

    def _on_clicked(self, index):
        item = self._model.itemFromIndex(index)
        if item:
            data = item.data(Qt.ItemDataRole.UserRole)
            if data and data[0] == "folder":
                self.folderSelected.emit(data[1], data[2], data[3])

    def _on_context_menu(self, position):
        index = self._tree.indexAt(position)
        if index.isValid():
            item = self._model.itemFromIndex(index)
            if item:
                data = item.data(Qt.ItemDataRole.UserRole)
                if data and data[0] == "folder":
                    self.contextMenuRequested.emit(
                        self._tree.mapToGlobal(position),
                        data[1], data[2], data[3]
                    )

    def set_folders_expanded(self, expanded: bool):
        if self._folders_group:
            index = self._model.indexFromItem(self._folders_group)
            if expanded:
                self._tree.expand(index)
            else:
                self._tree.collapse(index)

    def is_folders_expanded(self) -> bool:
        if self._folders_group:
            index = self._model.indexFromItem(self._folders_group)
            return self._tree.isExpanded(index)
        return True

    def set_account_expanded(self, expanded: bool):
        """Expand or collapse the account section."""
        if self._account_item:
            index = self._model.indexFromItem(self._account_item)
            if expanded:
                self._tree.expand(index)
            else:
                self._tree.collapse(index)

    def is_account_expanded(self) -> bool:
        """Check if account section is expanded."""
        if self._account_item:
            index = self._model.indexFromItem(self._account_item)
            return self._tree.isExpanded(index)
        return True

    def update_font(self):
        """Update font to match application font."""
        self._tree.setFont(QApplication.font())
