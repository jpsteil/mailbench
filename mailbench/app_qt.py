"""Main application window using PySide6."""

import sys
import os
import tempfile
import subprocess
import platform
from datetime import datetime, date, timedelta
from typing import Optional, Any
import threading

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QSplitter, QTreeView, QListView, QTextEdit, QLabel, QLineEdit,
    QToolBar, QStatusBar, QMenu, QMenuBar, QMessageBox, QFrame,
    QStyledItemDelegate, QStyle, QAbstractItemView, QSizePolicy,
    QPushButton, QDialog, QFileDialog, QGridLayout, QStackedWidget
)
try:
    from PySide6.QtWebEngineWidgets import QWebEngineView
    HAS_WEBENGINE = True
except ImportError:
    HAS_WEBENGINE = False
from PySide6.QtCore import (
    Qt, QAbstractListModel, QModelIndex, QSize, Signal, Slot,
    QTimer, QThread, QObject, QSettings, QMetaObject, Q_ARG, QGenericArgument,
    QEvent
)
from PySide6.QtGui import (
    QFont, QFontMetrics, QPainter, QColor, QPen, QBrush, QAction,
    QStandardItemModel, QStandardItem, QPalette, QIcon, QKeySequence,
    QShortcut, QPixmap
)

from mailbench.database import Database
from mailbench.kerio_client import KerioConnectionPool, SyncManager, KerioConfig
from mailbench.version import __version__


# Icons (Unicode)
ICON_FLAG = "\u2691"  # Filled flag
ICON_FLAG_EMPTY = "\u2690"  # Empty/hollow flag
ICON_REPLY = "\u21b5"
ICON_FORWARD = "\u21b7"
ICON_ATTACH = "\U0001F4CE"


class MessageData:
    """Container for message data."""
    def __init__(self, data: dict):
        self.item_id = data.get('item_id', '')
        self.sender_name = data.get('sender_name', '')
        self.sender_email = data.get('sender_email', '')
        self.subject = data.get('subject', '') or '(No Subject)'
        self.date_received = data.get('date_received', '')
        self.size = data.get('size', 0)
        self.is_read = data.get('is_read', True)
        self.is_flagged = data.get('is_flagged', False)
        self.is_answered = data.get('is_answered', False)
        self.is_forwarded = data.get('is_forwarded', False)
        self.has_attachments = data.get('has_attachments', False)
        self._raw = data

    @property
    def sender_display(self) -> str:
        return self.sender_name or self.sender_email or 'Unknown'

    @property
    def date_display(self) -> str:
        """Format date for display."""
        try:
            if not self.date_received or len(self.date_received) < 15:
                return self.date_received or ""
            date_part = self.date_received[:8]
            time_part = self.date_received[9:15]
            dt = datetime.strptime(f"{date_part}{time_part}", "%Y%m%d%H%M%S")

            today = date.today()
            msg_date = dt.date()

            if msg_date == today:
                hour = dt.hour % 12 or 12
                return f"{hour}:{dt.strftime('%M')} {dt.strftime('%p')}"
            elif msg_date == today - timedelta(days=1):
                return "Yesterday"
            else:
                return f"{dt.month}/{dt.day}/{dt.year}"
        except Exception:
            return self.date_received or ""

    @property
    def size_display(self) -> str:
        """Format size for display."""
        if not self.size:
            return ""
        if self.size < 1024:
            return f"{self.size} B"
        elif self.size < 1024 * 1024:
            return f"{self.size / 1024:.1f} kB"
        else:
            return f"{self.size / (1024 * 1024):.1f} MB"

    @property
    def icons(self) -> str:
        """Get status icons string."""
        result = ""
        if self.is_flagged:
            result += ICON_FLAG + " "
        if self.is_answered:
            result += ICON_REPLY + " "
        elif self.is_forwarded:
            result += ICON_FORWARD + " "
        if self.has_attachments:
            result += "*"
        return result.strip()


class MessageListModel(QAbstractListModel):
    """Model for the message list."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._messages: list[MessageData] = []
        self._message_map: dict[str, int] = {}  # item_id -> index
        self._filtered_messages: list[MessageData] = []
        self._filter_text = ""

    def rowCount(self, parent=QModelIndex()) -> int:
        return len(self._filtered_messages)

    def data(self, index: QModelIndex, role: int = Qt.ItemDataRole.DisplayRole) -> Any:
        if not index.isValid() or index.row() >= len(self._filtered_messages):
            return None

        msg = self._filtered_messages[index.row()]

        if role == Qt.ItemDataRole.DisplayRole:
            return msg
        elif role == Qt.ItemDataRole.UserRole:
            return msg.item_id

        return None

    def clear(self):
        """Clear all messages."""
        self.beginResetModel()
        self._messages.clear()
        self._filtered_messages.clear()
        self._message_map.clear()
        self.endResetModel()

    def set_filter(self, text: str):
        """Filter messages by text."""
        self._filter_text = text.lower().strip()
        self._apply_filter()

    def _apply_filter(self):
        """Apply current filter to messages."""
        self.beginResetModel()
        if not self._filter_text:
            self._filtered_messages = list(self._messages)
        else:
            self._filtered_messages = [
                msg for msg in self._messages
                if (self._filter_text in msg.sender_display.lower() or
                    self._filter_text in msg.subject.lower() or
                    self._filter_text in msg.sender_email.lower())
            ]
        self.endResetModel()

    def add_messages(self, messages: list[dict]):
        """Add multiple messages."""
        if not messages:
            return

        new_messages = []
        for msg_data in messages:
            item_id = msg_data.get('item_id')
            if item_id and item_id not in self._message_map:
                new_messages.append(MessageData(msg_data))

        if new_messages:
            for msg in new_messages:
                self._message_map[msg.item_id] = len(self._messages)
                self._messages.append(msg)
            self._apply_filter()

    def get_message(self, item_id: str) -> Optional[MessageData]:
        """Get message by item_id."""
        idx = self._message_map.get(item_id)
        if idx is not None and idx < len(self._messages):
            return self._messages[idx]
        return None

    def get_message_at(self, index: int) -> Optional[MessageData]:
        """Get message at filtered index."""
        if 0 <= index < len(self._filtered_messages):
            return self._filtered_messages[index]
        return None

    def update_message(self, item_id: str, is_read: bool = None, is_flagged: bool = None):
        """Update message state."""
        idx = self._message_map.get(item_id)
        if idx is not None and idx < len(self._messages):
            if is_read is not None:
                self._messages[idx].is_read = is_read
            if is_flagged is not None:
                self._messages[idx].is_flagged = is_flagged
            # Find in filtered list
            for i, msg in enumerate(self._filtered_messages):
                if msg.item_id == item_id:
                    model_index = self.index(i)
                    self.dataChanged.emit(model_index, model_index)
                    break

    def remove_message(self, item_id: str):
        """Remove a message."""
        idx = self._message_map.get(item_id)
        if idx is not None:
            del self._messages[idx]
            del self._message_map[item_id]
            # Rebuild map for shifted indices
            for i, msg in enumerate(self._messages[idx:], start=idx):
                self._message_map[msg.item_id] = i
            self._apply_filter()


class MessageDelegate(QStyledItemDelegate):
    """Custom delegate for Kerio-style message rows."""

    FLAG_WIDTH = 24  # Left column for clickable flag
    DATE_WIDTH = 90
    PADDING = 8
    RIGHT_PADDING = 28

    # Signal emitted when flag icon is clicked (item_id, new_flagged_state)
    flagClicked = Signal(str, bool)

    def __init__(self, parent=None, dark_mode=False, font_size=12):
        super().__init__(parent)
        self.dark_mode = dark_mode
        self.font_size = font_size
        self._row_height = max(48, int(font_size * 4))  # Scale row height with font
        self._setup_colors()

    def _setup_colors(self):
        """Setup colors based on theme."""
        if self.dark_mode:
            self.bg_color = QColor("#2b2b2b")
            self.bg_unread = QColor("#353535")
            self.fg_color = QColor("#a9b7c6")
            self.fg_bold = QColor("#ffffff")
            self.secondary = QColor("#888888")
            self.select_bg = QColor("#214283")
            self.border_color = QColor("#404040")
        else:
            self.bg_color = QColor("#ffffff")
            self.bg_unread = QColor("#f0f4f8")
            self.fg_color = QColor("#000000")
            self.fg_bold = QColor("#000000")
            self.secondary = QColor("#666666")
            self.select_bg = QColor("#0078d4")
            self.border_color = QColor("#e0e0e0")

    def set_dark_mode(self, dark: bool):
        """Update dark mode setting."""
        self.dark_mode = dark
        self._setup_colors()

    def set_font_size(self, size: int):
        """Update font size."""
        self.font_size = size
        self._row_height = max(48, int(size * 4))

    def sizeHint(self, option, index) -> QSize:
        return QSize(option.rect.width(), self._row_height)

    def paint(self, painter: QPainter, option, index: QModelIndex):
        msg: MessageData = index.data(Qt.ItemDataRole.DisplayRole)
        if not msg:
            return

        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        rect = option.rect
        is_selected = option.state & QStyle.StateFlag.State_Selected

        # Background
        if is_selected:
            painter.fillRect(rect, self.select_bg)
        elif not msg.is_read:
            painter.fillRect(rect, self.bg_unread)
        else:
            painter.fillRect(rect, self.bg_color)

        # Text colors
        if is_selected:
            sender_color = Qt.GlobalColor.white
            subject_color = Qt.GlobalColor.white
            meta_color = QColor("#cccccc")
        else:
            sender_color = self.fg_bold if not msg.is_read else self.fg_color
            subject_color = self.secondary
            meta_color = self.secondary

        # Fonts - scale based on font_size setting
        sender_font = painter.font()
        if not msg.is_read:
            sender_font.setBold(True)
        sender_font.setPointSize(self.font_size)

        subject_font = painter.font()
        subject_font.setPointSize(max(8, self.font_size - 1))

        meta_font = painter.font()
        meta_font.setPointSize(max(8, self.font_size - 1))

        small_font = painter.font()
        small_font.setPointSize(max(7, self.font_size - 2))

        # Calculate positions
        x = rect.x() + self.PADDING
        y = rect.y()
        h = rect.height()

        # Left column - clickable flag (always visible)
        flag_x = x
        flag_icon = ICON_FLAG if msg.is_flagged else ICON_FLAG_EMPTY
        flag_color = QColor("#e6a100") if msg.is_flagged else meta_color
        painter.setFont(meta_font)
        painter.setPen(QPen(flag_color))
        painter.drawText(flag_x, y, self.FLAG_WIDTH, h,
                        Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignCenter, flag_icon)

        # Right column - date on top, indicators below (fixed width)
        right_col_width = self.DATE_WIDTH + 30
        date_x = rect.x() + rect.width() - self.RIGHT_PADDING - right_col_width

        # Date (top right)
        painter.setFont(meta_font)
        painter.setPen(QPen(meta_color))
        painter.drawText(date_x, y + 6, right_col_width, 20,
                        Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop, msg.date_display)

        # Indicators (bottom right) - attachment, reply/forward (not flag, it's on left)
        painter.setFont(small_font)
        indicators = ""
        if msg.is_answered:
            indicators += ICON_REPLY + " "
        elif msg.is_forwarded:
            indicators += ICON_FORWARD + " "
        if msg.has_attachments:
            indicators += ICON_ATTACH
        if indicators:
            painter.drawText(date_x, y + 26, right_col_width, 20,
                            Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignTop, indicators.strip())

        # Sender and subject (middle)
        left_x = x + self.FLAG_WIDTH + 4
        left_w = date_x - left_x - 8

        # Sender
        painter.setFont(sender_font)
        painter.setPen(QPen(sender_color))
        sender_text = msg.sender_display
        fm = QFontMetrics(sender_font)
        if fm.horizontalAdvance(sender_text) > left_w:
            sender_text = fm.elidedText(sender_text, Qt.TextElideMode.ElideRight, left_w)
        painter.drawText(left_x, y + 6, left_w, 20,
                        Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop, sender_text)

        # Subject
        painter.setFont(subject_font)
        painter.setPen(QPen(subject_color))
        subject_text = msg.subject
        fm = QFontMetrics(subject_font)
        if fm.horizontalAdvance(subject_text) > left_w:
            subject_text = fm.elidedText(subject_text, Qt.TextElideMode.ElideRight, left_w)
        painter.drawText(left_x, y + 26, left_w, 20,
                        Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop, subject_text)

        # Bottom border
        painter.setPen(QPen(self.border_color))
        painter.drawLine(rect.bottomLeft(), rect.bottomRight())

        painter.restore()

    def editorEvent(self, event, model, option, index) -> bool:
        """Handle mouse events - detect clicks on flag area."""
        if event.type() == QEvent.Type.MouseButtonRelease:
            # Check if click was in flag area
            rect = option.rect
            flag_x = rect.x() + self.PADDING
            flag_end = flag_x + self.FLAG_WIDTH

            click_x = event.position().x()
            if flag_x <= click_x <= flag_end:
                msg: MessageData = index.data(Qt.ItemDataRole.DisplayRole)
                if msg:
                    # Toggle flag state
                    new_flagged = not msg.is_flagged
                    self.flagClicked.emit(msg.item_id, new_flagged)
                    return True

        return super().editorEvent(event, model, option, index)


class FolderTreeModel(QStandardItemModel):
    """Model for the folder tree."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setHorizontalHeaderLabels(["Accounts"])
        self._account_items: dict[int, QStandardItem] = {}
        self._folder_items: dict[str, QStandardItem] = {}  # "account_id:folder_id" -> item
        self._folders_group_items: dict[int, QStandardItem] = {}  # account_id -> "Folders" group item

    def add_account(self, account_id: int, name: str, connected: bool = False):
        """Add an account to the tree."""
        if account_id in self._account_items:
            item = self._account_items[account_id]
            item.setText(f"{name} (connected)" if connected else name)
            return item

        item = QStandardItem(f"{name} (connected)" if connected else name)
        item.setData(("account", account_id), Qt.ItemDataRole.UserRole)
        item.setEditable(False)
        self.appendRow(item)
        self._account_items[account_id] = item
        return item

    def add_folder(self, account_id: int, folder_id: str, name: str, unread_count: int = 0, is_base_folder: bool = True):
        """Add a folder under an account or under the Folders group."""
        if account_id not in self._account_items:
            return None

        key = f"{account_id}:{folder_id}"
        if key in self._folder_items:
            item = self._folder_items[key]
            display = f"{name} ({unread_count})" if unread_count > 0 else name
            item.setText(display)
            return item

        # Determine parent - base folders go under account, others under "Folders" group
        if is_base_folder:
            parent = self._account_items[account_id]
        else:
            parent = self._get_or_create_folders_group(account_id)

        display = f"{name} ({unread_count})" if unread_count > 0 else name
        item = QStandardItem(display)
        item.setData(("folder", account_id, folder_id), Qt.ItemDataRole.UserRole)
        item.setEditable(False)

        # Set icon for base folders
        if is_base_folder:
            icon = self._get_folder_icon(name.lower())
            if icon:
                item.setIcon(icon)

        parent.appendRow(item)
        self._folder_items[key] = item
        return item

    def _get_folder_icon(self, folder_name: str) -> Optional[QIcon]:
        """Get icon for a folder based on its name."""
        if 'inbox' in folder_name:
            return QIcon.fromTheme("mail-inbox", QIcon.fromTheme("folder"))
        elif 'sent' in folder_name:
            return QIcon.fromTheme("mail-sent", QIcon.fromTheme("folder"))
        elif 'draft' in folder_name:
            return QIcon.fromTheme("mail-drafts", QIcon.fromTheme("folder"))
        elif 'spam' in folder_name or 'junk' in folder_name:
            return QIcon.fromTheme("mail-mark-junk", QIcon.fromTheme("folder"))
        elif 'trash' in folder_name or 'deleted' in folder_name:
            return QIcon.fromTheme("user-trash", QIcon.fromTheme("folder"))
        return None

    def _get_or_create_folders_group(self, account_id: int) -> QStandardItem:
        """Get or create the 'Folders' group item for an account."""
        if account_id in self._folders_group_items:
            return self._folders_group_items[account_id]

        parent = self._account_items[account_id]
        group_item = QStandardItem("Folders")
        group_item.setData(("folders_group", account_id), Qt.ItemDataRole.UserRole)
        group_item.setEditable(False)
        parent.appendRow(group_item)
        self._folders_group_items[account_id] = group_item
        return group_item

    def get_folders_group_item(self, account_id: int) -> Optional[QStandardItem]:
        """Get the Folders group item for an account."""
        return self._folders_group_items.get(account_id)

    def clear_folders(self, account_id: int):
        """Clear folders for an account."""
        if account_id not in self._account_items:
            return
        parent = self._account_items[account_id]
        parent.removeRows(0, parent.rowCount())
        # Remove from folder map
        keys_to_remove = [k for k in self._folder_items if k.startswith(f"{account_id}:")]
        for k in keys_to_remove:
            del self._folder_items[k]
        # Remove folders group reference
        if account_id in self._folders_group_items:
            del self._folders_group_items[account_id]


class MessageWindow(QMainWindow):
    """Separate window for viewing a message."""

    def __init__(self, parent, msg_data: dict, full_data: dict = None, dark_mode: bool = False):
        super().__init__(parent)
        self.msg_data = msg_data
        self.full_data = full_data or {}
        self._dark_mode = dark_mode

        subject = msg_data.get('subject', '(No Subject)')
        self.setWindowTitle(subject)
        self.resize(800, 600)

        self._create_ui()
        self._apply_theme()
        self._restore_geometry()

    def _restore_geometry(self):
        """Restore window geometry from settings."""
        settings = QSettings("Mailbench", "Mailbench")
        geometry = settings.value("message_window_geometry")
        if geometry:
            self.restoreGeometry(geometry)

    def closeEvent(self, event):
        """Save geometry on close."""
        settings = QSettings("Mailbench", "Mailbench")
        settings.setValue("message_window_geometry", self.saveGeometry())
        event.accept()

    def _create_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(10, 10, 10, 10)

        # Header - fixed size, doesn't stretch
        sender_name = self.msg_data.get('sender_name', '')
        sender_email = self.msg_data.get('sender_email', '')
        if sender_name and sender_email:
            sender = f"{sender_name} <{sender_email}>"
        else:
            sender = sender_name or sender_email or 'Unknown'

        from_label = QLabel(f"From: {sender}")
        from_label.setFont(QFont("", 10, QFont.Weight.Bold))
        from_label.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        layout.addWidget(from_label)

        to_str = self.full_data.get('to', '')
        to_label = QLabel(f"To: {to_str}")
        to_label.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        to_label.setWordWrap(True)
        layout.addWidget(to_label)

        cc_str = self.full_data.get('cc', '')
        if cc_str:
            cc_label = QLabel(f"CC: {cc_str}")
            cc_label.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
            cc_label.setWordWrap(True)
            layout.addWidget(cc_label)

        subject = self.msg_data.get('subject', '(No Subject)')
        subject_label = QLabel(f"Subject: {subject}")
        subject_label.setFont(QFont("", 10, QFont.Weight.Bold))
        subject_label.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        layout.addWidget(subject_label)

        # Date
        date = self.msg_data.get('date_received', '')
        if date and len(date) >= 15:
            try:
                date_part = date[:8]
                time_part = date[9:15]
                dt = datetime.strptime(f"{date_part}{time_part}", "%Y%m%d%H%M%S")
                date = dt.strftime("%A, %B %d, %Y at %I:%M %p")
            except Exception:
                pass
        date_label = QLabel(date)
        date_label.setStyleSheet("color: gray;")
        date_label.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        layout.addWidget(date_label)

        # Separator
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        layout.addWidget(line)

        # Body - use WebEngine if available, stretches to fill space
        body = self.full_data.get('body', '')
        body_type = self.full_data.get('body_type', 'text')

        if HAS_WEBENGINE:
            self.body_text = QWebEngineView()
            if body_type == 'html':
                if not body.lower().strip().startswith('<!doctype') and not body.lower().strip().startswith('<html'):
                    body = f"<html><body>{body}</body></html>"
                self.body_text.setHtml(body)
            else:
                escaped_body = body.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('\n', '<br>')
                self.body_text.setHtml(f"<html><body><pre style='white-space: pre-wrap; font-family: sans-serif;'>{escaped_body}</pre></body></html>")
        else:
            self.body_text = QTextEdit()
            self.body_text.setReadOnly(True)
            if body_type == 'html':
                self.body_text.setHtml(body)
            else:
                self.body_text.setPlainText(body)
        layout.addWidget(self.body_text, 1)  # stretch factor = 1

    def _apply_theme(self):
        if self._dark_mode:
            self.setStyleSheet("""
                QMainWindow, QWidget { background-color: #2b2b2b; color: #a9b7c6; }
                QTextEdit { background-color: #313335; color: #a9b7c6; border: 1px solid #3c3f41; }
                QLabel { color: #a9b7c6; }
            """)
        else:
            self.setStyleSheet("")


class MailbenchWindow(QMainWindow):
    """Main application window."""

    # Signal for thread-safe callback execution
    _invoke_callback = Signal(object)

    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"Mailbench v{__version__}")
        self.resize(1200, 800)

        # Connect signal for thread-safe callbacks
        self._invoke_callback.connect(self._execute_callback, Qt.ConnectionType.QueuedConnection)

        # Initialize backend
        self.db = Database()
        self.kerio_pool = KerioConnectionPool()
        self.sync_manager = SyncManager(self.kerio_pool, self.db, self)

        # Track state
        self.connected_accounts: set[int] = set()
        self._current_account_id: Optional[int] = None
        self._address_book: list[dict] = []  # For email autocomplete
        self._signatures: dict[int, str] = {}  # account_id -> signature
        self._current_folder_id: Optional[str] = None
        self._current_message_id: Optional[str] = None
        self._messages_by_id: dict[str, dict] = {}
        self._current_message_full_data: Optional[dict] = None
        self._message_windows: list[MessageWindow] = []

        # Theme setting (System/Light/Dark)
        self._theme_setting = self.db.get_setting("theme", "System")
        self._dark_mode = self._detect_dark_mode()

        # Font size from settings
        self.font_size = int(self.db.get_setting("font_size", "12"))

        # Preview zoom level (50-200%, default 100)
        self._preview_zoom = int(self.db.get_setting("preview_zoom", "100"))

        # Build UI
        self._create_menu()
        self._create_toolbar()
        self._create_main_layout()
        self._create_statusbar()

        # Keyboard shortcuts
        self._create_shortcuts()

        # Apply theme, font, and zoom
        self._apply_theme()
        self._apply_font_size()
        self._apply_preview_zoom()

        # Restore geometry
        self._restore_geometry()

        # Load accounts
        QTimer.singleShot(100, self._load_accounts)

    @Slot(object)
    def _execute_callback(self, callback):
        """Execute a callback on the main thread."""
        if callable(callback):
            callback()

    def after(self, ms: int, callback):
        """Tkinter-compatible method for scheduling callbacks on main thread.

        The SyncManager uses root.after() - this provides Qt compatibility.
        Uses a signal to safely cross thread boundaries.
        """
        # Emit signal to execute callback on main thread (thread-safe)
        self._invoke_callback.emit(callback)

    def _create_shortcuts(self):
        """Create keyboard shortcuts."""
        # Delete key
        delete_shortcut = QShortcut(QKeySequence(Qt.Key.Key_Delete), self)
        delete_shortcut.activated.connect(self._delete_messages)

        # Ctrl+N for new message
        new_shortcut = QShortcut(QKeySequence("Ctrl+N"), self)
        new_shortcut.activated.connect(self._new_message)

        # F5 for refresh
        refresh_shortcut = QShortcut(QKeySequence(Qt.Key.Key_F5), self)
        refresh_shortcut.activated.connect(self._check_mail)

        # Ctrl+R for reply
        reply_shortcut = QShortcut(QKeySequence("Ctrl+R"), self)
        reply_shortcut.activated.connect(self._reply)

        # Ctrl+Shift+R for reply all
        reply_all_shortcut = QShortcut(QKeySequence("Ctrl+Shift+R"), self)
        reply_all_shortcut.activated.connect(self._reply_all)

        # Ctrl+F for forward
        forward_shortcut = QShortcut(QKeySequence("Ctrl+Shift+F"), self)
        forward_shortcut.activated.connect(self._forward)

        # Escape to clear filter
        escape_shortcut = QShortcut(QKeySequence(Qt.Key.Key_Escape), self)
        escape_shortcut.activated.connect(self._clear_filter)

        # Zoom shortcuts
        zoom_in_shortcut = QShortcut(QKeySequence("Ctrl+="), self)
        zoom_in_shortcut.activated.connect(self._preview_zoom_in)
        zoom_in_shortcut2 = QShortcut(QKeySequence("Ctrl++"), self)
        zoom_in_shortcut2.activated.connect(self._preview_zoom_in)
        zoom_out_shortcut = QShortcut(QKeySequence("Ctrl+-"), self)
        zoom_out_shortcut.activated.connect(self._preview_zoom_out)
        zoom_reset_shortcut = QShortcut(QKeySequence("Ctrl+0"), self)
        zoom_reset_shortcut.activated.connect(self._preview_zoom_reset)

    def _create_menu(self):
        """Create the menu bar."""
        menubar = self.menuBar()

        # File menu
        file_menu = menubar.addMenu("&File")

        new_action = QAction("&New Message", self)
        new_action.setShortcut("Ctrl+N")
        new_action.triggered.connect(self._new_message)
        file_menu.addAction(new_action)

        refresh_action = QAction("&Check Mail", self)
        refresh_action.setShortcut("F5")
        refresh_action.triggered.connect(self._check_mail)
        file_menu.addAction(refresh_action)

        file_menu.addSeparator()

        self._dark_mode_action = QAction("&Dark Mode", self)
        self._dark_mode_action.setCheckable(True)
        self._dark_mode_action.setChecked(self._dark_mode)
        self._dark_mode_action.triggered.connect(self._toggle_dark_mode)
        file_menu.addAction(self._dark_mode_action)

        settings_action = QAction("&Settings...", self)
        settings_action.triggered.connect(self._show_settings)
        file_menu.addAction(settings_action)

        file_menu.addSeparator()

        exit_action = QAction("E&xit", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        # Accounts menu
        accounts_menu = menubar.addMenu("&Accounts")

        add_account_action = QAction("&Add Account...", self)
        add_account_action.triggered.connect(self._add_account)
        accounts_menu.addAction(add_account_action)

        manage_accounts_action = QAction("&Manage Accounts...", self)
        manage_accounts_action.triggered.connect(self._manage_accounts)
        accounts_menu.addAction(manage_accounts_action)

        # Help menu
        help_menu = menubar.addMenu("&Help")

        about_action = QAction("&About", self)
        about_action.triggered.connect(self._show_about)
        help_menu.addAction(about_action)

    def _create_toolbar(self):
        """Create the toolbar."""
        toolbar = QToolBar("Main Toolbar")
        toolbar.setMovable(False)
        toolbar.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        toolbar.setIconSize(QSize(20, 20))
        self.addToolBar(toolbar)

        # Create actions with icons
        new_action = QAction(QIcon.fromTheme("mail-message-new", QIcon.fromTheme("document-new")), "New", self)
        new_action.triggered.connect(self._new_message)
        new_action.setToolTip("Compose new message (Ctrl+N)")
        toolbar.addAction(new_action)

        toolbar.addSeparator()

        self._reply_action = QAction(QIcon.fromTheme("mail-reply-sender", QIcon.fromTheme("go-previous")), "Reply", self)
        self._reply_action.triggered.connect(self._reply)
        self._reply_action.setToolTip("Reply to sender (Ctrl+R)")
        toolbar.addAction(self._reply_action)

        self._reply_all_action = QAction(QIcon.fromTheme("mail-reply-all"), "Reply All", self)
        self._reply_all_action.triggered.connect(self._reply_all)
        self._reply_all_action.setToolTip("Reply to all (Ctrl+Shift+R)")
        toolbar.addAction(self._reply_all_action)

        forward_action = QAction(QIcon.fromTheme("mail-forward", QIcon.fromTheme("go-next")), "Forward", self)
        forward_action.triggered.connect(self._forward)
        forward_action.setToolTip("Forward message (Ctrl+Shift+F)")
        toolbar.addAction(forward_action)

        toolbar.addSeparator()

        delete_action = QAction(QIcon.fromTheme("edit-delete", QIcon.fromTheme("user-trash")), "Delete", self)
        delete_action.triggered.connect(self._delete_messages)
        delete_action.setToolTip("Delete selected messages (Delete)")
        toolbar.addAction(delete_action)

        refresh_action = QAction(QIcon.fromTheme("view-refresh", QIcon.fromTheme("sync")), "Refresh", self)
        refresh_action.triggered.connect(self._check_mail)
        refresh_action.setToolTip("Check for new mail (F5)")
        toolbar.addAction(refresh_action)

    def _create_main_layout(self):
        """Create the 3-pane layout."""
        central = QWidget()
        self.setCentralWidget(central)

        layout = QHBoxLayout(central)
        layout.setContentsMargins(5, 5, 5, 5)

        # Main splitter
        self._splitter = QSplitter(Qt.Orientation.Horizontal)
        layout.addWidget(self._splitter)

        # Left pane - Folder tree
        folder_widget = QWidget()
        folder_layout = QVBoxLayout(folder_widget)
        folder_layout.setContentsMargins(0, 0, 0, 0)
        folder_layout.setSpacing(0)

        self._folder_model = FolderTreeModel()
        self._folder_tree = QTreeView()
        self._folder_tree.setModel(self._folder_model)
        self._folder_tree.setHeaderHidden(True)
        self._folder_tree.setAnimated(True)
        self._folder_tree.setIndentation(12)  # Reduce indent depth
        self._folder_tree.clicked.connect(self._on_folder_clicked)
        self._folder_tree.doubleClicked.connect(self._on_folder_double_clicked)
        self._folder_tree.expanded.connect(self._on_folder_expanded)
        self._folder_tree.collapsed.connect(self._on_folder_collapsed)
        folder_layout.addWidget(self._folder_tree)

        self._splitter.addWidget(folder_widget)

        # Right splitter (message list | preview)
        self._right_splitter = QSplitter(Qt.Orientation.Horizontal)
        self._splitter.addWidget(self._right_splitter)

        # Middle pane - Message list
        message_widget = QWidget()
        message_layout = QVBoxLayout(message_widget)
        message_layout.setContentsMargins(0, 0, 0, 0)

        # Filter box
        filter_layout = QHBoxLayout()
        filter_layout.setContentsMargins(8, 4, 4, 4)
        filter_layout.addWidget(QLabel("Filter:"))
        self._filter_entry = QLineEdit()
        self._filter_entry.textChanged.connect(self._filter_messages)
        self._filter_entry.setPlaceholderText("Type to filter...")
        filter_layout.addWidget(self._filter_entry)
        message_layout.addLayout(filter_layout)

        # Message list
        self._message_model = MessageListModel()
        self._message_delegate = MessageDelegate(dark_mode=self._dark_mode, font_size=self.font_size)
        self._message_delegate.flagClicked.connect(self._on_flag_clicked)

        self._message_list = QListView()
        self._message_list.setModel(self._message_model)
        self._message_list.setItemDelegate(self._message_delegate)
        self._message_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self._message_list.setUniformItemSizes(True)
        self._message_list.setSpacing(0)
        self._message_list.clicked.connect(self._on_message_clicked)
        self._message_list.doubleClicked.connect(self._on_message_double_clicked)
        self._message_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._message_list.customContextMenuRequested.connect(self._show_context_menu)
        # Handle keyboard navigation
        self._message_list.selectionModel().currentChanged.connect(self._on_message_selection_changed)
        message_layout.addWidget(self._message_list)

        self._right_splitter.addWidget(message_widget)

        # Right pane - Stacked widget for Preview/Compose
        self._preview_stack = QStackedWidget()

        # Index 0: Preview widget
        self._preview_widget = QWidget()
        preview_layout = QVBoxLayout(self._preview_widget)
        preview_layout.setContentsMargins(5, 5, 5, 5)
        preview_layout.setSpacing(2)

        # Preview header (fixed size area)
        header_widget = QWidget()
        header_layout = QVBoxLayout(header_widget)
        header_layout.setContentsMargins(0, 0, 0, 5)
        header_layout.setSpacing(2)

        self._preview_from = QLabel("From:")
        self._preview_from.setStyleSheet("font-weight: bold;")
        header_layout.addWidget(self._preview_from)

        self._preview_to = QLabel("To:")
        header_layout.addWidget(self._preview_to)

        self._preview_subject = QLabel("")
        self._preview_subject.setStyleSheet("font-weight: bold; font-size: 12pt;")
        self._preview_subject.setWordWrap(True)
        header_layout.addWidget(self._preview_subject)

        self._preview_date = QLabel("")
        self._preview_date.setStyleSheet("color: gray;")
        header_layout.addWidget(self._preview_date)

        # Attachments row
        self._attachments_layout = QHBoxLayout()
        header_layout.addLayout(self._attachments_layout)

        header_widget.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        preview_layout.addWidget(header_widget)

        # Separator
        self._preview_separator = QFrame()
        self._preview_separator.setFrameShape(QFrame.Shape.HLine)
        self._preview_separator.setFixedHeight(2)
        self._preview_separator.setObjectName("previewSeparator")
        preview_layout.addWidget(self._preview_separator)

        # Preview body - use WebEngine if available for proper HTML/image support
        if HAS_WEBENGINE:
            self._preview_body = QWebEngineView()
            self._preview_body.setHtml("<html><body></body></html>")
            self._use_webengine = True
        else:
            self._preview_body = QTextEdit()
            self._preview_body.setReadOnly(True)
            self._use_webengine = False

        # Body takes remaining space
        self._preview_body.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        preview_layout.addWidget(self._preview_body, 1)  # stretch factor = 1

        # Install event filter for Ctrl+scroll zoom
        self._preview_body.installEventFilter(self)

        self._preview_stack.addWidget(self._preview_widget)  # Index 0

        # Index 1: Compose widget (created on demand)
        self._compose_widget = None

        self._right_splitter.addWidget(self._preview_stack)

        # Set splitter sizes
        self._splitter.setSizes([200, 1000])
        self._right_splitter.setSizes([400, 600])

    def _create_statusbar(self):
        """Create the status bar."""
        self._statusbar = QStatusBar()
        self.setStatusBar(self._statusbar)
        self._statusbar.showMessage("Ready")

    def _apply_theme(self):
        """Apply dark or light theme."""
        if self._dark_mode:
            self.setStyleSheet("""
                QMainWindow, QWidget { background-color: #2b2b2b; color: #a9b7c6; }
                QTreeView, QListView, QTextEdit { background-color: #313335; color: #a9b7c6; border: 1px solid #3c3f41; }
                QTreeView::item { padding: 4px 2px; }
                QTreeView::item:selected, QListView::item:selected { background-color: #214283; }
                QTreeView::item:hover:!selected { background-color: #3c3f41; }
                QTreeView::branch { background-color: #313335; }
                QLineEdit { background-color: #313335; color: #a9b7c6; border: 1px solid #3c3f41; padding: 4px; }
                QToolBar { background-color: #2b2b2b; border: none; border-bottom: 1px solid #3c3f41; padding: 4px; spacing: 2px; }
                QToolBar QToolButton { background-color: transparent; color: #a9b7c6; padding: 6px 12px; border: none; border-radius: 4px; }
                QToolBar QToolButton:hover { background-color: #3c3f41; }
                QToolBar QToolButton:pressed { background-color: #4c5052; }
                QToolBar::separator { background-color: #3c3f41; width: 1px; margin: 4px 8px; }
                QMenuBar { background-color: #2b2b2b; color: #a9b7c6; }
                QMenuBar::item:selected { background-color: #214283; }
                QMenu { background-color: #2b2b2b; color: #a9b7c6; border: 1px solid #3c3f41; }
                QMenu::item:selected { background-color: #214283; }
                QStatusBar { background-color: #2b2b2b; color: #a9b7c6; }
                QSplitter::handle { background-color: #3c3f41; }
                QLabel { color: #a9b7c6; }
                QPushButton { background-color: #3c3f41; color: #a9b7c6; padding: 5px 10px; border: none; }
                QPushButton:hover { background-color: #4c5052; }
                QFrame#previewSeparator { background-color: #505050; }
            """)
        else:
            self.setStyleSheet("""
                QMainWindow, QWidget { background-color: #f0f0f0; color: #000000; }
                QTreeView, QListView, QTextEdit { background-color: #ffffff; color: #000000; border: 1px solid #c0c0c0; }
                QTreeView::item { padding: 4px 2px; }
                QTreeView::item:selected, QListView::item:selected { background-color: #0078d4; color: white; }
                QTreeView::item:hover:!selected { background-color: #e8e8e8; }
                QLineEdit { background-color: #ffffff; color: #000000; border: 1px solid #c0c0c0; padding: 4px; }
                QToolBar { background-color: #f5f5f5; border: none; border-bottom: 1px solid #d0d0d0; padding: 4px; spacing: 2px; }
                QToolBar QToolButton { background-color: transparent; color: #333333; padding: 6px 12px; border: none; border-radius: 4px; }
                QToolBar QToolButton:hover { background-color: #e0e0e0; }
                QToolBar QToolButton:pressed { background-color: #d0d0d0; }
                QToolBar::separator { background-color: #d0d0d0; width: 1px; margin: 4px 8px; }
                QSplitter::handle { background-color: #c0c0c0; }
                QFrame#previewSeparator { background-color: #c0c0c0; }
            """)

        # Update delegate
        self._message_delegate.set_dark_mode(self._dark_mode)
        self._message_list.viewport().update()

    def _apply_font_size(self):
        """Apply font size to UI elements."""
        # Set application-wide font
        font = QApplication.font()
        font.setPointSize(self.font_size)
        QApplication.setFont(font)

        # Update message delegate
        self._message_delegate.set_font_size(self.font_size)
        self._message_list.viewport().update()

    def eventFilter(self, obj, event):
        """Filter events for Ctrl+scroll zoom on child widgets."""
        if event.type() == event.Type.Wheel:
            if event.modifiers() == Qt.KeyboardModifier.ControlModifier:
                delta = event.angleDelta().y()
                if delta > 0:
                    self._preview_zoom_in()
                elif delta < 0:
                    self._preview_zoom_out()
                return True  # Event handled
        return super().eventFilter(obj, event)

    def _preview_zoom_in(self):
        """Increase preview zoom level."""
        if self._preview_zoom < 200:
            self._preview_zoom = min(200, self._preview_zoom + 10)
            self._apply_preview_zoom()

    def _preview_zoom_out(self):
        """Decrease preview zoom level."""
        if self._preview_zoom > 50:
            self._preview_zoom = max(50, self._preview_zoom - 10)
            self._apply_preview_zoom()

    def _preview_zoom_reset(self):
        """Reset preview zoom to 100%."""
        self._preview_zoom = 100
        self._apply_preview_zoom()

    def _apply_preview_zoom(self):
        """Apply current zoom level to preview and compose."""
        # Save zoom level
        self.db.set_setting("preview_zoom", str(self._preview_zoom))

        # Apply to preview body
        zoom_factor = self._preview_zoom / 100.0
        if HAS_WEBENGINE and isinstance(self._preview_body, QWebEngineView):
            self._preview_body.setZoomFactor(zoom_factor)
        else:
            # For QTextEdit, scale font
            font = self._preview_body.font()
            font.setPointSize(int(12 * zoom_factor))
            self._preview_body.setFont(font)

        # Apply to compose widget if open
        if self._compose_widget is not None:
            self._compose_widget.set_zoom(zoom_factor)

        # Update status bar
        self._statusbar.showMessage(f"Zoom: {self._preview_zoom}%", 2000)

    def _detect_dark_mode(self) -> bool:
        """Detect if dark mode should be used based on theme setting."""
        if self._theme_setting == "Dark":
            return True
        elif self._theme_setting == "Light":
            return False
        else:  # System
            # Detect system theme
            app = QApplication.instance()
            if app:
                palette = app.palette()
                # If window background is darker than text, it's dark mode
                bg_lightness = palette.color(QPalette.ColorRole.Window).lightness()
                return bg_lightness < 128
            return False

    def _apply_theme_setting(self, theme: str):
        """Apply a theme setting (System/Light/Dark)."""
        self._theme_setting = theme
        self._dark_mode = self._detect_dark_mode()
        self._dark_mode_action.setChecked(self._dark_mode)
        self._apply_theme()

    def _toggle_dark_mode(self):
        """Toggle dark mode (switches between Light and Dark, ignoring System)."""
        self._dark_mode = not self._dark_mode
        # Update theme setting to match
        self._theme_setting = "Dark" if self._dark_mode else "Light"
        self.db.set_setting("theme", self._theme_setting)
        self._dark_mode_action.setChecked(self._dark_mode)
        self._apply_theme()

    def _save_geometry(self):
        """Save window geometry and splitter positions."""
        settings = QSettings("Mailbench", "Mailbench")
        settings.setValue("geometry", self.saveGeometry())
        settings.setValue("splitter", self._splitter.saveState())
        settings.setValue("right_splitter", self._right_splitter.saveState())

    def _restore_geometry(self):
        """Restore window geometry and splitter positions."""
        settings = QSettings("Mailbench", "Mailbench")
        geometry = settings.value("geometry")
        if geometry:
            self.restoreGeometry(geometry)
        splitter = settings.value("splitter")
        if splitter:
            self._splitter.restoreState(splitter)
        right_splitter = settings.value("right_splitter")
        if right_splitter:
            self._right_splitter.restoreState(right_splitter)

    # ==================== Account/Folder Methods ====================

    def _load_accounts(self):
        """Load accounts from database."""
        accounts = self.db.get_accounts()
        for account in accounts:
            self._folder_model.add_account(account['id'], account['email'])

        # Auto-connect first account
        if accounts:
            self._connect_account(accounts[0]['id'])

    def _connect_account(self, account_id: int):
        """Connect to an account."""
        account = self.db.get_account(account_id)
        if not account:
            return

        self._statusbar.showMessage(f"Connecting to {account['email']}...")

        config = KerioConfig(
            email=account['email'],
            username=account['username'],
            password=account['password'],
            server=account['server']
        )

        def do_connect():
            try:
                self.kerio_pool.connect(account_id, config)
                self.after(0, lambda: self._on_connected(account_id, account['email'], True, None))
            except Exception as e:
                self.after(0, lambda: self._on_connected(account_id, account['email'], False, str(e)))

        threading.Thread(target=do_connect, daemon=True).start()

    def _on_connected(self, account_id: int, email: str, success: bool, error: str):
        """Handle connection result."""
        if success:
            self.connected_accounts.add(account_id)
            self._folder_model.add_account(account_id, email, connected=True)
            self._statusbar.showMessage(f"Connected to {email}")
            self._load_folders(account_id)
            self._load_address_book(account_id)
            self._load_signature(account_id)
        else:
            self._statusbar.showMessage(f"Connection failed: {error}")
            QMessageBox.warning(self, "Connection Failed", str(error))

    def _load_address_book(self, account_id: int):
        """Load contacts and users for email autocomplete."""
        self._address_book = []

        def on_users(success, error, users):
            if success and users:
                for u in users:
                    self._address_book.append({
                        'name': u.get('name', ''),
                        'email': u.get('email', ''),
                        'type': 'user'
                    })

        def on_contacts(success, error, contacts):
            if success and contacts:
                for c in contacts:
                    self._address_book.append({
                        'name': c.get('name', ''),
                        'email': c.get('email', ''),
                        'type': 'contact'
                    })
            # Also add cached emails from sent messages
            cached = self.db.get_cached_emails()
            for entry in cached:
                email = entry.get('email', '')
                # Check if already in address book
                if email and not any(a['email'].lower() == email.lower() for a in self._address_book):
                    self._address_book.append({
                        'name': entry.get('name', ''),
                        'email': email,
                        'type': 'recent',
                        'send_count': entry.get('send_count', 0)
                    })

        self.sync_manager.fetch_users(account_id, callback=on_users)
        self.sync_manager.fetch_contacts(account_id, callback=on_contacts)

    def _load_signature(self, account_id: int):
        """Load email signature from Kerio server."""
        def on_signature(success, error, signature):
            if success and signature:
                self._signatures[account_id] = signature

        self.sync_manager.fetch_signature(account_id, callback=on_signature)

    def _load_folders(self, account_id: int):
        """Load folders for an account."""
        def on_folders_synced(success, error=None):
            if success:
                # Read folders from database (sync_folders saves them there)
                folders = self.db.get_folders(account_id)
                self._folder_model.clear_folders(account_id)

                # Sort folders: standard folders first in specific order
                standard_order = ['inbox', 'sent', 'drafts', 'junk', 'spam', 'trash']
                # Folders to hide (non-mail folders and system folders)
                hidden_folders = [
                    'archive', 'public folders', 'public',
                    'calendar', 'contacts', 'notes', 'tasks',
                    'suggested contacts', 'infocenter', 'outbox'
                ]

                def folder_sort_key(f):
                    name = f['name'].lower()
                    # Handle "Deleted Items" as Trash
                    if name == 'deleted items':
                        name = 'trash'
                    # Check for standard folders
                    for i, std in enumerate(standard_order):
                        if std in name:
                            return (0, i, f['name'])
                    # Other folders come after, sorted alphabetically
                    return (1, 0, f['name'].lower())

                sorted_folders = sorted(folders, key=folder_sort_key)

                seen_trash = False
                for folder in sorted_folders:
                    name = folder['name'].lower()

                    # Skip hidden folders
                    if any(h in name for h in hidden_folders):
                        continue

                    # Skip folders starting with ~ (user root folders)
                    if folder['name'].startswith('~'):
                        continue

                    # Skip duplicate trash folder (Deleted Items = Trash)
                    if 'deleted items' in name or 'trash' in name:
                        if seen_trash:
                            continue
                        seen_trash = True

                    # Determine if this is a base folder
                    is_base = any(std in name for std in standard_order) or 'deleted items' in name

                    self._folder_model.add_folder(
                        account_id,
                        folder['folder_id'],
                        folder['name'],
                        folder.get('unread_count', 0),
                        is_base_folder=is_base
                    )

                # Expand account node
                for i in range(self._folder_model.rowCount()):
                    idx = self._folder_model.index(i, 0)
                    self._folder_tree.expand(idx)

                # Restore "Folders" group expanded state
                settings = QSettings("Mailbench", "Mailbench")
                folders_expanded = settings.value(f"folders_expanded_{account_id}", True, type=bool)
                folders_group = self._folder_model.get_folders_group_item(account_id)
                if folders_group:
                    folders_idx = self._folder_model.indexFromItem(folders_group)
                    if folders_expanded:
                        self._folder_tree.expand(folders_idx)
                    else:
                        self._folder_tree.collapse(folders_idx)

                # Select inbox
                self._select_inbox(account_id)
            else:
                self._statusbar.showMessage(f"Failed to load folders: {error}")

        self.sync_manager.sync_folders(account_id, callback=on_folders_synced)

    def _select_inbox(self, account_id: int):
        """Select the inbox folder."""
        # Find inbox in the folder items
        for key, item in self._folder_model._folder_items.items():
            if key.startswith(f"{account_id}:"):
                data = item.data(Qt.ItemDataRole.UserRole)
                if data and len(data) >= 3:
                    folder_name = item.text().lower()
                    if 'inbox' in folder_name:
                        # Select this item
                        idx = self._folder_model.indexFromItem(item)
                        self._folder_tree.setCurrentIndex(idx)
                        self._load_messages(account_id, data[2])
                        break

    def _on_folder_clicked(self, index: QModelIndex):
        """Handle folder selection."""
        item = self._folder_model.itemFromIndex(index)
        if not item:
            return

        data = item.data(Qt.ItemDataRole.UserRole)
        if not data:
            return

        if data[0] == "folder":
            account_id = data[1]
            folder_id = data[2]
            self._load_messages(account_id, folder_id)

    def _on_folder_double_clicked(self, index: QModelIndex):
        """Handle folder double-click (connect if account)."""
        item = self._folder_model.itemFromIndex(index)
        if not item:
            return

        data = item.data(Qt.ItemDataRole.UserRole)
        if data and data[0] == "account":
            account_id = data[1]
            if account_id not in self.connected_accounts:
                self._connect_account(account_id)

    def _on_folder_expanded(self, index: QModelIndex):
        """Save expanded state for Folders group."""
        item = self._folder_model.itemFromIndex(index)
        if item:
            data = item.data(Qt.ItemDataRole.UserRole)
            if data and data[0] == "folders_group":
                account_id = data[1]
                settings = QSettings("Mailbench", "Mailbench")
                settings.setValue(f"folders_expanded_{account_id}", True)

    def _on_folder_collapsed(self, index: QModelIndex):
        """Save collapsed state for Folders group."""
        item = self._folder_model.itemFromIndex(index)
        if item:
            data = item.data(Qt.ItemDataRole.UserRole)
            if data and data[0] == "folders_group":
                account_id = data[1]
                settings = QSettings("Mailbench", "Mailbench")
                settings.setValue(f"folders_expanded_{account_id}", False)

    # ==================== Message Methods ====================

    def _load_messages(self, account_id: int, folder_id: str):
        """Load messages for a folder."""
        self._current_account_id = account_id
        self._current_folder_id = folder_id
        self._message_model.clear()
        self._messages_by_id.clear()
        self._filter_entry.clear()

        if account_id not in self.connected_accounts:
            # Show cached messages
            messages = self.db.get_messages(account_id, folder_id, limit=50)
            self._display_messages(list(messages))
            return

        self._statusbar.showMessage("Loading messages...")

        def on_messages_synced(success, error=None, messages_data=None):
            if success and messages_data:
                self._display_messages(messages_data)
                self._statusbar.showMessage(f"{len(messages_data)} messages")
            elif error:
                self._statusbar.showMessage(f"Error: {error}")

        self.sync_manager.sync_messages(account_id, folder_id, callback=on_messages_synced)

    def _display_messages(self, messages: list[dict]):
        """Display messages in the list."""
        for msg in messages:
            item_id = msg.get('item_id')
            if item_id:
                self._messages_by_id[item_id] = msg

        self._message_model.add_messages(messages)

        # Auto-select first message
        if self._message_model.rowCount() > 0:
            first_index = self._message_model.index(0)
            self._message_list.setCurrentIndex(first_index)
            self._on_message_clicked(first_index)

    def _filter_messages(self, text: str):
        """Filter messages by search text."""
        self._message_model.set_filter(text)

    def _clear_filter(self):
        """Clear the filter."""
        self._filter_entry.clear()

    def _on_message_selection_changed(self, current: QModelIndex, previous: QModelIndex):
        """Handle message selection change (keyboard navigation)."""
        if current.isValid():
            self._on_message_clicked(current)

    def _on_message_clicked(self, index: QModelIndex):
        """Handle message selection."""
        msg: MessageData = index.data(Qt.ItemDataRole.DisplayRole)
        if not msg:
            return

        if msg.item_id == self._current_message_id:
            return

        self._current_message_id = msg.item_id
        self._show_message_preview(msg)

        # Fetch full body
        if self._current_account_id:
            self.sync_manager.fetch_message_body(
                self._current_account_id,
                msg.item_id,
                callback=self._on_body_fetched
            )

    def _on_message_double_clicked(self, index: QModelIndex):
        """Handle message double-click - open in new window."""
        msg: MessageData = index.data(Qt.ItemDataRole.DisplayRole)
        if not msg:
            return

        msg_data = self._messages_by_id.get(msg.item_id, {})
        full_data = self._current_message_full_data if msg.item_id == self._current_message_id else {}

        win = MessageWindow(self, msg_data, full_data, self._dark_mode)
        win.show()
        self._message_windows.append(win)

    def _show_message_preview(self, msg: MessageData):
        """Show message in preview pane."""
        sender = msg.sender_name
        if msg.sender_email:
            sender += f" <{msg.sender_email}>" if sender else msg.sender_email
        self._preview_from.setText(f"From: {sender}")
        self._preview_to.setText("To: Loading...")
        self._preview_subject.setText(msg.subject)
        self._preview_date.setText(msg.date_display)

        # Show loading message
        if self._use_webengine:
            self._preview_body.setHtml("<html><body><p>Loading...</p></body></html>")
        else:
            self._preview_body.setText("Loading...")

        # Clear attachments
        self._clear_attachments_display()

    def _on_body_fetched(self, success: bool, data: dict, error: str = None):
        """Handle message body fetch completion."""
        if not success or not data:
            if self._use_webengine:
                self._preview_body.setHtml(f"<html><body><p>Error loading message: {error}</p></body></html>")
            else:
                self._preview_body.setText(f"Error loading message: {error}")
            return

        # Store full data
        self._current_message_full_data = data

        # Update To field
        recipients = data.get('to', [])
        if recipients:
            if isinstance(recipients, list):
                to_str = ", ".join(r.get('address', '') if isinstance(r, dict) else str(r) for r in recipients[:3])
                if len(recipients) > 3:
                    to_str += f" (+{len(recipients) - 3} more)"
            else:
                to_str = str(recipients)
            self._preview_to.setText(f"To: {to_str}")

        # Show attachments
        attachments = data.get('attachments', [])
        self._show_attachments(attachments)

        # Show body
        body = data.get('body', '')
        body_type = data.get('body_type', 'text')

        if self._use_webengine:
            # WebEngineView handles HTML natively
            if body_type == 'html':
                # Wrap in proper HTML structure if needed
                if not body.lower().strip().startswith('<!doctype') and not body.lower().strip().startswith('<html'):
                    body = f"<html><body>{body}</body></html>"
                self._preview_body.setHtml(body)
            else:
                # Convert plain text to HTML
                escaped_body = body.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('\n', '<br>')
                self._preview_body.setHtml(f"<html><body><pre style='white-space: pre-wrap; font-family: sans-serif;'>{escaped_body}</pre></body></html>")
        else:
            # QTextEdit fallback
            if body_type == 'html':
                self._preview_body.setHtml(body)
            else:
                self._preview_body.setPlainText(body)

        # Mark as read
        item_id = data.get('item_id')
        if item_id:
            self._message_model.update_message(item_id, is_read=True)

    def _clear_attachments_display(self):
        """Clear attachments display."""
        while self._attachments_layout.count():
            child = self._attachments_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

    def _show_attachments(self, attachments: list):
        """Show attachments in preview header."""
        self._clear_attachments_display()

        if not attachments:
            return

        self._attachments_layout.addWidget(QLabel("Attachments:"))

        for att in attachments:
            name = att.get('name', 'attachment')
            size = att.get('size', 0)

            if size >= 1024 * 1024:
                size_str = f"({size / (1024 * 1024):.1f} MB)"
            elif size >= 1024:
                size_str = f"({size / 1024:.1f} KB)"
            else:
                size_str = f"({size} B)" if size else ""

            btn = QPushButton(f"📎 {name} {size_str}")
            btn.clicked.connect(lambda checked, a=att: self._download_attachment(a))
            self._attachments_layout.addWidget(btn)

        self._attachments_layout.addStretch()

    def _download_attachment(self, attachment: dict):
        """Download an attachment."""
        name = attachment.get('name', 'attachment')
        url = attachment.get('url', '')

        if not url:
            QMessageBox.warning(self, "Error", "Attachment URL not available")
            return

        if not self._current_account_id:
            return

        session = self.kerio_pool.get_session(self._current_account_id)
        if not session:
            QMessageBox.warning(self, "Error", "Account not connected")
            return

        # Ask user what to do
        result = QMessageBox.question(
            self, "Download Attachment",
            f"What do you want to do with '{name}'?",
            QMessageBox.StandardButton.Open | QMessageBox.StandardButton.Save | QMessageBox.StandardButton.Cancel
        )

        full_url = f"https://{session.config.server}{url}"

        if result == QMessageBox.StandardButton.Open:
            self._open_attachment(full_url, name, session)
        elif result == QMessageBox.StandardButton.Save:
            save_path, _ = QFileDialog.getSaveFileName(self, "Save Attachment", name)
            if save_path:
                self._save_attachment(full_url, save_path, session)

    def _open_attachment(self, url: str, filename: str, session):
        """Download attachment to temp file and open."""
        def do_open():
            try:
                import requests
                cookies = {session.cookie_name: session.cookie_value} if hasattr(session, 'cookie_name') else {}
                headers = {"X-Token": session.token} if hasattr(session, 'token') else {}

                response = requests.get(url, cookies=cookies, headers=headers, verify=False, timeout=60)
                response.raise_for_status()

                _, ext = os.path.splitext(filename)
                with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as f:
                    f.write(response.content)
                    temp_path = f.name

                system = platform.system()
                if system == "Darwin":
                    subprocess.run(["open", temp_path])
                elif system == "Windows":
                    os.startfile(temp_path)
                else:
                    subprocess.run(["xdg-open", temp_path])

                self.after(0, lambda: self._statusbar.showMessage("Opened attachment"))
            except Exception as e:
                self.after(0, lambda err=str(e): QMessageBox.warning(self, "Error", f"Failed to open: {err}"))

        self._statusbar.showMessage("Downloading attachment...")
        threading.Thread(target=do_open, daemon=True).start()

    def _save_attachment(self, url: str, save_path: str, session):
        """Save attachment to file."""
        def do_save():
            try:
                import requests
                cookies = {session.cookie_name: session.cookie_value} if hasattr(session, 'cookie_name') else {}
                headers = {"X-Token": session.token} if hasattr(session, 'token') else {}

                response = requests.get(url, cookies=cookies, headers=headers, verify=False, timeout=60)
                response.raise_for_status()

                with open(save_path, 'wb') as f:
                    f.write(response.content)

                self.after(0, lambda p=save_path: self._statusbar.showMessage(f"Saved: {p}"))
            except Exception as e:
                self.after(0, lambda err=str(e): QMessageBox.warning(self, "Error", f"Failed to save: {err}"))

        self._statusbar.showMessage("Downloading attachment...")
        threading.Thread(target=do_save, daemon=True).start()

    # ==================== Context Menu ====================

    def _show_context_menu(self, position):
        """Show context menu for message list."""
        indexes = self._message_list.selectedIndexes()
        if not indexes:
            return

        menu = QMenu(self)

        # Open
        open_action = menu.addAction("Open")
        open_action.triggered.connect(lambda: self._on_message_double_clicked(indexes[0]))

        menu.addSeparator()

        # Reply options
        reply_action = menu.addAction("Reply")
        reply_action.triggered.connect(self._reply)
        reply_all_action = menu.addAction("Reply All")
        reply_all_action.triggered.connect(self._reply_all)
        forward_action = menu.addAction("Forward")
        forward_action.triggered.connect(self._forward)

        menu.addSeparator()

        # Mark as read/unread
        msg: MessageData = indexes[0].data(Qt.ItemDataRole.DisplayRole)
        if msg and msg.is_read:
            mark_action = menu.addAction("Mark as Unread")
            mark_action.triggered.connect(lambda: self._mark_messages(is_read=False))
        else:
            mark_action = menu.addAction("Mark as Read")
            mark_action.triggered.connect(lambda: self._mark_messages(is_read=True))

        # Flag/unflag
        if msg and msg.is_flagged:
            flag_action = menu.addAction("Remove Flag")
            flag_action.triggered.connect(lambda: self._toggle_flag_selected(False))
        else:
            flag_action = menu.addAction("Flag")
            flag_action.triggered.connect(lambda: self._toggle_flag_selected(True))

        menu.addSeparator()

        # Move to folder submenu
        if self._current_account_id:
            move_menu = menu.addMenu("Move to")
            folders = self.db.get_folders(self._current_account_id)
            for folder in folders:
                folder_name = folder.get('name', '')
                folder_id = folder.get('folder_id', '')
                if folder_id and folder_id != self._current_folder_id:
                    action = move_menu.addAction(folder_name)
                    action.triggered.connect(lambda checked, fid=folder_id: self._move_to_folder(fid))

        menu.addSeparator()

        # Delete
        count = len(indexes)
        delete_label = f"Delete ({count} messages)" if count > 1 else "Delete"
        delete_action = menu.addAction(delete_label)
        delete_action.triggered.connect(self._delete_messages)

        menu.exec_(self._message_list.mapToGlobal(position))

    def _mark_messages(self, is_read: bool):
        """Mark selected messages as read/unread."""
        if not self._current_account_id:
            return

        for index in self._message_list.selectedIndexes():
            msg: MessageData = index.data(Qt.ItemDataRole.DisplayRole)
            if msg:
                self._message_model.update_message(msg.item_id, is_read=is_read)
                self.sync_manager.mark_as_read(
                    self._current_account_id, msg.item_id, is_read,
                    callback=lambda s, e: None
                )

    def _toggle_flag_selected(self, flagged: bool):
        """Toggle flag on selected messages."""
        if not self._current_account_id:
            return

        for index in self._message_list.selectedIndexes():
            msg: MessageData = index.data(Qt.ItemDataRole.DisplayRole)
            if msg:
                self._message_model.update_message(msg.item_id, is_flagged=flagged)
                self.sync_manager.set_flag(
                    self._current_account_id, msg.item_id, flagged,
                    callback=lambda s, e: None
                )

    def _on_flag_clicked(self, item_id: str, flagged: bool):
        """Handle flag icon click from delegate."""
        if not self._current_account_id:
            return

        self._message_model.update_message(item_id, is_flagged=flagged)
        self.sync_manager.set_flag(
            self._current_account_id, item_id, flagged,
            callback=lambda s, e: None
        )

    def _move_to_folder(self, folder_id: str):
        """Move selected messages to folder."""
        if not self._current_account_id:
            return

        for index in self._message_list.selectedIndexes():
            msg: MessageData = index.data(Qt.ItemDataRole.DisplayRole)
            if msg:
                self.sync_manager.move_message(
                    self._current_account_id, msg.item_id, folder_id,
                    callback=lambda s, e, iid=msg.item_id: self._on_message_moved(s, e, iid)
                )

    def _on_message_moved(self, success: bool, error: str, item_id: str):
        """Handle message moved."""
        if success:
            self._message_model.remove_message(item_id)
        else:
            self._statusbar.showMessage(f"Move failed: {error}")

    # ==================== Compose (Inline) ====================

    def _show_compose(self, reply_to=None, forward=None):
        """Show the compose widget inline, replacing the preview pane."""
        from mailbench.views.compose_qt import ComposeWidget

        # Remove old compose widget if it exists
        if self._compose_widget:
            self._preview_stack.removeWidget(self._compose_widget)
            self._compose_widget.deleteLater()
            self._compose_widget = None

        # Get signature from Kerio server (cached on connect)
        signature = self._signatures.get(self._current_account_id, "") if self._current_account_id else ""

        # Create new compose widget
        self._compose_widget = ComposeWidget(
            self, self.db, self.sync_manager, self._current_account_id,
            reply_to=reply_to, forward=forward, signature=signature,
            font_size=self.font_size, zoom=self._preview_zoom
        )
        self._compose_widget.message_sent.connect(self._on_compose_done)
        self._compose_widget.compose_cancelled.connect(self._on_compose_done)

        # Set address book for autocomplete
        if self._address_book:
            self._compose_widget.set_address_book(self._address_book)

        # Add to stack and switch to it
        self._preview_stack.addWidget(self._compose_widget)
        self._preview_stack.setCurrentWidget(self._compose_widget)

        # Focus the appropriate field after widget is shown
        QTimer.singleShot(0, self._compose_widget.focus_to_field)

    def _on_compose_done(self):
        """Handle compose completed (sent or cancelled)."""
        # Switch back to preview
        self._preview_stack.setCurrentWidget(self._preview_widget)

        # Clean up compose widget
        if self._compose_widget:
            self._preview_stack.removeWidget(self._compose_widget)
            self._compose_widget.deleteLater()
            self._compose_widget = None

    def _new_message(self):
        """Create new message."""
        self._show_compose()

    def _reply(self):
        """Reply to message."""
        if not self._current_message_id or not self._current_message_full_data:
            return

        msg_data = self._messages_by_id.get(self._current_message_id, {})
        full_data = self._current_message_full_data

        reply_data = {
            'from_name': msg_data.get('sender_name', ''),
            'from_email': msg_data.get('sender_email', ''),
            'subject': msg_data.get('subject', ''),
            'date': msg_data.get('date_received', ''),
            'to': full_data.get('to', ''),
            'body': full_data.get('body', ''),
            'reply_all': False
        }

        self._show_compose(reply_to=reply_data)

    def _reply_all(self):
        """Reply all."""
        if not self._current_message_id or not self._current_message_full_data:
            return

        msg_data = self._messages_by_id.get(self._current_message_id, {})
        full_data = self._current_message_full_data

        reply_data = {
            'from_name': msg_data.get('sender_name', ''),
            'from_email': msg_data.get('sender_email', ''),
            'subject': msg_data.get('subject', ''),
            'date': msg_data.get('date_received', ''),
            'to': full_data.get('to', ''),
            'cc': full_data.get('cc', ''),
            'body': full_data.get('body', ''),
            'reply_all': True
        }

        self._show_compose(reply_to=reply_data)

    def _forward(self):
        """Forward message."""
        if not self._current_message_id or not self._current_message_full_data:
            return

        msg_data = self._messages_by_id.get(self._current_message_id, {})
        full_data = self._current_message_full_data

        forward_data = {
            'from_name': msg_data.get('sender_name', ''),
            'from_email': msg_data.get('sender_email', ''),
            'subject': msg_data.get('subject', ''),
            'date': msg_data.get('date_received', ''),
            'to': full_data.get('to', ''),
            'body': full_data.get('body', '')
        }

        self._show_compose(forward=forward_data)

    def _delete_messages(self):
        """Delete selected messages."""
        selection = self._message_list.selectedIndexes()
        if not selection:
            return

        if len(selection) > 1:
            if QMessageBox.question(
                self, "Delete Messages",
                f"Delete {len(selection)} selected messages?"
            ) != QMessageBox.StandardButton.Yes:
                return

        # Remember the row to select after deletion (for single delete)
        next_row = selection[0].row() if len(selection) == 1 else -1

        for index in selection:
            msg: MessageData = index.data(Qt.ItemDataRole.DisplayRole)
            if msg and self._current_account_id:
                self.sync_manager.delete_message(
                    self._current_account_id,
                    msg.item_id,
                    callback=lambda s, e, iid=msg.item_id, row=next_row: self._on_message_deleted(s, e, iid, row)
                )

    def _on_message_deleted(self, success: bool, error: str, item_id: str, select_row: int = -1):
        """Handle message deletion."""
        if success:
            self._message_model.remove_message(item_id)

            # Select next message
            if select_row >= 0:
                row_count = self._message_model.rowCount()
                if row_count > 0:
                    # If we were at the last item, select the new last item
                    new_row = min(select_row, row_count - 1)
                    new_index = self._message_model.index(new_row)
                    self._message_list.setCurrentIndex(new_index)
                    self._on_message_clicked(new_index)
                else:
                    # No messages left, clear preview
                    self._current_message_id = None
                    self._preview_from.setText("From:")
                    self._preview_to.setText("To:")
                    self._preview_subject.setText("")
                    self._preview_date.setText("")
                    if self._use_webengine:
                        self._preview_body.setHtml("<html><body></body></html>")
                    else:
                        self._preview_body.clear()
        else:
            self._statusbar.showMessage(f"Delete failed: {error}")

    def _check_mail(self):
        """Refresh current folder."""
        if self._current_account_id and self._current_folder_id:
            self._load_messages(self._current_account_id, self._current_folder_id)

    def _add_account(self):
        """Add new account."""
        from mailbench.dialogs.dialogs_qt import AccountDialog
        dialog = AccountDialog(self, self.db, self.kerio_pool, app=self)
        dialog.exec()
        self._refresh_accounts()

    def _manage_accounts(self):
        """Manage accounts."""
        from mailbench.dialogs.dialogs_qt import AccountDialog
        dialog = AccountDialog(self, self.db, self.kerio_pool, app=self)
        dialog.exec()
        self._refresh_accounts()

    def _refresh_accounts(self):
        """Refresh account list after changes."""
        # Clear and reload
        self._folder_model.clear()
        self._folder_model._account_items.clear()
        self._folder_model._folder_items.clear()
        self._load_accounts()

    def _show_settings(self):
        """Show settings dialog."""
        from mailbench.dialogs.dialogs_qt import SettingsDialog
        dialog = SettingsDialog(self, self.db, self)
        dialog.exec()

    def _show_about(self):
        """Show about dialog."""
        QMessageBox.about(
            self, "About Mailbench",
            f"Mailbench v{__version__}\n\nA Python email client for Kerio Connect.\n\n"
            f"Built with PySide6 (Qt6)"
        )

    def closeEvent(self, event):
        """Handle window close."""
        # Capture actual zoom from WebEngineView before closing
        if HAS_WEBENGINE and isinstance(self._preview_body, QWebEngineView):
            actual_zoom = int(self._preview_body.zoomFactor() * 100)
            if actual_zoom != self._preview_zoom:
                self._preview_zoom = actual_zoom
                self.db.set_setting("preview_zoom", str(self._preview_zoom))
        # Save geometry
        self._save_geometry()

        # Hide window and release input
        self.hide()
        self.releaseKeyboard()
        self.releaseMouse()

        # Signal threads to stop and close sessions
        self.sync_manager._shutdown = True
        self.kerio_pool.close_all()

        # Accept and let Qt clean up properly (may help XWayland)
        event.accept()
        QApplication.instance().quit()


def main():
    """Main entry point."""
    # Force X11/XWayland to avoid Wayland keyboard grab issues on exit
    import os
    if os.environ.get("XDG_SESSION_TYPE") == "wayland":
        os.environ["QT_QPA_PLATFORM"] = "xcb"

    app = QApplication(sys.argv)
    app.setApplicationName("Mailbench")
    app.setApplicationVersion(__version__)
    app.setOrganizationName("Mailbench")

    # Enable high DPI scaling
    app.setStyle("Fusion")

    window = MailbenchWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
