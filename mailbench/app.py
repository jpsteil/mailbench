"""Main application window using PySide6."""

import sys
import os
import re
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
    QPushButton, QDialog, QFileDialog, QGridLayout, QStackedWidget,
    QToolButton
)
try:
    from PySide6.QtWebEngineWidgets import QWebEngineView
    from PySide6.QtWebEngineCore import QWebEngineSettings
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


def block_remote_images(html: str) -> str:
    """Block remote images to prevent tracking pixels and IP leakage.

    Security: Remote images are commonly used for email tracking:
    - Tracking pixels (1x1 images) confirm when/if an email was opened
    - Remote images reveal your IP address to the sender
    - Some tracking pixels include unique IDs per recipient

    This function replaces remote image URLs with about:blank.
    Keeps data: URLs for inline images (these are already in the email).
    """
    if not html:
        return html

    # Block remote image sources (http, https, protocol-relative)
    # Keep data: URLs as they're inline images, not remote
    html = re.sub(
        r'(<img[^>]*\s+src\s*=\s*["\']?)\s*(https?://[^"\'>\s]*|//[^"\'>\s]*)',
        r'\1about:blank',
        html, flags=re.IGNORECASE
    )

    # Also block remote sources in srcset
    html = re.sub(
        r'(<img[^>]*\s+srcset\s*=\s*["\'])[^"\']*(["\'])',
        r'\1\2',
        html, flags=re.IGNORECASE
    )

    # Block remote background images in inline styles
    html = re.sub(
        r'(style\s*=\s*["\'][^"\']*background[^:]*:\s*url\s*\(\s*["\']?)\s*(https?://[^"\')\s]*|//[^"\')\s]*)',
        r'\1about:blank',
        html, flags=re.IGNORECASE
    )

    return html


def sanitize_html(html: str) -> str:
    """Sanitize HTML email content to prevent XSS attacks.

    Removes:
    - script, iframe, object, embed, link, meta, base, form tags
    - Event handlers (onclick, onerror, onload, etc.)
    - javascript: and data: URLs
    - style tags (can contain CSS expressions)
    """
    if not html:
        return html

    # Remove dangerous tags completely (including contents for script/style)
    # Tags whose content should be removed entirely
    for tag in ['script', 'style']:
        html = re.sub(
            rf'<{tag}[^>]*>.*?</{tag}>',
            '', html, flags=re.IGNORECASE | re.DOTALL
        )

    # Tags to remove (but keep content for some)
    dangerous_tags = ['iframe', 'object', 'embed', 'link', 'meta', 'base', 'form', 'input', 'button', 'textarea']
    for tag in dangerous_tags:
        # Remove opening tags
        html = re.sub(rf'<{tag}[^>]*>', '', html, flags=re.IGNORECASE)
        # Remove closing tags
        html = re.sub(rf'</{tag}>', '', html, flags=re.IGNORECASE)

    # Remove event handlers (on* attributes)
    html = re.sub(
        r'\s+on\w+\s*=\s*["\'][^"\']*["\']',
        '', html, flags=re.IGNORECASE
    )
    html = re.sub(
        r'\s+on\w+\s*=\s*[^\s>]+',
        '', html, flags=re.IGNORECASE
    )

    # Remove javascript: URLs
    html = re.sub(
        r'(href|src|action)\s*=\s*["\']?\s*javascript:[^"\'>\s]*["\']?',
        r'\1=""', html, flags=re.IGNORECASE
    )

    # Remove data: URLs (can embed scripts)
    html = re.sub(
        r'(href|src)\s*=\s*["\']?\s*data:[^"\'>\s]*["\']?',
        r'\1=""', html, flags=re.IGNORECASE
    )

    return html


# Custom WebEngine page to intercept link clicks
if HAS_WEBENGINE:
    from PySide6.QtWebEngineCore import QWebEnginePage

    class SafeLinkPage(QWebEnginePage):
        """WebEngine page that intercepts link clicks for safety checks."""

        def acceptNavigationRequest(self, url, nav_type, is_main_frame):
            # Only intercept link clicks, not initial page loads
            if nav_type == QWebEnginePage.NavigationType.NavigationTypeLinkClicked:
                url_str = url.toString()

                # Skip empty or javascript URLs
                if not url_str or url_str.startswith('javascript:'):
                    return False

                # Build warning message
                warnings = []

                # Check for HTTP (not HTTPS)
                if url_str.startswith('http://'):
                    warnings.append("• This link uses HTTP (not secure)")

                # Show confirmation dialog
                from PySide6.QtWidgets import QMessageBox
                msg = QMessageBox()
                msg.setWindowTitle("Open Link")
                msg.setIcon(QMessageBox.Icon.Question)

                safety_tip = "Before opening, verify this goes where you expect. Attackers often disguise malicious links."

                if warnings:
                    msg.setText("This link has security concerns:")
                    msg.setInformativeText("\n".join(warnings) + f"\n\n{safety_tip}\n\nURL: {url_str}")
                else:
                    msg.setText("Open this link in your browser?")
                    msg.setInformativeText(f"{safety_tip}\n\nURL: {url_str}")

                msg.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                msg.setDefaultButton(QMessageBox.StandardButton.No if warnings else QMessageBox.StandardButton.Yes)

                if msg.exec() == QMessageBox.StandardButton.Yes:
                    # Open in external browser
                    from PySide6.QtGui import QDesktopServices
                    QDesktopServices.openUrl(url)

                return False  # Don't navigate in the email viewer

            # Allow other navigation (initial load, etc.)
            return super().acceptNavigationRequest(url, nav_type, is_main_frame)


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

    def flags(self, index: QModelIndex) -> Qt.ItemFlag:
        """Return item flags - enable dragging."""
        default_flags = super().flags(index)
        if index.isValid():
            return default_flags | Qt.ItemFlag.ItemIsDragEnabled
        return default_flags

    def mimeTypes(self) -> list[str]:
        """Return supported MIME types for drag."""
        return ["application/x-mailbench-message"]

    def mimeData(self, indexes: list[QModelIndex]):
        """Create MIME data for drag operation."""
        from PySide6.QtCore import QMimeData
        mime_data = QMimeData()
        item_ids = []
        for index in indexes:
            if index.isValid():
                msg = self._filtered_messages[index.row()]
                item_ids.append(msg.item_id)
        if item_ids:
            mime_data.setData("application/x-mailbench-message", ",".join(item_ids).encode())
        return mime_data

    def supportedDragActions(self) -> Qt.DropAction:
        """Return supported drag actions."""
        return Qt.DropAction.MoveAction

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

    def update_messages(self, messages_data: list[dict]):
        """Update multiple messages (for auto-refresh without disruption)."""
        changed = False
        for msg_data in messages_data:
            item_id = msg_data.get('item_id')
            idx = self._message_map.get(item_id)
            if idx is not None and idx < len(self._messages):
                old_msg = self._messages[idx]
                new_is_read = msg_data.get('is_read', old_msg.is_read)
                new_is_flagged = msg_data.get('is_flagged', old_msg.is_flagged)
                if old_msg.is_read != new_is_read or old_msg.is_flagged != new_is_flagged:
                    old_msg.is_read = new_is_read
                    old_msg.is_flagged = new_is_flagged
                    changed = True
        if changed:
            # Emit data changed for entire list
            if self._filtered_messages:
                self.dataChanged.emit(
                    self.index(0),
                    self.index(len(self._filtered_messages) - 1)
                )

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
            self.secondary = QColor("#000000")  # Black text
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

    # Signal emitted when messages are dropped on a folder (item_ids, account_id, folder_id)
    messagesDropped = Signal(list, int, str)

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

    def flags(self, index: QModelIndex) -> Qt.ItemFlag:
        """Return item flags - enable dropping on folder items."""
        default_flags = super().flags(index)
        if index.isValid():
            item = self.itemFromIndex(index)
            if item:
                data = item.data(Qt.ItemDataRole.UserRole)
                # Enable drop only on folder items
                if data and data[0] == "folder":
                    return default_flags | Qt.ItemFlag.ItemIsDropEnabled
        return default_flags

    def mimeTypes(self) -> list[str]:
        """Return supported MIME types for drop."""
        return ["application/x-mailbench-message"]

    def supportedDropActions(self) -> Qt.DropAction:
        """Return supported drop actions."""
        return Qt.DropAction.MoveAction

    def canDropMimeData(self, data, action, row, column, parent) -> bool:
        """Check if drop is allowed."""
        if not data.hasFormat("application/x-mailbench-message"):
            return False
        if not parent.isValid():
            return False
        item = self.itemFromIndex(parent)
        if item:
            item_data = item.data(Qt.ItemDataRole.UserRole)
            # Only allow drop on folder items
            return item_data and item_data[0] == "folder"
        return False

    def dropMimeData(self, data, action, row, column, parent) -> bool:
        """Handle the drop - move messages to folder."""
        if not data.hasFormat("application/x-mailbench-message"):
            return False
        if not parent.isValid():
            return False

        item = self.itemFromIndex(parent)
        if not item:
            return False

        item_data = item.data(Qt.ItemDataRole.UserRole)
        if not item_data or item_data[0] != "folder":
            return False

        account_id = item_data[1]
        folder_id = item_data[2]

        # Parse the dropped message IDs
        raw_data = data.data("application/x-mailbench-message").data().decode()
        item_ids = [id.strip() for id in raw_data.split(",") if id.strip()]

        if item_ids:
            # Emit signal with the message IDs and target folder
            self.messagesDropped.emit(item_ids, account_id, folder_id)

        return True


class MessageWindow(QMainWindow):
    """Separate window for viewing a message."""

    def __init__(self, parent, msg_data: dict, full_data: dict = None, dark_mode: bool = False, account_id: int = None):
        super().__init__(parent)
        self.msg_data = msg_data
        self.full_data = full_data or {}
        self.account_id = account_id
        self._dark_mode = dark_mode
        self._compose_widget = None

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
        # Toolbar matching main window style
        toolbar = QToolBar("Message Toolbar")
        toolbar.setMovable(False)
        toolbar.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        toolbar.setIconSize(QSize(20, 20))
        self.addToolBar(toolbar)

        reply_action = QAction(QIcon.fromTheme("mail-reply-sender", QIcon.fromTheme("go-previous")), "Reply", self)
        reply_action.triggered.connect(self._reply)
        toolbar.addAction(reply_action)

        reply_all_action = QAction(QIcon.fromTheme("mail-reply-all"), "Reply All", self)
        reply_all_action.triggered.connect(self._reply_all)
        toolbar.addAction(reply_all_action)

        forward_action = QAction(QIcon.fromTheme("mail-forward", QIcon.fromTheme("go-next")), "Forward", self)
        forward_action.triggered.connect(self._forward)
        toolbar.addAction(forward_action)

        toolbar.addSeparator()

        delete_action = QAction(QIcon.fromTheme("edit-delete", QIcon.fromTheme("user-trash")), "Delete", self)
        delete_action.triggered.connect(self._delete)
        toolbar.addAction(delete_action)

        # Use stacked widget to switch between message view and compose
        self._stack = QStackedWidget()
        self.setCentralWidget(self._stack)

        # Message view widget (index 0)
        self._message_widget = QWidget()
        layout = QVBoxLayout(self._message_widget)
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
            # Use custom page for link safety
            self._body_page = SafeLinkPage(self.body_text)
            self.body_text.setPage(self._body_page)
            # Disable JavaScript for security
            self.body_text.settings().setAttribute(QWebEngineSettings.WebAttribute.JavascriptEnabled, False)
            if body_type == 'html':
                # Sanitize HTML to prevent XSS attacks
                body = sanitize_html(body)
                # Block remote images to prevent tracking
                body = block_remote_images(body)
                # Inject sans-serif font style
                font_style = "<style>body { font-family: sans-serif; } p { margin: 0.3em 0; } div { margin: 0; }</style>"
                if body.lower().strip().startswith('<!doctype') or body.lower().strip().startswith('<html'):
                    for tag in ['<head>', '<HEAD>', '<html>', '<HTML>']:
                        if tag in body:
                            body = body.replace(tag, tag + font_style, 1)
                            break
                else:
                    body = f"<html><head>{font_style}</head><body>{body}</body></html>"
                self.body_text.setHtml(body)
            else:
                escaped_body = body.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('\n', '<br>')
                self.body_text.setHtml(f"<html><body><pre style='white-space: pre-wrap; font-family: sans-serif;'>{escaped_body}</pre></body></html>")
        else:
            self.body_text = QTextEdit()
            self.body_text.setReadOnly(True)
            self.body_text.setFont(QFont("Sans Serif", 12))
            if body_type == 'html':
                # Sanitize HTML to prevent XSS attacks
                body = sanitize_html(body)
                # Block remote images to prevent tracking
                body = block_remote_images(body)
                self.body_text.setHtml(body)
            else:
                self.body_text.setPlainText(body)
        layout.addWidget(self.body_text, 1)  # stretch factor = 1

        # Add message widget to stack
        self._stack.addWidget(self._message_widget)

    def _show_compose(self, reply_to=None, forward=None):
        """Show compose widget in this window."""
        from mailbench.views.compose import ComposeWidget

        # Remove old compose widget if exists
        if self._compose_widget:
            self._stack.removeWidget(self._compose_widget)
            self._compose_widget.deleteLater()
            self._compose_widget = None

        parent = self.parent()
        signature = parent._signatures.get(self.account_id, "") if self.account_id else ""

        self._compose_widget = ComposeWidget(
            self, parent.db, parent.sync_manager, self.account_id,
            reply_to=reply_to, forward=forward, signature=signature,
            font_size=parent.font_size, zoom=parent._preview_zoom
        )
        self._compose_widget.message_sent.connect(self._on_compose_done)
        self._compose_widget.compose_cancelled.connect(self._on_compose_cancelled)

        if parent._address_book:
            self._compose_widget.set_address_book(parent._address_book)

        self._stack.addWidget(self._compose_widget)
        self._stack.setCurrentWidget(self._compose_widget)

        # Update window title
        if reply_to:
            self.setWindowTitle("Re: " + reply_to.get('subject', ''))
        elif forward:
            self.setWindowTitle("Fwd: " + forward.get('subject', ''))

        QTimer.singleShot(0, self._compose_widget.focus_to_field)

    def _on_compose_done(self):
        """Message sent - close the window."""
        self.close()

    def _on_compose_cancelled(self):
        """Compose cancelled - go back to message view."""
        self._stack.setCurrentWidget(self._message_widget)
        self.setWindowTitle(self.msg_data.get('subject', '(No Subject)'))

        if self._compose_widget:
            self._stack.removeWidget(self._compose_widget)
            self._compose_widget.deleteLater()
            self._compose_widget = None

    def _reply(self):
        """Reply to the message."""
        reply_data = {
            'from_name': self.msg_data.get('sender_name', ''),
            'from_email': self.msg_data.get('sender_email', ''),
            'subject': self.msg_data.get('subject', ''),
            'date': self.msg_data.get('date_received', ''),
            'to': self.full_data.get('to', ''),
            'body': self.full_data.get('body', ''),
            'reply_all': False
        }
        self._show_compose(reply_to=reply_data)

    def _reply_all(self):
        """Reply all to the message."""
        reply_data = {
            'from_name': self.msg_data.get('sender_name', ''),
            'from_email': self.msg_data.get('sender_email', ''),
            'subject': self.msg_data.get('subject', ''),
            'date': self.msg_data.get('date_received', ''),
            'to': self.full_data.get('to', ''),
            'cc': self.full_data.get('cc', ''),
            'body': self.full_data.get('body', ''),
            'reply_all': True
        }
        self._show_compose(reply_to=reply_data)

    def _forward(self):
        """Forward the message."""
        forward_data = {
            'from_name': self.msg_data.get('sender_name', ''),
            'from_email': self.msg_data.get('sender_email', ''),
            'subject': self.msg_data.get('subject', ''),
            'date': self.msg_data.get('date_received', ''),
            'to': self.full_data.get('to', ''),
            'body': self.full_data.get('body', '')
        }
        self._show_compose(forward=forward_data)

    def _delete(self):
        """Delete the message."""
        item_id = self.msg_data.get('item_id')
        if not item_id or not self.account_id:
            return

        parent = self.parent()
        parent.sync_manager.delete_message(
            self.account_id,
            item_id,
            callback=lambda s, e: parent._on_message_deleted(s, e, item_id, -1)
        )
        self.close()

    def _apply_theme(self):
        """Use system defaults."""
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
        # Migrate any plaintext passwords to secure keyring storage
        migrated = self.db.migrate_passwords_to_keyring()
        if migrated > 0:
            print(f"Migrated {migrated} password(s) to secure keyring storage")
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

        # Theme setting - always light mode for now
        self._theme_setting = "Light"
        self._dark_mode = False

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

        # Check for updates after startup
        QTimer.singleShot(500, self._check_for_updates)

        # Auto-refresh timer (check for new mail every 10 seconds)
        self._refresh_timer = QTimer(self)
        self._refresh_timer.timeout.connect(self._auto_check_mail)
        self._refresh_timer.start(10000)  # 10 seconds

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

        self._forward_action = QAction(QIcon.fromTheme("mail-forward", QIcon.fromTheme("go-next")), "Forward", self)
        self._forward_action.triggered.connect(self._forward)
        self._forward_action.setToolTip("Forward message (Ctrl+Shift+F)")
        toolbar.addAction(self._forward_action)

        toolbar.addSeparator()

        self._delete_action = QAction(QIcon.fromTheme("edit-delete", QIcon.fromTheme("user-trash")), "Delete", self)
        self._delete_action.triggered.connect(self._delete_messages)
        self._delete_action.setToolTip("Delete selected messages (Delete)")
        toolbar.addAction(self._delete_action)

        refresh_action = QAction(QIcon.fromTheme("view-refresh", QIcon.fromTheme("sync")), "Refresh", self)
        refresh_action.triggered.connect(self._check_mail)
        refresh_action.setToolTip("Check for new mail (F5)")
        toolbar.addAction(refresh_action)

        # Initially disable message actions (no message selected)
        self._update_message_actions(False)

    def _update_message_actions(self, enabled: bool):
        """Enable or disable message-related toolbar actions."""
        self._reply_action.setEnabled(enabled)
        self._reply_all_action.setEnabled(enabled)
        self._forward_action.setEnabled(enabled)
        self._delete_action.setEnabled(enabled)

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
        self._folder_model.messagesDropped.connect(self._on_messages_dropped)
        self._folder_tree = QTreeView()
        self._folder_tree.setModel(self._folder_model)
        self._folder_tree.setHeaderHidden(True)
        self._folder_tree.setAnimated(True)
        self._folder_tree.setIndentation(12)  # Reduce indent depth
        self._folder_tree.clicked.connect(self._on_folder_clicked)
        self._folder_tree.doubleClicked.connect(self._on_folder_double_clicked)
        self._folder_tree.expanded.connect(self._on_folder_expanded)
        self._folder_tree.collapsed.connect(self._on_folder_collapsed)
        # Enable drop
        self._folder_tree.setAcceptDrops(True)
        self._folder_tree.setDropIndicatorShown(True)
        self._folder_tree.setDragDropMode(QAbstractItemView.DragDropMode.DropOnly)
        # Context menu for folder operations
        self._folder_tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self._folder_tree.customContextMenuRequested.connect(self._show_folder_context_menu)
        folder_layout.addWidget(self._folder_tree)

        folder_widget.setMinimumWidth(100)
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
        self._message_delegate = MessageDelegate(dark_mode=False, font_size=self.font_size)
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
        # Enable drag
        self._message_list.setDragEnabled(True)
        self._message_list.setDragDropMode(QAbstractItemView.DragDropMode.DragOnly)
        # Handle keyboard navigation
        self._message_list.selectionModel().currentChanged.connect(self._on_message_selection_changed)
        message_layout.addWidget(self._message_list)

        message_widget.setMinimumWidth(150)
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

        # From line with inline verification indicator
        from_layout = QHBoxLayout()
        from_layout.setContentsMargins(0, 0, 0, 0)
        from_layout.setSpacing(6)

        self._preview_from = QLabel("From:")
        self._preview_from.setStyleSheet("font-weight: bold; color: #000000;")
        from_layout.addWidget(self._preview_from)

        # Sender verification indicator (inline, hidden by default)
        self._sender_verification = QLabel("")
        self._sender_verification.setMinimumWidth(24)
        self._sender_verification.hide()
        from_layout.addWidget(self._sender_verification)

        from_layout.addStretch()
        header_layout.addLayout(from_layout)

        self._preview_to = QLabel("To:")
        self._preview_to.setStyleSheet("color: #000000;")
        header_layout.addWidget(self._preview_to)

        self._preview_subject = QLabel("")
        self._preview_subject.setStyleSheet("font-weight: bold; font-size: 12pt; color: #000000;")
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

        # Images blocked banner (hidden by default)
        self._images_banner = QWidget()
        banner_layout = QHBoxLayout(self._images_banner)
        banner_layout.setContentsMargins(8, 6, 8, 6)
        banner_layout.setSpacing(6)

        banner_label = QLabel("Images blocked.")
        banner_label.setStyleSheet("color: #856404; font-weight: 500;")
        banner_layout.addWidget(banner_label)

        # Modern link-style buttons
        link_style = """
            QPushButton {
                background: none;
                border: none;
                color: #0066cc;
                text-decoration: underline;
                padding: 0 4px;
                font-weight: 500;
            }
            QPushButton:hover {
                color: #004499;
            }
        """

        self._load_images_btn = QPushButton("Load images")
        self._load_images_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._load_images_btn.setStyleSheet(link_style)
        self._load_images_btn.clicked.connect(self._load_images_once)
        banner_layout.addWidget(self._load_images_btn)

        self._trust_sender_btn = QPushButton("")
        self._trust_sender_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self._trust_sender_btn.setStyleSheet(link_style)
        self._trust_sender_btn.clicked.connect(self._trust_sender)
        banner_layout.addWidget(self._trust_sender_btn)

        banner_layout.addStretch()
        self._images_banner.setStyleSheet("background-color: #fff3cd; border: 1px solid #ffc107; border-radius: 4px;")
        self._images_banner.hide()
        preview_layout.addWidget(self._images_banner)

        # Track current message data for image loading
        self._current_sender_email: Optional[str] = None
        self._current_body_html: Optional[str] = None
        self._images_blocked: bool = False

        # Preview body - use WebEngine if available for proper HTML/image support
        if HAS_WEBENGINE:
            self._preview_body = QWebEngineView()
            # Use custom page for link safety
            self._preview_page = SafeLinkPage(self._preview_body)
            self._preview_body.setPage(self._preview_page)
            # Disable JavaScript for security
            self._preview_body.settings().setAttribute(QWebEngineSettings.WebAttribute.JavascriptEnabled, False)
            self._preview_body.setHtml("<html><body></body></html>")
            self._use_webengine = True
        else:
            self._preview_body = QTextEdit()
            self._preview_body.setReadOnly(True)
            self._preview_body.setFont(QFont("Sans Serif", 12))
            self._use_webengine = False

        # Body takes remaining space
        self._preview_body.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        preview_layout.addWidget(self._preview_body, 1)  # stretch factor = 1

        # Install event filter for Ctrl+scroll zoom
        self._preview_body.installEventFilter(self)

        self._preview_stack.addWidget(self._preview_widget)  # Index 0

        # Index 1: Compose widget (created on demand)
        self._compose_widget = None

        self._preview_stack.setMinimumWidth(200)
        self._right_splitter.addWidget(self._preview_stack)

        # Set splitter sizes
        self._splitter.setSizes([200, 1000])
        self._right_splitter.setSizes([400, 600])

        # Allow widgets to be resized freely
        self._splitter.setChildrenCollapsible(False)
        self._right_splitter.setChildrenCollapsible(False)

    def _create_statusbar(self):
        """Create the status bar."""
        self._statusbar = QStatusBar()
        self.setStatusBar(self._statusbar)
        self._statusbar.showMessage("Ready")

    def _apply_theme(self):
        """Apply theme - force light mode."""
        app = QApplication.instance()

        # Create explicit light palette
        palette = QPalette()
        palette.setColor(QPalette.ColorRole.Window, QColor("#f0f0f0"))
        palette.setColor(QPalette.ColorRole.WindowText, QColor("#000000"))
        palette.setColor(QPalette.ColorRole.Base, QColor("#ffffff"))
        palette.setColor(QPalette.ColorRole.AlternateBase, QColor("#f7f7f7"))
        palette.setColor(QPalette.ColorRole.Text, QColor("#000000"))
        palette.setColor(QPalette.ColorRole.Button, QColor("#f0f0f0"))
        palette.setColor(QPalette.ColorRole.ButtonText, QColor("#000000"))
        palette.setColor(QPalette.ColorRole.BrightText, QColor("#ffffff"))
        palette.setColor(QPalette.ColorRole.Highlight, QColor("#0078d4"))
        palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#ffffff"))
        palette.setColor(QPalette.ColorRole.Link, QColor("#0066cc"))
        palette.setColor(QPalette.ColorRole.PlaceholderText, QColor("#808080"))
        palette.setColor(QPalette.ColorRole.ToolTipBase, QColor("#ffffdc"))
        palette.setColor(QPalette.ColorRole.ToolTipText, QColor("#000000"))
        # Disabled colors
        palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.WindowText, QColor("#808080"))
        palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.Text, QColor("#808080"))
        palette.setColor(QPalette.ColorGroup.Disabled, QPalette.ColorRole.ButtonText, QColor("#808080"))
        app.setPalette(palette)

        self.setStyleSheet("")

        # Delegate always uses light mode
        self._message_delegate.set_dark_mode(False)
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
            # For QTextEdit, scale font (preserve sans-serif)
            font = QFont("Sans Serif", int(12 * zoom_factor))
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
        """Apply a theme setting - light mode only for now."""
        self._theme_setting = "Light"
        self._dark_mode = False
        self._dark_mode_action.setChecked(False)
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
            # Start listening for mailbox changes (push notifications)
            self.sync_manager.start_change_listener(
                account_id,
                callback=lambda aid, changes: self._on_mailbox_changes(aid, changes)
            )
        else:
            self._statusbar.showMessage(f"Connection failed: {error}")
            QMessageBox.warning(self, "Connection Failed", str(error))

    def _on_mailbox_changes(self, account_id: int, changes: list):
        """Handle mailbox changes from server push."""
        if not changes:
            return

        # Check if any changes are for the current folder
        needs_refresh = False
        for change in changes:
            change_type = change.get("type", "")
            folder_id = change.get("folderId", "")

            # Refresh if changes affect current folder or if it's a general mail change
            if change_type in ("mtMail", "mtFolder"):
                if folder_id == self._current_folder_id or not folder_id:
                    needs_refresh = True
                    break

        if needs_refresh and self._current_account_id == account_id:
            # Refresh the current folder
            self._statusbar.showMessage("New mail received...")
            self._load_messages(account_id, self._current_folder_id)

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

        # Clear preview pane and current message state
        self._preview_from.setText("")
        self._preview_to.setText("")
        self._preview_subject.setText("")
        self._preview_date.setText("")
        self._clear_attachments_display()
        if self._use_webengine:
            self._preview_body.setHtml("<html><body></body></html>")
        else:
            self._preview_body.clear()
        self._current_message_id = None
        self._current_message_full_data = None
        self._update_message_actions(False)

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
        self._update_message_actions(True)
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

        win = MessageWindow(self, msg_data, full_data, False, self._current_account_id)
        win.show()
        self._message_windows.append(win)

    def _update_sender_verification(self, sender_email: str, sender_name: str):
        """Update sender verification indicator based on sender analysis.

        Only shows warnings - known contacts don't need an indicator.
        """
        if not sender_email:
            self._sender_verification.hide()
            return

        sender_email_lower = sender_email.lower()
        is_known = False
        is_external = False
        is_suspicious = False
        suspicious_reason = ""

        # Check if sender is in address book
        if self._address_book:
            for contact in self._address_book:
                if contact.get('email', '').lower() == sender_email_lower:
                    is_known = True
                    break

        # Check if sender domain is external (different from user's domain)
        if self._current_account_id:
            for acc in self.db.get_accounts():
                if acc['id'] == self._current_account_id:
                    user_email = acc.get('email', '')
                    if '@' in user_email and '@' in sender_email:
                        user_domain = user_email.split('@')[1].lower()
                        sender_domain = sender_email.split('@')[1].lower()
                        if user_domain != sender_domain:
                            is_external = True
                    break

        # Check for suspicious display name (name looks like email but doesn't match)
        if sender_name and '@' in sender_name:
            name_email = sender_name.lower()
            if sender_email_lower not in name_email:
                is_suspicious = True
                suspicious_reason = "Display name contains a different email address - possible spoofing attempt"

        # Build indicator - only show warnings, not positive indicators
        if is_suspicious:
            # Red yield sign for suspicious sender
            self._sender_verification.setText("⚠️")
            self._sender_verification.setToolTip(suspicious_reason)
            font = self._sender_verification.font()
            font.setPointSize(16)
            self._sender_verification.setFont(font)
            self._sender_verification.setStyleSheet("color: #ff0000;")
            self._sender_verification.setCursor(Qt.CursorShape.WhatsThisCursor)
            self._sender_verification.show()
        elif is_external and not is_known:
            # Yellow/orange yield sign for unknown external sender
            self._sender_verification.setText("⚠️")
            self._sender_verification.setToolTip("External sender - not from your organization")
            font = self._sender_verification.font()
            font.setPointSize(16)
            self._sender_verification.setFont(font)
            self._sender_verification.setStyleSheet("color: #ff8c00;")
            self._sender_verification.setCursor(Qt.CursorShape.WhatsThisCursor)
            self._sender_verification.show()
        else:
            # Known contact or internal sender - no indicator needed
            self._sender_verification.hide()
            self._sender_verification.setToolTip("")

    def _show_message_preview(self, msg: MessageData):
        """Show message in preview pane."""
        sender = msg.sender_name
        if msg.sender_email:
            sender += f" <{msg.sender_email}>" if sender else msg.sender_email
        self._preview_from.setText(f"From: {sender}")
        self._preview_to.setText("To: Loading...")
        self._preview_subject.setText(msg.subject)
        self._preview_date.setText(msg.date_display)

        # Update sender verification indicator
        self._update_sender_verification(msg.sender_email, msg.sender_name)

        # Store sender email for image loading banner
        self._current_sender_email = msg.sender_email

        # Hide images banner while loading
        self._images_banner.hide()

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

        # Track body for image loading (sender_email set in _show_message_preview)
        self._current_body_html = body if body_type == 'html' else None
        self._images_blocked = False

        # Check if sender is trusted for remote images
        sender_trusted = self.db.is_trusted_sender(self._current_sender_email) if self._current_sender_email else False

        if self._use_webengine:
            # WebEngineView handles HTML natively
            if body_type == 'html':
                # Sanitize HTML to prevent XSS attacks
                body = sanitize_html(body)
                # Block remote images unless sender is trusted
                if not sender_trusted:
                    original_body = body
                    body = block_remote_images(body)
                    # Check if any images were actually blocked
                    self._images_blocked = (body != original_body)
                # Inject sans-serif font style
                font_style = "<style>body { font-family: sans-serif; } p { margin: 0.3em 0; } div { margin: 0; }</style>"
                if body.lower().strip().startswith('<!doctype') or body.lower().strip().startswith('<html'):
                    # Insert style after <head> or <html> tag
                    for tag in ['<head>', '<HEAD>', '<html>', '<HTML>']:
                        if tag in body:
                            body = body.replace(tag, tag + font_style, 1)
                            break
                else:
                    body = f"<html><head>{font_style}</head><body>{body}</body></html>"
                self._preview_body.setHtml(body)
            else:
                # Convert plain text to HTML
                escaped_body = body.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('\n', '<br>')
                self._preview_body.setHtml(f"<html><body><pre style='white-space: pre-wrap; font-family: sans-serif;'>{escaped_body}</pre></body></html>")
        else:
            # QTextEdit fallback
            if body_type == 'html':
                # Sanitize HTML to prevent XSS attacks
                body = sanitize_html(body)
                # Block remote images unless sender is trusted
                if not sender_trusted:
                    original_body = body
                    body = block_remote_images(body)
                    self._images_blocked = (body != original_body)
                self._preview_body.setHtml(body)
            else:
                self._preview_body.setPlainText(body)

        # Show/hide images blocked banner
        if self._images_blocked:
            self._trust_sender_btn.setText(f"Always Load from {self._current_sender_email}")
            self._images_banner.show()
        else:
            self._images_banner.hide()

        # Mark as read (locally and on server)
        item_id = data.get('item_id')
        if item_id and self._current_account_id:
            self._message_model.update_message(item_id, is_read=True)
            # Sync to server
            self.sync_manager.mark_as_read(
                self._current_account_id, item_id, True,
                callback=lambda s, e: None  # Silent callback
            )
            # Update in local cache
            if item_id in self._messages_by_id:
                self._messages_by_id[item_id]['is_read'] = True

    def _load_images_once(self):
        """Load remote images for current message only."""
        if not self._current_body_html:
            return

        # Re-render with images (sanitize but don't block)
        body = sanitize_html(self._current_body_html)

        if self._use_webengine:
            font_style = "<style>body { font-family: sans-serif; } p { margin: 0.3em 0; } div { margin: 0; }</style>"
            if body.lower().strip().startswith('<!doctype') or body.lower().strip().startswith('<html'):
                for tag in ['<head>', '<HEAD>', '<html>', '<HTML>']:
                    if tag in body:
                        body = body.replace(tag, tag + font_style, 1)
                        break
            else:
                body = f"<html><head>{font_style}</head><body>{body}</body></html>"
            self._preview_body.setHtml(body)
        else:
            self._preview_body.setHtml(body)

        self._images_banner.hide()
        self._images_blocked = False

    def _trust_sender(self):
        """Add current sender to trusted list and load images."""
        if not self._current_sender_email:
            return

        self.db.add_trusted_sender(self._current_sender_email)
        self._statusbar.showMessage(f"Added {self._current_sender_email} to trusted senders")
        self._load_images_once()

    def _clear_attachments_display(self):
        """Clear attachments display."""
        while self._attachments_layout.count():
            child = self._attachments_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

    def _get_file_icon(self, filename: str) -> str:
        """Get an icon character based on file extension."""
        ext = os.path.splitext(filename.lower())[1]
        if ext in {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.webp', '.svg'}:
            return "🖼"
        elif ext in {'.pdf'}:
            return "📄"
        elif ext in {'.doc', '.docx', '.odt', '.rtf'}:
            return "📝"
        elif ext in {'.xls', '.xlsx', '.ods', '.csv'}:
            return "📊"
        elif ext in {'.ppt', '.pptx', '.odp'}:
            return "📽"
        elif ext in {'.zip', '.rar', '.7z', '.tar', '.gz'}:
            return "📦"
        elif ext in {'.mp3', '.wav', '.ogg', '.flac', '.m4a'}:
            return "🎵"
        elif ext in {'.mp4', '.avi', '.mkv', '.mov', '.webm'}:
            return "🎬"
        elif ext in {'.txt', '.log'}:
            return "📃"
        elif ext in {'.exe', '.msi', '.bat', '.cmd'}:
            return "⚙"
        else:
            return "📎"

    def _format_file_size(self, size_bytes: int) -> str:
        """Format file size in human-readable form."""
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.1f} KB"
        else:
            return f"{size_bytes / (1024 * 1024):.1f} MB"

    def _show_attachments(self, attachments: list):
        """Show attachments in preview header with card-style display."""
        self._clear_attachments_display()

        if not attachments:
            return

        is_dark = self.db.get_setting("dark_mode", "0") == "1"

        # Header
        count = len(attachments)
        header = QLabel(f"{count} Attachment{'s' if count > 1 else ''}")
        header_color = "#999" if is_dark else "#666"
        header.setStyleSheet(f"font-weight: bold; color: {header_color};")
        self._attachments_layout.addWidget(header)

        # Cards container
        cards_widget = QWidget()
        cards_layout = QHBoxLayout(cards_widget)
        cards_layout.setContentsMargins(0, 0, 0, 0)
        cards_layout.setSpacing(10)

        for att in attachments:
            card = self._create_preview_attachment_card(att, is_dark)
            cards_layout.addWidget(card)

        cards_layout.addStretch()
        self._attachments_layout.addWidget(cards_widget)

    def _create_preview_attachment_card(self, attachment: dict, is_dark: bool) -> QWidget:
        """Create a card widget for an attachment in preview pane."""
        card = QFrame()
        card.setFrameShape(QFrame.Shape.NoFrame)
        card.setObjectName("attachCard")

        if is_dark:
            card.setStyleSheet("""
                QFrame#attachCard {
                    background-color: #3c3f41;
                    border: 1px solid #555;
                    border-radius: 4px;
                }
                QFrame#attachCard:hover {
                    background-color: #4c5052;
                    border: 1px solid #666;
                }
            """)
        else:
            card.setStyleSheet("""
                QFrame#attachCard {
                    background-color: #f5f5f5;
                    border: 1px solid #ddd;
                    border-radius: 4px;
                }
                QFrame#attachCard:hover {
                    background-color: #e8e8e8;
                    border: 1px solid #ccc;
                }
            """)

        layout = QHBoxLayout(card)
        layout.setContentsMargins(8, 5, 8, 5)
        layout.setSpacing(8)

        # Icon
        name = attachment.get('name', 'attachment')
        icon_label = QLabel(self._get_file_icon(name))
        icon_label.setStyleSheet("font-size: 24px; background: transparent; border: none;")
        layout.addWidget(icon_label)

        # Name and size
        info_widget = QWidget()
        info_widget.setStyleSheet("background: transparent; border: none;")
        info_layout = QVBoxLayout(info_widget)
        info_layout.setContentsMargins(0, 0, 0, 0)
        info_layout.setSpacing(0)

        # Truncate long names
        display_name = name
        if len(display_name) > 25:
            display_name = display_name[:22] + "..."

        name_label = QLabel(display_name)
        name_label.setStyleSheet("font-weight: 500; background: transparent; border: none;")
        name_label.setToolTip(name)
        info_layout.addWidget(name_label)

        # Size
        size = attachment.get('size', 0)
        size_label = QLabel(self._format_file_size(size))
        size_color = "#777" if is_dark else "#888"
        size_label.setStyleSheet(f"color: {size_color}; font-size: 11px; background: transparent; border: none;")
        info_layout.addWidget(size_label)

        layout.addWidget(info_widget)

        # Menu button with dropdown
        menu_btn = QToolButton()
        menu_btn.setText("▼")
        menu_btn.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        menu_btn.setFixedSize(24, 24)

        btn_style = f"""
            QToolButton {{
                background: transparent;
                border: none;
                font-size: 10px;
                color: {'#888' if is_dark else '#666'};
            }}
            QToolButton:hover {{
                background-color: {'#4c5052' if is_dark else '#ddd'};
                border-radius: 4px;
            }}
            QToolButton::menu-indicator {{
                image: none;
            }}
        """
        menu_btn.setStyleSheet(btn_style)

        menu = QMenu(menu_btn)

        # For preview, offer Open and Save
        open_action = menu.addAction("📂 Open")
        open_action.triggered.connect(lambda checked, a=attachment: self._open_attachment_direct(a))

        save_action = menu.addAction("💾 Save As...")
        save_action.triggered.connect(lambda checked, a=attachment: self._save_attachment_direct(a))

        menu_btn.setMenu(menu)
        layout.addWidget(menu_btn)

        return card

    def _open_attachment_direct(self, attachment: dict):
        """Open attachment directly without prompting."""
        name = attachment.get('name', 'attachment')
        url = attachment.get('url', '')

        if not url or not self._current_account_id:
            return

        session = self.kerio_pool.get_session(self._current_account_id)
        if not session:
            return

        # Check for dangerous file types
        is_dangerous, warning = self._get_attachment_warning(name)
        if is_dangerous:
            msg = QMessageBox(self)
            msg.setWindowTitle("Security Warning")
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setText(f"Warning: {name}")
            msg.setInformativeText(f"{warning}\n\nDo you want to continue?")
            msg.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            msg.setDefaultButton(QMessageBox.StandardButton.No)
            if msg.exec() != QMessageBox.StandardButton.Yes:
                return

        full_url = f"https://{session.config.server}{url}"
        self._open_attachment(full_url, name, session)

    def _get_default_downloads_dir(self) -> str:
        """Get the default downloads directory from settings or OS default."""
        # Check saved setting first
        saved_dir = self.db.get_setting("default_save_directory", "")
        if saved_dir and os.path.isdir(saved_dir):
            return saved_dir

        # Fall back to OS-appropriate default
        system = platform.system()
        if system == "Darwin":
            return os.path.join(os.path.expanduser("~"), "Downloads")
        elif system == "Windows":
            return os.path.join(os.path.expanduser("~"), "Downloads")
        else:
            # Linux - try XDG
            try:
                result = subprocess.run(
                    ["xdg-user-dir", "DOWNLOAD"],
                    capture_output=True, text=True, timeout=2
                )
                if result.returncode == 0 and result.stdout.strip():
                    path = result.stdout.strip()
                    if os.path.isdir(path):
                        return path
            except Exception:
                pass
            return os.path.join(os.path.expanduser("~"), "Downloads")

    def _save_attachment_direct(self, attachment: dict):
        """Save attachment directly without prompting."""
        name = attachment.get('name', 'attachment')
        url = attachment.get('url', '')

        if not url or not self._current_account_id:
            return

        session = self.kerio_pool.get_session(self._current_account_id)
        if not session:
            return

        # Use default downloads directory
        default_dir = self._get_default_downloads_dir()
        default_path = os.path.join(default_dir, name)

        save_path, _ = QFileDialog.getSaveFileName(self, "Save Attachment", default_path)
        if save_path:
            full_url = f"https://{session.config.server}{url}"
            self._save_attachment(full_url, save_path, session)

    def _get_attachment_warning(self, filename: str) -> tuple[bool, str]:
        """Check if attachment is potentially dangerous and return warning message.

        Returns (is_dangerous, warning_message).
        """
        _, ext = os.path.splitext(filename.lower())

        # Executable extensions - very dangerous
        executables = {'.exe', '.msi', '.bat', '.cmd', '.com', '.scr', '.pif'}
        # Script extensions - dangerous
        scripts = {'.js', '.vbs', '.vbe', '.jse', '.wsf', '.wsh', '.ps1', '.psm1'}
        # Office macros - potentially dangerous
        office_macros = {'.docm', '.xlsm', '.pptm', '.dotm', '.xltm', '.potm'}
        # Archives that can hide malware
        archives = {'.zip', '.rar', '.7z', '.tar', '.gz', '.iso', '.img'}
        # Other potentially dangerous
        other_risky = {'.jar', '.hta', '.cpl', '.msc', '.lnk', '.reg', '.dll'}

        if ext in executables:
            return (True, f"This is an executable file ({ext}). Executable files can harm your computer if they contain malware. Only open if you trust the sender and were expecting this file.")
        elif ext in scripts:
            return (True, f"This is a script file ({ext}). Script files can run code on your computer. Only open if you trust the sender and were expecting this file.")
        elif ext in office_macros:
            return (True, f"This Office file ({ext}) may contain macros. Macros can run code on your computer. Only open if you trust the sender.")
        elif ext in archives:
            return (False, f"This is an archive file ({ext}). Archives can contain hidden executable files. Scan the contents before opening any files inside.")
        elif ext in other_risky:
            return (True, f"This file type ({ext}) can be used to install software or change settings. Only open if you trust the sender.")

        return (False, "")

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

        # Check for dangerous file types
        is_dangerous, warning = self._get_attachment_warning(name)

        if is_dangerous:
            # Show warning for dangerous files
            msg = QMessageBox(self)
            msg.setWindowTitle("Security Warning")
            msg.setIcon(QMessageBox.Icon.Warning)
            msg.setText(f"Warning: {name}")
            msg.setInformativeText(f"{warning}\n\nDo you want to continue?")
            msg.setStandardButtons(QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
            msg.setDefaultButton(QMessageBox.StandardButton.No)

            if msg.exec() != QMessageBox.StandardButton.Yes:
                return
        elif warning:
            # Show info for files that need caution (archives)
            msg = QMessageBox(self)
            msg.setWindowTitle("Caution")
            msg.setIcon(QMessageBox.Icon.Information)
            msg.setText(f"Note: {name}")
            msg.setInformativeText(warning)
            msg.setStandardButtons(QMessageBox.StandardButton.Ok)
            msg.exec()

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
            default_dir = self._get_default_downloads_dir()
            default_path = os.path.join(default_dir, name)
            save_path, _ = QFileDialog.getSaveFileName(self, "Save Attachment", default_path)
            if save_path:
                self._save_attachment(full_url, save_path, session)

    def _open_attachment(self, url: str, filename: str, session):
        """Download attachment to temp file and open."""
        def do_open():
            try:
                import requests
                cookies = {session.cookie_name: session.cookie_value} if hasattr(session, 'cookie_name') else {}
                headers = {"X-Token": session.token} if hasattr(session, 'token') else {}

                response = requests.get(url, cookies=cookies, headers=headers, verify=True, timeout=60)
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

                response = requests.get(url, cookies=cookies, headers=headers, verify=True, timeout=60)
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

    def _show_folder_context_menu(self, position):
        """Show context menu for folder tree."""
        index = self._folder_tree.indexAt(position)
        if not index.isValid():
            return

        item = self._folder_model.itemFromIndex(index)
        if not item:
            return

        data = item.data(Qt.ItemDataRole.UserRole)
        if not data or data[0] != "folder":
            return

        account_id = data[1]
        folder_id = data[2]
        folder_name = item.text().lower()

        # Check if this is a trash/deleted folder
        is_trash = 'trash' in folder_name or 'deleted' in folder_name

        if not is_trash:
            return  # No context menu for non-trash folders (for now)

        menu = QMenu(self)
        empty_action = menu.addAction("Empty Trash")
        empty_action.triggered.connect(lambda: self._empty_trash(account_id, folder_id))
        menu.exec_(self._folder_tree.mapToGlobal(position))

    def _empty_trash(self, account_id: int, folder_id: str):
        """Empty the trash folder - permanently delete all messages."""
        # Confirm with user
        result = QMessageBox.warning(
            self,
            "Empty Trash",
            "Permanently delete all messages in Trash?\n\nThis action cannot be undone.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No
        )

        if result != QMessageBox.StandardButton.Yes:
            return

        self._statusbar.showMessage("Emptying trash...")

        def on_empty_complete(success: bool, error: str, count: int):
            if success:
                self._statusbar.showMessage(f"Deleted {count} messages from trash")
                # Refresh the folder if we're viewing it
                if self._current_folder_id == folder_id:
                    self._message_model.clear()
                    self._current_message_id = None
                    self._clear_preview()
            else:
                self._statusbar.showMessage(f"Failed to empty trash: {error}")
                QMessageBox.warning(self, "Error", f"Failed to empty trash: {error}")

        self.sync_manager.empty_trash(account_id, folder_id, callback=on_empty_complete)

    def _clear_preview(self):
        """Clear the message preview pane."""
        self._preview_from.setText("From:")
        self._sender_verification.hide()
        self._preview_to.setText("To:")
        self._preview_subject.setText("")
        self._preview_date.setText("")
        if self._use_webengine:
            self._preview_body.setHtml("<html><body></body></html>")
        else:
            self._preview_body.setText("")

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

    def _on_messages_dropped(self, item_ids: list, account_id: int, folder_id: str):
        """Handle messages dropped on a folder (drag and drop)."""
        if not item_ids:
            return

        # Show status
        count = len(item_ids)
        self._statusbar.showMessage(f"Moving {count} message{'s' if count > 1 else ''}...")

        for item_id in item_ids:
            self.sync_manager.move_message(
                account_id, item_id, folder_id,
                callback=lambda s, e, iid=item_id: self._on_message_moved(s, e, iid)
            )

    # ==================== Compose (Inline) ====================

    def _show_compose(self, reply_to=None, forward=None):
        """Show the compose widget inline, replacing the preview pane."""
        from mailbench.views.compose import ComposeWidget

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
            # Get current selection before removing
            current_index = self._message_list.currentIndex()
            current_row = current_index.row() if current_index.isValid() else select_row

            self._message_model.remove_message(item_id)

            # Always select next available message
            row_count = self._message_model.rowCount()
            if row_count > 0:
                # Use the row we were at, clamped to valid range
                target_row = current_row if current_row >= 0 else select_row
                new_row = min(max(0, target_row), row_count - 1)
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
        """Refresh current folder (full refresh, clears preview)."""
        if self._current_account_id and self._current_folder_id:
            self._load_messages(self._current_account_id, self._current_folder_id)

    def _auto_check_mail(self):
        """Auto-refresh: update message list without disrupting current view."""
        if not self._current_account_id or not self._current_folder_id:
            return
        if self._current_account_id not in self.connected_accounts:
            return

        # Remember current selection
        current_selection = self._current_message_id

        def on_messages_synced(success, error=None, messages_data=None):
            if not success or not messages_data:
                return

            # Check if we have new messages or changes
            new_ids = {m.get('item_id') for m in messages_data}
            old_ids = set(self._messages_by_id.keys())

            if new_ids == old_ids:
                # No changes, just update read status etc.
                for msg in messages_data:
                    item_id = msg.get('item_id')
                    if item_id in self._messages_by_id:
                        self._messages_by_id[item_id] = msg
                # Update model for read status changes
                self._message_model.update_messages(messages_data)
                return

            # There are changes - update the list but preserve selection
            self._messages_by_id.clear()
            for msg in messages_data:
                item_id = msg.get('item_id')
                if item_id:
                    self._messages_by_id[item_id] = msg

            self._message_model.clear()
            self._message_model.add_messages(messages_data)

            # Restore selection if the message still exists
            if current_selection and current_selection in self._messages_by_id:
                for row in range(self._message_model.rowCount()):
                    idx = self._message_model.index(row, 0)
                    msg = self._message_model.data(idx, Qt.ItemDataRole.DisplayRole)
                    if msg and msg.item_id == current_selection:
                        self._message_list.setCurrentIndex(idx)
                        break

            # Update status
            new_count = len(new_ids - old_ids)
            if new_count > 0:
                self._statusbar.showMessage(f"{new_count} new message(s)")

        self.sync_manager.sync_messages(
            self._current_account_id,
            self._current_folder_id,
            callback=on_messages_synced
        )

    def _add_account(self):
        """Add new account."""
        from mailbench.dialogs.dialogs import AccountDialog
        dialog = AccountDialog(self, self.db, self.kerio_pool, app=self)
        dialog.exec()
        self._refresh_accounts()

    def _manage_accounts(self):
        """Manage accounts."""
        from mailbench.dialogs.dialogs import AccountDialog
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
        from mailbench.dialogs.dialogs import SettingsDialog
        dialog = SettingsDialog(self, self.db, self)
        dialog.exec()

    def _show_about(self):
        """Show about dialog."""
        QMessageBox.about(
            self, "About Mailbench",
            f"Mailbench v{__version__}\n\nA Python email client for Kerio Connect.\n\n"
            f"Built with PySide6 (Qt6)"
        )

    def _check_for_updates(self):
        """Check for updates in background."""
        from mailbench.version import get_pypi_version, is_newer_version

        def do_check():
            try:
                latest = get_pypi_version()
                if latest and is_newer_version(latest, __version__):
                    self._update_version = latest
            except Exception:
                pass

        def on_done():
            version = getattr(self, '_update_version', None)
            if version:
                self._show_update_dialog(version)

        thread = threading.Thread(target=do_check, daemon=True)
        thread.start()
        QTimer.singleShot(3000, on_done)

    def _show_update_dialog(self, latest_version: str):
        """Show update available dialog."""
        result = QMessageBox.question(
            self, "Update Available",
            f"A new version of Mailbench is available.\n\n"
            f"Installed: {__version__}\n"
            f"Latest: {latest_version}\n\n"
            f"Would you like to upgrade now?\n\n"
            f"(This will run: pipx upgrade mailbench)",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
        if result == QMessageBox.StandardButton.Yes:
            self._run_upgrade()

    def _run_upgrade(self):
        """Run pipx upgrade in background."""
        self._upgrade_result = None

        def do_upgrade():
            try:
                result = subprocess.run(
                    ["pipx", "upgrade", "mailbench", "--force"],
                    capture_output=True, text=True,
                    stdin=subprocess.DEVNULL, timeout=120)
                if result.returncode == 0:
                    # Give pipx time to finalize
                    import time
                    time.sleep(2)
                    self._upgrade_result = (True, "Mailbench has been upgraded.")
                else:
                    error = result.stderr or result.stdout or "Unknown error"
                    self._upgrade_result = (False, f"Failed to upgrade:\n{error}")
            except subprocess.TimeoutExpired:
                self._upgrade_result = (False, "Upgrade timed out. Please upgrade manually:\n\npipx upgrade mailbench")
            except FileNotFoundError:
                self._upgrade_result = (False, "pipx not found. Please upgrade manually:\n\npipx upgrade mailbench")
            except Exception as e:
                self._upgrade_result = (False, f"Failed to upgrade:\n{e}")

        def poll_result():
            result = self._upgrade_result
            if result is None:
                QTimer.singleShot(500, poll_result)
                return
            success, message = result
            self._statusbar.showMessage("Upgrade complete" if success else "Upgrade failed", 3000)
            if success:
                result = QMessageBox.question(
                    self, "Upgrade Complete",
                    "Mailbench has been upgraded.\n\nWould you like to restart now?",
                    QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)
                if result == QMessageBox.StandardButton.Yes:
                    self._restart_app()
            else:
                QMessageBox.warning(self, "Upgrade Failed", message)

        self._statusbar.showMessage("Upgrading Mailbench...")
        thread = threading.Thread(target=do_upgrade, daemon=True)
        thread.start()
        QTimer.singleShot(2000, poll_result)

    def _restart_app(self):
        """Restart the application."""
        self._save_geometry()

        # Small delay to ensure pipx has finished writing all files
        import time
        time.sleep(1)

        if sys.argv[0].endswith('mailbench') or 'mailbench' in sys.argv[0]:
            subprocess.Popen([sys.argv[0]])
        else:
            subprocess.Popen([sys.executable, '-m', 'mailbench'])

        self.close()
        os._exit(0)

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

        # Stop change listeners and signal threads to stop
        for account_id in self.connected_accounts:
            self.sync_manager.stop_change_listener(account_id)
        self.sync_manager._shutdown = True
        self.kerio_pool.close_all()

        # Accept and let Qt clean up properly (may help XWayland)
        event.accept()
        QApplication.instance().quit()


def main():
    """Main entry point."""
    app = QApplication(sys.argv)
    app.setApplicationName("Mailbench")
    app.setApplicationVersion(__version__)
    app.setOrganizationName("Mailbench")

    # Enable high DPI scaling
    # app.setStyle("Fusion")  # Disabled - was breaking menu hover on some systems

    window = MailbenchWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
