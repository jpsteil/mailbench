"""Contact list model and delegate for Qt views."""

import json
from typing import Optional, List

from PySide6.QtCore import Qt, QAbstractListModel, QModelIndex, QSize
from PySide6.QtWidgets import QStyledItemDelegate, QStyle
from PySide6.QtGui import QPainter, QColor, QFont, QPen, QFontMetrics


class ContactListModel(QAbstractListModel):
    """Model for displaying contacts in a list view."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._contacts: List[dict] = []
        self._filtered_contacts: List[dict] = []
        self._filter_text: str = ""

    def rowCount(self, parent=QModelIndex()) -> int:
        if parent.isValid():
            return 0
        return len(self._filtered_contacts)

    def data(self, index: QModelIndex, role: int = Qt.ItemDataRole.DisplayRole):
        if not index.isValid():
            return None

        contact = self._filtered_contacts[index.row()]

        if role == Qt.ItemDataRole.DisplayRole:
            return contact.get("common_name", "")
        elif role == Qt.ItemDataRole.UserRole:
            return contact
        elif role == Qt.ItemDataRole.UserRole + 1:
            # Primary email
            emails = contact.get("email_addresses", [])
            if isinstance(emails, str):
                try:
                    emails = json.loads(emails)
                except (json.JSONDecodeError, TypeError):
                    emails = []
            if emails:
                first = emails[0]
                if isinstance(first, dict):
                    return first.get("address", "")
                return str(first)
            return ""
        elif role == Qt.ItemDataRole.UserRole + 2:
            # Company
            return contact.get("company", "")

        return None

    def set_contacts(self, contacts: List[dict]):
        """Set the full contact list."""
        self.beginResetModel()
        # Sort contacts by name
        self._contacts = sorted(contacts, key=lambda c: (
            c.get("common_name") or
            f"{c.get('last_name', '')} {c.get('first_name', '')}".strip() or
            ""
        ).lower())
        self._apply_filter()
        self.endResetModel()

    def clear(self):
        """Clear all contacts."""
        self.beginResetModel()
        self._contacts = []
        self._filtered_contacts = []
        self.endResetModel()

    def set_filter(self, text: str):
        """Filter contacts by search text."""
        self.beginResetModel()
        self._filter_text = text.lower().strip()
        self._apply_filter()
        self.endResetModel()

    def _apply_filter(self):
        """Apply current filter to contacts."""
        if not self._filter_text:
            self._filtered_contacts = self._contacts[:]
            return

        self._filtered_contacts = []
        for contact in self._contacts:
            # Search in name
            name = contact.get("common_name", "") or ""
            first = contact.get("first_name", "") or ""
            last = contact.get("last_name", "") or ""
            company = contact.get("company", "") or ""

            # Search in emails
            emails_str = ""
            emails = contact.get("email_addresses", [])
            if isinstance(emails, str):
                emails_str = emails
            elif isinstance(emails, list):
                for e in emails:
                    if isinstance(e, dict):
                        emails_str += e.get("address", "") + " "
                    else:
                        emails_str += str(e) + " "

            search_text = f"{name} {first} {last} {company} {emails_str}".lower()
            if self._filter_text in search_text:
                self._filtered_contacts.append(contact)

    def get_contact(self, row: int) -> Optional[dict]:
        """Get contact data by row index."""
        if 0 <= row < len(self._filtered_contacts):
            return self._filtered_contacts[row]
        return None

    def find_index(self, item_id: str) -> QModelIndex:
        """Find model index for a contact by item_id."""
        for i, contact in enumerate(self._filtered_contacts):
            if contact.get("item_id") == item_id:
                return self.createIndex(i, 0)
        return QModelIndex()

    def refresh(self):
        """Refresh the view."""
        self.beginResetModel()
        self._apply_filter()
        self.endResetModel()


class ContactDelegate(QStyledItemDelegate):
    """Custom delegate for rendering contact list items."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._row_height = 56

    def sizeHint(self, option, index) -> QSize:
        return QSize(option.rect.width(), self._row_height)

    def paint(self, painter: QPainter, option, index: QModelIndex):
        painter.save()

        # Background
        if option.state & QStyle.StateFlag.State_Selected:
            painter.fillRect(option.rect, QColor("#d0e8ff"))
        elif option.state & QStyle.StateFlag.State_MouseOver:
            painter.fillRect(option.rect, QColor("#f0f0f0"))
        else:
            painter.fillRect(option.rect, Qt.GlobalColor.white)

        # Get contact data
        contact = index.data(Qt.ItemDataRole.UserRole)
        if not contact:
            painter.restore()
            return

        name = contact.get("common_name", "") or ""
        if not name:
            first = contact.get("first_name", "") or ""
            last = contact.get("last_name", "") or ""
            name = f"{first} {last}".strip()

        # Get primary email
        emails = contact.get("email_addresses", [])
        if isinstance(emails, str):
            try:
                emails = json.loads(emails)
            except (json.JSONDecodeError, TypeError):
                emails = []
        primary_email = ""
        if emails:
            first_email = emails[0]
            if isinstance(first_email, dict):
                primary_email = first_email.get("address", "")
            else:
                primary_email = str(first_email)

        # Get primary phone if no email
        phones = contact.get("phone_numbers", [])
        if isinstance(phones, str):
            try:
                phones = json.loads(phones)
            except (json.JSONDecodeError, TypeError):
                phones = []
        primary_phone = ""
        if phones:
            first_phone = phones[0]
            if isinstance(first_phone, dict):
                primary_phone = first_phone.get("number", "")
            else:
                primary_phone = str(first_phone)

        company = contact.get("company", "") or ""

        # Layout
        x = option.rect.x() + 12
        y = option.rect.y()
        w = option.rect.width() - 24
        h = option.rect.height()

        # Avatar placeholder (circle with initials)
        avatar_size = 36
        avatar_x = x
        avatar_y = y + (h - avatar_size) // 2
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QColor("#4a90d9"))
        painter.drawEllipse(avatar_x, avatar_y, avatar_size, avatar_size)

        # Initials
        initials = ""
        if name:
            parts = name.split()
            if len(parts) >= 2:
                initials = parts[0][0].upper() + parts[-1][0].upper()
            elif parts:
                initials = parts[0][0].upper()
        painter.setPen(Qt.GlobalColor.white)
        painter.setFont(QFont("sans-serif", 12, QFont.Weight.Bold))
        painter.drawText(avatar_x, avatar_y, avatar_size, avatar_size,
                        Qt.AlignmentFlag.AlignCenter, initials)

        # Text area
        text_x = avatar_x + avatar_size + 12
        text_w = w - avatar_size - 12

        # Name (bold)
        painter.setPen(Qt.GlobalColor.black)
        name_font = QFont("sans-serif", 11)
        name_font.setBold(True)
        painter.setFont(name_font)
        name_rect = painter.fontMetrics().boundingRect(name)
        painter.drawText(text_x, y + 14, text_w, name_rect.height(),
                        Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop,
                        name if name else "(No Name)")

        # Email, phone, or company (gray, below name)
        painter.setPen(QColor("#666666"))
        detail_font = QFont("sans-serif", 9)
        painter.setFont(detail_font)
        detail_text = primary_email or primary_phone or company
        if detail_text:
            painter.drawText(text_x, y + 32, text_w, 18,
                            Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop,
                            detail_text)

        # Bottom border
        painter.setPen(QPen(QColor("#e0e0e0")))
        painter.drawLine(option.rect.left(), option.rect.bottom(),
                        option.rect.right(), option.rect.bottom())

        painter.restore()
