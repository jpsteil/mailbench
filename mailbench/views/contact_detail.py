"""Contact detail view for viewing and editing contacts."""

import json
from typing import Optional, List

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit, QTextEdit,
    QPushButton, QScrollArea, QFrame, QFormLayout, QGridLayout,
    QSizePolicy, QMessageBox
)
from PySide6.QtGui import QFont, QPainter, QColor, QPen


class AvatarWidget(QWidget):
    """Circular avatar with initials."""

    def __init__(self, size=80, parent=None):
        super().__init__(parent)
        self._size = size
        self._initials = ""
        self._color = QColor("#4a90d9")
        self.setFixedSize(size, size)

    def set_initials(self, name: str):
        """Set initials from a name."""
        if not name:
            self._initials = ""
        else:
            parts = name.split()
            if len(parts) >= 2:
                self._initials = parts[0][0].upper() + parts[-1][0].upper()
            elif parts:
                self._initials = parts[0][0].upper()
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        # Draw circle
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(self._color)
        painter.drawEllipse(0, 0, self._size, self._size)

        # Draw initials
        painter.setPen(Qt.GlobalColor.white)
        font = QFont("sans-serif", self._size // 3, QFont.Weight.Bold)
        painter.setFont(font)
        painter.drawText(0, 0, self._size, self._size,
                        Qt.AlignmentFlag.AlignCenter, self._initials)


class MultiValueEditor(QWidget):
    """Editor for multiple values (emails, phones)."""

    valueChanged = Signal()

    def __init__(self, field_type: str = "email", parent=None):
        super().__init__(parent)
        self._field_type = field_type
        self._entries: List[dict] = []
        self._widgets: List[tuple] = []  # (line_edit, type_combo, remove_btn)
        self._edit_mode = False
        self._setup_ui()

    def _setup_ui(self):
        self._layout = QVBoxLayout(self)
        self._layout.setContentsMargins(0, 0, 0, 0)
        self._layout.setSpacing(4)

    def set_values(self, values: list):
        """Set the values to display."""
        # Clear existing
        for w in self._widgets:
            for widget in w:
                if widget:
                    widget.deleteLater()
        self._widgets.clear()

        self._entries = values if values else []

        for entry in self._entries:
            self._add_entry_widget(entry)

        # Add empty entry in edit mode
        if self._edit_mode:
            self._add_entry_widget({})

    def _add_entry_widget(self, entry: dict):
        """Add a widget row for one entry."""
        row = QHBoxLayout()
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(4)

        if self._field_type == "email":
            value = entry.get("address", "")
            placeholder = "Email address"
        else:
            value = entry.get("number", "")
            placeholder = "Phone number"

        line_edit = QLineEdit(value)
        line_edit.setPlaceholderText(placeholder)
        line_edit.setReadOnly(not self._edit_mode)
        line_edit.textChanged.connect(self.valueChanged.emit)
        row.addWidget(line_edit, 1)

        # Type label (work, home, etc.)
        type_label = QLabel(entry.get("type", "work").capitalize())
        type_label.setStyleSheet("color: gray; font-size: 10pt;")
        row.addWidget(type_label)

        # Remove button (only in edit mode)
        remove_btn = None
        if self._edit_mode:
            remove_btn = QPushButton("-")
            remove_btn.setFixedSize(24, 24)
            remove_btn.clicked.connect(lambda: self._remove_entry(line_edit))
            row.addWidget(remove_btn)

        self._widgets.append((line_edit, type_label, remove_btn))

        container = QWidget()
        container.setLayout(row)
        self._layout.addWidget(container)

    def _remove_entry(self, line_edit: QLineEdit):
        """Remove an entry."""
        for i, (le, _, _) in enumerate(self._widgets):
            if le == line_edit:
                # Remove from layout
                widget = self._layout.itemAt(i).widget()
                if widget:
                    widget.deleteLater()
                del self._widgets[i]
                self.valueChanged.emit()
                break

    def get_values(self) -> list:
        """Get current values."""
        values = []
        for line_edit, type_label, _ in self._widgets:
            text = line_edit.text().strip()
            if text:
                if self._field_type == "email":
                    values.append({"address": text, "type": "work"})
                else:
                    values.append({"number": text, "type": "work"})
        return values

    def set_edit_mode(self, enabled: bool):
        """Enable or disable edit mode."""
        self._edit_mode = enabled
        # Refresh the display
        self.set_values(self.get_values() if enabled else self._entries)

    def set_readonly_style(self, readonly: bool):
        """Style the line edits as read-only (no border)."""
        style = "QLineEdit { border: none; background: transparent; }" if readonly else ""
        for line_edit, _, _ in self._widgets:
            line_edit.setStyleSheet(style)


class ContactDetailView(QWidget):
    """Detail view for a contact with view/edit modes."""

    saveRequested = Signal(dict)  # contact_data
    deleteRequested = Signal(str)  # item_id
    cancelRequested = Signal()
    newMessageRequested = Signal(str, str)  # email, name

    def __init__(self, parent=None):
        super().__init__(parent)
        self._contact: Optional[dict] = None
        self._edit_mode = False
        self._setup_ui()

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 16, 16, 16)
        layout.setSpacing(16)

        # Header with avatar and name
        header = QHBoxLayout()
        header.setSpacing(16)

        self._avatar = AvatarWidget(80)
        header.addWidget(self._avatar)

        name_layout = QVBoxLayout()
        name_layout.setSpacing(4)

        self._name_label = QLabel()
        self._name_label.setFont(QFont("sans-serif", 18, QFont.Weight.Bold))
        name_layout.addWidget(self._name_label)

        self._name_edit = QLineEdit()
        self._name_edit.setPlaceholderText("Full Name")
        self._name_edit.setFont(QFont("sans-serif", 14))
        self._name_edit.hide()
        name_layout.addWidget(self._name_edit)

        self._company_label = QLabel()
        self._company_label.setStyleSheet("color: gray; font-size: 12pt;")
        name_layout.addWidget(self._company_label)

        header.addLayout(name_layout, 1)
        layout.addLayout(header)

        # Action buttons
        btn_layout = QHBoxLayout()
        btn_layout.setSpacing(8)

        self._edit_btn = QPushButton("Edit")
        self._edit_btn.clicked.connect(self._enter_edit_mode)
        btn_layout.addWidget(self._edit_btn)

        self._new_msg_btn = QPushButton("New Message")
        self._new_msg_btn.clicked.connect(self._compose_new_message)
        btn_layout.addWidget(self._new_msg_btn)

        self._save_btn = QPushButton("Save")
        self._save_btn.clicked.connect(self._save_contact)
        self._save_btn.hide()
        btn_layout.addWidget(self._save_btn)

        self._cancel_btn = QPushButton("Cancel")
        self._cancel_btn.clicked.connect(self._cancel_edit)
        self._cancel_btn.hide()
        btn_layout.addWidget(self._cancel_btn)

        self._delete_btn = QPushButton("Delete")
        self._delete_btn.setStyleSheet("color: red;")
        self._delete_btn.clicked.connect(self._delete_contact)
        self._delete_btn.hide()
        btn_layout.addWidget(self._delete_btn)

        btn_layout.addStretch()
        layout.addLayout(btn_layout)

        # Scroll area for details
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        scroll_content = QWidget()
        self._form_layout = QVBoxLayout(scroll_content)
        self._form_layout.setContentsMargins(0, 0, 0, 0)
        self._form_layout.setSpacing(16)

        # Email section
        self._add_section("Email")
        self._email_editor = MultiValueEditor("email")
        self._form_layout.addWidget(self._email_editor)

        # Phone section
        self._add_section("Phone")
        self._phone_editor = MultiValueEditor("phone")
        self._form_layout.addWidget(self._phone_editor)

        # Work section
        self._add_section("Work")
        work_form = QFormLayout()
        work_form.setSpacing(8)

        self._company_edit = QLineEdit()
        self._company_edit.setPlaceholderText("Company")
        work_form.addRow("Company:", self._company_edit)

        self._job_title_edit = QLineEdit()
        self._job_title_edit.setPlaceholderText("Job Title")
        work_form.addRow("Job Title:", self._job_title_edit)

        self._department_edit = QLineEdit()
        self._department_edit.setPlaceholderText("Department")
        work_form.addRow("Department:", self._department_edit)

        work_widget = QWidget()
        work_widget.setLayout(work_form)
        self._form_layout.addWidget(work_widget)

        # Other section
        self._add_section("Other")
        other_form = QFormLayout()
        other_form.setSpacing(8)

        self._website_edit = QLineEdit()
        self._website_edit.setPlaceholderText("Website")
        other_form.addRow("Website:", self._website_edit)

        self._birthday_edit = QLineEdit()
        self._birthday_edit.setPlaceholderText("YYYY-MM-DD")
        other_form.addRow("Birthday:", self._birthday_edit)

        other_widget = QWidget()
        other_widget.setLayout(other_form)
        self._form_layout.addWidget(other_widget)

        # Notes section
        self._add_section("Notes")
        self._notes_edit = QTextEdit()
        self._notes_edit.setPlaceholderText("Notes...")
        self._notes_edit.setMaximumHeight(100)
        self._form_layout.addWidget(self._notes_edit)

        self._form_layout.addStretch()

        scroll.setWidget(scroll_content)
        layout.addWidget(scroll, 1)

        # Initial state - no contact
        self._show_empty_state()

    def _add_section(self, title: str):
        """Add a section header to the form."""
        label = QLabel(title)
        label.setFont(QFont("sans-serif", 11, QFont.Weight.Bold))
        label.setStyleSheet("color: #333; margin-top: 8px;")
        self._form_layout.addWidget(label)

    def _show_empty_state(self):
        """Show empty state when no contact selected."""
        self._name_label.setText("Select a contact")
        self._company_label.setText("")
        self._avatar.set_initials("")
        self._edit_btn.hide()
        self._new_msg_btn.hide()

    def set_contact(self, contact: dict):
        """Display a contact."""
        self._contact = contact
        self._edit_mode = False
        self._update_display()

    def _update_display(self):
        """Update the display from contact data."""
        if not self._contact:
            self._show_empty_state()
            return

        # Name
        name = self._contact.get("common_name", "")
        if not name:
            first = self._contact.get("first_name", "") or ""
            last = self._contact.get("last_name", "") or ""
            name = f"{first} {last}".strip()

        self._name_label.setText(name if name else "(No Name)")
        self._name_edit.setText(name)
        self._avatar.set_initials(name)

        # Company
        company = self._contact.get("company", "") or ""
        job = self._contact.get("job_title", "") or ""
        if company and job:
            self._company_label.setText(f"{job} at {company}")
        elif company:
            self._company_label.setText(company)
        elif job:
            self._company_label.setText(job)
        else:
            self._company_label.setText("")

        # Emails
        emails = self._contact.get("email_addresses", [])
        if isinstance(emails, str):
            try:
                emails = json.loads(emails)
            except (json.JSONDecodeError, TypeError):
                emails = []
        self._email_editor.set_values(emails)

        # Phones
        phones = self._contact.get("phone_numbers", [])
        if isinstance(phones, str):
            try:
                phones = json.loads(phones)
            except (json.JSONDecodeError, TypeError):
                phones = []
        self._phone_editor.set_values(phones)

        # Work fields
        self._company_edit.setText(self._contact.get("company", "") or "")
        self._job_title_edit.setText(self._contact.get("job_title", "") or "")
        self._department_edit.setText(self._contact.get("department", "") or "")

        # Other
        self._website_edit.setText(self._contact.get("website", "") or "")
        self._birthday_edit.setText(self._contact.get("birthday", "") or "")

        # Notes
        self._notes_edit.setText(self._contact.get("notes", "") or "")

        # Check if this is a GAL contact (read-only)
        is_gal = self._contact.get("is_gal", False)

        # Set read-only state and styling
        self._set_fields_editable(self._edit_mode)
        self._set_readonly_style(is_gal and not self._edit_mode)

        # Show/hide buttons based on contact type
        self._edit_btn.setVisible(not is_gal)  # No edit for GAL contacts
        self._save_btn.setVisible(self._edit_mode)
        self._cancel_btn.setVisible(self._edit_mode)
        self._delete_btn.setVisible(self._edit_mode and not is_gal)
        self._name_label.setVisible(not self._edit_mode)
        self._name_edit.setVisible(self._edit_mode)

        # Show New Message button if contact has an email
        has_email = bool(emails)
        self._new_msg_btn.setVisible(has_email and not self._edit_mode)

    def _set_fields_editable(self, editable: bool):
        """Set whether fields are editable."""
        self._company_edit.setReadOnly(not editable)
        self._job_title_edit.setReadOnly(not editable)
        self._department_edit.setReadOnly(not editable)
        self._website_edit.setReadOnly(not editable)
        self._birthday_edit.setReadOnly(not editable)
        self._notes_edit.setReadOnly(not editable)
        self._email_editor.set_edit_mode(editable)
        self._phone_editor.set_edit_mode(editable)

    def _set_readonly_style(self, readonly: bool):
        """Style fields as read-only (no border, no placeholder) or editable."""
        if readonly:
            style = "QLineEdit { border: none; background: transparent; }"
            self._company_edit.setStyleSheet(style)
            self._job_title_edit.setStyleSheet(style)
            self._department_edit.setStyleSheet(style)
            self._website_edit.setStyleSheet(style)
            self._birthday_edit.setStyleSheet(style)
            self._notes_edit.setStyleSheet("QTextEdit { border: none; background: transparent; }")
        else:
            self._company_edit.setStyleSheet("")
            self._job_title_edit.setStyleSheet("")
            self._department_edit.setStyleSheet("")
            self._website_edit.setStyleSheet("")
            self._birthday_edit.setStyleSheet("")
            self._notes_edit.setStyleSheet("")
        # Also style the multi-value editors
        self._email_editor.set_readonly_style(readonly)
        self._phone_editor.set_readonly_style(readonly)

    def _enter_edit_mode(self):
        """Enter edit mode."""
        self._edit_mode = True
        self._update_display()

    def _cancel_edit(self):
        """Cancel editing and revert to view mode."""
        self._edit_mode = False
        self._update_display()
        self.cancelRequested.emit()

    def _save_contact(self):
        """Save the contact."""
        if not self._contact:
            return

        # Gather data from form
        contact_data = {
            "item_id": self._contact.get("item_id", ""),
            "common_name": self._name_edit.text().strip(),
            "company": self._company_edit.text().strip(),
            "job_title": self._job_title_edit.text().strip(),
            "department": self._department_edit.text().strip(),
            "website": self._website_edit.text().strip(),
            "birthday": self._birthday_edit.text().strip(),
            "notes": self._notes_edit.toPlainText().strip(),
            "email_addresses": self._email_editor.get_values(),
            "phone_numbers": self._phone_editor.get_values(),
        }

        self._edit_mode = False
        self.saveRequested.emit(contact_data)

    def _delete_contact(self):
        """Delete the contact after confirmation."""
        if not self._contact:
            return

        name = self._contact.get("common_name", "this contact")
        reply = QMessageBox.question(
            self, "Delete Contact",
            f"Are you sure you want to delete {name}?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if reply == QMessageBox.StandardButton.Yes:
            self.deleteRequested.emit(self._contact.get("item_id", ""))

    def _compose_new_message(self):
        """Compose a new message to this contact."""
        if not self._contact:
            return

        # Get primary email
        emails = self._contact.get("email_addresses", [])
        if isinstance(emails, str):
            try:
                emails = json.loads(emails)
            except (json.JSONDecodeError, TypeError):
                emails = []

        if not emails:
            return

        first_email = emails[0]
        email = first_email.get("address", "") if isinstance(first_email, dict) else str(first_email)

        # Get name
        name = self._contact.get("common_name", "")
        if not name:
            first = self._contact.get("first_name", "") or ""
            last = self._contact.get("last_name", "") or ""
            name = f"{first} {last}".strip()

        self.newMessageRequested.emit(email, name)

    def clear(self):
        """Clear the view."""
        self._contact = None
        self._edit_mode = False
        self._show_empty_state()

    def start_new_contact(self, folder_id: str):
        """Start creating a new contact."""
        self._contact = {
            "item_id": "",
            "folder_id": folder_id,
            "common_name": "",
            "email_addresses": [],
            "phone_numbers": [],
        }
        self._edit_mode = True
        self._update_display()
        self._name_edit.setFocus()
