"""Qt Compose email widget for inline composition in preview pane."""

import os
import re
import html as html_module
from datetime import datetime

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QLineEdit, QComboBox, QTextEdit, QPushButton,
    QFileDialog, QMessageBox, QSizePolicy, QFrame, QCompleter
)
from PySide6.QtCore import Qt, Signal, QStringListModel
from PySide6.QtGui import (
    QFont, QTextCharFormat, QAction, QKeySequence, QTextCursor
)


class AddressLineEdit(QLineEdit):
    """Line edit with email address autocomplete."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._addresses = []  # Display strings: "Name <email>" or "email"
        self._completer = QCompleter(self)
        self._completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self._completer.setFilterMode(Qt.MatchFlag.MatchContains)
        self._completer.activated.connect(self._insert_completion)
        self.setCompleter(self._completer)

    def set_address_book(self, addresses):
        """Set the address book for autocompletion."""
        # Sort: users (QLF) first, then contacts, then recent
        sorted_addrs = []
        for addr in addresses:
            addr_type = addr.get('type', '')
            if addr_type == 'user':
                priority = 0
            elif addr_type == 'contact':
                priority = 1
            else:
                priority = 2
            sorted_addrs.append((priority, addr))

        sorted_addrs.sort(key=lambda x: (x[0], x[1].get('name', '').lower()))

        # Build display strings
        self._addresses = []
        for _, addr in sorted_addrs:
            name = addr.get('name', '')
            email = addr.get('email', '')
            if name:
                display = f"{name} <{email}>"
            else:
                display = email
            self._addresses.append(display)

        self._completer.setModel(QStringListModel(self._addresses, self._completer))

    def _insert_completion(self, text):
        """Insert selected completion, handling comma-separated addresses."""
        current = self.text()
        if ',' in current:
            prefix = current.rsplit(',', 1)[0] + ', '
        else:
            prefix = ""

        self.setText(prefix + text)
        self.setCursorPosition(len(self.text()))


class ComposeWidget(QWidget):
    """Inline compose email widget that replaces the preview pane."""

    message_sent = Signal()
    compose_cancelled = Signal()

    def __init__(self, parent, db, sync_manager, account_id=None,
                 reply_to=None, forward=None, attachments=None, signature=None,
                 font_size=12, zoom=100):
        super().__init__(parent)
        self.db = db
        self.sync_manager = sync_manager
        self.account_id = account_id
        self._attachments = attachments or []
        self._accounts = []
        self._compose_type = "New Message"
        self._signature = signature or ""
        self._font_size = font_size
        self._zoom = zoom / 100.0  # Convert percentage to factor

        self._is_dark = db.get_setting("dark_mode", "0") == "1"

        self._create_ui()
        self._setup_shortcuts()
        self._apply_theme()

        if reply_to:
            self._setup_reply(reply_to)
        elif forward:
            self._setup_forward(forward)
        else:
            # New message - insert signature
            self._setup_new_message()

        # Store initial state for change detection
        self._initial_state = self._get_current_state()

    def _create_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(5, 5, 5, 5)
        layout.setSpacing(5)

        # Title bar with action buttons
        title_bar = QHBoxLayout()

        self._title_label = QLabel("New Message")
        self._title_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        title_bar.addWidget(self._title_label)
        title_bar.addStretch()

        send_btn = QPushButton("Send")
        send_btn.clicked.connect(self._send)
        title_bar.addWidget(send_btn)

        attach_btn = QPushButton("Attach")
        attach_btn.clicked.connect(self._attach)
        title_bar.addWidget(attach_btn)

        discard_btn = QPushButton("Discard")
        discard_btn.clicked.connect(self._discard)
        title_bar.addWidget(discard_btn)

        layout.addLayout(title_bar)

        # Separator
        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setFixedHeight(2)
        layout.addWidget(sep)

        # Header fields
        header = QGridLayout()
        header.setColumnStretch(1, 1)

        # From
        header.addWidget(QLabel("From:"), 0, 0)
        self.from_combo = QComboBox()
        self._accounts = self.db.get_accounts()
        for acc in self._accounts:
            self.from_combo.addItem(f"{acc['name']} <{acc['email']}>")

        # Select default or specified account
        for i, acc in enumerate(self._accounts):
            if acc.get('is_default') or (self.account_id and acc['id'] == self.account_id):
                self.from_combo.setCurrentIndex(i)
                break
        header.addWidget(self.from_combo, 0, 1)

        # To (with autocomplete)
        header.addWidget(QLabel("To:"), 1, 0)
        self.to_edit = AddressLineEdit()
        self.to_edit.setPlaceholderText("recipient@example.com")
        header.addWidget(self.to_edit, 1, 1)

        # CC (with autocomplete)
        header.addWidget(QLabel("CC:"), 2, 0)
        self.cc_edit = AddressLineEdit()
        header.addWidget(self.cc_edit, 2, 1)

        # Subject
        header.addWidget(QLabel("Subject:"), 3, 0)
        self.subject_edit = QLineEdit()
        header.addWidget(self.subject_edit, 3, 1)

        layout.addLayout(header)

        # Formatting toolbar
        format_bar = QHBoxLayout()

        self.bold_btn = QPushButton("B")
        self.bold_btn.setCheckable(True)
        self.bold_btn.setFixedWidth(30)
        self.bold_btn.setFont(QFont("", -1, QFont.Weight.Bold))
        self.bold_btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)  # Skip in tab order
        self.bold_btn.clicked.connect(self._toggle_bold)
        format_bar.addWidget(self.bold_btn)

        self.italic_btn = QPushButton("I")
        self.italic_btn.setCheckable(True)
        self.italic_btn.setFixedWidth(30)
        font = self.italic_btn.font()
        font.setItalic(True)
        self.italic_btn.setFont(font)
        self.italic_btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.italic_btn.clicked.connect(self._toggle_italic)
        format_bar.addWidget(self.italic_btn)

        self.underline_btn = QPushButton("U")
        self.underline_btn.setCheckable(True)
        self.underline_btn.setFixedWidth(30)
        font = self.underline_btn.font()
        font.setUnderline(True)
        self.underline_btn.setFont(font)
        self.underline_btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.underline_btn.clicked.connect(self._toggle_underline)
        format_bar.addWidget(self.underline_btn)

        format_bar.addWidget(QLabel("  Size:"))
        self.size_combo = QComboBox()
        self.size_combo.addItems(["10", "12", "14", "16", "18", "20", "24", "28"])
        self.size_combo.setCurrentText("12")
        self.size_combo.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        self.size_combo.currentTextChanged.connect(self._change_font_size)
        format_bar.addWidget(self.size_combo)

        bullet_btn = QPushButton("• List")
        bullet_btn.setFocusPolicy(Qt.FocusPolicy.NoFocus)
        bullet_btn.clicked.connect(self._insert_bullet)
        format_bar.addWidget(bullet_btn)

        format_bar.addStretch()

        # Attachments display
        self.attach_label = QLabel("")
        format_bar.addWidget(self.attach_label)

        layout.addLayout(format_bar)

        # Body editor
        self.body_edit = QTextEdit()
        # Use a proper sans-serif font with zoom applied
        body_font = QFont("Sans Serif")
        body_font.setPointSize(int(self._font_size * self._zoom))
        self.body_edit.setFont(body_font)
        self.body_edit.cursorPositionChanged.connect(self._update_format_buttons)
        layout.addWidget(self.body_edit, 1)

        self._update_attachments_display()

    def _setup_shortcuts(self):
        """Setup keyboard shortcuts."""
        # Ctrl+B for bold
        bold_shortcut = QAction(self)
        bold_shortcut.setShortcut(QKeySequence("Ctrl+B"))
        bold_shortcut.triggered.connect(self._toggle_bold)
        self.addAction(bold_shortcut)

        # Ctrl+I for italic
        italic_shortcut = QAction(self)
        italic_shortcut.setShortcut(QKeySequence("Ctrl+I"))
        italic_shortcut.triggered.connect(self._toggle_italic)
        self.addAction(italic_shortcut)

        # Ctrl+U for underline
        underline_shortcut = QAction(self)
        underline_shortcut.setShortcut(QKeySequence("Ctrl+U"))
        underline_shortcut.triggered.connect(self._toggle_underline)
        self.addAction(underline_shortcut)

        # Ctrl+Enter to send
        send_shortcut = QAction(self)
        send_shortcut.setShortcut(QKeySequence("Ctrl+Return"))
        send_shortcut.triggered.connect(self._send)
        self.addAction(send_shortcut)

        # Escape to discard
        esc_shortcut = QAction(self)
        esc_shortcut.setShortcut(QKeySequence("Escape"))
        esc_shortcut.triggered.connect(self._discard)
        self.addAction(esc_shortcut)

    def _toggle_bold(self):
        """Toggle bold formatting."""
        fmt = QTextCharFormat()
        if self.body_edit.fontWeight() == QFont.Weight.Bold:
            fmt.setFontWeight(QFont.Weight.Normal)
        else:
            fmt.setFontWeight(QFont.Weight.Bold)
        self._merge_format(fmt)
        self._update_format_buttons()

    def _toggle_italic(self):
        """Toggle italic formatting."""
        fmt = QTextCharFormat()
        fmt.setFontItalic(not self.body_edit.fontItalic())
        self._merge_format(fmt)
        self._update_format_buttons()

    def _toggle_underline(self):
        """Toggle underline formatting."""
        fmt = QTextCharFormat()
        fmt.setFontUnderline(not self.body_edit.fontUnderline())
        self._merge_format(fmt)
        self._update_format_buttons()

    def _change_font_size(self, size_str):
        """Change font size."""
        try:
            size = int(size_str)
            fmt = QTextCharFormat()
            fmt.setFontPointSize(size)
            self._merge_format(fmt)
        except ValueError:
            pass

    def _merge_format(self, fmt: QTextCharFormat):
        """Apply format to selection or current position."""
        cursor = self.body_edit.textCursor()
        if not cursor.hasSelection():
            cursor.select(QTextCursor.SelectionType.WordUnderCursor)
        cursor.mergeCharFormat(fmt)
        self.body_edit.mergeCurrentCharFormat(fmt)

    def _update_format_buttons(self):
        """Update button states based on current format."""
        self.bold_btn.setChecked(self.body_edit.fontWeight() == QFont.Weight.Bold)
        self.italic_btn.setChecked(self.body_edit.fontItalic())
        self.underline_btn.setChecked(self.body_edit.fontUnderline())

    def _insert_bullet(self):
        """Insert bullet point."""
        self.body_edit.insertPlainText("• ")

    def _attach(self):
        """Add attachments."""
        files, _ = QFileDialog.getOpenFileNames(self, "Select files to attach")
        for filepath in files:
            try:
                with open(filepath, 'rb') as f:
                    content = f.read()
                name = os.path.basename(filepath)
                self._attachments.append({'name': name, 'content': content, 'path': filepath})
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to read {filepath}: {e}")

        self._update_attachments_display()

    def _update_attachments_display(self):
        """Update attachments label."""
        if self._attachments:
            names = [a['name'] for a in self._attachments]
            self.attach_label.setText(f"Attachments: {', '.join(names)}")
        else:
            self.attach_label.setText("")

    def _html_to_plain_text(self, html_body):
        """Convert HTML to plain text, properly stripping style/script blocks."""
        if not html_body:
            return ""

        text = html_body

        # Remove style and script blocks entirely (including their content)
        text = re.sub(r'<style[^>]*>.*?</style>', '', text, flags=re.IGNORECASE | re.DOTALL)
        text = re.sub(r'<script[^>]*>.*?</script>', '', text, flags=re.IGNORECASE | re.DOTALL)
        text = re.sub(r'<head[^>]*>.*?</head>', '', text, flags=re.IGNORECASE | re.DOTALL)

        # Convert common block elements to newlines
        text = re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)
        text = re.sub(r'<p[^>]*>', '\n', text, flags=re.IGNORECASE)
        text = re.sub(r'</p>', '\n', text, flags=re.IGNORECASE)
        text = re.sub(r'<div[^>]*>', '\n', text, flags=re.IGNORECASE)
        text = re.sub(r'</div>', '', text, flags=re.IGNORECASE)
        text = re.sub(r'<tr[^>]*>', '\n', text, flags=re.IGNORECASE)
        text = re.sub(r'<li[^>]*>', '\n• ', text, flags=re.IGNORECASE)

        # Strip remaining tags
        text = re.sub(r'<[^>]+>', '', text)

        # Decode HTML entities
        text = html_module.unescape(text)

        # Clean up whitespace
        text = re.sub(r'\n{3,}', '\n\n', text)
        text = re.sub(r'[ \t]+', ' ', text)
        lines = [line.strip() for line in text.split('\n')]
        text = '\n'.join(lines).strip()

        return text

    def _get_signature_html(self):
        """Get signature formatted as HTML."""
        if not self._signature:
            return ""
        # Signature may be HTML or plain text
        sig = self._signature
        if '<' not in sig:
            # Plain text - escape and convert newlines
            sig = html_module.escape(sig).replace('\n', '<br>')
        # Return signature as-is - QTextEdit should handle <font> tags
        return f'<p><br></p>{sig}'

    def _setup_new_message(self):
        """Setup for new message with signature."""
        if self._signature:
            html_content = f'''<html><body>
<p><br></p>
{self._get_signature_html()}
</body></html>'''
            self.body_edit.setHtml(html_content)

            # Move cursor to beginning
            cursor = self.body_edit.textCursor()
            cursor.movePosition(QTextCursor.MoveOperation.Start)
            self.body_edit.setTextCursor(cursor)

        self.to_edit.setFocus()

    def _setup_reply(self, original_message):
        """Setup for reply."""
        reply_all = original_message.get('reply_all', False)
        self._compose_type = "Reply All" if reply_all else "Reply"
        self._title_label.setText(self._compose_type)

        from_name = original_message.get('from_name', '')
        from_email = original_message.get('from_email', '')

        # Set To field
        if from_name and from_email:
            self.to_edit.setText(f"{from_name} <{from_email}>")
        elif from_email:
            self.to_edit.setText(from_email)

        # Reply All - add CC
        if reply_all:
            my_email = ''
            idx = self.from_combo.currentIndex()
            if 0 <= idx < len(self._accounts):
                my_email = self._accounts[idx].get('email', '').lower()

            cc_list = []
            original_to = original_message.get('to', '')
            original_cc = original_message.get('cc', '')

            for addr in (original_to + ',' + original_cc).split(','):
                addr = addr.strip()
                if not addr:
                    continue
                addr_lower = addr.lower()
                if my_email and my_email in addr_lower:
                    continue
                if from_email and from_email.lower() in addr_lower:
                    continue
                cc_list.append(addr)

            if cc_list:
                self.cc_edit.setText(', '.join(cc_list))

        # Subject
        subject = original_message.get('subject', '')
        if not subject.lower().startswith('re:'):
            subject = f"Re: {subject}"
        self.subject_edit.setText(subject)

        # Quote original with left border bar style
        orig_body = original_message.get('body', '')
        date = original_message.get('date', '')

        if date and len(date) >= 15:
            try:
                date_part = date[:8]
                time_part = date[9:15]
                dt = datetime.strptime(f"{date_part}{time_part}", "%Y%m%d%H%M%S")
                date = dt.strftime("%B %d, %Y at %I:%M %p")
            except Exception:
                pass

        sender_str = f"{from_name} &lt;{from_email}&gt;" if from_name else html_module.escape(from_email)
        quote_header = f"On {date}, {sender_str} wrote:" if date else f"{sender_str} wrote:"

        # Convert HTML to plain text and escape for HTML
        plain_body = self._html_to_plain_text(orig_body)
        escaped_body = html_module.escape(plain_body).replace('\n', '<br>')

        # Build HTML with table-based left border (QTextEdit doesn't support border-left CSS)
        # Signature goes between reply area and quoted text
        signature_html = self._get_signature_html()

        html_content = f'''<html><body>
<p><br></p>
{signature_html}
<p style="color: #666;">{quote_header}</p>
<table border="0" cellspacing="0" cellpadding="0">
<tr>
<td style="background-color: #ccc; width: 3px;"></td>
<td style="padding-left: 10px; color: #555;">{escaped_body}</td>
</tr>
</table>
</body></html>'''

        self.body_edit.setHtml(html_content)

        # Move cursor to beginning
        cursor = self.body_edit.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.Start)
        self.body_edit.setTextCursor(cursor)
        self.body_edit.setFocus()

    def _setup_forward(self, original_message):
        """Setup for forward."""
        self._compose_type = "Forward"
        self._title_label.setText(self._compose_type)

        subject = original_message.get('subject', '')
        if not subject.lower().startswith('fwd:') and not subject.lower().startswith('fw:'):
            subject = f"Fwd: {subject}"
        self.subject_edit.setText(subject)

        # Build forwarded message body
        from_name = original_message.get('from_name', '')
        from_email = original_message.get('from_email', '')
        date = original_message.get('date', '')
        to = original_message.get('to', '')
        body = original_message.get('body', '')

        if date and len(date) >= 15:
            try:
                date_part = date[:8]
                time_part = date[9:15]
                dt = datetime.strptime(f"{date_part}{time_part}", "%Y%m%d%H%M%S")
                date = dt.strftime("%B %d, %Y at %I:%M %p")
            except Exception:
                pass

        # Escape for HTML
        sender_str = f"{html_module.escape(from_name)} &lt;{html_module.escape(from_email)}&gt;" if from_name else html_module.escape(from_email)
        escaped_subject = html_module.escape(original_message.get('subject', ''))
        escaped_to = html_module.escape(to)

        # Convert HTML to plain text and escape for HTML
        plain_body = self._html_to_plain_text(body)
        escaped_body = html_module.escape(plain_body).replace('\n', '<br>')

        # Build HTML with table-based left border (QTextEdit doesn't support border-left CSS)
        # Signature goes between message area and forwarded content
        signature_html = self._get_signature_html()

        html_content = f'''<html><body>
<p><br></p>
{signature_html}
<p style="color: #666;">---------- Forwarded message ----------</p>
<table border="0" cellspacing="0" cellpadding="0">
<tr>
<td style="background-color: #ccc; width: 3px;"></td>
<td style="padding-left: 10px;">
<p style="color: #666; margin: 0;">
<b>From:</b> {sender_str}<br>
<b>Date:</b> {html_module.escape(date)}<br>
<b>Subject:</b> {escaped_subject}<br>
<b>To:</b> {escaped_to}
</p>
<p style="color: #555;">{escaped_body}</p>
</td>
</tr>
</table>
</body></html>'''

        self.body_edit.setHtml(html_content)

        cursor = self.body_edit.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.Start)
        self.body_edit.setTextCursor(cursor)
        self.to_edit.setFocus()

    def _get_html_body(self):
        """Convert rich text to HTML."""
        return self.body_edit.toHtml()

    def _send(self):
        """Send the message."""
        to = self.to_edit.text().strip()
        if not to:
            QMessageBox.warning(self, "Error", "To field is required")
            return

        to_list = [addr.strip() for addr in to.split(',') if addr.strip()]
        cc = self.cc_edit.text().strip()
        cc_list = [addr.strip() for addr in cc.split(',') if addr.strip()] if cc else []

        subject = self.subject_edit.text().strip()
        body = self._get_html_body()

        from_idx = self.from_combo.currentIndex()
        if from_idx < 0 or from_idx >= len(self._accounts):
            QMessageBox.warning(self, "Error", "Please select a From account")
            return

        account_id = self._accounts[from_idx]['id']

        self._title_label.setText("Sending...")

        self.sync_manager.send_message(
            account_id=account_id,
            to=to_list,
            subject=subject,
            body=body,
            cc=cc_list,
            attachments=self._attachments,
            callback=self._on_sent
        )

    def _on_sent(self, success, message):
        """Handle send completion."""
        if success:
            self.message_sent.emit()
        else:
            QMessageBox.warning(self, "Error", f"Failed to send: {message}")
            self._title_label.setText(self._compose_type)

    def _get_current_state(self):
        """Get current state of all fields for change detection."""
        return {
            'to': self.to_edit.text(),
            'cc': self.cc_edit.text(),
            'subject': self.subject_edit.text(),
            'body': self.body_edit.toHtml(),
            'from_idx': self.from_combo.currentIndex(),
            'attachments': len(self._attachments)
        }

    def _has_changes(self):
        """Check if any fields have changed from initial state."""
        current = self._get_current_state()
        return current != self._initial_state

    def _discard(self):
        """Discard message."""
        # Only confirm if content has actually changed
        if self._has_changes():
            result = QMessageBox.question(self, "Discard", "Discard this message?")
            if result != QMessageBox.StandardButton.Yes:
                return

        self.compose_cancelled.emit()

    def focus_to_field(self):
        """Focus the To field."""
        self.to_edit.setFocus()

    def focus_body(self):
        """Focus the body editor."""
        self.body_edit.setFocus()

    def set_address_book(self, addresses):
        """Set the address book for To/CC autocomplete."""
        self.to_edit.set_address_book(addresses)
        self.cc_edit.set_address_book(addresses)

    def set_zoom(self, zoom_factor):
        """Set zoom level for the compose body."""
        self._zoom = zoom_factor
        size = int(self._font_size * zoom_factor)
        font = QFont("Sans Serif")
        font.setPointSize(size)
        self.body_edit.setFont(font)
        self.body_edit.viewport().update()

    def _apply_theme(self):
        """Apply dark/light theme."""
        if self._is_dark:
            self.setStyleSheet("""
                QWidget { background-color: #2b2b2b; color: #a9b7c6; }
                QTextEdit, QLineEdit { background-color: #313335; color: #a9b7c6; border: 1px solid #3c3f41; }
                QComboBox { background-color: #313335; color: #a9b7c6; border: 1px solid #3c3f41; }
                QPushButton { background-color: #3c3f41; color: #a9b7c6; padding: 5px 10px; border: none; }
                QPushButton:hover { background-color: #4c5052; }
                QPushButton:checked { background-color: #214283; }
            """)
        else:
            self.setStyleSheet("""
                QPushButton:checked { background-color: #0078d4; color: white; }
            """)
