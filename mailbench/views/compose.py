"""Qt Compose email widget for inline composition in preview pane."""

import os
import re
import html as html_module
from datetime import datetime

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QLineEdit, QComboBox, QTextEdit, QPushButton,
    QFileDialog, QMessageBox, QSizePolicy, QFrame, QCompleter,
    QMenu, QToolButton, QApplication
)
from PySide6.QtCore import Qt, Signal, QStringListModel, QUrl
from PySide6.QtGui import (
    QFont, QTextCharFormat, QAction, QKeySequence, QTextCursor,
    QTextBlockFormat
)

# Check for WebEngine availability
HAS_WEBENGINE = False
try:
    from PySide6.QtWebEngineWidgets import QWebEngineView
    from PySide6.QtWebEngineCore import QWebEnginePage, QWebEngineSettings
    from PySide6.QtWebChannel import QWebChannel
    HAS_WEBENGINE = True
except ImportError:
    pass


class RichTextEdit(QTextEdit):
    """QTextEdit with guaranteed clipboard support."""

    def keyPressEvent(self, event):
        """Handle key press events, ensuring clipboard operations work."""
        # Check for Ctrl+V (paste)
        if event.key() == Qt.Key.Key_V and event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            self._do_paste()
            event.accept()
            return
        # Check for Ctrl+C (copy)
        if event.key() == Qt.Key.Key_C and event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            self.copy()
            event.accept()
            return
        # Check for Ctrl+X (cut)
        if event.key() == Qt.Key.Key_X and event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            self.cut()
            event.accept()
            return
        # Check for Ctrl+A (select all)
        if event.key() == Qt.Key.Key_A and event.modifiers() == Qt.KeyboardModifier.ControlModifier:
            self.selectAll()
            event.accept()
            return
        super().keyPressEvent(event)

    def _do_paste(self):
        """Paste clipboard contents at cursor position."""
        clipboard = QApplication.clipboard()
        mime = clipboard.mimeData()

        # Try image first (for screenshots)
        if mime.hasImage():
            image = clipboard.image()
            if not image.isNull():
                # Convert image to base64 data URI
                from PySide6.QtCore import QBuffer, QByteArray
                buffer = QByteArray()
                qbuffer = QBuffer(buffer)
                qbuffer.open(QBuffer.OpenModeFlag.WriteOnly)
                image.save(qbuffer, "PNG")
                qbuffer.close()
                import base64
                b64_data = base64.b64encode(buffer.data()).decode('ascii')
                img_html = f'<img src="data:image/png;base64,{b64_data}"/>'
                self.insertHtml(img_html)
                return

        if mime.hasHtml():
            self.insertHtml(mime.html())
        elif mime.hasText():
            self.insertPlainText(mime.text())


if HAS_WEBENGINE:
    class WebEngineEditor(QWebEngineView):
        """QWebEngineView-based rich text editor with contentEditable."""

        def __init__(self, parent=None, font_size=12):
            super().__init__(parent)
            self._font_size = font_size
            self._html_content = ""
            self._ready = False

            # Create page and allow JavaScript
            page = QWebEnginePage(self)
            self.setPage(page)

            # Enable settings needed for editing
            settings = self.settings()
            settings.setAttribute(QWebEngineSettings.WebAttribute.JavascriptEnabled, True)
            settings.setAttribute(QWebEngineSettings.WebAttribute.LocalContentCanAccessRemoteUrls, True)
            settings.setAttribute(QWebEngineSettings.WebAttribute.FocusOnNavigationEnabled, True)

            # Load editable HTML template
            self._load_editor()

            # Connect load finished to set ready flag
            self.loadFinished.connect(self._on_load_finished)

        def _load_editor(self):
            """Load the editable HTML template."""
            html = f'''<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
body {{
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    font-size: {self._font_size}pt;
    margin: 8px;
    padding: 0;
    line-height: 1.2;
}}
p {{ margin: 0; padding: 0; }}
</style>
<script>
document.addEventListener('keydown', function(e) {{
    if (e.key === 'Tab') {{
        e.preventDefault();
        document.execCommand('insertText', false, '\\t');
    }}
}});
</script>
</head>
<body contenteditable="true" id="editor">
</body>
</html>'''
            self.setHtml(html)

        def _on_load_finished(self, ok):
            """Called when page finishes loading."""
            self._ready = ok
            if ok and self._html_content:
                # Set any pending content
                self._set_inner_html(self._html_content)
                self._html_content = ""

        def _set_inner_html(self, html):
            """Set the innerHTML of the editor body."""
            # Escape the HTML for JavaScript string
            escaped = html.replace('\\', '\\\\').replace('`', '\\`').replace('$', '\\$')
            js = f'document.getElementById("editor").innerHTML = `{escaped}`;'
            self.page().runJavaScript(js)

        def setHtml(self, html):
            """Set HTML content (compatible with QTextEdit API)."""
            # Extract body content if full HTML document
            body_match = re.search(r'<body[^>]*>(.*)</body>', html, re.IGNORECASE | re.DOTALL)
            if body_match:
                content = body_match.group(1)
            else:
                content = html

            if self._ready:
                self._set_inner_html(content)
            else:
                # Store for when ready
                self._html_content = content
                # Also call parent setHtml to load the template
                super().setHtml(self._get_template_with_content(content))

        def _get_template_with_content(self, content):
            """Get the editor template with content pre-filled."""
            return f'''<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
body {{
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    font-size: {self._font_size}pt;
    margin: 8px;
    padding: 0;
    line-height: 1.2;
}}
p {{ margin: 0; padding: 0; }}
</style>
<script>
document.addEventListener('keydown', function(e) {{
    if (e.key === 'Tab') {{
        e.preventDefault();
        document.execCommand('insertText', false, '\\t');
    }}
}});
</script>
</head>
<body contenteditable="true" id="editor">
{content}
</body>
</html>'''

        def toHtml(self, callback=None):
            """Get HTML content asynchronously.

            If callback is provided, calls callback(html) when ready.
            Otherwise returns cached content (may be stale).
            """
            if callback:
                def handle_result(html):
                    # Wrap in full HTML document
                    full_html = f'''<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
body {{
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    font-size: {self._font_size}pt;
    line-height: 1.2;
}}
p {{ margin: 0; padding: 0; }}
</style>
</head>
<body>
{html}
</body>
</html>'''
                    callback(full_html)

                self.page().runJavaScript(
                    'document.getElementById("editor").innerHTML',
                    handle_result
                )
            return self._html_content

        def toHtmlSync(self):
            """Get HTML content synchronously using event loop.

            This blocks until the JavaScript returns.
            """
            from PySide6.QtCore import QEventLoop
            loop = QEventLoop()
            result = [None]

            def callback(html):
                result[0] = f'''<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<style>
body {{
    font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;
    font-size: {self._font_size}pt;
    line-height: 1.2;
}}
p {{ margin: 0; padding: 0; }}
</style>
</head>
<body>
{html}
</body>
</html>'''
                loop.quit()

            self.page().runJavaScript(
                'document.getElementById("editor").innerHTML',
                callback
            )
            loop.exec()
            return result[0]

        def execCommand(self, command, value=None):
            """Execute a document.execCommand for formatting."""
            if value:
                js = f'document.execCommand("{command}", false, "{value}")'
            else:
                js = f'document.execCommand("{command}", false, null)'
            self.page().runJavaScript(js)

        def toggleBold(self):
            """Toggle bold formatting."""
            self.execCommand("bold")

        def toggleItalic(self):
            """Toggle italic formatting."""
            self.execCommand("italic")

        def toggleUnderline(self):
            """Toggle underline formatting."""
            self.execCommand("underline")

        def insertBullet(self):
            """Insert unordered list."""
            self.execCommand("insertUnorderedList")

        def setFontSize(self, size):
            """Set font size."""
            self._font_size = size
            # fontSize command uses 1-7 scale, not pt
            # We'll use CSS instead
            self.page().runJavaScript(
                f'document.getElementById("editor").style.fontSize = "{size}pt"'
            )

        def fontWeight(self):
            """Check if current selection is bold (for button state)."""
            # Return a placeholder - actual state checking is complex with WebEngine
            return QFont.Weight.Normal

        def fontItalic(self):
            """Check if current selection is italic."""
            return False

        def fontUnderline(self):
            """Check if current selection is underlined."""
            return False

        def setFocus(self):
            """Set focus to the editor."""
            super().setFocus()
            self.page().runJavaScript('document.getElementById("editor").focus()')

        def clear(self):
            """Clear the editor content."""
            if self._ready:
                self.page().runJavaScript('document.getElementById("editor").innerHTML = ""')
            self._html_content = ""

        def setFont(self, font):
            """Set the editor font (compatibility method)."""
            self._font_size = font.pointSize()
            if self._ready:
                self.page().runJavaScript(
                    f'document.getElementById("editor").style.fontSize = "{self._font_size}pt"'
                )

        def textCursor(self):
            """Return a dummy cursor for compatibility."""
            return _DummyCursor()

        def setTextCursor(self, cursor):
            """Set cursor position (no-op for WebEngine)."""
            pass

        def cursorPositionChanged(self):
            """Dummy signal for compatibility."""
            pass

    class _DummyCursor:
        """Dummy cursor class for API compatibility."""
        def movePosition(self, *args, **kwargs):
            pass
        def setBlockFormat(self, *args, **kwargs):
            pass


class AddressLineEdit(QLineEdit):
    """Line edit with email address autocomplete supporting multiple addresses."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._addresses = []  # Display strings: "Name <email>" or "email"
        self._completer = QCompleter(self)
        self._completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self._completer.setFilterMode(Qt.MatchFlag.MatchContains)
        self._completer.setWidget(self)
        self._completer.activated.connect(self._insert_completion)
        # Don't use setCompleter - we manage it manually for multi-address support
        self.textChanged.connect(self._on_text_changed)

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

        sorted_addrs.sort(key=lambda x: (x[0], (x[1].get('name') or '').lower()))

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

    def _get_current_prefix(self):
        """Get the text before the current address being typed."""
        text = self.text()
        cursor_pos = self.cursorPosition()
        text_before_cursor = text[:cursor_pos]
        if ',' in text_before_cursor:
            return text_before_cursor.rsplit(',', 1)[0] + ', '
        return ""

    def _get_current_address_text(self):
        """Get the address currently being typed (after the last comma)."""
        text = self.text()
        cursor_pos = self.cursorPosition()
        text_before_cursor = text[:cursor_pos]
        if ',' in text_before_cursor:
            return text_before_cursor.rsplit(',', 1)[1].strip()
        return text_before_cursor.strip()

    def _on_text_changed(self, text):
        """Handle text changes to show completer for current address."""
        current_addr = self._get_current_address_text()
        if len(current_addr) >= 1:
            self._completer.setCompletionPrefix(current_addr)
            if self._completer.completionCount() > 0:
                self._completer.complete()
        else:
            self._completer.popup().hide()

    def _insert_completion(self, text):
        """Insert selected completion, handling comma-separated addresses."""
        prefix = self._get_current_prefix()
        # Get any text after cursor (in case user is editing in the middle)
        cursor_pos = self.cursorPosition()
        full_text = self.text()
        text_after_cursor = full_text[cursor_pos:]

        # Find if there's more text after current address
        suffix = ""
        if ',' in text_after_cursor:
            suffix = ',' + text_after_cursor.split(',', 1)[1]

        self.setText(prefix + text + suffix)
        # Position cursor after the inserted address
        self.setCursorPosition(len(prefix + text))


class ComposeWidget(QWidget):
    """Inline compose email widget that replaces the preview pane."""

    message_sent = Signal()
    compose_cancelled = Signal()

    def __init__(self, parent, db, sync_manager, account_id=None,
                 reply_to=None, forward=None, attachments=None, signature=None,
                 font_size=12, zoom=100, default_attach_dir=None):
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
        self._default_attach_dir = default_attach_dir or ""

        self._is_dark = False  # Always light mode

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

        # Attachments area (hidden when empty)
        self._attachments_widget = QWidget()
        self._attachments_layout = QVBoxLayout(self._attachments_widget)
        self._attachments_layout.setContentsMargins(0, 5, 0, 5)
        self._attachments_layout.setSpacing(5)
        self._attachments_widget.hide()
        layout.addWidget(self._attachments_widget)

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
        layout.addLayout(format_bar)

        # Body editor - use WebEngine if available for proper HTML/image support
        # Use base font size for content, apply zoom separately for display only
        if HAS_WEBENGINE:
            self.body_edit = WebEngineEditor(font_size=self._font_size)
            self.body_edit.setZoomFactor(self._zoom)  # Zoom is display-only
            self._use_webengine = True
        else:
            self.body_edit = RichTextEdit()
            body_font = QFont("Sans Serif")
            body_font.setPointSize(int(self._font_size * self._zoom))  # QTextEdit needs font scaling
            self.body_edit.setFont(body_font)
            self.body_edit.cursorPositionChanged.connect(self._update_format_buttons)
            self.body_edit.setTabStopDistance(40)
            self._set_default_paragraph_format()
            self._use_webengine = False

        layout.addWidget(self.body_edit, 1)

        self._update_attachments_display()

    def _set_default_paragraph_format(self):
        """Set default paragraph format with tighter spacing."""
        # Set default block format for new paragraphs
        block_fmt = QTextBlockFormat()
        block_fmt.setTopMargin(2)
        block_fmt.setBottomMargin(2)

        # Apply to current cursor position (affects new text)
        cursor = self.body_edit.textCursor()
        cursor.setBlockFormat(block_fmt)
        self.body_edit.setTextCursor(cursor)

    def _setup_shortcuts(self):
        """Setup keyboard shortcuts."""
        # Use WidgetWithChildrenShortcut context so child widgets still get their shortcuts
        context = Qt.ShortcutContext.WidgetWithChildrenShortcut

        # Ctrl+B for bold
        bold_shortcut = QAction(self)
        bold_shortcut.setShortcut(QKeySequence("Ctrl+B"))
        bold_shortcut.setShortcutContext(context)
        bold_shortcut.triggered.connect(self._toggle_bold)
        self.addAction(bold_shortcut)

        # Ctrl+I for italic
        italic_shortcut = QAction(self)
        italic_shortcut.setShortcut(QKeySequence("Ctrl+I"))
        italic_shortcut.setShortcutContext(context)
        italic_shortcut.triggered.connect(self._toggle_italic)
        self.addAction(italic_shortcut)

        # Ctrl+U for underline
        underline_shortcut = QAction(self)
        underline_shortcut.setShortcut(QKeySequence("Ctrl+U"))
        underline_shortcut.setShortcutContext(context)
        underline_shortcut.triggered.connect(self._toggle_underline)
        self.addAction(underline_shortcut)

        # Ctrl+Enter to send
        send_shortcut = QAction(self)
        send_shortcut.setShortcut(QKeySequence("Ctrl+Return"))
        send_shortcut.setShortcutContext(context)
        send_shortcut.triggered.connect(self._send)
        self.addAction(send_shortcut)

        # Escape to discard
        esc_shortcut = QAction(self)
        esc_shortcut.setShortcut(QKeySequence("Escape"))
        esc_shortcut.setShortcutContext(context)
        esc_shortcut.triggered.connect(self._discard)
        self.addAction(esc_shortcut)


    def _toggle_bold(self):
        """Toggle bold formatting."""
        if self._use_webengine:
            self.body_edit.toggleBold()
        else:
            fmt = QTextCharFormat()
            if self.body_edit.fontWeight() == QFont.Weight.Bold:
                fmt.setFontWeight(QFont.Weight.Normal)
            else:
                fmt.setFontWeight(QFont.Weight.Bold)
            self._merge_format(fmt)
            self._update_format_buttons()

    def _toggle_italic(self):
        """Toggle italic formatting."""
        if self._use_webengine:
            self.body_edit.toggleItalic()
        else:
            fmt = QTextCharFormat()
            fmt.setFontItalic(not self.body_edit.fontItalic())
            self._merge_format(fmt)
            self._update_format_buttons()

    def _toggle_underline(self):
        """Toggle underline formatting."""
        if self._use_webengine:
            self.body_edit.toggleUnderline()
        else:
            fmt = QTextCharFormat()
            fmt.setFontUnderline(not self.body_edit.fontUnderline())
            self._merge_format(fmt)
            self._update_format_buttons()

    def _change_font_size(self, size_str):
        """Change font size."""
        try:
            size = int(size_str)
            if self._use_webengine:
                self.body_edit.setFontSize(size)
            else:
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
        if self._use_webengine:
            self.body_edit.insertBullet()
        else:
            self.body_edit.insertPlainText("• ")

    def _paste(self):
        """Paste clipboard contents at cursor position."""
        # Get the currently focused widget
        focused = QApplication.focusWidget()
        if focused == self.body_edit:
            self.body_edit.paste()
        elif focused == self.to_edit:
            self.to_edit.paste()
        elif focused == self.cc_edit:
            self.cc_edit.paste()
        elif focused == self.subject_edit:
            self.subject_edit.paste()
        elif hasattr(focused, 'paste'):
            focused.paste()

    def _select_all(self):
        """Select all text in focused widget."""
        focused = QApplication.focusWidget()
        if hasattr(focused, 'selectAll'):
            focused.selectAll()

    def _attach(self):
        """Add attachments."""
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select files to attach", self._default_attach_dir
        )
        for filepath in files:
            try:
                with open(filepath, 'rb') as f:
                    content = f.read()
                name = os.path.basename(filepath)
                self._attachments.append({'name': name, 'content': content, 'path': filepath})
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to read {filepath}: {e}")

        self._update_attachments_display()

    def _format_file_size(self, size_bytes: int) -> str:
        """Format file size in human-readable form."""
        if size_bytes < 1024:
            return f"{size_bytes} B"
        elif size_bytes < 1024 * 1024:
            return f"{size_bytes / 1024:.1f} KB"
        else:
            return f"{size_bytes / (1024 * 1024):.1f} MB"

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

    def _clear_attachments_layout(self):
        """Clear all widgets from attachments layout."""
        while self._attachments_layout.count():
            child = self._attachments_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

    def _remove_attachment(self, index: int):
        """Remove an attachment by index."""
        if 0 <= index < len(self._attachments):
            del self._attachments[index]
            self._update_attachments_display()

    def _update_attachments_display(self):
        """Update attachments display with cards."""
        self._clear_attachments_layout()

        if not self._attachments:
            self._attachments_widget.hide()
            return

        self._attachments_widget.show()

        # Header: "X Attachment(s)"
        count = len(self._attachments)
        header = QLabel(f"{count} Attachment{'s' if count > 1 else ''}")
        header_color = "#999" if self._is_dark else "#666"
        header.setStyleSheet(f"font-weight: bold; color: {header_color};")
        self._attachments_layout.addWidget(header)

        # Attachment cards in a horizontal flow
        cards_widget = QWidget()
        cards_layout = QHBoxLayout(cards_widget)
        cards_layout.setContentsMargins(0, 0, 0, 0)
        cards_layout.setSpacing(10)

        for i, att in enumerate(self._attachments):
            card = self._create_attachment_card(i, att)
            cards_layout.addWidget(card)

        cards_layout.addStretch()
        self._attachments_layout.addWidget(cards_widget)

    def _create_attachment_card(self, index: int, attachment: dict) -> QWidget:
        """Create a card widget for an attachment."""
        card = QFrame()
        card.setFrameShape(QFrame.Shape.NoFrame)

        if self._is_dark:
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
        card.setObjectName("attachCard")

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
        content = attachment.get('content', b'')
        size = len(content) if isinstance(content, bytes) else len(content.encode('utf-8') if content else b'')
        size_label = QLabel(self._format_file_size(size))
        size_color = "#777" if self._is_dark else "#888"
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
                color: {'#888' if self._is_dark else '#666'};
            }}
            QToolButton:hover {{
                background-color: {'#4c5052' if self._is_dark else '#ddd'};
                border-radius: 4px;
            }}
            QToolButton::menu-indicator {{
                image: none;
            }}
        """
        menu_btn.setStyleSheet(btn_style)

        menu = QMenu(menu_btn)

        # For compose window, only delete makes sense
        delete_action = menu.addAction("🗑 Remove")
        delete_action.triggered.connect(lambda checked, idx=index: self._remove_attachment(idx))

        menu_btn.setMenu(menu)
        layout.addWidget(menu_btn)

        return card

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

            # Apply default paragraph format for new text
            self._set_default_paragraph_format()

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
        body_type = original_message.get('body_type', 'text')
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

        # Preserve HTML body or convert plain text to HTML
        if body_type == 'html':
            # Strip outer html/body tags but keep the content
            body_html = orig_body
            # Remove doctype, html, head, body tags but keep content
            body_html = re.sub(r'<!DOCTYPE[^>]*>', '', body_html, flags=re.IGNORECASE)
            body_html = re.sub(r'<html[^>]*>', '', body_html, flags=re.IGNORECASE)
            body_html = re.sub(r'</html>', '', body_html, flags=re.IGNORECASE)
            body_html = re.sub(r'<head[^>]*>.*?</head>', '', body_html, flags=re.IGNORECASE | re.DOTALL)
            body_html = re.sub(r'<body[^>]*>', '', body_html, flags=re.IGNORECASE)
            body_html = re.sub(r'</body>', '', body_html, flags=re.IGNORECASE)
            # Remove remote images only for QTextEdit - WebEngine can load them
            if not self._use_webengine:
                body_html = re.sub(r'<img[^>]*src\s*=\s*["\']https?://[^"\']*["\'][^>]*>', '', body_html, flags=re.IGNORECASE)
            # Always remove about:blank images (blocked images)
            body_html = re.sub(r'<img[^>]*src\s*=\s*["\']about:blank["\'][^>]*>', '', body_html, flags=re.IGNORECASE)
            quoted_body = body_html.strip()
        else:
            # Convert plain text to HTML
            plain_body = self._html_to_plain_text(orig_body)
            quoted_body = html_module.escape(plain_body).replace('\n', '<br>')

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
<td style="padding-left: 10px;">{quoted_body}</td>
</tr>
</table>
</body></html>'''

        self.body_edit.setHtml(html_content)

        # Move cursor to beginning and focus body
        cursor = self.body_edit.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.Start)
        self.body_edit.setTextCursor(cursor)

        # Apply default paragraph format for new text
        self._set_default_paragraph_format()

        # Use timer to ensure focus after widget is fully shown
        from PySide6.QtCore import QTimer
        QTimer.singleShot(50, self.body_edit.setFocus)

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
        body_type = original_message.get('body_type', 'text')

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

        # Preserve HTML body or convert plain text to HTML
        if body_type == 'html':
            # Strip outer html/body tags but keep the content
            body_html = body
            # Remove doctype, html, head, body tags but keep content
            body_html = re.sub(r'<!DOCTYPE[^>]*>', '', body_html, flags=re.IGNORECASE)
            body_html = re.sub(r'<html[^>]*>', '', body_html, flags=re.IGNORECASE)
            body_html = re.sub(r'</html>', '', body_html, flags=re.IGNORECASE)
            body_html = re.sub(r'<head[^>]*>.*?</head>', '', body_html, flags=re.IGNORECASE | re.DOTALL)
            body_html = re.sub(r'<body[^>]*>', '', body_html, flags=re.IGNORECASE)
            body_html = re.sub(r'</body>', '', body_html, flags=re.IGNORECASE)
            # Remove remote images only for QTextEdit - WebEngine can load them
            if not self._use_webengine:
                body_html = re.sub(r'<img[^>]*src\s*=\s*["\']https?://[^"\']*["\'][^>]*>', '', body_html, flags=re.IGNORECASE)
            # Always remove about:blank images (blocked images)
            body_html = re.sub(r'<img[^>]*src\s*=\s*["\']about:blank["\'][^>]*>', '', body_html, flags=re.IGNORECASE)
            forwarded_body = body_html.strip()
        else:
            # Convert plain text to HTML
            plain_body = self._html_to_plain_text(body)
            forwarded_body = html_module.escape(plain_body).replace('\n', '<br>')

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
<div>{forwarded_body}</div>
</td>
</tr>
</table>
</body></html>'''

        self.body_edit.setHtml(html_content)

        cursor = self.body_edit.textCursor()
        cursor.movePosition(QTextCursor.MoveOperation.Start)
        self.body_edit.setTextCursor(cursor)

        # Apply default paragraph format for new text
        self._set_default_paragraph_format()

        self.to_edit.setFocus()

    def _get_html_body(self):
        """Convert rich text to HTML with proper styling."""
        if self._use_webengine:
            html = self.body_edit.toHtmlSync()
        else:
            html = self.body_edit.toHtml()

        # Web-safe sans-serif font stack
        font_stack = "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif"

        # Replace Qt's font-family declarations with web-safe fonts
        # Qt outputs things like: font-family:'Sans Serif'; or font-family:"Sans Serif";
        html = re.sub(
            r"font-family:['\"]?[^;'\"]+['\"]?;",
            f"font-family: {font_stack};",
            html
        )

        # Add inline style to body tag for clients that inherit from body
        if '<body' in html:
            if '<body style="' in html:
                html = html.replace('<body style="', f'<body style="font-family: {font_stack}; ')
            else:
                html = html.replace('<body', f'<body style="font-family: {font_stack};"', 1)

        # Add inline margin to p tags
        html = re.sub(
            r'<p([^>]*)style="([^"]*)"',
            r'<p\1style="margin: 0.3em 0; \2"',
            html
        )
        # Handle p tags without style attribute
        html = re.sub(
            r'<p(?![^>]*style=)([^>]*)>',
            r'<p style="margin: 0.3em 0;"\1>',
            html
        )

        # Convert tab characters to non-breaking spaces (4 spaces per tab)
        html = html.replace('\t', '&nbsp;&nbsp;&nbsp;&nbsp;')

        return html

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
        if self._use_webengine:
            body_html = self.body_edit.toHtmlSync()
        else:
            body_html = self.body_edit.toHtml()
        return {
            'to': self.to_edit.text(),
            'cc': self.cc_edit.text(),
            'subject': self.subject_edit.text(),
            'body': body_html,
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
        """Set zoom level for the compose body (display only, doesn't affect sent email)."""
        self._zoom = zoom_factor
        if self._use_webengine:
            # WebEngine has native zoom that doesn't affect content
            self.body_edit.setZoomFactor(zoom_factor)
        else:
            # QTextEdit doesn't have zoom, so we scale the font for display
            # Note: This affects the sent email for QTextEdit fallback
            font = QFont("Sans Serif")
            font.setPointSize(int(self._font_size * zoom_factor))
            self.body_edit.setFont(font)
            self.body_edit.viewport().update()

    def _apply_theme(self):
        """Apply dark/light theme - now using system defaults."""
        # Clear any stylesheet - use system theme
        self.setStyleSheet("")
