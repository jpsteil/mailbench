"""Compose email view with WYSIWYG editing."""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, font


class ComposeView:
    def __init__(self, parent, db, sync_manager, account_id=None, reply_to=None, forward=None):
        self.db = db
        self.sync_manager = sync_manager
        self.account_id = account_id
        self.attachments = []

        self.window = tk.Toplevel(parent)
        self.window.title("New Message")
        self.window.geometry("750x600")
        self.window.transient(parent)

        # Check dark mode
        self.is_dark = db.get_setting("dark_mode", "0") == "1"

        self._create_ui()
        self._setup_tags()
        self._apply_theme()

        if reply_to:
            self._setup_reply(reply_to)
        elif forward:
            self._setup_forward(forward)

        # Center on parent
        self.window.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.window.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.window.winfo_height()) // 2
        self.window.geometry(f"+{x}+{y}")

    def _create_ui(self):
        # Main toolbar
        toolbar = ttk.Frame(self.window)
        toolbar.pack(fill=tk.X, padx=5, pady=5)

        ttk.Button(toolbar, text="Send", command=self._send).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Attach", command=self._attach).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="Discard", command=self._discard).pack(side=tk.LEFT, padx=2)

        # Header fields
        header_frame = ttk.Frame(self.window)
        header_frame.pack(fill=tk.X, padx=10, pady=5)

        # From (account selector)
        ttk.Label(header_frame, text="From:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.from_var = tk.StringVar()
        accounts = self.db.get_accounts()
        account_names = [f"{a['name']} <{a['email']}>" for a in accounts]
        self.from_combo = ttk.Combobox(header_frame, textvariable=self.from_var,
                                       values=account_names, width=50, state="readonly")
        self.from_combo.grid(row=0, column=1, sticky=tk.W, pady=2)
        if account_names:
            for i, a in enumerate(accounts):
                if a.get('is_default') or (self.account_id and a['id'] == self.account_id):
                    self.from_combo.current(i)
                    break
            else:
                self.from_combo.current(0)
        self._accounts = accounts

        # To
        ttk.Label(header_frame, text="To:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.to_var = tk.StringVar()
        self.to_entry = ttk.Entry(header_frame, textvariable=self.to_var, width=55)
        self.to_entry.grid(row=1, column=1, sticky=tk.W, pady=2)

        # CC
        ttk.Label(header_frame, text="CC:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.cc_var = tk.StringVar()
        self.cc_entry = ttk.Entry(header_frame, textvariable=self.cc_var, width=55)
        self.cc_entry.grid(row=2, column=1, sticky=tk.W, pady=2)

        # Subject
        ttk.Label(header_frame, text="Subject:").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.subject_var = tk.StringVar()
        self.subject_entry = ttk.Entry(header_frame, textvariable=self.subject_var, width=55)
        self.subject_entry.grid(row=3, column=1, sticky=tk.W, pady=2)

        # Attachments display
        self.attach_frame = ttk.Frame(self.window)
        self.attach_frame.pack(fill=tk.X, padx=10)
        self.attach_label = ttk.Label(self.attach_frame, text="")
        self.attach_label.pack(side=tk.LEFT)

        # Formatting toolbar
        format_toolbar = ttk.Frame(self.window)
        format_toolbar.pack(fill=tk.X, padx=10, pady=5)

        # Bold, Italic, Underline buttons (use tk.Button for relief/state changes)
        btn_bg = "#4a4a4a" if self.is_dark else "#e0e0e0"
        btn_fg = "#ffffff" if self.is_dark else "#000000"
        btn_active_bg = "#666666" if self.is_dark else "#a0a0ff"

        self.bold_btn = tk.Button(format_toolbar, text="B", width=3, font=("TkDefaultFont", 10, "bold"),
                                  bg=btn_bg, fg=btn_fg, activebackground=btn_active_bg,
                                  relief=tk.RAISED, command=self._toggle_bold)
        self.bold_btn.pack(side=tk.LEFT, padx=1)
        self.italic_btn = tk.Button(format_toolbar, text="I", width=3, font=("TkDefaultFont", 10, "italic"),
                                    bg=btn_bg, fg=btn_fg, activebackground=btn_active_bg,
                                    relief=tk.RAISED, command=self._toggle_italic)
        self.italic_btn.pack(side=tk.LEFT, padx=1)
        self.underline_btn = tk.Button(format_toolbar, text="U", width=3, font=("TkDefaultFont", 10, "underline"),
                                       bg=btn_bg, fg=btn_fg, activebackground=btn_active_bg,
                                       relief=tk.RAISED, command=self._toggle_underline)
        self.underline_btn.pack(side=tk.LEFT, padx=1)

        # Store button colors for toggling
        self._btn_bg = btn_bg
        self._btn_active_bg = "#0078d7" if self.is_dark else "#0066cc"  # Blue when active

        ttk.Separator(format_toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)

        # Font size selector
        ttk.Label(format_toolbar, text="Size:").pack(side=tk.LEFT, padx=2)
        self.font_size_var = tk.StringVar(value="12")
        self.size_combo = ttk.Combobox(format_toolbar, textvariable=self.font_size_var,
                                       values=["10", "12", "14", "16", "18", "20", "24", "28", "32"],
                                       width=4, state="readonly")
        self.size_combo.pack(side=tk.LEFT, padx=2)
        self.size_combo.bind("<<ComboboxSelected>>", self._change_font_size)

        ttk.Separator(format_toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)

        # Bullet list
        ttk.Button(format_toolbar, text="• List", command=self._insert_bullet).pack(side=tk.LEFT, padx=1)

        # Body (rich text)
        body_frame = ttk.Frame(self.window)
        body_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        self.body_text = tk.Text(body_frame, wrap=tk.WORD, font=("TkDefaultFont", 12), undo=True)
        body_scroll = ttk.Scrollbar(body_frame, orient=tk.VERTICAL, command=self.body_text.yview)
        self.body_text.configure(yscrollcommand=body_scroll.set)

        body_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.body_text.pack(fill=tk.BOTH, expand=True)

        # Keyboard shortcuts
        self.body_text.bind("<Control-b>", lambda e: self._toggle_bold())
        self.body_text.bind("<Control-B>", lambda e: self._toggle_bold())
        self.body_text.bind("<Control-i>", lambda e: self._toggle_italic())
        self.body_text.bind("<Control-I>", lambda e: self._toggle_italic())
        self.body_text.bind("<Control-u>", lambda e: self._toggle_underline())
        self.body_text.bind("<Control-U>", lambda e: self._toggle_underline())

        # Ctrl+Enter to send
        self.window.bind("<Control-Return>", lambda e: self._send())

        # Track active formatting for typing
        self._active_tags = set()
        self.body_text.bind("<KeyPress>", self._on_key_press)

        # Focus on To field
        self.to_entry.focus()

    def _setup_tags(self):
        """Configure text tags for formatting."""
        # Get base font
        base_font = font.Font(family="TkDefaultFont", size=12)

        # Bold
        bold_font = font.Font(family="TkDefaultFont", size=12, weight="bold")
        self.body_text.tag_configure("bold", font=bold_font)

        # Italic
        italic_font = font.Font(family="TkDefaultFont", size=12, slant="italic")
        self.body_text.tag_configure("italic", font=italic_font)

        # Underline
        self.body_text.tag_configure("underline", underline=True)

        # Bold+Italic
        bold_italic_font = font.Font(family="TkDefaultFont", size=12, weight="bold", slant="italic")
        self.body_text.tag_configure("bold_italic", font=bold_italic_font)

        # Font sizes
        for size in [10, 12, 14, 16, 18, 20, 24, 28, 32]:
            self.body_text.tag_configure(f"size_{size}", font=font.Font(family="TkDefaultFont", size=size))

    def _on_key_press(self, event):
        """Apply active formatting to newly typed characters."""
        # Only handle printable characters
        if event.char and event.char.isprintable() and len(event.char) == 1:
            # Get current position
            insert_pos = self.body_text.index(tk.INSERT)

            # Schedule tag application after the character is inserted
            self.body_text.after(1, lambda: self._apply_active_tags(insert_pos))

    def _apply_active_tags(self, pos):
        """Apply active tags to the character just typed."""
        try:
            end_pos = self.body_text.index(f"{pos}+1c")
            for tag in self._active_tags:
                self.body_text.tag_add(tag, pos, end_pos)
        except tk.TclError:
            pass

    def _toggle_bold(self):
        """Toggle bold on selected text or for future typing."""
        self._toggle_tag("bold")
        return "break"

    def _toggle_italic(self):
        """Toggle italic on selected text or for future typing."""
        self._toggle_tag("italic")
        return "break"

    def _toggle_underline(self):
        """Toggle underline on selected text or for future typing."""
        self._toggle_tag("underline")
        return "break"

    def _toggle_tag(self, tag_name):
        """Toggle a tag on the selected text or for future typing."""
        try:
            sel_start = self.body_text.index(tk.SEL_FIRST)
            sel_end = self.body_text.index(tk.SEL_LAST)

            # Check if tag is already applied
            current_tags = self.body_text.tag_names(sel_start)
            if tag_name in current_tags:
                self.body_text.tag_remove(tag_name, sel_start, sel_end)
            else:
                self.body_text.tag_add(tag_name, sel_start, sel_end)
        except tk.TclError:
            # No selection - toggle for future typing
            if tag_name in self._active_tags:
                self._active_tags.discard(tag_name)
            else:
                self._active_tags.add(tag_name)

        # Update button appearance
        self._update_format_buttons()

    def _update_format_buttons(self):
        """Update formatting button appearance based on active tags."""
        # Bold button
        if "bold" in self._active_tags:
            self.bold_btn.config(relief=tk.SUNKEN, bg=self._btn_active_bg)
        else:
            self.bold_btn.config(relief=tk.RAISED, bg=self._btn_bg)

        # Italic button
        if "italic" in self._active_tags:
            self.italic_btn.config(relief=tk.SUNKEN, bg=self._btn_active_bg)
        else:
            self.italic_btn.config(relief=tk.RAISED, bg=self._btn_bg)

        # Underline button
        if "underline" in self._active_tags:
            self.underline_btn.config(relief=tk.SUNKEN, bg=self._btn_active_bg)
        else:
            self.underline_btn.config(relief=tk.RAISED, bg=self._btn_bg)

    def _change_font_size(self, event=None):
        """Change font size of selected text or for future typing."""
        size = self.font_size_var.get()
        size_tag = f"size_{size}"

        # Remove other size tags from active set
        self._active_tags = {t for t in self._active_tags if not t.startswith("size_")}
        if size != "12":
            self._active_tags.add(size_tag)

        try:
            sel_start = self.body_text.index(tk.SEL_FIRST)
            sel_end = self.body_text.index(tk.SEL_LAST)

            # Remove any existing size tags
            for s in [10, 12, 14, 16, 18, 20, 24, 28, 32]:
                self.body_text.tag_remove(f"size_{s}", sel_start, sel_end)

            # Apply new size
            if size != "12":
                self.body_text.tag_add(size_tag, sel_start, sel_end)
        except tk.TclError:
            pass

    def _insert_bullet(self):
        """Insert a bullet point."""
        self.body_text.insert(tk.INSERT, "• ")

    def _setup_reply(self, original_message):
        """Setup for reply."""
        reply_all = original_message.get('reply_all', False)
        self.window.title("Reply All" if reply_all else "Reply")

        # Get sender info
        from_name = original_message.get('from_name', '')
        from_email = original_message.get('from_email', '')

        # Set To: field to original sender
        if from_name and from_email:
            self.to_var.set(f"{from_name} <{from_email}>")
        elif from_email:
            self.to_var.set(from_email)

        # For Reply All, add CC recipients
        if reply_all:
            # Get my email to exclude from CC
            my_email = ''
            if self._accounts:
                idx = self.from_combo.current()
                if idx >= 0 and idx < len(self._accounts):
                    my_email = self._accounts[idx].get('email', '').lower()

            # Combine original To and CC, excluding myself and the sender
            cc_list = []
            original_to = original_message.get('to', '')
            original_cc = original_message.get('cc', '')

            for addr in (original_to + ',' + original_cc).split(','):
                addr = addr.strip()
                if not addr:
                    continue
                # Extract email from "Name <email>" format
                addr_lower = addr.lower()
                if my_email and my_email in addr_lower:
                    continue
                if from_email and from_email.lower() in addr_lower:
                    continue
                cc_list.append(addr)

            if cc_list:
                self.cc_var.set(', '.join(cc_list))

        # Set Subject with Re: prefix
        subject = original_message.get('subject', '')
        if not subject.lower().startswith('re:'):
            subject = f"Re: {subject}"
        self.subject_var.set(subject)

        # Add quoted original message to body
        orig_body = original_message.get('body', '')
        date = original_message.get('date', '')

        # Format date if it's in Kerio format
        if date and len(date) >= 15:
            try:
                from datetime import datetime
                date_part = date[:8]
                time_part = date[9:15]
                dt = datetime.strptime(f"{date_part}{time_part}", "%Y%m%d%H%M%S")
                date = dt.strftime("%B %d, %Y at %I:%M %p")
            except Exception:
                pass

        # Build reply header
        sender_str = f"{from_name} <{from_email}>" if from_name else from_email
        quote_header = f"\n\nOn {date}, {sender_str} wrote:\n" if date else f"\n\n{sender_str} wrote:\n"

        # Insert quote header first
        self.body_text.insert(tk.END, quote_header)

        # Parse and insert HTML with formatting preserved
        self._insert_html_as_quoted(orig_body)

        # Focus on body for typing reply
        self.body_text.focus()
        self.body_text.mark_set(tk.INSERT, "1.0")

    def _insert_html_as_quoted(self, html_body):
        """Insert HTML content as quoted text, preserving basic formatting."""
        import re
        import html as html_module

        if not html_body:
            return

        # Clean problematic Unicode characters
        def clean_text(t):
            t = t.replace('\u00a0', ' ')  # Non-breaking space
            t = t.replace('\u200b', '')   # Zero-width space
            t = t.replace('\u200c', '')   # Zero-width non-joiner
            t = t.replace('\u200d', '')   # Zero-width joiner
            t = t.replace('\u2003', ' ')  # Em space
            t = t.replace('\u2002', ' ')  # En space
            t = t.replace('\u2009', ' ')  # Thin space
            t = t.replace('\ufeff', '')   # BOM
            t = t.replace('\r\n', '\n').replace('\r', '\n')
            return t

        # Convert HTML to segments with formatting info
        segments = []  # List of (text, tags_set)
        current_tags = set()

        # Normalize line breaks first
        text = html_body
        text = re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)
        text = re.sub(r'<p[^>]*>', '\n', text, flags=re.IGNORECASE)
        text = re.sub(r'</p>', '\n', text, flags=re.IGNORECASE)
        text = re.sub(r'<div[^>]*>', '\n', text, flags=re.IGNORECASE)
        text = re.sub(r'</div>', '', text, flags=re.IGNORECASE)

        # Parse HTML tags and text
        tag_pattern = re.compile(r'<(/?)(\w+)[^>]*>', re.IGNORECASE)
        pos = 0

        for match in tag_pattern.finditer(text):
            # Add text before this tag
            if match.start() > pos:
                chunk = text[pos:match.start()]
                chunk = html_module.unescape(chunk)
                chunk = clean_text(chunk)
                if chunk:
                    segments.append((chunk, frozenset(current_tags)))

            # Process the tag
            is_closing = match.group(1) == '/'
            tag_name = match.group(2).lower()

            if tag_name in ('b', 'strong'):
                if is_closing:
                    current_tags.discard('bold')
                else:
                    current_tags.add('bold')
            elif tag_name in ('i', 'em'):
                if is_closing:
                    current_tags.discard('italic')
                else:
                    current_tags.add('italic')
            elif tag_name == 'u':
                if is_closing:
                    current_tags.discard('underline')
                else:
                    current_tags.add('underline')

            pos = match.end()

        # Add remaining text after last tag
        if pos < len(text):
            chunk = text[pos:]
            chunk = html_module.unescape(chunk)
            chunk = clean_text(chunk)
            if chunk:
                segments.append((chunk, frozenset(current_tags)))

        # Now insert segments with formatting, adding > prefix to lines
        # First, combine into lines so we can add > prefix
        full_text = ''.join(seg[0] for seg in segments)
        full_text = re.sub(r'\n{3,}', '\n\n', full_text).strip()

        # Build position map: for each char in full_text, what tags apply
        char_tags = []
        for seg_text, seg_tags in segments:
            for _ in seg_text:
                char_tags.append(seg_tags)

        # Insert line by line with > prefix
        lines = full_text.split('\n')
        char_idx = 0

        for i, line in enumerate(lines):
            # Insert > prefix (no formatting)
            self.body_text.insert(tk.END, '> ')

            # Insert each character with its formatting
            for ch in line:
                if char_idx < len(char_tags):
                    tags = char_tags[char_idx]
                    start_idx = self.body_text.index(tk.END + "-1c")
                    self.body_text.insert(tk.END, ch)
                    if tags:
                        end_idx = self.body_text.index(tk.END + "-1c")
                        for tag in tags:
                            self.body_text.tag_add(tag, start_idx, end_idx)
                else:
                    self.body_text.insert(tk.END, ch)
                char_idx += 1

            # Newline (skip the \n in char_tags)
            if i < len(lines) - 1:
                self.body_text.insert(tk.END, '\n')
                char_idx += 1  # Skip the \n character in char_tags

    def _setup_forward(self, original_message):
        """Setup for forward."""
        self.window.title("Forward")
        # Pre-fill body with original message
        pass

    def _attach(self):
        """Add attachment."""
        files = filedialog.askopenfilenames(parent=self.window, title="Select files to attach")
        for filepath in files:
            try:
                with open(filepath, 'rb') as f:
                    content = f.read()
                import os
                name = os.path.basename(filepath)
                self.attachments.append({'name': name, 'content': content})
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read {filepath}: {e}")

        if self.attachments:
            names = [a['name'] for a in self.attachments]
            self.attach_label.config(text=f"Attachments: {', '.join(names)}")

    def _get_html_body(self):
        """Convert the rich text content to HTML."""
        # Get all text with tags
        html_parts = []
        index = "1.0"

        while True:
            # Get tags at current position
            tags = self.body_text.tag_names(index)

            # Find next tag change or end
            next_index = self.body_text.index(f"{index}+1c")
            if self.body_text.compare(next_index, ">=", tk.END):
                break

            char = self.body_text.get(index, next_index)
            if char == "\n":
                html_parts.append("<br>\n")
            else:
                # Apply formatting
                formatted = char
                if "bold" in tags:
                    formatted = f"<b>{formatted}</b>"
                if "italic" in tags:
                    formatted = f"<i>{formatted}</i>"
                if "underline" in tags:
                    formatted = f"<u>{formatted}</u>"

                # Check for size tags
                for tag in tags:
                    if tag.startswith("size_"):
                        size = tag.replace("size_", "")
                        formatted = f'<span style="font-size:{size}px">{formatted}</span>'

                html_parts.append(formatted)

            index = next_index

        # Merge consecutive identical formatting
        body_html = "".join(html_parts)

        # Wrap in basic HTML structure
        return f"<html><body>{body_html}</body></html>"

    def _send(self):
        """Send the message."""
        to = self.to_var.get().strip()
        if not to:
            messagebox.showerror("Error", "To field is required")
            return

        to_list = [addr.strip() for addr in to.split(',') if addr.strip()]
        cc = self.cc_var.get().strip()
        cc_list = [addr.strip() for addr in cc.split(',') if addr.strip()] if cc else []

        subject = self.subject_var.get().strip()
        body = self._get_html_body()

        from_idx = self.from_combo.current()
        if from_idx < 0 or from_idx >= len(self._accounts):
            messagebox.showerror("Error", "Please select a From account")
            return

        account_id = self._accounts[from_idx]['id']

        self.sync_manager.send_message(
            account_id=account_id,
            to=to_list,
            subject=subject,
            body=body,
            cc=cc_list,
            callback=self._on_sent
        )

        self.window.title("Sending...")

    def _on_sent(self, success, message):
        """Called when send completes."""
        if success:
            self.window.destroy()
        else:
            messagebox.showerror("Error", f"Failed to send: {message}")
            self.window.title("New Message")

    def _discard(self):
        """Discard the message."""
        body = self.body_text.get("1.0", tk.END).strip()
        to = self.to_var.get().strip()
        subject = self.subject_var.get().strip()

        if body or to or subject:
            if not messagebox.askyesno("Discard", "Discard this message?"):
                return

        self.window.destroy()

    def _apply_theme(self):
        """Apply dark/light theme to compose window."""
        if self.is_dark:
            bg = "#2b2b2b"
            fg = "#a9b7c6"
            text_bg = "#313335"
            select_bg = "#214283"
        else:
            bg = "#d9d9d9"
            fg = "#000000"
            text_bg = "#ffffff"
            select_bg = "#4a6984"

        self.window.configure(bg=bg)

        self.body_text.configure(
            bg=text_bg, fg=fg,
            insertbackground=fg,
            selectbackground=select_bg,
            selectforeground=fg
        )

        # Style ttk widgets
        style = ttk.Style()
        style.configure('Compose.TEntry', fieldbackground=text_bg, foreground=fg)
        style.configure('Compose.TCombobox', fieldbackground=text_bg, foreground=fg)
        style.map('Compose.TCombobox',
                  fieldbackground=[('readonly', text_bg), ('disabled', bg)],
                  foreground=[('readonly', fg), ('disabled', fg)],
                  selectbackground=[('readonly', select_bg)],
                  selectforeground=[('readonly', fg)])

        self.to_entry.configure(style='Compose.TEntry')
        self.cc_entry.configure(style='Compose.TEntry')
        self.subject_entry.configure(style='Compose.TEntry')
        self.from_combo.configure(style='Compose.TCombobox')
        self.size_combo.configure(style='Compose.TCombobox')

        self.window.option_add('*TCombobox*Listbox.background', text_bg)
        self.window.option_add('*TCombobox*Listbox.foreground', fg)
        self.window.option_add('*TCombobox*Listbox.selectBackground', select_bg)
        self.window.option_add('*TCombobox*Listbox.selectForeground', fg)
