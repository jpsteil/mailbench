"""Main application window."""

import tkinter as tk
from tkinter import ttk, messagebox

# Font Awesome icons (requires FontAwesome font)
FA_PAPERCLIP = "\uf0c6"  # Paperclip icon
FA_FLAG = "\uf024"  # Flag icon
# Unicode fallback icons
UNI_FLAG = "\u2691"  # Black flag (widely available)
UNI_REPLY = "\u21b5"  # Downwards arrow with corner leftwards (reply)
UNI_FORWARD = "\u21b7"  # Clockwise top semicircle arrow (forward)

# Try to import HTML renderer
try:
    from tkinterweb import HtmlFrame
    HAS_HTML_VIEW = True
except ImportError:
    HAS_HTML_VIEW = False

from mailbench.database import Database
from mailbench.kerio_client import KerioConnectionPool, SyncManager, KerioConfig
from mailbench.version import __version__


class MailbenchApp:
    def __init__(self):
        self.root = tk.Tk(className="mailbench")
        self.root.title(f"Mailbench v{__version__}")

        # Hide window during setup
        self.root.withdraw()

        # Set window icon
        self._set_window_icon()

        # Initialize database and connections
        self.db = Database()
        self.kerio_pool = KerioConnectionPool()
        self.sync_manager = SyncManager(self.kerio_pool, self.db, self.root)

        # Track connected accounts
        self.connected_accounts = set()

        # UI setup
        self._restore_geometry()
        self._create_menu()
        self._create_toolbar()
        self._create_main_layout()
        self._create_statusbar()

        # Apply theme and font size
        self._apply_theme()
        self._apply_font_size()

        # Re-attach menu after theme
        self.root.config(menu=self.menubar)

        # Handle window close
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        # Delayed initialization
        self.root.after(50, self._ensure_visible_on_screen)
        self.root.after(100, self._restore_session)
        self.root.after(500, self._check_for_updates)
        self.root.after(2000, self._build_email_cache_background)  # Build cache after UI is ready

    def _set_window_icon(self):
        """Set the window icon from the bundled PNG."""
        try:
            from pathlib import Path
            icon_path = Path(__file__).parent / "resources" / "mailbench.png"
            if icon_path.exists():
                icon = tk.PhotoImage(file=str(icon_path))
                self.root.iconphoto(True, icon)
                self._icon = icon
        except Exception:
            pass

    def _create_menu(self):
        self.menubar = tk.Menu(self.root)
        self.root.config(menu=self.menubar)

        # File menu
        file_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New Message", command=self._new_message, accelerator="Ctrl+N")
        file_menu.add_command(label="Check Mail", command=self._check_mail, accelerator="F5")
        file_menu.add_separator()
        self.dark_mode_var = tk.BooleanVar(value=self.db.get_setting("dark_mode", "0") == "1")
        file_menu.add_checkbutton(label="Dark Mode", variable=self.dark_mode_var,
                                  command=self._toggle_dark_mode)
        file_menu.add_command(label="Settings...", command=self._show_settings)
        file_menu.add_separator()
        file_menu.add_command(label="Reset Layout", command=self._reset_layout)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self._on_close)

        # Accounts menu
        accounts_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Accounts", menu=accounts_menu)
        accounts_menu.add_command(label="Add Account...", command=self._add_account)
        accounts_menu.add_command(label="Manage Accounts...", command=self._manage_accounts)

        # View menu
        view_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="View", menu=view_menu)
        view_menu.add_command(label="Mail", command=lambda: self._show_view("mail"))
        view_menu.add_command(label="Calendar", command=lambda: self._show_view("calendar"), state="disabled")
        view_menu.add_command(label="Contacts", command=lambda: self._show_view("contacts"), state="disabled")

        # Load font size setting (default 12)
        self.font_size = int(self.db.get_setting("font_size", "12"))

        # Help menu
        help_menu = tk.Menu(self.menubar, tearoff=0)
        self.menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About", command=self._show_about)

        # Keyboard bindings
        self.root.bind("<Control-n>", lambda e: self._new_message())
        self.root.bind("<Control-N>", lambda e: self._new_message())
        self.root.bind("<F5>", lambda e: self._check_mail())

    def _create_toolbar(self):
        """Create the main toolbar."""
        self.toolbar = ttk.Frame(self.root)
        self.toolbar.pack(fill=tk.X, padx=5, pady=2)

        ttk.Button(self.toolbar, text="New", command=self._new_message).pack(side=tk.LEFT, padx=2)
        self.reply_btn = ttk.Button(self.toolbar, text="Reply", command=self._reply)
        self.reply_btn.pack(side=tk.LEFT, padx=2)
        self.reply_all_btn = ttk.Button(self.toolbar, text="Reply All", command=self._reply_all)
        self.reply_all_btn.pack(side=tk.LEFT, padx=2)
        ttk.Button(self.toolbar, text="Forward", command=self._forward).pack(side=tk.LEFT, padx=2)

        ttk.Separator(self.toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)

        ttk.Button(self.toolbar, text="Delete", command=self._delete_messages).pack(side=tk.LEFT, padx=2)
        ttk.Button(self.toolbar, text="Refresh", command=self._check_mail).pack(side=tk.LEFT, padx=2)

    def _create_main_layout(self):
        """Create the 3-pane layout."""
        # Main container for panes
        self.main_paned = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        self.main_paned.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Left pane - Folder tree (accounts and folders)
        self.folder_frame = ttk.Frame(self.main_paned)
        self.main_paned.add(self.folder_frame, weight=0)

        ttk.Label(self.folder_frame, text="Accounts", font=("TkDefaultFont", 9, "bold")).pack(
            anchor=tk.W, padx=2, pady=(2, 0))

        # Folder tree
        self.folder_tree = ttk.Treeview(self.folder_frame, show="tree", selectmode="browse")
        folder_scroll = ttk.Scrollbar(self.folder_frame, orient=tk.VERTICAL, command=self.folder_tree.yview)
        self.folder_tree.configure(yscrollcommand=folder_scroll.set)
        folder_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.folder_tree.pack(fill=tk.BOTH, expand=True)

        self.folder_tree.bind("<<TreeviewSelect>>", self._on_folder_select)
        self.folder_tree.bind("<Double-1>", self._on_folder_double_click)

        # Right paned window (message list | preview)
        self.content_paned = ttk.PanedWindow(self.main_paned, orient=tk.HORIZONTAL)
        self.main_paned.add(self.content_paned, weight=1)

        # Middle pane - Message list
        self.message_frame = ttk.Frame(self.content_paned)
        self.content_paned.add(self.message_frame, weight=1)

        # Search box at top of message list
        search_frame = ttk.Frame(self.message_frame)
        search_frame.pack(fill=tk.X, pady=(0, 2))
        ttk.Label(search_frame, text="Filter:").pack(side=tk.LEFT, padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        self.search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.search_var.trace_add("write", lambda *args: self._filter_messages())
        self.search_entry.bind("<Escape>", lambda e: self.search_var.set(""))
        self.search_entry.bind("<Down>", self._filter_to_first_message)

        # Message list columns
        columns = ("flag", "status", "attach", "from", "subject", "date")
        self.message_list = ttk.Treeview(self.message_frame, columns=columns, show="headings",
                                         selectmode="extended")

        # Check if FontAwesome is available for icons
        from tkinter import font as tkfont
        available_fonts = tkfont.families()
        self._has_fontawesome = 'FontAwesome' in available_fonts

        self.message_list.heading("flag", text="", anchor=tk.CENTER)
        self.message_list.heading("status", text="", anchor=tk.CENTER)
        self.message_list.heading("attach", text="", anchor=tk.CENTER)
        self.message_list.heading("from", text="From", anchor=tk.W)
        self.message_list.heading("subject", text="Subject", anchor=tk.W)
        self.message_list.heading("date", text="Date", anchor=tk.W)

        self.message_list.column("flag", width=24, minwidth=24, stretch=False)
        self.message_list.column("status", width=24, minwidth=24, stretch=False)
        self.message_list.column("attach", width=24, minwidth=24, stretch=False)
        self.message_list.column("from", width=150, minwidth=100, stretch=True)
        self.message_list.column("subject", width=300, minwidth=150, stretch=True)
        self.message_list.column("date", width=76, minwidth=76, stretch=False)

        msg_scroll = ttk.Scrollbar(self.message_frame, orient=tk.VERTICAL, command=self.message_list.yview)
        self.message_list.configure(yscrollcommand=self._on_message_list_scroll)
        self._msg_scrollbar = msg_scroll

        msg_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.message_list.pack(fill=tk.BOTH, expand=True)

        self.message_list.bind("<<TreeviewSelect>>", self._on_message_select)
        self.message_list.bind("<Double-1>", self._on_message_double_click)
        self.message_list.bind("<Button-1>", self._on_message_click)
        # Enable shift+arrow for multi-select
        self.message_list.bind("<Shift-Up>", self._on_shift_arrow_select)
        self.message_list.bind("<Shift-Down>", self._on_shift_arrow_select)
        # Delete key to delete message
        self.message_list.bind("<Delete>", self._on_delete_key)
        # Right-click context menu
        self.message_list.bind("<Button-3>", self._show_message_context_menu)

        # Create context menu
        self._message_context_menu = tk.Menu(self.root, tearoff=0)

        # Pagination tracking
        self._messages_loaded = 0
        self._messages_loading_more = False

        # Right pane - Preview
        self.preview_frame = ttk.Frame(self.content_paned)
        self.content_paned.add(self.preview_frame, weight=1)

        # Preview header
        self.preview_header = ttk.Frame(self.preview_frame)
        self.preview_header.pack(fill=tk.X, padx=5, pady=5)

        self.preview_from = ttk.Label(self.preview_header, text="", font=("TkDefaultFont", 9, "bold"))
        self.preview_from.pack(anchor=tk.W)
        self.preview_to = ttk.Label(self.preview_header, text="")
        self.preview_to.pack(anchor=tk.W)
        self.preview_subject = ttk.Label(self.preview_header, text="", font=("TkDefaultFont", 10, "bold"))
        self.preview_subject.pack(anchor=tk.W, pady=(5, 0))
        self.preview_date = ttk.Label(self.preview_header, text="", foreground="gray")
        self.preview_date.pack(anchor=tk.W)

        # Attachments frame
        self.preview_attach_frame = ttk.Frame(self.preview_header)
        self.preview_attach_frame.pack(anchor=tk.W, fill=tk.X, pady=(5, 0))

        ttk.Separator(self.preview_frame, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=5)

        # Preview body
        preview_body_frame = ttk.Frame(self.preview_frame)
        preview_body_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        # Load preview zoom level (100 = 100%, default)
        self._preview_zoom = int(self.db.get_setting("preview_zoom", "100"))

        self._preview_body_frame = preview_body_frame

        if HAS_HTML_VIEW:
            self.preview_html = HtmlFrame(preview_body_frame, messages_enabled=False)
            self.preview_html.pack(fill=tk.BOTH, expand=True)
            self.preview_text = None
            # Initialize with theme-appropriate background
            self.root.after(100, self._init_preview_background)
        else:
            self.preview_html = None
            self.preview_text = tk.Text(preview_body_frame, wrap=tk.WORD, state=tk.DISABLED,
                                        font=("TkDefaultFont", 12))
            preview_scroll = ttk.Scrollbar(preview_body_frame, orient=tk.VERTICAL,
                                           command=self.preview_text.yview)
            self.preview_text.configure(yscrollcommand=preview_scroll.set)
            preview_scroll.pack(side=tk.RIGHT, fill=tk.Y)
            self.preview_text.pack(fill=tk.BOTH, expand=True)
            # Enable Ctrl+A select all in preview
            self.preview_text.bind("<Control-a>", lambda e: self._preview_select_all())
            self.preview_text.bind("<Control-A>", lambda e: self._preview_select_all())

        # Ctrl+scroll to zoom - bind globally and check if over preview
        self.root.bind("<Control-Button-4>", self._on_ctrl_scroll_up)
        self.root.bind("<Control-Button-5>", self._on_ctrl_scroll_down)
        self.root.bind("<Control-MouseWheel>", self._on_ctrl_scroll_wheel)

        # Create compose frame (hidden initially, replaces preview when composing)
        self._create_compose_frame()

        # Populate folder tree
        self._refresh_folder_tree()

    def _create_compose_frame(self):
        """Create the inline compose frame (hidden by default)."""
        from tkinter import font

        self.compose_frame = ttk.Frame(self.content_paned)
        self._compose_visible = False

        # Compose toolbar
        compose_toolbar = ttk.Frame(self.compose_frame)
        compose_toolbar.pack(fill=tk.X, padx=5, pady=5)

        ttk.Button(compose_toolbar, text="Send", command=self._send_compose).pack(side=tk.LEFT, padx=2)
        ttk.Button(compose_toolbar, text="Attach", command=self._attach_compose).pack(side=tk.LEFT, padx=2)
        ttk.Button(compose_toolbar, text="Discard", command=self._discard_compose).pack(side=tk.LEFT, padx=2)

        # Header fields
        header_frame = ttk.Frame(self.compose_frame)
        header_frame.pack(fill=tk.X, padx=10, pady=5)

        # From
        ttk.Label(header_frame, text="From:").grid(row=0, column=0, sticky=tk.W, pady=2)
        self.compose_from_var = tk.StringVar()
        accounts = self.db.get_accounts()
        account_names = [f"{a['name']} <{a['email']}>" for a in accounts]
        self.compose_from_combo = ttk.Combobox(header_frame, textvariable=self.compose_from_var,
                                                values=account_names, width=50, state="readonly")
        self.compose_from_combo.grid(row=0, column=1, sticky=tk.W, pady=2)
        if account_names:
            self.compose_from_combo.current(0)
        self._compose_accounts = accounts

        # To
        ttk.Label(header_frame, text="To:").grid(row=1, column=0, sticky=tk.W, pady=2)
        self.compose_to_var = tk.StringVar()
        self.compose_to_entry = ttk.Entry(header_frame, textvariable=self.compose_to_var, width=55)
        self.compose_to_entry.grid(row=1, column=1, sticky=tk.W, pady=2)
        self.compose_to_entry.bind("<Control-Return>", lambda e: self._send_compose())
        self.compose_to_entry.bind("<Escape>", lambda e: self._hide_autocomplete() or self._discard_compose())
        self.compose_to_entry.bind("<Control-a>", lambda e: (e.widget.select_range(0, tk.END), "break")[1])
        self.compose_to_entry.bind("<Control-A>", lambda e: (e.widget.select_range(0, tk.END), "break")[1])
        self.compose_to_entry.bind("<KeyRelease>", lambda e: self._on_address_key(e, self.compose_to_entry))
        self.compose_to_entry.bind("<Down>", lambda e: self._autocomplete_navigate(1))
        self.compose_to_entry.bind("<Up>", lambda e: self._autocomplete_navigate(-1))
        self.compose_to_entry.bind("<Return>", lambda e: self._autocomplete_select() or "break")
        self.compose_to_entry.bind("<Tab>", lambda e: self._autocomplete_select())

        # CC
        ttk.Label(header_frame, text="CC:").grid(row=2, column=0, sticky=tk.W, pady=2)
        self.compose_cc_var = tk.StringVar()
        self.compose_cc_entry = ttk.Entry(header_frame, textvariable=self.compose_cc_var, width=55)
        self.compose_cc_entry.grid(row=2, column=1, sticky=tk.W, pady=2)
        self.compose_cc_entry.bind("<Control-Return>", lambda e: self._send_compose())
        self.compose_cc_entry.bind("<Escape>", lambda e: self._hide_autocomplete() or self._discard_compose())
        self.compose_cc_entry.bind("<Control-a>", lambda e: (e.widget.select_range(0, tk.END), "break")[1])
        self.compose_cc_entry.bind("<Control-A>", lambda e: (e.widget.select_range(0, tk.END), "break")[1])
        self.compose_cc_entry.bind("<KeyRelease>", lambda e: self._on_address_key(e, self.compose_cc_entry))
        self.compose_cc_entry.bind("<Down>", lambda e: self._autocomplete_navigate(1))
        self.compose_cc_entry.bind("<Up>", lambda e: self._autocomplete_navigate(-1))
        self.compose_cc_entry.bind("<Return>", lambda e: self._autocomplete_select() or "break")
        self.compose_cc_entry.bind("<Tab>", lambda e: self._autocomplete_select())

        # Address book cache and autocomplete
        self._address_book = []  # List of {name, email, type} dicts
        self._autocomplete_popup = None
        self._autocomplete_listbox = None
        self._autocomplete_entry = None

        # Subject
        ttk.Label(header_frame, text="Subject:").grid(row=3, column=0, sticky=tk.W, pady=2)
        self.compose_subject_var = tk.StringVar()
        self.compose_subject_entry = ttk.Entry(header_frame, textvariable=self.compose_subject_var, width=55)
        self.compose_subject_entry.grid(row=3, column=1, sticky=tk.W, pady=2)
        self.compose_subject_entry.bind("<Control-Return>", lambda e: self._send_compose())
        self.compose_subject_entry.bind("<Escape>", lambda e: self._discard_compose())
        self.compose_subject_entry.bind("<Control-a>", lambda e: (e.widget.select_range(0, tk.END), "break")[1])
        self.compose_subject_entry.bind("<Control-A>", lambda e: (e.widget.select_range(0, tk.END), "break")[1])

        # Attachments display
        self.compose_attach_frame = ttk.Frame(self.compose_frame)
        self.compose_attach_frame.pack(fill=tk.X, padx=10)
        ttk.Label(self.compose_attach_frame, text="Attachments:").pack(side=tk.LEFT)
        ttk.Button(self.compose_attach_frame, text="+ Add", width=6,
                   command=self._compose_add_attachment_dialog).pack(side=tk.LEFT, padx=5)
        # Container for attachment buttons
        self.compose_attach_list_frame = ttk.Frame(self.compose_attach_frame)
        self.compose_attach_list_frame.pack(side=tk.LEFT, fill=tk.X)

        # Formatting toolbar
        format_toolbar = ttk.Frame(self.compose_frame)
        format_toolbar.pack(fill=tk.X, padx=10, pady=5)

        is_dark = self.dark_mode_var.get()
        btn_bg = "#4a4a4a" if is_dark else "#e0e0e0"
        btn_fg = "#ffffff" if is_dark else "#000000"
        self._compose_btn_bg = btn_bg
        self._compose_btn_active_bg = "#0078d7" if is_dark else "#0066cc"

        self.compose_bold_btn = tk.Button(format_toolbar, text="B", width=3,
                                          font=("TkDefaultFont", 10, "bold"),
                                          bg=btn_bg, fg=btn_fg, relief=tk.RAISED,
                                          command=self._compose_toggle_bold)
        self.compose_bold_btn.pack(side=tk.LEFT, padx=1)

        self.compose_italic_btn = tk.Button(format_toolbar, text="I", width=3,
                                            font=("TkDefaultFont", 10, "italic"),
                                            bg=btn_bg, fg=btn_fg, relief=tk.RAISED,
                                            command=self._compose_toggle_italic)
        self.compose_italic_btn.pack(side=tk.LEFT, padx=1)

        self.compose_underline_btn = tk.Button(format_toolbar, text="U", width=3,
                                               font=("TkDefaultFont", 10, "underline"),
                                               bg=btn_bg, fg=btn_fg, relief=tk.RAISED,
                                               command=self._compose_toggle_underline)
        self.compose_underline_btn.pack(side=tk.LEFT, padx=1)

        ttk.Separator(format_toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)

        ttk.Label(format_toolbar, text="Size:").pack(side=tk.LEFT, padx=2)
        self.compose_font_size_var = tk.StringVar(value="12")
        self.compose_size_combo = ttk.Combobox(format_toolbar, textvariable=self.compose_font_size_var,
                                               values=["10", "12", "14", "16", "18", "20", "24", "28", "32"],
                                               width=4, state="readonly")
        self.compose_size_combo.pack(side=tk.LEFT, padx=2)
        self.compose_size_combo.bind("<<ComboboxSelected>>", self._compose_change_font_size)

        ttk.Separator(format_toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=5)
        ttk.Button(format_toolbar, text="• List", command=self._compose_insert_bullet).pack(side=tk.LEFT, padx=1)

        # Body text
        body_frame = ttk.Frame(self.compose_frame)
        body_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        compose_font_size = getattr(self, 'font_size', 12)
        zoom_factor = getattr(self, '_preview_zoom', 100) / 100
        zoomed_compose_size = int(compose_font_size * zoom_factor)
        self.compose_body = tk.Text(body_frame, wrap=tk.WORD, font=("TkDefaultFont", zoomed_compose_size), undo=True)
        body_scroll = ttk.Scrollbar(body_frame, orient=tk.VERTICAL, command=self.compose_body.yview)
        self.compose_body.configure(yscrollcommand=body_scroll.set)
        body_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.compose_body.pack(fill=tk.BOTH, expand=True)

        # Setup text tags for formatting
        self._setup_compose_tags()

        # Keyboard shortcuts
        self.compose_body.bind("<Control-b>", lambda e: self._compose_toggle_bold())
        self.compose_body.bind("<Control-B>", lambda e: self._compose_toggle_bold())
        self.compose_body.bind("<Control-i>", lambda e: self._compose_toggle_italic())
        self.compose_body.bind("<Control-I>", lambda e: self._compose_toggle_italic())
        self.compose_body.bind("<Control-u>", lambda e: self._compose_toggle_underline())
        self.compose_body.bind("<Control-U>", lambda e: self._compose_toggle_underline())
        self.compose_body.bind("<Control-Return>", lambda e: (self._send_compose(), "break")[1])
        self.compose_body.bind("<Escape>", lambda e: self._discard_compose())
        # Standard edit shortcuts
        self.compose_body.bind("<Control-a>", lambda e: (self.compose_body.tag_add(tk.SEL, "1.0", tk.END), "break")[1])
        self.compose_body.bind("<Control-A>", lambda e: (self.compose_body.tag_add(tk.SEL, "1.0", tk.END), "break")[1])

        # Track active formatting
        self._compose_active_tags = set()
        self.compose_body.bind("<KeyPress>", self._compose_on_key_press)

        # Attachments list
        self._compose_attachments = []

    def _setup_compose_tags(self):
        """Setup text formatting tags for compose body."""
        from tkinter import font as tkfont
        base_size = getattr(self, 'font_size', 12)
        zoom_factor = getattr(self, '_preview_zoom', 100) / 100
        zoomed_base = int(base_size * zoom_factor)

        base_font = tkfont.Font(family="TkDefaultFont", size=zoomed_base)
        bold_font = tkfont.Font(family="TkDefaultFont", size=zoomed_base, weight="bold")
        italic_font = tkfont.Font(family="TkDefaultFont", size=zoomed_base, slant="italic")

        self.compose_body.tag_configure("bold", font=bold_font)
        self.compose_body.tag_configure("italic", font=italic_font)
        self.compose_body.tag_configure("underline", underline=True)

        # Quoted text styling - gray color with left margin to indicate quoted content
        self.compose_body.tag_configure("quoted", foreground="#666666", lmargin1=10, lmargin2=10)

        # Size tags - apply zoom for display, but tag name preserves actual size for email
        for size in [10, 12, 14, 16, 18, 20, 24, 28, 32]:
            zoomed_size = int(size * zoom_factor)
            size_font = tkfont.Font(family="TkDefaultFont", size=zoomed_size)
            self.compose_body.tag_configure(f"size_{size}", font=size_font)

    def _show_compose(self, title="New Message", reply_data=None, forward_data=None, attachments=None):
        """Show compose frame, hiding preview."""
        if self._compose_visible:
            return

        # Remove preview from paned window and add compose
        self.content_paned.forget(self.preview_frame)
        self.content_paned.add(self.compose_frame, weight=1)
        self._compose_visible = True

        # Reset compose fields
        self.compose_to_var.set("")
        self.compose_cc_var.set("")
        self.compose_subject_var.set("")
        self.compose_body.delete("1.0", tk.END)
        self._compose_attachments = attachments if attachments else []
        self._refresh_compose_attachments()
        self._compose_active_tags = set()
        self._update_compose_format_buttons()

        # Track original message for reply/forward marking
        self._compose_original_id = None
        self._compose_mode = "new"  # "new", "reply", "forward"

        # Update account list
        accounts = self.db.get_accounts()
        account_names = [f"{a['name']} <{a['email']}>" for a in accounts]
        self.compose_from_combo['values'] = account_names
        self._compose_accounts = accounts

        # Select current account
        if hasattr(self, '_current_account_id'):
            for i, a in enumerate(accounts):
                if a['id'] == self._current_account_id:
                    self.compose_from_combo.current(i)
                    break

        # Handle reply or forward data
        if reply_data:
            self._compose_original_id = reply_data.get('original_id')
            self._compose_mode = "reply"
            self._setup_compose_reply(reply_data)
            self.statusbar.config(text=title)
            self.compose_body.focus()
            self.compose_body.mark_set(tk.INSERT, "1.0")
        elif forward_data:
            self._compose_original_id = forward_data.get('original_id')
            self._compose_mode = "forward"
            self._setup_compose_forward(forward_data)
            self.statusbar.config(text=title)
            self.compose_to_entry.focus()
        else:
            self.statusbar.config(text="Composing new message")
            self.compose_to_entry.focus()

    def _hide_compose(self):
        """Hide compose frame, showing preview."""
        if not self._compose_visible:
            return

        # Remove compose from paned window and add preview
        self.content_paned.forget(self.compose_frame)
        self.content_paned.add(self.preview_frame, weight=1)
        self._compose_visible = False
        self._hide_autocomplete()
        self.statusbar.config(text="Ready")

    def _load_address_book(self, account_id):
        """Load contacts and users for autocomplete."""
        self._address_book = []

        def on_users(success, error, users):
            print(f"[AddressBook] Users loaded: success={success}, count={len(users) if users else 0}, error={error}")
            if success and users:
                for u in users:
                    u['priority'] = 0  # Highest priority for internal users
                self._address_book.extend(users)
                print(f"[AddressBook] Total entries now: {len(self._address_book)}")

        def on_contacts(success, error, contacts):
            print(f"[AddressBook] Contacts loaded: success={success}, count={len(contacts) if contacts else 0}, error={error}")
            if success and contacts:
                for c in contacts:
                    c['priority'] = 1 if c.get('type') == 'gal' else 2
                self._address_book.extend(contacts)
                print(f"[AddressBook] Total entries now: {len(self._address_book)}")

        # Fetch users first (internal accounts), then contacts
        self.sync_manager.fetch_users(account_id, callback=on_users)
        self.sync_manager.fetch_contacts(account_id, callback=on_contacts)

    def _on_address_key(self, event, entry):
        """Handle key press in address entry for autocomplete."""
        # Ignore navigation and control keys
        if event.keysym in ('Down', 'Up', 'Return', 'Tab', 'Escape', 'Shift_L', 'Shift_R',
                           'Control_L', 'Control_R', 'Alt_L', 'Alt_R', 'Left', 'Right'):
            return

        print(f"[Autocomplete] Key pressed: {event.keysym}, address_book size: {len(self._address_book)}")

        # Get current text being typed (after the last comma/semicolon)
        text = entry.get()
        cursor_pos = entry.index(tk.INSERT)

        # Find the start of current address (after last separator before cursor)
        last_sep = max(text.rfind(',', 0, cursor_pos), text.rfind(';', 0, cursor_pos))
        if last_sep >= 0:
            current_text = text[last_sep + 1:cursor_pos].strip()
        else:
            current_text = text[:cursor_pos].strip()

        if len(current_text) < 2:
            self._hide_autocomplete()
            return

        # Search for matches
        matches = self._search_address_book(current_text)
        if matches:
            self._show_autocomplete(entry, matches)
        else:
            self._hide_autocomplete()

    def _search_address_book(self, query):
        """Search address book and cached emails for matching contacts."""
        query_lower = query.lower()
        matches = []
        seen_emails = set()

        # Search address book entries (internal users have priority=0)
        for entry in self._address_book:
            name = entry.get('name', '').lower()
            email = entry.get('email', '').lower()

            if query_lower in name or query_lower in email:
                matches.append(entry)
                seen_emails.add(email)

        # Search cached emails (add if not already in address book)
        cached = self.db.get_cached_emails()
        for cached_entry in cached:
            email = cached_entry.get('email', '').lower()
            if email in seen_emails:
                continue

            name = cached_entry.get('name', '') or ''
            send_count = cached_entry.get('send_count', 0)

            if query_lower in name.lower() or query_lower in email:
                # Priority: 50 base for cached, minus send_count to prioritize frequent contacts
                # This puts cached after QLF (priority 0) but sorted by usage
                priority = max(1, 50 - send_count)
                matches.append({
                    'name': name,
                    'email': cached_entry.get('email', ''),
                    'type': 'cached',
                    'priority': priority,
                    'send_count': send_count
                })
                seen_emails.add(email)

        # Sort by priority (internal users first), then by name
        matches.sort(key=lambda x: (x.get('priority', 99), x.get('name', '').lower()))

        return matches[:10]  # Limit to 10 suggestions

    def _show_autocomplete(self, entry, matches):
        """Show autocomplete dropdown."""
        self._autocomplete_entry = entry

        # Create or update popup
        if self._autocomplete_popup is None:
            self._autocomplete_popup = tk.Toplevel(self.root)
            self._autocomplete_popup.wm_overrideredirect(True)
            self._autocomplete_popup.wm_attributes("-topmost", True)

            self._autocomplete_listbox = tk.Listbox(
                self._autocomplete_popup,
                width=50,
                height=min(10, len(matches)),
                font=("TkDefaultFont", 10),
                selectmode=tk.SINGLE,
                activestyle='dotbox'
            )
            self._autocomplete_listbox.pack(fill=tk.BOTH, expand=True)
            self._autocomplete_listbox.bind("<Double-1>", lambda e: self._autocomplete_select())
            self._autocomplete_listbox.bind("<Return>", lambda e: self._autocomplete_select())

        # Clear and populate
        self._autocomplete_listbox.delete(0, tk.END)
        self._autocomplete_matches = matches

        for m in matches:
            name = m.get('name', '')
            email = m.get('email', '')
            type_marker = ""
            if m.get('type') == 'user':
                type_marker = " [QLF]"
            elif m.get('type') == 'gal':
                type_marker = " [GAL]"

            if name and email:
                display = f"{name} <{email}>{type_marker}"
            else:
                display = f"{email}{type_marker}"
            self._autocomplete_listbox.insert(tk.END, display)

        # Position popup below entry
        x = entry.winfo_rootx()
        y = entry.winfo_rooty() + entry.winfo_height()
        self._autocomplete_popup.geometry(f"+{x}+{y}")
        self._autocomplete_listbox.config(height=min(10, len(matches)))

        # Select first item
        if matches:
            self._autocomplete_listbox.selection_set(0)
            self._autocomplete_listbox.see(0)

        self._autocomplete_popup.deiconify()

    def _hide_autocomplete(self):
        """Hide autocomplete dropdown."""
        if self._autocomplete_popup:
            self._autocomplete_popup.withdraw()
        self._autocomplete_entry = None
        return False  # Return False so Escape can chain to discard_compose

    def _autocomplete_navigate(self, direction):
        """Navigate autocomplete list."""
        if not self._autocomplete_popup or not self._autocomplete_popup.winfo_viewable():
            return

        listbox = self._autocomplete_listbox
        selection = listbox.curselection()
        if not selection:
            new_idx = 0
        else:
            new_idx = selection[0] + direction

        # Clamp to valid range
        new_idx = max(0, min(new_idx, listbox.size() - 1))

        listbox.selection_clear(0, tk.END)
        listbox.selection_set(new_idx)
        listbox.see(new_idx)
        return "break"

    def _autocomplete_select(self):
        """Select current autocomplete item."""
        if not self._autocomplete_popup or not self._autocomplete_popup.winfo_viewable():
            return False

        listbox = self._autocomplete_listbox
        selection = listbox.curselection()
        if not selection:
            self._hide_autocomplete()
            return False

        idx = selection[0]
        if idx < len(self._autocomplete_matches):
            match = self._autocomplete_matches[idx]
            self._insert_address(match)

        self._hide_autocomplete()
        return True

    def _insert_address(self, match):
        """Insert selected address into entry."""
        entry = self._autocomplete_entry
        if not entry:
            return

        name = match.get('name', '')
        email = match.get('email', '')
        if name and email:
            address = f"{name} <{email}>"
        else:
            address = email

        # Get current text and cursor position
        text = entry.get()
        cursor_pos = entry.index(tk.INSERT)

        # Find where current partial address starts
        last_sep = max(text.rfind(',', 0, cursor_pos), text.rfind(';', 0, cursor_pos))

        if last_sep >= 0:
            # Replace text after last separator
            prefix = text[:last_sep + 1] + " "
            suffix = text[cursor_pos:]
            # Check if there's more text after cursor that's not another address
            next_sep = min(
                text.find(',', cursor_pos) if text.find(',', cursor_pos) >= 0 else len(text),
                text.find(';', cursor_pos) if text.find(';', cursor_pos) >= 0 else len(text)
            )
            suffix = text[next_sep:] if next_sep < len(text) else ""
        else:
            prefix = ""
            # Find end of current address
            next_sep = min(
                text.find(',', cursor_pos) if text.find(',', cursor_pos) >= 0 else len(text),
                text.find(';', cursor_pos) if text.find(';', cursor_pos) >= 0 else len(text)
            )
            suffix = text[next_sep:] if next_sep < len(text) else ""

        # Build new text
        new_text = prefix + address + suffix

        # Update entry
        if entry == self.compose_to_entry:
            self.compose_to_var.set(new_text)
        elif entry == self.compose_cc_entry:
            self.compose_cc_var.set(new_text)

        # Move cursor to end of inserted address
        entry.icursor(len(prefix) + len(address))

    def _setup_compose_reply(self, reply_data):
        """Setup compose for reply."""
        import re
        import html as html_module

        reply_all = reply_data.get('reply_all', False)

        # Set To field
        from_name = reply_data.get('from_name', '')
        from_email = reply_data.get('from_email', '')
        if from_name and from_email:
            self.compose_to_var.set(f"{from_name} <{from_email}>")
        elif from_email:
            self.compose_to_var.set(from_email)

        # For Reply All, add original To recipients to To, original CC to CC
        if reply_all:
            my_email = ''
            idx = self.compose_from_combo.current()
            if idx >= 0 and idx < len(self._compose_accounts):
                my_email = self._compose_accounts[idx].get('email', '').lower()

            def should_include(addr):
                """Check if address should be included (not me, not original sender)."""
                addr_lower = addr.lower()
                if my_email and my_email in addr_lower:
                    return False
                if from_email and from_email.lower() in addr_lower:
                    return False
                return True

            # Add original To recipients to our To field (after the sender)
            original_to = reply_data.get('to', '')
            to_additions = []
            for addr in original_to.split(','):
                addr = addr.strip()
                if addr and should_include(addr):
                    to_additions.append(addr)

            if to_additions:
                current_to = self.compose_to_var.get()
                self.compose_to_var.set(f"{current_to}, {', '.join(to_additions)}")

            # Original CC stays as CC
            original_cc = reply_data.get('cc', '')
            cc_list = []
            for addr in original_cc.split(','):
                addr = addr.strip()
                if addr and should_include(addr):
                    cc_list.append(addr)

            if cc_list:
                self.compose_cc_var.set(', '.join(cc_list))

        # Set Subject
        subject = reply_data.get('subject', '')
        if not subject.lower().startswith('re:'):
            subject = f"Re: {subject}"
        self.compose_subject_var.set(subject)

        # Format date
        date = reply_data.get('date', '')
        if date and len(date) >= 15:
            try:
                from datetime import datetime
                date_part = date[:8]
                time_part = date[9:15]
                dt = datetime.strptime(f"{date_part}{time_part}", "%Y%m%d%H%M%S")
                date = dt.strftime("%B %d, %Y at %I:%M %p")
            except Exception:
                pass

        # Build quote header
        sender_str = f"{from_name} <{from_email}>" if from_name else from_email
        quote_header = f"\n\nOn {date}, {sender_str} wrote:\n" if date else f"\n\n{sender_str} wrote:\n"

        self.compose_body.insert(tk.END, quote_header)

        # Insert quoted body with formatting
        self._insert_html_quoted(reply_data.get('body', ''))

    def _setup_compose_forward(self, forward_data):
        """Setup compose for forwarding."""
        # Set Subject with Fwd: prefix
        subject = forward_data.get('subject', '')
        if not subject.lower().startswith('fwd:'):
            subject = f"Fwd: {subject}"
        self.compose_subject_var.set(subject)

        # Format date
        date = forward_data.get('date', '')
        if date and len(date) >= 15:
            try:
                from datetime import datetime
                date_part = date[:8]
                time_part = date[9:15]
                dt = datetime.strptime(f"{date_part}{time_part}", "%Y%m%d%H%M%S")
                date = dt.strftime("%B %d, %Y at %I:%M %p")
            except Exception:
                pass

        # Build forward header
        from_name = forward_data.get('from_name', '')
        from_email = forward_data.get('from_email', '')
        sender_str = f"{from_name} <{from_email}>" if from_name else from_email
        original_to = forward_data.get('to', '')
        original_subject = forward_data.get('subject', '')

        forward_header = "\n\n---------- Forwarded message ----------\n"
        forward_header += f"From: {sender_str}\n"
        forward_header += f"Date: {date}\n"
        forward_header += f"Subject: {original_subject}\n"
        forward_header += f"To: {original_to}\n\n"

        self.compose_body.insert(tk.END, forward_header)

        # Insert original body with formatting
        self._insert_html_quoted(forward_data.get('body', ''))

    def _insert_html_quoted(self, html_body):
        """Insert HTML as quoted text with formatting preserved."""
        import re
        import html as html_module

        if not html_body:
            return

        def clean_text(t):
            t = t.replace('\u00a0', ' ')
            t = t.replace('\u200b', '').replace('\u200c', '').replace('\u200d', '')
            t = t.replace('\u2003', ' ').replace('\u2002', ' ').replace('\u2009', ' ')
            t = t.replace('\ufeff', '').replace('\r\n', '\n').replace('\r', '\n')
            return t

        text = html_body

        # Remove style tags and their content
        text = re.sub(r'<style[^>]*>.*?</style>', '', text, flags=re.IGNORECASE | re.DOTALL)
        # Remove script tags and their content
        text = re.sub(r'<script[^>]*>.*?</script>', '', text, flags=re.IGNORECASE | re.DOTALL)
        # Remove head section
        text = re.sub(r'<head[^>]*>.*?</head>', '', text, flags=re.IGNORECASE | re.DOTALL)
        # Remove HTML comments (including conditional comments)
        text = re.sub(r'<!--.*?-->', '', text, flags=re.DOTALL)

        text = re.sub(r'<br\s*/?>', '\n', text, flags=re.IGNORECASE)
        text = re.sub(r'<p[^>]*>', '\n', text, flags=re.IGNORECASE)
        text = re.sub(r'</p>', '\n', text, flags=re.IGNORECASE)
        text = re.sub(r'<div[^>]*>', '\n', text, flags=re.IGNORECASE)
        text = re.sub(r'</div>', '', text, flags=re.IGNORECASE)

        # Parse tags
        segments = []
        current_tags = set()
        tag_pattern = re.compile(r'<(/?)(\w+)[^>]*>', re.IGNORECASE)
        pos = 0

        for match in tag_pattern.finditer(text):
            if match.start() > pos:
                chunk = html_module.unescape(text[pos:match.start()])
                chunk = clean_text(chunk)
                if chunk:
                    segments.append((chunk, frozenset(current_tags)))

            is_closing = match.group(1) == '/'
            tag_name = match.group(2).lower()

            if tag_name in ('b', 'strong'):
                current_tags.discard('bold') if is_closing else current_tags.add('bold')
            elif tag_name in ('i', 'em'):
                current_tags.discard('italic') if is_closing else current_tags.add('italic')
            elif tag_name == 'u':
                current_tags.discard('underline') if is_closing else current_tags.add('underline')

            pos = match.end()

        if pos < len(text):
            chunk = html_module.unescape(text[pos:])
            chunk = clean_text(chunk)
            if chunk:
                segments.append((chunk, frozenset(current_tags)))

        # Build full text and char_tags
        full_text = ''.join(seg[0] for seg in segments)
        full_text = re.sub(r'\n{3,}', '\n\n', full_text).strip()

        char_tags = []
        for seg_text, seg_tags in segments:
            for _ in seg_text:
                char_tags.append(seg_tags)

        # Mark the start of quoted content
        quote_start = self.compose_body.index(tk.END + "-1c")

        # Insert with formatting
        lines = full_text.split('\n')
        char_idx = 0

        for i, line in enumerate(lines):
            for ch in line:
                if char_idx < len(char_tags):
                    tags = char_tags[char_idx]
                    start_idx = self.compose_body.index(tk.END + "-1c")
                    self.compose_body.insert(tk.END, ch)
                    if tags:
                        end_idx = self.compose_body.index(tk.END + "-1c")
                        for tag in tags:
                            self.compose_body.tag_add(tag, start_idx, end_idx)
                else:
                    self.compose_body.insert(tk.END, ch)
                char_idx += 1
            if i < len(lines) - 1:
                self.compose_body.insert(tk.END, '\n')
                char_idx += 1

        # Apply "quoted" tag to all the inserted content
        quote_end = self.compose_body.index(tk.END + "-1c")
        self.compose_body.tag_add("quoted", quote_start, quote_end)

    def _compose_on_key_press(self, event):
        """Handle keypress for active formatting."""
        if event.char and event.char.isprintable() and len(event.char) == 1:
            insert_pos = self.compose_body.index(tk.INSERT)
            self.compose_body.after(1, lambda: self._compose_apply_active_tags(insert_pos))

    def _compose_apply_active_tags(self, pos):
        """Apply active tags to just-typed character."""
        try:
            end_pos = self.compose_body.index(f"{pos}+1c")
            for tag in self._compose_active_tags:
                self.compose_body.tag_add(tag, pos, end_pos)
        except tk.TclError:
            pass

    def _compose_toggle_bold(self):
        self._compose_toggle_tag("bold")
        return "break"

    def _compose_toggle_italic(self):
        self._compose_toggle_tag("italic")
        return "break"

    def _compose_toggle_underline(self):
        self._compose_toggle_tag("underline")
        return "break"

    def _compose_toggle_tag(self, tag_name):
        """Toggle formatting tag."""
        try:
            sel_start = self.compose_body.index(tk.SEL_FIRST)
            sel_end = self.compose_body.index(tk.SEL_LAST)
            current_tags = self.compose_body.tag_names(sel_start)
            if tag_name in current_tags:
                self.compose_body.tag_remove(tag_name, sel_start, sel_end)
            else:
                self.compose_body.tag_add(tag_name, sel_start, sel_end)
        except tk.TclError:
            if tag_name in self._compose_active_tags:
                self._compose_active_tags.discard(tag_name)
            else:
                self._compose_active_tags.add(tag_name)
        self._update_compose_format_buttons()

    def _update_compose_format_buttons(self):
        """Update button appearance based on active tags."""
        if "bold" in self._compose_active_tags:
            self.compose_bold_btn.config(relief=tk.SUNKEN, bg=self._compose_btn_active_bg)
        else:
            self.compose_bold_btn.config(relief=tk.RAISED, bg=self._compose_btn_bg)

        if "italic" in self._compose_active_tags:
            self.compose_italic_btn.config(relief=tk.SUNKEN, bg=self._compose_btn_active_bg)
        else:
            self.compose_italic_btn.config(relief=tk.RAISED, bg=self._compose_btn_bg)

        if "underline" in self._compose_active_tags:
            self.compose_underline_btn.config(relief=tk.SUNKEN, bg=self._compose_btn_active_bg)
        else:
            self.compose_underline_btn.config(relief=tk.RAISED, bg=self._compose_btn_bg)

    def _compose_change_font_size(self, event=None):
        """Change font size."""
        size = self.compose_font_size_var.get()
        size_tag = f"size_{size}"
        self._compose_active_tags = {t for t in self._compose_active_tags if not t.startswith("size_")}
        if size != "12":
            self._compose_active_tags.add(size_tag)

        try:
            sel_start = self.compose_body.index(tk.SEL_FIRST)
            sel_end = self.compose_body.index(tk.SEL_LAST)
            for s in [10, 12, 14, 16, 18, 20, 24, 28, 32]:
                self.compose_body.tag_remove(f"size_{s}", sel_start, sel_end)
            if size != "12":
                self.compose_body.tag_add(size_tag, sel_start, sel_end)
        except tk.TclError:
            pass

    def _compose_insert_bullet(self):
        """Insert bullet point."""
        self.compose_body.insert(tk.INSERT, "• ")

    def _attach_compose(self):
        """Add attachment (called from menu)."""
        self._compose_add_attachment_dialog()

    def _compose_add_attachment_dialog(self):
        """Open file dialog to add attachments."""
        from tkinter import filedialog
        files = filedialog.askopenfilenames(parent=self.root, title="Select files to attach")
        for filepath in files:
            try:
                with open(filepath, 'rb') as f:
                    content = f.read()
                import os
                name = os.path.basename(filepath)
                self._compose_attachments.append({'name': name, 'content': content})
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read {filepath}: {e}")

        self._refresh_compose_attachments()

    def _refresh_compose_attachments(self):
        """Refresh the compose attachment display."""
        # Clear existing buttons
        for widget in self.compose_attach_list_frame.winfo_children():
            widget.destroy()

        # Add button for each attachment with remove option
        for i, att in enumerate(self._compose_attachments):
            name = att.get('name', 'attachment')
            frame = ttk.Frame(self.compose_attach_list_frame)
            frame.pack(side=tk.LEFT, padx=2)
            # Make attachment name clickable to open/view
            ttk.Button(frame, text=f"📎 {name}",
                      command=lambda a=att: self._open_compose_attachment(a)).pack(side=tk.LEFT)
            ttk.Button(frame, text="×", width=2,
                      command=lambda idx=i: self._remove_compose_attachment(idx)).pack(side=tk.LEFT)

    def _open_compose_attachment(self, attachment):
        """Open a compose attachment for viewing."""
        import tempfile
        import subprocess
        import platform
        import os

        name = attachment.get('name', 'attachment')
        content = attachment.get('content')

        if not content:
            # No content - maybe it has a URL (shouldn't happen for compose)
            messagebox.showinfo("Attachment", f"No content available for {name}")
            return

        try:
            # Save to temp file
            _, ext = os.path.splitext(name)
            with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as f:
                if isinstance(content, str):
                    f.write(content.encode('utf-8'))
                else:
                    f.write(content)
                temp_path = f.name

            # Open with OS default app
            system = platform.system()
            if system == "Darwin":
                subprocess.Popen(["open", temp_path])
            elif system == "Windows":
                os.startfile(temp_path)
            else:
                subprocess.Popen(["xdg-open", temp_path])

            self.statusbar.config(text=f"Opened: {name}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open attachment: {e}")

    def _remove_compose_attachment(self, index):
        """Remove an attachment from compose."""
        if 0 <= index < len(self._compose_attachments):
            del self._compose_attachments[index]
            self._refresh_compose_attachments()

    def _discard_compose(self):
        """Discard compose and return to preview."""
        if messagebox.askyesno("Discard", "Discard this message?"):
            self._hide_compose()

    def _send_compose(self):
        """Send the composed message."""
        to = self.compose_to_var.get().strip()
        if not to:
            messagebox.showwarning("Send", "Please enter a recipient")
            return

        subject = self.compose_subject_var.get().strip()
        body = self._get_compose_html_body()

        # Get selected account
        idx = self.compose_from_combo.current()
        if idx < 0 or idx >= len(self._compose_accounts):
            messagebox.showwarning("Send", "Please select an account")
            return

        account = self._compose_accounts[idx]
        account_id = account['id']

        # Parse recipients - handle commas within names like "Schultz, Alison <email>"
        def split_addresses(text):
            """Split address list, respecting commas inside <brackets>."""
            addresses = []
            current = ""
            in_brackets = False
            for char in text:
                if char == '<':
                    in_brackets = True
                    current += char
                elif char == '>':
                    in_brackets = False
                    current += char
                elif char == ',' and not in_brackets:
                    if current.strip():
                        addresses.append(current.strip())
                    current = ""
                else:
                    current += char
            if current.strip():
                addresses.append(current.strip())
            return addresses

        def extract_email(addr):
            addr = addr.strip()
            if '<' in addr and '>' in addr:
                # Extract email from "Name <email>" format
                start = addr.rfind('<') + 1
                end = addr.rfind('>')
                return addr[start:end].strip()
            return addr

        def extract_name_email(addr):
            """Extract both name and email from address."""
            addr = addr.strip()
            if '<' in addr and '>' in addr:
                start = addr.rfind('<') + 1
                end = addr.rfind('>')
                email = addr[start:end].strip()
                name = addr[:addr.rfind('<')].strip().strip('"')
                return name, email
            return None, addr

        to_addresses = split_addresses(to)
        to_list = [extract_email(addr) for addr in to_addresses if addr]
        cc = self.compose_cc_var.get().strip()
        cc_addresses = split_addresses(cc) if cc else []
        cc_list = [extract_email(addr) for addr in cc_addresses if addr]

        # Cache recipient addresses with send count increment
        for addr in to_addresses:
            if addr:
                name, email = extract_name_email(addr)
                self.db.add_email_to_cache(email, name, increment_send=True)
        for addr in cc_addresses:
            if addr:
                name, email = extract_name_email(addr)
                self.db.add_email_to_cache(email, name, increment_send=True)

        self.statusbar.config(text="Sending...")

        # Get original message info for reply/forward marking
        original_id = getattr(self, '_compose_original_id', None)
        compose_mode = getattr(self, '_compose_mode', 'new')

        self.sync_manager.send_message(
            account_id, to_list, subject, body, cc=cc_list,
            attachments=self._compose_attachments,
            original_id=original_id,
            is_reply=(compose_mode == "reply"),
            is_forward=(compose_mode == "forward"),
            callback=self._on_compose_sent
        )

    def _on_compose_sent(self, success, error=None):
        """Called when send completes."""
        if success:
            self.statusbar.config(text="Message sent")
            self._hide_compose()
        else:
            messagebox.showerror("Send Failed", str(error))
            self.statusbar.config(text="Send failed")

    def _get_compose_html_body(self):
        """Convert compose body to HTML."""
        html_parts = []
        index = "1.0"
        in_quote = False

        while True:
            next_idx = self.compose_body.index(f"{index}+1c")
            if self.compose_body.compare(next_idx, ">=", tk.END):
                break

            char = self.compose_body.get(index, next_idx)
            tags = self.compose_body.tag_names(index)
            is_quoted = 'quoted' in tags

            # Handle quote block transitions
            if is_quoted and not in_quote:
                html_parts.append('<blockquote style="border-left:2px solid #ccc;padding-left:10px;margin:10px 0 10px 5px;color:#555">')
                in_quote = True
            elif not is_quoted and in_quote:
                html_parts.append('</blockquote>')
                in_quote = False

            if char == '\n':
                html_parts.append('<br>')
            else:
                styled_char = char
                if char in ('&', '<', '>'):
                    styled_char = {'&': '&amp;', '<': '&lt;', '>': '&gt;'}[char]
                if 'bold' in tags:
                    styled_char = f'<b>{styled_char}</b>'
                if 'italic' in tags:
                    styled_char = f'<i>{styled_char}</i>'
                if 'underline' in tags:
                    styled_char = f'<u>{styled_char}</u>'
                for tag in tags:
                    if tag.startswith('size_'):
                        size = tag.split('_')[1]
                        styled_char = f'<span style="font-size:{size}px">{styled_char}</span>'
                        break
                html_parts.append(styled_char)
            index = next_idx

        # Close any open quote block
        if in_quote:
            html_parts.append('</blockquote>')

        return ''.join(html_parts)

    def _create_statusbar(self):
        """Create the status bar."""
        self.statusbar = ttk.Label(self.root, text="Ready", anchor=tk.W, relief=tk.SUNKEN)
        self.statusbar.pack(side=tk.BOTTOM, fill=tk.X, padx=5, pady=2)

    # ==================== Account Management ====================

    def _add_account(self):
        """Show add account dialog."""
        from mailbench.dialogs.account_dialog import AccountDialog
        dialog = AccountDialog(self.root, self.db, self.kerio_pool, app=self)
        self.root.wait_window(dialog.dialog)
        self._refresh_folder_tree()

    def _manage_accounts(self):
        """Show manage accounts dialog."""
        from mailbench.dialogs.account_dialog import AccountDialog
        dialog = AccountDialog(self.root, self.db, self.kerio_pool, app=self)
        self.root.wait_window(dialog.dialog)
        self._refresh_folder_tree()

    def _connect_account(self, account_id):
        """Connect to an account."""
        account = self.db.get_account(account_id)
        if not account:
            return

        config = KerioConfig(
            email=account['email'],
            username=account['username'],
            password=account['password'],
            server=account['server']
        )

        try:
            self.kerio_pool.connect(account_id, config)
            self.connected_accounts.add(account_id)
            self.statusbar.config(text=f"Connected to {account['name']}")

            # Sync folders
            self.sync_manager.sync_folders(account_id, callback=self._on_folders_synced)

            # Start listening for changes (new mail notifications)
            self.sync_manager.start_change_listener(account_id, callback=self._on_mailbox_changed)

            # Load address book for autocomplete
            self._load_address_book(account_id)

        except Exception as e:
            messagebox.showerror("Connection Error", str(e))

    def _on_folders_synced(self, success, error=None):
        """Called when folder sync completes."""
        if success:
            self._refresh_folder_tree()
            self._update_status()
            # Auto-select INBOX on first sync if nothing is selected
            if not hasattr(self, '_current_folder_id'):
                self.root.after(200, self._auto_select_inbox)
        else:
            self.statusbar.config(text=f"Sync failed: {error}")

    def _on_mailbox_changed(self, account_id: int, changes: list):
        """Called when mailbox changes are detected (new mail, etc.)."""
        if not changes:
            return

        # Check if any change affects the currently displayed folder
        current_account = getattr(self, '_current_account_id', None)
        current_folder = getattr(self, '_current_folder_id', None)

        should_refresh = False
        for change in changes:
            # For mail items, parentId is the folder; for folders, itemId is the folder
            change_folder = change.get('parentId') if not change.get('isFolder') else change.get('itemId')
            if change_folder == current_folder:
                should_refresh = True
                break

        if should_refresh and current_account == account_id:
            # Refresh the message list for current folder (incremental, no flashing)
            self.sync_manager.sync_messages(
                current_account, current_folder,
                callback=self._on_auto_refresh_messages
            )
            self.statusbar.config(text="New messages received")

    # ==================== Folder Tree ====================

    def _refresh_folder_tree(self, force_rebuild=False):
        """Refresh the folder tree with accounts and folders."""
        accounts = self.db.get_accounts()
        existing_children = set(self.folder_tree.get_children())

        # Check if we need a full rebuild
        expected_account_ids = {f"account_{a['id']}" for a in accounts}
        if force_rebuild or existing_children != expected_account_ids:
            self._rebuild_folder_tree(accounts)
            return

        # Check if any account is missing folders
        for account in accounts:
            account_id = account['id']
            account_node = f"account_{account_id}"
            folders = self.db.get_folders(account_id)

            # Check if folder count matches
            existing_folder_children = set(self.folder_tree.get_children(account_node))
            expected_folder_ids = {f"folder_{account_id}_{f['folder_id']}" for f in folders}
            if existing_folder_children != expected_folder_ids:
                self._rebuild_folder_tree(accounts)
                return

        # Just update existing items (no structural change)
        for account in accounts:
            account_id = account['id']
            account_node = f"account_{account_id}"

            # Update account status text
            status = " (connected)" if account_id in self.connected_accounts else ""
            new_text = f"{account['name']}{status}"
            if self.folder_tree.item(account_node, 'text') != new_text:
                self.folder_tree.item(account_node, text=new_text)

            # Update folder unread counts
            folders = self.db.get_folders(account_id)
            for folder in folders:
                folder_node = f"folder_{account_id}_{folder['folder_id']}"
                if self.folder_tree.exists(folder_node):
                    unread = folder.get('unread_count', 0)
                    display_name = folder['name']
                    if unread > 0:
                        display_name = f"{folder['name']} ({unread})"
                    if self.folder_tree.item(folder_node, 'text') != display_name:
                        self.folder_tree.item(folder_node, text=display_name)

    def _rebuild_folder_tree(self, accounts=None):
        """Full rebuild of folder tree."""
        self.folder_tree.delete(*self.folder_tree.get_children())

        if accounts is None:
            accounts = self.db.get_accounts()

        for account in accounts:
            account_id = account['id']
            name = account['name']

            # Add account node
            status = " (connected)" if account_id in self.connected_accounts else ""
            account_node = self.folder_tree.insert("", tk.END, iid=f"account_{account_id}",
                                                   text=f"{name}{status}", open=True)

            # Add folders
            folders = self.db.get_folders(account_id)

            # Filter out non-email folders (PIM folders like Calendar, Contacts, etc.)
            non_email_folders = {
                'calendar', 'contacts', 'tasks', 'journal', 'notes',
                'search root', 'suggested contacts'
            }
            folders = [f for f in folders if (
                # Keep standard email folders
                f.get('folder_type') in ('inbox', 'drafts', 'sent', 'junk', 'trash', 'outbox')
                # Or custom folders that aren't PIM folders
                or (f.get('folder_type') == 'custom'
                    and f['name'].lower() not in non_email_folders
                    and not f['name'].startswith('__')
                    and 'IPM_ROOT' not in f['name'])
            )]

            # Sort folders: standard folders first, then custom
            standard_order = {'inbox': 0, 'drafts': 1, 'sent': 2, 'junk': 3, 'trash': 4, 'outbox': 5}

            def folder_sort_key(f):
                ft = f.get('folder_type', 'custom')
                return (standard_order.get(ft, 100), f['name'].lower())

            folders.sort(key=folder_sort_key)

            for folder in folders:
                unread = folder.get('unread_count', 0)
                display_name = folder['name']
                if unread > 0:
                    display_name = f"{folder['name']} ({unread})"

                self.folder_tree.insert(account_node, tk.END,
                                        iid=f"folder_{account_id}_{folder['folder_id']}",
                                        text=display_name)

        # If no accounts, show hint
        if not accounts:
            self.folder_tree.insert("", tk.END, text="No accounts configured")
            self.folder_tree.insert("", tk.END, text="Use Accounts > Add Account...")

    def _on_folder_select(self, event):
        """Handle folder selection."""
        selection = self.folder_tree.selection()
        if not selection:
            return

        item_id = selection[0]
        if item_id.startswith("folder_"):
            # Parse folder_accountid_folderid
            parts = item_id.split("_", 2)
            if len(parts) >= 3:
                account_id = int(parts[1])
                folder_id = parts[2]
                # Skip if already loading this folder (prevents overwriting select_first flag)
                if (getattr(self, '_current_account_id', None) == account_id and
                    getattr(self, '_current_folder_id', None) == folder_id):
                    return
                # Clear filter when changing folders
                self.search_var.set("")
                self._load_messages(account_id, folder_id)

    def _on_folder_double_click(self, event):
        """Handle folder double-click."""
        selection = self.folder_tree.selection()
        if not selection:
            return

        item_id = selection[0]
        if item_id.startswith("account_"):
            account_id = int(item_id.replace("account_", ""))
            if account_id not in self.connected_accounts:
                self._connect_account(account_id)

    # ==================== Message List ====================

    def _load_messages(self, account_id, folder_id, select_first=False):
        """Load messages for a folder - always fetch from server."""
        self.message_list.delete(*self.message_list.get_children())
        self._current_account_id = account_id
        self._current_folder_id = folder_id
        self._select_first_after_load = select_first

        # Style unread messages
        self.message_list.tag_configure('unread', font=("TkDefaultFont", 10, "bold"))

        # Always sync from server if connected - don't show stale cache
        if account_id in self.connected_accounts:
            self.statusbar.config(text="Loading messages...")
            self.sync_manager.sync_messages(account_id, folder_id, callback=self._on_messages_synced)
        else:
            # Only show cache if not connected
            self._display_messages_from_data(self.db.get_messages(account_id, folder_id, limit=50))

    def _display_messages_from_data(self, messages, clear_first=True):
        """Display messages incrementally to keep UI responsive."""
        import time
        self._load_start_time = time.time()

        # Convert to list first (in case it's a generator)
        messages = list(messages)

        if clear_first:
            self.message_list.delete(*self.message_list.get_children())
            self._existing_msg_ids = set()
            self._messages_by_id = {}  # Store message data for quick lookup
        else:
            # Always refresh from treeview when appending to avoid duplicates
            self._existing_msg_ids = set(self.message_list.get_children())
        if not hasattr(self, '_messages_by_id'):
            self._messages_by_id = {}

        # Store messages to add incrementally
        self._pending_messages = list(messages)
        self._add_messages_batch()

    def _add_messages_batch(self):
        """Add a batch of messages to the list, then schedule next batch."""
        if not hasattr(self, '_pending_messages') or not self._pending_messages:
            # Done - select first message if requested
            if getattr(self, '_select_first_after_load', False):
                self._select_first_after_load = False
                self.root.after(50, self._select_first_message)
            # Show load time
            import time
            elapsed = time.time() - getattr(self, '_load_start_time', time.time())
            self.statusbar.config(text=f"Loaded {len(self.message_list.get_children())} messages in {elapsed:.2f}s")
            return

        # Add batch of 100 messages at a time
        batch_size = 100
        batch = self._pending_messages[:batch_size]
        self._pending_messages = self._pending_messages[batch_size:]

        for msg in batch:
            item_id = msg.get('item_id')
            if not item_id or item_id in self._existing_msg_ids:
                continue

            # Store message data for later retrieval
            self._messages_by_id[item_id] = msg

            sender = msg.get('sender_name') or msg.get('sender_email') or 'Unknown'
            subject = msg.get('subject') or '(No Subject)'
            flag = UNI_FLAG if msg.get('is_flagged') else ""
            # Status: replied takes precedence over forwarded
            if msg.get('is_answered'):
                status = UNI_REPLY
            elif msg.get('is_forwarded'):
                status = UNI_FORWARD
            else:
                status = ""
            attach = (FA_PAPERCLIP if self._has_fontawesome else "*") if msg.get('has_attachments') else ""
            date = msg.get('date_received', '')
            if date:
                date = self._format_kerio_date(date)

            tags = () if msg.get('is_read') else ('unread',)
            self.message_list.insert("", tk.END, iid=item_id,
                                     values=(flag, status, attach, sender, subject, date), tags=tags)
            self._existing_msg_ids.add(item_id)

        # Schedule next batch (1ms delay lets UI breathe)
        self.root.after(1, self._add_messages_batch)

    def _select_first_message(self):
        """Select the first message in the list."""
        children = self.message_list.get_children()
        if children:
            self.message_list.selection_set(children[0])
            self.message_list.see(children[0])
            self.message_list.focus(children[0])
            # Trigger the selection event
            self._on_message_select(None)

    def _on_message_list_scroll(self, *args):
        """Handle message list scroll - load more when near bottom."""
        # Update the scrollbar
        self._msg_scrollbar.set(*args)

        # Check if we're near the bottom (args is (first, last) as fractions)
        if len(args) >= 2:
            _, last = float(args[0]), float(args[1])
            # If we're 90% scrolled and not already loading
            if last > 0.9 and not self._messages_loading_more:
                self._load_more_messages()

    def _load_more_messages(self):
        """Load more messages when scrolling near bottom."""
        if not hasattr(self, '_current_account_id') or not hasattr(self, '_current_folder_id'):
            return
        if self._current_account_id not in self.connected_accounts:
            return
        if self._messages_loading_more:
            return

        current_count = len(self.message_list.get_children())
        if current_count == 0:
            return

        self._messages_loading_more = True
        self.statusbar.config(text="Loading more messages...")

        # Fetch more messages with offset
        self.sync_manager.sync_messages(
            self._current_account_id,
            self._current_folder_id,
            limit=current_count + 50,  # Load 50 more
            callback=self._on_more_messages_loaded
        )

    def _on_more_messages_loaded(self, success, error=None, messages_data=None):
        """Called when more messages are loaded."""
        self._messages_loading_more = False
        if success and messages_data:
            # Add only new messages (incremental)
            self._display_messages_from_data(messages_data, clear_first=False)
            self.statusbar.config(text=f"{len(self.message_list.get_children())} messages")
        elif error:
            self.statusbar.config(text=f"Error: {error}")

    def _on_messages_synced(self, success, error=None, messages_data=None):
        """Called when message sync completes - display messages directly."""
        if success:
            # Remember if we need to select first message BEFORE display clears it
            should_select_first = getattr(self, '_select_first_after_load', False)
            self._select_first_after_load = False  # Clear it now

            if messages_data:
                # Display messages directly from server response
                self._display_messages_from_data(messages_data)
            else:
                # Fallback to cache
                self._display_messages_from_data(
                    self.db.get_messages(self._current_account_id, self._current_folder_id, limit=50)
                )

            self.statusbar.config(text=f"{len(self.message_list.get_children())} messages")
            self._update_folder_unread_count()

            # Select first message if this was the initial load
            # Do this AFTER a delay to let batch loading complete
            if should_select_first:
                self.root.after(500, self._select_first_message)
        else:
            self.statusbar.config(text=f"Sync failed: {error}")

    def _update_folder_unread_count(self):
        """Update unread count display for current folder."""
        if not hasattr(self, '_current_account_id') or not hasattr(self, '_current_folder_id'):
            return

        # Get folder info from database
        folders = self.db.get_folders(self._current_account_id)
        for folder in folders:
            if folder['folder_id'] == self._current_folder_id:
                # Count unread in our local cache
                messages = self.db.get_messages(self._current_account_id, self._current_folder_id, limit=1000)
                unread_count = sum(1 for m in messages if not m.get('is_read'))

                # Update treeview display
                tree_id = f"folder_{self._current_account_id}_{self._current_folder_id}"
                if self.folder_tree.exists(tree_id):
                    display_name = folder['name']
                    if unread_count > 0:
                        display_name = f"{folder['name']} ({unread_count})"
                    self.folder_tree.item(tree_id, text=display_name)
                break

    def _format_kerio_date(self, date_str: str) -> str:
        """Format Kerio date string (YYYYMMDDTHHMMSS-TTTT) to readable format.

        Uses Odoo-style formatting:
        - Today: time only (e.g., "10:31 AM")
        - Yesterday: "Yesterday"
        - Older: date only (e.g., "12/31/2025")
        """
        try:
            from datetime import datetime, date, timedelta
            # Kerio format: 20251231T054339-0600
            if not date_str or len(date_str) < 15:
                return date_str
            # Extract date/time parts
            date_part = date_str[:8]  # YYYYMMDD
            time_part = date_str[9:15]  # HHMMSS (skip the T)
            # Parse
            dt = datetime.strptime(f"{date_part}{time_part}", "%Y%m%d%H%M%S")

            today = date.today()
            msg_date = dt.date()

            if msg_date == today:
                # Today: show time only (e.g., "10:31 AM")
                hour = dt.hour % 12 or 12
                return f"{hour}:{dt.strftime('%M')} {dt.strftime('%p')}"
            elif msg_date == today - timedelta(days=1):
                # Yesterday
                return "Yesterday"
            else:
                # Older: show date only (e.g., "12/31/2025")
                return f"{dt.month}/{dt.day}/{dt.year}"
        except Exception:
            return date_str

    def _on_message_select(self, event):
        """Handle message selection - show preview."""
        selection = self.message_list.selection()
        if not selection:
            return

        # Enable/disable Reply buttons based on selection count
        multi_select = len(selection) > 1
        self.reply_btn.config(state="disabled" if multi_select else "normal")
        self.reply_all_btn.config(state="disabled" if multi_select else "normal")

        # For multi-select, don't update preview
        if multi_select:
            self._clear_preview_pane()
            self.statusbar.config(text=f"{len(selection)} messages selected")
            return

        item_id = selection[0]  # This is the Kerio item_id

        # Skip if same message already displayed (prevents flash during refresh)
        if getattr(self, '_current_message_id', None) == item_id:
            return
        self._current_message_id = item_id

        # Get basic info from our in-memory message data
        msg_data = self._messages_by_id.get(item_id, {}) if hasattr(self, '_messages_by_id') else {}

        # Show header info immediately from list data
        self._show_message_header(item_id, msg_data)

        # Always fetch full message from API
        if hasattr(self, '_current_account_id'):
            self.statusbar.config(text="Loading message...")
            self.sync_manager.fetch_message_body(
                self._current_account_id, item_id,
                callback=lambda success, data, _: self._on_body_fetched(item_id, success, data)
            )

    def _show_message_header(self, item_id, msg_data):
        """Show message header info immediately from cached list data."""
        # Get sender info
        sender_name = msg_data.get('sender_name', '')
        sender_email = msg_data.get('sender_email', '')
        if sender_name and sender_email:
            sender = f"{sender_name} <{sender_email}>"
        elif sender_name:
            sender = sender_name
        elif sender_email:
            sender = sender_email
        else:
            sender = 'Unknown'

        self.preview_from.config(text=f"From: {sender}")
        self.preview_to.config(text="To: Loading...")  # Will be updated when body fetched
        self.preview_subject.config(text=msg_data.get('subject', '(No Subject)'))

        # Format date
        date = msg_data.get('date_received', '')
        if date:
            try:
                from datetime import datetime
                if isinstance(date, str) and len(date) >= 15:
                    date_part = date[:8]
                    time_part = date[9:15]
                    dt = datetime.strptime(f"{date_part}{time_part}", "%Y%m%d%H%M%S")
                    date = dt.strftime("%A, %B %d, %Y at %I:%M %p")
            except Exception:
                pass
        self.preview_date.config(text=date)

        # Clear body while loading
        if self.preview_html:
            is_dark = self.dark_mode_var.get()
            bg = "#313335" if is_dark else "#ffffff"
            fg = "#a9b7c6" if is_dark else "#000000"
            self.preview_html.load_html(
                f"<html><body style='background-color:{bg};color:{fg};font-family:sans-serif;'>"
                f"<p style='color:#888;'>Loading message...</p></body></html>"
            )
        elif self.preview_text:
            self.preview_text.config(state=tk.NORMAL)
            self.preview_text.delete("1.0", tk.END)
            self.preview_text.insert("1.0", "Loading message...")
            self.preview_text.config(state=tk.DISABLED)

    def _clear_preview_pane(self):
        """Clear the preview pane when no message is selected."""
        self._current_message_id = None
        self.preview_from.config(text="From:")
        self.preview_to.config(text="To:")
        self.preview_subject.config(text="")
        self.preview_date.config(text="")

        # Clear attachments
        for widget in self.preview_attach_frame.winfo_children():
            widget.destroy()

        if self.preview_html:
            is_dark = self.dark_mode_var.get()
            bg = "#313335" if is_dark else "#ffffff"
            fg = "#a9b7c6" if is_dark else "#000000"
            self.preview_html.load_html(
                f"<html><body style='background-color:{bg};color:{fg};font-family:sans-serif;'>"
                f"<p style='color:#888;text-align:center;margin-top:50px;'>No message selected</p></body></html>"
            )
        elif self.preview_text:
            self.preview_text.config(state=tk.NORMAL)
            self.preview_text.delete("1.0", tk.END)
            self.preview_text.insert("1.0", "No message selected")
            self.preview_text.config(state=tk.DISABLED)

    def _show_preview_attachments(self, attachments):
        """Show attachments in the preview pane header."""
        # Clear existing attachment widgets
        for widget in self.preview_attach_frame.winfo_children():
            widget.destroy()

        if not attachments:
            return

        # Add label
        ttk.Label(self.preview_attach_frame, text="Attachments:",
                  font=("TkDefaultFont", 9, "bold")).pack(side=tk.LEFT)

        # Add clickable attachment links
        for att in attachments:
            name = att.get('name', 'attachment')
            size = att.get('size', 0)
            # Format size
            if size >= 1024 * 1024:
                size_str = f"({size / (1024 * 1024):.1f} MB)"
            elif size >= 1024:
                size_str = f"({size / 1024:.1f} KB)"
            else:
                size_str = f"({size} bytes)" if size > 0 else ""

            btn = ttk.Button(self.preview_attach_frame, text=f"📎 {name} {size_str}",
                           command=lambda a=att: self._download_attachment(a))
            btn.pack(side=tk.LEFT, padx=3)

    def _download_attachment(self, attachment):
        """Download an attachment and open or save it."""
        from tkinter import filedialog
        import webbrowser

        name = attachment.get('name', 'attachment')
        url = attachment.get('url', '')

        if not url:
            messagebox.showerror("Error", "Attachment URL not available")
            return

        # Get server from current session
        if not hasattr(self, '_current_account_id'):
            messagebox.showerror("Error", "No account connected")
            return

        session = self.kerio_pool.get_session(self._current_account_id)
        if not session:
            messagebox.showerror("Error", "Account not connected")
            return

        # Build full URL
        full_url = f"https://{session.config.server}{url}"

        # Ask user what to do with themed dialog
        result = self._show_attachment_dialog(name)

        if result == "cancel":
            return
        elif result == "open":
            # Download to temp file and open with OS default app
            self._open_attachment(full_url, name, session)
        elif result == "save":
            save_path = self._show_save_dialog(name)
            if save_path:
                self._save_attachment(full_url, save_path, session)

    def _show_attachment_dialog(self, filename):
        """Show themed dialog for attachment download options."""
        result = {"value": "cancel"}

        dialog = tk.Toplevel(self.root)
        dialog.title("Download Attachment")
        dialog.transient(self.root)
        dialog.grab_set()

        # Make it modal and centered
        dialog.resizable(False, False)

        # Apply theme colors
        is_dark = self.dark_mode_var.get()
        bg = "#313335" if is_dark else "#f0f0f0"
        fg = "#a9b7c6" if is_dark else "#000000"
        dialog.configure(bg=bg)

        # Content frame
        frame = ttk.Frame(dialog, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)

        # Icon and message
        ttk.Label(frame, text="Attachment:", font=("TkDefaultFont", 10, "bold")).pack(anchor=tk.W)
        ttk.Label(frame, text=filename, wraplength=300).pack(anchor=tk.W, pady=(5, 15))
        ttk.Label(frame, text="What would you like to do?").pack(anchor=tk.W, pady=(0, 10))

        # Button frame
        btn_frame = ttk.Frame(frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))

        def on_open():
            result["value"] = "open"
            dialog.destroy()

        def on_save():
            result["value"] = "save"
            dialog.destroy()

        def on_cancel():
            result["value"] = "cancel"
            dialog.destroy()

        ttk.Button(btn_frame, text="Open", command=on_open).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="Save", command=on_save).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=on_cancel).pack(side=tk.RIGHT)

        # Handle escape key and window close
        dialog.bind("<Escape>", lambda e: on_cancel())
        dialog.protocol("WM_DELETE_WINDOW", on_cancel)

        # Center dialog on parent
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() - dialog.winfo_width()) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")

        # Wait for dialog to close
        dialog.wait_window()

        return result["value"]

    def _get_default_save_directory(self):
        """Get default save directory from settings or system Downloads folder."""
        import os
        # Check settings first
        save_dir = self.db.get_setting("default_save_directory", "")
        if save_dir and os.path.isdir(save_dir):
            return save_dir
        # Fall back to Downloads directory (works on Windows, Mac, Linux)
        downloads = os.path.join(os.path.expanduser("~"), "Downloads")
        if os.path.isdir(downloads):
            return downloads
        # Ultimate fallback to home directory
        return os.path.expanduser("~")

    def _show_save_dialog(self, initial_filename):
        """Show themed file save dialog."""
        import os
        result = {"path": None}

        dialog = tk.Toplevel(self.root)
        dialog.title("Save Attachment")
        dialog.transient(self.root)
        dialog.grab_set()

        # Apply theme colors
        is_dark = self.dark_mode_var.get()
        bg = "#313335" if is_dark else "#f0f0f0"
        fg = "#a9b7c6" if is_dark else "#000000"
        entry_bg = "#3c3f41" if is_dark else "#ffffff"
        list_bg = "#2b2b2b" if is_dark else "#ffffff"
        select_bg = "#4b6eaf" if is_dark else "#0078d7"
        dialog.configure(bg=bg)

        # Current directory - use default save directory
        current_dir = tk.StringVar(value=self._get_default_save_directory())

        # Main frame
        main_frame = ttk.Frame(dialog, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Directory selection row
        dir_frame = ttk.Frame(main_frame)
        dir_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(dir_frame, text="Save in:").pack(side=tk.LEFT)
        dir_entry = ttk.Entry(dir_frame, textvariable=current_dir, width=50)
        dir_entry.pack(side=tk.LEFT, padx=(5, 5), fill=tk.X, expand=True)

        def browse_dir():
            from tkinter import filedialog
            new_dir = filedialog.askdirectory(initialdir=current_dir.get())
            if new_dir:
                current_dir.set(new_dir)
                refresh_file_list()

        ttk.Button(dir_frame, text="...", width=3, command=browse_dir).pack(side=tk.LEFT)

        # File list
        list_frame = ttk.Frame(main_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Use Listbox for file listing
        file_listbox = tk.Listbox(list_frame, bg=list_bg, fg=fg, selectbackground=select_bg,
                                   selectforeground="#ffffff", height=10, width=60)
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=file_listbox.yview)
        file_listbox.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        def refresh_file_list():
            file_listbox.delete(0, tk.END)
            try:
                path = current_dir.get()
                # Add parent directory option
                file_listbox.insert(tk.END, "..")
                # List directories first, then files
                items = sorted(os.listdir(path))
                dirs = [d for d in items if os.path.isdir(os.path.join(path, d))]
                files = [f for f in items if os.path.isfile(os.path.join(path, f))]
                for d in dirs:
                    file_listbox.insert(tk.END, f"[{d}]")
                for f in files:
                    file_listbox.insert(tk.END, f)
            except PermissionError:
                file_listbox.insert(tk.END, "(Permission denied)")
            except Exception as e:
                file_listbox.insert(tk.END, f"(Error: {e})")

        def on_listbox_double_click(event):
            selection = file_listbox.curselection()
            if not selection:
                return
            item = file_listbox.get(selection[0])
            if item == "..":
                # Go up one directory
                parent = os.path.dirname(current_dir.get())
                if parent:
                    current_dir.set(parent)
                    refresh_file_list()
            elif item.startswith("[") and item.endswith("]"):
                # Navigate into directory
                dirname = item[1:-1]
                new_path = os.path.join(current_dir.get(), dirname)
                current_dir.set(new_path)
                refresh_file_list()
            else:
                # It's a file - put it in the filename entry
                filename_var.set(item)

        file_listbox.bind("<Double-1>", on_listbox_double_click)

        # Filename entry row
        name_frame = ttk.Frame(main_frame)
        name_frame.pack(fill=tk.X, pady=(0, 10))
        ttk.Label(name_frame, text="File name:").pack(side=tk.LEFT)
        filename_var = tk.StringVar(value=initial_filename)
        filename_entry = ttk.Entry(name_frame, textvariable=filename_var, width=50)
        filename_entry.pack(side=tk.LEFT, padx=(5, 0), fill=tk.X, expand=True)

        # Button row
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X)

        def on_save():
            filename = filename_var.get().strip()
            if filename:
                result["path"] = os.path.join(current_dir.get(), filename)
                dialog.destroy()

        def on_cancel():
            dialog.destroy()

        ttk.Button(btn_frame, text="Save", command=on_save).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(btn_frame, text="Cancel", command=on_cancel).pack(side=tk.RIGHT)

        # Handle escape key and window close
        dialog.bind("<Escape>", lambda e: on_cancel())
        dialog.bind("<Return>", lambda e: on_save())
        dialog.protocol("WM_DELETE_WINDOW", on_cancel)

        # Initialize file list
        refresh_file_list()

        # Center dialog on parent
        dialog.update_idletasks()
        dialog.minsize(500, 400)
        x = self.root.winfo_x() + (self.root.winfo_width() - 500) // 2
        y = self.root.winfo_y() + (self.root.winfo_height() - 400) // 2
        dialog.geometry(f"500x400+{x}+{y}")

        # Focus filename entry
        filename_entry.focus_set()
        filename_entry.select_range(0, tk.END)

        # Wait for dialog to close
        dialog.wait_window()

        return result["path"]

    def _save_attachment(self, url, save_path, session):
        """Save attachment to file."""
        import threading

        def do_save():
            try:
                import requests
                # Use session cookies for authentication
                cookies = {session.cookie_name: session.cookie_value} if hasattr(session, 'cookie_name') else {}
                headers = {"X-Token": session.token} if hasattr(session, 'token') else {}

                response = requests.get(url, cookies=cookies, headers=headers, verify=False, timeout=60)
                response.raise_for_status()

                with open(save_path, 'wb') as f:
                    f.write(response.content)

                self.root.after(0, lambda: self.statusbar.config(text=f"Saved: {save_path}"))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Error", f"Failed to save attachment: {e}"))

        self.statusbar.config(text="Downloading attachment...")
        threading.Thread(target=do_save, daemon=True).start()

    def _open_attachment(self, url, filename, session):
        """Download attachment to temp file and open with OS default app."""
        import threading
        import tempfile
        import subprocess
        import platform

        def do_open():
            try:
                import requests
                # Use session cookies for authentication
                cookies = {session.cookie_name: session.cookie_value} if hasattr(session, 'cookie_name') else {}
                headers = {"X-Token": session.token} if hasattr(session, 'token') else {}

                response = requests.get(url, cookies=cookies, headers=headers, verify=False, timeout=60)
                response.raise_for_status()

                # Save to temp file with original extension
                import os
                _, ext = os.path.splitext(filename)
                with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as f:
                    f.write(response.content)
                    temp_path = f.name

                # Open with OS default app
                system = platform.system()
                if system == "Darwin":  # macOS
                    subprocess.run(["open", temp_path])
                elif system == "Windows":
                    os.startfile(temp_path)
                else:  # Linux
                    subprocess.run(["xdg-open", temp_path])

                self.root.after(0, lambda: self.statusbar.config(text="Opened attachment"))
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Error", f"Failed to open attachment: {e}"))

        self.statusbar.config(text="Downloading attachment...")
        threading.Thread(target=do_open, daemon=True).start()

    def _preview_select_all(self):
        """Select all text in preview pane."""
        if self.preview_text:
            self.preview_text.tag_add(tk.SEL, "1.0", tk.END)
            self.preview_text.mark_set(tk.INSERT, "1.0")
            self.preview_text.see(tk.INSERT)
        return "break"

    def _on_shift_arrow_select(self, event):
        """Handle Shift+Arrow for multi-select in message list."""
        current = self.message_list.focus()
        if not current:
            return

        children = self.message_list.get_children()
        if not children:
            return

        try:
            idx = children.index(current)
        except ValueError:
            return

        # Determine direction
        if event.keysym == "Up":
            new_idx = max(0, idx - 1)
        else:  # Down
            new_idx = min(len(children) - 1, idx + 1)

        new_item = children[new_idx]

        # Add to selection (extend selection)
        current_selection = set(self.message_list.selection())
        current_selection.add(new_item)
        self.message_list.selection_set(list(current_selection))
        self.message_list.focus(new_item)
        self.message_list.see(new_item)

        return "break"

    def _is_over_preview(self, event):
        """Check if mouse event is over the preview pane."""
        try:
            # Get preview frame bounds
            frame = self._preview_body_frame
            fx = frame.winfo_rootx()
            fy = frame.winfo_rooty()
            fw = frame.winfo_width()
            fh = frame.winfo_height()
            # Check if event coordinates are within the frame
            return fx <= event.x_root <= fx + fw and fy <= event.y_root <= fy + fh
        except Exception:
            return False

    def _on_ctrl_scroll_up(self, event):
        """Handle Ctrl+scroll up globally."""
        if self._is_over_preview(event):
            self._preview_zoom_in()
            return "break"

    def _on_ctrl_scroll_down(self, event):
        """Handle Ctrl+scroll down globally."""
        if self._is_over_preview(event):
            self._preview_zoom_out()
            return "break"

    def _on_ctrl_scroll_wheel(self, event):
        """Handle Ctrl+mousewheel globally (Windows/macOS)."""
        if self._is_over_preview(event):
            if event.delta > 0:
                self._preview_zoom_in()
            elif event.delta < 0:
                self._preview_zoom_out()
            return "break"

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

    def _apply_preview_zoom(self):
        """Apply current zoom level and refresh display."""
        # Save zoom level
        self.db.set_setting("preview_zoom", str(self._preview_zoom))

        # For text widget fallback, update font size
        if self.preview_text:
            text_size = int(12 * self._preview_zoom / 100)
            self.preview_text.configure(font=("TkDefaultFont", text_size))

        # For HTML, re-render current message body if we have it cached
        if self.preview_html and hasattr(self, '_current_message_full_data') and self._current_message_full_data:
            # Re-render by calling body_fetched with cached data
            item_id = getattr(self, '_current_message_id', None)
            if item_id:
                self._on_body_fetched(item_id, True, self._current_message_full_data)

        # Update compose body zoom (display only, doesn't affect sent email)
        if hasattr(self, 'compose_body'):
            zoom_factor = self._preview_zoom / 100
            base_size = getattr(self, 'font_size', 12)
            zoomed_size = int(base_size * zoom_factor)
            self.compose_body.configure(font=("TkDefaultFont", zoomed_size))
            self._setup_compose_tags()  # Re-setup tags with zoomed sizes

        # Update status bar with zoom level (after re-render)
        self.statusbar.config(text=f"Zoom: {self._preview_zoom}%")

    def _on_message_click(self, event):
        """Handle single click - check if flag column was clicked to toggle flag."""
        # Identify what was clicked
        region = self.message_list.identify_region(event.x, event.y)
        if region != "cell":
            return

        column = self.message_list.identify_column(event.x)
        item = self.message_list.identify_row(event.y)

        if not item:
            return

        # Column #1 is the flag column
        if column == "#1":
            self._toggle_flag(item)

    def _toggle_flag(self, item_id):
        """Toggle the flag status of a message."""
        if not hasattr(self, '_current_account_id'):
            return

        msg = self._messages_by_id.get(item_id, {})
        current_flag = msg.get('is_flagged', False)
        new_flag = not current_flag

        # Update local state immediately for responsiveness
        msg['is_flagged'] = new_flag
        flag_icon = UNI_FLAG if new_flag else ""

        # Update the display
        values = list(self.message_list.item(item_id, 'values'))
        values[0] = flag_icon
        self.message_list.item(item_id, values=values)

        # Update server
        self.sync_manager.set_flag(self._current_account_id, item_id, new_flag,
                                   callback=lambda success, error: self._on_flag_toggled(success, error, item_id, new_flag))

    def _on_flag_toggled(self, success, error, item_id, new_flag):
        """Handle flag toggle response from server."""
        if not success:
            # Revert local state on failure
            msg = self._messages_by_id.get(item_id, {})
            msg['is_flagged'] = not new_flag
            flag_icon = UNI_FLAG if not new_flag else ""
            values = list(self.message_list.item(item_id, 'values'))
            values[0] = flag_icon
            self.message_list.item(item_id, values=values)
            self.statusbar.config(text=f"Failed to toggle flag: {error}")

    def _on_delete_key(self, event):
        """Handle Delete key press - delete selected messages."""
        self._delete_selected_messages()

    def _delete_selected_messages(self):
        """Delete the currently selected messages."""
        selection = self.message_list.selection()
        if not selection:
            return

        if not hasattr(self, '_current_account_id'):
            return

        # Confirm if multiple messages
        count = len(selection)
        if count > 1:
            if not messagebox.askyesno("Delete Messages",
                                       f"Delete {count} selected messages?"):
                return

        # Delete each message
        for item_id in selection:
            self.sync_manager.delete_message(
                self._current_account_id, item_id,
                callback=lambda success, error, iid=item_id: self._on_message_deleted(success, error, iid)
            )

    def _on_message_deleted(self, success, error, item_id):
        """Handle message deletion response."""
        if success:
            # Remove from list
            if self.message_list.exists(item_id):
                self.message_list.delete(item_id)
            # Remove from cache
            if item_id in self._messages_by_id:
                del self._messages_by_id[item_id]
            if item_id in self._existing_msg_ids:
                self._existing_msg_ids.discard(item_id)
        else:
            self.statusbar.config(text=f"Delete failed: {error}")

    def _show_message_context_menu(self, event):
        """Show context menu on right-click."""
        # Select the item under cursor if not already selected
        item = self.message_list.identify_row(event.y)
        if item:
            current_selection = self.message_list.selection()
            if item not in current_selection:
                self.message_list.selection_set(item)

        selection = self.message_list.selection()
        if not selection:
            return

        # Clear and rebuild menu based on selection
        menu = self._message_context_menu
        menu.delete(0, tk.END)

        # Get message info for context-aware menu items
        item_id = selection[0]
        msg = self._messages_by_id.get(item_id, {})
        is_read = msg.get('is_read', True)
        is_flagged = msg.get('is_flagged', False)
        count = len(selection)

        # Open
        menu.add_command(label="Open", command=self._context_open_message)
        menu.add_separator()

        # Reply options
        menu.add_command(label="Reply", command=self._context_reply)
        menu.add_command(label="Reply All", command=self._context_reply_all)
        menu.add_command(label="Forward", command=self._context_forward)
        menu.add_separator()

        # Mark as read/unread
        if is_read:
            menu.add_command(label="Mark as Unread", command=self._context_mark_unread)
        else:
            menu.add_command(label="Mark as Read", command=self._context_mark_read)

        # Flag/unflag
        if is_flagged:
            menu.add_command(label="Remove Flag", command=self._context_unflag)
        else:
            menu.add_command(label="Flag", command=self._context_flag)

        menu.add_separator()

        # Move to folder submenu
        move_menu = tk.Menu(menu, tearoff=0)
        if hasattr(self, '_current_account_id'):
            folders = self.db.get_folders(self._current_account_id)
            for folder in folders:
                folder_name = folder.get('name', '')
                folder_id = folder.get('folder_id', '')
                if folder_id and folder_id != getattr(self, '_current_folder_id', ''):
                    move_menu.add_command(
                        label=folder_name,
                        command=lambda fid=folder_id: self._context_move_to_folder(fid)
                    )
        menu.add_cascade(label="Move to", menu=move_menu)
        menu.add_separator()

        # Delete
        delete_label = f"Delete ({count} messages)" if count > 1 else "Delete"
        menu.add_command(label=delete_label, command=self._delete_selected_messages)

        # Show menu at cursor
        menu.tk_popup(event.x_root, event.y_root)

    def _context_open_message(self):
        """Open selected message in new window."""
        selection = self.message_list.selection()
        if selection:
            item_id = selection[0]
            msg_data = self._messages_by_id.get(item_id, {})
            full_data = getattr(self, '_current_message_full_data', None) if item_id == getattr(self, '_current_message_id', None) else None
            self._open_message_window(item_id, msg_data, full_data)

    def _context_reply(self):
        """Reply to selected message."""
        self._reply()

    def _context_reply_all(self):
        """Reply all to selected message."""
        self._reply_all()

    def _context_forward(self):
        """Forward selected message."""
        self._forward()

    def _context_mark_read(self):
        """Mark selected messages as read."""
        self._mark_selected_messages(is_read=True)

    def _context_mark_unread(self):
        """Mark selected messages as unread."""
        self._mark_selected_messages(is_read=False)

    def _context_flag(self):
        """Flag selected messages."""
        for item_id in self.message_list.selection():
            msg = self._messages_by_id.get(item_id, {})
            if not msg.get('is_flagged', False):
                self._toggle_flag(item_id)

    def _context_unflag(self):
        """Unflag selected messages."""
        for item_id in self.message_list.selection():
            msg = self._messages_by_id.get(item_id, {})
            if msg.get('is_flagged', False):
                self._toggle_flag(item_id)

    def _context_move_to_folder(self, folder_id):
        """Move selected messages to a folder."""
        if not hasattr(self, '_current_account_id'):
            return

        selection = self.message_list.selection()
        for item_id in selection:
            self.sync_manager.move_message(
                self._current_account_id, item_id, folder_id,
                callback=lambda success, error, iid=item_id: self._on_message_moved(success, error, iid)
            )

    def _on_message_moved(self, success, error, item_id):
        """Handle message move response."""
        if success:
            # Remove from current list
            if self.message_list.exists(item_id):
                self.message_list.delete(item_id)
            if item_id in self._messages_by_id:
                del self._messages_by_id[item_id]
            if item_id in self._existing_msg_ids:
                self._existing_msg_ids.discard(item_id)
        else:
            self.statusbar.config(text=f"Move failed: {error}")

    def _mark_selected_messages(self, is_read=True):
        """Mark selected messages as read or unread."""
        if not hasattr(self, '_current_account_id'):
            return

        selection = self.message_list.selection()
        for item_id in selection:
            # Update local state
            msg = self._messages_by_id.get(item_id, {})
            msg['is_read'] = is_read

            # Update display
            tags = () if is_read else ('unread',)
            self.message_list.item(item_id, tags=tags)

            # Update server
            self.sync_manager.mark_as_read(
                self._current_account_id, item_id, is_read,
                callback=lambda s, e: None  # Silent callback
            )

    def _on_message_double_click(self, event):
        """Handle message double-click - open in new window."""
        selection = self.message_list.selection()
        if not selection:
            return

        item_id = selection[0]

        # Get message data
        msg_data = self._messages_by_id.get(item_id, {}) if hasattr(self, '_messages_by_id') else {}
        if not msg_data:
            return

        # Get full message data if we have it cached
        full_data = getattr(self, '_current_message_full_data', None) if item_id == getattr(self, '_current_message_id', None) else None

        # Open message in new window
        self._open_message_window(item_id, msg_data, full_data)

    def _open_message_window(self, item_id, msg_data, full_data=None):
        """Open a message in its own window."""
        # Create new window
        win = tk.Toplevel(self.root)
        subject = msg_data.get('subject', '(No Subject)')
        win.title(subject)
        win.geometry("800x600")

        # Apply theme
        is_dark = self.dark_mode_var.get()
        if is_dark:
            bg = "#313335"
            fg = "#a9b7c6"
        else:
            bg = "#ffffff"
            fg = "#000000"

        win.configure(bg=bg)

        # Header frame
        header_frame = ttk.Frame(win)
        header_frame.pack(fill=tk.X, padx=10, pady=10)

        # Format sender
        sender_name = msg_data.get('sender_name', '')
        sender_email = msg_data.get('sender_email', '')
        if sender_name and sender_email:
            sender = f"{sender_name} <{sender_email}>"
        else:
            sender = sender_name or sender_email or 'Unknown'

        ttk.Label(header_frame, text=f"From: {sender}", font=("TkDefaultFont", 10, "bold")).pack(anchor=tk.W)

        # Get To/CC from full data if available
        to_str = full_data.get('to', '') if full_data else ''
        cc_str = full_data.get('cc', '') if full_data else ''
        ttk.Label(header_frame, text=f"To: {to_str}").pack(anchor=tk.W)
        if cc_str:
            ttk.Label(header_frame, text=f"CC: {cc_str}").pack(anchor=tk.W)

        ttk.Label(header_frame, text=f"Subject: {subject}", font=("TkDefaultFont", 10, "bold")).pack(anchor=tk.W, pady=(5, 0))

        # Format date
        date = msg_data.get('date_received', '')
        if date and len(date) >= 15:
            try:
                from datetime import datetime
                date_part = date[:8]
                time_part = date[9:15]
                dt = datetime.strptime(f"{date_part}{time_part}", "%Y%m%d%H%M%S")
                date = dt.strftime("%A, %B %d, %Y at %I:%M %p")
            except Exception:
                pass
        ttk.Label(header_frame, text=date, foreground="gray").pack(anchor=tk.W)

        ttk.Separator(win, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=10, pady=5)

        # Body frame
        body_frame = ttk.Frame(win)
        body_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Display body
        if HAS_HTML_VIEW:
            html_widget = HtmlFrame(body_frame, messages_enabled=False)
            html_widget.pack(fill=tk.BOTH, expand=True)

            # Get body content
            body = full_data.get('body', '') if full_data else ''

            if not body:
                # Fetch the body if not cached
                if hasattr(self, '_current_account_id'):
                    self._fetch_message_for_window(item_id, html_widget, win)
                    return
                body = "<p>Message body not available</p>"

            # Render HTML with theme
            self._render_html_in_widget(html_widget, body, is_dark)
        else:
            # Fallback to text
            text_widget = tk.Text(body_frame, wrap=tk.WORD, bg=bg, fg=fg)
            text_widget.pack(fill=tk.BOTH, expand=True)
            body = full_data.get('body', '') if full_data else 'Message body not available'
            text_widget.insert("1.0", self._html_to_text(body))
            text_widget.config(state=tk.DISABLED)

    def _fetch_message_for_window(self, item_id, html_widget, win):
        """Fetch message body for a separate window."""
        def on_fetched(success, data, _):
            if not success or not win.winfo_exists():
                return
            body = data.get('body', '') if isinstance(data, dict) else ''
            is_dark = self.dark_mode_var.get()
            self._render_html_in_widget(html_widget, body, is_dark)

        self.sync_manager.fetch_message_body(
            self._current_account_id, item_id,
            callback=on_fetched
        )

    def _render_html_in_widget(self, html_widget, body, is_dark):
        """Render HTML content in an HtmlFrame widget."""
        if is_dark:
            bg_color = "#313335"
            fg_color = "#a9b7c6"
            link_color = "#589df6"
            quote_bg = "#3a3a3a"
            quote_border = "#666666"
        else:
            bg_color = "#ffffff"
            fg_color = "#000000"
            link_color = "#0066cc"
            quote_bg = "#f5f5f5"
            quote_border = "#cccccc"

        zoom_factor = self._preview_zoom / 100
        html_font_size = int(14 * zoom_factor)
        css = f"""<style>
            html {{ background-color: {bg_color} !important; color: {fg_color} !important; overflow-x: auto; }}
            body {{ background-color: {bg_color} !important; color: {fg_color} !important;
                    font-family: sans-serif; font-size: {html_font_size}px !important; margin: 8px; line-height: 1.4; }}
            body, p, div, span, td, th, li, a, font, b, i, strong, em, h1, h2, h3, h4, h5, h6 {{
                font-size: {html_font_size}px !important; }}
            p {{ margin: 0.5em 0; color: {fg_color} !important; }}
            div {{ margin: 0.2em 0; color: {fg_color} !important; }}
            blockquote {{ border-left: 3px solid {quote_border}; margin: 1em 0; padding: 8px 12px;
                          background-color: {quote_bg}; color: {fg_color} !important; }}
            pre, code {{ white-space: pre-wrap; font-family: monospace; background-color: {quote_bg}; padding: 2px 4px; }}
            a {{ color: {link_color} !important; }}
            td, th {{ color: {fg_color} !important; }}
            hr {{ border: none; border-top: 1px solid {quote_border}; margin: 1em 0; }}
            .gmail_quote, .yahoo_quoted, .mailbench-quote-block {{
                border-left: 3px solid {quote_border}; margin: 1em 0;
                padding: 8px 12px; background-color: {quote_bg}; }}
        </style>"""

        import re
        html_body = body if body else "<p></p>"
        # Check for real HTML tags (not email addresses like <user@domain.com>)
        # Real tags have tagname followed by space, >, or /
        has_real_html = bool(re.search(r'<[a-zA-Z][a-zA-Z0-9]*[\s>/]', html_body))

        if not has_real_html:
            import html as html_module
            # Normalize line endings first
            normalized = html_body.replace('\r\n', '\n').replace('\r', '\n')
            escaped_body = html_module.escape(normalized).replace('\n', '<br>')
            html_body = f"<html><head>{css}</head><body>{escaped_body}</body></html>"
        elif '<html' not in html_body.lower():
            html_body = f"<html><head>{css}</head><body>{html_body}</body></html>"
        else:
            if '<head>' in html_body.lower():
                html_body = re.sub(r'(<head[^>]*>)', r'\1' + css, html_body, flags=re.IGNORECASE)
            elif '<html' in html_body.lower():
                html_body = re.sub(r'(<html[^>]*>)', r'\1<head>' + css + '</head>', html_body, flags=re.IGNORECASE)

        html_body = self._enhance_quote_styling(html_body)
        html_widget.load_html(html_body)

    def _show_message_preview(self, message_id):
        """Show message in preview pane."""
        msg = self.db.get_message(message_id)
        if not msg:
            return

        # Update header
        sender = msg.get('sender_name', '')
        if msg.get('sender_email'):
            sender = f"{sender} <{msg['sender_email']}>" if sender else msg['sender_email']
        self.preview_from.config(text=f"From: {sender}")

        # Get recipients from cache if available
        cached = getattr(self, '_message_cache', {}).get(message_id, {})
        to_str = cached.get('to', '')
        cc_str = cached.get('cc', '')

        self.preview_to.config(text=f"To: {to_str}" if to_str else "To:")
        # Could also show CC if needed

        self.preview_subject.config(text=msg.get('subject', '(No Subject)'))

        date = msg.get('date_received', '')
        if date:
            try:
                from datetime import datetime
                # Kerio format: 20251231T054339-0600
                if isinstance(date, str) and len(date) >= 15:
                    date_part = date[:8]  # YYYYMMDD
                    time_part = date[9:15]  # HHMMSS
                    dt = datetime.strptime(f"{date_part}{time_part}", "%Y%m%d%H%M%S")
                    date = dt.strftime("%A, %B %d, %Y at %I:%M %p")
            except Exception:
                pass
        self.preview_date.config(text=date)

        # Update body - prefer cached data from fetch_message_body
        body = cached.get('body') or msg.get('body') or msg.get('body_preview') or ''
        body_type = cached.get('body_type') or msg.get('body_type', 'html')

        if self.preview_html:
            # Use HTML view
            html_body = body if body else "<p></p>"

            # Get theme colors
            is_dark = self.dark_mode_var.get()
            if is_dark:
                bg_color = "#313335"
                fg_color = "#a9b7c6"
                link_color = "#589df6"
            else:
                bg_color = "#ffffff"
                fg_color = "#000000"
                link_color = "#0066cc"

            # CSS that forces dark/light mode - use !important to override email styles
            quote_bg = "#3a3a3a" if is_dark else "#f5f5f5"
            quote_border = "#666666" if is_dark else "#cccccc"
            # Zoom factor for scaling content
            zoom_factor = self._preview_zoom / 100
            html_font_size = int(14 * zoom_factor)
            css = f"""<style>
                html {{ background-color: {bg_color} !important; color: {fg_color} !important; overflow-x: auto; }}
                body {{ background-color: {bg_color} !important; color: {fg_color} !important;
                        font-family: sans-serif; font-size: {html_font_size}px !important; margin: 8px; line-height: 1.4; }}
                body, p, div, span, td, th, li, a, font, b, i, strong, em, h1, h2, h3, h4, h5, h6 {{
                    font-size: {html_font_size}px !important; }}
                p {{ margin: 0.5em 0; color: {fg_color} !important; }}
                div {{ margin: 0.2em 0; color: {fg_color} !important; }}
                blockquote {{ border-left: 3px solid {quote_border}; margin: 1em 0; padding: 8px 12px;
                              background-color: {quote_bg}; color: {fg_color} !important; }}
                blockquote blockquote {{ margin: 0.5em 0; }}
                pre, code {{ white-space: pre-wrap; word-wrap: break-word; font-family: monospace;
                            background-color: {quote_bg}; padding: 2px 4px; }}
                pre {{ padding: 8px; margin: 0.5em 0; }}
                a {{ color: {link_color} !important; }}
                span {{ color: inherit; }}
                td, th {{ color: {fg_color} !important; }}
                table {{ border-collapse: collapse; }}
                td, th {{ padding: 4px 8px; }}
                hr {{ border: none; border-top: 1px solid {quote_border}; margin: 1em 0; }}
                /* Common email quote styles */
                .gmail_quote, .yahoo_quoted {{ border-left: 3px solid {quote_border}; margin: 1em 0;
                                               padding: 8px 12px; background-color: {quote_bg}; }}
                /* Outlook/Office quote header style - div with border-top */
                .mailbench-quote-block {{ background-color: {quote_bg}; padding: 8px 12px; margin: 1em 0;
                                          border-left: 3px solid {quote_border}; }}
                /* Preserve original colored text */
                span[style*="color"] {{ color: inherit !important; }}
            </style>"""

            # Check if content is plain text - trust body_type from API, but also
            # verify by checking for actual HTML structure (not just angle brackets
            # which could be email addresses like <user@domain.com>)
            import re
            # Only consider it HTML if body_type says so AND it has real HTML tags
            # (tags followed by > or space/attr, not @ which indicates email)
            has_real_html = (body_type == 'html' and
                           bool(re.search(r'<[a-zA-Z][a-zA-Z0-9]*[\s>]', html_body)))

            if not has_real_html:
                # Plain text - escape and convert newlines to <br>
                import html as html_module
                # Normalize line endings first
                normalized = html_body.replace('\r\n', '\n').replace('\r', '\n')
                escaped_body = html_module.escape(normalized)
                escaped_body = escaped_body.replace('\n', '<br>')
                html_body = f"<html><head>{css}</head><body>{escaped_body}</body></html>"
            elif '<html' not in html_body.lower():
                # Has HTML tags but no wrapper - just wrap it
                html_body = f"<html><head>{css}</head><body>{html_body}</body></html>"
            else:
                # Complete HTML document - inject our CSS into head
                if '<head>' in html_body.lower():
                    html_body = re.sub(r'(<head[^>]*>)', r'\1' + css, html_body, flags=re.IGNORECASE)
                elif '<html' in html_body.lower():
                    html_body = re.sub(r'(<html[^>]*>)', r'\1<head>' + css + '</head>', html_body, flags=re.IGNORECASE)

            # Post-process: detect and style Outlook quote blocks
            # Outlook uses <div style='border:none;border-top:solid...'>  with From:/Sent: headers
            html_body = self._enhance_quote_styling(html_body)
            self.preview_html.load_html(html_body)
        else:
            # Fallback to plain text
            body = self._html_to_text(body)
            self.preview_text.config(state=tk.NORMAL)
            self.preview_text.delete("1.0", tk.END)
            self.preview_text.insert("1.0", body)
            self.preview_text.config(state=tk.DISABLED)

        # Mark as read (locally and on server)
        if not msg.get('is_read'):
            self.db.update_message_read(message_id, True)
            # Update display - use item_id from the message
            ews_item_id = msg.get('item_id')
            if ews_item_id and self.message_list.exists(ews_item_id):
                self.message_list.item(ews_item_id, tags=())
            # Sync read status to server
            if hasattr(self, '_current_account_id') and ews_item_id:
                self.sync_manager.mark_as_read(self._current_account_id, ews_item_id, True)
            # Update folder unread count
            self._update_folder_unread_count()

        # Fetch full body if not cached
        if not msg.get('body') and hasattr(self, '_current_account_id'):
            self.statusbar.config(text="Loading message body...")
            self.sync_manager.fetch_message_body(
                self._current_account_id, msg['item_id'],
                callback=lambda success, msg_data, _: self._on_body_fetched(message_id, success, msg_data)
            )

    def _on_body_fetched(self, item_id, success, msg_data):
        """Called when message body is fetched."""
        if not success:
            self.statusbar.config(text=f"Failed to load message: {msg_data}")
            return

        if not isinstance(msg_data, dict):
            self.statusbar.config(text="Invalid message data")
            return

        # Store full message data for reply/forward
        self._current_message_full_data = msg_data

        # Update To/CC in header
        to_str = msg_data.get('to', '')
        cc_str = msg_data.get('cc', '')
        self.preview_to.config(text=f"To: {to_str}" if to_str else "To:")

        # Display attachments
        self._show_preview_attachments(msg_data.get('attachments', []))

        # Get body content
        body = msg_data.get('body', '')
        body_type = msg_data.get('body_type', 'html')

        # Display the body
        if self.preview_html:
            html_body = body if body else "<p></p>"

            # Get theme colors
            is_dark = self.dark_mode_var.get()
            if is_dark:
                bg_color = "#313335"
                fg_color = "#a9b7c6"
                link_color = "#589df6"
            else:
                bg_color = "#ffffff"
                fg_color = "#000000"
                link_color = "#0066cc"

            # CSS for consistent theming with quote styling
            quote_bg = "#3a3a3a" if is_dark else "#f5f5f5"
            quote_border = "#666666" if is_dark else "#cccccc"
            # Zoom factor for scaling content
            zoom_factor = self._preview_zoom / 100
            html_font_size = int(14 * zoom_factor)
            css = f"""<style>
                html {{ background-color: {bg_color} !important; color: {fg_color} !important; overflow-x: auto; }}
                body {{ background-color: {bg_color} !important; color: {fg_color} !important;
                        font-family: sans-serif; font-size: {html_font_size}px !important; margin: 8px; line-height: 1.4; }}
                body, p, div, span, td, th, li, a, font, b, i, strong, em, h1, h2, h3, h4, h5, h6 {{
                    font-size: {html_font_size}px !important; }}
                p {{ margin: 0.5em 0; color: {fg_color} !important; }}
                div {{ margin: 0.2em 0; color: {fg_color} !important; }}
                blockquote {{ border-left: 3px solid {quote_border}; margin: 1em 0; padding: 8px 12px;
                              background-color: {quote_bg}; color: {fg_color} !important; }}
                pre, code {{ white-space: pre-wrap; word-wrap: break-word; font-family: monospace;
                            background-color: {quote_bg}; padding: 2px 4px; }}
                a {{ color: {link_color} !important; }}
                span {{ color: inherit; }}
                td, th {{ color: {fg_color} !important; }}
                hr {{ border: none; border-top: 1px solid {quote_border}; margin: 1em 0; }}
                .gmail_quote, .yahoo_quoted, .mailbench-quote-block {{
                    border-left: 3px solid {quote_border}; margin: 1em 0;
                    padding: 8px 12px; background-color: {quote_bg}; }}
            </style>"""

            # Check if content is plain text - trust body_type from API, but also
            # verify by checking for actual HTML structure (not just angle brackets
            # which could be email addresses like <user@domain.com>)
            import re
            # Only consider it HTML if body_type says so AND it has real HTML tags
            # (tags followed by > or space/attr, not @ which indicates email)
            has_real_html = (body_type == 'html' and
                           bool(re.search(r'<[a-zA-Z][a-zA-Z0-9]*[\s>]', html_body)))

            if not has_real_html:
                # Plain text - escape and convert newlines to <br>
                import html as html_module
                # Normalize line endings first
                normalized = html_body.replace('\r\n', '\n').replace('\r', '\n')
                escaped_body = html_module.escape(normalized)
                escaped_body = escaped_body.replace('\n', '<br>')
                html_body = f"<html><head>{css}</head><body>{escaped_body}</body></html>"
            elif '<html' not in html_body.lower():
                # Has HTML tags but no wrapper
                html_body = f"<html><head>{css}</head><body>{html_body}</body></html>"
            else:
                # Complete HTML document - inject our CSS
                if '<head>' in html_body.lower():
                    html_body = re.sub(r'(<head[^>]*>)', r'\1' + css, html_body, flags=re.IGNORECASE)
                elif '<html' in html_body.lower():
                    html_body = re.sub(r'(<html[^>]*>)', r'\1<head>' + css + '</head>', html_body, flags=re.IGNORECASE)

            # Enhance quote styling
            html_body = self._enhance_quote_styling(html_body)
            self.preview_html.load_html(html_body)
        else:
            # Fallback to plain text
            text_body = self._html_to_text(body)
            self.preview_text.config(state=tk.NORMAL)
            self.preview_text.delete("1.0", tk.END)
            self.preview_text.insert("1.0", text_body)
            self.preview_text.config(state=tk.DISABLED)

        # Mark message as read in the list (remove bold)
        if self.message_list.exists(item_id):
            self.message_list.item(item_id, tags=())
            # Also update on server
            if hasattr(self, '_current_account_id'):
                self.sync_manager.mark_as_read(self._current_account_id, item_id, True)

        self.statusbar.config(text="Message loaded")

    def _enhance_quote_styling(self, html_body):
        """Detect and enhance styling for quoted email content."""
        import re

        # Pattern 1: Outlook-style quote with border-top containing From:/Sent:
        # Look for <div style='...border-top:solid...'>...<b>From:</b>...
        outlook_quote_pattern = re.compile(
            r'(<div[^>]*style=["\'][^"\']*border-top[^"\']*["\'][^>]*>)',
            re.IGNORECASE
        )

        def add_quote_class(match):
            tag = match.group(1)
            # Add our class to the div
            if 'class=' in tag.lower():
                tag = re.sub(r'class=["\']([^"\']*)["\']',
                            r'class="\1 mailbench-quote-block"', tag, flags=re.IGNORECASE)
            else:
                tag = tag.replace('>', ' class="mailbench-quote-block">', 1)
            return tag

        html_body = outlook_quote_pattern.sub(add_quote_class, html_body)

        # Preserve blank lines: convert <br><br> to a visible empty line
        # Use a div with fixed height since tkhtmlview collapses consecutive <br> tags
        html_body = re.sub(r'(<br\s*/?>){2,}', r'<br><div style="height:0.5em"></div>', html_body, flags=re.IGNORECASE)

        return html_body

    def _html_to_text(self, html):
        """Convert HTML to plain text."""
        import re
        import html as html_module

        if not html:
            return ''

        text = html

        # Remove style and script blocks entirely
        text = re.sub(r'<style[^>]*>.*?</style>', '', text, flags=re.DOTALL | re.IGNORECASE)
        text = re.sub(r'<script[^>]*>.*?</script>', '', text, flags=re.DOTALL | re.IGNORECASE)
        text = re.sub(r'<head[^>]*>.*?</head>', '', text, flags=re.DOTALL | re.IGNORECASE)

        # Remove HTML comments
        text = re.sub(r'<!--.*?-->', '', text, flags=re.DOTALL)

        # Replace line breaks
        text = re.sub(r'<br[^>]*/?>', '\n', text, flags=re.IGNORECASE)

        # Block elements - only add newline on closing tag
        text = re.sub(r'</p>', '\n', text, flags=re.IGNORECASE)
        text = re.sub(r'</div>', '\n', text, flags=re.IGNORECASE)
        text = re.sub(r'</tr>', '\n', text, flags=re.IGNORECASE)
        text = re.sub(r'<li[^>]*>', '• ', text, flags=re.IGNORECASE)
        text = re.sub(r'</li>', '\n', text, flags=re.IGNORECASE)

        # Remove all remaining HTML tags
        text = re.sub(r'<[^>]+>', '', text)

        # Decode HTML entities
        text = html_module.unescape(text)

        # Clean up whitespace
        text = text.replace('\r\n', '\n').replace('\r', '\n')  # Normalize line endings
        text = re.sub(r'[ \t]+', ' ', text)  # Multiple spaces to single
        text = re.sub(r' *\n *', '\n', text)  # Remove spaces around newlines
        text = re.sub(r'\n{3,}', '\n\n', text)  # Max 2 consecutive newlines
        text = text.strip()

        # Remove emoji and other characters above BMP that tkinter can't render
        text = ''.join(c if ord(c) < 0x10000 else '' for c in text)

        return text

    # ==================== Actions ====================

    def _new_message(self):
        """Open inline compose for new message."""
        self._show_compose(title="New Message")

    def _reply(self):
        """Reply to selected message."""
        self._do_reply(reply_all=False)

    def _reply_all(self):
        """Reply all to selected message."""
        self._do_reply(reply_all=True)

    def _do_reply(self, reply_all=False):
        """Common reply handler."""
        selection = self.message_list.selection()
        if not selection:
            return

        # Get full message data
        msg_data = getattr(self, '_current_message_full_data', None)
        if not msg_data:
            messagebox.showwarning("Reply", "Please wait for message to load")
            return

        # Get list data for sender info
        item_id = selection[0]
        list_data = self._messages_by_id.get(item_id, {})

        # Combine data
        reply_data = {
            'from_name': list_data.get('sender_name', ''),
            'from_email': list_data.get('sender_email', ''),
            'to': msg_data.get('to', ''),
            'cc': msg_data.get('cc', ''),
            'subject': list_data.get('subject', ''),
            'date': list_data.get('date_received', ''),
            'body': msg_data.get('body', ''),
            'reply_all': reply_all,
            'original_id': item_id  # Track original message for marking as answered
        }

        title = "Reply All" if reply_all else "Reply"
        self._show_compose(title=title, reply_data=reply_data)

    def _forward(self):
        """Forward selected message(s)."""
        selection = self.message_list.selection()
        if not selection:
            return

        # Multi-select: forward as attachments
        if len(selection) > 1:
            self._forward_as_attachments(selection)
            return

        # Single message: inline forward
        msg_data = getattr(self, '_current_message_full_data', None)
        if not msg_data:
            messagebox.showwarning("Forward", "Please wait for message to load")
            return

        # Get list data for sender info
        item_id = selection[0]
        list_data = self._messages_by_id.get(item_id, {})

        # Combine data for forward
        forward_data = {
            'from_name': list_data.get('sender_name', ''),
            'from_email': list_data.get('sender_email', ''),
            'to': msg_data.get('to', ''),
            'subject': list_data.get('subject', ''),
            'date': list_data.get('date_received', ''),
            'body': msg_data.get('body', ''),
            'original_id': item_id  # Track original message for marking as forwarded
        }

        # Check for attachments to include
        attachments = msg_data.get('attachments', [])
        if attachments:
            # Download attachments before showing compose
            self._forward_with_attachments(forward_data, attachments)
        else:
            self._show_compose(title="Forward", forward_data=forward_data)

    def _forward_with_attachments(self, forward_data, attachments):
        """Download attachments and then show forward compose."""
        import threading

        # Get session for downloading
        if not hasattr(self, '_current_account_id'):
            self._show_compose(title="Forward", forward_data=forward_data)
            return

        session = self.kerio_pool.get_session(self._current_account_id)
        if not session:
            self._show_compose(title="Forward", forward_data=forward_data)
            return

        self._forward_inline_pending = len(attachments)
        self._forward_inline_attachments = []
        self._forward_inline_data = forward_data

        self.statusbar.config(text=f"Downloading {len(attachments)} attachment(s)...")

        for att in attachments:
            url = att.get('url', '')
            name = att.get('name', 'attachment')
            if url:
                full_url = f"https://{session.config.server}{url}"
                threading.Thread(
                    target=self._download_forward_attachment,
                    args=(full_url, name, session),
                    daemon=True
                ).start()
            else:
                self._forward_inline_pending -= 1

        # Check if no attachments to download
        if self._forward_inline_pending <= 0:
            self._show_compose(title="Forward", forward_data=forward_data)

    def _download_forward_attachment(self, url, name, session):
        """Download a single attachment for forwarding."""
        try:
            import requests
            cookies = {session.cookie_name: session.cookie_value} if hasattr(session, 'cookie_name') else {}
            headers = {"X-Token": session.token} if hasattr(session, 'token') else {}

            response = requests.get(url, cookies=cookies, headers=headers, verify=False, timeout=60)
            response.raise_for_status()

            self._forward_inline_attachments.append({
                'name': name,
                'content': response.content
            })
        except Exception as e:
            print(f"Failed to download attachment {name}: {e}")

        self._forward_inline_pending -= 1

        # Check if all downloads complete
        if self._forward_inline_pending <= 0:
            self.root.after(0, self._forward_show_compose_with_attachments)

    def _forward_show_compose_with_attachments(self):
        """Show compose after attachments downloaded."""
        attachments = getattr(self, '_forward_inline_attachments', [])
        forward_data = getattr(self, '_forward_inline_data', {})
        count = len(attachments)
        title = f"Forward ({count} attachment{'s' if count != 1 else ''})" if count else "Forward"
        self._show_compose(title=title, forward_data=forward_data, attachments=attachments)

    def _forward_as_attachments(self, selection):
        """Forward multiple messages as attachments."""
        self._forward_pending = len(selection)
        self._forward_attachments = []
        self._forward_subjects = []

        self.statusbar.config(text=f"Fetching {len(selection)} messages...")

        for item_id in selection:
            list_data = self._messages_by_id.get(item_id, {})
            subject = list_data.get('subject', 'message')
            self._forward_subjects.append(subject)

            # Fetch raw message for attachment
            self.sync_manager.fetch_message_raw(
                self._current_account_id, item_id,
                callback=lambda success, data, subj=subject: self._on_forward_message_fetched(success, data, subj)
            )

    def _on_forward_message_fetched(self, success, data, subject):
        """Called when a message is fetched for forwarding."""
        if success and data:
            # Clean subject for filename
            import re
            safe_subject = re.sub(r'[<>:"/\\|?*]', '_', subject)[:50]
            filename = f"{safe_subject}.eml"
            self._forward_attachments.append({
                'name': filename,
                'content': data.encode('utf-8') if isinstance(data, str) else data
            })

        self._forward_pending -= 1

        if self._forward_pending <= 0:
            # All messages fetched, open compose
            attachments = self._forward_attachments
            count = len(attachments)
            self._show_compose(
                title=f"Forward ({count} attachments)",
                attachments=attachments
            )
            self.compose_subject_var.set(f"Fwd: {count} messages")

    def _delete_messages(self):
        """Delete selected message(s)."""
        selection = self.message_list.selection()
        if not selection:
            return

        # Check if we're in trash folder - if so, permanently delete
        in_trash = False
        if hasattr(self, '_current_folder_id') and hasattr(self, '_current_account_id'):
            folders = self.db.get_folders(self._current_account_id)
            for folder in folders:
                if folder['folder_id'] == self._current_folder_id:
                    in_trash = folder.get('folder_type') == 'trash'
                    break

        # Only confirm for permanent deletion from trash
        if in_trash:
            msg = f"Permanently delete {len(selection)} message(s)?" if len(selection) > 1 else "Permanently delete selected message?"
            if not messagebox.askyesno("Delete", msg):
                return

        # Track how many we're deleting
        self._delete_pending = len(selection)
        self._delete_success_count = 0
        self._delete_in_trash = in_trash

        # Delete each selected message
        if hasattr(self, '_current_account_id'):
            for item_id in selection:
                self.sync_manager.delete_message(
                    self._current_account_id, item_id, hard_delete=in_trash,
                    callback=lambda success, error, iid=item_id: self._on_message_deleted(iid, success, error)
                )

    def _on_message_deleted(self, item_id, success, error):
        """Called when a message is deleted."""
        if success:
            # Remove from message list and caches
            if self.message_list.exists(item_id):
                self.message_list.delete(item_id)
            if hasattr(self, '_messages_by_id') and item_id in self._messages_by_id:
                del self._messages_by_id[item_id]
            self._delete_success_count = getattr(self, '_delete_success_count', 0) + 1

        # Track completion
        self._delete_pending = getattr(self, '_delete_pending', 1) - 1

        if self._delete_pending <= 0:
            # All deletes completed
            count = getattr(self, '_delete_success_count', 0)
            in_trash = getattr(self, '_delete_in_trash', False)

            # Clear current message tracking
            self._current_message_id = None

            # Select next message or clear preview
            children = self.message_list.get_children()
            if children:
                self.message_list.selection_set(children[0])
                self.message_list.see(children[0])
                self._on_message_select(None)
            else:
                self._clear_preview_pane()

            # Update folder unread count
            self._update_folder_unread_count()

            # Show status
            if in_trash:
                self.statusbar.config(text=f"{count} message(s) permanently deleted")
            else:
                self.statusbar.config(text=f"{count} message(s) moved to Deleted Items")

    def _init_preview_background(self):
        """Initialize preview pane with theme-appropriate background."""
        if hasattr(self, 'preview_html') and self.preview_html:
            is_dark = self.dark_mode_var.get()
            bg = "#313335" if is_dark else "#ffffff"
            self.preview_html.load_html(f"<html><body style='background-color:{bg}'></body></html>")
        elif hasattr(self, 'preview_text'):
            self.preview_text.config(state=tk.NORMAL)
            self.preview_text.delete("1.0", tk.END)
            self.preview_text.config(state=tk.DISABLED)

    def _check_mail(self):
        """Check for new mail on all connected accounts."""
        if not self.connected_accounts:
            self.statusbar.config(text="No accounts connected")
            return

        self.statusbar.config(text="Checking mail...")
        for account_id in self.connected_accounts:
            self.sync_manager.sync_folders(account_id, callback=self._on_folders_synced)

    def _filter_messages(self):
        """Filter messages in list based on search query."""
        query = self.search_var.get().lower().strip()

        # Get all messages from cache
        if not hasattr(self, '_messages_by_id'):
            return

        # Temporarily unbind selection event
        self.message_list.unbind("<<TreeviewSelect>>")

        # Remember current selection
        current_selection = self.message_list.selection()

        # Clear the list
        self.message_list.delete(*self.message_list.get_children())

        # Collect matching messages
        matching = []
        for item_id, msg in self._messages_by_id.items():
            # Check if message matches filter
            if query:
                from_str = f"{msg.get('sender_name', '')} {msg.get('sender_email', '')}".lower()
                subject = msg.get('subject', '').lower()
                if query not in from_str and query not in subject:
                    continue
            matching.append((item_id, msg))

        # Sort by date (newest first) - date format is YYYYMMDDTHHMMSS
        matching.sort(key=lambda x: x[1].get('date_received', ''), reverse=True)

        # Add to list in sorted order
        for item_id, msg in matching:
            date_str = msg.get('date_received', '')
            if len(date_str) >= 15:
                date_str = f"{date_str[0:4]}/{date_str[4:6]}/{date_str[6:8]} {date_str[9:11]}:{date_str[11:13]}"

            subject = msg.get('subject', '')
            flag = UNI_FLAG if msg.get('is_flagged') else ""
            # Status: replied takes precedence over forwarded
            if msg.get('is_answered'):
                status = UNI_REPLY
            elif msg.get('is_forwarded'):
                status = UNI_FORWARD
            else:
                status = ""
            attach = (FA_PAPERCLIP if self._has_fontawesome else "*") if msg.get('has_attachments') else ""

            sender = msg.get('sender_name') or msg.get('sender_email', '')
            values = (
                flag,
                status,
                attach,
                sender,
                subject,
                date_str
            )
            self.message_list.insert("", tk.END, iid=item_id, values=values)

        # Rebind selection event
        self.message_list.bind("<<TreeviewSelect>>", self._on_message_select)

        # Always select and display the first message
        children = self.message_list.get_children()
        if children:
            self.message_list.selection_set(children[0])
            self.message_list.see(children[0])
            self._on_message_select(None)  # Load preview

        # Update status
        if query:
            self.statusbar.config(text=f"Filtered: {len(matching)} messages matching '{query}'")

    def _filter_to_first_message(self, event=None):
        """Navigate from filter box to first message in list."""
        children = self.message_list.get_children()
        if children:
            self.message_list.focus_set()
            self.message_list.selection_set(children[0])
            self.message_list.focus(children[0])
            self.message_list.see(children[0])
        return "break"

    def _get_default_account_id(self):
        """Get the default account ID for sending."""
        accounts = self.db.get_accounts()
        for acc in accounts:
            if acc.get('is_default'):
                return acc['id']
        return accounts[0]['id'] if accounts else None

    # ==================== View Switching ====================

    def _show_view(self, view_type):
        """Switch between mail/calendar/contacts views."""
        # Currently only mail is implemented
        pass

    # ==================== Settings & Theme ====================

    def _show_settings(self):
        """Show settings dialog."""
        from mailbench.dialogs.settings_dialog import SettingsDialog
        SettingsDialog(self.root, self.db, self)

    def _show_about(self):
        """Show about dialog."""
        messagebox.showinfo(
            "About Mailbench",
            f"Mailbench v{__version__}\n\n"
            "A Python email client for Kerio Connect."
        )

    def _toggle_dark_mode(self):
        """Toggle dark mode on/off."""
        is_dark = self.dark_mode_var.get()
        self.db.set_setting("dark_mode", "1" if is_dark else "0")
        self._apply_theme()

    def _apply_theme(self):
        """Apply light or dark theme."""
        is_dark = self.dark_mode_var.get()
        style = ttk.Style()

        if is_dark:
            bg = "#2b2b2b"
            fg = "#a9b7c6"
            bg_light = "#313335"
            bg_dark = "#1e1e1e"
            select_bg = "#214283"
            border = "#3c3f41"

            style.theme_use("clam")

            style.configure(".", background=bg, foreground=fg, fieldbackground=bg_light,
                           troughcolor=bg_dark, bordercolor=border, lightcolor=bg_light,
                           darkcolor=bg_dark, insertcolor=fg)
            style.configure("TFrame", background=bg)
            style.configure("TLabel", background=bg, foreground=fg)
            style.configure("TLabelframe", background=bg, foreground=fg)
            style.configure("TLabelframe.Label", background=bg, foreground=fg)
            style.configure("TButton", background=bg_light, foreground=fg)
            style.configure("TEntry", fieldbackground=bg_light, foreground=fg, insertcolor=fg)
            style.configure("TCombobox", fieldbackground=bg_light, foreground=fg)
            style.map("TCombobox", fieldbackground=[("readonly", bg_light)])
            # Combobox dropdown list colors
            self.root.option_add("*TCombobox*Listbox.background", bg_light)
            self.root.option_add("*TCombobox*Listbox.foreground", fg)
            self.root.option_add("*TCombobox*Listbox.selectBackground", select_bg)
            self.root.option_add("*TCombobox*Listbox.selectForeground", fg)
            style.configure("TPanedwindow", background=bg)
            style.configure("Sash", sashthickness=6, gripcount=0)
            style.configure("TScrollbar", background="#5a5a5a", troughcolor=bg, arrowcolor="#6e6e6e",
                           bordercolor=bg, lightcolor="#5a5a5a", darkcolor="#5a5a5a")
            style.map("TScrollbar", background=[("pressed", "#6e6e6e"), ("active", "#6e6e6e")])
            style.configure("Treeview", background=bg_light, foreground=fg, fieldbackground=bg_light)
            style.configure("Treeview.Heading", background=bg, foreground=fg)
            style.configure("TCheckbutton", background=bg, foreground=fg)

            style.map("TButton", background=[("active", bg_light)])
            style.map("Treeview", background=[("selected", select_bg)], foreground=[("selected", fg)])

            self.root.configure(bg=bg)
        else:
            style.theme_use("clam")

            bg = "#d9d9d9"
            fg = "#000000"
            bg_light = "#ffffff"
            select_bg = "#4a6984"
            border = "#9e9e9e"

            style.configure(".", background=bg, foreground=fg, fieldbackground=bg_light,
                           troughcolor="#c3c3c3", bordercolor=border,
                           lightcolor="#ededed", darkcolor="#cfcfcf", insertcolor=fg)
            style.configure("TFrame", background=bg)
            style.configure("TLabel", background=bg, foreground=fg)
            style.configure("TLabelframe", background=bg, foreground=fg)
            style.configure("TLabelframe.Label", background=bg, foreground=fg)
            style.configure("TButton", background="#e1e1e1", foreground=fg)
            style.configure("TEntry", fieldbackground=bg_light, foreground=fg, insertcolor=fg)
            style.configure("TCombobox", fieldbackground=bg_light, foreground=fg)
            style.map("TCombobox", fieldbackground=[("readonly", bg_light)])
            # Combobox dropdown list colors
            self.root.option_add("*TCombobox*Listbox.background", bg_light)
            self.root.option_add("*TCombobox*Listbox.foreground", fg)
            self.root.option_add("*TCombobox*Listbox.selectBackground", select_bg)
            self.root.option_add("*TCombobox*Listbox.selectForeground", "#ffffff")
            style.configure("TPanedwindow", background=bg)
            style.configure("TScrollbar", background="#c3c3c3", troughcolor="#e6e6e6",
                           arrowcolor="#5a5a5a", bordercolor=border)
            style.map("TScrollbar", background=[("pressed", "#a0a0a0"), ("active", "#b0b0b0")])
            style.configure("Treeview", background=bg_light, foreground=fg, fieldbackground=bg_light)
            style.configure("Treeview.Heading", background=bg, foreground=fg)
            style.configure("TCheckbutton", background=bg, foreground=fg)

            style.map("TButton", background=[("active", "#ececec")])
            style.map("Treeview", background=[("selected", select_bg)], foreground=[("selected", "#ffffff")])

            self.root.configure(bg=bg)

        self._apply_theme_to_widgets()

    def _apply_theme_to_widgets(self):
        """Apply theme to non-ttk widgets."""
        is_dark = self.dark_mode_var.get()

        if is_dark:
            bg = "#2b2b2b"
            fg = "#a9b7c6"
            text_bg = "#313335"
            select_bg = "#214283"
        else:
            bg = "#f0f0f0"
            fg = "#000000"
            text_bg = "#ffffff"
            select_bg = "#0078d4"

        self._configure_widgets_recursive(self.root, text_bg, fg, select_bg, bg)

    def _configure_widgets_recursive(self, widget, bg, fg, select_bg, menu_bg):
        """Recursively configure non-ttk widgets."""
        widget_class = widget.winfo_class()

        try:
            if widget_class == "Text":
                widget.configure(bg=bg, fg=fg, insertbackground=fg,
                               selectbackground=select_bg, selectforeground=fg)
            elif widget_class == "Listbox":
                widget.configure(bg=bg, fg=fg,
                               selectbackground=select_bg, selectforeground=fg)
            elif widget_class == "Menu":
                widget.configure(bg=menu_bg, fg=fg,
                               activebackground=select_bg, activeforeground=fg)
        except tk.TclError:
            pass

        for child in widget.winfo_children():
            self._configure_widgets_recursive(child, bg, fg, select_bg, menu_bg)

    def _apply_font_size(self):
        """Apply the current font size to all UI elements."""
        from tkinter import font as tkfont

        size = self.font_size

        # Configure the default fonts globally
        default_font = tkfont.nametofont("TkDefaultFont")
        default_font.configure(size=size)

        menu_font = tkfont.nametofont("TkMenuFont")
        menu_font.configure(size=size)

        text_font = tkfont.nametofont("TkTextFont")
        text_font.configure(size=size)

        fixed_font = tkfont.nametofont("TkFixedFont")
        fixed_font.configure(size=size)

        # Update folder frame "Accounts" label
        for child in self.folder_frame.winfo_children():
            if isinstance(child, ttk.Label):
                child.configure(font=("TkDefaultFont", size - 1, "bold"))
                break

        # Update preview header labels
        if hasattr(self, 'preview_from'):
            self.preview_from.configure(font=("TkDefaultFont", size - 1, "bold"))
        if hasattr(self, 'preview_subject'):
            self.preview_subject.configure(font=("TkDefaultFont", size, "bold"))

        # Update compose labels
        if hasattr(self, 'compose_from_label'):
            self.compose_from_label.configure(font=("TkDefaultFont", size, "bold"))
        if hasattr(self, 'compose_to_label'):
            self.compose_to_label.configure(font=("TkDefaultFont", size, "bold"))
        if hasattr(self, 'compose_cc_label'):
            self.compose_cc_label.configure(font=("TkDefaultFont", size, "bold"))
        if hasattr(self, 'compose_subject_label'):
            self.compose_subject_label.configure(font=("TkDefaultFont", size, "bold"))

        # Update compose body font (with zoom)
        if hasattr(self, 'compose_body'):
            zoom_factor = getattr(self, '_preview_zoom', 100) / 100
            zoomed_size = int(size * zoom_factor)
            self.compose_body.configure(font=("TkDefaultFont", zoomed_size))
            self._setup_compose_tags()  # Re-setup tags with new size and zoom

    # ==================== Window Management ====================

    def _restore_geometry(self):
        """Restore window geometry from saved settings."""
        default_geometry = "1200x800"
        saved = self.db.get_setting("window_geometry", default_geometry)

        try:
            self.root.geometry(saved)
        except Exception:
            self.root.geometry(default_geometry)

    def _ensure_visible_on_screen(self):
        """Ensure window is visible on screen."""
        default_geometry = "1200x800"
        try:
            if not self._is_visible_on_screen():
                self.root.geometry(default_geometry)
                self._center_window()
        except Exception:
            pass

    def _is_visible_on_screen(self):
        """Check if window position is reasonable."""
        x = self.root.winfo_x()
        y = self.root.winfo_y()
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        max_x = screen_w * 3
        max_y = screen_h * 2
        return x > -500 and y > -100 and x < max_x and y < max_y

    def _center_window(self):
        """Center window on screen."""
        self.root.update_idletasks()
        w = self.root.winfo_width()
        h = self.root.winfo_height()
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        x = (screen_w - w) // 2
        y = (screen_h - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")

    def _save_geometry(self):
        """Save current window geometry."""
        geometry = self.root.geometry()
        self.db.set_setting("window_geometry", geometry)

    def _save_layout(self):
        """Save paned window sash positions as ratios."""
        try:
            sash_pos = self.main_paned.sashpos(0)
            pane_width = self.main_paned.winfo_width()
            if pane_width > 100:
                ratio = sash_pos / pane_width
                self.db.set_setting("layout_folder_ratio", f"{ratio:.4f}")
        except Exception:
            pass

        try:
            sash_pos = self.content_paned.sashpos(0)
            pane_width = self.content_paned.winfo_width()
            if pane_width > 100:
                ratio = sash_pos / pane_width
                self.db.set_setting("layout_list_ratio", f"{ratio:.4f}")
        except Exception:
            pass

    def _restore_layout(self):
        """Restore paned window sash positions from saved ratios."""
        self.root.update_idletasks()

        try:
            ratio_str = self.db.get_setting("layout_folder_ratio")
            if ratio_str:
                ratio = float(ratio_str)
                pane_width = self.main_paned.winfo_width()
                if pane_width > 100 and 0.05 <= ratio <= 0.95:
                    sash_pos = int(ratio * pane_width)
                    self.main_paned.sashpos(0, sash_pos)
        except Exception:
            pass

        try:
            ratio_str = self.db.get_setting("layout_list_ratio")
            if ratio_str:
                ratio = float(ratio_str)
                pane_width = self.content_paned.winfo_width()
                if pane_width > 100 and 0.05 <= ratio <= 0.95:
                    sash_pos = int(ratio * pane_width)
                    self.content_paned.sashpos(0, sash_pos)
        except Exception:
            pass

        self.root.after(50, self.root.deiconify)

    def _reset_layout(self):
        """Reset layout to defaults."""
        self.db.set_setting("layout_folder_ratio", "")
        self.db.set_setting("layout_list_ratio", "")

        self.root.update_idletasks()
        try:
            self.main_paned.sashpos(0, 200)
            width = self.content_paned.winfo_width()
            self.content_paned.sashpos(0, width // 2)
        except Exception:
            pass

        self.statusbar.config(text="Layout reset to defaults")

    def _restore_session(self):
        """Restore session state."""
        # Auto-connect all accounts
        accounts = self.db.get_accounts()
        for account in accounts:
            self._connect_account(account['id'])

        # Update status bar with connected count
        self._update_status()

        # Restore layout
        self.root.after(100, self._restore_layout)

        # Note: INBOX auto-selection happens in _on_folders_synced after connection
        # Note: New messages are detected via change listener (push notifications)

    def _auto_select_inbox(self):
        """Auto-select INBOX folder and load messages."""
        if not self.connected_accounts:
            return

        # Find first connected account
        account_id = next(iter(self.connected_accounts))
        folders = self.db.get_folders(account_id)

        # Find INBOX folder
        inbox_folder = None
        for folder in folders:
            if folder.get('folder_type') == 'inbox':
                inbox_folder = folder
                break

        if inbox_folder:
            folder_id = inbox_folder['folder_id']
            tree_id = f"folder_{account_id}_{folder_id}"

            # Select in tree
            if self.folder_tree.exists(tree_id):
                self.folder_tree.selection_set(tree_id)
                self.folder_tree.see(tree_id)

            # Load messages and select first
            self._load_messages(account_id, folder_id, select_first=True)

    def _on_auto_refresh_messages(self, success, error=None, messages_data=None):
        """Called when auto-refresh message sync completes."""
        if success:
            # Incremental update - no flashing
            self._incremental_update_messages(messages_data)

    def _incremental_update_messages(self, messages_data=None):
        """Incrementally update message list without flashing."""
        if not hasattr(self, '_current_account_id') or not hasattr(self, '_current_folder_id'):
            return

        # Use provided data or fall back to cache
        if not messages_data:
            messages_data = self.db.get_messages(
                self._current_account_id, self._current_folder_id, limit=50
            )

        # Save current selection
        selection = self.message_list.selection()

        # Get existing item_ids in treeview
        existing_ids = set(self.message_list.get_children())
        new_ids = {msg['item_id'] for msg in messages_data if msg.get('item_id')}

        # Check if filter is active
        filter_query = self.search_var.get().lower().strip() if hasattr(self, 'search_var') else ""

        # Remove messages no longer on server (quietly, no rebuild)
        deleted_ids = existing_ids - new_ids
        for item_id in deleted_ids:
            self.message_list.delete(item_id)
            # Also remove from cache
            if hasattr(self, '_messages_by_id') and item_id in self._messages_by_id:
                del self._messages_by_id[item_id]

        # Ensure _messages_by_id exists
        if not hasattr(self, '_messages_by_id'):
            self._messages_by_id = {}

        # Update existing and add new messages
        for msg in messages_data:
            item_id = msg.get('item_id')
            if not item_id:
                continue

            # Always update message cache (for unread counts, etc.)
            self._messages_by_id[item_id] = msg

            # Check if message matches filter
            if filter_query:
                from_str = f"{msg.get('sender_name', '')} {msg.get('sender_email', '')}".lower()
                subject_str = msg.get('subject', '').lower()
                if filter_query not in from_str and filter_query not in subject_str:
                    # Message doesn't match filter - skip treeview update but keep in cache
                    continue

            sender = msg.get('sender_name') or msg.get('sender_email') or 'Unknown'
            subject = msg.get('subject') or '(No Subject)'
            flag = UNI_FLAG if msg.get('is_flagged') else ""
            # Status: replied takes precedence over forwarded
            if msg.get('is_answered'):
                status = UNI_REPLY
            elif msg.get('is_forwarded'):
                status = UNI_FORWARD
            else:
                status = ""
            attach = (FA_PAPERCLIP if self._has_fontawesome else "*") if msg.get('has_attachments') else ""
            date_raw = msg.get('date_received', '')
            date = self._format_kerio_date(date_raw) if date_raw else ""

            tags = () if msg.get('is_read') else ('unread',)

            if item_id in existing_ids:
                # Update existing item (for read/unread changes)
                self.message_list.item(item_id, values=(flag, status, attach, sender, subject, date), tags=tags)
            else:
                # Find correct position to insert (maintain date sort, newest first)
                insert_pos = 0
                children = self.message_list.get_children()
                for i, child_id in enumerate(children):
                    child_msg = self._messages_by_id.get(child_id)
                    if child_msg:
                        child_date = child_msg.get('date_received', '')
                        if date_raw >= child_date:
                            # New message is newer or same, insert here
                            insert_pos = i
                            break
                        insert_pos = i + 1
                self.message_list.insert("", insert_pos, iid=item_id,
                                        values=(flag, status, attach, sender, subject, date), tags=tags)

        # Handle selection after deletions
        if selection:
            selected_item = selection[0]
            if selected_item in deleted_ids:
                # Selected message was deleted - clear current message and select next
                self._current_message_id = None
                children = self.message_list.get_children()
                if children:
                    # Select first available message
                    self.message_list.selection_set(children[0])
                    self.message_list.see(children[0])
                    self._on_message_select(None)
                else:
                    # No messages left - clear preview
                    self._clear_preview_pane()
            else:
                # Selection still exists - restore it without triggering reload
                self.message_list.unbind("<<TreeviewSelect>>")
                self.message_list.selection_set(selected_item)
                self.message_list.bind("<<TreeviewSelect>>", self._on_message_select)

        # Update folder unread count
        self._update_folder_unread_count()

    def _build_email_cache_background(self):
        """Build email cache from existing messages in background thread."""
        import threading

        def build_cache():
            try:
                # Get unique senders from messages table
                senders = self.db.get_unique_senders_from_messages()
                if senders:
                    # Convert to list of tuples
                    emails = [(s.get('sender_email'), s.get('sender_name')) for s in senders]
                    self.db.bulk_add_emails_to_cache(emails)
            except Exception:
                pass  # Silently fail - cache is nice-to-have

        thread = threading.Thread(target=build_cache, daemon=True)
        thread.start()

    def _update_status(self):
        """Update status bar with connection info."""
        count = len(self.connected_accounts)
        if count == 0:
            self.statusbar.config(text="No accounts connected")
        elif count == 1:
            self.statusbar.config(text="1 account connected")
        else:
            self.statusbar.config(text=f"{count} accounts connected")

    def _check_for_updates(self):
        """Check for updates in background."""
        from mailbench.version import check_for_updates

        def on_update_check(has_update, latest_version):
            if has_update:
                self.root.after(0, lambda: self._show_update_dialog(latest_version))

        check_for_updates(on_update_check)

    def _show_update_dialog(self, latest_version):
        """Show update available dialog."""
        result = messagebox.askyesno(
            "Update Available",
            f"A new version of Mailbench is available.\n\n"
            f"Installed: {__version__}\n"
            f"Latest: {latest_version}\n\n"
            f"Would you like to upgrade now?\n\n"
            f"(This will run: pipx upgrade mailbench)"
        )

        if result:
            self._run_upgrade()

    def _run_upgrade(self):
        """Run pipx upgrade in background."""
        import subprocess
        import threading

        def do_upgrade():
            try:
                result = subprocess.run(
                    ["pipx", "upgrade", "mailbench"],
                    capture_output=True,
                    text=True
                )
                if result.returncode == 0:
                    self.root.after(0, lambda: messagebox.showinfo(
                        "Upgrade Complete",
                        "Mailbench has been upgraded. Please restart the application."
                    ))
                else:
                    error = result.stderr or result.stdout or "Unknown error"
                    self.root.after(0, lambda: messagebox.showerror("Upgrade Failed", error))
            except FileNotFoundError:
                self.root.after(0, lambda: messagebox.showerror(
                    "Error",
                    "pipx not found. Please upgrade manually:\n\npipx upgrade mailbench"
                ))

        self.statusbar.config(text="Upgrading Mailbench...")
        thread = threading.Thread(target=do_upgrade, daemon=True)
        thread.start()

    def _on_close(self):
        """Handle window close event."""
        # Stop auto-refresh
        self._closing = True

        self._save_geometry()
        self._save_layout()

        # Save view state
        if hasattr(self, '_current_account_id') and hasattr(self, '_current_folder_id'):
            selection = self.message_list.selection()
            selected_msg = selection[0].replace("msg_", "") if selection else None
            self.db.save_view_state("mail", self._current_account_id, self._current_folder_id,
                                    selected_msg)

        # Disconnect all accounts
        self.kerio_pool.disconnect_all()
        self.sync_manager.shutdown()

        self.root.destroy()

        # Force exit in case background threads are still running
        import sys
        sys.exit(0)

    def run(self):
        self.root.mainloop()


def main():
    """Entry point for Mailbench."""
    app = MailbenchApp()
    app.run()


if __name__ == "__main__":
    main()
