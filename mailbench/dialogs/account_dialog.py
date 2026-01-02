"""Account management dialog."""

import tkinter as tk
from tkinter import ttk, messagebox

from mailbench.kerio_client import KerioConfig, KerioSession


class AccountDialog:
    def __init__(self, parent, db, kerio_pool, app=None, edit_account_id=None):
        self.db = db
        self.kerio_pool = kerio_pool
        self.app = app
        self.edit_account_id = edit_account_id

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Manage Accounts" if not edit_account_id else "Edit Account")
        self.dialog.geometry("800x400")
        self.dialog.transient(parent)
        self.dialog.grab_set()

        self._create_ui()
        self._load_accounts()

        # Center on parent
        self.dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.dialog.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.dialog.winfo_height()) // 2
        self.dialog.geometry(f"+{x}+{y}")

    def _create_ui(self):
        # Main container
        main_frame = ttk.Frame(self.dialog, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Left: Account list
        left_frame = ttk.LabelFrame(main_frame, text="Accounts", padding=5)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=False, padx=(0, 10))

        self.account_list = tk.Listbox(left_frame, width=25, height=15)
        account_scroll = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.account_list.yview)
        self.account_list.configure(yscrollcommand=account_scroll.set)

        self.account_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        account_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.account_list.bind("<<ListboxSelect>>", self._on_account_select)

        # List buttons
        list_btn_frame = ttk.Frame(left_frame)
        list_btn_frame.pack(fill=tk.X, pady=(5, 0))
        ttk.Button(list_btn_frame, text="New", command=self._new_account).pack(side=tk.LEFT, padx=2)
        ttk.Button(list_btn_frame, text="Delete", command=self._delete_account).pack(side=tk.LEFT, padx=2)

        # Right: Account details
        right_frame = ttk.LabelFrame(main_frame, text="Account Details", padding=10)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Form fields
        row = 0

        ttk.Label(right_frame, text="Account Name:").grid(row=row, column=0, sticky=tk.W, pady=2, padx=(0,10))
        self.name_var = tk.StringVar()
        self.name_entry = ttk.Entry(right_frame, textvariable=self.name_var, width=50)
        self.name_entry.grid(row=row, column=1, sticky=tk.EW, pady=2)
        row += 1

        ttk.Label(right_frame, text="Email Address:").grid(row=row, column=0, sticky=tk.W, pady=2, padx=(0,10))
        self.email_var = tk.StringVar()
        self.email_entry = ttk.Entry(right_frame, textvariable=self.email_var, width=50)
        self.email_entry.grid(row=row, column=1, sticky=tk.EW, pady=2)
        row += 1

        ttk.Label(right_frame, text="Server:").grid(row=row, column=0, sticky=tk.W, pady=2, padx=(0,10))
        self.server_var = tk.StringVar()
        self.server_entry = ttk.Entry(right_frame, textvariable=self.server_var, width=50)
        self.server_entry.grid(row=row, column=1, sticky=tk.EW, pady=2)
        row += 1

        # Hint for server
        hint_label = ttk.Label(right_frame, text="e.g., mail.company.com (Kerio Connect server)",
                              foreground="gray", font=("TkDefaultFont", 8))
        hint_label.grid(row=row, column=1, sticky=tk.W)
        row += 1

        ttk.Label(right_frame, text="Username:").grid(row=row, column=0, sticky=tk.W, pady=2, padx=(0,10))
        self.username_var = tk.StringVar()
        self.username_entry = ttk.Entry(right_frame, textvariable=self.username_var, width=50)
        self.username_entry.grid(row=row, column=1, sticky=tk.EW, pady=2)
        row += 1

        ttk.Label(right_frame, text="Password:").grid(row=row, column=0, sticky=tk.W, pady=2, padx=(0,10))
        self.password_var = tk.StringVar()
        self.password_entry = ttk.Entry(right_frame, textvariable=self.password_var, width=50, show="*")
        self.password_entry.grid(row=row, column=1, sticky=tk.EW, pady=2)
        row += 1

        # Make column 1 expand
        right_frame.columnconfigure(1, weight=1)

        # Checkboxes
        self.default_var = tk.BooleanVar(value=False)
        default_cb = ttk.Checkbutton(right_frame, text="Default account for sending",
                                     variable=self.default_var)
        default_cb.grid(row=row, column=0, columnspan=2, sticky=tk.W, pady=2)
        row += 1

        # Buttons
        btn_frame = ttk.Frame(right_frame)
        btn_frame.grid(row=row, column=0, columnspan=2, pady=(20, 0))

        ttk.Button(btn_frame, text="Test Connection", command=self._test_connection).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Save", command=self._save_account).pack(side=tk.LEFT, padx=5)

        # Bottom buttons
        bottom_frame = ttk.Frame(self.dialog)
        bottom_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Button(bottom_frame, text="Close", command=self.dialog.destroy).pack(side=tk.RIGHT)

        # Status label
        self.status_label = ttk.Label(bottom_frame, text="")
        self.status_label.pack(side=tk.LEFT)

    def _load_accounts(self):
        """Load accounts into the list."""
        self.account_list.delete(0, tk.END)
        accounts = self.db.get_accounts()
        self._accounts = accounts

        for acc in accounts:
            display = acc['name']
            if acc.get('is_default'):
                display = f"{display} (default)"
            self.account_list.insert(tk.END, display)

        if accounts:
            self.account_list.selection_set(0)
            self._on_account_select(None)

    def _on_account_select(self, event):
        """Handle account selection."""
        selection = self.account_list.curselection()
        if not selection:
            return

        idx = selection[0]
        if idx < len(self._accounts):
            acc = self.db.get_account(self._accounts[idx]['id'])
            if acc:
                self.name_var.set(acc['name'])
                self.email_var.set(acc['email'])
                self.server_var.set(acc['server'])
                self.username_var.set(acc['username'])
                self.password_var.set(acc['password'])
                self.default_var.set(bool(acc.get('is_default')))
                self._current_account_id = acc['id']
            else:
                self._current_account_id = None

    def _new_account(self):
        """Clear form for new account."""
        self.account_list.selection_clear(0, tk.END)
        self.name_var.set("")
        self.email_var.set("")
        self.server_var.set("")
        self.username_var.set("")
        self.password_var.set("")
        self.default_var.set(False)
        self._current_account_id = None
        self.name_entry.focus()

    def _delete_account(self):
        """Delete selected account."""
        selection = self.account_list.curselection()
        if not selection:
            return

        idx = selection[0]
        if idx >= len(self._accounts):
            return

        acc = self._accounts[idx]
        if not messagebox.askyesno("Delete Account",
                                   f"Delete account '{acc['name']}'?\n\n"
                                   "This will remove all cached messages for this account."):
            return

        self.db.delete_account(acc['id'])
        self._load_accounts()
        self._new_account()
        self.status_label.config(text="Account deleted")

    def _test_connection(self):
        """Test the connection settings."""
        if not self._validate():
            return

        self.status_label.config(text="Testing connection...")
        self.dialog.update()

        config = KerioConfig(
            email=self.email_var.get().strip(),
            username=self.username_var.get().strip(),
            password=self.password_var.get(),
            server=self.server_var.get().strip()
        )

        try:
            session = KerioSession(config)
            session.login()
            user_info = session.whoami()
            session.logout()
            self.status_label.config(
                text=f"Connected as {user_info.get('userName', 'user')}",
                foreground="green"
            )
        except Exception as e:
            self.status_label.config(text=f"Failed: {str(e)}", foreground="red")

    def _validate(self):
        """Validate form fields."""
        if not self.name_var.get().strip():
            messagebox.showerror("Validation Error", "Account name is required")
            return False
        if not self.email_var.get().strip():
            messagebox.showerror("Validation Error", "Email address is required")
            return False
        if not self.server_var.get().strip():
            messagebox.showerror("Validation Error", "Server is required")
            return False
        if not self.username_var.get().strip():
            messagebox.showerror("Validation Error", "Username is required")
            return False
        if not self.password_var.get():
            messagebox.showerror("Validation Error", "Password is required")
            return False
        return True

    def _save_account(self):
        """Save the account."""
        if not self._validate():
            return

        name = self.name_var.get().strip()
        email = self.email_var.get().strip()
        server = self.server_var.get().strip()
        username = self.username_var.get().strip()
        password = self.password_var.get()
        is_default = self.default_var.get()

        # Check for duplicate name
        existing = self.db.get_account_by_name(name)
        if existing and (not hasattr(self, '_current_account_id') or
                        existing['id'] != self._current_account_id):
            messagebox.showerror("Error", f"An account named '{name}' already exists")
            return

        account_id = getattr(self, '_current_account_id', None)

        self.db.save_account(
            name=name,
            email=email,
            server=server,
            username=username,
            password=password,
            is_default=is_default,
            account_id=account_id
        )

        self._load_accounts()
        self.status_label.config(text="Account saved", foreground="green")
