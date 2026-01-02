"""Settings dialog."""

import os
import tkinter as tk
from tkinter import ttk, filedialog


class SettingsDialog:
    def __init__(self, parent, db, app):
        self.db = db
        self.app = app

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Settings")
        self.dialog.geometry("500x250")
        self.dialog.transient(parent)
        self.dialog.grab_set()

        self._create_ui()

        # Center on parent
        self.dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.dialog.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.dialog.winfo_height()) // 2
        self.dialog.geometry(f"+{x}+{y}")

    def _create_ui(self):
        main_frame = ttk.Frame(self.dialog, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Configure column weights for proper expansion
        main_frame.columnconfigure(1, weight=1)

        # Font size
        row = 0
        ttk.Label(main_frame, text="Font Size:").grid(row=row, column=0, sticky=tk.W, pady=5)
        self.font_size_var = tk.IntVar(value=self.app.font_size)
        font_spinbox = ttk.Spinbox(main_frame, from_=6, to=24, width=5,
                                   textvariable=self.font_size_var)
        font_spinbox.grid(row=row, column=1, sticky=tk.W, pady=5)
        row += 1

        # Default save directory
        ttk.Label(main_frame, text="Default Save Directory:").grid(row=row, column=0, sticky=tk.W, pady=5)
        save_dir_frame = ttk.Frame(main_frame)
        save_dir_frame.grid(row=row, column=1, sticky=tk.EW, pady=5)

        # Get current setting or default to Downloads
        current_save_dir = self.db.get_setting("default_save_directory", "")
        if not current_save_dir:
            current_save_dir = os.path.join(os.path.expanduser("~"), "Downloads")
        self.save_dir_var = tk.StringVar(value=current_save_dir)

        save_dir_entry = ttk.Entry(save_dir_frame, textvariable=self.save_dir_var)
        save_dir_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(save_dir_frame, text="Browse...", command=self._browse_save_dir).pack(side=tk.LEFT)
        row += 1

        # Desktop launcher
        ttk.Label(main_frame, text="Desktop Launcher:").grid(row=row, column=0, sticky=tk.W, pady=5)
        launcher_frame = ttk.Frame(main_frame)
        launcher_frame.grid(row=row, column=1, sticky=tk.W, pady=5)
        ttk.Button(launcher_frame, text="Install", command=self._install_launcher).pack(side=tk.LEFT, padx=2)
        ttk.Button(launcher_frame, text="Remove", command=self._remove_launcher).pack(side=tk.LEFT, padx=2)
        row += 1

        # Buttons
        btn_frame = ttk.Frame(main_frame)
        btn_frame.grid(row=row, column=0, columnspan=2, pady=(20, 0))

        ttk.Button(btn_frame, text="OK", command=self._save).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", command=self.dialog.destroy).pack(side=tk.LEFT, padx=5)

    def _browse_save_dir(self):
        """Browse for default save directory."""
        initial_dir = self.save_dir_var.get()
        if not os.path.isdir(initial_dir):
            initial_dir = os.path.expanduser("~")
        new_dir = filedialog.askdirectory(initialdir=initial_dir, title="Select Default Save Directory")
        if new_dir:
            self.save_dir_var.set(new_dir)

    def _save(self):
        """Save settings."""
        font_size = self.font_size_var.get()
        self.db.set_setting("font_size", str(font_size))
        self.app.font_size = font_size
        self.app._apply_font_size()

        # Save default save directory
        save_dir = self.save_dir_var.get().strip()
        if save_dir and os.path.isdir(save_dir):
            self.db.set_setting("default_save_directory", save_dir)

        self.dialog.destroy()

    def _install_launcher(self):
        """Install desktop launcher."""
        from mailbench.launcher import create_launcher
        success = create_launcher()
        if success:
            tk.messagebox.showinfo("Success", "Desktop launcher installed")
        else:
            tk.messagebox.showerror("Error", "Failed to install launcher")

    def _remove_launcher(self):
        """Remove desktop launcher."""
        from mailbench.launcher import remove_launcher
        success = remove_launcher()
        if success:
            tk.messagebox.showinfo("Success", "Desktop launcher removed")
        else:
            tk.messagebox.showerror("Error", "No launcher found to remove")
