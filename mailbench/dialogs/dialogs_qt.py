"""Qt dialogs for account management and settings."""

import os
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QLineEdit,
    QPushButton, QListWidget, QGroupBox, QCheckBox, QSpinBox, QFileDialog,
    QMessageBox, QDialogButtonBox, QComboBox
)
from PySide6.QtCore import Qt

from mailbench.kerio_client import KerioConfig, KerioSession


class AccountDialog(QDialog):
    """Account management dialog."""

    def __init__(self, parent, db, kerio_pool, app=None, edit_account_id=None):
        super().__init__(parent)
        self.db = db
        self.kerio_pool = kerio_pool
        self.app = app
        self.edit_account_id = edit_account_id
        self._current_account_id = None
        self._accounts = []

        self.setWindowTitle("Manage Accounts" if not edit_account_id else "Edit Account")
        self.setMinimumSize(800, 400)
        self.setModal(True)

        self._create_ui()
        self._load_accounts()

    def _create_ui(self):
        main_layout = QVBoxLayout(self)

        # Content layout (horizontal)
        content_layout = QHBoxLayout()
        main_layout.addLayout(content_layout, 1)

        # Left: Account list
        left_group = QGroupBox("Accounts")
        left_layout = QVBoxLayout(left_group)

        self.account_list = QListWidget()
        self.account_list.itemSelectionChanged.connect(self._on_account_select)
        left_layout.addWidget(self.account_list)

        btn_layout = QHBoxLayout()
        new_btn = QPushButton("New")
        new_btn.clicked.connect(self._new_account)
        delete_btn = QPushButton("Delete")
        delete_btn.clicked.connect(self._delete_account)
        btn_layout.addWidget(new_btn)
        btn_layout.addWidget(delete_btn)
        left_layout.addLayout(btn_layout)

        content_layout.addWidget(left_group)

        # Right: Account details
        right_group = QGroupBox("Account Details")
        right_layout = QGridLayout(right_group)

        row = 0
        right_layout.addWidget(QLabel("Account Name:"), row, 0)
        self.name_edit = QLineEdit()
        right_layout.addWidget(self.name_edit, row, 1)
        row += 1

        right_layout.addWidget(QLabel("Email Address:"), row, 0)
        self.email_edit = QLineEdit()
        right_layout.addWidget(self.email_edit, row, 1)
        row += 1

        right_layout.addWidget(QLabel("Server:"), row, 0)
        self.server_edit = QLineEdit()
        self.server_edit.setPlaceholderText("e.g., mail.company.com")
        right_layout.addWidget(self.server_edit, row, 1)
        row += 1

        hint = QLabel("Kerio Connect server address")
        hint.setStyleSheet("color: gray; font-size: 10px;")
        right_layout.addWidget(hint, row, 1)
        row += 1

        right_layout.addWidget(QLabel("Username:"), row, 0)
        self.username_edit = QLineEdit()
        right_layout.addWidget(self.username_edit, row, 1)
        row += 1

        right_layout.addWidget(QLabel("Password:"), row, 0)
        self.password_edit = QLineEdit()
        self.password_edit.setEchoMode(QLineEdit.EchoMode.Password)
        right_layout.addWidget(self.password_edit, row, 1)
        row += 1

        self.default_check = QCheckBox("Default account for sending")
        right_layout.addWidget(self.default_check, row, 0, 1, 2)
        row += 1

        # Buttons
        btn_layout2 = QHBoxLayout()
        test_btn = QPushButton("Test Connection")
        test_btn.clicked.connect(self._test_connection)
        save_btn = QPushButton("Save")
        save_btn.clicked.connect(self._save_account)
        btn_layout2.addWidget(test_btn)
        btn_layout2.addWidget(save_btn)
        btn_layout2.addStretch()
        right_layout.addLayout(btn_layout2, row, 0, 1, 2)
        row += 1

        # Status
        self.status_label = QLabel("")
        right_layout.addWidget(self.status_label, row, 0, 1, 2)

        right_layout.setRowStretch(row + 1, 1)
        content_layout.addWidget(right_group, 1)

        # Bottom close button (right-aligned)
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        button_layout.addWidget(close_btn)
        main_layout.addLayout(button_layout)

    def _load_accounts(self):
        """Load accounts into list."""
        self.account_list.clear()
        self._accounts = self.db.get_accounts()

        for acc in self._accounts:
            display = acc['name']
            if acc.get('is_default'):
                display = f"{display} (default)"
            self.account_list.addItem(display)

        if self._accounts:
            self.account_list.setCurrentRow(0)

    def _on_account_select(self):
        """Handle account selection."""
        row = self.account_list.currentRow()
        if row < 0 or row >= len(self._accounts):
            return

        acc = self.db.get_account(self._accounts[row]['id'])
        if acc:
            self.name_edit.setText(acc['name'])
            self.email_edit.setText(acc['email'])
            self.server_edit.setText(acc['server'])
            self.username_edit.setText(acc['username'])
            self.password_edit.setText(acc['password'])
            self.default_check.setChecked(bool(acc.get('is_default')))
            self._current_account_id = acc['id']
        else:
            self._current_account_id = None

    def _new_account(self):
        """Clear form for new account."""
        self.account_list.clearSelection()
        self.name_edit.clear()
        self.email_edit.clear()
        self.server_edit.clear()
        self.username_edit.clear()
        self.password_edit.clear()
        self.default_check.setChecked(False)
        self._current_account_id = None
        self.name_edit.setFocus()

    def _delete_account(self):
        """Delete selected account."""
        row = self.account_list.currentRow()
        if row < 0 or row >= len(self._accounts):
            return

        acc = self._accounts[row]
        result = QMessageBox.question(
            self, "Delete Account",
            f"Delete account '{acc['name']}'?\n\nThis will remove all cached messages."
        )
        if result == QMessageBox.StandardButton.Yes:
            self.db.delete_account(acc['id'])
            self._load_accounts()
            self._new_account()
            self.status_label.setText("Account deleted")

    def _validate(self):
        """Validate form fields."""
        if not self.name_edit.text().strip():
            QMessageBox.warning(self, "Validation Error", "Account name is required")
            return False
        if not self.email_edit.text().strip():
            QMessageBox.warning(self, "Validation Error", "Email address is required")
            return False
        if not self.server_edit.text().strip():
            QMessageBox.warning(self, "Validation Error", "Server is required")
            return False
        if not self.username_edit.text().strip():
            QMessageBox.warning(self, "Validation Error", "Username is required")
            return False
        if not self.password_edit.text():
            QMessageBox.warning(self, "Validation Error", "Password is required")
            return False
        return True

    def _test_connection(self):
        """Test connection settings."""
        if not self._validate():
            return

        self.status_label.setText("Testing connection...")
        self.status_label.setStyleSheet("")

        config = KerioConfig(
            email=self.email_edit.text().strip(),
            username=self.username_edit.text().strip(),
            password=self.password_edit.text(),
            server=self.server_edit.text().strip()
        )

        try:
            session = KerioSession(config)
            session.login()
            user_info = session.whoami()
            session.logout()
            self.status_label.setText(f"Connected as {user_info.get('userName', 'user')}")
            self.status_label.setStyleSheet("color: green;")
        except Exception as e:
            self.status_label.setText(f"Failed: {str(e)}")
            self.status_label.setStyleSheet("color: red;")

    def _save_account(self):
        """Save the account."""
        if not self._validate():
            return

        name = self.name_edit.text().strip()
        email = self.email_edit.text().strip()
        server = self.server_edit.text().strip()
        username = self.username_edit.text().strip()
        password = self.password_edit.text()
        is_default = self.default_check.isChecked()

        # Check for duplicate
        existing = self.db.get_account_by_name(name)
        if existing and existing['id'] != self._current_account_id:
            QMessageBox.warning(self, "Error", f"An account named '{name}' already exists")
            return

        self.db.save_account(
            name=name,
            email=email,
            server=server,
            username=username,
            password=password,
            is_default=is_default,
            account_id=self._current_account_id
        )

        self._load_accounts()
        self.status_label.setText("Account saved")
        self.status_label.setStyleSheet("color: green;")

        # Notify app to refresh
        if self.app and hasattr(self.app, '_load_accounts'):
            self.app._load_accounts()


class SettingsDialog(QDialog):
    """Settings dialog."""

    def __init__(self, parent, db, app):
        super().__init__(parent)
        self.db = db
        self.app = app

        self.setWindowTitle("Settings")
        self.setMinimumSize(450, 250)
        self.setModal(True)

        self._create_ui()

    def _create_ui(self):
        layout = QVBoxLayout(self)

        # Font size
        font_layout = QHBoxLayout()
        font_layout.addWidget(QLabel("Font Size:"))
        self.font_spin = QSpinBox()
        self.font_spin.setRange(6, 24)
        self.font_spin.setValue(getattr(self.app, 'font_size', 12))
        font_layout.addWidget(self.font_spin)
        font_layout.addStretch()
        layout.addLayout(font_layout)

        # Theme
        theme_layout = QHBoxLayout()
        theme_layout.addWidget(QLabel("Theme:"))
        self.theme_combo = QComboBox()
        self.theme_combo.addItems(["System", "Light", "Dark"])
        current_theme = self.db.get_setting("theme", "System")
        index = self.theme_combo.findText(current_theme)
        if index >= 0:
            self.theme_combo.setCurrentIndex(index)
        theme_layout.addWidget(self.theme_combo)
        theme_layout.addStretch()
        layout.addLayout(theme_layout)

        # Default save directory
        save_layout = QHBoxLayout()
        save_layout.addWidget(QLabel("Default Save Directory:"))
        current_dir = self.db.get_setting("default_save_directory", "")
        if not current_dir:
            current_dir = os.path.join(os.path.expanduser("~"), "Downloads")
        self.save_dir_edit = QLineEdit(current_dir)
        save_layout.addWidget(self.save_dir_edit, 1)
        browse_btn = QPushButton("Browse...")
        browse_btn.clicked.connect(self._browse_save_dir)
        save_layout.addWidget(browse_btn)
        layout.addLayout(save_layout)

        # Desktop launcher
        launcher_layout = QHBoxLayout()
        launcher_layout.addWidget(QLabel("Desktop Launcher:"))
        install_btn = QPushButton("Install")
        install_btn.clicked.connect(self._install_launcher)
        remove_btn = QPushButton("Remove")
        remove_btn.clicked.connect(self._remove_launcher)
        launcher_layout.addWidget(install_btn)
        launcher_layout.addWidget(remove_btn)
        launcher_layout.addStretch()
        layout.addLayout(launcher_layout)

        layout.addStretch()

        # Buttons
        btn_box = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        btn_box.accepted.connect(self._save)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

    def _browse_save_dir(self):
        """Browse for save directory."""
        initial = self.save_dir_edit.text()
        if not os.path.isdir(initial):
            initial = os.path.expanduser("~")
        new_dir = QFileDialog.getExistingDirectory(self, "Select Default Save Directory", initial)
        if new_dir:
            self.save_dir_edit.setText(new_dir)

    def _save(self):
        """Save settings."""
        font_size = self.font_spin.value()
        self.db.set_setting("font_size", str(font_size))
        if hasattr(self.app, 'font_size'):
            self.app.font_size = font_size
            if hasattr(self.app, '_apply_font_size'):
                self.app._apply_font_size()

        # Save theme
        theme = self.theme_combo.currentText()
        self.db.set_setting("theme", theme)
        if hasattr(self.app, '_apply_theme_setting'):
            self.app._apply_theme_setting(theme)

        save_dir = self.save_dir_edit.text().strip()
        if save_dir and os.path.isdir(save_dir):
            self.db.set_setting("default_save_directory", save_dir)

        self.accept()

    def _install_launcher(self):
        """Install desktop launcher."""
        from mailbench.launcher import create_launcher
        if create_launcher():
            QMessageBox.information(self, "Success", "Desktop launcher installed")
        else:
            QMessageBox.warning(self, "Error", "Failed to install launcher")

    def _remove_launcher(self):
        """Remove desktop launcher."""
        from mailbench.launcher import remove_launcher
        if remove_launcher():
            QMessageBox.information(self, "Success", "Desktop launcher removed")
        else:
            QMessageBox.warning(self, "Error", "No launcher found to remove")
