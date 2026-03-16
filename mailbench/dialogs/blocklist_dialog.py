"""Blocked senders management dialog."""

from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTabWidget, QWidget,
    QTableWidget, QTableWidgetItem, QPushButton, QHeaderView,
    QInputDialog, QMessageBox, QAbstractItemView, QLineEdit, QLabel
)
from PySide6.QtCore import Qt


class BlocklistDialog(QDialog):
    """Dialog for managing blocked senders (domains and emails)."""

    def __init__(self, blocklist_manager, parent=None):
        super().__init__(parent)
        self.blocklist_manager = blocklist_manager
        self.setWindowTitle("Manage Blocked Senders")
        self.setMinimumSize(700, 500)
        self._setup_ui()
        self._load_data()

    def _create_search_box(self, table, placeholder="Search..."):
        """Create a search box that filters a table."""
        search = QLineEdit()
        search.setPlaceholderText(placeholder)
        search.setClearButtonEnabled(True)
        search.textChanged.connect(lambda text: self._filter_table(table, text))
        return search

    def _filter_table(self, table, text):
        """Filter table rows based on search text."""
        text = text.lower()
        for row in range(table.rowCount()):
            match = False
            for col in range(table.columnCount()):
                item = table.item(row, col)
                if item and text in item.text().lower():
                    match = True
                    break
            table.setRowHidden(row, not match)

    def _setup_ui(self):
        """Set up the dialog UI."""
        layout = QVBoxLayout(self)

        # Tab widget
        self._tabs = QTabWidget()
        layout.addWidget(self._tabs)

        # Domains tab
        domains_widget = QWidget()
        domains_layout = QVBoxLayout(domains_widget)

        self._domains_search = self._create_search_box(None, "Search domains...")
        domains_layout.addWidget(self._domains_search)

        self._domains_table = QTableWidget()
        self._domains_table.setColumnCount(3)
        self._domains_table.setHorizontalHeaderLabels(["Domain", "Blocked", "Last Blocked"])
        self._domains_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self._domains_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self._domains_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self._domains_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._domains_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._domains_table.setSortingEnabled(True)
        domains_layout.addWidget(self._domains_table)

        # Connect search after table is created
        self._domains_search.textChanged.disconnect()
        self._domains_search.textChanged.connect(
            lambda text: self._filter_table(self._domains_table, text)
        )

        domains_buttons = QHBoxLayout()
        add_domain_btn = QPushButton("Add Domain")
        add_domain_btn.clicked.connect(self._add_domain)
        domains_buttons.addWidget(add_domain_btn)

        remove_domain_btn = QPushButton("Remove Selected")
        remove_domain_btn.clicked.connect(self._remove_domain)
        domains_buttons.addWidget(remove_domain_btn)

        domains_buttons.addStretch()
        domains_layout.addLayout(domains_buttons)

        self._tabs.addTab(domains_widget, "Blocked Domains")

        # Emails tab
        emails_widget = QWidget()
        emails_layout = QVBoxLayout(emails_widget)

        self._emails_search = self._create_search_box(None, "Search emails...")
        emails_layout.addWidget(self._emails_search)

        self._emails_table = QTableWidget()
        self._emails_table.setColumnCount(3)
        self._emails_table.setHorizontalHeaderLabels(["Email", "Blocked", "Last Blocked"])
        self._emails_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self._emails_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        self._emails_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        self._emails_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._emails_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._emails_table.setSortingEnabled(True)
        emails_layout.addWidget(self._emails_table)

        # Connect search after table is created
        self._emails_search.textChanged.disconnect()
        self._emails_search.textChanged.connect(
            lambda text: self._filter_table(self._emails_table, text)
        )

        emails_buttons = QHBoxLayout()
        add_email_btn = QPushButton("Add Email")
        add_email_btn.clicked.connect(self._add_email)
        emails_buttons.addWidget(add_email_btn)

        remove_email_btn = QPushButton("Remove Selected")
        remove_email_btn.clicked.connect(self._remove_email)
        emails_buttons.addWidget(remove_email_btn)

        emails_buttons.addStretch()
        emails_layout.addLayout(emails_buttons)

        self._tabs.addTab(emails_widget, "Blocked Emails")

        # Allowed Domains tab (formerly Never Block)
        allowed_widget = QWidget()
        allowed_layout = QVBoxLayout(allowed_widget)

        self._allowed_search = self._create_search_box(None, "Search allowed domains...")
        allowed_layout.addWidget(self._allowed_search)

        self._allowed_table = QTableWidget()
        self._allowed_table.setColumnCount(1)
        self._allowed_table.setHorizontalHeaderLabels(["Domain"])
        self._allowed_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self._allowed_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self._allowed_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self._allowed_table.setSortingEnabled(True)
        allowed_layout.addWidget(self._allowed_table)

        # Connect search after table is created
        self._allowed_search.textChanged.disconnect()
        self._allowed_search.textChanged.connect(
            lambda text: self._filter_table(self._allowed_table, text)
        )

        allowed_buttons = QHBoxLayout()
        add_allowed_btn = QPushButton("Add Domain")
        add_allowed_btn.clicked.connect(self._add_allowed)
        allowed_buttons.addWidget(add_allowed_btn)

        remove_allowed_btn = QPushButton("Remove Selected")
        remove_allowed_btn.clicked.connect(self._remove_allowed)
        allowed_buttons.addWidget(remove_allowed_btn)

        allowed_buttons.addStretch()
        allowed_layout.addLayout(allowed_buttons)

        self._tabs.addTab(allowed_widget, "Allowed Domains")

        # Close button
        close_layout = QHBoxLayout()
        close_layout.addStretch()
        close_btn = QPushButton("Close")
        close_btn.clicked.connect(self.accept)
        close_layout.addWidget(close_btn)
        layout.addLayout(close_layout)

    def _load_data(self):
        """Load data into tables."""
        # Load blocked domains
        domains = self.blocklist_manager.get_domains()
        self._domains_table.setSortingEnabled(False)
        self._domains_table.setRowCount(len(domains))
        for i, entry in enumerate(domains):
            self._domains_table.setItem(i, 0, QTableWidgetItem(entry.get('domain', '')))
            self._domains_table.setItem(i, 1, QTableWidgetItem(str(entry.get('blocked_count', 0))))
            last_blocked = entry.get('last_blocked') or 'Never'
            self._domains_table.setItem(i, 2, QTableWidgetItem(str(last_blocked)))
        self._domains_table.setSortingEnabled(True)

        # Update tab label
        self._tabs.setTabText(0, f"Blocked Domains ({len(domains)})")

        # Load blocked emails
        emails = self.blocklist_manager.get_emails()
        self._emails_table.setSortingEnabled(False)
        self._emails_table.setRowCount(len(emails))
        for i, entry in enumerate(emails):
            self._emails_table.setItem(i, 0, QTableWidgetItem(entry.get('email', '')))
            self._emails_table.setItem(i, 1, QTableWidgetItem(str(entry.get('blocked_count', 0))))
            last_blocked = entry.get('last_blocked') or 'Never'
            self._emails_table.setItem(i, 2, QTableWidgetItem(str(last_blocked)))
        self._emails_table.setSortingEnabled(True)

        # Update tab label
        self._tabs.setTabText(1, f"Blocked Emails ({len(emails)})")

        # Load allowed domains
        allowed = self.blocklist_manager.get_allowed_domains()
        self._allowed_table.setSortingEnabled(False)
        self._allowed_table.setRowCount(len(allowed))
        for i, entry in enumerate(allowed):
            self._allowed_table.setItem(i, 0, QTableWidgetItem(entry.get('domain', '')))
        self._allowed_table.setSortingEnabled(True)

        # Update tab label
        self._tabs.setTabText(2, f"Allowed Domains ({len(allowed)})")

        # Re-apply search filters
        self._filter_table(self._domains_table, self._domains_search.text())
        self._filter_table(self._emails_table, self._emails_search.text())
        self._filter_table(self._allowed_table, self._allowed_search.text())

    def _add_domain(self):
        """Add a new domain to block list."""
        domain, ok = QInputDialog.getText(
            self, "Block Domain",
            "Enter domain to block (e.g., spam.com):"
        )
        if ok and domain:
            domain = domain.strip().lower()
            if not domain:
                return

            def on_added(success, error):
                if success:
                    self._load_data()
                else:
                    QMessageBox.warning(self, "Error", f"Failed to block domain: {error}")

            self.blocklist_manager.add_domain(domain, on_added)

    def _remove_domain(self):
        """Remove selected domain from block list."""
        row = self._domains_table.currentRow()
        if row < 0:
            return

        domain = self._domains_table.item(row, 0).text()

        result = QMessageBox.question(
            self, "Unblock Domain",
            f"Unblock '{domain}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if result == QMessageBox.StandardButton.Yes:
            def on_removed(success, error):
                if success:
                    self._load_data()
                else:
                    QMessageBox.warning(self, "Error", f"Failed to unblock domain: {error}")

            self.blocklist_manager.remove_domain(domain, on_removed)

    def _add_email(self):
        """Add a new email to block list."""
        email, ok = QInputDialog.getText(
            self, "Block Email",
            "Enter email address to block:"
        )
        if ok and email:
            email = email.strip().lower()
            if not email or '@' not in email:
                QMessageBox.warning(self, "Invalid Email", "Please enter a valid email address.")
                return

            def on_added(success, error):
                if success:
                    self._load_data()
                else:
                    QMessageBox.warning(self, "Error", f"Failed to block email: {error}")

            self.blocklist_manager.add_email(email, on_added)

    def _remove_email(self):
        """Remove selected email from block list."""
        row = self._emails_table.currentRow()
        if row < 0:
            return

        email = self._emails_table.item(row, 0).text()

        result = QMessageBox.question(
            self, "Unblock Email",
            f"Unblock '{email}'?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if result == QMessageBox.StandardButton.Yes:
            def on_removed(success, error):
                if success:
                    self._load_data()
                else:
                    QMessageBox.warning(self, "Error", f"Failed to unblock email: {error}")

            self.blocklist_manager.remove_email(email, on_removed)

    def _add_allowed(self):
        """Add a domain to allowed list."""
        domain, ok = QInputDialog.getText(
            self, "Add Allowed Domain",
            "Enter domain to always allow (e.g., company.com):"
        )
        if ok and domain:
            domain = domain.strip().lower()
            if not domain:
                return

            def on_added(success, error):
                if success:
                    self._load_data()
                else:
                    QMessageBox.warning(self, "Error", f"Failed to add domain: {error}")

            self.blocklist_manager.add_allowed_domain(domain, on_added)

    def _remove_allowed(self):
        """Remove selected domain from allowed list."""
        row = self._allowed_table.currentRow()
        if row < 0:
            return

        domain = self._allowed_table.item(row, 0).text()

        result = QMessageBox.question(
            self, "Remove Allowed Domain",
            f"Remove '{domain}' from allowed list?\n\nThis will allow blocking emails from this domain.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
        )

        if result == QMessageBox.StandardButton.Yes:
            def on_removed(success, error):
                if success:
                    self._load_data()
                else:
                    QMessageBox.warning(self, "Error", f"Failed to remove domain: {error}")

            self.blocklist_manager.remove_allowed_domain(domain, on_removed)
