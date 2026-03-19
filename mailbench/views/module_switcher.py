"""Module switcher panel for navigating between Mail, Contacts, Calendar, etc."""

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QToolButton, QButtonGroup, QFrame, QSizePolicy
)
from PySide6.QtGui import QIcon

from mailbench.icons import get_module_icon


class ModuleButton(QToolButton):
    """A styled button for the module switcher."""

    def __init__(self, module_id: str, icon: QIcon, tooltip: str, parent=None):
        super().__init__(parent)
        self.module_id = module_id
        self.setIcon(icon)
        self.setToolTip(tooltip)
        self.setCheckable(True)
        self.setAutoRaise(True)
        self.setIconSize(self.sizeHint())
        self.setFixedSize(40, 40)
        self.setCursor(Qt.CursorShape.PointingHandCursor)


class ModuleSwitcher(QWidget):
    """Vertical panel with module icons for navigation."""

    moduleSelected = Signal(str)  # Emits: 'mail', 'calendar', 'contacts', 'tasks', 'notes'

    def __init__(self, parent=None):
        super().__init__(parent)
        self._current_module = 'mail'
        self._setup_ui()

    def _setup_ui(self):
        self.setFixedWidth(48)
        self.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Expanding)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(4, 8, 4, 8)
        layout.setSpacing(4)

        # Create button group for exclusive selection
        self._button_group = QButtonGroup(self)
        self._button_group.setExclusive(True)

        # Module buttons
        modules = [
            ('mail', 'Mail', True),
            ('contacts', 'Contacts', True),
            ('calendar', 'Calendar', False),
            ('tasks', 'Tasks', False),
            ('notes', 'Notes', False),
        ]

        self._buttons = {}
        for module_id, tooltip, enabled in modules:
            icon = get_module_icon(module_id)
            btn = ModuleButton(module_id, icon, tooltip, self)
            btn.setEnabled(enabled)
            if not enabled:
                btn.setToolTip(f"{tooltip} (Coming Soon)")
            btn.clicked.connect(lambda checked, m=module_id: self._on_button_clicked(m))
            self._button_group.addButton(btn)
            self._buttons[module_id] = btn
            layout.addWidget(btn)

        # Add stretch to push buttons to top
        layout.addStretch()

        # Set initial selection
        self._buttons['mail'].setChecked(True)

        # Style the panel
        self.setStyleSheet("""
            ModuleSwitcher {
                background-color: #f0f0f0;
                border-right: 1px solid #d0d0d0;
            }
            ModuleButton {
                border: none;
                border-radius: 6px;
                padding: 6px;
            }
            ModuleButton:hover {
                background-color: #e0e0e0;
            }
            ModuleButton:checked {
                background-color: #d0d0d0;
            }
            ModuleButton:disabled {
                opacity: 0.4;
            }
        """)

    def _on_button_clicked(self, module_id: str):
        if module_id != self._current_module:
            self._current_module = module_id
            self.moduleSelected.emit(module_id)

    def current_module(self) -> str:
        """Return the currently selected module."""
        return self._current_module

    def set_module(self, module_id: str):
        """Programmatically select a module."""
        if module_id in self._buttons:
            self._buttons[module_id].setChecked(True)
            self._current_module = module_id

    def set_module_enabled(self, module_id: str, enabled: bool):
        """Enable or disable a module button."""
        if module_id in self._buttons:
            self._buttons[module_id].setEnabled(enabled)
