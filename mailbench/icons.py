"""Modern SVG icons for Mailbench."""

from PySide6.QtCore import QByteArray, QSize
from PySide6.QtGui import QIcon, QPixmap, QPainter
from PySide6.QtSvg import QSvgRenderer


# Icon color - muted gray that works well on light backgrounds
ICON_COLOR = "#5f6368"


def _svg_to_icon(svg_content: str, size: int = 16) -> QIcon:
    """Convert SVG string to QIcon."""
    # Replace placeholder color
    svg_content = svg_content.replace("{color}", ICON_COLOR)

    svg_bytes = QByteArray(svg_content.encode('utf-8'))
    renderer = QSvgRenderer(svg_bytes)

    icon = QIcon()
    for s in [16, 24, 32]:
        pixmap = QPixmap(QSize(s, s))
        pixmap.fill("transparent")
        painter = QPainter(pixmap)
        renderer.render(painter)
        painter.end()
        icon.addPixmap(pixmap)

    return icon


# Modern outline-style SVG icons
# All icons are 24x24 viewBox with 1.5px stroke

_SVG_INBOX = '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
  <polyline points="22 12 16 12 14 15 10 15 8 12 2 12"/>
  <path d="M5.45 5.11L2 12v6a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2v-6l-3.45-6.89A2 2 0 0 0 16.76 4H7.24a2 2 0 0 0-1.79 1.11z"/>
</svg>'''

_SVG_SENT = '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
  <line x1="22" y1="2" x2="11" y2="13"/>
  <polygon points="22 2 15 22 11 13 2 9 22 2"/>
</svg>'''

_SVG_DRAFTS = '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
  <path d="M11 4H4a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7"/>
  <path d="M18.5 2.5a2.121 2.121 0 0 1 3 3L12 15l-4 1 1-4 9.5-9.5z"/>
</svg>'''

_SVG_JUNK = '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
  <circle cx="12" cy="12" r="10"/>
  <line x1="4.93" y1="4.93" x2="19.07" y2="19.07"/>
</svg>'''

_SVG_TRASH = '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
  <polyline points="3 6 5 6 21 6"/>
  <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"/>
</svg>'''

_SVG_FOLDER = '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
  <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/>
</svg>'''

_SVG_FOLDER_OPEN = '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
  <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2v1"/>
  <path d="M3 10h18l-2 9H5l-2-9z"/>
</svg>'''

_SVG_ARCHIVE = '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
  <polyline points="21 8 21 21 3 21 3 8"/>
  <rect x="1" y="3" width="22" height="5"/>
  <line x1="10" y1="12" x2="14" y2="12"/>
</svg>'''

_SVG_ACCOUNT = '''<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round">
  <path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/>
  <polyline points="22,6 12,13 2,6"/>
</svg>'''


# Cache for created icons
_icon_cache = {}


def get_folder_icon(folder_type: str) -> QIcon:
    """Get a modern folder icon by type.

    Args:
        folder_type: One of 'inbox', 'sent', 'drafts', 'junk', 'trash',
                     'folder', 'folder_open', 'archive', 'account'

    Returns:
        QIcon for the folder type
    """
    if folder_type in _icon_cache:
        return _icon_cache[folder_type]

    svg_map = {
        'inbox': _SVG_INBOX,
        'sent': _SVG_SENT,
        'drafts': _SVG_DRAFTS,
        'junk': _SVG_JUNK,
        'spam': _SVG_JUNK,
        'trash': _SVG_TRASH,
        'deleted': _SVG_TRASH,
        'folder': _SVG_FOLDER,
        'folder_open': _SVG_FOLDER_OPEN,
        'archive': _SVG_ARCHIVE,
        'account': _SVG_ACCOUNT,
    }

    svg = svg_map.get(folder_type, _SVG_FOLDER)
    icon = _svg_to_icon(svg)
    _icon_cache[folder_type] = icon
    return icon
