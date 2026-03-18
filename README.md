# Mailbench

A desktop email client for Kerio Connect mail servers, built with Python and Qt.

## Features

- **Multi-account support** - Connect to multiple Kerio Connect servers
- **3-pane layout** - Folders, message list, and preview pane
- **Compose, Reply, Forward** - Full email composition with HTML support
- **WebEngine-based editor** - Rich HTML compose with inline image support
- **Attachments** - View, download, and attach files
- **Address book integration** - Autocomplete from Kerio contacts and cached addresses
- **Local caching** - SQLite database for offline message access
- **Drag and drop** - Move messages between folders
- **Block sender/domain** - Drag messages to block folders
- **Message flagging** - Flag important messages
- **Keyboard shortcuts** - Efficient navigation and actions
- **Zoom support** - Ctrl+scroll or shortcuts to adjust preview size

## Security Features

- **Remote image blocking** - Prevents email tracking pixels
- **External link warnings** - Shows domain before opening links
- **Phishing detection** - Warns about homograph/lookalike domains
- **Dangerous attachment warnings** - Alerts for executable files
- **HTML sanitization** - XSS protection for message display
- **Secure credential storage** - Uses system keyring for passwords

## Requirements

- Python 3.10+
- Kerio Connect mail server with JSON-RPC API access

## Installation

```bash
# Clone the repository
git clone https://github.com/jpsteil/mailbench.git
cd mailbench

# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -e .
```

## Usage

```bash
mailbench
```

On first run, go to **Accounts > Add Account** to configure your Kerio Connect server.

The status bar displays a permanent zoom indicator showing the current preview zoom level.

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| Ctrl+N | New message |
| Ctrl+R | Reply |
| Ctrl+Shift+R | Reply All |
| Ctrl+F | Forward |
| Ctrl+Scroll | Zoom preview |
| Ctrl++ or Ctrl+= | Zoom in |
| Ctrl+- | Zoom out |
| Ctrl+0 | Reset zoom |
| Ctrl+Enter | Send message (in compose) |
| Delete | Delete message |
| Escape | Clear filter / Discard compose |

## License

MIT License - see [LICENSE](LICENSE) file.
