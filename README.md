# Mailbench

A desktop email client for Kerio Connect mail servers, built with Python and Tkinter.

## Features

- **Multi-account support** - Connect to multiple Kerio Connect servers
- **3-pane layout** - Folders, message list, and preview pane
- **Dark/light themes** - Toggle between themes with Ctrl+D
- **Compose, Reply, Forward** - Full email composition with HTML support
- **Attachments** - View, download, and attach files
- **Address book integration** - Autocomplete from Kerio contacts and cached addresses
- **Local caching** - SQLite database for offline message access
- **Keyboard shortcuts** - Efficient navigation and actions
- **Zoom support** - Ctrl+scroll to adjust preview size

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

## Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| Ctrl+N | New message |
| Ctrl+R | Reply |
| Ctrl+Shift+R | Reply All |
| Ctrl+F | Forward |
| Ctrl+D | Toggle dark mode |
| Ctrl+Scroll | Zoom preview |
| Delete | Delete message |
| Escape | Clear filter / Discard compose |

## License

MIT License - see [LICENSE](LICENSE) file.
