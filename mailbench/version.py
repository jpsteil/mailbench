"""Version checking for Mailbench."""

import threading
import urllib.request
import json

__version__ = "0.3.0"


def get_installed_version():
    """Get the currently installed version."""
    return __version__


def get_pypi_version():
    """Fetch the latest version from PyPI."""
    try:
        url = "https://pypi.org/pypi/mailbench/json"
        with urllib.request.urlopen(url, timeout=5) as response:
            data = json.loads(response.read().decode())
            return data["info"]["version"]
    except Exception:
        return None


def parse_version(version_str):
    """Parse version string into tuple for comparison."""
    try:
        parts = version_str.split(".")
        return tuple(int(p) for p in parts)
    except (ValueError, AttributeError):
        return (0, 0, 0)


def is_newer_version(pypi_version, installed_version):
    """Check if PyPI version is newer than installed version."""
    return parse_version(pypi_version) > parse_version(installed_version)


def check_for_updates(callback):
    """Check for updates in background thread.

    Args:
        callback: Function to call with (has_update, latest_version) or (False, None) on error
    """
    def do_check():
        try:
            latest = get_pypi_version()
            if latest and is_newer_version(latest, __version__):
                callback(True, latest)
            else:
                callback(False, latest)
        except Exception:
            callback(False, None)

    thread = threading.Thread(target=do_check, daemon=True)
    thread.start()
