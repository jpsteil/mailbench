"""Entry point for mailbench."""

import sys


def check_pyside6():
    """Check if PySide6 is available."""
    try:
        import PySide6
        return True
    except ImportError:
        return False


def check_tkinter():
    """Check if tkinter is available and provide installation instructions if not."""
    try:
        import tkinter
        return True
    except ImportError:
        pass

    import platform
    system = platform.system().lower()

    print("Error: tkinter is not installed.")
    print()
    print("Mailbench requires tkinter for its graphical interface.")
    print()

    if system == "linux":
        distro = ""
        try:
            with open("/etc/os-release") as f:
                for line in f:
                    if line.startswith("ID="):
                        distro = line.strip().split("=")[1].strip('"').lower()
                        break
                    if line.startswith("ID_LIKE="):
                        distro_like = line.strip().split("=")[1].strip('"').lower()
                        if not distro:
                            distro = distro_like
        except Exception:
            pass

        if distro in ("ubuntu", "debian", "pop", "mint", "elementary") or "debian" in distro or "ubuntu" in distro:
            print("To install on Debian/Ubuntu-based systems:")
            print("  sudo apt install python3-tk")
        elif distro in ("fedora", "rhel", "centos", "rocky", "alma") or "fedora" in distro or "rhel" in distro:
            print("To install on Fedora/RHEL-based systems:")
            print("  sudo dnf install python3-tkinter")
        elif distro in ("arch", "manjaro", "endeavouros") or "arch" in distro:
            print("To install on Arch-based systems:")
            print("  sudo pacman -S tk")
        elif distro in ("opensuse", "suse") or "suse" in distro:
            print("To install on openSUSE:")
            print("  sudo zypper install python3-tk")
        else:
            print("To install tkinter, use your distribution's package manager:")
            print("  Debian/Ubuntu: sudo apt install python3-tk")
            print("  Fedora/RHEL:   sudo dnf install python3-tkinter")
            print("  Arch Linux:    sudo pacman -S tk")
            print("  openSUSE:      sudo zypper install python3-tk")
    elif system == "darwin":
        print("To install on macOS:")
        print("  brew install python-tk")
        print()
        print("Or reinstall Python with tkinter support:")
        print("  brew reinstall python")
    elif system == "windows":
        print("On Windows, tkinter should be included with Python.")
        print("Try reinstalling Python and ensure 'tcl/tk and IDLE' is selected.")
    else:
        print("Please install tkinter for your operating system.")

    print()
    print("After installing, run mailbench again.")
    return False


def main():
    """Main entry point with argument handling."""
    use_tk = False

    if len(sys.argv) > 1:
        arg = sys.argv[1].lower()

        if arg in ("--install-launcher", "--create-launcher"):
            from mailbench.launcher import create_launcher
            success = create_launcher()
            sys.exit(0 if success else 1)

        elif arg in ("--remove-launcher", "--uninstall-launcher"):
            from mailbench.launcher import remove_launcher
            success = remove_launcher()
            sys.exit(0 if success else 1)

        elif arg == "--tk":
            use_tk = True

        elif arg in ("--help", "-h"):
            print("Mailbench - Python Email Client for Kerio Connect")
            print()
            print("Usage: mailbench [options]")
            print()
            print("Options:")
            print("  --install-launcher   Create a desktop launcher for this OS")
            print("  --remove-launcher    Remove the desktop launcher")
            print("  --tk                 Use tkinter UI instead of Qt")
            print("  --help, -h           Show this help message")
            print()
            print("Run without arguments to start the application.")
            sys.exit(0)

    # Use Qt by default if available, fall back to tkinter
    if use_tk or not check_pyside6():
        if not check_tkinter():
            sys.exit(1)
        from mailbench.app import main as app_main
        app_main()
    else:
        from mailbench.app_qt import main as app_main
        app_main()


if __name__ == "__main__":
    main()
