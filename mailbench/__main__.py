"""Entry point for mailbench."""

import sys


def check_pyside6():
    """Check if PySide6 is available and provide installation instructions if not."""
    try:
        import PySide6
        return True
    except ImportError:
        pass

    import platform
    system = platform.system().lower()

    print("Error: PySide6 is not installed.")
    print()
    print("Mailbench requires PySide6 for its graphical interface.")
    print()

    if system == "linux":
        print("To install PySide6:")
        print("  pip install PySide6")
    elif system == "darwin":
        print("To install on macOS:")
        print("  pip install PySide6")
    elif system == "windows":
        print("To install on Windows:")
        print("  pip install PySide6")
    else:
        print("To install PySide6:")
        print("  pip install PySide6")

    print()
    print("After installing, run mailbench again.")
    return False


def main():
    """Main entry point with argument handling."""
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

        elif arg in ("--help", "-h"):
            print("Mailbench - Python Email Client for Kerio Connect")
            print()
            print("Usage: mailbench [options]")
            print()
            print("Options:")
            print("  --install-launcher   Create a desktop launcher for this OS")
            print("  --remove-launcher    Remove the desktop launcher")
            print("  --help, -h           Show this help message")
            print()
            print("Run without arguments to start the application.")
            sys.exit(0)

    if not check_pyside6():
        sys.exit(1)

    from mailbench.app import main as app_main
    app_main()


if __name__ == "__main__":
    main()
