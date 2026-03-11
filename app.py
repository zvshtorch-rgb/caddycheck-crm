"""CaddyCheck CRM – entry point."""
import sys
import logging
from pathlib import Path

from config.settings import LOG_LEVEL, LOG_FORMAT, OUTPUT_DIR, DATA_DIR

# ── Configure logging before importing Qt ──────────────────────────────────────
logging.basicConfig(level=LOG_LEVEL, format=LOG_FORMAT)
logger = logging.getLogger(__name__)


def check_data_files():
    """Warn if expected data files are missing."""
    from config.settings import PROJECTS_FILE, INVOICE_TEMPLATE
    missing = []
    if not PROJECTS_FILE.exists():
        missing.append(str(PROJECTS_FILE))
    if not INVOICE_TEMPLATE.exists():
        missing.append(str(INVOICE_TEMPLATE))
    return missing


def main():
    from PySide6.QtWidgets import QApplication, QMessageBox
    from PySide6.QtCore import Qt
    from PySide6.QtGui import QFont

    app = QApplication(sys.argv)
    app.setApplicationName("CaddyCheck CRM")
    app.setOrganizationName("Video Inform Ltd")

    # Set a clean, modern font
    font = QFont("Segoe UI", 10)
    app.setFont(font)

    # Ensure output directory exists
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # Check data files
    missing = check_data_files()
    if missing:
        QMessageBox.warning(
            None,
            "Missing Data Files",
            "The following required files were not found:\n\n"
            + "\n".join(missing)
            + "\n\nThe application will start but data may not load correctly.\n"
            "Place your Excel files in the 'data/' directory.",
        )

    from ui.main_window import MainWindow
    window = MainWindow()
    window.show()

    logger.info("CaddyCheck CRM started.")
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
