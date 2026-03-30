"""Main application window with sidebar navigation."""
import logging
from typing import List

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QHBoxLayout, QVBoxLayout,
    QPushButton, QStackedWidget, QLabel, QFrame, QSizePolicy,
    QApplication, QMessageBox,
)
from PySide6.QtCore import Qt, QThread, Signal, QObject
from PySide6.QtGui import QFont, QColor, QPalette

from models.project import Project
from models.invoice import Invoice, DebtSummary
from services.excel_service import (
    load_projects, load_invoices, compute_debt_summaries,
    get_yearly_summary, get_monthly_summary,
)

logger = logging.getLogger(__name__)

# ── Colour palette ─────────────────────────────────────────────────────────────
SIDEBAR_BG = "#1E2A3A"
SIDEBAR_HOVER = "#2D4A6A"
SIDEBAR_ACTIVE = "#2D6A9F"
CONTENT_BG = "#F5F7FA"
TEXT_LIGHT = "#FFFFFF"
TEXT_DARK = "#2C3E50"
ACCENT = "#2D6A9F"
BORDER_COLOR = "#DEE2E6"


NAV_ITEMS = [
    ("Dashboard",        "dashboard"),
    ("Projects",         "projects"),
    ("Invoice Details",  "invoices"),
    ("Monthly Invoice",  "monthly_invoice"),
    ("Settings",         "settings"),
]


class DataLoader(QObject):
    """Worker to load data off the main thread."""
    finished = Signal(list, list, list, dict)
    error = Signal(str)

    def run(self):
        try:
            projects = load_projects()
            invoices = load_invoices()
            debt_summaries = compute_debt_summaries(projects, invoices)
            yearly_summary = get_yearly_summary(invoices)
            self.finished.emit(projects, invoices, debt_summaries, yearly_summary)
        except Exception as e:
            logger.exception("Data load failed")
            self.error.emit(str(e))


class MainWindow(QMainWindow):
    """Top-level application window."""

    def __init__(self):
        super().__init__()
        self.setWindowTitle("CaddyCheck CRM")
        self.resize(1280, 800)
        self.setMinimumSize(1000, 600)

        # Data
        self.projects: List[Project] = []
        self.invoices: List[Invoice] = []
        self.debt_summaries: List[DebtSummary] = []
        self.yearly_summary: dict = {}

        self._nav_buttons: dict = {}
        self._current_page = "dashboard"

        self._build_ui()
        self._load_data()

    # ── UI construction ────────────────────────────────────────────────────────

    def _build_ui(self):
        root = QWidget()
        root_layout = QHBoxLayout(root)
        root_layout.setContentsMargins(0, 0, 0, 0)
        root_layout.setSpacing(0)

        # Sidebar
        sidebar = self._build_sidebar()
        root_layout.addWidget(sidebar)

        # Content area
        self._stack = QStackedWidget()
        self._stack.setStyleSheet(f"background: {CONTENT_BG};")
        root_layout.addWidget(self._stack, stretch=1)

        self.setCentralWidget(root)
        self._apply_global_style()

        # Pages are created lazily after data is loaded
        self._pages: dict = {}

    def _build_sidebar(self) -> QWidget:
        sidebar = QFrame()
        sidebar.setFixedWidth(200)
        sidebar.setStyleSheet(f"background: {SIDEBAR_BG}; border: none;")

        layout = QVBoxLayout(sidebar)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Logo / title
        title_label = QLabel("CaddyCheck\nCRM")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setFixedHeight(80)
        title_label.setStyleSheet(
            f"color: {TEXT_LIGHT}; font-size: 18px; font-weight: bold;"
            f" background: {ACCENT}; padding: 10px;"
        )
        layout.addWidget(title_label)

        # Spacer
        layout.addSpacing(10)

        # Nav buttons
        for label, key in NAV_ITEMS:
            btn = QPushButton(label)
            btn.setFixedHeight(48)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setStyleSheet(self._nav_btn_style(active=False))
            btn.clicked.connect(lambda checked, k=key: self._navigate(k))
            layout.addWidget(btn)
            self._nav_buttons[key] = btn

        layout.addStretch()

        # Version label
        version_label = QLabel("v1.0")
        version_label.setAlignment(Qt.AlignCenter)
        version_label.setStyleSheet(f"color: #6C757D; font-size: 11px; padding: 8px;")
        layout.addWidget(version_label)

        return sidebar

    @staticmethod
    def _nav_btn_style(active: bool) -> str:
        bg = SIDEBAR_ACTIVE if active else "transparent"
        return (
            f"QPushButton {{ background: {bg}; color: {TEXT_LIGHT}; "
            f"border: none; text-align: left; padding: 0 20px; "
            f"font-size: 14px; }}"
            f"QPushButton:hover {{ background: {SIDEBAR_HOVER}; }}"
        )

    def _apply_global_style(self):
        self.setStyleSheet(
            "QToolTip { background: #2D6A9F; color: white; border: none; }"
            "QScrollBar:vertical { background: #F0F0F0; width: 8px; }"
            "QScrollBar::handle:vertical { background: #AAAAAA; border-radius: 4px; }"
            "QHeaderView::section { background: #2D6A9F; color: white; "
            "  padding: 6px; font-weight: bold; border: 1px solid #1A5276; }"
        )

    # ── Data loading ───────────────────────────────────────────────────────────

    def _load_data(self):
        self._thread = QThread()
        self._loader = DataLoader()
        self._loader.moveToThread(self._thread)
        self._thread.started.connect(self._loader.run)
        self._loader.finished.connect(self._on_data_loaded)
        self._loader.error.connect(self._on_data_error)
        self._loader.finished.connect(self._thread.quit)
        self._loader.error.connect(self._thread.quit)
        self._thread.start()

    def _on_data_loaded(self, projects, invoices, debt_summaries, yearly_summary):
        self.projects = projects
        self.invoices = invoices
        self.debt_summaries = debt_summaries
        self.yearly_summary = yearly_summary

        self._create_pages()
        self._navigate("dashboard")

    def _on_data_error(self, message: str):
        QMessageBox.critical(
            self, "Data Load Error",
            f"Failed to load data:\n{message}\n\n"
            "Check that the Excel files exist in the 'data/' directory.",
        )
        # Still create pages (they will show empty state)
        self._create_pages()
        self._navigate("dashboard")

    # ── Page management ────────────────────────────────────────────────────────

    def _create_pages(self):
        from ui.dashboard_page import DashboardPage
        from ui.projects_page import ProjectsPage
        from ui.invoices_page import InvoicesPage
        from ui.monthly_invoice_page import MonthlyInvoicePage
        from ui.settings_page import SettingsPage

        self._pages["dashboard"] = DashboardPage(
            self.projects, self.invoices, self.debt_summaries, self.yearly_summary
        )
        self._pages["projects"] = ProjectsPage(self.projects, self.invoices)
        self._pages["invoices"] = InvoicesPage(
            self.invoices, self.projects, self.debt_summaries
        )
        self._pages["monthly_invoice"] = MonthlyInvoicePage(self.projects)
        self._pages["settings"] = SettingsPage()

        for page in self._pages.values():
            self._stack.addWidget(page)

    def _navigate(self, page_key: str):
        if page_key not in self._pages:
            return

        self._current_page = page_key
        self._stack.setCurrentWidget(self._pages[page_key])

        # Update sidebar button styles
        for key, btn in self._nav_buttons.items():
            btn.setStyleSheet(self._nav_btn_style(active=(key == page_key)))

    def refresh_data(self):
        """Reload data from disk and refresh all pages."""
        # Remove existing pages
        for page in self._pages.values():
            self._stack.removeWidget(page)
            page.deleteLater()
        self._pages.clear()
        self._load_data()
