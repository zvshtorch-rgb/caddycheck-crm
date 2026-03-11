"""Projects page – editable project list with revenue calculation and rate overrides."""
import datetime
import logging
from typing import List, Optional

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QFrame,
    QTableWidget, QTableWidgetItem, QHeaderView, QLineEdit,
    QPushButton, QComboBox, QGroupBox, QSplitter, QMessageBox,
    QAbstractItemView, QDialog, QFormLayout, QDoubleSpinBox,
    QDialogButtonBox, QSpinBox,
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont

from config.settings import MONTH_ORDER, RATE_Y1_PER_CAM, RATE_Y2_PLUS_PER_CAM
from models.project import Project
from models.invoice import Invoice

logger = logging.getLogger(__name__)
CURRENT_YEAR = datetime.datetime.now().year

# Columns shown in the table and whether they are editable
# (display_name, attr_name, editable, col_index)
PROJ_COLUMNS = [
    ("Project Name",    "project_name",     False,  0),
    ("Country",         "country",          True,   1),
    ("# Cams",          "num_cams",         True,   2),
    ("Payment Month",   "payment_month",    True,   3),
    ("Install Year",    "installation_year",True,   4),
    ("Activation Date", "activation_date",  True,   5),
    ("Status",          "status",           True,   6),
    ("License EOP",     "license_eop",      True,   7),
    ("Y1 Rate (€)",     "_y1_rate",         False,  8),   # computed, edited via dialog
    ("Y2+ Rate (€)",    "_y2_rate",         False,  9),
]
EDITABLE_COLS = {col[3] for col in PROJ_COLUMNS if col[2]}
NON_EDITABLE_COLS = {col[3] for col in PROJ_COLUMNS if not col[2]}


class RateOverrideDialog(QDialog):
    """Dialog to set per-project Y1/Y2+ rate overrides."""

    def __init__(self, project: Project, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Rate Override – {project.project_name}")
        self.setFixedWidth(380)

        from config.settings import RATE_Y1_PER_CAM, RATE_Y2_PLUS_PER_CAM
        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        info = QLabel(
            f"<b>{project.project_name}</b><br>"
            f"Override the per-camera rates for this project only.<br>"
            f"Leave blank to use the global default."
        )
        info.setWordWrap(True)
        layout.addWidget(info)

        form = QFormLayout()
        form.setSpacing(10)

        self._y1 = QDoubleSpinBox()
        self._y1.setRange(0, 99999)
        self._y1.setDecimals(2)
        self._y1.setSpecialValueText("Global default")
        self._y1.setValue(
            project.rate_y1_override if project.rate_y1_override is not None
            else RATE_Y1_PER_CAM
        )
        form.addRow(f"Y1 Rate (default €{RATE_Y1_PER_CAM}):", self._y1)

        self._y2 = QDoubleSpinBox()
        self._y2.setRange(0, 99999)
        self._y2.setDecimals(2)
        self._y2.setSpecialValueText("Global default")
        self._y2.setValue(
            project.rate_y2_override if project.rate_y2_override is not None
            else RATE_Y2_PLUS_PER_CAM
        )
        form.addRow(f"Y2+ Rate (default €{RATE_Y2_PLUS_PER_CAM}):", self._y2)
        layout.addLayout(form)

        reset_btn = QPushButton("Reset to Global Defaults")
        reset_btn.clicked.connect(self._reset)
        layout.addWidget(reset_btn)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _reset(self):
        from config.settings import RATE_Y1_PER_CAM, RATE_Y2_PLUS_PER_CAM
        self._y1.setValue(RATE_Y1_PER_CAM)
        self._y2.setValue(RATE_Y2_PLUS_PER_CAM)

    def get_values(self):
        """Return (y1_rate, y2_rate). None if same as global default."""
        from config.settings import RATE_Y1_PER_CAM, RATE_Y2_PLUS_PER_CAM
        y1 = self._y1.value()
        y2 = self._y2.value()
        return (
            y1 if y1 != RATE_Y1_PER_CAM else None,
            y2 if y2 != RATE_Y2_PLUS_PER_CAM else None,
        )


class ProjectsPage(QWidget):
    """Page listing all projects with inline editing and revenue breakdown."""

    def __init__(self, projects: List[Project], invoices: List[Invoice], parent=None):
        super().__init__(parent)
        self._projects = projects
        self._invoices = invoices
        self._selected_project: Optional[Project] = None
        self._dirty = False          # unsaved changes flag
        self._updating_table = False  # guard against recursive itemChanged

        from collections import defaultdict
        self._inv_map = defaultdict(list)
        for inv in invoices:
            self._inv_map[inv.project_name.lower().strip()].append(inv)

        self._build_ui()
        self._populate_table()

    # ── UI construction ────────────────────────────────────────────────────────

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)

        # Title + action buttons
        title_row = QHBoxLayout()
        title = QLabel("Projects")
        title.setStyleSheet("font-size: 24px; font-weight: bold; color: #2C3E50;")
        title_row.addWidget(title)
        title_row.addStretch()

        self._dirty_label = QLabel()
        self._dirty_label.setStyleSheet("color: #E67E22; font-weight: 600;")
        title_row.addWidget(self._dirty_label)

        self._override_btn = QPushButton("Override Rate…")
        self._override_btn.setFixedHeight(36)
        self._override_btn.setEnabled(False)
        self._override_btn.setStyleSheet(
            "QPushButton { background: #8E44AD; color: white; border-radius: 4px; padding: 0 14px; }"
            "QPushButton:hover { background: #6C3483; }"
            "QPushButton:disabled { background: #AAA; }"
        )
        self._override_btn.clicked.connect(self._open_override_dialog)
        title_row.addWidget(self._override_btn)

        self._save_btn = QPushButton("Save to Excel")
        self._save_btn.setFixedHeight(36)
        self._save_btn.setEnabled(False)
        self._save_btn.setStyleSheet(
            "QPushButton { background: #27AE60; color: white; border-radius: 4px; padding: 0 14px; font-weight: 600; }"
            "QPushButton:hover { background: #1E8449; }"
            "QPushButton:disabled { background: #AAA; }"
        )
        self._save_btn.clicked.connect(self._save_to_excel)
        title_row.addWidget(self._save_btn)

        layout.addLayout(title_row)

        # Edit hint banner
        hint = QLabel(
            "Double-click any highlighted cell to edit.  "
            "Changes are saved to Excel only when you click Save to Excel."
        )
        hint.setStyleSheet(
            "background: #EBF5FB; color: #1A5276; border: 1px solid #AED6F1; "
            "border-radius: 4px; padding: 6px 12px; font-size: 12px;"
        )
        layout.addWidget(hint)

        # Search / filter row
        filter_row = QHBoxLayout()
        self._search_box = QLineEdit()
        self._search_box.setPlaceholderText("Search projects…")
        self._search_box.setFixedHeight(34)
        self._search_box.setStyleSheet(
            "QLineEdit { border: 1px solid #CED4DA; border-radius: 4px; padding: 0 8px; }"
        )
        self._search_box.textChanged.connect(self._apply_filter)
        filter_row.addWidget(self._search_box, stretch=2)

        filter_row.addWidget(QLabel("Country:"))
        self._country_cb = QComboBox()
        countries = sorted({p.country for p in self._projects if p.country})
        self._country_cb.addItem("All")
        for c in countries:
            self._country_cb.addItem(c)
        self._country_cb.currentTextChanged.connect(self._apply_filter)
        filter_row.addWidget(self._country_cb)

        filter_row.addWidget(QLabel("Month:"))
        self._month_cb = QComboBox()
        self._month_cb.addItem("All")
        for m in MONTH_ORDER:
            self._month_cb.addItem(m)
        self._month_cb.currentTextChanged.connect(self._apply_filter)
        filter_row.addWidget(self._month_cb)

        filter_row.addWidget(QLabel("Status:"))
        self._status_cb = QComboBox()
        self._status_cb.addItems(["All", "Active", "Offline"])
        self._status_cb.currentTextChanged.connect(self._apply_filter)
        filter_row.addWidget(self._status_cb)

        filter_row.addStretch()
        layout.addLayout(filter_row)

        # Splitter
        splitter = QSplitter(Qt.Vertical)

        # Project table
        table_frame = QFrame()
        table_frame.setStyleSheet(
            "QFrame { background: white; border: 1px solid #DEE2E6; border-radius: 8px; }"
        )
        tf_layout = QVBoxLayout(table_frame)
        tf_layout.setContentsMargins(0, 0, 0, 0)

        self._table = QTableWidget()
        self._table.setColumnCount(len(PROJ_COLUMNS))
        self._table.setHorizontalHeaderLabels([c[0] for c in PROJ_COLUMNS])
        self._table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.AnyKeyPressed)
        self._table.setSelectionBehavior(QTableWidget.SelectRows)
        self._table.setSelectionMode(QAbstractItemView.SingleSelection)
        self._table.setAlternatingRowColors(True)
        self._table.setSortingEnabled(True)
        self._table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self._table.setStyleSheet(
            "QTableWidget { border: none; gridline-color: #DEE2E6; }"
            "QTableWidget::item:selected { background: #D6EAF8; color: #2C3E50; }"
        )
        self._table.itemSelectionChanged.connect(self._on_selection_changed)
        self._table.itemChanged.connect(self._on_item_changed)
        tf_layout.addWidget(self._table)
        splitter.addWidget(table_frame)

        # Revenue detail panel
        detail_frame = QFrame()
        detail_frame.setStyleSheet(
            "QFrame { background: white; border: 1px solid #DEE2E6; border-radius: 8px; }"
        )
        dl = QVBoxLayout(detail_frame)
        dl.setContentsMargins(16, 12, 16, 12)

        detail_title = QLabel("Expected Revenue by Maintenance Year")
        detail_title.setStyleSheet("font-size: 15px; font-weight: bold; color: #2C3E50;")
        dl.addWidget(detail_title)

        ctrl_row = QHBoxLayout()
        ctrl_row.addWidget(QLabel("From:"))
        self._from_year_cb = QComboBox()
        self._to_year_cb = QComboBox()
        for yr in range(2013, CURRENT_YEAR + 6):
            self._from_year_cb.addItem(str(yr))
            self._to_year_cb.addItem(str(yr))
        self._from_year_cb.setCurrentText("2024")
        self._to_year_cb.setCurrentText(str(CURRENT_YEAR + 2))
        ctrl_row.addWidget(self._from_year_cb)
        ctrl_row.addWidget(QLabel("To:"))
        ctrl_row.addWidget(self._to_year_cb)
        calc_btn = QPushButton("Calculate")
        calc_btn.setFixedHeight(30)
        calc_btn.setStyleSheet(
            "QPushButton { background: #2D6A9F; color: white; border-radius: 4px; padding: 0 12px; }"
            "QPushButton:hover { background: #1A5276; }"
        )
        calc_btn.clicked.connect(self._calc_revenue)
        ctrl_row.addWidget(calc_btn)
        ctrl_row.addStretch()
        dl.addLayout(ctrl_row)

        self._rev_table = QTableWidget()
        self._rev_table.setColumnCount(5)
        self._rev_table.setHorizontalHeaderLabels(
            ["Year", "Maint. Year", "Rate/Cam (€)", "# Cams", "Expected (€)"]
        )
        self._rev_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self._rev_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self._rev_table.setAlternatingRowColors(True)
        self._rev_table.setStyleSheet("QTableWidget { border: none; gridline-color: #DEE2E6; }")
        dl.addWidget(self._rev_table)

        splitter.addWidget(detail_frame)
        splitter.setSizes([420, 220])
        layout.addWidget(splitter)

        self._summary_label = QLabel()
        self._summary_label.setStyleSheet("color: #555; font-size: 12px;")
        layout.addWidget(self._summary_label)

    # ── Table population ───────────────────────────────────────────────────────

    def _populate_table(self, projects: Optional[List[Project]] = None):
        if projects is None:
            projects = self._projects

        self._updating_table = True
        self._table.setSortingEnabled(False)
        self._table.setRowCount(len(projects))

        for row, proj in enumerate(projects):
            act = proj.activation_date.strftime("%Y-%m-%d") if proj.activation_date else ""
            eop = proj.license_eop.strftime("%Y-%m-%d") if proj.license_eop else ""
            y1_rate = proj.rate_y1_override if proj.rate_y1_override is not None else RATE_Y1_PER_CAM
            y2_rate = proj.rate_y2_override if proj.rate_y2_override is not None else RATE_Y2_PLUS_PER_CAM

            values = [
                proj.project_name,
                proj.country,
                proj.num_cams,
                proj.payment_month,
                proj.installation_year if proj.installation_year else "",
                act,
                proj.status,
                eop,
                f"€{y1_rate:.0f}" + (" *" if proj.rate_y1_override is not None else ""),
                f"€{y2_rate:.0f}" + (" *" if proj.rate_y2_override is not None else ""),
            ]

            for col_idx, (_, _, editable, col_pos) in enumerate(PROJ_COLUMNS):
                item = QTableWidgetItem(str(values[col_idx]) if values[col_idx] is not None else "")
                item.setTextAlignment(Qt.AlignVCenter | (Qt.AlignLeft if col_idx == 0 else Qt.AlignCenter))
                item.setData(Qt.UserRole, proj.project_name)  # store key

                if not editable:
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    item.setForeground(QColor("#666"))

                # Color status cell
                if col_idx == 6:
                    item.setForeground(QColor("#27AE60") if proj.is_active() else QColor("#E74C3C"))

                # Highlight overridden rates
                if col_idx in (8, 9) and (proj.rate_y1_override is not None or proj.rate_y2_override is not None):
                    item.setBackground(QColor("#FEF9E7"))
                    item.setForeground(QColor("#D68910"))

                self._table.setItem(row, col_pos, item)

        self._table.setSortingEnabled(True)
        self._table.resizeRowsToContents()
        self._updating_table = False
        self._summary_label.setText(
            f"Showing {len(projects)} of {len(self._projects)} projects  |  "
            f"Double-click a cell to edit"
        )

    # ── Item changed (inline edit) ─────────────────────────────────────────────

    def _on_item_changed(self, item: QTableWidgetItem):
        if self._updating_table:
            return
        col = item.column()
        if col in NON_EDITABLE_COLS:
            return

        proj_name = item.data(Qt.UserRole)
        proj = next((p for p in self._projects if p.project_name == proj_name), None)
        if proj is None:
            return

        new_val = item.text().strip()

        try:
            if col == 1:   # Country
                proj.country = new_val
            elif col == 2: # # Cams
                proj.num_cams = int(new_val) if new_val else 0
            elif col == 3: # Payment Month
                from config.settings import normalize_month
                proj.payment_month = normalize_month(new_val)
                # Update display to normalized form
                self._updating_table = True
                item.setText(proj.payment_month)
                self._updating_table = False
            elif col == 4: # Install Year
                proj.installation_year = int(new_val) if new_val else None
            elif col == 5: # Activation Date
                if new_val:
                    import datetime as dt
                    proj.activation_date = dt.datetime.strptime(new_val, "%Y-%m-%d")
                else:
                    proj.activation_date = None
            elif col == 6: # Status
                proj.status = new_val
            elif col == 7: # License EOP
                if new_val:
                    import datetime as dt
                    proj.license_eop = dt.datetime.strptime(new_val, "%Y-%m-%d")
                else:
                    proj.license_eop = None
        except (ValueError, Exception) as e:
            logger.warning("Invalid value for col %d: %s – %s", col, new_val, e)

        self._mark_dirty()

    def _mark_dirty(self):
        self._dirty = True
        self._save_btn.setEnabled(True)
        self._dirty_label.setText("Unsaved changes")

    # ── Selection ──────────────────────────────────────────────────────────────

    def _on_selection_changed(self):
        rows = self._table.selectedItems()
        if not rows:
            self._selected_project = None
            self._override_btn.setEnabled(False)
            self._rev_table.setRowCount(0)
            return
        row = self._table.currentRow()
        item = self._table.item(row, 0)
        if not item:
            return
        proj_name = item.data(Qt.UserRole) or item.text()
        self._selected_project = next(
            (p for p in self._projects if p.project_name == proj_name), None
        )
        self._override_btn.setEnabled(self._selected_project is not None)
        self._calc_revenue()

    # ── Filter ─────────────────────────────────────────────────────────────────

    def _apply_filter(self, _=None):
        search = self._search_box.text().lower()
        country = self._country_cb.currentText()
        month = self._month_cb.currentText()
        status = self._status_cb.currentText()

        filtered = [
            p for p in self._projects
            if (not search or search in p.project_name.lower())
            and (country == "All" or p.country == country)
            and (month == "All" or p.payment_month == month)
            and (status == "All"
                 or (status == "Active" and p.is_active())
                 or (status == "Offline" and not p.is_active()))
        ]
        self._populate_table(filtered)

    # ── Revenue breakdown ──────────────────────────────────────────────────────

    def _calc_revenue(self):
        if not self._selected_project:
            return
        proj = self._selected_project
        try:
            from_yr = int(self._from_year_cb.currentText())
            to_yr = int(self._to_year_cb.currentText())
        except ValueError:
            return
        if from_yr > to_yr:
            from_yr, to_yr = to_yr, from_yr

        rows = [
            (yr, f"Y{proj.get_maintenance_year(yr)}", proj.get_rate(yr),
             proj.num_cams, proj.get_expected_amount(yr))
            for yr in range(from_yr, to_yr + 1)
            if proj.installation_year is None or yr >= proj.installation_year
        ]

        self._rev_table.setRowCount(len(rows) + (1 if rows else 0))
        total = 0.0
        for r, (yr, label, rate, cams, amount) in enumerate(rows):
            for c, val in enumerate([yr, label, f"€{rate:.0f}", cams, f"€{amount:,.0f}"]):
                cell = QTableWidgetItem(str(val))
                cell.setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)
                self._rev_table.setItem(r, c, cell)
            total += amount

        if rows:
            total_row = len(rows)
            for c, val in enumerate(["", "", "", "TOTAL", f"€{total:,.0f}"]):
                cell = QTableWidgetItem(val)
                cell.setTextAlignment(Qt.AlignCenter | Qt.AlignVCenter)
                cell.setFont(QFont("", -1, QFont.Bold))
                cell.setBackground(QColor("#D5F5E3"))
                self._rev_table.setItem(total_row, c, cell)

    # ── Override rate dialog ───────────────────────────────────────────────────

    def _open_override_dialog(self):
        if not self._selected_project:
            return
        dlg = RateOverrideDialog(self._selected_project, parent=self)
        if dlg.exec() == QDialog.Accepted:
            y1, y2 = dlg.get_values()
            self._selected_project.rate_y1_override = y1
            self._selected_project.rate_y2_override = y2

            # Persist to overrides file
            from config.settings import get_project_overrides, save_project_overrides
            overrides = get_project_overrides()
            key = self._selected_project.project_name.lower().strip()
            overrides[key] = {"y1_rate": y1, "y2_rate": y2}
            save_project_overrides(overrides)

            # Refresh table to show new rates
            self._apply_filter()
            self._calc_revenue()

    # ── Save to Excel ──────────────────────────────────────────────────────────

    def _save_to_excel(self):
        reply = QMessageBox.question(
            self, "Save to Excel",
            "This will overwrite the Projects overview sheet in:\n"
            "data/CaddyCheckProjectsInfo.xlsx\n\n"
            "Continue?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return
        try:
            from services.excel_service import save_projects_to_excel
            save_projects_to_excel(self._projects)
            self._dirty = False
            self._save_btn.setEnabled(False)
            self._dirty_label.setText("")
            QMessageBox.information(self, "Saved", "Projects saved to Excel successfully.")
        except Exception as e:
            logger.exception("Failed to save projects")
            QMessageBox.critical(self, "Save Error", f"Failed to save:\n{e}")
