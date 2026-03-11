"""Invoice details page – inline editing, paid/unpaid toggle, debt summary."""
import logging
from typing import List, Optional

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QFrame,
    QTableWidget, QTableWidgetItem, QHeaderView, QLineEdit,
    QPushButton, QComboBox, QGroupBox, QSplitter,
    QAbstractItemView, QMessageBox, QMenu,
)
from PySide6.QtCore import Qt, QPoint
from PySide6.QtGui import QColor, QFont, QAction

from config.settings import MONTH_ORDER
from models.project import Project
from models.invoice import Invoice, DebtSummary

logger = logging.getLogger(__name__)

# Invoice table columns
# (header, attr, editable)
INV_COLUMNS = [
    ("Invoice #",       "invoice_number",   True),
    ("Project Name",    "project_name",     False),
    ("Maint. Year",     "maintenance_year", True),
    ("Amount (€)",      "payment_amount",   True),
    ("Cameras",         "cameras_number",   True),
    ("Payment Date",    "payment_date",     True),
    ("Paid",            "paid",             True),
    ("Year",            "year",             True),
]
EDITABLE_COLS = {i for i, (_, _, ed) in enumerate(INV_COLUMNS) if ed}
NON_EDITABLE_COLS = {i for i, (_, _, ed) in enumerate(INV_COLUMNS) if not ed}


class InvoicesPage(QWidget):
    """Page showing invoice details and debt summary, both editable."""

    def __init__(
        self,
        invoices: List[Invoice],
        projects: List[Project],
        debt_summaries: List[DebtSummary],
        parent=None,
    ):
        super().__init__(parent)
        self._invoices = invoices
        self._projects = projects
        self._debt_summaries = debt_summaries

        self._dirty = False
        self._updating_table = False
        self._displayed_invoices: List[Invoice] = []

        self._build_ui()
        self._populate_debt_table()
        self._populate_invoice_table()

    # ── UI construction ────────────────────────────────────────────────────────

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(12)

        # Title + save button
        title_row = QHBoxLayout()
        title = QLabel("Invoice Details")
        title.setStyleSheet("font-size: 24px; font-weight: bold; color: #2C3E50;")
        title_row.addWidget(title)
        title_row.addStretch()

        self._dirty_label = QLabel()
        self._dirty_label.setStyleSheet("color: #E67E22; font-weight: 600;")
        title_row.addWidget(self._dirty_label)

        self._save_btn = QPushButton("Save to Excel")
        self._save_btn.setFixedHeight(36)
        self._save_btn.setEnabled(False)
        self._save_btn.setStyleSheet(
            "QPushButton { background: #27AE60; color: white; border-radius: 4px; "
            "padding: 0 14px; font-weight: 600; }"
            "QPushButton:hover { background: #1E8449; }"
            "QPushButton:disabled { background: #AAA; }"
        )
        self._save_btn.clicked.connect(self._save_to_excel)
        title_row.addWidget(self._save_btn)
        layout.addLayout(title_row)

        # Edit hint
        hint = QLabel(
            "Double-click a cell to edit.  Right-click a row to quickly mark as Paid / Unpaid / Cancelled."
        )
        hint.setStyleSheet(
            "background: #EBF5FB; color: #1A5276; border: 1px solid #AED6F1; "
            "border-radius: 4px; padding: 6px 12px; font-size: 12px;"
        )
        layout.addWidget(hint)

        # ── Summary row ────────────────────────────────────────────────────────
        summary_frame = QFrame()
        summary_frame.setStyleSheet(
            "QFrame { background: white; border: 1px solid #DEE2E6; border-radius: 8px; }"
        )
        sl = QHBoxLayout(summary_frame)
        sl.setContentsMargins(20, 10, 20, 10)
        sl.setSpacing(30)

        self._lbl_total_invoices = QLabel("Total: 0")
        self._lbl_total_paid     = QLabel("Paid: €0")
        self._lbl_total_unpaid   = QLabel("Unpaid: €0")
        self._lbl_total_cancelled = QLabel("Cancelled: €0")

        for lbl, color in [
            (self._lbl_total_invoices, "#2C3E50"),
            (self._lbl_total_paid,     "#27AE60"),
            (self._lbl_total_unpaid,   "#E74C3C"),
            (self._lbl_total_cancelled,"#888"),
        ]:
            lbl.setStyleSheet(f"font-size: 14px; font-weight: 600; color: {color};")
            sl.addWidget(lbl)
        sl.addStretch()
        layout.addWidget(summary_frame)

        # ── Splitter: invoices (top) + debt (bottom) ───────────────────────────
        splitter = QSplitter(Qt.Vertical)

        # Invoice table group
        inv_frame = QGroupBox("Invoices")
        inv_frame.setStyleSheet(
            "QGroupBox { font-weight: bold; font-size: 14px; color: #2C3E50; "
            "border: 1px solid #DEE2E6; border-radius: 8px; margin-top: 8px; }"
            "QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 4px; }"
        )
        inv_layout = QVBoxLayout(inv_frame)

        # Filter bar
        fr = QHBoxLayout()
        self._search_box = QLineEdit()
        self._search_box.setPlaceholderText("Search project name…")
        self._search_box.setFixedHeight(32)
        self._search_box.setStyleSheet(
            "QLineEdit { border: 1px solid #CED4DA; border-radius: 4px; padding: 0 8px; }"
        )
        self._search_box.textChanged.connect(self._apply_filter)
        fr.addWidget(self._search_box, stretch=2)

        fr.addWidget(QLabel("Year:"))
        self._year_cb = QComboBox()
        self._year_cb.setFixedHeight(32)
        years = sorted({inv.year for inv in self._invoices if inv.year}, reverse=True)
        self._year_cb.addItem("All")
        for y in years:
            self._year_cb.addItem(str(y))
        self._year_cb.currentTextChanged.connect(self._apply_filter)
        fr.addWidget(self._year_cb)

        fr.addWidget(QLabel("Maint. Year:"))
        self._my_cb = QComboBox()
        self._my_cb.setFixedHeight(32)
        my_vals = sorted({inv.maintenance_year for inv in self._invoices if inv.maintenance_year})
        self._my_cb.addItem("All")
        for v in my_vals:
            self._my_cb.addItem(v)
        self._my_cb.currentTextChanged.connect(self._apply_filter)
        fr.addWidget(self._my_cb)

        fr.addWidget(QLabel("Paid:"))
        self._paid_cb = QComboBox()
        self._paid_cb.setFixedHeight(32)
        self._paid_cb.addItems(["All", "Yes", "No", "cancelled"])
        self._paid_cb.currentTextChanged.connect(self._apply_filter)
        fr.addWidget(self._paid_cb)

        inv_layout.addLayout(fr)

        # Invoice table
        self._inv_table = QTableWidget()
        self._inv_table.setColumnCount(len(INV_COLUMNS))
        self._inv_table.setHorizontalHeaderLabels([c[0] for c in INV_COLUMNS])
        self._inv_table.setEditTriggers(QTableWidget.DoubleClicked | QTableWidget.AnyKeyPressed)
        self._inv_table.setSelectionBehavior(QTableWidget.SelectRows)
        self._inv_table.setSortingEnabled(True)
        self._inv_table.setAlternatingRowColors(True)
        self._inv_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self._inv_table.setContextMenuPolicy(Qt.CustomContextMenu)
        self._inv_table.customContextMenuRequested.connect(self._show_context_menu)
        self._inv_table.setStyleSheet(
            "QTableWidget { border: none; gridline-color: #DEE2E6; }"
            "QTableWidget::item:selected { background: #D6EAF8; color: #2C3E50; }"
        )
        self._inv_table.itemChanged.connect(self._on_item_changed)
        inv_layout.addWidget(self._inv_table)

        self._inv_count_label = QLabel()
        self._inv_count_label.setStyleSheet("color: #555; font-size: 12px;")
        inv_layout.addWidget(self._inv_count_label)

        splitter.addWidget(inv_frame)

        # Debt summary group
        debt_frame = QGroupBox("Debt Summary by Project")
        debt_frame.setStyleSheet(
            "QGroupBox { font-weight: bold; font-size: 14px; color: #2C3E50; "
            "border: 1px solid #DEE2E6; border-radius: 8px; margin-top: 8px; }"
            "QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 4px; }"
        )
        dl = QVBoxLayout(debt_frame)

        dfr = QHBoxLayout()
        self._debt_search = QLineEdit()
        self._debt_search.setPlaceholderText("Search project…")
        self._debt_search.setFixedHeight(30)
        self._debt_search.setStyleSheet(
            "QLineEdit { border: 1px solid #CED4DA; border-radius: 4px; padding: 0 8px; }"
        )
        self._debt_search.textChanged.connect(self._apply_debt_filter)
        dfr.addWidget(self._debt_search, stretch=1)

        self._only_debt_cb = QComboBox()
        self._only_debt_cb.addItems(["All Projects", "With Debt Only"])
        self._only_debt_cb.setFixedHeight(30)
        self._only_debt_cb.currentTextChanged.connect(self._apply_debt_filter)
        dfr.addWidget(self._only_debt_cb)
        dfr.addStretch()
        dl.addLayout(dfr)

        self._debt_table = QTableWidget()
        self._debt_table.setColumnCount(9)
        self._debt_table.setHorizontalHeaderLabels([
            "Project Name", "Country", "# Cams", "Month",
            "Install Year", "Status",
            "Expected (€)", "Paid (€)", "Unpaid (€)",
        ])
        self._debt_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self._debt_table.setSelectionBehavior(QTableWidget.SelectRows)
        self._debt_table.setSortingEnabled(True)
        self._debt_table.setAlternatingRowColors(True)
        self._debt_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self._debt_table.setStyleSheet(
            "QTableWidget { border: none; gridline-color: #DEE2E6; }"
            "QTableWidget::item:selected { background: #D6EAF8; color: #2C3E50; }"
        )
        dl.addWidget(self._debt_table)

        splitter.addWidget(debt_frame)
        splitter.setSizes([450, 280])
        layout.addWidget(splitter)

    # ── Populate invoice table ─────────────────────────────────────────────────

    def _populate_invoice_table(self, invoices: Optional[List[Invoice]] = None):
        if invoices is None:
            invoices = self._invoices
        self._displayed_invoices = invoices

        total_paid      = sum(inv.payment_amount for inv in invoices if inv.is_paid())
        total_unpaid    = sum(inv.payment_amount for inv in invoices if inv.is_unpaid())
        total_cancelled = sum(inv.payment_amount for inv in invoices if inv.is_cancelled())
        self._lbl_total_invoices.setText(f"Total: {len(invoices)}")
        self._lbl_total_paid.setText(f"Paid: €{total_paid:,.0f}")
        self._lbl_total_unpaid.setText(f"Unpaid: €{total_unpaid:,.0f}")
        self._lbl_total_cancelled.setText(f"Cancelled: €{total_cancelled:,.0f}")

        self._updating_table = True
        self._inv_table.setSortingEnabled(False)
        self._inv_table.setRowCount(len(invoices))

        for row, inv in enumerate(invoices):
            date_str = inv.payment_date.strftime("%Y-%m-%d") if inv.payment_date else ""
            cams_str = str(int(inv.cameras_number)) if inv.cameras_number else ""
            inv_no_str = str(int(inv.invoice_number)) if inv.invoice_number else ""

            values = [
                inv_no_str,
                inv.project_name,
                inv.maintenance_year,
                f"{inv.payment_amount:.2f}" if inv.payment_amount else "0.00",
                cams_str,
                date_str,
                inv.paid,
                str(inv.year) if inv.year else "",
            ]

            for col, (_, _, editable) in enumerate(INV_COLUMNS):
                item = QTableWidgetItem(values[col])
                item.setTextAlignment(Qt.AlignVCenter | (Qt.AlignLeft if col == 1 else Qt.AlignCenter))
                item.setData(Qt.UserRole, row)   # store index into _displayed_invoices

                if not editable:
                    item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                    item.setForeground(QColor("#666"))

                # Colour Paid column
                if col == 6:
                    if inv.is_paid():
                        item.setForeground(QColor("#27AE60"))
                        item.setFont(QFont("", -1, QFont.Bold))
                    elif inv.is_unpaid():
                        item.setForeground(QColor("#E74C3C"))
                        item.setFont(QFont("", -1, QFont.Bold))
                    else:
                        item.setForeground(QColor("#888"))

                self._inv_table.setItem(row, col, item)

        self._inv_table.setSortingEnabled(True)
        self._inv_table.resizeRowsToContents()
        self._updating_table = False
        self._inv_count_label.setText(
            f"Showing {len(invoices)} of {len(self._invoices)} invoices"
        )

    # ── Item changed (inline edit) ─────────────────────────────────────────────

    def _on_item_changed(self, item: QTableWidgetItem):
        if self._updating_table:
            return
        col = item.column()
        if col in NON_EDITABLE_COLS:
            return

        inv_row = item.data(Qt.UserRole)
        if inv_row is None or inv_row >= len(self._displayed_invoices):
            return
        inv = self._displayed_invoices[inv_row]

        new_val = item.text().strip()
        try:
            if col == 0:   # Invoice #
                inv.invoice_number = float(new_val) if new_val else None
            elif col == 2:   # Maintenance Year
                inv.maintenance_year = new_val
            elif col == 3: # Amount
                inv.payment_amount = float(new_val) if new_val else 0.0
            elif col == 4: # Cameras
                inv.cameras_number = float(new_val) if new_val else None
            elif col == 5: # Payment Date
                if new_val:
                    import datetime as dt
                    inv.payment_date = dt.datetime.strptime(new_val, "%Y-%m-%d")
                else:
                    inv.payment_date = None
            elif col == 6: # Paid
                val_lower = new_val.lower()
                if val_lower in ("yes", "y", "1", "true", "paid"):
                    inv.paid = "Yes"
                elif val_lower in ("no", "n", "0", "false", "unpaid"):
                    inv.paid = "No"
                elif val_lower in ("cancelled", "cancel", "c"):
                    inv.paid = "cancelled"
                else:
                    inv.paid = new_val
                # Refresh cell colour
                self._updating_table = True
                item.setText(inv.paid)
                if inv.is_paid():
                    item.setForeground(QColor("#27AE60"))
                    item.setFont(QFont("", -1, QFont.Bold))
                elif inv.is_unpaid():
                    item.setForeground(QColor("#E74C3C"))
                    item.setFont(QFont("", -1, QFont.Bold))
                else:
                    item.setForeground(QColor("#888"))
                    item.setFont(QFont())
                self._updating_table = False
            elif col == 7: # Year
                inv.year = int(new_val) if new_val else None
        except (ValueError, Exception) as e:
            logger.warning("Invalid value col %d: %s – %s", col, new_val, e)

        self._mark_dirty()

    def _mark_dirty(self):
        self._dirty = True
        self._save_btn.setEnabled(True)
        self._dirty_label.setText("Unsaved changes")

    # ── Right-click context menu ───────────────────────────────────────────────

    def _show_context_menu(self, pos: QPoint):
        row = self._inv_table.rowAt(pos.y())
        if row < 0 or row >= len(self._displayed_invoices):
            return

        inv = self._displayed_invoices[row]
        menu = QMenu(self)

        act_paid = QAction("Mark as Paid", self)
        act_paid.triggered.connect(lambda: self._set_paid_status(row, "Yes"))
        menu.addAction(act_paid)

        act_unpaid = QAction("Mark as Unpaid", self)
        act_unpaid.triggered.connect(lambda: self._set_paid_status(row, "No"))
        menu.addAction(act_unpaid)

        act_cancel = QAction("Mark as Cancelled", self)
        act_cancel.triggered.connect(lambda: self._set_paid_status(row, "cancelled"))
        menu.addAction(act_cancel)

        menu.exec(self._inv_table.viewport().mapToGlobal(pos))

    def _set_paid_status(self, row: int, status: str):
        if row >= len(self._displayed_invoices):
            return
        inv = self._displayed_invoices[row]
        inv.paid = status

        # Update the cell directly
        self._updating_table = True
        paid_item = self._inv_table.item(row, 6)
        if paid_item:
            paid_item.setText(status)
            if status == "Yes":
                paid_item.setForeground(QColor("#27AE60"))
                paid_item.setFont(QFont("", -1, QFont.Bold))
            elif status == "No":
                paid_item.setForeground(QColor("#E74C3C"))
                paid_item.setFont(QFont("", -1, QFont.Bold))
            else:
                paid_item.setForeground(QColor("#888"))
                paid_item.setFont(QFont())
        self._updating_table = False
        self._mark_dirty()

    # ── Populate debt table ────────────────────────────────────────────────────

    def _populate_debt_table(self, summaries: Optional[List[DebtSummary]] = None):
        if summaries is None:
            summaries = self._debt_summaries

        self._debt_table.setSortingEnabled(False)
        self._debt_table.setRowCount(len(summaries))

        for row, ds in enumerate(summaries):
            def _item(val, align=Qt.AlignLeft):
                item = QTableWidgetItem(str(val) if val not in (None, "") else "")
                item.setTextAlignment(align | Qt.AlignVCenter)
                item.setFlags(item.flags() & ~Qt.ItemIsEditable)
                return item

            self._debt_table.setItem(row, 0, _item(ds.project_name))
            self._debt_table.setItem(row, 1, _item(ds.country, Qt.AlignCenter))
            self._debt_table.setItem(row, 2, _item(ds.num_cams, Qt.AlignCenter))
            self._debt_table.setItem(row, 3, _item(ds.payment_month, Qt.AlignCenter))
            self._debt_table.setItem(row, 4, _item(ds.installation_year, Qt.AlignCenter))

            st = _item(ds.status, Qt.AlignCenter)
            st.setForeground(QColor("#27AE60") if str(ds.status).lower() == "active" else QColor("#E74C3C"))
            self._debt_table.setItem(row, 5, st)

            for col, (val, color) in enumerate([
                (ds.total_expected,  None),
                (ds.total_paid,      "#27AE60"),
                (ds.total_unpaid,    "#E74C3C" if ds.total_unpaid > 0 else None),
            ], start=6):
                cell = QTableWidgetItem()
                cell.setData(Qt.DisplayRole, round(val, 0))
                cell.setTextAlignment(Qt.AlignRight | Qt.AlignVCenter)
                cell.setFlags(cell.flags() & ~Qt.ItemIsEditable)
                if color:
                    cell.setForeground(QColor(color))
                if col == 8 and ds.total_unpaid > 0:
                    cell.setFont(QFont("", -1, QFont.Bold))
                self._debt_table.setItem(row, col, cell)

        self._debt_table.setSortingEnabled(True)
        self._debt_table.resizeRowsToContents()

    # ── Filtering ──────────────────────────────────────────────────────────────

    def _apply_filter(self, _=None):
        search = self._search_box.text().lower()
        year_text = self._year_cb.currentText()
        year = int(year_text) if year_text != "All" else None
        maint_year = self._my_cb.currentText()
        paid_filter = self._paid_cb.currentText()

        filtered = [
            inv for inv in self._invoices
            if (not search or search in inv.project_name.lower())
            and (year is None or inv.year == year)
            and (maint_year == "All" or inv.maintenance_year == maint_year)
            and (paid_filter == "All" or inv.paid.lower() == paid_filter.lower())
        ]
        self._populate_invoice_table(filtered)

    def _apply_debt_filter(self, _=None):
        search = self._debt_search.text().lower()
        only_debt = self._only_debt_cb.currentText() == "With Debt Only"

        filtered = [
            ds for ds in self._debt_summaries
            if (not search or search in ds.project_name.lower())
            and (not only_debt or ds.total_unpaid > 0)
        ]
        self._populate_debt_table(filtered)

    # ── Save to Excel ──────────────────────────────────────────────────────────

    def _save_to_excel(self):
        reply = QMessageBox.question(
            self, "Save to Excel",
            "This will overwrite the Invoice details sheet in:\n"
            "data/CaddyCheckProjectsInfo.xlsx\n\n"
            "Continue?",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return
        try:
            from services.excel_service import save_invoices_to_excel
            save_invoices_to_excel(self._invoices)
            self._dirty = False
            self._save_btn.setEnabled(False)
            self._dirty_label.setText("")
            QMessageBox.information(self, "Saved", "Invoices saved to Excel successfully.")
        except Exception as e:
            logger.exception("Failed to save invoices")
            QMessageBox.critical(self, "Save Error", f"Failed to save:\n{e}")
