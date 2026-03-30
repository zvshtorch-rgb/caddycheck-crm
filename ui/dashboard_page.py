"""Dashboard page – summary cards and filtered overview."""
import datetime
import logging
from typing import List, Optional

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QFrame,
    QGridLayout, QComboBox, QScrollArea, QPushButton,
    QSizePolicy, QGroupBox, QTableWidget, QTableWidgetItem,
    QHeaderView, QFileDialog, QMessageBox, QToolTip,
)
from PySide6.QtCore import Qt, QThread, Signal, QObject
from PySide6.QtGui import QFont, QColor, QCursor
from PySide6.QtCharts import (QChart, QChartView, QLineSeries, QBarSeries,
                               QBarSet, QBarCategoryAxis, QValueAxis, QDateTimeAxis)
from PySide6.QtCore import QPointF, QDate, QTime, QDateTime

from config.settings import MONTH_ORDER, normalize_month, OUTPUT_DIR
from models.project import Project
from models.invoice import Invoice, DebtSummary
from services.excel_service import get_yearly_summary, get_monthly_summary

logger = logging.getLogger(__name__)

CARD_COLORS = {
    "income":   ("#27AE60", "#EAFAF1"),
    "debt":     ("#E74C3C", "#FDEDEC"),
    "paid":     ("#2980B9", "#EBF5FB"),
    "projects": ("#8E44AD", "#F4ECF7"),
    "cameras":  ("#F39C12", "#FEF9E7"),
    "monthly":  ("#16A085", "#E8F8F5"),
    "yearly":   ("#2C3E50", "#EAECEE"),
}


class SummaryCard(QFrame):
    """A metric card widget."""

    def __init__(self, title: str, value: str, color_key: str = "income", parent=None):
        super().__init__(parent)
        accent, bg = CARD_COLORS.get(color_key, ("#2D6A9F", "#EBF5FB"))
        self.setFixedHeight(110)
        self.setMinimumWidth(160)
        self.setStyleSheet(
            f"QFrame {{ background: {bg}; border: 1px solid {accent}33; "
            f"border-left: 4px solid {accent}; border-radius: 8px; }}"
        )

        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 10, 14, 10)

        self._title_label = QLabel(title)
        self._title_label.setStyleSheet(f"color: {accent}; font-size: 12px; font-weight: 600; border: none;")
        layout.addWidget(self._title_label)

        self._value_label = QLabel(value)
        self._value_label.setStyleSheet(f"color: #2C3E50; font-size: 22px; font-weight: bold; border: none;")
        layout.addWidget(self._value_label)

        layout.addStretch()

    def update_value(self, value: str):
        self._value_label.setText(value)

    def update_title(self, title: str):
        self._title_label.setText(title)


class DashboardPage(QWidget):
    """Dashboard showing summary metrics with filters."""

    def __init__(
        self,
        projects: List[Project],
        invoices: List[Invoice],
        debt_summaries: List[DebtSummary],
        yearly_summary: dict,
        parent=None,
    ):
        super().__init__(parent)
        self._all_projects = projects
        self._all_invoices = invoices
        self._all_debt_summaries = debt_summaries
        self._yearly_summary = yearly_summary

        self._build_ui()
        self._init_chart_years()
        self._refresh()

    def _build_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(24, 20, 24, 20)
        main_layout.setSpacing(16)

        # ── Title bar ──────────────────────────────────────────────────────────
        title_row = QHBoxLayout()
        title = QLabel("Dashboard")
        title.setStyleSheet("font-size: 24px; font-weight: bold; color: #2C3E50;")
        title_row.addWidget(title)
        title_row.addStretch()

        export_btn = QPushButton("Export to Excel")
        export_btn.setFixedHeight(36)
        export_btn.setStyleSheet(
            "QPushButton { background: #2D6A9F; color: white; border-radius: 4px; "
            "padding: 0 16px; font-weight: 600; }"
            "QPushButton:hover { background: #1A5276; }"
        )
        export_btn.clicked.connect(self._export_excel)
        title_row.addWidget(export_btn)
        main_layout.addLayout(title_row)

        # ── Filters row ────────────────────────────────────────────────────────
        filter_frame = QFrame()
        filter_frame.setStyleSheet(
            "QFrame { background: white; border: 1px solid #DEE2E6; border-radius: 8px; }"
        )
        filter_layout = QHBoxLayout(filter_frame)
        filter_layout.setContentsMargins(16, 10, 16, 10)
        filter_layout.setSpacing(12)

        filter_layout.addWidget(QLabel("Filters:"))

        # Year filter
        filter_layout.addWidget(QLabel("Year:"))
        self._year_cb = QComboBox()
        self._year_cb.setFixedWidth(90)
        years = sorted({inv.year for inv in self._all_invoices if inv.year}, reverse=True)
        self._year_cb.addItem("All")
        for y in years:
            self._year_cb.addItem(str(y))
        self._year_cb.currentTextChanged.connect(self._on_filter_changed)
        filter_layout.addWidget(self._year_cb)

        # Month filter
        filter_layout.addWidget(QLabel("Month:"))
        self._month_cb = QComboBox()
        self._month_cb.setFixedWidth(110)
        self._month_cb.addItem("All")
        for m in MONTH_ORDER:
            self._month_cb.addItem(m)
        self._month_cb.currentTextChanged.connect(self._on_filter_changed)
        filter_layout.addWidget(self._month_cb)

        # Country filter
        filter_layout.addWidget(QLabel("Country:"))
        self._country_cb = QComboBox()
        self._country_cb.setFixedWidth(90)
        countries = sorted({p.country for p in self._all_projects if p.country})
        self._country_cb.addItem("All")
        for c in countries:
            self._country_cb.addItem(c)
        self._country_cb.currentTextChanged.connect(self._on_filter_changed)
        filter_layout.addWidget(self._country_cb)

        # Paid/unpaid filter
        filter_layout.addWidget(QLabel("Status:"))
        self._paid_cb = QComboBox()
        self._paid_cb.setFixedWidth(100)
        self._paid_cb.addItems(["All", "Paid", "Unpaid", "Cancelled"])
        self._paid_cb.currentTextChanged.connect(self._on_filter_changed)
        filter_layout.addWidget(self._paid_cb)

        filter_layout.addStretch()
        main_layout.addWidget(filter_frame)

        # ── Summary cards ──────────────────────────────────────────────────────
        cards_row = QHBoxLayout()
        cards_row.setSpacing(12)

        self._card_income    = SummaryCard("Total Income", "€0", "income")
        self._card_paid      = SummaryCard("Total Paid", "€0", "paid")
        self._card_debt      = SummaryCard("Total Debt", "€0", "debt")
        self._card_monthly   = SummaryCard("Monthly Income", "€0", "monthly")
        self._card_yearly    = SummaryCard("Yearly Income", "€0", "yearly")
        self._card_projects  = SummaryCard("Active Projects", "0", "projects")
        self._card_cameras   = SummaryCard("Total Cameras", "0", "cameras")

        for card in [
            self._card_income, self._card_paid, self._card_debt,
            self._card_monthly, self._card_yearly,
            self._card_projects, self._card_cameras,
        ]:
            card.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
            cards_row.addWidget(card)

        main_layout.addLayout(cards_row)

        # ── Cameras per project table ──────────────────────────────────────────
        cam_group = QGroupBox("Cameras by Project")
        cam_group.setStyleSheet(
            "QGroupBox { font-weight: bold; font-size: 14px; color: #2C3E50; "
            "border: 1px solid #DEE2E6; border-radius: 8px; margin-top: 8px; }"
            "QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 4px; }"
        )
        cam_layout = QVBoxLayout(cam_group)

        self._cam_table = QTableWidget()
        self._cam_table.setColumnCount(6)
        self._cam_table.setHorizontalHeaderLabels([
            "Project Name", "Country", "# Cams", "Payment Month",
            "Install Year", "Status",
        ])
        self._cam_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self._cam_table.setSelectionBehavior(QTableWidget.SelectRows)
        self._cam_table.setAlternatingRowColors(True)
        self._cam_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self._cam_table.setStyleSheet(
            "QTableWidget { border: none; gridline-color: #DEE2E6; }"
            "QTableWidget::item:selected { background: #D6EAF8; color: #2C3E50; }"
        )
        cam_layout.addWidget(self._cam_table)
        main_layout.addWidget(cam_group)

        # ── Charts ─────────────────────────────────────────────────────────────
        chart_group = QGroupBox("Trends")
        chart_group.setStyleSheet(
            "QGroupBox { font-weight: bold; font-size: 14px; color: #2C3E50; "
            "border: 1px solid #DEE2E6; border-radius: 8px; margin-top: 8px; }"
            "QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 4px; }"
        )
        chart_layout = QVBoxLayout(chart_group)

        # Chart controls
        chart_ctrl = QHBoxLayout()
        chart_ctrl.addWidget(QLabel("Show:"))
        self._chart_metric_cb = QComboBox()
        self._chart_metric_cb.addItems(["Income (Paid)", "Income (All)", "Active Projects", "Cameras"])
        self._chart_metric_cb.currentTextChanged.connect(self._update_chart)
        chart_ctrl.addWidget(self._chart_metric_cb)

        chart_ctrl.addWidget(QLabel("Resolution:"))
        self._chart_res_cb = QComboBox()
        self._chart_res_cb.addItems(["Yearly", "Monthly"])
        self._chart_res_cb.currentTextChanged.connect(self._update_chart)
        chart_ctrl.addWidget(self._chart_res_cb)

        chart_ctrl.addWidget(QLabel("From Year:"))
        self._chart_from_cb = QComboBox()
        chart_ctrl.addWidget(self._chart_from_cb)

        chart_ctrl.addWidget(QLabel("To Year:"))
        self._chart_to_cb = QComboBox()
        chart_ctrl.addWidget(self._chart_to_cb)

        chart_ctrl.addStretch()
        chart_layout.addLayout(chart_ctrl)

        # Chart view
        self._chart = QChart()
        self._chart.setAnimationOptions(QChart.SeriesAnimations)
        self._chart.legend().setVisible(False)
        self._chart_view = QChartView(self._chart)
        self._chart_view.setMinimumHeight(280)
        self._chart_view.setRenderHint(self._chart_view.renderHints() | self._chart_view.renderHints())
        chart_layout.addWidget(self._chart_view)

        main_layout.addWidget(chart_group)

    # ── Filtering and refresh ──────────────────────────────────────────────────

    def _on_filter_changed(self, _=None):
        self._refresh()

    def _get_filters(self):
        year_text = self._year_cb.currentText()
        year = int(year_text) if year_text != "All" else None
        month = self._month_cb.currentText()
        if month == "All":
            month = None
        country = self._country_cb.currentText()
        if country == "All":
            country = None
        paid_filter = self._paid_cb.currentText()
        if paid_filter == "All":
            paid_filter = None
        return year, month, country, paid_filter

    def _filter_invoices(self) -> List[Invoice]:
        year, month, country, paid_filter = self._get_filters()

        # Map month name to number
        month_num = None
        if month:
            month_num = MONTH_ORDER.index(month) + 1 if month in MONTH_ORDER else None

        # Build project country map
        proj_country = {p.project_name.lower(): p.country for p in self._all_projects}

        result = []
        for inv in self._all_invoices:
            if year and inv.year != year:
                continue
            if month_num and inv.payment_date and inv.payment_date.month != month_num:
                continue
            if country:
                pc = proj_country.get(inv.project_name.lower().strip(), "")
                if pc != country:
                    continue
            if paid_filter == "Paid" and not inv.is_paid():
                continue
            if paid_filter == "Unpaid" and not inv.is_unpaid():
                continue
            if paid_filter == "Cancelled" and not inv.is_cancelled():
                continue
            result.append(inv)
        return result

    def _filter_projects(self) -> List[Project]:
        _, _, country, _ = self._get_filters()
        if country:
            return [p for p in self._all_projects if p.country == country]
        return self._all_projects

    def _refresh(self):
        year_text = self._year_cb.currentText()
        selected_year = int(year_text) if year_text != "All" else None

        filtered_invoices = self._filter_invoices()
        filtered_projects = self._filter_projects()

        # Totals
        total_paid = sum(inv.payment_amount for inv in filtered_invoices if inv.is_paid())
        total_unpaid = sum(inv.payment_amount for inv in filtered_invoices if inv.is_unpaid())
        total_income = total_paid + total_unpaid

        self._card_income.update_value(f"€{total_income:,.0f}")
        self._card_paid.update_value(f"€{total_paid:,.0f}")
        self._card_debt.update_value(f"€{total_unpaid:,.0f}")

        # Monthly/yearly from paid invoices
        current_month = datetime.datetime.now().month
        if selected_year:
            ref_year = selected_year
        else:
            # Use the most recent year that has paid invoices
            paid_years = [inv.year for inv in filtered_invoices if inv.is_paid() and inv.year]
            ref_year = max(paid_years) if paid_years else datetime.datetime.now().year

        monthly = get_monthly_summary(filtered_invoices, year=ref_year)
        monthly_val = monthly.get(current_month, 0.0)
        yearly_val = sum(
            inv.payment_amount for inv in filtered_invoices
            if inv.is_paid() and inv.year == ref_year
        )

        self._card_monthly.update_title(f"Monthly Income ({ref_year})")
        self._card_yearly.update_title(f"Yearly Income ({ref_year})")
        self._card_monthly.update_value(f"€{monthly_val:,.0f}")
        self._card_yearly.update_value(f"€{yearly_val:,.0f}")

        # Active projects
        active_count = sum(1 for p in filtered_projects if p.is_active())
        self._card_projects.update_value(str(active_count))

        # Total cameras
        total_cams = sum(p.num_cams for p in filtered_projects)
        self._card_cameras.update_value(str(total_cams))

        # Camera table
        self._populate_cam_table(filtered_projects)

    def _populate_cam_table(self, projects: List[Project]):
        sorted_projects = sorted(projects, key=lambda p: (0 if p.is_active() else 1, p.project_name))
        self._cam_table.setRowCount(len(sorted_projects))

        for row, proj in enumerate(sorted_projects):
            def _item(text, align=Qt.AlignLeft):
                item = QTableWidgetItem(str(text) if text else "")
                item.setTextAlignment(align | Qt.AlignVCenter)
                return item

            self._cam_table.setItem(row, 0, _item(proj.project_name))
            self._cam_table.setItem(row, 1, _item(proj.country, Qt.AlignCenter))
            self._cam_table.setItem(row, 2, _item(proj.num_cams, Qt.AlignCenter))
            self._cam_table.setItem(row, 3, _item(proj.payment_month, Qt.AlignCenter))
            self._cam_table.setItem(row, 4, _item(proj.installation_year, Qt.AlignCenter))

            status_item = _item(proj.status, Qt.AlignCenter)
            if proj.is_active():
                status_item.setForeground(QColor("#27AE60"))
                status_item.setFont(QFont("", -1, QFont.Bold))
            else:
                status_item.setForeground(QColor("#E74C3C"))
            self._cam_table.setItem(row, 5, status_item)

        self._cam_table.resizeRowsToContents()

    # ── Chart ──────────────────────────────────────────────────────────────────

    def _init_chart_years(self):
        all_years = sorted({inv.year for inv in self._all_invoices if inv.year})
        if not all_years:
            all_years = [datetime.datetime.now().year]
        for cb in (self._chart_from_cb, self._chart_to_cb):
            cb.blockSignals(True)
            for y in all_years:
                cb.addItem(str(int(y)))
            cb.blockSignals(False)
        self._chart_from_cb.setCurrentText(str(int(min(all_years))))
        self._chart_to_cb.setCurrentText(str(int(max(all_years))))
        self._chart_from_cb.currentTextChanged.connect(self._update_chart)
        self._chart_to_cb.currentTextChanged.connect(self._update_chart)
        self._update_chart()

    def _update_chart(self):
        try:
            from_yr = int(self._chart_from_cb.currentText())
            to_yr   = int(self._chart_to_cb.currentText())
        except (ValueError, AttributeError):
            return
        if from_yr > to_yr:
            from_yr, to_yr = to_yr, from_yr

        metric     = self._chart_metric_cb.currentText()
        resolution = self._chart_res_cb.currentText()
        invoices   = self._all_invoices
        projects   = self._all_projects

        self._chart.removeAllSeries()
        for ax in self._chart.axes():
            self._chart.removeAxis(ax)

        is_income = metric.startswith("Income")

        if resolution == "Yearly":
            # ── Bar chart ──────────────────────────────────────────────────────
            labels, values = [], []
            for yr in range(from_yr, to_yr + 1):
                labels.append(str(yr))
                if metric == "Income (Paid)":
                    v = sum(i.payment_amount for i in invoices if i.is_paid() and i.year == yr)
                elif metric == "Income (All)":
                    v = sum(i.payment_amount for i in invoices if i.year == yr)
                elif metric == "Active Projects":
                    v = sum(1 for p in projects if p.installation_year and p.installation_year <= yr and p.is_active())
                else:
                    v = sum(p.num_cams for p in projects if p.installation_year and p.installation_year <= yr and p.is_active())
                values.append(float(v))

            bar_set = QBarSet(metric)
            bar_set.setColor(QColor("#2980B9"))
            for v in values:
                bar_set.append(v)
            series = QBarSeries()
            series.append(bar_set)
            self._chart.addSeries(series)

            x_axis = QBarCategoryAxis()
            x_axis.append(labels)
            x_axis.setLabelsAngle(-45)
            self._chart.addAxis(x_axis, Qt.AlignBottom)
            series.attachAxis(x_axis)

            max_val = max(values) if values else 1.0
            y_axis = QValueAxis()
            y_axis.setRange(0, max_val * 1.15)
            y_axis.setLabelFormat("%.0f")
            y_axis.setTitleText("EUR" if is_income else "Count")
            self._chart.addAxis(y_axis, Qt.AlignLeft)
            series.attachAxis(y_axis)

            # Hover tooltip: show exact value
            unit = " EUR" if is_income else ""
            _labels = labels  # capture for closure
            _values = values
            def _bar_hovered(status, index, barset, _lbl=_labels, _val=_values, _u=unit):
                if status and 0 <= index < len(_lbl):
                    QToolTip.showText(
                        QCursor.pos(),
                        f"{_lbl[index]}: {_val[index]:,.0f}{_u}"
                    )
                else:
                    QToolTip.hideText()
            series.hovered.connect(_bar_hovered)

        else:
            # ── Line chart with QDateTimeAxis ─────────────────────────────────
            series = QLineSeries()
            pen = series.pen()
            pen.setColor(QColor("#2980B9"))
            pen.setWidth(2)
            series.setPen(pen)
            series.setPointsVisible(True)

            max_val = 0.0
            for yr in range(from_yr, to_yr + 1):
                for mo in range(1, 13):
                    if metric == "Income (Paid)":
                        v = sum(i.payment_amount for i in invoices
                                if i.is_paid() and i.payment_date
                                and i.payment_date.year == yr and i.payment_date.month == mo)
                    elif metric == "Income (All)":
                        v = sum(i.payment_amount for i in invoices
                                if i.payment_date
                                and i.payment_date.year == yr and i.payment_date.month == mo)
                    elif metric == "Active Projects":
                        v = sum(1 for p in projects
                                if p.installation_year and p.installation_year <= yr and p.is_active())
                    else:
                        v = sum(p.num_cams for p in projects
                                if p.installation_year and p.installation_year <= yr and p.is_active())
                    v = float(v)
                    if v > max_val:
                        max_val = v
                    dt = QDateTime(QDate(yr, mo, 1), QTime(0, 0))
                    series.append(dt.toMSecsSinceEpoch(), v)

            self._chart.addSeries(series)

            x_axis = QDateTimeAxis()
            x_axis.setFormat("MMM yyyy")
            x_axis.setLabelsAngle(-60)
            num_years = to_yr - from_yr + 1
            x_axis.setTickCount(min(num_years + 1, 15))
            self._chart.addAxis(x_axis, Qt.AlignBottom)
            series.attachAxis(x_axis)

            if max_val == 0:
                max_val = 1.0
            y_axis = QValueAxis()
            y_axis.setRange(0, max_val * 1.15)
            y_axis.setLabelFormat("%.0f")
            y_axis.setTitleText("EUR" if is_income else "Count")
            self._chart.addAxis(y_axis, Qt.AlignLeft)
            series.attachAxis(y_axis)

            # Hover tooltip for line chart
            unit = " EUR" if is_income else ""
            def _line_hovered(point, state, _u=unit):
                if state:
                    ms = point.x()
                    dt_obj = QDateTime.fromMSecsSinceEpoch(int(ms))
                    label = dt_obj.toString("MMM yyyy")
                    QToolTip.showText(
                        QCursor.pos(),
                        f"{label}: {point.y():,.0f}{_u}"
                    )
                else:
                    QToolTip.hideText()
            series.hovered.connect(_line_hovered)

        self._chart.setTitle(f"{metric}  |  {resolution}  ({from_yr} - {to_yr})")
        self._chart.setTitleFont(QFont("", 11, QFont.Bold))

    # ── Export ─────────────────────────────────────────────────────────────────

    def _export_excel(self):
        from services.report_service import export_dashboard_excel
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Dashboard Export", str(OUTPUT_DIR / "dashboard_export.xlsx"),
            "Excel Files (*.xlsx)"
        )
        if not path:
            return
        try:
            from pathlib import Path
            out = export_dashboard_excel(
                self._all_projects,
                self._all_invoices,
                self._all_debt_summaries,
                self._yearly_summary,
                output_dir=Path(path).parent,
            )
            QMessageBox.information(self, "Export Complete", f"Saved to:\n{out}")
        except Exception as e:
            QMessageBox.critical(self, "Export Error", str(e))
