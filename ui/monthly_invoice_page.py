"""Monthly invoice generation page."""
import datetime
import logging
import os
from pathlib import Path
from typing import List, Optional

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QFrame,
    QTableWidget, QTableWidgetItem, QHeaderView,
    QPushButton, QComboBox, QSpinBox, QGroupBox,
    QFormLayout, QSplitter, QMessageBox, QFileDialog,
    QTextEdit, QSizePolicy, QLineEdit, QDialog,
    QDialogButtonBox, QProgressDialog,
)
from PySide6.QtCore import Qt, QThread, Signal, QObject
from PySide6.QtGui import QColor, QFont

from config.settings import MONTH_ORDER, normalize_month, OUTPUT_DIR
from models.project import Project
from services.excel_service import get_projects_for_month
from services.invoice_service import generate_monthly_invoice, get_invoice_preview_data

logger = logging.getLogger(__name__)

CURRENT_YEAR = datetime.datetime.now().year


class EmailDialog(QDialog):
    """Dialog to compose and send the invoice email."""

    def __init__(self, attachment_path: Path, month: str, year: int, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Send Invoice Email")
        self.setMinimumSize(580, 480)
        self._attachment_path = attachment_path

        from config.settings import get_email_config
        self._config = get_email_config()

        layout = QVBoxLayout(self)

        form = QFormLayout()
        form.setSpacing(10)

        self._to_edit = QLineEdit(", ".join(self._config.get("default_recipients", [])))
        self._cc_edit = QLineEdit(", ".join(self._config.get("default_cc", [])))
        subject_template = self._config.get(
            "default_subject_template", "Monthly Invoice - {month} {year}"
        )
        self._subject_edit = QLineEdit(
            subject_template.format(month=month, year=year)
        )
        form.addRow("To:", self._to_edit)
        form.addRow("CC:", self._cc_edit)
        form.addRow("Subject:", self._subject_edit)
        layout.addLayout(form)

        body_label = QLabel("Body:")
        layout.addWidget(body_label)
        body_template = self._config.get(
            "default_body_template",
            "Dear Team,\n\nPlease find attached the monthly invoice.\n\nBest regards"
        )
        self._body_edit = QTextEdit(
            body_template.format(month=month, year=year)
        )
        self._body_edit.setMinimumHeight(150)
        layout.addWidget(self._body_edit)

        attach_label = QLabel(f"Attachment: {attachment_path.name}")
        attach_label.setStyleSheet("color: #555; font-size: 12px;")
        layout.addWidget(attach_label)

        buttons = QDialogButtonBox(
            QDialogButtonBox.Ok | QDialogButtonBox.Cancel
        )
        buttons.button(QDialogButtonBox.Ok).setText("Send Email")
        buttons.accepted.connect(self._send)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _send(self):
        recipients = [r.strip() for r in self._to_edit.text().split(",") if r.strip()]
        cc = [c.strip() for c in self._cc_edit.text().split(",") if c.strip()]
        subject = self._subject_edit.text()
        body = self._body_edit.toPlainText()

        if not recipients:
            QMessageBox.warning(self, "Missing Recipients", "Please enter at least one recipient.")
            return

        try:
            from services.email_service import send_invoice_email
            send_invoice_email(
                attachment_path=self._attachment_path,
                subject=subject,
                body=body,
                recipients=recipients,
                cc=cc,
                config=self._config,
            )
            QMessageBox.information(self, "Email Sent", "Invoice email sent successfully!")
            self.accept()
        except Exception as e:
            QMessageBox.critical(self, "Email Error", f"Failed to send email:\n{e}")


class MonthlyInvoicePage(QWidget):
    """Page for selecting a month/year and generating a monthly invoice."""

    def __init__(self, projects: List[Project], parent=None):
        super().__init__(parent)
        self._projects = projects
        self._generated_path: Optional[Path] = None
        self._filtered_projects: List[Project] = []

        self._build_ui()
        self._refresh_preview()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(14)

        # Title
        title = QLabel("Monthly Invoice Generation")
        title.setStyleSheet("font-size: 24px; font-weight: bold; color: #2C3E50;")
        layout.addWidget(title)

        # ── Settings row ───────────────────────────────────────────────────────
        settings_frame = QFrame()
        settings_frame.setStyleSheet(
            "QFrame { background: white; border: 1px solid #DEE2E6; border-radius: 8px; }"
        )
        settings_layout = QHBoxLayout(settings_frame)
        settings_layout.setContentsMargins(20, 14, 20, 14)
        settings_layout.setSpacing(16)

        # Month selector
        settings_layout.addWidget(QLabel("Month:"))
        self._month_cb = QComboBox()
        self._month_cb.setFixedHeight(34)
        for m in MONTH_ORDER:
            self._month_cb.addItem(m)
        # Default to current month
        current_month_name = MONTH_ORDER[datetime.datetime.now().month - 1]
        self._month_cb.setCurrentText(current_month_name)
        self._month_cb.currentTextChanged.connect(self._refresh_preview)
        settings_layout.addWidget(self._month_cb)

        # Year selector
        settings_layout.addWidget(QLabel("Year:"))
        self._year_spin = QSpinBox()
        self._year_spin.setRange(2013, CURRENT_YEAR + 5)
        self._year_spin.setValue(CURRENT_YEAR)
        self._year_spin.setFixedHeight(34)
        self._year_spin.setFixedWidth(90)
        self._year_spin.valueChanged.connect(self._refresh_preview)
        settings_layout.addWidget(self._year_spin)

        # Invoice number
        settings_layout.addWidget(QLabel("Invoice #:"))
        self._inv_no_spin = QSpinBox()
        self._inv_no_spin.setRange(1000, 99999)
        self._inv_no_spin.setValue(8670)
        self._inv_no_spin.setFixedHeight(34)
        self._inv_no_spin.setFixedWidth(90)
        settings_layout.addWidget(self._inv_no_spin)

        settings_layout.addStretch()

        # Generate button
        self._gen_btn = QPushButton("Generate Invoice")
        self._gen_btn.setFixedHeight(38)
        self._gen_btn.setStyleSheet(
            "QPushButton { background: #27AE60; color: white; border-radius: 5px; "
            "padding: 0 18px; font-size: 14px; font-weight: 600; }"
            "QPushButton:hover { background: #1E8449; }"
        )
        self._gen_btn.clicked.connect(self._generate)
        settings_layout.addWidget(self._gen_btn)

        # Open button (only enabled after generation)
        self._open_btn = QPushButton("Open File")
        self._open_btn.setFixedHeight(38)
        self._open_btn.setEnabled(False)
        self._open_btn.setStyleSheet(
            "QPushButton { background: #2D6A9F; color: white; border-radius: 5px; "
            "padding: 0 14px; font-size: 14px; }"
            "QPushButton:hover { background: #1A5276; }"
            "QPushButton:disabled { background: #AAA; }"
        )
        self._open_btn.clicked.connect(self._open_file)
        settings_layout.addWidget(self._open_btn)

        # Email button
        self._email_btn = QPushButton("Send by Email")
        self._email_btn.setFixedHeight(38)
        self._email_btn.setEnabled(False)
        self._email_btn.setStyleSheet(
            "QPushButton { background: #8E44AD; color: white; border-radius: 5px; "
            "padding: 0 14px; font-size: 14px; }"
            "QPushButton:hover { background: #6C3483; }"
            "QPushButton:disabled { background: #AAA; }"
        )
        self._email_btn.clicked.connect(self._send_email)
        settings_layout.addWidget(self._email_btn)

        layout.addWidget(settings_frame)

        # ── Preview table ──────────────────────────────────────────────────────
        preview_group = QGroupBox("Invoice Preview")
        preview_group.setStyleSheet(
            "QGroupBox { font-weight: bold; font-size: 14px; color: #2C3E50; "
            "border: 1px solid #DEE2E6; border-radius: 8px; margin-top: 8px; }"
            "QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 4px; }"
        )
        preview_layout = QVBoxLayout(preview_group)

        # Title preview label
        self._title_label = QLabel()
        self._title_label.setStyleSheet(
            "font-size: 15px; font-weight: bold; color: #2D6A9F; padding: 4px 0;"
        )
        preview_layout.addWidget(self._title_label)

        self._preview_table = QTableWidget()
        self._preview_table.setColumnCount(5)
        self._preview_table.setHorizontalHeaderLabels([
            "Project / Store", "# Units", "Maint. Year", "Rate (€)", "Line Total (€)"
        ])
        self._preview_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self._preview_table.setAlternatingRowColors(True)
        self._preview_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self._preview_table.setStyleSheet(
            "QTableWidget { border: none; gridline-color: #DEE2E6; }"
        )
        preview_layout.addWidget(self._preview_table)

        # Totals bar
        self._totals_label = QLabel()
        self._totals_label.setStyleSheet(
            "font-size: 14px; font-weight: bold; color: #2C3E50; "
            "padding: 6px 0; text-align: right;"
        )
        self._totals_label.setAlignment(Qt.AlignRight)
        preview_layout.addWidget(self._totals_label)

        layout.addWidget(preview_group)

        # Status bar
        self._status_label = QLabel("Select a month and year, then click Generate Invoice.")
        self._status_label.setStyleSheet("color: #555; font-size: 12px;")
        layout.addWidget(self._status_label)

    # ── Refresh preview ────────────────────────────────────────────────────────

    def _refresh_preview(self, _=None):
        month = self._month_cb.currentText()
        year = self._year_spin.value()

        self._filtered_projects = [
            p for p in get_projects_for_month(self._projects, month)
            if p.is_active()
        ]
        preview_rows = get_invoice_preview_data(self._filtered_projects, month, year)

        self._title_label.setText(
            f"Invoice Title: Iretailcheck - Maintenance - {month[:3]} {year}"
        )

        self._preview_table.setRowCount(len(preview_rows))
        for row_idx, row_data in enumerate(preview_rows):
            is_total_row = row_data.get("rate") == "TOTAL"

            def _item(val, align=Qt.AlignLeft, bold=False):
                text = str(val) if val not in (None, "") else ""
                item = QTableWidgetItem(text)
                item.setTextAlignment(align | Qt.AlignVCenter)
                if bold or is_total_row:
                    item.setFont(QFont("", -1, QFont.Bold))
                if is_total_row:
                    item.setBackground(QColor("#D5F5E3"))
                return item

            self._preview_table.setItem(row_idx, 0, _item(row_data["project_name"]))
            self._preview_table.setItem(row_idx, 1, _item(
                row_data["num_cams"], Qt.AlignCenter
            ))
            self._preview_table.setItem(row_idx, 2, _item(
                row_data["maintenance_year"], Qt.AlignCenter
            ))
            self._preview_table.setItem(row_idx, 3, _item(
                f"€{row_data['rate']}" if row_data["rate"] != "TOTAL" else "TOTAL",
                Qt.AlignRight
            ))
            line_total = row_data["line_total"]
            self._preview_table.setItem(row_idx, 4, _item(
                f"€{line_total:,.0f}" if isinstance(line_total, (int, float)) else "",
                Qt.AlignRight,
                bold=is_total_row,
            ))

        # Find total from last row
        total = 0.0
        for rd in preview_rows:
            if rd.get("rate") == "TOTAL":
                total = rd.get("line_total", 0.0)

        self._totals_label.setText(
            f"Projects in {month}: {len(self._filtered_projects)}   |   "
            f"Grand Total: €{total:,.0f}"
        )

        self._preview_table.resizeRowsToContents()

    # ── Generate ───────────────────────────────────────────────────────────────

    def _generate(self):
        month = self._month_cb.currentText()
        year = self._year_spin.value()
        inv_no = self._inv_no_spin.value()

        if not self._filtered_projects:
            QMessageBox.warning(
                self, "No Projects",
                f"No projects found with payment month '{month}'.\n"
                "Please check the Projects overview data."
            )
            return

        try:
            output_path = generate_monthly_invoice(
                projects=self._filtered_projects,
                month_name=month,
                year=year,
                invoice_number=inv_no,
            )
            self._generated_path = output_path
            self._open_btn.setEnabled(True)
            self._email_btn.setEnabled(True)
            self._status_label.setText(f"Invoice generated: {output_path}")
            QMessageBox.information(
                self, "Invoice Generated",
                f"Invoice saved to:\n{output_path}"
            )
        except Exception as e:
            logger.exception("Invoice generation failed")
            QMessageBox.critical(self, "Generation Error", f"Failed to generate invoice:\n{e}")

    def _open_file(self):
        if self._generated_path and self._generated_path.exists():
            os.startfile(str(self._generated_path))
        else:
            QMessageBox.warning(self, "File Not Found", "Generated file not found.")

    def _send_email(self):
        if not self._generated_path:
            QMessageBox.warning(self, "No Invoice", "Please generate an invoice first.")
            return
        month = self._month_cb.currentText()
        year = self._year_spin.value()
        dlg = EmailDialog(self._generated_path, month, year, parent=self)
        dlg.exec()
