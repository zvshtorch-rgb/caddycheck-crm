"""Settings page – data paths, email SMTP configuration, and app preferences."""
import logging
from pathlib import Path

from PySide6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QFrame,
    QPushButton, QFormLayout, QLineEdit, QSpinBox,
    QCheckBox, QGroupBox, QMessageBox, QTextEdit,
    QScrollArea, QSizePolicy, QFileDialog,
)
from PySide6.QtCore import Qt

from config.settings import (
    get_email_config, save_email_config,
    get_data_paths, save_data_paths,
)

logger = logging.getLogger(__name__)


class SettingsPage(QWidget):
    """Application settings including SMTP email configuration."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self._config = get_email_config()
        self._paths_changed = False
        self._build_ui()
        self._load_values()

    def _build_ui(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("QScrollArea { border: none; }")

        content = QWidget()
        scroll.setWidget(content)

        outer_layout = QVBoxLayout(self)
        outer_layout.setContentsMargins(0, 0, 0, 0)
        outer_layout.addWidget(scroll)

        layout = QVBoxLayout(content)
        layout.setContentsMargins(24, 20, 24, 20)
        layout.setSpacing(20)

        # Title
        title = QLabel("Settings")
        title.setStyleSheet("font-size: 24px; font-weight: bold; color: #2C3E50;")
        layout.addWidget(title)

        # ── Data Paths (OneDrive / SharePoint) ─────────────────────────────────
        paths_group = self._make_group("Data File Paths  (OneDrive / SharePoint / Local)")
        paths_info = QLabel(
            "Point these paths to your shared OneDrive or SharePoint folder so the\n"
            "whole team reads and writes the same files. Each person sets this once\n"
            "on their own machine — OneDrive syncs the rest automatically."
        )
        paths_info.setStyleSheet(
            "background: #EBF5FB; color: #1A5276; border: 1px solid #AED6F1; "
            "border-radius: 4px; padding: 8px 12px; font-size: 12px;"
        )
        paths_group.layout().addWidget(paths_info)

        paths_form = QFormLayout()
        paths_form.setSpacing(10)
        paths_form.setLabelAlignment(Qt.AlignRight)

        # Projects file
        proj_row = QHBoxLayout()
        self._projects_path = QLineEdit()
        self._projects_path.setPlaceholderText("e.g. C:\\Users\\you\\OneDrive\\CaddyCheck\\CaddyCheckProjectsInfo.xlsx")
        proj_row.addWidget(self._projects_path)
        proj_browse = QPushButton("Browse…")
        proj_browse.setFixedWidth(80)
        proj_browse.clicked.connect(lambda: self._browse_file(
            self._projects_path, "Select Projects File", "Excel Files (*.xlsx)"
        ))
        proj_row.addWidget(proj_browse)
        paths_form.addRow("Projects File:", proj_row)

        # Invoice template
        tmpl_row = QHBoxLayout()
        self._template_path = QLineEdit()
        self._template_path.setPlaceholderText("e.g. C:\\Users\\you\\OneDrive\\CaddyCheck\\CC_M_inv_template.xlsx")
        tmpl_row.addWidget(self._template_path)
        tmpl_browse = QPushButton("Browse…")
        tmpl_browse.setFixedWidth(80)
        tmpl_browse.clicked.connect(lambda: self._browse_file(
            self._template_path, "Select Invoice Template", "Excel Files (*.xlsx)"
        ))
        tmpl_row.addWidget(tmpl_browse)
        paths_form.addRow("Invoice Template:", tmpl_row)

        # Output directory
        out_row = QHBoxLayout()
        self._output_dir = QLineEdit()
        self._output_dir.setPlaceholderText("e.g. C:\\Users\\you\\OneDrive\\CaddyCheck\\output")
        out_row.addWidget(self._output_dir)
        out_browse = QPushButton("Browse…")
        out_browse.setFixedWidth(80)
        out_browse.clicked.connect(lambda: self._browse_dir(self._output_dir))
        out_row.addWidget(out_browse)
        paths_form.addRow("Output Folder:", out_row)

        paths_group.layout().addLayout(paths_form)

        # Test + save paths row
        paths_btn_row = QHBoxLayout()
        test_paths_btn = QPushButton("Test Paths")
        test_paths_btn.setFixedHeight(32)
        test_paths_btn.setStyleSheet(
            "QPushButton { background: #2980B9; color: white; border-radius: 4px; padding: 0 12px; }"
            "QPushButton:hover { background: #1A5276; }"
        )
        test_paths_btn.clicked.connect(self._test_paths)
        paths_btn_row.addWidget(test_paths_btn)

        save_paths_btn = QPushButton("Save Paths & Reload Data")
        save_paths_btn.setFixedHeight(32)
        save_paths_btn.setStyleSheet(
            "QPushButton { background: #27AE60; color: white; border-radius: 4px; "
            "padding: 0 12px; font-weight: 600; }"
            "QPushButton:hover { background: #1E8449; }"
        )
        save_paths_btn.clicked.connect(self._save_paths)
        paths_btn_row.addWidget(save_paths_btn)
        paths_btn_row.addStretch()
        paths_group.layout().addLayout(paths_btn_row)

        layout.addWidget(paths_group)

        # ── SMTP Settings ──────────────────────────────────────────────────────
        smtp_group = self._make_group("Email / SMTP Configuration")
        smtp_form = QFormLayout()
        smtp_form.setSpacing(10)
        smtp_form.setLabelAlignment(Qt.AlignRight)

        self._smtp_host = QLineEdit()
        self._smtp_host.setPlaceholderText("smtp.gmail.com")
        smtp_form.addRow("SMTP Host:", self._smtp_host)

        self._smtp_port = QSpinBox()
        self._smtp_port.setRange(1, 65535)
        self._smtp_port.setValue(587)
        self._smtp_port.setFixedWidth(100)
        smtp_form.addRow("SMTP Port:", self._smtp_port)

        self._use_tls = QCheckBox("Use STARTTLS (recommended)")
        smtp_form.addRow("TLS:", self._use_tls)

        self._smtp_user = QLineEdit()
        self._smtp_user.setPlaceholderText("your@email.com")
        smtp_form.addRow("Username:", self._smtp_user)

        self._smtp_pass = QLineEdit()
        self._smtp_pass.setEchoMode(QLineEdit.Password)
        self._smtp_pass.setPlaceholderText("App password or SMTP password")
        smtp_form.addRow("Password:", self._smtp_pass)

        self._sender_name = QLineEdit()
        self._sender_name.setPlaceholderText("CaddyCheck CRM")
        smtp_form.addRow("Sender Name:", self._sender_name)

        self._sender_email = QLineEdit()
        self._sender_email.setPlaceholderText("sender@email.com")
        smtp_form.addRow("Sender Email:", self._sender_email)

        smtp_group.layout().addLayout(smtp_form)

        # Test connection button
        test_row = QHBoxLayout()
        test_btn = QPushButton("Test Connection")
        test_btn.setFixedHeight(34)
        test_btn.setStyleSheet(
            "QPushButton { background: #2980B9; color: white; border-radius: 4px; "
            "padding: 0 14px; }"
            "QPushButton:hover { background: #1A5276; }"
        )
        test_btn.clicked.connect(self._test_connection)
        test_row.addWidget(test_btn)
        test_row.addStretch()
        smtp_group.layout().addLayout(test_row)

        layout.addWidget(smtp_group)

        # ── Default Recipients ─────────────────────────────────────────────────
        rec_group = self._make_group("Default Email Recipients")
        rec_form = QFormLayout()
        rec_form.setSpacing(10)
        rec_form.setLabelAlignment(Qt.AlignRight)

        self._recipients = QLineEdit()
        self._recipients.setPlaceholderText("recipient1@example.com, recipient2@example.com")
        rec_form.addRow("To (comma-separated):", self._recipients)

        self._cc = QLineEdit()
        self._cc.setPlaceholderText("cc@example.com")
        rec_form.addRow("CC (comma-separated):", self._cc)

        self._subject_template = QLineEdit()
        self._subject_template.setPlaceholderText("Monthly Invoice - {month} {year}")
        rec_form.addRow("Subject Template:", self._subject_template)

        rec_group.layout().addLayout(rec_form)
        layout.addWidget(rec_group)

        # ── Default Body Template ──────────────────────────────────────────────
        body_group = self._make_group("Default Email Body Template")
        body_layout = QVBoxLayout()
        body_layout.addWidget(QLabel(
            "Use {month} and {year} as placeholders:"
        ))
        self._body_template = QTextEdit()
        self._body_template.setFixedHeight(120)
        body_layout.addWidget(self._body_template)
        body_group.layout().addLayout(body_layout)
        layout.addWidget(body_group)

        # ── Save button ────────────────────────────────────────────────────────
        save_row = QHBoxLayout()
        save_btn = QPushButton("Save Settings")
        save_btn.setFixedHeight(40)
        save_btn.setStyleSheet(
            "QPushButton { background: #27AE60; color: white; border-radius: 5px; "
            "padding: 0 24px; font-size: 15px; font-weight: 600; }"
            "QPushButton:hover { background: #1E8449; }"
        )
        save_btn.clicked.connect(self._save)
        save_row.addWidget(save_btn)
        save_row.addStretch()
        layout.addLayout(save_row)

        layout.addStretch()

    @staticmethod
    def _make_group(title: str) -> QGroupBox:
        group = QGroupBox(title)
        group.setStyleSheet(
            "QGroupBox { font-weight: bold; font-size: 14px; color: #2C3E50; "
            "border: 1px solid #DEE2E6; border-radius: 8px; margin-top: 8px; "
            "padding: 12px; }"
            "QGroupBox::title { subcontrol-origin: margin; left: 12px; padding: 0 4px; }"
        )
        group.setLayout(QVBoxLayout())
        group.layout().setSpacing(8)
        return group

    # ── Browse helpers ─────────────────────────────────────────────────────────

    def _browse_file(self, line_edit: QLineEdit, title: str, filter_str: str):
        current = line_edit.text().strip()
        start_dir = str(Path(current).parent) if current else ""
        path, _ = QFileDialog.getOpenFileName(self, title, start_dir, filter_str)
        if path:
            line_edit.setText(path)

    def _browse_dir(self, line_edit: QLineEdit):
        current = line_edit.text().strip()
        path = QFileDialog.getExistingDirectory(self, "Select Output Folder", current)
        if path:
            line_edit.setText(path)

    def _test_paths(self):
        proj = Path(self._projects_path.text().strip())
        tmpl = Path(self._template_path.text().strip())
        out  = Path(self._output_dir.text().strip())

        lines = []
        lines.append(f"Projects file:    {'EXISTS' if proj.exists() else 'NOT FOUND'}  →  {proj}")
        lines.append(f"Invoice template: {'EXISTS' if tmpl.exists() else 'NOT FOUND'}  →  {tmpl}")
        lines.append(f"Output folder:    {'EXISTS' if out.exists() else 'WILL BE CREATED'}  →  {out}")

        all_ok = proj.exists() and tmpl.exists()
        if all_ok:
            QMessageBox.information(self, "Path Test", "\n".join(lines))
        else:
            QMessageBox.warning(self, "Path Test – Issues Found", "\n".join(lines))

    def _save_paths(self):
        proj = self._projects_path.text().strip()
        tmpl = self._template_path.text().strip()
        out  = self._output_dir.text().strip()

        if not proj or not tmpl or not out:
            QMessageBox.warning(self, "Missing Paths", "Please fill in all three paths.")
            return

        if not Path(proj).exists():
            reply = QMessageBox.question(
                self, "File Not Found",
                f"Projects file not found:\n{proj}\n\nSave anyway?",
                QMessageBox.Yes | QMessageBox.No,
            )
            if reply != QMessageBox.Yes:
                return

        try:
            save_data_paths({
                "projects_file":    proj,
                "invoice_template": tmpl,
                "output_dir":       out,
            })
            self._paths_changed = True
            # Ask main window to reload
            main_win = self.window()
            if hasattr(main_win, "refresh_data"):
                reply = QMessageBox.question(
                    self, "Paths Saved",
                    "Paths saved. Reload data from the new location now?",
                    QMessageBox.Yes | QMessageBox.No,
                )
                if reply == QMessageBox.Yes:
                    main_win.refresh_data()
            else:
                QMessageBox.information(self, "Saved", "Paths saved. Restart the app to reload data.")
        except Exception as e:
            QMessageBox.critical(self, "Save Error", f"Failed to save paths:\n{e}")

    # ── Load / Save ────────────────────────────────────────────────────────────

    def _load_values(self):
        # Data paths
        paths = get_data_paths()
        self._projects_path.setText(str(paths["projects_file"]))
        self._template_path.setText(str(paths["invoice_template"]))
        self._output_dir.setText(str(paths["output_dir"]))

        c = self._config
        self._smtp_host.setText(c.get("smtp_host", ""))
        self._smtp_port.setValue(int(c.get("smtp_port", 587)))
        self._use_tls.setChecked(bool(c.get("smtp_use_tls", True)))
        self._smtp_user.setText(c.get("smtp_username", ""))
        self._smtp_pass.setText(c.get("smtp_password", ""))
        self._sender_name.setText(c.get("sender_name", "CaddyCheck CRM"))
        self._sender_email.setText(c.get("sender_email", ""))
        self._recipients.setText(", ".join(c.get("default_recipients", [])))
        self._cc.setText(", ".join(c.get("default_cc", [])))
        self._subject_template.setText(
            c.get("default_subject_template", "Monthly Invoice - {month} {year}")
        )
        self._body_template.setPlainText(
            c.get("default_body_template", "")
        )

    def _save(self):
        config = {
            "smtp_host": self._smtp_host.text().strip(),
            "smtp_port": self._smtp_port.value(),
            "smtp_use_tls": self._use_tls.isChecked(),
            "smtp_username": self._smtp_user.text().strip(),
            "smtp_password": self._smtp_pass.text(),
            "sender_name": self._sender_name.text().strip(),
            "sender_email": self._sender_email.text().strip(),
            "default_recipients": [
                r.strip() for r in self._recipients.text().split(",") if r.strip()
            ],
            "default_cc": [
                c.strip() for c in self._cc.text().split(",") if c.strip()
            ],
            "default_subject_template": self._subject_template.text().strip(),
            "default_body_template": self._body_template.toPlainText(),
        }
        try:
            save_email_config(config)
            self._config = config
            QMessageBox.information(self, "Saved", "Settings saved successfully.")
        except Exception as e:
            QMessageBox.critical(self, "Save Error", f"Failed to save settings:\n{e}")

    def _test_connection(self):
        config = {
            "smtp_host": self._smtp_host.text().strip(),
            "smtp_port": self._smtp_port.value(),
            "smtp_use_tls": self._use_tls.isChecked(),
            "smtp_username": self._smtp_user.text().strip(),
            "smtp_password": self._smtp_pass.text(),
        }
        try:
            from services.email_service import test_smtp_connection
            success, msg = test_smtp_connection(config)
            if success:
                QMessageBox.information(self, "Connection Test", msg)
            else:
                QMessageBox.warning(self, "Connection Failed", msg)
        except Exception as e:
            QMessageBox.critical(self, "Test Error", str(e))
