from __future__ import annotations

import re
import shutil
import sys
import threading
from pathlib import Path

from PySide6 import QtCore, QtGui, QtWidgets

# Allow running as a module or as a script
try:
    from . import backend as be  # type: ignore
except Exception:  # pragma: no cover
    import sys as _sys
    from pathlib import Path as _Path
    _root = _Path(__file__).resolve().parents[1]
    if str(_root) not in _sys.path:
        _sys.path.insert(0, str(_root))
    import ui_next.backend as be  # type: ignore


class ProcWorker(QtCore.QThread):
    line = QtCore.Signal(str)
    finished = QtCore.Signal(int)

    def __init__(self, popen_factory, parent=None):
        super().__init__(parent)
        self._popen_factory = popen_factory
        self._stop = threading.Event()
        self._proc = None

    def run(self):
        rc = 0
        try:
            self._proc = self._popen_factory()
            stream = self._proc.stdout
            if stream is not None:
                for line in stream:
                    if self._stop.is_set():
                        break
                    self.line.emit(line.rstrip("\n"))
            rc = self._proc.wait()
        except Exception as e:
            self.line.emit(f"[ERROR] {e}")
            rc = 1
        finally:
            self.finished.emit(rc)

    def stop(self):
        self._stop.set()
        try:
            if self._proc and self._proc.poll() is None:
                self._proc.terminate()
        except Exception:
            pass


class RunProgressDialog(QtWidgets.QDialog):
    """Large popup that visualizes term progress with a simple spinner animation."""

    canceled = QtCore.Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Running EIDP Scanner")
        self.setModal(True)
        self.setWindowModality(QtCore.Qt.WindowModality.ApplicationModal)
        self.setWindowFlag(QtCore.Qt.WindowType.WindowContextHelpButtonHint, False)
        self.setWindowFlag(QtCore.Qt.WindowType.WindowCloseButtonHint, False)
        self.resize(480, 260)
        self.setStyleSheet(
            """
            QDialog {
                background-color: #0b1526;
                color: #f6fbff;
            }
            QDialog QLabel {
                color: #f6fbff;
            }
            QProgressBar {
                background-color: #14253d;
                color: #0b1526;
                border: 1px solid #284364;
                border-radius: 4px;
            }
            QProgressBar::chunk {
                background-color: #3db6ff;
            }
            QPushButton[variant="ghost"] {
                background: transparent;
                color: #f6fbff;
                border: 1px solid #3db6ff;
            }
            QPushButton[variant="ghost"]:disabled {
                color: #94a6c3;
                border-color: #2a3b52;
            }
            """
        )

        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(28, 28, 28, 28)
        layout.setSpacing(16)

        self.lbl_heading = QtWidgets.QLabel("Executing run...")
        font = self.lbl_heading.font()
        font.setPointSize(18)
        font.setBold(True)
        self.lbl_heading.setFont(font)
        self.lbl_heading.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)

        self.lbl_status = QtWidgets.QLabel("Preparing scanner")
        self.lbl_status.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.lbl_status.setWordWrap(True)
        self.lbl_status.setStyleSheet("font-size: 13px;")

        self.spinner_label = QtWidgets.QLabel("•")
        spin_font = self.spinner_label.font()
        spin_font.setPointSize(24)
        spin_font.setBold(False)
        self.spinner_label.setFont(spin_font)
        self.spinner_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)

        self.progress_bar = QtWidgets.QProgressBar()
        self.progress_bar.setRange(0, 0)  # indeterminate until totals stream in
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("Working...")

        self.detail_label = QtWidgets.QLabel("Searching: 0 / 0 terms \u2022 Found: 0")
        self.detail_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.detail_label.setStyleSheet("font-size: 12px; color: #e0e7f0;")

        self.hint_label = QtWidgets.QLabel("This window closes automatically when the run finishes.")
        self.hint_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.hint_label.setStyleSheet("color: #a5b8d6; font-size: 11px;")

        self.btn_cancel = QtWidgets.QPushButton("Abort Run")
        self.btn_cancel.setProperty("variant", "ghost")
        self.btn_cancel.clicked.connect(self._on_cancel_clicked)

        layout.addWidget(self.lbl_heading)
        layout.addWidget(self.lbl_status)
        layout.addWidget(self.spinner_label)
        layout.addWidget(self.progress_bar)
        layout.addWidget(self.detail_label)
        layout.addWidget(self.hint_label)
        layout.addWidget(self.btn_cancel)

        self._spinner_frames = ["•", "·", "•", "·"]
        self._spinner_index = 0
        self._anim_timer = QtCore.QTimer(self)
        self._anim_timer.setInterval(170)
        self._anim_timer.timeout.connect(self._advance_spinner)

    def _advance_spinner(self):
        self._spinner_index = (self._spinner_index + 1) % len(self._spinner_frames)
        self.spinner_label.setText(self._spinner_frames[self._spinner_index])

    def begin(self, status_text: str):
        self.lbl_status.setText(status_text)
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setFormat("Working...")
        self.detail_label.setText("Searching: 0 / 0 terms \u2022 Found: 0")
        self._spinner_index = 0
        self.spinner_label.setText(self._spinner_frames[0])
        self._anim_timer.start()
        self.btn_cancel.setEnabled(True)
        self.btn_cancel.setText("Abort Run")
        self.show()
        try:
            self.raise_()
            self.activateWindow()
        except Exception:
            pass

    def update_progress(self, completed: int, total: int, found: int = 0):
        if total <= 0:
            if self.progress_bar.maximum() != 0:
                self.progress_bar.setRange(0, 0)
                self.progress_bar.setFormat("Working...")
            self.detail_label.setText(f"Searching: {completed} terms \u2022 Found: {found}")
            return
        if self.progress_bar.maximum() == 0:
            self.progress_bar.setRange(0, 100)
        pct = max(0, min(100, int(round((completed * 100) / max(1, total)))))
        remaining = max(0, total - completed)
        self.progress_bar.setValue(pct)
        self.progress_bar.setFormat(f"{pct}%")
        self.detail_label.setText(f"Searching: {completed} / {total} terms \u2022 Found: {found}")

    def finish(self, message: str, success: bool = True):
        self._anim_timer.stop()
        self.spinner_label.setText("✓" if success else "×")
        self.lbl_status.setText(message)
        if self.progress_bar.maximum() == 0:
            self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(100 if success else self.progress_bar.value())
        self.progress_bar.setFormat("Done")
        self.btn_cancel.setEnabled(False)
        QtCore.QTimer.singleShot(1200, self.hide)

    def abort(self):
        self._anim_timer.stop()
        self.btn_cancel.setEnabled(False)
        self.hide()

    def closeEvent(self, event: QtGui.QCloseEvent):  # type: ignore[override]
        self._anim_timer.stop()
        super().closeEvent(event)

    def _on_cancel_clicked(self):
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.setText("Aborting...")
        self.lbl_status.setText("Stopping run...")
        self.canceled.emit()


class _DropZone(QtWidgets.QFrame):
    def __init__(self, hint: str, on_drop, parent=None):
        super().__init__(parent)
        self._on_drop = on_drop
        self._label = QtWidgets.QLabel(hint)
        self._label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self._label.setWordWrap(True)
        lay = QtWidgets.QVBoxLayout(self)
        lay.addStretch(1)
        lay.addWidget(self._label)
        lay.addStretch(1)
        self.setAcceptDrops(True)
        self.setFrameStyle(
            QtWidgets.QFrame.Shape.StyledPanel | QtWidgets.QFrame.Shadow.Plain
        )
        self.setStyleSheet("""
            QFrame { border: 2px dashed #c9d3e3; border-radius: 8px; min-height: 140px; background: #fafbfd; }
            QLabel { color: #5b6b7a; font-size: 14px; }
        """)

    def set_hint(self, text: str):
        self._label.setText(text)

    def dragEnterEvent(self, e: QtGui.QDragEnterEvent):  # type: ignore[override]
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
        else:
            e.ignore()

    def dragMoveEvent(self, e: QtGui.QDragMoveEvent):  # type: ignore[override]
        e.acceptProposedAction()

    def dropEvent(self, e: QtGui.QDropEvent):  # type: ignore[override]
        try:
            urls = e.mimeData().urls()
            paths = [u.toLocalFile() for u in urls if u.isLocalFile()]
            if paths:
                self._on_drop(paths)
            e.acceptProposedAction()
        except Exception:
            e.ignore()


class ToastNotification(QtWidgets.QWidget):
    """Cookie banner / toast notification that appears at bottom of window."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(QtCore.Qt.WindowType.ToolTip | QtCore.Qt.WindowType.FramelessWindowHint)
        self.setAttribute(QtCore.Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setStyleSheet("""
            QWidget {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #1e293b, stop:1 #0f172a);
                border: 2px solid #3b82f6;
                border-radius: 12px;
                padding: 0px;
            }
            QLabel {
                color: #f8fafc;
                font-size: 14px;
                font-weight: 600;
                padding: 12px 24px;
            }
        """)

        layout = QtWidgets.QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        self.label = QtWidgets.QLabel()
        self.label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.label)

        self._timer = QtCore.QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self.hide)

    def show_message(self, message: str, duration: int = 5000):
        """Show toast notification with message for duration milliseconds."""
        self.label.setText(message)
        self.adjustSize()

        # Position at bottom center of parent
        if self.parent():
            parent_rect = self.parent().geometry()
            x = parent_rect.x() + (parent_rect.width() - self.width()) // 2
            y = parent_rect.y() + parent_rect.height() - self.height() - 40
            self.move(x, y)

        self.show()
        self.raise_()
        self._timer.start(duration)


class TermsEditorDialog(QtWidgets.QDialog):
    """In-app editor for Smart-Snap terms with grouped headers and curated fields."""

    VISIBLE_COLUMN_ORDER = [
        "Data Group",
        "Term Label",
        "Smart Snap Type",
        "Term",
        "Pages",
        "GroupAfter",
        "GroupBefore",
        "Units",
        "Range (min)",
        "Range (max)",
        "Format",
        "Secondary Term",
        "Smart Position",
    ]
    COLUMN_DISPLAY_NAMES = {
        "Data Group": "Data Grouping",
        "Term Label": "Term Label",
        "Smart Snap Type": "Smart Snap Type",
        "Term": "Search Term",
        "Pages": "Pages",
        "GroupAfter": "Group After",
        "GroupBefore": "Group Before",
        "Units": "Units",
        "Range (min)": "Range (min)",
        "Range (max)": "Range (max)",
        "Format": "Format / Pattern",
        "Secondary Term": "Secondary Term",
        "Smart Position": "Smart Position",
    }
    DEFAULT_HIDDEN = {"Return", "Mode", "Line", "Column", "Anchor", "FieldIndex", "FieldSplit"}

    def __init__(self, terms_path: Path, parent=None):
        super().__init__(parent)
        self._terms_path = Path(terms_path)
        self._all_headers: list[str] = []
        self._visible_headers: list[str] = []
        self._hidden_headers: list[str] = []
        self._hidden_rows: list[dict[str, str]] = []
        self._dirty = False
        self._loading = False
        self._pending_save_notice = False

        self.setWindowTitle("Smart-Snap Terms Editor")
        self.resize(1280, 680)
        self.setObjectName("termsEditorDialog")
        self.setStyleSheet("""
            #termsEditorDialog {
                background-color: #f8fafc;
                color: #0f172a;
            }
            #termsEditorDialog QLabel {
                color: #0f172a;
            }
            #termsEditorDialog QLabel#termsPathLabel {
                color: #475569;
            }
            #termsEditorDialog QTableWidget {
                background-color: #ffffff;
                border: 1px solid #d4d4d8;
                gridline-color: #e4e4e7;
                selection-background-color: #d1d5db;
                selection-color: #0f172a;
            }
            #termsEditorDialog QTableWidget::item:selected {
                background-color: #d1d5db;
                color: #0f172a;
            }
            #termsEditorDialog QHeaderView::section {
                background-color: #eef2ff;
                color: #0f172a;
                padding: 8px;
                border: 1px solid #cbd5f5;
                font-weight: 600;
            }
            #termsEditorDialog QPushButton {
                background-color: #ffffff;
                color: #0f172a;
                border: 1px solid #cbd5f5;
                border-radius: 6px;
                padding: 8px 14px;
            }
            #termsEditorDialog QPushButton:hover {
                background-color: #f1f5f9;
            }
            #termsEditorDialog QPushButton[variant="primary"] {
                background-color: #2563eb;
                border-color: #2563eb;
                color: #ffffff;
            }
            #termsEditorDialog QPushButton[variant="primary"]:hover {
                background-color: #1d4ed8;
            }
        """)

        self._combo_defs = {
            "Smart Snap Type": {
                "options": self._smart_type_options(),
                "default": "",
            },
        }

        layout = QtWidgets.QVBoxLayout(self)
        title = QtWidgets.QLabel("Edit the Smart-Snap input spreadsheet directly in the app.")
        title.setStyleSheet("font-size: 18px; font-weight: 700; color: #0f172a;")
        layout.addWidget(title)
        hint = QtWidgets.QLabel(f"File: {self._terms_path}")
        hint.setObjectName("termsPathLabel")
        hint.setStyleSheet("color: #556070;")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        self.table = QtWidgets.QTableWidget()
        self.table.setAlternatingRowColors(False)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.SingleSelection)
        self.table.verticalHeader().setVisible(True)
        self.table.verticalHeader().setDefaultAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.table.verticalHeader().setDefaultSectionSize(42)
        self.table.setStyleSheet("""
            QTableWidget {
                background-color: #ffffff;
                border: 1px solid #d4d4d8;
                gridline-color: #e4e4e7;
                selection-background-color: #d1d5db;
                selection-color: #0f172a;
            }
            QTableWidget::item {
                padding: 10px;
                color: #0f172a;
            }
            QTableWidget::item:selected {
                background-color: #d1d5db;
                color: #0f172a;
            }
            QTableWidget QLineEdit {
                background-color: #ffffff;
                border: 1px solid #94a3b8;
                border-radius: 4px;
                padding: 6px 8px;
                color: #0f172a;
                selection-background-color: #bfdbfe;
                min-height: 26px;
            }
        """)
        header = self.table.horizontalHeader()
        header.setStretchLastSection(False)
        # Clicking the row number highlights the full row
        self.table.verticalHeader().sectionClicked.connect(self._on_vertical_header_clicked)
        layout.addWidget(self.table, 1)

        row_btns = QtWidgets.QHBoxLayout()
        self.btn_add_row = QtWidgets.QPushButton("Add Row")
        self.btn_duplicate_row = QtWidgets.QPushButton("Duplicate Row")
        self.btn_move_up = QtWidgets.QPushButton("Move Up")
        self.btn_move_down = QtWidgets.QPushButton("Move Down")
        self.btn_delete_row = QtWidgets.QPushButton("Delete Selected")
        self.btn_add_row.clicked.connect(self._add_blank_row)
        self.btn_duplicate_row.clicked.connect(self._duplicate_row)
        self.btn_move_up.clicked.connect(lambda: self._move_rows(-1))
        self.btn_move_down.clicked.connect(lambda: self._move_rows(1))
        self.btn_delete_row.clicked.connect(self._delete_rows)
        row_btns.addWidget(self.btn_add_row)
        row_btns.addWidget(self.btn_duplicate_row)
        row_btns.addWidget(self.btn_move_up)
        row_btns.addWidget(self.btn_move_down)
        row_btns.addWidget(self.btn_delete_row)
        row_btns.addStretch(1)
        layout.addLayout(row_btns)

        bottom = QtWidgets.QHBoxLayout()
        self._status_label = QtWidgets.QLabel("Loading...")
        self._status_label.setObjectName("termsStatusLabel")
        self._status_label.setStyleSheet("color: #0f172a; font-weight: 600;")
        bottom.addWidget(self._status_label)
        bottom.addStretch(1)
        self.btn_open_excel = QtWidgets.QPushButton("Open in Excel")
        self.btn_open_excel.clicked.connect(self._open_in_excel)
        self.btn_save_close = QtWidgets.QPushButton("Save & Close")
        self.btn_save_close.setProperty("variant", "primary")
        self.btn_save_close.clicked.connect(self._save_and_close)
        bottom.addWidget(self.btn_open_excel)
        bottom.addWidget(self.btn_save_close)
        layout.addLayout(bottom)

        QtGui.QShortcut(QtGui.QKeySequence.StandardKey.Save, self, activated=self._save_and_close)
        self.table.itemChanged.connect(self._on_table_item_changed)

        self._load_rows()

    def _smart_type_options(self) -> list[tuple[str, str]]:
        opts: list[tuple[str, str]] = []
        seen: set[str] = set()
        values = be.TERMS_SMART_TYPE_CHOICES or ["", "auto", "number", "date", "time", "title"]
        for opt in values:
            val = opt or ""
            if val in seen:
                continue
            seen.add(val)
            if val == "":
                opts.append(("Auto-detect (blank)", ""))
            elif val == "number":
                opts.append(("Numeric", "number"))
            elif val == "title":
                opts.append(("Title / Text", "title"))
            else:
                opts.append((val.capitalize(), val))
        if not opts:
            opts.append(("Auto-detect (blank)", ""))
        return opts

    def _determine_visible_headers(self, headers: list[str]) -> None:
        order: list[str] = []
        present = set(headers)
        for col in self.VISIBLE_COLUMN_ORDER:
            if col in present and col not in order:
                order.append(col)
        for h in headers:
            if h in self.DEFAULT_HIDDEN:
                continue
            if h not in order:
                order.append(h)
        self._visible_headers = [h for h in order if h in headers and h not in self.DEFAULT_HIDDEN]
        self._hidden_headers = [h for h in headers if h not in self._visible_headers]

    def _load_rows(self) -> None:
        headers, rows = be.read_terms_rows(self._terms_path)
        self._loading = True
        try:
            self._all_headers = headers
            self._determine_visible_headers(headers)
            self._hidden_rows = []
            self.table.clear()
            self.table.setColumnCount(len(self._visible_headers))
            header_labels = [self._column_display_name(h) for h in self._visible_headers]
            self.table.setHorizontalHeaderLabels(header_labels)
            self.table.setRowCount(0)
            for row in rows:
                idx = self.table.rowCount()
                self.table.insertRow(idx)
                self._hidden_rows.append({h: row.get(h, "") for h in self._hidden_headers})
                self._populate_row(idx, row)
        finally:
            self._loading = False
        self._dirty = False
        self._status_label.setText("All changes saved")

        # Set custom column widths - double width for Data Group and Term Label
        self.table.resizeColumnsToContents()
        for col_idx, header in enumerate(self._visible_headers):
            if header in ("Data Group", "Term Label"):
                current_width = self.table.columnWidth(col_idx)
                self.table.setColumnWidth(col_idx, current_width * 2)

    def _column_display_name(self, header: str) -> str:
        return self.COLUMN_DISPLAY_NAMES.get(header, header)

    def _populate_row(self, row_idx: int, row_data: dict[str, str]) -> None:
        for col_idx, header in enumerate(self._visible_headers):
            value = row_data.get(header, "")
            if header in self._combo_defs:
                combo = self._build_combo(header, value)
                self.table.setCellWidget(row_idx, col_idx, combo)
            else:
                item = QtWidgets.QTableWidgetItem(value)
                self.table.setItem(row_idx, col_idx, item)

    def _build_combo(self, header: str, value: str) -> QtWidgets.QComboBox:
        config = self._combo_defs[header]
        combo = QtWidgets.QComboBox()
        combo.setEditable(False)
        combo.setStyleSheet("""
            QComboBox {
                background-color: #ffffff;
                border: 1px solid #94a3b8;
                border-radius: 4px;
                padding: 4px;
                color: #0f172a;
            }
            QComboBox::drop-down {
                border: none;
            }
            QComboBox QAbstractItemView {
                background-color: #ffffff;
                color: #0f172a;
                selection-background-color: #bfdbfe;
                selection-color: #0f172a;
            }
        """)
        seen_values = set()
        for label, val in config["options"]:
            combo.addItem(label, val)
            seen_values.add(val)
        if value and value not in seen_values:
            combo.addItem(value, value)
        target = value if value else config.get("default", "")
        combo.blockSignals(True)
        idx = combo.findData(target)
        combo.setCurrentIndex(idx if idx >= 0 else 0)
        combo.blockSignals(False)
        combo.currentIndexChanged.connect(self._on_combo_changed)
        return combo

    def _on_combo_changed(self, *args) -> None:
        if self._loading:
            return
        self._mark_dirty()

    def _on_table_item_changed(self, item: QtWidgets.QTableWidgetItem) -> None:
        if self._loading:
            return
        self._mark_dirty()

    def _mark_dirty(self) -> None:
        if self._dirty:
            return
        self._dirty = True
        self._status_label.setText("Unsaved changes")

    def _selected_rows(self) -> list[int]:
        rows = {idx.row() for idx in self.table.selectionModel().selectedRows()}
        return sorted(rows)

    def _on_vertical_header_clicked(self, logical_index: int) -> None:
        try:
            self.table.selectRow(logical_index)
        except Exception:
            pass

    def _add_blank_row(self) -> None:
        payload = {h: "" for h in self._all_headers}
        payload.setdefault("Mode", (be.TERMS_MODE_CHOICES[0] if be.TERMS_MODE_CHOICES else "smart"))
        self._insert_row(payload)
        self.table.scrollToBottom()
        self._mark_dirty()

    def _duplicate_row(self) -> None:
        rows = self._selected_rows()
        if not rows:
            return
        source = rows[0]
        payload = self._combine_row_payload(source)
        self._insert_row(payload)
        self.table.scrollToBottom()
        self._mark_dirty()

    def _delete_rows(self) -> None:
        rows = self._selected_rows()
        if not rows:
            return
        for row in reversed(rows):
            self.table.removeRow(row)
            if 0 <= row < len(self._hidden_rows):
                self._hidden_rows.pop(row)
        if self.table.rowCount() == 0:
            self._insert_row({h: "" for h in self._all_headers})
        self._mark_dirty()

    def _move_rows(self, direction: int) -> None:
        rows = self._selected_rows()
        if not rows:
            return
        max_row = self.table.rowCount() - 1
        if max_row <= 0:
            return
        if direction < 0 and rows[0] == 0:
            return
        if direction > 0 and rows[-1] >= max_row:
            return
        payloads = [self._combine_row_payload(i) for i in range(self.table.rowCount())]
        if direction < 0:
            for row in rows:
                payloads[row - 1], payloads[row] = payloads[row], payloads[row - 1]
        else:
            for row in reversed(rows):
                payloads[row + 1], payloads[row] = payloads[row], payloads[row + 1]
        self._reset_table_from_payloads(payloads)
        sel = self.table.selectionModel()
        sel.clearSelection()
        new_rows = [row + direction for row in rows]
        for row in new_rows:
            index = self.table.model().index(row, 0)
            sel.select(
                index,
                QtCore.QItemSelectionModel.SelectionFlag.Select | QtCore.QItemSelectionModel.SelectionFlag.Rows,
            )
        if new_rows:
            self.table.scrollTo(self.table.model().index(new_rows[0], 0))
        self._mark_dirty()

    def _insert_row(self, payload: dict[str, str]) -> None:
        self._loading = True
        try:
            idx = self.table.rowCount()
            self.table.insertRow(idx)
            hidden_payload = {h: payload.get(h, "") for h in self._hidden_headers}
            if idx <= len(self._hidden_rows):
                self._hidden_rows.insert(idx, hidden_payload)
            else:
                self._hidden_rows.append(hidden_payload)
            self._populate_row(idx, payload)
        finally:
            self._loading = False

    def _reset_table_from_payloads(self, payloads: list[dict[str, str]]) -> None:
        self._loading = True
        try:
            self.table.setRowCount(0)
            self._hidden_rows = []
            if not payloads:
                payloads = [{h: "" for h in self._all_headers}]
            for payload in payloads:
                idx = self.table.rowCount()
                self.table.insertRow(idx)
                self._hidden_rows.append({h: payload.get(h, "") for h in self._hidden_headers})
                self._populate_row(idx, payload)
        finally:
            self._loading = False

    def _row_data_from_table(self, row_idx: int) -> dict[str, str]:
        data: dict[str, str] = {}
        for col_idx, header in enumerate(self._visible_headers):
            widget = self.table.cellWidget(row_idx, col_idx)
            if isinstance(widget, QtWidgets.QComboBox):
                current = widget.currentData()
                val = current if current is not None else widget.currentText()
                data[header] = str(val)
                continue
            item = self.table.item(row_idx, col_idx)
            data[header] = item.text() if item else ""
        return data

    def _combine_row_payload(self, row_idx: int) -> dict[str, str]:
        payload: dict[str, str] = {}
        payload.update(self._hidden_rows[row_idx] if row_idx < len(self._hidden_rows) else {})
        payload.update(self._row_data_from_table(row_idx))
        for header in self._all_headers:
            payload.setdefault(header, "")
        return payload

    def _gather_rows(self) -> list[dict[str, str]]:
        rows_out: list[dict[str, str]] = []
        for row_idx in range(self.table.rowCount()):
            combined = self._combine_row_payload(row_idx)
            if not any((combined.get(h, "") or "").strip() for h in self._visible_headers if h in combined and h != "Return"):
                continue
            existing_ret = combined.get("Return")
            combined["Return"] = be.derive_return_value(combined, existing=existing_ret)
            if row_idx < len(self._hidden_rows):
                self._hidden_rows[row_idx]["Return"] = combined["Return"]
            rows_out.append({h: combined.get(h, "") for h in self._all_headers})
        if not rows_out:
            rows_out.append({h: "" for h in self._all_headers})
        return rows_out

    def _save_rows(self) -> bool:
        try:
            rows = self._gather_rows()
            be.write_terms_rows(rows, path=self._terms_path, headers=self._all_headers)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Save failed", str(exc))
            return False
        self._dirty = False
        self._status_label.setText("All changes saved")
        self._pending_save_notice = True
        return True

    def _save_and_close(self) -> None:
        """Save changes and close the dialog."""
        if self._save_rows():
            self.accept()  # Close with success code

    def _open_in_excel(self) -> None:
        """Open the terms file in Excel (or default spreadsheet application)."""
        try:
            import subprocess
            import platform

            # Save any pending changes first
            if self._dirty:
                resp = QtWidgets.QMessageBox.question(
                    self,
                    "Save before opening?",
                    "You have unsaved changes. Save before opening in Excel?",
                    QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No | QtWidgets.QMessageBox.StandardButton.Cancel,
                )
                if resp == QtWidgets.QMessageBox.StandardButton.Cancel:
                    return
                elif resp == QtWidgets.QMessageBox.StandardButton.Yes:
                    if not self._save_rows():
                        return

            file_path = str(self._terms_path.resolve())

            # Open file with default application
            if platform.system() == 'Windows':
                subprocess.Popen(['start', 'excel', file_path], shell=True)
            elif platform.system() == 'Darwin':  # macOS
                subprocess.Popen(['open', file_path])
            else:  # Linux
                subprocess.Popen(['xdg-open', file_path])

            self._status_label.setText("Opened in Excel")
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Open failed", f"Could not open file in Excel:\n{exc}")

    def reject(self) -> None:  # type: ignore[override]
        if self._dirty:
            resp = QtWidgets.QMessageBox.question(
                self,
                "Discard unsaved edits?",
                "You have unsaved changes. Close without saving?",
                QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No,
            )
            if resp != QtWidgets.QMessageBox.StandardButton.Yes:
                return
        super().reject()

    def keyPressEvent(self, event: QtGui.QKeyEvent) -> None:  # type: ignore[override]
        if event.matches(QtGui.QKeySequence.StandardKey.Save):
            if self._save_rows():
                event.accept()
                return
        super().keyPressEvent(event)


class SeriesDropdown(QtWidgets.QWidget):
    """Dropdown button with checkable items for selecting multiple series."""

    changed = QtCore.Signal()

    def __init__(self, options: list[dict], selected: list[str] | None = None, parent=None):
        super().__init__(parent)
        self._options = options or []
        self._actions: dict[str, QtGui.QAction] = {}

        layout = QtWidgets.QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        self.button = QtWidgets.QToolButton()
        self.button.setText("Select Series")
        self.button.setMinimumWidth(140)
        self.button.setMinimumHeight(32)
        self.button.setPopupMode(QtWidgets.QToolButton.ToolButtonPopupMode.InstantPopup)
        self.menu = QtWidgets.QMenu(self)
        self.button.setMenu(self.menu)

        self.summary = QtWidgets.QLabel("No series selected")
        self.summary.setStyleSheet("color: #5b6b7a;")
        self.summary.setMinimumWidth(260)

        layout.addWidget(self.button)
        layout.addWidget(self.summary, 1)

        self.set_options(self._options, selected or [])

    def set_options(self, options: list[dict], selected: list[str] | None = None) -> None:
        self._options = options or []
        existing = set(selected or self.selected_series())
        self.menu.clear()
        self._actions = {}
        for opt in self._options:
            name = opt.get("name") or ""
            if not name:
                continue
            act = QtGui.QAction(name, self.menu)
            act.setCheckable(True)
            act.setChecked(name in existing)
            act.toggled.connect(self._on_action_toggled)
            self.menu.addAction(act)
            self._actions[name] = act
        self._update_summary()

    def selected_series(self) -> list[str]:
        return [name for name, act in self._actions.items() if act.isChecked()]

    def set_selected(self, names: list[str]) -> None:
        target = set(names or [])
        for name, act in self._actions.items():
            act.blockSignals(True)
            act.setChecked(name in target)
            act.blockSignals(False)
        self._update_summary()

    def _on_action_toggled(self, _: bool) -> None:
        self._update_summary()
        self.changed.emit()

    def _update_summary(self) -> None:
        sel = self.selected_series()
        self.summary.setText(", ".join(sel) if sel else "No series selected")


class PlotRowWidget(QtWidgets.QWidget):
    """One proposed plot row with editable fields and dropdown series selector."""

    changed = QtCore.Signal()

    def __init__(self, series_options: list[dict], data: dict | None = None, parent=None):
        super().__init__(parent)
        data = data or {}
        self._series_options = series_options or []
        self._series_axis_map: dict[str, str] = {}
        self._manual_y_axis = bool(str(data.get("y_axis") or "").strip())
        self._x_axis = str(data.get("x_axis") or "SN").strip() or "SN"

        layout = QtWidgets.QGridLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setHorizontalSpacing(8)

        self.name_edit = QtWidgets.QLineEdit(str(data.get("name") or ""))
        self.name_edit.setPlaceholderText("Plot name")
        self.name_edit.setMinimumWidth(260)
        self.name_edit.setMinimumHeight(34)
        self.name_edit.setSizePolicy(QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Fixed)
        self.name_edit.textChanged.connect(self.changed.emit)
        layout.addWidget(self.name_edit, 0, 0)

        self.y_axis_edit = QtWidgets.QLineEdit(str(data.get("y_axis") or ""))
        self.y_axis_edit.setPlaceholderText("Y axis label")
        self.y_axis_edit.setMinimumWidth(220)
        self.y_axis_edit.setMinimumHeight(34)
        self.y_axis_edit.setSizePolicy(QtWidgets.QSizePolicy.Policy.Expanding, QtWidgets.QSizePolicy.Policy.Fixed)
        self.y_axis_edit.textEdited.connect(self._handle_y_axis_edited)
        layout.addWidget(self.y_axis_edit, 0, 1)

        self.series_dropdown = SeriesDropdown(self._series_options, list(data.get("series") or []))
        self.series_dropdown.changed.connect(self._on_series_toggled)
        layout.addWidget(self.series_dropdown, 0, 2)

        layout.setColumnStretch(0, 3)
        layout.setColumnStretch(1, 3)
        layout.setColumnStretch(2, 4)

        self.update_series_options(self._series_options, initial_selection=list(data.get("series") or []))
        if not self._manual_y_axis:
            self._auto_update_y_axis()

    def _handle_y_axis_edited(self, text: str) -> None:
        self._manual_y_axis = bool(text.strip())
        self.changed.emit()

    def _set_y_axis_text(self, value: str) -> None:
        current = self.y_axis_edit.text()
        if current == value:
            return
        self.y_axis_edit.blockSignals(True)
        self.y_axis_edit.setText(value)
        self.y_axis_edit.blockSignals(False)
        self.changed.emit()

    def selected_series(self) -> list[str]:
        return self.series_dropdown.selected_series()

    def update_series_options(self, options: list[dict], initial_selection: list[str] | None = None) -> None:
        self._series_options = options or []
        if initial_selection is None:
            initial_selection = self.selected_series()
        self._series_axis_map = {}
        for opt in self._series_options:
            name = opt.get("name")
            if not name:
                continue
            axis_value = (
                str(opt.get("default_y_axis") or "").strip()
                or str(opt.get("units") or "").strip()
            )
            self._series_axis_map[name] = axis_value
        self.series_dropdown.set_options(self._series_options, initial_selection)
        self._auto_update_y_axis()

    def _on_series_toggled(self, _: bool) -> None:
        self._auto_update_y_axis()
        self.changed.emit()

    def _auto_update_y_axis(self) -> None:
        if self._manual_y_axis:
            return
        candidates: list[str] = []
        for name in self.selected_series():
            axis = self._series_axis_map.get(name, "")
            if axis:
                candidates.append(axis)
        unique: list[str] = []
        seen: set[str] = set()
        for axis in candidates:
            key = axis.lower()
            if key and key not in seen:
                seen.add(key)
                unique.append(axis)
        if len(unique) == 1:
            self._set_y_axis_text(unique[0])
        elif not candidates:
            self._set_y_axis_text("")

    def to_dict(self) -> dict:
        name = self.name_edit.text().strip() or "Plot"
        return {
            "name": name,
            "series": self.selected_series(),
            "y_axis": self.y_axis_edit.text().strip(),
            "x_axis": self._x_axis,
        }


class ProposedPlotsDialog(QtWidgets.QDialog):
    """Popup dialog for managing proposed plot definitions."""

    def __init__(self, series_options: list[dict], plots: list[dict], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Manage Proposed Plots")
        self.resize(1100, 560)
        self.setObjectName("proposedPlotsDialog")
        self.setStyleSheet("""
            #proposedPlotsDialog {
                background-color: #f8fafc;
                color: #0f172a;
            }
            #proposedPlotsDialog QLabel {
                color: #0f172a;
            }
            #proposedPlotsDialog QListWidget {
                background-color: #ffffff;
                border: 1px solid #d4d4d8;
                color: #0f172a;
            }
            #proposedPlotsDialog QPushButton {
                background-color: #ffffff;
                color: #0f172a;
                border: 1px solid #cbd5f5;
                border-radius: 6px;
                padding: 8px 14px;
            }
            #proposedPlotsDialog QPushButton:hover {
                background-color: #f1f5f9;
            }
        """)
        self._series_options = series_options or []
        self._plots: list[dict] = [dict(p) for p in plots] if plots else []

        layout = QtWidgets.QVBoxLayout(self)
        intro = QtWidgets.QLabel(
            "Add plots, choose the plot series from the dropdown, and optionally override the Y-axis label."
        )
        intro.setWordWrap(True)
        layout.addWidget(intro)

        self.list_widget = QtWidgets.QListWidget()
        self.list_widget.setSpacing(6)
        self.list_widget.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.SingleSelection)
        layout.addWidget(self.list_widget, 1)

        controls = QtWidgets.QHBoxLayout()
        self.btn_add = QtWidgets.QPushButton("Add Plot")
        self.btn_remove = QtWidgets.QPushButton("Remove Selected")
        controls.addWidget(self.btn_add)
        controls.addWidget(self.btn_remove)
        controls.addStretch(1)
        layout.addLayout(controls)

        btn_box = QtWidgets.QDialogButtonBox(
            QtWidgets.QDialogButtonBox.StandardButton.Save | QtWidgets.QDialogButtonBox.StandardButton.Close
        )
        btn_box.accepted.connect(self._save_and_close)
        btn_box.rejected.connect(self.reject)
        layout.addWidget(btn_box)

        self.btn_add.clicked.connect(self._add_plot)
        self.btn_remove.clicked.connect(self._remove_selected)

        if self._plots:
            for plot in self._plots:
                self._append_row(plot)
        else:
            self._add_plot()

    def _append_row(self, data: dict | None = None) -> PlotRowWidget:
        item = QtWidgets.QListWidgetItem()
        widget = PlotRowWidget(self._series_options, data=data)
        widget.changed.connect(self._sync_item_size)
        item.setSizeHint(widget.sizeHint())
        self.list_widget.addItem(item)
        self.list_widget.setItemWidget(item, widget)
        return widget

    def _add_plot(self):
        name = self._generate_plot_name()
        widget = self._append_row({"name": name, "series": [], "y_axis": "", "x_axis": "SN"})
        widget.name_edit.setFocus(QtCore.Qt.FocusReason.OtherFocusReason)
        self.list_widget.setCurrentRow(self.list_widget.count() - 1)

    def _remove_selected(self):
        row = self.list_widget.currentRow()
        if row < 0:
            return
        item = self.list_widget.takeItem(row)
        if item:
            widget = self.list_widget.itemWidget(item)
            if widget:
                widget.deleteLater()
            del item
        if self.list_widget.count() == 0:
            self._add_plot()

    def _generate_plot_name(self) -> str:
        existing = set()
        for i in range(self.list_widget.count()):
            widget = self._row_widget(i)
            if widget:
                existing.add(widget.name_edit.text().strip())
        idx = max(1, len(existing) + 1)
        while True:
            candidate = f"Plot {idx}"
            if candidate not in existing:
                return candidate
            idx += 1

    def _row_widget(self, index: int) -> PlotRowWidget | None:
        item = self.list_widget.item(index)
        if not item:
            return None
        widget = self.list_widget.itemWidget(item)
        return widget if isinstance(widget, PlotRowWidget) else None

    def _sync_item_size(self):
        for i in range(self.list_widget.count()):
            item = self.list_widget.item(i)
            widget = self.list_widget.itemWidget(item)
            if widget:
                item.setSizeHint(widget.sizeHint())

    def _collect_plots(self) -> list[dict]:
        collected: list[dict] = []
        for i in range(self.list_widget.count()):
            widget = self._row_widget(i)
            if widget:
                collected.append(widget.to_dict())
        return collected

    def _save_and_close(self):
        self._plots = self._collect_plots()
        self.accept()

    def plots(self) -> list[dict]:
        return self._plots


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("EIDAT Prototype - Demonstration Only")
        self.resize(1280, 860)

        be.ensure_scaffold()
        self._refresh_plot_series_after_worker = False
        self._auto_update_plot_terms_on_success = False
        self._plot_terms_pending = False
        self._plot_terms_pending_reason = ""

        # Global styling - polished modern design with rounded corners
        self.setStyleSheet(
            """
            QMainWindow {
                background: #f0f4f8;
            }
            QWidget {
                background: transparent;
            }
            QLabel {
                color: #1f2937;
            }
            QLabel.subtle {
                color: #6b7280;
                font-size: 12px;
            }

            QLineEdit {
                background: #ffffff;
                border: 1px solid #d1d5db;
                border-radius: 8px;
                padding: 10px 14px;
                color: #374151;
                font-size: 13px;
            }
            QLineEdit:focus {
                border-color: #2563eb;
                border-width: 2px;
            }

            /* Improve QMessageBox legibility */
            QMessageBox {
                background-color: #ffffff;
                border-radius: 12px;
            }
            QMessageBox QLabel {
                color: #1f2937;
            }
            QMessageBox QPushButton {
                padding: 8px 20px;
                border-radius: 8px;
                background: #2563eb;
                color: #ffffff;
                border: none;
                font-weight: 600;
            }
            QMessageBox QPushButton:hover {
                background: #1d4ed8;
            }

            /* Plain text edit for debug console with rounded corners */
            QPlainTextEdit {
                background: #1f2937;
                color: #e5e7eb;
                border: 1px solid #374151;
                border-radius: 10px;
                font-family: 'Consolas', 'Monaco', monospace;
                font-size: 11px;
                padding: 12px;
            }

            /* Scrollbars */
            QScrollBar:vertical {
                border: none;
                background: #f3f4f6;
                width: 12px;
                border-radius: 6px;
                margin: 0px;
            }
            QScrollBar::handle:vertical {
                background: #d1d5db;
                border-radius: 6px;
                min-height: 20px;
            }
            QScrollBar::handle:vertical:hover {
                background: #9ca3af;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
            }

            QScrollBar:horizontal {
                border: none;
                background: #f3f4f6;
                height: 12px;
                border-radius: 6px;
                margin: 0px;
            }
            QScrollBar::handle:horizontal {
                background: #d1d5db;
                border-radius: 6px;
                min-width: 20px;
            }
            QScrollBar::handle:horizontal:hover {
                background: #9ca3af;
            }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                width: 0px;
            }
            """
        )

        central = QtWidgets.QWidget()
        self.setCentralWidget(central)

        # Header with gradient background - logo and tabs inline with modern design
        header = QtWidgets.QFrame()
        header.setObjectName("heroHeader")
        header.setStyleSheet("""
            #heroHeader {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #ffffff, stop:0.5 #f8fafc, stop:1 #ffffff);
                border: none;
                margin: 0px;
                padding: 0px;
            }
        """)
        hbox = QtWidgets.QHBoxLayout(header)
        hbox.setContentsMargins(24, 20, 24, 16)
        hbox.setSpacing(20)

        # Logo - simple and compact
        logo_container = QtWidgets.QFrame()
        logo_container.setObjectName("logoBadge")
        logo_container.setStyleSheet("""
            #logoBadge {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                    stop:0 #2563eb, stop:1 #1e40af);
                border-radius: 12px;
                border: none;
            }
        """)
        shadow = QtWidgets.QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(24)
        shadow.setOffset(0, 6)
        shadow.setColor(QtGui.QColor(15, 23, 42, 90))
        logo_container.setGraphicsEffect(shadow)
        logo_container.setFixedSize(48, 48)
        logo_layout = QtWidgets.QVBoxLayout(logo_container)
        logo_layout.setContentsMargins(0, 0, 0, 0)

        logo = QtWidgets.QLabel()
        logo_pix = self._build_logo_pixmap(size=44)
        logo.setPixmap(logo_pix)
        logo.setFixedSize(44, 44)
        logo.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        logo_layout.addWidget(logo)

        hbox.addWidget(logo_container)

        # Title section with improved typography
        title = QtWidgets.QLabel("EIDAT")
        font = title.font(); font.setPointSize(22); font.setBold(True); font.setLetterSpacing(QtGui.QFont.SpacingType.AbsoluteSpacing, 0.5); title.setFont(font)
        title.setStyleSheet("color: #0f172a; padding: 0px;")
        subtitle = QtWidgets.QLabel("End Item Data Analysis Tool (Prototype)")
        subtitle.setStyleSheet("color:#64748b; font-size: 12px; font-weight: 500; letter-spacing: 0.3px; border:none; background:transparent;")
        proto_badge = QtWidgets.QLabel("Prototype build - demonstration only")
        proto_badge.setStyleSheet("color:#991b1b; background:#fee2e2; border:1px solid #fecaca; border-radius:6px; font-size:11px; font-weight:600; padding:2px 8px; letter-spacing:0.5px;")
        proto_badge.setAlignment(QtCore.Qt.AlignmentFlag.AlignLeft)
        tbox = QtWidgets.QVBoxLayout();
        tbox.setSpacing(4)
        tbox.addWidget(title); tbox.addWidget(subtitle); tbox.addWidget(proto_badge)
        hbox.addLayout(tbox)

        hbox.addStretch(1)

        # Create clean, minimal tab buttons - larger and right-aligned
        self.tab_buttons = QtWidgets.QWidget()
        tab_btn_layout = QtWidgets.QHBoxLayout(self.tab_buttons)
        tab_btn_layout.setContentsMargins(0, 0, 0, 0)
        tab_btn_layout.setSpacing(0)

        self.btn_tab_setup = QtWidgets.QPushButton("⚙  Setup")
        self.btn_tab_process = QtWidgets.QPushButton("📄  EIDP Processing")
        self.btn_tab_plot = QtWidgets.QPushButton("📊  Analysis")

        for btn in [self.btn_tab_setup, self.btn_tab_process, self.btn_tab_plot]:
            btn.setCheckable(True)
            btn.setStyleSheet("""
                QPushButton {
                    padding: 14px 48px;
                    margin: 0;
                    font-weight: 500;
                    font-size: 15px;
                    color: #6b7280;
                    background: transparent;
                    border: none;
                    border-bottom: 3px solid transparent;
                }
                QPushButton:checked {
                    color: #2563eb;
                    font-weight: 600;
                    background: transparent;
                    border-bottom: 3px solid #2563eb;
                }
                QPushButton:hover:!checked {
                    color: #374151;
                    background: rgba(59, 130, 246, 0.05);
                    border-bottom: 3px solid #cbd5e1;
                }
            """)
            tab_btn_layout.addWidget(btn)

        self.btn_tab_setup.setChecked(True)
        hbox.addWidget(self.tab_buttons)

        # Store status label but don't add it to layout (hidden)
        self.lbl_ready = QtWidgets.QLabel("● System Ready");
        self.lbl_ready.setObjectName("statusBadge");
        self.lbl_ready.setStyleSheet("""
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #d1fae5, stop:1 #a7f3d0);
            color: #065f46;
            border-radius: 12px;
            padding: 6px 16px;
            font-size: 12px;
            font-weight: 700;
            border: 2px solid #10b981;
        """)
        self.lbl_ready.setVisible(False)

        # Create tabs widget (hidden, only used for content management)
        self.tabs = QtWidgets.QTabWidget()
        self.tab_setup = QtWidgets.QWidget()
        self.tab_process = QtWidgets.QWidget()
        self.tab_plot = QtWidgets.QWidget()
        self.tab_outputs = QtWidgets.QWidget()
        self.tabs.addTab(self.tab_setup, "Setup")
        self.tabs.addTab(self.tab_process, "EIDP Processing")
        self.tabs.addTab(self.tab_plot, "Analysis")

        # Connect tab buttons to switch content
        self.btn_tab_setup.clicked.connect(lambda: self._switch_tab(0))
        self.btn_tab_process.clicked.connect(lambda: self._switch_tab(1))
        self.btn_tab_plot.clicked.connect(lambda: self._switch_tab(2))

        self.tabs.currentChanged.connect(self._on_tab_changed)

        # Tab pane styling - clean look with no border-radius to match flat tabs
        self.tabs.setStyleSheet(
            """
            QTabWidget::pane {
                border: 1px solid #e5e7eb;
                border-radius: 0;
                border-top: none;
                margin: 0px;
                background: #ffffff;
                padding: 16px;
            }
            QTabBar::tab {
                width: 0px;
                height: 0px;
                margin: 0px;
                padding: 0px;
                border: none;
            }
            """
        )

        self.log = QtWidgets.QPlainTextEdit(); self.log.setReadOnly(True); self.log.setMaximumBlockCount(5000)
        # Toggle to show/hide the debug log panel on demand - styled with rounded corners
        self.btn_toggle_log = QtWidgets.QPushButton("\u25B6  Debug Console")
        self.btn_toggle_log.setCheckable(True)
        self.btn_toggle_log.setChecked(False)
        self.btn_toggle_log.clicked.connect(self._toggle_log_panel)
        self.btn_toggle_log.setStyleSheet("""
            QPushButton {
                background: #374151;
                color: #e5e7eb;
                border: 1px solid #4b5563;
                border-radius: 8px;
                padding: 10px 16px;
                text-align: left;
                font-size: 12px;
                font-weight: 500;
                margin: 8px;
            }
            QPushButton:hover {
                background: #4b5563;
            }
            QPushButton:checked {
                background: #1f2937;
                border-color: #374151;
            }
        """)
        pol_log = self.btn_toggle_log.sizePolicy(); pol_log.setHorizontalStretch(1); pol_log.setHorizontalPolicy(QtWidgets.QSizePolicy.Policy.Expanding); self.btn_toggle_log.setSizePolicy(pol_log)
        self.status_bar = self.statusBar()
        self._progress_dialog = RunProgressDialog(self)
        self._progress_dialog.canceled.connect(self._on_progress_canceled)
        # Pattern supports both old format (without Found) and new format (with Found)
        self._progress_pattern = re.compile(r"\[PROGRESS\]\s*Terms:\s*(\d+)%\s*\((\d+)/(\d+)\)(?:\s*\|\s*Found:\s*(\d+))?")
        self._progress_total = 0
        self._progress_completed = 0
        self._progress_found = 0
        self._progress_popup_active = False
        self._progress_was_canceled = False
        self._last_run_dir: Path | None = None

        # Create toast notification widget
        self._toast = ToastNotification(self)
        self._toast.hide()

        layout = QtWidgets.QVBoxLayout(central)
        layout.addWidget(header)
        layout.addWidget(self.tabs)
        layout.addWidget(self.btn_toggle_log)
        layout.addWidget(self.log, 1)
        self.log.setVisible(False)
        # Build tabs
        self._setup_tab_setup()
        self._setup_tab_process()
        self._setup_tab_plot()
        # Data Outputs tab replaced by top-level master button

        # Runtime
        self._worker: ProcWorker | None = None
        self._enrich_after_run: bool = False
        self._registry_cache: tuple[list[str], list[list[str]]] | None = None
        self._scan_refresh()
        # Periodic auto-sync every few minutes (no popup, no compile)
        try:
            self._sync_timer = QtCore.QTimer(self)
            self._sync_timer.setInterval(5 * 60 * 1000)  # 5 minutes
            self._sync_timer.timeout.connect(lambda: self._sync_workspace(auto=True))
            self._sync_timer.start()
        except Exception:
            pass

    # Tabs
    def _setup_tab_setup(self):
        grid = QtWidgets.QGridLayout(self.tab_setup)
        grid.setContentsMargins(24, 24, 24, 24)
        grid.setSpacing(16)

        # Program Health
        grp_env = QtWidgets.QGroupBox("Program Health")
        grp_env.setStyleSheet("""
            QGroupBox {
                font-weight: 900;
                font-size: 24px;
                border: 1px solid #e5e7eb;
                border-radius: 8px;
                margin-top: 16px;
                background: #ffffff;
                padding: 16px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 4px 8px;
                color: #111827;
            }
        """)
        l_env = QtWidgets.QVBoxLayout(grp_env)
        l_env.setSpacing(12)

        # Header row with description and badge
        header_row = QtWidgets.QHBoxLayout()
        desc_label = QtWidgets.QLabel("Ensure all dependencies are installed and up to date")
        desc_label.setStyleSheet("color: #6b7280; font-size: 13px; font-weight: 400;")
        header_row.addWidget(desc_label)
        header_row.addStretch()

        self.lbl_env_health = QtWidgets.QLabel("Healthy")
        self.lbl_env_health.setObjectName("healthBadge")
        self.lbl_env_health.setStyleSheet("background: #d1fae5; color: #065f46; border-radius: 12px; padding: 4px 12px; font-size: 12px; font-weight: 600;")
        header_row.addWidget(self.lbl_env_health)
        l_env.addLayout(header_row)

        # Button row
        button_row = QtWidgets.QHBoxLayout()
        button_row.setSpacing(12)
        self.btn_check = QtWidgets.QPushButton("Check Environment")
        self.btn_check.setStyleSheet("""
            QPushButton {
                padding: 10px 20px;
                border-radius: 6px;
                background: #ffffff;
                color: #374151;
                border: 1px solid #d1d5db;
                font-size: 13px;
            }
            QPushButton:hover {
                background: #f9fafb;
                border-color: #9ca3af;
            }
        """)
        self.btn_install = QtWidgets.QPushButton("Update Packages")
        self.btn_install.setStyleSheet("""
            QPushButton {
                padding: 10px 20px;
                border-radius: 6px;
                background: #2563eb;
                color: #ffffff;
                border: 1px solid #2563eb;
                font-size: 13px;
                font-weight: 600;
            }
            QPushButton:hover {
                background: #1d4ed8;
            }
            QPushButton:disabled {
                background: #93c5fd;
                border-color: #93c5fd;
            }
        """)
        self.btn_check.clicked.connect(self._act_check_env)
        self.btn_install.clicked.connect(self._act_install)
        button_row.addWidget(self.btn_check)
        button_row.addWidget(self.btn_install)
        button_row.addStretch()
        l_env.addLayout(button_row)

        # Environment path display
        self.lbl_env = QtWidgets.QLabel("Env: Unknown")
        self.lbl_env.setStyleSheet("color: #6b7280; font-size: 12px; font-weight: 400; padding: 8px; background: #f9fafb; border-radius: 4px;")
        self.lbl_env.setWordWrap(True)
        l_env.addWidget(self.lbl_env)

        # Extraction Settings
        grp_set = QtWidgets.QGroupBox("Extraction Settings")
        grp_set.setStyleSheet("""
            QGroupBox {
                font-weight: 900;
                font-size: 24px;
                border: 1px solid #e5e7eb;
                border-radius: 8px;
                margin-top: 16px;
                background: #ffffff;
                padding: 16px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 4px 8px;
                color: #111827;
            }
        """)
        ls = QtWidgets.QVBoxLayout(grp_set)
        ls.setSpacing(16)

        # Description
        hint = QtWidgets.QLabel("Configure how EIDAT processes and extracts data from documents")
        hint.setStyleSheet("color: #6b7280; font-size: 13px; font-weight: 400;")
        ls.addWidget(hint)

        # OCR Mode
        ocr_container = QtWidgets.QWidget()
        ocr_layout = QtWidgets.QVBoxLayout(ocr_container)
        ocr_layout.setContentsMargins(0, 0, 0, 0)
        ocr_layout.setSpacing(6)
        ocr_label = QtWidgets.QLabel("OCR Mode")
        ocr_label.setStyleSheet("color: #374151; font-size: 13px; font-weight: 500;")
        ocr_layout.addWidget(ocr_label)
        # Friendly OCR mode labels mapped to env values
        self._ocr_value_to_display = {
            "fallback": "Read PDF and fallback to OCR if needed",
            "ocr_only": "Only OCR read the selected documents",
            "no_ocr": "Read PDF, no OCR (may fail)",
        }
        self._ocr_display_to_value = {v: k for k, v in self._ocr_value_to_display.items()}
        self.cmb_ocr_mode = QtWidgets.QComboBox()
        self.cmb_ocr_mode.setStyleSheet("""
            QComboBox {
                background: #ffffff;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                padding: 8px 12px;
                color: #374151;
                font-size: 13px;
                min-height: 20px;
            }
            QComboBox:hover {
                border-color: #9ca3af;
            }
            QComboBox::drop-down {
                border: none;
                width: 20px;
            }
            QComboBox QAbstractItemView {
                color: #374151;
                background: #ffffff;
                selection-background-color: #dbeafe;
                border: 1px solid #d1d5db;
                padding: 4px;
            }
        """)
        self.cmb_ocr_mode.addItems(list(self._ocr_value_to_display.values()))
        self.cmb_ocr_mode.setSizeAdjustPolicy(QtWidgets.QComboBox.SizeAdjustPolicy.AdjustToContents)
        # Enable scrollbar for dropdown if needed
        self.cmb_ocr_mode.view().setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        ocr_layout.addWidget(self.cmb_ocr_mode)
        ls.addWidget(ocr_container)

        # XY Fuzz Tolerance with info icon
        xy_container = QtWidgets.QWidget()
        xy_layout = QtWidgets.QVBoxLayout(xy_container)
        xy_layout.setContentsMargins(0, 0, 0, 0)
        xy_layout.setSpacing(6)
        xy_header = QtWidgets.QHBoxLayout()
        xy_label = QtWidgets.QLabel("XY Fuzz Tolerance")
        xy_label.setStyleSheet("color: #374151; font-size: 13px; font-weight: 500;")
        xy_header.addWidget(xy_label)
        xy_info = QtWidgets.QLabel("\u24D8")  # Info icon
        xy_info.setStyleSheet("color: #9ca3af; font-size: 14px;")
        xy_info.setToolTip("Tolerance for matching table cell positions (higher = more lenient)")
        xy_header.addWidget(xy_info)
        xy_header.addStretch()
        self.lbl_xy_val = QtWidgets.QLabel("0.46")
        self.lbl_xy_val.setStyleSheet("color: #111827; font-size: 14px; font-weight: 600;")
        xy_header.addWidget(self.lbl_xy_val)
        xy_layout.addLayout(xy_header)
        self.sld_xy_fuzz = QtWidgets.QSlider(QtCore.Qt.Orientation.Horizontal)
        self.sld_xy_fuzz.setRange(0, 100)
        self.sld_xy_fuzz.setStyleSheet("""
            QSlider::groove:horizontal {
                height: 6px;
                background: #e5e7eb;
                border-radius: 3px;
            }
            QSlider::handle:horizontal {
                width: 16px;
                height: 16px;
                margin: -5px 0;
                border-radius: 8px;
                background: #2563eb;
            }
            QSlider::handle:horizontal:hover {
                background: #1d4ed8;
            }
        """)
        xy_layout.addWidget(self.sld_xy_fuzz)
        ls.addWidget(xy_container)

        # OCR DPI
        dpi_container = QtWidgets.QWidget()
        dpi_layout = QtWidgets.QVBoxLayout(dpi_container)
        dpi_layout.setContentsMargins(0, 0, 0, 0)
        dpi_layout.setSpacing(6)
        dpi_header = QtWidgets.QHBoxLayout()
        dpi_label = QtWidgets.QLabel("OCR DPI")
        dpi_label.setStyleSheet("color: #374151; font-size: 13px; font-weight: 500;")
        dpi_header.addWidget(dpi_label)
        dpi_info = QtWidgets.QLabel("\u24D8")
        dpi_info.setStyleSheet("color: #9ca3af; font-size: 14px;")
        dpi_info.setToolTip("Higher DPI may improve OCR accuracy at the cost of speed")
        dpi_header.addWidget(dpi_info)
        dpi_header.addStretch()
        self.lbl_dpi_val = QtWidgets.QLabel("763")
        self.lbl_dpi_val.setStyleSheet("color: #111827; font-size: 14px; font-weight: 600;")
        dpi_header.addWidget(self.lbl_dpi_val)
        dpi_layout.addLayout(dpi_header)
        self.sld_ocr_dpi = QtWidgets.QSlider(QtCore.Qt.Orientation.Horizontal)
        self.sld_ocr_dpi.setRange(500, 1000)
        self.sld_ocr_dpi.setSingleStep(25)
        self.sld_ocr_dpi.setStyleSheet("""
            QSlider::groove:horizontal {
                height: 6px;
                background: #e5e7eb;
                border-radius: 3px;
            }
            QSlider::handle:horizontal {
                width: 16px;
                height: 16px;
                margin: -5px 0;
                border-radius: 8px;
                background: #2563eb;
            }
            QSlider::handle:horizontal:hover {
                background: #1d4ed8;
            }
        """)
        dpi_layout.addWidget(self.sld_ocr_dpi)
        ls.addWidget(dpi_container)

        # Show detailed debug logs toggle
        debug_container = QtWidgets.QWidget()
        debug_layout = QtWidgets.QHBoxLayout(debug_container)
        debug_layout.setContentsMargins(0, 0, 0, 0)
        debug_layout.setSpacing(8)
        debug_left = QtWidgets.QVBoxLayout()
        debug_left.setSpacing(2)
        debug_title = QtWidgets.QLabel("Show detailed debug logs")
        debug_title.setStyleSheet("color: #374151; font-size: 13px; font-weight: 500;")
        debug_desc = QtWidgets.QLabel("(slightly slower)")
        debug_desc.setStyleSheet("color: #9ca3af; font-size: 12px;")
        debug_left.addWidget(debug_title)
        debug_left.addWidget(debug_desc)
        debug_layout.addLayout(debug_left)
        debug_layout.addStretch()
        self.chk_logging = QtWidgets.QCheckBox()
        self.chk_logging.setStyleSheet("""
            QCheckBox::indicator {
                width: 40px;
                height: 20px;
            }
            QCheckBox::indicator:unchecked {
                border-radius: 10px;
                background: #d1d5db;
            }
            QCheckBox::indicator:checked {
                border-radius: 10px;
                background: #2563eb;
            }
        """)
        debug_layout.addWidget(self.chk_logging)
        ls.addWidget(debug_container)

        # OCR Language
        lang_container = QtWidgets.QWidget()
        lang_layout = QtWidgets.QVBoxLayout(lang_container)
        lang_layout.setContentsMargins(0, 0, 0, 0)
        lang_layout.setSpacing(6)
        lang_label = QtWidgets.QLabel("OCR Language")
        lang_label.setStyleSheet("color: #374151; font-size: 13px; font-weight: 500;")
        lang_layout.addWidget(lang_label)
        # Language display mapping (user-friendly names)
        self._lang_display_to_code = {
            "English": "en",
            "French": "fr",
            "German": "de",
            "Spanish": "es",
        }
        self._lang_code_to_display = {v: k for k, v in self._lang_display_to_code.items()}
        self.cmb_lang = QtWidgets.QComboBox()
        self.cmb_lang.setStyleSheet("""
            QComboBox {
                background: #ffffff;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                padding: 8px 12px;
                color: #374151;
                font-size: 13px;
                min-height: 20px;
            }
            QComboBox:hover {
                border-color: #9ca3af;
            }
            QComboBox::drop-down {
                border: none;
                width: 20px;
            }
            QComboBox QAbstractItemView {
                color: #374151;
                background: #ffffff;
                selection-background-color: #dbeafe;
                border: 1px solid #d1d5db;
                padding: 4px;
            }
        """)
        self.cmb_lang.addItems(list(self._lang_display_to_code.keys()))
        self.cmb_lang.setSizeAdjustPolicy(QtWidgets.QComboBox.SizeAdjustPolicy.AdjustToContents)
        # Enable scrollbar for dropdown if needed
        self.cmb_lang.view().setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        lang_layout.addWidget(self.cmb_lang)
        ls.addWidget(lang_container)

        # Persist on change
        self.cmb_ocr_mode.currentTextChanged.connect(self._persist_settings_from_panel)
        self.sld_xy_fuzz.valueChanged.connect(self._on_xy_slider)
        self.sld_ocr_dpi.valueChanged.connect(self._on_dpi_slider)
        self.chk_logging.stateChanged.connect(self._persist_settings_from_panel)
        self.cmb_lang.currentTextChanged.connect(self._persist_settings_from_panel)

        grid.addWidget(grp_env, 0, 0, 1, 2)
        grid.addWidget(grp_set, 1, 0, 1, 2)
        grid.setRowStretch(2, 1)

    def _setup_tab_process(self):
        # Main container with two columns
        main_layout = QtWidgets.QHBoxLayout(self.tab_process)
        main_layout.setContentsMargins(24, 24, 24, 24)
        main_layout.setSpacing(16)

        # Left column
        left_column = QtWidgets.QVBoxLayout()
        left_column.setSpacing(16)

        # === Master Database Section ===
        grp_master = QtWidgets.QGroupBox("Master Database")
        grp_master.setStyleSheet("""
            QGroupBox {
                font-weight: 900;
                font-size: 24px;
                border: 1px solid #e5e7eb;
                border-radius: 8px;
                margin-top: 16px;
                background: #ffffff;
                padding: 16px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 4px 8px;
                color: #111827;
            }
        """)
        master_layout = QtWidgets.QVBoxLayout(grp_master)
        master_layout.setSpacing(12)

        master_desc = QtWidgets.QLabel("Access and manage the central EIDAT database")
        master_desc.setStyleSheet("color: #6b7280; font-size: 13px; font-weight: 400;")
        master_layout.addWidget(master_desc)

        self.btn_open_master_tab = QtWidgets.QPushButton("\U0001F5C4  Open Master Database")
        self.btn_open_master_tab.setStyleSheet("""
            QPushButton {
                padding: 10px 20px;
                border-radius: 6px;
                background: #ffffff;
                color: #374151;
                border: 1px solid #d1d5db;
                font-size: 13px;
                text-align: center;
            }
            QPushButton:hover {
                background: #f9fafb;
                border-color: #9ca3af;
            }
        """)
        self.btn_open_master_tab.clicked.connect(lambda: self._safe_open(be.open_master_workbook))
        master_layout.addWidget(self.btn_open_master_tab)

        left_column.addWidget(grp_master)

        # === Processing Controls Section (formerly Define Inputs) ===
        grp_inputs = QtWidgets.QGroupBox("Processing Controls")
        grp_inputs.setStyleSheet("""
            QGroupBox {
                font-weight: 900;
                font-size: 24px;
                border: 1px solid #e5e7eb;
                border-radius: 8px;
                margin-top: 16px;
                background: #ffffff;
                padding: 16px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 4px 8px;
                color: #111827;
            }
        """)
        inputs_layout = QtWidgets.QVBoxLayout(grp_inputs)
        inputs_layout.setSpacing(12)

        inputs_desc = QtWidgets.QLabel("Configure extraction terms and execute processing operations")
        inputs_desc.setStyleSheet("color: #6b7280; font-size: 13px; font-weight: 400;")
        inputs_layout.addWidget(inputs_desc)

        self.btn_terms_edit = QtWidgets.QPushButton("\u270E  Edit Smart-Snap Terms")
        self.btn_terms_edit.setStyleSheet("""
            QPushButton {
                padding: 10px 20px;
                border-radius: 6px;
                background: #2563eb;
                color: #ffffff;
                border: 1px solid #2563eb;
                font-size: 13px;
                font-weight: 600;
            }
            QPushButton:hover {
                background: #1d4ed8;
            }
            QPushButton:disabled {
                background: #93c5fd;
                border-color: #93c5fd;
            }
        """)
        self.btn_terms_edit.clicked.connect(self._open_terms_editor)
        inputs_layout.addWidget(self.btn_terms_edit)

        self.btn_terms_refresh = QtWidgets.QPushButton("\U0001F4C4  Create/Refresh Input Spreadsheet")
        self.btn_terms_refresh.setStyleSheet("""
            QPushButton {
                padding: 10px 20px;
                border-radius: 6px;
                background: #ffffff;
                color: #374151;
                border: 1px solid #d1d5db;
                font-size: 13px;
            }
            QPushButton:hover {
                background: #f9fafb;
                border-color: #9ca3af;
            }
        """)
        self.btn_terms_refresh.clicked.connect(self._act_generate_terms)
        inputs_layout.addWidget(self.btn_terms_refresh)

        # Add separator
        separator = QtWidgets.QFrame()
        separator.setFrameShape(QtWidgets.QFrame.Shape.HLine)
        separator.setStyleSheet("background-color: #e5e7eb; margin: 8px 0;")
        inputs_layout.addWidget(separator)

        # Smart Processing Controls (moved from bottom)
        proc_label = QtWidgets.QLabel("Batch Processing")
        proc_label.setStyleSheet("color: #374151; font-size: 13px; font-weight: 600; margin-top: 4px;")
        inputs_layout.addWidget(proc_label)

        self.btn_start = QtWidgets.QPushButton("\u25B6  Extract and Update All")
        self.btn_start.setStyleSheet("""
            QPushButton {
                padding: 12px 20px;
                border-radius: 6px;
                background: #2563eb;
                color: #ffffff;
                border: 1px solid #2563eb;
                font-size: 13px;
                font-weight: 600;
            }
            QPushButton:hover {
                background: #1d4ed8;
            }
            QPushButton:disabled {
                background: #93c5fd;
                border-color: #93c5fd;
            }
        """)
        self.btn_start.clicked.connect(self._show_extraction_options)  # Updated to show options
        inputs_layout.addWidget(self.btn_start)

        proc_secondary = QtWidgets.QHBoxLayout()
        proc_secondary.setSpacing(8)

        self.btn_stop = QtWidgets.QPushButton("\u23F9  Stop Scan")
        self.btn_stop.setStyleSheet("""
            QPushButton {
                padding: 10px 16px;
                border-radius: 6px;
                background: #ffffff;
                color: #374151;
                border: 1px solid #d1d5db;
                font-size: 13px;
            }
            QPushButton:hover {
                background: #f9fafb;
                border-color: #9ca3af;
            }
        """)
        self.btn_stop.clicked.connect(self._act_stop_scan)

        self.btn_open_last = QtWidgets.QPushButton("\U0001F4C2  Open Last Run Folder")
        self.btn_open_last.setStyleSheet("""
            QPushButton {
                padding: 10px 16px;
                border-radius: 6px;
                background: #ffffff;
                color: #374151;
                border: 1px solid #d1d5db;
                font-size: 13px;
            }
            QPushButton:hover {
                background: #f9fafb;
                border-color: #9ca3af;
            }
        """)
        self.btn_open_last.clicked.connect(lambda: self._safe_open(be.open_last_run_folder))

        proc_secondary.addWidget(self.btn_stop)
        proc_secondary.addWidget(self.btn_open_last)
        inputs_layout.addLayout(proc_secondary)

        # Table Extraction Section
        separator2 = QtWidgets.QFrame()
        separator2.setFrameShape(QtWidgets.QFrame.Shape.HLine)
        separator2.setStyleSheet("background-color: #e5e7eb; margin: 8px 0;")
        inputs_layout.addWidget(separator2)

        table_label = QtWidgets.QLabel("Table Extraction")
        table_label.setStyleSheet("color: #374151; font-size: 13px; font-weight: 600; margin-top: 4px;")
        inputs_layout.addWidget(table_label)

        self.btn_extract_tables = QtWidgets.QPushButton("\U0001F4CA  Extract Tables from PDFs")
        self.btn_extract_tables.setStyleSheet("""
            QPushButton {
                padding: 12px 20px;
                border-radius: 6px;
                background: #10b981;
                color: #ffffff;
                border: 1px solid #10b981;
                font-size: 13px;
                font-weight: 600;
            }
            QPushButton:hover {
                background: #059669;
            }
            QPushButton:disabled {
                background: #86efac;
                border-color: #86efac;
            }
        """)
        self.btn_extract_tables.clicked.connect(self._show_table_extraction_dialog)
        inputs_layout.addWidget(self.btn_extract_tables)

        # Keep internal fields for logic
        self.ed_terms = QtWidgets.QLineEdit(str(be.DEFAULT_TERMS_XLSX))
        self.ed_terms.setVisible(False)
        self.ed_pdfs = QtWidgets.QLineEdit(str(be.DEFAULT_PDF_DIR))
        self.ed_pdfs.setVisible(False)

        left_column.addWidget(grp_inputs)
        left_column.addStretch(1)

        # Right column
        right_column = QtWidgets.QVBoxLayout()
        right_column.setSpacing(16)

        # === Data Upload Section ===
        grp_upload = QtWidgets.QGroupBox("Data Controls and Status")
        grp_upload.setStyleSheet("""
            QGroupBox {
                font-weight: 900;
                font-size: 24px;
                border: 1px solid #e5e7eb;
                border-radius: 8px;
                margin-top: 16px;
                background: #ffffff;
                padding: 16px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 4px 8px;
                color: #111827;
            }
        """)
        upload_layout = QtWidgets.QVBoxLayout(grp_upload)
        upload_layout.setSpacing(12)

        upload_desc = QtWidgets.QLabel("Sync workspace and update the EIDAT database")
        upload_desc.setStyleSheet("color: #6b7280; font-size: 13px; font-weight: 400;")
        upload_layout.addWidget(upload_desc)

        self.btn_sync_workspace = QtWidgets.QPushButton("\u2B73  Sync Workspace Now")
        self.btn_sync_workspace.setStyleSheet("""
            QPushButton {
                padding: 10px 20px;
                border-radius: 6px;
                background: #2563eb;
                color: #ffffff;
                border: 1px solid #2563eb;
                font-size: 13px;
                font-weight: 600;
            }
            QPushButton:hover {
                background: #1d4ed8;
            }
            QPushButton:disabled {
                background: #93c5fd;
                border-color: #93c5fd;
            }
        """)
        self.btn_sync_workspace.clicked.connect(self._act_sync_workspace)
        upload_layout.addWidget(self.btn_sync_workspace)

        # Repository Root
        repo_label = QtWidgets.QLabel("Repository Root")
        repo_label.setStyleSheet("color: #374151; font-size: 13px; font-weight: 500; margin-top: 8px;")
        upload_layout.addWidget(repo_label)

        repo_row = QtWidgets.QHBoxLayout()
        repo_row.setSpacing(8)
        self.ed_repo = QtWidgets.QLineEdit(str(getattr(be, 'get_repo_root', lambda: be.DEFAULT_REPO_ROOT)()))
        self.ed_repo.setStyleSheet("""
            QLineEdit {
                background: #ffffff;
                border: 1px solid #d1d5db;
                border-radius: 6px;
                padding: 8px 12px;
                color: #374151;
                font-size: 13px;
            }
            QLineEdit:focus {
                border-color: #2563eb;
            }
        """)
        btn_repo = QtWidgets.QPushButton("Browse...")
        btn_repo.setStyleSheet("""
            QPushButton {
                padding: 8px 16px;
                border-radius: 6px;
                background: #ffffff;
                color: #374151;
                border: 1px solid #d1d5db;
                font-size: 13px;
            }
            QPushButton:hover {
                background: #f9fafb;
                border-color: #9ca3af;
            }
        """)
        btn_repo.clicked.connect(lambda: self._browse_folder(self.ed_repo, be.DEFAULT_PDF_DIR))
        repo_row.addWidget(self.ed_repo, 1)
        repo_row.addWidget(btn_repo)
        upload_layout.addLayout(repo_row)

        # Sync banner/status
        self.lbl_sync_banner = QtWidgets.QLabel("No sync run yet.")
        self.lbl_sync_banner.setObjectName("syncBanner")
        self.lbl_sync_banner.setWordWrap(True)
        self.lbl_sync_banner.setStyleSheet(
            "background: #fef3c7; color: #92400e; border: 1px solid #fcd34d; border-radius: 6px; padding: 10px 12px; font-size: 12px;"
        )
        upload_layout.addWidget(self.lbl_sync_banner)

        self.btn_view_outdated = QtWidgets.QPushButton("\U0001F4CB  View Data Package List and Update EIDAT Database")
        self.btn_view_outdated.setStyleSheet("""
            QPushButton {
                padding: 10px 20px;
                border-radius: 6px;
                background: #2563eb;
                color: #ffffff;
                border: 1px solid #2563eb;
                font-size: 13px;
                font-weight: 600;
            }
            QPushButton:hover {
                background: #1d4ed8;
            }
        """)
        self.btn_view_outdated.clicked.connect(self._show_outdated_popup)
        upload_layout.addWidget(self.btn_view_outdated)

        # Secondary buttons row
        secondary_row = QtWidgets.QHBoxLayout()
        secondary_row.setSpacing(8)

        self.btn_view_registry2 = QtWidgets.QPushButton("\U0001F4D6  View Registry")
        self.btn_view_registry2.setStyleSheet("""
            QPushButton {
                padding: 8px 16px;
                border-radius: 6px;
                background: #ffffff;
                color: #374151;
                border: 1px solid #d1d5db;
                font-size: 13px;
            }
            QPushButton:hover {
                background: #f9fafb;
                border-color: #9ca3af;
            }
        """)
        self.btn_view_registry2.clicked.connect(self._act_view_registry)

        self.btn_clear_old_runs = QtWidgets.QPushButton("\U0001F5D1  Clear Old Run Cache")
        self.btn_clear_old_runs.setStyleSheet("""
            QPushButton {
                padding: 8px 16px;
                border-radius: 6px;
                background: #ffffff;
                color: #374151;
                border: 1px solid #d1d5db;
                font-size: 13px;
            }
            QPushButton:hover {
                background: #f9fafb;
                border-color: #9ca3af;
            }
        """)
        self.btn_clear_old_runs.clicked.connect(self._act_clear_old_runs)

        secondary_row.addWidget(self.btn_view_registry2)
        secondary_row.addWidget(self.btn_clear_old_runs)
        upload_layout.addLayout(secondary_row)

        right_column.addWidget(grp_upload)
        right_column.addStretch(1)

        # Add columns to main layout
        main_layout.addLayout(left_column, 1)
        main_layout.addLayout(right_column, 1)
    def _setup_tab_plot(self):
        main_layout = QtWidgets.QVBoxLayout(self.tab_plot)
        main_layout.setContentsMargins(24, 24, 24, 24)
        main_layout.setSpacing(16)

        # === Plotting Dashboard Section ===
        grp_plot = QtWidgets.QGroupBox("Plotting Dashboard")
        grp_plot.setStyleSheet("""
            QGroupBox {
                font-weight: 900;
                font-size: 24px;
                border: 1px solid #e5e7eb;
                border-radius: 8px;
                margin-top: 16px;
                background: #ffffff;
                padding: 16px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 4px 8px;
                color: #111827;
            }
        """)
        plot_layout = QtWidgets.QVBoxLayout(grp_plot)
        plot_layout.setSpacing(16)

        # Description
        desc = QtWidgets.QLabel("Manage terms, configure plots, and generate reports")
        desc.setStyleSheet("color: #6b7280; font-size: 13px; font-weight: 400;")
        plot_layout.addWidget(desc)

        # === Terms Management Subsection ===
        terms_label = QtWidgets.QLabel("Terms Management")
        terms_label.setStyleSheet("color: #111827; font-size: 14px; font-weight: 600; margin-top: 8px;")
        plot_layout.addWidget(terms_label)

        self.btn_open_plot_terms = QtWidgets.QPushButton("\U0001F441  View Terms List")
        self.btn_open_plot_terms.setStyleSheet("""
            QPushButton {
                padding: 10px 20px;
                border-radius: 6px;
                background: #ffffff;
                color: #374151;
                border: 1px solid #d1d5db;
                font-size: 13px;
                text-align: left;
                padding-left: 16px;
            }
            QPushButton:hover {
                background: #f9fafb;
                border-color: #9ca3af;
            }
        """)
        self.btn_open_plot_terms.clicked.connect(lambda: self._safe_open(lambda: be.open_path(be.DEFAULT_PLOT_TERMS_XLSX)))
        plot_layout.addWidget(self.btn_open_plot_terms)

        # Separator
        separator1 = QtWidgets.QFrame()
        separator1.setFrameShape(QtWidgets.QFrame.Shape.HLine)
        separator1.setStyleSheet("background-color: #e5e7eb; margin: 8px 0;")
        plot_layout.addWidget(separator1)

        # === Proposed Plots Subsection ===
        plots_label = QtWidgets.QLabel("Plots")
        plots_label.setStyleSheet("color: #111827; font-size: 14px; font-weight: 600;")
        plot_layout.addWidget(plots_label)

        plots_desc = QtWidgets.QLabel("Manage and configure plot series: Isp1 Pre Test Functional - Specific Impulse")
        plots_desc.setStyleSheet("color: #6b7280; font-size: 12px; font-style: italic;")
        plots_desc.setWordWrap(True)
        self.lbl_proposed_summary = plots_desc  # Keep reference for updates
        plot_layout.addWidget(plots_desc)

        # Proposed plots buttons side by side
        proposed_row = QtWidgets.QHBoxLayout()
        proposed_row.setSpacing(12)

        self.btn_manage_proposed = QtWidgets.QPushButton("\u2699  Manage Plots...")
        self.btn_manage_proposed.setStyleSheet("""
            QPushButton {
                padding: 10px 20px;
                border-radius: 6px;
                background: #ffffff;
                color: #374151;
                border: 1px solid #d1d5db;
                font-size: 13px;
            }
            QPushButton:hover {
                background: #f9fafb;
                border-color: #9ca3af;
            }
        """)
        self.btn_manage_proposed.clicked.connect(self._open_proposed_plots_dialog)

        self.btn_refresh_series = QtWidgets.QPushButton("\U0001F504  Reload Series List")
        self.btn_refresh_series.setStyleSheet("""
            QPushButton {
                padding: 10px 20px;
                border-radius: 6px;
                background: #ffffff;
                color: #374151;
                border: 1px solid #d1d5db;
                font-size: 13px;
            }
            QPushButton:hover {
                background: #f9fafb;
                border-color: #9ca3af;
            }
        """)
        self.btn_refresh_series.clicked.connect(self._refresh_series_catalog)

        proposed_row.addWidget(self.btn_manage_proposed)
        proposed_row.addWidget(self.btn_refresh_series)
        proposed_row.addStretch()
        plot_layout.addLayout(proposed_row)

        # Separator
        separator2 = QtWidgets.QFrame()
        separator2.setFrameShape(QtWidgets.QFrame.Shape.HLine)
        separator2.setStyleSheet("background-color: #e5e7eb; margin: 8px 0;")
        plot_layout.addWidget(separator2)

        # === Generate Plots Subsection ===
        generate_label = QtWidgets.QLabel("Generate Plots")
        generate_label.setStyleSheet("color: #111827; font-size: 14px; font-weight: 600;")
        plot_layout.addWidget(generate_label)

        # Generate plots buttons
        generate_row = QtWidgets.QHBoxLayout()
        generate_row.setSpacing(12)

        self.btn_plots_generate = QtWidgets.QPushButton("\U0001F4CA  Generate Plots")
        self.btn_plots_generate.setStyleSheet("""
            QPushButton {
                padding: 12px 24px;
                border-radius: 6px;
                background: #2563eb;
                color: #ffffff;
                border: 1px solid #2563eb;
                font-size: 13px;
                font-weight: 600;
            }
            QPushButton:hover {
                background: #1d4ed8;
            }
            QPushButton:disabled {
                background: #93c5fd;
                border-color: #93c5fd;
            }
        """)
        self.btn_plots_generate.clicked.connect(self._act_generate_plots)

        self.btn_plot_summary = QtWidgets.QPushButton("\U0001F4C4  Plot Summary Report")
        self.btn_plot_summary.setStyleSheet("""
            QPushButton {
                padding: 12px 24px;
                border-radius: 6px;
                background: #ffffff;
                color: #374151;
                border: 1px solid #d1d5db;
                font-size: 13px;
            }
            QPushButton:hover {
                background: #f9fafb;
                border-color: #9ca3af;
            }
        """)
        self.btn_plot_summary.clicked.connect(self._act_export_plot_summary)

        self.btn_plots_open_folder = QtWidgets.QPushButton("\U0001F4C2  Open Plots Folder")
        self.btn_plots_open_folder.setStyleSheet("""
            QPushButton {
                padding: 12px 24px;
                border-radius: 6px;
                background: #ffffff;
                color: #374151;
                border: 1px solid #d1d5db;
                font-size: 13px;
            }
            QPushButton:hover {
                background: #f9fafb;
                border-color: #9ca3af;
            }
        """)
        self.btn_plots_open_folder.clicked.connect(lambda: self._safe_open(be.open_plots_folder))

        generate_row.addWidget(self.btn_plots_generate)
        generate_row.addWidget(self.btn_plot_summary)
        generate_row.addWidget(self.btn_plots_open_folder)
        generate_row.addStretch()
        plot_layout.addLayout(generate_row)

        main_layout.addWidget(grp_plot)
        main_layout.addStretch(1)

        self._plot_series_options: list[dict] = []
        self._proposed_plots_cache: list[dict] = []
        self._refresh_series_catalog()
        self._load_proposed_plots()

        self._plot_series_options: list[dict] = []
    def _setup_tab_outputs(self):
        grid = QtWidgets.QGridLayout(self.tab_outputs)
        grid.addWidget(QtWidgets.QLabel("Run Registry UI coming soon. Use buttons above to open registry and master."), 0, 0)
        grid.setRowStretch(1, 1)

    # Actions & helpers
    def _switch_tab(self, idx: int):
        """Switch to a tab by index and update button states."""
        # Update button checked states
        self.btn_tab_setup.setChecked(idx == 0)
        self.btn_tab_process.setChecked(idx == 1)
        self.btn_tab_plot.setChecked(idx == 2)
        # Switch the actual tab
        self.tabs.setCurrentIndex(idx)

    def _on_tab_changed(self, idx: int):
        try:
            if self.tabs.widget(idx) is self.tab_plot:
                self._refresh_series_catalog()
            if self.tabs.widget(idx) is self.tab_process or self.tabs.widget(idx) is self.tab_outputs:
                self._refresh_run_registry()
        except Exception:
            pass

    def _append_log(self, text: str):
        self.log.appendPlainText(text)
        if not self.log.isVisible():
            try:
                label = self.btn_toggle_log.text()
                if "\u2022" not in label and "•" not in label:
                    self.btn_toggle_log.setText("Show Debug Panel •")
            except Exception:
                pass
        self.log.verticalScrollBar().setValue(self.log.verticalScrollBar().maximum())

    def _on_worker_line(self, text: str):
        self._append_log(text)
        self._maybe_update_run_progress(text)
        self._maybe_track_run_dir(text)

    def _maybe_update_run_progress(self, text: str):
        if "[PROGRESS] Terms" not in text:
            return
        match = self._progress_pattern.search(text)
        if not match:
            return
        try:
            completed = int(match.group(2))
            total = int(match.group(3))
            # Group 4 is optional (for backward compatibility with old format)
            found_str = match.group(4)
            found = int(found_str) if found_str is not None else 0
        except Exception:
            return
        self._progress_total = max(total, 0)
        self._progress_completed = max(0, min(completed, self._progress_total or completed))
        self._progress_found = max(0, found)
        self._update_progress_widgets()

    def _maybe_track_run_dir(self, text: str):
        if "Outputs will be saved under:" not in text:
            return
        try:
            _, tail = text.split("Outputs will be saved under:", 1)
        except ValueError:
            return
        candidate = tail.strip().strip('"')
        if not candidate:
            return
        path = Path(candidate)
        if not path.is_absolute():
            path = Path(be.ROOT) / path
        self._last_run_dir = path

    def _on_progress_canceled(self):
        if not self._progress_popup_active or self._progress_was_canceled:
            return
        self._progress_was_canceled = True
        self._enrich_after_run = False
        try:
            self.lbl_ready.setText("Stopping run...")
        except Exception:
            pass
        self._append_log("[GUI] User requested run abort.")
        self._progress_dialog.lbl_status.setText("Stopping run...")
        self._act_stop_scan()

    def _update_progress_widgets(self):
        if not self._progress_popup_active:
            return
        total = self._progress_total
        completed = min(self._progress_completed, total if total else self._progress_completed)
        found = self._progress_found
        self._progress_dialog.update_progress(completed, total, found)

    def _finalize_run_progress(self, success: bool):
        if not self._progress_popup_active:
            return
        if self._progress_total and self._progress_completed < self._progress_total:
            self._progress_completed = self._progress_total
        self._update_progress_widgets()
        final_success = success and not self._progress_was_canceled
        if self._progress_was_canceled:
            message = "Run aborted"
        elif final_success:
            message = "Run complete"
        else:
            message = "Run finished with errors"
        self._progress_dialog.finish(message, success=final_success)
        self._progress_popup_active = False
        self._progress_was_canceled = False
        if final_success:
            self._last_run_dir = None

    def _cleanup_last_run_dir(self):
        path = self._last_run_dir
        if not path:
            return
        try:
            runs_root = Path(be.RUNS_DIR).resolve()
        except Exception:
            runs_root = Path(be.RUNS_DIR)
        try:
            target = path.resolve()
        except Exception:
            target = path
        if runs_root not in target.parents and target != runs_root:
            self._last_run_dir = None
            return
        try:
            shutil.rmtree(target, ignore_errors=True)
        finally:
            self._last_run_dir = None

    def _schedule_plot_terms_update(self, reason: str):
        label = "Refreshing available plot terms"
        if reason:
            label += f" ({reason})"
        if self._worker is not None and self._worker.isRunning():
            self._plot_terms_pending = True
            self._plot_terms_pending_reason = reason
            return
        self._plot_terms_pending = False
        self._plot_terms_pending_reason = ""
        self._refresh_plot_series_after_worker = True
        self._start_worker(be.generate_plot_terms, status_msg=label)

    def _start_worker(self, popen_factory, *, status_msg: str, show_run_progress: bool = False, refresh_plot_terms_after: bool = False):
        if self._worker is not None and self._worker.isRunning():
            return
        self._auto_update_plot_terms_on_success = bool(refresh_plot_terms_after)
        self._append_log(f"[GUI] {status_msg}")
        self.status_bar.showMessage(status_msg)
        self._progress_total = 0
        self._progress_completed = 0
        self._progress_found = 0
        self._progress_popup_active = show_run_progress
        self._progress_was_canceled = False
        self._last_run_dir = None
        # Track if this is an extraction run (show_run_progress indicates EIDP extraction)
        self._is_extraction_run = show_run_progress
        if show_run_progress:
            self._progress_dialog.begin(status_msg)
        else:
            self._progress_dialog.abort()
        self._worker = ProcWorker(popen_factory)
        self._worker.line.connect(self._on_worker_line)
        self._worker.finished.connect(self._on_worker_done)
        self._worker.start()

    def _toggle_log_panel(self):
        try:
            visible = bool(self.btn_toggle_log.isChecked())
        except Exception:
            visible = True
        self.log.setVisible(visible)
        try:
            # Update arrow direction based on visibility
            self.btn_toggle_log.setText("\u25BC  Debug Console" if visible else "\u25B6  Debug Console")
        except Exception:
            pass

    def _on_worker_done(self, rc: int):
        self._worker = None
        self.status_bar.showMessage("Ready.", 3000)
        self._append_log(f"[INFO] Process finished with code {rc}")
        was_canceled = self._progress_was_canceled
        is_extraction = getattr(self, "_is_extraction_run", False)
        self._finalize_run_progress(success=(rc == 0))

        # Show toast notification only for extraction runs
        if is_extraction:
            if was_canceled:
                self._show_toast("Extraction run aborted")
            elif rc == 0:
                self._show_toast("Extraction completed successfully")
            else:
                self._show_toast(f"Extraction failed with code {rc}")

        if was_canceled:
            self._cleanup_last_run_dir()

        # Refresh UI after a scan; no post-run enrichment step
        if getattr(self, "_enrich_after_run", False):
            self._enrich_after_run = False
            try:
                self._refresh_run_registry()
            except Exception:
                pass
        if getattr(self, "_refresh_plot_series_after_worker", False):
            self._refresh_plot_series_after_worker = False
            try:
                self._refresh_series_catalog()
            except Exception:
                pass
        auto_refresh = self._auto_update_plot_terms_on_success and not was_canceled and rc == 0
        self._auto_update_plot_terms_on_success = False
        if auto_refresh:
            self._schedule_plot_terms_update("latest scan")

        # Update master database after successful extraction run
        if is_extraction and rc == 0 and not was_canceled:
            try:
                self._append_log("[GUI] Compiling master database after extraction run...")
                be.compile_master()
                self._show_toast("Master database updated")
            except Exception as e:
                self._append_log(f"[ERROR] Failed to compile master database: {e}")

        self._scan_refresh()
        if self._plot_terms_pending and (self._worker is None or not self._worker.isRunning()):
            reason = self._plot_terms_pending_reason or "refresh"
            self._plot_terms_pending = False
            self._plot_terms_pending_reason = ""
            self._schedule_plot_terms_update(reason)

    def _act_install(self):
        self._start_worker(be.run_install_full, status_msg="Installing environment...")

    def _act_check_env(self):
        self._start_worker(be.check_environment, status_msg="Checking environment...")

    def _act_open_env(self):
        try:
            be.SCANNER_ENV.parent.mkdir(parents=True, exist_ok=True)
            if not be.SCANNER_ENV.exists():
                be.save_scanner_env({"QUIET": "1"})
            be.open_path(be.SCANNER_ENV)
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Open failed", str(e))

    def _open_terms_editor(self):
        raw = (self.ed_terms.text() or "").strip()
        try:
            target = Path(raw).expanduser() if raw else be.DEFAULT_TERMS_XLSX
        except Exception:
            target = be.DEFAULT_TERMS_XLSX
        if not target.exists():
            resp = QtWidgets.QMessageBox.question(
                self,
                "Terms spreadsheet missing",
                f"{target} was not found.\nGenerate a fresh Smart-Snap template now?",
                QtWidgets.QMessageBox.StandardButton.Yes | QtWidgets.QMessageBox.StandardButton.No,
            )
            if resp == QtWidgets.QMessageBox.StandardButton.Yes:
                self._act_generate_terms()
            return
        try:
            dlg = TermsEditorDialog(target, self)
        except Exception as exc:
            QtWidgets.QMessageBox.warning(self, "Unable to open editor", str(exc))
            return
        result = dlg.exec()
        # If dialog was accepted (saved successfully), show toast and sync
        if result == QtWidgets.QDialog.DialogCode.Accepted:
            self._show_toast("Smart-Snap terms saved successfully")
            # Auto-sync workspace after successful save
            QtCore.QTimer.singleShot(500, lambda: self._sync_workspace(auto=False))

    def _act_generate_terms(self):
        try:
            target = Path(self.ed_terms.text()).expanduser()
        except Exception:
            target = be.DEFAULT_TERMS_XLSX if hasattr(be, "DEFAULT_TERMS_XLSX") else Path("terms.xlsx")
        if target.exists():
            msg = (
                f"A Smart-Snap terms workbook already exists:\n\n{target}\n\n"
                "Generating a new template will overwrite it.\n\nContinue?"
            )
            if (
                QtWidgets.QMessageBox.question(
                    self, "Overwrite Smart-Snap workbook?", msg
                )
                != QtWidgets.QMessageBox.StandardButton.Yes
            ):
                return
        self._start_worker(be.generate_terms, status_msg="Generating Smart-Snap terms spreadsheet...")

    def _show_extraction_options(self):
        """Show dialog with extraction options: out-of-date only, force all, or choose files."""
        dlg = QtWidgets.QDialog(self)
        dlg.setWindowTitle("Extraction Options")
        dlg.resize(500, 280)
        dlg.setStyleSheet("""
            QDialog {
                background: #ffffff;
            }
            QPushButton {
                padding: 12px 20px;
                border-radius: 6px;
                font-size: 13px;
                font-weight: 500;
            }
        """)

        layout = QtWidgets.QVBoxLayout(dlg)
        layout.setSpacing(16)
        layout.setContentsMargins(24, 24, 24, 24)

        # Title and description
        title = QtWidgets.QLabel("Choose Extraction Method")
        title.setStyleSheet("font-size: 18px; font-weight: 700; color: #111827;")
        layout.addWidget(title)

        desc = QtWidgets.QLabel("Select how you want to process your documents:")
        desc.setStyleSheet("font-size: 13px; color: #6b7280; margin-bottom: 8px;")
        layout.addWidget(desc)

        # Option 1: Extract out-of-date only
        btn_outdated = QtWidgets.QPushButton("\u26A1  Extract All Out-of-Date")
        btn_outdated.setStyleSheet("""
            QPushButton {
                background: #2563eb;
                color: #ffffff;
                border: 1px solid #2563eb;
                text-align: left;
                padding-left: 16px;
            }
            QPushButton:hover {
                background: #1d4ed8;
            }
        """)
        btn_outdated.setToolTip("Process only files marked as 'New' or 'Out-of-Date'")

        # Option 2: Force extract all
        btn_force = QtWidgets.QPushButton("\U0001F504  Force Extract All")
        btn_force.setStyleSheet("""
            QPushButton {
                background: #dc2626;
                color: #ffffff;
                border: 1px solid #dc2626;
                text-align: left;
                padding-left: 16px;
            }
            QPushButton:hover {
                background: #b91c1c;
            }
        """)
        btn_force.setToolTip("Re-run extraction on the entire database (may take longer)")

        # Option 3: Choose specific files
        btn_choose = QtWidgets.QPushButton("\U0001F4CB  Choose Files to Extract")
        btn_choose.setStyleSheet("""
            QPushButton {
                background: #ffffff;
                color: #374151;
                border: 1px solid #d1d5db;
                text-align: left;
                padding-left: 16px;
            }
            QPushButton:hover {
                background: #f9fafb;
                border-color: #9ca3af;
            }
        """)
        btn_choose.setToolTip("Manually select which files to process")

        layout.addWidget(btn_outdated)
        layout.addWidget(btn_force)
        layout.addWidget(btn_choose)

        # Cancel button
        layout.addStretch()
        btn_cancel = QtWidgets.QPushButton("Cancel")
        btn_cancel.setStyleSheet("""
            QPushButton {
                background: #ffffff;
                color: #6b7280;
                border: 1px solid #d1d5db;
            }
            QPushButton:hover {
                background: #f9fafb;
            }
        """)
        layout.addWidget(btn_cancel)

        # Connect buttons
        def extract_outdated():
            dlg.accept()
            # Get all out-of-date files (new or pdf_newer or terms_newer)
            details = getattr(self, "_sync_details", None) or []
            rows = [d for d in details if d.get("reason") in ("new", "pdf_newer", "terms_newer")]
            paths = [Path(d.get("pdf")) for d in rows if d.get("pdf")]
            if not paths:
                QtWidgets.QMessageBox.information(self, "Nothing to extract", "No out-of-date files found. Run 'Sync Workspace Now' first.")
                return
            try:
                terms = Path(self.ed_terms.text()).expanduser()
            except Exception:
                terms = be.DEFAULT_TERMS_XLSX
            self._enrich_after_run = True
            self._start_worker(lambda: be.run_selected_pdfs(paths, terms), status_msg="Extracting out-of-date files...", show_run_progress=True, refresh_plot_terms_after=True)

        def force_extract_all():
            dlg.accept()
            # Force run on all PDFs
            self._act_start_scan()

        def choose_files():
            dlg.accept()
            # Show the file selection dialog
            self._show_outdated_popup()

        btn_outdated.clicked.connect(extract_outdated)
        btn_force.clicked.connect(force_extract_all)
        btn_choose.clicked.connect(choose_files)
        btn_cancel.clicked.connect(dlg.reject)

        dlg.exec()

    def _act_start_scan(self):
        terms = Path(self.ed_terms.text()).expanduser()
        pdfs = Path(self.ed_pdfs.text()).expanduser()
        if not terms.exists():
            QtWidgets.QMessageBox.critical(self, "Missing terms", f"Terms file not found:\n{terms}")
            return
        if not pdfs.exists():
            QtWidgets.QMessageBox.critical(self, "Missing PDFs folder", f"PDFs folder not found:\n{pdfs}")
            return
        # Only enrich registry after a scan completes
        self._enrich_after_run = True
        self._start_worker(lambda: be.run_scanner(terms, pdfs), status_msg="Scanning PDFs...", show_run_progress=True, refresh_plot_terms_after=True)

    def _act_stop_scan(self):
        try:
            if self._worker and self._worker.isRunning():
                self._worker.stop()
        except Exception:
            pass

    def _act_compile_master(self):
        self._start_worker(be.compile_master, status_msg="Compiling master workbook...")

    def _act_generate_plot_terms(self):
        self._refresh_plot_series_after_worker = True
        self._start_worker(be.generate_plot_terms, status_msg="Generating plot terms...")

    def _act_generate_plots(self):
        try:
            plots_dir = be.PLOTS_DIR
            if plots_dir.exists():
                for img in plots_dir.glob("*.png"):
                    try:
                        img.unlink()
                    except Exception:
                        pass
        except Exception:
            pass
        self._start_worker(be.generate_plots, status_msg="Generating plots...")

    def _act_export_plot_summary(self):
        self._start_worker(be.export_plots_summary, status_msg="Exporting plot summary...")

    def _get_registry_table_data(self) -> tuple[list[str], list[list[str]]]:
        try:
            base = (be.ROOT / "Product_Data_File") if hasattr(be, "ROOT") else Path("Product_Data_File")
            p_csv = base / "run_registry.csv"
            p_xlsx = base / "run_registry.xlsx"
            rows: list[list[str]] = []
            headers: list[str] = []
            # Keep registry tidy before reading; best-effort and safe (prunes invalid rows)
            try:
                be.ensure_run_registry_consistent()
            except Exception:
                pass
            if p_csv.exists():
                import csv
                with open(p_csv, newline="", encoding="utf-8") as f:
                    reader = csv.reader(f)
                    for i, r in enumerate(reader):
                        if i == 0:
                            headers = [str(x) for x in r]
                        else:
                            rows.append([str(x) for x in r])
            elif p_xlsx.exists():
                # Convert xlsx -> csv once via backend and retry
                try:
                    be.ensure_run_registry_consistent()
                except Exception:
                    pass
                if p_csv.exists():
                    import csv
                    with open(p_csv, newline="", encoding="utf-8") as f:
                        reader = csv.reader(f)
                        for i, r in enumerate(reader):
                            if i == 0:
                                headers = [str(x) for x in r]
                            else:
                                rows.append([str(x) for x in r])
            else:
                headers = ["Run Registry"]
                rows = [["No registry found. Run a scan to create it."]]

            # Show file as-is (no derived columns) for a clean, authoritative view

            return headers, rows
        except Exception:
            return [], []

    def _act_view_registry(self):
        """Refresh cache and show the run registry dialog."""
        try:
            self._refresh_run_registry()
        except Exception:
            pass
        self._show_registry_popup()

    def _act_clear_old_runs(self):
        try:
            confirm = QtWidgets.QMessageBox.question(
                self,
                "Confirm cleanup",
                "Delete old run_data folders not referenced by the current registry?",
            )
            if confirm != QtWidgets.QMessageBox.StandardButton.Yes:
                return
            # Ensure registry is consistent before cleanup
            try:
                be.ensure_run_registry_consistent()
            except Exception:
                pass
            deleted, kept = be.clear_stale_run_data()
            QtWidgets.QMessageBox.information(
                self,
                "Cleanup complete",
                f"Removed {deleted} old run folder(s). Kept {kept} current folder(s).",
            )
            # Refresh internal cache after cleanup
            self._refresh_run_registry()
        except Exception as e:
            QtWidgets.QMessageBox.information(self, "Cleanup", str(e))

    def _refresh_run_registry(self):
        try:
            self._registry_cache = self._get_registry_table_data()
        except Exception:
            self._registry_cache = ([], [])

    def _show_registry_popup(self):
        # Guard against duplicate dialogs
        if getattr(self, "_dlg_open_registry", False):
            return
        self._dlg_open_registry = True
        try:
            cache = getattr(self, "_registry_cache", None)
            if not cache:
                cache = self._get_registry_table_data()
            headers, rows = cache
            dlg = QtWidgets.QDialog(self)
            dlg.setWindowTitle("Run Registry")
            dlg.resize(900, 500)
            dlg.setObjectName("runRegistryDialog")
            dlg.setStyleSheet("""
                #runRegistryDialog {
                    background-color: #f8fafc;
                    color: #0f172a;
                }
                #runRegistryDialog QLabel {
                    color: #0f172a;
                }
                #runRegistryDialog QTableWidget {
                    background-color: #ffffff;
                    border: 1px solid #d4d4d8;
                    gridline-color: #e4e4e7;
                    selection-background-color: #d1d5db;
                    selection-color: #0f172a;
                    color: #0f172a;
                }
                #runRegistryDialog QTableWidget::item {
                    color: #0f172a;
                }
                #runRegistryDialog QTableWidget::item:selected {
                    background-color: #d1d5db;
                    color: #0f172a;
                }
                #runRegistryDialog QHeaderView::section {
                    background-color: #eef2ff;
                    color: #0f172a;
                    padding: 8px;
                    border: 1px solid #cbd5f5;
                    font-weight: 600;
                }
                #runRegistryDialog QPushButton {
                    background-color: #ffffff;
                    color: #0f172a;
                    border: 1px solid #cbd5f5;
                    border-radius: 6px;
                    padding: 8px 14px;
                }
                #runRegistryDialog QPushButton:hover {
                    background-color: #f1f5f9;
                }
            """)
            v = QtWidgets.QVBoxLayout(dlg)
            tbl = QtWidgets.QTableWidget(0, len(headers))
            tbl.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.AllEditTriggers)
            tbl.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
            tbl.setAlternatingRowColors(False)
            if headers:
                tbl.setHorizontalHeaderLabels(headers)
            for r, row in enumerate(rows):
                tbl.insertRow(r)
                for c, val in enumerate(row[: len(headers) or len(row)]):
                    tbl.setItem(r, c, QtWidgets.QTableWidgetItem(val))
            tbl.resizeColumnsToContents()
            v.addWidget(tbl)
            # Controls: Delete Selected and Close
            bar = QtWidgets.QHBoxLayout()
            btn_save = QtWidgets.QPushButton("Save Changes")
            btn_delete = QtWidgets.QPushButton("Delete Selected")
            btn_close = QtWidgets.QPushButton("Close")
            bar.addStretch(1)
            bar.addWidget(btn_save)
            bar.addWidget(btn_delete)
            bar.addWidget(btn_close)
            v.addLayout(bar)

            def _delete_selected():
                if not headers:
                    return
                hdr_lc = [h.strip().lower() for h in headers]
                try:
                    idx_sc = hdr_lc.index("serial_component") if "serial_component" in hdr_lc else (
                        hdr_lc.index("serial") if "serial" in hdr_lc else -1
                    )
                except Exception:
                    idx_sc = -1
                if idx_sc < 0:
                    QtWidgets.QMessageBox.information(dlg, "Delete", "Cannot locate 'serial_component' column.")
                    return
                sels = tbl.selectionModel().selectedRows()
                if not sels:
                    QtWidgets.QMessageBox.information(dlg, "Delete", "Select one or more rows to delete.")
                    return
                serials: list[str] = []
                for mi in sels:
                    it = tbl.item(mi.row(), idx_sc)
                    if it and it.text().strip():
                        serials.append(it.text().strip())
                if not serials:
                    return
                confirm = QtWidgets.QMessageBox.question(
                    dlg,
                    "Confirm deletion",
                    f"Delete {len(serials)} entr(ies) and their run_data folders?",
                )
                if confirm != QtWidgets.QMessageBox.StandardButton.Yes:
                    return
                try:
                    be.delete_registry_entries(serials)
                except Exception as e:
                    QtWidgets.QMessageBox.information(dlg, "Delete", str(e))
                try:
                    self._start_worker(be.compile_master, status_msg="Compiling master workbook...")
                except Exception:
                    pass
                dlg.accept()

            def _save_changes():
                if not headers:
                    return
                rows_out: list[dict[str, str]] = []
                for r in range(tbl.rowCount()):
                    row_map: dict[str, str] = {}
                    for c in range(len(headers)):
                        val = tbl.item(r, c).text() if tbl.item(r, c) else ""
                        row_map[headers[c]] = val
                    rows_out.append(row_map)
                try:
                    be.write_run_registry_rows(rows_out)
                except Exception as e:
                    QtWidgets.QMessageBox.information(dlg, "Save", str(e))
                    return
                try:
                    self._start_worker(be.compile_master, status_msg="Compiling master workbook...")
                except Exception:
                    pass
                dlg.accept()

            btn_save.clicked.connect(_save_changes)
            btn_delete.clicked.connect(_delete_selected)
            btn_close.clicked.connect(dlg.reject)
            
            dlg.exec()
        except Exception as e:
            QtWidgets.QMessageBox.information(self, "Registry", str(e))
        finally:
            self._dlg_open_registry = False

    def _safe_open(self, fn):
        try:
            fn()
        except Exception as e:
            QtWidgets.QMessageBox.information(self, "Open", str(e))

    # File/browser helpers
    def _browse_file(self, edit: QtWidgets.QLineEdit, initial_dir: Path, filter_str: str):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select file", str(initial_dir), filter_str)
        if path:
            edit.setText(path)
        self._scan_refresh()

    def _browse_folder(self, edit: QtWidgets.QLineEdit, initial_dir: Path):
        path = QtWidgets.QFileDialog.getExistingDirectory(self, "Select folder", str(initial_dir))
        if path:
            edit.setText(path)
            # Persist repository root when editing the repo picker
            if edit is getattr(self, 'ed_repo', None):
                try:
                    be.set_repo_root(Path(path))
                except Exception:
                    pass
        self._scan_refresh()

    # Workspace sync
    def _sync_workspace(self, auto: bool = False):
        try:
            repo = Path(self.ed_repo.text()).expanduser() if hasattr(self, "ed_repo") else Path(self.ed_pdfs.text()).expanduser()
        except Exception:
            repo = be.DEFAULT_PDF_DIR
        try:
            terms = Path(self.ed_terms.text()).expanduser()
        except Exception:
            terms = be.DEFAULT_TERMS_XLSX
        try:
            # For manual syncs, first rebuild the registry from run_data
            if not auto:
                try:
                    be.rebuild_registry_from_run_data()
                except Exception:
                    pass
            summary, details = be.compute_workspace_sync(repo, terms)
            self._sync_summary = summary
            self._sync_details = details
            new = int(summary.get("new", 0) or 0)
            pdf_newer = int(summary.get("pdf_newer", 0) or 0)
            terms_newer = int(summary.get("terms_newer", 0) or 0)
            up_to_date = int(summary.get("up_to_date", 0) or 0)
            total = int(summary.get("total", 0) or 0)
            last = str(summary.get("last_sync", ""))
            repo_s = str(summary.get("repo_root", repo))
            txt = (
                f"Repository: {repo_s}\n"
                f"Total PDFs: {total}  |  New: {new}  |  Out-of-date (PDF): {pdf_newer}  |  Out-of-date (Terms): {terms_newer}  |  Up-to-date: {up_to_date}\n"
                f"Last sync: {last}"
            )
            self.lbl_sync_banner.setText(txt)
            if (new + pdf_newer + terms_newer) > 0:
                self.lbl_sync_banner.setStyleSheet(
                    "#syncBanner { background: #fff4e5; color: #5b3100; border: 1px solid #ffd9a8; border-radius: 6px; padding: 8px 12px; }"
                )
            else:
                self.lbl_sync_banner.setStyleSheet(
                    "#syncBanner { background: #e8f5e9; color: #1b5e20; border: 1px solid #c8e6c9; border-radius: 6px; padding: 8px 12px; }"
                )
            flagged = (new + pdf_newer + terms_newer)
            # Optionally compile master when sync is user-initiated
            if not auto:
                try:
                    self._start_worker(be.compile_master, status_msg="Compiling master workbook from registry...")
                except Exception:
                    pass
            if not auto:
                self._append_log("[GUI] Workspace sync complete")
                # Show toast notification for manual sync
                if flagged > 0:
                    self._show_toast(f"Workspace synced - {flagged} out-of-date items found")
                else:
                    self._show_toast("Workspace synced - all up-to-date")
            self._schedule_plot_terms_update("workspace sync")
        except Exception as e:
            self.lbl_sync_banner.setText(f"Sync failed: {e}")
            if not auto:
                self._show_toast(f"Workspace sync failed: {e}")

    def _act_sync_workspace(self):
        self._sync_workspace(auto=False)

    def _show_toast(self, message: str, duration: int = 5000):
        """Show a toast/cookie banner notification."""
        try:
            self._toast.show_message(message, duration)
        except Exception:
            pass

    def _show_outdated_popup(self, auto: bool = False):
        # Guard against duplicate dialogs
        if getattr(self, "_dlg_open_outdated", False):
            return
        self._dlg_open_outdated = True
        try:
            details = getattr(self, "_sync_details", None) or []
            rows = [d for d in details if d.get("reason") in ("new", "pdf_newer", "terms_newer")]
            dlg = QtWidgets.QDialog(self)
            dlg.setWindowTitle("View Data Package List and Update EIDAT Database")
            dlg.resize(1100, 600)
            dlg.setStyleSheet("""
                QDialog {
                    background: #ffffff;
                }
            """)

            v = QtWidgets.QVBoxLayout(dlg)
            v.setContentsMargins(20, 20, 20, 20)
            v.setSpacing(16)

            # Title and description
            title = QtWidgets.QLabel("Select documents to process")
            title.setStyleSheet("font-size: 16px; font-weight: 600; color: #111827;")
            v.addWidget(title)

            desc = QtWidgets.QLabel("Check the boxes to select which documents to extract and update")
            desc.setStyleSheet("font-size: 13px; color: #6b7280;")
            v.addWidget(desc)

            # Toolbar with Select All/None buttons
            toolbar = QtWidgets.QHBoxLayout()
            toolbar.setSpacing(8)

            btn_sel_all = QtWidgets.QPushButton("Select All")
            btn_sel_all.setStyleSheet("""
                QPushButton {
                    padding: 8px 16px;
                    border-radius: 6px;
                    background: #ffffff;
                    color: #374151;
                    border: 1px solid #d1d5db;
                    font-size: 13px;
                }
                QPushButton:hover {
                    background: #f9fafb;
                    border-color: #9ca3af;
                }
            """)

            btn_sel_none = QtWidgets.QPushButton("Select None")
            btn_sel_none.setStyleSheet("""
                QPushButton {
                    padding: 8px 16px;
                    border-radius: 6px;
                    background: #ffffff;
                    color: #374151;
                    border: 1px solid #d1d5db;
                    font-size: 13px;
                }
                QPushButton:hover {
                    background: #f9fafb;
                    border-color: #9ca3af;
                }
            """)

            toolbar.addWidget(btn_sel_all)
            toolbar.addWidget(btn_sel_none)
            toolbar.addStretch(1)
            v.addLayout(toolbar)

            # Table with checkboxes
            cols = ["Select", "Serial", "Reason", "PDF", "Run Date", "PDF Modified", "Terms Modified"]
            tbl = QtWidgets.QTableWidget(0, len(cols))
            tbl.setHorizontalHeaderLabels(cols)
            tbl.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
            tbl.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.ExtendedSelection)
            tbl.setAlternatingRowColors(True)
            tbl.verticalHeader().setVisible(False)

            # Consistent styling matching Terms Editor
            tbl.setStyleSheet("""
                QTableWidget {
                    background-color: #ffffff;
                    alternate-background-color: #f9fafb;
                    selection-background-color: #dbeafe;
                    selection-color: #111827;
                    gridline-color: #e5e7eb;
                    border: 1px solid #d1d5db;
                    border-radius: 6px;
                }
                QTableWidget::item {
                    padding: 10px 8px;
                    color: #374151;
                }
                QTableWidget::item:selected {
                    background-color: #dbeafe;
                    color: #111827;
                }
                QHeaderView::section {
                    background-color: #f3f4f6;
                    color: #111827;
                    padding: 12px 8px;
                    border: none;
                    border-right: 1px solid #e5e7eb;
                    border-bottom: 2px solid #d1d5db;
                    font-weight: 600;
                    font-size: 13px;
                }
                QCheckBox {
                    spacing: 0px;
                }
                QCheckBox::indicator {
                    width: 20px;
                    height: 20px;
                    border-radius: 4px;
                    border: 2px solid #d1d5db;
                    background: #ffffff;
                }
                QCheckBox::indicator:hover {
                    border-color: #2563eb;
                }
                QCheckBox::indicator:checked {
                    background: #2563eb;
                    border-color: #2563eb;
                    image: url(none);
                }
                QCheckBox::indicator:checked:after {
                    content: "✓";
                    color: #ffffff;
                }
            """)

            v.addWidget(tbl, 1)

            # Populate table with checkboxes
            for r, d in enumerate(rows):
                tbl.insertRow(r)

                # Create checkbox widget for Select column
                checkbox = QtWidgets.QCheckBox()
                checkbox.setChecked(True)
                checkbox.setStyleSheet("""
                    QCheckBox {
                        margin-left: 8px;
                    }
                    QCheckBox::indicator {
                        width: 18px;
                        height: 18px;
                        border-radius: 3px;
                        border: 2px solid #d1d5db;
                        background: #ffffff;
                    }
                    QCheckBox::indicator:hover {
                        border-color: #2563eb;
                    }
                    QCheckBox::indicator:checked {
                        background: #2563eb;
                        border-color: #2563eb;
                    }
                """)

                # Center the checkbox
                checkbox_widget = QtWidgets.QWidget()
                checkbox_layout = QtWidgets.QHBoxLayout(checkbox_widget)
                checkbox_layout.addWidget(checkbox)
                checkbox_layout.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
                checkbox_layout.setContentsMargins(0, 0, 0, 0)
                tbl.setCellWidget(r, 0, checkbox_widget)

                # Add data to other columns
                tbl.setItem(r, 1, QtWidgets.QTableWidgetItem(d.get("serial_component", "")))
                tbl.setItem(r, 2, QtWidgets.QTableWidgetItem(d.get("reason", "")))
                tbl.setItem(r, 3, QtWidgets.QTableWidgetItem(d.get("pdf", "")))
                tbl.setItem(r, 4, QtWidgets.QTableWidgetItem(d.get("run_date", "")))
                tbl.setItem(r, 5, QtWidgets.QTableWidgetItem(d.get("pdf_mtime", "")))
                tbl.setItem(r, 6, QtWidgets.QTableWidgetItem(d.get("terms_mtime", "")))

            tbl.resizeColumnsToContents()
            tbl.setColumnWidth(0, 80)  # Fixed width for checkbox column

            # Update Select All/None to work with checkbox widgets
            def _set_all_checkboxes(checked: bool):
                for r in range(tbl.rowCount()):
                    widget = tbl.cellWidget(r, 0)
                    if widget:
                        checkbox = widget.findChild(QtWidgets.QCheckBox)
                        if checkbox:
                            checkbox.setChecked(checked)

            btn_sel_all.clicked.connect(lambda: _set_all_checkboxes(True))
            btn_sel_none.clicked.connect(lambda: _set_all_checkboxes(False))

            # Bottom buttons with consistent styling
            btns = QtWidgets.QHBoxLayout()
            btns.setSpacing(12)

            btn_run_all = QtWidgets.QPushButton("Run All Out-of-Date")
            btn_run_all.setStyleSheet("""
                QPushButton {
                    padding: 10px 20px;
                    border-radius: 6px;
                    background: #2563eb;
                    color: #ffffff;
                    border: 1px solid #2563eb;
                    font-size: 13px;
                    font-weight: 600;
                }
                QPushButton:hover {
                    background: #1d4ed8;
                }
            """)

            btn_run = QtWidgets.QPushButton("Run Selected")
            btn_run.setStyleSheet("""
                QPushButton {
                    padding: 10px 20px;
                    border-radius: 6px;
                    background: #2563eb;
                    color: #ffffff;
                    border: 1px solid #2563eb;
                    font-size: 13px;
                    font-weight: 600;
                }
                QPushButton:hover {
                    background: #1d4ed8;
                }
            """)

            btn_close = QtWidgets.QPushButton("Close")
            btn_close.setStyleSheet("""
                QPushButton {
                    padding: 10px 20px;
                    border-radius: 6px;
                    background: #ffffff;
                    color: #6b7280;
                    border: 1px solid #d1d5db;
                    font-size: 13px;
                }
                QPushButton:hover {
                    background: #f9fafb;
                }
            """)

            btns.addStretch(1)
            btns.addWidget(btn_run_all)
            btns.addWidget(btn_run)
            btns.addWidget(btn_close)
            v.addLayout(btns)

            # Update run functions to work with checkbox widgets
            def _run_selected():
                paths: list[Path] = []
                for r in range(tbl.rowCount()):
                    widget = tbl.cellWidget(r, 0)
                    if widget:
                        checkbox = widget.findChild(QtWidgets.QCheckBox)
                        if checkbox and checkbox.isChecked():
                            p = tbl.item(r, 3).text() if tbl.item(r, 3) else ""
                            if p:
                                paths.append(Path(p))
                if not paths:
                    QtWidgets.QMessageBox.information(dlg, "Nothing selected", "Choose at least one EIDP to run.")
                    return
                try:
                    terms = Path(self.ed_terms.text()).expanduser()
                except Exception:
                    terms = be.DEFAULT_TERMS_XLSX
                self._enrich_after_run = True
                self._start_worker(lambda: be.run_selected_pdfs(paths, terms), status_msg="Running selected EIDPs...", show_run_progress=True, refresh_plot_terms_after=True)
                dlg.accept()

            def _run_all():
                all_paths = [Path(d.get("pdf")) for d in rows if d.get("pdf")] if rows else []
                if not all_paths:
                    QtWidgets.QMessageBox.information(dlg, "Nothing to run", "No out-of-date EIDPs found.")
                    return
                try:
                    terms = Path(self.ed_terms.text()).expanduser()
                except Exception:
                    terms = be.DEFAULT_TERMS_XLSX
                self._enrich_after_run = True
                self._start_worker(lambda: be.run_selected_pdfs(all_paths, terms), status_msg="Running all out-of-date EIDPs...", show_run_progress=True, refresh_plot_terms_after=True)
                dlg.accept()

            btn_run.clicked.connect(_run_selected)
            btn_run_all.clicked.connect(_run_all)
            btn_close.clicked.connect(dlg.reject)
            dlg.exec()
        finally:
            self._dlg_open_outdated = False

    def _scan_refresh(self):
        # Env status + badge
        py = be.resolve_project_python()
        env_ok = Path(py).exists()
        self.lbl_env.setText(f"Env: {'OK' if env_ok else 'Not Installed'} ({py})")
        self.lbl_ready.setText("System Ready" if env_ok else "Setup Needed")
        try:
            self.lbl_env_health.setProperty("status", "ok" if env_ok else "bad")
            self.lbl_env_health.setText("Healthy" if env_ok else "Issues")
            self.lbl_env_health.style().unpolish(self.lbl_env_health); self.lbl_env_health.style().polish(self.lbl_env_health)
        except Exception:
            pass

        # Sync settings from env
        try:
            env = be.parse_scanner_env()
            mode = env.get("OCR_MODE", "fallback").strip().lower()
            display = self._ocr_value_to_display.get(mode, self._ocr_value_to_display["fallback"]) if hasattr(self, "_ocr_value_to_display") else mode
            if self.cmb_ocr_mode.currentText() != display:
                # Ensure the display choice exists (in case of dynamic mapping)
                if display not in [self.cmb_ocr_mode.itemText(i) for i in range(self.cmb_ocr_mode.count())]:
                    self.cmb_ocr_mode.addItem(display)
                self.cmb_ocr_mode.setCurrentText(display)
            xy = env.get("XY_FUZZ", "0.50")
            try:
                val = float(xy)
                self.sld_xy_fuzz.blockSignals(True)
                self.sld_xy_fuzz.setValue(int(round(val * 100)))
                self.sld_xy_fuzz.blockSignals(False)
                self.lbl_xy_val.setText(f"{val:0.2f}")
            except Exception:
                pass
            # OCR DPI (500-1000)
            try:
                dpi = int(env.get("OCR_DPI", "500"))
                dpi = max(500, min(1000, dpi))
                self.sld_ocr_dpi.blockSignals(True)
                self.sld_ocr_dpi.setValue(dpi)
                self.sld_ocr_dpi.blockSignals(False)
                self.lbl_dpi_val.setText(str(dpi))
            except Exception:
                pass
            enable_logging = not (env.get("QUIET", "1").strip().lower() in ("1", "true", "yes"))
            self.chk_logging.blockSignals(True)
            self.chk_logging.setChecked(enable_logging)
            self.chk_logging.blockSignals(False)
            lang_code = env.get("EASYOCR_LANGS", env.get("OCR_LANGS", "en")).split(",")[0].strip() or "en"
            # Map code -> display name
            display_lang = (
                self._lang_code_to_display.get(lang_code, lang_code)
                if hasattr(self, "_lang_code_to_display")
                else lang_code
            )
            # Ensure present and set
            if display_lang not in [self.cmb_lang.itemText(i) for i in range(self.cmb_lang.count())]:
                self.cmb_lang.addItem(display_lang)
            self.cmb_lang.setCurrentText(display_lang)
        except Exception:
            pass

        # Enable/disable core actions
        terms_path = Path(self.ed_terms.text()).expanduser()
        has_terms = terms_path.exists()
        has_pdfs = Path(self.ed_pdfs.text()).expanduser().exists()
        worker = getattr(self, "_worker", None)
        self.btn_start.setEnabled(has_terms and has_pdfs and (worker is None or not worker.isRunning()))

        # Workspace sync banner refresh handled by periodic timer

        # Plotting selection
        self._refresh_series_catalog()
        # Inline viewer removed; popup will build data on demand

    # Upload ingestion
    def _ingest_paths(self, paths: list[str]) -> None:
        import shutil
        dest_dir = Path(self.ed_pdfs.text()).expanduser()
        try:
            dest_dir.mkdir(parents=True, exist_ok=True)
        except Exception as e:
            QtWidgets.QMessageBox.critical(self, "Folder error", f"Cannot create PDFs folder:\n{dest_dir}\n\n{e}")
            return
        copied, skipped, errors = 0, 0, 0
        to_visit: list[Path] = []
        for p in paths:
            try:
                to_visit.append(Path(p))
            except Exception:
                continue
        def is_pdf(p: Path) -> bool:
            return p.suffix.lower() == ".pdf"
        visit_files: list[Path] = []
        for p in to_visit:
            try:
                if p.is_dir():
                    for sub in p.rglob("*.pdf"):
                        visit_files.append(sub)
                elif p.is_file():
                    visit_files.append(p)
            except Exception:
                errors += 1
        for src in visit_files:
            try:
                if not is_pdf(src):
                    skipped += 1
                    continue
                dst = self._unique_destination(dest_dir / src.name)
                shutil.copy2(str(src), str(dst))
                copied += 1
                self._append_log(f"[GUI] Added {src} -> {dst}")
            except Exception as e:
                errors += 1
                self._append_log(f"[ERROR] Copy failed for {src}: {e}")
        self._refresh_upload_list()
        msg = f"Added {copied} file(s)" + (f", skipped {skipped}" if skipped else "") + (f", errors {errors}" if errors else "")
        self.status_bar.showMessage(msg, 4000)

    def _unique_destination(self, initial: Path) -> Path:
        if not initial.exists():
            return initial
        stem = initial.stem; suf = initial.suffix; parent = initial.parent; i = 1
        while True:
            cand = parent / f"{stem} ({i}){suf}"
            if not cand.exists():
                return cand
            i += 1

    def _refresh_upload_list(self):
        folder = Path(self.ed_pdfs.text()).expanduser()
        rows: list[tuple[str, str, str]] = []
        if folder.exists():
            try:
                for p in sorted(folder.glob("*.pdf"), key=lambda x: x.name.lower()):
                    size = self._fmt_size(p.stat().st_size)
                    mtime = self._fmt_mtime(p.stat().st_mtime)
                    rows.append((p.name, size, mtime))
            except Exception:
                pass
        self.list_pdfs.setRowCount(0)
        for r, (name, size, mtime) in enumerate(rows):
            self.list_pdfs.insertRow(r)
            self.list_pdfs.setItem(r, 0, QtWidgets.QTableWidgetItem(name))
            self.list_pdfs.setItem(r, 1, QtWidgets.QTableWidgetItem(size))
            self.list_pdfs.setItem(r, 2, QtWidgets.QTableWidgetItem(mtime))

    def _fmt_size(self, n: int) -> str:
        size = float(n)
        for unit in ["B", "KB", "MB", "GB"]:
            if size < 1024.0:
                return f"{size:.0f} {unit}" if unit == "B" else f"{size:.1f} {unit}"
            size /= 1024.0
        return f"{size:.1f} TB"

    def _fmt_mtime(self, ts: float) -> str:
        try:
            import datetime as _dt
            return _dt.datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M")
        except Exception:
            return ""

    def _selected_pdf_paths(self) -> list[Path]:
        folder = Path(self.ed_pdfs.text()).expanduser()
        sel = []
        for idx in self.list_pdfs.selectionModel().selectedRows():
            item = self.list_pdfs.item(idx.row(), 0)
            if item is None:
                continue
            sel.append(folder / item.text())
        return sel

    def _act_remove_selected(self):
        files = self._selected_pdf_paths()
        if not files:
            return
        names = "\n".join(p.name for p in files[:10])
        extra = "" if len(files) <= 10 else f"\nÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¾ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¾ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â¦ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¾ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â¦ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¦ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â¦ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¾ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¦ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¾Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â¦ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â‚¬Å¾Ã‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â¦ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã¢â‚¬Â ÃƒÂ¢Ã¢â€šÂ¬Ã¢â€žÂ¢ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡Ãƒâ€šÃ‚Â¬ÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â¡ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€¦Ã‚Â¡ÃƒÆ’Ã†â€™ÃƒÂ¢Ã¢â€šÂ¬Ã…Â¡ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Â¦and {len(files)-10} more"
        if (
            QtWidgets.QMessageBox.question(
                self, "Remove files", f"Delete these from PDFs folder?\n\n{names}{extra}"
            )
            != QtWidgets.QMessageBox.StandardButton.Yes
        ):
            return
        removed, errors = 0, 0
        for p in files:
            try:
                p.unlink(missing_ok=True)
                removed += 1
            except Exception:
                errors += 1
        self._append_log(f"[GUI] Removed {removed} file(s); errors {errors}")
        self._refresh_upload_list()

    def _act_remove_all(self):
        folder = Path(self.ed_pdfs.text()).expanduser()
        if not folder.exists():
            return
        if (
            QtWidgets.QMessageBox.question(
                self, "Remove all", f"Delete ALL PDFs in:\n{folder}?"
            )
            != QtWidgets.QMessageBox.StandardButton.Yes
        ):
            return
        removed, errors = 0, 0
        for p in folder.glob("*.pdf"):
            try:
                p.unlink(missing_ok=True)
                removed += 1
            except Exception:
                errors += 1
        self._append_log(f"[GUI] Cleared {removed} file(s); errors {errors}")
        self._refresh_upload_list()

    def _act_add_files(self):
        paths, _ = QtWidgets.QFileDialog.getOpenFileNames(self, "Select PDFs", str(Path(self.ed_pdfs.text()).expanduser()), "PDF files (*.pdf);;All files (*.*)")
        if paths:
            self._ingest_paths(paths)

    def _act_add_folder(self):
        path = QtWidgets.QFileDialog.getExistingDirectory(self, "Select folder with PDFs", str(Path(self.ed_pdfs.text()).expanduser()))
        if path:
            self._ingest_paths([path])

    # Plot selection helpers
    def _refresh_series_catalog(self):
        try:
            options = be.list_plot_series_options()
        except Exception:
            options = []
        self._plot_series_options = options
        if not options and hasattr(self, "status_bar"):
            self.status_bar.showMessage("No plot series found. Generate the term list first.", 4000)

    def _load_proposed_plots(self):
        try:
            records = be.read_proposed_plots()
        except Exception:
            records = []
        self._proposed_plots_cache = records
        self._refresh_proposed_summary()

    def _refresh_proposed_summary(self):
        if not hasattr(self, "lbl_proposed_summary"):
            return
        if not self._proposed_plots_cache:
            self.lbl_proposed_summary.setText("No plots configured. Click 'Manage Proposed Plots…' to add some.")
            return
        lines: list[str] = []
        for entry in self._proposed_plots_cache:
            name = entry.get("name") or "Plot"
            series = entry.get("series") or []
            summary = ", ".join(series) if series else "No series selected"
            lines.append(f"{name}: {summary}")
        self.lbl_proposed_summary.setText("\n".join(lines))

    def _open_proposed_plots_dialog(self):
        dlg = ProposedPlotsDialog(self._plot_series_options, list(self._proposed_plots_cache), parent=self)
        if dlg.exec() == QtWidgets.QDialog.DialogCode.Accepted:
            new_plots = dlg.plots()
            try:
                be.write_proposed_plots(new_plots)
            except Exception as exc:
                QtWidgets.QMessageBox.warning(self, "Save failed", str(exc))
                return
            self._proposed_plots_cache = new_plots
            self._refresh_proposed_summary()

    def _apply_tab_widths(self) -> None:
        if not hasattr(self, "_tab_style_template"):
            return
        try:
            bar = self.tabs.tabBar()
        except Exception:
            bar = None
        if not bar:
            return
        count = max(1, bar.count())
        width = max(80, self.tabs.width() // count)
        css = self._tab_style_template.replace("{width}", str(width))
        self.tabs.setStyleSheet(css)

    def resizeEvent(self, event: QtGui.QResizeEvent) -> None:  # type: ignore[override]
        super().resizeEvent(event)

    def _build_logo_pixmap(self, size: int = 52) -> QtGui.QPixmap:
        """Load external app logo if present; otherwise draw the fallback glyph.

        Checks these paths in order and scales preserving aspect ratio:
        - ui_next/assets/app_logo.png
        - ui_next/assets/logo.png
        - user_inputs/app_logo.png
        """
        candidates = [
            (be.ROOT / "ui_next" / "assets" / "app_logo.png"),
            (be.ROOT / "ui_next" / "assets" / "logo.png"),
            (be.ROOT / "user_inputs" / "app_logo.png"),
        ]
        for path in candidates:
            try:
                if path.exists():
                    pix = QtGui.QPixmap(str(path))
                    if not pix.isNull():
                        scaled = pix.scaled(size, size, QtCore.Qt.AspectRatioMode.KeepAspectRatio, QtCore.Qt.TransformationMode.SmoothTransformation)
                        try:
                            # Also set the window icon for consistency
                            self.setWindowIcon(QtGui.QIcon(scaled))
                        except Exception:
                            pass
                        return scaled
            except Exception:
                pass
        # Fallback: draw an inline polished monogram
        pix = QtGui.QPixmap(size, size)
        pix.fill(QtCore.Qt.GlobalColor.transparent)
        painter = QtGui.QPainter(pix)
        painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing)

        rect = QtCore.QRectF(0.5, 0.5, size - 1, size - 1)
        radius = size * 0.24

        # Base gradient block with subtle border
        gradient = QtGui.QLinearGradient(0, 0, size, size)
        gradient.setColorAt(0, QtGui.QColor("#60a5fa"))
        gradient.setColorAt(1, QtGui.QColor("#1d4ed8"))
        painter.setBrush(QtGui.QBrush(gradient))
        painter.setPen(QtGui.QPen(QtGui.QColor("#0f172a"), max(2, size // 18)))
        painter.drawRoundedRect(rect, radius, radius)

        # Inner glow
        painter.setPen(QtCore.Qt.PenStyle.NoPen)
        painter.setBrush(QtGui.QColor(255, 255, 255, 35))
        painter.drawRoundedRect(rect.adjusted(size * 0.08, size * 0.08, -size * 0.08, -size * 0.25), radius * 0.8, radius * 0.8)

        # Accent ring
        inner = QtCore.QRectF(size * 0.22, size * 0.22, size * 0.56, size * 0.56)
        painter.setBrush(QtCore.Qt.BrushStyle.NoBrush)
        painter.setPen(QtGui.QPen(QtGui.QColor(255, 255, 255, 120), max(2, size // 25)))
        painter.drawEllipse(inner)

        # Stylized "E" glyph
        glyph_pen = QtGui.QPen(QtGui.QColor("#f8fafc"))
        glyph_pen.setWidth(max(3, size // 16))
        glyph_pen.setCapStyle(QtCore.Qt.PenCapStyle.RoundCap)
        painter.setPen(glyph_pen)
        mid_y = size * 0.5
        left_x = size * 0.28
        right_x = size * 0.72
        painter.drawLine(QtCore.QPointF(left_x, size * 0.26), QtCore.QPointF(left_x, size * 0.74))
        painter.drawLine(QtCore.QPointF(left_x, size * 0.3), QtCore.QPointF(right_x, size * 0.3))
        painter.drawLine(QtCore.QPointF(left_x, mid_y), QtCore.QPointF(size * 0.65, mid_y))
        painter.drawLine(QtCore.QPointF(left_x, size * 0.7), QtCore.QPointF(right_x, size * 0.7))

        # Small spark in corner
        painter.setPen(QtGui.QPen(QtGui.QColor(255, 255, 255, 180), max(2, size // 28)))
        painter.drawPoint(QtCore.QPointF(size * 0.78, size * 0.28))

        painter.end()
        try:
            self.setWindowIcon(QtGui.QIcon(pix))
        except Exception:
            pass
        return pix

    # Settings persistence
    def _on_xy_slider(self, value: int):
        try:
            self.lbl_xy_val.setText(f"{value/100.0:0.2f}")
        except Exception:
            pass
        self._persist_settings_from_panel()

    def _on_dpi_slider(self, value: int):
        try:
            self.lbl_dpi_val.setText(str(value))
        except Exception:
            pass
        self._persist_settings_from_panel()

    def _persist_settings_from_panel(self):
        try:
            env = be.parse_scanner_env()
            disp = self.cmb_ocr_mode.currentText()
            env["OCR_MODE"] = self._ocr_display_to_value.get(disp, "fallback") if hasattr(self, "_ocr_display_to_value") else disp.strip().lower()
            env["XY_FUZZ"] = f"{self.sld_xy_fuzz.value()/100.0:0.2f}"
            env["OCR_DPI"] = str(int(self.sld_ocr_dpi.value()))
            # QUIET is inverse of logging toggle
            env["QUIET"] = "0" if self.chk_logging.isChecked() else "1"
            # Map display name -> language code for env
            disp_lang = self.cmb_lang.currentText().strip()
            env["EASYOCR_LANGS"] = (
                self._lang_display_to_code.get(disp_lang, disp_lang)
                if hasattr(self, "_lang_display_to_code")
                else disp_lang
            )
            be.save_scanner_env(env)
            self.status_bar.showMessage("Settings saved", 2000)
        except Exception:
            pass

    def _show_table_extraction_dialog(self):
        """Show dialog for table extraction with PDF selection and parameters."""
        # Guard against duplicate dialogs
        if getattr(self, "_dlg_table_extraction", False):
            return
        self._dlg_table_extraction = True

        try:
            # Get available PDFs from workspace sync
            details = getattr(self, "_sync_details", None) or []
            all_pdfs = [(d.get("pdf"), d.get("serial_component", ""), d.get("reason", "")) for d in details if d.get("pdf")]

            dlg = QtWidgets.QDialog(self)
            dlg.setWindowTitle("Extract Tables from Data Packages")
            dlg.resize(1200, 700)
            dlg.setStyleSheet("""
                QDialog {
                    background: #ffffff;
                }
                QGroupBox {
                    font-weight: 600;
                    font-size: 13px;
                    border: 1px solid #e5e7eb;
                    border-radius: 6px;
                    margin-top: 12px;
                    padding: 12px;
                }
                QGroupBox::title {
                    subcontrol-origin: margin;
                    subcontrol-position: top left;
                    padding: 0 4px;
                }
            """)

            main_layout = QtWidgets.QVBoxLayout(dlg)
            main_layout.setContentsMargins(20, 20, 20, 20)
            main_layout.setSpacing(16)

            # Title
            title = QtWidgets.QLabel("Extract Tables from Data Packages")
            title.setStyleSheet("font-size: 18px; font-weight: 700; color: #111827;")
            main_layout.addWidget(title)

            desc = QtWidgets.QLabel("Select documents and configure table extraction parameters")
            desc.setStyleSheet("font-size: 13px; color: #6b7280;")
            main_layout.addWidget(desc)

            # Main content in horizontal split
            content_layout = QtWidgets.QHBoxLayout()
            content_layout.setSpacing(16)

            # ========== LEFT: PDF Selection ==========
            left_widget = QtWidgets.QWidget()
            left_layout = QtWidgets.QVBoxLayout(left_widget)
            left_layout.setContentsMargins(0, 0, 0, 0)
            left_layout.setSpacing(8)

            pdf_label = QtWidgets.QLabel("Select Data Packages")
            pdf_label.setStyleSheet("font-size: 14px; font-weight: 600; color: #111827;")
            left_layout.addWidget(pdf_label)

            # Toolbar
            toolbar = QtWidgets.QHBoxLayout()
            toolbar.setSpacing(8)

            btn_sel_all = QtWidgets.QPushButton("Select All")
            btn_sel_all.setStyleSheet("""
                QPushButton {
                    padding: 6px 12px;
                    border-radius: 4px;
                    background: #ffffff;
                    color: #374151;
                    border: 1px solid #d1d5db;
                    font-size: 12px;
                }
                QPushButton:hover { background: #f9fafb; }
            """)

            btn_sel_none = QtWidgets.QPushButton("Select None")
            btn_sel_none.setStyleSheet("""
                QPushButton {
                    padding: 6px 12px;
                    border-radius: 4px;
                    background: #ffffff;
                    color: #374151;
                    border: 1px solid #d1d5db;
                    font-size: 12px;
                }
                QPushButton:hover { background: #f9fafb; }
            """)

            toolbar.addWidget(btn_sel_all)
            toolbar.addWidget(btn_sel_none)
            toolbar.addStretch()
            left_layout.addLayout(toolbar)

            # PDF list with checkboxes
            pdf_scroll = QtWidgets.QScrollArea()
            pdf_scroll.setWidgetResizable(True)
            pdf_scroll.setStyleSheet("""
                QScrollArea {
                    border: 1px solid #e5e7eb;
                    border-radius: 6px;
                    background: #f9fafb;
                }
            """)

            pdf_container = QtWidgets.QWidget()
            pdf_layout = QtWidgets.QVBoxLayout(pdf_container)
            pdf_layout.setContentsMargins(8, 8, 8, 8)
            pdf_layout.setSpacing(4)

            checkboxes = []
            if not all_pdfs:
                no_pdfs_label = QtWidgets.QLabel("No PDFs found. Run 'Sync Workspace Now' first.")
                no_pdfs_label.setStyleSheet("color: #6b7280; font-style: italic; padding: 16px;")
                pdf_layout.addWidget(no_pdfs_label)
            else:
                for pdf_path, serial, reason in all_pdfs:
                    checkbox = QtWidgets.QCheckBox(f"{serial or Path(pdf_path).stem}")
                    checkbox.setToolTip(str(pdf_path))
                    checkbox.setProperty("pdf_path", pdf_path)
                    checkbox.setStyleSheet("""
                        QCheckBox {
                            padding: 4px;
                            color: #374151;
                        }
                        QCheckBox::indicator {
                            width: 16px;
                            height: 16px;
                            border-radius: 3px;
                            border: 2px solid #d1d5db;
                            background: #ffffff;
                        }
                        QCheckBox::indicator:hover { border-color: #10b981; }
                        QCheckBox::indicator:checked {
                            background: #10b981;
                            border-color: #10b981;
                        }
                    """)
                    pdf_layout.addWidget(checkbox)
                    checkboxes.append(checkbox)

            pdf_layout.addStretch()
            pdf_scroll.setWidget(pdf_container)
            left_layout.addWidget(pdf_scroll, 1)

            # Connect toolbar buttons
            btn_sel_all.clicked.connect(lambda: [cb.setChecked(True) for cb in checkboxes])
            btn_sel_none.clicked.connect(lambda: [cb.setChecked(False) for cb in checkboxes])

            content_layout.addWidget(left_widget, 1)

            # ========== RIGHT: Parameters ==========
            right_widget = QtWidgets.QWidget()
            right_layout = QtWidgets.QVBoxLayout(right_widget)
            right_layout.setContentsMargins(0, 0, 0, 0)
            right_layout.setSpacing(12)

            param_label = QtWidgets.QLabel("Extraction Parameters")
            param_label.setStyleSheet("font-size: 14px; font-weight: 600; color: #111827;")
            right_layout.addWidget(param_label)

            # Parameters form
            form_scroll = QtWidgets.QScrollArea()
            form_scroll.setWidgetResizable(True)
            form_scroll.setStyleSheet("""
                QScrollArea {
                    border: 1px solid #e5e7eb;
                    border-radius: 6px;
                    background: #ffffff;
                }
            """)

            # Helper function to create styled labels
            def make_label(text: str) -> QtWidgets.QLabel:
                lbl = QtWidgets.QLabel(text)
                lbl.setStyleSheet("color: #1f2937; font-size: 13px; font-weight: 600;")
                return lbl

            # Helper for input styling
            input_style = """
                QLineEdit, QSpinBox, QDoubleSpinBox {
                    padding: 6px 8px;
                    border: 1px solid #d1d5db;
                    border-radius: 4px;
                    font-size: 13px;
                    background: #ffffff;
                }
                QLineEdit:focus, QSpinBox:focus, QDoubleSpinBox:focus {
                    border-color: #10b981;
                }
            """

            form_container = QtWidgets.QWidget()
            form_layout = QtWidgets.QFormLayout(form_container)
            form_layout.setContentsMargins(16, 16, 16, 16)
            form_layout.setSpacing(16)
            form_layout.setLabelAlignment(QtCore.Qt.AlignmentFlag.AlignRight)

            # Pages (required)
            pages_input = QtWidgets.QLineEdit("1")
            pages_input.setPlaceholderText("e.g., 1,3-5,7")
            pages_input.setToolTip("Page range (1-indexed). Examples: 1 or 1,3-5 or 1-10")
            pages_input.setStyleSheet(input_style)
            form_layout.addRow(make_label("Pages* (required):"), pages_input)

            # Fixed column count
            num_cols_input = QtWidgets.QLineEdit("0")
            num_cols_input.setPlaceholderText("0 = Auto-detect")
            num_cols_input.setToolTip("Set exact number of columns (0 = auto-detect, range: 0-20)")
            num_cols_input.setStyleSheet(input_style)
            num_cols_input.setMaxLength(2)
            form_layout.addRow(make_label("Fixed Columns:"), num_cols_input)

            # Min columns
            min_cols_input = QtWidgets.QLineEdit("2")
            min_cols_input.setPlaceholderText("Default: 2")
            min_cols_input.setToolTip("Minimum columns required (auto-detect mode only, range: 1-20)")
            min_cols_input.setStyleSheet(input_style)
            min_cols_input.setMaxLength(2)
            form_layout.addRow(make_label("Min Columns:"), min_cols_input)

            # Min rows
            min_rows_input = QtWidgets.QLineEdit("3")
            min_rows_input.setPlaceholderText("Default: 3")
            min_rows_input.setToolTip("Minimum consecutive rows to form a table (range: 1-50)")
            min_rows_input.setStyleSheet(input_style)
            min_rows_input.setMaxLength(2)
            form_layout.addRow(make_label("Min Rows:"), min_rows_input)

            # Match threshold
            threshold_input = QtWidgets.QLineEdit("0.5")
            threshold_input.setPlaceholderText("Default: 0.5")
            threshold_input.setToolTip("Type matching threshold for column alignment (range: 0.0-1.0)")
            threshold_input.setStyleSheet(input_style)
            threshold_input.setMaxLength(4)
            form_layout.addRow(make_label("Match Threshold:"), threshold_input)

            # OCR mode
            ocr_checkbox = QtWidgets.QCheckBox("Force OCR mode")
            ocr_checkbox.setToolTip("Use EasyOCR instead of PyMuPDF text extraction")
            ocr_checkbox.setStyleSheet("font-size: 13px; color: #374151;")
            form_layout.addRow(make_label(""), ocr_checkbox)

            # DPI
            dpi_input = QtWidgets.QLineEdit("300")
            dpi_input.setPlaceholderText("Default: 300")
            dpi_input.setToolTip("DPI for OCR rendering if OCR mode enabled (range: 150-800)")
            dpi_input.setStyleSheet(input_style)
            dpi_input.setMaxLength(3)
            form_layout.addRow(make_label("OCR DPI:"), dpi_input)

            # Delimiter
            delimiter_input = QtWidgets.QLineEdit()
            delimiter_input.setPlaceholderText("Auto-detect")
            delimiter_input.setToolTip("Field delimiter (leave empty for auto-detect)")
            delimiter_input.setStyleSheet(input_style)
            form_layout.addRow(make_label("Delimiter:"), delimiter_input)

            form_scroll.setWidget(form_container)
            right_layout.addWidget(form_scroll, 1)

            # Help text
            help_text = QtWidgets.QLabel(
                "<b>Tips:</b><br>"
                "• Fixed Columns: Set to exact column count for smart alignment<br>"
                "• Match Threshold: Higher = stricter type matching<br>"
                "• Output saved to Data Packages/<filename>_table.xlsx"
            )
            help_text.setStyleSheet("color: #6b7280; font-size: 11px; padding: 8px; background: #f9fafb; border-radius: 4px;")
            help_text.setWordWrap(True)
            right_layout.addWidget(help_text)

            content_layout.addWidget(right_widget, 1)
            main_layout.addLayout(content_layout, 1)

            # ========== BOTTOM: Action buttons ==========
            button_layout = QtWidgets.QHBoxLayout()
            button_layout.setSpacing(8)

            btn_extract = QtWidgets.QPushButton("\U0001F4CA Extract Tables")
            btn_extract.setStyleSheet("""
                QPushButton {
                    padding: 10px 20px;
                    border-radius: 6px;
                    background: #10b981;
                    color: #ffffff;
                    border: none;
                    font-size: 13px;
                    font-weight: 600;
                }
                QPushButton:hover { background: #059669; }
                QPushButton:disabled {
                    background: #d1d5db;
                    color: #9ca3af;
                }
            """)

            btn_cancel = QtWidgets.QPushButton("Cancel")
            btn_cancel.setStyleSheet("""
                QPushButton {
                    padding: 10px 20px;
                    border-radius: 6px;
                    background: #ffffff;
                    color: #6b7280;
                    border: 1px solid #d1d5db;
                    font-size: 13px;
                }
                QPushButton:hover { background: #f9fafb; }
            """)

            button_layout.addStretch()
            button_layout.addWidget(btn_cancel)
            button_layout.addWidget(btn_extract)
            main_layout.addLayout(button_layout)

            # ========== Connect actions ==========
            def extract_tables():
                selected = [cb for cb in checkboxes if cb.isChecked()]
                if not selected:
                    QtWidgets.QMessageBox.warning(dlg, "No Selection", "Please select at least one PDF")
                    return

                pages = pages_input.text().strip()
                if not pages:
                    QtWidgets.QMessageBox.warning(dlg, "Missing Pages", "Please specify page range (e.g., 1 or 1,3-5)")
                    return

                # Extract and validate parameters
                try:
                    num_cols_val = int(num_cols_input.text().strip() or "0")
                    if not (0 <= num_cols_val <= 20):
                        raise ValueError("Fixed Columns must be between 0 and 20")
                    num_cols = num_cols_val if num_cols_val > 0 else None

                    min_cols = int(min_cols_input.text().strip() or "2")
                    if not (1 <= min_cols <= 20):
                        raise ValueError("Min Columns must be between 1 and 20")

                    min_rows = int(min_rows_input.text().strip() or "3")
                    if not (1 <= min_rows <= 50):
                        raise ValueError("Min Rows must be between 1 and 50")

                    match_threshold = float(threshold_input.text().strip() or "0.5")
                    if not (0.0 <= match_threshold <= 1.0):
                        raise ValueError("Match Threshold must be between 0.0 and 1.0")

                    dpi = int(dpi_input.text().strip() or "300")
                    if not (150 <= dpi <= 800):
                        raise ValueError("OCR DPI must be between 150 and 800")

                except ValueError as e:
                    QtWidgets.QMessageBox.warning(dlg, "Invalid Input", str(e))
                    return

                ocr = ocr_checkbox.isChecked()
                delimiter = delimiter_input.text().strip() or None

                dlg.accept()

                # Process each selected PDF
                for cb in selected:
                    pdf_path = Path(cb.property("pdf_path"))
                    try:
                        self._start_worker(
                            lambda p=pdf_path: be.extract_csv_tables(
                                p, pages, num_cols, min_cols, min_rows,
                                match_threshold, ocr, dpi, delimiter
                            ),
                            status_msg=f"Extracting tables from {pdf_path.name}..."
                        )
                    except Exception as e:
                        QtWidgets.QMessageBox.critical(self, "Extraction Error", f"Failed to extract from {pdf_path.name}:\n{e}")

                self._show_toast(f"Table extraction started for {len(selected)} document(s)")

            btn_extract.clicked.connect(extract_tables)
            btn_cancel.clicked.connect(dlg.reject)

            dlg.exec()

        finally:
            self._dlg_table_extraction = False


def main():
    app = QtWidgets.QApplication(sys.argv)
    w = MainWindow(); w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
