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

        self.spinner_label = QtWidgets.QLabel("◐")
        spin_font = self.spinner_label.font()
        spin_font.setPointSize(32)
        spin_font.setBold(True)
        self.spinner_label.setFont(spin_font)
        self.spinner_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)

        self.progress_bar = QtWidgets.QProgressBar()
        self.progress_bar.setRange(0, 0)  # indeterminate until totals stream in
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("Working...")

        self.detail_label = QtWidgets.QLabel("Terms found: 0 / 0 \u2022 0 remaining")
        self.detail_label.setAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        self.detail_label.setStyleSheet("font-size: 12px;")

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

        self._spinner_frames = ["◐", "◓", "◑", "◒"]
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
        self.detail_label.setText("Terms found: 0 / 0 \u2022 0 remaining")
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

    def update_progress(self, completed: int, total: int):
        if total <= 0:
            if self.progress_bar.maximum() != 0:
                self.progress_bar.setRange(0, 0)
                self.progress_bar.setFormat("Working...")
            self.detail_label.setText(f"Terms processed: {completed}")
            return
        if self.progress_bar.maximum() == 0:
            self.progress_bar.setRange(0, 100)
        pct = max(0, min(100, int(round((completed * 100) / max(1, total)))))
        remaining = max(0, total - completed)
        self.progress_bar.setValue(pct)
        self.progress_bar.setFormat(f"{pct}%")
        self.detail_label.setText(f"Terms found: {completed} / {total} \u2022 {remaining} remaining")

    def finish(self, message: str, success: bool = True):
        self._anim_timer.stop()
        self.spinner_label.setText("✓" if success else "!")
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


class TermsEditorDialog(QtWidgets.QDialog):
    """In-app editor for Smart-Snap terms with grouped headers and curated fields."""

    VISIBLE_COLUMN_ORDER = [
        "Data Group",
        "Term Label",
        "Mode",
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
        "Term": "Search Term",
        "Data Group": "Data Grouping",
        "Term Label": "Term Label",
        "Mode": "Mode",
        "Smart Snap Type": "Smart Snap Type",
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
    GROUP_LAYOUT = [
        ("Data Information", ["Data Group", "Term Label"]),
        ("Mode / Type", ["Mode", "Smart Snap Type"]),
        ("Search Index", ["Term", "Pages", "GroupAfter", "GroupBefore"]),
        (
            "Data Definition for Extraction",
            ["Units", "Range (min)", "Range (max)", "Format", "Secondary Term", "Smart Position"],
        ),
    ]
    DEFAULT_HIDDEN = {"Return"}

    def __init__(self, terms_path: Path, parent=None):
        super().__init__(parent)
        self._terms_path = Path(terms_path)
        self._all_headers: list[str] = []
        self._visible_headers: list[str] = []
        self._hidden_headers: list[str] = []
        self._hidden_rows: list[dict[str, str]] = []
        self._dirty = False
        self._loading = False

        self.setWindowTitle("Smart-Snap Terms Editor")
        self.resize(1280, 680)

        self._combo_defs = {
            "Mode": {
                "options": [(opt, opt) for opt in be.TERMS_MODE_CHOICES] or [("smart", "smart")],
                "default": (be.TERMS_MODE_CHOICES[0] if be.TERMS_MODE_CHOICES else "smart"),
            },
            "Smart Snap Type": {
                "options": self._smart_type_options(),
                "default": "",
            },
        }

        layout = QtWidgets.QVBoxLayout(self)
        title = QtWidgets.QLabel("Edit the Smart-Snap input spreadsheet directly in the app.")
        title.setStyleSheet("font-size: 16px; font-weight: 600;")
        layout.addWidget(title)
        hint = QtWidgets.QLabel(f"File: {self._terms_path}")
        hint.setObjectName("termsPathLabel")
        hint.setStyleSheet("color: #556070;")
        hint.setWordWrap(True)
        layout.addWidget(hint)

        self.group_header = QtWidgets.QTableWidget(1, 0)
        self._configure_group_header_widget()
        layout.addWidget(self.group_header)

        self.table = QtWidgets.QTableWidget()
        self.table.setAlternatingRowColors(True)
        self.table.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.ExtendedSelection)
        self.table.verticalHeader().setVisible(False)
        header = self.table.horizontalHeader()
        header.setStretchLastSection(False)
        header.sectionResized.connect(self._sync_group_header_section)
        self.table.horizontalScrollBar().valueChanged.connect(self.group_header.horizontalScrollBar().setValue)
        layout.addWidget(self.table, 1)

        row_btns = QtWidgets.QHBoxLayout()
        self.btn_add_row = QtWidgets.QPushButton("Add Row")
        self.btn_duplicate_row = QtWidgets.QPushButton("Duplicate Row")
        self.btn_delete_row = QtWidgets.QPushButton("Delete Selected")
        self.btn_add_row.clicked.connect(self._add_blank_row)
        self.btn_duplicate_row.clicked.connect(self._duplicate_row)
        self.btn_delete_row.clicked.connect(self._delete_rows)
        row_btns.addWidget(self.btn_add_row)
        row_btns.addWidget(self.btn_duplicate_row)
        row_btns.addWidget(self.btn_delete_row)
        row_btns.addStretch(1)
        layout.addLayout(row_btns)

        bottom = QtWidgets.QHBoxLayout()
        self._status_label = QtWidgets.QLabel("Loading...")
        self._status_label.setObjectName("termsStatusLabel")
        self._status_label.setStyleSheet("color: #1b5e20;")
        bottom.addWidget(self._status_label)
        bottom.addStretch(1)
        self.btn_save = QtWidgets.QPushButton("Save")
        self.btn_save.setProperty("variant", "primary")
        self.btn_done = QtWidgets.QPushButton("Done")
        self.btn_save.clicked.connect(self._save_rows)
        self.btn_done.clicked.connect(self.reject)
        bottom.addWidget(self.btn_save)
        bottom.addWidget(self.btn_done)
        layout.addLayout(bottom)

        QtGui.QShortcut(QtGui.QKeySequence.StandardKey.Save, self, activated=self._save_rows)
        self.table.itemChanged.connect(self._on_table_item_changed)

        self._load_rows()

    def _configure_group_header_widget(self) -> None:
        self.group_header.setEditTriggers(QtWidgets.QAbstractItemView.EditTriggers.NoEditTriggers)
        self.group_header.setFocusPolicy(QtCore.Qt.FocusPolicy.NoFocus)
        self.group_header.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.NoSelection)
        self.group_header.setFixedHeight(46)
        self.group_header.setShowGrid(False)
        self.group_header.horizontalHeader().setVisible(False)
        self.group_header.verticalHeader().setVisible(False)
        self.group_header.setHorizontalScrollMode(QtWidgets.QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.group_header.horizontalHeader().setSectionResizeMode(QtWidgets.QHeaderView.ResizeMode.Fixed)
        self.group_header.setStyleSheet(
            """
            QTableWidget {
                background: #050505;
                border: none;
                border-bottom: 4px solid #121212;
            }
            QTableWidget::item {
                border-right: 1px solid #1f1f1f;
                padding-top: 6px;
                padding-bottom: 6px;
            }
            """
        )
        self.group_header.setVerticalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.group_header.setHorizontalScrollBarPolicy(QtCore.Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

    def _make_group_cell(self, text: str = "") -> QtWidgets.QTableWidgetItem:
        item = QtWidgets.QTableWidgetItem(text)
        item.setTextAlignment(QtCore.Qt.AlignmentFlag.AlignCenter)
        item.setFlags(QtCore.Qt.ItemFlag.NoItemFlags)
        font = item.font()
        font.setBold(bool(text))
        item.setFont(font)
        item.setForeground(QtGui.QBrush(QtGui.QColor("#f4f4f4")))
        item.setBackground(QtGui.QBrush(QtGui.QColor("#050505")))
        return item

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
        self._rebuild_group_header()

    def _column_display_name(self, header: str) -> str:
        return self.COLUMN_DISPLAY_NAMES.get(header, header)

    def _rebuild_group_header(self) -> None:
        self.group_header.blockSignals(True)
        self.group_header.clear()
        cols = len(self._visible_headers)
        self.group_header.setColumnCount(cols)
        self.group_header.setRowCount(1)
        self.group_header.clearSpans()
        for idx in range(cols):
            self.group_header.setColumnWidth(idx, self.table.columnWidth(idx))
            self.group_header.setItem(0, idx, self._make_group_cell(""))
        for group_name, members in self.GROUP_LAYOUT:
            indices = [self._visible_headers.index(m) for m in members if m in self._visible_headers]
            if not indices:
                continue
            start = min(indices)
            span = len(indices)
            self.group_header.setSpan(0, start, 1, span)
            self.group_header.setItem(0, start, self._make_group_cell(group_name))
        self.group_header.blockSignals(False)

    def _sync_group_header_section(self, logical_index: int, _old_size: int, new_size: int) -> None:
        try:
            self.group_header.setColumnWidth(logical_index, new_size)
        except Exception:
            pass

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
        return True

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
        self.setWindowTitle("EIDAT - End Item Data Analysis Tool")
        self.resize(1280, 860)

        be.ensure_scaffold()
        self._refresh_plot_series_after_worker = False

        # Global styling to reflect the provided mock
        self.setStyleSheet(
            """
            QMainWindow { background: #f7f9fc; }
            QLabel { color: #1f2937; }
            QLabel.subtle { color: #5b6b7a; font-size: 12px; }
            QGroupBox { font-weight: 700; font-size: 15px; border: 1px solid #e1e6ef; border-radius: 8px; margin-top: 26px; background: #ffffff; }
            /* Make group titles pronounced, black, centered, and add padding below */
            QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top center; color: #000000; background: #ffffff; padding: 2px 8px 6px 8px; }
            QPushButton { padding: 10px 16px; border-radius: 6px; background: #ffffff; color: #17324d; border: 1px solid #d1dae6; }
            QPushButton:hover { background: #f1f4f9; }
            QPushButton[variant="primary"] { background: #113a70; color: #ffffff; border: 1px solid #0f335f; }
            QPushButton[variant="primary"]:disabled { background: #b9c6d9; color: #f7f9fc; }
            QPushButton[variant="ghost"] { background: #ffffff; color: #17324d; border: 1px solid #d1dae6; }

            QLabel#statusBadge { background: #e8f5e9; color: #1b5e20; border-radius: 10px; padding: 4px 10px; }
            QLabel#healthBadge { background: #e8f5e9; color: #1b5e20; border-radius: 10px; padding: 4px 10px; }
            QLabel#healthBadge[status="bad"] { background: #fdecea; color: #b00020; }

            QLineEdit, QComboBox { background: #ffffff; border: 1px solid #d1dae6; border-radius: 6px; padding: 8px; color: #17324d; }
            /* Ensure popup list text is visible */
            QComboBox QAbstractItemView { color: #17324d; background: #ffffff; selection-background-color: #e7edf6; }
            QComboBox::drop-down { width: 28px; }

            QSlider::groove:horizontal { height: 8px; background: #e7edf6; border-radius: 4px; }
            QSlider::handle:horizontal { width: 18px; height: 18px; margin: -6px 0; border-radius: 9px; background: #113a70; }

            QCheckBox::indicator { width: 44px; height: 24px; }
            QCheckBox::indicator:unchecked { border-radius: 12px; background: #dfe7f2; }
            QCheckBox::indicator:unchecked:hover { background: #cfd9e9; }
            QCheckBox::indicator:checked { border-radius: 12px; background: #113a70; }
            QCheckBox::indicator:checked:hover { background: #0f335f; }
            /* Improve QMessageBox legibility */
            QMessageBox { background-color: #113a70; }
            QMessageBox QLabel { color: #ffffff; }
            """
        )

        central = QtWidgets.QWidget()
        self.setCentralWidget(central)

        # Header
        header = QtWidgets.QFrame()
        hbox = QtWidgets.QHBoxLayout(header)
        logo = QtWidgets.QLabel()
        logo_pix = self._build_logo_pixmap()
        logo.setPixmap(logo_pix)
        logo.setFixedSize(logo_pix.size())
        hbox.addWidget(logo)

        title = QtWidgets.QLabel("EIDAT")
        font = title.font(); font.setPointSize(20); font.setBold(True); title.setFont(font)
        subtitle = QtWidgets.QLabel("End Item Data Analysis Tool"); subtitle.setStyleSheet("color:#5b6b7a; font-size: 12px;")
        tbox = QtWidgets.QVBoxLayout(); tbox.addWidget(title); tbox.addWidget(subtitle)
        hbox.addLayout(tbox); hbox.addStretch(1)
        self.lbl_ready = QtWidgets.QLabel("System Ready"); self.lbl_ready.setObjectName("statusBadge"); hbox.addWidget(self.lbl_ready)

        # Tabs and log
        self.tabs = QtWidgets.QTabWidget()
        self.tab_setup = QtWidgets.QWidget()
        self.tab_process = QtWidgets.QWidget()
        self.tab_plot = QtWidgets.QWidget()
        self.tab_outputs = QtWidgets.QWidget()
        self.tabs.addTab(self.tab_setup, "Setup")
        self.tabs.addTab(self.tab_process, "EIDP Processing")
        self.tabs.addTab(self.tab_plot, "Analysis")
        self.tabs.currentChanged.connect(self._on_tab_changed)
        # Center tabs across the top by expanding them
        try:
            self.tabs.tabBar().setExpanding(True)
        except Exception:
            pass

        self._tab_style_template = (
            "QTabWidget::pane { border: none; margin-top: -2px; background: transparent; border-top: 3px solid rgba(16, 52, 90, 0.85); }\n"
            "QTabWidget::tab-bar { alignment: center; }\n"
            "QTabBar { qproperty-drawBase: 0; }\n"
            "QTabBar::tab {\n"
            "    padding: 18px 0;\n"
            "    margin: 0;\n"
            "    font-weight: 700;\n"
            "    letter-spacing: 0.5px;\n"
            "    color: #ffffff;\n"
            "    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #173761, stop:1 #2c68af);\n"
            "    border: none;\n"
            "    border-right: 1px solid rgba(255,255,255,0.25);\n"
            "    border-bottom: 4px solid rgba(0, 0, 0, 0.12);\n"
            "    min-width: {width}px;\n"
            "}\n"
            "QTabBar::tab:last { border-right: none; }\n"
            "QTabBar::tab:selected { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #0d2a48, stop:1 #1b4d85); border-bottom: 4px solid rgba(255,255,255,0.45); }\n"
            "QTabBar::tab:hover { background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #1d4c82, stop:1 #2d79bd); }\n"
        )

        self._apply_tab_widths()
        QtCore.QTimer.singleShot(0, self._apply_tab_widths)

        self.log = QtWidgets.QPlainTextEdit(); self.log.setReadOnly(True); self.log.setMaximumBlockCount(5000)
        # Toggle to show/hide the debug log panel on demand
        self.btn_toggle_log = QtWidgets.QPushButton("Show Debug Panel")
        self.btn_toggle_log.setCheckable(True)
        self.btn_toggle_log.setChecked(False)
        self.btn_toggle_log.clicked.connect(self._toggle_log_panel)
        pol_log = self.btn_toggle_log.sizePolicy(); pol_log.setHorizontalStretch(1); pol_log.setHorizontalPolicy(QtWidgets.QSizePolicy.Policy.Expanding); self.btn_toggle_log.setSizePolicy(pol_log)
        self.status_bar = self.statusBar()
        self._progress_dialog = RunProgressDialog(self)
        self._progress_dialog.canceled.connect(self._on_progress_canceled)
        self._progress_pattern = re.compile(r"\[PROGRESS\]\s*Terms:\s*(\d+)%\s*\((\d+)/(\d+)\)")
        self._progress_total = 0
        self._progress_completed = 0
        self._progress_popup_active = False
        self._progress_was_canceled = False
        self._last_run_dir: Path | None = None

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

        # Program Health
        grp_env = QtWidgets.QGroupBox("Program Health")
        grp_env.setAlignment(QtCore.Qt.AlignmentFlag.AlignHCenter)
        l_env = QtWidgets.QGridLayout(grp_env)
        l_env.addWidget(QtWidgets.QLabel("Ensure all dependencies are installed and up to date"), 0, 0, 1, 3)
        self.lbl_env_health = QtWidgets.QLabel("Healthy"); self.lbl_env_health.setObjectName("healthBadge")
        l_env.addWidget(
            self.lbl_env_health,
            0,
            3,
            alignment=QtCore.Qt.AlignmentFlag.AlignRight,
        )
        # spacing and stretch to prevent label/button overlap
        l_env.setHorizontalSpacing(10)
        l_env.setVerticalSpacing(10)
        for c in range(3):
            l_env.setColumnStretch(c, 1)
        l_env.setColumnStretch(3, 0)
        self.btn_check = QtWidgets.QPushButton("Check Environment"); self.btn_check.setProperty("variant", "ghost")
        self.btn_install = QtWidgets.QPushButton("Update Packages"); self.btn_install.setProperty("variant", "primary")
        self.btn_check.clicked.connect(self._act_check_env)
        self.btn_install.clicked.connect(self._act_install)
        self.lbl_env = QtWidgets.QLabel("Env: Unknown")
        l_env.addWidget(self.btn_check, 1, 0); l_env.addWidget(self.btn_install, 1, 1)
        l_env.addWidget(self.lbl_env, 2, 0, 1, 4)

        # Extraction Settings
        grp_set = QtWidgets.QGroupBox("Extraction Settings")
        grp_set.setAlignment(QtCore.Qt.AlignmentFlag.AlignHCenter)
        ls = QtWidgets.QGridLayout(grp_set)
        hint = QtWidgets.QLabel("Configure how EIDAT processes and extracts data from documents"); hint.setProperty("class", "subtle")
        ls.addWidget(hint, 0, 0, 1, 3)
        ls.addWidget(QtWidgets.QLabel("OCR Mode"), 1, 0)
        # Friendly OCR mode labels mapped to env values
        self._ocr_value_to_display = {
            "fallback": "Read PDF and fallback to OCR if needed",
            "ocr_only": "Only OCR read the selected documents",
            "no_ocr": "Read PDF, no OCR (may fail)",
        }
        self._ocr_display_to_value = {v: k for k, v in self._ocr_value_to_display.items()}
        self.cmb_ocr_mode = QtWidgets.QComboBox();
        self.cmb_ocr_mode.addItems(list(self._ocr_value_to_display.values())); ls.addWidget(self.cmb_ocr_mode, 2, 0, 1, 3)
        # Separator
        sep1 = QtWidgets.QFrame(); sep1.setFrameShape(QtWidgets.QFrame.Shape.HLine); sep1.setFrameShadow(QtWidgets.QFrame.Shadow.Sunken)
        ls.addWidget(sep1, 3, 0, 1, 3)
        ls.addWidget(QtWidgets.QLabel("XY Fuzz Tolerance"), 4, 0)
        self.sld_xy_fuzz = QtWidgets.QSlider(QtCore.Qt.Orientation.Horizontal); self.sld_xy_fuzz.setRange(0, 100)
        self.lbl_xy_val = QtWidgets.QLabel("0.50")
        self.lbl_xy_val.setAlignment(QtCore.Qt.AlignmentFlag.AlignRight)
        self.lbl_xy_val.setMinimumWidth(40)
        ls.addWidget(self.sld_xy_fuzz, 5, 0, 1, 2); ls.addWidget(self.lbl_xy_val, 5, 2)
        lbl_xy_desc = QtWidgets.QLabel("Tolerance for matching table cell positions (higher = more lenient)"); lbl_xy_desc.setProperty("class", "subtle"); ls.addWidget(lbl_xy_desc, 6, 0, 1, 3)
        # Separator
        sep2 = QtWidgets.QFrame(); sep2.setFrameShape(QtWidgets.QFrame.Shape.HLine); sep2.setFrameShadow(QtWidgets.QFrame.Shadow.Sunken)
        ls.addWidget(sep2, 7, 0, 1, 3)
        # OCR DPI control (slider 500-1000)
        ls.addWidget(QtWidgets.QLabel("OCR DPI"), 8, 0)
        self.sld_ocr_dpi = QtWidgets.QSlider(QtCore.Qt.Orientation.Horizontal)
        self.sld_ocr_dpi.setRange(500, 1000)
        self.sld_ocr_dpi.setSingleStep(25)
        self.lbl_dpi_val = QtWidgets.QLabel("500")
        self.lbl_dpi_val.setAlignment(QtCore.Qt.AlignmentFlag.AlignRight)
        self.lbl_dpi_val.setMinimumWidth(48)
        ls.addWidget(self.sld_ocr_dpi, 9, 0, 1, 2)
        ls.addWidget(self.lbl_dpi_val, 9, 2)
        lbl_dpi_desc = QtWidgets.QLabel("Higher DPI may improve OCR accuracy at the cost of speed")
        lbl_dpi_desc.setProperty("class", "subtle"); ls.addWidget(lbl_dpi_desc, 10, 0, 1, 3)
        # Separator
        sep3 = QtWidgets.QFrame(); sep3.setFrameShape(QtWidgets.QFrame.Shape.HLine); sep3.setFrameShadow(QtWidgets.QFrame.Shadow.Sunken)
        ls.addWidget(sep3, 11, 0, 1, 3)
        self.chk_logging = QtWidgets.QCheckBox("Show Debug Logs")
        self.chk_logging.setToolTip("Writes detailed messages to help troubleshoot issues (slightly slower)")
        ls.addWidget(self.chk_logging, 12, 0, 1, 3)
        lbl_log_desc = QtWidgets.QLabel("Show detailed debug logs (slightly slower)"); lbl_log_desc.setProperty("class", "subtle"); ls.addWidget(lbl_log_desc, 13, 0, 1, 3)
        ls.addWidget(QtWidgets.QLabel("OCR Language"), 14, 0)
        # Language display mapping (user-friendly names)
        self._lang_display_to_code = {
            "English": "en",
            "French": "fr",
            "German": "de",
            "Spanish": "es",
        }
        self._lang_code_to_display = {v: k for k, v in self._lang_display_to_code.items()}
        self.cmb_lang = QtWidgets.QComboBox()
        self.cmb_lang.addItems(list(self._lang_display_to_code.keys()))
        ls.addWidget(self.cmb_lang, 15, 0, 1, 3)
        # layout spacing & stretch to avoid overlap
        ls.setHorizontalSpacing(10)
        ls.setVerticalSpacing(8)
        ls.setColumnStretch(0, 1); ls.setColumnStretch(1, 1); ls.setColumnStretch(2, 0)
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
        grid = QtWidgets.QGridLayout(self.tab_process)

        # Master button at top (above Define Inputs)
        self.btn_open_master_tab = QtWidgets.QPushButton("Open Master Database")
        # Black outline style for differentiation
        self.btn_open_master_tab.setStyleSheet("QPushButton { border: 2px solid #000; color: #000; background: #ffffff; padding: 10px 16px; border-radius: 6px; } QPushButton:hover { background: #f5f5f5; }")
        polm = self.btn_open_master_tab.sizePolicy(); polm.setHorizontalStretch(1); polm.setHorizontalPolicy(QtWidgets.QSizePolicy.Policy.Expanding); self.btn_open_master_tab.setSizePolicy(polm)
        self.btn_open_master_tab.clicked.connect(lambda: self._safe_open(be.open_master_workbook))
        grid.addWidget(self.btn_open_master_tab, 0, 0, 1, 2)

        # Define Inputs
        grp_inputs = QtWidgets.QGroupBox("Define Inputs")
        li = QtWidgets.QGridLayout(grp_inputs)
        self.ed_terms = QtWidgets.QLineEdit(str(be.DEFAULT_TERMS_XLSX))
        btn_browse_terms = QtWidgets.QPushButton("Browse...")
        btn_browse_terms.clicked.connect(lambda: self._browse_file(self.ed_terms, be.DEFAULT_TERMS_XLSX.parent, "Smart Snap Terms (*.xlsx);;All files (*.*)"))
        self.btn_terms_edit = QtWidgets.QPushButton("Edit Smart-Snap Terms")
        self.btn_terms_edit.setProperty("variant", "primary")
        self.btn_terms_refresh = QtWidgets.QPushButton("Create/Refresh Input Spreadsheet")
        self.btn_terms_edit.clicked.connect(self._open_terms_editor)
        self.btn_terms_refresh.clicked.connect(self._act_generate_terms)
        self.ed_pdfs = QtWidgets.QLineEdit(str(be.DEFAULT_PDF_DIR))
        li.addWidget(self.btn_terms_edit, 0, 0, 1, 2)
        li.addWidget(self.btn_terms_refresh, 1, 0, 1, 2)
        # Keep internal path field for logic, but do not show it

        # Upload (repurposed as workspace sync)
        grp_upload = QtWidgets.QGroupBox("Data Upload")
        up = QtWidgets.QGridLayout(grp_upload)
        self.btn_sync_workspace = QtWidgets.QPushButton("Sync Workspace Now")
        self.btn_sync_workspace.setProperty("variant", "primary")
        self.btn_sync_workspace.clicked.connect(self._act_sync_workspace)
        up.addWidget(self.btn_sync_workspace, 0, 0, 1, 2)

        # Repository root picker
        up.addWidget(QtWidgets.QLabel("Repository Root"), 1, 0)
        self.ed_repo = QtWidgets.QLineEdit(str(getattr(be, 'get_repo_root', lambda: be.DEFAULT_REPO_ROOT)()))
        btn_repo = QtWidgets.QPushButton("Browse…")
        btn_repo.clicked.connect(lambda: self._browse_folder(self.ed_repo, be.DEFAULT_PDF_DIR))
        row_repo = QtWidgets.QHBoxLayout(); row_repo.addWidget(self.ed_repo, 1); row_repo.addWidget(btn_repo)
        wrapper = QtWidgets.QWidget(); wrapper.setLayout(row_repo)
        up.addWidget(wrapper, 1, 1)

        self.lbl_sync_banner = QtWidgets.QLabel("No sync run yet.")
        self.lbl_sync_banner.setObjectName("syncBanner")
        self.lbl_sync_banner.setWordWrap(True)
        self.lbl_sync_banner.setStyleSheet(
            "#syncBanner { background: #f0f3f7; color: #0f2a46; border: 1px solid #c8d3e5; border-radius: 6px; padding: 8px 12px; }"
        )
        up.addWidget(self.lbl_sync_banner, 2, 0, 1, 2)

        self.btn_view_outdated = QtWidgets.QPushButton("View Data Package List and Update EIDAT Database")
        self.btn_view_outdated.setProperty("variant", "primary")
        pol = self.btn_view_outdated.sizePolicy(); pol.setHorizontalStretch(1); pol.setHorizontalPolicy(QtWidgets.QSizePolicy.Policy.Expanding); self.btn_view_outdated.setSizePolicy(pol)
        self.btn_view_outdated.clicked.connect(self._show_outdated_popup)
        up.addWidget(self.btn_view_outdated, 3, 0, 1, 2)

        self.btn_view_registry2 = QtWidgets.QPushButton("View Registry")
        self.btn_view_registry2.clicked.connect(self._act_view_registry)
        pol2 = self.btn_view_registry2.sizePolicy(); pol2.setHorizontalStretch(1); pol2.setHorizontalPolicy(QtWidgets.QSizePolicy.Policy.Expanding); self.btn_view_registry2.setSizePolicy(pol2)
        up.addWidget(self.btn_view_registry2, 4, 0, 1, 2)

        # Quick cleanup: remove run_data folders not referenced by registry
        self.btn_clear_old_runs = QtWidgets.QPushButton("Clear Old Run Cache")
        self.btn_clear_old_runs.clicked.connect(self._act_clear_old_runs)
        pol3 = self.btn_clear_old_runs.sizePolicy(); pol3.setHorizontalStretch(1); pol3.setHorizontalPolicy(QtWidgets.QSizePolicy.Policy.Expanding); self.btn_clear_old_runs.setSizePolicy(pol3)
        up.addWidget(self.btn_clear_old_runs, 5, 0, 1, 2)

        # Removed Open Repository Folder button per UX simplification

        # Processing + Outputs
        grp_proc = QtWidgets.QGroupBox("Smart Processing Controls")
        lp = QtWidgets.QHBoxLayout(grp_proc)
        self.btn_start = QtWidgets.QPushButton("Extract and Update All"); self.btn_start.setProperty("variant", "primary")
        self.btn_stop = QtWidgets.QPushButton("Stop Scan")
        self.btn_open_last = QtWidgets.QPushButton("Open Last Run Folder")
        self.btn_start.clicked.connect(self._act_start_scan)
        self.btn_stop.clicked.connect(self._act_stop_scan)
        self.btn_open_last.clicked.connect(lambda: self._safe_open(be.open_last_run_folder))
        lp.addWidget(self.btn_start); lp.addWidget(self.btn_stop); lp.addWidget(self.btn_open_last)

        grid.addWidget(grp_inputs, 1, 0, 1, 2)
        grid.addWidget(grp_upload, 2, 0, 1, 2)
        grid.addWidget(grp_proc, 3, 0, 1, 2)
        grid.setRowStretch(4, 1)
    def _setup_tab_plot(self):
        grid = QtWidgets.QGridLayout(self.tab_plot)
        grp_dash = QtWidgets.QGroupBox("Plotting Dashboard")
        ld = QtWidgets.QGridLayout(grp_dash)
        self.btn_plot_terms_create = QtWidgets.QPushButton("Generate Available Term List")
        self.btn_plot_terms_create.clicked.connect(self._act_generate_plot_terms)
        self.btn_open_plot_terms = QtWidgets.QPushButton("View Terms List")
        self.btn_open_plot_terms.clicked.connect(lambda: self._safe_open(lambda: be.open_path(be.DEFAULT_PLOT_TERMS_XLSX)))
        ld.addWidget(self.btn_plot_terms_create, 0, 0)
        ld.addWidget(self.btn_open_plot_terms, 0, 1)

        grp_gen = QtWidgets.QGroupBox("Proposed Plots")
        lg = QtWidgets.QVBoxLayout(grp_gen)
        self.lbl_proposed_summary = QtWidgets.QLabel("No plots configured.")
        self.lbl_proposed_summary.setWordWrap(True)
        lg.addWidget(self.lbl_proposed_summary)
        controls = QtWidgets.QHBoxLayout()
        self.btn_manage_proposed = QtWidgets.QPushButton("Manage Proposed Plots…")
        self.btn_manage_proposed.clicked.connect(self._open_proposed_plots_dialog)
        self.btn_refresh_series = QtWidgets.QPushButton("Reload Series List")
        self.btn_refresh_series.clicked.connect(self._refresh_series_catalog)
        controls.addWidget(self.btn_manage_proposed)
        controls.addWidget(self.btn_refresh_series)
        controls.addStretch(1)
        lg.addLayout(controls)

        grp_ops = QtWidgets.QGroupBox("Generate Plots")
        lo = QtWidgets.QHBoxLayout(grp_ops)
        self.btn_plots_generate = QtWidgets.QPushButton("Generate Plots")
        self.btn_plots_generate.setProperty("variant", "primary")
        self.btn_plots_generate.clicked.connect(self._act_generate_plots)
        self.btn_plot_summary = QtWidgets.QPushButton("Plot Summary Report")
        self.btn_plot_summary.clicked.connect(self._act_export_plot_summary)
        self.btn_plots_open_folder = QtWidgets.QPushButton("Open Plots Folder")
        self.btn_plots_open_folder.clicked.connect(lambda: self._safe_open(be.open_plots_folder))
        lo.addWidget(self.btn_plots_generate)
        lo.addWidget(self.btn_plot_summary)
        lo.addStretch(1)
        lo.addWidget(self.btn_plots_open_folder)

        grid.addWidget(grp_dash, 0, 0, 1, 2)
        grid.addWidget(grp_gen, 1, 0, 1, 2)
        grid.addWidget(grp_ops, 2, 0, 1, 2)
        grid.setRowStretch(3, 1)

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
    def _on_tab_changed(self, idx: int):
        try:
            self._apply_tab_widths()
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
        except Exception:
            return
        self._progress_total = max(total, 0)
        self._progress_completed = max(0, min(completed, self._progress_total or completed))
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
        self._progress_dialog.update_progress(completed, total)

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

    def _start_worker(self, popen_factory, *, status_msg: str, show_run_progress: bool = False):
        if self._worker is not None and self._worker.isRunning():
            return
        self._append_log(f"[GUI] {status_msg}")
        self.status_bar.showMessage(status_msg)
        self._progress_total = 0
        self._progress_completed = 0
        self._progress_popup_active = show_run_progress
        self._progress_was_canceled = False
        self._last_run_dir = None
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
            self.btn_toggle_log.setText("Hide Debug Panel" if visible else "Show Debug Panel")
        except Exception:
            pass

    def _on_worker_done(self, rc: int):
        self.status_bar.showMessage("Ready.", 3000)
        self._append_log(f"[INFO] Process finished with code {rc}")
        was_canceled = self._progress_was_canceled
        self._finalize_run_progress(success=(rc == 0))
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
        self._scan_refresh()

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
        dlg.exec()
        # No immediate auto-sync; the periodic timer will refresh the banner

    def _act_generate_terms(self):
        # Warn if a terms spreadsheet exists and will be overwritten
        try:
            target = Path(self.ed_terms.text()).expanduser()
        except Exception:
            target = be.DEFAULT_TERMS_XLSX if hasattr(be, "DEFAULT_TERMS_XLSX") else Path("terms.xlsx")
        if target.exists():
            msg = (
                f"A terms spreadsheet already exists and will be overwritten.\n\n"
                f"{target}\n\n"
                "If you wish to keep it, please back it up first.\n\n"
                "Do you want to overwrite it?"
            )
            if (
                QtWidgets.QMessageBox.question(
                    self, "Overwrite terms.xlsx?", msg
                )
                != QtWidgets.QMessageBox.StandardButton.Yes
            ):
                return
        # Kick off generation (backend creates terms.schema.smartsnap.xlsx)
        self._start_worker(be.generate_terms, status_msg="Generating Smart-Snap terms spreadsheet...")

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
        self._start_worker(lambda: be.run_scanner(terms, pdfs), status_msg="Scanning PDFs...", show_run_progress=True)

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
            v = QtWidgets.QVBoxLayout(dlg)
            tbl = QtWidgets.QTableWidget(0, len(headers))
            tbl.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.AllEditTriggers)
            tbl.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
            tbl.setAlternatingRowColors(True)
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
        except Exception as e:
            self.lbl_sync_banner.setText(f"Sync failed: {e}")

    def _act_sync_workspace(self):
        self._sync_workspace(auto=False)

    def _show_outdated_popup(self, auto: bool = False):
        # Guard against duplicate dialogs
        if getattr(self, "_dlg_open_outdated", False):
            return
        self._dlg_open_outdated = True
        try:
            details = getattr(self, "_sync_details", None) or []
            rows = [d for d in details if d.get("reason") in ("new", "pdf_newer", "terms_newer")]
            dlg = QtWidgets.QDialog(self)
            dlg.setWindowTitle("Out-of-Date EIDPs")
            dlg.resize(900, 520)
            v = QtWidgets.QVBoxLayout(dlg)
            toolbar = QtWidgets.QHBoxLayout()
            btn_sel_all = QtWidgets.QPushButton("Select All")
            btn_sel_none = QtWidgets.QPushButton("Select None")
            toolbar.addWidget(btn_sel_all)
            toolbar.addWidget(btn_sel_none)
            toolbar.addStretch(1)
            v.addLayout(toolbar)
            cols = ["Select", "Serial", "Reason", "PDF", "Run Date", "PDF Modified", "Terms Modified"]
            tbl = QtWidgets.QTableWidget(0, len(cols))
            tbl.setHorizontalHeaderLabels(cols)
            tbl.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows)
            tbl.setAlternatingRowColors(True)
            v.addWidget(tbl, 1)
            for r, d in enumerate(rows):
                tbl.insertRow(r)
                it = QtWidgets.QTableWidgetItem()
                it.setFlags(QtCore.Qt.ItemFlag.ItemIsUserCheckable | QtCore.Qt.ItemFlag.ItemIsEnabled)
                it.setCheckState(QtCore.Qt.CheckState.Checked)
                tbl.setItem(r, 0, it)
                tbl.setItem(r, 1, QtWidgets.QTableWidgetItem(d.get("serial_component", "")))
                tbl.setItem(r, 2, QtWidgets.QTableWidgetItem(d.get("reason", "")))
                tbl.setItem(r, 3, QtWidgets.QTableWidgetItem(d.get("pdf", "")))
                tbl.setItem(r, 4, QtWidgets.QTableWidgetItem(d.get("run_date", "")))
                tbl.setItem(r, 5, QtWidgets.QTableWidgetItem(d.get("pdf_mtime", "")))
                tbl.setItem(r, 6, QtWidgets.QTableWidgetItem(d.get("terms_mtime", "")))
            tbl.resizeColumnsToContents()

            def _set_all(state: QtCore.Qt.CheckState):
                for r in range(tbl.rowCount()):
                    it = tbl.item(r, 0)
                    if it:
                        it.setCheckState(state)

            btn_sel_all.clicked.connect(lambda: _set_all(QtCore.Qt.CheckState.Checked))
            btn_sel_none.clicked.connect(lambda: _set_all(QtCore.Qt.CheckState.Unchecked))

            btns = QtWidgets.QHBoxLayout()
            btn_run_all = QtWidgets.QPushButton("Run All Out-of-Date")
            btn_run = QtWidgets.QPushButton("Run Selected")
            btn_close = QtWidgets.QPushButton("Close")
            btns.addStretch(1)
            btns.addWidget(btn_run_all)
            btns.addWidget(btn_run)
            btns.addWidget(btn_close)
            v.addLayout(btns)

            def _run_selected():
                paths: list[Path] = []
                for r in range(tbl.rowCount()):
                    it = tbl.item(r, 0)
                    if it and it.checkState() == QtCore.Qt.CheckState.Checked:
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
                self._start_worker(lambda: be.run_selected_pdfs(paths, terms), status_msg="Running selected EIDPs...", show_run_progress=True)
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
                self._start_worker(lambda: be.run_selected_pdfs(all_paths, terms), status_msg="Running all out-of-date EIDPs...", show_run_progress=True)
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
        try:
            self._apply_tab_widths()
        except Exception:
            pass

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
        # Fallback: draw the original glyph
        pix = QtGui.QPixmap(size, size)
        pix.fill(QtCore.Qt.GlobalColor.transparent)
        painter = QtGui.QPainter(pix)
        painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing)
        bg = QtGui.QColor("#1f5c9a")
        painter.setBrush(QtGui.QBrush(bg))
        painter.setPen(QtCore.Qt.PenStyle.NoPen)
        painter.drawRoundedRect(0, 0, size, size, 8, 8)
        painter.setPen(QtGui.QPen(QtGui.QColor("#ffffff")))
        font = painter.font(); font.setBold(True); font.setPointSize(int(size * 0.42)); painter.setFont(font)
        painter.drawText(pix.rect(), QtCore.Qt.AlignmentFlag.AlignCenter, "E")
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


def main():
    app = QtWidgets.QApplication(sys.argv)
    w = MainWindow(); w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
