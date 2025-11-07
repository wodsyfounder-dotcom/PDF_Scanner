from __future__ import annotations

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

class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("EIDAT - Engineering End Item Data Analysis Tool")
        self.resize(1280, 860)

        be.ensure_scaffold()

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
        subtitle = QtWidgets.QLabel("Engineering End Item Data Analysis Tool"); subtitle.setStyleSheet("color:#5b6b7a; font-size: 12px;")
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
        self.status_bar = self.statusBar()

        layout = QtWidgets.QVBoxLayout(central)
        layout.addWidget(header)
        layout.addWidget(self.tabs)
        layout.addWidget(self.log, 1)
        # Build tabs
        self._setup_tab_setup()
        self._setup_tab_process()
        self._setup_tab_plot()
        self._setup_tab_outputs()

        # Runtime
        self._worker: ProcWorker | None = None
        self._enrich_after_run: bool = False
        self._registry_cache: tuple[list[str], list[list[str]]] | None = None
        self._scan_refresh()

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
        self.ed_scanned = QtWidgets.QLineEdit(str(be.DEFAULT_SCANNED_DIR))
        li.addWidget(self.btn_terms_edit, 0, 0, 1, 2)
        li.addWidget(self.btn_terms_refresh, 1, 0, 1, 2)
        # Keep internal path field for logic, but do not show it

        # Upload
        grp_upload = QtWidgets.QGroupBox("Data Upload")
        up = QtWidgets.QGridLayout(grp_upload)
        self.drop_zone = _DropZone("Drop PDFs here", on_drop=self._ingest_paths)
        up.addWidget(self.drop_zone, 0, 0)
        right_box = QtWidgets.QGroupBox("Staged PDFs in folder")
        rl = QtWidgets.QVBoxLayout(right_box)
        self.list_pdfs = QtWidgets.QTableWidget(0, 3)
        self.list_pdfs.setHorizontalHeaderLabels(["File", "Size", "Modified"])
        header_view = self.list_pdfs.horizontalHeader()
        header_view.setSectionResizeMode(0, QtWidgets.QHeaderView.ResizeMode.Stretch)
        header_view.setSectionResizeMode(
            1, QtWidgets.QHeaderView.ResizeMode.ResizeToContents
        )
        header_view.setSectionResizeMode(
            2, QtWidgets.QHeaderView.ResizeMode.ResizeToContents
        )
        self.list_pdfs.setSelectionBehavior(
            QtWidgets.QAbstractItemView.SelectionBehavior.SelectRows
        )
        self.list_pdfs.setSelectionMode(
            QtWidgets.QAbstractItemView.SelectionMode.ExtendedSelection
        )
        rl.addWidget(self.list_pdfs, 1)
        bbar = QtWidgets.QHBoxLayout();
        self.btn_remove_selected = QtWidgets.QPushButton("Remove Selected"); self.btn_remove_selected.clicked.connect(self._act_remove_selected)
        self.btn_clear_all = QtWidgets.QPushButton("Remove All"); self.btn_clear_all.clicked.connect(self._act_remove_all)
        self.btn_open_pdfs = QtWidgets.QPushButton("Open PDFs Folder"); self.btn_open_pdfs.clicked.connect(lambda: self._safe_open(lambda: be.open_path(Path(self.ed_pdfs.text()).expanduser())))
        bbar.addWidget(self.btn_remove_selected); bbar.addWidget(self.btn_clear_all); bbar.addStretch(1); bbar.addWidget(self.btn_open_pdfs)
        rl.addLayout(bbar)
        up.addWidget(right_box, 0, 1)
        cbar = QtWidgets.QHBoxLayout();
        self.btn_add_files = QtWidgets.QPushButton("Add Files..."); self.btn_add_files.clicked.connect(self._act_add_files)
        self.btn_add_folder = QtWidgets.QPushButton("Add Folder..."); self.btn_add_folder.clicked.connect(self._act_add_folder)
        cbar.addWidget(self.btn_add_files); cbar.addWidget(self.btn_add_folder); cbar.addStretch(1)
        up.addLayout(cbar, 1, 0, 1, 2)

        # Processing + Outputs
        grp_proc = QtWidgets.QGroupBox("Processing Controls")
        lp = QtWidgets.QHBoxLayout(grp_proc)
        self.btn_start = QtWidgets.QPushButton("Start Scan"); self.btn_start.setProperty("variant", "primary")
        self.btn_stop = QtWidgets.QPushButton("Stop Scan")
        self.btn_open_last = QtWidgets.QPushButton("Open Last Run Folder")
        self.btn_start.clicked.connect(self._act_start_scan)
        self.btn_stop.clicked.connect(self._act_stop_scan)
        self.btn_open_last.clicked.connect(lambda: self._safe_open(be.open_last_run_folder))
        lp.addWidget(self.btn_start); lp.addWidget(self.btn_stop); lp.addWidget(self.btn_open_last)

        grp_out = QtWidgets.QGroupBox("Data Outputs")
        lo = QtWidgets.QVBoxLayout(grp_out)
        # Buttons row
        btn_row = QtWidgets.QHBoxLayout()
        self.btn_compile = QtWidgets.QPushButton("Compile Master")
        self.btn_open_master = QtWidgets.QPushButton("Open Master")
        self.btn_open_registry = QtWidgets.QPushButton("Open Run Registry")
        self.btn_view_registry = QtWidgets.QPushButton("View Registry")
        self.btn_compile.clicked.connect(self._act_compile_master)
        self.btn_open_master.clicked.connect(lambda: self._safe_open(be.open_master_workbook))
        self.btn_open_registry.clicked.connect(lambda: self._safe_open(be.open_run_registry))
        self.btn_view_registry.clicked.connect(self._show_registry_popup)
        btn_row.addWidget(self.btn_compile); btn_row.addWidget(self.btn_open_master); btn_row.addWidget(self.btn_open_registry); btn_row.addWidget(self.btn_view_registry); btn_row.addStretch(1)
        lo.addLayout(btn_row)
        # Popup viewer is created on demand; no inline viewer

        grid.addWidget(grp_inputs, 0, 0, 1, 2)
        grid.addWidget(grp_upload, 1, 0, 1, 2)
        grid.addWidget(grp_proc, 2, 0, 1, 2)
        grid.addWidget(grp_out, 3, 0, 1, 2)
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
        lg = QtWidgets.QGridLayout(grp_gen)
        self.list_plot_names = QtWidgets.QListWidget()
        self.list_plot_names.setSelectionMode(QtWidgets.QAbstractItemView.SelectionMode.NoSelection)
        self.list_plot_names.itemChanged.connect(self._on_plot_name_toggled)
        lg.addWidget(self.list_plot_names, 0, 0, 1, 2)
        self.btn_apply_plot_names = QtWidgets.QPushButton("Use Selection")
        self.btn_apply_plot_names.clicked.connect(self._apply_plot_name_selection)
        self.btn_refresh_plot_names = QtWidgets.QPushButton("Refresh")
        self.btn_refresh_plot_names.clicked.connect(self._load_plot_names_list)
        lg.addWidget(self.btn_apply_plot_names, 1, 0)
        lg.addWidget(self.btn_refresh_plot_names, 1, 1)

        grp_ops = QtWidgets.QGroupBox("Generate Plots")
        lo = QtWidgets.QHBoxLayout(grp_ops)
        self.btn_plots_generate = QtWidgets.QPushButton("Generate Plots")
        self.btn_plots_generate.setProperty("variant", "primary")
        self.btn_plots_generate.clicked.connect(self._act_generate_plots_save_selection)
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

    def _setup_tab_outputs(self):
        grid = QtWidgets.QGridLayout(self.tab_outputs)
        grid.addWidget(QtWidgets.QLabel("Run Registry UI coming soon. Use buttons above to open registry and master."), 0, 0)
        grid.setRowStretch(1, 1)

    # Actions & helpers
    def _on_tab_changed(self, idx: int):
        try:
            self._apply_tab_widths()
            if self.tabs.widget(idx) is self.tab_plot:
                self._load_plot_names_list()
            if self.tabs.widget(idx) is self.tab_process or self.tabs.widget(idx) is self.tab_outputs:
                self._refresh_run_registry()
        except Exception:
            pass

    def _append_log(self, text: str):
        self.log.appendPlainText(text)
        self.log.verticalScrollBar().setValue(self.log.verticalScrollBar().maximum())

    def _start_worker(self, popen_factory, *, status_msg: str):
        if self._worker is not None and self._worker.isRunning():
            return
        self._append_log(f"[GUI] {status_msg}")
        self.status_bar.showMessage(status_msg)
        self._worker = ProcWorker(popen_factory)
        self._worker.line.connect(self._append_log)
        self._worker.finished.connect(self._on_worker_done)
        self._worker.start()

    def _on_worker_done(self, rc: int):
        self.status_bar.showMessage("Ready.", 3000)
        self._append_log(f"[INFO] Process finished with code {rc}")
        # Refresh UI after a scan; no post-run enrichment step
        if getattr(self, "_enrich_after_run", False):
            self._enrich_after_run = False
            try:
                self._refresh_run_registry()
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
        scanned = Path(self.ed_scanned.text()).expanduser()
        if not terms.exists():
            QtWidgets.QMessageBox.critical(self, "Missing terms", f"Terms file not found:\n{terms}")
            return
        if not pdfs.exists():
            QtWidgets.QMessageBox.critical(self, "Missing PDFs folder", f"PDFs folder not found:\n{pdfs}")
            return
        scanned.mkdir(parents=True, exist_ok=True)
        # Only enrich registry after a scan completes
        self._enrich_after_run = True
        self._start_worker(lambda: be.run_scanner(terms, pdfs, scanned), status_msg="Scanning PDFs...")

    def _act_stop_scan(self):
        try:
            if self._worker and self._worker.isRunning():
                self._worker.stop()
        except Exception:
            pass

    def _act_compile_master(self):
        self._start_worker(be.compile_master, status_msg="Compiling master workbook...")

    def _act_generate_plot_terms(self):
        self._start_worker(be.generate_plot_terms, status_msg="Generating plot terms...")

    def _act_generate_plots_save_selection(self):
        # Persist plot name selection before generating
        self._apply_plot_name_selection(show_status=False)
        self._act_generate_plots()

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
                try:
                    import openpyxl  # type: ignore
                    wb = openpyxl.load_workbook(str(p_xlsx), read_only=True, data_only=True)
                    ws = wb.active
                    if ws is None:
                        # No active worksheet; treat as empty workbook
                        headers = []
                        rows = []
                    else:
                        it = ws.iter_rows(values_only=True)
                        try:
                            headers = [str(x) if x is not None else "" for x in next(it)]
                        except StopIteration:
                            headers = []
                        for r in it:
                            rows.append([str(x) if x is not None else "" for x in r])
                except Exception:
                    headers = ["Run Registry"]
                    rows = [[f"Preview requires CSV or openpyxl: {p_xlsx.name}"]]
            else:
                headers = ["Run Registry"]
                rows = [["No registry found. Run a scan to create it."]]

            # Derive Program/Vehicle/Serial from PDF filename if possible
            try:
                hdr_lc = [h.strip().lower() for h in headers]
                file_candidates = {"file", "filename", "pdf", "document", "input", "input file"}
                idx_file = -1
                for i, h in enumerate(hdr_lc):
                    if h in file_candidates or h.endswith(" file") or h.endswith(" pdf"):
                        idx_file = i
                        break
                def _derive_from_row(row: list[str]) -> tuple[str, str, str]:
                    name = ""
                    if idx_file >= 0 and idx_file < len(row):
                        name = str(row[idx_file] or "")
                    if not name or ".pdf" not in name.lower():
                        for c in row:
                            s = str(c or "")
                            if s.lower().endswith(".pdf"):
                                name = s
                                break
                    base_name = Path(name).name
                    if base_name.lower().endswith(".pdf"):
                        base_name = base_name[:-4]
                    parts = [p for p in base_name.split("_") if p]
                    if len(parts) >= 3:
                        program, vehicle, serial = parts[0], parts[1], parts[2]
                    else:
                        program = vehicle = serial = ""
                    return program, vehicle, serial
                add_cols: list[str] = []
                for col in ["Program", "Vehicle", "Serial"]:
                    if col not in headers:
                        add_cols.append(col)
                if add_cols:
                    headers = headers + add_cols
                    new_rows: list[list[str]] = []
                    for row in rows:
                        prog, veh, ser = _derive_from_row(row)
                        new_rows.append(row + [prog, veh, ser][: len(add_cols)])
                    rows = new_rows
                else:
                    # Fill existing columns if present but empty
                    idx_prog = hdr_lc.index("program") if "program" in hdr_lc else -1
                    idx_veh = hdr_lc.index("vehicle") if "vehicle" in hdr_lc else -1
                    idx_ser = hdr_lc.index("serial") if "serial" in hdr_lc else -1
                    for i, row in enumerate(rows):
                        prog, veh, ser = _derive_from_row(row)
                        if idx_prog >= 0 and (idx_prog >= len(row) or not str(row[idx_prog]).strip()):
                            if idx_prog >= len(row):
                                row.extend([""] * (idx_prog - len(row) + 1))
                            row[idx_prog] = prog
                        if idx_veh >= 0 and (idx_veh >= len(row) or not str(row[idx_veh]).strip()):
                            if idx_veh >= len(row):
                                row.extend([""] * (idx_veh - len(row) + 1))
                            row[idx_veh] = veh
                        if idx_ser >= 0 and (idx_ser >= len(row) or not str(row[idx_ser]).strip()):
                            if idx_ser >= len(row):
                                row.extend([""] * (idx_ser - len(row) + 1))
                            row[idx_ser] = ser
            except Exception:
                pass

            return headers, rows
        except Exception:
            return [], []

    def _refresh_run_registry(self):
        try:
            self._registry_cache = self._get_registry_table_data()
        except Exception:
            self._registry_cache = ([], [])

    def _show_registry_popup(self):
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
            tbl.setEditTriggers(QtWidgets.QAbstractItemView.EditTrigger.NoEditTriggers)
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
            btns = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.StandardButton.Close)
            btns.rejected.connect(dlg.reject)
            btns.accepted.connect(dlg.accept)
            v.addWidget(btns)
            dlg.exec()
        except Exception as e:
            QtWidgets.QMessageBox.information(self, "Registry", str(e))

    def _safe_open(self, fn):
        try:
            fn()
        except Exception as e:
            QtWidgets.QMessageBox.information(self, "Open", str(e))

    # File/browser helpers
    def _browse_file(self, edit: QtWidgets.QLineEdit, initial_dir: Path, filter_str: str):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Select file", str(initial_dir), filter_str)
        if path:\
        









            self._scan_refresh()

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

        # Upload panel refresh
        self._refresh_upload_list()
        self.drop_zone.set_hint(f"Drop PDFs here\n-> {self.ed_pdfs.text()}")

        # Plotting selection\n        self._load_plot_names_list()
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
    def _load_plot_names_list(self):
        if not hasattr(self, "list_plot_names"):
            return
        try:
            items = be.read_plot_names()
        except Exception:
            items = []
        self.list_plot_names.blockSignals(True)
        self.list_plot_names.clear()
        for name, selected in items:
            it = QtWidgets.QListWidgetItem(name)
            it.setFlags(it.flags() | QtCore.Qt.ItemFlag.ItemIsUserCheckable)
            it.setCheckState(QtCore.Qt.CheckState.Checked if selected else QtCore.Qt.CheckState.Unchecked)
            self.list_plot_names.addItem(it)
        self.list_plot_names.blockSignals(False)

    def _on_plot_name_toggled(self, item: QtWidgets.QListWidgetItem):
        self._apply_plot_name_selection(show_status=False)

    def _apply_plot_name_selection(self, *, show_status: bool = True):
        if not hasattr(self, "list_plot_names"):
            return
        selected: set[str] = set()
        for i in range(self.list_plot_names.count()):
            it = self.list_plot_names.item(i)
            if it and it.checkState() == QtCore.Qt.CheckState.Checked:
                selected.add(it.text())
        try:
            be.set_plot_flags_by_plot_names(selected)
            if show_status and hasattr(self, "status_bar"):
                self.status_bar.showMessage("Plot selection saved", 1500)
        except Exception:
            pass

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
        pix = QtGui.QPixmap(size, size)
        pix.fill(QtCore.Qt.GlobalColor.transparent)
        painter = QtGui.QPainter(pix)
        painter.setRenderHint(QtGui.QPainter.RenderHint.Antialiasing)

        # Background disk
        gradient = QtGui.QConicalGradient(size / 2, size / 2, 45)
        gradient.setColorAt(0.0, QtGui.QColor("#1f5c9a"))
        gradient.setColorAt(0.5, QtGui.QColor("#0f3258"))
        gradient.setColorAt(1.0, QtGui.QColor("#1f5c9a"))
        painter.setBrush(QtGui.QBrush(gradient))
        painter.setPen(QtCore.Qt.PenStyle.NoPen)
        painter.drawEllipse(0, 0, size - 1, size - 1)

        # Inner ring
        ring_color = QtGui.QColor(255, 255, 255, 80)
        pen = QtGui.QPen(ring_color, size * 0.08)
        painter.setPen(pen)
        inset = size * 0.16
        painter.drawEllipse(QtCore.QRectF(inset, inset, size - inset * 2, size - inset * 2))

        # Flight lines
        painter.setPen(QtGui.QPen(QtGui.QColor(255, 255, 255, 140), size * 0.06, QtCore.Qt.PenStyle.SolidLine, QtCore.Qt.PenCapStyle.RoundCap))
        painter.drawArc(int(size * 0.18), int(size * 0.36), int(size * 0.64), int(size * 0.40), 30 * 16, 120 * 16)
        painter.drawArc(int(size * 0.10), int(size * 0.18), int(size * 0.80), int(size * 0.64), -40 * 16, -120 * 16)

        # Center glyph
        painter.setPen(QtGui.QPen(QtGui.QColor("#ffffff")))
        font = painter.font()
        font.setBold(True)
        font.setPointSize(int(size * 0.42))
        painter.setFont(font)
        painter.drawText(pix.rect(), QtCore.Qt.AlignmentFlag.AlignCenter, "E")
        painter.end()
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
