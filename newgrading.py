# grading_system.py  —  run with:  python grading_system.py
# Requirements:  pip install PyQt5 pandas openpyxl numpy
from __future__ import annotations

import sys
import re
import traceback
import os
import json
import datetime
import math
from typing import Any, List, Optional, Dict, Tuple
from pathlib import Path

import numpy as np
import pandas as pd

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QScrollArea, QCheckBox,
    QLineEdit, QTableWidget, QTableWidgetItem, QComboBox,
    QFrame, QHeaderView, QMessageBox, QSizePolicy, QStackedWidget,
    QAbstractItemView, QDialog, QFormLayout, QDoubleSpinBox,
    QDialogButtonBox, QTabWidget, QGridLayout, QToolTip,
    QSpinBox, QGroupBox, QRadioButton, QButtonGroup,
    QSplitter, QInputDialog, QTextEdit, QListWidget, QListWidgetItem,
)
from PyQt5.QtCore import (
    Qt, QPropertyAnimation, QEasingCurve, QSize, QRectF, QPointF,
    pyqtSignal, QTimer,
)
from PyQt5.QtGui import (
    QFont, QColor, QPalette, QPainter, QPainterPath, QPen, QBrush,
    QFontMetrics, QCursor,
)

# ═══════════════════════════════════════════════════════════════════════
#  DESIGN TOKENS
# ═══════════════════════════════════════════════════════════════════════
BG_MAIN   = "#0a0a0a"
BG_PANEL  = "#090909"
BG_CARD   = "#16161a"
BG_INPUT  = "#1c1c21"
BG_HOVER  = "#1f1f25"
BORDER    = "#232329"
BORDER_HI = "#2e2e38"
GOLD      = "#c9a84c"
GOLD_DIM  = "#7a6230"
GOLD_BG   = "#1a1608"
TEXT_PRI  = "#e2e2ec"
TEXT_SEC  = "#70708a"
TEXT_DIM  = "#3e3e52"
WHITE     = "#ffffff"

GRADE_PAL = {
    "A":  ("#0f2b10", "#4caf50"),
    "A-": ("#122b14", "#66bb6a"),
    "B":  ("#0d1c35", "#4a90d9"),
    "B-": ("#0a1e30", "#29b6f6"),
    "C":  ("#2a1f08", "#f5a623"),
    "C-": ("#2a1a05", "#e07b10"),
    "D":  ("#2a1008", "#e06030"),
    "F":  ("#2a0a0a", "#e53935"),
}
GRADE_FG = {g: v[1] for g, v in GRADE_PAL.items()}
GRADE_ORDER = ["A", "A-", "B", "B-", "C", "C-", "D", "F"]

STYLESHEET = f"""
QMainWindow, QWidget {{
    background-color: {BG_MAIN};
    color: {TEXT_PRI};
    font-family: 'Segoe UI', system-ui, sans-serif;
    font-size: 16px;
}}
QScrollArea {{ border: none; background: transparent; }}
QScrollBar:horizontal {{
    background: {BG_MAIN}; height: 5px; border-radius: 3px; margin: 0;
}}
QScrollBar::handle:horizontal {{
    background: {BORDER_HI}; border-radius: 3px; min-width: 20px;
}}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{ width: 0; }}
QScrollBar:vertical {{
    background: {BG_MAIN}; width: 5px; border-radius: 3px; margin: 0;
}}
QScrollBar::handle:vertical {{
    background: {BORDER_HI}; border-radius: 3px; min-height: 20px;
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{ height: 0; }}
QScrollBar:horizontal {{
    background: {BG_MAIN}; height: 12px; border-radius: 6px; margin: 0;
}}
QScrollBar::handle:horizontal {{
    background: {BORDER_HI}; border-radius: 6px; min-width: 28px;
}}
QTableWidget {{
    background-color: {BG_MAIN};
    alternate-background-color: {BG_CARD};
    color: {TEXT_PRI};
    border: none;
    gridline-color: {BORDER};
    font-size: 14px;
    selection-background-color: #1d1d28;
    outline: none;
}}
QTableWidget::item {{
    padding: 0px 12px;
    border-bottom: 1px solid {BORDER};
}}
QTableWidget::item:selected {{
    background-color: #1d1d28;
    color: {WHITE};
}}
QHeaderView::section {{
    background-color: {BG_CARD};
    color: {TEXT_SEC};
    padding: 9px 14px;
    border: none;
    border-bottom: 1px solid {BORDER};
    border-right: 1px solid {BORDER};
    font-size: 13px;
    font-weight: bold;
    letter-spacing: 1.2px;
}}
QLineEdit {{
    background-color: {BG_INPUT};
    color: {TEXT_PRI};
    border: 1px solid {BORDER};
    border-radius: 6px;
    padding: 6px 10px;
    font-size: 16px;
}}
QLineEdit:focus {{ border: 1px solid {GOLD_DIM}; }}
QLineEdit:disabled {{
    background-color: {BG_MAIN};
    color: {TEXT_DIM};
    border: 1px solid #181820;
}}
QComboBox {{
    background-color: {BG_INPUT};
    color: {TEXT_PRI};
    border: 1px solid {BORDER_HI};
    border-radius: 7px;
    padding: 7px 14px;
    font-size: 16px;
    min-height: 38px;
}}
QComboBox:focus {{ border: 1px solid {GOLD_DIM}; }}
QComboBox::drop-down {{ border: none; width: 26px; }}
QComboBox QAbstractItemView {{
    background-color: {BG_CARD};
    color: {TEXT_PRI};
    border: 1px solid {BORDER_HI};
    selection-background-color: {GOLD_DIM};
    selection-color: {WHITE};
    outline: none;
}}
QCheckBox {{ color: {TEXT_PRI}; spacing: 8px; background: transparent; }}
QCheckBox::indicator {{
    width: 19px; height: 19px;
    border-radius: 5px;
    border: 1px solid {BORDER_HI};
    background-color: {BG_INPUT};
}}
QCheckBox::indicator:checked {{
    background-color: {GOLD};
    border: 1px solid {GOLD};
}}
QStatusBar {{
    background-color: {BG_MAIN};
    color: {TEXT_DIM};
    border-top: 1px solid {BORDER};
    font-size: 12px;
    padding: 3px 10px;
}}
QToolTip {{
    background-color: {BG_CARD};
    color: {TEXT_PRI};
    border: 1px solid {BORDER_HI};
    padding: 6px 10px;
    border-radius: 4px;
}}
QTabWidget::pane {{
    border: 1px solid {BORDER};
    border-radius: 4px;
    background-color: {BG_MAIN};
    top: -1px;
}}
QTabBar::tab {{
    background-color: {BG_CARD};
    color: {TEXT_SEC};
    border: 1px solid {BORDER};
    border-bottom: none;
    border-top-left-radius: 6px;
    border-top-right-radius: 6px;
    padding: 8px 18px;
    margin-right: 2px;
    font-weight: 600;
    font-size: 13px;
}}
QTabBar::tab:selected {{
    background-color: {BG_MAIN};
    color: {GOLD};
    border-bottom: 2px solid {GOLD};
}}
QTabBar::tab:hover {{
    color: {TEXT_PRI};
    background-color: {BG_HOVER};
}}
QRadioButton {{
    color: {TEXT_PRI};
    spacing: 8px;
    background: transparent;
    font-size: 14px;
}}
QRadioButton::indicator {{
    width: 18px; height: 18px;
    border-radius: 9px;
    border: 1.5px solid {BORDER_HI};
    background-color: {BG_INPUT};
}}
QRadioButton::indicator:checked {{
    background-color: {GOLD};
    border: 1.5px solid {GOLD};
}}
QGroupBox {{
    background-color: {BG_CARD};
    border: 1px solid {BORDER};
    border-radius: 8px;
    margin-top: 12px;
    padding-top: 18px;
    font-weight: bold;
    color: {TEXT_SEC};
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    left: 14px;
    padding: 0 6px;
    color: {TEXT_SEC};
    font-size: 12px;
}}
QDoubleSpinBox {{
    background-color: {BG_INPUT};
    border: 1px solid {BORDER};
    border-radius: 6px;
    padding: 4px 8px;
    color: {TEXT_PRI};
    font-size: 14px;
}}
QDoubleSpinBox:focus {{
    border: 1px solid {GOLD_DIM};
}}
"""

# ═══════════════════════════════════════════════════════════════════════
#  PERSISTENT STORAGE  (NEW — Feature 1)
# ═══════════════════════════════════════════════════════════════════════

def _get_app_data_dir() -> Path:
    """
    Return a writable directory for app data that works both in dev mode
    and when packaged as a PyInstaller .exe.
    On Windows this is %APPDATA%/AutoGradingSystem; elsewhere ~/.AutoGradingSystem.
    We never rely on sys._MEIPASS (read-only bundle dir) for writing.
    """
    if sys.platform == "win32":
        base = Path(os.environ.get("APPDATA", Path.home()))
    else:
        base = Path.home()
    app_dir = base / "AutoGradingSystem"
    app_dir.mkdir(parents=True, exist_ok=True)
    return app_dir


APP_DATA_DIR   = _get_app_data_dir()
SESSIONS_FILE  = APP_DATA_DIR / "sessions.json"   # index of all past sessions
CURRENT_FILE   = APP_DATA_DIR / "current_session.json"  # auto-save slot
CLUSTERS_FILE  = APP_DATA_DIR / "clusters.json"   # NEW: cluster definitions


def _load_json(path: Path) -> Any:
    """Load JSON safely; return None on missing / corrupt file."""
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return None


def _save_json(path: Path, data: Any) -> bool:
    """Save data as JSON; returns True on success."""
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        return True
    except Exception as exc:
        print(f"[Storage] Could not save {path}: {exc}")
        return False


def _session_to_dict(app: "GradingApp") -> dict:
    """
    Serialize the current application state into a plain dict that
    can be round-tripped through JSON without losing type information.
    """
    if app.df_graded is None:
        return {}

    state: dict = {
        "timestamp": datetime.datetime.now().isoformat(timespec="seconds"),
        "id_col": app._id_col or "",
        "frozen_mu": app._frozen_mu,
        "frozen_sigma": app._frozen_sigma,
        "base_grade": app.combo_base.currentText(),
        "use_custom_cutoffs": app._use_custom_cutoffs,
        "thresholds": app._thresholds,          # list of [grade, value]
        "custom_thresholds": app._custom_thresholds,
        "components": app._components,
        # Store df_raw and df_graded as lists-of-dicts (all strings for safety)
        "df_raw": app.df_raw.to_dict(orient="records") if app.df_raw is not None else [],
        "df_graded": app.df_graded.to_dict(orient="records"),
        "edit_log": app._edit_log,
        "numeric_cols": app._numeric_cols,
        "clusters": app._clusters,  # NEW: persist cluster definitions
    }
    return state


def _session_from_dict(state: dict) -> dict:
    """
    Deserialize and return a plain state dict.  The caller (GradingApp)
    is responsible for applying the state to the UI widgets.
    """
    return state  # already a dict; reconstruction happens in _restore_session


def save_current_session(app: "GradingApp", custom_name: str = "") -> bool:
    """
    Write session data and append an entry to sessions index.
    NEW: supports custom_name for named saves; each session gets its own file.
    """
    if app.df_graded is None:
        return False
    state = _session_to_dict(app)
    if not state:
        return False

    # ── Determine file name for this session ──
    ts_safe = state["timestamp"].replace(":", "-").replace("T", "_")
    if custom_name:
        # Sanitize the custom name for use as a filename
        safe = re.sub(r'[\\/*?:"<>|]', "_", custom_name.strip())
        fname = f"session_{safe}_{ts_safe}.json"
        display_name = custom_name.strip()
    else:
        fname = f"session_{ts_safe}.json"
        display_name = ""

    session_path = APP_DATA_DIR / fname
    state["session_file"] = fname
    state["display_name"] = display_name

    # 1. Save to named session file
    ok = _save_json(session_path, state)

    # 2. Also keep the current_session slot updated (for quick resume)
    _save_json(CURRENT_FILE, state)

    # 3. Append to session index (keep last 50)
    index = _load_json(SESSIONS_FILE) or []
    entry = {
        "timestamp": state["timestamp"],
        "display_name": display_name,
        "session_file": fname,
        "id_col": state["id_col"],
        "student_count": len(state["df_graded"]),
        "components": [c["column"] for c in state["components"]],
    }
    # Remove any old entry with the same session file to avoid duplicates
    index = [e for e in index if e.get("session_file") != fname]
    index.append(entry)
    index = index[-50:]
    _save_json(SESSIONS_FILE, index)

    return ok


def delete_session_by_index(session_index: int) -> bool:
    """
    NEW — Delete a specific session entry (newest-first index).
    Removes the session file and updates the index.
    Returns True on success.
    """
    index = _load_json(SESSIONS_FILE) or []
    reversed_index = list(reversed(index))  # newest first

    if session_index < 0 or session_index >= len(reversed_index):
        return False

    entry = reversed_index[session_index]
    fname = entry.get("session_file", "")

    # Delete the session file if it exists
    if fname:
        session_path = APP_DATA_DIR / fname
        try:
            if session_path.exists():
                session_path.unlink()
        except Exception as exc:
            print(f"[Storage] Could not delete {session_path}: {exc}")

    # Remove from index
    reversed_index.pop(session_index)
    _save_json(SESSIONS_FILE, list(reversed(reversed_index)))
    return True


def load_sessions_index() -> list:
    """Return list of past session summaries (newest first)."""
    data = _load_json(SESSIONS_FILE) or []
    return list(reversed(data))


def load_session_by_index(session_index: int) -> Optional[dict]:
    """
    Load a session by its newest-first index.
    Uses per-session files when available, falls back to current_session.json.
    """
    sessions = load_sessions_index()  # newest first
    if session_index < 0 or session_index >= len(sessions):
        return None
    entry = sessions[session_index]
    fname = entry.get("session_file", "")
    if fname:
        path = APP_DATA_DIR / fname
        data = _load_json(path)
        if data:
            return data
    # Fallback: the old single-file slot
    return _load_json(CURRENT_FILE)


# ═══════════════════════════════════════════════════════════════════════
#  CLUSTER STORAGE  (NEW)
# ═══════════════════════════════════════════════════════════════════════

def load_clusters() -> Dict[str, list]:
    """Return dict of {cluster_name: [roll_number, ...]}."""
    return _load_json(CLUSTERS_FILE) or {}


def save_clusters(clusters: Dict[str, list]) -> bool:
    """Persist cluster definitions."""
    return _save_json(CLUSTERS_FILE, clusters)


# ═══════════════════════════════════════════════════════════════════════
#  MANAGE HISTORY DIALOG  (NEW — Feature: Delete History)
# ═══════════════════════════════════════════════════════════════════════

class ManageHistoryDialog(QDialog):
    """
    Lists all saved sessions; user can select one or more to delete.
    Similar to browser history management.
    """
    def __init__(self, sessions: list, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Manage Session History")
        self.setMinimumWidth(600)
        self.setMinimumHeight(400)
        self.setStyleSheet(f"QDialog {{ background-color: {BG_CARD}; color: {TEXT_PRI}; }}")
        self._sessions = list(sessions)  # newest-first copy
        self._to_delete: set = set()

        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        hdr = QLabel("Select sessions to delete, then click 'Delete Selected'.")
        hdr.setStyleSheet(f"color:{TEXT_SEC}; font-size:13px;")
        layout.addWidget(hdr)

        # Action row
        act_row = QHBoxLayout()
        sel_all = _small_btn("✓ Select All", height=28)
        sel_none = _small_btn("✗ Clear", height=28)
        act_row.addWidget(sel_all)
        act_row.addWidget(sel_none)
        act_row.addStretch()
        layout.addLayout(act_row)

        self._list = QListWidget()
        self._list.setSelectionMode(QAbstractItemView.MultiSelection)
        self._list.setStyleSheet(f"""
            QListWidget {{
                background-color: {BG_MAIN}; color: {TEXT_PRI};
                border: 1px solid {BORDER}; font-size: 13px;
            }}
            QListWidget::item:selected {{
                background-color: #2a0a0a; color: #e53935;
            }}
        """)
        for s in self._sessions:
            label = s.get("display_name") or s.get("timestamp", "—")
            extra = f"  ·  {s.get('student_count', '?')} students"
            comp_str = ", ".join(s.get("components", []))
            item = QListWidgetItem(f"📁  {label}{extra}  [{comp_str}]")
            self._list.addItem(item)
        layout.addWidget(self._list, 1)

        sel_all.clicked.connect(self._list.selectAll)
        sel_none.clicked.connect(self._list.clearSelection)

        btn_row = QHBoxLayout()
        del_btn = _primary_btn("🗑  Delete Selected", height=42)
        del_btn.setStyleSheet(del_btn.styleSheet().replace(GOLD, "#c0392b").replace("#d4b55a", "#e74c3c").replace("#b89038", "#922b21"))
        del_btn.clicked.connect(self._on_delete)
        btn_row.addWidget(del_btn)

        close_btn = _ghost_btn("Close", height=42)
        close_btn.clicked.connect(self.accept)
        btn_row.addWidget(close_btn)
        layout.addLayout(btn_row)

    def _on_delete(self):
        selected_rows = sorted(
            [self._list.row(item) for item in self._list.selectedItems()],
            reverse=True  # delete from highest index to preserve lower indices
        )
        if not selected_rows:
            QMessageBox.information(self, "None Selected", "Please select sessions to delete.")
            return

        names = [self._sessions[r].get("display_name") or self._sessions[r].get("timestamp", "?")
                 for r in selected_rows]
        confirm = QMessageBox.question(
            self, "Confirm Delete",
            f"Permanently delete {len(selected_rows)} session(s)?\n\n" + "\n".join(f"• {n}" for n in names),
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if confirm != QMessageBox.Yes:
            return

        for row in selected_rows:
            delete_session_by_index(row)
            self._list.takeItem(row)
            self._sessions.pop(row)

        QMessageBox.information(self, "Deleted", f"{len(selected_rows)} session(s) deleted.")


# ═══════════════════════════════════════════════════════════════════════
#  CLUSTER MANAGER DIALOG  (NEW — Feature: Cluster)
# ═══════════════════════════════════════════════════════════════════════

class ClusterManagerDialog(QDialog):
    """
    Create and manage clusters (groups of roll numbers).
    Also allows bulk mark assignment to a cluster.
    """
    def __init__(self, clusters: Dict[str, list], all_student_ids: list,
                 components: list, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Cluster Manager")
        self.setMinimumWidth(560)
        self.setMinimumHeight(500)
        self.setStyleSheet(f"QDialog {{ background-color: {BG_CARD}; color: {TEXT_PRI}; }}")

        self._clusters: Dict[str, list] = {k: list(v) for k, v in clusters.items()}
        self._all_ids = all_student_ids
        self._components = components

        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        # ── Header
        hdr = QLabel("🗂  Cluster Manager — group students by roll number")
        hdr.setStyleSheet(f"color:{GOLD}; font-size:14px; font-weight:bold;")
        layout.addWidget(hdr)
        layout.addWidget(_hline())

        # ── Create new cluster
        new_grp = QGroupBox("Create / Edit Cluster")
        new_grp.setStyleSheet(f"""
            QGroupBox {{ background-color:{BG_MAIN}; border:1px solid {BORDER};
                         border-radius:8px; margin-top:10px; padding-top:14px; }}
            QGroupBox::title {{ color:{TEXT_SEC}; font-size:11px; left:12px; }}
        """)
        new_lay = QVBoxLayout(new_grp)

        row1 = QHBoxLayout()
        row1.addWidget(QLabel("Cluster Name:"))
        self._name_edit = QLineEdit()
        self._name_edit.setPlaceholderText("e.g. Group A")
        row1.addWidget(self._name_edit, 1)
        new_lay.addLayout(row1)

        new_lay.addWidget(QLabel("Roll Numbers (one per line or comma-separated):"))
        self._rolls_edit = QTextEdit()
        self._rolls_edit.setFixedHeight(90)
        self._rolls_edit.setStyleSheet(f"""
            background:{BG_INPUT}; color:{TEXT_PRI};
            border:1px solid {BORDER}; border-radius:6px; font-size:13px;
        """)
        new_lay.addWidget(self._rolls_edit)

        add_btn = _primary_btn("➕  Save Cluster", height=38)
        add_btn.clicked.connect(self._on_save_cluster)
        new_lay.addWidget(add_btn)
        layout.addWidget(new_grp)

        # ── Existing clusters
        layout.addWidget(_lbl("EXISTING CLUSTERS", size=11))
        self._cluster_list = QListWidget()
        self._cluster_list.setMaximumHeight(120)
        self._cluster_list.setStyleSheet(f"""
            QListWidget {{ background:{BG_MAIN}; color:{TEXT_PRI};
                           border:1px solid {BORDER}; font-size:13px; }}
            QListWidget::item:selected {{ background:{BG_HOVER}; }}
        """)
        self._cluster_list.currentItemChanged.connect(self._on_cluster_selected)
        self._refresh_cluster_list()
        layout.addWidget(self._cluster_list)

        cl_btn_row = QHBoxLayout()
        del_cl = _small_btn("🗑 Delete Cluster", height=30)
        del_cl.clicked.connect(self._on_delete_cluster)
        cl_btn_row.addWidget(del_cl)
        cl_btn_row.addStretch()
        layout.addLayout(cl_btn_row)

        layout.addWidget(_hline())
        close_btn = _ghost_btn("Close", height=42)
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)

    def _refresh_cluster_list(self):
        self._cluster_list.clear()
        for name, rolls in self._clusters.items():
            item = QListWidgetItem(f"📦 {name}  ({len(rolls)} students): {', '.join(rolls[:5])}{'...' if len(rolls) > 5 else ''}")
            item.setData(Qt.UserRole, name)
            self._cluster_list.addItem(item)

    def _on_cluster_selected(self, current, _prev):
        if not current:
            return
        name = current.data(Qt.UserRole)
        rolls = self._clusters.get(name, [])
        self._name_edit.setText(name)
        self._rolls_edit.setPlainText("\n".join(rolls))

    def _on_save_cluster(self):
        name = self._name_edit.text().strip()
        if not name:
            QMessageBox.warning(self, "Error", "Please enter a cluster name.")
            return
        raw = self._rolls_edit.toPlainText().strip()
        # Parse comma or newline separated roll numbers
        rolls = [r.strip() for part in raw.replace(",", "\n").split("\n") for r in [part.strip()] if r]
        if not rolls:
            QMessageBox.warning(self, "Error", "Please enter at least one roll number.")
            return
        self._clusters[name] = rolls
        save_clusters(self._clusters)
        self._refresh_cluster_list()
        QMessageBox.information(self, "Saved", f"Cluster '{name}' saved with {len(rolls)} student(s).")

    def _on_delete_cluster(self):
        item = self._cluster_list.currentItem()
        if not item:
            QMessageBox.information(self, "None Selected", "Select a cluster to delete.")
            return
        name = item.data(Qt.UserRole)
        if QMessageBox.question(self, "Delete", f"Delete cluster '{name}'?",
                                QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
            self._clusters.pop(name, None)
            save_clusters(self._clusters)
            self._refresh_cluster_list()

    def get_clusters(self) -> Dict[str, list]:
        return self._clusters


# ═══════════════════════════════════════════════════════════════════════
#  CLUSTER BULK MARKS DIALOG  (NEW — Feature: Cluster bulk assign)
# ═══════════════════════════════════════════════════════════════════════

class ClusterBulkMarksDialog(QDialog):
    """
    Assign marks for a selected component to all students in a cluster.
    """
    def __init__(self, clusters: Dict[str, list], components: list, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Assign Marks to Cluster")
        self.setMinimumWidth(420)
        self.setStyleSheet(f"QDialog {{ background-color:{BG_CARD}; color:{TEXT_PRI}; }}")

        self._result: Optional[Dict] = None  # {cluster_name, component, mark}

        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        hdr = QLabel("Bulk-assign marks to all students in a cluster")
        hdr.setStyleSheet(f"color:{TEXT_SEC}; font-size:13px;")
        hdr.setWordWrap(True)
        layout.addWidget(hdr)

        form = QFormLayout()
        form.setSpacing(10)

        self._cluster_combo = QComboBox()
        for name, rolls in clusters.items():
            self._cluster_combo.addItem(f"{name}  ({len(rolls)} students)", userData=name)
        form.addRow("Cluster:", self._cluster_combo)

        self._comp_combo = QComboBox()
        for comp in components:
            self._comp_combo.addItem(f"{comp['column']}  (max {comp['max_marks']})", userData=comp)
        form.addRow("Component:", self._comp_combo)

        self._mark_spin = QDoubleSpinBox()
        self._mark_spin.setRange(0, 9999)
        self._mark_spin.setDecimals(2)
        self._mark_spin.setSingleStep(1.0)
        form.addRow("Mark to assign:", self._mark_spin)

        layout.addLayout(form)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.setStyleSheet(f"""
            QPushButton {{ background-color:{GOLD}; color:#060606;
                border:none; border-radius:6px; padding:8px 22px;
                font-weight:bold; font-size:14px; }}
            QPushButton:hover {{ background-color:#d4b55a; }}
        """)
        buttons.accepted.connect(self._on_ok)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _on_ok(self):
        self._result = {
            "cluster_name": self._cluster_combo.currentData(),
            "component": self._comp_combo.currentData(),
            "mark": self._mark_spin.value(),
        }
        self.accept()

    def get_result(self) -> Optional[Dict]:
        return self._result


# ═══════════════════════════════════════════════════════════════════════
#  DYNAMIC COMPONENT DIALOG  (NEW — Feature: Dynamic Component Creation)
# ═══════════════════════════════════════════════════════════════════════

class AddComponentDialog(QDialog):
    """
    Let user define a brand-new component column (e.g. 'Presentation')
    with max marks and weightage. Marks for each student start at 0.
    """
    def __init__(self, existing_columns: list, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Add New Component")
        self.setMinimumWidth(400)
        self.setStyleSheet(f"QDialog {{ background-color:{BG_CARD}; color:{TEXT_PRI}; }}")

        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        hdr = QLabel("Define a new grading component (e.g. Presentation, Quiz 3)")
        hdr.setStyleSheet(f"color:{TEXT_SEC}; font-size:13px;")
        hdr.setWordWrap(True)
        layout.addWidget(hdr)

        form = QFormLayout()
        form.setSpacing(10)

        self._name_edit = QLineEdit()
        self._name_edit.setPlaceholderText("e.g. Presentation")
        form.addRow("Component Name:", self._name_edit)

        self._max_spin = QDoubleSpinBox()
        self._max_spin.setRange(1, 9999)
        self._max_spin.setDecimals(1)
        self._max_spin.setValue(100)
        form.addRow("Max Marks:", self._max_spin)

        self._weight_spin = QDoubleSpinBox()
        self._weight_spin.setRange(0.1, 100)
        self._weight_spin.setDecimals(1)
        self._weight_spin.setValue(10)
        form.addRow("Weightage:", self._weight_spin)

        layout.addLayout(form)

        self._existing = existing_columns

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.setStyleSheet(f"""
            QPushButton {{ background-color:{GOLD}; color:#060606;
                border:none; border-radius:6px; padding:8px 22px;
                font-weight:bold; font-size:14px; }}
            QPushButton:hover {{ background-color:#d4b55a; }}
        """)
        buttons.accepted.connect(self._on_ok)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _on_ok(self):
        name = self._name_edit.text().strip()
        if not name:
            QMessageBox.warning(self, "Error", "Please enter a component name.")
            return
        if name in self._existing:
            QMessageBox.warning(self, "Duplicate", f"Column '{name}' already exists.")
            return
        self.accept()

    def get_component_info(self) -> dict:
        return {
            "name": self._name_edit.text().strip(),
            "max_marks": self._max_spin.value(),
            "weightage": self._weight_spin.value(),
        }



class ColumnSelectDialog(QDialog):
    """
    Shown before Excel export.  User checks which columns to include.
    All columns are checked by default.
    """
    def __init__(self, columns: list, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Select Columns to Export")
        self.setMinimumWidth(380)
        self.setStyleSheet(f"""
            QDialog {{
                background-color: {BG_CARD};
                color: {TEXT_PRI};
            }}
        """)
        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        header = QLabel("Choose which columns to include in the Excel export:")
        header.setStyleSheet(f"color:{TEXT_SEC}; font-size:13px;")
        header.setWordWrap(True)
        layout.addWidget(header)

        # Select All / None shortcuts
        btn_row = QHBoxLayout()
        all_btn = _small_btn("✓ Select All",  height=28)
        non_btn = _small_btn("✗ Clear All",   height=28)
        btn_row.addWidget(all_btn)
        btn_row.addWidget(non_btn)
        btn_row.addStretch()
        layout.addLayout(btn_row)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("background:transparent; border:none;")
        inner = QWidget()
        inner.setStyleSheet("background:transparent;")
        inner_lay = QVBoxLayout(inner)
        inner_lay.setContentsMargins(4, 4, 4, 4)
        inner_lay.setSpacing(6)

        self._checks: Dict[str, QCheckBox] = {}
        for col in columns:
            cb = QCheckBox(col)
            cb.setChecked(True)
            self._checks[col] = cb
            inner_lay.addWidget(cb)

        inner_lay.addStretch()
        scroll.setWidget(inner)
        scroll.setMinimumHeight(min(300, len(columns) * 32 + 20))
        layout.addWidget(scroll)

        # Wire select-all / clear-all
        all_btn.clicked.connect(lambda: [c.setChecked(True)  for c in self._checks.values()])
        non_btn.clicked.connect(lambda: [c.setChecked(False) for c in self._checks.values()])

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.setStyleSheet(f"""
            QPushButton {{
                background-color: {GOLD}; color: #060606;
                border: none; border-radius: 6px;
                padding: 8px 22px; font-weight: bold; font-size: 14px;
            }}
            QPushButton:hover {{ background-color: #d4b55a; }}
        """)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_selected_columns(self) -> list:
        """Return ordered list of checked column names."""
        return [col for col, cb in self._checks.items() if cb.isChecked()]


# ═══════════════════════════════════════════════════════════════════════
#  RESUME SESSION DIALOG  (NEW — Feature 1)
# ═══════════════════════════════════════════════════════════════════════

class ResumeSessionDialog(QDialog):
    """
    Shows the last N sessions and lets the user pick one to restore.
    """
    def __init__(self, sessions: list, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Resume Previous Session")
        self.setMinimumWidth(520)
        self.setStyleSheet(f"QDialog {{ background-color: {BG_CARD}; color: {TEXT_PRI}; }}")
        self._selected_index = None

        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        title = QLabel("Select a session to restore:")
        title.setStyleSheet(f"color:{TEXT_SEC}; font-size:13px;")
        layout.addWidget(title)

        self._list = QTableWidget()
        self._list.setColumnCount(5)
        self._list.setHorizontalHeaderLabels(["NAME", "TIMESTAMP", "ID COLUMN", "STUDENTS", "COMPONENTS"])
        self._list.setRowCount(len(sessions))
        self._list.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._list.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._list.setSelectionMode(QAbstractItemView.SingleSelection)
        self._list.verticalHeader().setVisible(False)
        self._list.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self._list.setShowGrid(True)
        self._list.setFrameShape(QFrame.NoFrame)
        self._list.setAlternatingRowColors(True)
        self._list.setStyleSheet(f"""
            QTableWidget {{
                background-color: {BG_MAIN}; color: {TEXT_PRI};
                border: none; gridline-color: {BORDER};
                font-size: 13px;
            }}
            QHeaderView::section {{
                background-color: {BG_CARD}; color: {TEXT_SEC};
                padding: 6px 10px; border: none;
                border-bottom: 1px solid {BORDER};
            }}
        """)

        for r, s in enumerate(sessions):
            self._list.setRowHeight(r, 32)
            # Col 0: display name (or timestamp if no name)
            name_text = s.get("display_name") or s.get("timestamp", "—")
            name_item = QTableWidgetItem(name_text)
            name_item.setForeground(QColor(GOLD))
            self._list.setItem(r, 0, name_item)
            # Col 1: timestamp
            self._list.setItem(r, 1, QTableWidgetItem(s.get("timestamp", "—")))
            # Col 2: id column
            self._list.setItem(r, 2, QTableWidgetItem(s.get("id_col", "—")))
            # Col 3: student count
            self._list.setItem(r, 3, QTableWidgetItem(str(s.get("student_count", "—"))))
            # Col 4: components
            comps = ", ".join(s.get("components", []))
            self._list.setItem(r, 4, QTableWidgetItem(comps))

        self._list.resizeColumnsToContents()
        layout.addWidget(self._list)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.setStyleSheet(f"""
            QPushButton {{
                background-color: {GOLD}; color: #060606;
                border: none; border-radius: 6px;
                padding: 8px 22px; font-weight: bold; font-size: 14px;
            }}
            QPushButton:hover {{ background-color: #d4b55a; }}
        """)
        buttons.accepted.connect(self._on_ok)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _on_ok(self):
        rows = self._list.selectedItems()
        if rows:
            self._selected_index = self._list.currentRow()
        self.accept()

    def get_selected_index(self) -> Optional[int]:
        """0-based index into the sessions list (newest first)."""
        return self._selected_index


# ═══════════════════════════════════════════════════════════════════════
#  WIDGET HELPERS
# ═══════════════════════════════════════════════════════════════════════

def _hline(color=BORDER):
    f = QFrame()
    f.setFrameShape(QFrame.HLine)
    f.setFixedHeight(1)
    f.setStyleSheet(f"background:{color};border:none;")
    return f


def _lbl(text, color=TEXT_SEC, size=17, bold=False, spacing="1.5px"):
    l = QLabel(text)
    w = "bold" if bold else "normal"
    l.setStyleSheet(
        f"color:{color};font-size:{size}px;font-weight:{w};"
        f"letter-spacing:{spacing};background:transparent;"
    )
    return l


def _primary_btn(text, height=50):
    b = QPushButton(text)
    b.setFixedHeight(height)
    b.setCursor(Qt.PointingHandCursor)
    b.setStyleSheet(f"""
        QPushButton {{
            background-color: {GOLD}; color: #060606;
            border: none; border-radius: 9px;
            font-size: 17px; font-weight: bold;
        }}
        QPushButton:hover   {{ background-color: #d4b55a; }}
        QPushButton:pressed {{ background-color: #b89038; }}
        QPushButton:disabled {{
            background-color: #1e1c10; color: {TEXT_DIM};
        }}
    """)
    return b


def _ghost_btn(text, height=46):
    b = QPushButton(text)
    b.setFixedHeight(height)
    b.setCursor(Qt.PointingHandCursor)
    b.setStyleSheet(f"""
        QPushButton {{
            background-color: {BG_CARD}; color: {TEXT_SEC};
            border: 1px solid {BORDER}; border-radius: 9px;
            font-size: 17px; font-weight: bold;
        }}
        QPushButton:hover {{
            background-color: {BG_HOVER};
            color: {TEXT_PRI}; border-color: {BORDER_HI};
        }}
        QPushButton:pressed {{ background-color: {BG_INPUT}; }}
        QPushButton:disabled {{
            background-color: {BG_MAIN};
            color: {TEXT_DIM}; border-color: {BORDER};
        }}
    """)
    return b


def _small_btn(text, height=34):
    b = QPushButton(text)
    b.setFixedHeight(height)
    b.setCursor(Qt.PointingHandCursor)
    b.setStyleSheet(f"""
        QPushButton {{
            background-color: {BG_CARD}; color: {TEXT_SEC};
            border: 1px solid {BORDER}; border-radius: 6px;
            font-size: 13px; font-weight: bold;
            padding: 4px 12px;
        }}
        QPushButton:hover {{
            background-color: {BG_HOVER};
            color: {TEXT_PRI}; border-color: {GOLD_DIM};
        }}
        QPushButton:pressed {{ background-color: {BG_INPUT}; }}
        QPushButton:disabled {{
            background-color: {BG_MAIN};
            color: {TEXT_DIM}; border-color: {BORDER};
        }}
    """)
    return b


# ═══════════════════════════════════════════════════════════════════════
#  DATA LOGIC
# ═══════════════════════════════════════════════════════════════════════

def is_valid_student_id(val: Any) -> bool:
    s = str(val).strip()
    return bool(re.search(r'[A-Za-z]', s)) and bool(re.search(r'\d', s))


def load_and_clean_data(filepath: str) -> pd.DataFrame:
    ext = filepath.rsplit('.', 1)[-1].lower()
    if ext == 'csv':
        df = pd.read_csv(filepath, dtype=str)
    elif ext in ('xlsx', 'xls'):
        df = pd.read_excel(filepath, dtype=str)
    else:
        raise ValueError("Unsupported format — use .csv, .xlsx or .xls")
    df.columns = [str(c).strip() for c in df.columns]
    df.fillna('0', inplace=True)
    mask = df[df.columns[0]].apply(is_valid_student_id)
    return df[mask].reset_index(drop=True)


def load_table_file(filepath: str) -> pd.DataFrame:
    """Load a CSV/Excel file without student filtering."""
    ext = filepath.rsplit('.', 1)[-1].lower()
    if ext == 'csv':
        df = pd.read_csv(filepath, dtype=str)
    elif ext in ('xlsx', 'xls'):
        df = pd.read_excel(filepath, dtype=str)
    else:
        raise ValueError("Unsupported format - use .csv, .xlsx or .xls")
    df.columns = [str(c).strip() for c in df.columns]
    return df.fillna('')


def extract_roll_numbers(value: Any, allow_numeric_only: bool = False) -> list:
    """
    Extract roll numbers from a cell. Normal auto-detection requires letters
    and digits; numeric-only IDs are accepted only after user column selection.
    """
    s = str(value).strip()
    if not s:
        return []
    if allow_numeric_only:
        pattern = r'[A-Za-z]*\d+[A-Za-z0-9]*'
    else:
        pattern = r'(?=[A-Za-z0-9]*[A-Za-z])(?=[A-Za-z0-9]*\d)[A-Za-z0-9]+'
    return [m.group(0).strip().upper() for m in re.finditer(pattern, s) if m.group(0).strip()]


def detect_team_columns(df: pd.DataFrame) -> tuple:
    cols = list(df.columns)
    lower = {c: c.lower() for c in cols}

    member_cols = [c for c in cols if "member" in lower[c]]
    team_col = next(
        (c for c in cols if (
            any(k in lower[c] for k in ("team", "group", "cluster")) or lower[c] in ("name", "id")
        ) and c not in member_cols),
        None,
    )
    inferred_member_cols = []
    numeric_only_candidates = []

    if not member_cols:
        for col in cols:
            hits = 0
            numeric_hits = 0
            for val in df[col].dropna().astype(str).head(25):
                if extract_roll_numbers(val, allow_numeric_only=False):
                    hits += 1
                elif extract_roll_numbers(val, allow_numeric_only=True):
                    numeric_hits += 1
            if hits:
                inferred_member_cols.append(col)
            elif numeric_hits:
                numeric_only_candidates.append(col)
        member_cols = inferred_member_cols

    marks_cols = []
    for col in cols:
        if col == team_col or col in member_cols:
            continue
        lc = lower[col]
        numeric_count = pd.to_numeric(df[col], errors='coerce').notna().sum()
        if numeric_count > 0 and any(k in lc for k in ("mark", "marks", "score", "points")):
            marks_cols.append(col)
    if not marks_cols:
        for col in cols:
            if col == team_col or col in member_cols:
                continue
            if pd.to_numeric(df[col], errors='coerce').notna().sum() > 0:
                marks_cols.append(col)

    return team_col, member_cols, marks_cols, numeric_only_candidates


def get_numeric_columns(df: pd.DataFrame) -> list:
    # BUG FIX: exclude computed output columns that may exist in the source file
    exclude = {'Total', 'Grade', 'Rank'}
    return [c for c in df.columns[1:]
            if c not in exclude and pd.to_numeric(df[c], errors='coerce').notna().sum() > 0]


def calculate_totals(df: pd.DataFrame, components: list) -> pd.DataFrame:
    """
    BUG FIX: Drop any pre-existing 'Total' column from the source data before
    computing a fresh weighted total, so the original column is never silently
    overwritten with stale or partial data.
    """
    result = df.copy()
    # Drop the old Total column if it already exists (e.g. column named 'Total' in source)
    if 'Total' in result.columns:
        result = result.drop(columns=['Total'])
    result['Total'] = 0.0
    for comp in components:
        marks = pd.to_numeric(result[comp['column']], errors='coerce').fillna(0)
        result['Total'] += (marks / comp['max_marks']) * comp['weightage']
    result['Total'] = result['Total'].round(2)
    return result


def compute_default_thresholds(mu: float, sigma: float, base_grade: str) -> list:
    if base_grade == 'B':
        thresholds = [
            ('A',  mu + 1.0 * sigma), ('A-', mu + 0.5 * sigma), ('B',  mu),
            ('B-', mu - 0.5 * sigma), ('C',  mu - 1.0 * sigma), ('C-', mu - 1.5 * sigma),
            ('D',  mu - 2.0 * sigma),
        ]
    else:
        thresholds = [
            ('A',  mu + 1.5 * sigma), ('A-', mu + 1.0 * sigma), ('B',  mu + 0.5 * sigma),
            ('B-', mu),               ('C',  mu - 0.5 * sigma),  ('C-', mu - 1.0 * sigma),
            ('D',  mu - 1.5 * sigma),
        ]
    return thresholds


def compute_cutoff_ranges(thresholds: list) -> Dict[str, Tuple[float, float]]:
    cutoffs = {}
    for i, (grade, val) in enumerate(thresholds):
        upper = float('inf') if i == 0 else thresholds[i - 1][1]
        cutoffs[grade] = (val, upper)
    last_val = thresholds[-1][1] if thresholds else 0
    cutoffs['F'] = (float('-inf'), last_val)
    return cutoffs


def assign_grades(df: pd.DataFrame, base_grade: str) -> tuple:
    totals = df['Total'].astype(float)
    mu, sigma = totals.mean(), totals.std()
    if sigma == 0:
        sigma = 0.001
    thresholds = compute_default_thresholds(mu, sigma, base_grade)

    def _g(s):
        for lbl, cut in thresholds:
            if s >= cut:
                return lbl
        return 'F'

    result = df.copy()
    result['Grade'] = result['Total'].astype(float).apply(_g)
    return result, mu, sigma, thresholds


def assign_grades_custom(df: pd.DataFrame, thresholds: list) -> pd.DataFrame:
    def _g(s):
        for lbl, cut in thresholds:
            if s >= cut:
                return lbl
        return 'F'

    result = df.copy()
    result['Grade'] = result['Total'].astype(float).apply(_g)
    return result


def grade_single_student(total: float, thresholds: list) -> str:
    for lbl, cut in thresholds:
        if total >= cut:
            return lbl
    return 'F'


# ═══════════════════════════════════════════════════════════════════════
#  STYLED EXCEL EXPORT
# ═══════════════════════════════════════════════════════════════════════

def save_styled_excel(df: pd.DataFrame, output_path: str, id_col: str,
                      cols_to_export: Optional[list] = None):
    """
    Write df to a styled .xlsx file.

    Feature 2 fix: numeric columns (marks, Total, Rank) are now written
    as Python int/float values so Excel treats them as numbers, enabling
    sorting, SUM(), AVERAGE() etc.

    Feature 3: pass cols_to_export to control which columns appear in
    the output.  Defaults to all columns when not supplied.
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "Grading Results"

    # Use caller-supplied column list or fall back to all columns
    if cols_to_export is None:
        cols_to_export = list(df.columns)
    else:
        # Keep only columns that actually exist in df
        cols_to_export = [c for c in cols_to_export if c in df.columns]

    header_fill = PatternFill(start_color="16161a", end_color="16161a", fill_type="solid")
    header_font = Font(bold=True, color="70708a", size=9)
    header_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        bottom=Side(style="thin", color="232329"),
        left=Side(style="thin", color="232329"),
        right=Side(style="thin", color="232329"),
        top=Side(style="thin", color="232329"),
    )

    for ci, col_name in enumerate(cols_to_export, 1):
        cell = ws.cell(row=1, column=ci, value=col_name)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_align
        cell.border = thin_border

    alt_fills = [
        PatternFill(start_color="16161a", end_color="16161a", fill_type="solid"),
        PatternFill(start_color="0c0c0f", end_color="0c0c0f", fill_type="solid"),
    ]
    gold_font = Font(bold=True, color="c9a84c")
    white_font = Font(color="e2e2ec")

    GRADE_BG_XL = {
        "A": "0f2b10", "A-": "122b14", "B": "0d1c35", "B-": "0a1e30",
        "C": "2a1f08", "C-": "2a1a05", "D": "2a1008", "F": "2a0a0a"
    }
    GRADE_FG_XL = {
        "A": "4caf50", "A-": "66bb6a", "B": "4a90d9", "B-": "29b6f6",
        "C": "f5a623", "C-": "e07b10", "D": "e06030", "F": "e53935"
    }

    for row_num, (_, row) in enumerate(df.iterrows()):
        fill = alt_fills[row_num % 2]
        grade = str(row.get('Grade', 'F'))

        for ci, col_name in enumerate(cols_to_export, 1):
            val = row[col_name]

            # ── Feature 2: store numbers as proper Python numeric types ──
            # This makes Excel treat them as numbers (sortable, formulae work).
            if col_name == 'Grade' or col_name == id_col:
                # Keep grade and ID column as strings
                val = str(val)
            else:
                # Attempt numeric conversion for every other column
                try:
                    numeric = float(val)
                    if col_name == 'Rank':
                        val = int(numeric)
                    elif col_name == 'Total':
                        val = round(numeric, 2)
                    else:
                        # Raw mark columns — preserve decimals but remove
                        # trailing ".0" for whole numbers to keep it clean
                        val = int(numeric) if numeric == int(numeric) else round(numeric, 4)
                except (ValueError, TypeError):
                    val = str(val)  # fallback to string if not parseable

            cell = ws.cell(row=row_num + 2, column=ci, value=val)
            cell.border = thin_border

            if col_name == 'Grade':
                gb = GRADE_BG_XL.get(grade, "16161a")
                gf = GRADE_FG_XL.get(grade, "e2e2ec")
                cell.fill = PatternFill(start_color=gb, end_color=gb, fill_type="solid")
                cell.font = Font(bold=True, color=gf, size=11)
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_name == 'Total':
                cell.fill = fill
                cell.font = gold_font
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_name == 'Rank':
                cell.fill = fill
                rank_f = Font(bold=True, color="c9a84c") if row_num < 3 else white_font
                cell.font = rank_f
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_name == id_col:
                cell.fill = fill
                cell.font = white_font
                cell.alignment = Alignment(horizontal="left", vertical="center")
            else:
                cell.fill = fill
                cell.font = white_font
                cell.alignment = Alignment(horizontal="center", vertical="center")

    for ci in range(1, len(cols_to_export) + 1):
        col_letter = get_column_letter(ci)
        ws.column_dimensions[col_letter].width = 16
    # Make the first column (whatever it is) a bit wider
    ws.column_dimensions[get_column_letter(1)].width = 28
    ws.freeze_panes = "A2"
    wb.save(output_path)


# ═══════════════════════════════════════════════════════════════════════
#  INTERACTIVE GRAPH WIDGET
# ═══════════════════════════════════════════════════════════════════════
class InteractiveGraphWidget(QWidget):
    student_clicked = pyqtSignal(int)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumHeight(380)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self._data = []
        self._cutoffs = {}
        self._thresholds = []
        self._blocks = []
        self._hovered = -1
        self._min_t = 0
        self._max_t = 100
        self._margin_l = 65
        self._margin_r = 25
        self._margin_t = 30
        self._margin_b = 55
        self.setMouseTracking(True)

    def set_data(self, data_list, cutoffs, thresholds=None):
        self._data = data_list
        self._cutoffs = cutoffs
        self._thresholds = thresholds or []
        self._blocks = []
        self._hovered = -1
        self.update()

    def clear_data(self):
        self._data = []
        self._cutoffs = {}
        self._thresholds = []
        self._blocks = []
        self._hovered = -1
        self.update()

    def _compute_layout(self):
        self._blocks = []
        if not self._data:
            return

        ml, mr, mt, mb = self._margin_l, self._margin_r, self._margin_t, self._margin_b
        w = self.width() - ml - mr
        h = self.height() - mt - mb

        if w <= 0 or h <= 0:
            return

        totals = [d[0] for d in self._data]
        self._min_t = min(totals) - 1 if totals else 0
        self._max_t = max(totals) + 1 if totals else 100

        if self._max_t <= self._min_t:
            self._max_t = self._min_t + 1

        min_t, max_t = self._min_t, self._max_t
        data_range = max_t - min_t

        # Group students by EXACT rounded mark (1 decimal)
        bins = {}
        for total, grade, id_str, row_idx in self._data:
            key = round(total, 1)
            if key not in bins:
                bins[key] = []
            bins[key].append((total, grade, id_str, row_idx))

        max_stack = max((len(v) for v in bins.values()), default=1)

        num_possible = int(max_t - min_t) + 1
        block_w = max(5, min(24, (w / max(num_possible, 1)) * 0.8))
        block_h = max(5, min(20, (h / max(max_stack, 1)) * 0.85))

        for mark_val, students in bins.items():
            if data_range > 0:
                cx = ml + ((mark_val - min_t) / data_range) * w
            else:
                cx = ml + w / 2

            for stack_i, (total, grade, id_str, row_idx) in enumerate(students):
                bx = cx - block_w / 2
                by = self.height() - mb - (stack_i + 1) * block_h
                by = max(mt, by)
                rect = QRectF(bx, by, block_w, block_h - 1)
                self._blocks.append((rect, row_idx, id_str, grade, total))

    def paintEvent(self, event):
        self._compute_layout()
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)

        w = self.width()
        h = self.height()
        ml, mr, mt, mb = self._margin_l, self._margin_r, self._margin_t, self._margin_b
        chart_w = w - ml - mr
        chart_h = h - mt - mb

        painter.fillRect(self.rect(), QColor(BG_MAIN))

        if not self._data or chart_w <= 0 or chart_h <= 0:
            painter.setPen(QColor(TEXT_DIM))
            painter.setFont(QFont("Segoe UI", 13))
            painter.drawText(self.rect(), Qt.AlignCenter, "Process data to see the graph")
            painter.end()
            return

        min_t, max_t = self._min_t, self._max_t

        chart_rect = QRectF(ml, mt, chart_w, chart_h)
        painter.fillRect(chart_rect, QColor(BG_CARD))

        painter.setPen(QPen(QColor(BORDER), 0.5))
        num_grid = 8
        for i in range(1, num_grid):
            y = mt + chart_h * i / num_grid
            painter.drawLine(int(ml), int(y), int(w - mr), int(y))

        painter.setPen(QPen(QColor(BORDER_HI), 1.2))
        painter.drawLine(ml, mt, ml, h - mb)
        painter.drawLine(ml, h - mb, w - mr, h - mb)

        painter.setPen(QColor(TEXT_SEC))
        font_small = QFont("Segoe UI", 9)
        painter.setFont(font_small)

        num_ticks = min(10, max(4, int(max_t - min_t)))
        for i in range(num_ticks + 1):
            val = min_t + (max_t - min_t) * i / max(num_ticks, 1)
            x = ml + (val - min_t) / (max_t - min_t) * chart_w
            painter.setPen(QColor(TEXT_SEC))
            painter.drawText(QRectF(x - 25, h - mb + 6, 50, 18), Qt.AlignCenter, f"{val:.1f}")
            painter.setPen(QPen(QColor(BORDER_HI), 0.5))
            painter.drawLine(int(x), int(h - mb), int(x), int(h - mb + 4))

        painter.setPen(QColor(TEXT_SEC))
        painter.setFont(QFont("Segoe UI", 10, QFont.Bold))
        painter.drawText(QRectF(ml, h - 18, chart_w, 18), Qt.AlignCenter, "Total Marks (Weighted)")

        painter.save()
        painter.translate(14, mt + chart_h / 2)
        painter.rotate(-90)
        painter.drawText(QRectF(-40, 0, 80, 18), Qt.AlignCenter, "Frequency")
        painter.restore()

        for grade, cutoff_val in self._thresholds:
            if min_t <= cutoff_val <= max_t:
                color = QColor(GRADE_FG.get(grade, TEXT_DIM))
                pen = QPen(color, 1.5, Qt.DashLine)
                painter.setPen(pen)
                x = ml + (cutoff_val - min_t) / (max_t - min_t) * chart_w
                painter.drawLine(int(x), mt, int(x), h - mb)
                painter.setFont(QFont("Segoe UI", 8, QFont.Bold))
                painter.drawText(int(x) - 12, mt - 4, f"≥{grade}")

        for i, (rect, row_idx, id_str, grade, total) in enumerate(self._blocks):
            bg_col, fg_col = GRADE_PAL.get(grade, (BG_CARD, TEXT_SEC))
            color = QColor(fg_col)
            if i == self._hovered:
                color = color.lighter(150)
            painter.setBrush(QBrush(color))
            painter.setPen(QPen(QColor(BG_MAIN), 0.8))
            painter.drawRoundedRect(rect, 2, 2)

        painter.setFont(QFont("Segoe UI", 8, QFont.Bold))
        lx = w - mr - 120
        ly = mt + 8
        for grade in GRADE_ORDER:
            _, fg = GRADE_PAL.get(grade, (BG_CARD, TEXT_SEC))
            painter.setBrush(QBrush(QColor(fg)))
            painter.setPen(Qt.NoPen)
            painter.drawRoundedRect(QRectF(lx, ly, 10, 10), 2, 2)
            painter.setPen(QColor(TEXT_SEC))
            painter.drawText(int(lx + 14), int(ly + 9), grade)
            ly += 14

        painter.end()

    def mouseMoveEvent(self, event):
        pos = event.pos()
        old_hovered = self._hovered
        self._hovered = -1
        for i, (rect, row_idx, id_str, grade, total) in enumerate(self._blocks):
            if rect.contains(QPointF(pos)):
                self._hovered = i
                QToolTip.showText(
                    event.globalPos(),
                    f"ID: {id_str}\nTotal: {total:.2f}\nGrade: {grade}\n\nClick to edit",
                    self
                )
                self.setCursor(Qt.PointingHandCursor)
                break
        else:
            self.setCursor(Qt.ArrowCursor)

        if old_hovered != self._hovered:
            self.update()

    def mousePressEvent(self, event):
        pos = event.pos()
        for i, (rect, row_idx, id_str, grade, total) in enumerate(self._blocks):
            if rect.contains(QPointF(pos)):
                self.student_clicked.emit(row_idx)
                return

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.update()


# ═══════════════════════════════════════════════════════════════════════
#  CUTOFF DISPLAY WIDGET
# ═══════════════════════════════════════════════════════════════════════
class NormalDistributionGraphWidget(QWidget):
    """Show the fitted normal distribution curve for weighted totals."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumHeight(380)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self._data = []
        self._thresholds = []
        self._mu = None
        self._sigma = None
        self._margin_l = 65
        self._margin_r = 25
        self._margin_t = 30
        self._margin_b = 55

    def set_data(self, data_list, thresholds=None, mu=None, sigma=None):
        self._data = data_list
        self._thresholds = thresholds or []
        totals = [float(d[0]) for d in data_list]
        self._mu = float(mu) if mu is not None else (sum(totals) / len(totals) if totals else None)
        if sigma is not None:
            self._sigma = float(sigma)
        elif len(totals) > 1:
            self._sigma = float(np.std(totals, ddof=1))
        else:
            self._sigma = 0.001
        if self._sigma == 0:
            self._sigma = 0.001
        self.update()

    def clear_data(self):
        self._data = []
        self._thresholds = []
        self._mu = None
        self._sigma = None
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, True)
        painter.fillRect(self.rect(), QColor(BG_MAIN))

        w = self.width()
        h = self.height()
        ml, mr, mt, mb = self._margin_l, self._margin_r, self._margin_t, self._margin_b
        chart_w = w - ml - mr
        chart_h = h - mt - mb

        if not self._data or self._mu is None or self._sigma is None or chart_w <= 0 or chart_h <= 0:
            painter.setPen(QColor(TEXT_DIM))
            painter.setFont(QFont("Segoe UI", 13))
            painter.drawText(self.rect(), Qt.AlignCenter, "Process data to see the normal curve")
            painter.end()
            return

        totals = [float(d[0]) for d in self._data]
        min_t = min(min(totals), self._mu - 3.2 * self._sigma)
        max_t = max(max(totals), self._mu + 3.2 * self._sigma)
        if max_t <= min_t:
            max_t = min_t + 1

        chart_rect = QRectF(ml, mt, chart_w, chart_h)
        painter.fillRect(chart_rect, QColor(BG_CARD))

        painter.setPen(QPen(QColor(BORDER), 0.5))
        for i in range(1, 8):
            y = mt + chart_h * i / 8
            painter.drawLine(int(ml), int(y), int(w - mr), int(y))

        painter.setPen(QPen(QColor(BORDER_HI), 1.2))
        painter.drawLine(ml, mt, ml, h - mb)
        painter.drawLine(ml, h - mb, w - mr, h - mb)

        def x_for(val: float) -> float:
            return ml + ((val - min_t) / (max_t - min_t)) * chart_w

        pdf_max = 1.0 / (self._sigma * math.sqrt(2 * math.pi))
        curve = QPainterPath()
        fill = QPainterPath()
        samples = 220
        first_x = x_for(min_t)
        for i in range(samples + 1):
            val = min_t + (max_t - min_t) * i / samples
            pdf = pdf_max * math.exp(-0.5 * ((val - self._mu) / self._sigma) ** 2)
            x = x_for(val)
            y = h - mb - (pdf / pdf_max) * (chart_h * 0.82)
            if i == 0:
                curve.moveTo(x, y)
                fill.moveTo(x, h - mb)
                fill.lineTo(x, y)
            else:
                curve.lineTo(x, y)
                fill.lineTo(x, y)
        fill.lineTo(x_for(max_t), h - mb)
        fill.lineTo(first_x, h - mb)

        painter.fillPath(fill, QColor(GOLD_BG))
        painter.setPen(QPen(QColor(GOLD), 2.2))
        painter.drawPath(curve)

        mu_x = x_for(self._mu)
        painter.setPen(QPen(QColor(WHITE), 1.4, Qt.DashLine))
        painter.drawLine(int(mu_x), mt, int(mu_x), h - mb)
        painter.setFont(QFont("Segoe UI", 8, QFont.Bold))
        painter.setPen(QColor(WHITE))
        painter.drawText(int(mu_x) - 18, mt + 14, "Mean")

        for grade, cutoff_val in self._thresholds:
            if min_t <= cutoff_val <= max_t:
                color = QColor(GRADE_FG.get(grade, TEXT_DIM))
                painter.setPen(QPen(color, 1.3, Qt.DashLine))
                x = x_for(cutoff_val)
                painter.drawLine(int(x), mt, int(x), h - mb)
                painter.setFont(QFont("Segoe UI", 8, QFont.Bold))
                painter.drawText(int(x) - 12, mt - 4, f">={grade}")

        painter.setFont(QFont("Segoe UI", 9))
        for i in range(9):
            val = min_t + (max_t - min_t) * i / 8
            x = x_for(val)
            painter.setPen(QColor(TEXT_SEC))
            painter.drawText(QRectF(x - 25, h - mb + 6, 50, 18), Qt.AlignCenter, f"{val:.1f}")
            painter.setPen(QPen(QColor(BORDER_HI), 0.5))
            painter.drawLine(int(x), int(h - mb), int(x), int(h - mb + 4))

        painter.setPen(QColor(TEXT_SEC))
        painter.setFont(QFont("Segoe UI", 10, QFont.Bold))
        painter.drawText(QRectF(ml, h - 18, chart_w, 18), Qt.AlignCenter, "Total Marks (Weighted)")

        painter.save()
        painter.translate(14, mt + chart_h / 2)
        painter.rotate(-90)
        painter.drawText(QRectF(-70, 0, 140, 18), Qt.AlignCenter, "Normal Density")
        painter.restore()

        painter.setFont(QFont("Segoe UI", 10, QFont.Bold))
        painter.setPen(QColor(TEXT_PRI))
        painter.drawText(
            QRectF(ml, mt + 8, chart_w, 22),
            Qt.AlignCenter,
            f"Normal curve: mean={self._mu:.2f}, sd={self._sigma:.2f}, n={len(totals)}",
        )
        painter.end()


class CutoffDisplayWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("background: transparent;")
        self._lay = QVBoxLayout(self)
        self._lay.setContentsMargins(0, 0, 0, 0)
        self._lay.setSpacing(4)
        self._grid = QGridLayout()
        self._grid.setSpacing(4)
        self._lay.addLayout(self._grid)

    def set_cutoffs(self, cutoff_ranges: Dict[str, Tuple[float, float]]):
        for i in reversed(range(self._grid.count())):
            w = self._grid.itemAt(i).widget()
            if w:
                w.deleteLater()

        row = 0
        for grade in GRADE_ORDER:
            if grade not in cutoff_ranges:
                continue
            mn, mx = cutoff_ranges[grade]
            bg, fg = GRADE_PAL.get(grade, (BG_CARD, TEXT_SEC))

            g_lbl = QLabel(f"  {grade}  ")
            g_lbl.setAlignment(Qt.AlignCenter)
            g_lbl.setFixedWidth(50)
            g_lbl.setStyleSheet(f"""
                background-color: {bg};
                color: {fg};
                border: 1px solid {fg}55;
                border-radius: 4px;
                font-weight: bold;
                font-size: 12px;
                padding: 2px 6px;
            """)

            if mx == float('inf'):
                range_text = f"≥ {mn:.2f}"
            elif mn == float('-inf'):
                range_text = f"< {mx:.2f}"
            else:
                range_text = f"{mn:.2f}  —  {mx:.2f}"

            r_lbl = QLabel(range_text)
            r_lbl.setStyleSheet(f"color: {TEXT_PRI}; font-size: 13px; background: transparent;")

            self._grid.addWidget(g_lbl, row, 0)
            self._grid.addWidget(r_lbl, row, 1)
            row += 1


# ═══════════════════════════════════════════════════════════════════════
#  CUSTOM CUTOFF DIALOG
# ═══════════════════════════════════════════════════════════════════════
class CustomCutoffDialog(QDialog):
    def __init__(self, mu: float, sigma: float, base_grade: str,
                 current_thresholds: list = None, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Custom Grade Cutoffs")
        self.setMinimumWidth(450)
        self.setStyleSheet(f"""
            QDialog {{
                background-color: {BG_CARD};
                color: {TEXT_PRI};
            }}
        """)

        layout = QVBoxLayout(self)
        layout.setSpacing(14)

        info = QLabel(
            f"Mean (μ): {mu:.2f}   |   Std Dev (σ): {sigma:.2f}\n"
            f"Enter the minimum marks for each grade.\n"
            f"Grades are checked top-down; first match wins."
        )
        info.setStyleSheet(f"color: {TEXT_SEC}; font-size: 13px;")
        info.setWordWrap(True)
        layout.addWidget(info)

        if current_thresholds is None:
            current_thresholds = compute_default_thresholds(mu, sigma, base_grade)
        thresh_dict = {g: v for g, v in current_thresholds}

        form = QFormLayout()
        form.setSpacing(10)
        self._spins = {}

        for grade in GRADE_ORDER:
            if grade == "F":
                continue
            spin = QDoubleSpinBox()
            spin.setRange(-9999, 9999)
            spin.setDecimals(2)
            spin.setSingleStep(0.5)
            if grade in thresh_dict:
                spin.setValue(thresh_dict[grade])
            else:
                spin.setValue(0)

            bg, fg = GRADE_PAL.get(grade, (BG_CARD, TEXT_SEC))
            lbl = QLabel(f"  {grade}  ")
            lbl.setAlignment(Qt.AlignCenter)
            lbl.setStyleSheet(f"""
                background-color: {bg};
                color: {fg};
                border-radius: 4px;
                font-weight: bold;
                font-size: 13px;
                padding: 3px 10px;
            """)
            form.addRow(lbl, spin)
            self._spins[grade] = spin

        layout.addLayout(form)

        f_label = QLabel("Grade F: Scores below the lowest cutoff above")
        f_label.setStyleSheet(f"color: {GRADE_FG.get('F', TEXT_SEC)}; font-size: 12px; font-style: italic;")
        layout.addWidget(f_label)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.setStyleSheet(f"""
            QPushButton {{
                background-color: {GOLD};
                color: #060606;
                border: none;
                border-radius: 6px;
                padding: 8px 22px;
                font-weight: bold;
                font-size: 14px;
            }}
            QPushButton:hover {{ background-color: #d4b55a; }}
        """)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_thresholds(self) -> list:
        result = []
        for grade in GRADE_ORDER:
            if grade == "F":
                continue
            if grade in self._spins:
                result.append((grade, self._spins[grade].value()))
        result.sort(key=lambda x: x[1], reverse=True)
        return result


# ═══════════════════════════════════════════════════════════════════════
#  EDIT STUDENT DIALOG
# ═══════════════════════════════════════════════════════════════════════
class EditStudentDialog(QDialog):
    def __init__(self, student_id: str, marks_dict: dict,
                 current_grade: str, current_total: float,
                 numeric_cols: list, components: list, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Edit Student: {student_id}")
        self.setMinimumWidth(460)
        self.setStyleSheet(f"""
            QDialog {{
                background-color: {BG_CARD};
                color: {TEXT_PRI};
            }}
        """)

        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        header = QLabel(f"🎓  {student_id}")
        header.setStyleSheet(f"font-size: 18px; font-weight: bold; color: {GOLD}; background: transparent;")
        layout.addWidget(header)

        bg, fg = GRADE_PAL.get(current_grade, (BG_CARD, TEXT_SEC))
        info_row = QHBoxLayout()
        total_lbl = QLabel(f"Total: {current_total:.2f}")
        total_lbl.setStyleSheet(f"color: {GOLD}; font-size: 14px; font-weight: bold; background: transparent;")
        info_row.addWidget(total_lbl)

        grade_lbl = QLabel(f"  {current_grade}  ")
        grade_lbl.setAlignment(Qt.AlignCenter)
        grade_lbl.setStyleSheet(f"""
            background-color: {bg}; color: {fg};
            border: 1px solid {fg}55; border-radius: 4px;
            font-weight: bold; font-size: 14px; padding: 2px 10px;
        """)
        info_row.addWidget(grade_lbl)
        info_row.addStretch()
        layout.addLayout(info_row)

        layout.addWidget(_hline())

        marks_header = QLabel("SUBJECT MARKS")
        marks_header.setStyleSheet(
            f"color:{TEXT_DIM};font-size:10px;font-weight:bold;"
            f"letter-spacing:1.8px;background:transparent;"
        )
        layout.addWidget(marks_header)

        comp_lookup = {c['column']: c for c in components}

        form = QFormLayout()
        form.setSpacing(8)
        self._mark_spins = {}

        for col in numeric_cols:
            if col == 'Total':
                continue
            comp = comp_lookup.get(col)
            max_val = comp['max_marks'] if comp else 999

            spin = QDoubleSpinBox()
            spin.setRange(0, max_val * 2)
            spin.setDecimals(2)
            spin.setSingleStep(1.0)
            spin.setValue(float(marks_dict.get(col, 0)))

            col_label = QLabel(col)
            col_label.setStyleSheet(f"color: {TEXT_PRI}; font-size: 14px; background: transparent;")

            form.addRow(col_label, spin)
            self._mark_spins[col] = spin

        layout.addLayout(form)
        layout.addWidget(_hline())

        grade_header = QLabel("GRADE OVERRIDE (OPTIONAL)")
        grade_header.setStyleSheet(
            f"color:{TEXT_DIM};font-size:10px;font-weight:bold;"
            f"letter-spacing:1.8px;background:transparent;"
        )
        layout.addWidget(grade_header)

        grade_row = QHBoxLayout()
        grade_row.addWidget(QLabel("Override:"))
        self._grade_combo = QComboBox()
        self._grade_combo.addItems(["Auto (from cutoffs)"] + GRADE_ORDER)
        self._grade_combo.setFixedWidth(200)
        grade_row.addWidget(self._grade_combo)
        grade_row.addStretch()
        layout.addLayout(grade_row)

        note = QLabel("⚠ Mean & Std Dev remain frozen after edits.")
        note.setStyleSheet(f"color: {TEXT_DIM}; font-size: 11px; font-style: italic; background: transparent;")
        layout.addWidget(note)

        buttons = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        buttons.setStyleSheet(f"""
            QPushButton {{
                background-color: {GOLD};
                color: #060606;
                border: none;
                border-radius: 6px;
                padding: 8px 22px;
                font-weight: bold;
                font-size: 14px;
            }}
            QPushButton:hover {{ background-color: #d4b55a; }}
        """)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_marks(self) -> dict:
        return {col: spin.value() for col, spin in self._mark_spins.items()}

    def get_grade_override(self) -> Optional[str]:
        text = self._grade_combo.currentText()
        if text.startswith("Auto"):
            return None
        return text


# ═══════════════════════════════════════════════════════════════════════
#  COMPONENT ROW
# ═══════════════════════════════════════════════════════════════════════
class ComponentRow(QWidget):
    def __init__(self, col_name: str, parent=None):
        super().__init__(parent)
        self.column_name = col_name
        self.setFixedHeight(54)
        self.setStyleSheet("background:transparent;")
        lay = QHBoxLayout(self)
        lay.setContentsMargins(2, 2, 8, 2)
        lay.setSpacing(10)

        self.checkbox = QCheckBox()
        self.checkbox.setChecked(True)
        self.checkbox.setFixedWidth(24)
        self.checkbox.stateChanged.connect(self._toggle)
        lay.addWidget(self.checkbox)

        self.lbl = QLabel(col_name)
        self.lbl.setStyleSheet(
            f"color:{TEXT_PRI};font-size:17px;background:transparent;"
        )
        self.lbl.setToolTip(col_name)
        self.lbl.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        lay.addWidget(self.lbl, 1)

        self.max_input = QLineEdit()
        self.max_input.setPlaceholderText("Max")
        self.max_input.setFixedWidth(78)
        self.max_input.setAlignment(Qt.AlignCenter)
        lay.addWidget(self.max_input)

        self.weight_input = QLineEdit()
        self.weight_input.setPlaceholderText("Wt")
        self.weight_input.setFixedWidth(78)
        self.weight_input.setAlignment(Qt.AlignCenter)
        lay.addWidget(self.weight_input)

    def _toggle(self, state: int):
        on = state == Qt.Checked
        self.max_input.setEnabled(on)
        self.weight_input.setEnabled(on)
        self.lbl.setStyleSheet(
            f"color:{TEXT_PRI if on else TEXT_DIM};font-size:17px;background:transparent;"
        )

    def get_config(self) -> Optional[dict]:
        if not self.checkbox.isChecked():
            return None
        return {
            'column':     self.column_name,
            'max_raw':    self.max_input.text().strip(),
            'weight_raw': self.weight_input.text().strip(),
        }


# ═══════════════════════════════════════════════════════════════════════
#  STAT CARD
# ═══════════════════════════════════════════════════════════════════════
class StatCard(QWidget):
    def __init__(self, label: str, accent: bool = False, parent=None):
        super().__init__(parent)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.setFixedHeight(72)
        self.setStyleSheet(f"""
            QWidget {{
                background-color: {BG_CARD};
                border-bottom: 1px solid {BORDER};
                border-right: 1px solid {BORDER};
            }}
        """)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(14, 9, 14, 9)
        lay.setSpacing(3)

        lbl = QLabel(label)
        lbl.setStyleSheet(
            f"color:{TEXT_SEC};font-size:9px;font-weight:bold;"
            f"letter-spacing:1.8px;background:transparent;border:none;"
        )
        lay.addWidget(lbl)

        self._val = QLabel("—")
        self._val.setStyleSheet(
            f"color:{GOLD if accent else TEXT_PRI};"
            f"font-size:24px;font-weight:bold;"
            f"background:transparent;border:none;"
        )
        lay.addWidget(self._val)

    def set_value(self, text: str):
        self._val.setText(text)


# ═══════════════════════════════════════════════════════════════════════
#  GRADE BADGE
# ═══════════════════════════════════════════════════════════════════════
class GradeBadge(QLabel):
    def __init__(self, grade: str, count: int, parent=None):
        super().__init__(parent)
        bg, fg = GRADE_PAL.get(grade, (BG_CARD, TEXT_SEC))
        self.setText(f"{grade}: {count}")
        self.setAlignment(Qt.AlignCenter)
        self.setContentsMargins(10, 3, 10, 3)
        self.setStyleSheet(f"""
            QLabel {{
                background-color: {bg};
                color: {fg};
                border: 1px solid {fg}55;
                border-radius: 4px;
                font-size: 11px;
                font-weight: bold;
                padding: 2px 4px;
            }}
        """)


# ═══════════════════════════════════════════════════════════════════════
#  FILE DROP ZONE
# ═══════════════════════════════════════════════════════════════════════
class FileDropZone(QWidget):
    def __init__(self, on_file, parent=None):
        super().__init__(parent)
        self.on_file = on_file
        self.setAcceptDrops(True)
        self.setCursor(Qt.PointingHandCursor)
        self.setFixedHeight(84)
        self._set_idle()
        lay = QVBoxLayout(self)
        lay.setContentsMargins(8, 8, 8, 8)
        lay.setSpacing(2)
        for txt, sz, col in [
            ("↑", 16, TEXT_SEC),
            ("Click to upload  or  drag & drop", 11, TEXT_DIM),
            (".csv  ·  .xlsx  ·  .xls",          10, TEXT_DIM),
        ]:
            l = QLabel(txt)
            l.setAlignment(Qt.AlignCenter)
            l.setStyleSheet(
                f"color:{col};font-size:{sz}px;background:transparent;border:none;"
            )
            lay.addWidget(l)

    def _set_idle(self):
        self.setStyleSheet(f"""
            background-color: {BG_INPUT};
            border: 1px dashed {BORDER_HI};
            border-radius: 7px;
        """)

    def _set_active(self):
        self.setStyleSheet(f"""
            background-color: {GOLD_BG};
            border: 1px dashed {GOLD_DIM};
            border-radius: 7px;
        """)

    def mousePressEvent(self, _):
        path, _ = QFileDialog.getOpenFileName(
            self, "Select File", "", "Data Files (*.csv *.xlsx *.xls)"
        )
        if path:
            self.on_file(path)

    def dragEnterEvent(self, e):
        if e.mimeData().hasUrls():
            e.acceptProposedAction()
            self._set_active()

    def dragLeaveEvent(self, _):
        self._set_idle()

    def dropEvent(self, e):
        self._set_idle()
        urls = e.mimeData().urls()
        if urls:
            self.on_file(urls[0].toLocalFile())


# ═══════════════════════════════════════════════════════════════════════
#  LOADED FILE WIDGET
# ═══════════════════════════════════════════════════════════════════════
class LoadedFileWidget(QWidget):
    def __init__(self, filename: str, row_count: int, on_clear, parent=None):
        super().__init__(parent)
        self.setFixedHeight(84)
        self.setStyleSheet(f"""
            background-color: {BG_INPUT};
            border: 1px solid {BORDER_HI};
            border-radius: 7px;
        """)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(12, 8, 12, 8)
        lay.setSpacing(5)

        row1 = QHBoxLayout()
        row1.setSpacing(8)

        icon = QLabel("📄")
        icon.setStyleSheet("background:transparent;font-size:14px;border:none;")
        row1.addWidget(icon)

        nm = QLabel(filename)
        nm.setStyleSheet(
            f"color:{GOLD};"
            f"font-family:'Consolas','Courier New',monospace;"
            f"font-size:12px;background:transparent;border:none;"
        )
        nm.setToolTip(filename)
        row1.addWidget(nm, 1)

        x = QPushButton("✕")
        x.setFixedSize(18, 18)
        x.setCursor(Qt.PointingHandCursor)
        x.setStyleSheet(
            f"background:transparent;color:{TEXT_DIM};"
            f"border:none;font-size:11px;padding:0;"
        )
        x.clicked.connect(on_clear)
        row1.addWidget(x)
        lay.addLayout(row1)

        status = QLabel(f"Columns loaded  ·  {row_count} students")
        status.setStyleSheet(
            f"color:{TEXT_DIM};font-size:11px;background:transparent;border:none;"
        )
        lay.addWidget(status)


# ═══════════════════════════════════════════════════════════════════════
#  MAIN WINDOW
# ═══════════════════════════════════════════════════════════════════════
class GradingApp(QMainWindow):

    df_raw:    Optional[pd.DataFrame]
    df_graded: Optional[pd.DataFrame]
    comp_rows: List[ComponentRow]

    def __init__(self):
        super().__init__()
        self.df_raw = self.df_graded = None
        self.comp_rows = []
        self.team_comp_rows: List[ComponentRow] = []
        self._sidebar_open = True
        self._sidebar_anim = None
        self._frozen_mu = None
        self._frozen_sigma = None
        self._thresholds = []
        self._cutoff_ranges = {}
        self._use_custom_cutoffs = False
        self._custom_thresholds = None
        self._components = []
        self._numeric_cols = []
        self._id_col = None
        self._edit_log = []
        self.df_team: Optional[pd.DataFrame] = None
        self._team_path = ""
        self._team_name_col = None
        self._team_member_cols: list = []
        self._team_marks_cols: list = []
        self._team_numeric_only_member = False
        self._team_roll_marks: Dict[str, float] = {}
        # NEW: cluster state
        self._clusters: Dict[str, list] = load_clusters()
        self._unsaved_changes = False  # NEW: track unsaved state

        self.setWindowTitle("Automated Grading System")
        self.setMinimumSize(1100, 700)
        self.resize(1380, 900)
        self.setStyleSheet(STYLESHEET)
        self._center()
        self._build()
        self.statusBar().showMessage(
            "Ready — upload a CSV or Excel file to begin."
        )

        # ── Feature 1: offer to resume last session (delayed so UI is visible) ──
        QTimer.singleShot(300, self._maybe_offer_resume)

    def closeEvent(self, event):
        """NEW: Prompt user to save before closing (like Notepad/Excel)."""
        if self.df_graded is not None and self._unsaved_changes:
            reply = QMessageBox(self)
            reply.setWindowTitle("Save Changes?")
            reply.setText("You have unsaved grading changes.\nDo you want to save before closing?")
            reply.setIcon(QMessageBox.Question)
            save_btn   = reply.addButton("💾  Save",       QMessageBox.AcceptRole)
            nosave_btn = reply.addButton("Don't Save",     QMessageBox.DestructiveRole)
            cancel_btn = reply.addButton("Cancel",         QMessageBox.RejectRole)
            reply.setDefaultButton(save_btn)
            reply.exec_()
            clicked = reply.clickedButton()
            if clicked == cancel_btn:
                event.ignore()
                return
            elif clicked == save_btn:
                saved = self._action_save_session()
                if not saved:
                    # User cancelled the save dialog — don't close
                    event.ignore()
                    return
            # else Don't Save — just exit without saving
        elif self.df_graded is not None:
            # Auto-save silently on close (no changes tracked but data exists)
            save_current_session(self)
        super().closeEvent(event)

    def _center(self):
        g = QApplication.primaryScreen().availableGeometry()
        self.move(
            (g.width()  - self.width())  // 2,
            (g.height() - self.height()) // 2,
        )

    # ── BUILD ─────────────────────────────────────────────────────────
    def _build(self):
        root_w = QWidget()
        self.setCentralWidget(root_w)
        root = QVBoxLayout(root_w)
        root.setContentsMargins(20, 16, 20, 10)
        root.setSpacing(14)
        root.addWidget(self._make_header())

        body = QHBoxLayout()
        body.setSpacing(0)

        self._left_panel = self._make_left_panel()
        body.addWidget(self._left_panel)

        body.addWidget(self._make_right_panel(), 1)
        root.addLayout(body, 1)

    # ── SIDEBAR TOGGLE ────────────────────────────────────────────────
    def _toggle_sidebar(self):
        if self._sidebar_anim and self._sidebar_anim.state() == QPropertyAnimation.Running:
            return

        panel = self._left_panel
        start = panel.width()

        if self._sidebar_open:
            end = 0
            self._sidebar_open = False
            self._toggle_btn.setText("▶")
            self._toggle_btn.setToolTip("Show sidebar")
        else:
            end = 470
            self._sidebar_open = True
            self._toggle_btn.setText("◀")
            self._toggle_btn.setToolTip("Hide sidebar")

        anim = QPropertyAnimation(panel, b"maximumWidth")
        anim.setDuration(220)
        anim.setStartValue(start)
        anim.setEndValue(end)
        anim.setEasingCurve(QEasingCurve.InOutCubic)
        anim2 = QPropertyAnimation(panel, b"minimumWidth")
        anim2.setDuration(220)
        anim2.setStartValue(start)
        anim2.setEndValue(end)
        anim2.setEasingCurve(QEasingCurve.InOutCubic)
        anim.start()
        anim2.start()
        self._sidebar_anim  = anim
        self._sidebar_anim2 = anim2

    # ── HEADER ────────────────────────────────────────────────────────
    def _make_header(self) -> QWidget:
        w = QWidget()
        w.setStyleSheet("background:transparent;")
        lay = QHBoxLayout(w)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(10)

        self._toggle_btn = QPushButton("◀")
        self._toggle_btn.setFixedSize(32, 32)
        self._toggle_btn.setToolTip("Hide sidebar")
        self._toggle_btn.setCursor(Qt.PointingHandCursor)
        self._toggle_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {BG_CARD};
                color: {TEXT_SEC};
                border: 1px solid {BORDER_HI};
                border-radius: 7px;
                font-size: 12px;
                font-weight: bold;
            }}
            QPushButton:hover  {{ background-color:{BG_HOVER}; color:{TEXT_PRI}; border-color:{GOLD_DIM}; }}
            QPushButton:pressed {{ background-color:{BG_INPUT}; }}
        """)
        self._toggle_btn.clicked.connect(self._toggle_sidebar)
        lay.addWidget(self._toggle_btn)

        cap = QLabel("🎓")
        cap.setStyleSheet("font-size:26px;background:transparent;")
        lay.addWidget(cap)

        title = QLabel("Automated Grading System")
        title.setStyleSheet(
            f"color:{TEXT_PRI};font-size:26px;font-weight:bold;"
            f"letter-spacing:-0.5px;background:transparent;"
        )
        lay.addWidget(title)

        sub = QLabel("Relative grading with statistical curve analysis")
        sub.setStyleSheet(
            f"color:{TEXT_DIM};font-size:12px;background:transparent;"
            f"margin-top:7px;margin-left:2px;"
        )
        lay.addWidget(sub)
        lay.addStretch()

        # ── NEW: Save Session button ──
        self._btn_save = _small_btn("💾  Save", height=32)
        self._btn_save.setToolTip("Save current session with a custom name")
        self._btn_save.clicked.connect(self._action_save_session)
        lay.addWidget(self._btn_save)

        # ── Feature 1: History / Resume button ──
        self._btn_history = _small_btn("🕒  History", height=32)
        self._btn_history.setToolTip("Resume a previous grading session")
        self._btn_history.clicked.connect(self._on_history)
        lay.addWidget(self._btn_history)

        # ── NEW: Manage History button ──
        self._btn_manage_history = _small_btn("🗑  Manage", height=32)
        self._btn_manage_history.setToolTip("Delete saved history entries")
        self._btn_manage_history.clicked.connect(self._on_manage_history)
        lay.addWidget(self._btn_manage_history)

        badge = QLabel("v2.0")
        badge.setStyleSheet(f"""
            background-color: {GOLD_BG};
            color: {GOLD};
            border: 1px solid {GOLD_DIM};
            border-radius: 10px;
            padding: 4px 12px;
            font-size: 11px;
            font-weight: bold;
        """)
        lay.addWidget(badge)

        return w

    # ── LEFT PANEL ────────────────────────────────────────────────────
    def _make_left_panel(self) -> QWidget:
        panel = QWidget()
        panel.setMinimumWidth(470)
        panel.setMaximumWidth(470)
        panel.setStyleSheet(f"""
            QWidget {{
                background-color: {BG_PANEL};
                border: 1px solid {BORDER};
                border-right: none;
            }}
        """)

        outer_scroll = QScrollArea()
        outer_scroll.setWidgetResizable(True)
        outer_scroll.setStyleSheet("background:transparent; border:none;")

        inner = QWidget()
        inner.setStyleSheet("background:transparent;")
        lay = QVBoxLayout(inner)
        lay.setContentsMargins(20, 20, 20, 20)
        lay.setSpacing(14)

        # DATA SOURCE
        lay.addWidget(_lbl("↑  DATA SOURCE", size=13))
        lay.addWidget(_hline())

        self._source_tabs = QTabWidget()
        self._source_tabs.setStyleSheet(f"QTabWidget {{ border: none; background: {BG_PANEL}; }}")

        main_src = QWidget()
        main_src.setStyleSheet("background:transparent;border:none;")
        main_lay = QVBoxLayout(main_src)
        main_lay.setContentsMargins(0, 6, 0, 0)
        self.file_stack = QStackedWidget()
        self.file_stack.setStyleSheet("background:transparent;border:none;")
        self.file_stack.setFixedHeight(95)
        self._drop_zone = FileDropZone(self._on_file)
        self.file_stack.addWidget(self._drop_zone)
        _ph = QWidget()
        _ph.setStyleSheet("background:transparent;border:none;")
        self.file_stack.addWidget(_ph)
        main_lay.addWidget(self.file_stack)
        self._source_tabs.addTab(main_src, "Main File")

        team_src = QWidget()
        team_src.setStyleSheet("background:transparent;border:none;")
        team_lay = QVBoxLayout(team_src)
        team_lay.setContentsMargins(0, 6, 0, 0)
        self.team_file_stack = QStackedWidget()
        self.team_file_stack.setStyleSheet("background:transparent;border:none;")
        self.team_file_stack.setFixedHeight(95)
        self._team_drop_zone = FileDropZone(self._on_team_file)
        self.team_file_stack.addWidget(self._team_drop_zone)
        _team_ph = QWidget()
        _team_ph.setStyleSheet("background:transparent;border:none;")
        self.team_file_stack.addWidget(_team_ph)
        team_lay.addWidget(self.team_file_stack)
        self._source_tabs.addTab(team_src, "Team File")

        lay.addWidget(self._source_tabs)

        # COMPONENTS
        lay.addWidget(_hline())
        ch = QHBoxLayout()
        ch.addWidget(_lbl("≡  COMPONENTS", size=13))
        ch.addStretch()
        self._sel_count = QLabel("")
        self._sel_count.setStyleSheet(
            f"color:{GOLD};font-size:17px;font-weight:bold;background:transparent;"
        )
        ch.addWidget(self._sel_count)
        lay.addLayout(ch)

        sh = QWidget()
        sh.setStyleSheet("background:transparent;")
        sh_lay = QHBoxLayout(sh)
        sh_lay.setContentsMargins(28, 0, 8, 0)
        sh_lay.setSpacing(10)
        col_l = _lbl("COLUMN", size=14)
        col_l.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        sh_lay.addWidget(col_l, 1)
        for t in ("MAX", "WEIGHT"):
            l = _lbl(t, size=14)
            l.setFixedWidth(78)
            l.setAlignment(Qt.AlignCenter)
            sh_lay.addWidget(l)

        self._component_tabs = QTabWidget()
        self._component_tabs.setStyleSheet(f"QTabWidget {{ border: none; background: {BG_PANEL}; }}")

        main_comp_tab = QWidget()
        main_comp_tab.setStyleSheet("background:transparent;border:none;")
        main_comp_lay = QVBoxLayout(main_comp_tab)
        main_comp_lay.setContentsMargins(0, 6, 0, 0)
        main_comp_lay.setSpacing(2)
        main_comp_lay.addWidget(sh)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setStyleSheet("background:transparent;border:none;")
        self.scroll_inner = QWidget()
        self.scroll_inner.setStyleSheet("background:transparent;")
        self.scroll_layout = QVBoxLayout(self.scroll_inner)
        self.scroll_layout.setContentsMargins(0, 2, 0, 2)
        self.scroll_layout.setSpacing(2)
        self.scroll_layout.addStretch()
        self.scroll_area.setWidget(self.scroll_inner)
        self.scroll_area.setMinimumHeight(130)
        main_comp_lay.addWidget(self.scroll_area, 1)
        self._component_tabs.addTab(main_comp_tab, "Main Components")

        team_comp_tab = QWidget()
        team_comp_tab.setStyleSheet("background:transparent;border:none;")
        team_comp_lay = QVBoxLayout(team_comp_tab)
        team_comp_lay.setContentsMargins(0, 8, 0, 0)
        team_comp_lay.setSpacing(2)

        team_sh = QWidget()
        team_sh.setStyleSheet("background:transparent;")
        team_sh_lay = QHBoxLayout(team_sh)
        team_sh_lay.setContentsMargins(28, 0, 8, 0)
        team_sh_lay.setSpacing(10)
        team_col_l = _lbl("COLUMN", size=14)
        team_col_l.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        team_sh_lay.addWidget(team_col_l, 1)
        for t in ("MAX", "WEIGHT"):
            l = _lbl(t, size=14)
            l.setFixedWidth(78)
            l.setAlignment(Qt.AlignCenter)
            team_sh_lay.addWidget(l)
        team_comp_lay.addWidget(team_sh)

        self.team_scroll_area = QScrollArea()
        self.team_scroll_area.setWidgetResizable(True)
        self.team_scroll_area.setStyleSheet("background:transparent;border:none;")
        self.team_scroll_inner = QWidget()
        self.team_scroll_inner.setStyleSheet("background:transparent;")
        self.team_scroll_layout = QVBoxLayout(self.team_scroll_inner)
        self.team_scroll_layout.setContentsMargins(0, 2, 0, 2)
        self.team_scroll_layout.setSpacing(2)
        self.team_scroll_layout.addStretch()
        self.team_scroll_area.setWidget(self.team_scroll_inner)
        self.team_scroll_area.setMinimumHeight(100)
        team_comp_lay.addWidget(self.team_scroll_area, 1)

        self.btn_apply_team_marks = _small_btn("Merge Selected Team Marks", height=34)
        self.btn_apply_team_marks.setEnabled(False)
        self.btn_apply_team_marks.clicked.connect(self._on_apply_team_marks)
        team_comp_lay.addWidget(self.btn_apply_team_marks)
        self._component_tabs.addTab(team_comp_tab, "Team Marks")

        lay.addWidget(self._component_tabs, 1)

        # Total weightage
        lay.addWidget(_hline())
        tw = QHBoxLayout()
        tw.addWidget(_lbl("Total Weightage", color=TEXT_SEC, size=16, spacing="0"))
        tw.addStretch()
        self._tw_lbl = QLabel("0.0")
        self._tw_lbl.setStyleSheet(
            f"color:{GOLD};font-size:20px;font-weight:bold;background:transparent;"
        )
        tw.addWidget(self._tw_lbl)
        lay.addLayout(tw)
        lay.addWidget(_hline())

        # Base Grade
        bg_row = QHBoxLayout()
        bg_row.addWidget(_lbl("Base Grade", color=TEXT_SEC, size=16, spacing="0"))
        bg_row.addStretch()
        self.combo_base = QComboBox()
        self.combo_base.addItems(["B", "B-"])
        self.combo_base.setFixedWidth(110)
        bg_row.addWidget(self.combo_base)
        lay.addLayout(bg_row)

        # Cutoff Mode
        lay.addWidget(_hline())
        lay.addWidget(_lbl("✂  CUTOFF MODE", size=13))

        self._radio_default = QRadioButton("Default (Mean ± σ × factor)")
        self._radio_custom = QRadioButton("Custom (Manual entry)")
        self._radio_default.setChecked(True)

        self._cutoff_group = QButtonGroup(self)
        self._cutoff_group.addButton(self._radio_default, 0)
        self._cutoff_group.addButton(self._radio_custom, 1)
        self._cutoff_group.buttonClicked.connect(self._on_cutoff_mode_changed)

        lay.addWidget(self._radio_default)
        lay.addWidget(self._radio_custom)

        self._btn_set_cutoffs = _small_btn("✏  Set Custom Cutoffs")
        self._btn_set_cutoffs.setVisible(False)
        self._btn_set_cutoffs.clicked.connect(self._on_set_custom_cutoffs)
        lay.addWidget(self._btn_set_cutoffs)

        lay.addWidget(_hline())

        # Buttons
        self.btn_calc = _primary_btn("⊞  Calculate Grades", height=50)
        self.btn_calc.setEnabled(False)
        self.btn_calc.clicked.connect(self.action_calculate)
        lay.addWidget(self.btn_calc)

        self.btn_export = _ghost_btn("↓  Download Excel", height=46)
        self.btn_export.setEnabled(False)
        self.btn_export.clicked.connect(self.action_export)
        lay.addWidget(self.btn_export)

        # ── NEW: Add Component button ──
        self.btn_add_comp = _ghost_btn("➕  Add Component", height=40)
        self.btn_add_comp.setEnabled(False)
        self.btn_add_comp.setToolTip("Add a new custom grading component column")
        self.btn_add_comp.clicked.connect(self._on_add_component)
        lay.addWidget(self.btn_add_comp)

        # ── NEW: Cluster Manager button ──
        self.btn_cluster = _ghost_btn("🗂  Clusters", height=40)
        self.btn_cluster.setToolTip("Manage student clusters and bulk-assign marks")
        self.btn_cluster.clicked.connect(self._on_cluster_manager)
        lay.addWidget(self.btn_cluster)

        # ── NEW: Bulk Assign Cluster Marks ──
        self.btn_cluster_marks = _ghost_btn("⚡  Assign Cluster Marks", height=40)
        self.btn_cluster_marks.setEnabled(False)
        self.btn_cluster_marks.setToolTip("Assign marks to all students in a cluster")
        self.btn_cluster_marks.clicked.connect(self._on_cluster_bulk_marks)
        lay.addWidget(self.btn_cluster_marks)

        lay.addStretch()
        outer_scroll.setWidget(inner)

        panel_lay = QVBoxLayout(panel)
        panel_lay.setContentsMargins(0, 0, 0, 0)
        panel_lay.addWidget(outer_scroll)

        return panel

    # ── RIGHT PANEL ───────────────────────────────────────────────────
    def _make_right_panel(self) -> QWidget:
        panel = QWidget()
        panel.setStyleSheet(f"""
            QWidget {{
                background-color: {BG_MAIN};
                border: 1px solid {BORDER};
            }}
        """)
        lay = QVBoxLayout(panel)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        # Stat Cards
        cards_w = QWidget()
        cards_w.setStyleSheet(f"background:{BG_MAIN};border:none;")
        cards_lay = QHBoxLayout(cards_w)
        cards_lay.setContentsMargins(0, 0, 0, 0)
        cards_lay.setSpacing(0)

        self.card_mean   = StatCard("MEAN",      accent=True)
        self.card_std    = StatCard("STD DEV",   accent=True)
        self.card_count  = StatCard("STUDENTS")
        self.card_min    = StatCard("MIN SCORE")
        self.card_max    = StatCard("MAX SCORE")
        self.card_weight = StatCard("TOTAL WT")

        all_cards = (
            self.card_mean, self.card_std, self.card_count,
            self.card_min,  self.card_max, self.card_weight,
        )
        self.card_weight.setStyleSheet(
            self.card_weight.styleSheet().replace(
                "border-right: 1px solid", "border-right: none; border-dummy: 1px solid"
            )
        )
        for card in all_cards:
            cards_lay.addWidget(card)
        lay.addWidget(cards_w)

        # Distribution badges
        dist_w = QWidget()
        dist_w.setFixedHeight(38)
        dist_w.setStyleSheet(
            f"background:{BG_MAIN};"
            f"border-bottom: 1px solid {BORDER};"
            f"border-top: none; border-left: none; border-right: none;"
        )
        dist_lay = QHBoxLayout(dist_w)
        dist_lay.setContentsMargins(16, 0, 16, 0)
        dist_lay.setSpacing(8)
        dist_lay.addWidget(_lbl("DISTRIBUTION:", color=TEXT_DIM, size=10))

        self._badge_container = QWidget()
        self._badge_container.setStyleSheet("background:transparent;border:none;")
        self._badge_lay = QHBoxLayout(self._badge_container)
        self._badge_lay.setContentsMargins(0, 0, 0, 0)
        self._badge_lay.setSpacing(6)
        self._badge_lay.addStretch()
        dist_lay.addWidget(self._badge_container, 1)
        lay.addWidget(dist_w)

        # Search bar
        search_w = QWidget()
        search_w.setFixedHeight(44)
        search_w.setStyleSheet(f"background:{BG_MAIN}; border:none;")
        search_lay = QHBoxLayout(search_w)
        search_lay.setContentsMargins(12, 4, 12, 4)
        search_lay.setSpacing(8)

        search_icon = QLabel("🔍")
        search_icon.setStyleSheet("font-size:16px; background:transparent; border:none;")
        search_lay.addWidget(search_icon)

        self._search_input = QLineEdit()
        self._search_input.setPlaceholderText("Search by roll number / student ID...")
        self._search_input.setFixedHeight(32)
        self._search_input.textChanged.connect(self._on_search)
        search_lay.addWidget(self._search_input)

        clear_search = _small_btn("✕ Clear", height=30)
        clear_search.setFixedWidth(75)
        clear_search.clicked.connect(lambda: self._search_input.clear())
        search_lay.addWidget(clear_search)

        lay.addWidget(search_w)

        # Tabs
        self._tabs = QTabWidget()
        self._tabs.setStyleSheet(f"QTabWidget {{ border: none; background: {BG_MAIN}; }}")

        # Tab 1: Results
        tab_results = QWidget()
        tab_results.setStyleSheet(f"background: {BG_MAIN};")
        tr_lay = QVBoxLayout(tab_results)
        tr_lay.setContentsMargins(0, 4, 0, 0)
        tr_lay.setSpacing(4)

        edit_row = QHBoxLayout()
        edit_hint = QLabel("💡 Double-click any row to edit marks / grade")
        edit_hint.setStyleSheet(f"color:{TEXT_DIM};font-size:11px;background:transparent;")
        edit_row.addWidget(edit_hint)
        edit_row.addStretch()
        tr_lay.addLayout(edit_row)

        self.result_table = QTableWidget()
        self.result_table.setAlternatingRowColors(True)
        self.result_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.result_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.result_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.result_table.verticalHeader().setVisible(False)
        # NEW: ResizeToContents for auto-width, but allow horizontal scroll
        self.result_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.result_table.horizontalHeader().setDefaultSectionSize(120)
        self.result_table.horizontalHeader().setMinimumSectionSize(70)
        self.result_table.setShowGrid(True)
        self.result_table.setFrameShape(QFrame.NoFrame)
        self.result_table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)  # NEW: smooth horizontal scroll
        self.result_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)          # NEW: show scrollbar when needed
        self.result_table.doubleClicked.connect(self._on_table_double_click)
        tr_lay.addWidget(self.result_table, 1)

        self._tabs.addTab(tab_results, "📊 Results")

        # Tab 2: Rankings
        tab_ranking = QWidget()
        tab_ranking.setStyleSheet(f"background: {BG_MAIN};")
        tk_lay = QVBoxLayout(tab_ranking)
        tk_lay.setContentsMargins(0, 4, 0, 0)
        tk_lay.setSpacing(4)

        self.ranking_table = QTableWidget()
        self.ranking_table.setAlternatingRowColors(True)
        self.ranking_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.ranking_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.ranking_table.setSelectionMode(QAbstractItemView.SingleSelection)
        self.ranking_table.verticalHeader().setVisible(False)
        self.ranking_table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.ranking_table.horizontalHeader().setDefaultSectionSize(120)
        self.ranking_table.horizontalHeader().setMinimumSectionSize(70)
        self.ranking_table.setShowGrid(True)
        self.ranking_table.setFrameShape(QFrame.NoFrame)
        self.ranking_table.setHorizontalScrollMode(QAbstractItemView.ScrollPerPixel)  # NEW
        self.ranking_table.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)          # NEW
        tk_lay.addWidget(self.ranking_table, 1)

        self._tabs.addTab(tab_ranking, "🏆 Rankings")

        # Tab 3: Cutoffs
        tab_cutoffs = QWidget()
        tab_cutoffs.setStyleSheet(f"background: {BG_MAIN};")
        tc_lay = QVBoxLayout(tab_cutoffs)
        tc_lay.setContentsMargins(12, 12, 12, 12)
        tc_lay.setSpacing(10)

        cutoff_title = QLabel("GRADE CUTOFF VALUES")
        cutoff_title.setStyleSheet(
            f"color:{TEXT_DIM};font-size:10px;font-weight:bold;"
            f"letter-spacing:1.8px;background:transparent;"
        )
        tc_lay.addWidget(cutoff_title)

        self._cutoff_display = CutoffDisplayWidget()
        tc_lay.addWidget(self._cutoff_display)

        cutoff_note = QLabel(
            "Each grade covers scores from its minimum (≥) up to the next grade's minimum.\n"
            "Grade F covers everything below the lowest cutoff."
        )
        cutoff_note.setStyleSheet(f"color:{TEXT_DIM};font-size:12px;background:transparent;")
        cutoff_note.setWordWrap(True)
        tc_lay.addWidget(cutoff_note)
        tc_lay.addStretch()

        self._tabs.addTab(tab_cutoffs, "✂ Cutoffs")

        # Tab 4: Graph
        tab_graph = QWidget()
        tab_graph.setStyleSheet(f"background: {BG_MAIN};")
        tg_lay = QVBoxLayout(tab_graph)
        tg_lay.setContentsMargins(4, 4, 4, 4)
        tg_lay.setSpacing(4)

        self._graph_tabs = QTabWidget()
        self._graph_tabs.setStyleSheet(f"QTabWidget {{ border: none; background: {BG_MAIN}; }}")

        tab_distribution = QWidget()
        tab_distribution.setStyleSheet(f"background: {BG_MAIN};")
        dist_lay = QVBoxLayout(tab_distribution)
        dist_lay.setContentsMargins(0, 4, 0, 0)
        dist_lay.setSpacing(4)

        graph_header = QHBoxLayout()
        graph_hint = QLabel("Click any block to view/edit student · Dashed lines = cutoffs")
        graph_hint.setStyleSheet(f"color:{TEXT_DIM};font-size:11px;background:transparent;")
        graph_header.addWidget(graph_hint)
        graph_header.addStretch()
        dist_lay.addLayout(graph_header)

        self._graph_widget = InteractiveGraphWidget()
        self._graph_widget.student_clicked.connect(self._on_graph_student_clicked)
        dist_lay.addWidget(self._graph_widget, 1)
        self._graph_tabs.addTab(tab_distribution, "Distribution Blocks")

        tab_normal = QWidget()
        tab_normal.setStyleSheet(f"background: {BG_MAIN};")
        normal_lay = QVBoxLayout(tab_normal)
        normal_lay.setContentsMargins(0, 4, 0, 0)
        normal_lay.setSpacing(4)
        normal_hint = QLabel("Normal distribution curve - dashed lines = mean and cutoffs")
        normal_hint.setStyleSheet(f"color:{TEXT_DIM};font-size:11px;background:transparent;")
        normal_lay.addWidget(normal_hint)
        self._normal_graph_widget = NormalDistributionGraphWidget()
        normal_lay.addWidget(self._normal_graph_widget, 1)
        self._graph_tabs.addTab(tab_normal, "Normal Curve")

        tg_lay.addWidget(self._graph_tabs, 1)

        self._tabs.addTab(tab_graph, "📈 Graph")

        # Tab 5: Edit Log
        tab_editlog = QWidget()
        tab_editlog.setStyleSheet(f"background: {BG_MAIN};")
        te_lay = QVBoxLayout(tab_editlog)
        te_lay.setContentsMargins(4, 4, 4, 4)
        te_lay.setSpacing(4)

        el_header = QHBoxLayout()
        el_title = QLabel("MANUALLY EDITED STUDENTS")
        el_title.setStyleSheet(
            f"color:{TEXT_DIM};font-size:10px;font-weight:bold;"
            f"letter-spacing:1.8px;background:transparent;"
        )
        el_header.addWidget(el_title)
        el_header.addStretch()

        self._btn_clear_log = _small_btn("🗑 Clear Log", height=28)
        self._btn_clear_log.clicked.connect(self._clear_edit_log)
        el_header.addWidget(self._btn_clear_log)
        te_lay.addLayout(el_header)

        self._edit_log_table = QTableWidget()
        self._edit_log_table.setAlternatingRowColors(True)
        self._edit_log_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._edit_log_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._edit_log_table.verticalHeader().setVisible(False)
        self._edit_log_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)
        self._edit_log_table.horizontalHeader().setMinimumSectionSize(80)
        self._edit_log_table.setShowGrid(True)
        self._edit_log_table.setFrameShape(QFrame.NoFrame)
        self._edit_log_table.setColumnCount(6)
        self._edit_log_table.setHorizontalHeaderLabels([
            "STUDENT ID", "OLD TOTAL", "NEW TOTAL",
            "OLD GRADE", "NEW GRADE", "CHANGES DETAIL"
        ])
        te_lay.addWidget(self._edit_log_table, 1)

        self._edit_log_count_label = QLabel("No edits yet")
        self._edit_log_count_label.setStyleSheet(
            f"color:{TEXT_DIM};font-size:11px;background:transparent;"
        )
        te_lay.addWidget(self._edit_log_count_label)

        self._tabs.addTab(tab_editlog, "📝 Edit Log")

        # ── NEW Tab 6: Cluster ───────────────────────────────────────
        tab_cluster = QWidget()
        tab_cluster.setStyleSheet(f"background: {BG_MAIN};")
        tcl_lay = QVBoxLayout(tab_cluster)
        tcl_lay.setContentsMargins(10, 10, 10, 10)
        tcl_lay.setSpacing(8)

        cl_hdr = QHBoxLayout()
        cl_title = QLabel("STUDENT CLUSTERS")
        cl_title.setStyleSheet(
            f"color:{TEXT_DIM};font-size:10px;font-weight:bold;"
            f"letter-spacing:1.8px;background:transparent;"
        )
        cl_hdr.addWidget(cl_title)
        cl_hdr.addStretch()
        cl_refresh = _small_btn("🔄 Refresh", height=28)
        cl_refresh.clicked.connect(self._refresh_cluster_tab)
        cl_hdr.addWidget(cl_refresh)
        tcl_lay.addLayout(cl_hdr)

        self._cluster_tab_table = QTableWidget()
        self._cluster_tab_table.setAlternatingRowColors(True)
        self._cluster_tab_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._cluster_tab_table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._cluster_tab_table.verticalHeader().setVisible(False)
        self._cluster_tab_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self._cluster_tab_table.setShowGrid(True)
        self._cluster_tab_table.setFrameShape(QFrame.NoFrame)
        self._cluster_tab_table.setColumnCount(3)
        self._cluster_tab_table.setHorizontalHeaderLabels(["CLUSTER NAME", "STUDENTS (COUNT)", "ROLL NUMBERS"])
        tcl_lay.addWidget(self._cluster_tab_table, 1)

        #ADD before cl_hint:
        cl_del_row = QHBoxLayout()
        self._btn_delete_cluster = _small_btn("🗑 Delete Selected Cluster", height=30)
        self._btn_delete_cluster.clicked.connect(self._on_delete_cluster_from_tab)
        cl_del_row.addWidget(self._btn_delete_cluster)
        cl_del_row.addStretch()
        tcl_lay.addLayout(cl_del_row) 

        cl_hint = QLabel("Use 🗂 Clusters button (left panel) to create/edit clusters · Use ⚡ Assign Cluster Marks to bulk-assign marks")
        cl_hint.setStyleSheet(f"color:{TEXT_DIM};font-size:11px;background:transparent;")
        cl_hint.setWordWrap(True)
        tcl_lay.addWidget(cl_hint)

        self._tabs.addTab(tab_cluster, "🗂 Clusters")

        lay.addWidget(self._tabs, 1)

        return panel

    # ── CUTOFF MODE ───────────────────────────────────────────────────
    def _on_cutoff_mode_changed(self, button):
        is_custom = self._cutoff_group.checkedId() == 1
        self._use_custom_cutoffs = is_custom
        self._btn_set_cutoffs.setVisible(is_custom)

    def _on_set_custom_cutoffs(self):
        mu = self._frozen_mu if self._frozen_mu is not None else 50.0
        sigma = self._frozen_sigma if self._frozen_sigma is not None else 10.0
        base = self.combo_base.currentText()

        dlg = CustomCutoffDialog(mu, sigma, base, self._thresholds, self)
        if dlg.exec_() == QDialog.Accepted:
            self._custom_thresholds = dlg.get_thresholds()
            QMessageBox.information(
                self, "Custom Cutoffs Set",
                "Custom cutoff values configured.\n"
                "Click 'Calculate Grades' to apply."
            )

    # ── FILE ──────────────────────────────────────────────────────────
    def _on_file(self, path: str):
        try:
            self.df_raw = load_and_clean_data(path)
        except Exception as exc:
            QMessageBox.critical(self, "File Error", f"Could not load:\n\n{exc}")
            return
        fname = path.replace("\\", "/").rsplit("/", 1)[-1]
        self._id_col = self.df_raw.columns[0]

        old = self.file_stack.widget(1)
        new_w = LoadedFileWidget(fname, len(self.df_raw), self._clear_file)
        self.file_stack.removeWidget(old)
        old.deleteLater()
        self.file_stack.insertWidget(1, new_w)
        self.file_stack.setCurrentIndex(1)

        self._load_columns()
        self.btn_apply_team_marks.setEnabled(bool(self.team_comp_rows))
        self.statusBar().showMessage(
            f"Loaded: {fname}  ·  {len(self.df_raw)} valid students"
        )

    def _on_team_file(self, path: str):
        try:
            self.df_team = load_table_file(path)
            team_col, member_cols, marks_cols, numeric_only_cols = detect_team_columns(self.df_team)
            self._team_numeric_only_member = False
            if not member_cols:
                choices = numeric_only_cols or list(self.df_team.columns)
                col, ok = QInputDialog.getItem(
                    self,
                    "Select Members Column",
                    "Roll numbers could not be detected automatically.\nSelect the column containing roll numbers:",
                    choices,
                    0,
                    False,
                )
                if not ok or not col:
                    self.df_team = None
                    return
                member_cols = [col]
                self._team_numeric_only_member = True
        except Exception as exc:
            QMessageBox.critical(self, "Team File Error", f"Could not load team file:\n\n{exc}")
            return

        self._team_path = path
        self._team_name_col = team_col
        self._team_member_cols = member_cols
        self._team_marks_cols = marks_cols

        fname = path.replace("\\", "/").rsplit("/", 1)[-1]
        old = self.team_file_stack.widget(1)
        new_w = LoadedFileWidget(fname, len(self.df_team), self._clear_team_file)
        self.team_file_stack.removeWidget(old)
        old.deleteLater()
        self.team_file_stack.insertWidget(1, new_w)
        self.team_file_stack.setCurrentIndex(1)

        self._load_team_component_rows()
        self.btn_apply_team_marks.setEnabled(bool(self.team_comp_rows) and self.df_raw is not None)

        warnings = self._build_clusters_from_team_file()
        msg = f"Team file loaded: {fname} - {len(self._clusters)} cluster(s) available"
        if warnings:
            msg += f" - {len(warnings)} warning(s)"
        self.statusBar().showMessage(msg)

    def _clear_team_file(self):
        self.df_team = None
        self._team_path = ""
        self._team_name_col = None
        self._team_member_cols = []
        self._team_marks_cols = []
        self._team_roll_marks = {}
        self._team_numeric_only_member = False
        self._clear_team_comp_rows()
        self.btn_apply_team_marks.setEnabled(False)

        old = self.team_file_stack.widget(1)
        ph = QWidget()
        ph.setStyleSheet("background:transparent;border:none;")
        self.team_file_stack.removeWidget(old)
        old.deleteLater()
        self.team_file_stack.insertWidget(1, ph)
        self.team_file_stack.setCurrentIndex(0)
        self.statusBar().showMessage("Team file cleared.")

    def _build_clusters_from_team_file(self) -> list:
        if self.df_team is None:
            return []
        clusters: Dict[str, list] = {}
        warnings = []
        for idx, row in self.df_team.iterrows():
            team_name = str(row.get(self._team_name_col, "")).strip() if self._team_name_col else ""
            if not team_name:
                team_name = f"Team {idx + 1}"
            rolls = []
            seen = set()
            for col in self._team_member_cols:
                for roll in extract_roll_numbers(row.get(col, ""), self._team_numeric_only_member):
                    if roll not in seen:
                        seen.add(roll)
                        rolls.append(roll)
            if not rolls:
                warnings.append(f"{team_name}: no roll numbers found")
                continue
            clusters[team_name] = rolls

        if clusters:
            self._clusters.update(clusters)
            save_clusters(self._clusters)
            self._refresh_cluster_tab()
        return warnings

    def _on_apply_team_marks(self):
        if self.df_raw is None:
            QMessageBox.warning(self, "Main File Needed", "Load the main student file before merging team marks.")
            return
        if self.df_team is None or not self.team_comp_rows:
            QMessageBox.warning(self, "No Team Marks", "Load a team file with a marks column first.")
            return

        selected = []
        for rw in self.team_comp_rows:
            cfg = rw.get_config()
            if cfg is None:
                continue
            col, mr, wr = cfg['column'], cfg['max_raw'], cfg['weight_raw']
            try:
                max_marks = float(mr)
                weightage = float(wr)
                if max_marks <= 0 or weightage <= 0:
                    raise ValueError
            except ValueError:
                QMessageBox.warning(
                    self,
                    "Invalid Team Component",
                    f"Enter positive Max and Weight values for '{col}'."
                )
                return
            selected.append((col, max_marks, weightage))

        if not selected:
            QMessageBox.warning(self, "No Team Components", "Select at least one team marks component.")
            return

        summaries = []
        for marks_col, max_marks, weightage in selected:
            summary = self._merge_single_team_marks_component(marks_col, max_marks, weightage)
            summaries.append(summary)

        self.btn_calc.setEnabled(True)
        self.btn_add_comp.setEnabled(True)
        self._refresh_weight()
        self._unsaved_changes = True
        QMessageBox.information(self, "Team Marks Merged", "\n".join(summaries))
        self.statusBar().showMessage(f"Merged {len(summaries)} team component(s).")

    def _merge_single_team_marks_component(self, marks_col: str, max_marks: float, weightage: float) -> str:
        component_col = marks_col if marks_col not in self.df_raw.columns else f"Team {marks_col}"
        roll_to_mark: Dict[str, float] = {}
        duplicate_kept = 0
        for _, row in self.df_team.iterrows():
            mark = pd.to_numeric(row.get(marks_col, ""), errors='coerce')
            if pd.isna(mark):
                continue
            mark = float(mark)
            rolls = []
            for col in self._team_member_cols:
                rolls.extend(extract_roll_numbers(row.get(col, ""), self._team_numeric_only_member))
            for roll in rolls:
                old = roll_to_mark.get(roll)
                if old is None or mark > old:
                    if old is not None:
                        duplicate_kept += 1
                    roll_to_mark[roll] = mark

        id_col = self._id_col or self.df_raw.columns[0]
        if component_col not in self.df_raw.columns:
            self.df_raw[component_col] = "0"
        col_idx = self.df_raw.columns.get_loc(component_col)
        loaded_ids = {str(v).strip().upper() for v in self.df_raw[id_col].astype(str)}
        matched = 0
        for i in range(len(self.df_raw)):
            sid = str(self.df_raw.iat[i, self.df_raw.columns.get_loc(id_col)]).strip().upper()
            if sid in roll_to_mark:
                self.df_raw.iat[i, col_idx] = str(roll_to_mark[sid])
                matched += 1
        unmatched = sorted(set(roll_to_mark) - loaded_ids)

        if component_col not in self._numeric_cols:
            self._numeric_cols.append(component_col)
        self._add_or_update_component_row(component_col, max_marks, weightage)

        detail = f"{matched} student(s) matched"
        if unmatched:
            detail += f"; {len(unmatched)} unmatched skipped"
        if duplicate_kept:
            detail += f"; {duplicate_kept} duplicate(s) used higher marks"
        return f"'{component_col}': {detail}."

    def _add_or_update_component_row(self, col_name: str, max_marks: float, weightage: float):
        for rw in self.comp_rows:
            if rw.column_name == col_name:
                rw.checkbox.setChecked(True)
                rw.max_input.setText(str(int(max_marks)) if max_marks == int(max_marks) else str(max_marks))
                rw.weight_input.setText(str(int(weightage)) if weightage == int(weightage) else str(weightage))
                return
        rw = ComponentRow(col_name)
        rw.max_input.setText(str(int(max_marks)) if max_marks == int(max_marks) else str(max_marks))
        rw.weight_input.setText(str(int(weightage)) if weightage == int(weightage) else str(weightage))
        rw.max_input.textChanged.connect(self._refresh_weight)
        rw.weight_input.textChanged.connect(self._refresh_weight)
        rw.checkbox.stateChanged.connect(self._refresh_weight)
        self.scroll_layout.insertWidget(self.scroll_layout.count() - 1, rw)
        self.comp_rows.append(rw)

    def _clear_file(self):
        self.df_raw = self.df_graded = None
        self._frozen_mu = None
        self._frozen_sigma = None
        self._thresholds = []
        self._cutoff_ranges = {}
        self._custom_thresholds = None
        self._components = []
        self._numeric_cols = []
        self._edit_log = []
        self._unsaved_changes = False  # NEW
        if self.df_team is None:
            self._clusters = {}
            save_clusters({})
            self._refresh_cluster_tab()
        else:
            self._build_clusters_from_team_file()

        self._clear_comp_rows()
        self._clear_table()
        self._clear_badges()
        self._graph_widget.clear_data()
        self._normal_graph_widget.clear_data()
        self._populate_edit_log_table()

        for c in (self.card_mean, self.card_std, self.card_count,
                  self.card_min, self.card_max, self.card_weight):
            c.set_value("—")
        self.btn_calc.setEnabled(False)
        self.btn_export.setEnabled(False)
        self.btn_add_comp.setEnabled(False)      # NEW
        self.btn_cluster_marks.setEnabled(False) # NEW
        self.btn_apply_team_marks.setEnabled(False)
        self._tw_lbl.setText("0.0")
        self._sel_count.setText("")

        old = self.file_stack.widget(1)
        ph = QWidget()
        ph.setStyleSheet("background:transparent;border:none;")
        self.file_stack.removeWidget(old)
        old.deleteLater()
        self.file_stack.insertWidget(1, ph)
        self.file_stack.setCurrentIndex(0)
        self.statusBar().showMessage("File cleared — upload a new file.")

    def _load_columns(self):
        if self.df_raw is None:
            return
        self._clear_comp_rows()
        cols = get_numeric_columns(self.df_raw)
        self._numeric_cols = cols
        if not cols:
            QMessageBox.warning(self, "No Numeric Columns", "No numeric columns found.")
            return
        for col in cols:
            rw = ComponentRow(col)
            rw.max_input.textChanged.connect(self._refresh_weight)
            rw.weight_input.textChanged.connect(self._refresh_weight)
            rw.checkbox.stateChanged.connect(self._refresh_weight)
            self.scroll_layout.insertWidget(
                self.scroll_layout.count() - 1, rw
            )
            self.comp_rows.append(rw)
        self.btn_calc.setEnabled(True)
        self._refresh_weight()

    def _refresh_weight(self):
        total, sel = 0.0, 0
        for rw in self.comp_rows:
            if not rw.checkbox.isChecked():
                continue
            try:
                total += float(rw.weight_input.text().strip())
                sel += 1
            except ValueError:
                pass
        self._tw_lbl.setText(f"{total:.1f}")
        color = GOLD if abs(total - 100.0) < 0.01 else TEXT_PRI
        self._tw_lbl.setStyleSheet(
            f"color:{color};font-size:20px;font-weight:bold;background:transparent;"
        )
        self._sel_count.setText(f"{sel} selected" if sel else "")

    # ── CALCULATE ─────────────────────────────────────────────────────
    def action_calculate(self):
        comps = self._validate()
        if comps is None:
            return
        self._components = comps
        base = self.combo_base.currentText()

        try:
            df_t = calculate_totals(self.df_raw, comps)

            if self._use_custom_cutoffs and self._custom_thresholds:
                totals = df_t['Total'].astype(float)
                mu, sigma = totals.mean(), totals.std()
                if sigma == 0:
                    sigma = 0.001
                self._thresholds = self._custom_thresholds
                df_g = assign_grades_custom(df_t, self._custom_thresholds)
            else:
                df_g, mu, sigma, thresholds = assign_grades(df_t, base)
                self._thresholds = thresholds
        except Exception as exc:
            QMessageBox.critical(self, "Calculation Error",
                                 f"{type(exc).__name__}: {exc}\n\n{traceback.format_exc()}")
            return

        self.df_graded = df_g
        self._frozen_mu = mu
        self._frozen_sigma = sigma
        self._cutoff_ranges = compute_cutoff_ranges(self._thresholds)

        # Add Rank column
        self.df_graded['Rank'] = self.df_graded['Total'].astype(float).rank(
            ascending=False, method='min'
        ).astype(int)

        # Clear edit log on new calculation
        self._edit_log = []
        self._populate_edit_log_table()
        self._unsaved_changes = True  # NEW: mark dirty

        totals = df_g['Total'].astype(float)

        self.card_mean.set_value(f"{mu:.2f}")
        self.card_std.set_value(f"{sigma:.2f}")
        self.card_count.set_value(str(len(df_g)))
        self.card_min.set_value(f"{totals.min():.2f}")
        self.card_max.set_value(f"{totals.max():.2f}")
        self.card_weight.set_value(f"{sum(c['weightage'] for c in comps):.1f}")

        self._populate_badges(df_g)
        self._populate_table(df_g)
        self._populate_ranking_table(df_g)
        self._cutoff_display.set_cutoffs(self._cutoff_ranges)
        self._update_graph()
        self.btn_export.setEnabled(True)
        self.btn_add_comp.setEnabled(True)      # NEW
        self.btn_cluster_marks.setEnabled(True) # NEW

        dist = df_g['Grade'].value_counts().to_dict()
        self.statusBar().showMessage(
            f"Done  ·  μ={mu:.2f}  σ={sigma:.2f}  ·  Base:{base}  ·  n={len(df_g)}  ·  "
            + "  ".join(f"{g}:{n}" for g, n in sorted(dist.items()))
        )

    # ── EXPORT ────────────────────────────────────────────────────────
    def action_export(self):
        if self.df_graded is None:
            return

        # ── Feature 3: let user pick which columns to include ──
        all_cols = list(self.df_graded.columns)
        col_dlg = ColumnSelectDialog(all_cols, self)
        if col_dlg.exec_() != QDialog.Accepted:
            return
        selected_cols = col_dlg.get_selected_columns()
        if not selected_cols:
            QMessageBox.warning(self, "No Columns",
                                "Please select at least one column to export.")
            return

        save_path, _ = QFileDialog.getSaveFileName(
            self, "Save Output", "graded_output.xlsx", "Excel Files (*.xlsx)"
        )
        if not save_path:
            return
        if not save_path.lower().endswith('.xlsx'):
            save_path += '.xlsx'

        try:
            id_col = self._id_col or self.df_graded.columns[0]
            # Feature 2 fix: pass selected_cols; save_styled_excel now writes
            # numeric columns as proper Python int/float values.
            save_styled_excel(self.df_graded, save_path, id_col,
                              cols_to_export=selected_cols)
        except Exception as exc:
            QMessageBox.critical(self, "Export Error",
                                 f"Could not save:\n\n{exc}")
            return
        QMessageBox.information(self, "Exported",
                                f"Styled Excel saved:\n{save_path}")
        self.statusBar().showMessage(f"Exported to {save_path}")

    # ── VALIDATION ────────────────────────────────────────────────────
    def _validate(self) -> Optional[list]:
        comps, total_w = [], 0.0
        for rw in self.comp_rows:
            cfg = rw.get_config()
            if cfg is None:
                continue
            col, mr, wr = cfg['column'], cfg['max_raw'], cfg['weight_raw']
            if not mr:
                QMessageBox.warning(self, "Missing", f"Max Marks needed for '{col}'")
                return None
            try:
                mm = float(mr)
                if mm <= 0:
                    raise ValueError
            except ValueError:
                QMessageBox.warning(self, "Invalid", f"Max Marks must be >0 for '{col}'")
                return None
            if not wr:
                QMessageBox.warning(self, "Missing", f"Weightage needed for '{col}'")
                return None
            try:
                wt = float(wr)
                if wt <= 0:
                    raise ValueError
            except ValueError:
                QMessageBox.warning(self, "Invalid", f"Weightage must be >0 for '{col}'")
                return None
            total_w += wt
            comps.append({'column': col, 'max_marks': mm, 'weightage': wt})
        if not comps:
            QMessageBox.warning(self, "No Components", "Select at least one component.")
            return None
        if abs(total_w - 100.0) > 0.01:
            r = QMessageBox.question(
                self, "Weight Mismatch",
                f"Total weightage is {total_w:.2f} (expected 100).\n\nContinue?",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
            )
            if r == QMessageBox.No:
                return None
        return comps

    # ── SEARCH ────────────────────────────────────────────────────────
    def _on_search(self, text: str):
        if self.df_graded is None:
            return
        self._populate_table(self.df_graded, text.strip())

    # ── TABLE ─────────────────────────────────────────────────────────
    def _populate_table(self, df: pd.DataFrame, filter_text: str = ""):
        display_df = df
        id_col = self._id_col or df.columns[0]

        if filter_text:
            mask = display_df[id_col].astype(str).str.contains(
                filter_text, case=False, na=False
            )
            display_df = display_df[mask]

        self.result_table.setSortingEnabled(False)
        self.result_table.clear()
        self.result_table.setRowCount(len(display_df))

        show_cols = list(display_df.columns)
        self.result_table.setColumnCount(len(show_cols))
        self.result_table.setHorizontalHeaderLabels(
            [c.upper() for c in show_cols]
        )
        bold_f = QFont('Segoe UI', 11, 75)
        def_f  = QFont('Segoe UI', 11)

        for r, (_, row) in enumerate(display_df.iterrows()):
            self.result_table.setRowHeight(r, 34)
            for c, col in enumerate(show_cols):
                val = row[col]
                if col == 'Total':
                    try:
                        item = QTableWidgetItem()
                        item.setData(Qt.DisplayRole, round(float(val), 2))
                    except (ValueError, TypeError):
                        item = QTableWidgetItem(str(val))
                    item.setForeground(QColor(GOLD))
                    item.setFont(bold_f)
                elif col == 'Grade':
                    item = QTableWidgetItem(str(val))
                    item.setForeground(QColor(GRADE_FG.get(str(val), WHITE)))
                    item.setFont(bold_f)
                elif col == 'Rank':
                    rank_val = 999
                    try:
                        rank_val = int(float(val))
                    except (ValueError, TypeError):
                        pass
                    medals = {1: "🥇", 2: "🥈", 3: "🥉"}
                    display_text = f"{medals.get(rank_val, '')} {rank_val}" if rank_val <= 3 else str(rank_val)
                    item = QTableWidgetItem(display_text)
                    item.setData(Qt.UserRole, rank_val)
                    item.setFont(bold_f)
                    if rank_val == 1:
                        item.setForeground(QColor("#FFD700"))
                    elif rank_val == 2:
                        item.setForeground(QColor("#C0C0C0"))
                    elif rank_val == 3:
                        item.setForeground(QColor("#CD7F32"))
                    else:
                        item.setForeground(QColor(TEXT_PRI))
                else:
                    item = QTableWidgetItem(str(val))
                    item.setFont(def_f)
                    item.setForeground(QColor(TEXT_PRI))
                item.setTextAlignment(Qt.AlignCenter)
                self.result_table.setItem(r, c, item)

        self.result_table.resizeColumnsToContents()
        for c in range(self.result_table.columnCount()):
            if self.result_table.columnWidth(c) < 100:
                self.result_table.setColumnWidth(c, 100)

    def _populate_ranking_table(self, df: pd.DataFrame):
        id_col = self._id_col or df.columns[0]
        ranked = df.sort_values('Rank', ascending=True).reset_index(drop=True)
        show_cols = ['Rank', id_col, 'Total', 'Grade']

        self.ranking_table.setSortingEnabled(False)
        self.ranking_table.clear()
        self.ranking_table.setRowCount(len(ranked))
        self.ranking_table.setColumnCount(len(show_cols))
        self.ranking_table.setHorizontalHeaderLabels(
            [c.upper() for c in show_cols]
        )

        bold_f = QFont('Segoe UI', 11, 75)
        big_f  = QFont('Segoe UI', 13, 75)

        for r, (_, row) in enumerate(ranked.iterrows()):
            self.ranking_table.setRowHeight(r, 36)
            for c, col in enumerate(show_cols):
                val = row.get(col, '')

                if col == 'Rank':
                    rank_val = 999
                    try:
                        rank_val = int(float(val))
                    except (ValueError, TypeError):
                        pass
                    medals = {1: "🥇", 2: "🥈", 3: "🥉"}
                    display = f"{medals.get(rank_val, '')}  {rank_val}" if rank_val <= 3 else str(rank_val)
                    item = QTableWidgetItem(display)
                    item.setFont(big_f)
                    if rank_val == 1:
                        item.setForeground(QColor("#FFD700"))
                    elif rank_val == 2:
                        item.setForeground(QColor("#C0C0C0"))
                    elif rank_val == 3:
                        item.setForeground(QColor("#CD7F32"))
                    else:
                        item.setForeground(QColor(TEXT_PRI))
                elif col == 'Total':
                    try:
                        item = QTableWidgetItem()
                        item.setData(Qt.DisplayRole, round(float(val), 2))
                    except (ValueError, TypeError):
                        item = QTableWidgetItem(str(val))
                    item.setForeground(QColor(GOLD))
                    item.setFont(bold_f)
                elif col == 'Grade':
                    grade = str(val)
                    item = QTableWidgetItem(grade)
                    item.setForeground(QColor(GRADE_FG.get(grade, WHITE)))
                    item.setFont(bold_f)
                else:
                    item = QTableWidgetItem(str(val))
                    item.setFont(QFont('Segoe UI', 11))
                    item.setForeground(QColor(TEXT_PRI))

                item.setTextAlignment(Qt.AlignCenter)
                self.ranking_table.setItem(r, c, item)

        self.ranking_table.resizeColumnsToContents()
        for c in range(self.ranking_table.columnCount()):
            if self.ranking_table.columnWidth(c) < 100:
                self.ranking_table.setColumnWidth(c, 100)

    def _clear_table(self):
        self.result_table.clear()
        self.result_table.setRowCount(0)
        self.result_table.setColumnCount(0)
        self.ranking_table.clear()
        self.ranking_table.setRowCount(0)
        self.ranking_table.setColumnCount(0)

    # ── GRAPH ─────────────────────────────────────────────────────────
    def _update_graph(self):
        if self.df_graded is None:
            self._graph_widget.clear_data()
            self._normal_graph_widget.clear_data()
            return

        id_col = self._id_col or self.df_graded.columns[0]
        data_list = []
        for row_idx, (_, row) in enumerate(self.df_graded.iterrows()):
            total = float(row.get('Total', 0))
            grade = str(row.get('Grade', 'F'))
            id_str = str(row.get(id_col, ''))
            data_list.append((total, grade, id_str, row_idx))

        self._graph_widget.set_data(data_list, self._cutoff_ranges, self._thresholds)
        self._normal_graph_widget.set_data(
            data_list,
            self._thresholds,
            self._frozen_mu,
            self._frozen_sigma,
        )

    def _on_graph_student_clicked(self, row_idx: int):
        if self.df_graded is None:
            return
        if row_idx >= len(self.df_graded):
            return
        self._open_edit_dialog(row_idx)

    # ── EDIT ──────────────────────────────────────────────────────────
    def _on_table_double_click(self, index):
        if self.df_graded is None:
            return

        filter_text = self._search_input.text().strip()
        if filter_text:
            id_col = self._id_col or self.df_graded.columns[0]
            mask = self.df_graded[id_col].astype(str).str.contains(
                filter_text, case=False, na=False
            )
            display_df = self.df_graded[mask]
            if index.row() >= len(display_df):
                return
            actual_idx = display_df.index[index.row()]
            row_idx = list(self.df_graded.index).index(actual_idx)
        else:
            row_idx = index.row()

        if row_idx >= len(self.df_graded):
            return
        self._open_edit_dialog(row_idx)

    def _open_edit_dialog(self, row_idx: int):
        """Open edit dialog for the student at positional row_idx."""
        if self.df_graded is None or self.df_raw is None:
            return
        if row_idx < 0 or row_idx >= len(self.df_graded):
            return

        try:
            id_col = self._id_col or self.df_graded.columns[0]
            student_row = self.df_graded.iloc[row_idx]
            student_id = str(student_row[id_col])
            current_grade = str(student_row.get('Grade', 'F'))
            current_total = float(student_row.get('Total', 0))

            # Get raw marks safely
            marks = {}
            if row_idx < len(self.df_raw):
                raw_row = self.df_raw.iloc[row_idx]
                for col in self._numeric_cols:
                    if col == 'Total':
                        continue
                    try:
                        val = raw_row.get(col, 0)
                        marks[col] = float(val) if val is not None and str(val).strip() != '' else 0.0
                    except (ValueError, TypeError):
                        marks[col] = 0.0

            dlg = EditStudentDialog(
                student_id, marks, current_grade, current_total,
                self._numeric_cols, self._components, self
            )

            if dlg.exec_() == QDialog.Accepted:
                new_marks = dlg.get_marks()
                grade_override = dlg.get_grade_override()

                # Track what changed for the edit log
                old_total = current_total
                old_grade = current_grade
                changes_made = {}

                # Update raw marks safely
                for col, val in new_marks.items():
                    if col in self.df_raw.columns and row_idx < len(self.df_raw):
                        old_val = 0.0
                        try:
                            raw_v = self.df_raw.iloc[row_idx][col]
                            old_val = float(raw_v) if raw_v is not None and str(raw_v).strip() != '' else 0.0
                        except (ValueError, TypeError):
                            old_val = 0.0
                        if abs(old_val - val) > 0.001:
                            changes_made[col] = (old_val, val)
                        self.df_raw.iat[row_idx, self.df_raw.columns.get_loc(col)] = str(val)

                # Recompute total for this student using components
                new_total = 0.0
                for comp in self._components:
                    col_name = comp['column']
                    if col_name in new_marks:
                        raw_val = new_marks[col_name]
                    elif row_idx < len(self.df_raw) and col_name in self.df_raw.columns:
                        try:
                            rv = self.df_raw.iloc[row_idx][col_name]
                            raw_val = float(rv) if rv is not None and str(rv).strip() != '' else 0.0
                        except (ValueError, TypeError):
                            raw_val = 0.0
                    else:
                        raw_val = 0.0
                    new_total += (raw_val / comp['max_marks']) * comp['weightage']
                new_total = round(new_total, 2)

                # Update graded df using actual index label
                df_idx = self.df_graded.index[row_idx]
                self.df_graded.at[df_idx, 'Total'] = new_total

                # Safely update raw column values in graded df
                for col, val in new_marks.items():
                    if col in self.df_graded.columns:
                        try:
                            self.df_graded.at[df_idx, col] = val
                        except Exception:
                            pass

                # Determine grade (frozen mean/std, frozen thresholds)
                if grade_override:
                    self.df_graded.at[df_idx, 'Grade'] = grade_override
                else:
                    new_grade = grade_single_student(new_total, self._thresholds)
                    self.df_graded.at[df_idx, 'Grade'] = new_grade

                final_grade = str(self.df_graded.at[df_idx, 'Grade'])

                # Re-rank
                self.df_graded['Rank'] = self.df_graded['Total'].astype(float).rank(
                    ascending=False, method='min'
                ).astype(int)

                # Log the edit if anything changed
                if abs(old_total - new_total) > 0.001 or old_grade != final_grade or changes_made:
                    edit_entry = {
                        'student_id': student_id,
                        'old_total': old_total,
                        'new_total': new_total,
                        'old_grade': old_grade,
                        'new_grade': final_grade,
                        'mark_changes': changes_made,
                        'grade_override': grade_override is not None,
                    }
                    self._edit_log.append(edit_entry)
                    self._populate_edit_log_table()
                    self._unsaved_changes = True  # NEW: mark dirty

                # Refresh UI
                self._refresh_after_edit()

        except Exception as exc:
            QMessageBox.critical(
                self, "Edit Error",
                f"An error occurred while editing:\n\n{type(exc).__name__}: {exc}\n\n{traceback.format_exc()}"
            )

    def _refresh_after_edit(self):
        """Refresh all UI after an edit, keeping frozen stats."""
        if self.df_graded is None:
            return

        self._populate_badges(self.df_graded)

        totals = self.df_graded['Total'].astype(float)
        self.card_min.set_value(f"{totals.min():.2f}")
        self.card_max.set_value(f"{totals.max():.2f}")

        filter_text = self._search_input.text().strip()
        self._populate_table(self.df_graded, filter_text)
        self._populate_ranking_table(self.df_graded)
        self._update_graph()

        self.statusBar().showMessage(
            f"Student edited  ·  μ={self._frozen_mu:.2f} (frozen)  "
            f"σ={self._frozen_sigma:.2f} (frozen)  ·  n={len(self.df_graded)}"
        )

    # ── EDIT LOG ──────────────────────────────────────────────────────
    def _populate_edit_log_table(self):
        """Refresh the edit log table from self._edit_log."""
        tbl = self._edit_log_table
        tbl.setRowCount(len(self._edit_log))

        bold_f = QFont('Segoe UI', 11, 75)
        def_f  = QFont('Segoe UI', 11)

        for r, entry in enumerate(self._edit_log):
            tbl.setRowHeight(r, 36)

            # Student ID
            item_id = QTableWidgetItem(entry['student_id'])
            item_id.setFont(bold_f)
            item_id.setForeground(QColor(TEXT_PRI))
            item_id.setTextAlignment(Qt.AlignCenter)
            tbl.setItem(r, 0, item_id)

            # Old Total
            item_ot = QTableWidgetItem(f"{entry['old_total']:.2f}")
            item_ot.setFont(def_f)
            item_ot.setForeground(QColor(TEXT_SEC))
            item_ot.setTextAlignment(Qt.AlignCenter)
            tbl.setItem(r, 1, item_ot)

            # New Total
            item_nt = QTableWidgetItem(f"{entry['new_total']:.2f}")
            item_nt.setFont(bold_f)
            changed_total = abs(entry['old_total'] - entry['new_total']) > 0.001
            item_nt.setForeground(QColor(GOLD if changed_total else TEXT_PRI))
            item_nt.setTextAlignment(Qt.AlignCenter)
            tbl.setItem(r, 2, item_nt)

            # Old Grade
            old_g = entry['old_grade']
            item_og = QTableWidgetItem(old_g)
            item_og.setFont(def_f)
            item_og.setForeground(QColor(GRADE_FG.get(old_g, TEXT_SEC)))
            item_og.setTextAlignment(Qt.AlignCenter)
            tbl.setItem(r, 3, item_og)

            # New Grade
            new_g = entry['new_grade']
            item_ng = QTableWidgetItem(new_g)
            item_ng.setFont(bold_f)
            item_ng.setForeground(QColor(GRADE_FG.get(new_g, TEXT_PRI)))
            item_ng.setTextAlignment(Qt.AlignCenter)
            tbl.setItem(r, 4, item_ng)

            # Changes Detail
            details = []
            for col, (old_v, new_v) in entry.get('mark_changes', {}).items():
                details.append(f"{col}: {old_v:.1f} → {new_v:.1f}")
            if entry.get('grade_override'):
                details.append("Grade manually overridden")
            if not details:
                details.append("No mark changes")
            detail_text = " | ".join(details)

            item_detail = QTableWidgetItem(detail_text)
            item_detail.setFont(QFont('Segoe UI', 10))
            has_real_changes = details and details[0] != "No mark changes"
            item_detail.setForeground(QColor("#e07b10" if has_real_changes else TEXT_DIM))
            item_detail.setTextAlignment(Qt.AlignLeft | Qt.AlignVCenter)
            tbl.setItem(r, 5, item_detail)

        tbl.resizeColumnsToContents()

        count = len(self._edit_log)
        if count == 0:
            self._edit_log_count_label.setText("No edits yet")
        else:
            self._edit_log_count_label.setText(
                f"{count} edit(s) recorded  ·  "
                f"Mean & Std Dev remain frozen from original calculation"
            )

    def _clear_edit_log(self):
        """Clear the edit log."""
        if not self._edit_log:
            return
        r = QMessageBox.question(
            self, "Clear Edit Log",
            f"Clear all {len(self._edit_log)} edit record(s)?\n\n"
            "This only clears the log — edits to marks/grades remain.",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No,
        )
        if r == QMessageBox.Yes:
            self._edit_log.clear()
            self._populate_edit_log_table()

    # ── NEW: SAVE SESSION WITH CUSTOM NAME ──────────────────────────────
    def _action_save_session(self) -> bool:
        """
        Prompt for a custom session name, then save.
        Returns True if saved successfully, False if the user cancelled.
        """
        if self.df_graded is None:
            QMessageBox.warning(self, "Nothing to Save",
                                "Please calculate grades first.")
            return False

        name, ok = QInputDialog.getText(
            self, "Save Session",
            "Enter a name for this session\n(leave blank to use timestamp):",
            QLineEdit.Normal, ""
        )
        if not ok:
            return False  # User pressed Cancel

        saved = save_current_session(self, custom_name=name.strip())
        if saved:
            self._unsaved_changes = False
            display = name.strip() if name.strip() else "auto-named"
            self.statusBar().showMessage(
                f"Session saved: '{display}'  →  {APP_DATA_DIR}"
            )
            QMessageBox.information(
                self, "Saved",
                f"Session saved successfully.\n\nLocation: {APP_DATA_DIR}"
            )
        else:
            QMessageBox.warning(
                self, "Save Failed",
                "Could not save the session.\n"
                "Check disk space / permissions."
            )
        return saved

    # ── NEW: MANAGE HISTORY (DELETE) ─────────────────────────────────────
    def _on_manage_history(self):
        """Open the Manage History dialog so the user can delete sessions."""
        sessions = load_sessions_index()
        if not sessions:
            QMessageBox.information(
                self, "No History",
                f"No saved sessions found.\n\nStorage folder:\n{APP_DATA_DIR}"
            )
            return
        dlg = ManageHistoryDialog(sessions, self)
        dlg.exec_()

    # ── NEW: CLUSTER MANAGER ─────────────────────────────────────────────
    def _on_cluster_manager(self):
        """Open the Cluster Manager dialog (create / edit / delete clusters)."""
        all_ids: list = []
        if self.df_graded is not None:
            id_col = self._id_col or self.df_graded.columns[0]
            all_ids = self.df_graded[id_col].astype(str).tolist()

        dlg = ClusterManagerDialog(
            self._clusters, all_ids, self._components, self
        )
        dlg.exec_()
        # Persist any changes the dialog made
        self._clusters = dlg.get_clusters()
        self._refresh_cluster_tab()

    # ── NEW: BULK ASSIGN MARKS TO A CLUSTER ──────────────────────────────
    def _on_cluster_bulk_marks(self):
        """Assign a single mark value for a component to every student in a cluster."""
        if not self._clusters:
            QMessageBox.warning(
                self, "No Clusters",
                "Create clusters first via 🗂 Clusters."
            )
            return
        if not self._components:
            QMessageBox.warning(
                self, "No Components",
                "Calculate grades first so that components are defined."
            )
            return

        dlg = ClusterBulkMarksDialog(self._clusters, self._components, self)
        if dlg.exec_() != QDialog.Accepted:
            return
        result = dlg.get_result()
        if not result:
            return

        cluster_name = result["cluster_name"]
        comp         = result["component"]
        mark         = result["mark"]
        roll_numbers = set(self._clusters.get(cluster_name, []))
        col_name     = comp["column"]
        max_marks    = comp["max_marks"]

        if mark > max_marks:
            QMessageBox.warning(
                self, "Invalid Mark",
                f"Mark {mark} exceeds maximum {max_marks} for '{col_name}'."
            )
            return

        if self.df_raw is None or self.df_graded is None:
            QMessageBox.warning(self, "No Data",
                                "Load a file and calculate grades first.")
            return

        id_col = self._id_col or self.df_raw.columns[0]
        updated = 0
        for i in range(len(self.df_raw)):
            sid = str(self.df_raw.iat[i, self.df_raw.columns.get_loc(id_col)]).strip()
            if sid in roll_numbers and col_name in self.df_raw.columns:
                self.df_raw.iat[i, self.df_raw.columns.get_loc(col_name)] = str(mark)
                updated += 1

        if updated == 0:
            QMessageBox.warning(
                self, "No Matches",
                f"No students in cluster '{cluster_name}' matched the loaded data.\n"
                "Check that roll numbers match exactly."
            )
            return

        # Recompute totals and grades for the entire dataset
        try:
            df_t = calculate_totals(self.df_raw, self._components)
            if self._use_custom_cutoffs and self._custom_thresholds:
                df_g = assign_grades_custom(df_t, self._custom_thresholds)
            else:
                df_g = assign_grades_custom(df_t, self._thresholds)
            df_g['Rank'] = (
                df_g['Total'].astype(float)
                .rank(ascending=False, method='min')
                .astype(int)
            )
            self.df_graded = df_g
            self._unsaved_changes = True
            self._refresh_after_edit()
            self.statusBar().showMessage(
                f"Cluster '{cluster_name}': {updated} student(s) updated  ·  "
                f"{col_name} = {mark}"
            )
        except Exception as exc:
            QMessageBox.critical(
                self, "Recalculation Error",
                f"Failed to recompute grades:\n\n{exc}\n\n{traceback.format_exc()}"
            )

    # ── NEW: REFRESH CLUSTER TAB ─────────────────────────────────────────
    def _refresh_cluster_tab(self):
        """Populate the Clusters tab table from self._clusters."""
        tbl = self._cluster_tab_table
        clusters = self._clusters
        tbl.setRowCount(len(clusters))
        bold_f = QFont('Segoe UI', 11, 75)
        sec_f  = QFont('Segoe UI', 11)

        for r, (name, rolls) in enumerate(clusters.items()):
            tbl.setRowHeight(r, 34)

            n_item = QTableWidgetItem(name)
            n_item.setFont(bold_f)
            n_item.setForeground(QColor(GOLD))
            n_item.setTextAlignment(Qt.AlignCenter)
            tbl.setItem(r, 0, n_item)

            c_item = QTableWidgetItem(str(len(rolls)))
            c_item.setFont(sec_f)
            c_item.setForeground(QColor(TEXT_PRI))
            c_item.setTextAlignment(Qt.AlignCenter)
            tbl.setItem(r, 1, c_item)

            r_item = QTableWidgetItem(", ".join(rolls))
            r_item.setFont(sec_f)
            r_item.setForeground(QColor(TEXT_SEC))
            tbl.setItem(r, 2, r_item)

        tbl.resizeColumnsToContents()

    def _on_delete_cluster_from_tab(self):
            """Delete the selected cluster row directly from the Clusters tab."""
            selected = self._cluster_tab_table.selectedItems()
            if not selected:
                QMessageBox.information(self, "None Selected", "Click a cluster row to select it first.")
                return
            row = self._cluster_tab_table.currentRow()
            name_item = self._cluster_tab_table.item(row, 0)
            if not name_item:   
                return
            name = name_item.text()
            if QMessageBox.question(self, "Delete Cluster", 
                                    f"Delete cluster '{name}'?",
                                    QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
                self._clusters.pop(name, None)
                save_clusters(self._clusters)
                self._refresh_cluster_tab()
                self.statusBar().showMessage(f"Cluster '{name}' deleted.")

    # ── NEW: ADD COMPONENT DYNAMICALLY ───────────────────────────────────
    def _on_add_component(self):
        """
        Let the user define a brand-new grading component column (e.g. 'Presentation').
        The column is added to df_raw with 0 marks for every student.
        A ComponentRow is inserted in the sidebar so it participates in grading.
        """
        if self.df_raw is None:
            QMessageBox.warning(self, "No Data", "Please load a file first.")
            return

        existing_cols = list(self.df_raw.columns)
        dlg = AddComponentDialog(existing_cols, self)
        if dlg.exec_() != QDialog.Accepted:
            return

        info      = dlg.get_component_info()
        col_name  = info["name"]
        max_marks = info["max_marks"]
        weightage = info["weightage"]

        # Add the column to df_raw with default mark = 0
        self.df_raw[col_name] = "0"
        if col_name not in self._numeric_cols:
            self._numeric_cols.append(col_name)

        # Create a ComponentRow pre-filled with the supplied values
        rw = ComponentRow(col_name)
        mm_str = str(int(max_marks)) if max_marks == int(max_marks) else str(max_marks)
        wt_str = str(int(weightage)) if weightage == int(weightage) else str(weightage)
        rw.max_input.setText(mm_str)
        rw.weight_input.setText(wt_str)
        rw.max_input.textChanged.connect(self._refresh_weight)
        rw.weight_input.textChanged.connect(self._refresh_weight)
        rw.checkbox.stateChanged.connect(self._refresh_weight)
        # Insert before the trailing stretch item
        self.scroll_layout.insertWidget(self.scroll_layout.count() - 1, rw)
        self.comp_rows.append(rw)
        self._refresh_weight()

        self.statusBar().showMessage(
            f"New component '{col_name}' added  ·  "
            "Assign marks then click ⊞ Calculate Grades"
        )
        QMessageBox.information(
            self, "Component Added",
            f"'{col_name}' added (max={max_marks}, weight={weightage}).\n\n"
            "• Use ⚡ Assign Cluster Marks to bulk-set marks for a group.\n"
            "• Or double-click individual rows to set marks per student.\n"
            "• Click ⊞ Calculate Grades to recompute everything."
        )

    # ── HISTORY / SESSION PERSISTENCE  (NEW — Feature 1) ──────────────

    def _maybe_offer_resume(self):
        """
        Called once at startup (via QTimer).  If a previous session exists,
        ask the user whether they want to resume it.
        """
        sessions = load_sessions_index()
        if not sessions:
            return
        last = sessions[0]   # newest first
        reply = QMessageBox.question(
            self,
            "Resume Previous Session?",
            f"A previous session was found:\n\n"
            f"  Date/Time : {last.get('timestamp', '—')}\n"
            f"  Students  : {last.get('student_count', '—')}\n"
            f"  Components: {', '.join(last.get('components', []))}\n\n"
            "Would you like to resume it?",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.Yes,
        )
        if reply == QMessageBox.Yes:
            # Index 0 in the UI list (newest first) → loads CURRENT_FILE
            self._restore_session(0)

    def _on_history(self):
        """
        Show the ResumeSessionDialog listing all saved sessions and let
        the user pick one to restore.
        """
        sessions = load_sessions_index()
        if not sessions:
            QMessageBox.information(
                self, "No History",
                "No previous sessions found.\n\n"
                f"Sessions are saved to:\n{APP_DATA_DIR}"
            )
            return

        dlg = ResumeSessionDialog(sessions, self)
        if dlg.exec_() != QDialog.Accepted:
            return
        idx = dlg.get_selected_index()
        if idx is None:
            return
        self._restore_session(idx)

    def _restore_session(self, session_index: int):
        """
        Load a persisted session and rebuild the full UI state from it.
        session_index is the 0-based position in the newest-first list.
        """
        state = load_session_by_index(session_index)
        if not state:
            QMessageBox.warning(self, "Restore Failed",
                                "Could not load the selected session.")
            return

        try:
            # ── Rebuild DataFrames ──────────────────────────────────────
            df_raw_records    = state.get("df_raw", [])
            df_graded_records = state.get("df_graded", [])
            if not df_graded_records:
                QMessageBox.warning(self, "Restore Failed",
                                    "Session data is empty or corrupt.")
                return

            df_raw    = pd.DataFrame(df_raw_records)
            df_graded = pd.DataFrame(df_graded_records)

            # Ensure Total column is float and Rank is int
            if 'Total' in df_graded.columns:
                df_graded['Total'] = pd.to_numeric(df_graded['Total'],
                                                   errors='coerce').fillna(0.0)
            if 'Rank' in df_graded.columns:
                df_graded['Rank'] = pd.to_numeric(df_graded['Rank'],
                                                  errors='coerce').fillna(0).astype(int)

            # ── Restore app state ───────────────────────────────────────
            self._clear_file_silent()   # reset without touching the file widget

            self.df_raw    = df_raw
            self.df_graded = df_graded
            self._id_col          = state.get("id_col") or (df_graded.columns[0] if len(df_graded.columns) else None)
            self._frozen_mu       = state.get("frozen_mu")
            self._frozen_sigma    = state.get("frozen_sigma")
            self._thresholds      = [tuple(t) for t in (state.get("thresholds") or [])]
            self._custom_thresholds = (
                [tuple(t) for t in state["custom_thresholds"]]
                if state.get("custom_thresholds") else None
            )
            self._use_custom_cutoffs = state.get("use_custom_cutoffs", False)
            self._components      = state.get("components", [])
            self._numeric_cols    = state.get("numeric_cols", [])
            self._edit_log        = state.get("edit_log", [])
            # NEW: restore clusters; merge with existing file-based clusters
            session_clusters = state.get("clusters", {})
            if session_clusters:
                self._clusters.update(session_clusters)
                save_clusters(self._clusters)

            # ── Restore combo / radio widgets ───────────────────────────
            base = state.get("base_grade", "B")
            idx_base = self.combo_base.findText(base)
            if idx_base >= 0:
                self.combo_base.setCurrentIndex(idx_base)

            if self._use_custom_cutoffs:
                self._radio_custom.setChecked(True)
                self._btn_set_cutoffs.setVisible(True)
            else:
                self._radio_default.setChecked(True)
                self._btn_set_cutoffs.setVisible(False)

            # ── Rebuild component rows ──────────────────────────────────
            self._clear_comp_rows()
            for comp in self._components:
                col = comp.get("column", "")
                rw = ComponentRow(col)
                rw.max_input.setText(str(comp.get("max_marks", "")))
                rw.weight_input.setText(str(comp.get("weightage", "")))
                rw.max_input.textChanged.connect(self._refresh_weight)
                rw.weight_input.textChanged.connect(self._refresh_weight)
                rw.checkbox.stateChanged.connect(self._refresh_weight)
                self.scroll_layout.insertWidget(
                    self.scroll_layout.count() - 1, rw
                )
                self.comp_rows.append(rw)
            self._refresh_weight()

            # ── Rebuild loaded-file widget (shows "Restored session") ──
            old = self.file_stack.widget(1)
            new_w = LoadedFileWidget(
                f"[Restored]  {state.get('timestamp', '')}",
                len(df_graded),
                self._clear_file,
            )
            self.file_stack.removeWidget(old)
            old.deleteLater()
            self.file_stack.insertWidget(1, new_w)
            self.file_stack.setCurrentIndex(1)

            # ── Update stats / badges / tables / graph ──────────────────
            self._cutoff_ranges = compute_cutoff_ranges(self._thresholds)

            totals = df_graded['Total'].astype(float)
            mu  = self._frozen_mu  or totals.mean()
            sig = self._frozen_sigma or totals.std()

            self.card_mean.set_value(f"{mu:.2f}")
            self.card_std.set_value(f"{sig:.2f}")
            self.card_count.set_value(str(len(df_graded)))
            self.card_min.set_value(f"{totals.min():.2f}")
            self.card_max.set_value(f"{totals.max():.2f}")
            total_w = sum(c.get('weightage', 0) for c in self._components)
            self.card_weight.set_value(f"{total_w:.1f}")

            self._populate_badges(df_graded)
            self._populate_table(df_graded)
            self._populate_ranking_table(df_graded)
            self._cutoff_display.set_cutoffs(self._cutoff_ranges)
            self._update_graph()
            self._populate_edit_log_table()
            self._refresh_cluster_tab()  # NEW

            self.btn_calc.setEnabled(True)
            self.btn_export.setEnabled(True)
            self.btn_add_comp.setEnabled(True)      # NEW
            self.btn_cluster_marks.setEnabled(True) # NEW
            self._unsaved_changes = False  # freshly restored = no unsaved changes

            self.statusBar().showMessage(
                f"Session restored  ·  {state.get('timestamp', '')}  ·  "
                f"μ={mu:.2f} (frozen)  σ={sig:.2f} (frozen)  ·  "
                f"n={len(df_graded)}"
            )

        except Exception as exc:
            QMessageBox.critical(
                self, "Restore Error",
                f"An error occurred while restoring the session:\n\n"
                f"{type(exc).__name__}: {exc}\n\n{traceback.format_exc()}"
            )

    def _clear_file_silent(self):
        """
        Reset all data state WITHOUT touching the file-stack widget.
        Used internally before restoring a session.
        """
        self.df_raw = self.df_graded = None
        self._frozen_mu = None
        self._frozen_sigma = None
        self._thresholds = []
        self._cutoff_ranges = {}
        self._custom_thresholds = None
        self._components = []
        self._numeric_cols = []
        self._edit_log = []
        self._unsaved_changes = False  # NEW

        self._clear_comp_rows()
        self._clear_table()
        self._clear_badges()
        self._graph_widget.clear_data()
        self._normal_graph_widget.clear_data()
        self._populate_edit_log_table()

        for c in (self.card_mean, self.card_std, self.card_count,
                  self.card_min, self.card_max, self.card_weight):
            c.set_value("—")
        self.btn_calc.setEnabled(False)
        self.btn_export.setEnabled(False)
        self.btn_add_comp.setEnabled(False)      # NEW
        self.btn_cluster_marks.setEnabled(False) # NEW
        self.btn_apply_team_marks.setEnabled(False)
        self._tw_lbl.setText("0.0")
        self._sel_count.setText("")

    # ── BADGES ────────────────────────────────────────────────────────
    def _populate_badges(self, df: pd.DataFrame):
        self._clear_badges()
        dist = df['Grade'].value_counts()
        for g in GRADE_ORDER:
            if g in dist:
                badge = GradeBadge(g, int(dist[g]))
                self._badge_lay.insertWidget(
                    self._badge_lay.count() - 1, badge
                )

    def _clear_badges(self):
        while self._badge_lay.count() > 1:
            item = self._badge_lay.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

    # ── COMP ROWS ─────────────────────────────────────────────────────
    def _clear_comp_rows(self):
        for w in self.comp_rows:
            self.scroll_layout.removeWidget(w)
            w.deleteLater()
        self.comp_rows.clear()

    def _load_team_component_rows(self):
        self._clear_team_comp_rows()
        for col in self._team_marks_cols:
            rw = ComponentRow(col)
            rw.max_input.textChanged.connect(self._refresh_weight)
            rw.weight_input.textChanged.connect(self._refresh_weight)
            rw.checkbox.stateChanged.connect(self._refresh_weight)
            self.team_scroll_layout.insertWidget(
                self.team_scroll_layout.count() - 1, rw
            )
            self.team_comp_rows.append(rw)

    def _clear_team_comp_rows(self):
        for w in self.team_comp_rows:
            self.team_scroll_layout.removeWidget(w)
            w.deleteLater()
        self.team_comp_rows.clear()


# ═══════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ═══════════════════════════════════════════════════════════════════════
def main():
    QApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
    QApplication.setAttribute(Qt.AA_UseHighDpiPixmaps,   True)
    app = QApplication(sys.argv)
    app.setApplicationName("Automated Grading System")
    app.setStyle("Fusion")

    pal = QPalette()
    pal.setColor(QPalette.Window,          QColor(BG_MAIN))
    pal.setColor(QPalette.WindowText,      QColor(TEXT_PRI))
    pal.setColor(QPalette.Base,            QColor(BG_PANEL))
    pal.setColor(QPalette.AlternateBase,   QColor(BG_CARD))
    pal.setColor(QPalette.ToolTipBase,     QColor(BG_CARD))
    pal.setColor(QPalette.ToolTipText,     QColor(TEXT_PRI))
    pal.setColor(QPalette.Text,            QColor(TEXT_PRI))
    pal.setColor(QPalette.Button,          QColor(BG_PANEL))
    pal.setColor(QPalette.ButtonText,      QColor(TEXT_PRI))
    pal.setColor(QPalette.BrightText,      QColor(WHITE))
    pal.setColor(QPalette.Highlight,       QColor(GOLD))
    pal.setColor(QPalette.HighlightedText, QColor("#060606"))
    pal.setColor(QPalette.Disabled, QPalette.Text,       QColor(TEXT_DIM))
    pal.setColor(QPalette.Disabled, QPalette.ButtonText, QColor(TEXT_DIM))
    app.setPalette(pal)

    w = GradingApp()
    w.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
