"""
USGS_StreamStats_FenceQuery.py
==============================
Bentley MicroStation (2024/2025+) Python Automation Script.

Features:
1. Fence Query & NHDPlus HR Stream Clipping with plan-readable text labels.
2. Native DGN pour-point (circle + crosshair) and culvert-inlet (square + X) markers,
   plus an optional longest-flowpath polyline (USGS when available, else NHD trace).
3. Culvert Analysis tab (batch mode) for multiple culvert locations (via CSV import,
   all clipped streams inside fence, or interactive CAD multi-point picking).
4. Automated EC Item Type attachment ("USGS_VDOT_Hydrology" -> "DrainageArea": IDs + drainage area only)
   including CulvertID, Drainage Area, and Q2-Q500 peak flows.
5. VDOT Rational Method (Q = Cf * C * i * A) vs. USGS Regression comparison.
6. Interactive Table Origin Picking in CAD view via DgnPrimitiveTool.
7. Multi-Sheet Excel (.xlsx) and Auto-CSV export directly into active DGN folder.
8. Zero-Dependency Google Earth KML (.kml) & ESRI Shapefile (.shp/.shx/.dbf/.prj)
   export for Watershed Polygons, Pour Points, and Clipped NHDPlus Flowlines.

Version: 2.0.0   Build date: 2026-10-30   Build By: Robert Ratliff
"""

# Bump APP_VERSION and BUILD_DATE whenever the script is changed.
APP_VERSION = "2.2.0"
BUILD_DATE = "2026-11-11"
BUILT_BY = "Robert Ratliff"

import csv
import heapq
import datetime
import gc
import json
import math
import os
import queue
import re
import struct
import sys
import threading
import time
import traceback
import urllib.parse
import urllib.request
import zipfile
from xml.sax.saxutils import escape as xml_escape
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog, colorchooser

# ---------------------------------------------------------------------------
# 1. MicroStation MSPy Imports
# ---------------------------------------------------------------------------
try:
    from MSPyBentley import *
    from MSPyBentleyGeom import *
    from MSPyECObjects import *
    from MSPyDgnPlatform import *
    from MSPyDgnView import *
    from MSPyMstnPlatform import *
    IN_MICROSTATION = True
except ImportError:
    IN_MICROSTATION = False
    print("[WARNING] MSPy modules not detected. Running in standalone UI/Test mode.")

# Optional External Libraries
try:
    import requests
    HAS_REQUESTS = True
except ImportError:
    HAS_REQUESTS = False

try:
    from pyproj import Transformer
    HAS_PYPROJ = True
except ImportError:
    HAS_PYPROJ = False

try:
    from shapely.geometry import Polygon as ShapelyPolygon, LineString as ShapelyLineString
    HAS_SHAPELY = True
except ImportError:
    HAS_SHAPELY = False

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


_LIBRARIES = (
    # (module, purpose) â€“ pip-installable packages this tool uses
    ("pyproj", "Coordinate transformations"),
    ("shapely", "Fence clipping geometry"),
    ("openpyxl", "Excel (.xlsx) export"),
)


def library_status():
    """Returns [(module, purpose, loaded, version, path)] for the libraries this tool uses."""
    rows = []
    for name, purpose in _LIBRARIES:
        mod = sys.modules.get(name)
        if mod is None:
            rows.append((name, purpose, False, "", "Not loaded \u2013 not installed in this Python"))
            continue
        path = getattr(mod, "__file__", None) or ""
        if path:
            path = os.path.dirname(path) if os.path.basename(path).startswith("__init__.") else path
        else:
            path = "(built-in)"
        ver = str(getattr(mod, "__version__", "") or getattr(mod, "VERSION", "") or getattr(mod, "TkVersion", "") or "")
        rows.append((name, purpose, True, ver, path))
    return rows


# Native-crash diagnostics: every DGN step is appended (and fsync'd) to this file so the
# last line identifies the failing call if MicroStation terminates.
TRACE_LOG_PATH = os.path.join(os.environ.get("TEMP", os.path.expanduser("~")), "USGS_StreamStats_trace.log")
try:
    import faulthandler
    _FAULT_FH = open(TRACE_LOG_PATH, "a", encoding="utf-8")
    faulthandler.enable(_FAULT_FH)
except Exception:
    _FAULT_FH = None


def _trace(step):
    try:
        with open(TRACE_LOG_PATH, "a", encoding="utf-8") as fh:
            fh.write(f"{datetime.datetime.now():%Y-%m-%d %H:%M:%S} {step}\n")
            fh.flush()
            os.fsync(fh.fileno())
    except Exception:
        pass


class _ThreadSafeStdStream:
    """MicroStation's Python IDE (Bentley.MsPythonIDE PythonEditor.OnOutput/OnError) updates a WPF TextBox
    for every write to sys.stdout/stderr. A write from a worker thread (e.g. a urllib3/requests warning,
    or a thread traceback) throws InvalidOperationException ('a different thread owns it') and terminates
    MicroStation (0xe0434352). Non-main-thread writes are diverted to the trace log instead."""

    def __init__(self, stream, main_thread, tag):
        self._stream, self._main, self._tag = stream, main_thread, tag

    def write(self, s):
        if threading.current_thread() is self._main:
            return self._stream.write(s)
        if s and s.strip():
            _trace(f"[{self._tag} from {threading.current_thread().name}] {s.rstrip()}")
        return len(s or "")

    def flush(self):
        if threading.current_thread() is self._main:
            try:
                self._stream.flush()
            except Exception:
                pass

    def __getattr__(self, name):
        return getattr(self._stream, name)


def _install_thread_safe_std_streams():
    main = threading.current_thread()
    saved = (sys.stdout, sys.stderr)
    if not isinstance(sys.stdout, _ThreadSafeStdStream):
        sys.stdout = _ThreadSafeStdStream(sys.stdout, main, "stdout")
    if not isinstance(sys.stderr, _ThreadSafeStdStream):
        sys.stderr = _ThreadSafeStdStream(sys.stderr, main, "stderr")
    try:
        import warnings
        warnings.simplefilter("ignore")
        import logging
        logging.lastResort = None
        for name in ("urllib3", "requests", "charset_normalizer", "chardet"):
            lg = logging.getLogger(name)
            lg.setLevel(logging.CRITICAL)
            lg.propagate = False
    except Exception:
        pass

    def _excepthook(args):
        _trace(f"[thread exception {getattr(args.thread, 'name', '?')}] {args.exc_type.__name__}: {args.exc_value}")

    threading.excepthook = _excepthook
    return saved


def _restore_std_streams(saved):
    if saved:
        sys.stdout, sys.stderr = saved


class _HostWindowLock:
    """
    While a background USGS request runs, Tk's message loop is pumping MicroStation's windows too.
    Clicking into MicroStation then re-enters its .NET UI from inside the Python command, which can
    terminate MicroStation (0xe0434352). Like a modal dialog, disable MicroStation's own top-level
    windows (never the Tk windows) until the request finishes.
    """
    _disabled = []
    ENABLED = True

    @classmethod
    def _host_windows(cls):
        import ctypes
        from ctypes import wintypes
        user32 = ctypes.windll.user32
        pid = os.getpid()
        found = []

        @ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
        def _enum(hwnd, _lp):
            owner = wintypes.DWORD()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(owner))
            if owner.value == pid and user32.IsWindowVisible(hwnd) and user32.IsWindowEnabled(hwnd):
                buf = ctypes.create_unicode_buffer(64)
                user32.GetClassNameW(hwnd, buf, 64)
                if not buf.value.startswith("Tk"):
                    found.append(hwnd)
            return True

        user32.EnumWindows(_enum, 0)
        return found

    @classmethod
    def lock(cls):
        if not cls.ENABLED or not IN_MICROSTATION or os.name != "nt" or cls._disabled:
            return
        try:
            import ctypes
            cls._disabled = cls._host_windows()
            for hwnd in cls._disabled:
                ctypes.windll.user32.EnableWindow(hwnd, False)
            _trace(f"host windows locked n={len(cls._disabled)}")
        except Exception as exc:
            _trace(f"host lock failed: {exc}")

    @classmethod
    def unlock(cls):
        if not cls._disabled:
            return
        try:
            import ctypes
            for hwnd in cls._disabled:
                ctypes.windll.user32.EnableWindow(hwnd, True)
        except Exception:
            pass
        cls._disabled = []
        _trace("host windows unlocked")


def _simplify_ring(pts, max_pts):
    """Douglas-Peucker with a growing tolerance until the ring fits MicroStation's vertex limit."""
    if len(pts) <= max_pts:
        return pts

    def _dp(seq, tol):
        keep = [False] * len(seq)
        keep[0] = keep[-1] = True
        stack = [(0, len(seq) - 1)]
        while stack:
            a, b = stack.pop()
            (ax, ay), (bx, by) = seq[a], seq[b]
            dx, dy = bx - ax, by - ay
            norm = math.hypot(dx, dy)
            idx, dmax = -1, tol
            for i in range(a + 1, b):
                px, py = seq[i][0] - ax, seq[i][1] - ay
                # Closed rings start and end on the same point: use the distance to that point.
                d = abs(dy * px - dx * py) / norm if norm > 1e-12 else math.hypot(px, py)
                if d > dmax:
                    idx, dmax = i, d
            if idx > 0:
                keep[idx] = True
                stack.extend(((a, idx), (idx, b)))
        return [p for p, k in zip(seq, keep) if k]

    xs = [p[0] for p in pts]
    ys = [p[1] for p in pts]
    tol = max(max(xs) - min(xs), max(ys) - min(ys)) * 1e-5
    out = pts
    while len(out) > max_pts:
        out = _dp(pts, tol)
        tol *= 2.0
    if out[0] != out[-1]:
        out.append(out[0])
    return out


def _repair_cleared_contextlib():
    """MicroStation can tear down modules from an earlier Python session, leaving contextlib's globals None
    (openpyxl/ElementTree then fail with "'NoneType' object has no attribute 'exc_info'").
    Re-runs contextlib's code inside any such cleared module so XML/XLSX writing works again."""
    mods = []
    try:
        import xml.etree.ElementTree as _et
        mods.append(_et.__dict__.get("contextlib"))
    except Exception:
        pass
    mods.append(sys.modules.get("contextlib"))
    fixed = False
    for mod in {id(m): m for m in mods if m is not None}.values():
        d = mod.__dict__
        if d.get("__name__") is not None and d.get("_sys", 0) is not None and d.get("sys", 0) is not None:
            continue
        try:
            import importlib.machinery
            spec = importlib.machinery.PathFinder.find_spec("contextlib")
            code = spec.loader.get_code("contextlib")
            import builtins as _bi
            d.update(__name__="contextlib", __spec__=spec, __loader__=spec.loader, __file__=spec.origin,
                     __builtins__=_bi)
            exec(code, d)
            fixed = True
        except Exception:
            _trace("contextlib repair failed:\n" + traceback.format_exc())
    if fixed:
        _trace("repaired cleared contextlib module")
    return fixed


UI_PREF_FILE = os.path.join(os.path.expanduser("~"), ".usgs_streamstats_ui.json")


def _load_ui_prefs():
    try:
        with open(UI_PREF_FILE, "r", encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def _save_ui_prefs(**updates):
    """Merges `updates` into the UI preferences file (theme, KML style location, ...)."""
    prefs = _load_ui_prefs()
    prefs.update(updates)
    try:
        with open(UI_PREF_FILE, "w", encoding="utf-8") as f:
            json.dump(prefs, f, indent=2)
    except Exception:
        pass


def _status_ok(status):
    """MSPy returns several different status enums (BentleyStatus, ReprojectStatus,
    TextBlockToElementResult...). All use 0 for success, but pybind enums of different
    types never compare equal, so normalise to int before checking."""
    try:
        return int(status) == 0
    except (TypeError, ValueError):
        return bool(status)


# ---------------------------------------------------------------------------
# 2. Interactive DGN Primitive Tools
# ---------------------------------------------------------------------------
if IN_MICROSTATION:
    class DgnPointPickerTool(DgnPrimitiveTool):
        """Interactive tool that captures a data click in the CAD view."""

        def __init__(self, on_picked_callback, prompt_msg="Click a point in the DGN:", multi=False):
            super().__init__(0, 0)
            self.on_picked_callback = on_picked_callback
            self.prompt_msg = prompt_msg
            self.multi = multi

        def _OnPostInstall(self):
            AccuSnap.GetInstance().EnableSnap(True)
            NotificationManager.OutputPrompt(self.prompt_msg)
            super()._OnPostInstall()

        def _OnDataButton(self, ev):
            pt = ev.GetPoint()
            cb = self.on_picked_callback
            if not self.multi:
                self.on_picked_callback = None
                self._ExitTool()
            if cb:
                cb(float(pt.x), float(pt.y), float(pt.z))
            if self.multi:
                NotificationManager.OutputPrompt(self.prompt_msg)
            return True

        def _OnResetButton(self, ev):
            NotificationManager.OutputPrompt("Marker placement finished." if self.multi else "Pick canceled.")
            cb, self.on_picked_callback = self.on_picked_callback, None
            self._ExitTool()
            if cb:
                cb(None, None, None)
            return True

        def _OnRestartTool(self):
            self._ExitTool()

    class DgnFencePlaceTool(DgnPrimitiveTool):
        """Places a block or shape fence, drawing the rubber-band outline in red while it is being placed
        (MicroStation's own Place Fence tool cannot change its dynamics color).
        on_done(points or None, error text or None) runs once, after the tool exits."""
        COLOR, WEIGHT = 3, 2

        def __init__(self, mode, on_done):
            super().__init__(0, 0)
            self.mode = "SHAPE" if str(mode).upper() == "SHAPE" else "BLOCK"
            self.on_done = on_done
            self.pts = []
            self.vp = None
            self._finished = False

        def _prompt(self):
            if not self.pts:
                msg = "Place Fence > Enter first point"
            elif self.mode == "BLOCK":
                msg = "Place Fence > Enter opposite corner"
            elif len(self.pts) < 3:
                msg = "Place Fence > Enter next vertex"
            else:
                msg = "Place Fence > Enter next vertex, or Reset (right-click) to close the fence"
            NotificationManager.OutputPrompt(msg)

        def _OnPostInstall(self):
            AccuSnap.GetInstance().EnableSnap(True)
            self._prompt()
            super()._OnPostInstall()

        def _outline(self, cursor):
            """Closed outline through the placed points and the cursor position."""
            if self.mode == "BLOCK":
                (x1, y1, z), (x2, y2) = self.pts[0], cursor[:2]
                return [(x1, y1, z), (x2, y1, z), (x2, y2, z), (x1, y2, z), (x1, y1, z)]
            out = list(self.pts) + [cursor]
            return out + [out[0]] if len(out) >= 3 else out

        def _OnDataButton(self, ev):
            pt = ev.GetPoint()
            p = (float(pt.x), float(pt.y), float(pt.z))
            self.vp = ev.GetViewport()
            if not self.pts:
                self.pts.append(p)
                self._BeginDynamics()
                self._prompt()
                return False
            if self.mode == "BLOCK":
                if p[0] == self.pts[0][0] or p[1] == self.pts[0][1]:
                    return False
                self._finish(self._outline(p))
                return True
            if p != self.pts[-1]:
                self.pts.append(p)
            self._prompt()
            return False

        def _OnResetButton(self, ev):
            if self.mode == "SHAPE" and len(self.pts) >= 3:
                self._finish(list(self.pts) + [self.pts[0]])
            else:
                self._finish(None)
            return True

        def _OnDynamicFrame(self, ev):
            try:
                if not self.pts:
                    return
                pt = ev.GetPoint()
                dpts = DPoint3dArray()
                for x, y, z in self._outline((float(pt.x), float(pt.y), float(pt.z))):
                    dpts.append(DPoint3d(x, y, z))
                model_ref = ISessionMgr.ActiveDgnModelRef
                eeh = EditElementHandle()
                if LineStringHandler.CreateLineStringElement(eeh, None, dpts, model_ref.Is3d(),
                                                             model_ref) != BentleyStatus.eSUCCESS:
                    return
                setter = ElementPropertiesSetter()
                setter.SetColor(self.COLOR)
                setter.SetWeight(self.WEIGHT)
                setter.SetLinestyle(0, None)
                setter.Apply(eeh)
                redraw = RedrawElems()
                redraw.SetDynamicsViews(IViewManager.GetActiveViewSet(), ev.GetViewport())
                redraw.SetDrawMode(eDRAW_MODE_TempDraw)
                redraw.SetDrawPurpose(DrawPurpose.eDynamics)
                redraw.DoRedraw(eeh)
            except Exception as exc:
                _trace(f"fence dynamics error: {exc}")

        def _finish(self, pts):
            if self._finished:
                return
            self._finished = True
            cb, self.on_done = self.on_done, None
            err = None
            if pts:
                try:
                    dpts = DPoint3dArray()
                    for x, y, z in pts:
                        dpts.append(DPoint3d(x, y, z))
                    if int(FenceManager.GetManager().DefineByPoints(dpts, self.vp)) != 0:
                        pts, err = None, "MicroStation rejected the fence points"
                except Exception as exc:
                    pts, err = None, str(exc)
            NotificationManager.OutputPrompt("Fence placed." if pts else "Fence placement cancelled.")
            self._ExitTool()
            if cb:
                cb(pts, err)

        def _OnRestartTool(self):
            self._finish(None)

        def _OnCleanup(self):
            # Another command replaced this tool before the fence was finished: report it as cancelled.
            if not self._finished:
                self._finished = True
                cb, self.on_done = self.on_done, None
                if cb:
                    cb(None, None)


# ---------------------------------------------------------------------------
# 3. VDOT Rational Method & Hydrology Comparison Calculator
# ---------------------------------------------------------------------------
class VDOTHydrologyCalculator:
    """Computes VDOT Rational discharges (Q = Cf * C * i * A) & regression comparisons."""

    CF_FACTORS = {2: 1.00, 5: 1.00, 10: 1.00, 25: 1.10, 50: 1.20, 100: 1.25, 200: 1.25, 500: 1.25}
    VA_IDF_COEFFS = {
        2: (48.5, 8.5, 0.78), 5: (63.0, 9.0, 0.78), 10: (75.2, 9.5, 0.78),
        25: (92.4, 10.0, 0.78), 50: (106.8, 10.5, 0.78), 100: (122.0, 11.0, 0.78),
        200: (138.0, 11.5, 0.78), 500: (160.5, 12.0, 0.78),
    }

    @classmethod
    def extract_basin_metrics(cls, parameters_list):
        metrics = {"drnarea_sqmi": 0.0, "drnarea_acres": 0.0, "slope_ft_mi": 50.0, "imperv_pct": 2.0, "forest_pct": 50.0}
        for p in parameters_list or []:
            code = str(p.get("code", "")).upper()
            try:
                val = float(p.get("value", 0.0))
            except (ValueError, TypeError):
                continue
            if code == "DRNAREA":
                metrics["drnarea_sqmi"] = val
                metrics["drnarea_acres"] = val * 640.0
            elif code in ("CSL10_85", "CSL1085LFP", "SLOP30_10M"):
                metrics["slope_ft_mi"] = max(val, 1.0)
            elif "IMP" in code:
                metrics["imperv_pct"] = max(0.0, min(val, 100.0))
            elif "FOR" in code or "WOOD" in code:
                metrics["forest_pct"] = max(0.0, min(val, 100.0))
        return metrics

    @classmethod
    def estimate_vdot_c_and_tc(cls, metrics, user_c_override=None, flow_length_ft=None):
        """Rational C and Kirpich Tc. Uses the measured longest-flowpath length when given,
        otherwise a length estimated from the drainage area."""
        if user_c_override is not None and user_c_override > 0.0:
            c_val = min(max(user_c_override, 0.05), 0.98)
        else:
            imp_frac = metrics["imperv_pct"] / 100.0
            for_frac = metrics["forest_pct"] / 100.0
            open_frac = max(0.0, 1.0 - imp_frac - for_frac)
            c_val = round(0.90 * imp_frac + 0.20 * for_frac + 0.35 * open_frac, 2)
            c_val = min(max(c_val, 0.15), 0.95)

        if flow_length_ft and flow_length_ft > 0:
            length_ft = float(flow_length_ft)
        else:
            da_sqmi = max(metrics["drnarea_sqmi"], 0.005)
            length_ft = 1.4 * (da_sqmi ** 0.6) * 5280.0
        slope_ft_ft = max(metrics["slope_ft_mi"] / 5280.0, 0.0005)
        tc_min = 0.0078 * (length_ft ** 0.77) * (slope_ft_ft ** -0.385)
        tc_min = round(min(max(tc_min, 5.0), 360.0), 1)
        return c_val, tc_min

    @staticmethod
    def atlas14_intensity(atlas14, yr, tc_min):
        """Rainfall intensity (in/hr) for a return period at duration Tc, log-log interpolated from a
        NOAA Atlas 14 table. None when the table is missing or does not cover the return period."""
        try:
            col = atlas14["ari"].index(yr)
            durs = atlas14["durations_min"]
            vals = [float(row[col]) for row in atlas14["intensity"]]
            if len(vals) != len(durs) or min(vals) <= 0:
                return None
            t = min(max(float(tc_min), durs[0]), durs[-1])
            for i in range(len(durs) - 1):
                if durs[i] <= t <= durs[i + 1]:
                    f = (math.log(t) - math.log(durs[i])) / (math.log(durs[i + 1]) - math.log(durs[i]))
                    return math.exp(math.log(vals[i]) + f * (math.log(vals[i + 1]) - math.log(vals[i])))
        except Exception:
            pass
        return None

    RAIN_SRC_ATLAS14 = "NOAA Atlas 14 (at the pour point)"
    RAIN_SRC_BUILTIN = "Built-in statewide IDF coefficients (placeholder)"
    TC_SRC_MEASURED = "Measured longest flowpath"
    TC_SRC_ESTIMATED = "Estimated from drainage area"
    REG_SRC_NSS = "USGS NSS regression equations"
    REG_SRC_BUILTIN = "Built-in check equation (placeholder)"

    # ---- Peak-discharge methods: VDOT Drainage Manual, Chapter 6 (Rev. 7/26), Table 6-2 -------------------
    # (key, label, can be calculated here, on by default, smallest area ac, largest area ac, note)
    METHODS = (
        ("rational", "Rational Method", True, True, 0.0, 200.0,
         "Q = Cf\u00b7C\u00b7i\u00b7A with NOAA Atlas 14 intensity at Tc"),
        ("nrcs", "NRCS Graphical Peak Discharge (TR-55)", True, True, 0.0, 2000.0,
         "Type II storm, NOAA Atlas 14 24-hr rainfall, CN from land cover and soils (or entered below)"),
        ("efh2", "NRCS EFH-2", False, False, 0.0, 2000.0, "Not calculated here \u2013 run the NRCS EFH-2 program"),
        ("anderson", "Anderson Method (USGS)", True, False, 200.0, None,
         "VDOT: only for urbanized areas in the Northern Virginia District"),
        ("rural", "USGS Regional Regression \u2013 Rural", True, True, 200.0, None,
         "USGS StreamStats / NSS equations (SIR 2025-5110)"),
        ("urban", "USGS Regional Regression \u2013 Urban", True, True, 200.0, None,
         "SIR 2014-5090; basin must be at least 10% urban and 1.2 to 2,400 sq mi"),
        ("gage", "Stream Gage Data (Log-Pearson III)", False, False, 200.0, None,
         "Not calculated here \u2013 needs 10+ years of gage records"),
        ("snyder", "Snyder Method", False, False, None, None,
         "Not in the current VDOT Drainage Manual (Rev. 7/26)"),
    )
    METHOD_LABELS = {
        "rational": "VDOT Rational (<=200 ac)", "nrcs": "NRCS Graphical Peak (TR-55)",
        "anderson": "Anderson Method (USGS)", "rural": "USGS Rural Regression",
        "urban": "USGS Urban Regression", "builtin": "Built-in regression check (placeholder)",
    }
    ANDERSON_TYPES = (("natural", "Natural rural watershed"), ("developed", "Developed, partially channelized"),
                      ("sewered", "Completely developed and sewered"))
    HSG_CHOICES = ("auto", "A", "B", "C", "D")

    @classmethod
    def default_options(cls):
        """Method settings as the VDOT Drainage Manual recommends them."""
        return {"enabled": {m[0]: m[3] for m in cls.METHODS if m[2]}, "use_ranges": True, "design": "auto",
                "cn": None, "hsg": "auto", "anderson_type": "natural"}

    options = None   # set by the Settings tab; None -> default_options()

    @classmethod
    def range_text(cls, lo, hi):
        if lo is None and hi is None:
            return "\u2014"
        if hi is None:
            return f"over {lo:,.0f} ac"
        return f"{lo:,.0f} to {hi:,.0f} ac"

    # Anderson (USGS WSP 2001-C; VDOT eq. 6.6-6.8, Tables 6-5 and 6-6). The 2.33-yr ratio is the mean annual
    # flood, so there is no 2-yr value. The manual prints K = 1 + 0.15I; Anderson's report and the 2.5*R100
    # term (Rf must equal R100 at 100% impervious) both give 0.015, which is used here.
    ANDERSON_K_PER_PCT = 0.015
    ANDERSON_LAG = {"natural": (4.64, 0.42), "developed": (0.90, 0.50), "sewered": (0.56, 0.52)}
    ANDERSON_RATIOS = {5: (1.65, 1.24), 10: (2.20, 1.45), 25: (3.30, 1.80), 50: (4.40, 2.00), 100: (5.50, 2.20)}

    @classmethod
    def anderson_q(cls, yr, area_sqmi, length_mi, slope_ft_mi, imperv_pct, basin_type="natural"):
        if yr == 500:
            q100 = cls.anderson_q(100, area_sqmi, length_mi, slope_ft_mi, imperv_pct, basin_type)
            return None if q100 is None else 1.7 * q100
        if yr not in cls.ANDERSON_RATIOS or min(area_sqmi, length_mi, slope_ft_mi) <= 0:
            return None
        coef, exp = cls.ANDERSON_LAG.get(basin_type, cls.ANDERSON_LAG["natural"])
        lag_hr = coef * (length_mi / math.sqrt(slope_ft_mi)) ** exp
        imp = min(max(imperv_pct, 0.0), 100.0)
        k = 1.0 + cls.ANDERSON_K_PER_PCT * imp
        rn, r100 = cls.ANDERSON_RATIOS[yr]
        rf = (rn + 0.01 * imp * (2.5 * r100 - rn)) / k
        return rf * 230.0 * k * area_sqmi ** 0.82 * lag_hr ** -0.48

    # USGS urban regression for Virginia (SIR 2014-5090; VDOT section 6.4.3.4.4): b0..b5 by return period.
    URBAN_BETA = {
        2: (2.027, 40.290, 1.216, -0.00414, 0.00468, -0.366), 5: (2.229, 39.370, 1.139, -0.00346, 0.00487, -0.338),
        10: (2.373, 38.706, 1.103, -0.00313, 0.00470, -0.334), 25: (2.557, 39.168, 1.083, -0.00224, 0.00434, -0.332),
        50: (2.697, 39.168, 1.083, -0.00219, 0.00390, -0.343), 100: (2.776, 38.765, 1.070, -0.00242, 0.00434, -0.342),
        200: (2.863, 39.063, 1.057, -0.00223, 0.00465, -0.329), 500: (2.961, 39.287, 0.904, -0.00049, 0.00636, -0.317),
    }
    URBAN_MIN_PCT, URBAN_MIN_SQMI, URBAN_MAX_SQMI = 10.0, 1.2, 2400.0

    @classmethod
    def urban_q(cls, yr, area_sqmi, urban_pct):
        b = cls.URBAN_BETA.get(yr)
        if not b or area_sqmi <= 0:
            return None
        la = math.log10(area_sqmi)
        log_q = b[0] + (urban_pct - b[1]) * ((la - b[2]) * b[3]) + urban_pct * b[4] + la * b[5]
        return (10.0 ** log_q) * area_sqmi

    # NRCS graphical peak discharge (TR-55 Appendix F, Type II storm): Ia/P, C0, C1, C2.
    TR55_TYPE2 = ((0.10, 2.55323, -0.61512, -0.16403), (0.30, 2.46532, -0.62257, -0.11657),
                  (0.35, 2.41896, -0.61594, -0.08820), (0.40, 2.36409, -0.59857, -0.05621),
                  (0.45, 2.29238, -0.57005, -0.02281), (0.50, 2.20282, -0.51599, -0.01259))

    @classmethod
    def nrcs_q(cls, rain_24h_in, cn, area_sqmi, tc_min):
        """qp = qu * A * Q (no pond/swamp adjustment). Tc is limited to TR-55's 0.1-10 hr range."""
        if not rain_24h_in or not cn or area_sqmi <= 0:
            return None
        cn = min(max(float(cn), 30.0), 98.0)
        s = 1000.0 / cn - 10.0
        ia = 0.2 * s
        if rain_24h_in <= ia:
            return 0.0
        runoff_in = (rain_24h_in - ia) ** 2 / (rain_24h_in - ia + s)
        ratio = min(max(ia / rain_24h_in, 0.10), 0.50)
        lt = math.log10(min(max(tc_min / 60.0, 0.1), 10.0))
        tab = cls.TR55_TYPE2
        for lo, hi in zip(tab, tab[1:]):
            if lo[0] <= ratio <= hi[0]:
                f = (ratio - lo[0]) / (hi[0] - lo[0])
                log_lo = lo[1] + lo[2] * lt + lo[3] * lt * lt
                log_hi = hi[1] + hi[2] * lt + hi[3] * lt * lt
                return (10.0 ** (log_lo + f * (log_hi - log_lo))) * area_sqmi * runoff_in
        return None

    # Runoff curve numbers by NLCD class for hydrologic soil groups A, B, C, D (TR-55 Table 2-2 cover types).
    NLCD_CN = {
        11: (98, 98, 98, 98), 21: (49, 69, 79, 84), 22: (57, 72, 81, 86), 23: (77, 85, 90, 92),
        24: (89, 92, 94, 95), 31: (77, 86, 91, 94), 41: (30, 55, 70, 77), 42: (30, 55, 70, 77),
        43: (30, 55, 70, 77), 52: (35, 56, 70, 77), 71: (30, 58, 71, 78), 81: (39, 61, 74, 80),
        82: (67, 78, 85, 89), 90: (77, 77, 77, 77), 95: (78, 78, 78, 78),
    }

    @classmethod
    def estimate_cn(cls, nlcd_classes, hsg_fractions):
        """Area-weighted CN from NLCD class fractions and soil-group fractions (treated as independent)."""
        tot_c = sum(v for k, v in (nlcd_classes or {}).items() if int(k) in cls.NLCD_CN)
        tot_h = sum(v for k, v in (hsg_fractions or {}).items() if k in "ABCD" and len(k) == 1)
        if tot_c <= 0 or tot_h <= 0:
            return None
        cn = 0.0
        for c, cf in nlcd_classes.items():
            row = cls.NLCD_CN.get(int(c))
            if row:
                cn += (cf / tot_c) * sum((hf / tot_h) * row["ABCD".index(h)]
                                         for h, hf in hsg_fractions.items() if h in "ABCD" and len(h) == 1)
        return round(cn, 1)

    @classmethod
    def compute_comparison_table(cls, ss_result, user_c_override=None):
        if not ss_result:
            return {}
        opts = dict(cls.default_options(), **(cls.options or {}))
        enabled = dict(cls.default_options()["enabled"], **(opts.get("enabled") or {}))
        meta = {m[0]: m for m in cls.METHODS}

        metrics = cls.extract_basin_metrics(ss_result.get("parameters", []))
        tc_length = ss_result.get("lfp_length_ft") or ss_result.get("tc_flow_length_ft") or None
        c_val, tc_min = cls.estimate_vdot_c_and_tc(metrics, user_c_override, tc_length)
        acres = metrics["drnarea_acres"]
        sqmi = max(metrics["drnarea_sqmi"], 0.0001)
        flow_stats = ss_result.get("flow_stats", {})
        atlas14 = ss_result.get("atlas14")
        nlcd = ss_result.get("nlcd") or {}

        urban_pct = None
        for p in ss_result.get("parameters", []) or []:
            if re.match(r"^LC\d\dDEV$", str(p.get("code", "")).upper()):
                try:
                    urban_pct = float(p.get("value"))
                except (TypeError, ValueError):
                    pass
        urban_src = "USGS StreamStats"
        if urban_pct is None and nlcd.get("dev_pct") is not None:
            urban_pct, urban_src = float(nlcd["dev_pct"]), "NLCD 2021 developed classes 21-24"
        metrics["urban_pct"] = urban_pct

        # Curve number for the NRCS method: Settings override, else NLCD land cover x soil group.
        cn, cn_src = None, ""
        if opts.get("cn"):
            cn, cn_src = float(opts["cn"]), "Entered in Settings"
        elif nlcd.get("classes"):
            hsg_opt = str(opts.get("hsg") or "auto").upper()
            if hsg_opt in ("A", "B", "C", "D"):
                hsg, hsg_txt = {hsg_opt: 1.0}, f"soil group {hsg_opt} (Settings)"
            elif ss_result.get("hsg"):
                hsg = ss_result["hsg"]
                tot = sum(hsg.values()) or 1.0
                src_name = str(ss_result.get("hsg_source") or "USDA SSURGO").split(";")[0].split(" (")[0]
                hsg_txt = f"{src_name} soils " + ", ".join(
                    f"{k} {100.0 * v / tot:.0f}%" for k, v in sorted(hsg.items()))
            else:
                hsg, hsg_txt = {"B": 1.0}, "soil group B assumed (SSURGO unavailable)"
            cn = cls.estimate_cn(nlcd["classes"], hsg)
            cn_src = f"Estimated: NLCD 2021 land cover, {hsg_txt}"

        # Which methods run: (active, why not).
        status = {}
        for key, label, can_calc, _d, lo, hi, note in cls.METHODS:
            if not can_calc:
                status[key] = (False, note)
            elif not enabled.get(key, False):
                status[key] = (False, "Turned off in Settings")
            elif opts.get("use_ranges", True) and acres > 0 and (
                    (lo and acres <= lo) or (hi is not None and acres > hi)):
                status[key] = (False, f"Outside the VDOT range ({cls.range_text(lo, hi)})")
            else:
                status[key] = (True, "")
        if status["nrcs"][0] and (cn is None or not atlas14):
            status["nrcs"] = (False, "Needs a curve number" if cn is None else "Needs NOAA Atlas 14 rainfall")
        if status["anderson"][0] and not tc_length:
            status["anderson"] = (False, "Needs the longest flowpath (turn it on in Settings)")
        if status["urban"][0]:
            if urban_pct is None:
                status["urban"] = (False, "Percent urban area is not available")
            elif urban_pct < cls.URBAN_MIN_PCT:
                status["urban"] = (False, f"Basin is {urban_pct:.0f}% urban (equations need at least 10%)")
            elif not cls.URBAN_MIN_SQMI <= sqmi <= cls.URBAN_MAX_SQMI:
                status["urban"] = (False, f"Drainage area {sqmi:.2f} sq mi is outside 1.2 to 2,400 sq mi")

        # Design-flow order (VDOT section 6.3.2): Rational up to 200 ac; above that the USGS regression,
        # urban when the basin is 30% or more developed (the rural equations are not valid there).
        if 0.0 < acres <= 200.0:
            order = ["rational", "nrcs", "rural", "urban", "anderson"]
        elif urban_pct is not None and urban_pct >= 30.0:
            order = ["urban", "rural", "nrcs", "anderson"]
        else:
            order = ["rural", "urban", "nrcs", "anderson"]
        design = str(opts.get("design") or "auto")
        if design in order:
            order.remove(design)
            order.insert(0, design)

        va_reg_check = {
            2: (108.0, 0.68), 5: (185.0, 0.67), 10: (248.0, 0.66), 25: (342.0, 0.65),
            50: (425.0, 0.64), 100: (518.0, 0.63), 200: (620.0, 0.62), 500: (775.0, 0.61),
        }

        comparison_rows = []
        rain_atlas = rain_builtin = 0
        used = {}
        for yr in (2, 5, 10, 25, 50, 100, 200, 500):
            pk_tag = f"PK{yr}"
            usgs_q = None
            for k, v in flow_stats.items():
                if k == pk_tag or k.startswith(f"{pk_tag}_") or k.endswith(f"_{pk_tag}"):
                    try:
                        usgs_q = float(v["value"])
                    except (ValueError, TypeError):
                        pass
                    break

            cf = cls.CF_FACTORS.get(yr, 1.0)
            intensity_in_hr = cls.atlas14_intensity(atlas14, yr, tc_min) if atlas14 else None
            if intensity_in_hr is None:
                a_idf, b_idf, c_idf = cls.VA_IDF_COEFFS.get(yr, (75.2, 9.5, 0.78))
                intensity_in_hr = a_idf / ((tc_min + b_idf) ** c_idf)
                rain_builtin += 1
            else:
                rain_atlas += 1
            # VDOT Table 6-4: Cf x C shall not exceed 1.0.
            q_rational_any = round(min(cf * c_val, 1.0) * intensity_in_hr * acres, 1) if acres > 0 else 0.0
            k_reg, exp_reg = va_reg_check.get(yr, (248.0, 0.66))
            q_va_check = round(k_reg * (sqmi ** exp_reg), 1) if sqmi > 0 else 0.0

            q = {"rational": q_rational_any if status["rational"][0] else None, "nrcs": None, "anderson": None,
                 "rural": usgs_q if (status["rural"][0] and usgs_q is not None and usgs_q > 0) else None,
                 "urban": None}
            rain_24h = cls.atlas14_intensity(atlas14, yr, 1440.0) if atlas14 else None
            rain_24h = rain_24h * 24.0 if rain_24h else None
            if status["nrcs"][0]:
                q["nrcs"] = cls.nrcs_q(rain_24h, cn, sqmi, tc_min)
            if status["anderson"][0]:
                q["anderson"] = cls.anderson_q(yr, sqmi, float(tc_length) / 5280.0, metrics["slope_ft_mi"],
                                               metrics["imperv_pct"], opts.get("anderson_type", "natural"))
            if status["urban"][0]:
                q["urban"] = cls.urban_q(yr, sqmi, urban_pct)
            q = {k: (round(v, 1) if v is not None else None) for k, v in q.items()}

            pick = next((k for k in order if q.get(k) is not None and q[k] > 0), None)
            if pick is None:   # nothing enabled produced a value: Rational up to 200 ac, else the check equation
                pick = "rational" if 0.0 < acres <= 200.0 else "builtin"
                vdot_q = q_rational_any if pick == "rational" else q_va_check
            else:
                vdot_q = q[pick]
            used[pick] = used.get(pick, 0) + 1

            pct_diff_str = "N/A"
            if usgs_q is not None and usgs_q > 0 and vdot_q > 0:
                pct_diff = ((vdot_q - usgs_q) / usgs_q) * 100.0
                pct_diff_str = f"{pct_diff:+.1f}%"

            comparison_rows.append({
                "return_yr": yr, "code": pk_tag,
                "usgs_q_cfs": round(usgs_q, 1) if usgs_q is not None else "N/A",
                "vdot_q_cfs": vdot_q, "q_rational_cfs": q["rational"], "q_va_reg_check_cfs": q_va_check,
                "q_nrcs_cfs": q["nrcs"], "q_anderson_cfs": q["anderson"], "q_urban_cfs": q["urban"],
                "design_method": pick, "rain_24h_in": round(rain_24h, 2) if rain_24h else None,
                "intensity_in_hr": round(intensity_in_hr, 2),
                "pct_diff": pct_diff_str,
            })

        # Governing method = the one behind the 10-yr design flow (the most common design event).
        gov_key = next((r["design_method"] for r in comparison_rows if r["return_yr"] == 10), "rational")
        governing_label = cls.METHOD_LABELS.get(gov_key, gov_key)
        if gov_key == "rural":
            governing_label = "USGS/VDOT Regression (>200 ac)" if acres > 200.0 else governing_label
        for r in comparison_rows:
            r["governing"] = governing_label

        if rain_atlas and rain_builtin:
            rain_src = f"{cls.RAIN_SRC_ATLAS14}; {cls.RAIN_SRC_BUILTIN} for {rain_builtin} event(s)"
        else:
            rain_src = cls.RAIN_SRC_ATLAS14 if rain_atlas else cls.RAIN_SRC_BUILTIN
        if not (used.get("rural") or used.get("urban") or used.get("builtin")):
            reg_src = f"Not used ({cls.METHOD_LABELS.get(gov_key, gov_key).split(' (')[0]} governs)"
        elif used.get("builtin") and (used.get("rural") or used.get("urban")):
            reg_src = f"{cls.REG_SRC_NSS}; {cls.REG_SRC_BUILTIN} for {used['builtin']} event(s)"
        elif used.get("builtin"):
            reg_src = cls.REG_SRC_BUILTIN
        else:
            reg_src = cls.REG_SRC_NSS if used.get("rural") else "USGS urban equations (SIR 2014-5090)"

        methods = []
        for key, label, can_calc, _d, lo, hi, note in cls.METHODS:
            ok, why = status[key]
            methods.append({"key": key, "label": label, "calculated": bool(ok), "range": cls.range_text(lo, hi),
                            "status": "Calculated" if ok else why})
        return {
            "metrics": metrics, "c_coeff": c_val, "tc_min": tc_min,
            "governing_method": governing_label, "rows": comparison_rows,
            "tc_length_ft": round(float(tc_length), 1) if tc_length else None,
            "tc_length_source": cls.TC_SRC_MEASURED if tc_length else cls.TC_SRC_ESTIMATED,
            "rainfall_source": rain_src, "regression_source": reg_src,
            "basin_source": ss_result.get("bc_source") or "USGS StreamStats",
            "methods": methods, "design_setting": design,
            "cn": cn, "cn_source": cn_src, "urban_pct": urban_pct, "urban_source": urban_src if urban_pct is not None else "",
            "anderson_type": dict(cls.ANDERSON_TYPES).get(opts.get("anderson_type", "natural"), ""),
        }


# ---------------------------------------------------------------------------
# 4. Pure-Python Polyline-to-Polygon Clipper
# ---------------------------------------------------------------------------
class PurePythonClipper:
    @staticmethod
    def point_in_polygon(pt, poly):
        x, y = pt
        inside = False
        n = len(poly)
        p1x, p1y = poly[0]
        for i in range(1, n + 1):
            p2x, p2y = poly[i % n]
            if y > min(p1y, p2y) and y <= max(p1y, p2y) and x <= max(p1x, p2x):
                if p1y != p2y:
                    xinters = (y - p1y) * (p2x - p1x) / (p2y - p1y) + p1x
                if p1x == p2x or x <= xinters:
                    inside = not inside
            p1x, p1y = p2x, p2y
        return inside

    @staticmethod
    def segment_intersection(p1, p2, q1, q2):
        r = (p2[0] - p1[0], p2[1] - p1[1])
        s = (q2[0] - q1[0], q2[1] - q1[1])
        rxs = r[0] * s[1] - r[1] * s[0]
        if abs(rxs) < 1e-14:
            return None
        qp = (q1[0] - p1[0], q1[1] - p1[1])
        t = (qp[0] * s[1] - qp[1] * s[0]) / rxs
        u = (qp[0] * r[1] - qp[1] * r[0]) / rxs
        if 0.0 <= t <= 1.0 and 0.0 <= u <= 1.0:
            return t, (p1[0] + t * r[0], p1[1] + t * r[1])
        return None

    @classmethod
    def clip_polyline_to_polygon(cls, polyline, polygon):
        if len(polyline) < 2 or len(polygon) < 3:
            return []
        if HAS_SHAPELY:
            poly_geom = ShapelyPolygon(polygon)
            line_geom = ShapelyLineString(polyline)
            clipped = line_geom.intersection(poly_geom)
            if clipped.is_empty:
                return []
            if clipped.geom_type == "LineString":
                return [list(clipped.coords)]
            elif clipped.geom_type in ("MultiLineString", "GeometryCollection"):
                return [list(g.coords) for g in clipped.geoms if g.geom_type == "LineString" and len(g.coords) >= 2]
            return []

        closed_poly = list(polygon)
        if closed_poly[0] != closed_poly[-1]:
            closed_poly.append(closed_poly[0])

        inside_subsegments = []
        for i in range(len(polyline) - 1):
            p1, p2 = polyline[i], polyline[i + 1]
            t_values = [0.0, 1.0]
            for j in range(len(closed_poly) - 1):
                hit = cls.segment_intersection(p1, p2, closed_poly[j], closed_poly[j + 1])
                if hit is not None:
                    t_values.append(hit[0])

            t_values = sorted(set(round(t, 9) for t in t_values))
            for k in range(len(t_values) - 1):
                t_a, t_b = t_values[k], t_values[k + 1]
                if t_b - t_a < 1e-8:
                    continue
                t_mid = 0.5 * (t_a + t_b)
                mid_pt = (p1[0] + t_mid * (p2[0] - p1[0]), p1[1] + t_mid * (p2[1] - p1[1]))
                if cls.point_in_polygon(mid_pt, closed_poly):
                    pt_a = (p1[0] + t_a * (p2[0] - p1[0]), p1[1] + t_a * (p2[1] - p1[1]))
                    pt_b = (p1[0] + t_b * (p2[0] - p1[0]), p1[1] + t_b * (p2[1] - p1[1]))
                    inside_subsegments.append((pt_a, pt_b))

        if not inside_subsegments:
            return []

        polylines = []
        chain = [inside_subsegments[0][0], inside_subsegments[0][1]]
        for seg_start, seg_end in inside_subsegments[1:]:
            last_pt = chain[-1]
            if math.hypot(seg_start[0] - last_pt[0], seg_start[1] - last_pt[1]) < 1e-8:
                chain.append(seg_end)
            else:
                polylines.append(chain)
                chain = [seg_start, seg_end]
        polylines.append(chain)
        return polylines


# ---------------------------------------------------------------------------
# 5. Coordinate Transformation Manager
# ---------------------------------------------------------------------------
class DgnCoordinateProjector:
    def __init__(self, fallback_epsg=2284):
        self.fallback_epsg = fallback_epsg
        self.uor_per_master = 10000.0
        self.global_origin = (0.0, 0.0, 0.0)
        self.dgn_gcs = None
        self.model_ref = None
        self.to_wgs84 = None
        self.from_wgs84 = None

        if IN_MICROSTATION:
            self.model_ref = ISessionMgr.ActiveDgnModelRef
            if self.model_ref is not None:
                dgn_model = self.model_ref.GetDgnModel()
                model_info = dgn_model.GetModelInfo()
                self.uor_per_master = self._read_uor_per_master(model_info)
                try:
                    go = model_info.GetGlobalOrigin()
                    self.global_origin = (float(go.x), float(go.y), float(go.z))
                except Exception:
                    pass
                try:
                    self.dgn_gcs = DgnGCS.FromModel(self.model_ref, True)
                except Exception:
                    self.dgn_gcs = None

        self._init_pyproj(fallback_epsg)

    @staticmethod
    def _read_uor_per_master(model_info):
        """ModelInfo has no GetUorPerMaster(); derive it from the storage unit (Bentley sample recipe)."""
        try:
            ok, factor = model_info.StorageUnit.ConvertDistanceFrom(model_info.UorPerStorage, model_info.MasterUnit)
            if factor and float(factor) > 0:
                return float(factor)
        except Exception:
            pass
        try:
            ok, factor = model_info.GetStorageUnit().ConvertDistanceFrom(
                model_info.GetUorPerStorage(), model_info.GetMasterUnit())
            if factor and float(factor) > 0:
                return float(factor)
        except Exception:
            pass
        return 10000.0

    @property
    def has_gcs(self):
        return self.dgn_gcs is not None

    def master_to_uor(self, x_m, y_m, z_m=0.0):
        gx, gy, gz = self.global_origin
        return x_m * self.uor_per_master + gx, y_m * self.uor_per_master + gy, z_m * self.uor_per_master + gz

    def uor_to_master(self, x_uor, y_uor, z_uor=0.0):
        gx, gy, gz = self.global_origin
        return (x_uor - gx) / self.uor_per_master, (y_uor - gy) / self.uor_per_master, (z_uor - gz) / self.uor_per_master

    def _init_pyproj(self, epsg_code):
        self.fallback_epsg = epsg_code
        self.to_wgs84 = None
        self.from_wgs84 = None
        if HAS_PYPROJ:
            try:
                self.to_wgs84 = Transformer.from_crs(f"EPSG:{epsg_code}", "EPSG:4326", always_xy=True)
                self.from_wgs84 = Transformer.from_crs("EPSG:4326", f"EPSG:{epsg_code}", always_xy=True)
            except Exception:
                pass

    def uor_to_wgs84(self, x_uor, y_uor, z_uor=0.0):
        # NOTE: DgnGCS returns lon/lat in the GCS datum (e.g. NAD83). NAD83 vs WGS84
        # differs by ~1 m in CONUS, which is acceptable for StreamStats/NHD queries.
        if IN_MICROSTATION and self.dgn_gcs is not None:
            try:
                uor_pt = DPoint3d(x_uor, y_uor, z_uor)
                ll_pt = GeoPoint()
                if _status_ok(self.dgn_gcs.LatLongFromUors(ll_pt, uor_pt)):
                    return float(ll_pt.longitude), float(ll_pt.latitude)
            except Exception:
                pass

        x_m, y_m, _ = self.uor_to_master(x_uor, y_uor, z_uor)
        if self.to_wgs84:
            lon, lat = self.to_wgs84.transform(x_m, y_m)
            return float(lon), float(lat)
        raise RuntimeError("No DGN GCS assigned and pyproj is unavailable; cannot transform UOR to WGS84.")

    def wgs84_to_uor(self, lon, lat, elev=0.0):
        if IN_MICROSTATION and self.dgn_gcs is not None:
            try:
                ll_pt = GeoPoint()
                ll_pt.Init(lon, lat, elev)
                uor_pt = DPoint3d()
                if _status_ok(self.dgn_gcs.UorsFromLatLong(uor_pt, ll_pt)):
                    return float(uor_pt.x), float(uor_pt.y), float(uor_pt.z)
            except Exception:
                pass

        if self.from_wgs84:
            xm, ym = self.from_wgs84.transform(lon, lat)
            return self.master_to_uor(float(xm), float(ym), 0.0)
        raise RuntimeError("No DGN GCS assigned and pyproj is unavailable; cannot transform WGS84 to UOR.")


# ---------------------------------------------------------------------------
# 6. USGS NHDPlus HR & StreamStats REST Client
# ---------------------------------------------------------------------------
class USGSWaterServicesClient:
    NHDPLUS_HR_URL = "https://hydro.nationalmap.gov/arcgis/rest/services/NHDPlus_HR/MapServer/3/query"
    NHD_STANDARD_URL = "https://hydro.nationalmap.gov/arcgis/rest/services/nhd/MapServer/6/query"
    # Legacy "streamstatsservices" was retired by USGS (HTTP 404). Current service chain:
    STREAMSTATS_SNAP_URL = "https://streamstats.usgs.gov/pourpoint/v2/snap/streamgrid/{region}"
    STREAMSTATS_DELINEATE_URL = "https://streamstats.usgs.gov/ss-delineate/v1/delineate/sshydro/{region}"
    STREAMSTATS_BASINCHAR_URL = "https://streamstats.usgs.gov/ss-hydro/v1/basin-characteristics/calculate"
    NSS_BASE = "https://streamstats.usgs.gov/nssservices"
    NSS_PEAKFLOW_GROUP = 2
    LFP_FEATURES_URL = "https://streamstats.usgs.gov/ss-hydro/v1/additional-features/calculate"
    NLDI_BASE = "https://api.water.usgs.gov/nldi/linked-data"
    NLDI_FLOWTRACE_URL = "https://api.water.usgs.gov/nldi/pygeoapi/processes/nldi-flowtrace/execution"
    MISSING_VALUE = -999.0

    def __init__(self, log_callback=None):
        self.log = log_callback if log_callback else print
        self.headers = {"User-Agent": "MicroStation-USGS-VDOT-Automation/4.0", "Accept": "application/json"}
        # One pooled session reuses keep-alive TLS connections across the many calls in a batch.
        self.http = requests.Session() if HAS_REQUESTS else None

    def _http_get_json(self, url, params, timeout=45, max_retries=3):
        q_str = urllib.parse.urlencode(params)
        full_url = f"{url}?{q_str}"
        for attempt in range(1, max_retries + 1):
            try:
                if HAS_REQUESTS:
                    resp = self.http.get(url, params=params, headers=self.headers, timeout=timeout)
                    if resp.status_code in (500, 502, 503, 504) and attempt < max_retries:
                        time.sleep(2 ** attempt)
                        continue
                    resp.raise_for_status()
                    return resp.json()
                else:
                    req = urllib.request.Request(full_url, headers=self.headers)
                    with urllib.request.urlopen(req, timeout=timeout) as response:
                        return json.loads(response.read().decode("utf-8"))
            except Exception as exc:
                if attempt < max_retries:
                    time.sleep(2 ** attempt)
                else:
                    raise RuntimeError(f"USGS API error: {exc}")

    # -----------------------------------------------------------------
    # Flow (Raindrop) Path - downstream trace through the NHDPlus network
    # -----------------------------------------------------------------
    @staticmethod
    def geo_length_ft(path):
        total = 0.0
        for (lon1, lat1), (lon2, lat2) in zip(path, path[1:]):
            p1, p2 = math.radians(lat1), math.radians(lat2)
            dp, dl = p2 - p1, math.radians(lon2 - lon1)
            a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
            total += 2.0 * 20902231.0 * math.asin(min(1.0, math.sqrt(a)))
        return total

    @staticmethod
    def _geom_points(geom):
        """GeoJSON LineString / MultiLineString -> single ordered [(lon, lat), ...]."""
        if not geom:
            return []
        gtype, coords = geom.get("type"), geom.get("coordinates") or []
        parts = [coords] if gtype == "LineString" else coords if gtype == "MultiLineString" else []
        out = []
        for part in parts:
            for c in part:
                pt = (float(c[0]), float(c[1]))
                if not out or out[-1] != pt:
                    out.append(pt)
        return out

    @staticmethod
    def _project_onto_path(lon, lat, path):
        """Returns (segment_index, (lon, lat)) of the closest point on `path` to (lon, lat)."""
        kx = math.cos(math.radians(lat))
        best = (float("inf"), 0, path[0])
        for i, ((x1, y1), (x2, y2)) in enumerate(zip(path, path[1:])):
            dx, dy = (x2 - x1) * kx, (y2 - y1)
            seg2 = dx * dx + dy * dy
            t = 0.0 if seg2 == 0 else max(0.0, min(1.0, (((lon - x1) * kx) * dx + (lat - y1) * dy) / seg2))
            px, py = x1 + (x2 - x1) * t, y1 + (y2 - y1) * t
            d = ((px - lon) * kx) ** 2 + (py - lat) ** 2
            if d < best[0]:
                best = (d, i, (px, py))
        return best[1], best[2]

    def _nldi_flowtrace(self, lon, lat):
        """USGS NLDI raindrop trace (DEM flow-direction path from the point to the nearest NHD flowline).
        Returns the overland path [(lon, lat), ...] or None when the service is unavailable."""
        body = {"inputs": [
            {"id": "lat", "value": f"{lat:.7f}", "type": "text/plain"},
            {"id": "lon", "value": f"{lon:.7f}", "type": "text/plain"},
            {"id": "raindroptrace", "value": "true", "type": "text/plain"},
            {"id": "direction", "value": "down", "type": "text/plain"},
        ]}
        hdr = dict(self.headers)
        hdr["Content-Type"] = "application/json"
        data = json.dumps(body).encode("utf-8")
        if HAS_REQUESTS:
            resp = self.http.post(self.NLDI_FLOWTRACE_URL, data=data, headers=hdr, timeout=45)
            resp.raise_for_status()
            result = resp.json()
        else:
            req = urllib.request.Request(self.NLDI_FLOWTRACE_URL, data=data, headers=hdr)
            with urllib.request.urlopen(req, timeout=45) as r:
                result = json.loads(r.read().decode("utf-8"))
        feats = (result or {}).get("features") or []
        # Match the feature id exactly: 'downstreamFlowline' also carries a 'raindrop_pathDist' property.
        ranked = sorted(feats, key=lambda f: str(f.get("id", "")).lower() != "raindroppath")
        for feat in ranked:
            fid = str(feat.get("id", "")).lower()
            if fid == "raindroppath" or ("raindrop" in fid and "flowline" not in fid):
                pts = self._geom_points(feat.get("geometry"))
                if len(pts) >= 2:
                    if math.hypot(pts[-1][0] - lon, pts[-1][1] - lat) < math.hypot(pts[0][0] - lon, pts[0][1] - lat):
                        pts = pts[::-1]
                    if pts[0] != (lon, lat):
                        pts = [(lon, lat)] + pts
                    return pts
        return None

    def trace_raindrop_path(self, lon, lat, max_miles=1.0, clipped_streams=None):
        """Traces water from (lon, lat) overland to the nearest NHDPlus flowline, then down the mainstem
        (USGS NLDI 'DM' navigation) for up to `max_miles`. Network only - safe on a worker thread."""
        lon, lat = float(lon), float(lat)
        overland, overland_method = None, ""
        self.log("Raindrop trace: requesting NLDI overland flow path...")
        try:
            overland = self._nldi_flowtrace(lon, lat)
            if overland:
                overland_method = "NLDI raindrop trace (DEM flow direction)"
        except Exception as exc:
            self.log(f"  Note: NLDI raindrop service unavailable ({str(exc)[:90]}); "
                     "using a straight overland leg to the nearest flowline.")

        reaches, source = [], "USGS NLDI downstream mainstem"
        try:
            hit = overland[-1] if overland else (lon, lat)
            pos = self._http_get_json(f"{self.NLDI_BASE}/comid/position",
                                      {"f": "json", "coords": f"POINT({hit[0]:.7f} {hit[1]:.7f})"},
                                      timeout=60, max_retries=2)
            feats = pos.get("features") or []
            if not feats:
                raise RuntimeError("no NHDPlus catchment found at this point")
            comid = str(feats[0]["properties"].get("comid"))
            self.log(f"  Start reach NHDPlus COMID {comid}; navigating downstream {max_miles:g} mi...")
            fl = self._http_get_json(f"{self.NLDI_BASE}/comid/{comid}/navigation/DM/flowlines",
                                     {"f": "json", "distance": f"{max(0.1, max_miles) * 1.609344:.3f}"},
                                     timeout=90, max_retries=2)
            lines = []
            for f in fl.get("features") or []:
                pts = self._geom_points(f.get("geometry"))
                if len(pts) >= 2:
                    lines.append({"comid": str((f.get("properties") or {}).get("nhdplus_comid", "")), "pts": pts})
            lines = self._order_downstream(lines, comid)
            reaches = [{"comid": ln["comid"], "name": "", "pts": ln["pts"]} for ln in lines]
        except Exception as exc:
            if not clipped_streams:
                raise RuntimeError(f"Downstream network trace failed: {exc}")
            self.log(f"  Note: NLDI navigation failed ({str(exc)[:90]}); tracing the fence NHD streams instead.")
            reaches = self._trace_clipped_streams(overland[-1] if overland else (lon, lat), clipped_streams)
            source = "Fence NHDPlus HR streams (endpoint connectivity)"
        if not reaches:
            raise RuntimeError("No downstream flowlines were found for this point.")

        hit = overland[-1] if overland else (lon, lat)
        seg_i, proj = self._project_onto_path(hit[0], hit[1], reaches[0]["pts"])
        reaches[0]["pts"] = [proj] + reaches[0]["pts"][seg_i + 1:]
        if not overland:
            overland = [(lon, lat), proj]
            overland_method = "Straight line to nearest flowline (approximation)"
        elif overland[-1] != proj:
            overland = overland + [proj]

        cum = self.geo_length_ft(overland)
        overland_ft = cum
        for r in reaches:
            r["length_ft"] = self.geo_length_ft(r["pts"])
            cum += r["length_ft"]
            r["cum_ft"] = cum
        path = list(overland)
        for r in reaches:
            path.extend(r["pts"][1:] if path and r["pts"][0] == path[-1] else r["pts"])
        return {
            "query_point": (lon, lat), "overland_wgs84": overland, "overland_method": overland_method,
            "overland_ft": overland_ft, "reaches": reaches, "path_wgs84": path, "total_ft": cum,
            "source": source, "max_miles": float(max_miles),
        }

    @staticmethod
    def _order_downstream(lines, start_comid):
        """Chains flowlines upstream->downstream by matching end/start vertices; keeps service order otherwise."""
        if len(lines) < 2:
            return lines
        key = lambda p: (round(p[0], 5), round(p[1], 5))
        by_start = {key(ln["pts"][0]): ln for ln in lines}
        cur = next((ln for ln in lines if ln["comid"] == str(start_comid)), lines[0])
        ordered, seen = [cur], {id(cur)}
        while True:
            nxt = by_start.get(key(cur["pts"][-1]))
            if nxt is None or id(nxt) in seen:
                break
            ordered.append(nxt)
            seen.add(id(nxt))
            cur = nxt
        return ordered if len(ordered) == len(lines) else lines

    def _trace_clipped_streams(self, start, clipped_streams):
        """Offline fallback: walks the fence-clipped NHD lines downstream by endpoint connectivity."""
        parts = []
        for s in clipped_streams or []:
            for ln in s.get("clipped_lines_wgs84") or []:
                if len(ln) >= 2:
                    parts.append({"comid": str(s.get("id", "")), "name": str(s.get("name", "")),
                                  "pts": [(float(p[0]), float(p[1])) for p in ln]})
        if not parts:
            return []
        kx = math.cos(math.radians(start[1]))

        def _d2(p):
            return ((p[0] - start[0]) * kx) ** 2 + (p[1] - start[1]) ** 2

        first = min(parts, key=lambda pr: _d2(USGSWaterServicesClient._project_onto_path(start[0], start[1], pr["pts"])[1]))
        ordered, seen, cur = [first], {id(first)}, first
        tol = 1e-5
        while True:
            end = cur["pts"][-1]
            nxt = next((pr for pr in parts if id(pr) not in seen
                        and abs(pr["pts"][0][0] - end[0]) < tol and abs(pr["pts"][0][1] - end[1]) < tol), None)
            if nxt is None:
                break
            ordered.append(nxt)
            seen.add(id(nxt))
            cur = nxt
        return ordered

    def query_and_clip_nhd_streams(self, fence_wgs84_poly):
        lons = [pt[0] for pt in fence_wgs84_poly]
        lats = [pt[1] for pt in fence_wgs84_poly]
        params = {
            "geometry": f"{min(lons)},{min(lats)},{max(lons)},{max(lats)}",
            "geometryType": "esriGeometryEnvelope", "inSR": "4326",
            "spatialRel": "esriSpatialRelIntersects", "outFields": "*",
            "returnGeometry": "true", "outSR": "4326", "f": "geojson",
        }
        self.log("Querying USGS NHDPlus HR flowlines in fence bounding box...")
        try:
            geojson = self._http_get_json(self.NHDPLUS_HR_URL, params, timeout=90, max_retries=2)
            if "error" in geojson:
                raise RuntimeError(geojson["error"])
            features = geojson.get("features", [])
        except Exception as exc:
            self.log(f"  NHDPlus HR unavailable ({exc}); falling back to medium-resolution NHD.")
            geojson = self._http_get_json(self.NHD_STANDARD_URL, params, timeout=90, max_retries=2)
            features = geojson.get("features", [])
        if geojson.get("exceededTransferLimit") or (geojson.get("properties") or {}).get("exceededTransferLimit"):
            self.log("  [WARNING] Server record limit reached - some flowlines may be missing. Use a smaller fence.")

        clipped_streams = []
        for idx, feat in enumerate(features):
            props = feat.get("properties") or {}
            lp = {str(k).lower(): v for k, v in props.items()}
            geom = feat.get("geometry") or {}
            coords = geom.get("coordinates", [])
            gtype = geom.get("type", "")

            raw_lines = [coords] if gtype == "LineString" else coords if gtype == "MultiLineString" else []
            clipped_parts = []
            for rline in raw_lines:
                segs = PurePythonClipper.clip_polyline_to_polygon([(c[0], c[1]) for c in rline], fence_wgs84_poly)
                clipped_parts.extend(segs)

            if not clipped_parts:
                continue

            name = props.get("GNIS_Name") or props.get("gnis_name") or f"Tributary #{idx+1}"
            order = props.get("StreamOrde") or props.get("streamorde") or 1
            tot_da = props.get("TotDASqKm") or props.get("totdasqkm") or 0.0

            total_len = 0.0
            for part in clipped_parts:
                for k in range(len(part) - 1):
                    lon1, lat1 = part[k]
                    lon2, lat2 = part[k + 1]
                    dx = (lon2 - lon1) * 364000.0 * math.cos(math.radians(0.5 * (lat1 + lat2)))
                    dy = (lat2 - lat1) * 364000.0
                    total_len += math.hypot(dx, dy)

            try:
                order_int = max(int(float(order)), 1)
            except (TypeError, ValueError):
                order_int = 1
            try:
                tot_da_f = float(tot_da)
            except (TypeError, ValueError):
                tot_da_f = 0.0

            clipped_streams.append({
                "id": idx + 1, "name": name,
                "stream_order": order_int,
                "tot_da_sqkm": tot_da_f,
                "length_ft": round(total_len, 1),
                "clipped_lines_wgs84": clipped_parts,
                "pour_point_wgs84": clipped_parts[-1][-1],
                "nhd": {k: lp.get(k) for k in self.NHD_ITEM_FIELDS},
            })

        by_seq = {}
        for s in clipped_streams:
            hs = s["nhd"].get("hydroseq")
            if hs not in (None, "", 0):
                by_seq[hs] = s["name"]
        for s in clipped_streams:
            dn = s["nhd"].get("dnhydroseq")
            if dn in (None, "", 0):
                s["flows_into"] = "Network outlet" if dn == 0 else ""
            else:
                s["flows_into"] = by_seq.get(dn, "Downstream reach outside fence")

        clipped_streams.sort(key=lambda s: (s["tot_da_sqkm"], s["stream_order"], s["length_ft"]), reverse=True)
        return clipped_streams

    # NHDPlus HR attributes kept for the stream Item Type (qema = EROM mean annual flow, vema = velocity).
    NHD_ITEM_FIELDS = ("nhdplusid", "permanent_identifier", "reachcode", "ftype", "fcode",
                       "qema", "vema", "slope", "hydroseq", "dnhydroseq", "flowdir")

    def _http_post_json(self, url, body, params=None, timeout=240, max_retries=3):
        full_url = f"{url}?{urllib.parse.urlencode(params)}" if params else url
        headers = dict(self.headers)
        headers.update({"Content-Type": "application/json", "X-Is-StreamStats": "true"})
        data = json.dumps(body).encode("utf-8")
        for attempt in range(1, max_retries + 1):
            try:
                if HAS_REQUESTS:
                    resp = self.http.post(full_url, data=data, headers=headers, timeout=timeout)
                    if resp.status_code in (500, 502, 503, 504) and attempt < max_retries:
                        time.sleep(2 ** attempt)
                        continue
                    resp.raise_for_status()
                    return resp.json()
                req = urllib.request.Request(full_url, data=data, headers=headers, method="POST")
                with urllib.request.urlopen(req, timeout=timeout) as response:
                    return json.loads(response.read().decode("utf-8"))
            except Exception as exc:
                if attempt < max_retries:
                    time.sleep(2 ** attempt)
                else:
                    raise RuntimeError(f"USGS API error ({url}): {exc}")

    def _snap_pour_point(self, region, lon, lat):
        try:
            data = self._http_get_json(self.STREAMSTATS_SNAP_URL.format(region=region),
                                       {"pt": f"POINT({lon:.6f} {lat:.6f})"}, timeout=60, max_retries=2)
            coords = ((data or {}).get("output") or {}).get("coordinates") or []
            if data.get("scope") and len(coords) >= 2:
                return float(coords[0]), float(coords[1]), True
        except Exception as exc:
            self.log(f"  [WARNING] Pour-point snap failed ({exc}); using the original point.")
        return lon, lat, False

    def delineate_streamstats_watershed(self, rcode, lon, lat, user_c_override=None, include_lfp=False):
        region = rcode.upper().strip()
        self.log(f"Calling USGS StreamStats ({region}, Lon={lon:.5f}, Lat={lat:.5f})...")

        s_lon, s_lat, snapped = self._snap_pour_point(region, lon, lat)
        if snapped:
            self.log(f"  Pour point snapped to stream grid: Lon={s_lon:.5f}, Lat={s_lat:.5f}")
        else:
            self.log("  [WARNING] Point did not snap to a StreamStats stream cell; results may be invalid.")

        ws_data = self._http_get_json(self.STREAMSTATS_DELINEATE_URL.format(region=region),
                                      {"lat": f"{s_lat:.6f}", "lon": f"{s_lon:.6f}"}, timeout=240, max_retries=2)
        wsresp = ((ws_data or {}).get("bcrequest") or {}).get("wsresp") or {}
        ws_id = wsresp.get("workspace_id", "") or ""

        fc_items = []
        for item in wsresp.get("featurecollection", []) or []:
            fc_items.extend(item if isinstance(item, list) else [item])

        basin_rings = []
        basin_geom = None
        snapped_pt = (s_lon, s_lat)
        for fc in fc_items:
            if not isinstance(fc, dict):
                continue
            fc_name = str(fc.get("name", "")).lower()
            for feat in (fc.get("feature") or {}).get("features", []) or []:
                geom = feat.get("geometry") or {}
                gtype = geom.get("type", "")
                coords = geom.get("coordinates", [])
                if "point" in fc_name and gtype == "Point" and len(coords) >= 2:
                    snapped_pt = (coords[0], coords[1])
                    warn = (feat.get("properties") or {}).get("WarningMsg")
                    if warn:
                        self.log(f"  [WARNING] StreamStats: {warn}")
                elif "watershed" in fc_name and gtype in ("Polygon", "MultiPolygon"):
                    if basin_geom is None:
                        basin_geom = geom
                    polys = [coords] if gtype == "Polygon" else coords
                    for poly in polys:
                        for ring in poly:
                            basin_rings.append([(c[0], c[1]) for c in ring])
        if not basin_rings:
            raise RuntimeError("StreamStats returned no watershed polygon for this point.")

        params_list = []
        bc_wait = self._down_minutes("bc")
        if bc_wait:
            self.log(f"  StreamStats basin characteristics are down (retrying in {bc_wait:.0f} min).")
        else:
            try:
                self.log("  Computing basin characteristics...")
                ws_data.setdefault("bcrequest", {})["bcLabels"] = "*"
                bc_data = self._http_post_json(self.STREAMSTATS_BASINCHAR_URL, ws_data, max_retries=2)
                for p in bc_data or []:
                    try:
                        val = float(p.get("value"))
                    except (TypeError, ValueError):
                        continue
                    if val <= self.MISSING_VALUE:
                        continue
                    params_list.append({"code": p.get("code", ""), "name": p.get("name", ""),
                                        "value": val, "unit": p.get("unit", "")})
            except Exception as exc:
                self._mark_down("bc", exc)
                self.log(f"  [WARNING] Basin characteristics failed: {exc}")

        bc_from_usgs = bool(params_list)
        nhd_lfp = None
        nlcd = hsg = None
        hsg_source = ""
        try:
            nlcd = self._nlcd_basin_cover(basin_geom)
            self.log(f"  NLCD 2021 (MRLC): impervious {nlcd['imperv_pct']:.1f}%, forest {nlcd['forest_pct']:.1f}%, "
                     f"developed {nlcd['dev_pct']:.1f}% from {nlcd['cells']:,} cells in the basin.")
        except Exception as exc:
            self.log(f"  [WARNING] MRLC NLCD land cover unavailable ({exc}).")
        # Soil groups feed the NRCS curve number and the drainage-area Item Type, so they are read for every basin.
        if basin_geom is not None:
            hsg, hsg_source = self.basin_soil_groups(basin_geom)
            if hsg:
                self.log(f"  Soil groups ({hsg_source.split(';')[0]}): "
                         + ", ".join(f"{k} {100.0 * v:.0f}%" for k, v in sorted(hsg.items())) + ".")
        if not any(str(p.get("code", "")).upper() == "DRNAREA" for p in params_list):
            area = self._polygon_area_sqmi(basin_geom)
            if area > 0:
                params_list.append({"code": "DRNAREA", "name": "Drainage Area (computed from basin polygon)",
                                    "value": round(area, 4), "unit": "square miles"})
                self.log(f"  Drainage area computed from the basin polygon: {area * 640.0:,.1f} ac "
                         f"({area:.4f} sq mi).")
        if not bc_from_usgs and basin_geom is not None:
            nhd_lfp = self._alternate_basin_characteristics(basin_geom, basin_rings, snapped_pt, params_list, nlcd)

        flow_stats = {}
        if basin_geom is not None and params_list:
            # Alternate-source slope/land cover are approximations, so only the drainage area feeds the regression.
            nss_params = params_list if bc_from_usgs else [
                p for p in params_list if str(p.get("code", "")).upper() == "DRNAREA"]
            nss_wait = self._down_minutes("nss")
            if nss_wait:
                self.log(f"  USGS NSS regression service is down (retrying in {nss_wait:.0f} min); "
                         "no USGS regression flows for this culvert.")
            else:
                try:
                    flow_stats = self._compute_peak_flows(region, basin_geom, nss_params)
                except Exception as exc:
                    skipped = self._mark_down("nss", exc)
                    self.log(f"  [WARNING] Peak-flow regression (NSS) failed: {exc}."
                             + (" It will be skipped for 10 min so the batch keeps moving." if skipped else ""))

        ss_result = {
            "rcode": region, "workspace_id": ws_id,
            "query_point": (lon, lat), "snapped_point": snapped_pt,
            "basin_rings_wgs84": basin_rings, "parameters": params_list, "flow_stats": flow_stats,
            "bc_source": "USGS StreamStats" if bc_from_usgs else "Alternate (MRLC NLCD 2021 / USGS 3DEP / basin polygon)",
            "nlcd": nlcd, "hsg": hsg, "hsg_source": hsg_source,
        }
        if include_lfp:
            try:
                lfp, method = self.compute_longest_flowpath(ws_data, basin_rings, snapped_pt, nhd_lfp=nhd_lfp)
                if lfp:
                    length_ft = self._line_len_ft(lfp)
                    ss_result.update({"longest_flowpath_wgs84": lfp, "lfp_length_ft": round(length_ft, 1),
                                      "lfp_method": method})
                    self.log(f"  Longest flowpath: {length_ft:,.0f} ft ({length_ft / 5280.0:.2f} mi) via {method}.")
                else:
                    self.log("  [WARNING] Longest flowpath could not be determined.")
            except Exception as exc:
                self.log(f"  [WARNING] Longest flowpath failed: {exc}")
        elif nhd_lfp:
            ss_result["tc_flow_length_ft"] = round(self._line_len_ft(nhd_lfp), 1)
        try:
            ss_result["atlas14"] = self._atlas14_intensities(snapped_pt[0], snapped_pt[1])
        except Exception as exc:
            self.log(f"  [WARNING] NOAA Atlas 14 rainfall unavailable ({exc}); using built-in IDF coefficients.")
        vcomp = VDOTHydrologyCalculator.compute_comparison_table(ss_result, user_c_override)
        ss_result["vdot_comparison"] = vcomp
        self.log(f"  Tc {vcomp.get('tc_min')} min ({str(vcomp.get('tc_length_source', '')).lower()}); "
                 f"rainfall: {vcomp.get('rainfall_source')}; basin data: {vcomp.get('basin_source')}.")
        ran = [m["label"] for m in vcomp.get("methods", []) if m["calculated"]]
        self.log(f"  Peak-flow methods calculated: {', '.join(ran) or 'none'}. "
                 f"Design flow: {vcomp.get('governing_method')}."
                 + (f" NRCS CN {vcomp['cn']:.0f} ({vcomp.get('cn_source')})." if vcomp.get("cn") else ""))
        return ss_result

    ATLAS14_URL = "https://hdsc.nws.noaa.gov/cgi-bin/hdsc/new/cgi_readH5.py"
    ATLAS14_ARI = (1, 2, 5, 10, 25, 50, 100, 200, 500, 1000)
    ATLAS14_DURATIONS_MIN = (5, 10, 15, 30, 60, 120, 180, 360, 720, 1440)

    def _atlas14_intensities(self, lon, lat):
        """NOAA Atlas 14 precipitation-frequency intensities (in/hr, partial-duration series) at a point."""
        txt = self._http_get_text(self.ATLAS14_URL, {
            "lat": f"{lat:.5f}", "lon": f"{lon:.5f}", "type": "pf", "data": "intensity",
            "units": "english", "series": "pds"}, timeout=45)
        m = re.search(r"quantiles\s*=\s*(\[\[.*?\]\])\s*;", txt, re.S)
        if not m:
            raise RuntimeError("no Atlas 14 data for this location")
        rows = [[float(v) for v in re.findall(r"'([^']*)'", r)] for r in re.findall(r"\[([^\[\]]+)\]", m.group(1))]
        nd, na = len(self.ATLAS14_DURATIONS_MIN), len(self.ATLAS14_ARI)
        if len(rows) < nd or any(len(r) < na for r in rows[:nd]):
            raise RuntimeError("unexpected Atlas 14 table size")
        return {"source": "NOAA Atlas 14, partial-duration series", "ari": list(self.ATLAS14_ARI),
                "durations_min": list(self.ATLAS14_DURATIONS_MIN), "intensity": [r[:na] for r in rows[:nd]]}

    # (label, URL, what stops working when it is down). Any HTTP status below 500 counts as responding.
    SERVICE_CHECKS = (
        ("StreamStats delineation", "https://streamstats.usgs.gov/ss-delineate/v1/",
         "watersheds cannot be delineated"),
        ("StreamStats pour-point snap", "https://streamstats.usgs.gov/pourpoint/v2/",
         "pour points are not snapped to the stream grid"),
        ("StreamStats basin characteristics", "https://streamstats.usgs.gov/ss-hydro/v1/",
         "slope and land cover come from MRLC NLCD / USGS 3DEP; the longest flowpath is traced from NHD"),
        ("USGS NSS regression", "https://streamstats.usgs.gov/nssservices/regions",
         "no USGS regression peak flows"),
        ("USGS NHD streams", "https://hydro.nationalmap.gov/arcgis/rest/services/nhd/MapServer?f=json",
         "streams cannot be loaded or traced"),
        ("USGS NLDI (raindrop path)", "https://api.water.usgs.gov/nldi/linked-data/",
         "raindrop paths cannot be traced"),
        ("MRLC NLCD land cover", "https://www.mrlc.gov/geoserver/ows?service=WCS&version=2.0.1&request="
                                 "DescribeCoverage&coverageId=mrlc_download__NLCD_2021_Impervious_L48",
         "no alternate impervious/forest values"),
        ("USGS 3DEP elevation", "https://elevation.nationalmap.gov/arcgis/rest/services/3DEPElevation/"
                                "ImageServer?f=json", "no alternate channel slope"),
        ("NOAA Atlas 14 rainfall", "https://hdsc.nws.noaa.gov/cgi-bin/hdsc/new/cgi_readH5.py?lat=37.5&lon=-78.5"
                                   "&type=pf&data=intensity&units=english&series=pds",
         "built-in IDF coefficients are used and the NRCS method is not calculated"),
        ("USDA SSURGO soils", "https://sdmdataaccess.sc.egov.usda.gov/",
         "soil groups come from the coarser UC Davis SoilWeb grid"),
        ("UC Davis SoilWeb soil grid (backup)", "https://casoilresource.lawr.ucdavis.edu/wcs/issr800.json",
         "no backup soil groups if USDA SSURGO is also down"),
        # The regression-region lookup is a separate back end from the NSS list above; it can hang on its own.
        ("USGS NSS region lookup", "https://streamstats.usgs.gov/nssservices/regressionregions/bylocation",
         "no USGS regression peak flows (each culvert skips them after one 45 s wait)",
         {"type": "Polygon", "coordinates": [[[-78.50, 37.50], [-78.49, 37.50], [-78.49, 37.51],
                                              [-78.50, 37.51], [-78.50, 37.50]]]}),
    )

    def check_services(self, timeout=15):
        """Probes every web service the app relies on. Returns [(label, ok, detail, impact)], where detail
        includes the response time. Network only."""
        results = [None] * len(self.SERVICE_CHECKS)

        def _probe(i, label, url, impact, body=None):
            headers = {"User-Agent": self.headers["User-Agent"]}
            data = None
            if body is not None:
                headers["Content-Type"] = "application/json"
                data = json.dumps(body).encode("utf-8")
            t0 = time.time()
            try:
                with urllib.request.urlopen(urllib.request.Request(url, data=data, headers=headers),
                                            timeout=timeout) as resp:
                    resp.read(1)
                    ok, detail = True, f"HTTP {resp.status}"
            except urllib.error.HTTPError as exc:
                ok, detail = exc.code < 500, f"HTTP {exc.code}"
            except Exception as exc:
                ok, detail = False, str(getattr(exc, "reason", exc))[:60]
            results[i] = (label, ok, f"{detail}, {time.time() - t0:.1f} s", impact)

        threads = [threading.Thread(target=_probe, args=(i, *chk), daemon=True)
                   for i, chk in enumerate(self.SERVICE_CHECKS)]
        for t in threads:
            t.start()
        for t in threads:
            t.join(timeout + 3)
        return [r or (chk[0], False, f"no answer in {timeout} s", chk[2])
                for r, chk in zip(results, self.SERVICE_CHECKS)]

    # Remembers services that just failed so a batch does not wait on them again for every culvert.
    OUTAGE_SKIP_SEC = 600
    _outages = {}

    def _down_minutes(self, key):
        """Minutes left before a failed service is tried again (0 = try it now)."""
        return max(0.0, (type(self)._outages.get(key, 0.0) - time.time()) / 60.0)

    @staticmethod
    def _is_outage_error(exc):
        """True for server errors, timeouts and dropped connections (not for a problem with one basin)."""
        msg = str(exc).lower()
        return ("timed out" in msg or "timeout" in msg or "connection" in msg
                or any(f"error {c}" in msg or f" {c} " in msg for c in (500, 502, 503, 504)))

    def _mark_down(self, key, exc=None):
        if exc is None or self._is_outage_error(exc):
            type(self)._outages[key] = time.time() + self.OUTAGE_SKIP_SEC
            return True
        return False

    @classmethod
    def clear_outages(cls):
        cls._outages.clear()
        cls._lfp_down_until = 0.0

    @classmethod
    def apply_service_check(cls, results):
        """Uses a finished check_services() result so the next run skips the services found down at once."""
        cls.clear_outages()
        until = time.time() + cls.OUTAGE_SKIP_SEC
        for label, ok, _detail, _impact in results or []:
            if ok:
                continue
            if label == "StreamStats basin characteristics":
                cls._outages["bc"] = until
                cls._lfp_down_until = until
            elif label.startswith("USGS NSS"):
                cls._outages["nss"] = until
            elif label == "USDA SSURGO soils":
                cls._outages["ssurgo"] = until

    def _compute_peak_flows(self, region, basin_geom, params_list):
        self.log("  Looking up NSS regression regions...")
        rr_list = self._http_post_json(f"{self.NSS_BASE}/regressionregions/bylocation", basin_geom,
                                       timeout=45, max_retries=1) or []
        if not rr_list:
            return {}
        weights = {str(r.get("code", "")).upper(): r.get("percentWeight") for r in rr_list}
        scenarios = self._http_get_json(f"{self.NSS_BASE}/scenarios", {
            "regions": f"{region},NA", "statisticgroups": self.NSS_PEAKFLOW_GROUP,
            "regressionregions": ",".join(r["code"] for r in rr_list if r.get("code")),
        }, timeout=45, max_retries=2) or []
        if not scenarios:
            return {}

        bc_map = {str(p["code"]).upper(): p["value"] for p in params_list}
        missing = set()
        for sc in scenarios:
            regs = [g for g in sc.get("regressionRegions", []) or [] if "area-averaged" not in str(g.get("name", "")).lower()]
            for g in regs:
                g.pop("results", None)
                g["percentWeight"] = weights.get(str(g.get("code", "")).upper())
                for p in g.get("parameters", []) or []:
                    code = str(p.get("code", "")).upper()
                    if code not in bc_map:
                        missing.add(code)
                    p["value"] = bc_map.get(code)
            sc["regressionRegions"] = regs

        if missing:
            self.log(f"  [WARNING] Skipping USGS peak flows - missing basin characteristic(s): {', '.join(sorted(missing))}. "
                     "Move the pour point onto the stream centerline.")
            return {}

        self.log("  Estimating peak flows (NSS)...")
        est = self._http_post_json(f"{self.NSS_BASE}/scenarios/estimate", scenarios, params={"regions": f"{region},NA"},
                                   timeout=60, max_retries=1)
        return self._parse_flow_stats(est)

    @staticmethod
    def _polygon_area_sqmi(geom):
        """Area of a GeoJSON (Multi)Polygon in WGS84, in sq mi (local equal-area projection; holes subtracted)."""
        if not isinstance(geom, dict):
            return 0.0
        coords = geom.get("coordinates") or []
        polys = [coords] if geom.get("type") == "Polygon" else coords if geom.get("type") == "MultiPolygon" else []
        r_mi = 3958.7613
        total = 0.0
        for poly in polys:
            for i, ring in enumerate(poly or []):
                if len(ring) < 3:
                    continue
                lat0 = math.radians(sum(c[1] for c in ring) / len(ring))
                pts = [(math.radians(c[0]) * math.cos(lat0) * r_mi, math.radians(c[1]) * r_mi) for c in ring]
                a = abs(sum(pts[k][0] * pts[k + 1][1] - pts[k + 1][0] * pts[k][1]
                            for k in range(len(pts) - 1))) / 2.0
                total += -a if i > 0 else a
        return max(total, 0.0)

    # -----------------------------------------------------------------
    # Alternate basin characteristics (used when the StreamStats ss-hydro service is down)
    # -----------------------------------------------------------------
    MRLC_WCS_URL = "https://www.mrlc.gov/geoserver/ows"
    NLCD_LANDCOVER_COVERAGE = "mrlc_download__NLCD_2021_Land_Cover_L48"
    NLCD_IMPERV_COVERAGE = "mrlc_download__NLCD_2021_Impervious_L48"
    NLCD_FOREST_CLASSES = (41, 42, 43)
    NLCD_MAX_CELLS = 500          # longest grid side; larger basins are resampled
    DEM_SAMPLES_URL = "https://elevation.nationalmap.gov/arcgis/rest/services/3DEPElevation/ImageServer/getSamples"
    NLDI_COMID_URL = "https://api.water.usgs.gov/nldi/linked-data/comid/position"
    STREAMCAT_URL = "https://api.epa.gov/StreamCat/streams/metrics"

    def _http_get_text(self, url, params, timeout=90):
        headers = {"User-Agent": self.headers["User-Agent"]}
        if HAS_REQUESTS:
            resp = self.http.get(url, params=params, headers=headers, timeout=timeout)
            resp.raise_for_status()
            return resp.text
        req = urllib.request.Request(f"{url}?{urllib.parse.urlencode(params, doseq=True)}", headers=headers)
        with urllib.request.urlopen(req, timeout=timeout) as response:
            return response.read().decode("utf-8", "replace")

    def _nlcd_grid(self, coverage, bbox, scale):
        """NLCD cells for a WGS84 bbox as rows of ints (north row first) plus (min lon, min lat, max lon, max lat)."""
        params = {"service": "WCS", "version": "2.0.1", "request": "GetCoverage", "coverageId": coverage,
                  "format": "text/plain", "subsettingCrs": "http://www.opengis.net/def/crs/EPSG/0/4326",
                  "subset": [f"Long({bbox[0]:.6f},{bbox[2]:.6f})", f"Lat({bbox[1]:.6f},{bbox[3]:.6f})"]}
        if scale < 1.0:
            params["scalefactor"] = f"{scale:.4f}"
        txt = self._http_get_text(self.MRLC_WCS_URL, params)
        head, sep, body = txt.partition("Band 0:")
        m = re.search(r"Grid bounds:([^\n]*)", head)
        nums = [float(v) for v in re.findall(r"-?\d+\.\d+(?:[eE][-+]?\d+)?", m.group(1))] if m else []
        lons, lats = sorted(v for v in nums if v < 0), sorted(v for v in nums if v > 0)
        if not sep or len(lons) != 2 or len(lats) != 2:
            raise RuntimeError("unexpected response from the MRLC coverage service")
        rows = [[int(float(v)) for v in ln.split()] for ln in body.splitlines() if ln.strip()]
        if not rows or not rows[0]:
            raise RuntimeError("MRLC coverage service returned no cells")
        return rows, (lons[0], lats[0], lons[1], lats[1])

    def _nlcd_basin_cover(self, basin_geom):
        """NLCD 2021 clipped to the basin polygon: impervious %, forest %, developed %, class fractions."""
        coords = basin_geom.get("coordinates") or []
        polys = [coords] if basin_geom.get("type") == "Polygon" else coords
        rings = [r for poly in polys for r in poly if len(r) >= 4]
        xs = [c[0] for r in rings for c in r]
        ys = [c[1] for r in rings for c in r]
        pad = 0.0006
        bbox = (min(xs) - pad, min(ys) - pad, max(xs) + pad, max(ys) + pad)
        side = max((bbox[2] - bbox[0]) * math.cos(math.radians(0.5 * (bbox[1] + bbox[3]))),
                   bbox[3] - bbox[1]) * 111320.0 / 30.0
        scale = min(1.0, self.NLCD_MAX_CELLS / max(side, 1.0))
        lc, (minx, miny, maxx, maxy) = self._nlcd_grid(self.NLCD_LANDCOVER_COVERAGE, bbox, scale)
        imp, _ = self._nlcd_grid(self.NLCD_IMPERV_COVERAGE, bbox, scale)
        nrows, ncols = len(lc), min(len(r) for r in lc)
        if len(imp) != nrows or min(len(r) for r in imp) != ncols:
            raise RuntimeError("NLCD land-cover and impervious grids do not match")
        dx, dy = (maxx - minx) / ncols, (maxy - miny) / nrows

        # Scanline fill: x-crossings of every ring edge per cell-centre row (even-odd, so holes are excluded).
        crossings = [[] for _ in range(nrows)]
        for ring in rings:
            for (x1, y1), (x2, y2) in zip(((c[0], c[1]) for c in ring), ((c[0], c[1]) for c in ring[1:])):
                if y1 == y2:
                    continue
                ylo, yhi = (y1, y2) if y1 < y2 else (y2, y1)
                r0 = max(int(math.floor((maxy - yhi) / dy - 0.5)) + 1, 0)
                r1 = min(int(math.floor((maxy - ylo) / dy - 0.5)), nrows - 1)
                for r in range(r0, r1 + 1):
                    yc = maxy - (r + 0.5) * dy
                    crossings[r].append(x1 + (yc - y1) * (x2 - x1) / (y2 - y1))

        n = forest = 0
        imp_sum = 0.0
        classes = {}

        def _take(r, c):
            nonlocal n, forest, imp_sum
            if lc[r][c] == 0 or not 0 <= imp[r][c] <= 100:
                return
            n += 1
            imp_sum += imp[r][c]
            classes[lc[r][c]] = classes.get(lc[r][c], 0) + 1
            if lc[r][c] in self.NLCD_FOREST_CLASSES:
                forest += 1

        for r, xcs in enumerate(crossings):
            xcs.sort()
            for xa, xb in zip(xcs[0::2], xcs[1::2]):
                c0 = max(int(math.ceil((xa - minx) / dx - 0.5)), 0)
                c1 = min(int(math.ceil((xb - minx) / dx - 0.5)) - 1, ncols - 1)
                for c in range(c0, c1 + 1):
                    _take(r, c)
        if n == 0:   # basin smaller than one 30 m cell: use the cell at its centre
            cx, cy = sum(xs) / len(xs), sum(ys) / len(ys)
            _take(min(max(int((maxy - cy) / dy), 0), nrows - 1), min(max(int((cx - minx) / dx), 0), ncols - 1))
        if n == 0:
            raise RuntimeError("no NLCD cells inside the basin")
        return {"imperv_pct": imp_sum / n, "forest_pct": 100.0 * forest / n, "cells": n,
                "classes": {k: v / n for k, v in classes.items()},
                "dev_pct": 100.0 * sum(v for k, v in classes.items() if k in (21, 22, 23, 24)) / n}

    SSURGO_URL = "https://sdmdataaccess.sc.egov.usda.gov/Tabular/post.rest"
    HSG_SRC_SSURGO = "USDA SSURGO (Soil Data Access); dual groups counted as D"
    HSG_SRC_ISSR800 = "UC Davis SoilWeb ISSR-800 grid (800 m, from SSURGO); dual groups counted as D"
    ISSR800_WCS_URL = "https://casoilresource.lawr.ucdavis.edu/cgi-bin/mapserv"
    ISSR800_HSG_CODES = {1: "A", 2: "D", 3: "B", 4: "D", 5: "C", 6: "D", 7: "D"}   # 2, 4, 6 = A/D, B/D, C/D
    ISSR800_SAMPLES = 60       # sample cells along the longest side of the basin

    def _http_get_bytes(self, url, params, timeout=60):
        headers = {"User-Agent": self.headers["User-Agent"]}
        if HAS_REQUESTS:
            resp = self.http.get(url, params=params, headers=headers, timeout=timeout)
            resp.raise_for_status()
            return resp.content
        req = urllib.request.Request(f"{url}?{urllib.parse.urlencode(params, doseq=True)}", headers=headers)
        with urllib.request.urlopen(req, timeout=timeout) as response:
            return response.read()

    @staticmethod
    def _read_byte_geotiff(data):
        """Minimal reader for an uncompressed, single-band, 8-bit, strip-organised GeoTIFF.
        Returns (rows of ints north row first, west lon, north lat, cell width, cell height)."""
        if data[:2] not in (b"II", b"MM"):
            raise RuntimeError("soil grid service did not return a TIFF image")
        bo = "<" if data[:2] == b"II" else ">"
        ifd = struct.unpack(bo + "I", data[4:8])[0]
        sizes = {1: ("B", 1), 3: ("H", 2), 4: ("I", 4), 12: ("d", 8)}
        tags = {}
        for i in range(struct.unpack(bo + "H", data[ifd:ifd + 2])[0]):
            tag, typ, cnt, val = struct.unpack(bo + "HHI4s", data[ifd + 2 + 12 * i: ifd + 14 + 12 * i])
            if typ not in sizes:
                continue
            code, size = sizes[typ]
            n = size * cnt
            raw = val[:n] if n <= 4 else data[struct.unpack(bo + "I", val)[0]:][:n]
            tags[tag] = struct.unpack(bo + code * cnt, raw)
        width, height = tags[256][0], tags[257][0]
        if tags.get(259, (1,))[0] != 1 or tags.get(258, (8,))[0] != 8 or tags.get(277, (1,))[0] != 1 \
                or 273 not in tags or 33550 not in tags or 33922 not in tags:
            raise RuntimeError("unsupported soil grid image format")
        pixels = b"".join(data[o:o + c] for o, c in zip(tags[273], tags[279]))
        if len(pixels) < width * height:
            raise RuntimeError("soil grid image is truncated")
        tie, scale = tags[33922], tags[33550]
        west, north = tie[3] - tie[0] * scale[0], tie[4] + tie[1] * scale[1]
        rows = [list(pixels[r * width:(r + 1) * width]) for r in range(height)]
        return rows, west, north, scale[0], scale[1]

    def _issr800_hsg(self, basin_geom):
        """Backup for _ssurgo_hsg: hydrologic soil group fractions from the UC Davis SoilWeb ISSR-800 grid
        (800 m cells aggregated from SSURGO), sampled inside the basin. Coarser than the SSURGO polygons."""
        coords = basin_geom.get("coordinates") or []
        polys = [coords] if basin_geom.get("type") == "Polygon" else coords
        ring = [(c[0], c[1]) for c in max((poly[0] for poly in polys if poly), key=len)]
        xs, ys = [p[0] for p in ring], [p[1] for p in ring]
        minx, maxx, miny, maxy = min(xs), max(xs), min(ys), max(ys)
        k = math.cos(math.radians(0.5 * (miny + maxy)))
        step = max((maxx - minx) * k, maxy - miny) / self.ISSR800_SAMPLES
        nx = max(int(round((maxx - minx) * k / step)), 1)
        ny = max(int(round((maxy - miny) / step)), 1)
        crs = "http://www.opengis.net/def/crs/EPSG/0/4326"
        data = self._http_get_bytes(self.ISSR800_WCS_URL, {
            "map": "issr800-wcs", "SERVICE": "WCS", "VERSION": "2.0.1", "REQUEST": "GetCoverage",
            "COVERAGEID": "hydgrp", "FORMAT": "image/tiff", "SUBSETTINGCRS": crs, "OUTPUTCRS": crs,
            "SUBSET": [f"x({minx:.6f},{maxx:.6f})", f"y({miny:.6f},{maxy:.6f})"],
            "SCALESIZE": f"x({nx}),y({ny})"})
        rows, west, north, dx, dy = self._read_byte_geotiff(data)

        def _inside(x, y):
            hit = False
            for (x1, y1), (x2, y2) in zip(ring, ring[1:]):
                if (y1 > y) != (y2 > y) and x < x1 + (y - y1) * (x2 - x1) / (y2 - y1):
                    hit = not hit
            return hit

        counts = {}
        for r, row in enumerate(rows):
            y = north - (r + 0.5) * dy
            for c, v in enumerate(row):
                grp = self.ISSR800_HSG_CODES.get(v)
                if grp and _inside(west + (c + 0.5) * dx, y):
                    counts[grp] = counts.get(grp, 0) + 1
        if not counts:   # sliver basin: take the cell at its centre
            cx, cy = 0.5 * (minx + maxx), 0.5 * (miny + maxy)
            r = min(max(int((north - cy) / dy), 0), len(rows) - 1)
            c = min(max(int((cx - west) / dx), 0), len(rows[0]) - 1)
            grp = self.ISSR800_HSG_CODES.get(rows[r][c])
            if grp:
                counts[grp] = 1
        total = sum(counts.values())
        if not total:
            raise RuntimeError("no soil group data for this basin")
        return {g: n / total for g, n in counts.items()}

    def basin_soil_groups(self, basin_geom):
        """Hydrologic soil group fractions for a basin: (fractions, source). USDA SSURGO polygons first; when
        that service is down, the UC Davis ISSR-800 grid. (None, '') when neither answers."""
        wait = self._down_minutes("ssurgo")
        if wait:
            self.log(f"  USDA SSURGO soils service is down (retrying in {wait:.0f} min).")
        else:
            try:
                return self._ssurgo_hsg(basin_geom), self.HSG_SRC_SSURGO
            except Exception as exc:
                self._mark_down("ssurgo", exc)
                self.log(f"  [WARNING] USDA SSURGO soils unavailable ({exc}); trying the UC Davis SoilWeb soil grid.")
        try:
            return self._issr800_hsg(basin_geom), self.HSG_SRC_ISSR800
        except Exception as exc:
            self.log(f"  [WARNING] UC Davis SoilWeb soil grid unavailable ({exc}); soil properties are left blank"
                     " and soil group B is assumed when the NRCS CN is estimated.")
            return None, ""

    def _ssurgo_hsg(self, basin_geom):
        """Hydrologic soil group area fractions {A, B, C, D} inside the basin, from USDA Soil Data Access.
        Dual groups (A/D, B/D, C/D) count as D, the undrained condition."""
        coords = basin_geom.get("coordinates") or []
        polys = [coords] if basin_geom.get("type") == "Polygon" else coords
        ring = max((poly[0] for poly in polys if poly), key=len)
        ring = _simplify_ring([(c[0], c[1]) for c in ring], 400)
        if ring[0] != ring[-1]:
            ring.append(ring[0])
        wkt = "polygon((" + ",".join(f"{x:.6f} {y:.6f}" for x, y in ring) + "))"
        query = ("~DeclareGeometry(@aoi)~\n"
                 f"select @aoi = geometry::STPolyFromText('{wkt}', 4326)\n"
                 "~DeclareIdGeomTable(@clip)~\n"
                 "~GetClippedMapunits(@aoi,polygon,geo,@clip)~\n"
                 "select c.hydgrp, sum(g.geom.STArea() * c.comppct_r) as w "
                 "from @clip g join component c on c.mukey = g.id "
                 "where c.hydgrp is not null group by c.hydgrp")
        data = self._http_post_json(self.SSURGO_URL, {"query": query, "format": "JSON"}, timeout=90, max_retries=2)
        out = {}
        for grp, weight in (data or {}).get("Table", []) or []:
            letter = str(grp).strip().upper().split("/")[-1]
            if letter in ("A", "B", "C", "D"):
                out[letter] = out.get(letter, 0.0) + float(weight)
        total = sum(out.values())
        if total <= 0:
            raise RuntimeError("no soil groups returned for this basin")
        return {k: v / total for k, v in out.items()}

    def _streamcat_cover(self, lon, lat):
        """(impervious %, forest %, COMID) for the NHDPlus watershed at the point - EPA StreamCat (NLCD 2019)."""
        fc = self._http_get_json(self.NLDI_COMID_URL, {"coords": f"POINT({lon:.6f} {lat:.6f})"},
                                 timeout=45, max_retries=2)
        props = fc["features"][0]["properties"]
        comid = str(props.get("comid") or props.get("identifier"))
        data = self._http_get_json(self.STREAMCAT_URL, {
            "name": "pctimp2019,pctdecid2019,pctconif2019,pctmxfst2019", "comid": comid,
            "areaOfInterest": "ws"}, timeout=45, max_retries=2)
        item = data["items"][0]
        forest = sum(float(item[k]) for k in ("pctdecid2019ws", "pctconif2019ws", "pctmxfst2019ws"))
        return float(item["pctimp2019ws"]), forest, comid

    def _channel_slope_10_85(self, lfp, outlet):
        """Channel slope (ft/mi) between 10% and 85% of the flowpath length above the outlet, from USGS 3DEP."""
        pts = [(float(p[0]), float(p[1])) for p in lfp]
        if self._ft_dist(pts[0], outlet) > self._ft_dist(pts[-1], outlet):
            pts.reverse()
        total = self._line_len_ft(pts)
        if total <= 0:
            raise RuntimeError("zero-length flowpath")

        def _at(frac):
            target, run = frac * total, 0.0
            for a, b in zip(pts, pts[1:]):
                seg = self._ft_dist(a, b)
                if seg > 0 and run + seg >= target:
                    t = (target - run) / seg
                    return [a[0] + t * (b[0] - a[0]), a[1] + t * (b[1] - a[1])]
                run += seg
            return list(pts[-1])

        data = self._http_get_json(self.DEM_SAMPLES_URL, {
            "geometry": json.dumps({"points": [_at(0.10), _at(0.85)], "spatialReference": {"wkid": 4326}}),
            "geometryType": "esriGeometryMultipoint", "returnFirstValueOnly": "true", "f": "json",
        }, timeout=90, max_retries=2)
        elev = {int(s["locationId"]): float(s["value"]) * 3.28084 for s in data.get("samples", [])}
        return abs(elev[1] - elev[0]) / (0.75 * total / 5280.0), total

    def _alternate_basin_characteristics(self, basin_geom, basin_rings, outlet, params_list, nlcd=None):
        """Fills impervious, forest and channel slope from other public services when StreamStats basin
        characteristics are unavailable. Returns the NHD-traced longest flowpath (reused later) or None."""
        self.log("  StreamStats basin characteristics are down; using alternate sources (MRLC NLCD 2021, USGS 3DEP).")

        def _add(code, name, value, unit):
            params_list.append({"code": code, "name": name, "value": round(value, 2), "unit": unit})

        if nlcd:
            _add("LC21IMP", "Impervious area, NLCD 2021 (MRLC, clipped to basin)", nlcd["imperv_pct"], "percent")
            _add("LC21FOREST", "Forest cover, NLCD 2021 classes 41-43 (MRLC, clipped to basin)", nlcd["forest_pct"],
                 "percent")
            _add("LC21DEV", "Developed (urban) area, NLCD 2021 classes 21-24 (MRLC, clipped to basin)",
                 nlcd["dev_pct"], "percent")
        else:
            try:
                imp, forest, comid = self._streamcat_cover(outlet[0], outlet[1])
                _add("LC19IMP", f"Impervious area, NLCD 2019 (EPA StreamCat, NHDPlus watershed {comid})", imp, "percent")
                _add("LC19FOREST", f"Forest cover, NLCD 2019 (EPA StreamCat, NHDPlus watershed {comid})", forest, "percent")
                self.log(f"  EPA StreamCat (NHDPlus watershed {comid}): impervious {imp:.1f}%, forest {forest:.1f}%.")
            except Exception as exc2:
                self.log(f"  [WARNING] EPA StreamCat lookup failed ({exc2}); using default impervious 2% and forest 50%.")

        lfp = None
        try:
            lfp = self._nhd_longest_flowpath(basin_rings, outlet)
            if not lfp or len(lfp) < 2:
                raise RuntimeError("no flowpath")
            slope, length = self._channel_slope_10_85(lfp, outlet)
            _add("CSL10_85", "Channel slope 10-85 (USGS 3DEP elevation along the NHD-traced longest flowpath)",
                 slope, "feet per mi")
            self.log(f"  Channel slope 10-85 (USGS 3DEP): {slope:,.1f} ft/mi along a {length:,.0f} ft flowpath.")
        except Exception as exc:
            self.log(f"  [WARNING] Channel slope from USGS 3DEP failed ({exc}); using the default 50 ft/mi.")
        return lfp

    @staticmethod
    def _ft_dist(a, b):
        dx = (b[0] - a[0]) * 364000.0 * math.cos(math.radians(0.5 * (a[1] + b[1])))
        dy = (b[1] - a[1]) * 364000.0
        return math.hypot(dx, dy)

    @classmethod
    def _line_len_ft(cls, pts):
        return sum(cls._ft_dist(pts[i], pts[i + 1]) for i in range(len(pts) - 1))

    def compute_longest_flowpath(self, ws_data, basin_rings, outlet, nhd_lfp=None):
        """USGS StreamStats longest flowpath where the state publishes it, else an NHD-traced approximation."""
        self.log("  Computing longest flowpath...")
        lfp = self._usgs_longest_flowpath(ws_data)
        if lfp:
            return lfp, "USGS StreamStats"
        lfp = nhd_lfp or self._nhd_longest_flowpath(basin_rings, outlet)
        if lfp:
            return lfp, ("NHD channel + overland (approx.)" if len(lfp) > 2 else "Straight line divide-to-outlet (approx.)")
        return None, ""

    LFP_OUTAGE_SKIP_SEC = 600
    _lfp_down_until = 0.0

    def _usgs_longest_flowpath(self, ws_data):
        now = time.time()
        if now < type(self)._lfp_down_until:
            mins = (type(self)._lfp_down_until - now) / 60.0
            self.log(f"  USGS longest-flowpath service is down (retrying in {mins:.0f} min); "
                     "using the NHD-traced flowpath.")
            return None
        try:
            data = self._http_post_json(self.LFP_FEATURES_URL, ws_data, timeout=120, max_retries=2)
        except Exception as exc:
            msg = str(exc)
            if any(f"HTTP Error {c}" in msg or f" {c} " in msg for c in (500, 502, 503, 504)):
                type(self)._lfp_down_until = time.time() + self.LFP_OUTAGE_SKIP_SEC
                self.log("  USGS longest-flowpath service is down (server error "
                         f"{next(c for c in (500, 502, 503, 504) if str(c) in msg)}); "
                         "using the NHD-traced flowpath. USGS will be retried in 10 min.")
            else:
                self.log(f"  USGS longest flowpath unavailable ({exc}); using the NHD-traced flowpath.")
            return None
        lines = []

        def _walk(obj, tagged=False):
            if isinstance(obj, dict):
                tag = tagged or "longestflowpath" in str(obj.get("name", obj.get("id", ""))).lower().replace("_", "")
                geom = obj.get("geometry")
                if tag and isinstance(geom, dict):
                    if geom.get("type") == "LineString":
                        lines.append([(c[0], c[1]) for c in geom.get("coordinates", [])])
                    elif geom.get("type") == "MultiLineString":
                        lines.extend([(c[0], c[1]) for c in part] for part in geom.get("coordinates", []))
                for v in obj.values():
                    if isinstance(v, (dict, list)):
                        _walk(v, tag)
            elif isinstance(obj, list):
                for v in obj:
                    _walk(v, tagged)

        _walk(data)
        lines = [ln for ln in lines if len(ln) >= 2]
        if not lines:
            self.log("  USGS does not publish a longest flowpath for this region; tracing NHD flowlines instead.")
            return None
        return max(lines, key=self._line_len_ft)

    _nhd_lfp_cache = None   # (outlet, ring size, path) of the last flowpath traced from NHD

    def _nhd_longest_flowpath(self, basin_rings, outlet):
        """Longest NHD channel path to the outlet, extended overland to the nearest upstream divide vertex."""
        if not basin_rings:
            return None
        ring = max(basin_rings, key=len)
        outlet = (float(outlet[0]), float(outlet[1]))
        cached = type(self)._nhd_lfp_cache
        if cached and cached[1] == len(ring) and self._ft_dist(cached[0], outlet) < 100.0:
            self.log("  Same outlet and basin as the previous culvert; reusing its NHD-traced flowpath.")
            return list(cached[2])
        clip_ring = _simplify_ring(list(ring), 600)
        lons = [p[0] for p in ring]
        lats = [p[1] for p in ring]
        params = {
            "geometry": f"{min(lons)},{min(lats)},{max(lons)},{max(lats)}",
            "geometryType": "esriGeometryEnvelope", "inSR": "4326",
            "spatialRel": "esriSpatialRelIntersects", "outFields": "OBJECTID",
            "returnGeometry": "true", "outSR": "4326", "f": "geojson",
        }
        self.log("  Tracing longest NHD flowpath inside the basin...")
        graph = {}
        features, got_nhd = [], False
        # The USGS NHD server often answers on a second or third try, so a failed query is repeated before
        # falling back to a straight line (which is only a rough stand-in for the real flowpath).
        for attempt in range(1, 4):
            t0 = time.time()
            try:
                gj = self._http_get_json(self.NHD_STANDARD_URL, params, timeout=120, max_retries=2)
                if "error" in gj:
                    raise RuntimeError("the service returned an error")
                if gj.get("exceededTransferLimit") or (gj.get("properties") or {}).get("exceededTransferLimit"):
                    self.log("  [WARNING] NHD record limit reached; flowpath may be incomplete.")
                features, got_nhd = gj.get("features", []), True
                break
            except Exception as exc:
                if attempt == 3 or time.time() - t0 > 60.0:
                    self.log(f"  [WARNING] NHD flowline query failed ({exc}); using a straight divide-to-outlet "
                             "line instead of the real flowpath.")
                    break
                self.log(f"  NHD flowline query failed ({exc}); trying again ({attempt + 1} of 3)...")
                time.sleep(5.0)

        def _key(p):
            return (round(p[0], 6), round(p[1], 6))

        for feat in features:
            geom = feat.get("geometry") or {}
            gtype = geom.get("type")
            if gtype == "LineString":
                parts = [geom.get("coordinates", [])]
            elif gtype == "MultiLineString":
                parts = geom.get("coordinates", [])
            else:
                parts = []
            for part in parts:
                pts = [(c[0], c[1]) for c in part]
                for seg in PurePythonClipper.clip_polyline_to_polygon(pts, clip_ring):
                    for a, b in zip(seg, seg[1:]):
                        ka, kb = _key(a), _key(b)
                        if ka == kb:
                            continue
                        d = self._ft_dist(ka, kb)
                        graph.setdefault(ka, {})[kb] = d
                        graph.setdefault(kb, {})[ka] = d

        if not graph:
            far = max(ring, key=lambda p: self._ft_dist(p, outlet))
            return [(float(far[0]), float(far[1])), outlet]

        start = min(graph, key=lambda k: self._ft_dist(k, outlet))
        dist, prev, pq = {start: 0.0}, {}, [(0.0, start)]
        while pq:
            d, u = heapq.heappop(pq)
            if d > dist.get(u, float("inf")):
                continue
            for v, w in graph[u].items():
                nd = d + w
                if nd < dist.get(v, float("inf")):
                    dist[v], prev[v] = nd, u
                    heapq.heappush(pq, (nd, v))
        head = max(dist, key=dist.get)
        path = [head]
        while path[-1] != start:
            path.append(prev[path[-1]])

        channel = dist[head]
        overland = self._overland_start(clip_ring, head, graph)
        if overland:
            path.insert(0, overland)
        if self._ft_dist(path[-1], outlet) > 1.0:
            path.append(outlet)
        path = [(float(p[0]), float(p[1])) for p in path]
        far = max(ring, key=lambda p: self._ft_dist(p, outlet))
        if self._ft_dist(far, outlet) > self._line_len_ft(path):
            self.log(f"  NHD channel ({channel:,.0f} ft) shorter than divide distance; using divide-to-outlet line.")
        type(self)._nhd_lfp_cache = (outlet, len(ring), list(path))
        return path

    @staticmethod
    def _overland_start(ring, head, graph):
        """Divide vertex where the overland leg of the longest flowpath starts: the farthest one that drains
        to the channel head, i.e. no other stream point is closer to it than the head, and the straight leg
        to the head crosses neither a stream nor the basin boundary. None when there is no such vertex."""
        k = 364000.0 * math.cos(math.radians(head[1]))

        def xy(p):
            return (p[0] * k, p[1] * 364000.0)

        def crosses(p1, p2, a, b):
            def side(o, q, r):
                return (q[0] - o[0]) * (r[1] - o[1]) - (q[1] - o[1]) * (r[0] - o[0])
            return side(p1, p2, a) * side(p1, p2, b) < 0 and side(a, b, p1) * side(a, b, p2) < 0

        h = xy(head)
        nodes = [xy(n) for n in graph if n != head]
        edges = [(xy(a), xy(b)) for a in graph for b in graph[a] if a < b]
        ring_xy = [xy(p) for p in ring]
        ring_edges = list(zip(ring_xy, ring_xy[1:]))
        candidates = []
        for p, (px, py) in zip(ring, ring_xy):
            d2 = (px - h[0]) ** 2 + (py - h[1]) ** 2
            if d2 < 1.0:
                continue
            limit = d2 - 1.0
            if any((px - nx) ** 2 + (py - ny) ** 2 < limit for nx, ny in nodes):
                continue
            candidates.append((d2, p, (px, py)))
        for _d2, p, pxy in sorted(candidates, reverse=True):
            if any(crosses(pxy, h, a, b) for a, b in edges):
                continue
            if any(crosses(pxy, h, a, b) for a, b in ring_edges if a != pxy and b != pxy):
                continue
            return p
        return None

    @staticmethod
    def _nss_code_to_pk(code):
        """Maps NSS AEP codes (e.g. PK1AEP, PK0_2AEP) to return-period tags (PK100, PK500)."""
        code = str(code or "").upper().strip()
        m = re.match(r"^PK(\d+(?:_\d+)?)AEP$", code)
        if not m:
            return code
        aep = float(m.group(1).replace("_", "."))
        return f"PK{int(round(100.0 / aep))}" if aep > 0 else code

    @classmethod
    def _parse_flow_stats(cls, fl_json):
        """Prefers the 'Area-Averaged' regression region when NSS returns several."""
        res_map = {}
        area_avg_map = {}
        items = fl_json if isinstance(fl_json, list) else [fl_json] if isinstance(fl_json, dict) else []
        for grp in items:
            for reg in grp.get("regressionRegions", grp.get("RegressionRegions", [])) or []:
                is_avg = "area-averaged" in str(reg.get("name", reg.get("Name", ""))).lower()
                for res in reg.get("results", reg.get("Results", [])) or []:
                    code = cls._nss_code_to_pk(res.get("code") or res.get("Name", ""))
                    val = res.get("value", res.get("Value"))
                    if val is None or not code:
                        continue
                    unit = res.get("unit") or res.get("Unit") or {}
                    entry = {"value": val, "unit": unit.get("abbreviation", unit.get("abbr", "cfs")) if isinstance(unit, dict) else "cfs",
                             "name": res.get("name", "")}
                    if is_avg:
                        area_avg_map[code] = entry
                    else:
                        res_map.setdefault(code, entry)
        res_map.update(area_avg_map)
        return res_map


# ---------------------------------------------------------------------------
# 7. Zero-Dependency KML (.kml) & ESRI Shapefile (.shp/.shx/.dbf/.prj) Exporter
# ---------------------------------------------------------------------------
class PurePythonShapefileWriter:
    """
    Pure-Python binary ESRI Shapefile (.shp, .shx, .dbf, .prj) writer using `struct`.
    Supports Point (1), PolyLine (3), and Polygon (5) in WGS84 (EPSG:4326) with zero
    external libraries required.
    """

    WGS84_PRJ = (
        'GEOGCS["GCS_WGS_1984",DATUM["D_WGS_1984",'
        'SPHEROID["WGS_1984",6378137.0,298.257223563]],'
        'PRIMEM["Greenwich",0.0],UNIT["Degree",0.0174532925199433]]'
    )

    @classmethod
    def write_shapefile(cls, base_filepath, shape_type, field_specs, records_with_parts):
        """
        Writes `<base>.shp`, `<base>.shx`, `<base>.dbf`, and `<base>.prj`.
        - shape_type: 1 = Point, 3 = PolyLine, 5 = Polygon
        - field_specs: list of (name_max10, type_'C'_or_'N', length, decimals)
        - records_with_parts: list of (parts_list, attr_dict)
            For Point (1): parts_list is [(lon, lat)]
            For PolyLine (3) / Polygon (5): parts_list is [ [(lon, lat), ...], ... ]
        """
        base = os.path.splitext(base_filepath)[0]
        if not records_with_parts:
            return None

        shp_records = []
        shx_records = []
        dbf_rows = []

        global_xmin, global_ymin = 1e30, 1e30
        global_xmax, global_ymax = -1e30, -1e30
        offset_words = 50  # 100-byte header = 50 16-bit words

        for rec_idx, (parts, attrs) in enumerate(records_with_parts, start=1):
            if shape_type == 1:
                lon, lat = parts[0]
                global_xmin = min(global_xmin, lon)
                global_ymin = min(global_ymin, lat)
                global_xmax = max(global_xmax, lon)
                global_ymax = max(global_ymax, lat)
                content = struct.pack("<idd", 1, float(lon), float(lat))
            else:
                valid_parts = []
                for ring in parts:
                    if len(ring) < 2:
                        continue
                    pts = [(float(p[0]), float(p[1])) for p in ring]
                    if shape_type == 5:
                        if pts[0] != pts[-1]:
                            pts.append(pts[0])
                        pts = cls._ensure_clockwise_ring(pts)
                    valid_parts.append(pts)

                if not valid_parts:
                    continue

                all_pts = [pt for part in valid_parts for pt in part]
                xs = [p[0] for p in all_pts]
                ys = [p[1] for p in all_pts]
                xmin, ymin, xmax, ymax = min(xs), min(ys), max(xs), max(ys)

                global_xmin = min(global_xmin, xmin)
                global_ymin = min(global_ymin, ymin)
                global_xmax = max(global_xmax, xmax)
                global_ymax = max(global_ymax, ymax)

                num_parts = len(valid_parts)
                num_points = len(all_pts)

                part_indices = []
                idx_cursor = 0
                for part in valid_parts:
                    part_indices.append(idx_cursor)
                    idx_cursor += len(part)

                header_buf = struct.pack(
                    "<iddddii",
                    shape_type,
                    xmin,
                    ymin,
                    xmax,
                    ymax,
                    num_parts,
                    num_points,
                )
                parts_buf = struct.pack(f"<{num_parts}i", *part_indices)
                pts_buf = b"".join(struct.pack("<dd", px, py) for px, py in all_pts)
                content = header_buf + parts_buf + pts_buf

            content_len_words = len(content) // 2
            rec_header = struct.pack(">ii", rec_idx, content_len_words)
            shp_records.append(rec_header + content)
            shx_records.append(struct.pack(">ii", offset_words, content_len_words))
            offset_words += 4 + content_len_words
            dbf_rows.append(attrs)

        if not shp_records:
            return None

        # 1. Write .shp
        shp_file_len_words = offset_words
        shp_header = cls._build_main_header(
            shp_file_len_words, shape_type, global_xmin, global_ymin, global_xmax, global_ymax
        )
        with open(base + ".shp", "wb") as f_shp:
            f_shp.write(shp_header)
            for rec in shp_records:
                f_shp.write(rec)

        # 2. Write .shx
        shx_file_len_words = 50 + 4 * len(shx_records)
        shx_header = cls._build_main_header(
            shx_file_len_words, shape_type, global_xmin, global_ymin, global_xmax, global_ymax
        )
        with open(base + ".shx", "wb") as f_shx:
            f_shx.write(shx_header)
            for rec in shx_records:
                f_shx.write(rec)

        # 3. Write .dbf
        cls._write_dbf(base + ".dbf", field_specs, dbf_rows)

        # 4. Write .prj and .cpg (UTF-8 attribute encoding hint for ArcGIS/QGIS)
        with open(base + ".prj", "w", encoding="utf-8") as f_prj:
            f_prj.write(cls.WGS84_PRJ)
        with open(base + ".cpg", "w", encoding="ascii") as f_cpg:
            f_cpg.write("UTF-8")

        return base + ".shp"

    @staticmethod
    def _ensure_clockwise_ring(pts):
        """ESRI Shapefile specification requires outer polygon rings to be clockwise."""
        area2 = 0.0
        for i in range(len(pts) - 1):
            x1, y1 = pts[i]
            x2, y2 = pts[i + 1]
            area2 += (x2 - x1) * (y2 + y1)
        # In standard Cartesian (lon, lat), area2 > 0 means clockwise
        if area2 < 0:
            return list(reversed(pts))
        return pts

    @staticmethod
    def _build_main_header(file_length_words, shape_type, xmin, ymin, xmax, ymax):
        part1 = struct.pack(">iiiiiii", 9994, 0, 0, 0, 0, 0, file_length_words)
        part2 = struct.pack(
            "<iidddddddd",
            1000,
            shape_type,
            xmin,
            ymin,
            xmax,
            ymax,
            0.0,
            0.0,
            0.0,
            0.0,
        )
        return part1 + part2

    @staticmethod
    def _write_dbf(dbf_path, field_specs, rows):
        now = datetime.datetime.now()
        num_records = len(rows)
        num_fields = len(field_specs)
        header_len = 32 + (32 * num_fields) + 1
        record_len = 1 + sum(f[2] for f in field_specs)

        with open(dbf_path, "wb") as f:
            # 32-byte dBASE III+ file header
            dbf_hdr = struct.pack(
                "<BBBBihH20s",
                0x03,
                now.year - 1900,
                now.month,
                now.day,
                num_records,
                header_len,
                record_len,
                b"\x00" * 20,
            )
            f.write(dbf_hdr)

            # Field descriptors (32 bytes each)
            for fname, ftype, flen, fdec in field_specs:
                name_bytes = fname[:10].encode("ascii", "ignore").ljust(11, b"\x00")
                f_desc = struct.pack(
                    "<11sc4xBB14s",
                    name_bytes,
                    ftype.encode("ascii"),
                    flen,
                    fdec,
                    b"\x00" * 14,
                )
                f.write(f_desc)

            # Header terminator
            f.write(b"\x0D")

            # Data records
            for row in rows:
                f.write(b" ")  # Deletion flag (space = active)
                for fname, ftype, flen, fdec in field_specs:
                    val = row.get(fname, "")
                    if ftype == "N":
                        try:
                            num_val = float(val)
                            if fdec > 0:
                                s_val = f"{num_val:{flen}.{fdec}f}"
                            else:
                                s_val = f"{int(round(num_val)):{flen}d}"
                        except (ValueError, TypeError):
                            s_val = " " * flen
                        raw = s_val[:flen].rjust(flen, " ").encode("ascii", "ignore")
                    else:
                        s_val = str(val if val is not None else "")
                        raw = s_val[:flen].ljust(flen, " ").encode("ascii", "ignore")
                    f.write(raw)

            f.write(b"\x1A")  # EOF marker


class GISExporter:
    """
    Exports Watershed Boundary Polygons, Snapped Pour Points, and Clipped NHDPlus
    Streams to Google Earth KML (.kml) and ESRI Shapefiles (.shp, .shx, .dbf, .prj).
    """

    BASIN_FIELD_SPECS = [
        ("CULVERT_ID", "C", 18, 0),
        ("STATE",      "C", 4,  0),
        ("WORKSPACE",  "C", 24, 0),
        ("DA_SQMI",    "N", 12, 4),
        ("DA_ACRES",   "N", 12, 2),
        ("SLOPE_FTMI", "N", 10, 2),
        ("VDOT_C",     "N", 8,  2),
        ("TC_MIN",     "N", 8,  1),
        ("Q2_USGS",    "N", 10, 1),
        ("Q10_USGS",   "N", 10, 1),
        ("Q25_USGS",   "N", 10, 1),
        ("Q50_USGS",   "N", 10, 1),
        ("Q100_USGS",  "N", 10, 1),
        ("Q10_VDOT",   "N", 10, 1),
        ("Q25_VDOT",   "N", 10, 1),
        ("Q100_VDOT",  "N", 10, 1),
        ("GOV_METHOD", "C", 32, 0),
    ]

    LFP_FIELD_SPECS = [
        ("CULVERT_ID", "C", 18, 0),
        ("LEN_FT",     "N", 12, 1),
        ("LEN_MI",     "N", 10, 3),
        ("METHOD",     "C", 40, 0),
    ]

    RAINDROP_FIELD_SPECS = [
        ("TRACE_ID",   "C", 18, 0),
        ("SEQ",        "N", 4,  0),
        ("SEGMENT",    "C", 10, 0),
        ("COMID",      "C", 14, 0),
        ("LEN_FT",     "N", 12, 1),
        ("CUM_MI",     "N", 10, 3),
        ("METHOD",     "C", 50, 0),
    ]

    @classmethod
    def _raindrop_rows(cls, res):
        """[(seq, segment, comid, pts, len_ft, cum_ft, method)] for a raindrop trace result."""
        rows = [(1, "OVERLAND", "", res["overland_wgs84"], res["overland_ft"], res["overland_ft"],
                 res.get("overland_method", ""))]
        for i, r in enumerate(res.get("reaches") or [], start=2):
            rows.append((i, "STREAM", str(r.get("comid", "")), r["pts"], r["length_ft"], r["cum_ft"],
                         res.get("source", "")))
        return rows

    @classmethod
    def export_raindrop(cls, base_filepath, trace_id, res, do_kml=True, do_shp=True, kmz=False, style=None):
        """Writes `<base>.kml|.kmz` and `<base>_RaindropPath.shp` for a Flow (Raindrop) Path trace."""
        base = os.path.splitext(base_filepath)[0]
        rows = cls._raindrop_rows(res)
        created = []
        if do_shp:
            recs = [([pts], {"TRACE_ID": str(trace_id), "SEQ": seq, "SEGMENT": seg, "COMID": comid,
                             "LEN_FT": float(lf), "CUM_MI": float(cf) / 5280.0, "METHOD": str(m)[:50]})
                    for seq, seg, comid, pts, lf, cf, m in rows if len(pts) >= 2]
            shp = PurePythonShapefileWriter.write_shapefile(base + "_RaindropPath.shp", 3,
                                                            cls.RAINDROP_FIELD_SPECS, recs)
            if shp:
                created.append(shp)
        if do_kml:
            st = dict(cls.load_kml_style())
            st.update(style or {})
            kc = cls._kml_color
            tid = xml_escape(str(trace_id))
            col, wid = kc(st["raindrop_color"]), float(st["raindrop_width"])
            out = [
                '<?xml version="1.0" encoding="UTF-8"?>',
                '<kml xmlns="http://www.opengis.net/kml/2.2">', "  <Document>",
                f"    <name>{tid} Flow (Raindrop) Path</name>",
                f'    <Style id="rdStream"><LineStyle><color>{col}</color><width>{wid}</width></LineStyle></Style>',
                f'    <Style id="rdOverland"><LineStyle><color>{kc(st["raindrop_color"], 70)}</color>'
                f"<width>{max(1.0, wid - 1.0)}</width></LineStyle></Style>",
                '    <Style id="rdPoint">'
                f"<IconStyle><color>{col}</color><scale>{float(st['point_scale'])}</scale>"
                f"<Icon><href>{xml_escape(str(st['point_icon']))}</href></Icon></IconStyle>"
                f"<LabelStyle><scale>{float(st['label_scale'])}</scale></LabelStyle></Style>",
            ]
            for seq, seg, comid, pts, lf, cf, m in rows:
                if len(pts) < 2:
                    continue
                name = "Overland flow" if seg == "OVERLAND" else f"Reach {seq - 1} (COMID {xml_escape(comid)})"
                c_str = " ".join(f"{p[0]:.6f},{p[1]:.6f},0" for p in pts)
                out += [
                    "    <Placemark>", f"      <name>{name}</name>",
                    f"      <description><![CDATA[Length: {lf:,.0f} ft<br/>Cumulative: {cf / 5280.0:.3f} mi<br/>"
                    f"Method: {xml_escape(str(m))}]]></description>",
                    f"      <styleUrl>#{'rdOverland' if seg == 'OVERLAND' else 'rdStream'}</styleUrl>",
                    f"      <LineString><tessellate>1</tessellate><coordinates>{c_str}</coordinates></LineString>",
                    "    </Placemark>",
                ]
            path = res["path_wgs84"]
            for label, (lon, lat) in ((f"{tid} Start", res["query_point"]),
                                      (f"{tid} End ({res['total_ft'] / 5280.0:.2f} mi)", path[-1])):
                out += ["    <Placemark>", f"      <name>{label}</name>", "      <styleUrl>#rdPoint</styleUrl>",
                        f"      <Point><coordinates>{lon:.6f},{lat:.6f},0</coordinates></Point>", "    </Placemark>"]
            out += ["  </Document>", "</kml>"]
            text = "\n".join(out)
            kpath = base + (".kmz" if kmz else ".kml")
            if kmz:
                with zipfile.ZipFile(kpath, "w", zipfile.ZIP_DEFLATED) as zf:
                    zf.writestr("doc.kml", text)
            else:
                with open(kpath, "w", encoding="utf-8") as fh:
                    fh.write(text)
            created.insert(0, kpath)
        return created

    DEFAULT_KML_STYLE_FILE = os.path.join(os.path.expanduser("~"), ".usgs_streamstats_kml_style.json")
    KML_STYLE_DEFAULTS = {
        "basin_line_color": "#FF0000", "basin_line_width": 2.5,
        "basin_fill": True, "basin_fill_color": "#FF0000", "basin_fill_opacity": 33,
        "stream_color": "#0088FF", "stream_width": 3.0,
        "flowpath_color": "#FF00FF", "flowpath_width": 3.5,
        "raindrop_color": "#00E5FF", "raindrop_width": 4.0,
        "point_color": "#00FF00", "point_scale": 1.1,
        "point_icon": "http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png",
        "label_scale": 0.8,
    }
    KML_ICONS = {
        "Circle": "http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png",
        "Square": "http://maps.google.com/mapfiles/kml/shapes/placemark_square.png",
        "Triangle": "http://maps.google.com/mapfiles/kml/shapes/triangle.png",
        "Target": "http://maps.google.com/mapfiles/kml/shapes/target.png",
        "Pushpin": "http://maps.google.com/mapfiles/kml/pushpin/wht-pushpin.png",
        "Water": "http://maps.google.com/mapfiles/kml/shapes/water.png",
    }

    @classmethod
    def kml_style_path(cls):
        """Style file location chosen in the KML / KMZ Style dialog (remembered in the UI prefs)."""
        return _load_ui_prefs().get("kml_style_file") or cls.DEFAULT_KML_STYLE_FILE

    @classmethod
    def set_kml_style_path(cls, path):
        _save_ui_prefs(kml_style_file=None if not path or os.path.normcase(os.path.abspath(path)) ==
                       os.path.normcase(cls.DEFAULT_KML_STYLE_FILE) else os.path.abspath(path))

    @classmethod
    def load_kml_style(cls, path=None):
        style = dict(cls.KML_STYLE_DEFAULTS)
        try:
            with open(path or cls.kml_style_path(), encoding="utf-8") as fh:
                saved = json.load(fh)
            style.update({k: v for k, v in saved.items() if k in style})
        except Exception:
            pass
        return style

    @classmethod
    def save_kml_style(cls, style, path=None):
        path = path or cls.kml_style_path()
        folder = os.path.dirname(path)
        if folder:
            os.makedirs(folder, exist_ok=True)
        with open(path, "w", encoding="utf-8") as fh:
            json.dump(style, fh, indent=2)

    @staticmethod
    def _kml_color(hex_rgb, opacity_pct=100):
        """'#RRGGBB' + opacity % -> KML 'aabbggrr'."""
        h = str(hex_rgb or "#FFFFFF").lstrip("#")
        if len(h) != 6:
            h = "FFFFFF"
        alpha = int(round(max(0.0, min(float(opacity_pct), 100.0)) * 2.55))
        return f"{alpha:02x}{h[4:6]}{h[2:4]}{h[0:2]}".lower()

    STREAM_FIELD_SPECS = [
        ("REACH_ID",   "N", 6,  0),
        ("GNIS_NAME",  "C", 40, 0),
        ("STRM_ORDER", "N", 6,  0),
        ("DA_SQKM",    "N", 12, 3),
        ("LEN_FT",     "N", 12, 1),
    ]

    @classmethod
    def _build_basin_attr_dict(cls, culvert_id, ss_result):
        vcomp = ss_result.get("vdot_comparison", {})
        met = vcomp.get("metrics", {})
        r_map = {r["return_yr"]: r for r in vcomp.get("rows", [])}

        def _num(val):
            return float(val) if isinstance(val, (int, float)) else 0.0

        return {
            "CULVERT_ID": str(culvert_id or "CULV-01"),
            "STATE":      str(ss_result.get("rcode", "VA")),
            "WORKSPACE":  str(ss_result.get("workspace_id", "")),
            "DA_SQMI":    _num(met.get("drnarea_sqmi", 0.0)),
            "DA_ACRES":   _num(met.get("drnarea_acres", 0.0)),
            "SLOPE_FTMI": _num(met.get("slope_ft_mi", 0.0)),
            "VDOT_C":     _num(vcomp.get("c_coeff", 0.0)),
            "TC_MIN":     _num(vcomp.get("tc_min", 0.0)),
            "Q2_USGS":    _num(r_map.get(2, {}).get("usgs_q_cfs", 0.0)),
            "Q10_USGS":   _num(r_map.get(10, {}).get("usgs_q_cfs", 0.0)),
            "Q25_USGS":   _num(r_map.get(25, {}).get("usgs_q_cfs", 0.0)),
            "Q50_USGS":   _num(r_map.get(50, {}).get("usgs_q_cfs", 0.0)),
            "Q100_USGS":  _num(r_map.get(100, {}).get("usgs_q_cfs", 0.0)),
            "Q10_VDOT":   _num(r_map.get(10, {}).get("vdot_q_cfs", 0.0)),
            "Q25_VDOT":   _num(r_map.get(25, {}).get("vdot_q_cfs", 0.0)),
            "Q100_VDOT":  _num(r_map.get(100, {}).get("vdot_q_cfs", 0.0)),
            "GOV_METHOD": str(vcomp.get("governing_method", "")),
        }

    @classmethod
    def export_kml(cls, kml_filepath, culvert_items, clipped_streams=None, style=None):
        """
        Writes a styled Google Earth KML (.kml) or zipped KMZ (.kmz) containing Watershed Polygons,
        Pour Points, Clipped NHDPlus Flowlines and Longest Flowpaths. Symbology comes from
        `style` or the saved KML style (see the "KML / KMZ Style" dialog).
        """
        st = dict(cls.load_kml_style())
        st.update(style or {})
        kc = cls._kml_color
        lines = [
            '<?xml version="1.0" encoding="UTF-8"?>',
            '<kml xmlns="http://www.opengis.net/kml/2.2">',
            "  <Document>",
            "    <name>USGS StreamStats &amp; VDOT Hydrology Deliverable</name>",
            '    <Style id="basinPolyStyle">',
            f"      <LineStyle><color>{kc(st['basin_line_color'])}</color><width>{float(st['basin_line_width'])}</width></LineStyle>",
            f"      <PolyStyle><color>{kc(st['basin_fill_color'], st['basin_fill_opacity'])}</color>"
            f"<fill>{1 if st['basin_fill'] else 0}</fill><outline>1</outline></PolyStyle>",
            "    </Style>",
            '    <Style id="pourPointStyle">',
            f"      <IconStyle><color>{kc(st['point_color'])}</color><scale>{float(st['point_scale'])}</scale>"
            f"<Icon><href>{xml_escape(str(st['point_icon']))}</href></Icon></IconStyle>",
            f"      <LabelStyle><scale>{float(st['label_scale'])}</scale></LabelStyle>",
            "    </Style>",
            '    <Style id="nhdStreamStyle">',
            f"      <LineStyle><color>{kc(st['stream_color'])}</color><width>{float(st['stream_width'])}</width></LineStyle>",
            "    </Style>",
            '    <Style id="flowpathStyle">',
            f"      <LineStyle><color>{kc(st['flowpath_color'])}</color><width>{float(st['flowpath_width'])}</width></LineStyle>",
            "    </Style>",
        ]

        # 1. Folder: Watershed Basins
        lines.append("    <Folder>")
        lines.append("      <name>1_Watershed_Basins</name>")
        for item in culvert_items or []:
            cid = xml_escape(str(item.get("culvert_id", "CULV-01")))
            ssr = item.get("ss_result")
            if not ssr or not ssr.get("basin_rings_wgs84"):
                continue
            attrs = cls._build_basin_attr_dict(cid, ssr)
            desc_html = (
                f"<![CDATA[<h3>{cid} - Watershed Summary</h3>"
                f"<table border='1' cellpadding='4' cellspacing='0'>"
                f"<tr><td><b>State / Workspace</b></td><td>{attrs['STATE']} / {attrs['WORKSPACE']}</td></tr>"
                f"<tr><td><b>Drainage Area</b></td><td>{attrs['DA_SQMI']:.3f} sq.mi ({attrs['DA_ACRES']:.1f} ac)</td></tr>"
                f"<tr><td><b>Channel Slope</b></td><td>{attrs['SLOPE_FTMI']:.1f} ft/mi</td></tr>"
                f"<tr><td><b>VDOT C / Tc</b></td><td>C = {attrs['VDOT_C']:.2f}, Tc = {attrs['TC_MIN']:.1f} min</td></tr>"
                f"<tr><td><b>Q10 (USGS / VDOT)</b></td><td>{attrs['Q10_USGS']:.1f} cfs / {attrs['Q10_VDOT']:.1f} cfs</td></tr>"
                f"<tr><td><b>Q25 (USGS / VDOT)</b></td><td>{attrs['Q25_USGS']:.1f} cfs / {attrs['Q25_VDOT']:.1f} cfs</td></tr>"
                f"<tr><td><b>Q100 (USGS / VDOT)</b></td><td>{attrs['Q100_USGS']:.1f} cfs / {attrs['Q100_VDOT']:.1f} cfs</td></tr>"
                f"<tr><td><b>Governing Method</b></td><td>{xml_escape(attrs['GOV_METHOD'])}</td></tr>"
                f"</table>]]>"
            )

            lines.append("      <Placemark>")
            lines.append(f"        <name>{cid} Basin ({attrs['DA_ACRES']:.1f} ac)</name>")
            lines.append(f"        <description>{desc_html}</description>")
            lines.append("        <styleUrl>#basinPolyStyle</styleUrl>")
            lines.append("        <MultiGeometry>")

            for ring in ssr["basin_rings_wgs84"]:
                if len(ring) < 3:
                    continue
                closed_ring = list(ring)
                if closed_ring[0] != closed_ring[-1]:
                    closed_ring.append(closed_ring[0])
                coord_str = " ".join(f"{pt[0]:.6f},{pt[1]:.6f},0" for pt in closed_ring)
                lines.append("          <Polygon>")
                lines.append("            <outerBoundaryIs><LinearRing>")
                lines.append(f"              <coordinates>{coord_str}</coordinates>")
                lines.append("            </LinearRing></outerBoundaryIs>")
                lines.append("          </Polygon>")

            lines.append("        </MultiGeometry>")
            lines.append("      </Placemark>")
        lines.append("    </Folder>")

        # 2. Folder: Pour Points / Culverts
        lines.append("    <Folder>")
        lines.append("      <name>2_Pour_Points_Culverts</name>")
        for item in culvert_items or []:
            cid = xml_escape(str(item.get("culvert_id", "CULV-01")))
            ssr = item.get("ss_result")
            if not ssr:
                continue
            plon, plat = ssr.get("snapped_point", ssr.get("query_point", (0.0, 0.0)))
            attrs = cls._build_basin_attr_dict(cid, ssr)
            lines.append("      <Placemark>")
            lines.append(f"        <name>{cid} Pour Point</name>")
            lines.append(
                f"        <description><![CDATA[<b>{cid}</b><br/>"
                f"Lon: {plon:.6f}, Lat: {plat:.6f}<br/>"
                f"Q25 VDOT: {attrs['Q25_VDOT']:.1f} cfs | Q100 VDOT: {attrs['Q100_VDOT']:.1f} cfs]]></description>"
            )
            lines.append("        <styleUrl>#pourPointStyle</styleUrl>")
            lines.append(
                f"        <Point><coordinates>{plon:.6f},{plat:.6f},0</coordinates></Point>"
            )
            lines.append("      </Placemark>")
        lines.append("    </Folder>")

        # 3. Folder: Clipped NHDPlus Streams
        if clipped_streams:
            lines.append("    <Folder>")
            lines.append("      <name>3_Clipped_NHDPlus_Streams</name>")
            for s in clipped_streams:
                sname = xml_escape(str(s.get("name", "Stream")))
                lines.append("      <Placemark>")
                lines.append(f"        <name>{sname} (Order {s.get('stream_order', 1)})</name>")
                lines.append(
                    f"        <description><![CDATA[Clipped Length: {s.get('length_ft', 0.0):.1f} ft<br/>"
                    f"Drainage Area: {s.get('tot_da_sqkm', 0.0):.2f} sq.km]]></description>"
                )
                lines.append("        <styleUrl>#nhdStreamStyle</styleUrl>")
                lines.append("        <MultiGeometry>")
                for part in s.get("clipped_lines_wgs84", []):
                    if len(part) < 2:
                        continue
                    c_str = " ".join(f"{pt[0]:.6f},{pt[1]:.6f},0" for pt in part)
                    lines.append(
                        f"          <LineString><coordinates>{c_str}</coordinates></LineString>"
                    )
                lines.append("        </MultiGeometry>")
                lines.append("      </Placemark>")
            lines.append("    </Folder>")

        # 4. Folder: Longest Flowpaths
        lfp_items = [i for i in culvert_items or [] if (i.get("ss_result") or {}).get("longest_flowpath_wgs84")]
        if lfp_items:
            lines.append("    <Folder>")
            lines.append("      <name>4_Longest_Flowpaths</name>")
            for item in lfp_items:
                cid = xml_escape(str(item.get("culvert_id", "CULV-01")))
                ssr = item["ss_result"]
                len_ft = float(ssr.get("lfp_length_ft", 0.0))
                c_str = " ".join(f"{pt[0]:.6f},{pt[1]:.6f},0" for pt in ssr["longest_flowpath_wgs84"])
                lines.append("      <Placemark>")
                lines.append(f"        <name>{cid} Longest Flowpath ({len_ft:,.0f} ft)</name>")
                lines.append(
                    f"        <description><![CDATA[Length: {len_ft:,.1f} ft ({len_ft / 5280.0:.3f} mi)<br/>"
                    f"Method: {xml_escape(str(ssr.get('lfp_method', '')))}]]></description>"
                )
                lines.append("        <styleUrl>#flowpathStyle</styleUrl>")
                lines.append(f"        <LineString><tessellate>1</tessellate><coordinates>{c_str}</coordinates></LineString>")
                lines.append("      </Placemark>")
            lines.append("    </Folder>")

        lines.append("  </Document>")
        lines.append("</kml>")

        text = "\n".join(lines)
        if kml_filepath.lower().endswith(".kmz"):
            with zipfile.ZipFile(kml_filepath, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.writestr("doc.kml", text)
        else:
            with open(kml_filepath, "w", encoding="utf-8") as f:
                f.write(text)
        return kml_filepath

    @classmethod
    def export_shapefiles(cls, base_filepath, culvert_items, clipped_streams=None):
        """
        Exports up to 3 ESRI Shapefile sets (.shp, .shx, .dbf, .prj):
        1. `<base>_Basins.shp` (Polygon)
        2. `<base>_PourPoints.shp` (Point)
        3. `<base>_ClippedStreams.shp` (PolyLine)
        4. `<base>_LongestFlowpath.shp` (PolyLine, only when computed)
        Returns a list of created `.shp` file paths.
        """
        base = os.path.splitext(base_filepath)[0]
        created_files = []

        # 1. Watershed Basins Polygon Shapefile & Pour Points Shapefile
        basin_records = []
        point_records = []
        for item in culvert_items or []:
            cid = item.get("culvert_id", "CULV-01")
            ssr = item.get("ss_result")
            if not ssr:
                continue
            attrs = cls._build_basin_attr_dict(cid, ssr)
            rings = ssr.get("basin_rings_wgs84", [])
            if rings:
                basin_records.append((rings, attrs))
            plon, plat = ssr.get("snapped_point", ssr.get("query_point", (0.0, 0.0)))
            point_records.append(([(plon, plat)], attrs))

        if basin_records:
            basin_shp = PurePythonShapefileWriter.write_shapefile(
                base + "_Basins.shp", 5, cls.BASIN_FIELD_SPECS, basin_records
            )
            if basin_shp:
                created_files.append(basin_shp)

        if point_records:
            pt_shp = PurePythonShapefileWriter.write_shapefile(
                base + "_PourPoints.shp", 1, cls.BASIN_FIELD_SPECS, point_records
            )
            if pt_shp:
                created_files.append(pt_shp)

        # 2. Clipped NHDPlus Streams PolyLine Shapefile
        stream_records = []
        for s in clipped_streams or []:
            parts = s.get("clipped_lines_wgs84", [])
            if not parts:
                continue
            s_attrs = {
                "REACH_ID":   int(s.get("id", 1)),
                "GNIS_NAME":  str(s.get("name", "Unnamed")),
                "STRM_ORDER": int(s.get("stream_order", 1)),
                "DA_SQKM":    float(s.get("tot_da_sqkm", 0.0)),
                "LEN_FT":     float(s.get("length_ft", 0.0)),
            }
            stream_records.append((parts, s_attrs))

        if stream_records:
            strm_shp = PurePythonShapefileWriter.write_shapefile(
                base + "_ClippedStreams.shp", 3, cls.STREAM_FIELD_SPECS, stream_records
            )
            if strm_shp:
                created_files.append(strm_shp)

        lfp_records = []
        for item in culvert_items or []:
            ssr = item.get("ss_result") or {}
            path = ssr.get("longest_flowpath_wgs84")
            if not path or len(path) < 2:
                continue
            len_ft = float(ssr.get("lfp_length_ft", 0.0))
            lfp_records.append(([path], {
                "CULVERT_ID": str(item.get("culvert_id", "CULV-01")),
                "LEN_FT": len_ft, "LEN_MI": len_ft / 5280.0,
                "METHOD": str(ssr.get("lfp_method", "")),
            }))
        if lfp_records:
            lfp_shp = PurePythonShapefileWriter.write_shapefile(
                base + "_LongestFlowpath.shp", 3, cls.LFP_FIELD_SPECS, lfp_records
            )
            if lfp_shp:
                created_files.append(lfp_shp)

        return created_files

    @classmethod
    def auto_save_gis_in_dgn_folder(cls, prefix_tag, culvert_items, clipped_streams=None):
        """
        Automatically writes `<DGN_Stem>_<prefix_tag>.kml` and Shapefiles into the
        active DGN file's directory.
        """
        dgn_dir, dgn_stem = MicroStationDgnManager.get_active_dgn_directory_and_stem()
        safe_tag = re.sub(r'[<>:"/\\|?*\s]+', "_", str(prefix_tag)).strip("_") or "Export"
        base_path = os.path.join(dgn_dir, f"{dgn_stem}_{safe_tag}")
        kml_path = cls.export_kml(base_path + ".kml", culvert_items, clipped_streams)
        shp_paths = cls.export_shapefiles(base_path, culvert_items, clipped_streams)
        return kml_path, shp_paths


# ---------------------------------------------------------------------------
# 8. MicroStation CAD Manager (Cell Placement, EC Item Types, Tables)
# ---------------------------------------------------------------------------
class MicroStationDgnManager:
    LEVEL_CONFIG = {
        "USGS_NHD_Streams_Clipped": {"color": 1, "weight": 2, "style": 0},
        "USGS_NHD_Streams_Labels":  {"color": 1, "weight": 1, "style": 0},
        "USGS_NHD_Flow_Arrows":     {"color": 1, "weight": 1, "style": 0},
        "USGS_StreamStats_Basin":   {"color": 3, "weight": 2, "style": 2},
        "USGS_StreamStats_Point":   {"color": 2, "weight": 3, "style": 0},
        "USGS_Culvert_Inlet":       {"color": 6, "weight": 2, "style": 0},
        "USGS_Longest_Flowpath":    {"color": 5, "weight": 3, "style": 0},
        "USGS_Raindrop_Path":       {"color": 7, "weight": 3, "style": 0},
        "USGS_StreamStats_Table":   {"color": 0, "weight": 1, "style": 0},
        "USGS_Fence_Boundary":      {"color": 3, "weight": 2, "style": 0},
    }
    # Item Types and native TextTables have crashed MicroStation 2026 when invoked from the
    # Tk event loop; keep them off unless explicitly re-enabled.
    SAFE_DGN_MODE = True
    # The drainage-area Item Type (IDs, basin characteristics and hydrology results) is attached even in safe mode.
    ATTACH_ITEM_TYPE = True
    MAX_SHAPE_VERTICES = 4500
    ITEM_LIBRARY_NAME = "USGS_VDOT_Hydrology"
    ITEM_TYPE_NAME = "DrainageArea"
    RETURN_PERIODS = (2, 5, 10, 25, 50, 100, 200, 500)
    # (property name, "s" string | "d" double). Missing properties are added to existing libraries.
    BASIN_ITEM_PROPS = [
        ("CulvertID", "s"), ("StateCode", "s"), ("WorkspaceID", "s"),
        ("DrainageArea_SqMi", "d"), ("DrainageArea_Acres", "d"),
        ("Slope_FtPerMi", "d"), ("Impervious_Pct", "d"), ("Forest_Pct", "d"),
        ("RunoffCoeff_C", "d"), ("Tc_Min", "d"),
        ("LongestFlowpath_Ft", "d"), ("LongestFlowpath_Method", "s"),
        ("GoverningMethod", "s"),
        ("SoilGroup_A_Pct", "d"), ("SoilGroup_B_Pct", "d"), ("SoilGroup_C_Pct", "d"), ("SoilGroup_D_Pct", "d"),
        ("DominantSoilGroup", "s"), ("SoilSource", "s"),
        ("CurveNumber_CN", "d"), ("CurveNumber_Source", "s"),
    ] + [(f"Q{yr}_{src}_cfs", "d") for yr in RETURN_PERIODS for src in ("USGS", "VDOT")] + [
        ("CalcDate", "s"),
    ]
    STREAM_ITEM_TYPE_NAME = "Tributary"
    STREAM_ITEM_PROPS = [
        ("StreamName", "s"), ("NHDPlusID", "s"), ("ReachCode", "s"), ("FeatureType", "s"),
        ("StreamOrder", "d"), ("FlowDirection", "s"), ("FlowBearing_Deg", "d"), ("FlowsInto", "s"),
        ("MeanAnnualFlow_cfs", "d"), ("MeanVelocity_fps", "d"), ("DrainageArea_SqMi", "d"),
        ("Slope_FtPerFt", "d"), ("ClippedLength_Ft", "d"), ("Source", "s"), ("QueryDate", "s"),
    ]
    FLOWPATH_ITEM_TYPE_NAME = "LongestFlowpath"
    FLOWPATH_ITEM_PROPS = [
        ("CulvertID", "s"), ("Length_Ft", "d"), ("Length_Mi", "d"), ("Method", "s"),
        ("FlowDirection", "s"), ("FlowBearing_Deg", "d"), ("Segment", "s"),
        ("BasinSlope_FtPerMi", "d"), ("Tc_Min", "d"), ("TcMethod", "s"), ("DrainageArea_Acres", "d"),
        ("Upstream_Lat", "d"), ("Upstream_Lon", "d"), ("Outlet_Lat", "d"), ("Outlet_Lon", "d"),
        ("CalcDate", "s"),
    ]
    RAINDROP_ITEM_TYPE_NAME = "RaindropPath"
    RAINDROP_ITEM_PROPS = [
        ("TraceID", "s"), ("Segment", "s"), ("Method", "s"), ("FlowDirection", "s"), ("FlowBearing_Deg", "d"),
        ("SegmentLength_Ft", "d"), ("OverlandLength_Ft", "d"), ("TotalLength_Ft", "d"), ("TotalLength_Mi", "d"),
        ("MaxDownstream_Mi", "d"), ("ReachCount", "d"), ("StartReachCOMID", "s"), ("EndReachCOMID", "s"),
        ("Start_Lat", "d"), ("Start_Lon", "d"), ("End_Lat", "d"), ("End_Lon", "d"), ("CalcDate", "s"),
    ]

    def __init__(self, projector: DgnCoordinateProjector, log_callback=None):
        self.projector = projector
        self.log = log_callback if log_callback else print
        self.ref_extent_uor = 0.0
        self._rec_ids = None

    @staticmethod
    def get_active_dgn_directory_and_stem():
        if IN_MICROSTATION:
            try:
                dgn_file = ISessionMgr.ActiveDgnFile
                if dgn_file:
                    fpath = str(dgn_file.GetFileName())
                    if fpath and os.path.dirname(fpath):
                        return os.path.dirname(fpath), os.path.splitext(os.path.basename(fpath))[0]
            except Exception:
                pass
        return os.path.expanduser("~"), "ActiveDGN"

    def get_active_fence_polygon_uor(self):
        if not IN_MICROSTATION:
            cx, cy = 11055000.0 * 10000.0, 3625000.0 * 10000.0
            dx, dy = 1200.0 * 10000.0, 900.0 * 10000.0
            return [(cx-dx, cy-dy), (cx+dx, cy-dy), (cx+dx, cy+dy), (cx-dx, cy+dy), (cx-dx, cy-dy)]

        fence_mgr = FenceManager.GetManager()
        if not fence_mgr.IsFenceActive():
            raise RuntimeError("No active Fence in MicroStation. Click \u25ad Place Fence in DGN (or place a Fence) first.")

        pts = []
        try:
            dpts = DPoint3dArray()
            fence_mgr.GetFencePoints(dpts)
            pts = [(float(p.x), float(p.y)) for p in dpts]
        except Exception:
            pts = []

        if len(pts) < 3:
            # Fallback: read the clip boundary through FenceParams.
            try:
                fparams = FenceParams(ISessionMgr.ActiveDgnModelRef)
                fence_mgr.InitFromActiveFence(fparams, False, False, FenceClipMode.eNone)
                cvec = fparams.GetClipVector()
                for prim in cvec or []:
                    poly = prim.GetPolygon() if hasattr(prim, "GetPolygon") else None
                    for p in poly or []:
                        pts.append((float(p.x), float(p.y)))
            except Exception:
                pass

        if len(pts) < 3:
            raise RuntimeError("Unable to read fence vertices. Re-place the fence as a Block or Shape and try again.")

        if pts[0] != pts[-1]:
            pts.append(pts[0])
        return pts

    def ensure_levels_exist(self):
        if not IN_MICROSTATION:
            return {}
        _trace("ensure_levels_exist")
        dgn_file = ISessionMgr.ActiveDgnFile
        cache = dgn_file.GetLevelCache()
        ids = {}
        for name, p in self.LEVEL_CONFIG.items():
            try:
                h = cache.GetLevelByName(name)
                if not h.IsValid():
                    eh = cache.CreateLevel(name, LEVEL_NULL_CODE, LEVEL_NULL_ID)
                    if eh.IsValid():
                        try:
                            eh.SetByLevelColor(LevelDefinitionColor(p["color"], dgn_file))
                            eh.SetByLevelWeight(p["weight"])
                        except Exception:
                            pass
                        h = eh
                if h.IsValid():
                    ids[name] = h.GetLevelId()
            except Exception as exc:
                self.log(f"  Level note ({name}): {exc}")
        cache.Write()
        return ids

    def _apply_symbology(self, eeh, level_id, color, weight, style=0):
        if not IN_MICROSTATION:
            return
        ps = ElementPropertiesSetter()
        if level_id:
            ps.SetLevel(level_id)
        ps.SetColor(color)
        ps.SetWeight(weight)
        ps.SetLinestyle(style, None)
        ps.Apply(eeh)

    def draw_clipped_streams(self, clipped_streams):
        if not IN_MICROSTATION:
            self.log(f"[Simulation] Would draw {len(clipped_streams)} clipped streams & labels in DGN.")
            return

        _trace(f"draw_clipped_streams n={len(clipped_streams)}")
        model_ref = ISessionMgr.ActiveDgnModelRef
        lvl_ids = self.ensure_levels_exist()
        stream_lvl = lvl_ids.get("USGS_NHD_Streams_Clipped")
        label_lvl = lvl_ids.get("USGS_NHD_Streams_Labels")
        arrow_lvl = lvl_ids.get("USGS_NHD_Flow_Arrows")
        arrow_cfg = self.LEVEL_CONFIG["USGS_NHD_Flow_Arrows"]
        cfg = self.LEVEL_CONFIG["USGS_NHD_Streams_Clipped"]
        lbl_cfg = self.LEVEL_CONFIG["USGS_NHD_Streams_Labels"]

        count, lbl_count, arrow_count, item_count = 0, 0, 0, 0
        stream_itype = None
        if self.ATTACH_ITEM_TYPE:
            try:
                stream_itype = self._ensure_item_type(self.STREAM_ITEM_TYPE_NAME, self.STREAM_ITEM_PROPS)
            except Exception:
                _trace("stream item type error:\n" + traceback.format_exc())
        for stream in clipped_streams:
            lbl_text = f"{stream['name']} (Order {stream['stream_order']})"
            for line_wgs84 in stream["clipped_lines_wgs84"]:
                dpts = DPoint3dArray()
                pts_uor = []
                for lon, lat in line_wgs84:
                    xu, yu, zu = self.projector.wgs84_to_uor(lon, lat)
                    dpts.append(DPoint3d(xu, yu, zu))
                    pts_uor.append((xu, yu))
                if len(dpts) >= 2:
                    eeh = EditElementHandle()
                    if LineStringHandler.CreateLineStringElement(eeh, None, dpts, model_ref.Is3d(), model_ref) == BentleyStatus.eSUCCESS:
                        self._apply_symbology(eeh, stream_lvl, cfg["color"], cfg["weight"], cfg["style"])
                        eeh.AddToModel()
                        count += 1
                        if stream_itype and self.attach_stream_item_type(eeh, stream, pts_uor, stream_itype):
                            item_count += 1
                    try:
                        if self._draw_stream_label_along_segment(dpts, lbl_text, model_ref, label_lvl, lbl_cfg["color"]):
                            lbl_count += 1
                    except Exception as exc:
                        self.log(f"  Stream label skipped ({exc})")
                    try:
                        arrow_count += self.draw_flow_arrows(dpts, model_ref, arrow_lvl, arrow_cfg["color"])
                    except Exception as exc:
                        self.log(f"  Flow arrows skipped ({exc})")

        self.log(f"  Added {count} clipped stream reaches, {lbl_count} rotated stream labels "
                 f"and {arrow_count} flow-direction arrows to DGN.")
        if item_count:
            self.log(f"  Attached Item Type '{self.ITEM_LIBRARY_NAME}:{self.STREAM_ITEM_TYPE_NAME}' "
                     f"to {item_count} tributary line(s).")

    FLOW_ARROW_LEN_TEXT = 1.5      # arrowhead length, in text heights
    FLOW_ARROW_SPACING_TEXT = 40.0  # distance between arrows along a stream, in text heights

    def draw_flow_arrows(self, dpts, model_ref, level_id, color):
        """Filled arrowheads along a stream pointing downstream (NHD flowlines are digitized in flow direction).
        Placed away from the midpoint, where the stream label sits. Returns the number of arrows drawn."""
        n_pts = len(dpts)
        if n_pts < 2:
            return 0
        seg = [math.hypot(dpts[i + 1].x - dpts[i].x, dpts[i + 1].y - dpts[i].y) for i in range(n_pts - 1)]
        total = sum(seg)
        alen = self.FLOW_ARROW_LEN_TEXT * self.text_height_uor()
        if total < 4.0 * alen:
            return 0
        n = max(1, int(total / (self.FLOW_ARROW_SPACING_TEXT * self.text_height_uor())))
        stations = [total * (i + 0.25) / n for i in range(n)]
        if n == 1:
            stations.append(total * 0.75)
        drawn = 0
        for st in stations:
            acc = 0.0
            for i, L in enumerate(seg):
                if acc + L >= st and L > 1e-6:
                    f = (st - acc) / L
                    ux, uy = (dpts[i + 1].x - dpts[i].x) / L, (dpts[i + 1].y - dpts[i].y) / L
                    tx, ty = dpts[i].x + f * (dpts[i + 1].x - dpts[i].x), dpts[i].y + f * (dpts[i + 1].y - dpts[i].y)
                    z = dpts[i].z
                    bx, by = tx - ux * alen, ty - uy * alen
                    hw = 0.4 * alen
                    tri = [(tx, ty, z), (bx - uy * hw, by + ux * hw, z), (bx + uy * hw, by - ux * hw, z), (tx, ty, z)]
                    if self._add_filled_shape(tri, model_ref, level_id, color):
                        drawn += 1
                    break
                acc += L
        return drawn

    def _add_filled_shape(self, pts_uor, model_ref, level_id, color):
        dpa = DPoint3dArray()
        for x, y, z in pts_uor:
            dpa.append(DPoint3d(x, y, z))
        eeh = EditElementHandle()
        if ShapeHandler.CreateShapeElement(eeh, None, dpa, model_ref.Is3d(), model_ref) != BentleyStatus.eSUCCESS:
            return False
        self._apply_symbology(eeh, level_id, color, 1, 0)
        try:
            eeh.GetHandler().AddSolidFill(eeh, color, True)
        except Exception:
            pass
        self._add_to_model(eeh)
        return True

    def _draw_stream_label_along_segment(self, dpts, label_text, model_ref, level_id, color):
        if len(dpts) < 2 or not label_text:
            return False
        seg_lengths = []
        total_len = 0.0
        for i in range(len(dpts) - 1):
            L = math.hypot(dpts[i + 1].x - dpts[i].x, dpts[i + 1].y - dpts[i].y)
            seg_lengths.append(L)
            total_len += L
        if total_len < 25.0 * self.projector.uor_per_master:
            return False

        half_len = 0.5 * total_len
        accum = 0.0
        mid_x, mid_y, angle_rad = dpts[0].x, dpts[0].y, 0.0
        for i, L in enumerate(seg_lengths):
            if accum + L >= half_len and L > 1e-6:
                frac = (half_len - accum) / L
                mid_x = dpts[i].x + frac * (dpts[i + 1].x - dpts[i].x)
                mid_y = dpts[i].y + frac * (dpts[i + 1].y - dpts[i].y)
                angle_rad = math.atan2(dpts[i + 1].y - dpts[i].y, dpts[i + 1].x - dpts[i].x)
                break
            accum += L

        if angle_rad > math.pi / 2.0:
            angle_rad -= math.pi
        elif angle_rad < -math.pi / 2.0:
            angle_rad += math.pi

        th = self.text_height_uor()
        offset_uor = 1.0 * th
        perp_x = -math.sin(angle_rad) * offset_uor
        perp_y = math.cos(angle_rad) * offset_uor

        dgn_model = model_ref.GetDgnModel()
        tprops = TextBlockProperties.Create(dgn_model)
        pprops = ParagraphProperties.Create(dgn_model)
        rprops = RunProperties.Create(DgnFontManager.GetDefaultTrueTypeFont(), DPoint2d(th, th), dgn_model)

        tb = TextBlock(tprops, pprops, rprops, dgn_model)
        tb.SetUserOrigin(DPoint3d(mid_x + perp_x, mid_y + perp_y, dpts[0].z))
        tb.SetOrientation(RotMatrix.FromAxisAndRotationAngle(2, angle_rad))
        tb.AppendText(label_text)

        teh = EditElementHandle()
        if _status_ok(TextHandlerBase.CreateElement(teh, None, tb)):
            self._apply_symbology(teh, level_id, color, 1, 0)
            teh.AddToModel()
            return True
        return False

    SYMBOL_UNITS_PER_TEXT = 7.0   # symbol size unit = Settings text height / this

    def _sym_uor(self):
        """UORs per symbol unit, from the Settings drawing scale (not the view zoom), so markers plot at a
        consistent size next to their labels."""
        return self.text_height_uor() / self.SYMBOL_UNITS_PER_TEXT

    def _add_to_model(self, eeh):
        """AddToModel, recording the new element ID while a marker group is being recorded."""
        eeh.AddToModel()
        if self._rec_ids is not None:
            try:
                self._rec_ids.append(int(eeh.GetElementId()))
            except Exception:
                pass

    def basin_extent_uor(self, ss_result):
        """Largest side (UOR) of the drainage-area bounding box, 0 when there is no basin geometry."""
        rings = (ss_result or {}).get("basin_rings_wgs84") or []
        if not rings or not rings[0]:
            return 0.0
        lons = [p[0] for p in rings[0]]
        lats = [p[1] for p in rings[0]]
        try:
            x1, y1, _ = self.projector.wgs84_to_uor(min(lons), min(lats))
            x2, y2, _ = self.projector.wgs84_to_uor(max(lons), max(lats))
        except Exception:
            return 0.0
        return max(abs(x2 - x1), abs(y2 - y1))

    ANNOT_UNITS_PER_TEXT = 4.0

    def annot_unit(self, extent_uor=0.0):
        """Annotation unit (UOR) for culvert markers and gaps: Settings text height / ANNOT_UNITS_PER_TEXT,
        so markers follow the drawing scale. `extent_uor` is accepted for older callers and ignored."""
        return self.text_height_uor() / self.ANNOT_UNITS_PER_TEXT

    DRAWING_SCALES = (50.0, 100.0, 200.0)   # 1" = N'
    drawing_scale = 100.0
    plot_text_in = 0.20

    def text_height_ft(self):
        """DGN text height (ft) = plotted height (in) x drawing scale (ft per in)."""
        return self.plot_text_in * self.drawing_scale

    def text_height_uor(self):
        return self.text_height_ft() * self.projector.uor_per_master

    def label_text_height(self, unit=None):
        """Marker label text height (UOR), from the Settings drawing scale."""
        return self.text_height_uor()

    _fence_outline_ids = None

    def draw_fence_outline(self, uor_poly):
        """Draws the fence boundary as a red outline on level 'USGS_Fence_Boundary' (MicroStation has no setting
        for its own fence color). Replaces the outline drawn earlier in this session. Returns True when drawn."""
        if not IN_MICROSTATION or len(uor_poly or []) < 3:
            return False
        pts = [(float(p[0]), float(p[1]), 0.0) for p in uor_poly]
        if pts[0] != pts[-1]:
            pts.append(pts[0])
        cfg = self.LEVEL_CONFIG["USGS_Fence_Boundary"]
        lvl = self.ensure_levels_exist().get("USGS_Fence_Boundary")
        if self._fence_outline_ids:
            self.delete_elements(self._fence_outline_ids)
        prev_rec, self._rec_ids = self._rec_ids, []
        try:
            ok = self._add_linestring(pts, ISessionMgr.ActiveDgnModelRef, lvl, cfg["color"], cfg["weight"],
                                      cfg["style"])
        finally:
            self._fence_outline_ids, self._rec_ids = list(self._rec_ids), prev_rec
        return ok

    def delete_elements(self, ids):
        """Deletes elements (by ID) from the active model; returns the number removed."""
        if not IN_MICROSTATION or not ids:
            return 0
        dgn_model = ISessionMgr.ActiveDgnModelRef.GetDgnModel()
        n = 0
        for eid in ids:
            try:
                ref = dgn_model.FindElementByID(eid)
                if ref is None:
                    continue
                eeh = EditElementHandle(ref, dgn_model)
                if eeh.IsValid() and eeh.DeleteFromModel() == BentleyStatus.eSUCCESS:
                    n += 1
            except Exception:
                pass
        return n

    def find_tool_elements(self):
        """IDs of every element in the active model on one of this tool's levels (LEVEL_CONFIG)."""
        if not IN_MICROSTATION:
            return []
        cache = ISessionMgr.ActiveDgnFile.GetLevelCache()
        lvl_ids = set()
        for name in self.LEVEL_CONFIG:
            try:
                h = cache.GetLevelByName(name)
                if h.IsValid():
                    lvl_ids.add(h.GetLevelId())
            except Exception:
                pass
        if not lvl_ids:
            return []
        found = []
        for ref in ISessionMgr.ActiveDgnModelRef.GetDgnModel().GetGraphicElements():
            try:
                if ElementPropertiesGetter(ElementHandle(ref)).GetLevel() in lvl_ids:
                    found.append(ref.GetElementId())
            except Exception:
                pass
        return found

    def _existing_ranges(self, level_name, elem_type=None):
        """(low.x, low.y, high.x, high.y) in UORs of each element already on `level_name` (optionally
        only `elem_type`). Used to avoid drawing the same drainage area, flowpath or marker twice."""
        out = []
        if not IN_MICROSTATION:
            return out
        try:
            h = ISessionMgr.ActiveDgnFile.GetLevelCache().GetLevelByName(level_name)
            if not h.IsValid():
                return out
            lvl = h.GetLevelId()
        except Exception:
            return out
        for ref in ISessionMgr.ActiveDgnModelRef.GetDgnModel().GetGraphicElements():
            try:
                eh = ElementHandle(ref)
                if ElementPropertiesGetter(eh).GetLevel() != lvl:
                    continue
                if elem_type is not None and eh.GetElementType() != elem_type:
                    continue
                cv = ICurvePathQuery.ElementToCurveVector(eh)
                if cv is None:
                    continue
                r = DRange3d()
                cv.GetRange(r)
                out.append((r.low.x, r.low.y, r.high.x, r.high.y))
            except Exception:
                pass
        return out

    def _already_drawn(self, level_name, pts_uor, elem_type=None):
        """True when an element on `level_name` already covers the same extent as `pts_uor` (within 0.5 master unit)."""
        if not pts_uor:
            return False
        xs = [p[0] for p in pts_uor]
        ys = [p[1] for p in pts_uor]
        box = (min(xs), min(ys), max(xs), max(ys))
        tol = 0.5 * self.projector.uor_per_master
        return any(all(abs(a - b) <= tol for a, b in zip(box, r))
                   for r in self._existing_ranges(level_name, elem_type))

    def _marker_at(self, level_name, x, y, elem_type):
        """True when a marker element of `elem_type` on `level_name` is already centred on (x, y)."""
        tol = 0.5 * self.projector.uor_per_master
        return any(abs((r[0] + r[2]) / 2 - x) <= tol and abs((r[1] + r[3]) / 2 - y) <= tol
                   for r in self._existing_ranges(level_name, elem_type))

    def delete_all_tool_elements(self):
        """Deletes everything this tool drew (all elements on its levels) and clears the active fence."""
        n = self.delete_elements(self.find_tool_elements())
        self._fence_outline_ids = None
        self._batch_table_ids = None
        if IN_MICROSTATION:
            try:
                FenceManager.GetManager().ClearFence()
            except Exception:
                pass
        return n

    def _add_linestring(self, pts_uor, model_ref, level_id, color, weight, style=0, closed=False, on_added=None):
        dpts = DPoint3dArray()
        for x, y, z in pts_uor:
            dpts.append(DPoint3d(x, y, z))
        eeh = EditElementHandle()
        if closed:
            ok = ShapeHandler.CreateShapeElement(eeh, None, dpts, model_ref.Is3d(), model_ref)
        else:
            ok = LineStringHandler.CreateLineStringElement(eeh, None, dpts, model_ref.Is3d(), model_ref)
        if ok == BentleyStatus.eSUCCESS:
            self._apply_symbology(eeh, level_id, color, weight, style)
            self._add_to_model(eeh)
            if on_added:
                on_added(eeh)
            return True
        return False

    def place_pour_point_marker(self, origin_uor, model_ref, level_id, color, label="", unit=None):
        """Pour point (snapped outlet): circle with crosshairs. `unit` = annotation unit (UOR)."""
        if not IN_MICROSTATION:
            self.log(f"  [Simulation] Placed pour-point marker for {label}")
            return
        _trace("place_pour_point_marker")
        u = unit or self.annot_unit()
        rad = 2.5 * u
        ox, oy, oz = origin_uor.x, origin_uor.y, origin_uor.z
        ell_eeh = EditElementHandle()
        if EllipseHandler.CreateEllipseElement(
            ell_eeh, None, origin_uor, rad, rad, 0.0, model_ref.Is3d(), model_ref
        ) == BentleyStatus.eSUCCESS:
            self._apply_symbology(ell_eeh, level_id, color, 3, 0)
            self._add_to_model(ell_eeh)
        self._add_linestring([(ox - rad * 1.5, oy, oz), (ox + rad * 1.5, oy, oz)], model_ref, level_id, color, 1)
        self._add_linestring([(ox, oy - rad * 1.5, oz), (ox, oy + rad * 1.5, oz)], model_ref, level_id, color, 1)
        if label:
            _trace("place_text_callout")
            self._place_text_callout(origin_uor, label, model_ref, level_id, color,
                                     th=self.label_text_height(u), offset=(rad * 1.6, 0.3 * self.label_text_height(u)))

    def place_culvert_inlet_marker(self, origin_uor, model_ref, level_id, color, label="", unit=None):
        """Culvert inlet (user/query location): square with an X. `unit` = annotation unit (UOR)."""
        if not IN_MICROSTATION:
            self.log(f"  [Simulation] Placed culvert-inlet marker for {label}")
            return
        _trace("place_culvert_inlet_marker")
        u = unit or self.annot_unit()
        h = 1.6 * u
        ox, oy, oz = origin_uor.x, origin_uor.y, origin_uor.z
        sq = [(ox - h, oy - h, oz), (ox + h, oy - h, oz), (ox + h, oy + h, oz), (ox - h, oy + h, oz), (ox - h, oy - h, oz)]
        self._add_linestring(sq, model_ref, level_id, color, 2, closed=True)
        self._add_linestring([sq[0], sq[2]], model_ref, level_id, color, 1)
        self._add_linestring([sq[1], sq[3]], model_ref, level_id, color, 1)
        if label:
            self._place_text_callout(DPoint3d(ox, oy, oz), label, model_ref, level_id, color,
                                     th=self.label_text_height(u),
                                     offset=(4.0 * u, -1.5 * self.label_text_height(u)))

    def draw_markers_at_uor(self, x, y, z, culvert_id="", unit=None):
        """Places pour-point and culvert-inlet markers at a picked DGN point. Returns the new element IDs."""
        if not IN_MICROSTATION:
            self.log(f"  [Simulation] Placed pour-point & culvert-inlet markers for {culvert_id}.")
            return []
        model_ref = ISessionMgr.ActiveDgnModelRef
        lvl_ids = self.ensure_levels_exist()
        p_cfg = self.LEVEL_CONFIG["USGS_StreamStats_Point"]
        i_cfg = self.LEVEL_CONFIG["USGS_Culvert_Inlet"]
        xm, ym, _ = self.projector.uor_to_master(x, y, z)
        inlet_lbl = f"{culvert_id} INLET  E {xm:,.1f}  N {ym:,.1f}" if culvert_id else ""
        self._rec_ids = []
        try:
            self.place_culvert_inlet_marker(DPoint3d(x, y, z), model_ref, lvl_ids.get("USGS_Culvert_Inlet"),
                                            i_cfg["color"], label=inlet_lbl, unit=unit)
            self.place_pour_point_marker(DPoint3d(x, y, z), model_ref, lvl_ids.get("USGS_StreamStats_Point"),
                                         p_cfg["color"], label=f"{culvert_id} PP" if culvert_id else "", unit=unit)
            return list(self._rec_ids)
        finally:
            self._rec_ids = None

    def draw_culvert_markers(self, ss_result, culvert_id=""):
        """Draws the pour-point marker (snapped outlet) and culvert-inlet marker (query point), sized to
        this culvert's drainage area. Returns 1 on success."""
        if not ss_result:
            return 0
        qlon, qlat = ss_result.get("query_point") or ss_result.get("snapped_point")
        plon, plat = ss_result.get("snapped_point") or (qlon, qlat)
        if not IN_MICROSTATION:
            self.log(f"  [Simulation] Pour-point & culvert-inlet markers for {culvert_id}.")
            return 1
        model_ref = ISessionMgr.ActiveDgnModelRef
        lvl_ids = self.ensure_levels_exist()
        p_cfg = self.LEVEL_CONFIG["USGS_StreamStats_Point"]
        i_cfg = self.LEVEL_CONFIG["USGS_Culvert_Inlet"]
        u = self.annot_unit(self.basin_extent_uor(ss_result))
        px, py, pz = self.projector.wgs84_to_uor(plon, plat)
        ix, iy, iz = self.projector.wgs84_to_uor(qlon, qlat)
        pp_exists = self._marker_at("USGS_StreamStats_Point", px, py, 15)
        in_exists = self._marker_at("USGS_Culvert_Inlet", ix, iy, 6)
        if pp_exists and in_exists:
            self.log(f"  Markers for {culvert_id or 'culvert'} are already in the DGN; not drawn again.")
            return 1
        if not pp_exists:
            self.place_pour_point_marker(DPoint3d(px, py, pz), model_ref, lvl_ids.get("USGS_StreamStats_Point"),
                                         p_cfg["color"], label=f"{culvert_id} PP" if culvert_id else "", unit=u)
        if not in_exists:
            self.place_culvert_inlet_marker(DPoint3d(ix, iy, iz), model_ref, lvl_ids.get("USGS_Culvert_Inlet"),
                                            i_cfg["color"], label=f"{culvert_id} INLET" if culvert_id else "", unit=u)
        self.log(f"  Drew pour-point & culvert-inlet markers for {culvert_id or 'culvert'}.")
        return 1

    def draw_longest_flowpath(self, ss_result, culvert_id=""):
        path = (ss_result or {}).get("longest_flowpath_wgs84")
        if not path or len(path) < 2:
            return False
        if not IN_MICROSTATION:
            self.log(f"  [Simulation] Would draw longest flowpath ({len(path)} vertices) for {culvert_id}.")
            return True
        _trace(f"draw_longest_flowpath pts={len(path)}")
        model_ref = ISessionMgr.ActiveDgnModelRef
        lvl_ids = self.ensure_levels_exist()
        cfg = self.LEVEL_CONFIG["USGS_Longest_Flowpath"]
        pts = [self.projector.wgs84_to_uor(lon, lat) for lon, lat in path]
        if self._already_drawn("USGS_Longest_Flowpath", pts[:4001]):
            self.log(f"  Longest flowpath for {culvert_id or 'culvert'} is already in the DGN; not drawn again.")
            return True
        step = 4000
        starts = list(range(0, len(pts) - 1, step))
        n_items = 0
        for seg_no, i in enumerate(starts, 1):
            seg_pts = pts[i:i + step + 1]
            attached = []
            on_added = None
            if self.ATTACH_ITEM_TYPE:
                def on_added(eeh, seg_no=seg_no):
                    attached.append(self.attach_flowpath_item_type(eeh, ss_result, culvert_id, pts,
                                                                   seg_no, len(starts)))
            self._add_linestring(seg_pts, model_ref, lvl_ids.get("USGS_Longest_Flowpath"),
                                 cfg["color"], cfg["weight"], cfg["style"], on_added=on_added)
            n_items += sum(1 for a in attached if a)
        self.log(f"  Drew longest flowpath ({ss_result.get('lfp_length_ft', 0.0):,.0f} ft) on level 'USGS_Longest_Flowpath'.")
        if n_items:
            self.log(f"  Attached Item Type '{self.ITEM_LIBRARY_NAME}:{self.FLOWPATH_ITEM_TYPE_NAME}' "
                     f"to the longest flowpath ({culvert_id or 'CULVERT-01'}).")
        return True

    def draw_labeled_point(self, x, y, z, label, level_name="USGS_Raindrop_Path"):
        """Draws a point marker (circle with an X) with a text label at a picked DGN location."""
        if not IN_MICROSTATION:
            self.log(f"  [Simulation] Drew point '{label}'.")
            return True
        _trace("draw_labeled_point")
        model_ref = ISessionMgr.ActiveDgnModelRef
        lvl = self.ensure_levels_exist().get(level_name)
        color = self.LEVEL_CONFIG[level_name]["color"]
        rad = 4.0 * self._sym_uor()
        self._draw_circle_x(x, y, z, rad, model_ref, lvl, color)
        self._marker_label(x, y, z, rad, label, model_ref, lvl, color)
        return True

    def _marker_label(self, x, y, z, rad, text_str, model_ref, lvl, color):
        """Label beside a circle marker at the Settings text height (drawing scale x plotted height)."""
        th = self.text_height_uor()
        self._place_text_callout(DPoint3d(x, y, z), text_str, model_ref, lvl, color,
                                 th=th, offset=(rad + 0.6 * th, -0.5 * th))

    def _draw_circle_x(self, x, y, z, rad, model_ref, lvl, color):
        """Circle of radius `rad` with an inscribed X (raindrop start symbol)."""
        eeh = EditElementHandle()
        if EllipseHandler.CreateEllipseElement(eeh, None, DPoint3d(x, y, z), rad, rad, 0.0,
                                               model_ref.Is3d(), model_ref) == BentleyStatus.eSUCCESS:
            self._apply_symbology(eeh, lvl, color, 3, 0)
            eeh.AddToModel()
        d = rad * math.sqrt(0.5)
        self._add_linestring([(x - d, y - d, z), (x + d, y + d, z)], model_ref, lvl, color, 2)
        self._add_linestring([(x - d, y + d, z), (x + d, y - d, z)], model_ref, lvl, color, 2)

    def draw_raindrop_path(self, res, trace_id="RD-01", draw_start=True):
        """Draws a Flow (Raindrop) Path: dashed overland leg, solid downstream reaches, start/end markers."""
        path = (res or {}).get("path_wgs84") or []
        if len(path) < 2:
            return False
        if not IN_MICROSTATION:
            self.log(f"  [Simulation] Would draw raindrop path ({len(path)} vertices) for {trace_id}.")
            return True
        _trace(f"draw_raindrop_path pts={len(path)}")
        model_ref = ISessionMgr.ActiveDgnModelRef
        lvl = self.ensure_levels_exist().get("USGS_Raindrop_Path")
        cfg = self.LEVEL_CONFIG["USGS_Raindrop_Path"]
        to_uor = self.projector.wgs84_to_uor

        overland = [to_uor(lon, lat) for lon, lat in res.get("overland_wgs84") or []]
        stream_pts = []
        for r in res.get("reaches") or []:
            pts = [to_uor(lon, lat) for lon, lat in r["pts"]]
            stream_pts.extend(pts[1:] if stream_pts and pts and pts[0] == stream_pts[-1] else pts)
        step = 4000
        chunks = [stream_pts[i:i + step + 1] for i in range(0, max(len(stream_pts) - 1, 0), step)]
        n_items = 0

        def _attach(kind, seg_pts, seg_no=1, n_segs=1):
            if not self.ATTACH_ITEM_TYPE:
                return None

            def cb(eeh):
                nonlocal n_items
                if self.attach_raindrop_item_type(eeh, res, trace_id, kind, seg_pts, seg_no, n_segs):
                    n_items += 1
            return cb

        if len(overland) >= 2:
            self._add_linestring(overland, model_ref, lvl, self.RAINDROP_OVERLAND_COLOR, 2, 2,
                                 on_added=_attach("Overland leg", overland))
        for k, chunk in enumerate(chunks, 1):
            self._add_linestring(chunk, model_ref, lvl, cfg["color"], cfg["weight"], cfg["style"],
                                 on_added=_attach("Downstream channel", chunk, k, len(chunks)))

        n_arrows = 0
        for line, acolor in [(overland, self.RAINDROP_OVERLAND_COLOR)] + [(c, cfg["color"]) for c in chunks]:
            if len(line) < 2:
                continue
            dpa = DPoint3dArray()
            for x, y, z in line:
                dpa.append(DPoint3d(x, y, z))
            try:
                n_arrows += self.draw_flow_arrows(dpa, model_ref, lvl, acolor)
            except Exception as exc:
                self.log(f"  Flow arrows skipped ({exc})")

        upm = self._sym_uor()
        sx, sy, sz = to_uor(*res["query_point"])
        if draw_start:
            self._draw_circle_x(sx, sy, sz, 4.0 * upm, model_ref, lvl, cfg["color"])
            self._marker_label(sx, sy, sz, 4.0 * upm, f"{trace_id} RAINDROP START", model_ref, lvl, cfg["color"])

        ex, ey, ez = to_uor(*path[-1])
        self._draw_circle_x(ex, ey, ez, 4.0 * upm, model_ref, lvl, cfg["color"])
        self._marker_label(ex, ey, ez, 4.0 * upm, f"{trace_id} END {res['total_ft'] / 5280.0:.2f} MI",
                           model_ref, lvl, cfg["color"])
        self.log(f"  Drew raindrop path ({res['total_ft']:,.0f} ft) on level 'USGS_Raindrop_Path' "
                 f"with {n_arrows} flow-direction arrow(s).")
        if n_items:
            self.log(f"  Attached Item Type '{self.ITEM_LIBRARY_NAME}:{self.RAINDROP_ITEM_TYPE_NAME}' "
                     f"to {n_items} raindrop path line(s).")
        try:
            xs = [p[0] for p in overland + stream_pts] + [sx, ex]
            ys = [p[1] for p in overland + stream_pts] + [sy, ey]
            th = self.text_height_uor()
            self.place_raindrop_table(res, trace_id, max(xs) + 4.0 * th, max(ys), model_ref, 4.0 * upm)
        except Exception as exc:
            self.log(f"  Raindrop table skipped ({exc})")
        return True

    RAINDROP_OVERLAND_COLOR = 4   # yellow (dashed); downstream channel uses the level color (cyan)

    def place_raindrop_table(self, res, trace_id, tx, ty, model_ref, marker_rad=None):
        """Bordered DGN table for a raindrop trace. The LEGEND column draws each line type / symbol
        in its real color and style, and each row's text uses the same color."""
        lvl_tbl = self.ensure_levels_exist().get("USGS_StreamStats_Table")
        cfg = self.LEVEL_CONFIG["USGS_Raindrop_Path"]
        ch_col, ov_col = cfg["color"], self.RAINDROP_OVERLAND_COLOR
        th = self.text_height_uor()
        f = lambda v, d=0: f"{v:,.{d}f}"
        rows = [("LEGEND", "SEGMENT", "NHDPLUS COMID", "LENGTH (FT)", "CUMULATIVE (MI)", "METHOD")]
        symbols, colors = {}, {}

        def _line_sym(color, weight, style):
            return lambda x0, x1, yc, h: self._add_linestring([(x0, yc, 0.0), (x1, yc, 0.0)], model_ref, lvl_tbl,
                                                               color, weight, style)

        def _arrow_sym(color):
            def fn(x0, x1, yc, h):
                self._add_linestring([(x0, yc, 0.0), (x1, yc, 0.0)], model_ref, lvl_tbl, color, 1, 0)
                a, xm = 0.35 * h, (x0 + x1) / 2.0
                self._add_filled_shape([(xm + a, yc, 0.0), (xm - a, yc + 0.6 * a, 0.0),
                                        (xm - a, yc - 0.6 * a, 0.0), (xm + a, yc, 0.0)], model_ref, lvl_tbl, color)
            return fn

        def _circle_x_sym(color):
            return lambda x0, x1, yc, h: self._draw_circle_x((x0 + x1) / 2.0, yc, 0.0, 0.32 * h,
                                                             model_ref, lvl_tbl, color)

        for seq, seg, comid, _pts, lf, cf, m in GISExporter._raindrop_rows(res):
            r = len(rows)
            is_ov = seg == "OVERLAND"
            rows.append(("", "OVERLAND FLOW" if is_ov else f"STREAM REACH {seq - 1}", comid or "-",
                         f(lf), f(cf / 5280.0, 3), m or "-"))
            symbols[(r, 0)] = _line_sym(ov_col, 2, 2) if is_ov else _line_sym(ch_col, cfg["weight"], cfg["style"])
            colors[r] = ov_col if is_ov else ch_col
        r = len(rows)
        rows.append(("", "TOTAL", "-", f(res["total_ft"]), f(res["total_ft"] / 5280.0, 3),
                     f"Max downstream {res.get('max_miles', 0):g} mi"))
        for label, desc, fn, col in (
                ("FLOW DIRECTION", "Arrowheads point downstream", _arrow_sym(ch_col), ch_col),
                ("START / END POINT", "Raindrop start and trace end", _circle_x_sym(ch_col), ch_col)):
            rr = len(rows)
            rows.append(("", label, "", "", "", desc))
            symbols[(rr, 0)] = fn
            colors[rr] = col
        lon, lat = res["query_point"]
        title = f"{trace_id} FLOW (RAINDROP) PATH  \u2013  start {lat:.6f}, {lon:.6f}  ({datetime.date.today():%Y-%m-%d})"
        self._build_fallback_line_table(rows, tx, ty, [], model_ref, lvl_tbl, title=title, th=th,
                                        symbols=symbols, min_widths={0: 6.0 * th}, row_colors=colors)
        self.log(f"  Placed raindrop table for {trace_id} ({len(rows) - 1} rows incl. legend).")
        return True

    def _place_text_callout(self, origin_uor, text_str, model_ref, level_id, color, th=None, offset=None):
        dgn_model = model_ref.GetDgnModel()
        tprops = TextBlockProperties.Create(dgn_model)
        pprops = ParagraphProperties.Create(dgn_model)
        sym = self._sym_uor()
        th = th or 7.0 * sym
        dx, dy = offset or (11.0 * sym, 5.0 * sym)
        rprops = RunProperties.Create(DgnFontManager.GetDefaultTrueTypeFont(), DPoint2d(th, th), dgn_model)

        tb = TextBlock(tprops, pprops, rprops, dgn_model)
        tb.SetUserOrigin(DPoint3d(origin_uor.x + dx, origin_uor.y + dy, origin_uor.z))
        tb.AppendText(text_str)

        teh = EditElementHandle()
        if _status_ok(TextHandlerBase.CreateElement(teh, None, tb)):
            self._apply_symbology(teh, level_id, color, 1, 0)
            self._add_to_model(teh)

    def _ensure_basin_item_type(self):
        return self._ensure_item_type(self.ITEM_TYPE_NAME, self.BASIN_ITEM_PROPS)

    def _ensure_item_type(self, type_name, item_props):
        if not IN_MICROSTATION:
            return None
        dgn_file = ISessionMgr.ActiveDgnFile
        lib = None
        try:
            lib = ItemTypeLibrary.FindByName(self.ITEM_LIBRARY_NAME, dgn_file)
        except Exception:
            lib = None
        if lib is None:
            try:
                lib = ItemTypeLibrary(self.ITEM_LIBRARY_NAME, dgn_file, False)
            except TypeError:
                lib = ItemTypeLibrary(self.ITEM_LIBRARY_NAME, dgn_file)

        itype = lib.GetItemTypeByName(type_name)
        if itype is None:
            try:
                itype = lib.AddItemType(type_name, False)
            except TypeError:
                itype = lib.AddItemType(type_name)

        S, D = CustomProperty.Type1.eString, CustomProperty.Type1.eDouble
        prop_specs = [(n, D if t == "d" else S) for n, t in item_props]
        changed = False
        for name, ptype in prop_specs:
            if itype.GetPropertyByName(name) is None:
                p = itype.AddProperty(name)
                if p:
                    p.SetType(ptype)
                changed = True
        if changed:
            lib.Write()
        # The ItemType is owned by the library: keep the library alive or the ItemType becomes a dangling
        # pointer and ApplyCustomItem crashes MicroStation (access violation).
        self._item_libs = getattr(self, "_item_libs", {})
        self._item_libs[type_name] = lib
        return lib.GetItemTypeByName(type_name) or itype

    @classmethod
    def _basin_item_values(cls, ss_result, culvert_id=""):
        """Item Type values for one drainage area: identity, basin characteristics, C, Tc, flowpath and peak flows."""
        vcomp = ss_result.get("vdot_comparison") or {}
        met = vcomp.get("metrics", {})

        def num(v, nd):
            return round(float(v), nd) if isinstance(v, (int, float)) and not isinstance(v, bool) else None

        sqmi = num(met.get("drnarea_sqmi"), 4)
        if sqmi is None:
            sqmi = num(ss_result.get("drainage_area_sqmi"), 4) or 0.0
        acres = num(met.get("drnarea_acres"), 2)
        if acres is None:
            acres = round(sqmi * 640.0, 2)
        vals = {
            "CulvertID": str(culvert_id or "CULVERT-01"),
            "StateCode": str(ss_result.get("rcode", "VA")),
            "WorkspaceID": str(ss_result.get("workspace_id", "N/A")),
            "DrainageArea_SqMi": sqmi, "DrainageArea_Acres": acres,
            "Slope_FtPerMi": num(met.get("slope_ft_mi"), 2),
            "Impervious_Pct": num(met.get("imperv_pct"), 2),
            "Forest_Pct": num(met.get("forest_pct"), 2),
            "RunoffCoeff_C": num(vcomp.get("c_coeff"), 3),
            "Tc_Min": num(vcomp.get("tc_min"), 1),
            "LongestFlowpath_Ft": num(ss_result.get("lfp_length_ft"), 1),
            "LongestFlowpath_Method": ss_result.get("lfp_method") or None,
            "GoverningMethod": vcomp.get("governing_method") or None,
            "CalcDate": datetime.date.today().isoformat(),
            "CurveNumber_CN": num(vcomp.get("cn"), 1),
            "CurveNumber_Source": vcomp.get("cn_source") or None,
        }
        hsg = {k: float(v) for k, v in (ss_result.get("hsg") or {}).items() if k in ("A", "B", "C", "D")}
        total = sum(hsg.values())
        if total > 0:
            for grp in "ABCD":
                vals[f"SoilGroup_{grp}_Pct"] = round(100.0 * hsg.get(grp, 0.0) / total, 1)
            vals["DominantSoilGroup"] = max(sorted(hsg), key=hsg.get)
            vals["SoilSource"] = ss_result.get("hsg_source") or "USDA SSURGO (Soil Data Access); dual groups counted as D"
        r_map = {r.get("return_yr"): r for r in vcomp.get("rows", [])}
        for yr in cls.RETURN_PERIODS:
            r = r_map.get(yr, {})
            vals[f"Q{yr}_USGS_cfs"] = num(r.get("usgs_q_cfs"), 1)
            vals[f"Q{yr}_VDOT_cfs"] = num(r.get("vdot_q_cfs"), 1)
        return vals

    def attach_basin_item_type(self, basin_eeh, ss_result, culvert_id=""):
        """Attaches the 'DrainageArea' Item Type (IDs, basin characteristics, C, Tc, flowpath, peak flows)."""
        if not IN_MICROSTATION or not self.ATTACH_ITEM_TYPE:
            return
        try:
            _trace("attach_basin_item_type")
            itype = self._ensure_basin_item_type()
            if not itype:
                return
            # The boundary was just created, so it has no item yet. Do not call GetCustomItem here:
            # in MSPy it returns a null instance that crashes MicroStation (access violation) when tested.
            host = CustomItemHost(basin_eeh, False)
            inst = host.ApplyCustomItem(itype)
            if inst is None:
                self.log("  Item Type note: could not attach to the drainage-area boundary.")
                return

            values = self._basin_item_values(ss_result, culvert_id)
            n_set = 0
            for name, kind in self.BASIN_ITEM_PROPS:
                v = values.get(name)
                if kind == "d":
                    if isinstance(v, bool) or not isinstance(v, (int, float)):
                        continue  # leave blank rather than writing a misleading 0
                    v = float(v)
                elif v is None:
                    continue
                else:
                    v = str(v)
                try:
                    inst.SetValue(name, ECValue(v))
                    n_set += 1
                except Exception as e:
                    _trace(f"item SetValue {name} failed: {e}")
            inst.WriteChanges()
            _trace("attach_basin_item_type done")
            q100 = values.get("Q100_VDOT_cfs")
            q_txt = f", Q100 {q100:,.1f} cfs" if isinstance(q100, (int, float)) else ""
            self.log(f"  Attached Item Type '{self.ITEM_LIBRARY_NAME}:{self.ITEM_TYPE_NAME}' to the drainage area "
                     f"({culvert_id or 'CULVERT-01'}, {values['DrainageArea_Acres']:,.1f} ac{q_txt}; "
                     f"{n_set} properties).")
        except Exception as e:
            _trace("attach_basin_item_type error:\n" + traceback.format_exc())
            self.log(f"  Item Type note: {e}")

    COMPASS_8 = ("N", "NE", "E", "SE", "S", "SW", "W", "NW")
    NHD_FTYPES = {"460": "Stream/River", "558": "Artificial Path", "336": "Canal/Ditch", "334": "Connector",
                  "420": "Underground Conduit", "428": "Pipeline", "566": "Coastline"}

    @classmethod
    def _stream_item_values(cls, stream, pts_uor):
        """Item Type values for one tributary reach: identity, flow direction, receiving stream and NHD flows."""
        nhd = stream.get("nhd") or {}

        def num(v, nd, lo=None):
            try:
                f = float(v)
            except (TypeError, ValueError):
                return None
            if lo is not None and f < lo:
                return None  # NHDPlus uses -9998/-9999 for missing values
            return round(f, nd)

        (x0, y0), (x1, y1) = pts_uor[0], pts_uor[-1]
        bearing = None
        direction = "Downstream (digitized direction)"
        if math.hypot(x1 - x0, y1 - y0) > 0:
            bearing = round(math.degrees(math.atan2(x1 - x0, y1 - y0)) % 360.0, 1)
            direction = f"Flows {cls.COMPASS_8[int((bearing + 22.5) // 45) % 8]} ({bearing:.0f}\u00b0 grid)"
        nid = nhd.get("nhdplusid") or nhd.get("permanent_identifier")
        if isinstance(nid, float) and nid.is_integer():
            nid = int(nid)
        return {
            "StreamName": str(stream.get("name", "")),
            "NHDPlusID": str(nid) if nid not in (None, "") else None,
            "ReachCode": str(nhd.get("reachcode")) if nhd.get("reachcode") else None,
            "FeatureType": cls.NHD_FTYPES.get(str(nhd.get("ftype")), str(nhd.get("ftype") or nhd.get("fcode") or "")) or None,
            "StreamOrder": num(stream.get("stream_order"), 0, 0),
            "FlowDirection": direction,
            "FlowBearing_Deg": bearing,
            "FlowsInto": stream.get("flows_into") or None,
            "MeanAnnualFlow_cfs": num(nhd.get("qema"), 2, 0),
            "MeanVelocity_fps": num(nhd.get("vema"), 2, 0),
            "DrainageArea_SqMi": num((stream.get("tot_da_sqkm") or 0) * 0.386102, 4, 0) or None,
            "Slope_FtPerFt": num(nhd.get("slope"), 5, 0),
            "ClippedLength_Ft": num(stream.get("length_ft"), 1, 0),
            "Source": "USGS NHDPlus HR",
            "QueryDate": datetime.date.today().isoformat(),
        }

    def _write_item_values(self, inst, item_props, values):
        n_set = 0
        for name, kind in item_props:
            v = values.get(name)
            if kind == "d":
                if isinstance(v, bool) or not isinstance(v, (int, float)):
                    continue
                v = float(v)
            elif v is None:
                continue
            else:
                v = str(v)
            try:
                inst.SetValue(name, ECValue(v))
                n_set += 1
            except Exception as e:
                _trace(f"item SetValue {name} failed: {e}")
        inst.WriteChanges()
        return n_set

    @classmethod
    def _flowpath_item_values(cls, ss_result, culvert_id, pts_uor, seg_no=1, n_segs=1):
        """Item Type values for the longest flowpath: length, method, direction, Tc and end points."""
        vcomp = ss_result.get("vdot_comparison") or {}
        met = vcomp.get("metrics", {})
        path = ss_result.get("longest_flowpath_wgs84") or []

        def num(v, nd):
            return round(float(v), nd) if isinstance(v, (int, float)) and not isinstance(v, bool) else None

        bearing, direction = None, None
        if len(pts_uor) >= 2:
            (x0, y0), (x1, y1) = pts_uor[0][:2], pts_uor[-1][:2]
            if math.hypot(x1 - x0, y1 - y0) > 0:
                bearing = round(math.degrees(math.atan2(x1 - x0, y1 - y0)) % 360.0, 1)
                direction = (f"Flows {cls.COMPASS_8[int((bearing + 22.5) // 45) % 8]} "
                             f"({bearing:.0f}\u00b0 grid) to outlet")
        length_ft = num(ss_result.get("lfp_length_ft"), 1)
        up = path[0] if path else (None, None)
        out = path[-1] if path else (None, None)
        return {
            "CulvertID": str(culvert_id or "CULVERT-01"),
            "Length_Ft": length_ft,
            "Length_Mi": round(length_ft / 5280.0, 3) if length_ft else None,
            "Method": ss_result.get("lfp_method") or None,
            "FlowDirection": direction,
            "FlowBearing_Deg": bearing,
            "Segment": f"{seg_no} of {n_segs}",
            "BasinSlope_FtPerMi": num(met.get("slope_ft_mi"), 2),
            "Tc_Min": num(vcomp.get("tc_min"), 1),
            "TcMethod": "Kirpich" if vcomp.get("tc_min") is not None else None,
            "DrainageArea_Acres": num(met.get("drnarea_acres"), 2),
            "Upstream_Lat": num(up[1], 6), "Upstream_Lon": num(up[0], 6),
            "Outlet_Lat": num(out[1], 6), "Outlet_Lon": num(out[0], 6),
            "CalcDate": datetime.date.today().isoformat(),
        }

    def attach_flowpath_item_type(self, eeh, ss_result, culvert_id, pts_uor, seg_no=1, n_segs=1):
        """Attaches the 'LongestFlowpath' Item Type to a just-added flowpath line."""
        if not IN_MICROSTATION or not self.ATTACH_ITEM_TYPE:
            return False
        try:
            _trace("attach_flowpath_item_type")
            itype = self._ensure_item_type(self.FLOWPATH_ITEM_TYPE_NAME, self.FLOWPATH_ITEM_PROPS)
            if not itype:
                return False
            inst = CustomItemHost(eeh, False).ApplyCustomItem(itype)  # never GetCustomItem (crashes)
            if inst is None:
                return False
            self._write_item_values(inst, self.FLOWPATH_ITEM_PROPS,
                                    self._flowpath_item_values(ss_result, culvert_id, pts_uor, seg_no, n_segs))
            return True
        except Exception:
            _trace("attach_flowpath_item_type error:\n" + traceback.format_exc())
            return False

    @classmethod
    def _raindrop_item_values(cls, res, trace_id, kind, pts_uor, seg_no=1, n_segs=1, uor_per_ft=1000.0):
        """Item Type values for one raindrop-path line (overland leg or downstream channel)."""
        def num(v, nd):
            return round(float(v), nd) if isinstance(v, (int, float)) and not isinstance(v, bool) else None

        bearing, direction = None, None
        if len(pts_uor) >= 2:
            (x0, y0), (x1, y1) = pts_uor[0][:2], pts_uor[-1][:2]
            if math.hypot(x1 - x0, y1 - y0) > 0:
                bearing = round(math.degrees(math.atan2(x1 - x0, y1 - y0)) % 360.0, 1)
                direction = f"Flows {cls.COMPASS_8[int((bearing + 22.5) // 45) % 8]} ({bearing:.0f}\u00b0 grid)"
        seg_len = sum(math.hypot(b[0] - a[0], b[1] - a[1]) for a, b in zip(pts_uor, pts_uor[1:])) / (uor_per_ft or 1.0)
        reaches = res.get("reaches") or []
        start = res.get("query_point") or (None, None)
        path = res.get("path_wgs84") or []
        end = path[-1] if path else (None, None)
        total = num(res.get("total_ft"), 1)
        overland = kind.startswith("Overland")
        return {
            "TraceID": str(trace_id),
            "Segment": kind if n_segs <= 1 else f"{kind} {seg_no} of {n_segs}",
            "Method": (res.get("overland_method") if overland else res.get("source")) or None,
            "FlowDirection": direction,
            "FlowBearing_Deg": bearing,
            "SegmentLength_Ft": round(seg_len, 1),
            "OverlandLength_Ft": num(res.get("overland_ft"), 1),
            "TotalLength_Ft": total,
            "TotalLength_Mi": round(total / 5280.0, 3) if total else None,
            "MaxDownstream_Mi": num(res.get("max_miles"), 2),
            "ReachCount": float(len(reaches)),
            "StartReachCOMID": str(reaches[0].get("comid") or "") or None if reaches else None,
            "EndReachCOMID": str(reaches[-1].get("comid") or "") or None if reaches else None,
            "Start_Lat": num(start[1], 6), "Start_Lon": num(start[0], 6),
            "End_Lat": num(end[1], 6), "End_Lon": num(end[0], 6),
            "CalcDate": datetime.date.today().isoformat(),
        }

    def attach_raindrop_item_type(self, eeh, res, trace_id, kind, pts_uor, seg_no=1, n_segs=1):
        """Attaches the 'RaindropPath' Item Type to a just-added raindrop path line."""
        if not IN_MICROSTATION or not self.ATTACH_ITEM_TYPE:
            return False
        try:
            _trace("attach_raindrop_item_type")
            itype = self._ensure_item_type(self.RAINDROP_ITEM_TYPE_NAME, self.RAINDROP_ITEM_PROPS)
            if not itype:
                return False
            inst = CustomItemHost(eeh, False).ApplyCustomItem(itype)  # never GetCustomItem (crashes)
            if inst is None:
                return False
            vals = self._raindrop_item_values(res, trace_id, kind, pts_uor, seg_no, n_segs,
                                              self.projector.uor_per_master)
            self._write_item_values(inst, self.RAINDROP_ITEM_PROPS, vals)
            return True
        except Exception:
            _trace("attach_raindrop_item_type error:\n" + traceback.format_exc())
            return False

    def attach_stream_item_type(self, stream_eeh, stream, pts_uor, itype=None):
        """Attaches the 'Tributary' Item Type (name, flow direction, receiving stream, NHD flows) to a stream line."""
        if not IN_MICROSTATION or not self.ATTACH_ITEM_TYPE:
            return False
        try:
            _trace("attach_stream_item_type")
            itype = self._ensure_item_type(self.STREAM_ITEM_TYPE_NAME, self.STREAM_ITEM_PROPS)
            if not itype:
                return False
            # Freshly created element: apply directly (never GetCustomItem - it crashes MicroStation).
            inst = CustomItemHost(stream_eeh, False).ApplyCustomItem(itype)
            if inst is None:
                return False
            self._write_item_values(inst, self.STREAM_ITEM_PROPS, self._stream_item_values(stream, pts_uor))
            return True
        except Exception:
            _trace("attach_stream_item_type error:\n" + traceback.format_exc())
            return False

    def draw_basin_rings(self, ss_result, culvert_id=""):
        """Draws StreamStats basin rings as closed shapes and attaches the drainage-area Item Type."""
        if not IN_MICROSTATION:
            n = len(ss_result.get("basin_rings_wgs84", []))
            self.log(f"  [Simulation] Would draw {n} basin ring(s) for {culvert_id}.")
            return 0
        model_ref = ISessionMgr.ActiveDgnModelRef
        lvl_ids = self.ensure_levels_exist()
        basin_lvl = lvl_ids.get("USGS_StreamStats_Basin")
        b_cfg = self.LEVEL_CONFIG["USGS_StreamStats_Basin"]
        drawn = 0
        existing = 0
        for ring_idx, ring in enumerate(ss_result.get("basin_rings_wgs84", [])):
            if len(ring) > self.MAX_SHAPE_VERTICES:
                n0 = len(ring)
                ring = _simplify_ring(list(ring), self.MAX_SHAPE_VERTICES)
                self.log(f"  Simplified basin ring {ring_idx + 1}: {n0} -> {len(ring)} vertices.")
            _trace(f"draw_basin_rings ring={ring_idx} pts={len(ring)}")
            dpts = DPoint3dArray()
            uor = []
            for lon, lat in ring:
                xu, yu, zu = self.projector.wgs84_to_uor(lon, lat)
                dpts.append(DPoint3d(xu, yu, zu))
                uor.append((xu, yu))
            if len(dpts) < 3:
                continue
            if self._already_drawn("USGS_StreamStats_Basin", uor, 6):
                existing += 1
                continue
            eeh = EditElementHandle()
            if ShapeHandler.CreateShapeElement(eeh, None, dpts, model_ref.Is3d(), model_ref) == BentleyStatus.eSUCCESS:
                self._apply_symbology(eeh, basin_lvl, b_cfg["color"], b_cfg["weight"], b_cfg["style"])
                eeh.AddToModel()
                drawn += 1
                if ring_idx == 0:
                    self.attach_basin_item_type(eeh, ss_result, culvert_id)
                    self.attach_sources_item_type(eeh, ss_result, culvert_id)
        if existing:
            self.log(f"  Drainage area for {culvert_id or 'culvert'} is already in the DGN "
                     f"(same outlet as an earlier culvert or run); not drawn again.")
        if drawn or existing:
            ss_result["_dgn_basin_drawn"] = True
            self.ref_extent_uor = max(self.ref_extent_uor, self.basin_extent_uor(ss_result))
        return drawn + existing

    @staticmethod
    def _fmt_num(v, dec=1, unit=""):
        if isinstance(v, bool) or not isinstance(v, (int, float)):
            try:
                v = float(str(v).replace(",", ""))
            except (TypeError, ValueError):
                return "-" if v in (None, "", "-") else str(v)
        return f"{v:,.{dec}f}{(' ' + unit) if unit else ''}"

    _batch_table_ids = None

    def place_batch_master_table(self, batch_results, tx, ty, replace_previous=False):
        """Draws the Culvert Analysis summary table and its legend. With `replace_previous`, the table drawn
        by the previous call is deleted first (used when only the failed culverts are re-run)."""
        if not IN_MICROSTATION:
            self.log(f"[Simulation] Placed Master Batch Table with {len(batch_results)} culverts.")
            return

        lvl_ids = self.ensure_levels_exist()
        model_ref = ISessionMgr.ActiveDgnModelRef
        rows = [
            ("CULVERT ID", "DRAINAGE (AC)", "SLOPE (FT/MI)", "Q10 (CFS)", "Q25 (CFS)", "Q50 (CFS)", "Q100 (CFS)",
             "GOVERNING METHOD", "BASIN DATA", "RAINFALL")
        ]
        f = self._fmt_num
        for item in batch_results:
            cid = item.get("culvert_id", "")
            ssr = item.get("ss_result")
            if not ssr:
                rows.append((cid, "FAILED", "-", "-", "-", "-", "-", "Error", "-", "-"))
                continue
            vcomp = ssr.get("vdot_comparison", {})
            met = vcomp.get("metrics", {})
            r_map = {r["return_yr"]: r for r in vcomp.get("rows", [])}
            qs = [f(r_map.get(yr, {}).get("vdot_q_cfs", "-"), 1) for yr in (10, 25, 50, 100)]
            rows.append((cid, f(met.get("drnarea_acres", 0.0), 1), f(met.get("slope_ft_mi", 0.0), 1), *qs,
                         vcomp.get("governing_method", "") or "-",
                         "ALTERNATE" if "Alternate" in str(vcomp.get("basin_source", "")) else "STREAMSTATS",
                         "ATLAS 14" if str(vcomp.get("rainfall_source", "")).startswith("NOAA") else "BUILT-IN"))

        widths = [110.0, 100.0, 100.0, 85.0, 85.0, 85.0, 95.0, 180.0, 100.0, 90.0]
        title = (f"BATCH CULVERT HYDROLOGY SUMMARY  \u2013  VDOT design peak flows  "
                 f"({len(batch_results)} culverts, {datetime.date.today():%Y-%m-%d})")
        th = self.table_text_height(batch_results)
        if replace_previous and self._batch_table_ids:
            try:
                self.delete_elements(self._batch_table_ids)
            except Exception as exc:
                self.log(f"  Previous summary table could not be removed ({exc}).")
        prev_rec, self._rec_ids = self._rec_ids, []
        try:
            info = self._build_text_table(rows, tx, ty, widths, model_ref, lvl_ids.get("USGS_StreamStats_Table"),
                                          title=title, th_uor=th)
            bottom = (info or {}).get("bottom")
            if bottom is None:
                bottom = ty - 2.4 * 1.33 * th - 2.2 * th * len(rows)
            try:
                leg = self.place_culvert_legend_table(tx, bottom - 2.0 * th, model_ref, lvl_ids, th)
                bottom = (leg or {}).get("bottom", bottom)
            except Exception as exc:
                self.log(f"  Culvert legend table skipped ({exc})")
            try:
                self.place_data_sources_table(batch_results, tx, bottom - 2.0 * th, model_ref, lvl_ids, th)
            except Exception as exc:
                self.log(f"  Data sources table skipped ({exc})")
        finally:
            self._batch_table_ids, self._rec_ids = list(self._rec_ids), prev_rec

    def place_culvert_legend_table(self, tx, ty, model_ref, lvl_ids, th):
        """Legend for the Culvert Analysis table: DGN symbols drawn in cells, plus column/method notes."""
        lvl = lvl_ids.get("USGS_StreamStats_Table")
        C = {k: v["color"] for k, v in self.LEVEL_CONFIG.items()}
        S = self.LEVEL_CONFIG

        def _line(color, weight, style):
            return lambda x0, x1, yc, h: self._add_linestring([(x0, yc, 0.0), (x1, yc, 0.0)], model_ref, lvl,
                                                              color, weight, style)

        def _inlet(x0, x1, yc, h):
            xm, a = (x0 + x1) / 2.0, 0.3 * h
            sq = [(xm - a, yc - a, 0.0), (xm + a, yc - a, 0.0), (xm + a, yc + a, 0.0), (xm - a, yc + a, 0.0),
                  (xm - a, yc - a, 0.0)]
            col = C["USGS_Culvert_Inlet"]
            self._add_linestring(sq, model_ref, lvl, col, 2, closed=True)
            self._add_linestring([sq[0], sq[2]], model_ref, lvl, col, 1)
            self._add_linestring([sq[1], sq[3]], model_ref, lvl, col, 1)

        def _pour(x0, x1, yc, h):
            xm, r = (x0 + x1) / 2.0, 0.25 * h
            col = C["USGS_StreamStats_Point"]
            eeh = EditElementHandle()
            if EllipseHandler.CreateEllipseElement(eeh, None, DPoint3d(xm, yc, 0.0), r, r, 0.0,
                                                   model_ref.Is3d(), model_ref) == BentleyStatus.eSUCCESS:
                self._apply_symbology(eeh, lvl, col, 3, 0)
                self._add_to_model(eeh)
            self._add_linestring([(xm - 1.5 * r, yc, 0.0), (xm + 1.5 * r, yc, 0.0)], model_ref, lvl, col, 1)
            self._add_linestring([(xm, yc - 1.5 * r, 0.0), (xm, yc + 1.5 * r, 0.0)], model_ref, lvl, col, 1)

        def _arrow(x0, x1, yc, h):
            col = C["USGS_NHD_Flow_Arrows"]
            self._add_linestring([(x0, yc, 0.0), (x1, yc, 0.0)], model_ref, lvl, C["USGS_NHD_Streams_Clipped"], 1, 0)
            a, xm = 0.35 * h, (x0 + x1) / 2.0
            self._add_filled_shape([(xm + a, yc, 0.0), (xm - a, yc + 0.6 * a, 0.0), (xm - a, yc - 0.6 * a, 0.0),
                                    (xm + a, yc, 0.0)], model_ref, lvl, col)

        def _cfg_line(name):
            cfg = S[name]
            return _line(cfg["color"], cfg["weight"], cfg["style"])

        entries = [  # (legend cell text, feature, description, symbol fn, text color)
            ("", "CULVERT INLET", "Culvert location (picked or imported)", _inlet, C["USGS_Culvert_Inlet"]),
            ("", "POUR POINT", "StreamStats snapped outlet", _pour, C["USGS_StreamStats_Point"]),
            ("", "DRAINAGE AREA", "StreamStats watershed boundary", _cfg_line("USGS_StreamStats_Basin"),
             C["USGS_StreamStats_Basin"]),
            ("", "NHD STREAM", "NHDPlus flowline within the drainage area", _cfg_line("USGS_NHD_Streams_Clipped"),
             None),
            ("", "FLOW DIRECTION", "Arrowheads point downstream", _arrow, None),
            ("", "LONGEST FLOWPATH", "Hydraulic length used for time of concentration",
             _cfg_line("USGS_Longest_Flowpath"), C["USGS_Longest_Flowpath"]),
            ("DRAINAGE (AC)", "COLUMN", "Delineated drainage area, acres", None, None),
            ("SLOPE (FT/MI)", "COLUMN", "Basin mean channel slope, feet per mile", None, None),
            ("Q10 - Q100", "COLUMN", "VDOT design peak discharge (cfs) for the 10- to 100-yr storm", None, None),
            ("RATIONAL", "GOVERNING METHOD", "VDOT Rational Q = Cf\u00b7C\u00b7i\u00b7A, drainage area <= 200 ac",
             None, None),
            ("REGRESSION", "GOVERNING METHOD", "USGS regional regression (NSS), drainage area > 200 ac", None, None),
            ("STREAMSTATS", "BASIN DATA", "Slope and land cover from USGS StreamStats", None, None),
            ("ALTERNATE", "BASIN DATA", "StreamStats down: NLCD 2021 land cover, 3DEP slope, polygon area",
             None, 4),
            ("ATLAS 14", "RAINFALL", "NOAA Atlas 14 intensity at the pour point", None, None),
            ("BUILT-IN", "RAINFALL", "Statewide IDF coefficients (placeholder) \u2013 verify before design",
             None, 4),
        ]
        rows = [("SYMBOL", "FEATURE", "DESCRIPTION")]
        symbols, colors = {}, {}
        for code, feat, desc, fn, col in entries:
            r = len(rows)
            rows.append((code, feat, desc))
            if fn:
                symbols[(r, 0)] = fn
            if col is not None:
                colors[r] = col
        return self._build_fallback_line_table(rows, tx, ty, [], model_ref, lvl, title="LEGEND", th=th,
                                               symbols=symbols, min_widths={0: 6.0 * th}, row_colors=colors)

    @classmethod
    def data_source_rows(cls, batch_results):
        """Rows (data, source, website) for the DATA SOURCES table: only the sources this batch actually used."""
        results = [i.get("ss_result") for i in batch_results or [] if i.get("ss_result")]
        return [("DATA", "SOURCE", "WEBSITE")] + [e[1:] for e in cls.data_source_entries(results)]

    # Keys of data_source_entries; each becomes <key>_Source and <key>_Website on the DataSources Item Type.
    DATA_SOURCE_KEYS = ("Watershed", "BasinCharacteristics", "LandCover", "ChannelSlope", "Streams",
                        "LongestFlowpath", "RuralRegression", "UrbanRegression", "Anderson", "Rainfall",
                        "SoilGroups", "Methods")
    SOURCES_ITEM_TYPE_NAME = "DataSources"
    SOURCES_ITEM_PROPS = ([("CulvertID", "s")]
                          + [(f"{k}_{part}", "s") for k in DATA_SOURCE_KEYS for part in ("Source", "Website")]
                          + [("CalcDate", "s")])

    @staticmethod
    def data_source_entries(results):
        """(key, data, source, website) for every data source used by the given delineation results."""
        vcomps = [r.get("vdot_comparison") or {} for r in results]

        def any_v(key, text):
            return any(text in str(v.get(key, "")) for v in vcomps)

        def any_method(*keys):
            return any(row.get("design_method") in keys or any(row.get(f"q_{k}_cfs") is not None for k in keys)
                       for v in vcomps for row in v.get("rows", []))

        rows = [("Watershed", "Watershed boundary, pour point", "USGS StreamStats", "https://streamstats.usgs.gov")]
        if any_v("basin_source", "StreamStats"):
            rows.append(("BasinCharacteristics", "Basin characteristics", "USGS StreamStats",
                         "https://streamstats.usgs.gov"))
        if any(r.get("nlcd") for r in results):
            rows.append(("LandCover", "Land cover, impervious, % urban", "MRLC National Land Cover Database 2021",
                         "https://www.mrlc.gov"))
        if any_v("basin_source", "Alternate"):
            rows.append(("ChannelSlope", "Channel slope (StreamStats down)", "USGS 3DEP elevation",
                         "https://elevation.nationalmap.gov"))
        rows.append(("Streams", "Streams, flow direction", "USGS National Hydrography Dataset (NHDPlus HR)",
                     "https://hydro.nationalmap.gov"))
        if any("NHD" in str(r.get("lfp_method", "")) for r in results):
            rows.append(("LongestFlowpath", "Longest flowpath", "Traced from USGS NHD flowlines",
                         "https://hydro.nationalmap.gov"))
        if any("StreamStats" in str(r.get("lfp_method", "")) for r in results):
            rows.append(("LongestFlowpath", "Longest flowpath", "USGS StreamStats", "https://streamstats.usgs.gov"))
        if any(r.get("flow_stats") for r in results):
            rows.append(("RuralRegression", "Rural regression peak flows",
                         "USGS National Streamflow Statistics (NSS)", "https://streamstats.usgs.gov/nssservices"))
        if any_method("urban"):
            rows.append(("UrbanRegression", "Urban regression peak flows", "USGS SIR 2014-5090 equations",
                         "https://pubs.usgs.gov/publication/sir20145090"))
        if any_method("anderson"):
            rows.append(("Anderson", "Anderson method", "USGS Water-Supply Paper 2001-C",
                         "https://pubs.usgs.gov/publication/wsp01C"))
        if any_v("rainfall_source", "NOAA"):
            rows.append(("Rainfall", "Rainfall intensity and depth",
                         "NOAA Atlas 14 Precipitation Frequency Data Server", "https://hdsc.nws.noaa.gov/pfds"))
        if any("USDA" in str(r.get("hsg_source", "")) for r in results):
            rows.append(("SoilGroups", "Hydrologic soil groups", "USDA NRCS SSURGO (Soil Data Access)",
                         "https://sdmdataaccess.sc.egov.usda.gov"))
        if any("UC Davis" in str(r.get("hsg_source", "")) for r in results):
            rows.append(("SoilGroups", "Hydrologic soil groups", "UC Davis SoilWeb ISSR-800 grid (from SSURGO)",
                         "https://casoilresource.lawr.ucdavis.edu"))
        rows.append(("Methods", "Methods, size ranges, Cf factors", "VDOT Drainage Manual, Chapter 6 (Rev. 7/26)",
                     "https://www.vdot.virginia.gov"))
        return rows

    def attach_sources_item_type(self, basin_eeh, ss_result, culvert_id=""):
        """Attaches the 'DataSources' Item Type (source and website of each kind of data used for this
        culvert) to the drainage-area boundary, next to its 'DrainageArea' item."""
        if not IN_MICROSTATION or not self.ATTACH_ITEM_TYPE:
            return
        try:
            _trace("attach_sources_item_type")
            itype = self._ensure_item_type(self.SOURCES_ITEM_TYPE_NAME, self.SOURCES_ITEM_PROPS)
            if not itype:
                return
            # New element: apply only. GetCustomItem returns a null instance in MSPy that crashes MicroStation.
            inst = CustomItemHost(basin_eeh, False).ApplyCustomItem(itype)
            if inst is None:
                self.log("  Item Type note: could not attach the data sources to the drainage-area boundary.")
                return
            values = {"CulvertID": str(culvert_id or "CULVERT-01"), "CalcDate": datetime.date.today().isoformat()}
            for key, _data, source, website in self.data_source_entries([ss_result]):
                values[f"{key}_Source"], values[f"{key}_Website"] = source, website
            n_set = 0
            for name, _kind in self.SOURCES_ITEM_PROPS:
                if values.get(name):
                    try:
                        inst.SetValue(name, ECValue(str(values[name])))
                        n_set += 1
                    except Exception as e:
                        _trace(f"item SetValue {name} failed: {e}")
            inst.WriteChanges()
            _trace("attach_sources_item_type done")
            self.log(f"  Attached Item Type '{self.ITEM_LIBRARY_NAME}:{self.SOURCES_ITEM_TYPE_NAME}' to the drainage "
                     f"area ({n_set} properties).")
        except Exception as e:
            _trace("attach_sources_item_type error:\n" + traceback.format_exc())
            self.log(f"  Item Type note (data sources): {e}")

    def place_data_sources_table(self, batch_results, tx, ty, model_ref, lvl_ids, th):
        """Table listing where each kind of data in the summary came from, with its website."""
        return self._build_fallback_line_table(self.data_source_rows(batch_results), tx, ty, [], model_ref,
                                               lvl_ids.get("USGS_StreamStats_Table"), title="DATA SOURCES", th=th)

    def table_text_height(self, batch_results=()):
        """Table text height (UOR), from the Settings drawing scale."""
        return self.text_height_uor()

    def _build_text_table(self, table_rows, tx, ty, col_w_masters, model_ref, lvl_id, title=None, th_uor=None):
        dgn_model = model_ref.GetDgnModel()
        dgn_file = ISessionMgr.ActiveDgnFile
        uor = self.projector.uor_per_master
        num_rows = len(table_rows)
        num_cols = len(table_rows[0])
        th_uor = th_uor or 7.0 * uor
        col_scale = th_uor / (7.0 * uor)

        if self.SAFE_DGN_MODE:
            _trace("fallback_line_table")
            return self._build_fallback_line_table(table_rows, tx, ty, col_w_masters, model_ref, lvl_id, title=title,
                                                   th=th_uor)

        try:
            _trace("native TextTable")
            dstyle = None
            try:
                dstyle = DgnTextStyle.GetByName("USGS_Table_Style", dgn_file)
            except Exception:
                dstyle = None
            if dstyle is None:
                dstyle = DgnTextStyle("USGS_Table_Style", dgn_file)
                dstyle.SetDoubleProperty(TextStyleProperty.eTextStyle_Height, th_uor)
                dstyle.SetDoubleProperty(TextStyleProperty.eTextStyle_Width, th_uor)
                dstyle.Add(dgn_file)
            sid = dstyle.GetID()

            tt = TextTable.Create(num_rows, num_cols, sid, th_uor, dgn_model)
            tt.SetOrigin(DPoint3d(tx, ty, 0.0))
            for c_idx, w in enumerate(col_w_masters):
                col = tt.GetColumn(c_idx)
                if col:
                    col.SetWidth(w * uor * col_scale)

            for r_idx, rvals in enumerate(table_rows):
                for c_idx in range(num_cols):
                    cell = tt.GetCell(TableCellIndex(r_idx, c_idx))
                    if cell:
                        cell.SetTextString(str(rvals[c_idx]))

            teh = EditElementHandle()
            if _status_ok(TextTableHandler.CreateTextTableElement(teh, tt)):
                self._apply_symbology(teh, lvl_id, 0, 1, 0)
                teh.AddToModel()
                self.log(f"  Placed MicroStation TextTable ({num_rows}x{num_cols}) on level 'USGS_StreamStats_Table'.")
                return
            self.log("  TextTable element creation failed; drawing bordered CAD table fallback.")
        except Exception as e:
            self.log(f"  TextTable API note ({e}); drawing bordered CAD table fallback.")

        return self._build_fallback_line_table(table_rows, tx, ty, col_w_masters, model_ref, lvl_id, title=title,
                                               th=th_uor)

    _NUM_CELL_RE = re.compile(r"^[-+]?[\d,]*\.?\d+\s*(%|cfs|ac|sq\.?\s?mi|ft/mi|min|ft)?$", re.I)
    TABLE_TEXT_COLOR = 0
    # Color 0 follows the view background (white on black, black on white), so headings always contrast.
    TABLE_TITLE_COLOR = 0
    TABLE_HEADER_COLOR = 0

    @classmethod
    def _cell_align(cls, text, r_idx, c_idx):
        """'C' for headers and placeholders, 'R' for numbers (so decimals line up), 'L' for text."""
        t = text.strip()
        if r_idx == 0 or t in ("", "-", "\u2014", "N/A"):
            return "C"
        if c_idx > 0 and cls._NUM_CELL_RE.match(t):
            return "R"
        return "L"

    @staticmethod
    def _estimate_text_width(text, th):
        w = 0.0
        for ch in text:
            if ch == " ":
                w += 0.38
            elif ch in "il.,:;'|!()[]/-I1":
                w += 0.42
            elif ch.isupper() or ch in "MW%&@mw":
                w += 0.92
            else:
                w += 0.74
        return w * th

    def _make_cell_text(self, text, th, just, dgn_model, font):
        tprops = TextBlockProperties.Create(dgn_model)
        pprops = ParagraphProperties.Create(dgn_model)
        try:
            pprops.SetJustification(just)
        except Exception:
            pass
        rprops = RunProperties.Create(font, DPoint2d(th, th), dgn_model)
        tb = TextBlock(tprops, pprops, rprops, dgn_model)
        tb.AppendText(text)
        est = self._estimate_text_width(text, th)
        width = 0.0
        try:
            rng = tb.GetNominalRange()
            width = abs(rng.high.x - rng.low.x)
        except Exception:
            width = 0.0
        if not (0.4 * est <= width <= 2.5 * est):
            width = est
        return tb, width

    def _build_fallback_line_table(self, table_rows, tx, ty, col_w_masters, model_ref, lvl_id, title=None, th=None,
                                   symbols=None, min_widths=None, row_colors=None):
        """Draws a bordered table of lines + text. Columns auto-size to the measured text; every cell is
        justified to a point inside the cell (headers centred, numbers right-aligned, text left-aligned,
        all vertically centred) so the text lines up regardless of the font.
        `symbols` {(row, col): fn(x0, x1, yc, row_h)} draws graphics (legend samples) inside cells;
        `min_widths` {col: UOR}; `row_colors` {row: text color}."""
        uor = self.projector.uor_per_master
        dgn_model = model_ref.GetDgnModel()
        is3d = model_ref.Is3d()
        font = DgnFontManager.GetDefaultTrueTypeFont()
        th = th or 6.0 * uor
        title_th = 1.33 * th
        pad = 0.9 * th
        row_h = 2.2 * th
        title_h = 2.4 * title_th if title else 0.0
        ncols = max(len(r) for r in table_rows)
        J = TextElementJustification
        just_map = {"L": J.eLeftMiddle, "C": J.eCenterMiddle, "R": J.eRightMiddle}

        cells, widths = [], [4.0 * th] * ncols
        for r_idx, rvals in enumerate(table_rows):
            row = []
            for c_idx in range(ncols):
                txt = str(rvals[c_idx]) if c_idx < len(rvals) and rvals[c_idx] is not None else ""
                align = self._cell_align(txt, r_idx, c_idx)
                tb, w = self._make_cell_text(txt, th, just_map[align], dgn_model, font)
                row.append((tb, align, txt))
                widths[c_idx] = max(widths[c_idx], w + 2.0 * pad)
            cells.append(row)
        for c_idx, mw in (min_widths or {}).items():
            widths[c_idx] = max(widths[c_idx], mw)

        title_tb = None
        if title:
            title_tb, tw = self._make_cell_text(title, title_th, J.eLeftMiddle, dgn_model, font)
            if tw + 2.0 * pad > sum(widths):
                widths[-1] += tw + 2.0 * pad - sum(widths)
        total_w = sum(widths)
        body_top = ty - title_h
        bottom = body_top - row_h * len(table_rows)

        def _line(x1, y1, x2, y2, weight=0):
            pts = DPoint3dArray([DPoint3d(x1, y1, 0.0), DPoint3d(x2, y2, 0.0)])
            leh = EditElementHandle()
            if LineStringHandler.CreateLineStringElement(leh, None, pts, is3d, model_ref) == BentleyStatus.eSUCCESS:
                self._apply_symbology(leh, lvl_id, self.TABLE_TEXT_COLOR, weight, 0)
                self._add_to_model(leh)

        def _text(tb, x, y, color):
            tb.SetUserOrigin(DPoint3d(x, y, 0.0))
            teh = EditElementHandle()
            if _status_ok(TextHandlerBase.CreateElement(teh, None, tb)):
                self._apply_symbology(teh, lvl_id, color, 0, 0)
                self._add_to_model(teh)

        # Outer border (heavy), header separator (medium), body grid (thin).
        _line(tx, ty, tx + total_w, ty, 3)
        _line(tx, bottom, tx + total_w, bottom, 3)
        _line(tx, ty, tx, bottom, 3)
        _line(tx + total_w, ty, tx + total_w, bottom, 3)
        if title:
            _line(tx, body_top, tx + total_w, body_top, 2)
        for r in range(1, len(table_rows)):
            y = body_top - r * row_h
            _line(tx, y, tx + total_w, y, 2 if r == 1 else 0)
        x = tx
        for w in widths[:-1]:
            x += w
            _line(x, body_top, x, bottom, 0)

        if title_tb is not None:
            _text(title_tb, tx + pad, ty - title_h / 2.0, self.TABLE_TITLE_COLOR)
        for r_idx, row in enumerate(cells):
            yc = body_top - (r_idx + 0.5) * row_h
            x0 = tx
            for c_idx, (tb, align, txt) in enumerate(row):
                x1 = x0 + widths[c_idx]
                if txt:
                    xa = x0 + pad if align == "L" else (x1 - pad if align == "R" else (x0 + x1) / 2.0)
                    color = self.TABLE_HEADER_COLOR if r_idx == 0 else (row_colors or {}).get(r_idx, self.TABLE_TEXT_COLOR)
                    _text(tb, xa, yc, color)
                sym_fn = (symbols or {}).get((r_idx, c_idx))
                if sym_fn:
                    try:
                        sym_fn(x0 + pad, x1 - pad, yc, row_h)
                    except Exception as exc:
                        self.log(f"  Table legend symbol skipped ({exc})")
                x0 = x1
        self.log(f"  Placed bordered CAD table ({len(table_rows)} rows) on level 'USGS_StreamStats_Table'.")
        return {"width": total_w, "bottom": bottom}


# ---------------------------------------------------------------------------
# 9. Batch & Single Tabular Export Manager (Excel .xlsx and CSV)
# ---------------------------------------------------------------------------
class BatchSummaryExporter:
    @staticmethod
    def export_results_workbook(xlsx_path, results):
        """Results & Comparison workbook: a Summary sheet plus one sheet per culvert (metrics + flow comparison)."""
        h_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        h_font = Font(bold=True, color="FFFFFF")
        warn_fill = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
        q_fmt = "#,##0.0"

        def num(v):
            return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None

        def header(ws, labels):
            ws.append(labels)
            for c in range(1, len(labels) + 1):
                cell = ws.cell(row=ws.max_row, column=c)
                cell.fill, cell.font = h_fill, h_font

        def autosize(ws):
            for col in ws.columns:
                width = max(len(str(c.value)) for c in col if c.value is not None) if any(
                    c.value is not None for c in col) else 8
                ws.column_dimensions[col[0].column_letter].width = min(max(width + 3, 11), 42)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Summary"
        ws.append(["USGS StreamStats vs VDOT Rational \u2013 Results & Comparison"])
        ws["A1"].font = Font(bold=True, size=13)
        ws.append(["Generated", datetime.datetime.now().strftime("%Y-%m-%d %H:%M"), "Build By", "Robert Ratliff"])
        ws.append([])
        yrs = (2, 5, 10, 25, 50, 100, 200, 500)
        header(ws, ["Culvert ID", "State", "Drainage (sq mi)", "Drainage (ac)", "Slope (ft/mi)", "Rational C",
                    "Tc (min)", "Longest Flowpath (ft)", "Governing Method"]
               + [f"Q{y} USGS" for y in yrs] + [f"Q{y} VDOT" for y in yrs]
               + ["Basin Data Source", "Rainfall Source", "Tc Length Source", "Regression Source"])
        ws.freeze_panes = f"B{ws.max_row + 1}"

        used = {"Summary"}
        for cid, ssr in results.items():
            vcomp = ssr.get("vdot_comparison", {}) or {}
            met = vcomp.get("metrics", {}) or {}
            rows = vcomp.get("rows", []) or []
            r_map = {r.get("return_yr"): r for r in rows}
            ws.append([cid, ssr.get("rcode", ""), num(met.get("drnarea_sqmi")), num(met.get("drnarea_acres")),
                       num(met.get("slope_ft_mi")), num(vcomp.get("c_coeff")), num(vcomp.get("tc_min")),
                       num(ssr.get("lfp_length_ft")), vcomp.get("governing_method", "")]
                      + [num(r_map.get(y, {}).get("usgs_q_cfs")) for y in yrs]
                      + [num(r_map.get(y, {}).get("vdot_q_cfs")) for y in yrs]
                      + [vcomp.get("basin_source", ""), vcomp.get("rainfall_source", ""),
                         vcomp.get("tc_length_source", ""), vcomp.get("regression_source", "")])
            for c in range(3, ws.max_column + 1):
                ws.cell(row=ws.max_row, column=c).number_format = "#,##0.000" if c == 3 else (
                    "0.00" if c == 6 else q_fmt)

            name = re.sub(r"[\[\]:*?/\\]", "_", str(cid))[:31] or "Culvert"
            base, n = name, 2
            while name in used:
                name = f"{base[:28]}_{n}"
                n += 1
            used.add(name)
            cs = wb.create_sheet(name)
            cs.append([f"Culvert {cid} \u2013 Results & Comparison"])
            cs["A1"].font = Font(bold=True, size=13)
            cs.append([])
            header(cs, ["Basin Metric", "Value", "Units"])
            snapped = ssr.get("snapped_point") or ssr.get("query_point") or (None, None)
            for label, val, unit in (
                ("Drainage Area", num(met.get("drnarea_sqmi")), "sq mi"),
                ("Drainage Area", num(met.get("drnarea_acres")), "ac"),
                ("Basin Slope", num(met.get("slope_ft_mi")), "ft/mi"),
                ("Impervious", num(met.get("imperv_pct")), "%"),
                ("Forest", num(met.get("forest_pct")), "%"),
                ("Rational C", num(vcomp.get("c_coeff")), ""),
                ("Time of Concentration", num(vcomp.get("tc_min")), "min"),
                ("Tc Flow Length", num(vcomp.get("tc_length_ft")), "ft"),
                ("Tc Length Source", vcomp.get("tc_length_source", ""), ""),
                ("Longest Flowpath", num(ssr.get("lfp_length_ft")), "ft"),
                ("Flowpath Method", ssr.get("lfp_method", ""), ""),
                ("Governing Method", vcomp.get("governing_method", ""), ""),
                ("Basin Data Source", vcomp.get("basin_source", ""), ""),
                ("Rainfall Source", vcomp.get("rainfall_source", ""), ""),
                ("Regression Source", vcomp.get("regression_source", ""), ""),
                ("Percent Urban (developed)", num(vcomp.get("urban_pct")), "%"),
                ("NRCS Curve Number", num(vcomp.get("cn")), ""),
                ("Curve Number Source", vcomp.get("cn_source", ""), ""),
                ("State / Workspace", f"{ssr.get('rcode', '')} / {ssr.get('workspace_id', '')}", ""),
                ("Pour Point (Lon, Lat)", f"{snapped[0]:.5f}, {snapped[1]:.5f}" if num(snapped[0]) is not None
                 else "", "WGS84"),
            ):
                cs.append([label, val, unit])
            cs.append([])
            header(cs, ["Event", "USGS Rural Regression (cfs)", "Rational (cfs)", "NRCS TR-55 (cfs)",
                        "Anderson (cfs)", "USGS Urban Regression (cfs)", "VDOT Design (cfs)", "Design vs USGS",
                        "Design Method", "Rainfall Intensity (in/hr)", "24-hr Rainfall (in)"])
            for r in rows:
                usgs = num(r.get("usgs_q_cfs"))
                cs.append([f"Q{r['return_yr']} ({r['return_yr']}-yr)", "N/A" if usgs is None else usgs,
                           num(r.get("q_rational_cfs")), num(r.get("q_nrcs_cfs")), num(r.get("q_anderson_cfs")),
                           num(r.get("q_urban_cfs")), num(r.get("vdot_q_cfs")), r.get("pct_diff", "N/A"),
                           VDOTHydrologyCalculator.METHOD_LABELS.get(r.get("design_method"), ""),
                           num(r.get("intensity_in_hr")), num(r.get("rain_24h_in"))])
                for c in range(2, 8):
                    cs.cell(row=cs.max_row, column=c).number_format = q_fmt
                try:
                    if abs(float(str(r.get("pct_diff", "")).rstrip("%"))) > 25.0:
                        cs.cell(row=cs.max_row, column=8).fill = warn_fill
                except ValueError:
                    pass
            cs.append([])
            header(cs, ["Peak Discharge Method (VDOT Table 6-2)", "VDOT Size Range", "Status"])
            for m in vcomp.get("methods", []) or []:
                cs.append([m.get("label", ""), m.get("range", ""), m.get("status", "")])
            cs.append([])
            sources = " ".join(str(vcomp.get(k, "")) for k in ("rainfall_source", "regression_source"))
            if "placeholder" in sources:
                cs.append(["Note: built-in (placeholder) IDF or regression coefficients were used \u2013 see the "
                           "source rows above and verify against the VDOT Drainage Manual before design use."])
            else:
                cs.append(["Note: rainfall is NOAA Atlas 14 at the pour point; the Rational C is estimated from "
                           "land cover. Verify against the VDOT Drainage Manual before design use."])
            cs.cell(row=cs.max_row, column=1).font = Font(italic=True, color="7F7F7F")
            autosize(cs)
            cs.column_dimensions["A"].width = 26
        autosize(ws)
        ws.column_dimensions["A"].width = 16
        _repair_cleared_contextlib()
        wb.save(xlsx_path)
        return xlsx_path

    @classmethod
    def export_batch_summary(cls, batch_results, rcode):
        dgn_dir, dgn_stem = MicroStationDgnManager.get_active_dgn_directory_and_stem()
        csv_path = cls._writable_path(os.path.join(dgn_dir, f"{dgn_stem}_Batch_Culvert_Hydrology.csv"))
        xlsx_path = cls._writable_path(os.path.join(dgn_dir, f"{dgn_stem}_Batch_Culvert_Hydrology.xlsx"))

        with open(csv_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(["BATCH CULVERT HYDROLOGY & DRAINAGE SUMMARY"])
            writer.writerow(["Date", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "Region", rcode])
            writer.writerow([])
            writer.writerow([
                "Culvert ID", "Status", "Drainage (sq.mi)", "Drainage (acres)",
                "Slope (ft/mi)", "VDOT C", "Tc (min)",
                "Q2 USGS (cfs)", "Q10 USGS (cfs)", "Q25 USGS (cfs)", "Q50 USGS (cfs)", "Q100 USGS (cfs)",
                "Q10 VDOT (cfs)", "Q25 VDOT (cfs)", "Q50 VDOT (cfs)", "Q100 VDOT (cfs)",
                "Governing Method", "Exit Lon", "Exit Lat",
                "Basin Data Source", "Rainfall Source", "Tc Length Source", "Regression Source"
            ])
            for item in batch_results:
                cid = item.get("culvert_id", "")
                ssr = item.get("ss_result")
                if not ssr:
                    writer.writerow([cid, "FAILED"] + [""] * 21)
                    continue
                vcomp = ssr.get("vdot_comparison", {})
                met = vcomp.get("metrics", {})
                r_map = {r["return_yr"]: r for r in vcomp.get("rows", [])}
                writer.writerow([
                    cid, "Success", f"{met.get('drnarea_sqmi',0):.3f}", f"{met.get('drnarea_acres',0):.1f}",
                    f"{met.get('slope_ft_mi',0):.1f}", vcomp.get("c_coeff", ""), vcomp.get("tc_min", ""),
                    r_map.get(2, {}).get("usgs_q_cfs", ""), r_map.get(10, {}).get("usgs_q_cfs", ""),
                    r_map.get(25, {}).get("usgs_q_cfs", ""), r_map.get(50, {}).get("usgs_q_cfs", ""),
                    r_map.get(100, {}).get("usgs_q_cfs", ""),
                    r_map.get(10, {}).get("vdot_q_cfs", ""), r_map.get(25, {}).get("vdot_q_cfs", ""),
                    r_map.get(50, {}).get("vdot_q_cfs", ""), r_map.get(100, {}).get("vdot_q_cfs", ""),
                    vcomp.get("governing_method", ""),
                    ssr["snapped_point"][0], ssr["snapped_point"][1],
                    vcomp.get("basin_source", ""), vcomp.get("rainfall_source", ""),
                    vcomp.get("tc_length_source", ""), vcomp.get("regression_source", ""),
                ])

        if HAS_OPENPYXL:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Batch_Culverts"
            h_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
            h_font = Font(bold=True, color="FFFFFF")

            ws.append(["VDOT / USGS Batch Culvert Hydrology Matrix"])
            ws["A1"].font = Font(bold=True, size=12)
            ws.append(["Generated", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
            ws.append([])
            headers = [
                "Culvert ID", "Status", "Drainage (ac)", "Slope (ft/mi)", "VDOT C", "Tc (min)",
                "Q10 USGS", "Q25 USGS", "Q50 USGS", "Q100 USGS",
                "Q10 VDOT", "Q25 VDOT", "Q50 VDOT", "Q100 VDOT", "Governing Method",
                "Basin Data Source", "Rainfall Source", "Tc Length Source", "Regression Source"
            ]
            ws.append(headers)
            for c in range(1, len(headers) + 1):
                cell = ws.cell(row=ws.max_row, column=c)
                cell.fill = h_fill
                cell.font = h_font

            for item in batch_results:
                cid = item.get("culvert_id", "")
                ssr = item.get("ss_result")
                if not ssr:
                    ws.append([cid, "FAILED"])
                    continue
                vcomp = ssr.get("vdot_comparison", {})
                met = vcomp.get("metrics", {})
                r_map = {r["return_yr"]: r for r in vcomp.get("rows", [])}
                ws.append([
                    cid, "Success", round(met.get("drnarea_acres", 0), 1), round(met.get("slope_ft_mi", 0), 1),
                    vcomp.get("c_coeff", ""), vcomp.get("tc_min", ""),
                    r_map.get(10, {}).get("usgs_q_cfs", ""), r_map.get(25, {}).get("usgs_q_cfs", ""),
                    r_map.get(50, {}).get("usgs_q_cfs", ""), r_map.get(100, {}).get("usgs_q_cfs", ""),
                    r_map.get(10, {}).get("vdot_q_cfs", ""), r_map.get(25, {}).get("vdot_q_cfs", ""),
                    r_map.get(50, {}).get("vdot_q_cfs", ""), r_map.get(100, {}).get("vdot_q_cfs", ""),
                    vcomp.get("governing_method", ""),
                    vcomp.get("basin_source", ""), vcomp.get("rainfall_source", ""),
                    vcomp.get("tc_length_source", ""), vcomp.get("regression_source", ""),
                ])

            for col in ws.columns:
                max_l = max(len(str(c.value or "")) for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max(max_l + 3, 11), 38)
            _repair_cleared_contextlib()
            wb.save(xlsx_path)
        else:
            xlsx_path = None

        return csv_path, xlsx_path

    @staticmethod
    def _writable_path(path):
        """If the target is locked (e.g. open in Excel), fall back to a timestamped filename."""
        if not os.path.exists(path):
            return path
        try:
            with open(path, "a"):
                return path
        except PermissionError:
            stem, ext = os.path.splitext(path)
            return f"{stem}_{datetime.datetime.now():%Y%m%d_%H%M%S}{ext}"


# ---------------------------------------------------------------------------
# 10. Modern Themed GUI (Single, Batch, Results + KML / Shapefile GIS Export)
# ---------------------------------------------------------------------------
# Header logos (base64 PNG, embedded so the macro needs no extra files).
# USGS logo (USGS green PMS 348 #00843D, 227x86) - from USGS_Basemap_Attach.py
USGS_LOGO_PNG = (
    "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"
    "jYGM/wPMlyq7qm2ZIgAAAABJRU5ErkJggg=="
)
# VDOT logo (341x91) - from VA_LiDAR_Downloader.py
VDOT_LOGO_PNG = (
    "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"
)


class ModernTheme:
    """Flat navy/slate palette with a teal accent built on the ttk 'clam' engine (light & dark modes)."""
    PALETTES = {
        "light": dict(
            BG="#eef2f7", CARD="#ffffff", NAVY="#0f1f3d", SLATE="#1e293b", TEXT="#1f2937",
            MUTED="#64748b", BORDER="#d6dde8", ACCENT="#0d9488", ACCENT_HOVER="#0f766e",
            ACCENT_PRESS="#115e59", STRIPE="#f5f7fb", SELECT="#ccfbf1", SEL_FG="#0f1f3d",
            OK="#16a34a", WARN="#d97706", ERR="#dc2626", LOG_BG="#0b1220", LOG_FG="#cbd5e1",
            TITLE="#0f1f3d", FIELD="#ffffff", BTN2="#e2e8f0", BTN2_HOVER="#d5dde8",
            BTN2_PRESS="#cbd5e1", BTN2_DIS="#f1f5f9", BTN2_DIS_FG="#94a3b8", TAB="#dfe5ee",
            DANGER_BG="#fee2e2", DANGER_HOVER="#fecaca", HEADING="#0f1f3d", HEADING_ACTIVE="#1e293b",
        ),
        "dark": dict(
            BG="#0f172a", CARD="#1e293b", NAVY="#0f1f3d", SLATE="#334155", TEXT="#e2e8f0",
            MUTED="#94a3b8", BORDER="#334155", ACCENT="#14b8a6", ACCENT_HOVER="#0d9488",
            ACCENT_PRESS="#0f766e", STRIPE="#243247", SELECT="#134e4a", SEL_FG="#f1f5f9",
            OK="#4ade80", WARN="#fbbf24", ERR="#f87171", LOG_BG="#020617", LOG_FG="#cbd5e1",
            TITLE="#f1f5f9", FIELD="#0f172a", BTN2="#334155", BTN2_HOVER="#475569",
            BTN2_PRESS="#1e293b", BTN2_DIS="#1e293b", BTN2_DIS_FG="#64748b", TAB="#172033",
            DANGER_BG="#450a0a", DANGER_HOVER="#7f1d1d", HEADING="#020617", HEADING_ACTIVE="#1e293b",
        ),
    }
    MODE = "light"
    PREF_FILE = UI_PREF_FILE

    @classmethod
    def set_mode(cls, mode):
        cls.MODE = mode if mode in cls.PALETTES else "light"
        for k, v in cls.PALETTES[cls.MODE].items():
            setattr(cls, k, v)

    @classmethod
    def load_mode(cls):
        return _load_ui_prefs().get("theme", "light")

    @classmethod
    def save_mode(cls):
        _save_ui_prefs(theme=cls.MODE)

    FONT = ("Segoe UI", 10)
    FONT_B = ("Segoe UI Semibold", 10)
    FONT_SM = ("Segoe UI", 9)
    FONT_TITLE = ("Segoe UI Semibold", 16)
    FONT_CARD = ("Segoe UI Semibold", 11)
    MONO = ("Consolas", 9)

    @classmethod
    def apply(cls, root):
        root.configure(bg=cls.BG)
        st = ttk.Style(root)
        try:
            st.theme_use("clam")
        except tk.TclError:
            pass

        st.configure(".", background=cls.BG, foreground=cls.TEXT, font=cls.FONT,
                     bordercolor=cls.BORDER, focuscolor=cls.ACCENT)
        st.configure("TFrame", background=cls.BG)
        st.configure("Card.TFrame", background=cls.CARD, relief="solid", borderwidth=1, bordercolor=cls.BORDER)
        st.configure("Inner.TFrame", background=cls.CARD)
        st.configure("TLabel", background=cls.BG, foreground=cls.TEXT)
        st.configure("Card.TLabel", background=cls.CARD, foreground=cls.TEXT)
        st.configure("CardTitle.TLabel", background=cls.CARD, foreground=cls.TITLE, font=cls.FONT_CARD)
        st.configure("Muted.TLabel", background=cls.CARD, foreground=cls.MUTED, font=cls.FONT_SM)
        st.configure("Metric.TLabel", background=cls.CARD, foreground=cls.TITLE, font=("Segoe UI Semibold", 13))
        s = max(root.winfo_fpixels("1i") / 96.0, 1.0)
        st.configure("Card.TCheckbutton", background=cls.CARD, foreground=cls.TEXT,
                     indicatorbackground=cls.FIELD, indicatorforeground=cls.TEXT,
                     indicatorsize=int(round(16 * s)), indicatormargin=(0, 2, int(round(8 * s)), 2),
                     padding=(0, int(round(3 * s))))
        st.map("Card.TCheckbutton", background=[("active", cls.CARD)])
        st.configure("TRadiobutton", background=cls.CARD, foreground=cls.TEXT,
                     indicatorbackground=cls.FIELD, indicatorforeground=cls.TEXT)
        st.map("TRadiobutton", background=[("active", cls.CARD)])

        st.configure("TEntry", fieldbackground=cls.FIELD, foreground=cls.TEXT, insertcolor=cls.TEXT, padding=4,
                     bordercolor=cls.BORDER, lightcolor=cls.BORDER, darkcolor=cls.BORDER)
        st.map("TEntry", bordercolor=[("focus", cls.ACCENT)], lightcolor=[("focus", cls.ACCENT)],
               fieldbackground=[("disabled", cls.BTN2_DIS)],
               foreground=[("disabled", cls.BTN2_DIS_FG)])
        st.map("TSpinbox", fieldbackground=[("disabled", cls.BTN2_DIS)], foreground=[("disabled", cls.BTN2_DIS_FG)])
        st.configure("TSpinbox", fieldbackground=cls.FIELD, foreground=cls.TEXT, arrowcolor=cls.TEXT,
                     background=cls.BTN2, bordercolor=cls.BORDER, lightcolor=cls.BORDER, darkcolor=cls.BORDER)
        st.configure("TCombobox", padding=3, fieldbackground=cls.FIELD, foreground=cls.TEXT, arrowcolor=cls.TEXT,
                     background=cls.BTN2, bordercolor=cls.BORDER, lightcolor=cls.BORDER, darkcolor=cls.BORDER)
        st.map("TCombobox", fieldbackground=[("disabled", cls.BTN2_DIS), ("readonly", cls.FIELD)],
               foreground=[("disabled", cls.BTN2_DIS_FG), ("readonly", cls.TEXT)],
               selectbackground=[("readonly", cls.FIELD)], selectforeground=[("readonly", cls.TEXT)],
               background=[("active", cls.BTN2_HOVER), ("pressed", cls.BTN2_PRESS)],
               arrowcolor=[("disabled", cls.BTN2_DIS_FG)],
               bordercolor=[("focus", cls.ACCENT)])
        root.option_add("*TCombobox*Listbox.background", cls.FIELD)
        root.option_add("*TCombobox*Listbox.foreground", cls.TEXT)
        root.option_add("*TCombobox*Listbox.selectBackground", cls.ACCENT)
        root.option_add("*TCombobox*Listbox.selectForeground", "#ffffff")

        st.configure("Accent.TButton", background=cls.ACCENT, foreground="#ffffff",
                     font=cls.FONT_B, padding=(14, 7), borderwidth=0)
        st.map("Accent.TButton",
               background=[("disabled", cls.BTN2_DIS_FG), ("pressed", cls.ACCENT_PRESS), ("active", cls.ACCENT_HOVER)],
               foreground=[("disabled", cls.BTN2_DIS)])
        st.configure("Secondary.TButton", background=cls.BTN2, foreground=cls.TITLE,
                     padding=(12, 6), borderwidth=0)
        st.map("Secondary.TButton",
               background=[("disabled", cls.BTN2_DIS), ("pressed", cls.BTN2_PRESS), ("active", cls.BTN2_HOVER)],
               foreground=[("disabled", cls.BTN2_DIS_FG)])
        st.configure("Danger.TButton", background=cls.DANGER_BG, foreground=cls.ERR, padding=(12, 6), borderwidth=0)
        st.map("Danger.TButton", background=[("disabled", cls.BTN2_DIS), ("active", cls.DANGER_HOVER)],
               foreground=[("disabled", cls.BTN2_DIS_FG)])

        st.configure("TNotebook", background=cls.BG, borderwidth=0, tabmargins=(0, 4, 0, 0))
        st.configure("TNotebook.Tab", background=cls.TAB, foreground=cls.MUTED,
                     padding=(18, 8), font=cls.FONT_B, borderwidth=0)
        st.map("TNotebook.Tab", background=[("selected", cls.CARD)], foreground=[("selected", cls.TITLE)])

        st.configure("Treeview", background=cls.CARD, fieldbackground=cls.CARD, foreground=cls.TEXT,
                     rowheight=26, borderwidth=0)
        st.map("Treeview", background=[("selected", cls.SELECT)], foreground=[("selected", cls.SEL_FG)])
        st.configure("Treeview.Heading", background=cls.HEADING, foreground="#ffffff",
                     font=cls.FONT_B, relief="flat", padding=6)
        st.map("Treeview.Heading", background=[("active", cls.HEADING_ACTIVE)])

        st.configure("TScrollbar", background=cls.BTN2, troughcolor=cls.CARD, arrowcolor=cls.TEXT,
                     bordercolor=cls.BORDER, lightcolor=cls.BTN2, darkcolor=cls.BTN2)
        st.map("TScrollbar", background=[("active", cls.BTN2_HOVER)])
        st.configure("Accent.Horizontal.TProgressbar", troughcolor="#24324f", background=cls.ACCENT,
                     bordercolor=cls.NAVY, lightcolor=cls.ACCENT, darkcolor=cls.ACCENT)
        st.configure("Batch.Horizontal.TProgressbar", troughcolor=cls.BTN2, background=cls.ACCENT, thickness=18,
                     bordercolor=cls.BORDER, lightcolor=cls.ACCENT, darkcolor=cls.ACCENT)
        st.configure("BatchDone.Horizontal.TProgressbar", troughcolor=cls.BTN2, background=cls.OK, thickness=18,
                     bordercolor=cls.BORDER, lightcolor=cls.OK, darkcolor=cls.OK)
        st.configure("BatchWarn.Horizontal.TProgressbar", troughcolor=cls.BTN2, background=cls.WARN, thickness=18,
                     bordercolor=cls.BORDER, lightcolor=cls.WARN, darkcolor=cls.WARN)
        st.configure("TPanedwindow", background=cls.BG)
        st.configure("Sash", sashthickness=6, gripcount=0, background=cls.BG, lightcolor=cls.BG, bordercolor=cls.BG)
        return st


ModernTheme.set_mode("light")


class Tooltip:
    """Hover help: shows `text` in a small popup after the pointer rests on `widget`."""
    DELAY_MS = 500
    WRAP_PX = 320

    def __init__(self, widget, text):
        self.widget, self.text = widget, text
        self._after = None
        self._tip = None
        widget.bind("<Enter>", self._schedule, add="+")
        widget.bind("<Leave>", self._hide, add="+")
        widget.bind("<ButtonPress>", self._hide, add="+")
        widget.bind("<Destroy>", self._hide, add="+")

    def _schedule(self, _e=None):
        self._cancel()
        self._after = self.widget.after(self.DELAY_MS, self._show)

    def _cancel(self):
        if self._after:
            try:
                self.widget.after_cancel(self._after)
            except tk.TclError:
                pass
            self._after = None

    def _show(self):
        self._after = None
        if self._tip or not self.text:
            return
        try:
            x = self.widget.winfo_rootx() + 12
            y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
            self._tip = tw = tk.Toplevel(self.widget)
            tw.wm_overrideredirect(True)
            tw.attributes("-topmost", True)
            tk.Label(tw, text=self.text, justify="left", wraplength=self.WRAP_PX, bg="#1e293b", fg="#f1f5f9",
                     font=ModernTheme.FONT, padx=9, pady=6, bd=1, relief="solid").pack()
            tw.update_idletasks()
            sw = tw.winfo_screenwidth()
            if x + tw.winfo_reqwidth() > sw:
                x = max(0, sw - tw.winfo_reqwidth() - 4)
            tw.wm_geometry(f"+{x}+{y}")
        except tk.TclError:
            self._tip = None

    def _hide(self, _e=None):
        self._cancel()
        if self._tip:
            try:
                self._tip.destroy()
            except tk.TclError:
                pass
            self._tip = None


class StreamStatsFenceDialog:
    T = ModernTheme
    APP_TITLE = "USGS StreamStats  \u00b7  VDOT Hydrology"
    POLL_MS = 60

    def __init__(self, root):
        self.root = root
        self.root.title(f"USGS StreamStats & VDOT Hydrology \u2013 Fence Query   v{APP_VERSION} ({BUILD_DATE})   Build By: {BUILT_BY}")
        self.root.geometry("1220x880")
        self.root.minsize(940, 700)
        self.T.set_mode(self.T.load_mode())
        self.T.apply(root)
        self._trees = []
        self._scroll_canvases = []

        self._ui_queue = queue.Queue()
        self._main_thread = threading.current_thread()
        self._busy = False
        self._busy_widgets = []
        self._cancel_event = threading.Event()
        self._active_tool = None  # Keeps the DgnPrimitiveTool alive (avoid garbage collection)

        self.projector = DgnCoordinateProjector(fallback_epsg=2284)
        self.api_client = USGSWaterServicesClient(log_callback=self.log_message)
        self.dgn_mgr = MicroStationDgnManager(self.projector, log_callback=self.log_message)

        self.fence_wgs84_poly = []
        self.clipped_streams = []
        self._streams_drawn = False
        self.batch_queue = []          # [{"culvert_id", "type", "x", "y", "ss_result", "iid"}]
        self.batch_table_origin_uor = None
        self.raindrop_result = None

        self._build_ui()
        _trace(f"App version {APP_VERSION} (build {BUILD_DATE})")
        self.log_message(f"USGS StreamStats Fence Query v{APP_VERSION} (build {BUILD_DATE})")
        self.root.after(self.POLL_MS, self._poll_ui_queue)
        self.root.after(250, lambda: self.load_fence_and_clip_streams(draw_outline=False))
        self.root.after(900, self.check_services)

    # =====================================================================
    # Layout
    # =====================================================================
    def _build_ui(self):
        self._build_header()

        self._build_status_bar()

        body = ttk.Frame(self.root, padding=(14, 10, 14, 6))
        body.pack(fill="both", expand=True)

        paned = ttk.Panedwindow(body, orient="vertical")
        paned.pack(fill="both", expand=True)

        nb_holder = ttk.Frame(paned)
        self.nb = ttk.Notebook(nb_holder)
        self.nb.pack(fill="both", expand=True)
        self.tab_batch = ttk.Frame(self.nb, style="Inner.TFrame", padding=12)
        self.tab_results = ttk.Frame(self.nb, style="Inner.TFrame", padding=12)
        self.tab_flow = ttk.Frame(self.nb, style="Inner.TFrame", padding=12)
        self.tab_settings = ttk.Frame(self.nb, style="Inner.TFrame", padding=12)
        self.tab_libs = ttk.Frame(self.nb, style="Inner.TFrame", padding=12)
        self.nb.add(self.tab_flow, text="  \u2248  Flow (Raindrop) Path  ")
        self.nb.add(self.tab_batch, text="  \u2630  Culvert Analysis  ")
        self.nb.add(self.tab_results, text="  \u2261  Results & Comparison  ")
        self.nb.add(self.tab_settings, text="  \u2699  Settings  ")
        self.nb.add(self.tab_libs, text="  \u25a3  Libraries  ")
        self._build_settings_card(self._scrollable(self.tab_settings))
        self._build_libraries_card(self.tab_libs)
        self._build_batch_tab()
        self._build_results_tab()
        self._build_flow_tab()
        paned.add(nb_holder, weight=4)

        log_holder = ttk.Frame(paned)
        self._build_log_card(log_holder)
        paned.add(log_holder, weight=1)

    def _build_header(self):
        T = self.T
        hdr = tk.Frame(self.root, bg=T.NAVY)
        hdr.pack(fill="x")

        tk.Frame(hdr, bg=T.ACCENT, width=6).pack(side="left", fill="y")
        self.theme_btn = tk.Button(
            hdr, command=self.toggle_theme, bg="#13315c", fg="#e2e8f0", activebackground="#1e4479",
            activeforeground="#ffffff", relief="flat", bd=0, padx=12, pady=4, cursor="hand2", font=T.FONT_B,
        )
        self.theme_btn.pack(side="right", padx=(0, 16))
        Tooltip(self.theme_btn, "Switches between light and dark mode. Your choice is remembered.")
        self._update_theme_btn()
        self.delete_all_btn = tk.Button(
            hdr, text="\u2716  Delete All", command=self.delete_all_drawings, bg="#dc2626",
            fg="#ffffff", activebackground="#b91c1c", activeforeground="#ffffff", relief="flat", bd=0, padx=12,
            pady=4, cursor="hand2", font=T.FONT_B,
        )
        self.delete_all_btn.pack(side="right", padx=(0, 8))
        Tooltip(self.delete_all_btn, "Deletes everything this tool drew in the active model (drainage areas, "
                                     "markers, labels, streams, flowpaths, raindrop paths, tables and the fence "
                                     "outline) and clears the fence. Asks first. Use MicroStation Undo to restore.")
        logos = tk.Frame(hdr, bg=T.NAVY)
        logos.pack(side="left", padx=(14, 0), pady=10)
        self._logo_imgs = []
        self._make_logo(logos, "usgs").pack(side="left", padx=(0, 12))
        self._make_logo(logos, "vdot").pack(side="left")
        left = tk.Frame(hdr, bg=T.NAVY)
        left.pack(side="left", fill="both", expand=True, padx=16, pady=(12, 12))
        title_row = tk.Frame(left, bg=T.NAVY)
        title_row.pack(anchor="w")
        tk.Label(title_row, text=self.APP_TITLE, bg=T.NAVY, fg="#ffffff", font=T.FONT_TITLE).pack(side="left")
        tk.Label(
            left, text="Fence query  \u2192  NHDPlus streams  \u2192  StreamStats delineation  \u2192  VDOT Rational comparison",
            bg=T.NAVY, fg="#94a3b8", font=T.FONT_SM, anchor="w",
        ).pack(anchor="w", fill="x", pady=(2, 0))
        tk.Label(left, text=f"v{APP_VERSION}  \u00b7  Build {BUILD_DATE}  \u00b7  Build By: {BUILT_BY}",
                 bg=T.NAVY, fg="#5eead4", font=T.FONT_SM).pack(anchor="w", pady=(1, 0))

    def _build_env_pills(self, parent):
        """Environment status pills (shown in the status bar)."""
        T = self.T
        pills = tk.Frame(parent, bg=T.NAVY)
        env = [
            ("MicroStation" if IN_MICROSTATION else "Standalone", IN_MICROSTATION),
            ("DGN GCS", self.projector.has_gcs),
            ("pyproj", HAS_PYPROJ),
            ("shapely", HAS_SHAPELY),
            ("openpyxl", HAS_OPENPYXL),
        ]
        for text, ok in env:
            tk.Label(
                pills, text="\u25cf " + text,
                bg="#13315c" if ok else "#3b2330", fg="#5eead4" if ok else "#fda4af",
                font=T.FONT_SM, padx=7, pady=0,
            ).pack(side="left", padx=2)
        return pills

    LOGO_H = 42
    LOGO_W = 104

    @staticmethod
    def _logo_dir():
        try:
            base = os.path.dirname(os.path.abspath(__file__))
        except NameError:
            base = os.getcwd()
        return os.path.join(base, "logos")

    @staticmethod
    def _scale_photo(img, max_h, max_w):
        """Scale a PhotoImage to fit max_h x max_w using the best zoom/subsample ratio (Tk has no resize)."""
        h, w = img.height(), img.width()
        best = (1, max(1, math.ceil(h / max_h), math.ceil(w / max_w)))
        best_r = best[0] / best[1]
        for d in range(1, 13):
            for z in range(1, 13):
                r = z / d
                if r > best_r and h * r <= max_h and w * r <= max_w:
                    best, best_r = (z, d), r
        z, d = best
        if z > 1:
            img = img.zoom(z, z)
        if d > 1:
            img = img.subsample(d, d)
        return img

    def _make_logo(self, parent, which):
        """Header logo. Uses logos\\usgs_logo.png / logos\\vdot_logo.png next to this script when present,
        otherwise the embedded official artwork; falls back to a drawn badge if PNG loading fails."""
        T = self.T
        s = max(1.0, self.root.winfo_fpixels("1i") / 96.0)
        img = None
        path = os.path.join(self._logo_dir(), f"{which}_logo.png")
        try:
            if os.path.isfile(path):
                img = tk.PhotoImage(file=path)
            else:
                img = tk.PhotoImage(data="".join(USGS_LOGO_PNG if which == "usgs" else VDOT_LOGO_PNG))
        except tk.TclError:
            img = None
        if img is not None:
            img = self._scale_photo(img, int(round(self.LOGO_H * s)), int(round(2.2 * self.LOGO_W * s)))
            self._logo_imgs.append(img)
            return tk.Label(parent, image=img, bg=T.NAVY, bd=0)
        w, h = self.LOGO_W, self.LOGO_H
        cv = tk.Canvas(parent, width=int(w * s), height=int(h * s), bg=T.NAVY, highlightthickness=0, bd=0)

        def _px(n):
            return -max(1, int(round(n * s)))

        def _rrect(x1, y1, x2, y2, r, **kw):
            pts = [x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r, x2, y2 - r, x2, y2,
                   x2 - r, y2, x1 + r, y2, x1, y2, x1, y2 - r, x1, y1 + r, x1, y1]
            return cv.create_polygon([p * s for p in pts], smooth=True, **kw)

        _rrect(1, 1, w - 1, h - 1, 10, fill="#ffffff", outline="")
        if which == "usgs":
            green = "#00704a"
            for i, y0 in enumerate((14, 21, 28)):
                cv.create_line(*[c * s for c in (8, y0, 14, y0 - 4, 20, y0, 26, y0 + 4, 32, y0)], smooth=True,
                               fill=green, width=max(2, int(2.5 * s)), capstyle="round")
            cv.create_text(68 * s, 21 * s, text="USGS", fill="#1f2937", font=("Segoe UI Black", _px(22), "bold"))
            cv.create_text(w / 2 * s, 42 * s, text="science for a changing world", fill=green,
                           font=("Segoe UI Semibold", _px(7.5)))
        else:
            blue, orange = "#00447c", "#f2a900"
            cv.create_text(w / 2 * s, 20 * s, text="VDOT", fill=blue, font=("Segoe UI Black", _px(26), "bold"))
            cv.create_line(*[c * s for c in (14, 35, w / 2, 32, w - 14, 35)], smooth=True, fill=orange,
                           width=max(2, int(2.5 * s)), capstyle="round")
            cv.create_text(w / 2 * s, 44 * s, text="Virginia DOT", fill="#475569",
                           font=("Segoe UI Semibold", _px(9)))
        return cv

    def _update_theme_btn(self):
        self.theme_btn.configure(text="\u2600  Light Mode" if self.T.MODE == "dark" else "\u263e  Dark Mode")

    def toggle_theme(self):
        """Switches between light and dark mode live and remembers the choice."""
        self.T.set_mode("light" if self.T.MODE == "dark" else "dark")
        self.T.apply(self.root)
        self.T.save_mode()
        for tree in self._trees:
            try:
                tree.tag_configure("odd", background=self.T.STRIPE)
                tree.tag_configure("ok", foreground=self.T.OK)
                tree.tag_configure("warn", foreground=self.T.WARN)
                tree.tag_configure("err", foreground=self.T.ERR)
                tree.tag_configure("busy", foreground=self.T.ACCENT)
            except tk.TclError:
                pass
        self._recolor_combobox_popdowns(self.root)
        for canvas in self._scroll_canvases:
            try:
                canvas.configure(bg=self.T.CARD)
            except tk.TclError:
                pass
        self._update_theme_btn()

    def _recolor_combobox_popdowns(self, widget):
        for child in widget.winfo_children():
            if isinstance(child, ttk.Combobox):
                try:
                    pop = self.root.tk.call("ttk::combobox::PopdownWindow", child)
                    self.root.tk.call(f"{pop}.f.l", "configure", "-background", self.T.FIELD,
                                      "-foreground", self.T.TEXT)
                except tk.TclError:
                    pass
            elif not isinstance(child, tk.Toplevel):
                self._recolor_combobox_popdowns(child)

    def _scrollable(self, parent):
        """Returns a frame inside `parent` that scrolls vertically (mouse wheel or scrollbar) when its
        content is taller than the tab."""
        canvas = tk.Canvas(parent, highlightthickness=0, borderwidth=0, bg=self.T.CARD)
        vbar = ttk.Scrollbar(parent, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vbar.set)
        vbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        body = ttk.Frame(canvas, style="Inner.TFrame")
        win = canvas.create_window((0, 0), window=body, anchor="nw")
        body.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(win, width=e.width))

        def _wheel(e):
            try:
                under = self.root.winfo_containing(e.x_root, e.y_root)
                if (under is not None and str(under).startswith(str(canvas))
                        and body.winfo_reqheight() > canvas.winfo_height()):
                    canvas.yview_scroll(-1 if e.delta > 0 else 1, "units")
            except (tk.TclError, KeyError):
                pass

        self.root.bind_all("<MouseWheel>", _wheel, add="+")
        self._scroll_canvases.append(canvas)
        return body

    def _card(self, parent, title, subtitle=None):
        card = ttk.Frame(parent, style="Card.TFrame", padding=(14, 10))
        head = ttk.Frame(card, style="Inner.TFrame")
        head.pack(fill="x", pady=(0, 8))
        ttk.Label(head, text=title, style="CardTitle.TLabel").pack(side="left")
        if subtitle:
            ttk.Label(head, text=subtitle, style="Muted.TLabel").pack(side="left", padx=10, pady=(3, 0))
        inner = ttk.Frame(card, style="Inner.TFrame")
        inner.pack(fill="both", expand=True)
        return card, inner, head

    def _field(self, parent, row, col, label, var, width=10, widget="entry", values=None):
        ttk.Label(parent, text=label, style="Muted.TLabel").grid(row=row, column=col, sticky="w", padx=(0, 4))
        if widget == "combo":
            w = ttk.Combobox(parent, textvariable=var, values=values or [], width=width)
        else:
            w = ttk.Entry(parent, textvariable=var, width=width)
        w.grid(row=row + 1, column=col, sticky="we", padx=(0, 14), pady=(1, 0))
        return w

    def _build_settings_card(self, parent):
        card, inner, _ = self._card(parent, "Project Settings", "Coordinates, markers & GIS output")
        card.pack(fill="x", anchor="n")

        self.rcode_var = tk.StringVar(value="VA")       # StreamStats state (fixed; no longer shown)
        self.epsg_var = tk.StringVar(value="2284")      # fallback EPSG if the DGN has no GCS (no longer shown)
        self.c_override_var = tk.StringVar(value="")
        self.auto_gis_var = tk.BooleanVar(value=False)
        self.auto_markers_var = tk.BooleanVar(value=True)
        self.lfp_var = tk.BooleanVar(value=True)

        self._field(inner, 0, 0, "Rational C override", self.c_override_var, width=8)

        prefs = _load_ui_prefs()
        try:
            scale = float(prefs.get("drawing_scale", 100.0))
            scale = scale if scale in MicroStationDgnManager.DRAWING_SCALES else 100.0
        except (TypeError, ValueError):
            scale = 100.0
        try:
            plot_in = float(prefs.get("plot_text_in", 0.20))
        except (TypeError, ValueError):
            plot_in = 0.20
        self._scale_labels = {f"1\" = {int(s)}'": s for s in MicroStationDgnManager.DRAWING_SCALES}
        self.scale_var = tk.StringVar(value=f"1\" = {int(scale)}'")
        self.plot_text_var = tk.StringVar(value=f"{plot_in:.2f}")
        self.text_ft_var = tk.StringVar()
        cb = self._field(inner, 2, 0, "Drawing scale", self.scale_var, width=10, widget="combo",
                         values=list(self._scale_labels))
        cb.configure(state="readonly")
        self._field(inner, 2, 1, "Text height (in, plotted)", self.plot_text_var, width=8)
        ttk.Label(inner, textvariable=self.text_ft_var, style="Muted.TLabel").grid(
            row=3, column=2, sticky="w")
        self.scale_var.trace_add("write", lambda *_: self._apply_text_settings())
        self.plot_text_var.trace_add("write", lambda *_: self._apply_text_settings())
        self._apply_text_settings(save=False)

        style_btn = ttk.Button(inner, text="KML / KMZ Style\u2026", style="Secondary.TButton",
                               command=self.open_kml_style_dialog)
        style_btn.grid(row=4, column=0, sticky="w", pady=(10, 0))
        Tooltip(style_btn, "Sets the colors, line widths and style-file location used for every KML/KMZ export.")
        inner.columnconfigure(5, weight=1)
        inner.rowconfigure(2, pad=14)

        for r, (txt, var) in enumerate((
            ("Auto-save KML & Shapefiles beside the active DGN after each run", self.auto_gis_var),
            ("Draw pour-point & culvert-inlet markers after each delineation", self.auto_markers_var),
            ("Compute longest flowpath (drawn in DGN, exported as a separate _LongestFlowpath shapefile)", self.lfp_var),
        )):
            ttk.Checkbutton(inner, text=txt, variable=var, style="Card.TCheckbutton").grid(
                row=5 + r, column=0, columnspan=6, sticky="w", pady=(10 if r == 0 else 2, 0))
        self._build_methods_card(parent)

    def _build_methods_card(self, parent):
        """Peak-discharge methods from VDOT Drainage Manual Table 6-2: a checkbox for each one this app can
        calculate, plus the inputs those methods need. Saved with the other UI preferences."""
        V = VDOTHydrologyCalculator
        card, inner, head = self._card(parent, "Peak Discharge Methods",
                                       "VDOT Drainage Manual Chapter 6, Table 6-2 (Rev. 7/26)")
        card.pack(fill="x", anchor="n", pady=(10, 0))
        rbtn = ttk.Button(head, text="Reset to VDOT Defaults", style="Secondary.TButton",
                          command=self._reset_method_settings)
        rbtn.pack(side="right")
        Tooltip(rbtn, "Restores the method choices the VDOT Drainage Manual recommends: Rational up to 200 ac, "
                      "NRCS up to 2,000 ac, USGS regression above 200 ac, Anderson off.")

        opts = V.default_options()
        saved = _load_ui_prefs().get("hydro_methods") or {}
        if isinstance(saved, dict):
            opts["enabled"].update({k: bool(v) for k, v in (saved.get("enabled") or {}).items()
                                    if k in opts["enabled"]})
            for k in ("use_ranges", "design", "cn", "hsg", "anderson_type"):
                if k in saved:
                    opts[k] = saved[k]

        for c, txt in enumerate(("Use", "Method", "VDOT size range", "Notes")):
            ttk.Label(inner, text=txt.upper(), style="Muted.TLabel").grid(row=0, column=c, sticky="w",
                                                                         padx=(0, 14), pady=(0, 2))
        self._method_vars = {}
        for r, (key, label, can_calc, _default, lo, hi, note) in enumerate(V.METHODS, start=1):
            if can_calc:
                var = tk.BooleanVar(value=bool(opts["enabled"].get(key)))
                ttk.Checkbutton(inner, variable=var, style="Card.TCheckbutton",
                                command=self._apply_method_settings).grid(row=r, column=0, sticky="w")
                self._method_vars[key] = var
            else:
                ttk.Label(inner, text="\u2014", style="Muted.TLabel").grid(row=r, column=0, sticky="w", padx=(4, 0))
            ttk.Label(inner, text=label, style="Card.TLabel" if can_calc else "Muted.TLabel").grid(
                row=r, column=1, sticky="w", padx=(0, 14))
            ttk.Label(inner, text=V.range_text(lo, hi), style="Muted.TLabel").grid(row=r, column=2, sticky="w",
                                                                                   padx=(0, 14))
            ttk.Label(inner, text=note, style="Muted.TLabel").grid(row=r, column=3, sticky="w")
        inner.columnconfigure(3, weight=1)

        opt = ttk.Frame(inner, style="Inner.TFrame")
        opt.grid(row=len(V.METHODS) + 1, column=0, columnspan=4, sticky="we", pady=(10, 0))
        self._design_labels = {"VDOT default (by drainage area)": "auto"}
        self._design_labels.update({m[1]: m[0] for m in V.METHODS if m[2]})
        self._anderson_labels = {label: key for key, label in V.ANDERSON_TYPES}
        self._hsg_labels = {"Auto (USDA SSURGO)": "auto", "A": "A", "B": "B", "C": "C", "D": "D"}

        def _label_for(mapping, value, fallback):
            return next((k for k, v in mapping.items() if v == value), fallback)

        self.design_var = tk.StringVar(value=_label_for(self._design_labels, opts.get("design"),
                                                        "VDOT default (by drainage area)"))
        self.cn_var = tk.StringVar(value="" if not opts.get("cn") else f"{float(opts['cn']):g}")
        self.hsg_var = tk.StringVar(value=_label_for(self._hsg_labels, opts.get("hsg"), "Auto (USDA SSURGO)"))
        self.anderson_var = tk.StringVar(value=_label_for(self._anderson_labels, opts.get("anderson_type"),
                                                          V.ANDERSON_TYPES[0][1]))
        self.use_ranges_var = tk.BooleanVar(value=bool(opts.get("use_ranges", True)))
        self.cn_note_var = tk.StringVar()

        w = self._field(opt, 0, 0, "Design flow method", self.design_var, width=34, widget="combo",
                        values=list(self._design_labels))
        w.configure(state="readonly")
        Tooltip(w, "Which method supplies the VDOT Design flow. The VDOT default is Rational up to 200 ac and "
                   "USGS regression above that (urban equations when the basin is 30% or more developed).")
        w = self._field(opt, 0, 1, "NRCS curve number (blank = estimate)", self.cn_var, width=10)
        Tooltip(w, "Runoff curve number (30-98) used for every culvert. Leave blank to estimate it for each "
                   "basin from NLCD 2021 land cover and the soil group at right.")
        w = self._field(opt, 0, 2, "Soil group for estimated CN", self.hsg_var, width=20, widget="combo",
                        values=list(self._hsg_labels))
        w.configure(state="readonly")
        Tooltip(w, "Hydrologic soil group used when the curve number is estimated. Auto reads the soil groups "
                   "inside each basin from the USDA SSURGO soil survey.")
        w = self._field(opt, 0, 3, "Anderson watershed type", self.anderson_var, width=32, widget="combo",
                        values=list(self._anderson_labels))
        w.configure(state="readonly")
        Tooltip(w, "Selects the Anderson lag-time equation (VDOT Table 6-5).")
        ttk.Label(opt, textvariable=self.cn_note_var, style="Muted.TLabel").grid(row=2, column=1, columnspan=3,
                                                                                 sticky="w")
        ttk.Checkbutton(opt, text="Only calculate each method inside its VDOT Table 6-2 size range",
                        variable=self.use_ranges_var, style="Card.TCheckbutton",
                        command=self._apply_method_settings).grid(row=3, column=0, columnspan=4, sticky="w",
                                                                  pady=(8, 0))
        for var in (self.design_var, self.cn_var, self.hsg_var, self.anderson_var):
            var.trace_add("write", lambda *_: self._apply_method_settings())
        self._apply_method_settings(save=False)

    def _apply_method_settings(self, save=True):
        """Pushes the Settings-tab method choices to the calculator (used by the next delineation)."""
        V = VDOTHydrologyCalculator
        cn, txt = None, self.cn_var.get().strip()
        if txt:
            try:
                cn = float(txt)
                if not 30.0 <= cn <= 98.0:
                    raise ValueError
                self.cn_note_var.set(f"CN {cn:g} is used for every culvert.")
            except ValueError:
                cn = None
                self.cn_note_var.set("Enter a curve number from 30 to 98, or leave blank to estimate it.")
        else:
            self.cn_note_var.set("CN is estimated for each basin from NLCD 2021 land cover and the soil group.")
        opts = {
            "enabled": {k: bool(v.get()) for k, v in self._method_vars.items()},
            "use_ranges": bool(self.use_ranges_var.get()),
            "design": self._design_labels.get(self.design_var.get(), "auto"),
            "cn": cn, "hsg": self._hsg_labels.get(self.hsg_var.get(), "auto"),
            "anderson_type": self._anderson_labels.get(self.anderson_var.get(), "natural"),
        }
        V.options = opts
        if save:
            _save_ui_prefs(hydro_methods=opts)

    def _reset_method_settings(self):
        V = VDOTHydrologyCalculator
        d = V.default_options()
        for k, var in self._method_vars.items():
            var.set(d["enabled"][k])
        self.use_ranges_var.set(True)
        self.design_var.set("VDOT default (by drainage area)")
        self.cn_var.set("")
        self.hsg_var.set("Auto (USDA SSURGO)")
        self.anderson_var.set(V.ANDERSON_TYPES[0][1])
        self._apply_method_settings()
        self.log_message("Peak-discharge methods reset to the VDOT Drainage Manual defaults.", "info")

    def _apply_text_settings(self, save=True):
        """Drawing scale x plotted text height -> DGN text height for stream labels, marker labels & tables."""
        scale = self._scale_labels.get(self.scale_var.get(), 100.0)
        try:
            plot_in = float(self.plot_text_var.get())
            if not 0.01 <= plot_in <= 2.0:
                raise ValueError
        except ValueError:
            self.text_ft_var.set("Enter 0.01 \u2013 2.00 in")
            return
        self.dgn_mgr.drawing_scale = scale
        self.dgn_mgr.plot_text_in = plot_in
        self.text_ft_var.set(f"= {self.dgn_mgr.text_height_ft():.1f} ft in DGN")
        if save:
            _save_ui_prefs(drawing_scale=scale, plot_text_in=plot_in)

    def _build_libraries_card(self, parent):
        card, inner, head = self._card(parent, "Libraries", "Python packages required by this tool")
        card.pack(fill="both", expand=True, anchor="n")
        copy_btn = ttk.Button(head, text="Copy", style="Secondary.TButton", command=self._copy_library_status)
        copy_btn.pack(side="right")
        Tooltip(copy_btn, "Copies the library status list to the clipboard.")
        ref_btn = ttk.Button(head, text="Refresh", style="Secondary.TButton", command=self._refresh_library_status)
        ref_btn.pack(side="right", padx=(0, 6))
        Tooltip(ref_btn, "Re-checks which Python packages are installed and their versions.")
        frame, self.lib_tree = self._make_tree(inner, [
            ("lib", "Library", 160, "w"),
            ("status", "Status", 135, "w"),
            ("ver", "Version", 80, "w"),
            ("purpose", "Used for", 225, "w"),
            ("path", "Path", 380, "w"),
        ], height=6)
        self.lib_path_var = tk.StringVar(value="Select a library to see its full path")
        ttk.Entry(inner, textvariable=self.lib_path_var, state="readonly").pack(side="bottom", fill="x", pady=(8, 0))
        frame.pack(fill="both", expand=True)
        self.lib_tree.bind("<<TreeviewSelect>>", self._on_library_select)
        self._refresh_library_status()

    def _on_library_select(self, _event=None):
        sel = self.lib_tree.selection()
        if sel:
            vals = self.lib_tree.item(sel[0], "values")
            self.lib_path_var.set(f"{vals[0]}:  {vals[4]}")

    def _refresh_library_status(self):
        tree = self.lib_tree
        tree.delete(*tree.get_children())
        for i, (name, purpose, ok, ver, path) in enumerate(library_status()):
            tags = ("ok" if ok else "err",) + (("odd",) if i % 2 else ())
            tree.insert("", "end", values=(name, "\u2714 Loaded" if ok else "\u2716 Not loaded",
                                           ver or ("\u2013" if ok else ""), purpose, path), tags=tags)

    def _copy_library_status(self):
        lines = [f"{n}\t{'Loaded' if ok else 'NOT LOADED'}\t{v}\t{p}" for n, _, ok, v, p in library_status()]
        self.root.clipboard_clear()
        self.root.clipboard_append("\n".join(lines))
        if hasattr(self, "status_var"):
            self.status_var.set("Library status copied to clipboard")

    def _make_tree(self, parent, columns, height=8):
        """columns: list of (key, heading, width, anchor)."""
        frame = ttk.Frame(parent, style="Inner.TFrame")
        tree = ttk.Treeview(frame, columns=[c[0] for c in columns], show="headings", height=height)
        for key, head, width, anchor in columns:
            tree.heading(key, text=head)
            tree.column(key, width=width, anchor=anchor, stretch=True)
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        tree.tag_configure("odd", background=self.T.STRIPE)
        tree.tag_configure("ok", foreground=self.T.OK)
        tree.tag_configure("warn", foreground=self.T.WARN)
        tree.tag_configure("err", foreground=self.T.ERR)
        tree.tag_configure("busy", foreground=self.T.ACCENT)
        self._trees.append(tree)
        return frame, tree

    def _button(self, parent, text, cmd, style="Secondary.TButton", side="left", busy_lock=True, tip=None, **pack):
        b = ttk.Button(parent, text=text, command=cmd, style=style)
        b.pack(side=side, padx=pack.pop("padx", 4), **pack)
        if busy_lock:
            self._busy_widgets.append(b)
        if tip:
            Tooltip(b, tip)
        return b

    def _build_batch_tab(self):
        top = ttk.Frame(self.tab_batch, style="Inner.TFrame")
        top.pack(fill="x", pady=(0, 8))
        self._button(top, "\u21bb  Reload Fence & Clip Streams", self.load_fence_and_clip_streams, padx=(0, 8),
                     tip="Re-reads the active MicroStation fence and downloads the NHD streams clipped to it.")
        self._button(top, "\u25ad  Place Fence in DGN", self.place_fence_click, style="Accent.TButton", padx=(0, 4),
                     tip="Place a fence in the DGN (Shape or Block, chosen at right). The outline shows in red "
                         "while you draw it and is then drawn in red on level USGS_Fence_Boundary; "
                         "streams reload automatically when the fence is placed.")
        self.fence_mode_var = tk.StringVar(value="Shape")
        ttk.Combobox(top, textvariable=self.fence_mode_var, values=["Shape", "Block"], state="readonly",
                     width=7).pack(side="left", padx=(0, 10))
        self.fence_status_lbl = ttk.Label(top, text="Fence: initializing\u2026", style="Muted.TLabel")
        self.fence_status_lbl.pack(side="left")

        bar = ttk.Frame(self.tab_batch, style="Inner.TFrame")
        bar.pack(fill="x", pady=(0, 8))
        self._button(bar, "Import CSV\u2026", self.batch_import_csv, padx=(0, 6),
                     tip="Loads culvert points from a CSV file (ID, X/Lon, Y/Lat) into the batch queue. "
                         "Accepts lat/lon or DGN master-unit coordinates.")
        self._button(bar, "Queue All Fence Streams", self.batch_queue_all_fence_streams, padx=(0, 6),
                     tip="Adds the outlet point of every NHD stream clipped by the fence to the batch queue.")
        self._button(bar, "\u25ce  Place Markers in DGN\u2026", self.batch_place_markers_click,
                     style="Accent.TButton", padx=(0, 6),
                     tip="Click culvert locations in the DGN. Each click places pour-point and culvert-inlet "
                         "markers and adds the culvert to the queue.")
        self._button(bar, "Remove Selected", self.batch_remove_selected, padx=(0, 6),
                     tip="Removes the selected row(s) from the batch queue.")
        self._button(bar, "Clear Queue", self.batch_clear_queue, style="Danger.TButton", side="right", padx=0,
                     tip="Removes every culvert from the batch queue.")

        cols = [
            ("cid", "Culvert ID", 150, "w"), ("coord_type", "Coordinate Type", 120, "center"),
            ("x_lon", "X / Lon", 130, "e"), ("y_lat", "Y / Lat", 130, "e"), ("status", "Status", 260, "w"),
        ]
        tf, self.tree_batch = self._make_tree(self.tab_batch, cols, height=9)
        note = ttk.Label(
            self.tab_batch, style="Muted.TLabel",
            text="CSV format: ID, X/Lon, Y/Lat  \u2014  header row optional; lat,lon order is auto-detected for US data; "
                 "values outside lat/lon range are treated as DGN master units.",
        )

        run = ttk.Frame(self.tab_batch, style="Inner.TFrame")
        self._button(run, "\u25b6  Run Batch", self.run_batch_workflow, style="Accent.TButton", padx=(0, 6),
                     tip="Delineates every queued culvert with USGS StreamStats, computes VDOT Rational flows, "
                         "and draws the drainage areas, markers and summary table in the DGN.")
        self.stop_btn = ttk.Button(run, text="\u25a0  Stop", style="Danger.TButton", command=self.batch_stop,
                                   state="disabled")
        self.stop_btn.pack(side="left", padx=(0, 6))
        Tooltip(self.stop_btn, "Stops the running batch after the current culvert finishes.")
        self._button(run, "\u21bb  Retry Failed", lambda: self.run_batch_workflow(failed_only=True), padx=(0, 6),
                     tip="Re-runs only the culverts that failed or were skipped. Completed culverts keep their "
                         "results, and the summary table is redrawn in place.")
        self._button(run, "\u2316  Pick Table Origin", self.pick_batch_table_origin, padx=(0, 8),
                     tip="Click a point in the DGN to set where the batch results table is placed "
                         "(otherwise it is placed automatically).")
        self.batch_origin_lbl = ttk.Label(run, text="Table origin: auto",
                                          style="Muted.TLabel")
        self.batch_origin_lbl.pack(side="left")
        self._button(run, "Export GIS\u2026", lambda: self.export_gis_dialog("batch"), side="right", padx=0,
                     tip="Exports the delineated watersheds, pour points, clipped streams and longest flowpaths "
                         "to KML/KMZ and/or Shapefile.")
        self._button(run, "\u25ce  Draw Areas & Markers", lambda: self.draw_markers_for("batch"), side="right",
                     padx=(0, 6), tip="Draws (or redraws) the drainage areas plus pour-point and culvert-inlet "
                                      "markers for every delineated culvert.")

        prog = ttk.Frame(self.tab_batch, style="Inner.TFrame")
        head = ttk.Frame(prog, style="Inner.TFrame")
        head.pack(fill="x")
        self.batch_prog_lbl = ttk.Label(head, text="Batch progress: idle", style="Muted.TLabel")
        self.batch_prog_lbl.pack(side="left")
        self.batch_prog_pct = ttk.Label(head, text="0%", style="Muted.TLabel")
        self.batch_prog_pct.pack(side="right")
        self.batch_prog = ttk.Progressbar(prog, style="Batch.Horizontal.TProgressbar", mode="determinate",
                                          maximum=100.0, value=0.0)
        self.batch_prog.pack(fill="x", pady=(4, 0))
        self._bp = None
        # Bottom rows are packed before the table so the table shrinks first and they never get clipped.
        prog.pack(side="bottom", fill="x", pady=(12, 0))
        run.pack(side="bottom", fill="x", pady=(10, 0))
        note.pack(side="bottom", anchor="w", pady=(6, 0))
        tf.pack(fill="both", expand=True)

    # ---- Culvert Analysis progress bar -------------------------------------------------------
    BATCH_ITEM_GUESS_S = 20.0   # expected seconds per culvert until a real average is known

    @staticmethod
    def _fmt_secs(s):
        s = int(max(0, s))
        return f"{s // 60}:{s % 60:02d}"

    def _batch_progress_start(self, total):
        self._bp = {"total": total, "done": 0, "ok": 0, "fail": 0, "t0": time.time(),
                    "cur": None, "cur_t0": None}
        self.batch_prog.configure(style="Batch.Horizontal.TProgressbar", value=0.0)
        self._batch_progress_tick()

    def _batch_progress_item_started(self, idx, cid):
        if self._bp:
            self._bp.update(cur=f"{cid} ({idx + 1}/{self._bp['total']})", cur_t0=time.time())

    def _batch_progress_item_done(self, ok):
        if self._bp:
            self._bp["done"] += 1
            self._bp["ok" if ok else "fail"] += 1
            self._bp.update(cur=None, cur_t0=None)

    def _batch_progress_tick(self):
        """Updates the bar ~4x a second; the current culvert creeps toward its expected finish so the bar keeps moving."""
        bp = self._bp
        if not bp or bp.get("final"):
            return
        now = time.time()
        total, done = bp["total"], bp["done"]
        avg = (now - bp["t0"]) / done if done else self.BATCH_ITEM_GUESS_S
        frac = 0.0
        if bp["cur_t0"]:
            frac = min(0.92, (now - bp["cur_t0"]) / max(avg, 1.0))
        value = min(total, done + frac)
        pct = 100.0 * value / total if total else 0.0
        remaining = max(0.0, (total - value) * avg)
        step = f"Delineating {bp['cur']}" if bp["cur"] else "Waiting for next culvert\u2026"
        fails = f"  \u00b7  {bp['fail']} failed" if bp["fail"] else ""
        try:
            self.batch_prog.configure(value=pct)
            self.batch_prog_pct.configure(text=f"{pct:.0f}%")
            self.batch_prog_lbl.configure(
                text=f"{done}/{total} done{fails}  \u00b7  {step}  \u00b7  elapsed {self._fmt_secs(now - bp['t0'])}"
                     f"  \u00b7  ~{self._fmt_secs(remaining)} left")
            self.root.after(250, self._batch_progress_tick)
        except tk.TclError:
            pass

    def _batch_progress_finish(self, cancelled, text=None):
        bp = self._bp
        if not bp:
            return
        bp["final"] = True
        total, done = bp["total"], bp["done"]
        elapsed = self._fmt_secs(time.time() - bp["t0"])
        if text is None:
            if cancelled:
                text = f"Stopped \u2013 {done}/{total} processed ({bp['ok']} OK, {bp['fail']} failed) in {elapsed}"
            else:
                text = f"Complete \u2013 {bp['ok']}/{total} delineated" + \
                       (f", {bp['fail']} failed" if bp["fail"] else "") + f" in {elapsed}"
        style = "BatchDone.Horizontal.TProgressbar" if not cancelled and not bp["fail"] \
            else "BatchWarn.Horizontal.TProgressbar"
        pct = 100.0 * done / total if total else 0.0
        self.batch_prog.configure(style=style, value=pct)
        self.batch_prog_pct.configure(text=f"{pct:.0f}%")
        self.batch_prog_lbl.configure(text=text)

    def _build_results_tab(self):
        top = ttk.Frame(self.tab_results, style="Inner.TFrame")
        top.pack(fill="x", pady=(0, 10))
        ttk.Label(top, text="Culvert", style="Muted.TLabel").pack(side="left", padx=(0, 6))
        self.result_sel_var = tk.StringVar()
        self.result_combo = ttk.Combobox(top, textvariable=self.result_sel_var, state="readonly", width=32)
        self.result_combo.pack(side="left")
        self.result_combo.bind("<<ComboboxSelected>>", lambda _e: self._show_result(self.result_sel_var.get()))
        xbtn = ttk.Button(top, text="\u2193  Export to Excel", style="Accent.TButton",
                          command=self.export_results_excel)
        xbtn.pack(side="right")
        Tooltip(xbtn, "Saves all delineation results (basin characteristics, StreamStats peak flows and "
                      "VDOT Rational comparison) to an Excel workbook.")
        metrics = ttk.Frame(self.tab_results, style="Inner.TFrame")
        metrics.pack(fill="x", pady=(0, 10))
        self.metric_vars = {}
        for i, (key, label) in enumerate((
            ("area", "Drainage Area"), ("slope", "Basin Slope"), ("c", "Rational C"),
            ("tc", "Time of Conc."), ("gov", "Governing Method"),
        )):
            tile = ttk.Frame(metrics, style="Card.TFrame", padding=(12, 8))
            tile.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 8, 0))
            metrics.columnconfigure(i, weight=1 if key != "gov" else 2)
            ttk.Label(tile, text=label.upper(), style="Muted.TLabel").pack(anchor="w")
            v = tk.StringVar(value="\u2014")
            ttk.Label(tile, textvariable=v, style="Metric.TLabel").pack(anchor="w")
            self.metric_vars[key] = v

        cols = [
            ("event", "Event", 110, "w"), ("rational", "Rational (cfs)", 110, "e"),
            ("nrcs", "NRCS TR-55 (cfs)", 120, "e"), ("anderson", "Anderson (cfs)", 110, "e"),
            ("usgs", "USGS Rural (cfs)", 120, "e"), ("urban", "USGS Urban (cfs)", 120, "e"),
            ("vdot", "VDOT Design (cfs)", 130, "e"), ("diff", "Design vs USGS", 110, "center"),
        ]
        tf, self.tree_results = self._make_tree(self.tab_results, cols, height=8)
        tf.pack(fill="both", expand=True)
        self.result_src_var = tk.StringVar(
            value="Data sources for the selected culvert are listed here after a delineation.")
        ttk.Label(self.tab_results, style="Muted.TLabel", textvariable=self.result_src_var, wraplength=1100,
                  justify="left").pack(anchor="w", pady=(6, 0))

    def _build_flow_tab(self):
        tab = self.tab_flow
        ttk.Label(tab, style="Muted.TLabel",
                  text="Trace the path a raindrop follows from a point: overland to the nearest NHDPlus flowline, "
                       "then downstream along the mainstem (USGS NLDI).").pack(anchor="w", pady=(0, 8))

        form = ttk.Frame(tab, style="Inner.TFrame")
        form.pack(fill="x")
        self.rd_id_var = tk.StringVar(value="RD-01")
        self.rd_lon_var = tk.StringVar()
        self.rd_lat_var = tk.StringVar()
        self.rd_dist_var = tk.StringVar(value="1")
        self.rd_autodraw_var = tk.BooleanVar(value=True)
        self._field(form, 0, 0, "Trace ID", self.rd_id_var, width=12)
        self._field(form, 0, 1, "Longitude (WGS84)", self.rd_lon_var, width=14)
        self._field(form, 0, 2, "Latitude (WGS84)", self.rd_lat_var, width=14)
        self._field(form, 0, 3, "Max downstream (mi)", self.rd_dist_var, width=8)
        pick = ttk.Button(form, text="\u2316  Pick Point in DGN", style="Secondary.TButton",
                          command=self.raindrop_pick_point)
        pick.grid(row=1, column=4, sticky="w")
        self._busy_widgets.append(pick)
        Tooltip(pick, "Click the raindrop start point in the DGN. The point is marked and you are asked "
                      "whether to trace the flow path now.")
        ttk.Checkbutton(form, text="Draw in DGN automatically after tracing", variable=self.rd_autodraw_var,
                        style="Card.TCheckbutton").grid(row=2, column=0, columnspan=5, sticky="w", pady=(8, 0))

        act = ttk.Frame(tab, style="Inner.TFrame")
        act.pack(fill="x", pady=(10, 8))
        self._button(act, "\u25b6  Trace Flow Path", self.run_raindrop_trace, style="Accent.TButton", padx=(0, 8),
                     tip="Traces where a raindrop at the start point flows: overland to the nearest stream (NLDI), "
                         "then downstream up to the max distance.")
        self._button(act, "Draw in DGN", self.raindrop_draw, padx=(0, 8),
                     tip="Draws the last trace in the DGN with flow arrows, start/end markers and the "
                         "RaindropPath Item Type.")
        self._button(act, "Export KML / Shapefile\u2026", self.raindrop_export, padx=(0, 8),
                     tip="Saves the last trace as KML/KMZ and/or Shapefile.")
        self._button(act, "Clear", self.raindrop_clear, style="Danger.TButton", side="right", padx=0,
                     tip="Clears the trace results table. Elements already drawn in the DGN are not deleted.")

        self.rd_summary_var = tk.StringVar(value="No trace yet.")
        ttk.Label(tab, textvariable=self.rd_summary_var, style="CardTitle.TLabel").pack(anchor="w", pady=(0, 6))
        cols = [
            ("seq", "#", 40, "center"), ("seg", "Segment", 100, "center"), ("comid", "NHDPlus COMID", 130, "center"),
            ("len", "Length (ft)", 110, "e"), ("cum", "Cumulative (mi)", 120, "e"), ("method", "Method", 330, "w"),
        ]
        tf, self.tree_flow = self._make_tree(tab, cols, height=8)
        tf.pack(fill="both", expand=True)

    def raindrop_pick_point(self):
        def _set(x, y, z):
            lon, lat = self.projector.uor_to_wgs84(x, y, z)
            self.rd_lon_var.set(f"{lon:.6f}")
            self.rd_lat_var.set(f"{lat:.6f}")
            self.log_message(f"Raindrop start captured at {lon:.6f}, {lat:.6f} {self._fmt_master(x, y)}.", "success")
            tid = self.rd_id_var.get().strip() or "RD-01"
            try:
                self.dgn_mgr.draw_labeled_point(x, y, z, f"{tid} START  {lat:.6f}, {lon:.6f}")
                self._rd_start_drawn = (round(lon, 6), round(lat, 6))
            except Exception as exc:
                self.log_message(f"[DGN Draw Error] Start point: {exc}", "error")
            self.root.after(150, _ask_trace)

        def _ask_trace():
            tid = self.rd_id_var.get().strip() or "RD-01"
            if messagebox.askyesno("Flow (Raindrop) Path",
                                   f"Start point captured:\n{self.rd_lat_var.get()}, {self.rd_lon_var.get()}\n\n"
                                   f"Trace flow path {tid} now?", parent=self.root):
                self.run_raindrop_trace()
        if not IN_MICROSTATION:
            _set(*self.projector.master_to_uor(11055000.0, 3625000.0))
            return
        self._install_picker("Raindrop path: click the start point in DGN (Reset to cancel):", _set)

    def run_raindrop_trace(self):
        try:
            lon_txt, lat_txt = self.rd_lon_var.get().strip(), self.rd_lat_var.get().strip()
            if not lon_txt or not lat_txt:
                raise ValueError("Enter a longitude and latitude, or use \u2316 Pick Point in DGN.")
            lon, lat = float(lon_txt), float(lat_txt)
            if not (-180.0 <= lon <= 180.0 and -90.0 <= lat <= 90.0):
                raise ValueError("Longitude/latitude are outside the valid WGS84 range.")
            dist = float(self.rd_dist_var.get().strip() or "5")
            if not 0.0 < dist <= 500.0:
                raise ValueError("Max downstream distance must be between 0 and 500 miles.")
        except ValueError as exc:
            msg = str(exc) if "could not convert" not in str(exc) else "Longitude, latitude and distance must be numbers."
            messagebox.showwarning("Flow (Raindrop) Path", msg, parent=self.root)
            return
        tid = self.rd_id_var.get().strip() or "RD-01"
        streams = list(self.clipped_streams)

        def _done(res):
            self.raindrop_result = {"trace_id": tid, "result": res}
            self._show_raindrop(tid, res)
            self.log_message(f"Raindrop trace {tid} completed: {res['total_ft'] / 5280.0:.2f} mi through "
                             f"{len(res['reaches'])} reach(es).", "success")
            if self.rd_autodraw_var.get():
                self.raindrop_draw()

        def _err(exc):
            self.rd_summary_var.set("Trace failed.")
            self.log_message(f"[Raindrop Error] {exc}", "error")
            messagebox.showerror("Flow (Raindrop) Path", str(exc), parent=self.root)

        self.log_message(f"Tracing flow (raindrop) path {tid} from {lon:.6f}, {lat:.6f} ({dist:g} mi max)...")
        self._run_async("Tracing flow (raindrop) path\u2026",
                        lambda: self.api_client.trace_raindrop_path(lon, lat, dist, streams), _done, _err)

    def _show_raindrop(self, tid, res):
        self.tree_flow.delete(*self.tree_flow.get_children())
        self.tree_flow.tag_configure("overland", foreground="#b7791f")
        self.tree_flow.tag_configure("stream", foreground="#0e7490")
        for n, (seq, seg, comid, _pts, lf, cf, m) in enumerate(GISExporter._raindrop_rows(res)):
            tags = ("overland" if seg == "OVERLAND" else "stream",) + (("odd",) if n % 2 else ())
            self.tree_flow.insert("", "end", values=(seq, seg, comid or "\u2014", f"{lf:,.0f}", f"{cf / 5280.0:,.3f}", m),
                                  tags=tags)
        self.rd_summary_var.set(
            f"{tid}:  {res['total_ft'] / 5280.0:,.2f} mi  ({res['total_ft']:,.0f} ft)  \u2022  "
            f"overland {res['overland_ft']:,.0f} ft  \u2022  {len(res['reaches'])} downstream reach(es)")

    def raindrop_draw(self):
        if not self.raindrop_result:
            messagebox.showinfo("Flow (Raindrop) Path", "Run a trace first.", parent=self.root)
            return
        try:
            res = self.raindrop_result["result"]
            qp = res.get("query_point") or (None, None)
            already = getattr(self, "_rd_start_drawn", None)
            start_drawn = bool(already and qp[0] is not None
                               and abs(already[0] - qp[0]) < 1e-5 and abs(already[1] - qp[1]) < 1e-5)
            self.dgn_mgr.draw_raindrop_path(res, self.raindrop_result["trace_id"], draw_start=not start_drawn)
        except Exception as exc:
            self.log_message(f"[Raindrop Draw Error] {exc}", "error")
            messagebox.showerror("Flow (Raindrop) Path", str(exc), parent=self.root)

    def raindrop_export(self):
        if not self.raindrop_result:
            messagebox.showinfo("Flow (Raindrop) Path", "Run a trace first.", parent=self.root)
            return
        fmt = self._ask_gis_format()
        if not fmt:
            return
        tid = self.raindrop_result["trace_id"]
        dgn_dir, dgn_stem = MicroStationDgnManager.get_active_dgn_directory_and_stem()
        safe = re.sub(r'[<>:"/\\|?*\s]+', "_", tid).strip("_") or "RD"
        do_kml, do_shp, kmz = "KM" in fmt, "Shapefile" in fmt, "KMZ" in fmt
        ext = (".kmz" if kmz else ".kml") if do_kml else ".shp"
        fpath = filedialog.asksaveasfilename(
            parent=self.root, title=f"Export Raindrop Path \u2013 {fmt}", initialdir=dgn_dir,
            initialfile=f"{dgn_stem}_{safe}{ext}", defaultextension=ext,
            filetypes=[("Google Earth KMZ (*.kmz)", "*.kmz")] if ext == ".kmz" else
                      [("Google Earth KML (*.kml)", "*.kml")] if ext == ".kml" else [("ESRI Shapefile (*.shp)", "*.shp")])
        if not fpath:
            return
        try:
            base = os.path.splitext(fpath)[0]
            created = GISExporter.export_raindrop(base, tid, self.raindrop_result["result"], do_kml, do_shp, kmz)
            for cp in created:
                self.log_message(f"Exported GIS file: {cp}", "success")
            messagebox.showinfo("GIS Export Complete", "Exported files:\n\n" + "\n".join(created), parent=self.root)
        except Exception as exc:
            self.log_message(f"[GIS Export Error] {exc}", "error")
            messagebox.showerror("GIS Export Error", str(exc), parent=self.root)

    def raindrop_clear(self):
        self.raindrop_result = None
        self.tree_flow.delete(*self.tree_flow.get_children())
        self.rd_summary_var.set("No trace yet.")

    def _build_log_card(self, parent):
        card, inner, head = self._card(parent, "Activity Log")
        card.pack(fill="both", expand=True, pady=(10, 0))
        cbtn = ttk.Button(head, text="Clear", style="Secondary.TButton",
                          command=lambda: self.log_box.delete("1.0", tk.END))
        cbtn.pack(side="right")
        Tooltip(cbtn, "Clears the Activity Log window.")
        T = self.T
        self.log_box = scrolledtext.ScrolledText(
            inner, height=7, font=T.MONO, bg=T.LOG_BG, fg=T.LOG_FG, insertbackground=T.LOG_FG,
            relief="flat", borderwidth=0, padx=10, pady=8, wrap="word",
        )
        self.log_box.pack(fill="both", expand=True)
        self.log_box.tag_configure("ts", foreground="#475569")
        self.log_box.tag_configure("info", foreground=T.LOG_FG)
        self.log_box.tag_configure("success", foreground="#4ade80")
        self.log_box.tag_configure("warn", foreground="#fbbf24")
        self.log_box.tag_configure("error", foreground="#f87171")

    def _build_status_bar(self):
        T = self.T
        bar = tk.Frame(self.root, bg=T.NAVY, height=30)
        bar.pack(side="bottom", fill="x")
        bar.pack_propagate(False)
        self.status_var = tk.StringVar(value="Ready")
        self.status_dot = tk.Label(bar, text="\u25cf", bg=T.NAVY, fg="#4ade80", font=T.FONT_SM)
        self.status_dot.pack(side="left", padx=(14, 4))
        tk.Label(bar, textvariable=self.status_var, bg=T.NAVY, fg="#e2e8f0", font=T.FONT_SM).pack(side="left")
        self.progress = ttk.Progressbar(bar, style="Accent.Horizontal.TProgressbar", length=240, mode="determinate")
        self.progress.pack(side="right", padx=14, pady=8)
        tk.Label(bar, text=f"v{APP_VERSION}  \u00b7  Build {BUILD_DATE}", bg=T.NAVY, fg="#94a3b8",
                 font=T.FONT_SM).pack(side="right", padx=(0, 6))
        self._build_env_pills(bar).pack(side="right", padx=(0, 14), pady=3)
        self.svc_var = tk.StringVar(value="\u25cf Services: not checked")
        self.svc_lbl = tk.Label(bar, textvariable=self.svc_var, bg="#13315c", fg="#94a3b8", font=T.FONT_SM,
                                padx=7, pady=0, cursor="hand2")
        self.svc_lbl.pack(side="right", padx=(0, 6), pady=3)
        self.svc_lbl.bind("<Button-1>", lambda _e: self.check_services())
        self._svc_tip = Tooltip(self.svc_lbl, "Web services used by this app. Click to check them again.")
        self._svc_checking = False

    def check_services(self):
        """Probes the USGS / MRLC / NOAA services on a worker thread and reports which are down."""
        if self._svc_checking:
            return
        self._svc_checking = True
        self.svc_var.set("\u25cf Services: checking\u2026")
        self.svc_lbl.configure(bg="#13315c", fg="#94a3b8")

        def _done(results):
            self._svc_checking = False
            USGSWaterServicesClient.apply_service_check(results)
            down = [r for r in results if not r[1]]
            lines = [f"{'OK  ' if ok else 'DOWN'}  {label} ({detail})" for label, ok, detail, _ in results]
            self._svc_tip.text = "\n".join(lines) + "\n\nClick to check again."
            if down:
                self.svc_var.set(f"\u25cf Services: {len(down)} down")
                self.svc_lbl.configure(bg="#3b2330", fg="#fda4af")
                for label, _ok, detail, impact in down:
                    self.log_message(f"[Service down] {label} ({detail}) \u2013 {impact}.", "warn")
            else:
                self.svc_var.set(f"\u25cf Services: all {len(results)} up")
                self.svc_lbl.configure(bg="#13315c", fg="#5eead4")
                self.log_message(f"All {len(results)} web services are responding.", "success")

        def _worker():
            try:
                res = self.api_client.check_services()
            except Exception as exc:
                res = [("Service check", False, str(exc)[:60], "service status is unknown")]
            self._post(_done, res)

        threading.Thread(target=_worker, daemon=True).start()

    # =====================================================================
    # Threading / logging infrastructure
    # =====================================================================
    @staticmethod
    def _classify(msg):
        low = msg.lower()
        if "error" in low or "failed" in low:
            return "error"
        if "warn" in low or "note" in low or "skipp" in low:
            return "warn"
        if any(k in low for k in ("saved", "completed", "exported", "placed", "success", "captured")):
            return "success"
        return "info"

    def log_message(self, msg, level=None):
        """Thread-safe: worker threads enqueue, the Tk thread writes."""
        msg = str(msg)
        level = level or self._classify(msg)
        _trace(f"LOG {msg}")
        if threading.current_thread() is self._main_thread:
            self._append_log(msg, level)
        else:
            self._ui_queue.put(("log", level, msg))

    def _append_log(self, msg, level):
        self.log_box.insert(tk.END, datetime.datetime.now().strftime("%H:%M:%S  "), "ts")
        self.log_box.insert(tk.END, f"{msg}\n", level)
        self.log_box.see(tk.END)

    def _post(self, fn, payload=None):
        self._ui_queue.put(("call", fn, payload))

    def _poll_ui_queue(self):
        try:
            while True:
                kind, a, b = self._ui_queue.get_nowait()
                try:
                    if kind == "log":
                        self._append_log(b, a)
                    elif kind == "call":
                        a(b)
                except Exception as exc:
                    self._append_log(f"[UI Error] {exc}", "error")
        except queue.Empty:
            pass
        # Automatic GC is disabled in main(): collecting on a worker thread runs tkinter
        # Variable.__del__ off the Tk thread and Tcl aborts MicroStation. Collect here instead.
        now = time.monotonic()
        if now - getattr(self, "_last_gc", 0.0) >= 2.0:
            self._last_gc = now
            gc.collect()
        self.root.after(self.POLL_MS, self._poll_ui_queue)

    def _set_busy(self, busy, status=None):
        self._busy = busy
        for w in self._busy_widgets:
            try:
                w.state(["disabled"] if busy else ["!disabled"])
            except tk.TclError:
                pass
        self.status_dot.configure(fg="#fbbf24" if busy else "#4ade80")
        self.status_var.set(status or ("Working\u2026" if busy else "Ready"))
        if not busy:
            _HostWindowLock.unlock()
            self.progress.stop()
            self.progress.configure(mode="determinate", value=0)
            self.root.configure(cursor="")
        else:
            _HostWindowLock.lock()
            self.root.configure(cursor="watch")

    def _run_async(self, status, work, on_done, on_error=None, indeterminate=True):
        """Runs `work()` (network only, no MSPy calls) on a worker thread; callbacks run on the Tk thread."""
        if self._busy:
            messagebox.showinfo("Busy", "Another operation is still running.", parent=self.root)
            return False
        self._set_busy(True, status)
        if indeterminate:
            self.progress.configure(mode="indeterminate")
            self.progress.start(12)

        def _finish(result):
            self._set_busy(False)
            on_done(result)

        def _fail(exc):
            self._set_busy(False)
            (on_error or self._default_error)(exc)

        def _worker():
            try:
                res = work()
            except Exception as exc:
                self._post(_fail, exc)
            else:
                self._post(_finish, res)

        threading.Thread(target=_worker, daemon=True).start()
        return True

    def _default_error(self, exc):
        self.log_message(f"[Error] {exc}", "error")
        messagebox.showerror("Error", str(exc), parent=self.root)

    def _fmt_master(self, x_uor, y_uor):
        xm, ym, _ = self.projector.uor_to_master(x_uor, y_uor)
        return f"({xm:,.1f}, {ym:,.1f})"

    def run_loop(self):
        """Runs the dialog. Inside MicroStation it uses Bentley's cooperative loop (root.update() +
        PyCadInputQueue.PythonMainLoop()) so MicroStation keeps servicing its own input queue and .NET UI;
        a blocking Tk mainloop starves them and can crash MicroStation (0xe0434352) during long requests."""
        pump = None
        if IN_MICROSTATION:
            try:
                pump = PyCadInputQueue.PythonMainLoop
            except NameError:
                pump = None
        _HostWindowLock.ENABLED = pump is None
        _trace(f"run_loop cooperative={pump is not None}")
        self._tk_loop_running = True
        try:
            if pump is None:
                self.root.mainloop()
                return
            while True:
                try:
                    if not self.root.winfo_exists():
                        break
                    self.root.update()
                except tk.TclError:
                    break
                try:
                    pump()
                except Exception as exc:
                    _trace(f"PythonMainLoop error: {exc}")
                time.sleep(0.005)
        finally:
            self._tk_loop_running = False
            _HostWindowLock.unlock()
            _trace("run_loop exit")

    def _install_picker(self, prompt, on_point, multi=False, on_finish=None):
        """Minimises the dialog and installs the DGN point picker. on_point(x, y, z) runs per click.
        With multi=True the tool stays active until Reset (right-click), then on_finish() runs.
        Clicks reach the tool through PyCadInputQueue.PythonMainLoop() in run_loop."""
        self.root.iconify()

        def _resume():
            self._active_tool = None
            try:
                self.root.deiconify()
                self.root.lift()
            except tk.TclError:
                pass

        def _cb(x, y, z):
            if x is not None:
                try:
                    # Picked z follows the view's active depth (can be far outside the model);
                    # draw at elevation 0 like the basins/streams so markers & labels stay visible.
                    on_point(x, y, 0.0)
                except Exception as exc:
                    self.log_message(f"[Pick Error] {exc}", "error")
                if not multi:
                    _resume()
                return
            try:
                if multi or on_finish:
                    if on_finish:
                        on_finish()
                else:
                    self.log_message("Pick cancelled.", "warn")
            finally:
                _resume()

        _trace(f"install picker multi={multi}")
        self._active_tool = DgnPointPickerTool(_cb, prompt, multi)
        self._active_tool.InstallTool()

    def _parse_c_override(self):
        txt = self.c_override_var.get().strip()
        if not txt:
            return None
        c = None
        try:
            c = float(txt)
        except ValueError:
            pass
        if c is None or not 0.0 < c <= 1.0:
            raise ValueError(f"Rational C override must be a number between 0 and 1 (got '{txt}'). "
                             f"Leave it blank to use the computed C.")
        return c

    def _ensure_streams_drawn(self):
        if self.clipped_streams and not self._streams_drawn:
            self.dgn_mgr.draw_clipped_streams(self.clipped_streams)
            self._streams_drawn = True

    def _auto_save_gis(self, prefix, items):
        if not self.auto_gis_var.get():
            return
        try:
            kml_p, shp_list = GISExporter.auto_save_gis_in_dgn_folder(prefix, items, self.clipped_streams)
            self.log_message(f"[Auto-GIS] Saved KML: {kml_p}", "success")
            for sp in shp_list:
                self.log_message(f"[Auto-GIS] Saved Shapefile: {sp}", "success")
        except Exception as exc:
            self.log_message(f"[Auto-GIS] Could not write GIS files beside the DGN ({exc}). "
                             "Use Export GIS\u2026 to choose another folder.", "warn")

    # =====================================================================
    # Fence + NHD streams
    # =====================================================================
    def place_fence_click(self):
        """Places a block or shape fence with the app's own tool (red rubber-band outline), minimises the
        dialog, and reloads streams once the fence is complete. Falls back to MicroStation's Place Fence
        tool if the custom tool cannot start."""
        if not IN_MICROSTATION:
            self.load_fence_and_clip_streams()
            return
        mode = "SHAPE" if self.fence_mode_var.get() == "Shape" else "BLOCK"
        try:
            FenceManager.GetManager().ClearFence()
        except Exception:
            pass
        self._fence_wait_id = getattr(self, "_fence_wait_id", 0) + 1

        def _done(pts, err):
            self._active_tool = None
            try:
                self.root.deiconify()
                self.root.lift()
            except tk.TclError:
                return
            if pts:
                self.log_message("Fence captured.", "success")
                self.load_fence_and_clip_streams()
            else:
                self.fence_status_lbl.configure(text="Fence: placement cancelled \u2014 no active fence")
                self.log_message(f"Fence placement failed: {err}." if err else "Fence placement cancelled.", "warn")

        try:
            tool = DgnFencePlaceTool(mode, _done)
            self._active_tool = tool
            tool.InstallTool()
        except Exception as exc:
            self._active_tool = None
            self.log_message(f"Red fence tool unavailable ({exc}); using MicroStation's Place Fence tool.", "warn")
            self._place_fence_native(mode)
            return
        self.fence_status_lbl.configure(text=f"Fence: waiting for {mode.lower()} fence in DGN\u2026")
        self.log_message(f"Place a {mode.lower()} fence in the DGN"
                         f"{' (Reset to close the shape)' if mode == 'SHAPE' else ' (two corner clicks)'}; "
                         "streams reload automatically.", "info")
        _trace(f"place fence {mode} (red tool)")
        self.root.iconify()

    def _place_fence_native(self, mode):
        """MicroStation's own Place Fence tool (its rubber-band color cannot be changed)."""
        try:
            PyCadInputQueue.SendKeyin(f"PLACE FENCE {mode}")
        except Exception as exc:
            messagebox.showerror("Place Fence", f"Could not start Place Fence: {exc}", parent=self.root)
            return
        self.fence_status_lbl.configure(text=f"Fence: waiting for {mode.lower()} fence in DGN\u2026")
        self.log_message(f"Place a {mode.lower()} fence in the DGN"
                         f"{' (Reset to close the shape)' if mode == 'SHAPE' else ' (two corner clicks)'}; "
                         "streams reload automatically.", "info")
        _trace(f"place fence {mode}")
        self.root.iconify()
        deadline = time.time() + 300.0
        self._fence_wait_id = getattr(self, "_fence_wait_id", 0) + 1
        wait_id = self._fence_wait_id

        def _poll():
            if wait_id != self._fence_wait_id:
                return
            try:
                active = FenceManager.GetManager().IsFenceActive()
            except Exception:
                active = False
            try:
                restored = self.root.state() != "iconic"
            except tk.TclError:
                return
            if active:
                self.root.deiconify()
                self.root.lift()
                self.log_message("Fence captured.", "success")
                self.load_fence_and_clip_streams()
            elif restored or time.time() > deadline:
                self.fence_status_lbl.configure(text="Fence: placement cancelled \u2014 no active fence")
                self.log_message("Fence placement cancelled or timed out.", "warn")
            else:
                self.root.after(400, _poll)

        self.root.after(600, _poll)

    def load_fence_and_clip_streams(self, then=None, draw_outline=True):
        try:
            epsg = int(self.epsg_var.get().strip())
            self.projector._init_pyproj(epsg)
            uor_poly = self.dgn_mgr.get_active_fence_polygon_uor()
            fence_poly = [self.projector.uor_to_wgs84(x, y) for x, y in uor_poly]
        except Exception as exc:
            self.fence_status_lbl.configure(text=f"Fence: {exc}")
            self.log_message(f"[Fence Error] {exc}", "error")
            return
        if draw_outline:
            try:
                if self.dgn_mgr.draw_fence_outline(uor_poly):
                    self.log_message("Fence boundary drawn in red on level 'USGS_Fence_Boundary'.", "info")
            except Exception as exc:
                self.log_message(f"[DGN Draw Error] Fence outline: {exc}", "warn")

        self.fence_wgs84_poly = fence_poly
        self.fence_status_lbl.configure(text="Fence: querying NHDPlus streams\u2026")

        def _done(streams):
            self.clipped_streams = streams or []
            self._streams_drawn = False
            n = len(self.clipped_streams)
            self.fence_status_lbl.configure(text=f"Fence: {n} stream reach(es) clipped inside boundary")
            self.log_message(f"Clipped {n} NHDPlus stream(s) inside the active fence.", "success" if n else "warn")
            if then:
                then()

        def _err(exc):
            self.fence_status_lbl.configure(text="Fence: NHD query failed")
            self.log_message(f"[Fence Error] {exc}", "error")

        self._run_async("Querying NHDPlus streams\u2026",
                        lambda: self.api_client.query_and_clip_nhd_streams(fence_poly), _done, _err)

    # =====================================================================
    # Batch queue
    # =====================================================================
    def _add_to_batch_queue(self, cid, coord_type, val_x, val_y):
        existing = {i["culvert_id"] for i in self.batch_queue}
        base, n = cid, 2
        while cid in existing:
            cid = f"{base}_{n}"
            n += 1
        iid = self.tree_batch.insert(
            "", tk.END, tags=("odd",) if len(self.batch_queue) % 2 else (),
            values=(cid, coord_type, f"{val_x:.6f}" if coord_type == "WGS84_LonLat" else f"{val_x:,.2f}",
                    f"{val_y:.6f}" if coord_type == "WGS84_LonLat" else f"{val_y:,.2f}", "Queued"),
        )
        self.batch_queue.append({"culvert_id": cid, "type": coord_type, "x": val_x, "y": val_y,
                                 "ss_result": None, "iid": iid})

    def _set_batch_status(self, item, text, tag=None):
        vals = list(self.tree_batch.item(item["iid"], "values"))
        vals[4] = text
        base_tags = [t for t in self.tree_batch.item(item["iid"], "tags") if t == "odd"]
        self.tree_batch.item(item["iid"], values=vals, tags=tuple(base_tags + ([tag] if tag else [])))

    def batch_queue_all_fence_streams(self):
        def _queue():
            for s in self.clipped_streams:
                safe = re.sub(r"[^A-Za-z0-9]+", "_", s["name"]).strip("_")[:12] or "Stream"
                lon, lat = s["pour_point_wgs84"]
                self._add_to_batch_queue(f"CULV-{int(s['id']):02d}_{safe}", "WGS84_LonLat", lon, lat)
            self.log_message(f"Queued {len(self.clipped_streams)} clipped stream outlet(s).")

        if self.clipped_streams:
            _queue()
        else:
            self.load_fence_and_clip_streams(then=_queue)

    def _next_culvert_id(self):
        existing = {i["culvert_id"] for i in self.batch_queue}
        n = len(self.batch_queue) + 1
        while f"CULV-{n:02d}" in existing:
            n += 1
        return f"CULV-{n:02d}"

    def batch_place_markers_click(self, _state=None):
        """Click a culvert location in the DGN: the click places pour-point & inlet markers and queues it,
        then asks whether to place another one. Answer No (or Reset/right-click) to stop; then Run Batch
        delineates and draws every drainage area."""
        st = _state if _state is not None else {"total": 0, "last": None}
        dup_tol = 0.5 * self.projector.uor_per_master if hasattr(self.projector, "uor_per_master") else 1.0

        def _done_msg():
            n = st["total"]
            if n:
                self.log_message(f"Placed markers for {n} culvert(s). Click \u25b6 Run Batch to delineate "
                                 f"and draw all drainage areas.", "success")
            else:
                self.log_message("Marker placement ended - no culverts placed.", "warn")

        def _ask_more():
            more = messagebox.askyesno(
                "Place Another Culvert?",
                f"{st['total']} culvert(s) placed.\n\nDo you want to place another culvert?",
                parent=self.root)
            if more:
                self.batch_place_markers_click(_state=st)
            else:
                _done_msg()

        def _place(x, y, z):
            last = st["last"]
            if last and abs(x - last[0]) <= dup_tol and abs(y - last[1]) <= dup_tol:
                self.log_message("Ignored repeated click at the same location.", "warn")
            else:
                st["last"] = (x, y)
                cid = self._next_culvert_id()
                self._add_to_batch_queue(cid, "DGN_UOR", x, y)
                item = self.batch_queue[-1]
                try:
                    item["marker_ids"] = self.dgn_mgr.draw_markers_at_uor(x, y, z, item["culvert_id"])
                    item["click_uor"] = (x, y, z)
                    item["markers_placed"] = True
                    self._set_batch_status(item, "Markers placed \u2013 ready to delineate")
                except Exception as exc:
                    self._set_batch_status(item, f"Queued (marker draw failed: {exc})", "warn")
                st["total"] += 1
                self.log_message(f"Placed markers for {item['culvert_id']} {self._fmt_master(x, y)}.", "success")
            # Ask from the Tk loop after the dialog is restored, not inside the MicroStation tool callback.
            self.root.after(150, _ask_more)

        if not IN_MICROSTATION:
            _place(*self.projector.master_to_uor(11055100.0 + 100.0 * st["total"], 3625050.0))
            return
        n = st["total"] + 1
        self._install_picker(f"Click culvert inlet #{n} (Reset/right-click to cancel):", _place,
                             multi=False, on_finish=_done_msg)

    @staticmethod
    def _classify_xy(v1, v2):
        """Returns (coord_type, x, y). Handles lon,lat and lat,lon (western hemisphere) ordering."""
        if abs(v1) <= 90.0 and abs(v2) <= 90.0:
            # Ambiguous range - use sign convention for US data (lon < 0, lat > 0).
            if v1 > 0 > v2:
                return "WGS84_LonLat", v2, v1
            return "WGS84_LonLat", v1, v2
        if abs(v1) <= 180.0 and abs(v2) <= 90.0:
            return "WGS84_LonLat", v1, v2
        if abs(v2) <= 180.0 and abs(v1) <= 90.0:
            return "WGS84_LonLat", v2, v1
        return "DGN_Master", v1, v2

    def batch_import_csv(self):
        fpath = filedialog.askopenfilename(
            parent=self.root, title="Select Culverts CSV File",
            filetypes=[("CSV Files (*.csv)", "*.csv"), ("Text Files (*.txt)", "*.txt"), ("All Files (*.*)", "*.*")],
        )
        if not fpath:
            return

        added, skipped = 0, []
        try:
            with open(fpath, "r", encoding="utf-8-sig", newline="") as f:
                for line_no, row in enumerate(csv.reader(f), start=1):
                    if len(row) < 3 or not any(c.strip() for c in row):
                        continue
                    try:
                        v1, v2 = float(row[1].strip()), float(row[2].strip())
                    except ValueError:
                        if line_no > 1:
                            skipped.append(line_no)
                        continue  # header row or bad numeric data
                    cid = row[0].strip() or f"CULV-{len(self.batch_queue) + 1:02d}"
                    ctype, x, y = self._classify_xy(v1, v2)
                    self._add_to_batch_queue(cid, ctype, x, y)
                    added += 1
        except (OSError, csv.Error) as exc:
            messagebox.showerror("CSV Import Error", str(exc), parent=self.root)
            return

        self.log_message(f"Imported {added} culvert point(s) from {os.path.basename(fpath)}.", "success")
        if skipped:
            self.log_message(f"Skipped non-numeric CSV line(s): {', '.join(map(str, skipped[:20]))}", "warn")

    def batch_remove_selected(self):
        sel = set(self.tree_batch.selection())
        if not sel:
            return
        self.batch_queue = [i for i in self.batch_queue if i["iid"] not in sel]
        self.tree_batch.delete(*sel)

    def batch_clear_queue(self):
        self.batch_queue.clear()
        self.tree_batch.delete(*self.tree_batch.get_children())
        self.log_message("Cleared batch queue.")

    def delete_all_drawings(self):
        """Delete All Drawings: removes every element on the tool's levels after a confirmation."""
        if self._busy:
            return
        if not IN_MICROSTATION:
            self.log_message("[Simulation] Would delete every element on the USGS_* levels.")
            return
        try:
            n = len(self.dgn_mgr.find_tool_elements())
        except Exception as exc:
            self.log_message(f"[DGN Error] Could not scan the model: {exc}", "error")
            return
        if n == 0:
            messagebox.showinfo("Delete All Drawings", "Nothing to delete \u2014 this tool has no drawings in "
                                "the active model.", parent=self.root)
            return
        levels = "\n".join("  \u2022 " + name for name in self.dgn_mgr.LEVEL_CONFIG)
        if not messagebox.askyesno(
                "Delete All Drawings",
                f"Delete all {n:,} element(s) this tool drew in the active model?\n\n"
                f"Everything on these levels is removed, and the active fence is cleared:\n{levels}\n\n"
                "Use MicroStation Undo to bring them back.", icon="warning", default="no", parent=self.root):
            return
        try:
            deleted = self.dgn_mgr.delete_all_tool_elements()
        except Exception as exc:
            self.log_message(f"[DGN Error] Delete All Drawings: {exc}", "error")
            return
        for it in self.batch_queue:
            it["marker_ids"] = None
            if it.get("ss_result"):
                it["ss_result"]["_dgn_basin_drawn"] = False
        self._streams_drawn = False
        self._rd_start_drawn = None
        try:
            self.fence_status_lbl.configure(text="Fence: cleared \u2014 place a new fence")
        except Exception:
            pass
        self.log_message(f"Deleted {deleted:,} element(s) drawn by this tool and cleared the fence. "
                         "Draw Areas & Markers redraws completed culverts.", "success")

    def pick_batch_table_origin(self):
        def _set(x, y, z):
            self.batch_table_origin_uor = (x, y, z)
            self.batch_origin_lbl.configure(text=f"Table origin: {self._fmt_master(x, y)}")
            self.log_message(f"Captured master batch table origin {self._fmt_master(x, y)}.", "success")

        if not IN_MICROSTATION:
            _set(*self.projector.master_to_uor(11055000.0, 3624800.0))
            return
        self._install_picker("Click in DGN for Master Culvert Table Origin:", _set)

    def batch_stop(self):
        self._cancel_event.set()
        self.log_message("Stop requested \u2013 finishing current culvert\u2026", "warn")

    def run_batch_workflow(self, failed_only=False):
        if self._busy:
            return
        if not self.batch_queue:
            messagebox.showwarning("Batch Queue Empty", "Add culvert points before running the batch.",
                                   parent=self.root)
            return
        if failed_only and all(i.get("ss_result") for i in self.batch_queue):
            messagebox.showinfo("Retry Failed", "Every culvert in the queue already has a result.",
                                parent=self.root)
            return
        try:
            c_override = self._parse_c_override()
        except ValueError as exc:
            messagebox.showwarning("Invalid input", str(exc), parent=self.root)
            return

        rcode = self.rcode_var.get().strip().upper()
        want_lfp = bool(self.lfp_var.get())
        self._batch_retry = failed_only

        # Coordinate transforms use MSPy, so resolve lon/lat on the Tk thread before going async.
        jobs = []
        for item in self.batch_queue:
            if failed_only and item.get("ss_result"):
                continue
            item["ss_result"] = None
            try:
                if item["type"] == "WGS84_LonLat":
                    lon, lat = item["x"], item["y"]
                elif item["type"] == "DGN_UOR":
                    lon, lat = self.projector.uor_to_wgs84(item["x"], item["y"])
                else:
                    xu, yu, _ = self.projector.master_to_uor(item["x"], item["y"])
                    lon, lat = self.projector.uor_to_wgs84(xu, yu)
                jobs.append((item, lon, lat))
                self._set_batch_status(item, "Queued")
            except Exception as exc:
                self._set_batch_status(item, f"Coordinate error: {exc}", "err")

        if not jobs:
            self.log_message("No batch items could be converted to WGS84.", "error")
            return

        total = len(jobs)
        self._cancel_event.clear()
        try:
            self._ensure_streams_drawn()
        except Exception as exc:
            self._streams_drawn = True
            self.log_message(f"[DGN Draw Error] Streams: {exc}", "error")
        _trace("batch start")
        self.log_message(f"{'Retrying' if failed_only else 'Starting batch delineation for'} {total} "
                         f"{'failed ' if failed_only else ''}culvert(s)\u2026")

        def _item_started(payload):
            idx, item = payload
            self._set_batch_status(item, "Delineating\u2026", "busy")
            self._batch_progress_item_started(idx, item["culvert_id"])

        def _item_done(payload):
            idx, item, ss_res, err = payload
            self._batch_progress_item_done(err is None)
            self.progress.configure(value=idx + 1)
            self.status_var.set(f"Batch {idx + 1}/{total} \u2013 {item['culvert_id']}")
            if err is not None:
                self._set_batch_status(item, f"Failed: {err}", "err")
                self.log_message(f"[{idx + 1}/{total}] Error on {item['culvert_id']}: {err}", "error")
                return
            item["ss_result"] = ss_res
            try:
                # Markers are drawn in _finalize_batch, after every drainage area is in the DGN.
                self.dgn_mgr.draw_basin_rings(ss_res, item["culvert_id"])
                self.dgn_mgr.draw_longest_flowpath(ss_res, item["culvert_id"])
                self._set_batch_status(item, "Completed", "ok")
            except Exception as exc:
                self._set_batch_status(item, f"Delineated; DGN draw failed: {exc}", "warn")
            self.log_message(f"[{idx + 1}/{total}] Delineated {item['culvert_id']}.", "success")

        def _worker():
            for idx, (item, lon, lat) in enumerate(jobs):
                if self._cancel_event.is_set():
                    break
                self._post(_item_started, (idx, item))
                try:
                    res = self.api_client.delineate_streamstats_watershed(rcode, lon, lat, c_override, include_lfp=want_lfp)
                    self._post(_item_done, (idx, item, res, None))
                except Exception as exc:
                    self._post(_item_done, (idx, item, None, exc))
                # Pause between requests to be polite to the USGS service; wakes at once on Stop.
                if idx < total - 1 and self._cancel_event.wait(1.2):
                    break
            return self._cancel_event.is_set()

        def _finish(cancelled):
            self.stop_btn.state(["disabled"])
            if self._bp:
                self._bp["cur"] = None
                self.batch_prog_lbl.configure(text="Drawing master table and saving reports\u2026")
                self.root.update_idletasks()
            try:
                self._finalize_batch(rcode, cancelled)
            finally:
                self._batch_progress_finish(cancelled)

        def _fail(exc):
            self.stop_btn.state(["disabled"])
            self._batch_progress_finish(True, text=f"Batch failed: {exc}")
            self._default_error(exc)

        self._batch_progress_start(total)
        if self._run_async(f"Running batch of {total}\u2026", _worker, _finish, _fail, indeterminate=False):
            self.progress.configure(maximum=total, value=0)
            self.stop_btn.state(["!disabled"])
        else:
            self._batch_progress_finish(True, text="Batch progress: idle (another task is running)")

    def _finalize_batch(self, rcode, cancelled):
        done_items = [i for i in self.batch_queue if i.get("ss_result")]
        if done_items:
            # Guarantees every delineated drainage area is in the DGN, then markers for culverts without pre-placed ones.
            self._draw_areas_then_markers(done_items, markers=bool(self.auto_markers_var.get()))
        if cancelled:
            for i in self.batch_queue:
                if i.get("ss_result") is None and self.tree_batch.item(i["iid"], "values")[4] == "Queued":
                    self._set_batch_status(i, "Skipped (stopped)", "warn")

        if not done_items:
            self.log_message("Batch finished with no successful delineations.", "error")
            return

        try:
            if self.batch_table_origin_uor:
                tx, ty = self.batch_table_origin_uor[0], self.batch_table_origin_uor[1]
            else:
                tx, ty = self._default_table_origin(done_items)
            self.dgn_mgr.place_batch_master_table(self.batch_queue, tx, ty,
                                                  replace_previous=getattr(self, "_batch_retry", False))
        except Exception as exc:
            self.log_message(f"[Batch Table Error] {exc}", "error")

        report_dir = ""
        try:
            csv_p, xlsx_p = BatchSummaryExporter.export_batch_summary(self.batch_queue, rcode)
            report_dir = os.path.dirname(csv_p)
            self.log_message(f"Saved batch CSV report: {csv_p}", "success")
            if xlsx_p:
                self.log_message(f"Saved batch Excel report: {xlsx_p}", "success")
            else:
                self.log_message("openpyxl not installed \u2013 Excel report skipped.", "warn")
        except Exception as exc:
            _trace("Batch report traceback:\n" + traceback.format_exc())
            self.log_message(f"[Batch Report Error] {exc}", "error")

        self._auto_save_gis("Batch_Watersheds", done_items)
        self._refresh_results(select=done_items[0]["culvert_id"])

        warnings = self._check_batch_results(done_items)
        n_fail = len(self.batch_queue) - len(done_items)
        self.status_var.set(f"Batch complete \u2013 {len(done_items)} ok, {n_fail} failed/skipped"
                            + (f", {len(warnings)} warning(s)" if warnings else ""))
        msg = (f"Delineated {len(done_items)} of {len(self.batch_queue)} culvert(s)."
               + (f"\n\nReports saved to:\n{report_dir}" if report_dir else ""))
        if warnings:
            messagebox.showwarning(
                "Batch Complete \u2013 Check Results",
                msg + "\n\nCHECK THESE CULVERTS:\n\n\u2022 " + "\n\n\u2022 ".join(warnings),
                parent=self.root,
            )
        else:
            messagebox.showinfo("Batch Complete", msg, parent=self.root)

    SMALL_BASIN_ACRES = 1.0

    def _check_batch_results(self, done_items):
        """Flags culverts that StreamStats snapped to the same outlet as another culvert (their rows are copies)
        and drainage areas under SMALL_BASIN_ACRES (the point is probably off the stream). Logs each warning,
        marks the batch row, and returns the warning texts."""
        warnings = []
        by_outlet = {}
        for it in done_items:
            sp = it["ss_result"].get("snapped_point") or (None, None)
            if sp[0] is None:
                continue
            by_outlet.setdefault((round(sp[0], 6), round(sp[1], 6)), []).append(it)
        flagged = {}
        for group in by_outlet.values():
            if len(group) < 2:
                continue
            ids = [g["culvert_id"] for g in group]
            warnings.append(f"{', '.join(ids)} snapped to the same StreamStats outlet, so they share one drainage "
                            f"area and identical flows. Move each culvert onto its own stream and re-run.")
            for g in group:
                flagged[g["culvert_id"]] = (g, "Check: shared outlet")
        for it in done_items:
            met = (it["ss_result"].get("vdot_comparison") or {}).get("metrics", {})
            ac = met.get("drnarea_acres")
            if ac is not None and ac < self.SMALL_BASIN_ACRES:
                warnings.append(f"{it['culvert_id']} drainage area is only {ac:.2f} ac. The point is probably "
                                f"off the stream channel. Move it onto the stream and re-run.")
                prev = flagged.get(it["culvert_id"])
                flagged[it["culvert_id"]] = (it, (prev[1] + ", tiny area") if prev else "Check: tiny area")
        for it, text in flagged.values():
            self._set_batch_status(it, text, "warn")
        for w in warnings:
            self.log_message("WARNING: " + w, "warn")
        return warnings

    # =====================================================================
    # Results tab
    # =====================================================================
    def _all_results(self):
        items = {}
        for i in self.batch_queue:
            if i.get("ss_result"):
                items[i["culvert_id"]] = i["ss_result"]
        return items

    def _refresh_results(self, select=None):
        names = list(self._all_results().keys())
        self.result_combo.configure(values=names)
        if select in names:
            self.result_sel_var.set(select)
        elif names and self.result_sel_var.get() not in names:
            self.result_sel_var.set(names[0])
        self._show_result(self.result_sel_var.get())

    def _show_result(self, cid):
        ss_res = self._all_results().get(cid)
        self.tree_results.delete(*self.tree_results.get_children())
        if not ss_res:
            for v in self.metric_vars.values():
                v.set("\u2014")
            self.result_src_var.set("Data sources for the selected culvert are listed here after a delineation.")
            return
        vcomp = ss_res.get("vdot_comparison", {}) or {}
        met = vcomp.get("metrics", {}) or {}
        tc_len = vcomp.get("tc_length_ft")
        dash = "\u2014"
        tc_note = f" ({tc_len:,.0f} ft)" if tc_len else ""
        sources = (f"Basin data: {vcomp.get('basin_source', dash)}   \u00b7   "
                   f"Rainfall: {vcomp.get('rainfall_source', dash)}   \u00b7   "
                   f"Tc length: {vcomp.get('tc_length_source', dash)}{tc_note}   \u00b7   "
                   f"Regression: {vcomp.get('regression_source', dash)}")
        if vcomp.get("cn"):
            sources += f"   \u00b7   NRCS CN {vcomp['cn']:.0f}: {vcomp.get('cn_source', dash)}"
        skipped = [f"{m['label']} \u2013 {m['status']}" for m in vcomp.get("methods", [])
                   if not m["calculated"] and m["key"] in VDOTHydrologyCalculator.default_options()["enabled"]]
        if skipped:
            sources += "   \u00b7   Not calculated: " + "; ".join(skipped)
        flagged = "placeholder" in sources or "Alternate" in sources or "Estimated" in sources
        self.result_src_var.set(("\u26a0  " if flagged else "") + sources
                                + "   \u00b7   Verify against the VDOT Drainage Manual before design use.")
        self.metric_vars["area"].set(f"{met.get('drnarea_sqmi', 0.0):.3f} mi\u00b2  /  {met.get('drnarea_acres', 0.0):,.1f} ac")
        self.metric_vars["slope"].set(f"{met.get('slope_ft_mi', 0.0):.1f} ft/mi")
        self.metric_vars["c"].set(f"{vcomp.get('c_coeff', 0.0):.2f}")
        self.metric_vars["tc"].set(f"{vcomp.get('tc_min', 0.0):.1f} min")
        self.metric_vars["gov"].set(vcomp.get("governing_method", "\u2014"))

        for i, r in enumerate(vcomp.get("rows", [])):
            usgs = r["usgs_q_cfs"]
            diff = r["pct_diff"]
            tags = ["odd"] if i % 2 else []
            try:
                if abs(float(diff.rstrip("%"))) > 25.0:
                    tags.append("warn")
            except (ValueError, AttributeError):
                pass
            num = lambda v: f"{v:,.1f}" if isinstance(v, (int, float)) and not isinstance(v, bool) else "\u2014"
            self.tree_results.insert("", tk.END, tags=tuple(tags), values=(
                f"Q{r['return_yr']}  ({r['return_yr']}-yr)",
                num(r.get("q_rational_cfs")), num(r.get("q_nrcs_cfs")), num(r.get("q_anderson_cfs")),
                num(usgs), num(r.get("q_urban_cfs")), num(r.get("vdot_q_cfs")), diff,
            ))

    def export_results_excel(self):
        results = {k: v for k, v in self._all_results().items() if v}
        if not results:
            messagebox.showwarning("No Results", "Run a delineation first, then export the results.", parent=self.root)
            return
        if not HAS_OPENPYXL:
            messagebox.showerror("Excel Export", "The 'openpyxl' library is not installed (see the Libraries tab).",
                                 parent=self.root)
            return
        dgn_dir, dgn_stem = MicroStationDgnManager.get_active_dgn_directory_and_stem()
        fpath = filedialog.asksaveasfilename(
            parent=self.root, title="Export Results & Comparison to Excel", initialdir=dgn_dir,
            initialfile=f"{dgn_stem}_Results_Comparison.xlsx", defaultextension=".xlsx",
            filetypes=[("Excel Workbook (*.xlsx)", "*.xlsx")])
        if not fpath:
            return
        try:
            BatchSummaryExporter.export_results_workbook(fpath, results)
            self.log_message(f"Exported Results & Comparison: {fpath}", "success")
            if messagebox.askyesno("Excel Export Complete", f"Saved:\n\n{fpath}\n\nOpen it now?", parent=self.root):
                os.startfile(fpath)
        except PermissionError:
            messagebox.showerror("Excel Export Error", f"Cannot write:\n{fpath}\n\nClose it in Excel and try again.",
                                 parent=self.root)
        except Exception as exc:
            self.log_message(f"[Excel Export Error] {exc}", "error")
            messagebox.showerror("Excel Export Error", str(exc), parent=self.root)

    # =====================================================================
    # Markers
    # =====================================================================
    def draw_markers_for(self, source="batch"):
        items = [i for i in self.batch_queue if i.get("ss_result")]
        if not items:
            messagebox.showwarning("No Delineations", "Run a delineation first, then draw its drainage area, "
                                   "pour-point and culvert-inlet markers.", parent=self.root)
            return
        self._draw_areas_then_markers(items)

    def _default_table_origin(self, done_items):
        """Top-left of the table: just right of the combined drainage-area extents, level with their top."""
        xs, ys = [], []
        for it in done_items:
            ssr = it["ss_result"]
            for ring in (ssr.get("basin_rings_wgs84") or [])[:1]:
                if ring:
                    lons = [p[0] for p in ring]
                    lats = [p[1] for p in ring]
                    for lon, lat in ((min(lons), min(lats)), (max(lons), max(lats))):
                        x, y, _ = self.projector.wgs84_to_uor(lon, lat)
                        xs.append(x)
                        ys.append(y)
        if xs:
            gap = 3.0 * self.dgn_mgr.annot_unit(max(max(xs) - min(xs), max(ys) - min(ys)))
            return max(xs) + gap, max(ys)
        plon, plat = done_items[0]["ss_result"].get("snapped_point") or done_items[0]["ss_result"]["query_point"]
        px, py, _ = self.projector.wgs84_to_uor(plon, plat)
        return px + 150.0 * self.projector.uor_per_master, py

    def _resize_preplaced_markers(self, it):
        """Markers placed before delineation were sized without a drainage area; once the basin exists,
        replace them with markers proportional to that drainage area."""
        ids, at = it.get("marker_ids"), it.get("click_uor")
        if not ids or not at or not it["ss_result"].get("_dgn_basin_drawn"):
            return
        try:
            u = self.dgn_mgr.annot_unit(self.dgn_mgr.basin_extent_uor(it["ss_result"]))
            self.dgn_mgr.delete_elements(ids)
            it["marker_ids"] = self.dgn_mgr.draw_markers_at_uor(*at, it["culvert_id"], unit=u)
        except Exception as exc:
            it["marker_ids"] = None
            self.log_message(f"[DGN Draw Error] Resizing markers for {it['culvert_id']}: {exc}", "warn")

    def _draw_areas_then_markers(self, items, markers=True):
        """Pass 1: every drainage area (basin) not yet in the DGN. Pass 2: pour-point & inlet markers
        (skipped for culverts whose markers were already placed with Place Markers in DGN)."""
        areas = 0
        for it in items:
            ssr = it["ss_result"]
            if ssr.get("_dgn_basin_drawn"):
                continue
            try:
                if self.dgn_mgr.draw_basin_rings(ssr, it["culvert_id"]):
                    areas += 1
                    self.log_message(f"Drew drainage area for {it['culvert_id']}.")
                else:
                    if IN_MICROSTATION:
                        self.log_message(f"[WARNING] No drainage area geometry for {it['culvert_id']}.", "warn")
            except Exception as exc:
                self.log_message(f"[DGN Draw Error] Drainage area for {it['culvert_id']}: {exc}", "error")
        total_areas = sum(1 for it in items if it["ss_result"].get("_dgn_basin_drawn"))
        for it in items:
            if it.get("markers_placed"):
                self._resize_preplaced_markers(it)
        n = 0
        if markers:
            for it in items:
                if it.get("markers_placed"):
                    continue
                try:
                    n += self.dgn_mgr.draw_culvert_markers(it["ss_result"], it["culvert_id"])
                    it["markers_placed"] = True
                except Exception as exc:
                    self.log_message(f"[DGN Draw Error] Markers for {it['culvert_id']}: {exc}", "error")
        self.log_message(f"Drainage areas in DGN: {total_areas}/{len(items)} ({areas} newly drawn); "
                         f"markers drawn for {n} more culvert(s).",
                         "success" if total_areas == len(items) or not IN_MICROSTATION else "warn")

    # =====================================================================
    # KML / KMZ symbology
    # =====================================================================
    def open_kml_style_dialog(self):
        T = self.T
        style = GISExporter.load_kml_style()
        dlg = tk.Toplevel(self.root)
        dlg.title("KML / KMZ Symbology")
        dlg.configure(bg=T.CARD)
        dlg.resizable(False, False)
        dlg.transient(self.root)
        frm = ttk.Frame(dlg, style="Inner.TFrame", padding=18)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="KML / KMZ Symbology", style="CardTitle.TLabel").grid(row=0, column=0, columnspan=4, sticky="w")
        ttk.Label(frm, text="Applies to every KML/KMZ export. Saved to:",
                  style="Muted.TLabel").grid(row=1, column=0, columnspan=4, sticky="w")
        path_var = tk.StringVar(value=GISExporter.kml_style_path())
        prow = ttk.Frame(frm, style="Inner.TFrame")
        prow.grid(row=2, column=0, columnspan=4, sticky="we", pady=(2, 12))
        prow.columnconfigure(0, weight=1)
        ttk.Entry(prow, textvariable=path_var, state="readonly", width=58).grid(row=0, column=0, sticky="we")

        def _browse():
            cur = path_var.get()
            p = filedialog.asksaveasfilename(
                parent=dlg, title="KML / KMZ Style File Location", initialdir=os.path.dirname(cur),
                initialfile=os.path.basename(cur), defaultextension=".json", confirmoverwrite=False,
                filetypes=[("Style file (*.json)", "*.json"), ("All files", "*.*")])
            if not p:
                return
            path_var.set(os.path.normpath(p))
            if os.path.isfile(p) and messagebox.askyesno(
                    "KML Style", "This style file already exists.\n\nLoad its colors and widths into this dialog?\n"
                    "(No keeps the current settings and overwrites it on Save.)", parent=dlg):
                _load_into(GISExporter.load_kml_style(p))

        def _default_loc():
            path_var.set(GISExporter.DEFAULT_KML_STYLE_FILE)

        b = ttk.Button(prow, text="Browse\u2026", style="Secondary.TButton", command=_browse)
        b.grid(row=0, column=1, padx=(8, 0))
        Tooltip(b, "Choose a different file/folder for the KML style settings (e.g. a shared network location).")
        b = ttk.Button(prow, text="Default", style="Secondary.TButton", command=_default_loc)
        b.grid(row=0, column=2, padx=(6, 0))
        Tooltip(b, "Resets the style file location to the default in your user folder.")

        vars_ = {}
        swatches = {}

        def _swatch(row, key, label):
            ttk.Label(frm, text=label, style="Muted.TLabel").grid(row=row, column=0, sticky="w", pady=3)
            sw = tk.Label(frm, width=6, bg=style[key], relief="solid", bd=1, cursor="hand2")
            sw.grid(row=row, column=1, sticky="w", padx=(8, 16))
            vars_[key] = tk.StringVar(value=style[key])
            swatches[key] = sw

            def _pick(_e=None):
                c = colorchooser.askcolor(color=vars_[key].get(), parent=dlg, title=label)[1]
                if c:
                    vars_[key].set(c.upper())
                    sw.configure(bg=c)
            sw.bind("<Button-1>", _pick)

        def _num(row, col, key, label, lo, hi, inc):
            ttk.Label(frm, text=label, style="Muted.TLabel").grid(row=row, column=col, sticky="w", pady=3)
            vars_[key] = tk.StringVar(value=str(style[key]))
            ttk.Spinbox(frm, from_=lo, to=hi, increment=inc, textvariable=vars_[key], width=7).grid(
                row=row, column=col + 1, sticky="w", padx=(8, 0))

        r = 3
        _swatch(r, "basin_line_color", "Basin outline color")
        _num(r, 2, "basin_line_width", "Width", 0.5, 10, 0.5)
        r += 1
        _swatch(r, "basin_fill_color", "Basin fill color")
        _num(r, 2, "basin_fill_opacity", "Opacity %", 0, 100, 5)
        r += 1
        vars_["basin_fill"] = tk.BooleanVar(value=bool(style["basin_fill"]))
        ttk.Checkbutton(frm, text="Fill basin polygons", variable=vars_["basin_fill"],
                        style="Card.TCheckbutton").grid(row=r, column=0, columnspan=2, sticky="w", pady=3)
        r += 1
        _swatch(r, "stream_color", "NHD stream color")
        _num(r, 2, "stream_width", "Width", 0.5, 10, 0.5)
        r += 1
        _swatch(r, "flowpath_color", "Longest flowpath color")
        _num(r, 2, "flowpath_width", "Width", 0.5, 10, 0.5)
        r += 1
        _swatch(r, "raindrop_color", "Raindrop path color")
        _num(r, 2, "raindrop_width", "Width", 0.5, 10, 0.5)
        r += 1
        _swatch(r, "point_color", "Pour-point icon color")
        _num(r, 2, "point_scale", "Scale", 0.2, 4, 0.1)
        r += 1

        ttk.Label(frm, text="Pour-point icon", style="Muted.TLabel").grid(row=r, column=0, sticky="w", pady=3)
        cur_icon = next((k for k, v in GISExporter.KML_ICONS.items() if v == style["point_icon"]), style["point_icon"])
        vars_["point_icon"] = tk.StringVar(value=cur_icon)
        ttk.Combobox(frm, textvariable=vars_["point_icon"], values=list(GISExporter.KML_ICONS), width=14).grid(
            row=r, column=1, sticky="w", padx=(8, 16))
        _num(r, 2, "label_scale", "Label scale", 0, 4, 0.1)
        r += 1

        def _collect():
            out = {}
            for k, v in vars_.items():
                val = v.get()
                d = GISExporter.KML_STYLE_DEFAULTS[k]
                if isinstance(d, bool):
                    out[k] = bool(val)
                elif isinstance(d, (int, float)):
                    try:
                        out[k] = float(val)
                    except ValueError:
                        out[k] = d
                else:
                    out[k] = val
            out["point_icon"] = GISExporter.KML_ICONS.get(out["point_icon"], out["point_icon"])
            return out

        def _save():
            try:
                path = path_var.get()
                GISExporter.save_kml_style(_collect(), path)
                GISExporter.set_kml_style_path(path)
                self.log_message(f"Saved KML/KMZ symbology: {path}", "success")
                dlg.destroy()
            except Exception as exc:
                messagebox.showerror("KML Style", str(exc), parent=dlg)

        def _load_into(src):
            for k, d in src.items():
                if k not in vars_:
                    continue
                if k == "point_icon":
                    vars_[k].set(next((n for n, u in GISExporter.KML_ICONS.items() if u == d), d))
                else:
                    vars_[k].set(d)
                if k in swatches:
                    swatches[k].configure(bg=d)

        def _reset():
            _load_into(GISExporter.KML_STYLE_DEFAULTS)

        btns = ttk.Frame(frm, style="Inner.TFrame")
        btns.grid(row=r, column=0, columnspan=4, sticky="we", pady=(16, 0))
        for txt, st, cmd, side, px, tip in (
                ("Save", "Accent.TButton", _save, "right", 0, "Saves these styles to the style file; used by every KML/KMZ export."),
                ("Cancel", "Secondary.TButton", dlg.destroy, "right", 6, "Closes without saving."),
                ("Reset Defaults", "Secondary.TButton", _reset, "left", 0,
                 "Restores the default colors and widths in this dialog (click Save to keep them).")):
            b = ttk.Button(btns, text=txt, style=st, command=cmd)
            b.pack(side=side, padx=px)
            Tooltip(b, tip)
        dlg.bind("<Escape>", lambda _e: dlg.destroy())
        dlg.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - dlg.winfo_width()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - dlg.winfo_height()) // 3
        dlg.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        dlg.grab_set()

    # =====================================================================
    # GIS export
    # =====================================================================
    def export_gis_dialog(self, source="auto"):
        """Exports delineated watersheds & clipped streams to KML, Shapefile, or both."""
        items = [i for i in self.batch_queue if i.get("ss_result")]

        if not items and not self.clipped_streams:
            messagebox.showwarning(
                "No GIS Geometry",
                "Delineate a watershed or clip NHDPlus streams before exporting.", parent=self.root)
            return

        fmt = self._ask_gis_format()
        if not fmt:
            return

        dgn_dir, dgn_stem = MicroStationDgnManager.get_active_dgn_directory_and_stem()
        kml_ext = ".kmz" if "KMZ" in fmt else ".kml"
        do_kml = "KM" in fmt
        do_shp = "Shapefile" in fmt
        if do_kml and not do_shp:
            ext = kml_ext
            ftypes = [("Google Earth KMZ (*.kmz)", "*.kmz")] if ext == ".kmz" else [("Google Earth KML (*.kml)", "*.kml")]
        else:
            ext = ".shp"
            ftypes = [("ESRI Shapefile (*.shp)", "*.shp")]
        fpath = filedialog.asksaveasfilename(
            parent=self.root, title=f"Export Watersheds & Streams \u2013 {fmt}",
            initialdir=dgn_dir, initialfile=f"{dgn_stem}_USGS_Watersheds{ext}",
            defaultextension=ext, filetypes=ftypes,
        )
        if not fpath:
            return

        try:
            base = os.path.splitext(fpath)[0]
            created = []
            if do_kml:
                created.append(GISExporter.export_kml(base + kml_ext, items, self.clipped_streams))
            if do_shp:
                created.extend(GISExporter.export_shapefiles(base, items, self.clipped_streams))
            for cp in created:
                self.log_message(f"Exported GIS file: {cp}", "success")
            messagebox.showinfo("GIS Export Complete", "Exported files:\n\n" + "\n".join(created), parent=self.root)
        except Exception as exc:
            self.log_message(f"[GIS Export Error] {exc}", "error")
            messagebox.showerror("GIS Export Error", str(exc), parent=self.root)

    def _ask_gis_format(self):
        T = self.T
        dlg = tk.Toplevel(self.root)
        dlg.title("Export GIS")
        dlg.configure(bg=T.CARD)
        dlg.resizable(False, False)
        dlg.transient(self.root)
        result = {"fmt": None}

        frm = ttk.Frame(dlg, style="Inner.TFrame", padding=18)
        frm.pack(fill="both", expand=True)
        ttk.Label(frm, text="Choose export format", style="CardTitle.TLabel").pack(anchor="w")
        ttk.Label(frm, text="Watershed polygons, pour points, clipped streams and longest flowpaths.",
                  style="Muted.TLabel").pack(anchor="w", pady=(0, 10))
        fmt_var = tk.StringVar(value="KML + Shapefile")
        combo = ttk.Combobox(frm, textvariable=fmt_var, state="readonly",
                             values=["KML + Shapefile", "KMZ + Shapefile", "KML", "KMZ", "Shapefile"], width=24)
        combo.pack(anchor="w", fill="x")

        btns = ttk.Frame(frm, style="Inner.TFrame")
        btns.pack(fill="x", pady=(16, 0))

        def _ok():
            result["fmt"] = fmt_var.get()
            dlg.destroy()

        b = ttk.Button(btns, text="Export", style="Accent.TButton", command=_ok)
        b.pack(side="right")
        Tooltip(b, "Continues to choose where to save the files in the selected format.")
        b = ttk.Button(btns, text="Cancel", style="Secondary.TButton", command=dlg.destroy)
        b.pack(side="right", padx=6)
        Tooltip(b, "Closes without exporting.")
        dlg.bind("<Return>", lambda _e: _ok())
        dlg.bind("<Escape>", lambda _e: dlg.destroy())

        dlg.update_idletasks()
        x = self.root.winfo_rootx() + (self.root.winfo_width() - dlg.winfo_width()) // 2
        y = self.root.winfo_rooty() + (self.root.winfo_height() - dlg.winfo_height()) // 3
        dlg.geometry(f"+{max(x, 0)}+{max(y, 0)}")
        dlg.grab_set()
        combo.focus_set()
        self._wait_dialog(dlg)
        return result["fmt"]

    def _wait_dialog(self, dlg):
        """Waits for a modal Toplevel to close. Inside MicroStation it keeps pumping
        PyCadInputQueue.PythonMainLoop() (like run_loop) so MicroStation doesn't stall meanwhile."""
        pump = None
        if IN_MICROSTATION and getattr(self, "_tk_loop_running", False):
            try:
                pump = PyCadInputQueue.PythonMainLoop
            except NameError:
                pump = None
        if pump is None:
            self.root.wait_window(dlg)
            return
        while True:
            try:
                if not dlg.winfo_exists():
                    break
                self.root.update()
            except tk.TclError:
                break
            try:
                pump()
            except Exception as exc:
                _trace(f"PythonMainLoop error (dialog): {exc}")
            time.sleep(0.01)


# ---------------------------------------------------------------------------
# 11. Main Entry Point
# ---------------------------------------------------------------------------
def main():
    saved_streams = _install_thread_safe_std_streams()
    gc_was_enabled = gc.isenabled()
    gc.disable()  # garbage collection runs only on the Tk thread (see _poll_ui_queue)
    root = tk.Tk()
    global _APP
    try:
        _APP = StreamStatsFenceDialog(root)
        _APP.run_loop()
    finally:
        _HostWindowLock.unlock()
        gc.collect()
        if gc_was_enabled:
            gc.enable()
        _restore_std_streams(saved_streams)


if __name__ == "__main__":
    main()
