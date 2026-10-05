from __future__ import annotations

import html
import json
import os
import re
import subprocess
import sys
import tempfile
import threading
import unicodedata
from dataclasses import dataclass
from urllib.parse import quote

import wx

# ConpubsWalker -- walks the whole conpubs tree (root -> series -> instances -> sub-pages) and,
# per the checkboxes in its dialog, gives every PDF one or both of:
#   * a page-header check/update (the same header ConEditor stamps on upload), either always
#     ("Always re-stamp") or only when the existing header is missing or differs from what the
#     page data says it should be ("Only when missing or incorrect"); and/or
#   * an OCR-quality check, OCR-ing NOT_OCRED PDFs with ABBYY FineReader 15.
# A PDF is downloaded once, both operations are applied to the same copy (OCR first -- ABBYY
# rewrites the file, which would destroy a header stamped before it), and it is re-uploaded only
# if something actually changed.
#
# Two done-logs record every PDF checked and are reloaded at startup so finished PDFs are skipped:
#   "ConpubsWalker Headers Done.txt" -- status (OK / STAMPED) + server path
#   "ConpubsWalker OCR Done.txt"     -- the OCR-quality survey columns + server path
# A PDF whose check failed (download error, ABBYY failure, failed upload) is NOT recorded, so the
# next run retries it. "Clear State And Rerun" deletes both logs and starts over.
#
# This absorbs ConEditorAuto (CPA): the tree traversal and series/instance parsing below are
# ported from CPAWalker.py, and the header stamping is CPA's planned Stage 2, built on ConEditor's
# shared PDFHelpers engine. Header text and dates come from the series/instance pages (dates from
# the series page's Dates column -- ConEditor used the instance page's own dates, which should be
# the same text; a difference shows up as an "incorrect" header and is re-stamped canonically).

# ---------- path setup -----------------------------------------------
# Run from the ConpubsWalker dir; the shared modules live in sibling directories.
_DIR = os.path.dirname(os.path.abspath(__file__))
_PYROOT = os.path.dirname(_DIR)
_CONEDITOR = os.path.join(_PYROOT, 'ConEditor')
for _sib in ('FTP', 'Settings', 'HelpersPackage', 'FanzineIssueSpecPackage', 'Locale', 'WxDataGrid', 'ConEditor'):
    _p = os.path.join(_PYROOT, _sib)
    if _p not in sys.path:
        sys.path.insert(0, _p)

from pypdf import PdfReader
from HelpersPackage import FindBracketedText2, FindLinkInString, RemoveAccents
from PDFHelpers import (LowQualityScan, OcrQuality, AddPdfPageHeader,
                        GetPdfPageHeaderText, PdfPageHeaderDisplayText)
from FTP import FTP
from Log import Log, LogOpen, LogClose, LogError
from ConInstance import ConInstance       # ConEditor's canonical (headless) con-instance page reader

# ---------- configuration --------------------------------------------
LIMIT_SERIES   = "xxxTest"    # walk only this con series ("" = the whole tree); display or dir name
LIMIT_INSTANCE = ""           # walk only this con instance within the series ("" = all)

CONPUBS_ROOT_URL = "https://www.fanac.org/conpubs/"
FTP_CREDS      = os.path.join(_CONEDITOR, "FTP Credentials.json")
LOGO_FILE      = os.path.join(_CONEDITOR, "Fanac logo for pdf headers.jpg")
FINECMD        = r"C:\Program Files (x86)\ABBYY FineReader 15\FineCmd.exe"
HDR_DONE_FILE  = os.path.join(_DIR, "ConpubsWalker Headers Done.txt")
OCR_DONE_FILE  = os.path.join(_DIR, "ConpubsWalker OCR Done.txt")
STATE_FILE     = os.path.join(_DIR, "ConpubsWalker State.json")

# OCR done-log columns (unchanged from the original survey format)
_W_ALPHA = 11   # "alpha count"  column — right-justified
_W_WORDS = 10   # "# in words"   column — right-justified
_W_RATIO =  5   # "ratio"        column — right-justified
_W_LONG  = 12   # "words>5 char" column — right-justified
_W_TAG   = 10   # "Status"       column — left-justified  (NOT_OCRED = 9)
_C_TAG   = _W_ALPHA + 2 + _W_WORDS + 2 + _W_RATIO + 2 + _W_LONG + 2   # = 46
_C_PATH  = _C_TAG + _W_TAG + 2                                        # = 58
_OCR_HEADER = (f"{'alpha count':>{_W_ALPHA}}  {'# in words':>{_W_WORDS}}  "
               f"{'ratio':>{_W_RATIO}}  {'words>5 char':>{_W_LONG}}  "
               f"{'Status':<{_W_TAG}}  pdf server path\n")
_HDR_HEADER = f"{'Status':<{_W_TAG}}  pdf server path\n"

_g_logo: bytes|None = None    # the page-header logo art, loaded once at startup (None = no logo)


# ---------- series/instance page parsing -----------------------------
# Ported from ConEditorAuto/CPAWalker.py. These mirror ConEditorFrame.DownloadMainConlist (root)
# and ConSeriesFrame.LoadConSeriesFromHTML + its ReadTableRow/ConNameInfoUnpack statics. They are
# duplicated (rather than importing the GUI frames) to keep this headless walker decoupled from wx
# frame classes; keep them in sync if those parsers change. Instance pages are read by the reused
# ConInstance class, not here.

@dataclass
class Instance:
    name: str       # display name of the con instance
    dir: str        # the instance's directory name on the server (under the series dir)
    dates: str      # the con's date(s) from the series page (used in the PDF page header)
    linked: bool    # True if the series-page entry actually links to a page (unlinked entries are common/intentional)


# Mirrors ConSeriesFrame.ReadTableRow
def _ReadTableRow(row: str, delim: str="td") -> list[str]:
    rest=row
    out: list[str]=[]
    while True:
        item, rest=FindBracketedText2(rest, delim, caseInsensitive=True)
        if item == "":
            break
        if f"<{delim}>" in item:    # corrects the malformed-but-displayable pattern '<td>xxx<td>yyy</td>'
            out.extend(item.split(f"<{delim}>"))
        else:
            out.append(item)
    return out


# Mirrors ConSeriesFrame.ConNameInfoUnpack -> (name, url, extra)
def _ConNameInfoUnpack(packed: str) -> tuple[str, str, str]:
    name=html.unescape(packed)
    url=""
    extra=""
    m=re.match('<a href=\"?(.*?)\"?>(.*?)</a>(.*)$', packed, re.IGNORECASE)
    if m is not None:
        url=m.groups()[0].strip()
        name=html.unescape(m.groups()[1].strip())
        extra=html.unescape(m.groups()[2].strip())
    if extra == "":
        m=re.match(r"(.*)\((.*)\)$", name)
        if m is not None:
            if len(m.groups(2)) > 0:
                name=m.groups()[0]
                extra=f"({m.groups()[1]})"
    if url == f"{name}/index.html" or url == name:
        url="index.html"
    return name, url, extra


# The server directory name for a con series, derived from its root-index link (e.g. "./Worldcon/index.html").
def _SeriesDirFromLink(link: str) -> str:
    s=link.strip()
    if s.startswith("./"):
        s=s[2:]
    s=s.rstrip("/")
    if s.lower().endswith("/index.html"):
        s=s[:-len("/index.html")]
    elif s.lower() == "index.html":
        s=""
    return s


# The server directory for a con instance. Prefer the page's href -- the authoritative path the browser
# uses -- fully unescaped to undo HTML-escaping (single, or the accumulated multi-level escaping from the
# old CE round-trip bug: "Heicon &#x27;70", "Conspiracy &amp;amp;#x27;87"). Using the href also preserves
# the full "xxx (yyy)" directory name, which ConNameInfoUnpack would otherwise split into name + extra
# (e.g. "Thylacon 2005 (Thylacon IV)" -> name "Thylacon 2005 "), giving the wrong directory. When the
# link is collapsed to "index.html" (or absent), fall back to the accent-stripped name.
def _InstanceDir(name: str, url: str) -> str:
    u=url.strip()
    if u and u.lower() != "index.html":
        while True:                     # fully unescape (handles single and accumulated escaping)
            nu=html.unescape(u)
            if nu == u:
                break
            u=nu
        if u.lower().endswith("/index.html"):
            u=u[:-len("/index.html")]
        return u.rstrip("/")
    return RemoveAccents(name)


# Parse the conpubs root index.html into [(display_name, series_dir), ...].
def ParseSeriesLinks(rootHtml: str) -> list[tuple[str, str]]:
    out: list[tuple[str, str]]=[]
    table, _=FindBracketedText2(rootHtml, "fanac-table", caseInsensitive=True)
    tbody, _=FindBracketedText2(table, "tbody", caseInsensitive=True)
    while True:
        tr, tbody=FindBracketedText2(tbody, "tr", caseInsensitive=True)
        if tr == "":
            break
        td, _=FindBracketedText2(tr, "td", caseInsensitive=True)
        if "----" in td:
            continue        # skip divider rows
        _, link, text, _=FindLinkInString(td)
        sdir=_SeriesDirFromLink(link)
        if sdir:
            out.append((html.unescape(text), sdir))
    return out


# Parse a con series index.html into (series_display_name, [Instance, ...]).
def ParseInstances(seriesHtml: str) -> tuple[str, list[Instance]]:
    head, rest=FindBracketedText2(seriesHtml, "head", caseInsensitive=True)
    seriesName, _=FindBracketedText2(head, "title", caseInsensitive=True)

    table, _=FindBracketedText2(rest, "fanac-table", caseInsensitive=True)
    if table == "":
        return html.unescape(seriesName or ""), []

    header, rest=FindBracketedText2(table, "thead", caseInsensitive=True)
    headers=_ReadTableRow(header, "th")

    instances: list[Instance]=[]
    while True:
        rowtext, rest=FindBracketedText2(rest, "tr", caseInsensitive=True)
        if rowtext == "":
            break
        row=_ReadTableRow(rowtext, "td")
        if len(row) < len(headers):
            row.extend(" "*(len(headers)-len(row)))

        name=url=dates=""
        for icol, h in enumerate(headers):
            cell=row[icol] if isinstance(row[icol], str) else str(row[icol])
            if h == "Convention":
                name, url, _=_ConNameInfoUnpack(cell)
            elif h == "Dates":
                dates=html.unescape(cell)

        name=name.strip()
        if not name or name.startswith("---"):
            continue        # skip empty/divider rows
        instances.append(Instance(name=name, dir=_InstanceDir(name, url), dates=dates, linked=bool(url.strip())))

    return html.unescape(seriesName or ""), instances


# Copied from ConEditor/ConInstanceFrame.py (module-level there, but importing it would drag in the
# whole wx frame module); keep in sync.
def _CleanTitle(display_title: str) -> str:
    t=display_title.strip().removesuffix(".pdf").removesuffix(".PDF")
    # Expand the fannish abbreviations "PR"/"pr" -> "Progress Report" and "PB"/"pb" -> "Program Book".
    # Match a standalone PR/PB token (any case) whether or not a number follows and regardless of the
    # spacing: "PR 5", "PR5", "pr 5", "pr5" -> "...Report 5"; a bare "PR" -> "Progress Report".
    # The leading \b keeps it from matching inside words (e.g. "expr", "PROGRAM", "Spring").
    t=re.sub(r"\b[Pp][Rr] *(\d+)", r"Progress Report \1", t)   # PR/pr with a number (any spacing)
    t=re.sub(r"\b[Pp][Rr](?![A-Za-z0-9])", "Progress Report", t)   # bare PR/pr (no number)
    t=re.sub(r"\b[Pp][Bb] *(\d+)", r"Program Book \1", t)      # PB/pb with a number (any spacing)
    t=re.sub(r"\b[Pp][Bb](?![A-Za-z0-9])", "Program Book", t)      # bare PB/pb (no number)
    return t


# ---------- page-header construction ---------------------------------
# These mirror ConInstanceFrame._BuildPageHeader/_BuildSubPageHeader so the walker writes exactly
# the header ConEditor writes; see AddPdfPageHeader for the format/items conventions (a URL item is
# rendered as a hyperlink whose display text is the immediately-following item).

def _BuildPageHeader(sdir: str, inst: Instance, item_title: str) -> tuple[str, list]:
    instance_url = CONPUBS_ROOT_URL + "/".join(quote(s, safe='') for s in (sdir, inst.dir)) + "/"
    items = [instance_url, inst.name, item_title]
    if inst.dates.strip():
        fmt = "{}: {} ({})  --  from {}"
        items.append(inst.dates.strip())
    else:
        fmt = "{}: {}  --  from {}"
    items += [CONPUBS_ROOT_URL, "fanac.org/conpubs"]
    return fmt, items


def _BuildSubPageHeader(sdir: str, inst: Instance, sp_folder: str, sp_display: str, item_title: str) -> tuple[str, list]:
    sp_url = CONPUBS_ROOT_URL + "/".join(quote(s, safe='') for s in (sdir, inst.dir, sp_folder)) + "/"
    fmt = "{}: {}  --  from {}"
    items = [sp_url, sp_display, item_title, CONPUBS_ROOT_URL, "fanac.org/conpubs"]
    return fmt, items


# ---------- FTP / OCR helpers ----------------------------------------

def _download_pdf(ftp_path: str, local_path: str) -> str | None:
    """Return None on success, 'MISSING' if the file is absent (550), '' for other errors."""
    directory, filename = ftp_path.rsplit('/', 1)
    if not FTP().SetDirectory(directory or '/'):
        LogError(f"_download_pdf: directory not found on server: '{directory}'")
        return ""
    try:
        with open(local_path, 'wb') as f:
            msg = FTP.g_ftp.retrbinary(f"RETR {filename}", f.write)
        if not msg.startswith("226"):
            LogError(f"_download_pdf: unexpected response for '{filename}': {msg}")
            return ""
        return None
    except Exception as e:
        if '550' in str(e):
            LogError(f"_download_pdf: file not found on server: '{ftp_path}'")
            return "MISSING"
        LogError(f"_download_pdf: transfer failed for '{ftp_path}': {e}")
        return ""


def _upload_pdf(local_path: str, ftp_path: str) -> bool:
    directory, filename = ftp_path.rsplit('/', 1)
    if not FTP().SetDirectory(directory or '/'):
        Log(f"_upload_pdf: SetDirectory('{directory}') failed")
        return False
    return FTP().PutFile(local_path, filename)


def _ocr_pdf(input_path: str, output_path: str) -> bool:
    """Run ABBYY FineReader 15 on input_path, writing a searchable PDF to output_path.
    Returns True on success."""
    if not os.path.exists(FINECMD):
        LogError(f"_ocr_pdf: FineCmd not found at {FINECMD!r}")
        return False
    try:
        result = subprocess.run(
            [FINECMD, '/if', input_path, '/of', output_path, '/ofmt', 'PDF'],
            timeout=600,
            capture_output=True,
            text=True,
        )
        if result.returncode != 0:
            LogError(f"_ocr_pdf: FineCmd returned {result.returncode}; stderr={result.stderr.strip()!r}")
            return False
        if not os.path.exists(output_path) or os.path.getsize(output_path) == 0:
            LogError(f"_ocr_pdf: FineCmd exited 0 but output is missing or empty: {output_path!r}")
            return False
        return True
    except subprocess.TimeoutExpired:
        LogError(f"_ocr_pdf: FineCmd timed out for {input_path!r}")
        return False
    except Exception as e:
        LogError(f"_ocr_pdf: FineCmd failed: {e}")
        return False


# ---------- progress dialog ------------------------------------------

class WalkerDialog(wx.Dialog):
    def __init__(self):
        super().__init__(None, title="ConpubsWalker",
                         style=wx.DEFAULT_DIALOG_STYLE | wx.STAY_ON_TOP | wx.RESIZE_BORDER)
        self.stop_requested = False

        self._l1 = wx.StaticText(self, label="Headers checked: 0 (stamped 0)  --  OCR checked: 0 (OCR'd 0)")
        self._l2 = wx.StaticText(self, label="Step: Starting…")
        self._l3 = wx.StaticText(self, label="File: ")

        self._chk_hdr       = wx.CheckBox(self,    label="Update page headers")
        self._rb_hdr_check  = wx.RadioButton(self, label="Only when missing or incorrect", style=wx.RB_GROUP)
        self._rb_hdr_always = wx.RadioButton(self, label="Always re-stamp")
        self._chk_ocr       = wx.CheckBox(self,    label="OCR PDFs that need it")

        self._btn       = wx.Button(self, wx.ID_STOP,  "Stop")
        self._btn_clear = wx.Button(self, wx.ID_ANY,   "Clear State And Rerun")

        btn_row = wx.BoxSizer(wx.HORIZONTAL)
        btn_row.Add(self._btn,       0, wx.ALL, 6)
        btn_row.Add(self._btn_clear, 0, wx.ALL, 6)

        sep = wx.StaticLine(self, style=wx.LI_HORIZONTAL)

        sz = wx.BoxSizer(wx.VERTICAL)
        for ctrl in (self._l1, self._l2, self._l3):
            sz.Add(ctrl, 0, wx.ALL | wx.EXPAND, 6)
        sz.Add(sep, 0, wx.EXPAND | wx.LEFT | wx.RIGHT, 6)
        sz.Add(self._chk_hdr,       0, wx.ALL, 6)
        sz.Add(self._rb_hdr_check,  0, wx.LEFT, 28)
        sz.Add(self._rb_hdr_always, 0, wx.LEFT, 28)
        sz.Add(self._chk_ocr,       0, wx.ALL, 6)
        sz.Add(btn_row, 0, wx.ALIGN_CENTER)
        self.SetSizer(sz)
        self.Fit()                          # compute natural size
        self.SetMinSize(self.GetSize())     # don't allow shrinking below it

        self._restore_state()
        self._sync_radio_enable()

        self.Bind(wx.EVT_CHECKBOX, self._on_hdr_check, self._chk_hdr)
        self.Bind(wx.EVT_BUTTON, self._on_stop,  self._btn)
        self.Bind(wx.EVT_BUTTON, self._on_clear, self._btn_clear)
        self.Bind(wx.EVT_CLOSE,  self._on_close)

    # -- window state -------------------------------------------------

    @staticmethod
    def _on_any_screen(x: int, y: int) -> bool:
        """Return True if (x, y) falls inside any connected display."""
        for i in range(wx.Display.GetCount()):
            if wx.Display(i).GetGeometry().Contains(wx.Point(x, y)):
                return True
        return False

    def _restore_state(self):
        natural = self.GetSize()
        default_w = int(natural.width * 1.5)   # 50 % wider than natural fit

        if os.path.exists(STATE_FILE):
            try:
                with open(STATE_FILE, encoding='utf-8') as f:
                    s = json.load(f)
                self._chk_hdr.SetValue(s.get('update_headers', False))
                self._rb_hdr_always.SetValue(s.get('header_always', False))
                self._rb_hdr_check.SetValue(not s.get('header_always', False))
                self._chk_ocr.SetValue(s.get('ocr_if_needed', False))
                w = s.get('window', {})
                x, y     = w.get('x', None), w.get('y', None)
                width    = w.get('width',  default_w)
                height   = w.get('height', natural.height)
                if x is not None and y is not None and self._on_any_screen(x, y):
                    self.SetSize(width, height)
                    self.SetPosition((x, y))
                    return
            except Exception:
                pass

        self.SetSize(default_w, natural.height)
        self.Centre()

    def _save_state(self):
        pos, size = self.GetPosition(), self.GetSize()
        try:
            with open(STATE_FILE, 'w', encoding='utf-8') as f:
                json.dump({'window':         {'x': pos.x, 'y': pos.y,
                                              'width': size.width, 'height': size.height},
                           'update_headers': self._chk_hdr.GetValue(),
                           'header_always':  self._rb_hdr_always.GetValue(),
                           'ocr_if_needed':  self._chk_ocr.GetValue()},
                          f, indent=2)
        except Exception as e:
            Log(f"WalkerDialog: could not save state: {e}")

    def _sync_radio_enable(self):
        on = self._chk_hdr.GetValue()
        self._rb_hdr_check.Enable(on)
        self._rb_hdr_always.Enable(on)

    def _on_hdr_check(self, _):
        self._sync_radio_enable()

    @property
    def hdr_enabled(self) -> bool:
        return self._chk_hdr.GetValue()

    @property
    def hdr_always(self) -> bool:
        return self._rb_hdr_always.GetValue()

    @property
    def ocr_enabled(self) -> bool:
        return self._chk_ocr.GetValue()

    def _on_close(self, event):
        self._save_state()
        wx.GetApp().ExitMainLoop()
        event.Skip()

    # -- buttons ------------------------------------------------------

    def _on_clear(self, _):
        self._save_state()      # keep the checkbox settings across the restart
        LogClose()
        for path in (HDR_DONE_FILE, OCR_DONE_FILE,
                     os.path.join(_DIR, "Log -- ConpubsWalker.txt"),
                     os.path.join(_DIR, "Log (Errors) -- ConpubsWalker.txt")):
            if os.path.exists(path):
                os.remove(path)
        subprocess.Popen([sys.executable] + sys.argv)
        self.Close()

    def _on_stop(self, _):
        self.stop_requested = True
        self.Close()

    # -- update from worker thread ------------------------------------

    def update(self, *, line1: str = None, line2: str = None, line3: str = None):
        if line1 is not None:
            self._l1.SetLabel(line1)
        if line2 is not None:
            self._l2.SetLabel(f"Step: {line2}")
            font = self._l2.GetFont()
            font.SetWeight(wx.FONTWEIGHT_BOLD if 'Complete' in line2 else wx.FONTWEIGHT_NORMAL)
            self._l2.SetFont(font)
        if line3 is not None:
            self._l3.SetLabel(f"File: {line3}")
        self.Layout()


# ---------- walker ---------------------------------------------------

class Walker:

    def __init__(self, dlg: WalkerDialog):
        self._dlg = dlg
        self._hdr_done: set[str] = set()    # server paths already header-checked (OK or STAMPED)
        self._ocr_done: set[str] = set()    # server paths already OCR-checked (quality-tagged)
        self._hdr_checked = 0
        self._hdr_stamped = 0
        self._ocr_checked = 0
        self._ocr_fixed   = 0
        self._load_done_logs()

    # -- done-log persistence -----------------------------------------
    # Each log line is columns + two spaces + the PDF's server path. Only success statuses count as
    # done (a MISSING/failed PDF is retried next run). Loading tolerates the header line and blanks.

    def _load_done_logs(self):
        if os.path.exists(HDR_DONE_FILE):
            with open(HDR_DONE_FILE, encoding='utf-8') as f:
                for ln in f:
                    if len(ln) > _W_TAG + 2:
                        status = ln[:_W_TAG].strip()
                        path   = ln[_W_TAG + 2:].strip()
                        if path and status in ('OK', 'STAMPED'):
                            self._hdr_done.add(path)
        if os.path.exists(OCR_DONE_FILE):
            with open(OCR_DONE_FILE, encoding='utf-8') as f:
                for ln in f:
                    if len(ln) > _C_PATH:
                        tag  = ln[_C_TAG:_C_TAG + _W_TAG].strip()
                        path = ln[_C_PATH:].strip()
                        if path and tag in ('HIGH', 'LOW', 'NOT_OCRED'):
                            self._ocr_done.add(path)

    def _record_hdr(self, path: str, status: str):
        if status in ('OK', 'STAMPED'):
            self._hdr_done.add(path)
        needs_header = not os.path.exists(HDR_DONE_FILE) or os.path.getsize(HDR_DONE_FILE) == 0
        with open(HDR_DONE_FILE, 'a', encoding='utf-8') as f:
            if needs_header:
                f.write(_HDR_HEADER)
            f.write(f"{status:<{_W_TAG}}  {path}\n")

    def _record_ocr(self, path: str, tag: str, stats: dict):
        if tag in ('HIGH', 'LOW', 'NOT_OCRED'):
            self._ocr_done.add(path)
        alpha    = stats.get('alpha')
        in_words = stats.get('in_words')
        ratio    = stats.get('ratio')
        long_w   = stats.get('long_words')

        alpha_str    = "n/a" if alpha    is None else str(alpha)
        in_words_str = "n/a" if in_words is None else str(in_words)
        ratio_str    = "n/a" if ratio    is None else f"{ratio:.2f}"
        long_str     = "n/a" if long_w   is None else str(long_w)

        line = (f"{alpha_str:>{_W_ALPHA}}  {in_words_str:>{_W_WORDS}}  {ratio_str:>{_W_RATIO}}  "
                f"{long_str:>{_W_LONG}}  {tag:<{_W_TAG}}  {path}\n")
        needs_header = not os.path.exists(OCR_DONE_FILE) or os.path.getsize(OCR_DONE_FILE) == 0
        with open(OCR_DONE_FILE, 'a', encoding='utf-8') as f:
            if needs_header:
                f.write(_OCR_HEADER)
            f.write(line)

    # -- display ------------------------------------------------------

    def _show(self, step: str = None, filename: str = None):
        wx.CallAfter(self._dlg.update,
                     line1=f"Headers checked: {self._hdr_checked} (stamped {self._hdr_stamped})  --  "
                           f"OCR checked: {self._ocr_checked} (OCR'd {self._ocr_fixed})",
                     line2=step,
                     line3=filename)

    # -- PDF processing -----------------------------------------------

    # PDF text extraction returns the FONT's canonical characters, not the ones originally written:
    # an ASCII '-' stamped into the header comes back as U+2010 HYPHEN. Fold the hyphen/dash family
    # to '-' and NFKC-normalize (ligatures etc.) on BOTH sides so such differences don't read as an
    # "incorrect" header.
    _HYPHENS = dict.fromkeys(map(ord, '‐‑‒–—―'), '-')

    @classmethod
    def _normalize(cls, s: str) -> str:
        return " ".join(unicodedata.normalize('NFKC', s).translate(cls._HYPHENS).split())

    def _process_pdf(self, path: str, fmt: str, items: list):
        """Download one PDF (server path '/series/instance[/subpage]/file.pdf'), apply whichever of
        the OCR check and the header check are enabled and not already logged done, and re-upload
        only if the file changed. Logs each completed check; a failure logs nothing so it retries."""
        need_ocr = self._dlg.ocr_enabled and path not in self._ocr_done
        need_hdr = self._dlg.hdr_enabled and path not in self._hdr_done
        if not need_ocr and not need_hdr:
            return

        parts = [p for p in path.split('/') if p]
        name  = '/'.join(parts[-2:]) if len(parts) >= 2 else parts[-1]
        Log(f"Walker: _process_pdf {path!r}  (ocr={need_ocr}, hdr={need_hdr})")

        tmp_files: list[str] = []
        def _tmp() -> str:
            t = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
            t.close()
            tmp_files.append(t.name)
            return t.name

        try:
            tmp_in = _tmp()
            self._show("Downloading", name)
            err = _download_pdf(path, tmp_in)
            if err is not None:
                if err == 'MISSING':
                    # Recorded for the report but with a non-done status, so it is retried if it appears.
                    if need_hdr:
                        self._record_hdr(path, 'MISSING')
                return

            src      = tmp_in      # the copy both checks work on (and that an upload sends)
            modified = False

            # OCR first: ABBYY rewrites the whole PDF, which would destroy a header stamped before it.
            ocr_result = None      # (tag, stats) once the check completes
            ocr_fixed  = False
            if need_ocr:
                self._show("Assessing OCR quality", name)
                try:
                    reader         = PdfReader(src)
                    quality, stats = LowQualityScan(reader, label=name)
                    del reader
                except Exception as e:
                    LogError(f"Walker: PDF assess failed for {name}: {e}")
                    need_ocr = False
                    quality, stats = None, None
                if quality is not None:
                    Log(f"Walker: {name} -- initial OCR quality: {quality.name}")
                    if quality == OcrQuality.NOT_OCRED:
                        self._show("OCR-ing with ABBYY", name)
                        tmp_ocr = _tmp()
                        os.unlink(tmp_ocr)   # FineCmd refuses to overwrite an existing file
                        if _ocr_pdf(src, tmp_ocr):
                            self._show("Re-assessing OCR quality", name)
                            try:
                                reader2        = PdfReader(tmp_ocr)
                                quality, stats = LowQualityScan(reader2, label=f"{name}[OCR'd]")
                                del reader2
                            except Exception as e:
                                LogError(f"Walker: re-assess after OCR failed for {name}: {e}")
                            Log(f"Walker: {name} -- post-OCR quality: {quality.name}")
                            src      = tmp_ocr
                            modified = True
                            ocr_fixed  = True
                            ocr_result = (quality.name, stats)
                        else:
                            LogError(f"Walker: ABBYY OCR failed for {name} -- not recorded, will retry")
                            # No ocr_result: the OCR check did not complete. The header check can still run.
                    else:
                        ocr_result = (quality.name, stats)

            hdr_result = None      # status string once the check completes
            if need_hdr:
                expected = PdfPageHeaderDisplayText(fmt, items)
                stamp    = True
                was      = 'forced'
                if not self._dlg.hdr_always:
                    self._show("Checking page header", name)
                    try:
                        actual = GetPdfPageHeaderText(src)
                    except Exception as e:
                        LogError(f"Walker: header check failed for {name}: {e}")
                        actual = ""        # unreadable -> treat as failed; stamp nothing, record nothing
                        stamp  = None
                    if stamp is not None:
                        if actual is None:
                            was = 'missing'
                        elif self._normalize(actual) != self._normalize(expected):
                            was = 'differed'
                            Log(f"Walker: {name} -- header differs.\n  expected: {self._normalize(expected)!r}\n  found:    {self._normalize(actual)!r}")
                        else:
                            stamp      = False
                            hdr_result = 'OK'
                if stamp:
                    self._show("Stamping page header", name)
                    try:
                        AddPdfPageHeader(src, fmt, items, logo=_g_logo)
                        modified   = True
                        hdr_result = 'STAMPED'
                        Log(f"Walker: {name} -- header stamped ({was})")
                    except Exception as e:
                        LogError(f"Walker: header stamping failed for {name}: {e}")

            if modified:
                self._show("Uploading", name)
                try:
                    ok = _upload_pdf(src, path)
                except Exception as e:
                    LogError(f"Walker: upload exception for {name}: {e}")
                    ok = False
                if not ok:
                    LogError(f"Walker: upload failed -- {path}; nothing recorded, will retry")
                    return

            # Record (and count) only after any upload has succeeded, so a failed run retries from scratch.
            if ocr_result is not None:
                self._ocr_checked += 1
                if ocr_fixed:
                    self._ocr_fixed += 1
                self._record_ocr(path, *ocr_result)
            if hdr_result is not None:
                self._hdr_checked += 1
                if hdr_result == 'STAMPED':
                    self._hdr_stamped += 1
                self._record_hdr(path, hdr_result)
            self._show("Done", name)

        except Exception as e:
            LogError(f"Walker: unexpected error processing {name}: {e}")
        finally:
            for p in tmp_files:
                if os.path.exists(p):
                    os.remove(p)

    # -- tree walk ----------------------------------------------------

    def _walk_instance_rows(self, ci: ConInstance, dirpath: str, header_for) -> None:
        """Process the PDF rows of one downloaded ConInstance. dirpath is its server dir
        ('/series/instance' or '/series/instance/subpage'); header_for(title) -> (fmt, items)."""
        for row in ci.ConInstanceRows:
            if self._dlg.stop_requested:
                return
            if row.IsTextRow or row.IsLinkRow or row.IsSubPageRow:
                continue
            site = (row.SiteFilename or "").strip()
            if not site.lower().endswith('.pdf'):
                continue
            fmt, items = header_for(_CleanTitle(row.DisplayTitle))
            self._process_pdf(f"{dirpath}/{site}", fmt, items)

    def _walk_instance(self, sdir: str, inst: Instance) -> None:
        instpath = f"/{sdir}/{inst.dir}"
        self._show("Reading instance page", instpath)
        try:
            ci = ConInstance(f"/{sdir}", sdir, inst.dir)
            if not ci.Download():
                LogError(f"Walker: instance index download/parse failed: {instpath}/index.html")
                return
        except Exception as e:
            LogError(f"Walker: instance index error for {instpath}: {e!r}")
            return

        self._walk_instance_rows(ci, instpath, lambda title: _BuildPageHeader(sdir, inst, title))

        # Sub-pages: a sub-page row's SiteFilename is its sub-folder (empty = not created yet).
        for row in ci.ConInstanceRows:
            if self._dlg.stop_requested:
                return
            if not row.IsSubPageRow:
                continue
            sp_folder = (row.SiteFilename or "").strip()
            if not sp_folder:
                continue
            sp_display = row.DisplayTitle.strip() or sp_folder
            sp_path    = f"{instpath}/{sp_folder}"
            self._show("Reading sub-page", sp_path)
            try:
                sp = ConInstance(instpath, sdir, sp_folder)
                if not sp.Download():
                    LogError(f"Walker: sub-page index download/parse failed: {sp_path}/index.html")
                    continue
            except Exception as e:
                LogError(f"Walker: sub-page index error for {sp_path}: {e!r}")
                continue
            self._walk_instance_rows(sp, sp_path,
                                     lambda title: _BuildSubPageHeader(sdir, inst, sp_folder, sp_display, title))

    def walk(self) -> None:
        self._show("Reading conpubs root index")
        try:
            rootHtml = FTP().GetFileAsString("", "index.html")
        except Exception as e:
            LogError(f"Walker: root index download failed: {e!r}")
            return
        if not rootHtml:
            LogError("Walker: root index download returned nothing")
            return

        allSeries = ParseSeriesLinks(rootHtml)      # [(display_name, series_dir), ...]
        if LIMIT_SERIES:
            lim = LIMIT_SERIES.lower()
            seriesDirs = [d for (disp, d) in allSeries if lim in (disp.lower(), d.lower())]
            if not seriesDirs:
                LogError(f"Walker: no con series matching '{LIMIT_SERIES}' in the root index")
                return
        else:
            seriesDirs = [d for (_, d) in allSeries]
        Log(f"Walker: root index lists {len(allSeries)} con series; walking {len(seriesDirs)}")

        for sdir in seriesDirs:
            if self._dlg.stop_requested:
                return
            self._show("Reading series page", f"/{sdir}")
            try:
                seriesHtml = FTP().GetFileAsString(f"/{sdir}", "index.html")
            except Exception as e:
                LogError(f"Walker: series index download failed for /{sdir}: {e!r}")
                continue
            if not seriesHtml:
                LogError(f"Walker: series index download returned nothing for /{sdir}")
                continue

            _, instances = ParseInstances(seriesHtml)
            linked = [i for i in instances if i.linked]
            Log(f"Walker: series /{sdir}: {len(instances)} instance(s), {len(linked)} linked")

            for inst in linked:
                if self._dlg.stop_requested:
                    return
                if LIMIT_INSTANCE and inst.dir != LIMIT_INSTANCE and inst.name != LIMIT_INSTANCE:
                    continue
                self._walk_instance(sdir, inst)


# ---------- app ------------------------------------------------------

class ConpubsWalkerApp(wx.App):

    def OnInit(self):
        global _g_logo
        os.chdir(_DIR)
        LogOpen("Log -- ConpubsWalker.txt", "Log (Errors) -- ConpubsWalker.txt")

        if not os.path.exists(FTP_CREDS):
            wx.MessageBox(f"FTP credentials not found:\n{FTP_CREDS}",
                          "ConpubsWalker")
            return False

        if not FTP().OpenConnection(FTP_CREDS):
            wx.MessageBox(f"Cannot open FTP connection.\n"
                          f"Credentials: {FTP_CREDS}",
                          "ConpubsWalker — FTP error")
            return False

        # The page-header logo, shared with ConEditor. Missing/unreadable is non-fatal: headers
        # simply carry no logo (matching ConEditor's behavior).
        try:
            with open(LOGO_FILE, "rb") as f:
                _g_logo = f.read()
        except Exception as e:
            Log(f"ConpubsWalker: could not load '{LOGO_FILE}'; headers will have no logo: {e}")

        self._dlg = WalkerDialog()
        self._dlg.Show()
        threading.Thread(target=self._run, daemon=True).start()
        return True

    def _run(self):
        Walker(self._dlg).walk()
        if self._dlg.stop_requested:
            wx.CallAfter(self._dlg.update, line2="Stopped")
        else:
            wx.CallAfter(self._dlg.update, line2="Complete — all pages processed")


if __name__ == '__main__':
    ConpubsWalkerApp().MainLoop()
