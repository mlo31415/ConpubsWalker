from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import threading
from html.parser import HTMLParser
from urllib.parse import urljoin, urlparse, quote, unquote
from urllib.request import urlopen, Request
from urllib.error import URLError

import wx

# ---------- path setup -----------------------------------------------
# Run from ConpubsWalker dir; ConEditor is a sibling directory with all shared modules.
_DIR        = os.path.dirname(os.path.abspath(__file__))
_CONEDITOR  = os.path.normpath(os.path.join(_DIR, '..', 'ConEditor'))
_HELPERSPKG = os.path.normpath(os.path.join(_DIR, '..', 'HelpersPackage'))
sys.path.insert(0, _CONEDITOR)
sys.path.insert(0, _HELPERSPKG)

from pypdf import PdfReader, PdfWriter
from PDFHelpers import LowQualityScan, OcrQuality
from FTP import FTP
from Log import Log, LogOpen, LogClose, LogError

# ---------- configuration --------------------------------------------
START_URL      = "https://fanac.org/conpubs/xxxTest/index.html"
CONPUBS_BASE   = "https://fanac.org/conpubs/"
FTP_CREDS      = os.path.join(_CONEDITOR, "FTP Credentials.json")
COMPLETED_FILE = os.path.join(_DIR, "ConpubsWalker Completed.json")
PROCESSED_FILE = os.path.join(_DIR, "ConpubsWalker Processed.txt")
STATE_FILE     = os.path.join(_DIR, "ConpubsWalker State.json")
FINECMD        = r"C:\Program Files (x86)\ABBYY FineReader 15\FineCmd.exe"
_W_ALPHA = 11   # "alpha count"  column — right-justified
_W_WORDS = 10   # "# in words"   column — right-justified
_W_RATIO =  5   # "ratio"        column — right-justified
_W_LONG  = 12   # "words>5 char" column — right-justified
_W_TAG   = 10   # "Status"       column — left-justified  (NOT_OCRED = 9)
_C_TAG   = _W_ALPHA + 2 + _W_WORDS + 2 + _W_RATIO + 2 + _W_LONG + 2   # = 46
_C_URL   = _C_TAG + _W_TAG + 2                                           # = 58
_HEADER  = (f"{'alpha count':>{_W_ALPHA}}  {'# in words':>{_W_WORDS}}  "
            f"{'ratio':>{_W_RATIO}}  {'words>5 char':>{_W_LONG}}  "
            f"{'Status':<{_W_TAG}}  pdf and path\n")
_BASE_LIMIT    = START_URL.rsplit('/', 1)[0] + '/'   # strip down to series folder
_EMPTY_STATS   = {'alpha': 0,    'in_words': 0,    'ratio': None, 'long_words': 0}
_MISSING_STATS = {'alpha': None, 'in_words': None, 'ratio': None, 'long_words': None}
HTTP_HEADERS   = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                  'AppleWebKit/537.36 (KHTML, like Gecko) '
                  'Chrome/120.0.0.0 Safari/537.36'
}

# ---------- HTML link extraction -------------------------------------

class _LinkExtractor(HTMLParser):
    def __init__(self):
        super().__init__()
        self.links: list[str] = []

    def handle_starttag(self, tag, attrs):
        if tag == 'a':
            for name, val in attrs:
                if name == 'href' and val:
                    self.links.append(val)


def _encode_url(url: str) -> str:
    """Percent-encode spaces and other illegal characters in the URL path."""
    p = urlparse(url)
    return p._replace(path=quote(p.path, safe='/')).geturl()


def _fetch_page(url: str) -> str | None:
    try:
        req = Request(_encode_url(url), headers=HTTP_HEADERS)
        with urlopen(req, timeout=30) as resp:
            return resp.read().decode('utf-8', errors='replace')
    except Exception as e:
        Log(f"ConpubsWalker: fetch failed for {url}: {e}")
        return None


def _page_links(html: str, base_url: str) -> list[str]:
    p = _LinkExtractor()
    p.feed(html)
    # Strip #fragment suffixes (e.g. #view=Fit) — they're browser hints, not part of the file URL.
    return [urljoin(base_url, h).split('#')[0] for h in p.links]


# ---------- FTP helpers ----------------------------------------------

def _url_to_ftp_path(url: str) -> str:
    """https://fanac.org/conpubs/X/Y/z.pdf  →  /X/Y/z.pdf"""
    path = unquote(urlparse(url).path)  # decode %23→# %20→space etc.
    if path.startswith('/conpubs'):
        path = path[len('/conpubs'):]
    return path or '/'


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

        self._l1 = wx.StaticText(self, label="Processed: 0 this run, 0 total")
        self._l2 = wx.StaticText(self, label="Step: Starting…")
        self._l3 = wx.StaticText(self, label="File: ")
        self._chk_ocr   = wx.CheckBox(self, label="OCR PDFs if needed")
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
        sz.Add(self._chk_ocr, 0, wx.ALL, 6)
        sz.Add(btn_row, 0, wx.ALIGN_CENTER)
        self.SetSizer(sz)
        self.Fit()                          # compute natural size
        self.SetMinSize(self.GetSize())     # don't allow shrinking below it

        self._restore_state()

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
                json.dump({'window':       {'x': pos.x, 'y': pos.y,
                                            'width': size.width, 'height': size.height},
                           'ocr_if_needed': self._chk_ocr.GetValue()},
                          f, indent=2)
        except Exception as e:
            Log(f"WalkerDialog: could not save state: {e}")

    @property
    def ocr_enabled(self) -> bool:
        return self._chk_ocr.GetValue()

    def _on_close(self, event):
        self._save_state()
        wx.GetApp().ExitMainLoop()
        event.Skip()

    # -- stop button --------------------------------------------------

    def _on_clear(self, _):
        LogClose()
        for path in (COMPLETED_FILE, PROCESSED_FILE,
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
        self._dlg       = dlg
        self._completed: set[str] = set()
        self._processed: set[str] = set()   # quality-tagged (HIGH/LOW/NOT_OCRED) — won't re-process
        self._seen:      set[str] = set()   # all URLs recorded in Processed.txt
        self._run_count = 0
        self._load_state()

    # -- state persistence --------------------------------------------

    def _load_state(self):
        if os.path.exists(COMPLETED_FILE):
            with open(COMPLETED_FILE, encoding='utf-8') as f:
                self._completed = set(json.load(f).get("completed", []))
        if os.path.exists(PROCESSED_FILE):
            with open(PROCESSED_FILE, encoding='utf-8') as f:
                lines = f.readlines()
            valid = []
            for ln in lines:
                if not ln.strip():
                    continue
                # 6-column format: display path at _C_URL, tag at _C_TAG
                if len(ln) > _C_URL:
                    tag     = ln[_C_TAG:_C_TAG + _W_TAG].strip()
                    display = ln[_C_URL:].strip()
                    if display:
                        # Reconstruct full URL (backward-compat: keep if already full URL)
                        if display.startswith('http'):
                            full_url = display
                        else:
                            full_url = _BASE_LIMIT + quote(display, safe='/')
                        valid.append(ln)
                        self._seen.add(full_url)
                        if tag in ('HIGH', 'LOW', 'NOT_OCRED'):
                            self._processed.add(full_url)
                        continue
                # Header line or unrecognized format — keep header, drop old data rows
                if ln.strip().startswith('alpha count'):
                    valid.append(ln)
            if len(valid) != len([ln for ln in lines if ln.strip()]):
                with open(PROCESSED_FILE, 'w', encoding='utf-8') as f:
                    f.writelines(valid)

    def _save_completed(self):
        with open(COMPLETED_FILE, 'w', encoding='utf-8') as f:
            json.dump({"completed": sorted(self._completed)}, f, indent=2)

    def _record_pdf(self, url: str, tag: str, stats: dict):
        self._seen.add(url)
        if tag in ('HIGH', 'LOW', 'NOT_OCRED'):
            self._processed.add(url)
        alpha    = stats.get('alpha')
        in_words = stats.get('in_words')
        ratio    = stats.get('ratio')
        long_w   = stats.get('long_words')

        alpha_str    = "n/a" if alpha    is None else str(alpha)
        in_words_str = "n/a" if in_words is None else str(in_words)
        ratio_str    = "n/a" if ratio    is None else f"{ratio:.2f}"
        long_str     = "n/a" if long_w   is None else str(long_w)

        display = unquote(url[len(_BASE_LIMIT):]) if url.startswith(_BASE_LIMIT) else url

        line = (f"{alpha_str:>{_W_ALPHA}}  {in_words_str:>{_W_WORDS}}  {ratio_str:>{_W_RATIO}}  "
                f"{long_str:>{_W_LONG}}  {tag:<{_W_TAG}}  {display}\n")
        needs_header = not os.path.exists(PROCESSED_FILE) or os.path.getsize(PROCESSED_FILE) == 0
        with open(PROCESSED_FILE, 'a', encoding='utf-8') as f:
            if needs_header:
                f.write(_HEADER)
            f.write(line)

    # -- display ------------------------------------------------------

    def _show(self, step: str = None, filename: str = None):
        wx.CallAfter(self._dlg.update,
                     line1=f"Processed: {self._run_count} this run, "
                           f"{len(self._processed)} total",
                     line2=step,
                     line3=filename)

    # -- PDF processing -----------------------------------------------

    def _process_pdf(self, pdf_url: str) -> tuple[str, dict]:
        """Process one PDF and return (tag, stats). tag is quality name, 'MISSING', or ''."""
        ftp_path = _url_to_ftp_path(pdf_url)
        parts    = [p for p in ftp_path.split('/') if p]
        name     = '/'.join(parts[-2:]) if len(parts) >= 2 else parts[-1]
        Log(f"Walker: _process_pdf url={pdf_url!r}  ftp={ftp_path!r}")

        tmp_files: list[str] = []
        def _tmp() -> str:
            t = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
            t.close()
            tmp_files.append(t.name)
            return t.name

        try:
            try:
                tmp_in = _tmp()
                self._show("Downloading", name)
                err = _download_pdf(ftp_path, tmp_in)
                if err is not None:
                    return err, (_MISSING_STATS if err == 'MISSING' else _EMPTY_STATS)

                self._show("Assessing OCR quality", name)
                try:
                    reader         = PdfReader(tmp_in)
                    quality, stats = LowQualityScan(reader, label=name)
                    del reader
                except Exception as e:
                    LogError(f"Walker: PDF assess failed for {name}: {e}")
                    return "", _EMPTY_STATS
                Log(f"Walker: {name} — initial OCR quality: {quality.name}")

                upload_src = None   # the file we'll actually upload

                if quality == OcrQuality.NOT_OCRED and self._dlg.ocr_enabled:
                    self._show("OCR-ing with ABBYY", name)
                    tmp_ocr = _tmp()
                    os.unlink(tmp_ocr)   # FineCmd refuses to overwrite an existing file
                    if _ocr_pdf(tmp_in, tmp_ocr):
                        self._show("Re-assessing OCR quality", name)
                        try:
                            reader2        = PdfReader(tmp_ocr)
                            quality, stats = LowQualityScan(reader2, label=f"{name}[OCR'd]")
                            del reader2
                        except Exception as e:
                            LogError(f"Walker: re-assess after OCR failed for {name}: {e}")
                        Log(f"Walker: {name} — post-OCR quality: {quality.name}")
                        upload_src = tmp_ocr
                    else:
                        LogError(f"Walker: ABBYY OCR failed for {name} — skipping upload")
                        return "", _EMPTY_STATS

                if upload_src is None:
                    # No OCR performed — clone through PdfWriter as usual
                    self._show(f"Writing PDF [{quality.name}]", name)
                    tmp_out = _tmp()
                    try:
                        PdfWriter(clone_from=tmp_in).write(tmp_out)
                    except Exception as e:
                        LogError(f"Walker: PDF write failed for {name}: {e}")
                        return "", _EMPTY_STATS
                    upload_src = tmp_out

                self._show("Uploading", name)
                try:
                    ok = _upload_pdf(upload_src, ftp_path)
                except Exception as e:
                    LogError(f"Walker: upload exception for {name}: {e}")
                    return "", _EMPTY_STATS
                if not ok:
                    LogError(f"Walker: upload failed — {ftp_path}")
                    return "", _EMPTY_STATS

                self._run_count += 1
                self._show("Done", name)
                Log(f"Walker: processed {pdf_url} — {quality.name}")
                return quality.name, stats

            except Exception as e:
                LogError(f"Walker: unexpected error processing {name}: {e}")
                return "", _EMPTY_STATS

        finally:
            for p in tmp_files:
                if os.path.exists(p):
                    os.remove(p)

    # -- tree walk ----------------------------------------------------

    def walk(self, url: str, base_limit: str):
        if self._dlg.stop_requested:
            return
        if url in self._completed:
            return
        if not url.startswith(base_limit):
            return

        self._show("Fetching page", url.replace(CONPUBS_BASE, ''))
        html = _fetch_page(url)
        if html is None:
            return

        all_links = _page_links(html, url)
        Log(f"Walker: {url} — {len(all_links)} links found")

        pdf_links = sorted(set(
            l for l in all_links
            if urlparse(l).path.lower().endswith('.pdf') and l.startswith(base_limit)
        ))
        sub_links = sorted(set(
            l for l in all_links
            if l.endswith('index.html') and l.startswith(base_limit)
            and l != url and len(l) > len(url)
        ))
        Log(f"Walker:   {len(pdf_links)} PDFs, {len(sub_links)} sub-pages")
        for l in pdf_links:
            Log(f"Walker:     PDF: {l}")
        for l in sub_links:
            Log(f"Walker:     Sub: {l}")

        for pdf_url in pdf_links:
            if self._dlg.stop_requested:
                return
            if pdf_url not in self._seen:
                tag, stats = self._process_pdf(pdf_url)
                self._record_pdf(pdf_url, tag, stats)

        for sub in sub_links:
            if self._dlg.stop_requested:
                return
            self.walk(sub, base_limit)

        # Mark this page complete only after all its PDFs and sub-pages are done.
        if not self._dlg.stop_requested:
            self._completed.add(url)
            self._save_completed()


# ---------- app ------------------------------------------------------

class ConpubsWalkerApp(wx.App):

    def OnInit(self):
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

        self._dlg = WalkerDialog()
        self._dlg.Show()
        threading.Thread(target=self._run, daemon=True).start()
        return True

    def _run(self):
        base_limit = START_URL.rsplit('/', 1)[0] + '/'
        Walker(self._dlg).walk(START_URL, base_limit)
        if self._dlg.stop_requested:
            wx.CallAfter(self._dlg.update, line2="Stopped")
        else:
            wx.CallAfter(self._dlg.update, line2="Complete — all pages processed")


if __name__ == '__main__':
    ConpubsWalkerApp().MainLoop()
