# -*- coding: utf-8 -*-
"""
PDF 용량 줄이기

원하는 최대 용량을 지정하면, 그 이하가 될 때까지 단계적으로 PDF를 압축합니다.
모든 처리는 이 PC 안에서만 이루어지며 파일이 외부로 전송되지 않습니다.

이 프로그램은 자유 소프트웨어이며 GNU Affero General Public License v3.0
조건에 따라 배포됩니다. 어떠한 보증도 제공되지 않습니다.
PDF 처리에는 Artifex Software의 PyMuPDF(AGPL-3.0)를 사용합니다.
"""

import json
import os
import queue
import shutil
import subprocess
import sys
import tempfile
import threading
import webbrowser
import tkinter as tk
import tkinter.font as tkfont
from tkinter import filedialog, messagebox, ttk

# 창 없는(windowed) exe에서는 표준 출력이 없어 라이브러리 경고가 오류를 일으킬 수 있음
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w")

import fitz  # PyMuPDF

# ---------------------------------------------------------------------------
# 배포 정보 (배포 전에 아래 두 값을 본인 정보로 바꿔 주세요)
# ---------------------------------------------------------------------------
APP_NAME = "PDF 용량 줄이기"
APP_VERSION = "1.0.0"
AUTHOR = "작성자 이름"
SOURCE_URL = "https://github.com/사용자명/저장소명"

# ---------------------------------------------------------------------------
# 압축 설정 (위에서 아래로 갈수록 강하게 압축, 목표를 만족하는 첫 단계에서 멈춤)
# ---------------------------------------------------------------------------
# (목표 해상도 dpi, JPEG 품질)
IMAGE_LEVELS = [
    (300, 85), (250, 80), (200, 75), (170, 70), (150, 65), (130, 60),
    (110, 55), (96, 50), (80, 45), (72, 40), (60, 30),
]
RASTER_LEVELS = [
    (200, 80), (170, 75), (150, 70), (130, 65), (110, 60), (96, 55),
    (80, 50), (72, 45), (60, 40), (50, 30),
]
MIN_TARGET_BYTES = 10 * 1024

STAGE_LABELS = {
    1: "1단계  무손실 최적화  (화질 변화 없음)",
    2: "2단계  이미지 압축  (이미지 화질만 낮춤, 텍스트 유지)",
    3: "3단계  페이지를 이미지로 변환  (가장 강력, 텍스트 선택 불가)",
}

try:
    HAS_REWRITE = hasattr(fitz.Document, "rewrite_images")
except Exception:
    HAS_REWRITE = False

try:
    fitz.TOOLS.mupdf_display_errors(False)
    fitz.TOOLS.mupdf_display_warnings(False)
except Exception:
    pass


class Cancelled(Exception):
    pass


# ---------------------------------------------------------------------------
# 유틸리티
# ---------------------------------------------------------------------------
def fmt_size(n):
    if n >= 1024 * 1024:
        return f"{n / 1024 / 1024:.2f} MB"
    return f"{n / 1024:.1f} KB"


def resource_path(name):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, name)


def settings_path():
    base = os.environ.get("APPDATA") or os.path.expanduser("~")
    return os.path.join(base, "PDFCompressor", "settings.json")


def load_settings():
    try:
        with open(settings_path(), "r", encoding="utf-8") as f:
            data = json.load(f)
            return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def save_settings(data):
    try:
        p = settings_path()
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with open(p, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def make_dst(src, out_dir):
    folder = out_dir or os.path.dirname(os.path.abspath(src))
    os.makedirs(folder, exist_ok=True)
    base = os.path.splitext(os.path.basename(src))[0]
    cand = os.path.join(folder, f"{base}_compressed.pdf")
    n = 1
    while os.path.exists(cand):
        cand = os.path.join(folder, f"{base}_compressed({n}).pdf")
        n += 1
    return cand


def open_path(path):
    try:
        if sys.platform == "win32":
            os.startfile(path)  # noqa
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except Exception:
        pass


# ---------------------------------------------------------------------------
# 압축 엔진
# ---------------------------------------------------------------------------
def _open(src):
    doc = fitz.open(src)
    if doc.needs_pass:
        doc.close()
        raise ValueError("암호가 걸린 PDF는 처리할 수 없습니다.")
    return doc


def save_lossless(src, dst):
    doc = _open(src)
    try:
        doc.save(dst, garbage=4, deflate=True, clean=True)
    finally:
        doc.close()


def save_images(src, dst, dpi, quality):
    doc = _open(src)
    try:
        doc.rewrite_images(dpi_threshold=dpi + 10, dpi_target=dpi, quality=quality)
        doc.save(dst, garbage=4, deflate=True, clean=True)
    finally:
        doc.close()


def save_raster(src, dst, dpi, quality):
    doc = _open(src)
    out = fitz.open()
    try:
        for page in doc:
            pix = page.get_pixmap(dpi=dpi)
            data = pix.tobytes("jpeg", jpg_quality=quality)
            new_page = out.new_page(width=page.rect.width, height=page.rect.height)
            new_page.insert_image(new_page.rect, stream=data)
        out.save(dst, garbage=4, deflate=True)
    finally:
        out.close()
        doc.close()


def compress_file(src, dst, target, stages, cancel, log, step):
    """
    src를 target 바이트 이하로 압축해 dst에 저장.
    stages: 사용할 단계의 집합 (1, 2, 3 중 선택). 선택한 단계를 1 → 2 → 3 순서로 시도.
    반환: {"status": ok|missed|skipped|nochange|error|cancelled, "size": int, "dst": str, "msg": str}
    """
    orig = os.path.getsize(src)
    log(f"  원본 {fmt_size(orig)}")

    if orig <= target:
        log("  이미 목표 용량 이하라 건너뜁니다.")
        return {"status": "skipped", "size": orig}

    try:
        d = _open(src)
        log(f"  {d.page_count}페이지")
        d.close()
    except Exception as e:
        msg = str(e) if isinstance(e, ValueError) else "PDF를 열 수 없습니다. 손상된 파일일 수 있어요."
        log(f"  실패: {msg}")
        return {"status": "error", "msg": msg}

    use1 = 1 in stages
    use2 = 2 in stages and HAS_REWRITE
    use3 = 3 in stages
    if 2 in stages and not HAS_REWRITE:
        log("  2단계는 설치된 PyMuPDF 버전에서 지원하지 않아 건너뜁니다.")
    total = max(1, (1 if use1 else 0)
                + (len(IMAGE_LEVELS) if use2 else 0)
                + (len(RASTER_LEVELS) if use3 else 0))
    state = {"done": 0, "best": None, "mode": "", "reached": False, "errors": 0}
    tmpdir = tempfile.mkdtemp(prefix="pdfcomp_")
    tmp = os.path.join(tmpdir, "t.pdf")
    best = os.path.join(tmpdir, "best.pdf")

    def attempt(mode, label, fn, *args):
        if cancel.is_set():
            raise Cancelled()
        step(state["done"] / total, label)
        try:
            fn(src, tmp, *args)
            size = os.path.getsize(tmp)
        except Exception as e:
            state["errors"] += 1
            state["done"] += 1
            log(f"  {label} → 실패 ({e})")
            return False
        state["done"] += 1
        log(f"  {label} → {fmt_size(size)}")
        if size < orig and (state["best"] is None or size < state["best"]):
            state["best"] = size
            state["mode"] = mode
            shutil.copyfile(tmp, best)
        if size <= target:
            state["reached"] = True
        return state["reached"]

    try:
        if use1:
            attempt("lossless", "1단계 무손실 최적화", save_lossless)

        if not state["reached"] and use2:
            for dpi, q in IMAGE_LEVELS:
                if attempt("images", f"2단계 이미지 압축 ({dpi}dpi, 품질 {q})", save_images, dpi, q):
                    break

        if not state["reached"] and use3:
            log("  페이지를 이미지로 변환합니다. (텍스트 선택과 검색이 안 됩니다)")
            for dpi, q in RASTER_LEVELS:
                if attempt("raster", f"3단계 페이지 이미지화 ({dpi}dpi, 품질 {q})", save_raster, dpi, q):
                    break

        if state["done"] == 0:
            return {"status": "error", "msg": "선택한 단계를 실행할 수 없습니다."}

        if state["best"] is None:
            if state["errors"] >= 1 and state["done"] == state["errors"]:
                return {"status": "error", "msg": "압축 중 오류가 발생했습니다."}
            log("  더 줄일 수 없는 파일입니다.")
            return {"status": "nochange", "size": orig}

        shutil.copyfile(best, dst)
        return {
            "status": "ok" if state["reached"] else "missed",
            "size": state["best"],
            "dst": dst,
            "mode": state["mode"],
        }
    except Cancelled:
        return {"status": "cancelled"}
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)


# ---------------------------------------------------------------------------
# 화면
# ---------------------------------------------------------------------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} {APP_VERSION}")
        try:
            self.iconbitmap(resource_path("app.ico"))
        except Exception:
            pass
        for name in ("TkDefaultFont", "TkTextFont", "TkMenuFont"):
            try:
                tkfont.nametofont(name).configure(family="Malgun Gothic", size=9)
            except Exception:
                pass
        self.minsize(540, 720)

        self.files = []
        self.cancel = threading.Event()
        self.q = queue.Queue()
        self.running = False
        self.last_outputs = []

        cfg = load_settings()
        self.size_var = tk.StringVar(value=str(cfg.get("size", "5")))
        self.unit_var = tk.StringVar(value=cfg.get("unit", "MB"))
        saved = cfg.get("stages", [1, 2, 3])
        if not isinstance(saved, list):
            saved = [1, 2, 3]
        self.stage_vars = {n: tk.BooleanVar(value=(n in saved)) for n in (1, 2, 3)}
        self.outmode = tk.StringVar(value=cfg.get("outmode", "same"))
        self.outdir = tk.StringVar(value=cfg.get("outdir", ""))
        self.status = tk.StringVar(value="PDF 파일을 추가해 주세요.")

        self._build()
        self.update_out_state()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(100, self.poll)

    def _build(self):
        pad = {"padx": 12, "pady": 5}
        root = ttk.Frame(self)
        root.pack(fill="both", expand=True, pady=(6, 8))
        self.controls = []

        # 파일 목록
        box = ttk.LabelFrame(root, text="PDF 파일", padding=8)
        box.pack(fill="x", **pad)
        lf = ttk.Frame(box)
        lf.pack(side="left", fill="both", expand=True)
        self.listbox = tk.Listbox(lf, height=6, selectmode="extended", activestyle="none")
        sb = ttk.Scrollbar(lf, orient="vertical", command=self.listbox.yview)
        self.listbox.configure(yscrollcommand=sb.set)
        self.listbox.pack(side="left", fill="both", expand=True)
        sb.pack(side="left", fill="y")
        bf = ttk.Frame(box)
        bf.pack(side="left", padx=(8, 0), anchor="n")
        for text, cmd in (("추가", self.add_files), ("제거", self.remove_selected), ("비우기", self.clear_files)):
            b = ttk.Button(bf, text=text, command=cmd, width=8)
            b.pack(pady=2)
            self.controls.append(b)

        # 목표 용량
        tb = ttk.LabelFrame(root, text="목표 용량", padding=8)
        tb.pack(fill="x", **pad)
        row = ttk.Frame(tb)
        row.pack(fill="x")
        self.size_entry = ttk.Entry(row, textvariable=self.size_var, width=9)
        self.size_entry.pack(side="left")
        self.unit_combo = ttk.Combobox(row, textvariable=self.unit_var, values=["MB", "KB"], width=5, state="readonly")
        self.unit_combo.pack(side="left", padx=6)
        ttk.Label(row, text="이하로 줄이기").pack(side="left")
        self.controls.append(self.size_entry)

        # 압축 단계 선택
        sg = ttk.LabelFrame(root, text="사용할 압축 단계", padding=8)
        sg.pack(fill="x", **pad)
        for n in (1, 2, 3):
            cb = ttk.Checkbutton(sg, text=STAGE_LABELS[n], variable=self.stage_vars[n])
            cb.pack(anchor="w", pady=1)
            self.controls.append(cb)
        ttk.Label(sg, text="선택한 단계를 순서대로 시도하고, 목표 용량에 도달하면 멈춥니다.",
                  foreground="gray").pack(anchor="w", pady=(4, 0))

        # 저장 위치
        ob = ttk.LabelFrame(root, text="저장 위치", padding=8)
        ob.pack(fill="x", **pad)
        self.rb_same = ttk.Radiobutton(ob, text="원본과 같은 폴더 (파일명_compressed.pdf)",
                                       variable=self.outmode, value="same", command=self.update_out_state)
        self.rb_same.pack(anchor="w")
        orow = ttk.Frame(ob)
        orow.pack(fill="x", pady=(4, 0))
        self.rb_dir = ttk.Radiobutton(orow, text="지정 폴더", variable=self.outmode, value="dir",
                                      command=self.update_out_state)
        self.rb_dir.pack(side="left")
        self.out_entry = ttk.Entry(orow, textvariable=self.outdir)
        self.out_entry.pack(side="left", fill="x", expand=True, padx=6)
        self.out_btn = ttk.Button(orow, text="찾기", width=6, command=self.pick_outdir)
        self.out_btn.pack(side="left")
        self.controls += [self.rb_same, self.rb_dir]

        # 실행 버튼
        ab = ttk.Frame(root)
        ab.pack(fill="x", **pad)
        self.btn_start = ttk.Button(ab, text="압축 시작", command=self.start)
        self.btn_start.pack(side="left")
        self.btn_cancel = ttk.Button(ab, text="취소", command=self.cancel_run, state="disabled")
        self.btn_cancel.pack(side="left", padx=6)
        ttk.Button(ab, text="정보", command=self.show_about, width=6).pack(side="right")
        self.btn_open = ttk.Button(ab, text="저장 폴더 열기", command=self.open_output, state="disabled")
        self.btn_open.pack(side="right", padx=6)

        # 진행 상황
        self.bar = ttk.Progressbar(root, mode="determinate", maximum=100)
        self.bar.pack(fill="x", padx=12, pady=(4, 2))
        ttk.Label(root, textvariable=self.status, anchor="w").pack(fill="x", padx=12)

        # 로그
        lg = ttk.Frame(root)
        lg.pack(fill="both", expand=True, padx=12, pady=(6, 0))
        self.logbox = tk.Text(lg, height=9, state="disabled", wrap="word")
        ls = ttk.Scrollbar(lg, orient="vertical", command=self.logbox.yview)
        self.logbox.configure(yscrollcommand=ls.set)
        self.logbox.pack(side="left", fill="both", expand=True)
        ls.pack(side="left", fill="y")

    # -- 상태 ---------------------------------------------------------------
    def update_out_state(self):
        st = "normal" if self.outmode.get() == "dir" and not self.running else "disabled"
        self.out_entry.configure(state=st)
        self.out_btn.configure(state=st)

    def set_running(self, flag):
        self.running = flag
        st = "disabled" if flag else "normal"
        for w in self.controls:
            w.configure(state=st)
        self.unit_combo.configure(state="disabled" if flag else "readonly")
        self.listbox.configure(state=st)
        self.btn_start.configure(state=st)
        self.btn_cancel.configure(state="normal" if flag else "disabled")
        self.update_out_state()

    def log(self, msg):
        self.logbox.configure(state="normal")
        self.logbox.insert("end", msg + "\n")
        self.logbox.see("end")
        self.logbox.configure(state="disabled")

    # -- 파일 목록 -----------------------------------------------------------
    def refresh_list(self):
        self.listbox.delete(0, "end")
        for p in self.files:
            try:
                size = fmt_size(os.path.getsize(p))
            except OSError:
                size = "?"
            self.listbox.insert("end", f"{os.path.basename(p)}   ({size})")
        n = len(self.files)
        self.status.set(f"{n}개 파일 선택됨" if n else "PDF 파일을 추가해 주세요.")

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="PDF 파일 선택", filetypes=[("PDF 파일", "*.pdf"), ("모든 파일", "*.*")])
        for p in paths:
            p = os.path.normpath(p)
            if p.lower().endswith(".pdf") and p not in self.files:
                self.files.append(p)
        self.refresh_list()

    def remove_selected(self):
        for i in sorted(self.listbox.curselection(), reverse=True):
            del self.files[i]
        self.refresh_list()

    def clear_files(self):
        self.files = []
        self.refresh_list()

    def pick_outdir(self):
        p = filedialog.askdirectory(title="저장 폴더 선택")
        if p:
            self.outdir.set(os.path.normpath(p))

    # -- 실행 ---------------------------------------------------------------
    def start(self):
        if not self.files:
            messagebox.showwarning(APP_NAME, "압축할 PDF 파일을 먼저 추가해 주세요.")
            return
        try:
            val = float(self.size_var.get().replace(",", ".").strip())
        except ValueError:
            val = 0
        if val <= 0:
            messagebox.showwarning(APP_NAME, "목표 용량을 올바른 숫자로 입력해 주세요.")
            return
        target = int(val * (1024 * 1024 if self.unit_var.get() == "MB" else 1024))
        if target < MIN_TARGET_BYTES:
            messagebox.showwarning(APP_NAME, "목표 용량은 10 KB 이상으로 입력해 주세요.")
            return
        stages = {n for n, v in self.stage_vars.items() if v.get()}
        if not stages:
            messagebox.showwarning(APP_NAME, "사용할 압축 단계를 하나 이상 선택해 주세요.")
            return
        out_dir = None
        if self.outmode.get() == "dir":
            out_dir = self.outdir.get().strip()
            if not out_dir:
                messagebox.showwarning(APP_NAME, "저장할 폴더를 지정해 주세요.")
                return

        save_settings({
            "size": self.size_var.get(), "unit": self.unit_var.get(),
            "stages": sorted(stages), "outmode": self.outmode.get(),
            "outdir": self.outdir.get(),
        })

        self.cancel.clear()
        self.last_outputs = []
        self.btn_open.configure(state="disabled")
        self.logbox.configure(state="normal")
        self.logbox.delete("1.0", "end")
        self.logbox.configure(state="disabled")
        self.bar["value"] = 0
        self.set_running(True)
        threading.Thread(
            target=self.work,
            args=(list(self.files), target, stages, out_dir),
            daemon=True,
        ).start()

    def cancel_run(self):
        if self.running:
            self.cancel.set()
            self.status.set("취소하는 중...")

    def work(self, files, target, stages, out_dir):
        q = self.q
        counts = {"ok": 0, "missed": 0, "skipped": 0, "nochange": 0, "error": 0}
        outputs = []
        cancelled = False
        n = len(files)
        try:
            for i, src in enumerate(files):
                if self.cancel.is_set():
                    cancelled = True
                    break
                q.put(("log", f"\n[{i + 1}/{n}] {os.path.basename(src)}"))

                def step(frac, text, i=i):
                    q.put(("progress", (i + frac) / n, f"[{i + 1}/{n}] {text}"))

                def log(msg):
                    q.put(("log", msg))

                try:
                    dst = make_dst(src, out_dir)
                    res = compress_file(src, dst, target, stages, self.cancel, log, step)
                except Exception as e:
                    res = {"status": "error", "msg": str(e)}
                    q.put(("log", f"  실패: {e}"))

                st = res["status"]
                if st == "cancelled":
                    cancelled = True
                    break
                counts[st] = counts.get(st, 0) + 1
                if st in ("ok", "missed"):
                    outputs.append(res["dst"])
                    if st == "ok":
                        q.put(("log", f"  완료: {fmt_size(res['size'])}"))
                    else:
                        rest = [str(s) for s in (2, 3) if s not in stages]
                        hint = f" {', '.join(rest)}단계도 함께 선택하면 더 줄일 수 있어요." if rest else ""
                        q.put(("log", f"  목표에 도달하지 못했습니다. 가장 작은 결과 {fmt_size(res['size'])}로 저장했습니다.{hint}"))
                    if res.get("mode") == "raster":
                        q.put(("log", "  참고: 이 파일은 텍스트 선택과 검색이 되지 않습니다."))
                    q.put(("log", f"  저장: {dst}"))
                q.put(("progress", (i + 1) / n, f"[{i + 1}/{n}] 처리 완료"))
        except Exception as e:
            q.put(("log", f"오류: {e}"))
        q.put(("done", {"counts": counts, "outputs": outputs, "cancelled": cancelled}))

    def poll(self):
        try:
            while True:
                kind, *payload = self.q.get_nowait()
                if kind == "log":
                    self.log(payload[0])
                elif kind == "progress":
                    self.bar["value"] = payload[0] * 100
                    self.status.set(payload[1])
                elif kind == "done":
                    self.finish(payload[0])
        except queue.Empty:
            pass
        self.after(100, self.poll)

    def finish(self, summary):
        self.set_running(False)
        self.last_outputs = summary["outputs"]
        if self.last_outputs:
            self.btn_open.configure(state="normal")
        c = summary["counts"]
        parts = []
        for key, label in (("ok", "성공"), ("missed", "목표 미달"), ("skipped", "건너뜀"),
                           ("nochange", "더 줄일 수 없음"), ("error", "실패")):
            if c.get(key):
                parts.append(f"{label} {c[key]}개")
        text = ", ".join(parts) if parts else "처리한 파일이 없습니다"
        if summary["cancelled"]:
            self.status.set(f"취소됨 ({text})")
            self.log("\n사용자가 취소했습니다.")
        else:
            self.bar["value"] = 100
            self.status.set(f"완료: {text}")
            self.log(f"\n완료: {text}")
            messagebox.showinfo(APP_NAME, f"작업이 끝났습니다.\n{text}")

    def open_output(self):
        if self.last_outputs:
            open_path(os.path.dirname(self.last_outputs[0]))

    def on_close(self):
        if self.running:
            if not messagebox.askyesno(APP_NAME, "압축이 진행 중입니다. 취소하고 종료할까요?"):
                return
            self.cancel.set()
        self.destroy()

    def show_about(self):
        win = tk.Toplevel(self)
        win.title("정보")
        win.resizable(False, False)
        win.transient(self)
        text = (
            f"{APP_NAME}  v{APP_VERSION}\n"
            f"만든 사람: {AUTHOR}\n\n"
            "모든 처리는 이 PC 안에서만 이루어지며,\n"
            "파일이 외부로 전송되지 않습니다.\n\n"
            "이 프로그램은 자유 소프트웨어이며\n"
            "GNU Affero General Public License v3.0 조건에 따라\n"
            "배포됩니다. 어떠한 보증도 제공되지 않습니다.\n\n"
            "PDF 처리에 Artifex Software의 PyMuPDF(AGPL-3.0)를 사용합니다.\n"
            "소스코드는 아래 주소에서 받을 수 있습니다.\n"
            f"{SOURCE_URL}"
        )
        ttk.Label(win, text=text, justify="left", padding=16).pack()
        bf = ttk.Frame(win, padding=(16, 0, 16, 14))
        bf.pack(fill="x")
        ttk.Button(bf, text="소스코드 열기", command=lambda: webbrowser.open(SOURCE_URL)).pack(side="left")
        ttk.Button(bf, text="닫기", command=win.destroy).pack(side="right")
        win.grab_set()


def main():
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    App().mainloop()


if __name__ == "__main__":
    main()
