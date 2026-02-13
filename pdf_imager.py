"""
PDFImager - PDFを1ページごとに画像化して保存する簡易アプリ
"""

import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import fitz  # PyMuPDF


class PDFImagerApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PDFImager")
        self.root.geometry("620x420")
        self.root.resizable(False, False)

        self.pdf_path = tk.StringVar()
        self.output_dir = tk.StringVar()
        self.dpi = tk.IntVar(value=200)
        self.fmt = tk.StringVar(value="PNG")
        self.is_running = False

        self._build_ui()

    # ── UI 構築 ──────────────────────────────────────────────
    def _build_ui(self):
        pad = {"padx": 10, "pady": 5}

        # --- PDFファイル選択 ---
        frame_pdf = ttk.LabelFrame(self.root, text="PDFファイル")
        frame_pdf.pack(fill="x", **pad)

        self.entry_pdf = ttk.Entry(frame_pdf, textvariable=self.pdf_path, width=60)
        self.entry_pdf.pack(side="left", padx=(10, 5), pady=8, fill="x", expand=True)

        ttk.Button(frame_pdf, text="参照…", command=self._browse_pdf).pack(
            side="right", padx=(0, 10), pady=8
        )

        # --- 出力先フォルダ ---
        frame_out = ttk.LabelFrame(self.root, text="出力先フォルダ")
        frame_out.pack(fill="x", **pad)

        self.entry_out = ttk.Entry(frame_out, textvariable=self.output_dir, width=60)
        self.entry_out.pack(side="left", padx=(10, 5), pady=8, fill="x", expand=True)

        ttk.Button(frame_out, text="参照…", command=self._browse_output).pack(
            side="right", padx=(0, 10), pady=8
        )

        # --- オプション ---
        frame_opt = ttk.LabelFrame(self.root, text="オプション")
        frame_opt.pack(fill="x", **pad)

        ttk.Label(frame_opt, text="DPI:").pack(side="left", padx=(10, 2), pady=8)
        spin_dpi = ttk.Spinbox(
            frame_opt, from_=72, to=600, increment=50, textvariable=self.dpi, width=6
        )
        spin_dpi.pack(side="left", padx=(0, 20), pady=8)

        ttk.Label(frame_opt, text="形式:").pack(side="left", padx=(0, 2), pady=8)
        combo_fmt = ttk.Combobox(
            frame_opt,
            textvariable=self.fmt,
            values=["PNG", "JPEG"],
            state="readonly",
            width=6,
        )
        combo_fmt.pack(side="left", padx=(0, 10), pady=8)

        # --- プログレスバー & ログ ---
        self.progress = ttk.Progressbar(self.root, mode="determinate")
        self.progress.pack(fill="x", **pad)

        self.label_status = ttk.Label(self.root, text="待機中")
        self.label_status.pack(**pad)

        self.text_log = tk.Text(self.root, height=6, state="disabled", font=("Consolas", 9))
        self.text_log.pack(fill="both", expand=True, **pad)

        # --- 実行ボタン ---
        self.btn_run = ttk.Button(
            self.root, text="変換開始", command=self._start_conversion
        )
        self.btn_run.pack(pady=(0, 10))

    # ── ファイル / フォルダ選択 ──────────────────────────────
    def _browse_pdf(self):
        path = filedialog.askopenfilename(
            title="PDFファイルを選択",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
        )
        if path:
            self.pdf_path.set(path)
            # 出力先が未設定なら PDF と同じフォルダを設定
            if not self.output_dir.get():
                self.output_dir.set(os.path.dirname(path))

    def _browse_output(self):
        path = filedialog.askdirectory(title="出力先フォルダを選択")
        if path:
            self.output_dir.set(path)

    # ── ログ出力 ─────────────────────────────────────────────
    def _log(self, msg: str):
        self.text_log.configure(state="normal")
        self.text_log.insert("end", msg + "\n")
        self.text_log.see("end")
        self.text_log.configure(state="disabled")

    # ── 変換処理 ─────────────────────────────────────────────
    def _start_conversion(self):
        pdf = self.pdf_path.get().strip()
        out = self.output_dir.get().strip()

        if not pdf or not os.path.isfile(pdf):
            messagebox.showerror("エラー", "有効なPDFファイルを選択してください。")
            return
        if not out:
            messagebox.showerror("エラー", "出力先フォルダを指定してください。")
            return
        if self.is_running:
            return

        os.makedirs(out, exist_ok=True)
        self.is_running = True
        self.btn_run.configure(state="disabled")

        # 別スレッドで変換 (UIフリーズ防止)
        thread = threading.Thread(
            target=self._convert, args=(pdf, out), daemon=True
        )
        thread.start()

    def _convert(self, pdf_path: str, output_dir: str):
        try:
            doc = fitz.open(pdf_path)
            total = len(doc)
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            ext = self.fmt.get().lower()  # png or jpeg
            dpi = self.dpi.get()
            zoom = dpi / 72.0
            mat = fitz.Matrix(zoom, zoom)

            self.root.after(0, lambda: self.progress.configure(maximum=total, value=0))
            self.root.after(0, lambda: self._log(f"変換開始: {total} ページ  (DPI={dpi}, 形式={ext.upper()})"))

            for i, page in enumerate(doc, start=1):
                pix = page.get_pixmap(matrix=mat)
                file_ext = "jpg" if ext == "jpeg" else "png"
                out_file = os.path.join(output_dir, f"{base_name}_page{i:04d}.{file_ext}")
                pix.save(out_file)

                self.root.after(0, lambda v=i, f=out_file: self._update_progress(v, total, f))

            doc.close()
            self.root.after(0, lambda: self._conversion_done(total, output_dir))

        except Exception as e:
            self.root.after(0, lambda: self._conversion_error(str(e)))

    def _update_progress(self, current: int, total: int, filepath: str):
        self.progress["value"] = current
        self.label_status.configure(text=f"変換中… {current}/{total}")
        self._log(f"  [{current}/{total}] {os.path.basename(filepath)}")

    def _conversion_done(self, total: int, output_dir: str):
        self.label_status.configure(text="完了")
        self._log(f"✔ 全 {total} ページの変換が完了しました。出力先: {output_dir}")
        self.is_running = False
        self.btn_run.configure(state="normal")
        messagebox.showinfo("完了", f"{total} ページの画像を保存しました。")

    def _conversion_error(self, err: str):
        self.label_status.configure(text="エラー")
        self._log(f"✖ エラー: {err}")
        self.is_running = False
        self.btn_run.configure(state="normal")
        messagebox.showerror("エラー", err)


# ── エントリポイント ─────────────────────────────────────────
def main():
    root = tk.Tk()
    PDFImagerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
