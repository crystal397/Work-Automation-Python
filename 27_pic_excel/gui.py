"""
사진대지 자동 삽입 GUI
======================
Tkinter 기반 간단한 GUI. 엑셀 파일 + 사진 폴더를 선택해서 실행 버튼만 누르면 끝.
기존 insert_photos.py는 손대지 않고, 그 insert_photos() 함수의 print 출력을
stdout 리다이렉트로 가로채서 로그창에 표시한다.

실행:
    python gui.py

exe 빌드:
    pyinstaller --onefile --windowed --name="사진대지삽입" gui.py
"""

import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from insert_photos import insert_photos


class QueueWriter:
    """print() 출력을 Queue에 담아 UI 스레드에서 소비할 수 있게 하는 래퍼."""

    def __init__(self, q):
        self.q = q

    def write(self, s):
        if s:
            self.q.put(s)

    def flush(self):
        pass


class PhotoInsertApp:
    def __init__(self, root):
        self.root = root
        root.title("사진대지 자동 삽입")
        root.geometry("760x560")

        self.log_queue = queue.Queue()

        self.excel_var = tk.StringVar()
        self._build_path_row("엑셀 파일:", self.excel_var, self._pick_excel)

        self.photo_var = tk.StringVar()
        self._build_path_row("사진 폴더:", self.photo_var, self._pick_photo_dir)

        self.output_var = tk.StringVar()
        self._build_path_row("출력 파일:", self.output_var, self._pick_output,
                             hint="(비워두면 자동 생성)")

        btn_frame = ttk.Frame(root, padding=(10, 5))
        btn_frame.pack(fill="x")
        self.run_btn = ttk.Button(btn_frame, text="사진 삽입 실행",
                                  command=self._start_run)
        self.run_btn.pack(side="left")
        ttk.Button(btn_frame, text="로그 지우기",
                   command=lambda: self.log_area.delete("1.0", "end")
                   ).pack(side="left", padx=5)

        self.log_area = scrolledtext.ScrolledText(root, height=18,
                                                  font=("Consolas", 10))
        self.log_area.pack(fill="both", expand=True, padx=10, pady=(5, 10))

        self.root.after(100, self._drain_log_queue)

    def _build_path_row(self, label, var, picker, hint=None):
        frame = ttk.Frame(self.root, padding=(10, 5))
        frame.pack(fill="x")
        ttk.Label(frame, text=label, width=10).pack(side="left")
        ttk.Entry(frame, textvariable=var).pack(
            side="left", fill="x", expand=True, padx=5)
        ttk.Button(frame, text="찾아보기", command=picker).pack(side="left")
        if hint:
            ttk.Label(frame, text=hint, foreground="gray").pack(
                side="left", padx=5)

    def _pick_excel(self):
        path = filedialog.askopenfilename(
            title="엑셀 파일 선택",
            filetypes=[("엑셀 파일", "*.xlsx *.xlsm"), ("모든 파일", "*.*")],
        )
        if path:
            self.excel_var.set(path)

    def _pick_photo_dir(self):
        path = filedialog.askdirectory(title="사진 폴더 선택")
        if path:
            self.photo_var.set(path)

    def _pick_output(self):
        path = filedialog.asksaveasfilename(
            title="출력 파일 저장 위치",
            defaultextension=".xlsx",
            filetypes=[("엑셀 파일", "*.xlsx")],
        )
        if path:
            self.output_var.set(path)

    def _drain_log_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_area.insert("end", msg)
                self.log_area.see("end")
        except queue.Empty:
            pass
        self.root.after(100, self._drain_log_queue)

    def _start_run(self):
        excel = self.excel_var.get().strip()
        photo = self.photo_var.get().strip()
        output = self.output_var.get().strip() or None

        if not excel or not Path(excel).is_file():
            messagebox.showerror("오류", "엑셀 파일을 올바르게 선택하세요.")
            return
        if not photo or not Path(photo).is_dir():
            messagebox.showerror("오류", "사진 폴더를 올바르게 선택하세요.")
            return

        self.run_btn.config(state="disabled")
        self.log_area.delete("1.0", "end")

        thread = threading.Thread(
            target=self._worker, args=(excel, photo, output), daemon=True)
        thread.start()

    def _worker(self, excel, photo, output):
        writer = QueueWriter(self.log_queue)
        old_stdout, old_stderr = sys.stdout, sys.stderr
        sys.stdout = writer
        sys.stderr = writer
        try:
            insert_photos(excel, photo, output)
            self.root.after(0, lambda: messagebox.showinfo(
                "완료", "사진 삽입이 완료되었습니다."))
        except Exception as e:
            err = str(e)
            self.log_queue.put(f"\n❌ 오류: {err}\n")
            self.root.after(0, lambda: messagebox.showerror("오류", err))
        finally:
            sys.stdout, sys.stderr = old_stdout, old_stderr
            self.root.after(0, lambda: self.run_btn.config(state="normal"))


def main():
    root = tk.Tk()
    PhotoInsertApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
