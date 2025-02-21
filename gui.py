import os
import queue
import subprocess
import threading
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk

from tkinterdnd2 import TkinterDnD, DND_FILES

from file_processor import process_file


class DragDropApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        self.title("Google Forms to PDF")
        self.geometry("500x500")
        self.configure(bg="white")

        # 드래그 앤 드롭 영역
        self.drop_area = tk.Label(
            self,
            text="여기에 CSV 파일을 드래그하세요",
            relief="solid",
            bg="white",
            fg="black",
            font=("Arial", 12),
            anchor="center",
            justify="center",
        )
        self.drop_area.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)  # 여백 추가

        # 프로그래스 바 설정
        self.progress = ttk.Progressbar(self, length=300, mode='determinate', maximum=100, value=0)
        self.progress.place(relx=0.5, rely=0.85, anchor="center")

        # 드래그 앤 드롭 이벤트 연결
        self.drop_area.drop_target_register(DND_FILES)
        self.drop_area.dnd_bind('<<Drop>>', self.on_drop)

    def on_drop(self, event):
        file_path = event.data.strip('{}')
        if file_path.endswith('.csv'):
            # 프로그래스 바 시작
            self.update()

            # 큐 설정 (메인 스레드에서 값 받기)
            self.progress_queue = queue.Queue()

            # 파일 처리 작업을 비동기적으로 실행
            threading.Thread(target=self.process_file_in_thread,
                             args=(file_path, self.progress_queue),
                             daemon=True).start()

            # 프로그래스 바 업데이트 및 메시지 박스 호출
            self.handle_progress_signal()

    def process_file_in_thread(self, file_path, progress_queue):
        # 파일 드롭 후, "처리중입니다."로 텍스트 변경
        self.progress['value'] = 0
        self.drop_area.config(text="처리중입니다.")

        # 파일 처리 및 프로그레스 업데이트
        output_pdf_path = process_file(file_path, lambda value: progress_queue.put(value))

        # 처리 완료 후 메시지 박스 호출
        self.progress['value'] = 100
        self.drop_area.config(text="완료되었습니다.")
        self.open_pdf(output_pdf_path)

    def open_pdf(self, pdf_path):
        """자동으로 PDF 파일을 기본 뷰어로 실행하는 함수"""
        try:
            if os.name == 'nt':
                os.startfile(pdf_path)
            else:
                subprocess.run(['open', pdf_path], check=True)  # Linux 및 macOS에서도 'open' 사용
        except Exception as e:
            messagebox.showerror("오류", f"PDF 파일을 열 수 없습니다: {e}")

    def handle_progress_signal(self):
        # 이벤트를 발생시켜 프로그래스 바를 업데이트
        self.event_generate("<<ProgressUpdate>>")

    def on_progress_update(self, event):
        try:
            value = self.progress_queue.get_nowait()
            self.progress['value'] = value
            self.after(200, self.handle_progress_signal)  # 100ms 후에 다시 시도
        except queue.Empty:
            # 큐가 비어 있으면 다시 시도
            self.after(200, self.handle_progress_signal)

    def bind_progress_event(self):
        # 프로그래스 업데이트 이벤트 바인딩
        self.bind("<<ProgressUpdate>>", self.on_progress_update)


if __name__ == "__main__":
    app = DragDropApp()
    app.bind_progress_event()  # 프로그래스 업데이트 이벤트 바인딩
    app.mainloop()
