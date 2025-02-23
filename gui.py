import os
import queue
import subprocess
import threading
import tkinter as tk
import webbrowser
from tkinter import messagebox
from tkinter import ttk

from tkinterdnd2 import TkinterDnD, DND_FILES

from file_processor import process_file

AUTHOR = "운양고등학교 이종환T"
VERSION = "2025.02.21."


class DragDropApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        self.title("Google Forms to PDF")
        self.geometry("500x500")
        self.configure(bg="white")
        self.create_menu()

        default_font = tkFont.nametofont("TkDefaultFont")
        default_font.configure(family="Arial", size=12)

        # 드래그 앤 드롭 영역
        self.drop_area = tk.Label(
            self,
            text="\n".join([
                "여기에 Google Forms 응답결과",
                "CSV 파일을 Drag & Drop 하세요"
            ]),
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

    def create_menu(self):
        """메뉴바 생성"""
        menubar = tk.Menu(self)

        # About 메뉴 생성
        about_menu = tk.Menu(menubar, tearoff=0)
        about_menu.add_command(label="프로그램 정보", command=self.show_program_info)
        about_menu.add_command(label="GitHub 바로가기", command=self.open_github)
        menubar.add_cascade(label="About", menu=about_menu)

        # 메뉴바를 메인 윈도우에 추가
        self.config(menu=menubar)

    def show_program_info(self):
        info_title = "프로그램 정보"
        info_message = (
            f"제작자: {AUTHOR}\n"
            f"버전: {VERSION}\n"
            "\n"
            "이 프로그램은 구글 폼 설문지 CSV 결과를 PDF로 변환하는 Python 프로그램입니다."
        )

        messagebox.showinfo(info_title, info_message)

    def open_github(self):
        webbrowser.open("https://github.com/itmir913/google-forms-to-pdf-gui/releases")


if __name__ == "__main__":
    app = DragDropApp()
    app.bind_progress_event()  # 프로그래스 업데이트 이벤트 바인딩
    app.mainloop()
