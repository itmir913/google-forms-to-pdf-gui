import subprocess
import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES
from tkinter import messagebox
from tkinter import ttk
from file_processor import process_file
import os


class DragDropApp(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()

        self.title("Google Forms to PDF")
        self.geometry("500x500")
        self.configure(bg="#f0f0f0")  # 전체 배경색 설정

        # 드래그 앤 드롭 영역 레이아웃 (전체 창 크기)
        self.drop_area = tk.Label(self, text="여기에 CSV 파일을 드래그하세요", relief="solid", bg="#f5f5f5", fg="black",
                                  font=("Arial", 12), anchor="center", justify="center")
        self.drop_area.pack(fill=tk.BOTH, expand=True)  # 전체 창을 채우도록 설정

        # 프로그래스 바 설정
        self.progress = ttk.Progressbar(self, length=300, mode='determinate', maximum=100, value=0)
        self.progress.place(relx=0.5, rely=0.85, anchor="center")  # 하단 중앙에 배치

        # 드래그 앤 드롭 이벤트 연결
        self.drop_area.drop_target_register(DND_FILES)
        self.drop_area.dnd_bind('<<Drop>>', self.on_drop)

    def on_drop(self, event):
        file_path = event.data
        if file_path.endswith('.csv'):
            # 프로그래스 바 시작
            self.progress.start()
            self.update()


            # 파일 처리 로직 실행
            output_pdf_path = process_file(file_path, self.update_progress)

            # 프로그래스 바 완료
            self.progress.stop()
            messagebox.showinfo("완료", "PDF 생성이 완료되었습니다!")

            if os.name == 'nt':
                os.startfile(output_pdf_path)
            else:
                subprocess.run(['open', output_pdf_path], check=True)  # Linux 및 macOS에서도 'open' 사용

    def update_progress(self, value):
        # 프로그래스 바의 값 업데이트
        self.progress['value'] = value
        self.update()


if __name__ == "__main__":
    app = DragDropApp()
    app.mainloop()
