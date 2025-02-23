import os
import shutil
import tempfile

import PyPDF2
import pandas as pd
import pdfkit
from jinja2 import Template


# Jinja2 템플릿 정의
def load_template(template_path):
    with open(template_path, "r", encoding="utf-8") as file:
        return file.read()


def process_file(file_path, update_progress, batch_size):
    print(f"파일 처리 시작: {file_path}")

    # CSV 파일 불러오기
    df = pd.read_csv(file_path, header=None)
    df = df.loc[:, df.iloc[0].notna()]
    df.columns = df.iloc[0]
    df = df.drop(0, axis=0)
    df = df.fillna("No answer")

    # 파일의 디렉토리 경로와 파일 이름 추출
    folder_path = os.path.dirname(file_path)
    file_name = os.path.splitext(os.path.basename(file_path))[0]  # 확장자 제외한 파일 이름

    # 최종 PDF 파일 이름 생성
    output_pdf_path = os.path.join(folder_path, f"output_{file_name}.pdf")

    # 실행 파일이 있는 디렉토리 경로 찾기
    base_path = os.path.abspath(".")

    # wkhtmltopdf 실행 파일 경로
    wkhtmltopdf_path = shutil.which("wkhtmltopdf") or os.path.join(base_path, "bin", "wkhtmltopdf.exe")

    # pdfkit 설정
    config = pdfkit.configuration(wkhtmltopdf=wkhtmltopdf_path)
    options = {
        "encoding": "UTF-8",
        "enable-local-file-access": "",
        "custom-header": [("Accept-Encoding", "gzip")],
        "no-stop-slow-scripts": "",
        "load-error-handling": "ignore",
        "load-media-error-handling": "ignore"
    }

    template_path = os.path.join(base_path, "bin", "template.html")
    html_template = load_template(template_path)

    # 중간 저장을 위해 파일을 열고 진행
    with open(output_pdf_path, "wb") as output_pdf:
        pdf_writer = PyPDF2.PdfWriter()

        # 각 응답자를 PDF로 처리 후 저장
        for idx, respondent in enumerate(df.to_dict(orient="records")):
            # HTML 렌더링
            template = Template(html_template)
            html_content = template.render(respondent=respondent)

            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_pdf:
                temp_pdf_path = temp_pdf.name
                pdfkit.from_string(html_content, temp_pdf_path, configuration=config, options=options)

                pdf_reader = PyPDF2.PdfReader(temp_pdf)
                num_pages = len(pdf_reader.pages)

                # 응답자 PDF 페이지를 output_pdf에 바로 기록
                for page_num in range(num_pages):
                    pdf_writer.add_page(pdf_reader.pages[page_num])

                # 응답자 PDF에 빈 페이지 추가
                pages_to_add = batch_size - (num_pages % batch_size) if num_pages % batch_size != 0 else 0
                for _ in range(pages_to_add):
                    pdf_writer.add_blank_page(width=pdf_reader.pages[0].mediabox[2],
                                              height=pdf_reader.pages[0].mediabox[3])

                # 바로 결과를 파일에 저장
                pdf_writer.write(output_pdf)

            # 처리한 임시 PDF 파일 삭제
            if os.path.exists(temp_pdf_path):
                os.remove(temp_pdf_path)

            update_progress((idx + 1) / len(df) * 100)

    print("PDF 생성 완료!")
    return output_pdf_path
