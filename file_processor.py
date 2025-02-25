import os
import shutil
import tempfile

import PyPDF2
import chardet
import pandas as pd
import pdfkit
from jinja2 import Template


# Jinja2 템플릿 정의
def load_template(template_path):
    with open(template_path, "r", encoding="utf-8") as file:
        return file.read()


def detect_csv_encoding(file_path):
    try:
        with open(file_path, 'rb') as f:
            result = chardet.detect(f.read())
        encoding = result['encoding']
        print(f"Detected encoding: {encoding}")
        return encoding
    except Exception as e:
        print(f"Encoding detection failed: {e}")
        return 'utf-8'


def count_valid_pages(pdf_path):
    # PDF 파일 열기
    with open(pdf_path, "rb") as file:
        pdf_reader = PyPDF2.PdfReader(file)
        num_pages = len(pdf_reader.pages)

        valid_pages = 0

        # 각 페이지의 텍스트 추출 후 텍스트가 있는지 검사
        for page_num in range(num_pages):
            page = pdf_reader.pages[page_num]
            text = page.extract_text()

            # 텍스트가 있으면 유효한 페이지로 카운트
            if text.strip():  # strip()으로 공백을 제거하고 텍스트가 있으면
                valid_pages += 1

        return valid_pages


def process_file(file_path, update_progress, batch_size):
    print(f"파일 처리 시작: {file_path}")

    # CSV 파일 불러오기
    df = pd.read_csv(file_path,
                     header=None,
                     encoding=detect_csv_encoding(file_path))
    df = df.loc[:, df.iloc[0].notna()]
    df.columns = df.iloc[0]
    df = df.drop(0, axis=0)
    df = df.dropna(how='all')
    df = df.fillna("No answer")
    df = df.apply(lambda col: col.map(lambda x: x.strip() if isinstance(x, str) else x))

    # 파일의 디렉토리 경로와 파일 이름 추출
    folder_path = os.path.dirname(file_path)
    file_name = os.path.splitext(os.path.basename(file_path))[0]

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
        "load-media-error-handling": "ignore",
        "no-images": "",  # Prevent images from being included
        "page-size": "A4",
        "dpi": 50,  # Lower resolution for smaller size
    }

    # template 설정
    template_path = os.path.join(base_path, "bin", "template.html")
    html_template = load_template(template_path)
    template = Template(html_template)

    # load pandas
    records = df.to_dict(orient="records")
    responses = [{"blank_pages": 0,
                  "records": record} for record in records]

    # 공백 페이지 수 계산
    if batch_size != 1:
        total_responses = len(responses)
        for idx, response in enumerate(responses):
            html_content = template.render(responses=[response])

            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as temp_pdf:
                temp_pdf_path = temp_pdf.name
                pdfkit.from_string(html_content, temp_pdf_path, configuration=config, options=options)

                num_pages = count_valid_pages(temp_pdf_path)
                response["blank_pages"] = batch_size - (num_pages % batch_size) if num_pages % batch_size != 0 else 0

            if os.path.exists(temp_pdf_path):
                os.remove(temp_pdf_path)

            update_progress(idx / total_responses * 100)

    # 최종 PDF 렌더링
    html_content = template.render(responses=responses)
    pdfkit.from_string(html_content, output_pdf_path, configuration=config, options=options)

    print("PDF 생성 완료!")
    return output_pdf_path
