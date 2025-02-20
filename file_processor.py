import os
import pandas as pd
from jinja2 import Template
import pdfkit
import PyPDF2
import shutil
import tempfile

# Jinja2 템플릿 정의
html_template = """
    <!DOCTYPE html>
    <html>
    <head>
        <style>
            body { font-family: "맑은 고딕", "Malgun Gothic", "Apple SD Gothic Neo", sans-serif; }
            h1 { color: #333; }
            .question { font-weight: bold; margin-top: 20px; }
            .answer { margin-bottom: 20px; }
            .section { margin-bottom: 40px; page-break-before: always; } /* 페이지 나누기 */
        </style>
    </head>
    <body>
        <div class="section">
            <h1>설문 조사 결과</h1>
            {% for question, answer in respondent.items() %}
                <h3 class="question">{{ question }}</h3>
                <h4 class="answer">{{ answer }}</h4>
            {% endfor %}
        </div>
    </body>
    </html>
    """


def process_file(file_path, update_progress):
    print(f"파일 처리 시작: {file_path}")

    # CSV 파일 불러오기
    df = pd.read_csv(file_path)

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

                # 응답자 PDF가 홀수 페이지일 경우 빈 페이지 추가
                if num_pages % 2 != 0:
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
