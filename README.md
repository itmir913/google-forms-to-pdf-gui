# Google Forms to PDF

![google_forms_to_pdf](https://github.com/user-attachments/assets/d1540151-6c3c-4fee-ace2-3d1df7ce87dc)

## 개요
이 프로그램은 Google Forms 설문지 응답 결과 CSV 파일을 PDF로 변환하는 Python 기반 GUI 프로그램입니다. 설문 결과를 인쇄하기 쉽게 정리하고, 양면 인쇄를 고려하여 빈 페이지를 자동으로 삽입하는 기능을 제공합니다.

## 주요 기능
- 직관적인 Drag & Drop 인터페이스 제공  
- CSV 설문 결과를 PDF로 변환  
- HTML 템플릿을 활용하여 PDF 양식을 자유롭게 변경 가능  
- 양면 인쇄를 지원하기 위해 자동으로 빈 페이지 삽입

## 사용 방법
1. Google Forms에서 응답 결과 CSV 파일 다운로드
2. `gui.exe` 프로그램 실행
3. [페이지 옵션](https://github.com/itmir913/google-forms-to-pdf-gui?tab=readme-ov-file#%ED%8E%98%EC%9D%B4%EC%A7%80-%EC%98%B5%EC%85%98) 선택
4. 다운로드 받은 CSV 파일을 본 프로그램에 Drag & Drop

## 페이지 옵션
양면 인쇄를 지원하기 위한 다양한 페이지 옵션이 구현되어 있습니다.

- 빈 페이지 추가하지 않음 (단면 인쇄용)  
  : PDF에 빈 페이지를 추가하지 않아 변환 속도가 빠르며, 단면 인쇄에 적합합니다.

- 2의 배수로 빈 페이지 추가 (양면 인쇄용)  
   : 응답자의 설문 결과 페이지 수가 홀수일 경우, 자동으로 빈 페이지를 추가합니다. 이를 통해 양면 인쇄 시, 새로운 응답자가 항상 새로운 용지에서 시작됩니다.

- 4의 배수로 빈 페이지 추가 (2쪽 모아찍기 & 양면 인쇄용)  
   : 응답자의 설문 결과 페이지 수를 4의 배수로 맞추기 위해 자동으로 빈 페이지를 추가합니다. 이를 통해 2쪽 모아찍기 및 양면 인쇄 시, 새로운 응답자가 항상 새로운 용지에서 시작됩니다.

## 빌드 방법
PyInstaller를 사용하여 프로그램을 빌드할 수 있습니다. 아래 명령어를 사용하세요:

```bash
pyinstaller -F -w gui.py --additional-hooks-dir=.
```
