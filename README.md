# Google Forms to PDF

## 개요
이 프로그램은 구글 폼 설문지의 응답 결과 CSV를 인쇄하기 편하도록 PDF로 변환해주는 Python 기반 GUI 프로그램입니다.

## 주요 기능
- HTML 템플릿을 활용하여 PDF 양식 자유롭게 변경 가능
- 응답자별 설문결과 페이지가 홀수일경우 자동으로 빈 페이지를 삽입하여 양면인쇄를 도와줌

## 빌드 방법
PyInstaller를 사용하여 프로그램을 빌드할 수 있습니다.

아래 명령어를 사용하세요:
```bash
pyinstaller -F -w gui.py --additional-hooks-dir=.
```
