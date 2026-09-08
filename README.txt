# 화천기공 Streamlit 배포 패키지 안내

1. 압축을 해제하면 다음 파일들이 포함되어 있습니다:
   - `streamlit_app.py` (메인 웹 앱 코드)
   - `requirements.txt` (필수 패키지 목록)
   - `README.txt` (본 안내 파일)

2. 폰트 파일 안내:
   - PDF 생성 시 한글 깨짐을 방지하기 위해 윈도우의 '맑은 고딕' 폰트 파일(`malgun.ttf`)을 이 폴더와 동일한 위치에 복사해 넣어주세요.
   - (경로: C:\Windows\Fonts\malgun.ttf 복사 후 본 폴더에 붙여넣기)

3. 스트림릿 클라우드 배포 방법:
   - GitHub에 본 파일들을 업로드합니다.
   - Streamlit Community Cloud (share.streamlit.io)에 로그인 후 New App을 생성하고 저장소와 `streamlit_app.py` 경로를 지정합니다.
