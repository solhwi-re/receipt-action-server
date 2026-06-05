# Receipt Action Server

MyGPT Action에서 호출할 DOCX 생성 서버입니다.

## 파일 구성
- `app.py` : FastAPI 서버
- `template.docx` : 중앙청년지원센터 법인카드 영수증 증빙자료 원본 양식
- `requirements.txt` : Render 설치 패키지
- `openapi.yaml` : MyGPT Actions에 붙여넣을 스키마

## Render 실행 명령
```bash
uvicorn app:app --host 0.0.0.0 --port $PORT
```

## API
POST `/create-receipt-docx`

영수증이 여러 장이면 하나의 DOCX 안에 페이지를 나누어 생성합니다.
