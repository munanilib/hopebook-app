# 🧪 자동 업데이트 시험판

- `app/` : 실제 기능 코드 (여기만 고치면 됨)
- `dist/app.js` : 빌드 결과 — 모든 사본이 이 파일을 읽어 감
- `dist/Loader.gs` : 사본에 넣는 껍데기 (거의 안 바뀜)
- `tools/build.js` : `node tools/build.js` 로 app → dist 묶기
