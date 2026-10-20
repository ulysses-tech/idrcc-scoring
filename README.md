# IDCC 2026 · DCC Scoring System

심판 채점 및 HQ 집계 웹앱 (React + Vite + Tailwind + Firebase Realtime Database)

## 배포
- **GitHub Pages**: `main` 브랜치에 커밋하면 `.github/workflows/deploy.yml`이 자동으로 빌드·배포합니다.
  최초 1회 **Settings → Pages → Source = GitHub Actions** 로 설정해야 합니다.
- **Netlify(선택)**: 이 저장소를 Netlify에 연결하면 `netlify.toml` 설정(빌드 `npm run build`, 게시 폴더 `dist`)으로 자동 배포됩니다.

## 수정할 파일
- 앱 코드: `src/App.jsx` (비밀번호·카테고리·점수 규칙은 파일 상단 상수)
- 아이콘: `public/` 폴더

## PC에서 직접 실행(개발자용)
```
npm install
npm run dev
```
