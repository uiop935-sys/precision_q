# PRECISION-Q 배포 가이드

## 1단계: GitHub에 올리기

```bash
cd precision-q
git init
git add .
git commit -m "PRECISION-Q v1.0"
```

GitHub에서 새 리포지토리 만들고:
```bash
git remote add origin https://github.com/너의아이디/precision-q.git
git branch -M main
git push -u origin main
```

## 2단계: Vercel 배포 (1분 소요)

1. **https://vercel.com** 접속 → GitHub로 로그인
2. **"Import Project"** 클릭 → `precision-q` 리포 선택
3. **Framework**: Vite 자동 감지됨
4. **"Deploy"** 클릭 → 끝!

배포 완료되면 `https://precision-q.vercel.app` 같은 URL이 나옴.

## 3단계: 휴대폰에서 앱처럼 설치

### iPhone (Safari)
1. 배포된 URL 접속
2. 하단 **공유 버튼** (□↑) 탭
3. **"홈 화면에 추가"** 선택
4. 앱 아이콘이 홈화면에 생김 → 풀스크린으로 실행됨

### Android (Chrome)
1. 배포된 URL 접속
2. 상단 **"앱 설치"** 배너 탭 (또는 메뉴 → "앱 설치")
3. 홈화면에 앱 아이콘 생성

## 참고사항

- **OCR 기능**: Claude API를 사용합니다. Vercel 환경에서 API 키 없이도 동작하지만, 실제 OCR 호출 시 Claude API 키가 필요합니다.
- **데모 모드**: API 키 없어도 "데모 데이터로 테스트" 버튼으로 전체 플로우 확인 가능합니다.
- **분석 이력**: localStorage에 저장되므로 기기별로 유지됩니다.
- **PWA**: 오프라인에서도 캐시된 페이지는 접근 가능합니다.
