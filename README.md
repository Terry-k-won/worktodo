# 🚀 Life & Task Management Dashboard

나의 삶과 목표, 당장의 과제를 한눈에 정리하고 **구글 스프레드시트(DB)** 및 **모바일 웹**에서 손쉽게 관리할 수 있는 모던 대시보드 웹 애플리케이션입니다.

---

## ✨ 핵심 기능

1. **5가지 카테고리 일괄 관리**
   - ☀️ **오늘 할 일** (Today's Tasks)
   - 🔥 **당장 내게 쌓여있는 과제** (Urgent Assignments)
   - 🚩 **앞으로 장기적으로 해야할 일** (Long-term Goals)
   - 🛠️ **내가 개발해야할 능력들** (Skills to Develop)
   - 🛡️ **나의 부족함 (개선점)** (Areas for Improvement)

2. **완료 및 휴지통 보관 / 복원 기능 (체크 처리)**
   - 메인 화면에서 체크박스를 누르면 **부드러운 애니메이션**과 함께 메인 목록에서 사라집니다.
   - 상단 📥 **완료 및 휴지통 목록** 창에서 언제든지 완료되거나 삭제된 항목을 확인 가능합니다.
   - **`복원`** 버튼을 누르면 실수로 완료/삭제한 항목을 원래 카테고리로 되돌릴 수 있습니다.

3. **구글 스프레드시트(Google Sheets) DB 연동**
   - 구글 앱스 스크립트(Google Apps Script) 웹 앱 URL을 등록하여 기기(PC/스마트폰) 간 실시간 데이터 동기화.
   - URL 미설정 시에도 브라우저 **LocalStorage에 안전하게 자동 저장**됩니다.

4. **모바일 웹 및 핑거 터치 최적화**
   - 스마트폰 화면에서도 손쉽게 탭하고 과제를 체크할 수 있는 반응형 뷰와 모바일 FAB 버튼 제공.

5. **GitHub Actions 자동 배포**
   - `main` 브랜치에 코드를 push하면 **GitHub Pages**로 웹사이트가 자동 배포됩니다.

---

## 🚨 "This app is blocked" (보안 차단) 문제 해결 방법

구글 앱스 스크립트 권한 승인 과정에서 **"This app is blocked"** (이 앱은 차단되었습니다) 또는 **"Google에서 이 앱을 검증하지 않았습니다"** 경고가 떠서 차단되는 경우, 아래 **3가지 단계**로 100% 해결할 수 있습니다.

### 1단계: 스프레드시트 내부에서 스크립트 생성 (가장 중요 ⭐)
> ⚠️ `script.google.com` 사이트에서 독립형 스크립트로 만들면 구글 드라이브 전체 접근 권한을 요구하여 구글 보안 시스템에 의해 차단됩니다!
- 구글 드라이브에서 만든 **구글 스프레드시트 파일 내부**로 들어갑니다.
- 상단 메뉴의 **[확장 프로그램] ➔ [Apps Script]**를 클릭하여 스크립트 창을 엽니다.

### 2단계: `@OnlyCurrentDoc` 보안 주석 확인
- 코드 최상단에 `/** @OnlyCurrentDoc */` 주석이 있는지 확인합니다.
- 이 구문은 해당 스크립트가 "현재 스프레드시트 1개"에만 접근하도록 권한 범위를 최소화하여 구글 보안 블락을 방지합니다.

### 3단계: 권한 승인 화면 패스 (고급 -> 이동)
1. **[배포] ➔ [새 배포]** (웹 앱 선택 / 실행: 나 / 액세스: 모든 사용자) 후 **[권한 검토]** 클릭
2. 본인의 구글 계정을 선택합니다.
3. *"Google에서 이 앱을 검증하지 않았습니다"* 또는 *"Google haven't verified this app"* 화면이 나오면:
   - 좌측 하단의 **`고급` (Advanced)** 글자를 클릭합니다.
   - 아래에 나타나는 **`TaskDB(안전하지 않음)으로 이동` (Go to TaskDB (unsafe))`** 링크를 클릭합니다.
   - **`허용` (Allow)** 버튼을 클릭합니다.
4. 이제 발급된 **`웹 앱 URL`**을 복사하여 웹 프로그램의 ⚙️ **구글 연동 및 설정**에 등록하시면 완벽히 작동합니다!

---

## 📊 구글 스프레드시트 연동 방법 (Google Apps Script)

1. [Google Sheets](https://sheets.google.com)에서 **새 스프레드시트**를 하나 만듭니다.
2. 메뉴의 **[확장 프로그램] ➔ [Apps Script]**를 클릭합니다.
3. 기존 코드를 모두 삭제하고, 프로젝트의 [`google_apps_script.gs`](google_apps_script.gs) 내용 전체를 붙여넣습니다.
4. **[배포] ➔ [새 배포]** (유형: 웹 앱, 실행: 나, 액세스: 모든 사용자)를 실행합니다.
5. 위 3단계 순서대로 권한 승인을 완료한 후 생성된 **Web App URL**을 웹 프로그램 설정에 등록합니다.

---

## 🌐 GitHub Actions로 배포하는 방법

1. 이 프로젝트 코드를 본인의 GitHub 저장소(Repository)로 push합니다.
   ```bash
   git init
   git add .
   git commit -m "Initial commit of Life Dashboard"
   git remote add origin https://github.com/사용자명/저장소명.git
   git push -u origin main
   ```
2. GitHub 저장소 페이지의 **Settings ➔ Pages**로 이동합니다.
3. **Build and deployment ➔ Source** 항목을 **`GitHub Actions`**로 변경합니다.
4. 코드가 push될 때마다 [.github/workflows/deploy.yml](.github/workflows/deploy.yml)이 자동으로 실행되어 나만의 웹사이트 주소(`https://사용자명.github.io/저장소명`)로 무료 배포됩니다.
