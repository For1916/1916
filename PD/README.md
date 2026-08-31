# Handoff: PD Portfolio — 도진우 인턴 포트폴리오 웹사이트

## Overview

교사에서 콘텐츠 PD로 커리어 전환 중인 지원자의 인턴 지원용 포트폴리오 웹사이트. 편집 룸/신문 그리드에서 영감을 받은 **다크 그리드 매거진 레이아웃**으로, 다음 세 가지 축을 축약해서 보여준다.

1. **About** — 이력, 통계, 스킬 칩
2. **Works** — 실제 진행한 4개 프로젝트(Classting SNS, MESA 학원, Instagram 개인 채널, 잘 틀린 문제집)
3. **Film / OS-Bio Cell** — 자작 단편 스크립트("OS-Bio Cell : 제 2의 초전도체")의 예고편·시나리오·스토리보드 서브페이지 3개

톤은 신문 티커 + 편집실의 조합. **파란 액센트(#3D68F6)** 와 모노스페이스 마이크로카피가 다크 배경 위에서 정보 밀도를 만든다.

## About the Design Files

이 번들의 HTML 파일들은 **디자인 레퍼런스**다 — 최종 룩앤필과 인터랙션을 정확히 보여주기 위한 프로토타입이며, 그대로 프로덕션에 붙이려는 코드는 아니다.

작업은 **이 HTML 디자인을 타깃 코드베이스의 기존 환경에서 재구현**하는 것. 프로젝트에 이미 정해진 프레임워크(Next.js / Nuxt / Astro / SvelteKit 등)가 있으면 그 컨벤션과 컴포넌트 시스템을 따라 다시 만들고, 새로 시작한다면 **정적 사이트에 가장 적합한 프레임워크**(Astro 또는 Next.js 정적 export 추천)를 선택해서 구현하면 된다.

## Fidelity

**High-fidelity (hifi)** — 픽셀 단위로 완성된 목업이다. 컬러, 타이포, 스페이싱, 인터랙션(호버, 애니메이션, 스크롤 잠금 오버레이)까지 최종 값으로 잡혀 있다. 그대로 재현하되, 프레임워크의 기존 컴포넌트/유틸을 활용해 코드 구조는 코드베이스 컨벤션에 맞추면 된다.

## Screens / Views

사이트는 **1개 메인 페이지 + 2개 서브 페이지**로 구성된다.

### 1. `Portfolio C - Grid.html` — 메인 랜딩 (허브)

한 페이지 스크롤 구조. 상단 → 하단 순서:

#### 1-1. Ticker Bar (최상단, sticky)
- 높이 32px, 배경 `--bg-2`, 하단 보더 `--line`
- 좌: 브랜드 마크 "◉ LIVE · DESIGNER" + 깜박이는 액센트 도트
- 중: 상태 문구 애니메이션(마퀴) — "AVAILABLE NOW · MATH TEACHER → PD INTERN · APPLICATION TO GENSPARK …"
- 우: 현재 UTC 시각(JS로 매초 업데이트) + "2026.08" 뱃지

#### 1-2. Nav (sticky, ticker 아래)
- 좌측 로고 "PD" (font-display, 800, 20px, 액센트 하이라이트)
- 중앙 인덱스 링크: `01 / ABOUT`, `02 / WORKS`, `03 / FILM`, `04 / CONTACT`
- 우측 CTA 버튼 "HIRE" (액센트 배경, 검정 텍스트, 호버 시 `--ink` 배경)
- 폰트: `--font-mono` 11px, 자간 .1em, 대문자
- 각 링크 좌우로 세로 보더

#### 1-3. Hero (허브 그리드)
매거진 커버 스타일 그리드. 대략 `grid-template-columns: 320px 1fr 260px` / `grid-template-rows: auto auto auto`.
- 좌상단 셀: `◉ ISSUE / 03` + 지원 회사(GENSPARK) / 롤(PD INTERN)
- 중앙 대형 헤드라인 h1: **"교실에서 콘텐츠 <span class='hl'>현장</span>으로."** (font-display, 800, `clamp(56px, 8vw, 128px)`, line-height 1.05, letter-spacing -.04em, `.hl`은 액센트)
- 우측: 인덱스 리스트(각 섹션 번호 · 제목 · 페이지 넘버, `--font-mono`)
- 하단 스트랩: `INSIDE` 라벨 + 3개 강조 문구
- 하단 lede 문단: 짧은 자기소개 리드 카피
- 모든 셀 사이 1px 보더 `--line`

#### 1-4. About 섹션 (id="about")
- 상단 헤더 바: 좌 `◉ 01 / ABOUT` (액센트) + 중앙 큰 타이틀 "지원자, <span class='hl'>도진우</span>." + 우 메타 "M · 26 / 서울"
- 본문 3컬럼:
  - 좌: 대형 숫자 `03` (전공 개수) + 캡션
  - 중: lede 문장 — "수학 교사로 시작해서, <span class='hl'>영상·기획</span>으로 확장했습니다..."
  - 우: 프로필 사진(스튜디오샷, `assets/profile-studio.jpg`, aspect 3:4)
- 스탯 로우 (4열, 각 셀 우측 보더):
  - `03` 전공 · Math / Media / Design
  - `04` 실전 프로젝트 · shipped
  - `04` 진행 년차 · 콘텐츠 제작
  - `3` 신문 기자 · 교지 편집 · 신문 편집장 (언론직 경력)
  - `24H` 회신 · Slack / Email
- 스킬 칩 로우: `Premiere Pro`, `Photoshop`, `Figma`, `Notion`, `Kakao Channel`, `Instagram`, `Google Analytics`, `Genspark AI`, `Cursor` 등 — 각 칩 `--line-2` 보더, mono 11px, 호버 시 액센트 채움

#### 1-5. Works 섹션 (id="works")
- 헤더 바: `◉ 02 / WORKS` + "직접 만든 <span class='hl'>네 개의 프로젝트</span>." + `SHIPPED · 2023–2026`
- 4개 프로젝트 카드가 세로 스택. 각 카드는 좌우 스플릿(대략 40/60):
  - **W01 · Classting SNS 콘텐츠 기획** — 좌: 넘버링 + 카피 + KPI 태그(피드/QnA/채팅) / 우: `assets/classting-*.jpg` 4장 콜라주
  - **W02 · MESA 학원 브랜딩** — `assets/mesa-*.jpg` (커리큘럼 시트 + 매거진), 하단 "3% 안 진입" 태그(`assets/olympiad-3pct.jpg`)
  - **W03 · @jinu._.heart 인스타그램 개인 채널** — `assets/instagram-profile.jpg`, 팔로워/릴스 수 KPI
  - **W04 · 잘 틀린 문제집** — `assets/wrongbook/*` (5권 표지 + 빈우 클라우드 스킴)
- 각 카드 헤더: `W0N` 넘버 + 프로젝트 코드 + 연도 + 롤(Direction / Editing / Copy 등)
- 카드 사이 1px 보더 구분

#### 1-6. Film Section (id="film") — OS-Bio Cell
자작 단편 스크립트. `.work-mod` 12컬럼 그리드 구조:
- **상단 좁은 좌측 aside** (`.wm-info`, grid-column 1/5, sticky) — `◉ WORK 05 / 05`, kind "Screenplay · Short film", h3 "OS-Bio Cell : 제 2의 초전도체", 로그라인 설명, ROLE/TYPE facts 바. **이 사이드에는 3카드를 넣지 않는다.**
- **상단 우측 미디어** (`.wm-media.osbio-slide`, grid-column 5/-1) — "OS-Bio Cell : 제 2의 초전도체", "2026 여름 대공개", "SHORT · MYSTERY / THRILLER" 포스터
- **하단 전체 폭 3카드** (`.film-triptych.film-triptych--full`, grid-column 1/-1) — 포스터 아래로 이동. 각 카드 padding `32px 32px 28px`, 스테이지 height `200px`, 타이틀 `clamp(28px, 2.4vw, 40px)`, 서브 14px:
  - **01 · 예고편** (`ft-card--trailer`) — CTA "Play trailer", 클릭 시 `#trailerPopup` 모달 팝업 (실제 영상 없음, "면접에서 보여드리겠습니다" 안내). 스테이지에 재생 아이콘 + 파형 애니메이션
  - **02 · 시나리오** (`ft-card--script`) — 스테이지에 스크립트 스니펫 렌더(INT. 연구소, 인물 큐, 액션 라인, 커서 깜박임). CTA "Read screenplay" → `scenario.html`
  - **03 · 스토리보드** (`ft-card--board`) — 미니 프레임 3개 프리뷰, 호버 시 부채꼴 효과. CTA "Open storyboard" → `storyboard.html`

#### 1-7. Working Style / Grow Block
- 2컬럼 그리드 `1.2fr auto` / gap 56px / padding `28px 60px` / 배경 `--bg-2`
- 좌: eyebrow `— Working style` (mono, 13px, 액센트) + 대형 인용문 h3 "이미 완성된 **PD**가 아니라, 빠르게 **성장하는 사람.**" (font-display, 600, `clamp(48px, 6vw, 96px)`, line-height 1.08, letter-spacing -.04em) + 서브 "교실에서 콘텐츠 현장으로. 배우면서 만들고, 만들면서 다시 배웁니다." (18~24px)
- 우: 프로필 사진 (`assets/profile-front.jpg`, 폭 `clamp(380px, 32vw, 500px)`, aspect 683/1024, `object-fit: contain`, `background: #ffffff` — 전신이 잘리지 않고 다 보이도록)

#### 1-8. Contact / Finale (id="contact")
- 좌 상단: `◉ 04 / GET IN TOUCH` + h2 "같이 만들 **이야기**가 아직 남아있습니다."
- 대형 메일 셀: `1916@mail4.uk ↗`
- 우측 CTA 카드: `◉ HIRE / AVAILABLE 2026` + "Let's talk." + "Reply within 24h →"
- 하단 4열 로우: Instagram(@jinu._.heart), Profile(도진우.enn.kr), Email, Role(PD Intern)
- 콜로폰: `© 2026 · PD PORTFOLIO` + `Set in Inter Tight + Pretendard + JetBrains Mono`

### 2. `scenario.html` — 시나리오 서브 페이지
- 상단에 동일 스타일 nav (로고 "OS-Bio Cell. Screenplay" + "← Portfolio / 스토리보드 →")
- 본문: 실제 시나리오 텍스트를 스크린플레이 포맷(INT/EXT 슬러그, 액션 라인, 캐릭터 큐, 대사)으로 렌더
- 페이지 색상 시스템은 메인과 동일한 다크 팔레트
- 하단 back-hub 플로팅 버튼 "← Portfolio"

### 3. `storyboard.html` — 스토리보드 서브 페이지
- 상단 nav (로고 "OS-Bio Cell. Storyboard" + "← 시나리오 / Portfolio →")
- 본문: 6개 클립 이미지를 세로 스택으로 풀블리드 렌더
  - `clip1.png` ~ `clip3.png`, `clip5.png`, `clip6.png` (각 ~1.5MB PNG, 1024×683)
  - `clip4.jpg` (2.13MB, 1600×5930 — 세로형 시트, 컬러 스토리보드)
- 각 프레임 컨테이너 `background: #fff`, 하단 보더 `--line`
- 하단 back-hub 플로팅 버튼

## Interactions & Behavior

### Global
- **커스텀 커서**: 화면에 8×8 액센트 도트 커서(`.cursor`). 링크/버튼 호버 시 80×80 아웃라인 커서로 확대 (`hover` 클래스 토글). 데스크탑만.
- **시간 티커**: 우상단 UTC 시각을 매초 갱신하는 JS 인터벌.
- **부드러운 스크롤**: nav 링크 클릭 시 앵커로 스무스 스크롤.
- **폰트 로딩**: Pretendard Variable + Inter Tight + JetBrains Mono, 웹폰트 CDN에서 로드.

### Film 섹션
- `ft-card--trailer` 클릭 → `#trailerPopup` 모달 표시 (`aria-hidden=false`). 카드 자체가 링크가 아니라 팝업 트리거.
  - 모달 내용: "면접에서 **보여드리겠습니다.**" + "30초 예고편은 대면 면접 자리에서 직접 공개합니다." + 확인 버튼
  - 배경 클릭/확인 버튼/Esc 키로 닫힘
- `ft-card--script` → `scenario.html`로 페이지 이동
- `ft-card--board` → `storyboard.html`로 페이지 이동
- 시나리오 카드 내부 커서 깜박임(`.ft-cursor`) 애니메이션 — 순수 CSS `@keyframes`

### Hover 상태
- Nav 링크: 배경 → 액센트, 텍스트 → 검정
- Nav CTA "HIRE": 배경 → `--ink`(오프화이트), 텍스트 검정
- 스킬 칩: 배경 → 액센트, 텍스트 → 검정
- Works 카드: 미세한 배경 톤 시프트 + 우측 화살표 슬라이드
- 인덱스 리스트 항목: 텍스트 → 액센트

### 스크롤 잠금 / 오버레이
- 페이지 상단 안내 오버레이 "빠르게 액세스하려면, 북마크 바에 북마크를 배치하세요"는 브라우저 크롬(genspark.site 서브도메인 프리뷰) 요소 — **실제 배포에서는 제거**할 것.

## State Management

정적 사이트. React/Vue로 재구현 시 필요한 로컬 UI 상태만:
- `trailerModalOpen: boolean` — Film 섹션 예고편 모달 열림/닫힘
- `cursorPos: {x, y}, cursorHover: boolean` — 커스텀 커서 위치와 호버 상태 (데스크탑만)
- `nowUtc: string` — 티커의 현재 시각(setInterval로 매초 갱신)
- `activeSection: string` — nav 하이라이트(옵션, IntersectionObserver로 스크롤 위치 추적)

데이터 페칭 불필요. Works/Film 데이터는 하드코딩된 정적 JSON으로 리팩토링해서 컴포넌트에서 map하면 확장성 확보 가능.

## Design Tokens

### Colors
```css
--bg:      #141414;   /* 페이지 배경(다크 챠콜) */
--bg-2:    #1e1e1e;   /* 티커 / 카드 / grow-block 배경 */
--bg-3:    #282828;   /* 한 단계 더 밝은 서피스 */
--ink:     #f0ede6;   /* 본문 텍스트(오프화이트, 약간 웜) */
--ink-dim: #b8b4a9;   /* 세컨더리 텍스트 */
--ink-sub: #7a766d;   /* 티어셔리(마이크로카피, 캡션) */
--line:    #333333;   /* 기본 보더 */
--line-2:  #3d3d3d;   /* 강조 보더(칩 등) */
--accent:  #3D68F6;   /* 브랜드 블루(하이라이트, CTA, 커서) */
--accent-2:#f2a03d;   /* 세컨더리 웜 액센트(태그, 뱃지) */
```

### Typography
```css
--font-sans:    'Pretendard Variable', Pretendard, sans-serif;   /* 한글 본문 */
--font-display: 'Inter Tight', 'Pretendard Variable', sans-serif; /* 헤드라인 */
--font-mono:    'JetBrains Mono', ui-monospace, monospace;        /* 마이크로카피, 티커, 넘버링 */
```

**주요 타입 스케일:**
- Hero h1: `clamp(56px, 8vw, 128px)` / weight 800 / line-height 1.05 / letter-spacing -.04em
- Section title: `clamp(36px, 4.5vw, 72px)` / weight 700-800 / letter-spacing -.04em
- Grow quote: `clamp(48px, 6vw, 96px)` / weight 600 / line-height 1.08 / letter-spacing -.04em
- 본문 lede: `clamp(20px, 2vw, 32px)` / weight 400
- 서브 카피: `clamp(15px, 1.4vw, 22px)`
- 마이크로카피(mono): 10-13px / letter-spacing .1em~.2em / uppercase
- 링크/nav mono: 11px / letter-spacing .1em / uppercase

### Spacing
- 셀 패딩: 20~28px
- 섹션 헤더 패딩: 32~40px
- 섹션 세로 여백: 60~80px (grow-block은 축소해서 상하 28px)
- 그리드 gap: 0 (셀 사이는 1px 보더로만 구분) / 카드 그리드는 32~56px

### Border / Line
- 기본: `1px solid var(--line)` — 그리드 셀 구분, 섹션 사이
- 강조: `1px solid var(--line-2)` — 칩, 인터랙티브 요소
- Border radius: **거의 사용 안 함**(0). 아이덴티티가 각진 매거진 그리드. 예외는 티커의 라이브 도트(circle).

### Shadow
- **거의 사용 안 함**. 트레일러 모달 등 오버레이에만 은은한 `box-shadow: 0 20px 60px rgba(0,0,0,.5)` 정도.

## Responsive Behavior

- **데스크탑 우선** (1440px+ 최적). 대부분의 그리드는 명시적 컬럼 수로 잡혀 있음.
- **≤ 960px 브레이크포인트**:
  - `grow-block`: 2컬럼 → 1컬럼, 사진 상단으로, 폰트 스케일 축소
  - Hero 그리드: 컬럼 재배치, h1 축소
  - Works 카드: 좌우 스플릿 → 세로 스택
  - Nav: 링크 축소 또는 햄버거 메뉴 필요(현재 프로토타입에는 미구현 — **개발 시 추가 필요**)
- **모바일 (≤ 480px)**: 커스텀 커서 비활성화, 티커 마퀴 속도 조정, 폰트 스케일 최소값으로 하강
- 커스텀 커서는 `(hover: hover) and (pointer: fine)` 미디어 쿼리로 데스크탑만 활성화 권장

## Assets

모든 이미지는 `assets/` 폴더 아래. 실제 프로젝트에서 캡처한 실사진과 스크린샷.

### Profile
- `profile-front.jpg` — 정면 전신샷 (683×1024, ~960KB, 흰 배경) — Working style 섹션
- `profile-studio.jpg` — 스튜디오 반신샷 (~1MB) — About 섹션

### Works — Classting
- `classting-feed.jpg`, `classting-chat.jpg`, `classting-qna.jpg`, `classting-quote.jpg` — 클래스팅 앱 실제 콘텐츠 스크린샷

### Works — MESA 학원
- `mesa-curriculum.jpg` — 커리큘럼 시트
- `mesa-magazine.jpg` — 학원 매거진 커버
- `olympiad-3pct.jpg` — "상위 3% 진입" 성과 캡처

### Works — Instagram
- `instagram-profile.jpg` — @jinu._.heart 채널 프로필 캡처

### Works — 잘 틀린 문제집 (wrongbook/)
- `workbook-01.jpg` ~ `workbook-05.jpg` — 문제집 5권 표지
- `binhwoo-clouds.jpg` — 스킴 이미지
- `yuje-1.jpg` — 유제 페이지

### Film — OS-Bio Cell
- `storyboard/clip1.png` ~ `clip3.png`, `clip5.png`, `clip6.png` (1024×683 PNG, ~1.5MB 각)
- `storyboard/clip4.jpg` (1600×5930 컬러 세로 시트, 2.13MB — 최근 고화질로 재업로드된 파일)
- `film-noir.jpg` — 필름 무드 참조

### 기타
- `hero-workspace.jpg` — 워크스페이스 무드
- `landing-page.jpg` — 랜딩 참고
- `kakao-channel.jpg` — 카카오 채널 캡처
- `book-beige.jpg` — 도서 무드
- `replay/replay-*.jpg` — 리플레이 시퀀스 (7장)
- `congrats_image.png` — 이스터에그/성과 이미지

### 웹폰트 (외부 CDN, 로컬 저장 안 함)
```html
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/variable/pretendardvariable-dynamic-subset.min.css">
<link href="https://fonts.googleapis.com/css2?family=Inter+Tight:wght@400;500;600;700;800;900&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
```

## Files

번들에 포함된 디자인 소스:

- **`Portfolio C - Grid.html`** — 메인 랜딩 페이지 (인라인 CSS + JS, 자기완결)
- **`scenario.html`** — 시나리오 서브 페이지 (스크립트 포맷 렌더)
- **`storyboard.html`** — 스토리보드 서브 페이지 (6개 클립 세로 스택)
- **`assets/`** — 모든 이미지 자산 (37개, 총 ~18MB)

각 HTML은 인라인 `<style>` + 인라인 `<script>`로 자기완결되어 있어, 브라우저에서 바로 열면 그대로 작동한다. 재구현 시 CSS 변수를 앱의 토큰 시스템(Tailwind config / CSS custom properties / Styled-components theme)으로 옮기고, 각 섹션을 컴포넌트로 분리하는 것을 권장한다.

## Implementation Notes

- **매거진 그리드가 핵심 아이덴티티**다. `border-radius: 0`, 셀 사이 1px 보더, 모노스페이스 마이크로카피 — 이 세 요소를 유지해야 룩이 산다.
- **한글 + 영문 혼용 타이포**: Pretendard로 한글 안정성 확보하면서, 큰 헤드라인은 Inter Tight의 조여진 letter-spacing으로 밀도를 만든다. 두 폰트의 x-height가 어느 정도 맞아서 함께 써도 자연스럽다.
- **컬러**: 파란 액센트(#3D68F6)는 절제해서 하이라이트/CTA/커서에만 사용. 남용하면 매거진 톤이 무너진다.
- **커스텀 커서**는 브랜드 시그니처지만, 접근성/모바일에서 반드시 폴백해야 한다.
- **콘텐츠 카피(한글)** 은 이력서 원본이므로 함부로 바꾸지 말 것. 개발자가 재구현할 때 텍스트 상수만 뽑아 i18n 파일로 옮기는 정도가 안전하다.
