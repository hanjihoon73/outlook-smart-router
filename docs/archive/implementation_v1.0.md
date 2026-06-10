# Outlook Smart Router (가칭) 크롬 확장앱 기획안 v1.0

Outlook 웹에서 다중 규칙 조합 시 'AND(그리고)' 조건만 가능한 한계를 극복하고, 더욱 유연하고 강력한 메일 자동 분류를 지원하는 크롬 확장 프로그램입니다.

## 1. 핵심 문제 및 해결 방안

*   **문제점:** Outlook 기본 규칙 설정은 서로 다른 조건 간에 'AND'만 지원하며, 'OR(또는)' 조건이나 복잡한 괄호 `( )` 묶음을 지원하지 않아 여러 규칙을 만들어야 하는 번거로움이 있습니다.
*   **해결 방안 (기술적 접근):**
    *   확장앱 내에서 사용자가 자유롭게 AND/OR 및 그룹화 조건(예: `(A AND B) OR C`)을 설정할 수 있는 **고급 규칙 빌더 UI**를 제공합니다.
    *   사용자가 설정한 복잡한 규칙을 논리식의 **DNF(논리합 정규형, OR of ANDs)** 형태로 변환합니다.
    *   **Microsoft Graph API**를 사용하여 변환된 각각의 AND 그룹을 여러 개의 개별 Outlook 순정 규칙으로 쪼개어 서버에 저장합니다.
    *   *효과:* 브라우저가 꺼져 있거나 모바일 앱을 사용할 때도 Outlook 서버 단에서 규칙이 정상적으로 작동하게 됩니다.

## 2. 주요 기능 (Features)

1.  **고급 규칙 빌더 (팝업 UI)**
    *   크롬 우측 상단의 확장앱 아이콘을 클릭하여 빠르고 간편하게 실행할 수 있는 팝업창 UI 제공.
    *   조건 간 AND / OR 선택 및 그룹화(괄호) 기능.
2.  **핵심 액션: 특정 폴더로 이동**
    *   조건이 일치할 때 수행할 작업으로, 가장 많이 사용되는 '지정된 폴더로 이동' 기능을 최우선으로 지원합니다.
3.  **Microsoft 계정 안전 연동**
    *   최초 1회 Microsoft 계정 로그인을 통해 Graph API 통신 권한을 획득합니다 (개인정보는 서버에 저장되지 않고 브라우저에 안전하게 보관됨).
4.  **통합 규칙 관리자**
    *   확장앱이 만든 복잡한 규칙은 뒷단에서 여러 개로 쪼개지지만, 사용자에게는 하나의 규칙으로 묶어 보여줍니다 (`[SmartRouter]` 접두사 사용).

## 3. 폴더 구조 및 파일 구성 (Folder Structure)

프로젝트는 React와 Vite를 기반으로 구축되며, 지침에 따라 각 역할별로 폴더를 명확히 분리합니다.

```text
outlook-smart-router/
├── docs/                     # 기획 및 프로젝트 관련 문서 보관
│   └── archive/              # 이전 버전의 문서 보관
├── public/                   # 확장앱 아이콘 및 정적 파일
├── src/                      # 소스 코드
│   ├── assets/               # 이미지, 글로벌 CSS 등
│   ├── components/           # UI 컴포넌트 (규칙 빌더, 버튼, 등)
│   ├── hooks/                # 커스텀 훅 (상태 관리, Graph API 훅 등)
│   ├── utils/                # 유틸리티 함수 (DNF 변환 알고리즘 등)
│   ├── api/                  # API 호출 함수 (MS Graph API 통신 전담)
│   ├── background/           # 확장앱 백그라운드 서비스 워커 (토큰 관리)
│   ├── popup/                # 팝업창 진입점 UI (index.html, main.tsx)
│   └── types/                # TypeScript 타입 정의
├── manifest.json             # 크롬 확장앱 설정 파일 (Manifest V3)
├── package.json              # 의존성 패키지 관리
├── vite.config.ts            # 빌드 도구 설정
└── tsconfig.json             # TypeScript 설정
```

## 4. 기술 스택 (Tech Stack)

*   **Frontend UI:** React.js + Tailwind CSS (빠른 UI 구성 및 상태 관리)
*   **Build Tool:** Vite (빠른 개발 및 빌드 환경)
*   **Chrome Extension API:** `Manifest V3` 기반
*   **연동 API:** Microsoft Graph API (`MailboxSettings.ReadWrite`, `User.Read` 권한)

## 5. 개발 로드맵 (Phases)

*   **Phase 1: 기반 세팅 및 인증 (PoC)**
    *   Vite 기반 React 프로젝트 및 Manifest V3 초기 설정
    *   MSAL.js를 활용한 Microsoft Graph API OAuth2 로그인 연동 구현
*   **Phase 2: UI 개발 및 논리 변환기 (Compiler) 제작**
    *   팝업창 UI (조건 블록 쌓기 형태) 개발
    *   `사용자 입력 -> DNF 논리 변환 로직` 알고리즘 개발
*   **Phase 3: 폴더 이동 액션 및 규칙 생성 연동**
    *   Graph API를 활용한 '사용자 폴더 목록' 조회 기능
    *   작성된 복잡한 규칙을 여러 개의 Graph API 페이로드로 생성 후 서버에 저장
*   **Phase 4: 통합 테스트**
    *   규칙 동기화 및 엣지 케이스 테스트
