# DARToSheeToDART

DART `.dsd` 파일을 구글 시트로 변환하고, 수정 후 다시 `.dsd`로 재생성하는 양방향 Google Apps Script 도구입니다.

## 주요 기능

- **DSD -> 시트 변환**: `contents.xml`을 파싱해 표지/목차/감사보고서/재무제표/주석을 시트 탭으로 분리 생성합니다.
- **시트 -> DSD 역변환**: 작업 시트 변경사항을 원본 XML 구조에 패치해 DART 업로드 가능한 `.dsd`를 만듭니다.
- **메타데이터 보존**: 숨김 시트 `__METADATA__`에 원본 블록/섹션/런타임 문맥을 보관해 roundtrip 안정성을 높입니다.
- **워크스페이스 자동화**: `source/worksheet/backupsheet/target` 폴더를 자동 관리하고 baseline 백업 시트를 생성합니다.
- **북마크 기반 주석 분리(개선)**: 주석 분리 시 북마크를 우선 사용하고, 부재 시 fallback 로직을 적용합니다.
- **서술문 숫자 추출(개선)**: 문장 내 숫자를 수식 + `[IGNORE]` 값으로 분리하며, 다중 숫자 참조 오프셋을 정확히 맞춥니다.

## 내부 구현 요약

- `src/Main.js`: 웹앱 API 엔트리포인트, 스프레드시트 메뉴, 워크스페이스 초기화/변환 실행
- `src/XmlParser.js`: `contents.xml` 파싱, 섹션/블록/주석 경계 식별, 제목/번호 정규화
- `src/SheetBuilder.js`: 작업 시트 생성, 표/서술문 렌더링, UX 포맷팅, 작업안내 시트 작성
- `src/SheetToDsd.js`: 시트 변경사항 기반 XML 패치, `.dsd` 재패키징
- `src/ZipManager.js`: `.dsd` 압축 해제/재생성 유틸리티
- `src/WebApp.html`: 사용자 웹앱 UI

## 빠른 시작

### 1) 테스트 웹앱으로 바로 사용 (권장)

- 테스트 URL: `https://script.google.com/macros/s/AKfycby6aHKq7d0_iOkhlF31FjBa5jy4sQI4Ez8Nm6LU2E_DKyRl74PstiiFkX7QtgZHGcV9wA/exec`
- 사용 순서:
  1. 웹앱에서 워크스페이스 루트 폴더 URL/ID 입력
  2. DSD 업로드 -> `worksheet/backupsheet/target` 자동 관리
  3. 시트 수정 후 `Sheet -> DSD` 실행
  4. `target` 폴더 결과 확인

### 2) `clasp`로 직접 배포해서 사용

```bash
npm install -g @google/clasp
clasp login
cp .clasp.json.sample .clasp.json
# .clasp.json의 scriptId를 내 Apps Script 프로젝트 ID로 변경
clasp push
clasp version "release note"
clasp deploy
```

- `.clasp.json` 원본 파일은 로컬 전용이며 git 추적 대상에서 제외됩니다.
- 저장소에는 `.clasp.json.sample`만 포함되어 있으며, 샘플의 `scriptId`를 교체해서 사용하면 됩니다.

#### 히스토리 초기화 운영 수칙

- GitHub 히스토리를 초기화해야 할 때는 먼저 로컬 백업 브랜치를 생성합니다.
  - 예: `git branch local-history-backup-YYYYMMDD`
- 백업 브랜치는 **로컬 보관 전용**으로 유지하고 원격에는 푸시하지 않습니다.
- 원격 초기화 이후 다른 작업 환경에서는 `git fetch --all --prune` 후 `git reset --hard origin/main`으로 재동기화합니다.

## 사용 방법

### 1) 웹앱에서 사용하는 방법 (권장)

1. 웹앱 URL을 열고 워크스페이스 루트 폴더 URL/ID를 입력합니다.
2. DSD 파일을 업로드하면 `worksheet`에 작업 시트, `backupsheet`에 baseline 백업이 생성됩니다.
3. 작업 시트에서 편집 후, 웹앱의 `Sheet -> DSD` 변환을 실행합니다.
4. 결과 `.dsd`는 `target` 폴더에서 확인합니다.

### 2) 코드를 복사해서 스크립트 프로젝트로 사용하는 방법

1. 구글 앱스 스크립트 프로젝트를 생성합니다.
2. `src/*.js`, `src/*.html`, `src/appsscript.json` 내용을 프로젝트에 반영합니다.
3. 웹앱 배포 후 `onOpen` 메뉴 또는 웹앱 화면에서 변환을 수행합니다.

### 3) 라이브러리 형태로 사용하는 방법 (고급)

- 라이브러리 방식은 호출부 스크립트를 별도로 작성해야 하므로 난이도가 있습니다.
- 필요 시 아래 공개 API를 호출하도록 구현합니다.
  - `apiInitializeWorkspace(payload)`
  - `apiUploadDsdAndCreateWorksheet(payload)`
  - `apiConvertSheetToDsd(spreadsheetInput)`
  - `runSheetsToDsd(spreadsheetInput)`

## 데이터 기록/보안 안내

- 본 프로젝트는 별도 백엔드 DB/분석 서버를 사용하지 않으며, 사용자 입력값을 외부 저장소에 수집하지 않습니다.
- 변환 과정에서 필요한 정보는 사용자 Google Drive 내부 파일(`worksheet`, `backupsheet`, `target`)과 숨김 시트(`__METADATA__`)에만 저장됩니다.

## 로고 사용 주의

- 금감원(FSS) 로고는 데모/테스트 목적의 재미 요소로 포함된 항목입니다.
- 실제 서비스/외부 배포 시에는 법적/브랜딩 이슈를 피하기 위해 반드시 자체 로고/아이콘으로 교체해서 사용하세요.

## 추후 업데이트 예정

- 비감사보고서/다양한 DSD 포맷 대응을 위한 분류 로직 일반화
- 북마크/주석 경계 회귀 테스트 fixture 확장
- 작업 시트 품질 점검(숫자 참조, span/빈셀 복원, note numbering) 자동 검증 스크립트 강화
- 운영 가이드 문서화(배포 ID 단일화 정책, 히스토리 기반 작업 절차)

## 참고 문서

- 구조/설계 메모: [CONTENTS.md](CONTENTS.md)
- 작업 규칙: `docs/WORK_INSTRUCTIONS.md`
