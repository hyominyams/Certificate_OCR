# 📄 Google Sheets Certificate OCR & Classification System

이 프로젝트는 **Google Forms**로 제출된 이수증 이미지를 **Upstage OCR API**를 통해 텍스트로 변환하고, 내용을 분석하여 **Google Sheets**의 적절한 카테고리 시트로 자동 분류해주는 자동화 시스템입니다.

## 🛠 작동 원리 (Architecture)

```mermaid
graph TD
    A[사용자 (교사)] -->|이수증 이미지 업로드| B(Google Forms)
    B -->|응답 저장| C{Google Sheets (Form Responses)}
    C -->|Trigger (onFormSubmit)| D[Google Apps Script]
    D -->|1. 이미지 전송| E[Upstage OCR API]
    E -->|2. 텍스트 추출| D
    D -->|3. 파싱 (Regex / LLM)| D
    D -->|4. 카테고리 매칭 (Config)| F[카테고리별 시트]
```

1.  **제출**: 사용자가 구글 폼을 통해 이수증 사진을 첨부하여 제출합니다.
2.  **트리거**: 폼이 제출되면 `onFormSubmit` 트리거가 작동하여 Apps Script를 깨웁니다.
3.  **OCR (광학 문자 인식)**: 스크립트가 이미지를 **Upstage Document Digitization API**로 보내 텍스트를 추출합니다.
4.  **파싱 (Parsing)**:
    *   **1차**: 정규표현식(Regex)을 사용하여 이수번호, 성명, 생년월일, 과정명 등을 추출합니다. (공백/줄바꿈에 강건하게 설계됨)
    *   **2차 (Fallback)**: 정규식 파싱이 불완전할 경우, **Upstage Solar LLM**을 호출하여 자연어 처리를 통해 정보를 보완합니다.
5.  **분류 (Classification)**: 'Config' 시트에 정의된 키워드 규칙에 따라 적절한 시트(예: 성희롱예방, 법정의무연수 등)로 데이터를 이동시킵니다.

---

## 🚀 시작하기 (Getting Started)

### 1. 사전 준비
*   Google 계정 (Forms, Sheets 사용)
*   **Upstage API Key**: [Upstage Console](https://console.upstage.ai/)에서 발급 필요.

### 2. 설치 방법
1.  **Google Form 생성**: 이수증을 업로드할 수 있는 질문(파일 업로드)을 포함하여 폼을 만듭니다.
2.  **Google Sheet 연결**: 폼 응답이 저장될 스프레드시트를 생성합니다.
3.  **스크립트 편집기 열기**: 스프레드시트 메뉴에서 `확장 프로그램` > `Apps Script`를 클릭합니다.
4.  **코드 복사**:
    *   `setup.gs`: 초기 설정 및 메뉴 관련 코드
    *   `code.gs`: 핵심 로직 (OCR, 파싱, 분류)
    *   위 두 파일의 내용을 복사하여 붙여넣습니다.

### 3. 초기 설정
1.  스프레드시트를 새로고침하면 상단에 **`Certificate OCR`** 메뉴가 나타납니다.
2.  **`1. 초기 설정 (시트 생성)`** 클릭: `Config` 시트와 분류별 시트들이 자동으로 생성됩니다.
3.  **`Upstage API Key 설정`** 클릭: 발급받은 API Key를 입력합니다.
4.  **`2. 트리거 설정 (자동 실행 켜기)`** 클릭: 폼 제출 시 스크립트가 자동 실행되도록 설정합니다.

---

## 📖 사용 방법 (Usage)

### 자동 처리
*   설정이 완료되면, 사용자가 폼을 제출하는 즉시 시스템이 백그라운드에서 작동합니다.
*   잠시 후 해당 카테고리 시트에 분석된 내용이 한 행으로 추가됩니다.

### 수동 처리 (에러 대응)
*   만약 자동 처리가 실패했거나, 트리거 설정 전의 데이터를 처리해야 한다면:
    *   메뉴 > **`미처리 응답 일괄 처리`**를 클릭하세요.
*   **상태(Status) 확인**:
    *   `OK`: 정상 처리됨.
    *   `NEEDS_REVIEW`: 파싱이 불완전함. 사람이 확인 필요.
    *   `OCR_FAIL`: OCR API 호출 실패 또는 이미지 읽기 실패.

### 설정 변경 (Config)
*   **`Config` 시트**에서 분류 규칙을 자유롭게 수정할 수 있습니다.
*   **A열 (시트이름)**: 데이터가 저장될 시트의 이름입니다.
*   **B열 (키워드)**: 해당 시트로 분류되기 위한 키워드들입니다. (쉼표 `,`로 구분)
*   *새로운 연수 종류가 생기면 Config 시트에 행을 추가하기만 하면 됩니다.*

---

## 💻 기술 스택 (Tech Stack)

*   **Platform**: Google Apps Script (Serverless JavaScript environment)
*   **Database**: Google Sheets
*   **AI & OCR**:
    *   **Upstage Document Parse**: 고성능 문서 OCR 모델 (`document-parse`)
    *   **Upstage Solar Pro**: 복잡한 비정형 텍스트 파싱을 위한 LLM (`solar-pro-2`)

### 주요 함수 설명
*   `onFormSubmit(e)`: 트리거 진입점.
*   `processSingleResponse_(row)`: 하나의 응답을 처리하는 메인 파이프라인.
*   `callUpstageOcr_(fileId)`: Upstage API와 통신.
*   `parseCertificateText_(text)`: 정규식 기반 텍스트 추출.
*   `classifyCategory_(...)`: 키워드 매칭 알고리즘.

---

## ⚠️ 주의사항
*   **API 비용**: Upstage API는 사용량에 따라 과금될 수 있습니다.
*   **이미지 권한**: 스크립트가 이미지를 읽으려면 해당 파일이 저장된 Google Drive 폴더에 접근 권한이 있어야 합니다. (폼으로 제출된 파일은 기본적으로 소유자에게 권한이 있으므로 문제없음)
