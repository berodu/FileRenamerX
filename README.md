# FileRenamerX

배관 검사 비디오 처리를 위한 파일명 자동 변경 유틸리티 프로그램입니다.

## API 키 설정

FileRenamerX를 사용하려면 다음 API 키가 필요합니다:

### 필요한 API 키

1. **Google Vision API** (필수)
   - OCR(텍스트 인식) 기능에 사용됩니다
   - 비디오 프레임에서 텍스트를 추출합니다

2. **OpenAI ChatGPT API** (필수)
   - 추출된 텍스트를 분석하여 파일명을 생성합니다

### API 키 설정 방법

#### 1. Google Vision API 키 설정

1. **Google Cloud Console에서 서비스 계정 키 생성**
   - [Google Cloud Console](https://console.cloud.google.com/) 접속
   - 프로젝트 선택 또는 새 프로젝트 생성
   - "API 및 서비스" > "사용자 인증 정보" 메뉴로 이동
   - "서비스 계정" 생성 또는 기존 계정 선택
   - "키" 탭에서 "키 추가" > "JSON 만들기" 선택
   - JSON 키 파일이 다운로드됩니다

2. **Vision API 활성화**
   - "API 및 서비스" > "라이브러리" 메뉴로 이동
   - "Cloud Vision API" 검색 후 활성화

3. **키 파일 배치**
   - 다운로드한 JSON 파일을 `vision-api-key/` 폴더에 복사
   - 파일명을 `vision-ocr-454121-572fb601794b.json`으로 변경 (또는 코드에서 지정한 파일명 사용)

**폴더 구조:**
```
FileRenamerX/
└── vision-api-key/
    └── vision-ocr-454121-572fb601794b.json
```

#### 2. ChatGPT API 키 설정

1. **OpenAI API 키 발급**
   - [OpenAI Platform](https://platform.openai.com/) 접속
   - 계정 로그인 또는 회원가입
   - "API keys" 메뉴로 이동
   - "Create new secret key" 클릭하여 새 API 키 생성
   - 생성된 키를 복사 (다시 볼 수 없으므로 안전하게 보관)

2. **키 파일 생성**
   - `chatgpt-api-key/` 폴더 생성
   - `chatgpt_api_key.txt` 파일 생성
   - 복사한 API 키를 파일에 저장 (공백이나 줄바꿈 없이)

**폴더 구조:**
```
FileRenamerX/
└── chatgpt-api-key/
    └── chatgpt_api_key.txt
```

### API 키 확인

프로그램 실행 시 API 키가 올바르게 설정되어 있는지 자동으로 확인합니다. 
API 키가 없거나 잘못된 경우 오류 메시지가 표시됩니다.

### API 사용량 및 비용

- **Google Vision API**: 사용량에 따라 과금됩니다. 자세한 내용은 [Google Cloud 가격 정책](https://cloud.google.com/vision/pricing) 참고
- **OpenAI ChatGPT API**: 사용량에 따라 과금됩니다. 자세한 내용은 [OpenAI 가격 정책](https://openai.com/pricing) 참고

**비용 절감 팁:**
- API 요청 사이에 최소 1초 간격이 자동으로 유지됩니다
- 불필요한 재시작을 피하여 API 호출 횟수를 줄이세요