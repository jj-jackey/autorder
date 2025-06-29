# 📋 자동 발주서 변환 시스템

라이브커머스 후 다양한 플랫폼에서 발생하는 주문서를 **표준 발주서로 자동 변환**하고 **이메일로 자동 전송**하는 시스템입니다.

## ✨ 주요 기능

### 🚀 **1단계: 핵심 기능 (MVP)**
- 📁 **파일 업로드 및 미리보기**: Excel(.xlsx, .xls), CSV 파일 지원
- 🔗 **필드 매핑**: 드래그 앤 드롭 방식으로 주문서-발주서 컬럼 연결
- 📋 **발주서 자동 생성**: 템플릿 기반 표준 발주서 생성
- 📧 **이메일 자동 전송**: 즉시/예약 전송 옵션

### 🛡️ **데이터 검증 기능**
- ✅ 필수 필드 검증
- 📞 전화번호 형식 검증  
- 💰 단가/수량 논리적 오류 검증
- 🔍 중복 데이터 감지
- 📊 상세 오류 리포트

## 🖥️ **사용법**

### 1. **설치 및 실행**

```bash
# 프로젝트 클론 또는 다운로드
cd agorder

# 의존성 설치
npm install

# 환경 변수 설정
cp .env.example .env
# .env 파일을 열어서 이메일 설정 입력

# 서버 실행
npm start

# 또는 개발 모드로 실행
npm run dev
```

### 2. **웹 브라우저 접속**
```
http://localhost:3000
```

### 3. **사용 단계**

#### 📁 **1단계: 파일 업로드**
- 주문서 파일(Excel 또는 CSV)을 드래그 앤 드롭
- 상위 20행 데이터 미리보기 확인
- 데이터 검증 결과 확인

#### 🔗 **2단계: 필드 매핑**
- 주문서 컬럼을 선택 후 발주서 컬럼과 연결
- 매핑 규칙 저장 (재사용 가능)

#### 📋 **3단계: 발주서 생성**
- "발주서 생성" 버튼 클릭
- 변환 결과 확인 및 다운로드

#### 📧 **4단계: 이메일 전송**
- 받는 사람, 제목, 내용 입력
- 즉시 전송 또는 예약 전송 선택
- 전송 결과 및 이력 확인

## 📋 **필요한 환경 설정**

### Gmail 사용 시
1. Gmail 계정에서 "2단계 인증" 활성화
2. "앱 비밀번호" 생성 
3. `.env` 파일에 설정:
```
EMAIL_USER=your-email@gmail.com
EMAIL_PASS=your-app-password
```

### 다른 이메일 서비스 사용 시
```
EMAIL_USER=your-email@company.com
EMAIL_PASS=your-password
SMTP_HOST=smtp.company.com
SMTP_PORT=587
```

## 📂 **프로젝트 구조**

```
agorder/
├── server.js              # 메인 서버
├── package.json           # 프로젝트 설정
├── .env                   # 환경 변수
├── routes/                # API 라우트
│   ├── orders.js         # 주문서 처리 API
│   └── email.js          # 이메일 API
├── utils/                 # 유틸리티 함수
│   ├── validation.js     # 데이터 검증
│   └── converter.js      # 파일 변환
├── public/               # 프론트엔드
│   ├── index.html        # 메인 페이지
│   └── app.js           # 클라이언트 로직
├── uploads/              # 업로드된 파일
├── file/                 # 템플릿 및 설정
│   ├── porder_template.xlsx  # 발주서 템플릿
│   ├── mappings/            # 매핑 규칙
│   └── email-templates/     # 이메일 템플릿
└── README.md             # 사용법 안내
```

## 🔧 **기술 스택**

- **Backend**: Node.js, Express
- **Frontend**: HTML5, CSS3, Vanilla JavaScript
- **Excel 처리**: ExcelJS
- **이메일**: Nodemailer
- **데이터 검증**: Yup
- **파일 업로드**: Multer

## 📈 **로드맵**

### ✅ **1단계 완료 (현재)**
- 기본 업로드-변환-전송 흐름
- 데이터 검증 및 오류 리포트
- 이메일 전송 및 이력 관리

### 🚧 **2단계 예정**
- 사용자 인증 시스템
- 템플릿 CRUD 관리
- 대시보드 고도화

### 🌟 **3단계 예정**
- 다수 거래처 동시 발송
- 외부 시스템 연동 API
- 고급 Excel 양식 지원

## 🐛 **문제 해결**

### 자주 발생하는 오류

1. **파일 업로드 실패**
   - 파일 크기 10MB 이하 확인
   - 지원 형식(.xlsx, .xls, .csv) 확인

2. **이메일 전송 실패**
   - .env 파일의 이메일 설정 확인
   - Gmail 앱 비밀번호 사용 확인

3. **발주서 생성 오류**
   - 템플릿 파일 경로 확인
   - 매핑 규칙 설정 확인

## 📞 **지원 및 문의**

시스템 사용 중 문제가 발생하면:
1. 브라우저 개발자 도구 콘솔 확인
2. 서버 로그 확인 (`npm start` 터미널)
3. GitHub Issues 또는 이메일로 문의

---

**🎯 목표**: "올리고 → 바꾸고 → 보내면 끝" - 3Click으로 완성되는 자동화 

## 🚀 배포 시 중요사항

### Render 배포 문제 해결
Render와 같은 클라우드 플랫폼은 **Ephemeral Filesystem**을 사용하여 파일이 임시로만 저장되고 재배포 시 사라집니다. 

**해결책**: `Supabase Storage` 사용
- 프로덕션 환경에서는 모든 파일이 Supabase Storage에 저장됩니다
- 개발 환경에서는 기존처럼 로컬 파일 시스템을 사용합니다

📖 **[Supabase 설정 가이드 보기](./SUPABASE_SETUP.md)**

## 📋 기능

- **📁 파일 업로드**: Excel(.xlsx, .xls), CSV 파일 지원
- **👀 데이터 미리보기**: 업로드된 파일의 상위 20행 미리보기
- **🔄 필드 매핑**: 소스 파일의 컬럼을 표준 발주서 형식으로 매핑
- **📋 발주서 생성**: 매핑 규칙에 따라 자동으로 표준 발주서 생성
- **📧 메일 전송**: 생성된 발주서를 지정된 이메일로 자동 전송
- **☁️ 클라우드 스토리지**: Supabase Storage를 통한 영구 파일 저장

## 🛠️ 설치 및 실행

### 1. 프로젝트 클론
```bash
git clone <repository-url>
cd agorder_v1
```

### 2. 패키지 설치
```bash
npm install
```

### 3. 환경 변수 설정
`.env` 파일을 생성하고 다음 내용을 추가:

```env
# 기본 설정
NODE_ENV=development
PORT=3000

# Supabase 설정 (필수)
SUPABASE_URL=https://your-project-id.supabase.co
SUPABASE_ANON_KEY=your-supabase-anon-key

# Gmail 설정 (메일 전송용)
GMAIL_USER=your-email@gmail.com
GMAIL_APP_PASSWORD=your-app-password
```

### 4. Supabase 설정
📖 **[상세 설정 가이드](./SUPABASE_SETUP.md)** 참고

### 5. 개발 서버 실행
```bash
npm run dev
```

### 6. 프로덕션 빌드
```bash
npm start
```

## 🌐 배포

### Render 배포
1. GitHub에 코드 푸시
2. Render에서 새 Web Service 생성
3. 환경 변수 설정:
   - `NODE_ENV=production`
   - `SUPABASE_URL=your-url`
   - `SUPABASE_ANON_KEY=your-key`
   - Gmail 설정 등
4. 배포 완료

## 📚 API 엔드포인트

| Method | Endpoint | Description |
|--------|----------|-------------|
| `POST` | `/api/orders/upload` | 파일 업로드 및 미리보기 |
| `POST` | `/api/orders/mapping` | 필드 매핑 규칙 저장 |
| `POST` | `/api/orders/generate` | 발주서 생성 |
| `GET` | `/api/orders/download/:fileName` | 생성된 발주서 다운로드 |
| `POST` | `/api/email/send` | 이메일 전송 |

## 🗂️ 프로젝트 구조

```
agorder_v1/
├── public/               # 클라이언트 파일
├── routes/              # API 라우트
├── utils/               # 유틸리티 함수
│   ├── converter.js     # 데이터 변환 로직
│   ├── validation.js    # 데이터 검증
│   └── supabase.js      # Supabase Storage 유틸리티
├── uploads/             # 로컬 업로드 폴더 (개발환경용)
├── file/                # 템플릿 및 매핑 파일
├── server.js            # 메인 서버 파일
└── SUPABASE_SETUP.md    # Supabase 설정 가이드
```

## ⚙️ 환경별 동작 방식

### 개발환경 (`NODE_ENV=development`)
- 파일이 로컬 `uploads/` 폴더에 저장
- 매핑 데이터가 `file/mappings/` 폴더에 저장
- 디스크 스토리지 사용

### 프로덕션환경 (`NODE_ENV=production`)
- 파일이 Supabase Storage에 저장
- 매핑 데이터가 Supabase Storage에 저장
- 메모리 스토리지 → 클라우드 업로드 방식

## 🚨 문제 해결

### 파일 업로드 실패
1. Supabase URL과 API Key 확인
2. 버킷이 올바르게 생성되었는지 확인
3. 네트워크 연결 상태 확인

### 이메일 전송 실패
1. Gmail 앱 비밀번호 확인
2. 환경 변수 설정 확인

## 📄 라이선스

MIT License 