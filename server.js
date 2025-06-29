const express = require('express');
const multer = require('multer');
const path = require('path');
const cors = require('cors');
const fs = require('fs');
const { createClient } = require('@supabase/supabase-js');
require('dotenv').config();

const app = express();
const PORT = process.env.PORT || 3000;

// Supabase 클라이언트 초기화
const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_ANON_KEY
);

// uploads 폴더는 임시 파일 처리용으로만 사용 (Supabase Storage가 메인)
const uploadsDir = path.join(__dirname, 'uploads');

// 기존 자동 폴더 생성 코드 (주석 처리 - 필요시에만 생성)
/*
if (process.env.NODE_ENV !== 'production' && !fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true });
  console.log('📁 uploads 폴더가 생성되었습니다:', uploadsDir);
}
*/

// 미들웨어 설정
app.use(cors());
app.use(express.json());
app.use(express.static('public'));

// 파일 업로드 설정 - Supabase Storage 사용 (로컬에서도 테스트)
const storage = multer.memoryStorage(); // 모든 환경에서 Supabase 사용

// 기존 로컬 파일 시스템 설정 (주석 처리)
/*
const storage = process.env.NODE_ENV === 'production' 
  ? multer.memoryStorage()  // 프로덕션: 메모리에 임시 저장 후 Supabase로 업로드
  : multer.diskStorage({    // 개발환경: 디스크 저장
      destination: function (req, file, cb) {
        cb(null, uploadsDir);
      },
      filename: function (req, file, cb) {
        const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
        cb(null, file.fieldname + '-' + uniqueSuffix + path.extname(file.originalname));
      }
    });
*/

const upload = multer({ 
  storage: storage,
  fileFilter: (req, file, cb) => {
    console.log('🔍 서버 파일 필터 검사:', {
      originalname: file.originalname,
      mimetype: file.mimetype
    });
    
    const allowedTypes = /xlsx|xls|csv/;
    const extname = allowedTypes.test(path.extname(file.originalname).toLowerCase());
    const mimetype = allowedTypes.test(file.mimetype) || 
                     file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
                     file.mimetype === 'application/vnd.ms-excel' ||
                     file.mimetype === 'text/csv' ||
                     file.mimetype === 'application/octet-stream'; // 일부 브라우저에서 Excel을 이렇게 인식
    
    if (mimetype && extname) {
      console.log('✅ 서버 파일 필터 통과');
      return cb(null, true);
    } else {
      console.log('❌ 서버 파일 필터 실패:', { mimetype, extname });
      cb(new Error('파일 형식이 지원되지 않습니다. Excel(.xlsx, .xls) 또는 CSV 파일만 업로드 가능합니다.'));
    }
  },
  limits: { 
    fileSize: 50 * 1024 * 1024, // 50MB로 증가
    fieldSize: 2 * 1024 * 1024   // 2MB
  }
});

// API 라우트
const orderRoutes = require('./routes/orders');
const emailRoutes = require('./routes/email');

app.use('/api/orders', orderRoutes);
app.use('/api/email', emailRoutes);

// 홈페이지 라우트
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// 에러 핸들링
app.use((error, req, res, next) => {
  console.error('🚨 서버 에러:', {
    error: error.message,
    code: error.code,
    type: error.constructor.name,
    timestamp: new Date().toISOString()
  });
  
  if (error instanceof multer.MulterError) {
    if (error.code === 'LIMIT_FILE_SIZE') {
      return res.status(400).json({ 
        error: '파일 크기가 너무 큽니다. 50MB 이하의 파일을 업로드해주세요.',
        code: 'LIMIT_FILE_SIZE'
      });
    }
    if (error.code === 'LIMIT_FIELD_SIZE') {
      return res.status(400).json({ 
        error: '필드 크기가 너무 큽니다.',
        code: 'LIMIT_FIELD_SIZE'
      });
    }
  }
  
  res.status(500).json({ 
    error: error.message,
    code: error.code || 'UNKNOWN_ERROR'
  });
});

app.listen(PORT, () => {
  console.log(`🚀 서버가 포트 ${PORT}에서 실행 중입니다.`);
  console.log(`📁 파일 업로드: http://localhost:${PORT}`);
  console.log(`☁️ 스토리지: Supabase Storage (모든 환경)`);
  console.log(`🔗 Supabase URL: ${process.env.SUPABASE_URL ? '✅ 연결됨' : '❌ 설정안됨'}`);
}); 