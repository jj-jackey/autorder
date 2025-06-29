const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const { validateOrderData } = require('../utils/validation');
const { convertToStandardFormat } = require('../utils/converter');
const { uploadFile, downloadFile, saveMappingData, loadMappingData } = require('../utils/supabase');

const router = express.Router();

// 업로드 디렉토리 설정 (개발환경용)
const uploadsDir = path.join(__dirname, '../uploads');

// 개발환경에서만 폴더 생성
if (process.env.NODE_ENV !== 'production' && !fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true });
  console.log('📁 업로드 폴더 생성됨:', uploadsDir);
}

// 파일 업로드 설정 - Supabase Storage 사용 (모든 환경)
const storage = multer.memoryStorage(); // 모든 환경에서 Supabase 사용

// 기존 환경별 스토리지 설정 (주석 처리)
/*
const storage = process.env.NODE_ENV === 'production' 
  ? multer.memoryStorage()  // 프로덕션: 메모리 스토리지 (Supabase로 업로드)
  : multer.diskStorage({    // 개발환경: 디스크 스토리지
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
  limits: { 
    fileSize: 50 * 1024 * 1024, // 50MB로 증가
    fieldSize: 2 * 1024 * 1024   // 2MB
  },
  fileFilter: (req, file, cb) => {
    console.log('🔍 파일 필터 검사:', {
      originalname: file.originalname,
      mimetype: file.mimetype,
      size: file.size
    });
    
    const allowedTypes = /xlsx|xls|csv/;
    const extname = allowedTypes.test(path.extname(file.originalname).toLowerCase());
    const mimetype = allowedTypes.test(file.mimetype) || 
                     file.mimetype === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
                     file.mimetype === 'application/vnd.ms-excel' ||
                     file.mimetype === 'text/csv' ||
                     file.mimetype === 'application/octet-stream'; // 일부 브라우저에서 Excel을 이렇게 인식
    
    if (mimetype && extname) {
      console.log('✅ 파일 필터 통과');
      return cb(null, true);
    } else {
      console.log('❌ 파일 필터 실패:', { mimetype, extname });
      cb(new Error('파일 형식이 지원되지 않습니다. Excel(.xlsx, .xls) 또는 CSV 파일만 업로드 가능합니다.'));
    }
  }
});

// 📁 파일 업로드 및 미리보기
router.post('/upload', upload.single('orderFile'), async (req, res) => {
  try {
    console.log('📁 파일 업로드 요청 수신');
    console.log('🌍 NODE_ENV:', process.env.NODE_ENV);
    
    if (!req.file) {
      console.log('❌ 파일이 업로드되지 않음');
      return res.status(400).json({ error: '파일이 업로드되지 않았습니다.' });
    }

    console.log('📋 업로드된 파일 정보:', {
      originalName: req.file.originalname,
      mimetype: req.file.mimetype,
      size: req.file.size,
      encoding: req.file.encoding
    });

    // 파일명 생성
    const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
    const fileName = 'orderFile-' + uniqueSuffix + path.extname(req.file.originalname);
    
    // Supabase Storage에 업로드 (모든 환경)
    console.log('📤 Supabase Storage 업로드 중...', {
      fileName: fileName,
      fileSize: req.file.buffer.length,
      timestamp: new Date().toISOString()
    });
    
    const uploadResult = await uploadFile(req.file.buffer, fileName);
    if (!uploadResult.success) {
      console.error('❌ Supabase Storage 업로드 실패:', uploadResult.error);
      return res.status(500).json({ 
        error: 'Supabase Storage 업로드 실패', 
        details: uploadResult.error 
      });
    }
    
    const filePath = fileName; // Supabase에서는 파일명만 저장
    const fileBuffer = req.file.buffer;
    
    console.log('✅ Supabase 업로드 성공:', {
      fileName: fileName,
      uploadTime: new Date().toISOString()
    });

    // 기존 환경별 파일 처리 (주석 처리)
    /*
    let filePath;
    let fileBuffer;

    if (process.env.NODE_ENV === 'production') {
      // 프로덕션: Supabase Storage에 업로드
      console.log('📤 Supabase Storage 업로드 중...');
      
      const uploadResult = await uploadFile(req.file.buffer, fileName);
      if (!uploadResult.success) {
        return res.status(500).json({ 
          error: 'Supabase Storage 업로드 실패', 
          details: uploadResult.error 
        });
      }
      
      filePath = fileName; // Supabase에서는 파일명만 저장
      fileBuffer = req.file.buffer;
      
      console.log('✅ Supabase 업로드 성공:', fileName);
    } else {
      // 개발환경: 로컬 디스크 저장
      filePath = req.file.path;
      fileBuffer = fs.readFileSync(filePath);
      
      console.log('✅ 로컬 파일 저장 성공:', {
        originalName: req.file.originalname,
        filename: req.file.filename,
        size: req.file.size,
        path: filePath
      });
    }
    */

    const fileExtension = path.extname(req.file.originalname).toLowerCase();
    
    let previewData = [];
    let headers = [];

    if (fileExtension === '.csv') {
      // CSV 파일 처리
      const csvData = fileBuffer.toString('utf8');
      const lines = csvData.split('\n').filter(line => line.trim());
      
      if (lines.length > 0) {
        headers = lines[0].split(',').map(h => h.trim());
        previewData = lines.slice(1, 21).map(line => {
          const values = line.split(',').map(v => v.trim());
          const rowData = {};
          headers.forEach((header, index) => {
            rowData[header] = values[index] || '';
          });
          return rowData;
        });
      }
    } else {
      // Excel 파일 처리 - 개선된 로직 사용
      try {
        console.log('🔄 Excel 파일 처리 시작:', {
          fileSize: fileBuffer.length,
          timestamp: new Date().toISOString()
        });

        // 프로덕션 환경에서는 /tmp 폴더 사용
        const tempDir = process.env.NODE_ENV === 'production' 
          ? '/tmp' 
          : path.join(__dirname, '../uploads');
        
        // 임시 파일로 저장하여 개선된 readExcelFile 함수 사용
        const tempFileName = `temp_${Date.now()}_${Math.random().toString(36).substr(2, 9)}.xlsx`;
        const tempFilePath = path.join(tempDir, tempFileName);
        
        console.log('📁 임시 파일 생성:', tempFilePath);
        
        // 폴더가 없으면 생성 (로컬에서만)
        if (process.env.NODE_ENV !== 'production' && !fs.existsSync(tempDir)) {
          fs.mkdirSync(tempDir, { recursive: true });
          console.log('📁 임시 폴더 생성됨:', tempDir);
        }
        
        try {
          fs.writeFileSync(tempFilePath, fileBuffer);
          console.log('✅ 임시 파일 쓰기 완료:', tempFilePath);
        } catch (writeError) {
          console.error('❌ 임시 파일 쓰기 실패:', writeError.message);
          throw writeError;
        }
        
        // 개선된 Excel 읽기 함수 사용
        const { readExcelFile } = require('../utils/converter');
        console.log('🔄 Excel 파일 읽기 시작...');
        
        const excelData = await readExcelFile(tempFilePath);
        
        headers = excelData.headers;
        previewData = excelData.data.slice(0, 20); // 상위 20행만
        
        console.log('✅ Excel 파일 처리 완료:', {
          worksheets: '자동 선택됨',
          headers: headers.length,
          dataRows: excelData.data.length,
          previewRows: previewData.length,
          processingTime: new Date().toISOString()
        });
        
        // 임시 파일 삭제
        try {
          if (fs.existsSync(tempFilePath)) {
            fs.unlinkSync(tempFilePath);
            console.log('🗑️ 임시 파일 삭제 완료:', tempFilePath);
          }
        } catch (deleteError) {
          console.warn('⚠️ 임시 파일 삭제 실패 (무시됨):', deleteError.message);
        }
        
      } catch (excelError) {
        console.error('❌ 개선된 Excel 처리 실패, 기본 방식 사용:', {
          error: excelError.message,
          stack: excelError.stack?.split('\n')[0],
          fileName: req.file.originalname,
          fileSize: fileBuffer.length
        });
        
        // 개선된 처리가 실패하면 기본 방식으로 fallback
        try {
          console.log('🔄 기본 Excel 처리 방식으로 fallback...');
          const workbook = new ExcelJS.Workbook();
          await workbook.xlsx.load(fileBuffer);
          const worksheet = workbook.getWorksheet(1);
          
          if (worksheet) {
            const firstRow = worksheet.getRow(1);
            headers = [];
            firstRow.eachCell((cell, colNumber) => {
              headers.push(cell.value ? cell.value.toString() : `컬럼${colNumber}`);
            });

            // 상위 20행까지 미리보기 데이터 생성
            for (let rowNumber = 2; rowNumber <= Math.min(21, worksheet.rowCount); rowNumber++) {
              const row = worksheet.getRow(rowNumber);
              const rowData = {};
              
              headers.forEach((header, index) => {
                const cell = row.getCell(index + 1);
                rowData[header] = cell.value ? cell.value.toString() : '';
              });
              
              previewData.push(rowData);
            }
            
            console.log('✅ 기본 Excel 처리 완료:', {
              headers: headers.length,
              previewRows: previewData.length
            });
          }
        } catch (fallbackError) {
          console.error('❌ 기본 Excel 처리도 실패:', fallbackError.message);
          throw new Error(`Excel 파일 처리 실패: ${fallbackError.message}`);
        }
      }
    }

    // 데이터 검증
    const validation = validateOrderData(previewData, headers);

    console.log('✅ 파일 처리 완료:', {
      headers: headers.length,
      previewRows: previewData.length,
      isValid: validation.isValid
    });

    res.json({
      success: true,
      fileName: req.file.originalname,
      fileId: fileName, // 모든 환경에서 Supabase 파일명 사용
      headers: headers,
      previewData: previewData,
      totalRows: previewData.length,
      validation: validation,
      message: `파일이 성공적으로 업로드되었습니다. ${previewData.length}행의 데이터를 확인했습니다.`
    });

    // 기존 환경별 fileId 설정 (주석 처리)
    // fileId: process.env.NODE_ENV === 'production' ? fileName : req.file.filename,

  } catch (error) {
    console.error('❌ 파일 업로드 오류:', {
      error: error.message,
      stack: error.stack?.split('\n')[0],
      fileName: req.file?.originalname,
      fileSize: req.file?.size,
      timestamp: new Date().toISOString()
    });
    
    res.status(500).json({ 
      error: '파일 처리 중 오류가 발생했습니다.', 
      details: error.message,
      fileName: req.file?.originalname || 'unknown'
    });
  }
});

// 🔄 필드 매핑 설정 저장
router.post('/mapping', async (req, res) => {
  try {
    const { mappingName, sourceFields, targetFields, mappingRules, fixedValues } = req.body;
    
    // 매핑 규칙 데이터
    const mappingData = {
      name: mappingName,
      createdAt: new Date().toISOString(),
      sourceFields,
      targetFields,
      rules: mappingRules,
      fixedValues: fixedValues || {} // 고정값 추가
    };

    // Supabase Storage에 저장 (모든 환경)
    const saveResult = await saveMappingData(mappingName, mappingData);
    if (!saveResult.success) {
      return res.status(500).json({ 
        error: 'Supabase Storage 매핑 저장 실패', 
        details: saveResult.error 
      });
    }
    console.log('✅ Supabase 매핑 저장 성공:', mappingName);

    // 기존 환경별 매핑 저장 (주석 처리)
    /*
    if (process.env.NODE_ENV === 'production') {
      // 프로덕션: Supabase Storage에 저장
      const saveResult = await saveMappingData(mappingName, mappingData);
      if (!saveResult.success) {
        return res.status(500).json({ 
          error: 'Supabase Storage 매핑 저장 실패', 
          details: saveResult.error 
        });
      }
      console.log('✅ Supabase 매핑 저장 성공:', mappingName);
    } else {
      // 개발환경: 로컬 파일로 저장
      const mappingPath = path.join(__dirname, '../file/mappings');
      
      if (!fs.existsSync(mappingPath)) {
        fs.mkdirSync(mappingPath, { recursive: true });
      }

      fs.writeFileSync(
        path.join(mappingPath, `${mappingName}.json`),
        JSON.stringify(mappingData, null, 2)
      );
      console.log('✅ 로컬 매핑 저장 성공:', path.join(mappingPath, `${mappingName}.json`));
    }
    */

    res.json({
      success: true,
      message: '매핑 규칙이 저장되었습니다.',
      mappingId: mappingName
    });

  } catch (error) {
    console.error('❌ 매핑 저장 오류:', error);
    res.status(500).json({ 
      error: '매핑 저장 중 오류가 발생했습니다.', 
      details: error.message 
    });
  }
});

// 📋 발주서 생성
router.post('/generate', async (req, res) => {
  try {
    const { fileId, mappingId, templateType } = req.body;
    
    console.log('📋 발주서 생성 요청:', { fileId, mappingId, templateType });
    
    // Supabase Storage에서 파일 다운로드 (모든 환경)
    console.log('📥 Supabase에서 파일 다운로드 중:', fileId);
    
    const downloadResult = await downloadFile(fileId);
    if (!downloadResult.success) {
      console.log('❌ Supabase 파일 다운로드 실패:', downloadResult.error);
      return res.status(404).json({ error: '업로드된 파일을 찾을 수 없습니다.' });
    }
    
    // 임시 파일로 저장 (변환 처리용)
    const uploadedFilePath = path.join(__dirname, '../uploads', fileId);
    // uploads 폴더가 없으면 생성
    if (!fs.existsSync(path.join(__dirname, '../uploads'))) {
      fs.mkdirSync(path.join(__dirname, '../uploads'), { recursive: true });
    }
    fs.writeFileSync(uploadedFilePath, downloadResult.data);
    console.log('✅ Supabase 파일 다운로드 완료');

    // 매핑 규칙 로드
    let mappingRules = {};
    const mappingResult = await loadMappingData(mappingId);
    if (mappingResult.success) {
      mappingRules = mappingResult.data;
      console.log('✅ Supabase 매핑 로드 완료');
    }

    // 기존 환경별 파일 처리 (주석 처리)
    /*
    let uploadedFilePath;
    let mappingRules = {};

    if (process.env.NODE_ENV === 'production') {
      // 프로덕션: Supabase Storage에서 파일 다운로드
      console.log('📥 Supabase에서 파일 다운로드 중:', fileId);
      
      const downloadResult = await downloadFile(fileId);
      if (!downloadResult.success) {
        console.log('❌ Supabase 파일 다운로드 실패:', downloadResult.error);
        return res.status(404).json({ error: '업로드된 파일을 찾을 수 없습니다.' });
      }
      
      // 임시 파일로 저장 (변환 처리용)
      uploadedFilePath = path.join('/tmp', fileId);
      fs.writeFileSync(uploadedFilePath, downloadResult.data);
      console.log('✅ Supabase 파일 다운로드 완료');

      // 매핑 규칙 로드
      const mappingResult = await loadMappingData(mappingId);
      if (mappingResult.success) {
        mappingRules = mappingResult.data;
        console.log('✅ Supabase 매핑 로드 완료');
      }
    } else {
      // 개발환경: 로컬 파일 시스템 사용
      uploadedFilePath = path.join(uploadsDir, fileId);
      if (!fs.existsSync(uploadedFilePath)) {
        console.log('❌ 업로드된 파일을 찾을 수 없음:', uploadedFilePath);
        return res.status(404).json({ error: '업로드된 파일을 찾을 수 없습니다.' });
      }

      // 매핑 규칙 로드
      const mappingPath = path.join(__dirname, '../file/mappings', `${mappingId}.json`);
      if (fs.existsSync(mappingPath)) {
        mappingRules = JSON.parse(fs.readFileSync(mappingPath, 'utf8'));
      }
    }
    */

    // 템플릿 파일 로드
    const templatePath = path.join(__dirname, '../file/porder_template.xlsx');
    
    // 데이터 변환 및 발주서 생성
    const result = await convertToStandardFormat(uploadedFilePath, templatePath, mappingRules);
    
    console.log('✅ 발주서 생성 완료:', result.fileName);

    // 생성된 발주서를 Supabase Storage에 업로드 (모든 환경)
    const generatedFileBuffer = fs.readFileSync(result.filePath);
    const uploadResult = await uploadFile(generatedFileBuffer, result.fileName, 'generated');
    
    if (uploadResult.success) {
      console.log('✅ 생성된 발주서 Supabase 업로드 완료');
      // 임시 파일들 정리
      if (fs.existsSync(uploadedFilePath)) fs.unlinkSync(uploadedFilePath);
      if (fs.existsSync(result.filePath)) fs.unlinkSync(result.filePath);
    } else {
      console.error('❌ 생성된 발주서 Supabase 업로드 실패:', uploadResult.error);
    }

    const downloadUrl = `/api/orders/download/${result.fileName}`;

    // 기존 환경별 업로드 처리 (주석 처리)
    /*
    let downloadUrl = `/api/orders/download/${result.fileName}`;
    
    if (process.env.NODE_ENV === 'production') {
      // 프로덕션: 생성된 발주서를 Supabase Storage에 업로드
      const generatedFileBuffer = fs.readFileSync(result.filePath);
      const uploadResult = await uploadFile(generatedFileBuffer, result.fileName, 'generated');
      
      if (uploadResult.success) {
        console.log('✅ 생성된 발주서 Supabase 업로드 완료');
        // 임시 파일들 정리
        if (fs.existsSync(uploadedFilePath)) fs.unlinkSync(uploadedFilePath);
        if (fs.existsSync(result.filePath)) fs.unlinkSync(result.filePath);
      } else {
        console.error('❌ 생성된 발주서 Supabase 업로드 실패:', uploadResult.error);
      }
    }
    */
    
    res.json({
      success: true,
      generatedFile: result.fileName,
      downloadUrl: downloadUrl,
      processedRows: result.processedRows,
      errors: result.errors,
      message: '발주서가 성공적으로 생성되었습니다.'
    });

  } catch (error) {
    console.error('❌ 발주서 생성 오류:', error);
    res.status(500).json({ 
      error: '발주서 생성 중 오류가 발생했습니다.', 
      details: error.message 
    });
  }
});

// 📥 생성된 발주서 다운로드
router.get('/download/:fileName', async (req, res) => {
  try {
    const fileName = req.params.fileName;
    
    console.log('📥 다운로드 요청:', fileName);
    
    // Supabase Storage에서 다운로드 (모든 환경)
    const downloadResult = await downloadFile(fileName, 'generated');
    
    if (!downloadResult.success) {
      console.log('❌ Supabase 파일 다운로드 실패:', downloadResult.error);
      return res.status(404).json({ error: '파일을 찾을 수 없습니다.' });
    }

    // 파일 헤더 설정 및 전송
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(downloadResult.data);
    
    console.log('✅ Supabase 파일 다운로드 완료:', fileName);

    // 기존 환경별 다운로드 처리 (주석 처리)
    /*
    if (process.env.NODE_ENV === 'production') {
      // 프로덕션: Supabase Storage에서 다운로드
      const downloadResult = await downloadFile(fileName, 'generated');
      
      if (!downloadResult.success) {
        console.log('❌ Supabase 파일 다운로드 실패:', downloadResult.error);
        return res.status(404).json({ error: '파일을 찾을 수 없습니다.' });
      }

      // 파일 헤더 설정 및 전송
      res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`);
      res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
      res.send(downloadResult.data);
      
      console.log('✅ Supabase 파일 다운로드 완료:', fileName);
    } else {
      // 개발환경: 로컬 파일 시스템에서 다운로드
      const filePath = path.join(uploadsDir, fileName);
      
      if (!fs.existsSync(filePath)) {
        console.log('❌ 다운로드 파일을 찾을 수 없음:', filePath);
        return res.status(404).json({ error: '파일을 찾을 수 없습니다.' });
      }

      res.download(filePath, fileName, (err) => {
        if (err) {
          console.error('❌ 파일 다운로드 오류:', err);
          res.status(500).json({ error: '파일 다운로드 중 오류가 발생했습니다.' });
        } else {
          console.log('✅ 파일 다운로드 완료:', fileName);
        }
      });
    }
    */

  } catch (error) {
    console.error('❌ 다운로드 오류:', error);
    res.status(500).json({ 
      error: '파일 다운로드 중 오류가 발생했습니다.', 
      details: error.message 
    });
  }
});

// 📋 템플릿 목록 조회 API
router.get('/templates', (req, res) => {
  try {
    const templatesConfigPath = path.join(__dirname, '../file/templates-config.json');
    
    if (!fs.existsSync(templatesConfigPath)) {
      return res.status(404).json({ 
        error: '템플릿 설정 파일을 찾을 수 없습니다.' 
      });
    }
    
    const templatesConfig = JSON.parse(fs.readFileSync(templatesConfigPath, 'utf8'));
    
    // 각 템플릿의 파일 존재 여부 확인
    const templates = Object.keys(templatesConfig.templates).map(key => {
      const template = templatesConfig.templates[key];
      const templateFilePath = path.join(__dirname, '../file', template.file);
      
      return {
        id: key,
        name: template.name,
        description: template.description,
        file: template.file,
        fields: template.fields,
        available: fs.existsSync(templateFilePath)
      };
    });
    
    res.json({
      success: true,
      templates: templates
    });
    
  } catch (error) {
    console.error('템플릿 목록 조회 오류:', error);
    res.status(500).json({ 
      error: '템플릿 목록을 불러오는 중 오류가 발생했습니다.',
      details: error.message 
    });
  }
});

// 📝 직접 입력 데이터로 발주서 생성
router.post('/generate-direct', async (req, res) => {
  try {
    const { mappingId, inputData, templateType } = req.body;
    
    console.log('📝 직접 입력 발주서 생성 요청:', { mappingId, inputData, templateType });
    
    // 매핑 규칙 로드
    let mappingRules = {};
    const mappingResult = await loadMappingData(mappingId);
    if (mappingResult.success) {
      mappingRules = mappingResult.data;
      console.log('✅ Supabase 매핑 로드 완료');
    }
    
    // 템플릿 파일 로드
    const templatePath = path.join(__dirname, '../file/porder_template.xlsx');
    
    // 직접 입력 데이터를 표준 형식으로 변환
    const { convertDirectInputToStandardFormat } = require('../utils/converter');
    const result = await convertDirectInputToStandardFormat(templatePath, inputData, mappingRules);
    
    console.log('✅ 직접 입력 발주서 생성 완료:', result.fileName);

    // 생성된 발주서를 Supabase Storage에 업로드
    const generatedFileBuffer = fs.readFileSync(result.filePath);
    const uploadResult = await uploadFile(generatedFileBuffer, result.fileName, 'generated');
    
    if (uploadResult.success) {
      console.log('✅ 생성된 발주서 Supabase 업로드 완료');
      // 임시 파일 정리
      if (fs.existsSync(result.filePath)) fs.unlinkSync(result.filePath);
    } else {
      console.error('❌ 생성된 발주서 Supabase 업로드 실패:', uploadResult.error);
    }

    const downloadUrl = `/api/orders/download/${result.fileName}`;

    res.json({
      success: true,
      message: '직접 입력으로 발주서가 성공적으로 생성되었습니다.',
      generatedFile: result.fileName,
      downloadUrl: downloadUrl,
      inputData: inputData,
      processedRows: 1
    });

  } catch (error) {
    console.error('❌ 직접 입력 발주서 생성 오류:', error);
    res.status(500).json({ 
      error: '직접 입력 발주서 생성 중 오류가 발생했습니다.', 
      details: error.message 
    });
  }
});

module.exports = router; 