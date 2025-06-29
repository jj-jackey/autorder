const { createClient } = require('@supabase/supabase-js');

// Supabase 클라이언트 초기화
const supabase = createClient(
  process.env.SUPABASE_URL,
  process.env.SUPABASE_ANON_KEY
);

/**
 * 파일을 Supabase Storage에 업로드
 * @param {Buffer} fileBuffer - 파일 버퍼
 * @param {string} fileName - 파일명
 * @param {string} bucket - 버킷명 (기본값: 'uploads')
 * @returns {Promise<{success: boolean, data?: any, error?: string}>}
 */
async function uploadFile(fileBuffer, fileName, bucket = 'uploads') {
  try {
    console.log('📤 Supabase Storage 업로드 시작:', fileName);
    
    // 파일 확장자에 따른 MIME 타입 설정
    const getContentType = (fileName) => {
      const ext = fileName.toLowerCase().split('.').pop();
      switch (ext) {
        case 'xlsx':
          return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        case 'xls':
          return 'application/vnd.ms-excel';
        case 'csv':
          return 'text/csv';
        case 'json':
          return 'application/json';
        default:
          return 'application/octet-stream';
      }
    };
    
    const { data, error } = await supabase.storage
      .from(bucket)
      .upload(`files/${fileName}`, fileBuffer, {
        cacheControl: '3600',
        upsert: false,
        contentType: getContentType(fileName)
      });

    if (error) {
      console.error('❌ Supabase 업로드 오류:', error);
      return { success: false, error: error.message };
    }

    console.log('✅ Supabase 업로드 성공:', data.path);
    return { success: true, data };
  } catch (error) {
    console.error('❌ 업로드 예외 오류:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Supabase Storage에서 파일 다운로드
 * @param {string} fileName - 파일명
 * @param {string} bucket - 버킷명 (기본값: 'uploads')
 * @returns {Promise<{success: boolean, data?: Buffer, error?: string}>}
 */
async function downloadFile(fileName, bucket = 'uploads') {
  try {
    console.log('📥 Supabase Storage 다운로드 시작:', fileName);
    
    const { data, error } = await supabase.storage
      .from(bucket)
      .download(`files/${fileName}`);

    if (error) {
      console.error('❌ Supabase 다운로드 오류:', error);
      return { success: false, error: error.message };
    }

    // Blob을 Buffer로 변환
    const arrayBuffer = await data.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    console.log('✅ Supabase 다운로드 성공');
    return { success: true, data: buffer };
  } catch (error) {
    console.error('❌ 다운로드 예외 오류:', error);
    return { success: false, error: error.message };
  }
}

/**
 * Supabase Storage에서 파일 삭제
 * @param {string} fileName - 파일명
 * @param {string} bucket - 버킷명 (기본값: 'uploads')
 * @returns {Promise<{success: boolean, error?: string}>}
 */
async function deleteFile(fileName, bucket = 'uploads') {
  try {
    console.log('🗑️ Supabase Storage 파일 삭제:', fileName);
    
    const { error } = await supabase.storage
      .from(bucket)
      .remove([`files/${fileName}`]);

    if (error) {
      console.error('❌ Supabase 삭제 오류:', error);
      return { success: false, error: error.message };
    }

    console.log('✅ Supabase 파일 삭제 성공');
    return { success: true };
  } catch (error) {
    console.error('❌ 삭제 예외 오류:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 파일의 공개 URL 생성
 * @param {string} fileName - 파일명
 * @param {string} bucket - 버킷명 (기본값: 'uploads')
 * @returns {string} 공개 URL
 */
function getPublicUrl(fileName, bucket = 'uploads') {
  const { data } = supabase.storage
    .from(bucket)
    .getPublicUrl(`files/${fileName}`);
  
  return data.publicUrl;
}

/**
 * 매핑 데이터를 Supabase에 저장
 * @param {string} mappingName - 매핑명
 * @param {Object} mappingData - 매핑 데이터
 * @returns {Promise<{success: boolean, data?: any, error?: string}>}
 */
async function saveMappingData(mappingName, mappingData) {
  try {
    console.log('💾 매핑 데이터 저장:', mappingName);
    
    const jsonData = JSON.stringify(mappingData, null, 2);
    const buffer = Buffer.from(jsonData, 'utf8');
    
    const result = await uploadFile(buffer, `${mappingName}.json`, 'mappings');
    
    if (result.success) {
      console.log('✅ 매핑 데이터 저장 성공');
    }
    
    return result;
  } catch (error) {
    console.error('❌ 매핑 저장 오류:', error);
    return { success: false, error: error.message };
  }
}

/**
 * 매핑 데이터를 Supabase에서 로드
 * @param {string} mappingName - 매핑명
 * @returns {Promise<{success: boolean, data?: Object, error?: string}>}
 */
async function loadMappingData(mappingName) {
  try {
    console.log('📖 매핑 데이터 로드:', mappingName);
    
    const result = await downloadFile(`${mappingName}.json`, 'mappings');
    
    if (!result.success) {
      return result;
    }
    
    const jsonData = result.data.toString('utf8');
    const mappingData = JSON.parse(jsonData);
    
    console.log('✅ 매핑 데이터 로드 성공');
    return { success: true, data: mappingData };
  } catch (error) {
    console.error('❌ 매핑 로드 오류:', error);
    return { success: false, error: error.message };
  }
}

module.exports = {
  uploadFile,
  downloadFile,
  deleteFile,
  getPublicUrl,
  saveMappingData,
  loadMappingData,
  supabase
}; 