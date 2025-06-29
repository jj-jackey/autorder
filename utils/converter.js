const ExcelJS = require('exceljs');
const path = require('path');
const fs = require('fs');

// 생성된 파일 저장 디렉토리 설정
const getOutputDir = () => {
  return process.env.NODE_ENV === 'production' 
    ? path.join('/tmp', 'uploads')  // Render에서는 /tmp 사용
    : path.join(__dirname, '../uploads');
};

// 🔄 주문서를 표준 발주서로 변환
async function convertToStandardFormat(sourceFilePath, templateFilePath, mappingRules) {
  try {
    console.log('🔄 데이터 변환 시작');
    console.log('📂 입력 파일:', sourceFilePath);
    console.log('📂 템플릿 파일:', templateFilePath);
    
    const outputDir = getOutputDir();
    
    // 출력 디렉토리 확인 및 생성
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
      console.log('📁 출력 디렉토리 생성됨:', outputDir);
    }
    
    // 1. 원본 주문서 데이터 읽기
    const sourceData = await readSourceFile(sourceFilePath);
    
    // 2. 매핑 규칙 적용하여 데이터 변환
    const transformedData = applyMappingRules(sourceData, mappingRules);
    
    // 3. 발주서 템플릿에 데이터 삽입
    const result = await generatePurchaseOrder(templateFilePath, transformedData);
    
    return result;
    
  } catch (error) {
    console.error('변환 처리 오류:', error);
    throw new Error(`파일 변환 중 오류가 발생했습니다: ${error.message}`);
  }
}

// 📖 원본 파일 읽기 (Excel 또는 CSV)
async function readSourceFile(filePath) {
  const extension = path.extname(filePath).toLowerCase();
  
  if (extension === '.csv') {
    return await readCSVFile(filePath);
  } else {
    return await readExcelFile(filePath);
  }
}

// 📊 Excel 파일 읽기 (개선된 버전 - 복잡한 구조 지원)
async function readExcelFile(filePath) {
  console.log('📊 Excel 파일 읽기 시작:', {
    path: filePath,
    timestamp: new Date().toISOString()
  });

  const workbook = new ExcelJS.Workbook();
  
  try {
    // 파일 존재 확인
    if (!fs.existsSync(filePath)) {
      throw new Error(`파일을 찾을 수 없습니다: ${filePath}`);
    }
    
    // 파일 크기 확인
    const stats = fs.statSync(filePath);
    const fileSizeMB = stats.size / 1024 / 1024;
    console.log('📊 파일 정보:', {
      size: stats.size,
      sizeInMB: fileSizeMB.toFixed(2) + 'MB'
    });
    
    // Render 환경에서 대용량 파일 경고
    if (process.env.NODE_ENV === 'production' && fileSizeMB > 10) {
      console.warn('⚠️ 대용량 파일 감지: 처리 시간이 오래 걸릴 수 있습니다.');
    }
    
    // 메모리 효율적인 옵션으로 파일 읽기
    await workbook.xlsx.readFile(filePath, {
      sharedStrings: 'cache',
      hyperlinks: 'ignore',
      worksheets: 'emit',
      styles: 'cache'
    });
    
    console.log('📊 총 워크시트 개수:', workbook.worksheets.length);
    
  } catch (readError) {
    console.error('❌ Excel 파일 읽기 실패:', readError.message);
    throw new Error(`Excel 파일을 읽을 수 없습니다: ${readError.message}`);
  }
  
  // 1. 가장 적합한 워크시트 찾기
  let bestWorksheet = null;
  let bestScore = 0;
  
  try {
    workbook.worksheets.forEach((worksheet, index) => {
      try {
        console.log(`📄 워크시트 ${index + 1} 분석: ${worksheet.name} (행:${worksheet.rowCount}, 열:${worksheet.columnCount})`);
        
        // 데이터가 없거나 너무 적은 워크시트 제외
        if (worksheet.rowCount < 2 || worksheet.columnCount === 0) {
          console.log(`❌ 워크시트 ${index + 1} 제외: 데이터 부족`);
          return;
        }
        
        // 워크시트 점수 계산
        let score = 0;
        
        // 이름으로 점수 추가
        const sheetName = worksheet.name.toLowerCase();
        if (sheetName.includes('sheet') || sheetName.includes('데이터') || sheetName.includes('주문')) {
          score += 10;
        }
        if (sheetName.includes('요약') || sheetName.includes('피벗')) {
          score -= 20; // 요약/피벗 테이블은 피함
        }
        
        // 데이터 양으로 점수 추가
        score += Math.min(worksheet.rowCount / 10, 20); // 최대 20점
        score += Math.min(worksheet.columnCount, 10); // 최대 10점
        
        console.log(`📊 워크시트 ${index + 1} 점수: ${score}`);
        
        if (score > bestScore) {
          bestScore = score;
          bestWorksheet = worksheet;
        }
      } catch (sheetError) {
        console.warn(`⚠️ 워크시트 ${index + 1} 분석 중 오류 (건너뜀):`, sheetError.message);
      }
    });
  } catch (worksheetError) {
    console.error('❌ 워크시트 분석 중 오류:', worksheetError.message);
    // 첫 번째 워크시트를 기본으로 사용
    bestWorksheet = workbook.getWorksheet(1);
    console.log('🔄 첫 번째 워크시트를 기본으로 사용');
  }
  
  if (!bestWorksheet) {
    throw new Error('적절한 워크시트를 찾을 수 없습니다.');
  }
  
  console.log(`✅ 선택된 워크시트: ${bestWorksheet.name}`);
  
  // 2. 헤더 행 찾기
  let headerRowNum = 1;
  let headers = [];
  let maxHeaderScore = 0;
  
  const maxRowsToCheck = Math.min(10, bestWorksheet.rowCount);
  console.log(`🔍 헤더 검색 범위: 1-${maxRowsToCheck}행`);
  
  for (let rowNumber = 1; rowNumber <= maxRowsToCheck; rowNumber++) {
    try {
      const row = bestWorksheet.getRow(rowNumber);
      const potentialHeaders = [];
      let headerScore = 0;
      
      // 현재 행의 셀들을 확인 (최대 50개 컬럼까지 확장)
      const maxColumnsToCheck = Math.min(50, bestWorksheet.columnCount);
      for (let colNumber = 1; colNumber <= maxColumnsToCheck; colNumber++) {
        try {
          const cell = row.getCell(colNumber);
          const value = cell.value ? cell.value.toString().trim() : '';
          potentialHeaders.push(value);
          
          // 헤더 키워드로 점수 계산
          if (value) {
            if (value.includes('상품') || value.includes('제품') || value.includes('품목')) headerScore += 10;
            if (value.includes('수량') || value.includes('qty')) headerScore += 10;
            if (value.includes('가격') || value.includes('단가') || value.includes('price')) headerScore += 10;
            if (value.includes('고객') || value.includes('주문자') || value.includes('이름') || value.includes('성')) headerScore += 8;
            if (value.includes('연락') || value.includes('전화') || value.includes('휴대폰')) headerScore += 8;
            if (value.includes('주소') || value.includes('배송')) headerScore += 8;
            if (value.includes('이메일') || value.includes('email')) headerScore += 5;
            if (value.length > 0) headerScore += 1; // 빈 값이 아니면 1점
          }
        } catch (cellError) {
          console.warn(`⚠️ 셀 읽기 오류 (${rowNumber}, ${colNumber}): ${cellError.message}`);
          potentialHeaders.push('');
        }
      }
      
      console.log(`행 ${rowNumber} 헤더 점수: ${headerScore}, 샘플: [${potentialHeaders.slice(0, 5).join(', ')}...]`);
      
      if (headerScore > maxHeaderScore && headerScore > 5) { // 최소 점수 조건
        maxHeaderScore = headerScore;
        headerRowNum = rowNumber;
        headers = potentialHeaders.filter(h => h !== ''); // 빈 값 제거
      }
    } catch (rowError) {
      console.warn(`⚠️ 행 ${rowNumber} 처리 중 오류 (건너뜀):`, rowError.message);
    }
  }
  
  if (headers.length === 0) {
    // 헤더를 찾지 못한 경우 기본 컬럼명 생성
    console.log('⚠️ 헤더를 찾지 못함, 기본 컬럼명 사용');
    const firstDataRow = bestWorksheet.getRow(1);
    for (let colNumber = 1; colNumber <= bestWorksheet.columnCount; colNumber++) {
      headers.push(`컬럼${colNumber}`);
    }
    headerRowNum = 0; // 데이터가 1행부터 시작
  }
  
  console.log(`✅ 헤더 행: ${headerRowNum}, 헤더 개수: ${headers.length}`);
  console.log(`📋 발견된 헤더: [${headers.slice(0, 8).join(', ')}...]`);
  
  // AA 컬럼 (27번째) 확인
  if (headers.length >= 27) {
    console.log(`🏠 AA 컬럼 (27번째) 헤더: "${headers[26]}"`);
  } else {
    console.log(`❌ AA 컬럼 (27번째)을 찾을 수 없음 - 총 헤더 개수: ${headers.length}`);
  }
  
  // 3. 데이터 읽기
  const data = [];
  const dataStartRow = headerRowNum + 1;
  const maxRowsToProcess = bestWorksheet.rowCount; // 모든 행 처리하도록 변경
  
  console.log(`📋 데이터 읽기 시작: ${dataStartRow}행부터 ${maxRowsToProcess}행까지 (총 ${bestWorksheet.rowCount}행)`);
  
  let processedRows = 0;
  let skippedRows = 0;
  
  for (let rowNumber = dataStartRow; rowNumber <= maxRowsToProcess; rowNumber++) {
    try {
      const row = bestWorksheet.getRow(rowNumber);
      const rowData = {};
      
      headers.forEach((header, index) => {
        try {
          const cell = row.getCell(index + 1);
          const value = cell.value ? cell.value.toString().trim() : '';
          rowData[header] = value;
        } catch (cellError) {
          console.warn(`⚠️ 셀 읽기 오류 (${rowNumber}, ${index + 1}): ${cellError.message}`);
          rowData[header] = '';
        }
      });
      
      // 빈 행 제외 (모든 값이 빈 문자열인 경우)
      if (Object.values(rowData).some(value => value !== '')) {
        data.push(rowData);
        processedRows++;
        
        // 첫 5개 데이터 행에서 AA 컬럼 값 확인
        if (processedRows <= 5 && headers.length >= 27) {
          const aaColumnValue = rowData[headers[26]];
          console.log(`🏠 행 ${rowNumber} AA 컬럼 데이터: "${aaColumnValue}"`);
        }
      } else {
        skippedRows++;
      }
      
      // 진행 상황 로그 (500행마다)
      if (rowNumber % 500 === 0) {
        console.log(`📊 진행 상황: ${rowNumber}/${maxRowsToProcess}행 처리됨`);
      }
      
    } catch (rowError) {
      console.warn(`⚠️ 행 ${rowNumber} 처리 중 오류 (건너뜀):`, rowError.message);
      skippedRows++;
    }
  }
  
  console.log(`✅ 데이터 읽기 완료:`, {
    processedRows: processedRows,
    skippedRows: skippedRows,
    totalDataRows: data.length,
    processingTime: new Date().toISOString()
  });
  
  // 전체 헤더 목록 출력 (AA 컬럼 확인용)
  console.log('📋 전체 헤더 목록:');
  headers.forEach((header, index) => {
    if (index === 26) { // AA 컬럼
      console.log(`  [${index + 1}] (AA 컬럼): "${header}"`);
    } else if (index < 30) { // 처음 30개만 출력
      console.log(`  [${index + 1}]: "${header}"`);
    }
  });
  
  return { headers, data };
}

// 📄 CSV 파일 읽기
async function readCSVFile(filePath) {
  const csvData = fs.readFileSync(filePath, 'utf8');
  const lines = csvData.split('\n').filter(line => line.trim());
  
  if (lines.length === 0) {
    throw new Error('CSV 파일이 비어있습니다.');
  }
  
  const headers = lines[0].split(',').map(h => h.trim());
  const data = [];
  
  for (let i = 1; i < lines.length; i++) {
    const values = lines[i].split(',').map(v => v.trim());
    const rowData = {};
    
    headers.forEach((header, index) => {
      rowData[header] = values[index] || '';
    });
    
    if (Object.values(rowData).some(value => value !== '')) {
      data.push(rowData);
    }
  }
  
  return { headers, data };
}

// 🗺️ 매핑 규칙 적용
function applyMappingRules(sourceData, mappingRules) {
  const { headers, data } = sourceData;
  const { rules, fixedValues } = mappingRules;
  
  if (!rules || Object.keys(rules).length === 0) {
    // 기본 매핑 적용
    return applyDefaultMapping(data);
  }
  
  return data.map(row => {
    const transformedRow = {};
    
    // 매핑 규칙에 따라 데이터 변환
    Object.keys(rules).forEach(targetField => {
      const sourceField = rules[targetField];
      
      // 고정값 패턴 확인 ([고정값: xxx] 형태)
      if (sourceField && sourceField.startsWith('[고정값:') && sourceField.endsWith(']')) {
        // 고정값에서 실제 값 추출
        const fixedValue = sourceField.substring(6, sourceField.length - 1); // '[고정값:' 제거하고 ']' 제거
        transformedRow[targetField] = fixedValue.trim();
      } else if (sourceField && row[sourceField] !== undefined) {
        // 일반 필드 매핑
        transformedRow[targetField] = row[sourceField];
      }
    });
    
    // 고정값이 별도로 전달된 경우 적용
    if (fixedValues && Object.keys(fixedValues).length > 0) {
      Object.keys(fixedValues).forEach(field => {
        transformedRow[field] = fixedValues[field];
      });
    }
    
    // 계산 필드 추가
    if (transformedRow.수량 && transformedRow.단가) {
      transformedRow.금액 = parseInt(transformedRow.수량) * parseFloat(transformedRow.단가);
    }
    
    return transformedRow;
  });
}

// 🔧 기본 매핑 적용 (매핑 규칙이 없는 경우)
function applyDefaultMapping(data) {
  const defaultMappings = {
    '상품명': ['상품명', '품목명', '제품명', 'product'],
    '수량': ['수량', '주문수량', 'quantity', 'qty'],
    '단가': ['단가', '가격', 'price', 'unit_price'],
    '고객명': ['고객명', '주문자', '배송받는분', 'customer'],
    '연락처': ['연락처', '전화번호', 'phone', 'tel'],
    '주소': ['주소', '배송지', 'address']
  };
  
  return data.map(row => {
    const transformedRow = {};
    
    Object.keys(defaultMappings).forEach(targetField => {
      const possibleFields = defaultMappings[targetField];
      
      for (const field of possibleFields) {
        if (row[field] !== undefined) {
          transformedRow[targetField] = row[field];
          break;
        }
      }
    });
    
    // 계산 필드 추가
    if (transformedRow.수량 && transformedRow.단가) {
      transformedRow.금액 = parseInt(transformedRow.수량) * parseFloat(transformedRow.단가);
    }
    
    return transformedRow;
  });
}

// 📋 발주서 생성
async function generatePurchaseOrder(templateFilePath, transformedData) {
  const outputDir = getOutputDir();
  const workbook = new ExcelJS.Workbook();
  let useTemplate = false;
  
  // 템플릿 사용을 시도하되, 오류 발생 시 새 워크북 생성
  try {
    if (fs.existsSync(templateFilePath)) {
      // 템플릿 파일이 있는 경우 - 공유 수식 오류 방지를 위한 옵션 추가
      await workbook.xlsx.readFile(templateFilePath, {
        sharedStrings: 'cache',
        hyperlinks: 'ignore',
        worksheets: 'emit',
        styles: 'cache'
      });
      useTemplate = true;
      console.log('템플릿 파일 로드 성공');
    }
  } catch (templateError) {
    console.log('템플릿 파일 사용 불가, 새 워크북 생성:', templateError.message);
    useTemplate = false;
  }
  
  // 템플릿 사용에 실패했거나 없는 경우 새 워크북 생성
  if (!useTemplate) {
    // 기존 워크북 초기화
    workbook.removeWorksheet(workbook.getWorksheet(1));
    workbook.addWorksheet('발주서');
  }
  
  const worksheet = workbook.getWorksheet(1) || workbook.addWorksheet('발주서');
  
  // 템플릿에 데이터 삽입
  const dataStartRow = findDataStartRow(worksheet) || 3;
  
  // 헤더 설정 (데이터 시작 행 바로 위)
  const headerRow = worksheet.getRow(dataStartRow - 1);
  const standardHeaders = ['발주번호', '발주일자', '품목명', '주문수량', '단가', '공급가액', '받는 분', '전화번호', '주소'];
  
  standardHeaders.forEach((header, index) => {
    headerRow.getCell(index + 1).value = header;
    headerRow.getCell(index + 1).font = { bold: true };
  });
  
  // 데이터 삽입
  const errors = [];
  const processedRows = [];
  
  transformedData.forEach((row, index) => {
    try {
      const dataRow = worksheet.getRow(dataStartRow + index);
      
      // 발주번호 생성 (ORD + 날짜 + 순번)
      const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
      const orderNumber = `ORD${today}-${String(index + 1).padStart(3, '0')}`;
      
      dataRow.getCell(1).value = orderNumber; // 발주번호
      dataRow.getCell(2).value = new Date(); // 발주일자
      dataRow.getCell(3).value = row.상품명 || ''; // 품목명
      dataRow.getCell(4).value = row.수량 ? parseInt(row.수량) : ''; // 주문수량
      dataRow.getCell(5).value = row.단가 ? parseFloat(row.단가) : ''; // 단가
      dataRow.getCell(6).value = row.금액 ? parseFloat(row.금액) : ''; // 공급가액
      dataRow.getCell(7).value = row.고객명 || ''; // 받는 분
      dataRow.getCell(8).value = row.연락처 || ''; // 전화번호
      dataRow.getCell(9).value = row.주소 || ''; // 주소
      
      processedRows.push(row);
      
    } catch (error) {
      errors.push({
        row: index + 1,
        error: error.message,
        data: row
      });
    }
  });
  
  // 합계 행 추가 - 수식 대신 계산된 값 사용
  if (processedRows.length > 0) {
    const totalRow = worksheet.getRow(dataStartRow + transformedData.length);
    totalRow.getCell(3).value = '합계'; // 품목명 위치에 합계 표시
    
    // 수식 대신 직접 계산한 값 사용
    const totalQuantity = processedRows.reduce((sum, row) => sum + (parseInt(row.수량) || 0), 0);
    const totalAmount = processedRows.reduce((sum, row) => sum + (parseFloat(row.금액) || 0), 0);
    
    totalRow.getCell(4).value = totalQuantity; // 주문수량
    totalRow.getCell(6).value = totalAmount; // 공급가액
    totalRow.font = { bold: true };
  }
  
  // 파일 저장 - 공유 수식 오류 방지
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
      const fileName = `purchase_order_${timestamp}.xlsx`;
  const outputPath = path.join(outputDir, fileName);
  
  // 안전한 파일 저장
  try {
    // 템플릿을 사용했다면 수식 문제를 해결
    if (useTemplate) {
      try {
        worksheet.eachRow((row, rowNumber) => {
          row.eachCell((cell, colNumber) => {
            try {
              // 수식이 있는지 안전하게 확인
              if (cell && typeof cell === 'object' && cell.type === 'formula') {
                // 수식을 값으로 변환
                const currentValue = cell.result || cell.value || 0;
                cell.type = 'number';
                cell.value = currentValue;
              }
            } catch (cellError) {
              // 개별 셀 오류는 무시하고 계속 진행
              console.log(`셀 처리 오류 (${rowNumber}, ${colNumber}):`, cellError.message);
            }
          });
        });
      } catch (worksheetError) {
        console.log('워크시트 수식 처리 중 오류, 단순 저장으로 변경:', worksheetError.message);
        // 수식 처리 실패 시 새 워크북으로 대체
        return await createSimpleWorkbook(transformedData, outputPath, fileName);
      }
    }
    
    await workbook.xlsx.writeFile(outputPath);
    
  } catch (writeError) {
    console.error('파일 저장 오류, 단순 워크북으로 재생성:', writeError.message);
    return await createSimpleWorkbook(transformedData, outputPath, fileName);
  }
  
  return {
    fileName,
    filePath: outputPath,
    processedRows: processedRows.length,
    totalRows: transformedData.length,
    errors
  };
}

// 🔍 템플릿에서 데이터 시작 행 찾기
function findDataStartRow(worksheet) {
  let dataStartRow = 3; // 기본값
  
  // 'NO' 또는 '번호' 헤더를 찾아서 데이터 시작 행 결정
  for (let rowNumber = 1; rowNumber <= 10; rowNumber++) {
    const row = worksheet.getRow(rowNumber);
    for (let colNumber = 1; colNumber <= 10; colNumber++) {
      const cell = row.getCell(colNumber);
      if (cell.value && ['NO', '번호', '순번'].includes(cell.value.toString().toUpperCase())) {
        return rowNumber + 1;
      }
    }
  }
  
  return dataStartRow;
}

// 📄 단순한 워크북 생성 (공유 수식 문제 회피)
async function createSimpleWorkbook(transformedData, outputPath, fileName) {
  const simpleWorkbook = new ExcelJS.Workbook();
  const simpleWorksheet = simpleWorkbook.addWorksheet('발주서');
  
  // 제목 설정
  simpleWorksheet.getCell('A1').value = '발주서';
  simpleWorksheet.getCell('A1').font = { size: 16, bold: true };
  simpleWorksheet.mergeCells('A1:H1');
  simpleWorksheet.getCell('A1').alignment = { horizontal: 'center' };
  
  // 헤더 설정
  const standardHeaders = ['발주번호', '발주일자', '품목명', '주문수량', '단가', '공급가액', '받는 분', '전화번호', '주소'];
  standardHeaders.forEach((header, index) => {
    const cell = simpleWorksheet.getCell(2, index + 1);
    cell.value = header;
    cell.font = { bold: true };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE0E0E0' } };
    cell.border = {
      top: { style: 'thin' },
      left: { style: 'thin' },
      bottom: { style: 'thin' },
      right: { style: 'thin' }
    };
  });
  
  // 데이터 입력
  const processedRows = [];
  const errors = [];
  
  transformedData.forEach((row, index) => {
    try {
      const dataRowNum = index + 3;
      
      // 발주번호 생성 (ORD + 날짜 + 순번)
      const today = new Date().toISOString().slice(0, 10).replace(/-/g, '');
      const orderNumber = `ORD${today}-${String(index + 1).padStart(3, '0')}`;
      
      simpleWorksheet.getCell(dataRowNum, 1).value = orderNumber; // 발주번호
      simpleWorksheet.getCell(dataRowNum, 2).value = new Date(); // 발주일자
      simpleWorksheet.getCell(dataRowNum, 3).value = row.상품명 || ''; // 품목명
      simpleWorksheet.getCell(dataRowNum, 4).value = row.수량 ? parseInt(row.수량) : ''; // 주문수량
      simpleWorksheet.getCell(dataRowNum, 5).value = row.단가 ? parseFloat(row.단가) : ''; // 단가
      simpleWorksheet.getCell(dataRowNum, 6).value = row.금액 ? parseFloat(row.금액) : ''; // 공급가액
      simpleWorksheet.getCell(dataRowNum, 7).value = row.고객명 || ''; // 받는 분
      simpleWorksheet.getCell(dataRowNum, 8).value = row.연락처 || ''; // 전화번호
      simpleWorksheet.getCell(dataRowNum, 9).value = row.주소 || ''; // 주소
      
      // 테두리 추가
      for (let col = 1; col <= 9; col++) {
        simpleWorksheet.getCell(dataRowNum, col).border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      }
      
      processedRows.push(row);
      
    } catch (error) {
      errors.push({
        row: index + 1,
        error: error.message,
        data: row
      });
    }
  });
  
  // 합계 행 추가
  if (processedRows.length > 0) {
    const totalRowNum = transformedData.length + 3;
    const totalQuantity = processedRows.reduce((sum, row) => sum + (parseInt(row.수량) || 0), 0);
    const totalAmount = processedRows.reduce((sum, row) => sum + (parseFloat(row.금액) || 0), 0);
    
    simpleWorksheet.getCell(totalRowNum, 3).value = '합계'; // 품목명 위치
    simpleWorksheet.getCell(totalRowNum, 4).value = totalQuantity; // 주문수량
    simpleWorksheet.getCell(totalRowNum, 6).value = totalAmount; // 공급가액
    
    for (let col = 1; col <= 9; col++) {
      const cell = simpleWorksheet.getCell(totalRowNum, col);
      cell.font = { bold: true };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF0F0F0' } };
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' }
      };
    }
  }
  
  // 열 너비 조정
  simpleWorksheet.columns = [
    { width: 15 },  // 발주번호
    { width: 12 },  // 발주일자
    { width: 20 },  // 품목명
    { width: 10 },  // 주문수량
    { width: 12 },  // 단가
    { width: 12 },  // 공급가액
    { width: 15 },  // 받는 분
    { width: 15 },  // 전화번호
    { width: 25 }   // 주소
  ];
  
  await simpleWorkbook.xlsx.writeFile(outputPath);
  
  return {
    fileName,
    filePath: outputPath,
    processedRows: processedRows.length,
    totalRows: transformedData.length,
    errors
  };
}

// 📝 직접 입력 데이터를 표준 발주서로 변환
async function convertDirectInputToStandardFormat(templateFilePath, inputData, mappingRules) {
  try {
    console.log('📝 직접 입력 데이터 변환 시작');
    console.log('📂 템플릿 파일:', templateFilePath);
    console.log('📝 입력 데이터:', inputData);
    
    const outputDir = getOutputDir();
    
    // 출력 디렉토리 확인 및 생성
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir, { recursive: true });
      console.log('📁 출력 디렉토리 생성됨:', outputDir);
    }
    
    // 직접 입력 데이터를 표준 형식으로 변환
    const transformedData = [inputData]; // 단일 행 데이터로 처리
    
    // 발주서 템플릿에 데이터 삽입
    const result = await generatePurchaseOrder(templateFilePath, transformedData);
    
    return result;
    
  } catch (error) {
    console.error('직접 입력 데이터 변환 오류:', error);
    throw new Error(`직접 입력 데이터 변환 중 오류가 발생했습니다: ${error.message}`);
  }
}

module.exports = {
  convertToStandardFormat,
  convertDirectInputToStandardFormat,
  readSourceFile,
  readExcelFile,
  applyMappingRules,
  generatePurchaseOrder
}; 