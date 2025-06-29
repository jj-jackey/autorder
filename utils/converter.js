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

// 📊 Excel 파일 읽기
async function readExcelFile(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  
  const worksheet = workbook.getWorksheet(1);
  const data = [];
  
  if (!worksheet) {
    throw new Error('워크시트를 찾을 수 없습니다.');
  }
  
  // 헤더 읽기
  const headerRow = worksheet.getRow(1);
  const headers = [];
  headerRow.eachCell((cell, colNumber) => {
    headers.push(cell.value ? cell.value.toString().trim() : `컬럼${colNumber}`);
  });
  
  // 데이터 읽기
  for (let rowNumber = 2; rowNumber <= worksheet.rowCount; rowNumber++) {
    const row = worksheet.getRow(rowNumber);
    const rowData = {};
    
    headers.forEach((header, index) => {
      const cell = row.getCell(index + 1);
      rowData[header] = cell.value ? cell.value.toString().trim() : '';
    });
    
    // 빈 행 제외
    if (Object.values(rowData).some(value => value !== '')) {
      data.push(rowData);
    }
  }
  
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
  const { rules } = mappingRules;
  
  if (!rules || Object.keys(rules).length === 0) {
    // 기본 매핑 적용
    return applyDefaultMapping(data);
  }
  
  return data.map(row => {
    const transformedRow = {};
    
    // 매핑 규칙에 따라 데이터 변환
    Object.keys(rules).forEach(targetField => {
      const sourceField = rules[targetField];
      if (sourceField && row[sourceField] !== undefined) {
        transformedRow[targetField] = row[sourceField];
      }
    });
    
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
  const standardHeaders = ['NO', '상품명', '수량', '단가', '금액', '고객명', '연락처', '주소'];
  
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
      
      dataRow.getCell(1).value = index + 1; // NO
      dataRow.getCell(2).value = row.상품명 || '';
      dataRow.getCell(3).value = row.수량 ? parseInt(row.수량) : '';
      dataRow.getCell(4).value = row.단가 ? parseFloat(row.단가) : '';
      dataRow.getCell(5).value = row.금액 ? parseFloat(row.금액) : '';
      dataRow.getCell(6).value = row.고객명 || '';
      dataRow.getCell(7).value = row.연락처 || '';
      dataRow.getCell(8).value = row.주소 || '';
      
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
    totalRow.getCell(2).value = '합계';
    
    // 수식 대신 직접 계산한 값 사용
    const totalQuantity = processedRows.reduce((sum, row) => sum + (parseInt(row.수량) || 0), 0);
    const totalAmount = processedRows.reduce((sum, row) => sum + (parseFloat(row.금액) || 0), 0);
    
    totalRow.getCell(3).value = totalQuantity;
    totalRow.getCell(5).value = totalAmount;
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
  const standardHeaders = ['NO', '상품명', '수량', '단가', '금액', '고객명', '연락처', '주소'];
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
      simpleWorksheet.getCell(dataRowNum, 1).value = index + 1;
      simpleWorksheet.getCell(dataRowNum, 2).value = row.상품명 || '';
      simpleWorksheet.getCell(dataRowNum, 3).value = row.수량 ? parseInt(row.수량) : '';
      simpleWorksheet.getCell(dataRowNum, 4).value = row.단가 ? parseFloat(row.단가) : '';
      simpleWorksheet.getCell(dataRowNum, 5).value = row.금액 ? parseFloat(row.금액) : '';
      simpleWorksheet.getCell(dataRowNum, 6).value = row.고객명 || '';
      simpleWorksheet.getCell(dataRowNum, 7).value = row.연락처 || '';
      simpleWorksheet.getCell(dataRowNum, 8).value = row.주소 || '';
      
      // 테두리 추가
      for (let col = 1; col <= 8; col++) {
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
    
    simpleWorksheet.getCell(totalRowNum, 2).value = '합계';
    simpleWorksheet.getCell(totalRowNum, 3).value = totalQuantity;
    simpleWorksheet.getCell(totalRowNum, 5).value = totalAmount;
    
    for (let col = 1; col <= 8; col++) {
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
    { width: 5 },   // NO
    { width: 20 },  // 상품명
    { width: 8 },   // 수량
    { width: 12 },  // 단가
    { width: 12 },  // 금액
    { width: 15 },  // 고객명
    { width: 15 },  // 연락처
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

module.exports = {
  convertToStandardFormat,
  readSourceFile,
  applyMappingRules,
  generatePurchaseOrder
}; 