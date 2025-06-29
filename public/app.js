// 전역 변수
let currentFileId = null;
let currentMapping = {};
let generatedFileName = null;

// 페이지 로드 시 초기화
document.addEventListener('DOMContentLoaded', function() {
    initializeApp();
    loadEmailHistory();
    updateDashboard();
    
    // 초기 상태 설정
    currentMapping = {};
    generatedFileName = null;
    resetAllSteps();
    
    // 매핑 상태 초기화
    sessionStorage.setItem('mappingSaved', 'false');
    
    // GENERATE ORDER 버튼 초기 비활성화
    setTimeout(() => {
        updateGenerateOrderButton();
    }, 100);
    
    // 진행률 초기 숨김
    hideProgress();
});

// 앱 초기화
function initializeApp() {
    const uploadArea = document.getElementById('uploadArea');
    const fileInput = document.getElementById('fileInput');
    
    // 드래그 앤 드롭 이벤트
    uploadArea.addEventListener('click', () => fileInput.click());
    uploadArea.addEventListener('dragover', handleDragOver);
    uploadArea.addEventListener('dragleave', handleDragLeave);
    uploadArea.addEventListener('drop', handleDrop);
    
    // 파일 선택 이벤트
    fileInput.addEventListener('change', handleFileSelect);
    
    // 전송 옵션 변경 이벤트
    document.querySelectorAll('input[name="sendOption"]').forEach(radio => {
        radio.addEventListener('change', function() {
            const scheduleTimeGroup = document.getElementById('scheduleTimeGroup');
            scheduleTimeGroup.style.display = this.value === 'scheduled' ? 'flex' : 'none';
        });
    });
}



// 드래그 오버 처리
function handleDragOver(e) {
    e.preventDefault();
    e.currentTarget.classList.add('drag-over');
}

// 드래그 떠남 처리
function handleDragLeave(e) {
    e.currentTarget.classList.remove('drag-over');
}

// 드롭 처리
function handleDrop(e) {
    e.preventDefault();
    e.currentTarget.classList.remove('drag-over');
    
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        processFile(files[0]);
    }
}

// 파일 선택 처리
function handleFileSelect(e) {
    const file = e.target.files[0];
    if (file) {
        processFile(file);
    }
}

// 파일 처리
async function processFile(file) {
    // 파일 형식 검증
    const allowedTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 
                         'application/vnd.ms-excel', 'text/csv'];
    const allowedExtensions = ['.xlsx', '.xls', '.csv'];
    
    if (!allowedTypes.includes(file.type) && !allowedExtensions.some(ext => file.name.toLowerCase().endsWith(ext))) {
        showAlert('error', '지원하지 않는 파일 형식입니다. Excel(.xlsx, .xls) 또는 CSV 파일을 업로드해주세요.');
        return;
    }
    
    // 파일 크기 검증 (10MB)
    if (file.size > 10 * 1024 * 1024) {
        showAlert('error', '파일 크기가 너무 큽니다. 10MB 이하의 파일을 업로드해주세요.');
        return;
    }
    
    try {
        showLoading('파일을 업로드하고 있습니다...');
        
        const formData = new FormData();
        formData.append('orderFile', file);
        
        const response = await fetch('/api/orders/upload', {
            method: 'POST',
            body: formData
        });
        
        const result = await response.json();
        
        hideLoading();
        
        if (result.success) {
            // 새 파일 업로드 시 초기화
            resetAllSteps();
            currentFileId = result.fileId;
            currentMapping = {}; // 매핑 초기화
            generatedFileName = null; // 생성된 파일명 초기화
            
            showUploadResult(result);
            showStep(2);
            setupMapping(result.headers);
        } else {
            showAlert('error', result.error || '파일 업로드에 실패했습니다.');
        }
        
    } catch (error) {
        hideLoading();
        console.error('업로드 오류:', error);
        showAlert('error', '파일 업로드 중 오류가 발생했습니다.');
    }
}

// 업로드 결과 표시
function showUploadResult(result) {
    const uploadResult = document.getElementById('uploadResult');
    const uploadAlert = document.getElementById('uploadAlert');
    const previewContainer = document.getElementById('previewContainer');
    
    uploadResult.classList.remove('hidden');
    
    // 검증 결과에 따른 알림 표시
    if (result.validation.isValid) {
        uploadAlert.innerHTML = `
            <div class="alert alert-success">
                ✅ ${result.message}<br>
                <strong>검증 결과:</strong> ${result.validation.validRows}/${result.validation.totalRows}행 처리 가능 
                (성공률: ${result.validation.summary.successRate}%)
            </div>
        `;
    } else {
        uploadAlert.innerHTML = `
            <div class="alert alert-warning">
                ⚠️ ${result.message}<br>
                <strong>오류:</strong> ${result.validation.errorRows}개 행에서 오류 발견<br>
                <strong>경고:</strong> ${result.validation.warningRows}개 행에서 경고 발견
            </div>
        `;
    }
    
    // 미리보기 테이블 생성
    if (result.previewData && result.previewData.length > 0) {
        let tableHtml = '<h5>DATA PREVIEW (상위 20행)</h5>';
        tableHtml += '<table class="preview-table"><thead><tr>';
        
        result.headers.forEach(header => {
            tableHtml += `<th>${header}</th>`;
        });
        tableHtml += '</tr></thead><tbody>';
        
        result.previewData.slice(0, 10).forEach(row => {
            tableHtml += '<tr>';
            result.headers.forEach(header => {
                tableHtml += `<td>${row[header] || ''}</td>`;
            });
            tableHtml += '</tr>';
        });
        
        tableHtml += '</tbody></table>';
        previewContainer.innerHTML = tableHtml;
    }
}

// 매핑 설정
function setupMapping(sourceHeaders) {
    // 소스 필드 초기화
    const sourceFieldsContainer = document.getElementById('sourceFields');
    sourceFieldsContainer.innerHTML = '';
    
    sourceHeaders.forEach(header => {
        const fieldDiv = document.createElement('div');
        fieldDiv.className = 'field-item';
        fieldDiv.textContent = header;
        fieldDiv.dataset.source = header;
        fieldDiv.onclick = () => selectSourceField(fieldDiv);
        sourceFieldsContainer.appendChild(fieldDiv);
    });
    
    // 표준 타겟 필드 설정
    setupStandardTargetFields();
    
    // 타겟 필드 초기화 (이전 매핑 상태 제거)
    resetTargetFields();
    
    // 타겟 필드 클릭 이벤트
    document.querySelectorAll('#targetFields .field-item').forEach(item => {
        item.onclick = () => selectTargetField(item);
    });
    
    // 매핑 상태 초기화
    sessionStorage.setItem('mappingSaved', 'false');
    
    // GENERATE ORDER 버튼 초기 비활성화
    updateGenerateOrderButton();
}

// 소스 필드 선택
function selectSourceField(element) {
    document.querySelectorAll('#sourceFields .field-item').forEach(item => {
        item.classList.remove('selected');
    });
    element.classList.add('selected');
}

// 타겟 필드 선택 및 매핑
function selectTargetField(element) {
    const targetField = element.dataset.target;
    
    // 이미 매핑된 필드인지 확인 (매핑 취소 기능)
    if (currentMapping[targetField]) {
        // 매핑 취소
        const sourceField = currentMapping[targetField];
        delete currentMapping[targetField];
        
        // 타겟 필드 원래대로 복원
        element.style.background = '';
        element.style.color = '';
        element.innerHTML = targetField;
        
        // 소스 필드를 다시 SOURCE FIELDS에 추가
        const sourceFieldsContainer = document.getElementById('sourceFields');
        const fieldDiv = document.createElement('div');
        fieldDiv.className = 'field-item';
        fieldDiv.textContent = sourceField;
        fieldDiv.dataset.source = sourceField;
        fieldDiv.onclick = () => selectSourceField(fieldDiv);
        sourceFieldsContainer.appendChild(fieldDiv);
        
        showAlert('info', `${sourceField} → ${targetField} 매핑이 취소되었습니다.`);
        
        // GENERATE ORDER 버튼 비활성화
        updateGenerateOrderButton();
        return;
    }
    
    // 새로운 매핑 생성
    const selectedSource = document.querySelector('#sourceFields .field-item.selected');
    
    if (!selectedSource) {
        showAlert('warning', '먼저 주문서 컬럼을 선택해주세요.');
        return;
    }
    
    const sourceField = selectedSource.dataset.source;
    
    // 매핑 저장
    currentMapping[targetField] = sourceField;
    
    // 시각적 표시
    element.style.background = '#28a745';
    element.style.color = 'white';
    element.innerHTML = `${targetField} ← ${sourceField}`;
    
    // 선택된 소스 필드 제거
    selectedSource.remove();
    
    showAlert('success', `${sourceField} → ${targetField} 매핑이 완료되었습니다.`);
    
    // GENERATE ORDER 버튼 상태 업데이트
    updateGenerateOrderButton();
}

// GENERATE ORDER 버튼 상태 업데이트
function updateGenerateOrderButton() {
    const generateBtn = document.querySelector('button[onclick="generateOrder()"]');
    const isMappingSaved = sessionStorage.getItem('mappingSaved') === 'true';
    
    if (isMappingSaved && Object.keys(currentMapping).length > 0) {
        generateBtn.disabled = false;
        generateBtn.style.opacity = '1';
        generateBtn.style.cursor = 'pointer';
    } else {
        generateBtn.disabled = true;
        generateBtn.style.opacity = '0.5';
        generateBtn.style.cursor = 'not-allowed';
    }
}

// 매핑 저장
async function saveMapping() {
    if (Object.keys(currentMapping).length === 0) {
        showAlert('warning', '매핑 규칙을 설정해주세요.');
        return;
    }
    
    // 필수 필드 검증
    const validation = validateRequiredFields(currentMapping);
    if (!validation.isValid) {
        // 필수 필드가 누락되었을 때 입력 폼 표시
        showMissingFieldsForm(validation.missingFields);
        return;
    }
    
    try {
        const mappingData = {
            mappingName: `mapping_${Date.now()}`,
            sourceFields: Object.values(currentMapping),
            targetFields: Object.keys(currentMapping),
            mappingRules: currentMapping
        };
        
        const response = await fetch('/api/orders/mapping', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(mappingData)
        });
        
        const result = await response.json();
        
        if (result.success) {
            showAlert('success', '✅ 매핑 규칙이 저장되었습니다. 모든 필수 필드가 올바르게 매핑되었습니다.');
            
            // 매핑 저장 상태 표시
            sessionStorage.setItem('mappingSaved', 'true');
            
            // GENERATE ORDER 버튼 활성화
            updateGenerateOrderButton();
            
        } else {
            showAlert('error', result.error || '매핑 저장에 실패했습니다.');
        }
        
    } catch (error) {
        console.error('매핑 저장 오류:', error);
        showAlert('error', '매핑 저장 중 오류가 발생했습니다.');
    }
}

// 발주서 생성
async function generateOrder() {
    if (!currentFileId) {
        showAlert('error', '업로드된 파일이 없습니다.');
        return;
    }
    
    try {
        // 진행률 표시 시작
        showProgress('발주서 생성을 준비하고 있습니다...');
        
        // 진행률 단계 정의
        const progressSteps = [
            { percent: 10, message: '매핑 규칙을 저장하고 있습니다...' },
            { percent: 30, message: '파일 데이터를 읽고 있습니다...' },
            { percent: 50, message: '데이터를 변환하고 있습니다...' },
            { percent: 75, message: '발주서를 생성하고 있습니다...' },
            { percent: 90, message: '최종 검증을 진행하고 있습니다...' },
            { percent: 100, message: '발주서 생성이 완료되었습니다!' }
        ];
        
        const requestData = {
            fileId: currentFileId,
            mappingId: `mapping_${Date.now()}`,
            templateType: 'standard'
        };
        
        // 진행률 시뮬레이션과 실제 작업을 병렬로 실행
        const progressPromise = simulateProgress(progressSteps, 2500);
        
        // 실제 API 호출
        const workPromise = (async () => {
            // 매핑 저장
            await fetch('/api/orders/mapping', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    mappingName: requestData.mappingId,
                    mappingRules: currentMapping
                })
            });
            
            // 발주서 생성
            const response = await fetch('/api/orders/generate', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(requestData)
            });
            
            return response.json();
        })();
        
        // 진행률과 실제 작업 모두 완료될 때까지 대기
        const [_, result] = await Promise.all([progressPromise, workPromise]);
        
        // 진행률 숨기기
        hideProgress();
        
        if (result.success) {
            generatedFileName = result.generatedFile;
            showGenerateResult(result);
            showStep(3);
            showStep(4);
        } else {
            showAlert('error', result.error || '발주서 생성에 실패했습니다.');
        }
        
    } catch (error) {
        hideProgress();
        console.error('발주서 생성 오류:', error);
        showAlert('error', '발주서 생성 중 오류가 발생했습니다.');
    }
}

// 발주서 생성 결과 표시
function showGenerateResult(result) {
    const generateResult = document.getElementById('generateResult');
    
    generateResult.innerHTML = `
        <div class="alert alert-success">
            ✅ 발주서가 성공적으로 생성되었습니다!<br>
            <strong>처리 결과:</strong> ${result.processedRows}/${result.processedRows}행 처리 완료<br>
            <strong>생성된 파일:</strong> ${result.generatedFile}
        </div>
        
        <div style="text-align: center; margin-top: 20px;">
            <a href="${result.downloadUrl}" class="btn btn-success" download>DOWNLOAD ORDER</a>
        </div>
    `;
    
    if (result.errors && result.errors.length > 0) {
        generateResult.innerHTML += `
            <div class="alert alert-warning" style="margin-top: 15px;">
                <strong>오류 내역:</strong><br>
                ${result.errors.map(err => `행 ${err.row}: ${err.error}`).join('<br>')}
            </div>
        `;
    }
}

// 이메일 전송
async function sendEmail() {
    const emailTo = document.getElementById('emailTo').value;
    const emailSubject = document.getElementById('emailSubject').value;
    const emailBody = document.getElementById('emailBody').value;
    const sendOption = document.querySelector('input[name="sendOption"]:checked').value;
    const scheduleTime = document.getElementById('scheduleTime').value;
    
    if (!emailTo || !emailSubject || !generatedFileName) {
        showAlert('error', '필수 항목을 모두 입력해주세요.');
        return;
    }
    
    try {
        showLoading('이메일을 전송하고 있습니다...');
        
        const emailData = {
            to: emailTo,
            subject: emailSubject,
            body: emailBody,
            attachmentPath: generatedFileName
        };
        
        if (sendOption === 'scheduled' && scheduleTime) {
            emailData.scheduleTime = scheduleTime;
        }
        
        const response = await fetch('/api/email/send', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(emailData)
        });
        
        const result = await response.json();
        
        hideLoading();
        
        if (result.success) {
            showEmailResult('success', result.message);
            loadEmailHistory();
            updateDashboard();
        } else {
            showEmailResult('error', result.error || '이메일 전송에 실패했습니다.');
        }
        
    } catch (error) {
        hideLoading();
        console.error('이메일 전송 오류:', error);
        showEmailResult('error', '이메일 전송 중 오류가 발생했습니다.');
    }
}

// 이메일 전송 결과 표시
function showEmailResult(type, message) {
    const emailResult = document.getElementById('emailResult');
    const alertClass = type === 'success' ? 'alert-success' : 'alert-error';
    const icon = type === 'success' ? '●' : '●';
    
    emailResult.innerHTML = `
        <div class="alert ${alertClass}" style="margin-top: 20px;">
            <span style="color: ${type === 'success' ? '#28a745' : '#dc3545'}">${icon}</span> ${message}
        </div>
    `;
}

// 이메일 이력 로드
async function loadEmailHistory() {
    try {
        const response = await fetch('/api/email/history');
        const result = await response.json();
        
        if (result.success && result.history.length > 0) {
            const historyList = document.getElementById('emailHistoryList');
            
            historyList.innerHTML = result.history.slice(0, 10).map((item, index) => {
                const statusClass = item.status === 'success' ? '' : 'failed';
                const statusIcon = item.status === 'success' ? '●' : '●';
                
                return `
                    <div class="history-item ${statusClass}" style="display: flex; align-items: center; justify-content: space-between;">
                        <div style="display: flex; align-items: center; flex: 1;">
                            <input type="checkbox" class="history-checkbox" data-index="${index}" onchange="updateDeleteButton()" style="margin-right: 10px;">
                            <div style="flex: 1;">
                                <div><strong><span style="color: ${item.status === 'success' ? '#28a745' : '#dc3545'}">${statusIcon}</span> ${item.to}</strong></div>
                                <div>${item.subject}</div>
                                <div class="history-time">${new Date(item.sentAt).toLocaleString()}</div>
                                ${item.error ? `<div style="color: #dc3545; font-size: 0.9em;">ERROR: ${item.error}</div>` : ''}
                            </div>
                        </div>
                        <button class="btn" onclick="deleteSingleHistory(${index})" style="background: linear-gradient(135deg, #dc3545 0%, #c82333 100%); margin-left: 10px; padding: 5px 10px; font-size: 0.8em;">삭제</button>
                    </div>
                `;
            }).join('');
        } else {
            const historyList = document.getElementById('emailHistoryList');
            historyList.innerHTML = '<p style="text-align: center; color: #6c757d;">전송 이력이 없습니다.</p>';
        }
        
        // 전체 선택 체크박스 초기화
        document.getElementById('selectAllHistory').checked = false;
        updateDeleteButton();
        
    } catch (error) {
        console.error('이력 로드 오류:', error);
    }
}

// 대시보드 업데이트
async function updateDashboard() {
    try {
        const response = await fetch('/api/email/history');
        const result = await response.json();
        
        if (result.success) {
            const today = new Date().toDateString();
            const todayEmails = result.history.filter(item => 
                new Date(item.sentAt).toDateString() === today
            );
            
            const successEmails = result.history.filter(item => item.status === 'success');
            const successRate = result.history.length > 0 ? 
                Math.round((successEmails.length / result.history.length) * 100) : 0;
            
            const lastProcessed = result.history.length > 0 ? 
                new Date(result.history[0].sentAt).toLocaleTimeString() : '-';
            
            document.getElementById('todayProcessed').textContent = todayEmails.length;
            document.getElementById('successRate').textContent = successRate + '%';
            document.getElementById('totalEmails').textContent = result.history.length;
            document.getElementById('lastProcessed').textContent = lastProcessed;
        }
    } catch (error) {
        console.error('대시보드 업데이트 오류:', error);
    }
}

// 유틸리티 함수들
function showStep(stepNumber) {
    document.getElementById(`step${stepNumber}`).classList.remove('hidden');
}

function showAlert(type, message) {
    const uploadAlert = document.getElementById('uploadAlert');
    const alertClass = type === 'success' ? 'alert-success' : 
                      type === 'warning' ? 'alert-warning' : 
                      type === 'info' ? 'alert-info' : 'alert-error';
    const icon = type === 'success' ? '●' : 
                type === 'warning' ? '▲' : 
                type === 'info' ? 'ℹ' : '●';
    
    uploadAlert.innerHTML = `
        <div class="alert ${alertClass}">
            ${icon} ${message}
        </div>
    `;
    
    // 3초 후 자동 제거
    setTimeout(() => {
        if (uploadAlert.innerHTML.includes(message)) {
            uploadAlert.innerHTML = '';
        }
    }, 3000);
}

function showLoading(message) {
    const uploadAlert = document.getElementById('uploadAlert');
    uploadAlert.innerHTML = `
        <div class="alert alert-success">
            <div class="loading"></div> ${message}
        </div>
    `;
}

function hideLoading() {
    const uploadAlert = document.getElementById('uploadAlert');
    uploadAlert.innerHTML = '';
}

// 진행률 표시 시작
function showProgress(message = '처리 중...') {
    const progressContainer = document.getElementById('progressContainer');
    const progressMessage = document.getElementById('progressMessage');
    const progressPercent = document.getElementById('progressPercent');
    const progressFill = document.getElementById('progressFill');
    
    progressMessage.textContent = message;
    progressPercent.textContent = '0%';
    progressFill.style.width = '0%';
    
    progressContainer.classList.remove('hidden');
}

// 진행률 업데이트
function updateProgress(percent, message = null) {
    const progressMessage = document.getElementById('progressMessage');
    const progressPercent = document.getElementById('progressPercent');
    const progressFill = document.getElementById('progressFill');
    
    if (message) {
        progressMessage.textContent = message;
    }
    
    progressPercent.textContent = `${percent}%`;
    progressFill.style.width = `${percent}%`;
}

// 진행률 숨기기
function hideProgress() {
    const progressContainer = document.getElementById('progressContainer');
    progressContainer.classList.add('hidden');
}

// 진행률 시뮬레이션 (실제 백엔드 진행률이 없을 경우)
function simulateProgress(steps, totalDuration = 3000) {
    return new Promise((resolve) => {
        let currentStep = 0;
        const stepDuration = totalDuration / steps.length;
        
        const processStep = () => {
            if (currentStep < steps.length) {
                const step = steps[currentStep];
                updateProgress(step.percent, step.message);
                currentStep++;
                setTimeout(processStep, stepDuration);
            } else {
                resolve();
            }
        };
        
        processStep();
    });
}

// 모든 단계 초기화
function resetAllSteps() {
    // 전역 변수 초기화 (중요!)
    currentFileId = null;
    currentMapping = {};
    generatedFileName = null;
    
    // STEP 2, 3, 4 숨기기
    document.getElementById('step2').classList.add('hidden');
    document.getElementById('step3').classList.add('hidden');
    document.getElementById('step4').classList.add('hidden');
    
    // 직접 입력 폼 숨기기
    const directInputStep = document.getElementById('directInputStep');
    if (directInputStep) {
        directInputStep.classList.add('hidden');
    }
    
    // 업로드 결과 초기화
    const uploadResult = document.getElementById('uploadResult');
    if (uploadResult) {
        uploadResult.classList.add('hidden');
    }
    
    // 생성 결과 초기화
    const generateResult = document.getElementById('generateResult');
    if (generateResult) {
        generateResult.innerHTML = '';
    }
    
    // 이메일 결과 초기화
    const emailResult = document.getElementById('emailResult');
    if (emailResult) {
        emailResult.innerHTML = '';
    }
    
    // 필수 필드 입력 폼 숨기기
    const missingFieldsForm = document.getElementById('missingFieldsForm');
    if (missingFieldsForm) {
        missingFieldsForm.classList.add('hidden');
    }
    
    // 파일 입력 초기화
    const fileInput = document.getElementById('fileInput');
    if (fileInput) {
        fileInput.value = '';
    }
    
    // 매핑 상태 초기화
    sessionStorage.setItem('mappingSaved', 'false');
    
    // 타겟 필드 초기화
    resetTargetFields();
    
    // GENERATE ORDER 버튼 비활성화
    setTimeout(() => {
        updateGenerateOrderButton();
    }, 100);
    
    // 진행률 숨기기
    hideProgress();
}

// 타겟 필드 초기화
function resetTargetFields() {
    const targetFields = document.querySelectorAll('#targetFields .field-item');
    targetFields.forEach(field => {
        // 원래 텍스트로 복원
        const targetName = field.dataset.target;
        field.innerHTML = targetName;
        
        // 스타일 초기화
        field.style.background = '';
        field.style.color = '';
        
        // 기본 클래스만 유지
        field.className = 'field-item';
    });
}

// 전체 선택/해제
function toggleSelectAll() {
    const selectAllCheckbox = document.getElementById('selectAllHistory');
    const historyCheckboxes = document.querySelectorAll('.history-checkbox');
    
    historyCheckboxes.forEach(checkbox => {
        checkbox.checked = selectAllCheckbox.checked;
    });
    
    updateDeleteButton();
}

// 삭제 버튼 상태 업데이트
function updateDeleteButton() {
    const checkedBoxes = document.querySelectorAll('.history-checkbox:checked');
    const deleteBtn = document.getElementById('deleteSelectedBtn');
    
    if (checkedBoxes.length > 0) {
        deleteBtn.style.display = 'inline-block';
    } else {
        deleteBtn.style.display = 'none';
    }
    
    // 전체 선택 체크박스 상태 업데이트
    const allCheckboxes = document.querySelectorAll('.history-checkbox');
    const selectAllCheckbox = document.getElementById('selectAllHistory');
    
    if (allCheckboxes.length > 0) {
        selectAllCheckbox.checked = checkedBoxes.length === allCheckboxes.length;
    }
}

// 선택된 이력 삭제
async function deleteSelectedHistory() {
    const checkedBoxes = document.querySelectorAll('.history-checkbox:checked');
    
    if (checkedBoxes.length === 0) {
        showAlert('warning', '삭제할 항목을 선택해주세요.');
        return;
    }
    
    if (!confirm(`선택된 ${checkedBoxes.length}개 항목을 삭제하시겠습니까?`)) {
        return;
    }
    
    try {
        showLoading('선택된 이력을 삭제하고 있습니다...');
        
        const indices = Array.from(checkedBoxes).map(checkbox => parseInt(checkbox.dataset.index));
        
        const response = await fetch('/api/email/history/delete', {
            method: 'DELETE',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ indices })
        });
        
        const result = await response.json();
        
        hideLoading();
        
        if (result.success) {
            showAlert('success', `${indices.length}개 항목이 삭제되었습니다.`);
            loadEmailHistory();
            updateDashboard();
        } else {
            showAlert('error', result.error || '이력 삭제에 실패했습니다.');
        }
        
    } catch (error) {
        hideLoading();
        console.error('이력 삭제 오류:', error);
        showAlert('error', '이력 삭제 중 오류가 발생했습니다.');
    }
}

// 단일 이력 삭제
async function deleteSingleHistory(index) {
    if (!confirm('이 이력을 삭제하시겠습니까?')) {
        return;
    }
    
    try {
        showLoading('이력을 삭제하고 있습니다...');
        
        const response = await fetch('/api/email/history/delete', {
            method: 'DELETE',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ indices: [index] })
        });
        
        const result = await response.json();
        
        hideLoading();
        
        if (result.success) {
            showAlert('success', '이력이 삭제되었습니다.');
            loadEmailHistory();
            updateDashboard();
        } else {
            showAlert('error', result.error || '이력 삭제에 실패했습니다.');
        }
        
    } catch (error) {
        hideLoading();
        console.error('이력 삭제 오류:', error);
        showAlert('error', '이력 삭제 중 오류가 발생했습니다.');
    }
}

// 전체 이력 삭제
async function clearAllHistory() {
    if (!confirm('모든 전송 이력을 삭제하시겠습니까? 이 작업은 되돌릴 수 없습니다.')) {
        return;
    }
    
    try {
        showLoading('모든 이력을 삭제하고 있습니다...');
        
        const response = await fetch('/api/email/history/clear', {
            method: 'DELETE'
        });
        
        const result = await response.json();
        
        hideLoading();
        
        if (result.success) {
            showAlert('success', '모든 이력이 삭제되었습니다.');
            loadEmailHistory();
            updateDashboard();
        } else {
            showAlert('error', result.error || '이력 삭제에 실패했습니다.');
        }
        
    } catch (error) {
        hideLoading();
        console.error('이력 삭제 오류:', error);
        showAlert('error', '이력 삭제 중 오류가 발생했습니다.');
    }
}

// 🎯 표준 타겟 필드 설정
function setupStandardTargetFields() {
    const targetFieldsContainer = document.getElementById('targetFields');
    targetFieldsContainer.innerHTML = '';
    
    // 표준 발주서 필수 필드 정의 (상품명, 연락처, 주소만 필수)
    const standardFields = [
        { name: '상품명', required: true },
        { name: '수량', required: false },
        { name: '단가', required: false },
        { name: '고객명', required: false },
        { name: '연락처', required: true },
        { name: '주소', required: true }
    ];
    
    standardFields.forEach(field => {
        const fieldDiv = document.createElement('div');
        fieldDiv.className = 'field-item';
        fieldDiv.dataset.target = field.name;
        fieldDiv.dataset.required = field.required ? 'true' : 'false';
        
        if (field.required) {
            fieldDiv.innerHTML = `${field.name} <span style="color: red;">*</span>`;
        } else {
            fieldDiv.textContent = field.name;
        }
        
        fieldDiv.onclick = () => selectTargetField(fieldDiv);
        targetFieldsContainer.appendChild(fieldDiv);
    });
    
    // 타겟 필드 초기화 (이전 매핑 상태 제거)
    resetTargetFields();
}

// 📊 필수 필드 검증 강화
function validateRequiredFields(mapping) {
    const requiredFields = ['상품명', '연락처', '주소'];
    const missingFields = [];
    
    requiredFields.forEach(field => {
        if (!mapping[field] || mapping[field].trim() === '') {
            missingFields.push(field);
        }
    });
    
    return {
        isValid: missingFields.length === 0,
        missingFields: missingFields,
        message: missingFields.length > 0 ? 
            `필수 필드가 매핑되지 않았습니다: ${missingFields.join(', ')}` : 
            '모든 필수 필드가 매핑되었습니다.'
    };
}

// 🔄 필수 필드 입력 폼 표시
function showMissingFieldsForm(missingFields) {
    const form = document.getElementById('missingFieldsForm');
    const container = document.getElementById('missingFieldsContainer');
    
    // 기존 내용 초기화
    container.innerHTML = '';
    
    // 각 누락된 필드에 대해 입력 필드 생성
    missingFields.forEach(field => {
        const fieldDiv = document.createElement('div');
        fieldDiv.className = 'form-group';
        fieldDiv.style.marginBottom = '15px';
        
        const label = document.createElement('label');
        label.textContent = field;
        label.style.fontWeight = '600';
        label.style.color = '#856404';
        label.style.marginBottom = '5px';
        label.style.display = 'block';
        
        const input = document.createElement('input');
        input.type = 'text';
        input.className = 'form-control';
        input.id = `missing_${field}`;
        input.placeholder = `${field}를 입력하세요`;
        input.style.width = '100%';
        input.style.padding = '8px 12px';
        input.style.border = '1px solid #dee2e6';
        input.style.borderRadius = '4px';
        input.style.fontSize = '0.9em';
        
        fieldDiv.appendChild(label);
        fieldDiv.appendChild(input);
        container.appendChild(fieldDiv);
    });
    
    // 폼 표시
    form.classList.remove('hidden');
    
    // 폼으로 스크롤
    form.scrollIntoView({ behavior: 'smooth' });
}

// 💾 필수 필드 저장
async function saveMissingFields() {
    const form = document.getElementById('missingFieldsForm');
    const inputs = form.querySelectorAll('input[id^="missing_"]');
    
    // 입력값 검증
    let hasEmptyFields = false;
    const fieldValues = {};
    
    inputs.forEach(input => {
        const fieldName = input.id.replace('missing_', '');
        const value = input.value.trim();
        
        if (value === '') {
            hasEmptyFields = true;
            input.style.borderColor = '#dc3545';
        } else {
            input.style.borderColor = '#dee2e6';
            fieldValues[fieldName] = value;
        }
    });
    
    if (hasEmptyFields) {
        showAlert('warning', '모든 필수 필드를 입력해주세요.');
        return;
    }
    
    try {
        // 현재 매핑에 입력값들을 추가 (고정값으로 설정)
        Object.keys(fieldValues).forEach(field => {
            currentMapping[field] = `[고정값: ${fieldValues[field]}]`;
        });
        
        // 매핑 저장
        const mappingData = {
            mappingName: `mapping_${Date.now()}`,
            sourceFields: Object.values(currentMapping),
            targetFields: Object.keys(currentMapping),
            mappingRules: currentMapping,
            fixedValues: fieldValues // 고정값들을 별도로 전송
        };
        
        showLoading('매핑 규칙을 저장하고 있습니다...');
        
        const response = await fetch('/api/orders/mapping', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(mappingData)
        });
        
        const result = await response.json();
        
        hideLoading();
        
        if (result.success) {
            // 타겟 필드들의 매핑 상태 업데이트
            Object.keys(fieldValues).forEach(field => {
                const targetField = document.querySelector(`[data-target="${field}"]`);
                if (targetField) {
                    targetField.classList.add('selected');
                    targetField.textContent = `${field} ← [고정값]`;
                }
            });
            
            showAlert('success', '✅ 필수 정보가 저장되었습니다. 매핑이 완료되었습니다.');
            
            // 매핑 저장 상태 표시
            sessionStorage.setItem('mappingSaved', 'true');
            
            // GENERATE ORDER 버튼 활성화
            updateGenerateOrderButton();
            
            // 폼 숨기기
            hideMissingFieldsForm();
            
        } else {
            showAlert('error', result.error || '매핑 저장에 실패했습니다.');
        }
        
    } catch (error) {
        hideLoading();
        console.error('필수 필드 저장 오류:', error);
        showAlert('error', '필수 필드 저장 중 오류가 발생했습니다.');
    }
}

// 🚫 필수 필드 입력 폼 숨기기
function hideMissingFieldsForm() {
    const form = document.getElementById('missingFieldsForm');
    form.classList.add('hidden');
}

// 📝 직접 입력 폼 표시
function showDirectInputForm() {
    // 모든 단계 숨기기
    resetAllSteps();
    
    // 직접 입력 폼 표시
    const directInputStep = document.getElementById('directInputStep');
    directInputStep.classList.remove('hidden');
    
    // 폼으로 스크롤
    directInputStep.scrollIntoView({ behavior: 'smooth' });
}

// 💾 직접 입력 데이터 저장 및 발주서 생성
async function saveDirectInput() {
    // 필수 필드 검증
    const requiredFields = ['상품명', '연락처', '주소'];
    const inputData = {};
    let hasEmptyRequired = false;
    
    // 모든 필드 값 수집
    ['상품명', '연락처', '주소', '수량', '단가', '고객명'].forEach(field => {
        const input = document.getElementById(`direct_${field}`);
        const value = input.value.trim();
        
        if (requiredFields.includes(field) && value === '') {
            hasEmptyRequired = true;
            input.style.borderColor = '#dc3545';
        } else {
            input.style.borderColor = '#dee2e6';
            if (value !== '') {
                inputData[field] = value;
            }
        }
    });
    
    if (hasEmptyRequired) {
        showAlert('warning', '필수 필드를 모두 입력해주세요. (상품명, 연락처, 주소)');
        return;
    }
    
    try {
        showLoading('직접 입력 데이터로 발주서를 생성하고 있습니다...');
        
        // 직접 입력 데이터를 매핑 형태로 변환
        const mappingData = {
            mappingName: `direct_input_${Date.now()}`,
            sourceFields: [],
            targetFields: Object.keys(inputData),
            mappingRules: {},
            fixedValues: inputData,
            isDirect: true // 직접 입력 플래그
        };
        
        // 매핑 저장
        const mappingResponse = await fetch('/api/orders/mapping', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(mappingData)
        });
        
        const mappingResult = await mappingResponse.json();
        
        if (!mappingResult.success) {
            throw new Error(mappingResult.error || '매핑 저장에 실패했습니다.');
        }
        
        // 직접 입력 데이터로 발주서 생성
        const generateResponse = await fetch('/api/orders/generate-direct', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({
                mappingId: mappingData.mappingName,
                inputData: inputData,
                templateType: 'standard'
            })
        });
        
        const generateResult = await generateResponse.json();
        
        hideLoading();
        
        if (generateResult.success) {
            generatedFileName = generateResult.generatedFile;
            
            // 성공 결과 표시
            showAlert('success', '✅ 직접 입력 데이터로 발주서가 생성되었습니다!');
            
            // 결과 표시 및 이메일 단계로 이동
            showDirectInputResult(generateResult);
            showStep(3);
            showStep(4);
            
        } else {
            showAlert('error', generateResult.error || '발주서 생성에 실패했습니다.');
        }
        
    } catch (error) {
        hideLoading();
        console.error('직접 입력 저장 오류:', error);
        showAlert('error', '직접 입력 처리 중 오류가 발생했습니다.');
    }
}

// 📋 직접 입력 결과 표시
function showDirectInputResult(result) {
    const generateResult = document.getElementById('generateResult');
    
    generateResult.innerHTML = `
        <div class="alert alert-success">
            ✅ 직접 입력 데이터로 발주서가 성공적으로 생성되었습니다!<br>
            <strong>입력된 정보:</strong> ${Object.keys(result.inputData || {}).length}개 필드<br>
            <strong>생성된 파일:</strong> ${result.generatedFile}
        </div>
        
        <div style="text-align: center; margin-top: 20px;">
            <a href="${result.downloadUrl}" class="btn btn-success" download>DOWNLOAD ORDER</a>
        </div>
    `;
}

// 🚫 직접 입력 취소
function cancelDirectInput() {
    // 직접 입력 폼의 입력값 초기화
    ['상품명', '연락처', '주소', '수량', '단가', '고객명'].forEach(field => {
        const input = document.getElementById(`direct_${field}`);
        if (input) {
            input.value = '';
            input.style.borderColor = '#dee2e6';
        }
    });
    
    // 모든 상태 초기화 (resetAllSteps 사용)
    resetAllSteps();
    
    // 1단계만 표시
    const step1 = document.getElementById('step1');
    if (step1) {
        step1.classList.remove('hidden');
    }
    
    console.log('🔄 직접 입력 취소: 모든 상태 초기화 완료');
} 