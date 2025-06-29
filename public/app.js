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
            showAlert('success', '매핑 규칙이 저장되었습니다.');
            
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
    // STEP 2, 3, 4 숨기기
    document.getElementById('step2').classList.add('hidden');
    document.getElementById('step3').classList.add('hidden');
    document.getElementById('step4').classList.add('hidden');
    
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
    
    // 매핑 상태 초기화
    sessionStorage.setItem('mappingSaved', 'false');
    
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