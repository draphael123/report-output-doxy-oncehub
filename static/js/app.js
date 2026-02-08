// File Upload State Management
const uploadState = {
    files: {
        doxy: null,
        account: null,
        gusto: null,
        booking: null
    },
    validated: {
        doxy: false,
        account: false,
        gusto: false,
        booking: false
    }
};

// Initialize on page load
document.addEventListener('DOMContentLoaded', () => {
    // Setup all dropzones
    setupDropzone('doxy-dropzone', 'doxy_file', 'doxy');
    setupDropzone('account-dropzone', 'account_file', 'account');
    setupDropzone('gusto-dropzone', 'gusto_file', 'gusto');
    setupDropzone('booking-dropzone', 'booking_file', 'booking');
    
    // Load report history
    loadReportHistory();
    
    // Setup preview button
    document.getElementById('preview-btn').addEventListener('click', showPreview);
    
    // Setup modal close handlers
    document.getElementById('close-preview').addEventListener('click', () => hideModal('preview-modal'));
    document.getElementById('cancel-preview').addEventListener('click', () => hideModal('preview-modal'));
    document.getElementById('confirm-generate').addEventListener('click', () => {
        hideModal('preview-modal');
        document.getElementById('upload-form').submit();
    });
    
    // Setup form submission
    document.getElementById('upload-form').addEventListener('submit', handleFormSubmit);
    
    // Update button states
    updateButtonStates();
});

// Drag-and-Drop Handlers
function setupDropzone(dropzoneId, fileInputId, fileType) {
    const dropzone = document.getElementById(dropzoneId);
    const fileInput = document.getElementById(fileInputId);
    const statusDiv = document.getElementById(`${fileType}-status`);
    const previewDiv = document.getElementById(`${fileType}-preview`);
    
    // Click to browse
    dropzone.addEventListener('click', (e) => {
        if (e.target.closest('.dropzone-link') || !fileInput.files.length) {
            fileInput.click();
        }
    });
    
    // Drag events
    dropzone.addEventListener('dragover', (e) => {
        e.preventDefault();
        dropzone.classList.add('dragover');
    });
    
    dropzone.addEventListener('dragleave', () => {
        dropzone.classList.remove('dragover');
    });
    
    dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropzone.classList.remove('dragover');
        
        if (e.dataTransfer.files.length > 0) {
            fileInput.files = e.dataTransfer.files;
            handleFileSelect(fileInput.files[0], fileType, statusDiv, previewDiv, dropzone);
        }
    });
    
    // File input change
    fileInput.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFileSelect(e.target.files[0], fileType, statusDiv, previewDiv, dropzone);
        }
    });
}

// Handle file selection
async function handleFileSelect(file, fileType, statusDiv, previewDiv, dropzone) {
    // Update state
    uploadState.files[fileType] = file;
    uploadState.validated[fileType] = false;
    
    // Show loading
    statusDiv.innerHTML = '<div class="loading"></div> Validating...';
    dropzone.classList.remove('success', 'error');
    
    // Client-side validation
    const clientValidation = validateFile(file, fileType);
    
    if (!clientValidation.valid) {
        showValidationResult(statusDiv, { valid: false, errors: clientValidation.errors, warnings: [] }, fileType);
        dropzone.classList.add('error');
        updateButtonStates();
        return;
    }
    
    // Server-side validation
    try {
        const result = await validateFileServer(file, fileType);
        showValidationResult(statusDiv, result, fileType);
        
        if (result.valid) {
            dropzone.classList.add('success');
            dropzone.classList.remove('error');
            uploadState.validated[fileType] = true;
            
            // Show preview if available
            if (result.preview) {
                showFilePreview(previewDiv, result.preview);
            }
        } else {
            dropzone.classList.add('error');
            dropzone.classList.remove('success');
        }
    } catch (error) {
        showValidationResult(statusDiv, { valid: false, errors: [error.message], warnings: [] }, fileType);
        dropzone.classList.add('error');
    }
    
    updateButtonStates();
}

// Client-Side File Validation
function validateFile(file, fileType) {
    const errors = [];
    const warnings = [];
    
    // File size check (max 10MB)
    const maxSize = 10 * 1024 * 1024; // 10MB
    if (file.size > maxSize) {
        errors.push(`File too large (${(file.size / 1024 / 1024).toFixed(1)}MB). Maximum is 10MB.`);
    }
    
    // File extension check
    const allowedExtensions = ['.csv', '.xls', '.xlsx'];
    const fileExt = '.' + file.name.split('.').pop().toLowerCase();
    if (!allowedExtensions.includes(fileExt)) {
        errors.push('Invalid file type. Please upload CSV or Excel file.');
    }
    
    return {
        valid: errors.length === 0,
        errors: errors,
        warnings: warnings
    };
}

// Server-Side Validation API Call
async function validateFileServer(file, fileType) {
    const formData = new FormData();
    formData.append('file', file);
    formData.append('file_type', fileType);
    
    try {
        const response = await fetch('/api/validate-file', {
            method: 'POST',
            body: formData
        });
        
        if (!response.ok) {
            // Try to get error message from response
            let errorMessage = 'Validation request failed';
            try {
                const errorData = await response.json();
                if (errorData.errors && errorData.errors.length > 0) {
                    errorMessage = errorData.errors[0];
                } else if (errorData.error) {
                    errorMessage = errorData.error;
                }
            } catch (e) {
                // If response is not JSON, use status text
                errorMessage = `Validation failed: ${response.statusText || response.status}`;
            }
            throw new Error(errorMessage);
        }
        
        return await response.json();
    } catch (error) {
        return {
            valid: false,
            errors: [error.message],
            warnings: [],
            info: []
        };
    }
}

// Show Validation Result
function showValidationResult(statusDiv, result, fileType) {
    if (result.valid) {
        let html = '<div class="validation-success"><strong>✓ File is valid</strong>';
        if (result.info && result.info.length > 0) {
            html += '<ul>';
            result.info.forEach(info => {
                html += `<li>${info}</li>`;
            });
            html += '</ul>';
        }
        html += '</div>';
        statusDiv.innerHTML = html;
    } else {
        let html = '<div class="validation-error"><strong>✗ Validation failed</strong><ul>';
        result.errors.forEach(error => {
            html += `<li class="error">${error}</li>`;
        });
        html += '</ul></div>';
        
        if (result.warnings && result.warnings.length > 0) {
            html += '<div class="validation-warnings"><strong>⚠ Warnings:</strong><ul>';
            result.warnings.forEach(warning => {
                html += `<li class="warning">${warning}</li>`;
            });
            html += '</ul></div>';
        }
        
        statusDiv.innerHTML = html;
    }
}

// Show File Preview
function showFilePreview(previewDiv, preview) {
    if (!preview || !preview.sample_data || preview.sample_data.length === 0) {
        previewDiv.style.display = 'none';
        return;
    }
    
    previewDiv.style.display = 'block';
    previewDiv.classList.add('show');
    
    let html = '<div class="file-preview-content"><h4>Preview (first 5 rows)</h4>';
    html += `<p><strong>Columns:</strong> ${preview.columns.join(', ')}</p>`;
    html += '<table class="preview-table"><thead><tr>';
    
    preview.columns.forEach(col => {
        html += `<th>${col}</th>`;
    });
    html += '</tr></thead><tbody>';
    
    preview.sample_data.forEach(row => {
        html += '<tr>';
        preview.columns.forEach(col => {
            html += `<td>${row[col] || ''}</td>`;
        });
        html += '</tr>';
    });
    
    html += '</tbody></table></div>';
    previewDiv.innerHTML = html;
}

// Update Button States
function updateButtonStates() {
    const requiredFiles = ['doxy', 'account', 'gusto'];
    const allRequiredValid = requiredFiles.every(type => uploadState.validated[type]);
    
    const previewBtn = document.getElementById('preview-btn');
    const generateBtn = document.getElementById('generate-btn');
    
    previewBtn.disabled = !allRequiredValid;
    generateBtn.disabled = !allRequiredValid;
    
    if (allRequiredValid) {
        updateProgressSteps(2);
    }
}

// Progress Steps Update
function updateProgressSteps(currentStep) {
    const steps = document.querySelectorAll('.step');
    steps.forEach((step, index) => {
        const stepNum = index + 1;
        step.classList.remove('active', 'complete');
        
        if (stepNum < currentStep) {
            step.classList.add('complete');
        } else if (stepNum === currentStep) {
            step.classList.add('active');
        }
    });
}

// Preview Functionality
async function showPreview() {
    const form = document.getElementById('upload-form');
    const formData = new FormData(form);
    
    // Check if required files are uploaded
    const requiredFiles = ['doxy', 'account', 'gusto'];
    const missingFiles = requiredFiles.filter(type => !uploadState.files[type]);
    
    if (missingFiles.length > 0) {
        showToast('Please upload all required files first', 'error');
        return;
    }
    
    const previewBtn = document.getElementById('preview-btn');
    previewBtn.disabled = true;
    previewBtn.textContent = '⏳ Loading...';
    
    try {
        const response = await fetch('/preview', {
            method: 'POST',
            body: formData
        });
        
        if (!response.ok) {
            const error = await response.json();
            throw new Error(error.error || 'Preview failed');
        }
        
        const data = await response.json();
        
        // Populate preview modal
        populatePreviewModal(data);
        showModal('preview-modal');
        
    } catch (error) {
        showToast(`Preview error: ${error.message}`, 'error');
    } finally {
        previewBtn.disabled = false;
        previewBtn.textContent = '👁️ Preview Data';
    }
}

// Populate Preview Modal
function populatePreviewModal(data) {
    // Summary stats
    const statsHtml = `
        <div class="preview-stat">
            <div class="preview-stat-number">${data.summary.total_providers || 0}</div>
            <div class="preview-stat-label">Providers</div>
        </div>
        <div class="preview-stat">
            <div class="preview-stat-number">${data.summary.total_doxy_visits || 0}</div>
            <div class="preview-stat-label">Total Visits</div>
        </div>
        <div class="preview-stat">
            <div class="preview-stat-number">${data.summary.total_hours_worked || 0}</div>
            <div class="preview-stat-label">Hours Worked</div>
        </div>
    `;
    document.getElementById('preview-stats').innerHTML = statsHtml;
    
    // Tabs and table
    let tabsHtml = '';
    let activeSheet = null;
    
    data.sheets.forEach((sheet, index) => {
        if (sheet.name === 'OnceHub Visits' && !sheet.available) {
            return;
        }
        
        const isActive = index === 0;
        const tabId = `tab-${index}`;
        
        tabsHtml += `
            <button class="preview-tab ${isActive ? 'active' : ''}" data-tab="${tabId}" data-sheet-index="${index}">
                ${sheet.name}
            </button>
        `;
        
        if (isActive) {
            activeSheet = sheet;
        }
    });
    
    document.getElementById('preview-tabs').innerHTML = tabsHtml;
    
    // Tab click handlers
    document.querySelectorAll('.preview-tab').forEach(tab => {
        tab.addEventListener('click', () => {
            const sheetIndex = parseInt(tab.dataset.sheetIndex);
            const sheet = data.sheets[sheetIndex];
            
            // Update active tab
            document.querySelectorAll('.preview-tab').forEach(t => t.classList.remove('active'));
            tab.classList.add('active');
            
            // Update table
            renderPreviewTable(sheet);
        });
    });
    
    // Render initial table
    if (activeSheet) {
        renderPreviewTable(activeSheet);
    }
    
    // Alerts
    const alertsHtml = '';
    document.getElementById('preview-alerts').innerHTML = alertsHtml;
}

// Render Preview Table
function renderPreviewTable(sheet) {
    const tableContainer = document.getElementById('preview-table').parentElement;
    
    if (!sheet.columns || sheet.columns.length === 0 || !sheet.sample_data || sheet.sample_data.length === 0) {
        tableContainer.innerHTML = '<div class="empty-sheet-message" style="padding: 2rem; text-align: center; color: var(--gray-500);">No data available for this sheet</div>';
        return;
    }
    
    let html = '<table class="preview-table"><thead><tr>';
    sheet.columns.forEach(col => {
        html += `<th>${col}</th>`;
    });
    html += '</tr></thead><tbody>';
    
    const rows = sheet.sample_data || [];
    const maxRows = Math.min(rows.length, 20);
    
    for (let i = 0; i < maxRows; i++) {
        html += '<tr>';
        sheet.columns.forEach(col => {
            const value = rows[i][col] !== undefined && rows[i][col] !== null ? rows[i][col] : '';
            html += `<td>${value}</td>`;
        });
        html += '</tr>';
    }
    
    if (sheet.rows > maxRows) {
        html += `<tr><td colspan="${sheet.columns.length}" style="text-align: center; color: var(--gray-500);">... and ${sheet.rows - maxRows} more rows</td></tr>`;
    }
    
    html += '</tbody></table>';
    tableContainer.innerHTML = html;
}

// Form Submission Handler
async function handleFormSubmit(e) {
    e.preventDefault();
    
    updateProgressSteps(3);
    showModal('progress-modal');
    
    const form = document.getElementById('upload-form');
    const formData = new FormData(form);
    
    // Update progress stages
    updateProgressStage('upload', 'active');
    
    try {
        const xhr = new XMLHttpRequest();
        
        // Track upload progress
        xhr.upload.addEventListener('progress', (e) => {
            if (e.lengthComputable) {
                const percentComplete = (e.loaded / e.total) * 100;
                if (percentComplete >= 100) {
                    updateProgressStage('upload', 'complete');
                    updateProgressStage('parse', 'active');
                }
            }
        });
        
        xhr.addEventListener('load', () => {
            if (xhr.status === 200) {
                updateProgressStage('parse', 'complete');
                updateProgressStage('validate', 'complete');
                updateProgressStage('calculate', 'complete');
                updateProgressStage('generate', 'complete');
                
                // Trigger download
                const blob = new Blob([xhr.response], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                
                const contentDisposition = xhr.getResponseHeader('Content-Disposition');
                const filename = contentDisposition 
                    ? contentDisposition.split('filename=')[1]?.replace(/"/g, '') 
                    : 'Weekly Report.xlsx';
                a.download = filename;
                
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                document.body.removeChild(a);
                
                // Hide progress modal
                setTimeout(() => {
                    hideModal('progress-modal');
                    updateProgressSteps(4);
                    showToast('Report generated successfully!', 'success');
                    
                    // Reload history
                    loadReportHistory();
                }, 500);
            } else {
                throw new Error('Report generation failed');
            }
        });
        
        xhr.addEventListener('error', () => {
            throw new Error('Network error');
        });
        
        xhr.open('POST', '/');
        xhr.responseType = 'blob';
        xhr.send(formData);
        
    } catch (error) {
        hideModal('progress-modal');
        showToast(`Error: ${error.message}`, 'error');
    }
}

// Update Progress Stage
function updateProgressStage(stage, status) {
    const stageEl = document.querySelector(`[data-stage="${stage}"]`);
    if (stageEl) {
        stageEl.classList.remove('active', 'complete');
        if (status === 'active') {
            stageEl.classList.add('active');
            stageEl.querySelector('.stage-icon').textContent = '⏳';
        } else if (status === 'complete') {
            stageEl.classList.add('complete');
            stageEl.querySelector('.stage-icon').textContent = '✓';
        }
    }
}

// Modal Controls
function showModal(modalId) {
    const modal = document.getElementById(modalId);
    if (modal) {
        modal.classList.add('show');
    }
}

function hideModal(modalId) {
    const modal = document.getElementById(modalId);
    if (modal) {
        modal.classList.remove('show');
    }
}

// Toast Notifications
function showToast(message, type = 'info') {
    const container = document.getElementById('toast-container');
    const toast = document.createElement('div');
    toast.className = `toast ${type}`;
    
    const icons = {
        success: '✓',
        error: '✗',
        warning: '⚠',
        info: 'ℹ'
    };
    
    toast.innerHTML = `
        <span class="toast-icon">${icons[type] || 'ℹ'}</span>
        <span class="toast-message">${message}</span>
        <button class="toast-close" onclick="this.parentElement.remove()">&times;</button>
    `;
    
    container.appendChild(toast);
    
    // Auto-dismiss after 5 seconds
    setTimeout(() => {
        toast.classList.add('hiding');
        setTimeout(() => toast.remove(), 300);
    }, 5000);
}

// Report History Loading
async function loadReportHistory() {
    try {
        const response = await fetch('/api/report-history');
        if (!response.ok) {
            return;
        }
        
        const data = await response.json();
        const historyDiv = document.getElementById('report-history');
        
        if (data.reports && data.reports.length > 0) {
            historyDiv.innerHTML = data.reports.map(report => `
                <div class="report-item" onclick="downloadReport('${report.filename}')">
                    <div class="report-item-name">${report.filename}</div>
                    <div class="report-item-date">${report.date}</div>
                </div>
            `).join('');
        } else {
            historyDiv.innerHTML = '<div class="history-empty">No reports yet</div>';
        }
    } catch (error) {
        console.error('Error loading history:', error);
    }
}

// Download Report
function downloadReport(filename) {
    window.location.href = `/download/${filename}`;
}

// Make functions globally available
window.showToast = showToast;
window.downloadReport = downloadReport;

