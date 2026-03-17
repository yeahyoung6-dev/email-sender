// 全局状态
let sessionId = null;
let excelData = null;
let progressInterval = null;

// DOM 元素
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');

// 初始化
document.addEventListener('DOMContentLoaded', () => {
    initUpload();
});

// 初始化上传区域
function initUpload() {
    // 点击上传
    uploadArea.addEventListener('click', () => {
        fileInput.click();
    });

    // 文件选择
    fileInput.addEventListener('change', handleFileSelect);

    // 拖拽上传
    uploadArea.addEventListener('dragover', (e) => {
        e.preventDefault();
        uploadArea.classList.add('drag-over');
    });

    uploadArea.addEventListener('dragleave', () => {
        uploadArea.classList.remove('drag-over');
    });

    uploadArea.addEventListener('drop', (e) => {
        e.preventDefault();
        uploadArea.classList.remove('drag-over');
        const files = e.dataTransfer.files;
        if (files.length > 0) {
            handleFile(files[0]);
        }
    });
}

// 处理文件选择
function handleFileSelect(e) {
    const file = e.target.files[0];
    if (file) {
        handleFile(file);
    }
}

// 处理文件上传
async function handleFile(file) {
    // 验证文件类型
    const validTypes = [
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        'application/vnd.ms-excel'
    ];
    const extension = file.name.split('.').pop().toLowerCase();

    if (!['xlsx', 'xls'].includes(extension)) {
        alert('请上传Excel文件（.xlsx 或 .xls）');
        return;
    }

    // 上传文件
    const formData = new FormData();
    formData.append('file', file);

    try {
        const response = await fetch('/api/upload', {
            method: 'POST',
            body: formData
        });

        const result = await response.json();

        if (!response.ok) {
            alert('上传失败: ' + result.error);
            return;
        }

        sessionId = result.sessionId;
        excelData = result.preview;

        // 显示文件信息
        document.getElementById('fileInfo').classList.remove('hidden');
        document.getElementById('fileName').textContent = file.name;
        document.getElementById('rowCount').textContent = result.rowCount;

        // 显示预览
        renderPreview(result.preview);

        // 启用下一步按钮
        document.getElementById('step2Next').disabled = false;

    } catch (error) {
        alert('上传出错: ' + error.message);
    }
}

// 渲染预览表格
function renderPreview(data) {
    const tbody = document.getElementById('previewBody');
    tbody.innerHTML = '';

    data.forEach(row => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${row['姓名'] || ''}</td>
            <td>${row['人员编码'] || ''}</td>
            <td>${row['卡号'] || ''}</td>
            <td>${row['邮箱地址'] || ''}</td>
        `;
        tbody.appendChild(tr);
    });

    document.getElementById('dataPreview').classList.remove('hidden');
}

// 步骤导航
function nextStep(current) {
    // 验证当前步骤
    if (current === 1) {
        const host = document.getElementById('smtpHost').value.trim();
        const port = document.getElementById('smtpPort').value;
        const user = document.getElementById('smtpUser').value.trim();
        const pass = document.getElementById('smtpPass').value;
        const senderName = document.getElementById('senderName').value.trim();

        if (!host || !port || !user || !pass || !senderName) {
            alert('请填写完整的SMTP配置信息');
            return;
        }
    }

    if (current === 2 && !sessionId) {
        alert('请先上传Excel文件');
        return;
    }

    if (current === 3) {
        const subject = document.getElementById('emailSubject').value.trim();
        const body = document.getElementById('emailBody').value.trim();

        if (!subject || !body) {
            alert('请填写邮件主题和正文');
            return;
        }

        // 更新预览
        updateEmailPreview();
    }

    // 切换到下一步
    document.getElementById(`step${current}`).classList.add('hidden');
    document.getElementById(`step${current + 1}`).classList.remove('hidden');

    // 更新步骤指示器
    updateStepIndicator(current + 1);
}

function prevStep(current) {
    document.getElementById(`step${current}`).classList.add('hidden');
    document.getElementById(`step${current - 1}`).classList.remove('hidden');
    updateStepIndicator(current - 1);
}

function updateStepIndicator(activeStep) {
    document.querySelectorAll('.step').forEach((step, index) => {
        step.classList.remove('active', 'completed');
        if (index + 1 < activeStep) {
            step.classList.add('completed');
        } else if (index + 1 === activeStep) {
            step.classList.add('active');
        }
    });
}

// 更新邮件预览
function updateEmailPreview() {
    const subject = document.getElementById('emailSubject').value;
    const body = document.getElementById('emailBody').value;

    if (excelData && excelData.length > 0) {
        const firstRow = excelData[0];
        document.getElementById('previewSubject').textContent = replaceVariables(subject, firstRow);
        document.getElementById('previewBody').textContent = replaceVariables(body, firstRow);

        // 更新摘要
        const rowCount = document.getElementById('rowCount').textContent;
        document.getElementById('summaryCount').textContent = rowCount;
        const minutes = Math.ceil((parseInt(rowCount) * 3) / 60);
        document.getElementById('estimatedTime').textContent = `约${minutes}分钟`;
    }
}

// 替换模板变量
function replaceVariables(template, data) {
    let result = template;
    const fields = ['姓名', '人员编码', '卡号', '密码', '邮箱地址'];
    fields.forEach(field => {
        const regex = new RegExp(`\\{${field}\\}`, 'g');
        result = result.replace(regex, data[field] || '');
    });
    return result;
}

// 发送邮件
async function sendEmails() {
    const smtpConfig = {
        host: document.getElementById('smtpHost').value.trim(),
        port: parseInt(document.getElementById('smtpPort').value),
        user: document.getElementById('smtpUser').value.trim(),
        pass: document.getElementById('smtpPass').value,
        senderName: document.getElementById('senderName').value.trim()
    };

    const emailTemplate = {
        subject: document.getElementById('emailSubject').value.trim(),
        body: document.getElementById('emailBody').value.trim()
    };

    // 确认发送
    const rowCount = document.getElementById('rowCount').textContent;
    if (!confirm(`确定要向 ${rowCount} 个收件人发送邮件吗？\n预计耗时约 ${Math.ceil((parseInt(rowCount) * 3) / 60)} 分钟`)) {
        return;
    }

    // 禁用按钮
    document.getElementById('sendBtn').disabled = true;
    document.getElementById('sendBtn').textContent = '发送中...';

    // 显示进度区域
    document.getElementById('progressSection').classList.remove('hidden');

    try {
        const response = await fetch('/api/send', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                sessionId,
                smtpConfig,
                emailTemplate
            })
        });

        const result = await response.json();

        if (!response.ok) {
            alert('发送失败: ' + result.error);
            document.getElementById('sendBtn').disabled = false;
            document.getElementById('sendBtn').textContent = '开始发送';
            return;
        }

        // 开始轮询进度
        startProgressPolling();

    } catch (error) {
        alert('发送出错: ' + error.message);
        document.getElementById('sendBtn').disabled = false;
        document.getElementById('sendBtn').textContent = '开始发送';
    }
}

// 轮询发送进度
function startProgressPolling() {
    progressInterval = setInterval(async () => {
        try {
            const response = await fetch(`/api/progress/${sessionId}`);
            const progress = await response.json();

            updateProgress(progress);

            if (progress.status === 'completed') {
                clearInterval(progressInterval);
                onSendComplete(progress);
            }
        } catch (error) {
            console.error('获取进度失败:', error);
        }
    }, 1000);
}

// 更新进度显示
function updateProgress(progress) {
    const percent = Math.round((progress.current / progress.total) * 100);
    document.getElementById('progressFill').style.width = percent + '%';
    document.getElementById('progressText').textContent = `${progress.current} / ${progress.total}`;
    document.getElementById('successCount').textContent = progress.successCount;
    document.getElementById('failCount').textContent = progress.failCount;
}

// 发送完成
function onSendComplete(progress) {
    // 显示结果表格
    document.getElementById('sendResults').classList.remove('hidden');

    const tbody = document.getElementById('resultsBody');
    tbody.innerHTML = '';

    progress.results.forEach(result => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${result.name}</td>
            <td>${result.email}</td>
            <td class="${result.success ? 'result-success' : 'result-fail'}">
                ${result.success ? '成功' : '失败'}
            </td>
            <td>${result.error || ''}</td>
        `;
        tbody.appendChild(tr);
    });

    // 切换按钮
    document.getElementById('sendBtns').classList.add('hidden');
    document.getElementById('restartBtns').classList.remove('hidden');

    // 显示完成提示
    alert(`发送完成！\n成功: ${progress.successCount}\n失败: ${progress.failCount}`);
}

// 重置
function resetAll() {
    sessionId = null;
    excelData = null;

    // 重置表单
    document.getElementById('smtpHost').value = '';
    document.getElementById('smtpPort').value = '465';
    document.getElementById('smtpUser').value = '';
    document.getElementById('smtpPass').value = '';
    document.getElementById('senderName').value = '';
    document.getElementById('emailSubject').value = '';
    document.getElementById('emailBody').value = '';

    // 重置文件上传
    fileInput.value = '';
    document.getElementById('fileInfo').classList.add('hidden');
    document.getElementById('dataPreview').classList.add('hidden');
    document.getElementById('step2Next').disabled = true;

    // 重置进度
    document.getElementById('progressSection').classList.add('hidden');
    document.getElementById('sendResults').classList.add('hidden');
    document.getElementById('progressFill').style.width = '0%';
    document.getElementById('successCount').textContent = '0';
    document.getElementById('failCount').textContent = '0';

    // 重置按钮
    document.getElementById('sendBtns').classList.remove('hidden');
    document.getElementById('restartBtns').classList.add('hidden');
    document.getElementById('sendBtn').disabled = false;
    document.getElementById('sendBtn').textContent = '开始发送';

    // 切换到第一步
    document.querySelectorAll('.panel').forEach(panel => {
        panel.classList.add('hidden');
    });
    document.getElementById('step1').classList.remove('hidden');
    updateStepIndicator(1);
}