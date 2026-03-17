const express = require('express');
const multer = require('multer');
const nodemailer = require('nodemailer');
const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');
const { v4: uuidv4 } = require('uuid');

const app = express();
const PORT = 3000;

// 确保uploads目录存在
const uploadsDir = path.join(__dirname, 'uploads');
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true });
}

// 中间件
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// 配置文件上传
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    cb(null, path.join(__dirname, 'uploads/'));
  },
  filename: (req, file, cb) => {
    cb(null, Date.now() + '-' + file.originalname);
  }
});
const upload = multer({ storage });

// 发送进度存储
const progressStore = new Map();

// 解析Excel文件
function parseExcel(filePath) {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(worksheet);
  return data;
}

// 验证Excel数据格式
function validateExcelData(data) {
  const requiredFields = ['姓名', '人员编码', '卡号', '密码', '邮箱地址'];
  const errors = [];

  if (!data || data.length === 0) {
    return { valid: false, errors: ['Excel文件为空'] };
  }

  const firstRow = data[0];
  const missingFields = requiredFields.filter(field => !(field in firstRow));

  if (missingFields.length > 0) {
    errors.push(`缺少必要字段: ${missingFields.join(', ')}`);
  }

  // 验证邮箱格式
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  data.forEach((row, index) => {
    if (row['邮箱地址'] && !emailRegex.test(row['邮箱地址'])) {
      errors.push(`第${index + 2}行邮箱格式不正确: ${row['邮箱地址']}`);
    }
  });

  return {
    valid: errors.length === 0,
    errors,
    rowCount: data.length
  };
}

// 替换模板变量
function replaceVariables(template, data) {
  let result = template;
  Object.keys(data).forEach(key => {
    const regex = new RegExp(`\\{${key}\\}`, 'g');
    result = result.replace(regex, String(data[key] || ''));
  });
  return result;
}

// API: 上传Excel文件
app.post('/api/upload', upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: '请上传文件' });
    }

    const filePath = req.file.path;
    const data = parseExcel(filePath);
    const validation = validateExcelData(data);

    if (!validation.valid) {
      fs.unlinkSync(filePath);
      return res.status(400).json({ error: validation.errors.join('; ') });
    }

    // 生成会话ID并存储数据
    const sessionId = uuidv4();
    progressStore.set(sessionId, {
      data,
      filePath,
      progress: { current: 0, total: data.length, status: 'pending' }
    });

    // 返回预览数据（隐藏敏感信息）
    const preview = data.slice(0, 5).map(row => {
      const cardNum = String(row['卡号'] || '');
      return {
        姓名: row['姓名'],
        人员编码: row['人员编码'],
        卡号: cardNum ? cardNum.substring(0, 4) + '****' : '',
        密码: '****',
        邮箱地址: row['邮箱地址']
      };
    });

    res.json({
      success: true,
      sessionId,
      rowCount: data.length,
      preview
    });
  } catch (error) {
    console.error('上传错误:', error);
    res.status(500).json({ error: '文件处理失败: ' + error.message });
  }
});

// API: 发送邮件
app.post('/api/send', async (req, res) => {
  const { sessionId, smtpConfig, emailTemplate } = req.body;

  if (!sessionId || !progressStore.has(sessionId)) {
    return res.status(400).json({ error: '无效的会话，请重新上传文件' });
  }

  const session = progressStore.get(sessionId);
  const { data, filePath } = session;

  // 创建邮件传输器
  const transporter = nodemailer.createTransport({
    host: smtpConfig.host,
    port: smtpConfig.port,
    secure: smtpConfig.port === 465,
    auth: {
      user: smtpConfig.user,
      pass: smtpConfig.pass
    }
  });

  // 更新状态
  session.progress.status = 'sending';
  session.progress.current = 0;
  session.progress.total = data.length;
  session.progress.results = [];
  session.progress.startTime = Date.now();

  res.json({ success: true, message: '开始发送...' });

  // 异步发送邮件
  const sendInterval = 3000; // 3秒间隔

  for (let i = 0; i < data.length; i++) {
    const row = data[i];

    try {
      const subject = replaceVariables(emailTemplate.subject, row);
      const body = replaceVariables(emailTemplate.body, row);

      await transporter.sendMail({
        from: `"${smtpConfig.senderName}" <${smtpConfig.user}>`,
        to: row['邮箱地址'],
        subject,
        text: body
      });

      session.progress.results.push({
        email: row['邮箱地址'],
        name: row['姓名'],
        success: true
      });
    } catch (error) {
      session.progress.results.push({
        email: row['邮箱地址'],
        name: row['姓名'],
        success: false,
        error: error.message
      });
    }

    session.progress.current = i + 1;

    // 等待间隔（最后一封不需要等待）
    if (i < data.length - 1) {
      await new Promise(resolve => setTimeout(resolve, sendInterval));
    }
  }

  session.progress.status = 'completed';
  session.progress.endTime = Date.now();

  // 删除临时文件
  try {
    fs.unlinkSync(filePath);
  } catch (e) {
    console.error('删除临时文件失败:', e);
  }
});

// API: 获取发送进度
app.get('/api/progress/:sessionId', (req, res) => {
  const { sessionId } = req.params;

  if (!progressStore.has(sessionId)) {
    return res.status(404).json({ error: '会话不存在' });
  }

  const session = progressStore.get(sessionId);
  const { progress } = session;

  const successCount = progress.results?.filter(r => r.success).length || 0;
  const failCount = progress.results?.filter(r => !r.success).length || 0;

  res.json({
    status: progress.status,
    current: progress.current,
    total: progress.total,
    successCount,
    failCount,
    results: progress.results || [],
    startTime: progress.startTime,
    endTime: progress.endTime
  });
});

// 启动服务器
app.listen(PORT, () => {
  console.log(`邮件发送服务已启动: http://localhost:${PORT}`);
  console.log('请在浏览器中打开上述地址使用');
});