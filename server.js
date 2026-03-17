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
    cb(null, uploadsDir);
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

// 验证Excel数据格式 - 只要求有邮箱地址列
function validateExcelData(data) {
  const errors = [];

  if (!data || data.length === 0) {
    return { valid: false, errors: ['Excel文件为空'] };
  }

  // 获取所有列名
  const columns = Object.keys(data[0]);

  // 查找邮箱列（支持多种命名）
  const emailColumnNames = ['邮箱地址', '邮箱', 'email', 'Email', 'E-mail', '电子邮件', '邮件地址'];
  let emailColumn = null;

  for (const name of emailColumnNames) {
    if (columns.includes(name)) {
      emailColumn = name;
      break;
    }
  }

  // 如果没找到明确邮箱列，尝试查找包含"邮箱"或"email"的列
  if (!emailColumn) {
    for (const col of columns) {
      if (col.includes('邮箱') || col.toLowerCase().includes('email')) {
        emailColumn = col;
        break;
      }
    }
  }

  if (!emailColumn) {
    return {
      valid: false,
      errors: ['Excel中未找到邮箱地址列，请确保有一列名为"邮箱地址"、"邮箱"或"email"']
    };
  }

  // 验证邮箱格式
  const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  data.forEach((row, index) => {
    const email = row[emailColumn];
    if (email && !emailRegex.test(String(email))) {
      errors.push(`第${index + 2}行邮箱格式不正确: ${email}`);
    }
  });

  return {
    valid: errors.length === 0,
    errors,
    rowCount: data.length,
    columns,
    emailColumn
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
      emailColumn: validation.emailColumn,
      columns: validation.columns,
      progress: { current: 0, total: data.length, status: 'pending' }
    });

    // 返回预览数据和列名
    const preview = data.slice(0, 5).map(row => {
      const obj = {};
      validation.columns.forEach(col => {
        let value = row[col];
        // 对敏感信息进行脱敏
        if (col.includes('密码') || col.toLowerCase().includes('password')) {
          obj[col] = '****';
        } else if (col.includes('卡号') || col.toLowerCase().includes('card')) {
          const strVal = String(value || '');
          obj[col] = strVal.length > 4 ? strVal.substring(0, 4) + '****' : strVal;
        } else {
          obj[col] = value;
        }
      });
      return obj;
    });

    res.json({
      success: true,
      sessionId,
      rowCount: data.length,
      columns: validation.columns,
      emailColumn: validation.emailColumn,
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
  const { data, filePath, emailColumn } = session;

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
    const email = row[emailColumn];

    try {
      const subject = replaceVariables(emailTemplate.subject, row);
      const body = replaceVariables(emailTemplate.body, row);

      await transporter.sendMail({
        from: `"${smtpConfig.senderName}" <${smtpConfig.user}>`,
        to: String(email),
        subject,
        text: body
      });

      session.progress.results.push({
        email: String(email),
        success: true
      });
    } catch (error) {
      session.progress.results.push({
        email: String(email),
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