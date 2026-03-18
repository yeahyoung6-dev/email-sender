const { app, BrowserWindow, shell, dialog } = require('electron');
const path = require('path');
const fs = require('fs');

const PORT = 3000;
let mainWindow = null;

// 获取正确的基础路径
function getBasePath() {
  // 打包后使用 app.getAppPath()，开发时使用 __dirname
  return app.isPackaged ? path.dirname(app.getAppPath()) : __dirname;
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    minWidth: 900,
    minHeight: 600,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js')
    },
    title: '邮件发送工具',
    show: false
  });

  mainWindow.loadURL(`http://localhost:${PORT}`);

  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
  });

  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    if (url.startsWith('http')) {
      shell.openExternal(url);
    }
    return { action: 'deny' };
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });

  mainWindow.webContents.on('did-fail-load', (event, errorCode, errorDescription) => {
    dialog.showErrorBox('加载失败', `无法加载页面: ${errorDescription}`);
  });
}

app.whenReady().then(() => {
  try {
    // 设置工作目录为应用所在目录
    const basePath = getBasePath();
    process.chdir(basePath);

    // 确保uploads目录存在
    const uploadsDir = path.join(basePath, 'uploads');
    if (!fs.existsSync(uploadsDir)) {
      fs.mkdirSync(uploadsDir, { recursive: true });
    }

    // 启动后端服务器
    require('./server.js');

    // 等待服务器启动
    setTimeout(() => {
      createWindow();
    }, 2000);

    app.on('activate', () => {
      if (BrowserWindow.getAllWindows().length === 0) {
        createWindow();
      }
    });
  } catch (error) {
    dialog.showErrorBox('启动失败', error.message + '\n' + error.stack);
    app.quit();
  }
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

process.on('uncaughtException', (error) => {
  dialog.showErrorBox('程序错误', error.message);
});