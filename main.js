const { app, BrowserWindow, shell, dialog } = require('electron');
const path = require('path');

const PORT = 3000;
let mainWindow = null;

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

  // 在默认浏览器中打开外部链接
  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    if (url.startsWith('http')) {
      shell.openExternal(url);
    }
    return { action: 'deny' };
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });

  // 显示错误
  mainWindow.webContents.on('did-fail-load', (event, errorCode, errorDescription) => {
    dialog.showErrorBox('加载失败', `无法加载页面: ${errorDescription}\n请检查服务器是否正常启动`);
  });
}

// 应用启动
app.whenReady().then(() => {
  try {
    // 启动后端服务器
    require('./server.js');
    console.log('服务器已启动');

    // 等待服务器启动完成后再创建窗口
    setTimeout(() => {
      createWindow();
    }, 2000);

    app.on('activate', () => {
      if (BrowserWindow.getAllWindows().length === 0) {
        createWindow();
      }
    });
  } catch (error) {
    dialog.showErrorBox('启动失败', error.message);
    app.quit();
  }
});

// 所有窗口关闭时退出应用
app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

// 捕获未处理的错误
process.on('uncaughtException', (error) => {
  dialog.showErrorBox('程序错误', error.message);
});