const { app, BrowserWindow } = require('electron');
const path = require('path');
const { spawn } = require('child_process');

let serverProcess;

function createWindow() {
  const win = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      nodeIntegration: false, // レンダラープロセスでのNode.js統合を無効化
      contextIsolation: true, // プリロードスクリプトを使用する場合にtrue
      preload: path.join(__dirname, 'preload.js') // 必要に応じてプリロードスクリプトを指定
    }
  });

  win.loadFile('index.html');

  // 開発ツールを開く
  win.webContents.openDevTools();
}

function startServer() {
  // server.jsを子プロセスとして起動
  serverProcess = spawn('node', [path.join(__dirname, 'server.js')], {
    stdio: 'inherit', // 親プロセスの標準入出力に接続
    env: { ...process.env, PORT: 3000 } // 環境変数を引き継ぎ、PORTを設定
  });

  serverProcess.on('close', (code) => {
    console.log(`server.js process exited with code ${code}`);
  });

  serverProcess.on('error', (err) => {
    console.error('Failed to start server.js process:', err);
  });
}

function stopServer() {
  if (serverProcess) {
    console.log('Stopping server.js process...');
    // SIGTERMを送信してgraceful shutdownを促す
    serverProcess.kill('SIGTERM');
  }
}

app.whenReady().then(() => {
  startServer(); // Electronアプリ起動時にサーバーを起動
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

app.on('will-quit', () => {
  stopServer(); // Electronアプリ終了時にサーバーを停止
});
