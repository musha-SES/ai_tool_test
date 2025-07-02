const { contextBridge } = require('electron');

// レンダラープロセスでNode.jsのAPIを直接公開しない場合、
// ここでは特に何もする必要はありません。
// 必要に応じて、IPC通信などを設定できます。

// 例: レンダラープロセスからメインプロセスにメッセージを送る場合
// contextBridge.exposeInMainWorld('electronAPI', {
//   sendMessage: (message) => ipcRenderer.send('message-from-renderer', message)
// });
