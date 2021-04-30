const { app, BrowserWindow, dialog } = require('electron');

function createWindow () {
  const win = new BrowserWindow({
    width: 800,
    height: 900,
    webPreferences: {
      nodeIntegration: true,
      enableRemoteModule: true
    },
  })
  win.webContents.openDevTools()
  win.loadFile('index.html')
}

app.whenReady().then(createWindow)

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit()
  }
})

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) {
    createWindow()
  }
})
