const {app, BrowserWindow, Menu} = require('electron');
const path = require('path');
const url = require('url');
  
// Keep a global reference of the window object, if you don't, the window will
// be closed automatically when the JavaScript object is garbage collected.
let win;

function createWindow () {
  // Create the browser window.
  win = new BrowserWindow({icon: __dirname + '/images/256x256.png'});

  win.maximize();

  win.on('focus', setMenu);
  win.once('show', setMenu);

  // and load the index.html of the app.
  win.loadURL(url.format({
    pathname: path.join(__dirname, 'index.html'),
    protocol: 'file:',
    slashes: true
  }));

  // Emitted when the window is closed.
  win.on('closed', () => {
    // Dereference the window object, usually you would store windows
    // in an array if your app supports multi windows, this is the time
    // when you should delete the corresponding element.
    win = null;
  })
}

// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.
app.on('ready', createWindow);

// Quit when all windows are closed.
app.on('window-all-closed', () => {
  app.quit();
});


function setMenu() {
  Menu.setApplicationMenu(Menu.buildFromTemplate([
    {
      label: "Edit",
      submenu: [
        { label: 'Undo', accelerator: 'CmdOrCtrl+Z', role: 'undo' },
        { label: 'Redo', accelerator: 'CmdOrCtrl+Shift+Z', role: 'redo' },
        { type: 'separator' },
        { label: 'Cut', accelerator: 'CmdOrCtrl+X', role: 'cut' },
        { label: 'Copy', accelerator: 'CmdOrCtrl+C', role: 'copy' },
        { label: 'Paste', accelerator: 'CmdOrCtrl+V', role: 'paste' },
        { label: 'Select All', accelerator: 'CmdOrCtrl+A', role: 'selectall' }
      ]
    },
    {
      label: "View",
      submenu: [
        {
          label: 'Reload',
          accelerator: 'CmdOrCtrl+R',
          click() { win.reload(); }
        },
        { type: 'separator' },
        {
          label: 'Toggle Dev Tools',
          accelerator: 'CmdOrCtrl+Alt+I',
          click() {
            var focusedWindow = BrowserWindow.getFocusedWindow();
            focusedWindow && focusedWindow.webContents.toggleDevTools();
          }
        },
        { type: 'separator' },
        {
          label: 'Toggle Fullscreen',
          role: 'togglefullscreen',
          click: function() {
            win.setFullScreen(!win.isFullScreen());
          }
        }
      ]
    }
  ]));
}
