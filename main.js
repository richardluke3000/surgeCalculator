// Modules to control application life and create native browser window
const { spawn } = require('child_process')
const {app, BrowserWindow, ipcMain, dialog} = require('electron')
const path = require('path')
function createWindow () {
  // Create the browser window.
  const mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js')
    }
  })

  // and load the index.html of the app.
  mainWindow.loadFile('index.html')

  // Open the DevTools.
  // mainWindow.webContents.openDevTools()
}

// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.
app.whenReady().then(() => {
  createWindow()
  
  app.on('activate', function () {
    // On macOS it's common to re-create a window in the app when the
    // dock icon is clicked and there are no other windows open.
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})

// Quit when all windows are closed, except on macOS. There, it's common
// for applications and their menu bar to stay active until the user quits
// explicitly with Cmd + Q.
app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') app.quit()
})

// In this file you can include the rest of your app's specific main process
// code. You can also put them in separate files and require them here.

// save dialog 
ipcMain.on('destination', (event, args)=>{
  let result = dialog.showOpenDialog({
    properties:['openDirectory']
  })
  result.then((value)=>{
    if(!value.canceled){
      console.log(value);
      event.sender.send('selected-destination', value.filePaths)
    }
  })
} )

ipcMain.on('extract', (event, arg)=>{
  extract = spawn('python',['sheet.py', arg[0], arg[1], arg[2] ] )

  console.log(typeof arg);
  console.log('extracting' + arg[0] +  arg[1] + arg [2] );

  extract.stdout.on('data', data => { console.log(data); event.sender.send('result', data.toString() )} )

  
  

})


ipcMain.on('open-file-dialog', (event,arg) => {
  console.log('files opened');

  let result = dialog.showOpenDialog({
    properties: ['openFile']
  })
  result.then( (value)=>{
    
      if (!value.canceled){
        
        loadEvent(event)
        
        let py = spawn('python',['main.py', value.filePaths] )

        py.stdout.on('data', data => event.sender.send('sheets', data.toString() ) )


        event.sender.send('selected-directory', value.filePaths)


      }
    
  } )

} )

function loadEvent(event) {
  console.log('loading');
  event.sender.send('loading')
}
