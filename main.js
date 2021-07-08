// for excel
var XLSX = require('xlsx')
// Modules to control application life and create native browser window
const { spawn } = require('child_process')

try {
  require('electron-reloader')(module)
} catch (_) {}

const { app, BrowserWindow, ipcMain, dialog } = require('electron')
const path = require('path')
const { log } = require('console')
function createWindow() {
  // Create the browser window.
  const mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntergration: true
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
ipcMain.on('destination', (event, args) => {
  let result = dialog.showOpenDialog({
    properties: ['openDirectory']
  })
  event.sender.send('result', 'selecting directory...')

  result.then((value) => {
    loadEvent(event)

    if (!value.canceled) {

      event.sender.send('selected-destination', value.filePaths)
    }

  })
})

ipcMain.on('extract', (event, arg) => {

  console.log(arg[2]);

  // first read fie
  let sheet = XLSX.readFile(arg[0], {  sheets: [arg[2]], cellStyles:true, cellFormula:false })


  // creating a new workbok
  var wb = XLSX.utils.book_new()

  // get the workbook data
  let data = sheet.Sheets[arg[2]]

  
  // append sheet to new workbook
  XLSX.utils.book_append_sheet(wb, data, arg[2]);

  // write to a new file
  XLSX.writeFile(wb, arg[1])

  // send message 
  event.sender.send('result', 'done')
})


ipcMain.on('open-file-dialog', (event, arg) => {




  let result = dialog.showOpenDialog({
    properties: ['openFile']
  })
  result.then((value) => {

    if (!value.canceled) {

      loadEvent(event)





      var workbook = XLSX.readFile(value.filePaths[0]);
      var sheet_name_list = workbook.SheetNames;
      event.sender.send('sheets', sheet_name_list)



      event.sender.send('selected-directory', value.filePaths)


    }

  })

})

function loadEvent(event) {

  event.sender.send('loading')
}

