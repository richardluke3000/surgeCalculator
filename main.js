// for excel
var XLSX = require('xlsx')

const {BrowserWindow, Menu, app, shell, dialog, ipcMain} = require('electron')

// MENU TEMPLATE
let template = [{
  label: 'Links',
  click: () => {
    app.emit('activate')
  }
}]



var fs = require('fs')
// Modules to control application life and create native browser window
const { spawn } = require('child_process')



const path = require('path')


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


// WHEN APP IS READY
app.on('ready', () => {
  // const menu = Menu.buildFromTemplate(template)
  // Menu.setApplicationMenu(menu)
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
  let sheet = XLSX.readFile(arg[0], {  sheets: [arg[2]] })


  // creating a new workbok
  var wb = XLSX.utils.book_new()

  // get the workbook data
  let data = sheet.Sheets[arg[2]]

  XLSX.utils.book_append_sheet(wb, data, arg[2])

  // write to new file
  XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });

  XLSX.writeFile(wb, `${arg[1]}`);

  event.sender.send('result', 'finished Extracting sheet')


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

ipcMain.on('filesFolder', (event, arg)=>{
  let result = dialog.showOpenDialog({
    properties: ['openDirectory']
  })

  result.then((value) => {
    
    if (!value.canceled) {

      loadEvent(event)

      console.log(value.filePaths);


      //variable to contain files in the directory
      let newFiles = [];


    fs.readdir( value.filePaths[0], function(err, files){
      
      if (err) {
        // do stuff to handle the error
        
        return event.sender.send('result', 'unable to do necessary stuff : ' + err)

      }
      let patt = new RegExp("^[A-z].*.xlsx")

      

      files.forEach(file => {

        const filtered = patt.exec(file)


        if( filtered ){
          newFiles.push( filtered[0] )
        }
        
      });
      event.sender.send('selected-consolidate', value.filePaths, newFiles)
      event.sender.send('result', `Found ${newFiles.length} files`)
    } )
      
    }

  })
})

ipcMain.on('consolidate', (event ,  dir, files , destination, startDate, sheetToConsolidate )=>{
  
  let arg = dir

  // collumn to search for a date
  let dateCollumn;

  loadEvent(event)


  let finalsheet = []

  files.forEach((element, index) => {
     

    let path = `${arg}\\${element}`

    if( sheetToConsolidate === 'Defaulter Q1' ){
      dateCollumn = 'F'
    }else{
      dateCollumn = 'E'
    }

    var sheet = XLSX.readFile( path, {sheets: sheetToConsolidate } );
   

    let data = sheet.Sheets[sheetToConsolidate]
  

    XLSX.utils.sheet_to_json(data)

    let check = Date.parse(startDate).toString().substr(0,5)
    
    let taken = [] ;
   
    // this will check every key / or cell in the sheeet
    for (const key in data) {
     

      if( key[0] == dateCollumn && Date.parse(data[key].w).toString().substr(0,5) ==  check ){
        let row = key.substr(1,100);
        
        
          for (const key in data) {
            
            if( row == key.substr(1,100)   )
              taken.push( data[key].v )
            
          }

      }
      if ( taken.length > 0  ){
        finalsheet.push(taken)
        taken = []
      }
    }

  });
  
  finalsheet = XLSX.utils.aoa_to_sheet(finalsheet);

  let wb = XLSX.utils.book_new();

  wb.Props = {
      Title: "surge cascade",
      Subject: "Output",
      Author: "UI-HEXED",
      CreatedDate: new Date(),
  }

  XLSX.utils.book_append_sheet(wb, finalsheet, sheetToConsolidate)

  XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });

  XLSX.writeFile(wb, destination);
  
  event.sender.send('result', 'finished consolidating')


})


function loadEvent(event) {

  event.sender.send('loading')
}

