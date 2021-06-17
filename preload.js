// All of the Node.js APIs are available in the preload process.
// It has the same sandbox as a Chrome extension.
var { ipcRenderer } = require('electron')

window.addEventListener('DOMContentLoaded', () => {
  const replaceText = (selector, text) => {
    const element = document.getElementById(selector)
    if (element) element.innerText = text
  }

  for (const type of ['chrome', 'node', 'electron']) {
    replaceText(`${type}-version`, process.versions[type])
  }

  Selected_file = 'path'
  // for open file
  const selectDirBtn = document.getElementById('select-directory')

  // for extracting
  const extract = document.getElementById('extract')

  //for opening destination

  const destination = document.getElementById('destination_btn')

  // request destination folder
  destination.addEventListener('click', (event) =>{
    ipcRenderer.send('destination')
  })

  // request extract action
  extract.addEventListener('click', (event)=>{
    let source = document.getElementById('selected-file').innerText
    
    let destination = document.getElementById('destination').innerText
    
    let sheet_index = document.getElementById('select').value
    
    ipcRenderer.send('extract', [source, destination, sheet_index])
  })
  

  selectDirBtn.addEventListener('click', (event) => {
    ipcRenderer.send('open-file-dialog')
    ipcRenderer.send('asynchronous-message')
  })


  // when destination folder is selected
  ipcRenderer.on('selected-destination', (event, path)=>{
    let destination = document.getElementById('destination')
    let rename = document.getElementById('rename')
    
    destination.innerText = path.toString()
    destination.innerText += "\\" + rename.value
  })

  // when the directory is selected
  ipcRenderer.on('selected-directory', (event, path) => {
    
    let selected_file = document.getElementById('selected-file')
    selected_file.innerText = `${path}`

    let  file = selected_file.innerText.split("\\")
    let rename = document.getElementById('rename')

    rename.value = file[file.length - 1 ]

  })

  // when the sheets have been retrieved
  ipcRenderer.on('sheets', (event, data) => {
   let list = data.split(',') //split the sheets into various sheets
      let sheets = document.getElementById('select');
      let x = 0
      let loading = document.getElementById('loading');

      loading.innerHTML = 'please a wait moment...'

      // for each sheet add to the drop down list
      list.forEach(item => {
        sheets.innerHTML += (`<option value='${x}'>${item}</option>`)
        
        x++
      });
      loading.innerHTML = ''
  })


  ipcRenderer.on('asynchronous-reply', (event, arg) => {
    const message = `Asynchronous message reply: ${arg}`
    console.log(message);
  })
  
  ipcRenderer.on('loading', (event, arg) => {
    let loading = document.getElementById('loading');

    loading.innerHTML = 'please a wait moment...'
    
  })




})





