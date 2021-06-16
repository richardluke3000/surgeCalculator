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

  //for opening destination

  const destination = document.getElementById('destination_btn')

  destination.addEventListener('click', (event) =>{
    ipcRenderer.send('destination')
  })
  
  selectDirBtn.addEventListener('click', (event) => {
    ipcRenderer.send('open-file-dialog')
    ipcRenderer.send('asynchronous-message')
  })

  ipcRenderer.on('selected-destination', (event, path)=>{
    let destination = document.getElementById('destination')
    console.log(path.toString());
    destination.innerText = path.toString()
  })
  ipcRenderer.on('selected-directory', (event, path) => {
    console.log(path);
    console.log(event);
    
    document.getElementById('selected-file').innerHTML = `${path}`
    let rename = document.getElementById('rename')
    console.log(rename);
    console.log(path.toString());
        rename.value = path.toString()
        console.log(rename.value);
  })
  ipcRenderer.on('sheets', (event, data) => {
   let list = data.split(',')
      let sheets = document.getElementById('select');
      list.forEach(item => {
        
        sheets.innerHTML += (`<option value='0'>${item}</option>`)
        // console.log(sheets);
        loading.innerHTML = ''
        
      });
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





