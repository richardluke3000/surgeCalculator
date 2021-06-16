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


  // open file
  const selectDirBtn = document.getElementById('select-directory')

  

  selectDirBtn.addEventListener('click', (event) => {
    ipcRenderer.send('open-file-dialog')
    ipcRenderer.send('asynchronous-message')
  })

  ipcRenderer.on('selected-directory', (event, path) => {
    console.log(path);
    console.log(event);
    document.getElementById('selected-file').innerHTML = `${path}`
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





