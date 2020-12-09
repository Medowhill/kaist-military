const { ipcRenderer } = require('electron');

function log(x) {
  document.getElementById('p').innerHTML = x;
}

document.getElementById('button-file1').addEventListener('click', e => {
  ipcRenderer.send('asynchronous-message', { typ: 1 });
});

document.getElementById('button-file2').addEventListener('click', e => {
  ipcRenderer.send('asynchronous-message', { typ: 2 });
});

document.getElementById('button-dir').addEventListener('click', e => {
  ipcRenderer.send('asynchronous-message', { typ: 3 });
});

document.getElementById('button-run').addEventListener('click', e => {
  const file1 = document.getElementById('p-file1').innerHTML;
  const file2 = document.getElementById('p-file2').innerHTML;
  const dir = document.getElementById('p-dir').innerHTML;
  const noFile = '선택된 파일 없음';
  const noDir = '선택된 폴더 없음';
  if (file1 !== noFile && file2 !== noFile && dir !== noDir)
    ipcRenderer.send('asynchronous-message', { typ: 4, file1, file2, dir });
});

ipcRenderer.on('asynchronous-reply', (event, arg) => {
  switch (arg.typ) {
    case 1:
      document.getElementById('p-file1').innerHTML = arg.file;
      break;
    case 2:
      document.getElementById('p-file2').innerHTML = arg.file;
      break;
    case 3:
      document.getElementById('p-dir').innerHTML = arg.file;
      break;
  }
});

