const ipc = require('electron').ipcRenderer;

// 输入文件
const inputButton = document.getElementById('input-button');
inputButton.addEventListener('click', function() {
    ipc.send('open-input-file');
});

ipc.on('input-file-path', function (event, path) {
    const inputSource = document.getElementById('inputSource');
    inputSource.value = path;
});

// 输出文件
const outputButton = document.getElementById('output-button');
outputButton.addEventListener('click', function() {
    ipc.send('open-output-directory');
});

ipc.on('output-file-path', function (event, path) {
    const outputSource = document.getElementById('outputSource');
    outputSource.value = path;
});

