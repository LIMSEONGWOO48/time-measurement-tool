const { ipcRenderer } = require('electron');
const path = require('path');
const fs = require('fs');

let originalData = [];
let selectedFilePath = '';
let selectedStandardTimesFile = '';
let totalStandardTime = 0;

document.getElementById('select-file').addEventListener('click', () => {
  ipcRenderer.send('select-file');
});

document.getElementById('process-files').addEventListener('click', () => {
  const standardTimesFileName = document.getElementById('standard-times-select').value;
  selectedStandardTimesFile = path.join(__dirname, standardTimesFileName);
  ipcRenderer.send('process-files', { filePath: selectedFilePath, standardTimesFile: selectedStandardTimesFile });
});

ipcRenderer.on('file-selected', (event, filePath) => {
  selectedFilePath = filePath;
  document.getElementById('process-files').disabled = false;
  const fileName = filePath.split('\\').pop().split('/').pop();
  document.getElementById('selected-file-name').innerText = `選択されたファイル: ${fileName}`;
});

ipcRenderer.on('process-result', (event, result) => {
  originalData = result.data;
  totalStandardTime = result.totalStandardTime;
  alert('学習時間計測ができました。');

  document.getElementById('name-filter').value = '';
  populateFilters(originalData);
  filterAndDisplayResults();
});

document.getElementById('name-filter').addEventListener('change', filterAndDisplayResults);
document.getElementById('mark-filter').addEventListener('change', filterAndDisplayResults);

document.getElementById('export-csv').addEventListener('click', () => {
  const tableElement = document.getElementById('result-table');
  const csvData = tableToCsv(tableElement);
  ipcRenderer.send('save-csv', csvData);
});

document.getElementById('save-pdf').addEventListener('click', () => {
  const nameFilterValue = document.getElementById('name-filter').value;
  if (!nameFilterValue) {
    alert('氏名を選択してください');
    return;
  }
  const details = extractDetailsForName(nameFilterValue, originalData);
  const detailsWithoutCompletionCertificate = details.filter(detail => !detail.contentName.includes('修了証'));
  ipcRenderer.send('save-pdf', { name: nameFilterValue, details: detailsWithoutCompletionCertificate, totalStandardTime });
});

function populateFilters(data) {
  const names = new Set(data.map(row => row['氏名']));

  const nameFilter = document.getElementById('name-filter');
  nameFilter.innerHTML = '<option value="">全員</option>'; // Clear existing options
  names.forEach(name => {
    const option = document.createElement('option');
    option.value = name;
    option.textContent = name;
    nameFilter.appendChild(option);
  });
}

function filterAndDisplayResults() {
  const nameFilterValue = document.getElementById('name-filter').value;
  const markFilterValue = document.getElementById('mark-filter').value;

  let filteredData = originalData;
  if (nameFilterValue) {
    filteredData = filteredData.filter(row => row['氏名'] === nameFilterValue);
  }
  if (markFilterValue) {
    filteredData = filteredData.filter(row => row['マーク'] === markFilterValue);
  }

  const processedData = processDuplicateContentNames(filteredData);
  document.getElementById('result-table').querySelector('tbody').innerHTML = jsonToTable(processedData);
}

function processDuplicateContentNames(data) {
  const groupedData = {};

  data.forEach(row => {
    const contentName = row['コンテンツ名'];
    const personName = row['氏名'];
    const key = `${contentName}-${personName}`; // Group by both content name and person name
    if (!groupedData[key]) {
      groupedData[key] = [];
    }
    groupedData[key].push(row);
  });

  const processedRows = [];

  Object.values(groupedData).forEach(group => {
    let totalDuration = group.reduce((sum, row) => sum + timeToSeconds(row['所要時間']), 0);
    console.log(`Total Duration (seconds): ${totalDuration}`); // デバッグ用ログ
    const standardTime = timeToSeconds(group[0]['標準学習時間']);
    console.log(`Standard Time (seconds): ${standardTime}`); // デバッグ用ログ
    
    group[0]['所要時間'] = secondsToTime(totalDuration); // 最初のエントリに合計時間を設定
    
    const contentName = group[0]['コンテンツ名'];
    const mark = contentName.includes('理解度テスト') ? 'O' : (totalDuration >= standardTime ? 'O' : 'X');
    group[0]['マーク'] = mark;

    // 足りない時間を計算
    const missingTime = Math.max(0, standardTime - totalDuration);
    group[0]['足りない時間'] = mark === 'X' ? secondsToTime(missingTime) : '';

    processedRows.push(group[0]);
  });

  return processedRows;
}

function jsonToTable(data) {
  const tableBody = document.createElement('tbody');
  
  data.forEach((row, rowIndex) => {
    const tr = document.createElement('tr');
    if (row['コンテンツ名'] && row['コンテンツ名'].includes('修了証')) {
      tr.classList.add('completion-certificate-row');
    }
    Object.values(row).forEach((cell, cellIndex) => {
      const td = document.createElement('td');
      td.textContent = cell;
      if (cellIndex === 12) { // Assuming 'URL' is the 13th column
        td.classList.add('url-cell');
      }
      tr.appendChild(td);
    });
    tableBody.appendChild(tr);
  });

  return tableBody.innerHTML;
}

function tableToCsv(tableElement) {
  const headers = Array.from(tableElement.querySelectorAll('thead th')).map(th => th.innerText);
  const rows = Array.from(tableElement.querySelectorAll('tbody tr'));
  let csv = headers.join(',') + '\n';
  rows.forEach(row => {
    const cells = Array.from(row.querySelectorAll('td')).map(td => td.innerText);
    csv += cells.join(',') + '\n';
  });
  return csv.trim();
}

ipcRenderer.on('pdf-saved', (event, filePath) => {
    console.log(`PDF saved to ${filePath}`);
    alert('PDFが保存されました');
  });
  
  ipcRenderer.on('pdf-error', (event, errorMessage) => {
    console.error(`PDF save error: ${errorMessage}`);
    alert('PDFの保存中にエラーが発生しました');
  });
  
  ipcRenderer.on('csv-saved', (event, filePath) => {
    console.log(`CSV saved to ${filePath}`);
    alert('CSVが保存されました');
  });
  
  ipcRenderer.on('csv-error', (event, errorMessage) => {
    console.error(`CSV save error: ${errorMessage}`);
    alert('CSVの保存中にエラーが発生しました');
  });

function extractDetailsForName(name, data) {
  return data.filter(row => row['氏名'] === name).map(row => ({
    folderName: row['フォルダ名'],
    contentName: row['コンテンツ名'],
    startTime: row['学習開始日時'],
    endTime: row['学習完了日'],
    duration: row['所要時間'],
    standardTime: row['標準学習時間']
  }));
}

// Helper functions to convert time
function timeToSeconds(timeStr) {
  const parts = timeStr.split(':');
  const hours = +parts[0];
  const minutes = +parts[1];
  const seconds = +parts[2];
  return hours * 3600 + minutes * 60 + seconds;
}

function secondsToTime(seconds) {
  const h = Math.floor(seconds / 3600);
  const m = Math.floor((seconds % 3600) / 60);
  const s = seconds % 60;
  return `${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}:${s.toString().padStart(2, '0')}`;
}
