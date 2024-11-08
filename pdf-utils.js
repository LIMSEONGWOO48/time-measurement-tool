const { BrowserWindow } = require('electron');
const fs = require('fs');
const path = require('path');

function mergeDuplicateContent(details) {
  const mergedDetails = {};

  details.forEach(detail => {
    const key = `${detail.contentName}-${detail.folderName}`;
    if (mergedDetails[key]) {
      mergedDetails[key].startTime = detail.startTime; // Update to latest start time if needed
      mergedDetails[key].endTime = detail.endTime; // Update to latest end time if needed
      mergedDetails[key].duration = detail.duration; // Adjust this as needed (e.g., sum durations)
      mergedDetails[key].standardTime = detail.standardTime; // Adjust this as needed (e.g., sum or keep the max)
    } else {
      mergedDetails[key] = { ...detail };
    }
  });

  return Object.values(mergedDetails);
}

function createPDFWindow(name, details, totalStandardTime) {
  const certificatePath = path.join(__dirname, 'certificate_template.html');
  let certificateContent = fs.readFileSync(certificatePath, 'utf-8');

  const mergedDetails = mergeDuplicateContent(details);

  let totalDuration = 0;
  let latestCompletionDate = null;

  let tableRows = mergedDetails
    .map(detail => {
      totalDuration += timeToSeconds(detail.duration);
      
      // 日付が有効かどうかを確認して最新の日付を取得
      const endTime = new Date(detail.endTime);
      if (!isNaN(endTime.getTime()) && (!latestCompletionDate || endTime > latestCompletionDate)) {
        latestCompletionDate = endTime; // 有効な場合に最新の日付を設定
      }

      return `
        <tr>
          <td>${detail.folderName}</td>
          <td>${detail.contentName}</td>
          <td>${detail.startTime}</td>
          <td>${detail.endTime}</td>
          <td>${detail.duration}</td>
          <td>${detail.standardTime}</td>
        </tr>
      `;
    }).join('');

  // 最新の学習完了日を発行日として設定
  let issueDate = latestCompletionDate instanceof Date && !isNaN(latestCompletionDate)
    ? latestCompletionDate.toISOString().split('T')[0]
    : 'N/A';

  certificateContent = certificateContent.replace('<!-- ここに証明書の内容を追加 -->', `
    <h1>修了証</h1>
    <p>${name} 殿</p>
    <p>あなたは、所定の課程を修了されたことをここに証します。</p>
    <p>発行日　${issueDate}</p>
    <p>発行元団体　本気ファクトリー株式会社</p>
    <p>発行者名　代表取締役　畠山和也</p>
    <table>
      <thead>
        <tr>
          <th>フォルダ名</th>
          <th>コンテンツ名</th>
          <th>学習開始日</th>
          <th>学習完了日</th>
          <th>所要時間</th>
          <th>標準学習時間</th>
        </tr>
      </thead>
      <tbody>
        ${tableRows}
      </tbody>
    </table>
    <p>所要時間の合計: ${secondsToTime(totalDuration)}</p>
    <p>標準学習時間の合計: ${secondsToTime(totalStandardTime)}</p>
  `);

  const pdfWindow = new BrowserWindow({ show: false });
  pdfWindow.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(certificateContent)}`);

  return pdfWindow;
}

// Helper functions to convert time
function timeToSeconds(timeStr) {
  const parts = timeStr.split(':');
  return (+parts[0]) * 3600 + (+parts[1]) * 60 + (+parts[2]);
}

function secondsToTime(seconds) {
  const h = Math.floor(seconds / 3600);
  const m = Math.floor((seconds % 3600) / 60);
  const s = seconds % 60;
  return `${h.toString().padStart(2, '0')}:${m.toString().padStart(2, '0')}:${s.toString().padStart(2, '0')}`;
}

module.exports = { createPDFWindow };
