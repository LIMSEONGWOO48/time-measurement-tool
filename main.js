const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const { createPDFWindow } = require('./pdf-utils');
const { spawn } = require('child_process');

// Disable GPU Acceleration
app.disableHardwareAcceleration();

function createWindow() {
  const mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    icon: path.join(__dirname, 'icon.ico'),
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntegration: true,
      contextIsolation: false
    }
  });

  mainWindow.loadFile('index.html');
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', function () {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') app.quit();
});

ipcMain.on('select-file', async (event) => {
  try {
    const result = await dialog.showOpenDialog({
      properties: ['openFile'],
      filters: [
        { name: 'CSV Files', extensions: ['csv'] }
      ]
    });

    if (!result.canceled) {
      event.sender.send('file-selected', result.filePaths[0]);
    }
  } catch (error) {
    console.error('Error selecting file:', error);
    event.reply('file-selection-error', 'Failed to select file.');
  }
});

ipcMain.on('process-files', (event, args) => {
  const { filePath, standardTimesFile } = args;

  // Pythonバイナリのパスを設定する
  const binaryPath = path.join(__dirname, 'dist', 'process_files'); // PyInstallerで作成されたバイナリファイル

  // Log paths for debugging
  console.log('Python script path:', binaryPath);
  console.log('CSV file path:', filePath);
  console.log('Standard times file path:', standardTimesFile);

  // Spawn process to run the script
  const pyProcess = spawn(binaryPath, [filePath, standardTimesFile]);

  let outputData = '';
  let errorData = '';

  pyProcess.stdout.on('data', (data) => {
    outputData += data.toString();
  });

  pyProcess.stderr.on('data', (data) => {
    errorData += data.toString();
  });

  pyProcess.on('close', (code) => {
    if (code === 0) {
      try {
        const result = JSON.parse(outputData);
        event.reply('process-result', result);
      } catch (e) {
        console.error('Error parsing Python output:', e);
        event.reply('process-error', 'Failed to parse Python output.');
      }
    } else {
      console.error(`Python script exited with code ${code}:`, errorData);
      event.reply('process-error', `Python script exited with code ${code}: ${errorData}`);
    }
  });
});

ipcMain.on('save-pdf', async (event, args) => {
  const { name, details, totalStandardTime } = args;
  try {
    const pdfWindow = createPDFWindow(name, details, totalStandardTime);

    pdfWindow.webContents.on('did-finish-load', async () => {
      try {
        const { filePath } = await dialog.showSaveDialog({
          title: 'Save PDF',
          defaultPath: `${name}_certificate.pdf`,
          filters: [{ name: 'PDF Files', extensions: ['pdf'] }]
        });

        if (filePath) {
          const data = await pdfWindow.webContents.printToPDF({});
          fs.writeFileSync(filePath, data);
          event.reply('pdf-saved', filePath);
        } else {
          event.reply('pdf-error', 'PDF save canceled');
        }
      } catch (printError) {
        console.error('Error during PDF generation:', printError);
        event.reply('pdf-error', printError.message);
      } finally {
        pdfWindow.close();
      }
    });
  } catch (error) {
    console.error('Error creating PDF window:', error);
    event.reply('pdf-error', error.message);
  }
});

// save-csv handler (冗長性を排除)
ipcMain.on('save-csv', async (event, csvData) => {
  try {
    const { filePath } = await dialog.showSaveDialog({
      title: 'Save CSV',
      defaultPath: 'output_result.csv',
      filters: [
        { name: 'CSV Files', extensions: ['csv'] }
      ]
    });

    if (filePath) {
      // Add BOM for UTF-8 to prevent character encoding issues
      const bom = '\uFEFF';
      fs.writeFileSync(filePath, bom + csvData, 'utf-8');
      event.reply('csv-saved', filePath);
    } else {
      event.reply('csv-error', 'CSV save canceled');
    }
  } catch (error) {
    console.error('Error during CSV save:', error);
    event.reply('csv-error', error.message);
  }
});
