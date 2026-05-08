const { app, BrowserWindow, dialog, ipcMain } = require('electron');
const XlsxPopulate = require('xlsx-populate');

let mainWindow;
const RESPONSE_COUNT_PER_SEC = 10;
let isProcessing = false;

// Функция для отправки логов в окно
function sendLog(message) {
    if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.webContents.send('log-message', message);
    }
    console.log(message); // Также выводим в консоль для отладки
}

// Функция проверки, открыт ли файл
function isFileLocked(filePath) {
    try {
        const fs = require('fs');
        const fd = fs.openSync(filePath, 'r+');
        fs.closeSync(fd);
        return false;
    } catch (error) {
        if (error.code === 'EBUSY') {
            return true;
        }
        return false;
    }
}

// ========== FUNCTIONS ==========

async function getData(phone) {
    try {
        const response = await fetch(`http://num.voxlink.ru/get/?num=${phone}&field=region`);
        if (!response.ok) {
            throw new Error('Failed to get region');
        }
        return await response.text();
    } catch (err) {
        return err.message;
    }
}

async function timeoutFetch(phones) {
    const formattedPhones = phones.map(item => String(item));
    const groupedArr = [];

    while (formattedPhones.length > 0) {
        let temp = formattedPhones.splice(0, RESPONSE_COUNT_PER_SEC);
        groupedArr.push(temp);
    }

    const totalBatches = groupedArr.length;
    sendLog(`📦 Всего пачек для обработки: ${totalBatches}`);

    async function fetchGroup(i) {
        return await Promise.all(groupedArr[i].map(el => getData(el)));
    }

    const allResults = [];
    for (let i = 0; i < totalBatches; i++) {
        sendLog(`🔄 Обработка пачки ${i + 1}/${totalBatches}...`);
        const res = await fetchGroup(i);
        allResults.push(...res);

        if ((i + 1) % 10 === 0 || i === totalBatches - 1) {
            const percent = Math.round((i + 1) / totalBatches * 100);
            sendLog(`📈 Прогресс: ${i + 1}/${totalBatches} пачек (${percent}%)`);
        }

        if (i < totalBatches - 1) {
            await new Promise(resolve => setTimeout(resolve, 1000));
        }
    }

    sendLog(`✅ Все ${totalBatches} пачек обработано!`);
    return allResults;
}

function getNextColumnLetter(columnLetter) {
    let num = 0;
    for (let i = 0; i < columnLetter.length; i++) {
        num = num * 26 + (columnLetter.charCodeAt(i) - 64);
    }
    num++;

    let result = '';
    while (num > 0) {
        num--;
        result = String.fromCharCode(65 + (num % 26)) + result;
        num = Math.floor(num / 26);
    }
    return result;
}

async function addColumnToExcel(excelPath, newColumnData, columnName = 'Регион') {
    const workbook = await XlsxPopulate.fromFileAsync(excelPath);
    const sheet = workbook.sheet(0);

    let lastColumnLetter = 'A';
    const usedRange = sheet.usedRange();

    if (usedRange) {
        const lastCell = usedRange.endCell();
        const lastColumnAddress = lastCell.address();
        const match = lastColumnAddress.match(/[A-Z]+/);
        if (match) {
            lastColumnLetter = match[0];
        }
    }

    const newColumnLetter = getNextColumnLetter(lastColumnLetter);
    sheet.cell(`${newColumnLetter}1`).value(columnName);

    for (let i = 0; i < newColumnData.length; i++) {
        const rowNumber = i + 2;
        sheet.cell(`${newColumnLetter}${rowNumber}`).value(newColumnData[i]);
    }

    await workbook.toFileAsync(excelPath);
    sendLog(`💾 Столбец "${columnName}" добавлен в файл`);
}

async function readPhonesFromExcel(filePath) {
    const workbook = await XlsxPopulate.fromFileAsync(filePath);
    const sheet = workbook.sheet(0);
    const usedRange = sheet.usedRange();

    if (!usedRange) return [];

    const rows = usedRange.value();
    const phones = [];

    for (let i = 0; i < rows.length; i++) {
        const cellValue = rows[i][0];
        if (cellValue) {
            const phoneStr = String(cellValue).replace(/[^0-9]/g, '');
            if (phoneStr.length >= 10) {
                phones.push(phoneStr);
            }
        }
    }
    return phones;
}

async function processExcelFile(filePath) {
    // Проверяем, не открыт ли файл
    if (isFileLocked(filePath)) {
        sendLog('❌ ОШИБКА: Файл открыт в Excel!');
        sendLog('📌 Пожалуйста, закройте Excel-файл и попробуйте снова.');
        return;
    }

    sendLog('📖 Читаем номера из Excel...');
    const phones = await readPhonesFromExcel(filePath);

    if (phones.length === 0) {
        sendLog('❌ Номера телефонов не найдены!');
        return;
    }

    sendLog(`📊 Найдено номеров: ${phones.length}`);
    sendLog(`⏱️ Обрабатываем по ${RESPONSE_COUNT_PER_SEC} номеров в секунду`);
    sendLog('🔄 Начинаем обработку...');
    sendLog('---');

    const regions = await timeoutFetch(phones);

    sendLog('---');
    sendLog('💾 Сохраняем результаты в Excel...');

    // Ещё раз проверяем перед сохранением
    if (isFileLocked(filePath)) {
        sendLog('❌ ОШИБКА: Файл был открыт во время обработки!');
        sendLog('📌 Пожалуйста, закройте Excel-файл и запустите снова.');
        return;
    }

    await addColumnToExcel(filePath, regions, 'Регион');

    sendLog('✅ ГОТОВО! Файл успешно обновлён.');
    sendLog(`📁 Расположение файла: ${filePath}`);
}

// ========== WINDOW ==========

function createWindow() {
    mainWindow = new BrowserWindow({
        width: 700,
        height: 600,
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false
        }
    });

    const html = `<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>Phone Region Checker</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Arial, sans-serif;
            display: flex;
            justify-content: center;
            align-items: center;
            min-height: 100vh;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        }
        .container {
            background: white;
            padding: 30px;
            border-radius: 15px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            width: 600px;
            text-align: center;
        }
        h1 {
            color: #333;
            margin-bottom: 10px;
        }
        p {
            color: #666;
            margin-bottom: 20px;
        }
        .warning {
            background: #fff3cd;
            border: 1px solid #ffc107;
            color: #856404;
            padding: 8px 12px;
            border-radius: 5px;
            font-size: 12px;
            margin-bottom: 15px;
        }
        button {
            background: #007bff;
            color: white;
            border: none;
            padding: 12px 30px;
            font-size: 16px;
            border-radius: 8px;
            cursor: pointer;
            transition: transform 0.2s, background 0.2s;
        }
        button:hover {
            background: #0056b3;
            transform: translateY(-2px);
        }
        button:disabled {
            background: #ccc;
            cursor: not-allowed;
            transform: none;
        }
        .log {
            margin-top: 20px;
            text-align: left;
            background: #1e1e1e;
            color: #d4d4d4;
            padding: 15px;
            border-radius: 8px;
            font-family: 'Consolas', 'Monaco', monospace;
            font-size: 12px;
            height: 350px;
            overflow-y: auto;
        }
        .log-entry {
            border-bottom: 1px solid #333;
            padding: 4px 0;
            white-space: pre-wrap;
            word-break: break-all;
        }
        .log-time {
            color: #858585;
            margin-right: 8px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>📞 Phone Region Checker</h1>
        <p>Выберите Excel файл с номерами телефонов в колонке A</p>
        <div class="warning">
            ⚠️ ВАЖНО: Закройте Excel-файл перед обработкой!
        </div>
        <button id="selectBtn">📁 Выбрать Excel файл</button>
        <div class="log" id="log"></div>
    </div>

     <script>
        const { ipcRenderer } = require('electron');
        const logDiv = document.getElementById('log');
        
        function addLog(message) {
            const entry = document.createElement('div');
            entry.className = 'log-entry';
            const time = new Date().toLocaleTimeString();
            entry.innerHTML = '<span class="log-time">[' + time + ']</span>' + message;
            logDiv.appendChild(entry);
            // Скролл НЕ двигается
        }
        
        ipcRenderer.on('log-message', (event, message) => {
            addLog(message);
        });
        
        document.getElementById('selectBtn').onclick = async () => {
            addLog('Выбор файла...');
            const result = await ipcRenderer.invoke('select-file');
            if (result.success) {
                addLog('✅ ' + result.message);
            } else {
                addLog('❌ ' + result.message);
            }
        };
        
        addLog('✅ Готов к работе. Нажмите кнопку для выбора Excel файла.');
    </script>
</body>
</html>`;

    mainWindow.loadURL(`data:text/html;charset=utf-8,${encodeURIComponent(html)}`);
}

async function handleFileSelect() {
    if (isProcessing) {
        sendLog('⏳ Обработка уже выполняется, пожалуйста, подождите...');
        return { success: false, message: 'Обработка уже выполняется. Пожалуйста, подождите.' };
    }

    const result = await dialog.showOpenDialog(mainWindow, {
        properties: ['openFile'],
        filters: [{ name: 'Excel Files', extensions: ['xlsx', 'xls'] }]
    });

    if (!result.canceled && result.filePaths.length > 0) {
        const filePath = result.filePaths[0];
        sendLog(`📁 Выбран файл: ${filePath}`);

        isProcessing = true;

        mainWindow.webContents.executeJavaScript(`
            document.getElementById('selectBtn').disabled = true;
            document.getElementById('selectBtn').textContent = '⏳ Обработка... Пожалуйста, подождите';
        `);

        try {
            await processExcelFile(filePath);
            return { success: true, message: 'Готово! Файл обновлён.' };
        } catch (error) {
            sendLog(`❌ Ошибка: ${error.message}`);
            return { success: false, message: error.message };
        } finally {
            isProcessing = false;
            mainWindow.webContents.executeJavaScript(`
                document.getElementById('selectBtn').disabled = false;
                document.getElementById('selectBtn').textContent = '📁 Выбрать Excel файл';
            `);
        }
    }
    return { success: false, message: 'Файл не выбран' };
}

// ========== START ==========

app.whenReady().then(() => {
    createWindow();
    ipcMain.handle('select-file', handleFileSelect);
});

app.on('window-all-closed', () => {
    app.quit();
});