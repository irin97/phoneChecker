const { app, BrowserWindow, dialog, ipcMain, Menu, shell } = require('electron');
const XlsxPopulate = require('xlsx-populate');
const path = require('path')

let mainWindow;
const RESPONSE_COUNT_PER_SEC = 10;
let isProcessing = false;

// Функция для отправки логов в окно
function sendLog(message) {
    if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.webContents.send('log-message', message);
    }
}

// Функция отправки прогресса в окно
function sendProgressUpdate(message) {
    if (mainWindow && !mainWindow.isDestroyed()) {
        mainWindow.webContents.send('progress-message', message);
    }
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

function flashIcon() {
    if (mainWindow && !mainWindow.isDestroyed()) mainWindow.flashFrame(true);
}

async function checkConnection() {
    const controller = new AbortController();
    const timer = setTimeout(() => controller.abort(), 3000)

    try {
        const response = await fetch('http://num.voxlink.ru/get/?num=79133350909&field=region', {
            signal: controller.signal
        })
        clearTimeout(timer)
        return response.ok
    } catch (err) {
        clearTimeout(timer)
        return false
    }
}

async function getData(phone) {
    const maxAttempts = 3;

    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
        try {
            const response = await fetch(`http://num.voxlink.ru/get/?num=${phone}&field=region`);
            if (!response.ok) {
                throw new Error(`Failed to get region: ${response.status} ${response.statusText}`);
            }
            return await response.text();
        } catch (err) {
            if (attempt === maxAttempts) {
                return err.message;
            }
            await new Promise(resolve => setTimeout(resolve, 300));
        }
    }
}

function getTime(timeInSec) {
    const min = Math.trunc(timeInSec / 60);
    const ces = Math.round(timeInSec % 60);
    return `${min} минут ${ces} секунд`
}

async function timeoutFetch(phones) {
    const formattedPhones = phones.map(item => String(item));
    const groupedArr = [];

    while (formattedPhones.length > 0) {
        let temp = formattedPhones.splice(0, RESPONSE_COUNT_PER_SEC);
        groupedArr.push(temp);
    }

    async function fetchGroup(i) {
        return await Promise.all(groupedArr[i].map(el => getData(el)));
    }

    const totalBatches = groupedArr.length;
    const totalTimeSec = totalBatches;
    let progressPercent = 0;
    let timeLeft = totalTimeSec;

    const allResults = [];

    for (let i = 0; i < totalBatches; i++) {
        const res = await fetchGroup(i);
        allResults.push(...res);

        progressPercent = Math.round((i + 1) / totalBatches * 100);
        timeLeft = timeLeft - 1;
        const formattedTime = getTime(timeLeft);

        const progressText = `⏳ Прогресс выполнения: ${progressPercent}%, осталось времени: ${formattedTime}`;

        sendProgressUpdate(progressText);

        if (i < totalBatches - 1) {
            await new Promise(resolve => setTimeout(resolve, 1000));
        }
    }

    sendLog(`✅ Все номера обработаны!`);
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
        throw new Error('ОШИБКА: Файл открыт в Excel! Пожалуйста, закройте Excel-файл и попробуйте снова.');
    }

    sendLog('📖 Читаем номера из Excel...');
    const phones = await readPhonesFromExcel(filePath);

    if (phones.length === 0) {
        throw new Error('❌ Номера телефонов не найдены!');
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
        throw new Error('Файл был открыт во время обработки! Пожалуйста, закройте Excel-файл и запустите снова.');
    }

    await addColumnToExcel(filePath, regions, 'Регион');

    sendLog('✅ Готово! Файл успешно обновлён.');
    const safePath = JSON.stringify(filePath).replace(/"/g, '&quot;');
    sendLog(`📁 Расположение файла: <span class="file-link" data-path="${safePath}">${filePath}</span>`);
    // sendLog(`📁 Расположение файла: ${filePath}`);
}

// ========== WINDOW ==========

function createWindow() {
    Menu.setApplicationMenu(null);
    mainWindow = new BrowserWindow({
        width: 700,
        height: 850,
        webPreferences: {
            nodeIntegration: true,
            contextIsolation: false
        }
    });
    
    mainWindow.loadFile(path.join(__dirname, 'index.html'));
}

async function handleFileSelect() {
    sendLog('🔍 Проверяем файл и доступность сервера')

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

        const isApiAvailable = await checkConnection();

        if (!isApiAvailable) {
            flashIcon();
            return {
                success: false,
                message: 'Нет связи с сервером! Возможно у вас включен VPN. ОТКЛЮЧИТЕ ВПН и попробуйте снова.'
            }
        }

        sendLog(`📁 Выбран файл: ${filePath}`);
        isProcessing = true;

        mainWindow.webContents.executeJavaScript(`
            document.getElementById('selectBtn').disabled = true;
            document.getElementById('selectBtn').textContent = '⏳ Обработка... Пожалуйста, подождите';
        `);

        try {
            await processExcelFile(filePath);
            flashIcon();
            return { success: true };
        } catch (error) {
            sendLog(`${error.message}`);
            flashIcon()
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

    ipcMain.on('open-folder-with-file', (event, filePath) => {
        if (filePath) {
            shell.showItemInFolder(filePath);
        }
    });
});

app.on('window-all-closed', () => {
    app.quit();
});