const dropZone = document.getElementById('drop-zone');
const fileInput = document.getElementById('file-input');
const statusContainer = document.getElementById('status-container');
const resultContainer = document.getElementById('result-container');
const sheetSelectionContainer = document.getElementById('sheet-selection-container');
const sheetList = document.getElementById('sheet-list');
const cancelSheetBtn = document.getElementById('cancel-sheet-btn');
const previewText = document.getElementById('preview-text');
const fileNameLabel = document.getElementById('file-name-label');
const downloadBtn = document.getElementById('download-btn');
const resetBtn = document.getElementById('reset-btn');

let processedData = '';
let currentWorkbook = null;
let currentSheetName = '';

// Drag and drop events
dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.classList.add('active');
});

dropZone.addEventListener('dragleave', () => {
    dropZone.classList.remove('active');
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    dropZone.classList.remove('active');
    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFileSelection(files[0]);
    }
});

// Click to select
dropZone.addEventListener('click', () => {
    fileInput.click();
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleFileSelection(e.target.files[0]);
    }
});

async function handleFileSelection(file) {
    const isExcel = file.name.endsWith('.xlsx') || file.name.endsWith('.xls');
    if (!isExcel) {
        alert('Por favor selecciona un archivo Excel (.xlsx o .xls)');
        return;
    }

    try {
        // Mostrar cargando
        dropZone.classList.add('hidden');
        statusContainer.classList.remove('hidden');

        const reader = new FileReader();
        reader.onload = (e) => {
            const data = new Uint8Array(e.target.result);
            currentWorkbook = XLSX.read(data, { type: 'array' });
            
            statusContainer.classList.add('hidden');
            renderSheetSelection(currentWorkbook.SheetNames);
        };
        reader.onerror = () => {
            throw new Error('No se pudo leer el archivo');
        };
        reader.readAsArrayBuffer(file);

    } catch (error) {
        alert('Error al leer el archivo: ' + error.message);
        resetUI();
    }
}

function renderSheetSelection(sheets) {
    sheetList.innerHTML = '';
    sheets.forEach(sheet => {
        const item = document.createElement('div');
        item.className = 'sheet-item';
        item.textContent = sheet;
        item.onclick = () => processSheet(sheet);
        sheetList.appendChild(item);
    });
    sheetSelectionContainer.classList.remove('hidden');
}

function processSheet(sheetName) {
    sheetSelectionContainer.classList.add('hidden');
    statusContainer.classList.remove('hidden');
    currentSheetName = sheetName;

    try {
        if (!currentWorkbook) throw new Error('No hay libro de trabajo cargado');
        const worksheet = currentWorkbook.Sheets[sheetName];
        if (!worksheet) throw new Error(`La hoja "${sheetName}" no existe`);

        // Convert to JSON first to handle pipe delimiter manually
        const data = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Join with pipe
        processedData = data.map(row => row.join('|')).join('\n');
        previewText.textContent = processedData;
        fileNameLabel.textContent = `${sheetName}.txt`;

        statusContainer.classList.add('hidden');
        resultContainer.classList.remove('hidden');
    } catch (error) {
        alert('Error al procesar la hoja: ' + error.message);
        sheetSelectionContainer.classList.remove('hidden');
        statusContainer.classList.add('hidden');
    }
}

const saveSuccess = document.getElementById('save-success');
const openFileLink = document.getElementById('open-file-link');

// Ocultar link de "Ver en carpeta" ya que en web no se puede abrir el explorador directamente de la misma forma
if (openFileLink) {
    openFileLink.classList.add('hidden');
}

// Download button
downloadBtn.addEventListener('click', () => {
    if (!processedData) return;

    const blob = new Blob([processedData], { type: 'text/plain' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${currentSheetName}.txt`;
    document.body.appendChild(a);
    a.click();
    
    setTimeout(() => {
        document.body.removeChild(a);
        window.URL.revokeObjectURL(url);
    }, 0);

    saveSuccess.classList.remove('hidden');
    saveSuccess.querySelector('span').textContent = 'Archivo descargado con éxito.';
});

// Cancel and Reset buttons
cancelSheetBtn.addEventListener('click', resetUI);
resetBtn.addEventListener('click', resetUI);

function resetUI() {
    processedData = '';
    currentWorkbook = null;
    currentSheetName = '';
    previewText.textContent = '';
    fileInput.value = '';
    dropZone.classList.remove('hidden');
    statusContainer.classList.add('hidden');
    resultContainer.classList.add('hidden');
    sheetSelectionContainer.classList.add('hidden');
    saveSuccess.classList.add('hidden');
}
