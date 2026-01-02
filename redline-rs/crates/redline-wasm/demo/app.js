import init, { 
    compare_word_documents_with_changes, 
    compare_spreadsheets, 
    compare_presentations,
    build_change_list,
    build_sml_change_list,
    build_pml_change_list
} from '../pkg/redline_wasm.js';

let wasmInitialized = false;
let originalFile = null;
let modifiedFile = null;
let lastResult = null;

async function initialize() {
    try {
        await init();
        wasmInitialized = true;
        console.log("WASM Initialized");
    } catch (e) {
        showError("Failed to initialize WebAssembly: " + e.message);
    }
}

function setupDropZone(id, isOriginal) {
    const zone = document.getElementById(id);
    const input = zone.querySelector('input');
    
    // Prevent default drag behaviors
    ['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
        zone.addEventListener(eventName, preventDefaults, false);
    });

    function preventDefaults(e) {
        e.preventDefault();
        e.stopPropagation();
    }

    // Highlight drop zone
    ['dragenter', 'dragover'].forEach(eventName => {
        zone.addEventListener(eventName, () => zone.classList.add('drag-over'), false);
    });

    ['dragleave', 'drop'].forEach(eventName => {
        zone.addEventListener(eventName, () => zone.classList.remove('drag-over'), false);
    });

    // Handle dropped files
    zone.addEventListener('drop', (e) => {
        const dt = e.dataTransfer;
        const files = dt.files;
        if (files.length > 0) {
            handleFile(files[0], isOriginal);
        }
    });
    
    // Handle click to upload
    zone.addEventListener('click', (e) => {
        input.click();
    });
    
    // Prevent input click from bubbling to zone (though input is hidden)
    input.addEventListener('click', (e) => {
        e.stopPropagation();
    });
    
    input.addEventListener('change', (e) => {
        if (e.target.files.length > 0) {
            handleFile(e.target.files[0], isOriginal);
        }
    });
}

function handleFile(file, isOriginal) {
    if (!file) return;
    
    if (isOriginal) {
        originalFile = file;
        document.getElementById('filename-original').textContent = file.name;
        document.getElementById('drop-zone-original').classList.add('has-file');
    } else {
        modifiedFile = file;
        document.getElementById('filename-modified').textContent = file.name;
        document.getElementById('drop-zone-modified').classList.add('has-file');
    }
    
    updateButtonState();
}

function updateButtonState() {
    const btn = document.getElementById('compare-btn');
    btn.disabled = !wasmInitialized || !originalFile || !modifiedFile;
}

function getFileType(filename) {
    const ext = filename.split('.').pop().toLowerCase();
    if (['docx', 'docm'].includes(ext)) return 'docx';
    if (['xlsx', 'xlsm'].includes(ext)) return 'xlsx';
    if (['pptx', 'pptm'].includes(ext)) return 'pptx';
    return null;
}

async function readFileAsUint8Array(file) {
    return new Uint8Array(await file.arrayBuffer());
}

async function compare() {
    if (!originalFile || !modifiedFile) return;
    
    const type1 = getFileType(originalFile.name);
    const type2 = getFileType(modifiedFile.name);
    
    if (!type1 || !type2 || type1 !== type2) {
        showError("Files must be of the same type (Word, Excel, or PowerPoint)");
        return;
    }
    
    showLoading(true);
    hideResults();
    hideError();
    
    try {
        const bytes1 = await readFileAsUint8Array(originalFile);
        const bytes2 = await readFileAsUint8Array(modifiedFile);
        
        const settings = {
            author_for_revisions: document.getElementById('author-name').value
        };
        const settingsJson = JSON.stringify(settings);
        
        let result;
        
        if (type1 === 'docx') {
            result = compare_word_documents_with_changes(bytes1, bytes2, settingsJson);
        } else if (type1 === 'xlsx') {
            result = compare_spreadsheets(bytes1, bytes2, settingsJson);
        } else if (type1 === 'pptx') {
            result = compare_presentations(bytes1, bytes2, settingsJson);
        }
        
        lastResult = result;
        lastResult.fileType = type1;
        showResults(result, type1);
        
    } catch (e) {
        showError("Comparison failed: " + e.message);
        console.error(e);
    } finally {
        showLoading(false);
    }
}

function showResults(result, type) {
    const grid = document.getElementById('stats-grid');
    grid.innerHTML = '';
    
    const addStat = (label, value) => {
        const div = document.createElement('div');
        div.className = 'stat-card';
        div.innerHTML = `<span class="stat-value">${value}</span><span class="stat-label">${label}</span>`;
        grid.appendChild(div);
    };
    
    if (type === 'docx') {
        addStat('Insertions', result.insertions);
        addStat('Deletions', result.deletions);
        addStat('Total Changes', result.total_revisions);
    } else if (type === 'xlsx') {
        addStat('Cells Added', result.insertions); // Mapped from cells_added
        addStat('Cells Deleted', result.deletions);
        addStat('Total Changes', result.revision_count);
    } else if (type === 'pptx') {
        addStat('Inserted', result.insertions);
        addStat('Deleted', result.deletions);
        addStat('Total Changes', result.revision_count);
    }
    
    document.getElementById('results').classList.remove('hidden');
    
    // Build change list for preview
    let changeList;
    if (type === 'docx') {
        changeList = build_change_list(result.changes, null);
    } else if (type === 'xlsx') {
        changeList = build_sml_change_list(result.changes, null);
    } else if (type === 'pptx') {
        changeList = build_pml_change_list(result.changes, null);
    }
    
    document.getElementById('change-list-json').textContent = JSON.stringify(changeList, null, 2);
}

function downloadResult() {
    if (!lastResult) return;
    
    const blob = new Blob([lastResult.document], { 
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' 
    }); // Mime type doesn't strictly matter for download
    
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `redline-compare.${lastResult.fileType}`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function showError(msg) {
    const el = document.getElementById('error');
    el.textContent = msg;
    el.classList.remove('hidden');
}

function hideError() {
    document.getElementById('error').classList.add('hidden');
}

function showLoading(show) {
    const el = document.getElementById('loading');
    if (show) el.classList.remove('hidden');
    else el.classList.add('hidden');
}

function hideResults() {
    document.getElementById('results').classList.add('hidden');
}

// Event Listeners
document.addEventListener('DOMContentLoaded', () => {
    initialize();
    setupDropZone('drop-zone-original', true);
    setupDropZone('drop-zone-modified', false);
    
    document.getElementById('compare-btn').addEventListener('click', compare);
    document.getElementById('download-btn').addEventListener('click', downloadResult);
    document.getElementById('view-changes-btn').addEventListener('click', () => {
        document.getElementById('change-list-container').classList.toggle('hidden');
    });
});
