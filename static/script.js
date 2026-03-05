// ─── DOM refs ────────────────────────────────────────────
const uploadArea    = document.getElementById('uploadArea');
const fileInput     = document.getElementById('fileInput');
const fileInfo      = document.getElementById('fileInfo');
const uploadBtn     = document.getElementById('uploadBtn');
const loader        = document.getElementById('loader');

const configToggle  = document.getElementById('configToggle');
const configContent = document.getElementById('configContent');

const uploadSection = document.getElementById('uploadSection');
const statusSection = document.getElementById('statusSection');
const reviewSection = document.getElementById('reviewSection');
const successSection= document.getElementById('successSection');
const errorSection  = document.getElementById('errorSection');

const reviewTableBody  = document.getElementById('reviewTableBody');
const keepHint         = document.getElementById('keepHint');
const extractHint      = document.getElementById('extractHint');
const reviewKeepBtn    = document.getElementById('reviewKeepBtn');
const reviewExtractBtn = document.getElementById('reviewExtractBtn');
const reviewSkipBtn    = document.getElementById('reviewSkipBtn');

const downloadBtn   = document.getElementById('downloadBtn');
const resetBtn      = document.getElementById('resetBtn');
const errorResetBtn = document.getElementById('errorResetBtn');
const successMsg    = document.getElementById('successMessage');
const errorMsg      = document.getElementById('errorMessage');

const addFabricBtn  = document.getElementById('addFabricBtn');
const fabricRows    = document.getElementById('fabricRows');

// ─── State ───────────────────────────────────────────────
let selectedFile = null;
let downloadId   = null;
let scanId       = null;       // Set after /scan succeeds
let formSnapshot = null;       // Saved form values for phase 2

// ─── Config toggle ────────────────────────────────────────
configToggle.addEventListener('click', () => {
    configToggle.classList.toggle('active');
    configContent.classList.toggle('show');
});

// ─── File drag-and-drop / click ───────────────────────────
uploadArea.addEventListener('click', () => fileInput.click());

uploadArea.addEventListener('dragover', e => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
});
uploadArea.addEventListener('dragleave', () => uploadArea.classList.remove('dragover'));
uploadArea.addEventListener('drop', e => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');
    if (e.dataTransfer.files.length > 0) handleFileSelect(e.dataTransfer.files[0]);
});
fileInput.addEventListener('change', e => {
    if (e.target.files.length > 0) handleFileSelect(e.target.files[0]);
});

function handleFileSelect(file) {
    const validExts = ['.xlsx', '.xls'];
    const ext = '.' + file.name.split('.').pop().toLowerCase();
    if (!validExts.includes(ext)) {
        showError('Invalid file type. Please upload .xlsx or .xls files only.');
        return;
    }
    if (file.size > 16 * 1024 * 1024) {
        showError('File too large. Maximum size is 16 MB.');
        return;
    }
    selectedFile = file;
    fileInfo.textContent = `📄 ${file.name}  (${(file.size / 1024).toFixed(1)} KB)`;
    fileInfo.classList.add('show');
    uploadBtn.disabled = false;
}

// ─── Dynamic Fabric Rows ──────────────────────────────────
function updateFabricPlaceholder() {
    const empty  = fabricRows.querySelector('.fabric-empty');
    const hasRows= fabricRows.querySelectorAll('.fabric-row').length > 0;
    if (hasRows && empty) empty.remove();
    if (!hasRows && !empty) {
        const p = document.createElement('p');
        p.className = 'fabric-empty';
        p.textContent = 'No additional fabrics added yet.';
        fabricRows.appendChild(p);
    }
}

addFabricBtn.addEventListener('click', () => {
    const row = document.createElement('div');
    row.className = 'fabric-row';
    row.innerHTML = `
        <input type="text" placeholder="Fabric name (e.g. Kitchen)" class="fabric-name">
        <input type="text" placeholder="Color code (e.g. KT4561)" class="fabric-color">
        <button class="btn-remove-fabric" title="Remove">×</button>
    `;
    row.querySelector('.btn-remove-fabric').addEventListener('click', () => {
        row.remove();
        updateFabricPlaceholder();
    });
    fabricRows.appendChild(row);
    updateFabricPlaceholder();
    row.querySelector('.fabric-name').focus();
});

function collectFabricColors() {
    const result = {};
    fabricRows.querySelectorAll('.fabric-row').forEach(row => {
        const name  = row.querySelector('.fabric-name').value.trim();
        const color = row.querySelector('.fabric-color').value.trim();
        if (name && color) result[name] = color;
    });
    return Object.keys(result).length > 0 ? result : null;
}

// ─── Collect form values ──────────────────────────────────
function collectFormData() {
    return {
        bedColor:      document.getElementById('bedColor').value.trim(),
        livColor:      document.getElementById('livColor').value.trim(),
        deductionCell: document.getElementById('deductionCell').value.trim(),
        fabricColors:  collectFabricColors()
    };
}

// ─── Phase 1: Scan ────────────────────────────────────────
uploadBtn.addEventListener('click', async () => {
    if (!selectedFile) return;

    formSnapshot = collectFormData();

    const fd = new FormData();
    fd.append('file', selectedFile);

    uploadBtn.disabled = true;
    uploadBtn.classList.add('loading');
    uploadSection.style.display = 'none';
    statusSection.style.display = 'block';

    try {
        const res  = await fetch('/scan', { method: 'POST', body: fd });
        const data = await res.json();

        if (!res.ok) throw new Error(data.error || 'Scan failed');

        scanId = data.scan_id;

        if (!data.has_text_tags) {
            // No problematic rows — go straight to processing
            await processFile('skip');
        } else {
            // Show review card
            showReviewCard(data.rows);
        }
    } catch (err) {
        statusSection.style.display = 'none';
        showError(err.message);
        uploadBtn.classList.remove('loading');
    }
});

// ─── Review card ─────────────────────────────────────────
function showReviewCard(rows) {
    statusSection.style.display = 'none';

    // Populate table
    reviewTableBody.innerHTML = '';
    rows.forEach(r => {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td>${r.tag}</td><td>${r.fabric}</td><td>${r.width ?? ''}</td><td>${r.height ?? ''}</td>`;
        reviewTableBody.appendChild(tr);
    });

    // Build hint text from first tag
    const firstTag = rows[0]?.tag ?? '';
    const numMatch = firstTag.match(/\d+/);
    keepHint.textContent    = firstTag ? `e.g. ${firstTag}-Bed` : '';
    extractHint.textContent = numMatch  ? `e.g. ${numMatch[0]}-Bed` : firstTag ? `e.g. ${firstTag}-Bed` : '';

    reviewSection.style.display = 'block';
    uploadBtn.classList.remove('loading');
}

reviewKeepBtn.addEventListener('click',    () => proceedWithAction('keep'));
reviewExtractBtn.addEventListener('click', () => proceedWithAction('extract'));
reviewSkipBtn.addEventListener('click',    () => proceedWithAction('skip'));

async function proceedWithAction(action) {
    reviewSection.style.display = 'none';
    statusSection.style.display = 'block';
    await processFile(action);
}

// ─── Phase 2: Process ─────────────────────────────────────
async function processFile(tagAction) {
    try {
        const fd = new FormData();
        fd.append('scan_id',   scanId);
        fd.append('tag_action', tagAction);

        const { bedColor, livColor, deductionCell, fabricColors } = formSnapshot;
        if (bedColor)      fd.append('bed_color',     bedColor);
        if (livColor)      fd.append('liv_color',     livColor);
        if (deductionCell) fd.append('deduction_cell', deductionCell);
        if (fabricColors)  fd.append('fabric_colors', JSON.stringify(fabricColors));

        const res    = await fetch('/upload', { method: 'POST', body: fd });
        const result = await res.json();

        statusSection.style.display = 'none';

        if (res.ok && result.success) {
            downloadId = result.download_id;
            successMsg.textContent = result.filename;
            successSection.style.display = 'block';
        } else {
            throw new Error(result.error || 'Unknown error occurred');
        }
    } catch (err) {
        statusSection.style.display = 'none';
        showError(err.message);
    } finally {
        uploadBtn.classList.remove('loading');
    }
}

// ─── Download ─────────────────────────────────────────────
downloadBtn.addEventListener('click', () => {
    if (downloadId) window.location.href = `/download/${downloadId}`;
});

// ─── Reset ────────────────────────────────────────────────
resetBtn.addEventListener('click', resetForm);
errorResetBtn.addEventListener('click', resetForm);

function resetForm() {
    selectedFile  = null;
    downloadId    = null;
    scanId        = null;
    formSnapshot  = null;

    fileInput.value = '';
    fileInfo.textContent = '';
    fileInfo.classList.remove('show');
    uploadBtn.disabled = true;

    document.getElementById('bedColor').value      = '';
    document.getElementById('livColor').value      = '';
    document.getElementById('deductionCell').value = 'I6';
    fabricRows.innerHTML = '';
    updateFabricPlaceholder();
    reviewTableBody.innerHTML = '';

    reviewSection.style.display  = 'none';
    successSection.style.display = 'none';
    errorSection.style.display   = 'none';
    statusSection.style.display  = 'none';
    uploadSection.style.display  = 'block';
}

// ─── Error helper ─────────────────────────────────────────
function showError(message) {
    errorMsg.textContent = message;
    errorSection.style.display  = 'block';
    uploadSection.style.display = 'none';
}

// ─── Init ─────────────────────────────────────────────────
updateFabricPlaceholder();
