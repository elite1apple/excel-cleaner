// ─── State ───────────────────────────────────────────────────
let selectedFile = null;
let tagAction    = 'skip';
let scanId       = null;
let downloadId   = null;
let pendingRows  = [];   // text-tag rows stored while waiting on missing-header decision

// ─── DOM refs ────────────────────────────────────────────────
const dropZone    = document.getElementById('dropZone');
const fileInput   = document.getElementById('fileInput');
const filePill    = document.getElementById('filePill');
const fileNameEl  = document.getElementById('fileName');
const runBtn      = document.getElementById('runBtn');
const statusBar   = document.getElementById('statusBar');
const statusMsg   = document.getElementById('statusMsg');
const spinner     = document.getElementById('spinner');
const dlBtn       = document.getElementById('dlBtn');
const scanPreview = document.getElementById('scanPreview');
const scanHeader  = document.getElementById('scanHeader');
const scanTable   = document.getElementById('scanTable');
const missingBar  = document.getElementById('missingBar');
const missingMsg  = document.getElementById('missingMsg');
const fabricRows  = document.getElementById('fabricRows');
const addFabricBtn = document.getElementById('addFabricBtn');

// ─── Add Fabric rows ─────────────────────────────────────
addFabricBtn.addEventListener('click', () => {
    const row = document.createElement('div');
    row.className = 'fabric-row';
    row.innerHTML = `
        <input type="text" placeholder="Fabric name (e.g. Kitchen)" class="fabric-name">
        <input type="text" placeholder="Color code (e.g. KT001)" class="fabric-color">
        <button class="fabric-row-remove" title="Remove" type="button">✕</button>
    `;
    row.querySelector('.fabric-row-remove').addEventListener('click', () => row.remove());
    fabricRows.appendChild(row);
    row.querySelector('.fabric-name').focus();
});

function collectFabricColors() {
    const result = {};
    fabricRows.querySelectorAll('.fabric-row').forEach(row => {
        const name  = row.querySelector('.fabric-name').value.trim();
        const color = row.querySelector('.fabric-color').value.trim();
        if (name && color) result[name] = color;
    });
    return Object.keys(result).length > 0 ? JSON.stringify(result) : '';
}

// ─── File handling ───────────────────────────────────────────
fileInput.addEventListener('change', e => {
    if (e.target.files[0]) setFile(e.target.files[0]);
});

dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag'));
dropZone.addEventListener('drop', e => {
    e.preventDefault();
    dropZone.classList.remove('drag');
    if (e.dataTransfer.files[0]) setFile(e.dataTransfer.files[0]);
});

document.getElementById('clearFile').addEventListener('click', clearFile);

function setFile(f) {
    const ext = '.' + f.name.split('.').pop().toLowerCase();
    if (!['.xlsx', '.xls'].includes(ext)) { setStatus('Invalid file type. Please upload .xlsx or .xls files.', 'error'); return; }
    if (f.size > 16 * 1024 * 1024) { setStatus('File too large. Maximum size is 16 MB.', 'error'); return; }

    selectedFile = f;
    fileNameEl.textContent = f.name;
    filePill.style.display = 'flex';
    runBtn.disabled = false;
    dlBtn.classList.remove('show');
    scanPreview.classList.remove('show');
    missingBar.style.display = 'none';
    scanId = null;
    downloadId = null;
    setStatus('');
}

function clearFile() {
    selectedFile = null;
    scanId = null;
    downloadId = null;
    fileInput.value = '';
    filePill.style.display = 'none';
    runBtn.disabled = true;
    dlBtn.classList.remove('show');
    scanPreview.classList.remove('show');
    missingBar.style.display = 'none';
    setStatus('');
}

// ─── Tag pills ───────────────────────────────────────────────
document.querySelectorAll('.tag-pill').forEach(pill => {
    pill.addEventListener('click', () => {
        document.querySelectorAll('.tag-pill').forEach(p => p.classList.remove('active'));
        pill.classList.add('active');
        tagAction = pill.dataset.val;
    });
});

// ─── Status helpers ──────────────────────────────────────────
function setStatus(msg, type = '') {
    statusBar.className = 'status-bar' + (type ? ' ' + type : '');
    statusMsg.textContent = msg;
    spinner.style.display = (type === 'info') ? 'block' : 'none';
    if (!type) statusBar.style.display = 'none';
}

function showScan(rows) {
    scanTable.innerHTML = '';
    if (!rows.length) { scanPreview.classList.remove('show'); return; }
    scanHeader.textContent = `${rows.length} non-numeric tag row${rows.length > 1 ? 's' : ''} detected — using "${tagAction}" action`;
    const keys = ['tag', 'fabric', 'width', 'height'];
    rows.slice(0, 8).forEach(row => {
        const tr = document.createElement('tr');
        keys.forEach(k => {
            const td = document.createElement('td');
            td.textContent = row[k] !== undefined ? String(row[k]).slice(0, 40) : '';
            tr.appendChild(td);
        });
        scanTable.appendChild(tr);
    });
    scanPreview.classList.add('show');
}

// ─── Missing headers buttons ─────────────────────────────────
document.getElementById('headersFixBtn').addEventListener('click', () => {
    missingBar.style.display = 'none';
    clearFile();
});

document.getElementById('headersSkipBtn').addEventListener('click', async () => {
    missingBar.style.display = 'none';
    if (pendingRows.length > 0) showScan(pendingRows);
    await runUpload();
});

// ─── Main run ────────────────────────────────────────────────
runBtn.addEventListener('click', handleRun);

async function handleRun() {
    if (!selectedFile) return;
    runBtn.disabled = true;
    dlBtn.classList.remove('show');
    missingBar.style.display = 'none';
    scanPreview.classList.remove('show');
    pendingRows = [];

    setStatus('Scanning file…', 'info');

    const fd1 = new FormData();
    fd1.append('file', selectedFile);

    try {
        const r1   = await fetch('/scan', { method: 'POST', body: fd1 });
        const d1   = await r1.json();

        if (d1.error) { setStatus(d1.error, 'error'); runBtn.disabled = false; return; }

        scanId = d1.scan_id;
        pendingRows = d1.rows || [];

        // Missing headers → pause and ask user
        if (d1.has_missing_headers && d1.missing_headers.length > 0) {
            setStatus('');
            missingMsg.textContent = 'Missing column' + (d1.missing_headers.length > 1 ? 's' : '') +
                ': ' + d1.missing_headers.join(', ') + '. Control side data may be missing from output.';
            missingBar.className = 'status-bar warn';
            missingBar.style.display = 'flex';
            runBtn.disabled = false;
            return;
        }

        // Show text-tag preview if any
        if (d1.has_text_tags && pendingRows.length > 0) {
            showScan(pendingRows);
            setStatus(`Found ${pendingRows.length} tag row(s) — processing with "${tagAction}" action.`, 'warn');
            await pause(600);
        }

        await runUpload();

    } catch (err) {
        setStatus('Network error: ' + err.message, 'error');
        runBtn.disabled = false;
    }
}

// ─── Upload & clean ──────────────────────────────────────────
async function runUpload() {
    setStatus('Cleaning file…', 'info');

    const fd2 = new FormData();
    fd2.append('scan_id',    scanId);
    fd2.append('tag_action', tagAction);

    const bc = document.getElementById('bedColor').value.trim();
    const lc = document.getElementById('livColor').value.trim();
    const dc = document.getElementById('dedCell').value.trim() || 'I6';
    const fc = collectFabricColors();
    if (bc) fd2.append('bed_color',      bc);
    if (lc) fd2.append('liv_color',      lc);
    if (dc) fd2.append('deduction_cell', dc);
    if (fc) fd2.append('fabric_colors',  fc);

    try {
        const r2 = await fetch('/upload', { method: 'POST', body: fd2 });
        const d2 = await r2.json();

        if (d2.error) { setStatus(d2.error, 'error'); runBtn.disabled = false; return; }

        downloadId = d2.download_id;
        spinner.style.display = 'none';
        setStatus('Done! Your file is ready.', 'info');
        dlBtn.textContent = 'Download ' + d2.filename;
        dlBtn.classList.add('show');
        runBtn.disabled = false;
    } catch (err) {
        setStatus('Network error: ' + err.message, 'error');
        runBtn.disabled = false;
    }
}

// ─── Download ────────────────────────────────────────────────
dlBtn.addEventListener('click', () => {
    if (downloadId) window.location.href = '/download/' + encodeURIComponent(downloadId);
});

// ─── Util ────────────────────────────────────────────────────
function pause(ms) { return new Promise(r => setTimeout(r, ms)); }
