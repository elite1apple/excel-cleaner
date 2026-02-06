// DOM Elements
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');
const uploadBtn = document.getElementById('uploadBtn');
const loader = document.getElementById('loader');

const configToggle = document.getElementById('configToggle');
const configContent = document.getElementById('configContent');

const uploadSection = document.getElementById('uploadSection');
const statusSection = document.getElementById('statusSection');
const successSection = document.getElementById('successSection');
const errorSection = document.getElementById('errorSection');

const downloadBtn = document.getElementById('downloadBtn');
const resetBtn = document.getElementById('resetBtn');
const errorResetBtn = document.getElementById('errorResetBtn');

const successMessage = document.getElementById('successMessage');
const errorMessage = document.getElementById('errorMessage');

// State
let selectedFile = null;
let downloadId = null;

// Config Toggle
configToggle.addEventListener('click', () => {
    configToggle.classList.toggle('active');
    configContent.classList.toggle('show');
});

// File Upload Handlers
uploadArea.addEventListener('click', () => fileInput.click());

uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
});

uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('dragover');
});

uploadArea.addEventListener('drop', (e) => {
    e.preventDefault();
    uploadArea.classList.remove('dragover');

    const files = e.dataTransfer.files;
    if (files.length > 0) {
        handleFileSelect(files[0]);
    }
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleFileSelect(e.target.files[0]);
    }
});

function handleFileSelect(file) {
    // Validate file type
    const validTypes = ['application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'application/vnd.ms-excel'];
    const validExtensions = ['.xlsx', '.xls'];
    const fileExtension = '.' + file.name.split('.').pop().toLowerCase();

    if (!validTypes.includes(file.type) && !validExtensions.includes(fileExtension)) {
        showError('Invalid file type. Please upload .xlsx or .xls files only.');
        return;
    }

    // Validate file size (16MB)
    if (file.size > 16 * 1024 * 1024) {
        showError('File too large. Maximum size is 16MB.');
        return;
    }

    selectedFile = file;
    fileInfo.textContent = `📄 ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
    fileInfo.classList.add('show');
    uploadBtn.disabled = false;
}

// Upload & Process
uploadBtn.addEventListener('click', async () => {
    if (!selectedFile) return;

    // Prepare form data
    const formData = new FormData();
    formData.append('file', selectedFile);

    // Add optional configuration
    const bedColor = document.getElementById('bedColor').value.trim();
    const livColor = document.getElementById('livColor').value.trim();
    const deductionCell = document.getElementById('deductionCell').value.trim();
    const fabricColors = document.getElementById('fabricColors').value.trim();

    if (bedColor) formData.append('bed_color', bedColor);
    if (livColor) formData.append('liv_color', livColor);
    if (deductionCell) formData.append('deduction_cell', deductionCell);
    if (fabricColors) {
        // Validate JSON
        try {
            JSON.parse(fabricColors);
            formData.append('fabric_colors', fabricColors);
        } catch (e) {
            showError('Invalid JSON format in Additional Fabric Colors field.');
            return;
        }
    }

    // Show loading state
    uploadBtn.disabled = true;
    uploadBtn.classList.add('loading');

    // Show status section
    uploadSection.style.display = 'none';
    statusSection.style.display = 'block';

    try {
        const response = await fetch('/upload', {
            method: 'POST',
            body: formData
        });

        const result = await response.json();

        if (response.ok && result.success) {
            // Success
            downloadId = result.download_id;
            successMessage.textContent = `File: ${result.filename}`;

            statusSection.style.display = 'none';
            successSection.style.display = 'block';
        } else {
            // Error from server
            throw new Error(result.error || 'Unknown error occurred');
        }
    } catch (error) {
        // Show error
        statusSection.style.display = 'none';
        showError(error.message);
    } finally {
        uploadBtn.classList.remove('loading');
    }
});

// Download
downloadBtn.addEventListener('click', () => {
    if (downloadId) {
        window.location.href = `/download/${downloadId}`;
    }
});

// Reset
resetBtn.addEventListener('click', resetForm);
errorResetBtn.addEventListener('click', resetForm);

function resetForm() {
    selectedFile = null;
    downloadId = null;

    fileInput.value = '';
    fileInfo.textContent = '';
    fileInfo.classList.remove('show');
    uploadBtn.disabled = true;

    document.getElementById('bedColor').value = '';
    document.getElementById('livColor').value = '';
    document.getElementById('deductionCell').value = 'I6';
    document.getElementById('fabricColors').value = '';

    successSection.style.display = 'none';
    errorSection.style.display = 'none';
    statusSection.style.display = 'none';
    uploadSection.style.display = 'block';
}

function showError(message) {
    errorMessage.textContent = message;
    errorSection.style.display = 'block';
    uploadSection.style.display = 'none';
}
