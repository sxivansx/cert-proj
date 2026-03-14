(() => {
    // State
    let templateFilename = null;
    let dataFilename = null;
    let templateWidth = 0;
    let templateHeight = 0;
    let columns = [];
    let activeColumn = null;
    let placedFields = []; // { column, x, y, fontSize, color, font }

    // DOM elements
    const templateDropzone = document.getElementById('template-dropzone');
    const templateInput = document.getElementById('template-input');
    const templateStatus = document.getElementById('template-status');
    const dataDropzone = document.getElementById('data-dropzone');
    const dataInput = document.getElementById('data-input');
    const dataStatus = document.getElementById('data-status');
    const mappingSection = document.getElementById('mapping-section');
    const columnsBar = document.getElementById('columns-bar');
    const previewWrapper = document.getElementById('preview-wrapper');
    const templatePreview = document.getElementById('template-preview');
    const markersLayer = document.getElementById('markers-layer');
    const fieldsList = document.getElementById('fields-list');
    const generateSection = document.getElementById('generate-section');
    const generateBtn = document.getElementById('generate-btn');
    const generateStatus = document.getElementById('generate-status');
    const previewSection = document.getElementById('preview-section');
    const previewTable = document.getElementById('data-preview-table');
    const fontSizeInput = document.getElementById('font-size');
    const textColorInput = document.getElementById('text-color');
    const fontSelect = document.getElementById('font-select');

    // Load available fonts
    fetch('/fonts-list')
        .then(r => r.json())
        .then(data => {
            data.fonts.forEach(f => {
                const opt = document.createElement('option');
                opt.value = f;
                opt.textContent = f.replace(/\.(ttf|otf)$/i, '');
                fontSelect.appendChild(opt);
            });
        });

    // --- Dropzone helpers ---
    function setupDropzone(zone, input, onFile) {
        zone.addEventListener('click', () => input.click());
        input.addEventListener('change', () => {
            if (input.files[0]) onFile(input.files[0]);
        });
        zone.addEventListener('dragover', e => {
            e.preventDefault();
            zone.classList.add('dragover');
        });
        zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
        zone.addEventListener('drop', e => {
            e.preventDefault();
            zone.classList.remove('dragover');
            if (e.dataTransfer.files[0]) onFile(e.dataTransfer.files[0]);
        });
    }

    // --- Template upload ---
    setupDropzone(templateDropzone, templateInput, async (file) => {
        const form = new FormData();
        form.append('template', file);
        templateStatus.textContent = 'Uploading...';

        try {
            const res = await fetch('/upload-template', { method: 'POST', body: form });
            const data = await res.json();
            if (data.error) {
                templateStatus.textContent = data.error;
                return;
            }
            templateFilename = data.filename;
            templateWidth = data.width;
            templateHeight = data.height;
            templateStatus.textContent = `${file.name} (${data.width}x${data.height})`;
            templateDropzone.classList.add('uploaded');

            templatePreview.src = data.url;
            checkReady();
        } catch (e) {
            templateStatus.textContent = 'Upload failed';
        }
    });

    // --- Data upload ---
    setupDropzone(dataDropzone, dataInput, async (file) => {
        const form = new FormData();
        form.append('datafile', file);
        dataStatus.textContent = 'Uploading...';

        try {
            const res = await fetch('/upload-data', { method: 'POST', body: form });
            const data = await res.json();
            if (data.error) {
                dataStatus.textContent = data.error;
                return;
            }
            dataFilename = data.filename;
            columns = data.columns;
            dataStatus.textContent = `${file.name} (${columns.length} columns)`;
            dataDropzone.classList.add('uploaded');

            renderColumns();
            renderDataPreview(data.columns, data.preview);
            checkReady();
        } catch (e) {
            dataStatus.textContent = 'Upload failed';
        }
    });

    function checkReady() {
        if (templateFilename && dataFilename && columns.length > 0) {
            mappingSection.classList.remove('hidden');
            generateSection.classList.remove('hidden');
        }
    }

    // --- Column chips ---
    function renderColumns() {
        columnsBar.innerHTML = '';
        columns.forEach(col => {
            const chip = document.createElement('button');
            chip.className = 'col-chip';
            chip.textContent = col;
            chip.dataset.column = col;

            if (placedFields.some(f => f.column === col)) {
                chip.classList.add('placed');
            }

            chip.addEventListener('click', () => {
                if (activeColumn === col) {
                    activeColumn = null;
                    chip.classList.remove('active');
                } else {
                    document.querySelectorAll('.col-chip.active').forEach(c => c.classList.remove('active'));
                    activeColumn = col;
                    chip.classList.add('active');
                }
            });
            columnsBar.appendChild(chip);
        });
    }

    // --- Click to place ---
    previewWrapper.addEventListener('click', (e) => {
        if (!activeColumn) return;

        const rect = templatePreview.getBoundingClientRect();
        const clickX = e.clientX - rect.left;
        const clickY = e.clientY - rect.top;

        const scaleX = templateWidth / templatePreview.clientWidth;
        const scaleY = templateHeight / templatePreview.clientHeight;

        const realX = Math.round(clickX * scaleX);
        const realY = Math.round(clickY * scaleY);

        // Remove existing placement for this column
        placedFields = placedFields.filter(f => f.column !== activeColumn);

        placedFields.push({
            column: activeColumn,
            x: realX,
            y: realY,
            fontSize: parseInt(fontSizeInput.value) || 30,
            color: textColorInput.value || '#684631',
            font: fontSelect.value || ''
        });

        activeColumn = null;
        renderColumns();
        renderMarkers();
        renderFieldsList();
    });

    // --- Markers on template ---
    function renderMarkers() {
        markersLayer.innerHTML = '';

        const scale = templatePreview.clientWidth / templateWidth;

        placedFields.forEach(field => {
            const pctX = (field.x / templateWidth) * 100;
            const pctY = (field.y / templateHeight) * 100;

            // Scale font size from real image coords to displayed size
            const displayFontSize = Math.max(8, Math.round(field.fontSize * scale));

            const marker = document.createElement('div');
            marker.className = 'marker';
            marker.style.left = pctX + '%';
            marker.style.top = pctY + '%';

            marker.innerHTML = `
                <div class="marker-text" style="font-size:${displayFontSize}px; color:${field.color}">${field.column}</div>
                <div class="marker-tag">${field.column} (${field.fontSize}px)</div>
                <button class="remove-marker" data-column="${field.column}">&times;</button>
            `;

            marker.querySelector('.remove-marker').addEventListener('click', (e) => {
                e.stopPropagation();
                removeField(field.column);
            });

            markersLayer.appendChild(marker);
        });
    }

    // --- Placed fields list ---
    function renderFieldsList() {
        fieldsList.innerHTML = '';
        placedFields.forEach(field => {
            const li = document.createElement('li');
            li.innerHTML = `
                <strong>${field.column}</strong>
                <span class="coord">(${field.x}, ${field.y})</span>
                <label class="field-size-label">size: <input type="number" class="field-size-input" value="${field.fontSize}" min="8" max="200"></label>
                <button class="remove-btn" data-column="${field.column}">&times;</button>
            `;
            li.querySelector('.remove-btn').addEventListener('click', () => removeField(field.column));
            li.querySelector('.field-size-input').addEventListener('input', (e) => {
                field.fontSize = parseInt(e.target.value) || 30;
                renderMarkers();
            });
            fieldsList.appendChild(li);
        });
    }

    function removeField(column) {
        placedFields = placedFields.filter(f => f.column !== column);
        renderColumns();
        renderMarkers();
        renderFieldsList();
    }

    // --- Data preview table ---
    function renderDataPreview(cols, rows) {
        previewSection.classList.remove('hidden');
        let html = '<thead><tr>';
        cols.forEach(c => html += `<th>${c}</th>`);
        html += '</tr></thead><tbody>';
        rows.forEach(row => {
            html += '<tr>';
            row.forEach(val => html += `<td>${val}</td>`);
            html += '</tr>';
        });
        html += '</tbody>';
        previewTable.innerHTML = html;
    }

    // --- Generate ---
    generateBtn.addEventListener('click', async () => {
        if (!templateFilename || !dataFilename || placedFields.length === 0) {
            generateStatus.textContent = 'Please upload files and place at least one field.';
            generateStatus.className = 'status-msg error';
            return;
        }

        generateBtn.disabled = true;
        generateStatus.textContent = 'Generating certificates...';
        generateStatus.className = 'status-msg loading';

        const form = new FormData();
        form.append('template_filename', templateFilename);
        form.append('data_filename', dataFilename);
        form.append('fields', JSON.stringify(placedFields));

        try {
            const res = await fetch('/generate', { method: 'POST', body: form });

            if (!res.ok) {
                const err = await res.json();
                generateStatus.textContent = err.error || 'Generation failed';
                generateStatus.className = 'status-msg error';
                generateBtn.disabled = false;
                return;
            }

            const blob = await res.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'certificates.zip';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);

            generateStatus.textContent = 'Done! Your ZIP is downloading.';
            generateStatus.className = 'status-msg success';
        } catch (e) {
            generateStatus.textContent = 'Network error. Please try again.';
            generateStatus.className = 'status-msg error';
        }

        generateBtn.disabled = false;
    });

    // Re-render markers on window resize
    window.addEventListener('resize', () => {
        if (placedFields.length > 0) renderMarkers();
    });
})();
