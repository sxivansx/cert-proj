(() => {
    // State
    let templateFilename = null;
    let dataFilename = null;
    let templateWidth = 0;
    let templateHeight = 0;
    let columns = [];
    let activeColumn = null;
    let placedFields = []; // { column, x, y, fontSize, color, font }
    let dataPreviewCollapsed = false;

    // DOM elements
    const templateDropzone = document.getElementById('template-dropzone');
    const templateInput = document.getElementById('template-input');
    const templateStatus = document.getElementById('template-status');
    const templateIcon = document.getElementById('template-icon');
    const dataDropzone = document.getElementById('data-dropzone');
    const dataInput = document.getElementById('data-input');
    const dataStatus = document.getElementById('data-status');
    const dataIcon = document.getElementById('data-icon');
    const mappingSection = document.getElementById('mapping-section');
    const columnsBar = document.getElementById('columns-bar');
    const previewWrapper = document.getElementById('preview-wrapper');
    const templatePreview = document.getElementById('template-preview');
    const markersLayer = document.getElementById('markers-layer');
    const fieldsList = document.getElementById('fields-list');
    const emptyFieldsMsg = document.getElementById('empty-fields-msg');
    const generateBtn = document.getElementById('generate-btn');
    const generateStatus = document.getElementById('generate-status');
    const previewSection = document.getElementById('preview-section');
    const previewTable = document.getElementById('data-preview-table');
    const previewToggle = document.getElementById('preview-toggle');
    const previewBody = document.getElementById('preview-body');
    const hintBanner = document.getElementById('hint-banner');
    const hintText = document.getElementById('hint-text');
    const fontSizeInput = document.getElementById('font-size');
    const textColorInput = document.getElementById('text-color');
    const fontSelect = document.getElementById('font-select');
    const uploadFontBtn = document.getElementById('upload-font-btn');
    const fontUploadInput = document.getElementById('font-upload-input');

    let availableFonts = [];

    const checkmarkSVG = `<svg width="40" height="40" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14"/>
        <polyline points="22 4 12 14.01 9 11.01"/>
    </svg>`;

    // ── Collapsible data preview ──
    previewToggle.addEventListener('click', () => {
        dataPreviewCollapsed = !dataPreviewCollapsed;
        previewToggle.classList.toggle('collapsed', dataPreviewCollapsed);
        previewBody.classList.toggle('collapsed', dataPreviewCollapsed);
    });

    // ── Load available fonts ──
    function loadFonts() {
        return fetch('/fonts-list')
            .then(r => r.json())
            .then(data => {
                availableFonts = data.fonts;
                refreshFontSelect();
            });
    }

    function refreshFontSelect() {
        const current = fontSelect.value;
        fontSelect.innerHTML = '<option value="">Default</option>';
        availableFonts.forEach(f => {
            const opt = document.createElement('option');
            opt.value = f;
            opt.textContent = f.replace(/\.(ttf|otf)$/i, '');
            fontSelect.appendChild(opt);
        });
        fontSelect.value = current;
    }

    loadFonts();

    // ── Font upload ──
    uploadFontBtn.addEventListener('click', () => fontUploadInput.click());
    fontUploadInput.addEventListener('change', async () => {
        const file = fontUploadInput.files[0];
        if (!file) return;
        const form = new FormData();
        form.append('font', file);
        try {
            const res = await fetch('/upload-font', { method: 'POST', body: form });
            const data = await res.json();
            if (data.error) { alert(data.error); return; }
            await loadFonts();
            fontSelect.value = data.filename;
        } catch (e) {
            alert('Font upload failed');
        }
        fontUploadInput.value = '';
    });

    // ── Dropzone helpers ──
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

    // ── Template upload ──
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
            templateStatus.textContent = `${file.name} (${data.width}\u00d7${data.height})`;
            templateDropzone.classList.add('uploaded');
            templateIcon.innerHTML = checkmarkSVG;

            templatePreview.src = data.url;
            checkReady();
        } catch (e) {
            templateStatus.textContent = 'Upload failed';
        }
    });

    // ── Data upload ──
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
            dataIcon.innerHTML = checkmarkSVG;

            renderColumns();
            renderDataPreview(data.columns, data.preview);
            checkReady();
        } catch (e) {
            dataStatus.textContent = 'Upload failed';
        }
    });

    function showSection(el) {
        if (!el.classList.contains('visible')) {
            el.classList.add('visible');
        }
    }

    function checkReady() {
        if (templateFilename && dataFilename && columns.length > 0) {
            showSection(mappingSection);
        }
    }

    // ── Column chips ──
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
                updateHintBanner();
            });
            columnsBar.appendChild(chip);
        });
    }

    // ── Active column guidance ──
    function updateHintBanner() {
        if (activeColumn) {
            hintText.textContent = `Click on the template to place "${activeColumn}"`;
            hintBanner.classList.add('active');
            previewWrapper.classList.add('awaiting-click');
        } else {
            hintBanner.classList.remove('active');
            previewWrapper.classList.remove('awaiting-click');
        }
    }

    // ── Click to place ──
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
        updateHintBanner();
        renderColumns();
        renderMarkers();
        renderFieldsList();
    });

    // ── Markers on template ──
    function renderMarkers() {
        markersLayer.innerHTML = '';

        const scale = templatePreview.clientWidth / templateWidth;

        placedFields.forEach(field => {
            const pctX = (field.x / templateWidth) * 100;
            const pctY = (field.y / templateHeight) * 100;

            const displayFontSize = Math.max(8, Math.round(field.fontSize * scale));

            const marker = document.createElement('div');
            marker.className = 'marker';
            marker.style.left = pctX + '%';
            marker.style.top = pctY + '%';

            marker.innerHTML = `
                <div class="marker-text" style="font-size:${displayFontSize}px; color:${field.color}">${field.column}</div>
                <div class="marker-tag">${field.column} (${field.fontSize}px, ${field.font ? field.font.replace(/\.(ttf|otf)$/i, '') : 'Default'})</div>
                <button class="remove-marker" data-column="${field.column}">&times;</button>
            `;

            marker.querySelector('.remove-marker').addEventListener('click', (e) => {
                e.stopPropagation();
                removeField(field.column);
            });

            // Prevent clicks on a marker from triggering "place field" on the wrapper
            marker.addEventListener('click', (e) => e.stopPropagation());

            // Drag to reposition
            marker.addEventListener('mousedown', (e) => {
                if (e.target.classList.contains('remove-marker')) return;
                e.preventDefault();
                e.stopPropagation();

                const startMouseX = e.clientX;
                const startMouseY = e.clientY;
                const startPctX = field.x / templateWidth;
                const startPctY = field.y / templateHeight;
                let dragged = false;

                marker.style.cursor = 'grabbing';

                const onMouseMove = (e) => {
                    dragged = true;
                    const layerW = markersLayer.clientWidth;
                    const layerH = markersLayer.clientHeight;
                    const newPctX = startPctX + (e.clientX - startMouseX) / layerW;
                    const newPctY = startPctY + (e.clientY - startMouseY) / layerH;
                    marker.style.left = (Math.max(0, Math.min(1, newPctX)) * 100) + '%';
                    marker.style.top = (Math.max(0, Math.min(1, newPctY)) * 100) + '%';
                };

                const onMouseUp = () => {
                    document.removeEventListener('mousemove', onMouseMove);
                    document.removeEventListener('mouseup', onMouseUp);
                    marker.style.cursor = '';
                    if (dragged) {
                        field.x = Math.round(parseFloat(marker.style.left) / 100 * templateWidth);
                        field.y = Math.round(parseFloat(marker.style.top) / 100 * templateHeight);
                        renderFieldsList();
                    }
                };

                document.addEventListener('mousemove', onMouseMove);
                document.addEventListener('mouseup', onMouseUp);
            });

            markersLayer.appendChild(marker);
        });
    }

    // ── Placed fields list (card rows) ──
    function buildFontOptions(selectedFont) {
        let html = `<option value=""${selectedFont === '' ? ' selected' : ''}>Default</option>`;
        availableFonts.forEach(f => {
            const label = f.replace(/\.(ttf|otf)$/i, '');
            html += `<option value="${f}"${f === selectedFont ? ' selected' : ''}>${label}</option>`;
        });
        return html;
    }

    function renderFieldsList() {
        fieldsList.innerHTML = '';

        if (placedFields.length === 0) {
            emptyFieldsMsg.style.display = '';
            return;
        }
        emptyFieldsMsg.style.display = 'none';

        placedFields.forEach(field => {
            const card = document.createElement('div');
            card.className = 'field-card';
            card.innerHTML = `
                <div class="field-card-row1">
                    <span class="field-card-name">${field.column}</span>
                    <input type="color" class="field-color-swatch" value="${field.color}" title="Change color">
                    <button class="field-remove-btn" data-column="${field.column}" title="Remove">&times;</button>
                </div>
                <div class="field-card-row2">
                    <span class="field-detail">(${field.x}, ${field.y})</span>
                    <label class="field-detail">size: <input type="number" value="${field.fontSize}" min="8" max="200"></label>
                    <label class="field-detail">font: <select>${buildFontOptions(field.font)}</select></label>
                </div>
            `;

            card.querySelector('.field-remove-btn').addEventListener('click', () => removeField(field.column));

            card.querySelector('.field-color-swatch').addEventListener('input', (e) => {
                field.color = e.target.value;
                renderMarkers();
            });

            card.querySelector('.field-card-row2 input[type="number"]').addEventListener('input', (e) => {
                field.fontSize = parseInt(e.target.value) || 30;
                renderMarkers();
            });

            card.querySelector('.field-card-row2 select').addEventListener('change', (e) => {
                field.font = e.target.value;
                renderMarkers();
            });

            fieldsList.appendChild(card);
        });
    }

    function removeField(column) {
        placedFields = placedFields.filter(f => f.column !== column);
        renderColumns();
        renderMarkers();
        renderFieldsList();
    }

    // ── Data preview table ──
    function renderDataPreview(cols, rows) {
        showSection(previewSection);
        // Collapse by default after first render
        dataPreviewCollapsed = true;
        previewToggle.classList.add('collapsed');
        previewBody.classList.add('collapsed');

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

    // ── Generate ──
    generateBtn.addEventListener('click', async () => {
        if (!templateFilename || !dataFilename || placedFields.length === 0) {
            generateStatus.textContent = 'Please upload files and place at least one field.';
            generateStatus.className = 'status-msg error';
            return;
        }

        generateBtn.disabled = true;
        generateBtn.classList.add('loading');
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
                generateBtn.classList.remove('loading');
                return;
            }

            // Try to get count from header
            const countHeader = res.headers.get('X-Certificate-Count');

            const blob = await res.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'certificates.zip';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);

            const countMsg = countHeader ? `${countHeader} certificates generated.` : 'Done!';
            generateStatus.textContent = `${countMsg} Your ZIP is downloading.`;
            generateStatus.className = 'status-msg success';
        } catch (e) {
            generateStatus.textContent = 'Network error. Please try again.';
            generateStatus.className = 'status-msg error';
        }

        generateBtn.disabled = false;
        generateBtn.classList.remove('loading');
    });

    // Re-render markers on window resize
    window.addEventListener('resize', () => {
        if (placedFields.length > 0) renderMarkers();
    });
})();
