(function () {
    'use strict';

    let availableOffers = [...window.APP_CONFIG.defaultOffers];
    let companyMappings = {};

    const $ = (sel) => document.querySelector(sel);
    const $$ = (sel) => document.querySelectorAll(sel);

    function showToast(message, type = 'info') {
        const container = $('#toastContainer');
        const toast = document.createElement('div');
        toast.className = `toast toast-${type}`;
        toast.textContent = message;
        container.appendChild(toast);
        setTimeout(() => {
            toast.style.opacity = '0';
            toast.style.transition = 'opacity 0.3s';
            setTimeout(() => toast.remove(), 300);
        }, 3500);
    }

    function loadLogo() {
        fetch('/api/logo')
            .then(r => {
                if (r.ok) {
                    const img = document.createElement('img');
                    img.src = '/api/logo?' + Date.now();
                    img.alt = 'Lotus Logo';
                    const navLogo = $('#navLogo');
                    if (navLogo) {
                        navLogo.innerHTML = '';
                        navLogo.appendChild(img);
                    }
                }
            })
            .catch(() => {});
    }

    function initUploadZone() {
        const zone = $('#uploadZone');
        const input = $('#fileInput');
        if (!zone || !input) return;

        zone.addEventListener('click', () => input.click());
        zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('dragover'); });
        zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
        zone.addEventListener('drop', (e) => {
            e.preventDefault();
            zone.classList.remove('dragover');
            if (e.dataTransfer.files.length) handleFileUpload(e.dataTransfer.files[0]);
        });
        input.addEventListener('change', () => {
            if (input.files.length) handleFileUpload(input.files[0]);
        });
    }

    function handleFileUpload(file) {
        const formData = new FormData();
        formData.append('file', file);
        showToast('Uploading file...', 'info');

        fetch('/api/upload', { method: 'POST', body: formData })
            .then(r => r.json())
            .then(data => {
                if (data.error) {
                    showToast(data.error, 'error');
                    return;
                }
                $('#fileInfo').classList.remove('hidden');
                $('#fileName').textContent = data.filename;
                $('#fileRows').textContent = `${data.row_count.toLocaleString()} rows loaded`;
                $('#fileCompanies').textContent = `${data.company_count} companies found`;

                if (data.row_count === 0) {
                    showToast('File appears empty — check format or sheet', 'error');
                } else if (!data.has_manufacturer_column) {
                    showToast('Warning: "Manufacturer Name" column not found', 'error');
                } else if (data.company_count === 0) {
                    showToast('No companies found in Manufacturer Name column', 'error');
                } else {
                    showToast(`Loaded ${data.row_count.toLocaleString()} rows, ${data.company_count} companies`, 'success');
                }

                buildMappingUI(data.companies);
            })
            .catch(() => showToast('Upload failed', 'error'));
    }

    function buildMappingUI(companies) {
        const empty = $('#mappingEmpty');
        const table = $('#mappingTable');
        const tbody = $('#mappingBody');
        companyMappings = {};

        if (!companies || companies.length === 0) {
            empty.classList.remove('hidden');
            table.classList.add('hidden');
            empty.innerHTML = '<p>No companies found. Check Manufacturer Name column.</p>';
            return;
        }

        empty.classList.add('hidden');
        table.classList.remove('hidden');
        tbody.innerHTML = '';

        companies.forEach((comp) => {
            companyMappings[comp] = 'Skip';
            const tr = document.createElement('tr');
            const options = availableOffers.map(o => `<option value="${o}">${o}</option>`).join('');
            tr.innerHTML = `<td><strong>${comp}</strong></td><td><select class="form-select offer-select" data-company="${comp}">${options}</select></td>`;
            tbody.appendChild(tr);
        });

        $$('.offer-select').forEach(sel => {
            sel.addEventListener('change', () => {
                companyMappings[sel.dataset.company] = sel.value;
            });
        });
    }

    function initDateFilter() {
        const chk = $('#filterByDate');
        const dateGroup = $('#dateRangeGroup');
        const startDate = $('#startDate');
        const endDate = $('#endDate');
        if (!chk) return;

        chk.addEventListener('change', () => {
            const enabled = chk.checked;
            dateGroup.classList.toggle('hidden', !enabled);
            startDate.disabled = !enabled;
            endDate.disabled = !enabled;
        });
    }

    function initAddOffer() {
        $('#btnAddOffer').addEventListener('click', () => {
            const input = $('#newOffer');
            const offer = input.value.trim().toUpperCase();
            if (!offer) return;
            if (availableOffers.includes(offer)) {
                showToast('Offer already exists', 'error');
                return;
            }
            availableOffers.push(offer);
            updateOfferDropdowns();
            input.value = '';
            showToast(`Offer '${offer}' added`, 'success');
        });
    }

    function updateOfferDropdowns() {
        const setAll = $('#setAllOffers');
        setAll.innerHTML = availableOffers.map(o => `<option value="${o}">${o}</option>`).join('');
        $$('.offer-select').forEach(sel => {
            const current = sel.value;
            sel.innerHTML = availableOffers.map(o =>
                `<option value="${o}" ${o === current ? 'selected' : ''}>${o}</option>`
            ).join('');
        });
    }

    function initSetAll() {
        $('#setAllOffers').addEventListener('change', (e) => {
            const val = e.target.value;
            $$('.offer-select').forEach(sel => {
                sel.value = val;
                companyMappings[sel.dataset.company] = val;
            });
        });
    }

    function initProcess() {
        $('#btnProcess').addEventListener('click', () => {
            $$('.offer-select').forEach(sel => {
                companyMappings[sel.dataset.company] = sel.value;
            });

            if (Object.keys(companyMappings).length === 0) {
                showToast('Please upload a file first', 'error');
                return;
            }

            const payload = {
                company_mappings: companyMappings,
                branch: $('#branch').value,
                filter_by_date: $('#filterByDate').checked,
                start_date: $('#startDate').value || null,
                end_date: $('#endDate').value || null,
                target_discount: $('#targetDiscount').value
            };

            $('#processingOverlay').classList.remove('hidden');

            fetch('/api/process', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload)
            })
                .then(r => r.json())
                .then(data => {
                    $('#processingOverlay').classList.add('hidden');
                    if (data.error) {
                        showToast(data.error, 'error');
                        return;
                    }
                    showResults(data);
                })
                .catch(() => {
                    $('#processingOverlay').classList.add('hidden');
                    showToast('Processing failed', 'error');
                });
        });
    }

    function showResults(data) {
        const body = $('#resultsBody');
        body.innerHTML = `
            <div class="stat"><span>Version</span><span class="stat-value">${data.version}</span></div>
            <div class="stat"><span>Target Discount</span><span class="stat-value">${data.target_discount}</span></div>
            <div class="stat"><span>Achieved Discount</span><span class="stat-value">${data.achieved_discount} EGP</span></div>
            <div class="stat"><span>Processed Items</span><span class="stat-value">${data.processed_count}</span></div>
            <div class="stat"><span>Unprocessed Items</span><span class="stat-value">${data.unprocessed_count}</span></div>
            <div class="stat"><span>Offer Sheets</span><span class="stat-value">${data.offer_sheets.join(', ')}</span></div>
        `;
        $('#resultsModal').classList.remove('hidden');
    }

    function initClear() {
        $('#btnClear').addEventListener('click', () => {
            fetch('/api/clear', { method: 'POST' })
                .then(() => {
                    $('#fileInfo').classList.add('hidden');
                    $('#fileName').textContent = '';
                    $('#fileRows').textContent = '';
                    $('#fileCompanies').textContent = '';
                    $('#targetDiscount').value = '';
                    $('#branch').selectedIndex = 0;
                    $('#filterByDate').checked = false;
                    $('#dateRangeGroup').classList.add('hidden');
                    $('#startDate').disabled = true;
                    $('#endDate').disabled = true;
                    $('#mappingEmpty').classList.remove('hidden');
                    $('#mappingTable').classList.add('hidden');
                    $('#mappingBody').innerHTML = '';
                    companyMappings = {};
                    $('#fileInput').value = '';
                    showToast('Data cleared', 'info');
                });
        });
    }

    function initModal() {
        $('#btnCloseModal').addEventListener('click', () => $('#resultsModal').classList.add('hidden'));
        const backdrop = $('.modal-backdrop');
        if (backdrop) backdrop.addEventListener('click', () => $('#resultsModal').classList.add('hidden'));
    }

    function initLogoUpload() {
        const btn = $('#btnUploadLogo');
        const input = $('#logoInput');
        if (!btn || !input) return;
        btn.addEventListener('click', () => input.click());
        input.addEventListener('change', () => {
            if (!input.files.length) return;
            const formData = new FormData();
            formData.append('logo', input.files[0]);
            fetch('/api/logo', { method: 'POST', body: formData })
                .then(r => r.json())
                .then(data => {
                    if (data.error) { showToast(data.error, 'error'); return; }
                    loadLogo();
                    showToast('Logo uploaded', 'success');
                })
                .catch(() => showToast('Logo upload failed', 'error'));
        });
    }

    document.addEventListener('DOMContentLoaded', () => {
        loadLogo();
        initUploadZone();
        initDateFilter();
        initAddOffer();
        initSetAll();
        initProcess();
        initClear();
        initModal();
        initLogoUpload();
    });
})();
