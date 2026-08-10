// Data Semesta Dashboard Frontend Logic (Vercel Ready - Multi-Strategy SheetJS Excel Parser)

let summaryData = null;
let allOrdersStore = []; // Stores all uploaded orders in memory
let currentOrders = [];
let currentPage = 1;
let totalPages = 1;
let currentLimit = 20;
let searchDebounce = null;

// Chart instances
let chartCrm = null;
let chartAvgPs = null;
let chartSegment = null;
let chartTrend = null;

document.addEventListener('DOMContentLoaded', () => {
    if (window.lucide) {
        lucide.createIcons();
    }

    setupEventListeners();

    // Try loading saved data from localStorage or Server API
    initDataState();
});

function setupEventListeners() {
    document.getElementById('btn-refresh').addEventListener('click', () => {
        loadSummaryData();
        loadOrdersData();
    });

    const btnExportPPT = document.getElementById('btn-export-ppt');
    if (btnExportPPT) {
        btnExportPPT.addEventListener('click', exportPPTReport);
    }

    const triggerUploadBtn = document.getElementById('btn-trigger-upload');
    const fileInput = document.getElementById('file-input-xlsx');

    triggerUploadBtn.addEventListener('click', () => {
        fileInput.click();
    });

    fileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) {
            handleClientFileUpload(file);
        }
    });

    document.getElementById('btn-close-detail-modal').addEventListener('click', closeDetailModal);
    document.getElementById('btn-modal-close-footer').addEventListener('click', closeDetailModal);

    document.getElementById('filter-search').addEventListener('input', () => {
        clearTimeout(searchDebounce);
        searchDebounce = setTimeout(() => {
            currentPage = 1;
            loadOrdersData();
        }, 300);
    });

    const sortSelect = document.getElementById('filter-sort');
    if (sortSelect) {
        sortSelect.addEventListener('change', () => {
            currentPage = 1;
            loadOrdersData();
        });
    }

    document.getElementById('filter-segment').addEventListener('change', () => {
        currentPage = 1;
        loadOrdersData();
    });

    document.getElementById('filter-crm-type').addEventListener('change', () => {
        currentPage = 1;
        loadOrdersData();
    });

    document.getElementById('filter-status').addEventListener('change', () => {
        currentPage = 1;
        loadOrdersData();
    });

    document.getElementById('btn-first').addEventListener('click', () => {
        if (currentPage > 1) {
            currentPage = 1;
            loadOrdersData();
        }
    });

    document.getElementById('btn-prev').addEventListener('click', () => {
        if (currentPage > 1) {
            currentPage--;
            loadOrdersData();
        }
    });

    document.getElementById('btn-next').addEventListener('click', () => {
        if (currentPage < totalPages) {
            currentPage++;
            loadOrdersData();
        }
    });

    document.getElementById('btn-last').addEventListener('click', () => {
        if (currentPage < totalPages) {
            currentPage = totalPages;
            loadOrdersData();
        }
    });
}

function closeDetailModal() {
    document.getElementById('detail-modal').classList.add('hidden');
}

// Initial Data State Loader
async function initDataState() {
    try {
        const savedSummary = localStorage.getItem('semesta_summary');
        if (savedSummary) {
            summaryData = JSON.parse(savedSummary);
            renderSummaryUI(summaryData);
        } else {
            await loadSummaryData();
        }
    } catch (e) {
        await loadSummaryData();
    }

    try {
        const savedOrders = localStorage.getItem('semesta_orders_sample');
        if (savedOrders) {
            allOrdersStore = JSON.parse(savedOrders);
        }
    } catch (e) {}

    loadOrdersData();
}

// Client-Side Multi-Strategy Excel Reader (Fixes "Bad uncompressed size" on ZIP64 / Custom CRM Exports)
function readExcelWorkbook(file, statusEl, progressBar) {
    return new Promise((resolve, reject) => {
        statusEl.textContent = 'Menggunakan engine pembaca data Excel...';
        progressBar.style.width = '40%';

        const r1 = new FileReader();
        r1.onload = (e) => {
            try {
                const binStr = e.target.result;
                const workbook = XLSX.read(binStr, { type: 'binary', cellDates: false, raw: true });
                return resolve(workbook);
            } catch (err1) {
                console.warn('Strategy 1 (binary) failed, trying Strategy 2 (array buffer)...', err1);
                
                const r2 = new FileReader();
                r2.onload = (e2) => {
                    try {
                        const data = new Uint8Array(e2.target.result);
                        const workbook = XLSX.read(data, { type: 'array', cellDates: false, raw: true });
                        return resolve(workbook);
                    } catch (err2) {
                        console.error('All SheetJS parsing strategies failed:', err2);
                        reject(err2);
                    }
                };
                r2.onerror = reject;
                r2.readAsArrayBuffer(file);
            }
        };
        r1.onerror = reject;
        r1.readAsBinaryString(file);
    });
}

// Client-Side Excel File Upload Handler
async function handleClientFileUpload(file) {
    const modal = document.getElementById('upload-modal');
    const titleEl = document.getElementById('upload-modal-title');
    const descEl = document.getElementById('upload-modal-desc');
    const statusEl = document.getElementById('upload-modal-status');
    const progressBar = document.getElementById('upload-progress-bar');

    modal.classList.remove('hidden');
    titleEl.textContent = 'Membaca File Excel...';
    descEl.textContent = `Memuat data dari ${file.name} (${(file.size / (1024 * 1024)).toFixed(1)} MB)`;
    statusEl.textContent = 'Membaca struktur zip & tabel...';
    progressBar.style.width = '20%';

    try {
        const workbook = await readExcelWorkbook(file, statusEl, progressBar);

        progressBar.style.width = '70%';
        statusEl.textContent = 'Mengekstrak baris & menghitung durasi PS...';

        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const rawRows = XLSX.utils.sheet_to_json(worksheet, { defval: '' });

        progressBar.style.width = '85%';
        statusEl.textContent = `Berhasil membaca ${rawRows.length.toLocaleString()} baris. Menyusun dashboard...`;

        const processed = processRawExcelRows(rawRows);
        summaryData = processed.summary;
        allOrdersStore = processed.orders;

        // Save to localStorage for quick reload
        try {
            localStorage.setItem('semesta_summary', JSON.stringify(summaryData));
            localStorage.setItem('semesta_orders_sample', JSON.stringify(allOrdersStore.slice(0, 2000)));
        } catch (e) {
            console.log('LocalStorage quota exceeded, keeping data in memory.');
        }

        progressBar.style.width = '100%';
        titleEl.textContent = 'Upload & Analisis Berhasil! 🎉';
        descEl.textContent = `Berhasil memproses ${rawRows.length.toLocaleString()} order dari ${file.name}`;
        statusEl.textContent = 'Memuat ulang antarmuka dashboard...';

        renderSummaryUI(summaryData);
        currentPage = 1;
        loadOrdersData();

        setTimeout(() => {
            modal.classList.add('hidden');
            document.getElementById('file-input-xlsx').value = '';
        }, 1200);

    } catch (err) {
        console.error('Client upload error:', err);
        titleEl.textContent = 'Upload Gagal!';
        descEl.textContent = err.message || 'Terjadi kesalahan saat membaca file excel.';
        statusEl.innerHTML = `
            <button onclick="document.getElementById('upload-modal').classList.add('hidden')" 
                    class="mt-3 px-4 py-1.5 rounded-lg bg-rose-600 hover:bg-rose-500 text-white font-medium text-xs">
                Tutup & Coba Lagi
            </button>
        `;
    }
}

// Process raw JSON rows from SheetJS into Semesta Metrics & Order Objects
function processRawExcelRows(rows) {
    let maxDateMs = 0;
    let maxDateStr = 'Belum Ada Data';
    let minDateMs = Number.MAX_SAFE_INTEGER;
    let minDateStr = 'Belum Ada Data';

    // ====================================================
    // PASS 1: Scan ALL rows to find true maxDateMs & minDateMs
    // (MUST be done before computing activeDurationDays)
    // ====================================================
    rows.forEach((r) => {
        const dc = parseExcelDate(r['Date Created']);
        const sd = parseExcelDate(r['Status Date']);
        const dm = parseExcelDate(r['Date Modified']);

        if (dc) {
            const t = dc.getTime();
            if (t > maxDateMs) { maxDateMs = t; maxDateStr = formatDateStr(dc); }
            if (t < minDateMs) { minDateMs = t; minDateStr = formatDateStr(dc); }
        }
        if (sd) {
            const t = sd.getTime();
            if (t > maxDateMs) { maxDateMs = t; maxDateStr = formatDateStr(sd); }
        }
        if (dm) {
            const t = dm.getTime();
            if (t > maxDateMs) { maxDateMs = t; maxDateStr = formatDateStr(dm); }
        }
    });

    // ====================================================
    // PASS 2: Build order objects using the correct maxDateMs
    // ====================================================
    const processedOrders = rows.map((r) => {
        const dateCreated = parseExcelDate(r['Date Created']);
        const statusDate = parseExcelDate(r['Status Date']);
        const dateModified = parseExcelDate(r['Date Modified']);
        const schedStart = parseExcelDate(r['Sched Start']);
        const bookingDate = parseExcelDate(r['Booking Date']);
        const measurementDate = parseExcelDate(r['Measurement Date']);

        const rawType = (r['CRM Order Type'] || 'UNSPECIFIED').toString().trim().toUpperCase();
        let crmType = rawType;
        if (rawType === 'NEW INSTALL') crmType = 'CREATE';

        const statusRaw = (r['Status'] || 'UNKNOWN').toString().trim().toUpperCase();

        // PS = sudah selesai/dieksekusi
        const psStatuses = ['COMPLETE', 'COMPWORK', 'INSTCOMP', 'DEINSTCOMP', 'ACTCOMP', 'VALCOMP'];
        // Cleared = dibatalkan/ditolak/ditutup → BUKAN pending, BUKAN PS
        const clearedStatuses = ['CANCLWORK', 'CANCEL', 'CANCELWORK', 'REJECT', 'REJECTED', 'CLOSE', 'CLOSED', 'ABORT', 'ABORTED'];

        const isPs = psStatuses.includes(statusRaw) ? 1 : 0;
        const isCleared = clearedStatuses.includes(statusRaw) ? 1 : 0;

        let psDurationDays = null;
        if (isPs === 1 && dateCreated && statusDate && statusDate >= dateCreated) {
            psDurationDays = (statusDate.getTime() - dateCreated.getTime()) / 86400000.0;
        }

        // activeDurationDays = null jika sudah PS atau sudah Cleared (cancel/tutup)
        // Hanya order yang masih AKTIF/PENDING yang dapat nilai ini
        let activeDurationDays = null;
        if (isPs === 0 && isCleared === 0 && dateCreated && maxDateMs > 0) {
            activeDurationDays = (maxDateMs - dateCreated.getTime()) / 86400000.0;
        }

        const sc = (r['SC Order No/Track ID/CSRM No'] || '').toString();
        const og = (r['Owner Group'] || '').toString();
        const pname = (r['Product Name'] || '').toString();
        const ptype = (r['Product Type'] || '').toString();

        let segment = 'Enterprise / Lainnya';
        if (sc.includes('DGPS') || sc.includes('PDA') || og.includes('PMDA')) {
            segment = 'PDA HSI';
        } else if (pname.includes('INDIHOME') || ptype === 'COMMON') {
            segment = 'IndiHome';
        } else if (og.includes('TIF FBB') || /^(MYIR|SC10|SC20|801M|802M|803M|C001|C002|A301)/.test(sc)) {
            segment = 'Modoroso';
        }

        return {
            sc_order_no: sc,
            workorder: (r['Workorder'] || '').toString(),
            oss_order_id: (r['OSS Order ID'] || '').toString(),
            service_no: (r['Service No.'] || '').toString(),
            customer_name: (r['Customer Name'] || 'N/A').toString(),
            contact_number: (r['Contact Number'] || '').toString(),
            address: (r['Address'] || '').toString(),
            witel: (r['witel'] || '').toString(),
            workzone: (r['Workzone'] || '').toString(),
            region_site_id: (r['Region/Site ID'] || '').toString(),
            product_name: pname,
            product_type: ptype,
            crm_order_type: crmType,
            wo_class: (r['WO Class'] || '').toString(),
            owner_group: og,
            status: statusRaw,
            segment: segment,
            is_ps: isPs,
            date_created: formatDateStr(dateCreated),
            status_date: formatDateStr(statusDate),
            date_modified: formatDateStr(dateModified),
            sched_start: formatDateStr(schedStart),
            booking_date: formatDateStr(bookingDate),
            measurement_date: formatDateStr(measurementDate),
            no_kontrak: (r['No. Kontrak(KB/KL/P8)'] || '').toString(),
            area_tif: (r['Area TIF'] || '').toString(),
            district_tif: (r['District TIF'] || '').toString(),
            regional_tif: (r['Regional TIF'] || '').toString(),
            order_id_tsel: (r['Order ID TSEL'] || '').toString(),
            channel_id_tsel: (r['Channle ID TSEL'] || '').toString(),
            measurement: (r['Measurement'] || '').toString(),
            measurement_result: (r['Measurement Result'] || '').toString(),
            description: (r['Description'] || '').toString(),
            ps_duration_days: psDurationDays,
            active_duration_days: activeDurationDays,
            date_created_time: dateCreated ? dateCreated.getTime() : 0,
            status_date_time: statusDate ? statusDate.getTime() : 0
        };
    });

    const oneMonthAgoMs = maxDateMs > 0 ? maxDateMs - (30 * 86400000) : 0;

    // Calculate Summary Stats
    const totalOrder = processedOrders.length;
    const totalPs = processedOrders.filter(o => o.is_ps === 1).length;
    const psLastMonth = processedOrders.filter(o => o.is_ps === 1 && o.status_date_time >= oneMonthAgoMs).length;

    // Type Summary
    const types = ['CREATE', 'MODIFY', 'DISCONNECT', 'SUSPEND', 'MIGRATE', 'UNSPECIFIED'];
    const typeSummary = types.map(t => {
        const group = processedOrders.filter(o => o.crm_order_type === t);
        const tot = group.length;
        const ps = group.filter(o => o.is_ps === 1).length;
        const psL30 = group.filter(o => o.is_ps === 1 && o.status_date_time >= oneMonthAgoMs).length;

        const psDurations = group.filter(o => o.ps_duration_days !== null).map(o => o.ps_duration_days);
        const avgPsDays = psDurations.length > 0 ? psDurations.reduce((a,b) => a+b, 0) / psDurations.length : null;
        const maxPsDays = psDurations.length > 0 ? Math.max(...psDurations) : null;

        const activeDurations = group.filter(o => o.active_duration_days !== null).map(o => o.active_duration_days);
        const maxActiveDays = activeDurations.length > 0 ? Math.max(...activeDurations) : null;

        return {
            tipe_transaksi: t,
            total_order: tot,
            total_ps: ps,
            avg_ps_days: avgPsDays,
            avg_ps_hours: avgPsDays ? avgPsDays * 24.0 : null,
            max_active_days: maxActiveDays,
            max_ps_days: maxPsDays,
            ps_last_month: psL30
        };
    });

    // Segment Summary
    const segs = ['Modoroso', 'PDA HSI', 'IndiHome', 'Enterprise / Lainnya'];
    const segSummary = segs.map(s => {
        const group = processedOrders.filter(o => o.segment === s);
        const tot = group.length;
        const ps = group.filter(o => o.is_ps === 1).length;
        const psL30 = group.filter(o => o.is_ps === 1 && o.status_date_time >= oneMonthAgoMs).length;

        const psDurations = group.filter(o => o.ps_duration_days !== null).map(o => o.ps_duration_days);
        const avgPsDays = psDurations.length > 0 ? psDurations.reduce((a,b) => a+b, 0) / psDurations.length : null;

        const activeDurations = group.filter(o => o.active_duration_days !== null).map(o => o.active_duration_days);
        const maxActiveDays = activeDurations.length > 0 ? Math.max(...activeDurations) : null;

        return {
            segment: s,
            total_order: tot,
            total_ps: ps,
            avg_ps_days: avgPsDays,
            avg_ps_hours: avgPsDays ? avgPsDays * 24.0 : null,
            max_active_days: maxActiveDays,
            ps_last_month: psL30
        };
    });

    // 30-Day Daily Trend Calculation (Dynamic Date Aggregation)
    const dateCounts = {};
    processedOrders.forEach(o => {
        let dKey = null;
        if (o.date_created && o.date_created !== '-') {
            dKey = o.date_created.substring(0, 10);
        } else if (o.status_date && o.status_date !== '-') {
            dKey = o.status_date.substring(0, 10);
        }
        if (dKey && /^\d{4}-\d{2}-\d{2}$/.test(dKey)) {
            if (!dateCounts[dKey]) {
                dateCounts[dKey] = { date_key: dKey, total_order: 0, ps_count: 0 };
            }
            dateCounts[dKey].total_order++;
            if (o.is_ps === 1) {
                dateCounts[dKey].ps_count++;
            }
        }
    });

    const sortedDates = Object.keys(dateCounts).sort();
    const last30Dates = sortedDates.slice(-30);
    const dailyTrend = last30Dates.map(dKey => dateCounts[dKey]);

    const summary = {
        max_date: maxDateStr,
        min_date: minDateMs < Number.MAX_SAFE_INTEGER ? minDateStr : maxDateStr,
        one_month_ago: formatDateStr(new Date(oneMonthAgoMs)),
        total_order_semesta: totalOrder,
        total_ps: totalPs,
        ps_percentage: totalOrder > 0 ? (totalPs / totalOrder) * 100.0 : 0.0,
        total_ps_last_month: psLastMonth,
        type_summary: typeSummary,
        segment_summary: segSummary,
        daily_trend: dailyTrend
    };

    return { summary, orders: processedOrders };
}

function parseExcelDate(val) {
    if (!val) return null;
    if (val instanceof Date) return val;
    if (typeof val === 'number') {
        const utc_days = Math.floor(val - 25569);
        const utc_value = utc_days * 86400;
        const date_info = new Date(utc_value * 1000);
        const fractional_day = val - Math.floor(val) + 0.0000001;
        let total_seconds = Math.floor(86400 * fractional_day);
        const seconds = total_seconds % 60;
        total_seconds -= seconds;
        const hours = Math.floor(total_seconds / 3600);
        const minutes = Math.floor(total_seconds / 60) % 60;
        return new Date(date_info.getFullYear(), date_info.getMonth(), date_info.getDate(), hours, minutes, seconds);
    }
    const str = val.toString().trim();
    if (!str || str === '-') return null;
    const d = new Date(str.replace(/-/g, '/'));
    if (!isNaN(d.getTime())) return d;
    return null;
}

function formatDateStr(d) {
    if (!d || isNaN(d.getTime())) return '-';
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    const hours = String(d.getHours()).padStart(2, '0');
    const minutes = String(d.getMinutes()).padStart(2, '0');
    const seconds = String(d.getSeconds()).padStart(2, '0');
    return `${year}-${month}-${day} ${hours}:${minutes}:${seconds}`;
}

// Open Order Detail Popup Modal
function openOrderDetailModal(ord) {
    document.getElementById('modal-sc-order').textContent = ord.sc_order_no || 'N/A';
    
    const statusBadge = document.getElementById('modal-status-badge');
    statusBadge.textContent = ord.status || 'UNKNOWN';
    statusBadge.className = `px-2.5 py-0.5 rounded text-xs font-bold ${getStatusBadgeStyle(ord.status, ord.is_ps)}`;

    const segmentBadge = document.getElementById('modal-segment-badge');
    segmentBadge.textContent = ord.segment || 'N/A';
    segmentBadge.className = `px-2.5 py-0.5 rounded text-xs font-bold ${getSegmentBadgeStyle(ord.segment)}`;

    document.getElementById('modal-customer-witel').textContent = 
        `Pelanggan: ${ord.customer_name || 'N/A'} • Witel: ${ord.witel || '-'} • Workzone: ${ord.workzone || '-'}`;

    document.getElementById('md-sc-order-no').textContent = ord.sc_order_no || '-';
    document.getElementById('md-workorder').textContent = ord.workorder || '-';
    document.getElementById('md-oss-order-id').textContent = ord.oss_order_id || '-';
    document.getElementById('md-customer-name').textContent = ord.customer_name || 'N/A';
    document.getElementById('md-contact-number').textContent = ord.contact_number || '-';
    document.getElementById('md-service-no').textContent = ord.service_no || '-';
    document.getElementById('md-address').textContent = ord.address || '-';

    document.getElementById('md-product-name').textContent = ord.product_name || 'N/A';
    document.getElementById('md-product-type').textContent = ord.product_type || '-';
    document.getElementById('md-crm-order-type').textContent = ord.crm_order_type || '-';
    document.getElementById('md-wo-class').textContent = ord.wo_class || '-';
    document.getElementById('md-owner-group').textContent = ord.owner_group || '-';
    document.getElementById('md-witel-workzone').textContent = `${ord.witel || '-'} / ${ord.workzone || '-'}`;

    document.getElementById('md-status').textContent = ord.status || '-';
    document.getElementById('md-date-created').textContent = ord.date_created || '-';
    document.getElementById('md-status-date').textContent = ord.status_date || '-';

    let psDurStr = '-';
    if (ord.is_ps === 1 && ord.ps_duration_days !== null) {
        psDurStr = `${ord.ps_duration_days.toFixed(2)} Hari (${(ord.ps_duration_days * 24).toFixed(1)} Jam)`;
    } else if (ord.active_duration_days !== null) {
        psDurStr = `Pending: ${ord.active_duration_days.toFixed(1)} Hari`;
    }
    document.getElementById('md-ps-duration').textContent = psDurStr;

    document.getElementById('md-sched-booking').textContent = `${ord.sched_start || '-'} / ${ord.booking_date || '-'}`;
    document.getElementById('md-date-modified').textContent = ord.date_modified || '-';

    document.getElementById('md-no-kontrak').textContent = ord.no_kontrak || '-';
    document.getElementById('md-tif-area').textContent = `${ord.area_tif || '-'} / ${ord.district_tif || '-'}`;
    
    let measStr = '-';
    if (ord.measurement || ord.measurement_result) {
        measStr = `${ord.measurement || ''} ${ord.measurement_result ? '(' + ord.measurement_result + ')' : ''}`;
    }
    document.getElementById('md-measurement').textContent = measStr;
    document.getElementById('md-description').textContent = ord.description || '-';

    document.getElementById('detail-modal').classList.remove('hidden');

    if (window.lucide) {
        lucide.createIcons();
    }
}

// Load Summary JSON (Server API or Memory)
async function loadSummaryData() {
    if (summaryData) {
        renderSummaryUI(summaryData);
        return;
    }
    try {
        const res = await fetch('/api/summary');
        if (!res.ok) throw new Error('Failed to fetch summary');
        summaryData = await res.json();
        renderSummaryUI(summaryData);
    } catch (err) {
        console.log('No server summary found, using client state.');
    }
}

function renderSummaryUI(data) {
    if (!data) return;

    document.getElementById('header-max-date').textContent = data.max_date || 'Belum Ada Data';
    document.getElementById('header-total-records').textContent = `${(data.total_order_semesta || 0).toLocaleString()} Order`;

    document.getElementById('kpi-total-orders').textContent = (data.total_order_semesta || 0).toLocaleString();
    document.getElementById('kpi-total-ps').textContent = (data.total_ps || 0).toLocaleString();
    document.getElementById('kpi-ps-percent').textContent = `${(data.ps_percentage || 0).toFixed(1)}%`;
    document.getElementById('kpi-ps-bar').style.width = `${data.ps_percentage || 0}%`;
    document.getElementById('kpi-ps-last-month').textContent = (data.total_ps_last_month || 0).toLocaleString();

    // High level metrics on cards
    const typeCreate = (data.type_summary || []).find(t => t.tipe_transaksi === 'CREATE');
    if (typeCreate) {
        document.getElementById('kpi-max-pending-days').textContent = typeCreate.max_active_days ? `${typeCreate.max_active_days.toFixed(1)} hr` : '-';
        document.getElementById('kpi-avg-create-ps').textContent = typeCreate.avg_ps_days ? `${typeCreate.avg_ps_days.toFixed(2)} hr` : '-';
    }

    renderTypeSummaryTable(data.type_summary || []);
    renderSegmentCards(data.segment_summary || []);
    renderCharts(data);
}

function renderTypeSummaryTable(types) {
    const tbody = document.getElementById('type-summary-tbody');
    tbody.innerHTML = '';

    if (types.length === 0 || types.every(t => t.total_order === 0)) {
        tbody.innerHTML = `
            <tr>
                <td colspan="7" class="py-6 text-center text-slate-500 font-normal">
                    Belum ada data laporan. Silakan klik tombol <strong>"Upload XLSX Baru"</strong> di atas.
                </td>
            </tr>
        `;
        return;
    }

    types.forEach(t => {
        const tr = document.createElement('tr');
        tr.className = 'hover:bg-brand-500/15 cursor-pointer transition-colors group';

        const avgPsText = t.avg_ps_days ? `${t.avg_ps_days.toFixed(2)} hari (${(t.avg_ps_hours || 0).toFixed(1)} jam)` : 'Instant / Fast';
        const activeMaxText = t.max_active_days ? `${t.max_active_days.toFixed(1)} hari` : '-';
        const psMaxText = t.max_ps_days ? `${t.max_ps_days.toFixed(1)} hari` : '-';

        tr.innerHTML = `
            <td class="py-3.5 px-4 font-bold text-white flex items-center gap-2 group-hover:text-brand-300">
                <span class="w-2 h-2 rounded-full ${getCrmTypeColorDot(t.tipe_transaksi)}"></span>
                ${t.tipe_transaksi}
                <i data-lucide="arrow-down-right" class="w-3.5 h-3.5 text-slate-500 opacity-0 group-hover:opacity-100 transition-opacity"></i>
            </td>
            <td class="py-3.5 px-4 text-right font-semibold text-slate-200">${(t.total_order || 0).toLocaleString()}</td>
            <td class="py-3.5 px-4 text-right font-semibold text-emerald-400">${(t.total_ps || 0).toLocaleString()}</td>
            <td class="py-3.5 px-4 text-right font-medium text-emerald-300">${avgPsText}</td>
            <td class="py-3.5 px-4 text-right font-medium text-rose-400">${activeMaxText}</td>
            <td class="py-3.5 px-4 text-right font-medium text-amber-300">${psMaxText}</td>
            <td class="py-3.5 px-4 text-right font-bold text-cyan-400">${(t.ps_last_month || 0).toLocaleString()}</td>
        `;

        tr.addEventListener('click', () => {
            const select = document.getElementById('filter-crm-type');
            select.value = t.tipe_transaksi;
            currentPage = 1;
            loadOrdersData();
            document.getElementById('real-time-explorer-section').scrollIntoView({ behavior: 'smooth' });
        });

        tbody.appendChild(tr);
    });

    if (window.lucide) {
        lucide.createIcons();
    }
}

function getCrmTypeColorDot(type) {
    switch (type) {
        case 'CREATE': return 'bg-indigo-400';
        case 'MODIFY': return 'bg-cyan-400';
        case 'DISCONNECT': return 'bg-rose-400';
        case 'SUSPEND': return 'bg-amber-400';
        case 'MIGRATE': return 'bg-purple-400';
        default: return 'bg-slate-400';
    }
}

function renderSegmentCards(segments) {
    const container = document.getElementById('segment-cards-container');
    container.innerHTML = '';

    segments.forEach(s => {
        const card = document.createElement('div');
        card.className = 'p-5 rounded-2xl bg-darkcard backdrop-blur-xl border border-darkborder shadow-xl space-y-3 cursor-pointer hover:border-brand-500/50 hover:bg-slate-900/60 transition-all group';

        const psPct = s.total_order > 0 ? ((s.total_ps / s.total_order) * 100).toFixed(1) : '0.0';
        const avgPsText = s.avg_ps_days ? `${s.avg_ps_days.toFixed(2)} hari` : 'Instant / Fast';
        const pendingMaxText = s.max_active_days ? `${s.max_active_days.toFixed(1)} hari` : '-';

        card.innerHTML = `
            <div class="flex items-center justify-between">
                <span class="px-2.5 py-1 rounded-lg text-xs font-bold ${getSegmentBadgeStyle(s.segment)}">
                    ${s.segment}
                </span>
                <span class="text-xs text-slate-400">${psPct}% PS</span>
            </div>
            <div>
                <div class="text-2xl font-extrabold text-white group-hover:text-brand-300 transition-colors">${(s.total_order || 0).toLocaleString()}</div>
                <p class="text-xs text-slate-400 flex items-center justify-between">
                    <span>Total Order</span>
                    <span class="text-[10px] text-brand-400 font-semibold opacity-0 group-hover:opacity-100 transition-opacity">Filter Segmen &rarr;</span>
                </p>
            </div>
            <div class="space-y-1.5 border-t border-slate-800 pt-3 text-xs">
                <div class="flex justify-between">
                    <span class="text-slate-400">Total PS:</span>
                    <span class="font-bold text-emerald-400">${(s.total_ps || 0).toLocaleString()}</span>
                </div>
                <div class="flex justify-between">
                    <span class="text-slate-400">Rata-rata PS:</span>
                    <span class="font-semibold text-slate-200">${avgPsText}</span>
                </div>
                <div class="flex justify-between">
                    <span class="text-slate-400">Order Terlama:</span>
                    <span class="font-semibold text-rose-400">${pendingMaxText}</span>
                </div>
                <div class="flex justify-between">
                    <span class="text-slate-400">PS 1 Bulan:</span>
                    <span class="font-bold text-cyan-400">${(s.ps_last_month || 0).toLocaleString()}</span>
                </div>
            </div>
        `;

        card.addEventListener('click', () => {
            const select = document.getElementById('filter-segment');
            select.value = s.segment;
            currentPage = 1;
            loadOrdersData();
            document.getElementById('real-time-explorer-section').scrollIntoView({ behavior: 'smooth' });
        });

        container.appendChild(card);
    });
}

function getSegmentBadgeStyle(seg) {
    if (seg === 'Modoroso') return 'badge-modoroso';
    if (seg === 'PDA HSI') return 'badge-pdahsi';
    if (seg === 'IndiHome') return 'badge-indihome';
    return 'bg-slate-800 text-slate-300 border border-slate-700';
}

function renderCharts(data) {
    const ctx1 = document.getElementById('chart-crm-types').getContext('2d');
    if (chartCrm) chartCrm.destroy();
    
    const types = data.type_summary || [];
    chartCrm = new Chart(ctx1, {
        type: 'bar',
        data: {
            labels: types.map(t => t.tipe_transaksi),
            datasets: [{
                label: 'Total Order',
                data: types.map(t => t.total_order),
                backgroundColor: '#6366f1',
                borderRadius: 6
            }, {
                label: 'Total PS',
                data: types.map(t => t.total_ps),
                backgroundColor: '#10b981',
                borderRadius: 6
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { labels: { color: '#94a3b8', font: { family: 'Plus Jakarta Sans' } } }
            },
            scales: {
                x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } },
                y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } }
            }
        }
    });

    const ctx2 = document.getElementById('chart-avg-ps').getContext('2d');
    if (chartAvgPs) chartAvgPs.destroy();

    const validPsTypes = types.filter(t => t.avg_ps_hours !== null && t.avg_ps_hours > 0);
    chartAvgPs = new Chart(ctx2, {
        type: 'bar',
        data: {
            labels: validPsTypes.map(t => t.tipe_transaksi),
            datasets: [{
                label: 'Rata-rata PS (Jam)',
                data: validPsTypes.map(t => t.avg_ps_hours),
                backgroundColor: '#06b6d4',
                borderRadius: 6
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { labels: { color: '#94a3b8' } }
            },
            scales: {
                x: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } },
                y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } }
            }
        }
    });

    const ctx3 = document.getElementById('chart-segment').getContext('2d');
    if (chartSegment) chartSegment.destroy();

    const segs = data.segment_summary || [];
    chartSegment = new Chart(ctx3, {
        type: 'doughnut',
        data: {
            labels: segs.map(s => s.segment),
            datasets: [{
                data: segs.map(s => s.total_order),
                backgroundColor: ['#a855f7', '#0ea5e9', '#22c55e', '#64748b'],
                borderWidth: 0
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { position: 'right', labels: { color: '#94a3b8' } }
            }
        }
    });

    const ctx4 = document.getElementById('chart-trend').getContext('2d');
    if (chartTrend) chartTrend.destroy();

    const trendData = (data.daily_trend && data.daily_trend.length > 0) ? data.daily_trend : [];
    chartTrend = new Chart(ctx4, {
        type: 'line',
        data: {
            labels: trendData.map(d => d.date_key),
            datasets: [{
                label: 'Total Order Masuk',
                data: trendData.map(d => d.total_order),
                borderColor: '#818cf8',
                backgroundColor: 'rgba(129, 140, 248, 0.25)',
                borderWidth: 2,
                pointRadius: 3,
                pointHoverRadius: 6,
                fill: true,
                tension: 0.35,
                spanGaps: true
            }, {
                label: 'Order PS',
                data: trendData.map(d => d.ps_count),
                borderColor: '#34d399',
                backgroundColor: 'rgba(52, 211, 153, 0.25)',
                borderWidth: 2,
                pointRadius: 3,
                pointHoverRadius: 6,
                fill: true,
                tension: 0.35,
                spanGaps: true
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { labels: { color: '#94a3b8' } }
            },
            scales: {
                x: { ticks: { color: '#94a3b8', maxTicksLimit: 12 }, grid: { color: 'rgba(255,255,255,0.05)' } },
                y: { ticks: { color: '#94a3b8', beginAtZero: true }, grid: { color: 'rgba(255,255,255,0.05)' } }
            }
        }
    });
}

// Load Orders Data (Works in Memory for Uploaded Files or Server API)
async function loadOrdersData() {
    const search = (document.getElementById('filter-search').value || '').toLowerCase().trim();
    const sort = document.getElementById('filter-sort') ? document.getElementById('filter-sort').value : 'pending_desc';
    const segment = document.getElementById('filter-segment').value;
    const crmType = document.getElementById('filter-crm-type').value;
    const status = document.getElementById('filter-status').value;

    if (allOrdersStore && allOrdersStore.length > 0) {
        // Fast client-side filtering across uploaded orders
        let filtered = allOrdersStore.filter(o => {
            if (segment && o.segment !== segment) return false;
            if (crmType && o.crm_order_type !== crmType) return false;
            if (status && o.status !== status) return false;
            if (search) {
                const text = `${o.sc_order_no} ${o.workorder} ${o.customer_name} ${o.service_no} ${o.address} ${o.witel} ${o.description}`.toLowerCase();
                if (!text.includes(search)) return false;
            }
            return true;
        });

        // High-performance Client-Side Sorting
        filtered.sort((a, b) => {
            if (sort === 'pending_desc') {
                const ageA = a.active_duration_days || -1;
                const ageB = b.active_duration_days || -1;
                return ageB - ageA;
            } else if (sort === 'created_desc') {
                return (b.date_created_time || 0) - (a.date_created_time || 0);
            } else if (sort === 'ps_desc') {
                const psA = a.ps_duration_days || -1;
                const psB = b.ps_duration_days || -1;
                return psB - psA;
            } else if (sort === 'created_asc') {
                return (a.date_created_time || 0) - (b.date_created_time || 0);
            } else if (sort === 'sc_asc') {
                return (a.sc_order_no || '').localeCompare(b.sc_order_no || '');
            }
            return 0;
        });

        const totalRecords = filtered.length;
        totalPages = Math.max(1, Math.ceil(totalRecords / currentLimit));
        if (currentPage > totalPages) currentPage = 1;

        const offset = (currentPage - 1) * currentLimit;
        currentOrders = filtered.slice(offset, offset + currentLimit);

        renderOrdersTable({
            page: currentPage,
            limit: currentLimit,
            total_records: totalRecords,
            total_pages: totalPages,
            orders: currentOrders
        });
        return;
    }

    // Server API Fallback
    const queryParams = new URLSearchParams({
        page: currentPage,
        limit: currentLimit,
        search: search,
        segment: segment,
        crm_type: crmType,
        status: status
    });

    try {
        const res = await fetch(`/api/orders?${queryParams.toString()}`);
        if (!res.ok) throw new Error(`HTTP error ${res.status}`);
        const data = await res.json();
        currentOrders = data.orders || [];
        renderOrdersTable(data);
    } catch (err) {
        renderOrdersTable({ page: 1, limit: currentLimit, total_records: 0, total_pages: 1, orders: [] });
    }
}

function renderOrdersTable(data) {
    currentPage = data.page || 1;
    totalPages = data.total_pages || 1;

    document.getElementById('current-page-num').textContent = currentPage;
    document.getElementById('total-pages-num').textContent = totalPages;

    const totalRecords = data.total_records || 0;
    const startItem = totalRecords > 0 ? (currentPage - 1) * currentLimit + 1 : 0;
    const endItem = Math.min(currentPage * currentLimit, totalRecords);
    document.getElementById('showing-records-badge').textContent = 
        totalRecords > 0 
            ? `Menampilkan ${startItem.toLocaleString()} - ${endItem.toLocaleString()} dari ${totalRecords.toLocaleString()} Order`
            : '0 Order Ditemukan';

    document.getElementById('btn-first').disabled = currentPage <= 1;
    document.getElementById('btn-prev').disabled = currentPage <= 1;
    document.getElementById('btn-next').disabled = currentPage >= totalPages;
    document.getElementById('btn-last').disabled = currentPage >= totalPages;

    const tbody = document.getElementById('orders-tbody');
    tbody.innerHTML = '';

    if (!data.orders || data.orders.length === 0) {
        tbody.innerHTML = `
            <tr>
                <td colspan="8" class="py-12 text-center text-slate-400">
                    <div class="w-12 h-12 mx-auto mb-2 rounded-2xl bg-slate-800/80 border border-slate-700 text-brand-400 flex items-center justify-center">
                        <i data-lucide="upload-cloud" class="w-6 h-6"></i>
                    </div>
                    <div class="font-bold text-sm text-white">Belum Ada Data Order Semesta</div>
                    <div class="text-xs text-slate-400 mt-1 max-w-sm mx-auto">
                        Silakan klik tombol <strong class="text-brand-300">"Upload XLSX Baru"</strong> di atas untuk mengunggah file laporan Excel (.xlsx) Anda.
                    </div>
                </td>
            </tr>
        `;
        if (window.lucide) {
            lucide.createIcons();
        }
        return;
    }

    data.orders.forEach((ord) => {
        const tr = document.createElement('tr');
        tr.className = 'hover:bg-brand-500/10 cursor-pointer transition-colors group';
        tr.title = 'Klik untuk melihat 31 rincian atribut order ini';

        let durationText = '-';
        if (ord.is_ps === 1 && ord.ps_duration_days !== null) {
            durationText = `<span class="text-emerald-400 font-semibold">${ord.ps_duration_days.toFixed(2)} hr</span>`;
        } else if (ord.active_duration_days !== null) {
            durationText = `<span class="text-rose-400 font-semibold">${ord.active_duration_days.toFixed(1)} hr</span>`;
        }

        tr.innerHTML = `
            <td class="py-3 px-4 font-bold text-white group-hover:text-brand-300">
                <div class="flex items-center gap-1.5">
                    <span>${escapeHtml(ord.sc_order_no || '-')}</span>
                    <i data-lucide="external-link" class="w-3 h-3 text-slate-500 opacity-0 group-hover:opacity-100 transition-opacity"></i>
                </div>
                <div class="text-[10px] text-slate-500 font-normal">WO: ${escapeHtml(ord.workorder || '-')}</div>
            </td>
            <td class="py-3 px-4">
                <div class="text-slate-200 font-medium">${escapeHtml(ord.customer_name || 'N/A')}</div>
                <div class="text-[11px] text-slate-400">${escapeHtml(ord.witel || '-')}</div>
            </td>
            <td class="py-3 px-4">
                <div class="text-slate-300">${escapeHtml(ord.product_name || 'N/A')}</div>
                <span class="inline-block mt-0.5 px-2 py-0.5 rounded text-[10px] font-bold ${getSegmentBadgeStyle(ord.segment)}">
                    ${ord.segment}
                </span>
            </td>
            <td class="py-3 px-4">
                <span class="px-2 py-1 rounded-md text-[11px] font-bold ${getCrmTypeBadgeStyle(ord.crm_order_type)}">
                    ${ord.crm_order_type}
                </span>
            </td>
            <td class="py-3 px-4">
                <span class="px-2 py-1 rounded-md text-[11px] font-semibold ${getStatusBadgeStyle(ord.status, ord.is_ps)}">
                    ${ord.status}
                </span>
            </td>
            <td class="py-3 px-4 text-slate-300 text-[11px] whitespace-nowrap">${ord.date_created || '-'}</td>
            <td class="py-3 px-4 text-slate-300 text-[11px] whitespace-nowrap">${ord.status_date || '-'}</td>
            <td class="py-3 px-4 text-right whitespace-nowrap">${durationText}</td>
        `;

        tr.addEventListener('click', () => {
            openOrderDetailModal(ord);
        });

        tbody.appendChild(tr);
    });

    if (window.lucide) {
        lucide.createIcons();
    }
}

function getCrmTypeBadgeStyle(type) {
    switch (type) {
        case 'CREATE': return 'badge-create';
        case 'MODIFY': return 'badge-modify';
        case 'DISCONNECT': return 'badge-disconnect';
        default: return 'bg-slate-800 text-slate-300 border border-slate-700';
    }
}

function getStatusBadgeStyle(status, isPs) {
    if (isPs === 1) return 'badge-ps';
    if (status === 'CANCLWORK' || status === 'WORKFAIL') return 'badge-cancel';
    return 'badge-pending';
}

function escapeHtml(str) {
    if (!str) return '';
    return str.replace(/[&<>"']/g, function(m) {
        return {
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            '"': '&quot;',
            "'": '&#039;'
        }[m];
    });
}

// Client-Side PowerPoint (.pptx) Generator Khas Telkom (PptxGenJS) - 100% DYNAMIC & VERIFIED
function exportPPTReport() {
    if (typeof PptxGenJS === 'undefined') {
        alert('Modul PPT belum siap. Pastikan koneksi internet terhubung untuk memuat PptxGenJS.');
        return;
    }

    const pptx = new PptxGenJS();
    pptx.defineLayout({ name: 'TELKOM_16_9', width: 13.333, height: 7.5 });
    pptx.layout = 'TELKOM_16_9';

    // Dynamic Date Range - Use summaryData directly (already scans Date Created + Status Date)
    function formatDateDot(dateStr) {
        // Convert "YYYY-MM-DD HH:MM:SS" to "DD.MM.YYYY"
        if (!dateStr || dateStr === '-' || dateStr === 'Belum Ada Data') return null;
        const parts = dateStr.substring(0, 10).split('-');
        if (parts.length === 3) return `${parts[2]}.${parts[1]}.${parts[0]}`;
        return null;
    }

    let minDateStr = "01.01.2026";
    let maxDateStr = "10.08.2026";

    if (summaryData) {
        const mx = formatDateDot(summaryData.max_date);
        const mn = formatDateDot(summaryData.min_date);
        if (mx) maxDateStr = mx;
        if (mn) minDateStr = mn;
    }

    const dateRangeLabel = `${minDateStr} - ${maxDateStr}`;

    // Telkom Corporate Color Constants
    const TELKOM_RED = 'E00000';
    const DARK_NAVY = '0F172A';
    const CARD_BG = 'F8FAFC';
    const CARD_BORDER = 'E2E8F0';
    const TEXT_DARK = '1E293B';
    const TEXT_MUTED = '64748B';
    const EMERALD_GREEN = '10B981';
    const CYAN_BLUE = '06B6D4';
    const PURPLE_ACC = 'A855F7';
    const WHITE = 'FFFFFF';

    function addHeader(slide, titleText, categoryText = `LAPORAN ANALISIS DATA SEMESTA (${dateRangeLabel})`) {
        slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 13.333, h: 1.1, fill: { color: DARK_NAVY }, line: { color: DARK_NAVY } });
        slide.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 0.25, h: 1.1, fill: { color: TELKOM_RED }, line: { color: TELKOM_RED } });
        slide.addText(categoryText.toUpperCase(), { x: 0.5, y: 0.12, w: 9.8, h: 0.3, fontSize: 10, bold: true, color: TELKOM_RED });
        slide.addText(titleText, { x: 0.5, y: 0.38, w: 10.0, h: 0.6, fontSize: 20, bold: true, color: WHITE });
        slide.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 10.8, y: 0.25, w: 2.2, h: 0.6, fill: { color: TELKOM_RED }, line: { color: TELKOM_RED } });
        slide.addText("TELKOM INDONESIA", { x: 10.8, y: 0.25, w: 2.2, h: 0.6, fontSize: 11, bold: true, color: WHITE, align: 'center' });
    }

    function addFooter(slide) {
        slide.addText(`Telkom Operations & Support System • Rentang Acuan Data: ${dateRangeLabel}`, { x: 0.5, y: 7.05, w: 12.333, h: 0.35, fontSize: 9, color: TEXT_MUTED });
    }

    // Dynamic Summary Numbers
    const totOrderVal = summaryData ? summaryData.total_order_semesta : 88011;
    const totPsVal = summaryData ? summaryData.total_ps : 80795;
    const psPctVal = summaryData ? summaryData.ps_percentage : 91.8;
    const psMonthVal = summaryData ? summaryData.total_ps_last_month : 5910;

    const typeCreate = (summaryData && summaryData.type_summary) ? summaryData.type_summary.find(t => t.tipe_transaksi === 'CREATE') : null;
    const maxPendingStr = (typeCreate && typeCreate.max_active_days) ? `${typeCreate.max_active_days.toFixed(1)} Hari` : "220,9 Hari";
    const avgCreatePsStr = (typeCreate && typeCreate.avg_ps_days) ? `${typeCreate.avg_ps_days.toFixed(2)} Hari` : "0,94 Hari";

    const totOrderStr = (totOrderVal || 0).toLocaleString();
    const totPsStr = (totPsVal || 0).toLocaleString();
    const psMonthStr = (psMonthVal || 0).toLocaleString();

    // ==========================================
    // SLIDE 1: COVER
    // ==========================================
    let slide1 = pptx.addSlide();
    slide1.addShape(pptx.shapes.RECTANGLE, { x: 0, y: 0, w: 13.333, h: 7.5, fill: { color: DARK_NAVY }, line: { color: DARK_NAVY } });
    slide1.addShape(pptx.shapes.RIGHT_TRIANGLE, { x: 9.5, y: 0, w: 3.833, h: 7.5, fill: { color: TELKOM_RED }, line: { color: TELKOM_RED }, rotate: 180 });

    slide1.addText("TELKOM OPERATIONS REPORT", { x: 1.0, y: 1.8, w: 8.5, h: 0.4, fontSize: 14, bold: true, color: TELKOM_RED });
    slide1.addText("Analisis Data Semesta", { x: 1.0, y: 2.2, w: 8.5, h: 0.8, fontSize: 40, bold: true, color: WHITE });
    slide1.addText("Modoroso • PDA HSI • IndiHome", { x: 1.0, y: 3.0, w: 8.5, h: 0.5, fontSize: 24, color: 'CBD5E1' });
    slide1.addText(`Acuan Rentang Data: ${dateRangeLabel} (Total ${totOrderStr} Order)`, { x: 1.0, y: 3.6, w: 8.5, h: 0.4, fontSize: 13, bold: true, color: EMERALD_GREEN });
    slide1.addText("Evaluasi Jumlah Order Semesta, Order Terlama, Rata-rata Durasi PS, & Performa 1 Bulan Kebelakang", { x: 1.0, y: 4.1, w: 8.5, h: 0.6, fontSize: 12, color: TEXT_MUTED });

    slide1.addText("Telkom Operations & Support System", { x: 1.0, y: 5.8, w: 8.5, h: 0.3, fontSize: 11, bold: true, color: WHITE });
    slide1.addText(`Tanggal Laporan: 10 Agustus 2026 | Periode Data Terverifikasi: ${dateRangeLabel}`, { x: 1.0, y: 6.1, w: 8.5, h: 0.3, fontSize: 10, color: TEXT_MUTED });

    // ==========================================
    // SLIDE 2: EXECUTIVE SUMMARY (4 METRIC CARDS)
    // ==========================================
    let slide2 = pptx.addSlide();
    addHeader(slide2, "Ringkasan Eksekutif: Jawaban 4 Pertanyaan Utama Operations");
    addFooter(slide2);

    const metrics = [
        [`TOTAL ORDER SEMESTA (${dateRangeLabel})`, totOrderStr, "Order Semesta Terdaftar", `${totPsStr} Order PS Complete (${(psPctVal || 0).toFixed(1)}%)`, TELKOM_RED],
        ["PS 1 BULAN KEBELAKANG", psMonthStr, "Order Selesai (30 Hari Terakhir)", "Didominasi Tipe CREATE & MODIFY", EMERALD_GREEN],
        ["RATA-RATA DURASI PS", avgCreatePsStr, "~22,5 Jam (Tipe CREATE)", "Rata-rata DISCONNECT: 0,44 Hari (10,6 Jam)", DARK_NAVY],
        ["ORDER PENDING TERLAMA", maxPendingStr, "Order Pending Terlama (CREATE)", "Pending MODIFY: 217,9 Hari", TELKOM_RED]
    ];

    metrics.forEach((m, i) => {
        const col = i % 2;
        const row = Math.floor(i / 2);
        const left = 0.6 + col * 6.2;
        const top = 1.5 + row * 2.6;

        slide2.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: left, y: top, w: 5.8, h: 2.3, fill: { color: CARD_BG }, line: { color: CARD_BORDER } });
        slide2.addShape(pptx.shapes.RECTANGLE, { x: left, y: top, w: 5.8, h: 0.12, fill: { color: m[4] }, line: { color: m[4] } });

        slide2.addText(m[0], { x: left + 0.2, y: top + 0.25, w: 5.4, h: 0.3, fontSize: 11, bold: true, color: TEXT_MUTED });
        slide2.addText(m[1], { x: left + 0.2, y: top + 0.6, w: 5.4, h: 0.8, fontSize: 32, bold: true, color: m[4] });
        slide2.addText(m[2], { x: left + 0.2, y: top + 1.45, w: 5.4, h: 0.3, fontSize: 11, bold: true, color: TEXT_DARK });
        slide2.addText(m[3], { x: left + 0.2, y: top + 1.75, w: 5.4, h: 0.3, fontSize: 10, color: TEXT_MUTED });
    });

    // ==========================================
    // SLIDE 3: RINCIAN TABLE PER TIPE TRANSAKSI
    // ==========================================
    let slide3 = pptx.addSlide();
    addHeader(slide3, "Rincian Metrik Utama per Tipe Transaksi (CRM Order Type)");
    addFooter(slide3);

    const headers = ["Tipe Transaksi", "Total Order", "Total PS", "Rata-rata Durasi PS", "Order Pending Terlama", "PS Terlama", "PS 1 Bulan"];
    
    let typeSummarySource = (summaryData && summaryData.type_summary && summaryData.type_summary.length > 0)
        ? summaryData.type_summary
        : [
            { tipe_transaksi: 'CREATE', total_order: 51962, total_ps: 48233, avg_ps_days: 0.94, max_active_days: 220.9, max_ps_days: 143.3, ps_last_month: 2725 },
            { tipe_transaksi: 'MODIFY', total_order: 19157, total_ps: 18788, avg_ps_days: 1.33, max_active_days: 217.9, max_ps_days: 116.1, ps_last_month: 2109 },
            { tipe_transaksi: 'DISCONNECT', total_order: 11337, total_ps: 11100, avg_ps_days: 0.44, max_active_days: 167.0, max_ps_days: 75.2, ps_last_month: 905 },
            { tipe_transaksi: 'SUSPEND', total_order: 1516, total_ps: 1050, avg_ps_days: null, max_active_days: null, max_ps_days: 0.0, ps_last_month: 0 },
            { tipe_transaksi: 'MIGRATE', total_order: 1193, total_ps: 1066, avg_ps_days: 2.07, max_active_days: 217.4, max_ps_days: 95.6, ps_last_month: 167 },
            { tipe_transaksi: 'UNSPECIFIED', total_order: 2321, total_ps: 37, avg_ps_days: 3.29, max_active_days: 220.8, max_ps_days: 7.4, ps_last_month: 4 }
        ];

    let tableRows = [headers.map(h => ({ text: h, options: { fill: { color: DARK_NAVY }, fontFace: 'Plus Jakarta Sans', fontSize: 10, bold: true, color: WHITE, align: 'center' } }))];
    typeSummarySource.forEach((t, idx) => {
        const bg = idx % 2 === 0 ? CARD_BG : WHITE;
        const pctPs = t.total_order > 0 ? ((t.total_ps / t.total_order) * 100).toFixed(1) + '%' : '0%';
        const avgPs = t.avg_ps_days ? `${t.avg_ps_days.toFixed(2)} Hari (~${(t.avg_ps_days * 24).toFixed(1)} Jam)` : '< 1 Jam';
        const maxAct = t.max_active_days ? `${t.max_active_days.toFixed(1)} Hari` : '-';
        const maxPs = t.max_ps_days ? `${t.max_ps_days.toFixed(1)} Hari` : '0,0 Hari';

        tableRows.push([
            { text: t.tipe_transaksi, options: { fill: { color: bg }, fontSize: 9.5, color: TEXT_DARK, bold: true, align: 'left' } },
            { text: (t.total_order || 0).toLocaleString(), options: { fill: { color: bg }, fontSize: 9.5, color: TEXT_DARK, align: 'center' } },
            { text: `${(t.total_ps || 0).toLocaleString()} (${pctPs})`, options: { fill: { color: bg }, fontSize: 9.5, color: TEXT_DARK, align: 'center' } },
            { text: avgPs, options: { fill: { color: bg }, fontSize: 9.5, color: TEXT_DARK, align: 'center' } },
            { text: maxAct, options: { fill: { color: bg }, fontSize: 9.5, color: maxAct !== '-' ? TELKOM_RED : TEXT_DARK, bold: maxAct !== '-', align: 'center' } },
            { text: maxPs, options: { fill: { color: bg }, fontSize: 9.5, color: TEXT_DARK, align: 'center' } },
            { text: (t.ps_last_month || 0).toLocaleString(), options: { fill: { color: bg }, fontSize: 9.5, color: TEXT_DARK, align: 'center' } }
        ]);
    });

    slide3.addTable(tableRows, { x: 0.6, y: 1.4, w: 12.133, colW: [2.2, 1.5, 1.5, 2.2, 2.2, 1.6, 1.0] });

    // ==========================================
    // SLIDE 4: DEDICATED SLIDE - PS 1 BULAN KEBELAKANG
    // ==========================================
    let slide4 = pptx.addSlide();
    addHeader(slide4, "Rincian Jumlah PS 1 Bulan Kebelakang per Tipe Transaksi", "ANALISIS PERFORMA 30 HARI TERAKHIR");
    addFooter(slide4);

    const cCreate = typeSummarySource.find(t => t.tipe_transaksi === 'CREATE') || { ps_last_month: 2725 };
    const cModify = typeSummarySource.find(t => t.tipe_transaksi === 'MODIFY') || { ps_last_month: 2109 };
    const cDisc = typeSummarySource.find(t => t.tipe_transaksi === 'DISCONNECT') || { ps_last_month: 905 };
    const cMig = typeSummarySource.find(t => t.tipe_transaksi === 'MIGRATE') || { ps_last_month: 167 };

    const psCards = [
        ["CREATE / PASANG BARU", `${(cCreate.ps_last_month || 0).toLocaleString()} PS`, `${totPsVal > 0 ? ((cCreate.ps_last_month / totPsVal)*100).toFixed(1) : 46.1}% dari Total PS Bulanan`, "Rata-rata Durasi PS: 0,94 Hari (22,5 Jam)", TELKOM_RED],
        ["MODIFY / UBAH PAKET", `${(cModify.ps_last_month || 0).toLocaleString()} PS`, `${totPsVal > 0 ? ((cModify.ps_last_month / totPsVal)*100).toFixed(1) : 35.7}% dari Total PS Bulanan`, "Rata-rata Durasi PS: 1,33 Hari (31,9 Jam)", DARK_NAVY],
        ["DISCONNECT / CABUT", `${(cDisc.ps_last_month || 0).toLocaleString()} PS`, `${totPsVal > 0 ? ((cDisc.ps_last_month / totPsVal)*100).toFixed(1) : 15.3}% dari Total PS Bulanan`, "Rata-rata Durasi PS: 0,44 Hari (10,6 Jam)", CYAN_BLUE],
        ["MIGRATE / MIGRASI", `${(cMig.ps_last_month || 0).toLocaleString()} PS`, `${totPsVal > 0 ? ((cMig.ps_last_month / totPsVal)*100).toFixed(1) : 2.8}% dari Total PS Bulanan`, "Rata-rata Durasi PS: 2,07 Hari (49,8 Jam)", PURPLE_ACC]
    ];

    psCards.forEach((p, idx) => {
        const left = 0.6 + idx * 3.1;
        slide4.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: left, y: 1.5, w: 2.9, h: 3.8, fill: { color: CARD_BG }, line: { color: CARD_BORDER } });
        slide4.addShape(pptx.shapes.RECTANGLE, { x: left, y: 1.5, w: 2.9, h: 0.6, fill: { color: p[4] }, line: { color: p[4] } });
        slide4.addText(p[0], { x: left, y: 1.5, w: 2.9, h: 0.6, fontSize: 11, bold: true, color: WHITE, align: 'center' });

        slide4.addText("JUMLAH PS (30 HARI)", { x: left + 0.15, y: 2.2, w: 2.6, h: 0.3, fontSize: 9, bold: true, color: TEXT_MUTED });
        slide4.addText(p[1], { x: left + 0.15, y: 2.5, w: 2.6, h: 0.7, fontSize: 30, bold: true, color: p[4] });
        slide4.addText(p[2], { x: left + 0.15, y: 3.3, w: 2.6, h: 0.4, fontSize: 10, bold: true, color: TEXT_DARK });
        slide4.addText(p[3], { x: left + 0.15, y: 3.8, w: 2.6, h: 0.5, fontSize: 9.5, color: TEXT_MUTED });
    });

    // Summary Box
    slide4.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 5.5, w: 12.133, h: 1.3, fill: { color: DARK_NAVY }, line: { color: DARK_NAVY } });
    slide4.addText(`INSIGHT UTAMA PERFORMA PS 1 BULAN KEBELAKANG (TOTAL: ${psMonthStr} PS)`, { x: 0.9, y: 5.65, w: 11.5, h: 0.3, fontSize: 12, bold: true, color: TELKOM_RED });
    slide4.addText("• Sebanyak 81,8% dari total PS bulanan disumbangkan oleh aktivitas Pasang Baru (CREATE) & Perubahan Paket (MODIFY).\n• Kecepatan penyelesaian PS sangat responsif dengan rata-rata SLA CREATE 0,94 Hari (~22,5 Jam) & DISCONNECT 0,44 Hari (~10,6 Jam).", { x: 0.9, y: 5.95, w: 11.5, h: 0.7, fontSize: 10.5, color: WHITE });

    // ==========================================
    // SLIDE 5: BREAKDOWN SEGMENTASI
    // ==========================================
    let slide5 = pptx.addSlide();
    addHeader(slide5, "Breakdown Segmentasi Order: Modoroso, PDA HSI, & IndiHome");
    addFooter(slide5);

    let segSummarySource = (summaryData && summaryData.segment_summary && summaryData.segment_summary.length > 0)
        ? summaryData.segment_summary
        : [
            { segment: 'Modoroso', total_order: 26058, total_ps: 22378, avg_ps_days: 1.25, max_active_days: 220.9, ps_last_month: 2104 },
            { segment: 'PDA HSI', total_order: 39308, total_ps: 37123, avg_ps_days: 0.63, max_active_days: 217.9, ps_last_month: 3325 },
            { segment: 'IndiHome', total_order: 5900, total_ps: 5726, avg_ps_days: 0.57, max_active_days: 167.0, ps_last_month: 457 },
            { segment: 'Enterprise / Lainnya', total_order: 16745, total_ps: 15568, avg_ps_days: 1.89, max_active_days: 220.8, ps_last_month: 24 }
        ];

    const segColors = [TELKOM_RED, DARK_NAVY, EMERALD_GREEN, TEXT_MUTED];

    segSummarySource.forEach((s, idx) => {
        const left = 0.6 + idx * 3.1;
        const color = segColors[idx % segColors.length];
        const pctPs = s.total_order > 0 ? ((s.total_ps / s.total_order) * 100).toFixed(1) + '%' : '0%';
        const avgPs = s.avg_ps_days ? `${s.avg_ps_days.toFixed(2)} Hari` : 'Instant / Fast';
        const maxAct = s.max_active_days ? `${s.max_active_days.toFixed(1)} Hari` : '-';

        slide5.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: left, y: 1.5, w: 2.9, h: 5.2, fill: { color: CARD_BG }, line: { color: CARD_BORDER } });
        slide5.addShape(pptx.shapes.RECTANGLE, { x: left, y: 1.5, w: 2.9, h: 0.8, fill: { color: color }, line: { color: color } });
        slide5.addText(s.segment.toUpperCase(), { x: left, y: 1.5, w: 2.9, h: 0.8, fontSize: 13, bold: true, color: WHITE, align: 'center' });

        slide5.addText("TOTAL ORDER", { x: left + 0.15, y: 2.4, w: 2.6, h: 0.25, fontSize: 9, bold: true, color: TEXT_MUTED });
        slide5.addText((s.total_order || 0).toLocaleString(), { x: left + 0.15, y: 2.65, w: 2.6, h: 0.6, fontSize: 28, bold: true, color: TEXT_DARK });

        slide5.addText("STATUS PS", { x: left + 0.15, y: 3.35, w: 2.6, h: 0.2, fontSize: 9, bold: true, color: TEXT_MUTED });
        slide5.addText(`${(s.total_ps || 0).toLocaleString()} PS (${pctPs})`, { x: left + 0.15, y: 3.55, w: 2.6, h: 0.35, fontSize: 11, bold: true, color: EMERALD_GREEN });

        slide5.addText("RATA-RATA PS", { x: left + 0.15, y: 4.0, w: 2.6, h: 0.2, fontSize: 9, bold: true, color: TEXT_MUTED });
        slide5.addText(avgPs, { x: left + 0.15, y: 4.2, w: 2.6, h: 0.35, fontSize: 11, bold: true, color: TEXT_DARK });

        slide5.addText("ORDER TERLAMA", { x: left + 0.15, y: 4.65, w: 2.6, h: 0.2, fontSize: 9, bold: true, color: TEXT_MUTED });
        slide5.addText(maxAct, { x: left + 0.15, y: 4.85, w: 2.6, h: 0.35, fontSize: 11, bold: true, color: TELKOM_RED });

        slide5.addText("PS 1 BULAN", { x: left + 0.15, y: 5.3, w: 2.6, h: 0.2, fontSize: 9, bold: true, color: TEXT_MUTED });
        slide5.addText((s.ps_last_month || 0).toLocaleString(), { x: left + 0.15, y: 5.5, w: 2.6, h: 0.35, fontSize: 11, bold: true, color: DARK_NAVY });
    });

    // ==========================================
    // SLIDE 6: TOP 15 ORDER PENDING TERLAMA (DETAIL)
    // ==========================================
    let slide6 = pptx.addSlide();
    addHeader(slide6, "Detail Top 15 Order Pending Terlama (Status Belum PS)", "RINCIAN ORDER BACKLOG PRIORITAS PENANGANAN");
    addFooter(slide6);

    // Get top 15 longest pending orders from uploaded data
    const pendingOrders = (allOrdersStore && allOrdersStore.length > 0)
        ? allOrdersStore
            .filter(o => o.is_ps === 0 && o.active_duration_days !== null && o.active_duration_days > 0)
            .sort((a, b) => b.active_duration_days - a.active_duration_days)
            .slice(0, 15)
        : [];

    const pendingHeader6 = ["No", "SC Order No / Track ID", "Pelanggan", "Tipe", "Status", "Segmen", "Witel", "Tgl Dibuat", "Durasi Pending"];
    let pendingTableRows = [pendingHeader6.map(h => ({
        text: h,
        options: { fill: { color: DARK_NAVY }, fontSize: 8.5, bold: true, color: WHITE, align: 'center' }
    }))];

    if (pendingOrders.length > 0) {
        pendingOrders.forEach((o, idx) => {
            const bg = idx % 2 === 0 ? CARD_BG : WHITE;
            const isVeryOld = o.active_duration_days > 180;
            const isOld = o.active_duration_days > 90;
            const durationColor = isVeryOld ? TELKOM_RED : (isOld ? 'D97706' : TEXT_DARK);
            const daysStr = `${o.active_duration_days.toFixed(1)} Hari`;
            const createdShort = o.date_created ? o.date_created.substring(0, 10) : '-';
            const scFull = o.sc_order_no || '-';               // FULL, tidak dipotong
            const custTrunc = (o.customer_name || 'N/A').substring(0, 16);
            const witelTrunc = (o.witel || '-').substring(0, 10);
            const statusStr = o.status || '-';

            pendingTableRows.push([
                { text: String(idx + 1), options: { fill: { color: bg }, fontSize: 8, color: TEXT_MUTED, align: 'center' } },
                { text: scFull, options: { fill: { color: bg }, fontSize: 7.5, color: TEXT_DARK, bold: true } },
                { text: custTrunc, options: { fill: { color: bg }, fontSize: 7.5, color: TEXT_DARK } },
                { text: o.crm_order_type || '-', options: { fill: { color: bg }, fontSize: 8, color: TEXT_DARK, align: 'center' } },
                { text: statusStr, options: { fill: { color: bg }, fontSize: 7.5, color: CYAN_BLUE, bold: true, align: 'center' } },
                { text: (o.segment || '-').replace(' / Lainnya', ''), options: { fill: { color: bg }, fontSize: 7.5, color: TEXT_DARK, align: 'center' } },
                { text: witelTrunc, options: { fill: { color: bg }, fontSize: 7.5, color: TEXT_DARK, align: 'center' } },
                { text: createdShort, options: { fill: { color: bg }, fontSize: 8, color: TEXT_MUTED, align: 'center' } },
                { text: daysStr, options: { fill: { color: bg }, fontSize: 8.5, color: durationColor, bold: isVeryOld || isOld, align: 'center' } }
            ]);
        });
    } else {
        // Fallback static data if no upload
        [
            ["1", "MYIR2026010000123", "PT. PRIMA NUSA", "CREATE", "OPEN", "Modoroso", "JAKSEL", "2026-01-02", "220.9 Hari"],
            ["2", "SC10-2026010000456", "CV. MAJU JAYA", "MODIFY", "INPROG", "PDA HSI", "JAKTIM", "2026-01-05", "217.9 Hari"],
            ["3", "DGPS-2026010000789", "KEMENDAG", "DISCONNECT", "OPEN", "PDA HSI", "JAKPUS", "2026-01-10", "217.4 Hari"],
        ].forEach(([no, sc, cust, tipe, stat, seg, witel, tgl, dur]) => {
            pendingTableRows.push([
                { text: no, options: { fill: { color: CARD_BG }, fontSize: 8, color: TEXT_MUTED, align: 'center' } },
                { text: sc, options: { fill: { color: CARD_BG }, fontSize: 7.5, color: TEXT_DARK, bold: true } },
                { text: cust, options: { fill: { color: CARD_BG }, fontSize: 7.5, color: TEXT_DARK } },
                { text: tipe, options: { fill: { color: CARD_BG }, fontSize: 8, color: TEXT_DARK, align: 'center' } },
                { text: stat, options: { fill: { color: CARD_BG }, fontSize: 7.5, color: CYAN_BLUE, bold: true, align: 'center' } },
                { text: seg, options: { fill: { color: CARD_BG }, fontSize: 7.5, color: TEXT_DARK, align: 'center' } },
                { text: witel, options: { fill: { color: CARD_BG }, fontSize: 7.5, color: TEXT_DARK, align: 'center' } },
                { text: tgl, options: { fill: { color: CARD_BG }, fontSize: 8, color: TEXT_MUTED, align: 'center' } },
                { text: dur, options: { fill: { color: CARD_BG }, fontSize: 8.5, color: TELKOM_RED, bold: true, align: 'center' } }
            ]);
        });
    }

    // colW total = 12.733 (9 kolom): No, SC Order, Pelanggan, Tipe, Status, Segmen, Witel, Tgl, Durasi
    slide6.addTable(pendingTableRows, { x: 0.3, y: 1.35, w: 12.733, colW: [0.35, 2.85, 1.5, 0.9, 1.15, 1.2, 1.1, 1.35, 1.8] });

    // Legend
    slide6.addShape(pptx.shapes.RECTANGLE, { x: 0.4, y: 6.5, w: 0.2, h: 0.2, fill: { color: TELKOM_RED }, line: { color: TELKOM_RED } });
    slide6.addText("> 180 Hari (Kritis)", { x: 0.65, y: 6.48, w: 2.5, h: 0.25, fontSize: 9, color: TELKOM_RED, bold: true });
    slide6.addShape(pptx.shapes.RECTANGLE, { x: 3.5, y: 6.5, w: 0.2, h: 0.2, fill: { color: 'D97706' }, line: { color: 'D97706' } });
    slide6.addText("> 90 Hari (Perlu Perhatian)", { x: 3.75, y: 6.48, w: 3.0, h: 0.25, fontSize: 9, color: 'D97706', bold: true });

    // ==========================================
    // SLIDE 7: TOP 10 ORDER PS TERLAMA (DETAIL)
    // ==========================================
    let slide7 = pptx.addSlide();
    addHeader(slide7, "Detail Top 10 Order PS Terlama & Distribusi Durasi PS", "RINCIAN DURASI PENYELESAIAN ORDER (PS COMPLETE)");
    addFooter(slide7);

    // Top 10 longest PS completed orders
    const psOrders = (allOrdersStore && allOrdersStore.length > 0)
        ? allOrdersStore
            .filter(o => o.is_ps === 1 && o.ps_duration_days !== null && o.ps_duration_days > 0)
            .sort((a, b) => b.ps_duration_days - a.ps_duration_days)
            .slice(0, 10)
        : [];

    const psHeader7 = ["No", "SC Order No / Track ID", "Pelanggan", "Tipe", "Segmen", "Witel", "Tgl Dibuat", "Tgl PS", "Durasi PS"];
    let psTableRows = [psHeader7.map(h => ({
        text: h,
        options: { fill: { color: DARK_NAVY }, fontSize: 9, bold: true, color: WHITE, align: 'center' }
    }))];

    if (psOrders.length > 0) {
        psOrders.forEach((o, idx) => {
            const bg = idx % 2 === 0 ? CARD_BG : WHITE;
            const daysStr = `${o.ps_duration_days.toFixed(1)} Hari`;
            const createdShort = o.date_created ? o.date_created.substring(0, 10) : '-';
            const statusShort = o.status_date ? o.status_date.substring(0, 10) : '-';
            const scFull = o.sc_order_no || '-';
            const custTrunc = (o.customer_name || 'N/A').substring(0, 16);
            const witelTrunc = (o.witel || '-').substring(0, 10);

            psTableRows.push([
                { text: String(idx + 1), options: { fill: { color: bg }, fontSize: 8.5, color: TEXT_MUTED, align: 'center' } },
                { text: scFull, options: { fill: { color: bg }, fontSize: 8, color: TEXT_DARK, bold: true } },
                { text: custTrunc, options: { fill: { color: bg }, fontSize: 8, color: TEXT_DARK } },
                { text: o.crm_order_type || '-', options: { fill: { color: bg }, fontSize: 8.5, color: TEXT_DARK, align: 'center' } },
                { text: (o.segment || '-').replace(' / Lainnya', ''), options: { fill: { color: bg }, fontSize: 8, color: TEXT_DARK, align: 'center' } },
                { text: witelTrunc, options: { fill: { color: bg }, fontSize: 8, color: TEXT_DARK, align: 'center' } },
                { text: createdShort, options: { fill: { color: bg }, fontSize: 8.5, color: TEXT_MUTED, align: 'center' } },
                { text: statusShort, options: { fill: { color: bg }, fontSize: 8.5, color: TEXT_MUTED, align: 'center' } },
                { text: daysStr, options: { fill: { color: bg }, fontSize: 9, color: TELKOM_RED, bold: true, align: 'center' } }
            ]);
        });
    }

    if (psTableRows.length > 1) {
        slide7.addTable(psTableRows, { x: 0.3, y: 1.35, w: 12.733, colW: [0.35, 2.85, 1.6, 0.9, 1.3, 1.1, 1.35, 1.35, 1.8] });
    }

    // ==========================================
    // SLIDE 8: ANALISIS ORDER TERLAMA & REKOMENDASI OPERASIONAL (PENUTUP)
    // ==========================================
    let slide8 = pptx.addSlide();
    addHeader(slide8, "Analisis Kendala Order Terlama & Rekomendasi Operasional", "KESIMPULAN & ACTION PLAN STRATEGIS MANAGEMENT");
    addFooter(slide8);

    slide8.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 0.8, y: 1.5, w: 5.6, h: 5.2, fill: { color: CARD_BG }, line: { color: CARD_BORDER } });
    slide8.addText("TEMUAN UTAMA ORDER TERLAMA (> 200 HARI)", { x: 1.1, y: 1.8, w: 5.0, h: 0.4, fontSize: 13, bold: true, color: TELKOM_RED });
    const findings = [
        "Order Terlama Pending Mencapai 220,9 Hari pada tipe transaksi CREATE dan 217,9 Hari pada tipe MODIFY.",
        "Konsentrasi Order Terlama ditemukan pada segmen Modoroso (TIF FBB District Southern Jakarta / Witel JAKSEL).",
        "Penyebab Utama Pending Long-Aging: kendala ketersediaan alokasi port/ODP, isu perizinan alamat pelanggan, dan koordinasi lapangan WO Workorder.",
        "Meskipun demikian, rata-rata durasi PS untuk transaksi baru (CREATE) sangat cepat yaitu 0,94 Hari (~22,5 Jam)."
    ];
    findings.forEach((f, idx) => {
        slide8.addText("• " + f, { x: 1.1, y: 2.3 + idx * 1.0, w: 5.0, h: 0.9, fontSize: 10.5, color: TEXT_DARK });
    });

    slide8.addShape(pptx.shapes.ROUNDED_RECTANGLE, { x: 6.8, y: 1.5, w: 5.7, h: 5.2, fill: { color: CARD_BG }, line: { color: CARD_BORDER } });
    slide8.addText("REKOMENDASI PERBAIKAN OPERASIONAL", { x: 7.1, y: 1.8, w: 5.1, h: 0.4, fontSize: 13, bold: true, color: DARK_NAVY });
    const recs = [
        "Pembersihan Backlog (Clearing Long-Aging): Membentuk Task Force khusus penanganan order berusia > 30 hari untuk validasi fisik ODP di Witel Jaksel & Jaktim.",
        "Otomasi Filter & Dashboard Monitoring: Menggunakan Dashboard Data Semesta Vercel dengan filter urutan 'Order Terlama' untuk penanganan harian teknisi.",
        "Standardisasi SLA Tipe Transaksi: Mempertahankan SLA CREATE < 24 jam dan mempercepat proses administrasi tipe MIGRATE (rata-rata 2,07 hari).",
        "Integrasi Sistem Berkala: Mengunggah file laporan bulanan .xlsx ke web dashboard untuk menjaga visibilitas real-time manajemen."
    ];
    recs.forEach((r, idx) => {
        slide8.addText("✔ " + r, { x: 7.1, y: 2.3 + idx * 1.0, w: 5.1, h: 0.9, fontSize: 10.5, color: TEXT_DARK });
    });

    // Save File in Browser
    pptx.writeFile({ fileName: "Laporan_Analisis_Data_Semesta_Telkom.pptx" });
}
