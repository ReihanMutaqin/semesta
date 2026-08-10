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

    window.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') closeDetailModal();
    });

    document.getElementById('detail-modal').addEventListener('click', (e) => {
        if (e.target.id === 'detail-modal') closeDetailModal();
    });

    const searchInput = document.getElementById('filter-search');
    searchInput.addEventListener('input', () => {
        clearTimeout(searchDebounce);
        searchDebounce = setTimeout(() => {
            currentPage = 1;
            loadOrdersData();
        }, 300);
    });

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
        // Strategy 1: BinaryString (Most Compatible with Telkom CRM / zlib ZIP64 exports)
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
                
                // Strategy 2: ArrayBuffer fallback
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

// Client-Side Excel File Upload & Processing
async function handleClientFileUpload(file) {
    const modal = document.getElementById('upload-modal');
    const titleEl = document.getElementById('upload-modal-title');
    const descEl = document.getElementById('upload-modal-desc');
    const statusEl = document.getElementById('upload-modal-status');
    const progressBar = document.getElementById('upload-progress-bar');

    modal.classList.remove('hidden');
    titleEl.textContent = 'Membaca & Memproses Excel...';
    descEl.textContent = `Mengurai file: ${file.name} (${(file.size / (1024 * 1024)).toFixed(2)} MB)...`;

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

    const processedOrders = rows.map((r) => {
        const dateCreated = parseExcelDate(r['Date Created']);
        const statusDate = parseExcelDate(r['Status Date']);
        const dateModified = parseExcelDate(r['Date Modified']);
        const schedStart = parseExcelDate(r['Sched Start']);
        const bookingDate = parseExcelDate(r['Booking Date']);
        const measurementDate = parseExcelDate(r['Measurement Date']);

        if (dateCreated && dateCreated.getTime() > maxDateMs) {
            maxDateMs = dateCreated.getTime();
            maxDateStr = formatDateStr(dateCreated);
        }
        if (statusDate && statusDate.getTime() > maxDateMs) {
            maxDateMs = statusDate.getTime();
            maxDateStr = formatDateStr(statusDate);
        }

        const rawType = (r['CRM Order Type'] || 'UNSPECIFIED').toString().trim().toUpperCase();
        let crmType = rawType;
        if (rawType === 'NEW INSTALL') crmType = 'CREATE';

        const statusRaw = (r['Status'] || 'UNKNOWN').toString().trim().toUpperCase();
        const psStatuses = ['COMPLETE', 'COMPWORK', 'INSTCOMP', 'DEINSTCOMP', 'ACTCOMP', 'VALCOMP'];
        const isPs = psStatuses.includes(statusRaw) ? 1 : 0;

        let psDurationDays = null;
        if (isPs === 1 && dateCreated && statusDate && statusDate >= dateCreated) {
            psDurationDays = (statusDate.getTime() - dateCreated.getTime()) / 86400000.0;
        }

        let activeDurationDays = null;
        if (isPs === 0 && dateCreated && maxDateMs > 0) {
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

    const oneMonthAgoMs = maxDateMs - (30 * 86400000);

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

    const summary = {
        max_date: maxDateStr,
        one_month_ago: formatDateStr(new Date(oneMonthAgoMs)),
        total_order_semesta: totalOrder,
        total_ps: totalPs,
        ps_percentage: totalOrder > 0 ? (totalPs / totalOrder) * 100.0 : 0.0,
        total_ps_last_month: psLastMonth,
        type_summary: typeSummary,
        segment_summary: segSummary,
        daily_trend: []
    };

    return { summary, orders: processedOrders };
}

function parseExcelDate(val) {
    if (!val) return null;
    if (val instanceof Date) return val;
    if (typeof val === 'number') {
        // Excel serial date number conversion
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
    const d = new Date(val);
    if (!isNaN(d.getTime())) return d;
    return null;
}

function formatDateStr(d) {
    if (!d || isNaN(d.getTime())) return '-';
    return d.toISOString().replace('T', ' ').substring(0, 19);
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
        tr.title = `Klik untuk memfilter data order tipe ${t.tipe_transaksi}`;

        const avgPsText = t.avg_ps_days 
            ? `${t.avg_ps_days.toFixed(2)} hari (${t.avg_ps_hours.toFixed(1)} jam)`
            : '-';

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
                label: 'Rata-rata Durasi PS (Jam)',
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

    const trendData = data.daily_trend || [];
    chartTrend = new Chart(ctx4, {
        type: 'line',
        data: {
            labels: trendData.map(d => d.date_key),
            datasets: [{
                label: 'Total Order Masuk',
                data: trendData.map(d => d.total_order),
                borderColor: '#818cf8',
                backgroundColor: 'rgba(129, 140, 248, 0.1)',
                fill: true,
                tension: 0.3
            }, {
                label: 'Order PS',
                data: trendData.map(d => d.ps_count),
                borderColor: '#34d399',
                backgroundColor: 'rgba(52, 211, 153, 0.1)',
                fill: true,
                tension: 0.3
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: { labels: { color: '#94a3b8' } }
            },
            scales: {
                x: { ticks: { color: '#94a3b8', maxTicksLimit: 10 }, grid: { color: 'rgba(255,255,255,0.05)' } },
                y: { ticks: { color: '#94a3b8' }, grid: { color: 'rgba(255,255,255,0.05)' } }
            }
        }
    });
}

// Load Orders Data (Works in Memory for Uploaded Files or Server API)
async function loadOrdersData() {
    const search = (document.getElementById('filter-search').value || '').toLowerCase().trim();
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
