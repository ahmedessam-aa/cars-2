/* =====================================================================
   استلام السيارات + الخصومات والجزاءات + ملخصات الأقسام
   (مدموج من نظام "إقرار استلام سيارة" في سيستم المركبات)
   ===================================================================== */

let editingHandoverId = null;
let editingPenaltyId = null;

const byId = (id) => document.getElementById(id);
const numVal = (v) => Number(v) || 0;
const moneyFmt = (v) => numVal(v).toLocaleString('en-US');
const toLatinDigits = (s) => String(s || '').replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d));
const fmtArDate = (d) => (d ? new Date(d + 'T00:00:00').toLocaleDateString('ar-EG') : '');

// عنصر الفورم -> اسم الحقل في السجل
const HO_FIELDS = {
    hoVehicleType: 'vehicle_type', hoVehicleName: 'vehicle_name', hoChassis: 'chassis_number',
    hoEngine: 'engine_number', hoColor: 'color', hoYear: 'manufacture_year',
    hoLicStart: 'license_start', hoLicEnd: 'license_expiry', hoAccessories: 'accessories',
    hoDelegate: 'delegate_name', hoJob: 'job_title', hoDept: 'department',
    hoNationalId: 'national_id', hoDate: 'delivery_date', hoSigner: 'signature_name',
    hoResponsible: 'responsible_name',
    hoFingerprint: 'fingerprint', hoNotes: 'notes'
};
const PN_FIELDS = {
    pnDate: 'penalty_date', pnEmployee: 'employee_name', pnDept: 'department',
    pnType: 'penalty_type', pnAmount: 'penalty_amount', pnIncidentDate: 'incident_date',
    pnReason: 'reason'
};

/* ---------------------------------------------------------------------
   بيانات الاستلام لكل مركبة
   --------------------------------------------------------------------- */
function handoverSort(a, b) {
    if (a.delivery_date !== b.delivery_date) return (b.delivery_date || '').localeCompare(a.delivery_date || '');
    return (Number(b.id) || 0) - (Number(a.id) || 0);
}

function handoverBelongsTo(h, vehicle) {
    return (h.vehicle_id && h.vehicle_id === vehicle.id) ||
        (!!h.plate_number && h.plate_number === vehicle.plate_number);
}

function getVehicleHandovers(vehicle) {
    return appData.handovers.filter(h => handoverBelongsTo(h, vehicle)).sort(handoverSort);
}

function getLastHandover(vehicle) {
    return getVehicleHandovers(vehicle)[0] || null;
}

function lastReceiverName(vehicle) {
    const h = getLastHandover(vehicle);
    return h ? (h.delegate_name || '') : '';
}

function lastReceiveDate(vehicle) {
    const h = getLastHandover(vehicle);
    return h ? (h.delivery_date || '') : '';
}

// بعد حفظ/حذف إقرار استلام: بيانات المركبة (السائق الحالي، رقم الشاسيه، رقم الموتور،
// بداية وانتهاء الرخصة) بتتحدث تلقائيًا من آخر إقرار استلام مسجّل للمركبة دي،
// عشان محتاجش تعدّل على بيانات المركبة يدويًا كل مرة.
function syncVehicleFromHandover(vehicleId) {
    if (!vehicleId) return;
    const vehicle = appData.vehicles.find(v => v.id === vehicleId);
    if (!vehicle) return;
    const latest = getLastHandover(vehicle);
    if (!latest) return;
    if (latest.delegate_name) vehicle.vin_number = latest.delegate_name;
    if (latest.chassis_number) vehicle.chassis_number = latest.chassis_number;
    if (latest.engine_number) vehicle.engine_number = latest.engine_number;
    if (latest.license_start) vehicle.license_start = latest.license_start;
    if (latest.license_expiry) vehicle.license_expiry = latest.license_expiry;
}

function fleetBadgeClass(status) {
    if (status === 'اخضر') return 'badge-green';
    if (status === 'مطحون') return 'badge-roasted';
    if (status === 'نقل موظفين') return 'badge-staff';
    if (status === 'ملاكي') return 'badge-private';
    return 'status-inactive';
}

/* ---------------------------------------------------------------------
   الملخصات (بتظهر فوق كل قسم وفي التقارير)
   --------------------------------------------------------------------- */
function licenseState(vehicle) {
    const days = Math.ceil((new Date(vehicle.license_expiry) - new Date()) / 86400000);
    if (days <= 0) return 'expired';
    if (days <= licenseWarningDays) return 'expiring';
    return 'valid';
}

function getSummary(kind) {
    const V = appData.vehicles;
    const countStatus = (s) => V.filter(v => v.status === s).length;

    if (kind === 'fleet' || kind === 'vehicles') {
        const items = [
            { label: 'إجمالي المركبات', value: V.length, cls: 'chip-total', filter: ['vehicleStatusFilter', '', 'vehicles'] },
            { label: 'أخضر', value: countStatus('اخضر'), cls: 'chip-green', filter: ['vehicleStatusFilter', 'اخضر', 'vehicles'] },
            { label: 'مطحون', value: countStatus('مطحون'), cls: 'chip-roasted', filter: ['vehicleStatusFilter', 'مطحون', 'vehicles'] },
            { label: 'نقل موظفين', value: countStatus('نقل موظفين'), cls: 'chip-staff', filter: ['vehicleStatusFilter', 'نقل موظفين', 'vehicles'] },
            { label: 'ملاكي', value: countStatus('ملاكي'), cls: 'chip-private', filter: ['vehicleStatusFilter', 'ملاكي', 'vehicles'] }
        ];
        const other = V.filter(v => !['اخضر', 'مطحون', 'نقل موظفين', 'ملاكي'].includes(v.status)).length;
        if (other) items.push({ label: 'حالات أخرى', value: other, cls: 'chip-total' });
        if (kind === 'vehicles') {
            items.push({ label: 'رخص منتهية', value: V.filter(v => licenseState(v) === 'expired').length, cls: 'chip-danger', filter: ['licenseStatusFilter', 'منتهية', 'licenses'] });
            items.push({ label: 'رخص قريبة الانتهاء', value: V.filter(v => licenseState(v) === 'expiring').length, cls: 'chip-warn', filter: ['licenseStatusFilter', 'قريبة الانتهاء', 'licenses'] });
        }
        return items;
    }

    if (kind === 'maintenance') {
        const M = appData.maintenance;
        const now = new Date();
        const monthCost = M.filter(m => {
            const d = new Date(m.maintenance_date);
            return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
        }).reduce((s, m) => s + numVal(m.cost), 0);
        return [
            { label: 'إجمالي العمليات', value: M.length, cls: 'chip-total', filter: ['maintenanceStatusFilter', ''] },
            { label: 'قيد الإجراء', value: M.filter(m => m.status === 'قيد الإجراء').length, cls: 'chip-warn', filter: ['maintenanceStatusFilter', 'قيد الإجراء'] },
            { label: 'مكتملة', value: M.filter(m => m.status === 'مكتملة').length, cls: 'chip-green', filter: ['maintenanceStatusFilter', 'مكتملة'] },
            { label: 'إجمالي التكلفة (جنيه)', value: moneyFmt(M.reduce((s, m) => s + numVal(m.cost), 0)), cls: 'chip-total' },
            { label: 'تكلفة هذا الشهر (جنيه)', value: moneyFmt(monthCost), cls: 'chip-staff' }
        ];
    }

    if (kind === 'violations') {
        const X = appData.violations;
        const pending = X.filter(v => v.status === 'معلقة');
        return [
            { label: 'إجمالي المخالفات', value: X.length, cls: 'chip-total', filter: ['violationStatusFilter', ''] },
            { label: 'معلقة', value: pending.length, cls: 'chip-danger', filter: ['violationStatusFilter', 'معلقة'] },
            { label: 'مسددة', value: X.filter(v => v.status === 'مسددة').length, cls: 'chip-green', filter: ['violationStatusFilter', 'مسددة'] },
            { label: 'مبلغ المعلق (جنيه)', value: moneyFmt(pending.reduce((s, v) => s + numVal(v.amount), 0)), cls: 'chip-warn' },
            { label: 'إجمالي المبالغ (جنيه)', value: moneyFmt(X.reduce((s, v) => s + numVal(v.amount), 0)), cls: 'chip-total' }
        ];
    }

    if (kind === 'licenses') {
        return [
            { label: 'إجمالي الرخص', value: V.length, cls: 'chip-total', filter: ['licenseStatusFilter', ''] },
            { label: 'سارية', value: V.filter(v => licenseState(v) === 'valid').length, cls: 'chip-green', filter: ['licenseStatusFilter', 'سارية'] },
            { label: 'قريبة الانتهاء', value: V.filter(v => licenseState(v) === 'expiring').length, cls: 'chip-warn', filter: ['licenseStatusFilter', 'قريبة الانتهاء'] },
            { label: 'منتهية', value: V.filter(v => licenseState(v) === 'expired').length, cls: 'chip-danger', filter: ['licenseStatusFilter', 'منتهية'] }
        ];
    }

    if (kind === 'tires') {
        const stock = getTireStock();
        return [
            { label: 'أنواع الكوتش', value: stock.length, cls: 'chip-total' },
            { label: 'إجمالي المضاف', value: stock.reduce((s, x) => s + x.added, 0), cls: 'chip-green' },
            { label: 'إجمالي المصروف', value: stock.reduce((s, x) => s + x.issued, 0), cls: 'chip-warn' },
            { label: 'الرصيد الحالي', value: stock.reduce((s, x) => s + x.balance, 0), cls: 'chip-staff' }
        ];
    }

    if (kind === 'movements') {
        const T = appData.movements;
        const movedVehicles = new Set(T.map(m => getRecordPlate(m)).filter(Boolean));
        return [
            { label: 'عدد التحركات', value: T.length, cls: 'chip-total' },
            { label: 'سيارات اتحركت', value: movedVehicles.size, cls: 'chip-staff' },
            { label: 'إجمالي الكيلومترات', value: moneyFmt(T.reduce((s, m) => s + numVal(m.km), 0)), cls: 'chip-green' },
            { label: 'إجمالي مدة التشغيل', value: formatDuration(T.reduce((s, m) => s + getMovementMinutes(m), 0)), cls: 'chip-warn' }
        ];
    }

    if (kind === 'expenses') {
        const E = appData.expenses;
        const sumGroup = (g) => E.filter(e => getRecordFleetGroup(e) === g).reduce((s, e) => s + numVal(e.amount), 0);
        return [
            { label: 'عدد النفقات', value: E.length, cls: 'chip-total' },
            { label: 'إجمالي النفقات (جنيه)', value: moneyFmt(E.reduce((s, e) => s + numVal(e.amount), 0)), cls: 'chip-staff' },
            { label: 'نفقات البن الأخضر (جنيه)', value: moneyFmt(sumGroup('green')), cls: 'chip-green' },
            { label: 'نفقات البن المطحون (جنيه)', value: moneyFmt(sumGroup('roasted')), cls: 'chip-roasted' }
        ];
    }

    if (kind === 'handovers') {
        const H = appData.handovers;
        const now = new Date();
        const monthKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
        const withReceipt = V.filter(v => getLastHandover(v)).length;
        return [
            { label: 'إجمالي الإقرارات', value: H.length, cls: 'chip-total' },
            { label: 'مركبات ليها استلام', value: withReceipt, cls: 'chip-green' },
            { label: 'مركبات بدون استلام مسجل', value: V.length - withReceipt, cls: 'chip-warn' },
            { label: 'إقرارات هذا الشهر', value: H.filter(h => (h.delivery_date || '').startsWith(monthKey)).length, cls: 'chip-staff' },
            { label: 'خصومات وجزاءات', value: appData.penalties.length, cls: 'chip-danger', filter: ['handoverTabPenalties', '', '', 'penalties'] }
        ];
    }

    return [];
}

const SUMMARY_TITLES = {
    vehicles: 'المركبات', maintenance: 'الصيانة', violations: 'المخالفات', licenses: 'الرخص',
    tires: 'مخزن الكوتش', movements: 'تحركات السيارات', expenses: 'النفقات', handovers: 'استلام السيارات'
};

const SUMMARY_TARGETS = {
    fleet: 'dashboardFleetSummary', vehicles: 'vehiclesSummary', maintenance: 'maintenanceSummary',
    violations: 'violationsSummary', licenses: 'licensesSummary', tires: 'tiresSummary',
    movements: 'movementsSummary', expenses: 'expensesSummary', handovers: 'handoversSummary'
};

function renderSummaryStrip(containerId, items) {
    const el = byId(containerId);
    if (!el) return;
    el.innerHTML = items.map(i => {
        const clickable = i.filter
            ? ` onclick="quickFilter('${i.filter[0]}','${i.filter[1]}','${i.filter[2] || ''}','${i.filter[3] || ''}')" role="button" tabindex="0"`
            : '';
        return `<div class="summary-chip ${i.cls || ''}${i.filter ? ' clickable' : ''}"${clickable}>
            <span class="chip-value">${i.value}</span><span class="chip-label">${i.label}</span></div>`;
    }).join('');
}

function renderAllSummaries() {
    try {
        Object.keys(SUMMARY_TARGETS).forEach(kind => renderSummaryStrip(SUMMARY_TARGETS[kind], getSummary(kind)));
    } catch (error) {
        console.error('خطأ في تحديث الملخصات:', error);
    }
}

// الضغط على مربع ملخص بيفلتر الجدول (وبينقلك للقسم لو محتاج)
function quickFilter(selectId, value, section, tab) {
    if (section) showSection(section);
    if (tab) { showSection('handovers'); showHandoverTab(tab); return; }
    const el = byId(selectId);
    if (!el) return;
    el.value = value;
    el.dispatchEvent(new Event('change'));
}

// نسخة للتقارير (HTML) والإكسل
function reportChips(items) {
    return `<div style="display:flex;flex-wrap:wrap;gap:8px;margin:10px 0 14px 0;">${items.map(i =>
        `<span style="border:1px solid #cfd8dc;background:#f4f8fb;padding:5px 12px;border-radius:6px;font-size:12px;"><b style="font-size:14px;color:#22313a;">${i.value}</b> ${i.label}</span>`
    ).join('')}</div>`;
}

function summaryRows(kind) {
    return [['ملخص'], ...getSummary(kind).map(i => [i.label, i.value])];
}

function allSummaryRows() {
    const rows = [['ملخص النظام'], ['التاريخ: ' + new Date().toLocaleDateString('ar-EG')], []];
    Object.keys(SUMMARY_TITLES).forEach(kind => {
        rows.push([SUMMARY_TITLES[kind]]);
        getSummary(kind).forEach(i => rows.push([i.label, i.value]));
        rows.push([]);
    });
    return rows;
}

/* ---------------------------------------------------------------------
   تبويبات القسم + الفلاتر
   --------------------------------------------------------------------- */
function showHandoverTab(tab) {
    document.querySelectorAll('#handovers .tab-btn').forEach(b => b.classList.toggle('active', b.dataset.tab === tab));
    byId('handoverTabReceipts').classList.toggle('active', tab === 'receipts');
    byId('handoverTabPenalties').classList.toggle('active', tab === 'penalties');
}

function populateHandovers() {
    fillVehicleSelect('handoverVehicleFilter', 'كل السيارات');
    filterHandovers();
    filterPenalties();
}

function getFilteredHandovers() {
    const q = (byId('handoverSearch').value || '').trim().toLowerCase();
    const vehicleId = byId('handoverVehicleFilter').value;
    const from = byId('handoverFrom').value;
    const to = byId('handoverTo').value;
    const vehicle = vehicleId ? appData.vehicles.find(v => v.id === vehicleId) : null;

    return appData.handovers.filter(h => {
        if (vehicleId && !(vehicle && handoverBelongsTo(h, vehicle))) return false;
        if ((from || to) && !dateInRange(h.delivery_date, from, to)) return false;
        if (q) {
            const hay = [h.plate_number, h.delegate_name, h.signature_name, h.job_title, h.department, h.vehicle_name]
                .join(' ').toLowerCase();
            if (!hay.includes(q)) return false;
        }
        return true;
    }).sort(handoverSort);
}

function filterHandovers() {
    const tbody = byId('handoversTable');
    if (!tbody) return;
    const rows = getFilteredHandovers();
    if (!rows.length) {
        tbody.innerHTML = '<tr class="empty-row"><td colspan="10">لا توجد إقرارات استلام</td></tr>';
        return;
    }
    tbody.innerHTML = rows.map(h => `
        <tr>
            <td>${escapeHtml(h.delivery_date) || '-'}</td>
            <td>${escapeHtml(h.plate_number)}</td>
            <td>${escapeHtml(h.vehicle_name) || '-'}</td>
            <td>${escapeHtml(h.delegate_name) || '-'}</td>
            <td>${escapeHtml([h.job_title, h.department].filter(Boolean).join(' - ')) || '-'}</td>
            <td><strong>${escapeHtml(h.signature_name) || '-'}</strong></td>
            <td>${escapeHtml(h.responsible_name) || '-'}</td>
            <td>${escapeHtml(h.fingerprint) || '-'}</td>
            <td>${escapeHtml(h.notes) || '-'}</td>
            <td>
                <div class="action-buttons">
                    <button class="btn btn-info btn-small" onclick="printHandover('${h.id}')">طباعة</button>
                    <button class="btn btn-success btn-small" onclick="handoverPdf('${h.id}')">PDF</button>
                    <button class="btn btn-primary btn-small" onclick="openHandoverModal('${h.id}')">تعديل</button>
                    <button class="btn btn-danger btn-small" onclick="deleteHandover('${h.id}')">حذف</button>
                </div>
            </td>
        </tr>`).join('');
}

function resetHandoverFilters() {
    ['handoverSearch', 'handoverVehicleFilter', 'handoverFrom', 'handoverTo'].forEach(id => { byId(id).value = ''; });
    filterHandovers();
}

function getFilteredPenalties() {
    const q = (byId('penaltySearch').value || '').trim().toLowerCase();
    const type = byId('penaltyTypeFilter').value;
    const from = byId('penaltyFrom').value;
    const to = byId('penaltyTo').value;
    return appData.penalties.filter(p => {
        if (type && p.penalty_type !== type) return false;
        if ((from || to) && !dateInRange(p.penalty_date, from, to)) return false;
        if (q && !`${p.employee_name} ${p.department} ${p.reason}`.toLowerCase().includes(q)) return false;
        return true;
    }).sort((a, b) => (b.penalty_date || '').localeCompare(a.penalty_date || '') || (Number(b.id) || 0) - (Number(a.id) || 0));
}

function filterPenalties() {
    const tbody = byId('penaltiesTable');
    if (!tbody) return;
    const rows = getFilteredPenalties();
    if (!rows.length) {
        tbody.innerHTML = '<tr class="empty-row"><td colspan="8">لا توجد خصومات أو جزاءات مسجلة</td></tr>';
        return;
    }
    tbody.innerHTML = rows.map(p => `
        <tr>
            <td>${escapeHtml(p.penalty_date) || '-'}</td>
            <td><strong>${escapeHtml(p.employee_name)}</strong></td>
            <td>${escapeHtml(p.department) || '-'}</td>
            <td>${escapeHtml(p.penalty_type)}</td>
            <td>${escapeHtml(p.penalty_amount) || '-'}</td>
            <td>${escapeHtml(p.incident_date) || '-'}</td>
            <td>${escapeHtml(p.reason) || '-'}</td>
            <td>
                <div class="action-buttons">
                    <button class="btn btn-info btn-small" onclick="printPenalty('${p.id}')">طباعة</button>
                    <button class="btn btn-success btn-small" onclick="penaltyPdf('${p.id}')">PDF</button>
                    <button class="btn btn-primary btn-small" onclick="openPenaltyModal('${p.id}')">تعديل</button>
                    <button class="btn btn-danger btn-small" onclick="deletePenalty('${p.id}')">حذف</button>
                </div>
            </td>
        </tr>`).join('');
}

function resetPenaltyFilters() {
    ['penaltySearch', 'penaltyTypeFilter', 'penaltyFrom', 'penaltyTo'].forEach(id => { byId(id).value = ''; });
    filterPenalties();
}

/* ---------------------------------------------------------------------
   إقرار الاستلام: إضافة / تعديل / حذف
   --------------------------------------------------------------------- */
function fillHandoverFromVehicle() {
    const vehicle = appData.vehicles.find(v => v.id === byId('hoVehicle').value);
    if (!vehicle) return;
    const fromVehicle = {
        hoVehicleType: vehicle.status, hoVehicleName: vehicle.model, hoChassis: vehicle.chassis_number,
        hoEngine: vehicle.engine_number, hoYear: vehicle.year, hoLicStart: getLicenseStart(vehicle),
        hoLicEnd: vehicle.license_expiry
    };
    Object.entries(fromVehicle).forEach(([id, val]) => { byId(id).value = val ?? ''; });
    if (!byId('hoDelegate').value) byId('hoDelegate').value = vehicle.vin_number || '';
    updateHandoverHint();
}

function updateHandoverHint() {
    const hint = byId('hoLastHint');
    const vehicle = appData.vehicles.find(v => v.id === byId('hoVehicle').value);
    if (!vehicle) { hint.textContent = ''; return; }
    const last = getVehicleHandovers(vehicle).find(h => h.id !== editingHandoverId);
    hint.textContent = last
        ? `آخر استلام للسيارة دي: ${last.delegate_name || '-'} بتاريخ ${last.delivery_date || '-'} (الموقّع: ${last.signature_name || '-'})`
        : 'مفيش استلام سابق مسجل للسيارة دي.';
}

function openHandoverModal(id = null, vehicleId = '') {
    editingHandoverId = id;
    byId('handoverForm').reset();
    fillVehicleSelect('hoVehicle', 'اختر مركبة', false);
    byId('handoverModalTitle').textContent = id ? 'تعديل إقرار استلام' : 'إقرار استلام سيارة جديد';

    if (id) {
        const h = appData.handovers.find(x => x.id === id);
        if (!h) return;
        byId('hoVehicle').value = h.vehicle_id || '';
        Object.entries(HO_FIELDS).forEach(([eid, key]) => { byId(eid).value = h[key] || ''; });
    } else {
        if (vehicleId) {
            byId('hoVehicle').value = vehicleId;
            fillHandoverFromVehicle();
        }
        byId('hoDate').value = todayStr();
    }
    updateHandoverHint();
    byId('handoverModal').classList.add('show');
}

function closeHandoverModal() {
    byId('handoverModal').classList.remove('show');
    byId('handoverForm').reset();
    editingHandoverId = null;
}

function handleHandoverSubmit(e) {
    e.preventDefault();
    const nationalId = toLatinDigits(byId('hoNationalId').value).trim();
    if (nationalId && !/^\d{14}$/.test(nationalId)) {
        alert('الرقم القومي لازم يكون 14 رقم');
        return;
    }
    const licStart = byId('hoLicStart').value;
    const licEnd = byId('hoLicEnd').value;
    if (licStart && licEnd && licStart > licEnd) {
        alert('تاريخ بداية الرخصة لازم يكون قبل تاريخ الانتهاء');
        return;
    }

    const existing = editingHandoverId ? appData.handovers.find(h => h.id === editingHandoverId) : null;
    const vehicle = appData.vehicles.find(v => v.id === byId('hoVehicle').value);
    const record = {
        ...(existing || {}),
        id: editingHandoverId || Date.now().toString(),
        vehicle_id: vehicle ? vehicle.id : (existing?.vehicle_id || ''),
        plate_number: vehicle ? vehicle.plate_number : (existing?.plate_number || '')
    };
    if (!record.plate_number) {
        alert('اختار المركبة الأول');
        return;
    }
    Object.entries(HO_FIELDS).forEach(([eid, key]) => { record[key] = byId(eid).value.trim(); });
    record.national_id = nationalId;
    record.created_at = existing?.created_at || new Date().toISOString();

    if (existing) {
        appData.handovers[appData.handovers.findIndex(h => h.id === existing.id)] = record;
    } else {
        appData.handovers.push(record);
    }
    syncVehicleFromHandover(record.vehicle_id);

    saveData();
    closeHandoverModal();
    populateHandovers();
    populateVehiclesList();
    renderDashboard();
    if (!existing && confirm('تم حفظ الإقرار. تحب تطبعه دلوقتي؟')) printHandover(record.id);
}

function deleteHandover(id) {
    if (!confirm('هل أنت متأكد من حذف إقرار الاستلام ده؟')) return;
    const removed = appData.handovers.find(h => h.id === id);
    appData.handovers = appData.handovers.filter(h => h.id !== id);
    if (removed) syncVehicleFromHandover(removed.vehicle_id);
    saveData();
    populateHandovers();
    populateVehiclesList();
    renderDashboard();
    const historyModal = byId('handoverHistoryModal');
    if (historyModal.classList.contains('show') && historyModal.dataset.vehicleId) showHandoverHistory(historyModal.dataset.vehicleId);
}

// تحديث يدوي لكل المركبات مرة واحدة: مفيد لو في إقرارات استلام اتسجلت
// قبل ما التحديث التلقائي يتفعّل، فبتحدّث بيانات كل مركبة من آخر إقرار استلام ليها.
function syncAllVehiclesFromHandovers() {
    if (!confirm('هيتم تحديث اسم السائق ورقم الشاسيه ورقم الموتور وبداية وانتهاء الرخصة لكل مركبة من آخر إقرار استلام مسجّل ليها. تكمل؟')) return;
    appData.vehicles.forEach(v => syncVehicleFromHandover(v.id));
    saveData();
    populateVehiclesList();
    renderDashboard();
    alert('تم تحديث بيانات المركبات من إقرارات الاستلام.');
}

/* ---------------------------------------------------------------------
   سجل الاستلام لمركبة (مين استلم قبل كده وبتاريخ كام ومين وقّع)
   --------------------------------------------------------------------- */
function showHandoverHistory(vehicleId) {
    const vehicle = appData.vehicles.find(v => v.id === vehicleId);
    if (!vehicle) return;
    const modal = byId('handoverHistoryModal');
    modal.dataset.vehicleId = vehicleId;
    const list = getVehicleHandovers(vehicle);

    byId('handoverHistoryTitle').textContent = `سجل استلام السيارة ${vehicle.plate_number}`;
    byId('handoverHistoryCurrent').innerHTML = list.length
        ? `المستلم الحالي: <strong>${escapeHtml(list[0].delegate_name) || '-'}</strong> — بتاريخ <strong>${escapeHtml(list[0].delivery_date) || '-'}</strong> — الموقّع: <strong>${escapeHtml(list[0].signature_name) || '-'}</strong>`
        : 'مفيش استلام مسجل للسيارة دي لحد دلوقتي.';
    byId('handoverHistoryNew').onclick = () => { closeHandoverHistoryModal(); openHandoverModal(null, vehicleId); };

    byId('handoverHistoryBody').innerHTML = list.length
        ? list.map((h, i) => `
            <tr>
                <td>${escapeHtml(h.delivery_date) || '-'}${i === 0 ? ' <span class="status-badge status-complete">الحالي</span>' : ''}</td>
                <td>${escapeHtml(h.delegate_name) || '-'}</td>
                <td>${escapeHtml(h.job_title) || '-'}</td>
                <td><strong>${escapeHtml(h.signature_name) || '-'}</strong></td>
                <td>${escapeHtml(h.notes) || '-'}</td>
                <td><div class="action-buttons">
                    <button class="btn btn-info btn-small" onclick="printHandover('${h.id}')">طباعة</button>
                    <button class="btn btn-primary btn-small" onclick="closeHandoverHistoryModal();openHandoverModal('${h.id}')">تعديل</button>
                </div></td>
            </tr>`).join('')
        : '<tr class="empty-row"><td colspan="6">لا توجد سجلات</td></tr>';
    modal.classList.add('show');
}

function closeHandoverHistoryModal() {
    byId('handoverHistoryModal').classList.remove('show');
}

/* ---------------------------------------------------------------------
   الخصومات والجزاءات: إضافة / تعديل / حذف
   --------------------------------------------------------------------- */
function openPenaltyModal(id = null) {
    editingPenaltyId = id;
    byId('penaltyForm').reset();
    byId('penaltyModalTitle').textContent = id ? 'تعديل ورقة خصم / جزاء' : 'ورقة خصم / جزاء جديدة';
    if (id) {
        const p = appData.penalties.find(x => x.id === id);
        if (!p) return;
        Object.entries(PN_FIELDS).forEach(([eid, key]) => { byId(eid).value = p[key] || ''; });
    } else {
        byId('pnDate').value = todayStr();
    }
    byId('penaltyModal').classList.add('show');
}

function closePenaltyModal() {
    byId('penaltyModal').classList.remove('show');
    byId('penaltyForm').reset();
    editingPenaltyId = null;
}

function handlePenaltySubmit(e) {
    e.preventDefault();
    const existing = editingPenaltyId ? appData.penalties.find(p => p.id === editingPenaltyId) : null;
    const record = { ...(existing || {}), id: editingPenaltyId || Date.now().toString() };
    Object.entries(PN_FIELDS).forEach(([eid, key]) => { record[key] = byId(eid).value.trim(); });
    record.created_at = existing?.created_at || new Date().toISOString();

    if (existing) {
        appData.penalties[appData.penalties.findIndex(p => p.id === existing.id)] = record;
    } else {
        appData.penalties.push(record);
    }
    saveData();
    closePenaltyModal();
    filterPenalties();
    renderDashboard();
    if (!existing && confirm('تم حفظ الورقة. تحب تطبعها دلوقتي؟')) printPenalty(record.id);
}

function deletePenalty(id) {
    if (!confirm('هل أنت متأكد من حذف الورقة دي؟')) return;
    appData.penalties = appData.penalties.filter(p => p.id !== id);
    saveData();
    filterPenalties();
    renderDashboard();
}

/* ---------------------------------------------------------------------
   طباعة / PDF للإقرار وورقة الجزاء
   --------------------------------------------------------------------- */
function logoUrl() {
    return new URL('5.jpg', document.baseURI).href;
}

const DOC_CSS = `
.dc{font-family:'Cairo',sans-serif;color:#111;padding:18px;direction:rtl}
.dc *{box-sizing:border-box}
.dc-head{display:flex;align-items:center;justify-content:space-between;gap:14px;border-bottom:2px solid #d4af37;padding-bottom:10px;margin-bottom:12px}
.dc-logo{width:100px;height:100px;object-fit:contain}
.dc-title{text-align:center;flex:1}
.dc-title h2{margin:0 0 4px;font-size:22px;color:#0f3d6b}
.dc-title div{font-size:13px;color:#555}
.dc-statement{background:#f8f9fb;border:1px solid #e5e7eb;border-radius:10px;padding:10px;text-align:center;font-size:15px;line-height:1.8;margin:8px 0 12px}
.dc-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px 18px;margin:10px 0 14px}
.dc-item{border-bottom:1px dashed #cbd5e1;padding-bottom:6px}
.dc-item b{display:block;font-size:13px;margin-bottom:2px}
.dc-item span{font-size:14px}
.dc-full{grid-column:1/-1}
.dc-sign{display:grid;grid-template-columns:repeat(3,1fr);gap:18px;margin-top:22px}
.dc-sign div{border-top:1px solid #999;padding-top:6px;min-height:80px;text-align:center;font-size:14px}
.dc-sign small{display:block;color:#444;margin-top:6px;font-size:12px}
.dc-foot{margin-top:14px;font-size:12px;color:#555;text-align:center}
.dc-table{width:100%;border-collapse:collapse;margin:14px 0}
.dc-table th,.dc-table td{border:1px solid #aeb8c4;padding:9px;text-align:right;font-size:14px}
.dc-table th{background:#eef2f6;width:26%}
.dc-box{border:2px solid #d4af37;padding:12px;margin:12px 0;line-height:1.9;font-size:14px}
`;

function docHeader(title) {
    return `<div class="dc-head">
        <img class="dc-logo" src="${logoUrl()}" alt="شعار المصنع">
        <div class="dc-title"><h2>${title}</h2><div>إدارة حركة مصنع بن البهنساوي</div></div>
        <div class="dc-logo"></div></div>`;
}

function handoverDocFragment(h) {
    const item = (label, value, full) =>
        `<div class="dc-item${full ? ' dc-full' : ''}"><b>${label}</b><span>${escapeHtml(value) || '-'}</span></div>`;
    return `<div class="dc" dir="rtl"><style>${DOC_CSS}</style>
        ${docHeader('إقرار استلام سيارة')}
        <div class="dc-statement">هذا إقرار باستلام السيارة من إدارة حركة مصنع بن البهنساوي، وذلك وفق البيانات الموضحة أدناه.</div>
        <div class="dc-grid">
            ${item('اسم المندوب', h.delegate_name)}
            ${item('الوظيفة', h.job_title)}
            ${item('الإدارة التابعة', h.department)}
            ${item('الرقم القومي', h.national_id)}
            ${item('تاريخ الاستلام', fmtArDate(h.delivery_date))}
            ${item('نوع السيارة', h.vehicle_type)}
            ${item('رقم السيارة / اللوحة', h.plate_number)}
            ${item('اسم / موديل السيارة', h.vehicle_name)}
            ${item('رقم الشاسيه', h.chassis_number)}
            ${item('رقم الموتور', h.engine_number)}
            ${item('اللون', h.color)}
            ${item('سنة الصنع', h.manufacture_year)}
            ${item('تاريخ بداية الرخصة', fmtArDate(h.license_start))}
            ${item('تاريخ نهاية الرخصة', fmtArDate(h.license_expiry))}
            ${item('العدة والملحقات', h.accessories, true)}
            ${item('اسم من قام بالتوقيع', h.signature_name)}
            ${item('مسؤول الحركة', h.responsible_name)}
            ${item('البصمة / ملاحظات البصمة', h.fingerprint)}
            ${item('ملاحظات', h.notes, true)}
        </div>
        <div class="dc-sign">
            <div>توقيع المستلم<small>${escapeHtml(h.signature_name)}</small><small>${escapeHtml(fmtArDate(h.delivery_date))}</small></div>
            <div>توقيع مسؤول الحركة</div>
            <div>البصمة</div>
        </div>
        <div class="dc-foot">تم إنشاء هذا الإقرار بتاريخ ${new Date().toLocaleString('ar-EG')}</div>
    </div>`;
}

function penaltyDocFragment(p) {
    return `<div class="dc" dir="rtl"><style>${DOC_CSS}</style>
        ${docHeader('ورقة خصومات وجزاءات')}
        <div class="dc-box">تم تحرير هذه الورقة لإثبات الإجراء الإداري والمالي المتخذ بحق الموظف، وذلك وفقًا للبيانات الموضحة أدناه، لاعتمادها من الجهات المختصة.</div>
        <table class="dc-table">
            <tr><th>تاريخ الورقة</th><td>${escapeHtml(fmtArDate(p.penalty_date)) || '-'}</td></tr>
            <tr><th>اسم الموظف</th><td>${escapeHtml(p.employee_name) || '-'}</td></tr>
            <tr><th>الإدارة / القسم</th><td>${escapeHtml(p.department) || '-'}</td></tr>
            <tr><th>نوع الإجراء</th><td>${escapeHtml(p.penalty_type) || '-'}</td></tr>
            <tr><th>قيمة الخصم / مدة الجزاء</th><td>${escapeHtml(p.penalty_amount) || '-'}</td></tr>
            <tr><th>تاريخ الواقعة</th><td>${escapeHtml(fmtArDate(p.incident_date)) || '-'}</td></tr>
            <tr><th>سبب الخصم أو الجزاء</th><td>${escapeHtml(p.reason) || '-'}</td></tr>
        </table>
        <div class="dc-sign"><div>توقيع المدير</div><div>توقيع HR</div><div>اعتماد الإدارة</div></div>
        <div class="dc-foot">تم إنشاء هذه الورقة بتاريخ ${new Date().toLocaleString('ar-EG')}</div>
    </div>`;
}

function printFragment(fragment, title) {
    const w = window.open('', '', 'width=1000,height=800');
    if (!w) { alert('المتصفح منع نافذة الطباعة. اسمح بالنوافذ المنبثقة للموقع وجرب تاني.'); return; }
    w.document.write(`<!DOCTYPE html><html dir="rtl" lang="ar"><head><meta charset="UTF-8"><title>${title}</title>
        <link href="https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap" rel="stylesheet">
        </head><body style="margin:0">${fragment}</body></html>`);
    w.document.close();
    const imgs = [...w.document.images];
    Promise.all(imgs.map(i => i.complete ? null : new Promise(r => { i.onload = i.onerror = r; })))
        .then(() => setTimeout(() => w.print(), 300));
}

function printHandover(id) {
    const h = appData.handovers.find(x => x.id === id);
    if (h) printFragment(handoverDocFragment(h), 'إقرار استلام سيارة');
}

function handoverPdf(id) {
    const h = appData.handovers.find(x => x.id === id);
    if (h) openHtmlAsPdf(handoverDocFragment(h), `إقرار_استلام_${h.plate_number}_${h.delivery_date || ''}.pdf`, 'portrait');
}

function printPenalty(id) {
    const p = appData.penalties.find(x => x.id === id);
    if (p) printFragment(penaltyDocFragment(p), 'ورقة خصومات وجزاءات');
}

function penaltyPdf(id) {
    const p = appData.penalties.find(x => x.id === id);
    if (p) openHtmlAsPdf(penaltyDocFragment(p), `ورقة_جزاء_${p.employee_name}_${p.penalty_date || ''}.pdf`, 'portrait');
}

/* ---------------------------------------------------------------------
   تقرير استلام السيارات (PDF / Excel / طباعة)
   --------------------------------------------------------------------- */
function getHandoversReportData() {
    const vehicleId = byId('handoversReportVehicle').value;
    const from = byId('handoversReportFrom').value;
    const to = byId('handoversReportTo').value;
    const vehicle = vehicleId ? appData.vehicles.find(v => v.id === vehicleId) : null;

    const rows = appData.handovers
        .filter(h => (!vehicle || handoverBelongsTo(h, vehicle)) && ((!from && !to) || dateInRange(h.delivery_date, from, to)))
        .sort(handoverSort);

    const vehicles = vehicle ? [vehicle] : appData.vehicles;
    const perVehicle = vehicles.map(v => {
        const all = getVehicleHandovers(v);
        const last = all[0];
        return {
            plate: v.plate_number, model: v.model, count: all.length,
            receiver: last ? last.delegate_name : '', date: last ? last.delivery_date : '',
            signer: last ? last.signature_name : ''
        };
    });
    return { rows, perVehicle, subtitle: (vehicle ? `السيارة: ${vehicle.plate_number} — ` : '') + periodLabel(from, to) };
}

function generateHandoversReportHTML() {
    const { rows, perVehicle, subtitle } = getHandoversReportData();
    const chips = reportChips([
        { label: 'إقرارات في التقرير', value: rows.length },
        { label: 'مركبات ليها استلام', value: perVehicle.filter(v => v.count).length },
        { label: 'مركبات بدون استلام', value: perVehicle.filter(v => !v.count).length }
    ]);
    const main = rows.length
        ? reportTable(
            ['تاريخ الاستلام', 'اللوحة', 'السيارة', 'المستلم', 'الوظيفة / الإدارة', 'اسم الموقّع', 'مسؤول الحركة', 'البصمة', 'الملاحظات'],
            rows.map(h => [
                escapeHtml(h.delivery_date) || '-', escapeHtml(h.plate_number), escapeHtml(h.vehicle_name) || '-',
                escapeHtml(h.delegate_name) || '-', escapeHtml([h.job_title, h.department].filter(Boolean).join(' - ')) || '-',
                `<b>${escapeHtml(h.signature_name) || '-'}</b>`, escapeHtml(h.responsible_name) || '-',
                escapeHtml(h.fingerprint) || '-', escapeHtml(h.notes) || '-'
            ]))
        : '<p style="text-align:center;color:#888;">لا توجد إقرارات في الفترة المحددة</p>';
    const last = reportTable(
        ['اللوحة', 'النموذج', 'المستلم الحالي', 'تاريخ آخر استلام', 'الموقّع', 'عدد مرات الاستلام'],
        perVehicle.map(v => [escapeHtml(v.plate), escapeHtml(v.model), escapeHtml(v.receiver) || '-', escapeHtml(v.date) || '-', escapeHtml(v.signer) || '-', v.count]));
    return reportShell('تقرير استلام السيارات', subtitle,
        chips + reportSectionTitle('سجل الاستلام') + main + reportSectionTitle('آخر مستلم لكل سيارة') + last);
}

function generateHandoversReportPDF() { openHtmlAsPdf(generateHandoversReportHTML(), 'تقرير_استلام_السيارات.pdf', 'landscape'); }
function printHandoversReport() { printHtmlReport(generateHandoversReportHTML()); }

function generateHandoversReportExcel() {
    const { rows, perVehicle, subtitle } = getHandoversReportData();
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
        ['تقرير استلام السيارات'], [subtitle], [],
        ['تاريخ الاستلام', 'اللوحة', 'السيارة', 'المستلم', 'الوظيفة', 'الإدارة', 'الرقم القومي', 'اسم الموقّع', 'مسؤول الحركة', 'البصمة', 'الملاحظات'],
        ...rows.map(h => [h.delivery_date || '', h.plate_number, h.vehicle_name || '', h.delegate_name || '', h.job_title || '',
            h.department || '', h.national_id || '', h.signature_name || '', h.responsible_name || '', h.fingerprint || '', h.notes || ''])
    ]), 'الاستلام');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
        ['آخر مستلم لكل سيارة'], [subtitle], [],
        ['اللوحة', 'النموذج', 'المستلم الحالي', 'تاريخ آخر استلام', 'الموقّع', 'عدد مرات الاستلام'],
        ...perVehicle.map(v => [v.plate, v.model, v.receiver || '', v.date || '', v.signer || '', v.count])
    ]), 'آخر مستلم');
    XLSX.writeFile(wb, 'تقرير_استلام_السيارات.xlsx');
}

/* ---------------------------------------------------------------------
   تقرير الخصومات والجزاءات
   --------------------------------------------------------------------- */
function getPenaltiesReportData() {
    const from = byId('penaltiesReportFrom').value;
    const to = byId('penaltiesReportTo').value;
    const rows = appData.penalties
        .filter(p => (!from && !to) || dateInRange(p.penalty_date, from, to))
        .sort((a, b) => (b.penalty_date || '').localeCompare(a.penalty_date || ''));
    return { rows, subtitle: periodLabel(from, to) };
}

function generatePenaltiesReportHTML() {
    const { rows, subtitle } = getPenaltiesReportData();
    const types = [...new Set(rows.map(p => p.penalty_type))];
    const chips = reportChips([{ label: 'إجمالي الأوراق', value: rows.length },
        ...types.map(t => ({ label: escapeHtml(t), value: rows.filter(p => p.penalty_type === t).length }))]);
    const table = rows.length
        ? reportTable(['التاريخ', 'الموظف', 'الإدارة', 'نوع الإجراء', 'القيمة / المدة', 'تاريخ الواقعة', 'السبب'],
            rows.map(p => [escapeHtml(p.penalty_date) || '-', `<b>${escapeHtml(p.employee_name)}</b>`, escapeHtml(p.department) || '-',
                escapeHtml(p.penalty_type), escapeHtml(p.penalty_amount) || '-', escapeHtml(p.incident_date) || '-', escapeHtml(p.reason) || '-']))
        : '<p style="text-align:center;color:#888;">لا توجد أوراق في الفترة المحددة</p>';
    return reportShell('تقرير الخصومات والجزاءات', subtitle, chips + table);
}

function generatePenaltiesReportPDF() { openHtmlAsPdf(generatePenaltiesReportHTML(), 'تقرير_الخصومات_والجزاءات.pdf', 'landscape'); }
function printPenaltiesReport() { printHtmlReport(generatePenaltiesReportHTML()); }

function generatePenaltiesReportExcel() {
    const { rows, subtitle } = getPenaltiesReportData();
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
        ['تقرير الخصومات والجزاءات'], [subtitle], [],
        ['التاريخ', 'الموظف', 'الإدارة', 'نوع الإجراء', 'القيمة / المدة', 'تاريخ الواقعة', 'السبب'],
        ...rows.map(p => [p.penalty_date || '', p.employee_name || '', p.department || '', p.penalty_type || '',
            p.penalty_amount || '', p.incident_date || '', p.reason || ''])
    ]), 'الجزاءات');
    XLSX.writeFile(wb, 'تقرير_الخصومات_والجزاءات.xlsx');
}

/* ---------------------------------------------------------------------
   استيراد بيانات النظام القديم (إقرار استلام سيارة)
   --------------------------------------------------------------------- */
function mapLegacyReceipt(o, seq) {
    const plate = String(o.plateNumber || '').trim();
    const vehicle = appData.vehicles.find(v => String(v.plate_number).trim() === plate);
    return {
        id: (Date.now() + seq).toString(),
        vehicle_id: vehicle ? vehicle.id : '',
        plate_number: plate,
        vehicle_type: o.vehicleType || '', vehicle_name: o.vehicleName || '',
        chassis_number: o.chassisNumber || '', engine_number: o.engineNumber || '',
        color: o.color || '', manufacture_year: o.manufactureYear || '',
        license_start: o.licenseStartDate || '', license_expiry: o.licenseEndDate || '',
        accessories: o.accessories || '', delegate_name: o.delegateName || '',
        job_title: o.delegateJobTitle || o.position || '', department: o.department || '',
        national_id: o.nationalId || '', delivery_date: o.deliveryDate || '',
        signature_name: o.signatureName || '', fingerprint: o.fingerprint || '',
        notes: o.notes || '', created_at: new Date().toISOString(), imported: true
    };
}

function mapLegacyPenalty(o, seq) {
    return {
        id: (Date.now() + seq + 100000).toString(),
        penalty_date: o.penaltyDate || '', employee_name: o.employeeName || '',
        department: o.employeeDepartment || '', penalty_type: o.penaltyType || '',
        penalty_amount: o.penaltyAmount || '', incident_date: o.incidentDate || '',
        reason: o.reason || '', created_at: new Date().toISOString(), imported: true
    };
}

// بيرجع عدد اللي اتضاف فعلاً (من غير تكرار)
function importLegacyRecords(receipts, penalties) {
    const receiptKey = (r) => JSON.stringify([r.plate_number, r.delivery_date, r.delegate_name, r.signature_name, r.chassis_number, r.notes]);
    const penaltyKey = (p) => JSON.stringify([p.penalty_date, p.employee_name, p.penalty_type, p.penalty_amount, p.incident_date]);
    const haveR = new Set(appData.handovers.map(receiptKey));
    const haveP = new Set(appData.penalties.map(penaltyKey));
    let addedR = 0, addedP = 0;

    receipts.forEach((o, i) => {
        const r = mapLegacyReceipt(o, i);
        if (!r.plate_number || haveR.has(receiptKey(r))) return;
        haveR.add(receiptKey(r));
        appData.handovers.push(r);
        addedR++;
    });
    penalties.forEach((o, i) => {
        const p = mapLegacyPenalty(o, i);
        if (!p.employee_name || haveP.has(penaltyKey(p))) return;
        haveP.add(penaltyKey(p));
        appData.penalties.push(p);
        addedP++;
    });
    return { addedR, addedP };
}

function triggerLegacyImport() {
    byId('legacyImportFile').click();
}

function handleLegacyImport(event) {
    const [file] = event.target.files || [];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => {
        try {
            const parsed = JSON.parse(reader.result);
            const list = Array.isArray(parsed) ? parsed : [];
            const receipts = Array.isArray(parsed?.carReceipts) ? parsed.carReceipts : list.filter(o => o && o.plateNumber !== undefined);
            const penalties = Array.isArray(parsed?.penaltyRecords) ? parsed.penaltyRecords : list.filter(o => o && o.employeeName !== undefined);
            const { addedR, addedP } = importLegacyRecords(receipts, penalties);
            saveData();
            populateHandovers();
            populateVehiclesList();
            renderDashboard();
            alert(`تم استيراد ${addedR} إقرار استلام و ${addedP} ورقة خصم/جزاء (المكرر اتخطى).`);
        } catch (error) {
            console.error('خطأ في استيراد النظام القديم:', error);
            alert('تعذر قراءة الملف. لازم يكون ملف JSON اتصدّر من نظام إقرار استلام السيارة (car-receipts.json أو penalty-records.json).');
        } finally {
            event.target.value = '';
        }
    };
    reader.readAsText(file, 'utf-8');
}

// لو النظامين على نفس الدومين، بنجيب الإقرارات القديمة المحفوظة في المتصفح مرة واحدة
function migrateLegacyMotorData() {
    try {
        if (localStorage.getItem('legacyMotorMigrated')) return;
        const receipts = JSON.parse(localStorage.getItem('carReceipts') || '[]');
        const penalties = JSON.parse(localStorage.getItem('penaltyRecords') || '[]');
        if (receipts.length || penalties.length) {
            importLegacyRecords(receipts, penalties);
            saveData();
        }
        localStorage.setItem('legacyMotorMigrated', '1');
    } catch (error) {
        console.error('خطأ في نقل بيانات الاستلام القديمة:', error);
    }
}

/* ---------------------------------------------------------------------
   ربط الأحداث
   --------------------------------------------------------------------- */
function setupHandoverListeners() {
    byId('handoverForm').addEventListener('submit', handleHandoverSubmit);
    byId('penaltyForm').addEventListener('submit', handlePenaltySubmit);
    byId('hoVehicle').addEventListener('change', fillHandoverFromVehicle);
    byId('hoDelegate').addEventListener('blur', () => {
        if (!byId('hoSigner').value) byId('hoSigner').value = byId('hoDelegate').value;
    });
    byId('legacyImportFile').addEventListener('change', handleLegacyImport);

    byId('handoverSearch').addEventListener('input', filterHandovers);
    ['handoverVehicleFilter', 'handoverFrom', 'handoverTo'].forEach(id => byId(id).addEventListener('change', filterHandovers));
    byId('penaltySearch').addEventListener('input', filterPenalties);
    ['penaltyTypeFilter', 'penaltyFrom', 'penaltyTo'].forEach(id => byId(id).addEventListener('change', filterPenalties));

    // كروت لوحة التحكم تفتح القسم بتاعها
    const cardTargets = {
        totalVehicles: ['vehicles'], expiringLicenses: ['licenses'], inMaintenance: ['maintenance'],
        monthlyExpenses: ['maintenance'], totalViolations: ['violations'],
        pendingViolations: ['violations', 'violationStatusFilter', 'معلقة'], totalHandovers: ['handovers']
    };
    Object.entries(cardTargets).forEach(([id, [section, selectId, value]]) => {
        const card = byId(id)?.closest('.stat-card');
        if (!card) return;
        card.addEventListener('click', () => {
            showSection(section);
            if (selectId) quickFilter(selectId, value);
        });
    });

    migrateLegacyMotorData();
}
