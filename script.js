let licenseWarningDays = 20;
let editingVehicleId = null;
let editingViolationId = null;
let listenersInitialized = false;



const appData = {
    vehicles: [],
    maintenance: [],
    violations: [],
    expenses: [],
    advance: [],
    tires_in: [],
    tires_out: [],
    movements: [],
    handovers: [],
    penalties: [],
};



function initializeApp() {
    loadData();
    setupEventListeners();
    updateDateTime();
    loadLicenseWarningDays();
    renderDashboard();
    setInterval(updateDateTime, 1000);
}

function saveData() {
    try {
        localStorage.setItem('vehicles', JSON.stringify(appData.vehicles));
        localStorage.setItem('maintenance', JSON.stringify(appData.maintenance));
        localStorage.setItem('violations', JSON.stringify(appData.violations));
        localStorage.setItem('expenses', JSON.stringify(appData.expenses));
        localStorage.setItem('advance', JSON.stringify(appData.advance));
        localStorage.setItem('tires_in', JSON.stringify(appData.tires_in));
        localStorage.setItem('tires_out', JSON.stringify(appData.tires_out));
        localStorage.setItem('movements', JSON.stringify(appData.movements));
        localStorage.setItem('handovers', JSON.stringify(appData.handovers));
        localStorage.setItem('penalties', JSON.stringify(appData.penalties));
        if (typeof renderAllSummaries === 'function') renderAllSummaries();
        if (window.CloudSync) window.CloudSync.schedulePush();
    } catch (error) {
        console.error('خطأ في حفظ البيانات:', error);
    }
}

function exportAppData() {
    try {
        const exportData = {
            version: 2,
            exported_at: new Date().toISOString(),
            licenseWarningDays,
            data: {
                vehicles: appData.vehicles,
                maintenance: appData.maintenance,
                violations: appData.violations,
                expenses: appData.expenses,
                advance: appData.advance,
                tires_in: appData.tires_in,
                tires_out: appData.tires_out,
                movements: appData.movements,
                handovers: appData.handovers,
                penalties: appData.penalties
            }
        };

        const blob = new Blob([JSON.stringify(exportData, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const link = document.createElement('a');
        const dateLabel = new Date().toISOString().slice(0, 10);

        link.href = url;
        link.download = `cars_backup_${dateLabel}.json`;
        document.body.appendChild(link);
        link.click();
        link.remove();
        URL.revokeObjectURL(url);
    } catch (error) {
        console.error('خطأ في تصدير البيانات:', error);
        alert('حدث خطأ أثناء تنزيل البيانات.');
    }
}

function exportAllDataToExcel() {
    try {
        const wb = XLSX.utils.book_new();
        const exportSheets = [
            {
                name: 'المركبات',
                headers: ['اللوحة', 'الموديل', 'السنة', 'اسم السائق', 'الحالة', 'رقم الموتور', 'رقم الشاسيه', 'بداية الرخصة', 'انتهاء الرخصة', 'المستلم الحالي', 'تاريخ آخر استلام', 'الملاحظات'],
                keys: ['plate_number', 'model', 'year', 'vin_number', 'status', 'engine_number', 'chassis_number', 'license_start', 'license_expiry', 'last_receiver', 'last_receive_date', 'notes'],
                data: appData.vehicles.map(v => ({ ...v, license_start: getLicenseStart(v), last_receiver: lastReceiverName(v), last_receive_date: lastReceiveDate(v) }))
            },
            {
                name: 'الصيانة',
                headers: ['اللوحة', 'نوع الصيانة', 'التاريخ', 'التكلفة', 'الحالة', 'الملاحظات'],
                keys: ['plate_number', 'maintenance_type', 'maintenance_date', 'cost', 'status', 'notes'],
                data: appData.maintenance
            },
            {
                name: 'المخالفات',
                headers: ['اللوحة', 'نوع المخالفة', 'التاريخ', 'المبلغ', 'الحالة', 'الملاحظات'],
                keys: ['plate_number', 'violation_type', 'violation_date', 'amount', 'status', 'notes'],
                data: appData.violations
            },
            {
                name: 'النفقات',
                headers: ['نوع السيارات', 'اللوحة', 'اسم السائق', 'نوع النفقة', 'التاريخ', 'المبلغ', 'رقم العهدة', 'الملاحظات'],
                keys: ['fleet_group', 'plate_number', 'driver_name', 'expense_type', 'expense_date', 'amount', 'advance_id', 'notes'],
                data: appData.expenses
            },
            {
                name: 'العهدة',
                headers: ['نوع العهدة', 'المبلغ', 'تاريخ العهدة', 'نشطة', 'الملاحظات'],
                keys: ['fleet_group', 'amount', 'advance_date', 'is_active', 'notes'],
                data: appData.advance
            },
            {
                name: 'مخزن الكوتش - إضافات',
                headers: ['التاريخ', 'نوع الكوتش', 'العدد', 'الملاحظات'],
                keys: ['date', 'tire_type', 'quantity', 'notes'],
                data: appData.tires_in
            },
            {
                name: 'مخزن الكوتش - صرف',
                headers: ['التاريخ', 'نوع الكوتش', 'العدد', 'اتصرف لعربية', 'اسم السائق', 'الملاحظات'],
                keys: ['date', 'tire_type', 'quantity', 'plate_number', 'driver_name', 'notes'],
                data: appData.tires_out.map(r => ({ ...r, plate_number: getRecordPlate(r) }))
            },
            {
                name: 'تحركات السيارات',
                headers: ['التاريخ', 'اللوحة', 'اسم السائق', 'الكيلومترات', 'مدة التشغيل', 'مكان الركن', 'الملاحظات'],
                keys: ['date', 'plate_number', 'driver_name', 'km', 'duration', 'parking_place', 'notes'],
                data: appData.movements.map(m => ({ ...m, plate_number: getRecordPlate(m), duration: formatDuration(getMovementMinutes(m)) }))
            },
            {
                name: 'استلام السيارات',
                headers: ['تاريخ الاستلام', 'اللوحة', 'نوع السيارة', 'السيارة', 'رقم الشاسيه', 'رقم الموتور', 'اللون', 'سنة الصنع', 'بداية الرخصة', 'انتهاء الرخصة', 'العدة والملحقات', 'المستلم', 'الوظيفة', 'الإدارة', 'الرقم القومي', 'اسم الموقّع', 'البصمة', 'الملاحظات'],
                keys: ['delivery_date', 'plate_number', 'vehicle_type', 'vehicle_name', 'chassis_number', 'engine_number', 'color', 'manufacture_year', 'license_start', 'license_expiry', 'accessories', 'delegate_name', 'job_title', 'department', 'national_id', 'signature_name', 'fingerprint', 'notes'],
                data: [...appData.handovers].sort(handoverSort)
            },
            {
                name: 'الخصومات والجزاءات',
                headers: ['التاريخ', 'الموظف', 'الإدارة', 'نوع الإجراء', 'القيمة / المدة', 'تاريخ الواقعة', 'السبب'],
                keys: ['penalty_date', 'employee_name', 'department', 'penalty_type', 'penalty_amount', 'incident_date', 'reason'],
                data: appData.penalties
            }
        ];

        exportSheets.forEach(sheet => {
            const rows = [sheet.headers];
            sheet.data.forEach(item => {
                rows.push(sheet.keys.map(key => {
                    const value = item[key];
                    if (key === 'is_active') {
                        return value ? 'نعم' : 'لا';
                    }
                    if (key === 'fleet_group') {
                        return getFleetGroupLabel(value);
                    }
                    if (key === 'driver_name') {
                        return getDriverNameByRecord(item);
                    }
                    return value !== undefined && value !== null ? value : '';
                }));
            });

            const ws = XLSX.utils.aoa_to_sheet(rows);
            XLSX.utils.book_append_sheet(wb, ws, sheet.name);
        });

        XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(allSummaryRows()), 'ملخص');

        const fileName = `cars_export_${new Date().toISOString().slice(0, 10)}.xlsx`;
        XLSX.writeFile(wb, fileName);
    } catch (error) {
        console.error('خطأ في تصدير البيانات إلى Excel:', error);
        alert('حدث خطأ أثناء تصدير البيانات إلى Excel.');
    }
}

function triggerImportData() {
    document.getElementById('importDataFile').click();
}

function handleImportData(event) {
    const [file] = event.target.files || [];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = () => {
        try {
            const imported = JSON.parse(reader.result);
            const importedData = imported.data || imported;

            if (
                !importedData ||
                !Array.isArray(importedData.vehicles) ||
                !Array.isArray(importedData.maintenance) ||
                !Array.isArray(importedData.violations) ||
                !Array.isArray(importedData.expenses) ||
                !Array.isArray(importedData.advance)
            ) {
                throw new Error('ملف البيانات غير صالح.');
            }

            appData.vehicles = importedData.vehicles;
            appData.maintenance = importedData.maintenance;
            appData.violations = importedData.violations;
            appData.expenses = importedData.expenses;
            appData.advance = importedData.advance;
            appData.tires_in = Array.isArray(importedData.tires_in) ? importedData.tires_in : [];
            appData.tires_out = Array.isArray(importedData.tires_out) ? importedData.tires_out : [];
            appData.movements = Array.isArray(importedData.movements) ? importedData.movements : [];
            appData.handovers = Array.isArray(importedData.handovers) ? importedData.handovers : [];
            appData.penalties = Array.isArray(importedData.penalties) ? importedData.penalties : [];

            if (Number.isInteger(imported.licenseWarningDays)) {
                licenseWarningDays = imported.licenseWarningDays;
                localStorage.setItem('licenseWarningDays', licenseWarningDays);
                document.getElementById('licenseWarningDays').value = licenseWarningDays;
            }

            saveData();
            renderDashboard();
            populateVehiclesList();
            populateMaintenanceList();
            populateViolationsList();
            populateLicensesList();
            populateExpensesList();
            populateAdvanceList();
            populateTiresSection();
            populateMovementsSection();
            populateHandovers();
            alert('تم استيراد البيانات بنجاح');
        } catch (error) {
            console.error('خطأ في استيراد البيانات:', error);
            alert('تعذر استيراد الملف. تأكد أنه ملف نسخة احتياطية صحيح.');
        } finally {
            event.target.value = '';
        }
    };

    reader.onerror = () => {
        console.error('خطأ في قراءة ملف الاستيراد');
        alert('تعذر قراءة ملف الاستيراد.');
        event.target.value = '';
    };

    reader.readAsText(file, 'utf-8');
}

function loadLicenseWarningDays() {
    const saved = localStorage.getItem('licenseWarningDays');
    if (saved) {
        licenseWarningDays = parseInt(saved);
        document.getElementById('licenseWarningDays').value = licenseWarningDays;
    }
}

function saveLicenseWarningDays() {
    const value = parseInt(document.getElementById('licenseWarningDays').value);
    if (value >= 1 && value <= 365) {
        licenseWarningDays = value;
        localStorage.setItem('licenseWarningDays', value);
        if (window.CloudSync) window.CloudSync.schedulePush();
        alert('تم حفظ الإعدادات بنجاح');
        renderDashboard();
    } else {
        alert('يرجى إدخال عدد أيام صحيح (بين 1 و 365)');
    }
}

function setupEventListeners() {
    if (listenersInitialized) return;

    document.querySelectorAll('.nav-link').forEach(link => {
        link.addEventListener('click', (e) => {
            e.preventDefault();
            const section = link.getAttribute('data-section');
            showSection(section);
        });
    });

    document.getElementById('vehicleForm').addEventListener('submit', handleVehicleSubmit);
    document.getElementById('maintenanceForm').addEventListener('submit', handleMaintenanceSubmit);
    document.getElementById('violationForm').addEventListener('submit', handleViolationSubmit);
    document.getElementById('violationImage').addEventListener('change', handleViolationImageChange);
    document.getElementById('expensesForm').addEventListener('submit', handleExpensesSubmit);
    document.getElementById('advanceForm').addEventListener('submit', handleAdvanceSubmit);
    document.getElementById('vehicleSearch').addEventListener('input', filterVehicles);
    document.getElementById('vehicleStatusFilter').addEventListener('change', filterVehicles);
    document.getElementById('maintenanceSearch').addEventListener('input', filterMaintenance);
    document.getElementById('maintenanceStatusFilter').addEventListener('change', filterMaintenance);
    document.getElementById('licenseSearch').addEventListener('input', filterLicenses);
    document.getElementById('licenseStatusFilter').addEventListener('change', filterLicenses);
    document.getElementById('violationSearch').addEventListener('input', filterViolations);
    document.getElementById('violationStatusFilter').addEventListener('change', filterViolations);
    document.getElementById('expensesSearch').addEventListener('input', filterExpenses);
    document.getElementById('importDataFile').addEventListener('change', handleImportData);

    document.addEventListener('click', (e) => {
        if (e.target.classList.contains('modal')) {
            closeVehicleModal();
            closeMaintenanceModal();
            closeViolationModal();
            closeExpensesModal();
            closeTireInModal();
            closeTireOutModal();
            closeMovementModal();
            closeParkingHistoryModal();
            closeHandoverModal();
            closePenaltyModal();
            closeHandoverHistoryModal();
        }
    });

    setupNewFeatureListeners();

    listenersInitialized = true;
}

function updateDateTime() {
    const now = new Date();
    const options = {
        weekday: 'long',
        year: 'numeric',
        month: 'long',
        day: 'numeric',
        hour: '2-digit',
        minute: '2-digit',
        second: '2-digit'
    };
    const formatted = now.toLocaleDateString('ar-EG', options);
    document.getElementById('dateTime').textContent = formatted;
    document.getElementById('lastUpdate').textContent = formatted.split(',')[0];
}

function showSection(sectionId) {
    document.querySelectorAll('.section').forEach(s => s.classList.remove('active'));
    document.getElementById(sectionId).classList.add('active');

    document.querySelectorAll('.nav-link').forEach(l => l.classList.remove('active'));
    document.querySelector(`[data-section="${sectionId}"]`).classList.add('active');

    if (sectionId === 'dashboard') {
        renderDashboard();
    } else if (sectionId === 'vehicles') {
        populateVehiclesList();
    } else if (sectionId === 'handovers') {
        populateHandovers();
    } else if (sectionId === 'maintenance') {
        populateMaintenanceList();
    } else if (sectionId === 'violations') {
        populateViolationsList();
    } else if (sectionId === 'licenses') {
        populateLicensesList();
    } else if (sectionId === 'expenses') {
        populateExpensesList();
    } else if (sectionId === 'advance') {
        populateAdvanceList();
    } else if (sectionId === 'tires') {
        populateTiresSection();
    } else if (sectionId === 'movements') {
        populateMovementsSection();
    } else if (sectionId === 'reports') {
        populateReportVehicleOptions();
    }
}

function openVehicleModal(vehicleId = null) {
    editingVehicleId = vehicleId;
    const modal = document.getElementById('vehicleModal');
    const form = document.getElementById('vehicleForm');
    const title = document.getElementById('vehicleModalTitle');

    if (vehicleId) {
        title.textContent = 'تعديل بيانات المركبة';
        const vehicle = appData.vehicles.find(v => v.id === vehicleId);
        if (vehicle) {
            document.getElementById('plateNumber').value = vehicle.plate_number;
            document.getElementById('vehicleModel').value = vehicle.model;
            document.getElementById('vehicleYear').value = vehicle.year;
            document.getElementById('vinNumber').value = vehicle.vin_number || '';
            document.getElementById('vehicleStatus').value = vehicle.status;
            document.getElementById('engineNumber').value = vehicle.engine_number || '';
            document.getElementById('chassisNumber').value = vehicle.chassis_number || '';
            document.getElementById('licenseStart').value = getLicenseStart(vehicle);
            document.getElementById('licenseExpiry').value = vehicle.license_expiry;
            document.getElementById('vehicleNotes').value = vehicle.notes || '';
        }
    } else {
        title.textContent = 'إضافة مركبة جديدة';
        form.reset();
    }

    modal.classList.add('show');
}

function closeVehicleModal() {
    document.getElementById('vehicleModal').classList.remove('show');
    editingVehicleId = null;
    document.getElementById('vehicleForm').reset();
}

function handleVehicleSubmit(e) {
    e.preventDefault();

    const licenseStartValue = document.getElementById('licenseStart').value;
    const licenseExpiryValue = document.getElementById('licenseExpiry').value;
    if (licenseStartValue && licenseExpiryValue && licenseStartValue > licenseExpiryValue) {
        alert('تاريخ بداية الرخصة لازم يكون قبل تاريخ الانتهاء');
        return;
    }

    const existingVehicle = editingVehicleId ? appData.vehicles.find(v => v.id === editingVehicleId) : null;

    // نحافظ على أي بيانات قديمة موجودة على المركبة (زي رقم الرخصة القديم) بدل ما نمسحها
    const vehicleData = {
        ...(existingVehicle || {}),
        id: editingVehicleId || Date.now().toString(),
        plate_number: document.getElementById('plateNumber').value,
        model: document.getElementById('vehicleModel').value,
        year: parseInt(document.getElementById('vehicleYear').value),
        vin_number: document.getElementById('vinNumber').value,
        status: document.getElementById('vehicleStatus').value,
        engine_number: document.getElementById('engineNumber').value.trim(),
        chassis_number: document.getElementById('chassisNumber').value.trim(),
        license_start: licenseStartValue,
        license_expiry: licenseExpiryValue,
        notes: document.getElementById('vehicleNotes').value
    };

    try {
        if (editingVehicleId) {
            const index = appData.vehicles.findIndex(v => v.id === editingVehicleId);
            if (index !== -1) {
                appData.vehicles[index] = vehicleData;
            }
        } else {
            appData.vehicles.push(vehicleData);
        }

        saveData();
        closeVehicleModal();
        populateVehiclesList();
        renderDashboard();
        alert('تم حفظ البيانات بنجاح');
    } catch (error) {
        console.error('خطأ:', error);
        alert('حدث خطأ في حفظ البيانات: ' + error.message);
    }
}

function deleteVehicle(vehicleId) {
    if (confirm('هل أنت متأكد من حذف هذه المركبة؟')) {
        try {
            appData.vehicles = appData.vehicles.filter(v => v.id !== vehicleId);
            saveData();
            populateVehiclesList();
            renderDashboard();
        } catch (error) {
            console.error('خطأ:', error);
            alert('حدث خطأ في حذف البيانات');
        }
    }
}

function openMaintenanceModal(maintenanceId = null) {
    const modal = document.getElementById('maintenanceModal');
    const form = document.getElementById('maintenanceForm');
    const vehicleSelect = document.getElementById('maintenanceVehicle');

    vehicleSelect.innerHTML = '<option value="">اختر مركبة</option>';
    appData.vehicles.forEach(v => {
        const option = document.createElement('option');
        option.value = v.id;
        option.textContent = `${v.plate_number} - ${v.model}`;
        vehicleSelect.appendChild(option);
    });

    if (maintenanceId) {
        const maintenance = appData.maintenance.find(m => m.id === maintenanceId);
        if (maintenance) {
            document.getElementById('maintenanceVehicle').value = maintenance.vehicle_id;
            document.getElementById('maintenanceType').value = maintenance.maintenance_type;
            document.getElementById('maintenanceDate').value = maintenance.maintenance_date;
            document.getElementById('maintenanceCost').value = maintenance.cost;
            document.getElementById('maintenanceStatus').value = maintenance.status;
            document.getElementById('maintenanceNotes').value = maintenance.notes || '';
        }
    } else {
        form.reset();
    }

    modal.classList.add('show');
}

function closeMaintenanceModal() {
    document.getElementById('maintenanceModal').classList.remove('show');
    document.getElementById('maintenanceForm').reset();
}

function openViolationModal(violationId = null) {
    editingViolationId = violationId;
    const modal = document.getElementById('violationModal');
    const form = document.getElementById('violationForm');
    const vehicleSelect = document.getElementById('violationVehicle');
    const violationTitle = document.getElementById('violationModalTitle');
    if (violationTitle) violationTitle.textContent = violationId ? 'تعديل المخالفة' : 'تسجيل مخالفة جديدة';

    vehicleSelect.innerHTML = '<option value="">اختر مركبة</option>';
    appData.vehicles.forEach(v => {
        const option = document.createElement('option');
        option.value = v.id;
        option.textContent = `${v.plate_number} - ${v.model}`;
        vehicleSelect.appendChild(option);
    });

    if (violationId) {
        const violation = appData.violations.find(v => v.id === violationId);
        if (violation) {
            document.getElementById('violationVehicle').value = violation.vehicle_id;
            document.getElementById('violationType').value = violation.violation_type;
            document.getElementById('violationDate').value = violation.violation_date;
            document.getElementById('violationAmount').value = violation.amount;
            document.getElementById('violationStatus').value = violation.status;
            document.getElementById('violationNotes').value = violation.notes || '';
            updateViolationImagePreview(violation.image || '');
        }
    } else {
        form.reset();
        clearViolationImagePreview();
    }

    modal.classList.add('show');
}

function closeViolationModal() {
    document.getElementById('violationModal').classList.remove('show');
    document.getElementById('violationForm').reset();
    document.getElementById('violationImage').value = '';
    editingViolationId = null;
    clearViolationImagePreview();
}

function handleViolationImageChange(event) {
    const [file] = event.target.files || [];

    if (!file) {
        clearViolationImagePreview();
        return;
    }

    const reader = new FileReader();
    reader.onload = () => updateViolationImagePreview(reader.result);
    reader.readAsDataURL(file);
}

function updateViolationImagePreview(imageSrc) {
    const previewContainer = document.getElementById('violationImagePreviewContainer');
    const preview = document.getElementById('violationImagePreview');

    if (!imageSrc) {
        clearViolationImagePreview();
        return;
    }

    preview.src = imageSrc;
    previewContainer.classList.remove('hidden');
}

function clearViolationImagePreview() {
    const previewContainer = document.getElementById('violationImagePreviewContainer');
    const preview = document.getElementById('violationImagePreview');

    preview.src = '';
    previewContainer.classList.add('hidden');
}

function resizeImageFile(file, maxWidth = 1200, maxHeight = 1200, quality = 0.82) {
    return new Promise((resolve, reject) => {
        if (!file) {
            resolve('');
            return;
        }

        const reader = new FileReader();
        reader.onload = () => {
            const img = new Image();
            img.onload = () => {
                const canvas = document.createElement('canvas');
                let { width, height } = img;
                const ratio = Math.min(maxWidth / width, maxHeight / height, 1);

                width = Math.round(width * ratio);
                height = Math.round(height * ratio);
                canvas.width = width;
                canvas.height = height;

                const context = canvas.getContext('2d');
                context.drawImage(img, 0, 0, width, height);
                resolve(canvas.toDataURL('image/jpeg', quality));
            };
            img.onerror = () => reject(new Error('تعذر قراءة الصورة المرفوعة.'));
            img.src = reader.result;
        };
        reader.onerror = () => reject(new Error('تعذر تحميل ملف الصورة.'));
        reader.readAsDataURL(file);
    });
}

let editingExpenseId = null;

function openExpensesModal(expenseId = null) {
    editingExpenseId = expenseId;
    const modal = document.getElementById('expensesModal');
    const form = document.getElementById('expensesForm');
    const vehicleSelect = document.getElementById('expensesVehicle');
    const advanceSelect = document.getElementById('expensesAdvance');
    const title = document.getElementById('expensesModalTitle');

    vehicleSelect.innerHTML = '<option value="">اختر مركبة (اختياري)</option>';
    appData.vehicles.forEach(v => {
        const option = document.createElement('option');
        option.value = v.id;
        option.textContent = `${v.plate_number} - ${v.model}`;
        vehicleSelect.appendChild(option);
    });

    advanceSelect.innerHTML = '<option value="">لم ترتبط بعهدة</option>';
    appData.advance.filter(a => a.is_active).forEach(a => {
        const option = document.createElement('option');
        option.value = a.id;
        option.textContent = `عهدة ${a.amount} جنيه - ${a.advance_date}`;
        advanceSelect.appendChild(option);
    });

    if (expenseId) {
        title.textContent = 'تعديل بيانات النفقة';
        const expense = appData.expenses.find(e => e.id === expenseId);
        if (expense) {
            document.getElementById('expensesVehicle').value = expense.vehicle_id || '';
            document.getElementById('expensesType').value = expense.expense_type || '';
            document.getElementById('customExpenseType').value = expense.is_custom_type ? expense.expense_type : '';
            document.getElementById('expensesDate').value = expense.expense_date;
            document.getElementById('expensesAmount').value = expense.amount;
            document.getElementById('expensesAdvance').value = expense.advance_id || '';
            document.getElementById('expensesNotes').value = expense.notes || '';
        }
    } else {
        title.textContent = 'إضافة نفقة جديدة';
        form.reset();
        document.getElementById('customExpenseType').value = '';
    }

    modal.classList.add('show');
}

function closeExpensesModal() {
    document.getElementById('expensesModal').classList.remove('show');
    editingExpenseId = null;
    document.getElementById('expensesForm').reset();
}

function handleMaintenanceSubmit(e) {
    e.preventDefault();

    const vehicleId = document.getElementById('maintenanceVehicle').value;
    const vehicle = appData.vehicles.find(v => v.id === vehicleId);

    const maintenanceData = {
        id: Date.now().toString(),
        vehicle_id: vehicleId,
        plate_number: vehicle ? vehicle.plate_number : '',
        maintenance_type: document.getElementById('maintenanceType').value,
        maintenance_date: document.getElementById('maintenanceDate').value,
        cost: parseFloat(document.getElementById('maintenanceCost').value),
        status: document.getElementById('maintenanceStatus').value,
        notes: document.getElementById('maintenanceNotes').value
    };

    try {
        appData.maintenance.push(maintenanceData);
        saveData();
        closeMaintenanceModal();
        populateMaintenanceList();
        renderDashboard();
        alert('تم تسجيل الصيانة بنجاح');
    } catch (error) {
        console.error('خطأ:', error);
        alert('حدث خطأ في تسجيل الصيانة: ' + error.message);
    }
}

async function handleViolationSubmit(e) {
    e.preventDefault();

    const vehicleId = document.getElementById('violationVehicle').value;
    const vehicle = appData.vehicles.find(v => v.id === vehicleId);
    const imageFile = document.getElementById('violationImage').files[0];
    const existingViolation = editingViolationId ? appData.violations.find(v => v.id === editingViolationId) : null;

    try {
        const violationData = {
            id: editingViolationId || Date.now().toString(),
            vehicle_id: vehicleId,
            plate_number: vehicle ? vehicle.plate_number : '',
            violation_type: document.getElementById('violationType').value,
            violation_date: document.getElementById('violationDate').value,
            amount: parseFloat(document.getElementById('violationAmount').value),
            status: document.getElementById('violationStatus').value,
            image: imageFile ? await resizeImageFile(imageFile) : (existingViolation?.image || ''),
            notes: document.getElementById('violationNotes').value
        };

        if (editingViolationId) {
            const index = appData.violations.findIndex(v => v.id === editingViolationId);
            if (index !== -1) {
                appData.violations[index] = violationData;
            }
        } else {
            appData.violations.push(violationData);
        }

        saveData();
        closeViolationModal();
        populateViolationsList();
        renderDashboard();
        alert('تم حفظ المخالفة بنجاح');
    } catch (error) {
        console.error('خطأ:', error);
        alert('حدث خطأ في حفظ المخالفة: ' + error.message);
    }
}

function deleteMaintenance(maintenanceId) {
    if (confirm('هل أنت متأكد من حذف سجل الصيانة؟')) {
        try {
            appData.maintenance = appData.maintenance.filter(m => m.id !== maintenanceId);
            saveData();
            populateMaintenanceList();
            renderDashboard();
        } catch (error) {
            console.error('خطأ:', error);
            alert('حدث خطأ في حذف البيانات');
        }
    }
}

function deleteViolation(violationId) {
    if (confirm('هل أنت متأكد من حذف المخالفة؟')) {
        try {
            appData.violations = appData.violations.filter(v => v.id !== violationId);
            saveData();
            populateViolationsList();
            renderDashboard();
        } catch (error) {
            console.error('خطأ:', error);
            alert('حدث خطأ في حذف البيانات');
        }
    }
}

function handleExpensesSubmit(e) {
    e.preventDefault();

    const vehicleId = document.getElementById('expensesVehicle').value;
    const vehicle = vehicleId ? appData.vehicles.find(v => v.id === vehicleId) : null;

    let expenseType = document.getElementById('expensesType').value;
    let isCustom = false;
    const customType = document.getElementById('customExpenseType').value;

    if (customType) {
        expenseType = customType;
        isCustom = true;
    }

    const expenseData = {
        id: editingExpenseId || Date.now().toString(),
        vehicle_id: vehicleId || null,
        plate_number: vehicle ? vehicle.plate_number : '',
        expense_type: expenseType,
        is_custom_type: isCustom,
        expense_date: document.getElementById('expensesDate').value,
        amount: parseFloat(document.getElementById('expensesAmount').value),
        advance_id: document.getElementById('expensesAdvance').value || null,
        notes: document.getElementById('expensesNotes').value
    };

    try {
        if (editingExpenseId) {
            const index = appData.expenses.findIndex(e => e.id === editingExpenseId);
            if (index !== -1) {
                appData.expenses[index] = expenseData;
            }
        } else {
            appData.expenses.push(expenseData);
        }

        saveData();
        closeExpensesModal();
        populateExpensesList();
        populateAdvanceList();
        renderDashboard();
        alert('تم حفظ بيانات النفقة بنجاح');
    } catch (error) {
        console.error('خطأ:', error);
        alert('حدث خطأ في حفظ البيانات: ' + error.message);
    }
}

function deleteExpense(expenseId) {
    if (confirm('هل أنت متأكد من حذف هذه النفقة؟')) {
        try {
            appData.expenses = appData.expenses.filter(e => e.id !== expenseId);
            saveData();
            populateExpensesList();
            renderDashboard();
        } catch (error) {
            console.error('خطأ:', error);
            alert('حدث خطأ في حذف البيانات');
        }
    }
}

function openAdvanceModal(advanceId = null) {
    const modal = document.getElementById('advanceModal');
    const form = document.getElementById('advanceForm');
    const title = document.getElementById('advanceModalTitle');

    if (advanceId) {
        title.textContent = 'تعديل بيانات العهدة';
        const advance = appData.advance.find(a => a.id === advanceId);
        if (advance) {
            document.getElementById('advanceAmount').value = advance.amount;
            document.getElementById('advanceDate').value = advance.advance_date;
            document.getElementById('advanceNotes').value = advance.notes || '';
        }
    } else {
        title.textContent = 'إضافة عهدة جديدة';
        form.reset();
    }

    modal.classList.add('show');
}

function closeAdvanceModal() {
    document.getElementById('advanceModal').classList.remove('show');
    document.getElementById('advanceForm').reset();
}

function handleAdvanceSubmit(e) {
    e.preventDefault();

    const advanceData = {
        id: Date.now().toString(),
        amount: parseFloat(document.getElementById('advanceAmount').value),
        advance_date: document.getElementById('advanceDate').value,
        is_active: true,
        notes: document.getElementById('advanceNotes').value
    };

    try {
        appData.advance.push(advanceData);
        saveData();
        closeAdvanceModal();
        populateAdvanceList();
        renderDashboard();
        alert('تم إضافة العهدة بنجاح');
    } catch (error) {
        console.error('خطأ:', error);
        alert('حدث خطأ في حفظ البيانات: ' + error.message);
    }
}

function deleteAdvance(advanceId) {
    if (confirm('هل أنت متأكد من حذف هذه العهدة؟')) {
        try {
            appData.advance = appData.advance.filter(a => a.id !== advanceId);
            saveData();
            populateAdvanceList();
            renderDashboard();
        } catch (error) {
            console.error('خطأ:', error);
            alert('حدث خطأ في حذف البيانات');
        }
    }
}

function populateAdvanceList() {
    const tbody = document.getElementById('advanceTable');
    tbody.innerHTML = '';

    if (appData.advance.length === 0) {
        tbody.innerHTML = '<tr class="empty-row"><td colspan="5">لا توجد عهد مسجلة</td></tr>';
        updateAdvanceSummary();
        return;
    }

    appData.advance.forEach(advance => {
        const row = document.createElement('tr');
        const statusText = advance.is_active ? 'نشطة' : 'مستخدمة';
        const statusBadge = advance.is_active ? 'status-in-progress' : 'status-complete';
        
        row.innerHTML = `
            <td>${advance.amount.toLocaleString('en-US')} جنيه</td>
            <td>${advance.advance_date}</td>
            <td><span class="status-badge ${statusBadge}">${statusText}</span></td>
            <td>${advance.notes || '-'}</td>
            <td>
                <div class="action-buttons">
                    <button class="btn btn-danger btn-small" onclick="deleteAdvance('${advance.id}')">حذف</button>
                </div>
            </td>
        `;
        tbody.appendChild(row);
    });

    updateAdvanceSummary();
}

function updateAdvanceSummary() {
    const totalAdvance = appData.advance.reduce((sum, a) => sum + a.amount, 0);
    const usedAdvance = appData.expenses.reduce((sum, e) => sum + e.amount, 0);
    const remainingAdvance = totalAdvance - usedAdvance;

    document.getElementById('totalAdvance').textContent = totalAdvance.toLocaleString('en-US') + ' جنيه';
    document.getElementById('usedAdvance').textContent = usedAdvance.toLocaleString('en-US') + ' جنيه';
    document.getElementById('remainingAdvance').textContent = remainingAdvance.toLocaleString('en-US') + ' جنيه';
}

function populateExpensesList() {
    filterExpenses();
}

function filterExpenses() {
    const searchValue = document.getElementById('expensesSearch').value.toLowerCase();

    const filtered = appData.expenses.filter(e => {
        const vehiclePlate = e.plate_number ? e.plate_number.toLowerCase() : '';
        const matchesSearch = vehiclePlate.includes(searchValue) || 
                            (e.expense_type ? e.expense_type.toLowerCase().includes(searchValue) : false);
        return matchesSearch;
    });

    const tbody = document.getElementById('expensesTable');
    tbody.innerHTML = '';

    if (filtered.length === 0) {
        tbody.innerHTML = '<tr class="empty-row"><td colspan="6">لا توجد نفقات مسجلة</td></tr>';
        return;
    }

    filtered.forEach(expense => {
        const vehiclePlate = expense.plate_number || 'عام';

        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${vehiclePlate}</td>
            <td>${expense.expense_type}</td>
            <td>${expense.expense_date}</td>
            <td>${expense.amount.toLocaleString('en-US')} جنيه</td>
            <td>${expense.notes || '-'}</td>
            <td>
                <div class="action-buttons">
                    <button class="btn btn-primary btn-small" onclick="openExpensesModal('${expense.id}')">تعديل</button>
                    <button class="btn btn-danger btn-small" onclick="deleteExpense('${expense.id}')">حذف</button>
                </div>
            </td>
        `;
        tbody.appendChild(row);
    });
}

function renderDashboard() {
    const now = new Date();
    document.getElementById('totalVehicles').textContent = appData.vehicles.length;

    const expiringLicenses = appData.vehicles.filter(v => {
        const expiry = new Date(v.license_expiry);
        const daysLeft = Math.ceil((expiry - now) / (1000 * 60 * 60 * 24));
        return daysLeft <= licenseWarningDays && daysLeft > 0;
    }).length;
    document.getElementById('expiringLicenses').textContent = expiringLicenses;

    const inMaintenance = appData.maintenance.filter(m => m.status !== 'مكتملة').length;
    document.getElementById('inMaintenance').textContent = inMaintenance;

    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();
    const monthlyExpenses = appData.maintenance
        .filter(m => {
            const mDate = new Date(m.maintenance_date);
            return mDate.getMonth() === currentMonth && mDate.getFullYear() === currentYear;
        })
        .reduce((sum, m) => sum + m.cost, 0);
    document.getElementById('monthlyExpenses').textContent = monthlyExpenses.toLocaleString('en-US') + ' جنيه';

    const totalViolations = appData.violations.length;
    document.getElementById('totalViolations').textContent = totalViolations;

    const pendingViolations = appData.violations.filter(v => v.status === 'معلقة').length;
    document.getElementById('pendingViolations').textContent = pendingViolations;

    const handoversEl = document.getElementById('totalHandovers');
    if (handoversEl) handoversEl.textContent = appData.handovers.length;

    renderAllSummaries();
    renderAlerts();
    displayDashboardVehicles();
}

function renderAlerts() {
    const alertsList = document.getElementById('alertsList');
    alertsList.innerHTML = '';

    const now = new Date();
    const alerts = [];

    appData.vehicles.forEach(v => {
        const expiry = new Date(v.license_expiry);
        const daysLeft = Math.ceil((expiry - now) / (1000 * 60 * 60 * 24));

        if (daysLeft <= 0) {
            alerts.push({
                text: `🚨 المركبة ${v.plate_number} - انتهت رخصتها في ${v.license_expiry}`,
                critical: true
            });
        } else if (daysLeft <= licenseWarningDays) {
            alerts.push({
                text: `⚠️ المركبة ${v.plate_number} - تنتهي الرخصة في ${daysLeft} أيام`,
                critical: false
            });
        }

        if (v.status === 'معطوبة') {
            alerts.push({
                text: `⛔ المركبة ${v.plate_number} - في حالة معطوبة`,
                critical: true
            });
        }
    });

    if (alerts.length === 0) {
        alertsList.innerHTML = '<p class="empty-message">لا توجد تنبيهات</p>';
    } else {
        alerts.forEach(alert => {
            const div = document.createElement('div');
            div.className = `alert-item ${alert.critical ? 'critical' : ''}`;
            div.textContent = alert.text;
            alertsList.appendChild(div);
        });
    }
}

function populateVehiclesList() {
    filterVehicles();
}

function displayDashboardVehicles() {
    const tbody = document.getElementById('dashboardVehiclesTable');
    tbody.innerHTML = '';

    if (appData.vehicles.length === 0) {
        tbody.innerHTML = '<tr class="empty-row"><td colspan="6">لا توجد مركبات</td></tr>';
        return;
    }

    appData.vehicles.forEach(vehicle => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${vehicle.plate_number}</td>
            <td>${vehicle.model}</td>
            <td>${vehicle.vin_number || '-'}</td>
            <td>${vehicle.status}</td>
            <td>${vehicle.license_expiry}</td>
            <td>${vehicle.notes || '-'}</td>
        `;
        tbody.appendChild(row);
    });
}

function filterVehicles() {
    const searchValue = document.getElementById('vehicleSearch').value.toLowerCase();
    const statusFilter = document.getElementById('vehicleStatusFilter').value;

    const filtered = appData.vehicles.filter(v => {
        const matchesSearch = v.plate_number.toLowerCase().includes(searchValue) ||
                            v.model.toLowerCase().includes(searchValue);
        const matchesStatus = statusFilter === '' || v.status === statusFilter;
        return matchesSearch && matchesStatus;
    });

    const tbody = document.getElementById('vehiclesTable');
    tbody.innerHTML = '';

    if (filtered.length === 0) {
        tbody.innerHTML = '<tr class="empty-row"><td colspan="11">لا توجد مركبات</td></tr>';
        return;
    }

    filtered.forEach(vehicle => {
        const expiry = new Date(vehicle.license_expiry);
        const now = new Date();
        const daysLeft = Math.ceil((expiry - now) / (1000 * 60 * 60 * 24));

        let licenseStatus = '';
        if (daysLeft <= 0) {
            licenseStatus = '<span class="status-badge license-status-expired">منتهية</span>';
        } else if (daysLeft <= licenseWarningDays) {
            licenseStatus = `<span class="status-badge license-status-expiring">${daysLeft} أيام</span>`;
        } else {
            licenseStatus = '<span class="status-badge license-status-valid">سارية</span>';
        }

        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${vehicle.plate_number}</td>
            <td>${vehicle.model}</td>
            <td>${vehicle.year}</td>
            <td><span class="status-badge ${fleetBadgeClass(vehicle.status)}">${escapeHtml(vehicle.status)}</span></td>
            <td>${escapeHtml(vehicle.engine_number) || '-'}</td>
            <td>${escapeHtml(vehicle.chassis_number) || '-'}</td>
            <td>${getLicenseStart(vehicle) || '-'}</td>
            <td>${licenseStatus}</td>
            <td>${escapeHtml(lastReceiverName(vehicle)) || '-'}</td>
            <td>${escapeHtml(lastReceiveDate(vehicle)) || '-'}</td>
            <td>
                <div class="action-buttons">
                    <button class="btn btn-info btn-small" onclick="showHandoverHistory('${vehicle.id}')">سجل الاستلام</button>
                    <button class="btn btn-primary btn-small" onclick="openVehicleModal('${vehicle.id}')">تعديل</button>
                    <button class="btn btn-danger btn-small" onclick="deleteVehicle('${vehicle.id}')">حذف</button>
                </div>
            </td>
        `;
        tbody.appendChild(row);
    });
}

function populateMaintenanceList() {
    filterMaintenance();
}

function filterMaintenance() {
    const searchValue = document.getElementById('maintenanceSearch').value.toLowerCase();
    const statusFilter = document.getElementById('maintenanceStatusFilter').value;

    const filtered = appData.maintenance.filter(m => {
        const vehiclePlate = m.plate_number ? m.plate_number.toLowerCase() : '';
        const matchesSearch = vehiclePlate.includes(searchValue);
        const matchesStatus = statusFilter === '' || m.status === statusFilter;
        return matchesSearch && matchesStatus;
    });

    const tbody = document.getElementById('maintenanceTable');
    tbody.innerHTML = '';

    if (filtered.length === 0) {
        tbody.innerHTML = '<tr class="empty-row"><td colspan="7">لا توجد سجلات صيانة</td></tr>';
        return;
    }

    filtered.forEach(maintenance => {
        const vehiclePlate = maintenance.plate_number || 'غير محدد';

        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${vehiclePlate}</td>
            <td>${maintenance.maintenance_type}</td>
            <td>${maintenance.maintenance_date}</td>
            <td>${maintenance.cost.toLocaleString('en-US')} جنيه</td>
            <td><span class="status-badge ${maintenance.status === 'مكتملة' ? 'status-complete' : 'status-in-progress'}">${maintenance.status}</span></td>
            <td>${maintenance.notes || '-'}</td>
            <td>
                <div class="action-buttons">
                    <button class="btn btn-danger btn-small" onclick="deleteMaintenance('${maintenance.id}')">حذف</button>
                </div>
            </td>
        `;
        tbody.appendChild(row);
    });
}

function populateLicensesList() {
    filterLicenses();
}

function filterLicenses() {
    const searchValue = document.getElementById('licenseSearch').value.toLowerCase();
    const statusFilter = document.getElementById('licenseStatusFilter').value;

    const now = new Date();

    const filtered = appData.vehicles.filter(v => {
        const matchesSearch = v.plate_number.toLowerCase().includes(searchValue) ||
                            v.model.toLowerCase().includes(searchValue);

        const expiry = new Date(v.license_expiry);
        const daysLeft = Math.ceil((expiry - now) / (1000 * 60 * 60 * 24));

        let licenseStatus = '';
        if (daysLeft <= 0) {
            licenseStatus = 'منتهية';
        } else if (daysLeft <= licenseWarningDays) {
            licenseStatus = 'قريبة الانتهاء';
        } else {
            licenseStatus = 'سارية';
        }

        const matchesStatus = statusFilter === '' || licenseStatus === statusFilter;
        return matchesSearch && matchesStatus;
    });

    const tbody = document.getElementById('licensesTable');
    tbody.innerHTML = '';

    if (filtered.length === 0) {
        tbody.innerHTML = '<tr class="empty-row"><td colspan="7">لا توجد بيانات رخص</td></tr>';
        return;
    }

    filtered.forEach(vehicle => {
        const expiry = new Date(vehicle.license_expiry);
        const daysLeft = Math.ceil((expiry - now) / (1000 * 60 * 60 * 24));

        let licenseStatusBadge = '';
        let statusText = '';

        if (daysLeft <= 0) {
            licenseStatusBadge = '<span class="status-badge license-status-expired">منتهية</span>';
            statusText = 'منتهية';
        } else if (daysLeft <= licenseWarningDays) {
            licenseStatusBadge = `<span class="status-badge license-status-expiring">قريبة الانتهاء</span>`;
            statusText = 'قريبة الانتهاء';
        } else {
            licenseStatusBadge = '<span class="status-badge license-status-valid">سارية</span>';
            statusText = 'سارية';
        }

        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${vehicle.plate_number}</td>
            <td>${escapeHtml(vehicle.engine_number) || '-'}</td>
            <td>${escapeHtml(vehicle.chassis_number) || '-'}</td>
            <td>${getLicenseStart(vehicle) || '-'}</td>
            <td>${vehicle.license_expiry}</td>
            <td>${daysLeft > 0 ? daysLeft : 'منتهية'} ${daysLeft > 0 ? 'يوم' : ''}</td>
            <td>${licenseStatusBadge}</td>
        `;
        tbody.appendChild(row);
    });
}

function populateViolationsList() {
    filterViolations();
}

function filterViolations() {
    const searchValue = document.getElementById('violationSearch').value.toLowerCase();
    const statusFilter = document.getElementById('violationStatusFilter').value;

    const filtered = appData.violations.filter(v => {
        const vehiclePlate = v.plate_number ? v.plate_number.toLowerCase() : '';
        const matchesSearch = vehiclePlate.includes(searchValue);
        const matchesStatus = statusFilter === '' || v.status === statusFilter;
        return matchesSearch && matchesStatus;
    });

    const tbody = document.getElementById('violationsTable');
    tbody.innerHTML = '';

    if (filtered.length === 0) {
        tbody.innerHTML = '<tr class="empty-row"><td colspan="8">لا توجد مخالفات مسجلة</td></tr>';
        return;
    }

    filtered.forEach(violation => {
        const vehiclePlate = violation.plate_number || 'غير محدد';
        const imageCell = violation.image
            ? `<a class="violation-image-link" href="${violation.image}" target="_blank" rel="noopener noreferrer" title="عرض الصورة"><img src="${violation.image}" alt="صورة المخالفة" class="violation-image-thumb"></a>`
            : '<span class="no-image-text">لا توجد</span>';

        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${vehiclePlate}</td>
            <td>${violation.violation_type}</td>
            <td>${violation.violation_date}</td>
            <td>${violation.amount.toLocaleString("ar-SA")} جنيه</td>
            <td><span class="status-badge ${violation.status === "مسددة" ? "status-complete" : "status-in-progress"}">${violation.status}</span></td>
            <td class="violation-image-cell">${imageCell}</td>
            <td>${violation.notes || "-"}</td>
            <td>
                <div class="action-buttons">
                    <button class="btn btn-primary btn-small" onclick="openViolationModal('${violation.id}')">تعديل</button>
                    <button class="btn btn-danger btn-small" onclick="deleteViolation('${violation.id}')">حذف</button>
                </div>
            </td>
        `;
        tbody.appendChild(row);
    });
}

function generateVehiclesReportPDF() {
    const html = generateVehiclesReportHTML();
    const element = document.createElement('div');
    element.innerHTML = html;

    const options = {
        margin: 10,
        filename: 'تقرير_المركبات.pdf',
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2 },
        jsPDF: { orientation: 'landscape', unit: 'mm', format: 'a4' }
    };

    html2pdf().set(options).from(element).save();
}

function generateVehiclesReportHTML() {
    let html = `
        <div dir="rtl" style="font-family: 'Cairo', sans-serif; padding: 4px; margin: 0; position: relative;">
            <style>body, div, table, th, td, h1, h2, p { font-family: 'Cairo', sans-serif !important; }</style>
            <img src="5.jpg" style="position: absolute; top: 0; left: 0; width: 80px; height: 60px; object-fit: contain;">
            <div style="text-align: center; margin-bottom: 4px; border-bottom: 2px solid #2c3e50; padding-bottom: 3px; margin-top: 20px;">
                <h1 style="margin: 0; font-size: 16px; color: #2c3e50; font-weight: bold;">مصنع البهنساوي</h1>
                <h2 style="margin: 2px 0 0 0; font-size: 12px; color: #34495e;">تقرير المركبات - ${new Date().toLocaleDateString('ar-EG')}</h2>
            </div>
            ${reportChips(getSummary('vehicles'))}
            <table style="width: 100%; border-collapse: collapse; margin: 4px 0; line-height: 1.4;">
                <thead>
                    <tr style="background-color: #2c3e50; color: white; height: 18px;">
                        <th style="padding: 4px 6px; text-align: right; border: 2px solid #34495e; font-size: 12px; font-weight: bold;">اللوحة</th>
                        <th style="padding: 4px 6px; text-align: right; border: 2px solid #34495e; font-size: 12px; font-weight: bold;">النموذج</th>
                        <th style="padding: 4px 6px; text-align: right; border: 2px solid #34495e; font-size: 12px; font-weight: bold;">السائق</th>
                        <th style="padding: 4px 6px; text-align: right; border: 2px solid #34495e; font-size: 12px; font-weight: bold;">الحالة</th>
                        <th style="padding: 4px 6px; text-align: right; border: 2px solid #34495e; font-size: 12px; font-weight: bold;">رقم الموتور</th>
                        <th style="padding: 4px 6px; text-align: right; border: 2px solid #34495e; font-size: 12px; font-weight: bold;">رقم الشاسيه</th>
                        <th style="padding: 4px 6px; text-align: right; border: 2px solid #34495e; font-size: 12px; font-weight: bold;">بداية الرخصة</th>
                        <th style="padding: 4px 6px; text-align: right; border: 2px solid #34495e; font-size: 12px; font-weight: bold;">انتهاء الرخصة</th>
                        <th style="padding: 4px 6px; text-align: right; border: 2px solid #34495e; font-size: 12px; font-weight: bold;">المستلم الحالي</th>
                        <th style="padding: 4px 6px; text-align: right; border: 2px solid #34495e; font-size: 12px; font-weight: bold;">تاريخ الاستلام</th>
                    </tr>
                </thead>
                <tbody>
    `;

    appData.vehicles.forEach((vehicle, index) => {
        const bgColor = index % 2 === 0 ? '#ffffff' : '#f8f9fa';
        const statusColor = vehicle.status === 'اخضر' ? '#27ae60' : vehicle.status === 'ملاكي' ? '#3498db' : vehicle.status === 'نقل موظفين' ? '#f39c12' : '#6c757d';
        html += `<tr style="background-color: ${bgColor}; height: 16px;"><td style="padding: 3px 4px; border: 1px solid #dee2e6; font-size: 12px; font-weight: bold; color: #2c3e50;">${vehicle.plate_number}</td><td style="padding: 3px 4px; border: 1px solid #dee2e6; font-size: 12px; color: #495057;">${vehicle.model}</td><td style="padding: 3px 4px; border: 1px solid #dee2e6; font-size: 11px; color: #6c757d;">${vehicle.vin_number || '-'}</td><td style="padding: 3px 4px; border: 1px solid #dee2e6; font-size: 12px; color: white; background-color: ${statusColor}; font-weight: bold; text-align: center;">${vehicle.status}</td><td style="padding: 3px 4px; border: 1px solid #dee2e6; font-size: 12px; color: #495057;">${escapeHtml(vehicle.engine_number) || '-'}</td><td style="padding: 3px 4px; border: 1px solid #dee2e6; font-size: 12px; color: #495057;">${escapeHtml(vehicle.chassis_number) || '-'}</td><td style="padding: 3px 4px; border: 1px solid #dee2e6; font-size: 12px; color: #495057;">${getLicenseStart(vehicle) || '-'}</td><td style="padding: 3px 4px; border: 1px solid #dee2e6; font-size: 12px; color: #495057;">${vehicle.license_expiry}</td><td style="padding: 3px 4px; border: 1px solid #dee2e6; font-size: 12px; color: #495057;">${escapeHtml(lastReceiverName(vehicle)) || '-'}</td><td style="padding: 3px 4px; border: 1px solid #dee2e6; font-size: 12px; color: #495057;">${escapeHtml(lastReceiveDate(vehicle)) || '-'}</td></tr>`;
    });

    html += `</tbody></table></div>`;

    return html;
}

function generateVehiclesReportExcel() {
    const data = [
        ['تقرير المركبات', '', '', '', '', '', '', '', ''],
        ['التاريخ: ' + new Date().toLocaleDateString('ar-EG'), '', '', '', '', '', '', '', ''],
        [],
        ['اللوحة', 'النموذج', 'اسم السائق', 'السنة', 'الحالة', 'رقم الموتور', 'رقم الشاسيه', 'بداية الرخصة', 'انتهاء الرخصة', 'المستلم الحالي', 'تاريخ آخر استلام', 'الملاحظات']
    ];
    data.splice(3, 0, ...summaryRows('vehicles'), []);

    appData.vehicles.forEach(vehicle => {
        data.push([
            vehicle.plate_number,
            vehicle.model,
            vehicle.vin_number || '',
            vehicle.year,
            vehicle.status,
            vehicle.engine_number || '',
            vehicle.chassis_number || '',
            getLicenseStart(vehicle),
            vehicle.license_expiry,
            lastReceiverName(vehicle),
            lastReceiveDate(vehicle),
            vehicle.notes || ''
        ]);
    });

    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'المركبات');
    XLSX.writeFile(wb, 'تقرير_المركبات.xlsx');
}

function generateMaintenanceReportPDF() {
    const html = generateMaintenanceReportHTML();
    const element = document.createElement('div');
    element.innerHTML = html;

    const options = {
        margin: 10,
        filename: 'تقرير_الصيانة.pdf',
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2 },
        jsPDF: { orientation: 'landscape', unit: 'mm', format: 'a4' }
    };

    html2pdf().set(options).from(element).save();
}

function generateMaintenanceReportHTML() {
    let html = `
        <div dir="rtl" style="font-family: 'Cairo', sans-serif; padding: 20px; position: relative;">
            <style>body, div, table, th, td, h1, h2, p { font-family: 'Cairo', sans-serif !important; }</style>
            <img src="5.jpg" style="position: absolute; top: 0; left: 0; width: 80px; height: 60px; object-fit: contain;">
            <div style="text-align: center; margin-bottom: 30px; border-bottom: 2px solid #333; padding-bottom: 15px; margin-top: 20px;">
                <h1 style="margin: 0; font-size: 24px; font-weight: bold;">مصنع البهنساوي</h1>
                <h2 style="margin: 5px 0; font-size: 18px; font-weight: bold;">تقرير الصيانة والتصليح</h2>
                <p style="margin: 5px 0; font-size: 12px; color: #666;">
                    التاريخ: ${new Date().toLocaleDateString('ar-EG')}
                </p>
            </div>
            ${reportChips(getSummary('maintenance'))}
            <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
                <thead>
                    <tr style="background-color: #2c3e50; color: white;">
                        <th style="padding: 12px; text-align: right; border: 1px solid #ddd;">اللوحة</th>
                        <th style="padding: 12px; text-align: right; border: 1px solid #ddd;">نوع الصيانة</th>
                        <th style="padding: 12px; text-align: right; border: 1px solid #ddd;">التاريخ</th>
                        <th style="padding: 12px; text-align: right; border: 1px solid #ddd;">التكلفة</th>
                        <th style="padding: 12px; text-align: right; border: 1px solid #ddd;">الحالة</th>
                        <th style="padding: 12px; text-align: right; border: 1px solid #ddd;">الملاحظات</th>
                    </tr>
                </thead>
                <tbody>
    `;

    let totalCost = 0;
    appData.maintenance.forEach((maintenance, index) => {
        const vehiclePlate = maintenance.plate_number || 'غير محدد';
        const bgColor = index % 2 === 0 ? '#f9f9f9' : 'white';
        totalCost += maintenance.cost;

        html += `
            <tr style="background-color: ${bgColor};">
                <td style="padding: 10px; border: 1px solid #ddd;">${vehiclePlate}</td>
                <td style="padding: 10px; border: 1px solid #ddd;">${maintenance.maintenance_type}</td>
                <td style="padding: 10px; border: 1px solid #ddd;">${maintenance.maintenance_date}</td>
                <td style="padding: 10px; border: 1px solid #ddd;">${maintenance.cost.toLocaleString('en-US')} جنيه</td>
                <td style="padding: 10px; border: 1px solid #ddd;">${maintenance.status}</td>
                <td style="padding: 10px; border: 1px solid #ddd;">${maintenance.notes || '-'}</td>
            </tr>
        `;
    });

    html += `
                </tbody>
                <tfoot>
                    <tr style="background-color: #ecf0f1; font-weight: bold;">
                        <td colspan="3" style="padding: 12px; border: 1px solid #ddd; text-align: left;">إجمالي التكاليف:</td>
                        <td style="padding: 12px; border: 1px solid #ddd;">${totalCost.toLocaleString('en-US')} جنيه</td>
                        <td colspan="2" style="padding: 12px; border: 1px solid #ddd;"></td>
                    </tr>
                </tfoot>
            </table>
            <div style="margin-top: 30px; padding-top: 15px; border-top: 1px solid #ddd; text-align: center; font-size: 12px; color: #666;">
                <p>تم إنشاء هذا التقرير بواسطة نظام إدارة المركبات - مصنع البهنساوي</p>
            </div>
        </div>
    `;

    return html;
}

function generateMaintenanceReportExcel() {
    const data = [
        ['تقرير الصيانة والتصليح', '', '', '', '', ''],
        ['التاريخ: ' + new Date().toLocaleDateString('ar-EG'), '', '', '', '', ''],
        [],
        ['اللوحة', 'نوع الصيانة', 'التاريخ', 'التكلفة', 'الحالة', 'الملاحظات']
    ];
    data.splice(3, 0, ...summaryRows('maintenance'), []);

    let totalCost = 0;
    appData.maintenance.forEach(maintenance => {
        const vehiclePlate = maintenance.plate_number || 'غير محدد';
        totalCost += maintenance.cost;

        data.push([
            vehiclePlate,
            maintenance.maintenance_type,
            maintenance.maintenance_date,
            maintenance.cost,
            maintenance.status,
            maintenance.notes || ''
        ]);
    });

    data.push([]);
    data.push(['', '', 'إجمالي التكاليف:', totalCost, '', '']);

    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'الصيانة');
    XLSX.writeFile(wb, 'تقرير_الصيانة.xlsx');
}

function generateExpensesReportPDF() {
    const html = generateExpensesReportHTML();
    const element = document.createElement('div');
    element.innerHTML = html;

    const options = {
        margin: 10,
        filename: 'تقرير_المصروفات.pdf',
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2 },
        jsPDF: { orientation: 'portrait', unit: 'mm', format: 'a4' }
    };

    html2pdf().set(options).from(element).save();
}

function generateExpensesReportHTML() {
    let html = `
        <div dir="rtl" style="font-family: 'Cairo', sans-serif; padding: 8px; margin: 0; position: relative;">
            <style>@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@300;400;500;600;700;800;900&display=swap'); body, div, table, th, td, h1, h2, p { font-family: 'Cairo', sans-serif !important; }</style>
            <img src="5.jpg" style="position: absolute; top: 0; left: 0; width: 80px; height: 60px; object-fit: contain;">
            <div style="text-align: center; margin-bottom: 8px; border-bottom: 2px solid #1a252f; padding-bottom: 6px; margin-top: 20px;">
                <h1 style="margin: 0; font-size: 16px; color: #1a252f; font-weight: bold;">مصنع البهنساوي</h1>
                <h2 style="margin: 3px 0 0 0; font-size: 12px; color: #34495e; font-weight: bold;">تقرير المصروفات والعهدة الشامل</h2>
                <p style="margin: 2px 0; font-size: 9px; color: #666;">التاريخ: ${new Date().toLocaleDateString('ar-EG')}</p>
            </div>
    `;

    let totalMaintenance = 0;
    let totalViolations = 0;
    let totalExpenses = 0;
    let totalAdvance = 0;

    // مصاريف الصيانة
    appData.maintenance.forEach(m => {
        totalMaintenance += m.cost;
    });

    // مصاريف المخالفات
    appData.violations.forEach(v => {
        if (v.status === 'مسددة') {
            totalViolations += v.amount;
        }
    });

    // مصاريف النفقات
    appData.expenses.forEach(e => {
        totalExpenses += e.amount;
    });

    // مصاريف العهدة
    appData.advance.forEach(a => {
        totalAdvance += a.amount;
    });

    const expensesGrandTotal = totalMaintenance + totalViolations + totalExpenses;

    // جدول مصاريف الصيانة
    if (appData.maintenance.length > 0) {
        html += `
            <div style="margin-top: 8px;">
                <h3 style="margin: 4px 0 2px 0; font-size: 11px; color: white; background-color: #34495e; padding: 3px 6px; font-weight: bold;">🔧 مصاريف الصيانة والتصليح</h3>
                <table style="width: 100%; border-collapse: collapse; margin: 2px 0; font-size: 10px;">
                    <thead>
                        <tr style="background-color: #5d6d7b; color: white; height: 13px;">
                            <th style="padding: 2px 3px; border: 1px solid #666; text-align: right; font-weight: bold;">اللوحة</th>
                            <th style="padding: 2px 3px; border: 1px solid #666; text-align: right; font-weight: bold;">نوع الصيانة</th>
                            <th style="padding: 2px 3px; border: 1px solid #666; text-align: right; font-weight: bold;">التاريخ</th>
                            <th style="padding: 2px 3px; border: 1px solid #666; text-align: center; font-weight: bold;">المبلغ</th>
                        </tr>
                    </thead>
                    <tbody>
        `;
        
        appData.maintenance.forEach((m, idx) => {
            const bgColor = idx % 2 === 0 ? '#fafafa' : '#ffffff';
            html += `<tr style="background-color: ${bgColor}; height: 12px;"><td style="padding: 2px 3px; border: 0.5px solid #ddd; font-size: 10px; font-weight: bold;">${m.plate_number || '-'}</td><td style="padding: 2px 3px; border: 0.5px solid #ddd; font-size: 9px;">${m.maintenance_type}</td><td style="padding: 2px 3px; border: 0.5px solid #ddd; font-size: 9px;">${m.maintenance_date}</td><td style="padding: 2px 3px; border: 0.5px solid #ddd; text-align: center; font-weight: bold; font-size: 10px;">${m.cost.toLocaleString('en-US')}</td></tr>`;
        });

        html += `
                    </tbody>
                    <tfoot>
                        <tr style="background-color: #d4edda; font-weight: bold; height: 13px;">
                            <td colspan="3" style="padding: 2px 3px; border: 1px solid #999; text-align: right; font-size: 10px;">إجمالي</td>
                            <td style="padding: 2px 3px; border: 1px solid #999; text-align: center; color: #27ae60; font-size: 11px;">${totalMaintenance.toLocaleString('en-US')}</td>
                        </tr>
                    </tfoot>
                </table>
            </div>
        `;
    }

    // جدول مصاريف المخالفات المسددة
    const paidViolations = appData.violations.filter(v => v.status === 'مسددة');
    if (paidViolations.length > 0) {
        html += `
            <div style="margin-top: 8px;">
                <h3 style="margin: 4px 0 2px 0; font-size: 11px; color: white; background-color: #e74c3c; padding: 3px 6px; font-weight: bold;">⚠️ مصاريف المخالفات المسددة</h3>
                <table style="width: 100%; border-collapse: collapse; margin: 2px 0; font-size: 10px;">
                    <thead>
                        <tr style="background-color: #f1948b; color: white; height: 13px;">
                            <th style="padding: 2px 3px; border: 1px solid #666; text-align: right; font-weight: bold;">اللوحة</th>
                            <th style="padding: 2px 3px; border: 1px solid #666; text-align: right; font-weight: bold;">نوع المخالفة</th>
                            <th style="padding: 2px 3px; border: 1px solid #666; text-align: right; font-weight: bold;">التاريخ</th>
                            <th style="padding: 2px 3px; border: 1px solid #666; text-align: center; font-weight: bold;">المبلغ</th>
                        </tr>
                    </thead>
                    <tbody>
        `;
        
        paidViolations.forEach((v, idx) => {
            const bgColor = idx % 2 === 0 ? '#fafafa' : '#ffffff';
            html += `<tr style="background-color: ${bgColor}; height: 12px;"><td style="padding: 2px 3px; border: 0.5px solid #ddd; font-size: 10px; font-weight: bold;">${v.plate_number || '-'}</td><td style="padding: 2px 3px; border: 0.5px solid #ddd; font-size: 9px;">${v.violation_type}</td><td style="padding: 2px 3px; border: 0.5px solid #ddd; font-size: 9px;">${v.violation_date}</td><td style="padding: 2px 3px; border: 0.5px solid #ddd; text-align: center; font-weight: bold; font-size: 10px;">${v.amount.toLocaleString('en-US')}</td></tr>`;
        });

        html += `
                    </tbody>
                    <tfoot>
                        <tr style="background-color: #ffebee; font-weight: bold; height: 13px;">
                            <td colspan="3" style="padding: 2px 3px; border: 1px solid #999; text-align: right; font-size: 10px;">إجمالي</td>
                            <td style="padding: 2px 3px; border: 1px solid #999; text-align: center; color: #e74c3c; font-size: 11px;">${totalViolations.toLocaleString('en-US')}</td>
                        </tr>
                    </tfoot>
                </table>
            </div>
        `;
    }

    // جدول مصاريف النفقات
    if (appData.expenses.length > 0) {
        html += `
            <div style="margin-top: 8px;">
                <h3 style="margin: 4px 0 2px 0; font-size: 11px; color: white; background-color: #f39c12; padding: 3px 6px; font-weight: bold;">💰 مصاريف النفقات</h3>
                <table style="width: 100%; border-collapse: collapse; margin: 2px 0; font-size: 10px;">
                    <thead>
                        <tr style="background-color: #f5b041; color: white; height: 13px;">
                            <th style="padding: 2px 3px; border: 1px solid #666; text-align: right; font-weight: bold;">اللوحة</th>
                            <th style="padding: 2px 3px; border: 1px solid #666; text-align: right; font-weight: bold;">نوع النفقة</th>
                            <th style="padding: 2px 3px; border: 1px solid #666; text-align: right; font-weight: bold;">التاريخ</th>
                            <th style="padding: 2px 3px; border: 1px solid #666; text-align: center; font-weight: bold;">المبلغ</th>
                        </tr>
                    </thead>
                    <tbody>
        `;
        
        appData.expenses.forEach((e, idx) => {
            const bgColor = idx % 2 === 0 ? '#fafafa' : '#ffffff';
            html += `<tr style="background-color: ${bgColor}; height: 12px;"><td style="padding: 2px 3px; border: 0.5px solid #ddd; font-size: 10px; font-weight: bold;">${e.plate_number || '-'}</td><td style="padding: 2px 3px; border: 0.5px solid #ddd; font-size: 9px;">${e.expense_type}</td><td style="padding: 2px 3px; border: 0.5px solid #ddd; font-size: 9px;">${e.expense_date}</td><td style="padding: 2px 3px; border: 0.5px solid #ddd; text-align: center; font-weight: bold; font-size: 10px;">${e.amount.toLocaleString('en-US')}</td></tr>`;
        });

        html += `
                    </tbody>
                    <tfoot>
                        <tr style="background-color: #fffbea; font-weight: bold; height: 13px;">
                            <td colspan="3" style="padding: 2px 3px; border: 1px solid #999; text-align: right; font-size: 10px;">إجمالي</td>
                            <td style="padding: 2px 3px; border: 1px solid #999; text-align: center; color: #f39c12; font-size: 11px;">${totalExpenses.toLocaleString('en-US')}</td>
                        </tr>
                    </tfoot>
                </table>
            </div>
        `;
    }



    // الإجمالي العام للمصروفات (بدون العهدة)
    html += `
        <div style="margin-top: 10px; padding: 8px; background-color: #27ae60; color: white; border-radius: 2px;">
            <table style="width: 100%; border-collapse: collapse; font-size: 11px;">
                <tr style="height: 13px;">
                    <td style="padding: 3px 4px; text-align: right; border-bottom: 1px solid rgba(255,255,255,0.3);">🔧 مصاريف الصيانة</td>
                    <td style="padding: 3px 4px; text-align: center; font-weight: bold; border-bottom: 1px solid rgba(255,255,255,0.3);">${totalMaintenance.toLocaleString('en-US')}</td>
                </tr>
                <tr style="height: 13px;">
                    <td style="padding: 3px 4px; text-align: right; border-bottom: 1px solid rgba(255,255,255,0.3);">⚠️ مصاريف المخالفات المسددة</td>
                    <td style="padding: 3px 4px; text-align: center; font-weight: bold; border-bottom: 1px solid rgba(255,255,255,0.3);">${totalViolations.toLocaleString('en-US')}</td>
                </tr>
                <tr style="height: 13px;">
                    <td style="padding: 3px 4px; text-align: right; border-bottom: 2px solid white;">💰 مصاريف النفقات</td>
                    <td style="padding: 3px 4px; text-align: center; font-weight: bold; border-bottom: 2px solid white;">${totalExpenses.toLocaleString('en-US')}</td>
                </tr>
                <tr style="height: 15px;">
                    <td style="padding: 4px; font-size: 12px; text-align: right; font-weight: bold;">💵 إجمالي المصروفات</td>
                    <td style="padding: 4px; font-size: 13px; text-align: center; font-weight: bold; color: #fff000;">${expensesGrandTotal.toLocaleString('en-US')} جنيه</td>
                </tr>
            </table>
        </div>
    `;



    html += `</div>`;

    return html;
}

function generateExpensesReportExcel() {
    const now = new Date();
    const currentYear = now.getFullYear();

    const monthlyData = {};
    for (let i = 0; i < 12; i++) {
        monthlyData[i] = 0;
    }

    appData.maintenance.forEach(m => {
        const mDate = new Date(m.maintenance_date);
        if (mDate.getFullYear() === currentYear) {
            monthlyData[mDate.getMonth()] += m.cost;
        }
    });

    const monthNames = ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو',
                       'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر'];

    const data = [
        ['تقرير المصروفات', ''],
        ['السنة: ' + currentYear, ''],
        ['التاريخ: ' + new Date().toLocaleDateString('ar-EG'), ''],
        [],
        ['الشهر', 'المصروفات (جنيه)']
    ];

    let totalExpenses = 0;
    monthNames.forEach((month, index) => {
        const expense = monthlyData[index];
        totalExpenses += expense;
        data.push([month, expense]);
    });

    data.push([]);
    data.push(['الإجمالي السنوي', totalExpenses]);
    data.push(['متوسط الشهري', totalExpenses / 12]);

    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'المصروفات');
    XLSX.writeFile(wb, 'تقرير_المصروفات.xlsx');
}

function generateLicensesReportPDF() {
    const html = generateLicensesReportHTML();
    const element = document.createElement('div');
    element.innerHTML = html;

    const options = {
        margin: 10,
        filename: 'تقرير_الرخص.pdf',
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2 },
        jsPDF: { orientation: 'landscape', unit: 'mm', format: 'a4' }
    };

    html2pdf().set(options).from(element).save();
}

function generateLicensesReportHTML() {
    const now = new Date();

    let html = `
        <div dir="rtl" style="font-family: 'Cairo', sans-serif; padding: 20px; position: relative;">
            <style>body, div, table, th, td, h1, h2, p { font-family: 'Cairo', sans-serif !important; }</style>
            <img src="5.jpg" style="position: absolute; top: 0; left: 0; width: 80px; height: 60px; object-fit: contain;">
            <div style="text-align: center; margin-bottom: 30px; border-bottom: 2px solid #333; padding-bottom: 15px; margin-top: 20px;">
                <h1 style="margin: 0; font-size: 24px; font-weight: bold;">مصنع البهنساوي</h1>
                <h2 style="margin: 5px 0; font-size: 18px; font-weight: bold;">تقرير الرخص</h2>
                <p style="margin: 5px 0; font-size: 12px; color: #666;">
                    التاريخ: ${new Date().toLocaleDateString('ar-EG')}
                </p>
            </div>
            ${reportChips(getSummary('licenses'))}
            <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
                <thead>
                    <tr style="background-color: #2c3e50; color: white;">
                        <th style="padding: 12px; text-align: right; border: 1px solid #ddd;">اللوحة</th>
                        <th style="padding: 12px; text-align: right; border: 1px solid #ddd;">رقم الموتور</th>
                        <th style="padding: 12px; text-align: right; border: 1px solid #ddd;">رقم الشاسيه</th>
                        <th style="padding: 12px; text-align: right; border: 1px solid #ddd;">بداية الرخصة</th>
                        <th style="padding: 12px; text-align: right; border: 1px solid #ddd;">تاريخ الانتهاء</th>
                        <th style="padding: 12px; text-align: right; border: 1px solid #ddd;">الأيام المتبقية</th>
                        <th style="padding: 12px; text-align: right; border: 1px solid #ddd;">الحالة</th>
                    </tr>
                </thead>
                <tbody>
    `;

    appData.vehicles.forEach((vehicle, index) => {
        const expiry = new Date(vehicle.license_expiry);
        const daysLeft = Math.ceil((expiry - now) / (1000 * 60 * 60 * 24));
        const bgColor = index % 2 === 0 ? '#f9f9f9' : 'white';

        let status = 'سارية';
        if (daysLeft <= 0) {
            status = 'منتهية';
        } else if (daysLeft <= licenseWarningDays) {
            status = 'قريبة الانتهاء';
        }

        html += `
            <tr style="background-color: ${bgColor};">
                <td style="padding: 10px; border: 1px solid #ddd;">${vehicle.plate_number}</td>
                <td style="padding: 10px; border: 1px solid #ddd;">${escapeHtml(vehicle.engine_number) || '-'}</td>
                <td style="padding: 10px; border: 1px solid #ddd;">${escapeHtml(vehicle.chassis_number) || '-'}</td>
                <td style="padding: 10px; border: 1px solid #ddd;">${getLicenseStart(vehicle) || '-'}</td>
                <td style="padding: 10px; border: 1px solid #ddd;">${vehicle.license_expiry}</td>
                <td style="padding: 10px; border: 1px solid #ddd; text-align: center;">${daysLeft > 0 ? daysLeft : 'منتهية'}</td>
                <td style="padding: 10px; border: 1px solid #ddd;">${status}</td>
            </tr>
        `;
    });

    html += `
                </tbody>
            </table>
            <div style="margin-top: 30px; padding-top: 15px; border-top: 1px solid #ddd; text-align: center; font-size: 12px; color: #666;">
                <p>تم إنشاء هذا التقرير بواسطة نظام إدارة المركبات - مصنع البهنساوي</p>
            </div>
        </div>
    `;

    return html;
}

function generateLicensesReportExcel() {
    const now = new Date();

    const data = [
        ['تقرير الرخص', '', '', '', '', '', ''],
        ['التاريخ: ' + new Date().toLocaleDateString('ar-EG'), '', '', '', '', '', ''],
        [],
        ['اللوحة', 'رقم الموتور', 'رقم الشاسيه', 'بداية الرخصة', 'تاريخ الانتهاء', 'الأيام المتبقية', 'الحالة']
    ];
    data.splice(3, 0, ...summaryRows('licenses'), []);

    appData.vehicles.forEach(vehicle => {
        const expiry = new Date(vehicle.license_expiry);
        const daysLeft = Math.ceil((expiry - now) / (1000 * 60 * 60 * 24));

        let status = 'سارية';
        if (daysLeft <= 0) {
            status = 'منتهية';
        } else if (daysLeft <= licenseWarningDays) {
            status = 'قريبة الانتهاء';
        }

        data.push([
            vehicle.plate_number,
            vehicle.engine_number || '',
            vehicle.chassis_number || '',
            getLicenseStart(vehicle),
            vehicle.license_expiry,
            daysLeft > 0 ? daysLeft : 'منتهية',
            status
        ]);
    });

    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'الرخص');
    XLSX.writeFile(wb, 'تقرير_الرخص.xlsx');
}

function generateViolationsReportPDF() {
    const html = generateViolationsReportHTML();
    const element = document.createElement('div');
    element.innerHTML = html;

    const options = {
        margin: 10,
        filename: 'تقرير_المخالفات.pdf',
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2 },
        jsPDF: { orientation: 'landscape', unit: 'mm', format: 'a4' }
    };

    html2pdf().set(options).from(element).save();
}

function generateViolationsReportHTML() {
    let html = `
        <div dir="rtl" style="font-family: 'Cairo', sans-serif; padding: 20px; position: relative;">
            <style>body, div, table, th, td, h1, h2, p { font-family: 'Cairo', sans-serif !important; }</style>
            <img src="5.jpg" style="position: absolute; top: 0; left: 0; width: 80px; height: 60px; object-fit: contain;">
            <div style="text-align: center; margin-bottom: 30px; border-bottom: 2px solid #333; padding-bottom: 15px; margin-top: 20px;">
                <h1 style="margin: 0; font-size: 24px; font-weight: bold;">مصنع البهنساوي</h1>
                <h2 style="margin: 5px 0; font-size: 18px; font-weight: bold;">تقرير المخالفات المرورية</h2>
                <p style="margin: 5px 0; font-size: 12px; color: #666;">
                    التاريخ: ${new Date().toLocaleDateString("ar-EG")}
                </p>
            </div>
            ${reportChips(getSummary('violations'))}
            <table style="width: 100%; border-collapse: collapse; margin: 20px 0;">
                <thead>
                    <tr style="background-color: #2c3e50; color: white;">
                        <th style="padding: 12px; text-align: right; border: 1px solid #ddd;">اللوحة</th>
                        <th style="padding: 12px; text-align: right; border: 1px solid #ddd;">نوع المخالفة</th>
                        <th style="padding: 12px; text-align: right; border: 1px solid #ddd;">التاريخ</th>
                        <th style="padding: 12px; text-align: right; border: 1px solid #ddd;">المبلغ (جنيه)</th>
                        <th style="padding: 12px; text-align: right; border: 1px solid #ddd;">الحالة</th>
                        <th style="padding: 12px; text-align: center; border: 1px solid #ddd;">الصورة</th>
                        <th style="padding: 12px; text-align: right; border: 1px solid #ddd;">الملاحظات</th>
                    </tr>
                </thead>
                <tbody>
    `;

    let totalAmount = 0;
    let totalPending = 0;

    appData.violations.forEach((violation, index) => {
        const vehiclePlate = violation.plate_number || 'غير محدد';
        const bgColor = index % 2 === 0 ? '#f9f9f9' : 'white';

        totalAmount += violation.amount;
        if (violation.status === 'معلقة') totalPending += violation.amount;

        html += `
            <tr style="background-color: ${bgColor};">
                <td style="padding: 10px; border: 1px solid #ddd;">${vehiclePlate}</td>
                <td style="padding: 10px; border: 1px solid #ddd;">${violation.violation_type}</td>
                <td style="padding: 10px; border: 1px solid #ddd;">${violation.violation_date}</td>
                <td style="padding: 10px; border: 1px solid #ddd;">${violation.amount.toLocaleString("ar-SA")}</td>
                <td style="padding: 10px; border: 1px solid #ddd;">${violation.status}</td>
                <td style="padding: 10px; border: 1px solid #ddd; text-align: center;">${violation.image ? `<img src="${violation.image}" alt="صورة المخالفة" style="width: 130px; height: 100px; object-fit: cover; border-radius: 8px; border: 1px solid #ddd;">` : "لا توجد"}</td>
                <td style="padding: 10px; border: 1px solid #ddd;">${violation.notes || "-"}</td>
            </tr>
        `;
    });

    html += `
                </tbody>
                <tfoot>
                    <tr style="background-color: #ecf0f1; font-weight: bold;">
                        <td colspan="3" style="padding: 12px; border: 1px solid #ddd; text-align: left;">الإجمالي:</td>
                        <td style="padding: 12px; border: 1px solid #ddd;">${totalAmount.toLocaleString("ar-SA")} جنيه</td>
                        <td colspan="3" style="padding: 12px; border: 1px solid #ddd;"></td>
                    </tr>
                    <tr style="background-color: #fff3cd;">
                        <td colspan="3" style="padding: 12px; border: 1px solid #ddd; text-align: left;">المعلق من المخالفات:</td>
                        <td style="padding: 12px; border: 1px solid #ddd;">${totalPending.toLocaleString("ar-SA")} جنيه</td>
                        <td colspan="3" style="padding: 12px; border: 1px solid #ddd;"></td>
                    </tr>
                </tfoot>
            </table>
            <div style="margin-top: 30px; padding-top: 15px; border-top: 1px solid #ddd; text-align: center; font-size: 12px; color: #666;">
                <p>تم إنشاء هذا التقرير بواسطة نظام إدارة المركبات - مصنع البهنساوي</p>
            </div>
        </div>
    `;

    return html;
}

function generateViolationsReportExcel() {
    const data = [
        ['تقرير المخالفات المرورية', '', '', '', '', '', ''],
        ['التاريخ: ' + new Date().toLocaleDateString('ar-EG'), '', '', '', '', '', ''],
        [],
        ['اللوحة', 'نوع المخالفة', 'التاريخ', 'المبلغ (جنيه)', 'الحالة', 'الصورة', 'الملاحظات']
    ];
    data.splice(3, 0, ...summaryRows('violations'), []);

    let totalAmount = 0;
    let totalPending = 0;

    appData.violations.forEach(violation => {
        const vehiclePlate = violation.plate_number || 'غير محدد';

        totalAmount += violation.amount;
        if (violation.status === 'معلقة') totalPending += violation.amount;

        data.push([
            vehiclePlate,
            violation.violation_type,
            violation.violation_date,
            violation.amount,
            violation.status,
            violation.image ? 'مرفقة' : 'لا توجد',
            violation.notes || ''
        ]);
    });

    data.push([]);
    data.push(['', '', 'الإجمالي:', totalAmount, '', '', '']);
    data.push(['', '', 'المعلق:', totalPending, '', '', '']);

    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'المخالفات');
    XLSX.writeFile(wb, 'تقرير_المخالفات.xlsx');
}

function generateAdvanceExpensesReportPDF() {
    const html = generateAdvanceExpensesReportHTML();
    const element = document.createElement('div');
    element.innerHTML = html;

    const options = {
        margin: 10,
        filename: 'تقرير_النفقات_والعهدة.pdf',
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2 },
        jsPDF: { orientation: 'portrait', unit: 'mm', format: 'a4' }
    };

    html2pdf().set(options).from(element).save();
}

function generateAdvanceExpensesReportHTML() {
    const totalAdvance = appData.advance.reduce((sum, a) => sum + a.amount, 0);
    const usedAdvance = appData.expenses.reduce((sum, e) => sum + e.amount, 0);
    const remainingAdvance = totalAdvance - usedAdvance;

    let expensesTable = '';
    appData.expenses.forEach((expense, index) => {
        const bgColor = index % 2 === 0 ? '#f9f9f9' : 'white';
        const vehiclePlate = expense.plate_number || 'عام';
        
        expensesTable += `
            <tr style="background-color: ${bgColor};">
                <td style="padding: 6px 8px; border: 1px solid #ccc;">${vehiclePlate}</td>
                <td style="padding: 6px 8px; border: 1px solid #ccc;">${expense.expense_type}</td>
                <td style="padding: 6px 8px; border: 1px solid #ccc;">${expense.expense_date}</td>
                <td style="padding: 6px 8px; border: 1px solid #ccc; text-align: center; font-weight: 500;">${expense.amount.toLocaleString('en-US')}</td>
                <td style="padding: 6px 8px; border: 1px solid #ccc;">${expense.notes || '-'}</td>
            </tr>
        `;
    });
    
    expensesTable += `
        <tr style="background-color: #2c3e50; color: white; font-weight: bold; font-size: 12px;">
            <td style="padding: 6px 8px; border: 1px solid #999;"></td>
            <td style="padding: 6px 8px; border: 1px solid #999;"></td>
            <td style="padding: 6px 8px; border: 1px solid #999; text-align: right;">الإجمالي:</td>
            <td style="padding: 6px 8px; border: 1px solid #999; text-align: center; font-weight: bold;">${usedAdvance.toLocaleString('en-US')}</td>
            <td style="padding: 6px 8px; border: 1px solid #999;"></td>
        </tr>
    `;

    let html = `
        <div dir="rtl" style="font-family: 'Cairo', sans-serif; padding: 10px; line-height: 1.4; position: relative;">
            <style>@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@300;400;500;600;700;800;900&display=swap'); body, div, table, th, td, h1, h2, p { font-family: 'Cairo', sans-serif !important; }</style>
            <img src="5.jpg" style="position: absolute; top: 0; left: 0; width: 80px; height: 60px; object-fit: contain;">
            <div style="text-align: center; margin-bottom: 10px; border-bottom: 3px solid #2c3e50; padding-bottom: 8px; margin-top: 20px;">
                <h1 style="margin: 0; font-size: 24px; font-weight: bold;">مصنع البهنساوي</h1>
                <h2 style="margin: 2px 0; font-size: 15px; font-weight: bold;">  تقرير المصروفات والعهدة لسيارات البن المطحون</h2>
                <p style="margin: 2px 0; font-size: 10px; color: #666;">
                    التاريخ: ${new Date().toLocaleDateString('ar-EG')}
                </p>
            </div>

            <div style="margin-bottom: 8px; background-color: #f5f5f5; padding: 6px; border-radius: 3px; border-left: 4px solid #2c3e50;">
                <h3 style="margin: 0 0 5px 0; font-size: 14px; font-weight: bold; color: #2c3e50;">ملخص العهدة والمصروفات </h3>
                <table style="width: 100%; border-collapse: collapse; font-size: 13px;">
                    <tr>
                        <td style="padding: 6px 8px; border: 1px solid #bbb; background-color: #e8f4f8; font-weight: bold; width: 50%;">العهدة المتاحة:</td>
                        <td style="padding: 6px 8px; border: 1px solid #bbb; text-align: center; font-weight: bold; font-size: 14px;">${totalAdvance.toLocaleString('en-US')}</td>
                    </tr>
                    <tr>
                        <td style="padding: 6px 8px; border: 1px solid #bbb; background-color: #fff3cd; font-weight: bold;">المصروف:</td>
                        <td style="padding: 6px 8px; border: 1px solid #bbb; text-align: center; font-weight: bold; color: #c00; font-size: 14px;">${usedAdvance.toLocaleString('en-US')}</td>
                    </tr>
                    <tr>
                        <td style="padding: 6px 8px; border: 1px solid #bbb; background-color: #d4edda; font-weight: bold;">المتبقي:</td>
                        <td style="padding: 6px 8px; border: 1px solid #bbb; text-align: center; font-weight: bold; color: #060; font-size: 14px;">${remainingAdvance.toLocaleString('en-US')}</td>
                    </tr>
                </table>
            </div>

            <div style="margin-bottom: 2px;">
                <h3 style="margin: 0 0 4px 0; font-size: 14px; font-weight: bold; color: #2c3e50; border-bottom: 2px solid #2c3e50; padding-bottom: 2px;">تفاصيل المصروفات </h3>
                <table style="width: 100%; border-collapse: collapse; font-size: 14px;">
                    <thead>
                        <tr style="background-color: #2c3e50; color: white; font-weight: bold;">
                            <th style="padding: 8px 10px; text-align: right; border: 1px solid #999;">اللوحة</th>
                            <th style="padding: 8px 10px; text-align: right; border: 1px solid #999;">نوع المصروفات</th>
                            <th style="padding: 8px 10px; text-align: right; border: 1px solid #999;">التاريخ</th>
                            <th style="padding: 8px 10px; text-align: right; border: 1px solid #999;">المبلغ</th>
                            <th style="padding: 8px 10px; text-align: right; border: 1px solid #999;">ملاحظات</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${expensesTable || '<tr><td colspan="5" style="padding: 8px; text-align: center; color: #999;">لا توجد نفقات</td></tr>'}
                    </tbody>
                </table>
            </div>

            <div style="margin-top: 8px; padding: 7px 0; border-top: 2px solid #333;">
                <div style="display: flex; justify-content: space-between; align-items: flex-start;">
                    <div style="text-align: center; width: 45%;">
                        <p style="margin: 0 0 16px 0; font-size: 12px;">التوقيع</p>
                        <p style="margin: 0; font-size: 12px; font-weight: bold; border-top: 2px solid #333; padding-top: 4px;">مدير الحركة</p>
                    </div>
                    <div style="text-align: center; width: 45%;">
                        <p style="margin: 0 0 16px 0; font-size: 12px;">التوقيع</p>
                        <p style="margin: 0; font-size: 12px; font-weight: bold; border-top: 2px solid #333; padding-top: 4px;">المدير المالي</p>
                    </div>
                </div>
            </div>

            <div style="margin-top: 3px; padding-top: 3px; border-top: 1px solid #ccc; text-align: center; font-size: 9px; color: #777;">
                <p style="margin: 0;">نظام إدارة المركبات - مصنع البهنساوي</p>
            </div>
        </div>
    `;

    return html;
}

function generateAdvanceExpensesReportExcel() {
    const totalAdvance = appData.advance.reduce((sum, a) => sum + a.amount, 0);
    const usedAdvance = appData.expenses.reduce((sum, e) => sum + e.amount, 0);
    const remainingAdvance = totalAdvance - usedAdvance;

    const data = [
        ['تقرير النفقات والعهدة', '', '', '', ''],
        ['التاريخ: ' + new Date().toLocaleDateString('ar-EG'), '', '', '', ''],
        [],
        ['ملخص العهدة والنفقات', '', '', '', ''],
        ['العهدة المتاحة', totalAdvance, '', '', ''],
        ['المصروف من العهدة', usedAdvance, '', '', ''],
        ['المتبقي من العهدة', remainingAdvance, '', '', ''],
        [],
        ['تفاصيل النفقات', '', '', '', ''],
        ['اللوحة', 'نوع النفقة', 'التاريخ', 'المبلغ (جنيه)', 'الملاحظات']
    ];

    appData.expenses.forEach(expense => {
        const vehiclePlate = expense.plate_number || 'عام';
        data.push([
            vehiclePlate,
            expense.expense_type,
            expense.expense_date,
            expense.amount,
            expense.notes || ''
        ]);
    });

    data.push(['الإجمالي', '', '', usedAdvance, '']);

    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'النفقات والعهدة');
    XLSX.writeFile(wb, 'تقرير_النفقات_والعهدة.xlsx');
}

function printAdvanceExpensesReport() {
    const html = generateAdvanceExpensesReportHTML();
    const printWindow = window.open('', '', 'width=1200,height=800');
    printWindow.document.write(html);
    printWindow.document.close();
    printWindow.print();
}

function confirmClearAllData() {
    if (confirm('تحذير: هذا سيحذف جميع بياناتك بشكل نهائي. هل أنت متأكد؟')) {
        if (confirm('هل أنت متأكد تماماً؟ لا يمكن التراجع عن هذا الإجراء!')) {
            try {
                appData.vehicles = [];
                appData.maintenance = [];
                appData.violations = [];
                appData.expenses = [];
                appData.advance = [];
                appData.tires_in = [];
                appData.tires_out = [];
                appData.movements = [];
                appData.handovers = [];
                appData.penalties = [];
                saveData();
                alert('تم حذف جميع البيانات');
                renderDashboard();
            } catch (error) {
                console.error('خطأ:', error);
                alert('حدث خطأ في حذف البيانات');
            }
        }
    }
}

function printVehiclesReport() {
    const html = generateVehiclesReportHTML();
    const printWindow = window.open('', '', 'width=1200,height=800');
    printWindow.document.write(html);
    printWindow.document.close();
    printWindow.print();
}

function printMaintenanceReport() {
    const html = generateMaintenanceReportHTML();
    const printWindow = window.open('', '', 'width=1200,height=800');
    printWindow.document.write(html);
    printWindow.document.close();
    printWindow.print();
}

function printExpensesReport() {
    const html = generateExpensesReportHTML();
    const printWindow = window.open('', '', 'width=1200,height=800');
    printWindow.document.write(html);
    printWindow.document.close();
    printWindow.print();
}

function printLicensesReport() {
    const html = generateLicensesReportHTML();
    const printWindow = window.open('', '', 'width=1200,height=800');
    printWindow.document.write(html);
    printWindow.document.close();
    printWindow.print();
}

function printViolationsReport() {
    const html = generateViolationsReportHTML();
    const printWindow = window.open('', '', 'width=1200,height=800');
    printWindow.document.write(html);
    printWindow.document.close();
    printWindow.print();
}

function printPage() {
    window.print();
}

function openExpensesFilterModal() {
    document.getElementById('expensesFilterModal').classList.add('show');
}

function closeExpensesFilterModal() {
    document.getElementById('expensesFilterModal').classList.remove('show');
}

function getFilteredExpensesData() {
    const fromDate = document.getElementById('expensesFromDate').value;
    const toDate = document.getElementById('expensesToDate').value;
    const fromAmount = parseFloat(document.getElementById('expensesFromAmount').value) || 0;
    const toAmount = parseFloat(document.getElementById('expensesToAmount').value) || Infinity;

    let filteredMaintenance = appData.maintenance;
    let filteredViolations = appData.violations.filter(v => v.status === 'مسددة');
    let filteredExpenses = appData.expenses;
    let filteredAdvance = appData.advance;

    if (fromDate) {
        filteredMaintenance = filteredMaintenance.filter(m => new Date(m.maintenance_date) >= new Date(fromDate));
        filteredViolations = filteredViolations.filter(v => new Date(v.violation_date) >= new Date(fromDate));
        filteredExpenses = filteredExpenses.filter(e => new Date(e.expense_date) >= new Date(fromDate));
        filteredAdvance = filteredAdvance.filter(a => new Date(a.advance_date) >= new Date(fromDate));
    }

    if (toDate) {
        filteredMaintenance = filteredMaintenance.filter(m => new Date(m.maintenance_date) <= new Date(toDate));
        filteredViolations = filteredViolations.filter(v => new Date(v.violation_date) <= new Date(toDate));
        filteredExpenses = filteredExpenses.filter(e => new Date(e.expense_date) <= new Date(toDate));
        filteredAdvance = filteredAdvance.filter(a => new Date(a.advance_date) <= new Date(toDate));
    }

    if (fromAmount > 0 || toAmount !== Infinity) {
        filteredMaintenance = filteredMaintenance.filter(m => m.cost >= fromAmount && m.cost <= toAmount);
        filteredViolations = filteredViolations.filter(v => v.amount >= fromAmount && v.amount <= toAmount);
        filteredExpenses = filteredExpenses.filter(e => e.amount >= fromAmount && e.amount <= toAmount);
        filteredAdvance = filteredAdvance.filter(a => a.amount >= fromAmount && a.amount <= toAmount);
    }

    return { filteredMaintenance, filteredViolations, filteredExpenses, filteredAdvance };
}

function generateFilteredExpensesReportPDF() {
    const html = generateFilteredExpensesReportHTML();
    const element = document.createElement('div');
    element.innerHTML = html;

    const options = {
        margin: 10,
        filename: 'تقرير_المصروفات_مصفاة.pdf',
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2 },
        jsPDF: { orientation: 'portrait', unit: 'mm', format: 'a4' }
    };

    html2pdf().set(options).from(element).save();
    closeExpensesFilterModal();
}

function generateFilteredExpensesReportHTML() {
    const { filteredMaintenance, filteredViolations, filteredExpenses, filteredAdvance } = getFilteredExpensesData();
    
    let html = `
        <div dir="rtl" style="font-family: 'Cairo', sans-serif; padding: 10px; margin: 0; position: relative;">
            <style>@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@300;400;500;600;700;800;900&display=swap'); body, div, table, th, td, h1, h2, p { font-family: 'Cairo', sans-serif !important; }</style>
            <img src="5.jpg" style="position: absolute; top: 0; left: 0; width: 80px; height: 60px; object-fit: contain;">
            <div style="text-align: center; margin-bottom: 12px; border-bottom: 2px solid #333; padding-bottom: 8px; margin-top: 20px;">
                <h1 style="margin: 0; font-size: 18px; color: #1a252f; font-weight: bold;">مصنع البهنساوي</h1>
                <h2 style="margin: 4px 0 0 0; font-size: 13px; color: #34495e; font-weight: bold;">تقرير المصروفات (مصفاة)</h2>
                <p style="margin: 3px 0; font-size: 10px; color: #666;">التاريخ: ${new Date().toLocaleDateString('ar-EG')}</p>
            </div>
    `;

    let totalMaintenance = 0;
    let totalViolations = 0;
    let totalExpenses = 0;
    let totalAdvance = 0;

    filteredMaintenance.forEach(m => { totalMaintenance += m.cost; });
    filteredViolations.forEach(v => { totalViolations += v.amount; });
    filteredExpenses.forEach(e => { totalExpenses += e.amount; });
    filteredAdvance.forEach(a => { totalAdvance += a.amount; });

    const grandTotal = totalMaintenance + totalViolations + totalExpenses;

    if (filteredMaintenance.length > 0) {
        html += `
            <div style="margin-top: 8px;">
                <h3 style="margin: 5px 0; font-size: 12px; color: white; background-color: #34495e; padding: 4px 8px; font-weight: bold;">🔧 مصاريف الصيانة والتصليح</h3>
                <table style="width: 100%; border-collapse: collapse; margin: 4px 0; font-size: 10px;">
                    <thead>
                        <tr style="background-color: #5d6d7b; color: white; height: 12px;">
                            <th style="padding: 3px 4px; border: 0.5px solid #999; text-align: right;">اللوحة</th>
                            <th style="padding: 3px 4px; border: 0.5px solid #999; text-align: right;">نوع الصيانة</th>
                            <th style="padding: 3px 4px; border: 0.5px solid #999; text-align: right;">التاريخ</th>
                            <th style="padding: 3px 4px; border: 0.5px solid #999; text-align: center;">المبلغ</th>
                        </tr>
                    </thead>
                    <tbody>
        `;
        
        filteredMaintenance.forEach((m, idx) => {
            const bgColor = idx % 2 === 0 ? '#ffffff' : '#f5f5f5';
            html += `<tr style="background-color: ${bgColor}; height: 11px;"><td style="padding: 2px 3px; border: 0.5px solid #ddd;">${m.plate_number || '-'}</td><td style="padding: 2px 3px; border: 0.5px solid #ddd; font-size: 9px;">${m.maintenance_type}</td><td style="padding: 2px 3px; border: 0.5px solid #ddd; font-size: 9px;">${m.maintenance_date}</td><td style="padding: 2px 3px; border: 0.5px solid #ddd; text-align: center; font-weight: bold;">${m.cost.toLocaleString('en-US')}</td></tr>`;
        });

        html += `
                    </tbody>
                    <tfoot>
                        <tr style="background-color: #ecf0f1; font-weight: bold; height: 12px;">
                            <td colspan="3" style="padding: 3px 4px; border: 0.5px solid #999; text-align: right;">الإجمالي</td>
                            <td style="padding: 3px 4px; border: 0.5px solid #999; text-align: center; color: #27ae60;">${totalMaintenance.toLocaleString('en-US')}</td>
                        </tr>
                    </tfoot>
                </table>
            </div>
        `;
    }

    if (filteredViolations.length > 0) {
        html += `
            <div style="margin-top: 8px;">
                <h3 style="margin: 5px 0; font-size: 12px; color: white; background-color: #e74c3c; padding: 4px 8px; font-weight: bold;">⚠️ مصاريف المخالفات</h3>
                <table style="width: 100%; border-collapse: collapse; margin: 4px 0; font-size: 10px;">
                    <thead>
                        <tr style="background-color: #f1948b; color: white; height: 12px;">
                            <th style="padding: 3px 4px; border: 0.5px solid #999; text-align: right;">اللوحة</th>
                            <th style="padding: 3px 4px; border: 0.5px solid #999; text-align: right;">نوع المخالفة</th>
                            <th style="padding: 3px 4px; border: 0.5px solid #999; text-align: right;">التاريخ</th>
                            <th style="padding: 3px 4px; border: 0.5px solid #999; text-align: center;">المبلغ</th>
                        </tr>
                    </thead>
                    <tbody>
        `;
        
        filteredViolations.forEach((v, idx) => {
            const bgColor = idx % 2 === 0 ? '#ffffff' : '#f5f5f5';
            html += `<tr style="background-color: ${bgColor}; height: 11px;"><td style="padding: 2px 3px; border: 0.5px solid #ddd;">${v.plate_number || '-'}</td><td style="padding: 2px 3px; border: 0.5px solid #ddd; font-size: 9px;">${v.violation_type}</td><td style="padding: 2px 3px; border: 0.5px solid #ddd; font-size: 9px;">${v.violation_date}</td><td style="padding: 2px 3px; border: 0.5px solid #ddd; text-align: center; font-weight: bold;">${v.amount.toLocaleString('en-US')}</td></tr>`;
        });

        html += `
                    </tbody>
                    <tfoot>
                        <tr style="background-color: #ecf0f1; font-weight: bold; height: 12px;">
                            <td colspan="3" style="padding: 3px 4px; border: 0.5px solid #999; text-align: right;">الإجمالي</td>
                            <td style="padding: 3px 4px; border: 0.5px solid #999; text-align: center; color: #e74c3c;">${totalViolations.toLocaleString('en-US')}</td>
                        </tr>
                    </tfoot>
                </table>
            </div>
        `;
    }

    if (filteredExpenses.length > 0) {
        html += `
            <div style="margin-top: 8px;">
                <h3 style="margin: 5px 0; font-size: 12px; color: white; background-color: #f39c12; padding: 4px 8px; font-weight: bold;">💰 مصاريف النفقات</h3>
                <table style="width: 100%; border-collapse: collapse; margin: 4px 0; font-size: 10px;">
                    <thead>
                        <tr style="background-color: #f5b041; color: white; height: 12px;">
                            <th style="padding: 3px 4px; border: 0.5px solid #999; text-align: right;">اللوحة</th>
                            <th style="padding: 3px 4px; border: 0.5px solid #999; text-align: right;">نوع النفقة</th>
                            <th style="padding: 3px 4px; border: 0.5px solid #999; text-align: right;">التاريخ</th>
                            <th style="padding: 3px 4px; border: 0.5px solid #999; text-align: center;">المبلغ</th>
                        </tr>
                    </thead>
                    <tbody>
        `;
        
        filteredExpenses.forEach((e, idx) => {
            const bgColor = idx % 2 === 0 ? '#ffffff' : '#f5f5f5';
            html += `<tr style="background-color: ${bgColor}; height: 11px;"><td style="padding: 2px 3px; border: 0.5px solid #ddd;">${e.plate_number || '-'}</td><td style="padding: 2px 3px; border: 0.5px solid #ddd; font-size: 9px;">${e.expense_type}</td><td style="padding: 2px 3px; border: 0.5px solid #ddd; font-size: 9px;">${e.expense_date}</td><td style="padding: 2px 3px; border: 0.5px solid #ddd; text-align: center; font-weight: bold;">${e.amount.toLocaleString('en-US')}</td></tr>`;
        });

        html += `
                    </tbody>
                    <tfoot>
                        <tr style="background-color: #ecf0f1; font-weight: bold; height: 12px;">
                            <td colspan="3" style="padding: 3px 4px; border: 0.5px solid #999; text-align: right;">الإجمالي</td>
                            <td style="padding: 3px 4px; border: 0.5px solid #999; text-align: center; color: #f39c12;">${totalExpenses.toLocaleString('en-US')}</td>
                        </tr>
                    </tfoot>
                </table>
            </div>
        `;
    }



    html += `
        <div style="margin-top: 12px; padding: 10px; background-color: #1a252f; color: white; border-radius: 3px;">
            <table style="width: 100%; border-collapse: collapse;">
                <tr style="height: 14px;"><td style="padding: 4px; font-size: 11px; text-align: right;">إجمالي الصيانة:</td><td style="padding: 4px; font-size: 11px; font-weight: bold; text-align: center;">${totalMaintenance.toLocaleString('en-US')}</td></tr>
                <tr style="height: 14px; background-color: rgba(255,255,255,0.1);"><td style="padding: 4px; font-size: 11px; text-align: right;">إجمالي المخالفات:</td><td style="padding: 4px; font-size: 11px; font-weight: bold; text-align: center;">${totalViolations.toLocaleString('en-US')}</td></tr>
                <tr style="height: 14px;"><td style="padding: 4px; font-size: 11px; text-align: right;">إجمالي النفقات:</td><td style="padding: 4px; font-size: 11px; font-weight: bold; text-align: center;">${totalExpenses.toLocaleString('en-US')}</td></tr>
                <tr style="height: 16px; border-top: 2px solid white;"><td style="padding: 5px; font-size: 12px; font-weight: bold; text-align: right;">🔴 الإجمالي العام:</td><td style="padding: 5px; font-size: 13px; font-weight: bold; text-align: center; color: #ffd700;">${grandTotal.toLocaleString('en-US')} جنيه</td></tr>
            </table>
        </div>
    `;

    html += `</div>`;
    return html;
}

function generateFilteredExpensesReportExcel() {
    const { filteredMaintenance, filteredViolations, filteredExpenses, filteredAdvance } = getFilteredExpensesData();
    
    const data = [
        ['تقرير المصروفات (مصفاة)', '', '', ''],
        ['التاريخ: ' + new Date().toLocaleDateString('ar-EG'), '', '', ''],
        []
    ];

    if (filteredMaintenance.length > 0) {
        data.push(['مصاريف الصيانة', '', '', '']);
        data.push(['اللوحة', 'نوع الصيانة', 'التاريخ', 'المبلغ']);
        filteredMaintenance.forEach(m => {
            data.push([m.plate_number || '-', m.maintenance_type, m.maintenance_date, m.cost]);
        });
        data.push([]);
    }

    if (filteredViolations.length > 0) {
        data.push(['مصاريف المخالفات', '', '', '']);
        data.push(['اللوحة', 'نوع المخالفة', 'التاريخ', 'المبلغ']);
        filteredViolations.forEach(v => {
            data.push([v.plate_number || '-', v.violation_type, v.violation_date, v.amount]);
        });
        data.push([]);
    }

    if (filteredExpenses.length > 0) {
        data.push(['مصاريف النفقات', '', '', '']);
        data.push(['اللوحة', 'نوع النفقة', 'التاريخ', 'المبلغ']);
        filteredExpenses.forEach(e => {
            data.push([e.plate_number || '-', e.expense_type, e.expense_date, e.amount]);
        });
        data.push([]);
    }

    if (filteredAdvance.length > 0) {
        data.push(['مصاريف العهدة', '', '', '']);
        data.push(['الوصف', 'البيان', 'التاريخ', 'المبلغ']);
        filteredAdvance.forEach(a => {
            data.push([a.description || '-', a.notes || '-', a.advance_date || '-', a.amount]);
        });
    }

    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'المصروفات');
    XLSX.writeFile(wb, 'تقرير_المصروفات_مصفاة.xlsx');
    closeExpensesFilterModal();
}

const FLEET_GROUPS = {
    green: {
        label: 'سيارات البن الأخضر',
        advanceLabel: 'عهدة البن الأخضر',
        statuses: ['اخضر'],
        color: '#16814d'
    },
    roasted: {
        label: 'سيارات البن المطحون',
        advanceLabel: 'عهدة البن المطحون',
        statuses: ['مطحون', 'نقل موظفين'],
        color: '#8a5a24'
    }
};

function normalizeFleetGroup(group) {
    return group === 'green' ? 'green' : 'roasted';
}

function getFleetGroupLabel(group) {
    return FLEET_GROUPS[normalizeFleetGroup(group)].label;
}

function getAdvanceGroupLabel(group) {
    return FLEET_GROUPS[normalizeFleetGroup(group)].advanceLabel;
}

function getVehicleByRecord(record) {
    if (!record) return null;
    return appData.vehicles.find(v => v.id === record.vehicle_id) ||
        appData.vehicles.find(v => v.plate_number === record.plate_number) ||
        null;
}

function getDriverNameByRecord(record) {
    const vehicle = getVehicleByRecord(record);
    return record?.driver_name || vehicle?.vin_number || '';
}

function getVehicleFleetGroup(vehicle) {
    if (!vehicle) return 'roasted';
    return FLEET_GROUPS.green.statuses.includes(vehicle.status) ? 'green' : 'roasted';
}

function getRecordFleetGroup(record) {
    if (record?.fleet_group) return normalizeFleetGroup(record.fleet_group);
    return getVehicleFleetGroup(getVehicleByRecord(record));
}

function recordMatchesFleetGroup(record, group) {
    return getRecordFleetGroup(record) === normalizeFleetGroup(group);
}

function getSelectedReportGroup(selectId, fallback = 'roasted') {
    return normalizeFleetGroup(document.getElementById(selectId)?.value || fallback);
}

function loadData() {
    try {
        appData.vehicles = JSON.parse(localStorage.getItem('vehicles')) || [];
        appData.maintenance = JSON.parse(localStorage.getItem('maintenance')) || [];
        appData.violations = JSON.parse(localStorage.getItem('violations')) || [];
        appData.expenses = JSON.parse(localStorage.getItem('expenses')) || [];
        appData.advance = JSON.parse(localStorage.getItem('advance')) || [];
        appData.tires_in = JSON.parse(localStorage.getItem('tires_in')) || [];
        appData.tires_out = JSON.parse(localStorage.getItem('tires_out')) || [];
        appData.movements = JSON.parse(localStorage.getItem('movements')) || [];
        appData.handovers = JSON.parse(localStorage.getItem('handovers')) || [];
        appData.penalties = JSON.parse(localStorage.getItem('penalties')) || [];

        let changed = false;
        appData.expenses.forEach(expense => {
            if (!expense.fleet_group) {
                expense.fleet_group = getRecordFleetGroup(expense);
                changed = true;
            }
        });
        appData.advance.forEach(advance => {
            if (!advance.fleet_group) {
                advance.fleet_group = 'roasted';
                changed = true;
            }
        });
        if (changed) saveData();
    } catch (error) {
        console.error('خطأ في تحميل البيانات:', error);
        appData.vehicles = [];
        appData.maintenance = [];
        appData.violations = [];
        appData.expenses = [];
        appData.advance = [];
        appData.tires_in = [];
        appData.tires_out = [];
        appData.movements = [];
        appData.handovers = [];
        appData.penalties = [];
    }
}

function refreshExpensesModalOptions(selectedVehicleId = '', selectedAdvanceId = '') {
    const group = normalizeFleetGroup(document.getElementById('expensesFleetGroup')?.value);
    const vehicleSelect = document.getElementById('expensesVehicle');
    const advanceSelect = document.getElementById('expensesAdvance');
    if (!vehicleSelect || !advanceSelect) return;

    vehicleSelect.innerHTML = '<option value="">اختر مركبة (اختياري)</option>';
    appData.vehicles
        .filter(v => getVehicleFleetGroup(v) === group)
        .forEach(v => {
            const option = document.createElement('option');
            option.value = v.id;
            option.textContent = `${v.plate_number} - ${v.model}`;
            vehicleSelect.appendChild(option);
        });
    vehicleSelect.value = selectedVehicleId || '';
    updateExpensesDriverName();

    advanceSelect.innerHTML = '<option value="">لم ترتبط بعهدة</option>';
    appData.advance
        .filter(a => a.is_active && normalizeFleetGroup(a.fleet_group) === group)
        .forEach(a => {
            const option = document.createElement('option');
            option.value = a.id;
            option.textContent = `${getAdvanceGroupLabel(a.fleet_group)} - ${Number(a.amount || 0).toLocaleString('en-US')} جنيه - ${a.advance_date}`;
            advanceSelect.appendChild(option);
        });
    advanceSelect.value = selectedAdvanceId || '';
}

function updateExpensesDriverName() {
    const vehicleId = document.getElementById('expensesVehicle')?.value;
    const driverInput = document.getElementById('expensesDriverName');
    if (!driverInput) return;

    const vehicle = appData.vehicles.find(v => v.id === vehicleId);
    driverInput.value = vehicle?.vin_number || '';
}

function openExpensesModal(expenseId = null) {
    editingExpenseId = expenseId;
    const modal = document.getElementById('expensesModal');
    const form = document.getElementById('expensesForm');
    const title = document.getElementById('expensesModalTitle');

    if (expenseId) {
        title.textContent = 'تعديل بيانات النفقة';
        const expense = appData.expenses.find(e => e.id === expenseId);
        if (expense) {
            document.getElementById('expensesFleetGroup').value = getRecordFleetGroup(expense);
            refreshExpensesModalOptions(expense.vehicle_id || '', expense.advance_id || '');
            document.getElementById('expensesDriverName').value = getDriverNameByRecord(expense);
            document.getElementById('expensesType').value = expense.is_custom_type ? '' : (expense.expense_type || '');
            document.getElementById('customExpenseType').value = expense.is_custom_type ? expense.expense_type : '';
            document.getElementById('expensesDate').value = expense.expense_date;
            document.getElementById('expensesAmount').value = expense.amount;
            document.getElementById('expensesNotes').value = expense.notes || '';
        }
    } else {
        title.textContent = 'إضافة نفقة جديدة';
        form.reset();
        document.getElementById('expensesFleetGroup').value = 'roasted';
        document.getElementById('customExpenseType').value = '';
        document.getElementById('expensesDriverName').value = '';
        refreshExpensesModalOptions();
    }

    modal.classList.add('show');
}

function handleExpensesSubmit(e) {
    e.preventDefault();

    const group = normalizeFleetGroup(document.getElementById('expensesFleetGroup').value);
    const vehicleId = document.getElementById('expensesVehicle').value;
    const vehicle = vehicleId ? appData.vehicles.find(v => v.id === vehicleId) : null;

    let expenseType = document.getElementById('expensesType').value;
    let isCustom = false;
    const customType = document.getElementById('customExpenseType').value.trim();

    if (customType) {
        expenseType = customType;
        isCustom = true;
    }

    const expenseData = {
        id: editingExpenseId || Date.now().toString(),
        fleet_group: group,
        vehicle_id: vehicleId || null,
        plate_number: vehicle ? vehicle.plate_number : '',
        driver_name: vehicle ? (vehicle.vin_number || '') : '',
        expense_type: expenseType,
        is_custom_type: isCustom,
        expense_date: document.getElementById('expensesDate').value,
        amount: parseFloat(document.getElementById('expensesAmount').value),
        advance_id: document.getElementById('expensesAdvance').value || null,
        notes: document.getElementById('expensesNotes').value
    };

    try {
        if (editingExpenseId) {
            const index = appData.expenses.findIndex(item => item.id === editingExpenseId);
            if (index !== -1) appData.expenses[index] = expenseData;
        } else {
            appData.expenses.push(expenseData);
        }

        saveData();
        closeExpensesModal();
        populateExpensesList();
        populateAdvanceList();
        renderDashboard();
        alert('تم حفظ بيانات النفقة بنجاح');
    } catch (error) {
        console.error('خطأ:', error);
        alert('حدث خطأ في حفظ البيانات: ' + error.message);
    }
}

function handleAdvanceSubmit(e) {
    e.preventDefault();

    const advanceData = {
        id: Date.now().toString(),
        fleet_group: normalizeFleetGroup(document.getElementById('advanceFleetGroup').value),
        amount: parseFloat(document.getElementById('advanceAmount').value),
        advance_date: document.getElementById('advanceDate').value,
        is_active: true,
        notes: document.getElementById('advanceNotes').value
    };

    try {
        appData.advance.push(advanceData);
        saveData();
        closeAdvanceModal();
        populateAdvanceList();
        renderDashboard();
        alert('تم إضافة العهدة بنجاح');
    } catch (error) {
        console.error('خطأ:', error);
        alert('حدث خطأ في حفظ البيانات: ' + error.message);
    }
}

function getAdvanceSpent(advanceId) {
    return appData.expenses
        .filter(e => e.advance_id === advanceId)
        .reduce((sum, e) => sum + Number(e.amount || 0), 0);
}

function populateAdvanceList() {
    const tbody = document.getElementById('advanceTable');
    tbody.innerHTML = '';

    const groupFilter = document.getElementById('advanceGroupFilter')?.value || '';
    const statusFilter = document.getElementById('advanceStatusFilter')?.value || '';

    let list = [...appData.advance];
    if (groupFilter) {
        list = list.filter(a => normalizeFleetGroup(a.fleet_group) === groupFilter);
    }
    if (statusFilter) {
        list = list.filter(a => {
            const isClosed = a.is_active === false;
            return statusFilter === 'active' ? !isClosed : isClosed;
        });
    }
    // الأحدث أولاً
    list.sort((a, b) => new Date(b.advance_date) - new Date(a.advance_date));

    if (list.length === 0) {
        tbody.innerHTML = '<tr class="empty-row"><td colspan="8">لا توجد عهد مسجلة</td></tr>';
        updateAdvanceSummary();
        return;
    }

    list.forEach(advance => {
        const row = document.createElement('tr');
        const group = normalizeFleetGroup(advance.fleet_group);
        const amount = Number(advance.amount || 0);
        const spent = getAdvanceSpent(advance.id);
        const remaining = amount - spent;
        const isClosed = advance.is_active === false;

        let statusText, statusBadge;
        if (isClosed) {
            statusText = 'مقفولة';
            statusBadge = 'status-inactive';
        } else if (remaining <= 0) {
            statusText = 'مستخدمة بالكامل';
            statusBadge = 'status-complete';
        } else {
            statusText = 'نشطة';
            statusBadge = 'status-in-progress';
        }

        const closeReopenBtn = isClosed
            ? `<button class="btn btn-secondary btn-small" onclick="reopenAdvance('${advance.id}')">إعادة فتح</button>`
            : `<button class="btn btn-warning btn-small" onclick="closeAdvanceEntry('${advance.id}')">إنهاء العهدة</button>`;

        row.innerHTML = `
            <td>${amount.toLocaleString('en-US')} جنيه</td>
            <td><span class="fleet-pill ${group}">${getAdvanceGroupLabel(group)}</span></td>
            <td>${advance.advance_date}</td>
            <td>${spent.toLocaleString('en-US')} جنيه</td>
            <td>${remaining.toLocaleString('en-US')} جنيه</td>
            <td><span class="status-badge ${statusBadge}">${statusText}</span></td>
            <td>${advance.notes || '-'}</td>
            <td>
                <div class="action-buttons">
                    <button class="btn btn-info btn-small" onclick="showAdvanceDetails('${advance.id}')">تفاصيل</button>
                    ${closeReopenBtn}
                    <button class="btn btn-danger btn-small" onclick="deleteAdvance('${advance.id}')">حذف</button>
                </div>
            </td>
        `;
        tbody.appendChild(row);
    });

    updateAdvanceSummary();
}

function closeAdvanceEntry(advanceId) {
    const advance = appData.advance.find(a => a.id === advanceId);
    if (!advance) return;

    const spent = getAdvanceSpent(advanceId);
    const remaining = Number(advance.amount || 0) - spent;
    const remainingNote = remaining > 0
        ? `\nملحوظة: هيفضل متبقي ${remaining.toLocaleString('en-US')} جنيه من العهدة دي متسجلش عليه مصروفات.`
        : '';

    if (!confirm(`هل أنت متأكد من إنهاء وتقفيل هذه العهدة؟ لن تظهر بعد ذلك عند إضافة نفقات جديدة.${remainingNote}`)) {
        return;
    }

    advance.is_active = false;
    saveData();
    populateAdvanceList();
    alert('تم إنهاء العهدة بنجاح. تقدر دلوقتي تضيف عهدة جديدة.');
}

function reopenAdvance(advanceId) {
    const advance = appData.advance.find(a => a.id === advanceId);
    if (!advance) return;

    advance.is_active = true;
    saveData();
    populateAdvanceList();
}

function showAdvanceDetails(advanceId) {
    const advance = appData.advance.find(a => a.id === advanceId);
    if (!advance) return;

    const group = normalizeFleetGroup(advance.fleet_group);
    const amount = Number(advance.amount || 0);
    const relatedExpenses = appData.expenses
        .filter(e => e.advance_id === advanceId)
        .sort((a, b) => new Date(a.expense_date) - new Date(b.expense_date));
    const spent = relatedExpenses.reduce((sum, e) => sum + Number(e.amount || 0), 0);
    const remaining = amount - spent;

    document.getElementById('advanceDetailsTitle').textContent =
        `تفاصيل ${getAdvanceGroupLabel(group)} - ${amount.toLocaleString('en-US')} جنيه بتاريخ ${advance.advance_date}`;

    document.getElementById('advanceDetailsSummary').innerHTML = `
        <div class="summary-item"><span class="summary-label">مبلغ العهدة</span><span class="summary-value">${amount.toLocaleString('en-US')} جنيه</span></div>
        <div class="summary-item"><span class="summary-label">المصروف منها</span><span class="summary-value">${spent.toLocaleString('en-US')} جنيه</span></div>
        <div class="summary-item"><span class="summary-label">المتبقي</span><span class="summary-value">${remaining.toLocaleString('en-US')} جنيه</span></div>
    `;

    const tbody = document.getElementById('advanceDetailsTable');
    if (relatedExpenses.length === 0) {
        tbody.innerHTML = '<tr class="empty-row"><td colspan="5">لا توجد نفقات مرتبطة بهذه العهدة</td></tr>';
    } else {
        tbody.innerHTML = relatedExpenses.map(e => {
            const vehicle = getVehicleByRecord(e);
            const plateDriver = [vehicle?.plate_number || e.plate_number, e.driver_name].filter(Boolean).join(' - ') || '-';
            return `
                <tr>
                    <td>${e.expense_date || '-'}</td>
                    <td>${e.expense_type || '-'}</td>
                    <td>${plateDriver}</td>
                    <td>${Number(e.amount || 0).toLocaleString('en-US')} جنيه</td>
                    <td>${e.notes || '-'}</td>
                </tr>
            `;
        }).join('');
    }

    document.getElementById('advanceDetailsModal').classList.add('show');
}

function closeAdvanceDetailsModal() {
    document.getElementById('advanceDetailsModal').classList.remove('show');
}

function updateAdvanceSummary() {
    ['green', 'roasted'].forEach(group => {
        const totalAdvance = appData.advance
            .filter(a => normalizeFleetGroup(a.fleet_group) === group)
            .reduce((sum, a) => sum + Number(a.amount || 0), 0);
        const usedAdvance = appData.expenses
            .filter(e => getRecordFleetGroup(e) === group)
            .reduce((sum, e) => sum + Number(e.amount || 0), 0);
        const remainingAdvance = totalAdvance - usedAdvance;
        const suffix = group === 'green' ? 'Green' : 'Roasted';

        document.getElementById(`totalAdvance${suffix}`).textContent = totalAdvance.toLocaleString('en-US') + ' جنيه';
        document.getElementById(`usedAdvance${suffix}`).textContent = usedAdvance.toLocaleString('en-US') + ' جنيه';
        document.getElementById(`remainingAdvance${suffix}`).textContent = remainingAdvance.toLocaleString('en-US') + ' جنيه';
    });

    if (document.getElementById('totalAdvance')) {
        const total = appData.advance.reduce((sum, a) => sum + Number(a.amount || 0), 0);
        const used = appData.expenses.reduce((sum, e) => sum + Number(e.amount || 0), 0);
        document.getElementById('totalAdvance').textContent = total.toLocaleString('en-US') + ' جنيه';
        document.getElementById('usedAdvance').textContent = used.toLocaleString('en-US') + ' جنيه';
        document.getElementById('remainingAdvance').textContent = (total - used).toLocaleString('en-US') + ' جنيه';
    }
}

function filterExpenses() {
    const searchValue = document.getElementById('expensesSearch').value.toLowerCase();
    const groupFilter = document.getElementById('expensesGroupFilter')?.value || '';

    const filtered = appData.expenses.filter(expense => {
        const vehiclePlate = expense.plate_number ? expense.plate_number.toLowerCase() : '';
        const driverName = getDriverNameByRecord(expense).toLowerCase();
        const matchesSearch = vehiclePlate.includes(searchValue) ||
            driverName.includes(searchValue) ||
            (expense.expense_type ? expense.expense_type.toLowerCase().includes(searchValue) : false) ||
            getFleetGroupLabel(getRecordFleetGroup(expense)).toLowerCase().includes(searchValue);
        const matchesGroup = !groupFilter || getRecordFleetGroup(expense) === groupFilter;
        return matchesSearch && matchesGroup;
    });

    const tbody = document.getElementById('expensesTable');
    tbody.innerHTML = '';

    if (filtered.length === 0) {
        tbody.innerHTML = '<tr class="empty-row"><td colspan="8">لا توجد نفقات مسجلة</td></tr>';
        return;
    }

    filtered.forEach(expense => {
        const group = getRecordFleetGroup(expense);
        const vehiclePlate = expense.plate_number || 'عام';
        const driverName = getDriverNameByRecord(expense) || '-';
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${vehiclePlate}</td>
            <td>${driverName}</td>
            <td><span class="fleet-pill ${group}">${getFleetGroupLabel(group)}</span></td>
            <td>${expense.expense_type}</td>
            <td>${expense.expense_date}</td>
            <td>${Number(expense.amount || 0).toLocaleString('en-US')} جنيه</td>
            <td>${expense.notes || '-'}</td>
            <td>
                <div class="action-buttons">
                    <button class="btn btn-primary btn-small" onclick="openExpensesModal('${expense.id}')">تعديل</button>
                    <button class="btn btn-danger btn-small" onclick="deleteExpense('${expense.id}')">حذف</button>
                </div>
            </td>
        `;
        tbody.appendChild(row);
    });
}

function getFinancialReportData(group) {
    const fleetGroup = normalizeFleetGroup(group);
    const maintenance = appData.maintenance.filter(item => recordMatchesFleetGroup(item, fleetGroup));
    const violations = appData.violations.filter(item => item.status === 'مسددة' && recordMatchesFleetGroup(item, fleetGroup));
    const expenses = appData.expenses.filter(item => getRecordFleetGroup(item) === fleetGroup);
    const advance = appData.advance.filter(item => normalizeFleetGroup(item.fleet_group) === fleetGroup);
    return { maintenance, violations, expenses, advance };
}

function sumBy(items, key) {
    return items.reduce((sum, item) => sum + Number(item[key] || 0), 0);
}

function buildProfessionalReportHTML({ group, title, includeAdvance, data }) {
    const totalMaintenance = sumBy(data.maintenance, 'cost');
    const totalViolations = sumBy(data.violations, 'amount');
    const totalExpenses = sumBy(data.expenses, 'amount');
    const totalAdvance = sumBy(data.advance, 'amount');
    const spentTotal = includeAdvance ? totalExpenses : totalMaintenance + totalViolations + totalExpenses;
    const remainingAdvance = totalAdvance - totalExpenses;
    const color = FLEET_GROUPS[normalizeFleetGroup(group)].color;

    const renderRows = (items, type) => items.map((item, index) => {
        const plate = item.plate_number || 'عام';
        const driverName = getDriverNameByRecord(item) || '-';
        const name = type === 'maintenance' ? item.maintenance_type : type === 'violation' ? item.violation_type : item.expense_type;
        const date = type === 'maintenance' ? item.maintenance_date : type === 'violation' ? item.violation_date : item.expense_date;
        const amount = type === 'maintenance' ? item.cost : item.amount;
        return `
            <tr style="background:${index % 2 ? '#ffffff' : '#f7f9fb'};">
                <td>${plate}</td>
                <td>${driverName}</td>
                <td>${name || '-'}</td>
                <td>${date || '-'}</td>
                <td class="amount">${Number(amount || 0).toLocaleString('en-US')}</td>
                <td>${item.notes || '-'}</td>
            </tr>
        `;
    }).join('');

    const section = (heading, rows, total, emptyText) => `
        <section class="report-section">
            <div class="section-title">
                <h3>${heading}</h3>
                <strong>${Number(total || 0).toLocaleString('en-US')} جنيه</strong>
            </div>
            <table>
                <thead>
                    <tr>
                        <th>اللوحة</th>
                        <th>اسم السائق</th>
                        <th>البيان</th>
                        <th>التاريخ</th>
                        <th>المبلغ</th>
                        <th>ملاحظات</th>
                    </tr>
                </thead>
                <tbody>
                    ${rows || `<tr><td colspan="6" class="empty-report">${emptyText}</td></tr>`}
                </tbody>
                ${rows ? `
                    <tfoot>
                        <tr class="total-row">
                            <td colspan="4">الإجمالي</td>
                            <td class="amount">${Number(total || 0).toLocaleString('en-US')}</td>
                            <td>جنيه</td>
                        </tr>
                    </tfoot>
                ` : ''}
            </table>
        </section>
    `;

    return `
        <div dir="rtl" class="professional-report ${includeAdvance ? 'advance-report' : ''}">
            <style>
                @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800&display=swap');
                .professional-report { font-family: 'Cairo', sans-serif; color: #17202a; padding: 20px; background: #fff; }
                .report-hero { border: 1px solid #d9e0e7; border-top: 6px solid ${color}; border-radius: 8px; padding: 18px; position: relative; margin-bottom: 14px; }
                .report-hero img { position: absolute; left: 18px; top: 14px; width: 82px; height: 62px; object-fit: contain; }
                .report-hero h1 { margin: 0; font-size: 25px; font-weight: 800; }
                .report-hero h2 { margin: 4px 0 0; font-size: 16px; color: ${color}; font-weight: 800; }
                .report-meta { display: flex; gap: 8px; flex-wrap: wrap; margin-top: 12px; font-size: 11px; color: #53616f; }
                .report-meta span { border: 1px solid #dce3ea; border-radius: 999px; padding: 4px 10px; background: #f8fafc; }
                .kpi-grid { display: grid; grid-template-columns: repeat(${includeAdvance ? 3 : 4}, 1fr); gap: 8px; margin-bottom: 14px; }
                .kpi { border: 1px solid #dce3ea; border-radius: 8px; padding: 10px; background: #fbfcfe; }
                .kpi span { display: block; font-size: 10px; color: #5f6f80; margin-bottom: 5px; }
                .advance-report .kpi span { font-size: 12px; font-weight: 700; }
                .kpi strong { font-size: 16px; color: #17202a; }
                .report-section { margin-top: 12px; break-inside: avoid; }
                .section-title { display: flex; justify-content: space-between; align-items: center; background: ${color}; color: white; padding: 7px 10px; border-radius: 7px 7px 0 0; }
                .section-title h3 { margin: 0; font-size: 13px; }
                .section-title strong { font-size: 12px; }
                table { width: 100%; border-collapse: collapse; font-size: 11px; }
                th { background: #eef3f7; color: #263645; text-align: right; padding: 7px; border: 1px solid #d7dee6; }
                td { padding: 6px 7px; border: 1px solid #e0e5eb; }
                .amount { text-align: center; font-weight: 800; color: ${color}; }
                .total-row td { background: #f1f5f8; font-weight: 800; border-top: 2px solid ${color}; }
                .empty-report { text-align: center; color: #7f8c8d; padding: 14px; }
                .signatures { display: flex; justify-content: space-between; margin-top: 18px; gap: 18px; }
                .signature { flex: 1; text-align: center; padding-top: 28px; border-top: 1.5px solid #17202a; font-size: 12px; font-weight: 800; }
            </style>
            <div class="report-hero">
                <img src="5.jpg">
                <h1>مصنع البهنساوي</h1>
                <h2>${title}</h2>
                <div class="report-meta">
                    <span>${getFleetGroupLabel(group)}</span>
                    <span>تاريخ التقرير: ${new Date().toLocaleDateString('ar-EG')}</span>
                    <span>عدد البنود: ${data.expenses.length + data.maintenance.length + data.violations.length}</span>
                </div>
            </div>
            <div class="kpi-grid">
                ${includeAdvance ? `
                    <div class="kpi"><span>العهدة المتاحة</span><strong>${totalAdvance.toLocaleString('en-US')}</strong></div>
                    <div class="kpi"><span>المصروف من العهدة</span><strong>${totalExpenses.toLocaleString('en-US')}</strong></div>
                    <div class="kpi"><span>المتبقي</span><strong>${remainingAdvance.toLocaleString('en-US')}</strong></div>
                ` : `
                    <div class="kpi"><span>الصيانة</span><strong>${totalMaintenance.toLocaleString('en-US')}</strong></div>
                    <div class="kpi"><span>المخالفات المسددة</span><strong>${totalViolations.toLocaleString('en-US')}</strong></div>
                    <div class="kpi"><span>النفقات</span><strong>${totalExpenses.toLocaleString('en-US')}</strong></div>
                    <div class="kpi"><span>الإجمالي</span><strong>${spentTotal.toLocaleString('en-US')}</strong></div>
                `}
            </div>
            ${includeAdvance ? '' : section('مصاريف الصيانة والتصليح', renderRows(data.maintenance, 'maintenance'), totalMaintenance, 'لا توجد مصاريف صيانة')}
            ${includeAdvance ? '' : section('المخالفات المسددة', renderRows(data.violations, 'violation'), totalViolations, 'لا توجد مخالفات مسددة')}
            ${section(includeAdvance ? 'تفاصيل المصروفات المرتبطة بالعهدة' : 'النفقات الإضافية', renderRows(data.expenses, 'expense'), totalExpenses, 'لا توجد نفقات مسجلة')}
            <div class="signatures">
                <div class="signature">مدير الحركة</div>
                <div class="signature">المدير المالي</div>
            </div>
        </div>
    `;
}

function generateExpensesReportHTML(group = getSelectedReportGroup('expensesReportGroup')) {
    const data = getFinancialReportData(group);
    return buildProfessionalReportHTML({
        group,
        title: `تقرير المصروفات - ${getFleetGroupLabel(group)}`,
        includeAdvance: false,
        data
    });
}

function generateExpensesReportPDF() {
    const group = getSelectedReportGroup('expensesReportGroup');
    const element = document.createElement('div');
    element.innerHTML = generateExpensesReportHTML(group);
    html2pdf().set({
        margin: 8,
        filename: `تقرير_مصروفات_${getFleetGroupLabel(group).replace(/\s+/g, '_')}.pdf`,
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2 },
        jsPDF: { orientation: 'portrait', unit: 'mm', format: 'a4' }
    }).from(element).save();
}

function generateExpensesReportExcel() {
    const group = getSelectedReportGroup('expensesReportGroup');
    exportFinancialReportExcel(group, false, `تقرير_مصروفات_${getFleetGroupLabel(group).replace(/\s+/g, '_')}.xlsx`);
}

function generateAdvanceExpensesReportHTML(group = getSelectedReportGroup('advanceReportGroup')) {
    const data = getFinancialReportData(group);
    return buildProfessionalReportHTML({
        group,
        title: `تقرير النفقات والعهدة - ${getAdvanceGroupLabel(group)}`,
        includeAdvance: true,
        data
    });
}

function generateAdvanceExpensesReportPDF() {
    const group = getSelectedReportGroup('advanceReportGroup');
    const element = document.createElement('div');
    element.innerHTML = generateAdvanceExpensesReportHTML(group);
    html2pdf().set({
        margin: 8,
        filename: `تقرير_العهدة_${getAdvanceGroupLabel(group).replace(/\s+/g, '_')}.pdf`,
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2 },
        jsPDF: { orientation: 'portrait', unit: 'mm', format: 'a4' }
    }).from(element).save();
}

function generateAdvanceExpensesReportExcel() {
    const group = getSelectedReportGroup('advanceReportGroup');
    exportFinancialReportExcel(group, true, `تقرير_العهدة_${getAdvanceGroupLabel(group).replace(/\s+/g, '_')}.xlsx`);
}

function printAdvanceExpensesReport() {
    const group = getSelectedReportGroup('advanceReportGroup');
    const printWindow = window.open('', '', 'width=1200,height=800');
    printWindow.document.write(generateAdvanceExpensesReportHTML(group));
    printWindow.document.close();
    printWindow.print();
}

function printExpensesReport() {
    const group = getSelectedReportGroup('expensesReportGroup');
    const printWindow = window.open('', '', 'width=1200,height=800');
    printWindow.document.write(generateExpensesReportHTML(group));
    printWindow.document.close();
    printWindow.print();
}

function exportFinancialReportExcel(group, includeAdvance, fileName) {
    const data = getFinancialReportData(group);
    const rows = [
        [includeAdvance ? 'تقرير النفقات والعهدة' : 'تقرير المصروفات', getFleetGroupLabel(group)],
        ['تاريخ التقرير', new Date().toLocaleDateString('ar-EG')],
        []
    ];

    if (includeAdvance) {
        const totalAdvance = sumBy(data.advance, 'amount');
        const totalExpenses = sumBy(data.expenses, 'amount');
        rows.push(['العهدة المتاحة', totalAdvance]);
        rows.push(['المصروف من العهدة', totalExpenses]);
        rows.push(['المتبقي', totalAdvance - totalExpenses]);
        rows.push([]);
    }

    const pushSection = (title, items, type) => {
        rows.push([title]);
        rows.push(['اللوحة', 'اسم السائق', 'البيان', 'التاريخ', 'المبلغ', 'ملاحظات']);
        items.forEach(item => {
            rows.push([
                item.plate_number || 'عام',
                getDriverNameByRecord(item) || '',
                type === 'maintenance' ? item.maintenance_type : type === 'violation' ? item.violation_type : item.expense_type,
                type === 'maintenance' ? item.maintenance_date : type === 'violation' ? item.violation_date : item.expense_date,
                type === 'maintenance' ? item.cost : item.amount,
                item.notes || ''
            ]);
        });
        rows.push([]);
    };

    if (!includeAdvance) {
        pushSection('الصيانة', data.maintenance, 'maintenance');
        pushSection('المخالفات المسددة', data.violations, 'violation');
    }
    pushSection('النفقات', data.expenses, 'expense');

    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, includeAdvance ? 'العهدة' : 'المصروفات');
    XLSX.writeFile(wb, fileName);
}

function openExpensesFilterModal() {
    const reportGroup = document.getElementById('expensesReportGroup')?.value || 'roasted';
    const filterGroup = document.getElementById('expensesFilterGroup');
    if (filterGroup) filterGroup.value = reportGroup;
    document.getElementById('expensesFilterModal').classList.add('show');
}

function getFilteredExpensesData() {
    const group = normalizeFleetGroup(document.getElementById('expensesFilterGroup')?.value || 'roasted');
    const fromDate = document.getElementById('expensesFromDate').value;
    const toDate = document.getElementById('expensesToDate').value;
    const fromAmount = parseFloat(document.getElementById('expensesFromAmount').value) || 0;
    const toAmount = parseFloat(document.getElementById('expensesToAmount').value) || Infinity;

    const data = getFinancialReportData(group);
    const byDate = (items, key) => items.filter(item => {
        const dateValue = item[key];
        if (fromDate && new Date(dateValue) < new Date(fromDate)) return false;
        if (toDate && new Date(dateValue) > new Date(toDate)) return false;
        return true;
    });
    const byAmount = (items, key) => items.filter(item => Number(item[key] || 0) >= fromAmount && Number(item[key] || 0) <= toAmount);

    return {
        group,
        filteredMaintenance: byAmount(byDate(data.maintenance, 'maintenance_date'), 'cost'),
        filteredViolations: byAmount(byDate(data.violations, 'violation_date'), 'amount'),
        filteredExpenses: byAmount(byDate(data.expenses, 'expense_date'), 'amount'),
        filteredAdvance: byAmount(byDate(data.advance, 'advance_date'), 'amount')
    };
}

function generateFilteredExpensesReportHTML() {
    const filtered = getFilteredExpensesData();
    return buildProfessionalReportHTML({
        group: filtered.group,
        title: `تقرير المصروفات المصفى - ${getFleetGroupLabel(filtered.group)}`,
        includeAdvance: false,
        data: {
            maintenance: filtered.filteredMaintenance,
            violations: filtered.filteredViolations,
            expenses: filtered.filteredExpenses,
            advance: filtered.filteredAdvance
        }
    });
}

function generateFilteredExpensesReportExcel() {
    const filtered = getFilteredExpensesData();
    const rows = [
        ['تقرير المصروفات المصفى', getFleetGroupLabel(filtered.group)],
        ['تاريخ التقرير', new Date().toLocaleDateString('ar-EG')],
        []
    ];
    [
        ['الصيانة', filtered.filteredMaintenance, 'maintenance'],
        ['المخالفات المسددة', filtered.filteredViolations, 'violation'],
        ['النفقات', filtered.filteredExpenses, 'expense']
    ].forEach(([title, items, type]) => {
        rows.push([title]);
        rows.push(['اللوحة', 'اسم السائق', 'البيان', 'التاريخ', 'المبلغ', 'ملاحظات']);
        items.forEach(item => rows.push([
            item.plate_number || 'عام',
            getDriverNameByRecord(item) || '',
            type === 'maintenance' ? item.maintenance_type : type === 'violation' ? item.violation_type : item.expense_type,
            type === 'maintenance' ? item.maintenance_date : type === 'violation' ? item.violation_date : item.expense_date,
            type === 'maintenance' ? item.cost : item.amount,
            item.notes || ''
        ]));
        rows.push([]);
    });
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'المصروفات');
    XLSX.writeFile(wb, `تقرير_مصروفات_مصفى_${getFleetGroupLabel(filtered.group).replace(/\s+/g, '_')}.xlsx`);
    closeExpensesFilterModal();
}

/* =====================================================================
   مساعدات عامة للميزات الجديدة (مخزن الكوتش + تحركات السيارات)
   ===================================================================== */

function escapeHtml(value) {
    return String(value ?? '').replace(/[&<>"']/g, ch => ({
        '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
    }[ch]));
}

// مقاسات الكوتش (295/80 R22.5) بتتعكس جوه النص العربي؛ bdi بيثبت اتجاهها
function bdi(value) {
    return `<bdi>${escapeHtml(value)}</bdi>`;
}

function todayStr() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

function getLicenseStart(vehicle) {
    if (!vehicle) return '';
    if (vehicle.license_start) return vehicle.license_start;
    if (!vehicle.license_expiry) return '';
    // مركبات مسجلة قبل إضافة الحقل: نفترض الرخصة سنة
    const d = new Date(vehicle.license_expiry);
    if (isNaN(d)) return '';
    d.setFullYear(d.getFullYear() - 1);
    return d.toISOString().split('T')[0];
}

function getRecordPlate(record) {
    const vehicle = getVehicleByRecord(record);
    return (vehicle && vehicle.plate_number) || (record && record.plate_number) || '';
}

function vehicleOptionLabel(vehicle) {
    return `${vehicle.plate_number} - ${vehicle.model}`;
}

function fillVehicleSelect(selectId, firstLabel, keepValue = true) {
    const select = document.getElementById(selectId);
    if (!select) return;
    const previous = keepValue ? select.value : '';
    select.innerHTML = `<option value="">${firstLabel}</option>`;
    appData.vehicles.forEach(v => {
        const option = document.createElement('option');
        option.value = v.id;
        option.textContent = vehicleOptionLabel(v);
        select.appendChild(option);
    });
    if (previous && appData.vehicles.some(v => v.id === previous)) {
        select.value = previous;
    }
}

function dateInRange(dateStr, from, to) {
    if (!dateStr) return !from && !to;
    if (from && dateStr < from) return false;
    if (to && dateStr > to) return false;
    return true;
}

function compareByDateDesc(a, b) {
    if (a.date !== b.date) return (b.date || '').localeCompare(a.date || '');
    return (Number(b.id) || 0) - (Number(a.id) || 0);
}

function periodLabel(from, to) {
    if (from && to) return `الفترة: من ${from} إلى ${to}`;
    if (from) return `الفترة: من ${from}`;
    if (to) return `الفترة: حتى ${to}`;
    return 'كل الفترات';
}

function openHtmlAsPdf(html, filename, orientation = 'landscape') {
    const element = document.createElement('div');
    element.innerHTML = html;
    html2pdf().set({
        margin: 10,
        filename,
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2 },
        jsPDF: { orientation, unit: 'mm', format: 'a4' }
    }).from(element).save();
}

function printHtmlReport(html) {
    const printWindow = window.open('', '', 'width=1200,height=800');
    printWindow.document.write(html);
    printWindow.document.close();
    printWindow.print();
}

const RPT_TH = 'padding: 10px; text-align: right; border: 1px solid #ddd;';
const RPT_TD = 'padding: 8px; border: 1px solid #ddd;';

function reportTable(headers, rows, footerHtml = '') {
    const head = headers.map(h => `<th style="${RPT_TH}">${h}</th>`).join('');
    const body = rows.map((cells, i) =>
        `<tr style="background-color: ${i % 2 === 0 ? '#f9f9f9' : 'white'};">${cells.map(c => `<td style="${RPT_TD}">${c}</td>`).join('')}</tr>`
    ).join('');
    return `
        <table style="width: 100%; border-collapse: collapse; margin: 10px 0 20px 0;">
            <thead><tr style="background-color: #2c3e50; color: white;">${head}</tr></thead>
            <tbody>${body}</tbody>
            ${footerHtml ? `<tfoot>${footerHtml}</tfoot>` : ''}
        </table>`;
}

function reportShell(title, subtitle, inner) {
    return `
        <div dir="rtl" style="font-family: 'Cairo', sans-serif; padding: 20px; position: relative;">
            <style>body, div, table, th, td, h1, h2, h3, p { font-family: 'Cairo', sans-serif !important; }</style>
            <img src="5.jpg" style="position: absolute; top: 0; left: 0; width: 80px; height: 60px; object-fit: contain;">
            <div style="text-align: center; margin-bottom: 20px; border-bottom: 2px solid #333; padding-bottom: 15px; margin-top: 20px;">
                <h1 style="margin: 0; font-size: 24px; font-weight: bold;">مصنع البهنساوي</h1>
                <h2 style="margin: 5px 0; font-size: 18px; font-weight: bold;">${title}</h2>
                <p style="margin: 5px 0; font-size: 12px; color: #666;">
                    التاريخ: ${new Date().toLocaleDateString('ar-EG')} — ${subtitle}
                </p>
            </div>
            ${inner}
            <div style="margin-top: 30px; padding-top: 15px; border-top: 1px solid #ddd; text-align: center; font-size: 12px; color: #666;">
                <p>تم إنشاء هذا التقرير بواسطة نظام إدارة المركبات - مصنع البهنساوي</p>
            </div>
        </div>`;
}

function reportSectionTitle(text) {
    return `<h3 style="margin: 18px 0 4px 0; font-size: 15px; color: #2c3e50;">${text}</h3>`;
}

function setupNewFeatureListeners() {
    setupHandoverListeners();
    document.getElementById('tireInForm').addEventListener('submit', handleTireInSubmit);
    document.getElementById('tireOutForm').addEventListener('submit', handleTireOutSubmit);
    document.getElementById('movementForm').addEventListener('submit', handleMovementSubmit);

    ['tireOutSearch'].forEach(id => document.getElementById(id).addEventListener('input', filterTireOut));
    ['tireOutVehicleFilter', 'tireOutFrom', 'tireOutTo'].forEach(id => document.getElementById(id).addEventListener('change', filterTireOut));

    document.getElementById('movementSearch').addEventListener('input', filterMovements);
    ['movementVehicleFilter', 'movementFrom', 'movementTo'].forEach(id => document.getElementById(id).addEventListener('change', filterMovements));
}

function populateReportVehicleOptions() {
    fillVehicleSelect('movementsReportVehicle', 'كل السيارات');
    fillVehicleSelect('handoversReportVehicle', 'كل السيارات');
}

/* =====================================================================
   مخزن الكوتش
   ===================================================================== */

function tireKey(type) {
    return String(type || '').trim().replace(/\s+/g, ' ').toLowerCase();
}

// لو نفس النوع اتكتب قبل كدة بحروف مختلفة، نستخدم نفس الاسم عشان الرصيد ميتفرقش
function canonicalTireName(type) {
    const clean = String(type || '').trim().replace(/\s+/g, ' ');
    const key = tireKey(clean);
    const found = appData.tires_in.find(r => tireKey(r.tire_type) === key);
    return found ? found.tire_type : clean;
}

function getTireStock() {
    const map = new Map();
    const bucket = (type) => {
        const key = tireKey(type);
        if (!map.has(key)) map.set(key, { type, added: 0, issued: 0 });
        return map.get(key);
    };
    appData.tires_in.forEach(r => { bucket(r.tire_type).added += Number(r.quantity) || 0; });
    appData.tires_out.forEach(r => { bucket(r.tire_type).issued += Number(r.quantity) || 0; });
    return [...map.values()]
        .map(x => ({ ...x, balance: x.added - x.issued }))
        .sort((a, b) => a.type.localeCompare(b.type, 'ar'));
}

function getTireBalance(type) {
    const key = tireKey(type);
    const found = getTireStock().find(x => tireKey(x.type) === key);
    return found ? found.balance : 0;
}

function populateTiresSection() {
    const stock = getTireStock();

    const stockBody = document.getElementById('tireStockTable');
    if (stock.length === 0) {
        stockBody.innerHTML = '<tr class="empty-row"><td colspan="4">لا توجد أنواع كوتش في المخزن</td></tr>';
    } else {
        stockBody.innerHTML = stock.map(x => `
            <tr>
                <td>${bdi(x.type)}</td>
                <td>${x.added}</td>
                <td>${x.issued}</td>
                <td class="${x.balance > 0 ? 'stock-ok' : 'stock-low'}">${x.balance}</td>
            </tr>`).join('');
    }

    const inBody = document.getElementById('tireInTable');
    const ins = [...appData.tires_in].sort(compareByDateDesc);
    if (ins.length === 0) {
        inBody.innerHTML = '<tr class="empty-row"><td colspan="5">لا توجد إضافات مسجلة</td></tr>';
    } else {
        inBody.innerHTML = ins.map(r => `
            <tr>
                <td>${escapeHtml(r.date)}</td>
                <td>${bdi(r.tire_type)}</td>
                <td>${Number(r.quantity) || 0}</td>
                <td>${escapeHtml(r.notes) || '-'}</td>
                <td><div class="action-buttons">
                    <button class="btn btn-danger btn-small" onclick="deleteTireIn('${r.id}')">حذف</button>
                </div></td>
            </tr>`).join('');
    }

    document.getElementById('tireTypesList').innerHTML =
        stock.map(x => `<option value="${escapeHtml(x.type)}"></option>`).join('');

    fillVehicleSelect('tireOutVehicleFilter', 'كل السيارات');
    filterTireOut();
}

function getFilteredTireOut() {
    const search = document.getElementById('tireOutSearch').value.trim().toLowerCase();
    const vehicleId = document.getElementById('tireOutVehicleFilter').value;
    const from = document.getElementById('tireOutFrom').value;
    const to = document.getElementById('tireOutTo').value;

    return appData.tires_out.filter(r => {
        if (vehicleId && r.vehicle_id !== vehicleId) return false;
        if (!dateInRange(r.date, from, to)) return false;
        if (search) {
            const hay = [r.tire_type, getRecordPlate(r), getDriverNameByRecord(r), r.notes].join(' ').toLowerCase();
            if (!hay.includes(search)) return false;
        }
        return true;
    }).sort(compareByDateDesc);
}

function filterTireOut() {
    const rows = getFilteredTireOut();
    const tbody = document.getElementById('tireOutTable');

    if (rows.length === 0) {
        tbody.innerHTML = '<tr class="empty-row"><td colspan="7">لا توجد عمليات صرف مسجلة</td></tr>';
        document.getElementById('tireOutSummary').textContent = '';
        return;
    }

    const totalQty = rows.reduce((sum, r) => sum + (Number(r.quantity) || 0), 0);
    document.getElementById('tireOutSummary').textContent =
        `عدد عمليات الصرف: ${rows.length} — إجمالي الكوتش المصروف: ${totalQty}`;

    tbody.innerHTML = rows.map(r => `
        <tr>
            <td>${escapeHtml(r.date)}</td>
            <td>${bdi(r.tire_type)}</td>
            <td>${Number(r.quantity) || 0}</td>
            <td>${escapeHtml(getRecordPlate(r)) || '-'}</td>
            <td>${escapeHtml(getDriverNameByRecord(r)) || '-'}</td>
            <td>${escapeHtml(r.notes) || '-'}</td>
            <td><div class="action-buttons">
                <button class="btn btn-danger btn-small" onclick="deleteTireOut('${r.id}')">حذف</button>
            </div></td>
        </tr>`).join('');
}

function resetTireOutFilters() {
    document.getElementById('tireOutSearch').value = '';
    document.getElementById('tireOutVehicleFilter').value = '';
    document.getElementById('tireOutFrom').value = '';
    document.getElementById('tireOutTo').value = '';
    filterTireOut();
}

function openTireInModal() {
    document.getElementById('tireInForm').reset();
    document.getElementById('tireInDate').value = todayStr();
    document.getElementById('tireTypesList').innerHTML =
        getTireStock().map(x => `<option value="${escapeHtml(x.type)}"></option>`).join('');
    document.getElementById('tireInModal').classList.add('show');
}

function closeTireInModal() {
    document.getElementById('tireInModal').classList.remove('show');
    document.getElementById('tireInForm').reset();
}

function handleTireInSubmit(e) {
    e.preventDefault();

    const quantity = parseInt(document.getElementById('tireInQty').value, 10);
    const type = canonicalTireName(document.getElementById('tireInType').value);
    if (!type || !(quantity > 0)) {
        alert('اكتب نوع الكوتش وعدد صحيح أكبر من صفر');
        return;
    }

    appData.tires_in.push({
        id: Date.now().toString(),
        tire_type: type,
        quantity,
        date: document.getElementById('tireInDate').value,
        notes: document.getElementById('tireInNotes').value.trim()
    });

    saveData();
    closeTireInModal();
    populateTiresSection();
    alert('تم إضافة الكوتش للمخزن بنجاح');
}

function deleteTireIn(id) {
    const record = appData.tires_in.find(r => r.id === id);
    if (!record) return;

    const balanceAfter = getTireBalance(record.tire_type) - (Number(record.quantity) || 0);
    if (balanceAfter < 0) {
        alert('مينفعش تحذف الإضافة دي: فيه كوتش من النوع ده اتصرف بالفعل، وحذفها هيخلي رصيد المخزن بالسالب.');
        return;
    }
    if (!confirm('هل أنت متأكد من حذف هذه الإضافة؟')) return;

    appData.tires_in = appData.tires_in.filter(r => r.id !== id);
    saveData();
    populateTiresSection();
}

function openTireOutModal() {
    const stock = getTireStock().filter(x => x.balance > 0);
    if (stock.length === 0) {
        alert('مفيش كوتش متاح في المخزن. أضف كوتش الأول.');
        return;
    }
    if (appData.vehicles.length === 0) {
        alert('لازم تضيف مركبة الأول عشان تصرف لها كوتش.');
        return;
    }

    document.getElementById('tireOutForm').reset();

    const typeSelect = document.getElementById('tireOutType');
    typeSelect.innerHTML = '<option value="">اختر النوع</option>' +
        stock.map(x => `<option value="${escapeHtml(x.type)}">${escapeHtml(x.type)} (المتاح: ${x.balance})</option>`).join('');

    fillVehicleSelect('tireOutVehicle', 'اختر مركبة', false);
    document.getElementById('tireOutDriver').value = '';
    document.getElementById('tireOutDate').value = todayStr();
    document.getElementById('tireOutAvailable').textContent = '';
    document.getElementById('tireOutModal').classList.add('show');
}

function closeTireOutModal() {
    document.getElementById('tireOutModal').classList.remove('show');
    document.getElementById('tireOutForm').reset();
}

function updateTireOutAvailable() {
    const type = document.getElementById('tireOutType').value;
    const hint = document.getElementById('tireOutAvailable');
    const qty = document.getElementById('tireOutQty');
    if (!type) {
        hint.textContent = '';
        qty.removeAttribute('max');
        return;
    }
    const balance = getTireBalance(type);
    hint.textContent = `المتاح في المخزن: ${balance}`;
    hint.className = 'form-hint stock-hint';
    qty.max = balance;
}

function updateTireOutDriver() {
    const vehicle = appData.vehicles.find(v => v.id === document.getElementById('tireOutVehicle').value);
    document.getElementById('tireOutDriver').value = vehicle ? (vehicle.vin_number || '') : '';
}

function handleTireOutSubmit(e) {
    e.preventDefault();

    const type = document.getElementById('tireOutType').value;
    const quantity = parseInt(document.getElementById('tireOutQty').value, 10);
    const vehicle = appData.vehicles.find(v => v.id === document.getElementById('tireOutVehicle').value);

    if (!type || !vehicle || !(quantity > 0)) {
        alert('اختار نوع الكوتش والعربية وعدد صحيح أكبر من صفر');
        return;
    }

    const balance = getTireBalance(type);
    if (quantity > balance) {
        alert(`العدد المطلوب (${quantity}) أكبر من المتاح في المخزن (${balance}).`);
        return;
    }

    appData.tires_out.push({
        id: Date.now().toString(),
        tire_type: type,
        quantity,
        vehicle_id: vehicle.id,
        plate_number: vehicle.plate_number,
        driver_name: vehicle.vin_number || '',
        date: document.getElementById('tireOutDate').value,
        notes: document.getElementById('tireOutNotes').value.trim()
    });

    saveData();
    closeTireOutModal();
    populateTiresSection();
    alert('تم تسجيل الصرف بنجاح');
}

function deleteTireOut(id) {
    if (!confirm('هل أنت متأكد من حذف عملية الصرف؟ الكمية هترجع للمخزن.')) return;
    appData.tires_out = appData.tires_out.filter(r => r.id !== id);
    saveData();
    populateTiresSection();
}

/* ---------- تقرير مخزن الكوتش ---------- */

function getTiresReportData() {
    const from = document.getElementById('tiresReportFrom').value;
    const to = document.getElementById('tiresReportTo').value;
    const ins = appData.tires_in.filter(r => dateInRange(r.date, from, to)).sort(compareByDateDesc);
    const outs = appData.tires_out.filter(r => dateInRange(r.date, from, to)).sort(compareByDateDesc);

    // ملخص الصرف لكل عربية
    const perVehicle = new Map();
    outs.forEach(r => {
        const key = r.vehicle_id || getRecordPlate(r);
        if (!perVehicle.has(key)) {
            perVehicle.set(key, { plate: getRecordPlate(r), driver: getDriverNameByRecord(r), total: 0, byType: {} });
        }
        const entry = perVehicle.get(key);
        const q = Number(r.quantity) || 0;
        entry.total += q;
        entry.byType[r.tire_type] = (entry.byType[r.tire_type] || 0) + q;
    });

    return { from, to, ins, outs, stock: getTireStock(), perVehicle: [...perVehicle.values()] };
}

function describeByType(byType) {
    return Object.entries(byType).map(([type, q]) => `${type}: ${q}`).join(' | ');
}

function describeByTypeHtml(byType) {
    return Object.entries(byType).map(([type, q]) => `${bdi(type)}: ${q}`).join(' | ');
}

function generateTiresReportHTML() {
    const { from, to, ins, outs, stock, perVehicle } = getTiresReportData();
    const totalIn = ins.reduce((s, r) => s + (Number(r.quantity) || 0), 0);
    const totalOut = outs.reduce((s, r) => s + (Number(r.quantity) || 0), 0);
    const empty = (text) => `<p style="color: #888; margin: 6px 0 16px 0;">${text}</p>`;

    let inner = reportSectionTitle('رصيد المخزن الحالي');
    inner += stock.length
        ? reportTable(['نوع الكوتش', 'إجمالي المضاف', 'إجمالي المصروف', 'المتبقي'],
            stock.map(x => [bdi(x.type), x.added, x.issued, `<strong>${x.balance}</strong>`]))
        : empty('لا توجد أنواع كوتش في المخزن');

    inner += reportSectionTitle(`الإضافات للمخزن (${periodLabel(from, to)})`);
    inner += ins.length
        ? reportTable(['التاريخ', 'نوع الكوتش', 'العدد', 'الملاحظات'],
            ins.map(r => [escapeHtml(r.date), bdi(r.tire_type), Number(r.quantity) || 0, escapeHtml(r.notes) || '-']),
            `<tr style="background-color: #ecf0f1; font-weight: bold;"><td colspan="2" style="${RPT_TD}">إجمالي المضاف</td><td style="${RPT_TD}">${totalIn}</td><td style="${RPT_TD}"></td></tr>`)
        : empty('لا توجد إضافات في هذه الفترة');

    inner += reportSectionTitle(`المصروف للسيارات (${periodLabel(from, to)})`);
    inner += outs.length
        ? reportTable(['التاريخ', 'نوع الكوتش', 'العدد', 'اتصرف لعربية', 'السائق', 'الملاحظات'],
            outs.map(r => [escapeHtml(r.date), bdi(r.tire_type), Number(r.quantity) || 0,
                escapeHtml(getRecordPlate(r)) || '-', escapeHtml(getDriverNameByRecord(r)) || '-', escapeHtml(r.notes) || '-']),
            `<tr style="background-color: #ecf0f1; font-weight: bold;"><td colspan="2" style="${RPT_TD}">إجمالي المصروف</td><td style="${RPT_TD}">${totalOut}</td><td colspan="3" style="${RPT_TD}"></td></tr>`)
        : empty('لا توجد عمليات صرف في هذه الفترة');

    if (perVehicle.length) {
        inner += reportSectionTitle('ملخص المصروف لكل عربية');
        inner += reportTable(['اللوحة', 'السائق', 'التفاصيل', 'إجمالي العدد'],
            perVehicle.map(v => [escapeHtml(v.plate) || '-', escapeHtml(v.driver) || '-', describeByTypeHtml(v.byType), `<strong>${v.total}</strong>`]));
    }

    return reportShell('تقرير مخزن الكوتش', periodLabel(from, to), inner);
}

function generateTiresReportPDF() { openHtmlAsPdf(generateTiresReportHTML(), 'تقرير_مخزن_الكوتش.pdf', 'landscape'); }
function printTiresReport() { printHtmlReport(generateTiresReportHTML()); }

function generateTiresReportExcel() {
    const { from, to, ins, outs, stock, perVehicle } = getTiresReportData();
    const wb = XLSX.utils.book_new();
    const add = (name, rows) => XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), name);

    add('رصيد المخزن', [
        ['تقرير مخزن الكوتش - رصيد المخزن الحالي'], ['التاريخ: ' + new Date().toLocaleDateString('ar-EG')], [],
        ['نوع الكوتش', 'إجمالي المضاف', 'إجمالي المصروف', 'المتبقي'],
        ...stock.map(x => [x.type, x.added, x.issued, x.balance])
    ]);
    add('الإضافات', [
        ['الإضافات للمخزن - ' + periodLabel(from, to)], [],
        ['التاريخ', 'نوع الكوتش', 'العدد', 'الملاحظات'],
        ...ins.map(r => [r.date, r.tire_type, Number(r.quantity) || 0, r.notes || '']),
        [], ['', 'إجمالي المضاف', ins.reduce((s, r) => s + (Number(r.quantity) || 0), 0)]
    ]);
    add('الصرف', [
        ['المصروف للسيارات - ' + periodLabel(from, to)], [],
        ['التاريخ', 'نوع الكوتش', 'العدد', 'اتصرف لعربية', 'السائق', 'الملاحظات'],
        ...outs.map(r => [r.date, r.tire_type, Number(r.quantity) || 0, getRecordPlate(r), getDriverNameByRecord(r), r.notes || '']),
        [], ['', 'إجمالي المصروف', outs.reduce((s, r) => s + (Number(r.quantity) || 0), 0)]
    ]);
    add('ملخص لكل عربية', [
        ['ملخص المصروف لكل عربية - ' + periodLabel(from, to)], [],
        ['اللوحة', 'السائق', 'التفاصيل', 'إجمالي العدد'],
        ...perVehicle.map(v => [v.plate, v.driver, describeByType(v.byType), v.total])
    ]);

    XLSX.writeFile(wb, 'تقرير_مخزن_الكوتش.xlsx');
}

/* =====================================================================
   تحركات السيارات
   ===================================================================== */

let editingMovementId = null;

function getMovementMinutes(m) {
    if (m.run_total_minutes !== undefined && m.run_total_minutes !== null) return Number(m.run_total_minutes) || 0;
    return (Number(m.run_hours) || 0) * 60 + (Number(m.run_minutes) || 0);
}

function formatDuration(totalMinutes) {
    const total = Math.round(Number(totalMinutes) || 0);
    const h = Math.floor(total / 60);
    const m = total % 60;
    if (!h && !m) return '0 دقيقة';
    const parts = [];
    if (h) parts.push(`${h} ساعة`);
    if (m) parts.push(`${m} دقيقة`);
    return parts.join(' و ');
}

function formatKm(value) {
    return (Number(value) || 0).toLocaleString('en-US', { maximumFractionDigits: 1 });
}

function getSortedMovements(list = appData.movements) {
    return [...list].sort(compareByDateDesc);
}

function movementBelongsToVehicle(m, vehicle) {
    return m.vehicle_id === vehicle.id || (!m.vehicle_id && m.plate_number === vehicle.plate_number);
}

function populateMovementsSection() {
    fillVehicleSelect('movementVehicleFilter', 'كل السيارات');

    const places = [...new Set(appData.movements.map(m => (m.parking_place || '').trim()).filter(Boolean))];
    document.getElementById('parkingPlacesList').innerHTML =
        places.map(p => `<option value="${escapeHtml(p)}"></option>`).join('');

    renderLastParking();
    filterMovements();
}

function renderLastParking() {
    const tbody = document.getElementById('lastParkingTable');
    if (appData.vehicles.length === 0) {
        tbody.innerHTML = '<tr class="empty-row"><td colspan="5">لا توجد مركبات</td></tr>';
        return;
    }

    const sorted = getSortedMovements();
    tbody.innerHTML = appData.vehicles.map(v => {
        const last = sorted.find(m => movementBelongsToVehicle(m, v));
        return `
            <tr>
                <td>${escapeHtml(v.plate_number)}</td>
                <td>${escapeHtml(v.vin_number) || '-'}</td>
                <td>${last ? escapeHtml(last.parking_place) : '-'}</td>
                <td>${last ? escapeHtml(last.date) : 'لا توجد تحركات'}</td>
                <td>${last ? `<button class="btn btn-info btn-small" onclick="showParkingHistory('${v.id}')">عرض السجل</button>` : '-'}</td>
            </tr>`;
    }).join('');
}

function getFilteredMovements() {
    const search = document.getElementById('movementSearch').value.trim().toLowerCase();
    const vehicleId = document.getElementById('movementVehicleFilter').value;
    const from = document.getElementById('movementFrom').value;
    const to = document.getElementById('movementTo').value;

    return getSortedMovements().filter(m => {
        if (vehicleId && m.vehicle_id !== vehicleId) return false;
        if (!dateInRange(m.date, from, to)) return false;
        if (search) {
            const hay = [getRecordPlate(m), getDriverNameByRecord(m), m.parking_place, m.notes].join(' ').toLowerCase();
            if (!hay.includes(search)) return false;
        }
        return true;
    });
}

function filterMovements() {
    const rows = getFilteredMovements();
    const tbody = document.getElementById('movementsTable');

    if (rows.length === 0) {
        tbody.innerHTML = '<tr class="empty-row"><td colspan="8">لا توجد تحركات مسجلة</td></tr>';
        document.getElementById('movementSummary').textContent = '';
        return;
    }

    const totalKm = rows.reduce((s, m) => s + (Number(m.km) || 0), 0);
    const totalMin = rows.reduce((s, m) => s + getMovementMinutes(m), 0);
    document.getElementById('movementSummary').textContent =
        `عدد التحركات: ${rows.length} — إجمالي الكيلومترات: ${formatKm(totalKm)} كم — إجمالي مدة التشغيل: ${formatDuration(totalMin)}`;

    tbody.innerHTML = rows.map(m => `
        <tr>
            <td>${escapeHtml(m.date)}</td>
            <td>${escapeHtml(getRecordPlate(m)) || '-'}</td>
            <td>${escapeHtml(getDriverNameByRecord(m)) || '-'}</td>
            <td>${formatKm(m.km)} كم</td>
            <td>${formatDuration(getMovementMinutes(m))}</td>
            <td>${escapeHtml(m.parking_place) || '-'}</td>
            <td>${escapeHtml(m.notes) || '-'}</td>
            <td><div class="action-buttons">
                <button class="btn btn-primary btn-small" onclick="openMovementModal('${m.id}')">تعديل</button>
                <button class="btn btn-danger btn-small" onclick="deleteMovement('${m.id}')">حذف</button>
            </div></td>
        </tr>`).join('');
}

function resetMovementFilters() {
    document.getElementById('movementSearch').value = '';
    document.getElementById('movementVehicleFilter').value = '';
    document.getElementById('movementFrom').value = '';
    document.getElementById('movementTo').value = '';
    filterMovements();
}

function openMovementModal(movementId = null) {
    if (appData.vehicles.length === 0) {
        alert('لازم تضيف مركبة الأول عشان تسجل لها تحركات.');
        return;
    }

    editingMovementId = movementId;
    document.getElementById('movementForm').reset();
    fillVehicleSelect('movementVehicle', 'اختر مركبة', false);

    const places = [...new Set(appData.movements.map(m => (m.parking_place || '').trim()).filter(Boolean))];
    document.getElementById('parkingPlacesList').innerHTML =
        places.map(p => `<option value="${escapeHtml(p)}"></option>`).join('');

    if (movementId) {
        const m = appData.movements.find(x => x.id === movementId);
        if (!m) return;
        document.getElementById('movementModalTitle').textContent = 'تعديل بيانات التحرك';
        document.getElementById('movementDate').value = m.date || '';
        document.getElementById('movementVehicle').value = m.vehicle_id || '';
        document.getElementById('movementDriver').value = getDriverNameByRecord(m);
        document.getElementById('movementKm').value = m.km;
        const total = getMovementMinutes(m);
        document.getElementById('movementHours').value = Math.floor(total / 60);
        document.getElementById('movementMinutes').value = total % 60;
        document.getElementById('movementParking').value = m.parking_place || '';
        document.getElementById('movementNotes').value = m.notes || '';
    } else {
        document.getElementById('movementModalTitle').textContent = 'تسجيل تحرك جديد';
        document.getElementById('movementDate').value = todayStr();
    }

    document.getElementById('movementModal').classList.add('show');
}

function closeMovementModal() {
    document.getElementById('movementModal').classList.remove('show');
    document.getElementById('movementForm').reset();
    editingMovementId = null;
}

function updateMovementDriver() {
    const vehicle = appData.vehicles.find(v => v.id === document.getElementById('movementVehicle').value);
    document.getElementById('movementDriver').value = vehicle ? (vehicle.vin_number || '') : '';
}

function handleMovementSubmit(e) {
    e.preventDefault();

    const vehicle = appData.vehicles.find(v => v.id === document.getElementById('movementVehicle').value);
    if (!vehicle) {
        alert('اختار المركبة');
        return;
    }

    const hours = Math.max(0, parseInt(document.getElementById('movementHours').value, 10) || 0);
    const minutes = Math.max(0, parseInt(document.getElementById('movementMinutes').value, 10) || 0);
    const km = parseFloat(document.getElementById('movementKm').value);
    if (!(km >= 0)) {
        alert('اكتب عدد الكيلومترات');
        return;
    }

    const record = {
        id: editingMovementId || Date.now().toString(),
        date: document.getElementById('movementDate').value,
        vehicle_id: vehicle.id,
        plate_number: vehicle.plate_number,
        driver_name: document.getElementById('movementDriver').value.trim() || vehicle.vin_number || '',
        km,
        run_hours: hours,
        run_minutes: minutes,
        run_total_minutes: hours * 60 + minutes,
        parking_place: document.getElementById('movementParking').value.trim(),
        notes: document.getElementById('movementNotes').value.trim()
    };

    if (editingMovementId) {
        const index = appData.movements.findIndex(m => m.id === editingMovementId);
        if (index !== -1) appData.movements[index] = record;
    } else {
        appData.movements.push(record);
    }

    saveData();
    closeMovementModal();
    populateMovementsSection();
    alert('تم حفظ التحرك بنجاح');
}

function deleteMovement(id) {
    if (!confirm('هل أنت متأكد من حذف هذا التحرك؟')) return;
    appData.movements = appData.movements.filter(m => m.id !== id);
    saveData();
    populateMovementsSection();
}

function showParkingHistory(vehicleId) {
    const vehicle = appData.vehicles.find(v => v.id === vehicleId);
    if (!vehicle) return;

    const rows = getSortedMovements().filter(m => movementBelongsToVehicle(m, vehicle));
    document.getElementById('parkingHistoryTitle').textContent =
        `سجل الركن - ${vehicle.plate_number}${vehicle.vin_number ? ' (' + vehicle.vin_number + ')' : ''}`;

    const tbody = document.getElementById('parkingHistoryTable');
    tbody.innerHTML = rows.length === 0
        ? '<tr class="empty-row"><td colspan="4">لا توجد بيانات</td></tr>'
        : rows.map(m => `
            <tr>
                <td>${escapeHtml(m.date)}</td>
                <td>${escapeHtml(getDriverNameByRecord(m)) || '-'}</td>
                <td>${escapeHtml(m.parking_place) || '-'}</td>
                <td>${formatKm(m.km)} كم</td>
            </tr>`).join('');

    document.getElementById('parkingHistoryModal').classList.add('show');
}

function closeParkingHistoryModal() {
    document.getElementById('parkingHistoryModal').classList.remove('show');
}

/* ---------- تقرير تحركات السيارات ---------- */

function getMovementsReportData() {
    const vehicleId = document.getElementById('movementsReportVehicle').value;
    const from = document.getElementById('movementsReportFrom').value;
    const to = document.getElementById('movementsReportTo').value;
    const vehicle = vehicleId ? appData.vehicles.find(v => v.id === vehicleId) : null;

    const rows = getSortedMovements().filter(m =>
        (!vehicleId || m.vehicle_id === vehicleId) && dateInRange(m.date, from, to));

    // ملخص لكل عربية (آخر مكان ركن داخل الفترة المحددة)
    const perVehicle = new Map();
    rows.forEach(m => {   // rows مرتبة من الأحدث للأقدم، فأول تحرك نقابله هو آخر ركنة
        const key = m.vehicle_id || getRecordPlate(m);
        if (!perVehicle.has(key)) {
            perVehicle.set(key, { plate: getRecordPlate(m), driver: getDriverNameByRecord(m), trips: 0, km: 0, minutes: 0, lastPlace: m.parking_place || '', lastDate: m.date });
        }
        const entry = perVehicle.get(key);
        entry.trips += 1;
        entry.km += Number(m.km) || 0;
        entry.minutes += getMovementMinutes(m);
    });

    return { vehicle, from, to, rows, perVehicle: [...perVehicle.values()] };
}

function movementsReportSubtitle(vehicle, from, to) {
    return `${vehicle ? 'المركبة: ' + escapeHtml(vehicle.plate_number) + ' — ' : ''}${periodLabel(from, to)}`;
}

function generateMovementsReportHTML() {
    const { vehicle, from, to, rows, perVehicle } = getMovementsReportData();
    const totalKm = rows.reduce((s, m) => s + (Number(m.km) || 0), 0);
    const totalMin = rows.reduce((s, m) => s + getMovementMinutes(m), 0);

    let inner = reportSectionTitle('سجل التحركات');
    inner += rows.length
        ? reportTable(['التاريخ', 'اللوحة', 'اسم السائق', 'الكيلومترات', 'مدة التشغيل', 'مكان الركن', 'الملاحظات'],
            rows.map(m => [escapeHtml(m.date), escapeHtml(getRecordPlate(m)) || '-', escapeHtml(getDriverNameByRecord(m)) || '-',
                `${formatKm(m.km)} كم`, formatDuration(getMovementMinutes(m)), escapeHtml(m.parking_place) || '-', escapeHtml(m.notes) || '-']),
            `<tr style="background-color: #ecf0f1; font-weight: bold;"><td colspan="3" style="${RPT_TD}">الإجمالي (${rows.length} تحرك)</td><td style="${RPT_TD}">${formatKm(totalKm)} كم</td><td style="${RPT_TD}">${formatDuration(totalMin)}</td><td colspan="2" style="${RPT_TD}"></td></tr>`)
        : '<p style="color: #888;">لا توجد تحركات في هذه الفترة</p>';

    if (perVehicle.length) {
        inner += reportSectionTitle('ملخص لكل سيارة');
        inner += reportTable(['اللوحة', 'السائق', 'عدد التحركات', 'إجمالي الكيلومترات', 'إجمالي مدة التشغيل', 'آخر مكان ركن', 'بتاريخ'],
            perVehicle.map(v => [escapeHtml(v.plate) || '-', escapeHtml(v.driver) || '-', v.trips, `${formatKm(v.km)} كم`,
                formatDuration(v.minutes), escapeHtml(v.lastPlace) || '-', escapeHtml(v.lastDate)]));
    }

    return reportShell('تقرير تحركات السيارات', movementsReportSubtitle(vehicle, from, to), inner);
}

function generateMovementsReportPDF() { openHtmlAsPdf(generateMovementsReportHTML(), 'تقرير_تحركات_السيارات.pdf', 'landscape'); }
function printMovementsReport() { printHtmlReport(generateMovementsReportHTML()); }

function generateMovementsReportExcel() {
    const { vehicle, from, to, rows, perVehicle } = getMovementsReportData();
    const totalKm = rows.reduce((s, m) => s + (Number(m.km) || 0), 0);
    const totalMin = rows.reduce((s, m) => s + getMovementMinutes(m), 0);
    const wb = XLSX.utils.book_new();
    const plainSubtitle = `${vehicle ? 'المركبة: ' + vehicle.plate_number + ' — ' : ''}${periodLabel(from, to)}`;

    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
        ['تقرير تحركات السيارات'], [plainSubtitle], [],
        ['التاريخ', 'اللوحة', 'اسم السائق', 'الكيلومترات', 'مدة التشغيل', 'مكان الركن', 'الملاحظات'],
        ...rows.map(m => [m.date, getRecordPlate(m), getDriverNameByRecord(m), Number(m.km) || 0, formatDuration(getMovementMinutes(m)), m.parking_place || '', m.notes || '']),
        [], ['الإجمالي', '', '', totalKm, formatDuration(totalMin)]
    ]), 'التحركات');

    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
        ['ملخص لكل سيارة'], [plainSubtitle], [],
        ['اللوحة', 'السائق', 'عدد التحركات', 'إجمالي الكيلومترات', 'إجمالي مدة التشغيل', 'آخر مكان ركن', 'بتاريخ'],
        ...perVehicle.map(v => [v.plate, v.driver, v.trips, v.km, formatDuration(v.minutes), v.lastPlace, v.lastDate])
    ]), 'ملخص السيارات');

    XLSX.writeFile(wb, 'تقرير_تحركات_السيارات.xlsx');
}


document.addEventListener('DOMContentLoaded', () => {
    // لو المزامنة السحابية موجودة، التطبيق مايشتغلش غير بعد تسجيل الدخول
    if (window.CloudSync) window.CloudSync.boot(initializeApp);
    else initializeApp();
});

