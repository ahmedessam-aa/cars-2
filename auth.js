/* تسجيل الدخول + مزامنة البيانات بين الأجهزة (Firebase Auth + Firestore) */
(function () {
    const KEYS = ['vehicles', 'maintenance', 'violations', 'expenses', 'advance',
        'tires_in', 'tires_out', 'movements', 'handovers', 'penalties', 'licenseWarningDays'];
    const MAX_DOC_CHARS = 900000; // حد Firestore للمستند الواحد ~1MB

    const $ = id => document.getElementById(id);
    const lastSynced = {};      // آخر قيمة معروفة في السحابة لكل مفتاح
    let db = null, auth = null, col = null, unsub = null;
    let startApp = null, appStarted = false, pushTimer = null, applyingRemote = false;

    function setStatus(state) {
        const el = $('syncStatus');
        if (!el) return;
        const map = {
            ok: ['✅ متزامن', 'sync-ok'],
            saving: ['⏳ جاري الحفظ...', 'sync-saving'],
            error: ['⚠️ فشل الحفظ - تأكد من الإنترنت', 'sync-error'],
            big: ['⚠️ حجم البيانات كبير جدًا', 'sync-error']
        };
        const [text, cls] = map[state] || ['', ''];
        el.textContent = text;
        el.className = 'sync-status ' + cls;
    }

    function showError(msg) { $('loginError').textContent = msg || ''; }
    function showLogin(msg) {
        $('authLoading').style.display = 'none';
        $('loginForm').classList.remove('busy');
        $('authOverlay').style.display = 'flex';
        document.body.classList.add('auth-locked');
        showError(msg);
    }
    function showLoading() {
        $('loginForm').classList.add('busy');
        $('authLoading').style.display = 'block';
    }
    function unlockApp() {
        $('authOverlay').style.display = 'none';
        document.body.classList.remove('auth-locked');
    }

    function toEmail(username) {
        const u = username.trim();
        return u.includes('@') ? u : `${u.toLowerCase()}@${window.LOGIN_EMAIL_DOMAIN || 'bahnasawy.app'}`;
    }

    function authMessage(err) {
        switch (err && err.code) {
            case 'auth/invalid-credential':
            case 'auth/user-not-found':
            case 'auth/wrong-password':
            case 'auth/invalid-email':
                return 'اسم المستخدم أو كلمة المرور غير صحيحة';
            case 'auth/too-many-requests': return 'محاولات كتير، جرّب بعد شوية';
            case 'auth/network-request-failed': return 'مفيش اتصال بالإنترنت';
            case 'auth/user-disabled': return 'الحساب ده متوقف';
            default: return 'حصل خطأ في تسجيل الدخول';
        }
    }

    function clearLocalCache() {
        KEYS.forEach(k => localStorage.removeItem(k));
        localStorage.removeItem('cloudUid');
    }

    function collectionFor(user) {
        return window.SHARED_DATA
            ? db.collection('shared')
            : db.collection('users').doc(user.uid).collection('store');
    }

    /* ---------- الرفع ---------- */
    function schedulePush() {
        if (!col || applyingRemote) return;
        setStatus('saving');
        clearTimeout(pushTimer);
        pushTimer = setTimeout(pushNow, 700);
    }

    async function pushNow() {
        if (!col) return;
        const batch = db.batch();
        let count = 0, tooBig = false;
        KEYS.forEach(k => {
            const val = localStorage.getItem(k);
            if (val === null || val === lastSynced[k]) return;
            if (val.length > MAX_DOC_CHARS) { tooBig = true; return; }
            batch.set(col.doc(k), { json: val, updatedAt: firebase.firestore.FieldValue.serverTimestamp() });
            count++;
        });
        if (tooBig) { setStatus('big'); }
        if (!count) { if (!tooBig) setStatus('ok'); return; }
        try {
            await batch.commit();
            KEYS.forEach(k => { const v = localStorage.getItem(k); if (v !== null && v.length <= MAX_DOC_CHARS) lastSynced[k] = v; });
            if (!tooBig) setStatus('ok');
        } catch (e) {
            console.error('فشل رفع البيانات:', e);
            setStatus('error');
        }
    }

    /* ---------- التنزيل والمزامنة اللحظية ---------- */
    function attach(user) {
        col = collectionFor(user);
        let first = true;
        unsub = col.onSnapshot(snap => {
            if (first) {
                first = false;
                const cloudEmpty = snap.empty;
                snap.forEach(d => {
                    const json = d.data().json;
                    if (typeof json === 'string') { localStorage.setItem(d.id, json); lastSynced[d.id] = json; }
                });
                localStorage.setItem('cloudUid', user.uid);
                unlockApp();
                if (!appStarted) { appStarted = true; startApp(); }
                // أول مرة: لو السحابة فاضية والجهاز فيه بيانات قديمة، ارفعها
                if (cloudEmpty) schedulePush(); else setStatus('ok');
                return;
            }
            let changed = false;
            snap.docChanges().forEach(ch => {
                if (ch.type === 'removed' || ch.doc.metadata.hasPendingWrites) return;
                const json = ch.doc.data().json;
                if (typeof json !== 'string' || json === lastSynced[ch.doc.id]) return;
                lastSynced[ch.doc.id] = json;
                if (localStorage.getItem(ch.doc.id) !== json) {
                    localStorage.setItem(ch.doc.id, json);
                    changed = true;
                }
            });
            if (changed) applyRemote();
        }, err => {
            console.error('خطأ في المزامنة:', err);
            if (!appStarted) showLogin(err.code === 'permission-denied'
                ? 'مفيش صلاحية للوصول للبيانات - راجع قواعد Firestore'
                : 'تعذر تحميل البيانات، تأكد من الإنترنت');
            else setStatus('error');
        });
    }

    function applyRemote() {
        applyingRemote = true;
        try {
            if (typeof loadData === 'function') loadData();
            if (typeof loadLicenseWarningDays === 'function') loadLicenseWarningDays();
            const active = document.querySelector('.nav-link.active');
            if (typeof showSection === 'function' && active) showSection(active.getAttribute('data-section'));
            if (typeof renderAllSummaries === 'function') renderAllSummaries();
        } catch (e) { console.error(e); }
        applyingRemote = false;
        setStatus('ok');
    }

    /* ---------- التشغيل ---------- */
    function boot(initializeApp) {
        startApp = initializeApp;
        const cfg = window.FIREBASE_CONFIG || {};
        if (!cfg.apiKey || cfg.apiKey.startsWith('PASTE')) {
            showLogin('لم يتم ضبط Firebase بعد - املأ بيانات firebase-config.js');
            $('loginBtn').disabled = true;
            return;
        }
        firebase.initializeApp(cfg);
        auth = firebase.auth();
        db = firebase.firestore();

        $('loginForm').addEventListener('submit', async e => {
            e.preventDefault();
            showError('');
            showLoading();
            try {
                await auth.signInWithEmailAndPassword(toEmail($('loginUser').value), $('loginPass').value);
            } catch (err) {
                showLogin(authMessage(err));
            }
        });

        $('logoutBtn').addEventListener('click', async () => {
            if (!confirm('تسجيل الخروج؟')) return;
            clearTimeout(pushTimer);
            await pushNow();
            if (unsub) unsub();
            clearLocalCache();
            await auth.signOut();
            location.reload();
        });

        showLoading();
        auth.onAuthStateChanged(user => {
            if (!user) { showLogin(); return; }
            if (appStarted) return;
            // لو الجهاز كان مسجّل بيه يوزر تاني، امسح نسخته المحلية عشان ماتتخلطش
            const prev = localStorage.getItem('cloudUid');
            if (prev && prev !== user.uid && !window.SHARED_DATA) clearLocalCache();
            showLoading();
            attach(user);
        });
    }

    window.CloudSync = { boot, schedulePush };
})();
