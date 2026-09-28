/* ريسبونسيف الموبايل: القائمة الجانبية + عناوين خلايا الجداول */
(function () {
    const $ = id => document.getElementById(id);

    /* ---------- القائمة الجانبية (درج) ---------- */
    function setMenu(open) {
        const sb = $('sidebar'), bd = $('sidebarBackdrop'), btn = $('menuToggle');
        if (!sb) return;
        sb.classList.toggle('open', open);
        if (bd) bd.classList.toggle('show', open);
        document.body.classList.toggle('menu-open', open);
        if (btn) btn.setAttribute('aria-expanded', open ? 'true' : 'false');
    }

    function initMenu() {
        const btn = $('menuToggle'), bd = $('sidebarBackdrop'), sb = $('sidebar');
        if (!btn || !sb) return;
        btn.addEventListener('click', () => setMenu(!sb.classList.contains('open')));
        if (bd) bd.addEventListener('click', () => setMenu(false));
        sb.addEventListener('click', e => { if (e.target.closest('.nav-link')) setMenu(false); });
        document.addEventListener('keydown', e => { if (e.key === 'Escape') setMenu(false); });
        window.addEventListener('resize', () => { if (window.innerWidth > 900) setMenu(false); });
    }

    /* ---------- تسمية خلايا الجداول (تستخدمها بطاقات الموبايل) ---------- */
    function labelTable(table) {
        const heads = Array.from(table.querySelectorAll('thead th')).map(th => th.textContent.trim());
        if (!heads.length) return;
        table.querySelectorAll('tbody tr').forEach(tr => {
            const cells = tr.children;
            if (cells.length === 1 && cells[0].hasAttribute('colspan')) return;
            Array.from(cells).forEach((td, i) => {
                if (td.tagName !== 'TD') return;
                const label = heads[i] || '';
                if (td.getAttribute('data-label') !== label) td.setAttribute('data-label', label);
            });
        });
    }

    let queued = false;
    function scheduleLabel(table) {
        if (queued) return;
        queued = true;
        requestAnimationFrame(() => {
            queued = false;
            document.querySelectorAll('.table-container .data-table').forEach(labelTable);
        });
    }

    function initTables() {
        const tables = document.querySelectorAll('.table-container .data-table');
        tables.forEach(t => {
            labelTable(t);
            const body = t.querySelector('tbody');
            if (body) new MutationObserver(() => scheduleLabel(t)).observe(body, { childList: true });
        });
    }

    function init() { initMenu(); initTables(); }
    if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init);
    else init();
})();
