// Resources/HubUI/js/views/home.js
import { $, clear, div, debounce, toast } from '../core/dom.js';
import { renderTopbar } from '../core/topbar.js';
import { getFavs, toggleFav, getLast, saveCardOrder, getCardOrder } from '../core/state.js';

const CATS = { dup: '검토', conn: '진단', export: '좌표' };
const QKEY = 'kky_q';
const LAYOUT_KEY = 'kky_home_layout';   // 'card' | 'list'
const CAT_KEY = 'kky_home_cat';         // 'fav' | 'all' | 'modeling' | 'property' | 'utility'

// 사이드바 카테고리
const GROUPS = {
    fav: { icon: '⭐', label: '즐겨찾기' },
    all: { icon: '🗂️', label: '전체' },
    modeling: { icon: '🏗️', label: '모델링 검토' },
    property: { icon: '🧾', label: '속성검토' },
    utility: { icon: '🛠️', label: '유틸리티' }
};

// 각 기능의 그룹 매핑
const CARD_GROUP = {
    dup: 'modeling',   // 중복검토 → 모델링
    conn: 'property',  // 커넥터 진단 → 속성
    export: 'utility'  // 좌표/북각 추출 → 유틸리티
};

export function renderHome() {
    const root = document.getElementById('app'); clear(root);
    renderTopbar(root, false);

    // ---------- 레이아웃 ----------
    const layout = div('home-layout');
    const aside = document.createElement('aside'); aside.className = 'home-sb';
    const main = document.createElement('section'); main.className = 'home-main';

    // ---------- 사이드바(고정) ----------
    const sbHead = div('sb-head');
    const sbTitle = document.createElement('div'); sbTitle.className = 'sb-title'; sbTitle.textContent = '카테고리';
    sbHead.append(sbTitle);

    const sbNav = document.createElement('nav'); sbNav.className = 'sb-nav';
    const items = [];
    function mkItem(id, icon, label, onClick) {
        const btn = document.createElement('button');
        btn.className = 'sb-item'; btn.type = 'button'; btn.dataset.sid = id;
        btn.innerHTML = `<span class="ico">${icon}</span><span class="label">${label}</span>`;
        btn.title = label; btn.onclick = onClick; items.push(btn); return btn;
    }

    // 순서: 즐겨찾기 → 전체 → 모델링 → 속성 → 유틸리티
    mkItem('fav', GROUPS.fav.icon, GROUPS.fav.label, () => { setCat('fav'); applyFilter(q.value); syncSidebarActive(); });
    mkItem('all', GROUPS.all.icon, GROUPS.all.label, () => { setCat('all'); applyFilter(q.value); syncSidebarActive(); });
    mkItem('modeling', GROUPS.modeling.icon, GROUPS.modeling.label, () => { setCat('modeling'); applyFilter(q.value); syncSidebarActive(); });
    mkItem('property', GROUPS.property.icon, GROUPS.property.label, () => { setCat('property'); applyFilter(q.value); syncSidebarActive(); });
    mkItem('utility', GROUPS.utility.icon, GROUPS.utility.label, () => { setCat('utility'); applyFilter(q.value); syncSidebarActive(); });

    const sbFoot = div('sb-foot');
    const helpBtn = mkItem('help', '❓', '도움말', () => { const b = document.querySelector('.help-chip'); if (b) b.click(); });
    sbFoot.append(helpBtn);

    sbNav.append(...items);
    aside.append(sbHead, sbNav, sbFoot);

    // ---------- 메인(탭 제거: 사이드바에서만 전환) ----------
    // 검색 + 보기 전환
    const sbar = div('home-search');
    const q = document.createElement('input'); q.type = 'search'; q.placeholder = '기능 검색 (옵션)';

    const viewToggle = div('view-toggle');
    const segCard = document.createElement('button'); segCard.type = 'button'; segCard.className = 'seg';
    segCard.innerHTML = '<span class="ico">▦</span><span>카드</span>'; segCard.setAttribute('aria-pressed', 'true');
    const segList = document.createElement('button'); segList.type = 'button'; segList.className = 'seg';
    segList.innerHTML = '<span class="ico">≣</span><span>목록</span>'; segList.setAttribute('aria-pressed', 'false');
    viewToggle.append(segCard, segList);
    sbar.append(q, viewToggle);

    // 카드들
    const grid = div('card-grid');
    const last = getLast();
    const order = getCardOrder(['dup', 'conn', 'export']);
    const cards = {
        dup: toolCard('🧭 중복검토', '그룹/삭제↔되돌리기/선택·줌/엑셀(고정 스키마)', 'dup', last),
        conn: toolCard('🔌 커넥터 진단', 'tol/unit/param 검사, Distance(inch) 일관, 엑셀 I/O', 'conn', last),
        export: toolCard('📍 좌표/북각 추출', '폴더 스캔→미리보기→엑셀 저장', 'export', last),
    };
    order.forEach(v => grid.append(cards[v]));

    const empty = div('empty'); empty.textContent = '조건에 맞는 기능이 없어요 🙈';

    main.append(sbar, grid, empty);
    layout.append(aside, main);
    root.append(layout);

    // ---------- 상태 복구 ----------
    if (!localStorage.getItem(CAT_KEY)) setCat('all');
    const savedQ = localStorage.getItem(QKEY) || '';
    if (savedQ) q.value = savedQ;

    const savedLayout = (localStorage.getItem(LAYOUT_KEY) === 'list') ? 'list' : 'card';
    setLayout(savedLayout);

    const debounced = debounce(() => applyFilter(q.value), 200);
    q.addEventListener('input', debounced);
    segCard.addEventListener('click', () => setLayout('card'));
    segList.addEventListener('click', () => setLayout('list'));

    applyFilter(savedQ);
    syncSidebarActive();

    // ===== helpers =====
    function getCat() {
        const c = localStorage.getItem(CAT_KEY);
        return (c === 'fav' || c === 'all' || c === 'modeling' || c === 'property' || c === 'utility') ? c : 'all';
    }
    function setCat(c) {
        const v = (c === 'fav' || c === 'all' || c === 'modeling' || c === 'property' || c === 'utility') ? c : 'all';
        localStorage.setItem(CAT_KEY, v);
    }

    function setLayout(mode) {
        const m = (mode === 'list') ? 'list' : 'card';
        grid.dataset.layout = m;
        segCard.classList.toggle('is-active', m === 'card');
        segList.classList.toggle('is-active', m === 'list');
        segCard.setAttribute('aria-pressed', String(m === 'card'));
        segList.setAttribute('aria-pressed', String(m === 'list'));
        localStorage.setItem(LAYOUT_KEY, m);
    }

    function syncSidebarActive() {
        const cat = getCat();
        items.forEach(b => b.classList.toggle('is-active', b.dataset.sid === cat));
    }

    function escapeHtml(s) { return String(s ?? '').replace(/[&<>"']/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m])); }
    function highlight(text, needle) {
        if (!needle) return escapeHtml(text);
        const re = new RegExp(`(${needle.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})`, 'ig');
        return escapeHtml(text).replace(re, '<mark class="hlt">$1</mark>');
    }

    function applyFilter(needle) {
        const n = (needle || '').trim().toLowerCase();
        localStorage.setItem(QKEY, n);

        const favs = new Set(getFavs());
        const cat = getCat();

        let visible = 0;
        Object.values(cards).forEach(card => {
            const isFav = favs.has(card.dataset.view);
            const byCat =
                (cat === 'all') ? true :
                    (cat === 'fav') ? isFav :
                        (card.dataset.group === cat);

            const matchQ = !n || (card.dataset.search || '').includes(n);
            const show = byCat && matchQ;

            card.style.display = show ? '' : 'none';
            if (show) {
                const t = card.querySelector('[data-title]'),
                    d = card.querySelector('[data-desc]'),
                    c = card.querySelector('[data-cat]');
                if (t) t.innerHTML = highlight(t.dataset.orig, n);
                if (d) d.innerHTML = highlight(d.dataset.orig, n);
                if (c) c.innerHTML = '# ' + highlight(c.dataset.orig, n);
                visible++;
            }
        });
        empty.style.display = visible ? 'none' : '';
    }

    function toolCard(title, desc, view, last) {
        const c = div('card');
        c.tabIndex = 0;
        c.setAttribute('role', 'button');
        c.dataset.view = view;
        c.style.position = 'relative';
        c.dataset.cat = view;
        c.dataset.group = CARD_GROUP[view] || 'modeling';

        const titleRow = div('title');
        const left = document.createElement('div');
        left.className = 'left-title';
        const t = document.createElement('span');
        t.setAttribute('data-title', '');
        t.dataset.orig = title; t.innerHTML = title;
        left.append(t);
        if (last && last.view === view) {
            const r = document.createElement('span');
            r.className = 'recent-badge';
            r.textContent = '최근 실행';
            left.append(r);
        }
        titleRow.append(left);

        // 즐겨찾기
        const star = document.createElement('button');
        const isFav = getFavs().includes(view);
        star.className = 'star' + (isFav ? ' on' : '');
        star.textContent = isFav ? '★' : '☆';
        star.title = '즐겨찾기';
        star.setAttribute('aria-pressed', isFav ? 'true' : 'false');
        star.onclick = e => {
            e.stopPropagation();
            const nowOn = !star.classList.contains('on');
            star.classList.toggle('on', nowOn);
            star.textContent = nowOn ? '★' : '☆';
            star.setAttribute('aria-pressed', nowOn ? 'true' : 'false');
            toggleFav(view);
            applyFilter(q.value);            // 즐겨찾기 화면에서 즉시 반영
        };
        c.append(star);

        const d = div('desc'); d.setAttribute('data-desc', ''); d.dataset.orig = desc; d.innerHTML = desc;

        const footer = div('footer');
        const cat = document.createElement('span'); cat.className = 'cat-badge';
        cat.setAttribute('data-cat', ''); cat.dataset.orig = CATS[view] || '기능';
        cat.innerHTML = '# ' + cat.dataset.orig;

        const open = document.createElement('button');
        open.className = 'kkyt-btn btn-sm open-btn';
        open.type = 'button'; open.textContent = '열기';
        open.onclick = e => { e.stopPropagation(); location.hash = '#' + view; window.dispatchEvent(new HashChangeEvent('hashchange')); };

        footer.append(cat, open);

        c.dataset.search = [title, desc, CATS[view] || ''].join(' ').toLowerCase();
        c.append(titleRow, d, footer);
        c.draggable = true;
        return c;
    }

    // 카드 DnD 저장
    enableCardDnD(grid);
    function enableCardDnD(gridEl) {
        let dragEl = null, placeholder = null;
        function persist() {
            const ids = [...gridEl.querySelectorAll('.card')].map(x => x.dataset.view);
            saveCardOrder(ids); toast('카드 순서가 저장됐어요', 'ok');
        }
        gridEl.addEventListener('dragstart', e => {
            const card = e.target.closest('.card'); if (!card) return;
            dragEl = card; dragEl.classList.add('dragging');
            placeholder = document.createElement('div'); placeholder.className = 'card placeholder'; placeholder.style.height = `${card.offsetHeight}px`;
            e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', card.dataset.view);
        });
        gridEl.addEventListener('dragend', () => {
            if (dragEl) { dragEl.classList.remove('dragging'); dragEl = null; }
            if (placeholder) { placeholder.remove(); placeholder = null; }
            persist();
        });
        gridEl.addEventListener('dragover', e => {
            e.preventDefault();
            const after = getAfter(gridEl, e.clientY);
            if (!placeholder) return;
            if (after == null) gridEl.appendChild(placeholder);
            else gridEl.insertBefore(placeholder, after);
        });
        gridEl.addEventListener('drop', e => {
            e.preventDefault();
            if (dragEl && placeholder) { gridEl.insertBefore(dragEl, placeholder); }
        });
        function getAfter(container, y) {
            const els = [...container.querySelectorAll('.card:not(.dragging)')];
            return els.reduce((closest, child) => {
                const box = child.getBoundingClientRect();
                const offset = y - box.top - box.height / 2;
                return (offset < 0 && offset > closest.offset) ? { offset, element: child } : closest;
            }, { offset: -Infinity }).element || null;
        }
    }
}
