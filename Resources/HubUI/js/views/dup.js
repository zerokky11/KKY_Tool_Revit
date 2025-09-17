// Resources/HubUI/js/views/dup.js
import { clear, div, toast } from '../core/dom.js';
import { renderTopbar } from '../core/topbar.js';
import { onHost, post } from '../core/bridge.js';

// 고정 스키마(열 순서)
const COLS = ['ElementId', '카테고리', '패밀리:타입', '연결수/연결객체', '상태', '액션'];

// ✅ 행 목록만 받는 이벤트(dup:result/dup:done 은 제외!)
const RESP_ROWS_EVENTS = [
    'dup:list', 'dup:rows',
    'duplicate:list'
];

// 액션 관련 이벤트명
const EV_DELETE_REQ = 'duplicate:delete';
const EV_RESTORE_REQ = 'duplicate:restore';
const EV_SELECT_REQ = 'duplicate:select';
const EV_EXPORT_REQ = 'duplicate:export';

// 호스트에서 오는 단건/다건 변화 이벤트(두 이름 다 수신)
const EV_DELETED_ONE = 'dup:deleted';
const EV_RESTORED_ONE = 'dup:restored';
const EV_DELETED_MULTI = 'duplicate:delete';
const EV_RESTORED_MULTI = 'duplicate:restore';

// 엑셀 결과 (두 이름 다 수신)
const EV_EXPORTED_A = 'duplicate:export';
const EV_EXPORTED_B = 'dup:exported';

export function renderDup() {
    const root = document.getElementById('app'); clear(root);
    renderTopbar(root, true);

    // ====== 컨테이너 ======
    const page = div('dup-page');
    const topbar = div('dup-toolbar');

    const runBtn = btn('검토 시작', 'primary', onRun);
    const exportBtn = btn('엑셀 내보내기', '', onExport); exportBtn.disabled = true;
    const bulkBtn = btn('선택 삭제 (0)', 'danger', onBulkDelete); bulkBtn.disabled = true;

    const filt = document.createElement('input');
    filt.type = 'search';
    filt.className = 'kkyt-input';
    filt.placeholder = '필터: 카테고리/패밀리/타입/ID';
    filt.oninput = () => paintGroups();

    const sortSel = document.createElement('select');
    sortSel.className = 'kkyt-input';
    sortSel.innerHTML = `
    <option value="size-desc">그룹 큰순</option>
    <option value="size-asc">그룹 작은순</option>
    <option value="cat">카테고리</option>
    <option value="fam">패밀리:타입</option>
  `;
    sortSel.onchange = () => paintGroups();

    const summary = div('dup-summary', `
    <span class="chip">그룹: <b data-sum-g>0</b></span>
    <span class="chip">요소: <b data-sum-r>0</b></span>
    <span class="chip warn">삭제후보: <b data-sum-c>0</b></span>
  `);

    topbar.append(runBtn, exportBtn, summary, bulkBtn, div('flex-spacer'), filt, sortSel);
    page.append(topbar);

    // 헤더(체크 올)
    const head = div('dup-head');
    const checkAll = document.createElement('input');
    checkAll.type = 'checkbox'; checkAll.className = 'ck-all';
    checkAll.onchange = () => {
        selected.clear();
        if (checkAll.checked) rows.forEach(r => selected.add(String(r.id)));
        syncChecks(); setBulkCount(selected.size);
    };
    head.append(cell(checkAll, 'ck'));
    COLS.forEach(label => head.append(cell(label, 'th')));
    page.append(head);

    // 그룹/본문
    const body = div('dup-body'); // 그룹 카드들이 들어감
    page.append(body);

    root.append(page);

    // ===== 상태 =====
    let rows = [];            // 정규화된 행
    let groups = [];          // 그루핑 결과
    let selected = new Set(); // 선택된 id
    let deleted = new Set();  // 삭제된 id
    let expanded = new Set(); // 펼친 그룹키
    let waitTimer = null;

    // ===== 에러/예외 공통 처리 =====
    onHost('revit:error', ({ message }) => { setLoading(false); toast(message || 'Revit 오류가 발생했습니다.', 'err', 3200); });
    onHost('host:error', ({ message }) => { setLoading(false); toast(message || '호스트 오류가 발생했습니다.', 'err', 3200); });

    // ===== 모든 호스트 이벤트 수신 =====
    onHost(({ ev, payload }) => {
        try { console.debug('[dup] IN', ev, payload); } catch { }

        // 1) 목록 응답: rows만 렌더
        if (RESP_ROWS_EVENTS.includes(ev)) {
            setLoading(false);
            const list = payload?.rows ?? payload?.data ?? payload ?? [];
            handleRows(list);
            return;
        }

        // 2) 통계 이벤트: 테이블 건드리지 말고 요약만 갱신
        if (ev === 'dup:result') {
            setLoading(false);
            const groupsCnt = Number(payload?.groups || 0);
            const candCnt = Number(payload?.candidates || 0);
            setSummary(groupsCnt, rows.length, candCnt);
            return;
        }

        // 3) 삭제/복구
        if (ev === EV_DELETED_ONE) { const id = String(payload?.id ?? ''); if (id) { deleted.add(id); updateRowStates(); } return; }
        if (ev === EV_RESTORED_ONE) { const id = String(payload?.id ?? ''); if (id) { deleted.delete(id); updateRowStates(); } return; }
        if (ev === EV_DELETED_MULTI) { toIdArray(payload?.ids).forEach(id => deleted.add(id)); updateRowStates(); return; }
        if (ev === EV_RESTORED_MULTI) { toIdArray(payload?.ids).forEach(id => deleted.delete(id)); updateRowStates(); return; }

        // 4) 선택/줌 응답은 무시 가능
        if (ev === EV_SELECT_REQ) return;

        // 5) 엑셀 저장 결과
        if (ev === EV_EXPORTED_A || ev === EV_EXPORTED_B) {
            if (payload?.ok || payload?.path) toast('엑셀로 내보냈습니다', 'ok');
            else toast('엑셀 내보내기 실패', 'err');
            return;
        }
    });

    // ===== 실행/로딩/타임아웃 =====
    function setLoading(on) {
        runBtn.disabled = on;
        runBtn.textContent = on ? '검토 중…' : '검토 시작';
        if (!on && waitTimer) { clearTimeout(waitTimer); waitTimer = null; }
    }

    function onRun() {
        // 초기화
        setLoading(true);
        exportBtn.disabled = true;
        body.innerHTML = '';
        selected.clear(); deleted.clear(); setBulkCount(0);
        setSummary(0, 0, 0);

        // 10초 타임아웃
        waitTimer = setTimeout(() => {
            setLoading(false);
            toast('응답이 없습니다. Add-in 이벤트명을 확인하세요 (예: dup:list).', 'err');
            console.warn('[dup] timeout: no response within 10s after dup:run');
        }, 10000);

        // 요청 송신
        post('dup:run', {});
    }

    function onExport() { post(EV_EXPORT_REQ, {}); }

    function onBulkDelete() {
        const ids = [...selected].filter(Boolean);
        if (!ids.length) return;
        post(EV_DELETE_REQ, { ids });
    }

    // ===== 응답 처리 =====
    function handleRows(listLike) {
        const list = Array.isArray(listLike) ? listLike : [];
        rows = list.map(normalizeRow);

        // 그룹 만들기: (카테고리/패밀리:타입 + 동일 연결세트) 기준
        groups = buildGroups(rows);

        const gCnt = groups.length;
        const candCnt = rows.filter(r => r.candidate).length;
        setSummary(gCnt, rows.length, candCnt);

        exportBtn.disabled = rows.length === 0;
        setLoading(false);

        // 펼침 상태 초기값: 그룹 수가 적으면 모두 펼침
        expanded = new Set(groups.slice(0, 10).map(g => g.key));
        paintGroups();
    }

    // ===== 렌더: 그룹 카드 =====
    function paintGroups() {
        body.innerHTML = '';

        const q = (filt.value || '').trim().toLowerCase();
        const ord = sortSel.value;

        let list = groups.slice();

        // 필터
        if (q) {
            list = list.filter(g => {
                const t = `${g.category} ${g.family} ${g.type} ${g.ids.join(' ')}`.toLowerCase();
                return t.includes(q);
            });
        }

        // 정렬
        list.sort((a, b) => {
            if (ord === 'size-asc') return a.rows.length - b.rows.length;
            if (ord === 'cat') return a.category.localeCompare(b.category) || a.family.localeCompare(b.family) || a.type.localeCompare(b.type);
            if (ord === 'fam') return (a.family + ' ' + a.type).localeCompare(b.family + ' ' + b.type) || a.category.localeCompare(b.category);
            return b.rows.length - a.rows.length; // size-desc
        });

        // 그룹 카드 생성
        list.forEach(g => {
            const card = div('dup-grp');

            // Header
            const h = div('grp-h');
            const left = div('grp-txt');
            const famType = (g.family || g.type) ? `${g.family || '—'}${g.type ? ' : ' + g.type : ''}` : '—';
            left.innerHTML = `
        <div class="grp-line">
          <span class="grp-cat mono">${esc(g.category || '—')}</span>
          <span class="grp-sep">·</span>
          <span class="grp-fam">${esc(famType)}</span>
        </div>
        <div class="grp-line sub">
          <span class="chip">요소 ${g.rows.length}</span>
          <span class="chip alt">연결세트 ${esc(g.ids.slice(0, 3).join(', '))}${g.ids.length > 3 ? ' 외 ' + (g.ids.length - 3) : ''}</span>
        </div>
      `;
            const right = div('grp-actions');
            const toggle = btn(expanded.has(g.key) ? '접기' : '펼치기', '', () => { toggleGroup(g.key); });
            const selectAll = btn('이 그룹 모두 선택', '', () => { g.rows.forEach(r => selected.add(String(r.id))); syncChecks(); setBulkCount(selected.size); });
            const delSel = btn('이 그룹 선택 삭제', 'danger', () => {
                const ids = g.rows.map(r => String(r.id));
                if (!ids.length) return;
                post(EV_DELETE_REQ, { ids });
            });
            right.append(toggle, selectAll, delSel);
            h.append(left, right);

            card.append(h);

            // Body (table)
            const tbl = div('grp-body');
            if (expanded.has(g.key)) {
                tbl.append(renderGroupHeader()); // 소그룹 헤더
                g.rows.forEach(r => tbl.append(renderRow(r)));
            }
            card.append(tbl);

            body.append(card);
        });

        syncChecks();
        updateRowStates();
    }

    function renderGroupHeader() {
        const row = div('dup-row head');
        row.append(cell('', 'ck'));
        COLS.forEach(label => row.append(cell(label, 'th')));
        return row;
    }

    function renderRow(r) {
        const row = div('dup-row'); row.dataset.id = r.id;

        const ck = document.createElement('input');
        ck.type = 'checkbox'; ck.className = 'ck';
        ck.onchange = () => {
            if (ck.checked) selected.add(String(r.id));
            else selected.delete(String(r.id));
            setBulkCount(selected.size);
        };
        row.append(cell(ck, 'ck'));

        row.append(cell(r.id ?? '-', 'td mono'));
        row.append(cell(r.category || '—', 'td'));
        row.append(cell(`${r.family || '—'}${r.type ? ' : ' + r.type : ''}`, 'td'));

        const conn = (r.connectedCount ?? 0);
        const connIds = r.connectedIds?.length
            ? r.connectedIds.slice(0, 3).join(', ') + (r.connectedIds.length > 3 ? ` 외 ${r.connectedIds.length - 3}` : '')
            : '—';
        row.append(cell(`${conn} / ${connIds}`, 'td'));

        const st = document.createElement('span');
        st.className = 'badge ' + (r.candidate ? 'warn' : 'ok');
        st.textContent = r.candidate ? '삭제후보' : '정상';
        row.append(cell(st, 'td'));

        const act = div('row-actions');
        const delBtn = btn(r.deleted ? '되돌리기' : '삭제', r.deleted ? '' : 'danger', () => {
            const ids = [r.id];
            if (delBtn.dataset.mode === 'restore') post(EV_RESTORE_REQ, { ids });
            else post(EV_DELETE_REQ, { ids });
        });
        delBtn.dataset.mode = r.deleted ? 'restore' : 'delete';

        const selBtn = btn('선택/줌', '', () => post(EV_SELECT_REQ, { id: r.id }));
        act.append(delBtn, selBtn);
        row.append(cell(act, 'td'));

        return row;
    }

    function toggleGroup(key) {
        if (expanded.has(key)) expanded.delete(key);
        else expanded.add(key);
        paintGroups();
    }

    function updateRowStates() {
        [...body.querySelectorAll('.dup-row')].forEach(row => {
            const id = row.dataset.id;
            if (!id) return;
            const isDel = deleted.has(String(id));
            row.classList.toggle('is-deleted', isDel);
            const delBtn = row.querySelector('.row-actions button');
            if (delBtn) {
                delBtn.textContent = isDel ? '되돌리기' : '삭제';
                delBtn.className = 'kkyt-btn ' + (isDel ? '' : 'danger');
                delBtn.dataset.mode = isDel ? 'restore' : 'delete';
            }
            const ck = row.querySelector('input.ck');
            if (ck) ck.checked = selected.has(String(id));
        });
    }

    function syncChecks() {
        // 상단 전체선택 체크상태 맞추기
        const allIds = rows.map(r => String(r.id));
        const allSelected = allIds.length > 0 && allIds.every(id => selected.has(id));
        const ckAll = page.querySelector('.dup-head .ck-all');
        if (ckAll) ckAll.checked = allSelected;
    }

    function setSummary(g, r, c) {
        const gEl = page.querySelector('[data-sum-g]');
        const rEl = page.querySelector('[data-sum-r]');
        const cEl = page.querySelector('[data-sum-c]');
        if (gEl) gEl.textContent = g;
        if (rEl) rEl.textContent = r;
        if (cEl) cEl.textContent = c;
    }

    function setBulkCount(n) {
        bulkBtn.dataset.count = String(n);
        bulkBtn.textContent = `선택 삭제 (${n})`;
        bulkBtn.disabled = n === 0;
    }

    // ===== 유틸 =====
    function btn(label, tone, handler) {
        const b = document.createElement('button');
        b.className = 'kkyt-btn ' + (tone || '');
        b.type = 'button';
        b.textContent = label;
        b.onclick = handler;
        return b;
    }
    function cell(content, cls) {
        const c = document.createElement('div');
        c.className = 'cell ' + (cls || '');
        if (content instanceof Node) c.append(content);
        else c.textContent = content;
        return c;
    }
    function toIdArray(val) {
        if (!val) return [];
        if (Array.isArray(val)) return val.map(v => String(v));
        return [String(val)];
    }
    function esc(s) { return String(s ?? '').replace(/[&<>"']/g, m => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[m])); }

    // payload → 내부 표준으로 정규화
    function normalizeRow(r) {
        const id = safeId(r.elementId ?? r.ElementId ?? r.id ?? r.Id);
        const category = val(r.category ?? r.Category);
        const family = val(r.family ?? r.Family);
        const type = val(r.type ?? r.Type);

        const connectedIdsRaw = r.connectedIds ?? r.ConnectedIds ?? r.links ?? r.Links ?? r.connected ?? [];
        const connectedIds = Array.isArray(connectedIdsRaw)
            ? connectedIdsRaw.map(String)
            : (typeof connectedIdsRaw === 'string' && connectedIdsRaw.length
                ? connectedIdsRaw.split(/[,\s]+/).filter(Boolean)
                : []);
        const connectedCount =
            Number.isFinite(r.connectedCount) ? r.connectedCount :
                Number.isFinite(r.ConnectedCount) ? r.ConnectedCount :
                    connectedIds.length;

        const candidate = !!(r.candidate ?? r.isCandidate ?? r.deleteCandidate ?? r.DeleteCandidate);
        const deletedFlag = !!(r.deleted ?? r.isDeleted ?? r.Deleted);

        return { id: id || '-', category, family, type, connectedIds, connectedCount, candidate, deleted: deletedFlag };
    }

    // (카테고리/패밀리:타입 + 동일 연결세트) 기준의 그룹 구성
    function buildGroups(rs) {
        const map = new Map();
        for (const r of rs) {
            const cluster = [String(r.id), ...r.connectedIds.map(String)]
                .filter(Boolean)
                .map(x => x.trim())
                .sort((a, b) => Number(a) - Number(b));    // 동일한 세트면 동일한 정렬 결과
            const clusterKey = cluster.join(',');

            const key = [
                r.category || '',
                r.family || '',
                r.type || '',
                clusterKey
            ].join('|');

            let g = map.get(key);
            if (!g) {
                g = {
                    key,
                    category: r.category || '',
                    family: r.family || '',
                    type: r.type || '',
                    ids: cluster,
                    rows: []
                };
                map.set(key, g);
            }
            g.rows.push(r);
        }
        return [...map.values()];
    }

    function safeId(v) { if (v === 0) return 0; if (v == null) return ''; return String(v); }
    function val(v) { return v == null || v === '' ? '' : String(v); }
}
