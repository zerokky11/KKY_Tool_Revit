import { clear, div, tdText, toast, setBusy } from '../core/dom.js';
import { renderTopbar } from '../core/topbar.js';
import { post, onHost } from '../core/bridge.js';

const SKEY = 'kky_conn_opts';
function loadOpts() {
    try {
        return Object.assign(
            { tol: 1.0, unit: 'inch', param: 'Comments' },
            JSON.parse(localStorage.getItem(SKEY) || '{}')
        );
    } catch {
        return { tol: 1.0, unit: 'inch', param: 'Comments' };
    }
}
function saveOpts(o) { localStorage.setItem(SKEY, JSON.stringify(o)); }

export function renderConn() {
    const root = document.getElementById('app'); clear(root);
    renderTopbar(root, true, () => { location.hash = ''; });

    const opts = loadOpts();

    // 옵션 입력 바
    const inputs = div('kkyt-toolbar');
    const tol = document.createElement('input'); tol.type = 'number'; tol.step = '0.0001'; tol.value = String(opts.tol ?? 1.0);
    const unit = document.createElement('select');
    unit.innerHTML = '<option value="inch">inch</option><option value="mm">mm</option>';
    unit.value = String(opts.unit || 'inch');
    const param = document.createElement('input'); param.type = 'text'; param.value = String(opts.param || 'Comments');
    inputs.append(spanChip('허용범위'), tol, spanChip('단위'), unit, spanChip('파라미터'), param);

    const commit = () => saveOpts({
        tol: parseFloat(tol.value || '1') || 1,
        unit: String(unit.value),
        param: String(param.value || 'Comments')
    });
    tol.addEventListener('change', commit);
    unit.addEventListener('change', commit);
    param.addEventListener('change', commit);

    // 실행/저장 바
    const bar = div('kkyt-toolbar');
    const run = button('검토 시작', {}, () => {
        commit();
        setBusy(true, '커넥터 검토…');
        post('connector:run', {
            tol: parseFloat(tol.value || '1'),
            unit: String(unit.value),
            param: String(param.value || 'Comments')
        });
    });
    const save = button('엑셀 저장…', { id: 'btnConnSave', disabled: true }, () =>
        post('connector:save-excel', { rows: state.rows || [] })
    );
    bar.append(run, save);

    // 결과 테이블
    const tbl = document.createElement('table'); tbl.className = 'kkyt-table';
    const thead = document.createElement('thead');
    thead.innerHTML = '<tr><th>Id1</th><th>Id2</th><th>Category1</th><th>Category2</th><th>Family1</th><th>Family2</th><th>Distance (inch)</th><th>ConnectionType</th><th>ParamName</th><th>Value1</th><th>Value2</th><th>Status</th></tr>';
    const tbody = document.createElement('tbody'); tbl.append(thead, tbody);

    root.append(inputs, bar, tbl);

    const state = { rows: [] };

    function fill(rows) {
        setBusy(false);
        state.rows = Array.isArray(rows) ? rows : [];
        while (tbody.firstChild) tbody.removeChild(tbody.firstChild);
        state.rows.forEach(r => {
            const tr = document.createElement('tr');
            [
                r.Id1, r.Id2, r.Category1, r.Category2, r.Family1, r.Family2,
                r.DistanceInch || r['Distance (inch)'],
                r.ConnectionType, r.ParamName, r.Value1, r.Value2, r.Status
            ].forEach(v => tr.append(tdText(v)));
            tbody.append(tr);
        });
        const saveBtn = document.getElementById('btnConnSave');
        if (saveBtn) saveBtn.disabled = state.rows.length === 0;
        toast(`행 ${state.rows.length}개`, 'ok');
    }

    // ✅ onHost는 (msg) => { ev, payload } 형태만 받음 — 이벤트 이름으로 분기
    onHost(({ ev, payload }) => {
        switch (ev) {
            case 'connector:done':
            case 'connector:loaded':
                fill((payload && payload.rows) || []);
                break;
            case 'connector:saved':
                toast(`엑셀 저장: ${(payload && payload.path) || ''}`, 'ok', 2600);
                break;
            case 'revit:error':
                setBusy(false);
                toast((payload && payload.message) || '오류가 발생했습니다.', 'err', 3200);
                break;
            default:
                // ignore
                break;
        }
    });
}

const spanChip = t => { const s = document.createElement('span'); s.className = 'kkyt-chip'; s.textContent = t; return s; };
function button(text, opts = {}, onClick) {
    const b = document.createElement('button');
    b.textContent = text;
    b.className = 'kkyt-btn ' + (opts.variant === 'danger' ? 'kkyt-danger' : (opts.variant === 'secondary' ? 'kkyt-secondary' : ''));
    if (opts.id) b.id = opts.id; if (opts.title) b.title = opts.title; if (opts.disabled) b.disabled = true;
    if (typeof onClick === 'function') b.addEventListener('click', onClick);
    return b;
}
