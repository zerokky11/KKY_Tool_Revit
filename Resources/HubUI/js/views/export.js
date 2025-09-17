import { clear, div, tdText, toast, setBusy } from '../core/dom.js';
import { renderTopbar } from '../core/topbar.js';
import { post, onHost } from '../core/bridge.js';

const state = { files: [], rows: [], folder: '' };

export function renderExport() {
    const root = document.getElementById('app'); clear(root);
    renderTopbar(root, true, () => { location.hash = ''; });

    const wrap = div('kkyt-split');

    const left = div('kkyt-left');
    const lbar = div('kkyt-toolbar');
    const pick = button('폴더 선택', {}, () => post('export:browse-folder', {}));
    const clearBtn = button('목록 지우기', { variant: 'secondary' }, () => { state.files = []; state.folder = ''; renderFiles(); });
    lbar.append(pick, clearBtn);
    const list = div('kkyt-list'); const info = div('kkyt-hint'); info.textContent = '파일 0개';
    left.append(lbar, list, info);

    const right = div('kkyt-right');
    const rbar = div('kkyt-toolbar');
    const preview = button('미리보기', { id: 'btnExPreview', disabled: true }, () => { setBusy(true, '미리보기 생성…'); post('export:preview', { files: state.files }); });
    const save = button('엑셀 저장…', { id: 'btnExSave', disabled: true }, () => post('export:save-excel', { rows: state.rows || [] }));
    rbar.append(preview, save);

    const tbl = document.createElement('table'); tbl.className = 'kkyt-table';
    const thead = document.createElement('thead');
    thead.innerHTML = '<tr><th>File</th><th>ProjectPoint_E(mm)</th><th>ProjectPoint_N(mm)</th><th>ProjectPoint_Z(mm)</th><th>SurveyPoint_E(mm)</th><th>SurveyPoint_N(mm)</th><th>SurveyPoint_Z(mm)</th><th>TrueNorthAngle(deg)</th></tr>';
    const tbody = document.createElement('tbody'); tbl.append(thead, tbody);

    right.append(rbar, tbl);
    wrap.append(left, right);
    root.append(wrap);

    // === 파일 선택 응답 ===
    onHost('export:files', ({ files }) => {
        state.files = Array.isArray(files) ? files : [];
        state.folder = state.files[0] ? state.files[0].replace(/[/\\][^/\\]+$/, '') : '';
        renderFiles();
    });

    // === 미리보기 결과 ===
    onHost('export:previewed', ({ rows }) => {
        setBusy(false);
        state.rows = Array.isArray(rows) ? rows : [];
        while (tbody.firstChild) tbody.removeChild(tbody.firstChild);
        state.rows.forEach(r => {
            const tr = document.createElement('tr');
            [r.File, r.ProjectPoint_E_mm, r.ProjectPoint_N_mm, r.ProjectPoint_Z_mm, r.SurveyPoint_E_mm, r.SurveyPoint_N_mm, r.SurveyPoint_Z_mm, r.TrueNorthAngle_deg]
                .forEach(v => tr.append(tdText(v)));
            tbody.append(tr);
        });
        document.getElementById('btnExSave').disabled = state.rows.length === 0;
        toast(`미리보기 ${state.rows.length}행`, 'ok');
    });

    // === 저장 결과 ===
    onHost('export:saved', ({ path }) => toast(`엑셀 저장: ${path || ''}`, 'ok', 2600));

    // === 에러 공통 처리(중요) ===
    onHost('revit:error', ({ message }) => { setBusy(false); toast(message || 'Revit 오류가 발생했습니다.', 'err', 3200); });
    onHost('host:error', ({ message }) => { setBusy(false); toast(message || '호스트 오류가 발생했습니다.', 'err', 3200); });

    function renderFiles() {
        while (list.firstChild) list.removeChild(list.firstChild);
        state.files.forEach(p => {
            const it = div('kkyt-file'); const nm = p.split(/[/\\]/).pop();
            it.textContent = `${nm} — ${p}`; list.append(it);
        });
        const folder = state.folder || (state.files[0] ? state.files[0].replace(/[/\\][^/\\]+$/, '') : '');
        info.textContent = `${folder ? ('경로: ' + folder + ' · ') : ''}파일 ${state.files.length}개`;
        document.getElementById('btnExPreview').disabled = state.files.length === 0;
    }
}

function button(text, opts = {}, onClick) {
    const b = document.createElement('button');
    b.textContent = text;
    b.className = 'kkyt-btn ' + (opts.variant === 'danger' ? 'kkyt-danger' : (opts.variant === 'secondary' ? 'kkyt-secondary' : ''));
    if (opts.id) b.id = opts.id; if (opts.title) b.title = opts.title; if (opts.disabled) b.disabled = true;
    if (typeof onClick === 'function') b.addEventListener('click', onClick);
    return b;
}
