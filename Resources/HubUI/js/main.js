import { initTheme } from './core/theme.js';
import { onHost, post } from './core/bridge.js';
import { updateTopMost } from './core/topbar.js';
import { renderHome } from './views/home.js';
import { renderDup } from './views/dup.js';
import { renderConn } from './views/conn.js';
import { renderExport } from './views/export.js';

initTheme();

// 직전 TopMost 값 기억(중복 수신 무시)
let _lastTop = null;

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', boot);
else boot();

function boot() {
    // 부팅 스켈레톤 제거
    const bootEl = document.getElementById('boot'); if (bootEl) bootEl.remove();
    const app = document.getElementById('app'); if (app) app.hidden = false;

    // 초기 TopMost 상태 동기화(이제 여기서만 전송)
    try { post('ui:query-topmost'); } catch { }

    route();
    window.addEventListener('hashchange', route);

    // 제네릭 onHost: { ev, payload }
    onHost((msg) => {
        try {
            if (!msg || !msg.ev) return;

            // (디버그용) 수신 로그
            try { console.log('[host] ←', msg.ev, msg.payload); } catch { }

            switch (msg.ev) {
                case 'host:topmost': {
                    const on = (msg && typeof msg.payload === 'object') ? !!msg.payload.on : !!msg.payload;
                    // ★ 디듀프: 이전 값과 같으면 무시
                    if (_lastTop === on) return;
                    _lastTop = on;
                    updateTopMost(on);
                    break;
                }
                default:
                    break;
            }
        } catch (e) {
            console.error('[main] onHost dispatch error:', e);
        }
    });
}

function route() {
    const hash = (location.hash || '').replace('#', '');
    switch (hash) {
        case 'dup': return renderDup();
        case 'conn': return renderConn();
        case 'export': return renderExport();
        default: return renderHome();
    }
}
