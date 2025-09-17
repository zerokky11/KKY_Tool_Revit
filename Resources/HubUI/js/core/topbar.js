// Resources/HubUI/js/core/topbar.js
import { div } from './dom.js';
import { toggleTheme } from './theme.js';
import { setConn, ping, post } from './bridge.js';

export const APP_VERSION = 'v0.9.2';

export function renderTopbar(root, withBack = false, onBack = null) {
    const topbar = div('topbar');
    const brand = div('brand');

    // ============== KKY TOOL 아이콘 (인라인 SVG) ==============
    const logoWrap = document.createElement('span');
    logoWrap.className = 'hub-logo-wrap';

    const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    svg.setAttribute('viewBox', '0 0 48 48');
    svg.setAttribute('class', 'hub-logo');
    svg.setAttribute('aria-hidden', 'true');
    svg.setAttribute('focusable', 'false');

    const NS = svg.namespaceURI;

    // gradient 정의 (외곽선 포인트)
    const defs = document.createElementNS(NS, 'defs');
    const lg = document.createElementNS(NS, 'linearGradient');
    lg.setAttribute('id', 'kkygrad');
    lg.setAttribute('x1', '0'); lg.setAttribute('y1', '0');
    lg.setAttribute('x2', '1'); lg.setAttribute('y2', '1');
    const s1 = document.createElementNS(NS, 'stop'); s1.setAttribute('offset', '0'); s1.setAttribute('stop-color', 'currentColor');
    const s2 = document.createElementNS(NS, 'stop'); s2.setAttribute('offset', '1'); s2.setAttribute('stop-color', 'currentColor'); s2.setAttribute('stop-opacity', '0.55');
    lg.append(s1, s2); defs.append(lg);

    // 육각형
    const hex = document.createElementNS(NS, 'path');
    hex.setAttribute('d', ['M', 24, 6, 'L', 38, 14, 'L', 38, 30, 'L', 24, 38, 'L', 10, 30, 'L', 10, 14, 'Z'].join(' '));
    hex.setAttribute('fill', 'none');
    hex.setAttribute('stroke', 'url(#kkygrad)');
    hex.setAttribute('stroke-width', '2.8');
    hex.setAttribute('stroke-linejoin', 'round');

    // K·K·Y 모노그램
    const mono = document.createElementNS(NS, 'g');
    mono.setAttribute('stroke', 'currentColor');
    mono.setAttribute('stroke-width', '2.6');
    mono.setAttribute('stroke-linecap', 'round');
    mono.setAttribute('stroke-linejoin', 'round');
    mono.setAttribute('fill', 'none');

    const k1a = document.createElementNS(NS, 'path'); k1a.setAttribute('d', 'M12 16 L12 32');
    const k1b = document.createElementNS(NS, 'path'); k1b.setAttribute('d', 'M12 24 L18 17');
    const k1c = document.createElementNS(NS, 'path'); k1c.setAttribute('d', 'M12 24 L18 31');

    const k2a = document.createElementNS(NS, 'path'); k2a.setAttribute('d', 'M22 16 L22 32');
    const k2b = document.createElementNS(NS, 'path'); k2b.setAttribute('d', 'M22 24 L28 17');
    const k2c = document.createElementNS(NS, 'path'); k2c.setAttribute('d', 'M22 24 L28 31');

    const y1 = document.createElementNS(NS, 'path'); y1.setAttribute('d', 'M34 16 L29 21');
    const y2 = document.createElementNS(NS, 'path'); y2.setAttribute('d', 'M34 16 L39 21');
    const y3 = document.createElementNS(NS, 'path'); y3.setAttribute('d', 'M34 21 L34 32');

    mono.append(k1a, k1b, k1c, k2a, k2b, k2c, y1, y2, y3);

    const shadow = document.createElementNS(NS, 'ellipse');
    shadow.setAttribute('cx', '24'); shadow.setAttribute('cy', '40');
    shadow.setAttribute('rx', '10'); shadow.setAttribute('ry', '1.5');
    shadow.setAttribute('fill', 'currentColor'); shadow.setAttribute('opacity', '0.12');

    svg.append(defs, hex, mono, shadow);
    logoWrap.append(svg);

    const h1 = document.createElement('h1'); h1.textContent = 'KKY Tool';
    const sub = document.createElement('span'); sub.className = 'brand-sub';
    sub.textContent = '– Revit 작업 보조 통합 도구';
    const ver = document.createElement('span'); ver.className = 'kkyt-badge'; ver.textContent = APP_VERSION;

    brand.append(logoWrap, h1, sub, ver);

    const actions = div('top-actions');

    // ─────────────────────────────────────────────────────────────
    // HUB 홈 버튼: onBack 미지정 시에도 "확실히" 허브로 이동
    //   1) 커스텀 이벤트로 SPA 라우터가 있으면 위임 (kkyt:go-home)
    //   2) history.back() 시도
    //   3) 실패하면 index.html(허브)로 강제 이동
    // ─────────────────────────────────────────────────────────────
    if (withBack) {
        const backBtn = document.createElement('button');
        backBtn.className = 'kkyt-btn btn-sm kkyt-secondary';
        backBtn.type = 'button';
        backBtn.textContent = '← HUB 홈';

        const smartGoHome = () => {
            try {
                // 1) 앱 레벨 핸들러에게 위임 (필요 시 main.js에서 수신해 처리)
                const ev = new CustomEvent('kkyt:go-home');
                window.dispatchEvent(ev);
            } catch (_) { /* ignore */ }

            // 2) 브라우저 히스토리로 복귀 시도
            const before = location.href;
            try { history.back(); } catch (_) { /* ignore */ }

            // 3) 히스토리가 없으면 루트로 강제 이동
            setTimeout(() => {
                if (location.href === before) {
                    // 현재 경로에서 파일명만 index.html 로 치환
                    const url = new URL(location.href);
                    const parts = url.pathname.split('/');
                    parts[parts.length - 1] = 'index.html';
                    url.pathname = parts.join('/');
                    location.href = url.toString();
                }
            }, 80);
        };

        backBtn.onclick = () => {
            if (typeof onBack === 'function') {
                try { onBack(); } catch (_) { smartGoHome(); }
            } else {
                smartGoHome();
            }
        };
        topbar.append(backBtn);
    }

    topbar.append(brand, actions);
    root.append(topbar);

    renderTopbarChips();
    setConn(true);
}

export function renderTopbarChips() {
    const actions = document.querySelector('.top-actions'); if (!actions) return;
    actions.innerHTML = '';

    const conn = document.createElement('button'); conn.className = 'chip'; conn.id = 'connChip';
    conn.innerHTML = `<span class="dot ok"></span> 연결됨`;
    conn.addEventListener('click', ping);
    actions.append(conn);

    const pin = document.createElement('button'); pin.className = 'chip pin off'; pin.id = 'pinChip'; pin.type = 'button';
    pin.setAttribute('aria-pressed', 'false'); pin.innerHTML = '<span class="lock-ico">🔓</span> 항상 위';
    pin.disabled = false;
    pin.onclick = () => { try { post('ui:toggle-topmost'); } catch (e) { console.error(e); } };
    actions.append(pin);

    const theme = document.createElement('button'); theme.className = 'theme-chip'; theme.id = 'themeBtn'; theme.type = 'button';
    const cur = (document.documentElement.dataset.theme || 'light');
    theme.textContent = cur === 'dark' ? '🌙 다크' : '☀ 라이트';
    theme.onclick = () => {
        toggleTheme();
        theme.textContent = (document.documentElement.dataset.theme === 'dark' ? '🌙 다크' : '☀ 라이트');
    };
    actions.append(theme);

    const help = document.createElement('button'); help.className = 'help-chip'; help.innerText = '❓ 도움말';
    help.onclick = showHelpModal; actions.append(help);
}

export function updateTopMost(on) {
    const pin = document.querySelector('#pinChip'); if (!pin) return;
    pin.disabled = false;
    pin.classList.toggle('on', !!on);
    pin.classList.toggle('off', !on);
    pin.setAttribute('aria-pressed', on ? 'true' : 'false');
    const ico = pin.querySelector('.lock-ico'); if (ico) ico.textContent = on ? '🔒' : '🔓';
}

function showHelpModal() {
    const bd = document.createElement('div'); bd.className = 'modal-backdrop';
    const md = document.createElement('div'); md.className = 'modal';
    const h = document.createElement('header'); h.innerHTML = '<span>도움말 — KKY Tool Hub</span>';
    const body = document.createElement('div'); body.className = 'body';
    body.innerHTML = `
    <div><b>단축키</b></div>
    <ul>
      <li><code>/</code> 검색 포커스</li>
      <li>카드 선택 후 <code>Enter</code> 실행, <code>F</code> 즐겨찾기</li>
      <li><code>Ctrl</code>+<code>Shift</code>+<code>L</code> 테마 전환</li>
      <li><code>Ctrl</code>+<code>Shift</code>+<code>T</code> 항상 위</li>
      <li><code>F1</code> 또는 <code>?</code> 도움말</li>
    </ul>`;
    const foot = document.createElement('div'); foot.className = 'actions';
    const ok = document.createElement('button'); ok.className = 'kkyt-btn btn-sm'; ok.textContent = '닫기'; ok.onclick = () => bd.remove();
    foot.append(ok); md.append(h, body, foot); bd.append(md); document.body.append(bd);
    bd.addEventListener('click', e => { if (e.target === bd) bd.remove(); });
}
