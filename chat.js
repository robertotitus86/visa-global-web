/* ═══════════════════════════════════════════════════════════════════
   Chat Widget — Asesoría Visa Global
   Primera consulta gratis · Powered by AI · Dirige al diagnóstico $50
   ═══════════════════════════════════════════════════════════════════ */
(function () {
  const BOT_URL  = 'https://visa-global-bot.onrender.com/web-chat';
  const WA_NUM   = '593994442512';
  const DIAG_URL = '/diagnostico.html';
  let history = [], isOpen = false, isTyping = false, sessionRef = 'CHAT-' + Math.floor(Math.random() * 999999);

  /* ── ESTILOS ─────────────────────────────────────────────────── */
  const css = `
    #vg-btn {
      position: fixed; bottom: 1.5rem; right: 1.5rem; z-index: 9000;
      width: 3.5rem; height: 3.5rem; border-radius: 50%; border: none;
      background: #060E1F; cursor: pointer; box-shadow: 0 4px 20px rgba(0,0,0,0.3);
      display: flex; align-items: center; justify-content: center;
      transition: transform 0.2s, box-shadow 0.2s;
    }
    #vg-btn:hover { transform: scale(1.08); box-shadow: 0 6px 28px rgba(0,0,0,0.35); }
    #vg-btn svg { width: 1.6rem; height: 1.6rem; }
    #vg-badge {
      position: absolute; top: -4px; right: -4px; background: #F0B429;
      color: #060E1F; font-family: 'DM Sans', sans-serif; font-size: 0.6rem;
      font-weight: 700; padding: 0.15rem 0.35rem; border-radius: 999px;
      white-space: nowrap; display: none;
    }
    #vg-badge.show { display: block; }
    #vg-panel {
      position: fixed; bottom: 5.5rem; right: 1.5rem; z-index: 8999;
      width: min(380px, calc(100vw - 2rem)); height: 560px;
      background: #fff; border-radius: 1rem; overflow: hidden;
      box-shadow: 0 8px 40px rgba(0,0,0,0.2); display: none;
      flex-direction: column; font-family: 'DM Sans', sans-serif;
      transform: translateY(12px); opacity: 0;
      transition: transform 0.25s ease, opacity 0.25s ease;
    }
    #vg-panel.open { display: flex; transform: translateY(0); opacity: 1; }
    #vg-head {
      background: #060E1F; color: #fff; padding: 1rem 1.1rem;
      display: flex; align-items: center; gap: 0.75rem; flex-shrink: 0;
    }
    #vg-avatar {
      width: 2.25rem; height: 2.25rem; border-radius: 50%;
      background: #F0B429; display: flex; align-items: center; justify-content: center;
      flex-shrink: 0;
    }
    #vg-avatar span { font-size: 0.95rem; font-weight: 700; color: #060E1F; }
    #vg-head-info { flex: 1; }
    #vg-head-info strong { display: block; font-size: 0.9rem; }
    #vg-head-info small { font-size: 0.72rem; color: rgba(255,255,255,0.6); }
    #vg-head-info small span { color: #4ade80; }
    #vg-close { background: none; border: none; color: rgba(255,255,255,0.6); font-size: 1.25rem; cursor: pointer; padding: 0.25rem; line-height: 1; }
    #vg-close:hover { color: #fff; }
    #vg-messages {
      flex: 1; overflow-y: auto; padding: 1rem; display: flex; flex-direction: column; gap: 0.6rem;
      background: #f0f2f5;
    }
    #vg-messages::-webkit-scrollbar { width: 4px; }
    #vg-messages::-webkit-scrollbar-thumb { background: #d1d5db; border-radius: 999px; }
    .vg-msg { max-width: 82%; padding: 0.6rem 0.85rem; border-radius: 0.75rem; font-size: 0.88rem; line-height: 1.5; word-wrap: break-word; }
    .vg-msg.bot { background: #fff; color: #111827; border-bottom-left-radius: 0.2rem; align-self: flex-start; box-shadow: 0 1px 3px rgba(0,0,0,0.06); }
    .vg-msg.user { background: #060E1F; color: #fff; border-bottom-right-radius: 0.2rem; align-self: flex-end; }
    .vg-msg a { color: #F0B429; text-decoration: underline; }
    .vg-msg a.vg-cta-btn {
      display: inline-block; margin-top: 0.5rem; background: #F0B429; color: #060E1F;
      padding: 0.45rem 1rem; border-radius: 0.4rem; font-weight: 700; font-size: 0.82rem;
      text-decoration: none;
    }
    .vg-typing { display: flex; align-items: center; gap: 4px; padding: 0.65rem 0.85rem; background: #fff; border-radius: 0.75rem; border-bottom-left-radius: 0.2rem; align-self: flex-start; box-shadow: 0 1px 3px rgba(0,0,0,0.06); }
    .vg-typing span { width: 7px; height: 7px; background: #9ca3af; border-radius: 50%; animation: vg-bounce 1.2s infinite; }
    .vg-typing span:nth-child(2) { animation-delay: 0.2s; }
    .vg-typing span:nth-child(3) { animation-delay: 0.4s; }
    @keyframes vg-bounce { 0%,80%,100%{transform:scale(0.6)} 40%{transform:scale(1)} }
    #vg-quick { display: flex; flex-wrap: wrap; gap: 0.4rem; padding: 0.5rem 1rem; background: #f0f2f5; flex-shrink: 0; }
    .vg-q-btn { background: #fff; border: 1.5px solid #d1d5db; border-radius: 999px; padding: 0.35rem 0.8rem; font-family: 'DM Sans',sans-serif; font-size: 0.78rem; color: #374151; cursor: pointer; transition: all 0.15s; white-space: nowrap; }
    .vg-q-btn:hover { border-color: #060E1F; color: #060E1F; }
    #vg-input-row { display: flex; gap: 0.5rem; padding: 0.75rem; background: #fff; border-top: 1px solid #f3f4f6; flex-shrink: 0; }
    #vg-input { flex: 1; border: 1.5px solid #e5e7eb; border-radius: 1.5rem; padding: 0.55rem 1rem; font-family: 'DM Sans',sans-serif; font-size: 0.9rem; outline: none; resize: none; transition: border-color 0.15s; }
    #vg-input:focus { border-color: #060E1F; }
    #vg-send { background: #060E1F; border: none; border-radius: 50%; width: 2.2rem; height: 2.2rem; cursor: pointer; display: flex; align-items: center; justify-content: center; flex-shrink: 0; transition: background 0.15s; }
    #vg-send:hover { background: #0d1e3a; }
    #vg-send svg { width: 1rem; height: 1rem; fill: #fff; }
    #vg-free-tag { background: #f0f2f5; text-align: center; padding: 0.35rem; font-size: 0.7rem; color: #9ca3af; flex-shrink: 0; }
  `;

  /* ── DOM ─────────────────────────────────────────────────────── */
  function inject() {
    const style = document.createElement('style');
    style.textContent = css;
    document.head.appendChild(style);

    document.body.insertAdjacentHTML('beforeend', `
      <button id="vg-btn" onclick="vgToggle()" title="Consulta gratuita">
        <span id="vg-badge"></span>
        <svg viewBox="0 0 24 24" fill="none" xmlns="http://www.w3.org/2000/svg">
          <path d="M20 2H4C2.9 2 2 2.9 2 4v18l4-4h14c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2z" fill="#F0B429"/>
          <path d="M7 9h10M7 13h7" stroke="#060E1F" stroke-width="1.5" stroke-linecap="round"/>
        </svg>
      </button>
      <div id="vg-panel">
        <div id="vg-head">
          <div id="vg-avatar"><span>VG</span></div>
          <div id="vg-head-info">
            <strong>Asesoria Visa Global</strong>
            <small><span>●</span> Consulta gratuita · En linea</small>
          </div>
          <button id="vg-close" onclick="vgToggle()">✕</button>
        </div>
        <div id="vg-messages"></div>
        <div id="vg-quick"></div>
        <div id="vg-free-tag">Consulta gratuita — sin compromiso</div>
        <div id="vg-input-row">
          <textarea id="vg-input" rows="1" placeholder="Escribe tu pregunta..." onkeydown="vgKey(event)"></textarea>
          <button id="vg-send" onclick="vgSend()">
            <svg viewBox="0 0 24 24"><path d="M2.01 21L23 12 2.01 3 2 10l15 2-15 2z"/></svg>
          </button>
        </div>
      </div>
    `);
  }

  /* ── MENSAJES ────────────────────────────────────────────────── */
  function addMsg(text, role) {
    const div = document.createElement('div');
    div.className = 'vg-msg ' + role;
    div.innerHTML = text;
    document.getElementById('vg-messages').appendChild(div);
    scrollBottom();
  }

  function showTyping() {
    const el = document.createElement('div');
    el.className = 'vg-typing'; el.id = 'vg-typing-el';
    el.innerHTML = '<span></span><span></span><span></span>';
    document.getElementById('vg-messages').appendChild(el);
    scrollBottom();
  }

  function hideTyping() {
    const el = document.getElementById('vg-typing-el');
    if (el) el.remove();
  }

  function scrollBottom() {
    const msgs = document.getElementById('vg-messages');
    msgs.scrollTop = msgs.scrollHeight;
  }

  function setQuickReplies(opts) {
    const qr = document.getElementById('vg-quick');
    qr.innerHTML = opts.map(o =>
      `<button class="vg-q-btn" onclick="vgQuick('${o.replace(/'/g,"\\'")}')"> ${o}</button>`
    ).join('');
  }
  function clearQuickReplies() { document.getElementById('vg-quick').innerHTML = ''; }

  /* ── SALUDO INICIAL ──────────────────────────────────────────── */
  function greet() {
    setTimeout(() => {
      addMsg('Hola, soy el asistente de <strong>Asesoria Visa Global</strong>. Esta consulta es totalmente gratuita.', 'bot');
    }, 400);
    setTimeout(() => {
      addMsg('¿A que pais deseas viajar?', 'bot');
      setQuickReplies(['Estados Unidos', 'España', 'Europa Schengen', 'Reino Unido', 'Otro destino']);
    }, 900);
  }

  /* ── ENVIAR MENSAJE ──────────────────────────────────────────── */
  async function sendMessage(text) {
    if (!text.trim() || isTyping) return;
    clearQuickReplies();
    addMsg(text, 'user');
    history.push({ role: 'user', parts: [{ text }] });
    document.getElementById('vg-input').value = '';

    isTyping = true;
    showTyping();

    try {
      const res = await fetch(BOT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ session_id: sessionRef, message: text })
      });
      const data = await res.json();
      hideTyping();
      if (data.reply) {
        addMsg(data.reply, 'bot');
        history.push({ role: 'assistant', content: data.reply });
        if (data.quick_replies && data.quick_replies.length) setQuickReplies(data.quick_replies);
      }
    } catch(e) {
      hideTyping();
      addMsg('Hubo un problema de conexion. Escribenos directamente al WhatsApp: <a href="https://wa.me/' + WA_NUM + '" target="_blank">+593 99 444 2512</a>', 'bot');
    }
    isTyping = false;
  }

  /* ── EXPOSTOS GLOBALES ───────────────────────────────────────── */
  window.vgToggle = function() {
    const panel = document.getElementById('vg-panel');
    isOpen = !isOpen;
    if (isOpen) {
      panel.style.display = 'flex';
      setTimeout(() => panel.classList.add('open'), 10);
      document.getElementById('vg-badge').classList.remove('show');
      if (history.length === 0) greet();
    } else {
      panel.classList.remove('open');
      setTimeout(() => { panel.style.display = 'none'; }, 260);
    }
  };

  window.vgSend = function() {
    sendMessage(document.getElementById('vg-input').value);
  };

  window.vgKey = function(e) {
    if (e.key === 'Enter' && !e.shiftKey) { e.preventDefault(); window.vgSend(); }
  };

  window.vgQuick = function(text) { sendMessage(text); };

  /* ── BADGE DE BIENVENIDA ─────────────────────────────────────── */
  setTimeout(() => {
    if (!isOpen) document.getElementById('vg-badge').textContent = 'Gratis';
    document.getElementById('vg-badge').classList.add('show');
  }, 3000);

  /* ── INIT ────────────────────────────────────────────────────── */
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', inject);
  else inject();
})();
