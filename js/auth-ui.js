/* ============================================================
   AUTH-UI — Barra superior de estado de conexión + acciones
   ============================================================
   Requiere: SP_CONFIG, Auth, SPAdapter, DB
   Estados:
   - Modo localStorage → barra naranja PROMINENTE invitando a conectar
   - Modo sharepoint sin login → naranja invitando a login M365
   - Modo sharepoint con login → verde con avatar + menú desplegable
============================================================ */
(function(){
  const STYLE = `
    .sp-statusbar {
      position: sticky; top: 0; z-index: 1000;
      display: flex; align-items: center; justify-content: space-between;
      gap: 12px; padding: 10px 18px;
      font: 13px system-ui, -apple-system, "Segoe UI", Arial, sans-serif;
      border-bottom: 1px solid rgba(0,0,0,.1);
      box-shadow: 0 2px 6px rgba(0,0,0,.04);
    }
    .sp-statusbar.local   { background: linear-gradient(90deg,#fff7ed,#fed7aa); color: #9a3412; }
    .sp-statusbar.sp-off  { background: linear-gradient(90deg,#fff7ed,#fdba74); color: #9a3412; }
    .sp-statusbar.sp-on   { background: linear-gradient(90deg,#ecfdf5,#d1fae5); color: #065f46; }
    .sp-statusbar.sp-load { background: linear-gradient(90deg,#eff6ff,#bfdbfe); color: #1e40af; }
    .sp-statusbar.sp-err  { background: linear-gradient(90deg,#fef2f2,#fecaca); color: #991b1b; }

    .sp-statusbar .left   { display: flex; align-items: center; gap: 12px; flex: 1; }
    .sp-statusbar .icono  { font-size: 18px; line-height: 1; }
    .sp-statusbar .titulo { font-weight: 700; font-size: 13.5px; }
    .sp-statusbar .subtxt { font-size: 11.5px; opacity: .85; margin-top: 2px; }
    .sp-statusbar .info-block { display: flex; flex-direction: column; line-height: 1.25; }

    .sp-statusbar .dot {
      width: 10px; height: 10px; border-radius: 50%; display: inline-block;
      flex-shrink: 0;
    }
    .sp-statusbar.local   .dot { background: #f97316; }
    .sp-statusbar.sp-off  .dot { background: #ea580c; animation: sp-pulse 1.5s ease-in-out infinite; }
    .sp-statusbar.sp-on   .dot { background: #10b981; }
    .sp-statusbar.sp-load .dot { background: #3b82f6; animation: sp-pulse 1s ease-in-out infinite; }
    .sp-statusbar.sp-err  .dot { background: #ef4444; }
    @keyframes sp-pulse { 0%,100%{opacity:1;transform:scale(1)} 50%{opacity:.4;transform:scale(1.4)} }

    .sp-statusbar .actions { display: flex; gap: 8px; align-items: center; }

    /* Botones genéricos */
    .sp-statusbar button {
      font: inherit; padding: 7px 14px; border-radius: 6px;
      border: 1px solid currentColor; background: transparent; color: inherit;
      cursor: pointer; font-weight: 600; font-size: 12.5px;
      display: inline-flex; align-items: center; gap: 6px;
      transition: all .15s ease;
    }
    .sp-statusbar button:hover { background: rgba(0,0,0,.06); transform: translateY(-1px); }
    .sp-statusbar button:active { transform: translateY(0); }
    .sp-statusbar button.primary {
      background: currentColor; color: white !important;
      border-color: transparent; padding: 8px 18px;
      box-shadow: 0 2px 8px rgba(0,0,0,.15);
    }
    .sp-statusbar button.primary:hover {
      filter: brightness(1.1);
      box-shadow: 0 4px 12px rgba(0,0,0,.2);
    }
    .sp-statusbar button.danger { color: #b91c1c; border-color: #b91c1c; }
    .sp-statusbar button.danger:hover { background: #fee2e2; }

    /* Avatar circular del usuario logueado */
    .sp-avatar {
      width: 36px; height: 36px; border-radius: 50%;
      background: #10b981; color: white;
      display: inline-flex; align-items: center; justify-content: center;
      font-weight: 700; font-size: 14px;
      box-shadow: 0 2px 6px rgba(16,185,129,.4);
      border: 2px solid white;
      flex-shrink: 0;
    }

    /* Menú desplegable de usuario */
    .sp-user-menu { position: relative; }
    .sp-user-menu-btn {
      display: inline-flex; align-items: center; gap: 8px;
      padding: 4px 10px 4px 4px; border-radius: 22px;
      background: rgba(255,255,255,.5);
      border: 1px solid rgba(0,0,0,.1);
      cursor: pointer;
      transition: all .15s;
    }
    .sp-user-menu-btn:hover { background: rgba(255,255,255,.8); }
    .sp-user-menu-btn .nombre { font-weight: 600; font-size: 12.5px; }
    .sp-user-menu-dropdown {
      display: none;
      position: absolute; right: 0; top: calc(100% + 6px);
      background: white; border: 1px solid #e5e7eb; border-radius: 8px;
      box-shadow: 0 8px 24px rgba(0,0,0,.12);
      min-width: 240px; padding: 4px;
      z-index: 1100;
      color: #1f2937;
    }
    .sp-user-menu.open .sp-user-menu-dropdown { display: block; }
    .sp-user-menu-dropdown .item {
      display: flex; align-items: center; gap: 10px;
      padding: 10px 12px; border-radius: 6px; cursor: pointer;
      font-size: 13px; color: #374151;
      transition: background .12s;
    }
    .sp-user-menu-dropdown .item:hover { background: #f3f4f6; }
    .sp-user-menu-dropdown .item.danger { color: #b91c1c; }
    .sp-user-menu-dropdown .item.danger:hover { background: #fee2e2; }
    .sp-user-menu-dropdown .item .ico { font-size: 16px; width: 20px; text-align: center; }
    .sp-user-menu-dropdown .header {
      padding: 12px 14px 10px; border-bottom: 1px solid #e5e7eb;
      margin-bottom: 4px;
    }
    .sp-user-menu-dropdown .header .nombre { font-weight: 700; font-size: 13.5px; color: #111827; }
    .sp-user-menu-dropdown .header .email  { font-size: 11.5px; color: #6b7280; margin-top: 2px; word-break: break-all; }
    .sp-user-menu-dropdown .divider { height: 1px; background: #e5e7eb; margin: 4px 8px; }

    /* Modal de bienvenida (primera vez en modo local) */
    .sp-welcome-overlay {
      position: fixed; inset: 0; background: rgba(0,0,0,.55);
      display: flex; align-items: center; justify-content: center;
      z-index: 5000; padding: 20px;
      animation: sp-fadein .25s ease;
    }
    @keyframes sp-fadein { from{opacity:0} to{opacity:1} }
    .sp-welcome {
      background: white; border-radius: 12px; padding: 28px 28px 22px;
      max-width: 520px; width: 100%;
      box-shadow: 0 20px 60px rgba(0,0,0,.3);
      font: 14px system-ui, -apple-system, "Segoe UI", Arial, sans-serif;
      color: #1f2937;
    }
    .sp-welcome h2 { margin: 0 0 8px; font-size: 22px; color: #0353a4; }
    .sp-welcome p  { margin: 6px 0 14px; color: #4b5563; line-height: 1.55; }
    .sp-welcome ol { margin: 12px 0 18px; padding-left: 22px; color: #374151; }
    .sp-welcome ol li { margin-bottom: 6px; line-height: 1.5; }
    .sp-welcome .actions-row { display: flex; gap: 10px; justify-content: flex-end; margin-top: 16px; }
    .sp-welcome button {
      font: inherit; padding: 10px 20px; border-radius: 6px; cursor: pointer;
      border: 1px solid #d1d5db; background: white; color: #374151;
      font-weight: 600; transition: all .15s;
    }
    .sp-welcome button:hover { background: #f3f4f6; }
    .sp-welcome button.primary {
      background: #0353a4; color: white; border-color: #0353a4;
      box-shadow: 0 2px 6px rgba(3,83,164,.3);
    }
    .sp-welcome button.primary:hover { background: #0466c8; box-shadow: 0 4px 10px rgba(3,83,164,.4); }

    @media (max-width: 720px){
      .sp-statusbar { padding: 8px 10px; flex-wrap: wrap; gap: 8px; }
      .sp-statusbar .subtxt { display: none; }
      .sp-user-menu-btn .nombre { display: none; }
    }
  `;

  function injectStyle(){
    if(document.getElementById('sp-statusbar-style')) return;
    const s = document.createElement('style');
    s.id = 'sp-statusbar-style';
    s.textContent = STYLE;
    document.head.appendChild(s);
  }

  function iniciales(s){
    if(!s) return '?';
    const partes = String(s).split(/[.\s@]/).filter(Boolean);
    return ((partes[0]?.[0]||'') + (partes[1]?.[0]||'')).toUpperCase() || 'U';
  }

  function render(state){
    let bar = document.getElementById('sp-statusbar');
    if(!bar){
      bar = document.createElement('div');
      bar.id = 'sp-statusbar';
      document.body.insertBefore(bar, document.body.firstChild);
    }
    bar.className = 'sp-statusbar ' + state.cls;
    bar.innerHTML = `
      <div class="left">
        <span class="dot"></span>
        ${state.icono ? `<span class="icono">${state.icono}</span>` : ''}
        <div class="info-block">
          <span class="titulo">${state.titulo || ''}</span>
          ${state.subtxt ? `<span class="subtxt">${state.subtxt}</span>` : ''}
        </div>
      </div>
      <div class="actions">${state.actions || ''}</div>
    `;
  }

  // ===== Modal de bienvenida (solo primera vez en modo local) =====
  function mostrarBienvenida(){
    if(localStorage.getItem('sp_welcome_visto') === '1') return;
    const overlay = document.createElement('div');
    overlay.className = 'sp-welcome-overlay';
    overlay.id = 'sp-welcome-overlay';
    overlay.innerHTML = `
      <div class="sp-welcome">
        <h2>👋 Bienvenido al ERP CONSITEC / Poderosa</h2>
        <p>Estás viendo el ERP en <b>modo local</b> — los datos solo existen en este navegador.</p>
        <p>Para acceder a la información compartida del equipo (clientes, cotizaciones, OCs, inventario), debes conectarte a <b>SharePoint con tu cuenta Microsoft 365</b>.</p>
        <ol>
          <li>Click en <b>"Conectar a SharePoint"</b> arriba a la derecha</li>
          <li>Acepta el cambio (recarga la página)</li>
          <li>Click en <b>"Iniciar sesión M365"</b></li>
          <li>Ingresa con tu cuenta <code>@consitecing.com</code></li>
        </ol>
        <div class="actions-row">
          <button onclick="SPUI._cerrarBienvenida(false)">Continuar en modo local</button>
          <button class="primary" onclick="SPUI._cerrarBienvenida(true)">🚀 Conectar ahora</button>
        </div>
      </div>
    `;
    document.body.appendChild(overlay);
  }
  function _cerrarBienvenida(activar){
    localStorage.setItem('sp_welcome_visto', '1');
    document.getElementById('sp-welcome-overlay')?.remove();
    if(activar) activarSharePoint();
  }

  // ===== Menú desplegable de usuario =====
  function toggleUserMenu(){
    document.querySelector('.sp-user-menu')?.classList.toggle('open');
  }
  // Cerrar menú al click fuera
  document.addEventListener('click', e => {
    const menu = document.querySelector('.sp-user-menu');
    if(menu && !menu.contains(e.target)) menu.classList.remove('open');
  });

  // ===== Acciones =====
  async function sincronizar(){
    if(!Auth || !Auth.isLoggedIn()){ alert('Inicia sesión primero'); return; }
    render({cls:'sp-load', titulo:'☁ Sincronizando con SharePoint...', subtxt:'Releyendo datos del servidor'});
    try {
      for(const [collKey, listName] of Object.entries(DB.SP_LISTAS || {})){
        try {
          const items = await SPAdapter.list(listName);
          DB[collKey].bulkSet(items);
        } catch(e){ console.warn('Sync fallo en', listName, e); }
      }
      window.dispatchEvent(new CustomEvent('db-ready', { detail: { synced: true }}));
      if(typeof window.refrescarUI === 'function') window.refrescarUI();
      pintarLogueado();
    } catch(e){
      render({cls:'sp-err', titulo:'Error de sincronización', subtxt: e.message || e,
        actions:`<button onclick="SPUI.sincronizar()">Reintentar</button>`});
    }
  }

  async function activarSharePoint(){
    if(!confirm('¿Activar modo SharePoint?\n\nLos datos se cargarán desde el servidor compartido del equipo.\n\nNecesitarás iniciar sesión con tu cuenta de Microsoft 365.')) return;
    SP_CONFIG.setModo('sharepoint');
    location.reload();
  }
  async function desactivarSharePoint(){
    if(!confirm('¿Volver a modo local?\n\nSolo verás los datos guardados en este navegador. No afecta los datos del equipo.')) return;
    SP_CONFIG.setModo('localStorage');
    location.reload();
  }
  async function iniciarSesion(){
    try {
      render({cls:'sp-load', titulo:'Iniciando sesión...', subtxt:'Abriendo ventana de Microsoft 365'});
      await Auth.login();
      render({cls:'sp-load', titulo:'☁ Cargando datos...', subtxt:'Sincronizando con SharePoint'});
      await DB.loadFromSharePoint();
      location.reload();
    } catch(e){
      console.error(e);
      render({
        cls:'sp-err',
        titulo:'Error al iniciar sesión',
        subtxt: e.message || String(e),
        actions:`<button onclick="location.reload()">Reintentar</button>
                 <button onclick="SP_CONFIG.setModo('localStorage'); location.reload()">Volver a modo local</button>`
      });
    }
  }
  async function cerrarSesion(){
    if(!confirm('¿Cerrar sesión de Microsoft 365?\n\nVolverás a la pantalla de login.')) return;
    await Auth.logout();
    location.reload();
  }

  // Acción reset bienvenida (debug útil)
  function mostrarBienvenidaForzar(){
    localStorage.removeItem('sp_welcome_visto');
    mostrarBienvenida();
  }

  // Render del estado logueado con avatar + menú
  function pintarLogueado(){
    const acc = Auth.getAccount();
    const email = (acc && (acc.username || acc.name)) || 'usuario';
    const nombre = (acc?.name || email.split('@')[0] || '').split(' ').slice(0,2).join(' ');
    const ini = iniciales(nombre || email);
    render({
      cls:'sp-on',
      titulo: `Conectado a SharePoint`,
      subtxt: `Datos del equipo · sincronizados en vivo`,
      actions: `
        <button onclick="SPUI.sincronizar()" title="Re-leer datos del servidor">🔄 Sincronizar</button>
        <div class="sp-user-menu">
          <div class="sp-user-menu-btn" onclick="SPUI.toggleUserMenu(event)">
            <div class="sp-avatar">${ini}</div>
            <span class="nombre">${nombre}</span>
            <span style="font-size:10px">▼</span>
          </div>
          <div class="sp-user-menu-dropdown">
            <div class="header">
              <div class="nombre">${nombre}</div>
              <div class="email">${email}</div>
            </div>
            <div class="item" onclick="SPUI.sincronizar()"><span class="ico">🔄</span> Sincronizar datos</div>
            <div class="divider"></div>
            <div class="item danger" onclick="SPUI.cerrarSesion()"><span class="ico">🚪</span> Cerrar sesión</div>
          </div>
        </div>
      `,
    });
  }

  // expone para los onclick
  window.SPUI = { activarSharePoint, desactivarSharePoint, iniciarSesion, cerrarSesion, sincronizar, toggleUserMenu, _cerrarBienvenida, mostrarBienvenidaForzar };

  async function init(){
    injectStyle();

    // Caso 1: modo localStorage
    if(!SP_CONFIG || SP_CONFIG.modo !== 'sharepoint'){
      render({
        cls:'local',
        titulo:`💾 Modo local · No estás conectado al equipo`,
        subtxt:`Los datos solo se guardan en este navegador. Conéctate a SharePoint para ver datos compartidos.`,
        actions:`<button class="primary" onclick="SPUI.activarSharePoint()">🔌 Conectar a SharePoint</button>`,
      });
      // Mostrar modal de bienvenida si es primera visita
      setTimeout(mostrarBienvenida, 600);
      return;
    }

    // Caso 2: scripts faltantes
    if(typeof Auth === 'undefined' || typeof SPAdapter === 'undefined'){
      render({
        cls:'sp-err',
        titulo:'⚠ Faltan scripts en este módulo',
        subtxt:'Auth/SPAdapter no cargados. Verifica el HTML.',
        actions:`<button onclick="SP_CONFIG.setModo('localStorage'); location.reload()">Volver a modo local</button>`,
      });
      return;
    }

    // Caso 3: init Auth
    try { await Auth.init(); } catch(e){
      render({
        cls:'sp-err',
        titulo:'Error iniciando Microsoft 365',
        subtxt: e.message || String(e),
        actions:`<button onclick="SP_CONFIG.setModo('localStorage'); location.reload()">Volver a modo local</button>`,
      });
      return;
    }

    // Caso 4: sin login
    if(!Auth.isLoggedIn()){
      render({
        cls:'sp-off',
        titulo:'🔐 No has iniciado sesión',
        subtxt:'Para ver los datos del equipo, inicia sesión con tu cuenta @consitecing.com',
        actions:`<button class="primary" onclick="SPUI.iniciarSesion()">🚀 Iniciar sesión M365</button>
                 <button onclick="SPUI.desactivarSharePoint()">Modo local</button>`,
      });
      return;
    }

    // Caso 5: logueado, cargar datos
    if(!DB.isReady()){
      const acc = Auth.getAccount();
      const email = (acc && (acc.username || acc.name)) || 'usuario';
      render({cls:'sp-load', titulo:`☁ Cargando datos del equipo...`, subtxt: `Conectado como ${email}`});
      try {
        await DB.loadFromSharePoint();
        window.dispatchEvent(new CustomEvent('db-ready', { detail: { synced: true }}));
        if(typeof window.refrescarUI === 'function') window.refrescarUI();
      } catch(e){
        render({
          cls:'sp-err',
          titulo:'No tienes permiso al sitio',
          subtxt: 'Pide al admin que te dé acceso al sitio SharePoint del ERP.',
          actions:`<button onclick="location.reload()">Reintentar</button>
                   <button onclick="SPUI.cerrarSesion()">Cerrar sesión</button>`,
        });
        return;
      }
    }
    pintarLogueado();
  }

  if(document.readyState === 'loading'){
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
