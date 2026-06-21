/* ============================================================
   AUTH-UI — Barra superior de estado de conexión + acciones
   ============================================================
   Requiere: SP_CONFIG, Auth, SPAdapter, DB
   Se auto-inicializa en DOMContentLoaded.
   - Modo localStorage → barra gris con botón "Conectar a SharePoint"
   - Modo sharepoint sin login → barra naranja con botón "Iniciar sesión M365"
   - Modo sharepoint con login → barra verde con email y "Cerrar sesión"
============================================================ */
(function(){
  const STYLE = `
    .sp-statusbar {
      position: sticky; top: 0; z-index: 1000;
      display: flex; align-items: center; justify-content: space-between;
      gap: 12px; padding: 8px 16px;
      font: 13px system-ui, -apple-system, Segoe UI, Arial, sans-serif;
      border-bottom: 1px solid rgba(0,0,0,.08);
    }
    .sp-statusbar.local   { background: #f3f4f6; color: #374151; }
    .sp-statusbar.sp-off  { background: #fff7ed; color: #9a3412; }
    .sp-statusbar.sp-on   { background: #ecfdf5; color: #065f46; }
    .sp-statusbar.sp-load { background: #eff6ff; color: #1e40af; }
    .sp-statusbar.sp-err  { background: #fef2f2; color: #991b1b; }
    .sp-statusbar .left   { display: flex; align-items: center; gap: 10px; }
    .sp-statusbar .dot    { width: 8px; height: 8px; border-radius: 50%; display:inline-block; }
    .sp-statusbar.local   .dot { background: #9ca3af; }
    .sp-statusbar.sp-off  .dot { background: #f97316; }
    .sp-statusbar.sp-on   .dot { background: #10b981; }
    .sp-statusbar.sp-load .dot { background: #3b82f6; animation: sp-pulse 1s ease-in-out infinite; }
    .sp-statusbar.sp-err  .dot { background: #ef4444; }
    @keyframes sp-pulse { 0%,100%{opacity:1} 50%{opacity:.3} }
    .sp-statusbar .actions { display: flex; gap: 6px; }
    .sp-statusbar button {
      font: inherit; padding: 4px 10px; border-radius: 4px;
      border: 1px solid currentColor; background: transparent; color: inherit;
      cursor: pointer;
    }
    .sp-statusbar button:hover { background: rgba(0,0,0,.05); }
    .sp-statusbar button.primary {
      background: currentColor; color: white;
    }
    .sp-statusbar button.primary:hover { opacity: .9; }
  `;

  function injectStyle(){
    if(document.getElementById('sp-statusbar-style')) return;
    const s = document.createElement('style');
    s.id = 'sp-statusbar-style';
    s.textContent = STYLE;
    document.head.appendChild(s);
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
        <span>${state.text}</span>
      </div>
      <div class="actions">${state.actions || ''}</div>
    `;
  }

  async function sincronizar(){
    if(!Auth || !Auth.isLoggedIn()){ alert('Inicia sesión primero'); return; }
    render({cls:'sp-load', text:'☁ Sincronizando con SharePoint...'});
    try {
      // forzar reload del cache
      if(DB.SP_LISTAS){
        // marcar spReady=false para que loadFromSharePoint vuelva a correr
        // (uso un truco: recargar la página después de re-cachear todo)
      }
      // Re-leer cada lista
      for(const [collKey, listName] of Object.entries(DB.SP_LISTAS || {})){
        try {
          const items = await SPAdapter.list(listName);
          // empujar al storage para que actualice cache+mirror
          DB[collKey].bulkSet(items);
        } catch(e){ console.warn('Sync fallo en', listName, e); }
      }
      window.dispatchEvent(new CustomEvent('db-ready', { detail: { synced: true }}));
      const acc = Auth.getAccount();
      const email = (acc && (acc.username || acc.name)) || 'usuario';
      render({
        cls:'sp-on',
        text:`☁ SharePoint · ${email}`,
        actions:`<button onclick="SPUI.sincronizar()" title="Re-leer datos del servidor">🔄 Sincronizar</button>
                 <button onclick="SPUI.cerrarSesion()">Cerrar sesión</button>`,
      });
      if(typeof window.refrescarUI === 'function') window.refrescarUI();
    } catch(e){
      render({cls:'sp-err', text:'Error sync: '+(e.message||e),
        actions:`<button onclick="SPUI.sincronizar()">Reintentar</button>`});
    }
  }

  async function activarSharePoint(){
    if(!confirm('¿Activar modo SharePoint?\n\nLos datos se cargarán desde el servidor compartido.\n\nNecesitarás iniciar sesión con tu cuenta de Microsoft 365.')) return;
    SP_CONFIG.setModo('sharepoint');
    location.reload();
  }
  async function desactivarSharePoint(){
    if(!confirm('¿Volver a modo local (solo este navegador)?\n\nNo perderás datos. El último estado quedó en cache local.')) return;
    SP_CONFIG.setModo('localStorage');
    location.reload();
  }
  async function iniciarSesion(){
    try {
      render({cls:'sp-load', text:'Iniciando sesión...'});
      await Auth.login();
      render({cls:'sp-load', text:'Sincronizando con SharePoint...'});
      await DB.loadFromSharePoint();
      // Recargar la página para que los módulos repinten con los datos de SP
      location.reload();
    } catch(e){
      console.error(e);
      render({
        cls:'sp-err',
        text:'Error: ' + (e.message || e),
        actions:`<button onclick="location.reload()">Reintentar</button>
                 <button onclick="SP_CONFIG.setModo('localStorage'); location.reload()">Volver a modo local</button>`
      });
    }
  }
  async function cerrarSesion(){
    if(!confirm('¿Cerrar sesión de M365?')) return;
    await Auth.logout();
    location.reload();
  }

  // expone para los onclick
  window.SPUI = { activarSharePoint, desactivarSharePoint, iniciarSesion, cerrarSesion, sincronizar };

  async function init(){
    injectStyle();

    // Caso 1: modo localStorage
    if(!SP_CONFIG || SP_CONFIG.modo !== 'sharepoint'){
      render({
        cls:'local',
        text:'💾 Modo local · Datos solo en este navegador',
        actions:`<button class="primary" onclick="SPUI.activarSharePoint()">Conectar a SharePoint</button>`,
      });
      return;
    }

    // Caso 2: modo SharePoint pero MSAL/Auth no están cargados
    if(typeof Auth === 'undefined' || typeof SPAdapter === 'undefined'){
      render({
        cls:'sp-err',
        text:'⚠ Modo SharePoint activo pero faltan scripts (Auth/SPAdapter). Revisar HTML.',
        actions:`<button onclick="SP_CONFIG.setModo('localStorage'); location.reload()">Volver a modo local</button>`,
      });
      return;
    }

    // Caso 3: modo SP, esperar a que Auth esté listo
    try {
      await Auth.init();
    } catch(e){
      render({
        cls:'sp-err',
        text:'Error iniciando Auth: ' + (e.message || e),
        actions:`<button onclick="SP_CONFIG.setModo('localStorage'); location.reload()">Volver a modo local</button>`,
      });
      return;
    }

    // Caso 4: modo SP, sin login
    if(!Auth.isLoggedIn()){
      render({
        cls:'sp-off',
        text:'🔐 Modo SharePoint · Sesión cerrada',
        actions:`<button class="primary" onclick="SPUI.iniciarSesion()">Iniciar sesión M365</button>
                 <button onclick="SPUI.desactivarSharePoint()">Modo local</button>`,
      });
      return;
    }

    // Caso 5: modo SP, logueado → cargar datos
    const acc = Auth.getAccount();
    const email = (acc && (acc.username || acc.name)) || 'usuario';
    if(!DB.isReady()){
      render({cls:'sp-load', text:`☁ Sincronizando datos para ${email}...`});
      try {
        await DB.loadFromSharePoint();
        // Notificar a la página que los datos están listos
        window.dispatchEvent(new CustomEvent('db-ready', { detail: { synced: true }}));
        if(typeof window.refrescarUI === 'function') window.refrescarUI();
      } catch(e){
        render({
          cls:'sp-err',
          text:'Error cargando SharePoint: ' + (e.message || e),
          actions:`<button onclick="location.reload()">Reintentar</button>
                   <button onclick="SPUI.cerrarSesion()">Cerrar sesión</button>`,
        });
        return;
      }
    }
    render({
      cls:'sp-on',
      text:`☁ SharePoint · ${email}`,
      actions:`<button onclick="SPUI.sincronizar()" title="Re-leer datos del servidor">🔄 Sincronizar</button>
               <button onclick="SPUI.cerrarSesion()">Cerrar sesión</button>`,
    });
  }

  if(document.readyState === 'loading'){
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }
})();
