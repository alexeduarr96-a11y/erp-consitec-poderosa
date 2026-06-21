/* ============================================================
   AUTH — Login con Microsoft 365 usando MSAL.js
   ============================================================
   Requiere que SP_CONFIG esté cargado (de sharepoint-config.js)
   Requiere que msal-browser.min.js esté cargado (de CDN o local)
============================================================ */
const Auth = (function(){
  let msalInstance = null;
  let account = null;
  let listeners = [];

  async function init(){
    if(msalInstance) return msalInstance;
    if(typeof msal === 'undefined'){
      throw new Error('MSAL.js no está cargado. Agrega el <script> de msal-browser en el HTML.');
    }
    msalInstance = new msal.PublicClientApplication({
      auth: {
        clientId: SP_CONFIG.clientId,
        authority: 'https://login.microsoftonline.com/' + SP_CONFIG.tenantId,
        // Usamos solo origin (sin path) para que coincida con el URI registrado en Azure
        redirectUri: window.location.origin,
      },
      cache: {
        // localStorage para que la sesión persista entre pestañas/módulos del ERP
        cacheLocation: 'localStorage',
        storeAuthStateInCookie: false,
      }
    });
    await msalInstance.initialize();
    // Restaurar sesión previa si existe
    const accounts = msalInstance.getAllAccounts();
    if(accounts.length > 0){
      account = accounts[0];
      notifyListeners();
    }
    return msalInstance;
  }

  async function login(){
    if(!msalInstance) await init();
    try {
      const result = await msalInstance.loginPopup({
        scopes: SP_CONFIG.scopes || ['User.Read','Sites.ReadWrite.All'],
        prompt: 'select_account',
      });
      account = result.account;
      msalInstance.setActiveAccount(account);
      notifyListeners();
      return account;
    } catch(e){
      console.error('Login error:', e);
      throw e;
    }
  }

  async function logout(){
    if(!msalInstance) return;
    try {
      await msalInstance.logoutPopup({ account });
    } catch(e){ console.warn(e); }
    account = null;
    notifyListeners();
  }

  async function getToken(scopes){
    if(!msalInstance) await init();
    if(!account){
      const accounts = msalInstance.getAllAccounts();
      if(accounts.length > 0) account = accounts[0];
      else throw new Error('No hay sesión activa. Inicia sesión primero.');
    }
    const request = {
      scopes: scopes || ['Sites.ReadWrite.All'],
      account: account,
    };
    try {
      const result = await msalInstance.acquireTokenSilent(request);
      return result.accessToken;
    } catch(e){
      if(e instanceof msal.InteractionRequiredAuthError){
        const result = await msalInstance.acquireTokenPopup(request);
        return result.accessToken;
      }
      throw e;
    }
  }

  function getAccount(){ return account; }
  function isLoggedIn(){ return account !== null; }

  function onChange(cb){
    listeners.push(cb);
    cb(account);
    return () => { listeners = listeners.filter(x => x !== cb); };
  }
  function notifyListeners(){
    listeners.forEach(cb => { try { cb(account); } catch(e){ console.error(e); } });
  }

  return { init, login, logout, getToken, getAccount, isLoggedIn, onChange };
})();

if(typeof window !== 'undefined') window.Auth = Auth;
