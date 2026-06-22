/* ============================================================
   CONFIGURACIÓN SHAREPOINT — ERP CONSITEC / Poderosa
   ============================================================
   Edita este archivo cuando tengas:
   - URL del sitio SharePoint
   - Client ID y Tenant ID de la app registrada en Azure AD

   Por defecto el ERP funciona en modo 'localStorage' (sin SharePoint).
   Cuando quieras activar SharePoint, cambia modo: 'sharepoint'.
============================================================ */

const SP_CONFIG = {
  // ============== MODO DE OPERACIÓN ==============
  // 'localStorage' = datos solo en el navegador del usuario
  // 'sharepoint'   = datos compartidos en SharePoint Lists
  // El usuario puede cambiarlo desde la UI (botón "Conectar a SharePoint") y se persiste en localStorage.
  modo: (typeof localStorage !== 'undefined' && localStorage.getItem('sp_modo')) || 'localStorage',

  // ============== SHAREPOINT ==============
  siteUrl: 'https://consitecingenierossac.sharepoint.com/sites/erp-consitec-poderosa',

  // ============== AZURE AD (app "ERP CONSITEC Poderosa") ==============
  tenantId: '5b2acc65-31fe-49c9-aa4e-8066ac119e65',
  clientId: '66443139-9fb4-48a5-b68a-3e035af97fa9',

  // ============== MAPEO DE LISTAS ==============
  // Las claves DEBEN coincidir con los nombres de DB.X en storage.js
  // Los valores son los nombres exactos de las listas en SharePoint
  listas: {
    clientes:          'Clientes',
    proveedores:       'Proveedores',
    cargos:            'Cargos',
    personal:          'Personal',
    contactosConsitec: 'Gestion',
    materiales:        'Materiales',
    cotizaciones:      'Cotizaciones',
    ocs:               'OrdenesCompra',
    proyectos:         'Proyectos',
    // Reportes diario/semanal en una sola lista con campo 'tipo'
    reportesDiarios:   'Reportes',
    reportesSemanales: 'Reportes',
    // Inventario de unidades mineras (Vijus, Chaparrosa, etc.) — 1 lista, filtro por UnidadMinera
    inventario:        'Inventario',
    // Movimientos del inventario: entradas, salidas, ajustes (kardex con historial)
    movimientosInventario: 'MovimientosInventario',
  },

  // ============== SCOPES PARA AUTENTICACIÓN ==============
  scopes: [
    'User.Read',
    'Sites.ReadWrite.All',
  ],

  // ============== CACHE ==============
  // Si true, cachea respuestas en memoria por 60 segundos
  // (mejora performance al cambiar de pestañas)
  cache: true,
  cacheTTL: 60_000, // ms
};

// Helper para que la UI pueda cambiar el modo en runtime y recargar la página
SP_CONFIG.setModo = function(modo){
  if(modo !== 'localStorage' && modo !== 'sharepoint') return;
  localStorage.setItem('sp_modo', modo);
  this.modo = modo;
};

// Exponer en window para acceso global
if(typeof window !== 'undefined') window.SP_CONFIG = SP_CONFIG;
