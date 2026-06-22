/* ============================================================
   SHAREPOINT ADAPTER — CRUD sobre SharePoint Lists vía Graph API
   ============================================================
   Requiere: Auth (auth.js) + SP_CONFIG (sharepoint-config.js)

   Uso:
     await SPAdapter.init();
     const items = await SPAdapter.list('Clientes');
     await SPAdapter.save('Clientes', {razonSocial:'Foo', ruc:'123'});
     await SPAdapter.remove('Clientes', spId);
============================================================ */
const SPAdapter = (function(){
  const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
  let siteId = null;
  let listIds = {};   // { 'Clientes': '<guid>', ... }
  const cache = {};   // cache simple por lista
  let initialized = false;

  // Campos del item SP que NO debemos enviar al hacer update/insert
  // Incluye tanto los nombres PascalCase originales de SP como los camelCase post-deserialize.
  const SP_INTERNAL_FIELDS = new Set([
    'spId','id','createdAt','updatedAt','@odata.etag',
    'Title','LinkTitle','LinkTitleNoMenu',
    'Edit','_UIVersionString','Attachments','ContentType',
    'Modified','Created','Author','Editor','_HasCopyDestinations',
    '_CopySource','owshiddenversion','WorkflowVersion','_UIVersion',
    'WorkflowInstanceID','FileRef','FileDirRef','Last_x0020_Modified',
    'Created_x0020_Date','FSObjType','SortBehavior','PermMask',
    'FileLeafRef','UniqueId','SyncClientId','ProgId','ScopeId',
    'MetaInfo','_Level','_IsCurrentVersion','ItemChildCount',
    'FolderChildCount','Restricted','OriginatorId','NoExecute',
    'ContentVersion','_ComplianceFlags','_ComplianceTag',
    '_ComplianceTagWrittenTime','_ComplianceTagUserId',
    'AccessPolicy','AppAuthor','AppEditor','SMTotalSize',
    'SMLastModifiedDate','SMTotalFileStreamSize','SMTotalFileCount',
    'ParentVersionString','ParentLeafName',
    // camelCase tras deserialize — read-only en SP
    'attachments','contentType','contentTypeId',
    'authorLookupId','editorLookupId','appAuthorLookupId','appEditorLookupId',
    'attachmentsLookupId','linkTitle','linkTitleNoMenu',
    'fileLeafRef','fileRef','fileDirRef',
  ]);

  // Campos cuyo valor es array/objeto y se almacena como JSON-string
  const JSON_FIELDS = new Set([
    'contactos','areas','items','pagos','tareas','baseline','calendario',
    'mo1','mov1','eq1','mo2','mov2','eq2','mat','notas','excl','resp','fichas',
    'totales','actividades','personalPresente','fotos',
  ]);

  // Mapeo del campo JS principal de cada lista → SharePoint "Title" (única columna obligatoria por defecto)
  // Si el objeto contiene este campo, su valor va a `Title` y NO se envía como columna separada.
  const TITLE_MAP = {
    'Clientes':      'razonSocial',
    'Proveedores':   'razonSocial',
    'Cargos':        'cargo',
    'Personal':      'nombreCompleto',   // se compone de nombres + apellidos si no viene
    'Gestion':       'nombreCompleto',
    'Materiales':    'descripcion',
    'Cotizaciones':  'codigo',
    'OrdenesCompra': 'codigo',
    'Proyectos':     'codigo',
    'Reportes':      'codigo',
    'Inventario':    'codigo',     // código del item (MNIVE1, MMACH1, etc) va al Title
  };

  // Renombres de columnas: JS key → SharePoint internal name (cuando NO sigue la regla camel→Pascal)
  // Estructurado por lista, porque la misma JS-key puede mapear a distinta columna en cada lista.
  // '_global' aplica a todas; el resto solo a la lista nombrada.
  const COLUMN_RENAME = {
    _global: {
      'logo':          'LogoBase64',
    },
    'Personal': {
      'nombre':        'Nombres',     // JS usa 'nombre' singular, SP tiene 'Nombres' plural
      'cargoId':       'CargoIdSP',
    },
    'Gestion': {
      'nombre':        'Nombres',     // mismo caso
    },
    'Cotizaciones': {
      'clienteId':     'ClienteIdSP',
      'proyectoId':    'ProyectoIdSP',
      'cargo':         'CargoContacto',  // JS usa 'cargo' (singular del contacto), SP es 'CargoContacto'
      'notas':         'NotasJson',
      'mo1':           'Mo1Json',
      'mov1':          'Mov1Json',
      'eq1':           'Eq1Json',
      'mo2':           'Mo2Json',
      'mov2':          'Mov2Json',
      'eq2':           'Eq2Json',
      'mat':           'MatJson',
      'excl':          'ExclJson',
      'resp':          'RespJson',
      'fichas':        'FichasJson',
    },
    'OrdenesCompra': {
      'gestorId':      'GestorIdSP',
      'solicitanteId': 'SolicitanteIdSP',
      'aprobadorId':   'AprobadorIdSP',
      'proveedorId':   'ProveedorIdSP',
      'proyectoId':    'ProyectoIdSP',
      'desc':          'Descuento',   // JS: 'desc', SP: 'Descuento'
      'items':         'ItemsJson',
      'pagos':         'PagosJson',
    },
    'Proyectos': {
      'clienteId':     'ClienteIdSP',
      'cotizacionId':  'CotizacionIdSP',
      'tareas':        'TareasJson',
      'baseline':      'BaselineJson',
      'calendario':    'CalendarioJson',
    },
    'Reportes': {
      'proyectoId':      'ProyectoIdSP',
      'responsableId':   'ResponsableIdSP',
      'actividades':     'ActividadesJson',
      'personalPresente':'PersonalJson',
      'fotos':           'FotosJson',
    },
  };

  // Campos que el adapter NUNCA envía a SP (porque no tienen columna o son derivados)
  // Pueden tener subcampos que sí van como columnas separadas (ver expansión más abajo).
  const SKIP_FIELDS_BY_LIST = {
    // Cotizaciones: campos derivados o sin columna en SP (se mantienen en localStorage mirror)
    'Cotizaciones': new Set([
      'totales',           // totalFinal y ganancia van como columnas separadas
      'lugarEntrega',      // campo de cotizador suministro sin columna SP
      'plazoEntrega',      // idem
      'codigoCliente',     // cliente snapshot
      'rucCliente',        // cliente snapshot
    ]),
    'Personal': new Set(['cargo','precioHH','costoRealHH','ganancia','categoria']),
    'OrdenesCompra': new Set(['firmaId','firmaNombre','firmaCargo','proyectoCliente']),
  };

  // Expansión de objetos anidados → columnas top-level SP
  // Ej: en Cotizaciones, item.totales = {totalFinal, ganancia} se expande a TotalFinal, Ganancia
  const EXPAND_NESTED = {
    'Cotizaciones': {
      'totales': {
        'totalFinal': 'TotalFinal',
        'ganancia':   'Ganancia',
      },
    },
  };

  // JS key → SP column, según lista
  function renameToSP(listaNombre, jsKey){
    const listMap = COLUMN_RENAME[listaNombre];
    if(listMap && listMap[jsKey]) return listMap[jsKey];
    if(COLUMN_RENAME._global[jsKey]) return COLUMN_RENAME._global[jsKey];
    return null;
  }
  // SP column → JS key, según lista
  function renameToJS(listaNombre, spKey){
    const listMap = COLUMN_RENAME[listaNombre];
    if(listMap){
      for(const [js, sp] of Object.entries(listMap)) if(sp === spKey) return js;
    }
    for(const [js, sp] of Object.entries(COLUMN_RENAME._global)) if(sp === spKey) return js;
    return null;
  }

  async function init(){
    if(initialized) return;
    if(!SP_CONFIG || !SP_CONFIG.clientId || SP_CONFIG.clientId.startsWith('PEGAR_')){
      throw new Error('SP_CONFIG no está configurado. Edita js/sharepoint-config.js');
    }
    await Auth.init();
    initialized = true;
  }

  // Helper: llamadas a Microsoft Graph con auth automático
  async function graphFetch(path, options = {}){
    const token = await Auth.getToken();
    const url = path.startsWith('http') ? path : GRAPH_BASE + path;
    const res = await fetch(url, {
      ...options,
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'application/json',
        'Accept': 'application/json',
        ...(options.headers || {})
      }
    });
    if(!res.ok){
      let detail = '';
      try { detail = JSON.stringify(await res.json()); } catch(e){ detail = await res.text(); }
      throw new Error(`Graph ${res.status} ${res.statusText}: ${detail.slice(0,500)}`);
    }
    if(res.status === 204) return null;
    return res.json();
  }

  // Site ID
  async function getSiteId(){
    if(siteId) return siteId;
    const url = new URL(SP_CONFIG.siteUrl);
    const hostname = url.hostname;
    const path = url.pathname; // /sites/erp-consitec-poderosa
    const r = await graphFetch(`/sites/${hostname}:${path}`);
    siteId = r.id;
    return siteId;
  }

  // List ID
  async function getListId(listaNombre){
    if(listIds[listaNombre]) return listIds[listaNombre];
    const sid = await getSiteId();
    const r = await graphFetch(`/sites/${sid}/lists?$filter=displayName eq '${listaNombre}'`);
    if(!r.value || r.value.length === 0){
      throw new Error(`Lista '${listaNombre}' no existe en SharePoint`);
    }
    listIds[listaNombre] = r.value[0].id;
    return listIds[listaNombre];
  }

  // Pre-cargar todos los IDs de listas
  async function loadAllListIds(){
    const sid = await getSiteId();
    const r = await graphFetch(`/sites/${sid}/lists?$select=id,displayName`);
    (r.value || []).forEach(l => { listIds[l.displayName] = l.id; });
    return listIds;
  }

  // Genera el "campo Title" de la lista a partir del objeto JS
  function computeTitle(listaNombre, obj){
    const titleField = TITLE_MAP[listaNombre];
    if(titleField){
      // Caso especial: nombreCompleto = (nombre|nombres) + apellidos — JS usa 'nombre' singular
      if(titleField === 'nombreCompleto'){
        const full = [obj.nombre || obj.nombres, obj.apellidos].filter(Boolean).join(' ').trim();
        if(full) return full;
      } else if(obj[titleField]){
        return obj[titleField];
      }
    }
    return obj.razonSocial || obj.codigo || obj.nombre || obj.cargo || obj.descripcion || obj.titulo || '—';
  }

  // Inverso: dado el Title leído y la lista, devuelve los campos JS que deben rellenarse
  function expandTitle(listaNombre, titleValue){
    const titleField = TITLE_MAP[listaNombre];
    if(!titleField || !titleValue) return {};
    if(titleField === 'nombreCompleto'){
      const parts = String(titleValue).split(' ');
      // Devuelvo 'nombre' singular (lo que el JS espera). Los apellidos van por la columna Apellidos.
      return {
        nombre: parts[0] || '',
        apellidos: parts.slice(1).join(' ') || '',
      };
    }
    return { [titleField]: titleValue };
  }

  // ===== Serialización JS object → SharePoint fields =====
  function serialize(obj, listaNombre){
    const fields = {};
    const titleField = TITLE_MAP[listaNombre];
    const skipSet = SKIP_FIELDS_BY_LIST[listaNombre];
    const expandMap = EXPAND_NESTED[listaNombre] || {};

    Object.entries(obj || {}).forEach(([key, value]) => {
      if(SP_INTERNAL_FIELDS.has(key)) return;
      // Filtro genérico para cualquier campo de lookup que llegue tras deserializar
      // (Author, Editor, ContentType, etc.) — son read-only en SP.
      if(key.endsWith('LookupId')) return;
      if(value === undefined) return;
      // El UUID local 'id' viaja a SP como 'LocalId' para preservar referencias entre usuarios
      if(key === 'id'){
        if(value && typeof value === 'string' && value.startsWith('id_')){
          fields.LocalId = value;
        }
        return;
      }
      // Si este campo va al Title, NO lo enviamos como columna separada
      if(titleField && key === titleField) return;
      // Expansión de subcampos anidados (ej: totales → TotalFinal + Ganancia)
      if(expandMap[key] && value && typeof value === 'object'){
        Object.entries(expandMap[key]).forEach(([subKey, spCol]) => {
          if(value[subKey] !== undefined) fields[spCol] = value[subKey];
        });
        return; // no enviar el objeto entero
      }
      // Campos que el adapter explícitamente omite
      if(skipSet && skipSet.has(key)) return;
      // Renombre explícito (por lista o global) o regla camel→Pascal
      const spKey = renameToSP(listaNombre, key) || (key.charAt(0).toUpperCase() + key.slice(1));
      // SP rechaza string vacío "" en columnas DateTime. Si la key es una fecha y value es '', enviar null.
      // Si tiene formato YYYY-MM-DD sin hora, convertir a ISO completo.
      const isDateKey = /fecha/i.test(key) || key === 'createdAt' || key === 'updatedAt';
      if(isDateKey && typeof value === 'string'){
        if(value === '') value = null;
        else if(/^\d{4}-\d{2}-\d{2}$/.test(value)) value = value + 'T00:00:00Z';
      }
      if(value !== null && typeof value === 'object' && !(value instanceof Date)){
        fields[spKey] = JSON.stringify(value);
      } else if(typeof value === 'boolean'){
        fields[spKey] = value;
      } else if(value === null){
        fields[spKey] = null;
      } else {
        fields[spKey] = value;
      }
    });
    // Title obligatorio
    if(!fields.Title){
      fields.Title = computeTitle(listaNombre, obj);
    }
    if(fields.Title && fields.Title.length > 255){
      fields.Title = fields.Title.slice(0, 252) + '...';
    }
    return fields;
  }

  // ===== Deserialización SharePoint item → JS object =====
  function deserialize(spItem, listaNombre){
    const fields = spItem.fields || {};
    // El id de la app es LocalId (UUID local) si existe; si no, el id interno de SP
    const obj = {
      spId: spItem.id,
      id: fields.LocalId || spItem.id,
      createdAt: spItem.createdDateTime ? new Date(spItem.createdDateTime).getTime() : Date.now(),
      updatedAt: spItem.lastModifiedDateTime ? new Date(spItem.lastModifiedDateTime).getTime() : Date.now(),
    };
    // Mapa inverso de EXPAND_NESTED para esta lista: SPcol → {parentKey, subKey}
    const expandMap = EXPAND_NESTED[listaNombre] || {};
    const reverseExpand = {};
    Object.entries(expandMap).forEach(([parentKey, subMap]) => {
      Object.entries(subMap).forEach(([subKey, spCol]) => {
        reverseExpand[spCol] = { parentKey, subKey };
      });
    });

    Object.entries(fields).forEach(([key, value]) => {
      if(SP_INTERNAL_FIELDS.has(key) && key !== 'Title') return;
      if(key.startsWith('_') || key.startsWith('@')) return;
      // LocalId ya se procesó arriba
      if(key === 'LocalId') return;
      // Title → expandir al campo JS correspondiente de la lista
      if(key === 'Title'){
        Object.assign(obj, expandTitle(listaNombre, value));
        return;
      }
      // Reagrupar columnas tipo TotalFinal/Ganancia → obj.totales.{totalFinal, ganancia}
      if(reverseExpand[key]){
        const { parentKey, subKey } = reverseExpand[key];
        obj[parentKey] = obj[parentKey] || {};
        obj[parentKey][subKey] = value;
        return;
      }
      // Renombre inverso (por lista o global), o regla camelCase
      const jsKey = renameToJS(listaNombre, key) || (key.charAt(0).toLowerCase() + key.slice(1));
      // Si parece JSON (array/objeto), deserializar
      if(typeof value === 'string' && (value.startsWith('[') || value.startsWith('{'))){
        try { value = JSON.parse(value); } catch(e){ /* dejar como string */ }
      }
      obj[jsKey] = value;
    });
    return obj;
  }

  // ===== CRUD =====
  async function list(listaNombre, options = {}){
    await init();
    const sid = await getSiteId();
    const lid = await getListId(listaNombre);
    const r = await graphFetch(`/sites/${sid}/lists/${lid}/items?expand=fields&$top=999`);
    const items = (r.value || []).map(it => deserialize(it, listaNombre));
    cache[listaNombre] = items;
    return items;
  }

  async function get(listaNombre, spId){
    await init();
    const sid = await getSiteId();
    const lid = await getListId(listaNombre);
    const r = await graphFetch(`/sites/${sid}/lists/${lid}/items/${spId}?expand=fields`);
    return deserialize(r, listaNombre);
  }

  async function save(listaNombre, item){
    await init();
    const sid = await getSiteId();
    const lid = await getListId(listaNombre);
    const fields = serialize(item, listaNombre);
    if(item.spId){
      // UPDATE
      await graphFetch(`/sites/${sid}/lists/${lid}/items/${item.spId}/fields`, {
        method: 'PATCH',
        body: JSON.stringify(fields),
      });
      return { ...item };
    } else {
      // INSERT
      const r = await graphFetch(`/sites/${sid}/lists/${lid}/items`, {
        method: 'POST',
        body: JSON.stringify({ fields }),
      });
      return { ...item, spId: r.id, id: r.id };
    }
  }

  async function remove(listaNombre, spId){
    await init();
    const sid = await getSiteId();
    const lid = await getListId(listaNombre);
    await graphFetch(`/sites/${sid}/lists/${lid}/items/${spId}`, { method: 'DELETE' });
  }

  // Verificación rápida de conexión
  async function ping(){
    await init();
    if(!Auth.isLoggedIn()) return { ok: false, error: 'No logueado' };
    try {
      const sid = await getSiteId();
      const ids = await loadAllListIds();
      return {
        ok: true,
        siteId: sid,
        listas: Object.keys(ids),
        usuario: Auth.getAccount(),
      };
    } catch(e){
      return { ok: false, error: e.message };
    }
  }

  return {
    init, list, get, save, remove,
    getSiteId, getListId, loadAllListIds,
    serialize, deserialize, graphFetch, ping,
  };
})();

if(typeof window !== 'undefined') window.SPAdapter = SPAdapter;
