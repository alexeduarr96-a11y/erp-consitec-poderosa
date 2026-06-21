/* ===== ERP CONSITEC - Persistencia centralizada =====
   Modo dual:
     - localStorage (por defecto): todo se guarda en el navegador
     - sharepoint: cache en memoria + writeback async a SharePoint Lists
   El API es 100% compatible — los módulos no necesitan saber qué modo está activo.
   Toggle en js/sharepoint-config.js → SP_CONFIG.modo
============================================================ */
const DB = (() => {
  const PREFIX = 'consitec_erp_';
  const KEYS = {
    clientes:    PREFIX + 'clientes',
    proyectos:   PREFIX + 'proyectos',
    cotizaciones:PREFIX + 'cotizaciones',
    ocs:         PREFIX + 'ocs',
    cargos:      PREFIX + 'cargos',
    personal:    PREFIX + 'personal',
    materiales:  PREFIX + 'materiales',
    contactosConsitec: PREFIX + 'contactos_consitec',
    proveedores:       PREFIX + 'proveedores',
    asignaciones:      PREFIX + 'asignaciones',
    reportesDiarios:   PREFIX + 'reportes_diarios',
    catMovilizacion:   PREFIX + 'cat_movilizacion',
    catEquipos:        PREFIX + 'cat_equipos',
    catNotas:          PREFIX + 'cat_notas',
    catExclusiones:    PREFIX + 'cat_exclusiones',
    catResponsabilidades: PREFIX + 'cat_responsabilidades',
    inventario:        PREFIX + 'inventario',     // inventario de unidades mineras
    config:      PREFIX + 'config',
    counters:    PREFIX + 'counters',
  };

  // Mapeo colección local → lista SharePoint (las que se sincronizan)
  // Las colecciones que NO están aquí (catálogos pequeños, config, counters) se quedan en localStorage.
  const SP_LISTAS = {
    clientes:          'Clientes',
    proveedores:       'Proveedores',
    cargos:            'Cargos',
    personal:          'Personal',
    contactosConsitec: 'Gestion',
    materiales:        'Materiales',
    cotizaciones:      'Cotizaciones',
    ocs:               'OrdenesCompra',
    proyectos:         'Proyectos',
    reportesDiarios:   'Reportes',
    inventario:        'Inventario',
  };

  // ¿Está el modo SharePoint activo y disponible?
  function isSPMode(){
    return typeof SP_CONFIG !== 'undefined'
        && SP_CONFIG.modo === 'sharepoint'
        && typeof SPAdapter !== 'undefined'
        && typeof Auth !== 'undefined'
        && Auth.isLoggedIn();
  }

  // Cache en memoria por colección (espejo de SP). Solo se usa en modo sharepoint.
  const spCache = {};
  let spReady = false;
  let spReadyPromise = null;

  function read(key, fallback){
    try { const raw = localStorage.getItem(key); return raw ? JSON.parse(raw) : fallback; }
    catch(e){ console.error('DB read error', key, e); return fallback; }
  }
  function write(key, data){
    try { localStorage.setItem(key, JSON.stringify(data)); return true; }
    catch(e){ console.error('DB write error', key, e); return false; }
  }
  function uid(prefix='id'){
    return prefix + '_' + Date.now().toString(36) + Math.random().toString(36).slice(2,7);
  }

  // Reporta error sin romper la UI
  function reportSPError(op, listName, err){
    console.error(`[SP ${op}] ${listName}:`, err);
    if(typeof toast === 'function'){
      toast(`SharePoint ${op} falló: ${err.message || err}`, 'err');
    }
  }

  // ===== Pre-carga (espejo SP → cache) =====
  // Debe llamarse después del login con M365. Carga TODAS las listas mapeadas.
  async function loadFromSharePoint(){
    if(!isSPMode()) throw new Error('No estás en modo SharePoint o sin login.');
    if(spReadyPromise) return spReadyPromise;
    spReadyPromise = (async () => {
      const entries = Object.entries(SP_LISTAS);
      for(const [collKey, listName] of entries){
        try {
          const items = await SPAdapter.list(listName);
          spCache[collKey] = items;
          // mirror a localStorage como backup
          write(KEYS[collKey], items);
        } catch(e){
          reportSPError('LIST', listName, e);
          spCache[collKey] = read(KEYS[collKey], []);
        }
      }
      spReady = true;
      return spCache;
    })();
    return spReadyPromise;
  }
  function isReady(){ return !isSPMode() || spReady; }

  // ===== CONFIG (datos empresa) =====
  function getConfig(){
    return read(KEYS.config, {
      empresa: 'CONSITEC INGENIEROS S.A.C.',
      slogan:  'CONTROL, SISTEMAS Y TECNOLOGÍA',
      ruc:     '20602322816',
      direccion:'Cal. German Schreiber Nro. 276 San Isidro - Lima',
      telefono:'+51 949348302',
      email:   'comercial@consitecing.com',
      gerente: '',
      cuentaBanco: '',
      // márgenes default para cotizador
      margenMaterial: 1.3,
      margenGratting: 1.5,
      gastosGenerales: 0.20,
      utilidad: 0.10,
      igv: 0.18,
      moneda: 'PEN',
      validezOferta: '15 Días',
      formaPago: 'Factura a 15 días',
      metodoPago: 'Transferencia bancaria',
      modalidad: 'Todo costo',
    });
  }
  function setConfig(cfg){ write(KEYS.config, cfg); }

  // ===== COUNTERS (correlativos) =====
  function getCounter(key){
    const c = read(KEYS.counters, {});
    return c[key] || 0;
  }
  function nextCounter(key){
    const c = read(KEYS.counters, {});
    c[key] = (c[key]||0) + 1;
    write(KEYS.counters, c);
    return c[key];
  }
  // Genera código tipo CI26-CMP-OE-061
  function nextCotizacionCode(clienteCorto='XXX'){
    const year = new Date().getFullYear().toString().slice(-2);
    const n = nextCounter('cotizacion_' + year);
    return `CI${year}-${clienteCorto}-OE-${String(n).padStart(3,'0')}`;
  }
  function nextOCCode(){
    const year = new Date().getFullYear().toString().slice(-2);
    const n = nextCounter('oc_' + year);
    return `OC-${year}-${String(n).padStart(4,'0')}`;
  }

  // ===== CRUD genérico (dual: localStorage o SharePoint) =====
  // collKey: clave de KEYS (ej 'clientes'). Si está en SP_LISTAS y modo=sharepoint, usa SP.
  function makeCRUD(collKey){
    const key = KEYS[collKey];
    const spListName = SP_LISTAS[collKey]; // undefined si la colección no se sincroniza

    // ----- modo SharePoint -----
    // Si el cache aún no se llenó (módulo recién abierto, antes del sync), usar el mirror localStorage
    function spList(){
      if(spCache[collKey]) return spCache[collKey];
      const mirror = read(key, []);
      if(mirror && mirror.length) spCache[collKey] = mirror;
      return mirror;
    }
    function spGet(id){ return spList().find(x => x.id === id); }
    function spSave(item){
      if(!item.id){ item.id = uid(); item.createdAt = Date.now(); }
      item.updatedAt = Date.now();
      const arr = spCache[collKey] = spCache[collKey] || [];
      const i = arr.findIndex(x => x.id === item.id);
      if(i >= 0) arr[i] = item; else arr.push(item);
      // mirror local inmediato (offline-friendly)
      write(key, arr);
      // push async a SP
      SPAdapter.save(spListName, item).then(saved => {
        // captura spId asignado por SP
        if(saved && saved.spId && !item.spId){
          item.spId = saved.spId;
          write(key, arr);
        }
      }).catch(err => reportSPError('SAVE', spListName, err));
      return item;
    }
    function spRemove(id){
      const arr = spCache[collKey] || [];
      const item = arr.find(x => x.id === id);
      spCache[collKey] = arr.filter(x => x.id !== id);
      write(key, spCache[collKey]);
      if(item && item.spId){
        SPAdapter.remove(spListName, item.spId)
          .catch(err => reportSPError('REMOVE', spListName, err));
      }
    }
    function spBulkSet(items){
      spCache[collKey] = items || [];
      write(key, spCache[collKey]);
      // No hacemos sincronización masiva (sería peligroso); usar migrate() para eso
    }

    // ----- modo localStorage -----
    function lsList(){ return read(key, []); }
    function lsGet(id){ return read(key, []).find(x => x.id === id); }
    function lsSave(item){
      const arr = read(key, []);
      if(!item.id){ item.id = uid(); item.createdAt = Date.now(); }
      item.updatedAt = Date.now();
      const i = arr.findIndex(x => x.id === item.id);
      if(i >= 0) arr[i] = item; else arr.push(item);
      write(key, arr);
      return item;
    }
    function lsRemove(id){
      const arr = read(key, []).filter(x => x.id !== id);
      write(key, arr);
    }
    function lsBulkSet(items){ write(key, items); }

    return {
      list:    () => (isSPMode() && spListName) ? spList()   : lsList(),
      get:     (id) => (isSPMode() && spListName) ? spGet(id)   : lsGet(id),
      save:    (item) => (isSPMode() && spListName) ? spSave(item) : lsSave(item),
      remove:  (id) => (isSPMode() && spListName) ? spRemove(id) : lsRemove(id),
      bulkSet: (items) => (isSPMode() && spListName) ? spBulkSet(items) : lsBulkSet(items),
    };
  }

  const clientes     = makeCRUD('clientes');
  const proyectos    = makeCRUD('proyectos');
  const cotizaciones = makeCRUD('cotizaciones');
  const ocs          = makeCRUD('ocs');
  const cargos       = makeCRUD('cargos');
  const personal     = makeCRUD('personal');
  const materiales   = makeCRUD('materiales');
  const contactosConsitec  = makeCRUD('contactosConsitec');
  const proveedores        = makeCRUD('proveedores');
  const asignaciones       = makeCRUD('asignaciones');
  const reportesDiarios    = makeCRUD('reportesDiarios');
  const catMovilizacion    = makeCRUD('catMovilizacion');
  const catEquipos         = makeCRUD('catEquipos');
  const catNotas           = makeCRUD('catNotas');
  const catExclusiones     = makeCRUD('catExclusiones');
  const catResponsabilidades = makeCRUD('catResponsabilidades');
  const inventario           = makeCRUD('inventario');

  // ===== Helpers para sub-listas de cliente (contactos + areas) =====
  function clienteAddContacto(clienteId, contacto){
    const c = clientes.get(clienteId); if(!c) return;
    c.contactos = c.contactos || [];
    if(!contacto.id) contacto.id = uid('co');
    const i = c.contactos.findIndex(x => x.id === contacto.id);
    if(i >= 0) c.contactos[i] = contacto; else c.contactos.push(contacto);
    clientes.save(c);
    return contacto;
  }
  function clienteAddArea(clienteId, area){
    const c = clientes.get(clienteId); if(!c) return;
    c.areas = c.areas || [];
    const exists = c.areas.find(a => a.toLowerCase() === area.toLowerCase());
    if(!exists){ c.areas.push(area); clientes.save(c); }
  }
  // Crear/encontrar item de catálogo por descripción (auto-guardado)
  function ensureCatItem(catCRUD, descripcion, extra={}){
    if(!descripcion || !descripcion.trim()) return null;
    const items = catCRUD.list();
    const existing = items.find(x => (x.descripcion||'').toLowerCase() === descripcion.toLowerCase());
    if(existing) return existing;
    return catCRUD.save({descripcion, ...extra});
  }

  // Convertir cotización aprobada en proyecto (o actualizar existente)
  function crearProyectoDesdeCotizacion(cotizacionId){
    const c = cotizaciones.get(cotizacionId);
    if(!c) return null;
    // Si ya tiene proyecto vinculado, retornarlo
    if(c.proyectoId){
      const existente = proyectos.get(c.proyectoId);
      if(existente) return existente;
    }
    // Generar código de proyecto
    const year = new Date().getFullYear().toString().slice(-2);
    const n = proyectos.list().length + 1;
    const codigo = `PRY-${year}-${String(n).padStart(3,'0')}`;
    const tipoLbl = c.tipo === 'suministro' ? 'Suministro' : 'Servicio';
    const proyecto = proyectos.save({
      codigo,
      nombre: c.servicio || `${tipoLbl} para ${c.clienteRazon||''}`,
      clienteId: c.clienteId || '',
      clienteRazon: c.clienteRazon || '',
      lugar: c.lugar || '',
      cotizacionId: c.id,
      cotizacionCodigo: c.codigo || '',
      cotizacionTipo: c.tipo || 'servicios',
      estado: 'planificado',
      fechaInicio: '',
      fechaFin: '',
      avance: 0,
      monto: c.totales?.totalFinal || 0,
      responsable: '',
      notas: `Generado automáticamente desde cotización ${c.codigo||''}.`,
    });
    // Vincular cotización al proyecto
    c.proyectoId = proyecto.id;
    cotizaciones.save(c);
    return proyecto;
  }

  // ===== Seed inicial (datos demo si vacío) =====
  function seedIfEmpty(){
    if(cargos.list().length === 0){
      // Catálogo de cargos (con tarifas HH). NO contiene personas, solo el tipo de puesto.
      [
        {cargo:'Supervisor de Obra',     categoria:'Supervisión', costoRealHH: 52.94, ganancia: 1.3, precioHH: 68.82},
        {cargo:'Supervisor de Campo',    categoria:'Supervisión', costoRealHH: 42.38, ganancia: 1.3, precioHH: 55.10},
        {cargo:'Ingeniero de Seguridad', categoria:'Seguridad',   costoRealHH: 35.46, ganancia: 1.3, precioHH: 46.10},
        {cargo:'Proyectista CAD',        categoria:'Ingeniería',  costoRealHH: 29.73, ganancia: 1.3, precioHH: 38.65},
        {cargo:'Soldador',               categoria:'Técnico',     costoRealHH: 29.38, ganancia: 1.3, precioHH: 38.20},
        {cargo:'Técnico Mecánico',       categoria:'Técnico',     costoRealHH: 24.62, ganancia: 1.3, precioHH: 32.00},
        {cargo:'Técnico Electricista',   categoria:'Técnico',     costoRealHH: 24.62, ganancia: 1.3, precioHH: 32.00},
        {cargo:'Técnico Instrumentista', categoria:'Técnico',     costoRealHH: 26.92, ganancia: 1.3, precioHH: 35.00},
        {cargo:'Maestro Operario',       categoria:'Operario',    costoRealHH: 30.92, ganancia: 1.3, precioHH: 40.20},
        {cargo:'Operario',               categoria:'Operario',    costoRealHH: 24.77, ganancia: 1.3, precioHH: 32.20},
        {cargo:'Peón',                   categoria:'Operario',    costoRealHH: 15.38, ganancia: 1.3, precioHH: 20.00},
      ].forEach(c => cargos.save(c));
    }
    // Migración: si la colección personal tenía cargos viejos (sin nombre), migrar a cargos
    const personalLegacy = personal.list().filter(p => !p.nombre && p.cargo);
    if(personalLegacy.length > 0 && cargos.list().length === personalLegacy.length){
      // Ya migrado en seed anterior; eliminar los registros viejos sin nombre del personal
      personalLegacy.forEach(p => personal.remove(p.id));
    }
    if(materiales.list().length === 0){
      // ganancia + logistica reemplazan al "margen". margen_total = ganancia * logistica
      [
        {descripcion:'CANAL "U" A-36 X 6 MT - ASTM A36 3"X 1.41" X 4.10 LBS/PIE', um:'UND.', costoReal: 146.66, ganancia: 1.3, logistica: 1.0, fabricante:''},
        {descripcion:'TUBO CUADRADO LAC A500 X 3 MM X 6 MT (3" X 3")', um:'UND.', costoReal: 128.81, ganancia: 1.3, logistica: 1.0, fabricante:''},
        {descripcion:'TUBO REDONDO LAC 2.5MM X 6 MT - 1.1/2"', um:'UND.', costoReal: 69.00, ganancia: 1.3, logistica: 1.0, fabricante:''},
        {descripcion:'PLANCHA ACERO A-36 (12.0 MM) 1/2" X 1200 MM X 2400 MM', um:'UND.', costoReal: 805.00, ganancia: 1.3, logistica: 1.0, fabricante:''},
        {descripcion:'MALLA GRATING ASTM A36 W19-4, NORMA NAAM MBG 531, 1MX1M', um:'UND.', costoReal: 130.00, ganancia: 1.5, logistica: 1.0, fabricante:'DL SUPPLY & LOGISTIC SAC'},
        {descripcion:'ELECTRODOS 7018 1/8"', um:'KG.',  costoReal: 25.00,  ganancia: 1.3, logistica: 1.0, fabricante:''},
        {descripcion:'ELECTRODO 6011 1/8"',  um:'KG.',  costoReal: 22.60,  ganancia: 1.3, logistica: 1.0, fabricante:''},
        {descripcion:'DISCO DE CORTE 7"',    um:'UND.', costoReal: 6.00,   ganancia: 1.4, logistica: 1.0, fabricante:''},
        {descripcion:'DISCO DE CORTE 4.5"',  um:'UND.', costoReal: 4.00,   ganancia: 1.4, logistica: 1.0, fabricante:''},
        {descripcion:'DISCO DE DESBASTE 4.5"', um:'UND.', costoReal: 4.20, ganancia: 1.3, logistica: 1.0, fabricante:''},
        {descripcion:'PINTURA BASE EPÓXICO COLOR GRIS', um:'GAL.', costoReal: 368.64, ganancia: 1.3, logistica: 1.0, fabricante:''},
        {descripcion:'PINTURA DE ACABADO COLOR ROJO',    um:'GAL.', costoReal: 372.88, ganancia: 1.3, logistica: 1.0, fabricante:''},
        {descripcion:'PERNO DE ANCLAJE PARA CONCRETO 3/8" X 4 1/2"', um:'UND.', costoReal: 3.80, ganancia: 1.3, logistica: 1.0, fabricante:''},
      ].forEach(m => materiales.save(m));
    }
    if(catMovilizacion.list().length === 0){
      // Precios unitarios HH (igual al Excel de CONSITEC)
      [
        {descripcion:'Movilización Terrestre', um:'hh', precio: 1.51},
        {descripcion:'Alimentación en mina',   um:'hh', precio: 3.52},
        {descripcion:'Habitabilidad',          um:'hh', precio: 5.00},
        {descripcion:'Movilización Aérea',     um:'hh', precio: 20.00},
      ].forEach(x => catMovilizacion.save(x));
    }
    if(catEquipos.list().length === 0){
      // Precios unitarios HH (igual al Excel de CONSITEC)
      [
        {descripcion:'Equipos',      um:'hh', precio: 2.68},
        {descripcion:'Herramientas', um:'hh', precio: 1.91},
        {descripcion:'Software',     um:'hh', precio: 1.34},
        {descripcion:'Andamios Multidireccionales Certificados (2 cuerpos)', um:'hh', precio: 5.60},
        {descripcion:'Máquina de soldar', um:'hh', precio: 8.00},
        {descripcion:'Equipo de oxicorte', um:'hh', precio: 6.00},
      ].forEach(x => catEquipos.save(x));
    }
    if(catNotas.list().length === 0){
      [
        {descripcion:'Los precios están expresados en Soles (PEN) y no incluye el IGV.'},
        {descripcion:'Los trabajos se ejecutarán conforme a los planos, especificaciones técnicas y alcances entregados por el cliente.'},
        {descripcion:'Cualquier trabajo adicional o fuera del alcance será cotizado y aprobado previamente.'},
        {descripcion:'Los plazos de ejecución pueden variar por causas externas (clima, accesos, permisos, material provisto por el cliente).'},
        {descripcion:'La oferta económica es válida por la totalidad del servicio.'},
      ].forEach(x => catNotas.save(x));
    }
    if(catExclusiones.list().length === 0){
      [
        {descripcion:'Reemplazo de infraestructura existente (tablero eléctrico y poste).'},
        {descripcion:'Suministro de energía eléctrica en obra.'},
        {descripcion:'Permisos y licencias de obra.'},
        {descripcion:'Tránsito de personal y equipos por zonas restringidas.'},
      ].forEach(x => catExclusiones.save(x));
    }
    if(catResponsabilidades.list().length === 0){
      [
        {descripcion:'Disponer de los equipos detenidos para los trabajos respectivos.'},
        {descripcion:'Brindar el soporte necesario para los trabajos respectivos.'},
        {descripcion:'Habitabilidad del personal.'},
        {descripcion:'Proveer accesos y permisos de ingreso a planta.'},
        {descripcion:'Suministrar planos as-built y especificaciones técnicas.'},
      ].forEach(x => catResponsabilidades.save(x));
    }
    if(contactosConsitec.list().length === 0){
      [
        {nombre:'Elvia Rebeca', apellidos:'Ramírez Mendez', cargo:'Gerente General', area:'Gerencia',
         telefono:'+51 949348302', email:'gerencia@consitecing.com',
         rolFirma:true, rolAprobador:true},
        {nombre:'Monica', apellidos:'Balarezo', cargo:'Gestor de Compras', area:'Logística',
         telefono:'', email:'comercial@consitecing.com',
         rolGestor:true},
        {nombre:'Omar', apellidos:'Medina', cargo:'Solicitante / Comercial', area:'Comercial',
         telefono:'', email:'omar.medina@consitecing.com',
         rolSolicitante:true},
        {nombre:'Celeste', apellidos:'Castillo', cargo:'Aprobador', area:'Administración',
         telefono:'', email:'celeste.castillo@consitecing.com',
         rolAprobador:true},
        {nombre:'Melfin', apellidos:'Villacorta', cargo:'Comercial', area:'Comercial',
         telefono:'+51 988651543', email:'comercial@consitecing.com'},
      ].forEach(x => contactosConsitec.save(x));
    }
    if(proveedores.list().length === 0){
      [
        {razonSocial:'DL SUPPLY & LOGISTIC SAC',
         ruc:'20603549217', rubro:'Materiales metálicos',
         contacto:'Ventas DL', telefono:'+51 987 654 321',
         email:'ventas@dlsupply.com.pe',
         direccion:'Av. Industrial 123, Ate, Lima',
         diasCredito: 30, formaPagoDefault:'Crédito 30 días'},
        {razonSocial:'INPROCESS AUTOMATIZACION INDUSTRIAL',
         ruc:'20514706302', rubro:'Eléctrico / Instrumentación',
         contacto:'Ventas Inprocess', telefono:'959 144 276',
         email:'ventas@inprocess.com.pe',
         direccion:'Cal. Pablo de Olavide N° 110, San Isidro - Lima',
         diasCredito: 0, formaPagoDefault:'Al contado'},
      ].forEach(x => proveedores.save(x));
    }
    if(clientes.list().length === 0){
      clientes.save({
        razonSocial:'Compañía Minera Poderosa S.A.',
        ruc:'20137025354',
        corto:'CMP',
        direccion:'Av. Primavera 834, Surco, Lima',
        sector:'Minería',
        areas:['Zona de Transportes','Planta Chaparrosa','Taller Eléctrico Vijus'],
        contactos:[{
          id: uid('co'),
          nombre:'Ing. Miguel Herrera',
          cargo:'Jefe de Instrumentación y Control de procesos / Taller Eléctrico Vijus',
          telefono:'+51 948329768',
          email:'mherrerac@poderosa.com.pe',
        }],
      });
    }
  }

  // ===== Export / Import (backup JSON) =====
  function exportAll(){
    const data = {};
    Object.entries(KEYS).forEach(([k,v]) => { data[k] = read(v, null); });
    data._exportedAt = new Date().toISOString();
    return data;
  }
  function importAll(data){
    Object.entries(KEYS).forEach(([k,v]) => {
      if(data[k] !== undefined && data[k] !== null) write(v, data[k]);
    });
  }
  function clearAll(){
    Object.values(KEYS).forEach(k => localStorage.removeItem(k));
  }

  // ===== Migración (localStorage → SharePoint, una sola vez) =====
  // Sube TODOS los items del localStorage a las listas de SP.
  // Útil para usuarios que ya tenían datos antes de activar el modo SP.
  async function migrateLocalToSharePoint(onProgress){
    if(!isSPMode()) throw new Error('Activa SP_CONFIG.modo = "sharepoint" y haz login antes de migrar.');
    const reporte = [];
    for(const [collKey, listName] of Object.entries(SP_LISTAS)){
      const items = read(KEYS[collKey], []);
      if(!items.length){ reporte.push({list:listName, ok:0, err:0, total:0}); continue; }
      let ok = 0, err = 0;
      for(const item of items){
        try {
          const saved = await SPAdapter.save(listName, { ...item, spId: undefined });
          if(saved && saved.spId){ item.spId = saved.spId; }
          ok++;
        } catch(e){ console.error('Migración fallo en', listName, item, e); err++; }
        if(onProgress) onProgress({ list:listName, done:ok+err, total:items.length });
      }
      // Persistir spId en localStorage para que el próximo load no duplique
      write(KEYS[collKey], items);
      reporte.push({ list:listName, ok, err, total:items.length });
    }
    // Limpiar cache y recargar
    spReady = false; spReadyPromise = null;
    await loadFromSharePoint();
    return reporte;
  }

  return {
    KEYS, uid,
    SP_LISTAS, isSPMode, isReady, loadFromSharePoint, migrateLocalToSharePoint,
    getConfig, setConfig,
    getCounter, nextCounter, nextCotizacionCode, nextOCCode,
    clientes, proyectos, cotizaciones, ocs, cargos, personal, materiales,
    contactosConsitec, proveedores, asignaciones, reportesDiarios,
    catMovilizacion, catEquipos,
    catNotas, catExclusiones, catResponsabilidades,
    inventario,
    clienteAddContacto, clienteAddArea, ensureCatItem, crearProyectoDesdeCotizacion,
    seedIfEmpty, exportAll, importAll, clearAll,
  };
})();

// ===== Utilidades globales =====
const fmtPEN = (n) => 'S/ ' + (Number(n)||0).toLocaleString('es-PE',{minimumFractionDigits:2, maximumFractionDigits:2});

// Etiqueta para mostrar personal: "Nombre · Cargo" o solo lo que tenga
const personalLabel = (p) => {
  if(!p) return '(borrado)';
  if(p.nombre && p.cargo) return `${p.nombre} · ${p.cargo}`;
  return p.nombre || p.cargo || '(sin datos)';
};
const personalNombre = (p) => p?.nombre || p?.cargo || '(sin datos)';
const fmtNum = (n, d=2) => (Number(n)||0).toLocaleString('es-PE',{minimumFractionDigits:d, maximumFractionDigits:d});
const fmtDate = (ts) => {
  if(!ts) return '-';
  const d = new Date(ts);
  return d.toLocaleDateString('es-PE',{day:'2-digit',month:'2-digit',year:'numeric'});
};
const todayISO = () => new Date().toISOString().slice(0,10);
const toast = (msg, type='') => {
  const el = document.createElement('div');
  el.className = 'toast ' + type;
  el.textContent = msg;
  document.body.appendChild(el);
  setTimeout(() => el.remove(), 2800);
};
const numEnLetras = (n) => {
  // Conversor de números a letras (soles peruanos)
  const u = ['','UNO','DOS','TRES','CUATRO','CINCO','SEIS','SIETE','OCHO','NUEVE','DIEZ','ONCE','DOCE','TRECE','CATORCE','QUINCE','DIECISÉIS','DIECISIETE','DIECIOCHO','DIECINUEVE','VEINTE'];
  const d = ['','','VEINTI','TREINTA','CUARENTA','CINCUENTA','SESENTA','SETENTA','OCHENTA','NOVENTA'];
  const c = ['','CIENTO','DOSCIENTOS','TRESCIENTOS','CUATROCIENTOS','QUINIENTOS','SEISCIENTOS','SETECIENTOS','OCHOCIENTOS','NOVECIENTOS'];
  function seccion(num){
    if(num === 0) return '';
    if(num <= 20) return u[num];
    if(num < 30) return num === 20 ? 'VEINTE' : 'VEINTI' + u[num-20].toLowerCase().toUpperCase();
    if(num < 100){
      const dec = Math.floor(num/10), uni = num%10;
      return d[dec] + (uni ? ' Y ' + u[uni] : '');
    }
    if(num === 100) return 'CIEN';
    if(num < 1000){
      const cen = Math.floor(num/100), rest = num%100;
      return c[cen] + (rest ? ' ' + seccion(rest) : '');
    }
    return '';
  }
  function miles(num){
    if(num < 1000) return seccion(num);
    const mil = Math.floor(num/1000), rest = num%1000;
    let mTxt = mil === 1 ? 'MIL' : seccion(mil) + ' MIL';
    return mTxt + (rest ? ' ' + seccion(rest) : '');
  }
  function millones(num){
    if(num < 1000000) return miles(num);
    const mll = Math.floor(num/1000000), rest = num%1000000;
    let mTxt = mll === 1 ? 'UN MILLÓN' : miles(mll) + ' MILLONES';
    return mTxt + (rest ? ' ' + miles(rest) : '');
  }
  const entero = Math.floor(n);
  const cents = Math.round((n - entero)*100);
  return millones(entero) + ' CON ' + String(cents).padStart(2,'0') + '/100 SOLES';
};

// Inicializa seed solo en modo localStorage. En modo SP, los datos vienen del servidor.
if(typeof window !== 'undefined'){
  if(!DB.isSPMode()) DB.seedIfEmpty();
}
