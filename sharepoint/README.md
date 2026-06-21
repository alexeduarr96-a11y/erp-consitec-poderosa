# Migración del ERP CONSITEC / Poderosa a SharePoint + Azure

Este folder contiene todo lo necesario para migrar el ERP de localStorage local a SharePoint Lists hospedado en Azure Static Web Apps.

## 🌐 URL del sitio SharePoint

```
https://consitecingenierossac.sharepoint.com/sites/erp-consitec-poderosa
```

## 📋 Pasos de migración (en orden)

### Paso 1 — Crear las 10 listas en SharePoint
- **Opción A (recomendada):** ejecutar `02-crear-listas.ps1` (script automático)
- **Opción B (manual):** crear cada lista siguiendo `01-listas-esquema.md`

### Paso 2 — Registrar la app en Azure AD
- Ir a [portal.azure.com](https://portal.azure.com) → Azure AD → App registrations → New registration
- Nombre: `ERP CONSITEC Poderosa`
- Tipo de cuenta: **Solo cuentas de mi organización** (single tenant)
- Redirect URI: `https://erp-consitec-poderosa.azurestaticapps.net` (lo cambiamos después)
- Anotar: **Client ID** y **Tenant ID**

### Paso 3 — Configurar permisos API
- En la app registrada → API permissions → Add permission → Microsoft Graph
- Delegados:
  - `User.Read` (login básico)
  - `Sites.ReadWrite.All` (leer/escribir listas SharePoint)
- Click en "Grant admin consent"

### Paso 4 — Configurar `sharepoint-config.js`
Llenar los datos en `js/sharepoint-config.js`:
```js
const SP_CONFIG = {
  tenantId: 'XXXX-XXXX-XXXX',          // Del paso 2
  clientId: 'XXXX-XXXX-XXXX',          // Del paso 2
  siteUrl: 'https://consitecingenierossac.sharepoint.com/sites/erp-consitec-poderosa',
};
```

### Paso 5 — Subir a Azure Static Web Apps
- Crear cuenta gratis en [github.com](https://github.com) si no tienes
- Subir el repo del ERP a GitHub
- Ir a [portal.azure.com](https://portal.azure.com) → Static Web Apps → Create
- Conectar con el repo de GitHub
- Build presets: **Custom** (HTML estático, sin build)
- App location: `/`
- Output: vacío

### Paso 6 — Invitar a los 15 usuarios
- En SharePoint → sitio → Settings → Site permissions
- Agregar los 15 usuarios al grupo "Miembros"
- Pueden acceder a la URL: `https://erp-consitec-poderosa.azurestaticapps.net`

---

## 📂 Estructura de archivos

```
sharepoint/
├── README.md                    ← este archivo
├── 01-listas-esquema.md         ← detalle de cada lista
├── 02-crear-listas.ps1          ← script PowerShell automático
└── 03-permisos-azure.md         ← pasos detallados Azure AD
```

---

## ⚠ Modo de operación

La app `storage.js` ahora tiene un **adapter** que permite usar:
- **`localStorage`** (modo actual, hasta que esté SharePoint listo)
- **`sharepoint`** (modo destino, cuando todo esté configurado)

Para cambiar de modo, edita `js/sharepoint-config.js`:
```js
const SP_CONFIG = {
  modo: 'localStorage',  // ← cámbialo a 'sharepoint' cuando esté todo listo
  ...
};
```
