# Paso 3 — Registrar la app en Azure AD

Necesitamos registrar la app en Azure Active Directory para que los usuarios puedan iniciar sesión con su `@consitecing.com` y la app pueda leer/escribir en las listas SharePoint en nombre del usuario.

## 🌐 Paso a paso

### 1. Ir al Portal de Azure
- URL: [https://portal.azure.com](https://portal.azure.com)
- Inicia sesión con tu cuenta de admin

### 2. Ir a Azure Active Directory
- En el buscador arriba: escribe **"Azure Active Directory"** o **"Microsoft Entra ID"**
- Click en el resultado

### 3. Registrar nueva app
- En el menú lateral: **App registrations** → **+ New registration**
- Completa:
  - **Name:** `ERP CONSITEC Poderosa`
  - **Supported account types:** `Accounts in this organizational directory only` (single tenant)
  - **Redirect URI:**
    - Plataforma: **Single-page application (SPA)**
    - URI: `http://localhost:5500` (para pruebas locales)
    - Después agregaremos la URL de producción
- Click **Register**

### 4. Anotar los IDs
En la pantalla de la app registrada, **copia y guarda**:
- **Application (client) ID:** `XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX`
- **Directory (tenant) ID:** `XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX`

### 5. Configurar permisos API
- En el menú lateral: **API permissions** → **+ Add a permission**
- Click **Microsoft Graph** → **Delegated permissions**
- Marca estos permisos:
  - ✅ `User.Read` (login básico)
  - ✅ `Sites.ReadWrite.All` (leer/escribir listas SharePoint)
- Click **Add permissions**
- **IMPORTANTE:** Click **Grant admin consent for [tu organización]** y confirma

### 6. Configurar Authentication (opcional pero recomendado)
- En el menú lateral: **Authentication**
- Bajo "Single-page application" agrega más Redirect URIs:
  - `https://erp-consitec-poderosa.azurestaticapps.net` (URL final cuando esté deployado)
  - Cualquier otra URL si tienes dominio personalizado
- Marca: **ID tokens** (used for implicit and hybrid flows)
- Click **Save**

---

## ✅ Resultado

Al final tendrás:
- ✅ App registrada en Azure AD con nombre `ERP CONSITEC Poderosa`
- ✅ Client ID y Tenant ID anotados (los pongo en `sharepoint-config.js`)
- ✅ Permisos para leer/escribir SharePoint Lists
- ✅ Admin consent otorgado

---

## 📝 Lo que necesito de ti para continuar

Cuando termines, mándame:

```
Client ID:  XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX
Tenant ID:  XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX
```

Con eso configuro el archivo `js/sharepoint-config.js` y el ERP queda listo para conectarse a SharePoint.
