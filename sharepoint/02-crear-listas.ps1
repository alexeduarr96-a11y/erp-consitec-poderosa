# =====================================================
# Crea las 10 listas SharePoint para el ERP CONSITEC / Poderosa
# Compatible con PowerShell 5.1 y PowerShell 7+
# =====================================================

# CONFIGURACION
$SiteUrl = "https://consitecingenierossac.sharepoint.com/sites/erp-consitec-poderosa"
# Client ID de la app "ERP CONSITEC PnP" (Mobile/Desktop, dedicada a PowerShell)
$PnPClientId = "b421dc62-3f38-4512-91ff-bed9f98c1534"
# Tenant ID (Id. de directorio / inquilino)
$TenantId    = "5b2acc65-31fe-49c9-aa4e-8066ac119e65"

# Instalar PnP.PowerShell si no esta
if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
    Write-Host "Instalando PnP.PowerShell..." -ForegroundColor Yellow
    Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -SkipPublisherCheck -AllowClobber
}
Import-Module PnP.PowerShell -ErrorAction Stop

# Conectar a SharePoint usando el ClientId publico de PnP Management Shell
Write-Host "Conectando a SharePoint: $SiteUrl" -ForegroundColor Cyan
Write-Host "(Se abrira una ventana de login con tu cuenta de M365)" -ForegroundColor Gray
try {
    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $PnPClientId -Tenant $TenantId
    Write-Host "[OK] Conectado correctamente" -ForegroundColor Green
} catch {
    Write-Host "[ERROR] No se pudo conectar: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Write-Host "SOLUCION: Antes de ejecutar este script debes autorizar PnP en tu tenant." -ForegroundColor Yellow
    Write-Host "Ejecuta este comando (requiere ser admin global de M365):" -ForegroundColor Yellow
    Write-Host "  Register-PnPManagementShellAccess" -ForegroundColor White
    exit 1
}

# Funcion helper para crear listas
function Crear-Lista {
    param(
        [string]$Nombre,
        [array]$Columnas
    )

    $listaExistente = Get-PnPList -Identity $Nombre -ErrorAction SilentlyContinue
    if ($listaExistente) {
        $rta = Read-Host "La lista '$Nombre' ya existe. Recrear? (s/N)"
        if ($rta -eq "s") {
            Remove-PnPList -Identity $Nombre -Force
            Write-Host "  [X] Lista eliminada" -ForegroundColor Yellow
        } else {
            Write-Host "  -> Saltando '$Nombre'" -ForegroundColor Gray
            return
        }
    }

    Write-Host "Creando lista: $Nombre" -ForegroundColor Cyan
    New-PnPList -Title $Nombre -Template GenericList -OnQuickLaunch | Out-Null

    foreach ($col in $Columnas) {
        try {
            Add-PnPField -List $Nombre -DisplayName $col.Nombre -InternalName $col.Nombre -Type $col.Tipo | Out-Null
            Write-Host "  + $($col.Nombre) ($($col.Tipo))" -ForegroundColor Gray
        } catch {
            Write-Host "  ! Error agregando $($col.Nombre): $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
    Write-Host "[OK] Lista '$Nombre' creada" -ForegroundColor Green
}

# ============== 1. CLIENTES ==============
Crear-Lista -Nombre "Clientes" -Columnas @(
    @{Nombre="Ruc"; Tipo="Text"},
    @{Nombre="Corto"; Tipo="Text"},
    @{Nombre="Sector"; Tipo="Text"},
    @{Nombre="Direccion"; Tipo="Text"},
    @{Nombre="Contactos"; Tipo="Note"},
    @{Nombre="Areas"; Tipo="Note"},
    @{Nombre="LogoBase64"; Tipo="Note"},
    @{Nombre="Notas"; Tipo="Note"}
)

# ============== 2. PROVEEDORES ==============
Crear-Lista -Nombre "Proveedores" -Columnas @(
    @{Nombre="Ruc"; Tipo="Text"},
    @{Nombre="Rubro"; Tipo="Text"},
    @{Nombre="Contacto"; Tipo="Text"},
    @{Nombre="Telefono"; Tipo="Text"},
    @{Nombre="Email"; Tipo="Text"},
    @{Nombre="Direccion"; Tipo="Text"},
    @{Nombre="Web"; Tipo="Text"},
    @{Nombre="Cuenta"; Tipo="Text"},
    @{Nombre="DiasCredito"; Tipo="Number"},
    @{Nombre="FormaPagoDefault"; Tipo="Text"},
    @{Nombre="Calificacion"; Tipo="Text"},
    @{Nombre="Contactos"; Tipo="Note"},
    @{Nombre="Notas"; Tipo="Note"}
)

# ============== 3. CARGOS ==============
Crear-Lista -Nombre "Cargos" -Columnas @(
    @{Nombre="Categoria"; Tipo="Text"},
    @{Nombre="CostoRealHH"; Tipo="Number"},
    @{Nombre="Ganancia"; Tipo="Number"},
    @{Nombre="PrecioHH"; Tipo="Number"},
    @{Nombre="UsaHerramientas"; Tipo="Boolean"},
    @{Nombre="UsaSoftware"; Tipo="Boolean"},
    @{Nombre="UsaEquipos"; Tipo="Boolean"},
    @{Nombre="Notas"; Tipo="Note"}
)

# ============== 4. PERSONAL ==============
Crear-Lista -Nombre "Personal" -Columnas @(
    @{Nombre="Nombres"; Tipo="Text"},
    @{Nombre="Apellidos"; Tipo="Text"},
    @{Nombre="Dni"; Tipo="Text"},
    @{Nombre="Telefono"; Tipo="Text"},
    @{Nombre="Email"; Tipo="Text"},
    @{Nombre="CargoIdSP"; Tipo="Text"},
    @{Nombre="FechaIngreso"; Tipo="DateTime"},
    @{Nombre="Activo"; Tipo="Boolean"},
    @{Nombre="Especialidad"; Tipo="Text"},
    @{Nombre="Notas"; Tipo="Note"}
)

# ============== 5. GESTION ==============
Crear-Lista -Nombre "Gestion" -Columnas @(
    @{Nombre="Nombres"; Tipo="Text"},
    @{Nombre="Apellidos"; Tipo="Text"},
    @{Nombre="Cargo"; Tipo="Text"},
    @{Nombre="Area"; Tipo="Text"},
    @{Nombre="Email"; Tipo="Text"},
    @{Nombre="Telefono"; Tipo="Text"},
    @{Nombre="RolGestor"; Tipo="Boolean"},
    @{Nombre="RolSolicitante"; Tipo="Boolean"},
    @{Nombre="RolAprobador"; Tipo="Boolean"},
    @{Nombre="RolFirma"; Tipo="Boolean"},
    @{Nombre="Notas"; Tipo="Note"}
)

# ============== 6. MATERIALES ==============
Crear-Lista -Nombre "Materiales" -Columnas @(
    @{Nombre="Marca"; Tipo="Text"},
    @{Nombre="Modelo"; Tipo="Text"},
    @{Nombre="Fabricante"; Tipo="Text"},
    @{Nombre="Um"; Tipo="Text"},
    @{Nombre="CostoReal"; Tipo="Number"},
    @{Nombre="Ganancia"; Tipo="Number"},
    @{Nombre="Logistica"; Tipo="Number"},
    @{Nombre="Categoria"; Tipo="Text"},
    @{Nombre="Notas"; Tipo="Note"}
)

# ============== 7. COTIZACIONES ==============
Crear-Lista -Nombre "Cotizaciones" -Columnas @(
    @{Nombre="Tipo"; Tipo="Text"},
    @{Nombre="Estado"; Tipo="Text"},
    @{Nombre="Version"; Tipo="Number"},
    @{Nombre="Fecha"; Tipo="DateTime"},
    @{Nombre="Servicio"; Tipo="Note"},
    @{Nombre="Lugar"; Tipo="Text"},
    @{Nombre="Plazo"; Tipo="Text"},
    @{Nombre="ClienteIdSP"; Tipo="Text"},
    @{Nombre="ClienteRazon"; Tipo="Text"},
    @{Nombre="ContactoClienteId"; Tipo="Text"},
    @{Nombre="Atencion"; Tipo="Text"},
    @{Nombre="CargoContacto"; Tipo="Text"},
    @{Nombre="Contacto"; Tipo="Text"},
    @{Nombre="ContactoConsitec"; Tipo="Text"},
    @{Nombre="LogoCliente"; Tipo="Note"},
    @{Nombre="Alcance1"; Tipo="Note"},
    @{Nombre="Alcance2"; Tipo="Note"},
    @{Nombre="Mo1Json"; Tipo="Note"},
    @{Nombre="Mov1Json"; Tipo="Note"},
    @{Nombre="Eq1Json"; Tipo="Note"},
    @{Nombre="Mo2Json"; Tipo="Note"},
    @{Nombre="Mov2Json"; Tipo="Note"},
    @{Nombre="Eq2Json"; Tipo="Note"},
    @{Nombre="MatJson"; Tipo="Note"},
    @{Nombre="Gg"; Tipo="Number"},
    @{Nombre="Util"; Tipo="Number"},
    @{Nombre="Modalidad"; Tipo="Text"},
    @{Nombre="Validez"; Tipo="Text"},
    @{Nombre="FormaPago"; Tipo="Text"},
    @{Nombre="MetodoPago"; Tipo="Text"},
    @{Nombre="NotasJson"; Tipo="Note"},
    @{Nombre="ExclJson"; Tipo="Note"},
    @{Nombre="RespJson"; Tipo="Note"},
    @{Nombre="FichasJson"; Tipo="Note"},
    @{Nombre="TotalFinal"; Tipo="Number"},
    @{Nombre="Ganancia"; Tipo="Number"},
    @{Nombre="ProyectoIdSP"; Tipo="Text"}
)

# ============== 8. ORDENES DE COMPRA ==============
Crear-Lista -Nombre "OrdenesCompra" -Columnas @(
    @{Nombre="Fecha"; Tipo="DateTime"},
    @{Nombre="FechaEntrega"; Tipo="DateTime"},
    @{Nombre="FechaVencimiento"; Tipo="DateTime"},
    @{Nombre="CentroCompra"; Tipo="Text"},
    @{Nombre="GestorIdSP"; Tipo="Text"},
    @{Nombre="GestorNombre"; Tipo="Text"},
    @{Nombre="GestorEmail"; Tipo="Text"},
    @{Nombre="SolicitanteIdSP"; Tipo="Text"},
    @{Nombre="SolicitanteNombre"; Tipo="Text"},
    @{Nombre="SolicitanteEmail"; Tipo="Text"},
    @{Nombre="AprobadorIdSP"; Tipo="Text"},
    @{Nombre="AprobadorNombre"; Tipo="Text"},
    @{Nombre="AprobadorEmail"; Tipo="Text"},
    @{Nombre="ProveedorIdSP"; Tipo="Text"},
    @{Nombre="ProveedorRazon"; Tipo="Text"},
    @{Nombre="ProveedorRuc"; Tipo="Text"},
    @{Nombre="ProveedorDir"; Tipo="Text"},
    @{Nombre="ProveedorContacto"; Tipo="Text"},
    @{Nombre="ProveedorTel"; Tipo="Text"},
    @{Nombre="ProveedorEmail"; Tipo="Text"},
    @{Nombre="ProyectoIdSP"; Tipo="Text"},
    @{Nombre="ProyectoNombre"; Tipo="Text"},
    @{Nombre="ProyectoTexto"; Tipo="Text"},
    @{Nombre="Almacen"; Tipo="Text"},
    @{Nombre="Pedido"; Tipo="Text"},
    @{Nombre="NroCotizacion"; Tipo="Text"},
    @{Nombre="TratadoCon"; Tipo="Text"},
    @{Nombre="FechaEntregaText"; Tipo="Text"},
    @{Nombre="LugarEntrega"; Tipo="Note"},
    @{Nombre="FormaPago"; Tipo="Text"},
    @{Nombre="Moneda"; Tipo="Text"},
    @{Nombre="ItemsJson"; Tipo="Note"},
    @{Nombre="Igv"; Tipo="Number"},
    @{Nombre="Descuento"; Tipo="Number"},
    @{Nombre="Notas"; Tipo="Note"},
    @{Nombre="PagosJson"; Tipo="Note"},
    @{Nombre="Estado"; Tipo="Text"},
    @{Nombre="Total"; Tipo="Number"},
    @{Nombre="Pagado"; Tipo="Number"},
    @{Nombre="Saldo"; Tipo="Number"}
)

# ============== 9. PROYECTOS ==============
Crear-Lista -Nombre "Proyectos" -Columnas @(
    @{Nombre="Codigo"; Tipo="Text"},
    @{Nombre="ClienteIdSP"; Tipo="Text"},
    @{Nombre="ClienteRazon"; Tipo="Text"},
    @{Nombre="Lugar"; Tipo="Text"},
    @{Nombre="CotizacionIdSP"; Tipo="Text"},
    @{Nombre="CotizacionCodigo"; Tipo="Text"},
    @{Nombre="CotizacionTipo"; Tipo="Text"},
    @{Nombre="Estado"; Tipo="Text"},
    @{Nombre="FechaInicio"; Tipo="DateTime"},
    @{Nombre="FechaFin"; Tipo="DateTime"},
    @{Nombre="Avance"; Tipo="Number"},
    @{Nombre="Monto"; Tipo="Number"},
    @{Nombre="Responsable"; Tipo="Text"},
    @{Nombre="TareasJson"; Tipo="Note"},
    @{Nombre="BaselineJson"; Tipo="Note"},
    @{Nombre="CalendarioJson"; Tipo="Note"},
    @{Nombre="Notas"; Tipo="Note"}
)

# ============== 10. REPORTES ==============
Crear-Lista -Nombre "Reportes" -Columnas @(
    @{Nombre="Tipo"; Tipo="Text"},
    @{Nombre="ProyectoIdSP"; Tipo="Text"},
    @{Nombre="Fecha"; Tipo="DateTime"},
    @{Nombre="FechaIni"; Tipo="DateTime"},
    @{Nombre="FechaFin"; Tipo="DateTime"},
    @{Nombre="ResponsableIdSP"; Tipo="Text"},
    @{Nombre="ResponsableNombre"; Tipo="Text"},
    @{Nombre="Clima"; Tipo="Text"},
    @{Nombre="Turno"; Tipo="Text"},
    @{Nombre="ActividadesJson"; Tipo="Note"},
    @{Nombre="PersonalJson"; Tipo="Note"},
    @{Nombre="FotosJson"; Tipo="Note"},
    @{Nombre="Incidentes"; Tipo="Note"},
    @{Nombre="Observaciones"; Tipo="Note"},
    @{Nombre="AvancePorc"; Tipo="Number"},
    @{Nombre="Notas"; Tipo="Note"}
)

# ============== FINAL ==============
Write-Host ""
Write-Host "===========================================" -ForegroundColor Green
Write-Host "  TODAS LAS LISTAS CREADAS CORRECTAMENTE" -ForegroundColor Green
Write-Host "===========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Sitio: $SiteUrl" -ForegroundColor Cyan
Write-Host "Veras 10 listas en el panel izquierdo de SharePoint" -ForegroundColor Cyan
Write-Host ""
Write-Host "SIGUIENTE PASO: Registrar la app en Azure AD" -ForegroundColor Yellow
Write-Host "Ver: 03-permisos-azure.md" -ForegroundColor Yellow

Disconnect-PnPOnline
