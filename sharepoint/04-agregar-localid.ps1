# =====================================================
# Agrega columna 'LocalId' (Text) a las 10 listas del ERP
# Esta columna almacena el UUID local generado por la app
# para mantener referencias entre entidades (cliente -> oc, etc).
# =====================================================
$SiteUrl     = "https://consitecingenierossac.sharepoint.com/sites/erp-consitec-poderosa"
# Client ID de la app "ERP CONSITEC PnP" (dedicada para PowerShell - Mobile/Desktop)
$PnPClientId = "b421dc62-3f38-4512-91ff-bed9f98c1534"
$TenantId    = "5b2acc65-31fe-49c9-aa4e-8066ac119e65"

Import-Module PnP.PowerShell -ErrorAction Stop

Write-Host "Conectando a SharePoint: $SiteUrl" -ForegroundColor Cyan
Write-Host "(Se abrira una ventana para iniciar sesion con tu cuenta M365)" -ForegroundColor Gray
try {
    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $PnPClientId -Tenant $TenantId -ErrorAction Stop
} catch {
    Write-Host ""
    Write-Host "[ERROR] No se pudo conectar:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}

# Verificar conexion
try {
    $ctx = Get-PnPContext -ErrorAction Stop
    if (-not $ctx) { throw "Sin contexto" }
    Write-Host "[OK] Conectado correctamente" -ForegroundColor Green
} catch {
    Write-Host ""
    Write-Host "[ERROR] Login no completado. Aborta." -ForegroundColor Red
    exit 1
}

$listas = @(
    "Clientes","Proveedores","Cargos","Personal","Gestion",
    "Materiales","Cotizaciones","OrdenesCompra","Proyectos","Reportes"
)

foreach ($lista in $listas) {
    Write-Host ""
    Write-Host "Lista: $lista" -ForegroundColor Cyan

    # Verificar si LocalId ya existe
    $existing = Get-PnPField -List $lista -Identity "LocalId" -ErrorAction SilentlyContinue
    if ($existing) {
        Write-Host "  -> LocalId ya existe, omitiendo" -ForegroundColor Gray
        continue
    }

    try {
        Add-PnPField -List $lista -DisplayName "LocalId" -InternalName "LocalId" -Type Text | Out-Null
        Write-Host "  [+] LocalId agregada" -ForegroundColor Green
    } catch {
        Write-Host "  [!] Error: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "=============================================" -ForegroundColor Green
Write-Host "  COLUMNAS LocalId AGREGADAS" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green

Disconnect-PnPOnline
