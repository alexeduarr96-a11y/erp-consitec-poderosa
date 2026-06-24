# =====================================================
# Crea la lista 'CargosEntrega' en SharePoint
# Para registrar entregas de equipos/herramientas a supervisores
# con numero correlativo, datos del receptor, proyecto, items y firma
# =====================================================
$SiteUrl     = "https://consitecingenierossac.sharepoint.com/sites/erp-consitec-poderosa"
$PnPClientId = "b421dc62-3f38-4512-91ff-bed9f98c1534"
$TenantId    = "5b2acc65-31fe-49c9-aa4e-8066ac119e65"

Import-Module PnP.PowerShell -ErrorAction Stop

Write-Host "Conectando a SharePoint: $SiteUrl" -ForegroundColor Cyan
try {
    Connect-PnPOnline -Url $SiteUrl -Interactive -ClientId $PnPClientId -Tenant $TenantId -ErrorAction Stop
} catch {
    Write-Host "[ERROR] $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# Crear lista si no existe
$lista = Get-PnPList -Identity "CargosEntrega" -ErrorAction SilentlyContinue
if ($lista) {
    Write-Host "[!] La lista 'CargosEntrega' ya existe. Verificando columnas..." -ForegroundColor Yellow
} else {
    Write-Host "Creando lista 'CargosEntrega'..." -ForegroundColor Cyan
    New-PnPList -Title "CargosEntrega" -Template GenericList -OnQuickLaunch | Out-Null
    Write-Host "[OK] Lista creada" -ForegroundColor Green
}

# Columnas requeridas
$columnas = @(
    @{Nombre="LocalId"; Tipo="Text"},
    @{Nombre="Numero"; Tipo="Text"},                # CARGO-2026-001
    @{Nombre="Fecha"; Tipo="DateTime"},
    @{Nombre="FechaDevolucion"; Tipo="DateTime"},
    @{Nombre="SupervisorId"; Tipo="Text"},          # PersonalId
    @{Nombre="SupervisorNombre"; Tipo="Text"},
    @{Nombre="SupervisorDni"; Tipo="Text"},
    @{Nombre="SupervisorCargo"; Tipo="Text"},
    @{Nombre="ProyectoId"; Tipo="Text"},
    @{Nombre="ProyectoCodigo"; Tipo="Text"},
    @{Nombre="ProyectoNombre"; Tipo="Text"},
    @{Nombre="UnidadMineraDestino"; Tipo="Text"},
    @{Nombre="ItemsJson"; Tipo="Note"},             # array de equipos
    @{Nombre="Observaciones"; Tipo="Note"},
    @{Nombre="Estado"; Tipo="Text"}                 # emitido/devuelto/parcial
)

$existentes = Get-PnPField -List "CargosEntrega" | Select-Object -ExpandProperty InternalName
foreach ($col in $columnas) {
    if ($existentes -contains $col.Nombre) {
        Write-Host "  -> $($col.Nombre) ya existe" -ForegroundColor Gray
        continue
    }
    try {
        Add-PnPField -List "CargosEntrega" -DisplayName $col.Nombre -InternalName $col.Nombre -Type $col.Tipo | Out-Null
        Write-Host "  [+] $($col.Nombre) ($($col.Tipo))" -ForegroundColor Green
    } catch {
        Write-Host "  [!] $($col.Nombre): $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "=============================================" -ForegroundColor Green
Write-Host "  LISTA CargosEntrega LISTA" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green

Disconnect-PnPOnline
