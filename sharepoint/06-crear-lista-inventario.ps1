# =====================================================
# Crea la lista 'Inventario' en SharePoint
# Para el modulo de Inventario de Unidades Mineras
# (1 lista con campo 'UnidadMinera' para distinguir Vijus, Chaparrosa, etc)
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

try {
    $ctx = Get-PnPContext -ErrorAction Stop
    if (-not $ctx) { throw "Sin contexto" }
    Write-Host "[OK] Conectado" -ForegroundColor Green
} catch {
    Write-Host "[ERROR] Login no completado" -ForegroundColor Red
    exit 1
}

# Crear lista si no existe
$listaExistente = Get-PnPList -Identity "Inventario" -ErrorAction SilentlyContinue
if ($listaExistente) {
    Write-Host "[!] La lista 'Inventario' ya existe. Solo agregare columnas faltantes." -ForegroundColor Yellow
} else {
    Write-Host "Creando lista 'Inventario'..." -ForegroundColor Cyan
    New-PnPList -Title "Inventario" -Template GenericList -OnQuickLaunch | Out-Null
    Write-Host "[OK] Lista creada" -ForegroundColor Green
}

# Columnas a asegurar (idempotente)
$columnas = @(
    @{Nombre="LocalId"; Tipo="Text"},
    @{Nombre="UnidadMinera"; Tipo="Text"},
    @{Nombre="Herramienta"; Tipo="Text"},
    @{Nombre="Descripcion"; Tipo="Note"},
    @{Nombre="Marca"; Tipo="Text"},
    @{Nombre="Modelo"; Tipo="Text"},
    @{Nombre="Serie"; Tipo="Text"},
    @{Nombre="Color"; Tipo="Text"},
    @{Nombre="Aplicacion"; Tipo="Text"},
    @{Nombre="Ubicacion"; Tipo="Text"},
    @{Nombre="Estado"; Tipo="Text"},
    @{Nombre="Cantidad"; Tipo="Number"},
    @{Nombre="Unidad"; Tipo="Text"},
    @{Nombre="CodigoBarras1"; Tipo="Text"},
    @{Nombre="CodigoBarras2"; Tipo="Text"},
    @{Nombre="CodigoQR"; Tipo="Text"},
    @{Nombre="Foto"; Tipo="Note"},
    @{Nombre="Notas"; Tipo="Note"}
)

$existentes = Get-PnPField -List "Inventario" | Select-Object -ExpandProperty InternalName

foreach ($col in $columnas) {
    if ($existentes -contains $col.Nombre) {
        Write-Host "  -> $($col.Nombre) ya existe" -ForegroundColor Gray
        continue
    }
    try {
        Add-PnPField -List "Inventario" -DisplayName $col.Nombre -InternalName $col.Nombre -Type $col.Tipo | Out-Null
        Write-Host "  [+] $($col.Nombre) ($($col.Tipo))" -ForegroundColor Green
    } catch {
        Write-Host "  [!] $($col.Nombre): $($_.Exception.Message)" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "=============================================" -ForegroundColor Green
Write-Host "  LISTA Inventario LISTA" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green
Write-Host "  Total filas: $((Get-PnPListItem -List Inventario).Count)" -ForegroundColor Cyan

Disconnect-PnPOnline
