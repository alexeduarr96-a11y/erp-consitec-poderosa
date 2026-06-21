# =====================================================
# Agrega columnas que faltan en las listas SP
# Idempotente — si la columna ya existe, la omite
# Basado en auditoria de campos guardados por los modulos JS
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

# Mapa: lista -> array de columnas a asegurar
$columnasPorLista = @{
    "Cotizaciones" = @(
        @{Nombre="ClienteCorto"; Tipo="Text"},
        @{Nombre="ContactoConsitecId"; Tipo="Text"}
    )
    "Proyectos" = @(
        @{Nombre="Nombre"; Tipo="Text"}
    )
}

foreach ($lista in $columnasPorLista.Keys) {
    Write-Host ""
    Write-Host "Lista: $lista" -ForegroundColor Cyan
    # Obtener TODAS las columnas reales de la lista (chequeo estricto por InternalName)
    $allFields = Get-PnPField -List $lista | Select-Object -ExpandProperty InternalName
    foreach ($col in $columnasPorLista[$lista]) {
        if ($allFields -contains $col.Nombre) {
            Write-Host "  -> $($col.Nombre) ya existe" -ForegroundColor Gray
            continue
        }
        try {
            Add-PnPField -List $lista -DisplayName $col.Nombre -InternalName $col.Nombre -Type $col.Tipo | Out-Null
            Write-Host "  [+] $($col.Nombre) ($($col.Tipo)) agregada" -ForegroundColor Green
        } catch {
            Write-Host "  [!] $($col.Nombre): $($_.Exception.Message)" -ForegroundColor Yellow
        }
    }
}

Write-Host ""
Write-Host "=============================================" -ForegroundColor Green
Write-Host "  COLUMNAS COMPLETADAS" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green

Disconnect-PnPOnline
