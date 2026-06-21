# =====================================================
# Servidor HTTP local para el ERP CONSITEC / Poderosa
# Permite probar la integracion SharePoint sin instalar nada
# =====================================================
$port = 5500
$root = (Resolve-Path "$PSScriptRoot\..").Path

Write-Host ""
Write-Host "=============================================" -ForegroundColor Green
Write-Host "  Servidor ERP CONSITEC / Poderosa" -ForegroundColor Green
Write-Host "=============================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Servir desde: $root" -ForegroundColor Cyan
Write-Host "  Puerto:       $port" -ForegroundColor Cyan
Write-Host ""
Write-Host "  Abre estas URLs en tu navegador:" -ForegroundColor Yellow
Write-Host "  Test SharePoint: http://localhost:$port/sharepoint/test-conexion.html" -ForegroundColor White
Write-Host "  ERP completo:    http://localhost:$port/index.html" -ForegroundColor White
Write-Host ""
Write-Host "  Para DETENER: presiona Ctrl+C en esta ventana" -ForegroundColor Magenta
Write-Host "=============================================" -ForegroundColor Green
Write-Host ""

$listener = New-Object System.Net.HttpListener
$listener.Prefixes.Add("http://localhost:$port/")
try {
    $listener.Start()
} catch {
    Write-Host "ERROR: No se pudo iniciar el servidor en puerto $port" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""
    Write-Host "Posibles causas:" -ForegroundColor Yellow
    Write-Host "  - El puerto $port esta en uso (cierra otras apps)" -ForegroundColor White
    Write-Host "  - Firewall bloqueando (acepta el permiso si aparece)" -ForegroundColor White
    pause
    exit 1
}

$mimeTypes = @{
    ".html" = "text/html; charset=utf-8"
    ".htm"  = "text/html; charset=utf-8"
    ".js"   = "application/javascript; charset=utf-8"
    ".mjs"  = "application/javascript; charset=utf-8"
    ".css"  = "text/css; charset=utf-8"
    ".json" = "application/json; charset=utf-8"
    ".png"  = "image/png"
    ".jpg"  = "image/jpeg"
    ".jpeg" = "image/jpeg"
    ".gif"  = "image/gif"
    ".svg"  = "image/svg+xml"
    ".webp" = "image/webp"
    ".ico"  = "image/x-icon"
    ".pdf"  = "application/pdf"
    ".xml"  = "application/xml; charset=utf-8"
    ".txt"  = "text/plain; charset=utf-8"
    ".woff" = "font/woff"
    ".woff2"= "font/woff2"
    ".ttf"  = "font/ttf"
}

Write-Host "[OK] Servidor iniciado correctamente. Esperando solicitudes..." -ForegroundColor Green
Write-Host ""

# Loop principal
while ($listener.IsListening) {
    try {
        $context = $listener.GetContext()
        $req = $context.Request
        $res = $context.Response

        $path = [System.Net.WebUtility]::UrlDecode($req.Url.LocalPath)
        if ($path -eq "/") { $path = "/index.html" }

        # Sanitizar para evitar path traversal
        $relativePath = $path.TrimStart("/").Replace("/", [System.IO.Path]::DirectorySeparatorChar)
        $filePath = [System.IO.Path]::Combine($root, $relativePath)

        # Verificar que el path este dentro de $root
        $fullPath = [System.IO.Path]::GetFullPath($filePath)
        if (-not $fullPath.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) {
            $res.StatusCode = 403
            $body = [System.Text.Encoding]::UTF8.GetBytes("403 Forbidden")
            $res.OutputStream.Write($body, 0, $body.Length)
            $res.Close()
            Write-Host "[403] $path" -ForegroundColor Red
            continue
        }

        if (Test-Path $fullPath -PathType Leaf) {
            $bytes = [System.IO.File]::ReadAllBytes($fullPath)
            $ext = [System.IO.Path]::GetExtension($fullPath).ToLower()
            $mime = if ($mimeTypes.ContainsKey($ext)) { $mimeTypes[$ext] } else { "application/octet-stream" }
            $res.ContentType = $mime
            $res.ContentLength64 = $bytes.Length
            $res.Headers.Add("Cache-Control", "no-cache")
            $res.OutputStream.Write($bytes, 0, $bytes.Length)
            Write-Host "[200] $path ($($bytes.Length) bytes)" -ForegroundColor DarkGray
        } else {
            $res.StatusCode = 404
            $body = [System.Text.Encoding]::UTF8.GetBytes("404 - Not Found: $path")
            $res.ContentType = "text/plain; charset=utf-8"
            $res.OutputStream.Write($body, 0, $body.Length)
            Write-Host "[404] $path" -ForegroundColor Yellow
        }
        $res.Close()
    } catch {
        # Si Ctrl+C, salir limpiamente
        if ($_.Exception.GetType().Name -eq "HttpListenerException") { break }
        Write-Host "Error: $_" -ForegroundColor Red
    }
}

$listener.Stop()
Write-Host ""
Write-Host "Servidor detenido." -ForegroundColor Yellow
