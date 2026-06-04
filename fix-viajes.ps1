$file = "C:\Users\rober\visa-global-web\index.html"
$content = [System.IO.File]::ReadAllText($file, [System.Text.Encoding]::UTF8)

# Corregir link del botón
$content = $content.Replace('wa.me/593987846751', 'wa.me/593994442512')

# Mejorar estilo del botón
$content = $content.Replace(
    'background:var(--gold,#F0B429);color:#060E1F;',
    'background:linear-gradient(135deg,#F5C842,#C8861A);color:#060E1C;'
)
$content = $content.Replace(
    'padding:0.95rem 2rem;border-radius:0.5rem;text-decoration:none;transition:background 0.2s,transform 0.15s;',
    'padding:16px 36px;border-radius:100px;text-decoration:none;box-shadow:0 8px 40px rgba(245,200,66,.25);'
)

# Número 01 — reemplazar caja de icono 1
$num1 = '<div style="flex-shrink:0;font-family:''Playfair Display'',serif;font-size:2rem;font-weight:900;color:rgba(245,200,66,.2);line-height:1;min-width:2.5rem;text-align:right;">01</div>'
$num2 = '<div style="flex-shrink:0;font-family:''Playfair Display'',serif;font-size:2rem;font-weight:900;color:rgba(245,200,66,.2);line-height:1;min-width:2.5rem;text-align:right;">02</div>'
$num3 = '<div style="flex-shrink:0;font-family:''Playfair Display'',serif;font-size:2rem;font-weight:900;color:rgba(245,200,66,.2);line-height:1;min-width:2.5rem;text-align:right;">03</div>'
$num4 = '<div style="flex-shrink:0;font-family:''Playfair Display'',serif;font-size:2rem;font-weight:900;color:rgba(245,200,66,.2);line-height:1;min-width:2.5rem;text-align:right;">04</div>'
$iconBox = '<div style="flex-shrink:0;width:2.2rem;height:2.2rem;background:rgba(240,180,41,0.12);border-radius:0.45rem;display:flex;align-items:center;justify-content:center;font-size:1rem;">'

# Contar cuántos iconos hay
$count = ([regex]::Matches($content, [regex]::Escape($iconBox))).Count
Write-Output "Iconos encontrados: $count"

# Reemplazar uno por uno con sus números
if ($count -ge 1) { $content = [regex]::Replace($content, [regex]::Escape($iconBox) + '[^<]+</div>', { param($m)
    $script:iconNum = if ($script:iconNum -eq $null) { 1 } else { $script:iconNum + 1 }
    switch ($script:iconNum) {
        1 { return $num1 }
        2 { return $num2 }
        3 { return $num3 }
        4 { return $num4 }
        default { return $m.Value }
    }
}) }

# Guardar
[System.IO.File]::WriteAllText($file, $content, [System.Text.Encoding]::UTF8)
Write-Output "Guardado. Conteo final iconos: $(([regex]::Matches($content, [regex]::Escape($iconBox))).Count)"
