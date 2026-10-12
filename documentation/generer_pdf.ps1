# Génère documentation.html (avec numéros de page du sommaire) puis GPIEXALT_Documentation.pdf.
# Les numéros de page viennent de pages.json (mesurés sur le PDF précédent) : relancer une seconde fois si la pagination change.
param([string]$PagesJson = "")
$ErrorActionPreference = 'Stop'
$dir = Split-Path -Parent $MyInvocation.MyCommand.Path
$tpl = [IO.File]::ReadAllText("$dir\documentation.template.html", [Text.Encoding]::UTF8)
$pages = @{}
$pj = if($PagesJson){ $PagesJson } else { "$dir\pages.json" }
if(Test-Path $pj){ (Get-Content $pj -Raw | ConvertFrom-Json).PSObject.Properties | ForEach-Object { $pages[$_.Name] = $_.Value } }
$html = [regex]::Replace($tpl, '\{\{P:(\w+)\}\}', { param($m) $k = $m.Groups[1].Value; if($pages.ContainsKey($k)){ [string]$pages[$k] } else { '0' } })
[IO.File]::WriteAllText("$dir\documentation.html", $html, (New-Object Text.UTF8Encoding($false)))
$edge = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
$out = "$dir\GPIEXALT_Documentation.pdf"
if([IO.File]::Exists($out)){ [IO.File]::Delete($out) }
$argList = @('--headless=new','--disable-gpu','--no-pdf-header-footer',"--user-data-dir=$env:TEMP\edge_pdf_prof","--print-to-pdf=`"$out`"",('file:///' + $dir.Replace('\','/').Replace(' ','%20') + '/documentation.html'))
Start-Process -FilePath $edge -ArgumentList $argList -Wait -WindowStyle Hidden
Start-Sleep -Seconds 3
"PDF : " + (Get-Item $out).Length + " octets"
