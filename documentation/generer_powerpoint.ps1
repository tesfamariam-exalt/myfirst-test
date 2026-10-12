$ErrorActionPreference = 'Stop'
$outDir  = "C:\Users\TesfamariamYEBIO\OneDrive - EXALT\Bureau\app\documentation"
$pptxOut = Join-Path $outDir "GPIEXALT_Documentation.pptx"
$imgDoc  = Join-Path $outDir "images"
$imgDir  = Join-Path $env:TEMP "gpiexalt_slides"
if(Test-Path $imgDir){ Get-ChildItem $imgDir | ForEach-Object { [IO.File]::Delete($_.FullName) } } else { New-Item -ItemType Directory -Force $imgDir | Out-Null }
Get-Process POWERPNT -ErrorAction SilentlyContinue | Stop-Process -Force
Start-Sleep -Milliseconds 800

function Rgb([string]$h){ [int]([Convert]::ToInt32($h.Substring(0,2),16) + 256*[Convert]::ToInt32($h.Substring(2,2),16) + 65536*[Convert]::ToInt32($h.Substring(4,2),16)) }
function Nb([string]$t){ $nb = [string][char]0xA0; return ($t -replace '« ', ('«' + $nb)) -replace ' »', ($nb + '»') }

$NAVY='1E1C8C'; $INK='12162B'; $SOFT='586080'; $ACC='5853FF'; $BLUE='2F7DD1'; $GOOD='1B9E6B'; $WARN='B67608'; $BAD='C43D3D'
$BG='F5F6FB'; $LINE='DADDEE'; $PURPLE='7C4DBF'; $TEAL='0E8E82'
$TG='E3F5EC'; $TW='FDF0D5'; $TB='FBE4E4'; $TI='E4EEFA'

$ppt = New-Object -ComObject PowerPoint.Application
$pres = $ppt.Presentations.Add(-1)
$pres.PageSetup.SlideWidth = 960; $pres.PageSetup.SlideHeight = 540
$pres.SlideMaster.Theme.ThemeFontScheme.MajorFont.Item(1).Name = 'Calibri'
$pres.SlideMaster.Theme.ThemeFontScheme.MinorFont.Item(1).Name = 'Calibri'
$pres.SlideMaster.Background.Fill.Solid(); $pres.SlideMaster.Background.Fill.ForeColor.RGB = (Rgb $BG)
$layTitle = $pres.SlideMaster.CustomLayouts.Item(1)
$layOnly  = $pres.SlideMaster.CustomLayouts.Item(6)

function Shadow($b,$blur=14,$oy=4,$tr=0.84){ $b.Shadow.Visible = -1; $b.Shadow.ForeColor.RGB = (Rgb '1E1C8C'); $b.Shadow.Blur = $blur; $b.Shadow.OffsetX = 0; $b.Shadow.OffsetY = $oy; $b.Shadow.Transparency = $tr }
function Box($s,$x,$y,$w,$h,$fill,$shape=5,$line=$null,$radius=0.06){
  $b = $s.Shapes.AddShape($shape,$x,$y,$w,$h)
  $b.Fill.Solid(); $b.Fill.ForeColor.RGB = (Rgb $fill)
  if($line){ $b.Line.ForeColor.RGB = (Rgb $line); $b.Line.Weight = 1 } else { $b.Line.Visible = 0 }
  if($shape -eq 5){ $b.Adjustments.Item(1) = $radius }
  $b.Shadow.Visible = 0
  return $b
}
function Card($s,$x,$y,$w,$h,$fill='FFFFFF'){ $c = Box $s $x $y $w $h $fill 5 $LINE 0.045; Shadow $c; return $c }
function SetText($shape,$text,$size,$color,$bold=$false,$align=2,$anchor=3){
  $text = Nb $text
  $f = $shape.TextFrame; $f.WordWrap = -1; $f.AutoSize = 0
  $f.MarginLeft = 10; $f.MarginRight = 10; $f.MarginTop = 4; $f.MarginBottom = 4; $f.VerticalAnchor = $anchor
  $r = $f.TextRange; $r.Text = $text
  $r.Font.Name = 'Calibri'; $r.Font.Size = $size; $r.Font.Color.RGB = (Rgb $color)
  $r.Font.Bold = $(if($bold){-1}else{0}); $r.ParagraphFormat.Alignment = $align
}
function Txt($s,$x,$y,$w,$h,$text,$size=14,$color=$INK,$bold=$false,$align=1){
  $text = Nb $text
  $t = $s.Shapes.AddTextbox(1,$x,$y,$w,$h)
  $f = $t.TextFrame; $f.WordWrap = -1; $f.AutoSize = 0
  $f.MarginLeft = 0; $f.MarginRight = 0; $f.MarginTop = 0; $f.MarginBottom = 0
  $t.Height = $h
  $r = $f.TextRange; $r.Text = $text
  $r.Font.Name = 'Calibri'; $r.Font.Size = $size; $r.Font.Color.RGB = (Rgb $color)
  $r.Font.Bold = $(if($bold){-1}else{0}); $r.ParagraphFormat.Alignment = $align
  return $t
}
function Bul($s,$x,$y,$w,$h,$items,$size=15,$color=$INK,$after=7){
  $t = Txt $s $x $y $w $h ($items -join "`r") $size $color
  $r = $t.TextFrame.TextRange
  $r.ParagraphFormat.Bullet.Visible = -1; $r.ParagraphFormat.Bullet.Character = 8226
  $r.ParagraphFormat.Bullet.Font.Color.RGB = (Rgb $ACC)
  $r.ParagraphFormat.SpaceAfter = $after
  $t.TextFrame.Ruler.Levels.Item(1).FirstMargin = 0; $t.TextFrame.Ruler.Levels.Item(1).LeftMargin = 16
  return $t
}
function Circle($s,$x,$y,$d,$fill,$label,$size=16){
  $c = Box $s $x $y $d $d $fill 9
  SetText $c $label $size 'FFFFFF' $true 2 3
  $c.TextFrame.MarginLeft = 0; $c.TextFrame.MarginRight = 0; $c.TextFrame.MarginTop = 0; $c.TextFrame.MarginBottom = 0
  return $c
}
function Arrow($s,$x1,$y1,$x2,$y2,$color=$SOFT){
  $l = $s.Shapes.AddLine($x1,$y1,$x2,$y2)
  $l.Line.ForeColor.RGB = (Rgb $color); $l.Line.Weight = 2; $l.Line.EndArrowheadStyle = 2
  return $l
}
function Pill($s,$x,$y,$w,$text,$fill,$color,$h=24,$size=12){
  $p = Box $s $x $y $w $h $fill 5 $null 0.5
  SetText $p $text $size $color $true 2 3
  return $p
}
function Btn($s,$x,$y,$w,$text,$fill,$color,$line=$null,$h=30){
  $p = Box $s $x $y $w $h $fill 5 $line 0.25
  SetText $p $text 13 $color $true 2 3
  return $p
}
function Prism($s,[double]$x,[double]$y,[double]$k,[double]$tr1=0,[double]$tr2=0){
  $null = $s.Shapes.AddPicture((Join-Path $imgDoc 'exalt_x_soft.png'), 0, -1, $x, $y, (100*$k), (100*$k))
}
function MakeTable($s,$x,$y,$w,$colW,$rows,$size=12,$rowH=26,$centerFrom=0){
  $nr = $rows.Count; $nc = $rows[0].Count
  $sh = $s.Shapes.AddTable($nr,$nc,$x,$y,$w,$nr*$rowH); $t = $sh.Table
  $t.ApplyStyle('{2D5ABB26-0587-4C30-8999-92F81FD0307C}')
  for($c=1;$c -le $nc;$c++){ $t.Columns.Item($c).Width = $colW[$c-1] }
  for($r=1;$r -le $nr;$r++){
    for($c=1;$c -le $nc;$c++){
      $cell = $t.Cell($r,$c); $f = $cell.Shape.TextFrame; $tr = $f.TextRange
      $tr.Text = (Nb ([string]$rows[$r-1][$c-1]))
      $tr.Font.Name = 'Calibri'; $tr.Font.Size = $size
      $f.MarginLeft = 10; $f.MarginRight = 10; $f.MarginTop = 3; $f.MarginBottom = 3; $f.VerticalAnchor = 3
      $tr.ParagraphFormat.Alignment = $(if($centerFrom -gt 0 -and $c -ge $centerFrom){2}else{1})
      $cell.Shape.Fill.Solid()
      if($r -eq 1){ $cell.Shape.Fill.ForeColor.RGB = (Rgb $NAVY); $tr.Font.Color.RGB = (Rgb 'FFFFFF'); $tr.Font.Bold = -1 }
      else { $cell.Shape.Fill.ForeColor.RGB = (Rgb $(if($r % 2 -eq 0){'FFFFFF'}else{'F0F2FA'})); $tr.Font.Color.RGB = (Rgb $INK); $tr.Font.Bold = $(if($c -eq 1){-1}else{0}) }
    }
    $t.Rows.Item($r).Height = $rowH
  }
  return $sh
}
# Capture d'écran dans un cadre de navigateur, avec repères numérotés (coordonnées d'affichage de la capture)
function Shot($s,$file,$x,$y,$w,$origW,$origH,$dispW,$markers){
  $h = $w * $origH / $origW
  $fr = Box $s $x $y $w ($h+22) 'FFFFFF' 5 $LINE 0.02; Shadow $fr 18 5 0.8
  $bar = Box $s $x $y $w 22 'ECEEF8' 5 $LINE 0.02
  $d1 = Box $s ($x+10) ($y+7) 8 8 'F26B5E' 9; $d2 = Box $s ($x+23) ($y+7) 8 8 'F5C04A' 9; $d3 = Box $s ($x+36) ($y+7) 8 8 '5BC77A' 9
  $url = Box $s ($x+$w/2-90) ($y+4) 180 14 'FFFFFF' 5 $null 0.5
  SetText $url 'GPIEXALT' 9 $SOFT $false 2 3; $url.TextFrame.MarginTop = 0; $url.TextFrame.MarginBottom = 0
  $p = $s.Shapes.AddPicture((Join-Path $imgDoc $file), 0, -1, $x, $y+22, $w, $h)
  $k = $w / $dispW
  foreach($m in $markers){
    $cx = $x + $m[1]*$k; $cy = $y + 22 + $m[2]*$k
    $c = Circle $s ($cx-10) ($cy-10) 20 $ACC ([string]$m[0]) 11
    $c.Line.Visible = -1; $c.Line.ForeColor.RGB = (Rgb 'FFFFFF'); $c.Line.Weight = 1.5
  }
  return ($y + 22 + $h)
}
function ShotAuto($s,$file,$x,$y,$w,$dispW,$markers){
  Add-Type -AssemblyName System.Drawing
  $im = [System.Drawing.Image]::FromFile((Join-Path $imgDoc $file)); $ow = $im.Width; $oh = $im.Height; $im.Dispose()
  return (Shot $s $file $x $y $w $ow $oh $dispW $markers)
}
function Legend($s,$x,$y,$w,$items,$size=14,$gap=34){
  $i = 0
  foreach($it in $items){
    $yy = $y + $i*$gap
    $null = Circle $s $x $yy 22 $ACC ([string]($i+1)) 12
    $null = Txt $s ($x+32) ($yy+1) ($w-32) ($gap-4) $it $size $INK
    $i++
  }
}
function NewSlide($kicker,$title,$sub,$notes){
  $s = $pres.Slides.AddSlide($pres.Slides.Count + 1, $layOnly)
  $t = $s.Shapes.Title
  $t.Left = 48; $t.Top = 40; $t.Width = 864; $t.Height = 44
  $t.TextFrame.TextRange.Text = (Nb $title)
  $t.TextFrame.TextRange.Font.Name = 'Calibri'; $t.TextFrame.TextRange.Font.Size = 30; $t.TextFrame.TextRange.Font.Bold = -1
  $t.TextFrame.TextRange.Font.Color.RGB = (Rgb $NAVY); $t.TextFrame.TextRange.ParagraphFormat.Alignment = 1
  $t.TextFrame.VerticalAnchor = 3; $t.TextFrame.MarginLeft = 0
  $k = Txt $s 48 24 600 16 $kicker.ToUpper() 11 $ACC $true
  $k.TextFrame2.TextRange.Font.Spacing = 1.5
  if($sub){ $null = Txt $s 48 88 864 22 $sub 15 $SOFT }
  $null = Txt $s 48 512 700 14 "GPIEXALT  ·  DOC-GPIEXALT-001  ·  Version 7.0  ·  12/10/2026  ·  Classification : Interne" 9 $SOFT
  $num = Txt $s 812 512 100 14 "" 9 $SOFT $false 3
  $num.TextFrame.TextRange.InsertSlideNumber() | Out-Null
  if($notes){ $s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = $notes }
  return $s
}
function Divider($num,$title,$sub){
  $s = $pres.Slides.AddSlide($pres.Slides.Count + 1, $layOnly)
  $s.FollowMasterBackground = 0; $s.Background.Fill.Solid(); $s.Background.Fill.ForeColor.RGB = (Rgb $NAVY)
  $t = $s.Shapes.Title
  $t.Left = 96; $t.Top = 236; $t.Width = 620; $t.Height = 70
  $t.TextFrame.TextRange.Text = (Nb $title)
  $t.TextFrame.TextRange.Font.Name = 'Calibri'; $t.TextFrame.TextRange.Font.Size = 44; $t.TextFrame.TextRange.Font.Bold = -1
  $t.TextFrame.TextRange.Font.Color.RGB = (Rgb 'FFFFFF'); $t.TextFrame.TextRange.ParagraphFormat.Alignment = 1; $t.TextFrame.VerticalAnchor = 3; $t.TextFrame.MarginLeft = 0
  $null = Txt $s 96 150 300 80 $num 80 $ACC $true
  $null = Txt $s 96 316 620 60 $sub 20 'AEB9E3'
  Prism $s 660 130 2.6
  $null = Txt $s 96 500 700 14 "GPIEXALT  ·  DOC-GPIEXALT-001  ·  Version 7.0" 9 '7C86B8'
  return $s
}

# ================= 1. Couverture =================
$s = $pres.Slides.AddSlide(1, $layTitle)
$s.FollowMasterBackground = 0; $s.Background.Fill.Solid(); $s.Background.Fill.ForeColor.RGB = (Rgb $NAVY)
Prism $s 640 100 2.9
$t1 = $s.Shapes.Item(1); $t2 = $s.Shapes.Item(2)
$t1.Left = 96; $t1.Top = 200; $t1.Width = 560; $t1.Height = 80
$t1.TextFrame.TextRange.Text = 'GPIEXALT'; $t1.TextFrame.TextRange.Font.Name='Calibri'; $t1.TextFrame.TextRange.Font.Size = 66; $t1.TextFrame.TextRange.Font.Bold = -1
$t1.TextFrame.TextRange.Font.Color.RGB = (Rgb 'FFFFFF'); $t1.TextFrame.TextRange.ParagraphFormat.Alignment = 1; $t1.TextFrame.VerticalAnchor = 3; $t1.TextFrame.MarginLeft = 0
$t2.Left = 96; $t2.Top = 286; $t2.Width = 520; $t2.Height = 80
$t2.TextFrame.TextRange.Text = "Dossier de documentation : gestion du parc informatique eXalt, prêts, restitutions et attestations de matériel"
$t2.TextFrame.TextRange.Font.Name='Calibri'; $t2.TextFrame.TextRange.Font.Size = 22; $t2.TextFrame.TextRange.Font.Color.RGB = (Rgb 'CADCFC'); $t2.TextFrame.TextRange.ParagraphFormat.Alignment = 1; $t2.TextFrame.VerticalAnchor = 1; $t2.TextFrame.MarginLeft = 0
$k = Txt $s 96 168 600 18 'DOCUMENTATION FONCTIONNELLE, TECHNIQUE ET D''EXPLOITATION' 12 'B9B6FF' $true
$k.TextFrame2.TextRange.Font.Spacing = 1.5
$null = Txt $s 96 430 620 18 "DOC-GPIEXALT-001  ·  Version 7.0  ·  12/10/2026  ·  Application v167" 13 'AEB9E3'
$null = Txt $s 96 452 620 18 "Propriétaire : Service informatique (DSI)  ·  Classification : Interne  ·  Statut : diffusé pour validation" 13 'AEB9E3'
$s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = "Présentation de GPIEXALT (Gestion du Parc Informatique eXalt, anciennement Prisme) : fonctions, procédures, règles de gestion, architecture, sécurité, sauvegarde, recette et risques. Version 6.0 : thème et logo eXalt, retour prévu et alertes de retard, actions groupées, scan rapide, rapport par email, destinataire au choix. Version 7.0 : rapport propre à chaque page, fiche d'emprunt réorganisée."

# ================= 2. Contrôle du document =================
$s = NewSlide 'Gouvernance' 'Contrôle du document' 'Référence, version, propriétaire et circuit de validation' "Les noms du vérificateur et de l'approbateur restent à renseigner."
$null = Card $s 48 126 420 366
$null = Txt $s 68 140 380 22 'Informations générales' 17 $NAVY $true
$null = MakeTable $s 62 170 392 @(120,272) @(
  @('Rubrique','Valeur'),@('Référence','DOC-GPIEXALT-001'),@('Version','7.0 — 12/10/2026'),@('Application','GPIEXALT, version 167'),
  @('Propriétaire','Service informatique (DSI)'),@('Classification','Interne'),@('Statut','Diffusé pour validation'),@('Diffusion','Équipe IT, DSI ; RSSI et DPO (sécurité, données)')) 12 36
$null = Card $s 492 126 420 214
$null = Txt $s 512 138 380 22 'Historique des révisions' 17 $NAVY $true
$null = MakeTable $s 506 168 392 @(56,82,254) @(@('Version','Date','Nature'),@('1.0','05/10/2026','Première édition'),@('2.0','05/10/2026','Édition entreprise'),@('3.0','06/10/2026','App 106 : accueil, QR, signature'),@('4.0','07/10/2026','App 128 : rôles, rapports, recherche'),@('5.0','08/10/2026','App 147 : annuaire, envoi signé'),@('6.0','09/10/2026','App 161 : GPIEXALT, retards, rapport'),@('7.0','12/10/2026','App 167 : rapport par page')) 10 21
$null = Card $s 492 354 420 138
$null = Txt $s 512 364 380 22 'Circuit de validation' 17 $NAVY $true
$null = MakeTable $s 506 392 392 @(100,190,102) @(@('Étape','Nom et fonction','Date'),@('Rédaction','Service informatique','12/10/2026'),@('Vérification','À renseigner',''),@('Approbation','À renseigner (DSI)','')) 12 24

# ================= 3. Sommaire =================
$s = NewSlide 'Gouvernance' 'Sommaire' 'Sept parties, de la vue d''ensemble aux actions attendues' "Le dossier PDF détaille chaque partie (71 pages, procédures, plan de recette, registre des risques)."
$sm = @(
  @('01','Vue d''ensemble','Finalité, cycle de vie, les sept onglets'),
  @('02','Fonctions','Prêts, retours, attestations, archives, QR, recherche, rapports, rôles, annuaire, thèmes'),
  @('03','Gouvernance','Rôles et responsabilités, règles de gestion'),
  @('04','Architecture et sécurité','Composants, accès, données personnelles'),
  @('05','Sauvegarde','Où, comment, en cas d''échec'),
  @('06','Qualité','Plan de recette, risques'),
  @('07','Exploitation','Mises à jour, limites, actions attendues')
)
for($i=0;$i -lt 7;$i++){
  $y = 128 + $i*52
  $null = Txt $s 48 ($y+2) 70 40 $sm[$i][0] 30 $ACC $true
  $null = Txt $s 130 ($y+2) 300 24 $sm[$i][1] 19 $NAVY $true
  $null = Txt $s 130 ($y+26) 520 20 $sm[$i][2] 13 $SOFT
}
$cd = Card $s 660 128 252 360 $NAVY
$null = Txt $s 682 146 210 20 'EN BREF' 11 'B9B6FF' $true
$null = Txt $s 682 172 210 300 "Un outil pour savoir qui détient quel matériel, prouver ce qui a été restitué par une attestation datée et signée, et suivre les pertes.`r`rPC, tablette, téléphone.`rChaque action est tracée." 15 'FFFFFF'

# ================= 4. Divider 01 =================
$null = Divider '01' 'Vue d''ensemble' 'Finalité, cycle de vie et onglets'

# ================= 5. À quoi sert =================
$s = NewSlide 'Vue d''ensemble' 'À quoi sert GPIEXALT' 'Trois questions, une seule application' "Qui a quel matériel, qu'est-ce qui a été rendu (avec preuve), qu'est-ce qui a été perdu."
$cards = @(@('Qui a quoi',"Chaque équipement prêté est suivi, personne par personne, avec la date et l'auteur de l'enregistrement.",$BLUE),
  @('Preuve de restitution',"Une attestation PDF datée et signée, envoyée par email ou par téléphone, puis archivée.",$GOOD),
  @('Pertes sous contrôle',"Le matériel perdu, volé ou abîmé est déclaré et compté par type d'équipement.",$BAD))
for($i=0;$i -lt 3;$i++){
  $x = 48 + $i*297
  $null = Card $s $x 130 270 250
  $null = Circle $s ($x+22) 150 52 $cards[$i][2] ([string]($i+1)) 24
  $null = Txt $s ($x+22) 222 230 28 $cards[$i][0] 21 $NAVY $true
  $null = Txt $s ($x+22) 258 228 110 $cards[$i][1] 15 $INK
}
$bar = Box $s 48 400 864 76 $NAVY
SetText $bar "PC, tablette et téléphone   ·   Données sauvegardées et partagées   ·   Chaque action est tracée" 18 'FFFFFF' $false 2 3

# ================= 6. Cycle =================
$s = NewSlide 'Vue d''ensemble' 'Le cycle de vie d''un matériel' 'Quatre étapes et une variante' "La variante perdu / volé / abîmé remplace l'attestation et alimente l'onglet Matériel perdu."
$steps = @(@('1','Emprunter','Le prêt est enregistré',$BLUE),@('2','Restituer','Envoi par email ou téléphone',$PURPLE),@('3','Attestation','« En attente », puis relances',$WARN),@('4','Signée','PDF daté, signé et archivé',$GOOD))
for($i=0;$i -lt 4;$i++){
  $x = 48 + $i*224.67
  $null = Card $s $x 132 190 150
  $null = Circle $s ($x+16) 148 38 $steps[$i][3] $steps[$i][0] 17
  $null = Txt $s ($x+16) 196 160 28 $steps[$i][1] 22 $NAVY $true
  $null = Txt $s ($x+16) 230 162 44 $steps[$i][2] 14 $SOFT
  if($i -lt 3){ $null = Arrow $s ($x+194) 207 ($x+220) 207 }
}
$null = Arrow $s 343 284 343 334 $BAD
$v = Box $s 183 336 320 70 $TB 5 $BAD
SetText $v "Variante : perdu, volé ou abîmé`r→ onglet « Matériel perdu »" 16 $BAD $true 2 3
$q = Box $s 48 430 864 52 'FFFFFF' 5 $LINE
SetText $q "À chaque étape, l'application enregistre qui a fait l'action et quand : prêt, envoi, relance, signature, déclaration." 15 $INK $false 1 3
$q.TextFrame.MarginLeft = 18

# ================= 7. Onglets =================
$s = NewSlide 'Vue d''ensemble' 'Les sept onglets' 'Piloter, opérer, suivre' "Sur PC : menu latéral repliable ; sur téléphone et tablette en portrait : barre d'onglets en bas. Raccourcis 1 à 7."
$tabs = @(@('Accueil','Ce qui demande une action, répartition du parc',$TEAL),@('Emprunter','Enregistrer les équipements remis à une personne',$BLUE),@('Restituer','Attestation de retour, ou perte, vol, casse',$PURPLE),@('Matériel perdu','Synthèse du matériel perdu, volé ou abîmé',$BAD),@('Attestations','Suivre les envois : en attente, à relancer, signées',$WARN),@('Archives PDF','Retrouver le PDF signé de chaque attestation',$GOOD),@('Activités récentes','Journal de tout ce qui a été enregistré','8A93C0'))
for($i=0;$i -lt 7;$i++){
  $col = $i % 4; $row = [math]::Floor($i / 4)
  $x = 48 + $col*219; $y = 124 + $row*136
  $null = Card $s $x $y 207 124
  $null = Circle $s ($x+14) ($y+12) 30 $tabs[$i][2] ([string]($i+1)) 14
  $null = Txt $s ($x+54) ($y+17) 148 22 $tabs[$i][0] 16 $NAVY $true
  $null = Txt $s ($x+14) ($y+54) 182 62 $tabs[$i][1] 13 $INK
}
$null = Card $s 486 260 426 124 $NAVY
$null = Txt $s 506 272 390 20 'AUSSI DANS L''APPLICATION' 11 'B9B6FF' $true
$null = Txt $s 506 296 390 84 "Recherche rapide · Rapport PDF ou Excel de chaque page (envoi par email) · Alertes de retard · Rôles et accès · Annuaire Microsoft 365 · Filtres par période · Scan rapide · Thèmes eXalt" 15 'FFFFFF'
$null = Card $s 48 398 864 92
$null = Txt $s 68 408 400 22 'Aides à la navigation' 17 $NAVY $true
$null = Bul $s 68 434 400 54 @("Raccourcis : Ctrl K, 1 à 7 (onglets), / (liste), ? (aide), Échap","Recherche insensible aux accents et à la casse") 13 $INK 3
$null = Bul $s 492 434 400 54 @("Douze thèmes (dont eXalt) et sept couleurs d'accent","Pastille de sauvegarde : « Synchronisé » ou « Mode local »") 13 $INK 3

# ================= 8. Divider 02 =================
$null = Divider '02' 'Fonctions' 'Chaque écran, chaque bouton, chaque donnée enregistrée'

# ================= 9. Accueil =================
$s = NewSlide 'Onglet 1' 'Accueil : le tableau de bord' 'Ce qui demande une action, en un coup d''œil' "Les indicateurs (dont « Retours en retard », visible seulement s'il y en a) ouvrent l'onglet concerné avec le filtre déjà appliqué. Le bandeau « PDF à archiver » n'apparaît que s'il y en a. Au survol des couleurs : libellé et nombre. La date et le lieu s'affichent en haut à droite de chaque page."
$null = ShotAuto $s 'home.png' 48 126 410 1440 @(@(1,306,88),@(2,1246,162),@(3,306,299),@(4,306,432),@(5,306,605),@(6,857,605),@(7,306,837))
$null = Card $s 500 126 412 366
$null = Txt $s 518 138 376 22 'Repères' 17 $NAVY $true
Legend $s 518 168 376 @('Phrase de situation et voyant (vert, orange, rouge)','Accès rapides : recherche, prêt, retour, QR, rapport','Indicateurs cliquables, dont Retours en retard','Bandeau « PDF à archiver » (si besoin)','Parc matériel : chez les personnes, en retard, en attente, perdu','Attestations : signées, en attente, à relancer','Statistiques par type de matériel') 13 36

# ================= 10. Emprunter =================
$s = NewSlide 'Onglet 2' 'Emprunter : enregistrer un prêt' 'Rechercher la personne, cocher le matériel, choisir la date, enregistrer' "La fiche s'ouvre dans une fenêtre avec trois étapes. Un équipement peut être emprunté et restitué plusieurs fois. Annuler (10 secondes) supprime l'emprunt qui vient d'être saisi."
$null = ShotAuto $s 'loan.png' 48 126 552 1440 @(@(1,306,299),@(2,335,407),@(3,306,482),@(4,306,531),@(5,1192,482),@(6,1094,635),@(7,1336,635))
$null = Card $s 636 126 276 366
$null = Txt $s 654 138 240 22 'Repères' 17 $NAVY $true
Legend $s 654 168 240 @('Quatre indicateurs','Recherche (accents ignorés)','Filtres : à emprunter, tout emprunté','Période (entre deux dates…)','Tri','Jauge « x / y remis »','État de la personne') 13 33
$null = Txt $s 654 408 240 18 'Dans la fiche' 14 $NAVY $true
$null = Txt $s 654 428 240 62 "3 étapes ; carte « Retour prévu » facultative, bouton Enregistrer en bas à droite. Annuler pendant 10 s." 12 $SOFT

# ================= 11. Restituer : le suivi =================
$s = NewSlide 'Onglet 3' 'Restituer : suivre les retours' 'Indicateurs, filtres et jauge à quatre couleurs' "La jauge de chaque ligne distingue : restitué et signé (vert), encore chez la personne (bleu), en attente de signature (jaune), perdu volé ou abîmé (rouge)."
$null = ShotAuto $s 'returnlist.png' 48 126 552 1440 @(@(1,306,322),@(2,1185,428),@(3,306,504),@(4,1087,666),@(5,1339,658))
$null = Card $s 636 126 276 366
$null = Txt $s 654 138 240 22 'Repères' 17 $NAVY $true
Legend $s 654 168 240 @('Quatre indicateurs cliquables','Recherche et scanner QR','Filtres : à traiter, en retard de retour, en attente, perdu…','Jauge : restitué, chez la personne, en attente, perdu','Pastille d''état de la ligne') 13 36
$null = Txt $s 654 372 240 18 'Priorité d''une ligne' 14 $NAVY $true
$null = Txt $s 654 392 240 96 "Le bord gauche suit la priorité : perte (rouge), attente de signature (jaune), en cours (bleu), tout rendu (vert). « Retour en retard » en rouge sur la ligne." 12 $SOFT

# ================= 12. Restituer : actions =================
$s = NewSlide 'Onglet 3' 'Restituer : trois actions possibles' 'Attestation par email ou téléphone, ou déclaration de perte, vol, casse' "Rien n'est envoyé automatiquement : la personne connectée envoie elle-même le message préparé. Par email : un aperçu du PDF s'affiche d'abord. Les déclarations sont irréversibles (confirmation demandée)."
$bottom = ShotAuto $s 'return.png' 48 124 272 920 @(@(1,22,129),@(2,22,336),@(3,22,385),@(4,22,765),@(5,22,909),@(6,22,1020),@(7,810,129))
$null = Card $s 432 124 480 368
$null = Txt $s 452 136 440 22 'Repères' 17 $NAVY $true
Legend $s 452 168 440 @('Compteurs : restitués, à signer, perdus','Étapes : matériel, date de retour, envoi','Matériel emprunté : cases et quantités','Date de retour, « Sélectionner tout »','Par email / Par téléphone : attestation en attente','Déclarer perdu / volé / abîmé (confirmation)','Recherche d''une autre personne') 13 29
$null = Txt $s 452 376 440 20 'Résultat' 15 $NAVY $true
$null = Bul $s 452 398 440 92 @("Par email : aperçu du PDF, choix du destinataire (adresse enregistrée ou saisie à la main), puis brouillon Outlook, PDF joint","Par téléphone : SMS préparé, numéro de test 06 95 17 58 74","Déclaration : sort de chez la personne, visible dans Matériel perdu") 13 $INK 3

# ================= Retour prévu et alertes de retard =================
$s = NewSlide 'Onglets 2 et 3' 'Date de retour prévue et alertes de retard' 'Facultative à l''emprunt, surveillée ensuite : fenêtre, carte, pastilles, filtre' "La date de retour prévue est facultative. Dépassée, elle déclenche une fenêtre d'alerte à l'ouverture (au plus une fois par jour, ou si le nombre augmente), la carte rouge Retours en retard, les pastilles du menu et le filtre En retard de retour. Aucune notification n'est envoyée à l'extérieur : l'alerte n'apparaît que lorsque quelqu'un ouvre GPIEXALT."
$null = ShotAuto $s 'duedate.png' 48 126 330 910 @()
$null = Card $s 48 372 330 120
$null = Txt $s 62 382 304 104 "Carte « Retour prévu le » (facultatif) sous la date de remise : pastilles + 1 semaine, + 1 mois, + 3 mois, Effacer. Une date antérieure à l'emprunt est refusée. Le bouton Enregistrer est en bas à droite de la carte." 13 $INK
$null = ShotAuto $s 'latealert.png' 396 126 250 570 @()
$wl = Box $s 396 352 250 140 $TW 5 $WARN
SetText $wl "Limite : pas de notification automatique. L'alerte n'apparaît que lorsque quelqu'un ouvre l'application (risque R-19)." 13 $INK $false 1 3
$wl.TextFrame.MarginLeft = 14
$null = Card $s 668 126 244 366
$null = Txt $s 684 138 212 22 'Où voir les retards' 17 $NAVY $true
$null = Bul $s 684 168 214 320 @("Fenêtre à l'ouverture : personnes, équipements, jours de retard","Accueil : carte rouge, phrase de situation, segment violet dans le parc","Menu : pastille Restituer en rouge","Restituer : filtre « En retard de retour », personnes en retard en tête") 13 $INK 8

# ================= 13. Attestations =================
$s = NewSlide 'Onglet 5' 'Attestations : suivi et cycle de vie' 'Toutes les preuves, avec leur statut et leurs actions' "Chaque ligne affiche un cycle de vie : Envoyée, Signée, PDF archivé. En attente : jours d'attente, rouge à partir de 3 jours (REMINDER_AFTER_DAYS). Le bouton Relancer est masqué par défaut (SHOW_REMIND_BUTTON)."
$null = ShotAuto $s 'attest.png' 48 126 570 1440 @(@(1,306,338),@(2,306,524),@(3,905,524),@(4,425,739),@(5,942,700),@(6,1087,700),@(7,1305,700),@(8,1260,578),@(9,310,618))
$null = Card $s 668 126 244 366
$null = Txt $s 686 138 210 22 'Repères' 17 $NAVY $true
Legend $s 686 168 210 @('Indicateurs : envoyées, en attente, à relancer, délai moyen','Filtres : à relancer, signées…','Période (date d''envoi)','Cycle de vie : envoyée, signée, archivée','En attente · N j (rouge dès 3 jours)','Marquer comme signée','PDF de l''attestation','Exporter CSV','Sélection multiple (cases à cocher)') 12 33

# ================= Actions groupées =================
$s = NewSlide 'Onglet 5' 'Attestations : actions groupées' 'Cocher plusieurs attestations, puis les signer ou les exporter d''un coup' "Une case par ligne et « Sélectionner toutes les attestations en attente de cette liste » (suit la recherche, les filtres et la période). La barre d'actions apparaît en bas de la liste. La signature groupée exige une date (jamais dans le futur) et une mention écrite ; elle archive un PDF par attestation et n'envoie aucun message."
$null = ShotAuto $s 'bulk.png' 48 126 520 1440 @()
$null = ShotAuto $s 'bulkmodal.png' 592 126 320 500 @()
$null = Card $s 48 480 520 28
$null = Txt $s 58 485 500 20 "Marquer N comme signées  ·  Exporter la sélection (CSV)  ·  Désélectionner" 12 $INK

# ================= 14. Signature =================
$s = NewSlide 'Onglet 5' '« Marquer comme signée » pas à pas' 'Une date, une signature, une mention, un PDF archivé' "La date ne peut pas être dans le futur. Un tracé (doigt, stylet, souris) ou une mention écrite est obligatoire. La mention (140 caractères) permet de signer à la place de quelqu'un. Annuler pendant 10 secondes."
$st = @(@('1','Clic sur le bouton','Dans Attestations ou dans la fiche de la personne'),@('2','Date, et signature ou mention','Un tracé ou une mention est obligatoire : le bouton Valider reste grisé sinon'),@('3','PDF signé généré','Signature et mention dans le PDF ; « DSI » côté IT'),@('4','Archivage, message et Annuler','PDF rangé dans Archives ; envoi au collaborateur proposé ; Annuler 10 s'))
for($i=0;$i -lt 4;$i++){
  $y = 128 + $i*72
  $null = Circle $s 48 $y 40 $GOOD $st[$i][0] 18
  $null = Txt $s 102 ($y-1) 330 24 $st[$i][1] 18 $NAVY $true
  $null = Txt $s 102 ($y+24) 320 40 $st[$i][2] 13 $SOFT
}
$w1 = Box $s 48 424 400 66 $TW 5 $WARN
SetText $w1 "Si l'archivage échoue, la signature est quand même enregistrée : « Archiver le PDF » permet de recommencer." 13 $INK $false 1 3
$null = ShotAuto $s 'sign.png' 516 124 215 460 @(@(1,14,200),@(2,14,373),@(3,375,291),@(4,14,502),@(5,280,502),@(6,14,631),@(7,179,702),@(8,270,702))
$null = Card $s 772 124 140 366
$null = Txt $s 782 132 120 18 'Repères' 13 $NAVY $true
Legend $s 784 158 126 @('Date','Signature','Effacer','Mention','Pré-remplir','Message d''aide','Annuler','Valider') 11 30

# ================= Envoi de l'attestation signée =================
$s = NewSlide 'Onglet 5' 'Envoyer l''attestation signée par email' 'Après la signature : un message prêt à envoyer, avec le PDF signé joint' "Le message part de la boîte Outlook de l'utilisateur par le connecteur Microsoft 365, après un clic sur Envoyer maintenant. Le destinataire est l'adresse de l'annuaire : à contrôler dans la fenêtre. L'envoi réel reste à éprouver (CR-37)."
$null = ShotAuto $s 'sendmail.png' 48 124 250 580 @()
$null = Card $s 330 126 582 176
$null = Txt $s 350 138 540 22 'Quatre choix dans la fenêtre' 18 $NAVY $true
$null = Bul $s 350 168 546 130 @("Envoyer maintenant : le message part avec le PDF signé joint, copie dans les éléments envoyés","Modifier : l'objet et le texte (pas le destinataire ni le PDF)","Créer un brouillon : dossier Brouillons d'Outlook","Ne pas envoyer : la signature reste enregistrée") 14 $INK 4
$null = Card $s 330 316 582 176 $TW
$null = Txt $s 350 328 540 22 'À savoir' 18 $WARN $true
$null = Bul $s 350 358 546 130 @("Rien ne part sans clic ; la date d'envoi est tracée sur l'attestation","Sans connecteur : fichier .eml à ouvrir dans Outlook","Outlook sur le web ne s'ouvre pas depuis la page : brouillon ou .eml","Premier envoi réel à contrôler avec un destinataire de test") 14 $INK 4

# ================= 15. Aperçu et annulation =================
$s = NewSlide 'Confort' 'Aperçu du PDF, destinataire et bouton Annuler' 'Vérifier avant d''envoyer, corriger tout de suite' "L'aperçu utilise la bibliothèque pdf.js chargée à la demande. L'annulation est refusée si une autre action a été enregistrée depuis pour la même personne."
$null = ShotAuto $s 'rcp.png' 48 126 215 760 @()
$null = Card $s 290 126 622 176
$null = Txt $s 308 138 586 22 'Aperçu du PDF et destinataire du message' 18 $BLUE $true
$null = Bul $s 308 168 586 130 @("Par email : le PDF s'affiche tel qu'il sera joint","Destinataire : adresse enregistrée, ou autre adresse saisie à la main (personne sans compte professionnel) ; conservée avec l'attestation","« Continuer et envoyer » enregistre et prépare le brouillon ; « Retour » ne change rien") 13 $INK 4
$null = Card $s 290 316 622 176 $TW
$null = Txt $s 308 328 586 22 'Annuler une saisie qui vient d''être faite' 18 $WARN $true
$null = Bul $s 308 358 586 130 @("Après un emprunt ou une signature, le message propose « Annuler » pendant 10 secondes au moins","Emprunt supprimé ; attestation « En attente », archive et mention supprimées","Refusé si une autre action a été enregistrée depuis pour la personne","Au-delà : commentaire et administrateur") 13 $INK 4

# ================= 16. PDF =================
$s = NewSlide 'Le document' 'Le PDF d''attestation' 'Deux versions du même document, deux cases toujours remplies' "Le PDF est créé dans le navigateur, une page A4. Version émise (à l'envoi) et version signée (après Marquer comme signée) ; la mention de signature figure dans le bloc de suivi et sous la case du collaborateur."
$pg = Box $s 48 126 262 366 'FFFFFF' 1 $LINE; Shadow $pg 16 4 0.78
$null = $s.Shapes.AddPicture((Join-Path $imgDoc 'pdfsigned.png'), 0, -1, 52, 130, 254, 359)
$null = Card $s 330 126 270 170
$null = Txt $s 348 138 234 24 'Émise' 19 $BLUE $true
$null = Bul $s 348 170 238 120 @("À l'envoi par email ou téléphone","Date d'envoi dans les cases","Bloc « Attestation émise »") 14 $INK 5
$null = Card $s 620 126 292 170
$null = Txt $s 638 138 256 24 'Signée' 19 $GOOD $true
$null = Bul $s 638 170 260 120 @("Après « Marquer comme signée »","Date de signature choisie","Signature tracée ou nom","Mention écrite éventuelle") 14 $INK 4
$null = Card $s 330 312 582 180
$null = Txt $s 348 324 546 22 'Contenu' 17 $NAVY $true
$null = Bul $s 348 354 546 100 @("Collaborateur, texte d'attestation, tableau du matériel et total","Dates d'emprunt et de restitution, envoi, signature, mention","Cases de signature : date, signature ou nom, et « DSI » côté IT ; logo eXalt en en-tête") 14 $INK 5
$null = Txt $s 348 460 546 20 'Fichier : attestation_restitution_prenom-nom_ATT-3-5_signee.pdf' 12 $SOFT
# ================= 17. Archives =================
$s = NewSlide 'Onglet 6' 'Archives PDF : retrouver une preuve' 'Le PDF signé de chaque attestation, à tout moment' "Le PDF archivé est stocké dans la base (collection documents), de 6 à 15 Ko. Télécharger restitue exactement le fichier produit à la signature. Dans cette capture, trois attestations signées attendent encore leur archivage."
$null = ShotAuto $s 'archives.png' 48 126 570 1440 @(@(1,306,338),@(2,306,524),@(3,577,524),@(4,382,693),@(5,1105,654),@(6,1231,654))
$null = Card $s 668 126 244 366
$null = Txt $s 686 138 210 22 'Repères' 17 $NAVY $true
Legend $s 686 168 210 @('Indicateurs : archivés, à archiver, espace, dernière signature','Filtres : tous, archivés, à archiver','Période (date de signature)','Cycle de vie de l''attestation','Statut : Archivé ou Non archivé','Télécharger / Archiver le PDF') 12 40

# ================= 18. Perdu =================
$s = NewSlide 'Onglet 4' 'Matériel perdu : la synthèse' 'Perdu, volé ou abîmé, par type d''équipement' "Liste déroulante pour choisir le type ; les deux premiers indicateurs suivent le filtre. La date de perte n'est volontairement pas affichée."
$null = ShotAuto $s 'lost.png' 48 126 570 1440 @(@(1,306,322),@(2,300,520),@(3,700,520),@(4,385,667),@(5,1084,640),@(6,1260,577))
$null = Card $s 668 126 244 366
$null = Txt $s 686 138 210 22 'Repères' 17 $NAVY $true
Legend $s 686 168 210 @('Indicateurs : unités, déclarations, type le plus touché, ce mois-ci','Type de matériel (liste déroulante)','Période','Une pastille par déclaration','Jauge : part des pertes','Exporter CSV') 12 40
$null = Txt $s 686 424 210 62 "Au-delà de six pastilles : « + N autres »." 12 $SOFT

# ================= 19. Activités =================
$s = NewSlide 'Onglet 7' 'Activités récentes : le journal' 'Tout ce qui a été enregistré, du plus récent au plus ancien' "Les quatre indicateurs servent aussi de filtres ; un second clic retire le filtre. 50 événements affichés à la fois. La mention de signature figure sous l'événement « attestation signée »."
$null = ShotAuto $s 'activity.png' 48 126 570 1440 @(@(1,306,315),@(2,306,501),@(3,812,501),@(4,306,547),@(5,346,596))
$null = Card $s 668 126 244 366
$null = Txt $s 686 138 210 22 'Repères' 17 $NAVY $true
Legend $s 686 168 210 @('Indicateurs et filtres : prêts, envois, signatures, pertes','Filtres par type d''événement','Période','Regroupement par jour','Événement : personne, matériel, heure, auteur') 12 44

# ================= 20. Période =================
$s = NewSlide 'Transversal' 'Filtrer par période' 'Sur chaque liste, entre deux dates, deux mois ou deux années' "Le filtre ne modifie aucune donnée : il restreint l'affichage, les indicateurs et les exports de l'onglet. Réinitialiser rétablit toutes les dates."
$null = ShotAuto $s 'period.png' 48 126 864 1087 @()
$null = Card $s 48 296 420 196
$null = Txt $s 68 308 380 22 'Quatre modes' 17 $NAVY $true
$null = Bul $s 68 338 384 150 @("Toutes les dates (par défaut)","Entre deux dates : du … au …","De mois à mois","D'année à année") 14 $INK 5
$null = Card $s 492 296 420 196
$null = Txt $s 512 308 380 22 'Date utilisée selon l''onglet' 17 $NAVY $true
$null = Bul $s 512 338 384 150 @("Emprunter : date de remise","Restituer, Matériel perdu : retour ou déclaration","Attestations : date d'envoi","Archives PDF : date de signature","Activités : date de l'événement") 13 $INK 4

# ================= 21. QR =================
$s = NewSlide 'Transversal' 'Codes QR et scanner' 'Identifier un équipement en quelques secondes' "Les étiquettes se génèrent depuis la fiche d'une personne ; le scanner est dans le menu latéral et sur l'Accueil. La caméra n'est utilisée que pendant le scan et n'est jamais enregistrée. Bibliothèques jsQR et qrcode-generator chargées à la demande."
$null = ShotAuto $s 'scan.png' 48 126 270 500 @()
$null = Card $s 432 126 480 170
$null = Txt $s 450 138 440 22 'Étiquettes QR (PDF)' 18 $NAVY $true
$null = Bul $s 450 168 444 126 @("Bouton dans la fiche d'une personne","24 étiquettes par page A4 (3 × 8)","Chaque étiquette : QR, libellé, code, personne, date") 14 $INK 5
$null = Card $s 432 310 480 182
$null = Txt $s 450 322 440 22 'Scanner un QR' 18 $NAVY $true
$null = Bul $s 450 352 444 136 @("Caméra, photo du QR ou saisie du code (ex. AP-1)","Résultat : équipement, personne et état (disponible, emprunté, en attente, perdu)","Prêter maintenant, Restituer ce matériel ou ouvrir la fiche (scan rapide)") 14 $INK 5

# ================= Scan rapide et panneau latéral =================
$s = NewSlide 'Transversal' 'Scan rapide et panneau latéral' 'Prêter en un geste, et garder la liste sous les yeux sur grand écran' "Scan rapide : « Prêter maintenant » enregistre l'emprunt du jour (Annuler possible) puis la caméra repart ; « Restituer ce matériel » ouvre Restituer avec l'équipement coché. Panneau latéral : à partir de 1 480 px de large, la fiche s'affiche à droite ; confirmation avant d'abandonner une sélection non enregistrée."
$null = ShotAuto $s 'scanquick.png' 48 126 250 500 @()
$null = Card $s 48 424 250 68
$null = Txt $s 60 430 228 58 "Prêter maintenant : emprunt du jour, Annuler 10 s. Restituer ce matériel : fiche ouverte, équipement coché." 12 $INK
$null = ShotAuto $s 'sidepanel.png' 322 126 590 1700 @()
$null = Card $s 322 462 590 30
$null = Txt $s 334 468 568 20 "≥ 1 480 px : fiche à droite, liste cliquable ; en dessous, la fiche reste une fenêtre." 12 $INK

# ================= 16. Fiche et commentaires =================
$s = NewSlide 'Fiche personne' 'Fiche personne et commentaires' 'Le contexte de chaque personne, au même endroit' "Les commentaires sont un fil horodaté par personne ; chacun peut modifier ses propres commentaires."
$null = Card $s 48 126 420 366
$null = Circle $s 68 144 40 $BLUE '1' 17
$null = Txt $s 120 152 330 24 'La fiche' 20 $NAVY $true
$null = Bul $s 68 202 380 280 @("S'ouvre dans une fenêtre (ou un panneau à droite sur grand écran) ; croix, Échap ou clic à côté ; flèches précédent / suivant","Clic sur l'email ou le téléphone : copie","Compteurs : restitués, à signer, perdus","Historiques (emprunts, attestations, pertes) avec l'auteur de chaque action","Matériel emprunté et disponible") 16 $INK 10
$null = Card $s 492 126 420 366
$null = Circle $s 512 144 40 $PURPLE '2' 17
$null = Txt $s 564 152 330 24 'Les commentaires' 20 $NAVY $true
$null = Bul $s 512 202 380 280 @("Fil horodaté propre à chaque personne","Auteur, date et heure sur chaque message","Chacun modifie ses propres commentaires","Un brouillon en cours n'est jamais effacé","Visibles de tous les utilisateurs : pas d'information sensible") 16 $INK 10

# ================= 23. Envoi des messages =================
$s = NewSlide 'Envoi' 'Email et SMS : comment partent les messages' 'L''application prépare, l''utilisateur envoie' "Avec le connecteur Microsoft 365 autorisé : aperçu du PDF puis brouillon Outlook avec pièce jointe. Sinon : repli. La création du brouillon a été testée avec un connecteur simulé ; le premier essai réel est à vérifier (CR-12)."
$em = @(@('Connecteur Microsoft 365 autorisé',$GOOD,$TG,"Brouillon dans Outlook (dossier Brouillons), attestation jointe. La fenêtre « Message prêt à envoyer » propose un fichier .eml à ouvrir dans Outlook. Rien n'est envoyé sans clic."),
  @('Connecteur absent, expiré ou refusé',$WARN,$TW,"Même fenêtre : fichier .eml (message adressé, PDF joint) et raison de l'échec affichée."),
  @('Page ouverte hors de Claude',$BLUE,$TI,"Pas de connecteur : la messagerie par défaut s'ouvre sur le message préparé, PDF téléchargé."))
for($i=0;$i -lt 3;$i++){
  $x = 48 + $i*297
  $null = Card $s $x 128 270 220
  $null = Txt $s ($x+20) 148 232 50 $em[$i][0] 17 $em[$i][1] $true
  $null = Txt $s ($x+20) 204 232 140 $em[$i][3] 14 $INK
}
$sm2 = Card $s 48 366 864 56
$null = Txt $s 68 378 830 36 "SMS : application de messages, numéro de test 06 95 17 58 74.  Email : destinataire au choix, adresse enregistrée ou saisie à la main (personne sans compte professionnel)." 14 $INK
$wn = Box $s 48 436 864 54 $TW 5 $WARN
SetText $wn "Outlook sur le web ne s'ouvre pas depuis la page (bloqué par le navigateur) : brouillon ou fichier .eml. Premier essai réel à contrôler." 14 $INK $false 1 3
$wn.TextFrame.MarginLeft = 18

# ================= 24. Apparence =================
$s = NewSlide 'Confort' 'Apparence : thèmes et couleurs' 'Douze thèmes, sept couleurs d''accent, au choix de chacun' "Le choix est mémorisé dans le navigateur de chaque personne (pas partagé, pas synchronisé entre appareils). Le rouge et le vert ne sont pas proposés comme accents car ils signalent perdu et signé."
$null = ShotAuto $s 'themes.png' 48 126 180 470 @()
$null = Card $s 330 126 270 366
$null = Txt $s 348 138 234 22 'Thèmes' 18 $NAVY $true
$null = Bul $s 348 168 238 320 @("Système : suit l'appareil","Clair et Sombre","Minuit, Ardoise, Noir (sombres)","Océan et Forêt (sombres teintés)","Crème (clair chaud)","Contraste : noir sur blanc, lecture maximale","eXalt et eXalt nuit : charte de l'entreprise (Poppins, bleu eXalt, logo)") 13 $INK 5
$null = Card $s 622 126 290 170
$null = Txt $s 640 138 254 22 'Couleurs d''accent' 18 $NAVY $true
$null = Bul $s 640 168 256 126 @("Orange (défaut), Bleu, Violet","Turquoise, Rose, Graphite, eXalt (bleu de la charte)","Boutons, filtres actifs, liens et menu") 14 $INK 5
$null = Card $s 622 310 290 182 $TI
$null = Txt $s 640 322 254 22 'Où le choisir' 18 $BLUE $true
$null = Txt $s 640 352 254 134 "Bouton « Thème » en bas du menu latéral : un panneau propose les thèmes et les couleurs. 84 combinaisons possibles ; un thème eXalt sélectionne l'accent eXalt." 14 $INK

# ================= Recherche rapide =================
$s = NewSlide 'Transversal' 'Recherche rapide et raccourcis' 'Trouver une personne, un équipement ou une page en deux touches' "Le bouton Rechercher figure en haut de chaque page ; Ctrl K ouvre la même fenêtre. Un code (AP-1) ou un numéro de série affiche l'équipement et ouvre la fiche de la personne. Entrée ouvre l'action la plus utile sur une personne, Maj + Entrée l'autre."
$null = ShotAuto $s 'cmdk.png' 48 126 300 700 @()
$null = ShotAuto $s 'cmdksearch.png' 364 126 300 700 @()
$null = Card $s 364 304 300 188
$null = Txt $s 378 312 272 22 'Une ligne par personne' 15 $NAVY $true
$null = Bul $s 378 336 274 154 @("Entrée : l'action la plus utile (Restituer s'il y a du matériel à rendre, sinon Emprunter)","Maj + Entrée : l'autre action","Repères : En retard, 8 à rendre, Signature en attente") 12 $INK 4
$null = ShotAuto $s 'keys.png' 680 126 232 700 @()
$null = Txt $s 680 336 232 156 "Liste de départ : À traiter (chiffres), Activité récente, Aller à, Actions. Touche ? : aide clavier. Alt + flèches : personne précédente ou suivante." 12 $SOFT
$null = Card $s 48 404 300 88
$null = Txt $s 60 410 276 22 'À traiter, en un clic' 15 $NAVY $true
$null = Txt $s 60 434 276 56 "Retours en retard, attestations à relancer ou en attente, PDF à archiver, matériel à restituer : la page s'ouvre déjà filtrée." 12 $INK
# ================= Rapports PDF et Excel =================
$s = NewSlide 'Transversal' 'Rapports PDF et Excel' 'Un état à partager, en un clic, propre à la page où l''on se trouve' "Les rapports sont créés dans le navigateur (jsPDF et autotable pour le PDF, ExcelJS pour Excel, chargées à la demande). Ils sont réservés aux rôles Administrateur full, Super utilisateur et Lecteur."
$null = ShotAuto $s 'reportmodal.png' 48 126 190 500 @()
$null = ShotAuto $s 'reportpdf.png' 254 126 232 793 @()
$null = Card $s 506 126 406 170
$null = Txt $s 524 138 370 22 'Quatre clics' 18 $NAVY $true
$null = Bul $s 524 168 372 124 @("Bouton Rapport, en haut à droite de chaque page","Choisir la période d'activité et le format : PDF ou Excel","Générer le rapport, puis confirmer le téléchargement, ou « Envoyer par email »") 13 $INK 4
$null = Card $s 506 310 406 182 $TI
$null = Txt $s 524 322 370 22 'Contenu' 18 $BLUE $true
$null = Bul $s 524 352 372 138 @("PDF : indicateurs, parc par type, attestations à traiter, signatures, pertes, emprunts, situation par personne ; logo eXalt","Excel : 8 feuilles, filtres et première ligne figée") 13 $INK 4

# ================= Un rapport par page =================
$s = NewSlide 'Transversal' 'Un rapport propre à chaque page' 'Le contenu, le titre et le nom du fichier suivent la page choisie' "Le bouton Rapport figure en haut à droite d'Accueil, Emprunter, Restituer, Attestations, Archives PDF et Matériel perdu (masqué pour un Agent IT). La période s'applique à l'activité de la page ; les situations (parc, attestations en attente) sont celles d'aujourd'hui. Fichiers : rapport_gpiexalt_emprunts, restitutions, attestations, archives, perdu ; l'Accueil produit le rapport complet."
$null = ShotAuto $s 'reportpage.png' 48 126 180 500 @()
$null = ShotAuto $s 'reportreturn.png' 244 126 220 793 @()
$null = MakeTable $s 488 126 424 @(104,320) @(@('Page','Contenu du rapport'),
  @('Accueil','Rapport complet : parc par type, attestations, signatures, pertes, emprunts, personnes'),
  @('Emprunter','Emprunts de la période avec retour prévu et état ; matériel à remettre'),
  @('Restituer','Retours en retard, retours prévus, parc par type, situation par personne'),
  @('Attestations','À traiter, envoyées et signées de la période'),
  @('Archives PDF','PDF à archiver ; signatures avec l''état de l''archive'),
  @('Matériel perdu','Pertes par type et déclarations de la période')) 11 50

# ================= Envoyer le rapport par email =================
$s = NewSlide 'Transversal' 'Envoyer le rapport par email' 'Destinataire saisi à la main avec suggestions : rien ne part sans clic' "Le message part de la boîte Outlook de l'utilisateur par le connecteur Microsoft 365, après un clic sur Envoyer maintenant. Jusqu'à 10 adresses. Au-delà de 2,9 Mo (limite de Microsoft Graph), l'envoi direct est refusé : télécharger le rapport et le joindre à un message. Sans connecteur : fichier .eml."
$null = ShotAuto $s 'reportmail.png' 48 126 230 620 @()
$null = Card $s 300 126 612 170
$null = Txt $s 318 138 576 22 'Dans la fenêtre' 18 $NAVY $true
$null = Bul $s 318 168 576 124 @("Destinataire(s) : saisie à la main, suggestions dès 2 lettres (personnes de GPIEXALT, puis annuaire Microsoft 365)","Objet et message préremplis (titre de la page, date, format, période, expéditeur), modifiables","Pièce jointe : rapport_gpiexalt_AAAA-MM-JJ.pdf ou .xlsx") 14 $INK 5
$null = Card $s 300 310 612 182 $TW
$null = Txt $s 318 322 576 22 'Quatre boutons, rien sans clic' 18 $WARN $true
$null = Bul $s 318 352 576 136 @("Envoyer maintenant : part de votre boîte Outlook, copie dans les éléments envoyés","Créer un brouillon : dossier Brouillons d'Outlook","Télécharger le message : fichier .eml à ouvrir dans Outlook","Annuler : retour à la fenêtre du rapport") 14 $INK 5
# ================= Rôles et accès =================
$s = NewSlide 'Rôles et accès' 'Cinq rôles, du collaborateur à l''administrateur' 'Le partage de la page fixe la limite ; le rôle précise ce que chacun voit et fait' "Seul « Peut consulter » interdit techniquement toute écriture ; les rôles Agent IT et Lecteur restreignent l'affichage. Les droits réels passeront par Entra ID."
$null = MakeTable $s 48 126 620 @(130,290,110,90) @(@('Rôle','Ce que la personne voit et fait','Partage','Garantie'),
  @('Administrateur full','Tout, y compris Rôles et accès','Modifier','Plateforme'),
  @('Super utilisateur','Toutes les opérations, rapports et exports (rôle par défaut)','Interagir','Affichage'),
  @('Agent IT','Prêts, restitutions, attestations, signatures ; ni pertes ni rapports','Interagir','Affichage'),
  @('Lecteur','Tous les écrans et les rapports, sans saisie','Interagir','Affichage'),
  @('Collaborateur','Uniquement Mon matériel, en consultation','Consulter','Plateforme')) 12 42
$tgb = Box $s 48 396 620 96 $TI 5 $BLUE
SetText $tgb "Recherche par prénom et nom (début de mots accepté, nom complet affiché). Tous les utilisateurs : rattache chaque fiche au compte correspondant, sans toucher aux rôles attribués à la main ; un bilan liste les fiches sans compte trouvé." 13 $INK $false 1 3
$tgb.TextFrame.MarginLeft = 18
$null = ShotAuto $s 'roles.png' 692 126 220 1178 @()

# ================= Mon matériel =================
$s = NewSlide 'Rôles et accès' 'Mon matériel : la vue du collaborateur' 'Chaque collaborateur consulte son propre matériel, sans pouvoir modifier' "Un compte non rattaché voit « Compte non rattaché » et son identifiant, à communiquer à l'administrateur."
$null = ShotAuto $s 'mine.png' 48 126 330 1178 @()
$null = Card $s 410 126 502 190
$null = Txt $s 428 138 466 22 'Ce que voit le collaborateur' 18 $NAVY $true
$null = Bul $s 428 168 466 144 @("Quatre indicateurs : chez moi, en attente de signature, restitués et signés, perdu volé ou abîmé","Son matériel avec l'état de chaque équipement","Ses attestations, ses prêts et ses pertes déclarées","Aucune saisie ni signature") 13 $INK 4
$null = Card $s 410 330 502 162 $TW
$null = Txt $s 428 342 466 22 'Protection des données' 18 $WARN $true
$null = Bul $s 428 372 466 116 @("Résumé personnel tenu à jour à chaque action sur sa fiche (mine/compte)","La base refuse l'accès aux fiches des autres collaborateurs","Rattachement automatique par le nom, ou par l'identifiant du compte") 13 $INK 4

# ================= Annuaire Microsoft 365 =================
$s = NewSlide 'Rôles et accès' 'Annuaire Microsoft 365 : importer les utilisateurs' 'Les vrais collaborateurs dans Emprunter et Restituer, sans rien écrire dans Entra ID' "Lecture seule via le connecteur Microsoft 365 (list-users, list-groups, list-group-members). Le parc réel n'est pas importé : le connecteur ne donne pas accès à Intune. Les données importées sont personnelles : validation du DPO avant un import complet."
$null = ShotAuto $s 'dir.png' 48 126 540 1114 @()
$null = Card $s 48 366 540 126
$null = Txt $s 66 376 504 22 'Trois façons d''importer' 16 $NAVY $true
$null = Bul $s 66 402 504 88 @("Tout l'annuaire : comptes actifs, membres, avec email","Quelques personnes : recherche par nom, pour un essai","Un groupe Entra ID : moins de 999 membres") 13 $INK 3
$null = Card $s 612 126 300 366
$null = Txt $s 630 138 264 22 'Après l''import' 18 $NAVY $true
$null = Bul $s 630 168 266 320 @("Copiés : nom, fonction, email, mobile","Kit standard de 9 équipements par personne, sans emprunt","Synchroniser à nouveau : arrivées ajoutées, emprunts conservés","Un compte retiré reste listé pour restituer le matériel","« Tous les utilisateurs » est trop grand (plus de 999)") 13 $INK 8

# ================= Confort d'utilisation =================
$s = NewSlide 'Confort' 'Confort d''utilisation' 'Des détails qui font gagner du temps au quotidien' "Barre de progression et compteurs animés au chargement, impression en noir sur blanc, focus géré dans les fenêtres, animations réduites si le système le demande. Thème, couleur, densité et visite guidée sont mémorisés par navigateur."
$cf = @(@('Recherche collante',"La barre de recherche reste en haut pendant le défilement ; filtres actifs en pastille avec Réinitialiser.",$BLUE),
  @('Fiche et recherche',"Champ de recherche pour changer de personne, flèches précédent et suivant, Alt + flèches.",$PURPLE),
  @('Dates rapides',"Aujourd'hui et Hier sous les champs de date ; confirmation si du matériel coché n'est pas enregistré.",$TEAL),
  @('Tris',"Attestations et archives : en attente d'abord, plus récentes, plus anciennes, nom A à Z.",$WARN),
  @('Densité et thèmes',"Confortable ou compacte ; douze thèmes (dont eXalt) et sept couleurs, avec un fondu doux au changement.",$GOOD),
  @('Visite guidée',"Six étapes à la première ouverture, relançable depuis l'aide (?) ou la recherche rapide.",$ACC))
for($i=0;$i -lt 6;$i++){
  $x = 48 + ($i % 3)*297; $y = 128 + [math]::Floor($i / 3)*190
  $null = Card $s $x $y 270 172
  $null = Circle $s ($x+18) ($y+18) 16 $cf[$i][2] '' 8
  $null = Txt $s ($x+44) ($y+14) 214 26 $cf[$i][0] 16 $NAVY $true
  $null = Txt $s ($x+18) ($y+56) 236 108 $cf[$i][1] 14 $INK
}

# ================= 18. Téléphone et tablette =================
$s = NewSlide 'Compatibilité' 'Téléphone et tablette' 'La même application, adaptée à chaque écran' "Barre d'onglets en bas sur téléphone et tablette en portrait ; menu latéral sur PC et tablette en paysage."
$ph = Box $s 90 120 214 384 '14183A' 5 $null 0.12; Shadow $ph 20 6 0.7
$null = $s.Shapes.AddPicture((Join-Path $imgDoc 'mobile.png'), 0, -1, 109, 132, 176, 360)
$null = Card $s 340 126 572 172
$null = Txt $s 360 138 520 22 'Téléphone' 18 $NAVY $true
$null = Bul $s 360 168 530 126 @("Barre d'onglets fixée en bas, à portée du pouce, avec les compteurs","Boutons de 42 px minimum ; champs à 16 px (pas de zoom automatique sur iPhone)","Les boutons email et SMS ouvrent les applications de l'appareil") 14 $INK 5
$null = Card $s 340 314 572 178
$null = Txt $s 360 326 520 22 'Tablette et PC' 18 $NAVY $true
$null = Bul $s 360 356 530 130 @("Tablette en paysage et PC : menu latéral, repliable en icônes","Tablette en portrait : barre d'onglets en bas","PC : raccourcis 1 à 7, Ctrl K pour rechercher") 14 $INK 5

# ================= 19. Divider 03 =================
$null = Divider '03' 'Gouvernance' 'Rôles, responsabilités et règles de gestion'

# ================= 20. RACI =================
$s = NewSlide 'Gouvernance' 'Rôles et responsabilités' 'Matrice proposée, à valider par la DSI' "Cinq rôles dans l'application (Administrateur full, Super utilisateur, Agent IT, Lecteur, Collaborateur) ; le partage de la page reste la limite appliquée par la plateforme."
$null = MakeTable $s 48 126 864 @(414,112,112,114,112) @(
  @('Activité','Agent IT','DSI','Admin. appli','Collaborateur'),
  @('Enregistrer un prêt','R','I','–','C'),@('Enregistrer une restitution et envoyer l''attestation','R','I','–','I'),@('Relancer une attestation non signée','R','I','–','I'),
  @('Marquer comme signée et archiver','R','A','–','C'),@('Déclarer un matériel perdu, volé ou abîmé','R','A','–','C'),@('Fournir un PDF signé (audit, litige)','R','A','–','–'),
  @('Publier une nouvelle version de l''application','C','A','R','–'),@('Gérer les accès (partage, connecteurs)','–','A','R','–'),@('Tenir la documentation à jour','C','A','R','–')) 13 31 2
$null = Txt $s 48 448 864 40 "R = réalise  ·  A = approuve  ·  C = consulté  ·  I = informé.   Cinq rôles dans l'application ; le partage de la page reste la limite réelle." 13 $SOFT

# ================= 21. Règles de gestion =================
$s = NewSlide 'Gouvernance' 'Règles de gestion appliquées' 'Trente règles dans le dossier PDF ; les six essentielles ici' "RG-01 à RG-30 : chapitre 5 du dossier PDF."
$rg = @(@('RG-01 · Disponibilité',"Un équipement ne peut être emprunté que s'il est disponible : jamais emprunté, ou redevenu libre.",$BLUE),
  @('RG-03 · Signature',"Une attestation ne passe à « Signée » que par l'action « Marquer comme signée ».",$GOOD),
  @('RG-04 · Restitution',"Le matériel d'une attestation signée est restitué et redevient disponible.",$GOOD),
  @('RG-17 et 18 · Rôles',"Cinq rôles ; un collaborateur ne voit que sa propre fiche, en consultation seule.",$PURPLE),
  @('RG-06, 25 et 26 · Retards',"À relancer dès 3 jours ; retour prévu dépassé : alerte (fenêtre, carte, pastilles, filtre).",$WARN),
  @('RG-08 et 13 · Annulation',"Perte irréversible ; un emprunt ou une signature s'annule par le bouton Annuler (10 secondes).",$BAD))
for($i=0;$i -lt 6;$i++){
  $x = 48 + ($i % 3)*297; $y = 128 + [math]::Floor($i / 3)*190
  $null = Card $s $x $y 270 172
  $null = Circle $s ($x+18) ($y+18) 16 $rg[$i][2] '' 8
  $null = Txt $s ($x+44) ($y+14) 214 26 $rg[$i][0] 16 $NAVY $true
  $null = Txt $s ($x+18) ($y+56) 236 108 $rg[$i][1] 14 $INK
}

# ================= 22. Divider 04 =================
$null = Divider '04' 'Architecture et sécurité' 'Composants, accès, traçabilité et données personnelles'

# ================= 23. Architecture =================
$s = NewSlide 'Architecture et sécurité' 'Architecture et dépendances' 'Une page web, une base, un connecteur' "jsPDF est chargé depuis cdnjs.cloudflare.com : dépendance externe à faire valider par le RSSI. Le connecteur Microsoft 365 est autorisé par chaque utilisateur."
$null = Card $s 48 170 200 170
$null = Txt $s 66 184 166 24 'Navigateur' 18 $NAVY $true
$null = Txt $s 66 214 168 120 "PC, tablette, téléphone`rPage unique HTML / JavaScript`rPDF (jsPDF), rapports (ExcelJS)" 13 $SOFT
$db = Box $s 340 126 270 112 $NAVY; Shadow $db
$null = Txt $s 358 138 234 24 "Base de l'application" 17 'FFFFFF' $true
$null = Txt $s 358 166 236 68 "people : prêts, attestations, pertes, commentaires`rdocuments : PDF signésdirectory, roles, links, mine" 13 'CADCFC'
$cn = Box $s 340 252 270 84 $TI 5 $BLUE; Shadow $cn
SetText $cn "Connecteur Microsoft 365`rBrouillon Outlook + pièce jointe" 15 $INK $true 2 3
$cdn = Box $s 340 350 270 84 $TW 5 $WARN; Shadow $cdn
SetText $cdn "cdnjs · jsDelivr · Google Fonts`rjsPDF, pdf.js, QR, ExcelJS, Poppins" 14 $INK $true 2 3
$ap = Card $s 700 170 212 170 $TG
$null = Txt $s 716 184 182 44 "Applications de l'appareil" 17 $NAVY $true
$null = Txt $s 716 236 182 90 "Messagerie (mailto)`rSMS (sms:)" 14 $SOFT
$null = Arrow $s 252 210 336 178
$null = Arrow $s 252 262 336 292
$null = Arrow $s 252 320 336 390
$l = Arrow $s 252 245 696 245; $l.Line.DashStyle = 4
$null = Txt $s 48 450 864 40 "Chaque modification crée une version numérotée à la même adresse ; la base conserve les données d'une version à l'autre." 14 $SOFT

# ================= 24. Sécurité =================
$s = NewSlide 'Architecture et sécurité' 'Accès, traçabilité et données personnelles' 'Les points à valider relèvent du DPO et du RSSI' "Base légale, durée de conservation et hébergement : à valider. Aucune purge automatique à ce jour."
$null = Card $s 48 126 420 366
$null = Txt $s 70 140 380 26 'Accès et traçabilité' 20 $NAVY $true
$null = Bul $s 70 176 378 310 @("Accès réservé aux personnes avec lesquelles l'application est partagée","Cinq rôles ; un collaborateur ne voit que sa fiche (Mon matériel)","Chaque événement garde son auteur et sa date","PDF signés conservés tels que produits","Connecteur Microsoft 365 : consentement de chaque utilisateur, jetons jamais visibles de la page") 15 $INK 9
$null = Card $s 492 126 420 366 $TW
$null = Txt $s 514 140 380 26 'Données personnelles' 20 $WARN $true
$null = Bul $s 514 176 378 310 @("Données : nom, fonction, email, mobile (annuaire Entra ID), adresse saisie à la main, matériel, dates, signatures déclarées, commentaires","Finalité : gestion du parc, preuve de restitution, suivi des pertes","Base légale : à renseigner par le DPO","Durée de conservation : à définir (aucune purge automatique)","Hébergement : plateforme Claude, localisation à confirmer","Aucune donnée sensible au sens du RGPD") 15 $INK 9

# ================= 25. Divider 05 =================
$null = Divider '05' 'Sauvegarde' 'Où sont les données, comment elles sont protégées'

# ================= 26. Où =================
$s = NewSlide 'Sauvegarde' 'Où sont enregistrées les données' 'Dans la base de l''application, jamais sur le poste' "people et documents ; roles, links et mine pour les accès. Une personne n'apparaît dans la base qu'à sa première action."
$null = Card $s 48 150 200 130
$null = Txt $s 66 164 166 24 'Application' 18 $NAVY $true
$null = Txt $s 66 194 166 80 "Navigateur : PC, tablette ou téléphone" 14 $SOFT
$null = Arrow $s 254 215 326 215 $ACC
$null = Txt $s 250 180 80 28 'à chaque action' 11 $SOFT $false 2
$dbx = Box $s 332 126 580 230 $NAVY; Shadow $dbx
$null = Txt $s 354 140 540 26 "Base de l'application (hébergée par Claude)" 18 'FFFFFF' $true
$d1 = Box $s 354 176 262 164 '2F2F9B'
$null = Txt $s 370 188 230 24 'people' 19 'B9B6FF' $true
$null = Txt $s 370 218 230 118 "Un document par personne touchée : prêts, attestations (envois, signatures, relances), pertes, commentaires." 14 'FFFFFF'
$d2 = Box $s 632 176 262 164 '2F2F9B'
$null = Txt $s 648 188 230 24 'documents' 19 'B9B6FF' $true
$null = Txt $s 648 218 230 118 "Un document par attestation signée : le PDF signé, son nom, la date et l'auteur. Environ 18 Ko chacun." 14 'FFFFFF'
$null = Card $s 48 378 864 112
$null = Bul $s 68 392 826 96 @("Rien n'est stocké sur le poste ni dans OneDrive : toutes les personnes ayant accès voient les mêmes données","Une personne n'apparaît dans la base qu'à sa première action ; l'ouverture de l'application n'écrit rien","Limite de la base : 256 Ko par document ; collections directory (annuaire), roles, links et mine") 14 $INK 5

# ================= 27. Comment =================
$s = NewSlide 'Sauvegarde' 'Comment c''est enregistré et protégé' 'Relire, appliquer, enregistrer : pas d''écrasement entre collègues' "Au-delà de 6 secondes, l'enregistrement est abandonné et la pastille passe en Mode local."
$sv = @(@('1','Relire','La dernière version de la personne est lue dans la base'),@('2','Appliquer','Votre modification est ajoutée par-dessus'),@('3','Enregistrer','Le résultat est sauvegardé ; la pastille affiche « Synchronisé »'))
for($i=0;$i -lt 3;$i++){
  $x = 48 + $i*297
  $null = Card $s $x 126 270 130
  $null = Circle $s ($x+18) 142 36 $BLUE $sv[$i][0] 16
  $null = Txt $s ($x+66) 148 190 26 $sv[$i][1] 20 $NAVY $true
  $null = Txt $s ($x+18) 190 236 62 $sv[$i][2] 14 $INK
  if($i -lt 2){ $null = Arrow $s ($x+274) 191 ($x+293) 191 }
}
$null = Card $s 48 276 420 214 $TW
$null = Txt $s 68 290 380 24 'Si l''enregistrement échoue' 18 $WARN $true
$null = Bul $s 68 324 382 160 @("Au-delà de 6 secondes, l'action est abandonnée pour ne pas bloquer l'écran","La pastille passe à « Mode local »","L'action n'est pas conservée : recharger la page et recommencer") 14 $INK 6
$null = Card $s 492 276 420 214 $TI
$null = Txt $s 512 290 380 24 'Copies de secours' 18 $BLUE $true
$null = Bul $s 512 324 382 160 @("Exports CSV : attestations, matériel perdu","PDF signés téléchargeables un par un","Pas de sauvegarde automatique hors de l'application : export complet à prévoir","Fichier gpiexalt.html ouvert en local : aucune sauvegarde") 14 $INK 6

# ================= 28. Divider 06 =================
$null = Divider '06' 'Qualité' 'Plan de recette et registre des risques'

# ================= 29. Recette =================
$s = NewSlide 'Qualité' 'Plan de recette : état au 12/10/2026' '52 cas : 40 réalisés ou simulés en local, 11 prioritaires en production' "CR-01 à CR-52 dans le dossier PDF, chapitre 10."
$kp = @(@('40','cas réalisés ou simulés',$GOOD),@('9','cas à réaliser en local',$WARN),@('11','cas prioritaires en production',$BAD))
for($i=0;$i -lt 3;$i++){
  $x = 48 + $i*297
  $null = Card $s $x 122 270 82
  $null = Txt $s ($x+22) 128 100 70 $kp[$i][0] 48 $kp[$i][2] $true
  $null = Txt $s ($x+112) 146 150 52 $kp[$i][1] 15 $INK
}
$null = Txt $s 48 212 864 22 'Cas prioritaires en production' 17 $NAVY $true
$null = MakeTable $s 48 238 864 @(70,125,399,270) @(@('Réf.','Fonction','Scénario','Résultat attendu'),
  @('CR-12','Outlook','« Par email » avec le connecteur Microsoft 365 autorisé','Brouillon, PDF joint, bon destinataire'),
  @('CR-14','Sauvegarde','Enregistrer une action puis recharger la page','Données conservées, « Synchronisé »'),
  @('CR-18','Accès','Ouvrir le lien avec un compte non autorisé','Accès refusé'),
  @('CR-28','Rôles','Attribuer le rôle Collaborateur, puis « Tous les utilisateurs »','Fiches rattachées, bilan des absents'),
  @('CR-29','Collaborateur','Ouvrir avec un compte « Peut consulter » rattaché','Seule la vue Mon matériel ; autres fiches refusées'),
  @('CR-33','Annuaire','Importer 3 personnes choisies, ouvrir Emprunter','Seules ces 3 personnes sont proposées'),
  @('CR-34','Annuaire','Importer tout l''annuaire (plus de 999 comptes)','Tous les comptes actifs membres avec email'),
  @('CR-37','Message signé','Valider la signature, puis Envoyer maintenant','Message reçu avec le PDF signé joint'),
  @('CR-47','Rapport par email','Saisir un destinataire, puis Envoyer maintenant','Message reçu avec le rapport joint'),
  @('CR-48','Destinataire','Restituer « Par email » avec une adresse saisie à la main','Brouillon adressé à cette adresse'),
  @('CR-51','Rôles','Chercher « prénom nom » dans Rôles et accès','Nom complet affiché, personne trouvée')) 10 20
$null = Txt $s 48 486 864 24 "Les 52 cas couvrent toutes les fonctions : détail au chapitre 10 du dossier PDF." 12 $SOFT

# ================= 30. Risques =================
$s = NewSlide 'Qualité' 'Principaux risques' 'Vingt-trois risques suivis ; le plus élevé : parc réel non connecté' "Registre complet : chapitre 11 du dossier PDF."
$null = MakeTable $s 48 122 864 @(70,330,94,370) @(@('Réf.','Risque','Niveau','Mesure / état'),
  @('R-06','Parc réel non connecté : kit standard de 9 équipements par personne','Élevé','Choisir la source du parc (Intune, SharePoint, Excel) avant 1 500 collaborateurs'),
  @('R-01','Données non conservées si la base est injoignable','Moyen','Pastille « Mode local », export CSV régulier'),
  @('R-02','Signature enregistrée par erreur','Faible','Confirmation de la date ; Annuler pendant 10 s ; sinon administrateur'),
  @('R-04','Rôles Agent IT et Lecteur appliqués à l''affichage seulement','Moyen','Seul « Peut consulter » interdit l''écriture ; rôles réels via Entra ID'),
  @('R-15','Annuaire importé : données personnelles lisibles par tous les utilisateurs de la page','Moyen','Validation du DPO ; essais sur sélection ou groupe ; lecture réservée aux admins à ajouter'),
  @('R-19','Alertes de retard visibles seulement à l''ouverture (pas de notification)','Moyen','Fenêtre, carte et pastilles rouges ; tâche planifiée à étudier'),
  @('R-09','Pas de durée de conservation ni de purge','Moyen','Durée à définir avec le DPO ; purge à développer'),
  @('R-10','Pas de sauvegarde externe automatique','Moyen','Export complet périodique à mettre en place')) 12 40
# ================= 31. Divider 07 =================
$null = Divider '07' 'Exploitation' 'Mises à jour, limites connues et actions attendues'

# ================= 32. Mises à jour =================
$s = NewSlide 'Exploitation' 'Mises à jour de l''application et de la documentation' 'Les données sont conservées ; la documentation se met à jour sur demande' "Chaque modification de l'application crée une nouvelle version à la même adresse. La documentation ne se met pas à jour toute seule."
$null = Card $s 48 126 420 262
$null = Txt $s 68 140 380 26 'L''application' 20 $NAVY $true
$null = Bul $s 68 176 382 206 @("Chaque modification crée une nouvelle version, à la même adresse","Les données sont conservées d'une version à l'autre","Historique des versions : retour en arrière possible","Dans le code : annuaire de démonstration (avant import), catalogue, textes, relance (3 jours), SMS de test") 14 $INK 6
$null = Card $s 492 126 420 262 $TI
$null = Txt $s 512 140 380 26 'La documentation' 20 $BLUE $true
$null = Bul $s 512 176 382 206 @("Elle ne se met pas à jour automatiquement","Après une évolution : « mets à jour la documentation »","Les sections concernées sont ajustées, la version du document est incrémentée","PDF et PowerPoint sont régénérés") 14 $INK 6
$jv = Box $s 48 408 864 80 $NAVY
SetText $jv "Journal : Doc 1.0 (05/10/2026) première édition  ·  Doc 2.0 / App 81 (05/10/2026) édition entreprise  ·  Doc 3.0 / App 106 (06/10/2026) accueil, période, QR, signature, thèmes  ·  Doc 4.0 / App 128 (07/10/2026) rôles, rapports, recherche  ·  Doc 5.0 / App 147 (08/10/2026) annuaire Entra ID, envoi signé  ·  Doc 6.0 / App 161 (09/10/2026) GPIEXALT, retards, rapport par email  ·  Doc 7.0 / App 167 (12/10/2026) rapport par page" 12 'FFFFFF' $false 1 3
$jv.TextFrame.MarginLeft = 20

# ================= 33. Limites =================
$s = NewSlide 'Exploitation' 'Limites connues et évolutions prévues' 'Ce qui n''est pas encore là, et ce qui est prévu' "Données de démonstration, rôles Agent IT et Lecteur appliqués à l'affichage, envoi à relire, SMS vers un numéro de test, pas de sauvegarde externe automatique. La signature enregistrée est déclarée par l'équipe IT, non qualifiée."
$li = @(@('Annuaire et parc',"Annuaire : importé depuis Entra ID. Parc : kit standard de 9 équipements, source réelle à choisir.",$BLUE),
  @('Envoi',"Attestation signée : envoi après confirmation, ou brouillon. Aucun envoi sans clic.",$PURPLE),
  @('SMS',"Numéro de test, via l'application de messages de l'appareil.",$WARN),
  @('Droits et rôles',"Agent IT et Lecteur : affichage seulement. Droits réels via Entra ID à prévoir.",$BAD),
  @('Départs et retards',"Date de départ par collaborateur : non réalisé. Alertes de retard visibles à l'ouverture seulement.",$TEAL),
  @('Sauvegarde externe',"Export complet régulier à prévoir avant la mise en production.",$GOOD))
for($i=0;$i -lt 6;$i++){
  $x = 48 + ($i % 3)*297; $y = 128 + [math]::Floor($i / 3)*190
  $null = Card $s $x $y 270 172
  $null = Circle $s ($x+18) ($y+18) 16 $li[$i][2] '' 8
  $null = Txt $s ($x+44) ($y+14) 214 26 $li[$i][0] 16 $NAVY $true
  $null = Txt $s ($x+18) ($y+56) 236 108 $li[$i][1] 14 $INK
}

# ================= 34. Décisions =================
$s = NewSlide 'Prochaines étapes' 'Décisions et actions attendues' 'Six actions pour passer à une mise en production encadrée' "Les responsables sont des propositions. Après décision : mise à jour de la documentation en version 7.1."
$null = MakeTable $s 48 126 864 @(46,470,200,148) @(@('N°','Action','Responsable proposé','Échéance'),
  @('1','Désigner le vérificateur et l''approbateur du dossier','DSI','À définir'),
  @('2','Renseigner les contacts de support (niveaux 1 à 3)','DSI','À définir'),
  @('3','Fixer la durée de conservation et la base légale','DPO','À définir'),
  @('4','Valider l''hébergement et le chargement des bibliothèques (jsPDF, ExcelJS…)','RSSI','À définir'),
  @('5','Réaliser l''essai réel du brouillon et des envois Outlook (CR-12, CR-37, CR-47, CR-48) et les cas prioritaires','Administrateur de l''application','À définir'),
  @('6','Choisir la source du parc (Intune, SharePoint, Excel), valider l''annuaire avec le DPO','DSI','À définir')) 14 46
$q = Box $s 48 458 864 38 $TI 5 $BLUE
SetText $q "Après décision, la documentation est mise à jour (version 7.1) : demande « mets à jour la documentation »." 14 $INK $false 1 3
$q.TextFrame.MarginLeft = 18


$pres.SaveAs($pptxOut)
$pres.Export($imgDir, 'PNG', 1600, 900)
$count = $pres.Slides.Count
$pres.Close(); $ppt.Quit()
"OK slides=$count"
