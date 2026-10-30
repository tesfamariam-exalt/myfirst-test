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
  $dy = 0; if(@('loan.png','returnlist.png','attest.png','lost.png','inv_price.png') -contains $file){ $dy = 23 }   # la ligne de date en haut des pages a été retirée : les repères montent d'autant
  foreach($m in $markers){
    $cx = $x + $m[1]*$k; $cy = $y + 22 + ($m[2] - $dy)*$k
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
  $null = Txt $s 48 512 700 14 "GPIEXALT  ·  DOC-GPIEXALT-001  ·  Version 7.8  ·  30/10/2026  ·  Classification : Interne" 9 $SOFT
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
  $null = Txt $s 96 500 700 14 "GPIEXALT  ·  DOC-GPIEXALT-001  ·  Version 7.8" 9 '7C86B8'
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
$null = Txt $s 96 430 620 18 "DOC-GPIEXALT-001  ·  Version 7.8  ·  30/10/2026  ·  Application v452" 13 'AEB9E3'
$null = Txt $s 96 452 620 18 "Propriétaire : Service informatique (DSI)  ·  Classification : Interne  ·  Statut : diffusé pour validation" 13 'AEB9E3'
$s.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text = "Présentation de GPIEXALT (Gestion du Parc Informatique eXalt, anciennement Prisme) : fonctions, procédures, règles de gestion, architecture, sécurité, sauvegarde, recette et risques. Version 6.0 : thème et logo eXalt, retour prévu et alertes de retard, actions groupées, scan rapide, rapport par email, destinataire au choix. Version 7.0 : rapport propre à chaque page, fiche d'emprunt réorganisée. Version 7.4 : onglet Inventaire (stock, achats, prix et dépenses), stock jamais négatif et matériel épuisé grisé dans Emprunter, calendrier et listes de l'application, rapport de l'Inventaire."

# ================= 2. Contrôle du document =================
$s = NewSlide 'Gouvernance' 'Contrôle du document' 'Référence, version, propriétaire et circuit de validation' "Les noms du vérificateur et de l'approbateur restent à renseigner."
$null = Card $s 48 126 420 366
$null = Txt $s 68 140 380 22 'Informations générales' 17 $NAVY $true
$null = MakeTable $s 62 170 392 @(120,272) @(
  @('Rubrique','Valeur'),@('Référence','DOC-GPIEXALT-001'),@('Version','7.8 — 30/10/2026'),@('Application','GPIEXALT, version 452'),
  @('Propriétaire','Service informatique (DSI)'),@('Classification','Interne'),@('Statut','Diffusé pour validation'),@('Diffusion','Équipe IT, DSI ; RSSI et DPO (sécurité, données)')) 12 36
$null = Card $s 492 126 420 234
$null = Txt $s 512 138 380 22 'Historique des révisions' 17 $NAVY $true
$null = MakeTable $s 506 168 392 @(56,82,254) @(@('Version','Date','Nature'),@('1.0-2.0','05/10/2026','Première édition, édition entreprise'),@('3.0','06/10/2026','App 106 : accueil, QR, signature'),@('4.0','07/10/2026','App 128 : rôles, rapports, recherche'),@('5.0','08/10/2026','App 147 : annuaire, envoi signé'),@('6.0','09/10/2026','App 161 : GPIEXALT, retards, rapport'),@('7.0-7.2','12/10/2026','App 167 à 192 : rapport par page, finition, catalogue'),@('7.3-7.4','13/10/2026','App 206 à 242 : Par email, Inventaire, prix, stock'),@('7.5','16/10/2026','App 299 : Nouvel emprunt, quantités, machines, achats'),@('7.6','27/10/2026','App 409 : Teams, recherche globale, anneaux, familles'),@('7.7','28/10/2026','App 417 : import de facture, À commander'),@('7.8','30/10/2026','App 452 : prolonger, relances, Guide, TVA')) 8 14
$null = Card $s 492 368 420 124
$null = Txt $s 512 374 380 22 'Circuit de validation' 17 $NAVY $true
$null = MakeTable $s 506 400 392 @(100,190,102) @(@('Étape','Nom et fonction','Date'),@('Rédaction','Service informatique','30/10/2026'),@('Vérification','À renseigner',''),@('Approbation','À renseigner (DSI)','')) 12 22

# ================= 3. Sommaire =================
$s = NewSlide 'Gouvernance' 'Sommaire' 'Sept parties, de la vue d''ensemble aux actions attendues' "Le dossier PDF détaille chaque partie (100 pages, procédures, plan de recette, registre des risques)."
$sm = @(
  @('01','Vue d''ensemble','Finalité, cycle de vie, les huit onglets'),
  @('02','Fonctions','Prêts, retours, inventaire, attestations, archives, QR, recherche, rapports, rôles, thèmes'),
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
$s = NewSlide 'Vue d''ensemble' 'Les huit onglets' 'Piloter, opérer, suivre' "Sur PC : menu latéral repliable ; sur téléphone et tablette en portrait : barre d'onglets en bas. Raccourcis 1 à 8."
$tabs = @(@('Accueil','Ce qui demande une action, parc, stock à commander',$TEAL),@('Emprunter','Enregistrer les équipements remis à une personne',$BLUE),@('Restituer','Attestation de retour, ou perte, vol, casse',$PURPLE),@('Matériel perdu','Synthèse du matériel perdu, volé ou abîmé',$BAD),@('Inventaire','Acheté, en stock, à commander, prix et dépenses','2D9CDB'),@('Attestations','Suivre les envois : en attente, à relancer, signées',$WARN),@('Archives PDF','Retrouver le PDF signé de chaque attestation',$GOOD),@('Activités récentes','Journal de tout ce qui a été enregistré','8A93C0'))
for($i=0;$i -lt 8;$i++){
  $col = $i % 4; $row = [math]::Floor($i / 4)
  $x = 48 + $col*219; $y = 120 + $row*130
  $null = Card $s $x $y 207 118
  $null = Circle $s ($x+14) ($y+12) 30 $tabs[$i][2] ([string]($i+1)) 14
  $null = Txt $s ($x+54) ($y+17) 148 22 $tabs[$i][0] 16 $NAVY $true
  $null = Txt $s ($x+14) ($y+54) 182 60 $tabs[$i][1] 13 $INK
}
$null = Card $s 486 392 426 98 $NAVY
$null = Txt $s 506 402 390 20 'AUSSI DANS L''APPLICATION' 11 'B9B6FF' $true
$null = Txt $s 506 424 390 64 "Recherche rapide · Rapport PDF ou Excel de chaque page · Alertes de retard · Rôles et accès · Annuaire Microsoft 365 · Filtres par période · Calendrier et listes · Thèmes eXalt" 14 'FFFFFF'
$null = Card $s 48 392 424 98
$null = Txt $s 68 400 400 22 'Aides à la navigation' 17 $NAVY $true
$null = Bul $s 68 426 396 62 @("Raccourcis : Ctrl K, 1 à 8 (onglets), / (liste), ? (aide), Échap","Douze thèmes (dont eXalt) et sept couleurs d'accent","Pastille de sauvegarde : « Synchronisé » ou « Mode local »") 12 $INK 2
# ================= 8. Divider 02 =================
$null = Divider '02' 'Fonctions' 'Chaque écran, chaque bouton, chaque donnée enregistrée'

# ================= 9. Accueil =================
$s = NewSlide 'Onglet 1' 'Accueil : le tableau de bord' 'Ce qui demande une action, en un coup d''œil' "Les indicateurs (dont « Retours en retard », visible seulement s'il y en a, et « À commander ») ouvrent l'onglet concerné avec le filtre déjà appliqué. Le bandeau « PDF à archiver » n'apparaît que s'il y en a. Le parc matériel affiche aussi le stock ; la synthèse du stock par type est la même que dans l'Inventaire. Les dernières activités ne sont plus sur l'Accueil : elles se consultent dans Activités récentes."
$null = ShotAuto $s 'home.png' 48 126 330 1440 @(@(1,763,60),@(2,280,196),@(3,1180,198),@(4,350,310),@(5,356,525),@(6,945,525),@(7,365,950))
$null = Card $s 400 126 512 366
$null = Txt $s 418 138 476 22 'Repères' 17 $NAVY $true
Legend $s 418 168 476 @('Accès rapides : recherche, prêt, retour, rapport','Indicateurs cliquables, dont Retours en retard','Carte « À commander » : types sans stock (Inventaire)','Bandeau « PDF à archiver » (si besoin)','Parc matériel : anneau cliquable, avec « En stock »','Stock par type de matériel : tous les types','Dépenses : répartition par type et par mois') 13 40
# ================= 10. Emprunter =================
$s = NewSlide 'Onglet 2' 'Emprunter : enregistrer un prêt' 'Rechercher la personne, cocher le matériel, choisir la date, enregistrer' "La fiche s'ouvre dans une fenêtre avec trois étapes. Un équipement peut être emprunté et restitué plusieurs fois. Annuler (10 secondes) supprime l'emprunt qui vient d'être saisi."
$null = ShotAuto $s 'loan.png' 48 126 552 1440 @(@(1,284,238),@(2,308,353),@(3,301,433),@(4,749,433),@(5,1080,572),@(6,1314,561),@(7,905,433),@(8,927,79))
$null = Card $s 636 126 276 366
$null = Txt $s 654 138 240 22 'Repères' 17 $NAVY $true
Legend $s 654 168 240 @('Quatre indicateurs','Recherche (accents ignorés)','Filtres : à emprunter, tout emprunté','Période (entre deux dates…)','Dernier mouvement et retour prévu','Historique (N) : matériel et dates','Nouvel emprunt : personne absente','Guide : mode d''emploi') 13 29
$null = Txt $s 654 408 240 18 'Dans la fiche' 14 $NAVY $true
$null = Txt $s 654 428 240 62 "3 étapes ; carte « Retour prévu » facultative. Prolonger un prêt, pastille « Déjà 1 ordinateur » (doublon). Matériel sans stock : grisé. Annuler pendant 10 s." 12 $SOFT

# ================= Nouvel emprunt et quantités =================
$s = NewSlide 'Onglet 2' 'Nouvel emprunt et emprunt par quantité' 'Un collaborateur absent de la liste, des quantités, un PC par exemplaire' "Le bouton + Nouvel emprunt ouvre directement la fiche : nom et prénom obligatoires, fonction, email et téléphone facultatifs. À l'enregistrement, le collaborateur est ajouté à la liste et l'emprunt enregistré ; un nom déjà présent n'est pas dupliqué. Tout le matériel se prête par quantité, dans la limite du stock (20 si aucun achat). Pour un ordinateur portable ou un Mac, un champ de recherche de PC par exemplaire (nom ou numéro de série) ; un PC déjà choisi ou prêté n'est plus proposé."
$null = ShotAuto $s 'loannew.png' 48 126 276 1100 @()
$null = ShotAuto $s 'loanqty.png' 336 126 276 1100 @()
$null = Card $s 632 126 280 366
$null = Txt $s 650 138 244 22 'À retenir' 17 $NAVY $true
$null = Bul $s 650 168 246 320 @("+ Nouvel emprunt : la fiche s'ouvre, le collaborateur se saisit en haut","Nom et prénom obligatoires ; email, fonction et téléphone facultatifs","Quantité par équipement : compteur − / + (stock de l'Inventaire)","PC remis : un champ de recherche par exemplaire","Un nom déjà présent n'est pas dupliqué") 13 $INK 6

# ================= Chargeur et Écran : familles =================
$s = NewSlide 'Onglet 2' 'Chargeur et Écran : un modèle par exemplaire' 'Une seule ligne dans la liste, le modèle ou la marque au choix' "Chargeur (USB-C, HP, Asus, Mac, Acer, Dell) et Écran (Dell, MiniFire, Samsung) sont des familles : une seule ligne dans Emprunter, Restituer, l'Inventaire et à l'achat. Quand la ligne est cochée, une liste Modèle (un par chargeur) ou Marque (une par écran) s'affiche pour chaque exemplaire, avec le stock restant ; le choix est obligatoire. Chaque modèle est ensuite enregistré, restitué et compté sous son propre nom. Dans l'Inventaire, un seul anneau additionne les modèles."
$null = ShotAuto $s 'chgform.png' 48 126 380 1500 @()
$null = ShotAuto $s 'loanhist.png' 440 126 172 1450 @()
$null = Card $s 632 126 280 366
$null = Txt $s 650 138 244 22 'À retenir' 17 $NAVY $true
$null = Bul $s 650 168 246 316 @("Une ligne « Chargeur » et une ligne « Écran »","Un modèle ou une marque par exemplaire, obligatoire","Stock restant affiché dans la liste","Enregistré sous le modèle choisi","Historique (N) : tout le matériel d'une personne, avec ses dates") 13 $INK 6

# ================= Prolonger un prêt et doublons =================
$s = NewSlide 'Onglets 2 et 3' 'Prolonger un prêt et éviter les doublons' 'Repousser une date sans recréer le prêt, être prévenu d''un doublon' "Prolonger : bouton sur la ligne du matériel (Restituer et fiche d'Emprunter) quand l'emprunt a une date de retour. La fenêtre propose + 1 semaine, + 2 semaines, + 1 mois, + 3 mois ou une date libre, avec un motif facultatif ; si l'échéance est dépassée, le calcul part d'aujourd'hui. L'ancienne et la nouvelle date, l'auteur et le motif sont consignés sur l'emprunt ; Annuler rétablit la date d'origine. Doublons : une pastille orange « Déjà 1 ordinateur » ou « Déjà 3 chargeurs » prévient quand la personne a déjà un ordinateur (PC ou Mac) ou un chargeur ; une confirmation à l'enregistrement laisse décider, sans jamais bloquer."
$null = ShotAuto $s 'prolong.png' 48 126 300 940 @()
$null = ShotAuto $s 'dupwarn.png' 362 126 270 750 @()
$null = Card $s 362 292 270 200
$null = Txt $s 378 302 240 22 'Doublons' 17 $NAVY $true
$null = Bul $s 378 330 242 158 @("Pastille orange sur la ligne du catalogue","Ordinateur (PC ou Mac) ou chargeur déjà détenu","Confirmation à l'enregistrement : jamais bloquante","Nom du PC en info-bulle") 13 $INK 5
$null = Card $s 652 126 260 366
$null = Txt $s 668 138 228 22 'Prolonger' 17 $NAVY $true
$null = Bul $s 668 168 230 320 @("Bouton sur chaque matériel qui a une date de retour","+ 1 semaine, + 2 semaines, + 1 mois, + 3 mois ou une date libre","Motif facultatif","Échéance dépassée : calcul à partir d'aujourd'hui","Ancienne et nouvelle date consignées sur l'emprunt","Annuler dans la bulle rétablit la date d'origine") 13 $INK 6

# ================= 11. Restituer : le suivi =================
$s = NewSlide 'Onglet 3' 'Restituer : suivre les retours' 'Indicateurs, filtres et dernier mouvement de chaque personne' "Chaque ligne affiche le dernier mouvement de la personne : le matériel, la date d'emprunt, puis le retour prévu (en rouge s'il est dépassé) ou la date de rendu. Le bouton Historique ouvre tout le matériel de la personne. Le sélecteur de tri a été retiré."
$null = ShotAuto $s 'returnlist.png' 48 126 552 1440 @(@(1,292,216),@(2,308,371),@(3,301,451),@(4,1080,590),@(5,1314,580),@(6,1134,451))
$null = Card $s 636 126 276 366
$null = Txt $s 654 138 240 22 'Repères' 17 $NAVY $true
Legend $s 654 168 240 @('Quatre indicateurs cliquables','Recherche rapide','Filtres : à traiter, en retard de retour, en attente, perdu…','Dernier mouvement : matériel, emprunté le, retour prévu ou rendu le','Historique (N) : tout le matériel de la personne','Période : sur la ligne des filtres (le tri a été retiré)') 13 31
$null = Txt $s 654 372 240 18 'Priorité d''une ligne' 14 $NAVY $true
$null = Txt $s 654 392 240 96 "Le bord gauche suit la priorité : perte (rouge), attente de signature (jaune), en cours (bleu), tout rendu (vert). « Retour en retard » en rouge sur la ligne." 12 $SOFT

# ================= 12. Restituer : actions =================
$s = NewSlide 'Onglet 3' 'Restituer : trois actions possibles' 'Attestation par email ou téléphone, ou déclaration de perte, vol, casse' "Rien n'est envoyé automatiquement : la personne connectée envoie elle-même le message préparé. Par email : un aperçu du PDF s'affiche d'abord. Les déclarations sont irréversibles (confirmation demandée)."
$bottom = ShotAuto $s 'return.png' 48 124 272 920 @(@(1,22,148),@(2,22,403),@(3,22,461),@(4,22,887),@(5,22,1031),@(6,22,1142),@(7,870,158))
$null = Card $s 432 124 480 368
$null = Txt $s 452 136 440 22 'Repères' 17 $NAVY $true
Legend $s 452 168 440 @('Compteurs : restitués, à signer, perdus','Étapes : matériel, date de retour, envoi','Matériel emprunté : cases et quantités','Date de retour, « Sélectionner tout »','Par email / Par téléphone : attestation en attente','Déclarer perdu / volé / abîmé (confirmation)','Recherche d''une autre personne') 13 29
$null = Txt $s 452 376 440 20 'Résultat' 15 $NAVY $true
$null = Bul $s 452 398 440 92 @("Par email : aperçu du PDF et destinataire, puis signature (Marquer comme signée), puis message prêt à envoyer","Par téléphone : SMS préparé, numéro de test 06 95 17 58 74","Déclaration : sort de chez la personne, visible dans Matériel perdu") 13 $INK 3

# ================= Relances Teams par paliers =================
$s = NewSlide 'Onglet 3' 'Rappels et relances Teams par paliers' 'Trois jours avant, le jour même, une semaine, plus de quinze jours' "L'application prépare chaque relance selon l'échéance : rappel aimable à J-3, message neutre le jour du retour, message plus ferme après une semaine, puis au-delà de 15 jours avec le responsable en copie. Un bandeau dans Restituer annonce les relances à envoyer ; le bouton Préparer les relances ouvre la file (Préparer, Ignorer). Le message s'ouvre dans Teams, déjà écrit : il suffit de relire et d'appuyer sur Entrée. Une relance ouverte dans Teams est consignée sur l'emprunt et n'est plus reproposée tant que la date de retour ne change pas. Rien n'est envoyé automatiquement."
$null = ShotAuto $s 'relbandeau.png' 48 126 440 1870 @()
$null = ShotAuto $s 'relfile.png' 48 326 200 1090 @()
$null = ShotAuto $s 'teams.png' 262 326 212 1800 @()
$null = Card $s 540 126 372 366
$null = Txt $s 558 138 336 22 'Quatre paliers' 17 $NAVY $true
$null = MakeTable $s 554 168 344 @(128,216) @(@('Palier','Quand'),@('Avant l''échéance','3 jours avant'),@('Échéance','Jour du retour, jusqu''à J+6'),@('Retard d''une semaine','De 7 à 15 jours'),@('Plus de 15 jours','Responsable en copie')) 12 26
$null = Bul $s 558 312 336 178 @("Bandeau « relances Teams à envoyer » dans Restituer","File : Préparer ou Ignorer, par personne","Message modifiable, jamais envoyé automatiquement","Relance consignée : « Relancé le … », plus reproposée") 13 $INK 6

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
$s = NewSlide 'Onglet 6' 'Attestations : suivi et cycle de vie' 'Toutes les preuves, avec leur statut et leurs actions' "Chaque ligne affiche un cycle de vie : Envoyée, Signée, PDF archivé. En attente : jours d'attente, rouge à partir de 3 jours (REMINDER_AFTER_DAYS). Le bouton Relancer est masqué par défaut (SHOW_REMIND_BUTTON)."
$null = ShotAuto $s 'attest.png' 48 126 570 1440 @(@(1,284,238),@(2,307,433),@(3,950,433),@(4,399,631),@(5,1030,593),@(6,1212,593),@(7,1364,593),@(8,1132,433),@(9,282,511))
$null = Card $s 668 126 244 366
$null = Txt $s 686 138 210 22 'Repères' 17 $NAVY $true
Legend $s 686 168 210 @('Indicateurs : envoyées, en attente, à relancer, délai moyen','Filtres : à relancer, signées…','Période (date d''envoi)','Cycle de vie : envoyée, signée, archivée','En attente · N j (rouge dès 3 jours)','Marquer comme signée','PDF de l''attestation','Exporter CSV','Sélection multiple (cases à cocher)') 12 33

# ================= Actions groupées =================
$s = NewSlide 'Onglet 6' 'Attestations : actions groupées' 'Cocher plusieurs attestations, puis les signer ou les exporter d''un coup' "Une case par ligne et « Sélectionner toutes les attestations en attente de cette liste » (suit la recherche, les filtres et la période). La barre d'actions apparaît en bas de la liste. La signature groupée exige une date (jamais dans le futur) et une mention écrite ; elle archive un PDF par attestation et n'envoie aucun message."
$null = ShotAuto $s 'bulk.png' 48 126 520 1440 @()
$null = ShotAuto $s 'bulkmodal.png' 592 126 320 500 @()
$null = Card $s 48 480 520 28
$null = Txt $s 58 485 500 20 "Marquer N comme signées  ·  Exporter la sélection (CSV)  ·  Désélectionner" 12 $INK

# ================= 14. Signature =================
$s = NewSlide 'Onglet 6' '« Marquer comme signée » pas à pas' 'Une date, une signature, une mention, un PDF archivé' "La date ne peut pas être dans le futur. Un tracé (doigt, stylet, souris) ou une mention écrite est obligatoire. La mention (140 caractères) permet de signer à la place de quelqu'un. Annuler pendant 10 secondes."
$st = @(@('1','Clic sur le bouton','Dans Attestations ou dans la fiche de la personne'),@('2','Date, et signature ou mention','Un tracé ou une mention est obligatoire : le bouton Valider reste grisé sinon'),@('3','PDF signé généré','Signature et mention dans le PDF ; « DSI » côté IT'),@('4','Archivage, message et Annuler','PDF rangé dans Archives ; envoi au collaborateur proposé ; Annuler 10 s'))
for($i=0;$i -lt 4;$i++){
  $y = 128 + $i*72
  $null = Circle $s 48 $y 40 $GOOD $st[$i][0] 18
  $null = Txt $s 102 ($y-1) 330 24 $st[$i][1] 18 $NAVY $true
  $null = Txt $s 102 ($y+24) 320 40 $st[$i][2] 13 $SOFT
}
$w1 = Box $s 48 424 400 66 $TW 5 $WARN
SetText $w1 "Si l'archivage échoue, la signature est quand même enregistrée : « Archiver le PDF » permet de recommencer." 13 $INK $false 1 3
$null = ShotAuto $s 'sign.png' 516 124 195 464 @(@(1,14,249),@(2,14,443),@(3,375,344),@(4,14,634),@(5,272,588),@(6,14,746),@(7,179,824),@(8,428,824))
$null = Card $s 772 124 140 366
$null = Txt $s 782 132 120 18 'Repères' 13 $NAVY $true
Legend $s 784 158 126 @('Date','Signature','Effacer','Mention','Pré-remplir','Message d''aide','Annuler','Valider') 11 30

# ================= Envoi de l'attestation signée =================
$s = NewSlide 'Onglet 6' 'Envoyer l''attestation signée par email' 'Après la signature : un message prêt à envoyer, avec le PDF signé joint' "Le message part de la boîte Outlook de l'utilisateur par le connecteur Microsoft 365, après un clic sur Envoyer maintenant. Le destinataire est l'adresse de l'annuaire : à contrôler dans la fenêtre. L'envoi réel reste à éprouver (CR-37)."
$null = ShotAuto $s 'sendmail.png' 48 124 250 580 @()
$null = Card $s 330 126 582 176
$null = Txt $s 350 138 540 22 'Quatre choix dans la fenêtre' 18 $NAVY $true
$null = Bul $s 350 168 546 130 @("Envoyer maintenant : le message part avec le PDF signé joint, copie dans les éléments envoyés","Modifier : l'objet et le texte (pas le destinataire ni le PDF)","Créer un brouillon : dossier Brouillons d'Outlook","Ne pas envoyer : la signature reste enregistrée") 14 $INK 4
$null = Card $s 330 316 582 176 $TW
$null = Txt $s 350 328 540 22 'À savoir' 18 $WARN $true
$null = Bul $s 350 358 546 130 @("Rien ne part sans clic ; la date d'envoi est tracée sur l'attestation","Sans connecteur : fichier .eml à ouvrir dans Outlook","Outlook sur le web ne s'ouvre pas depuis la page : brouillon ou .eml","Aperçu du message dans la fenêtre ; connecteur indisponible : raison claire, Copier le texte, .eml") 14 $INK 4

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
$s = NewSlide 'Onglet 7' 'Archives PDF : retrouver une preuve' 'Le PDF signé de chaque attestation, à tout moment' "Le PDF archivé est stocké dans la base (collection documents), de 6 à 15 Ko. Télécharger restitue exactement le fichier produit à la signature. Dans cette capture, trois attestations signées attendent encore leur archivage."
$null = ShotAuto $s 'archives.png' 48 126 570 1440 @(@(1,284,238),@(2,284,433),@(3,680,433),@(4,389,601),@(5,1184,562),@(6,1325,562))
$null = Card $s 668 126 244 366
$null = Txt $s 686 138 210 22 'Repères' 17 $NAVY $true
Legend $s 686 168 210 @('Indicateurs : archivés, à archiver, espace, dernière signature','Filtres : tous, archivés, à archiver','Période (date de signature)','Cycle de vie de l''attestation','Statut : Archivé ou Non archivé','Télécharger / Archiver le PDF') 12 40

# ================= 18. Perdu =================
$s = NewSlide 'Onglet 4' 'Matériel perdu : la synthèse' 'Perdu, volé ou abîmé, par type d''équipement' "Liste déroulante pour choisir le type ; les deux premiers indicateurs suivent le filtre. La date de perte n'est volontairement pas affichée."
$null = ShotAuto $s 'lost.png' 48 126 570 1440 @(@(1,294,217),@(2,493,422),@(3,790,422),@(4,435,546),@(5,1218,520),@(6,977,422))
$null = Card $s 668 126 244 366
$null = Txt $s 686 138 210 22 'Repères' 17 $NAVY $true
Legend $s 686 168 210 @('Indicateurs : unités, déclarations, type le plus touché, ce mois-ci','Type de matériel (liste déroulante)','Période','Une pastille par déclaration','Jauge : part des pertes','Exporter CSV') 12 40
$null = Txt $s 686 424 210 62 "Au-delà de six pastilles : « + N autres »." 12 $SOFT

# ================= Inventaire : le stock =================
$s = NewSlide 'Onglet 5' 'Inventaire : la vue d''ensemble en anneaux' 'Un anneau par type de matériel : le stock au centre, la répartition autour' "Stock = acheté − chez les personnes − restitutions en attente de signature − perdu ; jamais négatif (affiché à 0). Seuls les achats se saisissent : prêts, retours signés et pertes sont lus dans les fiches. À commander = stock à 0 (un type sans stock ne peut plus être prêté). Le filtre Type de matériel limite aussi la vue d'ensemble à ce type. Les types sans achat enregistré affichent un tiret. Chargeur et Écran sont des familles : un seul anneau additionne leurs modèles. Survoler une pastille de la légende (À commander, En stock, Chez les personnes, À signer, Perdu) change le chiffre au centre de chaque anneau ; un clic épingle la couleur. Le volet Achats et la liste détaillée par type ont été retirés."
$null = ShotAuto $s 'inv_stock.png' 48 126 450 1440 @(@(1,310,173),@(2,670,173),@(3,1296,173),@(4,541,360),@(5,936,251),@(6,554,251),@(7,541,454))
$null = Card $s 520 126 392 366
$null = Txt $s 538 138 356 22 'Repères' 17 $NAVY $true
Legend $s 538 168 356 @('Deux volets : Stock, Prix et dépenses','Période, type de matériel et Factures','Enregistrer un achat','Un anneau par type : le stock au centre','Légende : À commander, stock, chez les personnes, à signer, perdu','Survol d''une pastille : le chiffre central change (prêtés, à commander…) ; un clic l''épingle','Sous l''anneau : la quantité achetée') 13 40

# ================= Inventaire : la carte d'un type =================
$s = NewSlide 'Onglet 5' 'Inventaire : la carte détaillée d''un type' 'Au survol d''un anneau, chaque ligne mène à la page concernée' "Au survol d'un anneau (ou au toucher sur tablette) s'affiche une carte : acheté au total, en stock, chez les personnes, à signer, perdu ou abîmé, à remettre aux personnes. Chaque ligne est cliquable : Contrôle des prix (acheté), Emprunter (en stock, à remettre), Restituer (chez les personnes), Attestations (à signer), Matériel perdu. Pour Chargeur et Écran, la carte détaille aussi chaque modèle."
$null = ShotAuto $s 'invring.png' 48 126 450 1500 @(@(1,456,36),@(2,540,84),@(3,465,124),@(4,525,283),@(5,457,368),@(6,515,329))
$null = Card $s 520 126 392 366
$null = Txt $s 538 138 356 22 'Repères' 17 $NAVY $true
Legend $s 538 168 356 @('Nom du type et état (stock suffisant, à commander)','Acheté au total : mène au contrôle des prix','En stock : mène à Emprunter','À remettre aux personnes : mène à Emprunter','Détail par modèle (familles Chargeur et Écran)','Cliquer sur une ligne pour y aller') 13 36
$null = Txt $s 538 392 356 18 'Au toucher' 14 $NAVY $true
$null = Txt $s 538 412 356 74 "Sur tablette, un appui sur l'anneau affiche la carte ; un second appui sur une ligne ouvre la page. L'anneau change de curseur (main) au survol." 12 $SOFT

# ================= Inventaire : prix et dépenses =================
$s = NewSlide 'Onglet 5' 'Inventaire : prix et dépenses' 'Tous les prix du matériel au même endroit, sans lien avec les prêts' "Les montants viennent uniquement des entrées de stock : ils ne changent jamais avec un prêt, un retour ou une perte. Chaque mois du graphique est cliquable : un clic filtre la page sur ce mois, un second clic retire le filtre. Le contrôle des prix signale une entrée dont le prix unitaire s'écarte d'au moins 25 % de la moyenne du type."
$null = ShotAuto $s 'inv_price.png' 48 126 400 1440 @(@(1,292,275),@(2,378,590),@(3,1224,612),@(4,619,401))
$null = Card $s 470 126 442 366
$null = Txt $s 488 138 406 22 'Repères' 17 $NAVY $true
Legend $s 488 168 406 @('Indicateurs : total dépensé, ce mois-ci, type le plus coûteux, entrées sans prix','Répartition des dépenses par type (anneau)','Dépenses par mois : chaque mois est cliquable','Deux boutons : Prix par type de matériel, Contrôle des prix entrée par entrée') 13 46
$null = Txt $s 488 372 406 18 'Aussi dans ce volet' 14 $NAVY $true
$null = Txt $s 488 392 406 96 "Entrées à compléter (achats sans prix) · contrôle des prix entrée par entrée (écart à la moyenne) · filtre de période et de type · export CSV · rapport PDF ou Excel de l'Inventaire." 12 $SOFT

# ================= Inventaire : achats, marque, machines, factures =================
$s = NewSlide 'Onglet 5' 'Inventaire : enregistrer un achat' 'Importer la facture (remplissage automatique) ou saisir à la main' "À l'achat, le type de matériel se choisit dans la liste ; pour Chargeur (USB-C, HP, Asus, Mac, Acer, Dell) et Écran (Dell, MiniFire, Samsung), la liste Modèle ou Marque est obligatoire. Pour un ordinateur portable, la liste Marque (Acer, HP, Asus, Lenovo, Dell) est obligatoire. On saisit ensuite le nom du PC et le numéro de série de chaque exemplaire (60 au plus) : ils sont proposés à l'emprunt (Ordinateur portable 1, 2, …), affichés dans Restituer et sur l'attestation. En haut de la fenêtre, Importer la facture : un PDF glissé ou choisi remplit le formulaire (type, marque ou modèle, quantité, date, prix TTC TVA comprise, fournisseur, numéros de série) ; chaque champ rempli est contouré et reste modifiable, et ce qui n'est pas trouvé est signalé. Seuls les PDF contenant du texte sont lus. En dessous, Ou saisir à la main. Le bouton Factures permet de relire, télécharger et rechercher les factures."
$null = ShotAuto $s 'buyfam.png' 48 126 276 1200 @()
$null = ShotAuto $s 'invimp.png' 336 126 250 1100 @()
$null = Card $s 632 126 280 366
$null = Txt $s 650 138 244 22 'À retenir' 17 $NAVY $true
$null = Bul $s 650 168 246 316 @("Chargeur et Écran : modèle ou marque obligatoire","Ordinateur portable : marque obligatoire (Acer, HP, Asus, Lenovo, Dell)","Nom du PC et n° de série par exemplaire (facultatif)","1 · Saisir à la main : tout reste modifiable","2 · Ou importer la facture (en bas) : le PDF remplit le formulaire","Prix total TTC (TVA comprise) lu ou calculé : à vérifier","Image ou PDF scanné : joint, mais à saisir à la main","Bouton Factures : lire, télécharger, rechercher") 12 $INK 6

# ================= Inventaire : acheter puis prêter =================
$s = NewSlide 'Onglet 5' 'Acheter, puis prêter : le stock suit' 'Un matériel sans stock est grisé dans Emprunter' "Enregistrer un achat : type, nature (achat ou retrait), quantité, date, prix facultatif : par unité ou au total, TTC ou HT, taux de TVA au choix (liste, autre taux ou TTC seul) ; les quatre prix se calculent entre eux, note. Un retrait supérieur au stock est refusé. Dans Emprunter, un type dont le stock est à 0 est grisé avec la pastille « Plus en stock » ; « Sélectionner tout » l'ignore et l'enregistrement est refusé s'il dépasse le stock. Un type sans achat enregistré n'est pas limité."
$null = ShotAuto $s 'inv_modal.png' 48 126 200 540 @()
$null = ShotAuto $s 'nostock.png' 264 126 330 880 @()
$null = Card $s 614 126 298 366
$null = Txt $s 632 138 262 22 'Comment ça marche' 17 $NAVY $true
$null = Bul $s 632 168 266 320 @("Enregistrer un achat : quantité, date, prix (facultatif)","Prix par unité ou total, TTC ou HT, taux de TVA : liés par la quantité","Retrait refusé s'il dépasse le stock","Emprunter : « Plus en stock », ligne grisée et non cochable","Retour signé : le matériel revient en stock","Perte : le matériel sort du stock") 13 $INK 6
# ================= 19. Activités =================
$s = NewSlide 'Onglet 8' 'Activités récentes : le journal' 'Tout ce qui a été enregistré, du plus récent au plus ancien' "Les quatre indicateurs servent aussi de filtres ; un second clic retire le filtre. 50 événements affichés à la fois. La mention de signature figure sous l'événement « attestation signée »."
$null = ShotAuto $s 'activity.png' 48 126 570 1440 @(@(1,284,238),@(2,504,433),@(3,926,433),@(4,360,477),@(5,432,521))
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
$s = NewSlide 'Transversal' 'Codes QR : étiquettes' 'Identifier physiquement chaque équipement' "Les étiquettes se génèrent depuis la fiche d'une personne. Le scanner de QR (caméra, photo, saisie) a été retiré de l'interface à la version 7.1 : aucun bouton « Scanner un QR » n'existe plus. Bibliothèque qrcode-generator chargée à la demande."
$null = Card $s 48 126 430 230
$null = Txt $s 66 140 394 22 'Étiquettes QR (PDF)' 18 $NAVY $true
$null = Bul $s 66 172 394 300 @("Bouton dans la fiche d'une personne (Emprunter ou Restituer)","24 étiquettes par page A4 (3 × 8)","Chaque étiquette : QR, libellé, code (ex. AP-1), personne, date","À imprimer, découper et coller sur le matériel") 15 $INK 8
$null = Card $s 494 126 418 230 $TW
$null = Txt $s 512 140 382 22 'Scanner : retiré' 18 $WARN $true
$null = Bul $s 512 172 382 300 @("La lecture par caméra, photo ou saisie du code n'existe plus dans l'interface","Plus de bouton « Scanner un QR » : Accueil, menu, Emprunter, Restituer, recherche rapide","Le code QR d'une étiquette contient l'identifiant de l'équipement (ex. AP-1)") 15 $INK 8
# ================= Panneau latéral =================
$s = NewSlide 'Transversal' 'Panneau latéral sur grand écran' 'Garder la liste sous les yeux pendant que l''on remplit la fiche' "À partir de 1 480 px de large, la fiche s'affiche à droite ; une confirmation est demandée avant d'abandonner un matériel coché mais non enregistré. En dessous de 1 480 px, la fiche reste une fenêtre."
$null = ShotAuto $s 'sidepanel.png' 48 126 590 1700 @()
$null = Card $s 662 126 250 366
$null = Txt $s 678 140 218 22 'Comportement' 18 $NAVY $true
$null = Bul $s 678 172 222 310 @("≥ 1 480 px : fiche à droite, liste cliquable","Autre personne : la fiche change","Matériel coché non enregistré : confirmation","< 1 480 px : la fiche reste une fenêtre") 13 $INK 8
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
  @('Connecteur absent, expiré ou refusé',$WARN,$TW,"Même fenêtre : aperçu du message, Copier le texte, fichier .eml (PDF joint), Envoyer maintenant (retente la connexion) et raison de l'échec affichée."),
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
$null = Bul $s 348 168 238 320 @("Système : suit l'appareil","Clair et Sombre","Minuit, Ardoise, Noir (sombres)","Océan et Forêt (sombres teintés)","Crème (clair chaud)","Contraste : noir sur blanc, lecture maximale","eXalt et eXalt nuit : charte de l'entreprise (Poppins, bleu eXalt, logo)","Finition premium : relief, jauges et chiffres animés (désactivés si votre système demande moins d'animations)") 13 $INK 5
$null = Card $s 622 126 290 170
$null = Txt $s 640 138 254 22 'Couleurs d''accent' 18 $NAVY $true
$null = Bul $s 640 168 256 126 @("Orange (défaut), Bleu, Violet","Turquoise, Rose, Graphite, eXalt (bleu de la charte)","Boutons, filtres actifs, liens et menu") 14 $INK 5
$null = Card $s 622 310 290 182 $TI
$null = Txt $s 640 322 254 22 'Où le choisir' 18 $BLUE $true
$null = Txt $s 640 352 254 134 "Bouton « Thème » en bas du menu latéral : un panneau propose les thèmes et les couleurs. 84 combinaisons possibles ; un thème eXalt sélectionne l'accent eXalt." 14 $INK

# ================= Recherche rapide =================
$s = NewSlide 'Transversal' 'Recherche rapide et raccourcis' 'Ctrl K : une personne, un équipement, du stock, un achat ou une page' "Le bouton Rechercher figure en haut de chaque page ; Ctrl K ouvre la même fenêtre, qui cherche dans toute l'application (stock, achats et factures, PC et Mac, attestations, archives, pertes, activités  ; voir la diapositive suivante). Un code (AP-1) ou un numéro de série affiche l'équipement et ouvre la fiche de la personne. Entrée ouvre l'action la plus utile sur une personne, Maj + Entrée l'autre."
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
# ================= Recherche rapide globale =================
$s = NewSlide 'Transversal' 'Recherche rapide : tout trouver, partout' 'Une seule barre pour le stock, les achats, les machines, les personnes et les pages' "La barre de recherche rapide (Ctrl K ou bouton Rechercher) cherche dans toute l'application : types de matériel avec leur stock, achats et factures (type, date, montant, note, fournisseur, nom de fichier, marque, nom ou numéro de série), PC et Mac, personnes, et le nombre de résultats dans Attestations, Archives PDF, Matériel perdu et Activités récentes. Un clic ouvre la page déjà filtrée. Les boutons de l'Inventaire (Stock, Prix et dépenses, Factures, Enregistrer un achat) sont aussi proposés comme actions."
$null = ShotAuto $s 'cmdkglobal.png' 48 126 400 1200 @()
$null = Card $s 470 126 442 366
$null = Txt $s 488 138 406 22 'Ce que la barre trouve' 17 $NAVY $true
$null = Bul $s 488 168 410 316 @("Inventaire : un type de matériel, son stock, acheté, prêté, perdu","Achats et factures : les cinq plus récents qui correspondent","PC / Mac : par nom de machine, numéro de série ou marque","Personnes : Entrée ouvre l'action la plus utile","Dans les pages : nombre de résultats par page, un clic pour l'ouvrir déjà filtrée","Actions : boutons de l'Inventaire, rapport, thème, visite guidée") 13 $INK 6

# ================= Guide et démonstrations =================
$s = NewSlide 'Transversal' 'Guide et démonstrations dans chaque page' 'Un bouton « Guide » : mode d''emploi détaillé et démonstration sur l''écran réel' "Le bouton Guide figure tout à gauche des boutons d'en-tête de chaque page. Il ouvre un guide en rubriques (par exemple Emprunter : Faire un emprunt, Cas particuliers, Historique et corrections), un encadré À retenir et une rubrique Partout commune à toutes les pages. Le bouton orange lance une démonstration : le vrai écran est mis en surbrillance étape par étape, avec une carte explicative (Précédent, Suivant, Passer ou la croix). La démonstration ouvre une fiche, coche une ligne et affiche les vraies fenêtres, mais n'enregistre rien : tout est refermé à la fin ou avec Échap. Démonstrations : Emprunter, Restituer sans attestation, Restituer avec attestation par email (aperçu, destinataire, signature, message), Enregistrer un achat, Prix et dépenses, tour de l'Accueil."
$null = ShotAuto $s 'guide.png' 48 126 232 930 @()
$null = ShotAuto $s 'demo.png' 292 126 330 2250 @()
$null = Card $s 292 368 330 124
$null = Txt $s 308 376 300 22 'Une démonstration' 17 $NAVY $true
$null = Bul $s 308 402 304 88 @("Le vrai écran est entouré, étape par étape","Précédent, Suivant, Passer ou la croix ✕","N'enregistre rien : tout est refermé à la fin") 13 $INK 5
$null = Card $s 642 126 270 366
$null = Txt $s 658 138 238 22 'Dans chaque guide' 17 $NAVY $true
$null = Bul $s 658 168 240 320 @("Rubriques en onglets, étapes numérotées","« À retenir » : un à trois points importants","Rubrique « Partout » : recherche, raccourcis, période, rapport…","Six démonstrations : emprunt, restitution (avec et sans attestation), achat, prix, Accueil","Dix pages couvertes") 13 $INK 6

# ================= Rapports PDF et Excel =================
$s = NewSlide 'Transversal' 'Rapports PDF et Excel' 'Un état à partager, en un clic, propre à la page où l''on se trouve' "Les rapports sont créés dans le navigateur (jsPDF et autotable pour le PDF, ExcelJS pour Excel, chargées à la demande). Ils sont réservés aux rôles Administrateur full, Super utilisateur et Lecteur."
$null = ShotAuto $s 'reportmodal.png' 48 126 190 500 @()
$null = ShotAuto $s 'reportpdf.png' 254 126 232 793 @()
$null = Card $s 506 126 406 170
$null = Txt $s 524 138 370 22 'Quatre clics' 18 $NAVY $true
$null = Bul $s 524 168 372 124 @("Bouton Rapport, en haut à droite de chaque page","Choisir la période, la personne (facultatif) et le format : PDF ou Excel","Générer le rapport, puis confirmer le téléchargement, ou « Envoyer par email »") 13 $INK 4
$null = Card $s 506 310 406 182 $TI
$null = Txt $s 524 322 370 22 'Contenu' 18 $BLUE $true
$null = Bul $s 524 352 372 138 @("PDF : indicateurs, parc par type, attestations à traiter, signatures, pertes, emprunts, situation par personne ; logo eXalt","Excel : 8 feuilles, filtres et première ligne figée") 13 $INK 4

# ================= Un rapport par page =================
$s = NewSlide 'Transversal' 'Un rapport propre à chaque page' 'Le contenu, le titre et le nom du fichier suivent la page choisie' "Le bouton Rapport figure en haut à droite d'Accueil, Emprunter, Restituer, Attestations, Archives PDF, Matériel perdu et Inventaire (masqué pour un Agent IT). La période s'applique à l'activité de la page ; les situations (parc, attestations en attente) sont celles d'aujourd'hui. Fichiers : rapport_gpiexalt_emprunts, restitutions, attestations, archives, perdu ; l'Accueil produit le rapport complet."
$null = ShotAuto $s 'reportpage.png' 48 126 180 500 @()
$null = ShotAuto $s 'reportreturn.png' 244 126 220 793 @()
$null = MakeTable $s 488 126 424 @(104,320) @(@('Page','Contenu du rapport'),
  @('Accueil','Rapport complet : parc par type, attestations, signatures, pertes, emprunts, personnes'),
  @('Emprunter','Emprunts de la période avec retour prévu et état ; matériel à remettre'),
  @('Restituer','Retours en retard, retours prévus, parc par type, situation par personne'),
  @('Attestations','À traiter, envoyées et signées de la période'),
  @('Archives PDF','PDF à archiver ; signatures avec l''état de l''archive'),
  @('Matériel perdu','Pertes par type et déclarations de la période'),
  @('Inventaire','Stock par type, dépenses par type, entrées de stock de la période')) 11 44

# ================= Envoyer le rapport par email =================
$s = NewSlide 'Transversal' 'Envoyer le rapport par email' 'Destinataire saisi à la main avec suggestions : rien ne part sans clic' "Le message part de la boîte Outlook de l'utilisateur par le connecteur Microsoft 365, après un clic sur Envoyer maintenant. Jusqu'à 10 adresses. Au-delà de 2,9 Mo (limite de Microsoft Graph), l'envoi direct est refusé : télécharger le rapport et le joindre à un message. Sans connecteur : fichier .eml."
$null = ShotAuto $s 'reportmail.png' 48 126 230 620 @()
$null = Card $s 300 126 612 170
$null = Txt $s 318 138 576 22 'Dans la fenêtre' 18 $NAVY $true
$null = Bul $s 318 168 576 124 @("Destinataire(s) : saisie à la main, suggestions dès 2 lettres (personnes de GPIEXALT, puis annuaire Microsoft 365)","Objet et message préremplis (titre de la page, date, format, période, expéditeur), modifiables","Pièce jointe : rapport_gpiexalt_AAAA-MM-JJ.pdf ou .xlsx") 14 $INK 5
$null = Card $s 300 310 612 182 $TW
$null = Txt $s 318 322 576 22 'Quatre boutons, rien sans clic' 18 $WARN $true
$null = Bul $s 318 352 576 136 @("Envoyer maintenant : part de votre boîte Outlook, copie dans les éléments envoyés","Créer un brouillon : dossier Brouillons d'Outlook","Télécharger le message : fichier .eml à ouvrir dans Outlook","Annuler : retour à la fenêtre du rapport ; connecteur indisponible : Réessayer la connexion") 14 $INK 5
# ================= Rôles et accès =================
$s = NewSlide 'Rôles et accès' 'Cinq rôles, du collaborateur à l''administrateur' 'Le partage de la page fixe la limite ; le rôle précise ce que chacun voit et fait' "Seul « Peut consulter » interdit techniquement toute écriture ; les rôles Agent IT et Lecteur restreignent l'affichage. Les droits réels passeront par Entra ID."
$null = MakeTable $s 48 126 620 @(130,290,110,90) @(@('Rôle','Ce que la personne voit et fait','Partage','Garantie'),
  @('Administrateur full','Tout, y compris Rôles et accès','Modifier','Plateforme'),
  @('Super utilisateur','Toutes les opérations, rapports et exports (rôle par défaut)','Interagir','Affichage'),
  @('Agent IT','Prêts, restitutions, attestations, signatures ; ni pertes ni rapports','Interagir','Affichage'),
  @('Lecteur','Tous les écrans et les rapports, sans saisie','Interagir','Affichage'),
  @('Collaborateur','Mon matériel et ses PDF, en consultation','Consulter','Plateforme')) 12 42
$tgb = Box $s 48 396 620 96 $TI 5 $BLUE
SetText $tgb "Recherche par prénom et nom (début de mots accepté, nom complet affiché). Tous les utilisateurs : rattache chaque fiche au compte correspondant, sans toucher aux rôles attribués à la main ; un bilan liste les fiches sans compte trouvé." 13 $INK $false 1 3
$tgb.TextFrame.MarginLeft = 18
$null = ShotAuto $s 'roles.png' 692 126 220 1178 @()

# ================= Mon matériel =================
$s = NewSlide 'Rôles et accès' 'Mon matériel : la vue du collaborateur' 'Chaque collaborateur consulte son propre matériel, sans pouvoir modifier' "Un compte non rattaché voit « Compte non rattaché » et son identifiant, à communiquer à l'administrateur."
$null = ShotAuto $s 'mine.png' 48 126 306 1178 @()
$null = Card $s 410 126 502 190
$null = Txt $s 428 138 466 22 'Ce que voit le collaborateur' 18 $NAVY $true
$null = Bul $s 428 168 466 144 @("Quatre indicateurs : chez moi, à signer, restitués, perdu","Son matériel avec l'état de chaque équipement","Ses attestations (bouton PDF pour les télécharger), ses prêts et ses pertes déclarées","Aucune saisie ni signature","Menu : Mon matériel, Mes prêts, Mes attestations, Matériel perdu") 13 $INK 4
$null = Card $s 410 330 502 162 $TW
$null = Txt $s 428 342 466 22 'Protection des données' 18 $WARN $true
$null = Bul $s 428 372 466 116 @("Résumé personnel tenu à jour à chaque action sur sa fiche (mine/compte)","Aucune donnée des autres personnes : la base refuse les autres fiches, rien n'est affiché","Rattachement automatique par le nom, ou par l'identifiant du compte") 13 $INK 4

# ================= Annuaire Microsoft 365 =================
$s = NewSlide 'Rôles et accès' 'Annuaire Microsoft 365 : importer les utilisateurs' 'Les vrais collaborateurs dans Emprunter et Restituer, sans rien écrire dans Entra ID' "Lecture seule via le connecteur Microsoft 365 (list-users, list-groups, list-group-members). Le parc réel n'est pas importé : le connecteur ne donne pas accès à Intune. Les données importées sont personnelles : validation du DPO avant un import complet."
$null = ShotAuto $s 'dir.png' 48 126 540 1114 @()
$null = Card $s 48 366 540 126
$null = Txt $s 66 376 504 22 'Trois façons d''importer' 16 $NAVY $true
$null = Bul $s 66 402 504 88 @("Tout l'annuaire : comptes actifs, membres, avec email","Quelques personnes : recherche par nom, pour un essai","Un groupe Entra ID : moins de 999 membres") 13 $INK 3
$null = Card $s 612 126 300 366
$null = Txt $s 630 138 264 22 'Après l''import' 18 $NAVY $true
$null = Bul $s 630 168 266 320 @("Copiés : nom, fonction, email, mobile","Kit standard de 13 équipements par personne (dont HDMI, multiprise, clavier, chargeur Mac), sans emprunt","Synchroniser à nouveau : arrivées ajoutées, emprunts conservés","Un compte retiré reste listé pour restituer le matériel","« Tous les utilisateurs » est trop grand (plus de 999)") 13 $INK 8

# ================= Confort d'utilisation =================
$s = NewSlide 'Confort' 'Confort d''utilisation' 'Des détails qui font gagner du temps au quotidien' "Barre de progression et compteurs animés au chargement, impression en noir sur blanc, focus géré dans les fenêtres, animations réduites si le système le demande. Thème, couleur, densité et visite guidée sont mémorisés par navigateur."
$cf = @(@('Recherche et Accueil',"Recherche collante en haut ; bouton Accueil en évidence sur chaque page, avec Rapport et Rechercher.",$BLUE),
  @('Fiche et recherche',"Champ de recherche pour changer de personne, flèches précédent et suivant, Alt + flèches.",$PURPLE),
  @('Dates rapides',"Aujourd'hui et Hier sous les champs de date ; confirmation si du matériel coché n'est pas enregistré.",$TEAL),
  @('Fenêtres soignées',"Rapport, signature, message, recherche et fiche : même relief ; listes déroulantes lisibles dans tous les thèmes.",$WARN),
  @('Densité et thèmes',"Confortable ou compacte ; douze thèmes (dont eXalt) et sept couleurs, avec un fondu doux au changement.",$GOOD),
  @('Visite guidée',"Onze étapes à la première ouverture (stock compris), relançable depuis l'aide (?) ou la recherche rapide.",$ACC))
for($i=0;$i -lt 6;$i++){
  $x = 48 + ($i % 3)*297; $y = 128 + [math]::Floor($i / 3)*190
  $null = Card $s $x $y 270 172
  $null = Circle $s ($x+18) ($y+18) 16 $cf[$i][2] '' 8
  $null = Txt $s ($x+44) ($y+14) 214 26 $cf[$i][0] 16 $NAVY $true
  $null = Txt $s ($x+18) ($y+56) 236 108 $cf[$i][1] 14 $INK
}

# ================= Dates, listes et fichiers =================
$s = NewSlide 'Confort' 'Dates, listes et fichiers' 'Un calendrier et des listes aux couleurs de l''application, un téléchargement expliqué' "Le calendrier remplace le sélecteur du navigateur : mois précédent et suivant, choix rapide du mois et de l'année, Aujourd'hui, Effacer, dates interdites grisées. La fenêtre « Autoriser cet artefact à télécharger un fichier ? » est celle de la plateforme Claude : elle ne peut pas être modifiée. L'application affiche une carte « Fichier prêt » (format, nom, taille) et propose Enregistrer si la demande est annulée par mégarde."
$null = ShotAuto $s 'calendar.png' 48 124 290 840 @()
$null = ShotAuto $s 'dlcard.png' 354 124 340 800 @()
$null = Card $s 710 124 202 150
$null = Txt $s 726 134 172 22 'Listes' 16 $NAVY $true
$null = Txt $s 726 160 172 110 "Menu aux couleurs du thème, option cochée, navigation au clavier (flèches, Entrée, Échap)." 12 $SOFT
$null = Card $s 48 392 282 98
$null = Txt $s 64 400 252 20 'Calendrier' 15 $NAVY $true
$null = Txt $s 64 424 252 64 "Tous les champs de date : remise, retour prévu, signature, période, achat." 12 $INK
$null = Card $s 346 392 282 98
$null = Txt $s 362 400 252 20 'Carte « Fichier prêt »' 15 $NAVY $true
$null = Txt $s 362 424 252 64 "S'affiche pendant la confirmation d'enregistrement de la plateforme." 12 $INK
$null = Card $s 644 392 268 98 $TW
$null = Txt $s 660 400 238 20 'Annulé par mégarde ?' 15 $WARN $true
$null = Txt $s 660 424 238 64 "Le bouton Enregistrer de la carte relance l'enregistrement, sans régénérer le fichier." 12 $INK
# ================= 18. Téléphone et tablette =================
$s = NewSlide 'Compatibilité' 'Tous les appareils, tous les écrans' 'Téléphone, tablette et PC : une seule application, rapide pour 1 800 collaborateurs' "Mise en page vérifiée de 320 à plus de 2 000 px dans des fenêtres simulées (pas encore sur de vrais iPhone, Android ou iPad : cas CR-74). Barre d'onglets en bas sur téléphone et tablette en portrait, menu latéral sur PC et tablette en paysage. Écrans tactiles : zones de 36 à 40 px, champs à 16 px (pas de zoom automatique sur iPhone), hauteur réelle de l'écran (barre d'adresse de iOS Safari et de Chrome Android). Un bandeau prévient d'un navigateur trop ancien. Rapidité mesurée sur 1 800 personnes fictives avec un historique chacune : Restituer environ 120 ms (790 ms avant optimisation), fiche environ 110 ms, Accueil environ 240 ms, Inventaire 220 à 270 ms ; calculs mis en cache et recherche rapide avec délai de frappe de 0,1 s."
$ph = Box $s 90 120 214 384 '14183A' 5 $null 0.12; Shadow $ph 20 6 0.7
$null = $s.Shapes.AddPicture((Join-Path $imgDoc 'mobile.png'), 0, -1, 109, 132, 176, 360)
$null = Card $s 340 126 572 112
$null = Txt $s 360 134 520 22 'Téléphone et tablette' 17 $NAVY $true
$null = Bul $s 360 160 530 78 @("Barre d'onglets en bas, avec les compteurs ; boutons d'action fixés en bas des fenêtres","Zones à toucher de 36 à 40 px ; champs à 16 px (pas de zoom automatique sur iPhone)","Hauteur réelle de l'écran : barre d'adresse de iOS Safari et de Chrome Android") 12 $INK 3
$null = Card $s 340 248 572 104
$null = Txt $s 360 256 520 22 'PC et grands écrans' 17 $NAVY $true
$null = Bul $s 360 282 530 70 @("Menu latéral repliable ; boutons d'en-tête qui passent à la ligne quand la place manque","Raccourcis 1 à 8 et Ctrl K ; mise en page vérifiée de 320 à plus de 2 000 px") 12 $INK 3
$null = Card $s 340 362 572 130
$null = Txt $s 360 370 520 22 'Navigateurs et rapidité' 17 $NAVY $true
$null = Bul $s 360 396 530 94 @("Chrome et Edge 111, Safari 16.4, Firefox 121 ou plus ; bandeau d'information sur un navigateur plus ancien","1 800 personnes (test) : Restituer ≈ 120 ms, fiche ≈ 110 ms, Accueil ≈ 240 ms, Inventaire ≈ 220 à 270 ms","Calculs mis en cache ; recherche rapide avec délai de frappe ; listes de 50 lignes") 12 $INK 3

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
$s = NewSlide 'Gouvernance' 'Règles de gestion appliquées' 'Quarante-six règles dans le dossier PDF ; les six essentielles ici' "RG-01 à RG-46 : chapitre 5 du dossier PDF."
$rg = @(@('RG-01 · Disponibilité',"Un équipement ne peut être emprunté que s'il est disponible : jamais emprunté, ou redevenu libre.",$BLUE),
  @('RG-03 · Signature',"Une attestation ne passe à « Signée » que par l'action « Marquer comme signée ».",$GOOD),
  @('RG-34 et 35 · Stock',"Stock = acheté − chez les personnes − en attente − perdu, jamais négatif. Sans stock, un matériel ne peut plus être prêté.",$TEAL),
  @('RG-17 et 18 · Rôles',"Cinq rôles ; un collaborateur ne voit que son matériel et ses documents, en consultation seule.",$PURPLE),
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
$null = Txt $s 358 166 236 68 "people : prêts, attestations, pertes, commentaires`rdocuments : PDF signésdirectory, roles, links, mine, inventory" 13 'CADCFC'
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
$null = Bul $s 70 176 378 310 @("Accès réservé aux personnes avec lesquelles l'application est partagée","Cinq rôles ; un collaborateur ne voit que son matériel et ses PDF (Mon matériel)","Chaque événement garde son auteur et sa date","PDF signés conservés tels que produits","Connecteur Microsoft 365 : consentement de chaque utilisateur, jetons jamais visibles de la page") 15 $INK 9
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
$null = Bul $s 68 392 826 96 @("Rien n'est stocké sur le poste ni dans OneDrive : toutes les personnes ayant accès voient les mêmes données","Une personne n'apparaît dans la base qu'à sa première action ; l'ouverture de l'application n'écrit rien","Limite de la base : 256 Ko par document ; collections directory (annuaire), roles, links, mine et inventory (entrées de stock)") 14 $INK 5

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
$s = NewSlide 'Qualité' 'Plan de recette : état au 30/10/2026' '75 cas : 62 réalisés ou simulés en local ; CR-70 à CR-75 à valider en production' "CR-01 à CR-75 dans le dossier PDF, chapitre 10."
$kp = @(@('57','cas réalisés ou simulés',$GOOD),@('9','cas à réaliser en local',$WARN),@('13','cas prioritaires en production',$BAD))
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
  @('CR-51','Rôles','Chercher « prénom nom » dans Rôles et accès','Nom complet affiché, personne trouvée'),
  @('CR-53','Mon matériel','Compte collaborateur : bouton PDF d''une attestation','PDF téléchargé ; rien des autres personnes'),
  @('CR-54','Message prêt à envoyer','Aperçu, Copier le texte, Envoyer maintenant, Réessayer','Raison claire ; envoi réel après reconnexion')) 9 17
$null = Txt $s 48 486 864 24 "Les 66 cas couvrent toutes les fonctions, dont l'Inventaire : détail au chapitre 10 du dossier PDF." 12 $SOFT

# ================= 30. Risques =================
$s = NewSlide 'Qualité' 'Principaux risques' 'Trente risques suivis ; le plus élevé : parc réel non connecté' "Registre complet : chapitre 11 du dossier PDF."
$null = MakeTable $s 48 122 864 @(70,330,94,370) @(@('Réf.','Risque','Niveau','Mesure / état'),
  @('R-06','Parc réel non connecté : kit standard de 13 équipements par personne','Élevé','Choisir la source du parc (Intune, SharePoint, Excel) avant 1 500 collaborateurs'),
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
SetText $jv "Journal : Doc 1.0 (05/10/2026) première édition  ·  Doc 2.0 / App 81 (05/10/2026) édition entreprise  ·  Doc 3.0 / App 106 (06/10/2026) accueil, période, QR, signature, thèmes  ·  Doc 4.0 / App 128 (07/10/2026) rôles, rapports, recherche  ·  Doc 5.0 / App 147 (08/10/2026) annuaire Entra ID, envoi signé  ·  Doc 6.0 / App 161 (09/10/2026) GPIEXALT, retards, rapport par email  ·  Doc 7.0 / App 167 (12/10/2026) rapport par page  ·  Doc 7.1 / App 180 (12/10/2026) finition premium, Mon matériel PDF, retrait du scanner  ·  Doc 7.2 / App 192 (12/10/2026) rapport par personne, catalogue complété  ·  Doc 7.3 / App 206 (13/10/2026) parcours Par email, menu collaborateur, habillage  ·  Doc 7.4 / App 242 (14/10/2026) Inventaire, prix et dépenses, stock, calendrier, rapport de l'Inventaire  ·  Doc 7.5 / App 299 (16/10/2026) Nouvel emprunt, quantités, machines, achats, factures  ·  Doc 7.6 / App 409 (27/10/2026) relance Teams, recherche globale, anneaux de l'Inventaire, familles Chargeur et Écran  ·  Doc 7.7 / App 417 (28/10/2026) import de facture, À commander  ·  Doc 7.8 / App 452 (30/10/2026) prolonger un prêt, relances Teams par paliers, Guide et démonstrations, prix HT et TVA, compatibilité et 1 800 personnes" 10 'FFFFFF' $false 1 3
$jv.TextFrame.MarginLeft = 20

# ================= 33. Limites =================
$s = NewSlide 'Exploitation' 'Limites connues et évolutions prévues' 'Ce qui n''est pas encore là, et ce qui est prévu' "Données de démonstration, rôles Agent IT et Lecteur appliqués à l'affichage, envoi à relire, SMS vers un numéro de test, pas de sauvegarde externe automatique. La signature enregistrée est déclarée par l'équipe IT, non qualifiée."
$li = @(@('Annuaire et parc',"Annuaire : importé depuis Entra ID. Parc : kit standard de 13 équipements, source réelle à choisir.",$BLUE),
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
$s = NewSlide 'Prochaines étapes' 'Décisions et actions attendues' 'Six actions pour passer à une mise en production encadrée' "Les responsables sont des propositions. Après décision : mise à jour de la documentation en version 7.5."
$null = MakeTable $s 48 126 864 @(46,470,200,148) @(@('N°','Action','Responsable proposé','Échéance'),
  @('1','Désigner le vérificateur et l''approbateur du dossier','DSI','À définir'),
  @('2','Renseigner les contacts de support (niveaux 1 à 3)','DSI','À définir'),
  @('3','Fixer la durée de conservation et la base légale','DPO','À définir'),
  @('4','Valider l''hébergement et le chargement des bibliothèques (jsPDF, ExcelJS…)','RSSI','À définir'),
  @('5','Réaliser l''essai réel du brouillon et des envois Outlook (CR-12, CR-37, CR-47, CR-48) et les cas prioritaires','Administrateur de l''application','À définir'),
  @('6','Choisir la source du parc (Intune, SharePoint, Excel), valider l''annuaire avec le DPO','DSI','À définir')) 14 46
$q = Box $s 48 458 864 38 $TI 5 $BLUE
SetText $q "Après décision, la documentation est mise à jour (version 7.8) : demande « mets à jour la documentation »." 14 $INK $false 1 3
$q.TextFrame.MarginLeft = 18


$pres.SaveAs($pptxOut)
$pres.Export($imgDir, 'PNG', 1600, 900)
$count = $pres.Slides.Count
$pres.Close(); $ppt.Quit()
"OK slides=$count"
