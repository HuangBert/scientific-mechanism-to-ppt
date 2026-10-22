param(
 [Parameter(Mandatory=$true)][string]$PlanPath,
 [Parameter(Mandatory=$true)][string]$OutputDir,
 [ValidatePattern('^[A-Za-z0-9_-]+$')][string]$BaseName='mechanisms',
 [ValidateRange(0.01,4)][double]$Scale=0.75,
 [string]$DefaultFont='Times New Roman',
 [ValidateRange(256,8192)][int]$PngWidth=3072,
 [switch]$Overwrite
)
$ErrorActionPreference='Stop'
$OutputDir=[IO.Path]::GetFullPath($OutputDir)
$Plan=@(Get-Content -LiteralPath $PlanPath -Raw -Encoding utf8 | ConvertFrom-Json)
if($Plan.Count -eq 0){throw 'Empty drawing plan'}
$CanvasWidth=[double]$Plan[0].width;$CanvasHeight=[double]$Plan[0].height
if($CanvasWidth -le 0 -or $CanvasHeight -le 0){throw 'Invalid canvas'}
if([Math]::Max($CanvasWidth,$CanvasHeight)*$Scale -gt 4032){throw 'Canvas exceeds PowerPoint maximum size'}
foreach($sp in $Plan){
 if($sp.width -ne $CanvasWidth -or $sp.height -ne $CanvasHeight){throw 'All slides must share canvas size'}
 foreach($o in $sp.objects){if($o.kind -eq 'text' -and -not $o.font){$o|Add-Member -NotePropertyName font -NotePropertyValue $DefaultFont -Force}}
}
[IO.Directory]::CreateDirectory($OutputDir)|Out-Null
$targets=@((Join-Path $OutputDir ($BaseName+'.pptx')),(Join-Path $OutputDir ($BaseName+'.pdf')))
foreach($sp in $Plan){$targets+=(Join-Path $OutputDir ('Fig'+$sp.figure+'_editable.png'))}
if(-not $Overwrite){foreach($target in $targets){if(Test-Path -LiteralPath $target){throw "Output exists; choose a new version directory or pass -Overwrite: $target"}}}
$HadPowerPoint=[bool](Get-Process -Name POWERPNT -ErrorAction SilentlyContinue)
function RGBValue([string]$hex) {
    return [Convert]::ToInt32($hex.Substring(0,2),16) + 256*[Convert]::ToInt32($hex.Substring(2,2),16) + 65536*[Convert]::ToInt32($hex.Substring(4,2),16)
}
function StyleShape($sh,$o) {
    if ($o.fill) {
        $sh.Fill.Visible=-1
        $sh.Fill.ForeColor.RGB=RGBValue $o.fill
        if ($o.grad) {$sh.Fill.TwoColorGradient(1,1);$sh.Fill.ForeColor.RGB=RGBValue $o.fill;$sh.Fill.BackColor.RGB=RGBValue $o.grad} else {$sh.Fill.Solid()}
    } else {$sh.Fill.Visible=0}
    if ($o.stroke -and $o.sw -gt 0) {$sh.Line.Visible=-1;$sh.Line.ForeColor.RGB=RGBValue $o.stroke;$sh.Line.Weight=[single]($o.sw*$Scale)} else {$sh.Line.Visible=0}
    if ($o.dash) {$sh.Line.DashStyle=4}
    if ($o.round -and $sh.Adjustments.Count -gt 0) {$sh.Adjustments.Item(1)=[single]($o.round/[Math]::Min($o.w,$o.h))}
}
function AddNativeText($sl,$o) {
    $sh=$sl.Shapes.AddTextbox(1,[single]($o.x*$Scale),[single]($o.y*$Scale),[single]($o.w*$Scale),[single]($o.h*$Scale))
    $sh.Name=$o.id
    $tf=$sh.TextFrame2
    $tf.MarginLeft=0;$tf.MarginRight=0;$tf.MarginTop=0;$tf.MarginBottom=0
    $tf.WordWrap=0;$tf.AutoSize=0;$tf.VerticalAnchor=3
    # Parse subscript spans into real editable text runs.
    $raw=[string]$o.text
    $plain='';$spans=@();$cursor=0
    foreach($m in [regex]::Matches($raw,'_\{([^}]+)\}')){
        $plain+=$raw.Substring($cursor,$m.Index-$cursor)
        $spans+=@{start=$plain.Length+1;length=$m.Groups[1].Value.Length}
        $plain+=$m.Groups[1].Value;$cursor=$m.Index+$m.Length
    }
    $plain+=$raw.Substring($cursor)
    $tf.TextRange.Text=$plain
    $tf.TextRange.Font.Name=$o.font;$tf.TextRange.Font.Size=[single]($o.size*$Scale)
    $tf.TextRange.Font.NameAscii=$o.font
    $tf.TextRange.Font.NameFarEast=$o.font
    $tf.TextRange.Font.NameOther=$o.font
    $tf.TextRange.Font.NameComplexScript=$o.font
    $tf.TextRange.Font.Bold=if($o.bold){-1}else{0}
    if($o.italic){$tf.TextRange.Font.Italic=-1}
    $tf.TextRange.Font.Fill.ForeColor.RGB=RGBValue $o.color
    $tf.TextRange.ParagraphFormat.Alignment=if($o.align -eq 'center'){2}elseif($o.align -eq 'right'){3}else{1}
    $tf.TextRange.ParagraphFormat.SpaceBefore=0;$tf.TextRange.ParagraphFormat.SpaceAfter=0
    foreach($sp in $spans){$part=$tf.TextRange.Characters($sp.start,$sp.length);$part.Font.Size=[single]($o.size*$Scale*.68);$part.Font.BaselineOffset=-0.20}
    # PowerPoint initially auto-sizes empty textboxes while margins are assigned.
    # Restore the reference rectangle after formatting to keep center anchors exact.
    $sh.Left=[single]($o.x*$Scale);$sh.Top=[single]($o.y*$Scale)
    $sh.Width=[single]($o.w*$Scale);$sh.Height=[single]($o.h*$Scale)
    if($o.angle){$sh.Rotation=$o.angle}
    return $sh
}
$Ppt = New-Object -ComObject PowerPoint.Application
$Ppt.DisplayAlerts=1
$Deck = $null
try {
    $Deck=$Ppt.Presentations.Add(0)
    foreach($scriptFont in @(1,2,3)){
        $Deck.SlideMaster.Theme.ThemeFontScheme.MajorFont($scriptFont).Name=$DefaultFont
        $Deck.SlideMaster.Theme.ThemeFontScheme.MinorFont($scriptFont).Name=$DefaultFont
    }
    $Deck.PageSetup.SlideWidth=[single]($CanvasWidth*$Scale);$Deck.PageSetup.SlideHeight=[single]($CanvasHeight*$Scale)
    foreach($sp in $Plan){
        $sl=$Deck.Slides.Add($Deck.Slides.Count+1,12)
        $sl.Name="Figure_$($sp.figure)"
        $sl.FollowMasterBackground=0;$sl.Background.Fill.ForeColor.RGB=16777215
        foreach($o in $sp.objects){
            $sh=$null
            switch($o.kind){
                'shape' {
                    $type=if($o.shape -eq 'ellipse'){9}elseif($o.shape -eq 'roundRect'){5}else{1}
                    $sh=$sl.Shapes.AddShape($type,[single]($o.x*$Scale),[single]($o.y*$Scale),[single]($o.w*$Scale),[single]($o.h*$Scale))
                    StyleShape $sh $o
                }
                'text' {$sh=AddNativeText $sl $o}
                'line' {
                    $sh=$sl.Shapes.AddConnector(1,[single]($o.x1*$Scale),[single]($o.y1*$Scale),[single]($o.x2*$Scale),[single]($o.y2*$Scale))
                    $sh.Line.ForeColor.RGB=RGBValue $o.color;$sh.Line.Weight=[single]($o.sw*$Scale)
                    if($o.arrow){$sh.Line.EndArrowheadStyle=2}
                    if($o.startArrow){$sh.Line.BeginArrowheadStyle=2}
                    if($o.dash){$sh.Line.DashStyle=4}
                }
                'path' {
                    $pts=$o.points
                    $builder=$sl.Shapes.BuildFreeform(1,[single]($pts[0][0]*$Scale),[single]($pts[0][1]*$Scale))
                    for($i=1;$i -lt $pts.Count;$i++){
                        $q=$pts[$i]
                        $rounded=$false
                        if(-not $o.fill -and $i -lt $pts.Count-1){
                            $a=$pts[$i-1];$b=$pts[$i+1]
                            $ux=[double]$q[0]-[double]$a[0];$uy=[double]$q[1]-[double]$a[1]
                            $vx=[double]$b[0]-[double]$q[0];$vy=[double]$b[1]-[double]$q[1]
                            $ul=[Math]::Sqrt($ux*$ux+$uy*$uy);$vl=[Math]::Sqrt($vx*$vx+$vy*$vy)
                            if($ul -gt 8 -and $vl -gt 8 -and [Math]::Abs($ux*$vx+$uy*$vy) -lt 0.02*$ul*$vl){
                                $r=[Math]::Min(16,[Math]::Min($ul/3,$vl/3));$ux/=$ul;$uy/=$ul;$vx/=$vl;$vy/=$vl
                                $px=[double]$q[0]-$ux*$r;$py=[double]$q[1]-$uy*$r
                                $ex=[double]$q[0]+$vx*$r;$ey=[double]$q[1]+$vy*$r
                                $builder.AddNodes(0,1,[single]($px*$Scale),[single]($py*$Scale))
                                $builder.AddNodes(1,1,[single](($px+$ux*$r*.552285)*$Scale),[single](($py+$uy*$r*.552285)*$Scale),[single](($ex-$vx*$r*.552285)*$Scale),[single](($ey-$vy*$r*.552285)*$Scale),[single]($ex*$Scale),[single]($ey*$Scale))
                                $rounded=$true
                            }
                        }
                        if(-not $rounded){$builder.AddNodes(0,1,[single]($q[0]*$Scale),[single]($q[1]*$Scale))}
                    }
                    if($o.fill){$builder.AddNodes(0,1,[single]($pts[0][0]*$Scale),[single]($pts[0][1]*$Scale))}
                    $sh=$builder.ConvertToShape()
                    $styles=@{fill=$o.fill;grad=$o.grad;stroke=$o.color;sw=$o.sw;dash=$o.dash}
                    StyleShape $sh $styles
                    if($o.arrow){$sh.Line.EndArrowheadStyle=2}
                }
            }
            if($sh){$sh.Name=$o.id}
        }
        $sl.NotesPage.Shapes.Placeholders.Item(2).TextFrame.TextRange.Text=$sp.notes
        $sl.Export((Join-Path $OutputDir "Fig$($sp.figure)_editable.png"),'PNG',$PngWidth,[int][Math]::Round($PngWidth*$CanvasHeight/$CanvasWidth))
        Write-Output "Rendered Figure $($sp.figure): $($sl.Shapes.Count) editable objects"
    }
    $PptxPath=Join-Path $OutputDir ($BaseName+'.pptx')
    $Deck.SaveAs($PptxPath,24)
    $Deck.SaveAs((Join-Path $OutputDir ($BaseName+'.pdf')),32)
    Write-Output "Saved $PptxPath"
} finally {
    if($Deck){$Deck.Close();[void][Runtime.InteropServices.Marshal]::ReleaseComObject($Deck)}
    if($Ppt){if(-not $HadPowerPoint -and $Ppt.Presentations.Count -eq 0){$Ppt.Quit()};[void][Runtime.InteropServices.Marshal]::ReleaseComObject($Ppt)}
}
