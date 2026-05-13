<#
.SYNOPSIS
    Проверка и автоисправление оформления Word-документа по «Памятке».

.DESCRIPTION
    Использует COM Word.Application. Ничего ставить не нужно.

.PARAMETER Path
    Путь к .docx / .doc. Если не указан — диалог выбора файла.

.PARAMETER Fix
    Если задан — после проверки сразу применит автоисправления и сохранит
    копию рядом с исходным файлом с суффиксом _fixed.docx.

.PARAMETER InPlace
    Применяется вместе с -Fix. Сохраняет изменения в исходный файл
    (предварительно делается резервная копия .bak).

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File Check-WordDoc.ps1 -Path "C:\письмо.docx"
    powershell -ExecutionPolicy Bypass -File Check-WordDoc.ps1 -Path "C:\письмо.docx" -Fix
#>

[CmdletBinding()]
param(
    [string]$Path,
    [switch]$Fix,
    [switch]$InPlace
)

# ---------- параметры из Памятки ----------
$MARGIN_LEFT_MM   = 30
$MARGIN_TOP_MM    = 20
$MARGIN_BOTTOM_MM = 20
$MARGIN_RIGHT_MM  = 10
$MM_TOL           = 0.5
$CM_TOL           = 0.05
$FONT_REQ         = 'Times New Roman'
$INDENT_REQ_CM    = 1.25
$FIX_FONT_SIZE    = 14
$FIX_LINE_SPACING = 1.5

# константы Word
$wdLineSpaceSingle     = 0
$wdLineSpace1pt5       = 1
$wdLineSpaceMultiple   = 5
$wdAlignParagraphCenter= 1
$wdHeaderFooterPrimary = 1
$wdFieldPage           = 33
$wdStatisticPages      = 2

function PointsToMillimeters([double]$pts) { [Math]::Round($pts * 25.4 / 72, 2) }
function PointsToCentimeters([double]$pts) { [Math]::Round($pts * 2.54 / 72, 3) }
function MillimetersToPoints([double]$mm)  { $mm * 72 / 25.4 }
function CentimetersToPoints([double]$cm)  { $cm * 72 / 2.54 }

function Show-OpenDialog {
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Документы Word (*.docx;*.doc)|*.docx;*.doc|Все файлы|*.*'
    $dlg.Title  = 'Выберите документ для проверки'
    if ($dlg.ShowDialog() -ne 'OK') { return $null }
    $dlg.FileName
}

function Collect-Issues($doc) {
    $issues = New-Object System.Collections.Generic.List[object]
    $ps = $doc.PageSetup
    $leftMM   = PointsToMillimeters $ps.LeftMargin
    $topMM    = PointsToMillimeters $ps.TopMargin
    $bottomMM = PointsToMillimeters $ps.BottomMargin
    $rightMM  = PointsToMillimeters $ps.RightMargin

    if ([Math]::Abs($leftMM   - $MARGIN_LEFT_MM)   -gt $MM_TOL) { $issues.Add(@{ Code='MARGIN_L'; Fixable=$true; Text="Левое поле $leftMM мм (нужно 30 мм)." }) }
    if ([Math]::Abs($topMM    - $MARGIN_TOP_MM)    -gt $MM_TOL) { $issues.Add(@{ Code='MARGIN_T'; Fixable=$true; Text="Верхнее поле $topMM мм (нужно 20 мм)." }) }
    if ([Math]::Abs($bottomMM - $MARGIN_BOTTOM_MM) -gt $MM_TOL) { $issues.Add(@{ Code='MARGIN_B'; Fixable=$true; Text="Нижнее поле $bottomMM мм (нужно 20 мм)." }) }
    if ([Math]::Abs($rightMM  - $MARGIN_RIGHT_MM)  -gt $MM_TOL) { $issues.Add(@{ Code='MARGIN_R'; Fixable=$true; Text="Правое поле $rightMM мм (нужно 10 мм)." }) }

    $badFonts   = New-Object System.Collections.Generic.HashSet[string]
    $badSizes   = New-Object System.Collections.Generic.HashSet[string]
    $badSpacing = $false
    $bodyParas  = 0
    $badIndent  = 0

    foreach ($p in $doc.Paragraphs) {
        $txt = $p.Range.Text.Trim()
        if ($txt.Length -le 1) { continue }
        $f = $p.Range.Font
        if ($f.Name -and $f.Name -ne $FONT_REQ) { [void]$badFonts.Add($f.Name) }
        $sz = $f.Size
        if ($sz -is [double] -and $sz -ne 13 -and $sz -ne 14) { [void]$badSizes.Add("$sz") }

        $rule = $p.Format.LineSpacingRule
        $ls   = $p.Format.LineSpacing
        $ok = $false
        switch ($rule) {
            0 { $ok = $true }
            1 { $ok = $true }
            5 { if ($f.Size -gt 0 -and ($ls / $f.Size) -ge 0.95 -and ($ls / $f.Size) -le 1.55) { $ok = $true } }
        }
        if (-not $ok) { $badSpacing = $true }

        $styleName = "$($p.Style.NameLocal)"
        if ($txt.Length -gt 80 -and $styleName -notmatch 'аголовок') {
            $bodyParas++
            $indCm = PointsToCentimeters $p.Format.FirstLineIndent
            if ([Math]::Abs($indCm - $INDENT_REQ_CM) -gt $CM_TOL) { $badIndent++ }
        }
    }

    if ($badFonts.Count -gt 0) { $issues.Add(@{ Code='FONT';    Fixable=$true; Text="Шрифт не $FONT_REQ: $($badFonts -join ', ')." }) }
    if ($badSizes.Count -gt 0) { $issues.Add(@{ Code='SIZE';    Fixable=$true; Text="Размер шрифта вне 13–14 пт: $($badSizes -join ', ')." }) }
    if ($badSpacing)           { $issues.Add(@{ Code='LINESP';  Fixable=$true; Text='Межстрочный интервал вне 1.0–1.5.' }) }
    if ($bodyParas -gt 0 -and $badIndent -gt 0) {
        $issues.Add(@{ Code='INDENT'; Fixable=$true; Text="Абзацный отступ ≠ 1.25 см: $badIndent из $bodyParas абзацев." })
    }

    if ($doc.ComputeStatistics($wdStatisticPages) -ge 2) {
        $hasNumber=$false; $centered=$false; $decor=$false
        foreach ($sec in $doc.Sections) {
            $hdr = $sec.Headers.Item($wdHeaderFooterPrimary)
            foreach ($fld in $hdr.Range.Fields) {
                if ($fld.Type -eq $wdFieldPage) {
                    $hasNumber = $true
                    if ($hdr.Range.ParagraphFormat.Alignment -eq $wdAlignParagraphCenter) { $centered = $true }
                    $t = $hdr.Range.Text
                    if ($t -match '-' -or $t -match 'стр' -or $t -match 'с\.') { $decor = $true }
                }
            }
        }
        if (-not $hasNumber)  { $issues.Add(@{ Code='PAGE_ADD';   Fixable=$true; Text='Документ многостраничный, но нет номеров страниц.' }) }
        else {
            if (-not $centered) { $issues.Add(@{ Code='PAGE_ALIGN'; Fixable=$true; Text='Номер страницы не по центру.' }) }
            if ($decor)         { $issues.Add(@{ Code='PAGE_DECOR'; Fixable=$true; Text='В номере страницы лишние символы.' }) }
        }
    }

    $bodyText = $doc.Range().Text
    if ($bodyText -notmatch '8\(\d{3,5}\)\s?\d') {
        $issues.Add(@{ Code='EXEC'; Fixable=$false; Text='Не найдена отметка об исполнителе (телефон 8(код)номер).' })
    }
    if ($bodyText -match '(?i)для служебного пользования' -and $bodyText -notmatch '(?i)экз\.?\s*№') {
        $issues.Add(@{ Code='DSP'; Fixable=$false; Text='«Для служебного пользования» без номера экземпляра.' })
    }

    return ,$issues
}

function Apply-Fixes($doc, $issues) {
    foreach ($iss in $issues) {
        if (-not $iss.Fixable) { continue }
        switch ($iss.Code) {
            'MARGIN_L'  { $doc.PageSetup.LeftMargin   = MillimetersToPoints $MARGIN_LEFT_MM }
            'MARGIN_T'  { $doc.PageSetup.TopMargin    = MillimetersToPoints $MARGIN_TOP_MM }
            'MARGIN_B'  { $doc.PageSetup.BottomMargin = MillimetersToPoints $MARGIN_BOTTOM_MM }
            'MARGIN_R'  { $doc.PageSetup.RightMargin  = MillimetersToPoints $MARGIN_RIGHT_MM }
            'FONT' {
                foreach ($p in $doc.Paragraphs) {
                    if ($p.Range.Text.Trim().Length -gt 1 -and $p.Range.Font.Name -ne $FONT_REQ) {
                        $p.Range.Font.Name = $FONT_REQ
                    }
                }
            }
            'SIZE' {
                foreach ($p in $doc.Paragraphs) {
                    if ($p.Range.Text.Trim().Length -gt 1) {
                        $sz = $p.Range.Font.Size
                        if ($sz -is [double] -and $sz -ne 13 -and $sz -ne 14) { $p.Range.Font.Size = $FIX_FONT_SIZE }
                    }
                }
            }
            'LINESP' {
                foreach ($p in $doc.Paragraphs) {
                    if ($p.Range.Text.Trim().Length -gt 1) {
                        $p.Format.LineSpacingRule = $wdLineSpaceMultiple
                        $p.Format.LineSpacing     = $p.Range.Font.Size * $FIX_LINE_SPACING
                    }
                }
            }
            'INDENT' {
                foreach ($p in $doc.Paragraphs) {
                    $txt = $p.Range.Text.Trim()
                    $styleName = "$($p.Style.NameLocal)"
                    if ($txt.Length -gt 80 -and $styleName -notmatch 'аголовок') {
                        $p.Format.FirstLineIndent = CentimetersToPoints $INDENT_REQ_CM
                    }
                }
            }
            { $_ -in 'PAGE_ADD','PAGE_DECOR' } {
                foreach ($sec in $doc.Sections) {
                    $hdr = $sec.Headers.Item($wdHeaderFooterPrimary)
                    $hdr.Range.Text = ''
                    $hdr.Range.ParagraphFormat.Alignment = $wdAlignParagraphCenter
                    [void]$hdr.Range.Fields.Add($hdr.Range, $wdFieldPage)
                }
            }
            'PAGE_ALIGN' {
                foreach ($sec in $doc.Sections) {
                    $sec.Headers.Item($wdHeaderFooterPrimary).Range.ParagraphFormat.Alignment = $wdAlignParagraphCenter
                }
            }
        }
    }
}

function Print-Issues($issues) {
    if ($issues.Count -eq 0) {
        Write-Host 'OK. Все автоматические проверки пройдены.' -ForegroundColor Green
        return
    }
    $fixable = ($issues | Where-Object Fixable).Count
    $manual  = $issues.Count - $fixable
    for ($i = 0; $i -lt $issues.Count; $i++) {
        $iss = $issues[$i]
        $tag = if ($iss.Fixable) { '[можно исправить]' } else { '[ручная правка] ' }
        Write-Host ("{0}. {1} {2}" -f ($i + 1), $tag, $iss.Text) -ForegroundColor Yellow
    }
    Write-Host ''
    Write-Host "Автоисправимых: $fixable; требуют ручной правки: $manual" -ForegroundColor Cyan
}

# ---------- основной поток ----------
if (-not $Path) { $Path = Show-OpenDialog }
if (-not $Path -or -not (Test-Path -LiteralPath $Path)) {
    Write-Host 'Файл не выбран. Выход.' -ForegroundColor Yellow
    return
}
$Path = (Resolve-Path -LiteralPath $Path).Path

$word = $null; $doc = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $readOnly = -not $Fix
    $doc = $word.Documents.Open($Path, [ref]$false, [ref]$readOnly)

    Write-Host ''
    Write-Host "Документ: $Path" -ForegroundColor Cyan
    Write-Host ('-' * 60)

    $issues = Collect-Issues $doc
    Print-Issues $issues

    if (-not $Fix -or $issues.Count -eq 0) { return }
    $fixable = $issues | Where-Object Fixable
    if ($fixable.Count -eq 0) {
        Write-Host ''
        Write-Host 'Автоисправление неприменимо — все замечания требуют ручной правки.' -ForegroundColor Yellow
        return
    }

    if (-not $PSBoundParameters.ContainsKey('Fix')) {
        $ans = Read-Host "`nИсправить $($fixable.Count) автоисправимых замечаний? [Y/N]"
        if ($ans -notmatch '^(y|yes|д|да)$') { return }
    }

    Apply-Fixes $doc $issues

    if ($InPlace) {
        $bak = "$Path.bak"
        Copy-Item -LiteralPath $Path -Destination $bak -Force
        $doc.Save()
        Write-Host ''
        Write-Host "Сохранено в исходный файл. Резервная копия: $bak" -ForegroundColor Green
    } else {
        $dir  = Split-Path -Parent $Path
        $name = [IO.Path]::GetFileNameWithoutExtension($Path)
        $out  = Join-Path $dir "$name`_fixed.docx"
        $wdFormatDocumentDefault = 16
        $doc.SaveAs([ref]$out, [ref]$wdFormatDocumentDefault)
        Write-Host ''
        Write-Host "Исправленная копия: $out" -ForegroundColor Green
    }

    $issues2 = Collect-Issues $doc
    Write-Host ''
    Write-Host 'Перепроверка после исправления:' -ForegroundColor Cyan
    Print-Issues $issues2
}
finally {
    if ($doc)  { $doc.Close([ref]$false)   | Out-Null; [Runtime.InteropServices.Marshal]::ReleaseComObject($doc)  | Out-Null }
    if ($word) { $word.Quit()              | Out-Null; [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
