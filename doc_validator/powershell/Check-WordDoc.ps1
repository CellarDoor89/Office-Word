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
    [switch]$InPlace,
    [string]$BlankTemplate
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

function Letterhead-TopOffsetMM($doc, $topMarginMM) {
    # Сумма: верхнее поле + высоты ведущих пустых строк первой таблицы.
    $offset = [double]$topMarginMM
    if ($doc.Tables.Count -eq 0) { return $offset }
    $t = $doc.Tables.Item(1)
    $wdRowHeightAuto = 0
    foreach ($row in $t.Rows) {
        $rowText = ''
        foreach ($cell in $row.Cells) {
            $txt = $cell.Range.Text -replace "`a", '' -replace "`r", ''
            $rowText += $txt
        }
        if ($rowText.Trim().Length -gt 0) { return $offset }
        # пустая — добавляем высоту
        try {
            if ($row.HeightRule -ne $wdRowHeightAuto) {
                $offset += (PointsToMillimeters $row.Height)
            } else {
                $offset += 6.0
            }
        } catch { $offset += 6.0 }
    }
    return $offset
}

function Has-Letterhead($doc) {
    # Документ начинается с таблицы (бланк с гербом)? Тогда верхнее поле
    # интенционально уменьшено и проверяться не должно.
    if ($doc.Tables.Count -eq 0) { return $false }
    $firstTable = $doc.Tables.Item(1)
    foreach ($p in $doc.Paragraphs) {
        if ($p.Range.Text.Trim().Length -eq 0) { continue }
        # первый непустой параграф
        try {
            if ($p.Range.Information(12)) { # wdWithInTable
                $tbl = $p.Range.Tables.Item(1)
                if ($tbl.Range.Start -eq $firstTable.Range.Start) { return $true }
            }
        } catch { }
        return $false
    }
    return $false
}

function Collect-Issues($doc) {
    $issues = New-Object System.Collections.Generic.List[object]
    $ps = $doc.PageSetup
    $leftMM   = PointsToMillimeters $ps.LeftMargin
    $topMM    = PointsToMillimeters $ps.TopMargin
    $bottomMM = PointsToMillimeters $ps.BottomMargin
    $rightMM  = PointsToMillimeters $ps.RightMargin

    $hasLetterhead = Has-Letterhead $doc

    if ([Math]::Abs($leftMM   - $MARGIN_LEFT_MM)   -gt $MM_TOL) { $issues.Add(@{ Code='MARGIN_L'; Fixable=$true; Text="Левое поле $leftMM мм (нужно 30 мм)." }) }
    if ($hasLetterhead) {
        $offsetMM = Letterhead-TopOffsetMM $doc $topMM
        if ($offsetMM -lt ($MARGIN_TOP_MM - $MM_TOL)) {
            $offFmt = '{0:N1}' -f $offsetMM
            $issues.Add(@{ Code='MARGIN_T_LETTERHEAD'; Fixable=$false; Text="Содержимое бланка начинается на $offFmt мм от верха страницы (нужно ≥20 мм). Регистрационный номер при подписи может наложиться на герб/текст бланка." })
        }
    } elseif ([Math]::Abs($topMM - $MARGIN_TOP_MM) -gt $MM_TOL) {
        $issues.Add(@{ Code='MARGIN_T'; Fixable=$true; Text="Верхнее поле $topMM мм (нужно 20 мм)." })
    }
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

    Check-Header $doc $issues
    Check-Blank  $doc $issues

    return ,$issues
}

function Is-RightBlockPara($p) {
    $wdAlignRight   = 2
    $wdWithInTable  = 12
    if ($p.Format.Alignment -eq $wdAlignRight) { return $true }
    $indCm = PointsToCentimeters $p.Format.LeftIndent
    if ($indCm -gt 7) { return $true }
    try {
        if ($p.Range.Information($wdWithInTable)) {
            $cells = $p.Range.Cells
            if ($cells -and $cells.Count -gt 0) {
                if ($cells.Item(1).ColumnIndex -ge 2) { return $true }
            }
        }
    } catch { }
    return $false
}

function Check-Header($doc, $issues) {
    $limit = 60
    $topParas = New-Object System.Collections.Generic.List[object]
    $i = 0
    foreach ($p in $doc.Paragraphs) {
        $i++
        if ($i -gt $limit) { break }
        $topParas.Add($p)
    }
    if ($topParas.Count -eq 0) { return }

    $topText = ''
    $rightParas = New-Object System.Collections.Generic.List[object]
    $rightText = ''
    foreach ($p in $topParas) {
        $topText += "$($p.Range.Text)`n"
        if ((Is-RightBlockPara $p) -and $p.Range.Text.Trim().Length -gt 0) {
            $rightParas.Add($p)
            $rightText += "$($p.Range.Text)`n"
        }
    }

    # --- Гриф ДСП ---
    if ($topText -match '(?i)для служебного пользования') {
        if ($rightText -notmatch '(?i)для служебного пользования') {
            $issues.Add(@{ Code='DSP_POS'; Fixable=$false; Text='«Для служебного пользования» найдено, но не в правом верхнем углу.' })
        }
        if ($topText -notmatch '(?i)экз\.?\s*№\s*\d') {
            $issues.Add(@{ Code='DSP'; Fixable=$false; Text='«Для служебного пользования» без номера экземпляра («Экз. № …»).' })
        }
    }

    # --- Адресат ---
    if ($rightParas.Count -eq 0) {
        $issues.Add(@{ Code='ADRESAT_MISSING'; Fixable=$false; Text='Не найден блок адресата в правой верхней части документа.' })
    } else {
        $adresatParas = $rightParas | Where-Object { $_.Range.Text -notmatch '(?i)для служебного пользования|экз\.?\s*№' }

        $hasFio = ($rightText -cmatch '[А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.\s?[А-ЯЁ]\.') -or `
                  ($rightText -cmatch '[А-ЯЁ][а-яё]{2,}\s+[А-ЯЁ][а-яё]{2,}\s+[А-ЯЁ][а-яё]{2,}')
        $hasOrg = $rightText -match 'ООО|АО|ПАО|ОАО|Министерств|Управлени|Федеральн|Правительств|Администраци|Канцеляри|ФНС|УФНС|ИФНС|Межрегиональн|Инспекци|Департамент|Комитет|Служб'

        if ($adresatParas -and -not ($hasFio -or $hasOrg)) {
            $issues.Add(@{ Code='ADRESAT_MISSING'; Fixable=$false; Text='Блок справа есть, но в нём не найдены ни ФИО, ни наименование организации.' })
        }

        if ($adresatParas -and $adresatParas.Count -ge 2) {
            $aligns  = $adresatParas | ForEach-Object { $_.Format.Alignment } | Select-Object -Unique
            $indents = $adresatParas | ForEach-Object { [Math]::Round((PointsToCentimeters $_.Format.LeftIndent), 1) } | Select-Object -Unique
            if ($aligns.Count -gt 1 -or $indents.Count -gt 1) {
                $issues.Add(@{ Code='ADRESAT_ALIGN'; Fixable=$false; Text='Строки адресата выровнены неодинаково (должны быть по левому краю блока).' })
            }
        }

        $badAdrSpacing = $false
        foreach ($p in $adresatParas) {
            $rule = $p.Format.LineSpacingRule
            $ls   = $p.Format.LineSpacing
            $sz   = $p.Range.Font.Size
            if ($rule -eq 1 -or $rule -eq 2) { $badAdrSpacing = $true; break }
            if ($rule -eq 5 -and $sz -gt 0 -and ($ls / $sz) -gt 1.05) { $badAdrSpacing = $true; break }
        }
        if ($badAdrSpacing) {
            $issues.Add(@{ Code='ADRESAT_SPACING'; Fixable=$false; Text='Межстрочный интервал в адресате не 1.0.' })
        }
    }

    # --- Формат инициалов ---
    if ($rightText -cmatch '[А-ЯЁ][а-яё]+[А-ЯЁ]\.') {
        $issues.Add(@{ Code='INITIALS_BAD'; Fixable=$false; Text='Инициалы записаны слитно с фамилией (нужно «Фамилия И.О.» через пробел).' })
    } elseif (($rightText -cmatch '[А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.') -and -not ($rightText -cmatch '[А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.\s?[А-ЯЁ]\.')) {
        $issues.Add(@{ Code='INITIALS_BAD'; Fixable=$false; Text='У фамилии указан только один инициал (нужно два: «Фамилия И.О.»).' })
    }

    # --- Заголовок «О …»/«Об …» (в пределах всей шапки) ---
    $titleFound = $false
    foreach ($p in $topParas) {
        $t = $p.Range.Text.Trim()
        if ($t.Length -eq 0) { continue }
        if ($t -cmatch '^\s*Об?\s+[а-яёА-ЯЁ]') { $titleFound = $true; break }
    }
    if (-not $titleFound) {
        $issues.Add(@{ Code='TITLE_MISSING'; Fixable=$false; Text='Не найден заголовок к тексту «О …»/«Об …» в шапке (п. 4 Памятки).' })
    }
}

function Get-BlankTemplatePath {
    if ($script:BlankTemplate -and (Test-Path -LiteralPath $script:BlankTemplate)) {
        return (Resolve-Path -LiteralPath $script:BlankTemplate).Path
    }
    $candidates = @(
        Join-Path $PSScriptRoot 'blank_template.txt'
        Join-Path (Split-Path -Parent $PSScriptRoot) 'blank_template.txt'
    )
    foreach ($c in $candidates) {
        if (Test-Path -LiteralPath $c) { return (Resolve-Path -LiteralPath $c).Path }
    }
    return $null
}

function Load-BlankTemplate($path) {
    $items = @()
    foreach ($raw in [IO.File]::ReadAllLines($path, [Text.Encoding]::UTF8)) {
        $line = $raw.Trim()
        if (-not $line -or $line.StartsWith('#')) { continue }
        $optional = $false; $isRegex = $false
        if ($line.StartsWith('?'))  { $optional = $true; $line = $line.Substring(1).TrimStart() }
        if ($line.StartsWith('re:')){ $isRegex  = $true; $line = $line.Substring(3) }
        if (-not $line) { continue }
        $items += [pscustomobject]@{ Pattern = $line; IsRegex = $isRegex; Optional = $optional }
    }
    return ,$items
}

function Get-BlankText($doc) {
    if ($doc.Tables.Count -gt 0) {
        $t = $doc.Tables.Item(1)
        $maxTcs = 0
        foreach ($row in $t.Rows) { if ($row.Cells.Count -gt $maxTcs) { $maxTcs = $row.Cells.Count } }
        $half = [Math]::Max(1, [Math]::Floor($maxTcs / 2))
        if ($maxTcs -lt 2) { $half = 1 }
        $sb = New-Object Text.StringBuilder
        foreach ($row in $t.Rows) {
            $idx = 0
            foreach ($cell in $row.Cells) {
                $idx++
                if ($idx -gt $half) { break }
                $txt = $cell.Range.Text -replace "`a", ''  # remove cell marker
                [void]$sb.AppendLine($txt)
            }
        }
        return $sb.ToString()
    }
    # без таблицы: «нелевые» (т.е. не правые) параграфы шапки
    $sb = New-Object Text.StringBuilder
    $limit = 60; $i = 0
    foreach ($p in $doc.Paragraphs) {
        $i++; if ($i -gt $limit) { break }
        if (-not (Is-RightBlockPara $p) -and $p.Range.Text.Trim().Length -gt 0) {
            [void]$sb.AppendLine($p.Range.Text)
        }
    }
    return $sb.ToString()
}

function Check-Blank($doc, $issues) {
    $path = Get-BlankTemplatePath
    if (-not $path) { return }
    $items = Load-BlankTemplate $path
    if ($items.Count -eq 0) { return }
    $blankText = Get-BlankText $doc
    foreach ($it in $items) {
        $found = $false
        if ($it.IsRegex) {
            try { $found = [bool]([regex]::IsMatch($blankText, $it.Pattern)) }
            catch {
                $issues.Add(@{ Code='BLANK_BAD_REGEX'; Fixable=$false; Text="Некорректный regexp в blank_template.txt: «$($it.Pattern)»." })
                continue
            }
        } else {
            $found = $blankText.Contains($it.Pattern)
        }
        if (-not $found) {
            $label = if ($it.IsRegex) { "(regex) $($it.Pattern)" } else { $it.Pattern }
            if ($it.Optional) {
                $issues.Add(@{ Code='BLANK_OPTIONAL'; Fixable=$false; Text="В бланке не найдено (необязательно): «$label»." })
            } else {
                $issues.Add(@{ Code='BLANK_MISSING'; Fixable=$false; Text="В бланке не найдено: «$label»." })
            }
        }
    }
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
