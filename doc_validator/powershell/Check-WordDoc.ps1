<#
.SYNOPSIS
    Проверка оформления Word-документа по «Памятке».

.DESCRIPTION
    Использует COM-объект Word.Application (входит в установленный Microsoft Word).
    Ничего ставить не нужно. Запускается двойным кликом из проводника
    (если разрешено выполнение .ps1) либо командой:
        powershell -ExecutionPolicy Bypass -File Check-WordDoc.ps1 -Path "C:\путь\к\файлу.docx"

.PARAMETER Path
    Путь к проверяемому .docx / .doc. Если не указан, появится диалог выбора файла.
#>

[CmdletBinding()]
param(
    [string]$Path
)

# ---------- параметры из Памятки ----------
$MARGIN_LEFT_MM   = 30
$MARGIN_TOP_MM    = 20
$MARGIN_BOTTOM_MM = 20
$MARGIN_RIGHT_MM  = 10
$MM_TOLERANCE     = 0.5
$CM_TOLERANCE     = 0.05
$FONT_REQ         = 'Times New Roman'
$INDENT_REQ_CM    = 1.25

# ---------- утилиты ----------
function PointsToMillimeters([double]$pts) { return [Math]::Round($pts * 25.4 / 72, 2) }
function PointsToCentimeters([double]$pts) { return [Math]::Round($pts * 2.54 / 72, 3) }

function Show-OpenDialog {
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Документы Word (*.docx;*.doc)|*.docx;*.doc|Все файлы|*.*'
    $dlg.Title  = 'Выберите документ для проверки'
    if ($dlg.ShowDialog() -ne 'OK') { return $null }
    return $dlg.FileName
}

# ---------- получение пути ----------
if (-not $Path) { $Path = Show-OpenDialog }
if (-not $Path -or -not (Test-Path -LiteralPath $Path)) {
    Write-Host 'Файл не выбран. Выход.' -ForegroundColor Yellow
    return
}
$Path = (Resolve-Path -LiteralPath $Path).Path

# ---------- запуск Word ----------
$word = $null
$doc  = $null
try {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0  # wdAlertsNone
    $doc = $word.Documents.Open($Path, [ref]$false, [ref]$true)  # ReadOnly

    $issues = New-Object System.Collections.Generic.List[string]

    # --- поля ---
    $ps = $doc.PageSetup
    $leftMM   = PointsToMillimeters $ps.LeftMargin
    $topMM    = PointsToMillimeters $ps.TopMargin
    $bottomMM = PointsToMillimeters $ps.BottomMargin
    $rightMM  = PointsToMillimeters $ps.RightMargin

    if ([Math]::Abs($leftMM   - $MARGIN_LEFT_MM)   -gt $MM_TOLERANCE) { $issues.Add("Левое поле $leftMM мм (требуется 30 мм).") }
    if ([Math]::Abs($topMM    - $MARGIN_TOP_MM)    -gt $MM_TOLERANCE) { $issues.Add("Верхнее поле $topMM мм (требуется 20 мм).") }
    if ([Math]::Abs($bottomMM - $MARGIN_BOTTOM_MM) -gt $MM_TOLERANCE) { $issues.Add("Нижнее поле $bottomMM мм (требуется 20 мм).") }
    if ([Math]::Abs($rightMM  - $MARGIN_RIGHT_MM)  -gt $MM_TOLERANCE) { $issues.Add("Правое поле $rightMM мм (требуется 10 мм).") }

    # --- шрифт, кегль, интервал, отступ ---
    $badFonts    = New-Object System.Collections.Generic.HashSet[string]
    $badSizes    = New-Object System.Collections.Generic.HashSet[string]
    $badSpacing  = $false
    $bodyParas   = 0
    $badIndent   = 0

    foreach ($p in $doc.Paragraphs) {
        $txt = $p.Range.Text.Trim()
        if ($txt.Length -le 1) { continue }

        $f = $p.Range.Font
        if ($f.Name -and $f.Name -ne $FONT_REQ) { [void]$badFonts.Add($f.Name) }

        $sz = $f.Size
        if ($sz -is [double] -and $sz -ne 13 -and $sz -ne 14) { [void]$badSizes.Add("$sz") }

        # межстрочный
        $lsRule = $p.Format.LineSpacingRule
        $ls     = $p.Format.LineSpacing
        # wdLineSpaceSingle = 0, wdLineSpace1pt5 = 1, wdLineSpaceMultiple = 5
        $okSpacing = $false
        switch ($lsRule) {
            0 { $okSpacing = $true }
            1 { $okSpacing = $true }
            5 {
                if ($f.Size -gt 0) {
                    $ratio = $ls / $f.Size
                    if ($ratio -ge 0.95 -and $ratio -le 1.55) { $okSpacing = $true }
                }
            }
        }
        if (-not $okSpacing) { $badSpacing = $true }

        # отступ первой строки (только для длинных абзацев, не заголовков)
        $styleName = "$($p.Style.NameLocal)"
        if ($txt.Length -gt 80 -and $styleName -notmatch 'аголовок') {
            $bodyParas++
            $indCm = PointsToCentimeters $p.Format.FirstLineIndent
            if ([Math]::Abs($indCm - $INDENT_REQ_CM) -gt $CM_TOLERANCE) { $badIndent++ }
        }
    }

    if ($badFonts.Count -gt 0)   { $issues.Add("Используется не $FONT_REQ: $($badFonts -join ', ').") }
    if ($badSizes.Count -gt 0)   { $issues.Add("Размер шрифта вне 13–14 пт: $($badSizes -join ', ').") }
    if ($badSpacing)             { $issues.Add('Межстрочный интервал вне диапазона 1.0–1.5.') }
    if ($bodyParas -gt 0 -and $badIndent -gt 0) {
        $issues.Add("Абзацный отступ ≠ 1.25 см в $badIndent из $bodyParas абзацев основного текста.")
    }

    # --- номера страниц на многостраничных документах ---
    $wdStatisticPages = 2
    $pages = $doc.ComputeStatistics($wdStatisticPages)
    if ($pages -ge 2) {
        $wdHeaderFooterPrimary = 1
        $wdFieldPage           = 33
        $hasNumber = $false
        $alignedCenter = $false
        $hasDecor = $false
        foreach ($sec in $doc.Sections) {
            $hdr = $sec.Headers.Item($wdHeaderFooterPrimary)
            foreach ($fld in $hdr.Range.Fields) {
                if ($fld.Type -eq $wdFieldPage) {
                    $hasNumber = $true
                    if ($hdr.Range.ParagraphFormat.Alignment -eq 1) { $alignedCenter = $true }
                    $hdrText = $hdr.Range.Text
                    if ($hdrText -match '-' -or $hdrText -match 'стр' -or $hdrText -match 'с\.') { $hasDecor = $true }
                }
            }
        }
        if (-not $hasNumber) {
            $issues.Add('Документ многостраничный, но номера страниц не найдены.')
        } else {
            if (-not $alignedCenter) { $issues.Add('Номер страницы должен быть по центру верхнего поля.') }
            if ($hasDecor)           { $issues.Add('В номере страницы есть лишние символы (тире / «стр.» / «с.»).') }
        }
    }

    # --- отметка об исполнителе ---
    $bodyText = $doc.Range().Text
    if ($bodyText -notmatch '8\(\d{3,5}\)\s?\d') {
        $issues.Add('Не найдена отметка об исполнителе (телефон в формате 8(код)номер).')
    }

    # --- гриф ДСП + номер экземпляра ---
    if ($bodyText -match '(?i)для служебного пользования') {
        if ($bodyText -notmatch '(?i)экз\.?\s*№') {
            $issues.Add('Указано «Для служебного пользования», но не найден номер экземпляра («Экз. № …»).')
        }
    }

    # ---------- вывод ----------
    Write-Host ''
    Write-Host "Документ: $Path" -ForegroundColor Cyan
    Write-Host ('-' * 60)
    if ($issues.Count -eq 0) {
        Write-Host 'OK. Все автоматические проверки пройдены.' -ForegroundColor Green
    } else {
        for ($i = 0; $i -lt $issues.Count; $i++) {
            Write-Host ("{0}. {1}" -f ($i + 1), $issues[$i]) -ForegroundColor Yellow
        }
        Write-Host ''
        Write-Host "Найдено замечаний: $($issues.Count)" -ForegroundColor Red
    }
}
finally {
    if ($doc)  { $doc.Close([ref]$false)   | Out-Null; [Runtime.InteropServices.Marshal]::ReleaseComObject($doc)  | Out-Null }
    if ($word) { $word.Quit()              | Out-Null; [Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
