Attribute VB_Name = "DocCheck"
' =============================================================
'  DocCheck - проверка и автоисправление оформления документа
'             по Памятке.
'  Запуск:  Alt+F8 -> DocCheck.CheckDocument
' =============================================================
Option Explicit

' ----- параметры из Памятки -----
Private Const MM_TOLERANCE As Double = 0.5
Private Const CM_TOLERANCE As Double = 0.05
Private Const FONT_NAME_REQ As String = "Times New Roman"
Private Const INDENT_REQ_CM As Double = 1.25

Private Const MARGIN_LEFT_MM As Double = 30
Private Const MARGIN_TOP_MM As Double = 20
Private Const MARGIN_BOTTOM_MM As Double = 20
Private Const MARGIN_RIGHT_MM As Double = 10

' Целевые значения при автоисправлении
Private Const FIX_FONT_SIZE As Single = 14         ' если кегль не 13/14 — ставим 14
Private Const FIX_LINE_SPACING As Single = 1.5     ' множитель для wdLineSpaceMultiple

' Путь к файлу с образцом бланка (левая часть шапки).
' ОТРЕДАКТИРУЙТЕ под себя или положите blank_template.txt по этому пути.
' Если файл не найден — проверка бланка пропускается.
Private Const BLANK_TEMPLATE_PATH As String = "C:\Users\Public\Documents\blank_template.txt"

' Структура замечания: текст + признак автоисправляемого + код
Private Type Issue
    Text As String
    Fixable As Boolean
    Code As String
End Type

Public Sub CheckDocument()
    If Documents.Count = 0 Then
        MsgBox "Нет открытого документа.", vbExclamation, "DocCheck"
        Exit Sub
    End If

    Dim doc As Document
    Set doc = ActiveDocument

    Dim issues() As Issue
    Dim cnt As Long
    cnt = 0
    ReDim issues(1 To 100)

    Dim m As MarginsState
    GatherMargins doc, issues, cnt, m
    Dim fs As FontState
    GatherFontStyling doc, issues, cnt, fs
    GatherPageNumbers doc, issues, cnt
    GatherExecutor doc, issues, cnt
    GatherHeader doc, issues, cnt

    If cnt = 0 Then
        MsgBox "Документ: " & doc.Name & vbCrLf & _
               String(50, "-") & vbCrLf & _
               "OK. Все автоматические проверки пройдены.", _
               vbInformation, "DocCheck — OK"
        Exit Sub
    End If

    Dim msg As String, fixable As Long, manual As Long
    msg = "Документ: " & doc.Name & vbCrLf & String(50, "-") & vbCrLf
    Dim i As Long
    For i = 1 To cnt
        Dim mark As String
        If issues(i).Fixable Then
            mark = "  [можно исправить] "
            fixable = fixable + 1
        Else
            mark = "  [ручная правка]   "
            manual = manual + 1
        End If
        msg = msg & i & "." & mark & issues(i).Text & vbCrLf
    Next i

    msg = msg & vbCrLf & _
          "Автоматически исправляются: " & fixable & vbCrLf & _
          "Требуют ручной правки: " & manual & vbCrLf & vbCrLf

    If fixable = 0 Then
        msg = msg & "Автоисправление не применимо — править нужно вручную."
        MsgBox msg, vbExclamation, "DocCheck — замечания"
        Exit Sub
    End If

    msg = msg & "Исправить автоматически устранимые замечания?" & vbCrLf & _
                "(изменения можно откатить Ctrl+Z; документ не сохраняется)"
    Dim answer As VbMsgBoxResult
    answer = MsgBox(msg, vbYesNo + vbQuestion, "DocCheck — исправить?")
    If answer <> vbYes Then Exit Sub

    ApplyFixes doc, issues, cnt, m, fs

    ' Перепроверим
    Dim issues2() As Issue, cnt2 As Long
    ReDim issues2(1 To 100)
    cnt2 = 0
    Dim m2 As MarginsState, fs2 As FontState
    GatherMargins doc, issues2, cnt2, m2
    GatherFontStyling doc, issues2, cnt2, fs2
    GatherPageNumbers doc, issues2, cnt2
    GatherExecutor doc, issues2, cnt2
    GatherHeader doc, issues2, cnt2

    Dim res As String
    res = "Исправление выполнено." & vbCrLf & String(50, "-") & vbCrLf
    If cnt2 = 0 Then
        res = res & "Все замечания устранены."
    Else
        res = res & "Осталось замечаний: " & cnt2 & vbCrLf
        For i = 1 To cnt2
            res = res & "  - " & issues2(i).Text & vbCrLf
        Next i
        res = res & vbCrLf & "Эти пункты нужно поправить руками."
    End If
    res = res & vbCrLf & vbCrLf & "Сохраните документ (Ctrl+S), если результат устраивает."
    MsgBox res, vbInformation, "DocCheck — результат"
End Sub

' --------------------------------------------------------------
'  Сбор замечаний + сохранение исходного состояния
' --------------------------------------------------------------
Private Type MarginsState
    LeftOk As Boolean
    TopOk As Boolean
    BottomOk As Boolean
    RightOk As Boolean
End Type

Private Type FontState
    HasBadFont As Boolean
    HasBadSize As Boolean
    HasBadSpacing As Boolean
    HasBadIndent As Boolean
End Type

Private Sub GatherMargins(doc As Document, issues() As Issue, ByRef cnt As Long, ByRef m As MarginsState)
    Dim ps As PageSetup
    Set ps = doc.PageSetup
    Dim L As Double, T As Double, B As Double, R As Double
    L = PointsToMillimeters(ps.LeftMargin)
    T = PointsToMillimeters(ps.TopMargin)
    B = PointsToMillimeters(ps.BottomMargin)
    R = PointsToMillimeters(ps.RightMargin)
    m.LeftOk = (Abs(L - MARGIN_LEFT_MM) <= MM_TOLERANCE)
    m.TopOk = (Abs(T - MARGIN_TOP_MM) <= MM_TOLERANCE)
    m.BottomOk = (Abs(B - MARGIN_BOTTOM_MM) <= MM_TOLERANCE)
    m.RightOk = (Abs(R - MARGIN_RIGHT_MM) <= MM_TOLERANCE)

    Dim hasLetterhead As Boolean
    hasLetterhead = HasLetterhead(doc)

    If Not m.LeftOk Then AddIssue issues, cnt, "Левое поле " & Format$(L, "0.0") & " мм (нужно 30 мм).", True, "MARGIN_L"
    If Not m.TopOk And Not hasLetterhead Then AddIssue issues, cnt, "Верхнее поле " & Format$(T, "0.0") & " мм (нужно 20 мм).", True, "MARGIN_T"
    If Not m.BottomOk Then AddIssue issues, cnt, "Нижнее поле " & Format$(B, "0.0") & " мм (нужно 20 мм).", True, "MARGIN_B"
    If Not m.RightOk Then AddIssue issues, cnt, "Правое поле " & Format$(R, "0.0") & " мм (нужно 10 мм).", True, "MARGIN_R"
End Sub

' Документ начинается с таблицы (бланк с гербом)? Тогда верхнее поле
' интенционально уменьшено и проверяться не должно.
Private Function HasLetterhead(doc As Document) As Boolean
    HasLetterhead = False
    If doc.Tables.Count = 0 Then Exit Function
    Dim firstTable As Table
    Set firstTable = doc.Tables(1)
    ' Найти первый непустой параграф документа и сравнить с первым параграфом таблицы.
    Dim firstNonEmpty As Paragraph
    Dim p As Paragraph
    For Each p In doc.Paragraphs
        If Len(Trim$(p.Range.Text)) > 0 Then
            Set firstNonEmpty = p
            Exit For
        End If
    Next p
    If firstNonEmpty Is Nothing Then Exit Function
    ' Если первый непустой параграф находится внутри первой таблицы — это бланк
    On Error Resume Next
    If firstNonEmpty.Range.Information(wdWithInTable) Then
        Dim tbl As Table
        Set tbl = firstNonEmpty.Range.Tables(1)
        If Not tbl Is Nothing Then
            If tbl.Range.Start = firstTable.Range.Start Then HasLetterhead = True
        End If
    End If
    On Error GoTo 0
End Function

Private Sub GatherFontStyling(doc As Document, issues() As Issue, ByRef cnt As Long, ByRef fs As FontState)
    Dim para As Paragraph
    Dim badFonts As Object, badSizes As Object
    Set badFonts = CreateObject("Scripting.Dictionary")
    Set badSizes = CreateObject("Scripting.Dictionary")
    Dim badSpacing As Boolean
    Dim bodyParas As Long, badIndent As Long

    For Each para In doc.Paragraphs
        Dim txt As String
        txt = Trim$(para.Range.Text)
        If Len(txt) > 1 Then
            Dim f As Font
            Set f = para.Range.Font

            If Len(f.Name) > 0 And StrComp(f.Name, FONT_NAME_REQ, vbTextCompare) <> 0 Then
                If Not badFonts.Exists(f.Name) Then badFonts.Add f.Name, 1
            End If
            Dim sz As Variant
            sz = f.Size
            If IsNumeric(sz) Then
                If sz <> 13 And sz <> 14 Then
                    If Not badSizes.Exists(CStr(sz)) Then badSizes.Add CStr(sz), 1
                End If
            End If

            Dim lsr As Long, ls As Single, okSpacing As Boolean
            lsr = para.Format.LineSpacingRule
            ls = para.Format.LineSpacing
            okSpacing = False
            Select Case lsr
                Case wdLineSpaceSingle, wdLineSpace1pt5
                    okSpacing = True
                Case wdLineSpaceMultiple
                    If f.Size > 0 Then
                        Dim ratio As Double
                        ratio = ls / f.Size
                        If ratio >= 0.95 And ratio <= 1.55 Then okSpacing = True
                    End If
            End Select
            If Not okSpacing Then badSpacing = True

            If Len(txt) > 80 And InStr(LCase$(para.Style.NameLocal), "аголовок") = 0 Then
                bodyParas = bodyParas + 1
                Dim indCm As Double
                indCm = PointsToCentimeters(para.Format.FirstLineIndent)
                If Abs(indCm - INDENT_REQ_CM) > CM_TOLERANCE Then badIndent = badIndent + 1
            End If
        End If
    Next para

    fs.HasBadFont = (badFonts.Count > 0)
    fs.HasBadSize = (badSizes.Count > 0)
    fs.HasBadSpacing = badSpacing
    fs.HasBadIndent = (bodyParas > 0 And badIndent > 0)

    If fs.HasBadFont Then AddIssue issues, cnt, "Шрифт не " & FONT_NAME_REQ & ": " & JoinKeys(badFonts) & ".", True, "FONT"
    If fs.HasBadSize Then AddIssue issues, cnt, "Размер шрифта вне 13–14 пт: " & JoinKeys(badSizes) & ".", True, "SIZE"
    If fs.HasBadSpacing Then AddIssue issues, cnt, "Межстрочный интервал вне 1.0–1.5.", True, "LINESP"
    If fs.HasBadIndent Then AddIssue issues, cnt, "Абзацный отступ ≠ 1.25 см: " & badIndent & " из " & bodyParas & " абзацев.", True, "INDENT"
End Sub

Private Sub GatherPageNumbers(doc As Document, issues() As Issue, ByRef cnt As Long)
    If doc.ComputeStatistics(wdStatisticPages) < 2 Then Exit Sub
    Dim sec As Section, hdr As HeaderFooter
    Dim hasNumber As Boolean, alignedCenter As Boolean, hasDecor As Boolean
    For Each sec In doc.Sections
        Set hdr = sec.Headers(wdHeaderFooterPrimary)
        Dim fld As Field
        For Each fld In hdr.Range.Fields
            If fld.Type = wdFieldPage Then
                hasNumber = True
                If hdr.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter Then alignedCenter = True
                Dim t As String
                t = hdr.Range.Text
                If InStr(t, "-") > 0 Or InStr(LCase$(t), "стр") > 0 Or InStr(LCase$(t), "с.") > 0 Then hasDecor = True
            End If
        Next fld
    Next sec
    If Not hasNumber Then
        AddIssue issues, cnt, "Документ многостраничный, но нет номеров страниц.", True, "PAGE_ADD"
    Else
        If Not alignedCenter Then AddIssue issues, cnt, "Номер страницы не по центру верхнего поля.", True, "PAGE_ALIGN"
        If hasDecor Then AddIssue issues, cnt, "В номере страницы лишние символы (-, стр., с.).", True, "PAGE_DECOR"
    End If
End Sub

Private Sub GatherExecutor(doc As Document, issues() As Issue, ByRef cnt As Long)
    Dim rx As Object
    Set rx = CreateObject("VBScript.RegExp")
    rx.Pattern = "8\(\d{3,5}\)\s?\d"
    If Not rx.Test(doc.Range.Text) Then
        AddIssue issues, cnt, "Не найдена отметка об исполнителе (телефон 8(код)номер).", False, "EXEC"
    End If
End Sub

Private Sub GatherHeader(doc As Document, issues() As Issue, ByRef cnt As Long)
    ' Собираем «шапочные» параграфы: первые 60 параграфов документа.
    ' В Word doc.Paragraphs включает параграфы в ячейках таблиц, так что
    ' табличная шапка тоже попадёт сюда.
    Const LIMIT As Long = 60
    Dim topParas As Collection
    Set topParas = New Collection
    Dim i As Long, para As Paragraph
    i = 0
    For Each para In doc.Paragraphs
        i = i + 1
        If i > LIMIT Then Exit For
        topParas.Add para
    Next para
    If topParas.Count = 0 Then Exit Sub

    ' Правый блок: alignment=Right или левый отступ > 7 см, или ячейка таблицы — последний столбец
    Dim rightText As String, topText As String
    Dim rightParas As Collection
    Set rightParas = New Collection
    Dim p As Paragraph
    For Each p In topParas
        topText = topText & p.Range.Text & vbLf
        If IsRightBlockPara(p) Then
            If Len(Trim$(p.Range.Text)) > 0 Then
                rightParas.Add p
                rightText = rightText & p.Range.Text & vbLf
            End If
        End If
    Next p

    Dim rx As Object
    Set rx = CreateObject("VBScript.RegExp")
    rx.IgnoreCase = True
    rx.Global = False

    ' --- Гриф ДСП ---
    rx.Pattern = "для служебного пользования"
    If rx.Test(topText) Then
        If Not rx.Test(rightText) Then
            AddIssue issues, cnt, "«Для служебного пользования» найдено, но не в правом верхнем углу.", False, "DSP_POS"
        End If
        rx.Pattern = "экз\.?\s*№\s*\d"
        If Not rx.Test(topText) Then
            AddIssue issues, cnt, "«Для служебного пользования» без номера экземпляра («Экз. № …»).", False, "DSP"
        End If
    End If

    ' --- Адресат ---
    If rightParas.Count = 0 Then
        AddIssue issues, cnt, "Не найден блок адресата в правой верхней части документа.", False, "ADRESAT_MISSING"
    Else
        Dim adresatParas As Collection
        Set adresatParas = New Collection
        rx.Pattern = "для служебного пользования|экз\.?\s*№"
        For Each p In rightParas
            If Not rx.Test(p.Range.Text) Then adresatParas.Add p
        Next p

        rx.IgnoreCase = False
        rx.Pattern = "[А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.\s?[А-ЯЁ]\."
        Dim hasFio As Boolean: hasFio = rx.Test(rightText)
        rx.Pattern = "[А-ЯЁ][а-яё]{2,}\s+[А-ЯЁ][а-яё]{2,}\s+[А-ЯЁ][а-яё]{2,}"
        If Not hasFio Then hasFio = rx.Test(rightText)
        rx.Pattern = "ООО|АО|ПАО|ОАО|Министерств|Управлени|Федеральн|Правительств|Администраци|Канцеляри|ФНС|УФНС|ИФНС|Межрегиональн|Инспекци|Департамент|Комитет|Служб"
        Dim hasOrg As Boolean: hasOrg = rx.Test(rightText)

        If adresatParas.Count > 0 And Not (hasFio Or hasOrg) Then
            AddIssue issues, cnt, "Блок справа есть, но в нём не найдены ни ФИО, ни наименование организации.", False, "ADRESAT_MISSING"
        End If

        ' Выравнивание / отступы внутри блока
        If adresatParas.Count >= 2 Then
            Dim firstAlign As Long, firstIndent As Double
            Dim ap As Paragraph, mixedAlign As Boolean, mixedIndent As Boolean
            Set ap = adresatParas(1)
            firstAlign = ap.Format.Alignment
            firstIndent = PointsToCentimeters(ap.Format.LeftIndent)
            Dim j As Long
            For j = 2 To adresatParas.Count
                Set ap = adresatParas(j)
                If ap.Format.Alignment <> firstAlign Then mixedAlign = True
                If Abs(PointsToCentimeters(ap.Format.LeftIndent) - firstIndent) > 0.1 Then mixedIndent = True
            Next j
            If mixedAlign Or mixedIndent Then
                AddIssue issues, cnt, "Строки адресата выровнены неодинаково (должны быть по левому краю блока).", False, "ADRESAT_ALIGN"
            End If
        End If

        ' Межстрочный в адресате должен быть одинарный
        Dim badAdrSpacing As Boolean
        For Each p In adresatParas
            Dim lsr As Long, ls As Single
            lsr = p.Format.LineSpacingRule
            If lsr = wdLineSpace1pt5 Or lsr = wdLineSpaceDouble Then badAdrSpacing = True
            If lsr = wdLineSpaceMultiple And p.Range.Font.Size > 0 Then
                If (p.Format.LineSpacing / p.Range.Font.Size) > 1.05 Then badAdrSpacing = True
            End If
        Next p
        If badAdrSpacing Then
            AddIssue issues, cnt, "Межстрочный интервал в адресате не 1.0.", False, "ADRESAT_SPACING"
        End If
    End If

    ' --- Формат инициалов ---
    rx.IgnoreCase = False
    rx.Pattern = "[А-ЯЁ][а-яё]+[А-ЯЁ]\."
    If rx.Test(rightText) Then
        AddIssue issues, cnt, "Инициалы записаны слитно с фамилией (нужно «Фамилия И.О.» через пробел).", False, "INITIALS_BAD"
    Else
        rx.Pattern = "[А-ЯЁ][а-яё]+\s+[А-ЯЁ]\.\s?[А-ЯЁ]\."
        Dim hasGood As Boolean: hasGood = rx.Test(rightText)
        rx.Pattern = "[А-ЯЁ][а-яё]+\s+[А-ЯЁ]\."
        If rx.Test(rightText) And Not hasGood Then
            AddIssue issues, cnt, "У фамилии указан только один инициал (нужно два: «Фамилия И.О.»).", False, "INITIALS_BAD"
        End If
    End If

    ' --- Заголовок «О …»/«Об …» (в пределах всей шапки) ---
    rx.IgnoreCase = False
    rx.Pattern = "^\s*Об?\s+[а-яёА-ЯЁ]"
    Dim titleFound As Boolean
    For Each p In topParas
        Dim t2 As String
        t2 = Trim$(p.Range.Text)
        If Len(t2) > 0 Then
            If rx.Test(t2) Then titleFound = True: Exit For
        End If
    Next p
    If Not titleFound Then
        AddIssue issues, cnt, "Не найден заголовок к тексту «О …»/«Об …» в шапке (п. 4 Памятки).", False, "TITLE_MISSING"
    End If

    ' --- Проверка бланка (левая часть шапки) по образцу ---
    CheckBlank doc, topParas, issues, cnt
End Sub

Private Sub CheckBlank(doc As Document, topParas As Collection, issues() As Issue, ByRef cnt As Long)
    ' Читаем образец из файла. Если файла нет — тихо пропускаем.
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(BLANK_TEMPLATE_PATH) Then Exit Sub

    Dim ts As Object
    On Error GoTo readFail
    Set ts = fso.OpenTextFile(BLANK_TEMPLATE_PATH, 1, False, -1) ' Unicode
    Dim content As String
    content = ts.ReadAll
    ts.Close
    On Error GoTo 0

    ' Разбираем строки шаблона
    Dim items() As String, optFlags() As Boolean, isRegex() As Boolean
    ReDim items(1 To 100): ReDim optFlags(1 To 100): ReDim isRegex(1 To 100)
    Dim nItems As Long: nItems = 0
    Dim lines() As String, ln As Variant, raw As String
    lines = Split(content, vbCrLf)
    Dim li As Long
    For li = LBound(lines) To UBound(lines)
        raw = lines(li)
        Dim t As String: t = Trim$(raw)
        If Len(t) = 0 Then GoTo nextLine
        If Left$(t, 1) = "#" Then GoTo nextLine
        Dim optF As Boolean: optF = False
        Dim rgxF As Boolean: rgxF = False
        If Left$(t, 1) = "?" Then optF = True: t = LTrim$(Mid$(t, 2))
        If Left$(t, 3) = "re:" Then rgxF = True: t = Mid$(t, 4)
        If Len(t) = 0 Then GoTo nextLine
        nItems = nItems + 1
        If nItems > UBound(items) Then
            ReDim Preserve items(1 To nItems + 50)
            ReDim Preserve optFlags(1 To nItems + 50)
            ReDim Preserve isRegex(1 To nItems + 50)
        End If
        items(nItems) = t
        optFlags(nItems) = optF
        isRegex(nItems) = rgxF
nextLine:
    Next li

    If nItems = 0 Then Exit Sub

    ' Собираем текст левой части бланка
    Dim blankText As String
    blankText = GetBlankText(doc, topParas)

    Dim rx As Object
    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = False

    Dim i As Long, found As Boolean
    For i = 1 To nItems
        found = False
        If isRegex(i) Then
            rx.Pattern = items(i)
            On Error Resume Next
            found = rx.Test(blankText)
            If Err.Number <> 0 Then
                AddIssue issues, cnt, "Некорректный regexp в blank_template.txt: «" & items(i) & "».", False, "BLANK_BAD_REGEX"
                Err.Clear
                On Error GoTo 0
                GoTo nextItem
            End If
            On Error GoTo 0
        Else
            found = (InStr(blankText, items(i)) > 0)
        End If
        If Not found Then
            Dim label As String
            If isRegex(i) Then label = "(regex) " & items(i) Else label = items(i)
            If optFlags(i) Then
                AddIssue issues, cnt, "В бланке не найдено (необязательно): «" & label & "».", False, "BLANK_OPTIONAL"
            Else
                AddIssue issues, cnt, "В бланке не найдено: «" & label & "».", False, "BLANK_MISSING"
            End If
        End If
nextItem:
    Next i
    Exit Sub
readFail:
    AddIssue issues, cnt, "Не удалось прочитать blank_template.txt (" & BLANK_TEMPLATE_PATH & ").", False, "BLANK_READ_FAIL"
End Sub

Private Function GetBlankText(doc As Document, topParas As Collection) As String
    ' Если в шапке есть таблица — берём левую половину её ячеек.
    ' Иначе — собираем «нелевые» параграфы из шапки.
    Dim s As String
    If doc.Tables.Count > 0 Then
        Dim t As Table
        Set t = doc.Tables(1)
        Dim maxTcs As Long: maxTcs = 0
        Dim row As row
        For Each row In t.Rows
            If row.Cells.Count > maxTcs Then maxTcs = row.Cells.Count
        Next row
        Dim half As Long: half = maxTcs \ 2
        If half < 1 Then half = 1
        For Each row In t.Rows
            Dim c As cell
            Dim cIdx As Long: cIdx = 0
            For Each c In row.Cells
                cIdx = cIdx + 1
                If cIdx > half Then Exit For
                Dim ctext As String
                ctext = c.Range.Text
                ' убираем символ конца ячейки (Chr(13)+Chr(7))
                ctext = Replace(ctext, Chr(7), "")
                s = s & ctext & vbLf
            Next c
        Next row
    Else
        Dim p As Paragraph
        For Each p In topParas
            If Not IsRightBlockPara(p) Then
                If Len(Trim$(p.Range.Text)) > 0 Then s = s & p.Range.Text & vbLf
            End If
        Next p
    End If
    GetBlankText = s
End Function

Private Function IsRightBlockPara(p As Paragraph) As Boolean
    IsRightBlockPara = False
    If p.Format.Alignment = wdAlignParagraphRight Then
        IsRightBlockPara = True
        Exit Function
    End If
    If PointsToCentimeters(p.Format.LeftIndent) > 7 Then
        IsRightBlockPara = True
        Exit Function
    End If
    ' Параграф внутри таблицы: правый блок = не первый столбец строки
    On Error Resume Next
    If p.Range.Information(wdWithInTable) Then
        Dim c As cell
        Set c = p.Range.Cells(1)
        If Err.Number = 0 Then
            If c.ColumnIndex >= 2 Then IsRightBlockPara = True
        End If
        Err.Clear
    End If
    On Error GoTo 0
End Function

' --------------------------------------------------------------
'  Применение исправлений
' --------------------------------------------------------------
Private Sub ApplyFixes(doc As Document, issues() As Issue, cnt As Long, m As MarginsState, fs As FontState)
    Dim i As Long
    For i = 1 To cnt
        If Not issues(i).Fixable Then GoTo nxt
        Select Case issues(i).Code
            Case "MARGIN_L": doc.PageSetup.LeftMargin = MillimetersToPoints(MARGIN_LEFT_MM)
            Case "MARGIN_T": doc.PageSetup.TopMargin = MillimetersToPoints(MARGIN_TOP_MM)
            Case "MARGIN_B": doc.PageSetup.BottomMargin = MillimetersToPoints(MARGIN_BOTTOM_MM)
            Case "MARGIN_R": doc.PageSetup.RightMargin = MillimetersToPoints(MARGIN_RIGHT_MM)
            Case "FONT": FixFontName doc
            Case "SIZE": FixFontSize doc
            Case "LINESP": FixLineSpacing doc
            Case "INDENT": FixIndent doc
            Case "PAGE_ADD": AddPageNumber doc
            Case "PAGE_ALIGN": CenterPageNumber doc
            Case "PAGE_DECOR": CleanPageNumber doc
        End Select
nxt:
    Next i
End Sub

Private Sub FixFontName(doc As Document)
    Dim para As Paragraph
    For Each para In doc.Paragraphs
        If Len(Trim$(para.Range.Text)) > 1 Then
            If StrComp(para.Range.Font.Name, FONT_NAME_REQ, vbTextCompare) <> 0 Then
                para.Range.Font.Name = FONT_NAME_REQ
            End If
        End If
    Next para
End Sub

Private Sub FixFontSize(doc As Document)
    Dim para As Paragraph
    For Each para In doc.Paragraphs
        If Len(Trim$(para.Range.Text)) > 1 Then
            Dim sz As Variant
            sz = para.Range.Font.Size
            If IsNumeric(sz) Then
                If sz <> 13 And sz <> 14 Then para.Range.Font.Size = FIX_FONT_SIZE
            End If
        End If
    Next para
End Sub

Private Sub FixLineSpacing(doc As Document)
    Dim para As Paragraph
    For Each para In doc.Paragraphs
        If Len(Trim$(para.Range.Text)) > 1 Then
            para.Format.LineSpacingRule = wdLineSpaceMultiple
            para.Format.LineSpacing = para.Range.Font.Size * FIX_LINE_SPACING
        End If
    Next para
End Sub

Private Sub FixIndent(doc As Document)
    Dim para As Paragraph
    For Each para In doc.Paragraphs
        Dim txt As String
        txt = Trim$(para.Range.Text)
        If Len(txt) > 80 And InStr(LCase$(para.Style.NameLocal), "аголовок") = 0 Then
            para.Format.FirstLineIndent = CentimetersToPoints(INDENT_REQ_CM)
        End If
    Next para
End Sub

Private Sub AddPageNumber(doc As Document)
    Dim sec As Section, hdr As HeaderFooter
    For Each sec In doc.Sections
        Set hdr = sec.Headers(wdHeaderFooterPrimary)
        hdr.Range.Text = ""
        hdr.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        hdr.Range.Fields.Add Range:=hdr.Range, Type:=wdFieldPage
    Next sec
End Sub

Private Sub CenterPageNumber(doc As Document)
    Dim sec As Section
    For Each sec In doc.Sections
        sec.Headers(wdHeaderFooterPrimary).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    Next sec
End Sub

Private Sub CleanPageNumber(doc As Document)
    ' Перезаписываем колонтитул чистым полем PAGE
    AddPageNumber doc
End Sub

' --------------------------------------------------------------
'  Утилиты
' --------------------------------------------------------------
Private Sub AddIssue(arr() As Issue, ByRef cnt As Long, text As String, fixable As Boolean, code As String)
    cnt = cnt + 1
    If cnt > UBound(arr) Then ReDim Preserve arr(1 To cnt + 50)
    arr(cnt).Text = text
    arr(cnt).Fixable = fixable
    arr(cnt).Code = code
End Sub

Private Function PointsToMillimeters(pts As Single) As Double
    PointsToMillimeters = pts * 25.4 / 72
End Function

Private Function PointsToCentimeters(pts As Single) As Double
    PointsToCentimeters = pts * 2.54 / 72
End Function

Private Function JoinKeys(d As Object) As String
    Dim k As Variant, s As String
    For Each k In d.Keys
        If Len(s) > 0 Then s = s & ", "
        s = s & k
    Next k
    JoinKeys = s
End Function
