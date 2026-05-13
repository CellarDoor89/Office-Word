Attribute VB_Name = "DocCheck"
' =============================================================
'  DocCheck - проверка оформления документа по Памятке
'  Запуск: Alt+F8 -> DocCheck.CheckDocument
'  Или назначить кнопку на панели быстрого доступа.
' =============================================================
Option Explicit

' ----- допуски -----
Private Const MM_TOLERANCE As Double = 0.5       ' мм
Private Const CM_TOLERANCE As Double = 0.05      ' см
Private Const FONT_NAME_REQ As String = "Times New Roman"
Private Const INDENT_REQ_CM As Double = 1.25

' Поля по Памятке: левое 30, верхнее 20, нижнее 20, правое 10 мм
Private Const MARGIN_LEFT_MM As Double = 30
Private Const MARGIN_TOP_MM As Double = 20
Private Const MARGIN_BOTTOM_MM As Double = 20
Private Const MARGIN_RIGHT_MM As Double = 10

Public Sub CheckDocument()
    If Documents.Count = 0 Then
        MsgBox "Нет открытого документа.", vbExclamation, "DocCheck"
        Exit Sub
    End If

    Dim doc As Document
    Set doc = ActiveDocument

    Dim issues As Collection
    Set issues = New Collection

    CheckMargins doc, issues
    CheckFontAndSpacing doc, issues
    CheckIndent doc, issues
    CheckPageNumbers doc, issues
    CheckExecutor doc, issues
    CheckAttachmentMark doc, issues
    CheckRestrictedMark doc, issues

    ShowReport doc, issues
End Sub

' --------------------------------------------------------------
'  Проверки
' --------------------------------------------------------------
Private Sub CheckMargins(doc As Document, issues As Collection)
    Dim ps As PageSetup
    Set ps = doc.PageSetup

    Dim leftMM As Double, topMM As Double, bottomMM As Double, rightMM As Double
    leftMM = PointsToMillimeters(ps.LeftMargin)
    topMM = PointsToMillimeters(ps.TopMargin)
    bottomMM = PointsToMillimeters(ps.BottomMargin)
    rightMM = PointsToMillimeters(ps.RightMargin)

    If Abs(leftMM - MARGIN_LEFT_MM) > MM_TOLERANCE Then _
        issues.Add "Левое поле " & FormatMM(leftMM) & " мм (требуется 30 мм)."
    If Abs(topMM - MARGIN_TOP_MM) > MM_TOLERANCE Then _
        issues.Add "Верхнее поле " & FormatMM(topMM) & " мм (требуется 20 мм)."
    If Abs(bottomMM - MARGIN_BOTTOM_MM) > MM_TOLERANCE Then _
        issues.Add "Нижнее поле " & FormatMM(bottomMM) & " мм (требуется 20 мм)."
    If Abs(rightMM - MARGIN_RIGHT_MM) > MM_TOLERANCE Then _
        issues.Add "Правое поле " & FormatMM(rightMM) & " мм (требуется 10 мм)."
End Sub

Private Sub CheckFontAndSpacing(doc As Document, issues As Collection)
    Dim para As Paragraph
    Dim badFonts As Object, badSizes As Object, badSpacing As Object
    Set badFonts = CreateObject("Scripting.Dictionary")
    Set badSizes = CreateObject("Scripting.Dictionary")
    Set badSpacing = CreateObject("Scripting.Dictionary")

    Dim totalChecked As Long
    totalChecked = 0

    For Each para In doc.Paragraphs
        Dim txt As String
        txt = Trim$(para.Range.Text)
        If Len(txt) > 1 Then
            totalChecked = totalChecked + 1
            Dim f As Font
            Set f = para.Range.Font

            If Len(f.Name) > 0 Then
                If StrComp(f.Name, FONT_NAME_REQ, vbTextCompare) <> 0 Then
                    If Not badFonts.Exists(f.Name) Then badFonts.Add f.Name, 1
                End If
            End If

            Dim sz As Variant
            sz = f.Size
            If IsNumeric(sz) Then
                If sz <> 13 And sz <> 14 Then
                    If Not badSizes.Exists(CStr(sz)) Then badSizes.Add CStr(sz), 1
                End If
            End If

            Dim ls As Single, lsr As Long
            lsr = para.Format.LineSpacingRule
            ls = para.Format.LineSpacing
            ' Допустимо: одинарный, 1.5, или Multiple в [1.0, 1.5]
            Dim okSpacing As Boolean
            okSpacing = False
            Select Case lsr
                Case wdLineSpaceSingle, wdLineSpace1pt5
                    okSpacing = True
                Case wdLineSpaceMultiple
                    ' LineSpacing в Multiple хранится как кегль*интервал/12; чаще как 12, 15, 18
                    ' Нормализуем относительно текущего размера шрифта
                    Dim ratio As Double
                    If f.Size > 0 Then
                        ratio = ls / f.Size
                    Else
                        ratio = 1
                    End If
                    If ratio >= 0.95 And ratio <= 1.55 Then okSpacing = True
                Case Else
                    okSpacing = False
            End Select
            If Not okSpacing Then
                Dim key As String
                key = "правило=" & lsr & "; значение=" & ls
                If Not badSpacing.Exists(key) Then badSpacing.Add key, 1
            End If
        End If
    Next para

    If badFonts.Count > 0 Then
        issues.Add "Используется не " & FONT_NAME_REQ & ": " & JoinKeys(badFonts) & "."
    End If
    If badSizes.Count > 0 Then
        issues.Add "Размер шрифта вне 13–14 пт: " & JoinKeys(badSizes) & "."
    End If
    If badSpacing.Count > 0 Then
        issues.Add "Межстрочный интервал вне диапазона 1.0–1.5."
    End If
End Sub

Private Sub CheckIndent(doc As Document, issues As Collection)
    Dim para As Paragraph
    Dim bodyParas As Long, badIndent As Long
    bodyParas = 0
    badIndent = 0

    For Each para In doc.Paragraphs
        Dim txt As String
        txt = Trim$(para.Range.Text)
        ' Считаем «телом» абзацы длиной > 80 символов и не заголовки
        If Len(txt) > 80 And InStr(LCase$(para.Style.NameLocal), "заголовок") = 0 Then
            bodyParas = bodyParas + 1
            Dim indentCm As Double
            indentCm = PointsToCentimeters(para.Format.FirstLineIndent)
            If Abs(indentCm - INDENT_REQ_CM) > CM_TOLERANCE Then badIndent = badIndent + 1
        End If
    Next para

    If bodyParas > 0 And badIndent > 0 Then
        issues.Add "Абзацный отступ ≠ 1.25 см в " & badIndent & " из " & bodyParas & " абзацев основного текста."
    End If
End Sub

Private Sub CheckPageNumbers(doc As Document, issues As Collection)
    If doc.ComputeStatistics(wdStatisticPages) < 2 Then Exit Sub

    Dim sec As Section, hdr As HeaderFooter
    Dim hasNumber As Boolean, hasDecor As Boolean, alignedCenter As Boolean
    hasNumber = False: hasDecor = False: alignedCenter = False

    For Each sec In doc.Sections
        Set hdr = sec.Headers(wdHeaderFooterPrimary)
        Dim numField As Field
        For Each numField In hdr.Range.Fields
            If numField.Type = wdFieldPage Then
                hasNumber = True
                If hdr.Range.ParagraphFormat.Alignment = wdAlignParagraphCenter Then alignedCenter = True
                Dim hdrText As String
                hdrText = hdr.Range.Text
                If InStr(hdrText, "-") > 0 Or InStr(LCase$(hdrText), "стр") > 0 Or InStr(LCase$(hdrText), "с.") > 0 Then
                    hasDecor = True
                End If
            End If
        Next numField
    Next sec

    If Not hasNumber Then
        issues.Add "Документ многостраничный, но номера страниц не найдены (требуется сверху по центру, арабские)."
    Else
        If Not alignedCenter Then issues.Add "Номер страницы должен быть выровнен по центру верхнего поля."
        If hasDecor Then issues.Add "В номере страницы есть лишние символы (тире / «стр.» / «с.») — должно быть только число."
    End If
End Sub

Private Sub CheckExecutor(doc As Document, issues As Collection)
    Dim txt As String
    txt = doc.Range.Text
    ' Грубая проверка: телефон вида 8(....)... в нижней четверти документа
    Dim phoneRegex As Object
    Set phoneRegex = CreateObject("VBScript.RegExp")
    phoneRegex.Pattern = "8\(\d{3,5}\)\s?\d"
    phoneRegex.Global = False
    If Not phoneRegex.Test(txt) Then
        issues.Add "Не найдена отметка об исполнителе (ФИО + телефон в формате 8(код)номер)."
    End If
End Sub

Private Sub CheckAttachmentMark(doc As Document, issues As Collection)
    ' Информационная: если в тексте упоминается «приложение», проверяем формат
    Dim txt As String
    txt = doc.Range.Text
    If InStr(LCase$(txt), "приложение") > 0 Then
        Dim rx As Object
        Set rx = CreateObject("VBScript.RegExp")
        rx.Pattern = "Приложение(:|\s)"
        rx.IgnoreCase = False
        If Not rx.Test(txt) Then
            issues.Add "В тексте упомянуто приложение, но отметка «Приложение:» в требуемом формате не найдена."
        End If
    End If
End Sub

Private Sub CheckRestrictedMark(doc As Document, issues As Collection)
    ' Если в тексте есть «для служебного пользования», должен быть номер экземпляра
    Dim txt As String
    txt = LCase$(doc.Range.Text)
    If InStr(txt, "для служебного пользования") > 0 Then
        If InStr(txt, "экз.") = 0 And InStr(txt, "экз №") = 0 Then
            issues.Add "Указано «Для служебного пользования», но не найден номер экземпляра («Экз. № …»)."
        End If
    End If
End Sub

' --------------------------------------------------------------
'  Отчёт
' --------------------------------------------------------------
Private Sub ShowReport(doc As Document, issues As Collection)
    Dim msg As String
    msg = "Документ: " & doc.Name & vbCrLf & String(50, "-") & vbCrLf
    If issues.Count = 0 Then
        msg = msg & "OK. Все автоматические проверки пройдены." & vbCrLf & vbCrLf & _
              "Помните: проверки не заменяют визуальный осмотр перед отправкой на подпись."
        MsgBox msg, vbInformation, "DocCheck — OK"
    Else
        Dim i As Long
        For i = 1 To issues.Count
            msg = msg & i & ". " & issues(i) & vbCrLf
        Next i
        msg = msg & vbCrLf & "Найдено замечаний: " & issues.Count
        MsgBox msg, vbExclamation, "DocCheck — замечания"
    End If
End Sub

' --------------------------------------------------------------
'  Утилиты
' --------------------------------------------------------------
Private Function PointsToMillimeters(pts As Single) As Double
    PointsToMillimeters = pts * 25.4 / 72
End Function

Private Function PointsToCentimeters(pts As Single) As Double
    PointsToCentimeters = pts * 2.54 / 72
End Function

Private Function FormatMM(v As Double) As String
    FormatMM = Format$(v, "0.0")
End Function

Private Function JoinKeys(d As Object) As String
    Dim k As Variant, s As String
    For Each k In d.Keys
        If Len(s) > 0 Then s = s & ", "
        s = s & k
    Next k
    JoinKeys = s
End Function
