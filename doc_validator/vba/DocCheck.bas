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
    GatherRestrictedMark doc, issues, cnt

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
    GatherRestrictedMark doc, issues2, cnt2

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
    If Not m.LeftOk Then AddIssue issues, cnt, "Левое поле " & Format$(L, "0.0") & " мм (нужно 30 мм).", True, "MARGIN_L"
    If Not m.TopOk Then AddIssue issues, cnt, "Верхнее поле " & Format$(T, "0.0") & " мм (нужно 20 мм).", True, "MARGIN_T"
    If Not m.BottomOk Then AddIssue issues, cnt, "Нижнее поле " & Format$(B, "0.0") & " мм (нужно 20 мм).", True, "MARGIN_B"
    If Not m.RightOk Then AddIssue issues, cnt, "Правое поле " & Format$(R, "0.0") & " мм (нужно 10 мм).", True, "MARGIN_R"
End Sub

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

Private Sub GatherRestrictedMark(doc As Document, issues() As Issue, ByRef cnt As Long)
    Dim t As String
    t = LCase$(doc.Range.Text)
    If InStr(t, "для служебного пользования") > 0 Then
        If InStr(t, "экз.") = 0 And InStr(t, "экз №") = 0 Then
            AddIssue issues, cnt, "«Для служебного пользования» без номера экземпляра.", False, "DSP"
        End If
    End If
End Sub

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
