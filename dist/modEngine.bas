Attribute VB_Name = "modEngine"
Option Explicit
'==============================================================
' modEngine - Chuan hoa va Ra soat theo mot TRule.
' Chi dung object model cua Word, chay duoc tren Windows va Mac.
'==============================================================

Private Const MAX_COMMENTS As Long = 40
Private Const UNDEF As Long = 9999999

Private mIssueCount As Long
Private mCommentCount As Long
Private mNotes As String

'--------------------------------------------------------------
' Tien ich
'--------------------------------------------------------------
Private Function Mm(ByVal v As Single) As Single
    Mm = Application.MillimetersToPoints(v)
End Function

Private Function ToMm(ByVal pt As Single) As Single
    ToMm = Application.PointsToMillimeters(pt)
End Function

Private Function InRange(ByVal v As Single, ByVal lo As Single, ByVal hi As Single, _
                         ByVal tol As Single) As Boolean
    InRange = (v >= lo - tol And v <= hi + tol)
End Function

Private Function CleanText(ByVal s As String) As String
    s = Replace(s, vbCr, "")
    s = Replace(s, Chr(7), "")
    s = Replace(s, Chr(11), " ")
    CleanText = Trim$(s)
End Function

Private Function TextRange(ByVal p As Paragraph) As Range
    Dim rg As Range
    Set rg = p.Range
    If rg.End > rg.Start Then rg.MoveEnd wdCharacter, -1
    Set TextRange = rg
End Function

Private Function MatchAnchor(ByVal txt As String, ByRef r As TRule, ByVal i As Integer) As Boolean
    Dim t As String, k As String
    t = CleanText(txt)
    k = r.AnchorKey(i)
    If Len(t) < Len(k) Then Exit Function
    If StrComp(Left$(t, Len(k)), k, vbTextCompare) <> 0 Then Exit Function
    If r.AnchorUpper(i) Then
        If t <> UCase$(t) Then Exit Function
    End If
    MatchAnchor = True
End Function

Private Function FindAnchor(ByRef r As TRule, ByVal i As Integer) As Paragraph
    Dim doc As Document, n As Long, k As Long, lim As Long
    Set doc = ActiveDocument
    n = doc.Paragraphs.Count
    If r.AnchorFromEnd(i) Then
        lim = n - 80
        If lim < 1 Then lim = 1
        For k = n To lim Step -1
            If MatchAnchor(doc.Paragraphs(k).Range.Text, r, i) Then
                Set FindAnchor = doc.Paragraphs(k)
                Exit Function
            End If
        Next k
    Else
        lim = 40
        If lim > n Then lim = n
        For k = 1 To lim
            If MatchAnchor(doc.Paragraphs(k).Range.Text, r, i) Then
                Set FindAnchor = doc.Paragraphs(k)
                Exit Function
            End If
        Next k
    End If
End Function

' Doan noi dung: dai, khong phai tieu de, khong canh giua/phai, khong trong bang
Private Function IsBodyParagraph(ByVal p As Paragraph) As Boolean
    Dim t As String
    If p.Alignment = wdAlignParagraphCenter Or p.Alignment = wdAlignParagraphRight Then Exit Function
    If p.OutlineLevel <> wdOutlineLevelBodyText Then Exit Function
    t = CleanText(p.Range.Text)
    If Len(t) < 40 Then Exit Function
    If p.Range.Information(wdWithInTable) Then Exit Function
    IsBodyParagraph = True
End Function

Private Function LineMultOf(ByVal p As Paragraph) As Single
    Select Case p.Format.LineSpacingRule
        Case wdLineSpaceSingle: LineMultOf = 1
        Case wdLineSpace1pt5: LineMultOf = 1.5
        Case wdLineSpaceDouble: LineMultOf = 2
        Case wdLineSpaceMultiple: LineMultOf = p.Format.LineSpacing / 12
        Case Else: LineMultOf = -1   ' AtLeast / Exactly
    End Select
End Function

'--------------------------------------------------------------
' Nhan xet
'--------------------------------------------------------------
Private Sub AddIssue(ByVal rg As Range, ByVal msg As String)
    mIssueCount = mIssueCount + 1
    If mCommentCount >= MAX_COMMENTS Then Exit Sub
    On Error Resume Next
    Dim c As Comment
    Set c = ActiveDocument.Comments.Add(rg, msg)
    If Not c Is Nothing Then
        c.Author = REVIEW_AUTHOR
        c.Initial = "VF"
        mCommentCount = mCommentCount + 1
    End If
    On Error GoTo 0
End Sub

Private Sub AddNote(ByVal msg As String)
    mIssueCount = mIssueCount + 1
    If Len(mNotes) < 1500 Then mNotes = mNotes & "- " & msg & vbCrLf
End Sub

Public Function ClearReviewComments() As Long
    Dim i As Long, n As Long
    For i = ActiveDocument.Comments.Count To 1 Step -1
        If ActiveDocument.Comments(i).Author = REVIEW_AUTHOR Then
            ActiveDocument.Comments(i).Delete
            n = n + 1
        End If
    Next i
    ClearReviewComments = n
End Function

'--------------------------------------------------------------
' RA SOAT
'--------------------------------------------------------------
Public Sub ReviewDocument(ByVal ruleId As String)
    Dim r As TRule
    r = GetRule(ruleId)
    If Not r.Available Then
        Info U("Quy chu\u1ea9n ") & r.Title & U(" ch\u01b0a c\u00f3 s\u1ed1 li\u1ec7u \u0111\u00e3 x\u00e1c minh n\u00ean ch\u01b0a th\u1ec3 r\u00e0 so\u00e1t.") & _
             vbCrLf & vbCrLf & r.Notes
        Exit Sub
    End If

    mIssueCount = 0: mCommentCount = 0: mNotes = ""
    PerfOn
    On Error GoTo EH
    ClearReviewComments

    CheckPageSetup r
    CheckAnchors r
    CheckBody r
    CheckPageNumber r

    PerfOff
    Dim msg As String
    If mIssueCount = 0 Then
        msg = U("Kh\u00f4ng ph\u00e1t hi\u1ec7n l\u1ed7i th\u1ec3 th\u1ee9c theo ") & r.Title & "."
    Else
        msg = U("Ph\u00e1t hi\u1ec7n ") & mIssueCount & U(" \u0111i\u1ec3m ch\u01b0a \u0111\u00fang theo ") & r.Title & "."
        If mCommentCount > 0 Then msg = msg & vbCrLf & U("\u0110\u00e3 ghi ") & mCommentCount & U(" nh\u1eadn x\u00e9t trong v\u0103n b\u1ea3n (t\u00e1c gi\u1ea3 ViFormat).")
        If mIssueCount > mCommentCount And mCommentCount >= MAX_COMMENTS Then _
            msg = msg & vbCrLf & U("Ch\u1ec9 ghi t\u1ed1i \u0111a ") & MAX_COMMENTS & U(" nh\u1eadn x\u00e9t \u0111\u1ea7u ti\u00ean.")
        If Len(mNotes) > 0 Then msg = msg & vbCrLf & vbCrLf & mNotes
    End If
    If Len(r.Notes) > 0 Then msg = msg & vbCrLf & vbCrLf & U("L\u01b0u \u00fd: ") & r.Notes
    Notify msg
    Exit Sub
EH:
    PerfOff
    Info U("L\u1ed7i khi r\u00e0 so\u00e1t: ") & Err.Description
End Sub

Private Sub CheckPageSetup(ByRef r As TRule)
    Dim s As Section, k As Integer, idx As Integer, v As Single, rg As Range
    Dim names As Variant
    names = Array("", U("tr\u00ean"), U("d\u01b0\u1edbi"), U("tr\u00e1i"), U("ph\u1ea3i"))
    Set rg = ActiveDocument.Paragraphs(1).Range
    For Each s In ActiveDocument.Sections
        With s.PageSetup
            If s.Index = 1 Or .Orientation = wdOrientPortrait Then
                If Abs(.PageWidth - Application.MillimetersToPoints(210)) > 2 Or _
                   Abs(.PageHeight - Application.MillimetersToPoints(297)) > 2 Then
                    If .Orientation = wdOrientPortrait Then _
                        AddNote "Section " & s.Index & U(": kh\u1ed5 gi\u1ea5y kh\u00f4ng ph\u1ea3i A4 (210 x 297 mm).")
                End If
                For idx = 1 To 4
                    Select Case idx
                        Case M_TOP: v = ToMm(.TopMargin)
                        Case M_BOTTOM: v = ToMm(.BottomMargin)
                        Case M_LEFT: v = ToMm(.LeftMargin)
                        Case M_RIGHT: v = ToMm(.RightMargin)
                    End Select
                    If Not InRange(v, r.MarginMin(idx), r.MarginMax(idx), 0.6) Then
                        AddNote "Section " & s.Index & U(": l\u1ec1 ") & names(idx) & " = " & Format$(v, "0.0") & _
                                U(" mm, quy \u0111\u1ecbnh ") & RangeText(r.MarginMin(idx), r.MarginMax(idx)) & " mm."
                    End If
                Next idx
            End If
        End With
    Next s
End Sub

Private Function RangeText(ByVal lo As Single, ByVal hi As Single) As String
    If lo = hi Then
        RangeText = Format$(lo, "0.#")
    Else
        RangeText = Format$(lo, "0.#") & "-" & Format$(hi, "0.#")
    End If
End Function

Private Sub CheckAnchors(ByRef r As TRule)
    Dim i As Integer, p As Paragraph, rg As Range, sz As Single
    Dim sizes(1 To 4) As Single
    For i = 1 To r.AnchorCount
        Set p = FindAnchor(r, i)
        If p Is Nothing Then
            AddNote U("Kh\u00f4ng t\u00ecm th\u1ea5y ") & r.AnchorName(i) & U(" (b\u1ecf qua n\u1ebfu v\u0103n b\u1ea3n kh\u00f4ng c\u00f3 th\u00e0nh ph\u1ea7n n\u00e0y).")
        Else
            Set rg = TextRange(p)
            sz = rg.Font.Size
            sizes(i) = sz
            If sz = UNDEF Then
                AddIssue rg, r.AnchorName(i) & U(": c\u1ee1 ch\u1eef kh\u00f4ng \u0111\u1ed3ng nh\u1ea5t, quy \u0111\u1ecbnh ") & _
                         RangeText(r.AnchorSizeMin(i), r.AnchorSizeMax(i)) & "."
            ElseIf Not InRange(sz, r.AnchorSizeMin(i), r.AnchorSizeMax(i), 0.01) Then
                AddIssue rg, r.AnchorName(i) & U(": c\u1ee1 ch\u1eef ") & sz & U(", quy \u0111\u1ecbnh ") & _
                         RangeText(r.AnchorSizeMin(i), r.AnchorSizeMax(i)) & "."
            End If
            If r.AnchorBold(i) And rg.Font.Bold <> True Then _
                AddIssue rg, r.AnchorName(i) & U(": ph\u1ea3i in \u0111\u1eadm.")
            If r.AnchorItalic(i) And rg.Font.Italic <> True Then _
                AddIssue rg, r.AnchorName(i) & U(": ph\u1ea3i in nghi\u00eang.")
            If r.AnchorUnder(i) And rg.Font.Underline = wdUnderlineNone Then _
                AddIssue rg, r.AnchorName(i) & U(": ph\u1ea3i g\u1ea1ch ch\u00e2n.")
        End If
    Next i

    If r.TieuNguPlus1 And sizes(1) > 0 And sizes(1) <> UNDEF And sizes(2) > 0 And sizes(2) <> UNDEF Then
        If Abs(sizes(2) - (sizes(1) + 1)) > 0.01 Then
            Set p = FindAnchor(r, 2)
            If Not p Is Nothing Then
                AddIssue TextRange(p), U("Ti\u00eau ng\u1eef ph\u1ea3i l\u1edbn h\u01a1n Qu\u1ed1c hi\u1ec7u 1 c\u1ee1 (Qu\u1ed1c hi\u1ec7u ") & sizes(1) & _
                         U(" th\u00ec Ti\u00eau ng\u1eef ") & (sizes(1) + 1) & ")."
            End If
        End If
    End If
End Sub

Private Sub CheckBody(ByRef r As TRule)
    Dim doc As Document, n As Long, k As Long, p As Paragraph, rg As Range
    Dim fn As String, sz As Single, m As Single
    Dim badFont As Long, badSize As Long, badAlign As Long, badIndent As Long
    Dim badLine As Long, badSpace As Long
    Dim firstFont As Range, firstSize As Range, firstAlign As Range
    Dim firstIndent As Range, firstLine As Range, firstSpace As Range

    Set doc = ActiveDocument
    n = doc.Paragraphs.Count
    For k = 1 To n
        If k Mod 50 = 0 Then Status U("ViFormat: \u0111ang r\u00e0 so\u00e1t \u0111o\u1ea1n ") & k & "/" & n
        Set p = doc.Paragraphs(k)
        If IsBodyParagraph(p) Then
            Set rg = TextRange(p)
            fn = rg.Font.Name
            If fn <> r.FontName Then
                badFont = badFont + 1
                If firstFont Is Nothing Then Set firstFont = rg
            End If
            sz = rg.Font.Size
            If sz = UNDEF Or Not InRange(sz, r.BodySizeMin, r.BodySizeMax, 0.01) Then
                badSize = badSize + 1
                If firstSize Is Nothing Then Set firstSize = rg
            End If
            If r.ParaCheck And r.Justify And p.Alignment <> wdAlignParagraphJustify Then
                badAlign = badAlign + 1
                If firstAlign Is Nothing Then Set firstAlign = rg
            End If
            If r.FirstIndentCheck Then
                If Abs(p.Format.FirstLineIndent - Mm(r.FirstIndentMm)) > Mm(1.5) Then
                    badIndent = badIndent + 1
                    If firstIndent Is Nothing Then Set firstIndent = rg
                End If
            End If
            If r.LineCheck Then
                If r.LineExact Then
                    If p.Format.LineSpacingRule <> wdLineSpaceExactly Or _
                       Not InRange(p.Format.LineSpacing, r.LineExactMin, r.LineExactMax, 0.1) Then
                        badLine = badLine + 1
                        If firstLine Is Nothing Then Set firstLine = rg
                    End If
                Else
                    m = LineMultOf(p)
                    If m < 0 Or Not InRange(m, r.LineMultMin, r.LineMultMax, 0.01) Then
                        badLine = badLine + 1
                        If firstLine Is Nothing Then Set firstLine = rg
                    End If
                End If
            End If
            If r.ParaCheck And p.Format.SpaceAfter + p.Format.SpaceBefore < r.SpaceAfterPt - 0.1 Then
                badSpace = badSpace + 1
                If firstSpace Is Nothing Then Set firstSpace = rg
            End If
        End If
    Next k

    If badFont > 0 Then AddIssue firstFont, U("Ph\u00f4ng ch\u1eef n\u1ed9i dung ph\u1ea3i l\u00e0 ") & r.FontName & _
        U(". C\u00f3 ") & badFont & U(" \u0111o\u1ea1n sai (\u0111\u00e2y l\u00e0 \u0111o\u1ea1n \u0111\u1ea7u ti\u00ean).")
    If badSize > 0 Then AddIssue firstSize, U("C\u1ee1 ch\u1eef n\u1ed9i dung ph\u1ea3i l\u00e0 ") & _
        RangeText(r.BodySizeMin, r.BodySizeMax) & U(". C\u00f3 ") & badSize & U(" \u0111o\u1ea1n sai.")
    If badAlign > 0 Then AddIssue firstAlign, U("N\u1ed9i dung ph\u1ea3i c\u0103n \u0111\u1ec1u hai b\u00ean. C\u00f3 ") & badAlign & U(" \u0111o\u1ea1n sai.")
    If badIndent > 0 Then AddIssue firstIndent, U("L\u00f9i \u0111\u1ea7u d\u00f2ng ph\u1ea3i kho\u1ea3ng ") & r.FirstIndentMm & _
        U(" mm. C\u00f3 ") & badIndent & U(" \u0111o\u1ea1n sai.")
    If badLine > 0 Then
        If r.LineExact Then
            AddIssue firstLine, U("Gi\u00e3n d\u00f2ng ph\u1ea3i l\u00e0 Exactly ") & RangeText(r.LineExactMin, r.LineExactMax) & _
                U(" pt. C\u00f3 ") & badLine & U(" \u0111o\u1ea1n sai.")
        Else
            AddIssue firstLine, U("Gi\u00e3n d\u00f2ng ph\u1ea3i trong kho\u1ea3ng ") & RangeText(r.LineMultMin, r.LineMultMax) & _
                U(" d\u00f2ng. C\u00f3 ") & badLine & U(" \u0111o\u1ea1n sai.")
        End If
    End If
    If badSpace > 0 Then AddIssue firstSpace, U("Kho\u1ea3ng c\u00e1ch gi\u1eefa c\u00e1c \u0111o\u1ea1n t\u1ed1i thi\u1ec3u ") & r.SpaceAfterPt & _
        U(" pt. C\u00f3 ") & badSpace & U(" \u0111o\u1ea1n sai.")
End Sub

Private Function HasPageField(ByVal rg As Range) As Boolean
    Dim f As Field
    For Each f In rg.Fields
        If f.Type = wdFieldPage Then HasPageField = True: Exit Function
    Next f
End Function

Private Sub CheckPageNumber(ByRef r As TRule)
    Dim s As Section, hdr As HeaderFooter, rg As Range
    Set s = ActiveDocument.Sections(1)
    Set hdr = s.Headers(wdHeaderFooterPrimary)
    Set rg = ActiveDocument.Paragraphs(1).Range
    If ActiveDocument.ComputeStatistics(wdStatisticPages) <= 1 Then Exit Sub
    If Not HasPageField(hdr.Range) Then
        AddNote U("Ch\u01b0a c\u00f3 s\u1ed1 trang \u1edf ph\u1ea7n l\u1ec1 tr\u00ean (canh gi\u1eefa). D\u00f9ng n\u00fat \u0110\u00e1nh s\u1ed1 trang.")
    Else
        If hdr.Range.ParagraphFormat.Alignment <> wdAlignParagraphCenter Then _
            AddNote U("S\u1ed1 trang ph\u1ea3i canh gi\u1eefa.")
    End If
    If Not s.PageSetup.DifferentFirstPageHeaderFooter Then _
        AddNote U("Trang \u0111\u1ea7u kh\u00f4ng \u0111\u01b0\u1ee3c hi\u1ec3n th\u1ecb s\u1ed1 trang (b\u1eadt Different First Page).")
End Sub

'--------------------------------------------------------------
' CHUAN HOA
'--------------------------------------------------------------
Public Sub NormalizeDocument(ByVal ruleId As String)
    Dim r As TRule
    r = GetRule(ruleId)
    If Not r.Available Then
        Info U("Quy chu\u1ea9n ") & r.Title & U(" ch\u01b0a c\u00f3 s\u1ed1 li\u1ec7u \u0111\u00e3 x\u00e1c minh n\u00ean ch\u01b0a th\u1ec3 chu\u1ea9n h\u00f3a.") & _
             vbCrLf & vbCrLf & r.Notes
        Exit Sub
    End If

    BeginUndo U("ViFormat: Chu\u1ea9n h\u00f3a")
    PerfOn
    On Error GoTo EH

    ApplyPageSetup r
    ApplyNormalStyle r
    ActiveDocument.Content.Font.Name = r.FontName
    ActiveDocument.Content.Font.Color = wdColorAutomatic
    ApplyBody r
    ApplyAnchors r
    ApplyPageNumber r

    PerfOff
    EndUndo
    Notify U("\u0110\u00e3 chu\u1ea9n h\u00f3a theo ") & r.Title & U(". C\u00f3 th\u1ec3 ho\u00e0n t\u00e1c b\u1eb1ng Undo; n\u00ean ch\u1ea1y R\u00e0 so\u00e1t th\u1ec3 th\u1ee9c \u0111\u1ec3 ki\u1ec3m tra l\u1ea1i.")
    Exit Sub
EH:
    PerfOff
    EndUndo
    Info U("L\u1ed7i khi chu\u1ea9n h\u00f3a: ") & Err.Description
End Sub

Public Sub ApplyPageSetup(ByRef r As TRule)
    Dim s As Section
    For Each s In ActiveDocument.Sections
        With s.PageSetup
            If .Orientation = wdOrientPortrait Then
                .PaperSize = wdPaperA4
                .TopMargin = Mm(r.MarginDef(M_TOP))
                .BottomMargin = Mm(r.MarginDef(M_BOTTOM))
                .LeftMargin = Mm(r.MarginDef(M_LEFT))
                .RightMargin = Mm(r.MarginDef(M_RIGHT))
            End If
        End With
    Next s
End Sub

Private Sub ApplyNormalStyle(ByRef r As TRule)
    With ActiveDocument.Styles(wdStyleNormal)
        .Font.Name = r.FontName
        .Font.Size = r.BodySizeDef
        .Font.Color = wdColorAutomatic
    End With
End Sub

Private Sub ApplyParagraphLayout(ByVal p As Paragraph, ByRef r As TRule)
    With p.Format
        If r.Justify Then .Alignment = wdAlignParagraphJustify
        If r.LineExact Then
            .LineSpacingRule = wdLineSpaceExactly
            .LineSpacing = r.LineExactDef
        Else
            .LineSpacingRule = wdLineSpaceMultiple
            .LineSpacing = Application.LinesToPoints(r.LineMultDef)
        End If
        If .SpaceAfter < r.SpaceAfterPt Then .SpaceAfter = r.SpaceAfterPt
        .FirstLineIndent = Mm(r.FirstIndentMm)
    End With
End Sub

Private Sub ApplyBody(ByRef r As TRule)
    Dim doc As Document, n As Long, k As Long, p As Paragraph, rg As Range, sz As Single
    Set doc = ActiveDocument
    n = doc.Paragraphs.Count
    For k = 1 To n
        If k Mod 50 = 0 Then Status U("ViFormat: \u0111ang chu\u1ea9n h\u00f3a \u0111o\u1ea1n ") & k & "/" & n
        Set p = doc.Paragraphs(k)
        If IsBodyParagraph(p) Then
            Set rg = TextRange(p)
            sz = rg.Font.Size
            If sz = UNDEF Or Not InRange(sz, r.BodySizeMin, r.BodySizeMax, 0.01) Then
                rg.Font.Size = r.BodySizeDef
            End If
            ApplyParagraphLayout p, r
        End If
    Next k
End Sub

Private Sub ApplyAnchors(ByRef r As TRule)
    Dim i As Integer, p As Paragraph, rg As Range
    For i = 1 To r.AnchorCount
        Set p = FindAnchor(r, i)
        If Not p Is Nothing Then
            Set rg = TextRange(p)
            rg.Font.Size = r.AnchorSizeDef(i)
            rg.Font.Bold = r.AnchorBold(i)
            rg.Font.Italic = r.AnchorItalic(i)
            If r.AnchorUnder(i) Then
                rg.Font.Underline = wdUnderlineSingle
            Else
                rg.Font.Underline = wdUnderlineNone
            End If
        End If
    Next i
End Sub

Public Sub ApplyPageNumber(ByRef r As TRule)
    Dim s As Section, hdr As HeaderFooter, rg As Range
    For Each s In ActiveDocument.Sections
        s.PageSetup.DifferentFirstPageHeaderFooter = True
        s.PageSetup.HeaderDistance = Mm(r.PageNumTopMm)
    Next s
    Set hdr = ActiveDocument.Sections(1).Headers(wdHeaderFooterPrimary)
    If Not HasPageField(hdr.Range) Then
        Set rg = hdr.Range
        rg.Text = ""
        rg.ParagraphFormat.Alignment = wdAlignParagraphCenter
        ActiveDocument.Fields.Add rg, wdFieldPage
    End If
    With hdr.Range
        .ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Font.Name = r.FontName
        .Font.Size = r.PageNumSize
        .Font.Bold = False
    End With
End Sub
