Attribute VB_Name = "modEngine"
Option Explicit
'==============================================================
' modEngine - Chuẩn hóa và Rà soát theo một TRule.
' Chỉ dùng object model của Word, chạy được trên Windows và Mac.
'==============================================================

Private Const MAX_COMMENTS As Long = 40
Private Const UNDEF As Long = 9999999

Private mIssueCount As Long
Private mCommentCount As Long
Private mNotes As String

'--------------------------------------------------------------
' Tiện ích
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

' Đoạn nội dung: dài, không phải tiêu đề, không canh giữa/phải, không trong bảng
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
' Nhận xét
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
' RÀ SOÁT
'--------------------------------------------------------------
Public Sub ReviewDocument(ByVal ruleId As String)
    Dim r As TRule
    r = GetRule(ruleId)
    If Not r.Available Then
        Info "Quy chuẩn " & r.Title & " chưa có số liệu đã xác minh nên chưa thể rà soát." & _
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
        msg = "Không phát hiện lỗi thể thức theo " & r.Title & "."
    Else
        msg = "Phát hiện " & mIssueCount & " điểm chưa đúng theo " & r.Title & "."
        If mCommentCount > 0 Then msg = msg & vbCrLf & "Đã ghi " & mCommentCount & " nhận xét trong văn bản (tác giả ViFormat)."
        If mIssueCount > mCommentCount And mCommentCount >= MAX_COMMENTS Then _
            msg = msg & vbCrLf & "Chỉ ghi tối đa " & MAX_COMMENTS & " nhận xét đầu tiên."
        If Len(mNotes) > 0 Then msg = msg & vbCrLf & vbCrLf & mNotes
    End If
    If Len(r.Notes) > 0 Then msg = msg & vbCrLf & vbCrLf & "Lưu ý: " & r.Notes
    Notify msg
    Exit Sub
EH:
    PerfOff
    Info "Lỗi khi rà soát: " & Err.Description
End Sub

Private Sub CheckPageSetup(ByRef r As TRule)
    Dim s As Section, k As Integer, idx As Integer, v As Single, rg As Range
    Dim names As Variant
    names = Array("", "trên", "dưới", "trái", "phải")
    Set rg = ActiveDocument.Paragraphs(1).Range
    For Each s In ActiveDocument.Sections
        With s.PageSetup
            If s.Index = 1 Or .Orientation = wdOrientPortrait Then
                If Abs(.PageWidth - Application.MillimetersToPoints(210)) > 2 Or _
                   Abs(.PageHeight - Application.MillimetersToPoints(297)) > 2 Then
                    If .Orientation = wdOrientPortrait Then _
                        AddNote "Section " & s.Index & ": khổ giấy không phải A4 (210 x 297 mm)."
                End If
                For idx = 1 To 4
                    Select Case idx
                        Case M_TOP: v = ToMm(.TopMargin)
                        Case M_BOTTOM: v = ToMm(.BottomMargin)
                        Case M_LEFT: v = ToMm(.LeftMargin)
                        Case M_RIGHT: v = ToMm(.RightMargin)
                    End Select
                    If Not InRange(v, r.MarginMin(idx), r.MarginMax(idx), 0.6) Then
                        AddNote "Section " & s.Index & ": lề " & names(idx) & " = " & Format$(v, "0.0") & _
                                " mm, quy định " & RangeText(r.MarginMin(idx), r.MarginMax(idx)) & " mm."
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
            AddNote "Không tìm thấy " & r.AnchorName(i) & " (bỏ qua nếu văn bản không có thành phần này)."
        Else
            Set rg = TextRange(p)
            sz = rg.Font.Size
            sizes(i) = sz
            If sz = UNDEF Then
                AddIssue rg, r.AnchorName(i) & ": cỡ chữ không đồng nhất, quy định " & _
                         RangeText(r.AnchorSizeMin(i), r.AnchorSizeMax(i)) & "."
            ElseIf Not InRange(sz, r.AnchorSizeMin(i), r.AnchorSizeMax(i), 0.01) Then
                AddIssue rg, r.AnchorName(i) & ": cỡ chữ " & sz & ", quy định " & _
                         RangeText(r.AnchorSizeMin(i), r.AnchorSizeMax(i)) & "."
            End If
            If r.AnchorBold(i) And rg.Font.Bold <> True Then _
                AddIssue rg, r.AnchorName(i) & ": phải in đậm."
            If r.AnchorItalic(i) And rg.Font.Italic <> True Then _
                AddIssue rg, r.AnchorName(i) & ": phải in nghiêng."
            If r.AnchorUnder(i) And rg.Font.Underline = wdUnderlineNone Then _
                AddIssue rg, r.AnchorName(i) & ": phải gạch chân."
        End If
    Next i

    If r.TieuNguPlus1 And sizes(1) > 0 And sizes(1) <> UNDEF And sizes(2) > 0 And sizes(2) <> UNDEF Then
        If Abs(sizes(2) - (sizes(1) + 1)) > 0.01 Then
            Set p = FindAnchor(r, 2)
            If Not p Is Nothing Then
                AddIssue TextRange(p), "Tiêu ngữ phải lớn hơn Quốc hiệu 1 cỡ (Quốc hiệu " & sizes(1) & _
                         " thì Tiêu ngữ " & (sizes(1) + 1) & ")."
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
        If k Mod 50 = 0 Then Status "ViFormat: đang rà soát đoạn " & k & "/" & n
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

    If badFont > 0 Then AddIssue firstFont, "Phông chữ nội dung phải là " & r.FontName & _
        ". Có " & badFont & " đoạn sai (đây là đoạn đầu tiên)."
    If badSize > 0 Then AddIssue firstSize, "Cỡ chữ nội dung phải là " & _
        RangeText(r.BodySizeMin, r.BodySizeMax) & ". Có " & badSize & " đoạn sai."
    If badAlign > 0 Then AddIssue firstAlign, "Nội dung phải căn đều hai bên. Có " & badAlign & " đoạn sai."
    If badIndent > 0 Then AddIssue firstIndent, "Lùi đầu dòng phải khoảng " & r.FirstIndentMm & _
        " mm. Có " & badIndent & " đoạn sai."
    If badLine > 0 Then
        If r.LineExact Then
            AddIssue firstLine, "Giãn dòng phải là Exactly " & RangeText(r.LineExactMin, r.LineExactMax) & _
                " pt. Có " & badLine & " đoạn sai."
        Else
            AddIssue firstLine, "Giãn dòng phải trong khoảng " & RangeText(r.LineMultMin, r.LineMultMax) & _
                " dòng. Có " & badLine & " đoạn sai."
        End If
    End If
    If badSpace > 0 Then AddIssue firstSpace, "Khoảng cách giữa các đoạn tối thiểu " & r.SpaceAfterPt & _
        " pt. Có " & badSpace & " đoạn sai."
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
        AddNote "Chưa có số trang ở phần lề trên (canh giữa). Dùng nút Đánh số trang."
    Else
        If hdr.Range.ParagraphFormat.Alignment <> wdAlignParagraphCenter Then _
            AddNote "Số trang phải canh giữa."
    End If
    If Not s.PageSetup.DifferentFirstPageHeaderFooter Then _
        AddNote "Trang đầu không được hiển thị số trang (bật Different First Page)."
End Sub

'--------------------------------------------------------------
' CHUẨN HÓA
'--------------------------------------------------------------
Public Sub NormalizeDocument(ByVal ruleId As String)
    Dim r As TRule
    r = GetRule(ruleId)
    If Not r.Available Then
        Info "Quy chuẩn " & r.Title & " chưa có số liệu đã xác minh nên chưa thể chuẩn hóa." & _
             vbCrLf & vbCrLf & r.Notes
        Exit Sub
    End If

    BeginUndo "ViFormat: Chuẩn hóa"
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
    Notify "Đã chuẩn hóa theo " & r.Title & ". Có thể hoàn tác bằng Undo; nên chạy Rà soát thể thức để kiểm tra lại."
    Exit Sub
EH:
    PerfOff
    EndUndo
    Info "Lỗi khi chuẩn hóa: " & Err.Description
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
        If k Mod 50 = 0 Then Status "ViFormat: đang chuẩn hóa đoạn " & k & "/" & n
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
