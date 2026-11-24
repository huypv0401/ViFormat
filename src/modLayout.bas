Attribute VB_Name = "modLayout"
Option Explicit
'==============================================================
' modLayout - bố cục trang: A4, trang ngang/dọc, trang trắng, số trang
'==============================================================

Public Sub PageA4()
    If Not HasDocument() Then Exit Sub
    Dim r As TRule
    r = GetRule(CurrentRuleId())
    If Not r.Available Then r = GetRule(RULE_ND30)
    BeginUndo "ViFormat: Căn lề A4"
    On Error GoTo EH
    ApplyPageSetup r
    EndUndo
    Notify "Đã căn lề A4 theo " & r.Title & ": trên " & r.MarginDef(M_TOP) & ", dưới " & r.MarginDef(M_BOTTOM) & _
         ", trái " & r.MarginDef(M_LEFT) & ", phải " & r.MarginDef(M_RIGHT) & " mm."
    Exit Sub
EH:
    EndUndo
    Info "Lỗi: " & Err.Description
End Sub

Public Sub InsertPageOrientation(ByVal landscape As Boolean)
    If Not HasDocument() Then Exit Sub
    BeginUndo "ViFormat: Chèn trang " & IIf(landscape, "ngang", "dọc")
    On Error GoTo EH
    Selection.InsertBreak wdSectionBreakNextPage
    With Selection.Sections(1).PageSetup
        If landscape Then
            .Orientation = wdOrientLandscape
        Else
            .Orientation = wdOrientPortrait
        End If
    End With
    EndUndo
    Exit Sub
EH:
    EndUndo
    Info "Lỗi: " & Err.Description
End Sub

' Xóa trang trắng thừa: đoạn trống ở cuối văn bản và ngắt trang kép liên tiếp.
' Không đụng đến đoạn trống ở giữa văn bản.
Public Sub RemoveBlankPages()
    If Not HasDocument() Then Exit Sub
    Dim doc As Document, p As Paragraph, removed As Long, n As Long
    Set doc = ActiveDocument
    BeginUndo "ViFormat: Xóa trang trắng"
    PerfOn
    On Error GoTo EH

    ' 1) Ngắt trang kép: ^m^p^m  ->  ^m
    With doc.Content.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = "^m^p^m"
        .Replacement.Text = "^m"
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = False
        Do While .Execute(Replace:=wdReplaceOne)
            removed = removed + 1
            If removed > 500 Then Exit Do
        Loop
    End With

    ' 2) Đoạn trống ở cuối văn bản (chừa lại đoạn cuối cùng bắt buộc)
    n = doc.Paragraphs.Count
    Do While n > 1
        Set p = doc.Paragraphs(n - 1)
        If Len(Replace(Replace(p.Range.Text, vbCr, ""), Chr(12), "")) = 0 _
           And p.Range.InlineShapes.Count = 0 _
           And Not p.Range.Information(wdWithInTable) Then
            p.Range.Delete
            removed = removed + 1
            n = doc.Paragraphs.Count
        Else
            Exit Do
        End If
        If removed > 2000 Then Exit Do
    Loop

    PerfOff
    EndUndo
    If removed = 0 Then
        Notify "Không tìm thấy trang trắng thừa ở cuối văn bản hoặc ngắt trang kép."
    Else
        Notify "Đã xử lý " & removed & " mục (ngắt trang kép và đoạn trống cuối văn bản)."
    End If
    Exit Sub
EH:
    PerfOff
    EndUndo
    Info "Lỗi: " & Err.Description
End Sub

Public Sub AddPageNumber()
    If Not HasDocument() Then Exit Sub
    Dim r As TRule
    r = GetRule(CurrentRuleId())
    If Not r.Available Then r = GetRule(RULE_ND30)
    BeginUndo "ViFormat: Đánh số trang"
    On Error GoTo EH
    ApplyPageNumber r
    EndUndo
    Notify "Đã đánh số trang (canh giữa, không hiển thị ở trang đầu) theo " & r.Title & "."
    Exit Sub
EH:
    EndUndo
    Info "Lỗi: " & Err.Description
End Sub
