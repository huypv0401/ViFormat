Attribute VB_Name = "modLayout"
Option Explicit
'==============================================================
' modLayout - bo cuc trang: A4, trang ngang/doc, trang trang, so trang
'==============================================================

Public Sub PageA4()
    If Not HasDocument() Then Exit Sub
    Dim r As TRule
    r = GetRule(CurrentRuleId())
    If Not r.Available Then r = GetRule(RULE_ND30)
    BeginUndo U("ViFormat: C\u0103n l\u1ec1 A4")
    On Error GoTo EH
    ApplyPageSetup r
    EndUndo
    Notify U("\u0110\u00e3 c\u0103n l\u1ec1 A4 theo ") & r.Title & U(": tr\u00ean ") & r.MarginDef(M_TOP) & U(", d\u01b0\u1edbi ") & r.MarginDef(M_BOTTOM) & _
         U(", tr\u00e1i ") & r.MarginDef(M_LEFT) & U(", ph\u1ea3i ") & r.MarginDef(M_RIGHT) & " mm."
    Exit Sub
EH:
    EndUndo
    Info U("L\u1ed7i: ") & Err.Description
End Sub

Public Sub InsertPageOrientation(ByVal landscape As Boolean)
    If Not HasDocument() Then Exit Sub
    BeginUndo U("ViFormat: Ch\u00e8n trang ") & IIf(landscape, "ngang", U("d\u1ecdc"))
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
    Info U("L\u1ed7i: ") & Err.Description
End Sub

' Xoa trang trang thua: doan trong o cuoi van ban va ngat trang kep lien tiep.
' Khong dung den doan trong o giua van ban.
Public Sub RemoveBlankPages()
    If Not HasDocument() Then Exit Sub
    Dim doc As Document, p As Paragraph, removed As Long, n As Long
    Set doc = ActiveDocument
    BeginUndo U("ViFormat: X\u00f3a trang tr\u1eafng")
    PerfOn
    On Error GoTo EH

    ' 1) Ngat trang kep: ^m^p^m  ->  ^m
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

    ' 2) Doan trong o cuoi van ban (chua lai doan cuoi cung bat buoc)
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
        Notify U("Kh\u00f4ng t\u00ecm th\u1ea5y trang tr\u1eafng th\u1eeba \u1edf cu\u1ed1i v\u0103n b\u1ea3n ho\u1eb7c ng\u1eaft trang k\u00e9p.")
    Else
        Notify U("\u0110\u00e3 x\u1eed l\u00fd ") & removed & U(" m\u1ee5c (ng\u1eaft trang k\u00e9p v\u00e0 \u0111o\u1ea1n tr\u1ed1ng cu\u1ed1i v\u0103n b\u1ea3n).")
    End If
    Exit Sub
EH:
    PerfOff
    EndUndo
    Info U("L\u1ed7i: ") & Err.Description
End Sub

Public Sub AddPageNumber()
    If Not HasDocument() Then Exit Sub
    Dim r As TRule
    r = GetRule(CurrentRuleId())
    If Not r.Available Then r = GetRule(RULE_ND30)
    BeginUndo U("ViFormat: \u0110\u00e1nh s\u1ed1 trang")
    On Error GoTo EH
    ApplyPageNumber r
    EndUndo
    Notify U("\u0110\u00e3 \u0111\u00e1nh s\u1ed1 trang (canh gi\u1eefa, kh\u00f4ng hi\u1ec3n th\u1ecb \u1edf trang \u0111\u1ea7u) theo ") & r.Title & "."
    Exit Sub
EH:
    EndUndo
    Info U("L\u1ed7i: ") & Err.Description
End Sub
