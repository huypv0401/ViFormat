Attribute VB_Name = "modRibbon"
Option Explicit
'==============================================================
' modRibbon - callback cho customUI14.xml
' Tên Sub phải khớp đúng thuộc tính trong XML.
'==============================================================

Public gRibbon As Object

Public Sub Ribbon_OnLoad(ByVal ribbon As Object)
    Set gRibbon = ribbon
End Sub

'---- Dropdown chọn quy chuẩn ----------------------------------
Public Sub DD_Rule_GetItemCount(ByVal control As Object, ByRef returnedVal As Variant)
    returnedVal = RuleCount()
End Sub

Public Sub DD_Rule_GetItemID(ByVal control As Object, ByVal index As Integer, ByRef returnedVal As Variant)
    returnedVal = RuleIdAt(index)
End Sub

Public Sub DD_Rule_GetItemLabel(ByVal control As Object, ByVal index As Integer, ByRef returnedVal As Variant)
    returnedVal = RuleLabelAt(index)
End Sub

Public Sub DD_Rule_GetSelectedItemID(ByVal control As Object, ByRef returnedVal As Variant)
    returnedVal = CurrentRuleId()
End Sub

Public Sub DD_Rule_OnAction(ByVal control As Object, ByVal id As String, ByVal index As Integer)
    SavePref "Rule", id
    Dim r As TRule
    r = GetRule(id)
    If Not r.Available Then
        Info r.Title & ": chưa có số liệu thể thức đã xác minh." & vbCrLf & vbCrLf & r.Notes
    End If
End Sub

'---- Nút ------------------------------------------------------
Public Sub Btn_RuleDetails(ByVal control As Object)
    Dim r As TRule
    r = GetRule(CurrentRuleId())
    ShowText RuleDetails(r)
End Sub

Public Sub Btn_Normalize(ByVal control As Object)
    If Not HasDocument() Then Exit Sub
    NormalizeDocument CurrentRuleId()
End Sub

Public Sub Btn_Review(ByVal control As Object)
    If Not HasDocument() Then Exit Sub
    ReviewDocument CurrentRuleId()
End Sub

Public Sub Btn_ClearReview(ByVal control As Object)
    If Not HasDocument() Then Exit Sub
    Dim n As Long
    n = ClearReviewComments()
    Notify "Đã xóa " & n & " nhận xét của ViFormat."
End Sub

Public Sub Btn_PageA4(ByVal control As Object)
    PageA4
End Sub

Public Sub Btn_Landscape(ByVal control As Object)
    InsertPageOrientation True
End Sub

Public Sub Btn_Portrait(ByVal control As Object)
    InsertPageOrientation False
End Sub

Public Sub Btn_RemoveBlank(ByVal control As Object)
    RemoveBlankPages
End Sub

Public Sub Btn_PageNumber(ByVal control As Object)
    AddPageNumber
End Sub

Public Sub Btn_About(ByVal control As Object)
    Dim s As String, i As Integer, r As TRule
    s = APP_NAME & " " & APP_VERSION & vbCrLf & _
        "Tác giả: " & AppAuthor() & vbCrLf & _
        "Nền tảng: " & IIf(IsMac(), "macOS", "Windows") & vbCrLf & vbCrLf & "Quy chuẩn:" & vbCrLf
    For i = 0 To RuleCount() - 1
        r = GetRule(RuleIdAt(i))
        s = s & "- " & r.Title & ": " & IIf(r.Available, "đã có số liệu", "CHƯA có số liệu") & vbCrLf
    Next i
    s = s & vbCrLf & "Mọi xử lý chạy trên máy, không gửi nội dung văn bản ra ngoài."
    Info s
End Sub
