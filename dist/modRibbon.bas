Attribute VB_Name = "modRibbon"
Option Explicit
'==============================================================
' modRibbon - callback cho customUI14.xml
' Ten Sub phai khop dung thuoc tinh trong XML.
'==============================================================

Public gRibbon As Object

Public Sub Ribbon_OnLoad(ByVal ribbon As Object)
    Set gRibbon = ribbon
End Sub

'---- Dropdown chon quy chuan ----------------------------------
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
        Info r.Title & U(": ch\u01b0a c\u00f3 s\u1ed1 li\u1ec7u th\u1ec3 th\u1ee9c \u0111\u00e3 x\u00e1c minh.") & vbCrLf & vbCrLf & r.Notes
    End If
End Sub

'---- Nut ------------------------------------------------------
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
    Notify U("\u0110\u00e3 x\u00f3a ") & n & U(" nh\u1eadn x\u00e9t c\u1ee7a ViFormat.")
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
        U("T\u00e1c gi\u1ea3: ") & AppAuthor() & vbCrLf & _
        U("N\u1ec1n t\u1ea3ng: ") & IIf(IsMac(), "macOS", "Windows") & vbCrLf & vbCrLf & U("Quy chu\u1ea9n:") & vbCrLf
    For i = 0 To RuleCount() - 1
        r = GetRule(RuleIdAt(i))
        s = s & "- " & r.Title & ": " & IIf(r.Available, U("\u0111\u00e3 c\u00f3 s\u1ed1 li\u1ec7u"), U("CH\u01afA c\u00f3 s\u1ed1 li\u1ec7u")) & vbCrLf
    Next i
    s = s & vbCrLf & U("M\u1ecdi x\u1eed l\u00fd ch\u1ea1y tr\u00ean m\u00e1y, kh\u00f4ng g\u1eedi n\u1ed9i dung v\u0103n b\u1ea3n ra ngo\u00e0i.")
    Info s
End Sub
