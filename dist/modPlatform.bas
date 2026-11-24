Attribute VB_Name = "modPlatform"
Option Explicit
'==============================================================
' modPlatform - tach rieng moi thu khac nhau giua Windows va Mac.
' Quy tac: KHONG dung Scripting.Dictionary, FileSystemObject,
' Declare (Win32 API), ActiveX trong toan bo project.
'==============================================================

Public Const APP_NAME As String = "ViFormat"
Public Const APP_VERSION As String = "0.1.0"
Public Const REVIEW_AUTHOR As String = "ViFormat"

' VBA chi luu ma nguon theo bang ma ANSI cua may, nen chuoi tieng Viet viet thang
' trong code se hong tren may khong dung bang ma 1258 (va co the hong tren Mac).
' tools/make_dist.py tu doi moi chuoi co dau thanh U("...\uXXXX...") (thuan ASCII);
' ham U() duoi day giai ma lai thanh Unicode luc chay. Nguon trong src/ van viet
' tieng Viet binh thuong cho de doc.
Public Function U(ByVal s As String) As String
    Dim i As Long, out As String
    i = 1
    Do While i <= Len(s)
        If Mid$(s, i, 2) = "\u" And i + 5 <= Len(s) Then
            out = out & ChrW(CLng("&H" & Mid$(s, i + 2, 4)))
            i = i + 6
        Else
            out = out & Mid$(s, i, 1)
            i = i + 1
        End If
    Loop
    U = out
End Function

Public Function AppAuthor() As String
    AppAuthor = U("Ph\u1ea1m V\u0103n Huy")   'khong dung Const vi dist doi thanh U(...)
End Function

Public Function IsMac() As Boolean
#If Mac Then
    IsMac = True
#Else
    IsMac = False
#End If
End Function

' Application.UndoRecord co tu Word 2010 tren Windows; tren Mac chua chac co.
' Goi qua CallByName de neu thieu thi chi loi luc chay (bo qua duoc),
' khong gay loi bien dich ca project.
Public Sub BeginUndo(ByVal title As String)
    On Error Resume Next
    Dim u As Object
    Set u = CallByName(Application, "UndoRecord", VbGet)
    If Not u Is Nothing Then
        If Not u.IsRecordingCustomRecord Then u.StartCustomRecord title
    End If
    On Error GoTo 0
End Sub

Public Sub EndUndo()
    On Error Resume Next
    Dim u As Object
    Set u = CallByName(Application, "UndoRecord", VbGet)
    If Not u Is Nothing Then
        If u.IsRecordingCustomRecord Then u.EndCustomRecord
    End If
    On Error GoTo 0
End Sub

Public Sub PerfOn()
    On Error Resume Next
    Application.ScreenUpdating = False
    On Error GoTo 0
End Sub

Public Sub PerfOff()
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.StatusBar = False
    On Error GoTo 0
End Sub

Public Sub Status(ByVal s As String)
    On Error Resume Next
    Application.StatusBar = s
    On Error GoTo 0
End Sub

' MsgBox tren Mac doi chuoi sang MacRoman nen mat chu Viet (a, e, D... thanh "_").
' Mac: hien hop thoai qua AppleScriptTask (file ViFormat.scpt do install.command cai,
' ho tro Unicode). Neu chua cai script thi dung MsgBox voi tieng Viet khong dau.
Public Sub Info(ByVal msg As String)
#If Mac Then
    Dim ok As Boolean
    ok = False
    On Error Resume Next
    AppleScriptTask "ViFormat.scpt", "ShowMsg", msg
    ok = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    If Not ok Then MsgBox NoAccent(msg), vbInformation, APP_NAME
#Else
    MsgBox msg, vbInformation, APP_NAME
#End If
End Sub

' Hien van ban dai co thanh cuon. Dong bat dau bang "#" la tieu de (in dam).
' Mac: hop thoai cuon qua AppleScriptTask (ViFormat.scpt + ViFormat.js do install.command cai).
' Neu khong goi duoc hoac tren Windows: dung MsgBox (bo dau "#").
Public Sub ShowText(ByVal msg As String)
#If Mac Then
    Dim ok As Boolean
    ok = False
    On Error Resume Next
    AppleScriptTask "ViFormat.scpt", "ShowText", msg
    ok = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    If Not ok Then MsgBox NoAccent(Replace(msg, "#", "")), vbInformation, APP_NAME
#Else
    MsgBox Replace(msg, "#", ""), vbInformation, APP_NAME
#End If
End Sub

' Bo dau tieng Viet (dung lam phuong an du phong tren Mac).
Public Function NoAccent(ByVal s As String) As String
    Dim src As String, dst As String, i As Long, p As Long, ch As String, lc As String, r As String
    src = U("\u00e0\u00e1\u1ea3\u00e3\u1ea1\u0103\u1eb1\u1eaf\u1eb3\u1eb5\u1eb7\u00e2\u1ea7\u1ea5\u1ea9\u1eab\u1ead\u00e8\u00e9\u1ebb\u1ebd\u1eb9\u00ea\u1ec1\u1ebf\u1ec3\u1ec5\u1ec7\u00ec\u00ed\u1ec9\u0129\u1ecb\u00f2\u00f3\u1ecf\u00f5\u1ecd\u00f4\u1ed3\u1ed1\u1ed5\u1ed7\u1ed9\u01a1\u1edd\u1edb\u1edf\u1ee1\u1ee3\u00f9\u00fa\u1ee7\u0169\u1ee5\u01b0\u1eeb\u1ee9\u1eed\u1eef\u1ef1\u1ef3\u00fd\u1ef7\u1ef9\u1ef5\u0111")
    dst = String$(17, "a") & String$(11, "e") & String$(5, "i") & String$(17, "o") & _
          String$(11, "u") & String$(5, "y") & "d"
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        lc = LCase$(ch)
        p = InStr(1, src, lc, vbBinaryCompare)
        If p > 0 Then
            r = Mid$(dst, p, 1)
            If lc <> ch Then r = UCase$(r)
            NoAccent = NoAccent & r
        Else
            NoAccent = NoAccent & ch
        End If
    Next i
End Function

' Thong bao ket qua khong chan thao tac: hien o thanh trang thai, khong popup.
' Mac: bo dau vi thanh trang thai/MsgBox cua VBA Mac khong chac hien thi Unicode.
Public Sub Notify(ByVal msg As String)
    Dim t As String
    t = Replace(Replace(msg, vbCrLf, " "), vbLf, " ")
#If Mac Then
    t = NoAccent(t)
#End If
    Status t
End Sub

Public Function HasDocument() As Boolean
    HasDocument = (Documents.Count > 0)
    If Not HasDocument Then Info U("Ch\u01b0a m\u1edf v\u0103n b\u1ea3n n\u00e0o.")
End Function

' Luu lua chon nguoi dung (quy chuan dang chon). GetSetting/SaveSetting
' hoat dong tren ca Windows (registry) va Mac (plist trong sandbox).
Public Function LoadPref(ByVal key As String, ByVal def As String) As String
    On Error Resume Next
    LoadPref = GetSetting("ViFormat", "Prefs", key, def)
    If Err.Number <> 0 Then LoadPref = def
    On Error GoTo 0
End Function

Public Sub SavePref(ByVal key As String, ByVal value As String)
    On Error Resume Next
    SaveSetting "ViFormat", "Prefs", key, value
    On Error GoTo 0
End Sub
