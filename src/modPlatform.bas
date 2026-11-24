Attribute VB_Name = "modPlatform"
Option Explicit
'==============================================================
' modPlatform - tách riêng mọi thứ khác nhau giữa Windows và Mac.
' Quy tắc: KHÔNG dùng Scripting.Dictionary, FileSystemObject,
' Declare (Win32 API), ActiveX trong toàn bộ project.
'==============================================================

Public Const APP_NAME As String = "ViFormat"
Public Const APP_VERSION As String = "0.1.0"
Public Const REVIEW_AUTHOR As String = "ViFormat"

' VBA chỉ lưu mã nguồn theo bảng mã ANSI của máy, nên chuỗi tiếng Việt viết thẳng
' trong code sẽ hỏng trên máy không dùng bảng mã 1258 (và có thể hỏng trên Mac).
' tools/make_dist.py tự đổi mọi chuỗi có dấu thành U("...\uXXXX...") (thuần ASCII);
' hàm U() dưới đây giải mã lại thành Unicode lúc chạy. Nguồn trong src/ vẫn viết
' tiếng Việt bình thường cho dễ đọc.
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
    AppAuthor = "Phạm Văn Huy"   'không dùng Const vì dist đổi thành U(...)
End Function

Public Function IsMac() As Boolean
#If Mac Then
    IsMac = True
#Else
    IsMac = False
#End If
End Function

' Application.UndoRecord có từ Word 2010 trên Windows; trên Mac chưa chắc có.
' Gọi qua CallByName để nếu thiếu thì chỉ lỗi lúc chạy (bỏ qua được),
' không gây lỗi biên dịch cả project.
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

' MsgBox trên Mac đổi chuỗi sang MacRoman nên mất chữ Việt (ă, ề, Đ... thành "_").
' Mac: hiện hộp thoại qua AppleScriptTask (file ViFormat.scpt do install.command cài,
' hỗ trợ Unicode). Nếu chưa cài script thì dùng MsgBox với tiếng Việt không dấu.
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

' Hiện văn bản dài có thanh cuộn. Dòng bắt đầu bằng "#" là tiêu đề (in đậm).
' Mac: hộp thoại cuộn qua AppleScriptTask (ViFormat.scpt + ViFormat.js do install.command cài).
' Nếu không gọi được hoặc trên Windows: dùng MsgBox (bỏ dấu "#").
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

' Bỏ dấu tiếng Việt (dùng làm phương án dự phòng trên Mac).
Public Function NoAccent(ByVal s As String) As String
    Dim src As String, dst As String, i As Long, p As Long, ch As String, lc As String, r As String
    src = "àáảãạăằắẳẵặâầấẩẫậèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵđ"
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

' Thông báo kết quả không chặn thao tác: hiện ở thanh trạng thái, không popup.
' Mac: bỏ dấu vì thanh trạng thái/MsgBox của VBA Mac không chắc hiển thị Unicode.
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
    If Not HasDocument Then Info "Chưa mở văn bản nào."
End Function

' Lưu lựa chọn người dùng (quy chuẩn đang chọn). GetSetting/SaveSetting
' hoạt động trên cả Windows (registry) và Mac (plist trong sandbox).
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
