' ImportModules.vbs  (chi chay tren Windows)
' Nap cac module trong thu muc dist\ vao ViFormat.dotm roi luu lai.
'
' Dieu kien: Word > File > Options > Trust Center > Trust Center Settings >
' Macro Settings > tick "Trust access to the VBA project object model".
'
' Cach dung: mo CMD tai thu muc project, chay:
'   python tools\make_dist.py
'   python build\build_dotm.py
'   cscript //nologo build\ImportModules.vbs

Option Explicit

Dim fso, root, dotm, distDir, word, doc, vbp, f, n

Set fso = CreateObject("Scripting.FileSystemObject")
root = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))
dotm = root & "\ViFormat.dotm"
distDir = root & "\dist"

If Not fso.FileExists(dotm) Then
    WScript.Echo "Khong thay " & dotm & ". Chay build\build_dotm.py truoc."
    WScript.Quit 1
End If

Set word = CreateObject("Word.Application")
word.Visible = False
word.DisplayAlerts = 0

On Error Resume Next
Set doc = word.Documents.Open(dotm)
If Err.Number <> 0 Then
    WScript.Echo "Khong mo duoc file: " & Err.Description
    word.Quit
    WScript.Quit 1
End If

Set vbp = doc.VBProject
If Err.Number <> 0 Then
    WScript.Echo "Word chan truy cap VBA project." & vbCrLf & _
        "Bat: File > Options > Trust Center > Trust Center Settings > Macro Settings > " & _
        "'Trust access to the VBA project object model', roi chay lai."
    doc.Close 0
    word.Quit
    WScript.Quit 1
End If
On Error GoTo 0

' Xoa module cu cung ten (neu chay lai)
Dim i, comp
For i = vbp.VBComponents.Count To 1 Step -1
    Set comp = vbp.VBComponents(i)
    If comp.Type = 1 Then vbp.VBComponents.Remove comp   ' 1 = module chuan
Next

n = 0
For Each f In fso.GetFolder(distDir).Files
    If LCase(fso.GetExtensionName(f.Name)) = "bas" Then
        vbp.VBComponents.Import f.Path
        n = n + 1
    End If
Next

doc.Save
doc.Close 0
word.Quit
WScript.Echo "Da nap " & n & " module vao " & dotm
