Attribute VB_Name = "modRules"
Option Explicit
'==============================================================
' modRules - bo quy chuan the thuc.
'
' Don vi: le tinh bang mm, co chu va khoang cach dong tinh bang pt.
' Available = False nghia la CHUA co so lieu da xac minh. Khi do
' Chuan hoa / Ra soat se tu choi chay thay vi ap so lieu doan.
' So lieu nao chua doi chieu voi van ban goc thi ghi vao Notes.
'==============================================================

Public Const RULE_ND30 As String = "ND30"
Public Const RULE_HD05 As String = "HD05"

Public Const M_TOP As Integer = 1
Public Const M_BOTTOM As Integer = 2
Public Const M_LEFT As Integer = 3
Public Const M_RIGHT As Integer = 4

Public Type TRule
    Id As String
    Title As String
    Available As Boolean
    Notes As String

    ' Le trang (mm), chi so 1..4 = tren, duoi, trai, phai
    MarginMin(1 To 4) As Single
    MarginMax(1 To 4) As Single
    MarginDef(1 To 4) As Single

    FontName As String

    ' Noi dung van ban
    BodySizeMin As Single
    BodySizeMax As Single
    BodySizeDef As Single
    Justify As Boolean
    SpaceAfterPt As Single

    FirstIndentCheck As Boolean   ' True: da co so lieu de kiem tra
    FirstIndentMm As Single

    LineExact As Boolean          ' True: Exactly (pt); False: Multiple (so dong)
    LineCheck As Boolean
    LineExactMin As Single
    LineExactMax As Single
    LineExactDef As Single
    LineMultMin As Single
    LineMultMax As Single
    LineMultDef As Single

    ' So trang
    PageNumTopMm As Single
    PageNumSize As Single
    PageNumSizeMin As Single
    PageNumSizeMax As Single

    ' True: van ban goc quy dinh cu the can deu / khoang cach doan nen Ra soat duoc kiem tra
    ParaCheck As Boolean

    ' Noi dung quy chuan hien thi o nut "Xem quy chuan" (nguyen van do nguoi dung cung cap)
    DetailText As String

    ' Cac thanh phan khoi dau / noi nhan tim theo chu mo dau cua doan
    AnchorCount As Integer
    AnchorKey(1 To 4) As String      ' chu mo dau de nhan dien
    AnchorName(1 To 4) As String     ' ten hien thi
    AnchorUpper(1 To 4) As Boolean
    AnchorBold(1 To 4) As Boolean
    AnchorItalic(1 To 4) As Boolean
    AnchorUnder(1 To 4) As Boolean
    AnchorSizeMin(1 To 4) As Single
    AnchorSizeMax(1 To 4) As Single
    AnchorSizeDef(1 To 4) As Single
    AnchorFromEnd(1 To 4) As Boolean ' True: tim tu cuoi van ban (Noi nhan)

    ' ND 30: Quoc hieu co N thi Tieu ngu co N+1 (anchor 1 va 2)
    TieuNguPlus1 As Boolean
End Type

'--------------------------------------------------------------
' Danh sach hien thi tren Ribbon
'--------------------------------------------------------------
Public Function RuleCount() As Integer
    RuleCount = 2
End Function

Public Function RuleIdAt(ByVal i As Integer) As String
    Select Case i
        Case 0: RuleIdAt = RULE_ND30
        Case 1: RuleIdAt = RULE_HD05
    End Select
End Function

Public Function RuleLabelAt(ByVal i As Integer) As String
    Select Case i
        Case 0: RuleLabelAt = U("N\u0110 30/2020/N\u0110-CP (H\u00e0nh ch\u00ednh nh\u00e0 n\u01b0\u1edbc)")
        Case 1: RuleLabelAt = U("HD 05-HD/VPTW (V\u0103n b\u1ea3n c\u1ee7a \u0110\u1ea3ng)")
    End Select
End Function

Public Function RuleIndexOf(ByVal ruleId As String) As Integer
    Dim i As Integer
    RuleIndexOf = 0
    For i = 0 To RuleCount() - 1
        If RuleIdAt(i) = ruleId Then RuleIndexOf = i: Exit Function
    Next i
End Function

Public Function CurrentRuleId() As String
    CurrentRuleId = LoadPref("Rule", RULE_ND30)
    If RuleIndexOf(CurrentRuleId) = 0 And CurrentRuleId <> RULE_ND30 Then
        CurrentRuleId = RULE_ND30
    End If
End Function

Public Function GetRule(ByVal ruleId As String) As TRule
    Select Case ruleId
        Case RULE_ND30: GetRule = BuildND30()
        Case RULE_HD05: GetRule = BuildHD05()
        Case Else: GetRule = BuildND30()
    End Select
End Function

Private Sub Ln(ByRef r As TRule, ByVal s As String)
    r.DetailText = r.DetailText & s & vbLf
End Sub

'--------------------------------------------------------------
' ND 30/2020/ND-CP
' Gian dong, khoang cach doan, lui dau dong: van ban goc KHONG quy dinh gia tri
' bat buoc, nen Ra soat khong kiem tra (chi Chuan hoa ap gia tri mac dinh thong dung).
'--------------------------------------------------------------
Private Function BuildND30() As TRule
    Dim r As TRule
    r.Id = RULE_ND30
    r.Title = U("N\u0110 30/2020/N\u0110-CP")
    r.Available = True
    r.Notes = U("Gi\u00e3n d\u00f2ng, kho\u1ea3ng c\u00e1ch \u0111o\u1ea1n, l\u00f9i \u0111\u1ea7u d\u00f2ng kh\u00f4ng c\u00f3 gi\u00e1 tr\u1ecb b\u1eaft bu\u1ed9c n\u00ean R\u00e0 so\u00e1t kh\u00f4ng ki\u1ec3m tra; ") & _
              U("Chu\u1ea9n h\u00f3a ch\u1ec9 \u00e1p gi\u00e1 tr\u1ecb m\u1eb7c \u0111\u1ecbnh th\u00f4ng d\u1ee5ng.")
    r.FontName = "Times New Roman"

    r.MarginMin(M_TOP) = 20: r.MarginMax(M_TOP) = 25: r.MarginDef(M_TOP) = 20
    r.MarginMin(M_BOTTOM) = 20: r.MarginMax(M_BOTTOM) = 25: r.MarginDef(M_BOTTOM) = 20
    r.MarginMin(M_LEFT) = 30: r.MarginMax(M_LEFT) = 35: r.MarginDef(M_LEFT) = 30
    r.MarginMin(M_RIGHT) = 15: r.MarginMax(M_RIGHT) = 20: r.MarginDef(M_RIGHT) = 15

    r.BodySizeMin = 13: r.BodySizeMax = 14: r.BodySizeDef = 14
    r.Justify = True
    r.ParaCheck = False
    r.SpaceAfterPt = 6
    r.FirstIndentCheck = False
    r.FirstIndentMm = 10
    r.LineExact = False
    r.LineCheck = False
    r.LineMultMin = 1: r.LineMultMax = 1.5: r.LineMultDef = 1.15

    r.PageNumTopMm = 10
    r.PageNumSize = 14: r.PageNumSizeMin = 13: r.PageNumSizeMax = 14

    r.AnchorCount = 2
    ' 1. Quoc hieu: in hoa, dung, dam, co 12-13
    r.AnchorKey(1) = U("C\u1ed8NG H\u00d2A X\u00c3 H\u1ed8I CH\u1ee6 NGH\u0128A VI\u1ec6T NAM")
    r.AnchorName(1) = U("Qu\u1ed1c hi\u1ec7u")
    r.AnchorUpper(1) = True: r.AnchorBold(1) = True
    r.AnchorSizeMin(1) = 12: r.AnchorSizeMax(1) = 13: r.AnchorSizeDef(1) = 13
    ' 2. Tieu ngu: in thuong, dung, dam, co 13-14
    r.AnchorKey(2) = U("\u0110\u1ed9c l\u1eadp")
    r.AnchorName(2) = U("Ti\u00eau ng\u1eef")
    r.AnchorBold(2) = True
    r.AnchorSizeMin(2) = 13: r.AnchorSizeMax(2) = 14: r.AnchorSizeDef(2) = 14

    r.TieuNguPlus1 = True
    Ln r, U("KH\u1ed4 GI\u1ea4Y")
    Ln r, U("  A4 (210 x 297 mm), chi\u1ec1u d\u1ecdc")
    Ln r, U("  Tr\u01b0\u1eddng h\u1ee3p c\u00f3 b\u1ea3ng, bi\u1ec3u kh\u00f4ng l\u1eadp th\u00e0nh ph\u1ee5 l\u1ee5c ri\u00eang: c\u00f3 th\u1ec3 tr\u00ecnh b\u00e0y ngang")
    Ln r, ""
    Ln r, U("L\u1ec0 TRANG")
    Ln r, U("  Tr\u00ean: 20 - 25 mm")
    Ln r, U("  D\u01b0\u1edbi: 20 - 25 mm")
    Ln r, U("  Tr\u00e1i: 30 - 35 mm")
    Ln r, U("  Ph\u1ea3i: 15 - 20 mm")
    Ln r, ""
    Ln r, U("PH\u00d4NG CH\u1eee")
    Ln r, "  Times New Roman"
    Ln r, U("  B\u1ed9 m\u00e3 Unicode theo TCVN 6909:2001")
    Ln r, U("  M\u00e0u ch\u1eef: \u0111en")
    Ln r, U("  Ki\u1ec3u ch\u1eef: \u0111\u1ee9ng")
    Ln r, ""
    Ln r, U("S\u1ed0 TRANG")
    Ln r, U("  \u0110\u00e1nh t\u1eeb s\u1ed1 1, ch\u1eef s\u1ed1 \u1ea2 R\u1eadp")
    Ln r, U("  C\u1ee1 ch\u1eef: 13 - 14 pt")
    Ln r, U("  Ki\u1ec3u ch\u1eef: \u0111\u1ee9ng")
    Ln r, U("  V\u1ecb tr\u00ed: canh gi\u1eefa theo chi\u1ec1u ngang, n\u1eb1m trong l\u1ec1 tr\u00ean")
    Ln r, U("  Trang \u0111\u1ea7u ti\u00ean: kh\u00f4ng hi\u1ec3n th\u1ecb s\u1ed1 trang")
    Ln r, ""
    Ln r, U("QU\u1ed0C HI\u1ec6U")
    Ln r, U("  N\u1ed9i dung: \u201cC\u1ed8NG H\u00d2A X\u00c3 H\u1ed8I CH\u1ee6 NGH\u0128A VI\u1ec6T NAM\u201d")
    Ln r, U("  V\u1ecb tr\u00ed: tr\u00ean c\u00f9ng, b\u00ean ph\u1ea3i trang \u0111\u1ea7u")
    Ln r, U("  C\u1ee1 ch\u1eef: 12 - 13 pt")
    Ln r, U("  Ki\u1ec3u ch\u1eef: \u0111\u1ee9ng, in hoa, \u0111\u1eadm")
    Ln r, ""
    Ln r, U("TI\u00caU NG\u1eee")
    Ln r, U("  N\u1ed9i dung: \u201c\u0110\u1ed9c l\u1eadp - T\u1ef1 do - H\u1ea1nh ph\u00fac\u201d")
    Ln r, U("  V\u1ecb tr\u00ed: canh gi\u1eefa, d\u01b0\u1edbi Qu\u1ed1c hi\u1ec7u")
    Ln r, U("  C\u1ee1 ch\u1eef: 13 - 14 pt")
    Ln r, U("  Ki\u1ec3u ch\u1eef: \u0111\u1ee9ng, in th\u01b0\u1eddng, \u0111\u1eadm")
    Ln r, U("  Vi\u1ebft hoa ch\u1eef c\u00e1i \u0111\u1ea7u t\u1eebng c\u1ee5m t\u1eeb")
    Ln r, U("  Gi\u1eefa c\u00e1c c\u1ee5m t\u1eeb c\u00f3 g\u1ea1ch n\u1ed1i \u201c-\u201d v\u00e0 kho\u1ea3ng c\u00e1ch")
    Ln r, U("  D\u01b0\u1edbi Ti\u00eau ng\u1eef c\u00f3 \u0111\u01b0\u1eddng k\u1ebb ngang li\u1ec1n, d\u00e0i b\u1eb1ng d\u00f2ng ch\u1eef")
    Ln r, U("  Qu\u1ed1c hi\u1ec7u v\u00e0 Ti\u00eau ng\u1eef c\u00e1ch nhau m\u1ed9t d\u00f2ng \u0111\u01a1n")
    Ln r, ""
    Ln r, U("T\u00caN C\u01a0 QUAN BAN H\u00c0NH")
    Ln r, U("  T\u00ean c\u01a1 quan ch\u1ee7 qu\u1ea3n tr\u1ef1c ti\u1ebfp: in hoa, 12 - 13 pt, ki\u1ec3u \u0111\u1ee9ng")
    Ln r, U("  T\u00ean c\u01a1 quan ban h\u00e0nh: in hoa, 12 - 13 pt, ki\u1ec3u \u0111\u1ee9ng, \u0111\u1eadm")
    Ln r, U("  Canh gi\u1eefa; d\u01b0\u1edbi t\u00ean c\u00f3 \u0111\u01b0\u1eddng k\u1ebb ngang li\u1ec1n")
    Ln r, U("  \u0110\u1ed9 d\u00e0i \u0111\u01b0\u1eddng k\u1ebb: 1/3 - 1/2 \u0111\u1ed9 d\u00e0i d\u00f2ng ch\u1eef")
    Ln r, U("  C\u00e1c d\u00f2ng t\u00ean c\u01a1 quan c\u00e1ch nhau d\u00f2ng \u0111\u01a1n")
    Ln r, ""
    Ln r, U("S\u1ed0, K\u00dd HI\u1ec6U V\u0102N B\u1ea2N")
    Ln r, U("  S\u1ed1 v\u0103n b\u1ea3n: ch\u1eef s\u1ed1 \u1ea2 R\u1eadp")
    Ln r, U("  S\u1ed1 v\u00e0 k\u00fd hi\u1ec7u l\u00e0 duy nh\u1ea5t trong n\u0103m")
    Ln r, U("  Th\u1ed1ng nh\u1ea5t gi\u1eefa v\u0103n b\u1ea3n gi\u1ea5y v\u00e0 v\u0103n b\u1ea3n \u0111i\u1ec7n t\u1eed")
    Ln r, ""
    Ln r, U("N\u1ed8I DUNG V\u0102N B\u1ea2N")
    Ln r, U("  C\u1ee1 ch\u1eef: 13 - 14 pt, theo quy \u0111\u1ecbnh ri\u00eang cho t\u1eebng th\u00e0nh ph\u1ea7n th\u1ec3 th\u1ee9c")
    Ln r, U("  Ph\u00f4ng ch\u1eef: Times New Roman, m\u00e0u \u0111en")
    Ln r, U("  Gi\u00e3n d\u00f2ng, kho\u1ea3ng c\u00e1ch \u0111o\u1ea1n, l\u00f9i \u0111\u1ea7u d\u00f2ng: Ngh\u1ecb \u0111\u1ecbnh 30/2020/N\u0110-CP")
    Ln r, U("  KH\u00d4NG quy \u0111\u1ecbnh gi\u00e1 tr\u1ecb b\u1eaft bu\u1ed9c c\u1ee5 th\u1ec3")
    Ln r, U("  => Kh\u00f4ng n\u00ean ki\u1ec3m tra c\u1ee9ng c\u00e1c tham s\u1ed1 n\u00e0y; ch\u1ec9 ki\u1ec3m tra khi c\u01a1 quan c\u00f3 quy ch\u1ebf n\u1ed9i b\u1ed9 ri\u00eang.")
    Ln r, ""
    Ln r, U("CH\u1eee K\u00dd, D\u1ea4U")
    Ln r, U("  V\u0103n b\u1ea3n gi\u1ea5y: k\u00fd b\u1eb1ng b\u00fat m\u1ef1c m\u00e0u xanh, kh\u00f4ng d\u00f9ng m\u1ef1c d\u1ec5 phai")
    Ln r, U("  V\u0103n b\u1ea3n \u0111i\u1ec7n t\u1eed: k\u00fd s\u1ed1 theo quy \u0111\u1ecbnh")
    Ln r, U("  D\u1ea5u \u0111\u00f3ng r\u00f5, ngay ng\u1eafn, \u0111\u00fang chi\u1ec1u, m\u1ef1c \u0111\u1ecf")
    Ln r, U("  D\u1ea5u tr\u00f9m kho\u1ea3ng 1/3 ch\u1eef k\u00fd v\u1ec1 ph\u00eda b\u00ean tr\u00e1i")
    Ln r, U("  D\u1ea5u gi\u00e1p lai: \u0111\u1eb7t gi\u1eefa m\u00e9p ph\u1ea3i, m\u1ed7i d\u1ea5u t\u1ed1i \u0111a 05 t\u1edd")
    BuildND30 = r
End Function

'--------------------------------------------------------------
' Huong dan 05-HD/VPTW
' So lieu duoi day la toan bo phan da duoc cung cap; gian dong, lui dau dong va co chu
' cac tieu de khong nam trong do nen Ra soat khong kiem tra.
'--------------------------------------------------------------
Private Function BuildHD05() As TRule
    Dim r As TRule
    r.Id = RULE_HD05
    r.Title = "HD 05-HD/VPTW"
    r.Available = True
    r.Notes = U("Gi\u00e3n d\u00f2ng, l\u00f9i \u0111\u1ea7u d\u00f2ng v\u00e0 c\u1ee1 ch\u1eef c\u00e1c ti\u00eau \u0111\u1ec1 ch\u01b0a c\u00f3 trong s\u1ed1 li\u1ec7u n\u00ean R\u00e0 so\u00e1t kh\u00f4ng ki\u1ec3m tra; ") & _
              U("Chu\u1ea9n h\u00f3a ch\u1ec9 \u00e1p gi\u00e1 tr\u1ecb m\u1eb7c \u0111\u1ecbnh th\u00f4ng d\u1ee5ng.")
    r.FontName = "Times New Roman"

    r.MarginMin(M_TOP) = 20: r.MarginMax(M_TOP) = 25: r.MarginDef(M_TOP) = 20
    r.MarginMin(M_BOTTOM) = 20: r.MarginMax(M_BOTTOM) = 25: r.MarginDef(M_BOTTOM) = 20
    r.MarginMin(M_LEFT) = 30: r.MarginMax(M_LEFT) = 35: r.MarginDef(M_LEFT) = 30
    r.MarginMin(M_RIGHT) = 15: r.MarginMax(M_RIGHT) = 20: r.MarginDef(M_RIGHT) = 15

    r.BodySizeMin = 13: r.BodySizeMax = 14: r.BodySizeDef = 14
    r.Justify = True
    r.ParaCheck = False
    r.SpaceAfterPt = 6
    r.FirstIndentCheck = False
    r.FirstIndentMm = 10
    r.LineExact = False
    r.LineCheck = False
    r.LineMultMin = 1: r.LineMultMax = 1.5: r.LineMultDef = 1.15

    r.PageNumTopMm = 10
    r.PageNumSize = 14: r.PageNumSizeMin = 13: r.PageNumSizeMax = 14

    r.AnchorCount = 0
    r.TieuNguPlus1 = False
    Ln r, U("KH\u1ed4 GI\u1ea4Y")
    Ln r, "  A4 (210 x 297 mm)"
    Ln r, U("  Tr\u00ecnh b\u00e0y theo chi\u1ec1u d\u1ecdc")
    Ln r, ""
    Ln r, U("L\u1ec0 TRANG")
    Ln r, U("  Tr\u00ean: 20 - 25 mm")
    Ln r, U("  D\u01b0\u1edbi: 20 - 25 mm")
    Ln r, U("  Tr\u00e1i: 30 - 35 mm")
    Ln r, U("  Ph\u1ea3i: 15 - 20 mm")
    Ln r, ""
    Ln r, U("PH\u00d4NG CH\u1eee")
    Ln r, "  Times New Roman"
    Ln r, U("  M\u00e0u \u0111en")
    Ln r, ""
    Ln r, U("C\u1ee0 CH\u1eee")
    Ln r, U("  N\u1ed9i dung v\u0103n b\u1ea3n: 13 - 14 pt")
    Ln r, U("  S\u1ed1 trang: 13 - 14 pt")
    Ln r, ""
    Ln r, U("S\u1ed0 TRANG")
    Ln r, U("  Ch\u1eef s\u1ed1 \u1ea2 R\u1eadp")
    Ln r, U("  Canh gi\u1eefa trong l\u1ec1 tr\u00ean")
    Ln r, U("  Kh\u00f4ng hi\u1ec3n th\u1ecb \u1edf trang \u0111\u1ea7u")
    Ln r, ""
    Ln r, U("TI\u00caU \u0110\u1ec0 B\u1eaeT BU\u1ed8C")
    Ln r, U("  \u201c\u0110\u1ea3ng C\u1ed9ng s\u1ea3n Vi\u1ec7t Nam\u201d")
    Ln r, U("  T\u00ean c\u01a1 quan ban h\u00e0nh")
    Ln r, U("  S\u1ed1 v\u00e0 k\u00fd hi\u1ec7u v\u0103n b\u1ea3n")
    Ln r, U("  \u0110\u1ecba danh, ng\u00e0y/th\u00e1ng/n\u0103m ban h\u00e0nh")
    Ln r, U("  T\u00ean lo\u1ea1i v\u00e0 tr\u00edch y\u1ebfu n\u1ed9i dung")
    Ln r, U("  N\u1ed9i dung v\u0103n b\u1ea3n")
    Ln r, U("  Ch\u1ee9c v\u1ee5, h\u1ecd t\u00ean, ch\u1eef k\u00fd ng\u01b0\u1eddi c\u00f3 th\u1ea9m quy\u1ec1n")
    Ln r, U("  D\u1ea5u c\u01a1 quan ban h\u00e0nh")
    Ln r, ""
    Ln r, U("L\u01afU \u00dd")
    Ln r, U("  V\u0103n b\u1ea3n ch\u1ee9a b\u00ed m\u1eadt nh\u00e0 n\u01b0\u1edbc: \u00e1p d\u1ee5ng H\u01b0\u1edbng d\u1eabn n\u00e0y")
    Ln r, U("  k\u1ebft h\u1ee3p quy \u0111\u1ecbnh b\u1ea3o v\u1ec7 b\u00ed m\u1eadt nh\u00e0 n\u01b0\u1edbc")
    Ln r, U("  V\u0103n b\u1ea3n chuy\u00ean ng\u00e0nh, s\u00e1ch, \u1ea5n ph\u1ea9m: kh\u00f4ng thu\u1ed9c ph\u1ea1m vi n\u00e0y")
    BuildHD05 = r
End Function


' Noi dung nut "Xem quy chuan". Dong sat le trai = tieu de muc (danh dau "#" de in dam),
' dong thut vao = noi dung.
Public Function RuleDetails(ByRef r As TRule) As String
    Dim lines() As String, i As Long, out As String, t As String
    If Not r.Available Then
        RuleDetails = U("#QUY CHU\u1ea8N: ") & r.Title & vbLf & vbLf & U("#CH\u01afA C\u00d3 S\u1ed0 LI\u1ec6U") & vbLf & "     " & r.Notes & vbLf
        Exit Function
    End If
    out = U("#QUY CHU\u1ea8N: ") & r.Title & vbLf
    lines = Split(r.DetailText, vbLf)
    For i = 0 To UBound(lines)
        t = lines(i)
        If Len(t) = 0 Then
            ' bo dong trong
        ElseIf Left$(t, 1) = " " Then
            out = out & "     " & Trim$(t) & vbLf
        Else
            out = out & vbLf & "#" & t & vbLf
        End If
    Next i
    RuleDetails = out
End Function
