Attribute VB_Name = "modRules"
Option Explicit
'==============================================================
' modRules - bộ quy chuẩn thể thức.
'
' Đơn vị: lề tính bằng mm, cỡ chữ và khoảng cách dòng tính bằng pt.
' Available = False nghĩa là CHƯA có số liệu đã xác minh. Khi đó
' Chuẩn hóa / Rà soát sẽ từ chối chạy thay vì áp số liệu đoán.
' Số liệu nào chưa đối chiếu với văn bản gốc thì ghi vào Notes.
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

    ' Lề trang (mm), chỉ số 1..4 = trên, dưới, trái, phải
    MarginMin(1 To 4) As Single
    MarginMax(1 To 4) As Single
    MarginDef(1 To 4) As Single

    FontName As String

    ' Nội dung văn bản
    BodySizeMin As Single
    BodySizeMax As Single
    BodySizeDef As Single
    Justify As Boolean
    SpaceAfterPt As Single

    FirstIndentCheck As Boolean   ' True: đã có số liệu để kiểm tra
    FirstIndentMm As Single

    LineExact As Boolean          ' True: Exactly (pt); False: Multiple (số dòng)
    LineCheck As Boolean
    LineExactMin As Single
    LineExactMax As Single
    LineExactDef As Single
    LineMultMin As Single
    LineMultMax As Single
    LineMultDef As Single

    ' Số trang
    PageNumTopMm As Single
    PageNumSize As Single
    PageNumSizeMin As Single
    PageNumSizeMax As Single

    ' True: văn bản gốc quy định cụ thể căn đều / khoảng cách đoạn nên Rà soát được kiểm tra
    ParaCheck As Boolean

    ' Nội dung quy chuẩn hiển thị ở nút "Xem quy chuẩn" (nguyên văn do người dùng cung cấp)
    DetailText As String

    ' Các thành phần khối đầu / nơi nhận tìm theo chữ mở đầu của đoạn
    AnchorCount As Integer
    AnchorKey(1 To 4) As String      ' chữ mở đầu để nhận diện
    AnchorName(1 To 4) As String     ' tên hiển thị
    AnchorUpper(1 To 4) As Boolean
    AnchorBold(1 To 4) As Boolean
    AnchorItalic(1 To 4) As Boolean
    AnchorUnder(1 To 4) As Boolean
    AnchorSizeMin(1 To 4) As Single
    AnchorSizeMax(1 To 4) As Single
    AnchorSizeDef(1 To 4) As Single
    AnchorFromEnd(1 To 4) As Boolean ' True: tìm từ cuối văn bản (Nơi nhận)

    ' NĐ 30: Quốc hiệu cỡ N thì Tiêu ngữ cỡ N+1 (anchor 1 và 2)
    TieuNguPlus1 As Boolean
End Type

'--------------------------------------------------------------
' Danh sách hiển thị trên Ribbon
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
        Case 0: RuleLabelAt = "NĐ 30/2020/NĐ-CP (Hành chính nhà nước)"
        Case 1: RuleLabelAt = "HD 05-HD/VPTW (Văn bản của Đảng)"
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
' NĐ 30/2020/NĐ-CP
' Giãn dòng, khoảng cách đoạn, lùi đầu dòng: văn bản gốc KHÔNG quy định giá trị
' bắt buộc, nên Rà soát không kiểm tra (chỉ Chuẩn hóa áp giá trị mặc định thông dụng).
'--------------------------------------------------------------
Private Function BuildND30() As TRule
    Dim r As TRule
    r.Id = RULE_ND30
    r.Title = "NĐ 30/2020/NĐ-CP"
    r.Available = True
    r.Notes = "Giãn dòng, khoảng cách đoạn, lùi đầu dòng không có giá trị bắt buộc nên Rà soát không kiểm tra; " & _
              "Chuẩn hóa chỉ áp giá trị mặc định thông dụng."
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
    ' 1. Quốc hiệu: in hoa, đứng, đậm, cỡ 12-13
    r.AnchorKey(1) = "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM"
    r.AnchorName(1) = "Quốc hiệu"
    r.AnchorUpper(1) = True: r.AnchorBold(1) = True
    r.AnchorSizeMin(1) = 12: r.AnchorSizeMax(1) = 13: r.AnchorSizeDef(1) = 13
    ' 2. Tiêu ngữ: in thường, đứng, đậm, cỡ 13-14
    r.AnchorKey(2) = "Độc lập"
    r.AnchorName(2) = "Tiêu ngữ"
    r.AnchorBold(2) = True
    r.AnchorSizeMin(2) = 13: r.AnchorSizeMax(2) = 14: r.AnchorSizeDef(2) = 14

    r.TieuNguPlus1 = True
    Ln r, "KHỔ GIẤY"
    Ln r, "  A4 (210 x 297 mm), chiều dọc"
    Ln r, "  Trường hợp có bảng, biểu không lập thành phụ lục riêng: có thể trình bày ngang"
    Ln r, ""
    Ln r, "LỀ TRANG"
    Ln r, "  Trên: 20 - 25 mm"
    Ln r, "  Dưới: 20 - 25 mm"
    Ln r, "  Trái: 30 - 35 mm"
    Ln r, "  Phải: 15 - 20 mm"
    Ln r, ""
    Ln r, "PHÔNG CHỮ"
    Ln r, "  Times New Roman"
    Ln r, "  Bộ mã Unicode theo TCVN 6909:2001"
    Ln r, "  Màu chữ: đen"
    Ln r, "  Kiểu chữ: đứng"
    Ln r, ""
    Ln r, "SỐ TRANG"
    Ln r, "  Đánh từ số 1, chữ số Ả Rập"
    Ln r, "  Cỡ chữ: 13 - 14 pt"
    Ln r, "  Kiểu chữ: đứng"
    Ln r, "  Vị trí: canh giữa theo chiều ngang, nằm trong lề trên"
    Ln r, "  Trang đầu tiên: không hiển thị số trang"
    Ln r, ""
    Ln r, "QUỐC HIỆU"
    Ln r, "  Nội dung: “CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM”"
    Ln r, "  Vị trí: trên cùng, bên phải trang đầu"
    Ln r, "  Cỡ chữ: 12 - 13 pt"
    Ln r, "  Kiểu chữ: đứng, in hoa, đậm"
    Ln r, ""
    Ln r, "TIÊU NGỮ"
    Ln r, "  Nội dung: “Độc lập - Tự do - Hạnh phúc”"
    Ln r, "  Vị trí: canh giữa, dưới Quốc hiệu"
    Ln r, "  Cỡ chữ: 13 - 14 pt"
    Ln r, "  Kiểu chữ: đứng, in thường, đậm"
    Ln r, "  Viết hoa chữ cái đầu từng cụm từ"
    Ln r, "  Giữa các cụm từ có gạch nối “-” và khoảng cách"
    Ln r, "  Dưới Tiêu ngữ có đường kẻ ngang liền, dài bằng dòng chữ"
    Ln r, "  Quốc hiệu và Tiêu ngữ cách nhau một dòng đơn"
    Ln r, ""
    Ln r, "TÊN CƠ QUAN BAN HÀNH"
    Ln r, "  Tên cơ quan chủ quản trực tiếp: in hoa, 12 - 13 pt, kiểu đứng"
    Ln r, "  Tên cơ quan ban hành: in hoa, 12 - 13 pt, kiểu đứng, đậm"
    Ln r, "  Canh giữa; dưới tên có đường kẻ ngang liền"
    Ln r, "  Độ dài đường kẻ: 1/3 - 1/2 độ dài dòng chữ"
    Ln r, "  Các dòng tên cơ quan cách nhau dòng đơn"
    Ln r, ""
    Ln r, "SỐ, KÝ HIỆU VĂN BẢN"
    Ln r, "  Số văn bản: chữ số Ả Rập"
    Ln r, "  Số và ký hiệu là duy nhất trong năm"
    Ln r, "  Thống nhất giữa văn bản giấy và văn bản điện tử"
    Ln r, ""
    Ln r, "NỘI DUNG VĂN BẢN"
    Ln r, "  Cỡ chữ: 13 - 14 pt, theo quy định riêng cho từng thành phần thể thức"
    Ln r, "  Phông chữ: Times New Roman, màu đen"
    Ln r, "  Giãn dòng, khoảng cách đoạn, lùi đầu dòng: Nghị định 30/2020/NĐ-CP"
    Ln r, "  KHÔNG quy định giá trị bắt buộc cụ thể"
    Ln r, "  => Không nên kiểm tra cứng các tham số này; chỉ kiểm tra khi cơ quan có quy chế nội bộ riêng."
    Ln r, ""
    Ln r, "CHỮ KÝ, DẤU"
    Ln r, "  Văn bản giấy: ký bằng bút mực màu xanh, không dùng mực dễ phai"
    Ln r, "  Văn bản điện tử: ký số theo quy định"
    Ln r, "  Dấu đóng rõ, ngay ngắn, đúng chiều, mực đỏ"
    Ln r, "  Dấu trùm khoảng 1/3 chữ ký về phía bên trái"
    Ln r, "  Dấu giáp lai: đặt giữa mép phải, mỗi dấu tối đa 05 tờ"
    BuildND30 = r
End Function

'--------------------------------------------------------------
' Hướng dẫn 05-HD/VPTW
' Số liệu dưới đây là toàn bộ phần đã được cung cấp; giãn dòng, lùi đầu dòng và cỡ chữ
' các tiêu đề không nằm trong đó nên Rà soát không kiểm tra.
'--------------------------------------------------------------
Private Function BuildHD05() As TRule
    Dim r As TRule
    r.Id = RULE_HD05
    r.Title = "HD 05-HD/VPTW"
    r.Available = True
    r.Notes = "Giãn dòng, lùi đầu dòng và cỡ chữ các tiêu đề chưa có trong số liệu nên Rà soát không kiểm tra; " & _
              "Chuẩn hóa chỉ áp giá trị mặc định thông dụng."
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
    Ln r, "KHỔ GIẤY"
    Ln r, "  A4 (210 x 297 mm)"
    Ln r, "  Trình bày theo chiều dọc"
    Ln r, ""
    Ln r, "LỀ TRANG"
    Ln r, "  Trên: 20 - 25 mm"
    Ln r, "  Dưới: 20 - 25 mm"
    Ln r, "  Trái: 30 - 35 mm"
    Ln r, "  Phải: 15 - 20 mm"
    Ln r, ""
    Ln r, "PHÔNG CHỮ"
    Ln r, "  Times New Roman"
    Ln r, "  Màu đen"
    Ln r, ""
    Ln r, "CỠ CHỮ"
    Ln r, "  Nội dung văn bản: 13 - 14 pt"
    Ln r, "  Số trang: 13 - 14 pt"
    Ln r, ""
    Ln r, "SỐ TRANG"
    Ln r, "  Chữ số Ả Rập"
    Ln r, "  Canh giữa trong lề trên"
    Ln r, "  Không hiển thị ở trang đầu"
    Ln r, ""
    Ln r, "TIÊU ĐỀ BẮT BUỘC"
    Ln r, "  “Đảng Cộng sản Việt Nam”"
    Ln r, "  Tên cơ quan ban hành"
    Ln r, "  Số và ký hiệu văn bản"
    Ln r, "  Địa danh, ngày/tháng/năm ban hành"
    Ln r, "  Tên loại và trích yếu nội dung"
    Ln r, "  Nội dung văn bản"
    Ln r, "  Chức vụ, họ tên, chữ ký người có thẩm quyền"
    Ln r, "  Dấu cơ quan ban hành"
    Ln r, ""
    Ln r, "LƯU Ý"
    Ln r, "  Văn bản chứa bí mật nhà nước: áp dụng Hướng dẫn này"
    Ln r, "  kết hợp quy định bảo vệ bí mật nhà nước"
    Ln r, "  Văn bản chuyên ngành, sách, ấn phẩm: không thuộc phạm vi này"
    BuildHD05 = r
End Function


' Nội dung nút "Xem quy chuẩn". Dòng sát lề trái = tiêu đề mục (đánh dấu "#" để in đậm),
' dòng thụt vào = nội dung.
Public Function RuleDetails(ByRef r As TRule) As String
    Dim lines() As String, i As Long, out As String, t As String
    If Not r.Available Then
        RuleDetails = "#QUY CHUẨN: " & r.Title & vbLf & vbLf & "#CHƯA CÓ SỐ LIỆU" & vbLf & "     " & r.Notes & vbLf
        Exit Function
    End If
    out = "#QUY CHUẨN: " & r.Title & vbLf
    lines = Split(r.DetailText, vbLf)
    For i = 0 To UBound(lines)
        t = lines(i)
        If Len(t) = 0 Then
            ' bỏ dòng trống
        ElseIf Left$(t, 1) = " " Then
            out = out & "     " & Trim$(t) & vbLf
        Else
            out = out & vbLf & "#" & t & vbLf
        End If
    Next i
    RuleDetails = out
End Function
