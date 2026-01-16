Attribute VB_Name = "模块_生成销售清单"
Option Explicit

'==================== 你可能要改的名字（只改这里） ====================
Private Const DATA_SHEET As String = "DataSheet"
Private Const TABLE_NAME As String = "DataTable"

Private Const TEMPLATE_SHEET As String = "TemplateSheet"
Private Const PRICE_SHEET As String = "PriceList"
Private Const AR_SHEET As String = "CustomerAR"
Private Const SETTINGS_SHEET As String = "Settings"

' DataTable 里列名（必须与表头完全一致）
Private Const COL_SHIPDATE As String = "出库日期"
Private Const COL_CUSTOMER As String = "出库对象"
Private Const COL_SPEC As String = "规格"
Private Const COL_NETW As String = "净重"
Private Const COL_BOOKED As String = "入账"
'=====================================================================

'==================== 固定版式参数（固定 15 行） ====================
Private Const START_ROW As Long = 5
Private Const FIXED_DETAIL_ROWS As Long = 15                ' 固定 15 行：5~19
Private Const DETAIL_END_ROW As Long = START_ROW + FIXED_DETAIL_ROWS - 1   ' 19
Private Const TOTALS_ROW As Long = DETAIL_END_ROW + 1        ' 20
Private Const TOTALS_ROW2 As Long = DETAIL_END_ROW + 2       ' 21

Private Const DETAIL_FIRST_COL As Long = 6                   ' F
Private Const DETAIL_COLS As Long = 10                       ' F~O (10列，满10个换行)
'=====================================================================

'=========================================================
' 固定15行规格区 + A4打印固定
'
' 1) 规格区固定为 15 行（5~19）。超出就提示并退出。
' 2) 每次生成前刷新清空固定区域，避免上一张残留。
' 3) A4打印设置固定：A4、纵向、强制一页、水平居中，保证大小位置一致。
'
' 数值格式：
' - 总重量/明细重量：1位小数
' - 单价：1位小数
' - 金额：2位小数（公式=ROUND(C*D,2)）
'
' 公式：
' - 金额 = 总重量 * 单价
' - 合计金额 = 各规格金额求和
' - 累计货款 = 合计金额 + 前欠货款
'
' 大写金额：动态绑定合计金额
'
' 【已兼容 Mac】：不使用 Scripting.Dictionary（避免 429 ActiveX 错误）
'=========================================================

'Public Sub 生成销货清单_Mac()
'    CreateSalesInvoice_Mac
'End Sub

Public Sub 选这个CreateSalesInvoice_Mac()
    On Error GoTo EH

    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim wsData As Worksheet, wsTpl As Worksheet, wsPrice As Worksheet, wsAR As Worksheet, wsSet As Worksheet
    Set wsData = wb.Worksheets(DATA_SHEET)
    Set wsTpl = wb.Worksheets(TEMPLATE_SHEET)
    Set wsPrice = wb.Worksheets(PRICE_SHEET)
    Set wsAR = wb.Worksheets(AR_SHEET)
    Set wsSet = wb.Worksheets(SETTINGS_SHEET)

    Dim lo As ListObject
    Set lo = wsData.ListObjects(TABLE_NAME)

    '==== 0) 每次生成前：先刷新固定区域（杜绝残留） ====
    RefreshTemplateFixedAreas wsTpl
    ApplyA4PrintSetup wsTpl

    '---- 从当前选中行获取 客户/日期（优先）
    Dim shipDate As Date, customer As String
    customer = ""
    shipDate = 0

    If ActiveSheet Is wsData Then
        If Not Intersect(ActiveCell, lo.DataBodyRange) Is Nothing Then
            Dim rr As Long
            rr = ActiveCell.Row - lo.HeaderRowRange.Row

            customer = CStr(lo.DataBodyRange.Cells(rr, lo.ListColumns(COL_CUSTOMER).Index).Value)

            Dim tmpD As Variant
            tmpD = lo.DataBodyRange.Cells(rr, lo.ListColumns(COL_SHIPDATE).Index).Value
            If IsDate(tmpD) Then shipDate = DateValue(tmpD)
        End If
    End If

    ' 兜底：手动输入
    If Trim$(customer) = "" Or shipDate = 0 Then
        customer = InputBox("请输入客户名称（出库对象）：", "生成销货清单")
        If Trim$(customer) = "" Then Exit Sub

        Dim sDate As String
        sDate = InputBox("请输入出库日期（yyyy-mm-dd）：", "生成销货清单")
        If Trim$(sDate) = "" Then Exit Sub
        shipDate = DateValue(Replace(Replace(sDate, ".", "-"), "/", "-"))
    End If

    '---- 汇总：按规格聚合（Mac/Win 通用：Collection + 类）
    Dim aggs As Collection
    Set aggs = New Collection

    Dim rowsToMark As Collection
    Set rowsToMark = New Collection

    Dim i As Long
    For i = 1 To lo.DataBodyRange.Rows.Count
        Dim d As Variant, cust As String, booked As String
        d = lo.DataBodyRange.Cells(i, lo.ListColumns(COL_SHIPDATE).Index).Value
        cust = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns(COL_CUSTOMER).Index).Value)
        booked = Trim$(CStr(lo.DataBodyRange.Cells(i, lo.ListColumns(COL_BOOKED).Index).Value))

        If cust = customer Then
            If IsDate(d) Then
                If DateValue(d) = shipDate Then
                    If booked <> "是" Then
                        Dim spec As String
                        spec = CStr(lo.DataBodyRange.Cells(i, lo.ListColumns(COL_SPEC).Index).Value)

                        Dim netw As Double
                        netw = CDbl(Val(lo.DataBodyRange.Cells(i, lo.ListColumns(COL_NETW).Index).Value))

                        If Trim$(spec) <> "" Then
                            Dim agg As CSpecAgg
                            Set agg = GetOrCreateAgg(aggs, spec)

                            agg.qty = agg.qty + 1
                            agg.totalW = agg.totalW + netw
                            agg.details.Add netw

                            rowsToMark.Add i
                        End If
                    End If
                End If
            End If
        End If
    Next i

    If aggs.Count = 0 Then
        MsgBox "没有找到符合条件且未入账的记录（同日期 + 同客户）。", vbInformation
        Exit Sub
    End If

    '---- 1) 计算本次需要占用多少“行”（10个明细=1行），并检查是否超出 15 行
    Dim keys() As String
    keys = AggKeysSorted(aggs)

    Dim requiredRows As Long: requiredRows = 0
    Dim detailInfo As String: detailInfo = ""

    For i = LBound(keys) To UBound(keys)
        Dim agg2 As CSpecAgg
        Set agg2 = aggs(keys(i))

        Dim qty As Long, needRows As Long
        qty = agg2.qty
        needRows = (qty + DETAIL_COLS - 1) \ DETAIL_COLS  ' ceil(qty/10)
        If needRows < 1 Then needRows = 1

        requiredRows = requiredRows + needRows
        detailInfo = detailInfo & keys(i) & "：箱数" & qty & "，占用行数" & needRows & vbCrLf
    Next i

    If requiredRows > FIXED_DETAIL_ROWS Then
        MsgBox "规格/明细超出固定 15 行，已停止生成！" & vbCrLf & vbCrLf & _
               "本次需要行数：" & requiredRows & "（最多允许 15）" & vbCrLf & _
               "请减少规格/箱数或改为分两张单打印。" & vbCrLf & vbCrLf & _
               "明细占用：" & vbCrLf & detailInfo, vbExclamation
        Exit Sub
    End If

    '---- 2) 通过检查后，再生成清单号（避免超出时也把编号+1）
    Dim invoiceNo As String
    invoiceNo = NextInvoiceNo(wsSet)

    '---- 表头（中文）
    wsTpl.Range("A3").Value = "客户：" & customer
    wsTpl.Range("F3").Value = "日期：" & Format(shipDate, "yyyy-mm-dd")
    wsTpl.Range("I2").Value = "No. " & invoiceNo

    '---- 3) 写入产品行 + 明细网格（固定范围内）
    Dim curRow As Long: curRow = START_ROW

    For i = LBound(keys) To UBound(keys)
        Dim k As String: k = keys(i)

        Dim agg3 As CSpecAgg
        Set agg3 = aggs(k)

        Dim q As Long: q = agg3.qty
        Dim totalW As Double: totalW = agg3.totalW
        Dim details As Collection: Set details = agg3.details

        Dim blockRows As Long
        blockRows = (q + DETAIL_COLS - 1) \ DETAIL_COLS
        If blockRows < 1 Then blockRows = 1

        ' 产品行（只写第一行）
        wsTpl.Cells(curRow, "A").Value = k
        wsTpl.Cells(curRow, "B").Value = q

        Dim unitPrice As Double
        unitPrice = Round(GetUnitPrice(k, wsPrice), 1)

        wsTpl.Cells(curRow, "C").Value = Round(totalW, 1)
        wsTpl.Cells(curRow, "D").Value = unitPrice

        ' 金额公式（动态）
        wsTpl.Cells(curRow, "E").Formula = "=ROUND(C" & curRow & "*D" & curRow & ",2)"

        ' 数字格式
        wsTpl.Cells(curRow, "C").NumberFormat = "0.0"
        wsTpl.Cells(curRow, "D").NumberFormat = "0.0"
        wsTpl.Cells(curRow, "E").NumberFormat = "0.00"

        ' 明细网格：F~O
        Dim j As Long
        For j = 1 To details.Count
            Dim rOff As Long, cOff As Long
            rOff = (j - 1) \ DETAIL_COLS
            cOff = (j - 1) Mod DETAIL_COLS

            wsTpl.Cells(curRow + rOff, DETAIL_FIRST_COL + cOff).Value = Round(CDbl(details(j)), 1)
            wsTpl.Cells(curRow + rOff, DETAIL_FIRST_COL + cOff).NumberFormat = "0.0"
        Next j

        ' 清空续行的 A~E（避免样式残留）
        Dim r As Long
        For r = 1 To blockRows - 1
            wsTpl.Range("A" & (curRow + r) & ":E" & (curRow + r)).ClearContents
        Next r

        curRow = curRow + blockRows
    Next i

    '---- 4) 合计/累计（辅助 R1~R3）
    Dim prevDebt As Double
    prevDebt = GetCustomerDebt(customer, wsAR)

    wsTpl.Range("R1").Value = Round(prevDebt, 2)
    wsTpl.Range("R1").NumberFormat = "0.00"

    wsTpl.Range("R2").Formula = "=SUM($E$" & START_ROW & ":$E$" & DETAIL_END_ROW & ")"
    wsTpl.Range("R2").NumberFormat = "0.00"

    wsTpl.Range("R3").Formula = "=ROUND($R$1+$R$2,2)"
    wsTpl.Range("R3").NumberFormat = "0.00"

    ' 合计区显示（固定位置：TOTALS_ROW / TOTALS_ROW2）
    wsTpl.Calculate                  ' 先让 R1~R3 的公式算完
    RefreshTotalsDisplay wsTpl       ' 再用 VBA 写入显示文本（避免 @ 和 #NAME?）


    '---- 5) 回写入账=是 + 更新欠款表
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim idx As Variant
    For Each idx In rowsToMark
        lo.DataBodyRange.Cells(CLng(idx), lo.ListColumns(COL_BOOKED).Index).Value = "是"
    Next idx

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    ' 更新欠款表：新欠款 = 前欠 + 本单合计
    Dim totalAmount As Double
    totalAmount = Round(CDbl(wsTpl.Range("R2").Value), 2)
    SetCustomerDebt customer, Round(prevDebt + totalAmount, 2), wsAR

    MsgBox "已生成清单：" & invoiceNo & vbCrLf & _
           "客户：" & customer & vbCrLf & _
           "日期：" & Format(shipDate, "yyyy-mm-dd"), vbInformation
    Exit Sub

EH:
    MsgBox "错误 " & Err.Number & "：" & Err.Description, vbExclamation
    On Error Resume Next
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

'===================== 每次生成前：刷新固定区域 ======================
Private Sub RefreshTemplateFixedAreas(ByVal wsTpl As Worksheet)
    wsTpl.Range("A" & START_ROW & ":O" & DETAIL_END_ROW).ClearContents
    wsTpl.Range("A" & TOTALS_ROW & ":O" & TOTALS_ROW2).ClearContents
    wsTpl.Range("R1:R3").ClearContents
End Sub

'===================== A4打印固定设置（保证大小位置一致） ======================
Private Sub ApplyA4PrintSetup(ByVal ws As Worksheet)
    On Error Resume Next

    With ws.PageSetup
        .PaperSize = xlPaperA4
        .Orientation = xlPortrait

        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1

        .CenterHorizontally = True
        .CenterVertically = False

        .LeftMargin = Application.CentimetersToPoints(0.8)
        .RightMargin = Application.CentimetersToPoints(0.8)
        .TopMargin = Application.CentimetersToPoints(0.8)
        .BottomMargin = Application.CentimetersToPoints(0.8)

        .PrintGridlines = False
        .PrintHeadings = False

        If .PrintArea = "" Then
            .PrintArea = ws.Range("A1:O60").Address
        End If
    End With
End Sub
Public Sub RefreshTotalsDisplay(ByVal wsTpl As Worksheet)
    ' 先让 R1~R3 计算完成（你手动改单价后也会依赖它们）
    wsTpl.Calculate

    Dim prevDebt As Double, totalAmount As Double, cumDebt As Double
    prevDebt = CDbl(Val(wsTpl.Range("R1").Value))
    totalAmount = CDbl(Val(wsTpl.Range("R2").Value))
    cumDebt = CDbl(Val(wsTpl.Range("R3").Value))

    wsTpl.Range("A" & TOTALS_ROW).Value = "合计金额(大写)： " & RMBUPPER(totalAmount)
    wsTpl.Range("A" & TOTALS_ROW2).Value = "合计金额（小写）：￥" & Format$(totalAmount, "#,##0.00")
    wsTpl.Range("E" & TOTALS_ROW2).Value = "前欠货款：￥" & Format$(prevDebt, "#,##0.00")
    wsTpl.Range("I" & TOTALS_ROW2).Value = "累计货款：￥" & Format$(cumDebt, "#,##0.00")
End Sub


'======================== 兼容 Mac 的“按规格聚合”辅助 ========================
Private Function GetOrCreateAgg(ByVal aggs As Collection, ByVal spec As String) As CSpecAgg
    Dim a As CSpecAgg

    On Error Resume Next
    Set a = aggs(spec)      ' 不存在会报错
    On Error GoTo 0

    If a Is Nothing Then
        Set a = New CSpecAgg
        a.spec = spec
        aggs.Add a, spec
    End If

    Set GetOrCreateAgg = a
End Function

Private Function AggKeysSorted(ByVal aggs As Collection) As String()
    Dim arr() As String
    ReDim arr(0 To aggs.Count - 1)

    Dim i As Long: i = 0
    Dim a As CSpecAgg
    For Each a In aggs
        arr(i) = a.spec
        i = i + 1
    Next a

    Dim x As Long, y As Long
    For x = LBound(arr) To UBound(arr) - 1
        For y = x + 1 To UBound(arr)
            If arr(x) > arr(y) Then
                Dim tmp As String
                tmp = arr(x): arr(x) = arr(y): arr(y) = tmp
            End If
        Next y
    Next x

    AggKeysSorted = arr
End Function

'======================== 依赖函数：单价/欠款/编号/大写金额 ========================
Private Function GetUnitPrice(ByVal spec As String, ByVal wsPrice As Worksheet) As Double
    On Error GoTo EH
    GetUnitPrice = Application.WorksheetFunction.VLookup(spec, wsPrice.Range("A:B"), 2, False)
    Exit Function
EH:
    GetUnitPrice = 0#
End Function

Private Function GetCustomerDebt(ByVal customer As String, ByVal wsAR As Worksheet) As Double
    Dim lastRow As Long: lastRow = wsAR.Cells(wsAR.Rows.Count, "A").End(xlUp).Row
    Dim r As Long

    For r = 2 To lastRow
        If Trim$(CStr(wsAR.Cells(r, "A").Value)) = Trim$(customer) Then
            GetCustomerDebt = CDbl(Val(wsAR.Cells(r, "B").Value))
            Exit Function
        End If
    Next r

    GetCustomerDebt = 0#
End Function

Private Sub SetCustomerDebt(ByVal customer As String, ByVal newDebt As Double, ByVal wsAR As Worksheet)
    Dim lastRow As Long: lastRow = wsAR.Cells(wsAR.Rows.Count, "A").End(xlUp).Row
    Dim r As Long

    For r = 2 To lastRow
        If Trim$(CStr(wsAR.Cells(r, "A").Value)) = Trim$(customer) Then
            wsAR.Cells(r, "B").Value = newDebt
            wsAR.Cells(r, "B").NumberFormat = "0.00"
            Exit Sub
        End If
    Next r

    wsAR.Cells(lastRow + 1, "A").Value = customer
    wsAR.Cells(lastRow + 1, "B").Value = newDebt
    wsAR.Cells(lastRow + 1, "B").NumberFormat = "0.00"
End Sub

Private Function NextInvoiceNo(ByVal wsSet As Worksheet) As String
    Dim prefix As String: prefix = Trim$(CStr(wsSet.Range("B1").Value))
    If prefix = "" Then prefix = "0001"

    Dim n As Long: n = CLng(Val(wsSet.Range("B2").Value))
    n = n + 1
    wsSet.Range("B2").Value = n

    NextInvoiceNo = prefix & "-" & Format$(n, "00000000")
End Function

' 工作表公式可调用：=RMBUPPER(R2)
Public Function RMBUPPER(ByVal amount As Double) As String
    RMBUPPER = RMBUpperInternal(amount)
End Function

' 人民币大写（到亿元，角分）
Private Function RMBUpperInternal(ByVal amount As Double) As String
    Dim CNNum, CNUnit
    CNNum = Array("零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖")
    CNUnit = Array("", "拾", "佰", "仟", "万", "拾", "佰", "仟", "亿")

    Dim s As String
    s = Format$(amount, "0.00")

    Dim intPart As String, decPart As String
    intPart = Split(s, ".")(0)
    decPart = Split(s, ".")(1)

    Dim res As String: res = ""
    Dim i As Long, n As Integer, pos As Long
    pos = 0

    For i = Len(intPart) To 1 Step -1
        n = CInt(Mid$(intPart, i, 1))

        Dim unitStr As String: unitStr = CNUnit(pos)

        If n = 0 Then
            If Left$(res, 1) <> "零" And res <> "" Then res = "零" & res
        Else
            res = CNNum(n) & unitStr & res
        End If

        If pos = 4 And Len(intPart) > 4 Then
            If Left$(res, 1) <> "万" Then res = "万" & res
        ElseIf pos = 8 And Len(intPart) > 8 Then
            If Left$(res, 1) <> "亿" Then res = "亿" & res
        End If

        pos = pos + 1
        If pos > 8 Then pos = 0
    Next i

    Do While InStr(res, "零零") > 0
        res = Replace(res, "零零", "零")
    Loop
    If Right$(res, 1) = "零" Then res = Left$(res, Len(res) - 1)
    If res = "" Then res = "零"
    res = res & "元"

    Dim jiao As Integer: jiao = CInt(Mid$(decPart, 1, 1))
    Dim fen As Integer: fen = CInt(Mid$(decPart, 2, 1))

    If jiao = 0 And fen = 0 Then
        RMBUpperInternal = res & "整"
        Exit Function
    End If

    If jiao <> 0 Then
        res = res & CNNum(jiao) & "角"
    Else
        res = res & "零"
    End If

    If fen <> 0 Then
        res = res & CNNum(fen) & "分"
    End If

    RMBUpperInternal = res
End Function


