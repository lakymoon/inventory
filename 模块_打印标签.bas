Attribute VB_Name = "模块_打印标签"
Option Explicit

'================= 配置（只改这里） =================
Private Const DATA_SHEET As String = "DataSheet"

' 默认模板（每次打印默认用它）
Private Const DEFAULT_BTW_FILE As String = "空白标签.btw"

' 打印机名（留空=使用btw里保存的默认打印机）
Private Const PRINTER_NAME As String = ""

' 打印成功回写值
Private Const PRINT_DONE_VALUE As String = "是"

' 表头列名（必须与表头完全一致）
Private Const COL_SHIPDATE As String = "出库日期"
Private Const COL_PRINTFLAG As String = "是否打印"
'====================================================

'================= 运行一次切换：只对下一次打印生效 =================
Private gTempBTWFile As String  ' 临时模板：下一次打印使用；打印结束自动清空


'================= 主入口：一键打印标签（按当前行出库日期） =================
Public Sub 一键打印标签()
    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(DATA_SHEET)

    If ActiveSheet.Name <> ws.Name Then
        MsgBox "请先切换到工作表：" & DATA_SHEET, vbExclamation
        Exit Sub
    End If

    If ActiveCell.Row < 2 Then
        MsgBox "请先点击任意一条数据行（第2行及以后），再执行打印。", vbExclamation
        Exit Sub
    End If

    ' 找列
    Dim colShip As Long, colFlag As Long
    colShip = FindHeaderCol(ws, COL_SHIPDATE)
    colFlag = FindHeaderCol(ws, COL_PRINTFLAG)

    If colShip = 0 Or colFlag = 0 Then
        MsgBox "找不到必要列：[" & COL_SHIPDATE & "] 或 [" & COL_PRINTFLAG & "]。" & vbCrLf & _
               "请检查表头是否完全一致。", vbCritical
        Exit Sub
    End If

    ' 目标出库日期 = 当前行出库日期
    Dim targetShipDate As Variant
    targetShipDate = ws.Cells(ActiveCell.Row, colShip).Value

    If IsEmpty(targetShipDate) Or Trim$(CStr(targetShipDate)) = "" Then
        MsgBox "当前行【出库日期】为空，无法按组打印。", vbExclamation
        Exit Sub
    End If

    ' 收集本组需要打印的行（同出库日期 且 是否打印不是“是”）
    Dim rowsToPrint() As Long
    Dim needCount As Long
    needCount = CollectRowsToPrint(ws, colShip, colFlag, targetShipDate, rowsToPrint)

    If needCount = 0 Then
        MsgBox "该【出库日期】下没有需要打印的记录。" & vbCrLf & _
               "出库日期 = " & SafeDateText(targetShipDate), vbInformation
        Exit Sub
    End If

    ' ===== 选择本次使用的模板（默认空白；若已切换则用临时模板）=====
    Dim btwFile As String
    btwFile = DEFAULT_BTW_FILE
    If Len(Trim$(gTempBTWFile)) > 0 Then btwFile = gTempBTWFile

    Dim ans As VbMsgBoxResult
    ans = MsgBox("本次使用模板：" & btwFile & vbCrLf & _
                 "将打印【出库日期】= " & SafeDateText(targetShipDate) & vbCrLf & _
                 "共 " & needCount & " 条未打印记录。" & vbCrLf & vbCrLf & _
                 "是否继续？（点“否”可取消，重新选择行）", _
                 vbQuestion + vbYesNo, "确认打印")
    If ans <> vbYes Then GoTo CLEANUP_RESET_DEFAULT

    ' btw 路径（同目录）
    Dim btwPath As String
    btwPath = ThisWorkbook.Path & "\" & btwFile
    If Dir(btwPath) = "" Then
        MsgBox "未找到标签文件：" & vbCrLf & btwPath & vbCrLf & _
               "请确认btw与库存.xlsm在同一目录。", vbCritical
        GoTo CLEANUP_RESET_DEFAULT
    End If

    ' 连接 BarTender（老版本兼容：Late Binding）
    Dim btApp As Object, btFmt As Object
    Set btApp = CreateObject("BarTender.Application")
    btApp.Visible = False

    Set btFmt = btApp.Formats.Open(btwPath, False, "")
    If Len(PRINTER_NAME) > 0 Then
        On Error Resume Next
        btFmt.Printer = PRINTER_NAME
        On Error GoTo EH
    End If

    ' 执行打印：逐行写入变量 -> 打印 -> 成功回写“是”
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Dim i As Long, r As Long
    Dim printed As Long, failed As Long
    Dim errMsg As String

    For i = LBound(rowsToPrint) To UBound(rowsToPrint)
        r = rowsToPrint(i)
        If r = 0 Then GoTo NextI

        ' 再次防呆：如果已经是“是”，跳过（避免中途有人改）
        If IsPrintedFlag(ws.Cells(r, colFlag).Value, PRINT_DONE_VALUE) Then GoTo NextI

        ' 写入 NamedSubStrings（变量名=表头）
        SetNamedSubStringsFromRow btFmt, ws, r

        ' 兼容老 BarTender 的打印调用：多种签名依次尝试
        If PrintFormat_Compat(btFmt, errMsg) Then
            ws.Cells(r, colFlag).Value = PRINT_DONE_VALUE
            printed = printed + 1
        Else
            failed = failed + 1
            ' 失败不回写，保留原值，方便重新打印
        End If

NextI:
    Next i

    ' 保存
    ThisWorkbook.Save

    MsgBox "打印完成：" & vbCrLf & _
           "模板 = " & btwFile & vbCrLf & _
           "出库日期 = " & SafeDateText(targetShipDate) & vbCrLf & _
           "成功打印: " & printed & " 条" & vbCrLf & _
           "失败: " & failed & " 条" & IIf(failed > 0, vbCrLf & "失败原因（最后一次）: " & errMsg, ""), _
           IIf(failed > 0, vbExclamation, vbInformation)

CLEANUP:
    On Error Resume Next
    btFmt.Close False
    btApp.Quit
    Set btFmt = Nothing
    Set btApp = Nothing

    Application.EnableEvents = True
    Application.ScreenUpdating = True

CLEANUP_RESET_DEFAULT:
    ' ===== 打印结束/取消/失败：自动恢复默认（清空临时模板）=====
    gTempBTWFile = ""

    Exit Sub

EH:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "发生错误：" & Err.Number & vbCrLf & Err.Description, vbCritical
    Resume CLEANUP_RESET_DEFAULT
End Sub


'================= 切换标签模板：列出同目录所有 .btw（仅对下一次打印生效） =================
Public Sub 切换标签模板()
    Dim list As Collection
    Set list = ListBTWFilesInSameFolder()

    If list.Count = 0 Then
        MsgBox "当前目录没有找到 .btw 文件。" & vbCrLf & ThisWorkbook.Path, vbExclamation
        Exit Sub
    End If

    Dim msg As String, i As Long
    msg = "请输入要使用的模板编号（仅对下一次打印生效）：" & vbCrLf & vbCrLf
    For i = 1 To list.Count
        msg = msg & i & ". " & CStr(list(i)) & vbCrLf
    Next i

    Dim pick As String
    pick = InputBox(msg, "切换标签模板（临时）", "1")
    If Trim$(pick) = "" Then Exit Sub

    If Not IsNumeric(pick) Then
        MsgBox "请输入数字编号。", vbExclamation
        Exit Sub
    End If

    Dim idx As Long
    idx = CLng(pick)
    If idx < 1 Or idx > list.Count Then
        MsgBox "编号超出范围。", vbExclamation
        Exit Sub
    End If

    gTempBTWFile = CStr(list(idx))

    MsgBox "本次已临时切换模板为：" & gTempBTWFile & vbCrLf & _
           "（打印结束后将自动恢复为默认：" & DEFAULT_BTW_FILE & "）", vbInformation
    Call 一键打印标签
End Sub


'================= 列出同目录所有 .btw 文件 =================
Private Function ListBTWFilesInSameFolder() As Collection
    Dim c As New Collection
    Dim folderPath As String
    folderPath = ThisWorkbook.Path & "\"

    Dim f As String
    f = Dir(folderPath & "*.btw")
    Do While Len(f) > 0
        c.Add f
        f = Dir()
    Loop

    Set ListBTWFilesInSameFolder = c
End Function


'================= 收集待打印行：同出库日期 & 未打印 =================
Private Function CollectRowsToPrint(ws As Worksheet, colShip As Long, colFlag As Long, _
                                   targetShipDate As Variant, ByRef outRows() As Long) As Long
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, colShip).End(xlUp).Row

    Dim tmp() As Long
    ReDim tmp(1 To Application.Max(1, lastRow - 1))
    Dim n As Long: n = 0

    Dim r As Long
    For r = 2 To lastRow
        If SameDate(ws.Cells(r, colShip).Value, targetShipDate) Then
            If Not IsPrintedFlag(ws.Cells(r, colFlag).Value, PRINT_DONE_VALUE) Then
                n = n + 1
                tmp(n) = r
            End If
        End If
    Next r

    If n = 0 Then
        ReDim outRows(1 To 1)
        outRows(1) = 0
        CollectRowsToPrint = 0
        Exit Function
    End If

    ReDim outRows(1 To n)
    For r = 1 To n
        outRows(r) = tmp(r)
    Next r

    CollectRowsToPrint = n
End Function


'================= BarTender 老接口兼容打印（多签名尝试） =================
Private Function PrintFormat_Compat(btFmt As Object, ByRef outErr As String) As Boolean
    On Error Resume Next
    Err.Clear
    outErr = ""

    btFmt.PrintOut
    If Err.Number = 0 Then PrintFormat_Compat = True: Exit Function
    outErr = "PrintOut() 失败：" & Err.Number & " " & Err.Description

    Err.Clear
    btFmt.PrintOut False
    If Err.Number = 0 Then PrintFormat_Compat = True: Exit Function
    outErr = "PrintOut(False) 失败：" & Err.Number & " " & Err.Description

    Err.Clear
    btFmt.PrintOut False, False
    If Err.Number = 0 Then PrintFormat_Compat = True: Exit Function
    outErr = "PrintOut(False,False) 失败：" & Err.Number & " " & Err.Description

    Err.Clear
    btFmt.Print "", False
    If Err.Number = 0 Then PrintFormat_Compat = True: Exit Function
    outErr = "Print("""",False) 失败：" & Err.Number & " " & Err.Description

    Err.Clear
    btFmt.Print "", False, 1
    If Err.Number = 0 Then PrintFormat_Compat = True: Exit Function
    outErr = "Print("""",False,1) 失败：" & Err.Number & " " & Err.Description

    PrintFormat_Compat = False
End Function


'================= 写入 NamedSubStrings（变量名=表头名） =================
Private Sub SetNamedSubStringsFromRow(btFmt As Object, ws As Worksheet, r As Long)
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        Dim key As String
        key = Trim$(CStr(ws.Cells(1, c).Value))
        If Len(key) = 0 Then GoTo NextC

        On Error Resume Next
        btFmt.NamedSubStrings(key).Value = ExcelValueToText(ws.Cells(r, c).Value)
        On Error GoTo 0

NextC:
    Next c
End Sub


'================= 判断“是否打印是否等于是”（更强健） =================
Private Function IsPrintedFlag(cellVal As Variant, doneText As String) As Boolean
    Dim s As String
    s = NormalizeText(cellVal)

    Dim d As String
    d = NormalizeText(doneText)

    If Len(d) > 0 Then
        If InStr(1, s, d, vbTextCompare) > 0 Then
            IsPrintedFlag = True
            Exit Function
        End If
    End If

    If s = "true" Or s = "1" Or s = "y" Or s = "yes" Then
        IsPrintedFlag = True
        Exit Function
    End If

    IsPrintedFlag = False
End Function

Private Function NormalizeText(v As Variant) As String
    Dim s As String

    If IsEmpty(v) Then
        NormalizeText = ""
        Exit Function
    End If

    If IsObject(v) Then
        If v Is Nothing Then
            NormalizeText = ""
            Exit Function
        End If
        s = CStr(v)
    Else
        s = CStr(v)
    End If

    s = Replace(s, ChrW(&H3000), "") ' 全角空格
    s = Replace(s, " ", "")
    s = Replace(s, vbTab, "")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")

    NormalizeText = LCase$(Trim$(s))
End Function


'================= 日期相等（兼容日期/文本） =================
Private Function SameDate(a As Variant, b As Variant) As Boolean
    On Error GoTo TxtCompare
    If IsDate(a) And IsDate(b) Then
        SameDate = (CLng(CDate(a)) = CLng(CDate(b)))
        Exit Function
    End If
TxtCompare:
    SameDate = (Trim$(CStr(a)) = Trim$(CStr(b)))
End Function

Private Function SafeDateText(v As Variant) As String
    If IsDate(v) Then
        SafeDateText = Format$(CDate(v), "yyyy-mm-dd")
    Else
        SafeDateText = Trim$(CStr(v))
    End If
End Function


'================= 表头找列号 =================
Private Function FindHeaderCol(ws As Worksheet, headerName As String) As Long
    Dim lastCol As Long, c As Long
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        If Trim$(CStr(ws.Cells(1, c).Value)) = headerName Then
            FindHeaderCol = c
            Exit Function
        End If
    Next c

    FindHeaderCol = 0
End Function


'================= Excel值转字符串（日期格式统一） =================
Private Function ExcelValueToText(v As Variant) As String
    If IsEmpty(v) Then
        ExcelValueToText = ""
    ElseIf IsDate(v) Then
        ExcelValueToText = Format$(CDate(v), "yyyy-mm-dd")


Public Sub 显示悬浮打印按钮()
    On Error Resume Next
    frmFloatingprint.Show vbModeless   ' vbModeless = 不阻塞Excel，可边滚动边点
End Sub

Public Sub 隐藏悬浮打印按钮()
    On Error Resume Next
    Unload frmFloatingprint
End Sub

