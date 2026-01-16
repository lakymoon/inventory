VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit

    Dim lo As ListObject
    Set lo = Me.ListObjects("DataTable")

    ' 只响应表格数据区的改动
    Dim rng As Range
    Set rng = Intersect(Target, lo.DataBodyRange)
    If rng Is Nothing Then Exit Sub

    Application.EnableEvents = False

    Dim dateCol As Long, specCol As Long
    dateCol = lo.ListColumns("出库日期").Index
    specCol = lo.ListColumns("规格").Index

    ' 只当“规格”列被改动时才触发
    Dim rngSpec As Range
    Set rngSpec = Intersect(Target, lo.DataBodyRange.Columns(specCol))
    If rngSpec Is Nothing Then GoTo SafeExit

    Dim c As Range
    For Each c In rngSpec.Cells
        ' 规格必须是“填写”（非空）才打日期；清空规格不触发
        If Trim$(CStr(c.Value)) <> "" Then
            ' 当前变动所在的“表格行序号”（DataBodyRange 第几行）
            Dim r As Long
            r = c.Row - lo.HeaderRowRange.Row

            If r >= 1 And r <= lo.DataBodyRange.Rows.Count Then
                ' 只有当这行“出库日期为空”时才打日期（避免覆盖旧数据）
                If Trim$(CStr(lo.DataBodyRange.Cells(r, dateCol).Value)) = "" Then
                    lo.DataBodyRange.Cells(r, dateCol).Value = Date
                    lo.DataBodyRange.Cells(r, dateCol).NumberFormat = "yyyy-mm-dd"
                End If
            End If
        End If
    Next c

SafeExit:
    Application.EnableEvents = True
End Sub

