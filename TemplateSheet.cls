VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo SafeExit

    ' 只监听你手动会改的区域：
    ' 单价列 D5:D19（你说你手动输入单价）
    ' 如果你还会手改总重量/件数，也可以把 C5:C19 / B5:B19 加进去
    Dim watchRng As Range
    Set watchRng = Union(Me.Range("D5:D19"), Me.Range("C5:C19"), Me.Range("B5:B19"))

    If Intersect(Target, watchRng) Is Nothing Then Exit Sub

    Application.EnableEvents = False

    ' 先让金额列、R2/R3 等公式刷新
    Me.Calculate

    ' 调用标准模块里的刷新显示（会写入 A20/A21/E21/I21）
    RefreshTotalsDisplay Me

SafeExit:
    Application.EnableEvents = True
End Sub

