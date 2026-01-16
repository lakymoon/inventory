VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFloatingprint 
   Caption         =   "打印标签"
   ClientHeight    =   2996
   ClientLeft      =   119
   ClientTop       =   462
   ClientWidth     =   4571
   OleObjectBlob   =   "frmFloatingprint.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "frmFloatingprint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub UserForm_Initialize()
    ' 让窗体更像悬浮工具条
    Me.Caption = "打印"
    Me.StartUpPosition = 0   ' 手动定位
    Me.Top = 20
    Me.Left = Application.UsableWidth - Me.Width - 200
End Sub

Private Sub btnPrint_Click()
    ' 直接调用你已经能工作的主宏
    Call 一键打印标签
End Sub

Private Sub btnClose_Click()
    Call 切换标签模板
End Sub


