VERSION 5.00
Object = "{C73B7301-0087-43FF-97F1-456B1090BF45}#1.0#0"; "ATLListView.dll"
Begin VB.UserControl UserControl1 
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3000
   ScaleWidth      =   4800
   Begin ATLLISTVIEWLibCtl.ListControl lv 
      Height          =   2565
      Left            =   150
      OleObjectBlob   =   "UserControl1.ctx":0000
      TabIndex        =   1
      Top             =   150
      Width           =   4485
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Sub ListControl1_ColumnClick(ByVal whichHeader As ATLLISTVIEWLibCtl.IColumnHeader)

End Sub

Private Sub UserControl_GotFocus()
    lv.SetFocus
End Sub

Private Sub UserControl_Resize()
    lv.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
End Sub

Public Property Get ListView() As ListControl
    Set ListView = lv
End Property
