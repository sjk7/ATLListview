Attribute VB_Name = "modMain"
Option Explicit

Private Declare Function InitShell Lib "shell32" Alias "IsUserAnAdmin" () As Long
Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As Any) As Boolean

Private Type tagINITCOMMONCONTROLSEX
    dwSize  As Long
    dwICC   As Long
End Type

Private Const ICC_STANDARD_CLASSES As Long = &H4000&

Sub Main()
    On Error Resume Next
    
    Dim ICC As tagINITCOMMONCONTROLSEX

    With ICC
        .dwSize = Len(ICC)
        .dwICC = ICC_STANDARD_CLASSES
    End With

    InitShell
    Dim lr As Long
    lr = InitCommonControlsEx(ICC)

    If lr = 0 Or Err.Number <> 0 Then
        InitCommonControls ' 9x version
    End If
    
    Dim f As New Form1
    Form1.Show
End Sub
