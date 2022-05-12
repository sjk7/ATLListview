VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Stopwatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' clsStopWatch
Private Declare Function timeGetTime Lib "winmm.dll" () As Long
Private Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Long) As Long
Dim m_Info As String
Dim m_Start As Long
Dim m_End As Long
Dim m_Default As Boolean

Friend Property Let Info(s As String)
    m_Info = s
End Property

Public Property Get Now() As Long
    Now = timeGetTime
End Property

Friend Sub StartTiming(Optional doDefault As Boolean = True)
    Static was As Boolean
    m_Default = doDefault
    If Not was Then
        was = True
        timeBeginPeriod 1
    End If
    m_Start = timeGetTime
End Sub

Public Function Create(sInfo As String, Optional doDefault As Boolean = True) As Stopwatch
    Dim c As New Stopwatch
    c.Info = sInfo
    m_Default = False
    Set Create = c
    c.StartTiming

End Function

Private Sub showProgress()
    m_End = timeGetTime
    If (m_Start And LenB(m_Info)) Then
        Form1.Loginfo m_Info + " taken " & m_End - m_Start & " ms, so far ..."
    End If
End Sub

Private Sub doOutput()
    m_End = timeGetTime
    If (m_Start <> 0 And LenB(m_Info) > 0) Then
        Form1.Loginfo m_Info + " took " & m_End - m_Start & " ms."
    End If
End Sub

Private Sub Class_Initialize()
    m_Default = True
End Sub

Private Sub Class_Terminate()
    Debug.Print "terminate when m_Default  = " & m_Default & " info = " & m_Info & vbCrLf
    If m_Default Then
        doOutput
    End If
End Sub
