VERSION 5.00
Object = "{C73B7301-0087-43FF-97F1-456B1090BF45}#1.0#0"; "ATLListView.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{57FBBA6B-40B6-443A-96CF-305EBC7C3802}#184.1#0"; "listviewapi.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   177
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   13680
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.ListView lvMS 
      Height          =   1965
      Left            =   9180
      TabIndex        =   18
      Top             =   4290
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   3466
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   16761024
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin ATLLISTVIEWLibCtl.ListControl lv 
      Height          =   1815
      Left            =   540
      OleObjectBlob   =   "Form1.frx":0000
      TabIndex        =   17
      Top             =   390
      Width           =   4035
   End
   Begin Project1.UserControl1 UserControl11 
      Height          =   1725
      Left            =   9210
      TabIndex        =   8
      Top             =   2400
      Width           =   4035
      _extentx        =   7117
      _extenty        =   3043
   End
   Begin VB.TextBox txt 
      Appearance      =   0  'Flat
      Height          =   2925
      Left            =   390
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   3300
      Width           =   8505
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show Sel"
      Height          =   285
      Left            =   4890
      TabIndex        =   12
      Top             =   30
      Width           =   1125
   End
   Begin VB.CommandButton cmdsel 
      Caption         =   "Show Sel"
      Height          =   285
      Left            =   3420
      TabIndex        =   16
      Top             =   30
      Width           =   1125
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   285
      Left            =   3210
      TabIndex        =   15
      Top             =   2340
      Width           =   1125
   End
   Begin VB.CheckBox chkMultiSelect 
      Caption         =   "MultiSelect"
      Height          =   315
      Left            =   1770
      TabIndex        =   14
      Top             =   0
      Width           =   1575
   End
   Begin VB.CheckBox chkCollate 
      Caption         =   "Collate Sort"
      Height          =   315
      Left            =   90
      TabIndex        =   13
      Top             =   0
      Value           =   1  'Checked
      Width           =   1575
   End
   Begin VB.CommandButton cmdAddSomelvOnly 
      Caption         =   "me only"
      Height          =   285
      Left            =   330
      TabIndex        =   11
      Top             =   2340
      Width           =   1125
   End
   Begin ListViewAPI.ListView lvapi 
      Height          =   1815
      Left            =   9180
      TabIndex        =   10
      Top             =   390
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   3201
   End
   Begin VB.CommandButton cmdRange 
      Caption         =   "Add Range"
      Height          =   375
      Left            =   2670
      TabIndex        =   9
      Top             =   6480
      Width           =   1605
   End
   Begin VB.CommandButton cmdEnum 
      Caption         =   "Enum"
      Height          =   375
      Left            =   420
      TabIndex        =   7
      Top             =   6480
      Width           =   1605
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Top             =   2700
      Width           =   1815
   End
   Begin VB.CheckBox chkOptim 
      Caption         =   "Optimized Add"
      Height          =   285
      Left            =   4770
      TabIndex        =   4
      Top             =   2310
      Width           =   1875
   End
   Begin VB.CommandButton cmdAddSome 
      Caption         =   "AddSome"
      Height          =   405
      Left            =   4770
      TabIndex        =   3
      Top             =   2670
      Width           =   1515
   End
   Begin VB.TextBox txtHowMany 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   6
      Text            =   "20000"
      Top             =   2670
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   2580
      TabIndex        =   2
      Top             =   2670
      Width           =   1815
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   1815
      Left            =   4800
      TabIndex        =   0
      Top             =   390
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   3201
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16761087
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   22080
      Y1              =   6390
      Y2              =   6390
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "How many"
      Height          =   315
      Left            =   6510
      TabIndex        =   5
      Top             =   2730
      Width           =   1005
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const WM_SETREDRAW = &HB
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const SB_BOTTOM = 7
Private Const EM_SCROLL As Integer = &HB5
Private WithEvents m_ColumnHeaders As ATLLISTVIEWLibCtl.ColumnHeaders
Attribute m_ColumnHeaders.VB_VarHelpID = -1


Private Sub CheckIndexes(lv As Object, litems As Variant)
    Dim li
    Dim i  As Long
    i = 1
    Dim sw As Stopwatch
    Set sw = New Stopwatch
    Set sw = sw.Create("Checking index sanity for: " & TypeName(lv))
    
    With litems
    Do While (i < .Count)
        Set li = litems(i)
        Debug.Assert (li.Index = i)
        i = i + 1
    Loop

    End With

End Sub
Private Sub lvActions(lv As Object)
    Dim col
    
    If TypeName(lv) = "ListControl" And lv.ListItems.Count = 0 Then
        Set col = lv.ColumnHeaders.Add(, "WTF", "Static String")
        Debug.Assert col.Text = "Static String"
    End If
    
   If TypeName(lv) = "ListControl" And lv.Name = "lv" Then
        Set m_ColumnHeaders = Me.lv.ColumnHeaders
    End If
    
    Dim s As String
    s = "Column" & lv.ColumnHeaders.Count + 1
    Set col = lv.ColumnHeaders.Add(, , s)
    'Exit Sub
    Debug.Print lv.hwnd
    Dim c
    Set c = lv.ColumnHeaders
    Debug.Print c.Count
    Debug.Assert (col.Text = s)
    Set col = c.Item(1)
    col.Width = 1440
    Debug.Print col.Text
    Dim Color As OLE_COLOR
    Color = RGB(255, 255, 0)
    lv.BackColor = Color
    Dim actual As Long
    actual = lv.BackColor
    Debug.Assert actual = Color
    Static times_clicked As Long
    times_clicked = times_clicked + 1
    Dim li
    Dim litems
    Set litems = lv.ListItems
    Set li = litems.Add(, , "Listitem " & litems.Count + 1)
    If TypeName(lv) = "ListControl" Then
        lv.ColumnHeaders.HeightInPixels = lv.ColumnHeaders.HeightInPixels + 1
        ' Set m_ColumnHeaders = lv.ColumnHeaders
    End If


End Sub

Sub DoColumnEnum(lv As Object)
    Dim colHeaders
    Set colHeaders = lv.ColumnHeaders
    Dim colHeader
    
    For Each colHeader In colHeaders
        Debug.Print colHeader.Text
        Debug.Print colHeader.Index
    Next
    
End Sub

Private Sub chkMultiSelect_Click()
    If chkMultiSelect.Value = vbChecked Then
        lvw.MultiSelect = True
        lv.MultiSelect = True
    Else
        lvw.MultiSelect = False
        lv.MultiSelect = False
    End If
End Sub

Private Sub cmdAddSomelvOnly_Click()

    Dim howMany As Long
    If IsNumeric(txtHowMany.Text) Then
        howMany = txtHowMany.Text
    Else
        howMany = 50000
    End If
    Dim lck As New ATLLISTVIEWLibCtl.RedrawLock
    Set lck.ListControlObject = lv

    Call AddLitemsEx(howMany, lv)
    CheckIndexes lv, lv.ListItems
    Loginfo "There are " & lv.ListItems.Count & " items in the ListControl"
    Loginfo "Virtual mode is: " & lv.VirtualMode
    Loginfo vbNullString
    
    
    
End Sub

Private Sub cmdClear_Click()
    ClearLvw lvw
    ClearLvw lv
    ClearLvw lvapi
End Sub

Private Sub ClearLvw(lv As Object)
    Dim sw As New Stopwatch
    Set sw = sw.Create("Clearing " & lv.ListItems.Count & " items, for " & _
        TypeName(lv) & " items took: ")
    
    lv.ListItems.Clear
    Debug.Assert lv.ListItems.Count = 0
End Sub
Private Sub ShowSelLvwApi()
    Dim sel As Collection
    Dim sw As New Stopwatch
    Set sw = sw.Create("Creating Selected items (for DPS' lvw) took: ")
    Set sel = lvapi.SelectedItems
    Loginfo "LvAPI has " & sel.Count & " selected items."
    
    Set sw = Nothing
End Sub
Private Sub ShowSel()
    Dim sel As ATLLISTVIEWLibCtl.SelItemCollection
    Dim sw As New Stopwatch
    Set sw = sw.Create("Creating Selected items for ATL ListControl took: ")
    Set sel = lv.SelectedItems
    Loginfo "Listcontrol has " & sel.Count & " selected items."
    
    Set sw = Nothing
End Sub

Private Sub ShowSelLvw()
    Dim sel As Collection
    Dim sw As New Stopwatch
    Set sw = sw.Create("Creating (MSLV) Selected items took: ")
    Set sel = New Collection
    Dim litem As MSComctlLib.ListItem
    
    For Each litem In lvw.ListItems
        If litem.Selected Then
            sel.Add litem
        End If
    Next litem
    
    Loginfo "Listview has " & sel.Count & " selected items."
    
    Set sw = Nothing
End Sub

Private Sub cmdEnum_Click()
    DoColumnEnum lv
    DoColumnEnum lvw
    
    Dim l1 As Currency, l2 As Currency
    l1 = DoEnum(lv)
    l2 = DoEnum(lvw)
    'Debug.Assert (l1 = l2)
    Loginfo vbNullString
    l1 = DoEnum(lvapi)
    
End Sub

Private Sub cmdRange_Click()
    Dim ln As Long: ln = CLng(Me.txtHowMany)
    AddRange2 lv, ln
    AddRange lv, ln
    AddRange lvapi, ln
    AddRange lvw, ln
    
End Sub

Private Sub AddRange2(lview As ListControl, nItems As Long)
    Dim c As New Stopwatch
    Set c = c.Create("AddRange (api compat) on new listview")
    
    Dim colAdded As InterfaceCollection
    Dim idx As Long
    idx = lview.ListItems.Count - 1
    
    Set colAdded = lview.ListItems.AddRange(nItems)
    Debug.Assert colAdded.Count = nItems
    
    Dim litem As ATLLISTVIEWLibCtl.ListItem
    For Each litem In colAdded
        litem.Text = "Hello Listitem! @ " & idx + 1
        idx = idx + 1
    Next
    Set c = Nothing
End Sub

Private Sub AddRange(lv As Object, nItems As Long)
    If lv.Name = "lvw" Then
        Loginfo "MS Listview does not have an addrange method, so adding them ..."
        AddLitemsEx nItems, lv
        Exit Sub
    End If

    Dim ar() As String
    ReDim ar(nItems - 1)
    Dim i As Long
    Dim c As New Stopwatch
    SendMessage lv.hwnd, WM_SETREDRAW, 0, 0
    Set c = c.Create("Adding Range, for " & lv.Name & "(" & TypeName(lv) & ")" & " with " & nItems & " items")
    For i = 0 To nItems - 1
        ar(i) = "Item in array " & i + 1
    Next i
    If lv.Name = "lv" Then
        lv.ListItems.AddRangeEx ar
    Else
        lv.ListItems.AddRange nItems
    End If
    SendMessage lv.hwnd, WM_SETREDRAW, 1, 0

End Sub

Private Sub cmdsel_Click()
    ShowSel
End Sub

Private Sub cmdsel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        lv.StartLabelEditEx 0, 1
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
    If Index = 0 Then

        lvActions lv
    Else
        lvActions lvw
        lvActions lvapi
        lvActions lvMS
    End If
    
    
    lvActions UserControl11.ListView
    Exit Sub

End Sub


Private Sub AddLitemsEx(howMany As Long, lv As Object)

    Dim c As Stopwatch
    Set c = New Stopwatch
    Set c = c.Create("Adding " & howMany & " items to: " & lv.Name & " (" & TypeName(lv) & ")")
    Dim litems, li
    Set litems = lv.ListItems
     Dim oldPointer As Integer: oldPointer = Me.MousePointer
    Me.MousePointer = vbHourglass
    Me.Refresh
    
    'SendMessage lv.hwnd, WM_SETREDRAW, 0, 0
    Dim i As Long
    Dim t As Long
    t = c.Now
    
    If lv.Name = "lvapi" Then
        lv.SuspendLayout
    End If
    Dim isListControl As Boolean
    isListControl = TypeName(lv) = "ListControl"
    Dim ctr As Long
    ctr = litems.Count
    Dim colCount As Long
    colCount = lv.ColumnHeaders.Count
    'Debug.Assert ObjPtr(lv) <> ObjPtr(lvMS)
    
    With litems
        For i = 0 To howMany
            Set li = .Add(, , "Hello Listitem! @ " & i + 1)
            If (isListControl) Then
                
                Dim myCtr As Long
                myCtr = 1
                Do While (myCtr < colCount)
                    Dim lsub As ATLLISTVIEWLibCtl.ListSubItem
                    Set lsub = li.ListSubItems(myCtr)
                    lsub.Text = "Subitem " & myCtr & ",for item:" & ctr
                    myCtr = myCtr + 1
                Loop
             Else
                myCtr = 1
                Do While (myCtr < colCount)
                    li.SubItems(myCtr) = "Subitem " & myCtr & ",for item:" & ctr
                    myCtr = myCtr + 1
                Loop
            End If
            
            If t - c.Now > 250 Then
                t = c.Now
                DoEvents
            End If
            ctr = ctr + 1
        Next i
    End With

    If lv.Name = "lvapi" Then
        lv.ResumeLayout
    End If
    'SendMessage lv.hwnd, WM_SETREDRAW, 1, 0
   
    Me.MousePointer = oldPointer

End Sub
Private Sub CmdAddSome_Click()

    Dim howMany As Long
    If IsNumeric(txtHowMany.Text) Then
        howMany = txtHowMany.Text
    Else
        howMany = 50000
    End If
    
    Dim lck As New ATLLISTVIEWLibCtl.RedrawLock
    Set lck.ListControlObject = lv
    Call AddLitemsEx(howMany, lvw)
    Debug.Print lck.ListControlObject.Appearance
    CheckIndexes lvw, lvw.ListItems
    
    Call AddLitemsEx(howMany, lv)
    CheckIndexes lv, lv.ListItems
    lv.SetRedraw True
   
    Call AddLitemsEx(howMany, lvapi)
    
    Loginfo vbNullString
    
    Call AddLitemsEx(howMany, lvMS)
    Set lck = Nothing
End Sub

Private Sub Command2_Click()
    ShowSelLvw
    ShowSelLvwApi
End Sub



Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        lv.StartLabelEdit
    End If
End Sub

Private Sub Form_Load()
    Debug.Assert lv.VirtualMode = False
    lv.VirtualMode = True
    lvapi.ColumnHeaders(1).Width = 1440
    Call lvMS.ColumnHeaders.Add(, , "Compare against me!")
    lvapi.BorderStyle = 1
End Sub

Friend Sub Loginfo(s As String)
    If LenB(s) = 0 Then
        txt.Text = txt.Text & vbCrLf
        Debug.Print vbCrLf
    Else
        txt.Text = txt.Text & Now & vbTab & s & vbNewLine
       Debug.Print s
    End If
    SendMessageLong txt.hwnd, EM_SCROLL, SB_BOTTOM, 0
End Sub


Function DoEnum(lv As Object) As Currency
    Dim litems
    Set litems = lv.ListItems
    Dim litem
    Dim idxs As Currency
    Dim c As New Stopwatch
    Set c = c.Create("Iterating in lv " & lv.Name & ", count = " & litems.Count & " items")
again:
    For Each litem In litems
        idxs = idxs + litem.Index
    Next litem
    
    DoEnum = idxs
    Set c = Nothing
    
End Function



Private Sub lv_AfterLabelEdit(Cancel As Boolean, NewString As String)
    Loginfo "AfterLabedEdit(): Cancel is " & Cancel & ", NewString is: " & NewString
    If CLng(Timer) Mod 2 = 0 Then
        Cancel = True
    End If
End Sub

Private Sub lv_BeforeLabelEdit(Cancel As Integer)
    'Cancel = True
    Loginfo "BeforeLabelEdit: Cancel = " & Cancel
End Sub

Private Sub lv_Click()
    Loginfo "ListControl Click"
End Sub

Private Sub lv_ClickEx(ByVal X As Long, ByVal Y As Long)
    Loginfo "ListControl ClickEx, with x and y: " & X & "," & Y
End Sub

Private Sub lv_ColumnClick(ByVal ColumnHeader As ATLLISTVIEWLibCtl.IColumnHeader)
    Debug.Print "Colheader clicked: " & ColumnHeader.Index & " " & ColumnHeader.Text & " width: " & ColumnHeader.Width
    Dim c As New Stopwatch
    Set c = c.Create("ListControl: Sorting " & lv.ListItems.Count & " items took")
    Dim ord As ATLLISTVIEWLibCtl.ListSortOrderFlags
    ord = lv.SortOrder
    If (ord And lvwDescending) Then
        ord = lvwAscending Or lvwNatural
    Else
        ord = lvwDescending Or lvwNatural
    End If
    
    If chkCollate.Value = vbChecked Then
        ord = ord Or lvwNatural
    Else
        ord = ord Xor lvwNatural
    End If
    
    lv.SortOrder = ord
    lv.SortKey = 0
    lv.Sorted = True
    
End Sub

Private Sub lv_ColumnRemoved(ByVal colRemoved As ATLLISTVIEWLibCtl.ColumnHeader)
    Debug.Print "Column Removed: " & colRemoved.Index
End Sub

'                               (DoDefault As Boolean, Shift As Integer, x As Single, y As Single, ColumnHeader As ListViewAPI.ColumnHeader)


Private Sub lv_ColumnRightClick(doDefault As Boolean, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, ByVal ColumnHeader As ATLLISTVIEWLibCtl.IColumnHeader)

    If ObjPtr(ColumnHeader) Then
        Debug.Assert (doDefault = True)
        Debug.Print "Right click on columnheader: " & ColumnHeader.Text & ", at index: " & ColumnHeader.Index
    
        If ColumnHeader.Index Mod 2 = 0 Then
            doDefault = False
        End If
    Else
        Loginfo "Clicked on the header, but not on any column (probably no columns to show)"
    End If
End Sub

Private Sub lv_DblClick()
    Loginfo "ListControl " & lv.Name & " double clicked."

 
End Sub

Private Sub lv_GotFocus()
    Loginfo "ListControl " & lv.Name & " Got focus."
End Sub

Private Sub lv_ItemClick(ByVal Item As ATLLISTVIEWLibCtl.ListItem)
    Loginfo "Item " & Item.Index & " clicked"
End Sub

Private Sub lv_ItemClicked(ByVal Item As ATLLISTVIEWLibCtl.ListItem, ByVal Button As ATLLISTVIEWLibCtl.vbMouseButtonConstants)
    Loginfo "Item " & Item.Index & " clicked, with button = " & Button
End Sub

Private Sub lv_ItemClickEx(ByVal Item As ATLLISTVIEWLibCtl.ListItem, ByVal SubItemIndex As Integer)
    Loginfo "Item " & Item.Index & " clicked(Ex) " & "Subitem Index is: " & SubItemIndex
   
End Sub

Private Sub lv_ItemClickRight(ByVal Item As ATLLISTVIEWLibCtl.ListItem, ByVal SubItemIndex As Integer)
    Loginfo "Item " & Item.Index & " RIGHT clicked(Ex) " & "Subitem Index is: " & SubItemIndex
        
    lv.StartLabelEditEx Item.Index, SubItemIndex
End Sub

Private Sub lv_ItemSelectionChanged(ByVal Item As ATLLISTVIEWLibCtl.IListItem, ByVal SelState As Boolean)
    Exit Sub
    If lv.LayoutSuspended Then Exit Sub
    Debug.Print "Selection changed, for item: " & Item.Index
    Debug.Print "New state is "
    If (SelState) Then
       Debug.Print "SELECTED"
    Else
        Debug.Print "NOT SELECTED"
    End If

    
    Dim litem As ATLLISTVIEWLibCtl.ListItem
    For Each litem In lv.SelectedItems
        Loginfo "Litem " & litem.Index & " is selected."
    Next litem
    
    
End Sub
Private Sub checkItemCount()
    Debug.Assert lv.SelectedItemCount = lv.SelectedItems.Count
    Debug.Assert lv.SelectedItems.Count = lv.SelectedItemCount
End Sub

Private Sub lv_KeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer)
    Dim s As String
    s = "Key down on Listcontrol: " & Chr$(KeyCode)
    Loginfo s
    PrintShiftKeys Shift, s
    
End Sub

Private Sub lv_KeyPress(KeyAscii As Integer)
    Dim s As String
    s = "KeyPress on ListControl: " & Chr$(KeyAscii)
    Loginfo s
    PrintShiftKeys 0, s
    ' KeyAscii = 0
End Sub

Private Sub lv_KeyUp(ByVal KeyCode As Integer, ByVal Shift As Integer)
    Dim s As String
    s = "Key up on Listcontrol: " & Chr$(KeyCode)
    Loginfo s
    PrintShiftKeys Shift, s
    
End Sub

Private Sub lv_LostFocus()
    Loginfo "ListControl " & lv.Name & " Lost focus."
End Sub

Private Sub lv_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    PrintMouse Button, "MouseDown", Shift, X, Y
End Sub

Private Sub lv_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
   PrintMouse Button, "MouseMove", Shift, X, Y
   
   Dim li As ATLLISTVIEWLibCtl.ListItem
   Dim liEx As ATLLISTVIEWLibCtl.ListItem
   
    Dim SubItemIndex As Long
    If Button = vbLeftButton Then
        Set li = lv.HitTest(X, Y)
        Set liEx = lv.HitTestEx(X, Y, SubItemIndex)
        If ObjPtr(li) Then
            Loginfo "HitTest: ListItem Index: " & li.Index
            Debug.Assert ObjPtr(liEx)
            Loginfo "HitTestEx: Listitem Index: " & li.Index & " subitem: " & SubItemIndex
        Else
            Debug.Assert liEx Is Nothing
        End If
    End If
End Sub

Private Sub PrintShiftKeys(Shift As Integer, action As String)
    If (Shift And vbShiftMask) Then
        Loginfo "SHIFT KEY IS DOWN WHEN " & action
    End If
    If (Shift And vbAltMask) Then
        Loginfo "ALT KEY IS DOWN WHEN " & action
    End If
    If (Shift And vbCtrlMask) Then
        Loginfo "CTRL KEY IS DOWN WHEN " & action
    End If
End Sub

Private Sub PrintMouse(Button As Integer, action As String, Shift As Integer, X As Single, Y As Single)
   Debug.Print action & ": " & X & ":" & Y
    PrintShiftKeys Shift, action
    'Debug.Assert action <> "MouseUp"
    If Button And vbRightButton Then
        Debug.Print "RIGHT MOUSE DOWN WHEN " & action
    End If
    
    If Button And vbLeftButton Then
        Debug.Print "LEFT MOUSE DOWN WHEN " & action
    End If
End Sub


Private Sub lv_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    PrintMouse Button, "MouseUp", Shift, X, Y

End Sub

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Debug.Print "Colheader clicked: " & ColumnHeader.Index & " " & ColumnHeader.Text & " width: " & ColumnHeader.Width
    Dim c As New Stopwatch
    Set c = c.Create("MS Listview: Sorting " & lvw.ListItems.Count & " items took")
    Dim ord As ListSortOrderConstants
    ord = lvw.SortOrder
    If (ord = MSComctlLib.lvwDescending) Then
        ord = MSComctlLib.lvwAscending
    Else
        ord = MSComctlLib.lvwDescending
    End If
    
    lvw.SortOrder = ord
    lvw.SortKey = 0
    lvw.Sorted = True
End Sub

Private Sub lvw_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim s As String
    s = "KeyDown on MS Listview: " & Chr$(KeyCode)
    Loginfo s
    PrintShiftKeys Shift, s
End Sub

Private Sub lvw_KeyPress(KeyAscii As Integer)
    Dim s As String
    s = "KeyPress on MS Listview: " & Chr$(KeyAscii)
    Loginfo s
    PrintShiftKeys 0, s
    KeyAscii = 0
End Sub

Private Sub lvw_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim s As String
    s = "KeyUp on MS Listview: " & Chr$(KeyCode)
    Loginfo s
    PrintShiftKeys Shift, s
End Sub

Private Sub m_ColumnHeaders_ColumnClick(ByVal whichHeader As ATLLISTVIEWLibCtl.IColumnHeader)
    If ObjPtr(whichHeader) Then
        Debug.Print "Clicked on Column " & whichHeader.Index
    End If
End Sub

Private Sub m_ColumnHeaders_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, _
                                        ByVal X As Single, ByVal Y As Single)
    Loginfo "ColumnHeaders_MouseDown: " & Button & ":" & Shift & ":" & X & ":" & Y
End Sub

Private Sub m_ColumnHeaders_MouseEvent(ByVal iMsg As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single, doDefault As Boolean)
    Debug.Print iMsg
    Debug.Print doDefault
    
End Sub

Private Sub m_ColumnHeaders_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    
    Loginfo "Columnheaders mousemove " & X & ":" & Y
    Dim chdr As ColumnHeader
    Set chdr = m_ColumnHeaders.HitTest(X, Y)
    If ObjPtr(chdr) = 0 Then
    Else
        Loginfo "ColumnHeader hitTest: " & chdr.Index
    End If
End Sub

Private Sub txt_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim s As String
    s = "Key down on Text " & Chr$(KeyCode)
    Loginfo s
    PrintShiftKeys Shift, s
End Sub

Private Sub txt_KeyPress(KeyAscii As Integer)
    Dim s As String
    s = "KeyPress on ListControl: " & Chr$(KeyAscii)
    Loginfo s
    PrintShiftKeys 0, s
    'KeyAscii = 0
End Sub

Private Sub txt_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim s As String
    s = "Key up on Text: " & Chr$(KeyCode)
    Loginfo s
    PrintShiftKeys Shift, s
End Sub
