VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl ALDropDownList 
   ClientHeight    =   270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2670
   FillStyle       =   0  'Solid
   Picture         =   "DropDownList.ctx":0000
   ScaleHeight     =   270
   ScaleWidth      =   2670
   ToolboxBitmap   =   "DropDownList.ctx":02CE
   Begin MSComctlLib.ListView lvwDropDown 
      Height          =   2295
      Left            =   0
      TabIndex        =   3
      Top             =   285
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   4048
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.PictureBox picButton 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawStyle       =   5  'Transparent
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   230
      Left            =   2400
      Picture         =   "DropDownList.ctx":05E0
      ScaleHeight     =   4.084
      ScaleMode       =   0  'User
      ScaleTop        =   30
      ScaleWidth      =   4.43
      TabIndex        =   1
      Top             =   30
      Width           =   230
   End
   Begin VB.TextBox txtDropDown 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
   Begin VB.TextBox txtCharLen 
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "ALDropDownList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

   Private objDropDownValue As ADODB.Recordset
   Private strDropDownState As String
   Private strResizeWhen As String
   Private intRowCount As Integer
   
'Events
   Public Event Change()
   Public Event Click()
   Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Public Event ListDropDown(vDropDownState As String)
   Public Event Populated(vPopulated As Boolean)
   Public Event Resize()
   Public Event ColumnCountChange()

Private Sub lvwDropDown_Click()
  
  txtDropDown.Text = lvwDropDown.SelectedItem
  txtDropDown.SelStart = Len(txtDropDown.Text)
  Call DropDown("")
  RaiseEvent Click
  
End Sub 'Private Sub lvwDropDown_Click()

Private Sub picButton_Click()
 
 Call DropDown("")
 
End Sub 'Private Sub picButton_Click()

Private Sub txtDropDown_Change()

   Dim lvwColumSearch As ListItem
  
   Set lvwColumSearch = lvwDropDown.FindItem(txtDropDown.Text, , , lvwPartial)
   
  If lvwColumSearch Is Nothing Then
    Call DropDown("Closed")
    Exit Sub
   Else
    lvwColumSearch.EnsureVisible
    lvwColumSearch.Selected = True
     Set lvwColumSearch = Nothing
    Call DropDown("DropedDown")
  End If 'If lvwColumSearch Is Nothing Then
  
  RaiseEvent Change

End Sub 'Private Sub txtDropDown_Change()

Private Sub lvwDropDown_KeyDown(KeyCode As Integer, Shift As Integer)
  
 If KeyCode = vbEnter Or KeyCode = 13 Then
   txtDropDown.Text = lvwDropDown.SelectedItem
  ElseIf (KeyCode >= 65 And KeyCode <= 90) Or (KeyCode >= 97 And KeyCode <= 122) Then
   txtDropDown.Text = txtDropDown.Text & Chr(KeyCode)
   txtDropDown.SelStart = Len(txtDropDown.Text)
 End If 'If KeyCode = vbEnter Then
 
End Sub 'Private Sub txtDropDown_KeyDown(KeyCode As Integer, Shift As Integer)

Private Sub UserControl_Initialize()

  strDropDownState = "Closed"
  
End Sub 'Private Sub UserControl_Initialize()

Public Sub DropDownValue(objData As ADODB.Recordset)
  
  Dim i As Integer
  Dim iLenHead As Integer
  Dim iLenData As Integer
  Dim iTotalWith As Integer
  Dim sColumnWith As String
  Dim lvwColumn As ListItem
      
  Set objDropDownValue = objData.Clone(adLockReadOnly)
  
 Screen.MousePointer = vbHourglass
  
 objData.Close
  
 lvwDropDown.ListItems.Clear
 lvwDropDown.ColumnHeaders.Clear
  
 objDropDownValue.MoveFirst
 
 iTotalWith = 0
 
 For i = 0 To objDropDownValue.Fields.Count - 1
   
   iLenHead = Len(objDropDownValue.Fields(i).Name)
   iLenData = Len(objDropDownValue.Fields(i).Value)
   
   If iLenHead >= iLenData Then
     txtCharLen.Text = String(iLenHead, "A")
    Else
     txtCharLen.Text = String(iLenData, "A")
   End If 'If iLenHead >= iLenData Then
   
   iTotalWith = iTotalWith + TextWidth(txtCharLen)
   
   lvwDropDown.ColumnHeaders.Add , , objDropDownValue.Fields(i).Name, TextWidth(txtCharLen)
 Next i 'For i = 0 To objDropDownValue.Fields.Count - 1
 
 Do Until objDropDownValue.EOF
    Set lvwColumn = lvwDropDown.ListItems.Add(, , objDropDownValue.Fields(0))
  For i = 1 To objDropDownValue.Fields.Count - 1
   lvwColumn.SubItems(i) = objDropDownValue.Fields(i).Value
  Next i
  objDropDownValue.MoveNext
 Loop 'Do Until objDropDownValue.EOF
 
 lvwDropDown.Width = iTotalWith + 190
 
 If UserControl.Width < lvwDropDown.Width Then
   strResizeWhen = "Runtime"
   UserControl.Width = lvwDropDown.Width + 190
 End If 'If UserControl.Width < lvwDropDown.Width Then
 
 objDropDownValue.Close
  
  Set objDropDownValue = Nothing
  
  Set lvwColumn = Nothing
 
 Screen.MousePointer = vbDefault
 
 RaiseEvent Populated(True)
 
End Sub 'Public Property Let DropDownValue(objData As Object)

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  RaiseEvent MouseDown(Button, Shift, X, Y)
 
End Sub 'Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  RaiseEvent MouseMove(Button, Shift, X, Y)
  
End Sub 'Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  
  RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub 'Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Sub UserControl_Resize()
  
 If strResizeWhen = "" Then
   txtDropDown.Width = UserControl.Width
   picButton.Left = UserControl.Width - picButton.Width
   
   RaiseEvent Resize
 End If 'If strResizeWhen = "" Then
 
 strResizeWhen = ""
  
End Sub 'Private Sub UserControl_Resize()

Private Sub DropDown(ByVal strToBeState As String)

  If strToBeState = "DropedDown" Then
    lvwDropDown.Visible = True
    strResizeWhen = "Runtime"
    UserControl.Height = lvwDropDown.Height + txtDropDown.Height
    strDropDownState = "DropedDown"
   ElseIf strToBeState = "Closed" Then
     lvwDropDown.Visible = False
     strResizeWhen = "Runtime"
     UserControl.Height = txtDropDown.Height
     strDropDownState = "Closed"
    Else
     If strDropDownState = "Closed" Then
      lvwDropDown.Visible = True
      strResizeWhen = "Runtime"
      UserControl.Height = lvwDropDown.Height + txtDropDown.Height
      strDropDownState = "DropedDown"
     ElseIf strDropDownState = "DropedDown" Then
      lvwDropDown.Visible = False
      strResizeWhen = "Runtime"
      UserControl.Height = txtDropDown.Height
      strDropDownState = "Closed"
    End If 'If strDropDownState = "Closed" Then
  End If 'If strToBeState = "DropedDown" Then
  
  RaiseEvent ListDropDown(strDropDownState)

End Sub 'Private Sub DropDown()

Public Property Get DropDownState() As String
 
 DropDownState = strDropDownState
 
End Property 'Public Property Get DropDownState() As String

Public Property Let DropDownState(ByVal vDropDownState As String)

  strDropDownState = vDropDownState
  Call DropDown(strDropDownState)
  
End Property 'Public Property Let DropDownState(ByVal vDropDownState As String)

Public Property Let RowCount(ByVal vRowCount As Integer)
  
   Dim lvwColumn As ListItem
  
  intRowCount = vRowCount
  
  If intRowCount > 0 Then
    If lvwDropDown.ListItems.Count > 0 Then
      lvwDropDown.Height = lvwDropDown.ListItems.Item(1).Height * (intRowCount + 1.5)
     Else
       Set lvwColumn = lvwDropDown.ListItems.Add(, , "test")
      lvwDropDown.Height = lvwColumn.Height * (intRowCount + 1.5)
       Set lvwColumn = lvwDropDown.FindItem("test", , , lvwPartial)
      lvwColumn.EnsureVisible
      lvwColumn.Selected = True
      lvwDropDown.ListItems.Remove (lvwColumn.Index)
       Set lvwColumn = Nothing
    End If 'If lvwDropDown.ListItems.Count > 0 Then
  End If 'If intRowCount > 0 Then
  
  RaiseEvent ColumnCountChange
      
End Property 'Public Property Let RowCount(ByVal vRowCount As String)

Public Property Get Text() As String
  
  Text = txtDropDown.Text
  
End Property 'Public Property Get Text() As String

Public Property Let Text(ByVal vText As String)
  
  txtDropDown.Text = vText
  
End Property 'Public Property Let Text(ByVal vText As String)
