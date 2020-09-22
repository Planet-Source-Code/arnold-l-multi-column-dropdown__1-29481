VERSION 5.00
Object = "*\ADropDownList.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2880
      TabIndex        =   3
      Text            =   "8"
      Top             =   1680
      Width           =   1215
   End
   Begin DropDownList.ALDropDownList ALDropDownList1 
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   450
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Populate DropDown"
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Change No Of Rows"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Enter Number Of Columns"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   1320
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

   Dim cn  As ADODB.Connection
   Dim rs As ADODB.Recordset
   
Private Sub Command1_Click()
  
   Dim sSQLtxt As String

   Set cn = New ADODB.Connection
   Set rs = New ADODB.Recordset
   
  With cn
   .Provider = "Microsoft.Jet.OLEDB.3.51"
   .ConnectionString = App.Path & "\DropDownList.mdb"
   .Open
  End With
  
    sSQLtxt = "SELECT * FROM Calendar"

  With rs
   .Source = sSQLtxt
   .ActiveConnection = cn
   .CursorType = adOpenStatic
   .LockType = adLockReadOnly
   .Open Options:=adCmdText
  End With 'With rs
  
  ALDropDownList1.DropDownValue rs.Clone
   

End Sub 'Private Sub Command1_Click()

Private Sub Command2_Click()
 
 If Not IsNumeric(Text1.Text) Then
   Exit Sub
 End If 'If Not IsNumeric(Text1.Text) Then
 
 ALDropDownList1.RowCount = Text1.Text
 
End Sub 'Private Sub Command2_Click()
