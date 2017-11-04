VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form masregForm 
   BackColor       =   &H00FFFFC0&
   Caption         =   "Vehicle Registration  Entry Form"
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10185
   Icon            =   "MasregForm.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "MasregForm.frx":06C2
   ScaleHeight     =   7320
   ScaleWidth      =   10185
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      MaxLength       =   50
      TabIndex        =   1
      Top             =   3600
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      MaxLength       =   300
      TabIndex        =   0
      Top             =   3000
      Width           =   2655
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   6135
      Left            =   6840
      TabIndex        =   3
      Top             =   240
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   10821
      _Version        =   393216
      Rows            =   25
      Cols            =   3
      FixedCols       =   0
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   5640
      Width           =   6375
      Begin VB.CommandButton butlt 
         Caption         =   "&List"
         Height          =   375
         Left            =   4560
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton butC 
         Caption         =   "Close"
         Height          =   375
         Left            =   5520
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton butDel 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   3600
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton butUp 
         Caption         =   "&Modify"
         Height          =   375
         Left            =   2520
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton butSv 
         Caption         =   "&Save"
         Height          =   375
         Left            =   1320
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton butNw 
         Caption         =   "&New"
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Year Of Model"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   5
      Top             =   3600
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Vehicle Registration No"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   3120
      Width           =   2535
   End
End
Attribute VB_Name = "masregForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub butC_Click()
Unload Me
End Sub

Private Sub butDel_Click()
If RIRS.State = 1 Then RIRS.Close
    RIRS.Open "select * from masreg where vregno='" & Text1 & "'", Conn
    If vbYes = MsgBox("Are you sure you want to delete this Record", vbYesNo) Then
    If RIRS.EOF = False Then
    Conn.Execute "delete from masreg where vregno='" & Text1 & "'"
    Else
    MsgBox ("Please check vregno")
    End If
    End If
        ClearTxtControls Me, TextBox
        ListDisplay
        butNw.SetFocus
        butDel.Enabled = False
        butUp.Enabled = False
        butlt.Enabled = True
End Sub
Sub ListDisplay()
If RIRS.State = 1 Then RIRS.Close
RIRS.Open "Select vregno,yearmodel from masreg", Conn
MSFInit
MSF.Row = 1
Do While Not RIRS.EOF
MSF.TextMatrix(MSF.Row, 0) = RIRS(0)
MSF.TextMatrix(MSF.Row, 1) = RIRS(1)
RIRS.MoveNext
MSF.Row = MSF.Row + 1
Loop
End Sub

Private Sub butlt_Click()
ListDisplay
End Sub

Private Sub butNw_Click()
ClearTxtControls Me, TextBox
Text1.SetFocus
butNw.Enabled = False
butSv.Enabled = True
End Sub

Private Sub butSv_Click()
If Text1 = "" Then
MsgBox ("Please enter Correct data")
Exit Sub
End If
    If RIRS.State = 1 Then RIRS.Close
    RIRS.Open "select * from masreg  where vregno='" & Text1 & "'", Conn
    If RIRS.EOF = True Then
    Conn.Execute "insert into masreg (vregno,yearmodel) values('" & Text1 & "','" & Text2 & "')"
    MsgBox ("Record inserted")
    Else
    MsgBox ("Please check Vehicle No")
    End If
    ClearTxtControls Me, TextBox
    ListDisplay
    butNw.Enabled = True
    butSv.Enabled = False
End Sub
Private Sub butUp_Click()
If RIRS.State = 1 Then RIRS.Close
    RIRS.Open "select * from masreg where vregno='" & Text1 & "'", Conn
    If RIRS.EOF = False Then
    Conn.Execute "update masreg set vregno='" & Text1 & "',yearmodel='" & Text2 & "' where vregno='" & Text1 & "'"
    MsgBox ("Record Updated")
    Else
    MsgBox ("Please check vregno")
    End If
    
        ClearTxtControls Me, TextBox
        ListDisplay
        butNew.SetFocus
        butDelete.Enabled = False
        butUpdate.Enabled = False
        butList.Enabled = True
End Sub

Private Sub Form_Load()
Me.Height = 8000
Me.Width = 10000
Me.Top = 0
Me.Left = 0
MSFInit
End Sub
Sub MSFInit()
MSF.Clear
MSF.ColWidth(0) = 2000
MSF.ColWidth(1) = 2500
MSF.TextMatrix(0, 0) = "vregno"
MSF.TextMatrix(0, 1) = "yearmodel"

End Sub

Private Sub MSF_DblClick()
        If MSF.TextMatrix(MSF.Row, 0) = "" Then Exit Sub
        VregVar = MSF.TextMatrix(MSF.Row, 0)
        If RIRS.State = 1 Then RIRS.Close
        RIRS.Open "select * from masreg where vregno='" & VregVar & "'", Conn
        If RIRS.EOF = False Then
        Text1 = RIRS(0) & ""
        Text2 = RIRS(1) & ""
        End If
  butlt.Enabled = False
  butDel.Enabled = True
  butUp.Enabled = True
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
Call ValidNumeric(KeyAscii)
'If Len((Text2.Text)) > 4 Then
'MsgBox "Limited to 4 Numerics", vbInformation
'Call selectTextControl(TxtMobile)
'End If
'If KeyAscii = 13 Then
'TxtEmail.SetFocus
'End If
End Sub

Private Sub Text2_LostFocus()
Text2 = UCase(Text2)
End Sub



Private Sub Text1_LostFocus()
Text1 = UCase(Text1)
End Sub


'Private Sub TxtPh_KeyPress(KeyAscii As Integer)
'Call ValidNumeric(KeyAscii)
'End Sub

'Private Sub TxtPin_Change()
'If Len((TxtPin.Text)) > 6 Then
'MsgBox "Limited to 6 Numerics", vbInformation
'Call selectTextControl(TxtPin)
'Exit Sub
'End If
'End Sub

Public Sub selectTextControl(txtCtrl As TextBox)
    txtCtrl.SelStart = 0
    txtCtrl.SelLength = Len(txtCtrl.Text)
End Sub

'Private Sub Text2_KeyPress(KeyAscii)
'Call ValidNumeric(KeyAscii)
'If Len((Text2.Text)) > 4 Then
'MsgBox "Limited to 4 Numerics", vbInformation
'Call selectTextControl(TxtMobile)
'End If
'If KeyAscii = 13 Then
'TxtEmail.SetFocus
'End If
'End Sub


