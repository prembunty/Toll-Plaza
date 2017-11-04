VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form CustomerForm 
   BackColor       =   &H00C0E0FF&
   Caption         =   "VEHICLE MASTER"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7560
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Customer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5655
   ScaleMode       =   0  'User
   ScaleWidth      =   7560
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   960
      TabIndex        =   30
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox TxtMobile 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   960
      TabIndex        =   21
      Top             =   3600
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   3975
      Left            =   4320
      TabIndex        =   10
      Top             =   360
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   7011
      _Version        =   393216
      Rows            =   100
      FixedCols       =   0
      ForeColorSel    =   0
   End
   Begin VB.ComboBox Combo 
      Appearance      =   0  'Flat
      Height          =   360
      ItemData        =   "Customer.frx":000C
      Left            =   960
      List            =   "Customer.frx":0016
      TabIndex        =   7
      Text            =   "Select"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.TextBox TxtAge 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   960
      TabIndex        =   6
      Top             =   2400
      Width           =   855
   End
   Begin VB.TextBox TxtEmail 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   960
      TabIndex        =   9
      Top             =   3960
      Width           =   3255
   End
   Begin VB.TextBox TxtPh 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   960
      TabIndex        =   8
      Top             =   3240
      Width           =   1695
   End
   Begin VB.TextBox TxtPin 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   3000
      TabIndex        =   5
      Top             =   1950
      Width           =   1215
   End
   Begin VB.TextBox TxtCity 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   960
      TabIndex        =   4
      Top             =   1950
      Width           =   1575
   End
   Begin VB.TextBox TxtAdd3 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   960
      TabIndex        =   3
      Top             =   1590
      Width           =   3255
   End
   Begin VB.TextBox TxtAdd2 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   960
      TabIndex        =   2
      Top             =   1230
      Width           =   3255
   End
   Begin VB.TextBox TxtAdd1 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   960
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox TxtCName 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   3255
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   735
      Left            =   120
      TabIndex        =   22
      Top             =   4680
      Width           =   7215
      Begin VB.CommandButton butList 
         Caption         =   "&List"
         Height          =   375
         Left            =   4200
         TabIndex        =   28
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton butClose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   6240
         TabIndex        =   27
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton butDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2880
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton butUpdate 
         Caption         =   "&Update"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1920
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton butSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   375
         Left            =   960
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton butNew 
         Caption         =   "&New"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender"
      Height          =   300
      Left            =   120
      TabIndex        =   29
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Entered Vehicle Reg.Number"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   3960
      TabIndex        =   20
      Top             =   120
      Width           =   3375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Age"
      Height          =   300
      Left            =   120
      TabIndex        =   19
      Top             =   2440
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   4080
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone "
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   3300
      Width           =   615
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PIN"
      Height          =   255
      Left            =   2640
      TabIndex        =   15
      Top             =   2100
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1000
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Owner"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   690
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Veh.NO"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   375
   End
End
Attribute VB_Name = "CustomerForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub butClose_Click()
Unload Me
End Sub
Private Sub butDelete_Click()
    If RIRS.State = 1 Then RIRS.Close
    RIRS.Open "select * from Customerdb where CNo='" & Combo1 & "'", Conn
    If vbYes = MsgBox("Aru you sure you want to delete this Record", vbYesNo) Then
    If RIRS.EOF = False Then
    Conn.Execute "delete from Customerdb where CNo='" & Combo1 & "'"
    Else
    MsgBox ("Please check CNo")
    End If
    End If
        ClearTxtControls Me, TextBox
        ListDisplay
        butNew.SetFocus
        butDelete.Enabled = False
        butUpdate.Enabled = False
        butList.Enabled = True
        
End Sub

Private Sub butNew_Click()
ClearTxtControls Me, TextBox
Combo1.SetFocus
butNew.Enabled = False
butSave.Enabled = True
End Sub

Private Sub butList_Click()
  ListDisplay
  
End Sub
Sub ListDisplay()
If RIRS.State = 1 Then RIRS.Close
RIRS.Open "Select CNo,CName from Customerdb", Conn
MSFInit
MSF.Row = 1
Do While Not RIRS.EOF
MSF.TextMatrix(MSF.Row, 0) = RIRS(0)
MSF.TextMatrix(MSF.Row, 1) = RIRS(1)
RIRS.MoveNext
MSF.Row = MSF.Row + 1
Loop
End Sub

Private Sub butSave_Click()
If Combo1 = "" Then
MsgBox ("Please enter Correct data")
Exit Sub
End If
    If RIRS.State = 1 Then RIRS.Close
    RIRS.Open "select * from Customerdb  where CNo='" & Combo1 & "'", Conn
    If RIRS.EOF = True Then
    Conn.Execute "insert into Customerdb (CNo,CName,Add1,Add2,Add3,City,Pin,Age,Sex,Phone,Mobile,Email) values('" & Combo1 & "','" & TxtCName & "','" & TxtAdd1 & "','" & TxtAdd2 & "','" & TxtAdd3 & "','" & TxtCity & "'," & Val(TxtPin) & "," & Val(TxtAge) & ",'" & Combo & "','" & TxtPh & "','" & TxtMobile & "','" & TxtEmail & "')"
    MsgBox ("Record inserted")
    Else
    MsgBox ("Please check Vehicle No")
    End If
    ClearTxtControls Me, TextBox
    ListDisplay
    butNew.Enabled = True
    butSave.Enabled = False
    
End Sub

Private Sub butUpdate_Click()
    If RIRS.State = 1 Then RIRS.Close
    RIRS.Open "select * from Customerdb where CNo='" & Combo1 & "'", Conn
    If RIRS.EOF = False Then
    Conn.Execute "update Customerdb set CNo='" & Combo1 & "',CName='" & TxtCName & "',Add1='" & TxtAdd1 & "',Add2='" & TxtAdd2 & "',Add3='" & TxtAdd3 & "',City='" & TxtCity & "',Pin=" & Val(TxtPin) & ",Age=" & Val(TxtAge) & ",Sex='" & Combo & "',Phone='" & TxtPh & "',Mobile=" & Val(TxtMobile) & ",Email='" & TxtEmail & "' where CNo='" & Combo1 & "'"
    MsgBox ("Record Updated")
    Else
    MsgBox ("Please check CNo")
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
Me.Width = 7740
Me.Top = 0
Me.Left = 0
MSFInit
NameDis


If RIRS.State = 1 Then RIRS.Close
RIRS.Open "select vregno from masreg order by vregno", Conn
'Conn.Execute "select regno from masreg order by regno"
'If TempRS.State = 1 Then TempRS.Close
'TempRS.Open "select regno from masreg order by regno", Conn
Combo1.Clear
Do While Not RIRS.EOF
Combo1.AddItem (RIRS(0))
RIRS.MoveNext
Loop


'If RIRS.State = 1 Then RIRS.Close
'RIRS.Open "Select regno from masreg order by regno", Conn
'MSFInit
'Combo1.AddItem (RIRS(0))
'MSF.Row = 1
'Do While Not RIRS.EOF
'Combo1.AddItem (RIRS(0))
'RIRS.MoveNext
'MSF.Row = MSF.Row + 1
'Loop
End Sub
Sub MSFInit()
MSF.Clear
MSF.ColWidth(0) = 500
MSF.ColWidth(1) = 2300
MSF.TextMatrix(0, 0) = "CNo"
MSF.TextMatrix(0, 1) = "CName"

End Sub

Private Sub MSF_DblClick()
        If MSF.TextMatrix(MSF.Row, 0) = "" Then Exit Sub
        CNoVar = MSF.TextMatrix(MSF.Row, 0)
        If RIRS.State = 1 Then RIRS.Close
        RIRS.Open "select * from customerdb where CNo='" & CNoVar & "'", Conn
        If RIRS.EOF = False Then
        Combo1 = RIRS(0) & ""
        TxtCName = RIRS(1) & ""
        TxtAdd1 = RIRS(2) & ""
        TxtAdd2 = RIRS(3) & ""
        TxtAdd3 = RIRS(4) & ""
        TxtCity = RIRS(5) & ""
        TxtPin = RIRS(6) & ""
        TxtAge = RIRS(7) & ""
        Combo = RIRS(8) & ""
        TxtPh = RIRS(9) & ""
        TxtMobile = RIRS(10) & ""
        TxtEmail = RIRS(11) & ""
        End If
  butList.Enabled = False
  butDelete.Enabled = True
  butUpdate.Enabled = True
End Sub

Private Sub TxtCName_LostFocus()
TxtCName = UCase(TxtCName)
End Sub



Private Sub Combo1_LostFocus()
Combo1 = UCase(Combo1)
End Sub


Private Sub TxtPh_KeyPress(KeyAscii As Integer)
Call ValidNumeric(KeyAscii)
End Sub

Private Sub TxtPin_Change()
If Len((TxtPin.Text)) > 6 Then
MsgBox "Limited to 6 Numerics", vbInformation
Call selectTextControl(TxtPin)
Exit Sub
End If
End Sub



Public Sub selectTextControl(txtCtrl As TextBox)
    txtCtrl.SelStart = 0
    txtCtrl.SelLength = Len(txtCtrl.Text)
End Sub

Private Sub TxtMobile_KeyPress(KeyAscii As Integer)
Call ValidNumeric(KeyAscii)
If Len((TxtMobile.Text)) > 9 Then
MsgBox "Limited to 10 Numerics", vbInformation
Call selectTextControl(TxtMobile)
End If
If KeyAscii = 13 Then
TxtEmail.SetFocus
End If
End Sub



Sub NameDis()
If RIRS.State = 1 Then RIRS.Close
RIRS.Open "select vregno from masreg ", Conn
K = 1
Do While Not RIRS.EOF
MSF.TextMatrix(K, 0) = RIRS(0)
RIRS.MoveNext
K = K + 1
Loop
End Sub

