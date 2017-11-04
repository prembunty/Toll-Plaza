VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form LoginAdmin 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Login"
   ClientHeight    =   6225
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   5880
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5655
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   5040
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   3000
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   5040
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   1560
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid MSF 
         Height          =   4335
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   7646
         _Version        =   393216
         Rows            =   25
         FixedCols       =   0
         ScrollBars      =   0
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.CommandButton butExit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Exit"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5520
         Width           =   975
      End
      Begin VB.CommandButton butClose 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5520
         Width           =   975
      End
      Begin VB.CommandButton butSave 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5520
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Administrator Name :"
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
         Left            =   240
         TabIndex        =   9
         Top             =   4800
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Password :"
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
         Left            =   3000
         TabIndex        =   8
         Top             =   4800
         Width           =   1575
      End
   End
End
Attribute VB_Name = "LoginAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public TempRS As New ADODB.Recordset

Private Sub ButClose_Click()
Unload Me
End Sub

Private Sub butExit_Click()
End
End Sub

Private Sub butSave_Click()
For I = 1 To 10
Conn.Execute "update users set username='" & UCase(MSF.TextMatrix(I, 0)) & "', Pword='" & MSF.TextMatrix(I, 1) & "' where rno = " & I
Next
Conn.Execute "update users set username='" & UCase(Text3) & "', Pword='" & Text2 & "' where rno = 0"
End Sub

Private Sub Form_Activate()
Me.Height = 6500
Me.Width = 6000
Me.Left = (MDIForm1.Width - Me.Width) / 2
Me.Top = (MDIForm1.Height - Me.Height) / 2 - 400


MSF.Clear
MSF.ColWidth(0) = 2500
MSF.ColWidth(1) = 2500
MSF.TextMatrix(0, 0) = "User Name "
MSF.TextMatrix(0, 1) = "Password "

If TempRS.State = 1 Then TempRS.Close
TempRS.Open "select * from users where rno > 0", Conn
RowNo = 1
Do While Not TempRS.EOF
MSF.TextMatrix(RowNo, 0) = TempRS(1) & ""
MSF.TextMatrix(RowNo, 1) = TempRS(2) & ""
RowNo = RowNo + 1
TempRS.MoveNext
Loop
If TempRS.State = 1 Then TempRS.Close
TempRS.Open "select * from users where rno = 0", Conn
Text3 = TempRS(1) & ""
Text2 = TempRS(2) & ""

End Sub

Private Sub MSF_Click()
Text1.Height = MSF.RowHeight(MSF.Row)
Text1.Width = MSF.ColWidth(MSF.Col)
Text1.Top = MSF.Top + MSF.CellTop
Text1.Left = MSF.CellLeft + MSF.Left
Text1 = MSF.TextMatrix(MSF.Row, MSF.Col)
Text1.Visible = True
Text1.SelStart = 0
Text1.SelLength = Len(Text1)
Text1.SetFocus
End Sub

Private Sub MSF_LeaveCell()
MSF.TextMatrix(MSF.Row, MSF.Col) = Text1
Text1.Visible = False
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
MSF.TextMatrix(MSF.Row, MSF.Col) = Text1
Text1.Visible = False
End If
End Sub
