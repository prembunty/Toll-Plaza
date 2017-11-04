VERSION 5.00
Begin VB.Form Login 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Login"
   ClientHeight    =   2310
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4875
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4815
      Begin VB.CommandButton butAdmin 
         BackColor       =   &H00FFFFC0&
         Caption         =   "&Admin"
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
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton butCancel 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Cancel"
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
         TabIndex        =   6
         Top             =   1680
         Width           =   975
      End
      Begin VB.CommandButton butLogin 
         BackColor       =   &H00C0FFC0&
         Caption         =   "&Login"
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
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1680
         Width           =   975
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
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox Text1 
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
         Left            =   1920
         TabIndex        =   2
         Top             =   360
         Width           =   2415
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
         Left            =   360
         TabIndex        =   4
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name :"
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
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LNo As Byte
'Public TempRS As New ADODB.Recordset

Private Sub butAdmin_Click()
LNo = 2
If TempRS.State = 1 Then TempRS.Close
TempRS.Open "select * from users where username='" & UCase(Text1) & "' and pword = '" & Text2 & "'", Conn
'TempRS.Open "select * from users where UserName = 'ADMIN' and pword = '" & Text2 & "' and rno=0", Conn
If TempRS.EOF = True Then
MsgBox ("The entered UserName or Password is not Correct")
Text1.SetFocus
LNo = 1
Else

UserNameVar = "ADMIN"
LoginAdmin.Show
End If

End Sub

Private Sub butCancel_Click()
LNo = 1
End
End Sub

Private Sub butLogin_Click()
LNo = 2


If TempRS.State = 1 Then TempRS.Close
TempRS.Open "select * from users where username='" & UCase(Text1) & "' and pword = '" & Text2 & "'", Conn
If TempRS.EOF = True Then
MsgBox ("The entered UserName or Password is not Correct")
Text1.SetFocus
LNo = 1
Else
If TempRS.State = 1 Then TempRS.Close
TempRS.Open "select * from users where username='ADMIN' and pword = '" & Text2 & "'", Conn
If Not TempRS.EOF = True Then
UserNameVar = "ADMIN"
End If

Unload Me

End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Activate()
Me.Height = 2740
Me.Width = 5000
Me.Left = 3500
Me.Top = 2500

LNo = 1

'Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Toll_Plaza\MAINDATA.MDB;Persist Security Info=False"
Conn.Open "Provider=MSDAORA.1;Password=tpas;User ID=tpas;Persist Security Info=True"
End Sub

Private Sub Form_Unload(Cancel As Integer)
If LNo = 1 Then End
End Sub

Private Sub Text1_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Text2.SetFocus
End Sub

Private Sub Text2_GotFocus()
Text1.SelStart = 0
Text1.SelLength = Len(Text1)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then butLogin_Click
End Sub
