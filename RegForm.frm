VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form EdForm 
   BackColor       =   &H00C0E0FF&
   Caption         =   "RTO Details"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   Icon            =   "RegForm.frx":0000
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   9480
   Begin VB.TextBox Text4 
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
      Left            =   4800
      MaxLength       =   4
      TabIndex        =   21
      Top             =   2400
      Width           =   735
   End
   Begin VB.ComboBox Text2 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      TabIndex        =   20
      Top             =   2400
      Width           =   1215
   End
   Begin VB.ComboBox Text1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   19
      Top             =   2400
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Add"
      Height          =   375
      Left            =   6240
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2400
      Width           =   495
   End
   Begin VB.TextBox Text3 
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
      Left            =   2280
      MaxLength       =   300
      TabIndex        =   1
      Top             =   2400
      Width           =   2415
   End
   Begin VB.TextBox Text5 
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
      Left            =   5640
      MaxLength       =   3
      TabIndex        =   2
      Top             =   2400
      Width           =   495
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   6135
      Left            =   6960
      TabIndex        =   9
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
      BackColor       =   &H00C0E0FF&
      Height          =   735
      Left            =   240
      TabIndex        =   8
      Top             =   5640
      Width           =   6495
      Begin VB.CommandButton Button 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&New"
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
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Button 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Save"
         Enabled         =   0   'False
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
         Index           =   1
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Button 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Modify"
         Enabled         =   0   'False
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
         Index           =   2
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Button 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Delete"
         Enabled         =   0   'False
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
         Index           =   3
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Button 
         BackColor       =   &H00C0E0FF&
         Caption         =   "&Close"
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
         Index           =   4
         Left            =   5520
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   2175
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   3836
      _Version        =   393216
      Rows            =   25
      Cols            =   6
      FixedCols       =   0
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "%"
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
      Left            =   5760
      TabIndex        =   22
      Top             =   2160
      Width           =   375
   End
   Begin VB.Label LblNo 
      BackColor       =   &H00C0E0FF&
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
      Left            =   1080
      TabIndex        =   18
      Top             =   480
      Width           =   975
   End
   Begin VB.Label LblName 
      BackColor       =   &H00C0E0FF&
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
      Left            =   3480
      TabIndex        =   15
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Name :"
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
      Left            =   2520
      TabIndex        =   14
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "App No. "
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
      Left            =   120
      TabIndex        =   13
      Top             =   480
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "RTO  Name"
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
      Left            =   2280
      TabIndex        =   12
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Year"
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
      Left            =   4800
      TabIndex        =   11
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "State"
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
      Left            =   120
      TabIndex        =   10
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Branch"
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
      Left            =   960
      TabIndex        =   7
      Top             =   2160
      Width           =   855
   End
End
Attribute VB_Name = "EdForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Num As Long

Private Sub Button_Click(Index As Integer)
Select Case Index
Case 0
Num = 1
MSF1Init
Button(1).Enabled = True
Button(0).Enabled = False
Case 1
If Num <= 1 Then MsgBox ("Experience is not entered Please Check"): Exit Sub
For I = 1 To Num
strSql = "insert into AppEd values "
strSql = strSql & " (" & Val(LblNo) & ",'" & LblName & "'," & I & ",'" & MSF1.TextMatrix(I, 1) & "','" & MSF1.TextMatrix(I, 2) & "','" & MSF1.TextMatrix(I, 3) & "','" & MSF1.TextMatrix(I, 4) & "','" & MSF1.TextMatrix(I, 5) & "')"
Conn.Execute strSql
Next
DisRec
Button(1).Enabled = False
Button(0).Enabled = True

Case 2
If vbYes = MsgBox("Do you want to Modify this Record", vbYesNo) Then
Conn.Execute "delete from  appEx where rno=" & AppNo & ""
For I = 1 To Num
If Not MSF1.TextMatrix(I, 1) = "" Then
strSql = "insert into AppEd values "
strSql = strSql & " (" & Val(LblNo) & ",'" & LblName & "'," & I & ",'" & MSF1.TextMatrix(I, 1) & "','" & MSF1.TextMatrix(I, 2) & "','" & MSF1.TextMatrix(I, 3) & "','" & MSF1.TextMatrix(I, 4) & "','" & MSF1.TextMatrix(I, 5) & "')"
Conn.Execute strSql
End If
Next
DisRec
End If
Case 3
If vbYes = MsgBox("Do you want to delete this Record", vbYesNo) Then
Conn.Execute "delete from  appEd where rno=" & AppNo & ""
DisRec
End If

Case 4
Unload Me
End Select
Num = 1
End Sub

Private Sub Command1_Click()

MSF1.TextMatrix(Num, 0) = Num
MSF1.TextMatrix(Num, 1) = Text2
MSF1.TextMatrix(Num, 2) = Text3
MSF1.TextMatrix(Num, 3) = Text4
MSF1.TextMatrix(Num, 4) = Text5
Num = Num + 1

ClearTxtControls Me, TextBox
End Sub
Sub MSF1Init()
MSF1.Clear
MSF1.ColWidth(0) = 600
MSF1.ColWidth(1) = 1000
MSF1.ColWidth(2) = 1000
MSF1.ColWidth(3) = 1600
MSF1.ColWidth(4) = 800
MSF1.ColWidth(5) = 600
MSF1.TextMatrix(0, 0) = "SlNO"
MSF1.TextMatrix(0, 1) = "Degree"
MSF1.TextMatrix(0, 2) = "Branch"
MSF1.TextMatrix(0, 3) = "School/College"
MSF1.TextMatrix(0, 4) = "Year"
MSF1.TextMatrix(0, 5) = "%"
End Sub
Private Sub Form_Load()
Me.Width = 10230
Me.Height = 6975
Me.Left = 100
Me.Top = 100
DisRec
Degree
Branch
End Sub

Sub MSFInit()
MSF.Clear
MSF.ColWidth(0) = 1600
MSF.ColWidth(1) = 800
MSF.TextMatrix(0, 0) = "Name"
MSF.TextMatrix(0, 1) = "App No"

End Sub

Sub DisRec()
If TempRS.State = 1 Then TempRS.Close
TempRS.Open "select rName,rno from Resume order by rName"
MSFInit
I = 1
Do While Not TempRS.EOF
MSF.TextMatrix(I, 0) = TempRS(0)
MSF.TextMatrix(I, 1) = TempRS(1)
I = I + 1
TempRS.MoveNext
Loop

End Sub

Private Sub MSF_DblClick()
If MSF.TextMatrix(MSF.Row, 0) = "" Then Exit Sub
LblNo = MSF.TextMatrix(MSF.Row, 1)
AppNo = MSF.TextMatrix(MSF.Row, 1)
LblName = MSF.TextMatrix(MSF.Row, 0)
If TempRS.State = 1 Then TempRS.Close
TempRS.Open "select * from AppEd where rno=" & AppNo & " order by slno ", Conn
I = 1
MSF1Init
Do While Not TempRS.EOF
MSF1.TextMatrix(I, 0) = I
MSF1.TextMatrix(I, 1) = TempRS(3) & ""
MSF1.TextMatrix(I, 2) = TempRS(4) & ""
MSF1.TextMatrix(I, 3) = TempRS(5) & ""
MSF1.TextMatrix(I, 4) = TempRS(6) & ""
MSF1.TextMatrix(I, 5) = TempRS(7) & ""
TempRS.MoveNext
I = I + 1
Loop
Num = I - 1

Button(2).Enabled = True
Button(3).Enabled = True
End Sub

Sub Degree()
If TempRS.State = 1 Then TempRS.Close
TempRS.Open "select Cname from MasEd order by cName"
Do While Not TempRS.EOF
Text1.AddItem (TempRS(0))
TempRS.MoveNext
Loop

End Sub
Sub Branch()
If TempRS.State = 1 Then TempRS.Close
TempRS.Open "select Cname from MasBr order by cName"
Do While Not TempRS.EOF
Text2.AddItem (TempRS(0))
TempRS.MoveNext
Loop
End Sub

