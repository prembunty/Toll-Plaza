VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ReportForm 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Report"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8520
   Icon            =   "ReportForm.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6030
   ScaleWidth      =   8520
   Begin VB.CommandButton Command6 
      Caption         =   "Print"
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
      Left            =   6720
      TabIndex        =   7
      Top             =   2640
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Print"
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
      Left            =   6720
      TabIndex        =   6
      Top             =   1200
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Close"
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
      Left            =   6720
      TabIndex        =   5
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Billing"
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
      Left            =   6720
      TabIndex        =   4
      Top             =   3600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Route Detail"
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
      Left            =   6720
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin MSComCtl2.DTPicker bDate 
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MMM/yyyy"
      Format          =   16580611
      CurrentDate     =   41640
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Deposits"
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
      Left            =   6720
      TabIndex        =   1
      Top             =   840
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   9975
      _Version        =   393216
      Rows            =   250
      Cols            =   10
      FixedCols       =   0
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
End
Attribute VB_Name = "ReportForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MSF.Clear
MSF.ColWidth(0) = 600
MSF.ColWidth(1) = 1000
MSF.ColWidth(2) = 1600
MSF.ColWidth(3) = 1000
MSF.TextMatrix(0, 0) = "Sl No"
MSF.TextMatrix(0, 1) = "Route No"
MSF.TextMatrix(0, 2) = "Veh.Owner"
MSF.TextMatrix(0, 3) = "Amount"
I = 1
tAmt = 0
If RS.State = 1 Then RS.Close
RS.Open "select rNo,cName,Advance from Allotmentdb where dateIn='" & bDate & "' order by RNo", Conn
Do While Not RS.EOF
MSF.TextMatrix(I, 0) = I
MSF.TextMatrix(I, 1) = RS(0)
MSF.TextMatrix(I, 2) = RS(1) & ""
MSF.TextMatrix(I, 3) = RS(2)
I = I + 1
tAmt = tAmt + RS(2)
RS.MoveNext
Loop
MSF.TextMatrix(I + 1, 2) = "Total"
MSF.TextMatrix(I + 1, 3) = tAmt

End Sub

Private Sub Command2_Click()
MSF.Clear
MSF.ColWidth(0) = 600
MSF.ColWidth(1) = 1000
MSF.ColWidth(2) = 1600
MSF.ColWidth(3) = 1400
MSF.ColWidth(4) = 1400
MSF.TextMatrix(0, 0) = "Sl No"
MSF.TextMatrix(0, 1) = "Route No"
MSF.TextMatrix(0, 2) = "Veh.Owner"
MSF.TextMatrix(0, 3) = "From Date"
MSF.TextMatrix(0, 4) = "To Date"
I = 1
tAmt = 0
If RS.State = 1 Then RS.Close
RS.Open "select rNo,cName,DateIn,DateOut from Allotmentdb where checkout='N' order by RNo", Conn
Do While Not RS.EOF
MSF.TextMatrix(I, 0) = I
MSF.TextMatrix(I, 1) = RS(0)
MSF.TextMatrix(I, 2) = RS(1) & ""
MSF.TextMatrix(I, 3) = DateFormat(RS(2))
MSF.TextMatrix(I, 4) = DateFormat(RS(3))
I = I + 1
RS.MoveNext
Loop
End Sub

Private Sub Command3_Click()
MSF.Clear
MSF.ColWidth(0) = 600
MSF.ColWidth(1) = 1000
MSF.ColWidth(2) = 1600
MSF.ColWidth(3) = 1000
MSF.TextMatrix(0, 0) = "Sl No"
MSF.TextMatrix(0, 1) = "Route No"
MSF.TextMatrix(0, 2) = "Veh. Owner"
MSF.TextMatrix(0, 3) = "DepoitAmt"
MSF.TextMatrix(0, 4) = "Bill Amt"
MSF.TextMatrix(0, 5) = "Bal Amt"
I = 1
tAmt = 0
If RS.State = 1 Then RS.Close
RS.Open "select roomNo,cName,Advamt,Gtotalamt,BAmt from SaMain where sdate=#" & bDate & "# order by RNo", Conn
Do While Not RS.EOF
MSF.TextMatrix(I, 0) = I
MSF.TextMatrix(I, 1) = RS(0)
MSF.TextMatrix(I, 2) = RS(1)
MSF.TextMatrix(I, 3) = RS(2)
MSF.TextMatrix(I, 4) = RS(3)
MSF.TextMatrix(I, 5) = RS(4)
tAmt = tAmt + RS(4)
I = I + 1
RS.MoveNext
Loop
MSF.TextMatrix(I + 1, 2) = "Total"
MSF.TextMatrix(I + 1, 5) = tAmt
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()

If RS.State = 1 Then RS.Close
RS.Open "select rNo,cName,Advance from Allotmentdb where dateIn='" & bDate & "' order by RNo", Conn
Set AdvReport.DataSource = RS

AdvReport.Sections("section4").Controls("L1").Caption = bDate
AdvReport.Show

End Sub

Private Sub Command6_Click()
If RS.State = 1 Then RS.Close
RS.Open "select rNo,cName,DateIn,DateOut from Allotmentdb where checkout='N' order by RNo", Conn
Set roomReport.DataSource = RS

roomReport.Sections("section4").Controls("L1").Caption = bDate
roomReport.Show
End Sub

Private Sub Form_Load()
Me.Width = 8640
Me.Height = 8000
Me.Top = 0
Me.Left = 0
bDate = Date
End Sub
