VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form AllotmentForm 
   BackColor       =   &H00C0E0FF&
   Caption         =   "DEPOSIT MASTER"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7650
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Allotment2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   7650
   Begin VB.TextBox AltNo 
      Height          =   375
      Left            =   3480
      TabIndex        =   30
      Top             =   2760
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   975
      Left            =   120
      TabIndex        =   25
      Top             =   4920
      Width           =   7455
      Begin VB.CommandButton ButLogOut 
         Caption         =   "&LogOut"
         Height          =   375
         Left            =   3480
         TabIndex        =   28
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton butClose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   6480
         TabIndex        =   27
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton butSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox TxtAdd2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   1125
      Width           =   3375
   End
   Begin MSComCtl2.DTPicker TimeIn 
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   2040
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "hh:mm"
      Format          =   70254595
      CurrentDate     =   0.5
      MaxDate         =   0.5
      MinDate         =   4.16666666666667E-02
   End
   Begin VB.TextBox TxtAdd1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox TxtAdvance 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   11
      Top             =   2760
      Width           =   975
   End
   Begin MSComCtl2.DTPicker TimeOut 
      Height          =   375
      Left            =   3480
      TabIndex        =   10
      Top             =   3600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "hh:mm"
      Format          =   70254595
      CurrentDate     =   0.833333333333333
      MaxDate         =   0.5
      MinDate         =   4.16666666666667E-02
   End
   Begin MSComCtl2.DTPicker DateOut 
      Height          =   375
      Left            =   1200
      TabIndex        =   9
      Top             =   3600
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   70254595
      CurrentDate     =   38065
      MaxDate         =   2958131
      MinDate         =   32874
   End
   Begin MSComCtl2.DTPicker DateIn 
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   2040
      Width           =   1455
      _ExtentX        =   2566
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
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   70254595
      CurrentDate     =   41640
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   30
      Left            =   5400
      TabIndex        =   22
      Top             =   2280
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
      _Version        =   393216
      Format          =   70254593
      CurrentDate     =   38065
   End
   Begin VB.TextBox TxtCNo 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   1080
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox TxtCName 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   3375
   End
   Begin VB.TextBox TxtAdd3 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   960
      TabIndex        =   4
      Top             =   1410
      Width           =   3375
   End
   Begin VB.TextBox TxtCity 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   5
      Top             =   1680
      Width           =   1695
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   4695
      Left            =   4560
      TabIndex        =   12
      Top             =   120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   8281
      _Version        =   393216
      Rows            =   100
      Cols            =   3
      FixedCols       =   0
      ForeColorSel    =   0
      Appearance      =   0
   End
   Begin VB.TextBox TxtPin 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3000
      TabIndex        =   6
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Alt No"
      Height          =   375
      Left            =   2760
      TabIndex        =   29
      Top             =   2760
      Width           =   615
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs."
      Height          =   255
      Left            =   2280
      TabIndex        =   24
      Top             =   2760
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000007&
      DrawMode        =   12  'Nop
      X1              =   960
      X2              =   4320
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Deposit"
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   2805
      Width           =   975
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Valid up to"
      Height          =   255
      Left            =   0
      TabIndex        =   21
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "TimeOut"
      Height          =   210
      Left            =   2640
      TabIndex        =   20
      Top             =   3720
      Width           =   855
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Time In"
      Height          =   210
      Left            =   2640
      TabIndex        =   19
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   1755
      Width           =   495
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Veh. No"
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Owner"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   585
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   915
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "PIN"
      Height          =   255
      Left            =   2655
      TabIndex        =   14
      Top             =   1755
      Width           =   375
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Valid from"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   2160
      Width           =   975
   End
End
Attribute VB_Name = "AllotmentForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ButClose_Click()
Unload Me
End Sub
Private Sub ListRec()
If RIRS.State = 1 Then RIRS.Close
RIRS.Open "select CNo,cName from Customerdb order by cNo", Conn
MSFInit
I = 1
Do While Not RIRS.EOF
MSF.TextMatrix(I, 0) = RIRS(0)
MSF.TextMatrix(I, 1) = RIRS(1)
RIRS.MoveNext
I = I + 1
Loop
End Sub

Private Sub ButLogOut_Click()
'    If RIRS.State = 1 Then RIRS.Close
'    RIRS.Open "select * from Allotmentdb where rNo=" & Val(TxtRNo) & " and checkOut='N'", Conn
'If RIRS.EOF = False Then
'Conn.Execute "update AllotmentDb set DateOut=#" & Format(DateOut, "dd/MMM/yyyy") & "#,advance=" & Val(TxtAdvance) & ",checkout='Y',LOutdate=#" & Format(DateOut, "dd/MMM/yyyy") & "# where RNo=" & Val(TxtRNo) & " and CheckOut='N'"
'End If
'CNo = Val(TxtCNo)
'CName = TxtCName
'RoomNo = Val(TxtRNo)
BillForm.Show
Unload Me
End Sub

Private Sub butSave_Click()
    If RIRS.State = 1 Then RIRS.Close
    RIRS.Open "select * from Allotmentdb where rNo=" & Val(TxtRNo) & " and checkOut='N'", Conn
If RIRS.EOF = False Then
'conn.Execute "update AllotmentDb"
Else
    If RIRS.State = 1 Then RIRS.Close
    RIRS.Open "select max(AltNo) from Allotmentdb", Conn
    AltNoVar = IIf(IsNull(RIRS(0)), 1001, RIRS(0)) + 1
    Conn.Execute "insert into Allotmentdb (AltNo,CNo,DateIn,TimeIn,DateOut,TimeOut,Advance,CheckOut,CName) values(" & Val(AltNoVar) & "," & Val(TxtCNo) & ",'" & DateFormat(DateIn) & "','" & TimeIn & "','" & DateFormat(DateOut) & "','" & TimeOut & "'," & Val(TxtAdvance) & ",'N','" & TxtCName & "')"
    End If
 
End Sub

Sub allotmentformactivate()

    LODate = Format(Date, "dd/MMM/yyyy")
    LOTime = Format(Time, "ShortTime")
    
End Sub



Private Sub Form_Load()
    Me.Height = 8000
    Me.Width = 7740
    Me.Top = 0
    Me.Left = 0
    TxtRNo = RoomNo
    DateIn = Format(Date, "dd/MMM/yyyy")
    DateOut = Format(Date, "dd/MMM/yyyy")
    timevar = Format(Now, "ShortTime")
    
  
If RIRS.State = 1 Then RIRS.Close
RIRS.Open "select * from allotmentdb where RNo=" & Val(RoomNo) & " and checkOut='N'", Conn
If RIRS.EOF = False Then
TxtCNo = RIRS(1) & ""
TxtRNo = RIRS(3) & ""
DateIn = RIRS(4)
TimeIn = RIRS(5) & ""
DateOut = RIRS(6) & ""
TimeOut = RIRS(7) & ""
TxtAdvance = RIRS(8) & ""

    If RIRS.State = 1 Then RIRS.Close
    'RIRS.Open "select CNo,CName,Add1,Add2,Add3,City,Pin from customerdb where CNo=" & Val(TxtCNo) & "", Conn
    RIRS.Open "select CNo,CName,Add1,Add2,Add3,City,Pin from customerdb where CNo= '" & Combo1 & "' ", Conn
    If RIRS.EOF = False Then
    TxtCNo = RIRS(0)
    TxtCName = RIRS(1)
    TxtAdd1 = RIRS(2)
    TxtAdd2 = RIRS(3)
    TxtAdd3 = RIRS(4)
    TxtCity = RIRS(5)
    TxtPin = RIRS(6)
    End If

End If

MSFInit
ListRec
End Sub
Sub MSFInit()
MSF.Clear
MSF.ColWidth(0) = 500
MSF.ColWidth(1) = 1150

MSF.TextMatrix(0, 0) = "CNo"
MSF.TextMatrix(0, 1) = "Name"

End Sub


Private Sub MSF_Click()
    If MSF.TextMatrix(MSF.Row, 0) = "" Then Exit Sub
    CNoVar = MSF.TextMatrix(MSF.Row, 0)
    If RIRS.State = 1 Then RIRS.Close
    RIRS.Open "select CNo,CName,Add1,Add2,Add3,City,Pin from customerdb where CNo='" & CNoVar & "'", Conn
    If RIRS.EOF = False Then
    TxtCNo = RIRS(0)
    TxtCName = RIRS(1)
    TxtAdd1 = RIRS(2)
    TxtAdd2 = RIRS(3)
    TxtAdd3 = RIRS(4)
    TxtCity = RIRS(5)
    TxtPin = RIRS(6)
    End If

        
End Sub

