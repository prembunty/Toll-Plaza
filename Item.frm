VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form ItemForm 
   BackColor       =   &H008080FF&
   Caption         =   "VEHICLE TYPE MASTER"
   ClientHeight    =   7710
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8835
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Item.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   8835
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Height          =   6015
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   6975
      Begin VB.TextBox TxtCNo 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   1800
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox TxtAdv 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   1320
         TabIndex        =   29
         Top             =   5640
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid MSF 
         Height          =   2775
         Left            =   120
         TabIndex        =   10
         Top             =   2760
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   4895
         _Version        =   393216
         Rows            =   50
         Cols            =   6
         FixedCols       =   0
         ForeColorSel    =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox TxtGTotal 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   5640
         TabIndex        =   14
         Top             =   5640
         Width           =   1215
      End
      Begin VB.ListBox List 
         ForeColor       =   &H8000000D&
         Height          =   1740
         Left            =   3480
         TabIndex        =   15
         Top             =   480
         Width           =   3255
      End
      Begin VB.TextBox TxtTotal 
         Appearance      =   0  'Flat
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   1560
         TabIndex        =   7
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox TxtRate 
         Appearance      =   0  'Flat
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   840
         TabIndex        =   6
         Top             =   2400
         Width           =   735
      End
      Begin VB.TextBox TxtRNo 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
      Begin VB.TextBox TxtIName 
         Appearance      =   0  'Flat
         Height          =   270
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   3255
      End
      Begin VB.TextBox TxtQty 
         Appearance      =   0  'Flat
         ForeColor       =   &H8000000D&
         Height          =   270
         Left            =   120
         TabIndex        =   5
         Top             =   2400
         Width           =   735
      End
      Begin MSComCtl2.DTPicker IDate 
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1080
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
         CustomFormat    =   "dd-MMM-yy"
         Format          =   71368707
         CurrentDate     =   38065
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Veh. No"
         Height          =   255
         Left            =   1800
         TabIndex        =   30
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Deposit"
         Height          =   255
         Left            =   360
         TabIndex        =   28
         Top             =   5640
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "GrandTotal"
         Height          =   255
         Left            =   4440
         TabIndex        =   24
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Type of Vehicle"
         Height          =   255
         Left            =   3480
         TabIndex        =   23
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Total"
         Height          =   255
         Left            =   1560
         TabIndex        =   22
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   495
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Way"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Route No"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Vehicle Type"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Rate"
         Height          =   255
         Left            =   840
         TabIndex        =   17
         Top             =   2160
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   25
      Top             =   5880
      Width           =   6975
      Begin VB.CommandButton butNew 
         Caption         =   "&New"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   735
      End
      Begin VB.CommandButton butSave 
         Caption         =   "&Save"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1080
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton butUpdate 
         Caption         =   "&Update"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         TabIndex        =   11
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton butDelete 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3360
         TabIndex        =   12
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton butClose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   5880
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton butList 
         Caption         =   "&List"
         Height          =   375
         Left            =   4680
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox TxtRefNo 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   270
      Left            =   720
      TabIndex        =   26
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "RefNo"
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "ItemForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim QtyVar, RateVar As String
Dim GTotal As Currency
Dim RefNoVar, RINoVar, I As Integer
Dim ImRs As Recordset
Private Sub butClose_Click()
Unload Me
End Sub

Private Sub butDelete_Click()
    If vbYes = MsgBox("Are You sure,You want delete this record", vbYesNo) Then
    If RIRS.State = 1 Then RIRS.Close
    RIRS.Open "Select * from Itemdb where RefNo=" & Val(TxtRefNo) & " and CNo=" & Val(TxtCNo) & " and RNo=" & Val(TxtRNo) & "", conn
    If RIRS.EOF = False Then
    conn.Execute " Delete from Itemdb where RefNo=" & Val(TxtRefNo) & " and CNo=" & Val(TxtCNo) & " and RNo=" & Val(TxtRNo) & ""
    MsgBox ("record deleted")
    Else
    MsgBox ("Please check IDate")
    End If
    End If
    butUpdate.Enabled = False
    butDelete.Enabled = False
    butList.Enabled = True
End Sub

Private Sub butList_Click()
Dim DateInVar, MyDate, DateDiffVar As Date
MsfInIt
' to get items from itemdb
        If RIRS.State = 1 Then RIRS.Close
        RIRS.Open "Select RefNo,IDate,IName,IQty,IRate,ITotal from Itemdb where RNo=" & Val(TxtRNo) & " and cNo=" & Val(TxtCNo) & "", conn
        If RIRS.EOF = False Then
        I = 1
        GTotal = 0
        Do While Not RIRS.EOF
        MSF.TextMatrix(I, 0) = RIRS(0)
        MSF.TextMatrix(I, 1) = Format(RIRS(1), "dd-MMM-yy")
        MSF.TextMatrix(I, 2) = RIRS(2)
        MSF.TextMatrix(I, 3) = RIRS(3)
        MSF.TextMatrix(I, 4) = RIRS(4)
        MSF.TextMatrix(I, 5) = RIRS(5)
        GTotal = GTotal + RIRS(5)
        RIRS.MoveNext
        I = I + 1
        Loop
        Else
        MsgBox ("Please enter RNo")
        End If
        'to get room rent from roomdb
If RIRS.State = 1 Then RIRS.Close
RIRS.Open "select Rate1 from RoomDb where RNo=" & Val(TxtRNo) & "", conn
If RIRS.EOF = False Then
If RIRS1.State = 1 Then RIRS1.Close
RIRS1.Open "select Datein,Advance from Allotmentdb where RNo=" & Val(TxtRNo) & " and CNo=" & Val(TxtCNo) & "", conn
MyDate = Format(Date, "dd-MMM-yyyy")
DateInVar = Format(RIRS1(0), "dd-MMM-yyyy")
DateDiffVar = DateDiff("d", DateInVar, MyDate)
I = I + 1
MSF.TextMatrix(I, 2) = "Room Rent"
MSF.TextMatrix(I, 5) = (RIRS(0) * DateDiffVar)
GTotal = GTotal + MSF.TextMatrix(I, 5)
I = I + 1
TxtAdv = RIRS1(1)
GTotal = GTotal - RIRS1(1)
End If
txtGTotal = GTotal
End Sub

Private Sub butNew_Click()
ClearTxtControls Me, TextBox
TxtRNo.SetFocus
    If RIRS.State = 1 Then RIRS.Close
    RIRS.Open " select max(RefNo) from Itemdb ", conn
    RefNoVar = IIf(IsNull(RIRS(0)), 0, RIRS(0)) + 1
    TxtRefNo = RefNoVar
    MsgBox RefNoVar
    butNew.Enabled = False
    butSave.Enabled = True
    
End Sub

Private Sub butSave_Click()
    If RIRS.State = 1 Then RIRS.Close
    RIRS.Open "Select * from Itemdb where RefNo=" & Val(RefNoVar) & "", conn
    If RIRS.EOF = True Then
    conn.Execute "insert into itemdb (RefNo,CNo,RNo,IDate,IName,IQty,IRate,ITotal,IGrandTotal) values(" & Val(TxtRefNo) & "," & Val(TxtCNo) & "," & Val(TxtRNo) & ",#" & Format(IDate, "dd-MMM-yy") & "#,'" & TxtIName & "'," & Val(TxtQty) & ",'" & txtRate & "','" & txtTotal & "','" & txtGTotal & "')"
    MsgBox ("record inserted")
    Else
    MsgBox ("Please check RNo")
    End If
    butSave.Enabled = False
    butNew.Enabled = True
End Sub


Private Sub butUpdate_Click()
If RIRS.State = 1 Then RIRS.Close
RIRS.Open "Select * from Itemdb where RefNo=" & Val(TxtRefNo) & " and IDate=#" & Format(IDate, "dd-MMM-yy") & "# and IName='" & TxtIName & "' ", conn
If RIRS.EOF = False Then
conn.Execute "update Itemdb set RefNo=" & Val(TxtRefNo) & ",CNo=" & Val(TxtCNo) & ",RNo=" & Val(TxtRNo) & ",IDate=#" & Format(IDate, "dd-MMM-yy") & "#,IName='" & TxtIName & "',IQty=" & Val(TxtQty) & ",IRate='" & txtRate & "',ITotal='" & txtTotal & "',IGrandtotal='" & txtGTotal & "' where RefNo=" & Val(TxtRefNo) & " and CNo=" & Val(TxtCNo) & " and RNo=" & Val(TxtRNo) & ""
MsgBox ("Record Updated")
Else
MsgBox ("Please check IDate")
End If
    butUpdate.Enabled = False
    butDelete.Enabled = False
    butList.Enabled = True
End Sub
Private Sub Form_Load()
Me.Height = 8000
Me.Width = 7740
Me.Top = 0
Me.Left = 0
MsfInIt
End Sub
Sub MsfInIt()
MSF.Clear
MSF.ColWidth(0) = 0
MSF.ColWidth(1) = 1200
MSF.ColWidth(2) = 2350
MSF.ColWidth(3) = 750
MSF.ColWidth(4) = 950
MSF.ColWidth(5) = 950
MSF.TextMatrix(0, 0) = "RefNo"
MSF.TextMatrix(0, 1) = "Date"
MSF.TextMatrix(0, 2) = "Items"
MSF.TextMatrix(0, 3) = "Qty"
MSF.TextMatrix(0, 4) = "Rate"
MSF.TextMatrix(0, 5) = "Total"

End Sub

Private Sub List_DblClick()
    If RIRS.State = 1 Then RIRS.Close
    RIRS.Open "Select * from Itemdb Where RNo=" & Val(TxtRNo) & " and CNo=" & Val(TxtCNo) & "", conn
    If RIRS.EOF = False Then
    TxtRefNo = RIRS(0)
    TxtCNo = RIRS(1)
    TxtRNo = RIRS(2)
    IDate = RIRS(3)
    TxtIName = RIRS(4)
    TxtQty = RIRS(5)
    txtRate = RIRS(6)
    Txtitotal = RIRS(7)
    End If
  
End Sub

Private Sub MSF_Click()
    If MSF.TextMatrix(MSF.Row, 0) = "" Then Exit Sub
    RefNoVar = MSF.TextMatrix(MSF.Row, 0)
    If RIRS.State = 1 Then RIRS.Close
    RIRS.Open "select * from itemdb where RefNo=" & Val(RefNoVar) & "", conn
    If RIRS.EOF = False Then
    TxtRefNo = RIRS(0) & ""
    TxtCNo = RIRS(1) & ""
    TxtRNo = RIRS(2) & ""
    IDate = RIRS(3) & ""
    TxtIName = RIRS(4) & ""
    TxtQty = RIRS(5) & ""
    txtRate = RIRS(6) & ""
    txtTotal = RIRS(6) & ""
    txtGTotal = RIRS(7) & ""
    End If
    butList.Enabled = False
    butUpdate.Enabled = True
    butDelete.Enabled = True
End Sub

Private Sub TxtCNo_Change()
    List.Clear
    If RIRS.State = 1 Then RIRS.Close
    RIRS.Open "Select IName,RefNo from Itemdb where RNo=" & Val(TxtRNo) & " and CNo=" & Val(TxtCNo) & " order by RNo ", conn
    Do While Not RIRS.EOF
    List.AddItem (RIRS(0))
    RIRS.MoveNext
    Loop
End Sub

Private Sub TxtRate_Change()
QtyVar = TxtQty.Text
If txtRate = "" Then
txtRate = ""
txtRate.SetFocus
Exit Sub
Else
RateVar = txtRate.Text
End If
txtTotal = (QtyVar * RateVar)
txtTotal = Format$(txtTotal.Text, "#.00")
End Sub

    
