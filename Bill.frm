VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form BillForm 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Toll Fee Bill"
   ClientHeight    =   7575
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7305
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Bill.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   7305
   Visible         =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   495
      Left            =   1320
      TabIndex        =   37
      Top             =   6840
      Width           =   4215
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFC0C0&
      ForeColor       =   &H00FFC0C0&
      Height          =   615
      Left            =   120
      TabIndex        =   36
      Top             =   1320
      Visible         =   0   'False
      Width           =   7095
   End
   Begin VB.ComboBox Combo2 
      Height          =   360
      Left            =   3720
      TabIndex        =   33
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   960
      TabIndex        =   32
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Vehicle Verification"
      Height          =   495
      Left            =   1320
      TabIndex        =   31
      Top             =   6840
      Width           =   4215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   855
      Left            =   120
      TabIndex        =   24
      Top             =   5880
      Width           =   7095
      Begin VB.CommandButton butDelete 
         Caption         =   "Delete"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3360
         TabIndex        =   30
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton butModify 
         Caption         =   "Modify"
         Enabled         =   0   'False
         Height          =   375
         Left            =   2280
         TabIndex        =   29
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton butSave 
         Caption         =   "Save"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1200
         TabIndex        =   28
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton butNew 
         Caption         =   "New"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton butClose 
         Caption         =   "Close"
         Height          =   375
         Left            =   5760
         TabIndex        =   26
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton butPrint 
         Caption         =   "Print"
         Height          =   375
         Left            =   4440
         TabIndex        =   25
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton butAdd 
      Caption         =   "Add"
      Height          =   255
      Left            =   6240
      TabIndex        =   16
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtTotal 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   5040
      TabIndex        =   15
      Top             =   1560
      Width           =   1095
   End
   Begin VB.TextBox txtRate 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   4080
      TabIndex        =   14
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox TxtQty 
      Appearance      =   0  'Flat
      Height          =   270
      Left            =   3360
      TabIndex        =   13
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox TxtBillNo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5640
      TabIndex        =   9
      Top             =   600
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   7095
      Begin VB.TextBox TxtBal 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         TabIndex        =   17
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox TxtAdv 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   3480
         Width           =   1095
      End
      Begin VB.TextBox txtGTotal 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   2
         Top             =   3480
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid MSF 
         Height          =   3135
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   5530
         _Version        =   393216
         Rows            =   50
         Cols            =   7
         FixedCols       =   0
         BackColor       =   -2147483639
         ForeColorSel    =   0
         Appearance      =   0
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
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Balance"
         Height          =   255
         Left            =   4800
         TabIndex        =   18
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "GTotal"
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   3480
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Deposited"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   3480
         Width           =   975
      End
   End
   Begin MSComCtl2.DTPicker BDate 
      Height          =   375
      Left            =   5640
      TabIndex        =   10
      Top             =   120
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
      Format          =   16646147
      CurrentDate     =   38065
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   240
      TabIndex        =   35
      Top             =   1560
      Width           =   2895
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   960
      TabIndex        =   34
      Top             =   600
      Width           =   3255
   End
   Begin VB.Label LabelMsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "TOLL PLAZA - NICE ROAD -TOLL BILL"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   960
      Width           =   6975
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   255
      Left            =   5040
      TabIndex        =   22
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Rate"
      Height          =   255
      Left            =   4080
      TabIndex        =   21
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Way"
      Height          =   255
      Left            =   3360
      TabIndex        =   20
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Route Name"
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill Date"
      Height          =   255
      Left            =   4560
      TabIndex        =   12
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No"
      Height          =   255
      Left            =   4560
      TabIndex        =   11
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Route No"
      Height          =   255
      Left            =   2760
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Owner"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Veh.No"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "BillForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tAmtVar As Currency
Dim BillNovar As Long
Private Sub butAdd_Click()
'If TxtItemName = "" Then
'MsgVar = MsgBox("Please Enter the Route Name ", vbDefaultButton1, "Select Item")
'TxtItemName.SetFocus
'Exit Sub
'End If
'If TxtQty = "" Or Val(TxtQty) <= 0 Then
'MsgVar = MsgBox("Please Enter the Quantity ", vbDefaultButton1, "Enter Quantity")
'TxtQty.SetFocus
'Exit Sub
'End If
'If txtRate = "" Or Val(txtRate) <= 0 Then
'MsgVar = MsgBox("Please Enter the Rate ", vbDefaultButton1, "Enter Rate")
'txtRate.SetFocus
'Exit Sub
'End If
'RINoVar = 0



MSF.TextMatrix(RINoVar, 0) = Combo2.Text
MSF.TextMatrix(RINoVar, 1) = Label14.Caption
MSF.TextMatrix(RINoVar, 2) = TxtQty.Text
MSF.TextMatrix(RINoVar, 3) = TxtRate.Text
MSF.TextMatrix(RINoVar, 4) = TxtTotal.Text
RINoVar = RINoVar + 1




'tAmtVar = tAmtVar + Val(txtTotal)
'txtGTotal = tAmtVar
'TxtBal = tAmtVar - Val(TxtAdv)

If vbYes = MsgBox("Do you want enter another Item", vbYesNo, "Another Item") Then
TxtItemName.SetFocus
Else
TxtAdv.SetFocus
End If
End Sub

Private Sub combo1_LostFocus()
If Not Combo1 = "" Then
If RIRS.State = 1 Then RIRS.Close
RIRS.Open "select * from customerdb where CNo = '" & Combo1 & "'", Conn
If Not RIRS.EOF Then
Label13.Caption = RIRS(1)
Combo2.SetFocus
End If
End If
End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Not Combo1 = "" Then
Combo2.SetFocus
End If
End If
End Sub





Private Sub combo2_LostFocus()
If Not Combo2 = "" Then
If RIRS.State = 1 Then RIRS.Close
RIRS.Open "select * from roomdb where RNo = " & Val(Combo2) & " ", Conn
If Not RIRS.EOF Then
Label14.Caption = RIRS(1)
TxtQty.Text = RIRS(2)
TxtRate.Text = RIRS(3)
TxtTotal.Text = RIRS(4)
butAdd.SetFocus
End If
End If
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Not Combo2 = "" Then
butAdd.SetFocus
End If
End If
End Sub

Private Sub Command1_Click()
MsgVar = MsgBox("Vehile Verified OK", vbYesNo)
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub ButClose_Click()
Unload Me
End Sub

Private Sub butNew_Click()
    If RS.State = 1 Then RS.Close
    RS.Open "select max(rNo) from saMain", Conn
    TxtBillNo = IIf(IsNull(RS(0)), 0, RS(0)) + 1
    BillNovar = Val(TxtBillNo)
    RINoVar = 1
    tAmtVar = 0
    butNew.Enabled = False
    butSave.Enabled = True
    butModify.Enabled = False
    butDelete.Enabled = False
    butPrint.Enabled = False
    Combo1.SetFocus
    
End Sub

Private Sub butPrint_Click()

Frame3.Visible = True
Frame2.Visible = False
On Error GoTo errhand
MsgBox "Do you want to print?", vbYesNo, "NICE"
If vbYes Then
Frame3.Visible = False
Frame2.Visible = False
Label1.Visible = False
Label2.Visible = False
Label8.Caption = "Amount"
TxtAdv.Visible = False
TxtGTotal.Visible = False
'Frame2.Caption = ""
Command2.Visible = True
'CommonDialog1.PrinterDefault = True
'CommonDialog1.ShowPrinter
'Command1.Visible = False
cmdAdd.Visible = False
cmdRemove.Visible = False
Label7.Visible = False
cmbItemName.Visible = False
Label8.Visible = False
txtItemQty.Visible = False
Label9.Visible = False
txtItemRate.Visible = False
Label12.Visible = True
PrintForm
'Unload Me
Else
Frame3.Visible = True
End If
errhand:
Frame3.Visible = True
Frame2.Visible = False
Label1.Visible = False
Label2.Visible = False
Label8.Caption = "Amount"
TxtAdv.Visible = False
TxtGTotal.Visible = False



MsgBox "printer is not available", 64, "NICE"
Frame3.Visible = True
'Frame2.Caption = "Item Details ( To remove select and press Remove button)"
Frame2.Visible = False
Command2.Visible = True
'cmdAdd.Visible = True
'cmdRemove.Visible = True
Label7.Visible = True
'cmbItemName.Visible = True
Label8.Visible = True
'txtItemQty.Visible = True
Label9.Visible = True
'txtItemRate.Visible = True
Label12.Visible = False


'BillPrint.Sections("section4").Controls("L1").Caption = TxtCName
'BillPrint.Sections("section4").Controls("L2").Caption = TxtRNo
'BillPrint.Sections("section4").Controls("L3").Caption = bDate
'BillPrint.Sections("section4").Controls("L4").Caption = TxtBillNo

'If RS.State = 1 Then RS.Close
'RS.Open "select * from sadet where rno=" & Val(Combo2.Text) & "", Conn
'RS.Open "select * from sadet where rno=" & Combo2.Text & "", Conn
'Set BillPrint.DataSource = RS
'BillPrint.Show
End Sub

Private Sub butSave_Click()
    
    If Combo1.Text = "" Then
    MsgVar = MsgBox("Please select the coustomer No ", vbDefaultButton1, "Supplier Selection")
    Combo1.SetFocus
    Exit Sub
    End If
    
    'If MSF.TextMatrix(1, 1) = "" Then
    'If vbNo = MsgBox("No Item is selected Please check", vbYesNo, "No Items") Then Exit Sub
    'Exit Sub
    'End If
    
    'RINoVar = RINoVar - 1
        
    
    'For I = 0 To RINoVar
    'strSql = "insert into sadet(RNo,itemName,Qty,Rate,total) values("
    'strSql = strSql & Val(MSF.TextMatrix(1, 0)) & ","
    'strSql = strSql & MSF.TextMatrix(1, 1) & "','"
    'strSql = strSql & MSF.TextMatrix(1, 2) & "','"
    'strSql = strSql & MSF.TextMatrix(1, 3) & "','"
    'strSql = strSql & MSF.TextMatrix(1, 4) & "','"
    'strSql = strSql & MSF.TextMatrix(1, 5) & "','"
    'Conn.Execute strSql
    'Conn.Execute "insert into sadet values(" & Val(MSF.TextMatrix(1, 0)) & ", '" & MSF.TextMatrix(1, 1) & "','" & MSF.TextMatrix(1, 2) & "','" & MSF.TextMatrix(1, 3) & "','" & MSF.TextMatrix(1, 4) & "' "

    Conn.Execute "insert into sadet values(" & Val(Combo2.Text) & ", '" & Label14.Caption & "','" & TxtQty.Text & "','" & TxtRate.Text & "','" & TxtTotal.Text & "') "
    MsgBox "Record Saved successfully"
    'Next
    
   ' strSql = "insert into samain(RNo,RoomNo,sdate,cNo,cname,advAmt,GTotalAmt,BAmt) values("
   ' strSql = strSql & BillNovar & ",'" & TxtRNo & "', #" & bDate.Value & "#, '" & TxtCNo & "', '" & TxtCName & "', " & Val(TxtAdvance) & "," & Val(txtGTotal) & "," & Val(TxtBal) & ")"

   ' Conn.Execute strSql
    
    butNew.Enabled = True
    butSave.Enabled = False
    butModify.Enabled = False
    butDelete.Enabled = False
    butPrint.Enabled = True
    butPrint.Enabled = True
    butNew.SetFocus
    
End Sub

Private Sub Form_Load()
Me.Height = 8000
Me.Width = 7740
Me.Top = 0
Me.Left = 0
MSFInit
bDate = Date
'TxtCNo = Val(CNo)
TxtCName = CName
'TxtRNo = Val(RoomNo)
RINoVar = 1
tAmtVar = 0



If RIRS.State = 1 Then RIRS.Close
RIRS.Open "select CNo from customerdb order by CNo", Conn
'Conn.Execute "select regno from masreg order by regno"
'If TempRS.State = 1 Then TempRS.Close
'TempRS.Open "select regno from masreg order by regno", Conn
Combo1.Clear
Do While Not RIRS.EOF
Combo1.AddItem (RIRS(0))
RIRS.MoveNext
Loop



If RIRS.State = 1 Then RIRS.Close
RIRS.Open "select RNo from Roomdb order by RNo", Conn
'Conn.Execute "select regno from masreg order by regno"
'If TempRS.State = 1 Then TempRS.Close
'TempRS.Open "select regno from masreg order by regno", Conn
Combo2.Clear
Do While Not RIRS.EOF
Combo2.AddItem (RIRS(0))
RIRS.MoveNext
Loop


















'If RS.State = 1 Then RS.Close
'RS.Open "select DateIn,LOutDate,advance from allotmentDB where altNo=(select max(altNo) from allotmentDb where cNo=" & CNo & " and rNo = " & Val(TxtRNo) & ")", Conn
'If RS.EOF = False Then
'LabelMsg = "Date In : " & Format(RS(0), "dd/MMM/yyyy") & "  Date Out : " & Format(RS(1), "dd/MMM/yyyy") & ""
'TxtAdv = RS(2)
'End If
End Sub

Sub MSFInit()
MSF.Clear
MSF.ColWidth(0) = 600
MSF.ColWidth(1) = 2600
MSF.ColWidth(2) = 1000
MSF.ColWidth(3) = 1000
MSF.ColWidth(4) = 1000

MSF.TextMatrix(0, 0) = "Sl No"
MSF.TextMatrix(0, 1) = "Particulars"
MSF.TextMatrix(0, 2) = "Qty"
MSF.TextMatrix(0, 3) = "Price"
MSF.TextMatrix(0, 4) = "Total"

End Sub

Private Sub TxtAdv_LostFocus()
TxtGTotal = TxtTotal
TxtBal.Text = Abs(Val(TxtAdv.Text) - Val(TxtGTotal.Text))
butSave.SetFocus
End Sub

Private Sub TxtBillNo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
BillNovar = Val(TxtBillNo)
End If

End Sub

Private Sub TxtCNo_Change()

End Sub

Private Sub TxtQty_Change()
TxtTotal = Val(TxtQty) * Val(TxtRate)
End Sub

Private Sub TxtRate_Change()
TxtTotal = Val(TxtQty) * Val(TxtRate)
End Sub



