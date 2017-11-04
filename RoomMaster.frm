VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form RoomForm 
   BackColor       =   &H00C0FFFF&
   Caption         =   "ROUTE  MASTER"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7605
   Icon            =   "RoomMaster.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   7605
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Height          =   735
      Left            =   120
      TabIndex        =   14
      Top             =   4800
      Width           =   7335
      Begin VB.CommandButton butList 
         Caption         =   "&List"
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
         Left            =   4440
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton butClose 
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
         Left            =   6360
         TabIndex        =   19
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton butDelete 
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
         Left            =   3360
         TabIndex        =   18
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton butUpdate 
         Caption         =   "&Update"
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
         Left            =   2280
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton butNew 
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
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton butSave 
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
         Left            =   1200
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   3735
      Left            =   4920
      TabIndex        =   6
      Top             =   240
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   6588
      _Version        =   393216
      Rows            =   100
      FixedCols       =   0
      ForeColorSel    =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TxtRNo 
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
      Left            =   1920
      TabIndex        =   0
      Top             =   195
      Width           =   735
   End
   Begin VB.TextBox TxtDes 
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
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   2895
   End
   Begin VB.TextBox TxtType 
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
      Left            =   1920
      TabIndex        =   2
      Top             =   1035
      Width           =   2895
   End
   Begin VB.TextBox TxtRate1 
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
      Left            =   1920
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox TxtRate2 
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
      Left            =   1920
      TabIndex        =   4
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox TxtFacility 
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
      Left            =   1920
      TabIndex        =   5
      Top             =   2400
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Available Route Numbers"
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
      Height          =   240
      Left            =   4680
      TabIndex        =   13
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Two Way"
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
      Height          =   210
      Left            =   120
      TabIndex        =   12
      Top             =   1965
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "One Way"
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
      Left            =   120
      TabIndex        =   11
      Top             =   1515
      Width           =   1215
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Facility Available"
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
      Height          =   240
      Left            =   120
      TabIndex        =   10
      Top             =   2430
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Route No"
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
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Left            =   120
      TabIndex        =   8
      Top             =   645
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
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
      Left            =   120
      TabIndex        =   7
      Top             =   1080
      Width           =   615
   End
End
Attribute VB_Name = "RoomForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub butClose_Click()
Unload Me
End Sub

Private Sub butDelete_Click()
If RIRS.State = 1 Then RIRS.Close
RIRS.Open "Select * from roomdb where RNo=" & Val(TxtRNo) & " ", Conn
If vbYes = MsgBox("Are you sure,you want to delete this Record", vbYesNo) Then
If RIRS.EOF = False Then
Conn.Execute "delete from roomdb where RNo=" & Val(TxtRNo) & ""
Else
MsgBox ("Please check Room Number")
End If
End If

        MSFInit
        List
        ClearTxtControls Me, TextBox
        butNew.SetFocus
        butDelete.Enabled = False
        butUpdate.Enabled = False
        butList.Enabled = True
End Sub

Private Sub butList_Click()
MSFInit
List

End Sub
Sub List()
If RIRS.State = 1 Then RIRS.Close
RIRS.Open "Select RNo,Rate1 from roomdb order by RNo ", Conn
MSFInit
MSF.Row = 1
Do While Not RIRS.EOF
MSF.TextMatrix(MSF.Row, 0) = RIRS(0)
MSF.TextMatrix(MSF.Row, 1) = RIRS(1)
RIRS.MoveNext
MSF.Row = MSF.Row + 1
Loop
End Sub

Private Sub butNew_Click()
ClearTxtControls Me, TextBox
TxtRNo.SetFocus
butNew.Enabled = False
butSave.Enabled = True
End Sub

Private Sub butSave_Click()
If TxtRNo = "" Then
MsgBox ("Please enter correct data")
butNew.Enabled = True
butSave.Enabled = False
Exit Sub
End If
If RIRS.State = 1 Then RIRS.Close
RIRS.Open "Select * from roomdb where RNo=" & TxtRNo & "", Conn
If RIRS.EOF = True Then
Conn.Execute "insert into roomdb (RNo,Description,type,rate1,Rate2,facility) values(" & Val(TxtRNo) & ",'" & TxtDes & "','" & TxtType & "','" & TxtRate1 & "','" & TxtRate2 & "','" & TxtFacility & "')"
MsgBox ("Record inserted")
Else
MsgBox ("Please check Route Number")
End If

        MSFInit
        List
        ClearTxtControls Me, TextBox

        butNew.Enabled = True
        butSave.Enabled = False
End Sub

Private Sub butUpdate_Click()
If RIRS.State = 1 Then RIRS.Close
RIRS.Open "Select * from roomdb where RNo=" & Val(TxtRNo) & "", Conn
If RIRS.EOF = False Then
Conn.Execute "update Roomdb set RNo=" & Val(TxtRNo) & ",Description='" & TxtDes & "',Type='" & TxtType & "',Rate1='" & TxtRate1 & "',Rate2='" & TxtRate2 & "',Facility='" & TxtFacility & "' where RNo=" & Val(TxtRNo) & " "
MsgBox ("Record updated")
Else
MsgBox ("Please check Route Number")
End If
        MSFInit
        List
        ClearTxtControls Me, TextBox
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
End Sub
Sub MSFInit()
MSF.Clear
MSF.ColWidth(0) = 750
MSF.ColWidth(1) = 1200
MSF.TextMatrix(0, 0) = "RNum"
MSF.TextMatrix(0, 1) = "Rate"
End Sub

Private Sub MSF_DblClick()
If MSF.TextMatrix(MSF.Row, 0) = "" Then Exit Sub
RINoVar = MSF.TextMatrix(MSF.Row, 0)
If RIRS.State = 1 Then RIRS.Close
RIRS.Open "select * from Roomdb where RNo=" & Val(RINoVar) & " ", Conn
If RIRS.EOF = False Then
TxtRNo = RIRS(0) & ""
TxtDes = RIRS(1) & ""
TxtType = RIRS(2) & ""
TxtRate1 = RIRS(3) & ""
TxtRate2 = RIRS(4) & ""
TxtFacility = RIRS(5) & ""
End If
butList.Enabled = False
butDelete.Enabled = True
butUpdate.Enabled = True
End Sub

