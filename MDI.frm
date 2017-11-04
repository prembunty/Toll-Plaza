VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "NICE ROAD TOLL FEE Automation"
   ClientHeight    =   5685
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   10200
   Icon            =   "MDI.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDI.frx":000C
   ScrollBars      =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Menu RegMenu 
      Caption         =   "REGISTRATION MASTER"
      Index           =   6
   End
   Begin VB.Menu CustomerMenu 
      Caption         =   "VEHICLE MASTER"
      Index           =   1
   End
   Begin VB.Menu RoomMenu 
      Caption         =   "ROUTE MASTER"
      Index           =   2
   End
   Begin VB.Menu AllotmentMenu 
      Caption         =   "DEPOSIT MASTER"
      Index           =   3
   End
   Begin VB.Menu BillingMenu 
      Caption         =   "BILLING"
      Index           =   4
   End
   Begin VB.Menu RMENU 
      Caption         =   "REPORT"
      Index           =   5
   End
   Begin VB.Menu EXIT 
      Caption         =   "EXIT"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AllotmentMenu_Click(Index As Integer)
AllotmentForm.Show
End Sub

Private Sub BillingMenu_Click(Index As Integer)
BillForm.Show
End Sub

Private Sub CustomerMenu_Click(Index As Integer)
CustomerForm.Show
End Sub

Private Sub ItemMenu_Click(Index As Integer)
ItemForm.Show
End Sub

Private Sub EXIT_Click()
End
End Sub

Private Sub MDIForm_Load()
Me.Width = 7785
Me.Height = 8000
'dbpath = App.Path & "\MainData.mdb"
'Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info = False; Data Source=" & dbpath & "", "Admin", "", False
'Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Toll_Plaza\MAINDATA.MDB;Persist Security Info=False"
'Conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\Toll_Plaza\MAINDATA.MDB;Persist Security Info=False"
End Sub


Private Sub RegMenu_Click(Index As Integer)
masregForm.Show
End Sub


Private Sub RMENU_Click(Index As Integer)
ReportForm.Show
End Sub

Private Sub RoomMenu_Click(Index As Integer)
RoomForm.Show
End Sub
