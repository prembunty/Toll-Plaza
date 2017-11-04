Attribute VB_Name = "Module1"

Option Explicit
'Public Conn As New ADODB.Connection
'Public RS As New ADODB.Recordset
'Public TempRS As New ADODB.Recordset
'Public TempRS1 As New ADODB.Recordset

Public RNo, RNoVar As Long
Public K As Long
Public strSql, MsgVar, CatVar, ItemVar, TranMain As String
Public VTVar, BGColor, UserNameVar, DatabasePath As String
Public AppNo As Long


Public Conn As New ADODB.Connection
Public RS As New ADODB.Recordset
Public RIRS As New ADODB.Recordset
Public rsTemp As Recordset
Public TempRS As New ADODB.Recordset

Public dbpath, CName As String
Public AltNoVar As Long
Public RINoVar, RINoVar1, RoomNo, CNo As Integer
Public Var1, Var2 As Integer
Public I, J As Integer
Public tAmt As Currency

Enum CtrlType
     TextBox = 1
     ComboBox = 2
End Enum

Public Sub ClearTxtControls(frm As Object, ControlType As CtrlType, Optional Tagstr As Variant)
Dim Contrl As Object
For Each Contrl In frm.Controls
         If Not (IsMissing(Tagstr)) Then
         If Trim(UCase(Contrl.Tag)) = Trim(UCase(Tagstr)) Then
            Contrl.Text = ""
            Exit For
          End If
          Else
          Select Case ControlType
                 Case CtrlType.ComboBox
                   If TypeOf Contrl Is ComboBox Then Contrl.Text = ""
                 Case CtrlType.TextBox
                   If TypeOf Contrl Is TextBox Then Contrl.Text = ""
          End Select
          End If
    Next
Set Contrl = Nothing
End Sub
Public Function DateFormat(DateVar)
DateFormat = Format(DateVar, "dd/MMM/yyyy")
End Function


Function CheckNum(KeyNum)
If KeyNum = 8 Then CheckNum = KeyNum: Exit Function
If KeyNum < 46 Or KeyNum > 57 Then
CheckNum = 0
MsgBox ("Please Enter Numbers Only")
Else
CheckNum = KeyNum
End If
End Function

Public Sub ValidNumeric(KeyAscii As Integer)
Select Case KeyAscii
Case 8
Case 97
Case 110
Case 47
Case 13
Case 32
Case 48 To 57
 Case Else
  MsgBox "Invalid Input.Please Enter Numeric Types Only..", vbOKOnly + vbExclamation
  KeyAscii = 0
End Select
End Sub

Public Sub ValidNonNumeric(KeyAscii As Integer)
KeyAscii = Asc(UCase$(Chr$(KeyAscii)))
Select Case KeyAscii
 Case Asc(" ")
 Case 65 To 90
 Case 97 To 122
 Case 32
 Case 13
 Case 8
 Case 127
 Case Else
  MsgBox "Invalid Input. Please Don't Enter Numerics...", vbOKOnly + vbExclamation
  KeyAscii = 0
End Select
End Sub


