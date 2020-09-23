Attribute VB_Name = "MYSQL"

' Code is free to be redistributed, and to be modified.
' 100% Free Code. By: Nathan Huebner (admin@sellchain.com)
' Visit my web site www.sellchain.com ;)

' These will setup your connection.
' Don't forget to add the ADO Control to your form
' Microsoft ADO Control. Drag & Drop the control to any form.
' Do not remove the control from your components.
' To add the control, go to Project > References > Then select the Microsoft ADO Control and Check it, click ok.
' Then drag the control onto a form to create it.

Public conn As ADODB.Connection
Public rs As ADODB.Recordset
Public isMySQLConnected As Boolean


' Public settings, simply call on the public variables
' from any form.

Public MySQL_User As String
Public MySQL_Password As String
Public MySQL_Server As String
Public MySQL_Database As String



Public Sub ConnectMySQL()

Dim ErrorDescription As String
Dim ErrorNumber As String
Dim Verifycount As Long
Dim VerifyStamp As String

On Error GoTo PerformErr
isMySQLConnected = False
Set conn = New ADODB.Connection
Set rs = New ADODB.Recordset

conn.CursorLocation = adUseClient
conn.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
& "SERVER=" & MySQL_Server & ";" _
& "DATABASE=" & MySQL_Database & ";" _
& "UID=" & MySQL_User & ";" _
& "PWD=" & MySQL_Password & ";" _
& "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384

conn.Open
isMySQLConnected = True
Verifycount = 0
' Verify Connection Once

Exit Sub

PerformErr:
' Failed to connect
ErrorDescription = Err.Description
ErrorNumber = Err.Number

MsgBox "A MySQL Connection Error has been returned: " & ErrorDescription & "[" & ErrorNumber & "]", vbCritical, "Could not connect to MySQL"


PerformErr_NoStandardError:
isMySQLConnected = False
End Sub

Public Sub MySQLConnect()
If isMySQLConnected = True Then
Else

ReconnectMySQL:

ConnectMySQL

If isMySQLConnected = False Then
' Connection failed.. Try again...
PauseMe 10 ' Waits 10 seconds, then reconnects
GoTo ReconnectMySQL
End If
End If
End Sub

Public Function QueryDatabase(QueryString As String) As Boolean
On Error Resume Next
MySQLConnect
rs.Close
rs.Open QueryString, conn, adOpenStatic, adLockOptimistic
End Function

Public Function RecordCount() As Long
RecordCount = 0
MySQLConnect
RecordCount = rs.RecordCount
End Function

Public Function getCell(FieldName As String, Row As Long) As String
MySQLConnect
rs.Move (Row - 1)
getCell = rs.Fields(FieldName).value
End Function


Sub PauseMe(Duration)
Dim FullTimer

StartTime = Timer
FullTimer = Timer - StartTime
If Duration = 0 Then
Duration = 1
End If

Do While (Timer - StartTime) < Duration
If StartTime - Timer > 20 Then Exit Do
DoEvents
Loop
End Sub



