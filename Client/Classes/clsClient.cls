VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public strSQL As String
Dim conn As ADODB.Connection
Dim rsData As ADODB.Recordset

Private Sub Class_Initialize()
    Set conn = New ADODB.Connection
End Sub

Function showClient() As ADODB.Recordset
On Error GoTo showClientErr
    If conn.State = adStateOpen Then conn.Close
    conn.Open cnStr
    Set rsData = New ADODB.Recordset
    
    strSQL = "SELECT ClientID, Name, DOB, Gender=case Gender when 'F' then 'Female' when 'M' then 'Male' end from Client order by ClientID"
             
    With rsData
        .CursorLocation = adUseClient
        .Open strSQL, conn, adOpenKeyset, adLockBatchOptimistic
    End With
    If rsData.RecordCount = 0 Then GoTo showClientExit
    Set showClient = rsData
    
showClientExit:
    rsData.ActiveConnection = Nothing
    conn.Close
    Exit Function
    
showClientErr:
    MsgBox Err.Description, vbCritical, "Client Module"
    GoTo showClientExit
End Function

Function getDataClient(ClientID As String) As ADODB.Recordset
On Error GoTo getDataClientErr
    If conn.State = adStateOpen Then conn.Close
    conn.Open cnStr
    Set rsData = New ADODB.Recordset
    
    strSQL = "SELECT ClientID, Name, DOB, Gender from Client where ClientID='" & ClientID & "'"
             
    With rsData
        .CursorLocation = adUseClient
        .Open strSQL, conn, adOpenKeyset, adLockBatchOptimistic
    End With
    If rsData.RecordCount = 0 Then GoTo getDataClientExit
    Set getDataClient = rsData
    
getDataClientExit:
    Exit Function
    
getDataClientErr:
    MsgBox Err.Description, vbCritical, "Client Module"
    GoTo getDataClientExit
End Function

Public Function saveDataClient(ClientName As String, Gender As String, DOB As Date, Status As String, Optional ClientID As String)
    If conn.State = adStateOpen Then conn.Close
    conn.Open cnStr
    
    If Status = "NEW" Then
        strSQL = "insert into Client(ClientID,Name,Gender,DOB,CreatedBy,CreatedDate) " & _
                 "select ClientID=dbo.fnGenClientID(), '" & ClientName & "', '" & Left(Gender, 1) & "', '" & Format(DOB, "MM/dd/yyyy") & "', 'SYST', getdate()"
    Else
        strSQL = "update client set Name='" & ClientName & "',Gender='" & Left(Gender, 1) & "',DOB='" & Format(DOB, "MM/dd/yyyy") & "'," & _
                 "UpdatedBy='SYST', UpdatedDate=getdate() where clientid='" & ClientID & "'"
    End If
    conn.Execute strSQL
End Function

Public Function deleteDataClient(ClientID As String)
    If conn.State = adStateOpen Then conn.Close
    conn.Open cnStr
    
    strSQL = "delete client where clientid='" & ClientID & "'"
    conn.Execute strSQL
End Function

