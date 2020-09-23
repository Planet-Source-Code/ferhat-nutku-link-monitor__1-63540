VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CADO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'CADO
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Members
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private mConnection As Connection
Private mRecordSet As Recordset


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Members
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Private Methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub Class_Initialize()

End Sub


Private Sub Class_Terminate()
   mConnection.Close
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Methods
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetConnectionAccess(ByVal pdbname As String) As Connection
   'Returns a Connection to a Access database
   
   'EXAMPLE:
   'Dim oConnection as Connection
   'Set oConnection = oCADO.GetConnectionAccess(strDBpath)
   
   If (mConnection Is Nothing) Then
      Set mConnection = New Connection
      mConnection.CursorLocation = adUseClient 'res...
      Call mConnection.Open("PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" & pdbname & ";") 'res..
   End If
      
   Set GetConnectionAccess = mConnection
End Function


Public Function GetRecordSetByConnection(ByVal pConnection As Connection, ByVal psqlcmd As String) As Recordset
  'Returns a RecordSet object from a Connection & an sql string
  
  'EXAMPLE:
  'Dim oADO As New ClassADO
  'Public oRecordSet As Recordset
  'Set oRecordSet = oADO.GetRecordSet(oConnection, "SELECT * FROM Sozlesme ORDER BY musAdSoyad ASC")

   Dim mRecordSet As New ADODB.Recordset
   Call mRecordSet.Open(psqlcmd, pConnection, adOpenStatic, adLockOptimistic)
   mRecordSet.Requery
 
   GetRecordSet = mRecordSet
End Function


Public Function GetRecordSetByFileName(ByVal pdbname As String, ByVal psqlcmd As String) As Recordset
  'Returns a RecordSet object from a Connection & an sql string
  
  'EXAMPLE:
  'Dim oADO As New ClassADO
  'Public oRecordSet As Recordset
  'Set oRecordSet = oADO.GetRecordSet(oConnection, "SELECT * FROM Sozlesme ORDER BY musAdSoyad ASC")

   Set mRecordSet = New Recordset
   Call mRecordSet.Open(psqlcmd, GetConnectionAccess(pdbname), adOpenStatic, adLockOptimistic)
   mRecordSet.Requery

   Set GetRecordSetByFileName = mRecordSet
End Function


Public Sub ExecuteCommand(ByVal pdbname As String, ByVal psqlcmd As String)
   'Executes an sql string over a ADODB.Command
   
   'EXAMPLE:
   'Call goAdo.ExecuteCommand(gDBName, sSql)
     
   ' Create a command object.
   Dim cmdADO As New ADODB.Command
   Set cmdADO.ActiveConnection = GetConnectionAccess(pdbname)
   cmdADO.CommandText = psqlcmd
   cmdADO.Execute
End Sub

