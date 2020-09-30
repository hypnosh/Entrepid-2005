VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private dbUser As String
Private dbPassword As String
Private dbSource As String
Private dbServer As String
Private rowCount As String
Private colCount As String
Public recSet As ADODB.Recordset

Public Function Savedata(strsql As String) As Boolean
'This method to execute insert,update and delete statements in SQL
'It returns true if database is updated otherwise false
On Error GoTo err_handler
Dim noofRecords As Double
gObjConn.Execute strsql, noofRecords
If noofRecords = 0 Then
 Savedata = False
Else
  Savedata = True
End If
Exit Function
err_handler:
 If Err.Number = -2147217873 Then
     MsgBox "Duplicate Record Entry", vbInformation, Err.Description
 End If
Savedata = False
End Function
Public Function getDbConn() As Boolean
On Error GoTo ErrHandler
    
    'Connect to the Database
    Set gObjConn = New ADODB.Connection
     gObjConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\\b112\db\entrepid_beta1.mdb;Persist Security Info=False"
     'gObjConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=E:\ENTREPID\entrepid_beta1.mdb;Persist Security Info=False"
    If gObjConn.State = adStateClosed Then
    gObjConn.Open
    getDbConn = True
    Exit Function
  End If
ErrHandler:
    getDbConn = False
  
End Function
Public Function closeDbConn()
    'Close the database connection
  If gObjConn.State = adStateOpen Then
    gObjConn.Close
  End If
End Function
Public Function executeSQL(strsql As String, Optional errormsg As String) As Clsresultset
'This method is used to getrecords from the database as per sql query passed
 'This method returns a two dimensional array of records
 On Error GoTo err_handler
    Dim objrs As ADODB.Recordset
    Dim columnarray As Variant
    Set objrs = New ADODB.Recordset
    Dim Result As Variant
    Dim gClsresult As Clsresultset
    
    Set recSet = Nothing
    objrs.CursorLocation = adUseClient

    objrs.Open strsql, gObjConn, adOpenKeyset, adLockOptimistic
    
    Set recSet = objrs
    
    If objrs.EOF = True And objrs.BOF Then
        'MsgBox "No records found"
        Set executeSQL = Nothing
        Exit Function
        
    Else
        'get column and rowcount
       Set gClsresult = New Clsresultset
        gClsresult.columnCount = objrs.Fields.Count
        columnCount = objrs.Fields.Count
        gClsresult.rowCount = objrs.RecordCount
        rowCount = objrs.RecordCount
        colCount = objrs.Fields.Count
        ReDim columnarray(objrs.Fields.Count - 1)
        For j = 0 To colCount - 1
            columnarray(j) = objrs.Fields(j).Name
        Next j
        gClsresult.resultArray = objrs.GetRows()
        gClsresult.columnarray = columnarray
        
        
       Set executeSQL = gClsresult
    End If
    Exit Function
err_handler:
  errormsg = CStr(Err.Number) + Err.Description

End Function
Public Function startTrans()
gObjConn.BeginTrans
End Function
Public Function commitTrans()
gObjConn.commitTrans
End Function

Public Function rollbackTrans()
gObjConn.rollbackTrans
End Function


