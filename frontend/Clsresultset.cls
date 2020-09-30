VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Clsresultset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public resultArray As Variant
Public columnarray As Variant
Public columnCount As Double
Public rowCount As Double
Public Function GetValue(ByVal rownumber As Double, colname As String)
For i = LBound(columnarray) To UBound(columnarray)
   If (Trim(UCase(colname)) = Trim(UCase(columnarray(i)))) Then
          Exit For
      End If
 Next i
 rownumber = rownumber - 1
'GetValue = resultArray(rownumber, i)
  If i > UBound(columnarray) Then
   i = UBound(columnarray)
  End If
 GetValue = Trim(resultArray(i, rownumber))
End Function
