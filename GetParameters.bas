Attribute VB_Name = "GetParameters"
Public Function getValor(id As Integer) As String
Dim rs As ADODB.Recordset
Dim conn As ADODB.Connection
Set conn = New ADODB.Connection
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.path & "\ws.mdb"
Set rs = New ADODB.Recordset
rs.Open "SELECT valor FROM parametros WHERE id = " & id, conn

getValor = rs("valor")
conn.Close
End Function


Public Function getActivo(id As Integer) As Boolean
Dim rs As ADODB.Recordset
Dim conn As ADODB.Connection

Set conn = New ADODB.Connection
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.path & "\ws.mdb"
Set rs = New ADODB.Recordset
rs.Open "SELECT activado FROM parametros WHERE id = " & id, conn

getActivo = rs("activado")
conn.Close
End Function

Public Function log(descripcion As String)
Dim rs As ADODB.Recordset
Dim conn As ADODB.Connection
'On Error GoTo 0

Set conn = New ADODB.Connection
conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.path & "\ws.mdb"
Set rs = New ADODB.Recordset
rs.Open "INSERT INTO logs (descripcion) " & _
        "VALUES ('" & descripcion & "') ", conn, 3, 3

conn.Close

End Function
