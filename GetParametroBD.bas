Attribute VB_Name = "GetParametroBD"
Public Function getValor(id As Integer) As String
Dim rs As ADODB.Recordset
Dim conn As ADODB.Connection
Set conn = New ADODB.Connection
conn.ConnectionString = "dsn=" & Xconexrmt
conn.Open
Set rs = New ADODB.Recordset
rs.Open "SELECT valor FROM parametros WHERE id = " & id, conn

getValor = rs("valor")
conn.Close
End Function

Public Function getValorFromParametro(parametro As String) As String
Dim rs As ADODB.Recordset
Dim conn As ADODB.Connection
Set conn = New ADODB.Connection
conn.ConnectionString = "dsn=" & "sappnew"
conn.Open
Set rs = New ADODB.Recordset
rs.Open "SELECT valor FROM parametros WHERE parametro = '" & parametro & "'", conn

getValorFromParametro = rs("valor")
conn.Close
End Function

Public Function getParametroDesc(parametro As String) As String
Dim rs As ADODB.Recordset
Dim conn As ADODB.Connection
Set conn = New ADODB.Connection
conn.ConnectionString = "dsn=" & Xconexrmt
conn.Open
Set rs = New ADODB.Recordset
rs.Open "SELECT parametro_desc FROM parametros WHERE parametro = '" & parametro & "'", conn

getParametroDesc = rs("parametro_desc")
conn.Close
End Function

Public Function getActivo(id As Integer) As Boolean
Dim rs As ADODB.Recordset
Dim conn As ADODB.Connection

Set conn = New ADODB.Connection
conn.ConnectionString = "dsn=" & Xconexrmt
conn.Open
Set rs = New ADODB.Recordset
rs.Open "SELECT activado FROM parametros WHERE id = " & id, conn

getActivo = rs("activado")
conn.Close
End Function

