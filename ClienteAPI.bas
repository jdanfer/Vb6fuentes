Attribute VB_Name = "ClienteAPI"
Function consumirServicio(metodo As String, urlWS As String, body As String) As Object
    On Error GoTo WSError
    Set obj = CreateObject("Microsoft.XMLHTTP")
    obj.Open metodo, urlWS, False
    obj.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    obj.setRequestHeader "User-Agent", "SAPP VB6"
    bodyVal = body
    If (metodo = "GET") Then
        obj.send
    Else
        obj.send bodyVal
    End If
    
    Set consumirServicio = obj
    Exit Function

WSError:
  MsgBox Err.Description & " (servicio de reservas cado o sin conexion, favor, contactarse con computos) " & "body: " & metodo & " " & urlWS & " " & bodyVal
  GetParameters.log Err.Description & " " & "body: " & metodo & " " & urlWS & " " & bodyVal
  On Error GoTo 0
End Function

Function consumirServicio2(metodo As String, urlWS As String, body As String, timeout As Integer) As Object
    ' no imprime mensaje de error, delega, y manda application/json
    Dim auth As String
    Dim clave As String
    Dim rs As ADODB.Recordset
    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection
    Set obj = CreateObject("Msxml2.ServerXMLHTTP")
    
    obj.Open metodo, urlWS, False
    obj.setTimeouts timeout * 1000, timeout * 1000, timeout * 1000, timeout * 1000
    obj.setRequestHeader "Content-Type", "application/json"
    obj.setRequestHeader "User-Agent", "SAPP VB6"
    ' agrego authentication header
    conn.ConnectionString = "dsn=" & Xconexrmt
    conn.Open
    Set rs = New ADODB.Recordset
    rs.Open "SELECT clave FROM usuarios WHERE usuario = '" & WElusuario & "'", conn
    
    clave = rs("clave")
    conn.Close
    
    auth = Base64EncodeString(WElusuario & ":" & clave)
    obj.setRequestHeader "Authorization", "Basic " & auth
    '
    
    bodyVal = body
    If (metodo = "GET") And body = "" Then
        obj.send
    Else
        obj.send bodyVal
    End If
    
    Set consumirServicio2 = obj
    Exit Function

End Function


Function consumirServicioAsync(metodo As String, urlWS As String, body As String) As Object
    ' consume servicio de forma asincrona y no espera la respuesta
    ' no imprime mensaje de error, delega
    Dim auth As String
    Dim clave As String
    Dim rs As ADODB.Recordset
    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection
    Set obj = CreateObject("Msxml2.ServerXMLHTTP")
    
    obj.Open metodo, urlWS, True
    obj.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    obj.setRequestHeader "User-Agent", "SAPP VB6"
    ' agrego authentication header
    conn.ConnectionString = "dsn=" & Xconexrmt
    conn.Open
    Set rs = New ADODB.Recordset
    rs.Open "SELECT clave FROM usuarios WHERE usuario = '" & WElusuario & "'", conn
    
    clave = rs("clave")
    conn.Close
    
    auth = Base64EncodeString(WElusuario & ":" & clave)
    obj.setRequestHeader "Authorization", "Basic " & auth
    '
    bodyVal = body
    If (metodo = "GET") Then
        obj.send
    Else
        obj.send bodyVal
    End If
    
    Set consumirServicioAsync = obj
    Exit Function

End Function

Function consumirServicioAsyncJSON(metodo As String, urlWS As String, body As String) As Object
    ' consume servicio de forma asincrona y no espera la respuesta
    ' no imprime mensaje de error, delega
    Dim auth As String
    Dim clave As String
    Dim rs As ADODB.Recordset
    Dim conn As ADODB.Connection
    Set conn = New ADODB.Connection
    Set obj = CreateObject("Microsoft.XMLHTTP")
    
    obj.Open metodo, urlWS
    obj.setRequestHeader "Content-Type", "application/json"
    obj.setRequestHeader "User-Agent", "SAPP VB6"
    ' agrego authentication header
    conn.ConnectionString = "dsn=" & Xconexrmt
    conn.Open
    Set rs = New ADODB.Recordset
    rs.Open "SELECT clave FROM usuarios WHERE usuario = '" & WElusuario & "'", conn
    
    clave = rs("clave")
    conn.Close
    
    auth = Base64EncodeString(WElusuario & ":" & clave)
    obj.setRequestHeader "Authorization", "Basic " & auth
    '
    bodyVal = body
    If (metodo = "GET") Then
        obj.send
    Else
        obj.send bodyVal
    End If
    
    Set consumirServicioAsyncJSON = obj
    Exit Function

End Function

