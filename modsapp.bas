Attribute VB_Name = "Module1"
Public Registro1 As New ADODB.Recordset
Public Sqlconsulta As String
Public Registro2 As New ADODB.Recordset

Public Xtwips As Integer, Ytwips As Integer
      Public Xpixels As Integer, Ypixels As Integer

      Type FRMSIZE
         Height As Long
         Width As Long
      End Type

      Public RePosForm As Boolean
      Public DoResize As Boolean



    
      Sub Resize_For_Resolution(ByVal SFX As Single, _
       ByVal SFY As Single, MyForm As Form)
      Dim i As Integer
      Dim SFFont As Single

      SFFont = (SFX + SFY) / 2  ' average scale
      ' Size the Controls for the new resolution
      On Error Resume Next  ' for read-only or nonexistent properties
      With MyForm
        For i = 0 To .count - 1
         If TypeOf .Controls(i) Is ComboBox Then   ' cannot change Height
           .Controls(i).Left = .Controls(i).Left * SFX
           .Controls(i).Top = .Controls(i).Top * SFY
           .Controls(i).Width = .Controls(i).Width * SFX
         Else
           .Controls(i).Move .Controls(i).Left * SFX, _
            .Controls(i).Top * SFY, _
            .Controls(i).Width * SFX, _
            .Controls(i).Height * SFY
         End If
           ' Be sure to resize and reposition before changing the FontSize
           .Controls(i).FontSize = .Controls(i).FontSize * SFFont
        Next i
        If RePosForm Then
          ' Now size the Form
          .Move .Left * SFX, .Top * SFY, .Width * SFX, .Height * SFY
        End If
      End With
      End Sub

Public Function ConectarBD()
ConbdSapp.ConnectionString = "driver={MySQL ODBC 8.0 Unicode Driver};SERVER=" & Xipsrv & ";PORT=3306;DATABASE=mmsyssapp;USER=root;PASSWORD=$.Sapp1987;OPTION=3;"
'MySQL ODBC 8.0 Unicode Driver
''ConbdSapp.ConnectionString = "driver={MySQL ODBC 5.1 Driver};SERVER=" & Xipsrv & ";PORT=3306;DATABASE=mmsyssapp;USER=root;PASSWORD=$.Sapp1987;OPTION=3;"

End Function
Public Function ConectarAviso()
ConbdSappAviso.ConnectionString = "driver={MySQL ODBC 8.0 Unicode Driver};SERVER=" & Xipsrv & ";PORT=3306;DATABASE=mmsyssapp;USER=root;PASSWORD=$.Sapp1987;OPTION=3;"
'MySQL ODBC 8.0 Unicode Driver
''ConbdSapp.ConnectionString = "driver={MySQL ODBC 5.1 Driver};SERVER=" & Xipsrv & ";PORT=3306;DATABASE=mmsyssapp;USER=root;PASSWORD=$.Sapp1987;OPTION=3;"

End Function
Public Function ConectarAvisoF()
ConbdSappAvisoF.ConnectionString = "driver={MySQL ODBC 8.0 Unicode Driver};SERVER=" & Xipsrv & ";PORT=3306;DATABASE=mmsyssapp;USER=root;PASSWORD=$.Sapp1987;OPTION=3;"
'MySQL ODBC 8.0 Unicode Driver
''ConbdSapp.ConnectionString = "driver={MySQL ODBC 5.1 Driver};SERVER=" & Xipsrv & ";PORT=3306;DATABASE=mmsyssapp;USER=root;PASSWORD=$.Sapp1987;OPTION=3;"

End Function

Public Function ConectarBD2()
''ConbdSapp2.ConnectionString = "driver={MySQL ODBC 5.1 Driver};SERVER=" & Xipsrv & ";PORT=3306;DATABASE=sappbd;USER=root;PASSWORD=$.Sapp1987;OPTION=3;"
ConbdSapp2.ConnectionString = "driver={MySQL ODBC 8.0 Unicode Driver};SERVER=" & Xipsrv & ";PORT=3306;DATABASE=sappbd;USER=root;PASSWORD=$.Sapp1987;OPTION=3;"

End Function

Public Function ConectarBDVpn()
''ConbdSappvpn.ConnectionString = "driver={MySQL ODBC 5.1 Driver};SERVER=" & Xipsrv & ";PORT=3306;DATABASE=sappfact;USER=root;PASSWORD=$.Sapp1987;OPTION=3;"
ConbdSappvpn.ConnectionString = "driver={MySQL ODBC 8.0 Unicode Driver};SERVER=" & Xipsrv & ";PORT=3306;DATABASE=sappfact;USER=root;PASSWORD=$.Sapp1987;OPTION=3;"

End Function

Public Function ControlUsuario(ByVal opcionmenu As String) As Integer

Dim XrecUsuario As New ADODB.Recordset
Dim Xsqlusua As String
Dim PuedeEntrar As Integer

ConectarBD
ConbdSapp.Open

Xsqlusua = "Select * from usua_permisos where id_usuario =" & Welnrou & " and opcion ='" & opcionmenu & "'"
With XrecUsuario
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlusua, ConbdSapp, , , adCmdText
End With

If XrecUsuario.RecordCount > 0 Then
   ControlUsuario = 1
Else
' se excluyen permisos especiales
   If opcionmenu = "frm_bloqueos" Or opcionmenu = "Command1" Or opcionmenu = "b_afil" Or opcionmenu = "b_infAfiliaciones" Or opcionmenu = "b_anular_reg" Or opcionmenu = "Utilitarios despacho" Or _
      opcionmenu = "Informática" Or opcionmenu = "Despachador Edita" Or opcionmenu = "Especialistas" Or opcionmenu = "Marketing" Or opcionmenu = "Datos tarjetas" Or opcionmenu = "Historial Adm" Then
   Else
      MsgBox "Sin permisos para acceder a esta opción", vbInformation
   End If
   ControlUsuario = 0
End If
XrecUsuario.Close
ConbdSapp.Close


End Function
Function InvokeWebService(strSoap, strSOAPAction, strURL, ByRef xmlResponse) As Boolean
Dim xmlhttp As MSXML2.XMLHTTP30
Dim blnSuccess As Boolean

Set xmlhttp = New MSXML2.XMLHTTP30
xmlhttp.Open "POST", strURL, False
xmlhttp.setRequestHeader "Man", "POST " & strURL & " HTTP/1.1"
xmlhttp.setRequestHeader "Content-Type", "text/xml; charset=utf-8"
xmlhttp.setRequestHeader "SOAPAction", strSOAPAction
Call xmlhttp.send(strSoap)

If xmlhttp.Status = 200 Then
blnSuccess = True
Else
blnSuccess = False
End If

Set xmlResponse = xmlhttp.responseXML
InvokeWebService = blnSuccess
Set xmlhttp = Nothing
End Function

