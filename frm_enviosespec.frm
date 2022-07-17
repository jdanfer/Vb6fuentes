VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frm_envioespec 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Envío de correos de reservas de especialistas"
   ClientHeight    =   2025
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8955
   Icon            =   "frm_enviosespec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   8955
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5640
      Top             =   1560
   End
   Begin VB.Data data_envios 
      Caption         =   "data_envios"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   480
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Enviar"
      Height          =   495
      Left            =   6120
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   6120
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Data data_cons 
      Caption         =   "data_cons"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data data_datos 
      Caption         =   "data_datos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_queenvio 
      Caption         =   "data_queenvio"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   840
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   480
      Visible         =   0   'False
      Width           =   4935
   End
   Begin VB.Data data_lista 
      Caption         =   "data_lista"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1200
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data data_ctrl 
      Caption         =   "data_ctrl"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   5760
      Top             =   0
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1440
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Height          =   495
      Left            =   360
      Picture         =   "frm_enviosespec.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Enviar los correos con los archivos de la lista"
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "ENVIO DE CORREOS DE CONFIRMACION RESERVA DE ESPECIALISTAS"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
   End
End
Attribute VB_Name = "frm_envioespec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim MenCorreo As String
Dim oMail As Class1
Dim Xcanenvio As Long
Dim XX As Integer
Dim XFechaEnvio As Date
XFechaEnvio = Date + 2
On Error GoTo Enenvio

'Data1.RecordSource = "Select * from correos where fecha_env is null"
Data1.RecordSource = "select * from t_fechas where fecha_cons =#" & Format(XFechaEnvio, "yyyy/mm/dd") & "# and fecha_correo is null and cel_pac is not null and ced_pac is not null"
Data1.Refresh
Timer1.Enabled = False

'If Data1.Recordset.RecordCount > 0 Then
'   Set oMail = New Class1
'   Data1.Recordset.MoveFirst
'   Do While Not Data1.Recordset.EOF
'      If IsNull(Data1.Recordset("direccion")) = False Then
'         With oMail
'              .servidor = "vera.com.uy"
'              .puerto = 465
'              .UseAuntentificacion = True
'              .ssl = True
'              .Usuario = "reservasdesapp@vera.com.uy"
'              .PassWord = "sapp1987"
'              .Asunto = Data1.Recordset("asunto")
'              .de = "reservasdesapp@vera.com.uy"
'              .para = Data1.Recordset("direccion")
'              .Adjunto = ""
'              .Mensaje = Data1.Recordset("mensaje")
'              .Enviar_Backup ' manda el mail
'         End With
''''          '  MsgBox "Correo enviado..."
'         Data1.Recordset.Edit
'         Data1.Recordset("fecha_env") = Format(Date, "dd/mm/yyyy")
'         Data1.Recordset("hora_env") = Format(Time, "HH:mm")
'         Data1.Recordset.Update
'      End If
'      Data1.Recordset.MoveNext
'   Loop
'   Set oMail = Nothing
'End If

'If Format(Time, "HH:mm") > "18:00" Then
'    If Format(data_ctrl.Recordset("fecha_env"), "yyyy/mm/dd") < Format(Date, "yyyy/mm/dd") Then
'       Dim Xfecped, Xfecesp As Date
'       Dim Xlleno As Integer
'       Xfecped = Date + 7
'       Xfecesp = Date + 30
'       Data1.RecordSource = "Select * from t_cabfechas where especial ='" & "PEDIATRIA" & "' and cdate(fecha) >=#" & Format(Date, "yyyy/mm/dd") & "# and cdate(fecha) <=#" & Format(Xfecped, "yyyy/mm/dd") & "# and enviado is null"
'       Data1.Refresh
'       If Data1.Recordset.RecordCount > 0 Then
''          Data1.Recordset.MoveFirst
'          Do While Not Data1.Recordset.EOF
'             data_lista.RecordSource = "select * from t_fechas where fecha ='" & Data1.Recordset("fecha") & "' and cod_cons =" & Data1.Recordset("cod_cons") & " and nom_pac is null order by nro"
'             data_lista.Refresh
'             If data_lista.Recordset.RecordCount > 0 Then
'             Else
'                data_queenvio.Recordset.AddNew
'                data_queenvio.Recordset("fecha_env") = Format(Date, "dd/mm/yyyy")
'                data_queenvio.Recordset("hora_env") = Format(Time, "HH:mm")
'                data_queenvio.Recordset("fecha_cons") = Format(Data1.Recordset("fecha"), "dd/mm/yyyy")
'                data_queenvio.Recordset("hora") = Data1.Recordset("hora")
'                data_queenvio.Recordset("base") = Data1.Recordset("base")
'                data_queenvio.Recordset("especial") = Data1.Recordset("especial")
'                data_queenvio.Recordset("medico") = Mid(Data1.Recordset("nom_med"), 1, 50)
'                data_queenvio.Recordset.Update
'                Data1.Recordset.Edit
'                Data1.Recordset("enviado") = "SI"
'                Data1.Recordset.Update
'             End If
'             Data1.Recordset.MoveNext
'          Loop
'          Data1.RecordSource = "Select * from t_cabfechas where especial not in ('PEDIATRIA','AFILIACIONES') and cdate(fecha) >=#" & Format(Date, "yyyy/mm/dd") & "# and cdate(fecha) <=#" & Format(Xfecesp, "yyyy/mm/dd") & "# and enviado is null"
'          Data1.Refresh
'          If Data1.Recordset.RecordCount > 0 Then
'             Data1.Recordset.MoveFirst
'             Do While Not Data1.Recordset.EOF
'                data_lista.RecordSource = "select * from t_fechas where fecha ='" & Data1.Recordset("fecha") & "' and cod_cons =" & Data1.Recordset("cod_cons") & " and nom_pac is null order by nro"
'                data_lista.Refresh
'                If data_lista.Recordset.RecordCount > 0 Then
'                Else
'                   data_queenvio.Recordset.AddNew
'                   data_queenvio.Recordset("fecha_env") = Format(Date, "dd/mm/yyyy")
'                   data_queenvio.Recordset("hora_env") = Format(Time, "HH:mm")
'                   data_queenvio.Recordset("fecha_cons") = Format(Data1.Recordset("fecha"), "dd/mm/yyyy")
'                   data_queenvio.Recordset("hora") = Data1.Recordset("hora")
'                   data_queenvio.Recordset("base") = Data1.Recordset("base")
'                   data_queenvio.Recordset("especial") = Data1.Recordset("especial")
'                   data_queenvio.Recordset("medico") = Mid(Data1.Recordset("nom_med"), 1, 50)
'                   data_queenvio.Recordset.Update
'                   Data1.Recordset.Edit
'                   Data1.Recordset("enviado") = "SI"
'                   Data1.Recordset.Update
'                End If
'                Data1.Recordset.MoveNext
'             Loop
'          End If
'          data_queenvio.RecordSource = "Select * from ctrlenv where fecha_env =#" & Format(Date, "yyyy/mm/dd") & "#"
'          data_queenvio.Refresh
'          If data_queenvio.Recordset.RecordCount > 0 Then
'             data_queenvio.Recordset.MoveFirst
'             Text1.Text = "ESPECIALISTAS QUE TIENEN CONSULTA COMPLETA" & Chr(13)
'             Do While Not data_queenvio.Recordset.EOF
'                Text1.Text = Text1.Text & data_queenvio.Recordset("medico") & " " & data_queenvio.Recordset("especial") & " FECHA CONS:" & data_queenvio.Recordset("fecha_cons") & " BASE: " & data_queenvio.Recordset("base") & Chr(13)
'                data_queenvio.Recordset.MoveNext
'             Loop
'
'             Set oMail = New Class1
'             With oMail
'''        '                 .servidor = "adinet.com.uy"
'                  .servidor = "vera.com.uy"
'                  .puerto = 465
'                  .UseAuntentificacion = True
'                  .ssl = True
'                  .Usuario = "reservasdesapp@vera.com.uy"
'                  .PassWord = "sapp1987"
'                  .Asunto = "FECHAS DE ESPECIALISTAS COMPLETAS"
'                  .de = "reservasdesapp@vera.com.uy"
'                  .para = "jefedepartamentoti@sapp.com.uy; subdirectortecnico@sapp.com.uy; directortecnico@sapp.com.uy"
'                  .Adjunto = ""
'                  .Mensaje = Text1.Text
'                  .Enviar_Backup ' manda el mail
'             End With
''''                  '  MsgBox "Correo enviado..."
'             Set oMail = Nothing
'          End If
         
'       End If
       
       Dim Xlinn, Xhoracomi As String
'       Dim Xfecdes, Xenvio As Date
'       Xfecdes = Date + 2
'       Xenvio = Date + 1
       Dim Xlac As String
       Xlac = ""
'       Xhoracomi = "X"

'''''https://api.mensajeroautomatico.com.uy/?token=3e690a8df1e92b4a3160dc4a86d03047&
'''''pin=1234&texto=hola+como+estan&destino=097136278&hora=noche

'       If data_inf.Recordset.RecordCount > 0 Then
'          data_inf.Recordset.MoveFirst
'          Do While Not data_inf.Recordset.EOF
'             data_inf.Recordset.Delete
'             data_inf.Recordset.MoveNext
'          Loop
'       End If
   
'       data_cons.Connect = "ODBC;DSN=sappespecial;"
   
'       data_datos.Connect = "ODBC;DSN=sappespecial;"
'       data_datos.RecordSource = "Select * from t_cabfechas where especial not in ('AFILIACIONES') and cdate(fecha) >=#" & Format(Xfecdes, "yyyy/mm/dd") & "# and cdate(fecha) <=#" & Format(Xfecdes, "yyyy/mm/dd") & "# and cancela not in ('SI')"
'       data_datos.Refresh
       
       If Data1.Recordset.RecordCount > 0 Then
          Open App.Path & "\smssapp.txt" For Output As #1
'          data_datos.Recordset.MoveLast
'          data_datos.Recordset.MoveFirst
          Xlac = "0099"
          Data1.Recordset.MoveFirst
          Do While Not Data1.Recordset.EOF
             If Len(Data1.Recordset("cel_pac")) >= 9 Then
                If Xlac = Data1.Recordset("ced_pac") Then
                Else
                   If Mid(Trim(Data1.Recordset("cel_pac")), 1, 2) = "09" And Len(Trim(Data1.Recordset("cel_pac"))) = 9 Then
                      Xlinn = Trim(Data1.Recordset("cel_pac")) + ";"
                      Xlinn = Xlinn & "1;"
                      If Data1.Recordset("especial") = "FISIOTERAPIA" Or _
                         Data1.Recordset("especial") = "PEDIATRIA" Or Data1.Recordset("especial") = "MED.GRAL." Or _
                         Data1.Recordset("especial") = "ODONTOLOGIA" Or _
                         Data1.Recordset("especial") = "SICOLOGIA" Then
                         Xlinn = Xlinn & Data1.Recordset("nom_pac") & " SAPP le RECUERDA consulta con:" & Data1.Recordset("especial")
                         Xlinn = Xlinn & " el dia " & Format(Data1.Recordset("fecha"), "dd/mm/yyyy") & " Hora: " & Data1.Recordset("hora")
                      Else
                         Xlinn = Xlinn & Data1.Recordset("nom_pac") & " SAPP le RECUERDA consulta con:" & Data1.Recordset("especial")
                         Xlinn = Xlinn & " el dia " & Format(Data1.Recordset("fecha"), "dd/mm/yyyy") & " Hora: " & Data1.Recordset("hora_com") & " NRO." & Trim(Str(Data1.Recordset("nro")))
                      End If
                      If Data1.Recordset("base") = 1 Then
                         Xlinn = Xlinn & " en P.Plata;"
                      End If
                      If Data1.Recordset("base") = 2 Then
                         Xlinn = Xlinn & " en Floresta;"
                      End If
                      If Data1.Recordset("base") = 3 Then
                         Xlinn = Xlinn & " en Salinas Base;"
                      End If
                      If Data1.Recordset("base") = 4 Then
                         Xlinn = Xlinn & " en Atlantida;"
                      End If
                      If Data1.Recordset("base") = 6 Then
                         Xlinn = Xlinn & " en Pando;"
                      End If
                      If Data1.Recordset("base") = 8 Then
                         Xlinn = Xlinn & " en Suarez;"
                      End If
                      If Data1.Recordset("base") = 11 Then
                         Xlinn = Xlinn & " en Bs.Bs. Base;"
                      End If
                      If Data1.Recordset("base") = 12 Then
                         Xlinn = Xlinn & " en San Jacinto;"
                      End If
                      If Data1.Recordset("base") = 13 Then
                         Xlinn = Xlinn & " en Tala;"
                      End If
                      If Data1.Recordset("base") = 16 Then
                         Xlinn = Xlinn & " en Toledo Sede;"
                      End If
                      If Data1.Recordset("base") = 17 Then
                         Xlinn = Xlinn & " en Barros Bs;"
                      End If
                      If Data1.Recordset("base") = 18 Then
                         Xlinn = Xlinn & " en Salinas Sede;"
                      End If
                      Xlinn = Xlinn & ";;;"
                      Print #1, Xlinn
                      data_inf.Recordset.AddNew
                      data_inf.Recordset("cl_apellid") = Mid(Data1.Recordset("nom_pac"), 1, 40)
                      data_inf.Recordset("cl_dpto") = Mid(Data1.Recordset("especial"), 1, 12)
                      data_inf.Recordset("cl_fnac") = CDate(Data1.Recordset("fecha"))
                      data_inf.Recordset("cl_nombre") = Mid(Data1.Recordset("nom_med"), 1, 30)
                      data_inf.Recordset.Update
                   End If
                End If
             End If
             Xlac = Data1.Recordset("ced_pac")
             Data1.Recordset.MoveNext
          Loop
          Close #1

''' acá comienza fonasa
'      Command3_Click
          Command2_Click
'''''          Timer2.Enabled = True
'       Else
'          data_ctrl.Recordset.Edit
'          data_ctrl.Recordset("fecha_env") = Format(Date, "dd/mm/yyyy")
'          data_ctrl.Recordset("hora_env") = Format(Time, "HH:mm")
'          data_ctrl.Recordset.Update
       Else
          Timer1.Enabled = True
       End If
'''   End If
'''End If
''''Timer1.Enabled = True
Exit Sub

Enenvio:
        If Err.Number = 91 Then
           MsgBox "Error al enviar el correo, verifique si continúa en ejecución el programa correoespec.exe", vbInformation
        Else
           MsgBox "Error al enviar el correo, verifique si continúa en ejecución el programa correoespec.exe", vbInformation
        End If

End Sub

Private Sub DBGrid1_DblClick()


End Sub

Private Sub Command2_Click()
 Dim El_Host As String
 Dim Xlocal, Xremoto As String
 Xlocal = App.Path & "\smssapp.txt"
 Xremoto = "/Sms.txt"
 On Error GoTo Elsms
'    List1.AddItem " ..Subiendo archivo "
'''''  Timer2.Enabled = False
'    El_Host = "ftp://192.168.20.22"
    El_Host = "ftp://192.168.10.25"

    With Inet1

        .URL = "ftp://190.64.83.74" 'servidor del mensajeroautomatico
        .Protocol = icFTP
        .RemoteHost = "190.64.83.74"
        .UserName = "sapp"
        .PassWord = "4Pn6ih8DsS"
        .Execute .URL, "put " + Xlocal + " " + Xremoto
        Do While .StillExecuting
            DoEvents
        Loop
'        Dim texto As String
'        UploadFile = (.ResponseCode = 0)
'        .Execute , "quit"                   'Logoff
'        DoEvents

     End With

     data_envios.DatabaseName = App.Path & "\envsms.mdb"
     data_envios.RecordSource = "envsms"
     data_envios.Refresh
'     data_inf.Refresh
'     If data_inf.Recordset.RecordCount > 0 Then
'        data_inf.Recordset.MoveFirst
'        Do While Not data_inf.Recordset.EOF
           data_envios.Recordset.AddNew
           data_envios.Recordset("fecha") = Date
           data_envios.Recordset("hora") = Format(Time, "HH:mm")
           data_envios.Recordset("usuario") = WElusuario
           data_envios.Recordset("nompac") = "Envío de SMS"
           data_envios.Recordset("resultado") = Label5.Caption
           data_envios.Recordset.Update
           'data_inf.Recordset.MoveNext
        ''Loop
     ''End If
     
     
'       data_ctrl.Recordset.Edit
'       data_ctrl.Recordset("fecha_env") = Format(Date, "dd/mm/yyyy")
'       data_ctrl.Recordset("hora_env") = Format(Time, "HH:mm")
'       data_ctrl.Recordset.Update
     
Timer1.Enabled = True
     
Exit Sub

Elsms:
      If Err.Number = 91 Then
         MsgBox "Error al subir el archivo de SMS " & Err.Description & Err.Number, vbInformation
      Else
         MsgBox "Error al subir el archivo de SMS " & Err.Description & Err.Number, vbInformation
      End If
End Sub

Private Sub Form_Load()
Data1.Connect = "ODBC;DSN=sappnew;"
Xenvia = 0
Timer1.Enabled = True
data_ctrl.DatabaseName = App.Path & "\ctrlenv.mdb"
data_ctrl.RecordSource = "ctrlenv"
data_ctrl.Refresh
data_queenvio.DatabaseName = App.Path & "\ctrlqueenvio.mdb"
data_queenvio.RecordSource = "ctrlenv"
data_queenvio.Refresh

data_lista.Connect = "ODBC;DSN=sappnew;"

data_inf.DatabaseName = App.Path & "\informes.mdb"
data_inf.RecordSource = "infcli"
data_inf.Refresh

End Sub

Private Sub List1_Click()

End Sub

Private Sub Inet1_StateChanged(ByVal State As Integer)
Select Case State

        'Dependiendo del valor recibido de State _
         muestra en el List1 la información de estado

        Case 0: Label5.Caption = " Nothing "
        Case 1: Label5.Caption = " Resolviendo Host "
        Case 2: Label5.Caption = " Host Resuelto "
        Case 3: Label5.Caption = " ..Conectando a: "
        Case 4: Label5.Caption = ".. Conectado a "
        Case 5: Label5.Caption = " Petición"
        Case 6: Label5.Caption = " ..enviando petición"
        Case 7: Label5.Caption = " Recibiendo Respuesta "
        Case 8: Label5.Caption = " Respuesta recibida "
        Case 9: Label5.Caption = " ..Desconectando "
        Case 10: Label5.Caption = " Estado : Desconectado"
        Case 11: Label5.Caption = " Error: " & Inet1.ResponseInfo
        Case 12: Label5.Caption = Inet1.ResponseInfo

       Case Else: Label5.Caption = " Estado -> " & Format$(State)
    End Select

    DoEvents


End Sub

Private Sub Timer1_Timer()
On Error GoTo Entimer

If Xenvia < 19 Then
   Xenvia = Xenvia + 1
Else
   If Xenvia = 19 Then
      Xenvia = 0
      Command1_Click
   End If
End If

Exit Sub

Entimer:
        If Err.Number = 91 Then
           MsgBox "Error en proceso automático, verifique si continúa en ejecución el programa correoespec.exe", vbInformation
        Else
           MsgBox "Error en proceso automático, verifique si continúa en ejecución el programa correoespec.exe", vbInformation
        End If
End Sub

Private Sub Timer2_Timer()

Command2_Click

End Sub
