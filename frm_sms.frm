VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm_sms 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes para comunicación vía SMS"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6180
   Icon            =   "frm_sms.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   6180
   StartUpPosition =   1  'CenterOwner
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2880
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Data data_fact 
      Caption         =   "data_fact"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   2400
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_envios 
      Caption         =   "data_envios"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
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
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4080
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Enviar"
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   3840
      Width           =   1095
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   360
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Data data_datos 
      Caption         =   "data_datos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_medicos 
      Caption         =   "data_medicos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
      Height          =   495
      Left            =   5400
      Picture         =   "frm_sms.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Height          =   495
      Left            =   4440
      Picture         =   "frm_sms.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Caption         =   "Datos para informe"
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.CheckBox Check2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Mayor a 15hs."
         Height          =   255
         Left            =   3360
         TabIndex        =   15
         Top             =   2760
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Total de pacientes anotados"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   2760
         Width           =   2655
      End
      Begin VB.Data data_cons 
         Caption         =   "data_cons"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   120
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         ItemData        =   "frm_sms.frx":109E
         Left            =   1560
         List            =   "frm_sms.frx":10CC
         Style           =   2  'Dropdown List
         TabIndex        =   11
         ToolTipText     =   "Selección de zona (OPCIONAL)"
         Top             =   2160
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frm_sms.frx":118E
         Left            =   1560
         List            =   "frm_sms.frx":119E
         TabIndex        =   6
         Text            =   "Combo2"
         Top             =   1680
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frm_sms.frx":11C8
         Left            =   1560
         List            =   "frm_sms.frx":11D8
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   3855
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   3240
         TabIndex        =   3
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox md 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Opcional ZONAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   2160
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "Opcional Gpos."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Origen:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "FECHA:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   13
      Top             =   3240
      Width           =   3975
   End
End
Attribute VB_Name = "frm_sms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
If Combo1.ListIndex = 2 Then
   Combo2.Clear
   Combo2.AddItem "SIN OPCIONES"
   Combo4.Clear
Else
   Combo2.Clear
   Combo2.AddItem "CCOU"
   Combo2.AddItem "SMI"
   Combo2.AddItem "H.EVANGELICO"
   Combo2.AddItem "UNIVERSAL"
   Combo2.AddItem "CASA DE GALICIA"
   Combo2.AddItem "SEMM"
   Combo2.AddItem "CASH"
   Combo2.AddItem "CPS"
   Combo2.AddItem "CASMU"
   Combo2.AddItem "AMBULATORIOS"
   Combo2.AddItem "EMERGENCIA"
   Combo2.AddItem "SELECCION"
End If

End Sub


Private Sub Command1_Click()
Dim Xlinn, Xhoracomi As String
Dim Xfecdes, Xenvio As Date
Xfecdes = Date + 2
Xenvio = Date + 1
Dim Xlac As String
Xlac = ""
Xhoracomi = "X"
'https://api.mensajeroautomatico.com.uy/?token=3e690a8df1e92b4a3160dc4a86d03047&
'pin=1234&texto=hola+como+estan&destino=097136278&hora=noche

On Error GoTo Queerftp
Command1.Enabled = False

If data_inf.Recordset.RecordCount > 0 Then
   data_inf.Recordset.MoveFirst
   Do While Not data_inf.Recordset.EOF
      data_inf.Recordset.Delete
      data_inf.Recordset.MoveNext
   Loop
End If
If Combo1.Text = "Padrón Activos" Then
   MsgBox "Opción no habilitada"
End If
If Combo1.Text = "Padrón Bajas" Then
   MsgBox "Opción no habilitada"
End If
frm_sms.MousePointer = 11
If Combo1.Text = "Especialistas" Then
'   data_medicos.RecordSource = "Select * from medicos where med_nombre ='" & Combo3.Text & "'"
'   data_medicos.Refresh
'      data_datos.RecordSource = "Select * from lineas where cod_prod =" & data_medicos.Recordset("med_cod")
'      data_datos.Refresh
'   data_cons.RecordSource = "Select * from mant_sol where cl_fnac =#" & Format(Xfecdes, "yyyy/mm/dd") & "# order by cl_fax,cl_ruc"
   data_datos.DatabaseName = ""
   data_datos.Connect = "ODBC;DSN=sappespecial;"
   If Check1.Value = 1 Then
      data_datos.RecordSource = "Select * from t_cabfechas where especial in ('AFILIACIONES') and cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cancela not in ('SI')"
      data_datos.Refresh
   Else
      data_datos.RecordSource = "Select * from t_cabfechas where especial in ('AFILIACIONES') and cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cancela not in ('SI')"
      data_datos.Refresh
   End If
   Open App.path & "\smssapp.txt" For Output As #1
   If data_datos.Recordset.RecordCount > 0 Then
      data_datos.Recordset.MoveLast
      data_datos.Recordset.MoveFirst
      Do While Not data_datos.Recordset.EOF
         data_cons.RecordSource = "Select * from t_fechas where cod_cons =" & data_datos.Recordset("id") & " and base =" & data_datos.Recordset("base") & " and cel_pac is not null order by cod_med,fecha,base"
         data_cons.Refresh
         If data_cons.Recordset.RecordCount > 0 Then
            data_cons.Recordset.MoveLast
            data_cons.Recordset.MoveFirst
            Xlac = "0099"
            Do While Not data_cons.Recordset.EOF
                  If Mid(Trim(data_cons.Recordset("cel_pac")), 1, 2) = "09" And Len(Trim(data_cons.Recordset("cel_pac"))) = 9 Then
                     Xlinn = Trim(data_cons.Recordset("cel_pac")) + ";"
                     Xlinn = Xlinn & "1;"
                     If data_cons.Recordset("especial") = "FISIOTERAPIA" Or _
                        data_cons.Recordset("especial") = "PEDIATRIA" Or _
                        data_cons.Recordset("especial") = "ODONTOLOGIA" Or _
                        data_cons.Recordset("especial") = "SICOLOGIA" Then
                        Xlinn = Xlinn & data_cons.Recordset("nom_pac") & " le recordamos agenda para " & data_cons.Recordset("especial") & " el dia " & Format(data_cons.Recordset("fecha"), "dd/mm/yyyy") & " Hora: " & data_cons.Recordset("hora")
                     End If
                     Xlinn = Xlinn & " Cancelacion al 097215427;"
                     Xlinn = Xlinn & ";;;"
'                     If Check2.Value = 1 Then
'                        if data_cons.Recordset("hora") >='15:01' then
                        
                     Print #1, Xlinn
                     data_inf.Recordset.AddNew
                     data_inf.Recordset("cl_apellid") = Mid(data_cons.Recordset("nom_pac"), 1, 40)
                     data_inf.Recordset("cl_dpto") = Mid(data_cons.Recordset("especial"), 1, 12)
                     data_inf.Recordset("cl_fnac") = CDate(data_cons.Recordset("fecha"))
                     data_inf.Recordset("cl_nombre") = Mid(data_cons.Recordset("nom_med"), 1, 30)
                     data_inf.Recordset.Update
                  End If
               data_cons.Recordset.MoveNext
            Loop
         End If
         data_datos.Recordset.MoveNext
         Xhoracomi = "X"
      Loop
      Close #1
      Command3_Click
      frm_sms.MousePointer = 0
      MsgBox "Proceso terminado"
   Else
      Close #1
      frm_sms.MousePointer = 0
      MsgBox "No hay datos para procesar"
   End If
End If
If Combo1.Text = "Multiconsultas" Then
   Dim Xlamat As Double
   Dim xcantt As Integer
   Dim Xfechatres, Xfechauno As Date
   Xfechatres = Date - 3
   Xfechauno = Date - 1
   data_datos.DatabaseName = ""
   data_datos.Connect = "ODBC;DSN=sappnew;"
   data_datos.RecordSource = "Select * from llamado where fecha >=#" & Format(Xfechatres, "yyyy/mm/dd") & "# and fecha <=#" & Format(Xfechauno, "yyyy/mm/dd") & "# and matric not in (0) order by matric"
   data_datos.Refresh
    
   data_inf.DatabaseName = App.path & "\informes.mdb"
   data_inf.RecordSource = "Select * from inflla"
   data_inf.Refresh
   If data_inf.Recordset.RecordCount > 0 Then
      data_inf.Recordset.MoveFirst
      Do While Not data_inf.Recordset.EOF
         data_inf.Recordset.Delete
         data_inf.Recordset.MoveNext
      Loop
   End If
   If data_datos.Recordset.RecordCount > 0 Then
      data_datos.Recordset.MoveFirst
      Xlamat = data_datos.Recordset("matric")
      Do While Not data_datos.Recordset.EOF
         If data_datos.Recordset("matric") = Xlamat Then
            xcantt = xcantt + 1
         Else
            If xcantt >= 4 Then
               data_datos.Recordset.MovePrevious
               data_cons.DatabaseName = ""
               data_cons.Connect = "ODBC;DSN=sappnew;"
               data_cons.RecordSource = "Select * from llamado where fecha >=#" & Format(Xfechatres, "yyyy/mm/dd") & "# and fecha <=#" & Format(Xfechauno, "yyyy/mm/dd") & "# and matric =" & data_datos.Recordset("matric")
               data_cons.Refresh
               If data_cons.Recordset.RecordCount > 0 Then
                  data_cons.Recordset.MoveFirst
                  Do While Not data_cons.Recordset.EOF
                     data_inf.Recordset.AddNew
                     data_inf.Recordset("matric") = data_cons.Recordset("matric")
                     data_inf.Recordset("fecha") = data_cons.Recordset("fecha")
                     data_inf.Recordset("hora") = data_cons.Recordset("hora")
                     data_inf.Recordset("nombre") = data_cons.Recordset("nombre")
                     data_inf.Recordset("categ") = data_cons.Recordset("categ")
                     data_inf.Recordset("edad") = data_cons.Recordset("edad")
                     data_inf.Recordset("codmot") = data_cons.Recordset("codmot")
                     data_inf.Recordset("obsmot") = data_cons.Recordset("obsmot")
                     data_inf.Recordset("nommed") = data_cons.Recordset("nommed")
                     data_inf.Recordset.Update
                     data_cons.Recordset.MoveNext
                  Loop
               End If
               data_datos.Recordset.MoveNext
            End If
            xcantt = 1
         End If
         Xlamat = data_datos.Recordset("matric")
         data_datos.Recordset.MoveNext
      Loop
   End If
   
   data_inf.Refresh
   If data_inf.Recordset.RecordCount > 0 Then
      Open App.path & "\multiconsulta.txt" For Output As #1
      data_inf.Recordset.MoveFirst
      Print #1, "FECHA-------HORA---CONVENIO-EDAD-MATRICULA---NOMBRE---------------MOTIVO ULTIMA CONSULTA----"
      Do While Not data_inf.Recordset.EOF
         Xlinn = Format(data_inf.Recordset("fecha"), "dd/mm/yyyy") & "   " & data_inf.Recordset("hora") & "  " & data_inf.Recordset("categ") & "  " & Trim(str(data_inf.Recordset("edad"))) & " " & Trim(str(data_inf.Recordset("matric"))) & " " & data_inf.Recordset("nombre") & "    " & data_inf.Recordset("obsmot")
         Print #1, Xlinn
         data_inf.Recordset.MoveNext
      Loop
      Close #1
      
      Dim MenCorreo As String
      Dim oMail As Class1
         
      Set oMail = New Class1
      With oMail
           .servidor = "smtp.gmail.com"
           .puerto = 465
           .UseAuntentificacion = True
           .ssl = True
           .Usuario = "despachosapp@gmail.com"
           .PassWord = "sapp1987"
           .Asunto = "Llamados multiconsultas últimos 3 días " & Format(Date, "dd/mm/yyyy")
           .de = "despachosapp@gmail.com"
           .para = "directortecnico@sapp.com.uy; subdirectortecnico@sapp.com.uy; jefedepartamentoti@sapp.com.uy"
''            .para = "sappjorge@hotmail.com; jefedepartamentoti@sapp.com.uy"
           .Adjunto = App.path & "\multiconsulta.txt"
           .Mensaje = "Multiconsultas por cliente."
           .Enviar_Backup ' manda el mail
      End With
      Set oMail = Nothing
      frm_sms.MousePointer = 0
      MsgBox "Proceso terminado"
   Else
      frm_sms.MousePointer = 0
      MsgBox "No hay registros", vbInformation
   End If
End If
Command1.Enabled = True
frm_sms.MousePointer = 0

Exit Sub

Queerftp:
          If Err.Number = 3155 Then
             MsgBox "Error al grabar"
          Else
             MsgBox "ERROR: " & Err.Description, Err.Number
          End If
         

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
 Dim El_Host As String
 Dim Xlocal, Xremoto As String
 Xlocal = App.path & "\smssapp.txt"
 Xremoto = "/Sms.txt"
'    List1.AddItem " ..Subiendo archivo "

    El_Host = "ftp://192.168.20.22"

    With Inet1

'        .URL = El_Host
        'nombre de usuario y password de la cuanta FTP
'        .UserName = "sapp"
'        .UserName = "jfernan"
'        .PassWord = "4Pn6ih8DsS"
'        .PassWord = "Sapp1987"
        'Escribe el fichero en el servidor con el comando Put
'        .Execute .url El_Host, "Put " & Xlocal & " " & Xremoto)
        .url = "ftp://190.64.83.74"
        .Protocol = icFTP
        .RemoteHost = "190.64.83.74"
        .UserName = "sapp"
        .PassWord = "4Pn6ih8DsS"
        .Execute .url, "put " + Xlocal + " " + Xremoto
        Do While .StillExecuting
            DoEvents
        Loop
'        Dim texto As String
'        UploadFile = (.ResponseCode = 0)
'        .Execute , "quit"                   'Logoff
'        DoEvents

     End With
     data_envios.RecordSource = "envsms"
     data_envios.Refresh
     data_inf.Refresh
     If data_inf.Recordset.RecordCount > 0 Then
        data_inf.Recordset.MoveFirst
        Do While Not data_inf.Recordset.EOF
           data_envios.Recordset.AddNew
           data_envios.Recordset("fecha") = Date
           data_envios.Recordset("hora") = Format(Time, "HH:mm")
           data_envios.Recordset("usuario") = WElusuario
           data_envios.Recordset("nompac") = data_inf.Recordset("cl_apellid")
           data_envios.Recordset("feccons") = data_inf.Recordset("cl_fnac")
           data_envios.Recordset("especial") = data_inf.Recordset("cl_dpto")
           data_envios.Recordset("resultado") = Label5.Caption
           data_envios.Recordset.Update
           data_inf.Recordset.MoveNext
        Loop
        data_envios.RecordSource = "Select * from envsms order by fecha"
        data_envios.Refresh
        cr1.ReportFileName = App.path & "\infenvsms.rpt"
        cr1.ReportTitle = "Informe del resultado de envío de SMS Fecha:" & md.Text & " hasta:" & mh.Text
        cr1.Action = 1
        
     End If

End Sub

Private Sub Form_Load()
data_medicos.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_datos.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cons.Connect = "ODBC;DSN=sappespecial;"
data_inf.DatabaseName = App.path & "\informes.mdb"
data_inf.RecordSource = "infcli"
data_inf.Refresh
data_envios.DatabaseName = App.path & "\envsms.mdb"
data_fact.Connect = "odbc;dsn=" & Xconexrmt & ";"

End Sub

Private Sub Inet1_StateChanged(ByVal state As Integer)
Select Case state

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

       Case Else: Label5.Caption = " Estado -> " & Format$(state)
    End Select

    DoEvents



End Sub

Private Sub md_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mh.SetFocus
End If

End Sub

Private Sub mh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo1.SetFocus
End If

End Sub
