VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_afilbusca 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar afiliaciones"
   ClientHeight    =   7230
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11025
   Icon            =   "frm_afilbusca.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   11025
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command6 
      Caption         =   "Correo Cancela"
      Height          =   375
      Left            =   3960
      TabIndex        =   20
      Top             =   6720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data data_abm 
      Caption         =   "data_abm"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox t_cedbusca 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   405
      Left            =   7440
      TabIndex        =   19
      ToolTipText     =   "Digitar todos los números (Ejemplo: para la cédula 1234567-8 digitar 12345678) y luego ENTER "
      Top             =   6720
      Width           =   2055
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ver afiliaciones canceladas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   6720
      Width           =   3375
   End
   Begin VB.Data data_cancelar 
      Caption         =   "data_cancelar"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton b_cancelar 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   9360
      Picture         =   "frm_afilbusca.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Anular afiliación"
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton b_infAfiliaciones 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   10320
      Picture         =   "frm_afilbusca.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Generar informe en planilla excel"
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton b_correo 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   10200
      Picture         =   "frm_afilbusca.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Enviar correo con copia del contrato a la cuenta de correo del titular"
      Top             =   6240
      Width           =   615
   End
   Begin VB.Data data_fact 
      Caption         =   "data_fact"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Confirmar Pago"
      Height          =   375
      Left            =   7680
      Picture         =   "frm_afilbusca.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Confirmar la realización de facturación de la afiliación"
      Top             =   6240
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1935
      Left            =   4560
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "frm_afilbusca.frx":1BB2
      Top             =   4200
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ver Observación"
      Height          =   375
      Left            =   2640
      Picture         =   "frm_afilbusca.frx":1BB8
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6240
      Width           =   1815
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   2205
      Left            =   3000
      TabIndex        =   9
      Top             =   3960
      Visible         =   0   'False
      Width           =   7815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Historial"
      Height          =   375
      Left            =   5160
      Picture         =   "frm_afilbusca.frx":2142
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Ver las acciones realizadas a la afiliación"
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Data data_hist 
      Caption         =   "data_hist"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin Crystal.CrystalReport cr2pant 
      Left            =   6240
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowPrintBtn=   0   'False
      WindowShowExportBtn=   0   'False
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Imprimir afiliación"
      Height          =   375
      Left            =   120
      Picture         =   "frm_afilbusca.frx":26CC
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6240
      Width           =   1815
   End
   Begin MSFlexGridLib.MSFlexGrid ms2 
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   3625
      _Version        =   393216
      FocusRect       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid DBGrid1 
      Height          =   3615
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   6376
      _Version        =   393216
      FocusRect       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Data data_afilcons 
      Caption         =   "data_afilcons"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   360
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   6360
      Picture         =   "frm_afilbusca.frx":2C56
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Buscar"
      Top             =   120
      Width           =   495
   End
   Begin VB.TextBox t_busca 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frm_afilbusca.frx":31E0
      Left            =   1560
      List            =   "frm_afilbusca.frx":31ED
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
   Begin Crystal.CrystalReport cr1print 
      Left            =   4440
      Top             =   2520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      WindowControlBox=   0   'False
      WindowMaxButton =   0   'False
      WindowMinButton =   0   'False
      DiscardSavedData=   -1  'True
      WindowState     =   1
      ProgressDialog  =   0   'False
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   0   'False
      WindowShowCancelBtn=   0   'False
      WindowShowPrintBtn=   0   'False
      WindowShowExportBtn=   0   'False
      WindowShowZoomCtl=   0   'False
      WindowShowProgressCtls=   0   'False
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404040&
      Caption         =   "Busca rápida por cédula:"
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
      Left            =   5160
      TabIndex        =   18
      Top             =   6720
      Width           =   2295
   End
   Begin VB.Label labpermiso 
      Height          =   255
      Left            =   7560
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      Height          =   255
      Left            =   10200
      TabIndex        =   7
      Top             =   120
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Buscar por:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frm_afilbusca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub b_cancelar_Click()
Dim DeseaCancelar As String
DeseaCancelar = ""


If DBGrid1.TextMatrix(DBGrid1.RowSel, 2) = WElusuario Or WElusuario = "COMPUTOS" Then
    b_cancelar.Enabled = False
    DeseaCancelar = MsgBox("Desea anular el contrato Nro:" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1)) & " ?", vbInformation + vbYesNo)
    If DeseaCancelar = vbYes Then
       data_cancelar.RecordSource = "Select * from afiliaciones_new where afilia_nro =" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1)) & " and pendiente not in (20)"
       data_cancelar.Refresh
       If data_cancelar.Recordset.RecordCount > 0 Then
          data_abm.RecordSource = "select * from abmsocio where fecha =#" & Format(Date, "yyyy/mm/dd") & "#"
          data_abm.Refresh
          data_cancelar.Recordset.MoveFirst
          data_hist.RecordSource = "select * from afiliaciones_impre where nro_afilia =" & data_cancelar.Recordset("afilia_nro")
          data_hist.Refresh
          data_hist.Recordset.AddNew
          data_hist.Recordset("fecha") = Date
          data_hist.Recordset("hora") = Format(Time, "HH:mm")
          data_hist.Recordset("usuario") = WElusuario
          data_hist.Recordset("base") = frm_menu.data_parse.Recordset("base")
          data_hist.Recordset("nro_afilia") = data_cancelar.Recordset("afilia_nro")
          data_hist.Recordset("accion") = "AFILIACION CANCELADA"
          data_hist.Recordset.Update
          Do While Not data_cancelar.Recordset.EOF
             data_cancelar.Recordset.Edit
             data_cancelar.Recordset("pendiente") = 20
             data_cancelar.Recordset("convenio") = "CANCELADO"
             data_cancelar.Recordset.Update
             If IsNull(data_cancelar.Recordset("matricula")) = False Then
                data_hist.RecordSource = "select * from clientes where cl_codigo =" & data_cancelar.Recordset("matricula")
                data_hist.Refresh
                If data_hist.Recordset.RecordCount > 0 Then
                   data_hist.Recordset.Edit
                   data_hist.Recordset("cl_codconv") = "PART"
                   data_hist.Recordset("cl_nomconv") = "PARTICULARES"
                   data_hist.Recordset("estado") = 2
                   data_hist.Recordset("fecha_baja") = Date
                   data_hist.Recordset.Update
                   data_abm.Recordset.AddNew
                   data_abm.Recordset("usuario") = WElusuario
                   data_abm.Recordset("fecha") = Date
                   data_abm.Recordset("hora") = Format(Time, "HH:mm")
                   data_abm.Recordset("cl_codigo") = data_cancelar.Recordset("matricula")
                   data_abm.Recordset("desc") = "BAJA"
                   data_abm.Recordset("cl_motivo") = "CANCELA AFIL."
                   data_abm.Recordset("convenio") = data_cancelar.Recordset("categ")
                   data_abm.Recordset("base") = frm_menu.data_parse.Recordset("base")
                   data_abm.Recordset.Update
                   MsgBox "Avise a Padrón social de la cancelación. Socio pasó como PARTICULAR!", vbInformation
                End If
             End If
             data_cancelar.Recordset.MoveNext
          Loop
          data_cancelar.Recordset.MoveFirst
          Command6_Click
          MsgBox "Terminado"
          Carga_grid
          ms2.Clear
       Else
          MsgBox "No existe datos para cancelar.", vbInformation
       End If
    End If
    b_cancelar.Enabled = True
Else
   MsgBox "No es el usuario creador de la afiliación.", vbExclamation
End If

End Sub

Private Sub b_correo_Click()

Dim Xarchtex As String
Dim Xsiimpafil As String
Dim Correo As String

data_inf.DatabaseName = App.path & "\informes.mdb"

data_inf.RecordSource = "select * from infcli"
data_inf.Refresh
If data_inf.Recordset.RecordCount > 0 Then
   data_inf.Recordset.MoveFirst
   Do While Not data_inf.Recordset.EOF
      data_inf.Recordset.Delete
      data_inf.Recordset.MoveNext
   Loop
End If
data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1))
data_afilcons.Refresh
data_afilcons.Recordset.MoveFirst
Do While Not data_afilcons.Recordset.EOF
   If IsNull(data_afilcons.Recordset("dondef")) = False Then
      If Val(Label3.Caption) = 10 Or Val(Label3.Caption) = 2 Then
      Else
         Label3.Caption = data_afilcons.Recordset("dondef")
      End If
   End If
   data_afilcons.Recordset.MoveNext
Loop
If Dir("C:\planillas\contrato.pdf") <> "" Then
   Kill ("C:\planillas\contrato.pdf")
End If
b_correo.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
Command4.Enabled = False
frm_afilbusca.MousePointer = 11
If Val(Label3.Caption) = 0 Or Val(Label3.Caption) = 3 Then
'If Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1)) > 0 Then
'If labnro.Caption <> "" Then
   data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1)) & " and integra_nro in (1)"
   data_afilcons.Refresh
   If IsNull(data_afilcons.Recordset("correo")) = False Then
      If Trim(data_afilcons.Recordset("correo")) = "NO APLICA" Then
         frm_afilbusca.MousePointer = 0
         MsgBox "No tiene cuenta de correo"
      Else
          Correo = data_afilcons.Recordset("correo")
          If IsNull(data_afilcons.Recordset("correo_enviado")) = True Then
            data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1)) & " order by integra_nro"
            data_afilcons.Refresh
            If data_afilcons.Recordset.RecordCount > 0 Then
               If IsNull(data_afilcons.Recordset("sifact")) = False Then
                  data_afilcons.Recordset.MoveFirst
                  Do While Not data_afilcons.Recordset.EOF
                     Genera_contrato
                     data_afilcons.Recordset.Edit
                     data_afilcons.Recordset("correo_enviado") = 1
                     data_afilcons.Recordset.Update
                     data_afilcons.Recordset.MoveNext
                  Loop
                  data_afilcons.Recordset.MovePrevious
                  data_hist.RecordSource = "select * from afiliaciones_impre"
                  data_hist.Refresh
                  data_hist.Recordset.AddNew
                  data_hist.Recordset("fecha") = Date
                  data_hist.Recordset("hora") = Format(Time, "HH:mm")
                  data_hist.Recordset("usuario") = WElusuario
                  data_hist.Recordset("base") = frm_menu.data_parse.Recordset("base")
                  data_hist.Recordset("nro_afilia") = data_afilcons.Recordset("afilia_nro")
                  data_hist.Recordset("accion") = "ENVIO POR CORREO"
                  data_hist.Recordset.Update
                  data_afilcons.Recordset.MoveNext
                  data_inf.RecordSource = "select * from infcli order by cl_cantpag"
                  data_inf.Refresh
                  If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
                     If data_inf.Recordset("cl_nom_sup") = "Cobro por Tarjeta" Then
                        cr1print.ReportFileName = App.path & "\contrato_deb.rpt"
                     Else
                        cr1print.ReportFileName = App.path & "\contrato_afil.rpt"
                     End If
                  Else
                     cr1print.ReportFileName = App.path & "\contrato_afil.rpt"
                  End If
                  cr1print.Action = 1
                  Sleep 20000
'2693002-1
                  frm_afilbusca.MousePointer = 0
                  MsgBox "Se enviará el correo con el contrato adjunto", vbInformation
                  frm_afilbusca.MousePointer = 11
                  If Dir("C:\planillas\contrato_Nro_" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1)) & ".pdf") <> "" Then
                     Kill ("C:\planillas\contrato_Nro_" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1)) & ".pdf")
                  End If
                  
                  If Dir("C:\planillas\contrato.pdf") <> "" Then
                     Name "C:\planillas\contrato.pdf" As "C:\planillas\contrato_Nro_" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1)) & ".pdf"
                     Xarchtex = "C:\planillas\contrato_Nro_" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1)) & ".pdf"
'IMAP
'Nombre de servidor: outlook.office365.com
'Puerto: 993
'Método de cifrado: TLS

'vIERNESCONONDA123
'Nombre de servidor: smtp.office365.com
'Puerto: 587
'Método de cifrado: STARTTLS
                     Dim MenCorreo As String
                     Dim oMail As Class1
                     Set oMail = New Class1
                         With oMail
                             .UseAuntentificacion = True
                             .servidor = "smtp.office365.com"
                             .puerto = 25
                             .UseAuntentificacion = True
                             .ssl = True
                             .Usuario = "facturacion@sapp.com.uy"
                             .PassWord = "vIERNESCONONDA123"
    '                         .PassWord = "PpasJfsh8719"
                             .Asunto = "Contrato SAPP"
                             .de = "facturacion@sapp.com.uy"
                             .para = Correo
                    '         .para = "sappjorge@hotmail.com; despachosapp@hotmail.com; sappsusanadominguez@hotmail.com; sappdirecciontecnica@hotmail.com; sappenrique@hotmail.com"
                             .Adjunto = Xarchtex
                             .Mensaje = "Se adjunta contrato SAPP Nro. " & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1))
                             .Enviar_Backup ' manda el mail
                         End With
                         Set oMail = Nothing
                         frm_afilbusca.MousePointer = 0
                         MsgBox "Correo enviado.", vbInformation
                         
                  Else
                     frm_afilbusca.MousePointer = 0
                     MsgBox "No se pudo generar el archivo correctamente, reintente nuevamente.", vbCritical
                  End If
               
               Else
                   frm_afilbusca.MousePointer = 0
                   MsgBox "No se puede enviar porque falta la facturación.", vbCritical
               End If
            Else
               frm_afilbusca.MousePointer = 0
               MsgBox "No se encuentra afiliación."
            End If
          Else
            MsgBox "El correo ya fue enviado. Solo se puede enviar una vez.", vbCritical
          End If
      End If
   Else
        frm_afilbusca.MousePointer = 0
        MsgBox "La afiliación no tiene correo ingresado el titular.", vbCritical
       
   End If
Else
   frm_afilbusca.MousePointer = 0
   MsgBox "La Afiliación está pendiente de autorizar.", vbExclamation

End If
b_correo.Enabled = True
Command2.Enabled = True
Command3.Enabled = True
Command4.Enabled = True
frm_afilbusca.MousePointer = 0

End Sub

Private Sub b_infAfiliaciones_Click()
Dim desde, hasta, Promo As String
Dim Xobjexel22 As Excel.Application
Dim Xlibexel22 As Excel.Workbook
Dim Xarchexel22 As New Excel.Worksheet
Dim ImprimirDetalle As String
Dim Xlin, XCol As Integer
Dim Xtotreg, Xsub As Long
Dim Xarchtex As String
Dim Xlabrir3 As New Excel.Application
Dim Fecha1, Fecha2 As String
ImprimirDetalle = ""
If Month(Date) < 10 Then
   Fecha1 = "01/0" & Trim(str(Month(Date))) & "/" & Trim(str(Year(Date)))
   If Day(Date) < 10 Then
      Fecha2 = "0" & Trim(str(Day(Date))) & "/0" & Trim(str(Month(Date))) & "/" & Trim(str(Year(Date)))
   Else
      Fecha2 = Trim(str(Day(Date))) & "/0" & Trim(str(Month(Date))) & "/" & Trim(str(Year(Date)))
   End If
Else
   Fecha1 = "01/" & Trim(str(Month(Date))) & "/" & Trim(str(Year(Date)))
   If Day(Date) < 10 Then
      Fecha2 = "0" & Trim(str(Day(Date))) & "/" & Trim(str(Month(Date))) & "/" & Trim(str(Year(Date)))
   Else
      Fecha2 = Trim(str(Day(Date))) & "/" & Trim(str(Month(Date))) & "/" & Trim(str(Year(Date)))
   End If

End If

desde = InputBox("Ingrese fecha de inicio (formato: DD/MM/AAAA):", "FECHA INICIAL", Fecha1)
hasta = InputBox("Ingrese fecha final (formato: DD/MM/AAAA):", "FECHA FINAL", Fecha2)
Promo = InputBox("INGRESE CÓDIGO DE PROMOTOR (CERO PARA LISTAR TODOS)", "CODIGO de PROMOTOR", 0)
If Trim(Promo) = "" Then
   Promo = "0"
End If
frm_afilbusca.MousePointer = 11
Xlin = 1
XCol = 1
Xtotreg = 0
Xsub = 0

If desde <> "" And hasta <> "" Then
   ImprimirDetalle = MsgBox("Desea Imprimir todos los integrantes?", vbInformation + vbYesNo, "Afiliaciones")
   If ImprimirDetalle = vbYes Then
      If Val(Promo) = 0 Then
         data_afilcons.RecordSource = "select * from afiliaciones_new where fecha >=#" & Format(CDate(desde), "yyyy/mm/dd") & "# and fecha <=#" & Format(CDate(hasta), "yyyy/mm/dd") & "# order by codvende,fecha"
         data_afilcons.Refresh
      Else
         data_afilcons.RecordSource = "select * from afiliaciones_new where fecha >=#" & Format(CDate(desde), "yyyy/mm/dd") & "# and fecha <=#" & Format(CDate(hasta), "yyyy/mm/dd") & "# and codvende =" & Val(Promo) & " order by codvende,fecha"
         data_afilcons.Refresh
      End If
   Else
      If Val(Promo) = 0 Then
         data_afilcons.RecordSource = "select * from afiliaciones_new where fecha >=#" & Format(CDate(desde), "yyyy/mm/dd") & "# and fecha <=#" & Format(CDate(hasta), "yyyy/mm/dd") & "# and integra_nro in (1) order by codvende,fecha"
         data_afilcons.Refresh
      Else
         data_afilcons.RecordSource = "select * from afiliaciones_new where fecha >=#" & Format(CDate(desde), "yyyy/mm/dd") & "# and fecha <=#" & Format(CDate(hasta), "yyyy/mm/dd") & "# and codvende =" & Val(Promo) & " and integra_nro in (1) order by codvende,fecha"
         data_afilcons.Refresh
      End If
   End If
   If data_afilcons.Recordset.RecordCount > 0 Then
      data_afilcons.Recordset.MoveFirst
      Set Xobjexel22 = New Excel.Application
      Set Xlibexel22 = Xobjexel22.Workbooks.Add
      Set Xarchexel22 = Xlibexel22.Worksheets.Add
      Xarchexel22.Name = Trim("Afiliaciones")
      Xlibexel22.SaveAs ("C:\planillas\Afiliaciones.xls")
      Xarchtex = "C:\planillas\Afiliaciones.xls"
      Xarchexel22.Cells(Xlin, XCol) = "DEPARTAMENTO TI SAPP S.A."
      Xlin = Xlin + 1
      XCol = XCol + 1
      Xarchexel22.Range("A1", "C3").Font.Size = 16
      Xarchexel22.Range("A" & Trim(str(Xlin)), "J" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
      Xarchexel22.Cells(Xlin, XCol) = "INFORME DE AFILIACIONES POR PROMOTOR DESDE: " & desde & " HASTA: " & hasta
      XCol = 1
      Xlin = Xlin + 2
      Xnrocan = Xnrocan + Xlin
      Xarchexel22.Range("A" & Trim(str(Xlin)), "J" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
      Xarchexel22.Range("A" & Trim(str(Xlin))).ColumnWidth = 12
      Xarchexel22.Cells(Xlin, XCol) = "FECHA"
      XCol = XCol + 1
      Xarchexel22.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
      Xarchexel22.Cells(Xlin, XCol) = "CEDULA"
      XCol = XCol + 1
      Xarchexel22.Range("C" & Trim(str(Xlin))).ColumnWidth = 35
      Xarchexel22.Cells(Xlin, XCol) = "NOMBRE"
      XCol = XCol + 1
      Xarchexel22.Range("D" & Trim(str(Xlin))).ColumnWidth = 16
      Xarchexel22.Cells(Xlin, XCol) = "CATEGORIA"
      XCol = XCol + 1
      Xarchexel22.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "COD.PROM."
      XCol = XCol + 1
      Xarchexel22.Range("F" & Trim(str(Xlin))).ColumnWidth = 20
      Xarchexel22.Cells(Xlin, XCol) = "PROMOTOR"
      XCol = XCol + 1
      Xarchexel22.Range("G" & Trim(str(Xlin))).ColumnWidth = 16
      Xarchexel22.Cells(Xlin, XCol) = "ZONA"
      XCol = XCol + 1
      Xarchexel22.Range("H" & Trim(str(Xlin))).ColumnWidth = 10
      Xarchexel22.Cells(Xlin, XCol) = "BASE"
      XCol = XCol + 1
      Xarchexel22.Range("I" & Trim(str(Xlin))).ColumnWidth = 26
      Xarchexel22.Cells(Xlin, XCol) = "CELULAR/TELEF."
      XCol = XCol + 1
      Xarchexel22.Range("J" & Trim(str(Xlin))).ColumnWidth = 18
      Xarchexel22.Cells(Xlin, XCol) = "ESTADO"
      
      Xlin = Xlin + 1
      XCol = 1
        
      Do While Not data_afilcons.Recordset.EOF
         Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_afilcons.Recordset("fecha"), "dd/mm/yyyy")
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = data_afilcons.Recordset("cedula")
         XCol = XCol + 1
         If IsNull(data_afilcons.Recordset("nom2")) = False Then
            If IsNull(data_afilcons.Recordset("ape2")) = False Then
               Xarchexel22.Cells(Xlin, XCol) = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("ape2") & " " & data_afilcons.Recordset("nom1") & " " & data_afilcons.Recordset("nom2")
            Else
               Xarchexel22.Cells(Xlin, XCol) = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("nom1") & " " & data_afilcons.Recordset("nom2")
            End If
         Else
            If IsNull(data_afilcons.Recordset("ape2")) = False Then
               Xarchexel22.Cells(Xlin, XCol) = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("ape2") & " " & data_afilcons.Recordset("nom1")
            Else
               Xarchexel22.Cells(Xlin, XCol) = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("nom1")
            End If
         End If
         XCol = XCol + 1
         If IsNull(data_afilcons.Recordset("convenio")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = data_afilcons.Recordset("convenio")
         Else
            Xarchexel22.Cells(Xlin, XCol) = "S/D"
         End If
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = data_afilcons.Recordset("codvende")
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = Devuelve_vende()
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = data_afilcons.Recordset("nomzona")
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = data_afilcons.Recordset("wbase")
         XCol = XCol + 1
         Xarchexel22.Cells(Xlin, XCol) = "Cel:" & data_afilcons.Recordset("celular") & "/Tel:" & data_afilcons.Recordset("telef")
         XCol = XCol + 1
         If data_afilcons.Recordset("pendiente") = 0 Then
            Xarchexel22.Cells(Xlin, XCol) = "Pendiente Padrón social"
         Else
            If data_afilcons.Recordset("pendiente") = 10 Then
               Xarchexel22.Cells(Xlin, XCol) = "Procesada"
            Else
               If data_afilcons.Recordset("pendiente") = 2 Then
                  Xarchexel22.Cells(Xlin, XCol) = "Pendiente AUTORIZAR"
               Else
                  If data_afilcons.Recordset("pendiente") = 11 Then
                     Xarchexel22.Cells(Xlin, XCol) = "CANCELADA"
                  Else
                     If data_afilcons.Recordset("pendiente") = 4 Then
                        Xarchexel22.Cells(Xlin, XCol) = "Falta facturar"
                     Else
                        If data_afilcons.Recordset("pendiente") = 20 Then
                           Xarchexel22.Cells(Xlin, XCol) = "CANCELADA"
                        Else
                           Xarchexel22.Cells(Xlin, XCol) = "VERIFICAR"
                        End If
                     End If
                  End If
               End If
            End If
         End If
         Xlin = Xlin + 1
         XCol = 1
         Xtotreg = Xtotreg + 1
         data_afilcons.Recordset.MoveNext
      Loop
      Xlin = Xlin + 1
      XCol = 1
      Xarchexel22.Cells(Xlin, XCol) = "Total Registros: " & Trim(str(Xtotreg))
      Xlin = Xlin + 1
      XCol = 1
      Xarchexel22.Cells(Xlin, XCol) = "FECHA DE EMISION:" & Format(Date, "dd/mm/yyyy")
      Xlibexel22.Save
      Xlibexel22.Close
      Xobjexel22.Quit
      Xlabrir3.Workbooks.Open Xarchtex, , False
      Xlabrir3.Visible = True
      Xlabrir3.WindowState = xlMaximized
      frm_afilbusca.MousePointer = 0
      MsgBox "Terminado"
   Else
      frm_afilbusca.MousePointer = 0
      MsgBox "No hay registros"
   End If
Else
   frm_afilbusca.MousePointer = 0
   MsgBox "Faltan fechas"
End If

End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
   Carga_grid_cancelado
Else
   Carga_grid
End If

End Sub

Private Sub Command1_Click()
DBGrid1.Clear
ms2.Clear
Carga_grid

End Sub

Private Sub Command2_Click()
Dim ImprimeContra As String
ImprimeContra = ""
data_inf.DatabaseName = App.path & "\informes.mdb"

data_inf.RecordSource = "select * from infcli"
data_inf.Refresh
If data_inf.Recordset.RecordCount > 0 Then
   data_inf.Recordset.MoveFirst
   Do While Not data_inf.Recordset.EOF
      data_inf.Recordset.Delete
      data_inf.Recordset.MoveNext
   Loop
End If
data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1))
data_afilcons.Refresh
If data_afilcons.Recordset.RecordCount > 0 Then
    data_afilcons.Recordset.MoveFirst
    Do While Not data_afilcons.Recordset.EOF
       If IsNull(data_afilcons.Recordset("dondef")) = False Then
          If Val(Label3.Caption) = 10 Or Val(Label3.Caption) = 2 Then
          Else
             Label3.Caption = data_afilcons.Recordset("dondef")
          End If
       End If
       data_afilcons.Recordset.MoveNext
    Loop
    If Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1)) > 0 Then
       ImprimeContra = MsgBox("Desea imprimir contrato?", vbInformation + vbYesNo, "Afiliaciones SAPP")
       If Val(Label3.Caption) = 10 Or Val(Label3.Caption) = 2 Then
          ImprimeContra = vbNo
       End If
        If ImprimeContra = vbYes Then
           If Val(Label3.Caption) <> 0 Then
              MsgBox "La afiliación está sin facturar, luego de firmada deberá ser facturada en BASE " & Label3.Caption
           End If
           data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1))
           data_afilcons.Refresh
           data_afilcons.Recordset.MoveFirst
           Do While Not data_afilcons.Recordset.EOF
              
              Genera_contrato
              data_afilcons.Recordset.Edit
              If IsNull(data_afilcons.Recordset("cant_impre")) = False Then
                 data_afilcons.Recordset("cant_impre") = data_afilcons.Recordset("cant_impre") + 1
              Else
                 data_afilcons.Recordset("cant_impre") = 1
              End If
              data_afilcons.Recordset.Update
              data_afilcons.Recordset.MoveNext
           Loop
           data_afilcons.Recordset.MovePrevious
           data_hist.RecordSource = "select * from afiliaciones_impre"
           data_hist.Refresh
           data_hist.Recordset.AddNew
           data_hist.Recordset("fecha") = Date
           data_hist.Recordset("hora") = Format(Time, "HH:mm")
           data_hist.Recordset("usuario") = WElusuario
           data_hist.Recordset("base") = frm_menu.data_parse.Recordset("base")
           data_hist.Recordset("nro_afilia") = data_afilcons.Recordset("afilia_nro")
           data_hist.Recordset("accion") = "IMPRESION"
           data_hist.Recordset.Update
           data_afilcons.Recordset.MoveNext
           data_inf.RecordSource = "select * from infcli order by cl_cantpag"
           data_inf.Refresh
           If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
              If data_inf.Recordset("cl_nom_sup") = "Cobro por Tarjeta" Then
                 cr1print.ReportFileName = App.path & "\contrato_debprint.rpt"
              Else
                 cr1print.ReportFileName = App.path & "\contrato_afilprint.rpt"
              End If
           Else
              cr1print.ReportFileName = App.path & "\contrato_afilprint.rpt"
           End If
           cr1print.Action = 1
        Else
           data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1))
           data_afilcons.Refresh
           data_afilcons.Recordset.MoveFirst
           Do While Not data_afilcons.Recordset.EOF
              
              Genera_contrato
              data_afilcons.Recordset.MoveNext
           Loop
           data_inf.RecordSource = "select * from infcli order by cl_cantpag"
           data_inf.Refresh
           If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
              If data_inf.Recordset("cl_nom_sup") = "Cobro por Tarjeta" Then
                 cr2pant.ReportFileName = App.path & "\contrato_debprint.rpt"
              Else
                 cr2pant.ReportFileName = App.path & "\contrato_afilprint.rpt"
              End If
           Else
              cr2pant.ReportFileName = App.path & "\contrato_afilprint.rpt"
           End If
           cr2pant.Action = 1
        End If
    Else
       MsgBox "Falta seleccionar"
       
    End If
End If

End Sub

Private Sub Command3_Click()

If List1.Visible = False Then
    List1.Visible = True
    data_hist.RecordSource = "select * from afiliaciones_impre where nro_afilia =" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1)) & " order by fecha"
    data_hist.Refresh
    List1.Clear
    If data_hist.Recordset.RecordCount > 0 Then
       data_hist.Recordset.MoveFirst
       Do While Not data_hist.Recordset.EOF
          List1.AddItem Format(data_hist.Recordset("fecha"), "dd/mm/yyyy") & " " & data_hist.Recordset("hora") & "-->" & data_hist.Recordset("usuario") & "-->" & data_hist.Recordset("accion") & "-->BASE: " & data_hist.Recordset("base")
          data_hist.Recordset.MoveNext
       Loop
       List1.AddItem "-----Presione nuevamente el botón historial para ocultar---"
       
    Else
       List1.AddItem "-----Presione nuevamente el botón historial para ocultar---"
    
    End If
    Command4.Enabled = False
    Command2.Enabled = False
Else
    List1.Clear
    List1.Visible = False
    Command4.Enabled = True
    Command2.Enabled = True

End If

End Sub

Private Sub Command4_Click()
If Text1.Visible = False Then

    data_hist.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1)) & " and integra_nro =" & Val(ms2.TextMatrix(ms2.RowSel, 0))
    data_hist.Refresh
    If data_hist.Recordset.RecordCount > 0 Then
       If IsNull(data_hist.Recordset("obs_adm")) = False Then
          Text1.Visible = True
          Text1.Text = ""
          Text1.Text = data_hist.Recordset("fec_obsadm") & " -->"
          Text1.Text = Text1.Text & data_hist.Recordset("obs_adm")
          Command4.Enabled = True
          Command3.Enabled = False
          Command2.Enabled = False
          ms2.Enabled = False
       Else
          Text1.Text = ""
       End If
    Else
       Text1.Text = ""
    End If
Else
    Text1.Visible = False
    Text1.Text = ""
    Command4.Enabled = True
    Command3.Enabled = True
    Command2.Enabled = True
    ms2.Enabled = True
End If


End Sub

Private Sub Command5_Click()
Dim IngreseFactura As String
IngreseFactura = InputBox("Ingrese el número de la factura realizada:", "Afiliaciones SAPP")
If Trim(IngreseFactura) <> "" Then
   data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1)) & " and integra_nro =" & 1
   data_afilcons.Refresh
   If data_afilcons.Recordset.RecordCount > 0 Then
      data_afilcons.Recordset.MoveFirst
      If IsNull(data_afilcons.Recordset("matricula")) = False Then
         data_fact.RecordSource = "select * from linmmdd where factura =" & Val(xingresefactura) & " and cod_cli =" & data_afilcons.Recordset("matricula") & " and cod_prod in (992)"
         data_fact.Refresh
         If data_fact.Recordset.RecordCount > 0 Then
            If IsNull(data_afilcons.Recordset("sifact")) = False Then
               If data_afilcons.Recordset("sifact") <> 1 Then
                  data_afilcons.Recordset.Edit
                  data_afilcons.Recordset("sifact") = 1
                  data_afilcons.Recordset.Update
               Else
                  MsgBox "Ya figura registrado el pago", vbInformation
               End If
            Else
               data_afilcons.Recordset.Edit
               data_afilcons.Recordset("sifact") = 1
               data_afilcons.Recordset.Update
            End If
            MsgBox "Pago confirmado.", vbInformation
         Else
            MsgBox "No se encuentra el número de factura ingresada, verifique!", vbCritical
         End If
      Else
         MsgBox "La afiliación no figura con número de matrícula, VERIFIQUE!", vbCritical
      End If
   Else
      MsgBox "No se encuentra la afiliación, reintente!", vbCritical
   End If
Else
   MsgBox "No ingresó número de factura", vbCritical
End If

End Sub

Private Sub Command6_Click()
Dim MenCorreo2 As String
Dim oMail2 As Class1
Set oMail2 = New Class1
MenCorreo2 = "AFILIACION CANCELADA Nro. " & data_cancelar.Recordset("afilia_nro") & vbCrLf
MenCorreo2 = MenCorreo2 & "FECHA ACTUAL: " & Format(Date, "dd/mm/yyyy") & " HORA:" & Format(Time, "HH:mm") & vbCrLf
MenCorreo2 = MenCorreo2 & "USUARIO: " & WElusuario & " BASE:" & frm_menu.data_parse.Recordset("base") & vbCrLf
If IsNull(data_cancelar.Recordset("matricula")) = False Then
   MenCorreo2 = MenCorreo2 & "MATRICULA: " & data_cancelar.Recordset("matricula")
Else
   MenCorreo2 = MenCorreo2 & "MATRICULA: " & "SIN ASIGNAR"
End If

With oMail2
       .servidor = "smtp.gmail.com"
       .puerto = 465
       .UseAuntentificacion = True
       .ssl = True
       .Usuario = "sappsistemas@gmail.com"
       .PassWord = "sapp1987"
       .Asunto = "AFILIACION CANCELADA Número: " & data_cancelar.Recordset("afilia_nro")
       .de = "sappsistemas@gmail.com"
       .para = "jefefacturacion@sapp.com.uy"
'             .Adjunto = Xarchtex
       .Mensaje = MenCorreo2
       .Enviar_Backup ' manda el mail
End With
Set oMail2 = Nothing


End Sub

Private Sub DBGrid1_DblClick()
'Xnro = Val(DBGrid1.TextMatrix(flex1.RowSel, 1))
'Xnroh = Val(flex1.TextMatrix(flex1.RowSel, 3))
Dim Xsqlpromos As String
Dim Xreccliis As New ADODB.Recordset
Dim Xcann As Integer

ConectarBD
ConbdSapp.Open
             
''''Grabar matrícula del socio para poder buscar acá

Xsqlpromos = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.nom1,afiliaciones_new.nom2,afiliaciones_new.ape1,afiliaciones_new.ape2,afiliaciones_new.pendiente," & _
"afiliaciones_new.sifact,afiliaciones_new.obs_noaut,afiliaciones_new.obs_adm,afiliaciones_new.matricula,afiliaciones_new.fnac,afiliaciones_new.integra_nro,afiliaciones_new.convenio,afiliaciones_new.celular,afiliaciones_new.cedula,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
"from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.afilia_nro =" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1)) & " order by afiliaciones_new.integra_nro"
             
With Xreccliis
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromos, ConbdSapp, , , adCmdText
End With
ms2.Clear
ms2.rows = 2
ms2.Cols = 9
ms2.TextMatrix(0, 0) = "NRO."
ms2.ColWidth(0) = 1300
ms2.TextMatrix(0, 1) = "CEDULA"
ms2.ColWidth(1) = 1500
ms2.TextMatrix(0, 2) = "NOMBRES"
ms2.ColWidth(2) = 2500
ms2.TextMatrix(0, 3) = "APELLIDOS"
ms2.ColWidth(3) = 2500
ms2.TextMatrix(0, 4) = "F.NAC."
ms2.ColWidth(4) = 1500
ms2.TextMatrix(0, 5) = "CELULAR"
ms2.ColWidth(5) = 1500
ms2.TextMatrix(0, 6) = "MATRICULA"
ms2.ColWidth(6) = 1500
ms2.TextMatrix(0, 7) = "OBS.NO AUTORIZADO"
ms2.ColWidth(7) = 3500
ms2.TextMatrix(0, 8) = "OBSERVACION ADM."
ms2.ColWidth(8) = 3500

Xcann = 1
Dim Xpendiente As Integer
Dim Xsifactura As Integer

Xpendiente = 0
Xsifactura = 0
If Xreccliis.RecordCount > 0 Then
   Xreccliis.MoveFirst
   Xpendiente = Xreccliis("pendiente")
   If IsNull(Xreccliis("sifact")) = False Then
      Xsifactura = Xreccliis("sifact")
   End If
   If Xpendiente = 2 Or Xpendiente = 10 Then
      If Xpendiente = 2 Then
         MsgBox "La afiliación está pendiente de autorización, no se puede editar los socios.", vbCritical
'         ms2.Enabled = False
         Command2.Enabled = False
         Label3.Caption = 2
         b_correo.Enabled = False
      Else
         MsgBox "La afiliación ya fue procesada al padrón, solo puede visualizar la afiliación.", vbInformation
         ms2.Enabled = True
         Command2.Enabled = True
         Label3.Caption = 10
         b_correo.Enabled = True
      End If
   Else
      If Xpendiente = 11 Or Xpendiente = 20 Then
         If Xpendiente = 20 Then
            MsgBox "La afiliación está cancelada. Puede verificar el historial.", vbCritical
            ms2.Enabled = False
            Command2.Enabled = False
            Label3.Caption = 11
            b_correo.Enabled = False
         Else
            MsgBox "La afiliación está cancelada. Puede verificar el historial.", vbCritical
            ms2.Enabled = False
            Command2.Enabled = False
            Label3.Caption = 11
            b_correo.Enabled = False
         End If
      Else
         If Xpendiente = 4 Then
            MsgBox "La afiliación está sin facturar. Sólo puede visualizar contrato por pantalla.", vbCritical
            ms2.Enabled = False
            Command2.Enabled = False
            Label3.Caption = 4
            b_correo.Enabled = False
         Else
            ms2.Enabled = True
            Command2.Enabled = True
            Label3.Caption = Xreccliis("pendiente")
            b_correo.Enabled = True
         End If
      End If
   End If
   If Xsifactura = 0 Then
      MsgBox "La afiliación está pendiente de FACTURAR.", vbCritical
'      Command2.Enabled = False
'      Label3.Caption = 2
      b_correo.Enabled = False
   End If
   Do While Not Xreccliis.EOF
      ms2.TextMatrix(Xcann, 0) = Xreccliis("integra_nro")
      ms2.TextMatrix(Xcann, 1) = Xreccliis("cedula")
      If IsNull(Xreccliis("nom2")) = False Then
         ms2.TextMatrix(Xcann, 2) = Xreccliis("nom1") & " " & Xreccliis("nom2")
      Else
         ms2.TextMatrix(Xcann, 2) = Xreccliis("nom1")
      End If
      If IsNull(Xreccliis("ape2")) = False Then
         ms2.TextMatrix(Xcann, 3) = Xreccliis("ape1") & " " & Xreccliis("ape2")
      Else
         ms2.TextMatrix(Xcann, 3) = Xreccliis("ape1")
      End If
      ms2.TextMatrix(Xcann, 4) = Xreccliis("fnac")
      ms2.TextMatrix(Xcann, 5) = Xreccliis("celular")
      If IsNull(Xreccliis("matricula")) = False Then
         ms2.TextMatrix(Xcann, 6) = Xreccliis("matricula")
      Else
         ms2.TextMatrix(Xcann, 6) = "0"
      End If
      If IsNull(Xreccliis("obs_noaut")) = False Then
         ms2.TextMatrix(Xcann, 7) = Xreccliis("obs_noaut")
      End If
      If IsNull(Xreccliis("obs_adm")) = False Then
         ms2.TextMatrix(Xcann, 7) = Xreccliis("obs_adm")
      End If
      
      ms2.rows = ms2.rows + 1
      Xreccliis.MoveNext
      Xcann = Xcann + 1
   Loop
End If

Xreccliis.Close
ConbdSapp.Close

End Sub

Private Sub Form_Load()
data_afilcons.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_hist.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_fact.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_abm.Connect = "odbc;dsn=" & Xconexrmt & ";"

'data_buscasi.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cancelar.Connect = "odbc;dsn=" & Xconexrmt & ";"

If ControlUsuario(b_infAfiliaciones.Name) = 1 Then
   b_infAfiliaciones.Enabled = True
   labpermiso.Caption = "1"
Else
   b_infAfiliaciones.Enabled = False
   labpermiso.Caption = "0"
End If

Carga_grid

End Sub

Public Sub Carga_grid()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Xcann, Xpendiente As Integer
Dim Ladate As Date
Ladate = Date - 30
Xpendiente = 0
ConectarBD
ConbdSapp.Open
             
'Data1.ConnectionString = "dsn=sappnew"
If labpermiso.Caption = "1" Then
   If t_busca.Text = "" Then
      Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.nom1,afiliaciones_new.ape1,afiliaciones_new.pendiente,afiliaciones_new.integra_nro,afiliaciones_new.wusuario," & _
      "afiliaciones_new.convenio,afiliaciones_new.celular,afiliaciones_new.telef,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
      "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.fecha >='" & Format(Ladate, "yyyy/mm/dd") & "' and afiliaciones_new.pendiente not in (20) and afiliaciones_new.integra_nro in (1) order by afiliaciones_new.fecha DESC"
   Else
      If Combo1.Text = "CEDULA" Then
         Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.nom1,afiliaciones_new.ape1,afiliaciones_new.pendiente,afiliaciones_new.integra_nro,afiliaciones_new.wusuario," & _
         "afiliaciones_new.cedula,afiliaciones_new.convenio,afiliaciones_new.celular,afiliaciones_new.telef,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
         "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.pendiente not in (20) and afiliaciones_new.integra_nro in (1) and afiliaciones_new.cedula ='" & t_busca.Text & "'  order by afiliaciones_new.fecha DESC"
      Else
         If Combo1.Text = "NRO.AFILIACION" Then
            Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.nom1,afiliaciones_new.ape1,afiliaciones_new.pendiente,afiliaciones_new.integra_nro,afiliaciones_new.wusuario," & _
            "afiliaciones_new.cedula,afiliaciones_new.convenio,afiliaciones_new.celular,afiliaciones_new.telef,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
            "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.integra_nro in (1) and afiliaciones_new.afilia_nro =" & t_busca.Text & " and afiliaciones_new.pendiente not in (20) order by afiliaciones_new.fecha DESC"
         Else
            If Combo1.Text = "FECHA" Then
               Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.nom1,afiliaciones_new.ape1,afiliaciones_new.pendiente,afiliaciones_new.integra_nro,afiliaciones_new.wusuario," & _
               "afiliaciones_new.convenio,afiliaciones_new.celular,afiliaciones_new.telef,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
               "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.fecha >='" & Format(t_busca.Text, "yyyy/mm/dd") & "' and afiliaciones_new.integra_nro in (1) and afiliaciones_new.pendiente not in (20) order by afiliaciones_new.fecha DESC"
            Else
               Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.nom1,afiliaciones_new.ape1,afiliaciones_new.pendiente,afiliaciones_new.integra_nro,afiliaciones_new.wusuario," & _
               "afiliaciones_new.convenio,afiliaciones_new.celular,afiliaciones_new.telef,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
               "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.fecha >='" & Format(Ladate, "yyyy/mm/dd") & "' and afiliaciones_new.integra_nro in (1) and afiliaciones_new.pendiente not in (20) order by afiliaciones_new.fecha DESC"
            End If
         End If
      End If
   End If
Else
   If t_busca.Text = "" Then
      Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.nom1,afiliaciones_new.ape1,afiliaciones_new.pendiente,afiliaciones_new.integra_nro,afiliaciones_new.wusuario," & _
      "afiliaciones_new.convenio,afiliaciones_new.celular,afiliaciones_new.telef,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
      "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.fecha >='" & Format(Ladate, "yyyy/mm/dd") & "' and afiliaciones_new.pendiente not in (20) and afiliaciones_new.integra_nro in (1) order by afiliaciones_new.fecha DESC"
   Else
      If Combo1.Text = "CEDULA" Then
         Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.nom1,afiliaciones_new.ape1,afiliaciones_new.pendiente,afiliaciones_new.integra_nro,afiliaciones_new.wusuario," & _
         "afiliaciones_new.cedula,afiliaciones_new.convenio,afiliaciones_new.celular,afiliaciones_new.telef,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
         "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.pendiente not in (20) and afiliaciones_new.integra_nro in (1) and afiliaciones_new.cedula ='" & t_busca.Text & "'  order by afiliaciones_new.fecha DESC"
      Else
         If Combo1.Text = "NRO.AFILIACION" Then
            Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.nom1,afiliaciones_new.ape1,afiliaciones_new.pendiente,afiliaciones_new.integra_nro,afiliaciones_new.wusuario," & _
            "afiliaciones_new.cedula,afiliaciones_new.convenio,afiliaciones_new.celular,afiliaciones_new.telef,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
            "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.integra_nro in (1) and afiliaciones_new.afilia_nro =" & t_busca.Text & " and afiliaciones_new.pendiente not in (20) order by afiliaciones_new.fecha DESC"
         Else
            If Combo1.Text = "FECHA" Then
               Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.nom1,afiliaciones_new.ape1,afiliaciones_new.pendiente,afiliaciones_new.integra_nro,afiliaciones_new.wusuario," & _
               "afiliaciones_new.convenio,afiliaciones_new.celular,afiliaciones_new.telef,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
               "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.fecha >='" & Format(t_busca.Text, "yyyy/mm/dd") & "' and afiliaciones_new.integra_nro in (1) and afiliaciones_new.pendiente not in (20) order by afiliaciones_new.fecha DESC"
            Else
               Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.nom1,afiliaciones_new.ape1,afiliaciones_new.pendiente,afiliaciones_new.integra_nro,afiliaciones_new.wusuario," & _
               "afiliaciones_new.convenio,afiliaciones_new.celular,afiliaciones_new.telef,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
               "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.fecha >='" & Format(Ladate, "yyyy/mm/dd") & "' and afiliaciones_new.integra_nro in (1) and afiliaciones_new.pendiente not in (20) order by afiliaciones_new.fecha DESC"
            End If
         End If
      End If
   End If

End If


With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
DBGrid1.Clear
ms2.Clear
DBGrid1.rows = 2
DBGrid1.Cols = 9
DBGrid1.TextMatrix(0, 0) = "FECHA"
DBGrid1.ColWidth(0) = 1300
DBGrid1.TextMatrix(0, 1) = "Nro.Af."
DBGrid1.ColWidth(1) = 1300
DBGrid1.TextMatrix(0, 2) = "USUARIO"
DBGrid1.ColWidth(2) = 1500

DBGrid1.TextMatrix(0, 3) = "NOMBRE"
DBGrid1.ColWidth(3) = 2900
DBGrid1.TextMatrix(0, 4) = "CATEGORIA"
DBGrid1.ColWidth(4) = 1500
DBGrid1.TextMatrix(0, 5) = "CELULAR"
DBGrid1.ColWidth(5) = 1500
DBGrid1.TextMatrix(0, 6) = "TELEFONO"
DBGrid1.ColWidth(6) = 1500
DBGrid1.TextMatrix(0, 7) = "PROMOTOR"
DBGrid1.ColWidth(7) = 1500
DBGrid1.TextMatrix(0, 8) = "ESTADO ACTUAL"
DBGrid1.ColWidth(8) = 1700

Xcann = 1

If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   Do While Not Xrecclii.EOF
      DBGrid1.TextMatrix(Xcann, 0) = Xrecclii("fecha")
      DBGrid1.TextMatrix(Xcann, 1) = Xrecclii("afilia_nro")
      DBGrid1.TextMatrix(Xcann, 2) = Xrecclii("wusuario")
      
      DBGrid1.TextMatrix(Xcann, 3) = Xrecclii("nom1") & " " & Xrecclii("ape1")
      DBGrid1.TextMatrix(Xcann, 4) = Xrecclii("convenio")
      If IsNull(Xrecclii("celular")) = False Then
         DBGrid1.TextMatrix(Xcann, 5) = Xrecclii("celular")
      End If
      If IsNull(Xrecclii("telef")) = False Then
         DBGrid1.TextMatrix(Xcann, 6) = Xrecclii("telef")
      End If
      If IsNull(Xrecclii("nombre")) = False Then
         DBGrid1.TextMatrix(Xcann, 7) = Xrecclii("nombre")
      End If
      Xpendiente = Xrecclii("pendiente")
      If Xpendiente = 0 Then
         DBGrid1.TextMatrix(Xcann, 8) = "PENDIENTE"
      Else
         If Xpendiente = 2 Then
            DBGrid1.TextMatrix(Xcann, 8) = "FALTA AUTORIZAR"
         Else
            If Xpendiente = 3 Then
               DBGrid1.TextMatrix(Xcann, 8) = "FALTA VERIFICAR"
            Else
               If Xpendiente = 4 Then
                  DBGrid1.TextMatrix(Xcann, 8) = "FALTA FACTURAR"
                  If Xpendiente = 10 Then
                      DBGrid1.TextMatrix(Xcann, 8) = "PROCESADA"
                  Else
                      DBGrid1.TextMatrix(Xcann, 8) = "SIN DATOS"
                  End If
               End If
            End If
         End If
      End If
      
      DBGrid1.rows = DBGrid1.rows + 1
      Xrecclii.MoveNext
      Xcann = Xcann + 1
   Loop
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Sub Genera_contrato()
Dim Direc, Xcontrato As String
Direc = ""


If data_afilcons.Recordset("convenio") = "COMPLEMENTO" Or data_afilcons.Recordset("convenio") = "COMPLEMENTO C.GALICIA" Or data_afilcons.Recordset("convenio") = "AMBULATORIO" Then
'   data_inf.DatabaseName = App.path & "\contrato.mdb"
'   data_inf.RecordSource = "contrato"
'   data_inf.Refresh
'   If IsNull(data_inf.Recordset("contrato")) = False Then
'      Xcontrato = data_inf.Recordset("contrato")
'   Else
'      Xcontrato = "Sin datos"
'   End If
   data_inf.DatabaseName = ""
   data_inf.Connect = "odbc;dsn=sappnew;"
   data_inf.RecordSource = "afiliaciones_contratoa"
   data_inf.Refresh
   If IsNull(data_inf.Recordset("descrip")) = False Then
      Xcontrato = data_inf.Recordset("descrip")
   Else
      Xcontrato = "Sin datos"
   End If
   data_inf.Connect = ""

Else
   data_inf.DatabaseName = ""
   data_inf.Connect = "odbc;dsn=sappnew;"
   data_inf.RecordSource = "afiliaciones_contratoe"
   data_inf.Refresh
   If IsNull(data_inf.Recordset("descrip")) = False Then
      Xcontrato = data_inf.Recordset("descrip")
   Else
      Xcontrato = "Sin datos"
   End If
   data_inf.Connect = ""
End If

data_inf.DatabaseName = App.path & "\informes.mdb"

data_inf.RecordSource = "select * from infcli"
data_inf.Refresh

If data_afilcons.Recordset.RecordCount > 0 Then
   data_inf.Recordset.AddNew
   data_inf.Recordset("cl_codigo") = data_afilcons.Recordset("afilia_nro")
   If IsNull(data_afilcons.Recordset("catcontrato")) = False Then
      data_inf.Recordset("cl_descpag") = data_afilcons.Recordset("catcontrato")
   Else
      If data_afilcons.Recordset("convenio") = "COMPLEMENTO" Or data_afilcons.Recordset("convenio") = "COMPLEMENTO C.GALICIA" Then
         data_inf.Recordset("cl_descpag") = "AMBULATORIO"
      Else
         data_inf.Recordset("cl_descpag") = data_afilcons.Recordset("convenio")
      End If
   End If
   data_inf.Recordset("cl_fecing") = data_afilcons.Recordset("fecha")
   data_inf.Recordset("cl_cantpag") = data_afilcons.Recordset("integra_nro")
   data_inf.Recordset("cl_apellid") = data_afilcons.Recordset("ape1")
   data_inf.Recordset("cl_medflia") = Mid(Devuelve_titular(), 1, 30)
   data_inf.Recordset("tit_tarj") = Mid(Devuelve_titularApe(), 1, 30)
   If IsNull(data_afilcons.Recordset("ape2")) = False Then
      data_inf.Recordset("cl_localid") = Mid(data_afilcons.Recordset("ape2"), 1, 35)
   End If
   data_inf.Recordset("cl_nomvend") = Mid(data_afilcons.Recordset("nom1"), 1, 35)
   If IsNull(data_afilcons.Recordset("nom2")) = False Then
      data_inf.Recordset("cl_nombre") = Mid(data_afilcons.Recordset("nom2"), 1, 30)
   End If
   data_inf.Recordset("cl_fnac") = data_afilcons.Recordset("fnac")
   If Len(data_afilcons.Recordset("cedula")) = 7 Then
      data_inf.Recordset("cl_fax") = Mid(data_afilcons.Recordset("cedula"), 1, 6) & "-" & Mid(data_afilcons.Recordset("cedula"), 7, 1)
   Else
      data_inf.Recordset("cl_fax") = Mid(data_afilcons.Recordset("cedula"), 1, 7) & "-" & Mid(data_afilcons.Recordset("cedula"), 8, 1)
   End If
   data_inf.Recordset("cl_tjemi_n") = Devuelve_titularCed()
   If IsNull(data_afilcons.Recordset("telef")) = False Then
      data_inf.Recordset("cl_telefon") = data_afilcons.Recordset("telef")
   End If
   data_inf.Recordset("cl_celular") = data_afilcons.Recordset("celular")
   If IsNull(data_afilcons.Recordset("correo")) = False Then
      data_inf.Recordset("cl_dircobr") = data_afilcons.Recordset("correo")
   End If
   If IsNull(data_afilcons.Recordset("codmut")) = False Then
      data_inf.Recordset("cl_socmnom") = Devuelve_mut()
   End If
   If IsNull(data_afilcons.Recordset("direc2")) = False Then
      Direc = data_afilcons.Recordset("direc1") & " E/" & data_afilcons.Recordset("direc2")
   Else
      Direc = data_afilcons.Recordset("direc1")
   End If
   data_inf.Recordset("cl_direcci") = Mid(Direc, 1, 80)
   If IsNull(data_afilcons.Recordset("manz")) = False Then
      data_inf.Recordset("cl_estadoc") = data_afilcons.Recordset("manz")
   End If
   If IsNull(data_afilcons.Recordset("solar")) = False Then
      data_inf.Recordset("cl_tipcli") = Mid(data_afilcons.Recordset("solar"), 1, 3)
   End If
   If IsNull(data_afilcons.Recordset("nomzona")) = False Then
      data_inf.Recordset("cl_zona") = Mid(data_afilcons.Recordset("nomzona"), 1, 25)
   End If
   data_inf.Recordset("cl_atrasoa") = data_afilcons.Recordset("valorcuota")
   If IsNull(data_afilcons.Recordset("desc_imp")) = False Then
      data_inf.Recordset("cl_seg_vto") = data_afilcons.Recordset("desc_imp")
   Else
      data_inf.Recordset("cl_seg_vto") = 0
   End If
   If IsNull(data_afilcons.Recordset("importe_fin")) = False Then
      data_inf.Recordset("cl_ter_vto") = data_afilcons.Recordset("importe_fin")
   Else
      data_inf.Recordset("cl_ter_vto") = 0
   End If
   
   If IsNull(data_afilcons.Recordset("tarj_nro")) = False Then
      data_inf.Recordset("cl_nom_sup") = "Cobro por Tarjeta"
      data_inf.Recordset("info_debit") = "COBRO POR DÉBITO AUTOMÁTICO:" & chr(13)
      data_inf.Recordset("info_debit") = data_inf.Recordset("info_debit") & "Se adjunta autorización débito automático al final del contrato."
      If IsNull(data_afilcons.Recordset("codvende")) = False Then
         data_inf.Recordset("cl_entre") = Devuelve_vende()
      Else
         data_inf.Recordset("cl_entre") = "Sin promotor"
      End If
      If IsNull(data_afilcons.Recordset("fec_desde")) = False Then
         data_inf.Recordset("cl_referen") = "PLAZOS----> DESDE:" & Format(data_afilcons.Recordset("fec_desde"), "dd/mm/yyyy") & " HASTA:" & Format(data_afilcons.Recordset("fec_hasta"), "dd/mm/yyyy")
      End If
      data_inf.Recordset("cl_tipclin") = data_afilcons.Recordset("tarj_sello")
      data_inf.Recordset("cl_email") = Mid(data_afilcons.Recordset("tarj_titular"), 1, 30)
      data_inf.Recordset("cl_nrovend") = data_afilcons.Recordset("tarj_cedtit")
      data_inf.Recordset("cl_forpago") = data_afilcons.Recordset("tarj_codced")
      data_inf.Recordset("cl_nomconv") = Mid(data_afilcons.Recordset("tarj_domic"), 1, 30)
      data_inf.Recordset("cl_nomcobr") = Mid(data_afilcons.Recordset("tarj_telef"), 1, 25)
      data_inf.Recordset("cl_nrotarj") = Mid(data_afilcons.Recordset("tarj_nro"), 1, 20)
      data_inf.Recordset("cl_ultmesp") = data_afilcons.Recordset("tarj_vencmes")
      data_inf.Recordset("cl_ultanop") = data_afilcons.Recordset("tarj_vencanio")
   Else
      If IsNull(data_afilcons.Recordset("debito_brou")) = False Then
         data_inf.Recordset("cl_nom_sup") = "Débito BROU"
         data_inf.Recordset("info_debit") = "CONFIRMA QUE REALIZÓ FORMULARIO PARA DÉBITO BROU?:--->SI" & chr(13)
         data_inf.Recordset("info_debit") = data_inf.Recordset("info_debit") & "NOMBRE DE TITULAR DE LA CUENTA:" & data_afilcons.Recordset("tarj_titular")
         If IsNull(data_afilcons.Recordset("codvende")) = False Then
            data_inf.Recordset("cl_entre") = Devuelve_vende()
         Else
            data_inf.Recordset("cl_entre") = "Sin promotor"
         End If
         If IsNull(data_afilcons.Recordset("fec_desde")) = False Then
            data_inf.Recordset("cl_referen") = "PLAZOS----> DESDE:" & Format(data_afilcons.Recordset("fec_desde"), "dd/mm/yyyy") & " HASTA:" & Format(data_afilcons.Recordset("fec_hasta"), "dd/mm/yyyy")
         End If
      Else
         data_inf.Recordset("cl_nom_sup") = "Cobrador a domicilio"
         data_inf.Recordset("info_debit") = "DOMICILIO DE COBRO:" & chr(13)
         If IsNull(data_afilcons.Recordset("direc_cobro")) = False Then
            data_inf.Recordset("info_debit") = data_inf.Recordset("info_debit") & data_afilcons.Recordset("direc_cobro") & chr(13)
            If IsNull(data_afilcons.Recordset("zonacobro")) = False Then
               data_inf.Recordset("info_debit") = data_inf.Recordset("info_debit") & "ZONA: " & data_afilcons.Recordset("zonacobro")
               If IsNull(data_afilcons.Recordset("dia_cobro")) = False Then
                  data_inf.Recordset("info_debit") = data_inf.Recordset("info_debit") & "----->FECHA DE COBRO: " & data_afilcons.Recordset("dia_cobro") & " c/mes" & chr(13)
               Else
                  data_inf.Recordset("info_debit") = data_inf.Recordset("info_debit") & "----->FECHA DE COBRO: sin datos." & chr(13)
               End If
            Else
               If IsNull(data_afilcons.Recordset("dia_cobro")) = False Then
                  data_inf.Recordset("info_debit") = data_inf.Recordset("info_debit") & "----->FECHA DE COBRO: " & data_afilcons.Recordset("dia_cobro") & " c/mes" & chr(13)
               Else
                  data_inf.Recordset("info_debit") = data_inf.Recordset("info_debit") & "----->FECHA DE COBRO: sin datos." & chr(13)
               End If
            End If
         Else
            data_inf.Recordset("info_debit") = "Misma dirección." & chr(13)
         End If
         If IsNull(data_afilcons.Recordset("codvende")) = False Then
            data_inf.Recordset("cl_entre") = Devuelve_vende()
         Else
            data_inf.Recordset("cl_entre") = "Sin promotor"
         End If
         If IsNull(data_afilcons.Recordset("fec_desde")) = False Then
            data_inf.Recordset("cl_referen") = "PLAZOS----> DESDE:" & Format(data_afilcons.Recordset("fec_desde"), "dd/mm/yyyy") & " HASTA:" & Format(data_afilcons.Recordset("fec_hasta"), "dd/mm/yyyy")
         End If
         
      End If
   End If
   data_inf.Recordset("obsp") = Xcontrato
   data_inf.Recordset.Update
Else
   MsgBox "No hay datos de afiliación para imprimir. Verifique!", vbCritical
   
End If

End Sub

Public Function Devuelve_titular() As String
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from afiliaciones_new where afilia_nro =" & data_afilcons.Recordset("afilia_nro") & " and integra_nro in (1)"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Devuelve_titular = Xrecclii("nom1")
Else
   Devuelve_titular = ""
End If

Xrecclii.Close
ConbdSapp.Close


End Function
Public Function Devuelve_titularApe() As String
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from afiliaciones_new where afilia_nro =" & data_afilcons.Recordset("afilia_nro") & " and integra_nro in (1)"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Devuelve_titularApe = Xrecclii("ape1")
Else
   Devuelve_titularApe = ""
End If

Xrecclii.Close
ConbdSapp.Close


End Function

Public Function Devuelve_mut() As String

Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from ca_adm where id =" & data_afilcons.Recordset("codmut")
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Devuelve_mut = Xrecclii("ca_nom")
Else
   Devuelve_mut = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Function

Public Function Devuelve_vende() As String

Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from vende_func where idfunc =" & data_afilcons.Recordset("codvende")
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Devuelve_vende = Xrecclii("nombre")
Else
   Devuelve_vende = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Function

Private Sub List1_DblClick()
    
    Command3.Enabled = True
    Command4.Enabled = True
    Command2.Enabled = True
    List1.Clear
    List1.Visible = False
    
End Sub

Private Sub ms2_DblClick()
Dim Xsiquieromodif As String

If Val(Label3.Caption) <> 10 Then
   If Val(Label3.Caption) = 2 Then
      MsgBox "Afiliación pendiente de autorización.", vbCritical
   Else
      Xsiquieromodif = MsgBox("Desea modificar datos de la afiliación?", vbExclamation + vbYesNo, "Afiliaciones SAPP")
      data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(DBGrid1.TextMatrix(DBGrid1.RowSel, 1)) & " and integra_nro =" & Val(ms2.TextMatrix(ms2.RowSel, 0)) & " and wusuario ='" & WElusuario & "'"
      data_afilcons.Refresh
      If data_afilcons.Recordset.RecordCount > 0 Then
         frm_afilia.Label3.Caption = Format(data_afilcons.Recordset("fecha"), "dd/mm/yyyy")
         frm_afilia.labnro.Caption = data_afilcons.Recordset("afilia_nro")
         frm_afilia.labintegra.Caption = data_afilcons.Recordset("integra_nro")
         frm_afilia.t_nom1.Text = data_afilcons.Recordset("nom1")
         If IsNull(data_afilcons.Recordset("nom2")) = False Then
            frm_afilia.t_nom2.Text = data_afilcons.Recordset("nom2")
         Else
            frm_afilia.t_nom2.Text = ""
         End If
         frm_afilia.t_ape1.Text = data_afilcons.Recordset("ape1")
         If IsNull(data_afilcons.Recordset("ape2")) = False Then
            frm_afilia.t_ape2.Text = data_afilcons.Recordset("ape2")
         Else
            frm_afilia.t_ape2.Text = ""
         End If
         frm_afilia.mfnac.Text = data_afilcons.Recordset("fnac")
         frm_afilia.t_telef.Text = data_afilcons.Recordset("telef")
         frm_afilia.t_celu.Text = data_afilcons.Recordset("celular")
         frm_afilia.t_correo.Text = data_afilcons.Recordset("correo")
         frm_afilia.t_calle.Text = data_afilcons.Recordset("direc1")
         If IsNull(data_afilcons.Recordset("direc2")) = False Then
            frm_afilia.t_entre.Text = data_afilcons.Recordset("direc2")
         Else
            frm_afilia.t_entre.Text = ""
         End If
         If IsNull(data_afilcons.Recordset("manz")) = False Then
            frm_afilia.t_manz.Text = data_afilcons.Recordset("manz")
         Else
            frm_afilia.t_manz.Text = ""
         End If
         If IsNull(data_afilcons.Recordset("solar")) = False Then
            frm_afilia.t_sol.Text = data_afilcons.Recordset("solar")
         Else
            frm_afilia.t_sol.Text = ""
         End If
         If IsNull(data_afilcons.Recordset("casa")) = False Then
            frm_afilia.t_casa.Text = data_afilcons.Recordset("casa")
         Else
            frm_afilia.t_casa.Text = ""
         End If
         frm_afilia.labcodzon.Caption = data_afilcons.Recordset("codzon")
         frm_afilia.cbozona.Text = data_afilcons.Recordset("nomzona")
         frm_afilia.labcodmut.Caption = data_afilcons.Recordset("codmut")
         Consulta_mutual
         If IsNull(data_afilcons.Recordset("sexo")) = False Then
            If data_afilcons.Recordset("sexo") = 1 Then
               frm_afilia.cbosexo.ListIndex = 1
            Else
               frm_afilia.cbosexo.ListIndex = 0
            End If
         End If
         If Xsiquieromodif = vbYes Then
            frm_afilia.Command3.Visible = True
            frm_afilia.b_imp.Visible = False
            frm_afilia.Frame1.Enabled = True
            frm_afilia.cbocat.Enabled = False
            frm_afilia.cbopromo.Enabled = False
            frm_afilia.cbovende.Enabled = False
         Else
            frm_afilia.cbocat.Enabled = True
            frm_afilia.cbopromo.Enabled = True
            frm_afilia.cbovende.Enabled = True
            frm_afilia.Command3.Visible = False
            frm_afilia.Frame1.Enabled = True
         
         End If
         Unload Me
      Else
         MsgBox "Verifique si es el usuario creador de la afiliación.", vbInformation
      End If
   End If
Else
   MsgBox "Afiliación ya ingresada y confirmada en el sistema.", vbCritical
End If

End Sub

Public Sub Consulta_mutual()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from ca_adm where id =" & data_afilcons.Recordset("codmut")
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   frm_afilia.cbomut.Text = Xrecclii("ca_nom")
Else
   frm_afilia.cbomut.Text = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Public Function Devuelve_titularCed() As String
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from afiliaciones_new where afilia_nro =" & data_afilcons.Recordset("afilia_nro") & " and integra_nro in (1)"
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   If Len(Xrecclii("cedula")) = 7 Then
      Devuelve_titularCed = Mid(Trim(str(Xrecclii("cedula"))), 1, 6) & "-" & Mid(Trim(str(Xrecclii("cedula"))), 7, 1)
   Else
      Devuelve_titularCed = Mid(Trim(str(Xrecclii("cedula"))), 1, 7) & "-" & Mid(Trim(str(Xrecclii("cedula"))), 8, 1)
   End If
Else
   Devuelve_titularCed = ""
End If

Xrecclii.Close
ConbdSapp.Close


End Function


Public Sub Carga_grid_cancelado()
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Xcann, Xpendiente As Integer
Dim Ladate As Date
Ladate = Date - 30
Xpendiente = 0
ConectarBD
ConbdSapp.Open
             
'Data1.ConnectionString = "dsn=sappnew"
If labpermiso.Caption = "1" Then
   If t_busca.Text = "" Then
      Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.nom1,afiliaciones_new.ape1,afiliaciones_new.pendiente,afiliaciones_new.integra_nro,afiliaciones_new.wusuario," & _
      "afiliaciones_new.convenio,afiliaciones_new.celular,afiliaciones_new.telef,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
      "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.fecha >='" & Format(Ladate, "yyyy/mm/dd") & "' and afiliaciones_new.pendiente in (20) and afiliaciones_new.integra_nro in (1) order by afiliaciones_new.fecha DESC"
   Else
      If Combo1.Text = "CEDULA" Then
         Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.nom1,afiliaciones_new.ape1,afiliaciones_new.pendiente,afiliaciones_new.integra_nro,afiliaciones_new.wusuario," & _
         "afiliaciones_new.cedula,afiliaciones_new.convenio,afiliaciones_new.celular,afiliaciones_new.telef,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
         "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.pendiente in (20) and afiliaciones_new.integra_nro in (1) and afiliaciones_new.cedula ='" & t_busca.Text & "'  order by afiliaciones_new.fecha DESC"
      Else
         If Combo1.Text = "NRO.AFILIACION" Then
            Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.nom1,afiliaciones_new.ape1,afiliaciones_new.pendiente,afiliaciones_new.integra_nro,afiliaciones_new.wusuario," & _
            "afiliaciones_new.cedula,afiliaciones_new.convenio,afiliaciones_new.celular,afiliaciones_new.telef,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
            "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.integra_nro in (1) and afiliaciones_new.afilia_nro =" & t_busca.Text & " and afiliaciones_new.pendiente in (20) order by afiliaciones_new.fecha DESC"
         Else
            If Combo1.Text = "FECHA" Then
               Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.nom1,afiliaciones_new.ape1,afiliaciones_new.pendiente,afiliaciones_new.integra_nro,afiliaciones_new.wusuario," & _
               "afiliaciones_new.convenio,afiliaciones_new.celular,afiliaciones_new.telef,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
               "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.fecha >='" & Format(t_busca.Text, "yyyy/mm/dd") & "' and afiliaciones_new.integra_nro in (1) and afiliaciones_new.pendiente in (20) order by afiliaciones_new.fecha DESC"
            Else
               Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.nom1,afiliaciones_new.ape1,afiliaciones_new.pendiente,afiliaciones_new.integra_nro,afiliaciones_new.wusuario," & _
               "afiliaciones_new.convenio,afiliaciones_new.celular,afiliaciones_new.telef,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
               "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.fecha >='" & Format(Ladate, "yyyy/mm/dd") & "' and afiliaciones_new.integra_nro in (1) and afiliaciones_new.pendiente in (20) order by afiliaciones_new.fecha DESC"
            End If
         End If
      End If
   End If
Else
   If t_busca.Text = "" Then
      Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.nom1,afiliaciones_new.ape1,afiliaciones_new.pendiente,afiliaciones_new.integra_nro,afiliaciones_new.wusuario," & _
      "afiliaciones_new.convenio,afiliaciones_new.celular,afiliaciones_new.telef,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
      "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.fecha >='" & Format(Ladate, "yyyy/mm/dd") & "' and afiliaciones_new.pendiente in (20) and afiliaciones_new.integra_nro in (1) order by afiliaciones_new.fecha DESC"
   Else
      If Combo1.Text = "CEDULA" Then
         Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.nom1,afiliaciones_new.ape1,afiliaciones_new.pendiente,afiliaciones_new.integra_nro,afiliaciones_new.wusuario," & _
         "afiliaciones_new.cedula,afiliaciones_new.convenio,afiliaciones_new.celular,afiliaciones_new.telef,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
         "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.pendiente in (20) and afiliaciones_new.integra_nro in (1) and afiliaciones_new.cedula ='" & t_busca.Text & "'  order by afiliaciones_new.fecha DESC"
      Else
         If Combo1.Text = "NRO.AFILIACION" Then
            Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.nom1,afiliaciones_new.ape1,afiliaciones_new.pendiente,afiliaciones_new.integra_nro,afiliaciones_new.wusuario," & _
            "afiliaciones_new.cedula,afiliaciones_new.convenio,afiliaciones_new.celular,afiliaciones_new.telef,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
            "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.integra_nro in (1) and afiliaciones_new.afilia_nro =" & t_busca.Text & " and afiliaciones_new.pendiente in (20) order by afiliaciones_new.fecha DESC"
         Else
            If Combo1.Text = "FECHA" Then
               Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.nom1,afiliaciones_new.ape1,afiliaciones_new.pendiente,afiliaciones_new.integra_nro,afiliaciones_new.wusuario," & _
               "afiliaciones_new.convenio,afiliaciones_new.celular,afiliaciones_new.telef,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
               "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.fecha >='" & Format(t_busca.Text, "yyyy/mm/dd") & "' and afiliaciones_new.integra_nro in (1) and afiliaciones_new.pendiente in (20) order by afiliaciones_new.fecha DESC"
            Else
               Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.nom1,afiliaciones_new.ape1,afiliaciones_new.pendiente,afiliaciones_new.integra_nro,afiliaciones_new.wusuario," & _
               "afiliaciones_new.convenio,afiliaciones_new.celular,afiliaciones_new.telef,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
               "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.fecha >='" & Format(Ladate, "yyyy/mm/dd") & "' and afiliaciones_new.integra_nro in (1) and afiliaciones_new.pendiente in (20) order by afiliaciones_new.fecha DESC"
            End If
         End If
      End If
   End If

End If


With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
DBGrid1.Clear
DBGrid1.rows = 2
DBGrid1.Cols = 9
DBGrid1.TextMatrix(0, 0) = "FECHA"
DBGrid1.ColWidth(0) = 1300
DBGrid1.TextMatrix(0, 1) = "Nro.Af."
DBGrid1.ColWidth(1) = 1300
DBGrid1.TextMatrix(0, 2) = "USUARIO"
DBGrid1.ColWidth(2) = 1500

DBGrid1.TextMatrix(0, 3) = "NOMBRE"
DBGrid1.ColWidth(3) = 2900
DBGrid1.TextMatrix(0, 4) = "CATEGORIA"
DBGrid1.ColWidth(4) = 1500
DBGrid1.TextMatrix(0, 5) = "CELULAR"
DBGrid1.ColWidth(5) = 1500
DBGrid1.TextMatrix(0, 6) = "TELEFONO"
DBGrid1.ColWidth(6) = 1500
DBGrid1.TextMatrix(0, 7) = "PROMOTOR"
DBGrid1.ColWidth(7) = 1500
DBGrid1.TextMatrix(0, 8) = "ESTADO ACTUAL"
DBGrid1.ColWidth(8) = 1700

Xcann = 1

If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   Do While Not Xrecclii.EOF
      DBGrid1.TextMatrix(Xcann, 0) = Xrecclii("fecha")
      DBGrid1.TextMatrix(Xcann, 1) = Xrecclii("afilia_nro")
      DBGrid1.TextMatrix(Xcann, 2) = Xrecclii("wusuario")
      
      DBGrid1.TextMatrix(Xcann, 3) = Xrecclii("nom1") & " " & Xrecclii("ape1")
      DBGrid1.TextMatrix(Xcann, 4) = Xrecclii("convenio")
      If IsNull(Xrecclii("celular")) = False Then
         DBGrid1.TextMatrix(Xcann, 5) = Xrecclii("celular")
      End If
      If IsNull(Xrecclii("telef")) = False Then
         DBGrid1.TextMatrix(Xcann, 6) = Xrecclii("telef")
      End If
      If IsNull(Xrecclii("nombre")) = False Then
         DBGrid1.TextMatrix(Xcann, 7) = Xrecclii("nombre")
      End If
      Xpendiente = Xrecclii("pendiente")
      If Xpendiente = 20 Then
         DBGrid1.TextMatrix(Xcann, 8) = "CANCELADA"
      Else
         DBGrid1.TextMatrix(Xcann, 8) = "VERIFICAR"
      End If
      
      DBGrid1.rows = DBGrid1.rows + 1
      Xrecclii.MoveNext
      Xcann = Xcann + 1
   Loop
End If

Xrecclii.Close
ConbdSapp.Close

End Sub

Private Sub t_cedbusca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Trim(t_cedbusca.Text) <> "" Then
   
       Dim Xsqlpromo As String
       Dim Xrecclii As New ADODB.Recordset
       Dim Xcann, Xpendiente As Integer
       Dim Ladate As Date
       Dim NroAf As Long
       NroAf = 0
       Ladate = Date - 30
       Xpendiente = 0
       ConectarBD
       ConbdSapp.Open
       
       Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.nom1,afiliaciones_new.ape1,afiliaciones_new.pendiente,afiliaciones_new.integra_nro,afiliaciones_new.wusuario," & _
       "afiliaciones_new.cedula,afiliaciones_new.convenio,afiliaciones_new.celular,afiliaciones_new.telef,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
       "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.pendiente not in (20) and afiliaciones_new.cedula ='" & t_cedbusca.Text & "'  order by afiliaciones_new.fecha DESC"
    
       With Xrecclii
           .CursorLocation = adUseClient
           .CursorType = adOpenKeyset
           .LockType = adLockOptimistic
           .Open Xsqlpromo, ConbdSapp, , , adCmdText
       End With
       If Xrecclii.RecordCount > 0 Then
          NroAf = Xrecclii("afilia_nro")
          Xsqlpromo = "Select afiliaciones_new.fecha,afiliaciones_new.afilia_nro,afiliaciones_new.nom1,afiliaciones_new.ape1,afiliaciones_new.pendiente,afiliaciones_new.integra_nro,afiliaciones_new.wusuario," & _
          "afiliaciones_new.cedula,afiliaciones_new.convenio,afiliaciones_new.celular,afiliaciones_new.telef,afiliaciones_new.codvende,vende_func.idfunc,vende_func.nombre " & _
          "from afiliaciones_new inner join vende_func on afiliaciones_new.codvende=vende_func.idfunc where afiliaciones_new.pendiente not in (20) and afiliaciones_new.integra_nro in (1) and afiliaciones_new.afilia_nro ='" & NroAf & "'  order by afiliaciones_new.fecha DESC"
    
            DBGrid1.Clear
            DBGrid1.rows = 2
            DBGrid1.Cols = 9
            DBGrid1.TextMatrix(0, 0) = "FECHA"
            DBGrid1.ColWidth(0) = 1300
            DBGrid1.TextMatrix(0, 1) = "Nro.Af."
            DBGrid1.ColWidth(1) = 1300
            DBGrid1.TextMatrix(0, 2) = "USUARIO"
            DBGrid1.ColWidth(2) = 1500
            
            DBGrid1.TextMatrix(0, 3) = "NOMBRE"
            DBGrid1.ColWidth(3) = 2900
            DBGrid1.TextMatrix(0, 4) = "CATEGORIA"
            DBGrid1.ColWidth(4) = 1500
            DBGrid1.TextMatrix(0, 5) = "CELULAR"
            DBGrid1.ColWidth(5) = 1500
            DBGrid1.TextMatrix(0, 6) = "TELEFONO"
            DBGrid1.ColWidth(6) = 1500
            DBGrid1.TextMatrix(0, 7) = "PROMOTOR"
            DBGrid1.ColWidth(7) = 1500
            DBGrid1.TextMatrix(0, 8) = "ESTADO ACTUAL"
            DBGrid1.ColWidth(8) = 1700
            
            Xcann = 1
            
            If Xrecclii.RecordCount > 0 Then
               Xrecclii.MoveFirst
               Do While Not Xrecclii.EOF
                  DBGrid1.TextMatrix(Xcann, 0) = Xrecclii("fecha")
                  DBGrid1.TextMatrix(Xcann, 1) = Xrecclii("afilia_nro")
                  DBGrid1.TextMatrix(Xcann, 2) = Xrecclii("wusuario")
                  
                  DBGrid1.TextMatrix(Xcann, 3) = Xrecclii("nom1") & " " & Xrecclii("ape1")
                  DBGrid1.TextMatrix(Xcann, 4) = Xrecclii("convenio")
                  If IsNull(Xrecclii("celular")) = False Then
                     DBGrid1.TextMatrix(Xcann, 5) = Xrecclii("celular")
                  End If
                  If IsNull(Xrecclii("telef")) = False Then
                     DBGrid1.TextMatrix(Xcann, 6) = Xrecclii("telef")
                  End If
                  If IsNull(Xrecclii("nombre")) = False Then
                     DBGrid1.TextMatrix(Xcann, 7) = Xrecclii("nombre")
                  End If
                  Xpendiente = Xrecclii("pendiente")
                  If Xpendiente = 0 Then
                     DBGrid1.TextMatrix(Xcann, 8) = "PENDIENTE"
                  Else
                     If Xpendiente = 2 Then
                        DBGrid1.TextMatrix(Xcann, 8) = "FALTA AUTORIZAR"
                     Else
                        If Xpendiente = 3 Then
                           DBGrid1.TextMatrix(Xcann, 8) = "FALTA VERIFICAR"
                        Else
                           If Xpendiente = 4 Then
                              DBGrid1.TextMatrix(Xcann, 8) = "FALTA FACTURAR"
                              If Xpendiente = 10 Then
                                  DBGrid1.TextMatrix(Xcann, 8) = "PROCESADA"
                              Else
                                  DBGrid1.TextMatrix(Xcann, 8) = "SIN DATOS"
                              End If
                           End If
                        End If
                     End If
                  End If
                  
                  DBGrid1.rows = DBGrid1.rows + 1
                  Xrecclii.MoveNext
                  Xcann = Xcann + 1
               Loop
            End If
            
            Xrecclii.Close
            ConbdSapp.Close
        Else
            MsgBox "No se encuentra la cédula en afiliaciones activas.", vbInformation
            Xrecclii.Close
            ConbdSapp.Close
        End If
    Else
        Carga_grid
    End If
    t_busca.SetFocus
End If

End Sub
