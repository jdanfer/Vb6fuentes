VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_ctrolsald 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación de saldos de clientes"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5445
   Icon            =   "frm_ctrolsald.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5445
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox t_mat 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3120
      TabIndex        =   11
      Top             =   840
      Width           =   1815
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Procesar mes atrasado"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3120
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   4560
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   2100
   End
   Begin Crystal.CrystalReport cr11 
      Left            =   4320
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_infdeu 
      Caption         =   "data_infdeu"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
      Visible         =   0   'False
      Width           =   2580
   End
   Begin VB.Data data_lin 
      Caption         =   "data_lin"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1920
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Data data_deu 
      Caption         =   "data_deu"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Listar deudas al terminar."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Visible         =   0   'False
      Width           =   3495
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox tan 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox tme 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Salir..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3360
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Procesar..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3360
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Procesar solo para la matrícula:"
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Este proceso puede tardar varios minutos."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mes/Año de última emisión:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frm_ctrolsald"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Dim Xmat, Xmes, Xano, Xsaldo, Xcant, Xbancuantos As Long
'data_cli.DatabaseName = App.Path & "\sapp.mdb"
'data_cli.RecordSource = "Select * from clientes where estado =" & 1
'data_cli.Refresh
'data_cli.RecordSource = "clientes"
'data_cli.Refresh
'data_deu.DatabaseName = App.Path & "\sapp.mdb"
'data_deu.RecordSource = "Select * from deudas order by cliente,ano,mes"
'data_deu.Refresh
'data_deu.Recordset.MoveLast
'pb1.Max = pb1.Value + data_deu.Recordset.RecordCount + 100
'data_deu.Recordset.MoveFirst
'DoEvents
'Xmat = data_deu.Recordset("cliente")
'Xmes = 0
'Xano = 0
'Xbancuantos = 0
Command3_Click

'Do While Not data_deu.Recordset.EOF
'   If Xmat = data_deu.Recordset("cliente") Then
'      If Xbancuantos = 0 Then
'         If IsNull(data_deu.Recordset("mes")) = True Then
'            If Month(Date) = 1 Then
'               Xmes = 11
'               Xano = Year(Date) - 1
'            Else
'               Xmes = Month(Date) - 3
'               Xano = Year(Date)
'            End If
'         Else
'            If data_deu.Recordset("mes") = 0 Then
'               If Month(Date) = 1 Then
'                  Xmes = 11
'                  Xano = Year(Date) - 1
'               Else
'                  Xmes = Month(Date) - 3
'                  Xano = Year(Date)
'               End If
'            Else
'               If data_deu.Recordset("mes") = 1 Then
'                  Xmes = 12
'                  Xano = data_deu.Recordset("ano") - 1
'               Else
'                  Xmes = data_deu.Recordset("mes") - 1
'                  Xano = data_deu.Recordset("ano")
'               End If
'            End If
'         End If
'         Xbancuantos = 1
'      End If
''25673
'      If IsNull(data_deu.Recordset("total")) = False And IsNull(data_deu.Recordset("fecha_pago")) = True Then
'         Xsaldo = Xsaldo + data_deu.Recordset("total")
'         If data_deu.Recordset("mes") > 0 Then
'            Xcant = Xcant + 1
'         End If
'         If IsNull(data_deu.Recordset("saldo_cc")) = False Then
'            If data_deu.Recordset("saldo_cc") <> Xsaldo Then
'               data_deu.Recordset.Edit
'               data_deu.Recordset("saldo_cc") = Xsaldo
'               data_deu.Recordset.Update
'            End If
'         Else
'            data_deu.Recordset.Edit
'            data_deu.Recordset("saldo_cc") = Xsaldo
'            data_deu.Recordset.Update
'         End If
'         Xmat = data_deu.Recordset("cliente")
'         data_deu.Recordset.MoveNext
'      Else
'         Xsaldo = Xsaldo + 0
'         If data_deu.Recordset("mes") > 0 Then
'            Xcant = Xcant + 1
'         End If
'         If data_deu.Recordset("saldo_cc") <> Xsaldo Then
'            data_deu.Recordset.Edit
'            data_deu.Recordset("saldo_cc") = Xsaldo
'            data_deu.Recordset.Update
'         End If
'         Xmat = data_deu.Recordset("cliente")
'         data_deu.Recordset.MoveNext
'      End If
'   Else
'      Xbancuantos = 0
'      data_deu.Recordset.MovePrevious
'      If IsNull(data_deu.Recordset("saldo_cc")) = False Then
''         If data_deu.Recordset("saldo_cc") <> Xsaldo Then
 '           data_deu.Recordset.Edit
 '           data_deu.Recordset("saldo_cc") = Xsaldo
 '           data_deu.Recordset.Update
 '        End If
 '     Else
'         data_deu.Recordset.Edit
'         data_deu.Recordset("saldo_cc") = Xsaldo
'         data_deu.Recordset.Update
'      End If
'      data_cli.RecordSource = "Select * from clientes where cl_codigo =" & Xmat
'      data_cli.Refresh
'      If data_cli.Recordset.RecordCount > 0 Then
'         If data_cli.Recordset("cl_ultmesp") = Xmes And data_cli.Recordset("cl_ultanop") = Xano Then
'            If IsNull(data_cli.Recordset("cl_atrasoa")) = False Then
'               If data_cli.Recordset("cl_atrasoa") = Xcant Then
'                  If IsNull(data_cli.Recordset("saldo_cc")) = True Then
'                     data_cli.Recordset.Edit
'                     data_cli.Recordset("saldo_cc") = Xsaldo
'                     data_cli.Recordset.Update
'                  Else
'                     If data_cli.Recordset("saldo_cc") = Xsaldo Then
'                     Else
'                        data_cli.Recordset.Edit
'                        data_cli.Recordset("saldo_cc") = Xsaldo
'                        data_cli.Recordset.Update
'                     End If
'                  End If
'               Else
'                  If data_cli.Recordset("saldo_cc") <> Xsaldo Then
'                     data_cli.Recordset.Edit
'                     data_cli.Recordset("cl_atrasoa") = Xcant
'                     data_cli.Recordset("saldo_cc") = Xsaldo
'                     data_cli.Recordset.Update
'                  End If
'               End If
'            Else
'               If data_cli.Recordset("saldo_cc") <> Xsaldo Then
'                  data_cli.Recordset.Edit
'                  data_cli.Recordset("cl_atrasoa") = Xcant
'                  data_cli.Recordset("saldo_cc") = Xsaldo
'                  data_cli.Recordset.Update
'               End If
'            End If
'         Else
'            If data_cli.Recordset("cl_atrasoa") <> Xcant Then
'               data_cli.Recordset.Edit
'               data_cli.Recordset("cl_atrasoa") = Xcant
'               data_cli.Recordset.Update
'            End If
'            If data_cli.Recordset("saldo_cc") <> Xsaldo Then
'               data_cli.Recordset.Edit
'               data_cli.Recordset("saldo_cc") = Xsaldo
'               data_cli.Recordset.Update
'            End If
'            If data_cli.Recordset("cl_ultmesp") <> Xmes Then
'               data_cli.Recordset.Edit
'               data_cli.Recordset("cl_ultmesp") = Xmes
'               data_cli.Recordset("cl_ultanop") = Xano
'               data_cli.Recordset.Update
'            End If
'         End If
'      End If
'      If data_deu.Recordset("nro_cobr") = 615 Or data_deu.Recordset("nro_cobr") = 616 Or _
'         data_deu.Recordset("nro_cobr") = 636 Or _
'         data_deu.Recordset("nro_cobr") = 635 Or _
'         data_deu.Recordset("nro_cobr") = 602 Or _
'         data_deu.Recordset("nro_cobr") = 653 Or _
'         data_deu.Recordset("nro_cobr") = 672 Or _
'         data_deu.Recordset("nro_cobr") = 113 Or _
'         data_deu.Recordset("nro_cobr") = 685 Or _
'         data_deu.Recordset("nro_cobr") = 10 Or _
'         data_deu.Recordset("nro_cobr") = 1 Or _
''         data_deu.Recordset("nro_cobr") = 679 Or _
 '        data_deu.Recordset("nro_cobr") = 685 Or _
 '        data_deu.Recordset("nro_cobr") = 512 Or _
 '        data_deu.Recordset("nro_cobr") = 201 Or _
 '        data_deu.Recordset("nro_cobr") = 208 Or _
 '        data_deu.Recordset("nro_cobr") = 209 Then
 '        If Xcant = 1 Then
'            data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & data_deu.Recordset("cliente") & " And mes_paga =" & tme.Text & " And ano_paga =" & tan.Text
'            data_lin.Refresh
'            If data_lin.Recordset.RecordCount > 0 Then
'               data_deu.Recordset.Delete
'               If IsNull(data_cli.Recordset("cl_atrasoa")) = False Then
'                  If data_cli.Recordset("cl_atrasoa") <> 0 Then
'                    data_cli.Recordset.Edit
'                    data_cli.Recordset("cl_atrasoa") = 0
'                    data_cli.Recordset("saldo_cc") = 0
'                    data_cli.Recordset("cl_ultmesp") = tme.Text
'                    data_cli.Recordset("cl_ultanop") = tan.Text
'                    data_cli.Recordset.Update
'                  End If
'               End If
'            End If
'         End If
'      End If
'      data_deu.Recordset.MoveNext
'      Xsaldo = 0
'      Xcant = 0
'      Xmes = 0
'      Xano = 0
'      Xmat = data_deu.Recordset("cliente")
'   End If
'   pb1.Value = pb1.Value + 1
'Loop
'Xbancuantos = 0
'DoEvents
'data_deu.Recordset.MovePrevious
'If IsNull(data_deu.Recordset("saldo_cc")) = False Then
'   If data_deu.Recordset("saldo_cc") <> Xsaldo Then
'      data_deu.Recordset.Edit
'      data_deu.Recordset("saldo_cc") = Xsaldo
'      data_deu.Recordset.Update
'   End If
'Else
'   data_deu.Recordset.Edit
'   data_deu.Recordset("saldo_cc") = Xsaldo
'   data_deu.Recordset.Update
'End If
''data_cli.Recordset.FindFirst "cl_codigo =" & Xmat
'data_cli.RecordSource = "Select * from clientes where cl_codigo =" & Xmat
'data_cli.Refresh
'If data_cli.Recordset.RecordCount > 0 Then
''''''''''''If Not data_cli.Recordset.NoMatch Then
'   If data_cli.Recordset("cl_ultmesp") = Xmes And data_cli.Recordset("cl_ultanop") = Xano Then
'      If IsNull(data_cli.Recordset("cl_atrasoa")) = False Then
'         If data_cli.Recordset("cl_atrasoa") = Xcant Then
'            If IsNull(data_cli.Recordset("saldo_cc")) = True Then
'               data_cli.Recordset.Edit
'               data_cli.Recordset("saldo_cc") = Xsaldo
'               data_cli.Recordset.Update
'            Else
'               If data_cli.Recordset("saldo_cc") = Xsaldo Then
'               Else
'                  data_cli.Recordset.Edit
'                  data_cli.Recordset("saldo_cc") = Xsaldo
'                  data_cli.Recordset.Update
'               End If
'            End If
'         Else
'            data_cli.Recordset.Edit
'            data_cli.Recordset("cl_atrasoa") = Xcant
'            data_cli.Recordset("saldo_cc") = Xsaldo
'            data_cli.Recordset.Update
'         End If
'      Else
'         data_cli.Recordset.Edit
'         data_cli.Recordset("cl_atrasoa") = Xcant
'         data_cli.Recordset("saldo_cc") = Xsaldo
'         data_cli.Recordset.Update
'      End If
'   Else
'      data_cli.Recordset.Edit
'      If data_cli.Recordset("cl_atrasoa") <> Xcant Then
'         data_cli.Recordset("cl_atrasoa") = Xcant
''      End If
 '     If data_cli.Recordset("saldo_cc") <> Xsaldo Then
 '        data_cli.Recordset("saldo_cc") = Xsaldo
 '     End If
'      data_cli.Recordset("cl_ultmesp") = Xmes
'      data_cli.Recordset("cl_ultanop") = Xano
'      data_cli.Recordset.Update
'   End If
'End If
'If data_deu.Recordset("nro_cobr") = 615 Or data_deu.Recordset("nro_cobr") = 616 Or _
'   data_deu.Recordset("nro_cobr") = 636 Or _
'   data_deu.Recordset("nro_cobr") = 635 Or _
'   data_deu.Recordset("nro_cobr") = 602 Or _
'   data_deu.Recordset("nro_cobr") = 653 Or _
'   data_deu.Recordset("nro_cobr") = 672 Or _
'   data_deu.Recordset("nro_cobr") = 113 Or _
'   data_deu.Recordset("nro_cobr") = 685 Or _
'   data_deu.Recordset("nro_cobr") = 10 Or _
'   data_deu.Recordset("nro_cobr") = 1 Or _
'   data_deu.Recordset("nro_cobr") = 679 Or _
'   data_deu.Recordset("nro_cobr") = 685 Or _
'   data_deu.Recordset("nro_cobr") = 512 Or _
'   data_deu.Recordset("nro_cobr") = 201 Or _
'   data_deu.Recordset("nro_cobr") = 208 Or _
'   data_deu.Recordset("nro_cobr") = 209 Then
'   If Xcant = 1 Then
'      data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & data_deu.Recordset("cliente") & " And mes_paga =" & tme.Text & " And ano_paga =" & tan.Text
'      data_lin.Refresh
'      If data_lin.Recordset.RecordCount > 0 Then
'         data_deu.Recordset.Delete
'         data_cli.Recordset.Edit
'         data_cli.Recordset("cl_atrasoa") = 0
'         data_cli.Recordset("saldo_cc") = 0
'         data_cli.Recordset("cl_ultmesp") = txt_mes.Text
'         data_cli.Recordset("cl_ultanop") = txt_ano.Text
'         data_cli.Recordset.Update
'      End If
'   End If
'End If
'pb1.Value = pb1.Value + 99

'MsgBox "Proceso de generación de saldos terminado."

'If Check1.Value = 1 Then
'   data_infdeu.DatabaseName = App.Path & "\informes.mdb"
'   data_infdeu.RecordSource = "infemis"
'   data_infdeu.Refresh
'   If data_infdeu.Recordset.RecordCount > 0 Then
'      data_infdeu.Recordset.MoveFirst
'      Do While Not data_infdeu.Recordset.EOF
'         data_infdeu.Recordset.Delete
'         data_infdeu.Recordset.MoveNext
'      Loop
'   End If
'   data_deu.Recordset.MoveFirst
'   Do While Not data_deu.Recordset.EOF
'      data_infdeu.Recordset.AddNew
'      data_infdeu.Recordset("cliente") = data_deu.Recordset("cliente")
'      data_infdeu.Recordset("nombre") = data_deu.Recordset("nombre")
'      data_infdeu.Recordset("cod_cnv") = data_deu.Recordset("cod_cnv")
'      data_infdeu.Recordset("nro_cobr") = data_deu.Recordset("nro_cobr")
'      data_infdeu.Recordset("nom_cobr") = data_deu.Recordset("nom_cobr")
'      data_infdeu.Recordset("importe") = data_deu.Recordset("importe")
'      data_infdeu.Recordset("deudas") = data_deu.Recordset("deudas")
'      data_infdeu.Recordset("iva") = data_deu.Recordset("iva")
'      data_infdeu.Recordset("tiquet") = data_deu.Recordset("tiquet")
'      data_infdeu.Recordset("total") = data_deu.Recordset("total")
'      data_infdeu.Recordset.Update
'      data_deu.Recordset.MoveNext
'   Loop
'   data_infdeu.RecordSource = "Select * from infdeu"
'   data_infdeu.Refresh
'   cr11.ReportFileName = App.Path & "\infdeuda2.rpt"
'   cr11.Action = 1
'
'End If

' Cargar a Deudas de mysql
'Command1.Enabled = False
'Command2.Enabled = True

'frm_procdeu.MousePointer = 0

'MsgBox "Proceso terminado"

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
Dim Xmat, Xmes, Xano, Xsaldo, Xcant, Xbancuantos As Long
Dim Fec1, Fec2 As Date
Dim TextFec1, TextFec2 As String

'data_cli.DatabaseName = App.Path & "\sapp.mdb"
data_cli.Connect = "ODBC;DSN=sapp;"

'data_cli.RecordSource = "Select * from clientes where cl_codconv not in ('PART','UCM')"
'data_cli.RecordSource = "Select * from clientes where cl_codigo =" & 10128126
'data_cli.Refresh

Command1.Enabled = False
Command2.Enabled = False

frm_ctrolsald.MousePointer = 11

'If data_cli.Recordset.RecordCount > 0 Then
'   data_cli.Recordset.MoveFirst
'   Do While Not data_cli.Recordset.EOF
'      If IsNull(data_cli.Recordset("saldo_cc")) = False Then
'         If data_cli.Recordset("saldo_cc") > 0 Then
'            data_cli.Recordset.Edit
'            data_cli.Recordset("saldo_cc") = 0
'            data_cli.Recordset("cl_atrasoa") = 0
'            data_cli.Recordset("cl_ultmesp") = 0
'            data_cli.Recordset("cl_ultanop") = 0
'            data_cli.Recordset.Update
'         End If
'      End If
 '  Loop
 '  data_cli.Refresh
'End If
If t_mat.Text <> "" Then
   data_deu.RecordSource = "Select * from deudas where cliente =" & t_mat.Text & " and fecha_pago is null and mes >" & 0 & " order by cliente,ano,mes"
Else
   data_deu.RecordSource = "Select * from deudas where fecha_pago is null and mes >" & 0 & " order by cliente,ano,mes"
End If
data_deu.Refresh

If data_deu.Recordset.RecordCount > 0 Then
   data_deu.Recordset.MoveFirst
   Do While Not data_deu.Recordset.EOF
      data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & data_deu.Recordset("cliente") & " and cod_prod =" & 999 & " and mes_paga =" & data_deu.Recordset("mes") & " and ano_paga =" & data_deu.Recordset("ano")
      data_lin.Refresh
      If data_lin.Recordset.RecordCount > 0 Then
         data_deu.Recordset.Edit
         data_deu.Recordset("fecha_pago") = data_lin.Recordset("fecha")
         data_deu.Recordset.Update
      End If
      data_deu.Recordset.MoveNext
   Loop
End If

If t_mat.Text <> "" Then
   data_deu.RecordSource = "Select * from deudas where cliente =" & t_mat.Text & " and fecha_pago is null order by cliente,ano,mes"
Else
   data_deu.RecordSource = "Select * from deudas where fecha_pago is null order by cliente,ano,mes"
End If
data_deu.Refresh
Xmat = data_deu.Recordset("cliente")
Xsaldo = 0
Xcant = 0
Xmes = 0
Xano = 0

If data_deu.Recordset.RecordCount > 0 Then
   data_deu.Recordset.MoveFirst
   Do While Not data_deu.Recordset.EOF
      If Xmat = data_deu.Recordset("cliente") Then
         Xsaldo = Xsaldo + data_deu.Recordset("total")
         If data_deu.Recordset("mes") > 0 Then
            Xcant = Xcant + 1
            If Xmes = 0 Then
               Xmes = data_deu.Recordset("mes")
               Xano = data_deu.Recordset("ano")
            Else
               If data_deu.Recordset("mes") > 9 Then
                  TextFec1 = "01/" & Trim(Str(data_deu.Recordset("mes"))) & "/" & Trim(Str(data_deu.Recordset("ano")))
               Else
                  TextFec1 = "01/" & "0" & Trim(Str(data_deu.Recordset("mes"))) & "/" & Trim(Str(data_deu.Recordset("ano")))
               End If
               Fec1 = CDate(TextFec1)
               If Xmes > 9 Then
                  TextFec2 = "01/" & Trim(Str(Xmes)) & "/" & Trim(Str(Xano))
               Else
                  TextFec2 = "01/0" & Trim(Str(Xmes)) & "/" & Trim(Str(Xano))
               End If
               Fec2 = CDate(TextFec2)
               If Format(Fec2, "yyyy/mm/dd") >= Format(Fec1, "yyyy/mm/dd") Then
               Else
                  Xmes = data_deu.Recordset("mes")
                  Xano = data_deu.Recordset("ano")
               End If
            End If
         End If
         Xmat = data_deu.Recordset("cliente")
         data_deu.Recordset.MoveNext
      Else
         data_cli.RecordSource = "Select * from clientes where cl_codigo =" & Xmat
         data_cli.Refresh
         If data_cli.Recordset.RecordCount > 0 Then
            If IsNull(data_cli.Recordset("saldo_cc")) = False Then
               If data_cli.Recordset("saldo_cc") = Xsaldo Then
               Else
                  data_cli.Recordset.Edit
                  data_cli.Recordset("saldo_cc") = Xsaldo
                  data_cli.Recordset.Update
               End If
            Else
               data_cli.Recordset.Edit
               data_cli.Recordset("saldo_cc") = Xsaldo
               data_cli.Recordset.Update
            End If
            If IsNull(data_cli.Recordset("cl_atrasoa")) = False Then
               If data_cli.Recordset("cl_atrasoa") = Xcant Then
               Else
                  data_cli.Recordset.Edit
                  data_cli.Recordset("cl_atrasoa") = Xcant
                  data_cli.Recordset.Update
               End If
            Else
               data_cli.Recordset.Edit
               data_cli.Recordset("cl_atrasoa") = Xcant
               data_cli.Recordset.Update
            End If
            If IsNull(data_cli.Recordset("cl_ultmesp")) = False Then
               If data_cli.Recordset("cl_ultmesp") = Xmes Then
               Else
                  data_cli.Recordset.Edit
                  data_cli.Recordset("cl_ultmesp") = Xmes
                  data_cli.Recordset.Update
               End If
            Else
               data_cli.Recordset.Edit
               data_cli.Recordset("cl_ultmesp") = Xmes
               data_cli.Recordset.Update
            End If
            If IsNull(data_cli.Recordset("cl_ultanop")) = False Then
               If data_cli.Recordset("cl_ultanop") = Xano Then
               Else
                  data_cli.Recordset.Edit
                  data_cli.Recordset("cl_ultanop") = Xano
                  data_cli.Recordset.Update
               End If
            Else
               data_cli.Recordset.Edit
               data_cli.Recordset("cl_ultanop") = Xano
               data_cli.Recordset.Update
            End If
         End If
         Xmat = data_deu.Recordset("cliente")
         Xsaldo = 0
         Xmes = 0
         Xano = 0
         Xcant = 0
         MsgBox "MAT:" & Xmat
      End If
   Loop
End If
      
frm_ctrolsald.MousePointer = 0
MsgBox "Proceso incialización terminado"
Unload Me
        

End Sub

Private Sub Form_Load()
'data_lin.DatabaseName = App.Path & "\sapp.mdb"
data_lin.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_deu.Connect = "odbc;dsn=" & Xconexrmt & ";"

tme.Text = Month(Date)
tan.Text = Year(Date)

End Sub

Private Sub t_mat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command1.SetFocus
End If

End Sub
