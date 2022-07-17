VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_inflab 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes para fertilab"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_inflab.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   5640
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport cr1 
      Left            =   2040
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_par 
      Caption         =   "data_par"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5040
      Picture         =   "frm_inflab.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Cerrar ésta ventana"
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      Picture         =   "frm_inflab.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Procesar Informe"
      Top             =   4080
      Width           =   375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      Caption         =   "Datos para el informe"
      ForeColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5175
      Begin MSAdodcLib.Adodc data_cli 
         Height          =   375
         Left            =   2760
         Top             =   1800
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "data_cli"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc data_lin 
         Height          =   375
         Left            =   2760
         Top             =   1320
         Visible         =   0   'False
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   661
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "data_lin"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0080FFFF&
         Caption         =   "Omitir chequeo de facturados"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   3360
         Width           =   4695
      End
      Begin VB.Data data_reserv 
         Caption         =   "data_reserv"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   960
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00808000&
         Caption         =   "Todos"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2760
         Width           =   2415
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00808000&
         Caption         =   "No Fertilab"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00808000&
         Caption         =   "Solo Fertilab"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Value           =   -1  'True
         Width           =   2415
      End
      Begin VB.TextBox t_b 
         BackColor       =   &H00808000&
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   1200
         Width           =   735
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   3480
         TabIndex        =   3
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   8421376
         ForeColor       =   16777215
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox md 
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   8421376
         ForeColor       =   16777215
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "(99 = Todas las bases)"
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2400
         TabIndex        =   6
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808080&
         Caption         =   "BASES:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808080&
         Caption         =   "FECHAS:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1575
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   960
      Picture         =   "frm_inflab.frx":109E
      Stretch         =   -1  'True
      Top             =   4080
      Width           =   2055
   End
End
Attribute VB_Name = "frm_inflab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Xlamatlab As Long
Dim Xfecdesdelab As Date
Xfecdesdelab = CDate(md.Text) - 25

Command1.Enabled = False
data_inf.RecordSource = "infcli"
data_inf.Refresh
If data_inf.Recordset.RecordCount > 0 Then
   If data_inf.Recordset.RecordCount > 0 Then
      data_inf.Recordset.MoveFirst
      Do While Not data_inf.Recordset.EOF
         data_inf.Recordset.Delete
         data_inf.Recordset.MoveNext
      Loop
   End If
End If

If mh.Text <> "__/__/____" And md.Text <> "__/__/____" And t_b.Text <> "" Then
   If t_b.Text = 99 Then
'      If Option1.value = True Then
'         data_lin.RecordSource = "Select * from linmmdd where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and nro_flia =" & 3 & " order by cod_cli,fecha"
'         data_lin.Refresh
'      Else
'         If Option2.value = True Then
'            data_lin.RecordSource = "Select * from linmmdd where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and nro_flia =" & 3 & " and tcambio <>" & 8 & " order by cod_cli,fecha"
'            data_lin.Refresh
'         Else
'            data_lin.RecordSource = "Select * from linmmdd where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and nro_flia =" & 3 & " order by cod_cli,fecha"
'            data_lin.Refresh
'         End If
'      End If
      data_reserv.RecordSource = "Select * from t_fechas where fecha >='" & md.Text & "' and fecha <='" & mh.Text & "' and especial ='" & "LABORATORIO" & "' and ced_pac is not null"
      data_reserv.Refresh
   Else
'      If Option1.value = True Then
'         data_lin.RecordSource = "Select * from linmmdd where vto >=#" & Format(md.Text, "yyyy/mm/dd") & "# and vto <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and nro_flia =" & 3 & " and tcambio =" & 8 & " and base =" & t_b.Text & " order by cod_cli,fecha"
'         data_lin.Refresh
'      Else
'         If Option2.value = True Then
'            data_lin.RecordSource = "Select * from linmmdd where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and nro_flia =" & 3 & " and tcambio <>" & 8 & " and base =" & t_b.Text & " order by cod_cli,fecha"
'            data_lin.Refresh
'         Else
'            data_lin.RecordSource = "Select * from linmmdd where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and nro_flia =" & 3 & " and base =" & t_b.Text & " order by cod_cli,fecha"
'            data_lin.Refresh
'         End If
'      End If
      data_reserv.RecordSource = "Select * from t_fechas where fecha >='" & md.Text & "' and fecha <='" & mh.Text & "' and especial ='" & "LABORATORIO" & "' and base =" & t_b.Text & " and ced_pac is not null"
      data_reserv.Refresh
   End If
   If data_reserv.Recordset.RecordCount > 0 Then
      data_reserv.Recordset.MoveFirst
      Dim Xlaceddelpac As Long
      Xlaceddelpac = 0
      Do While Not data_reserv.Recordset.EOF
         If Len(data_reserv.Recordset("ced_pac")) = 8 Then
            Xlaceddelpac = Val(Mid(Trim(data_reserv.Recordset("ced_pac")), 1, 7))
         Else
            If Len(data_reserv.Recordset("ced_pac")) = 7 Then
               Xlaceddelpac = Val(Mid(Trim(data_reserv.Recordset("ced_pac")), 1, 6))
            Else
               Xlaceddelpac = Val(Mid(Trim(data_reserv.Recordset("ced_pac")), 1, 5))
            End If
         End If
         data_cli.RecordSource = "Select * from clientes where cl_cedula =" & Xlaceddelpac
         data_cli.Refresh
         If data_cli.Recordset.RecordCount > 0 Then
            If Check1.Value = 1 Then
               data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & data_cli.Recordset("cl_codigo") & " and fecha >='" & Format(Xfecdesdelab, "yyyy-mm-dd") & "' and nro_flia =" & 3 & " and tcambio =" & 9
            Else
               data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & data_cli.Recordset("cl_codigo") & " and fecha >='" & Format(Xfecdesdelab, "yyyy-mm-dd") & "' and nro_flia =" & 3 & " and tcambio =" & 8
            End If
            data_lin.Refresh
            If data_lin.Recordset.RecordCount > 0 Then
               data_lin.Recordset.MoveFirst
               Do While Not data_lin.Recordset.EOF
                  data_inf.RecordSource = "Select * from infcli where cl_fnac =#" & Format(data_lin.Recordset("fecha"), "yyyy/mm/dd") & "# and cl_codigo =" & data_lin.Recordset("cod_cli")
                  data_inf.Refresh
                  If data_inf.Recordset.RecordCount > 0 Then
                     data_inf.Recordset.Edit
                     data_inf.Recordset("info_debit") = data_inf.Recordset("info_debit") & Chr(13) & data_lin.Recordset("nom_prod")
                     data_inf.Recordset.Update
                  Else
                     data_inf.Recordset.AddNew
                     data_inf.Recordset("cl_codigo") = data_lin.Recordset("cod_cli")
                     data_inf.Recordset("cl_apellid") = data_lin.Recordset("nom_cli")
                     data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                     data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                     data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fnac")
                     data_inf.Recordset("cl_nrocobr") = data_lin.Recordset("cod_prod")
                     data_inf.Recordset("cl_fnac") = data_lin.Recordset("fecha")
                     data_inf.Recordset("cl_nrovend") = data_lin.Recordset("base")
                     If IsNull(data_cli.Recordset("cl_referen")) = False Then
                        If Mid(Trim(data_cli.Recordset("cl_referen")), 1, 8) = "NOAPLICA" Or _
                           Mid(Trim(data_cli.Recordset("cl_referen")), 1, 8) = "NO APLIC" Then
                           data_inf.Recordset("cl_dircobr") = "Sin Datos"
                        Else
                           data_inf.Recordset("cl_dircobr") = data_cli.Recordset("cl_referen")
                        End If
                     Else
                        data_inf.Recordset("cl_dircobr") = "Sin datos"
                     End If
                     data_inf.Recordset("info_debit") = data_lin.Recordset("nom_prod")
                     data_inf.Recordset("cl_zona") = Trim(str(data_lin.Recordset("ced_socio"))) & "-" & Trim(str(data_lin.Recordset("fact")))
                     data_inf.Recordset("cl_codconv") = data_lin.Recordset("convenio")
                     data_inf.Recordset("cl_nombre") = WElusuario
                     data_inf.Recordset.Update
                  End If
                  If Check1.Value = 1 Then
                  Else
'                     data_lin.Recordset.Edit
                     data_lin.Recordset("tcambio") = 9
                     data_lin.Recordset.Update
                  End If
                  data_lin.Recordset.MoveNext
               Loop
            End If
         Else
            MsgBox "No se encontró la CEDULA: " & data_reserv.Recordset("ced_pac") & " VERIFIQUE!!", vbInformation
         End If
         data_reserv.Recordset.MoveNext
      Loop
      MsgBox "Proceso terminado", vbInformation
      cr1.ReportFileName = App.path & "\inflabos.rpt"
      cr1.ReportTitle = "Informe de LABORATORIOS desde: " & md.Text & " hasta:" & mh.Text & "BASE:" & t_b.Text
      cr1.Action = 1
   Else
      MsgBox "No existe reserva de laboratorio con éstas fechas"
   End If
Else
   MsgBox "Faltan datos para el informe, verifique", vbInformation
End If
Command1.Enabled = True

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
data_par.DatabaseName = App.path & "\parse.mdb"
data_par.RecordSource = "parsec0"
data_par.Refresh
t_b.Text = data_par.Recordset("base")
data_inf.DatabaseName = App.path & "\informes.mdb"

'data_lin.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_lin.ConnectionString = "dsn=" & Xconexrmt
'data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cli.ConnectionString = "dsn=" & Xconexrmt
data_reserv.Connect = "ODBC;DSN=" & Xconexrmt & ";"
md.Text = Format(Date, "dd/mm/yyyy")
mh.Text = Format(Date, "dd/mm/yyyy")


End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub md_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mh.SetFocus
End If

End Sub

Private Sub mh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_b.SetFocus
End If

End Sub
