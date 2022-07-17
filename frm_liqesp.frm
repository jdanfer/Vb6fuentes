VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_liqesp 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liquidación de especialistas"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6030
   Icon            =   "frm_liqesp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   6030
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data_lin 
      Height          =   330
      Left            =   2040
      Top             =   3840
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   582
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Data data_linfec 
      Caption         =   "data_linfec"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
      Visible         =   0   'False
      Width           =   3015
   End
   Begin Crystal.CrystalReport crl 
      Left            =   2520
      Top             =   3600
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
      Height          =   345
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton b_ca 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4200
      Picture         =   "frm_liqesp.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton b_ac 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      Picture         =   "frm_liqesp.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Datos para liquidación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.Data data_cabfec 
         Caption         =   "data_cabfec"
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
         Top             =   1680
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H000080FF&
         Caption         =   "Ver"
         Height          =   255
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Ver la lista de los médicos"
         Top             =   1200
         Width           =   735
      End
      Begin VB.Data data_med 
         Caption         =   "data_med"
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
         Top             =   840
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FF8080&
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
         Height          =   315
         ItemData        =   "frm_liqesp.frx":0F56
         Left            =   2640
         List            =   "frm_liqesp.frx":0F60
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2400
         Width           =   2535
      End
      Begin VB.TextBox t_b 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   7
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox t_codm 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         MaxLength       =   3
         TabIndex        =   5
         Top             =   1200
         Width           =   855
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox md 
         Height          =   375
         Left            =   2160
         TabIndex        =   2
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FF8080&
         Caption         =   "Opción de informe:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   2400
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF8080&
         Caption         =   "Base (99 = TODAS )"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
         Caption         =   "Cod.Medico(999=Todos)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
         Caption         =   "Rango de fechas:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   3000
      Picture         =   "frm_liqesp.frx":0F90
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1575
   End
End
Attribute VB_Name = "frm_liqesp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_ac_Click()
Dim Xcansoc As Long
Dim Xcodmed, Xtipomed, Xcantpac, Xcantpol, Xcodmedsapp, Xcantecg, Xcantecos, Xcantrefra As Integer
Dim Xnommed, Xespecial As String

frm_liqesp.MousePointer = 11

data_inf.RecordSource = "liqesp"
data_inf.Refresh
If data_inf.Recordset.RecordCount > 0 Then
   data_inf.Recordset.MoveFirst
   Do While Not data_inf.Recordset.EOF
      data_inf.Recordset.Delete
      data_inf.Recordset.MoveNext
   Loop
End If
'                  data_lin.RecordSource = "Select * from linmmdd where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cod_prod =" & 2 & " and nro_med_a =" & t_codm.Text
'                  data_lin.Refresh

If md.Text <> "__/__/____" Then
   If mh.Text <> "__/__/____" Then
      If t_codm.Text <> "" Then
         If t_b.Text <> "" Then
            If Combo1.ListIndex = 0 Then
                If t_b.Text = 99 Then
                   If t_codm.Text = 999 Then
                      data_linfec.RecordSource = "Select * from t_fechas where CDate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# And CDate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and especial not in ('LABORATORIO','PEDIATRIA','VACUNACION') and mat_pac is not null order by cod_med,cdate(fecha)"
                      data_linfec.Refresh
                   Else
                      data_linfec.RecordSource = "Select * from t_fechas where cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# And cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and especial not in ('LABORATORIO','PEDIATRIA','VACUNACION') and cod_med =" & t_codm.Text & " order by cod_med,cdate(fecha)"
                      data_linfec.Refresh
                   End If
                Else
                   If t_codm.Text = 999 Then
                      data_linfec.RecordSource = "Select * from t_fechas where cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# And cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and especial not in ('LABORATORIO','PEDIATRIA','VACUNACION') and base =" & t_b.Text & " order by cod_med,cdate(fecha)"
                      data_linfec.Refresh
                   Else
                      data_linfec.RecordSource = "Select * from t_fechas where cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# And cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and especial not in ('LABORATORIO','PEDIATRIA','VACUNACION') and cod_med =" & t_codm.Text & " and base =" & t_b.Text & " order by cod_med,cdate(fecha)"
                      data_linfec.Refresh
                   End If
                End If
                If data_linfec.Recordset.RecordCount > 0 Then
                   data_linfec.Recordset.MoveLast
                   data_linfec.Recordset.MoveFirst
                   Xcantpac = 0
                   Xcantpol = 0
                   Xcantecg = 0
                   Xcantecos = 0
                   Xcantrefra = 0
                   Xcodmed = data_linfec.Recordset("cod_med")
                   Xnommed = data_linfec.Recordset("nom_med")
                   data_med.RecordSource = "Select * from medicos_esp where id =" & Xcodmed
                   data_med.Refresh
                   If data_med.Recordset.RecordCount > 0 Then
                      Xcodmedsapp = data_med.Recordset("cod_sapp")
                      Xtipomed = data_med.Recordset("cod_liq")
                      Xespecial = data_med.Recordset("esp_med")
                   Else
                      Xcodmedsapp = 0
                      Xtipomed = 0
                      Xespecial = ""
                   End If
                   Do While Not data_linfec.Recordset.EOF
                      If Xcodmed = data_linfec.Recordset("cod_med") Then
                         If IsNull(data_linfec.Recordset("mat_pac")) = False Then
                            data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & data_linfec.Recordset("mat_pac") & " and fecha ='" & Format(data_linfec.Recordset("fecha"), "yyyy-mm-dd") & "' and nro_med_a =" & Xcodmedsapp & " and nro_flia =" & 10 & " and cod_prod in (13056)"
                            data_lin.Refresh
                            If data_lin.Recordset.RecordCount > 0 Then
                               Xcantpac = Xcantpac + 1
                            End If
                            data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & data_linfec.Recordset("mat_pac") & " and fecha ='" & Format(data_linfec.Recordset("fecha"), "yyyy-mm-dd") & "' and nro_med_a =" & Xcodmedsapp & " and cod_prod =" & 11001
                            data_lin.Refresh
                            If data_lin.Recordset.RecordCount > 0 Then
                               Xcantecg = Xcantecg + 1
                            End If
                            data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & data_linfec.Recordset("mat_pac") & " and fecha ='" & Format(data_linfec.Recordset("fecha"), "yyyy-mm-dd") & "' and nro_med_a =" & Xcodmedsapp & " and nro_flia =" & 5
                            data_lin.Refresh
                            If data_lin.Recordset.RecordCount > 0 Then
                               Xcantecos = Xcantecos + 1
                            End If
                            data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & data_linfec.Recordset("mat_pac") & " and fecha ='" & Format(data_linfec.Recordset("fecha"), "yyyy-mm-dd") & "' and nro_med_a =" & Xcodmedsapp & " and cod_prod =" & 80004
                            data_lin.Refresh
                            If data_lin.Recordset.RecordCount > 0 Then
                               Xcantrefra = Xcantrefra + 1
                            End If
                         End If
                         Xcodmed = data_linfec.Recordset("cod_med")
                         data_linfec.Recordset.MoveNext
                      Else
                         data_linfec.Recordset.MovePrevious
                         data_cabfec.RecordSource = "Select * from t_cabfechas where cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cod_med =" & Xcodmed
                         data_cabfec.Refresh
                         If data_cabfec.Recordset.RecordCount > 0 Then
                            data_cabfec.Recordset.MoveLast
                            data_cabfec.Recordset.MoveFirst
                            Xcantpol = data_cabfec.Recordset.RecordCount
                         End If
                         data_inf.Recordset.AddNew
                         data_inf.Recordset("fecha") = Date
                         data_inf.Recordset("mes") = Month(md.Text)
                         data_inf.Recordset("ano") = Year(md.Text)
                         data_inf.Recordset("codmed") = Xcodmed
                         data_inf.Recordset("nommed") = Xnommed
                         data_inf.Recordset("especial") = Xespecial
                         data_inf.Recordset("tipoesp") = Xtipomed
                         data_inf.Recordset("cantpac") = Xcantpac
                         data_inf.Recordset("cantpoli") = Xcantpol
                         data_inf.Recordset("cantecg") = Xcantecg
                         data_inf.Recordset("cantecos") = Xcantecos
                         data_inf.Recordset("cantrefra") = Xcantrefra
                         data_inf.Recordset.Update
                         data_linfec.Recordset.MoveNext
                         Xcantpac = 0
                         Xcantpol = 0
                         Xcantecg = 0
                         Xcantecos = 0
                         Xcantrefra = 0
                         Xcodmed = data_linfec.Recordset("cod_med")
                         Xnommed = data_linfec.Recordset("nom_med")
                         data_med.RecordSource = "Select * from medicos_esp where id =" & Xcodmed
                         data_med.Refresh
                         If data_med.Recordset.RecordCount > 0 Then
                            Xcodmedsapp = data_med.Recordset("cod_sapp")
                            Xtipomed = data_med.Recordset("cod_liq")
                            Xespecial = data_med.Recordset("esp_med")
                         Else
                            Xcodmedsapp = 0
                            Xtipomed = 0
                            Xespecial = ""
                         End If
                      End If
                   Loop
                   frm_liqesp.MousePointer = 0
                   MsgBox "Proceso terminado"
                   data_inf.Refresh
                   crl.ReportFileName = App.path & "\infliqesp.rpt"
                   crl.ReportTitle = "PLANILLA LIQ.ESPECIALISTAS -MES: " & Month(md.Text) & "/" & Year(md.Text)
                   crl.Action = 1
                Else
                   frm_liqesp.MousePointer = 0
                   MsgBox "No existen registros con esta selección", vbInformation, "Mensaje"
                End If
            End If
            If Combo1.ListIndex = 1 Then
               data_inf.RecordSource = "infpac"
               data_inf.Refresh
               If data_inf.Recordset.RecordCount > 0 Then
                  data_inf.Recordset.MoveFirst
                  Do While Not data_inf.Recordset.EOF
                     data_inf.Recordset.Delete
                     data_inf.Recordset.MoveNext
                  Loop
               End If
                If t_b.Text = 99 Then
                   If t_codm.Text = 999 Then
                      data_linfec.RecordSource = "Select * from t_fechas where cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# And cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and especial not in ('LABORATORIO','PEDIATRIA','VACUNACION') order by cod_med,cdate(fecha)"
                      data_linfec.Refresh
                   Else
                      data_linfec.RecordSource = "Select * from t_fechas where cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# And cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and especial not in ('LABORATORIO','PEDIATRIA','VACUNACION') and cod_med =" & t_codm.Text & " order by cod_med,cdate(fecha)"
                      data_linfec.Refresh
                   End If
                Else
                   If t_codm.Text = 999 Then
                      data_linfec.RecordSource = "Select * from t_fechas where cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# And cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and especial not in ('LABORATORIO','PEDIATRIA','VACUNACION') and base =" & t_b.Text & " order by cod_med,cdate(fecha)"
                      data_linfec.Refresh
                   Else
                      data_linfec.RecordSource = "Select * from t_fechas where cdate(fecha) >=#" & Format(md.Text, "yyyy/mm/dd") & "# And cdate(fecha) <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and especial not in ('LABORATORIO','PEDIATRIA','VACUNACION') and cod_med =" & t_codm.Text & " and base =" & t_b.Text & " order by cod_med,cdate(fecha)"
                      data_linfec.Refresh
                   End If
                End If
                If data_linfec.Recordset.RecordCount > 0 Then
                   data_linfec.Recordset.MoveLast
                   data_linfec.Recordset.MoveFirst
                   Do While Not data_linfec.Recordset.EOF
                      Xcodmed = data_linfec.Recordset("cod_med")
                      Xnommed = data_linfec.Recordset("nom_med")
                      data_med.RecordSource = "Select * from medicos_esp where id =" & Xcodmed
                      data_med.Refresh
                      If data_med.Recordset.RecordCount > 0 Then
                         Xcodmedsapp = data_med.Recordset("cod_sapp")
                      Else
                         Xcodmedsapp = 0
                      End If
                      If IsNull(data_linfec.Recordset("mat_pac")) = False Then
                         data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & data_linfec.Recordset("mat_pac") & " and fecha ='" & Format(data_linfec.Recordset("fecha"), "yyyy-mm-dd") & "' and nro_med_a =" & Xcodmedsapp & " and cod_prod =" & 2
                         data_lin.Refresh
                         If data_lin.Recordset.RecordCount > 0 Then
                            data_inf.Recordset.AddNew
                            data_inf.Recordset("fecha") = CDate(data_linfec.Recordset("fecha"))
                            data_inf.Recordset("base") = data_linfec.Recordset("base")
                            data_inf.Recordset("codmed") = Xcodmedsapp
                            data_inf.Recordset("nommed") = data_linfec.Recordset("nom_med")
                            data_inf.Recordset("matpac") = data_linfec.Recordset("mat_pac")
                            data_inf.Recordset("nompac") = data_linfec.Recordset("nom_pac")
                            data_inf.Recordset("catpac") = data_linfec.Recordset("convenio")
                            data_inf.Recordset.Update
                            data_inf.Refresh
                         End If
                      End If
                      data_linfec.Recordset.MoveNext
                   Loop
                   frm_liqesp.MousePointer = 0
                   data_inf.Refresh
                   crl.ReportFileName = App.path & "\infliqpac.rpt"
                   crl.ReportTitle = "CONSULTAS DETALLADAS DESDE: " & md.Text & " HASTA: " & mh.Text
                   crl.Action = 1
                   
                   MsgBox "Proceso terminado"
                   
                End If
            End If
         Else
            frm_liqesp.MousePointer = 0
            MsgBox "Ingrese Base", vbInformation, "Mensaje"
            txt_b.SetFocus
         End If
      Else
         frm_liqesp.MousePointer = 0
         MsgBox "Número de médico incorrecto", vbInformation, "Mensaje"
         DBCombo1.SetFocus
      End If
   Else
      frm_liqesp.MousePointer = 0
      MsgBox "Ingrese Fecha", vbInformation, "Mensaje"
      mh.SetFocus
   End If
Else
   frm_liqesp.MousePointer = 0
   MsgBox "Ingrese fecha", vbInformation, "Mensaje"
   md.SetFocus
End If

End Sub

Private Sub b_ca_Click()
Unload Me

End Sub

Private Sub Command1_Click()
frm_especialistas.Show vbModal

End Sub

Private Sub Form_Load()
'data_lin.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_lin.ConnectionString = "dsn=" & Xconexrmt
data_inf.DatabaseName = App.path & "\liqesp.mdb"

data_linfec.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_med.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cabfec.Connect = "odbc;dsn=" & Xconexrmt & ";"

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
   t_codm.SetFocus
End If

End Sub

Private Sub t_b_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_ac.SetFocus
End If

End Sub

Private Sub t_codm_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_b.SetFocus
End If

End Sub
