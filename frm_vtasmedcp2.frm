VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_vtasmedcp2 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas por Médico"
   ClientHeight    =   5115
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5775
   Icon            =   "frm_vtasmedcp2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5115
   ScaleWidth      =   5775
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4560
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSComctlLib.ProgressBar barr 
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   3960
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin Crystal.CrystalReport crv 
      Left            =   3480
      Top             =   4200
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
      Left            =   1800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "infvtas"
      Top             =   3000
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton b_canc 
      BackColor       =   &H00FFFFFF&
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
      Left            =   4800
      MouseIcon       =   "frm_vtasmedcp2.frx":0442
      MousePointer    =   99  'Custom
      Picture         =   "frm_vtasmedcp2.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir"
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton b_acep 
      BackColor       =   &H00FFFFFF&
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
      Left            =   240
      MouseIcon       =   "frm_vtasmedcp2.frx":0CD6
      MousePointer    =   99  'Custom
      Picture         =   "frm_vtasmedcp2.frx":0FE0
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Aceptar"
      Top             =   4440
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Datos para informe"
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
      Height          =   3495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5175
      Begin MSAdodcLib.Adodc data_lin 
         Height          =   375
         Left            =   2520
         Top             =   1920
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H0080FFFF&
         Caption         =   "Informe desde historial"
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
         Left            =   240
         TabIndex        =   14
         Top             =   3000
         Width           =   4695
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Emitir informe por Médico que Ordena."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   2400
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Informe sin detalle"
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
         Left            =   2640
         TabIndex        =   11
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox tm 
         Height          =   285
         Left            =   480
         TabIndex        =   10
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Data data_med 
         Caption         =   "data_med"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2520
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   1320
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox txt_b 
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
         Height          =   405
         Left            =   1560
         MaxLength       =   3
         TabIndex        =   7
         Top             =   1680
         Width           =   735
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "frm_vtasmedcp2.frx":156A
         Height          =   360
         Left            =   1560
         TabIndex        =   5
         Top             =   960
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   12648384
         ForeColor       =   0
         ListField       =   "MED_NOMBRE"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   3600
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
         Left            =   2040
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
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "BASE: (99=TODAS)"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Médico:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Rango de fecha:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   1440
      Picture         =   "frm_vtasmedcp2.frx":1581
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1575
   End
End
Attribute VB_Name = "frm_vtasmedcp2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_acep_Click()
b_acep.Enabled = False
b_canc.Enabled = False
frm_vtasmed.MousePointer = 11
If DBCombo1.Text = "" Then
   tm.Text = 9999
End If
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infvtas"
data_inf.RecordSource = "infvtas"
data_inf.Refresh

If md.Text <> "__/__/____" Then
   If mh.Text <> "__/__/____" Then
      If tm.Text <> "" Then
         If txt_b.Text <> "" Then
            If txt_b.Text = 99 Then
               If tm.Text = 9999 Then
                  If Check2.Value = 1 Then
                     If Check3.Value = 1 Then
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And nro_med_s >=" & 1 & " And nro_med_s <=" & 900 & " and tipo <>'" & "NOTA CR" & "' and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6) order by nro_med_s"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And nro_med_s >=" & 1 & " And nro_med_s <=" & 2000 & " and pendiente not in ('N','C','R') and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6) order by nro_med_s"
                        data_lin.Refresh
                     End If
                  Else
                     If Check3.Value = 1 Then
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And nro_med_a >=" & 1 & " And nro_med_a <=" & 900 & " and tipo <>'" & "NOTA CR" & "' and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6) order by nro_med_a"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And nro_med_a >=" & 1 & " And nro_med_a <=" & 2000 & " and pendiente not in ('N','C','R') and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6) order by nro_med_a"
                        data_lin.Refresh
                     End If
                  End If
               Else
                  If Check2.Value = 1 Then
                     If Check3.Value = 1 Then
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_med_s =" & tm.Text & " and tipo <>'" & "NOTA CR" & "' and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6)"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_med_s =" & tm.Text & " and pendiente not in ('N','C','R') and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6)"
                        data_lin.Refresh
                     End If
                  Else
                     If Check3.Value = 1 Then
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_med_a =" & tm.Text & " and tipo <>'" & "NOTA CR" & "' and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6)"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_med_a =" & tm.Text & " and pendiente not in ('N','C','R') and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6)"
                        data_lin.Refresh
                     End If
                  End If
               End If
            Else
               If tm.Text = 9999 Then
                  If Check2.Value = 1 Then
                     If Check3.Value = 1 Then
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & txt_b.Text & " And nro_med_s >=" & 1 & " And nro_med_s <=" & 900 & " and tipo <>'" & "NOTA CR" & "' and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6) order by nro_med_a"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & txt_b.Text & " And nro_med_s >=" & 1 & " And nro_med_s <=" & 2000 & " and pendiente not in ('N','C','R') and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6) order by nro_med_a"
                        data_lin.Refresh
                     End If
                  Else
                     If Check3.Value = 1 Then
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & txt_b.Text & " And nro_med_a >=" & 1 & " And nro_med_a <=" & 900 & " and tipo <>'" & "NOTA CR" & "' and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6) order by nro_med_a"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & txt_b.Text & " And nro_med_a >=" & 1 & " And nro_med_a <=" & 2000 & " and pendiente not in ('N','C','R') and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6) order by nro_med_a"
                        data_lin.Refresh
                     End If
                  End If
               Else
                  If Check2.Value = 1 Then
                     If Check3.Value = 1 Then
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_med_s =" & tm.Text & " And base =" & txt_b.Text & " and tipo <>'" & "NOTA CR" & "' and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6)"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_med_s =" & tm.Text & " And base =" & txt_b.Text & " and pendiente not in ('N','C','R') and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6)"
                        data_lin.Refresh
                     End If
                  Else
                     If Check3.Value = 1 Then
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_med_a =" & tm.Text & " And base =" & txt_b.Text & " and tipo <>'" & "NOTA CR" & "' and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6)"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and nro_med_a =" & tm.Text & " And base =" & txt_b.Text & " and pendiente not in ('N','C','R') and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6)"
                        data_lin.Refresh
                     End If
                  End If
               End If
            End If
            If data_lin.Recordset.RecordCount > 0 Then
               data_lin.Recordset.MoveLast
               barr.Visible = True
               barr.Max = data_lin.Recordset.RecordCount
               barr.Value = 0
               data_lin.Recordset.MoveFirst
               DoEvents
               Do While Not data_lin.Recordset.EOF
                  data_inf.Recordset.AddNew
                  data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
                  data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                  data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                  data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
                  data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
                  data_inf.Recordset("nro_flia") = data_lin.Recordset("nro_flia")
                  data_inf.Recordset("nom_flia") = data_lin.Recordset("nom_flia")
                  data_inf.Recordset("convenio") = data_lin.Recordset("convenio")
                  data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                  data_inf.Recordset("hora") = data_lin.Recordset("hora")
                  data_inf.Recordset("nro_med_a") = data_lin.Recordset("nro_med_a")
                  data_inf.Recordset("nom_med_a") = data_lin.Recordset("nom_med_a")
                  data_inf.Recordset("mes_paga") = data_lin.Recordset("mes_paga")
                  data_inf.Recordset("ano_paga") = data_lin.Recordset("ano_paga")
                  data_inf.Recordset("base") = data_lin.Recordset("base")
                  data_inf.Recordset("ced_socio") = data_lin.Recordset("ced_socio")
                  data1.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("convenio") & "'"
                  data1.Refresh
                  If data1.Recordset.RecordCount > 0 Then
                     If IsNull(data1.Recordset("cnv_grupo")) = False Then
                        If data1.Recordset("cnv_grupo") = "UNIVERSAL" Or _
                           data1.Recordset("cnv_grupo") = "CCOU" Or _
                           data1.Recordset("cnv_grupo") = "SMI" Or _
                           data1.Recordset("cnv_grupo") = "H.EVANGELICO" Or _
                           data1.Recordset("cnv_grupo") = "CASA DE GALICIA" Or _
                           data1.Recordset("cnv_grupo") = "IMPASA" Then
                           data_inf.Recordset("usa_timbre") = "S"
                        Else
                           data_inf.Recordset("usa_timbre") = "N"
                        End If
                     Else
                        data_inf.Recordset("usa_timbre") = "N"
                     End If
                  Else
                     data_inf.Recordset("usa_timbre") = "N"
                  End If
                  
                  If IsNull(data_lin.Recordset("imp_iva")) = False Then
                     data_inf.Recordset("imp_iva") = Format(data_lin.Recordset("imp_iva"), "Standard")
                  Else
                     data_inf.Recordset("imp_iva") = 0
                  End If
                  If IsNull(data_lin.Recordset("pendiente")) = False Then
                    If data_lin.Recordset("pendiente") = "T" Then
                       data_inf.Recordset("tipo") = "e-Tck " & data_lin.Recordset("tipo")
                    Else
                       If data_lin.Recordset("pendiente") = "F" Then
                          data_inf.Recordset("tipo") = "e-Fct " & data_lin.Recordset("tipo")
                       Else
                          If data_lin.Recordset("pendiente") = "C" Then
                             data_inf.Recordset("tipo") = "NC e-Tck " & data_lin.Recordset("tipo")
                             data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin") * -1
                             data_inf.Recordset("imp_iva") = data_lin.Recordset("imp_iva") * -1
                             data_inf.Recordset("imp_iva") = Format(data_inf.Recordset("imp_iva"), "Standard")
                          Else
                             If data_lin.Recordset("pendiente") = "N" Then
                                data_inf.Recordset("tipo") = "NC e-Fct " & data_lin.Recordset("tipo")
                                data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin") * -1
                                data_inf.Recordset("imp_iva") = data_lin.Recordset("imp_iva") * -1
                                data_inf.Recordset("imp_iva") = Format(data_inf.Recordset("imp_iva"), "Standard")
                             Else
                                If data_lin.Recordset("pendiente") = "A" Then
                                   data_inf.Recordset("tipo") = "ND e-Fct " & data_lin.Recordset("tipo")
                                Else
                                   If data_lin.Recordset("pendiente") = "B" Then
                                      data_inf.Recordset("tipo") = "ND e-Tck " & data_lin.Recordset("tipo")
                                   Else
                                      If data_lin.Recordset("pendiente") = "R" Then
                                         data_inf.Recordset("tipo") = "Dev.Recibo"
                                         data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin") * -1
                                      Else
                                         If data_lin.Recordset("pendiente") = "Z" Then
                                            data_inf.Recordset("tipo") = "Recibo"
                                         Else
                                            data_inf.Recordset("tipo") = "Registro"
                                         End If
                                      End If
                                   End If
                                End If
                             End If
                          End If
                       End If
                    End If
                  Else
                    data_inf.Recordset("tipo") = data_lin.Recordset("tipo")
                  End If
                  
                  data_inf.Recordset.Update
                  data_lin.Recordset.MoveNext
                  barr.Value = barr.Value + 1
               Loop
               DoEvents
               barr.Visible = False
               MiBaseact.Execute "Delete * from infvtas where usa_timbre ='" & "S" & "'"
               data_inf.RecordSource = "Select * from infvtas order by nro_med_a"
               data_inf.Refresh
               If Check1.Value = 1 Then
                  crv.ReportFileName = App.path & "\infvtasxmedncp.rpt"
                  If txt_b.Text = 99 Then
                     If Check2.Value = 1 Then
                        crv.ReportTitle = "INFORME DE VENTAS MEDICO QUE ORDENA FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy") & " ==TODAS LAS BASES=="
                     Else
                        crv.ReportTitle = "INFORME DE VENTAS POR MEDICO FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy") & " ==TODAS LAS BASES=="
                     End If
                  Else
                     If Check2.Value = 1 Then
                        crv.ReportTitle = "INFORME DE VENTAS MEDICO QUE ORDENA FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy") & " ==BASE: " & txt_b.Text
                     Else
                        crv.ReportTitle = "INFORME DE VENTAS POR MEDICO FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy") & " ==BASE: " & txt_b.Text
                     End If
                  End If
                  crv.Action = 1
               Else
                  crv.ReportFileName = App.path & "\infvtasxmedcp.rpt"
                  If txt_b.Text = 99 Then
                     If Check2.Value = 1 Then
                        crv.ReportTitle = "INFORME DE VENTAS MEDICO QUE ORDENA FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy") & "  ==TODAS LAS BASES=="
                     Else
                        crv.ReportTitle = "INFORME DE VENTAS POR MEDICO FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy") & "  ==TODAS LAS BASES=="
                     End If
                  Else
                     If Check2.Value = 1 Then
                        crv.ReportTitle = "INFORME DE VENTAS MEDICO QUE ORDENA FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy") & " ==BASE: " & txt_b.Text
                     Else
                        crv.ReportTitle = "INFORME DE VENTAS POR MEDICO FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy") & " ==BASE: " & txt_b.Text
                     End If
                  End If
                  crv.Action = 1
               End If
            
            Else
               MsgBox "No existen registros con esta selección", vbInformation, "Mensaje"
            End If
         Else
            MsgBox "Ingrese Base", vbInformation, "Mensaje"
            txt_b.SetFocus
         End If
      Else
         MsgBox "Número de médico incorrecto", vbInformation, "Mensaje"
         DBCombo1.SetFocus
      End If
   Else
      MsgBox "Ingrese Fecha", vbInformation, "Mensaje"
      mh.SetFocus
   End If
Else
   MsgBox "Ingrese fecha", vbInformation, "Mensaje"
   md.SetFocus
End If
b_acep.Enabled = True
b_canc.Enabled = True
frm_vtasmed.MousePointer = 0

End Sub

Private Sub b_canc_Click()
Unload Me

End Sub

Private Sub DBCombo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_b.SetFocus
End If

End Sub

Private Sub DBCombo1_LostFocus()

If IsNumeric(DBCombo1.Text) = True Then
   data_med.Recordset.FindFirst "med_cod =" & DBCombo1.Text
   If Not data_med.Recordset.NoMatch Then
      tm.Text = data_med.Recordset("med_cod")
      DBCombo1.Text = data_med.Recordset("med_nombre")
   Else
      MsgBox "No encontrado, consulte por nombre", vbInformation, "Mensaje"
      DBCombo1.SetFocus
   End If
Else
   data_med.Recordset.FindFirst "med_nombre ='" & DBCombo1.Text & "'"
   If Not data_med.Recordset.NoMatch Then
      tm.Text = data_med.Recordset("med_cod")
      DBCombo1.Text = data_med.Recordset("med_nombre")
   Else
      MsgBox "No ha seleccionado médico, se emitirán TODOS", vbInformation, "Mensaje"
'      DBCombo1.SetFocus
      tm.Text = 9999
      txt_b.SetFocus
   End If
End If


End Sub

Private Sub Form_Load()
data_med.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_med.RecordSource = "Select * from medicos order by med_nombre"
data_med.Refresh
data_inf.DatabaseName = App.path & "\informes.mdb"
data_lin.ConnectionString = "dsn=" & Xconexrmt
data1.Connect = "odbc;dsn=" & Xconexrmt & ";"

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
   DBCombo1.SetFocus
End If

End Sub

Private Sub txt_b_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_acep.SetFocus
End If

End Sub
