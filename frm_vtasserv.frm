VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_vtasserv 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas por Servicio"
   ClientHeight    =   7020
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7815
   Icon            =   "frm_vtasserv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7020
   ScaleWidth      =   7815
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_facafil 
      Caption         =   "data_facafil"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSComctlLib.ProgressBar barra 
      Height          =   375
      Left            =   720
      TabIndex        =   12
      Top             =   5880
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin Crystal.CrystalReport crs2 
      Left            =   6600
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport crs 
      Left            =   6000
      Top             =   5400
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
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "infvtas"
      Top             =   1320
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
      Left            =   6720
      MouseIcon       =   "frm_vtasserv.frx":0442
      MousePointer    =   99  'Custom
      Picture         =   "frm_vtasserv.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir"
      Top             =   6360
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
      Left            =   360
      MouseIcon       =   "frm_vtasserv.frx":0CD6
      MousePointer    =   99  'Custom
      Picture         =   "frm_vtasserv.frx":0FE0
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Procesar"
      Top             =   6360
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
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
      Height          =   5535
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7335
      Begin VB.Data data_emi 
         Caption         =   "data_emi"
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
         Top             =   1440
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Data data_lin 
         Caption         =   "data_lin"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   3480
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2400
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Data Data2 
         Caption         =   "Data2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   4800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Plani covid"
         Height          =   375
         Left            =   3120
         TabIndex        =   27
         Top             =   2760
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "planilla"
         Height          =   495
         Left            =   2520
         TabIndex        =   26
         Top             =   1440
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox Check11 
         BackColor       =   &H00800000&
         Caption         =   "Generar planilla (Promotor)"
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
         Left            =   2400
         TabIndex        =   25
         Top             =   5040
         Width           =   2895
      End
      Begin VB.TextBox t_rub 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1560
         TabIndex        =   24
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CheckBox Check10 
         BackColor       =   &H00000000&
         Caption         =   "Solo mutuales"
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
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   4560
         Width           =   3255
      End
      Begin VB.CheckBox Check9 
         BackColor       =   &H00000000&
         Caption         =   "Sin mutuales"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3720
         TabIndex        =   21
         Top             =   4560
         Width           =   3255
      End
      Begin VB.CheckBox chtim 
         BackColor       =   &H00000000&
         Caption         =   "Solo timbres cobranza"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   4080
         Width           =   3255
      End
      Begin VB.CheckBox Check8 
         BackColor       =   &H0080FFFF&
         Caption         =   "Informe desde historial"
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
         Left            =   3840
         TabIndex        =   19
         Top             =   2160
         Width           =   3135
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00000000&
         Caption         =   "Emitir solo notas de crédito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3720
         TabIndex        =   18
         Top             =   4080
         Width           =   3255
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00000000&
         Caption         =   "Socios en convenios diferentes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   3600
         Width           =   3255
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00000000&
         Caption         =   "Socios Bajas con servicios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3720
         TabIndex        =   16
         Top             =   3600
         Width           =   3255
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00000000&
         Caption         =   "Incluir datos de cobro."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   3120
         Width           =   3255
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00000000&
         Caption         =   "Emitir SOLO facturas a CREDITO."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3720
         TabIndex        =   14
         Top             =   3120
         Width           =   3255
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00000000&
         Caption         =   "Emitir SOLO facturas manuales."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3720
         TabIndex        =   13
         Top             =   2640
         Width           =   3255
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
         Left            =   3840
         TabIndex        =   11
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox tm 
         Height          =   285
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Data data_med 
         Caption         =   "data_med"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   4320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "estudios"
         Top             =   1200
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
         Bindings        =   "frm_vtasserv.frx":156A
         Height          =   360
         Left            =   1560
         TabIndex        =   5
         Top             =   960
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   12648384
         ForeColor       =   0
         ListField       =   "DESCRIP"
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
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Nro.Rubro:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   23
         Top             =   2400
         Width           =   1335
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
         Caption         =   "Servicio:"
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
      Height          =   1095
      Left            =   2880
      Picture         =   "frm_vtasserv.frx":1581
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   1815
   End
End
Attribute VB_Name = "frm_vtasserv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub b_acep_Click()
b_acep.Enabled = False
b_canc.Enabled = False
If DBCombo1.Text = "" Then
   tm.Text = 999999
End If
Dim Xemitim As String
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
               If tm.Text = 999999 Then
                  If Check2.Value = 1 Then
                     If Check8.Value = 1 Then
                        If t_rub.Text <> "" Then
                           data_lin.RecordSource = "Select * from resplin where fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# And tot_lin >=" & 1 & " and tipo <>'" & "NOTA CR" & "' and cod_prod =" & 10001 & " and rub_cont =" & t_rub.Text
                        Else
                           data_lin.RecordSource = "Select * from resplin where fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# And tot_lin >=" & 1 & " and tipo <>'" & "NOTA CR" & "' and cod_prod =" & 10001
                        End If
                        data_lin.Refresh
                     Else
                        If t_rub.Text <> "" Then
                           data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.factura,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.cod_prod," & _
                           "linmmdd.nom_prod,linmmdd.nro_flia,linmmdd.nom_flia,linmmdd.convenio,linmmdd.tipo_mov," & _
                           "linmmdd.unidad,linmmdd.zona,linmmdd.servicio,linmmdd.tot_lin,linmmdd.nro_med_a,linmmdd.nom_med_a," & _
                           "linmmdd.mes_paga,linmmdd.ano_paga,linmmdd.costo_prod,linmmdd.base,linmmdd.ced_socio," & _
                           "linmmdd.grupo,linmmdd.tipo,linmmdd.hora,linmmdd.imp_iva,linmmdd.pendiente,linmmdd.rub_cont," & _
                           "linmmdd.fact,linmmdd.rub_nomb,linmmdd.reg_cab,linmmdd.nro_superv,linmmdd.nom_medic,linmmdd.ruc," & _
                           "linmmdd.nom_med_s,linmmdd.nom_superv,clientes.cl_codigo,clientes.estado,clientes.cl_telefon," & _
                           "clientes.cl_dpto,clientes.cl_codconv from linmmdd inner join clientes on " & _
                           "linmmdd.cod_cli = clientes.cl_codigo where linmmdd.fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And linmmdd.fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# And linmmdd.tot_lin >=" & 1 & " and linmmdd.cod_prod =" & 10001 & " and linmmdd.rub_cont =" & t_rub.Text
                        End If
                        data_lin.Refresh
                     End If
                  Else
                     If Check3.Value = 1 Then
                        If Check8.Value = 1 Then
                           data_lin.RecordSource = "Select * from resplin where fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and tipo ='" & "CREDITO" & "' and cod_prod =" & 10001
                           data_lin.Refresh
                        Else
                           data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.factura,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.cod_prod," & _
                           "linmmdd.nom_prod,linmmdd.nro_flia,linmmdd.nom_flia,linmmdd.convenio,linmmdd.tipo_mov," & _
                           "linmmdd.unidad,linmmdd.zona,linmmdd.servicio,linmmdd.tot_lin,linmmdd.nro_med_a,linmmdd.nom_med_a," & _
                           "linmmdd.mes_paga,linmmdd.ano_paga,linmmdd.costo_prod,linmmdd.base,linmmdd.ced_socio," & _
                           "linmmdd.grupo,linmmdd.tipo,linmmdd.hora,linmmdd.imp_iva,linmmdd.pendiente,linmmdd.rub_cont," & _
                           "linmmdd.fact,linmmdd.rub_nomb,linmmdd.reg_cab,linmmdd.nro_superv,linmmdd.nom_medic,linmmdd.ruc," & _
                           "linmmdd.nom_med_s,linmmdd.nom_superv,clientes.cl_codigo,clientes.estado,clientes.cl_telefon," & _
                           "clientes.cl_dpto,clientes.cl_codconv from linmmdd inner join clientes on " & _
                           "linmmdd.cod_cli = clientes.cl_codigo where linmmdd.fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And linmmdd.fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and linmmdd.tipo ='" & "CREDITO" & "' and linmmdd.cod_prod =" & 10001
                           data_lin.Refresh
                        End If
                     Else
                        If Check8.Value = 1 Then
                           If t_rub.Text <> "" Then
                              data_lin.RecordSource = "Select * from resplin where fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and tipo <>'" & "NOTA CR" & "' and cod_prod =" & 10001 & " and rub_cont =" & t_rub.Text
                           Else
                              data_lin.RecordSource = "Select * from resplin where fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and tipo <>'" & "NOTA CR" & "' and cod_prod =" & 10001
                           End If
                           data_lin.Refresh
                        Else
                           If t_rub.Text <> "" Then
                              data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.factura,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.cod_prod," & _
                              "linmmdd.nom_prod,linmmdd.nro_flia,linmmdd.nom_flia,linmmdd.convenio,linmmdd.tipo_mov," & _
                              "linmmdd.unidad,linmmdd.zona,linmmdd.servicio,linmmdd.tot_lin,linmmdd.nro_med_a,linmmdd.nom_med_a," & _
                              "linmmdd.mes_paga,linmmdd.ano_paga,linmmdd.costo_prod,linmmdd.base,linmmdd.ced_socio," & _
                              "linmmdd.grupo,linmmdd.tipo,linmmdd.hora,linmmdd.imp_iva,linmmdd.pendiente,linmmdd.rub_cont," & _
                              "linmmdd.fact,linmmdd.rub_nomb,linmmdd.reg_cab,linmmdd.nro_superv,linmmdd.nom_medic,linmmdd.ruc," & _
                              "linmmdd.nom_med_s,linmmdd.nom_superv,clientes.cl_codigo,clientes.estado,clientes.cl_telefon," & _
                              "clientes.cl_dpto,clientes.cl_codconv from linmmdd inner join clientes on " & _
                              "linmmdd.cod_cli = clientes.cl_codigo where linmmdd.fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And linmmdd.fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and linmmdd.rub_cont =" & t_rub.Text
                           Else
                              If Check5.Value = 1 Then
                                data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.factura,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.cod_prod," & _
                                "linmmdd.nom_prod,linmmdd.nro_flia,linmmdd.nom_flia,linmmdd.convenio,linmmdd.tipo_mov," & _
                                "linmmdd.unidad,linmmdd.zona,linmmdd.servicio,linmmdd.tot_lin,linmmdd.nro_med_a,linmmdd.nom_med_a," & _
                                "linmmdd.mes_paga,linmmdd.ano_paga,linmmdd.costo_prod,linmmdd.base,linmmdd.ced_socio," & _
                                "linmmdd.grupo,linmmdd.tipo,linmmdd.hora,linmmdd.imp_iva,linmmdd.pendiente,linmmdd.rub_cont," & _
                                "linmmdd.fact,linmmdd.rub_nomb,linmmdd.reg_cab,linmmdd.nro_superv,linmmdd.nom_medic,linmmdd.ruc," & _
                                "linmmdd.nom_med_s,linmmdd.nom_superv,clientes.cl_codigo,clientes.estado,clientes.cl_telefon," & _
                                "clientes.cl_dpto,clientes.cl_codconv from linmmdd inner join clientes on " & _
                                "linmmdd.cod_cli = clientes.cl_codigo where linmmdd.fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And linmmdd.fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and linmmdd.cod_prod not in (8000,803,804,805,806,802,999,998,997,993,994,992,10200) and linmmdd.base <=" & 99 & " and linmmdd.pendiente not in ('N','R','C') and clientes.estado in (2)"
                              Else
                                data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.factura,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.cod_prod," & _
                                "linmmdd.nom_prod,linmmdd.nro_flia,linmmdd.nom_flia,linmmdd.convenio,linmmdd.tipo_mov," & _
                                "linmmdd.unidad,linmmdd.zona,linmmdd.servicio,linmmdd.tot_lin,linmmdd.nro_med_a,linmmdd.nom_med_a," & _
                                "linmmdd.mes_paga,linmmdd.ano_paga,linmmdd.costo_prod,linmmdd.base,linmmdd.ced_socio," & _
                                "linmmdd.grupo,linmmdd.tipo,linmmdd.hora,linmmdd.imp_iva,linmmdd.pendiente,linmmdd.rub_cont," & _
                                "linmmdd.fact,linmmdd.rub_nomb,linmmdd.reg_cab,linmmdd.nro_superv,linmmdd.nom_medic,linmmdd.ruc," & _
                                "linmmdd.nom_med_s,linmmdd.nom_superv,clientes.cl_codigo,clientes.estado,clientes.cl_telefon," & _
                                "clientes.cl_dpto,clientes.cl_codconv from linmmdd inner join clientes on " & _
                                "linmmdd.cod_cli = clientes.cl_codigo where linmmdd.fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And linmmdd.fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and linmmdd.cod_prod not in (8000,803,804,805,806,802,999,998,997,993,994,992,10200) and linmmdd.base <=" & 99 & " and linmmdd.pendiente not in ('N','R','C')"
                              End If
                           End If
                           data_lin.Refresh
                        End If
                     End If
                  End If
               Else
                  If Check2.Value = 1 Then
                     If Check8.Value = 1 Then
                        data_lin.RecordSource = "Select * from resplin where fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and cod_prod =" & tm.Text & " And tot_lin >=" & 1 & " and tipo <>'" & "NOTA CR" & "'"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.factura,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.cod_prod," & _
                        "linmmdd.nom_prod,linmmdd.nro_flia,linmmdd.nom_flia,linmmdd.convenio,linmmdd.tipo_mov," & _
                        "linmmdd.unidad,linmmdd.zona,linmmdd.servicio,linmmdd.tot_lin,linmmdd.nro_med_a,linmmdd.nom_med_a," & _
                        "linmmdd.mes_paga,linmmdd.ano_paga,linmmdd.costo_prod,linmmdd.base,linmmdd.ced_socio," & _
                        "linmmdd.grupo,linmmdd.tipo,linmmdd.hora,linmmdd.imp_iva,linmmdd.pendiente,linmmdd.rub_cont," & _
                        "linmmdd.fact,linmmdd.rub_nomb,linmmdd.reg_cab,linmmdd.nro_superv,linmmdd.nom_medic,linmmdd.ruc," & _
                        "linmmdd.nom_med_s,linmmdd.nom_superv,clientes.cl_codigo,clientes.estado,clientes.cl_telefon," & _
                        "clientes.cl_dpto,clientes.cl_codconv from linmmdd inner join clientes on " & _
                        "linmmdd.cod_cli = clientes.cl_codigo where linmmdd.fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And linmmdd.fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and linmmdd.cod_prod =" & tm.Text & " And linmmdd.tot_lin >=" & 1
                        data_lin.Refresh
                     End If
                  Else
                     If Check3.Value = 1 Then
                        If Check8.Value = 1 Then
                           data_lin.RecordSource = "Select * from resplin where fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and cod_prod =" & tm.Text & " And tipo ='" & "CREDITO" & "'"
                           data_lin.Refresh
                        Else
                           data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.factura,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.cod_prod," & _
                           "linmmdd.nom_prod,linmmdd.nro_flia,linmmdd.nom_flia,linmmdd.convenio,linmmdd.tipo_mov," & _
                           "linmmdd.unidad,linmmdd.zona,linmmdd.servicio,linmmdd.tot_lin,linmmdd.nro_med_a,linmmdd.nom_med_a," & _
                           "linmmdd.mes_paga,linmmdd.ano_paga,linmmdd.costo_prod,linmmdd.base,linmmdd.ced_socio," & _
                           "linmmdd.grupo,linmmdd.tipo,linmmdd.hora,linmmdd.imp_iva,linmmdd.pendiente,linmmdd.rub_cont," & _
                           "linmmdd.fact,linmmdd.rub_nomb,linmmdd.reg_cab,linmmdd.nro_superv,linmmdd.nom_medic,linmmdd.ruc," & _
                           "linmmdd.nom_med_s,linmmdd.nom_superv,clientes.cl_codigo,clientes.estado,clientes.cl_telefon," & _
                           "clientes.cl_dpto,clientes.cl_codconv from linmmdd inner join clientes on " & _
                           "linmmdd.cod_cli = clientes.cl_codigo where linmmdd.fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And linmmdd.fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and linmmdd.cod_prod =" & tm.Text & " And linmmdd.tipo ='" & "CREDITO" & "'"
                           data_lin.Refresh
                        End If
                     Else
                        If Check8.Value = 1 Then
                           data_lin.RecordSource = "Select * from resplin where fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and cod_prod =" & tm.Text & " and tipo <>'" & "NOTA CR" & "'"
                           data_lin.Refresh
                        Else
                           If Check5.Value = 1 Then
                              data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.factura,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.cod_prod," & _
                              "linmmdd.nom_prod,linmmdd.nro_flia,linmmdd.nom_flia,linmmdd.convenio,linmmdd.tipo_mov," & _
                              "linmmdd.unidad,linmmdd.zona,linmmdd.servicio,linmmdd.tot_lin,linmmdd.nro_med_a,linmmdd.nom_med_a," & _
                              "linmmdd.mes_paga,linmmdd.ano_paga,linmmdd.costo_prod,linmmdd.base,linmmdd.ced_socio," & _
                              "linmmdd.grupo,linmmdd.tipo,linmmdd.hora,linmmdd.imp_iva,linmmdd.pendiente,linmmdd.rub_cont," & _
                              "linmmdd.fact,linmmdd.rub_nomb,linmmdd.reg_cab,linmmdd.nro_superv,linmmdd.nom_medic,linmmdd.ruc," & _
                              "linmmdd.nom_med_s,linmmdd.nom_superv,clientes.cl_codigo,clientes.estado,clientes.cl_telefon," & _
                              "clientes.cl_dpto,clientes.cl_codconv from linmmdd inner join clientes on " & _
                              "linmmdd.cod_cli = clientes.cl_codigo where linmmdd.fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And linmmdd.fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and linmmdd.cod_prod =" & tm.Text & " and clientes.estado in (2)"
                           Else
                              data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.factura,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.cod_prod," & _
                              "linmmdd.nom_prod,linmmdd.nro_flia,linmmdd.nom_flia,linmmdd.convenio,linmmdd.tipo_mov," & _
                              "linmmdd.unidad,linmmdd.zona,linmmdd.servicio,linmmdd.tot_lin,linmmdd.nro_med_a,linmmdd.nom_med_a," & _
                              "linmmdd.mes_paga,linmmdd.ano_paga,linmmdd.costo_prod,linmmdd.base,linmmdd.ced_socio," & _
                              "linmmdd.grupo,linmmdd.tipo,linmmdd.hora,linmmdd.imp_iva,linmmdd.pendiente,linmmdd.rub_cont," & _
                              "linmmdd.fact,linmmdd.rub_nomb,linmmdd.reg_cab,linmmdd.nro_superv,linmmdd.nom_medic,linmmdd.ruc," & _
                              "linmmdd.nom_med_s,linmmdd.nom_superv,clientes.cl_codigo,clientes.estado,clientes.cl_telefon," & _
                              "clientes.cl_dpto,clientes.cl_codconv from linmmdd inner join clientes on " & _
                              "linmmdd.cod_cli = clientes.cl_codigo where linmmdd.fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And linmmdd.fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and linmmdd.cod_prod =" & tm.Text
                           End If
                           data_lin.Refresh
                        End If
                     End If
                  End If
               End If
            Else
               If tm.Text = 999999 Then
                  If Check2.Value = 1 Then
                     If Check8.Value = 1 Then
                        data_lin.RecordSource = "Select * from resplin where fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# And base =" & txt_b.Text & " And tot_lin >=" & 1 & " and tipo <>'" & "NOTA CR" & "'"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.factura,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.cod_prod," & _
                        "linmmdd.nom_prod,linmmdd.nro_flia,linmmdd.nom_flia,linmmdd.convenio,linmmdd.tipo_mov," & _
                        "linmmdd.unidad,linmmdd.zona,linmmdd.servicio,linmmdd.tot_lin,linmmdd.nro_med_a,linmmdd.nom_med_a," & _
                        "linmmdd.mes_paga,linmmdd.ano_paga,linmmdd.costo_prod,linmmdd.base,linmmdd.ced_socio," & _
                        "linmmdd.grupo,linmmdd.tipo,linmmdd.hora,linmmdd.imp_iva,linmmdd.pendiente,linmmdd.rub_cont," & _
                        "linmmdd.fact,linmmdd.rub_nomb,linmmdd.reg_cab,linmmdd.nro_superv,linmmdd.nom_medic,linmmdd.ruc," & _
                        "linmmdd.nom_med_s,linmmdd.nom_superv,clientes.cl_codigo,clientes.estado,clientes.cl_telefon," & _
                        "clientes.cl_dpto,clientes.cl_codconv from linmmdd inner join clientes on " & _
                        "linmmdd.cod_cli = clientes.cl_codigo where linmmdd.fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And linmmdd.fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# And linmmdd.base =" & txt_b.Text & " And linmmdd.tot_lin >=" & 1
                        data_lin.Refresh
                     End If
                  Else
                     If Check3.Value = 1 Then
                        If Check8.Value = 1 Then
                           data_lin.RecordSource = "Select * from resplin where fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# And base =" & txt_b.Text & " and tipo ='" & "CREDITO" & "'"
                           data_lin.Refresh
                        Else
                           data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.factura,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.cod_prod," & _
                           "linmmdd.nom_prod,linmmdd.nro_flia,linmmdd.nom_flia,linmmdd.convenio,linmmdd.tipo_mov," & _
                           "linmmdd.unidad,linmmdd.zona,linmmdd.servicio,linmmdd.tot_lin,linmmdd.nro_med_a,linmmdd.nom_med_a," & _
                           "linmmdd.mes_paga,linmmdd.ano_paga,linmmdd.costo_prod,linmmdd.base,linmmdd.ced_socio," & _
                           "linmmdd.grupo,linmmdd.tipo,linmmdd.hora,linmmdd.imp_iva,linmmdd.pendiente,linmmdd.rub_cont," & _
                           "linmmdd.fact,linmmdd.rub_nomb,linmmdd.reg_cab,linmmdd.nro_superv,linmmdd.nom_medic,linmmdd.ruc," & _
                           "linmmdd.nom_med_s,linmmdd.nom_superv,clientes.cl_codigo,clientes.estado,clientes.cl_telefon," & _
                           "clientes.cl_dpto,clientes.cl_codconv from linmmdd inner join clientes on " & _
                           "linmmdd.cod_cli = clientes.cl_codigo where linmmdd.fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And linmmdd.fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# And linmmdd.base =" & txt_b.Text & " and linmmdd.tipo ='" & "CREDITO" & "'"
                           data_lin.Refresh
                        End If
                     Else
                        If Check8.Value = 1 Then
                           data_lin.RecordSource = "Select * from resplin where fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# And base =" & txt_b.Text & " and tipo <>'" & "NOTA CR" & "'"
                           data_lin.Refresh
                        Else
                           If Check5.Value = 1 Then
                              data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.factura,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.cod_prod," & _
                              "linmmdd.nom_prod,linmmdd.nro_flia,linmmdd.nom_flia,linmmdd.convenio,linmmdd.tipo_mov," & _
                              "linmmdd.unidad,linmmdd.zona,linmmdd.servicio,linmmdd.tot_lin,linmmdd.nro_med_a,linmmdd.nom_med_a," & _
                              "linmmdd.mes_paga,linmmdd.ano_paga,linmmdd.costo_prod,linmmdd.base,linmmdd.ced_socio," & _
                              "linmmdd.grupo,linmmdd.tipo,linmmdd.hora,linmmdd.imp_iva,linmmdd.pendiente,linmmdd.rub_cont," & _
                              "linmmdd.fact,linmmdd.rub_nomb,linmmdd.reg_cab,linmmdd.nro_superv,linmmdd.nom_medic,linmmdd.ruc," & _
                              "linmmdd.nom_med_s,linmmdd.nom_superv,clientes.cl_codigo,clientes.estado,clientes.cl_telefon," & _
                              "clientes.cl_dpto,clientes.cl_codconv from linmmdd inner join clientes on " & _
                              "linmmdd.cod_cli = clientes.cl_codigo where linmmdd.fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And linmmdd.fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# And linmmdd.base =" & txt_b.Text & " and clientes.estado in (2)"
                           Else
                              data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.factura,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.cod_prod," & _
                              "linmmdd.nom_prod,linmmdd.nro_flia,linmmdd.nom_flia,linmmdd.convenio,linmmdd.tipo_mov," & _
                              "linmmdd.unidad,linmmdd.zona,linmmdd.servicio,linmmdd.tot_lin,linmmdd.nro_med_a,linmmdd.nom_med_a," & _
                              "linmmdd.mes_paga,linmmdd.ano_paga,linmmdd.costo_prod,linmmdd.base,linmmdd.ced_socio," & _
                              "linmmdd.grupo,linmmdd.tipo,linmmdd.hora,linmmdd.imp_iva,linmmdd.pendiente,linmmdd.rub_cont," & _
                              "linmmdd.fact,linmmdd.rub_nomb,linmmdd.reg_cab,linmmdd.nro_superv,linmmdd.nom_medic,linmmdd.ruc," & _
                              "linmmdd.nom_med_s,linmmdd.nom_superv,clientes.cl_codigo,clientes.estado,clientes.cl_telefon," & _
                              "clientes.cl_dpto,clientes.cl_codconv from linmmdd inner join clientes on " & _
                              "linmmdd.cod_cli = clientes.cl_codigo where linmmdd.fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And linmmdd.fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# And linmmdd.base =" & txt_b.Text
                           End If
                           data_lin.Refresh
                        End If
                     End If
                  End If
               Else
                  If Check2.Value = 1 Then
                     If Check8.Value = 1 Then
                        data_lin.RecordSource = "Select * from resplin where fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and cod_prod =" & tm.Text & " And base =" & txt_b.Text & " And tot_lin >=" & 1 & " and tipo <>'" & "NOTA CR" & "'"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.factura,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.cod_prod," & _
                        "linmmdd.nom_prod,linmmdd.nro_flia,linmmdd.nom_flia,linmmdd.convenio,linmmdd.tipo_mov," & _
                        "linmmdd.unidad,linmmdd.zona,linmmdd.servicio,linmmdd.tot_lin,linmmdd.nro_med_a,linmmdd.nom_med_a," & _
                        "linmmdd.mes_paga,linmmdd.ano_paga,linmmdd.costo_prod,linmmdd.base,linmmdd.ced_socio," & _
                        "linmmdd.grupo,linmmdd.tipo,linmmdd.hora,linmmdd.imp_iva,linmmdd.pendiente,linmmdd.rub_cont," & _
                        "linmmdd.fact,linmmdd.rub_nomb,linmmdd.reg_cab,linmmdd.nro_superv,linmmdd.nom_medic,linmmdd.ruc," & _
                        "linmmdd.nom_med_s,linmmdd.nom_superv,clientes.cl_codigo,clientes.estado,clientes.cl_telefon," & _
                        "clientes.cl_dpto,clientes.cl_codconv from linmmdd inner join clientes on " & _
                        "linmmdd.cod_cli = clientes.cl_codigo where linmmdd.fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And linmmdd.fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and linmmdd.cod_prod =" & tm.Text & " And linmmdd.base =" & txt_b.Text & " And linmmdd.tot_lin >=" & 1 & " and linmmdd.tipo <>'" & "NOTA CR" & "'"
                        data_lin.Refresh
                     End If
                  Else
                     If Check3.Value = 1 Then
                        If Check8.Value = 1 Then
                           data_lin.RecordSource = "Select * from resplin where fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and cod_prod =" & tm.Text & " And base =" & txt_b.Text & " and tipo ='" & "CREDITO" & "'"
                           data_lin.Refresh
                        Else
                           data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.factura,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.cod_prod," & _
                           "linmmdd.nom_prod,linmmdd.nro_flia,linmmdd.nom_flia,linmmdd.convenio,linmmdd.tipo_mov," & _
                           "linmmdd.unidad,linmmdd.zona,linmmdd.servicio,linmmdd.tot_lin,linmmdd.nro_med_a,linmmdd.nom_med_a," & _
                           "linmmdd.mes_paga,linmmdd.ano_paga,linmmdd.costo_prod,linmmdd.base,linmmdd.ced_socio," & _
                           "linmmdd.grupo,linmmdd.tipo,linmmdd.hora,linmmdd.imp_iva,linmmdd.pendiente,linmmdd.rub_cont," & _
                           "linmmdd.fact,linmmdd.rub_nomb,linmmdd.reg_cab,linmmdd.nro_superv,linmmdd.nom_medic,linmmdd.ruc," & _
                           "linmmdd.nom_med_s,linmmdd.nom_superv,clientes.cl_codigo,clientes.estado,clientes.cl_telefon," & _
                           "clientes.cl_dpto,clientes.cl_codconv from linmmdd inner join clientes on " & _
                           "linmmdd.cod_cli = clientes.cl_codigo where linmmdd.fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And linmmdd.fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and linmmdd.cod_prod =" & tm.Text & " And linmmdd.base =" & txt_b.Text & " and linmmdd.tipo ='" & "CREDITO" & "'"
                           data_lin.Refresh
                        End If
                     Else
                        If Check8.Value = 1 Then
                           data_lin.RecordSource = "Select * from resplin where fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and cod_prod =" & tm.Text & " And base =" & txt_b.Text & " and tipo <>'" & "NOTA CR" & "'"
                           data_lin.Refresh
                        Else
                           If Check5.Value = 1 Then
                              data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.factura,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.cod_prod," & _
                              "linmmdd.nom_prod,linmmdd.nro_flia,linmmdd.nom_flia,linmmdd.convenio,linmmdd.tipo_mov," & _
                              "linmmdd.unidad,linmmdd.zona,linmmdd.servicio,linmmdd.tot_lin,linmmdd.nro_med_a,linmmdd.nom_med_a," & _
                              "linmmdd.mes_paga,linmmdd.ano_paga,linmmdd.costo_prod,linmmdd.base,linmmdd.ced_socio," & _
                              "linmmdd.grupo,linmmdd.tipo,linmmdd.hora,linmmdd.imp_iva,linmmdd.pendiente,linmmdd.rub_cont," & _
                              "linmmdd.fact,linmmdd.rub_nomb,linmmdd.reg_cab,linmmdd.nro_superv,linmmdd.nom_medic,linmmdd.ruc," & _
                              "linmmdd.nom_med_s,linmmdd.nom_superv,clientes.cl_codigo,clientes.estado,clientes.cl_telefon," & _
                              "clientes.cl_dpto,clientes.cl_codconv from linmmdd inner join clientes on " & _
                              "linmmdd.cod_cli = clientes.cl_codigo where linmmdd.fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And linmmdd.fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and linmmdd.cod_prod =" & tm.Text & " And linmmdd.base =" & txt_b.Text & " and clientes.estado in (2)"
                           Else
                              data_lin.RecordSource = "Select linmmdd.fecha,linmmdd.factura,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.cod_prod," & _
                              "linmmdd.nom_prod,linmmdd.nro_flia,linmmdd.nom_flia,linmmdd.convenio,linmmdd.tipo_mov," & _
                              "linmmdd.unidad,linmmdd.zona,linmmdd.servicio,linmmdd.tot_lin,linmmdd.nro_med_a,linmmdd.nom_med_a," & _
                              "linmmdd.mes_paga,linmmdd.ano_paga,linmmdd.costo_prod,linmmdd.base,linmmdd.ced_socio," & _
                              "linmmdd.grupo,linmmdd.tipo,linmmdd.hora,linmmdd.imp_iva,linmmdd.pendiente,linmmdd.rub_cont," & _
                              "linmmdd.fact,linmmdd.rub_nomb,linmmdd.reg_cab,linmmdd.nro_superv,linmmdd.nom_medic,linmmdd.ruc," & _
                              "linmmdd.nom_med_s,linmmdd.nom_superv,clientes.cl_codigo,clientes.estado,clientes.cl_telefon," & _
                              "clientes.cl_dpto,clientes.cl_codconv from linmmdd inner join clientes on " & _
                              "linmmdd.cod_cli = clientes.cl_codigo where linmmdd.fecha >=#" & Format(md.Text, "yyyy-mm-dd") & "# And linmmdd.fecha <=#" & Format(mh.Text, "yyyy-mm-dd") & "# and linmmdd.cod_prod =" & tm.Text & " And linmmdd.base =" & txt_b.Text
                           End If
                           data_lin.Refresh
                        End If
                     End If
                  End If
               End If
            End If
            If data_lin.Recordset.RecordCount > 0 Then
'               data_lin.Recordset.MoveLast
               barra.Visible = True
               barra.Max = data_lin.Recordset.RecordCount
               barra.Value = 0
               data_lin.Recordset.MoveFirst
               frm_vtasserv.MousePointer = 11
               Do While Not data_lin.Recordset.EOF
                  data_inf.Recordset.AddNew
                  data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
                  data_inf.Recordset("factura") = data_lin.Recordset("factura")
                  data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                  data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                  data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
                  data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
                  data_inf.Recordset("nro_flia") = data_lin.Recordset("nro_flia")
                  data_inf.Recordset("nom_flia") = data_lin.Recordset("nom_flia")
                  If Check9.Value = 1 Or Check10.Value = 1 Then
                     Data1.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("convenio") & "'"
                     Data1.Refresh
                     If Data1.Recordset.RecordCount > 0 Then
                        If IsNull(Data1.Recordset("cnv_grupo")) = False Then
                           If Trim(Data1.Recordset("cnv_grupo")) <> "" Then
                              data_inf.Recordset("tipo_mov") = "M"
                           Else
                              data_inf.Recordset("tipo_mov") = "S"
                           End If
                        Else
                           data_inf.Recordset("tipo_mov") = "S"
                        End If
                     Else
                        data_inf.Recordset("tipo_mov") = "S"
                     End If
                  End If
                  
                  If tm.Text = 992 Or tm.Text = 20033 Or t_rub.Text <> "" Or tm.Text = 802 Or tm.Text = 803 Or _
                     tm.Text = 984 Or tm.Text = 985 Or tm.Text = 986 Or tm.Text = 987 Or tm.Text = 989 Or _
                     tm.Text = 804 Or tm.Text = 805 Or tm.Text = 806 Then
                     Data1.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("cod_cli")
                     Data1.Refresh
                     If Data1.Recordset.RecordCount > 0 Then
                        If Data1.Recordset("estado") = 2 Or Data1.Recordset("estado") = 3 Then
                           data_inf.Recordset("libro") = "B"
                        Else
                           data_inf.Recordset("libro") = "A"
                        End If
                        If tm.Text = 20033 Then
                           If IsNull(Data1.Recordset("cl_telefon")) = False Then
                              If IsNull(Data1.Recordset("cl_dpto")) = False Then
                                 data_inf.Recordset("nom_flia") = Trim(Data1.Recordset("cl_telefon")) & "//" & Trim(Data1.Recordset("cl_dpto"))
                              Else
                                 data_inf.Recordset("nom_flia") = Trim(Data1.Recordset("cl_telefon"))
                              End If
                           Else
                              If IsNull(Data1.Recordset("cl_dpto")) = False Then
                                 data_inf.Recordset("nom_flia") = Trim(Data1.Recordset("cl_telefon")) & "//" & Trim(Data1.Recordset("cl_dpto"))
                              Else
                                 data_inf.Recordset("nom_flia") = Trim(Data1.Recordset("cl_telefon"))
                              End If
                           End If
                        Else
                           If Data1.Recordset("cl_codconv") <> data_lin.Recordset("convenio") Then
                              data_inf.Recordset("convenio") = Data1.Recordset("cl_codconv")
                           Else
                              data_inf.Recordset("convenio") = data_lin.Recordset("convenio")
                           End If
                        End If
                        data_inf.Recordset("zona") = Trim(str(Data1.Recordset("cl_cedula"))) & "-" & Trim(str(Data1.Recordset("cl_codced")))
                     Else
                        data_inf.Recordset("convenio") = data_lin.Recordset("convenio")
                     End If
                     If tm.Text = 20033 Then
                     Else
                        data_facafil.RecordSource = "select * from linmmdd_afil where factura =" & data_lin.Recordset("factura")
                        data_facafil.Refresh
                        If data_facafil.Recordset.RecordCount > 0 Then
                           data_inf.Recordset("nom_flia") = Mid(data_facafil.Recordset("nombre"), 1, 40)
                           data_inf.Recordset("servicio") = data_facafil.Recordset("codfunc")
                        Else
                           data_inf.Recordset("nom_flia") = "Sin promotor"
                           data_inf.Recordset("servicio") = 0
                        End If
                     End If
                  Else
                     data_inf.Recordset("convenio") = data_lin.Recordset("convenio")
                  End If
                  data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                  data_inf.Recordset("nro_med_a") = data_lin.Recordset("nro_med_a")
                  data_inf.Recordset("nom_med_a") = data_lin.Recordset("nom_med_a")
                  data_inf.Recordset("mes_paga") = data_lin.Recordset("mes_paga")
                  data_inf.Recordset("ano_paga") = data_lin.Recordset("ano_paga")
                  data_inf.Recordset("costo_prod") = 0
                  data_inf.Recordset("base") = data_lin.Recordset("base")
                  data_inf.Recordset("ced_socio") = data_lin.Recordset("ced_socio")
                  data_inf.Recordset("libro_rub") = data_lin.Recordset("unidad")
                  data_inf.Recordset("tipo") = data_lin.Recordset("tipo")
                  data_inf.Recordset("hora") = data_lin.Recordset("hora")
                  If IsNull(data_lin.Recordset("imp_iva")) = False Then
                     If tm.Text = 999 Then
'                        If data_lin.Recordset("imp_iva") <= 0 Then
'                           data_inf.Recordset("imp_iva") = data_lin.Recordset("tot_lin") / 1.1 * 0.1
'                        Else
                           data_inf.Recordset("imp_iva") = Format(data_lin.Recordset("imp_iva"), "Standard")
'                        End If
                     Else
                        data_inf.Recordset("imp_iva") = Format(data_lin.Recordset("imp_iva"), "Standard")
                     End If
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
                                If IsNull(data_inf.Recordset("imp_iva")) = False Then
                                   data_inf.Recordset("imp_iva") = data_lin.Recordset("imp_iva") * -1
                                   data_inf.Recordset("imp_iva") = Format(data_inf.Recordset("imp_iva"), "Standard")
                                End If
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
                  If Check2.Value = 1 Then
                     data_inf.Recordset("reg_cab") = data_lin.Recordset("reg_cab")
                  End If
                  data_inf.Recordset("rub_cont") = data_lin.Recordset("rub_cont")
                  data_inf.Recordset("rub_nomb") = data_lin.Recordset("rub_nomb")
                  If IsNull(data_lin.Recordset("grupo")) = False Then
                     data_inf.Recordset("nro_superv") = data_lin.Recordset("grupo")
                  Else
                     data_inf.Recordset("nro_superv") = 0
                  End If
                  If Check4.Value = 1 Then
                     Data2.RecordSource = "Select * from cabezal where factura =" & data_lin.Recordset("factura")
                     Data2.Refresh
                     If Data2.Recordset.RecordCount > 0 Then
                        data_inf.Recordset("nom_medic") = Mid(Data2.Recordset("dir_cli"), 1, 50)
                        data_inf.Recordset("ruc") = Mid(Data2.Recordset("loc_cli"), 1, 20)
                        data_inf.Recordset("nom_med_s") = Mid(Data2.Recordset("nom_cli"), 1, 40)
                        data_inf.Recordset("nom_superv") = Trim(str(Data2.Recordset("cod_cli"))) & "-" & Trim(str(Data2.Recordset("dias")))
                     End If
                  End If
                  
                  data_inf.Recordset("nom_med_s") = Mid(data_lin.Recordset("nom_medic"), 1, 40)
                  If t_rub.Text <> "" Then
                  Else
                     If IsNull(data_lin.Recordset("ced_socio")) = True Then
                        data_inf.Recordset("zona") = "0"
                     Else
                        If IsNull(data_lin.Recordset("fact")) = True Then
                           data_inf.Recordset("zona") = Trim(str(data_lin.Recordset("ced_socio"))) & "-0"
                        Else
                           data_inf.Recordset("zona") = Trim(str(data_lin.Recordset("ced_socio"))) & "-" & Trim(str(data_lin.Recordset("fact")))
                        End If
                     End If
                  End If
                  If tm.Text = 992 And data_lin.Recordset("tipo") = "CREDITO" Then
                     Data2.RecordSource = "Select * from linmmdd where factura =" & data_lin.Recordset("factura") & " and cod_prod in (883)"
                     Data2.Refresh
                     If Data2.Recordset.RecordCount > 0 Then
                        data_inf.Recordset("costo") = Data2.Recordset("tot_lin")
                        data_inf.Recordset("costo_prod") = data_lin.Recordset("tot_lin") + Data2.Recordset("tot_lin")
                     Else
                        data_inf.Recordset("costo") = 0
                        data_inf.Recordset("costo_prod") = data_lin.Recordset("tot_lin")
                     End If
                  End If
                  data_inf.Recordset.Update
                  data_lin.Recordset.MoveNext
                  barra.Max = data_lin.Recordset.RecordCount
                  barra.Value = barra.Value + 1
               Loop
               data_inf.Refresh
               DoEvents
               If tm.Text = 999 Or tm.Text = 993 Or tm.Text = 994 Then
               End If
               If Check6.Value = 1 Then
                  If data_inf.Recordset.RecordCount > 0 Then
                     data_inf.Recordset.MoveFirst
                     Do While Not data_inf.Recordset.EOF
                        If IsNull(data_inf.Recordset("libro_rub")) = True Then
                           data_inf.Recordset.Delete
                        End If
                        data_inf.Recordset.MoveNext
                     Loop
                  End If
               End If
               If Check7.Value = 1 Then
                  If data_inf.Recordset.RecordCount > 0 Then
                     data_inf.Recordset.MoveFirst
                     Do While Not data_inf.Recordset.EOF
                        If data_inf.Recordset("tipo") = "NOTA CR" Then
                        Else
                           data_inf.Recordset.Delete
                        End If
                        data_inf.Recordset.MoveNext
                     Loop
                     data_inf.RecordSource = "Select * from infvtas order by cod_prod"
                     data_inf.Refresh
                  End If
               Else
                  data_inf.RecordSource = "Select * from infvtas order by cod_prod"
                  data_inf.Refresh
               End If
               If tm.Text = 802 Or tm.Text = 803 Or tm.Text = 804 Or tm.Text = 805 Or tm.Text = 806 Then
                  data_inf.RecordSource = "Select * from infvtas order by cod_prod"
                  data_inf.Refresh
                  If data_inf.Recordset.RecordCount > 0 Then
                     data_inf.Recordset.MoveFirst
                     Do While Not data_inf.Recordset.EOF
                        If IsNull(data_inf.Recordset("libro")) = False Then
                           If data_inf.Recordset("libro") = "B" Then
                              data_inf.Recordset.Delete
                           End If
                        End If
                        data_inf.Recordset.MoveNext
                     Loop
                     data_inf.RecordSource = "Select * from infvtas order by cod_prod"
                     data_inf.Refresh
                  End If
               End If
               frm_vtasserv.MousePointer = 0
               If Check11.Value = 1 Then
                  If tm.Text = 30081 Then
                     Command2_Click
                  Else
                     Command1_Click
                  End If
               Else
                    If tm.Text = 999 Or tm.Text = 993 Or tm.Text = 994 Or Check9.Value = 1 Or Check10.Value = 1 Then
                       If Check1.Value = 1 Then
                          crs.ReportFileName = App.path & "\infvtasxser3.rpt"
                          crs.ReportTitle = "INFORME DE VENTAS POR SERVICIO FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy")
                          crs.Action = 1
                       
                          crs2.ReportFileName = App.path & "\infvtasxser5.rpt"
                          crs2.ReportTitle = "INFORME DE VENTAS POR SERVICIO FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy")
                          crs2.Action = 1
                       
                       Else
                          If Check9.Value = 1 Then
                             crs.ReportFileName = App.path & "\infvtasxsermm.rpt"
                             crs.ReportTitle = "INFORME DE VENTAS POR SERVICIO FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy")
                             crs.Action = 1
                          Else
                             If Check10.Value = 1 Then
                                crs.ReportFileName = App.path & "\infvtasxserms.rpt"
                                crs.ReportTitle = "INFORME DE VENTAS POR SERVICIO FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy")
                                crs.Action = 1
                             Else
                                crs.ReportFileName = App.path & "\infvtasxser2.rpt"
                                crs.ReportTitle = "INFORME DE VENTAS POR SERVICIO FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy")
                                crs.Action = 1
                             End If
                          End If
                       End If
                    Else
                       If tm.Text = 60103 Or tm.Text = 60105 Or tm.Text = 60106 Or _
                          tm.Text = 60107 Or tm.Text = 60108 Or tm.Text = 60109 Then
                          crs.ReportFileName = App.path & "\infvtasxsemm.rpt"
                          If txt_b.Text = 99 Then
                             crs.ReportTitle = "INFORME VENTAS POR SERVICIO FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy")
                          Else
                             crs.ReportTitle = "INFORME DE VENTAS POR SERVICIO FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy") & " BASE: " & txt_b.Text
                          End If
                          crs.Action = 1
                       Else
                         If Check1.Value = 1 Then
                            crs.ReportFileName = App.path & "\infvtasxsern.rpt"
                            If txt_b.Text = 99 Then
                               crs.ReportTitle = "INFORME VENTAS POR SERVICIO FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy")
                            Else
                               crs.ReportTitle = "INFORME DE VENTAS POR SERVICIO FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy") & " BASE: " & txt_b.Text
                            End If
                            crs.Action = 1
                         Else
                            If Check4.Value = 1 Then
                               crs.ReportFileName = App.path & "\infvtasxser6.rpt"
                               If txt_b.Text = 99 Then
                                  crs.ReportTitle = "INFORME DE VENTAS SERVICIOS CRÉDITO: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy")
                               Else
                                  crs.ReportTitle = "INFORME DE VENTAS SERVICIOS CRÉDITO: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy") & "  ==BASE: " & txt_b.Text
                               End If
                               crs.Action = 1
                            Else
                               If Check6.Value = 1 Then
                                  crs.ReportFileName = App.path & "\infvtasxser7.rpt"
                                  If txt_b.Text = 99 Then
                                     crs.ReportTitle = "INFORME DE VENTAS SERVICIOS SOCIOS CATEG.DIFERENTE: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy")
                                  Else
                                     crs.ReportTitle = "INFORME DE VENTAS SERVICIOS SOCIOS CATEG.DIFERENTE: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy") & "  ==BASE: " & txt_b.Text
                                  End If
                                  crs.Action = 1
                               Else
                                  If tm.Text = 992 Or tm.Text = 20033 Or t_rub.Text <> "" Or tm.Text = 802 Or _
                                     tm.Text = 984 Or tm.Text = 985 Or tm.Text = 986 Or tm.Text = 987 Or tm.Text = 989 Or _
                                     tm.Text = 803 Or tm.Text = 804 Or tm.Text = 805 Or tm.Text = 806 Then
                                     If t_rub.Text <> "" Then
                                        crs.ReportFileName = App.path & "\infvtasxser29.rpt"
                                        crs.ReportTitle = "INFORME VTAS.DE SERVICIO POR RUBRO: " & t_rub.Text & " DESDE: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy")
                                        
                                        crs.Action = 1
                                     Else
                                         If tm.Text = 20033 Then
                                            crs.ReportFileName = App.path & "\infvtasxser2tel.rpt"
                                            If txt_b.Text = 99 Then
                                               crs.ReportTitle = "INFORME DE SERVICIO VACUNACIÓN DESDE: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy")
                                            Else
                                               crs.ReportTitle = "INFORME DE SERVICIO VACUNACIÓN DESDE: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy") & "  ==BASE: " & txt_b.Text
                                            End If
                                            crs.Action = 1
                                         Else
                                            crs.ReportFileName = App.path & "\infvtasxafipro.rpt"
                                            If txt_b.Text = 99 Then
                                               crs.ReportTitle = "INFORME DE AFILIACIONES DESDE: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy")
                                            Else
                                               crs.ReportTitle = "INFORME DE AFILIACIONES DESDE: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy") & "  ==BASE: " & txt_b.Text
                                            End If
                                            crs.Action = 1
                                         End If
                                      End If
                                   Else
                                      If t_rub.Text <> "" Then
                                         crs.ReportFileName = App.path & "\infvtasxser29.rpt"
                                      Else
                                         crs.ReportFileName = App.path & "\infvtasxser.rpt"
                                      End If
                                      If txt_b.Text = 99 Then
                                         crs.ReportTitle = "INFORME DE VENTAS POR SERVICIO FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy")
                                      Else
                                         crs.ReportTitle = "INFORME DE VENTAS POR SERVICIO FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy") & "  ==BASE: " & txt_b.Text
                                      End If
                                      crs.Action = 1
                                   End If
                               End If
                            End If
                         End If
                      End If
                    End If
               End If
               barra.Visible = False
            Else
               MsgBox "No existen registros con esta selección", vbInformation, "Mensaje"
            End If
            data_lin.Recordset.Close
            
         Else
            MsgBox "Ingrese Base", vbInformation, "Mensaje"
            txt_b.SetFocus
         End If
      Else
         MsgBox "Número de SERVICIO incorrecto", vbInformation, "Mensaje"
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
               
End Sub

Private Sub b_canc_Click()
Unload Me

End Sub

Private Sub Command1_Click()
Dim Xobjexel22 As Excel.Application
Dim Xlibexel22 As Excel.Workbook

Dim Xarchexel22 As New Excel.Worksheet
Dim Xlin, XCol As Integer
Dim Xtotreg, Xsub As Long
Dim Xarchtex As String
Dim Textofecha As String
Dim Xpromotor As Integer

Dim Xlabrir3 As New Excel.Application

frm_vtasserv.MousePointer = 11

Xlin = 1
XCol = 1
Xtotreg = 0
Xsub = 0
Set Xobjexel22 = New Excel.Application
Set Xlibexel22 = Xobjexel22.Workbooks.Add
Set Xarchexel22 = Xlibexel22.Worksheets.Add
Xarchexel22.Name = Trim("Ventas")
Xlibexel22.SaveAs ("C:\planillas\InfoVentas.xls")
Xarchtex = "C:\planillas\InfoVentas.xls"

Xarchexel22.Cells(Xlin, XCol) = "DEPARTAMENTO TI SAPP S.A."
Xlin = Xlin + 1
XCol = XCol + 1
Xarchexel22.Range("A1", "C3").Font.Size = 16
Xarchexel22.Range("A" & Trim(str(Xlin)), "O" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)

Xarchexel22.Cells(Xlin, XCol) = "VENTAS POR SERVICIO: " & DBCombo1.Text & " DESDE: " & md.Text & " HASTA: " & mh.Text
        
XCol = 1
Xlin = Xlin + 2
Xnrocan = Xnrocan + Xlin
        
Xarchexel22.Range("A" & Trim(str(Xlin)), "O" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
Xarchexel22.Range("A" & Trim(str(Xlin))).ColumnWidth = 12
Xarchexel22.Cells(Xlin, XCol) = "FECHA"
XCol = XCol + 1
Xarchexel22.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
Xarchexel22.Cells(Xlin, XCol) = "MATRICULA"
XCol = XCol + 1
Xarchexel22.Range("C" & Trim(str(Xlin))).ColumnWidth = 35
Xarchexel22.Cells(Xlin, XCol) = "NOMBRE"
XCol = XCol + 1
Xarchexel22.Range("D" & Trim(str(Xlin))).ColumnWidth = 13
Xarchexel22.Cells(Xlin, XCol) = "CEDULA"
XCol = XCol + 1
Xarchexel22.Range("E" & Trim(str(Xlin))).ColumnWidth = 20
Xarchexel22.Cells(Xlin, XCol) = "ZONA"
XCol = XCol + 1
Xarchexel22.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
Xarchexel22.Cells(Xlin, XCol) = "CONVENIO"
XCol = XCol + 1
Xarchexel22.Range("G" & Trim(str(Xlin))).ColumnWidth = 10
Xarchexel22.Cells(Xlin, XCol) = "CUOTA"
XCol = XCol + 1
Xarchexel22.Range("H" & Trim(str(Xlin))).ColumnWidth = 12
Xarchexel22.Cells(Xlin, XCol) = "DESCUENTO"
XCol = XCol + 1
Xarchexel22.Range("I" & Trim(str(Xlin))).ColumnWidth = 10
Xarchexel22.Cells(Xlin, XCol) = "TOTAL $."
XCol = XCol + 1
Xarchexel22.Range("J" & Trim(str(Xlin))).ColumnWidth = 8
Xarchexel22.Cells(Xlin, XCol) = "BASE"
XCol = XCol + 1
Xarchexel22.Range("K" & Trim(str(Xlin))).ColumnWidth = 10
Xarchexel22.Cells(Xlin, XCol) = "BOLETA"
XCol = XCol + 1
Xarchexel22.Range("L" & Trim(str(Xlin))).ColumnWidth = 13
Xarchexel22.Cells(Xlin, XCol) = "TIPO DOC"
XCol = XCol + 1
Xarchexel22.Range("M" & Trim(str(Xlin))).ColumnWidth = 13
Xarchexel22.Cells(Xlin, XCol) = "COD.PROMO"

XCol = XCol + 1
Xarchexel22.Range("N" & Trim(str(Xlin))).ColumnWidth = 30
If tm.Text = 992 Or _
   tm.Text = 984 Or _
   tm.Text = 985 Or _
   tm.Text = 986 Or _
   tm.Text = 987 Or _
   tm.Text = 989 Or _
   tm.Text = 802 Or _
   tm.Text = 803 Or _
   tm.Text = 804 Or _
   tm.Text = 805 Or _
   tm.Text = 806 Then
   Xarchexel22.Cells(Xlin, XCol) = "PROMOTOR"
Else
   Xarchexel22.Cells(Xlin, XCol) = "SERVICIO"
End If
XCol = XCol + 1
Xarchexel22.Range("O" & Trim(str(Xlin))).ColumnWidth = 15
Xarchexel22.Cells(Xlin, XCol) = "RECIBO PAGO"

Xlin = Xlin + 1
XCol = 1
        
'data_infccou.DatabaseName = App.path & "\informess.mdb"
If tm.Text = 992 Or _
   tm.Text = 984 Or _
   tm.Text = 985 Or _
   tm.Text = 986 Or _
   tm.Text = 987 Or _
   tm.Text = 989 Or _
   tm.Text = 802 Or _
   tm.Text = 803 Or _
   tm.Text = 804 Or _
   tm.Text = 805 Or _
   tm.Text = 806 Then
   
   data_inf.RecordSource = "select * from infvtas order by servicio"
Else
   data_inf.RecordSource = "select * from infvtas order by cod_prod"
End If

data_inf.Refresh
   
If data_inf.Recordset.RecordCount > 0 Then
   data_inf.Recordset.MoveFirst
   If tm.Text = 992 Or _
      tm.Text = 984 Or _
      tm.Text = 985 Or _
      tm.Text = 986 Or _
      tm.Text = 987 Or _
      tm.Text = 989 Or _
      tm.Text = 802 Or _
      tm.Text = 803 Or _
      tm.Text = 804 Or _
      tm.Text = 805 Or _
      tm.Text = 806 Then
      If IsNull(data_inf.Recordset("servicio")) = False Then
         Xpromotor = data_inf.Recordset("servicio")
      Else
         Xpromotor = 0
      End If
      Xarchexel22.Cells(Xlin, XCol) = "PROMOTOR:"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_inf.Recordset("nom_flia")
      Xlin = Xlin + 1
      XCol = 1
   Else
      Xpromotor = data_inf.Recordset("cod_prod")
      Xarchexel22.Cells(Xlin, XCol) = "SERVICIO:"
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_inf.Recordset("nom_prod")
      Xlin = Xlin + 1
      XCol = 1
   End If
   Xsub = 0
   Do While Not data_inf.Recordset.EOF
      If tm.Text = 992 Or _
         tm.Text = 984 Or _
         tm.Text = 985 Or _
         tm.Text = 986 Or _
         tm.Text = 987 Or _
         tm.Text = 989 Or _
         tm.Text = 802 Or _
         tm.Text = 803 Or _
         tm.Text = 804 Or _
         tm.Text = 805 Or _
         tm.Text = 806 Then
         If Xpromotor = data_inf.Recordset("servicio") Then
            Xsub = Xsub + 1
         Else
            Xarchexel22.Cells(Xlin, XCol) = "Sub-Total promotor:" & Trim(str(Xsub))
            Xlin = Xlin + 1
            XCol = 1
            Xarchexel22.Cells(Xlin, XCol) = "PROMOTOR:"
            XCol = XCol + 1
            Xarchexel22.Cells(Xlin, XCol) = data_inf.Recordset("nom_flia")
            Xlin = Xlin + 1
            XCol = 1
            Xsub = 1
            If IsNull(data_inf.Recordset("servicio")) = False Then
               Xpromotor = data_inf.Recordset("servicio")
            Else
               Xpromotor = 0
            End If
         End If
      Else
         If Xpromotor = data_inf.Recordset("servicio") Then
            Xsub = Xsub + 1
         Else
            Xarchexel22.Cells(Xlin, XCol) = "Sub-Total servicio:" & Trim(str(Xsub))
            Xlin = Xlin + 1
            XCol = 1
            Xarchexel22.Cells(Xlin, XCol) = "SERVICIO:"
            XCol = XCol + 1
            Xarchexel22.Cells(Xlin, XCol) = data_inf.Recordset("nom_prod")
            Xlin = Xlin + 1
            XCol = 1
            Xsub = 1
            Xpromotor = data_inf.Recordset("cod_prod")
         End If
      End If
      
      Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_inf.Recordset("fecha"), "dd/mm/yyyy")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_inf.Recordset("cod_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_inf.Recordset("nom_cli")
      XCol = XCol + 1
      If IsNull(data_inf.Recordset("zona")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_inf.Recordset("zona")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Data1.RecordSource = "select * from clientes where cl_codigo =" & data_inf.Recordset("cod_cli")
      Data1.Refresh
      If Data1.Recordset.RecordCount > 0 Then
         If IsNull(Data1.Recordset("cl_zona")) = False Then
            Xarchexel22.Cells(Xlin, XCol) = Data1.Recordset("cl_zona")
         Else
            Xarchexel22.Cells(Xlin, XCol) = "Sin Zona"
         End If
      Else
         Xarchexel22.Cells(Xlin, XCol) = "Sin datos"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_inf.Recordset("convenio")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_inf.Recordset("tot_lin")
      XCol = XCol + 1
      If IsNull(data_inf.Recordset("costo")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_inf.Recordset("costo")
      End If
      XCol = XCol + 1
      If IsNull(data_inf.Recordset("costo_prod")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_inf.Recordset("costo_prod")
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_inf.Recordset("base")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_inf.Recordset("factura")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_inf.Recordset("tipo")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_inf.Recordset("servicio")
      
      XCol = XCol + 1
      If tm.Text = 992 Or _
         tm.Text = 984 Or _
         tm.Text = 985 Or _
         tm.Text = 986 Or _
         tm.Text = 987 Or _
         tm.Text = 989 Or _
         tm.Text = 802 Or _
         tm.Text = 803 Or _
         tm.Text = 804 Or _
         tm.Text = 805 Or _
         tm.Text = 806 Then
         Xarchexel22.Cells(Xlin, XCol) = data_inf.Recordset("nom_flia")
      Else
         Xarchexel22.Cells(Xlin, XCol) = data_inf.Recordset("nom_prod")
      End If
      XCol = XCol + 1
      If data_inf.Recordset("tipo") = "e-Tck CREDITO" Then
         If Format(mh.Text, "yyyy/mm/dd") <= Format("30/09/2022", "yyyy/mm/dd") Then
            Data1.RecordSource = "select linmmdd.factura,linmmdd.pendiente,linmmdd.porce_est,linmmdd.tot_lin,linmmdd.fecha,linmmdd.cod_cli,linmmdd.cod_prod from linmmdd where linmmdd.tot_lin =" & data_inf.Recordset("costo_prod") & " and cod_cli=" & data_inf.Recordset("cod_cli") & " and linmmdd.pendiente in ('Z') and fecha >=#" & Format(data_inf.Recordset("fecha"), "yyyy/mm/dd") & "# and cod_prod in (997)"
         Else
            Data1.RecordSource = "select linmmdd.factura,linmmdd.pendiente,linmmdd.porce_est from linmmdd where linmmdd.porce_est =" & data_inf.Recordset("factura") & " and linmmdd.pendiente in ('Z')"
         End If
         Data1.Refresh
         If Data1.Recordset.RecordCount > 0 Then
            Xarchexel22.Cells(Xlin, XCol) = Data1.Recordset("factura")
         Else
            Xarchexel22.Cells(Xlin, XCol) = "Sin recibo"
         End If
      End If
         
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      data_inf.Recordset.MoveNext
   Loop
End If
   
Xarchexel22.Cells(Xlin, XCol) = "Total Registros: " & Trim(str(Xtotreg))
Xsub = 0
Xlin = Xlin + 1
XCol = 1
   
Xarchexel22.Cells(Xlin, XCol) = "FECHA DE EMISION:" & Format(Date, "dd/mm/yyyy")

Xlibexel22.Save
Xlibexel22.Close
Xobjexel22.Quit
Xlabrir3.Workbooks.Open Xarchtex, , False
Xlabrir3.Visible = True
Xlabrir3.WindowState = xlMaximized

frm_vtasserv.MousePointer = 0

MsgBox "Proceso terminado"

End Sub

Private Sub Command2_Click()
Dim Xobjexel22 As Excel.Application
Dim Xlibexel22 As Excel.Workbook

Dim Xarchexel22 As New Excel.Worksheet
Dim Xlin, XCol As Integer
Dim Xtotreg, Xsub As Long
Dim Xarchtex As String
Dim Textofecha As String
Dim Xpromotor As Integer
Dim SoloMutNO As String

Dim Xlabrir3 As New Excel.Application

frm_vtasserv.MousePointer = 11

Xlin = 1
XCol = 1
Xtotreg = 0
Xsub = 0
Set Xobjexel22 = New Excel.Application
Set Xlibexel22 = Xobjexel22.Workbooks.Add
Set Xarchexel22 = Xlibexel22.Worksheets.Add
Xarchexel22.Name = Trim("Ventas")
Xlibexel22.SaveAs ("C:\planillas\InfoVentas.xls")
Xarchtex = "C:\planillas\InfoVentas.xls"

Xarchexel22.Cells(Xlin, XCol) = "DEPARTAMENTO TI SAPP S.A."
Xlin = Xlin + 1
XCol = XCol + 1
Xarchexel22.Range("A1", "C3").Font.Size = 16
Xarchexel22.Range("A" & Trim(str(Xlin)), "K" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)

Xarchexel22.Cells(Xlin, XCol) = "VENTAS POR SERVICIO: " & DBCombo1.Text & " DESDE: " & md.Text & " HASTA: " & mh.Text
        
XCol = 1
Xlin = Xlin + 2
Xnrocan = Xnrocan + Xlin
        
Xarchexel22.Range("A" & Trim(str(Xlin)), "K" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
Xarchexel22.Range("A" & Trim(str(Xlin))).ColumnWidth = 12
Xarchexel22.Cells(Xlin, XCol) = "FECHA"
XCol = XCol + 1
Xarchexel22.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
Xarchexel22.Cells(Xlin, XCol) = "MATRICULA"
XCol = XCol + 1
Xarchexel22.Range("C" & Trim(str(Xlin))).ColumnWidth = 35
Xarchexel22.Cells(Xlin, XCol) = "NOMBRE"
XCol = XCol + 1
Xarchexel22.Range("D" & Trim(str(Xlin))).ColumnWidth = 16
Xarchexel22.Cells(Xlin, XCol) = "CEDULA"
XCol = XCol + 1
Xarchexel22.Range("E" & Trim(str(Xlin))).ColumnWidth = 15
Xarchexel22.Cells(Xlin, XCol) = "TIENE CARTA"
XCol = XCol + 1
Xarchexel22.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
Xarchexel22.Cells(Xlin, XCol) = "CONVENIO"
XCol = XCol + 1
Xarchexel22.Range("G" & Trim(str(Xlin))).ColumnWidth = 10
Xarchexel22.Cells(Xlin, XCol) = "CODSERV"
XCol = XCol + 1
Xarchexel22.Range("H" & Trim(str(Xlin))).ColumnWidth = 10
Xarchexel22.Cells(Xlin, XCol) = "BASE"
Xlin = Xlin + 1
XCol = 1
        
'data_infccou.DatabaseName = App.path & "\informess.mdb"
   
SoloMutNO = MsgBox("Desea incluir solo categorías NOSAPP?", vbInformation + vbYesNo, "Ventas")
If SoloMutNO = vbYes Then
   data_inf.RecordSource = "select * from infvtas where convenio in ('SMIN','SMINA','UNIVS','UNNSAM','HEVANO','EVNSAM','CCNOS','CCNSAM','GANOS','CASANO','CASNSA') order by cod_prod"
Else
   data_inf.RecordSource = "select * from infvtas order by cod_prod"
End If
data_inf.Refresh

If data_inf.Recordset.RecordCount > 0 Then
   data_inf.Recordset.MoveFirst
   Xtotreg = 0
   Do While Not data_inf.Recordset.EOF
      
      Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_inf.Recordset("fecha"), "dd/mm/yyyy")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_inf.Recordset("cod_cli")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_inf.Recordset("nom_cli")
      XCol = XCol + 1
      If IsNull(data_inf.Recordset("zona")) = False Then
         Xarchexel22.Cells(Xlin, XCol) = data_inf.Recordset("zona")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "0-0"
      End If
      XCol = XCol + 1
      Data1.RecordSource = "Select * from linmmdd where cod_cli =" & data_inf.Recordset("cod_cli") & " and cod_prod in (802,803,804,805,806)"
      Data1.Refresh
      If Data1.Recordset.RecordCount > 0 Then
         Xarchexel22.Cells(Xlin, XCol) = "SI"
      Else
         Xarchexel22.Cells(Xlin, XCol) = "NO"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_inf.Recordset("convenio")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_inf.Recordset("cod_prod")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = data_inf.Recordset("base")
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      data_inf.Recordset.MoveNext
   Loop
End If
   
Xarchexel22.Cells(Xlin, XCol) = "Total Registros: " & Trim(str(Xtotreg))
Xsub = 0
Xlin = Xlin + 1
XCol = 1
   
Xarchexel22.Cells(Xlin, XCol) = "FECHA DE EMISION:" & Format(Date, "dd/mm/yyyy")

Xlibexel22.Save
Xlibexel22.Close
Xobjexel22.Quit
Xlabrir3.Workbooks.Open Xarchtex, , False
Xlabrir3.Visible = True
Xlabrir3.WindowState = xlMaximized

frm_vtasserv.MousePointer = 0

MsgBox "Proceso terminado"

End Sub

Private Sub DBCombo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_b.SetFocus
End If

End Sub

Private Sub DBCombo1_LostFocus()

If DBCombo1.Text <> "" Then
    If IsNumeric(DBCombo1.Text) = True Then
       data_med.Recordset.FindFirst "codest =" & DBCombo1.Text
       If Not data_med.Recordset.NoMatch Then
          tm.Text = data_med.Recordset("codest")
          DBCombo1.Text = data_med.Recordset("descrip")
       Else
          MsgBox "No encontrado, consulte por nombre", vbInformation, "Mensaje"
          DBCombo1.SetFocus
       End If
    Else
       data_med.Recordset.FindFirst "descrip ='" & DBCombo1.Text & "'"
       If Not data_med.Recordset.NoMatch Then
          tm.Text = data_med.Recordset("codest")
          DBCombo1.Text = data_med.Recordset("descrip")
       Else
          MsgBox "No encontrado, consulte por nombre", vbInformation, "Mensaje"
          DBCombo1.SetFocus
       End If
    End If
Else
    tm.Text = 999999
    
End If

End Sub

Private Sub Form_Load()
data_med.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_facafil.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_med.RecordSource = "Select * from estudios order by descrip"
data_med.Refresh
data_inf.DatabaseName = App.path & "\informes.mdb"
data_lin.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
'Data1.RecordSource = "clientes"
'Data1.Refresh
'Data2.DatabaseName = ""
'Data2.ConnectionString = "dsn=" & Xconexrmt
Data2.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_emi.Connect = "odbc;dsn=" & Xconexrmt & ";"
'Data2.RecordSource = "Select * from cabezal"
'Data2.Refresh


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
