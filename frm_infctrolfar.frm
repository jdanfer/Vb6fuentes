VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infctrolfar 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes controles de farmacia"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6345
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_infctrolfar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   6345
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "Excluir móviles"
      Height          =   255
      Left            =   1440
      TabIndex        =   31
      Top             =   7560
      Width           =   2055
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   1980
   End
   Begin MSAdodcLib.Adodc data_lin 
      Height          =   330
      Left            =   2520
      Top             =   7800
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
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
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   5160
      Top             =   5640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Tipo de informe"
      Height          =   855
      Left            =   240
      TabIndex        =   11
      Top             =   6600
      Width           =   5895
      Begin VB.OptionButton Option7 
         BackColor       =   &H00FF0000&
         Caption         =   "Detalle"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2760
         TabIndex        =   13
         Top             =   360
         Width           =   2055
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FF0000&
         Caption         =   "Resumen"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5640
      Picture         =   "frm_infctrolfar.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Salir"
      Top             =   7560
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_infctrolfar.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Procesar"
      Top             =   7560
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Opciones de informes"
      Height          =   6495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      Begin VB.Data data_sql 
         Caption         =   "data_sql"
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
         Top             =   2280
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Data data_mdb 
         Caption         =   "data_mdb"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   840
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2160
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton Command3 
         Caption         =   "con prescrip"
         Height          =   375
         Left            =   4680
         TabIndex        =   30
         Top             =   6000
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.OptionButton Option13 
         BackColor       =   &H00C00000&
         Caption         =   "Cantidad de socios con prescripción"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   6120
         Width           =   4335
      End
      Begin VB.OptionButton Option12 
         BackColor       =   &H00C00000&
         Caption         =   "Bajas de recetas por motivo"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   5760
         Width           =   4335
      End
      Begin VB.OptionButton Option11 
         BackColor       =   &H00C00000&
         Caption         =   "Próximas recetas para retirar"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   5400
         Width           =   4335
      End
      Begin MSAdodcLib.Adodc adocli 
         Height          =   375
         Left            =   2760
         Top             =   480
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
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
         Caption         =   "adocli"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.TextBox t_edh 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   2880
         TabIndex        =   26
         Top             =   1920
         Width           =   855
      End
      Begin VB.TextBox t_edd 
         Alignment       =   2  'Center
         Height          =   360
         Left            =   1800
         TabIndex        =   25
         Top             =   1920
         Width           =   855
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
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
         Top             =   1560
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1800
         TabIndex        =   23
         Top             =   1440
         Width           =   855
      End
      Begin VB.Data data_inf 
         Caption         =   "data_inf"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   3120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   120
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.OptionButton Option10 
         BackColor       =   &H00C00000&
         Caption         =   "Total medicación facturada"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   5040
         Width           =   4335
      End
      Begin VB.TextBox t_base 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   4440
         TabIndex        =   20
         Top             =   840
         Width           =   735
      End
      Begin MSMask.MaskEdBox mhh 
         Height          =   375
         Left            =   2760
         TabIndex        =   18
         Top             =   840
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   5
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mhd 
         Height          =   375
         Left            =   1800
         TabIndex        =   17
         Top             =   840
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   5
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin VB.OptionButton Option9 
         BackColor       =   &H00C00000&
         Caption         =   "Devoluciones facturadas en base"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   4680
         Width           =   4335
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00C00000&
         Caption         =   "Medicación solicitada a mutualistas"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   4320
         Width           =   4335
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C00000&
         Caption         =   "Medicación enviada a base"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   3240
         Width           =   4335
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C00000&
         Caption         =   "Medicación sin entregar (Devoluciones)"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   3960
         Width           =   4335
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C00000&
         Caption         =   "Medicación solicitada a Farm.central"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   3600
         Width           =   4335
      End
      Begin MSMask.MaskEdBox mfh 
         Height          =   375
         Left            =   3840
         TabIndex        =   7
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfd 
         Height          =   375
         Left            =   2280
         TabIndex        =   6
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C00000&
         Caption         =   "Medicación pendiente"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         ToolTipText     =   "Se puede omitir ingresar rango de fechas y listará TODO lo pendiente"
         Top             =   2880
         Width           =   4335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C00000&
         Caption         =   "Medicación entregada"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   2520
         Width           =   4335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
         Caption         =   "Rango edad:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   1920
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Cant.Regs:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   22
         Top             =   1440
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "BASE:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3720
         TabIndex        =   19
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Rango horario:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   840
         Width           =   1575
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   0
         X2              =   5880
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Rango de fechas:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   4080
      Picture         =   "frm_infctrolfar.frx":0F56
      Stretch         =   -1  'True
      Top             =   7200
      Width           =   1695
   End
End
Attribute VB_Name = "frm_infctrolfar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Xtotmedsol, Xtotmedent As Long
Dim Xelporce As Double
Dim Xed As Double
Xed = 0

Xtotmedsol = 0
Xtotmedent = 0
Command1.Enabled = False
Command2.Enabled = False
frm_infctrolfar.MousePointer = 11

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")
Dim Xdocmedica As Integer

MiBaseact.Execute "Delete * from infvtas"

data_inf.RecordSource = "infvtas"
data_inf.Refresh

data_lin.ConnectionString = "dsn=" & Xconexrmt
Data1.DatabaseName = App.path & "\informes.mdb"

If mfd.Text <> "__/__/____" Then
   If mfh.Text <> "__/__/____" Then
      If Option4.Value = True Then
         If t_base.Text <> "" Then
            data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and tot_lin >=" & 0 & " and dias in (0,1,2,3,5,6,7,8) and base =" & t_base.Text
            data_lin.Refresh
         Else
            data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and tot_lin >=" & 0 & " and dias in (0,1,2,3,5,6,7,8)"
            data_lin.Refresh
         End If
      Else
         If Option9.Value = True Then
            If t_base.Text <> "" Then
               data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and tot_lin >=" & 0 & " and dias in (0,1,2,3,5,6,7,8) and base =" & t_base.Text
               data_lin.Refresh
            Else
               data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and tot_lin >=" & 0 & " and dias in (0,1,2,3,5,6,7,8)"
               data_lin.Refresh
            End If
         Else
            If Option10.Value = True Then
               If t_base.Text <> "" Then
                  data_lin.RecordSource = "Select linmmdd.dias,linmmdd.tot_lin,linmmdd.fecha,linmmdd.hora,linmmdd.cod_cli," & _
                  "linmmdd.nom_cli,linmmdd.cod_prod,linmmdd.nom_prod,linmmdd.base,linmmdd.factura,linmmdd.nom_medic,linmmdd.ced_socio," & _
                  "linmmdd.fact,linmmdd.nom_med_s,linmmdd.vto,linmmdd.margen_prd,linmmdd.pre_prod,linmmdd.costo,linmmdd.nro_superv,linmmdd.zona,pendiente,linmmdd.nro_flia," & _
                  "clientes.cl_codigo,clientes.cl_fnac from linmmdd inner join clientes on linmmdd.cod_cli=clientes.cl_codigo where linmmdd.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And linmmdd.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and linmmdd.nro_flia =" & 6 & " and linmmdd.tot_lin >=" & 0 & " and linmmdd.base =" & t_base.Text
'                  data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and nro_flia =" & 6 & " and tot_lin >=" & 0 & " and base =" & t_base.Text
                  data_lin.Refresh
               Else
                  data_lin.RecordSource = "Select linmmdd.dias,linmmdd.tot_lin,linmmdd.fecha,linmmdd.hora,linmmdd.cod_cli," & _
                  "linmmdd.nom_cli,linmmdd.cod_prod,linmmdd.nom_prod,linmmdd.base,linmmdd.factura,linmmdd.nom_medic,linmmdd.ced_socio," & _
                  "linmmdd.fact,linmmdd.nom_med_s,linmmdd.vto,linmmdd.margen_prd,linmmdd.pre_prod,linmmdd.costo,linmmdd.nro_superv,linmmdd.zona,pendiente,linmmdd.nro_flia," & _
                  "clientes.cl_codigo,clientes.cl_fnac from linmmdd inner join clientes on linmmdd.cod_cli=clientes.cl_codigo where linmmdd.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And linmmdd.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and linmmdd.nro_flia =" & 6 & " and linmmdd.tot_lin >=" & 0
                  data_lin.Refresh
               End If
            Else
               If t_base.Text <> "" Then
                  data_lin.RecordSource = "Select linmmdd.dias,linmmdd.tot_lin,linmmdd.fecha,linmmdd.hora,linmmdd.cod_cli," & _
                  "linmmdd.nom_cli,linmmdd.cod_prod,linmmdd.nom_prod,linmmdd.base,linmmdd.factura,linmmdd.nom_medic,linmmdd.ced_socio," & _
                  "linmmdd.fact,linmmdd.nom_med_s,linmmdd.vto,linmmdd.margen_prd,linmmdd.pre_prod,linmmdd.costo,linmmdd.nro_superv,linmmdd.zona,pendiente,linmmdd.nro_flia," & _
                  "clientes.cl_codigo,clientes.cl_fnac from linmmdd inner join clientes on linmmdd.cod_cli=clientes.cl_codigo where linmmdd.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And linmmdd.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and linmmdd.nro_flia =" & 6 & " and linmmdd.tot_lin >=" & 0 & " and linmmdd.dias in (0,1,3,5,2,7,4,8,9,10,11) and linmmdd.base =" & t_base.Text
                  data_lin.Refresh
               Else
                  data_lin.RecordSource = "Select linmmdd.dias,linmmdd.tot_lin,linmmdd.fecha,linmmdd.hora,linmmdd.cod_cli," & _
                  "linmmdd.nom_cli,linmmdd.cod_prod,linmmdd.nom_prod,linmmdd.base,linmmdd.factura,linmmdd.nom_medic,linmmdd.ced_socio," & _
                  "linmmdd.fact,linmmdd.nom_med_s,linmmdd.vto,linmmdd.margen_prd,linmmdd.pre_prod,linmmdd.costo,linmmdd.nro_superv,linmmdd.zona,pendiente,linmmdd.nro_flia," & _
                  "clientes.cl_codigo,clientes.cl_fnac from linmmdd inner join clientes on linmmdd.cod_cli=clientes.cl_codigo where linmmdd.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And linmmdd.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and linmmdd.nro_flia =" & 6 & " and linmmdd.tot_lin >=" & 0 & " and linmmdd.dias in (0,1,3,5,2,7,4,8,9,10,11)"
                  data_lin.Refresh
               End If
            End If
         End If
      End If
      If Option11.Value = True Then
         data_lin.RecordSource = "select * from hc_prescrip where hc_comfec >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and hc_comfec <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and hc_tippresd in ('MEDICACION','RECETA PACIENTE CRONICO') and hc_fecentrega is null and hc_codmedica is not null order by hc_comfec"
         data_lin.Refresh
      End If
      If Option12.Value = True Then
         data_lin.RecordSource = "select * from hc_prescrip where hc_comfec >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and hc_comfec <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and hc_tippresd in ('MEDICACION','RECETA PACIENTE CRONICO') and motivo_cance is not null order by hc_comfec"
         data_lin.Refresh
      End If
      
      If data_lin.Recordset.RecordCount > 0 Then
         If Option13.Value = False Then
            data_lin.Recordset.MoveLast
            Xtotmedsol = data_lin.Recordset.RecordCount
            data_lin.Recordset.MoveFirst
            Do While Not data_lin.Recordset.EOF
               If Option11.Value = True Then
                  Xdocmedica = data_lin.Recordset("hc_codmedica")
                  If Xdocmedica > 0 Then
                     data_inf.Recordset.AddNew
                     data_inf.Recordset("fecha") = data_lin.Recordset("hc_comfec")
                     data_inf.Recordset("cod_cli") = data_lin.Recordset("hc_mat")
                     data_inf.Recordset("cod_prod") = data_lin.Recordset("hc_nro")
                     data_inf.Recordset("nom_prod") = Mid(data_lin.Recordset("hc_descrip"), 1, 50)
                     data_inf.Recordset("nom_med_s") = data_lin.Recordset("hc_tippresd")
                     data_inf.Recordset("nom_med_a") = Mid(data_lin.Recordset("hc_indicanom"), 1, 40)
                     data_inf.Recordset("realizada") = data_lin.Recordset("hc_hastaf")
                     data_inf.Recordset.Update
                  End If
               End If
               If Option12.Value = True Then
                  data_inf.Recordset.AddNew
                  data_inf.Recordset("fecha") = data_lin.Recordset("hc_comfec")
                  data_inf.Recordset("cod_cli") = data_lin.Recordset("hc_mat")
                  data_inf.Recordset("cod_prod") = data_lin.Recordset("hc_nro")
                  data_inf.Recordset("nom_prod") = Mid(data_lin.Recordset("hc_descrip"), 1, 50)
                  data_inf.Recordset("nom_med_s") = data_lin.Recordset("hc_tippresd")
                  data_inf.Recordset("nom_med_a") = Mid(data_lin.Recordset("hc_indicanom"), 1, 40)
                  data_inf.Recordset("realizada") = data_lin.Recordset("hc_hastaf")
                  data_inf.Recordset("nom_flia") = Mid(data_lin.Recordset("motivo_cance"), 1, 40)
                  data_inf.Recordset.Update
               End If
               
               If Option1.Value = True Then
                  If IsNull(data_lin.Recordset("dias")) = False Then
                     If data_lin.Recordset("dias") = 1 Or _
                        data_lin.Recordset("dias") = 3 Or _
                        data_lin.Recordset("dias") = 8 Then
                        If data_lin.Recordset("tot_lin") >= 0 Then
                           If mhd.Text = "__:__" And mhh.Text = "__:__" Then
                              data_inf.Recordset.AddNew
                              data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
                              data_inf.Recordset("hora") = data_lin.Recordset("hora")
                              data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                              data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                              data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
                              data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
                              data_inf.Recordset("base") = data_lin.Recordset("base")
                              data_inf.Recordset("factura") = data_lin.Recordset("factura")
                              data_inf.Recordset("nom_medic") = data_lin.Recordset("nom_medic")
                              data_inf.Recordset("realizada") = data_lin.Recordset("cl_fnac")
                              If IsNull(data_lin.Recordset("ced_socio")) = False Then
                                 If IsNull(data_lin.Recordset("fact")) = False Then
                                    data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-" & Trim(str(data_lin.Recordset("fact")))
                                 Else
                                    data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-"
                                 End If
                              Else
                                 data_inf.Recordset("nom_med_s") = "0"
                              End If
                              data_inf.Recordset("nro_med_s") = Xtotmedsol
                              data_inf.Recordset("vto") = data_lin.Recordset("vto")
                              data_inf.Recordset("costo") = data_lin.Recordset("margen_prd")
                              data_inf.Recordset("nro_superv") = data_lin.Recordset("pre_prod")
                              data_inf.Recordset("zona") = data_lin.Recordset("zona")
                              If IsNull(data_lin.Recordset("pendiente")) = False Then
                                 If data_lin.Recordset("pendiente") = "N" Or data_lin.Recordset("pendiente") = "C" Or data_lin.Recordset("pendiente") = "R" Then
                                    data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin") * -1
                                 Else
                                    data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                                 End If
                              Else
                                 data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                              End If
                              data_inf.Recordset.Update
                              Xtotmedent = Xtotmedent + 1
                           Else
                              If data_lin.Recordset("hora") >= mhd.Text Then
                                 If data_lin.Recordset("hora") <= mhh.Text Then
                                   data_inf.Recordset.AddNew
                                   data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
                                   data_inf.Recordset("hora") = data_lin.Recordset("hora")
                                   data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                                   data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                                   data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
                                   data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
                                   data_inf.Recordset("base") = data_lin.Recordset("base")
                                   data_inf.Recordset("factura") = data_lin.Recordset("factura")
                                   data_inf.Recordset("nom_medic") = data_lin.Recordset("nom_medic")
                                   data_inf.Recordset("realizada") = data_lin.Recordset("cl_fnac")
                                   If IsNull(data_lin.Recordset("ced_socio")) = False Then
                                      If IsNull(data_lin.Recordset("fact")) = False Then
                                         data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-" & Trim(str(data_lin.Recordset("fact")))
                                      Else
                                         data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-"
                                      End If
                                   Else
                                      data_inf.Recordset("nom_med_s") = "0"
                                   End If
                                   data_inf.Recordset("nro_med_s") = Xtotmedsol
                                   data_inf.Recordset("vto") = data_lin.Recordset("vto")
                                   data_inf.Recordset("costo") = data_lin.Recordset("margen_prd")
                                   data_inf.Recordset("nro_superv") = data_lin.Recordset("pre_prod")
                                   data_inf.Recordset("zona") = data_lin.Recordset("zona")
                                   If IsNull(data_lin.Recordset("pendiente")) = False Then
                                      If data_lin.Recordset("pendiente") = "N" Or data_lin.Recordset("pendiente") = "C" Or data_lin.Recordset("pendiente") = "R" Then
                                         data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin") * -1
                                      Else
                                         data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                                      End If
                                   Else
                                      data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                                   End If
                                   data_inf.Recordset.Update
                                   Xtotmedent = Xtotmedent + 1
                                 
                                 End If
                              End If
                           End If
                        End If
                     End If
                  End If
               End If
               If Option2.Value = True Then
                  If IsNull(data_lin.Recordset("dias")) = False Then
                     If data_lin.Recordset("dias") = 0 Or _
                        data_lin.Recordset("dias") = 2 Or _
                        data_lin.Recordset("dias") = 4 Then
                        If data_lin.Recordset("tot_lin") >= 0 Then
                           If mhd.Text = "__:__" And mhh.Text = "__:__" Then
                              data_inf.Recordset.AddNew
                              data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
                              data_inf.Recordset("hora") = data_lin.Recordset("hora")
                              data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                              data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                              data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
                              data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
                              data_inf.Recordset("base") = data_lin.Recordset("base")
                              data_inf.Recordset("factura") = data_lin.Recordset("factura")
                              data_inf.Recordset("nom_medic") = data_lin.Recordset("nom_medic")
                              If IsNull(data_lin.Recordset("ced_socio")) = False Then
                                 If IsNull(data_lin.Recordset("fact")) = False Then
                                    data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-" & Trim(str(data_lin.Recordset("fact")))
                                 Else
                                    data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-"
                                 End If
                              Else
                                 data_inf.Recordset("nom_med_s") = "0"
                              End If
                              data_inf.Recordset("nro_med_s") = Xtotmedsol
                              data_inf.Recordset("vto") = data_lin.Recordset("vto")
                              data_inf.Recordset("costo") = data_lin.Recordset("margen_prd")
                              data_inf.Recordset("nro_superv") = data_lin.Recordset("pre_prod")
                              data_inf.Recordset("zona") = data_lin.Recordset("zona")
                              If IsNull(data_lin.Recordset("pendiente")) = False Then
                                 If data_lin.Recordset("pendiente") = "N" Or data_lin.Recordset("pendiente") = "C" Or data_lin.Recordset("pendiente") = "R" Then
                                    data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin") * -1
                                 Else
                                    data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                                 End If
                              Else
                                 data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                              End If
                              data_inf.Recordset.Update
                              Xtotmedent = Xtotmedent + 1
                           Else
                              If data_lin.Recordset("hora") >= mhd.Text Then
                                 If data_lin.Recordset("hora") <= mhh.Text Then
                                   data_inf.Recordset.AddNew
                                   data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
                                   data_inf.Recordset("hora") = data_lin.Recordset("hora")
                                   data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                                   data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                                   data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
                                   data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
                                   data_inf.Recordset("base") = data_lin.Recordset("base")
                                   data_inf.Recordset("factura") = data_lin.Recordset("factura")
                                   data_inf.Recordset("nom_medic") = data_lin.Recordset("nom_medic")
                                   If IsNull(data_lin.Recordset("ced_socio")) = False Then
                                      If IsNull(data_lin.Recordset("fact")) = False Then
                                         data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-" & Trim(str(data_lin.Recordset("fact")))
                                      Else
                                         data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-"
                                      End If
                                   Else
                                      data_inf.Recordset("nom_med_s") = "0"
                                   End If
                                   data_inf.Recordset("nro_med_s") = Xtotmedsol
                                   data_inf.Recordset("vto") = data_lin.Recordset("vto")
                                   data_inf.Recordset("costo") = data_lin.Recordset("margen_prd")
                                   data_inf.Recordset("nro_superv") = data_lin.Recordset("pre_prod")
                                   data_inf.Recordset("zona") = data_lin.Recordset("zona")
                                   If IsNull(data_lin.Recordset("pendiente")) = False Then
                                      If data_lin.Recordset("pendiente") = "N" Or data_lin.Recordset("pendiente") = "C" Or data_lin.Recordset("pendiente") = "R" Then
                                         data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin") * -1
                                      Else
                                         data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                                      End If
                                   Else
                                      data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                                   End If
                                   data_inf.Recordset.Update
                                   Xtotmedent = Xtotmedent + 1
                                 End If
                              End If
                              
                           End If
                        End If
                     End If
                  End If
               End If
               If Option8.Value = True Then
                  If IsNull(data_lin.Recordset("dias")) = False Then
                     If data_lin.Recordset("dias") = 7 Then
                        If data_lin.Recordset("tot_lin") >= 0 Then
                           If mhd.Text = "__:__" And mhh.Text = "__:__" Then
                               data_inf.Recordset.AddNew
                               data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
                               data_inf.Recordset("hora") = data_lin.Recordset("hora")
                               data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                               data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                               data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
                               data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
                               data_inf.Recordset("base") = data_lin.Recordset("base")
                               data_inf.Recordset("factura") = data_lin.Recordset("factura")
                               data_inf.Recordset("nom_medic") = data_lin.Recordset("nom_medic")
                               If IsNull(data_lin.Recordset("ced_socio")) = False Then
                                  If IsNull(data_lin.Recordset("fact")) = False Then
                                     data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-" & Trim(str(data_lin.Recordset("fact")))
                                  Else
                                     data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-"
                                  End If
                               Else
                                  data_inf.Recordset("nom_med_s") = "0"
                               End If
                               data_inf.Recordset("nro_med_s") = Xtotmedsol
                               data_inf.Recordset("vto") = data_lin.Recordset("vto")
                               data_inf.Recordset("costo") = data_lin.Recordset("margen_prd")
                               data_inf.Recordset("nro_superv") = data_lin.Recordset("pre_prod")
                               data_inf.Recordset("zona") = data_lin.Recordset("zona")
                               If IsNull(data_lin.Recordset("pendiente")) = False Then
                                  If data_lin.Recordset("pendiente") = "N" Or data_lin.Recordset("pendiente") = "C" Or data_lin.Recordset("pendiente") = "R" Then
                                     data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin") * -1
                                  Else
                                     data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                                  End If
                               Else
                                  data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                               End If
                               data_inf.Recordset.Update
                               Xtotmedent = Xtotmedent + 1
                           Else
                               If data_lin.Recordset("hora") >= mhd.Text Then
                                  If data_lin.Recordset("hora") <= mhh.Text Then
                                       data_inf.Recordset.AddNew
                                       data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
                                       data_inf.Recordset("hora") = data_lin.Recordset("hora")
                                       data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                                       data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                                       data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
                                       data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
                                       data_inf.Recordset("base") = data_lin.Recordset("base")
                                       data_inf.Recordset("factura") = data_lin.Recordset("factura")
                                       data_inf.Recordset("nom_medic") = data_lin.Recordset("nom_medic")
                                       If IsNull(data_lin.Recordset("ced_socio")) = False Then
                                          If IsNull(data_lin.Recordset("fact")) = False Then
                                             data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-" & Trim(str(data_lin.Recordset("fact")))
                                          Else
                                             data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-"
                                          End If
                                       Else
                                          data_inf.Recordset("nom_med_s") = "0"
                                       End If
                                       data_inf.Recordset("nro_med_s") = Xtotmedsol
                                       data_inf.Recordset("vto") = data_lin.Recordset("vto")
                                       data_inf.Recordset("costo") = data_lin.Recordset("margen_prd")
                                       data_inf.Recordset("nro_superv") = data_lin.Recordset("pre_prod")
                                       data_inf.Recordset("zona") = data_lin.Recordset("zona")
                                       If IsNull(data_lin.Recordset("pendiente")) = False Then
                                          If data_lin.Recordset("pendiente") = "N" Or data_lin.Recordset("pendiente") = "C" Or data_lin.Recordset("pendiente") = "R" Then
                                             data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin") * -1
                                          Else
                                             data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                                          End If
                                       Else
                                          data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                                       End If
                                       data_inf.Recordset.Update
                                       Xtotmedent = Xtotmedent + 1
                           
                                  End If
                               End If
                           End If
                        End If
                     End If
                  End If
               End If
               If Option5.Value = True Then
                  If IsNull(data_lin.Recordset("dias")) = False Then
                     If data_lin.Recordset("dias") = 5 Or _
                        data_lin.Recordset("dias") = 3 Then
                        If data_lin.Recordset("tot_lin") >= 0 Then
                           If mhd.Text = "__:__" And mhh.Text = "__:__" Then
                               data_inf.Recordset.AddNew
                               data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
                               data_inf.Recordset("hora") = data_lin.Recordset("hora")
                               data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                               data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                               data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
                               data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
                               data_inf.Recordset("base") = data_lin.Recordset("base")
                               data_inf.Recordset("factura") = data_lin.Recordset("factura")
                               data_inf.Recordset("nom_medic") = data_lin.Recordset("nom_medic")
                               If IsNull(data_lin.Recordset("ced_socio")) = False Then
                                  If IsNull(data_lin.Recordset("fact")) = False Then
                                     data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-" & Trim(str(data_lin.Recordset("fact")))
                                  Else
                                     data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-"
                                  End If
                               Else
                                  data_inf.Recordset("nom_med_s") = "0"
                               End If
                               data_inf.Recordset("nro_med_s") = Xtotmedsol
                               data_inf.Recordset("vto") = data_lin.Recordset("vto")
                               data_inf.Recordset("costo") = data_lin.Recordset("margen_prd")
                               data_inf.Recordset("nro_superv") = data_lin.Recordset("pre_prod")
                               data_inf.Recordset("zona") = data_lin.Recordset("zona")
                               If IsNull(data_lin.Recordset("pendiente")) = False Then
                                  If data_lin.Recordset("pendiente") = "N" Or data_lin.Recordset("pendiente") = "C" Or data_lin.Recordset("pendiente") = "R" Then
                                     data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin") * -1
                                  Else
                                     data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                                  End If
                               Else
                                  data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                               End If
                               data_inf.Recordset.Update
                               Xtotmedent = Xtotmedent + 1
                           Else
                              If data_lin.Recordset("hora") >= mhd.Text Then
                                 If data_lin.Recordset("hora") <= mhh.Text Then
                                   data_inf.Recordset.AddNew
                                   data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
                                   data_inf.Recordset("hora") = data_lin.Recordset("hora")
                                   data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                                   data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                                   data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
                                   data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
                                   data_inf.Recordset("base") = data_lin.Recordset("base")
                                   data_inf.Recordset("factura") = data_lin.Recordset("factura")
                                   data_inf.Recordset("nom_medic") = data_lin.Recordset("nom_medic")
                                   If IsNull(data_lin.Recordset("ced_socio")) = False Then
                                      If IsNull(data_lin.Recordset("fact")) = False Then
                                         data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-" & Trim(str(data_lin.Recordset("fact")))
                                      Else
                                         data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-"
                                      End If
                                   Else
                                      data_inf.Recordset("nom_med_s") = "0"
                                   End If
                                   data_inf.Recordset("nro_med_s") = Xtotmedsol
                                   data_inf.Recordset("vto") = data_lin.Recordset("vto")
                                   data_inf.Recordset("costo") = data_lin.Recordset("margen_prd")
                                   data_inf.Recordset("nro_superv") = data_lin.Recordset("pre_prod")
                                   data_inf.Recordset("zona") = data_lin.Recordset("zona")
                                   If IsNull(data_lin.Recordset("pendiente")) = False Then
                                      If data_lin.Recordset("pendiente") = "N" Or data_lin.Recordset("pendiente") = "C" Or data_lin.Recordset("pendiente") = "R" Then
                                         data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin") * -1
                                      Else
                                         data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                                      End If
                                   Else
                                      data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                                   End If
                                   data_inf.Recordset.Update
                                   Xtotmedent = Xtotmedent + 1
                           
                                 End If
                              End If
                           End If
                        End If
                     End If
                  End If
               End If
               If Option3.Value = True Then
                  If IsNull(data_lin.Recordset("dias")) = False Then
                     If data_lin.Recordset("dias") = 2 Then
                        If data_lin.Recordset("tot_lin") >= 0 Then
                           If mhd.Text = "__:__" And mhh.Text = "__:__" Then
                               data_inf.Recordset.AddNew
                               data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
                               data_inf.Recordset("hora") = data_lin.Recordset("hora")
                               data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                               data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                               data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
                               data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
                               data_inf.Recordset("base") = data_lin.Recordset("base")
                               data_inf.Recordset("factura") = data_lin.Recordset("factura")
                               data_inf.Recordset("nom_medic") = data_lin.Recordset("nom_medic")
                               If IsNull(data_lin.Recordset("ced_socio")) = False Then
                                  If IsNull(data_lin.Recordset("fact")) = False Then
                                     data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-" & Trim(str(data_lin.Recordset("fact")))
                                  Else
                                     data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-"
                                  End If
                               Else
                                  data_inf.Recordset("nom_med_s") = "0"
                               End If
                               data_inf.Recordset("nro_med_s") = Xtotmedsol
                               data_inf.Recordset("vto") = data_lin.Recordset("vto")
                               data_inf.Recordset("costo") = data_lin.Recordset("margen_prd")
                               data_inf.Recordset("nro_superv") = data_lin.Recordset("pre_prod")
                               data_inf.Recordset("zona") = data_lin.Recordset("zona")
                               If IsNull(data_lin.Recordset("pendiente")) = False Then
                                  If data_lin.Recordset("pendiente") = "N" Or data_lin.Recordset("pendiente") = "C" Or data_lin.Recordset("pendiente") = "R" Then
                                     data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin") * -1
                                  Else
                                     data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                                  End If
                               Else
                                  data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                               End If
                               data_inf.Recordset.Update
                               Xtotmedent = Xtotmedent + 1
                           Else
                              If data_lin.Recordset("hora") >= mhd.Text Then
                                 If data_lin.Recordset("hora") <= mhh.Text Then
                                   data_inf.Recordset.AddNew
                                   data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
                                   data_inf.Recordset("hora") = data_lin.Recordset("hora")
                                   data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                                   data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                                   data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
                                   data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
                                   data_inf.Recordset("base") = data_lin.Recordset("base")
                                   data_inf.Recordset("factura") = data_lin.Recordset("factura")
                                   data_inf.Recordset("nom_medic") = data_lin.Recordset("nom_medic")
                                   If IsNull(data_lin.Recordset("ced_socio")) = False Then
                                      If IsNull(data_lin.Recordset("fact")) = False Then
                                         data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-" & Trim(str(data_lin.Recordset("fact")))
                                      Else
                                         data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-"
                                      End If
                                   Else
                                      data_inf.Recordset("nom_med_s") = "0"
                                   End If
                                   data_inf.Recordset("nro_med_s") = Xtotmedsol
                                   data_inf.Recordset("vto") = data_lin.Recordset("vto")
                                   data_inf.Recordset("costo") = data_lin.Recordset("margen_prd")
                                   data_inf.Recordset("nro_superv") = data_lin.Recordset("pre_prod")
                                   data_inf.Recordset("zona") = data_lin.Recordset("zona")
                                   If IsNull(data_lin.Recordset("pendiente")) = False Then
                                      If data_lin.Recordset("pendiente") = "N" Or data_lin.Recordset("pendiente") = "C" Or data_lin.Recordset("pendiente") = "R" Then
                                         data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin") * -1
                                      Else
                                         data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                                      End If
                                   Else
                                      data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                                   End If
                                   data_inf.Recordset.Update
                                   Xtotmedent = Xtotmedent + 1
                           
                                 End If
                              End If
                           End If
                        End If
                       End If
                  End If
               End If
               If Option4.Value = True Then
                  If IsNull(data_lin.Recordset("dias")) = False Then
                     If data_lin.Recordset("dias") = 6 Then
                        If mhd.Text = "__:__" And mhh.Text = "__:__" Then
                           data_inf.Recordset.AddNew
                           data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
                           data_inf.Recordset("hora") = data_lin.Recordset("hora")
                           data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                           data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                           data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
                           data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
                           data_inf.Recordset("base") = data_lin.Recordset("base")
                           data_inf.Recordset("factura") = data_lin.Recordset("factura")
                           data_inf.Recordset("nom_medic") = data_lin.Recordset("nom_medic")
                           If IsNull(data_lin.Recordset("ced_socio")) = False Then
                              If IsNull(data_lin.Recordset("fact")) = False Then
                                 data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-" & Trim(str(data_lin.Recordset("fact")))
                              Else
                                 data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-"
                              End If
                           Else
                              data_inf.Recordset("nom_med_s") = "0"
                           End If
                           data_inf.Recordset("nro_med_s") = Xtotmedsol
                           data_inf.Recordset("vto") = data_lin.Recordset("vto")
                           data_inf.Recordset("costo") = data_lin.Recordset("margen_prd")
                           data_inf.Recordset("nro_superv") = data_lin.Recordset("pre_prod")
                           data_inf.Recordset("zona") = data_lin.Recordset("zona")
                           If IsNull(data_lin.Recordset("pendiente")) = False Then
                              If data_lin.Recordset("pendiente") = "N" Or data_lin.Recordset("pendiente") = "C" Or data_lin.Recordset("pendiente") = "R" Then
                                 data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin") * -1
                              Else
                                 data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                              End If
                           Else
                              data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                           End If
                           data_inf.Recordset.Update
                           Xtotmedent = Xtotmedent + 1
                        Else
                           If data_lin.Recordset("hora") >= mhd.Text Then
                              If data_lin.Recordset("hora") <= mhh.Text Then
                                   data_inf.Recordset.AddNew
                                   data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
                                   data_inf.Recordset("hora") = data_lin.Recordset("hora")
                                   data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                                   data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                                   data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
                                   data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
                                   data_inf.Recordset("base") = data_lin.Recordset("base")
                                   data_inf.Recordset("factura") = data_lin.Recordset("factura")
                                   data_inf.Recordset("nom_medic") = data_lin.Recordset("nom_medic")
                                   If IsNull(data_lin.Recordset("ced_socio")) = False Then
                                      If IsNull(data_lin.Recordset("fact")) = False Then
                                         data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-" & Trim(str(data_lin.Recordset("fact")))
                                      Else
                                         data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-"
                                      End If
                                   Else
                                      data_inf.Recordset("nom_med_s") = "0"
                                   End If
                                   data_inf.Recordset("nro_med_s") = Xtotmedsol
                                   data_inf.Recordset("vto") = data_lin.Recordset("vto")
                                   data_inf.Recordset("costo") = data_lin.Recordset("margen_prd")
                                   data_inf.Recordset("nro_superv") = data_lin.Recordset("pre_prod")
                                   data_inf.Recordset("zona") = data_lin.Recordset("zona")
                                   If IsNull(data_lin.Recordset("pendiente")) = False Then
                                      If data_lin.Recordset("pendiente") = "N" Or data_lin.Recordset("pendiente") = "C" Or data_lin.Recordset("pendiente") = "R" Then
                                         data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin") * -1
                                      Else
                                         data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                                      End If
                                   Else
                                      data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                                   End If
                                   data_inf.Recordset.Update
                                   Xtotmedent = Xtotmedent + 1
                        
                              End If
                           End If
                        End If
                     End If
                  End If
               End If
               If Option9.Value = True Then
                  If IsNull(data_lin.Recordset("dias")) = False Then
                     If data_lin.Recordset("dias") = 1 Or _
                        data_lin.Recordset("dias") = 0 Then
                        If data_lin.Recordset("tot_lin") < 0 Then
                           If mhd.Text = "__:__" And mhh.Text = "__:__" Then
                               data_inf.Recordset.AddNew
                               data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
                               data_inf.Recordset("hora") = data_lin.Recordset("hora")
                               data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                               data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                               data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
                               data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
                               data_inf.Recordset("base") = data_lin.Recordset("base")
                               data_inf.Recordset("factura") = data_lin.Recordset("factura")
                               data_inf.Recordset("nom_medic") = data_lin.Recordset("nom_medic")
                               If IsNull(data_lin.Recordset("ced_socio")) = False Then
                                  If IsNull(data_lin.Recordset("fact")) = False Then
                                     data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-" & Trim(str(data_lin.Recordset("fact")))
                                  Else
                                     data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-"
                                  End If
                               Else
                                  data_inf.Recordset("nom_med_s") = "0"
                               End If
                               data_inf.Recordset("nro_med_s") = Xtotmedsol
                               data_inf.Recordset("vto") = data_lin.Recordset("vto")
                               data_inf.Recordset("costo") = data_lin.Recordset("margen_prd")
                               data_inf.Recordset("nro_superv") = data_lin.Recordset("pre_prod")
                               data_inf.Recordset("zona") = data_lin.Recordset("zona")
                               If IsNull(data_lin.Recordset("pendiente")) = False Then
                                  If data_lin.Recordset("pendiente") = "N" Or data_lin.Recordset("pendiente") = "C" Or data_lin.Recordset("pendiente") = "R" Then
                                     data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin") * -1
                                  Else
                                     data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                                  End If
                               Else
                                  data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                               End If
                               data_inf.Recordset.Update
                               Xtotmedent = Xtotmedent + 1
                           Else
                               If data_lin.Recordset("hora") >= mhd.Text Then
                                  If data_lin.Recordset("hora") <= mhh.Text Then
                                       data_inf.Recordset.AddNew
                                       data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
                                       data_inf.Recordset("hora") = data_lin.Recordset("hora")
                                       data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                                       data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                                       data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
                                       data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
                                       data_inf.Recordset("base") = data_lin.Recordset("base")
                                       data_inf.Recordset("factura") = data_lin.Recordset("factura")
                                       data_inf.Recordset("nom_medic") = data_lin.Recordset("nom_medic")
                                       If IsNull(data_lin.Recordset("ced_socio")) = False Then
                                          If IsNull(data_lin.Recordset("fact")) = False Then
                                             data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-" & Trim(str(data_lin.Recordset("fact")))
                                          Else
                                             data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-"
                                          End If
                                       Else
                                          data_inf.Recordset("nom_med_s") = "0"
                                       End If
                                       data_inf.Recordset("nro_med_s") = Xtotmedsol
                                       data_inf.Recordset("vto") = data_lin.Recordset("vto")
                                       data_inf.Recordset("costo") = data_lin.Recordset("margen_prd")
                                       data_inf.Recordset("nro_superv") = data_lin.Recordset("pre_prod")
                                       data_inf.Recordset("zona") = data_lin.Recordset("zona")
                                       If IsNull(data_lin.Recordset("pendiente")) = False Then
                                          If data_lin.Recordset("pendiente") = "N" Or data_lin.Recordset("pendiente") = "C" Or data_lin.Recordset("pendiente") = "R" Then
                                             data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin") * -1
                                          Else
                                             data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                                          End If
                                       Else
                                          data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                                       End If
                                       data_inf.Recordset.Update
                                       Xtotmedent = Xtotmedent + 1
                           
                                  End If
                               End If
                           End If
                        End If
                     End If
                  End If
               End If
               If Option10.Value = True Then
                  If mhd.Text = "__:__" And mhh.Text = "__:__" Then
                       data_inf.Recordset.AddNew
                       data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
                       data_inf.Recordset("hora") = data_lin.Recordset("hora")
                       data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                       data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                       data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
                       data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
                       data_inf.Recordset("base") = data_lin.Recordset("base")
                       data_inf.Recordset("factura") = data_lin.Recordset("factura")
                       data_inf.Recordset("nom_medic") = data_lin.Recordset("nom_medic")
                       If IsNull(data_lin.Recordset("ced_socio")) = False Then
                          If IsNull(data_lin.Recordset("fact")) = False Then
                             data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-" & Trim(str(data_lin.Recordset("fact")))
                          Else
                             data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-"
                          End If
                       Else
                          data_inf.Recordset("nom_med_s") = "0"
                       End If
                       data_inf.Recordset("nro_med_s") = Xtotmedsol
                       data_inf.Recordset("vto") = data_lin.Recordset("vto")
                       data_inf.Recordset("costo") = data_lin.Recordset("margen_prd")
                       data_inf.Recordset("nro_superv") = data_lin.Recordset("pre_prod")
                       data_inf.Recordset("zona") = data_lin.Recordset("zona")
                       If IsNull(data_lin.Recordset("pendiente")) = False Then
                          If data_lin.Recordset("pendiente") = "N" Or data_lin.Recordset("pendiente") = "C" Or data_lin.Recordset("pendiente") = "R" Then
                             data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin") * -1
                          Else
                             data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                          End If
                       Else
                          data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                       End If
                       data_inf.Recordset.Update
                       Xtotmedent = Xtotmedent + 1
                  Else
                       If data_lin.Recordset("hora") >= mhd.Text Then
                          If data_lin.Recordset("hora") <= mhh.Text Then
                               data_inf.Recordset.AddNew
                               data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
                               data_inf.Recordset("hora") = data_lin.Recordset("hora")
                               data_inf.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                               data_inf.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                               data_inf.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
                               data_inf.Recordset("nom_prod") = data_lin.Recordset("nom_prod")
                               data_inf.Recordset("base") = data_lin.Recordset("base")
                               data_inf.Recordset("factura") = data_lin.Recordset("factura")
                               data_inf.Recordset("nom_medic") = data_lin.Recordset("nom_medic")
                               If IsNull(data_lin.Recordset("ced_socio")) = False Then
                                  If IsNull(data_lin.Recordset("fact")) = False Then
                                     data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-" & Trim(str(data_lin.Recordset("fact")))
                                  Else
                                     data_inf.Recordset("nom_med_s") = Trim(str(data_lin.Recordset("ced_socio"))) & "-"
                                  End If
                               Else
                                  data_inf.Recordset("nom_med_s") = "0"
                               End If
                               data_inf.Recordset("nro_med_s") = Xtotmedsol
                               data_inf.Recordset("vto") = data_lin.Recordset("vto")
                               data_inf.Recordset("costo") = data_lin.Recordset("margen_prd")
                               data_inf.Recordset("nro_superv") = data_lin.Recordset("pre_prod")
                               data_inf.Recordset("zona") = data_lin.Recordset("zona")
                               If IsNull(data_lin.Recordset("pendiente")) = False Then
                                  If data_lin.Recordset("pendiente") = "N" Or data_lin.Recordset("pendiente") = "C" Or data_lin.Recordset("pendiente") = "R" Then
                                     data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin") * -1
                                  Else
                                     data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                                  End If
                               Else
                                  data_inf.Recordset("tot_lin") = data_lin.Recordset("tot_lin")
                               End If
                               data_inf.Recordset.Update
                               Xtotmedent = Xtotmedent + 1
                          End If
                       End If
                  End If
               End If
               data_lin.Recordset.MoveNext
            Loop
         End If
         Dim Xcannn As Integer
         If Text1.Text <> "" Then
            If data_inf.Recordset.RecordCount > 0 Then
               data_inf.RecordSource = "select * from infvtas order by cod_cli"
               data_inf.Refresh
               data_inf.Recordset.MoveFirst
               Do While Not data_inf.Recordset.EOF
                  Data1.RecordSource = "Select * from infvtas where cod_cli =" & data_inf.Recordset("cod_cli")
                  Data1.Refresh
                  Data1.Recordset.MoveFirst
                  Do While Not Data1.Recordset.EOF
                     Xcannn = Xcannn + 1
                     Data1.Recordset.MoveNext
                  Loop
                  If Xcannn >= Val(Text1.Text) Then
                  Else
                     data_inf.Recordset.Edit
                     data_inf.Recordset("base") = 66
                     data_inf.Recordset.Update
                  End If
                  Xcannn = 0
                  data_inf.Recordset.MoveNext
               Loop
               data_inf.Refresh
               MiBaseact.Execute "Delete * from infvtas where base =" & 66
               data_inf.RecordSource = "Select * from infvtas"
               data_inf.Refresh
            End If
         End If
         If t_edd.Text <> "" Then
            If t_edh.Text <> "" Then
               If data_inf.Recordset.RecordCount > 0 Then
                  data_inf.Recordset.MoveFirst
                  Do While Not data_inf.Recordset.EOF
                     If IsNull(data_inf.Recordset("realizada")) = False Then
                        Xed = Date - data_inf.Recordset("realizada")
                        Xed = Xed / 365
                        If Xed >= t_edd.Text Then
                           If Xed <= t_edh.Text Then
                           Else
                              data_inf.Recordset.Delete
                           End If
                        Else
                            data_inf.Recordset.Delete
                        End If
                     Else
                        data_inf.Recordset.Delete
                     End If
                     data_inf.Recordset.MoveNext
                  Loop
                  data_inf.RecordSource = "Select * from infvtas"
                  data_inf.Refresh
               End If
            End If
         End If
         If Xtotmedsol <> 0 Then
            Xelporce = Xtotmedent / Xtotmedsol
            Xelporce = Xelporce * 100
         Else
            Xelporce = 0
         End If
         If Option13.Value = True Then
            Command3_Click
            data_inf.RecordSource = "select * from infcli"
            data_inf.Refresh
            cr1.ReportFileName = App.path & "\infmedporsocio.rpt"
            cr1.ReportTitle = "Informe de socios solicitaron medicación desde:" & mfd.Text & " hasta:" & mfh.Text
            cr1.Action = 1
         Else
            If data_inf.Recordset.RecordCount > 0 Then
               If Option1.Value = True Then
                   data_inf.Recordset.MoveFirst
                   Do While Not data_inf.Recordset.EOF
                      data_inf.Recordset.Edit
                      data_inf.Recordset("nro_med_a") = Xelporce
                      If Xelporce >= 90 Then
                         data_inf.Recordset("nom_med_a") = "CUMPLIDO"
                      Else
                         data_inf.Recordset("nom_med_a") = "NO CUMPLIDO"
                      End If
                      data_inf.Recordset.Update
                      data_inf.Recordset.MoveNext
                   Loop
                   frm_infctrolfar.MousePointer = 0
                   data_inf.RecordSource = "Select * from infvtas"
                   data_inf.Refresh
                   If Option6.Value = True Then
                      cr1.ReportFileName = App.path & "\infmedentn.rpt"
                   Else
                      If t_edd.Text <> "" Then
                         cr1.ReportFileName = App.path & "\infmedentded.rpt"
                      Else
                         If Text1.Text <> "" Then
                            cr1.ReportFileName = App.path & "\infmedentd2.rpt"
                         Else
                            cr1.ReportFileName = App.path & "\infmedentd.rpt"
                         End If
                      End If
                   End If
                   cr1.ReportTitle = "INFORME DE MEDICACION ENTREGADA DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                   cr1.Action = 1
               End If
               If Option2.Value = True Then
                   frm_infctrolfar.MousePointer = 0
                   data_inf.RecordSource = "Select * from infvtas"
                   data_inf.Refresh
                   If Option6.Value = True Then
                      cr1.ReportFileName = App.path & "\infmednoenn.rpt"
                   Else
                      If t_edd.Text <> "" Then
                         cr1.ReportFileName = App.path & "\infmedentded.rpt"
                      Else
                         cr1.ReportFileName = App.path & "\infmednoend.rpt"
                      End If
                   End If
                   cr1.ReportTitle = "INFORME DE MEDICACION SIN ENTREGAR DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                   cr1.Action = 1
               
               End If
               If Option5.Value = True Then
                   frm_infctrolfar.MousePointer = 0
                   data_inf.RecordSource = "Select * from infvtas"
                   data_inf.Refresh
                   If Option6.Value = True Then
                      cr1.ReportFileName = App.path & "\infmedenvbn.rpt"
                   Else
                      If t_edd.Text <> "" Then
                         cr1.ReportFileName = App.path & "\infmedentded.rpt"
                      Else
                         cr1.ReportFileName = App.path & "\infmedenvbd.rpt"
                      End If
                   End If
                   cr1.ReportTitle = "INFORME DE MEDICACION ENVIADA A BASES DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                   cr1.Action = 1
               
               End If
               If Option3.Value = True Then
                   frm_infctrolfar.MousePointer = 0
                   data_inf.RecordSource = "Select * from infvtas"
                   data_inf.Refresh
                   If Option6.Value = True Then
                      cr1.ReportFileName = App.path & "\infmedsolfn.rpt"
                   Else
                      If t_edd.Text <> "" Then
                         cr1.ReportFileName = App.path & "\infmedentded.rpt"
                      Else
                         cr1.ReportFileName = App.path & "\infmedsolfd.rpt"
                      End If
                   End If
                   cr1.ReportTitle = "MEDICACION SOLICITADA A FARMACIA CENTRAL DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                   cr1.Action = 1
               
               End If
               If Option4.Value = True Then
                   frm_infctrolfar.MousePointer = 0
                   data_inf.RecordSource = "Select * from infvtas"
                   data_inf.Refresh
                   If Option6.Value = True Then
                      cr1.ReportFileName = App.path & "\infmedsolfn.rpt"
                   Else
                      If t_edd.Text <> "" Then
                         cr1.ReportFileName = App.path & "\infmedentded.rpt"
                      Else
                         cr1.ReportFileName = App.path & "\infmedsolfd.rpt"
                      End If
                   End If
                   cr1.ReportTitle = "MEDICACION QUE NO SE RETIRO (DEVOLUCIONES) DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                   cr1.Action = 1
               
               End If
               If Option8.Value = True Then
                   frm_infctrolfar.MousePointer = 0
                   data_inf.RecordSource = "Select * from infvtas"
                   data_inf.Refresh
                   If Option6.Value = True Then
                      cr1.ReportFileName = App.path & "\infmedsolfn.rpt"
                   Else
                      If t_edd.Text <> "" Then
                         cr1.ReportFileName = App.path & "\infmedentded.rpt"
                      Else
                         cr1.ReportFileName = App.path & "\infmedsolfd.rpt"
                      End If
                   End If
                   cr1.ReportTitle = "MEDICACION SOLICITADA A LAS MUTUALISTAS DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                   cr1.Action = 1
               End If
               If Option9.Value = True Then
                   frm_infctrolfar.MousePointer = 0
                   data_inf.RecordSource = "Select * from infvtas"
                   data_inf.Refresh
                   If Option6.Value = True Then
                      cr1.ReportFileName = App.path & "\infmedsolfn.rpt"
                   Else
                      If t_edd.Text <> "" Then
                         cr1.ReportFileName = App.path & "\infmedentded.rpt"
                      Else
                         cr1.ReportFileName = App.path & "\infmedsolfd.rpt"
                      End If
                   End If
                   cr1.ReportTitle = "MEDICACION FACTURADA COMO DEVOLUCION DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                   cr1.Action = 1
               
               End If
               If Option10.Value = True Then
                   frm_infctrolfar.MousePointer = 0
                   data_inf.RecordSource = "Select * from infvtas"
                   data_inf.Refresh
                   If t_edd.Text <> "" Then
                      cr1.ReportFileName = App.path & "\infmedentded.rpt"
                      cr1.ReportTitle = "MEDICACION FACTURADA DESDE: " & mfd.Text & " HASTA: " & mfh.Text & " HORA:" & mhd.Text & " A:" & mhh.Text
                      cr1.Action = 1
                   Else
                      cr1.ReportFileName = App.path & "\infmedfac.rpt"
                      cr1.ReportTitle = "MEDICACION FACTURADA DESDE: " & mfd.Text & " HASTA: " & mfh.Text & " HORA:" & mhd.Text & " A:" & mhh.Text
                      cr1.Action = 1
                   End If
               End If
               If Option11.Value = True Then
                   frm_infctrolfar.MousePointer = 0
                   data_inf.RecordSource = "Select * from infvtas"
                   data_inf.Refresh
                   cr1.ReportFileName = App.path & "\infmedprox.rpt"
                   cr1.ReportTitle = "MEDICACION A SER RETIRADA ENTRE: " & mfd.Text & " HASTA: " & mfh.Text
                   cr1.Action = 1
               End If
               If Option12.Value = True Then
                   frm_infctrolfar.MousePointer = 0
                   data_inf.RecordSource = "Select * from infvtas"
                   data_inf.Refresh
                   cr1.ReportFileName = App.path & "\infmedbajmot.rpt"
                   cr1.ReportTitle = "RECETAS DE BAJA DE FECHAS: " & mfd.Text & " HASTA: " & mfh.Text
                   cr1.Action = 1
               End If
            Else
               MsgBox "No hay registros"
            End If
         End If
      Else
         If Option13.Value = True Then
            Command3_Click
            data_inf.RecordSource = "select * from infcli"
            data_inf.Refresh
            cr1.ReportFileName = App.path & "\infmedporsocio.rpt"
            cr1.ReportTitle = "Informe de socios solicitaron medicación desde:" & mfd.Text & " hasta:" & mfh.Text
            cr1.Action = 1
         Else
            MsgBox "No hay registros"
         End If
      End If
   End If
End If

Command1.Enabled = True
Command2.Enabled = True
frm_infctrolfar.MousePointer = 0

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()

Dim Xcantsocio, Xcantsocio2, Xcantsocio3, XSocio, Xdias As Integer
Dim Xmat As Double
Xcantsocio = 0
Xcantsocio2 = 0
Xcantsocio3 = 0
XSocio = 0
data_inf.DatabaseName = App.path & "\informes.mdb"
data_inf.RecordSource = "infcli"
data_inf.Refresh
If data_inf.Recordset.RecordCount > 0 Then
   data_inf.Recordset.MoveFirst
   Do While Not data_inf.Recordset.EOF
      data_inf.Recordset.Delete
      data_inf.Recordset.MoveNext
   Loop
End If

Data3.Connect = "odbc;dsn=sappnew;"

data_mdb.DatabaseName = App.path & "\informes.mdb"
data_mdb.RecordSource = "inflla"
data_mdb.Refresh
If data_mdb.Recordset.RecordCount > 0 Then
   data_mdb.Recordset.MoveFirst
   Do While Not data_mdb.Recordset.EOF
      data_mdb.Recordset.Delete
      data_mdb.Recordset.MoveNext
   Loop
End If

data_sql.Connect = "odbc;dsn=sappnew;"
'data_sql.RecordSource = "select from hc_prescrip where hc_fecha >=#" & Format("01/07/2020", "yyyy/mm/dd") & "# and hc_fecha <=#" & Format("06/07/2020", "yyyy/mm/dd") & "# and hc_tippresd in ('RECETA PACIENTE CRONICO') and hc_codmedica >" & 0 & " and motivo_cance is null order by hc_mat"
'data_sql.Refresh
If t_base.Text = "" Then
   If Check1.Value = 1 Then
        data_sql.RecordSource = "select hc_prescrip.hc_fecha,hc_prescrip.hc_nro,hc_prescrip.hc_mat,hc_prescrip.hc_tippresd," & _
        "hc_prescrip.hc_codmedica,hc_prescrip.motivo_cance,hc_prescrip.hc_indicanom,hc_prescrip.hc_descrip,hc_prescrip.hc_hastaf," & _
        "hc_prescrip.hc_comfec,hc_prescrip.hc_fecentrega,cabezal_hcdig.id,cabezal_hcdig.hc_base from hc_prescrip " & _
        "inner join cabezal_hcdig on hc_prescrip.hc_nro=cabezal_hcdig.id where " & _
        "hc_prescrip.hc_fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# and hc_prescrip.hc_fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "# and hc_prescrip.hc_tippresd in ('RECETA PACIENTE CRONICO') and hc_prescrip.hc_codmedica >" & 0 & " and hc_prescrip.motivo_cance is null and cabezal_hcdig.hc_base not in (5,117,217,203,206,301,202,306,208,310,213) order by hc_prescrip.hc_mat"
        data_sql.Refresh
   Else
        data_sql.RecordSource = "select hc_prescrip.hc_fecha,hc_prescrip.hc_nro,hc_prescrip.hc_mat,hc_prescrip.hc_tippresd," & _
        "hc_prescrip.hc_codmedica,hc_prescrip.motivo_cance,hc_prescrip.hc_indicanom,hc_prescrip.hc_descrip,hc_prescrip.hc_hastaf," & _
        "hc_prescrip.hc_comfec,hc_prescrip.hc_fecentrega,cabezal_hcdig.id,cabezal_hcdig.hc_base from hc_prescrip " & _
        "inner join cabezal_hcdig on hc_prescrip.hc_nro=cabezal_hcdig.id where " & _
        "hc_prescrip.hc_fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# and hc_prescrip.hc_fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "# and hc_prescrip.hc_tippresd in ('RECETA PACIENTE CRONICO') and hc_prescrip.hc_codmedica >" & 0 & " and hc_prescrip.motivo_cance is null order by hc_prescrip.hc_mat"
        data_sql.Refresh
   End If
Else
    data_sql.RecordSource = "select hc_prescrip.hc_fecha,hc_prescrip.hc_nro,hc_prescrip.hc_mat,hc_prescrip.hc_tippresd," & _
    "hc_prescrip.hc_codmedica,hc_prescrip.motivo_cance,hc_prescrip.hc_indicanom,hc_prescrip.hc_descrip,hc_prescrip.hc_hastaf," & _
    "hc_prescrip.hc_comfec,hc_prescrip.hc_fecentrega,cabezal_hcdig.id,cabezal_hcdig.hc_base from hc_prescrip " & _
    "inner join cabezal_hcdig on hc_prescrip.hc_nro=cabezal_hcdig.id where " & _
    "hc_prescrip.hc_fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# and hc_prescrip.hc_fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "# and hc_prescrip.hc_tippresd in ('RECETA PACIENTE CRONICO') and hc_prescrip.hc_codmedica >" & 0 & " and hc_prescrip.motivo_cance is null order by hc_prescrip.hc_mat"
    data_sql.Refresh
End If

If data_sql.Recordset.RecordCount > 0 Then
   data_sql.Recordset.MoveFirst
   Do While Not data_sql.Recordset.EOF
'      If Check1.Value = 1 Then
'         If data_sql.Recordset("hc_base") = 5 Or data_sql.Recordset("hc_base") = 117 Or _
'            data_sql.Recordset("hc_base") = 217 Or data_sql.Recordset("hc_base") = 203 Or _
'            data_sql.Recordset("hc_base") = 206 Or data_sql.Recordset("hc_base") = 301 Or _
'            data_sql.Recordset("hc_base") = 202 Or data_sql.Recordset("hc_base") = 306 Or _
'            data_sql.Recordset("hc_base") = 208 Or data_sql.Recordset("hc_base") = 310 Or _
'            data_sql.Recordset("hc_base") = 213 Then
'         Else
            data_mdb.Recordset.AddNew
            data_mdb.Recordset("matric") = data_sql.Recordset("hc_mat")
            data_mdb.Recordset("nombre") = data_sql.Recordset("hc_tippresd")
            data_mdb.Recordset("fecha") = data_sql.Recordset("hc_fecha")
            data_mdb.Recordset("fecpas") = data_sql.Recordset("hc_comfec")
            data_mdb.Recordset("fecsali") = data_sql.Recordset("hc_hastaf")
            data_mdb.Recordset("diag") = data_sql.Recordset("hc_indicanom")
            data_mdb.Recordset("fec_rea") = data_sql.Recordset("hc_fecentrega")
            data_mdb.Recordset("motcon") = data_sql.Recordset("hc_descrip")
            data_mdb.Recordset("movilpas") = data_sql.Recordset("hc_nro")
            data_mdb.Recordset.Update
'         End If
'      Else
'         data_mdb.Recordset.AddNew
'         data_mdb.Recordset("matric") = data_sql.Recordset("hc_mat")
'         data_mdb.Recordset("nombre") = data_sql.Recordset("hc_tippresd")
'         data_mdb.Recordset("fecha") = data_sql.Recordset("hc_fecha")
'         data_mdb.Recordset("fecpas") = data_sql.Recordset("hc_comfec")
'         data_mdb.Recordset("fecsali") = data_sql.Recordset("hc_hastaf")
'         data_mdb.Recordset("diag") = data_sql.Recordset("hc_indicanom")
'         data_mdb.Recordset("fec_rea") = data_sql.Recordset("hc_fecentrega")
'         data_mdb.Recordset("motcon") = data_sql.Recordset("hc_descrip")
'         data_mdb.Recordset("movilpas") = data_sql.Recordset("hc_nro")
'         data_mdb.Recordset.Update
'      End If
      data_sql.Recordset.MoveNext
   Loop
   Xcantsocio = 0
   Xcantsocio2 = 0
   Xcantsocio3 = 0
   data_sql.Recordset.MoveFirst
   Xmat = data_sql.Recordset("hc_mat")
   Do While Not data_sql.Recordset.EOF
      If Xmat = data_sql.Recordset("hc_mat") Then
         Xdias = DateDiff("d", data_sql.Recordset("hc_fecha"), data_sql.Recordset("hc_comfec"))
         If Xdias < 30 Then
            Xcantsocio = Xcantsocio + 1
         Else
            If Xdias >= 30 And Xdias < 60 Then
               Xcantsocio2 = Xcantsocio2 + 1
            Else
               Xcantsocio3 = Xcantsocio3 + 1
            End If
         End If
         Xmat = data_sql.Recordset("hc_mat")
         data_sql.Recordset.MoveNext
      Else
         XSocio = XSocio + 1
         data_sql.Recordset.MovePrevious
         data_inf.Recordset.AddNew
         data_inf.Recordset("cl_codigo") = data_sql.Recordset("hc_mat")
         If Xcantsocio > 0 Then
            data_inf.Recordset("cl_codced") = 1
         End If
         If Xcantsocio2 > 0 Then
            data_inf.Recordset("cl_cedula") = 1
         End If
         If Xcantsocio3 > 0 Then
            data_inf.Recordset("cl_nrovend") = 1
         End If
         data_inf.Recordset("cl_direcci") = data_sql.Recordset("hc_tippresd")
         data_inf.Recordset("cl_fecing") = data_sql.Recordset("hc_fecha")
         data_inf.Recordset("cl_fnac") = data_sql.Recordset("hc_comfec")
         data_inf.Recordset("cl_fultmov") = data_sql.Recordset("hc_hastaf")
         data_inf.Recordset("cl_apellid") = data_sql.Recordset("hc_indicanom")
         data_inf.Recordset("cl_fultvta") = data_sql.Recordset("hc_fecentrega")
         data_inf.Recordset("cl_grupo") = data_sql.Recordset("hc_base")
         data_inf.Recordset("cl_atrasop") = Xcantsocio
         data_inf.Recordset("cl_cantpag") = Xcantsocio2
         data_inf.Recordset("cl_cantdia") = Xcantsocio3
         Data3.RecordSource = "select * from clientes where cl_codigo =" & data_sql.Recordset("hc_mat")
         Data3.Refresh
         If Data3.Recordset.RecordCount > 0 Then
            data_inf.Recordset("cl_nombre") = Mid(Data3.Recordset("cl_apellid"), 1, 30)
            data_inf.Recordset("cl_codconv") = Data3.Recordset("cl_codconv")
         End If
         data_inf.Recordset.Update
         data_sql.Recordset.MoveNext
         Xmat = data_sql.Recordset("hc_mat")
         Xcantsocio = 0
         Xcantsocio2 = 0
         Xcantsocio3 = 0
      End If
   Loop
End If

If t_base = "" Then
   If Check1.Value = 1 Then
      data_sql.RecordSource = "select hc_prescrip.hc_fecha,hc_prescrip.hc_nro,hc_prescrip.hc_mat,hc_prescrip.hc_tippresd," & _
      "hc_prescrip.hc_codmedica,hc_prescrip.motivo_cance,hc_prescrip.hc_indicanom,hc_prescrip.hc_descrip,hc_prescrip.hc_hastaf," & _
      "hc_prescrip.hc_comfec,hc_prescrip.hc_fecentrega,cabezal_hcdig.id,cabezal_hcdig.hc_base from hc_prescrip " & _
      "inner join cabezal_hcdig on hc_prescrip.hc_nro=cabezal_hcdig.id where " & _
      "hc_prescrip.hc_fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# and hc_prescrip.hc_fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "# and hc_prescrip.hc_tippresd in ('MEDICACION') and hc_prescrip.hc_codmedica >" & 0 & " and hc_prescrip.motivo_cance is null and cabezal_hcdig.hc_base not in (5,117,217,203,206,301,202,306,208,310,213) order by hc_prescrip.hc_mat"
      data_sql.Refresh
   Else
      data_sql.RecordSource = "select hc_prescrip.hc_fecha,hc_prescrip.hc_nro,hc_prescrip.hc_mat,hc_prescrip.hc_tippresd," & _
      "hc_prescrip.hc_codmedica,hc_prescrip.motivo_cance,hc_prescrip.hc_indicanom,hc_prescrip.hc_descrip,hc_prescrip.hc_hastaf," & _
      "hc_prescrip.hc_comfec,hc_prescrip.hc_fecentrega,cabezal_hcdig.id,cabezal_hcdig.hc_base from hc_prescrip " & _
      "inner join cabezal_hcdig on hc_prescrip.hc_nro=cabezal_hcdig.id where " & _
      "hc_prescrip.hc_fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# and hc_prescrip.hc_fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "# and hc_prescrip.hc_tippresd in ('MEDICACION') and hc_prescrip.hc_codmedica >" & 0 & " and hc_prescrip.motivo_cance is null order by hc_prescrip.hc_mat"
      data_sql.Refresh
   End If
Else
    data_sql.RecordSource = "select hc_prescrip.hc_fecha,hc_prescrip.hc_nro,hc_prescrip.hc_mat,hc_prescrip.hc_tippresd," & _
    "hc_prescrip.hc_codmedica,hc_prescrip.motivo_cance,hc_prescrip.hc_indicanom,hc_prescrip.hc_descrip,hc_prescrip.hc_hastaf," & _
    "hc_prescrip.hc_comfec,hc_prescrip.hc_fecentrega,cabezal_hcdig.id,cabezal_hcdig.hc_base from hc_prescrip " & _
    "inner join cabezal_hcdig on hc_prescrip.hc_nro=cabezal_hcdig.id where " & _
    "hc_prescrip.hc_fecha >=#" & Format(mfd.Text, "yyyy/mm/dd") & "# and hc_prescrip.hc_fecha <=#" & Format(mfh.Text, "yyyy/mm/dd") & "# and hc_prescrip.hc_tippresd in ('MEDICACION') and hc_prescrip.hc_codmedica >" & 0 & " and cabezal_hcdig.hc_base=" & t_base.Text & " and hc_prescrip.motivo_cance is null order by hc_prescrip.hc_mat"
    data_sql.Refresh
End If

If data_sql.Recordset.RecordCount > 0 Then
   data_sql.Recordset.MoveFirst
   Do While Not data_sql.Recordset.EOF
      data_mdb.Recordset.AddNew
      data_mdb.Recordset("matric") = data_sql.Recordset("hc_mat")
      data_mdb.Recordset("nombre") = data_sql.Recordset("hc_tippresd")
      data_mdb.Recordset("fecha") = data_sql.Recordset("hc_fecha")
      data_mdb.Recordset("fecpas") = data_sql.Recordset("hc_comfec")
      data_mdb.Recordset("fecsali") = data_sql.Recordset("hc_hastaf")
      data_mdb.Recordset("diag") = data_sql.Recordset("hc_indicanom")
      data_mdb.Recordset("fec_rea") = data_sql.Recordset("hc_fecentrega")
      data_mdb.Recordset("motcon") = data_sql.Recordset("hc_descrip")
      data_mdb.Recordset("movilpas") = data_sql.Recordset("hc_nro")
      data_mdb.Recordset.Update
      data_sql.Recordset.MoveNext
   Loop
   
   data_sql.Recordset.MoveFirst
   Xmat = data_sql.Recordset("hc_mat")
   Do While Not data_sql.Recordset.EOF
      If Xmat = data_sql.Recordset("hc_mat") Then
         Xcantsocio = Xcantsocio + 1
      Else
         XSocio = XSocio + 1
         data_sql.Recordset.MovePrevious
         data_inf.Recordset.AddNew
         data_inf.Recordset("cl_codigo") = data_sql.Recordset("hc_mat")
         data_inf.Recordset("cl_codced") = Xcantsocio
         data_inf.Recordset("cl_direcci") = data_sql.Recordset("hc_tippresd")
         data_inf.Recordset("cl_fecing") = data_sql.Recordset("hc_fecha")
         data_inf.Recordset("cl_apellid") = data_sql.Recordset("hc_indicanom")
         data_inf.Recordset("cl_grupo") = data_sql.Recordset("hc_base")
         data_inf.Recordset("cl_atrasop") = Xcantsocio
         Data3.RecordSource = "select * from clientes where cl_codigo =" & data_sql.Recordset("hc_mat")
         Data3.Refresh
         If Data3.Recordset.RecordCount > 0 Then
            data_inf.Recordset("cl_nombre") = Mid(Data3.Recordset("cl_apellid"), 1, 30)
            data_inf.Recordset("cl_codconv") = Data3.Recordset("cl_codconv")
         End If
         data_inf.Recordset.Update
         data_sql.Recordset.MoveNext
         Xcantsocio = 1
      End If
      Xmat = data_sql.Recordset("hc_mat")
      data_sql.Recordset.MoveNext
   Loop
End If

If data_inf.Recordset.RecordCount > 0 Then
   data_inf.Recordset.MoveFirst
   Do While Not data_inf.Recordset.EOF
      If IsNull(data_inf.Recordset("cl_codced")) = False Then
         If data_inf.Recordset("cl_codced") >= 1 Then
            data_inf.Recordset.Edit
            data_inf.Recordset("cl_atrasoa") = 30
            data_inf.Recordset.Update
         End If
      End If
      If IsNull(data_inf.Recordset("cl_cedula")) = False Then
         If data_inf.Recordset("cl_cedula") >= 1 Then
            data_inf.Recordset.Edit
            data_inf.Recordset("cl_atrasoa") = 60
            data_inf.Recordset.Update
         End If
      End If
      If IsNull(data_inf.Recordset("cl_nrovend")) = False Then
         If data_inf.Recordset("cl_nrovend") >= 1 Then
            data_inf.Recordset.Edit
            data_inf.Recordset("cl_atrasoa") = 90
            data_inf.Recordset.Update
         End If
      End If
      data_inf.Recordset.MoveNext
   Loop
End If

End Sub

Private Sub Form_Load()
'data_inf.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"
data_inf.DatabaseName = App.path & "\informes.mdb"
adocli.ConnectionString = "dsn=" & Xconexrmt

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub mfd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfh.SetFocus
End If

End Sub

Private Sub mfh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mhd.SetFocus
End If

End Sub

Private Sub mhd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mhh.SetFocus
End If

End Sub

Private Sub mhh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_base.SetFocus
End If

End Sub

Private Sub t_base_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Option1.SetFocus
End If

End Sub

Private Sub t_edd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_edh.SetFocus
End If

End Sub
