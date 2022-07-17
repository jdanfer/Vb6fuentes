VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_vtasservcp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas por Servicio"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7815
   Icon            =   "frm_vtasservcp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7815
   StartUpPosition =   1  'CenterOwner
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
      Top             =   3480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSComctlLib.ProgressBar barra 
      Height          =   375
      Left            =   840
      TabIndex        =   12
      Top             =   3120
      Visible         =   0   'False
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin Crystal.CrystalReport crs2 
      Left            =   6240
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport crs 
      Left            =   4560
      Top             =   3720
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
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "infvtas"
      Top             =   2040
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
      MouseIcon       =   "frm_vtasservcp.frx":0442
      MousePointer    =   99  'Custom
      Picture         =   "frm_vtasservcp.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir"
      Top             =   3840
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
      MouseIcon       =   "frm_vtasservcp.frx":0CD6
      MousePointer    =   99  'Custom
      Picture         =   "frm_vtasservcp.frx":0FE0
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Procesar"
      Top             =   3840
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
      Height          =   2775
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7335
      Begin VB.CheckBox Check11 
         BackColor       =   &H0080FFFF&
         Caption         =   "Solo policlínicas"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   2280
         Width           =   3255
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
         Top             =   4800
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
         Top             =   4320
         Width           =   3255
      End
      Begin MSAdodcLib.Adodc data_emi 
         Height          =   330
         Left            =   720
         Top             =   4800
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
         Caption         =   "data_emi"
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
      Begin MSAdodcLib.Adodc data2 
         Height          =   375
         Left            =   4680
         Top             =   360
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
         Caption         =   "data2"
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
      Begin MSAdodcLib.Adodc data1 
         Height          =   375
         Left            =   2280
         Top             =   1440
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
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
         Caption         =   "data1"
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
      Begin MSAdodcLib.Adodc data_lin 
         Height          =   375
         Left            =   4560
         Top             =   3720
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
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
         Top             =   4320
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
         Top             =   3960
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
         Top             =   3960
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
         Top             =   3480
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
         Top             =   3480
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
         Top             =   3000
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
         Left            =   240
         TabIndex        =   13
         Top             =   3000
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
         Left            =   3360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "estudios"
         Top             =   1800
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
         Bindings        =   "frm_vtasservcp.frx":156A
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
      Left            =   2640
      Picture         =   "frm_vtasservcp.frx":1581
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1815
   End
End
Attribute VB_Name = "frm_vtasservcp"
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
'and convenio not in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','CCNR','SMINR','MSP') and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31)
If md.Text <> "__/__/____" Then
   If mh.Text <> "__/__/____" Then
      If tm.Text <> "" Then
         If txt_b.Text <> "" Then
            If txt_b.Text = 99 Then
               If tm.Text = 999999 Then
                  If Check2.Value = 1 Then
                     If Check8.Value = 1 Then
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio not in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','CCNR','SMINR','MSP') And tot_lin >=" & 1 & " and tipo <>'" & "NOTA CR" & "' and cod_prod in (10001,10003,10005) and nro_flia not in (19,6) and base not in (16,17,18)"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio not in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','CCNR','SMINR','MSP') And tot_lin >=" & 1 & " and cod_prod in (10001,10003,10005) and nro_flia not in (19,6) and base not in (16,17,18)"
                        data_lin.Refresh
                     End If
                  Else
                     If Check3.Value = 1 Then
                        If Check8.Value = 1 Then
                           data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio not in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','CCNR','SMINR','MSP') and tipo ='" & "CREDITO" & "' and cod_prod in (10001,10003,10005) and nro_flia not in (19,6) and base not in (16,17,18)"
                           data_lin.Refresh
                        Else
                           data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio not in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','CCNR','SMINR','MSP') and tipo ='" & "CREDITO" & "' and cod_prod in (10001,10003,10005) and nro_flia not in (19,6) and base not in (16,17,18)"
                           data_lin.Refresh
                        End If
                     Else
                        If Check8.Value = 1 Then
                           data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio not in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','CCNR','SMINR','MSP') and tipo <>'" & "NOTA CR" & "' and cod_prod in (10001,10003,10005) and nro_flia not in (19,6) and base not in (16,17,18)"
                           data_lin.Refresh
                        Else
                           data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio not in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','CCNR','SMINR','MSP') and cod_prod in (10001,10003,10005,2,14001) and nro_flia not in (19,6)"
                           data_lin.Refresh
                        End If
                     End If
                  End If
               Else
                  If Check2.Value = 1 Then
                     If Check8.Value = 1 Then
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio not in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','CCNR','SMINR','MSP') and cod_prod =" & tm.Text & " And tot_lin >=" & 1 & " and tipo <>'" & "NOTA CR" & "' and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6) and base not in (16,17,18)"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio not in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','CCNR','SMINR','MSP') and cod_prod =" & tm.Text & " And tot_lin >=" & 1 & " and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6) and base not in (16,17,18)"
                        data_lin.Refresh
                     End If
                  Else
                     If Check3.Value = 1 Then
                        If Check8.Value = 1 Then
                           data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio not in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','CCNR','SMINR','MSP') and cod_prod =" & tm.Text & " And tipo ='" & "CREDITO" & "' and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6) and base not in (16,17,18)"
                           data_lin.Refresh
                        Else
                           data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio not in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','CCNR','SMINR','MSP') and cod_prod =" & tm.Text & " And tipo ='" & "CREDITO" & "' and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6) and base not in (16,17,18)"
                           data_lin.Refresh
                        End If
                     Else
                        If Check8.Value = 1 Then
                           data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio not in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','CCNR','SMINR','MSP') and cod_prod =" & tm.Text & " and tipo <>'" & "NOTA CR" & "' and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6) and base not in (16,17,18)"
                           data_lin.Refresh
                        Else
                           data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio not in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','CCNR','SMINR','MSP') and cod_prod =" & tm.Text & " and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6) and base not in (16,17,18)"
                           data_lin.Refresh
                        End If
                     End If
                  End If
               End If
            Else
               If tm.Text = 999999 Then
                  If Check2.Value = 1 Then
                     If Check8.Value = 1 Then
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio not in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','CCNR','SMINR','MSP') And base =" & txt_b.Text & " And tot_lin >=" & 1 & " and tipo <>'" & "NOTA CR" & "' and cod_prod in (10001,10003,10005) and nro_flia not in (19,6) and base not in (16,17,18)"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio not in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','CCNR','SMINR','MSP') And base =" & txt_b.Text & " And tot_lin >=" & 1 & " and cod_prod in (10001,10003,10005) and nro_flia not in (19,6) and base not in (16,17,18)"
                        data_lin.Refresh
                     End If
                  Else
                     If Check3.Value = 1 Then
                        If Check8.Value = 1 Then
                           data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio not in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','CCNR','SMINR','MSP') And base =" & txt_b.Text & " and tipo ='" & "CREDITO" & "' and cod_prod in (10001,10003,10005) and nro_flia not in (19,6) and base not in (16,17,18)"
                           data_lin.Refresh
                        Else
                           data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio not in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','CCNR','SMINR','MSP') And base =" & txt_b.Text & " and tipo ='" & "CREDITO" & "' and cod_prod in (10001,10003,10005) and nro_flia not in (19,6) and base not in (16,17,18)"
                           data_lin.Refresh
                        End If
                     Else
                        If Check8.Value = 1 Then
                           data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio not in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','CCNR','SMINR','MSP') And base =" & txt_b.Text & " and tipo <>'" & "NOTA CR" & "' and cod_prod in (10001,10003,10005) and nro_flia not in (19,6) and base not in (16,17,18)"
                           data_lin.Refresh
                        Else
                           data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio not in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','CCNR','SMINR','MSP') And base =" & txt_b.Text & " and cod_prod in (10001,10003,10005) and nro_flia not in (19,6) and base not in (16,17,18)"
                           data_lin.Refresh
                        End If
                     End If
                  End If
               Else
                  If Check2.Value = 1 Then
                     If Check8.Value = 1 Then
                        data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio not in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','CCNR','SMINR','MSP') and cod_prod =" & tm.Text & " And base =" & txt_b.Text & " And tot_lin >=" & 1 & " and tipo <>'" & "NOTA CR" & "' and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6) and base not in (16,17,18)"
                        data_lin.Refresh
                     Else
                        data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio not in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','CCNR','SMINR','MSP') and cod_prod =" & tm.Text & " And base =" & txt_b.Text & " And tot_lin >=" & 1 & " and tipo <>'" & "NOTA CR" & "' and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6) and base not in (16,17,18)"
                        data_lin.Refresh
                     End If
                  Else
                     If Check3.Value = 1 Then
                        If Check8.Value = 1 Then
                           data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio not in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','CCNR','SMINR','MSP') and cod_prod =" & tm.Text & " And base =" & txt_b.Text & " and tipo ='" & "CREDITO" & "' and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6) and base not in (16,17,18)"
                           data_lin.Refresh
                        Else
                           data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio not in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','CCNR','SMINR','MSP') and cod_prod =" & tm.Text & " And base =" & txt_b.Text & " and tipo ='" & "CREDITO" & "' and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6) and base not in (16,17,18)"
                           data_lin.Refresh
                        End If
                     Else
                        If Check8.Value = 1 Then
                           data_lin.RecordSource = "Select * from resplin where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio not in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','CCNR','SMINR','MSP') and cod_prod =" & tm.Text & " And base =" & txt_b.Text & " and tipo <>'" & "NOTA CR" & "' and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6) and base not in (16,17,18)"
                           data_lin.Refresh
                        Else
                           data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and convenio not in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','CCNR','SMINR','MSP') and cod_prod =" & tm.Text & " And base =" & txt_b.Text & " and cod_prod not in (999,992,800,881,882,991,993,994,996,997,8000,995,30,31) and nro_flia not in (19,6) and base not in (16,17,18)"
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
                  Data1.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("convenio") & "'"
                  Data1.Refresh
                  If Data1.Recordset.RecordCount > 0 Then
                     If IsNull(Data1.Recordset("cnv_grupo")) = False Then
                        If Data1.Recordset("cnv_grupo") = "UNIVERSAL" Or _
                           Data1.Recordset("cnv_grupo") = "CCOU" Or _
                           Data1.Recordset("cnv_grupo") = "SMI" Or _
                           Data1.Recordset("cnv_grupo") = "H.EVANGELICO" Or _
                           Data1.Recordset("cnv_grupo") = "CASA DE GALICIA" Or _
                           Data1.Recordset("cnv_grupo") = "IMPASA" Then
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
                  data_inf.Recordset("convenio") = data_lin.Recordset("convenio")
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
                  data_inf.Recordset("rub_cont") = data_lin.Recordset("rub_cont")
                  data_inf.Recordset("rub_nomb") = data_lin.Recordset("rub_nomb")
                  If IsNull(data_lin.Recordset("grupo")) = False Then
                     data_inf.Recordset("nro_superv") = data_lin.Recordset("grupo")
                  Else
                     data_inf.Recordset("nro_superv") = 0
                  End If
                  data_inf.Recordset("nom_med_s") = Mid(data_lin.Recordset("nom_medic"), 1, 40)
                  If IsNull(data_lin.Recordset("ced_socio")) = True Then
                     data_inf.Recordset("zona") = "0"
                  Else
                     If IsNull(data_lin.Recordset("fact")) = True Then
                        data_inf.Recordset("zona") = Trim(str(data_lin.Recordset("ced_socio"))) & "-0"
                     Else
                        data_inf.Recordset("zona") = Trim(str(data_lin.Recordset("ced_socio"))) & "-" & Trim(str(data_lin.Recordset("fact")))
                     End If
                  End If
                  data_inf.Recordset.Update
                  data_lin.Recordset.MoveNext
                  barra.Max = data_lin.Recordset.RecordCount
                  barra.Value = barra.Value + 1
               Loop
               data_inf.Refresh
               DoEvents
               MiBaseact.Execute "Delete * from infvtas where usa_timbre ='" & "S" & "'"
               data_inf.RecordSource = "Select * from infvtas order by cod_prod"
               data_inf.Refresh
               frm_vtasserv.MousePointer = 0
               If Check1.Value = 1 Then
                  crs.ReportFileName = App.path & "\infvtasxserncp.rpt"
                  crs.ReportTitle = "INFORME VENTAS POR SERVICIO FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy")
               Else
                  crs.ReportFileName = App.path & "\infvtasxsercp.rpt"
                  crs.ReportTitle = "INFORME DE VENTAS POR SERVICIO FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA: " & Format(mh.Text, "dd/mm/yyyy")
               End If
               crs.Action = 1
               barra.Visible = False
            Else
               MsgBox "No existen registros con esta selección", vbInformation, "Mensaje"
            End If
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

data_med.RecordSource = "Select * from estudios where codest in (2,10001,10002,10003,10004,10005,10006,10009,10010,10011,10012,10013,10014,10015,10016,14001) order by descrip"
data_med.Refresh
data_inf.DatabaseName = App.path & "\informes.mdb"
data_lin.ConnectionString = "dsn=" & Xconexrmt
Data1.ConnectionString = "dsn=" & Xconexrmt
'Data1.RecordSource = "clientes"
'Data1.Refresh
'Data2.DatabaseName = ""
Data2.ConnectionString = "dsn=" & Xconexrmt
data_emi.ConnectionString = "dsn=" & Xconexrmt
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
