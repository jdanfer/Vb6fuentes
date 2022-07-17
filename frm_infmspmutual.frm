VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_infmspmutual 
   BackColor       =   &H00000040&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informes por Sexo x Edad mutuales"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7875
   Icon            =   "frm_infmspmutual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command9 
      Caption         =   "Command9"
      Height          =   495
      Left            =   5760
      TabIndex        =   23
      Top             =   4440
      Width           =   1935
   End
   Begin Crystal.CrystalReport crinf 
      Left            =   3240
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc data_cli 
      Height          =   375
      Left            =   4440
      Top             =   3600
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
      Caption         =   "data_cli"
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
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   495
      Left            =   6120
      TabIndex        =   14
      Top             =   1920
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   3120
      TabIndex        =   12
      Top             =   4320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   2280
      TabIndex        =   11
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   6480
      TabIndex        =   10
      Top             =   3720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1 Benef"
      Height          =   495
      Left            =   6360
      TabIndex        =   9
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data data_inflin 
      Caption         =   "data_inflin"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4080
      Visible         =   0   'False
      Width           =   3495
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   3600
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_msp 
      Caption         =   "data_msp"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.CommandButton b_sale 
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
      Height          =   615
      Left            =   4680
      TabIndex        =   7
      Top             =   3840
      Width           =   2055
   End
   Begin VB.CommandButton b_proc 
      Caption         =   "Procesar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   3840
      Width           =   2055
   End
   Begin MSMask.MaskEdBox mh 
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
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
      Height          =   255
      Left            =   2640
      TabIndex        =   2
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   450
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Opciones de informe"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7455
      Begin VB.CommandButton Command10 
         Height          =   375
         Left            =   6600
         Picture         =   "frm_infmspmutual.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1080
         Width           =   495
      End
      Begin VB.Data data_tempg 
         Caption         =   "data_tempg"
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
         Top             =   1920
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Data data_lin2 
         Caption         =   "data_lin2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3600
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   840
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSAdodcLib.Adodc data_lin 
         Height          =   375
         Left            =   0
         Top             =   120
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
      Begin MSAdodcLib.Adodc data_conv 
         Height          =   375
         Left            =   3720
         Top             =   1320
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
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
         Caption         =   "data_conv"
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
      Begin MSAdodcLib.Adodc data_emi 
         Height          =   375
         Left            =   480
         Top             =   1320
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
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Consultaron al menos 1 vez al año"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   21
         ToolTipText     =   "Opción BENEFICIARIOS"
         Top             =   2160
         Width           =   2175
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Solo consultas de Sedes Sec."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   20
         Top             =   2280
         Width           =   3135
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   255
         Left            =   6240
         TabIndex        =   18
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Command7"
         Height          =   375
         Left            =   6360
         TabIndex        =   17
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Solo consultas (opción B.Salida)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   16
         Top             =   1800
         Width           =   3135
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Command6"
         Height          =   495
         Left            =   3840
         TabIndex        =   15
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Desde archivo"
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
         TabIndex        =   13
         Top             =   1800
         Width           =   2175
      End
      Begin VB.ComboBox cbocat 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_infmspmutual.frx":09CC
         Left            =   2520
         List            =   "frm_infmspmutual.frx":09DF
         TabIndex        =   5
         Top             =   1080
         Width           =   3735
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Planilla de:"
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
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Rango de Fechas:"
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
         TabIndex        =   1
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "El proceso demora aproximadamente 30 minutos."
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
      TabIndex        =   19
      Top             =   4560
      Width           =   6615
   End
End
Attribute VB_Name = "frm_infmspmutual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_proc_Click()
Dim Xnomem As String
Dim Xmes, Xano As Integer
Dim Xcuandias, Xqueedadt As Long
Dim Xcantsoc As Double
Dim Xcansdm, Xcansdf, Xcansd As Long
Dim Xcan1m, Xcan14m, Xcan514m, Xcan1519m, Xcan2044m, Xcan4564m, Xcan6574m, Xcan74m As Long
Dim Xcan1f, Xcan14f, Xcan514f, Xcan1519f, Xcan2044f, Xcan4564f, Xcan6574f, Xcan74f As Long
Dim Xsubtott, Xsubtottt As Double
XcuentasinC = 0

'If Check1.value = 1 Then
'   data_lin.DatabaseName = App.Path & "\llamado.mdb"
'Else
'   data_lin.DatabaseName = App.Path & "\sapp.mdb"
'End If

data_msp.DatabaseName = App.Path & "\infmsp.mdb"
data_msp.RecordSource = "plani"
data_msp.Refresh

Xcan1m = 0
Xcan14m = 0
Xcan514m = 0
Xcan1519m = 0
Xcan2044m = 0
Xcan4564m = 0
Xcan6574m = 0
Xcan74m = 0
Xcansdm = 0
Xcan1f = 0
Xcan14f = 0
Xcan514f = 0
Xcan1519f = 0
Xcan2044f = 0
Xcan4564f = 0
Xcan6574f = 0
Xcan74f = 0
Xcansdf = 0

Xnomem = "emi"
Xmes = Month(mh.Text)
Xano = Year(mh.Text)
frm_infmspmutual.MousePointer = 11
If Xmes > 9 Then
   Xnomem = Xnomem + Trim(Str(Xmes)) + Mid(Trim(Str(Xano)), 3, 2)
Else
   Xnomem = Xnomem + "0" + Trim(Str(Xmes)) + Mid(Trim(Str(Xano)), 3, 2)
End If
Dim Xlafechaem As Date
Xlafechaem = CDate(md.Text) + 2

If cbocat.ListIndex = 3 Or cbocat.ListIndex = 2 Or cbocat.ListIndex = 4 Then
   If cbocat.ListIndex = 3 Then
      data_emi.RecordSource = "Select * from " & Xnomem & " where fecha <='" & Format(Xlafechaem, "yyyy-mm-dd") & "' and grupo >=" & 300 & " and grupo <=" & 325 & " order by nro_cobr"
      data_emi.Refresh
   End If
   If cbocat.ListIndex = 2 Then
      data_emi.RecordSource = "Select * from " & Xnomem & " where fecha <='" & Format(Xlafechaem, "yyyy-mm-dd") & "' and grupo >=" & 100 & " and grupo <=" & 115 & " order by nro_cobr"
      data_emi.Refresh
   End If
   If cbocat.ListIndex = 4 Then
      data_emi.RecordSource = "Select * from " & Xnomem & " where fecha <=#" & Format(Xlafechaem, "yyyy/mm/dd") & "# and grupo =" & 810 & " order by nro_cobr"
''''      data_emi.RecordSource = "Select * from " & Xnomem & " where fecha <=#" & Format(Xlafechaem, "yyyy/mm/dd") & "# and grupo in (810) order by nro_cobr"
      
'      data_emi.RecordSource = "Select * from " & Xnomem & " where fecha <=#" & Format(Xlafechaem, "yyyy/mm/dd") & "# and grupo in (815) order by nro_cobr"
'      data_emi.RecordSource = "Select * from " & Xnomem & " where fecha <=#" & Format(Xlafechaem, "yyyy/mm/dd") & "# and grupo =" & 815 & " order by nro_cobr"
'      data_emi.RecordSource = "Select * from " & Xnomem & " where fecha <=#" & Format(Xlafechaem, "yyyy/mm/dd") & "# and grupo in (650,800,801,802,803) order by nro_cobr"
''''      data_emi.RecordSource = "Select * from " & Xnomem & " where fecha <=#" & Format(Xlafechaem, "yyyy/mm/dd") & "# and grupo in (650,800,801) order by nro_cobr"

'      data_emi.RecordSource = "Select * from " & Xnomem & " where fecha <='" & Format(Xlafechaem, "yyyy/mm/dd") & "' and grupo >=" & 600 & " and grupo <=" & 606 & " order by nro_cobr"
'      data_emi.RecordSource = "Select * from " & Xnomem & " where fecha <=#" & Format(Xlafechaem, "yyyy/mm/dd") & "# and grupo in (670,672,673,674) order by nro_cobr"
'      data_emi.RecordSource = "Select * from " & Xnomem & " where fecha <=#" & Format(Xlafechaem, "yyyy/mm/dd") & "# and grupo in (500) order by nro_cobr"
'      data_emi.RecordSource = "Select * from " & Xnomem & " where fecha <='" & Format(Xlafechaem, "yyyy/mm/dd") & "' and grupo in (401,402,403,404,405,406) order by nro_cobr"
'''      data_emi.RecordSource = "Select * from " & Xnomem & " where fecha <='" & Format(Xlafechaem, "yyyy/mm/dd") & "' and grupo in (630) order by nro_cobr"
      
''      data_emi.RecordSource = "Select * from " & Xnomem & " where fecha <='" & Format(Xlafechaem, "yyyy/mm/dd") & "' and grupo in (201,202,203,204,205,206,207,208) order by nro_cobr"
'''      data_emi.RecordSource = "Select * from " & Xnomem & " where fecha <='" & Format(Xlafechaem, "yyyy-mm-dd") & "' and grupo in (301,302,303,304,305,306,307,308,309,310) order by nro_cobr"
'      data_emi.RecordSource = "Select * from " & Xnomem & " where fecha <=#" & Format(Xlafechaem, "yyyy/mm/dd") & "# and grupo in (700,701,702,703,704,705,706,707,708,709,710,711,712,713,714,715,716,717,722) order by nro_cobr"
'      data_emi.RecordSource = "Select * from " & Xnomem & " where fecha <=#" & Format(Xlafechaem, "yyyy/mm/dd") & "# and grupo in (672,673,674) order by nro_cobr"
''      data_emi.RecordSource = "Select * from " & Xnomem & " where fecha <=#" & Format(Xlafechaem, "yyyy/mm/dd") & "# and grupo in (811,640) order by nro_cobr"
'''      data_emi.RecordSource = "Select * from " & Xnomem & " where fecha <=#" & Format(Xlafechaem, "yyyy/mm/dd") & "# and grupo in (101,102,103,104,700,701,702,703,704,705,706,707,708,709,710,711,712,713,714,715,716,717,722) order by nro_cobr"
'      data_emi.RecordSource = "Select * from " & Xnomem & " where fecha <='" & Format(Xlafechaem, "yyyy-mm-dd") & "' and grupo in (101,102,103,104,700,701,702,703,704,705,706,707,708,709,710,711,712,713,714,715,716,717,722) order by nro_cobr"
'      data_emi.RecordSource = "Select * from " & Xnomem & " where fecha <='" & Format(Xlafechaem, "yyyy-mm-dd") & "' and grupo in (670,672,673,674) order by nro_cobr"

'      data_emi.RecordSource = "Select * from " & Xnomem & " where fecha <=#" & Format(Xlafechaem, "yyyy/mm/dd") & "#"
'      data_emi.RecordSource = "Select * from " & Xnomem & " where fecha <=#" & Format(Xlafechaem, "yyyy/mm/dd") & "# and grupo in (650,800,801,802,803) order by nro_cobr"
      
      data_emi.Refresh
   End If
Else
   
'   data_emi.RecordSource = "Select * from " & Xnomem & " where fecha <=#" & Format(Xlafechaem, "yyyy/mm/dd") & "# order by nro_cobr"
   data_emi.RecordSource = "Select * from " & Xnomem & " where fecha <='" & Format(Xlafechaem, "yyyy-mm-dd") & "' order by nro_cobr"
''   data_emi.RecordSource = "Select * from " & Xnomem & " where fecha <='" & Format(Xlafechaem, "yyyy-mm-dd") & "' and nro_cobr not in (6,5,11) order by nro_cobr"
   data_emi.Refresh
End If

frm_infmspmutual.MousePointer = 11
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.Path & "\infmsp.mdb")

MiBaseact.Execute "Delete * from plani"
MiBaseact.Execute "Delete * from res"
MiBaseact.Execute "Delete * from benef"

data_msp.RecordSource = "res"
data_msp.Refresh
data_msp.RecordSource = "plani"
data_msp.Refresh

If cbocat.ListIndex = 1 Then
   data_msp.RecordSource = "benef"
   data_msp.Refresh
End If
Dim Xlamatunavez As Long

Set MiBaseact = Unasesact.OpenDatabase(App.Path & "\informes.mdb")
                
MiBaseact.Execute "Delete * from infvtas"
MiBaseact.Execute "Delete * from infcli"
MiBaseact.Execute "delete * from infcaja"

data_tempg.RecordSource = "infcaja"
data_tempg.Refresh

data_inflin.RecordSource = "infvtas"
data_inflin.Refresh
Dim Xsigrabar As Integer
Xsigrabar = 0
If md.Text <> "__/__/____" Then
   If mh.Text <> "__/__/____" Then
      If cbocat.ListIndex = 1 Or Check4.Value = 1 Then
         If Check4.Value = 1 Then
'            data_lin.DatabaseName = ""
            data_lin.ConnectionString = "dsn=" & Xconexrmt
            data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cod_prod in (10001,10002,10003,10004,10005,10006,2,14001) and convenio not in ('911','911B','SEMM','SEMM1','CASH','SJ01','SJ02') order by cod_cli"
            data_lin.Refresh
            If data_lin.Recordset.RecordCount > 0 Then
               data_lin.Recordset.MoveFirst
               Xlamatunavez = 0
               Do While Not data_lin.Recordset.EOF
                  If data_lin.Recordset("cod_cli") = Xlamatunavez Then
                  Else
                     data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("convenio") & "'"
                     data_conv.Refresh
                     If data_conv.Recordset.RecordCount > 0 Then
                        If IsNull(data_conv.Recordset("cnv_grupo")) = False Then
                           If data_conv.Recordset("cnv_grupo") = "" Then
                              If data_conv.Recordset("cnv_cant_r") = 2 And data_conv.Recordset("cnv_precio") > 0 Then
                                 data_inflin.Recordset.AddNew
                                 data_inflin.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                                 data_inflin.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                                 data_inflin.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
                                 data_inflin.Recordset("fecha") = data_lin.Recordset("fecha")
                                 data_inflin.Recordset.Update
                              End If
                           Else
                              If data_conv.Recordset("cnv_grupo") = "UNIVERSAL" Or data_conv.Recordset("cnv_grupo") = "CCOU" Or _
                                 data_conv.Recordset("cnv_grupo") = "SMI" Or data_conv.Recordset("cnv_grupo") = "CASA DE GALICIA" Or _
                                 data_conv.Recordset("cnv_grupo") = "H.EVANGELICO" Then
                                 If data_conv.Recordset("cnv_codigo") = "CCNOS" Or data_conv.Recordset("cnv_codigo") = "CASANO" Or _
                                    data_conv.Recordset("cnv_codigo") = "UNIVS" Or data_conv.Recordset("cnv_codigo") = "IMPNO" Or _
                                    data_conv.Recordset("cnv_codigo") = "HEVAN" Or data_conv.Recordset("cnv_codigo") = "HEVANO" Or _
                                    data_conv.Recordset("cnv_codigo") = "CCNRE" Or data_conv.Recordset("cnv_codigo") = "SMIN" Or data_conv.Recordset("cnv_codigo") = "UNIVNR" Or _
                                    data_conv.Recordset("cnv_codigo") = "SMINR" Or data_conv.Recordset("cnv_codigo") = "GANOS" Then
                                 Else
                                    data_inflin.Recordset.AddNew
                                    data_inflin.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                                    data_inflin.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                                    data_inflin.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
                                    data_inflin.Recordset("fecha") = data_lin.Recordset("fecha")
                                    data_inflin.Recordset.Update
                                 End If
                              End If
                           End If
                        Else
                           If data_conv.Recordset("cnv_cant_r") = 2 And data_conv.Recordset("cnv_precio") > 0 Then
                              data_inflin.Recordset.AddNew
                              data_inflin.Recordset("cod_cli") = data_lin.Recordset("cod_cli")
                              data_inflin.Recordset("nom_cli") = data_lin.Recordset("nom_cli")
                              data_inflin.Recordset("cod_prod") = data_lin.Recordset("cod_prod")
                              data_inflin.Recordset("fecha") = data_lin.Recordset("fecha")
                              data_inflin.Recordset.Update
                           End If
                        End If
                     End If
                  End If
                  Xlamatunavez = data_lin.Recordset("cod_cli")
                  data_lin.Recordset.MoveNext
               Loop
               If data_inflin.Recordset.RecordCount > 0 Then
                  data_inflin.Recordset.MoveLast
                  MsgBox "ANOTE!!! TOTAL DE REGISTROS: " & data_inflin.Recordset.RecordCount
               End If
               Unload Me
            End If
         Else
            Command1_Click
         End If
      Else
        If cbocat.ListIndex = 0 Then
           
           data_lin.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codzon in (1,2,3) and base =" & 0 & " and cancela is null and codmot ='" & "R" & "' and categ not in ('CASH','SJ01','SJ02','SJ03','SJ04','SJ05')"
           data_lin.Refresh
           If data_lin.Recordset.RecordCount > 0 Then
              data_lin.Recordset.MoveFirst
              Do While Not data_lin.Recordset.EOF
'                 data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("categ") & "'"
                    data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("categ") & "' and cnv_emite ='" & "SI" & "' and (cnv_cant_r =" & 2 & " or cnv_cant_r=" & 1 & ") and (cnv_grupo is null or cnv_grupo='" & "" & "')"
                    data_conv.Refresh
                    If data_conv.Recordset.RecordCount > 0 Then
                       If data_lin.Recordset("movilpas") <> 99 Then
                          If IsNull(data_lin.Recordset("hh")) = True Then
                             If data_lin.Recordset("matric") > 0 Then
                                data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("matric")
                                data_cli.Refresh
                                If data_cli.Recordset.RecordCount > 0 Then
                                   If IsNull(data_cli.Recordset("cl_sexo")) = False Then
                                      If data_cli.Recordset("cl_sexo") = 1 Then
                                         If data_lin.Recordset("unied") < 3 Then
                                            Xcan1m = Xcan1m + 1
                                            Xsigrabar = 9
                                         Else
                                            If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                               Xcan14m = Xcan14m + 1
                                               Xsigrabar = 9
                                            Else
                                               If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                                  Xcan514m = Xcan514m + 1
                                                  Xsigrabar = 9
                                               Else
                                                  If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                                     Xcan1519m = Xcan1519m + 1
                                                     Xsigrabar = 9
                                                  Else
                                                     If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                                        Xcan2044m = Xcan2044m + 1
                                                        Xsigrabar = 9
                                                     Else
                                                        If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                           Xcan4564m = Xcan4564m + 1
                                                           Xsigrabar = 9
                                                        Else
                                                           If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                              Xcan6574m = Xcan6574m + 1
                                                              Xsigrabar = 9
                                                           Else
                                                              If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                                 Xcan74m = Xcan74m + 1
                                                                 Xsigrabar = 9
                                                              Else
                                                                 Xcansdm = Xcansdm + 1
                                                                 Xsigrabar = 9
                                                              End If
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      Else
                                         If data_lin.Recordset("unied") < 3 Then
                                            Xcan1f = Xcan1f + 1
                                            Xsigrabar = 9
                                         Else
                                            If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                               Xcan14f = Xcan14f + 1
                                               Xsigrabar = 9
                                            Else
                                               If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                                  Xcan514f = Xcan514f + 1
                                                  Xsigrabar = 9
                                               Else
                                                  If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                                     Xcan1519f = Xcan1519f + 1
                                                     Xsigrabar = 9
                                                  Else
                                                     If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                                        Xcan2044f = Xcan2044f + 1
                                                        Xsigrabar = 9
                                                     Else
                                                        If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                           Xcan4564f = Xcan4564f + 1
                                                           Xsigrabar = 9
                                                        Else
                                                           If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                              Xcan6574f = Xcan6574f + 1
                                                              Xsigrabar = 9
                                                           Else
                                                              If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                                 Xcan74f = Xcan74f + 1
                                                                 Xsigrabar = 9
                                                              Else
                                                                 Xcansdf = Xcansdf + 1
                                                                 Xsigrabar = 9
                                                              End If
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      End If
                                   Else
                                      Xcansd = Xcansd + 1
                                      Xsigrabar = 9
                                   End If
                                Else
                                   Xcansd = Xcansd + 1
                                   Xsigrabar = 9
                                End If
                             Else
                                Xcansd = Xcansd + 1
                                Xsigrabar = 9
                             End If
                          Else
                             If data_lin.Recordset("hh") = 0 Then ' MASC
                                If data_lin.Recordset("unied") < 3 Then
                                   Xcan1m = Xcan1m + 1
                                   Xsigrabar = 9
                                Else
                                   If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                      Xcan14m = Xcan14m + 1
                                      Xsigrabar = 9
                                   Else
                                      If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                         Xcan514m = Xcan514m + 1
                                         Xsigrabar = 9
                                      Else
                                         If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                            Xcan1519m = Xcan1519m + 1
                                            Xsigrabar = 9
                                         Else
                                            If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                               Xcan2044m = Xcan2044m + 1
                                               Xsigrabar = 9
                                            Else
                                               If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                  Xcan4564m = Xcan4564m + 1
                                                  Xsigrabar = 9
                                               Else
                                                  If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                     Xcan6574m = Xcan6574m + 1
                                                     Xsigrabar = 9
                                                  Else
                                                     If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                        Xcan74m = Xcan74m + 1
                                                        Xsigrabar = 9
                                                     Else
                                                        Xcansdm = Xcansdm + 1
                                                        Xsigrabar = 9
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      End If
                                   End If
                                End If
                             Else
                                If data_lin.Recordset("hh") = 1 Then ' FEM
                                   If data_lin.Recordset("unied") < 3 Then
                                      Xcan1f = Xcan1f + 1
                                      Xsigrabar = 9
                                   Else
                                      If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                         Xcan14f = Xcan14f + 1
                                         Xsigrabar = 9
                                      Else
                                         If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                            Xcan514f = Xcan514f + 1
                                            Xsigrabar = 9
                                         Else
                                            If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                               Xcan1519f = Xcan1519f + 1
                                               Xsigrabar = 9
                                            Else
                                               If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                                  Xcan2044f = Xcan2044f + 1
                                                  Xsigrabar = 9
                                               Else
                                                  If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                     Xcan4564f = Xcan4564f + 1
                                                     Xsigrabar = 9
                                                  Else
                                                     If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                        Xcan6574f = Xcan6574f + 1
                                                        Xsigrabar = 9
                                                     Else
                                                        If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                           Xcan74f = Xcan74f + 1
                                                           Xsigrabar = 9
                                                        Else
                                                           Xcansdf = Xcansdf + 1
                                                           Xsigrabar = 9
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      End If
                                   End If
                                Else
                                   Xcansd = Xcansd + 1
                                   Xsigrabar = 9
                                End If
                             End If
                          End If
                       End If
                    Else
                       If data_lin.Recordset("categ") = "SEMM1" Or data_lin.Recordset("categ") = "SEMM" Or _
                          data_lin.Recordset("categ") = "CERSEM" Or data_lin.Recordset("categ") = "UDEMM" Or _
                          data_lin.Recordset("categ") = "UCM" Then
                          If data_lin.Recordset("movilpas") <> 99 Then
                             If IsNull(data_lin.Recordset("hh")) = True Then
                                If data_lin.Recordset("matric") > 0 Then
                                   data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("matric")
                                   data_cli.Refresh
                                   If data_cli.Recordset.RecordCount > 0 Then
                                      If IsNull(data_cli.Recordset("cl_sexo")) = False Then
                                         If data_cli.Recordset("cl_sexo") = 1 Then
                                            If data_lin.Recordset("unied") < 3 Then
                                               Xcan1m = Xcan1m + 1
                                               Xsigrabar = 9
                                            Else
                                               If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                                  Xcan14m = Xcan14m + 1
                                                  Xsigrabar = 9
                                               Else
                                                  If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                                     Xcan514m = Xcan514m + 1
                                                     Xsigrabar = 9
                                                  Else
                                                     If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                                        Xcan1519m = Xcan1519m + 1
                                                        Xsigrabar = 9
                                                     Else
                                                        If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                                           Xcan2044m = Xcan2044m + 1
                                                           Xsigrabar = 9
                                                        Else
                                                           If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                              Xcan4564m = Xcan4564m + 1
                                                              Xsigrabar = 9
                                                           Else
                                                              If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                                 Xcan6574m = Xcan6574m + 1
                                                                 Xsigrabar = 9
                                                              Else
                                                                 If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                                    Xcan74m = Xcan74m + 1
                                                                    Xsigrabar = 9
                                                                 Else
                                                                    Xcansdm = Xcansdm + 1
                                                                    Xsigrabar = 9
                                                                 End If
                                                              End If
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         Else
                                            If data_lin.Recordset("unied") < 3 Then
                                               Xcan1f = Xcan1f + 1
                                               Xsigrabar = 9
                                            Else
                                               If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                                  Xcan14f = Xcan14f + 1
                                                  Xsigrabar = 9
                                               Else
                                                  If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                                     Xcan514f = Xcan514f + 1
                                                     Xsigrabar = 9
                                                  Else
                                                     If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                                        Xcan1519f = Xcan1519f + 1
                                                        Xsigrabar = 9
                                                     Else
                                                        If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                                           Xcan2044f = Xcan2044f + 1
                                                           Xsigrabar = 9
                                                        Else
                                                           If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                              Xcan4564f = Xcan4564f + 1
                                                              Xsigrabar = 9
                                                           Else
                                                              If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                                 Xcan6574f = Xcan6574f + 1
                                                                 Xsigrabar = 9
                                                              Else
                                                                 If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                                    Xcan74f = Xcan74f + 1
                                                                    Xsigrabar = 9
                                                                 Else
                                                                    Xcansdf = Xcansdf + 1
                                                                    Xsigrabar = 9
                                                                 End If
                                                              End If
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      Else
                                         Xcansd = Xcansd + 1
                                         Xsigrabar = 9
                                      End If
                                   Else
                                      Xcansd = Xcansd + 1
                                      Xsigrabar = 9
                                   End If
                                Else
                                   Xcansd = Xcansd + 1
                                   Xsigrabar = 9
                                End If
                             Else
                                If data_lin.Recordset("hh") = 0 Then ' MASC
                                   If data_lin.Recordset("unied") < 3 Then
                                      Xcan1m = Xcan1m + 1
                                      Xsigrabar = 9
                                   Else
                                      If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                         Xcan14m = Xcan14m + 1
                                         Xsigrabar = 9
                                      Else
                                         If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                            Xcan514m = Xcan514m + 1
                                            Xsigrabar = 9
                                         Else
                                            If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                               Xcan1519m = Xcan1519m + 1
                                               Xsigrabar = 9
                                            Else
                                               If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                                  Xcan2044m = Xcan2044m + 1
                                                  Xsigrabar = 9
                                               Else
                                                  If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                     Xcan4564m = Xcan4564m + 1
                                                     Xsigrabar = 9
                                                  Else
                                                     If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                        Xcan6574m = Xcan6574m + 1
                                                        Xsigrabar = 9
                                                     Else
                                                        If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                           Xcan74m = Xcan74m + 1
                                                           Xsigrabar = 9
                                                        Else
                                                           Xcansdm = Xcansdm + 1
                                                           Xsigrabar = 9
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      End If
                                   End If
                                Else
                                   If data_lin.Recordset("hh") = 1 Then ' FEM
                                      If data_lin.Recordset("unied") < 3 Then
                                         Xcan1f = Xcan1f + 1
                                         Xsigrabar = 9
                                      Else
                                         If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                            Xcan14f = Xcan14f + 1
                                            Xsigrabar = 9
                                         Else
                                            If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                               Xcan514f = Xcan514f + 1
                                               Xsigrabar = 9
                                            Else
                                               If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                                  Xcan1519f = Xcan1519f + 1
                                                  Xsigrabar = 9
                                               Else
                                                  If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                                     Xcan2044f = Xcan2044f + 1
                                                     Xsigrabar = 9
                                                  Else
                                                     If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                        Xcan4564f = Xcan4564f + 1
                                                        Xsigrabar = 9
                                                     Else
                                                        If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                           Xcan6574f = Xcan6574f + 1
                                                           Xsigrabar = 9
                                                        Else
                                                           If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                              Xcan74f = Xcan74f + 1
                                                              Xsigrabar = 9
                                                           Else
                                                              Xcansdf = Xcansdf + 1
                                                              Xsigrabar = 9
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      End If
                                   Else
                                      Xcansd = Xcansd + 1
                                      Xsigrabar = 9
                                   End If
                                End If
                             End If
                          End If
                       End If
                    End If
                 If Xsigrabar = 9 Then
                 End If
                 data_lin.Recordset.MoveNext
                 Xsigrabar = 0
              Loop
           End If
           data_msp.Recordset.AddNew
           data_msp.Recordset("nro") = 1
           data_msp.Recordset("mes") = Month(md.Text)
           data_msp.Recordset("ano") = Year(mh.Text)
           data_msp.Recordset("desc") = "ACTIVIDAD DOMICILIARIA"
           data_msp.Recordset("desc2") = "Segun clasificación en recepción"
           data_msp.Recordset("DESC4") = "EMERGENCIA"
           data_msp.Recordset("m1") = Xcan1m
           data_msp.Recordset("m1a4") = Xcan14m
           data_msp.Recordset("m5a14") = Xcan514m
           data_msp.Recordset("m15a19") = Xcan1519m
           data_msp.Recordset("m20a44") = Xcan2044m
           data_msp.Recordset("m45a64") = Xcan4564m
           data_msp.Recordset("m65a74") = Xcan6574m
           data_msp.Recordset("m74") = Xcan74m
           data_msp.Recordset("msd") = Xcansdm
           Xsubtott = Xcan1m + Xcan14m + Xcan514m + Xcan1519m + Xcan2044m + Xcan4564m + Xcan6574m + Xcan74m + Xcansdm
           data_msp.Recordset("mt") = Xsubtott
           data_msp.Recordset("f1") = Xcan1f
           data_msp.Recordset("f1a4") = Xcan14f
           data_msp.Recordset("f5a14") = Xcan514f
           data_msp.Recordset("f15a19") = Xcan1519f
           data_msp.Recordset("f20a44") = Xcan2044f
           data_msp.Recordset("f45a64") = Xcan4564f
           data_msp.Recordset("f65a74") = Xcan6574f
           data_msp.Recordset("f74") = Xcan74f
           data_msp.Recordset("fsd") = Xcansdf
           Xsubtottt = Xcan1f + Xcan14f + Xcan514f + Xcan1519f + Xcan2044f + Xcan4564f + Xcan6574f + Xcan74f + Xcansdf
           data_msp.Recordset("ft") = Xsubtottt
           data_msp.Recordset("tot") = Xsubtott + Xsubtottt
           data_msp.Recordset.Update
           data_msp.Refresh
           Xcan1m = 0
           Xcan14m = 0
           Xcan514m = 0
           Xcan1519m = 0
           Xcan2044m = 0
           Xcan4564m = 0
           Xcan6574m = 0
           Xcan74m = 0
           Xcansdm = 0
           Xsubtott = 0
           Xsubtott = 0
           Xcan1f = 0
           Xcan14f = 0
           Xcan514f = 0
           Xcan1519f = 0
           Xcan2044f = 0
           Xcan4564f = 0
           Xcan6574f = 0
           Xcan74f = 0
           Xcansdf = 0
           Xsubtottt = 0
''''' AMARILLOS
           data_lin.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codzon in (1,2,3) and base =" & 0 & " and cancela is null and codmot ='" & "A" & "' and categ not in ('CASH','SJ01','SJ02','SJ03','SJ04','SJ05')"
           data_lin.Refresh
           If data_lin.Recordset.RecordCount > 0 Then
              data_lin.Recordset.MoveLast
              data_lin.Recordset.MoveFirst
              Do While Not data_lin.Recordset.EOF
                    data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("categ") & "' and cnv_emite ='" & "SI" & "' and (cnv_cant_r =" & 2 & " or cnv_cant_r=" & 1 & ") and (cnv_grupo is null or cnv_grupo='" & "" & "')"
                    data_conv.Refresh
                    If data_conv.Recordset.RecordCount > 0 Then
                       If data_lin.Recordset("movilpas") <> 99 Then
                          If IsNull(data_lin.Recordset("hh")) = True Then
                             If data_lin.Recordset("matric") > 0 Then
                                data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("matric")
                                data_cli.Refresh
                                If data_cli.Recordset.RecordCount > 0 Then
                                   If IsNull(data_cli.Recordset("cl_sexo")) = False Then
                                      If data_cli.Recordset("cl_sexo") = 1 Then
                                         If data_lin.Recordset("unied") < 3 Then
                                            Xcan1m = Xcan1m + 1
                                            Xsigrabar = 9
                                         Else
                                            If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                               Xcan14m = Xcan14m + 1
                                               Xsigrabar = 9
                                            Else
                                               If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                                  Xcan514m = Xcan514m + 1
                                                  Xsigrabar = 9
                                               Else
                                                  If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                                     Xcan1519m = Xcan1519m + 1
                                                     Xsigrabar = 9
                                                  Else
                                                     If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                                        Xcan2044m = Xcan2044m + 1
                                                        Xsigrabar = 9
                                                     Else
                                                        If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                           Xcan4564m = Xcan4564m + 1
                                                           Xsigrabar = 9
                                                        Else
                                                           If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                              Xcan6574m = Xcan6574m + 1
                                                              Xsigrabar = 9
                                                           Else
                                                              If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                                 Xcan74m = Xcan74m + 1
                                                                 Xsigrabar = 9
                                                              Else
                                                                 Xcansdm = Xcansdm + 1
                                                                 Xsigrabar = 9
                                                              End If
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      Else
                                         If data_lin.Recordset("unied") < 3 Then
                                            Xcan1f = Xcan1f + 1
                                            Xsigrabar = 9
                                         Else
                                            If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                               Xcan14f = Xcan14f + 1
                                               Xsigrabar = 9
                                            Else
                                               If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                                  Xcan514f = Xcan514f + 1
                                                  Xsigrabar = 9
                                               Else
                                                  If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                                     Xcan1519f = Xcan1519f + 1
                                                     Xsigrabar = 9
                                                  Else
                                                     If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                                        Xcan2044f = Xcan2044f + 1
                                                        Xsigrabar = 9
                                                     Else
                                                        If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                           Xcan4564f = Xcan4564f + 1
                                                           Xsigrabar = 9
                                                        Else
                                                           If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                              Xcan6574f = Xcan6574f + 1
                                                              Xsigrabar = 9
                                                           Else
                                                              If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                                 Xcan74f = Xcan74f + 1
                                                                 Xsigrabar = 9
                                                              Else
                                                                 Xcansdf = Xcansdf + 1
                                                                 Xsigrabar = 9
                                                              End If
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      End If
                                   Else
                                      Xcansd = Xcansd + 1
                                      Xsigrabar = 9
                                   End If
                                Else
                                   Xcansd = Xcansd + 1
                                   Xsigrabar = 9
                                End If
                             Else
                                Xcansd = Xcansd + 1
                                Xsigrabar = 9
                             End If
                          Else
                             If data_lin.Recordset("hh") = 0 Then ' MASC
                                If data_lin.Recordset("unied") < 3 Then
                                   Xcan1m = Xcan1m + 1
                                   Xsigrabar = 9
                                Else
                                   If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                      Xcan14m = Xcan14m + 1
                                      Xsigrabar = 9
                                   Else
                                      If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                         Xcan514m = Xcan514m + 1
                                         Xsigrabar = 9
                                      Else
                                         If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                            Xcan1519m = Xcan1519m + 1
                                            Xsigrabar = 9
                                         Else
                                            If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                               Xcan2044m = Xcan2044m + 1
                                               Xsigrabar = 9
                                            Else
                                               If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                  Xcan4564m = Xcan4564m + 1
                                                  Xsigrabar = 9
                                               Else
                                                  If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                     Xcan6574m = Xcan6574m + 1
                                                     Xsigrabar = 9
                                                  Else
                                                     If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                        Xcan74m = Xcan74m + 1
                                                        Xsigrabar = 9
                                                     Else
                                                        Xcansdm = Xcansdm + 1
                                                        Xsigrabar = 9
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      End If
                                   End If
                                End If
                             Else
                                If data_lin.Recordset("hh") = 1 Then ' FEM
                                   If data_lin.Recordset("unied") < 3 Then
                                      Xcan1f = Xcan1f + 1
                                      Xsigrabar = 9
                                   Else
                                      If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                         Xcan14f = Xcan14f + 1
                                         Xsigrabar = 9
                                      Else
                                         If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                            Xcan514f = Xcan514f + 1
                                            Xsigrabar = 9
                                         Else
                                            If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                               Xcan1519f = Xcan1519f + 1
                                               Xsigrabar = 9
                                            Else
                                               If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                                  Xcan2044f = Xcan2044f + 1
                                                  Xsigrabar = 9
                                               Else
                                                  If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                     Xcan4564f = Xcan4564f + 1
                                                     Xsigrabar = 9
                                                  Else
                                                     If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                        Xcan6574f = Xcan6574f + 1
                                                        Xsigrabar = 9
                                                     Else
                                                        If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                           Xcan74f = Xcan74f + 1
                                                           Xsigrabar = 9
                                                        Else
                                                           Xcansdf = Xcansdf + 1
                                                           Xsigrabar = 9
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      End If
                                   End If
                                Else
                                   Xcansd = Xcansd + 1
                                   Xsigrabar = 9
                                End If
                             End If
                          End If
                       End If
                    Else
                       If data_lin.Recordset("categ") = "SEMM1" Or data_lin.Recordset("categ") = "SEMM" Or _
                          data_lin.Recordset("categ") = "CERSEM" Or data_lin.Recordset("categ") = "UDEMM" Or _
                          data_lin.Recordset("categ") = "UCM" Then
                          If data_lin.Recordset("movilpas") <> 99 Then
                             If IsNull(data_lin.Recordset("hh")) = True Then
                                If data_lin.Recordset("matric") > 0 Then
                                   data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("matric")
                                   data_cli.Refresh
                                   If data_cli.Recordset.RecordCount > 0 Then
                                      If IsNull(data_cli.Recordset("cl_sexo")) = False Then
                                         If data_cli.Recordset("cl_sexo") = 1 Then
                                            If data_lin.Recordset("unied") < 3 Then
                                               Xcan1m = Xcan1m + 1
                                               Xsigrabar = 9
                                            Else
                                               If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                                  Xcan14m = Xcan14m + 1
                                                  Xsigrabar = 9
                                               Else
                                                  If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                                     Xcan514m = Xcan514m + 1
                                                     Xsigrabar = 9
                                                  Else
                                                     If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                                        Xcan1519m = Xcan1519m + 1
                                                        Xsigrabar = 9
                                                     Else
                                                        If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                                           Xcan2044m = Xcan2044m + 1
                                                           Xsigrabar = 9
                                                        Else
                                                           If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                              Xcan4564m = Xcan4564m + 1
                                                              Xsigrabar = 9
                                                           Else
                                                              If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                                 Xcan6574m = Xcan6574m + 1
                                                                 Xsigrabar = 9
                                                              Else
                                                                 If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                                    Xcan74m = Xcan74m + 1
                                                                    Xsigrabar = 9
                                                                 Else
                                                                    Xcansdm = Xcansdm + 1
                                                                    Xsigrabar = 9
                                                                 End If
                                                              End If
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         Else
                                            If data_lin.Recordset("unied") < 3 Then
                                               Xcan1f = Xcan1f + 1
                                               Xsigrabar = 9
                                            Else
                                               If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                                  Xcan14f = Xcan14f + 1
                                                  Xsigrabar = 9
                                               Else
                                                  If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                                     Xcan514f = Xcan514f + 1
                                                     Xsigrabar = 9
                                                  Else
                                                     If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                                        Xcan1519f = Xcan1519f + 1
                                                        Xsigrabar = 9
                                                     Else
                                                        If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                                           Xcan2044f = Xcan2044f + 1
                                                           Xsigrabar = 9
                                                        Else
                                                           If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                              Xcan4564f = Xcan4564f + 1
                                                              Xsigrabar = 9
                                                           Else
                                                              If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                                 Xcan6574f = Xcan6574f + 1
                                                                 Xsigrabar = 9
                                                              Else
                                                                 If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                                    Xcan74f = Xcan74f + 1
                                                                    Xsigrabar = 9
                                                                 Else
                                                                    Xcansdf = Xcansdf + 1
                                                                    Xsigrabar = 9
                                                                 End If
                                                              End If
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      Else
                                         Xcansd = Xcansd + 1
                                         Xsigrabar = 9
                                      End If
                                   Else
                                      Xcansd = Xcansd + 1
                                      Xsigrabar = 9
                                   End If
                                Else
                                   Xcansd = Xcansd + 1
                                   Xsigrabar = 9
                                End If
                             Else
                                If data_lin.Recordset("hh") = 0 Then ' MASC
                                   If data_lin.Recordset("unied") < 3 Then
                                      Xcan1m = Xcan1m + 1
                                      Xsigrabar = 9
                                   Else
                                      If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                         Xcan14m = Xcan14m + 1
                                         Xsigrabar = 9
                                      Else
                                         If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                            Xcan514m = Xcan514m + 1
                                            Xsigrabar = 9
                                         Else
                                            If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                               Xcan1519m = Xcan1519m + 1
                                               Xsigrabar = 9
                                            Else
                                               If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                                  Xcan2044m = Xcan2044m + 1
                                                  Xsigrabar = 9
                                               Else
                                                  If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                     Xcan4564m = Xcan4564m + 1
                                                     Xsigrabar = 9
                                                  Else
                                                     If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                        Xcan6574m = Xcan6574m + 1
                                                        Xsigrabar = 9
                                                     Else
                                                        If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                           Xcan74m = Xcan74m + 1
                                                           Xsigrabar = 9
                                                        Else
                                                           Xcansdm = Xcansdm + 1
                                                           Xsigrabar = 9
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      End If
                                   End If
                                Else
                                   If data_lin.Recordset("hh") = 1 Then ' FEM
                                      If data_lin.Recordset("unied") < 3 Then
                                         Xcan1f = Xcan1f + 1
                                         Xsigrabar = 9
                                      Else
                                         If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                            Xcan14f = Xcan14f + 1
                                            Xsigrabar = 9
                                         Else
                                            If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                               Xcan514f = Xcan514f + 1
                                               Xsigrabar = 9
                                            Else
                                               If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                                  Xcan1519f = Xcan1519f + 1
                                                  Xsigrabar = 9
                                               Else
                                                  If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                                     Xcan2044f = Xcan2044f + 1
                                                     Xsigrabar = 9
                                                  Else
                                                     If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                        Xcan4564f = Xcan4564f + 1
                                                        Xsigrabar = 9
                                                     Else
                                                        If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                           Xcan6574f = Xcan6574f + 1
                                                           Xsigrabar = 9
                                                        Else
                                                           If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                              Xcan74f = Xcan74f + 1
                                                              Xsigrabar = 9
                                                           Else
                                                              Xcansdf = Xcansdf + 1
                                                              Xsigrabar = 9
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      End If
                                   Else
                                      Xcansd = Xcansd + 1
                                      Xsigrabar = 9
                                   End If
                                End If
                             End If
                          End If
                       End If
                    End If
                 
                 If Xsigrabar = 9 Then
                 End If
                 data_lin.Recordset.MoveNext
                 Xsigrabar = 0
              Loop
           End If
           
           data_msp.Recordset.Edit
           data_msp.Recordset("DESC5") = "URGENCIA"
           data_msp.Recordset("m1a") = Xcan1m
           data_msp.Recordset("m1a4a") = Xcan14m
           data_msp.Recordset("m5a14a") = Xcan514m
           data_msp.Recordset("m15a19a") = Xcan1519m
           data_msp.Recordset("m20a44a") = Xcan2044m
           data_msp.Recordset("m45a64a") = Xcan4564m
           data_msp.Recordset("m65a74a") = Xcan6574m
           data_msp.Recordset("m74a") = Xcan74m
           data_msp.Recordset("msda") = Xcansdm
           Xsubtott = Xcan1m + Xcan14m + Xcan514m + Xcan1519m + Xcan2044m + Xcan4564m + Xcan6574m + Xcan74m + Xcansdm
           data_msp.Recordset("mta") = Xsubtott
           data_msp.Recordset("f1a") = Xcan1f
           data_msp.Recordset("f1a4a") = Xcan14f
           data_msp.Recordset("f5a14a") = Xcan514f
           data_msp.Recordset("f15a19a") = Xcan1519f
           data_msp.Recordset("f20a44a") = Xcan2044f
           data_msp.Recordset("f45a64a") = Xcan4564f
           data_msp.Recordset("f65a74a") = Xcan6574f
           data_msp.Recordset("f74a") = Xcan74f
           data_msp.Recordset("fsda") = Xcansdf
           Xsubtottt = Xcan1f + Xcan14f + Xcan514f + Xcan1519f + Xcan2044f + Xcan4564f + Xcan6574f + Xcan74f + Xcansdf
           data_msp.Recordset("fta") = Xsubtottt
           data_msp.Recordset("tota") = Xsubtott + Xsubtottt
           data_msp.Recordset.Update
           Xcan1m = 0
           Xcan14m = 0
           Xcan514m = 0
           Xcan1519m = 0
           Xcan2044m = 0
           Xcan4564m = 0
           Xcan6574m = 0
           Xcan74m = 0
           Xcansdm = 0
           Xsubtott = 0
           Xsubtott = 0
           Xcan1f = 0
           Xcan14f = 0
           Xcan514f = 0
           Xcan1519f = 0
           Xcan2044f = 0
           Xcan4564f = 0
           Xcan6574f = 0
           Xcan74f = 0
           Xcansdf = 0
           Xsubtottt = 0
           data_lin.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codzon in (1,2,3) and base =" & 0 & " and cancela is null and codmot in ('V','Z') and categ not in ('CASH','SJ01','SJ02','SJ03','SJ04','SJ05')"
           data_lin.Refresh
           If data_lin.Recordset.RecordCount > 0 Then
              data_lin.Recordset.MoveFirst
              Do While Not data_lin.Recordset.EOF
                    data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("categ") & "' and cnv_emite ='" & "SI" & "' and (cnv_cant_r =" & 2 & " or cnv_cant_r=" & 1 & ") and (cnv_grupo is null or cnv_grupo='" & "" & "')"
                    data_conv.Refresh
                    If data_conv.Recordset.RecordCount > 0 Then
                       If data_lin.Recordset("movilpas") <> 99 Then
                          If IsNull(data_lin.Recordset("hh")) = True Then
                             If data_lin.Recordset("matric") > 0 Then
                                data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("matric")
                                data_cli.Refresh
                                If data_cli.Recordset.RecordCount > 0 Then
                                   If IsNull(data_cli.Recordset("cl_sexo")) = False Then
                                      If data_cli.Recordset("cl_sexo") = 1 Then
                                         If data_lin.Recordset("unied") < 3 Then
                                            Xcan1m = Xcan1m + 1
                                            Xsigrabar = 9
                                         Else
                                            If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                               Xcan14m = Xcan14m + 1
                                               Xsigrabar = 9
                                            Else
                                               If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                                  Xcan514m = Xcan514m + 1
                                                  Xsigrabar = 9
                                               Else
                                                  If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                                     Xcan1519m = Xcan1519m + 1
                                                     Xsigrabar = 9
                                                  Else
                                                     If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                                        Xcan2044m = Xcan2044m + 1
                                                        Xsigrabar = 9
                                                     Else
                                                        If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                           Xcan4564m = Xcan4564m + 1
                                                           Xsigrabar = 9
                                                        Else
                                                           If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                              Xcan6574m = Xcan6574m + 1
                                                              Xsigrabar = 9
                                                           Else
                                                              If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                                 Xcan74m = Xcan74m + 1
                                                                 Xsigrabar = 9
                                                              Else
                                                                 Xcansdm = Xcansdm + 1
                                                                 Xsigrabar = 9
                                                              End If
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      Else
                                         If data_lin.Recordset("unied") < 3 Then
                                            Xcan1f = Xcan1f + 1
                                            Xsigrabar = 9
                                         Else
                                            If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                               Xcan14f = Xcan14f + 1
                                               Xsigrabar = 9
                                            Else
                                               If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                                  Xcan514f = Xcan514f + 1
                                                  Xsigrabar = 9
                                               Else
                                                  If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                                     Xcan1519f = Xcan1519f + 1
                                                     Xsigrabar = 9
                                                  Else
                                                     If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                                        Xcan2044f = Xcan2044f + 1
                                                        Xsigrabar = 9
                                                     Else
                                                        If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                           Xcan4564f = Xcan4564f + 1
                                                           Xsigrabar = 9
                                                        Else
                                                           If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                              Xcan6574f = Xcan6574f + 1
                                                              Xsigrabar = 9
                                                           Else
                                                              If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                                 Xcan74f = Xcan74f + 1
                                                                 Xsigrabar = 9
                                                              Else
                                                                 Xcansdf = Xcansdf + 1
                                                                 Xsigrabar = 9
                                                              End If
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      End If
                                   Else
                                      Xcansd = Xcansd + 1
                                      Xsigrabar = 9
                                   End If
                                Else
                                   Xcansd = Xcansd + 1
                                   Xsigrabar = 9
                                End If
                             Else
                                Xcansd = Xcansd + 1
                                Xsigrabar = 9
                             End If
                          Else
                             If data_lin.Recordset("hh") = 0 Then ' MASC
                                If data_lin.Recordset("unied") < 3 Then
                                   Xcan1m = Xcan1m + 1
                                   Xsigrabar = 9
                                Else
                                   If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                      Xcan14m = Xcan14m + 1
                                      Xsigrabar = 9
                                   Else
                                      If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                         Xcan514m = Xcan514m + 1
                                         Xsigrabar = 9
                                      Else
                                         If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                            Xcan1519m = Xcan1519m + 1
                                            Xsigrabar = 9
                                         Else
                                            If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                               Xcan2044m = Xcan2044m + 1
                                               Xsigrabar = 9
                                            Else
                                               If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                  Xcan4564m = Xcan4564m + 1
                                                  Xsigrabar = 9
                                               Else
                                                  If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                     Xcan6574m = Xcan6574m + 1
                                                     Xsigrabar = 9
                                                  Else
                                                     If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                        Xcan74m = Xcan74m + 1
                                                        Xsigrabar = 9
                                                     Else
                                                        Xcansdm = Xcansdm + 1
                                                        Xsigrabar = 9
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      End If
                                   End If
                                End If
                             Else
                                If data_lin.Recordset("hh") = 1 Then ' FEM
                                   If data_lin.Recordset("unied") < 3 Then
                                      Xcan1f = Xcan1f + 1
                                      Xsigrabar = 9
                                   Else
                                      If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                         Xcan14f = Xcan14f + 1
                                         Xsigrabar = 9
                                      Else
                                         If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                            Xcan514f = Xcan514f + 1
                                            Xsigrabar = 9
                                         Else
                                            If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                               Xcan1519f = Xcan1519f + 1
                                               Xsigrabar = 9
                                            Else
                                               If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                                  Xcan2044f = Xcan2044f + 1
                                                  Xsigrabar = 9
                                               Else
                                                  If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                     Xcan4564f = Xcan4564f + 1
                                                     Xsigrabar = 9
                                                  Else
                                                     If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                        Xcan6574f = Xcan6574f + 1
                                                        Xsigrabar = 9
                                                     Else
                                                        If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                           Xcan74f = Xcan74f + 1
                                                           Xsigrabar = 9
                                                        Else
                                                           Xcansdf = Xcansdf + 1
                                                           Xsigrabar = 9
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      End If
                                   End If
                                Else
                                   Xcansd = Xcansd + 1
                                   Xsigrabar = 9
                                End If
                             End If
                          End If
                       End If
                    Else
                       If data_lin.Recordset("categ") = "SEMM1" Or data_lin.Recordset("categ") = "SEMM" Or _
                          data_lin.Recordset("categ") = "CERSEM" Or data_lin.Recordset("categ") = "UDEMM" Or _
                          data_lin.Recordset("categ") = "UCM" Then
                          If data_lin.Recordset("movilpas") <> 99 Then
                             If IsNull(data_lin.Recordset("hh")) = True Then
                                If data_lin.Recordset("matric") > 0 Then
                                   data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("matric")
                                   data_cli.Refresh
                                   If data_cli.Recordset.RecordCount > 0 Then
                                      If IsNull(data_cli.Recordset("cl_sexo")) = False Then
                                         If data_cli.Recordset("cl_sexo") = 1 Then
                                            If data_lin.Recordset("unied") < 3 Then
                                               Xcan1m = Xcan1m + 1
                                               Xsigrabar = 9
                                            Else
                                               If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                                  Xcan14m = Xcan14m + 1
                                                  Xsigrabar = 9
                                               Else
                                                  If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                                     Xcan514m = Xcan514m + 1
                                                     Xsigrabar = 9
                                                  Else
                                                     If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                                        Xcan1519m = Xcan1519m + 1
                                                        Xsigrabar = 9
                                                     Else
                                                        If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                                           Xcan2044m = Xcan2044m + 1
                                                           Xsigrabar = 9
                                                        Else
                                                           If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                              Xcan4564m = Xcan4564m + 1
                                                              Xsigrabar = 9
                                                           Else
                                                              If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                                 Xcan6574m = Xcan6574m + 1
                                                                 Xsigrabar = 9
                                                              Else
                                                                 If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                                    Xcan74m = Xcan74m + 1
                                                                    Xsigrabar = 9
                                                                 Else
                                                                    Xcansdm = Xcansdm + 1
                                                                    Xsigrabar = 9
                                                                 End If
                                                              End If
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         Else
                                            If data_lin.Recordset("unied") < 3 Then
                                               Xcan1f = Xcan1f + 1
                                               Xsigrabar = 9
                                            Else
                                               If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                                  Xcan14f = Xcan14f + 1
                                                  Xsigrabar = 9
                                               Else
                                                  If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                                     Xcan514f = Xcan514f + 1
                                                     Xsigrabar = 9
                                                  Else
                                                     If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                                        Xcan1519f = Xcan1519f + 1
                                                        Xsigrabar = 9
                                                     Else
                                                        If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                                           Xcan2044f = Xcan2044f + 1
                                                           Xsigrabar = 9
                                                        Else
                                                           If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                              Xcan4564f = Xcan4564f + 1
                                                              Xsigrabar = 9
                                                           Else
                                                              If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                                 Xcan6574f = Xcan6574f + 1
                                                                 Xsigrabar = 9
                                                              Else
                                                                 If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                                    Xcan74f = Xcan74f + 1
                                                                    Xsigrabar = 9
                                                                 Else
                                                                    Xcansdf = Xcansdf + 1
                                                                    Xsigrabar = 9
                                                                 End If
                                                              End If
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      Else
                                         Xcansd = Xcansd + 1
                                         Xsigrabar = 9
                                      End If
                                   Else
                                      Xcansd = Xcansd + 1
                                      Xsigrabar = 9
                                   End If
                                Else
                                   Xcansd = Xcansd + 1
                                   Xsigrabar = 9
                                End If
                             Else
                                If data_lin.Recordset("hh") = 0 Then ' MASC
                                   If data_lin.Recordset("unied") < 3 Then
                                      Xcan1m = Xcan1m + 1
                                      Xsigrabar = 9
                                   Else
                                      If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                         Xcan14m = Xcan14m + 1
                                         Xsigrabar = 9
                                      Else
                                         If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                            Xcan514m = Xcan514m + 1
                                            Xsigrabar = 9
                                         Else
                                            If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                               Xcan1519m = Xcan1519m + 1
                                               Xsigrabar = 9
                                            Else
                                               If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                                  Xcan2044m = Xcan2044m + 1
                                                  Xsigrabar = 9
                                               Else
                                                  If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                     Xcan4564m = Xcan4564m + 1
                                                     Xsigrabar = 9
                                                  Else
                                                     If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                        Xcan6574m = Xcan6574m + 1
                                                        Xsigrabar = 9
                                                     Else
                                                        If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                           Xcan74m = Xcan74m + 1
                                                           Xsigrabar = 9
                                                        Else
                                                           Xcansdm = Xcansdm + 1
                                                           Xsigrabar = 9
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      End If
                                   End If
                                Else
                                   If data_lin.Recordset("hh") = 1 Then ' FEM
                                      If data_lin.Recordset("unied") < 3 Then
                                         Xcan1f = Xcan1f + 1
                                         Xsigrabar = 9
                                      Else
                                         If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                            Xcan14f = Xcan14f + 1
                                            Xsigrabar = 9
                                         Else
                                            If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                               Xcan514f = Xcan514f + 1
                                               Xsigrabar = 9
                                            Else
                                               If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                                  Xcan1519f = Xcan1519f + 1
                                                  Xsigrabar = 9
                                               Else
                                                  If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                                     Xcan2044f = Xcan2044f + 1
                                                     Xsigrabar = 9
                                                  Else
                                                     If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                        Xcan4564f = Xcan4564f + 1
                                                        Xsigrabar = 9
                                                     Else
                                                        If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                           Xcan6574f = Xcan6574f + 1
                                                           Xsigrabar = 9
                                                        Else
                                                           If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                              Xcan74f = Xcan74f + 1
                                                              Xsigrabar = 9
                                                           Else
                                                              Xcansdf = Xcansdf + 1
                                                              Xsigrabar = 9
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      End If
                                   Else
                                      Xcansd = Xcansd + 1
                                      Xsigrabar = 9
                                   End If
                                End If
                             End If
                          End If
                       End If
                    End If
                 
                 If Xsigrabar = 9 Then
                 End If
                 data_lin.Recordset.MoveNext
                 Xsigrabar = 0
              Loop
           End If
            data_msp.Recordset.Edit
            data_msp.Recordset("DESC6") = "RADIO"
            data_msp.Recordset("m1v") = Xcan1m
            data_msp.Recordset("m1a4v") = Xcan14m
            data_msp.Recordset("m5a14v") = Xcan514m
            data_msp.Recordset("m15a19v") = Xcan1519m
            data_msp.Recordset("m20a44v") = Xcan2044m
            data_msp.Recordset("m45a64v") = Xcan4564m
            data_msp.Recordset("m65a74v") = Xcan6574m
            data_msp.Recordset("m74v") = Xcan74m
            data_msp.Recordset("msdv") = Xcansdm
            Xsubtott = Xcan1m + Xcan14m + Xcan514m + Xcan1519m + Xcan2044m + Xcan4564m + Xcan6574m + Xcan74m + Xcansdm
            data_msp.Recordset("mtv") = Xsubtott
            data_msp.Recordset("f1v") = Xcan1f
            data_msp.Recordset("f1a4v") = Xcan14f
            data_msp.Recordset("f5a14v") = Xcan514f
            data_msp.Recordset("f15a19v") = Xcan1519f
            data_msp.Recordset("f20a44v") = Xcan2044f
            data_msp.Recordset("f45a64v") = Xcan4564f
            data_msp.Recordset("f65a74v") = Xcan6574f
            data_msp.Recordset("f74v") = Xcan74f
            data_msp.Recordset("fsdv") = Xcansdf
            Xsubtottt = Xcan1f + Xcan14f + Xcan514f + Xcan1519f + Xcan2044f + Xcan4564f + Xcan6574f + Xcan74f + Xcansdf
            data_msp.Recordset("ftv") = Xsubtottt
            data_msp.Recordset("totv") = Xsubtott + Xsubtottt
            data_msp.Recordset.Update
            Xcan1m = 0
            Xcan14m = 0
            Xcan514m = 0
            Xcan1519m = 0
            Xcan2044m = 0
            Xcan4564m = 0
            Xcan6574m = 0
            Xcan74m = 0
            Xcansdm = 0
            Xsubtott = 0
            Xsubtott = 0
            Xcan1f = 0
            Xcan14f = 0
            Xcan514f = 0
            Xcan1519f = 0
            Xcan2044f = 0
            Xcan4564f = 0
            Xcan6574f = 0
            Xcan74f = 0
            Xcansdf = 0
            Xsubtottt = 0
 '           cr1.ReportFileName = App.Path & "\infmspok.rpt"
 '           cr1.ReportTitle = "INFORME DE SERVICIOS POR SEXO POR EDAD --CONVENIO: " & cbocat.Text & " DESDE: " & md.Text & " HASTA: " & mh.Text
 '           cr1.Action = 1
                 
            ''' aca el otro command para las consultas en domicilio segun codigo final
            Command2_Click
        Else
            If cbocat.ListIndex = 5 Then
            
            Else
               If cbocat.ListIndex = 3 Or cbocat.ListIndex = 2 Or cbocat.ListIndex = 4 Then
                  Command5_Click
               End If
            End If
        End If
      End If
   Else
      MsgBox "Verifique FECHAS"
   End If
Else
   MsgBox "Verifique FECHAS"
End If
frm_infmspmutual.MousePointer = 0

If cbocat.ListIndex = 0 Then
   MsgBox "Proceso terminado"
   cr1.ReportFileName = App.Path & "\infplani.rpt"
   cr1.Action = 1
End If

End Sub

Private Sub b_sale_Click()
Unload Me

End Sub

Private Sub cbocat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_proc.SetFocus
End If

End Sub

Private Sub Command1_Click()
Dim m1, m1a4, m5a14, m15a19, m20a44, m45a64, m65a74, m74, msd As Long
Dim f1, f1a4, f5a14, f15a19, f20a44, f45a64, f65a74, f74, fsd As Long
Dim m1ia, m1a4ia, m5a14ia, m15a19ia, m20a44ia, m45a64ia, m65a74ia, m74ia, msdia As Long
Dim f1ia, f1a4ia, f5a14ia, f15a19ia, f20a44ia, f45a64ia, f65a74ia, f74ia, fsdia As Long
Dim Xanos As Double
Dim Xsindatos, Xsindatosia As Long
Dim Grupo As String



Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.Path & "\informes.mdb")

MiBaseact.Execute "Delete * from infcli"
'MiBaseact.Execute "Delete * from benef"

data_inf.RecordSource = "infcli"
data_inf.Refresh

Set MiBaseact = Unasesact.OpenDatabase(App.Path & "\mat.mdb")
MiBaseact.Execute "Delete * from mat"

'Data1.RecordSource = "mat"
'Data1.Refresh

Xsindatos = 0
If data_emi.Recordset.RecordCount > 0 Then
   data_emi.Recordset.MoveLast
   pb.Max = data_emi.Recordset.RecordCount
   data_emi.Recordset.MoveFirst
   DoEvents
   Do While Not data_emi.Recordset.EOF
      data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_emi.Recordset("cliente")
      data_cli.Refresh
      If data_cli.Recordset.RecordCount > 0 Then
         If data_emi.Recordset("nro_cobr") = 5 Or _
            data_emi.Recordset("nro_cobr") = 6 Or _
            data_emi.Recordset("nro_cobr") = 11 Then
         Else
            data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_emi.Recordset("cod_cnv") & "'"
            data_conv.Refresh
            If data_conv.Recordset.RecordCount > 0 Then
               If data_conv.Recordset("cnv_cant_r") = 1 Then

               Else
               
                    If IsNull(data_conv.Recordset("cnv_grupo")) = False Then
                       If data_conv.Recordset("cnv_grupo") = "" Then
                            If IsNull(data_cli.Recordset("cl_sexo")) = True Then
                               Xsindatos = Xsindatos + 1
                            Else
                               If data_cli.Recordset("cl_sexo") = 2 Then ' FEM
                                  If IsNull(data_cli.Recordset("cl_fnac")) = True Then
                                     fsd = fsd + 1
                                  Else
                                     Xanos = Date - data_cli.Recordset("cl_fnac")
                                     Xanos = Xanos / 365
                                     If Xanos < 1 Then
                                        f1 = f1 + 1
                                     Else
                                        If Xanos >= 1 And Xanos <= 4 Then
                                           f1a4 = f1a4 + 1
                                        Else
                                           If Xanos >= 5 And Xanos <= 14 Then
                                              f5a14 = f5a14 + 1
                                           Else
                                              If Xanos >= 15 And Xanos <= 19 Then
                                                 f15a19 = f15a19 + 1
                                              Else
                                                 If Xanos >= 20 And Xanos <= 44 Then
                                                    f20a44 = f20a44 + 1
                                                 Else
                                                    If Xanos >= 45 And Xanos <= 64 Then
                                                       f45a64 = f45a64 + 1
                                                    Else
                                                       If Xanos >= 65 And Xanos <= 74 Then
                                                          f65a74 = f65a74 + 1
                                                       Else
                                                          If Xanos >= 75 And Xanos <= 110 Then
                                                             f74 = f74 + 1
                                                          Else
                                                             fsd = fsd + 1
                                                          End If
                                                       End If
                                                    End If
                                                 End If
                                              End If
                                           End If
                                        End If
                                     End If
                                  End If
                               Else
                                  If data_cli.Recordset("cl_sexo") = 1 Then ' MASC
                                     If IsNull(data_cli.Recordset("cl_fnac")) = True Then
                                        msd = msd + 1
                                     Else
                                        Xanos = Date - data_cli.Recordset("cl_fnac")
                                        Xanos = Xanos / 365
                                        If Xanos < 1 Then
                                           m1 = m1 + 1
                                        Else
                                           If Xanos >= 1 And Xanos <= 4 Then
                                              m1a4 = m1a4 + 1
                                           Else
                                              If Xanos >= 5 And Xanos <= 14 Then
                                                 m5a14 = m5a14 + 1
                                              Else
                                                 If Xanos >= 15 And Xanos <= 19 Then
                                                    m15a19 = m15a19 + 1
                                                 Else
                                                    If Xanos >= 20 And Xanos <= 44 Then
                                                       m20a44 = m20a44 + 1
                                                    Else
                                                       If Xanos >= 45 And Xanos <= 64 Then
                                                          m45a64 = m45a64 + 1
                                                       Else
                                                          If Xanos >= 65 And Xanos <= 74 Then
                                                             m65a74 = m65a74 + 1
                                                          Else
                                                             If Xanos >= 75 And Xanos <= 110 Then
                                                                m74 = m74 + 1
                                                             Else
                                                                msd = msd + 1
                                                             End If
                                                          End If
                                                       End If
                                                    End If
                                                 End If
                                              End If
                                           End If
                                        End If
                                     End If
                                  Else
                                     Xsindatos = Xsindatos + 1
                                  End If
                               End If
                            End If
                       Else
                            If IsNull(data_cli.Recordset("cl_sexo")) = True Then
                               Xsindatos = Xsindatos + 1
                            Else
                               If data_cli.Recordset("cl_sexo") = 2 Then ' FEM
                                  If IsNull(data_cli.Recordset("cl_fnac")) = True Then
                                     fsd = fsd + 1
                                  Else
                                     Xanos = Date - data_cli.Recordset("cl_fnac")
                                     Xanos = Xanos / 365
                                     If Xanos < 1 Then
                                        f1 = f1 + 1
                                     Else
                                        If Xanos >= 1 And Xanos <= 4 Then
                                           f1a4 = f1a4 + 1
                                        Else
                                           If Xanos >= 5 And Xanos <= 14 Then
                                              f5a14 = f5a14 + 1
                                           Else
                                              If Xanos >= 15 And Xanos <= 19 Then
                                                 f15a19 = f15a19 + 1
                                              Else
                                                 If Xanos >= 20 And Xanos <= 44 Then
                                                    f20a44 = f20a44 + 1
                                                 Else
                                                    If Xanos >= 45 And Xanos <= 64 Then
                                                       f45a64 = f45a64 + 1
                                                    Else
                                                       If Xanos >= 65 And Xanos <= 74 Then
                                                          f65a74 = f65a74 + 1
                                                       Else
                                                          If Xanos >= 75 And Xanos <= 110 Then
                                                             f74 = f74 + 1
                                                          Else
                                                             fsd = fsd + 1
                                                          End If
                                                       End If
                                                    End If
                                                 End If
                                              End If
                                           End If
                                        End If
                                     End If
                                  End If
                               Else
                                  If data_cli.Recordset("cl_sexo") = 1 Then ' MASC
                                     If IsNull(data_cli.Recordset("cl_fnac")) = True Then
                                        msd = msd + 1
                                     Else
                                        Xanos = Date - data_cli.Recordset("cl_fnac")
                                        Xanos = Xanos / 365
                                        If Xanos < 1 Then
                                           m1 = m1 + 1
                                        Else
                                           If Xanos >= 1 And Xanos <= 4 Then
                                              m1a4 = m1a4 + 1
                                           Else
                                              If Xanos >= 5 And Xanos <= 14 Then
                                                 m5a14 = m5a14 + 1
                                              Else
                                                 If Xanos >= 15 And Xanos <= 19 Then
                                                    m15a19 = m15a19 + 1
                                                 Else
                                                    If Xanos >= 20 And Xanos <= 44 Then
                                                       m20a44 = m20a44 + 1
                                                    Else
                                                       If Xanos >= 45 And Xanos <= 64 Then
                                                          m45a64 = m45a64 + 1
                                                       Else
                                                          If Xanos >= 65 And Xanos <= 74 Then
                                                             m65a74 = m65a74 + 1
                                                          Else
                                                             If Xanos >= 75 And Xanos <= 110 Then
                                                                m74 = m74 + 1
                                                             Else
                                                                msd = msd + 1
                                                             End If
                                                          End If
                                                       End If
                                                    End If
                                                 End If
                                              End If
                                           End If
                                        End If
                                     End If
                                  Else
                                     Xsindatos = Xsindatos + 1
                                  End If
                               End If
                            End If
                       
                       End If
                    Else
                            If IsNull(data_cli.Recordset("cl_sexo")) = True Then
                               Xsindatos = Xsindatos + 1
                            Else
                               If data_cli.Recordset("cl_sexo") = 2 Then ' FEM
                                  If IsNull(data_cli.Recordset("cl_fnac")) = True Then
                                     fsd = fsd + 1
                                  Else
                                     Xanos = Date - data_cli.Recordset("cl_fnac")
                                     Xanos = Xanos / 365
                                     If Xanos < 1 Then
                                        f1 = f1 + 1
                                     Else
                                        If Xanos >= 1 And Xanos <= 4 Then
                                           f1a4 = f1a4 + 1
                                        Else
                                           If Xanos >= 5 And Xanos <= 14 Then
                                              f5a14 = f5a14 + 1
                                           Else
                                              If Xanos >= 15 And Xanos <= 19 Then
                                                 f15a19 = f15a19 + 1
                                              Else
                                                 If Xanos >= 20 And Xanos <= 44 Then
                                                    f20a44 = f20a44 + 1
                                                 Else
                                                    If Xanos >= 45 And Xanos <= 64 Then
                                                       f45a64 = f45a64 + 1
                                                    Else
                                                       If Xanos >= 65 And Xanos <= 74 Then
                                                          f65a74 = f65a74 + 1
                                                       Else
                                                          If Xanos >= 75 And Xanos <= 110 Then
                                                             f74 = f74 + 1
                                                          Else
                                                             fsd = fsd + 1
                                                          End If
                                                       End If
                                                    End If
                                                 End If
                                              End If
                                           End If
                                        End If
                                     End If
                                  End If
                               Else
                                  If data_cli.Recordset("cl_sexo") = 1 Then ' MASC
                                     If IsNull(data_cli.Recordset("cl_fnac")) = True Then
                                        msd = msd + 1
                                     Else
                                        Xanos = Date - data_cli.Recordset("cl_fnac")
                                        Xanos = Xanos / 365
                                        If Xanos < 1 Then
                                           m1 = m1 + 1
                                        Else
                                           If Xanos >= 1 And Xanos <= 4 Then
                                              m1a4 = m1a4 + 1
                                           Else
                                              If Xanos >= 5 And Xanos <= 14 Then
                                                 m5a14 = m5a14 + 1
                                              Else
                                                 If Xanos >= 15 And Xanos <= 19 Then
                                                    m15a19 = m15a19 + 1
                                                 Else
                                                    If Xanos >= 20 And Xanos <= 44 Then
                                                       m20a44 = m20a44 + 1
                                                    Else
                                                       If Xanos >= 45 And Xanos <= 64 Then
                                                          m45a64 = m45a64 + 1
                                                       Else
                                                          If Xanos >= 65 And Xanos <= 74 Then
                                                             m65a74 = m65a74 + 1
                                                          Else
                                                             If Xanos >= 75 And Xanos <= 110 Then
                                                                m74 = m74 + 1
                                                             Else
                                                                msd = msd + 1
                                                             End If
                                                          End If
                                                       End If
                                                    End If
                                                 End If
                                              End If
                                           End If
                                        End If
                                     End If
                                  Else
                                     Xsindatos = Xsindatos + 1
                                  End If
                               End If
                            End If
                    End If
               End If
                
            End If
         
         End If
      Else
         Xsindatos = Xsindatos + 1
      End If
      data_emi.Recordset.MoveNext
      pb.Value = pb.Value + 1
   Loop



   data_cli.RecordSource = "Select * from clientes where estado in (1) and cl_codconv not in ('PART','EMERN','SEMM','SEMM1','CCASMU','CASH','UDEMM')"
'    data_cli.RecordSource = "Select * from clientes where estado in (1,0) and cl_grupo >=" & 600 & " and cl_grupo <=" & 640
'    data_cli.RecordSource = "Select * from clientes where estado in (1,0) and cl_grupo =" & 999
   data_cli.Refresh
    
   If data_cli.Recordset.RecordCount > 0 Then
      data_cli.Recordset.MoveFirst
      DoEvents
       Do While Not data_cli.Recordset.EOF
          If IsNull(data_cli.Recordset("cl_codconv")) = False Then
             If data_cli.Recordset("cl_codconv") = "CASH" Or _
                data_cli.Recordset("cl_codconv") = "CPS" Or data_cli.Recordset("cl_codconv") = "UNIVR" Or _
                data_cli.Recordset("cl_codconv") = "SEMM1" Or data_cli.Recordset("cl_codconv") = "HEVANR" Or _
                data_cli.Recordset("cl_codconv") = "1727" Or data_cli.Recordset("cl_codconv") = "MUCAMM" Or _
                data_cli.Recordset("cl_codconv") = "CCNOS" Or data_cli.Recordset("cl_codconv") = "MUCATA" Or _
                data_cli.Recordset("cl_codconv") = "UNIVS" Or data_cli.Recordset("cl_codconv") = "MUCAMT" Or _
                data_cli.Recordset("cl_codconv") = "IMPNO" Or data_cli.Recordset("cl_codconv") = "MUCAMS" Or _
                data_cli.Recordset("cl_codconv") = "HEVAN" Or data_cli.Recordset("cl_codconv") = "MUCAMI" Or _
                data_cli.Recordset("cl_codconv") = "HEVANO" Or data_cli.Recordset("cl_codconv") = "MUCAMA" Or _
                data_cli.Recordset("cl_codconv") = "CCNRE" Or data_cli.Recordset("cl_codconv") = "CAAM" Or _
                data_cli.Recordset("cl_codconv") = "SMINR" Or data_cli.Recordset("cl_codconv") = "CAAMEP" Or _
                data_cli.Recordset("cl_codconv") = "SMIN" Or data_cli.Recordset("cl_codconv") = "CASANR" Or _
                data_cli.Recordset("cl_codconv") = "GANOS" Or _
                data_cli.Recordset("cl_codconv") = "1727B" Or data_cli.Recordset("cl_codconv") = "CASANO" Or _
                data_cli.Recordset("cl_codconv") = "1727C1" Or _
                data_cli.Recordset("cl_codconv") = "SEMM" Or _
                data_cli.Recordset("cl_codconv") = "911" Or _
                data_cli.Recordset("cl_codconv") = "911B" Or _
                data_cli.Recordset("cl_codconv") = "RETMI" Or _
                data_cli.Recordset("cl_codconv") = "MSP" Then
             Else
'                If data_cli.Recordset("cl_grupo") = 999 Then
'                   Xsindatosia = Xsindatosia + 1
                
'                Else
                    data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                    data_conv.Refresh
                    If data_conv.Recordset.RecordCount > 0 Then
                       If IsNull(data_conv.Recordset("cnv_grupo")) = False Then
                          If data_conv.Recordset("cnv_grupo") = "CASA DE GALICIA" Then
                             'If data_conv.Recordset("cnv_cant_r") = 2 And data_conv.Recordset("cnv_precio") > 0 Then
                             'Else
                                If IsNull(data_cli.Recordset("cl_sexo")) = True Then
                                   Xsindatosia = Xsindatosia + 1
                                   Grupo = "sd"
                                   data_inf.Recordset.AddNew
                                   data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                                   data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                                   data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                                   data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                                   data_inf.Recordset("estado") = data_cli.Recordset("estado")
                                   data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                                   data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                                   data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                   data_inf.Recordset("cl_nombre") = Grupo
                                   data_inf.Recordset.Update
                                Else
                                   If data_cli.Recordset("cl_sexo") = 2 Then ' FEM
                                      If IsNull(data_cli.Recordset("cl_fnac")) = True Then
                                         fsdia = fsdia + 1
                                         Grupo = "sd"
                                      Else
                                         Xanos = Date - data_cli.Recordset("cl_fnac")
                                         Xanos = Xanos / 365
                                         If Xanos < 1 Then
                                            f1ia = f1ia + 1
                                            Grupo = "1"
                                         Else
                                            If Xanos >= 1 And Xanos <= 4 Then
                                               f1a4ia = f1a4ia + 1
                                               Grupo = "2"
                                            Else
                                               If Xanos >= 5 And Xanos <= 14 Then
                                                  f5a14ia = f5a14ia + 1
                                                  Grupo = "3"
                                               Else
                                                  If Xanos >= 15 And Xanos <= 19 Then
                                                     f15a19ia = f15a19ia + 1
                                                     Grupo = "4"
                                                  Else
                                                     If Xanos >= 20 And Xanos <= 44 Then
                                                        f20a44ia = f20a44ia + 1
                                                        Grupo = "5"
                                                     Else
                                                        If Xanos >= 45 And Xanos <= 64 Then
                                                           f45a64ia = f45a64ia + 1
                                                           Grupo = "6"
                                                        Else
                                                           If Xanos >= 65 And Xanos <= 74 Then
                                                              f65a74ia = f65a74ia + 1
                                                              Grupo = "7"
                                                           Else
                                                              If Xanos >= 75 And Xanos <= 110 Then
                                                                 f74ia = f74ia + 1
                                                                 Grupo = "8"
                                                              Else
                                                                 fsdia = fsdia + 1
                                                                 Grupo = "sd"
                                                              End If
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      End If
                                         data_inf.Recordset.AddNew
                                         data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                                         data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                                         data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                                         data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                                         data_inf.Recordset("estado") = data_cli.Recordset("estado")
                                         data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                                         data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                                         data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                         data_inf.Recordset("cl_nombre") = Grupo
                                         data_inf.Recordset.Update
                                   Else
                                      If data_cli.Recordset("cl_sexo") = 1 Then ' MASC
                                         If IsNull(data_cli.Recordset("cl_fnac")) = True Then
                                            msdia = msdia + 1
                                            Grupo = "sd"
                                         Else
                                            Xanos = Date - data_cli.Recordset("cl_fnac")
                                            Xanos = Xanos / 365
                                            If Xanos < 1 Then
                                               m1ia = m1ia + 1
                                               Grupo = "1"
                                            Else
                                               If Xanos >= 1 And Xanos <= 4 Then
                                                  m1a4ia = m1a4ia + 1
                                                  Grupo = "2"
                                               Else
                                                  If Xanos >= 5 And Xanos <= 14 Then
                                                     m5a14ia = m5a14ia + 1
                                                     Grupo = "3"
                                                  Else
                                                     If Xanos >= 15 And Xanos <= 19 Then
                                                        m15a19ia = m15a19ia + 1
                                                        Grupo = "4"
                                                     Else
                                                        If Xanos >= 20 And Xanos <= 44 Then
                                                           m20a44ia = m20a44ia + 1
                                                           Grupo = "5"
                                                        Else
                                                           If Xanos >= 45 And Xanos <= 64 Then
                                                              m45a64ia = m45a64ia + 1
                                                              Grupo = "6"
                                                           Else
                                                              If Xanos >= 65 And Xanos <= 74 Then
                                                                 m65a74ia = m65a74ia + 1
                                                                 Grupo = "7"
                                                              Else
                                                                 If Xanos >= 75 And Xanos <= 110 Then
                                                                    m74ia = m74ia + 1
                                                                    Grupo = "8"
                                                                 Else
                                                                    msdia = msdia + 1
                                                                    Grupo = "sd"
                                                                 End If
                                                              End If
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      Else
                                         Xsindatosia = Xsindatosia + 1
                                         Grupo = "sd"
                                      End If
                                      data_inf.Recordset.AddNew
                                      data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                                      data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                                      data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                                      data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                                      data_inf.Recordset("estado") = data_cli.Recordset("estado")
                                      data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                                      data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                                      data_inf.Recordset("cl_sexo") = data_cli.Recordset("cl_sexo")
                                      data_inf.Recordset("cl_nombre") = Grupo
                                      data_inf.Recordset.Update
                                   End If
                                End If
                             'End If
                          End If
                       End If
                    End If
                ''End If
             End If
          End If
          data_cli.Recordset.MoveNext
          Grupo = ""
       Loop
   End If

End If


If Check4.Value = 1 Then
   Dim Xladesde As String
   Xladesde = InputBox("Ingrese a partir de que fecha, socios que consultaron al menos una vez", "Informes SINADI")
   Dim Xlaca As Double
   If Xladesde = "" Then
      Xladesde = "01/01/2012"
   End If
'   data_lin.RecordSource = "Select * from linmmdd where fecha >=#" & Format(Xladesde, "yyyy/mm/dd") & "# And fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and nro_flia in (1,10,14) order by cod_cli"
'   data_lin.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
'         data_lin.Recordset.FindFirst "cod_cli =" & Data1.Recordset("mat")
         data_lin.RecordSource = "Select * from linmmdd where cod_cli =" & Data1.Recordset("mat") & " and fecha >=#" & Format(Xladesde, "yyyy/mm/dd") & "# And fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and nro_flia in (1,10,14)"
         data_lin.Refresh
'         If Not data_lin.Recordset.NoMatch Then
         If data_lin.Recordset.RecordCount > 0 Then
            Xlaca = Xlaca + 1
         End If
         Data1.Recordset.MoveNext
      Loop
   End If
   MsgBox "Terminado...Cantidad: " & Trim(Str(Xlaca)), vbInformation, "Final"
   
End If


data_inf.DatabaseName = App.Path & "\informes.mdb"
data_inf.RecordSource = "inflla"
data_inf.Refresh

data_msp.RecordSource = "benef"
data_msp.Refresh
If data_msp.Recordset.RecordCount > 0 Then
   data_msp.Recordset.MoveFirst
   Do While Not data_msp.Recordset.EOF
      data_msp.Recordset.Delete
      data_msp.Recordset.MoveNext
   Loop
End If
data_msp.Recordset.AddNew
data_msp.Recordset("ano") = Year(mh.Text)
data_msp.Recordset("m1") = m1
data_msp.Recordset("m1a4") = m1a4
data_msp.Recordset("m5a14") = m5a14
data_msp.Recordset("m15a19") = m15a19
data_msp.Recordset("m20a44") = m20a44
data_msp.Recordset("m45a64") = m45a64
data_msp.Recordset("m65a74") = m65a74
data_msp.Recordset("m74") = m74
data_msp.Recordset("msd") = msd
data_msp.Recordset("f1") = f1
data_msp.Recordset("f1a4") = f1a4
data_msp.Recordset("f5a14") = f5a14
data_msp.Recordset("f15a19") = f15a19
data_msp.Recordset("f20a44") = f20a44
data_msp.Recordset("f45a64") = f45a64
data_msp.Recordset("f65a74") = f65a74
data_msp.Recordset("f74") = f74
data_msp.Recordset("fsd") = fsd

data_msp.Recordset("m1ia") = m1ia
data_msp.Recordset("m1a4ia") = m1a4ia
data_msp.Recordset("m5a14ia") = m5a14ia
data_msp.Recordset("m15a19ia") = m15a19ia
data_msp.Recordset("m20a44ia") = m20a44ia
data_msp.Recordset("m45a64ia") = m45a64ia
data_msp.Recordset("m65a74ia") = m65a74ia
data_msp.Recordset("m74ia") = m74ia
data_msp.Recordset("msdia") = msdia
data_msp.Recordset("f1ia") = f1ia
data_msp.Recordset("f1a4ia") = f1a4ia
data_msp.Recordset("f5a14ia") = f5a14ia
data_msp.Recordset("f15a19ia") = f15a19ia
data_msp.Recordset("f20a44ia") = f20a44ia
data_msp.Recordset("f45a64ia") = f45a64ia
data_msp.Recordset("f65a74ia") = f65a74ia
data_msp.Recordset("f74ia") = f74ia
data_msp.Recordset("fsdia") = fsdia
data_msp.Recordset("sd") = Xsindatos
data_msp.Recordset("sdia") = Xsindatosia

data_msp.Recordset.Update

MsgBox "Proceso de Beneficiarios, TERMINADO!!"
cr1.ReportFileName = App.Path & "\infbenef.rpt"
cr1.ReportTitle = "Al: " & mh.Text
cr1.Action = 1


End Sub

Private Sub Command10_Click()
crinf.ReportFileName = App.Path & "\infvtasms.rpt"
crinf.ReportTitle = "Informe de Ventas por servicio y por BASE FECHA:" & md.Text & " HASTA:" & mh.Text
crinf.Action = 1

End Sub

Private Sub Command2_Click()

Xcan1m = 0
Xcan14m = 0
Xcan514m = 0
Xcan1519m = 0
Xcan2044m = 0
Xcan4564m = 0
Xcan6574m = 0
Xcan74m = 0
Xcansdm = 0
Xcan1f = 0
Xcan14f = 0
Xcan514f = 0
Xcan1519f = 0
Xcan2044f = 0
Xcan4564f = 0
Xcan6574f = 0
Xcan74f = 0
Xcansdf = 0

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)

Set MiBaseact = Unasesact.OpenDatabase(App.Path & "\informes.mdb")

MiBaseact.Execute "Delete * from inflla"

data_lin.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codzon in (1,2,3) and cancela is null and base =" & 0 & " and categ not in ('CASH','SJ01','SJ02','SJ03','SJ04','SJ05')"
data_lin.Refresh
If data_lin.Recordset.RecordCount > 0 Then
   data_lin.Recordset.MoveFirst
   Do While Not data_lin.Recordset.EOF
      data_inf.Recordset.AddNew
      data_inf.Recordset("fecha") = data_lin.Recordset("fecha")
      data_inf.Recordset("hora") = data_lin.Recordset("hora")
      data_inf.Recordset("usuario") = data_lin.Recordset("usuario")
      data_inf.Recordset("nombre") = data_lin.Recordset("nombre")
      data_inf.Recordset("matric") = data_lin.Recordset("matric")
      data_inf.Recordset("categ") = data_lin.Recordset("categ")
      data_inf.Recordset("nomcat") = data_lin.Recordset("nomcat")
      data_inf.Recordset("edad") = data_lin.Recordset("edad")
      data_inf.Recordset("direcc") = data_lin.Recordset("direcc")
      data_inf.Recordset("telef") = data_lin.Recordset("telef")
      data_inf.Recordset("codzon") = data_lin.Recordset("codzon")
      data_inf.Recordset("motcon") = data_lin.Recordset("motcon")
      data_inf.Recordset("codmot") = data_lin.Recordset("codmot")
      data_inf.Recordset("movilpas") = data_lin.Recordset("movilpas")
      data_inf.Recordset("nro") = data_lin.Recordset("ncobr")
      data_inf.Recordset("movil_rea") = data_lin.Recordset("movil_rea")
      data_inf.Recordset("hor_rea") = data_lin.Recordset("hor_rea")
      data_inf.Recordset("fec_Rea") = data_lin.Recordset("fec_rea")
      data_inf.Recordset("movtras") = data_lin.Recordset("movtras")
      data_inf.Recordset("ci") = data_lin.Recordset("ci")
      If IsNull(data_lin.Recordset("colormot")) = False Then
         If data_lin.Recordset("colormot") = "" Then
            data_inf.Recordset("colormot") = data_lin.Recordset("codmot")
         Else
            data_inf.Recordset("colormot") = data_lin.Recordset("colormot")
         End If
      Else
         data_inf.Recordset("colormot") = data_lin.Recordset("codmot")
      End If
      data_inf.Recordset("lugar") = data_lin.Recordset("lugar")
      data_inf.Recordset("codmed") = data_lin.Recordset("codmed")
      data_inf.Recordset("nommed") = data_lin.Recordset("nommed")
      data_inf.Recordset("obs") = data_lin.Recordset("obs")
      data_inf.Recordset("diag") = data_lin.Recordset("diag")
      data_inf.Recordset("motmov") = data_lin.Recordset("motmov")
      data_inf.Recordset("mm") = data_lin.Recordset("mm")
      data_inf.Recordset("thh") = data_lin.Recordset("thh")
      data_inf.Recordset("tmm") = data_lin.Recordset("tmm")
      data_inf.Recordset("pasado") = data_lin.Recordset("pasado")
      data_inf.Recordset("ano") = data_lin.Recordset("ano")
      data_inf.Recordset("trasla") = data_lin.Recordset("trasla")
      data_inf.Recordset("mes") = data_lin.Recordset("mes")
      data_inf.Recordset("hsald") = data_lin.Recordset("hsald")
      data_inf.Recordset("hllega") = data_lin.Recordset("hllega")
      data_inf.Recordset("hzona") = data_lin.Recordset("hzona")
      data_inf.Recordset("hor_cance") = data_lin.Recordset("hor_cance")
      data_inf.Recordset("timdes") = data_lin.Recordset("timdes")
      data_inf.Recordset("descol") = data_lin.Recordset("descol")
      data_inf.Recordset("hh") = data_lin.Recordset("hh")
      data_inf.Recordset("cancela") = data_lin.Recordset("cancela")
      data_inf.Recordset("base") = data_lin.Recordset("base")
      data_inf.Recordset("unied") = data_lin.Recordset("unied")
      data_inf.Recordset.Update
      data_lin.Recordset.MoveNext
   Loop
   data_inf.Refresh
End If
'MsgBox "Uno"
data_lin2.DatabaseName = App.Path & "\informes.mdb"
 data_lin2.RecordSource = "Select * from inflla where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and codzon in (1,2,3) and base =" & 0 & " and colormot ='" & "R" & "'"
 data_lin2.Refresh
 If data_lin2.Recordset.RecordCount > 0 Then
    data_lin2.Recordset.MoveFirst
    Do While Not data_lin2.Recordset.EOF
       data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin2.Recordset("categ") & "' and cnv_emite ='" & "SI" & "' and (cnv_cant_r =" & 2 & " or cnv_cant_r=" & 1 & ") and (cnv_grupo is null or cnv_grupo='" & "" & "')"
       data_conv.Refresh
       If data_conv.Recordset.RecordCount > 0 Then
          If data_lin2.Recordset("movilpas") <> 99 Then
             If IsNull(data_lin2.Recordset("hh")) = True Then
                If data_lin2.Recordset("matric") > 0 Then
                   data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin2.Recordset("matric")
                   data_cli.Refresh
                   If data_cli.Recordset.RecordCount > 0 Then
                      If IsNull(data_cli.Recordset("cl_sexo")) = False Then
                         If data_cli.Recordset("cl_sexo") = 1 Then
                            If data_lin2.Recordset("unied") < 3 Then
                               Xcan1m = Xcan1m + 1
                            Else
                               If data_lin2.Recordset("edad") >= 1 And data_lin2.Recordset("edad") <= 4 Then
                                  Xcan14m = Xcan14m + 1
                               Else
                                  If data_lin2.Recordset("edad") >= 5 And data_lin2.Recordset("edad") <= 14 Then
                                     Xcan514m = Xcan514m + 1
                                  Else
                                     If data_lin2.Recordset("edad") >= 15 And data_lin2.Recordset("edad") <= 19 Then
                                        Xcan1519m = Xcan1519m + 1
                                     Else
                                        If data_lin2.Recordset("edad") >= 20 And data_lin2.Recordset("edad") <= 44 Then
                                           Xcan2044m = Xcan2044m + 1
                                        Else
                                           If data_lin2.Recordset("edad") >= 45 And data_lin2.Recordset("edad") <= 64 Then
                                              Xcan4564m = Xcan4564m + 1
                                           Else
                                              If data_lin2.Recordset("edad") >= 65 And data_lin2.Recordset("edad") <= 74 Then
                                                 Xcan6574m = Xcan6574m + 1
                                              Else
                                                 If data_lin2.Recordset("edad") >= 75 And data_lin2.Recordset("edad") <= 120 Then
                                                    Xcan74m = Xcan74m + 1
                                                 Else
                                                    Xcansdm = Xcansdm + 1
                                                 End If
                                              End If
                                           End If
                                        End If
                                     End If
                                  End If
                               End If
                            End If
                         Else
                            If data_lin2.Recordset("unied") < 3 Then
                               Xcan1f = Xcan1f + 1
                            Else
                               If data_lin2.Recordset("edad") >= 1 And data_lin2.Recordset("edad") <= 4 Then
                                  Xcan14f = Xcan14f + 1
                               Else
                                  If data_lin2.Recordset("edad") >= 5 And data_lin2.Recordset("edad") <= 14 Then
                                     Xcan514f = Xcan514f + 1
                                  Else
                                     If data_lin2.Recordset("edad") >= 15 And data_lin2.Recordset("edad") <= 19 Then
                                        Xcan1519f = Xcan1519f + 1
                                     Else
                                        If data_lin2.Recordset("edad") >= 20 And data_lin2.Recordset("edad") <= 44 Then
                                           Xcan2044f = Xcan2044f + 1
                                        Else
                                           If data_lin2.Recordset("edad") >= 45 And data_lin2.Recordset("edad") <= 64 Then
                                              Xcan4564f = Xcan4564f + 1
                                           Else
                                              If data_lin2.Recordset("edad") >= 65 And data_lin2.Recordset("edad") <= 74 Then
                                                 Xcan6574f = Xcan6574f + 1
                                              Else
                                                 If data_lin2.Recordset("edad") >= 75 And data_lin2.Recordset("edad") <= 120 Then
                                                    Xcan74f = Xcan74f + 1
                                                 Else
                                                    Xcansdf = Xcansdf + 1
                                                 End If
                                              End If
                                           End If
                                        End If
                                     End If
                                  End If
                               End If
                            End If
                         End If
                      Else
                         Xcansd = Xcansd + 1
                      End If
                   Else
                      Xcansd = Xcansd + 1
                   End If
                Else
                   Xcansd = Xcansd + 1
                End If
             Else
                If data_lin2.Recordset("hh") = 0 Then ' MASC
                   If data_lin2.Recordset("unied") < 3 Then
                      Xcan1m = Xcan1m + 1
                   Else
                      If data_lin2.Recordset("edad") >= 1 And data_lin2.Recordset("edad") <= 4 Then
                         Xcan14m = Xcan14m + 1
                      Else
                         If data_lin2.Recordset("edad") >= 5 And data_lin2.Recordset("edad") <= 14 Then
                            Xcan514m = Xcan514m + 1
                         Else
                            If data_lin2.Recordset("edad") >= 15 And data_lin2.Recordset("edad") <= 19 Then
                               Xcan1519m = Xcan1519m + 1
                            Else
                               If data_lin2.Recordset("edad") >= 20 And data_lin2.Recordset("edad") <= 44 Then
                                  Xcan2044m = Xcan2044m + 1
                               Else
                                  If data_lin2.Recordset("edad") >= 45 And data_lin2.Recordset("edad") <= 64 Then
                                     Xcan4564m = Xcan4564m + 1
                                  Else
                                    If data_lin2.Recordset("edad") >= 65 And data_lin2.Recordset("edad") <= 74 Then
                                       Xcan6574m = Xcan6574m + 1
                                    Else
                                       If data_lin2.Recordset("edad") >= 75 And data_lin2.Recordset("edad") <= 120 Then
                                          Xcan74m = Xcan74m + 1
                                       Else
                                          Xcansdm = Xcansdm + 1
                                       End If
                                    End If
                                  End If
                               End If
                            End If
                         End If
                      End If
                   End If
                Else
                   If data_lin2.Recordset("hh") = 1 Then ' FEM
                      If data_lin2.Recordset("unied") < 3 Then
                         Xcan1f = Xcan1f + 1
                      Else
                         If data_lin2.Recordset("edad") >= 1 And data_lin2.Recordset("edad") <= 4 Then
                            Xcan14f = Xcan14f + 1
                         Else
                            If data_lin2.Recordset("edad") >= 5 And data_lin2.Recordset("edad") <= 14 Then
                               Xcan514f = Xcan514f + 1
                            Else
                               If data_lin2.Recordset("edad") >= 15 And data_lin2.Recordset("edad") <= 19 Then
                                  Xcan1519f = Xcan1519f + 1
                               Else
                                  If data_lin2.Recordset("edad") >= 20 And data_lin2.Recordset("edad") <= 44 Then
                                     Xcan2044f = Xcan2044f + 1
                                  Else
                                     If data_lin2.Recordset("edad") >= 45 And data_lin2.Recordset("edad") <= 64 Then
                                        Xcan4564f = Xcan4564f + 1
                                     Else
                                        If data_lin2.Recordset("edad") >= 65 And data_lin2.Recordset("edad") <= 74 Then
                                           Xcan6574f = Xcan6574f + 1
                                        Else
                                           If data_lin2.Recordset("edad") >= 75 And data_lin2.Recordset("edad") <= 120 Then
                                              Xcan74f = Xcan74f + 1
                                           Else
                                              Xcansdf = Xcansdf + 1
                                           End If
                                        End If
                                     End If
                                  End If
                               End If
                            End If
                         End If
                      End If
                   Else
                      Xcansd = Xcansd + 1
                   End If
                End If
             End If
          End If
       Else
          If data_lin2.Recordset("categ") = "SEMM1" Or data_lin2.Recordset("categ") = "SEMM" Or _
             data_lin2.Recordset("categ") = "CERSEM" Or data_lin2.Recordset("categ") = "UDEMM" Or _
             data_lin2.Recordset("categ") = "UCM" Then
             If data_lin2.Recordset("movilpas") <> 99 Then
                If IsNull(data_lin2.Recordset("hh")) = True Then
                   If data_lin2.Recordset("matric") > 0 Then
                      data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin2.Recordset("matric")
                      data_cli.Refresh
                      If data_cli.Recordset.RecordCount > 0 Then
                         If IsNull(data_cli.Recordset("cl_sexo")) = False Then
                            If data_cli.Recordset("cl_sexo") = 1 Then
                               If data_lin2.Recordset("unied") < 3 Then
                                  Xcan1m = Xcan1m + 1
                               Else
                                  If data_lin2.Recordset("edad") >= 1 And data_lin2.Recordset("edad") <= 4 Then
                                     Xcan14m = Xcan14m + 1
                                  Else
                                     If data_lin2.Recordset("edad") >= 5 And data_lin2.Recordset("edad") <= 14 Then
                                        Xcan514m = Xcan514m + 1
                                     Else
                                        If data_lin2.Recordset("edad") >= 15 And data_lin2.Recordset("edad") <= 19 Then
                                           Xcan1519m = Xcan1519m + 1
                                        Else
                                           If data_lin2.Recordset("edad") >= 20 And data_lin2.Recordset("edad") <= 44 Then
                                              Xcan2044m = Xcan2044m + 1
                                           Else
                                              If data_lin2.Recordset("edad") >= 45 And data_lin2.Recordset("edad") <= 64 Then
                                                 Xcan4564m = Xcan4564m + 1
                                              Else
                                                 If data_lin2.Recordset("edad") >= 65 And data_lin2.Recordset("edad") <= 74 Then
                                                    Xcan6574m = Xcan6574m + 1
                                                 Else
                                                    If data_lin2.Recordset("edad") >= 75 And data_lin2.Recordset("edad") <= 120 Then
                                                       Xcan74m = Xcan74m + 1
                                                    Else
                                                       Xcansdm = Xcansdm + 1
                                                    End If
                                                 End If
                                              End If
                                           End If
                                        End If
                                     End If
                                  End If
                               End If
                            Else
                               If data_lin2.Recordset("unied") < 3 Then
                                  Xcan1f = Xcan1f + 1
                               Else
                                  If data_lin2.Recordset("edad") >= 1 And data_lin2.Recordset("edad") <= 4 Then
                                     Xcan14f = Xcan14f + 1
                                  Else
                                     If data_lin2.Recordset("edad") >= 5 And data_lin2.Recordset("edad") <= 14 Then
                                        Xcan514f = Xcan514f + 1
                                     Else
                                        If data_lin2.Recordset("edad") >= 15 And data_lin2.Recordset("edad") <= 19 Then
                                           Xcan1519f = Xcan1519f + 1
                                        Else
                                           If data_lin2.Recordset("edad") >= 20 And data_lin2.Recordset("edad") <= 44 Then
                                              Xcan2044f = Xcan2044f + 1
                                           Else
                                              If data_lin2.Recordset("edad") >= 45 And data_lin2.Recordset("edad") <= 64 Then
                                                 Xcan4564f = Xcan4564f + 1
                                              Else
                                                 If data_lin2.Recordset("edad") >= 65 And data_lin2.Recordset("edad") <= 74 Then
                                                    Xcan6574f = Xcan6574f + 1
                                                 Else
                                                    If data_lin2.Recordset("edad") >= 75 And data_lin2.Recordset("edad") <= 120 Then
                                                       Xcan74f = Xcan74f + 1
                                                    Else
                                                       Xcansdf = Xcansdf + 1
                                                    End If
                                                 End If
                                              End If
                                           End If
                                        End If
                                     End If
                                  End If
                               End If
                            End If
                         Else
                            Xcansd = Xcansd + 1
                         End If
                      Else
                         Xcansd = Xcansd + 1
                      End If
                   Else
                      Xcansd = Xcansd + 1
                   End If
                Else
                   If data_lin2.Recordset("hh") = 0 Then ' MASC
                      If data_lin2.Recordset("unied") < 3 Then
                         Xcan1m = Xcan1m + 1
                      Else
                         If data_lin2.Recordset("edad") >= 1 And data_lin2.Recordset("edad") <= 4 Then
                            Xcan14m = Xcan14m + 1
                         Else
                            If data_lin2.Recordset("edad") >= 5 And data_lin2.Recordset("edad") <= 14 Then
                               Xcan514m = Xcan514m + 1
                            Else
                               If data_lin2.Recordset("edad") >= 15 And data_lin2.Recordset("edad") <= 19 Then
                                  Xcan1519m = Xcan1519m + 1
                               Else
                                  If data_lin2.Recordset("edad") >= 20 And data_lin2.Recordset("edad") <= 44 Then
                                     Xcan2044m = Xcan2044m + 1
                                  Else
                                     If data_lin2.Recordset("edad") >= 45 And data_lin2.Recordset("edad") <= 64 Then
                                        Xcan4564m = Xcan4564m + 1
                                     Else
                                        If data_lin2.Recordset("edad") >= 65 And data_lin2.Recordset("edad") <= 74 Then
                                           Xcan6574m = Xcan6574m + 1
                                        Else
                                           If data_lin2.Recordset("edad") >= 75 And data_lin2.Recordset("edad") <= 120 Then
                                              Xcan74m = Xcan74m + 1
                                           Else
                                              Xcansdm = Xcansdm + 1
                                           End If
                                        End If
                                     End If
                                  End If
                               End If
                            End If
                         End If
                      End If
                   Else
                      If data_lin2.Recordset("hh") = 1 Then ' FEM
                         If data_lin2.Recordset("unied") < 3 Then
                            Xcan1f = Xcan1f + 1
                         Else
                            If data_lin2.Recordset("edad") >= 1 And data_lin2.Recordset("edad") <= 4 Then
                               Xcan14f = Xcan14f + 1
                            Else
                               If data_lin2.Recordset("edad") >= 5 And data_lin2.Recordset("edad") <= 14 Then
                                  Xcan514f = Xcan514f + 1
                               Else
                                  If data_lin2.Recordset("edad") >= 15 And data_lin2.Recordset("edad") <= 19 Then
                                     Xcan1519f = Xcan1519f + 1
                                  Else
                                     If data_lin2.Recordset("edad") >= 20 And data_lin2.Recordset("edad") <= 44 Then
                                        Xcan2044f = Xcan2044f + 1
                                     Else
                                        If data_lin2.Recordset("edad") >= 45 And data_lin2.Recordset("edad") <= 64 Then
                                           Xcan4564f = Xcan4564f + 1
                                        Else
                                           If data_lin2.Recordset("edad") >= 65 And data_lin2.Recordset("edad") <= 74 Then
                                              Xcan6574f = Xcan6574f + 1
                                           Else
                                              If data_lin2.Recordset("edad") >= 75 And data_lin2.Recordset("edad") <= 120 Then
                                                 Xcan74f = Xcan74f + 1
                                              Else
                                                 Xcansdf = Xcansdf + 1
                                              End If
                                           End If
                                        End If
                                     End If
                                  End If
                               End If
                            End If
                         End If
                      Else
                         Xcansd = Xcansd + 1
                      End If
                   End If
                End If
             End If
          End If
       End If
       data_lin2.Recordset.MoveNext
    Loop
 End If
 data_msp.Recordset.AddNew
 data_msp.Recordset("nro") = 2
 data_msp.Recordset("desc") = "ACTIVIDAD DOMICILIARIA"
 data_msp.Recordset("desc2") = "Segun clasificación luego de realizado"
 data_msp.Recordset("DESC4") = "EMERGENCIA"
 data_msp.Recordset("m1") = Xcan1m
 data_msp.Recordset("m1a4") = Xcan14m
 data_msp.Recordset("m5a14") = Xcan514m
 data_msp.Recordset("m15a19") = Xcan1519m
 data_msp.Recordset("m20a44") = Xcan2044m
 data_msp.Recordset("m45a64") = Xcan4564m
 data_msp.Recordset("m65a74") = Xcan6574m
 data_msp.Recordset("m74") = Xcan74m
 data_msp.Recordset("msd") = Xcansdm
 Xsubtott = Xcan1m + Xcan14m + Xcan514m + Xcan1519m + Xcan2044m + Xcan4564m + Xcan6574m + Xcan74m + Xcansdm
 data_msp.Recordset("mt") = Xsubtott
 data_msp.Recordset("f1") = Xcan1f
 data_msp.Recordset("f1a4") = Xcan14f
 data_msp.Recordset("f5a14") = Xcan514f
 data_msp.Recordset("f15a19") = Xcan1519f
 data_msp.Recordset("f20a44") = Xcan2044f
 data_msp.Recordset("f45a64") = Xcan4564f
 data_msp.Recordset("f65a74") = Xcan6574f
 data_msp.Recordset("f74") = Xcan74f
 data_msp.Recordset("fsd") = Xcansdf
 Xsubtottt = Xcan1f + Xcan14f + Xcan514f + Xcan1519f + Xcan2044f + Xcan4564f + Xcan6574f + Xcan74f + Xcansdf
 data_msp.Recordset("ft") = Xsubtottt
 data_msp.Recordset("tot") = Xsubtott + Xsubtottt
 data_msp.Recordset.Update
 data_msp.Refresh
 Xcan1m = 0
 Xcan14m = 0
 Xcan514m = 0
 Xcan1519m = 0
 Xcan2044m = 0
 Xcan4564m = 0
 Xcan6574m = 0
 Xcan74m = 0
 Xcansdm = 0
 Xsubtott = 0
 Xsubtott = 0
 Xcan1f = 0
 Xcan14f = 0
 Xcan514f = 0
 Xcan1519f = 0
 Xcan2044f = 0
 Xcan4564f = 0
 Xcan6574f = 0
 Xcan74f = 0
 Xcansdf = 0
 Xsubtottt = 0
''''' AMARILLOS
'MsgBox "Dos"
 data_lin2.RecordSource = "Select * from inflla where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and codzon in (1,2,3) and base =" & 0 & " and colormot ='" & "A" & "'"
 data_lin2.Refresh
 If data_lin2.Recordset.RecordCount > 0 Then
    data_lin2.Recordset.MoveFirst
    Do While Not data_lin2.Recordset.EOF
       data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin2.Recordset("categ") & "' and cnv_emite ='" & "SI" & "' and (cnv_cant_r =" & 2 & " or cnv_cant_r=" & 1 & ") and (cnv_grupo is null or cnv_grupo='" & "" & "')"
       data_conv.Refresh
       If data_conv.Recordset.RecordCount > 0 Then
          If data_lin2.Recordset("movilpas") <> 99 Then
             If IsNull(data_lin2.Recordset("hh")) = True Then
                If data_lin2.Recordset("matric") > 0 Then
                   data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin2.Recordset("matric")
                   data_cli.Refresh
                   If data_cli.Recordset.RecordCount > 0 Then
                      If IsNull(data_cli.Recordset("cl_sexo")) = False Then
                         If data_cli.Recordset("cl_sexo") = 1 Then
                            If data_lin2.Recordset("unied") < 3 Then
                               Xcan1m = Xcan1m + 1
                            Else
                               If data_lin2.Recordset("edad") >= 1 And data_lin2.Recordset("edad") <= 4 Then
                                  Xcan14m = Xcan14m + 1
                               Else
                                  If data_lin2.Recordset("edad") >= 5 And data_lin2.Recordset("edad") <= 14 Then
                                     Xcan514m = Xcan514m + 1
                                  Else
                                     If data_lin2.Recordset("edad") >= 15 And data_lin2.Recordset("edad") <= 19 Then
                                        Xcan1519m = Xcan1519m + 1
                                     Else
                                        If data_lin2.Recordset("edad") >= 20 And data_lin2.Recordset("edad") <= 44 Then
                                           Xcan2044m = Xcan2044m + 1
                                        Else
                                           If data_lin2.Recordset("edad") >= 45 And data_lin2.Recordset("edad") <= 64 Then
                                              Xcan4564m = Xcan4564m + 1
                                           Else
                                              If data_lin2.Recordset("edad") >= 65 And data_lin2.Recordset("edad") <= 74 Then
                                                 Xcan6574m = Xcan6574m + 1
                                              Else
                                                 If data_lin2.Recordset("edad") >= 75 And data_lin2.Recordset("edad") <= 120 Then
                                                    Xcan74m = Xcan74m + 1
                                                 Else
                                                    Xcansdm = Xcansdm + 1
                                                 End If
                                              End If
                                           End If
                                        End If
                                     End If
                                  End If
                               End If
                            End If
                         Else
                            If data_lin2.Recordset("unied") < 3 Then
                               Xcan1f = Xcan1f + 1
                            Else
                               If data_lin2.Recordset("edad") >= 1 And data_lin2.Recordset("edad") <= 4 Then
                                  Xcan14f = Xcan14f + 1
                               Else
                                  If data_lin2.Recordset("edad") >= 5 And data_lin2.Recordset("edad") <= 14 Then
                                     Xcan514f = Xcan514f + 1
                                  Else
                                     If data_lin2.Recordset("edad") >= 15 And data_lin2.Recordset("edad") <= 19 Then
                                        Xcan1519f = Xcan1519f + 1
                                     Else
                                        If data_lin2.Recordset("edad") >= 20 And data_lin2.Recordset("edad") <= 44 Then
                                           Xcan2044f = Xcan2044f + 1
                                        Else
                                           If data_lin2.Recordset("edad") >= 45 And data_lin2.Recordset("edad") <= 64 Then
                                              Xcan4564f = Xcan4564f + 1
                                           Else
                                              If data_lin2.Recordset("edad") >= 65 And data_lin2.Recordset("edad") <= 74 Then
                                                 Xcan6574f = Xcan6574f + 1
                                              Else
                                                 If data_lin2.Recordset("edad") >= 75 And data_lin2.Recordset("edad") <= 120 Then
                                                    Xcan74f = Xcan74f + 1
                                                 Else
                                                    Xcansdf = Xcansdf + 1
                                                 End If
                                              End If
                                           End If
                                        End If
                                     End If
                                  End If
                               End If
                            End If
                         End If
                      Else
                         Xcansd = Xcansd + 1
                      End If
                   Else
                      Xcansd = Xcansd + 1
                   End If
                Else
                   Xcansd = Xcansd + 1
                End If
             Else
                If data_lin2.Recordset("hh") = 0 Then ' MASC
                   If data_lin2.Recordset("unied") < 3 Then
                      Xcan1m = Xcan1m + 1
                   Else
                      If data_lin2.Recordset("edad") >= 1 And data_lin2.Recordset("edad") <= 4 Then
                         Xcan14m = Xcan14m + 1
                      Else
                         If data_lin2.Recordset("edad") >= 5 And data_lin2.Recordset("edad") <= 14 Then
                            Xcan514m = Xcan514m + 1
                         Else
                            If data_lin2.Recordset("edad") >= 15 And data_lin2.Recordset("edad") <= 19 Then
                               Xcan1519m = Xcan1519m + 1
                            Else
                               If data_lin2.Recordset("edad") >= 20 And data_lin2.Recordset("edad") <= 44 Then
                                  Xcan2044m = Xcan2044m + 1
                               Else
                                  If data_lin2.Recordset("edad") >= 45 And data_lin2.Recordset("edad") <= 64 Then
                                     Xcan4564m = Xcan4564m + 1
                                  Else
                                    If data_lin2.Recordset("edad") >= 65 And data_lin2.Recordset("edad") <= 74 Then
                                       Xcan6574m = Xcan6574m + 1
                                    Else
                                       If data_lin2.Recordset("edad") >= 75 And data_lin2.Recordset("edad") <= 120 Then
                                          Xcan74m = Xcan74m + 1
                                       Else
                                          Xcansdm = Xcansdm + 1
                                       End If
                                    End If
                                  End If
                               End If
                            End If
                         End If
                      End If
                   End If
                Else
                   If data_lin2.Recordset("hh") = 1 Then ' FEM
                      If data_lin2.Recordset("unied") < 3 Then
                         Xcan1f = Xcan1f + 1
                      Else
                         If data_lin2.Recordset("edad") >= 1 And data_lin2.Recordset("edad") <= 4 Then
                            Xcan14f = Xcan14f + 1
                         Else
                            If data_lin2.Recordset("edad") >= 5 And data_lin2.Recordset("edad") <= 14 Then
                               Xcan514f = Xcan514f + 1
                            Else
                               If data_lin2.Recordset("edad") >= 15 And data_lin2.Recordset("edad") <= 19 Then
                                  Xcan1519f = Xcan1519f + 1
                               Else
                                  If data_lin2.Recordset("edad") >= 20 And data_lin2.Recordset("edad") <= 44 Then
                                     Xcan2044f = Xcan2044f + 1
                                  Else
                                     If data_lin2.Recordset("edad") >= 45 And data_lin2.Recordset("edad") <= 64 Then
                                        Xcan4564f = Xcan4564f + 1
                                     Else
                                        If data_lin2.Recordset("edad") >= 65 And data_lin2.Recordset("edad") <= 74 Then
                                           Xcan6574f = Xcan6574f + 1
                                        Else
                                           If data_lin2.Recordset("edad") >= 75 And data_lin2.Recordset("edad") <= 120 Then
                                              Xcan74f = Xcan74f + 1
                                           Else
                                              Xcansdf = Xcansdf + 1
                                           End If
                                        End If
                                     End If
                                  End If
                               End If
                            End If
                         End If
                      End If
                   Else
                      Xcansd = Xcansd + 1
                   End If
                End If
             End If
          End If
       Else
          If data_lin2.Recordset("categ") = "SEMM1" Or data_lin2.Recordset("categ") = "SEMM" Or _
             data_lin2.Recordset("categ") = "CERSEM" Or data_lin2.Recordset("categ") = "UDEMM" Or _
             data_lin2.Recordset("categ") = "UCM" Then
             If data_lin2.Recordset("movilpas") <> 99 Then
                If IsNull(data_lin2.Recordset("hh")) = True Then
                   If data_lin2.Recordset("matric") > 0 Then
                      data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin2.Recordset("matric")
                      data_cli.Refresh
                      If data_cli.Recordset.RecordCount > 0 Then
                         If IsNull(data_cli.Recordset("cl_sexo")) = False Then
                            If data_cli.Recordset("cl_sexo") = 1 Then
                               If data_lin2.Recordset("unied") < 3 Then
                                  Xcan1m = Xcan1m + 1
                               Else
                                  If data_lin2.Recordset("edad") >= 1 And data_lin2.Recordset("edad") <= 4 Then
                                     Xcan14m = Xcan14m + 1
                                  Else
                                     If data_lin2.Recordset("edad") >= 5 And data_lin2.Recordset("edad") <= 14 Then
                                        Xcan514m = Xcan514m + 1
                                     Else
                                        If data_lin2.Recordset("edad") >= 15 And data_lin2.Recordset("edad") <= 19 Then
                                           Xcan1519m = Xcan1519m + 1
                                        Else
                                           If data_lin2.Recordset("edad") >= 20 And data_lin2.Recordset("edad") <= 44 Then
                                              Xcan2044m = Xcan2044m + 1
                                           Else
                                              If data_lin2.Recordset("edad") >= 45 And data_lin2.Recordset("edad") <= 64 Then
                                                 Xcan4564m = Xcan4564m + 1
                                              Else
                                                 If data_lin2.Recordset("edad") >= 65 And data_lin2.Recordset("edad") <= 74 Then
                                                    Xcan6574m = Xcan6574m + 1
                                                 Else
                                                    If data_lin2.Recordset("edad") >= 75 And data_lin2.Recordset("edad") <= 120 Then
                                                       Xcan74m = Xcan74m + 1
                                                    Else
                                                       Xcansdm = Xcansdm + 1
                                                    End If
                                                 End If
                                              End If
                                           End If
                                        End If
                                     End If
                                  End If
                               End If
                            Else
                               If data_lin2.Recordset("unied") < 3 Then
                                  Xcan1f = Xcan1f + 1
                               Else
                                  If data_lin2.Recordset("edad") >= 1 And data_lin2.Recordset("edad") <= 4 Then
                                     Xcan14f = Xcan14f + 1
                                  Else
                                     If data_lin2.Recordset("edad") >= 5 And data_lin2.Recordset("edad") <= 14 Then
                                        Xcan514f = Xcan514f + 1
                                     Else
                                        If data_lin2.Recordset("edad") >= 15 And data_lin2.Recordset("edad") <= 19 Then
                                           Xcan1519f = Xcan1519f + 1
                                        Else
                                           If data_lin2.Recordset("edad") >= 20 And data_lin2.Recordset("edad") <= 44 Then
                                              Xcan2044f = Xcan2044f + 1
                                           Else
                                              If data_lin2.Recordset("edad") >= 45 And data_lin2.Recordset("edad") <= 64 Then
                                                 Xcan4564f = Xcan4564f + 1
                                              Else
                                                 If data_lin2.Recordset("edad") >= 65 And data_lin2.Recordset("edad") <= 74 Then
                                                    Xcan6574f = Xcan6574f + 1
                                                 Else
                                                    If data_lin2.Recordset("edad") >= 75 And data_lin2.Recordset("edad") <= 120 Then
                                                       Xcan74f = Xcan74f + 1
                                                    Else
                                                       Xcansdf = Xcansdf + 1
                                                    End If
                                                 End If
                                              End If
                                           End If
                                        End If
                                     End If
                                  End If
                               End If
                            End If
                         Else
                            Xcansd = Xcansd + 1
                         End If
                      Else
                         Xcansd = Xcansd + 1
                      End If
                   Else
                      Xcansd = Xcansd + 1
                   End If
                Else
                   If data_lin2.Recordset("hh") = 0 Then ' MASC
                      If data_lin2.Recordset("unied") < 3 Then
                         Xcan1m = Xcan1m + 1
                      Else
                         If data_lin2.Recordset("edad") >= 1 And data_lin2.Recordset("edad") <= 4 Then
                            Xcan14m = Xcan14m + 1
                         Else
                            If data_lin2.Recordset("edad") >= 5 And data_lin2.Recordset("edad") <= 14 Then
                               Xcan514m = Xcan514m + 1
                            Else
                               If data_lin2.Recordset("edad") >= 15 And data_lin2.Recordset("edad") <= 19 Then
                                  Xcan1519m = Xcan1519m + 1
                               Else
                                  If data_lin2.Recordset("edad") >= 20 And data_lin2.Recordset("edad") <= 44 Then
                                     Xcan2044m = Xcan2044m + 1
                                  Else
                                     If data_lin2.Recordset("edad") >= 45 And data_lin2.Recordset("edad") <= 64 Then
                                        Xcan4564m = Xcan4564m + 1
                                     Else
                                        If data_lin2.Recordset("edad") >= 65 And data_lin2.Recordset("edad") <= 74 Then
                                           Xcan6574m = Xcan6574m + 1
                                        Else
                                           If data_lin2.Recordset("edad") >= 75 And data_lin2.Recordset("edad") <= 120 Then
                                              Xcan74m = Xcan74m + 1
                                           Else
                                              Xcansdm = Xcansdm + 1
                                           End If
                                        End If
                                     End If
                                  End If
                               End If
                            End If
                         End If
                      End If
                   Else
                      If data_lin2.Recordset("hh") = 1 Then ' FEM
                         If data_lin2.Recordset("unied") < 3 Then
                            Xcan1f = Xcan1f + 1
                         Else
                            If data_lin2.Recordset("edad") >= 1 And data_lin2.Recordset("edad") <= 4 Then
                               Xcan14f = Xcan14f + 1
                            Else
                               If data_lin2.Recordset("edad") >= 5 And data_lin2.Recordset("edad") <= 14 Then
                                  Xcan514f = Xcan514f + 1
                               Else
                                  If data_lin2.Recordset("edad") >= 15 And data_lin2.Recordset("edad") <= 19 Then
                                     Xcan1519f = Xcan1519f + 1
                                  Else
                                     If data_lin2.Recordset("edad") >= 20 And data_lin2.Recordset("edad") <= 44 Then
                                        Xcan2044f = Xcan2044f + 1
                                     Else
                                        If data_lin2.Recordset("edad") >= 45 And data_lin2.Recordset("edad") <= 64 Then
                                           Xcan4564f = Xcan4564f + 1
                                        Else
                                           If data_lin2.Recordset("edad") >= 65 And data_lin2.Recordset("edad") <= 74 Then
                                              Xcan6574f = Xcan6574f + 1
                                           Else
                                              If data_lin2.Recordset("edad") >= 75 And data_lin2.Recordset("edad") <= 120 Then
                                                 Xcan74f = Xcan74f + 1
                                              Else
                                                 Xcansdf = Xcansdf + 1
                                              End If
                                           End If
                                        End If
                                     End If
                                  End If
                               End If
                            End If
                         End If
                      Else
                         Xcansd = Xcansd + 1
                      End If
                   End If
                End If
             End If
          End If
       End If
       data_lin2.Recordset.MoveNext
    Loop
 End If
 data_msp.Recordset.MoveLast
 data_msp.Recordset.Edit
 data_msp.Recordset("DESC5") = "URGENCIA"
 data_msp.Recordset("m1a") = Xcan1m
 data_msp.Recordset("m1a4a") = Xcan14m
 data_msp.Recordset("m5a14a") = Xcan514m
 data_msp.Recordset("m15a19a") = Xcan1519m
 data_msp.Recordset("m20a44a") = Xcan2044m
 data_msp.Recordset("m45a64a") = Xcan4564m
 data_msp.Recordset("m65a74a") = Xcan6574m
 data_msp.Recordset("m74a") = Xcan74m
 data_msp.Recordset("msda") = Xcansdm
 Xsubtott = Xcan1m + Xcan14m + Xcan514m + Xcan1519m + Xcan2044m + Xcan4564m + Xcan6574m + Xcan74m + Xcansdm
 data_msp.Recordset("mta") = Xsubtott
 data_msp.Recordset("f1a") = Xcan1f
 data_msp.Recordset("f1a4a") = Xcan14f
 data_msp.Recordset("f5a14a") = Xcan514f
 data_msp.Recordset("f15a19a") = Xcan1519f
 data_msp.Recordset("f20a44a") = Xcan2044f
 data_msp.Recordset("f45a64a") = Xcan4564f
 data_msp.Recordset("f65a74a") = Xcan6574f
 data_msp.Recordset("f74a") = Xcan74f
 data_msp.Recordset("fsda") = Xcansdf
 Xsubtottt = Xcan1f + Xcan14f + Xcan514f + Xcan1519f + Xcan2044f + Xcan4564f + Xcan6574f + Xcan74f + Xcansdf
 data_msp.Recordset("fta") = Xsubtottt
 data_msp.Recordset("tota") = Xsubtott + Xsubtottt
 data_msp.Recordset.Update
 Xcan1m = 0
 Xcan14m = 0
 Xcan514m = 0
 Xcan1519m = 0
 Xcan2044m = 0
 Xcan4564m = 0
 Xcan6574m = 0
 Xcan74m = 0
 Xcansdm = 0
 Xsubtott = 0
 Xsubtott = 0
 Xcan1f = 0
 Xcan14f = 0
 Xcan514f = 0
 Xcan1519f = 0
 Xcan2044f = 0
 Xcan4564f = 0
 Xcan6574f = 0
 Xcan74f = 0
 Xcansdf = 0
 Xsubtottt = 0

'MsgBox "Tres"
 data_lin2.RecordSource = "Select * from inflla where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and codzon in (1,2,3) and base =" & 0 & " and colormot in ('V','Z')"
 data_lin2.Refresh
 If data_lin2.Recordset.RecordCount > 0 Then
    data_lin2.Recordset.MoveFirst
    Do While Not data_lin2.Recordset.EOF
        data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin2.Recordset("categ") & "' and cnv_emite ='" & "SI" & "' and (cnv_cant_r =" & 2 & " or cnv_cant_r=" & 1 & ") and (cnv_grupo is null or cnv_grupo='" & "" & "')"
        data_conv.Refresh
        If data_conv.Recordset.RecordCount > 0 Then
           If data_lin2.Recordset("movilpas") <> 99 Then
              If IsNull(data_lin2.Recordset("hh")) = True Then
                 If data_lin2.Recordset("matric") > 0 Then
                    data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin2.Recordset("matric")
                    data_cli.Refresh
                    If data_cli.Recordset.RecordCount > 0 Then
                       If IsNull(data_cli.Recordset("cl_sexo")) = False Then
                          If data_cli.Recordset("cl_sexo") = 1 Then
                             If data_lin2.Recordset("unied") < 3 Then
                                Xcan1m = Xcan1m + 1
                             Else
                                If data_lin2.Recordset("edad") >= 1 And data_lin2.Recordset("edad") <= 4 Then
                                   Xcan14m = Xcan14m + 1
                                Else
                                   If data_lin2.Recordset("edad") >= 5 And data_lin2.Recordset("edad") <= 14 Then
                                      Xcan514m = Xcan514m + 1
                                   Else
                                      If data_lin2.Recordset("edad") >= 15 And data_lin2.Recordset("edad") <= 19 Then
                                         Xcan1519m = Xcan1519m + 1
                                      Else
                                         If data_lin2.Recordset("edad") >= 20 And data_lin2.Recordset("edad") <= 44 Then
                                            Xcan2044m = Xcan2044m + 1
                                         Else
                                            If data_lin2.Recordset("edad") >= 45 And data_lin2.Recordset("edad") <= 64 Then
                                               Xcan4564m = Xcan4564m + 1
                                            Else
                                               If data_lin2.Recordset("edad") >= 65 And data_lin2.Recordset("edad") <= 74 Then
                                                  Xcan6574m = Xcan6574m + 1
                                               Else
                                                  If data_lin2.Recordset("edad") >= 75 And data_lin2.Recordset("edad") <= 120 Then
                                                     Xcan74m = Xcan74m + 1
                                                  Else
                                                     Xcansdm = Xcansdm + 1
                                                  End If
                                               End If
                                            End If
                                         End If
                                      End If
                                   End If
                                End If
                             End If
                          Else
                             If data_lin2.Recordset("unied") < 3 Then
                                Xcan1f = Xcan1f + 1
                             Else
                                If data_lin2.Recordset("edad") >= 1 And data_lin2.Recordset("edad") <= 4 Then
                                   Xcan14f = Xcan14f + 1
                                Else
                                   If data_lin2.Recordset("edad") >= 5 And data_lin2.Recordset("edad") <= 14 Then
                                      Xcan514f = Xcan514f + 1
                                   Else
                                      If data_lin2.Recordset("edad") >= 15 And data_lin2.Recordset("edad") <= 19 Then
                                         Xcan1519f = Xcan1519f + 1
                                      Else
                                         If data_lin2.Recordset("edad") >= 20 And data_lin2.Recordset("edad") <= 44 Then
                                            Xcan2044f = Xcan2044f + 1
                                         Else
                                            If data_lin2.Recordset("edad") >= 45 And data_lin2.Recordset("edad") <= 64 Then
                                               Xcan4564f = Xcan4564f + 1
                                            Else
                                               If data_lin2.Recordset("edad") >= 65 And data_lin2.Recordset("edad") <= 74 Then
                                                  Xcan6574f = Xcan6574f + 1
                                               Else
                                                  If data_lin2.Recordset("edad") >= 75 And data_lin2.Recordset("edad") <= 120 Then
                                                     Xcan74f = Xcan74f + 1
                                                  Else
                                                     Xcansdf = Xcansdf + 1
                                                  End If
                                               End If
                                            End If
                                         End If
                                      End If
                                   End If
                                End If
                             End If
                          End If
                       Else
                          Xcansd = Xcansd + 1
                       End If
                    Else
                       Xcansd = Xcansd + 1
                    End If
                 Else
                    Xcansd = Xcansd + 1
                 End If
              Else
                 If data_lin2.Recordset("hh") = 0 Then ' MASC
                    If data_lin2.Recordset("unied") < 3 Then
                       Xcan1m = Xcan1m + 1
                    Else
                       If data_lin2.Recordset("edad") >= 1 And data_lin2.Recordset("edad") <= 4 Then
                          Xcan14m = Xcan14m + 1
                       Else
                          If data_lin2.Recordset("edad") >= 5 And data_lin2.Recordset("edad") <= 14 Then
                             Xcan514m = Xcan514m + 1
                          Else
                             If data_lin2.Recordset("edad") >= 15 And data_lin2.Recordset("edad") <= 19 Then
                                Xcan1519m = Xcan1519m + 1
                             Else
                                If data_lin2.Recordset("edad") >= 20 And data_lin2.Recordset("edad") <= 44 Then
                                   Xcan2044m = Xcan2044m + 1
                                Else
                                   If data_lin2.Recordset("edad") >= 45 And data_lin2.Recordset("edad") <= 64 Then
                                      Xcan4564m = Xcan4564m + 1
                                   Else
                                     If data_lin2.Recordset("edad") >= 65 And data_lin2.Recordset("edad") <= 74 Then
                                        Xcan6574m = Xcan6574m + 1
                                     Else
                                        If data_lin2.Recordset("edad") >= 75 And data_lin2.Recordset("edad") <= 120 Then
                                           Xcan74m = Xcan74m + 1
                                        Else
                                           Xcansdm = Xcansdm + 1
                                        End If
                                     End If
                                   End If
                                End If
                             End If
                          End If
                       End If
                    End If
                 Else
                    If data_lin2.Recordset("hh") = 1 Then ' FEM
                       If data_lin2.Recordset("unied") < 3 Then
                          Xcan1f = Xcan1f + 1
                       Else
                          If data_lin2.Recordset("edad") >= 1 And data_lin2.Recordset("edad") <= 4 Then
                             Xcan14f = Xcan14f + 1
                          Else
                             If data_lin2.Recordset("edad") >= 5 And data_lin2.Recordset("edad") <= 14 Then
                                Xcan514f = Xcan514f + 1
                             Else
                                If data_lin2.Recordset("edad") >= 15 And data_lin2.Recordset("edad") <= 19 Then
                                   Xcan1519f = Xcan1519f + 1
                                Else
                                   If data_lin2.Recordset("edad") >= 20 And data_lin2.Recordset("edad") <= 44 Then
                                      Xcan2044f = Xcan2044f + 1
                                   Else
                                      If data_lin2.Recordset("edad") >= 45 And data_lin2.Recordset("edad") <= 64 Then
                                         Xcan4564f = Xcan4564f + 1
                                      Else
                                         If data_lin2.Recordset("edad") >= 65 And data_lin2.Recordset("edad") <= 74 Then
                                            Xcan6574f = Xcan6574f + 1
                                         Else
                                            If data_lin2.Recordset("edad") >= 75 And data_lin2.Recordset("edad") <= 120 Then
                                               Xcan74f = Xcan74f + 1
                                            Else
                                               Xcansdf = Xcansdf + 1
                                            End If
                                         End If
                                      End If
                                   End If
                                End If
                             End If
                          End If
                       End If
                    Else
                       Xcansd = Xcansd + 1
                    End If
                 End If
              End If
           End If
        Else
           If data_lin2.Recordset("categ") = "SEMM1" Or data_lin2.Recordset("categ") = "SEMM" Or _
              data_lin2.Recordset("categ") = "CERSEM" Or data_lin2.Recordset("categ") = "UDEMM" Or _
              data_lin2.Recordset("categ") = "UCM" Then
              If data_lin2.Recordset("movilpas") <> 99 Then
                 If IsNull(data_lin2.Recordset("hh")) = True Then
                    If data_lin2.Recordset("matric") > 0 Then
                       data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin2.Recordset("matric")
                       data_cli.Refresh
                       If data_cli.Recordset.RecordCount > 0 Then
                          If IsNull(data_cli.Recordset("cl_sexo")) = False Then
                             If data_cli.Recordset("cl_sexo") = 1 Then
                                If data_lin2.Recordset("unied") < 3 Then
                                   Xcan1m = Xcan1m + 1
                                Else
                                   If data_lin2.Recordset("edad") >= 1 And data_lin2.Recordset("edad") <= 4 Then
                                      Xcan14m = Xcan14m + 1
                                   Else
                                      If data_lin2.Recordset("edad") >= 5 And data_lin2.Recordset("edad") <= 14 Then
                                         Xcan514m = Xcan514m + 1
                                      Else
                                         If data_lin2.Recordset("edad") >= 15 And data_lin2.Recordset("edad") <= 19 Then
                                            Xcan1519m = Xcan1519m + 1
                                         Else
                                            If data_lin2.Recordset("edad") >= 20 And data_lin2.Recordset("edad") <= 44 Then
                                               Xcan2044m = Xcan2044m + 1
                                            Else
                                               If data_lin2.Recordset("edad") >= 45 And data_lin2.Recordset("edad") <= 64 Then
                                                  Xcan4564m = Xcan4564m + 1
                                               Else
                                                  If data_lin2.Recordset("edad") >= 65 And data_lin2.Recordset("edad") <= 74 Then
                                                     Xcan6574m = Xcan6574m + 1
                                                  Else
                                                     If data_lin2.Recordset("edad") >= 75 And data_lin2.Recordset("edad") <= 120 Then
                                                        Xcan74m = Xcan74m + 1
                                                     Else
                                                        Xcansdm = Xcansdm + 1
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      End If
                                   End If
                                End If
                             Else
                                If data_lin2.Recordset("unied") < 3 Then
                                   Xcan1f = Xcan1f + 1
                                Else
                                   If data_lin2.Recordset("edad") >= 1 And data_lin2.Recordset("edad") <= 4 Then
                                      Xcan14f = Xcan14f + 1
                                   Else
                                      If data_lin2.Recordset("edad") >= 5 And data_lin2.Recordset("edad") <= 14 Then
                                         Xcan514f = Xcan514f + 1
                                      Else
                                         If data_lin2.Recordset("edad") >= 15 And data_lin2.Recordset("edad") <= 19 Then
                                            Xcan1519f = Xcan1519f + 1
                                         Else
                                            If data_lin2.Recordset("edad") >= 20 And data_lin2.Recordset("edad") <= 44 Then
                                               Xcan2044f = Xcan2044f + 1
                                            Else
                                               If data_lin2.Recordset("edad") >= 45 And data_lin2.Recordset("edad") <= 64 Then
                                                  Xcan4564f = Xcan4564f + 1
                                               Else
                                                  If data_lin2.Recordset("edad") >= 65 And data_lin2.Recordset("edad") <= 74 Then
                                                     Xcan6574f = Xcan6574f + 1
                                                  Else
                                                     If data_lin2.Recordset("edad") >= 75 And data_lin2.Recordset("edad") <= 120 Then
                                                        Xcan74f = Xcan74f + 1
                                                     Else
                                                        Xcansdf = Xcansdf + 1
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      End If
                                   End If
                                End If
                             End If
                          Else
                             Xcansd = Xcansd + 1
                          End If
                       Else
                          Xcansd = Xcansd + 1
                       End If
                    Else
                       Xcansd = Xcansd + 1
                    End If
                 Else
                    If data_lin2.Recordset("hh") = 0 Then ' MASC
                       If data_lin2.Recordset("unied") < 3 Then
                          Xcan1m = Xcan1m + 1
                       Else
                          If data_lin2.Recordset("edad") >= 1 And data_lin2.Recordset("edad") <= 4 Then
                             Xcan14m = Xcan14m + 1
                          Else
                             If data_lin2.Recordset("edad") >= 5 And data_lin2.Recordset("edad") <= 14 Then
                                Xcan514m = Xcan514m + 1
                             Else
                                If data_lin2.Recordset("edad") >= 15 And data_lin2.Recordset("edad") <= 19 Then
                                   Xcan1519m = Xcan1519m + 1
                                Else
                                   If data_lin2.Recordset("edad") >= 20 And data_lin2.Recordset("edad") <= 44 Then
                                      Xcan2044m = Xcan2044m + 1
                                   Else
                                      If data_lin2.Recordset("edad") >= 45 And data_lin2.Recordset("edad") <= 64 Then
                                         Xcan4564m = Xcan4564m + 1
                                      Else
                                         If data_lin2.Recordset("edad") >= 65 And data_lin2.Recordset("edad") <= 74 Then
                                            Xcan6574m = Xcan6574m + 1
                                         Else
                                            If data_lin2.Recordset("edad") >= 75 And data_lin2.Recordset("edad") <= 120 Then
                                               Xcan74m = Xcan74m + 1
                                            Else
                                               Xcansdm = Xcansdm + 1
                                            End If
                                         End If
                                      End If
                                   End If
                                End If
                             End If
                          End If
                       End If
                    Else
                       If data_lin2.Recordset("hh") = 1 Then ' FEM
                          If data_lin2.Recordset("unied") < 3 Then
                             Xcan1f = Xcan1f + 1
                          Else
                             If data_lin2.Recordset("edad") >= 1 And data_lin2.Recordset("edad") <= 4 Then
                                Xcan14f = Xcan14f + 1
                             Else
                                If data_lin2.Recordset("edad") >= 5 And data_lin2.Recordset("edad") <= 14 Then
                                   Xcan514f = Xcan514f + 1
                                Else
                                   If data_lin2.Recordset("edad") >= 15 And data_lin2.Recordset("edad") <= 19 Then
                                      Xcan1519f = Xcan1519f + 1
                                   Else
                                      If data_lin2.Recordset("edad") >= 20 And data_lin2.Recordset("edad") <= 44 Then
                                         Xcan2044f = Xcan2044f + 1
                                      Else
                                         If data_lin2.Recordset("edad") >= 45 And data_lin2.Recordset("edad") <= 64 Then
                                            Xcan4564f = Xcan4564f + 1
                                         Else
                                            If data_lin2.Recordset("edad") >= 65 And data_lin2.Recordset("edad") <= 74 Then
                                               Xcan6574f = Xcan6574f + 1
                                            Else
                                               If data_lin2.Recordset("edad") >= 75 And data_lin2.Recordset("edad") <= 120 Then
                                                  Xcan74f = Xcan74f + 1
                                               Else
                                                  Xcansdf = Xcansdf + 1
                                               End If
                                            End If
                                         End If
                                      End If
                                   End If
                                End If
                             End If
                          End If
                       Else
                          Xcansd = Xcansd + 1
                       End If
                    End If
                 End If
              End If
           End If
        End If
       data_lin2.Recordset.MoveNext
     Loop
  End If
    
  data_msp.Recordset.Edit
  data_msp.Recordset("DESC6") = "RADIO"
  data_msp.Recordset("m1v") = Xcan1m
  data_msp.Recordset("m1a4v") = Xcan14m
  data_msp.Recordset("m5a14v") = Xcan514m
  data_msp.Recordset("m15a19v") = Xcan1519m
  data_msp.Recordset("m20a44v") = Xcan2044m
  data_msp.Recordset("m45a64v") = Xcan4564m
  data_msp.Recordset("m65a74v") = Xcan6574m
  data_msp.Recordset("m74v") = Xcan74m
  data_msp.Recordset("msdv") = Xcansdm
  Xsubtott = Xcan1m + Xcan14m + Xcan514m + Xcan1519m + Xcan2044m + Xcan4564m + Xcan6574m + Xcan74m + Xcansdm
  data_msp.Recordset("mtv") = Xsubtott
  data_msp.Recordset("f1v") = Xcan1f
  data_msp.Recordset("f1a4v") = Xcan14f
  data_msp.Recordset("f5a14v") = Xcan514f
  data_msp.Recordset("f15a19v") = Xcan1519f
  data_msp.Recordset("f20a44v") = Xcan2044f
  data_msp.Recordset("f45a64v") = Xcan4564f
  data_msp.Recordset("f65a74v") = Xcan6574f
  data_msp.Recordset("f74v") = Xcan74f
'  Xcansdf = Xcansdf + XcuentasinC
  data_msp.Recordset("fsdv") = Xcansdf
  Xsubtottt = Xcan1f + Xcan14f + Xcan514f + Xcan1519f + Xcan2044f + Xcan4564f + Xcan6574f + Xcan74f + Xcansdf
  data_msp.Recordset("ftv") = Xsubtottt
  data_msp.Recordset("totv") = Xsubtott + Xsubtottt
  data_msp.Recordset.Update

  Xcan1m = 0
  Xcan14m = 0
  Xcan514m = 0
  Xcan1519m = 0
  Xcan2044m = 0
  Xcan4564m = 0
  Xcan6574m = 0
  Xcan74m = 0
  Xcansdm = 0
  Xsubtott = 0
  Xsubtott = 0
  Xcan1f = 0
  Xcan14f = 0
  Xcan514f = 0
  Xcan1519f = 0
  Xcan2044f = 0
  Xcan4564f = 0
  Xcan6574f = 0
  Xcan74f = 0
  Xcansdf = 0
  Xsubtottt = 0

'  MsgBox "Termina el commando2"
  Command3_Click
  

End Sub

Private Sub Command3_Click()
Dim Xlaedad As Double
data_lin.ConnectionString = "dsn=" & Xconexrmt
'data_lin.DatabaseName = App.Path & "\sapp.mdb"
Xcan1m = 0
Xcan14m = 0
Xcan514m = 0
Xcan1519m = 0
Xcan2044m = 0
Xcan4564m = 0
Xcan6574m = 0
Xcan74m = 0
Xcansdm = 0
Xcan1f = 0
Xcan14f = 0
Xcan514f = 0
Xcan1519f = 0
Xcan2044f = 0
Xcan4564f = 0
Xcan6574f = 0
Xcan74f = 0
Xcansdf = 0
Dim Xsigrabar As Integer
Xsigrabar = 0

'MsgBox "Comienza el commando3"
'clave 2 y clave 1
'CP convenio emerg not
 
 data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cod_prod in (10003,10005) and convenio not in ('EMERN','EMERJ','EMERNE')"
 data_lin.Refresh
 If data_lin.Recordset.RecordCount > 0 Then
    data_lin.Recordset.MoveFirst
    Do While Not data_lin.Recordset.EOF
       data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("convenio") & "'"
       data_conv.Refresh
       If data_conv.Recordset.RecordCount > 0 Then
          If IsNull(data_conv.Recordset("cnv_grupo")) = False Then
             If data_conv.Recordset("cnv_grupo") = "" Then
                If data_conv.Recordset("cnv_cant_r") = 2 And data_conv.Recordset("cnv_precio") > 0 Then
                      If data_lin.Recordset("cod_cli") > 0 Then
                         data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("cod_cli")
                         data_cli.Refresh
                         If data_cli.Recordset.RecordCount > 0 Then
                               If IsNull(data_cli.Recordset("cl_sexo")) = False Then
                                  If data_cli.Recordset("cl_sexo") = 1 Then
                                     If IsNull(data_cli.Recordset("cl_fnac")) = False Then
                                        Xlaedad = CDate(mh.Text) - data_cli.Recordset("cl_fnac")
                                        Xlaedad = Xlaedad / 365
                                        If Xlaedad < 1 Then
                                           Xcan1m = Xcan1m + 1
                                           Xsigrabar = 9
                                        Else
                                           If Xlaedad >= 1 And Xlaedad <= 4 Then
                                              Xcan14m = Xcan14m + 1
                                              Xsigrabar = 9
                                           Else
                                              If Xlaedad >= 5 And Xlaedad <= 14 Then
                                                 Xcan514m = Xcan514m + 1
                                                 Xsigrabar = 9
                                              Else
                                                 If Xlaedad >= 15 And Xlaedad <= 19 Then
                                                    Xcan1519m = Xcan1519m + 1
                                                    Xsigrabar = 9
                                                 Else
                                                    If Xlaedad >= 20 And Xlaedad <= 44 Then
                                                       Xcan2044m = Xcan2044m + 1
                                                       Xsigrabar = 9
                                                    Else
                                                       If Xlaedad >= 45 And Xlaedad <= 64 Then
                                                          Xcan4564m = Xcan4564m + 1
                                                          Xsigrabar = 9
                                                       Else
                                                          If Xlaedad >= 65 And Xlaedad <= 74 Then
                                                             Xcan6574m = Xcan6574m + 1
                                                             Xsigrabar = 9
                                                          Else
                                                             If Xlaedad >= 75 And Xlaedad <= 120 Then
                                                                Xcan74m = Xcan74m + 1
                                                                Xsigrabar = 9
                                                             Else
                                                                Xcansdm = Xcansdm + 1
                                                                Xsigrabar = 9
                                                             End If
                                                          End If
                                                      End If
                                                    End If
                                                 End If
                                              End If
                                           End If
                                        End If
                                     Else
                                        Xcansdm = Xcansdm + 1
                                        Xsigrabar = 9
                                     End If
                                  Else
                                     If IsNull(data_cli.Recordset("cl_fnac")) = False Then
                                        Xlaedad = CDate(mh.Text) - data_cli.Recordset("cl_fnac")
                                        Xlaedad = Xlaedad / 365
                                        If Xlaedad < 1 Then
                                           Xcan1f = Xcan1f + 1
                                           Xsigrabar = 9
                                        Else
                                           If Xlaedad >= 1 And Xlaedad <= 4 Then
                                              Xcan14f = Xcan14f + 1
                                              Xsigrabar = 9
                                           Else
                                              If Xlaedad >= 5 And Xlaedad <= 14 Then
                                                 Xcan514f = Xcan514f + 1
                                                 Xsigrabar = 9
                                              Else
                                                 If Xlaedad >= 15 And Xlaedad <= 19 Then
                                                    Xcan1519f = Xcan1519f + 1
                                                    Xsigrabar = 9
                                                 Else
                                                    If Xlaedad >= 20 And Xlaedad <= 44 Then
                                                       Xcan2044f = Xcan2044f + 1
                                                       Xsigrabar = 9
                                                    Else
                                                       If Xlaedad >= 45 And Xlaedad <= 64 Then
                                                          Xcan4564f = Xcan4564f + 1
                                                          Xsigrabar = 9
                                                       Else
                                                          If Xlaedad >= 65 And Xlaedad <= 74 Then
                                                             Xcan6574f = Xcan6574f + 1
                                                             Xsigrabar = 9
                                                          Else
                                                             If Xlaedad >= 75 And Xlaedad <= 120 Then
                                                                Xcan74f = Xcan74f + 1
                                                                Xsigrabar = 9
                                                             Else
                                                                Xcansdf = Xcansdf + 1
                                                                Xsigrabar = 9
                                                             End If
                                                          End If
                                                       End If
                                                    End If
                                                 End If
                                              End If
                                           End If
                                        End If
                                     Else
                                        Xcansdf = Xcansdf + 1
                                        Xsigrabar = 9
                                     End If
                                  End If
                               Else
                                  Xcansd = Xcansd + 1
                                  Xsigrabar = 9
                               End If
                         Else
                            Xcansd = Xcansd + 1
                            Xsigrabar = 9
                         End If
                      Else
                          Xcansd = Xcansd + 1
                          Xsigrabar = 9
                      End If
                End If
             Else
                If data_conv.Recordset("cnv_grupo") = "UNIVERSAL" Or data_conv.Recordset("cnv_grupo") = "CCOU" Or _
                   data_conv.Recordset("cnv_grupo") = "SMI" Or data_conv.Recordset("cnv_grupo") = "CASA DE GALICIA" Or _
                   data_conv.Recordset("cnv_grupo") = "H.EVANGELICO" Then
                   If data_conv.Recordset("cnv_codigo") = "CCNOS" Or data_conv.Recordset("cnv_codigo") = "CASANO" Or _
                      data_conv.Recordset("cnv_codigo") = "UNIVS" Or data_conv.Recordset("cnv_codigo") = "IMPNO" Or _
                      data_conv.Recordset("cnv_codigo") = "HEVAN" Or data_conv.Recordset("cnv_codigo") = "HEVANO" Or _
                      data_conv.Recordset("cnv_codigo") = "CCNRE" Or data_conv.Recordset("cnv_codigo") = "SMIN" Or _
                      data_conv.Recordset("cnv_codigo") = "SMINR" Or data_conv.Recordset("cnv_codigo") = "GANOS" Then
                   Else
                      If data_lin.Recordset("cod_cli") > 0 Then
                         data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("cod_cli")
                         data_cli.Refresh
                         If data_cli.Recordset.RecordCount > 0 Then
                               If IsNull(data_cli.Recordset("cl_sexo")) = False Then
                                  If data_cli.Recordset("cl_sexo") = 1 Then
                                     If IsNull(data_cli.Recordset("cl_fnac")) = False Then
                                        Xlaedad = CDate(mh.Text) - data_cli.Recordset("cl_fnac")
                                        Xlaedad = Xlaedad / 365
                                        If Xlaedad < 1 Then
                                           Xcan1m = Xcan1m + 1
                                           Xsigrabar = 9
                                        Else
                                           If Xlaedad >= 1 And Xlaedad <= 4 Then
                                              Xcan14m = Xcan14m + 1
                                              Xsigrabar = 9
                                           Else
                                              If Xlaedad >= 5 And Xlaedad <= 14 Then
                                                 Xcan514m = Xcan514m + 1
                                                 Xsigrabar = 9
                                              Else
                                                 If Xlaedad >= 15 And Xlaedad <= 19 Then
                                                    Xcan1519m = Xcan1519m + 1
                                                    Xsigrabar = 9
                                                 Else
                                                    If Xlaedad >= 20 And Xlaedad <= 44 Then
                                                       Xcan2044m = Xcan2044m + 1
                                                       Xsigrabar = 9
                                                    Else
                                                       If Xlaedad >= 45 And Xlaedad <= 64 Then
                                                          Xcan4564m = Xcan4564m + 1
                                                          Xsigrabar = 9
                                                       Else
                                                          If Xlaedad >= 65 And Xlaedad <= 74 Then
                                                             Xcan6574m = Xcan6574m + 1
                                                             Xsigrabar = 9
                                                          Else
                                                             If Xlaedad >= 75 And Xlaedad <= 120 Then
                                                                Xcan74m = Xcan74m + 1
                                                                Xsigrabar = 9
                                                             Else
                                                                Xcansdm = Xcansdm + 1
                                                                Xsigrabar = 9
                                                             End If
                                                          End If
                                                      End If
                                                    End If
                                                 End If
                                              End If
                                           End If
                                        End If
                                     Else
                                        Xcansdm = Xcansdm + 1
                                        Xsigrabar = 9
                                     End If
                                  Else
                                     If IsNull(data_cli.Recordset("cl_fnac")) = False Then
                                        Xlaedad = CDate(mh.Text) - data_cli.Recordset("cl_fnac")
                                        Xlaedad = Xlaedad / 365
                                        If Xlaedad < 1 Then
                                           Xcan1f = Xcan1f + 1
                                           Xsigrabar = 9
                                        Else
                                           If Xlaedad >= 1 And Xlaedad <= 4 Then
                                              Xcan14f = Xcan14f + 1
                                              Xsigrabar = 9
                                           Else
                                              If Xlaedad >= 5 And Xlaedad <= 14 Then
                                                 Xcan514f = Xcan514f + 1
                                                 Xsigrabar = 9
                                              Else
                                                 If Xlaedad >= 15 And Xlaedad <= 19 Then
                                                    Xcan1519f = Xcan1519f + 1
                                                    Xsigrabar = 9
                                                 Else
                                                    If Xlaedad >= 20 And Xlaedad <= 44 Then
                                                       Xcan2044f = Xcan2044f + 1
                                                       Xsigrabar = 9
                                                    Else
                                                       If Xlaedad >= 45 And Xlaedad <= 64 Then
                                                          Xcan4564f = Xcan4564f + 1
                                                          Xsigrabar = 9
                                                       Else
                                                          If Xlaedad >= 65 And Xlaedad <= 74 Then
                                                             Xcan6574f = Xcan6574f + 1
                                                             Xsigrabar = 9
                                                          Else
                                                             If Xlaedad >= 75 And Xlaedad <= 120 Then
                                                                Xcan74f = Xcan74f + 1
                                                                Xsigrabar = 9
                                                             Else
                                                                Xcansdf = Xcansdf + 1
                                                                Xsigrabar = 9
                                                             End If
                                                          End If
                                                       End If
                                                    End If
                                                 End If
                                              End If
                                           End If
                                        End If
                                     Else
                                        Xcansdf = Xcansdf + 1
                                        Xsigrabar = 9
                                     End If
                                  End If
                               Else
                                  Xcansd = Xcansd + 1
                                  Xsigrabar = 9
                               End If
                         Else
                            Xcansd = Xcansd + 1
                            Xsigrabar = 9
                         End If
                      Else
                          Xcansd = Xcansd + 1
                          Xsigrabar = 9
                      End If
                   End If
                End If
             End If
          Else
             If data_conv.Recordset("cnv_cant_r") = 2 And data_conv.Recordset("cnv_precio") > 0 Then
                   If data_lin.Recordset("cod_cli") > 0 Then
                      data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("cod_cli")
                      data_cli.Refresh
                      If data_cli.Recordset.RecordCount > 0 Then
                            If IsNull(data_cli.Recordset("cl_sexo")) = False Then
                               If data_cli.Recordset("cl_sexo") = 1 Then
                                  If IsNull(data_cli.Recordset("cl_fnac")) = False Then
                                     Xlaedad = CDate(mh.Text) - data_cli.Recordset("cl_fnac")
                                     Xlaedad = Xlaedad / 365
                                     If Xlaedad < 1 Then
                                        Xcan1m = Xcan1m + 1
                                        Xsigrabar = 9
                                     Else
                                        If Xlaedad >= 1 And Xlaedad <= 4 Then
                                           Xcan14m = Xcan14m + 1
                                           Xsigrabar = 9
                                        Else
                                           If Xlaedad >= 5 And Xlaedad <= 14 Then
                                              Xcan514m = Xcan514m + 1
                                              Xsigrabar = 9
                                           Else
                                              If Xlaedad >= 15 And Xlaedad <= 19 Then
                                                 Xcan1519m = Xcan1519m + 1
                                                 Xsigrabar = 9
                                              Else
                                                 If Xlaedad >= 20 And Xlaedad <= 44 Then
                                                    Xcan2044m = Xcan2044m + 1
                                                    Xsigrabar = 9
                                                 Else
                                                    If Xlaedad >= 45 And Xlaedad <= 64 Then
                                                       Xcan4564m = Xcan4564m + 1
                                                       Xsigrabar = 9
                                                    Else
                                                       If Xlaedad >= 65 And Xlaedad <= 74 Then
                                                          Xcan6574m = Xcan6574m + 1
                                                          Xsigrabar = 9
                                                       Else
                                                          If Xlaedad >= 75 And Xlaedad <= 120 Then
                                                             Xcan74m = Xcan74m + 1
                                                             Xsigrabar = 9
                                                          Else
                                                             Xcansdm = Xcansdm + 1
                                                             Xsigrabar = 9
                                                          End If
                                                       End If
                                                   End If
                                                 End If
                                              End If
                                           End If
                                        End If
                                     End If
                                  Else
                                     Xcansdm = Xcansdm + 1
                                     Xsigrabar = 9
                                  End If
                               Else
                                  If IsNull(data_cli.Recordset("cl_fnac")) = False Then
                                     Xlaedad = CDate(mh.Text) - data_cli.Recordset("cl_fnac")
                                     Xlaedad = Xlaedad / 365
                                     If Xlaedad < 1 Then
                                        Xcan1f = Xcan1f + 1
                                        Xsigrabar = 9
                                     Else
                                        If Xlaedad >= 1 And Xlaedad <= 4 Then
                                           Xcan14f = Xcan14f + 1
                                           Xsigrabar = 9
                                        Else
                                           If Xlaedad >= 5 And Xlaedad <= 14 Then
                                              Xcan514f = Xcan514f + 1
                                              Xsigrabar = 9
                                           Else
                                              If Xlaedad >= 15 And Xlaedad <= 19 Then
                                                 Xcan1519f = Xcan1519f + 1
                                                 Xsigrabar = 9
                                              Else
                                                 If Xlaedad >= 20 And Xlaedad <= 44 Then
                                                    Xcan2044f = Xcan2044f + 1
                                                    Xsigrabar = 9
                                                 Else
                                                    If Xlaedad >= 45 And Xlaedad <= 64 Then
                                                       Xcan4564f = Xcan4564f + 1
                                                       Xsigrabar = 9
                                                    Else
                                                       If Xlaedad >= 65 And Xlaedad <= 74 Then
                                                          Xcan6574f = Xcan6574f + 1
                                                          Xsigrabar = 9
                                                       Else
                                                          If Xlaedad >= 75 And Xlaedad <= 120 Then
                                                             Xcan74f = Xcan74f + 1
                                                             Xsigrabar = 9
                                                          Else
                                                             Xcansdf = Xcansdf + 1
                                                             Xsigrabar = 9
                                                          End If
                                                       End If
                                                    End If
                                                 End If
                                              End If
                                           End If
                                        End If
                                     End If
                                  Else
                                     Xcansdf = Xcansdf + 1
                                     Xsigrabar = 9
                                  End If
                               End If
                            Else
                               Xcansd = Xcansd + 1
                               Xsigrabar = 9
                            End If
                      Else
                         Xcansd = Xcansd + 1
                         Xsigrabar = 9
                      End If
                   Else
                       Xcansd = Xcansd + 1
                       Xsigrabar = 9
                   End If
             End If
             
          End If
       End If
       If Xsigrabar = 9 Then
          If data_lin.Recordset("cod_prod") = 14001 Then
          Else
            data_tempg.Recordset.AddNew
            data_tempg.Recordset("fecha") = data_lin.Recordset("fecha")
            data_tempg.Recordset("base") = data_lin.Recordset("base")
            data_tempg.Recordset("cod_socio") = data_lin.Recordset("cod_cli")
            data_tempg.Recordset("nom_socio") = data_lin.Recordset("nom_cli")
            data_tempg.Recordset("cod_serv") = data_lin.Recordset("cod_prod")
            data_tempg.Recordset("nom_serv") = data_lin.Recordset("nom_prod")
            data_tempg.Recordset("imp_fact") = data_lin.Recordset("tot_lin")
            data_tempg.Recordset("usuario") = data_lin.Recordset("convenio")
            data_tempg.Recordset.Update
          End If
       End If
       Xsigrabar = 0
       data_lin.Recordset.MoveNext
    Loop
 End If
 '' EMERGENCIA Y URGENCIA
 data_msp.Recordset.AddNew
 data_msp.Recordset("nro") = 3
 data_msp.Recordset("desc") = "ACTIVIDAD CENTRALIZADA"
 data_msp.Recordset("desc2") = "Urgencia/Emergencia"
 data_msp.Recordset("desc4") = "URGENCIA/EMERGENCIA"
 data_msp.Recordset("m1") = Xcan1m
 data_msp.Recordset("m1a4") = Xcan14m
 data_msp.Recordset("m5a14") = Xcan514m
 data_msp.Recordset("m15a19") = Xcan1519m
 data_msp.Recordset("m20a44") = Xcan2044m
 data_msp.Recordset("m45a64") = Xcan4564m
 data_msp.Recordset("m65a74") = Xcan6574m
 data_msp.Recordset("m74") = Xcan74m
 data_msp.Recordset("msd") = Xcansdm
 Xsubtott = Xcan1m + Xcan14m + Xcan514m + Xcan1519m + Xcan2044m + Xcan4564m + Xcan6574m + Xcan74m + Xcansdm
 data_msp.Recordset("mt") = Xsubtott
 data_msp.Recordset("f1") = Xcan1f
 data_msp.Recordset("f1a4") = Xcan14f
 data_msp.Recordset("f5a14") = Xcan514f
 data_msp.Recordset("f15a19") = Xcan1519f
 data_msp.Recordset("f20a44") = Xcan2044f
 data_msp.Recordset("f45a64") = Xcan4564f
 data_msp.Recordset("f65a74") = Xcan6574f
 data_msp.Recordset("f74") = Xcan74f
 data_msp.Recordset("fsd") = Xcansdf
 Xsubtottt = Xcan1f + Xcan14f + Xcan514f + Xcan1519f + Xcan2044f + Xcan4564f + Xcan6574f + Xcan74f + Xcansdf
 data_msp.Recordset("ft") = Xsubtottt
 data_msp.Recordset("tot") = Xsubtott + Xsubtottt
 data_msp.Recordset.Update
 data_msp.Refresh
 Xcan1m = 0
 Xcan14m = 0
 Xcan514m = 0
 Xcan1519m = 0
 Xcan2044m = 0
 Xcan4564m = 0
 Xcan6574m = 0
 Xcan74m = 0
 Xcansdm = 0
 Xsubtott = 0
 Xsubtott = 0
 Xcan1f = 0
 Xcan14f = 0
 Xcan514f = 0
 Xcan1519f = 0
 Xcan2044f = 0
 Xcan4564f = 0
 Xcan6574f = 0
 Xcan74f = 0
 Xcansdf = 0
 Xsubtottt = 0


''''' POLICLINICA MED.GRAL. Y PEDIATRIA SIN PED CP 19/9/2018
 data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cod_prod in (10001) and convenio not in ('EMERN','EMERJ','EMERNE')"
 data_lin.Refresh
 If data_lin.Recordset.RecordCount > 0 Then
    data_lin.Recordset.MoveFirst
    Do While Not data_lin.Recordset.EOF
       data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("convenio") & "'"
       data_conv.Refresh
       If data_conv.Recordset.RecordCount > 0 Then
          If IsNull(data_conv.Recordset("cnv_grupo")) = False Then
             If data_conv.Recordset("cnv_grupo") = "" Then
                If data_conv.Recordset("cnv_cant_r") = 2 And data_conv.Recordset("cnv_precio") > 0 Then
                   If data_lin.Recordset("cod_cli") > 0 Then
                      data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("cod_cli")
                      data_cli.Refresh
                      If data_cli.Recordset.RecordCount > 0 Then
                           If IsNull(data_cli.Recordset("cl_sexo")) = False Then
                              If data_cli.Recordset("cl_sexo") = 1 Then
                                 If IsNull(data_cli.Recordset("cl_fnac")) = False Then
                                    Xlaedad = CDate(mh.Text) - data_cli.Recordset("cl_fnac")
                                    Xlaedad = Xlaedad / 365
                                    If Xlaedad < 1 Then
                                       Xcan1m = Xcan1m + 1
                                       Xsigrabar = 9
                                    Else
                                       If Xlaedad >= 1 And Xlaedad <= 4 Then
                                          Xcan14m = Xcan14m + 1
                                          Xsigrabar = 9
                                       Else
                                          If Xlaedad >= 5 And Xlaedad <= 14 Then
                                             Xcan514m = Xcan514m + 1
                                             Xsigrabar = 9
                                          Else
                                             If Xlaedad >= 15 And Xlaedad <= 19 Then
                                                Xcan1519m = Xcan1519m + 1
                                                Xsigrabar = 9
                                             Else
                                                If Xlaedad >= 20 And Xlaedad <= 44 Then
                                                   Xcan2044m = Xcan2044m + 1
                                                   Xsigrabar = 9
                                                Else
                                                   If Xlaedad >= 45 And Xlaedad <= 64 Then
                                                      Xcan4564m = Xcan4564m + 1
                                                      Xsigrabar = 9
                                                   Else
                                                      If Xlaedad >= 65 And Xlaedad <= 74 Then
                                                         Xcan6574m = Xcan6574m + 1
                                                         Xsigrabar = 9
                                                      Else
                                                         If Xlaedad >= 75 And Xlaedad <= 120 Then
                                                            Xcan74m = Xcan74m + 1
                                                            Xsigrabar = 9
                                                         Else
                                                            Xcansdm = Xcansdm + 1
                                                            Xsigrabar = 9
                                                         End If
                                                      End If
                                                  End If
                                                End If
                                             End If
                                          End If
                                       End If
                                    End If
                                 Else
                                    Xcansdm = Xcansdm + 1
                                    Xsigrabar = 9
                                 End If
                              Else
                                 If IsNull(data_cli.Recordset("cl_fnac")) = False Then
                                    Xlaedad = CDate(mh.Text) - data_cli.Recordset("cl_fnac")
                                    Xlaedad = Xlaedad / 365
                                    If Xlaedad < 1 Then
                                       Xcan1f = Xcan1f + 1
                                       Xsigrabar = 9
                                    Else
                                       If Xlaedad >= 1 And Xlaedad <= 4 Then
                                          Xcan14f = Xcan14f + 1
                                          Xsigrabar = 9
                                       Else
                                          If Xlaedad >= 5 And Xlaedad <= 14 Then
                                             Xcan514f = Xcan514f + 1
                                             Xsigrabar = 9
                                          Else
                                             If Xlaedad >= 15 And Xlaedad <= 19 Then
                                                Xcan1519f = Xcan1519f + 1
                                                Xsigrabar = 9
                                             Else
                                                If Xlaedad >= 20 And Xlaedad <= 44 Then
                                                   Xcan2044f = Xcan2044f + 1
                                                   Xsigrabar = 9
                                                Else
                                                   If Xlaedad >= 45 And Xlaedad <= 64 Then
                                                      Xcan4564f = Xcan4564f + 1
                                                      Xsigrabar = 9
                                                   Else
                                                      If Xlaedad >= 65 And Xlaedad <= 74 Then
                                                         Xcan6574f = Xcan6574f + 1
                                                         Xsigrabar = 9
                                                      Else
                                                         If Xlaedad >= 75 And Xlaedad <= 120 Then
                                                            Xcan74f = Xcan74f + 1
                                                            Xsigrabar = 9
                                                         Else
                                                            Xcansdf = Xcansdf + 1
                                                            Xsigrabar = 9
                                                         End If
                                                      End If
                                                   End If
                                                End If
                                             End If
                                          End If
                                       End If
                                    End If
                                 Else
                                    Xcansdf = Xcansdf + 1
                                    Xsigrabar = 9
                                 End If
                              End If
                           Else
                              Xcansd = Xcansd + 1
                              Xsigrabar = 9
                           End If
                      End If
                   End If
                End If
             Else
                If data_conv.Recordset("cnv_grupo") = "UNIVERSAL" Or data_conv.Recordset("cnv_grupo") = "CCOU" Or _
                   data_conv.Recordset("cnv_grupo") = "SMI" Or data_conv.Recordset("cnv_grupo") = "CASA DE GALICIA" Or _
                   data_conv.Recordset("cnv_grupo") = "H.EVANGELICO" Then
                   If data_conv.Recordset("cnv_codigo") = "CCNOS" Or data_conv.Recordset("cnv_codigo") = "CASANO" Or _
                      data_conv.Recordset("cnv_codigo") = "UNIVS" Or data_conv.Recordset("cnv_codigo") = "IMPNO" Or _
                      data_conv.Recordset("cnv_codigo") = "HEVAN" Or data_conv.Recordset("cnv_codigo") = "HEVANO" Or _
                      data_conv.Recordset("cnv_codigo") = "CCNRE" Or data_conv.Recordset("cnv_codigo") = "SMIN" Or _
                      data_conv.Recordset("cnv_codigo") = "SMINR" Or data_conv.Recordset("cnv_codigo") = "GANOS" Then
                   Else
                        If data_lin.Recordset("cod_cli") > 0 Then
                           data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("cod_cli")
                           data_cli.Refresh
                           If data_cli.Recordset.RecordCount > 0 Then
                                If IsNull(data_cli.Recordset("cl_sexo")) = False Then
                                   If data_cli.Recordset("cl_sexo") = 1 Then
                                      If IsNull(data_cli.Recordset("cl_fnac")) = False Then
                                         Xlaedad = CDate(mh.Text) - data_cli.Recordset("cl_fnac")
                                         Xlaedad = Xlaedad / 365
                                         If Xlaedad < 1 Then
                                            Xcan1m = Xcan1m + 1
                                            Xsigrabar = 9
                                         Else
                                            If Xlaedad >= 1 And Xlaedad <= 4 Then
                                               Xcan14m = Xcan14m + 1
                                               Xsigrabar = 9
                                            Else
                                               If Xlaedad >= 5 And Xlaedad <= 14 Then
                                                  Xcan514m = Xcan514m + 1
                                                  Xsigrabar = 9
                                               Else
                                                  If Xlaedad >= 15 And Xlaedad <= 19 Then
                                                     Xcan1519m = Xcan1519m + 1
                                                     Xsigrabar = 9
                                                  Else
                                                     If Xlaedad >= 20 And Xlaedad <= 44 Then
                                                        Xcan2044m = Xcan2044m + 1
                                                        Xsigrabar = 9
                                                     Else
                                                        If Xlaedad >= 45 And Xlaedad <= 64 Then
                                                           Xcan4564m = Xcan4564m + 1
                                                           Xsigrabar = 9
                                                        Else
                                                           If Xlaedad >= 65 And Xlaedad <= 74 Then
                                                              Xcan6574m = Xcan6574m + 1
                                                              Xsigrabar = 9
                                                           Else
                                                              If Xlaedad >= 75 And Xlaedad <= 120 Then
                                                                 Xcan74m = Xcan74m + 1
                                                                 Xsigrabar = 9
                                                              Else
                                                                 Xcansdm = Xcansdm + 1
                                                                 Xsigrabar = 9
                                                              End If
                                                           End If
                                                       End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      Else
                                         Xcansdm = Xcansdm + 1
                                         Xsigrabar = 9
                                      End If
                                   Else
                                      If IsNull(data_cli.Recordset("cl_fnac")) = False Then
                                         Xlaedad = CDate(mh.Text) - data_cli.Recordset("cl_fnac")
                                         Xlaedad = Xlaedad / 365
                                         If Xlaedad < 1 Then
                                            Xcan1f = Xcan1f + 1
                                            Xsigrabar = 9
                                         Else
                                            If Xlaedad >= 1 And Xlaedad <= 4 Then
                                               Xcan14f = Xcan14f + 1
                                               Xsigrabar = 9
                                            Else
                                               If Xlaedad >= 5 And Xlaedad <= 14 Then
                                                  Xcan514f = Xcan514f + 1
                                                  Xsigrabar = 9
                                               Else
                                                  If Xlaedad >= 15 And Xlaedad <= 19 Then
                                                     Xcan1519f = Xcan1519f + 1
                                                     Xsigrabar = 9
                                                  Else
                                                     If Xlaedad >= 20 And Xlaedad <= 44 Then
                                                        Xcan2044f = Xcan2044f + 1
                                                        Xsigrabar = 9
                                                     Else
                                                        If Xlaedad >= 45 And Xlaedad <= 64 Then
                                                           Xcan4564f = Xcan4564f + 1
                                                           Xsigrabar = 9
                                                        Else
                                                           If Xlaedad >= 65 And Xlaedad <= 74 Then
                                                              Xcan6574f = Xcan6574f + 1
                                                              Xsigrabar = 9
                                                           Else
                                                              If Xlaedad >= 75 And Xlaedad <= 120 Then
                                                                 Xcan74f = Xcan74f + 1
                                                                 Xsigrabar = 9
                                                              Else
                                                                 Xcansdf = Xcansdf + 1
                                                                 Xsigrabar = 9
                                                              End If
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      Else
                                         Xcansdf = Xcansdf + 1
                                         Xsigrabar = 9
                                      End If
                                   End If
                                Else
                                   Xcansd = Xcansd + 1
                                   Xsigrabar = 9
                                End If
                           End If
                        End If
                   End If
                End If
             End If
          Else
             If data_conv.Recordset("cnv_cant_r") = 2 And data_conv.Recordset("cnv_precio") > 0 Then
                If data_lin.Recordset("cod_cli") > 0 Then
                   data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("cod_cli")
                   data_cli.Refresh
                   If data_cli.Recordset.RecordCount > 0 Then
                        If IsNull(data_cli.Recordset("cl_sexo")) = False Then
                           If data_cli.Recordset("cl_sexo") = 1 Then
                              If IsNull(data_cli.Recordset("cl_fnac")) = False Then
                                 Xlaedad = CDate(mh.Text) - data_cli.Recordset("cl_fnac")
                                 Xlaedad = Xlaedad / 365
                                 If Xlaedad < 1 Then
                                    Xcan1m = Xcan1m + 1
                                    Xsigrabar = 9
                                 Else
                                    If Xlaedad >= 1 And Xlaedad <= 4 Then
                                       Xcan14m = Xcan14m + 1
                                       Xsigrabar = 9
                                    Else
                                       If Xlaedad >= 5 And Xlaedad <= 14 Then
                                          Xcan514m = Xcan514m + 1
                                          Xsigrabar = 9
                                       Else
                                          If Xlaedad >= 15 And Xlaedad <= 19 Then
                                             Xcan1519m = Xcan1519m + 1
                                             Xsigrabar = 9
                                          Else
                                             If Xlaedad >= 20 And Xlaedad <= 44 Then
                                                Xcan2044m = Xcan2044m + 1
                                                Xsigrabar = 9
                                             Else
                                                If Xlaedad >= 45 And Xlaedad <= 64 Then
                                                   Xcan4564m = Xcan4564m + 1
                                                   Xsigrabar = 9
                                                Else
                                                   If Xlaedad >= 65 And Xlaedad <= 74 Then
                                                      Xcan6574m = Xcan6574m + 1
                                                      Xsigrabar = 9
                                                   Else
                                                      If Xlaedad >= 75 And Xlaedad <= 120 Then
                                                         Xcan74m = Xcan74m + 1
                                                         Xsigrabar = 9
                                                      Else
                                                         Xcansdm = Xcansdm + 1
                                                         Xsigrabar = 9
                                                      End If
                                                   End If
                                               End If
                                             End If
                                          End If
                                       End If
                                    End If
                                 End If
                              Else
                                 Xcansdm = Xcansdm + 1
                                 Xsigrabar = 9
                              End If
                           Else
                              If IsNull(data_cli.Recordset("cl_fnac")) = False Then
                                 Xlaedad = CDate(mh.Text) - data_cli.Recordset("cl_fnac")
                                 Xlaedad = Xlaedad / 365
                                 If Xlaedad < 1 Then
                                    Xcan1f = Xcan1f + 1
                                    Xsigrabar = 9
                                 Else
                                    If Xlaedad >= 1 And Xlaedad <= 4 Then
                                       Xcan14f = Xcan14f + 1
                                       Xsigrabar = 9
                                    Else
                                       If Xlaedad >= 5 And Xlaedad <= 14 Then
                                          Xcan514f = Xcan514f + 1
                                          Xsigrabar = 9
                                       Else
                                          If Xlaedad >= 15 And Xlaedad <= 19 Then
                                             Xcan1519f = Xcan1519f + 1
                                             Xsigrabar = 9
                                          Else
                                             If Xlaedad >= 20 And Xlaedad <= 44 Then
                                                Xcan2044f = Xcan2044f + 1
                                                Xsigrabar = 9
                                             Else
                                                If Xlaedad >= 45 And Xlaedad <= 64 Then
                                                   Xcan4564f = Xcan4564f + 1
                                                   Xsigrabar = 9
                                                Else
                                                   If Xlaedad >= 65 And Xlaedad <= 74 Then
                                                      Xcan6574f = Xcan6574f + 1
                                                      Xsigrabar = 9
                                                   Else
                                                      If Xlaedad >= 75 And Xlaedad <= 120 Then
                                                         Xcan74f = Xcan74f + 1
                                                         Xsigrabar = 9
                                                      Else
                                                         Xcansdf = Xcansdf + 1
                                                         Xsigrabar = 9
                                                      End If
                                                   End If
                                                End If
                                             End If
                                          End If
                                       End If
                                    End If
                                 End If
                              Else
                                 Xcansdf = Xcansdf + 1
                                 Xsigrabar = 9
                              End If
                           End If
                        Else
                           Xcansd = Xcansd + 1
                           Xsigrabar = 9
                        End If
                   End If
                End If
             End If
          
          End If
       End If
       If Xsigrabar = 9 Then
          If data_lin.Recordset("cod_prod") = 14001 Then
          Else
            data_tempg.Recordset.AddNew
            data_tempg.Recordset("fecha") = data_lin.Recordset("fecha")
            data_tempg.Recordset("base") = data_lin.Recordset("base")
            data_tempg.Recordset("cod_socio") = data_lin.Recordset("cod_cli")
            data_tempg.Recordset("nom_socio") = data_lin.Recordset("nom_cli")
            data_tempg.Recordset("cod_serv") = data_lin.Recordset("cod_prod")
            data_tempg.Recordset("nom_serv") = data_lin.Recordset("nom_prod")
            data_tempg.Recordset("imp_fact") = data_lin.Recordset("tot_lin")
            data_tempg.Recordset("usuario") = data_lin.Recordset("convenio")
            data_tempg.Recordset.Update
          End If
       End If
       Xsigrabar = 0
       data_lin.Recordset.MoveNext
    Loop
 End If
 data_msp.Recordset.MoveLast
 data_msp.Recordset.Edit
 data_msp.Recordset("DESC5") = "MED.GRAL/PEDIATRIA"
 data_msp.Recordset("m1a") = Xcan1m
 data_msp.Recordset("m1a4a") = Xcan14m
 data_msp.Recordset("m5a14a") = Xcan514m
 data_msp.Recordset("m15a19a") = Xcan1519m
 data_msp.Recordset("m20a44a") = Xcan2044m
 data_msp.Recordset("m45a64a") = Xcan4564m
 data_msp.Recordset("m65a74a") = Xcan6574m
 data_msp.Recordset("m74a") = Xcan74m
 data_msp.Recordset("msda") = Xcansdm
 Xsubtott = Xcan1m + Xcan14m + Xcan514m + Xcan1519m + Xcan2044m + Xcan4564m + Xcan6574m + Xcan74m + Xcansdm
 data_msp.Recordset("mta") = Xsubtott
 data_msp.Recordset("f1a") = Xcan1f
 data_msp.Recordset("f1a4a") = Xcan14f
 data_msp.Recordset("f5a14a") = Xcan514f
 data_msp.Recordset("f15a19a") = Xcan1519f
 data_msp.Recordset("f20a44a") = Xcan2044f
 data_msp.Recordset("f45a64a") = Xcan4564f
 data_msp.Recordset("f65a74a") = Xcan6574f
 data_msp.Recordset("f74a") = Xcan74f
 data_msp.Recordset("fsda") = Xcansdf
 Xsubtottt = Xcan1f + Xcan14f + Xcan514f + Xcan1519f + Xcan2044f + Xcan4564f + Xcan6574f + Xcan74f + Xcansdf
 data_msp.Recordset("fta") = Xsubtottt
 data_msp.Recordset("tota") = Xsubtott + Xsubtottt
 data_msp.Recordset.Update
 Xcan1m = 0
 Xcan14m = 0
 Xcan514m = 0
 Xcan1519m = 0
 Xcan2044m = 0
 Xcan4564m = 0
 Xcan6574m = 0
 Xcan74m = 0
 Xcansdm = 0
 Xsubtott = 0
 Xsubtott = 0
 Xcan1f = 0
 Xcan14f = 0
 Xcan514f = 0
 Xcan1519f = 0
 Xcan2044f = 0
 Xcan4564f = 0
 Xcan6574f = 0
 Xcan74f = 0
 Xcansdf = 0
 Xsubtottt = 0

''' POLICLINICA ESPECIALISTAS
 data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cod_prod =" & 2 & " and convenio not in ('EMERN','EMERJ','EMERNE')"
 data_lin.Refresh
 If data_lin.Recordset.RecordCount > 0 Then
    data_lin.Recordset.MoveFirst
    Do While Not data_lin.Recordset.EOF
       data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("convenio") & "'"
       data_conv.Refresh
       If data_conv.Recordset.RecordCount > 0 Then
          If IsNull(data_conv.Recordset("cnv_grupo")) = False Then
             If data_conv.Recordset("cnv_grupo") = "" Then
                If data_conv.Recordset("cnv_cant_r") = 2 And data_conv.Recordset("cnv_precio") > 0 Then
                   If data_lin.Recordset("cod_cli") > 0 Then
                      data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("cod_cli")
                      data_cli.Refresh
                      If data_cli.Recordset.RecordCount > 0 Then
                           If IsNull(data_cli.Recordset("cl_sexo")) = False Then
                              If data_cli.Recordset("cl_sexo") = 1 Then
                                 If IsNull(data_cli.Recordset("cl_fnac")) = False Then
                                    Xlaedad = CDate(mh.Text) - data_cli.Recordset("cl_fnac")
                                    Xlaedad = Xlaedad / 365
                                    If Xlaedad < 1 Then
                                       Xcan1m = Xcan1m + 1
                                       Xsigrabar = 9
                                    Else
                                       If Xlaedad >= 1 And Xlaedad <= 4 Then
                                          Xcan14m = Xcan14m + 1
                                          Xsigrabar = 9
                                       Else
                                          If Xlaedad >= 5 And Xlaedad <= 14 Then
                                             Xcan514m = Xcan514m + 1
                                             Xsigrabar = 9
                                          Else
                                             If Xlaedad >= 15 And Xlaedad <= 19 Then
                                                Xcan1519m = Xcan1519m + 1
                                                Xsigrabar = 9
                                             Else
                                                If Xlaedad >= 20 And Xlaedad <= 44 Then
                                                   Xcan2044m = Xcan2044m + 1
                                                   Xsigrabar = 9
                                                Else
                                                   If Xlaedad >= 45 And Xlaedad <= 64 Then
                                                      Xcan4564m = Xcan4564m + 1
                                                      Xsigrabar = 9
                                                   Else
                                                      If Xlaedad >= 65 And Xlaedad <= 74 Then
                                                         Xcan6574m = Xcan6574m + 1
                                                         Xsigrabar = 9
                                                      Else
                                                         If Xlaedad >= 75 And Xlaedad <= 120 Then
                                                            Xcan74m = Xcan74m + 1
                                                            Xsigrabar = 9
                                                         Else
                                                            Xcansdm = Xcansdm + 1
                                                            Xsigrabar = 9
                                                         End If
                                                      End If
                                                  End If
                                                End If
                                             End If
                                          End If
                                       End If
                                    End If
                                 Else
                                    Xcansdm = Xcansdm + 1
                                    Xsigrabar = 9
                                 End If
                              Else
                                 If IsNull(data_cli.Recordset("cl_fnac")) = False Then
                                    Xlaedad = CDate(mh.Text) - data_cli.Recordset("cl_fnac")
                                    Xlaedad = Xlaedad / 365
                                    If Xlaedad < 1 Then
                                       Xcan1f = Xcan1f + 1
                                       Xsigrabar = 9
                                    Else
                                       If Xlaedad >= 1 And Xlaedad <= 4 Then
                                          Xcan14f = Xcan14f + 1
                                          Xsigrabar = 9
                                       Else
                                          If Xlaedad >= 5 And Xlaedad <= 14 Then
                                             Xcan514f = Xcan514f + 1
                                             Xsigrabar = 9
                                          Else
                                             If Xlaedad >= 15 And Xlaedad <= 19 Then
                                                Xcan1519f = Xcan1519f + 1
                                                Xsigrabar = 9
                                             Else
                                                If Xlaedad >= 20 And Xlaedad <= 44 Then
                                                   Xcan2044f = Xcan2044f + 1
                                                   Xsigrabar = 9
                                                Else
                                                   If Xlaedad >= 45 And Xlaedad <= 64 Then
                                                      Xcan4564f = Xcan4564f + 1
                                                      Xsigrabar = 9
                                                   Else
                                                      If Xlaedad >= 65 And Xlaedad <= 74 Then
                                                         Xcan6574f = Xcan6574f + 1
                                                         Xsigrabar = 9
                                                      Else
                                                         If Xlaedad >= 75 And Xlaedad <= 120 Then
                                                            Xcan74f = Xcan74f + 1
                                                            Xsigrabar = 9
                                                         Else
                                                            Xcansdf = Xcansdf + 1
                                                            Xsigrabar = 9
                                                         End If
                                                      End If
                                                   End If
                                                End If
                                             End If
                                          End If
                                       End If
                                    End If
                                 Else
                                    Xcansdf = Xcansdf + 1
                                    Xsigrabar = 9
                                 End If
                              End If
                           Else
                              Xcansd = Xcansd + 1
                              Xsigrabar = 9
                           End If
                      Else
                          Xcansd = Xcansd + 1
                          Xsigrabar = 9
                      End If
                   End If
                End If
             Else
                If data_conv.Recordset("cnv_grupo") = "UNIVERSAL" Or data_conv.Recordset("cnv_grupo") = "CCOU" Or _
                   data_conv.Recordset("cnv_grupo") = "SMI" Or data_conv.Recordset("cnv_grupo") = "CASA DE GALICIA" Or _
                   data_conv.Recordset("cnv_grupo") = "H.EVANGELICO" Then
                   If data_conv.Recordset("cnv_codigo") = "CCNOS" Or data_conv.Recordset("cnv_codigo") = "CASANO" Or _
                      data_conv.Recordset("cnv_codigo") = "UNIVS" Or data_conv.Recordset("cnv_codigo") = "IMPNO" Or _
                      data_conv.Recordset("cnv_codigo") = "HEVAN" Or data_conv.Recordset("cnv_codigo") = "HEVANO" Or _
                      data_conv.Recordset("cnv_codigo") = "CCNRE" Or data_conv.Recordset("cnv_codigo") = "SMIN" Or _
                      data_conv.Recordset("cnv_codigo") = "SMINR" Or data_conv.Recordset("cnv_codigo") = "GANOS" Then
                   Else
                        If data_lin.Recordset("cod_cli") > 0 Then
                           data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("cod_cli")
                           data_cli.Refresh
                           If data_cli.Recordset.RecordCount > 0 Then
                                If IsNull(data_cli.Recordset("cl_sexo")) = False Then
                                   If data_cli.Recordset("cl_sexo") = 1 Then
                                      If IsNull(data_cli.Recordset("cl_fnac")) = False Then
                                         Xlaedad = CDate(mh.Text) - data_cli.Recordset("cl_fnac")
                                         Xlaedad = Xlaedad / 365
                                         If Xlaedad < 1 Then
                                            Xcan1m = Xcan1m + 1
                                            Xsigrabar = 9
                                         Else
                                            If Xlaedad >= 1 And Xlaedad <= 4 Then
                                               Xcan14m = Xcan14m + 1
                                               Xsigrabar = 9
                                            Else
                                               If Xlaedad >= 5 And Xlaedad <= 14 Then
                                                  Xcan514m = Xcan514m + 1
                                                  Xsigrabar = 9
                                               Else
                                                  If Xlaedad >= 15 And Xlaedad <= 19 Then
                                                     Xcan1519m = Xcan1519m + 1
                                                     Xsigrabar = 9
                                                  Else
                                                     If Xlaedad >= 20 And Xlaedad <= 44 Then
                                                        Xcan2044m = Xcan2044m + 1
                                                        Xsigrabar = 9
                                                     Else
                                                        If Xlaedad >= 45 And Xlaedad <= 64 Then
                                                           Xcan4564m = Xcan4564m + 1
                                                           Xsigrabar = 9
                                                        Else
                                                           If Xlaedad >= 65 And Xlaedad <= 74 Then
                                                              Xcan6574m = Xcan6574m + 1
                                                              Xsigrabar = 9
                                                           Else
                                                              If Xlaedad >= 75 And Xlaedad <= 120 Then
                                                                 Xcan74m = Xcan74m + 1
                                                                 Xsigrabar = 9
                                                              Else
                                                                 Xcansdm = Xcansdm + 1
                                                                 Xsigrabar = 9
                                                              End If
                                                           End If
                                                       End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      Else
                                         Xcansdm = Xcansdm + 1
                                         Xsigrabar = 9
                                      End If
                                   Else
                                      If IsNull(data_cli.Recordset("cl_fnac")) = False Then
                                         Xlaedad = CDate(mh.Text) - data_cli.Recordset("cl_fnac")
                                         Xlaedad = Xlaedad / 365
                                         If Xlaedad < 1 Then
                                            Xcan1f = Xcan1f + 1
                                            Xsigrabar = 9
                                         Else
                                            If Xlaedad >= 1 And Xlaedad <= 4 Then
                                               Xcan14f = Xcan14f + 1
                                               Xsigrabar = 9
                                            Else
                                               If Xlaedad >= 5 And Xlaedad <= 14 Then
                                                  Xcan514f = Xcan514f + 1
                                                  Xsigrabar = 9
                                               Else
                                                  If Xlaedad >= 15 And Xlaedad <= 19 Then
                                                     Xcan1519f = Xcan1519f + 1
                                                     Xsigrabar = 9
                                                  Else
                                                     If Xlaedad >= 20 And Xlaedad <= 44 Then
                                                        Xcan2044f = Xcan2044f + 1
                                                        Xsigrabar = 9
                                                     Else
                                                        If Xlaedad >= 45 And Xlaedad <= 64 Then
                                                           Xcan4564f = Xcan4564f + 1
                                                           Xsigrabar = 9
                                                        Else
                                                           If Xlaedad >= 65 And Xlaedad <= 74 Then
                                                              Xcan6574f = Xcan6574f + 1
                                                              Xsigrabar = 9
                                                           Else
                                                              If Xlaedad >= 75 And Xlaedad <= 120 Then
                                                                 Xcan74f = Xcan74f + 1
                                                                 Xsigrabar = 9
                                                              Else
                                                                 Xcansdf = Xcansdf + 1
                                                                 Xsigrabar = 9
                                                              End If
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      Else
                                         Xcansdf = Xcansdf + 1
                                         Xsigrabar = 9
                                      End If
                                   End If
                                Else
                                   Xcansd = Xcansd + 1
                                   Xsigrabar = 9
                                End If
                           Else
                               Xcansd = Xcansd + 1
                               Xsigrabar = 9
                           End If
                        End If
                   End If
                End If
             End If
          Else
             If data_conv.Recordset("cnv_cant_r") = 2 And data_conv.Recordset("cnv_precio") > 0 Then
                If data_lin.Recordset("cod_cli") > 0 Then
                   data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("cod_cli")
                   data_cli.Refresh
                   If data_cli.Recordset.RecordCount > 0 Then
                        If IsNull(data_cli.Recordset("cl_sexo")) = False Then
                           If data_cli.Recordset("cl_sexo") = 1 Then
                              If IsNull(data_cli.Recordset("cl_fnac")) = False Then
                                 Xlaedad = CDate(mh.Text) - data_cli.Recordset("cl_fnac")
                                 Xlaedad = Xlaedad / 365
                                 If Xlaedad < 1 Then
                                    Xcan1m = Xcan1m + 1
                                    Xsigrabar = 9
                                 Else
                                    If Xlaedad >= 1 And Xlaedad <= 4 Then
                                       Xcan14m = Xcan14m + 1
                                       Xsigrabar = 9
                                    Else
                                       If Xlaedad >= 5 And Xlaedad <= 14 Then
                                          Xcan514m = Xcan514m + 1
                                          Xsigrabar = 9
                                       Else
                                          If Xlaedad >= 15 And Xlaedad <= 19 Then
                                             Xcan1519m = Xcan1519m + 1
                                             Xsigrabar = 9
                                          Else
                                             If Xlaedad >= 20 And Xlaedad <= 44 Then
                                                Xcan2044m = Xcan2044m + 1
                                                Xsigrabar = 9
                                             Else
                                                If Xlaedad >= 45 And Xlaedad <= 64 Then
                                                   Xcan4564m = Xcan4564m + 1
                                                   Xsigrabar = 9
                                                Else
                                                   If Xlaedad >= 65 And Xlaedad <= 74 Then
                                                      Xcan6574m = Xcan6574m + 1
                                                      Xsigrabar = 9
                                                   Else
                                                      If Xlaedad >= 75 And Xlaedad <= 120 Then
                                                         Xcan74m = Xcan74m + 1
                                                         Xsigrabar = 9
                                                      Else
                                                         Xcansdm = Xcansdm + 1
                                                         Xsigrabar = 9
                                                      End If
                                                   End If
                                               End If
                                             End If
                                          End If
                                       End If
                                    End If
                                 End If
                              Else
                                 Xcansdm = Xcansdm + 1
                                 Xsigrabar = 9
                              End If
                           Else
                              If IsNull(data_cli.Recordset("cl_fnac")) = False Then
                                 Xlaedad = CDate(mh.Text) - data_cli.Recordset("cl_fnac")
                                 Xlaedad = Xlaedad / 365
                                 If Xlaedad < 1 Then
                                    Xcan1f = Xcan1f + 1
                                    Xsigrabar = 9
                                 Else
                                    If Xlaedad >= 1 And Xlaedad <= 4 Then
                                       Xcan14f = Xcan14f + 1
                                       Xsigrabar = 9
                                    Else
                                       If Xlaedad >= 5 And Xlaedad <= 14 Then
                                          Xcan514f = Xcan514f + 1
                                          Xsigrabar = 9
                                       Else
                                          If Xlaedad >= 15 And Xlaedad <= 19 Then
                                             Xcan1519f = Xcan1519f + 1
                                             Xsigrabar = 9
                                          Else
                                             If Xlaedad >= 20 And Xlaedad <= 44 Then
                                                Xcan2044f = Xcan2044f + 1
                                                Xsigrabar = 9
                                             Else
                                                If Xlaedad >= 45 And Xlaedad <= 64 Then
                                                   Xcan4564f = Xcan4564f + 1
                                                   Xsigrabar = 9
                                                Else
                                                   If Xlaedad >= 65 And Xlaedad <= 74 Then
                                                      Xcan6574f = Xcan6574f + 1
                                                      Xsigrabar = 9
                                                   Else
                                                      If Xlaedad >= 75 And Xlaedad <= 120 Then
                                                         Xcan74f = Xcan74f + 1
                                                         Xsigrabar = 9
                                                      Else
                                                         Xcansdf = Xcansdf + 1
                                                         Xsigrabar = 9
                                                      End If
                                                   End If
                                                End If
                                             End If
                                          End If
                                       End If
                                    End If
                                 End If
                              Else
                                 Xcansdf = Xcansdf + 1
                                 Xsigrabar = 9
                              End If
                           End If
                        Else
                           Xcansd = Xcansd + 1
                           Xsigrabar = 9
                        End If
                   Else
                       Xcansd = Xcansd + 1
                       Xsigrabar = 9
                   End If
                End If
             End If
          End If
       End If
       Xsigrabar = 0
       data_lin.Recordset.MoveNext
    Loop
 End If
  data_msp.Recordset.Edit
  data_msp.Recordset("DESC6") = "ESPECIALISTAS"
  data_msp.Recordset("m1v") = Xcan1m
  data_msp.Recordset("m1a4v") = Xcan14m
  data_msp.Recordset("m5a14v") = Xcan514m
  data_msp.Recordset("m15a19v") = Xcan1519m
  data_msp.Recordset("m20a44v") = Xcan2044m
  data_msp.Recordset("m45a64v") = Xcan4564m
  data_msp.Recordset("m65a74v") = Xcan6574m
  data_msp.Recordset("m74v") = Xcan74m
  data_msp.Recordset("msdv") = Xcansdm
  Xsubtott = Xcan1m + Xcan14m + Xcan514m + Xcan1519m + Xcan2044m + Xcan4564m + Xcan6574m + Xcan74m + Xcansdm
  data_msp.Recordset("mtv") = Xsubtott
  data_msp.Recordset("f1v") = Xcan1f
  data_msp.Recordset("f1a4v") = Xcan14f
  data_msp.Recordset("f5a14v") = Xcan514f
  data_msp.Recordset("f15a19v") = Xcan1519f
  data_msp.Recordset("f20a44v") = Xcan2044f
  data_msp.Recordset("f45a64v") = Xcan4564f
  data_msp.Recordset("f65a74v") = Xcan6574f
  data_msp.Recordset("f74v") = Xcan74f
  data_msp.Recordset("fsdv") = Xcansdf
  Xsubtottt = Xcan1f + Xcan14f + Xcan514f + Xcan1519f + Xcan2044f + Xcan4564f + Xcan6574f + Xcan74f + Xcansdf
  data_msp.Recordset("ftv") = Xsubtottt
  data_msp.Recordset("totv") = Xsubtott + Xsubtottt
  data_msp.Recordset.Update
  Xcan1m = 0
  Xcan14m = 0
  Xcan514m = 0
  Xcan1519m = 0
  Xcan2044m = 0
  Xcan4564m = 0
  Xcan6574m = 0
  Xcan74m = 0
  Xcansdm = 0
  Xsubtott = 0
  Xsubtott = 0
  Xcan1f = 0
  Xcan14f = 0
  Xcan514f = 0
  Xcan1519f = 0
  Xcan2044f = 0
  Xcan4564f = 0
  Xcan6574f = 0
  Xcan74f = 0
  Xcansdf = 0
  Xsubtottt = 0

If Check1.Value = 1 Then
'   data_lin.DatabaseName = App.Path & "\llamado.mdb"
Else
    data_lin.ConnectionString = "dsn=" & Xconexrmt
'   data_lin.DatabaseName = App.Path & "\sapp.mdb"
End If
'MsgBox "Termina el commando3"

Command4_Click

End Sub

Private Sub Command4_Click()
Dim Xcantras As Double
Xcan1m = 0
Xcan14m = 0
Xcan514m = 0
Xcan1519m = 0
Xcan2044m = 0
Xcan4564m = 0
Xcan6574m = 0
Xcan74m = 0
Xcansdm = 0
Xcan1f = 0
Xcan14f = 0
Xcan514f = 0
Xcan1519f = 0
Xcan2044f = 0
Xcan4564f = 0
Xcan6574f = 0
Xcan74f = 0
Xcansdf = 0
'MsgBox "Comienza el commando4"

Xcantras = 0
 data_lin.RecordSource = "Select * from llamado where codzon in (1,2,3) and fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,4,5) and cancela is null"
 data_lin.Refresh
 If data_lin.Recordset.RecordCount > 0 Then
    data_lin.Recordset.MoveFirst
    Do While Not data_lin.Recordset.EOF
       data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("categ") & "'"
       data_conv.Refresh
       If data_conv.Recordset.RecordCount > 0 Then
          If IsNull(data_conv.Recordset("cnv_grupo")) = False Then
             If data_conv.Recordset("cnv_grupo") = "" Then
                If data_conv.Recordset("cnv_cant_r") = 2 And data_conv.Recordset("cnv_precio") > 0 Then
                   Xcantras = Xcantras + 1
                End If
             Else
                If data_conv.Recordset("cnv_grupo") = "UNIVERSAL" Or data_conv.Recordset("cnv_grupo") = "CCOU" Or _
                   data_conv.Recordset("cnv_grupo") = "SMI" Or data_conv.Recordset("cnv_grupo") = "CASA DE GALICIA" Or _
                   data_conv.Recordset("cnv_grupo") = "H.EVANGELICO" Then
                   If data_conv.Recordset("cnv_codigo") = "CCNOS" Or data_conv.Recordset("cnv_codigo") = "CASANO" Or _
                      data_conv.Recordset("cnv_codigo") = "UNIVS" Or data_conv.Recordset("cnv_codigo") = "IMPNO" Or _
                      data_conv.Recordset("cnv_codigo") = "HEVAN" Or data_conv.Recordset("cnv_codigo") = "HEVANO" Or _
                      data_conv.Recordset("cnv_codigo") = "CCNRE" Or data_conv.Recordset("cnv_codigo") = "SMIN" Or _
                      data_conv.Recordset("cnv_codigo") = "SMINR" Or data_conv.Recordset("cnv_codigo") = "GANOS" Then
                   Else
                      Xcantras = Xcantras + 1
                   End If
                End If
             End If
          Else
             If data_conv.Recordset("cnv_cant_r") = 2 And data_conv.Recordset("cnv_precio") > 0 Then
                Xcantras = Xcantras + 1
             End If
          End If
       End If
       data_lin.Recordset.MoveNext
    Loop
 End If
 data_msp.RecordSource = "res"
 data_msp.Refresh
 
 data_msp.Recordset.AddNew
 data_msp.Recordset("nro") = 4
' data_msp.Recordset("desc") = "TRASLADOS"
' data_msp.Recordset("desc2") = ""
 data_msp.Recordset("cant") = Xcantras
 data_msp.Recordset.Update
 data_msp.Refresh
 Xcantras = 0

''''' 911
 data_lin.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and categ in ('911','911B') and cancela is null"
 data_lin.Refresh
 If data_lin.Recordset.RecordCount > 0 Then
    data_lin.Recordset.MoveFirst
    Do While Not data_lin.Recordset.EOF
       If data_lin.Recordset("categ") = "911" Or _
          data_lin.Recordset("categ") = "911B" Then
          Xcantras = Xcantras + 1
       End If
       data_lin.Recordset.MoveNext
    Loop
 End If
 data_msp.Recordset.Edit
 data_msp.Recordset("cant2") = Xcantras
 data_msp.Recordset.Update
 data_msp.Refresh
 Xcantras = 0
 
'a.p.
data_lin.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codzon in (1,2,3) and base =" & 0 & " and cancela is null and codmot ='" & "C" & "'"
data_lin.Refresh
If data_lin.Recordset.RecordCount > 0 Then
'   data_lin.Recordset.MoveLast
   data_lin.Recordset.MoveFirst
'   Xcantras = data_lin.Recordset.RecordCount
   Do While Not data_lin.Recordset.EOF
      If data_lin.Recordset("movilpas") = 202 Or data_lin.Recordset("movilpas") = 203 Then
         Xcantras = Xcantras + 1
      End If
      data_lin.Recordset.MoveNext
   Loop
End If
''' Areas protegidas
 data_msp.Recordset.Edit
 data_msp.Recordset("cant3") = Xcantras
 data_msp.Recordset.Update
 data_msp.Refresh
 Xcantras = 0

Dim XXlamat, Xxveces As Double
XXlamat = 0

data_lin.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codzon in (1,2,3) and base =" & 0 & " and cancela is null order by matric"
data_lin.Refresh
If data_lin.Recordset.RecordCount > 0 Then
   data_lin.Recordset.MoveFirst
   Do While Not data_lin.Recordset.EOF
      data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("categ") & "'"
      data_conv.Refresh
      If data_conv.Recordset.RecordCount > 0 Then
         If IsNull(data_conv.Recordset("cnv_grupo")) = False Then
            If data_conv.Recordset("cnv_grupo") = "" Then
                If data_conv.Recordset("cnv_cant_r") = 2 And data_conv.Recordset("cnv_precio") > 0 Then
                   If data_lin.Recordset("matric") > 0 Then
                   Else
                      If XXlamat = data_lin.Recordset("matric") Then
                         Xcantras = Xcantras + 1
                      End If
                   End If
                   XXlamat = data_lin.Recordset("matric")
                End If
            Else
                If data_conv.Recordset("cnv_grupo") = "UNIVERSAL" Or data_conv.Recordset("cnv_grupo") = "CCOU" Or _
                   data_conv.Recordset("cnv_grupo") = "SMI" Or data_conv.Recordset("cnv_grupo") = "CASA DE GALICIA" Or _
                   data_conv.Recordset("cnv_grupo") = "H.EVANGELICO" Then
                   If data_conv.Recordset("cnv_codigo") = "CCNOS" Or data_conv.Recordset("cnv_codigo") = "CASANO" Or _
                      data_conv.Recordset("cnv_codigo") = "UNIVS" Or data_conv.Recordset("cnv_codigo") = "IMPNO" Or _
                      data_conv.Recordset("cnv_codigo") = "HEVAN" Or data_conv.Recordset("cnv_codigo") = "HEVANO" Or _
                      data_conv.Recordset("cnv_codigo") = "CCNRE" Or data_conv.Recordset("cnv_codigo") = "SMIN" Or _
                      data_conv.Recordset("cnv_codigo") = "SMINR" Or data_conv.Recordset("cnv_codigo") = "GANOS" Then
                   Else
                      If data_lin.Recordset("matric") > 0 Then
                      Else
                         If XXlamat = data_lin.Recordset("matric") Then
                            Xcantras = Xcantras + 1
                         End If
                      End If
                      XXlamat = data_lin.Recordset("matric")
                   End If
               End If
            End If
         Else
            If data_conv.Recordset("cnv_cant_r") = 2 And data_conv.Recordset("cnv_precio") > 0 Then
               If data_lin.Recordset("matric") > 0 Then
               Else
                  If XXlamat = data_lin.Recordset("matric") Then
                     Xcantras = Xcantras + 1
                  End If
               End If
               XXlamat = data_lin.Recordset("matric")
            End If
         End If
      End If
      data_lin.Recordset.MoveNext
   Loop
End If
''' Reconsultas
 data_msp.Recordset.Edit
 data_msp.Recordset("cant4") = Xcantras
 data_msp.Recordset.Update
 data_msp.Refresh
 Xcantras = 0

 

End Sub

Private Sub Command5_Click()
Dim m1, m1a4, m5a14, m15a19, m20a44, m45a64, m65a74, m74, msd As Long
Dim f1, f1a4, f5a14, f15a19, f20a44, f45a64, f65a74, f74, fsd As Long
Dim m1ia, m1a4ia, m5a14ia, m15a19ia, m20a44ia, m45a64ia, m65a74ia, m74ia, msdia As Long
Dim f1ia, f1a4ia, f5a14ia, f15a19ia, f20a44ia, f45a64ia, f65a74ia, f74ia, fsdia As Long
Dim Xanos As Double
Dim Xsindatos, Xsindatosia As Long


data_inf.DatabaseName = App.Path & "\informes.mdb"
data_inf.RecordSource = "infcli"
data_inf.Refresh
If Check2.Value = 1 Then
   Command6_Click
Else
    Xsindatos = 0
    If data_emi.Recordset.RecordCount > 0 Then
       data_emi.Recordset.MoveLast
       pb.Max = data_emi.Recordset.RecordCount
       data_emi.Recordset.MoveFirst
       DoEvents
       Do While Not data_emi.Recordset.EOF
          data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_emi.Recordset("cliente")
          data_cli.Refresh
          If data_cli.Recordset.RecordCount > 0 Then
             If data_emi.Recordset("nro_cobr") = 26 Or _
                data_emi.Recordset("nro_cobr") = 27 Or _
                data_emi.Recordset("nro_cobr") = 28 Then
             Else
                data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_emi.Recordset("cod_cnv") & "'"
                data_conv.Refresh
                If data_conv.Recordset.RecordCount > 0 Then
    '               If data_conv.Recordset("cnv_cant_r") = 2 Or data_conv.Recordset("cnv_cant_r") = 1 Then
                        If IsNull(data_cli.Recordset("cl_sexo")) = True Then
                           Xsindatos = Xsindatos + 1
                        Else
                           If data_cli.Recordset("cl_sexo") = 2 Then ' FEM
                              If IsNull(data_cli.Recordset("cl_fnac")) = True Then
                                 fsd = fsd + 1
                              Else
                                 Xanos = Date - data_cli.Recordset("cl_fnac")
                                 Xanos = Xanos / 365
                                 If Xanos < 1 Then
                                    f1 = f1 + 1
                                 Else
                                    If Xanos >= 1 And Xanos <= 4 Then
                                       f1a4 = f1a4 + 1
                                    Else
                                       If Xanos >= 5 And Xanos <= 14 Then
                                          f5a14 = f5a14 + 1
                                       Else
                                          If Xanos >= 15 And Xanos <= 19 Then
                                             f15a19 = f15a19 + 1
                                          Else
                                             If Xanos >= 20 And Xanos <= 44 Then
                                                f20a44 = f20a44 + 1
                                             Else
                                                If Xanos >= 45 And Xanos <= 64 Then
                                                   f45a64 = f45a64 + 1
                                                Else
                                                   If Xanos >= 65 And Xanos <= 74 Then
                                                      f65a74 = f65a74 + 1
                                                   Else
                                                      If Xanos >= 75 And Xanos <= 110 Then
                                                         f74 = f74 + 1
                                                      Else
                                                         fsd = fsd + 1
                                                      End If
                                                   End If
                                                End If
                                             End If
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                           Else
                              If data_cli.Recordset("cl_sexo") = 1 Then ' MASC
                                 If IsNull(data_cli.Recordset("cl_fnac")) = True Then
                                    msd = msd + 1
                                 Else
                                    Xanos = Date - data_cli.Recordset("cl_fnac")
                                    Xanos = Xanos / 365
                                    If Xanos < 1 Then
                                       m1 = m1 + 1
                                    Else
                                       If Xanos >= 1 And Xanos <= 4 Then
                                          m1a4 = m1a4 + 1
                                       Else
                                          If Xanos >= 5 And Xanos <= 14 Then
                                             m5a14 = m5a14 + 1
                                          Else
                                             If Xanos >= 15 And Xanos <= 19 Then
                                                m15a19 = m15a19 + 1
                                             Else
                                                If Xanos >= 20 And Xanos <= 44 Then
                                                   m20a44 = m20a44 + 1
                                                Else
                                                   If Xanos >= 45 And Xanos <= 64 Then
                                                      m45a64 = m45a64 + 1
                                                   Else
                                                      If Xanos >= 65 And Xanos <= 74 Then
                                                         m65a74 = m65a74 + 1
                                                      Else
                                                         If Xanos >= 75 And Xanos <= 110 Then
                                                            m74 = m74 + 1
                                                         Else
                                                            msd = msd + 1
                                                         End If
                                                      End If
                                                   End If
                                                End If
                                             End If
                                          End If
                                       End If
                                    End If
                                 End If
                              Else
                                 Xsindatos = Xsindatos + 1
                              End If
                           End If
                        End If
    '               Else
    '                    Xsindatos = Xsindatos + 1
    '               End If
                End If
             End If
          Else
             Xsindatos = Xsindatos + 1
          End If
          data_emi.Recordset.MoveNext
          pb.Value = pb.Value + 1
       Loop
    End If
    If cbocat.ListIndex = 2 Then
       data_cli.RecordSource = "Select * from clientes where estado in (1,0) and cl_grupo >=" & 100 & " and cl_grupo <=" & 115
    End If
    If cbocat.ListIndex = 3 Then
       data_cli.RecordSource = "Select * from clientes where estado in (1,0) and cl_grupo >=" & 300 & " and cl_grupo <=" & 320
    End If
    If cbocat.ListIndex = 4 Then
''''''''''       data_cli.RecordSource = "Select * from clientes where estado in (1,0)"
'       data_cli.RecordSource = "Select * from clientes where estado in (1,0) and cl_grupo in (500)"
'       data_cli.RecordSource = "Select * from clientes where estado in (1,0) and cl_grupo in (670,672,673,674)"
'       data_cli.RecordSource = "Select * from clientes where estado in (1,0) and cl_grupo in (101,102,103,104,700,701,702,703,704,705,706,707,708,709,710,711,712,713,714,715,716,717,722)"
''       data_cli.RecordSource = "Select * from clientes where estado in (1,0) and cl_grupo in (201,202,203,204,205,206,207,208)"
''       data_cli.RecordSource = "Select * from clientes where estado in (1,0) and cl_grupo in (401,402,403,404,405,406)"
'       data_cli.RecordSource = "Select * from clientes where estado in (1,0) and cl_grupo in (630)"

''''       data_cli.RecordSource = "Select * from clientes where estado in (1,0) and cl_grupo in (301,302,303,304,305,306,307,308,309,310)"
'       data_cli.RecordSource = "Select * from clientes where estado in (1,0) and cl_grupo in (700,701,702,703,704,705,706,707,708,709,710,711,712,713,714,715,716,717,722)"
'       data_cli.RecordSource = "Select * from clientes where estado in (1,0) and cl_grupo in (672,673,674)"
'       data_cli.RecordSource = "Select * from clientes where estado in (1,0) and cl_grupo in (811,640)"
'       data_cli.RecordSource = "Select * from clientes where estado in (1,0) and cl_grupo in (101,102,103,104)"
'       data_cli.RecordSource = "Select * from clientes where estado in (1,0) and cl_grupo in (670,672,673,674)"
'       data_cli.RecordSource = "Select * from clientes where estado in (1,0) and cl_grupo in (650,800,801)"
       data_cli.RecordSource = "Select * from clientes where estado in (1,0) and cl_grupo in (810)"

'       data_cli.RecordSource = "Select * from clientes where estado in (1,0) and cl_grupo in (650,800,801,802,803)"
       
'       data_cli.RecordSource = "Select * from clientes where estado in (1,0) and cl_grupo >=" & 600 & " and cl_grupo <=" & 606
    End If
'    data_cli.RecordSource = "Select * from clientes where estado in (1,0) and cl_grupo >=" & 600 & " and cl_grupo <=" & 640
'    data_cli.RecordSource = "Select * from clientes where estado in (1,0) and cl_grupo =" & 999
    data_cli.Refresh
    
    If data_cli.Recordset.RecordCount > 0 Then
    '   data_cli.Recordset.MoveLast
    '   pb.Max = pb.Max + data_cli.Recordset.RecordCount
       data_cli.Recordset.MoveFirst
       DoEvents
       Do While Not data_cli.Recordset.EOF
          If IsNull(data_cli.Recordset("cl_codconv")) = False Then
             If data_cli.Recordset("cl_codconv") = "CASH" Or _
                data_cli.Recordset("cl_codconv") = "CPS" Or _
                data_cli.Recordset("cl_codconv") = "SEMM1" Or _
                data_cli.Recordset("cl_codconv") = "1727" Or _
                data_cli.Recordset("cl_codconv") = "CCNOS" Or _
                data_cli.Recordset("cl_codconv") = "UNIVS" Or _
                data_cli.Recordset("cl_codconv") = "IMPNO" Or _
                data_cli.Recordset("cl_codconv") = "HEVAN" Or _
                data_cli.Recordset("cl_codconv") = "HEVANO" Or _
                data_cli.Recordset("cl_codconv") = "CCNRE" Or _
                data_cli.Recordset("cl_codconv") = "SMINR" Or _
                data_cli.Recordset("cl_codconv") = "SMIN" Or _
                data_cli.Recordset("cl_codconv") = "GANOS" Or _
                data_cli.Recordset("cl_codconv") = "1727B" Or _
                data_cli.Recordset("cl_codconv") = "1727C1" Or _
                data_cli.Recordset("cl_codconv") = "SEMM" Or _
                data_cli.Recordset("cl_codconv") = "911" Or _
                data_cli.Recordset("cl_codconv") = "911B" Or _
                data_cli.Recordset("cl_codconv") = "RETMI" Or _
                data_cli.Recordset("cl_codconv") = "MSP" Then
             Else
                If data_cli.Recordset("cl_grupo") = 999 Then
                   Xsindatosia = Xsindatosia + 1
'                    data_inf.Recordset.AddNew
'                    data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
'                    data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
'                    data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
'                    data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
'                    data_inf.Recordset("estado") = data_cli.Recordset("estado")
'                    data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
'                    data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
'                    data_inf.Recordset.Update
                
                Else
                    data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                    data_conv.Refresh
                    If data_conv.Recordset.RecordCount > 0 Then
                       If IsNull(data_conv.Recordset("cnv_grupo")) = False Then
                          If data_conv.Recordset("cnv_grupo") <> "" Then
                             If data_conv.Recordset("cnv_cant_r") = 2 And data_conv.Recordset("cnv_precio") > 0 Then
                             Else
                                If IsNull(data_cli.Recordset("cl_sexo")) = True Then
                                   Xsindatosia = Xsindatosia + 1
                                   data_inf.Recordset.AddNew
                                   data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                                   data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                                   data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                                   data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                                   data_inf.Recordset("estado") = data_cli.Recordset("estado")
                                   data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                                   data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                                   data_inf.Recordset.Update
                                Else
                                   If data_cli.Recordset("cl_sexo") = 2 Then ' FEM
                                      If IsNull(data_cli.Recordset("cl_fnac")) = True Then
                                         fsdia = fsdia + 1
                                      Else
                                         Xanos = Date - data_cli.Recordset("cl_fnac")
                                         Xanos = Xanos / 365
                                         If Xanos < 1 Then
                                            f1ia = f1ia + 1
                                         Else
                                            If Xanos >= 1 And Xanos <= 4 Then
                                               f1a4ia = f1a4ia + 1
                                            Else
                                               If Xanos >= 5 And Xanos <= 14 Then
                                                  f5a14ia = f5a14ia + 1
                                               Else
                                                  If Xanos >= 15 And Xanos <= 19 Then
                                                     f15a19ia = f15a19ia + 1
                                                  Else
                                                     If Xanos >= 20 And Xanos <= 44 Then
                                                        f20a44ia = f20a44ia + 1
                                                     Else
                                                        If Xanos >= 45 And Xanos <= 64 Then
                                                           f45a64ia = f45a64ia + 1
                                                        Else
                                                           If Xanos >= 65 And Xanos <= 74 Then
                                                              f65a74ia = f65a74ia + 1
                                                           Else
                                                              If Xanos >= 75 And Xanos <= 110 Then
                                                                 f74ia = f74ia + 1
                                                              Else
                                                                 fsdia = fsdia + 1
                                                              End If
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                         data_inf.Recordset.AddNew
                                         data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                                         data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                                         data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                                         data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                                         data_inf.Recordset("estado") = data_cli.Recordset("estado")
                                         data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                                         data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                                         data_inf.Recordset.Update
                                         
                                      End If
                                   Else
                                      If data_cli.Recordset("cl_sexo") = 1 Then ' MASC
                                         If IsNull(data_cli.Recordset("cl_fnac")) = True Then
                                            msdia = msdia + 1
                                         Else
                                            Xanos = Date - data_cli.Recordset("cl_fnac")
                                            Xanos = Xanos / 365
                                            If Xanos < 1 Then
                                               m1ia = m1ia + 1
                                            Else
                                               If Xanos >= 1 And Xanos <= 4 Then
                                                  m1a4ia = m1a4ia + 1
                                               Else
                                                  If Xanos >= 5 And Xanos <= 14 Then
                                                     m5a14ia = m5a14ia + 1
                                                  Else
                                                     If Xanos >= 15 And Xanos <= 19 Then
                                                        m15a19ia = m15a19ia + 1
                                                     Else
                                                        If Xanos >= 20 And Xanos <= 44 Then
                                                           m20a44ia = m20a44ia + 1
                                                        Else
                                                           If Xanos >= 45 And Xanos <= 64 Then
                                                              m45a64ia = m45a64ia + 1
                                                           Else
                                                              If Xanos >= 65 And Xanos <= 74 Then
                                                                 m65a74ia = m65a74ia + 1
                                                              Else
                                                                 If Xanos >= 75 And Xanos <= 110 Then
                                                                    m74ia = m74ia + 1
                                                                 Else
                                                                    msdia = msdia + 1
                                                                 End If
                                                              End If
                                                           End If
                                                        End If
                                                     End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      Else
                                         Xsindatosia = Xsindatosia + 1
                                      End If
                                      data_inf.Recordset.AddNew
                                      data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                                      data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                                      data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                                      data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                                      data_inf.Recordset("estado") = data_cli.Recordset("estado")
                                      data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                                      data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                                      data_inf.Recordset.Update
                                   End If
                                End If
                             End If
                          End If
                       End If
                    End If
                End If
             End If
          End If
          data_cli.Recordset.MoveNext
    '      pb.Value = pb.Value + 1
       Loop
    End If
    data_inf.DatabaseName = App.Path & "\informes.mdb"
    data_inf.RecordSource = "inflla"
    data_inf.Refresh
    
    data_msp.RecordSource = "benef"
    data_msp.Refresh
    data_msp.Recordset.AddNew
    data_msp.Recordset("ano") = Year(mh.Text)
    data_msp.Recordset("m1") = m1
    data_msp.Recordset("m1a4") = m1a4
    data_msp.Recordset("m5a14") = m5a14
    data_msp.Recordset("m15a19") = m15a19
    data_msp.Recordset("m20a44") = m20a44
    data_msp.Recordset("m45a64") = m45a64
    data_msp.Recordset("m65a74") = m65a74
    data_msp.Recordset("m74") = m74
    data_msp.Recordset("msd") = msd
    data_msp.Recordset("f1") = f1
    data_msp.Recordset("f1a4") = f1a4
    data_msp.Recordset("f5a14") = f5a14
    data_msp.Recordset("f15a19") = f15a19
    data_msp.Recordset("f20a44") = f20a44
    data_msp.Recordset("f45a64") = f45a64
    data_msp.Recordset("f65a74") = f65a74
    data_msp.Recordset("f74") = f74
    data_msp.Recordset("fsd") = fsd
    
    data_msp.Recordset("m1ia") = m1ia
    data_msp.Recordset("m1a4ia") = m1a4ia
    data_msp.Recordset("m5a14ia") = m5a14ia
    data_msp.Recordset("m15a19ia") = m15a19ia
    data_msp.Recordset("m20a44ia") = m20a44ia
    data_msp.Recordset("m45a64ia") = m45a64ia
    data_msp.Recordset("m65a74ia") = m65a74ia
    data_msp.Recordset("m74ia") = m74ia
    data_msp.Recordset("msdia") = msdia
    data_msp.Recordset("f1ia") = f1ia
    data_msp.Recordset("f1a4ia") = f1a4ia
    data_msp.Recordset("f5a14ia") = f5a14ia
    data_msp.Recordset("f15a19ia") = f15a19ia
    data_msp.Recordset("f20a44ia") = f20a44ia
    data_msp.Recordset("f45a64ia") = f45a64ia
    data_msp.Recordset("f65a74ia") = f65a74ia
    data_msp.Recordset("f74ia") = f74ia
    data_msp.Recordset("fsdia") = fsdia
    
    data_msp.Recordset("sd") = Xsindatos
    data_msp.Recordset("sdia") = Xsindatosia
    
    data_msp.Recordset.Update
    
    MsgBox "Proceso de Beneficiarios, TERMINADO!!"
    cr1.ReportFileName = App.Path & "\infbenef.rpt"
    cr1.ReportTitle = "POLICLINICA DE SALINAS Al: " & mh.Text
    cr1.Action = 1
End If


End Sub

Private Sub Command6_Click()
Dim Xcuandias, Xqueedadt As Long
Dim Xcantsoc As Double
Dim Xcansdm, Xcansdf, Xcansd As Long
Dim Xcan1m, Xcan14m, Xcan514m, Xcan1519m, Xcan2044m, Xcan4564m, Xcan6574m, Xcan74m As Long
Dim Xcan1f, Xcan14f, Xcan514f, Xcan1519f, Xcan2044f, Xcan4564f, Xcan6574f, Xcan74f As Long
Dim Xsubtott, Xsubtottt As Double
XcuentasinC = 0
'138 - Traslados
' - 59
' 11 Eventos en vía

If Check1.Value = 1 Then
'   data_lin.DatabaseName = App.Path & "\llamado.mdb"
Else
   data_lin.ConnectionString = "dsn=" & Xconexrmt
'   data_lin.DatabaseName = App.Path & "\sapp.mdb"
End If

data_msp.DatabaseName = App.Path & "\infmsp.mdb"
data_msp.RecordSource = "plani"
data_msp.Refresh

Xcan1m = 0
Xcan14m = 0
Xcan514m = 0
Xcan1519m = 0
Xcan2044m = 0
Xcan4564m = 0
Xcan6574m = 0
Xcan74m = 0
Xcansdm = 0
Xcan1f = 0
Xcan14f = 0
Xcan514f = 0
Xcan1519f = 0
Xcan2044f = 0
Xcan4564f = 0
Xcan6574f = 0
Xcan74f = 0
Xcansdf = 0
           
data_lin.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codzon in (1,2,3) and base =" & 0 & " and codmot ='" & "R" & "' and cancela is null"
data_lin.Refresh
If data_lin.Recordset.RecordCount > 0 Then
   data_lin.Recordset.MoveFirst
   Do While Not data_lin.Recordset.EOF
      If data_lin.Recordset("categ") = "CASH" Or _
         data_lin.Recordset("categ") = "CPS" Or _
         data_lin.Recordset("categ") = "SEMM1" Or _
         data_lin.Recordset("categ") = "1727" Or _
         data_lin.Recordset("categ") = "CCNOS" Or _
         data_lin.Recordset("categ") = "UNIVS" Or _
         data_lin.Recordset("categ") = "IMPNO" Or _
         data_lin.Recordset("categ") = "HEVAN" Or _
         data_lin.Recordset("categ") = "HEVANO" Or _
         data_lin.Recordset("categ") = "CCNRE" Or _
         data_lin.Recordset("categ") = "SMINR" Or _
         data_lin.Recordset("categ") = "SMIN" Or _
         data_lin.Recordset("categ") = "GANOS" Or _
         data_lin.Recordset("categ") = "911" Or _
         data_lin.Recordset("categ") = "911B" Or _
         data_lin.Recordset("categ") = "1727B" Or _
         data_lin.Recordset("categ") = "UCM" Or _
         data_lin.Recordset("categ") = "RETMI" Then
      Else
         If IsNull(data_lin.Recordset("cancela")) = True Then
'            If IsNull(data_lin.Recordset("hh")) = True Then
               If data_lin.Recordset("matric") > 0 Then
                  If cbocat.ListIndex = 2 Then
                     data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("matric") & " and cl_grupo >=" & 100 & " and cl_grupo <=" & 115
                  End If
                  If cbocat.ListIndex = 3 Then
                     data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("matric") & " and cl_grupo >=" & 300 & " and cl_grupo <=" & 320
                  End If
                  If cbocat.ListIndex = 4 Then
'                     data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("matric") & " and cl_grupo >=" & 600 & " and cl_grupo <=" & 640
                     data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("matric") & " and cl_grupo >=" & 600 & " and cl_grupo <=" & 640
                  
                  End If
                  data_cli.Refresh
                  If data_cli.Recordset.RecordCount > 0 Then
                     If Check3.Value = 1 Then
                        data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("categ") & "' and cnv_grupo <>'" & Null & "'"
                        data_conv.Refresh
                     Else
                        data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("categ") & "'"
                        data_conv.Refresh
                     End If
                     If data_conv.Recordset.RecordCount > 0 Then
                        If data_conv.Recordset("cnv_cant_r") = 1 And data_conv.Recordset("cnv_colrec") = "M" Then
                        Else
                           If IsNull(data_cli.Recordset("cl_sexo")) = False Then
                              If data_cli.Recordset("cl_sexo") = 1 Then
                                 If data_lin.Recordset("unied") < 3 Then
                                    Xcan1m = Xcan1m + 1
                                 Else
                                    If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                       Xcan14m = Xcan14m + 1
                                    Else
                                       If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                          Xcan514m = Xcan514m + 1
                                       Else
                                          If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                             Xcan1519m = Xcan1519m + 1
                                          Else
                                             If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                                Xcan2044m = Xcan2044m + 1
                                             Else
                                                If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                   Xcan4564m = Xcan4564m + 1
                                                Else
                                                   If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                      Xcan6574m = Xcan6574m + 1
                                                   Else
                                                      If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                         Xcan74m = Xcan74m + 1
                                                      Else
                                                         Xcansdm = Xcansdm + 1
                                                      End If
                                                   End If
                                                End If
                                             End If
                                          End If
                                       End If
                                    End If
                                 End If
                              Else
                                 If data_lin.Recordset("unied") < 3 Then
                                    Xcan1f = Xcan1f + 1
                                 Else
                                    If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                       Xcan14f = Xcan14f + 1
                                    Else
                                       If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                          Xcan514f = Xcan514f + 1
                                       Else
                                          If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                             Xcan1519f = Xcan1519f + 1
                                          Else
                                             If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                                Xcan2044f = Xcan2044f + 1
                                             Else
                                                If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                   Xcan4564f = Xcan4564f + 1
                                                Else
                                                   If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                      Xcan6574f = Xcan6574f + 1
                                                   Else
                                                      If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                         Xcan74f = Xcan74f + 1
                                                      Else
                                                         Xcansdf = Xcansdf + 1
                                                      End If
                                                   End If
                                                End If
                                             End If
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                           Else
                              Xcansd = Xcansd + 1
                           End If
                        End If
                     End If
                  Else
'                     Xcansd = Xcansd + 1
                  End If
               Else
'                  Xcansd = Xcansd + 1
               End If
         End If
      End If
      data_lin.Recordset.MoveNext
   Loop
End If
data_msp.Recordset.AddNew
data_msp.Recordset("nro") = 1
data_msp.Recordset("mes") = Month(md.Text)
data_msp.Recordset("ano") = Year(mh.Text)
data_msp.Recordset("desc") = "ACTIVIDAD DOMICILIARIA"
data_msp.Recordset("desc2") = "Segun clasificación en recepción"
data_msp.Recordset("DESC4") = "EMERGENCIA"
data_msp.Recordset("m1") = Xcan1m
data_msp.Recordset("m1a4") = Xcan14m
 data_msp.Recordset("m5a14") = Xcan514m
 data_msp.Recordset("m15a19") = Xcan1519m
 data_msp.Recordset("m20a44") = Xcan2044m
 data_msp.Recordset("m45a64") = Xcan4564m
 data_msp.Recordset("m65a74") = Xcan6574m
 data_msp.Recordset("m74") = Xcan74m
 data_msp.Recordset("msd") = Xcansdm
 Xsubtott = Xcan1m + Xcan14m + Xcan514m + Xcan1519m + Xcan2044m + Xcan4564m + Xcan6574m + Xcan74m + Xcansdm
 data_msp.Recordset("mt") = Xsubtott
 data_msp.Recordset("f1") = Xcan1f
 data_msp.Recordset("f1a4") = Xcan14f
 data_msp.Recordset("f5a14") = Xcan514f
 data_msp.Recordset("f15a19") = Xcan1519f
 data_msp.Recordset("f20a44") = Xcan2044f
 data_msp.Recordset("f45a64") = Xcan4564f
 data_msp.Recordset("f65a74") = Xcan6574f
 data_msp.Recordset("f74") = Xcan74f
 data_msp.Recordset("fsd") = Xcansdf
 Xsubtottt = Xcan1f + Xcan14f + Xcan514f + Xcan1519f + Xcan2044f + Xcan4564f + Xcan6574f + Xcan74f + Xcansdf
 data_msp.Recordset("ft") = Xsubtottt
 data_msp.Recordset("tot") = Xsubtott + Xsubtottt
 data_msp.Recordset.Update
 data_msp.Refresh
 Xcan1m = 0
 Xcan14m = 0
 Xcan514m = 0
 Xcan1519m = 0
 Xcan2044m = 0
 Xcan4564m = 0
 Xcan6574m = 0
 Xcan74m = 0
 Xcansdm = 0
 Xsubtott = 0
 Xsubtott = 0
 Xcan1f = 0
 Xcan14f = 0
 Xcan514f = 0
 Xcan1519f = 0
 Xcan2044f = 0
 Xcan4564f = 0
 Xcan6574f = 0
 Xcan74f = 0
 Xcansdf = 0
 Xsubtottt = 0
'' AMARILLOS
data_lin.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "# And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codzon in (1,2,3) and base =" & 0 & " and codmot ='" & "A" & "' and cancela is null"
data_lin.Refresh
If data_lin.Recordset.RecordCount > 0 Then
   data_lin.Recordset.MoveFirst
   Do While Not data_lin.Recordset.EOF
      If data_lin.Recordset("categ") = "CASH" Or _
         data_lin.Recordset("categ") = "CPS" Or _
         data_lin.Recordset("categ") = "SEMM1" Or _
         data_lin.Recordset("categ") = "1727" Or _
         data_lin.Recordset("categ") = "CCNOS" Or _
         data_lin.Recordset("categ") = "UNIVS" Or _
         data_lin.Recordset("categ") = "IMPNO" Or _
         data_lin.Recordset("categ") = "HEVAN" Or _
         data_lin.Recordset("categ") = "HEVANO" Or _
         data_lin.Recordset("categ") = "CCNRE" Or _
         data_lin.Recordset("categ") = "SMINR" Or _
         data_lin.Recordset("categ") = "SMIN" Or _
         data_lin.Recordset("categ") = "GANOS" Or _
         data_lin.Recordset("categ") = "911" Or _
         data_lin.Recordset("categ") = "911B" Or _
         data_lin.Recordset("categ") = "1727B" Or _
         data_lin.Recordset("categ") = "UCM" Or _
         data_lin.Recordset("categ") = "RETMI" Then
      Else
         If IsNull(data_lin.Recordset("cancela")) = True Then
'            If IsNull(data_lin.Recordset("hh")) = True Then
               If data_lin.Recordset("matric") > 0 Then
                  If cbocat.ListIndex = 2 Then
                     data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("matric") & " and cl_grupo >=" & 100 & " and cl_grupo <=" & 115
                  End If
                  If cbocat.ListIndex = 3 Then
                     data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("matric") & " and cl_grupo >=" & 300 & " and cl_grupo <=" & 320
                  End If
                  If cbocat.ListIndex = 4 Then
                     data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("matric") & " and cl_grupo >=" & 600 & " and cl_grupo <=" & 640
                  End If
                  data_cli.Refresh
                  If data_cli.Recordset.RecordCount > 0 Then
                     If Check3.Value = 1 Then
                        data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("categ") & "' and cnv_grupo <>'" & Null & "'"
                        data_conv.Refresh
                     Else
                        data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("categ") & "'"
                        data_conv.Refresh
                     End If
                     If data_conv.Recordset.RecordCount > 0 Then
                        If data_conv.Recordset("cnv_cant_r") = 1 And data_conv.Recordset("cnv_colrec") = "M" Then
                        Else
                           If IsNull(data_cli.Recordset("cl_sexo")) = False Then
                              If data_cli.Recordset("cl_sexo") = 1 Then
                                 If data_lin.Recordset("unied") < 3 Then
                                    Xcan1m = Xcan1m + 1
                                 Else
                                    If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                       Xcan14m = Xcan14m + 1
                                    Else
                                       If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                          Xcan514m = Xcan514m + 1
                                       Else
                                          If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                             Xcan1519m = Xcan1519m + 1
                                          Else
                                             If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                                Xcan2044m = Xcan2044m + 1
                                             Else
                                                If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                   Xcan4564m = Xcan4564m + 1
                                                Else
                                                   If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                      Xcan6574m = Xcan6574m + 1
                                                   Else
                                                      If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                         Xcan74m = Xcan74m + 1
                                                      Else
                                                         Xcansdm = Xcansdm + 1
                                                      End If
                                                   End If
                                                End If
                                             End If
                                          End If
                                       End If
                                    End If
                                 End If
                              Else
                                 If data_lin.Recordset("unied") < 3 Then
                                    Xcan1f = Xcan1f + 1
                                 Else
                                    If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                       Xcan14f = Xcan14f + 1
                                    Else
                                       If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                          Xcan514f = Xcan514f + 1
                                       Else
                                          If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                             Xcan1519f = Xcan1519f + 1
                                          Else
                                             If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                                Xcan2044f = Xcan2044f + 1
                                             Else
                                                If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                   Xcan4564f = Xcan4564f + 1
                                                Else
                                                   If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                      Xcan6574f = Xcan6574f + 1
                                                   Else
                                                      If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                         Xcan74f = Xcan74f + 1
                                                      Else
                                                         Xcansdf = Xcansdf + 1
                                                      End If
                                                   End If
                                                End If
                                             End If
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                           Else
                              Xcansd = Xcansd + 1
                           End If
                        End If
                     End If
                  Else
'                     Xcansd = Xcansd + 1
                  End If
               Else
'                  Xcansd = Xcansd + 1
               End If
         End If
      End If
      data_lin.Recordset.MoveNext
   Loop
End If
 data_msp.Recordset.Edit
 data_msp.Recordset("DESC5") = "URGENCIA"
 data_msp.Recordset("m1a") = Xcan1m
 data_msp.Recordset("m1a4a") = Xcan14m
 data_msp.Recordset("m5a14a") = Xcan514m
 data_msp.Recordset("m15a19a") = Xcan1519m
 data_msp.Recordset("m20a44a") = Xcan2044m
 data_msp.Recordset("m45a64a") = Xcan4564m
 data_msp.Recordset("m65a74a") = Xcan6574m
 data_msp.Recordset("m74a") = Xcan74m
 data_msp.Recordset("msda") = Xcansdm
 Xsubtott = Xcan1m + Xcan14m + Xcan514m + Xcan1519m + Xcan2044m + Xcan4564m + Xcan6574m + Xcan74m + Xcansdm
 data_msp.Recordset("mta") = Xsubtott
 data_msp.Recordset("f1a") = Xcan1f
 data_msp.Recordset("f1a4a") = Xcan14f
 data_msp.Recordset("f5a14a") = Xcan514f
 data_msp.Recordset("f15a19a") = Xcan1519f
 data_msp.Recordset("f20a44a") = Xcan2044f
 data_msp.Recordset("f45a64a") = Xcan4564f
 data_msp.Recordset("f65a74a") = Xcan6574f
 data_msp.Recordset("f74a") = Xcan74f
 data_msp.Recordset("fsda") = Xcansdf
 Xsubtottt = Xcan1f + Xcan14f + Xcan514f + Xcan1519f + Xcan2044f + Xcan4564f + Xcan6574f + Xcan74f + Xcansdf
 data_msp.Recordset("fta") = Xsubtottt
 data_msp.Recordset("tota") = Xsubtott + Xsubtottt
 data_msp.Recordset.Update
 Xcan1m = 0
 Xcan14m = 0
 Xcan514m = 0
 Xcan1519m = 0
 Xcan2044m = 0
 Xcan4564m = 0
 Xcan6574m = 0
 Xcan74m = 0
 Xcansdm = 0
 Xsubtott = 0
 Xsubtott = 0
 Xcan1f = 0
 Xcan14f = 0
 Xcan514f = 0
 Xcan1519f = 0
 Xcan2044f = 0
 Xcan4564f = 0
 Xcan6574f = 0
 Xcan74f = 0
 Xcansdf = 0
 Xsubtottt = 0
''VERDES
data_lin.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codzon in (1,2,3) and base =" & 0 & " and codmot ='" & "V" & "' and cancela is null"
data_lin.Refresh
If data_lin.Recordset.RecordCount > 0 Then
   data_lin.Recordset.MoveFirst
   Do While Not data_lin.Recordset.EOF
      If data_lin.Recordset("categ") = "CASH" Or _
         data_lin.Recordset("categ") = "CPS" Or _
         data_lin.Recordset("categ") = "SEMM1" Or _
         data_lin.Recordset("categ") = "1727" Or _
         data_lin.Recordset("categ") = "CCNOS" Or _
         data_lin.Recordset("categ") = "UNIVS" Or _
         data_lin.Recordset("categ") = "IMPNO" Or _
         data_lin.Recordset("categ") = "HEVAN" Or _
         data_lin.Recordset("categ") = "HEVANO" Or _
         data_lin.Recordset("categ") = "CCNRE" Or _
         data_lin.Recordset("categ") = "SMINR" Or _
         data_lin.Recordset("categ") = "SMIN" Or _
         data_lin.Recordset("categ") = "GANOS" Or _
         data_lin.Recordset("categ") = "911" Or _
         data_lin.Recordset("categ") = "911B" Or _
         data_lin.Recordset("categ") = "1727B" Or _
         data_lin.Recordset("categ") = "UCM" Or _
         data_lin.Recordset("categ") = "RETMI" Then
      Else
         If IsNull(data_lin.Recordset("cancela")) = True Then
'            If IsNull(data_lin.Recordset("hh")) = True Then
               If data_lin.Recordset("matric") > 0 Then
                  If cbocat.ListIndex = 2 Then
                     data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("matric") & " and cl_grupo >=" & 100 & " and cl_grupo <=" & 115
                  End If
                  If cbocat.ListIndex = 3 Then
                     data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("matric") & " and cl_grupo >=" & 300 & " and cl_grupo <=" & 320
                  End If
                  If cbocat.ListIndex = 4 Then
                     data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("matric") & " and cl_grupo >=" & 600 & " and cl_grupo <=" & 640
                  End If
                  data_cli.Refresh
                  If data_cli.Recordset.RecordCount > 0 Then
                     If Check3.Value = 1 Then
                        data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("categ") & "' and cnv_grupo <>'" & Null & "'"
                        data_conv.Refresh
                     Else
                        data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("categ") & "'"
                        data_conv.Refresh
                     End If
                     If data_conv.Recordset.RecordCount > 0 Then
                        If data_conv.Recordset("cnv_cant_r") = 1 And data_conv.Recordset("cnv_colrec") = "M" Then
                        Else
                           If IsNull(data_cli.Recordset("cl_sexo")) = False Then
                              If data_cli.Recordset("cl_sexo") = 1 Then
                                 If data_lin.Recordset("unied") < 3 Then
                                    Xcan1m = Xcan1m + 1
                                 Else
                                    If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                       Xcan14m = Xcan14m + 1
                                    Else
                                       If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                          Xcan514m = Xcan514m + 1
                                       Else
                                          If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                             Xcan1519m = Xcan1519m + 1
                                          Else
                                             If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                                Xcan2044m = Xcan2044m + 1
                                             Else
                                                If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                   Xcan4564m = Xcan4564m + 1
                                                Else
                                                   If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                      Xcan6574m = Xcan6574m + 1
                                                   Else
                                                      If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                         Xcan74m = Xcan74m + 1
                                                      Else
                                                         Xcansdm = Xcansdm + 1
                                                      End If
                                                   End If
                                                End If
                                             End If
                                          End If
                                       End If
                                    End If
                                 End If
                              Else
                                 If data_lin.Recordset("unied") < 3 Then
                                    Xcan1f = Xcan1f + 1
                                 Else
                                    If data_lin.Recordset("edad") >= 1 And data_lin.Recordset("edad") <= 4 Then
                                       Xcan14f = Xcan14f + 1
                                    Else
                                       If data_lin.Recordset("edad") >= 5 And data_lin.Recordset("edad") <= 14 Then
                                          Xcan514f = Xcan514f + 1
                                       Else
                                          If data_lin.Recordset("edad") >= 15 And data_lin.Recordset("edad") <= 19 Then
                                             Xcan1519f = Xcan1519f + 1
                                          Else
                                             If data_lin.Recordset("edad") >= 20 And data_lin.Recordset("edad") <= 44 Then
                                                Xcan2044f = Xcan2044f + 1
                                             Else
                                                If data_lin.Recordset("edad") >= 45 And data_lin.Recordset("edad") <= 64 Then
                                                   Xcan4564f = Xcan4564f + 1
                                                Else
                                                   If data_lin.Recordset("edad") >= 65 And data_lin.Recordset("edad") <= 74 Then
                                                      Xcan6574f = Xcan6574f + 1
                                                   Else
                                                      If data_lin.Recordset("edad") >= 75 And data_lin.Recordset("edad") <= 120 Then
                                                         Xcan74f = Xcan74f + 1
                                                      Else
                                                         Xcansdf = Xcansdf + 1
                                                      End If
                                                   End If
                                                End If
                                             End If
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                           Else
                              Xcansd = Xcansd + 1
                           End If
                        End If
                     End If
                  Else
'                     Xcansd = Xcansd + 1
                  End If
               Else
'                  Xcansd = Xcansd + 1
               End If
         End If
      End If
      data_lin.Recordset.MoveNext
   Loop
End If
  data_msp.Recordset.Edit
  data_msp.Recordset("DESC6") = "RADIO"
  data_msp.Recordset("m1v") = Xcan1m
  data_msp.Recordset("m1a4v") = Xcan14m
  data_msp.Recordset("m5a14v") = Xcan514m
  data_msp.Recordset("m15a19v") = Xcan1519m
  data_msp.Recordset("m20a44v") = Xcan2044m
  data_msp.Recordset("m45a64v") = Xcan4564m
  data_msp.Recordset("m65a74v") = Xcan6574m
  data_msp.Recordset("m74v") = Xcan74m
  data_msp.Recordset("msdv") = Xcansdm
  Xsubtott = Xcan1m + Xcan14m + Xcan514m + Xcan1519m + Xcan2044m + Xcan4564m + Xcan6574m + Xcan74m + Xcansdm
  data_msp.Recordset("mtv") = Xsubtott
  data_msp.Recordset("f1v") = Xcan1f
  data_msp.Recordset("f1a4v") = Xcan14f
  data_msp.Recordset("f5a14v") = Xcan514f
  data_msp.Recordset("f15a19v") = Xcan1519f
  data_msp.Recordset("f20a44v") = Xcan2044f
  data_msp.Recordset("f45a64v") = Xcan4564f
  data_msp.Recordset("f65a74v") = Xcan6574f
  data_msp.Recordset("f74v") = Xcan74f
  data_msp.Recordset("fsdv") = Xcansdf
  Xsubtottt = Xcan1f + Xcan14f + Xcan514f + Xcan1519f + Xcan2044f + Xcan4564f + Xcan6574f + Xcan74f + Xcansdf
  data_msp.Recordset("ftv") = Xsubtottt
  data_msp.Recordset("totv") = Xsubtott + Xsubtottt
  data_msp.Recordset.Update
  Xcan1m = 0
  Xcan14m = 0
  Xcan514m = 0
  Xcan1519m = 0
  Xcan2044m = 0
  Xcan4564m = 0
  Xcan6574m = 0
  Xcan74m = 0
  Xcansdm = 0
  Xsubtott = 0
  Xsubtott = 0
  Xcan1f = 0
  Xcan14f = 0
  Xcan514f = 0
  Xcan1519f = 0
  Xcan2044f = 0
  Xcan4564f = 0
  Xcan6574f = 0
  Xcan74f = 0
  Xcansdf = 0
  Xsubtottt = 0

  Command7_Click
  
End Sub

Private Sub Command7_Click()
Dim Xlaedad As Double
data_lin.ConnectionString = "dsn=" & Xconexrmt

'data_lin.DatabaseName = App.Path & "\sapp.mdb"
Xcan1m = 0
Xcan14m = 0
Xcan514m = 0
Xcan1519m = 0
Xcan2044m = 0
Xcan4564m = 0
Xcan6574m = 0
Xcan74m = 0
Xcansdm = 0
Xcan1f = 0
Xcan14f = 0
Xcan514f = 0
Xcan1519f = 0
Xcan2044f = 0
Xcan4564f = 0
Xcan6574f = 0
Xcan74f = 0
Xcansdf = 0


 data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cod_prod in (10003,10005)"
 data_lin.Refresh
 If data_lin.Recordset.RecordCount > 0 Then
    data_lin.Recordset.MoveFirst
    Do While Not data_lin.Recordset.EOF
       If data_lin.Recordset("cod_cli") > 0 Then
          If cbocat.ListIndex = 2 Then
             data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("cod_cli") & " and cl_grupo >=" & 100 & " and cl_grupo <=" & 115
          End If
          If cbocat.ListIndex = 3 Then
             data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("cod_cli") & " and cl_grupo >=" & 300 & " and cl_grupo <=" & 320
          End If
          If cbocat.ListIndex = 4 Then
             If Check3.Value = 1 Then
                data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("cod_cli") & " and cl_grupo >=" & 600 & " and cl_grupo <=" & 640
             Else
                data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("cod_cli") & " and cl_grupo >=" & 600 & " and cl_grupo <=" & 640
             End If
          End If
          data_cli.Refresh
          If data_cli.Recordset.RecordCount > 0 Then
             If Check3.Value = 1 Then
                data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "' and cnv_grupo <>'" & Null & "'"
                data_conv.Refresh
             Else
                data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                data_conv.Refresh
             End If
             If data_conv.Recordset.RecordCount > 0 Then
                If IsNull(data_cli.Recordset("cl_sexo")) = False Then
                   If data_cli.Recordset("cl_sexo") = 1 Then
                      If IsNull(data_cli.Recordset("cl_fnac")) = False Then
                         Xlaedad = CDate(mh.Text) - data_cli.Recordset("cl_fnac")
                         Xlaedad = Xlaedad / 365
                         If Xlaedad < 1 Then
                            Xcan1m = Xcan1m + 1
                         Else
                            If Xlaedad >= 1 And Xlaedad <= 4 Then
                               Xcan14m = Xcan14m + 1
                            Else
                               If Xlaedad >= 5 And Xlaedad <= 14 Then
                                  Xcan514m = Xcan514m + 1
                               Else
                                  If Xlaedad >= 15 And Xlaedad <= 19 Then
                                     Xcan1519m = Xcan1519m + 1
                                  Else
                                     If Xlaedad >= 20 And Xlaedad <= 44 Then
                                        Xcan2044m = Xcan2044m + 1
                                     Else
                                        If Xlaedad >= 45 And Xlaedad <= 64 Then
                                           Xcan4564m = Xcan4564m + 1
                                        Else
                                           If Xlaedad >= 65 And Xlaedad <= 74 Then
                                              Xcan6574m = Xcan6574m + 1
                                           Else
                                              If Xlaedad >= 75 And Xlaedad <= 120 Then
                                                 Xcan74m = Xcan74m + 1
                                              Else
                                                 Xcansdm = Xcansdm + 1
                                              End If
                                           End If
                                       End If
                                     End If
                                  End If
                               End If
                            End If
                         End If
                      Else
                         Xcansdm = Xcansdm + 1
                      End If
                   Else
                      If IsNull(data_cli.Recordset("cl_fnac")) = False Then
                         Xlaedad = CDate(mh.Text) - data_cli.Recordset("cl_fnac")
                         Xlaedad = Xlaedad / 365
                         If Xlaedad < 1 Then
                            Xcan1f = Xcan1f + 1
                         Else
                            If Xlaedad >= 1 And Xlaedad <= 4 Then
                               Xcan14f = Xcan14f + 1
                            Else
                               If Xlaedad >= 5 And Xlaedad <= 14 Then
                                  Xcan514f = Xcan514f + 1
                               Else
                                  If Xlaedad >= 15 And Xlaedad <= 19 Then
                                     Xcan1519f = Xcan1519f + 1
                                  Else
                                     If Xlaedad >= 20 And Xlaedad <= 44 Then
                                        Xcan2044f = Xcan2044f + 1
                                     Else
                                        If Xlaedad >= 45 And Xlaedad <= 64 Then
                                           Xcan4564f = Xcan4564f + 1
                                        Else
                                           If Xlaedad >= 65 And Xlaedad <= 74 Then
                                              Xcan6574f = Xcan6574f + 1
                                           Else
                                              If Xlaedad >= 75 And Xlaedad <= 120 Then
                                                 Xcan74f = Xcan74f + 1
                                              Else
                                                 Xcansdf = Xcansdf + 1
                                              End If
                                           End If
                                        End If
                                     End If
                                  End If
                               End If
                            End If
                         End If
                      Else
                         Xcansdf = Xcansdf + 1
                      End If
                   End If
                Else
                   Xcansd = Xcansd + 1
                End If
             End If
          Else
             Xcansd = Xcansd + 1
          End If
       Else
          Xcansd = Xcansd + 1
       End If
       data_lin.Recordset.MoveNext
    Loop
 End If
 '' EMERGENCIA Y URGENCIA
 data_msp.Recordset.AddNew
 data_msp.Recordset("nro") = 3
 data_msp.Recordset("desc") = "ACTIVIDAD CENTRALIZADA"
 data_msp.Recordset("desc2") = "Urgencia/Emergencia"
 data_msp.Recordset("desc4") = "URGENCIA/EMERGENCIA"
 data_msp.Recordset("m1") = Xcan1m
 data_msp.Recordset("m1a4") = Xcan14m
 data_msp.Recordset("m5a14") = Xcan514m
 data_msp.Recordset("m15a19") = Xcan1519m
 data_msp.Recordset("m20a44") = Xcan2044m
 data_msp.Recordset("m45a64") = Xcan4564m
 data_msp.Recordset("m65a74") = Xcan6574m
 data_msp.Recordset("m74") = Xcan74m
 data_msp.Recordset("msd") = Xcansdm
 Xsubtott = Xcan1m + Xcan14m + Xcan514m + Xcan1519m + Xcan2044m + Xcan4564m + Xcan6574m + Xcan74m + Xcansdm
 data_msp.Recordset("mt") = Xsubtott
 data_msp.Recordset("f1") = Xcan1f
 data_msp.Recordset("f1a4") = Xcan14f
 data_msp.Recordset("f5a14") = Xcan514f
 data_msp.Recordset("f15a19") = Xcan1519f
 data_msp.Recordset("f20a44") = Xcan2044f
 data_msp.Recordset("f45a64") = Xcan4564f
 data_msp.Recordset("f65a74") = Xcan6574f
 data_msp.Recordset("f74") = Xcan74f
 data_msp.Recordset("fsd") = Xcansdf
 Xsubtottt = Xcan1f + Xcan14f + Xcan514f + Xcan1519f + Xcan2044f + Xcan4564f + Xcan6574f + Xcan74f + Xcansdf
 data_msp.Recordset("ft") = Xsubtottt
 data_msp.Recordset("tot") = Xsubtott + Xsubtottt
 data_msp.Recordset.Update
 data_msp.Refresh
 Xcan1m = 0
 Xcan14m = 0
 Xcan514m = 0
 Xcan1519m = 0
 Xcan2044m = 0
 Xcan4564m = 0
 Xcan6574m = 0
 Xcan74m = 0
 Xcansdm = 0
 Xsubtott = 0
 Xsubtott = 0
 Xcan1f = 0
 Xcan14f = 0
 Xcan514f = 0
 Xcan1519f = 0
 Xcan2044f = 0
 Xcan4564f = 0
 Xcan6574f = 0
 Xcan74f = 0
 Xcansdf = 0
 Xsubtottt = 0

''''' POLICLINICA MED.GRAL. Y PEDIATRIA
 data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cod_prod in (10001,14001)"
 data_lin.Refresh
 If data_lin.Recordset.RecordCount > 0 Then
    data_lin.Recordset.MoveFirst
    Do While Not data_lin.Recordset.EOF
       If data_lin.Recordset("cod_cli") > 0 Then
          If cbocat.ListIndex = 2 Then
             data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("cod_cli") & " and cl_grupo >=" & 100 & " and cl_grupo <=" & 115
          End If
          If cbocat.ListIndex = 3 Then
             data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("cod_cli") & " and cl_grupo >=" & 300 & " and cl_grupo <=" & 320
          End If
          If cbocat.ListIndex = 4 Then
             data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("cod_cli") & " and cl_grupo >=" & 600 & " and cl_grupo <=" & 640
          End If
          data_cli.Refresh
          If data_cli.Recordset.RecordCount > 0 Then
             If data_cli.Recordset("cl_codconv") = "CASH" Or _
                data_cli.Recordset("cl_codconv") = "CPS" Or _
                data_cli.Recordset("cl_codconv") = "SEMM1" Or _
                data_cli.Recordset("cl_codconv") = "1727" Or _
                data_cli.Recordset("cl_codconv") = "CCNOS" Or _
                data_cli.Recordset("cl_codconv") = "UNIVS" Or _
                data_cli.Recordset("cl_codconv") = "IMPNO" Or _
                data_cli.Recordset("cl_codconv") = "HEVAN" Or _
                data_cli.Recordset("cl_codconv") = "HEVANO" Or _
                data_cli.Recordset("cl_codconv") = "CCNRE" Or _
                data_cli.Recordset("cl_codconv") = "SMINR" Or _
                data_cli.Recordset("cl_codconv") = "SMIN" Or _
                data_cli.Recordset("cl_codconv") = "GANOS" Or _
                data_cli.Recordset("cl_codconv") = "911" Or _
                data_cli.Recordset("cl_codconv") = "911B" Or _
                data_cli.Recordset("cl_codconv") = "1727B" Or _
                data_cli.Recordset("cl_codconv") = "UCM" Or _
                data_cli.Recordset("cl_codconv") = "RETMI" Then
             Else
                If Check3.Value = 1 Then
                   data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "' and cnv_grupo <>'" & Null & "'"
                   data_conv.Refresh
                Else
                   data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("cl_codconv") & "'"
                   data_conv.Refresh
                End If
                If data_conv.Recordset.RecordCount > 0 Then
                
                    If IsNull(data_cli.Recordset("cl_sexo")) = False Then
                       If data_cli.Recordset("cl_sexo") = 1 Then
                          If IsNull(data_cli.Recordset("cl_fnac")) = False Then
                             Xlaedad = CDate(mh.Text) - data_cli.Recordset("cl_fnac")
                             Xlaedad = Xlaedad / 365
                             If Xlaedad < 1 Then
                                Xcan1m = Xcan1m + 1
                             Else
                                If Xlaedad >= 1 And Xlaedad <= 4 Then
                                   Xcan14m = Xcan14m + 1
                                Else
                                   If Xlaedad >= 5 And Xlaedad <= 14 Then
                                      Xcan514m = Xcan514m + 1
                                   Else
                                      If Xlaedad >= 15 And Xlaedad <= 19 Then
                                         Xcan1519m = Xcan1519m + 1
                                      Else
                                         If Xlaedad >= 20 And Xlaedad <= 44 Then
                                            Xcan2044m = Xcan2044m + 1
                                         Else
                                            If Xlaedad >= 45 And Xlaedad <= 64 Then
                                               Xcan4564m = Xcan4564m + 1
                                            Else
                                               If Xlaedad >= 65 And Xlaedad <= 74 Then
                                                  Xcan6574m = Xcan6574m + 1
                                               Else
                                                  If Xlaedad >= 75 And Xlaedad <= 120 Then
                                                     Xcan74m = Xcan74m + 1
                                                  Else
                                                     Xcansdm = Xcansdm + 1
                                                  End If
                                               End If
                                           End If
                                         End If
                                      End If
                                   End If
                                End If
                             End If
                          Else
                             Xcansdm = Xcansdm + 1
                          End If
                       Else
                          If IsNull(data_cli.Recordset("cl_fnac")) = False Then
                             Xlaedad = CDate(mh.Text) - data_cli.Recordset("cl_fnac")
                             Xlaedad = Xlaedad / 365
                             If Xlaedad < 1 Then
                                Xcan1f = Xcan1f + 1
                             Else
                                If Xlaedad >= 1 And Xlaedad <= 4 Then
                                   Xcan14f = Xcan14f + 1
                                Else
                                   If Xlaedad >= 5 And Xlaedad <= 14 Then
                                      Xcan514f = Xcan514f + 1
                                   Else
                                      If Xlaedad >= 15 And Xlaedad <= 19 Then
                                         Xcan1519f = Xcan1519f + 1
                                      Else
                                         If Xlaedad >= 20 And Xlaedad <= 44 Then
                                            Xcan2044f = Xcan2044f + 1
                                         Else
                                            If Xlaedad >= 45 And Xlaedad <= 64 Then
                                               Xcan4564f = Xcan4564f + 1
                                            Else
                                               If Xlaedad >= 65 And Xlaedad <= 74 Then
                                                  Xcan6574f = Xcan6574f + 1
                                               Else
                                                  If Xlaedad >= 75 And Xlaedad <= 120 Then
                                                     Xcan74f = Xcan74f + 1
                                                  Else
                                                     Xcansdf = Xcansdf + 1
                                                  End If
                                               End If
                                            End If
                                         End If
                                      End If
                                   End If
                                End If
                             End If
                          Else
                             Xcansdf = Xcansdf + 1
                          End If
                       End If
                    Else
                       Xcansd = Xcansd + 1
                    End If
                End If
             End If
          End If
       End If
       data_lin.Recordset.MoveNext
    Loop
 End If
 data_msp.Recordset.MoveLast
 data_msp.Recordset.Edit
 data_msp.Recordset("DESC5") = "MED.GRAL/PEDIATRIA"
 data_msp.Recordset("m1a") = Xcan1m
 data_msp.Recordset("m1a4a") = Xcan14m
 data_msp.Recordset("m5a14a") = Xcan514m
 data_msp.Recordset("m15a19a") = Xcan1519m
 data_msp.Recordset("m20a44a") = Xcan2044m
 data_msp.Recordset("m45a64a") = Xcan4564m
 data_msp.Recordset("m65a74a") = Xcan6574m
 data_msp.Recordset("m74a") = Xcan74m
 data_msp.Recordset("msda") = Xcansdm
 Xsubtott = Xcan1m + Xcan14m + Xcan514m + Xcan1519m + Xcan2044m + Xcan4564m + Xcan6574m + Xcan74m + Xcansdm
 data_msp.Recordset("mta") = Xsubtott
 data_msp.Recordset("f1a") = Xcan1f
 data_msp.Recordset("f1a4a") = Xcan14f
 data_msp.Recordset("f5a14a") = Xcan514f
 data_msp.Recordset("f15a19a") = Xcan1519f
 data_msp.Recordset("f20a44a") = Xcan2044f
 data_msp.Recordset("f45a64a") = Xcan4564f
 data_msp.Recordset("f65a74a") = Xcan6574f
 data_msp.Recordset("f74a") = Xcan74f
 data_msp.Recordset("fsda") = Xcansdf
 Xsubtottt = Xcan1f + Xcan14f + Xcan514f + Xcan1519f + Xcan2044f + Xcan4564f + Xcan6574f + Xcan74f + Xcansdf
 data_msp.Recordset("fta") = Xsubtottt
 data_msp.Recordset("tota") = Xsubtott + Xsubtottt
 data_msp.Recordset.Update
 Xcan1m = 0
 Xcan14m = 0
 Xcan514m = 0
 Xcan1519m = 0
 Xcan2044m = 0
 Xcan4564m = 0
 Xcan6574m = 0
 Xcan74m = 0
 Xcansdm = 0
 Xsubtott = 0
 Xsubtott = 0
 Xcan1f = 0
 Xcan14f = 0
 Xcan514f = 0
 Xcan1519f = 0
 Xcan2044f = 0
 Xcan4564f = 0
 Xcan6574f = 0
 Xcan74f = 0
 Xcansdf = 0
 Xsubtottt = 0

''' POLICLINICA ESPECIALISTAS
 data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cod_prod =" & 2
 data_lin.Refresh
 If data_lin.Recordset.RecordCount > 0 Then
    data_lin.Recordset.MoveFirst
    Do While Not data_lin.Recordset.EOF
       If data_lin.Recordset("cod_cli") > 0 Then
          If cbocat.ListIndex = 2 Then
             data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("cod_cli") & " and cl_grupo >=" & 100 & " and cl_grupo <=" & 115
          End If
          If cbocat.ListIndex = 3 Then
             data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("cod_cli") & " and cl_grupo >=" & 300 & " and cl_grupo <=" & 320
          End If
          If cbocat.ListIndex = 4 Then
             data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("cod_cli") & " and cl_grupo >=" & 600 & " and cl_grupo <=" & 640
          End If
          data_cli.Refresh
          If data_cli.Recordset.RecordCount > 0 Then
             If data_cli.Recordset("cl_codconv") = "CASH" Or _
                data_cli.Recordset("cl_codconv") = "CPS" Or _
                data_cli.Recordset("cl_codconv") = "SEMM1" Or _
                data_cli.Recordset("cl_codconv") = "1727" Or _
                data_cli.Recordset("cl_codconv") = "CCNOS" Or _
                data_cli.Recordset("cl_codconv") = "UNIVS" Or _
                data_cli.Recordset("cl_codconv") = "IMPNO" Or _
                data_cli.Recordset("cl_codconv") = "HEVAN" Or _
                data_cli.Recordset("cl_codconv") = "HEVANO" Or _
                data_cli.Recordset("cl_codconv") = "CCNRE" Or _
                data_cli.Recordset("cl_codconv") = "SMINR" Or _
                data_cli.Recordset("cl_codconv") = "SMIN" Or _
                data_cli.Recordset("cl_codconv") = "GANOS" Or _
                data_cli.Recordset("cl_codconv") = "911" Or _
                data_cli.Recordset("cl_codconv") = "911B" Or _
                data_cli.Recordset("cl_codconv") = "1727B" Or _
                data_cli.Recordset("cl_codconv") = "UCM" Or _
                data_cli.Recordset("cl_codconv") = "RETMI" Then
             Else
                If Check3.Value = 1 Then
                   data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "' and cnv_grupo <>'" & Null & "'"
                   data_conv.Refresh
                Else
                   data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("cl_codconv") & "'"
                   data_conv.Refresh
                End If
                If data_conv.Recordset.RecordCount > 0 Then
                
                    If IsNull(data_cli.Recordset("cl_sexo")) = False Then
                       If data_cli.Recordset("cl_sexo") = 1 Then
                          If IsNull(data_cli.Recordset("cl_fnac")) = False Then
                             Xlaedad = CDate(mh.Text) - data_cli.Recordset("cl_fnac")
                             Xlaedad = Xlaedad / 365
                             If Xlaedad < 1 Then
                                Xcan1m = Xcan1m + 1
                             Else
                                If Xlaedad >= 1 And Xlaedad <= 4 Then
                                   Xcan14m = Xcan14m + 1
                                Else
                                   If Xlaedad >= 5 And Xlaedad <= 14 Then
                                      Xcan514m = Xcan514m + 1
                                   Else
                                      If Xlaedad >= 15 And Xlaedad <= 19 Then
                                         Xcan1519m = Xcan1519m + 1
                                      Else
                                         If Xlaedad >= 20 And Xlaedad <= 44 Then
                                            Xcan2044m = Xcan2044m + 1
                                         Else
                                            If Xlaedad >= 45 And Xlaedad <= 64 Then
                                               Xcan4564m = Xcan4564m + 1
                                            Else
                                               If Xlaedad >= 65 And Xlaedad <= 74 Then
                                                  Xcan6574m = Xcan6574m + 1
                                               Else
                                                  If Xlaedad >= 75 And Xlaedad <= 120 Then
                                                     Xcan74m = Xcan74m + 1
                                                  Else
                                                     Xcansdm = Xcansdm + 1
                                                  End If
                                               End If
                                           End If
                                         End If
                                      End If
                                   End If
                                End If
                             End If
                          Else
                             Xcansdm = Xcansdm + 1
                          End If
                       Else
                          If IsNull(data_cli.Recordset("cl_fnac")) = False Then
                             Xlaedad = CDate(mh.Text) - data_cli.Recordset("cl_fnac")
                             Xlaedad = Xlaedad / 365
                             If Xlaedad < 1 Then
                                Xcan1f = Xcan1f + 1
                             Else
                                If Xlaedad >= 1 And Xlaedad <= 4 Then
                                   Xcan14f = Xcan14f + 1
                                Else
                                   If Xlaedad >= 5 And Xlaedad <= 14 Then
                                      Xcan514f = Xcan514f + 1
                                   Else
                                      If Xlaedad >= 15 And Xlaedad <= 19 Then
                                         Xcan1519f = Xcan1519f + 1
                                      Else
                                         If Xlaedad >= 20 And Xlaedad <= 44 Then
                                            Xcan2044f = Xcan2044f + 1
                                         Else
                                            If Xlaedad >= 45 And Xlaedad <= 64 Then
                                               Xcan4564f = Xcan4564f + 1
                                            Else
                                               If Xlaedad >= 65 And Xlaedad <= 74 Then
                                                  Xcan6574f = Xcan6574f + 1
                                               Else
                                                  If Xlaedad >= 75 And Xlaedad <= 120 Then
                                                     Xcan74f = Xcan74f + 1
                                                  Else
                                                     Xcansdf = Xcansdf + 1
                                                  End If
                                               End If
                                            End If
                                         End If
                                      End If
                                   End If
                                End If
                             End If
                          Else
                             Xcansdf = Xcansdf + 1
                          End If
                       End If
                    Else
                       Xcansd = Xcansd + 1
                    End If
                End If
             End If
          End If
       End If
       data_lin.Recordset.MoveNext
    Loop
 End If
  data_msp.Recordset.Edit
  data_msp.Recordset("DESC6") = "ESPECIALISTAS"
  data_msp.Recordset("m1v") = Xcan1m
  data_msp.Recordset("m1a4v") = Xcan14m
  data_msp.Recordset("m5a14v") = Xcan514m
  data_msp.Recordset("m15a19v") = Xcan1519m
  data_msp.Recordset("m20a44v") = Xcan2044m
  data_msp.Recordset("m45a64v") = Xcan4564m
  data_msp.Recordset("m65a74v") = Xcan6574m
  data_msp.Recordset("m74v") = Xcan74m
  data_msp.Recordset("msdv") = Xcansdm
  Xsubtott = Xcan1m + Xcan14m + Xcan514m + Xcan1519m + Xcan2044m + Xcan4564m + Xcan6574m + Xcan74m + Xcansdm
  data_msp.Recordset("mtv") = Xsubtott
  data_msp.Recordset("f1v") = Xcan1f
  data_msp.Recordset("f1a4v") = Xcan14f
  data_msp.Recordset("f5a14v") = Xcan514f
  data_msp.Recordset("f15a19v") = Xcan1519f
  data_msp.Recordset("f20a44v") = Xcan2044f
  data_msp.Recordset("f45a64v") = Xcan4564f
  data_msp.Recordset("f65a74v") = Xcan6574f
  data_msp.Recordset("f74v") = Xcan74f
  data_msp.Recordset("fsdv") = Xcansdf
  Xsubtottt = Xcan1f + Xcan14f + Xcan514f + Xcan1519f + Xcan2044f + Xcan4564f + Xcan6574f + Xcan74f + Xcansdf
  data_msp.Recordset("ftv") = Xsubtottt
  data_msp.Recordset("totv") = Xsubtott + Xsubtottt
  data_msp.Recordset.Update
  Xcan1m = 0
  Xcan14m = 0
  Xcan514m = 0
  Xcan1519m = 0
  Xcan2044m = 0
  Xcan4564m = 0
  Xcan6574m = 0
  Xcan74m = 0
  Xcansdm = 0
  Xsubtott = 0
  Xsubtott = 0
  Xcan1f = 0
  Xcan14f = 0
  Xcan514f = 0
  Xcan1519f = 0
  Xcan2044f = 0
  Xcan4564f = 0
  Xcan6574f = 0
  Xcan74f = 0
  Xcansdf = 0
  Xsubtottt = 0


  Command8_Click
  
End Sub

Private Sub Command8_Click()
Dim Xcantras As Double
Xcan1m = 0
Xcan14m = 0
Xcan514m = 0
Xcan1519m = 0
Xcan2044m = 0
Xcan4564m = 0
Xcan6574m = 0
Xcan74m = 0
Xcansdm = 0
Xcan1f = 0
Xcan14f = 0
Xcan514f = 0
Xcan1519f = 0
Xcan2044f = 0
Xcan4564f = 0
Xcan6574f = 0
Xcan74f = 0
Xcansdf = 0

'data_inf.DatabaseName = App.Path & "\informes.mdb"
'data_inf.RecordSource = "infcli"
'data_inf.Refresh
'If data_inf.Recordset.RecordCount > 0 Then
'   data_inf.Recordset.MoveFirst
'   Do While Not data_inf.Recordset.EOF
'      data_inf.Recordset.Delete
'      data_inf.Recordset.MoveNext
'   Loop
'End If

If Check1.Value = 1 Then
'   data_lin.DatabaseName = App.Path & "\llamado.mdb"
Else
   data_lin.ConnectionString = "dsn=" & Xconexrmt
'   data_lin.DatabaseName = App.Path & "\sapp.mdb"
End If

Xcantras = 0
 data_msp.RecordSource = "res"
 data_msp.Refresh
 If data_msp.Recordset.RecordCount > 0 Then
    data_msp.Recordset.MoveFirst
    Do While Not data_msp.Recordset.EOF
       data_msp.Recordset.Delete
       data_msp.Recordset.MoveNext
    Loop
 End If
 data_msp.Recordset.AddNew
 data_msp.Recordset("nro") = 4
' data_msp.Recordset("desc") = "TRASLADOS"
' data_msp.Recordset("desc2") = ""
 data_msp.Recordset("cant") = Xcantras
 data_msp.Recordset.Update
 data_msp.Refresh
 Xcantras = 0

''''' 911
 
 data_msp.Recordset.Edit
 data_msp.Recordset("cant2") = Xcantras
 data_msp.Recordset.Update
 data_msp.Refresh
 Xcantras = 0
 
data_lin.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codzon in (1,2,3) and base =" & 0 & " and cancela is null"
data_lin.Refresh
If data_lin.Recordset.RecordCount > 0 Then
   data_lin.Recordset.MoveFirst
   Do While Not data_lin.Recordset.EOF
      If data_lin.Recordset("categ") = "CASH" Or _
         data_lin.Recordset("categ") = "CPS" Or _
         data_lin.Recordset("categ") = "SEMM1" Or _
         data_lin.Recordset("categ") = "1727" Or _
         data_lin.Recordset("categ") = "CCNOS" Or _
         data_lin.Recordset("categ") = "UNIVS" Or _
         data_lin.Recordset("categ") = "IMPNO" Or _
         data_lin.Recordset("categ") = "HEVAN" Or _
         data_lin.Recordset("categ") = "HEVANO" Or _
         data_lin.Recordset("categ") = "CCNRE" Or _
         data_lin.Recordset("categ") = "SMINR" Or _
         data_lin.Recordset("categ") = "SMIN" Or _
         data_lin.Recordset("categ") = "CASANO" Or _
         data_lin.Recordset("categ") = "911" Or _
         data_lin.Recordset("categ") = "911B" Or _
         data_lin.Recordset("categ") = "1727B" Or _
         data_lin.Recordset("categ") = "UCM" Then
      Else
         If data_lin.Recordset("matric") > 0 Then
            If cbocat.ListIndex = 2 Then
               data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("matric") & " and cl_grupo >=" & 100 & " and cl_grupo <=" & 115
            End If
            If cbocat.ListIndex = 3 Then
               data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("matric") & " and cl_grupo >=" & 300 & " and cl_grupo <=" & 320
            End If
            If cbocat.ListIndex = 4 Then
               data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("matric") & " and cl_grupo >=" & 600 & " and cl_grupo <=" & 640
            End If
            data_cli.Refresh
            If data_cli.Recordset.RecordCount > 0 Then
               If Check3.Value = 1 Then
                  data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "' and cnv_grupo <>'" & Null & "'"
                  data_conv.Refresh
               Else
                  data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("cl_codconv") & "'"
                  data_conv.Refresh
               End If
               If data_conv.Recordset.RecordCount > 0 Then
                  If data_conv.Recordset("cnv_cant_r") = 1 And data_conv.Recordset("cnv_colrec") = "M" Then
                     If IsNull(data_conv.Recordset("cnv_grupo")) = True Then
                        Xcantras = Xcantras + 1
                        data_inf.Recordset.AddNew
                        data_inf.Recordset("cl_codigo") = data_lin.Recordset("matric")
                        data_inf.Recordset("cl_Fecing") = data_lin.Recordset("fecha")
                        data_inf.Recordset("cl_codconv") = data_lin.Recordset("categ")
                        data_inf.Recordset("cl_nomconv") = Mid(data_lin.Recordset("nomcat"), 1, 25)
                        data_inf.Recordset("cl_cedula") = data_lin.Recordset("codzon")
                        data_inf.Recordset("cl_telefon") = Mid(data_lin.Recordset("telef"), 1, 15)
                        data_inf.Recordset("cl_direcci") = Mid(data_lin.Recordset("direcc"), 1, 25)
                        data_inf.Recordset.Update
                     Else
                        If data_conv.Recordset("cnv_grupo") = "" Then
                            Xcantras = Xcantras + 1
                            data_inf.Recordset.AddNew
                            data_inf.Recordset("cl_Fecing") = data_lin.Recordset("fecha")
                            data_inf.Recordset("cl_codigo") = data_lin.Recordset("matric")
                            data_inf.Recordset("cl_codconv") = data_lin.Recordset("categ")
                            data_inf.Recordset("cl_nomconv") = Mid(data_lin.Recordset("nomcat"), 1, 25)
                            data_inf.Recordset("cl_cedula") = data_lin.Recordset("codzon")
                            data_inf.Recordset("cl_telefon") = Mid(data_lin.Recordset("telef"), 1, 15)
                            data_inf.Recordset("cl_direcci") = Mid(data_lin.Recordset("direcc"), 1, 25)
                            data_inf.Recordset.Update
                        End If
                     End If
                  End If
               End If
            End If
         End If
      End If
      data_lin.Recordset.MoveNext
   Loop
End If
''' Areas protegidas
 data_msp.Recordset.Edit
 data_msp.Recordset("cant3") = Xcantras
 data_msp.Recordset.Update
 data_msp.Refresh
 Xcantras = 0

Dim XXlamat, Xxveces As Double
XXlamat = 0

data_lin.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and codzon in (1,2,3) and base =" & 0 & " and cancela is null order by matric"
data_lin.Refresh
If data_lin.Recordset.RecordCount > 0 Then
   data_lin.Recordset.MoveFirst
   Do While Not data_lin.Recordset.EOF
      If data_lin.Recordset("categ") = "CASH" Or _
         data_lin.Recordset("categ") = "CPS" Or _
         data_lin.Recordset("categ") = "SEMM1" Or _
         data_lin.Recordset("categ") = "1727" Or _
         data_lin.Recordset("categ") = "CCNOS" Or _
         data_lin.Recordset("categ") = "UNIVS" Or _
         data_lin.Recordset("categ") = "IMPNO" Or _
         data_lin.Recordset("categ") = "HEVAN" Or _
         data_lin.Recordset("categ") = "HEVANO" Or _
         data_lin.Recordset("categ") = "CCNRE" Or _
         data_lin.Recordset("categ") = "SMINR" Or _
         data_lin.Recordset("categ") = "SMIN" Or _
         data_lin.Recordset("categ") = "CASANO" Or _
         data_lin.Recordset("categ") = "911" Or _
         data_lin.Recordset("categ") = "911B" Or _
         data_lin.Recordset("categ") = "1727B" Or _
         data_lin.Recordset("categ") = "UCM" Then
      Else
       If data_lin.Recordset("matric") > 0 Then
          If cbocat.ListIndex = 2 Then
             data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("matric") & " and cl_grupo >=" & 100 & " and cl_grupo <=" & 115
          End If
          If cbocat.ListIndex = 3 Then
             data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("matric") & " and cl_grupo >=" & 300 & " and cl_grupo <=" & 320
          End If
          If cbocat.ListIndex = 4 Then
             data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("matric") & " and cl_grupo >=" & 600 & " and cl_grupo <=" & 640
          End If
          data_cli.Refresh
          If data_cli.Recordset.RecordCount > 0 Then
             If Check3.Value = 1 Then
                data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "' and cnv_grupo <>'" & Null & "'"
                data_conv.Refresh
             Else
                data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("cl_codconv") & "'"
                data_conv.Refresh
             End If
          Else
             data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & "NADA" & "'"
             data_conv.Refresh
          
          End If
          If data_conv.Recordset.RecordCount > 0 Then
          
''          data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("matric") & " and cl_grupo >=" & 301 & " and cl_grupo <=" & 419
''          data_cli.Refresh
            If data_cli.Recordset.RecordCount > 0 Then
               If XXlamat = data_lin.Recordset("matric") Then
                  Xcantras = Xcantras + 1
                  data_inf.Recordset.AddNew
                  data_inf.Recordset("cl_codigo") = data_lin.Recordset("matric")
                  data_inf.Recordset("cl_apellid") = data_lin.Recordset("nombre")
                  data_inf.Recordset("cl_Fecing") = data_lin.Recordset("fecha")
                  data_inf.Recordset("cl_codconv") = data_lin.Recordset("categ")
                  data_inf.Recordset("cl_nomconv") = Mid(data_lin.Recordset("nomcat"), 1, 25)
                  data_inf.Recordset("cl_cedula") = data_lin.Recordset("codzon")
                  data_inf.Recordset("cl_telefon") = Mid(data_lin.Recordset("telef"), 1, 15)
                  data_inf.Recordset("cl_direcci") = Mid(data_lin.Recordset("direcc"), 1, 25)
                  data_inf.Recordset.Update
               End If
            End If
          End If
       End If
     End If
     XXlamat = data_lin.Recordset("matric")
     data_lin.Recordset.MoveNext
   Loop
End If
''' Reconsultas
 data_msp.Recordset.Edit
 data_msp.Recordset("cant4") = Xcantras
 data_msp.Recordset.Update
 data_msp.Refresh
 Xcantras = 0

   MsgBox "Proceso terminado"
   cr1.ReportFileName = App.Path & "\infplani.rpt"
   cr1.Action = 1

End Sub

Private Sub Command9_Click()
'clientes


'fin de consultas primera parte


End Sub

Private Sub Form_Load()
data_cli.ConnectionString = "dsn=sappnew"
'data_cli.RecordSource = "clientes"
'data_cli.Refresh
data_msp.DatabaseName = App.Path & "\infmsp.mdb"
data_msp.RecordSource = "plani"
data_msp.Refresh
data_conv.ConnectionString = "dsn=sappnew"
data_conv.RecordSource = "convenio"
data_conv.Refresh
'data_emi.DatabaseName = App.Path & "\emisiones.mdb"
data_emi.ConnectionString = "dsn=sappnew"
'data_lin.ConnectionString = "dsn=" & Xconexrmt
data_lin.ConnectionString = "dsn=sappnew"
'data_lin.RecordSource = "linmmdd"
'data_lin.Refresh
data_inflin.DatabaseName = App.Path & "\informes.mdb"
data_inflin.RecordSource = "infvtas"
data_inflin.Refresh
data_inf.DatabaseName = App.Path & "\informes.mdb"
data_inf.RecordSource = "inflla"
data_inf.Refresh

data_tempg.DatabaseName = App.Path & "\informes.mdb"

cr1.ReportFileName = App.Path & "\infporsexo.rpt"

'data_mdbn.ConnectionString = "DSN=sappnew"
'Data3.Connect = "odbc;dsn=sappnew;"



End Sub

Private Sub md_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mh.SetFocus
End If

End Sub

Private Sub mh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cbocat.SetFocus
End If

End Sub

Private Sub txt_flia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'   txt_zd.SetFocus
   b_proc.SetFocus
End If

End Sub

Private Sub txt_zd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_zh.SetFocus
End If

End Sub

Private Sub txt_zh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_proc.SetFocus
End If

End Sub
