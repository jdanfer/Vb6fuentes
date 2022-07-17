VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infabm 
   BackColor       =   &H00808000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe de ALTAS, BAJAS, MODIFICACIONES de socios"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7200
   Icon            =   "infabm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport crpromosn 
      Left            =   2880
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport crpromos 
      Left            =   2880
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_infc 
      Caption         =   "data_infc"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5760
      Visible         =   0   'False
      Width           =   2775
   End
   Begin Crystal.CrystalReport crmcobd 
      Left            =   5760
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport crbconn 
      Left            =   5040
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport crbcond 
      Left            =   4920
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport craconn 
      Left            =   4320
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport cracond 
      Left            =   3840
      Top             =   5880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport crbajradn 
      Left            =   4320
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport crbajradd 
      Left            =   4320
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport crbajpron 
      Left            =   3840
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport crbajprod 
      Left            =   3360
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport crbajcobn 
      Left            =   3840
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport crbajcobd 
      Left            =   3360
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport crradn 
      Left            =   5640
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport crradd 
      Left            =   5160
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport crpron 
      Left            =   3720
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport crprod 
      Left            =   3240
      Top             =   5160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data datau 
      Caption         =   "datau"
      Connect         =   "Access"
      DatabaseName    =   "C:\WINDOWS\usapp.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "usuarioact"
      Top             =   5160
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FF8080&
      Caption         =   "Opciones de informe:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   4440
      Width           =   6735
      Begin VB.OptionButton opdet 
         BackColor       =   &H00C0C000&
         Caption         =   "Detallado"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   18
         Top             =   360
         Width           =   2775
      End
      Begin VB.OptionButton opnum 
         BackColor       =   &H00C0C000&
         Caption         =   "Numérico"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   2775
      End
   End
   Begin Crystal.CrystalReport crcobnum 
      Left            =   3360
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3360
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport crcobd 
      Left            =   2880
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton btn_sale 
      BackColor       =   &H00FFFFFF&
      DownPicture     =   "infabm.frx":0442
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5760
      MaskColor       =   &H00FFC0FF&
      Picture         =   "infabm.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Salir"
      Top             =   5520
      UseMaskColor    =   -1  'True
      Width           =   735
   End
   Begin VB.CommandButton btn_acep 
      BackColor       =   &H00FFFFFF&
      DownPicture     =   "infabm.frx":0CD6
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      MaskColor       =   &H00FFC0C0&
      Picture         =   "infabm.frx":1118
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Procesar"
      Top             =   5520
      Width           =   735
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Caption         =   "Ordenado por..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1815
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   6735
      Begin VB.OptionButton oppromos 
         BackColor       =   &H00C0C000&
         Caption         =   "Por promociones"
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
         Left            =   120
         TabIndex        =   19
         Top             =   1440
         Width           =   2775
      End
      Begin MSAdodcLib.Adodc data_cli 
         Height          =   330
         Left            =   3600
         Top             =   480
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
      Begin VB.OptionButton opconv 
         BackColor       =   &H00C0C000&
         Caption         =   "Por Convenio"
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
         Left            =   3480
         TabIndex        =   13
         Top             =   960
         Width           =   2775
      End
      Begin VB.OptionButton oprad 
         BackColor       =   &H00C0C000&
         Caption         =   "Por Radio"
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
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   2775
      End
      Begin VB.OptionButton oppro 
         BackColor       =   &H00C0C000&
         Caption         =   "Por Promotor"
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
         Left            =   3480
         TabIndex        =   11
         Top             =   480
         Width           =   2775
      End
      Begin VB.OptionButton opcob 
         BackColor       =   &H00C0C000&
         Caption         =   "Por Cobrador"
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
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "Actividad de socios a listar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   6735
      Begin VB.OptionButton opbaja 
         BackColor       =   &H00C0C000&
         Caption         =   "Bajas"
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
         Left            =   2400
         TabIndex        =   7
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton opmod 
         BackColor       =   &H00C0C000&
         Caption         =   "Modificaciones"
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
         Left            =   4680
         TabIndex        =   8
         Top             =   480
         Width           =   1815
      End
      Begin VB.OptionButton opalta 
         BackColor       =   &H00C0C000&
         Caption         =   "Altas"
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
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Rango de Fechas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6735
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   720
         Visible         =   0   'False
         Width           =   2535
      End
      Begin MSMask.MaskEdBox fhasta 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14346
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   4920
         TabIndex        =   4
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
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
      Begin MSMask.MaskEdBox fdesde 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14346
            SubFormatType   =   3
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
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
      Begin VB.Label Label2 
         BackColor       =   &H00808000&
         Caption         =   "Hasta:"
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
         Left            =   3600
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00808000&
         Caption         =   "Desde:"
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
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   1560
      Picture         =   "infabm.frx":16A2
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   1335
   End
End
Attribute VB_Name = "frm_infabm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btn_acep_Click()
btn_acep.Enabled = False
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")
Dim Xpromos As Integer
Xpromos = 0
MiBaseact.Execute "Delete * from infcli"

data_infc.RecordSource = "infcli"
data_infc.Refresh
'If data_infc.Recordset.RecordCount > 0 Then
'   data_infc.Recordset.MoveFirst
'   Do While Not data_infc.Recordset.EOF
'      data_infc.Recordset.Delete
'      data_infc.Recordset.MoveNext
'   Loop
'End If
frm_infabm.MousePointer = 11
If fdesde.Text <> "__/__/____" Then
   If fhasta.Text <> "__/__/____" Then
      If opalta.Value = True Then
         If oppromos.Value = True Then
            data_cli.RecordSource = "select * from clientes where cl_fecing >= '" & Format(fdesde.Text, "yyyy-mm-dd") & "' And cl_fecing <= '" & Format(fhasta.Text, "yyyy-mm-dd") & "' and idpromos is not null order by cl_nrocobr"
            data_cli.Refresh
         Else
            data_cli.RecordSource = "Select * from clientes where estado <>" & 2 & " and estado <>" & 3 & " and cl_fecing >= '" & Format(fdesde.Text, "yyyy-mm-dd") & "' And cl_fecing <= '" & Format(fhasta.Text, "yyyy-mm-dd") & "' order by cl_nrocobr"
            data_cli.Refresh
         End If
      Else
        If opbaja.Value = True Then
           data_cli.RecordSource = "Select * from clientes where fecha_baja >= '" & Format(fdesde.Text, "yyyy-mm-dd") & "' And fecha_baja <= '" & Format(fhasta.Text, "yyyy-mm-dd") & "' order by cl_nrocobr"
           data_cli.Refresh
        Else
           If opmod.Value = True Then
              data_cli.RecordSource = "Select * from clientes where fecha_modi >= '" & Format(fdesde.Text, "yyyy-mm-dd") & "' And fecha_modi <= '" & Format(fhasta.Text, "yyyy-mm-dd") & "' order by cl_nrocobr"
              data_cli.Refresh
           Else
              MsgBox "No seleccionó opción de listado", vbInformation, "Mensaje"
           End If
        End If
      End If
        If data_cli.Recordset.RecordCount > 0 Then
           data_cli.Recordset.MoveFirst
           Do While Not data_cli.Recordset.EOF
              data_infc.Recordset.AddNew
              data_infc.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
              data_infc.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
              data_infc.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
              data_infc.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
              data_infc.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
              data_infc.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
              data_infc.Recordset("cl_nomconv") = data_cli.Recordset("cl_nomconv")
              data_infc.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
              data_infc.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
              data_infc.Recordset("cl_grupo") = data_cli.Recordset("cl_grupo")
              data_infc.Recordset("cl_nro_sup") = 0
              If oppromos.Value = True Then
                 Xpromos = data_cli.Recordset("idpromos")
                 If Xpromos > 0 Then
                    Data1.RecordSource = "select * from promocion_gpo where id =" & data_cli.Recordset("idpromos")
                    Data1.Refresh
                    If Data1.Recordset.RecordCount > 0 Then
                       data_infc.Recordset("cl_zona") = Data1.Recordset("descrip")
                       data_infc.Recordset("cl_nro_sup") = Xpromos
                    End If
                 Else
                    data_infc.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                 End If
              Else
                 data_infc.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
              End If
              data_infc.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
              data_infc.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
              data_infc.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
              data_infc.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
              data_infc.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
              If IsNull(data_cli.Recordset("fecha_baja")) = False Then
                 data_infc.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
              End If
              If IsNull(data_cli.Recordset("fecha_modi")) = False Then
                 data_infc.Recordset("fecha_modi") = data_cli.Recordset("fecha_modi")
              End If
              data_infc.Recordset.Update
              Xpromos = 0
              data_cli.Recordset.MoveNext
           Loop
        Else
           MsgBox "No se encontraron registros", vbInformation, "Informes"
        End If
   End If
End If
data_infc.RecordSource = "select * from infcli order by cl_codigo"
data_infc.Refresh
If opalta.Value = True Then
    If opcob.Value = True Or oppromos.Value = True Then
       If oppromos.Value = True Then
          If opnum.Value = True Then
             crpromosn.ReportTitle = "ALTAS DESDE: " + fdesde.Text + "  HASTA: " + fhasta.Text
             crpromosn.Action = 1
          Else
             If opdet.Value = True Then
                crpromos.ReportTitle = "ALTAS DESDE: " + fdesde.Text + "  HASTA: " + fhasta.Text
                crpromos.Action = 1
             End If
          End If
       Else
          If opnum.Value = True Then
             crcobnum.ReportTitle = "DESDE: " + fdesde.Text + "  HASTA: " + fhasta.Text
             crcobnum.Action = 1
          Else
             If opdet.Value = True Then
                crcobd.ReportTitle = "DESDE: " + fdesde.Text + "  HASTA: " + fhasta.Text
                crcobd.Action = 1
             End If
          End If
       End If
    End If
    If oppro.Value = True Then
       If opnum.Value = True Then
          crpron.ReportTitle = "DESDE: " + fdesde.Text + "  HASTA: " + fhasta.Text
          crpron.Action = 1
       Else
          If opdet.Value = True Then
             crprod.ReportTitle = "DESDE: " + fdesde.Text + "  HASTA: " + fhasta.Text
             crprod.Action = 1
          End If
       End If
    End If
    If oprad.Value = True Then
       If opnum.Value = True Then
          crradn.ReportTitle = "DESDE: " + fdesde.Text + "  HASTA: " + fhasta.Text
          crradn.Action = 1
       Else
          If opdet.Value = True Then
             crradd.ReportTitle = "DESDE: " + fdesde.Text + "  HASTA: " + fhasta.Text
             crradd.Action = 1
          End If
       End If
    End If
    If opconv.Value = True Then
       If opnum.Value = True Then
          craconn.ReportTitle = "DESDE: " + fdesde.Text + "  HASTA: " + fhasta.Text
          craconn.Action = 1
       Else
          If opdet.Value = True Then
             cracond.ReportTitle = "DESDE: " + fdesde.Text + "  HASTA: " + fhasta.Text
             cracond.Action = 1
          End If
       End If
    End If
End If

If opbaja.Value = True Then
    If opcob.Value = True Or oppromos.Value = True Then
       If oppromos.Value = True Then
          If opnum.Value = True Then
             crpromosn.ReportTitle = "BAJAS DESDE: " + fdesde.Text + "  HASTA: " + fhasta.Text
             crpromosn.Action = 1
          Else
             If opdet.Value = True Then
                crpromos.ReportTitle = "BAJAS DESDE: " + fdesde.Text + "  HASTA: " + fhasta.Text
                crpromos.Action = 1
             End If
          End If
       Else
          If opnum.Value = True Then
             crbajcobn.ReportTitle = "DESDE: " + fdesde.Text + "  HASTA: " + fhasta.Text
             crbajcobn.Action = 1
          Else
             If opdet.Value = True Then
                crbajcobd.ReportTitle = "DESDE: " + fdesde.Text + "  HASTA: " + fhasta.Text
                crbajcobd.Action = 1
             End If
          End If
       End If
    End If
    If oppro.Value = True Then
       If opnum.Value = True Then
          crbajpron.ReportTitle = "DESDE: " + fdesde.Text + "  HASTA: " + fhasta.Text
          crbajpron.Action = 1
       Else
          If opdet.Value = True Then
             crbajprod.ReportTitle = "DESDE: " + fdesde.Text + "  HASTA: " + fhasta.Text
             crbajprod.Action = 1
          End If
       End If
    End If
    If oprad.Value = True Then
       If opnum.Value = True Then
          crbajradn.ReportTitle = "DESDE: " + fdesde.Text + "  HASTA: " + fhasta.Text
          crbajradn.Action = 1
       Else
          If opdet.Value = True Then
             crbajradd.ReportTitle = "DESDE: " + fdesde.Text + "  HASTA: " + fhasta.Text
             crbajradd.Action = 1
          End If
       End If
    End If
    If opconv.Value = True Then
       If opnum.Value = True Then
          crbconn.ReportTitle = "DESDE: " + fdesde.Text + "  HASTA: " + fhasta.Text
          crbconn.Action = 1
       Else
          If opdet.Value = True Then
             crbcond.ReportTitle = "DESDE: " + fdesde.Text + "  HASTA: " + fhasta.Text
             crbcond.Action = 1
          End If
       End If
    End If
End If
If opmod.Value = True Then
    If opcob.Value = True Then
       If opnum.Value = True Then
          crmcobd.ReportTitle = "DESDE: " + fdesde.Text + "  HASTA: " + fhasta.Text
          crmcobd.Action = 1
       Else
          If opdet.Value = True Then
             crmcobd.ReportTitle = "DESDE: " + fdesde.Text + "  HASTA: " + fhasta.Text
             crmcobd.Action = 1
          End If
       End If
    End If
End If
btn_acep.Enabled = True
frm_infabm.MousePointer = 0

End Sub

Private Sub btn_sale_Click()
frm_infabm.Hide

End Sub

Private Sub fdesde_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   fhasta.SetFocus
End If

End Sub

Private Sub fhasta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   opalta.SetFocus
End If

End Sub

Private Sub fhasta_LostFocus()
'If IsDate(Format(fhasta.Text, "dd/mm/yyyy")) = False Then
'   MsgBox "Verifique Fecha", vbCritical, "Error"
'   opalta.SetFocus
'End If

End Sub

Private Sub Form_Load()
'data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cli.ConnectionString = "DSN=" & Xconexrmt
'data_infc.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"
data_infc.DatabaseName = App.path & "\informes.mdb"
crbajcobd.ReportFileName = App.path & "\infabmbcob.rpt"
crbajcobn.ReportFileName = App.path & "\infabmbcobn.rpt"
crbajradd.ReportFileName = App.path & "\infabmbrad.rpt"
crbcond.ReportFileName = App.path & "\infabmbcon.rpt"
crbajprod.ReportFileName = App.path & "\infabmbpro.rpt"
crbajpron.ReportFileName = App.path & "\infabmbpron.rpt"
crbajradn.ReportFileName = App.path & "\infabmbradn.rpt"
crbconn.ReportFileName = App.path & "\infabmbconn.rpt"
crmcobd.ReportFileName = App.path & "\infabmmcob.rpt"
crradd.ReportFileName = App.path & "\infabmrad.rpt"
crradn.ReportFileName = App.path & "\infabmradn.rpt"
crprod.ReportFileName = App.path & "\infabmpro.rpt"
crpron.ReportFileName = App.path & "\infabmpron.rpt"
crcobd.ReportFileName = App.path & "\infabmcob.rpt"
crcobnum.ReportFileName = App.path & "\infabmcobn.rpt"
cracond.ReportFileName = App.path & "\infabmcon.rpt"
craconn.ReportFileName = App.path & "\infabmconn.rpt"
crpromos.ReportFileName = App.path & "\infabmbpromos.rpt"
crpromosn.ReportFileName = App.path & "\infabmbpromosn.rpt"

Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub
