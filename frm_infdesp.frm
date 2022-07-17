VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infdesp 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes Despacho"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7455
   Icon            =   "frm_infdesp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   7455
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_inflla 
      Caption         =   "data_inflla"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6960
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_inf2 
      Caption         =   "data_inf2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc data_conv 
      Height          =   330
      Left            =   1320
      Top             =   6600
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
   Begin MSAdodcLib.Adodc data_llam 
      Height          =   375
      Left            =   4200
      Top             =   120
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
      Caption         =   "data_llam"
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
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2160
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C000&
      Caption         =   "Seleccione los datos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   480
      TabIndex        =   21
      Top             =   2280
      Visible         =   0   'False
      Width           =   6495
      Begin VB.ComboBox Combo1 
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
         Left            =   240
         TabIndex        =   24
         Top             =   720
         Width           =   5775
      End
      Begin VB.CommandButton Command1 
         Height          =   495
         Left            =   240
         Picture         =   "frm_infdesp.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1920
         Width           =   615
      End
      Begin VB.CommandButton Command2 
         Height          =   495
         Left            =   1200
         Picture         =   "frm_infdesp.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1920
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Convenios:"
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
         TabIndex        =   25
         Top             =   480
         Width           =   3015
      End
   End
   Begin MSComctlLib.ProgressBar pbar 
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   6120
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSMask.MaskEdBox mhh 
      Height          =   375
      Left            =   5280
      TabIndex        =   16
      Top             =   720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "HH:mm"
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mhd 
      Height          =   375
      Left            =   3720
      TabIndex        =   15
      Top             =   720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "HH:mm"
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   3600
      Top             =   5520
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2880
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton bcan 
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
      Left            =   6360
      Picture         =   "frm_infdesp.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Salir"
      Top             =   6600
      Width           =   615
   End
   Begin VB.CommandButton bace 
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
      Picture         =   "frm_infdesp.frx":1250
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Procesar"
      Top             =   6600
      Width           =   615
   End
   Begin VB.Frame Frame1 
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
      ForeColor       =   &H00FF0000&
      Height          =   4695
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   7215
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Receptor"
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
         Left            =   5640
         TabIndex        =   30
         Top             =   3960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Largador"
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
         Left            =   3840
         TabIndex        =   29
         Top             =   3960
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.OptionButton Option12 
         BackColor       =   &H00C00000&
         Caption         =   "Por usuarios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3840
         TabIndex        =   28
         Top             =   3480
         Width           =   3135
      End
      Begin VB.OptionButton Option11 
         BackColor       =   &H00C00000&
         Caption         =   "Tiempos en recepción"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   240
         TabIndex        =   27
         Top             =   3480
         Width           =   3135
      End
      Begin VB.OptionButton Option10 
         BackColor       =   &H00C00000&
         Caption         =   "Cancelados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3840
         TabIndex        =   20
         Top             =   2880
         Width           =   3135
      End
      Begin VB.OptionButton Option9 
         BackColor       =   &H00C00000&
         Caption         =   "Llamados con costo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   2880
         Width           =   3135
      End
      Begin VB.OptionButton Option8 
         BackColor       =   &H00C00000&
         Caption         =   "Por Socio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3840
         TabIndex        =   17
         ToolTipText     =   "Informe de socios que consultaron más de una vez en 72 hs"
         Top             =   2280
         Width           =   3135
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Emitir informe SIN DETALLE"
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
         Top             =   4080
         Width           =   3135
      End
      Begin VB.OptionButton Option7 
         BackColor       =   &H00C00000&
         Caption         =   "Por CONVENIO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3840
         TabIndex        =   12
         ToolTipText     =   "Informe por categorías  (AMB,EMERG, etc)"
         Top             =   1680
         Width           =   3135
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C00000&
         Caption         =   "Por CODIGO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3840
         TabIndex        =   9
         ToolTipText     =   "Informe por Clasificación del llamado"
         Top             =   1080
         Width           =   3135
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C00000&
         Caption         =   "Traslados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3840
         TabIndex        =   8
         Top             =   480
         Width           =   3135
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C00000&
         Caption         =   "Por MOVIL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   2280
         Width           =   3135
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C00000&
         Caption         =   "Por Médico"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   3135
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C00000&
         Caption         =   "Adultos/Pediátricos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1080
         Width           =   3135
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C00000&
         Caption         =   "Llamados por Zona"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   3135
      End
   End
   Begin MSMask.MaskEdBox mh 
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   240
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
      Left            =   3720
      TabIndex        =   1
      Top             =   240
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
      BackColor       =   &H00C00000&
      Caption         =   "Rango de horario:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   720
      Width           =   3615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      X1              =   0
      X2              =   7440
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "Rango de Fechas a informar:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
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
      Top             =   240
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   5160
      Picture         =   "frm_infdesp.frx":17DA
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   1575
   End
End
Attribute VB_Name = "frm_infdesp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bace_Click()
Dim queop, Queopdos As String
Dim Xcantmat As Integer
Dim Xlamatric As Long
Dim Xmin, Xmind, Xdif As Integer
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Dim Xcantveces As String

Dim MiBaseactll As Database
Dim Unasesactll As Workspace
Set Unasesactll = Workspaces(0)
Set MiBaseactll = Unasesactll.OpenDatabase(App.path & "\informes.mdb")

MiBaseactll.Execute "Delete * from inflla"
MiBaseactll.Execute "Delete * from infvtas"
data_inf2.RecordSource = "infvtas"
data_inf2.Refresh
data_inflla.RecordSource = "inflla"
data_inflla.Refresh
frm_infdesp.MousePointer = 11

If md.Text <> "__/__/____" Then
   If mh.Text <> "__/__/____" Then
      frm_infdesp.MousePointer = 11
      If mhd.Text <> "__:__" Then
         If Option5.Value = True Then
            queop = InputBox("Ingrese documento de socio (999 = TODOS):", "Solicitud de datos del socio")
            If queop = "" Then
               data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And hora >='" & mhd.Text & "' And hora <='" & mhh.Text & "' And trasla in (1,2,4,5) And base =" & 0
               data_llam.Refresh
            Else
               Xcantveces = InputBox("Ingrese >= CANTIDAD DE CONSULTAS")
               data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And hora >='" & mhd.Text & "' And hora <='" & mhh.Text & "' And trasla in (1,2,4,5) And base =" & 0
               data_llam.Refresh
            End If
         Else
            If Option8.Value = True Then
               queop = InputBox("Ingrese documento de socio (999 = TODOS):", "Solicitud de datos del socio")
               If queop <> "" Then
                  If Val(queop) <> 999 Then
                     data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And hora >='" & mhd.Text & "' And hora <='" & mhh.Text & "' And matric =" & Val(queop) & " And Base =" & 0 & " and trasla not in(9,10,11)"
                     data_llam.Refresh
                  Else
                     queop = "0"
                     data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And hora >='" & mhd.Text & "' And hora <='" & mhh.Text & "' And matric =" & Val(queop) & " And Base =" & 0 & " and trasla not in(9,10,11) and codmot <>'" & "C" & "'"
                     data_llam.Refresh
                  End If
               Else
                  data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And hora >='" & mhd.Text & "' And hora <='" & mhh.Text & "' And base =" & 0 & " and trasla not in(9,10,11)"
                  data_llam.Refresh
               End If
            Else
               data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And hora >='" & mhd.Text & "' And hora <='" & mhh.Text & "' And base =" & 0 & " and trasla not in(9,10,11)"
               data_llam.Refresh
            End If
         End If
      Else
         If Option5.Value = True Then
            queop = InputBox("Ingrese documento de socio (999 = TODOS):", "Solicitud de datos del socio")
            If queop = "" Then
               data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And trasla in (1,2,4,5) And base =" & 0
               data_llam.Refresh
            Else
               Xcantveces = InputBox("Ingrese >= CANTIDAD DE CONSULTAS")
               data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And trasla in (1,2,4,5) And base =" & 0
               data_llam.Refresh
            End If
         Else
            If Option8.Value = True Then
               queop = InputBox("Ingrese documento de socio (999 = TODOS):", "Solicitud de datos del socio")
               If queop <> "" Then
                  If Val(queop) <> 999 Then
                     data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And matric =" & Val(queop) & " And base =" & 0 & " and trasla not in(9,10,11)"
                     data_llam.Refresh
                  Else
                     data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and trasla not in(9,10,11) and codmot <>'" & "C" & "' order by matric"
                     data_llam.Refresh
                  End If
               Else
                  data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And base =" & 0 & " and trasla not in(9,10,11) order by matric"
                  data_llam.Refresh
               End If
            Else
               If Option7.Value = True Then
                  If Text1.Text = "TODOS" Or Text1.Text = "" Then
                     data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and base =" & 0 & " and trasla not in(9,10,11) order by matric"
                     data_llam.Refresh
                  Else
                     data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and categ ='" & Text1.Text & "' order by matric"
                     data_llam.Refresh
                  End If
               Else
                  data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and base =" & 0 & " and trasla not in(9,10,11) order by matric"
                  data_llam.Refresh
               End If
            End If
         End If
      End If
      If Option3.Value = True Then
         Queopdos = InputBox("Ingrese código de médico a listar (999 = TODOS)", "Datos para informe")
         If Queopdos = "" Then
         Else
            If Val(Queopdos) <> 999 Then
               If mhd.Text <> "__:__" Then
                  data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And hora >='" & mhd.Text & "' And hora <='" & mhh.Text & "' And codmed =" & Val(Queopdos) & " And base =" & 0
                  data_llam.Refresh
               Else
                  data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And codmed =" & Val(Queopdos) & " And base =" & 0
                  data_llam.Refresh
               End If
            End If
         End If
      End If
      If Option4.Value = True Then
         Queopdos = InputBox("Ingrese número de móvil (999 = TODOS)", "Datos para informe")
         If Queopdos = "" Then
         Else
            If Val(Queopdos) <> 999 Then
               If mhd.Text <> "__:__" Then
                  data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And hora >='" & mhd.Text & "' And hora <='" & mhh.Text & "' And movilpas =" & Val(Queopdos) & " And base =" & 0
                  data_llam.Refresh
               Else
                  data_llam.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And movilpas =" & Val(Queopdos) & " And base =" & 0
                  data_llam.Refresh
               End If
            End If
         End If
      End If
      DoEvents
      If Option8.Value = True Then
         Xcantveces = InputBox("Ingrese >= CANTIDAD DE CONSULTAS")
      End If
      If data_llam.Recordset.RecordCount > 0 Then
         data_llam.Recordset.MoveLast
         pbar.Max = data_llam.Recordset.RecordCount
         pbar.Value = 0
         data_llam.Recordset.MoveFirst
         If Option8.Value = False Then
            Do While Not data_llam.Recordset.EOF
               data_inflla.Recordset.AddNew
               data_inflla.Recordset("fecha") = data_llam.Recordset("fecha")
               data_inflla.Recordset("hora") = data_llam.Recordset("hora")
               data_inflla.Recordset("matric") = data_llam.Recordset("matric")
               data_inflla.Recordset("nombre") = data_llam.Recordset("nombre")
               data_inflla.Recordset("categ") = data_llam.Recordset("categ")
               data_inflla.Recordset("nomcat") = data_llam.Recordset("nomcat")
               data_inflla.Recordset("motcon") = data_llam.Recordset("motcon")
               data_inflla.Recordset("motmov") = data_llam.Recordset("motmov")
               data_inflla.Recordset("enfer") = data_llam.Recordset("enfer")
               If IsNull(data_llam.Recordset("cancela")) = True Then
                  data_inflla.Recordset("cancela") = 0
               Else
                  data_inflla.Recordset("cancela") = data_llam.Recordset("cancela")
               End If
               data_inflla.Recordset("fec_cance") = data_llam.Recordset("fec_cance")
               data_inflla.Recordset("hor_cance") = data_llam.Recordset("hor_cance")
               data_inflla.Recordset("motcance") = data_llam.Recordset("motcance")
               If Option10.Value = True Then
                  If IsNull(data_llam.Recordset("user_cance")) = False Then
                     data_inflla.Recordset("timdes") = data_llam.Recordset("user_cance")
                  End If
               Else
                  data_inflla.Recordset("timdes") = data_llam.Recordset("timdes")
               End If
               data_inflla.Recordset("activo") = data_llam.Recordset("activo")
               data_inflla.Recordset("usuario") = data_llam.Recordset("usuario")
               If IsNull(data_llam.Recordset("activo")) = False Then
                  Xmin = Val(Mid(data_llam.Recordset("activo"), 4, 2))
               Else
                  Xmin = 0
               End If
               If IsNull(data_llam.Recordset("hora")) = False Then
                  Xmind = Val(Mid(data_llam.Recordset("hora"), 4, 2))
               Else
                  Xmind = 0
               End If
               Xdif = Xmin - Xmind
               If Xdif < 0 Then
                  Xdif = Xdif + 60
               End If
               data_inflla.Recordset("timsi") = Xdif
'               data_conv.Recordset.FindFirst "cnv_codigo ='" & data_llam.Recordset("categ") & "'"
'               data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_llam.Recordset("categ") & "'"
'               data_conv.Refresh
'               If data_conv.Recordset.RecordCount > 0 Then
'                   If data_conv.Recordset("cnv_colrec") = "R" Then
'                      data_inflla.Recordset("ncobr") = 1
'                      data_inflla.Recordset("dcobr") = "AMBULATORIO"
'                   Else
'                      If data_conv.Recordset("cnv_colrec") = "A" Then
'                         data_inflla.Recordset("ncobr") = 2
'                         data_inflla.Recordset("dcobr") = "PARCIAL"
'                      Else
'                         If data_conv.Recordset("cnv_colrec") = "M" Or data_conv.Recordset("cnv_colrec") = "V" Then
'                            data_inflla.Recordset("ncobr") = 3
'                            data_inflla.Recordset("dcobr") = "EMERGENCIA"
'                         Else
'                            data_inflla.Recordset("ncobr") = 4
'                            data_inflla.Recordset("dcobr") = "OTROS"
'                         End If
'                      End If
'                   End If
'               Else
'                   data_inflla.Recordset("ncobr") = 4
'                   data_inflla.Recordset("dcobr") = "OTROS"
'               End If
               data_inflla.Recordset("edad") = data_llam.Recordset("edad")
               If data_llam.Recordset("unied") = 3 Then
                  If data_llam.Recordset("edad") <= 14 Then
                     data_inflla.Recordset("unied") = 1
                     data_inflla.Recordset("motmov") = "PEDIATRICOS"
                  Else
                     data_inflla.Recordset("unied") = 3
                     data_inflla.Recordset("motmov") = "ADULTOS"
                  End If
               Else
                  data_inflla.Recordset("unied") = 1
                  data_inflla.Recordset("motmov") = "PEDIATRICOS"
               End If
               data_inflla.Recordset("codzon") = data_llam.Recordset("codzon")
               data_inflla.Recordset("codmot") = data_llam.Recordset("codmot")
               If data_llam.Recordset("codmot") = "V" Then
                  data_inflla.Recordset("descol") = "VERDE"
               End If
               If data_llam.Recordset("codmot") = "A" Then
                  data_inflla.Recordset("descol") = "AMARILLO"
               End If
               If data_llam.Recordset("codmot") = "R" Then
                  data_inflla.Recordset("descol") = "ROJO"
               End If
               If Option7.Value = True Then
                  data_inflla.Recordset("ncobr") = 1
                  data_inflla.Recordset("dcobr") = data_llam.Recordset("categ")
               End If
               data_inflla.Recordset("horpas") = data_llam.Recordset("horpas")
               data_inflla.Recordset("movilpas") = data_llam.Recordset("movilpas")
               data_inflla.Recordset("fec_rea") = data_llam.Recordset("fec_rea")
               data_inflla.Recordset("hor_rea") = data_llam.Recordset("hor_rea")
               data_inflla.Recordset("fec_llega") = data_llam.Recordset("fec_llega")
               data_inflla.Recordset("hor_llega") = data_llam.Recordset("hor_llega")
               data_inflla.Recordset("trasla") = data_llam.Recordset("trasla")
               data_inflla.Recordset("lugar") = data_llam.Recordset("lugar")
               data_inflla.Recordset("totdem") = data_llam.Recordset("totdem")
               data_inflla.Recordset("codmed") = data_llam.Recordset("codmed")
               data_inflla.Recordset("nommed") = data_llam.Recordset("nommed")
               data_inflla.Recordset("totend") = data_llam.Recordset("totend")
               data_inflla.Recordset("hsald") = data_llam.Recordset("hsald")
               data_inflla.Recordset("hzona") = data_llam.Recordset("hzona")
               data_inflla.Recordset("realiza") = data_llam.Recordset("realiza")
               data_inflla.Recordset("mes") = data_llam.Recordset("mes")
               data_inflla.Recordset("ano") = data_llam.Recordset("ano")
               data_inflla.Recordset("hh") = data_llam.Recordset("hh")
               data_inflla.Recordset("colormot") = data_llam.Recordset("colormot")
               data_inflla.Recordset("horsali") = data_llam.Recordset("horsali")
               data_inflla.Recordset("diag") = data_llam.Recordset("diag")
               data_inflla.Recordset("ci") = data_llam.Recordset("ci")
               data_inflla.Recordset.Update
               data_llam.Recordset.MoveNext
               pbar.Value = pbar.Value + 1
            Loop
         Else
            Do While Not data_llam.Recordset.EOF
                data_inflla.Recordset.AddNew
                data_inflla.Recordset("fecha") = data_llam.Recordset("fecha")
                data_inflla.Recordset("hora") = data_llam.Recordset("hora")
                data_inflla.Recordset("matric") = data_llam.Recordset("matric")
                data_inflla.Recordset("nombre") = data_llam.Recordset("nombre")
                data_inflla.Recordset("categ") = data_llam.Recordset("categ")
                data_inflla.Recordset("nomcat") = data_llam.Recordset("nomcat")
                data_inflla.Recordset("motcon") = data_llam.Recordset("motcon")
                data_inflla.Recordset("movilpas") = data_llam.Recordset("movilpas")
                data_inflla.Recordset("codmot") = data_llam.Recordset("codmot")
                data_inflla.Recordset("timdes") = data_llam.Recordset("timdes")
                data_inflla.Recordset("activo") = data_llam.Recordset("activo")
                data_inflla.Recordset("usuario") = data_llam.Recordset("usuario")
                data_inflla.Recordset("diag") = data_llam.Recordset("diag")
                data_inflla.Recordset("ci") = data_llam.Recordset("ci")
                data_inflla.Recordset("motmov") = data_llam.Recordset("motmov")
                data_inflla.Recordset("lugar") = data_llam.Recordset("lugar")
                data_inflla.Recordset("totdem") = data_llam.Recordset("totdem")
                data_inflla.Recordset("codmed") = data_llam.Recordset("codmed")
                data_inflla.Recordset("nommed") = data_llam.Recordset("nommed")
                data_inflla.Recordset("codzon") = data_llam.Recordset("codzon")
                data_inflla.Recordset("enfer") = data_llam.Recordset("enfer")
                If IsNull(data_llam.Recordset("cancela")) = True Then
                   data_inflla.Recordset("cancela") = 0
                Else
                   data_inflla.Recordset("cancela") = data_llam.Recordset("cancela")
                End If
                data_inflla.Recordset("fec_cance") = data_llam.Recordset("fec_cance")
                data_inflla.Recordset("hor_cance") = data_llam.Recordset("hor_cance")
                data_inflla.Recordset("motcance") = data_llam.Recordset("motcance")
                If IsNull(data_llam.Recordset("activo")) = False Then
                   Xmin = Val(Mid(data_llam.Recordset("activo"), 4, 2))
                Else
                   Xmin = 0
                End If
                If IsNull(data_llam.Recordset("hora")) = False Then
                   Xmind = Val(Mid(data_llam.Recordset("hora"), 4, 2))
                Else
                   Xmind = 0
                End If
                Xdif = Xmin - Xmind
                If Xdif < 0 Then
                   Xdif = Xdif + 60
                End If
                data_inflla.Recordset("timsi") = Xdif
                data_inflla.Recordset("colormot") = data_llam.Recordset("colormot")
                data_inflla.Recordset("horsali") = data_llam.Recordset("horsali")
                data_inflla.Recordset.Update
                data_llam.Recordset.MoveNext
                pbar.Value = pbar.Value + 1
            Loop
         End If
         If Option10.Value = False Then
            data_inflla.RecordSource = "select * from inflla"
            data_inflla.Refresh
            If data_inflla.Recordset.RecordCount > 0 Then
               data_inflla.Recordset.MoveFirst
               Do While Not data_inflla.Recordset.EOF
                  If IsNull(data_inflla.Recordset("cancela")) = True Then
                  Else
                     If data_inflla.Recordset("cancela") = 0 Then
                     Else
                        If data_inflla.Recordset("cancela") = 1 Then
                           data_inflla.Recordset.Delete
                        End If
                     End If
                  End If
                  data_inflla.Recordset.MoveNext
               Loop
               data_inflla.Refresh
            End If
         Else
            data_inflla.RecordSource = "select * from inflla"
            data_inflla.Refresh
            If data_inflla.Recordset.RecordCount > 0 Then
               data_inflla.Recordset.MoveFirst
               Do While Not data_inflla.Recordset.EOF
                  If IsNull(data_inflla.Recordset("cancela")) = False Then
                     If data_inflla.Recordset("cancela") = 1 Then
                     Else
                        data_inflla.Recordset.Delete
                     End If
                  Else
                     data_inflla.Recordset.Delete
                  End If
                  data_inflla.Recordset.MoveNext
               Loop
               data_inflla.Refresh
            End If
         End If
         If Option8.Value = True Then
            Dim Xlama As Long
            Dim Xcuentomat As Integer
            Xcuentomat = 0
            Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

            If Xcantveces <> "" Then
               data_inflla.RecordSource = "select * from inflla order by matric"
               data_inflla.Refresh
               If data_inflla.Recordset.RecordCount > 0 Then
                  data_inflla.Recordset.MoveFirst
                  Xlama = data_inflla.Recordset("matric")
                  Do While Not data_inflla.Recordset.EOF
                     If data_inflla.Recordset("matric") = Xlama Then
                        Xcuentomat = Xcuentomat + 1
                     Else
                        If Xcuentomat >= Xcantveces Then
                           If data_inflla.Recordset("matric") = 0 Or data_inflla.Recordset("matric") = 999999999 Then
                           Else
                              data_inflla.Recordset.MovePrevious
                              data_inf2.Recordset.AddNew
                              data_inf2.Recordset("cod_cli") = data_inflla.Recordset("matric")
                              data_inf2.Recordset("nom_cli") = Mid(data_inflla.Recordset("nombre"), 1, 30)
                              data_inf2.Recordset("factura") = data_inflla.Recordset("edad")
                              data_inf2.Recordset("convenio") = data_inflla.Recordset("categ")
                              data_inf2.Recordset("nom_prod") = data_inflla.Recordset("nomcat")
                              
                              data_inf2.Recordset.Update
                              
                              data_inflla.Recordset.MoveNext
                           End If
                            
                           Xcuentomat = 1
                        Else
                           Xcuentomat = 1
                           MiBaseact.Execute "Update inflla set cancela = 99 where matric =" & Xlama
                           
                           'data_inflla.Recordset.Edit
                           'data_inflla.Recordset("cancela") = 99
                           'data_inflla.Recordset.Update
                        End If
                    End If
                    Xlama = data_inflla.Recordset("matric")
                    data_inflla.Recordset.MoveNext
                  Loop
                  If data_inflla.Recordset.RecordCount > 0 Then
                     data_inflla.Recordset.MoveFirst
                     Do While Not data_inflla.Recordset.EOF
                        If data_inflla.Recordset("cancela") = 99 Then
                           data_inflla.Recordset.Delete
                        Else
                           If data_inflla.Recordset("matric") = 0 Or data_inflla.Recordset("matric") = 999999999 Then
                              data_inflla.Recordset.Delete
                           Else
                              If IsNull(data_inflla.Recordset("enfer")) = False Then
                                 If data_inflla.Recordset("enfer") = 1 Then
                                    data_inflla.Recordset.Delete
                                 End If
                              End If
                           End If
                        End If
                        data_inflla.Recordset.MoveNext
                     Loop
                  End If
                  data_inflla.Refresh
                  
               End If
            End If
         End If
         If Option5.Value = True And Val(queop) = 999 Then
            Dim Xlama2 As Long
            Dim Xcuentomat2 As Integer
            Xcuentomat2 = 0
            Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")
            If Xcantveces <> "" Then
               data_inflla.RecordSource = "select * from inflla order by matric"
               data_inflla.Refresh
               If data_inflla.Recordset.RecordCount > 0 Then
                  data_inflla.Recordset.MoveFirst
                  Xlama2 = data_inflla.Recordset("matric")
                  Do While Not data_inflla.Recordset.EOF
                     If data_inflla.Recordset("matric") = Xlama Then
                        Xcuentomat2 = Xcuentomat2 + 1
                     Else
                        If Xcuentomat2 >= Xcantveces Then
                           Xcuentomat2 = 1
                        Else
                           Xcuentomat2 = 1
                           MiBaseact.Execute "Update inflla set cancela = 99 where matric =" & Xlama2
                           
                           'data_inflla.Recordset.Edit
                           'data_inflla.Recordset("cancela") = 99
                           'data_inflla.Recordset.Update
                        End If
                    End If
                    Xlama2 = data_inflla.Recordset("matric")
                    data_inflla.Recordset.MoveNext
                  Loop
                  If data_inflla.Recordset.RecordCount > 0 Then
                     data_inflla.Recordset.MoveFirst
                     Do While Not data_inflla.Recordset.EOF
                        If data_inflla.Recordset("cancela") = 99 Then
                           data_inflla.Recordset.Delete
                        Else
                           If data_inflla.Recordset("matric") = 0 Or data_inflla.Recordset("matric") = 999999999 Then
                              data_inflla.Recordset.Delete
                           Else
                              If IsNull(data_inflla.Recordset("enfer")) = False Then
                                 If data_inflla.Recordset("enfer") = 1 Then
                                    data_inflla.Recordset.Delete
                                 End If
                              End If
                           End If
                        End If
                        data_inflla.Recordset.MoveNext
                     Loop
                  End If
                  data_inflla.Refresh
               End If
            End If
         End If
         data_inflla.RecordSource = "select * from inflla order by fecha"
         data_inflla.Refresh
         If Option1.Value = True Then
            If Check1.Value = 1 Then
               CrystalReport1.ReportFileName = App.path & "\infdzn.rpt"
               CrystalReport1.ReportTitle = "INFORME LLAMADOS POR ZONA DESDE " & md.Text & " HASTA " & mh.Text
               CrystalReport1.Action = 1
            Else
               CrystalReport1.ReportFileName = App.path & "\infdz.rpt"
               CrystalReport1.ReportTitle = "INFORME LLAMADOS POR ZONA DESDE " & md.Text & " HASTA " & mh.Text
               CrystalReport1.Action = 1
            End If
         End If
         If Option2.Value = True Then
            If Check1.Value = 1 Then
               CrystalReport1.ReportFileName = App.path & "\infdapn.rpt"
               CrystalReport1.ReportTitle = "INFORME LLAMADOS POR EDAD DESDE " & md.Text & " HASTA " & mh.Text
               CrystalReport1.Action = 1
            Else
               CrystalReport2.ReportFileName = App.path & "\infdap.rpt"
               CrystalReport2.ReportTitle = "INFORME LLAMADOS POR EDAD DESDE " & md.Text & " HASTA " & mh.Text
               CrystalReport2.Action = 1
        '            CrystalReport2.WindowState = crptMaximized
            End If
         End If
         If Option3.Value = True Then
            If Check1.Value = 1 Then
               CrystalReport1.ReportFileName = App.path & "\infdmedn.rpt"
               CrystalReport1.ReportTitle = "INFORME LLAMADOS POR MEDICO DESDE " & md.Text & " HASTA " & mh.Text
               CrystalReport1.Action = 1
            Else
               CrystalReport2.ReportFileName = App.path & "\infdmedd.rpt"
               CrystalReport2.ReportTitle = "INFORME LLAMADOS POR MEDICO DE LLAMADO DESDE " & md.Text & " HASTA " & mh.Text
               CrystalReport2.Action = 1
            End If
         End If
         If Option4.Value = True Then
            If Check1.Value = 1 Then
               CrystalReport1.ReportFileName = App.path & "\infdmovn.rpt"
               CrystalReport1.ReportTitle = "INFORME LLAMADOS POR MOVIL DESDE " & md.Text & " HASTA " & mh.Text
               CrystalReport1.Action = 1
            Else
               CrystalReport2.ReportFileName = App.path & "\infdmovd.rpt"
               CrystalReport2.ReportTitle = "INFORME LLAMADOS POR MOVIL DE LLAMADO DESDE " & md.Text & " HASTA " & mh.Text
               CrystalReport2.Action = 1
            End If
         End If
         If Option5.Value = True Then
            If Val(queop) = 999 Then
               CrystalReport2.ReportFileName = App.path & "\infdcliz.rpt"
               CrystalReport2.ReportTitle = "INFORME DE INTERNACION POR SOCIO DESDE " & md.Text & " HASTA " & mh.Text
               CrystalReport2.Action = 1
            Else
               If Check1.Value = 1 Then
                  CrystalReport1.ReportFileName = App.path & "\infdtran.rpt"
                  CrystalReport1.ReportTitle = "INFORME de TRASLADOS DESDE " & md.Text & " HASTA " & mh.Text
                  CrystalReport1.Action = 1
               Else
                  CrystalReport2.ReportFileName = App.path & "\infdtrad.rpt"
                  CrystalReport2.ReportTitle = "INFORME DE TRASLADOS DESDE " & md.Text & " HASTA " & mh.Text
                  CrystalReport2.Action = 1
    '            CrystalReport2.WindowState = crptMaximized
               End If
            End If
         End If
         If Option6.Value = True Then
             If Check1.Value = 1 Then
                CrystalReport1.ReportFileName = App.path & "\infdcodn.rpt"
                CrystalReport1.ReportTitle = "INFORME POR CODIGO DE LLAMADO DESDE " & md.Text & " HASTA " & mh.Text
                CrystalReport1.Action = 1
             Else
                CrystalReport2.ReportFileName = App.path & "\infdcodd.rpt"
                CrystalReport2.ReportTitle = "INFORME POR CODIGO DE LLAMADO DESDE " & md.Text & " HASTA " & mh.Text
                CrystalReport2.Action = 1
    '            CrystalReport2.WindowState = crptMaximized
             End If
         End If
         If Option7.Value = True Then
             If Check1.Value = 1 Then
                CrystalReport1.ReportFileName = App.path & "\infdconvn.rpt"
                CrystalReport1.ReportTitle = "INFORME POR CONVENIOS DESDE " & md.Text & " HASTA " & mh.Text
                CrystalReport1.Action = 1
             Else
                CrystalReport2.ReportFileName = App.path & "\infdconvd.rpt"
                CrystalReport2.ReportTitle = "INFORME POR CONVENIOS DESDE " & md.Text & " HASTA " & mh.Text
                CrystalReport2.Action = 1
    '            CrystalReport2.WindowState = crptMaximized
             End If
         End If
         If Option8.Value = True Then
             If Check1.Value = 1 Then
                CrystalReport1.ReportFileName = App.path & "\infdcliz.rpt"
                CrystalReport1.ReportTitle = "INFORME DE CONSULTAS POR SOCIO DESDE " & md.Text & " HASTA " & mh.Text
                CrystalReport1.Action = 1
             Else
                CrystalReport2.ReportFileName = App.path & "\infdcliz.rpt"
                CrystalReport2.ReportTitle = "INFORME DE CONSULTAS POR SOCIO DESDE " & md.Text & " HASTA " & mh.Text
                CrystalReport2.Action = 1
    '            CrystalReport2.WindowState = crptMaximized
             End If
         End If
         If Option9.Value = True Then
             If Check1.Value = 1 Then
                CrystalReport1.ReportFileName = App.path & "\infdcosn.rpt"
                CrystalReport1.ReportTitle = "INFORME DE CONSULTAS CON COSTO DESDE " & md.Text & " HASTA " & mh.Text
                CrystalReport1.Action = 1
             Else
                CrystalReport2.ReportFileName = App.path & "\infdcosd.rpt"
                CrystalReport2.ReportTitle = "INFORME DE CONSULTAS CON COSTO DESDE " & md.Text & " HASTA " & mh.Text
                CrystalReport2.Action = 1
    '            CrystalReport2.WindowState = crptMaximized
             End If
         End If
         If Option10.Value = True Then
             If Check1.Value = 1 Then
                CrystalReport1.ReportFileName = App.path & "\infdcance.rpt"
                CrystalReport1.ReportTitle = "INFORME DE LLAMADOS CANCELADOS DESDE " & md.Text & " HASTA " & mh.Text
                CrystalReport1.Action = 1
             Else
                CrystalReport2.ReportFileName = App.path & "\infdcance.rpt"
                CrystalReport2.ReportTitle = "INFORME DE LLAMADOS CANCELADOS DESDE " & md.Text & " HASTA " & mh.Text
                CrystalReport2.Action = 1
    '            CrystalReport2.WindowState = crptMaximized
             End If
         End If
         If Option11.Value = True Then
            CrystalReport1.ReportFileName = App.path & "\infddemr.rpt"
            CrystalReport1.ReportTitle = "INFORME DEMORAS EN GRABAR LLAMADO DESDE " & md.Text & " HASTA " & mh.Text
            CrystalReport1.Action = 1
         End If
         If Option12.Value = True Then
            If Check2.Value = 1 Then
               CrystalReport1.ReportFileName = App.path & "\infdusul.rpt"
               CrystalReport1.ReportTitle = "INFORME DE LLAMADOS POR LARGADOR DESDE " & md.Text & " HASTA " & mh.Text
               CrystalReport1.Action = 1
            End If
            If Check3.Value = 1 Then
               CrystalReport1.ReportFileName = App.path & "\infdusur.rpt"
               CrystalReport1.ReportTitle = "INFORME DE LLAMADOS POR RECEPTOR DESDE " & md.Text & " HASTA " & mh.Text
               CrystalReport1.Action = 1
            End If
            
         End If
      
      Else
         MsgBox "No existen registros", vbInformation, "Mensaje"
      End If
   Else
      MsgBox "No ingresó fecha", vbInformation, "Mensaje"
   End If
Else
   MsgBox "No ingresó fecha", vbInformation, "Mensaje"
End If
pbar.Value = 0
'pbar.Max = 0
frm_infdesp.MousePointer = 0

End Sub

Private Sub bcan_Click()
Unload Me

End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
   Check3.Value = 0
End If

End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
   Check2.Value = 0
End If

End Sub

Private Sub Command1_Click()
If Combo1.Text = "" Then
   Combo1.ListIndex = 0
End If

If Combo1.Text <> "TODOS" Then
'   data_conv.Recordset.FindFirst "cnv_desc ='" & Combo1.Text & "'"
   data_conv.RecordSource = "Select * from convenio where cnv_desc ='" & Combo1.Text & "'"
   data_conv.Refresh
   If data_conv.Recordset.RecordCount > 0 Then
      Text1.Text = data_conv.Recordset("cnv_codigo")
   Else
'      data_conv.Recordset.FindFirst "cnv_codigo ='" & Combo1.Text & "'"
      data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & Combo1.Text & "'"
      data_conv.Refresh
      If data_conv.Recordset.RecordCount > 0 Then
         Text1.Text = data_conv.Recordset("cnv_codigo")
      Else
         Text1.Text = "TODOS"
      End If
   End If
Else
   Text1.Text = "TODOS"
End If
Frame2.Visible = False
bace.SetFocus
End Sub

Private Sub Command2_Click()
Frame2.Visible = False
Text1.Text = ""

End Sub

Private Sub Form_Load()
Dim Xlafella As Date
'    data_inf2.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"
    data_inf2.DatabaseName = App.path & "\informes.mdb"
'    data_inf2.RecordSource = "infvtas"
'    data_inf2.Refresh

Xlafella = Date - 1
'    data_inflla.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"
    data_inflla.DatabaseName = App.path & "\informes.mdb"
'    data_inflla.RecordSource = "inflla"
'    data_inflla.Refresh
    data_llam.ConnectionString = "dsn=" & Xconexrmt
'    data_llam.RecordSource = "Select * from llamado where fecha >=#" & Format(Xlafella, "yyyy/mm/dd") & "#"
'    data_llam.Refresh
    data_conv.ConnectionString = "dsn=" & Xconexrmt
    data_conv.RecordSource = "select * from convenio order by cnv_desc"
    data_conv.Refresh
Combo1.AddItem "TODOS"

If data_conv.Recordset.RecordCount > 0 Then
   data_conv.Recordset.MoveFirst
   Do While Not data_conv.Recordset.EOF
      If IsNull(data_conv.Recordset("cnv_desc")) = False Then
         Combo1.AddItem data_conv.Recordset("cnv_Desc")
      End If
      data_conv.Recordset.MoveNext
   Loop
End If

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
   Option1.SetFocus
End If

End Sub

Private Sub Option12_Click()
If Option12.Value = True Then
   Check2.Visible = True
   Check3.Visible = True
Else
   Check2.Visible = False
   Check3.Visible = False
End If

End Sub
