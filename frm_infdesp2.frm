VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infdesp2 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes despacho"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6795
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   177
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_infdesp2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   6795
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command7 
      Caption         =   "CP SJ"
      Height          =   375
      Left            =   2760
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2880
      Visible         =   0   'False
      Width           =   2700
   End
   Begin MSAdodcLib.Adodc data_lla 
      Height          =   330
      Left            =   2880
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
      Caption         =   "data_lla"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc data_conv 
      Height          =   330
      Left            =   480
      Top             =   5040
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command4 
      Caption         =   "CP"
      Height          =   375
      Left            =   4560
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   5280
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   2040
      TabIndex        =   16
      Top             =   5400
      Visible         =   0   'False
      Width           =   855
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   6120
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_movil 
      Caption         =   "data_movil"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5640
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5760
      Picture         =   "frm_infdesp2.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Salir"
      Top             =   5760
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Picture         =   "frm_infdesp2.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Procesar"
      Top             =   5760
      Width           =   615
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Tipo de informe"
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   2880
      Width           =   6495
      Begin VB.CheckBox Check9 
         BackColor       =   &H00C00000&
         Caption         =   "Agregar móvil"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3600
         TabIndex        =   31
         Top             =   1920
         Width           =   2415
      End
      Begin VB.CheckBox Check8 
         BackColor       =   &H00C00000&
         Caption         =   "Control de socios"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   30
         Top             =   1920
         Width           =   2415
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00C00000&
         Caption         =   "Tiempos Traslados"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   1560
         Width           =   2415
      End
      Begin VB.CheckBox Check6 
         BackColor       =   &H00C00000&
         Caption         =   "Incluir actos de enf."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3600
         TabIndex        =   25
         Top             =   1200
         Width           =   2415
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00C00000&
         Caption         =   "Solo los CELESTES"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1200
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C00000&
         Caption         =   "Sin llamados de base"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3600
         TabIndex        =   23
         Top             =   840
         Width           =   2415
      End
      Begin VB.Data data_inf 
         Caption         =   "data_inf"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   480
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   -120
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSAdodcLib.Adodc data_moviles 
         Height          =   330
         Left            =   1320
         Top             =   0
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
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
         Caption         =   "data_moviles"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc data_lla22 
         Height          =   330
         Left            =   4080
         Top             =   120
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
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
         Caption         =   "data_lla22"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.Data data_llaenf 
         Caption         =   "data_llaenf"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   5040
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   960
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H0000FFFF&
         Caption         =   "Desde respaldos"
         Height          =   255
         Left            =   3600
         TabIndex        =   20
         Top             =   1560
         Width           =   2415
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Solo demoras >2hs traslados"
         Height          =   255
         Left            =   2760
         TabIndex        =   19
         Top             =   600
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C00000&
         Caption         =   "Generar planilla"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   840
         Width           =   2415
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Resumen"
         Height          =   255
         Left            =   3600
         TabIndex        =   10
         Top             =   360
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Detalle"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
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
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6495
      Begin VB.CommandButton Command8 
         Caption         =   "cmt medic"
         Height          =   375
         Left            =   240
         TabIndex        =   29
         Top             =   2040
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Data data_buscacov 
         Caption         =   "data_buscacov"
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
         Top             =   360
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Data data_covid 
         Caption         =   "data_covid"
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
         Top             =   600
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton Command6 
         Caption         =   "COVID"
         Height          =   375
         Left            =   1440
         TabIndex        =   27
         Top             =   720
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CommandButton Command5 
         Caption         =   "lo viejo"
         Height          =   615
         Left            =   4560
         TabIndex        =   22
         Top             =   600
         Visible         =   0   'False
         Width           =   1575
      End
      Begin MSAdodcLib.Adodc data_med 
         Height          =   375
         Left            =   480
         Top             =   1320
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
         Caption         =   "data_med"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.ComboBox Combo3 
         Height          =   360
         ItemData        =   "frm_infdesp2.frx":0F56
         Left            =   2640
         List            =   "frm_infdesp2.frx":0F8A
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   2280
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   360
         Left            =   1440
         TabIndex        =   14
         Top             =   2280
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1080
         TabIndex        =   13
         Top             =   2040
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   360
         ItemData        =   "frm_infdesp2.frx":1026
         Left            =   2640
         List            =   "frm_infdesp2.frx":1069
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1680
         Width           =   3375
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "frm_infdesp2.frx":1196
         Left            =   2640
         List            =   "frm_infdesp2.frx":11A6
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   3375
      End
      Begin MSMask.MaskEdBox mfh 
         Height          =   375
         Left            =   4560
         TabIndex        =   3
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfd 
         Height          =   375
         Left            =   2640
         TabIndex        =   2
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "Agrupar por:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Informe de....:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Rango de fechas:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   1200
      Picture         =   "frm_infdesp2.frx":11E2
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   1095
   End
End
Attribute VB_Name = "frm_infdesp2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
Check1.Value = 1

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo2.SetFocus
End If

End Sub

Private Sub Combo2_Click()
If Combo2.ListIndex = 7 Then
'   If WElusuario = "JFERNAN" Or WElusuario = "SPEREZ" Or WElusuario = "MCOSTA" Or WElusuario = "SDOMINGUEZ" Then
'   Else
'      Combo2.ListIndex = 0
'      MsgBox "Usuario no autorizado para ésta opción de listado", vbInformation
'   End If
Else
    If Combo2.ListIndex = 1 Then
       Text2.Visible = True
       Combo3.Visible = False
       frm_buscondesp.Show vbModal
    Else
       If Combo2.ListIndex = 2 Or Combo2.ListIndex = 6 Then
          Text1.Text = ""
          Text2.Text = ""
          Text2.Visible = False
          Combo3.Visible = True
          Combo3.SetFocus
       Else
          Text2.Visible = True
          Combo3.Visible = False
          Text1.Text = ""
          Text2.Text = ""
       End If
    End If
End If

End Sub

Private Sub Combo3_Click()
If Combo3.ListIndex = 12 Then
   Check5.Visible = True
Else
   Check5.Visible = False
End If

End Sub

Private Sub Command1_Click()
Dim Xobjexel As Excel.Application
Dim Xlibexel As Excel.Workbook
Dim Xarchexel As New Excel.Worksheet
Dim Xnomlla, Xnomcatlla, Xobsmotlla, Xnommedlla, Xobslla, Xdiaglla, Xmotmovlla, Xtellla As String
Dim XCol, Xlin, Xnrocan, Xcolfija As Long
Dim Xdeudas As Double
Dim Xqdia, Xcanxdia As Long
Dim Xarchtex As String
Dim Xmovtr, Xcodchof, Xcodcolor, Xcodzonn As String
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
'10
Dim Xlabrir As New Excel.Application
Dim Xcodmedlla, Xcancella As Integer

Dim Xdifmin33, Xdifhor33 As Integer

Xdeudas = 0
''''On Error GoTo Queesinfdesp

If Check3.Value = 1 Then
'   data_lla.Connect = "odbc;dsn=" & Xconexrmt & ";"
'   data_lla.DatabaseName = App.Path & "\llamado.mdb"
   data_lla.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.path & "\llamado.mdb"
Else
'   data_lla.Connect = "odbc;dsn=" & Xconexrmt & ";"
   data_lla.ConnectionString = "dsn=" & Xconexrmt
   data_lla22.ConnectionString = "dsn=" & Xconexrmt
'   data_lla22.Connect = "odbc;dsn=" & Xconexrmt & ";"
End If
XCol = 1
Xlin = 1
Xnrocan = 1
Command1.Enabled = False
Command2.Enabled = False
frm_infdesp2.MousePointer = 11

Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from inflla"

Set MiBaseact = Unasesact.OpenDatabase(App.path & "\infllavpn.mdb")

MiBaseact.Execute "Delete * from inflla"
Data1.DatabaseName = App.path & "\infllavpn.mdb"
Data1.RecordSource = "inflla"
Data1.Refresh
Dim Xtotnumero As Long
Dim XsionoSamc As String
Dim Xmsp1, Xmsp2, Xmsp3 As Integer
Dim Xmsp1f, Xmsp2f, Xmsp3f As Integer
Dim X9111, X9112, X9113 As Integer
Dim X9111f, X9112f, X9113f As Integer
Dim Xterce1, Xterce2, Xterce3 As Integer
Dim Xterce1f, Xterce2f, Xterce3f As Integer
Dim Xtr1, Xtr2, Xtr3 As Integer
Dim Xtr1f, Xtr2f, Xtr3f As Integer

Xtotnumero = 0

'If data_inf.Recordset.RecordCount > 0 Then
'   data_inf.Recordset.MoveFirst
'   Do While Not data_inf.Recordset.EOF
'      data_inf.Recordset.Delete
'      data_inf.Recordset.MoveNext
'   Loop
'End If
pb.Max = 1
pb.Value = 0

If mfd.Text = "__/__/____" Then
Else
'If Combo2.ListIndex = 3 Or Combo3.ListIndex = 11 Or Check1.Value = 1 Then
   If Combo2.ListIndex = 5 Then
      Xmovtr = InputBox("Ingrese móvil de TRASLADO (0=TODOS)", "MOVIL DE TRASLADO")
      If Xmovtr = 215 Or Xmovtr = 315 Or Xmovtr = 415 Then
         XsionoSamc = MsgBox("Desea filtrar datos de Sauce?", vbInformation + vbYesNo, "Informes")
         If XsionoSamc = vbYes Then
         
         End If
      Else
         XsionoSamc = vbNo
      End If
   Else
      If Combo2.ListIndex = 6 Then
         Xmovtr = InputBox("Ingrese móvil de LLAMADO (0=TODOS)", "MOVIL DE LLAMADO")
         If Xmovtr = 215 Or Xmovtr = 315 Or Xmovtr = 415 Then
            XsionoSamc = MsgBox("Desea filtrar datos de Sauce?", vbInformation + vbYesNo, "Informes")
            If XsionoSamc = vbYes Then
           
            End If
         Else
            XsionoSamc = vbNo
         End If
      
      End If
   End If
   If Combo2.ListIndex = 11 Then
      Xmovtr = InputBox("Ingrese CODIGO del MEDICO (0=TODOS)", "MEDICO DEL LLAMADO")
   End If
   
   If Combo2.ListIndex = 12 Then
      Xcodcolor = InputBox("Ingrese CODIGO (T=TODOS)", "CLAVE DE LLAMADOS")
   Else
      Xcodcolor = "T"
   End If
   
   If Combo2.ListIndex = 13 Then
      Xcodzonn = InputBox("Ingrese CODIGO DE ZONA (0=TODAS)", "ZONAS DE LLAMADOS")
   Else
      Xcodzonn = ""
   End If
   
   If Xmovtr = "" Then
      Xmovtr = "0"
   End If
   If Check1.Value = 1 Then

  'Set xlsfileout = Xlibro.Worksheets.Add
  'xlsfileout.Name = "MIHOJADOS"
      If Combo1.ListIndex = 2 And Combo2.ListIndex = 3 Then
      Else
         If Combo2.ListIndex = 7 Or Combo2.ListIndex = 18 Or Combo2.ListIndex = 19 Then
         Else
            Set Xobjexel = New Excel.Application
    
            Set Xlibexel = Xobjexel.Workbooks.Add
            Set Xarchexel = Xlibexel.Worksheets.Add
    
            Xarchexel.Name = Trim(Combo1.Text)
            
            Xlibexel.SaveAs ("C:\planillas\" & Trim(Combo1.Text) & "\" & Trim(str(Month(mfd.Text))) & Trim(str(Year(mfd.Text))) & ".xls")
            Xarchtex = "C:\planillas\" & Trim(Combo1.Text) & "\" & Trim(str(Month(mfd.Text))) & Trim(str(Year(mfd.Text))) & ".xls"
         End If
      End If
'      Set Xarchexel = Xlibexel.Worksheets.Add
'      Xarchexel.Name = "HOJAUNA"
            
'      Set Xarchexel = Xobjexel.Sheets("HOJAUNA")
      
   End If
   If mfh.Text = "__/__/____" Then
   Else
      If Combo1.ListIndex = 0 Then
         If Combo2.ListIndex = 0 Then
            data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and categ in" & _
            " ('UDEMM','CERSEM','CERADU','CERDGI','CERSEM','CERHEV','CERCAS','CERMAT','CERKEV','CERIMP','CERSEV','CERVIS') and base =" & 0 & " and cancela is null order by fecha"
            data_lla.Refresh
         Else
            If Combo2.ListIndex = 14 Then
               data_lla.RecordSource = "select llamado.fecha,llamado.categ,llamado.nrolla,llamado.codmot,llamado.nombre,llamado.nomcat," & _
               "llamado.unied,llamado.edad,llamado.matric,llamado.motmov,llamado.lugar,llamado.movtras,llamado.nommed,llamado.hor_llega,llamado.obsmot," & _
               "llamado.motcon,llamado.dcobr,llamado.diag,llamado.telef,llamado.hsald,llamado.hllega,llamado.hor_cance,llamado.codzon,llamado.obs," & _
               "llamado.hzona,llamado.timdes,llamado.descol,llamado.movilpas,llamado.hor_rea,llamado.ncobr,llamado.mes,llamado.hora,llamado.ci," & _
               "llamado.ano,llamado.totend,llamado.referen,llamado.colormot,llamado.base,resplla.trasla,resplla.fecpas,resplla.mes from llamado " & _
               "inner join resplla on llamado.nrolla=resplla.nro where resplla.trasla is not null and " & _
               "llamado.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and llamado.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and llamado.base >=" & 0 & " and llamado.movilpas <>" & 99 & " and llamado.categ in " & _
               "('UDEMM','CERSEM','CERADU','CERDGI','CERSEM','CERHEV','CERCAS','CERMAT','CERKEV','CERIMP','CERSEV','CERVIS') and llamado.cancela is null order by llamado.fecha"
               data_lla.Refresh
            Else
               If Combo2.ListIndex = 16 Or Combo2.ListIndex = 17 Then
                  data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and movilpas =" & 597
                  data_lla.Refresh
               Else
                  data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and categ ='" & Text1.Text & "' and base =" & 0 & " and cancela is null order by fecha"
                  data_lla.Refresh
               End If
            End If
         End If
      Else
'
         If Combo1.ListIndex = 1 Then
            If Combo2.ListIndex = 0 Or Combo2.ListIndex = 2 Then
               data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and enfer =" & 1 & " and cancela is null order by fecha"
               data_lla.Refresh
            Else
               If Combo2.ListIndex = 14 Then
                  data_lla.RecordSource = "select llamado.fecha,llamado.categ,llamado.nrolla,llamado.codmot,llamado.nombre,llamado.nomcat," & _
                  "llamado.unied,llamado.edad,llamado.matric,llamado.motmov,llamado.lugar,llamado.movtras,llamado.nommed,llamado.codzon,llamado.obsmot," & _
                  "llamado.motcon,llamado.dcobr,llamado.diag,llamado.telef,llamado.hsald,llamado.hllega,llamado.hor_cance,llamado.hor_llega,llamado.obs," & _
                  "llamado.hzona,llamado.timdes,llamado.descol,llamado.movilpas,llamado.hor_rea,llamado.ncobr,llamado.mes,llamado.hora,llamado.ci," & _
                  "llamado.ano,llamado.totend,llamado.referen,llamado.colormot,llamado.base,resplla.trasla,resplla.fecpas,resplla.mes from llamado " & _
                  "inner join resplla on llamado.nrolla=resplla.nro where resplla.trasla is not null and " & _
                  "llamado.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and llamado.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and llamado.base >=" & 0 & " and llamado.cancela is null and llamado.enfer =" & 1 & " order by llamado.fecha"
                  data_lla.Refresh
               Else
                  If Combo2.ListIndex = 16 Or Combo2.ListIndex = 17 Then
                     data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and movilpas =" & 597
                     data_lla.Refresh
                  Else
                     data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and enfer =" & 1 & " and categ ='" & Text1.Text & "' and cancela is null order by fecha"
                     data_lla.Refresh
                  End If
               End If
            End If
         Else
            If Combo1.ListIndex = 2 Then 'Traslados
               If Combo2.ListIndex = 0 Or Combo2.ListIndex = 2 Or Combo2.ListIndex = 3 Or Combo2.ListIndex = 4 Or Combo2.ListIndex = 9 Then
                  If Combo3.ListIndex = 11 Or Combo3.ListIndex = 8 Then
                     data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16) and cancela is null and categ in ('911','911B') order by fecha"
                     data_lla.Refresh
                  Else
                     If Combo3.ListIndex = 12 Or Combo3.ListIndex = 13 Then
                        If Combo3.ListIndex = 13 Then
                           data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16) and codmed =" & 959 & " and enfer not in (1) and cancela is null order by fecha"
                           data_lla.Refresh
                        Else
                           data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16) and cancela is null and codmot ='" & "C" & "' order by fecha"
                           data_lla.Refresh
                        End If
                     Else
                        If Combo3.ListIndex = 7 Then
                           data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16) and cancela is null and categ in ('MSP') order by fecha"
                           data_lla.Refresh
                        Else
                           If Combo2.ListIndex = 9 Then
                              data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and base >" & 0 & " and cancela is null order by fecha"
                              data_lla.Refresh
                           Else
                              If Check4.Value = 1 Then
                                 data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and base =" & 0 & " and cancela is null order by fecha"
                                 data_lla.Refresh
                              Else
                                 data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and base >=" & 0 & " and cancela is null order by fecha"
                                 data_lla.Refresh
                              End If
                           End If
                        End If
                     End If
                  End If
               Else
                  If Combo2.ListIndex = 5 Or Combo2.ListIndex = 6 Or Combo2.ListIndex = 11 Then
                     If Xmovtr = "0" Then
                        data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and base >=" & 0 & " and cancela is null order by fecha"
                        data_lla.Refresh
                     Else
                        If Combo2.ListIndex = 5 Then
                           If XsionoSamc = vbYes Then
                              data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and base >=" & 0 & " and movtras =" & Val(Xmovtr) & " and cancela is null and categ not in ('SAMCB') order by fecha"
                              data_lla.Refresh
                           Else
                              data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and base >=" & 0 & " and movtras =" & Val(Xmovtr) & " and cancela is null order by fecha"
                              data_lla.Refresh
                           End If
                        Else
                           If Combo2.ListIndex = 11 Then
                              data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and base >=" & 0 & " and codmed =" & Val(Xmovtr) & " and cancela is null order by fecha"
                              data_lla.Refresh
                           Else
                              If Combo2.ListIndex = 1 Then
                                 data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and base >=" & 0 & " and movilpas =" & Val(Xmovtr) & " and categ ='" & Text1.Text & "' and cancela is null order by fecha"
                                 data_lla.Refresh
                              Else
                                 If XsionoSamc = vbYes Then 'and categ not in ('SAMCB')
                                    data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and base >=" & 0 & " and movilpas =" & Val(Xmovtr) & " and categ not in ('SAMCB') and cancela is null order by fecha"
                                    data_lla.Refresh
                                 Else
                                    data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and base >=" & 0 & " and movilpas =" & Val(Xmovtr) & " and cancela is null order by fecha"
                                    data_lla.Refresh
                                 End If
                              End If
                           End If
                        End If
                     End If
                  Else
                     If Combo2.ListIndex = 7 Or Combo2.ListIndex = 19 Then 'cp
                        If Combo2.ListIndex = 7 Then
                           If Format(mfd.Text, "yyyy/mm/dd") >= Format("01/08/2020", "yyyy/mm/dd") Then
                              If Format(mfd.Text, "yyyy/mm/dd") >= Format("01/07/2021", "yyyy/mm/dd") Then
                                 data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,13,14,15) and base >=" & 0 & " and categ <>'" & "SEMM" & "'" & _
                                 " and categ <>'" & "SEMM1" & "' and categ <>'" & "CCASMU" & "' and categ <>'" & "CPS" & "' and categ <>'" & "CPSSA" & "' and cancela is null order by fecha"
                              Else
                                 data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,13,14,15) and base >=" & 0 & " and categ <>'" & "SEMM" & "'" & _
                                 " and categ <>'" & "SEMM1" & "' and categ <>'" & "CCASMU" & "' and categ <>'" & "CPS" & "' and categ <>'" & "CPSSA" & "' and cancela is null and codzon not in (5) order by fecha"
                              End If
                           Else
                              data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,13,14,15) and base >=" & 0 & " and categ <>'" & "SEMM" & "'" & _
                              " and categ <>'" & "SEMM1" & "' and categ <>'" & "CCASMU" & "' and categ <>'" & "CPS" & "' and categ <>'" & "CPSSA" & "' and cancela is null order by fecha"
                           End If
                        Else
                           If Format(mfd.Text, "yyyy/mm/dd") <= Format("01/07/2020", "yyyy/mm/dd") Then
                              data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,13,14,15) and base >=" & 0 & " and categ <>'" & "SEMM" & "'" & _
                              " and categ <>'" & "SEMM1" & "' and categ <>'" & "CCASMU" & "' and categ <>'" & "CPS" & "' and categ <>'" & "CPSSA" & "' and cancela is null and movilpas in (312) order by fecha"
                           Else
                              data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,13,14,15) and base >=" & 0 & " and categ <>'" & "SEMM" & "'" & _
                              " and categ <>'" & "SEMM1" & "' and categ <>'" & "CCASMU" & "' and categ <>'" & "CPS" & "' and categ <>'" & "CPSSA" & "' and cancela is null and movilpas in (312) order by fecha"
                           End If
                        End If
                        data_lla.Refresh
                     Else
                        If Combo2.ListIndex = 1 Then 'seleccion
                           data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and categ ='" & Text1.Text & "' and cancela is null order by fecha"
                           data_lla.Refresh
                        Else
                           If Combo2.ListIndex = 12 Then
                              If Xcodcolor = "T" Or Xcodcolor = "" Then
                                 data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and base >=" & 0 & " and movilpas =" & Val(Xmovtr) & " and cancela is null order by fecha"
                                 data_lla.Refresh
                              Else
                                 data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and base >=" & 0 & " and movilpas =" & Val(Xmovtr) & " and cancela is null and codmot ='" & Xcodcolor & "' order by fecha"
                                 data_lla.Refresh
                              End If
                           Else
                              If Combo2.ListIndex = 13 Then
                                 If Xcodzonn <> "" Then
                                    If Val(Xcodzonn) = 0 Then
                                       data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and base >=" & 0 & " and cancela is null order by fecha"
                                       data_lla.Refresh
                                    Else
                                       data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and base >=" & 0 & " and cancela is null and codzon =" & Val(Xcodzonn) & " order by fecha"
                                       data_lla.Refresh
                                    End If
                                 Else
                                    data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and base >=" & 0 & " and cancela is null order by fecha"
                                    data_lla.Refresh
                                 End If
                              Else
                                 If Combo2.ListIndex = 15 Then
                                    data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and base >=" & 0 & " and cancela is null order by fecha"
                                    data_lla.Refresh
                                 Else
                                    If Combo2.ListIndex = 14 Then
                                       data_lla.RecordSource = "select llamado.fecha,llamado.categ,llamado.nrolla,llamado.codmot,llamado.nombre,llamado.nomcat," & _
                                       "llamado.unied,llamado.edad,llamado.matric,llamado.motmov,llamado.lugar,llamado.movtras,llamado.nommed,llamado.codzon,llamado.obsmot," & _
                                       "llamado.motcon,llamado.dcobr,llamado.diag,llamado.telef,llamado.hsald,llamado.hllega,llamado.hor_cance,llamado.hor_llega,llamado.obs," & _
                                       "llamado.hzona,llamado.timdes,llamado.descol,llamado.movilpas,llamado.hor_rea,llamado.ncobr,llamado.mes,llamado.hora,llamado.ci," & _
                                       "llamado.ano,llamado.totend,llamado.referen,llamado.colormot,llamado.base,resplla.trasla,resplla.fecpas,resplla.mes from llamado " & _
                                       "inner join resplla on llamado.nrolla=resplla.nro where resplla.trasla is not null and " & _
                                       "llamado.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and llamado.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and llamado.base >=" & 0 & " and llamado.cancela is null and llamado.trasla in (1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) order by llamado.fecha"
                                       data_lla.Refresh
                                    Else
                                       If Combo2.ListIndex = 16 Or Combo2.ListIndex = 17 Then
                                          data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and movilpas =" & 597
                                          data_lla.Refresh
                                       Else
                                          If Check4.Value = 1 Then 'sin llamados de base
                                             data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and base >=" & 0 & " and cancela is null order by fecha"
                                             data_lla.Refresh
                                          Else
                                             data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,7,8,9,10,11,13,14,15,16) and cancela is null order by fecha"
                                             data_lla.Refresh
                                          End If
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
'llamados
               If Combo2.ListIndex = 0 Or Combo2.ListIndex = 2 Or Combo2.ListIndex = 3 Or Combo2.ListIndex = 4 Or Combo2.ListIndex = 10 Or Combo2.ListIndex = 9 Then
                  If Combo3.ListIndex = 7 Then
                     data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cancela is null and categ ='" & "MSP" & "' and enfer not in (1) order by fecha"
                     data_lla.Refresh
                  Else
                     If Combo3.ListIndex = 11 Or Combo3.ListIndex = 8 Then
                        data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cancela is null and categ in ('911','911B') order by fecha"
                        data_lla.Refresh
                     Else
                        If Combo3.ListIndex = 12 Or Combo3.ListIndex = 13 Then
                           If Combo3.ListIndex = 13 Then
                              data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base =" & 0 & " and cancela is null and codmed =" & 959 & " and enfer not in (1) order by fecha"
                              data_lla.Refresh
                           Else
                              If Check5.Value = 1 Then
                                 data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base =" & 0 & " and cancela is null and codmot ='" & "C" & "' and codzon in (1,2,3,5) and enfer not in (1) order by fecha"
                                 data_lla.Refresh
                              Else
                                 data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base =" & 0 & " and cancela is null and codzon in (1,2,3,5) and categ not in ('UDEMM','CASH','HEVANO','CCNOS','CAUTE','SOLEME','CPS','SMI','CCSD','911','911B','SA','SAP','SAPP','EMERN','PART','SMIN','MSP','50','CAAM','CAAMEP','CCASMU','HMIL','MUCATA','MUCAMT','SAMCB','SEMM','SEMM1','SMI4','UNIDI') and enfer not in (1) order by fecha"
                                 data_lla.Refresh
                              End If
                           End If
                        Else
                           If Combo2.ListIndex = 9 Then
                              data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base >" & 0 & " and cancela is null and categ not in ('MSP','50','55') and codzon in (1,2,3,5) and enfer not in (1) order by fecha"
                              data_lla.Refresh
                           Else
                              If Check4.Value = 1 Then
                                 data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base =" & 0 & " and cancela is null and categ not in ('MSP','50','55') and codzon in (1,2,3,5) and enfer not in (1) order by fecha"
                                 data_lla.Refresh
                              Else
                                 data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base >=" & 0 & " and cancela is null and categ not in ('MSP','50','55') and codzon in (1,2,3,5) and enfer not in (1) order by fecha"
                                 data_lla.Refresh
                              End If
                           End If
                        End If
                     End If
                  End If
               Else
                  If Combo2.ListIndex = 5 Or Combo2.ListIndex = 6 Or Combo2.ListIndex = 11 Then
                     If Xmovtr = "0" Then
                        data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base =" & 0 & " and cancela is null and enfer not in (1) order by fecha"
                        data_lla.Refresh
                     Else
                        If Combo2.ListIndex = 5 Then
                           If XsionoSamc = vbYes Then 'and categ not in ('SAMCB')
                              data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,7,8,9,10,11,13) and base >=" & 0 & " and movtras =" & Val(Xmovtr) & " and enfer not in (1) and categ not in ('SAMCB') and cancela is null order by fecha"
                              data_lla.Refresh
                           Else
                              data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and trasla in (1,2,3,4,5,6,7,8,9,10,11,13) and base >=" & 0 & " and movtras =" & Val(Xmovtr) & " and enfer not in (1) and cancela is null order by fecha"
                              data_lla.Refresh
                           End If
                        Else
                           If Combo2.ListIndex = 11 Then
                              data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base >=" & 0 & " and codmed =" & Val(Xmovtr) & " and cancela is null and enfer not in (1) order by fecha"
                              data_lla.Refresh
                           Else
                              If XsionoSamc = vbYes Then 'and categ not in ('SAMCB')
                                 data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base >=" & 0 & " and movilpas =" & Val(Xmovtr) & " and enfer not in (1) and categ not in ('SAMCB') and cancela is null order by fecha"
                                 data_lla.Refresh
                              Else
                                 data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base >=" & 0 & " and movilpas =" & Val(Xmovtr) & " and enfer not in (1) and cancela is null order by fecha"
                                 data_lla.Refresh
                              End If
                           End If
                        End If
                     End If
                  Else
                     If Combo2.ListIndex = 7 Or Combo2.ListIndex = 19 Then ' aquí planilla mensual
                        If Combo2.ListIndex = 7 Then
                           If Year(mfh.Text) <= 2019 Then
                              data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base =" & 0 & " and movilpas not in (99) and enfer not in (1) order by fecha"
                           Else
                              If Format(mfd.Text, "yyyy/mm/dd") >= Format("01/08/2020", "yyyy/mm/dd") Then
                                 If Format(mfd.Text, "yyyy/mm/dd") >= Format("01/07/2021", "yyyy/mm/dd") Then
                                    If Format(mfd.Text, "yyyy/mm/dd") >= Format("01/12/2021", "yyyy/mm/dd") Then
                                       data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base =" & 0 & " and cancela is null and enfer not in (1) and movilpas not in (2015) order by fecha"
'                                      data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base =" & 0 & " and cancela is null and enfer not in (1) and codzon not in (5) order by fecha"
                                    Else
                                       data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base =" & 0 & " and cancela is null and enfer not in (1) order by fecha"
                                    End If
                                 Else
                                    data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base =" & 0 & " and cancela is null and enfer not in (1) and codzon not in (5) order by fecha"
                                 End If
                              Else
                                 data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base =" & 0 & " and cancela is null and enfer not in (1) order by fecha"
                              End If
                           End If
                        Else
                           data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base >=" & 0 & " and cancela is null and enfer not in (1) and movilpas in (312) order by fecha"
                        End If
                        data_lla.Refresh
                     Else
                        If Combo2.ListIndex = 9 Then
                           data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base =" & 0 & " and cancela is null and enfer not in (1) order by fecha"
                           data_lla.Refresh
                        Else
                           If Combo2.ListIndex = 1 Then
                           'acá modif
                              data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and categ ='" & Text1.Text & "' and cancela is null and enfer not in (1) order by fecha"
                              data_lla.Refresh
                           Else
                              If Combo2.ListIndex = 12 Then
                                 If Xcodcolor = "T" Or Xcodcolor = "" Then
                                    data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base >=" & 0 & " and cancela is null and categ not in ('MSP','50','55') and codzon in (1,2,3,5) and enfer not in (1) order by fecha"
                                    data_lla.Refresh
                                 Else
                                    data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base >=" & 0 & " and cancela is null and categ not in ('MSP','50','55') and codzon in (1,2,3,5) and codmot ='" & Xcodcolor & "' and enfer not in (1) order by fecha"
                                    data_lla.Refresh
                                 End If
                              Else
                                 If Combo2.ListIndex = 13 Then
                                    If Xcodzonn <> "" Then
                                       If Val(Xcodzonn) = 0 Then
                                          data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base >=" & 0 & " and cancela is null and categ not in ('MSP','50','55') and enfer not in (1) order by fecha"
                                          data_lla.Refresh
                                       Else
                                          data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base >=" & 0 & " and cancela is null and codzon =" & Val(Xcodzonn) & " and enfer not in (1) order by fecha"
                                          data_lla.Refresh
                                       End If
                                    Else
                                       data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base >=" & 0 & " and cancela is null and categ not in ('MSP','50','55') and enfer not in (1) order by fecha"
                                       data_lla.Refresh
                                    End If
                                 Else
                                    If Combo2.ListIndex = 15 Then
                                       data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base >=" & 0 & " and cancela in (1) and movilpas not in (0) and enfer not in (1) order by fecha"
                                       data_lla.Refresh
                                    Else
                                       If Combo2.ListIndex = 14 Then
                                          data_lla.RecordSource = "select llamado.fecha,llamado.categ,llamado.nrolla,llamado.codmot,llamado.nombre,llamado.hora,llamado.usuario,llamado.nomcat," & _
                                          "llamado.unied,llamado.edad,llamado.matric,llamado.motmov,llamado.lugar,llamado.movtras,llamado.nommed,llamado.codzon,llamado.obsmot," & _
                                          "llamado.motcon,llamado.dcobr,llamado.diag,llamado.telef,llamado.hsald,llamado.hllega,llamado.hor_cance,llamado.hor_llega,llamado.obs," & _
                                          "llamado.hzona,llamado.timdes,llamado.descol,llamado.movilpas,llamado.hor_rea,llamado.ncobr,llamado.mes,llamado.enfer,llamado.ci," & _
                                          "llamado.ano,llamado.totend,llamado.referen,llamado.colormot,llamado.base,resplla.trasla,resplla.fecpas,resplla.mes from llamado " & _
                                          "inner join resplla on llamado.nrolla=resplla.nro where resplla.trasla is not null and " & _
                                          "llamado.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and llamado.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and llamado.base >=" & 0 & " and llamado.cancela is null and llamado.categ not in ('MSP','50','55') and llamado.enfer not in (1) order by llamado.fecha"
                                          data_lla.Refresh
                                       Else
                                          If Combo2.ListIndex = 16 Then
                                             data_lla.RecordSource = "select llamado.fecha,llamado.categ,llamado.nrolla,llamado.codmot,llamado.nombre,llamado.hora,llamado.usuario,llamado.nomcat," & _
                                             "llamado.unied,llamado.edad,llamado.matric,llamado.usuario,llamado.motmov,llamado.lugar,llamado.movtras,llamado.nommed,llamado.codzon,llamado.obsmot," & _
                                             "llamado.motcon,llamado.dcobr,llamado.diag,llamado.telef,llamado.hsald,llamado.hllega,llamado.hor_cance,llamado.hor_llega,llamado.obs," & _
                                             "llamado.hzona,llamado.timdes,llamado.descol,llamado.movilpas,llamado.hor_rea,llamado.ncobr,llamado.mes,llamado.enfer,llamado.ci," & _
                                             "llamado.ano,llamado.totend,llamado.referen,llamado.colormot,llamado.base,resplla.trasla,resplla.hzona,resplla.timdes,resplla.mes from llamado " & _
                                             "inner join resplla on llamado.nrolla=resplla.nro where resplla.hzona is not null and " & _
                                             "llamado.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and llamado.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and llamado.base >=" & 0 & " and llamado.cancela is null and llamado.categ not in ('MSP','50','55') and llamado.enfer not in (1) order by llamado.fecha"
                                             data_lla.Refresh
                                          Else
                                             If Combo2.ListIndex = 17 Or Combo2.ListIndex = 20 Then
                                                data_lla.RecordSource = "select llamado.fecha,llamado.categ,llamado.nrolla,llamado.codmot,llamado.nombre,llamado.hora,llamado.usuario,llamado.nomcat," & _
                                                "llamado.unied,llamado.edad,llamado.matric,llamado.usuario,llamado.motmov,llamado.lugar,llamado.movtras,llamado.nommed,llamado.codzon,llamado.obsmot," & _
                                                "llamado.motcon,llamado.dcobr,llamado.diag,llamado.telef,llamado.hsald,llamado.hllega,llamado.hor_cance,llamado.hor_llega,llamado.obs," & _
                                                "llamado.hzona,llamado.timdes,llamado.descol,llamado.movilpas,llamado.hor_rea,llamado.ncobr,llamado.mes,llamado.enfer,llamado.ci," & _
                                                "llamado.ano,llamado.totend,llamado.referen,llamado.colormot,llamado.base,resplla.trasla,resplla.mm,resplla.timdes,resplla.mes from llamado " & _
                                                "inner join resplla on llamado.nrolla=resplla.nro where resplla.mm is not null and " & _
                                                "llamado.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and llamado.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and llamado.base >=" & 0 & " and llamado.cancela is null and llamado.categ not in ('MSP','50','55') and llamado.enfer not in (1) order by llamado.fecha"
                                                data_lla.Refresh
                                             Else
                                                If Check4.Value = 1 Then 'sin llamados de base
                                                   data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base =" & 0 & " and cancela is null and categ not in ('MSP','50','55') and enfer not in (1) order by fecha"
                                                   data_lla.Refresh
                                                Else
                                                   data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cancela is null and categ not in ('MSP','50','55') and enfer not in (1) order by fecha"
                                                   data_lla.Refresh
                                                End If
                                             End If
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                           End If
                        End If
                     End If
                  End If
               End If
            End If
         End If
      End If
      If Combo2.ListIndex = 18 Then
'         data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and segui_covid in (1) order by fecha"
         data_lla.RecordSource = "Select * from seguimiento_covid where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' order by fecha"
         data_lla.Refresh
      End If
      If data_lla.Recordset.RecordCount > 0 Then
         If Check1.Value = 1 Then
            If Combo1.ListIndex = 2 And Combo2.ListIndex = 3 Or Combo2.ListIndex = 7 Or Combo2.ListIndex = 18 Or Combo2.ListIndex = 19 Or Combo2.ListIndex = 20 Then
               If Combo2.ListIndex = 7 Or Combo2.ListIndex = 19 Or Combo2.ListIndex = 20 Then
                  If Combo2.ListIndex = 7 Then
                     Command4_Click 'CP
                  Else
                     If Combo2.ListIndex = 20 Then
'                        Command8_Click
                     Else
                        Command7_Click
                     End If
                  End If
               Else
                  If Combo2.ListIndex = 18 Then
                     Command6_Click
                  Else
                     Command3_Click
                  End If
               End If
            Else
                data_lla.Recordset.MoveLast
                pb.Max = data_lla.Recordset.RecordCount + 1
                data_lla.Recordset.MoveFirst
                Xqdia = 0
                Xcanxdia = 0
                Xarchexel.Cells(Xlin, XCol) = "DEPARTAMENTO de TI SAPP"
                XCol = 10
                Xarchexel.Cells(Xlin, XCol) = "FECHA:" & Format(Date, "dd/mm/yyyy")
                Xlin = Xlin + 1
                XCol = 2
                Xarchexel.Range("A1", "C3").Font.Size = 16
                Xarchexel.Cells(Xlin, XCol) = "PLANILLA DE " & Trim(Combo1.Text) & " DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
                XCol = 1
                Xlin = Xlin + 2
                Xnrocan = Xnrocan + Xlin
                Xarchexel.Range("A4", "AO" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
                Xarchexel.Range("A4", "AO" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
                Xarchexel.Range("A4", "AO" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
                Xarchexel.Range("A4", "AO" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
                Xarchexel.Range("A4", "AO" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
                Xarchexel.Range("A4", "AO" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
                Xarchexel.Range("A" & Trim(str(Xlin)), "AO" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
                Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 5
                Xarchexel.Cells(Xlin, XCol) = "NRO."
                XCol = XCol + 1
                Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
                Xarchexel.Cells(Xlin, XCol) = "FECHA"
                XCol = XCol + 1
                Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 35
                Xarchexel.Cells(Xlin, XCol) = "NOMBRE"
                XCol = XCol + 1
                Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 12
                Xarchexel.Cells(Xlin, XCol) = "CEDULA"
                XCol = XCol + 1
                Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 5
                Xarchexel.Cells(Xlin, XCol) = "EDAD"
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = "MAT."
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = "CAT."
                XCol = XCol + 1
                Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 35
                Xarchexel.Cells(Xlin, XCol) = "DESDE"
                Xarchexel.Range("I" & Trim(str(Xlin))).ColumnWidth = 20
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = "DESTINO"
                XCol = XCol + 1
                Xarchexel.Range("J" & Trim(str(Xlin))).ColumnWidth = 6
                Xarchexel.Cells(Xlin, XCol) = "MOVIL TR."
                XCol = XCol + 1
                Xarchexel.Range("K" & Trim(str(Xlin))).ColumnWidth = 25
                Xarchexel.Cells(Xlin, XCol) = "INDICA"
                XCol = XCol + 1
                Xarchexel.Range("L" & Trim(str(Xlin))).ColumnWidth = 40
                Xarchexel.Cells(Xlin, XCol) = "MOTIVO CONSULTA"
                XCol = XCol + 1
                Xarchexel.Range("M" & Trim(str(Xlin))).ColumnWidth = 20
                Xarchexel.Cells(Xlin, XCol) = "ASISTE"
                XCol = XCol + 1
                Xarchexel.Range("N" & Trim(str(Xlin))).ColumnWidth = 40
                Xarchexel.Cells(Xlin, XCol) = "DIAGNOSTICO"
                XCol = XCol + 1
                Xarchexel.Range("O" & Trim(str(Xlin))).ColumnWidth = 15
                Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = "H.SAL TR"
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = "H.LLEGA CA"
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = "H.SALE CA"
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = "EN ZONA"
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = "DESPACHA"
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = "CODIGO"
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = "MOV.LLAM"
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = "HORA REA"
                XCol = XCol + 1
                Xarchexel.Range("X" & Trim(str(Xlin))).ColumnWidth = 18
                Xarchexel.Cells(Xlin, XCol) = "DEMORA EN CA"
                XCol = XCol + 1
                Xarchexel.Range("Y" & Trim(str(Xlin))).ColumnWidth = 18
                Xarchexel.Cells(Xlin, XCol) = "DEMORA TRASL"
                XCol = XCol + 1
                Xarchexel.Range("Z" & Trim(str(Xlin))).ColumnWidth = 12
                Xarchexel.Cells(Xlin, XCol) = "AFT"
                XCol = XCol + 1
                Xarchexel.Range("AA" & Trim(str(Xlin))).ColumnWidth = 15
                Xarchexel.Cells(Xlin, XCol) = "COSTO"
                XCol = XCol + 1
                Xarchexel.Range("AB" & Trim(str(Xlin))).ColumnWidth = 15
                Xarchexel.Cells(Xlin, XCol) = "NRO.FACT."
                XCol = XCol + 1
                Xarchexel.Range("AC" & Trim(str(Xlin))).ColumnWidth = 15
                Xarchexel.Cells(Xlin, XCol) = "TIPO FACT"
                XCol = XCol + 1
                Xarchexel.Range("AD" & Trim(str(Xlin))).ColumnWidth = 60
                Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
                XCol = XCol + 1
                Xarchexel.Range("AE" & Trim(str(Xlin))).ColumnWidth = 12
                Xarchexel.Cells(Xlin, XCol) = "COD.FIN"
                XCol = XCol + 1
                Xarchexel.Range("AF" & Trim(str(Xlin))).ColumnWidth = 10
                Xarchexel.Cells(Xlin, XCol) = "ZONA"
                XCol = XCol + 1
                Xarchexel.Range("AG" & Trim(str(Xlin))).ColumnWidth = 20
                Xarchexel.Cells(Xlin, XCol) = "TIPO TRASLADO"
                XCol = XCol + 1
                Xarchexel.Range("AH" & Trim(str(Xlin))).ColumnWidth = 12
                Xarchexel.Cells(Xlin, XCol) = "T.RESP"
                XCol = XCol + 1
                Xarchexel.Range("AI" & Trim(str(Xlin))).ColumnWidth = 12
                Xarchexel.Cells(Xlin, XCol) = "HORA_REC"
                XCol = XCol + 1
                Xarchexel.Range("AJ" & Trim(str(Xlin))).ColumnWidth = 12
                Xarchexel.Cells(Xlin, XCol) = "DEMORA ASIG.TR"
                XCol = XCol + 1
                Xarchexel.Range("AK" & Trim(str(Xlin))).ColumnWidth = 70
                Xarchexel.Cells(Xlin, XCol) = "OBSERV."
                XCol = XCol + 1
                Xarchexel.Range("AL" & Trim(str(Xlin))).ColumnWidth = 14
                Xarchexel.Cells(Xlin, XCol) = "TELEFONISTA"
                XCol = XCol + 1
                Xarchexel.Range("AM" & Trim(str(Xlin))).ColumnWidth = 14
                Xarchexel.Cells(Xlin, XCol) = "CONV.PADRÓN"
                XCol = XCol + 1
                Xarchexel.Range("AN" & Trim(str(Xlin))).ColumnWidth = 14
                Xarchexel.Cells(Xlin, XCol) = "ESTADO"
                XCol = XCol + 1
                Xarchexel.Range("AO" & Trim(str(Xlin))).ColumnWidth = 14
                Xarchexel.Cells(Xlin, XCol) = "DEUDA $"
                XCol = XCol + 1
                Xarchexel.Range("AP" & Trim(str(Xlin))).ColumnWidth = 40
                Xarchexel.Cells(Xlin, XCol) = "CONVENIO DESC"
                
                If Check7.Value = 1 Then
                   XCol = XCol + 1
                   Xarchexel.Range("AM" & Trim(str(Xlin))).ColumnWidth = 14
                   Xarchexel.Cells(Xlin, XCol) = "HORA ASIG.TR"
                   
'                   XCol = XCol + 1
'                   Xarchexel.Range("AL" & Trim(str(Xlin))).ColumnWidth = 12
'                   Xarchexel.Cells(Xlin, XCol) = "CANT_DIA"
                Else
'                   XCol = XCol + 1
'                   Xarchexel.Range("AK" & Trim(str(Xlin))).ColumnWidth = 12
'                   Xarchexel.Cells(Xlin, XCol) = "CANT_DIA"
                End If
                
                Xlin = Xlin + 1
                XCol = 1
                Dim Xnumera, Xbandquehago As Integer
                Xnumera = 1
                Xbandquehago = 0
                If data_lla.Recordset.RecordCount > 0 Then
                   Xqdia = Day(data_lla.Recordset("fecha"))
'            Data1.Recordset("descol") = t_cod.Text
                   Do While Not data_lla.Recordset.EOF
                      If Combo3.ListIndex = 0 Or Combo3.ListIndex = 1 Or Combo3.ListIndex = 2 Or _
                         Combo3.ListIndex = 3 Or Combo3.ListIndex = 4 Or Combo3.ListIndex = 5 Or _
                         Combo3.ListIndex = 6 Or Combo3.ListIndex = 15 Or Combo3.ListIndex = 10 Then
                         If data_lla.Recordset("categ") = "SA" Or data_lla.Recordset("categ") = "SAP" Or _
                            data_lla.Recordset("categ") = "EMERN" Or data_lla.Recordset("categ") = "911" Or _
                            data_lla.Recordset("categ") = "MSP" Or data_lla.Recordset("codmot") = "C" Or _
                            data_lla.Recordset("categ") = "UDEMM" Or data_lla.Recordset("categ") = "CERSEM" Then
                            Xbandquehago = 3
                         Else
                            data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lla.Recordset("categ") & "'"
                            data_conv.Refresh
                            If data_conv.Recordset.RecordCount > 0 Then
                               If IsNull(data_conv.Recordset("cnv_grupo")) = False Then
                                  If data_conv.Recordset("cnv_grupo") = Combo3.Text Then
                                     Xbandquehago = 0
                                  Else
                                     Xbandquehago = 3
                                  End If
                               Else
                                  Xbandquehago = 3
                               End If
                            Else
                               Xbandquehago = 3
                            End If
                         End If
                      Else
                         If Combo2.ListIndex = 14 Then
                            If data_lla.Recordset.EOF = False Then
                               data_lla22.RecordSource = "Select * from resplla where nro =" & data_lla.Recordset("nrolla")
                               data_lla22.Refresh
                                If data_lla22.Recordset.RecordCount > 0 Then
                                   If IsNull(data_lla22.Recordset("fecpas")) = False And IsNull(data_lla22.Recordset("horpas")) = False Then
                                      Xbandquehago = 0
                                   Else
                                      Xbandquehago = 3
                                   End If
                                Else
                                   Xbandquehago = 3
                                End If
                            Else
                               Xbandquehago = 3
                            End If
                         Else
                            If Combo2.ListIndex = 16 Then
                               data_lla22.RecordSource = "Select * from resplla where nro =" & data_lla.Recordset("nrolla")
                               data_lla22.Refresh
                               If data_lla22.Recordset.RecordCount > 0 Then
                                  If IsNull(data_lla22.Recordset("hzona")) = False Then
                                     Xbandquehago = 0
'                                        data_inf.Recordset("activo") = data_lla22.Recordset("hzona")
'                                        data_inf.Recordset("timdes") = data_lla22.Recordset("timdes")
                                  Else
                                     Xbandquehago = 3
                                  End If
                               Else
                                  Xbandquehago = 3
                               End If
                            Else
                               If Combo2.ListIndex = 17 Then
                                  If data_lla.Recordset.EOF = False Then
                                     If IsNull(data_lla.Recordset("nrolla")) = False Then
                                        data_lla22.RecordSource = "Select * from resplla where nro =" & data_lla.Recordset("nrolla")
                                        data_lla22.Refresh
                                        If data_lla22.Recordset.RecordCount > 0 Then
                                           If IsNull(data_lla22.Recordset("mm")) = False Then
                                              If data_lla22.Recordset("mm") = 2 Then
                                                 Xbandquehago = 0
                                              Else
                                                 Xbandquehago = 3
                                              End If
                                           Else
                                              Xbandquehago = 3
                                           End If
                                        Else
                                           Xbandquehago = 3
                                        End If
                                     Else
                                        Xbandquehago = 3
                                     End If
                                  Else
                                     Xbandquehago = 3
                                  End If
                               Else
                                  If Combo2.ListIndex = 2 Then
                                     If Check5.Value <> 1 Then
                                        If Combo3.ListIndex = 12 Then
                                           data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lla.Recordset("categ") & "' and cnv_colrec ='" & "M" & "'"
                                           data_conv.Refresh
                                           If data_conv.Recordset.RecordCount > 0 Then
                                              If IsNull(data_conv.Recordset("cnv_cant_r")) = False Then
                                                 If data_conv.Recordset("cnv_cant_r") = 1 Or data_conv.Recordset("cnv_pmserv") = 1 Then
                                                    Xbandquehago = 0
                                                 Else
                                                    Xbandquehago = 3
                                                 End If
                                              Else
                                                 If data_conv.Recordset("cnv_pmserv") = 1 Then
                                                    Xbandquehago = 0
                                                 Else
                                                    Xbandquehago = 3
                                                 End If
                                              End If
                                           Else
                                              Xbandquehago = 3
                                           End If
                                        Else
                                           Xbandquehago = 0
                                        End If
                                     Else
                                        Xbandquehago = 0
                                     End If
                                  Else
                                     Xbandquehago = 0
                                  End If
                               End If
                            End If
                         End If
                      End If
                      If Xbandquehago = 3 Then
                      Else
                         If data_lla.Recordset.EOF = False Then
                            Xarchexel.Cells(Xlin, XCol) = Xnumera
                            XCol = XCol + 1
                            Xarchexel.Cells(Xlin, XCol) = "'" & Format(data_lla.Recordset("fecha"), "dd/mm/yyyy")
                            XCol = XCol + 1
                            If IsNull(data_lla.Recordset("nombre")) = False Then
                               If Trim(data_lla.Recordset("nombre")) <> "" Then
                                  Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("nombre")
                               Else
                                  Xarchexel.Cells(Xlin, XCol) = "NN"
                               End If
                            Else
                               Xarchexel.Cells(Xlin, XCol) = "NN"
                            End If
                            XCol = XCol + 1
                            If Combo1.ListIndex = 3 And Combo2.ListIndex = 0 Then
                               data_lla22.RecordSource = "Select * from resplla where nro =" & data_lla.Recordset("nrolla")
                               data_lla22.Refresh
                               If IsNull(data_lla.Recordset("ci")) = False Then
                                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_lla.Recordset("ci"))) & "-0"
                               Else
                                  Xarchexel.Cells(Xlin, XCol) = "0-0"
                               End If
                            Else
                               If IsNull(data_lla.Recordset("ci")) = False Then
                                  data_lla22.RecordSource = "Select * from resplla where nro =" & data_lla.Recordset("nrolla")
                                  data_lla22.Refresh
                                  If data_lla22.Recordset.RecordCount > 0 Then
                                     If IsNull(data_lla22.Recordset("mes")) = False Then
                                        Xarchexel.Cells(Xlin, XCol) = Trim(str(data_lla.Recordset("ci"))) & "-" & Trim(str(data_lla22.Recordset("mes")))
                                     Else
                                        Xarchexel.Cells(Xlin, XCol) = Trim(str(data_lla.Recordset("ci"))) & "-0"
                                     End If
                                  Else
                                     Xarchexel.Cells(Xlin, XCol) = Trim(str(data_lla.Recordset("ci"))) & "-0"
                                  End If
                               Else
                                  data_lla22.RecordSource = "Select * from resplla where nro =" & data_lla.Recordset("nrolla")
                                  data_lla22.Refresh
                                  Xarchexel.Cells(Xlin, XCol) = "0-0"
                               End If
                            End If
                            XCol = XCol + 1
                            If IsNull(data_lla.Recordset("unied")) = False Then
                               If data_lla.Recordset("unied") = 3 Then
                                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_lla.Recordset("edad"))) & " AÑOS"
                               Else
                                  If data_lla.Recordset("unied") = 2 Then
                                     Xarchexel.Cells(Xlin, XCol) = Trim(str(data_lla.Recordset("edad"))) & " MESES"
                                  Else
                                     If data_lla.Recordset("unied") = 1 Then
                                        Xarchexel.Cells(Xlin, XCol) = Trim(str(data_lla.Recordset("edad"))) & " DIAS"
                                     Else
                                        Xarchexel.Cells(Xlin, XCol) = Trim(str(data_lla.Recordset("edad"))) & " AÑOS"
                                     End If
                                  End If
                               End If
                            Else
                               Xarchexel.Cells(Xlin, XCol) = Trim(str(data_lla.Recordset("edad"))) & " A"
                            End If
                            XCol = XCol + 1
                            Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("matric")
                            XCol = XCol + 1
                            Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("categ")
                            XCol = XCol + 1
                            Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("motmov")
                            XCol = XCol + 1
                            Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("lugar")
                            XCol = XCol + 1
                            Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("movtras")
                            XCol = XCol + 1
                            Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("nommed")
                            XCol = XCol + 1
                            Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("obsmot")
                            XCol = XCol + 1
                            Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("dcobr")
                            XCol = XCol + 1
                            If Combo2.ListIndex = 17 Or Combo2.ListIndex = 16 Then
                               data_lla22.RecordSource = "Select * from resplla where nro =" & data_lla.Recordset("nrolla")
                               data_lla22.Refresh
                               If data_lla22.Recordset.RecordCount > 0 Then
                                   Xarchexel.Cells(Xlin, XCol) = Mid(data_lla22.Recordset("obsmot"), 1, 45)
                               
                               End If
                            Else
                               Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("diag")
                            End If
                            XCol = XCol + 1
                            Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("telef")
                            XCol = XCol + 1
                            Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("hsald")
                            XCol = XCol + 1
                            Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("hllega")
                            XCol = XCol + 1
                            Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("hor_cance")
                            XCol = XCol + 1
                            Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("hzona")
                            XCol = XCol + 1
                            If Combo2.ListIndex = 16 Or Combo2.ListIndex = 17 Then
                               If data_lla22.Recordset.RecordCount > 0 Then
                                  If IsNull(data_lla22.Recordset("timdes")) = False Then
                                     Xarchexel.Cells(Xlin, XCol) = data_lla22.Recordset("timdes")
''                                     Xarchexel.Cells(Xlin, XCol) = "00"
                                  Else
                                     Xarchexel.Cells(Xlin, XCol) = "S/D"
                                  End If
                               Else
                                  Xarchexel.Cells(Xlin, XCol) = "S/D"
                               End If
                            Else
                               Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("timdes")
                            End If
                            XCol = XCol + 1
                            Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("descol")
                            XCol = XCol + 1
                            If Combo2.ListIndex = 14 Then
                               If data_lla22.Recordset.RecordCount > 0 Then
                                  If IsNull(data_lla22.Recordset("trasla")) = False Then
                                     Xarchexel.Cells(Xlin, XCol) = data_lla22.Recordset("trasla")
                                  Else
                                     Xarchexel.Cells(Xlin, XCol) = 0
                                  End If
                               Else
                                  Xarchexel.Cells(Xlin, XCol) = 0
                               End If
                            Else
                               Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("movilpas")
                            End If
                            XCol = XCol + 1
                            Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("hor_rea")
                            XCol = XCol + 1
                            
                            If IsNull(data_lla.Recordset("hllega")) = False Then
                               If IsNull(data_lla.Recordset("hor_cance")) = False Then
                                  Xdeshor = Val(Mid(data_lla.Recordset("hllega"), 1, 2))
                                  Xhashor = Val(Mid(data_lla.Recordset("hor_cance"), 1, 2))
                                  Xdifhs = Xhashor - Xdeshor
                                  If Xdifhs < 0 Then
                                     Xdifhs = Xdifhs + 24
                                  End If
                                  Xdesmin = Val(Mid(data_lla.Recordset("hllega"), 4, 2))
                                  Xhasmin = Val(Mid(data_lla.Recordset("hor_cance"), 4, 2))
                                  Xdifmin = Xhasmin - Xdesmin
                                  If Xdifmin < 0 Then
                                     Xdifhs = Xdifhs - 1
                                     Xdifmin = Xdifmin + 60
                                  End If
                                  Xtotnumero = Xdifhs * 60
                                  Xtotnumero = Xtotnumero + Xdifmin
                                  Xarchexel.Cells(Xlin, XCol) = Xtotnumero
                               Else
                                  Xarchexel.Cells(Xlin, XCol) = 0
                               End If
                            Else
                               Xarchexel.Cells(Xlin, XCol) = 0
                            End If
                            XCol = XCol + 1
                            If IsNull(data_lla.Recordset("hsald")) = False Then
                               If IsNull(data_lla.Recordset("hzona")) = False Then
                                  Xdeshor = Val(Mid(data_lla.Recordset("hsald"), 1, 2))
                                  Xhashor = Val(Mid(data_lla.Recordset("hzona"), 1, 2))
                                  Xdifhs = Xhashor - Xdeshor
                                  If Xdifhs < 0 Then
                                     Xdifhs = Xdifhs + 24
                                  End If
                                  Xdesmin = Val(Mid(data_lla.Recordset("hsald"), 4, 2))
                                  Xhasmin = Val(Mid(data_lla.Recordset("hzona"), 4, 2))
                                  Xdifmin = Xhasmin - Xdesmin
                                  If Xdifmin < 0 Then
                                     Xdifhs = Xdifhs - 1
                                     Xdifmin = Xdifmin + 60
                                  End If
                                  Xtotnumero = Xdifhs * 60
                                  Xtotnumero = Xtotnumero + Xdifmin
                                  Xarchexel.Cells(Xlin, XCol) = Xtotnumero
                               Else
                                  Xarchexel.Cells(Xlin, XCol) = 0
                               End If
                            Else
                               Xarchexel.Cells(Xlin, XCol) = 0
                            End If
                            XCol = XCol + 1
                            If Combo2.ListIndex = 10 And Combo1.ListIndex = 2 Then
                               If IsNull(data_lla.Recordset("aft")) = False Then
                                  Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("aft")
                               Else
                                  Xarchexel.Cells(Xlin, XCol) = "Sin COD"
                               End If
                            Else
                                Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("ncobr")
                            End If
                            XCol = XCol + 1
                            Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("mes")
                            XCol = XCol + 1
                            Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("ano")
                            XCol = XCol + 1
                            If data_lla22.Recordset.RecordCount > 0 Then
                               If IsNull(data_lla22.Recordset("telef")) = False Then
                                  Xarchexel.Cells(Xlin, XCol) = data_lla22.Recordset("telef")
                               Else
                                  Xarchexel.Cells(Xlin, XCol) = "S/D"
                               End If
                            Else
                               Xarchexel.Cells(Xlin, XCol) = "S/C"
                            End If
                            XCol = XCol + 1
                            Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("referen")
                            
                            XCol = XCol + 1
                            Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("colormot")
                            
                            XCol = XCol + 1
                            Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("codzon")
                            XCol = XCol + 1
                            
                            If IsNull(data_lla.Recordset("trasla")) = False Then
                               If data_lla.Recordset("trasla") > 0 Then
                                  data_lla22.RecordSource = "Select * from traslados where idtrasl =" & data_lla.Recordset("trasla")
                                  data_lla22.Refresh
                                  If data_lla22.Recordset.RecordCount > 0 Then
                                     Xarchexel.Cells(Xlin, XCol) = data_lla22.Recordset("descrip")
                                  Else
                                     Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
                                  End If
                               End If
                            End If
                            XCol = XCol + 1
                            If IsNull(data_lla.Recordset("hora")) = False Then
                               If IsNull(data_lla.Recordset("hor_llega")) = False Then
                                  Xdeshor = Val(Mid(data_lla.Recordset("hora"), 1, 2))
                                  Xhashor = Val(Mid(data_lla.Recordset("hor_llega"), 1, 2))
                                  If Xdeshor <= Xhashor Then
                                     Xdifhs = Xhashor - Xdeshor
                                  Else
                                     Xdifhs = Xhashor - Xdeshor
                                     Xdifhs = Xdifhs + 24
                                  End If
                                  Xdesmin = Val(Mid(data_lla.Recordset("hora"), 4, 2))
                                  Xhasmin = Val(Mid(data_lla.Recordset("hor_llega"), 4, 2))
                                  If Xdesmin <= Xhasmin Then
                                     Xdifmin = Xhasmin - Xdesmin
                                  Else
                                     Xdifmin = Xhasmin - Xdesmin
                                     Xdifmin = Xdifmin + 60
                                     Xdifhs = Xdifhs - 1
                                  End If
                                  Xtotnumero = Xdifhs * 60
                                  Xtotnumero = Xtotnumero + Xdifmin
                                  Xarchexel.Cells(Xlin, XCol) = Xtotnumero
                               Else
                                  Xarchexel.Cells(Xlin, XCol) = "00:00"
                               End If
                            Else
                               Xarchexel.Cells(Xlin, XCol) = "00:00"
                            End If
                            XCol = XCol + 1
                            Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("hora")
                            XCol = XCol + 1
                            If Check7.Value = 1 Then
                               data_lla22.RecordSource = "Select * from resplla where nro =" & data_lla.Recordset("nrolla")
                               data_lla22.Refresh
                               If data_lla22.Recordset.RecordCount > 0 Then
                                  If IsNull(data_lla22.Recordset("hor_llega")) = False Then
                                     If IsNull(data_lla.Recordset("hsald")) = False Then
                                        Xdeshor = Val(Mid(data_lla22.Recordset("hor_llega"), 1, 2))
                                        Xhashor = Val(Mid(data_lla.Recordset("hsald"), 1, 2))
                                        Xdifhs = Xhashor - Xdeshor
                                        If Xdifhs < 0 Then
                                           Xdifhs = Xdifhs + 24
                                        End If
                                        Xdesmin = Val(Mid(data_lla22.Recordset("hor_llega"), 4, 2))
                                        Xhasmin = Val(Mid(data_lla.Recordset("hsald"), 4, 2))
                                        Xdifmin = Xhasmin - Xdesmin
                                        If Xdifmin < 0 Then
                                           Xdifhs = Xdifhs - 1
                                           Xdifmin = Xdifmin + 60
                                        End If
                                        Xtotnumero = Xdifhs * 60
                                        Xtotnumero = Xtotnumero + Xdifmin
                                        Xarchexel.Cells(Xlin, XCol) = Xtotnumero
                                     Else
                                        Xarchexel.Cells(Xlin, XCol) = 0
                                     End If
                                  Else
                                     Xarchexel.Cells(Xlin, XCol) = 0
                                  End If
                               Else
                                  Xarchexel.Cells(Xlin, XCol) = 0
                               End If
                            Else
                                Xarchexel.Cells(Xlin, XCol) = 0
                            End If
                            XCol = XCol + 1
                            If IsNull(data_lla.Recordset("obs")) = False Then
                               Xarchexel.Cells(Xlin, XCol) = Mid(data_lla.Recordset("obs"), 1, 190)
                            Else
                               Xarchexel.Cells(Xlin, XCol) = "Sin Obs"
                            End If
                            XCol = XCol + 1
                            If IsNull(data_lla.Recordset("usuario")) = False Then
                               Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("usuario")
                            Else
                               Xarchexel.Cells(Xlin, XCol) = "Sin Dato"
                            End If
                            If Check7.Value = 1 Then
                               XCol = XCol + 1
                               If data_lla22.Recordset.RecordCount > 0 Then
                                  If IsNull(data_lla22.Recordset("hor_llega")) = False Then
                                     Xarchexel.Cells(Xlin, XCol) = Mid(data_lla22.Recordset("hor_llega"), 1, 5)
                                  Else
                                     Xarchexel.Cells(Xlin, XCol) = "00:00"
                                  End If
                               Else
                                  Xarchexel.Cells(Xlin, XCol) = "00:00"
                               End If
                            Else
                               XCol = XCol + 1
                            End If
                            XCol = XCol + 1
                            If Check8.Value = 1 Then
                               If IsNull(data_lla.Recordset("ci")) = False Then
                                  data_lla22.RecordSource = "select * from clientes where cl_cedula =" & data_lla.Recordset("ci")
                                  data_lla22.Refresh
                                  If data_lla22.Recordset.RecordCount > 0 Then
                                     Xarchexel.Cells(Xlin, XCol) = data_lla22.Recordset("cl_codconv")
                                     XCol = XCol + 1
                                     If IsNull(data_lla22.Recordset("fecha_baja")) = False Then
                                        Xarchexel.Cells(Xlin, XCol) = "BAJA"
                                     Else
                                        Xarchexel.Cells(Xlin, XCol) = "ACTIVO"
                                     End If
                                     Xdeudas = 0
                                     data_conv.RecordSource = "select * from deudas where cliente =" & data_lla22.Recordset("cl_codigo") & " and fecha_pago is null"
                                     data_conv.Refresh
                                     If data_conv.Recordset.RecordCount > 0 Then
                                        data_conv.Recordset.MoveFirst
                                        Do While Not data_conv.Recordset.EOF
                                           Xdeudas = Xdeudas + data_conv.Recordset("total")
                                           data_conv.Recordset.MoveNext
                                        Loop
                                     End If
                                     XCol = XCol + 1
                                     Xarchexel.Cells(Xlin, XCol) = Xdeudas
                                  Else
                                     Xarchexel.Cells(Xlin, XCol) = "NO FIGURA"
                                  End If
                               Else
                                  Xarchexel.Cells(Xlin, XCol) = "NO FIGURA"
                               End If
                            Else
                               XCol = XCol + 1
                            End If
                            XCol = XCol + 1
                            Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("nomcat")
''aca
                            If Combo2.ListIndex = 10 Then
                               If data_lla.Recordset("categ") = "MSP" Then
                                  If IsNull(data_lla.Recordset("codmot")) = False Then
                                     If data_lla.Recordset("codmot") = "V" Or data_lla.Recordset("codmot") = "Z" Then
                                        Xmsp3 = Xmsp3 + 1
                                     Else
                                        If data_lla.Recordset("codmot") = "A" Or data_lla.Recordset("codmot") = "C" Then
                                           Xmsp2 = Xmsp2 + 1
                                        Else
                                           If data_lla.Recordset("codmot") = "R" Or data_lla.Recordset("codmot") = "N" Then
                                              Xmsp1 = Xmsp1 + 1
                                           Else
                                              Xmsp3 = Xmsp3 + 1
                                           End If
                                        End If
                                     End If
                                  Else
                                     Xmsp3 = Xmsp3 + 1
                                  End If
                                  If IsNull(data_lla.Recordset("colormot")) = False Then
                                     If data_lla.Recordset("colormot") = "V" Or data_lla.Recordset("colormot") = "Z" Then
                                        Xmsp3f = Xmsp3f + 1
                                     Else
                                        If data_lla.Recordset("colormot") = "A" Or data_lla.Recordset("colormot") = "C" Then
                                           Xmsp2f = Xmsp2f + 1
                                        Else
                                           If data_lla.Recordset("colormot") = "R" Or data_lla.Recordset("colormot") = "N" Then
                                              Xmsp1f = Xmsp1f + 1
                                           Else
                                              Xmsp3f = Xmsp3f + 1
                                           End If
                                        End If
                                     End If
                                  Else
                                     Xmsp3f = Xmsp3f + 1
                                  End If
                               Else
                                  If data_lla.Recordset("categ") = "911" Or data_lla.Recordset("categ") = "911B" Then
                                     If IsNull(data_lla.Recordset("codmot")) = False Then
                                        If data_lla.Recordset("codmot") = "V" Or data_lla.Recordset("codmot") = "Z" Then
                                           X9113 = X9113 + 1
                                        Else
                                           If data_lla.Recordset("codmot") = "A" Or data_lla.Recordset("codmot") = "C" Then
                                              X9112 = X9112 + 1
                                           Else
                                              If data_lla.Recordset("codmot") = "R" Or data_lla.Recordset("codmot") = "N" Then
                                                 X9111 = X9111 + 1
                                              Else
                                                 X9113 = X9113 + 1
                                              End If
                                           End If
                                        End If
                                     Else
                                        X9113 = X9113 + 1
                                     End If
                                     If IsNull(data_lla.Recordset("colormot")) = False Then
                                        If data_lla.Recordset("colormot") = "V" Or data_lla.Recordset("colormot") = "Z" Then
                                           X9113f = X9113f + 1
                                        Else
                                           If data_lla.Recordset("colormot") = "A" Or data_lla.Recordset("colormot") = "C" Then
                                              X9112f = X9112f + 1
                                           Else
                                              If data_lla.Recordset("colormot") = "R" Or data_lla.Recordset("colormot") = "N" Then
                                                 X9111f = X9111f + 1
                                              Else
                                                 X9113f = X9113f + 1
                                              End If
                                           End If
                                        End If
                                     Else
                                        X9113f = X9113f + 1
                                     End If
                                  Else
                                     If data_lla.Recordset("codmed") = 959 Then
                                        If IsNull(data_lla.Recordset("codmot")) = False Then
                                           If data_lla.Recordset("codmot") = "V" Or data_lla.Recordset("codmot") = "Z" Then
                                              Xterce3 = Xterce3 + 1
                                           Else
                                              If data_lla.Recordset("codmot") = "A" Or data_lla.Recordset("codmot") = "C" Then
                                                 Xterce2 = Xterce2 + 1
                                              Else
                                                 If data_lla.Recordset("codmot") = "R" Or data_lla.Recordset("codmot") = "N" Then
                                                    Xterce1 = Xterce1 + 1
                                                 Else
                                                    Xterce3 = Xterce3 + 1
                                                 End If
                                              End If
                                           End If
                                        Else
                                           Xterce3 = Xterce3 + 1
                                        End If
                                        If IsNull(data_lla.Recordset("colormot")) = False Then
                                           If data_lla.Recordset("colormot") = "V" Or data_lla.Recordset("colormot") = "Z" Then
                                              Xterce3f = Xterce3f + 1
                                           Else
                                              If data_lla.Recordset("colormot") = "A" Or data_lla.Recordset("colormot") = "C" Then
                                                 Xterce2f = Xterce2f + 1
                                              Else
                                                 If data_lla.Recordset("colormot") = "R" Or data_lla.Recordset("colormot") = "N" Then
                                                    Xterce1f = Xterce1f + 1
                                                 Else
                                                    Xterce3f = Xterce3f + 1
                                                 End If
                                              End If
                                           End If
                                        Else
                                           Xterce3f = Xterce3f + 1
                                        End If
                                     Else
                                        If IsNull(data_lla.Recordset("codmot")) = False Then
                                           If data_lla.Recordset("codmot") = "V" Or data_lla.Recordset("codmot") = "Z" Then
                                              Xtr3 = Xtr3 + 1
                                           Else
                                              If data_lla.Recordset("codmot") = "A" Or data_lla.Recordset("codmot") = "C" Then
                                                 Xtr2 = Xtr2 + 1
                                              Else
                                                 If data_lla.Recordset("codmot") = "R" Or data_lla.Recordset("codmot") = "N" Then
                                                    Xtr1 = Xtr1 + 1
                                                 Else
                                                    Xtr3 = Xtr3 + 1
                                                 End If
                                              End If
                                           End If
                                        Else
                                           Xtr3 = Xtr3 + 1
                                        End If
                                        If IsNull(data_lla.Recordset("colormot")) = False Then
                                           If data_lla.Recordset("colormot") = "V" Or data_lla.Recordset("colormot") = "Z" Then
                                              Xtr3f = Xtr3f + 1
                                           Else
                                              If data_lla.Recordset("colormot") = "A" Or data_lla.Recordset("colormot") = "C" Then
                                                 Xtr2f = Xtr2f + 1
                                              Else
                                                 If data_lla.Recordset("colormot") = "R" Or data_lla.Recordset("colormot") = "N" Then
                                                    Xtr1f = Xtr1f + 1
                                                 Else
                                                    Xtr3f = Xtr3f + 1
                                                 End If
                                              End If
                                           End If
                                        Else
                                           Xtr3f = Xtr3f + 1
                                        End If
                                     End If
                                  End If
                               End If
                            End If
                            Xlin = Xlin + 1
                            Xcolfija = XCol
                            XCol = 1
                            Xnumera = Xnumera + 1
                         End If
                      End If
                      Xbandquehago = 0
                                              
                      If data_lla.Recordset.EOF = False Then
                         data_lla.Recordset.MoveNext
                         pb.Value = pb.Value + 1
                      End If
                      Xtotnumero = 0
                   Loop
                   DoEvents
                End If
                Xlin = Xlin + 2
                XCol = 3
                Xnrocan = Xlin + 6
                Xarchexel.Range("A" & Trim(str(Xlin)), "AI" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
                Xarchexel.Range("A" & Trim(str(Xlin)), "AI" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
                Xarchexel.Range("A" & Trim(str(Xlin)), "AI" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
                Xarchexel.Range("A" & Trim(str(Xlin)), "AI" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
                Xarchexel.Range("A" & Trim(str(Xlin)), "AI" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
                Xarchexel.Range("A" & Trim(str(Xlin)), "AI" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
                               
                Xarchexel.Cells(Xlin, XCol) = "TIPO DE TRASLADO"
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = "CL-1 INICIAL"
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = "CL-2"
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = "CL-3"
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = "CL-1 FINAL"
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = "CL-2"
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = "CL-3"
                Xlin = Xlin + 1
                XCol = 3
                Xarchexel.Cells(Xlin, XCol) = "TRASLADOS MSP"
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = Trim(str(Xmsp1))
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = Trim(str(Xmsp2))
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = Trim(str(Xmsp3))
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = Trim(str(Xmsp1f))
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = Trim(str(Xmsp2f))
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = Trim(str(Xmsp3f))
                Xlin = Xlin + 1
                XCol = 3
                Xarchexel.Cells(Xlin, XCol) = "TRASLADOS 911"
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = Trim(str(X9111))
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = Trim(str(X9112))
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = Trim(str(X9113))
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = Trim(str(X9111f))
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = Trim(str(X9112f))
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = Trim(str(X9113f))
                Xlin = Xlin + 1
                XCol = 3
                Xarchexel.Cells(Xlin, XCol) = "TRASL.MÉDICOS DE TERCEROS"
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = Trim(str(Xterce1))
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = Trim(str(Xterce2))
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = Trim(str(Xterce3))
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = Trim(str(Xterce1f))
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = Trim(str(Xterce2f))
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = Trim(str(Xterce3f))
                Xlin = Xlin + 1
                XCol = 3
                Xarchexel.Cells(Xlin, XCol) = "RESTO DE TRASLADOS"
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = Trim(str(Xtr1))
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = Trim(str(Xtr2))
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = Trim(str(Xtr3))
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = Trim(str(Xtr1f))
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = Trim(str(Xtr2f))
                XCol = XCol + 1
                Xarchexel.Cells(Xlin, XCol) = Trim(str(Xtr3f))
                Xlibexel.Save
                Xlibexel.Close
                Xobjexel.Quit
                Xlabrir.Workbooks.Open Xarchtex, , False
                Xlabrir.Visible = True
                Xlabrir.WindowState = xlMaximized
            End If
         End If
         Command1.Enabled = True
         Command2.Enabled = True
         frm_infdesp2.MousePointer = 0
         MsgBox "Proceso terminado"
         Combo3.ListIndex = -1
         If Check1.Value = 1 Then
         Else
            If Combo1.ListIndex = 0 Then
               cr1.ReportTitle = "CERTIFICACIONES REALIZADAS DESDE: " & mfd.Text & " HASTA: " & mfh.Text
               cr1.ReportFileName = App.path & "\infcertifd.rpt"
               cr1.Action = 1
            End If
            If Combo1.ListIndex = 1 Then
               cr1.ReportTitle = "ACTOS DE ENFERMERIA DESDE: " & mfd.Text & " HASTA: " & mfh.Text
               cr1.ReportFileName = App.path & "\infenferd.rpt"
               cr1.Action = 1
            End If
            If Combo1.ListIndex = 2 Then
               If Combo2.ListIndex = 4 Then
                  cr1.ReportTitle = "TRASLADOS ACCIDENTES DE TRANSITO DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                  cr1.ReportFileName = App.path & "\infllaacc.rpt"
                  cr1.Action = 1
               Else
                  If Combo2.ListIndex = 3 Then
                     cr1.ReportTitle = "TRASLADOS PARA DIRECCION TECNICA DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                     cr1.ReportFileName = App.path & "\inftrasdt.rpt"
                     cr1.Action = 1
                  Else
                     If Combo2.ListIndex = 13 Then
                        cr1.ReportTitle = "INFORME DE TRASLADOS POR CLAVE ---FECHA DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                        cr1.ReportFileName = App.path & "\inftras11.rpt"
                        cr1.Action = 1
                     Else
                        If Combo2.ListIndex = 7 Then
                           cr1.ReportTitle = "INFORME DE TRASLADOS ---FECHA DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                           cr1.ReportFileName = App.path & "\infllamadcp.rpt"
                           cr1.Action = 1
                        Else
                           cr1.ReportTitle = "INFORME DE TRASLADOS ---FECHA DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                           cr1.ReportFileName = App.path & "\inftras1.rpt"
                           cr1.Action = 1
                        End If
                     End If
                  End If
               End If
            End If
            If Combo1.ListIndex = 3 Then
               If Combo3.Text = "911" Then
                  cr1.ReportTitle = "LLAMADOS DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                  cr1.ReportFileName = App.path & "\infllamad.rpt"
                  cr1.Action = 1
               Else
                  If Combo3.ListIndex = 11 Then
                     cr1.ReportFileName = App.path & "\infsamedesp2.rpt"
                     cr1.Action = 1
                  Else
                     If Combo2.ListIndex = 4 Then
                        cr1.ReportTitle = "LLAMADOS ACCIDENTES DE TRANSITO DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                        cr1.ReportFileName = App.path & "\infllaacc.rpt"
                        cr1.Action = 1
                     Else
                        If Combo3.ListIndex = 14 Then
                        
                           cr1.ReportTitle = "LLAMADOS DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                           cr1.ReportFileName = App.path & "\infllamadf.rpt"
                           cr1.DiscardSavedData = True
                           cr1.Action = 1
                        Else
                           If Combo2.ListIndex = 9 Then
                              cr1.ReportTitle = "LLAMADOS POR BASES DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                              cr1.ReportFileName = App.path & "\infllamadb.rpt"
                              cr1.Action = 1
                           Else
                              If Combo2.ListIndex = 12 Then
                                 Dim Xsionodemoras As String
                                 Xsionodemoras = MsgBox("Desea imprimir SOLO DEMORAS >12MINUTOS?", vbInformation + vbYesNo)
                                 If Xsionodemoras = vbYes Then
                                    cr1.ReportTitle = "DEMORAS CHOFERES >12 MINUTOS DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                                    cr1.ReportFileName = App.path & "\infllachof2.rpt"
                                    cr1.Action = 1
                                 Else
                                    cr1.ReportTitle = "DEMORAS CHOFERES DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                                    cr1.ReportFileName = App.path & "\infllachof.rpt"
                                    cr1.Action = 1
                                 End If
                              Else
                                 If Combo2.ListIndex = 7 Then
                                    cr1.ReportTitle = "INFORME DE LLAMADOS ---FECHA DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                                    cr1.ReportFileName = App.path & "\infllamadcp.rpt"
                                    cr1.Action = 1
                                 Else
                                    If Combo2.ListIndex = 13 Then
                                       If Option1.Value = True Then
                                          cr1.ReportTitle = "INFORME DE LLAMADOS POR CLAVE DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                                          cr1.ReportFileName = App.path & "\infllamad11.rpt"
                                          cr1.Action = 1
                                       Else
                                          cr1.ReportTitle = "INFORME DE LLAMADOS POR CLAVE DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                                          cr1.ReportFileName = App.path & "\infllamad11n.rpt"
                                          cr1.Action = 1
                                       End If
                                    Else
                                       If Combo2.ListIndex = 14 Then
                                          If Option1.Value = True Then
                                             cr1.ReportTitle = "INFORME DE LLAMADOS POR ZONAS DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                                             cr1.ReportFileName = App.path & "\infllamad12.rpt"
                                             cr1.Action = 1
                                          Else
                                             cr1.ReportTitle = "INFORME DE LLAMADOS POR ZONAS DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                                             cr1.ReportFileName = App.path & "\infllamad12n.rpt"
                                             cr1.Action = 1
                                          End If
                                       Else
                                          If Combo2.ListIndex = 16 Then
                                             cr1.ReportTitle = "LLAMADOS CANCELADOS DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                                             cr1.ReportFileName = App.path & "\infllamad.rpt"
                                             cr1.Action = 1
                                          Else
                                             If Option2.Value = True Then
                                                cr1.ReportTitle = "LLAMADOS DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                                                cr1.ReportFileName = App.path & "\infllamadn.rpt"
                                                cr1.Action = 1
                                             Else
                                                If Combo2.ListIndex = 17 Or Combo2.ListIndex = 18 Then 'llamados a CMT
                                                   cr1.ReportTitle = "LLAMADOS CMT DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                                                   cr1.ReportFileName = App.path & "\infllamadcmt.rpt"
                                                   cr1.Action = 1
                                                Else
                                                   cr1.ReportTitle = "LLAMADOS DESDE: " & mfd.Text & " HASTA: " & mfh.Text
                                                   cr1.ReportFileName = App.path & "\infllamad.rpt"
                                                   cr1.Action = 1
                                                End If
                                             End If
                                          End If
                                       End If
                                    End If
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
         MsgBox "No existen registros"
         If Combo2.ListIndex <> 18 Then
            If Check1.Value = 1 Then
               Xlibexel.Close
               Xobjexel.Quit
            End If
         End If
      End If
   End If
End If

Command1.Enabled = True
Command2.Enabled = True
frm_infdesp2.MousePointer = 0

'Exit Sub
'Queesinfdesp:
'             If Err.Number = 3155 Then
'                frm_infdesp2.MousePointer = 0
'                MsgBox "Error al generar, cierre el programa y vuelva a intentar " & Err.Description
'                Xlibexel.Close
'                Xobjexel.Quit
''             Else
'                MsgBox "Error al generar, verifique datos o salga del programa y vuelva a intentar " & Err.Description
'                frm_infdesp2.MousePointer = 0
'                Xlibexel.Close
'                Xobjexel.Quit
'             End If
           
             
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
Dim Xobjexel As Excel.Application
Dim Xlibexel As Excel.Workbook
Dim Xarchexel As New Excel.Worksheet
Dim XCol, Xlin, Xnrocan, Xcolfija As Long
Dim Xqdia, Xcanxdia As Long
Dim Xarchtex As String
XCol = 1
Xlin = 1
Xnrocan = 1
Set Xobjexel = New Excel.Application

Dim Xlabrir2 As New Excel.Application

Set Xlibexel = Xobjexel.Workbooks.Add
'Set Xarchexel = Xlibexel.Worksheets.Add
Xlibexel.SaveAs ("C:\planillas\" & Trim(Combo1.Text) & "\" & Trim(str(Month(mfd.Text))) & Trim(str(Year(mfd.Text))) & ".xls")
Xarchtex = "C:\planillas\" & Trim(Combo1.Text) & "\" & Trim(str(Month(mfd.Text))) & Trim(str(Year(mfd.Text))) & ".xls"

'      Set Xarchexel = Xlibexel.Worksheets.Add
'      Xarchexel.Name = "HOJAUNA"
Dim Xtr1, Xtr2, Xtr3 As Integer
Xtr1 = 0
Xtr2 = 0
Xtre3 = 0

If Combo1.ListIndex = 2 And Combo2.ListIndex = 3 Then
   Xqdia = 0
    Xcanxdia = 0
    If data_lla.Recordset.RecordCount > 0 Then
       data_lla.Recordset.MoveLast
       pb.Max = data_lla.Recordset.RecordCount
       data_lla.Recordset.MoveFirst
       Xqdia = Day(data_lla.Recordset("fecha"))
       Set Xarchexel = Xlibexel.Worksheets.Add
       Xlin = 1
       XCol = 1
       Xarchexel.Name = Trim(str(Xqdia))
       Xarchexel.Cells(Xlin, XCol) = "SAPP S.A."
       Xlin = Xlin + 1
       XCol = XCol + 1
       Xarchexel.Range("A1", "C3").Font.Size = 16
       Xarchexel.Cells(Xlin, XCol) = "PLANILLA DE " & Trim(Combo1.Text) & " DESDE: " & mfd.Text & " HASTA: " & mfh.Text
       Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 120)
       XCol = 1
       Xlin = Xlin + 2
       Xnrocan = Xnrocan + Xlin
       Xarchexel.Range("A4", "V" & Trim(str(data_lla.Recordset.RecordCount + 4))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "V" & Trim(str(data_lla.Recordset.RecordCount + 4))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "V" & Trim(str(data_lla.Recordset.RecordCount + 4))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "V" & Trim(str(data_lla.Recordset.RecordCount + 4))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "V" & Trim(str(data_lla.Recordset.RecordCount + 4))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "V" & Trim(str(data_lla.Recordset.RecordCount + 4))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "V" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 5
         Xarchexel.Cells(Xlin, XCol) = "NRO."
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 6
         Xarchexel.Cells(Xlin, XCol) = "FECHA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 5
         Xarchexel.Cells(Xlin, XCol) = "EDAD"
         XCol = XCol + 1
         Xarchexel.Cells(Xlin, XCol) = "MAT."
         XCol = XCol + 1
         Xarchexel.Cells(Xlin, XCol) = "CAT."
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 20
         Xarchexel.Cells(Xlin, XCol) = "DESDE"
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 20
         XCol = XCol + 1
         Xarchexel.Cells(Xlin, XCol) = "DESTINO"
         XCol = XCol + 1
         Xarchexel.Range("I" & Trim(str(Xlin))).ColumnWidth = 6
         Xarchexel.Cells(Xlin, XCol) = "MOVIL TRAS"
         XCol = XCol + 1
         Xarchexel.Range("J" & Trim(str(Xlin))).ColumnWidth = 25
         Xarchexel.Cells(Xlin, XCol) = "INDICA"
         XCol = XCol + 1
         Xarchexel.Range("K" & Trim(str(Xlin))).ColumnWidth = 20
         Xarchexel.Cells(Xlin, XCol) = "ASISTE"
         XCol = XCol + 1
         Xarchexel.Range("L" & Trim(str(Xlin))).ColumnWidth = 40
         Xarchexel.Cells(Xlin, XCol) = "DIAGNOSTICO"
         XCol = XCol + 1
         Xarchexel.Range("M" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Cells(Xlin, XCol) = "H.SAL"
         XCol = XCol + 1
         Xarchexel.Cells(Xlin, XCol) = "H.LLEGA"
         XCol = XCol + 1
         Xarchexel.Cells(Xlin, XCol) = "H.SALE"
         XCol = XCol + 1
         Xarchexel.Cells(Xlin, XCol) = "EN ZONA"
         XCol = XCol + 1
         Xarchexel.Cells(Xlin, XCol) = "DESPACHA"
         XCol = XCol + 1
         Xarchexel.Cells(Xlin, XCol) = "MOVIL LL"
         XCol = XCol + 1
         Xarchexel.Range("T" & Trim(str(Xlin))).ColumnWidth = 18
        Xarchexel.Cells(Xlin, XCol) = "DEMORA EN CA"
         XCol = XCol + 1
         Xarchexel.Range("U" & Trim(str(Xlin))).ColumnWidth = 18
         Xarchexel.Cells(Xlin, XCol) = "DEMORA TRASL"
        
        Xlin = Xlin + 1
        XCol = 1
       Do While Not data_lla.Recordset.EOF
       'trasla in (1,2,3,4,5,6,7,8,9,10,11,12)
'          If data_lla.Recordset("trasla") = 1 Or data_lla.Recordset("trasla") = 2 Or data_lla.Recordset("trasla") = 3 Or _
'             data_lla.Recordset("trasla") = 4 Or data_lla.Recordset("trasla") = 5 Or data_lla.Recordset("trasla") = 6 Or _
'             data_lla.Recordset("trasla") = 7 Or data_lla.Recordset("trasla") = 8 Or data_lla.Recordset("trasla") = 9 Or _
'             data_lla.Recordset("trasla") = 10 Or data_lla.Recordset("trasla") = 11 Or data_lla.Recordset("trasla") = 12 Then
'             Xtr1 = Xtr1 + 1
'             If data_lla.Recordset("trasla") = 9 Or data_lla.Recordset("trasla") = 10 Or data_lla.Recordset("trasla") = 11 Then
'                Xtr2 = Xtr2 + 1
'             Else
'                If data_lla.Recordset("trasla") = 3 Then
'                   Xtr3 = Xtr3 + 1
'                End If
'             End If
'          End If
          Data1.Recordset.AddNew
          Data1.Recordset("fecha") = data_lla.Recordset("fecha")
          Data1.Recordset("trasla") = data_lla.Recordset("trasla")
          Data1.Recordset.Update
          If Xqdia = Day(data_lla.Recordset("fecha")) Then
          Else
            Xqdia = Day(data_lla.Recordset("fecha"))
            Set Xarchexel = Xlibexel.Worksheets.Add
            Xlin = 1
            XCol = 1
            Xarchexel.Name = Trim(str(Xqdia))
             Xarchexel.Cells(Xlin, XCol) = "SAPP S.A."
             Xlin = Xlin + 1
             XCol = XCol + 1
             Xarchexel.Range("A1", "C3").Font.Size = 16
             Xarchexel.Cells(Xlin, XCol) = "PLANILLA DE " & Trim(Combo1.Text) & " DESDE: " & mfd.Text & " HASTA: " & mfh.Text
             Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(115, 0, 120)
             XCol = 1
             Xlin = Xlin + 2
             Xnrocan = Xnrocan + Xlin
             Xarchexel.Range("A4", "V" & Trim(str(data_lla.Recordset.RecordCount + 4))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
             Xarchexel.Range("A4", "V" & Trim(str(data_lla.Recordset.RecordCount + 4))).Borders(xlInsideVertical).LineStyle = xlContinuous
             Xarchexel.Range("A4", "V" & Trim(str(data_lla.Recordset.RecordCount + 4))).Borders(xlEdgeBottom).LineStyle = xlContinuous
             Xarchexel.Range("A4", "V" & Trim(str(data_lla.Recordset.RecordCount + 4))).Borders(xlEdgeTop).LineStyle = xlContinuous
             Xarchexel.Range("A4", "V" & Trim(str(data_lla.Recordset.RecordCount + 4))).Borders(xlEdgeLeft).LineStyle = xlContinuous
             Xarchexel.Range("A4", "V" & Trim(str(data_lla.Recordset.RecordCount + 4))).Borders(xlEdgeRight).LineStyle = xlContinuous
             Xarchexel.Range("A" & Trim(str(Xlin)), "V" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 120)
             Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 5
             Xarchexel.Cells(Xlin, XCol) = "NRO."
             XCol = XCol + 1
             Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 6
             Xarchexel.Cells(Xlin, XCol) = "FECHA"
             XCol = XCol + 1
             Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 35
             Xarchexel.Cells(Xlin, XCol) = "NOMBRE"
             XCol = XCol + 1
             Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 5
             Xarchexel.Cells(Xlin, XCol) = "EDAD"
             XCol = XCol + 1
             Xarchexel.Cells(Xlin, XCol) = "MAT."
             XCol = XCol + 1
             Xarchexel.Cells(Xlin, XCol) = "CAT."
             XCol = XCol + 1
             Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 20
             Xarchexel.Cells(Xlin, XCol) = "DESDE"
             Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 20
             XCol = XCol + 1
             Xarchexel.Cells(Xlin, XCol) = "DESTINO"
             XCol = XCol + 1
             Xarchexel.Range("I" & Trim(str(Xlin))).ColumnWidth = 6
             Xarchexel.Cells(Xlin, XCol) = "MOVIL TRAS"
             XCol = XCol + 1
             Xarchexel.Range("J" & Trim(str(Xlin))).ColumnWidth = 25
             Xarchexel.Cells(Xlin, XCol) = "INDICA"
             XCol = XCol + 1
             Xarchexel.Range("K" & Trim(str(Xlin))).ColumnWidth = 20
             Xarchexel.Cells(Xlin, XCol) = "ASISTE"
             XCol = XCol + 1
             Xarchexel.Range("L" & Trim(str(Xlin))).ColumnWidth = 40
             Xarchexel.Cells(Xlin, XCol) = "DIAGNOSTICO"
             XCol = XCol + 1
             Xarchexel.Range("M" & Trim(str(Xlin))).ColumnWidth = 15
             Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
             XCol = XCol + 1
             Xarchexel.Cells(Xlin, XCol) = "H.SAL"
             XCol = XCol + 1
             Xarchexel.Cells(Xlin, XCol) = "H.LLEGA"
             XCol = XCol + 1
             Xarchexel.Cells(Xlin, XCol) = "H.SALE"
             XCol = XCol + 1
             Xarchexel.Cells(Xlin, XCol) = "EN ZONA"
             XCol = XCol + 1
             Xarchexel.Cells(Xlin, XCol) = "DESPACHA"
             XCol = XCol + 1
             Xarchexel.Cells(Xlin, XCol) = "MOVIL LL"
             XCol = XCol + 1
             Xarchexel.Range("T" & Trim(str(Xlin))).ColumnWidth = 18
             Xarchexel.Cells(Xlin, XCol) = "DEMORA EN CA"
             XCol = XCol + 1
             Xarchexel.Range("U" & Trim(str(Xlin))).ColumnWidth = 18
             Xarchexel.Cells(Xlin, XCol) = "DEMORA TRASL"
            
            Xlin = Xlin + 1
            XCol = 1
          End If
       
'       Do While Not data_inf.Recordset.EOF
          Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("nro")
          XCol = XCol + 1
          Xarchexel.Cells(Xlin, XCol) = Day(data_lla.Recordset("fecha"))
          XCol = XCol + 1
          Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("nombre")
          XCol = XCol + 1
          Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("edad")
          XCol = XCol + 1
          Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("matric")
          XCol = XCol + 1
          Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("categ")
          XCol = XCol + 1
          Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("motmov")
          XCol = XCol + 1
          Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("lugar")
          XCol = XCol + 1
          Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("movtras")
          XCol = XCol + 1
          Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("nommed")
          XCol = XCol + 1
          Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("dcobr")
          XCol = XCol + 1
          Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("diag")
          XCol = XCol + 1
          Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("telef")
          XCol = XCol + 1
          Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("hsald")
          XCol = XCol + 1
          Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("hllega")
          XCol = XCol + 1
          Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("hor_cance")
          XCol = XCol + 1
          Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("hzona")
          XCol = XCol + 1
          Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("timdes")
          XCol = XCol + 1
          Xarchexel.Cells(Xlin, XCol) = data_lla.Recordset("movilpas")
          XCol = XCol + 1
          If IsNull(data_lla.Recordset("hllega")) = False Then
             If IsNull(data_lla.Recordset("hor_cance")) = False Then
                Dim Xdeshor, Xhashor, Xdifhs As Long
                Dim Xdesmin, Xhasmin, Xdifmin As Long
                Xdeshor = Val(Mid(data_lla.Recordset("hllega"), 1, 2))
                Xhashor = Val(Mid(data_lla.Recordset("hor_cance"), 1, 2))
                If Xdeshor <= Xhashor Then
                   Xdifhs = Xhashor - Xdeshor
                Else
                   Xdifhs = Xhashor - Xdeshor
                   Xdifhs = Xdifhs + 24
                End If
                Xdesmin = Val(Mid(data_lla.Recordset("hllega"), 4, 2))
                Xhasmin = Val(Mid(data_lla.Recordset("hor_cance"), 4, 2))
                If Xdesmin <= Xhasmin Then
                   Xdifmin = Xhasmin - Xdesmin
                Else
                   Xdifmin = Xhasmin - Xdesmin
                   Xdifmin = Xdifmin + 60
                   Xdifhs = Xdifhs - 1
                End If
                Xarchexel.Cells(Xlin, XCol) = Trim(str(Xdifhs)) & ":" & Trim(str(Xdifmin))
             
             Else
                Xarchexel.Cells(Xlin, XCol) = "00:00"
             End If
          Else
             Xarchexel.Cells(Xlin, XCol) = "00:00"
          End If
          XCol = XCol + 1
          If IsNull(data_lla.Recordset("hor_rea")) = False Then
             If IsNull(data_lla.Recordset("hsald")) = False Then
                Xdeshor = Val(Mid(data_lla.Recordset("hor_rea"), 1, 2))
                Xhashor = Val(Mid(data_lla.Recordset("hsald"), 1, 2))
                If Xdeshor <= Xhashor Then
                   Xdifhs = Xhashor - Xdeshor
                Else
                   Xdifhs = Xhashor - Xdeshor
                   Xdifhs = Xdifhs + 24
                End If
                Xdesmin = Val(Mid(data_lla.Recordset("hor_rea"), 4, 2))
                Xhasmin = Val(Mid(data_lla.Recordset("hsald"), 4, 2))
                If Xdesmin <= Xhasmin Then
                   Xdifmin = Xhasmin - Xdesmin
                Else
                   Xdifmin = Xhasmin - Xdesmin
                   Xdifmin = Xdifmin + 60
                   Xdifhs = Xdifhs - 1
                End If
                Xarchexel.Cells(Xlin, XCol) = Trim(str(Xdifhs)) & ":" & Trim(str(Xdifmin))
             Else
                Xarchexel.Cells(Xlin, XCol) = "00:00"
             End If
          Else
             Xarchexel.Cells(Xlin, XCol) = "00:00"
          End If
          XCol = XCol + 1
          If data_lla.Recordset.EOF = False Then
             Xqdia = Day(data_lla.Recordset("fecha"))
          Else
             Xqdia = 0
          End If
          data_lla.Recordset.MoveNext
          pb.Value = pb.Value + 1
          Xlin = Xlin + 1
          Xcolfija = XCol
          XCol = 1
       Loop
       Xlin = Xlin - 1
       Xarchexel.Cells(Xlin, Xcolfija) = Xcanxdia
       Set Xarchexel = Xlibexel.Worksheets.Add
       Xlin = 1
       XCol = 1
       Xarchexel.Name = Trim("Resumen")
       Xarchexel.Cells(Xlin, XCol) = "SAPP S.A."
       Xlin = Xlin + 2
       XCol = 1
       Xarchexel.Range("A1", "C3").Font.Size = 16
       Xarchexel.Cells(Xlin, XCol) = "RESUMEN " & Month(mfd.Text) & "/" & Year(mfd.Text)
        'Xarchexel.Range("B" & Trim(Str(Xlin)), "I" & Trim(Str(Xlin))).Interior.color = RGB(0, 200, 120)
       Xarchexel.Range("A3").Interior.color = RGB(115, 120, 0)
        
        XCol = 1
        Xlin = Xlin + 1
        'Xnrocan = Xnrocan + Xlin
       Xarchexel.Range("B4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("B4", "AN" & Trim(str(15))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
       Xarchexel.Range("B4", "AN" & Trim(str(15))).Borders(xlInsideVertical).LineStyle = xlContinuous
       Xarchexel.Range("B4", "AN" & Trim(str(15))).Borders(xlEdgeBottom).LineStyle = xlContinuous
       Xarchexel.Range("B4", "AN" & Trim(str(15))).Borders(xlEdgeTop).LineStyle = xlContinuous
       Xarchexel.Range("B4", "AN" & Trim(str(15))).Borders(xlEdgeLeft).LineStyle = xlContinuous
       Xarchexel.Range("B4", "AN" & Trim(str(15))).Borders(xlEdgeRight).LineStyle = xlContinuous
       Xarchexel.Range("B4" & Trim(str(Xlin)), "AN" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
       Xarchexel.Range("A4" & Trim(str(Xlin))).ColumnWidth = 45
       Xarchexel.Range("B4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("C4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("D4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("E4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("F4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("G4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("H4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("I4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("J4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("K4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("L4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("M4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("N4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("O4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("P4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("Q4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("R4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("S4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("T4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("U4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("V4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("W4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("X4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("Y4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("Z4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("AA4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("AB4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("AC4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("AD4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("AE4" & Trim(str(Xlin))).ColumnWidth = 4
       Xarchexel.Range("AF4" & Trim(str(Xlin))).ColumnWidth = 4
       Dim Xdiass As Long
       Xdiass = 1
       XCol = 2
       Do While Xdiass <= 31
          Xarchexel.Cells(Xlin, XCol) = Trim(str(Xdiass))
          Xdiass = Xdiass + 1
          XCol = XCol + 1
       Loop
       Xarchexel.Cells(Xlin, XCol) = "TOTAL"
       
       XCol = 1
       Xlin = 5
       Xarchexel.Cells(Xlin, XCol) = "TOTAL DE TRASLADOS:"
       Xlin = Xlin + 1
       Xarchexel.Cells(Xlin, XCol) = "TOTAL COORDINADOS:"
       Xlin = Xlin + 1
       Xarchexel.Cells(Xlin, XCol) = "TOTAL EN ZONA:"
       Xlin = Xlin + 1
       Xarchexel.Cells(Xlin, XCol) = "TOTAL A MONTEVIDEO:"
       Xlin = Xlin + 1
       Xarchexel.Cells(Xlin, XCol) = "COORDINADOS A MDEO:"
       Xlin = Xlin + 1
       Xarchexel.Cells(Xlin, XCol) = "TOTAL DE A.PROT.:"
       Xlin = 5
       XCol = 2
       Dim Xtottrasg, Xtottrasgdos As Long
       Dim Xfectras As Date
       Xtottrasg = 0
       Xtottrasgdos = 0
       Data1.RecordSource = "Select * from inflla where trasla in (1,2,3,4,5,6,7,8,9,10,11,12) order by fecha"
       Data1.Refresh
       If Data1.Recordset.RecordCount > 0 Then
          Data1.Recordset.MoveFirst
          Xfectras = Data1.Recordset("fecha")
          Do While Not Data1.Recordset.EOF
             If Data1.Recordset("fecha") = Xfectras Then
                Xtottrasg = Xtottrasg + 1
                Xtottrasgdos = Xtottrasgdos + 1
             Else
                Xarchexel.Cells(Xlin, XCol) = Xtottrasg
                Xtottrasg = 1
                Xtottrasgdos = Xtottrasgdos + 1
                XCol = XCol + 1
             End If
             Xfectras = Data1.Recordset("fecha")
             Data1.Recordset.MoveNext
          Loop
          Xarchexel.Cells(Xlin, XCol) = Xtottrasg
          Xtottrasg = 0
          XCol = 33
          Xarchexel.Cells(Xlin, XCol) = Xtottrasgdos
          XCol = 2
          Xlin = Xlin + 1
       End If
       Xtottrasgdos = 0
       Data1.RecordSource = "Select * from inflla where trasla in (9,10,11) order by fecha"
       Data1.Refresh
       If Data1.Recordset.RecordCount > 0 Then
          Data1.Recordset.MoveFirst
          Xfectras = Data1.Recordset("fecha")
          Do While Not Data1.Recordset.EOF
             If Data1.Recordset("fecha") = Xfectras Then
                Xtottrasg = Xtottrasg + 1
                Xtottrasgdos = Xtottrasgdos + 1
             Else
                Xarchexel.Cells(Xlin, XCol) = Xtottrasg
                Xtottrasg = 1
                Xtottrasgdos = Xtottrasgdos + 1
                XCol = XCol + 1
             End If
             Xfectras = Data1.Recordset("fecha")
             Data1.Recordset.MoveNext
          Loop
          Xarchexel.Cells(Xlin, XCol) = Xtottrasg
          Xtottrasg = 0
          XCol = 33
          Xarchexel.Cells(Xlin, XCol) = Xtottrasgdos
          XCol = 2
          Xlin = Xlin + 1
       End If
       ' En zona
       Xtottrasgdos = 0
       Data1.RecordSource = "Select * from inflla where trasla =" & 3 & " order by fecha"
       Data1.Refresh
       If Data1.Recordset.RecordCount > 0 Then
          Data1.Recordset.MoveFirst
          Xfectras = Data1.Recordset("fecha")
          Do While Not Data1.Recordset.EOF
             If Data1.Recordset("fecha") = Xfectras Then
                Xtottrasg = Xtottrasg + 1
                Xtottrasgdos = Xtottrasgdos + 1
             Else
                Xarchexel.Cells(Xlin, XCol) = Xtottrasg
                Xtottrasg = 1
                Xtottrasgdos = Xtottrasgdos + 1
                XCol = XCol + 1
             End If
             Xfectras = Data1.Recordset("fecha")
             Data1.Recordset.MoveNext
          Loop
          Xarchexel.Cells(Xlin, XCol) = Xtottrasg
          Xtottrasg = 0
          XCol = 33
          Xarchexel.Cells(Xlin, XCol) = Xtottrasgdos
          XCol = 2
          Xlin = Xlin + 1
       End If
       
       DoEvents
    End If
    
    Xlibexel.Save
    Xlibexel.Close
    
    Xobjexel.Quit
'    Shell frm_menu.data_usuac.Recordset("destino") & "excel.exe " & Xarchtex, vbMaximizedFocus
    Xlabrir2.Workbooks.Open Xarchtex, , False
    Xlabrir2.Visible = True
    Xlabrir2.WindowState = xlMaximized

End If

End Sub

Private Sub Command4_Click()
Dim Xobjexel22 As Excel.Application
Dim Xlibexel22 As Excel.Workbook

Dim Xarchexel22 As New Excel.Worksheet
Dim Xqdia, Xcanxdia, Xlin, XCol As Integer
Dim Xtotreg As Long
Dim Xarchtex As String
Dim Textofecha As String

Dim Xlabrir3 As New Excel.Application
Dim Xlabancp As Integer
Dim Contar911 As Integer



Xlin = 1
XCol = 1

      Set Xobjexel22 = New Excel.Application

      Set Xlibexel22 = Xobjexel22.Workbooks.Add
      Set Xarchexel22 = Xlibexel22.Worksheets.Add

      Xarchexel22.Name = Trim(Combo1.Text)

      Xlibexel22.SaveAs ("C:\planillas\" & Trim(Combo1.Text) & "\" & Trim(str(Month(mfd.Text))) & Trim(str(Year(mfd.Text))) & ".xls")
      Xarchtex = "C:\planillas\" & Trim(Combo1.Text) & "\" & Trim(str(Month(mfd.Text))) & Trim(str(Year(mfd.Text))) & ".xls"
        
        Xqdia = 0
        Xcanxdia = 0
        Xarchexel22.Cells(Xlin, XCol) = "DEPARTAMENTO TI SAPP S.A."
        Xlin = Xlin + 1
        XCol = XCol + 1
        Xarchexel22.Range("A1", "C3").Font.Size = 16
        Xarchexel22.Cells(Xlin, XCol) = "PLANILLA DE " & Trim(Combo1.Text) & " DESDE: " & mfd.Text & " HASTA: " & mfh.Text
        Xarchexel22.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
        
        XCol = 1
        Xlin = Xlin + 2
        Xnrocan = Xnrocan + Xlin
        Xarchexel22.Range("A" & Trim(str(Xlin)), "E" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
        Xarchexel22.Range("A" & Trim(str(Xlin))).ColumnWidth = 6
        Xarchexel22.Cells(Xlin, XCol) = "DIA"
        XCol = XCol + 1
        Xarchexel22.Range("B" & Trim(str(Xlin))).ColumnWidth = 6
        Xarchexel22.Cells(Xlin, XCol) = "MES"
        XCol = XCol + 1
        Xarchexel22.Range("C" & Trim(str(Xlin))).ColumnWidth = 6
        Xarchexel22.Cells(Xlin, XCol) = "AÑO"
        XCol = XCol + 1
        Xarchexel22.Range("D" & Trim(str(Xlin))).ColumnWidth = 35
        Xarchexel22.Cells(Xlin, XCol) = "NOMBRE"
        XCol = XCol + 1
        Xarchexel22.Range("E" & Trim(str(Xlin))).ColumnWidth = 12
        Xarchexel22.Cells(Xlin, XCol) = "ZONA"
        If Check9.Value = 1 Then
           XCol = XCol + 1
           Xarchexel22.Range("F" & Trim(str(Xlin))).ColumnWidth = 12
           Xarchexel22.Cells(Xlin, XCol) = "MOVIL"
        End If
'        XCol = XCol + 1
'        Xarchexel22.Range("G" & Trim(str(Xlin))).ColumnWidth = 20
'        Xarchexel22.Cells(Xlin, XCol) = "CONVENIO"
        
        Xlin = Xlin + 1
        XCol = 1
        If data_lla.Recordset.RecordCount > 0 Then
           data_lla.Recordset.MoveLast
           pb.Max = pb.Max + data_lla.Recordset.RecordCount + 1
           data_lla.Recordset.MoveFirst
           Xqdia = Day(data_lla.Recordset("fecha"))
           Dim Xcantre As Integer
           Xcantre = 0
           If Combo1.ListIndex = 2 Then 'traslados
               Do While Not data_lla.Recordset.EOF
                  If data_lla.Recordset("categ") = "911" Or data_lla.Recordset("categ") = "911B" Then
                     Xlabancp = 3
                  Else
                     data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lla.Recordset("categ") & "'"
                     data_conv.Refresh
                     If data_conv.Recordset.RecordCount > 0 Then
                        If IsNull(data_conv.Recordset("cnv_grupo")) = False Then
                           If data_conv.Recordset("cnv_grupo") = "CCOU" Or _
                              data_conv.Recordset("cnv_grupo") = "SMI" Or _
                              data_conv.Recordset("cnv_grupo") = "UNIVERSAL" Or _
                              data_conv.Recordset("cnv_grupo") = "IMPASA" Or _
                              data_conv.Recordset("cnv_codigo") = "UNIVA" Or data_conv.Recordset("cnv_codigo") = "CGALIC" Or _
                              data_conv.Recordset("cnv_grupo") = "CASA DE GALICIA" Or _
                              data_conv.Recordset("cnv_codigo") = "IMP2" Or data_conv.Recordset("cnv_codigo") = "CASA3" Or _
                              data_conv.Recordset("cnv_codigo") = "SMI3" Or data_conv.Recordset("cnv_codigo") = "SMI4" Or _
                              data_conv.Recordset("cnv_codigo") = "CCSA" Or data_conv.Recordset("cnv_codigo") = "SMI5" Or _
                              data_conv.Recordset("cnv_grupo") = "H.EVANGELICO" Then
                              If data_lla.Recordset("movilpas") = 3 Or _
                                 data_lla.Recordset("movilpas") = 206 Or data_lla.Recordset("movilpas") = 208 Or _
                                 data_lla.Recordset("movilpas") = 301 Or data_lla.Recordset("movilpas") = 207 Or _
                                 data_lla.Recordset("movilpas") = 306 Or data_lla.Recordset("movilpas") = 202 Or _
                                 data_lla.Recordset("movilpas") = 203 Or data_lla.Recordset("movilpas") = 501 Or _
                                 data_lla.Recordset("movilpas") = 620 Then
                                 If data_lla.Recordset("categ") = "MSP" Then
                                    Xlabancp = 3
                                 End If
                              Else
                                 Xlabancp = 3
                              End If
                           Else
                              Xlabancp = 3
                           End If
                        Else
                           Xlabancp = 3
                        End If
                     Else
                        Xlabancp = 3
                     End If
                  End If
                  If Xlabancp = 3 Then
                  Else
                     Xarchexel22.Cells(Xlin, XCol) = Trim(str(Day(data_lla.Recordset("fecha"))))
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = Trim(str(Month(data_lla.Recordset("fecha"))))
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = Trim(str(Year(data_lla.Recordset("fecha"))))
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = data_lla.Recordset("nombre")
                     XCol = XCol + 1
                     If data_lla.Recordset("codzon") = 4 Or data_lla.Recordset("codzon") = 5 Or data_lla.Recordset("codzon") = 6 Then
                        Xarchexel22.Cells(Xlin, XCol) = "Zona:3"
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "Zona:" & Trim(str(data_lla.Recordset("codzon")))
                     End If
                     If Check9.Value = 1 Then
                        XCol = XCol + 1
                        Xarchexel22.Cells(Xlin, XCol) = data_lla.Recordset("movtras")
                     End If
                     Xcantre = Xcantre + 1
                     Xlin = Xlin + 1
                     Xcolfija = XCol
                     XCol = 1
                  End If
                  Xlabancp = 0
                  data_lla.Recordset.MoveNext
                  pb.Value = pb.Value + 1
               Loop
               Xlin = Xlin + 1
               XCol = 2
               Xarchexel22.Cells(Xlin, XCol) = "TOTAL DE REGISTROS:" & Xcantre
               DoEvents
           Else
               Contar911 = 0
               Do While Not data_lla.Recordset.EOF
'                  If data_lla.Recordset("categ") = "911" Or data_lla.Recordset("categ") = "911B" Or data_lla.Recordset("categ") = "SEMM1" Or _
'                     data_lla.Recordset("categ") = "CERSEM" Or data_lla.Recordset("categ") = "UDEMM" Or data_lla.Recordset("categ") = "SEMM" Or _
'                     data_lla.Recordset("categ") = "UCM" Then
''                     Xlabancp = 0
'                  Else
                   If data_lla.Recordset("categ") = "CAAMEP" Then
                       Xlabancp = 3
                   Else
                        data_lla22.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lla.Recordset("categ") & "' and cnv_grupo not in ('CCOU','UNIVERSAL','CASA DE GALICIA','H.EVANGELICO','SMI','IMPASA')"
                        data_lla22.Refresh
                        If data_lla22.Recordset.RecordCount > 0 Then
                           If data_lla22.Recordset("cnv_codigo") = "SAMCB" Then
                              Xlabancp = 3
                           Else
                              Xlabancp = 0
                           End If
                        Else
                           Xlabancp = 3
                        End If
                    End If
'                  End If
                                    
                  If Xlabancp = 3 Then
                  Else
                     Xarchexel22.Cells(Xlin, XCol) = Trim(str(Day(data_lla.Recordset("fecha"))))
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = Trim(str(Month(data_lla.Recordset("fecha"))))
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = Trim(str(Year(data_lla.Recordset("fecha"))))
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = data_lla.Recordset("nombre")
                     XCol = XCol + 1
                     If data_lla.Recordset("codzon") = 4 Or data_lla.Recordset("codzon") = 5 Or data_lla.Recordset("codzon") = 6 Then
                        Xarchexel22.Cells(Xlin, XCol) = "Zona:3"
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "Zona:" & Trim(str(data_lla.Recordset("codzon")))
                     End If
'                     XCol = XCol + 1
'                     Xarchexel22.Cells(Xlin, XCol) = data_lla.Recordset("categ")
'                     XCol = XCol + 1
'                     Xarchexel22.Cells(Xlin, XCol) = data_lla.Recordset("nomcat")
                     
                     Xcantre = Xcantre + 1
                     If data_lla.Recordset("categ") = "911" Or data_lla.Recordset("categ") = "911B" Then
                        Contar911 = Contar911 + 1
                     End If
                     Xlin = Xlin + 1
                     Xcolfija = XCol
                     XCol = 1
                  End If
                  Xlabancp = 0
                  data_lla.Recordset.MoveNext
                  pb.Value = pb.Value + 1
               Loop
               Xlin = Xlin + 1
               XCol = 2
               Xarchexel22.Cells(Xlin, XCol) = "TOTAL DE REGISTROS:" & Xcantre
               Xlin = Xlin + 1
               XCol = 2
               Xarchexel22.Cells(Xlin, XCol) = "TOTAL DE REGISTROS 911:" & Trim(str(Contar911))
               
               DoEvents
           
           
           End If
        End If
        
        Xlibexel22.Save
'            Xlibexel.Application
        Xlibexel22.Close
        
        Xobjexel22.Quit
'        Shell frm_menu.data_usuac.Recordset("destino") & "excel.exe " & Xarchtex, vbMaximizedFocus
        If WElusuario = "CONTADURIA" Then
           If Combo1.Text = "LLAMADOS" Then
              If Year(mfh.Text) = 2021 Then
                 If Month(mfh.Text) = 9 Or Month(mfh.Text) = 10 Or Month(mfh.Text) = 11 Or _
                    Month(mfh.Text) = 8 Or Month(mfh.Text) = 7 Then
                    Xarchtex = "C:\planillas\cp\llamados " & Mid(mfh.Text, 4, 2) & Mid(mfh.Text, 7, 4) & ".xls"
                 End If
              End If
           End If
        Else
           If Combo1.Text = "LLAMADOS" Then
              If Year(mfh.Text) = 2018 Then
                 If Month(mfh.Text) = 9 Or Month(mfh.Text) = 10 Or Month(mfh.Text) = 11 Or _
                    Month(mfh.Text) = 12 Then
                    Xarchtex = "C:\planillas\Llamados " & Mid(mfh.Text, 4, 2) & Mid(mfh.Text, 7, 4) & ".xls"
                 End If
              Else
                 If Year(mfh.Text) = 2019 Then
                    If Month(mfh.Text) = 1 Or Month(mfh.Text) = 2 Or Month(mfh.Text) = 3 Or _
                       Month(mfh.Text) = 4 Or Month(mfh.Text) = 5 Or Month(mfh.Text) = 6 Then
                       Xarchtex = "C:\planillas\Llamados " & Mid(mfh.Text, 4, 2) & Mid(mfh.Text, 7, 4) & ".xls"
                    End If
                 End If
              End If
           End If
        End If
        Xlabrir3.Workbooks.Open Xarchtex, , False
        Xlabrir3.Visible = True
        Xlabrir3.WindowState = xlMaximized
        

End Sub

Private Sub Command5_Click()
          If IsNull(data_mdb.Recordset("hllega")) = False And IsNull(data_mdb.Recordset("hor_cance")) = False Then
             
             Xhh1 = Val(Mid(data_mdb.Recordset("hor_cance"), 1, 2))
             Xmm1 = Val(Mid(data_mdb.Recordset("hor_cance"), 4, 2))
         
             Xhh2 = Val(Mid(data_mdb.Recordset("hllega"), 1, 2))
             Xmm2 = Val(Mid(data_mdb.Recordset("hllega"), 4, 2))
          
             Xtoth = Xhh1 - Xhh2
             If Xtoth < 0 Then
                Xtoth = Xtoth + 24
             End If
             Xtotm = Xmm1 - Xmm2
             If Xtotm < 0 Then
                Xtotm = Xtotm + 60
                Xtoth = Xtoth - 1
             End If
             If Xtoth = 1 Then
                Xtotm = Xtotm + 60
             End If
             If Xtoth = 2 Then
                Xtotm = Xtotm + 120
             End If
             If Xtoth = 3 Then
                Xtotm = Xtotm + 180
             End If
             If Xtoth = 4 Then
                Xtotm = Xtotm + 240
             End If
             If Xtoth = 5 Then
                Xtotm = Xtotm + 300
             End If
             If Xtoth = 6 Then
                Xtotm = Xtotm + 360
             End If
             If Xtoth = 7 Then
                Xtotm = Xtotm + 420
             End If
             If Xtoth = 8 Then
                Xtotm = Xtotm + 480
             End If
             
             data_mdb.Recordset.Edit
             data_mdb.Recordset("hh") = Xtotm
             data_mdb.Recordset.Update
          Else
             data_mdb.Recordset.Edit
             data_mdb.Recordset("hh") = 0
             data_mdb.Recordset.Update
    
          End If
          Xtotm = 0
          Xtoth = 0
          Xhh1 = 0
          Xhh2 = 0
          Xmm1 = 0
          Xmm2 = 0

End Sub

Private Sub Command6_Click()
Dim Xobjexel22 As Excel.Application
Dim Xlibexel22 As Excel.Workbook

Dim Xarchexel22 As New Excel.Worksheet
Dim Xqdia, Xcanxdia, Xlin, XCol As Integer
Dim Xtotreg As Long
Dim Xarchtex As String
Dim Textofecha As String
Dim Xtempo As Integer
Dim Cuentactrol As Integer
Dim Xelnrocovid As Integer
Dim Xlabrir3 As New Excel.Application
Dim Xlabancp As Integer
Dim Contar911 As Integer
Cuentactrol = 0
Xlin = 1
XCol = 1
Xelnrocovid = 0
data_buscacov.Connect = "odbc;dsn=sappnew;"

      Set Xobjexel22 = New Excel.Application

      Set Xlibexel22 = Xobjexel22.Workbooks.Add
      Set Xarchexel22 = Xlibexel22.Worksheets.Add

      Xarchexel22.Name = Trim(Combo1.Text)

      Xlibexel22.SaveAs ("C:\planillas\COVID-19.xls")
      Xarchtex = "C:\planillas\COVID-19.xls"
        
        Xqdia = 0
        Xcanxdia = 0
        Xarchexel22.Cells(Xlin, XCol) = "DEPARTAMENTO TI SAPP S.A."
        Xlin = Xlin + 1
        XCol = XCol + 1
        Xarchexel22.Range("A1", "C3").Font.Size = 16
        Xarchexel22.Cells(Xlin, XCol) = "PLANILLA DE SEGUIMIENTO COVID-19  DESDE: " & mfd.Text & " HASTA: " & mfh.Text
        
        XCol = 1
        Xlin = Xlin + 2
        Xnrocan = Xnrocan + Xlin
'        Xarchexel22.Range("A4", "AJ" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
'        Xarchexel22.Range("A4", "AJ" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
'        Xarchexel22.Range("A4", "AJ" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
'        Xarchexel22.Range("A4", "AJ" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
'        Xarchexel22.Range("A4", "AJ" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
'        Xarchexel22.Range("A4", "AJ" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
        Xarchexel22.Range("A" & Trim(str(Xlin)), "AJ" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
        Xarchexel22.Range("A" & Trim(str(Xlin))).ColumnWidth = 5
        Xarchexel22.Cells(Xlin, XCol) = "NRO."
        XCol = XCol + 1
        Xarchexel22.Range("B" & Trim(str(Xlin))).ColumnWidth = 14
        Xarchexel22.Cells(Xlin, XCol) = "FECHA CTROL."
        
        XCol = XCol + 1
        Xarchexel22.Range("C" & Trim(str(Xlin))).ColumnWidth = 14
        Xarchexel22.Cells(Xlin, XCol) = "MEDICO"
        
        XCol = XCol + 1
        Xarchexel22.Range("D" & Trim(str(Xlin))).ColumnWidth = 35
        Xarchexel22.Cells(Xlin, XCol) = "APELLIDO Y NOMBRE"
        
        XCol = XCol + 1
        Xarchexel22.Range("E" & Trim(str(Xlin))).ColumnWidth = 12
        Xarchexel22.Cells(Xlin, XCol) = "CEDULA"
        
        XCol = XCol + 1
        Xarchexel22.Range("F" & Trim(str(Xlin))).ColumnWidth = 5
        Xarchexel22.Cells(Xlin, XCol) = "EDAD"
        XCol = XCol + 1
        Xarchexel22.Range("G" & Trim(str(Xlin))).ColumnWidth = 20
        Xarchexel22.Cells(Xlin, XCol) = "ZONA"
        XCol = XCol + 1
        Xarchexel22.Range("H" & Trim(str(Xlin))).ColumnWidth = 15
        Xarchexel22.Cells(Xlin, XCol) = "MUTUALISTA"
        XCol = XCol + 1
        Xarchexel22.Range("I" & Trim(str(Xlin))).ColumnWidth = 20
        Xarchexel22.Cells(Xlin, XCol) = "TELEFONO"
        XCol = XCol + 1
        Xarchexel22.Range("J" & Trim(str(Xlin))).ColumnWidth = 12
        Xarchexel22.Cells(Xlin, XCol) = "VIAJE"
        XCol = XCol + 1
        Xarchexel22.Range("K" & Trim(str(Xlin))).ColumnWidth = 12
        Xarchexel22.Cells(Xlin, XCol) = "CONTACTO"
        XCol = XCol + 1
        Xarchexel22.Range("L" & Trim(str(Xlin))).ColumnWidth = 15
        Xarchexel22.Cells(Xlin, XCol) = "INICIO SINTOMAS"
        XCol = XCol + 1
        Xarchexel22.Range("M" & Trim(str(Xlin))).ColumnWidth = 25
        Xarchexel22.Cells(Xlin, XCol) = "SINTOMAS"
        XCol = XCol + 1
        Xarchexel22.Range("N" & Trim(str(Xlin))).ColumnWidth = 20
        Xarchexel22.Cells(Xlin, XCol) = "COMUNICACION EPI"
        XCol = XCol + 1
        Xarchexel22.Range("O" & Trim(str(Xlin))).ColumnWidth = 15
        Xarchexel22.Cells(Xlin, XCol) = "CTROL.TELEF"
        XCol = XCol + 1
        Xarchexel22.Range("P" & Trim(str(Xlin))).ColumnWidth = 15
        Xarchexel22.Cells(Xlin, XCol) = "CTROL.MEDICO"
        XCol = XCol + 1
        Xarchexel22.Cells(Xlin, XCol) = "TRASLADO"
        XCol = XCol + 1
        Xarchexel22.Cells(Xlin, XCol) = "FECHA de TEST"
        Xarchexel22.Range("R" & Trim(str(Xlin))).ColumnWidth = 14
        XCol = XCol + 1
        Xarchexel22.Cells(Xlin, XCol) = "RESULTADO"
        Xarchexel22.Range("S" & Trim(str(Xlin))).ColumnWidth = 13
        XCol = XCol + 1
        Xarchexel22.Cells(Xlin, XCol) = "FECHA ALTA"
        Xarchexel22.Range("T" & Trim(str(Xlin))).ColumnWidth = 14

        XCol = XCol + 1
        Xarchexel22.Range("U" & Trim(str(Xlin))).ColumnWidth = 100
        Xarchexel22.Cells(Xlin, XCol) = "CONTROL"
        
        Xlin = Xlin + 1
        XCol = 1
        
        If data_lla.Recordset.RecordCount > 0 Then
           Xelnrocovid = 1
           data_lla.Recordset.MoveLast
           pb.Max = pb.Max + data_lla.Recordset.RecordCount + 1
           data_lla.Recordset.MoveFirst
           Do While Not data_lla.Recordset.EOF
              data_buscacov.RecordSource = "Select * from llamado where nrolla =" & data_lla.Recordset("id_llamado")
              data_buscacov.Refresh
              If data_buscacov.Recordset.RecordCount > 0 Then
                  Xarchexel22.Cells(Xlin, XCol) = Xelnrocovid
                  XCol = XCol + 1
                  Xarchexel22.Cells(Xlin, XCol) = "'" & Trim(Format(data_lla.Recordset("fecha"), "dd/mm/yyyy"))
                  XCol = XCol + 1
                  Xarchexel22.Cells(Xlin, XCol) = data_lla.Recordset("nom_usu")
                  XCol = XCol + 1
                  Xarchexel22.Cells(Xlin, XCol) = data_buscacov.Recordset("nombre")
                  XCol = XCol + 1
                  If IsNull(data_buscacov.Recordset("ci")) = False Then
                     data_lla22.RecordSource = "Select * from resplla where nro =" & data_lla.Recordset("id_llamado")
                     data_lla22.Refresh
                     If data_lla22.Recordset.RecordCount > 0 Then
                        If IsNull(data_lla22.Recordset("mes")) = False Then
                           Xarchexel22.Cells(Xlin, XCol) = Trim(str(data_buscacov.Recordset("ci"))) & "-" & Trim(str(data_lla22.Recordset("mes")))
                        Else
                           Xarchexel22.Cells(Xlin, XCol) = Trim(str(data_buscacov.Recordset("ci"))) & "-0"
                        End If
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = Trim(str(data_buscacov.Recordset("ci"))) & "-0"
                     End If
                  Else
                     Xarchexel22.Cells(Xlin, XCol) = Trim(str(data_buscacov.Recordset("ci"))) & "-0"
                  End If
                  XCol = XCol + 1
                  Xarchexel22.Cells(Xlin, XCol) = data_buscacov.Recordset("edad")
                  XCol = XCol + 1
                  If IsNull(data_buscacov.Recordset("motmov")) = False Then
                     Xarchexel22.Cells(Xlin, XCol) = data_buscacov.Recordset("motmov")
                  Else
                     Xarchexel22.Cells(Xlin, XCol) = "Sin Datos"
                  End If
                  XCol = XCol + 1
                  data_covid.RecordSource = "select * from convenio where cnv_codigo ='" & data_buscacov.Recordset("categ") & "'"
                  data_covid.Refresh
                  If data_covid.Recordset.RecordCount > 0 Then
                     If IsNull(data_covid.Recordset("cnv_grupo")) = False Then
                        If data_covid.Recordset("cnv_grupo") <> "" Then
                           Xarchexel22.Cells(Xlin, XCol) = data_covid.Recordset("cnv_grupo")
                        Else
                           Xarchexel22.Cells(Xlin, XCol) = "SAPP"
                        End If
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "SAPP"
                     End If
                  Else
                     Xarchexel22.Cells(Xlin, XCol) = "SAPP"
                  End If
                  XCol = XCol + 1
                  Xarchexel22.Cells(Xlin, XCol) = data_buscacov.Recordset("telef")
                  XCol = XCol + 1
                  If IsNull(data_buscacov.Recordset("viaje")) = False Then
                     Xtempo = data_buscacov.Recordset("viaje")
                     If Xtempo = 1 Then
                        Xarchexel22.Cells(Xlin, XCol) = "SI"
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                  Else
                     Xarchexel22.Cells(Xlin, XCol) = "Sin Datos"
                  End If
                  XCol = XCol + 1
                  If IsNull(data_buscacov.Recordset("contacto")) = False Then
                     Xarchexel22.Cells(Xlin, XCol) = data_buscacov.Recordset("contacto")
                  Else
                     Xarchexel22.Cells(Xlin, XCol) = "Sin Datos"
                  End If
                  XCol = XCol + 1
                  If IsNull(data_buscacov.Recordset("inicio_sint")) = False Then
                     Xarchexel22.Cells(Xlin, XCol) = data_buscacov.Recordset("inicio_sint")
                  Else
                     Xarchexel22.Cells(Xlin, XCol) = "Sin Datos"
                  End If
                  XCol = XCol + 1
                  If IsNull(data_buscacov.Recordset("sintomas")) = False Then
                     Xarchexel22.Cells(Xlin, XCol) = data_buscacov.Recordset("sintomas")
                  Else
                     Xarchexel22.Cells(Xlin, XCol) = "Sin Datos"
                  End If
                  XCol = XCol + 1
                  If IsNull(data_buscacov.Recordset("comunic_epi")) = False Then
                     Xtempo = data_buscacov.Recordset("comunic_epi")
                     If Xtempo = 1 Then
                        Xarchexel22.Cells(Xlin, XCol) = "SI"
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                  Else
                     Xarchexel22.Cells(Xlin, XCol) = "Sin Datos"
                  End If
                  XCol = XCol + 1
                  If IsNull(data_buscacov.Recordset("ctrol_telef")) = False Then
                     Xtempo = data_buscacov.Recordset("ctrol_telef")
                     If Xtempo = 1 Then
                        Xarchexel22.Cells(Xlin, XCol) = "SI"
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                  Else
                     Xarchexel22.Cells(Xlin, XCol) = "Sin Datos"
                  End If
                  XCol = XCol + 1
                  If IsNull(data_buscacov.Recordset("ctrol_medic")) = False Then
                     Xtempo = data_buscacov.Recordset("ctrol_medic")
                     If Xtempo = 1 Then
                        Xarchexel22.Cells(Xlin, XCol) = "SI"
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                  Else
                     Xarchexel22.Cells(Xlin, XCol) = "Sin Datos"
                  End If
                  XCol = XCol + 1
                  If IsNull(data_buscacov.Recordset("trasla")) = False Then
                     If data_buscacov.Recordset("trasla") >= 0 Then
                        Xarchexel22.Cells(Xlin, XCol) = "SI"
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "NO"
                     End If
                  Else
                     Xarchexel22.Cells(Xlin, XCol) = "Sin Datos"
                  End If
                  XCol = XCol + 1
                  If IsNull(data_buscacov.Recordset("isopa_fecrea")) = False Then
                     Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_buscacov.Recordset("isopa_fecrea"), "dd/mm/yyyy")
                  End If
                  XCol = XCol + 1
                  If IsNull(data_buscacov.Recordset("isopa_result")) = False Then
                     Xarchexel22.Cells(Xlin, XCol) = data_buscacov.Recordset("isopa_result")
                  End If
                  XCol = XCol + 1
                  If IsNull(data_buscacov.Recordset("cierre_fec")) = False Then
                     Xarchexel22.Cells(Xlin, XCol) = "'" & Format(data_buscacov.Recordset("cierre_fec"), "dd/mm/yyyy")
                  End If
                  XCol = XCol + 1
                  If IsNull(data_lla.Recordset("texto")) = False Then
                     Xarchexel22.Cells(Xlin, XCol) = "Ctrol.NRO:" & Trim(str(data_lla.Recordset("dia"))) & " " & data_lla.Recordset("nom_usu") & ": " & data_lla.Recordset("texto")
                  Else
                     Xarchexel22.Cells(Xlin, XCol) = "Sin Datos"
                  End If
                  
                  Cuentactrol = 0
                  Xlin = Xlin + 1
                  XCol = 1
                  Xelnrocovid = Xelnrocovid + 1
              End If
              data_lla.Recordset.MoveNext
              pb.Value = pb.Value + 1
           Loop
           Xlin = Xlin + 1
           
           Xarchexel22.Cells(Xlin, XCol) = "FECHA DE EMISION:" & Format(Date, "dd/mm/yyyy")
        
        End If
        
        Xlibexel22.Save
'            Xlibexel.Application
        Xlibexel22.Close
        
        Xobjexel22.Quit
'        Shell frm_menu.data_usuac.Recordset("destino") & "excel.exe " & Xarchtex, vbMaximizedFocus
        Xlabrir3.Workbooks.Open Xarchtex, , False
        Xlabrir3.Visible = True
        Xlabrir3.WindowState = xlMaximized


End Sub

Private Sub Command7_Click()
Dim Xobjexel22 As Excel.Application
Dim Xlibexel22 As Excel.Workbook

Dim Xarchexel22 As New Excel.Worksheet
Dim Xqdia, Xcanxdia, Xlin, XCol As Integer
Dim Xtotreg As Long
Dim Xarchtex As String
Dim Textofecha As String

Dim Xlabrir3 As New Excel.Application
Dim Xlabancp As Integer
Dim Contar911 As Integer

Xlin = 1
XCol = 1

      Set Xobjexel22 = New Excel.Application

      Set Xlibexel22 = Xobjexel22.Workbooks.Add
      Set Xarchexel22 = Xlibexel22.Worksheets.Add

      Xarchexel22.Name = Trim(Combo1.Text)

      Xlibexel22.SaveAs ("C:\planillas\" & Trim(Combo1.Text) & "\" & Trim(str(Month(mfd.Text))) & Trim(str(Year(mfd.Text))) & ".xls")
      Xarchtex = "C:\planillas\" & Trim(Combo1.Text) & "\" & Trim(str(Month(mfd.Text))) & Trim(str(Year(mfd.Text))) & ".xls"
        
        Xqdia = 0
        Xcanxdia = 0
        Xarchexel22.Cells(Xlin, XCol) = "DEPARTAMENTO TI SAPP S.A."
        Xlin = Xlin + 1
        XCol = XCol + 1
        Xarchexel22.Range("A1", "C3").Font.Size = 16
        Xarchexel22.Cells(Xlin, XCol) = "PLANILLA DE " & Trim(Combo1.Text) & " DESDE: " & mfd.Text & " HASTA: " & mfh.Text
        Xarchexel22.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
        
        XCol = 1
        Xlin = Xlin + 2
        Xnrocan = Xnrocan + Xlin
        Xarchexel22.Range("A" & Trim(str(Xlin)), "E" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
        Xarchexel22.Range("A" & Trim(str(Xlin))).ColumnWidth = 6
        Xarchexel22.Cells(Xlin, XCol) = "DIA"
        XCol = XCol + 1
        Xarchexel22.Range("B" & Trim(str(Xlin))).ColumnWidth = 6
        Xarchexel22.Cells(Xlin, XCol) = "MES"
        XCol = XCol + 1
        Xarchexel22.Range("C" & Trim(str(Xlin))).ColumnWidth = 6
        Xarchexel22.Cells(Xlin, XCol) = "AÑO"
        XCol = XCol + 1
        Xarchexel22.Range("D" & Trim(str(Xlin))).ColumnWidth = 35
        Xarchexel22.Cells(Xlin, XCol) = "NOMBRE"
        XCol = XCol + 1
        Xarchexel22.Range("E" & Trim(str(Xlin))).ColumnWidth = 12
        Xarchexel22.Cells(Xlin, XCol) = "ZONA"
        Xlin = Xlin + 1
        XCol = 1
        If data_lla.Recordset.RecordCount > 0 Then
           data_lla.Recordset.MoveLast
           pb.Max = pb.Max + data_lla.Recordset.RecordCount + 1
           data_lla.Recordset.MoveFirst
           Xqdia = Day(data_lla.Recordset("fecha"))
           Dim Xcantre As Integer
           Xcantre = 0
           If Combo1.ListIndex = 2 Then 'traslados
               Do While Not data_lla.Recordset.EOF
'                  If data_lla.Recordset("categ") = "911" Or data_lla.Recordset("categ") = "911B" Then
                     Xlabancp = 0
'                  Else
'                     data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lla.Recordset("categ") & "'"
'                     data_conv.Refresh
'                     If data_conv.Recordset.RecordCount > 0 Then
'                        If IsNull(data_conv.Recordset("cnv_grupo")) = False Then
'                           If data_conv.Recordset("cnv_grupo") = "CCOU" Or _
'                              data_conv.Recordset("cnv_grupo") = "SMI" Or _
'                              data_conv.Recordset("cnv_grupo") = "UNIVERSAL" Or _
'                              data_conv.Recordset("cnv_grupo") = "IMPASA" Or _
'                              data_conv.Recordset("cnv_codigo") = "UNIVA" Or data_conv.Recordset("cnv_codigo") = "CGALIC" Or _
'                              data_conv.Recordset("cnv_grupo") = "CASA DE GALICIA" Or _
'                              data_conv.Recordset("cnv_codigo") = "IMP2" Or data_conv.Recordset("cnv_codigo") = "CASA3" Or _
'                              data_conv.Recordset("cnv_codigo") = "SMI3" Or data_conv.Recordset("cnv_codigo") = "SMI4" Or _
'                              data_conv.Recordset("cnv_codigo") = "CCSA" Or data_conv.Recordset("cnv_codigo") = "SMI5" Or _
'                              data_conv.Recordset("cnv_grupo") = "H.EVANGELICO" Then
'                           Else
'                              Xlabancp = 3
'                           End If
'                        Else
'                           Xlabancp = 3
'                        End If
'                     Else
'                        Xlabancp = 3
'                     End If
'                  End If
                  If Xlabancp = 3 Then
                  Else
                     Xarchexel22.Cells(Xlin, XCol) = Trim(str(Day(data_lla.Recordset("fecha"))))
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = Trim(str(Month(data_lla.Recordset("fecha"))))
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = Trim(str(Year(data_lla.Recordset("fecha"))))
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = data_lla.Recordset("nombre")
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = "Zona:3"
                     Xcantre = Xcantre + 1
                     Xlin = Xlin + 1
                     Xcolfija = XCol
                     XCol = 1
                  End If
                  Xlabancp = 0
                  data_lla.Recordset.MoveNext
                  pb.Value = pb.Value + 1
               Loop
               Xlin = Xlin + 1
               XCol = 2
               Xarchexel22.Cells(Xlin, XCol) = "TOTAL DE REGISTROS:" & Xcantre
               DoEvents
           
           Else
               Contar911 = 0
               Do While Not data_lla.Recordset.EOF
                  If data_lla.Recordset("categ") = "911" Or data_lla.Recordset("categ") = "911B" Or data_lla.Recordset("categ") = "SEMM1" Or _
                     data_lla.Recordset("categ") = "CERSEM" Or data_lla.Recordset("categ") = "UDEMM" Or data_lla.Recordset("categ") = "SEMM" Or _
                     data_lla.Recordset("categ") = "UCM" Then
                     Xlabancp = 0
                  Else
                     If data_lla.Recordset("categ") = "SJ01" Or data_lla.Recordset("categ") = "SJ02" Or data_lla.Recordset("categ") = "CASH" Then
                        Xlabancp = 0
                     Else
                        data_lla22.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lla.Recordset("categ") & "' and cnv_emite ='" & "SI" & "' and (cnv_cant_r =" & 2 & " or cnv_cant_r=" & 1 & ") and (cnv_grupo is null or cnv_grupo='" & "" & "')"
                        data_lla22.Refresh
                        If data_lla22.Recordset.RecordCount > 0 Then
                           Xlabancp = 0
                        Else
'                           Xlabancp = 3
                        End If
                     End If
                  End If
                                    
                  If Xlabancp = 3 Then
                  Else
                     Xarchexel22.Cells(Xlin, XCol) = Trim(str(Day(data_lla.Recordset("fecha"))))
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = Trim(str(Month(data_lla.Recordset("fecha"))))
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = Trim(str(Year(data_lla.Recordset("fecha"))))
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = data_lla.Recordset("nombre")
                     XCol = XCol + 1
                     If data_lla.Recordset("codzon") = 4 Or data_lla.Recordset("codzon") = 5 Or data_lla.Recordset("codzon") = 6 Then
                        Xarchexel22.Cells(Xlin, XCol) = "Zona:5"
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "Zona:" & Trim(str(data_lla.Recordset("codzon")))
                     End If
                     Xcantre = Xcantre + 1
                     If data_lla.Recordset("categ") = "911" Or data_lla.Recordset("categ") = "911B" Then
                        Contar911 = Contar911 + 1
                     End If
                     Xlin = Xlin + 1
                     Xcolfija = XCol
                     XCol = 1
                  End If
                  Xlabancp = 0
                  data_lla.Recordset.MoveNext
                  pb.Value = pb.Value + 1
               Loop
               Xlin = Xlin + 1
               XCol = 2
               Xarchexel22.Cells(Xlin, XCol) = "TOTAL DE REGISTROS:" & Xcantre
               Xlin = Xlin + 1
               XCol = 2
               Xarchexel22.Cells(Xlin, XCol) = "TOTAL DE REGISTROS 911:" & Trim(str(Contar911))
               
               DoEvents
           
           
           End If
        End If
        
        Xlibexel22.Save
'            Xlibexel.Application
        Xlibexel22.Close
        
        Xobjexel22.Quit
'        Shell frm_menu.data_usuac.Recordset("destino") & "excel.exe " & Xarchtex, vbMaximizedFocus
        If Combo1.Text = "LLAMADOS" Then
            If Year(mfh.Text) = 2018 Then
               If Month(mfh.Text) = 9 Or Month(mfh.Text) = 10 Or Month(mfh.Text) = 11 Or _
                  Month(mfh.Text) = 12 Then
                  Xarchtex = "C:\planillas\Llamados " & Mid(mfh.Text, 4, 2) & Mid(mfh.Text, 7, 4) & ".xls"
               End If
            Else
               If Year(mfh.Text) = 2019 Then
                  If Month(mfh.Text) = 1 Or Month(mfh.Text) = 2 Or Month(mfh.Text) = 3 Or _
                     Month(mfh.Text) = 4 Or Month(mfh.Text) = 5 Or Month(mfh.Text) = 6 Then
                     Xarchtex = "C:\planillas\Llamados " & Mid(mfh.Text, 4, 2) & Mid(mfh.Text, 7, 4) & ".xls"
                  End If
               End If
            End If
        End If
        Xlabrir3.Workbooks.Open Xarchtex, , False
        Xlabrir3.Visible = True
        Xlabrir3.WindowState = xlMaximized

End Sub

Private Sub Form_Load()
If Check3.Value = 1 Then
'   data_lla.Connect = "odbc;dsn=" & Xconexrmt & ";"
'   data_lla.DatabaseName = App.Path & "\llamado.mdb"
   data_lla.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.path & "\llamado.mdb"
Else
'   data_lla.Connect = "odbc;dsn=" & Xconexrmt & ";"
   data_lla.ConnectionString = "dsn=" & Xconexrmt
   data_lla22.ConnectionString = "dsn=" & Xconexrmt
'   data_lla22.Connect = "odbc;dsn=" & Xconexrmt & ";"
End If
'data_inf.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"
data_inf.DatabaseName = App.path & "\informes.mdb"
data_movil.DatabaseName = App.path & "\moviles.mdb"
data_movil.RecordSource = "movil"
data_movil.Refresh
data_conv.ConnectionString = "dsn=" & Xconexrmt
Check1.Value = 1
data_moviles.ConnectionString = "dsn=" & Xconexrmt
data_covid.Connect = "odbc;dsn=sappnew;"
'data_moviles.RecordSource = "movil"
'data_moviles.Refresh

'data_conv.RecordSource = "convenio"
'data_conv.Refresh

data_med.ConnectionString = "dsn=" & Xconexrmt
'data_med.RecordSource = "medicos"
'data_med.Refresh
data_llaenf.Connect = "odbc;dsn=" & Xconexrmt & ";"

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
   Combo1.SetFocus
End If

End Sub
