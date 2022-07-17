VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infestud 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de estudios y servicios con el costo correspondiente"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6570
   Icon            =   "frm_infestud.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   6570
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data_est 
      Height          =   375
      Left            =   2400
      Top             =   1200
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
      Caption         =   "data_est"
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
      RecordSource    =   ""
      Top             =   1800
      Visible         =   0   'False
      Width           =   2820
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   3600
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
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
      Left            =   5760
      Picture         =   "frm_infestud.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command1 
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
      Picture         =   "frm_infestud.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Procesar"
      Top             =   1800
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "frm_infestud.frx":0F56
      Left            =   960
      List            =   "frm_infestud.frx":0F8A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   4215
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   0
      X2              =   6840
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   0
      X2              =   6840
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Seleccione Familia del servicio:"
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
      Left            =   960
      TabIndex        =   0
      Top             =   360
      Width           =   4215
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   240
      Picture         =   "frm_infestud.frx":106D
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1215
   End
End
Attribute VB_Name = "frm_infestud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command1.Enabled = False
Command2.Enabled = False
frm_infestud.MousePointer = 11
If Combo1.ListIndex = 0 Then
   data_est.RecordSource = "estudios"
   data_est.Refresh
Else
   data_est.RecordSource = "Select * from estudios where flia =" & Combo1.ListIndex
   data_est.Refresh
End If

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.Path & "\informes.mdb")

MiBaseact.Execute "Delete * from infcli"

data_inf.RecordSource = "infcli"
data_inf.Refresh

If data_est.Recordset.RecordCount > 0 Then
   data_est.Recordset.MoveFirst
   Do While Not data_est.Recordset.EOF
      data_inf.Recordset.AddNew
      data_inf.Recordset("cl_codigo") = data_est.Recordset("codest")
      data_inf.Recordset("cl_apellid") = data_est.Recordset("descrip")
      data_inf.Recordset("saldo_cc") = data_est.Recordset("cons")
      data_inf.Recordset("saldo_cc2") = data_est.Recordset("uc")
      data_inf.Recordset("cl_cedula") = data_est.Recordset("part")
      data_inf.Recordset("saldo_doc") = data_est.Recordset("ucfh")
      data_inf.Recordset.Update
      data_est.Recordset.MoveNext
   Loop
End If

frm_infestud.MousePointer = 0
Command1.Enabled = True
Command2.Enabled = True
data_inf.RecordSource = "Select * from infcli"
data_inf.Refresh

cr1.ReportFileName = App.Path & "\infestud.rpt"
cr1.Action = 1


End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
data_est.ConnectionString = "dsn=" & Xconexrmt
data_inf.DatabaseName = App.Path & "\informes.mdb"

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub
