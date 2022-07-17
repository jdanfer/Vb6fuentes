VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infabmpromo 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe de socios por promotores"
   ClientHeight    =   4230
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7605
   Icon            =   "frm_infabmpromo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   7605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc data_prom 
      Height          =   375
      Left            =   3840
      Top             =   3840
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
      Caption         =   "data_prom"
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
   Begin MSAdodcLib.Adodc data_cli 
      Height          =   375
      Left            =   1320
      Top             =   3480
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
   Begin Crystal.CrystalReport cr1 
      Left            =   3240
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton bfin 
      BackColor       =   &H00FFFFFF&
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
      Left            =   6480
      Picture         =   "frm_infabmpromo.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir"
      Top             =   3360
      Width           =   615
   End
   Begin VB.CommandButton bproc 
      BackColor       =   &H00FFFFFF&
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
      Picture         =   "frm_infabmpromo.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Procesar"
      Top             =   3360
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Datos de informe:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3015
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6735
      Begin VB.ComboBox DBCombo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1080
         Width           =   3975
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Numérico"
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
         Left            =   4440
         TabIndex        =   7
         Top             =   2280
         Width           =   1935
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Detallado"
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
         Left            =   2040
         TabIndex        =   6
         Top             =   2280
         Value           =   -1  'True
         Width           =   2055
      End
      Begin MSMask.MaskEdBox mfh 
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
         Left            =   4320
         TabIndex        =   3
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
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
      Begin MSMask.MaskEdBox mfd 
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
         Left            =   2160
         TabIndex        =   2
         Top             =   360
         Width           =   1815
         _ExtentX        =   3201
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
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
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
         TabIndex        =   10
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "Tipo de Informe:"
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
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   0
         X2              =   6720
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Promotor:"
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
         TabIndex        =   4
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   120
      Picture         =   "frm_infabmpromo.frx":0F56
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   1335
   End
End
Attribute VB_Name = "frm_infabmpromo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub bfin_Click()

Unload Me

End Sub

Private Sub bproc_Click()
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.Path & "\informes.mdb")

MiBaseact.Execute "Delete * from infcli"

If mfd.Text <> "__/__/____" Then
   If mfh.Text <> "__/__/____" Then
      frm_infabmpromo.MousePointer = 11
      If DBCombo1.Text <> "*TODOS" Then
         data_cli.RecordSource = "Select * from clientes where cl_fecing >= '" & Format(mfd.Text, "yyyy-mm-dd") & "' And cl_fecing <= '" & Format(mfh.Text, "yyyy-mm-dd") & "' And cl_nrovend =" & Label4.Caption & " order by cl_nrovend"
         data_cli.Refresh
      Else
         data_cli.RecordSource = "Select * from clientes where cl_fecing >= '" & Format(mfd.Text, "yyyy-mm-dd") & "' And cl_fecing <= '" & Format(mfh.Text, "yyyy-mm-dd") & "' order by cl_nrovend"
         data_cli.Refresh
      End If
      If data_cli.Recordset.RecordCount > 0 Then
         data_cli.Recordset.MoveFirst
         Do While Not data_cli.Recordset.EOF
            data_inf.Recordset.AddNew
            data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
            data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
            data_inf.Recordset("cl_nrovend") = data_cli.Recordset("cl_nrovend")
            data_inf.Recordset("cl_nomvend") = data_cli.Recordset("cl_nomvend")
            data_inf.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
            data_inf.Recordset("cl_nomcobr") = data_cli.Recordset("cl_nomcobr")
            data_inf.Recordset("saldo_cc") = data_cli.Recordset("saldo_cc")
            data_inf.Recordset("cl_ultmesp") = data_cli.Recordset("cl_ultmesp")
            data_inf.Recordset("cl_ultanop") = data_cli.Recordset("cl_ultanop")
            data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
            data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
            data_inf.Recordset("fecha_baja") = data_cli.Recordset("fecha_baja")
            data_inf.Recordset.Update
            data_cli.Recordset.MoveNext
         Loop
         frm_infabmpromo.MousePointer = 0
         MsgBox "Proceso terminado", vbInformation, "Mensaje"
         If Option1.value = True Then
            cr1.ReportFileName = App.Path & "\infabmpromo.rpt"
         Else
            cr1.ReportFileName = App.Path & "\infabmpromon.rpt"
         End If
         cr1.ReportTitle = "PERIODO DE INFORME: " + Format(mfd.Text, "dd/mm/yyyy") + " HASTA: " + Format(mfh.Text, "dd/mm/yyyy")
         cr1.Action = 1
      Else
         frm_infabmpromo.MousePointer = 0
         MsgBox "No existen registros para ésta selección", vbInformation, "Mensaje"
      End If
      frm_infabmpromo.MousePointer = 0
   End If
End If

End Sub

Private Sub DBCombo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   bproc.SetFocus
End If

End Sub

Private Sub DBCombo1_LostFocus()
Dim Xtex As String
Xtex = DBCombo1.Text
If Trim(Xtex) <> "" Then
'   data_prom.Recordset.FindFirst "vn_nombre ='" & DBCombo1.Text & "'"
   data_prom.RecordSource = "Select * from vendedor where vn_nombre ='" & Xtex & "'"
   data_prom.Refresh
   If data_prom.Recordset.RecordCount > 0 Then
      Label4.Caption = data_prom.Recordset("vn_numero")
   Else
      Label4.Caption = "799"
   End If
Else
   Label4.Caption = "799"
End If

End Sub

Private Sub Form_Load()
'data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cli.ConnectionString = "dsn=" & Xconexrmt
data_inf.DatabaseName = App.Path & "\informes.mdb"
'data_inf.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"
data_inf.RecordSource = "infcli"
data_inf.Refresh
data_prom.ConnectionString = "dsn=" & Xconexrmt
data_prom.RecordSource = "select * from vendedor order by vn_nombre"
data_prom.Refresh
If data_prom.Recordset.RecordCount > 0 Then
   data_prom.Recordset.MoveFirst
   Do While Not data_prom.Recordset.EOF
      DBCombo1.AddItem data_prom.Recordset("vn_nombre")
      data_prom.Recordset.MoveNext
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

Private Sub mfd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfh.SetFocus
End If

End Sub

Private Sub mfh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   DBCombo1.SetFocus
End If

End Sub
