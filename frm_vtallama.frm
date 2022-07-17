VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_vtallama 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de llamados con costo"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6270
   Icon            =   "frm_vtallama.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6270
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data_lla 
      Height          =   330
      Left            =   4200
      Top             =   3000
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
   Begin MSComctlLib.ProgressBar pb 
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   5040
      Top             =   2160
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
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   3255
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
      Left            =   5640
      Picture         =   "frm_vtallama.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   3000
      Width           =   495
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
      Left            =   120
      Picture         =   "frm_vtallama.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Procesar"
      Top             =   3000
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
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
      Height          =   2295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6015
      Begin MSAdodcLib.Adodc data_lla2 
         Height          =   375
         Left            =   1800
         Top             =   1560
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
         Caption         =   "data_lla2"
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
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C00000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox txt_mov 
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
         Height          =   375
         Left            =   2280
         TabIndex        =   5
         Text            =   "999"
         Top             =   1080
         Width           =   1095
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   4080
         TabIndex        =   3
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
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
         Left            =   2280
         TabIndex        =   2
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
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
         BackColor       =   &H00C00000&
         Caption         =   "Móvil (999=TODOS)"
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
         TabIndex        =   4
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Rango de fechas:"
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
         Width           =   2175
      End
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   2040
      Picture         =   "frm_vtallama.frx":0F56
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   1935
   End
End
Attribute VB_Name = "frm_vtallama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frm_vtallama.MousePointer = 11
Command1.Enabled = False

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infvtas"
data_inf.RecordSource = "infvtas"
data_inf.Refresh

If md.Text = "__/__/____" Or mh.Text = "__/__/____" Then
   MsgBox "No ingresó fechas"
Else
   If txt_mov.Text = 999 Then
      If Check1.Value = 1 Then
         data_lla.RecordSource = "Select * from resplla where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and mes >" & 50 & " order by fecha"
         data_lla.Refresh
      Else
         data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and mes >" & 50 & " and movilpas not in (99,0) order by fecha"
         data_lla.Refresh
      End If
   Else
      If Check1.Value = 1 Then
         data_lla.RecordSource = "Select * from resplla where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and mes >" & 50 & " order by fecha"
         data_lla.Refresh
      Else
         data_lla.RecordSource = "Select * from llamado where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and mes >" & 50 & " and movilpas not in (99,0) order by fecha"
         data_lla.Refresh
      End If
   End If
   If data_lla.Recordset.RecordCount > 0 Then
      data_lla.Recordset.MoveLast
      pb.Max = data_lla.Recordset.RecordCount
      data_lla.Recordset.MoveFirst
      DoEvents
      Do While Not data_lla.Recordset.EOF
         data_inf.Recordset.AddNew
         data_inf.Recordset("fecha") = data_lla.Recordset("fecha")
         data_inf.Recordset("cod_cli") = data_lla.Recordset("matric")
         data_inf.Recordset("nom_cli") = Mid(data_lla.Recordset("nombre"), 1, 30)
         data_inf.Recordset("convenio") = data_lla.Recordset("categ")
         data_inf.Recordset("base") = data_lla.Recordset("movilpas")
         data_inf.Recordset("nro_med_a") = data_lla.Recordset("codmed")
         data_inf.Recordset("nom_med_a") = Mid(data_lla.Recordset("nommed"), 1, 40)
         data_inf.Recordset("tot_lin") = data_lla.Recordset("mes")
         data_inf.Recordset("fact") = data_lla.Recordset("ano")
         data_lla2.RecordSource = "Select * from resplla where nro =" & data_lla.Recordset("nrolla")
         data_lla2.Refresh
         If data_lla2.Recordset.RecordCount > 0 Then
            data_inf.Recordset("tipo") = data_lla2.Recordset("telef")
         Else
            data_inf.Recordset("tipo") = "N/R"
         End If
         data_inf.Recordset("hora") = data_lla.Recordset("codmot")
         data_inf.Recordset("nom_superv") = data_lla.Recordset("timdes")
         data_inf.Recordset("nom_flia") = data_lla.Recordset("usuario")
         data_inf.Recordset.Update
         data_lla.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      data_inf.RecordSource = "Select * from infvtas order by fecha"
      data_inf.Refresh
      frm_vtallama.MousePointer = 0
      MsgBox "Proceso terminado"
      cr1.ReportFileName = App.path & "\infvtalla.rpt"
      cr1.ReportTitle = "INFORME DE LLAMADOS CON COSTO DESDE: " & md.Text & " HASTA: " & mh.Text
      cr1.Action = 1
      
   End If
End If
Command1.Enabled = True
      
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
data_inf.DatabaseName = App.path & "\informes.mdb"
data_lla.ConnectionString = "dsn=" & Xconexrmt
data_lla2.ConnectionString = "dsn=" & Xconexrmt

If Xaltaaa = 8 Then
   md.Text = frm_infvtascre.mfd.Text
   mh.Text = frm_infvtascre.mfh.Text
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
   Command1.SetFocus
End If

End Sub
