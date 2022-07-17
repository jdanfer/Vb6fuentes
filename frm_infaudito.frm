VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_infaudito 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de solicitudes de auditoría"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6060
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_infaudito.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6060
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_busca 
      Caption         =   "data_busca"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3480
      Visible         =   0   'False
      Width           =   2535
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   5640
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_acc 
      Caption         =   "data_acc"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5400
      Picture         =   "frm_infaudito.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir"
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      Picture         =   "frm_infaudito.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Procesar"
      Top             =   3360
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Datos de informe"
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5775
      Begin VB.OptionButton Option2 
         BackColor       =   &H00800000&
         Caption         =   "Resumen"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   2880
         TabIndex        =   7
         Top             =   2160
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00800000&
         Caption         =   "Detalle"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   390
         ItemData        =   "frm_infaudito.frx":0F56
         Left            =   1800
         List            =   "frm_infaudito.frx":0F63
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1200
         Width           =   3375
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   3600
         TabIndex        =   3
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
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
         Left            =   1800
         TabIndex        =   2
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
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
         BackColor       =   &H00800000&
         Caption         =   "Informe:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "Fechas:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   3240
      Picture         =   "frm_infaudito.frx":0FAC
      Stretch         =   -1  'True
      Top             =   3480
      Width           =   1455
   End
End
Attribute VB_Name = "frm_infaudito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command1.Enabled = False
Command2.Enabled = False
frm_infaudito.MousePointer = 11
If md.Text <> "__/__/____" Then
   If mh.Text <> "__/__/____" Then
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         Do While Not data_inf.Recordset.EOF
            data_inf.Recordset.Delete
            data_inf.Recordset.MoveNext
         Loop
      End If
      If Combo1.ListIndex = 0 Then
         data_acc.RecordSource = "Select * from env_soc where cl_codigo >=" & 80000 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# "
         data_acc.Refresh
      Else
         If Combo1.ListIndex = 1 Then
            data_acc.RecordSource = "Select * from env_soc where cl_codigo >=" & 80000 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cl_tipocli =" & 0
            data_acc.Refresh
         Else
            If Combo1.ListIndex = 2 Then
               data_acc.RecordSource = "Select * from env_soc where cl_codigo >=" & 80000 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cl_tipocli =" & 1
               data_acc.Refresh
            Else
               data_acc.RecordSource = "Select * from env_soc where cl_codigo >=" & 80000 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "#"
               data_acc.Refresh
            End If
         End If
      End If
      If data_acc.Recordset.RecordCount > 0 Then
         data_acc.Recordset.MoveFirst
         Do While Not data_acc.Recordset.EOF
            data_inf.Recordset.AddNew
            data_inf.Recordset("cl_codigo") = data_acc.Recordset("cl_codigo")
            data_inf.Recordset("cl_fnac") = data_acc.Recordset("cl_fnac")
            data_inf.Recordset("cl_nomvend") = data_acc.Recordset("cl_nomvend")
            data_inf.Recordset("info_debit") = data_acc.Recordset("info_debit")
            If IsNull(data_acc.Recordset("cl_tipocli")) = False Then
               If data_acc.Recordset("cl_tipocli") = 1 Then
                  data_inf.Recordset("cl_zona") = "CONFORME"
                  data_inf.Recordset("cl_dpto") = "CONFORME"
                  data_inf.Recordset("estado") = 1
               Else
                  If data_acc.Recordset("cl_tipocli") = 2 Then
                     data_inf.Recordset("cl_zona") = "NO CONFORME"
                     data_inf.Recordset("cl_dpto") = "NO CONFORME"
                     data_inf.Recordset("estado") = 2
                  
                  Else
                     data_inf.Recordset("cl_zona") = "EN PROCESO"
                     data_inf.Recordset("estado") = 0
                  End If
               End If
            End If
            data_inf.Recordset.Update
            data_acc.Recordset.MoveNext
         Loop
         frm_infmejoras.MousePointer = 0
         MsgBox "Proceso terminado"
         
         data_inf.RecordSource = "Select * from infcli order by cl_fnac"
         data_inf.Refresh
         If Combo1.ListIndex = 0 Then
            cr1.ReportFileName = App.Path & "\infmejor4.rpt"
            cr1.ReportTitle = "INFORME TOTAL de SOLICITUDES DESDE: " & md.Text & " HASTA: " & mh.Text
            cr1.Action = 1
         End If
         If Combo1.ListIndex = 1 Then
            cr1.ReportFileName = App.Path & "\infmejor4.rpt"
            cr1.ReportTitle = "INFORME SOLICITUDES PENDIENTES DESDE: " & md.Text & " HASTA: " & mh.Text
            cr1.Action = 1
         End If
         If Combo1.ListIndex = 2 Then
            cr1.ReportFileName = App.Path & "\infmejor4.rpt"
            cr1.ReportTitle = "INFORME SOLICITUDES CUMPLIDAS DESDE: " & md.Text & " HASTA: " & mh.Text
            cr1.Action = 1
         End If
      Else
         frm_infaudito.MousePointer = 0
         MsgBox "No existen registros"
      End If
   End If
End If
   
Command1.Enabled = True
Command2.Enabled = True
frm_infaudito.MousePointer = 0
   
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
data_acc.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_inf.DatabaseName = App.Path & "\informes.mdb"
data_inf.RecordSource = "infcli"
data_inf.Refresh
data_busca.Connect = "odbc;dsn=" & Xconexrmt & ";"
Combo1.ListIndex = 0

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
   Combo1.SetFocus
End If

End Sub
