VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infmejoras 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de solicitudes de mejoras"
   ClientHeight    =   6645
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
   Icon            =   "frm_infmejoras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   6060
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data_acc 
      Height          =   375
      Left            =   840
      Top             =   6000
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
      Caption         =   "data_acc"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Data data_cargo 
      Caption         =   "data_cargo"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5400
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Filtrar solo por:"
      Height          =   1695
      Left            =   120
      TabIndex        =   10
      Top             =   4200
      Width           =   5775
      Begin MSAdodcLib.Adodc data_busca 
         Height          =   375
         Left            =   2880
         Top             =   240
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
         Caption         =   "data_busca"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C00000&
         Caption         =   "Oportunidad de Mejora"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   3975
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C00000&
         Caption         =   "No conformidad"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   3975
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C00000&
         Caption         =   "Observaciones"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   3975
      End
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   5640
      Top             =   5880
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
      Height          =   375
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6000
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5280
      Picture         =   "frm_infmejoras.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Salir"
      Top             =   5880
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      Picture         =   "frm_infmejoras.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Procesar"
      Top             =   5880
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Datos de informe"
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5775
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_infmejoras.frx":0F56
         Left            =   1800
         List            =   "frm_infmejoras.frx":0F58
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1800
         Width           =   3375
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C00000&
         Caption         =   "Ordenar por acción"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   120
         TabIndex        =   15
         Top             =   3360
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C00000&
         Caption         =   "Ordenar por origen"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   120
         TabIndex        =   14
         Top             =   2880
         Width           =   5055
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C00000&
         Caption         =   "Resumen"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   2880
         TabIndex        =   7
         Top             =   2400
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C00000&
         Caption         =   "Detalle"
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   390
         ItemData        =   "frm_infmejoras.frx":0F5A
         Left            =   1800
         List            =   "frm_infmejoras.frx":0F6D
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
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "Usuario:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Informe:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
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
      Height          =   615
      Left            =   1800
      Picture         =   "frm_infmejoras.frx":0FD0
      Stretch         =   -1  'True
      Top             =   5880
      Width           =   2415
   End
End
Attribute VB_Name = "frm_infmejoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check2.Value = 1 Then
   Check2.Value = 0
End If

End Sub

Private Sub Check2_Click()
If Check1.Value = 1 Then
   Check1.Value = 0
End If

End Sub

Private Sub Combo1_Click()
Option3.Value = False
Option4.Value = False
Option5.Value = False

End Sub

Private Sub Command1_Click()
Command1.Enabled = False
Command2.Enabled = False
Dim Xelu As String
Dim Xlafecve As Date
''On Error GoTo Alinfmam

Xlafecve = mh.Text
frm_infmejoras.MousePointer = 11
''data_busca.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and estado =" & 99 & " and cl_nomcobr =" & 1
'''data_busca.Refresh
'data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_nomcobr =" & 1
'data_acc.Refresh
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

If Combo2.ListIndex >= 0 Then
   data_cargo.Connect = "odbc;dsn=" & Xconexrmt & ";"
   data_cargo.RecordSource = "Select * from movil where chofer ='" & Combo2.Text & "'"
   data_cargo.Refresh
   If data_cargo.Recordset.RecordCount > 0 Then
      Xelu = data_cargo.Recordset("medico")
   Else
      Xelu = ""
   End If
Else
   Xelu = ""
End If

'If data_busca.Recordset.RecordCount > 0 Then
'   data_busca.Recordset.MoveFirst
'   Do While Not data_busca.Recordset.EOF
'      If IsNull(data_busca.Recordset("cl_val3")) = False Then
'         If data_busca.Recordset("cl_val3") = 1 Then
'            data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and estado =" & data_busca.Recordset("cl_nrovend") & " and cl_nomcobr =" & 1
'            data_acc.Refresh
'            If data_acc.Recordset.RecordCount > 0 Then
'               If data_acc.Recordset("cl_val3") = 1 Then
'               Else
'                  data_acc.Recordset("cl_val3") = 1
'                  data_acc.Recordset.Update
'               End If
'            End If
'         End If
'      End If
'
'      If IsNull(data_busca.Recordset("cl_nro_sup")) = False Then
'         If data_busca.Recordset("cl_nro_sup") >= 0 Then
'            data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and estado =" & data_busca.Recordset("cl_nrovend") & " and cl_nomcobr =" & 1
'            data_acc.Refresh
'            If data_acc.Recordset.RecordCount > 0 Then
'               If IsNull(data_acc.Recordset("cl_nro_sup")) = False Then
'                  If data_acc.Recordset("cl_nro_sup") = data_busca.Recordset("cl_nro_sup") Then
'                  Else
'                     data_acc.Recordset.Edit
'                     data_acc.Recordset("cl_nro_sup") = data_busca.Recordset("cl_nro_sup")
'                     data_acc.Recordset.Update
'                  End If
'               Else
'                  data_acc.Recordset.Edit
'                  data_acc.Recordset("cl_nro_sup") = data_busca.Recordset("cl_nro_sup")
'                  data_acc.Recordset.Update
 '              End If
'            End If
'         End If
'      End If
'      data_busca.Recordset.MoveNext
'   Loop
'End If

'data_busca.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_nomcobr =" & 1
'data_busca.Refresh

MiBaseact.Execute "Delete * from infcli"
data_inf.RecordSource = "infcli"
data_inf.Refresh

If md.Text <> "__/__/____" Then
   If mh.Text <> "__/__/____" Then
      If Combo1.ListIndex = 0 Then
         If Option3.Value = True Then
            If Xelu = "" Then
               data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and estado not in (98,99) and cl_fnac >='" & Format(md.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_atrasop =" & 1 & " and cl_nomcobr =" & 1
               data_acc.Refresh
            Else
               data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and estado not in (98,99) and cl_fnac >='" & Format(md.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_atrasop =" & 1 & " and cl_nom_sup ='" & Xelu & "' and cl_nomcobr =" & 1
               data_acc.Refresh
            End If
         Else
            If Option4.Value = True Then
               If Xelu = "" Then
                  data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and estado not in (98,99) and cl_fnac >='" & Format(md.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_atrasop =" & 2 & " and cl_nomcobr =" & 1
                  data_acc.Refresh
               Else
                  data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and estado not in (98,99) and cl_fnac >='" & Format(md.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_atrasop =" & 2 & " and cl_nom_sup ='" & Xelu & "' and cl_nomcobr =" & 1
                  data_acc.Refresh
               End If
            Else
               If Option5.Value = True Then
                  If Xelu = "" Then
                     data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and estado not in (98,99) and cl_fnac >='" & Format(md.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_atrasop =" & 3 & " and cl_nomcobr =" & 1
                     data_acc.Refresh
                  Else
                     data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and estado not in (98,99) and cl_fnac >='" & Format(md.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_atrasop =" & 3 & " and cl_nom_sup ='" & Xelu & "' and cl_nomcobr =" & 1
                     data_acc.Refresh
                  End If
               Else
                  If Xelu = "" Then
                     data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and estado not in (98,99) and cl_fnac >='" & Format(md.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_nomcobr =" & 1
                     data_acc.Refresh
                  Else
                     data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and estado not in (98,99) and cl_fnac >='" & Format(md.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_nom_sup ='" & Xelu & "' and cl_nomcobr =" & 1
                     data_acc.Refresh
                  End If
               End If
            End If
         End If
      Else
         If Combo1.ListIndex = 1 Then
            If Option3.Value = True Then
               If Xelu = "" Then
                  data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >='" & Format(md.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_atrasop =" & 1 & " and cl_nomcobr =" & 1
                  data_acc.Refresh
               Else
                  data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >='" & Format(md.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_atrasop =" & 1 & " and cl_nom_sup ='" & Xelu & "' and cl_nomcobr =" & 1
                  data_acc.Refresh
               End If
            Else
               If Option4.Value = True Then
                  If Xelu = "" Then
                     data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >='" & Format(md.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_atrasop =" & 2 & " and cl_nomcobr =" & 1
                     data_acc.Refresh
                  Else
                     data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >='" & Format(md.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_atrasop =" & 2 & " and cl_nom_sup ='" & Xelu & "' and cl_nomcobr =" & 1
                     data_acc.Refresh
                  End If
               Else
                  If Option5.Value = True Then
                     If Xelu = "" Then
                        data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >='" & Format(md.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_atrasop =" & 3 & " and cl_nomcobr =" & 1
                        data_acc.Refresh
                     Else
                        data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >='" & Format(md.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_atrasop =" & 3 & " and cl_nom_sup ='" & Xelu & "' and cl_nomcobr =" & 1
                        data_acc.Refresh
                     End If
                  Else
                     If Xelu = "" Then
                        data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >='" & Format(md.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_nomcobr =" & 1
                        data_acc.Refresh
                     Else
                        data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >='" & Format(md.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_nom_sup ='" & Xelu & "' and cl_nomcobr =" & 1
                        data_acc.Refresh
                     End If
                  End If
               End If
            End If
         Else '2 pendientes de evaluacion
            If Combo1.ListIndex = 2 Or Combo1.ListIndex = 3 Then '3 Conformidad
               If Option3.Value = True Then
                  If Xelu = "" Then
                     data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >='" & Format(md.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_atrasop =" & 1 & " and cl_nomcobr =" & 1
                     data_acc.Refresh
                  Else
                     data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >='" & Format(md.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_atrasop =" & 1 & " and cl_nom_sup ='" & Xelu & "' and cl_nomcobr =" & 1
                     data_acc.Refresh
                  End If
               Else
                  If Option4.Value = True Then
                     If Xelu = "" Then
                        data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >='" & Format(md.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_atrasop =" & 2 & " and cl_nomcobr =" & 1
                        data_acc.Refresh
                     Else
                        data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >='" & Format(md.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_atrasop =" & 2 & " and cl_nom_sup ='" & Xelu & "' and cl_nomcobr =" & 1
                        data_acc.Refresh
                     End If
                  Else
                     If Option5.Value = True Then
                        If Xelu = "" Then
                           data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >='" & Format(md.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_atrasop =" & 3 & " and cl_nomcobr =" & 1
                           data_acc.Refresh
                        Else
                           data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >='" & Format(md.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_atrasop =" & 3 & " and cl_nom_sup ='" & Xelu & "' and cl_nomcobr =" & 1
                           data_acc.Refresh
                        End If
                     Else
                        If Xelu = "" Then
                        'acá
                           data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and estado not in (98,99,97) and cl_fnac >='" & Format(md.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_nomcobr =" & 1
                           data_acc.Refresh
                        Else
                           data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and estado not in (98,99,97) and cl_fnac >='" & Format(md.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_nom_sup ='" & Xelu & "' and cl_nomcobr =" & 1
                           data_acc.Refresh
                        End If
                     End If
                  End If
               End If
            Else
               If Combo1.ListIndex = 4 Then
                  If Xelu = "" Then
                     data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >='" & Format(md.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_nomcobr =" & 1 & " and cl_fec1 <='" & Format(Xlafecve, "yyyy-mm-dd") & "'"
                     data_acc.Refresh
                  Else
                     data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >='" & Format(md.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_nom_sup ='" & Xelu & "' and cl_nomcobr =" & 1 & " and cl_fec1 <='" & Format(Xlafecve, "yyyy-mm-dd") & "'"
                     data_acc.Refresh
                  End If
               Else
                  If Xelu = "" Then
                     data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >='" & Format(md.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_nomcobr =" & 1
                     data_acc.Refresh
                  Else
                     data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >='" & Format(md.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_nom_sup ='" & Xelu & "' and cl_nomcobr =" & 1
                     data_acc.Refresh
                  End If
               End If
            End If
         End If
      End If
      If data_acc.Recordset.RecordCount > 0 Then
         data_acc.Recordset.MoveFirst
         Do While Not data_acc.Recordset.EOF
            data_inf.Recordset.AddNew
            data_inf.Recordset("cl_codigo") = data_acc.Recordset("estado")
            data_inf.Recordset("cl_fnac") = data_acc.Recordset("cl_fnac")
            data_inf.Recordset("cl_cantpag") = data_acc.Recordset("cl_nrovend")
            data_inf.Recordset("cl_descpag") = data_acc.Recordset("cl_descpag") ' usuario que lo crea
            data_inf.Recordset("cl_forpago") = data_acc.Recordset("cl_nro_sup")
            data_inf.Recordset("cl_localid") = data_acc.Recordset("cl_descpag")
            data_inf.Recordset("cl_apellid") = data_acc.Recordset("cl_desc1")
            data_inf.Recordset("cl_dircobr") = Mid(data_acc.Recordset("info_debit"), 1, 100)
            data_inf.Recordset("cl_nom_sup") = data_acc.Recordset("cl_nom_sup") ' usuario que modifica
            data_inf.Recordset("cl_nomvend") = data_acc.Recordset("cl_nom_sup")
            If Combo1.ListIndex = 4 Or Combo1.ListIndex = 1 Then
               data_inf.Recordset("cl_fultvta") = data_acc.Recordset("cl_fec1")
            Else
               data_inf.Recordset("cl_fultvta") = data_acc.Recordset("cl_fultpag")
            End If
            data_inf.Recordset("cl_cua_vto") = data_acc.Recordset("cl_val3")
            data_inf.Recordset("cl_grupo") = data_acc.Recordset("cl_grupo")
            If Option3.Value = True Then
               data_inf.Recordset("tit_tarj") = "FILTRADO POR OBSERVACIONES"
            Else
               If Option4.Value = True Then
                  data_inf.Recordset("tit_tarj") = "FILTRADO NO CONFORMIDADES"
               Else
                  If Option5.Value = True Then
                     data_inf.Recordset("tit_tarj") = "FILTRADO OP.DE MEJORA"
                  Else
                     data_inf.Recordset("tit_tarj") = ""
                  End If
               End If
            End If
            If IsNull(data_acc.Recordset("cl_grupo")) = False Then
               If data_acc.Recordset("cl_grupo") = 0 Then
                  data_inf.Recordset("cl_tipclin") = "AUDITORIA INT."
               Else
                  If data_acc.Recordset("cl_grupo") = 1 Then
                     data_inf.Recordset("cl_tipclin") = "AUDITORIA EXT."
                  Else
                     If data_acc.Recordset("cl_grupo") = 2 Then
                        data_inf.Recordset("cl_tipclin") = "SERVICIO"
                     Else
                        If data_inf.Recordset("cl_grupo") = 3 Then
                           data_inf.Recordset("cl_tipclin") = "RECLAMO"
                        Else
                           If data_inf.Recordset("cl_grupo") = 4 Then
                              data_inf.Recordset("cl_tipclin") = "INDICADORES"
                           Else
                              If data_inf.Recordset("cl_grupo") = 5 Then
                                 data_inf.Recordset("cl_tipclin") = "OTROS"
                              Else
                                 data_inf.Recordset("cl_tipclin") = "ORGANO CONTRALOR"
                              End If
                           End If
                        End If
                     End If
                  End If
               End If
            Else
               data_inf.Recordset("cl_tipclin") = "OTROS"
            End If
            If IsNull(data_acc.Recordset("cl_val3")) = False Then
               If data_acc.Recordset("cl_val3") = 1 Then
                  data_inf.Recordset("cl_zona") = "TERMINADO C/CONFORME"
                  If Combo1.ListIndex = 0 Then
                     If IsNull(data_acc.Recordset("cl_val1")) = False Then
                        If data_acc.Recordset("cl_val1") < 0 Then
                           data_inf.Recordset("cl_zona") = "TERMINADO S/CONFORME"
                           data_inf.Recordset("estado") = 2
                        Else
                           data_inf.Recordset("estado") = 0
                        End If
                     Else
                        data_inf.Recordset("cl_zona") = "TERMINADO S/CONFORME"
                        data_inf.Recordset("estado") = 2
                     End If
                  Else
                     If Combo1.ListIndex = 2 Then
                        If IsNull(data_acc.Recordset("cl_val1")) = False Then
                           If data_acc.Recordset("cl_val1") < 0 Then
                              data_inf.Recordset("cl_dpto") = "S/REGISTRAR"
                              data_inf.Recordset("cl_zona") = "TERMINADO S/CONFORME"
                           End If
                        Else
                           data_inf.Recordset("cl_dpto") = "S/REGISTRAR"
                           data_inf.Recordset("cl_zona") = "TERMINADO S/CONFORME"
                        End If
                     Else
                        If IsNull(data_acc.Recordset("cl_val1")) = False Then
                           If data_acc.Recordset("cl_val1") = 0 Then
                              data_inf.Recordset("cl_dpto") = "CONFORME"
                           Else
                              If data_acc.Recordset("cl_val1") = 1 Then
                                 data_inf.Recordset("cl_dpto") = "CON DEMORA"
                              Else
                                 If data_acc.Recordset("cl_val1") = 2 Then
                                    data_inf.Recordset("cl_dpto") = "NO CONFORME"
                                 Else
                                    data_inf.Recordset("cl_dpto") = "S/REGISTRAR"
                                    data_inf.Recordset("cl_zona") = "TERMINADO S/CONFORME"
    ''                                 data_inf.Recordset("estado") = 4
                                 End If
                              End If
                           End If
                        Else
                           data_inf.Recordset("cl_dpto") = "S/REGISTRAR"
                        End If
                     End If
                     data_inf.Recordset("estado") = data_acc.Recordset("cl_val1")
                  End If
               Else
                  data_inf.Recordset("cl_zona") = "EN PROCESO C/REGISTRO"
                  data_inf.Recordset("estado") = -1
               End If
            Else
               data_inf.Recordset("cl_zona") = "EN PROCESO C/REGISTRO"
               data_inf.Recordset("estado") = -1
            End If
            data_inf.Recordset("saldo_cc") = data_acc.Recordset("cl_val1")
            data_inf.Recordset.Update
            data_acc.Recordset.MoveNext
         Loop
         If data_inf.Recordset.RecordCount > 0 Then
            data_inf.Recordset.MoveFirst
            If Combo1.ListIndex = 0 Then ''' todas
               Do While Not data_inf.Recordset.EOF
                  data_busca.RecordSource = "Select * from infor_sol where cl_nrovend =" & data_inf.Recordset("cl_codigo") & " and cl_nomcobr =" & 1 & " and estado in (99,98)"
                  data_busca.Refresh
                  If data_busca.Recordset.RecordCount > 0 Then
                     If IsNull(data_busca.Recordset("cl_descpag")) = False Then
                        data_inf.Recordset.Edit
                        data_inf.Recordset("cl_nomcobr") = data_busca.Recordset("cl_descpag")
                        data_inf.Recordset.Update
                     End If
                  Else
                     data_inf.Recordset.Edit
                     data_inf.Recordset("cl_zona") = "EN PROCESO S/REGISTRO"
                     data_inf.Recordset("estado") = 1
                     data_inf.Recordset.Update
                  End If
                  data_inf.Recordset.MoveNext
               Loop
            End If
            If Combo1.ListIndex = 1 Or Combo1.ListIndex = 4 Then ''' pendientes
               Do While Not data_inf.Recordset.EOF
                  data_busca.RecordSource = "Select * from infor_sol where cl_nrovend =" & data_inf.Recordset("cl_codigo") & " and cl_nomcobr =" & 1
                  data_busca.Refresh
                  If data_busca.Recordset.RecordCount > 0 Then
                     If IsNull(data_busca.Recordset("cl_val3")) = False Then
                        If data_busca.Recordset("cl_val3") = 1 Then
                           data_inf.Recordset.Edit
                           data_inf.Recordset("cl_codced") = 8
                           data_inf.Recordset.Update
                        End If
                     End If
                  Else
                     data_inf.Recordset.Edit
                     data_inf.Recordset("cl_zona") = "EN PROCESO S/REGISTRO"
                     data_inf.Recordset("estado") = 1
                     data_inf.Recordset.Update
                  End If
                  data_inf.Recordset.MoveNext
               Loop
               MiBaseact.Execute "Delete * from infcli where cl_codced =" & 8
               data_inf.Refresh
            End If
               
            If Combo1.ListIndex = 2 Or Combo1.ListIndex = 3 Then ''' pend.evaluación y conformes
               Do While Not data_inf.Recordset.EOF
                  data_busca.RecordSource = "Select * from infor_sol where cl_nrovend =" & data_inf.Recordset("cl_codigo") & " and estado in (99,98)"
                  data_busca.Refresh
                  If data_busca.Recordset.RecordCount > 0 Then
                     If IsNull(data_busca.Recordset("cl_descpag")) = False Then
                        data_inf.Recordset.Edit
                        data_inf.Recordset("cl_nomcobr") = data_busca.Recordset("cl_descpag")
                        data_inf.Recordset.Update
                     End If
                     If IsNull(data_busca.Recordset("cl_val3")) = False Then
                        If data_busca.Recordset("cl_val3") = 1 Then
                           If IsNull(data_inf.Recordset("saldo_cc")) = False Then
                              If data_inf.Recordset("saldo_cc") < 0 Then
                                 data_inf.Recordset.Edit
                                 data_inf.Recordset("cl_zona") = "TERMINADO S/CONFORME"
                                 data_inf.Recordset.Update
                              Else
                                 data_inf.Recordset.Edit
                                 data_inf.Recordset("cl_zona") = "TERMINADO C/CONFORME"
                                 data_inf.Recordset.Update
                              End If
                           Else
                              data_inf.Recordset.Edit
                              data_inf.Recordset("cl_zona") = "TERMINADO S/CONFORME"
                              data_inf.Recordset.Update
                           End If
                        Else
                           data_inf.Recordset.Delete
                        End If
                     Else
                        data_inf.Recordset.Delete
'                     data_inf.Recordset.Edit
'                     data_inf.Recordset("cl_zona") = "EN PROCESO S/REGISTRO"
'                     data_inf.Recordset.Update
                     
                     End If
                  Else
                     data_inf.Recordset.Delete
'                     data_inf.Recordset.Edit
'                     data_inf.Recordset("cl_zona") = "EN PROCESO C/REGISTRO"
'                     data_inf.Recordset.Update
                  End If
                  data_inf.Recordset.MoveNext
               Loop
            End If
            
            If Combo1.ListIndex = 4 Then
               If data_inf.Recordset.RecordCount > 0 Then
                  data_inf.Recordset.MoveFirst
                  Do While Not data_inf.Recordset.EOF
                     If data_inf.Recordset("cl_zona") = "TERMINADO C/CONFORME" Then
                        data_inf.Recordset.Delete
                     End If
                     data_inf.Recordset.MoveNext
                  Loop
               End If
            End If
         End If
'If Combo1.ListIndex = 3 Then
'            If data_inf.Recordset.RecordCount > 0 Then
'               data_inf.Recordset.MoveFirst
'            End If
'            Do While Not data_inf.Recordset.EOF
'               If IsNull(data_inf.Recordset("cl_cua_vto")) = True Then
'                  data_inf.Recordset.Delete
''               Else
'                  If data_inf.Recordset("cl_cua_vto") = 1 Then
'                  Else
'                     data_inf.Recordset.Delete
'                  End If
'               End If
'               data_inf.Recordset.MoveNext
'            Loop
'         End If
         
         frm_infmejoras.MousePointer = 0
         MsgBox "Proceso terminado"
         
         data_inf.RecordSource = "Select * from infcli order by cl_fnac"
         data_inf.Refresh
         If Combo1.ListIndex = 0 Or Combo1.ListIndex = 4 Then
            If Combo1.ListIndex = 4 Then
               cr1.ReportFileName = App.path & "\infmejorvenc.rpt"
            Else
               If Check1.Value = 1 Then
                  cr1.ReportFileName = App.path & "\infmejor3.rpt"
               Else
                  If Check2.Value = 1 Then
                     cr1.ReportFileName = App.path & "\infmejor5.rpt"
                  Else
                     cr1.ReportFileName = App.path & "\infmejor1.rpt"
                  End If
               End If
            End If
            cr1.ReportTitle = "INFORME SOLICITUDES DE MEJORA CONTINUA DESDE: " & md.Text & " HASTA: " & mh.Text
            cr1.Action = 1
         End If
         If Combo1.ListIndex = 1 Then
            If Check1.Value = 1 Then
               cr1.ReportFileName = App.path & "\infmejor3.rpt"
            Else
               If Check2.Value = 1 Then
                  cr1.ReportFileName = App.path & "\infmejor5.rpt"
               Else
                  cr1.ReportFileName = App.path & "\infmejor1.rpt"
               End If
            End If
            cr1.ReportTitle = "INFORME SOLICITUDES PENDIENTES DESDE: " & md.Text & " HASTA: " & mh.Text
            cr1.Action = 1
         End If
         If Combo1.ListIndex = 2 Then
            If Check1.Value = 1 Then
               cr1.ReportFileName = App.path & "\infmejorpor.rpt"
            Else
               If Check2.Value = 1 Then
                  cr1.ReportFileName = App.path & "\infmejorpac.rpt"
               Else
                  cr1.ReportFileName = App.path & "\infmejor1cd.rpt"
               End If
            End If
            cr1.ReportTitle = "INFORME SOLICITUDES CUMPLIDAS SIN CONFORMIDAD DESDE: " & md.Text & " HASTA: " & mh.Text
            cr1.Action = 1
         End If
         If Combo1.ListIndex = 3 Then
            If Check1.Value = 1 Then
               cr1.ReportFileName = App.path & "\infmejor6.rpt"
            Else
               If Check2.Value = 1 Then
                  cr1.ReportFileName = App.path & "\infmejor7.rpt"
               Else
                  cr1.ReportFileName = App.path & "\infmejorconf.rpt"
               End If
            End If
            cr1.ReportTitle = "INFORME CONFORMIDAD DE SOLICITUDES DESDE: " & md.Text & " HASTA: " & mh.Text
            cr1.Action = 1
         End If
         
      Else
         frm_infmejoras.MousePointer = 0
         MsgBox "No existen registros"
      End If
   End If
End If
   
Command1.Enabled = True
Command2.Enabled = True
frm_infmejoras.MousePointer = 0
'Exit Sub

'Alinfmam:
'        If Err.Number = 13 Then
'           MsgBox "Error en los datos, verifique!", vbInformation
'        Else
'           MsgBox "Verifique datos para el informe", vbInformation
'        End If
        

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
'data_acc.DatabaseName = App.Path & "\sapp.mdb"
data_acc.ConnectionString = "dsn=" & Xconexrmt
data_inf.DatabaseName = App.path & "\informes.mdb"
'data_inf.RecordSource = "infcli"
'data_inf.Refresh
data_busca.ConnectionString = "dsn=" & Xconexrmt

Combo1.ListIndex = 0
Combo2.Clear
If WElusuario = "BDD" Or WElusuario = "SPEREZ" Or WElusuario = "JFERNAN" Or WElusuario = "SDOMINGUEZ" Or WElusuario = "AGUILLEN" Or WElusuario = "MEUGENIA" Then
    Combo2.AddItem "TODOS"
    Combo2.AddItem "DIRECTOR GENERAL"
    Combo2.AddItem "GERENTE GENERAL"
    Combo2.AddItem "DIRECCION TECNICA"
    Combo2.AddItem "SUB-DIREC.TECNICA"
    'Combo1.AddItem "GERENTE COMERCIAL"
    'Combo1.AddItem "JEFE DE MEDICOS DE MOVIL"
    Combo2.AddItem "JEFE CHOFERES Y MANT."
    Combo2.AddItem "JEFE ADMINISTRACION"
    Combo2.AddItem "JEFE DPTO.TI"
    Combo2.AddItem "JEFE REGIONAL COSTA"
    Combo2.AddItem "JEFE FARMACIA/ECONOMATO"
    Combo2.AddItem "JEFE DESPACHO"
    'Combo1.AddItem "ENCARGADO METAS"
    'Combo1.AddItem "JEFE ATENCION AL CLIENTE"
    Combo2.AddItem "SUB-JEFE FACTURACION"
    Combo2.AddItem "JEFE CONTADURIA"
    Combo2.AddItem "JEFE REGIONAL NORTE"
    Combo2.AddItem "JEFE COMERCIAL"
    Combo2.AddItem "RESPONSABLE CALIDAD"
    Combo2.AddItem "JEFE ASISTENCIAL"
    Combo1.AddItem "SUB-JEFE TESORERIA" 'Paola
    Combo2.ListIndex = -1
Else
    Combo2.AddItem "DIRECCION TECNICA"
    Combo2.AddItem "SUB-DIREC.TECNICA"
    'Combo1.AddItem "GERENTE COMERCIAL"
    'Combo1.AddItem "JEFE DE MEDICOS DE MOVIL"
    Combo2.AddItem "JEFE CHOFERES Y MANT."
    Combo2.AddItem "JEFE ADMINISTRACION"
    Combo2.AddItem "JEFE DPTO.TI"
    Combo2.AddItem "JEFE REGIONAL COSTA"
    Combo2.AddItem "JEFE FARMACIA/ECONOMATO"
    Combo2.AddItem "JEFE DESPACHO"
    'Combo1.AddItem "ENCARGADO METAS"
    'Combo1.AddItem "JEFE ATENCION AL CLIENTE"
    Combo2.AddItem "SUB-JEFE FACTURACION"
    Combo2.AddItem "JEFE CONTADURIA"
    Combo2.AddItem "JEFE REGIONAL NORTE"
    Combo2.AddItem "JEFE COMERCIAL"
    Combo2.AddItem "JEFE ASISTENCIAL"
    Combo1.AddItem "SUB-JEFE TESORERIA" 'Paola
    Combo2.ListIndex = -1

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
   Combo1.SetFocus
End If

End Sub
