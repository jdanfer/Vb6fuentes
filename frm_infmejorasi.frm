VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_infmejorasi 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de solicitudes"
   ClientHeight    =   6510
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
   Icon            =   "frm_infmejorasi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6510
   ScaleWidth      =   6060
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_cargo 
      Caption         =   "data_cargo"
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
      Top             =   5520
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H000040C0&
      Caption         =   "Filtrar solo por:"
      Height          =   1695
      Left            =   120
      TabIndex        =   10
      Top             =   4200
      Width           =   5775
      Begin VB.OptionButton Option5 
         BackColor       =   &H0000FF00&
         Caption         =   "Oportunidad de Mejora"
         Height          =   270
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   3975
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H0000FF00&
         Caption         =   "No conformidad"
         Height          =   270
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   3975
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H0000FF00&
         Caption         =   "Observaciones"
         Height          =   270
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   3975
      End
   End
   Begin VB.Data data_busca 
      Caption         =   "data_busca"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4080
      Visible         =   0   'False
      Width           =   2535
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
   Begin VB.Data data_acc 
      Caption         =   "data_acc"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5520
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
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5160
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5280
      Picture         =   "frm_infmejorasi.frx":0442
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
      Picture         =   "frm_infmejorasi.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Procesar"
      Top             =   5880
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000040C0&
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
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1800
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H0000FF00&
         Caption         =   "Ordenar por acción"
         Height          =   270
         Left            =   120
         TabIndex        =   15
         Top             =   3360
         Width           =   5055
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H0000FF00&
         Caption         =   "Ordenar por origen"
         Height          =   270
         Left            =   120
         TabIndex        =   14
         Top             =   2880
         Width           =   5055
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H0000FF00&
         Caption         =   "Resumen"
         Height          =   270
         Left            =   2880
         TabIndex        =   7
         Top             =   2400
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H0000FF00&
         Caption         =   "Detalle"
         Height          =   270
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   390
         ItemData        =   "frm_infmejorasi.frx":0F56
         Left            =   1800
         List            =   "frm_infmejorasi.frx":0F63
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
         BackColor       =   &H00FF8080&
         Caption         =   "Usuario:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF8080&
         Caption         =   "Informe:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF8080&
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
      Left            =   2400
      Picture         =   "frm_infmejorasi.frx":0FA2
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   2415
   End
End
Attribute VB_Name = "frm_infmejorasi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check2.value = 1 Then
   Check2.value = 0
End If

End Sub

Private Sub Check2_Click()
If Check1.value = 1 Then
   Check1.value = 0
End If

End Sub

Private Sub Combo1_Click()
Option3.value = False
Option4.value = False
Option5.value = False

End Sub

Private Sub Command1_Click()
Command1.Enabled = False
Command2.Enabled = False
Dim Xelu As String

frm_infmejoras.MousePointer = 11
data_busca.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and estado =" & 99 & " and cl_nomcobr =" & 2
data_busca.Refresh
data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_nomcobr =" & 2
data_acc.Refresh
If Combo2.ListIndex > 0 Then
   data_cargo.Connect = "odbc;dsn=" & Xconexrmt & ";"
   data_cargo.RecordSource = "Select * from movil where chofer ='" & Combo2.Text & "'"
   data_cargo.Refresh
   If data_cargo.Recordset.RecordCount > 0 Then
      Xelu = data_cargo.Recordset("medico")
   Else
      Xelu = WElusuario
   End If
Else
   Xelu = ""
End If

If data_busca.Recordset.RecordCount > 0 Then
   data_busca.Recordset.MoveFirst
   Do While Not data_busca.Recordset.EOF
      If IsNull(data_busca.Recordset("cl_val3")) = False Then
         If data_busca.Recordset("cl_val3") = 1 Then
'            data_acc.Recordset.FindFirst "estado =" & data_busca.Recordset("cl_nrovend")
'            If Not data_acc.Recordset.NoMatch Then
            data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and estado =" & data_busca.Recordset("cl_nrovend") & " and cl_nomcobr =" & 2
            data_acc.Refresh
            If data_acc.Recordset.RecordCount > 0 Then
               If data_acc.Recordset("cl_val3") = 1 Then
               Else
                  data_acc.Recordset.Edit
                  data_acc.Recordset("cl_val3") = 1
                  data_acc.Recordset.Update
               End If
            End If
         End If
      End If
      
      If IsNull(data_busca.Recordset("cl_nro_sup")) = False Then
         If data_busca.Recordset("cl_nro_sup") >= 0 Then
'            data_acc.Recordset.FindFirst "estado =" & data_busca.Recordset("cl_nrovend")
'            If Not data_acc.Recordset.NoMatch Then
            data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and estado =" & data_busca.Recordset("cl_nrovend") & " and cl_nomcobr =" & 2
            data_acc.Refresh
            If data_acc.Recordset.RecordCount > 0 Then
               If IsNull(data_acc.Recordset("cl_nro_sup")) = False Then
                  If data_acc.Recordset("cl_nro_sup") = data_busca.Recordset("cl_nro_sup") Then
                  Else
                     data_acc.Recordset.Edit
                     data_acc.Recordset("cl_nro_sup") = data_busca.Recordset("cl_nro_sup")
                     data_acc.Recordset.Update
                  End If
               Else
                  data_acc.Recordset.Edit
                  data_acc.Recordset("cl_nro_sup") = data_busca.Recordset("cl_nro_sup")
                  data_acc.Recordset.Update
               End If
            End If
         End If
      End If
      data_busca.Recordset.MoveNext
   Loop
End If
data_busca.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_nomcobr =" & 2
data_busca.Refresh

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
         If Option3.value = True Then
            If Xelu = "" Then
               data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and estado <>" & 99 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cl_atrasop =" & 1 & " and cl_nomcobr =" & 2
               data_acc.Refresh
            Else
               data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and estado <>" & 99 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cl_atrasop =" & 1 & " and cl_nomcobr =" & 2 & " and cl_nom_sup ='" & Xelu & "'"
               data_acc.Refresh
            End If
         Else
            If Option4.value = True Then
               If Xelu = "" Then
                  data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and estado <>" & 99 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cl_atrasop =" & 2 & " and cl_nomcobr =" & 2
                  data_acc.Refresh
               Else
                  data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and estado <>" & 99 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cl_atrasop =" & 2 & " and cl_nomcobr =" & 2 & " and cl_nom_sup ='" & Xelu & "'"
                  data_acc.Refresh
               End If
            Else
               If Option5.value = True Then
                  If Xelu = "" Then
                     data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and estado <>" & 99 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cl_atrasop =" & 3 & " and cl_nomcobr =" & 2
                     data_acc.Refresh
                  Else
                     data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and estado <>" & 99 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cl_atrasop =" & 3 & " and cl_nomcobr =" & 2 & " and cl_nom_sup ='" & Xelu & "'"
                     data_acc.Refresh
                  End If
               Else
                  If Xelu = "" Then
                     data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and estado <>" & 99 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cl_nomcobr =" & 2
                     data_acc.Refresh
                  Else
                     data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and estado <>" & 99 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cl_nomcobr =" & 2 & " and cl_nom_sup ='" & Xelu & "'"
                     data_acc.Refresh
                  End If
               End If
            End If
         End If
      Else
         If Combo1.ListIndex = 1 Then
            If Option3.value = True Then
               If Xelu = "" Then
                  data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cl_atrasop =" & 1 & " and cl_nomcobr =" & 2
                  data_acc.Refresh
               Else
                  data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cl_atrasop =" & 1 & " and cl_nomcobr =" & 2 & " and cl_nom_sup ='" & Xelu & "'"
                  data_acc.Refresh
               End If
            Else
               If Option4.value = True Then
                  If Xelu = "" Then
                     data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cl_atrasop =" & 2 & " and cl_nomcobr =" & 2
                     data_acc.Refresh
                  Else
                     data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cl_nomcobr =" & 2 & " and cl_atrasop =" & 2 & " and cl_nom_sup ='" & Xelu & "'"
                     data_acc.Refresh
                  End If
               Else
                  If Option5.value = True Then
                     If Xelu = "" Then
                        data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cl_atrasop =" & 3 & " and cl_nomcobr =" & 2
                        data_acc.Refresh
                     Else
                        data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cl_atrasop =" & 3 & " and cl_nomcobr =" & 2 & " and cl_nom_sup ='" & Xelu & "'"
                        data_acc.Refresh
                     End If
                  Else
                     If Xelu = "" Then
                        data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cl_nomcobr =" & 2
                        data_acc.Refresh
                     Else
                        data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cl_nomcobr =" & 2 & " and cl_nom_sup ='" & Xelu & "'"
                        data_acc.Refresh
                     End If
                  End If
               End If
            End If
         Else
            If Combo1.ListIndex = 2 Or Combo1.ListIndex = 3 Then
               If Option3.value = True Then
                  If Xelu = "" Then
                     data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cl_atrasop =" & 1 & " and cl_nomcobr =" & 2
                     data_acc.Refresh
                  Else
                     data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cl_atrasop =" & 1 & " and cl_nomcobr =" & 2 & " and cl_nom_sup ='" & Xelu & "'"
                     data_acc.Refresh
                  End If
               Else
                  If Option4.value = True Then
                     If Xelu = "" Then
                        data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cl_atrasop =" & 2 & " and cl_nomcobr =" & 2
                        data_acc.Refresh
                     Else
                        data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cl_atrasop =" & 2 & " and cl_nomcobr =" & 2 & " and cl_nom_sup ='" & Xelu & "'"
                        data_acc.Refresh
                     End If
                  Else
                     If Option5.value = True Then
                        If Xelu = "" Then
                           data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cl_atrasop =" & 3 & " and cl_nomcobr =" & 2
                           data_acc.Refresh
                        Else
                           data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cl_atrasop =" & 3 & " and cl_nomcobr =" & 2 & " and cl_nom_sup ='" & Xelu & "'"
                           data_acc.Refresh
                        End If
                     Else
                        If Xelu = "" Then
                           data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and estado <>" & 99 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cl_nomcobr =" & 2
                           data_acc.Refresh
                        Else
                           data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and estado <>" & 99 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cl_nomcobr =" & 2 & " and cl_nom_sup ='" & Xelu & "'"
                           data_acc.Refresh
                        End If
                     End If
                  End If
               End If
            Else
               If Xelu = "" Then
                  data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cl_nomcobr =" & 2
                  data_acc.Refresh
               Else
                  data_acc.RecordSource = "Select * from infor_sol where cl_val2 =" & 7 & " and cl_fnac >=#" & Format(md.Text, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and cl_nomcobr =" & 2 & " and cl_nom_sup ='" & Xelu & "'"
                  data_acc.Refresh
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
            data_inf.Recordset("cl_fultvta") = data_acc.Recordset("cl_fultpag")
            data_inf.Recordset("cl_cua_vto") = data_acc.Recordset("cl_val3")
            data_inf.Recordset("cl_grupo") = data_acc.Recordset("cl_grupo")
            If Option3.value = True Then
               data_inf.Recordset("tit_tarj") = "FILTRADO POR OBSERVACIONES"
            Else
               If Option4.value = True Then
                  data_inf.Recordset("tit_tarj") = "FILTRADO NO CONFORMIDADES"
               Else
                  If Option5.value = True Then
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
                              data_inf.Recordset("cl_tipclin") = "OTROS"
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
            data_inf.Recordset("cl_nro_sup") = data_acc.Recordset("cl_nro_sup")
            If IsNull(data_acc.Recordset("cl_nro_sup")) = False Then
               If data_acc.Recordset("cl_nro_sup") = 0 Then
                  data_inf.Recordset("cl_nomcobr") = "C.INMEDIATA"
               End If
               If data_acc.Recordset("cl_nro_sup") = 1 Then
                  data_inf.Recordset("cl_nomcobr") = "CORRECTIVA"
               End If
               If data_acc.Recordset("cl_nro_sup") = 2 Then
                  data_inf.Recordset("cl_nomcobr") = "PREVENTIVA"
               End If
               If data_acc.Recordset("cl_nro_sup") = 3 Then
                  data_inf.Recordset("cl_nomcobr") = "PLAN DE ACCION"
               End If
               If data_acc.Recordset("cl_nro_sup") = 4 Then
                  data_inf.Recordset("cl_nomcobr") = "OTRO"
               End If
            Else
            End If
            data_inf.Recordset.Update
            data_acc.Recordset.MoveNext
         Loop
         If data_inf.Recordset.RecordCount > 0 Then
            If Combo1.ListIndex = 0 Then ''' todas
               data_inf.Recordset.MoveFirst
               Do While Not data_inf.Recordset.EOF
                  data_busca.RecordSource = "Select * from infor_sol where cl_nrovend =" & data_inf.Recordset("cl_codigo") & " and cl_nomcobr =" & 2
                  data_busca.Refresh
                  If data_busca.Recordset.RecordCount > 0 Then
                     data_busca.Recordset.MoveFirst
                     Do While Not data_busca.Recordset.EOF
                        If IsNull(data_busca.Recordset("cl_val3")) = False Then
                           If data_busca.Recordset("cl_val3") = 1 Then
'                              data_inf.Recordset.Edit
'                              data_inf.Recordset("cl_codced") = 8
'                              data_inf.Recordset.Update
'                              data_inf.Recordset.Delete
                           End If
                        End If
                        data_busca.Recordset.MoveNext
                     Loop
                  Else
                     data_inf.Recordset.Edit
                     data_inf.Recordset("cl_zona") = "EN PROCESO S/REGISTRO"
                     data_inf.Recordset("estado") = 1
                     data_inf.Recordset.Update
                  End If
                  data_inf.Recordset.MoveNext
               Loop
            End If
         End If
         If data_inf.Recordset.RecordCount > 0 Then
            data_inf.Recordset.MoveFirst
            If Combo1.ListIndex = 1 Then ''' pendientes
               Do While Not data_inf.Recordset.EOF
                  data_busca.RecordSource = "Select * from infor_sol where cl_nrovend =" & data_inf.Recordset("cl_codigo") & " and cl_nomcobr =" & 2
                  data_busca.Refresh
                  If data_busca.Recordset.RecordCount > 0 Then
                     data_busca.Recordset.MoveFirst
                     Do While Not data_busca.Recordset.EOF
                        If IsNull(data_busca.Recordset("cl_val3")) = False Then
                           If data_busca.Recordset("cl_val3") = 1 Then
                              data_inf.Recordset.Edit
                              data_inf.Recordset("cl_codced") = 8
                              data_inf.Recordset.Update
'                              data_inf.Recordset.Delete
                           End If
                        End If
                        data_busca.Recordset.MoveNext
                     Loop
                  Else
                     data_inf.Recordset.Edit
                     data_inf.Recordset("cl_zona") = "EN PROCESO S/REGISTRO"
                     data_inf.Recordset("estado") = 1
                     data_inf.Recordset.Update
                  End If
                  data_inf.Recordset.MoveNext
               Loop
               data_inf.Recordset.MoveFirst
               Do While Not data_inf.Recordset.EOF
                  If data_inf.Recordset("cl_codced") = 8 Then
                     data_inf.Recordset.Delete
                  End If
                  data_inf.Recordset.MoveNext
               Loop
            End If
            If Combo1.ListIndex = 2 Or Combo1.ListIndex = 3 Then ''' cumplidas
               Do While Not data_inf.Recordset.EOF
                  data_busca.RecordSource = "Select * from infor_sol where cl_nrovend =" & data_inf.Recordset("cl_codigo") & " and cl_nomcobr =" & 2
                  data_busca.Refresh
                  If data_busca.Recordset.RecordCount > 0 Then
                     data_busca.Recordset.MoveFirst
                     Do While Not data_busca.Recordset.EOF
                        If IsNull(data_busca.Recordset("cl_val3")) = False Then
                           If data_busca.Recordset("cl_val3") = 1 Then
                           Else
                              data_inf.Recordset.Edit
                              data_inf.Recordset("cl_codced") = 8
                              data_inf.Recordset.Update
                           End If
                        Else
                           data_inf.Recordset.Edit
                           data_inf.Recordset("cl_codced") = 8
                           data_inf.Recordset.Update
                        End If
                        If IsNull(data_busca.Recordset("cl_nro_sup")) = False Then
                           data_inf.Recordset.Edit
                           If data_busca.Recordset("cl_nro_sup") = 0 Then
                              data_inf.Recordset("cl_nomcobr") = "C.INMEDIATA"
                           End If
                           If data_busca.Recordset("cl_nro_sup") = 1 Then
                              data_inf.Recordset("cl_nomcobr") = "CORRECTIVA"
                           End If
                           If data_busca.Recordset("cl_nro_sup") = 2 Then
                              data_inf.Recordset("cl_nomcobr") = "PREVENTIVA"
                           End If
                           If data_busca.Recordset("cl_nro_sup") = 3 Then
                              data_inf.Recordset("cl_nomcobr") = "PLAN DE ACCION"
                           End If
                           If data_busca.Recordset("cl_nro_sup") = 4 Then
                              data_inf.Recordset("cl_nomcobr") = "OTRO"
                           End If
                           data_inf.Recordset.Update
                        Else
                        End If
                        
                        data_busca.Recordset.MoveNext
                     Loop
                  Else
                     data_inf.Recordset.Edit
                     data_inf.Recordset("cl_codced") = 8
                     data_inf.Recordset.Update
                  End If
                  
                  data_inf.Recordset.MoveNext
               Loop
               If Combo1.ListIndex = 2 Then
               End If
            End If
            If Combo1.ListIndex = 3 Then ''' conformidad
               Do While Not data_inf.Recordset.EOF
                  data_busca.RecordSource = "Select * from infor_sol where cl_nrovend =" & data_inf.Recordset("cl_codigo") & " and cl_nomcobr =" & 2
                  data_busca.Refresh
                  If data_busca.Recordset.RecordCount > 0 Then
                     data_busca.Recordset.MoveFirst
                     Do While Not data_busca.Recordset.EOF
                        If IsNull(data_busca.Recordset("cl_val3")) = False Then
                           If data_busca.Recordset("cl_val3") = 1 Then
                              If IsNull(data_busca.Recordset("cl_val1")) = False Then
                                 If data_busca.Recordset("cl_val1") = 0 Then
                                    data_inf.Recordset.Edit
                                    data_inf.Recordset("cl_dpto") = "CONFORME"
                                    data_inf.Recordset.Update
                                 Else
                                    If data_busca.Recordset("cl_val1") = 1 Then
                                       data_inf.Recordset.Edit
                                       data_inf.Recordset("cl_dpto") = "CON DEMORA"
                                       data_inf.Recordset.Update
                                    Else
                                       If data_busca.Recordset("cl_val1") = 2 Then
                                          data_inf.Recordset.Edit
                                          data_inf.Recordset("cl_dpto") = "NO CONFORME"
                                          data_inf.Recordset.Update
                                       Else
                                          data_inf.Recordset.Edit
                                          data_inf.Recordset("cl_dpto") = "S/REGISTRAR"
                                          data_inf.Recordset.Update
                                       End If
                                    End If
                                 End If
                              Else
                                 data_inf.Recordset.Edit
                                 data_inf.Recordset("cl_dpto") = "S/REGISTRAR"
                                 data_inf.Recordset.Update
                              End If
                           Else
                              data_inf.Recordset.Edit
                              data_inf.Recordset("cl_codced") = 8
                              data_inf.Recordset.Update
                           End If
                        Else
                           data_inf.Recordset.Edit
                           data_inf.Recordset("cl_codced") = 8
                           data_inf.Recordset.Update
                        End If
                        If IsNull(data_busca.Recordset("cl_nro_sup")) = False Then
                           data_inf.Recordset.Edit
                           If data_busca.Recordset("cl_nro_sup") = 0 Then
                              data_inf.Recordset("cl_nomcobr") = "C.INMEDIATA"
                           End If
                           If data_busca.Recordset("cl_nro_sup") = 1 Then
                              data_inf.Recordset("cl_nomcobr") = "CORRECTIVA"
                           End If
                           If data_busca.Recordset("cl_nro_sup") = 2 Then
                              data_inf.Recordset("cl_nomcobr") = "PREVENTIVA"
                           End If
                           If data_busca.Recordset("cl_nro_sup") = 3 Then
                              data_inf.Recordset("cl_nomcobr") = "PLAN DE ACCION"
                           End If
                           If data_busca.Recordset("cl_nro_sup") = 4 Then
                              data_inf.Recordset("cl_nomcobr") = "OTRO"
                           End If
                           data_inf.Recordset.Update
                        Else
                        End If
                        
                        data_busca.Recordset.MoveNext
                     Loop
                  Else
                     data_inf.Recordset.Edit
                     data_inf.Recordset("cl_codced") = 8
                     data_inf.Recordset.Update
                  End If
                  
                  data_inf.Recordset.MoveNext
               Loop
            End If
         
         End If
         If Combo1.ListIndex = 3 Then
            If data_inf.Recordset.RecordCount > 0 Then
               data_inf.Recordset.MoveFirst
            End If
            Do While Not data_inf.Recordset.EOF
               If IsNull(data_inf.Recordset("cl_cua_vto")) = True Then
                  data_inf.Recordset.Delete
               Else
                  If data_inf.Recordset("cl_cua_vto") = 1 Then
                  Else
                     data_inf.Recordset.Delete
                  End If
               End If
               data_inf.Recordset.MoveNext
            Loop
         End If
         If Combo1.ListIndex = 2 Then
         End If
         
         frm_infmejoras.MousePointer = 0
         MsgBox "Proceso terminado"
         
         data_inf.RecordSource = "Select * from infcli order by cl_fnac"
         data_inf.Refresh
         If Combo1.ListIndex = 0 Then
            If Check1.value = 1 Then
               cr1.ReportFileName = App.Path & "\infmejor3.rpt"
            Else
               If Check2.value = 1 Then
                  cr1.ReportFileName = App.Path & "\infmejor5.rpt"
               Else
                  cr1.ReportFileName = App.Path & "\infmejor1.rpt"
               End If
            End If
            cr1.ReportTitle = "INFORME SOLICITUDES DE MEJORA CONTINUA DESDE: " & md.Text & " HASTA: " & mh.Text
            cr1.Action = 1
         End If
         If Combo1.ListIndex = 1 Then
            If Check1.value = 1 Then
               cr1.ReportFileName = App.Path & "\infmejor3.rpt"
            Else
               If Check2.value = 1 Then
                  cr1.ReportFileName = App.Path & "\infmejor5.rpt"
               Else
                  cr1.ReportFileName = App.Path & "\infmejor1.rpt"
               End If
            End If
            cr1.ReportTitle = "INFORME SOLICITUDES PENDIENTES DESDE: " & md.Text & " HASTA: " & mh.Text
            cr1.Action = 1
         End If
         If Combo1.ListIndex = 2 Then
            If Check1.value = 1 Then
               cr1.ReportFileName = App.Path & "\infmejor3.rpt"
            Else
               If Check2.value = 1 Then
                  cr1.ReportFileName = App.Path & "\infmejor5.rpt"
               Else
                  cr1.ReportFileName = App.Path & "\infmejor1cd.rpt"
               End If
            End If
            cr1.ReportTitle = "INFORME SOLICITUDES CUMPLIDAS SIN CONFORMIDAD DESDE: " & md.Text & " HASTA: " & mh.Text
            cr1.Action = 1
         End If
         If Combo1.ListIndex = 3 Then
            If Check1.value = 1 Then
               cr1.ReportFileName = App.Path & "\infmejor6.rpt"
            Else
               If Check2.value = 1 Then
                  cr1.ReportFileName = App.Path & "\infmejor7.rpt"
               Else
                  cr1.ReportFileName = App.Path & "\infmejor2.rpt"
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
If WElusuario = "BDD" Or WElusuario = "SPEREZ" Or WElusuario = "BRUNO" Or WElusuario = "MEUGENIA" Then
   Label3.Visible = True
   Combo2.Visible = True
    Combo2.Clear
    Combo2.AddItem "TODOS"
    Combo2.AddItem "DIRECTOR GENERAL"
    Combo2.AddItem "GERENTE GENERAL"
    Combo2.AddItem "DIRECCION TECNICA"
    Combo2.AddItem "SUB-DIREC.TECNICA"
    Combo2.AddItem "GERENTE COMERCIAL"
    Combo2.AddItem "JEFE DE MEDICOS DE MOVIL"
    Combo2.AddItem "JEFE CHOFERES Y MANT."
    Combo2.AddItem "JEFE TESORERIA/CONT."
    Combo2.AddItem "JEFE C.COMPUTOS"
    Combo2.AddItem "JEFE BASES Y ENF."
    Combo2.AddItem "JEFE FARMACIA/ECONOMATO"
    Combo2.AddItem "JEFE DESPACHO"
    Combo2.AddItem "JEFE PRUEBAS"
    Combo2.ListIndex = -1
Else
    Label3.Visible = False
    Combo2.Visible = False
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
