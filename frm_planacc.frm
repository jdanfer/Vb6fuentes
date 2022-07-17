VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_planacc 
   BackColor       =   &H00404000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registros de Planes de acción"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9735
   Icon            =   "frm_planacc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   9735
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_his2 
      Caption         =   "data_his2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4920
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_par 
      Caption         =   "data_par"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7080
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CheckBox chcod 
      BackColor       =   &H0080FFFF&
      Caption         =   "Buscar por número"
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
      Left            =   5040
      TabIndex        =   31
      Top             =   4680
      Width           =   2775
   End
   Begin VB.CheckBox chfec 
      BackColor       =   &H0080FFFF&
      Caption         =   "Buscar por fecha"
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
      Left            =   5040
      TabIndex        =   30
      Top             =   4320
      Width           =   2775
   End
   Begin VB.CommandButton b_nover 
      Enabled         =   0   'False
      Height          =   495
      Left            =   5760
      Picture         =   "frm_planacc.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Cancelar la visualización del cuadro descripción."
      Top             =   7200
      Width           =   615
   End
   Begin VB.CommandButton b_ver 
      Height          =   495
      Left            =   4920
      Picture         =   "frm_planacc.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Editar el cuadro DESCRIPCION para leer los datos ingresados."
      Top             =   7200
      Width           =   615
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7320
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Data data_cargo 
      Caption         =   "data_cargo"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6240
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ver solo planes sin cerrar"
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
      Left            =   120
      TabIndex        =   25
      Top             =   4920
      Width           =   3255
   End
   Begin VB.CommandButton b_histo 
      BackColor       =   &H0000FF00&
      Caption         =   "Registrar ACCIONES"
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7320
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_graba 
      Caption         =   "data_graba"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton b_buscafec 
      BackColor       =   &H00C00000&
      Height          =   615
      Left            =   8040
      Picture         =   "frm_planacc.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4320
      Width           =   735
   End
   Begin VB.Data data_accion 
      Caption         =   "data_accion"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6360
      Visible         =   0   'False
      Width           =   2895
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_planacc.frx":1108
      Height          =   2055
      Left            =   120
      OleObjectBlob   =   "frm_planacc.frx":1122
      TabIndex        =   21
      Top             =   5160
      Width           =   9495
   End
   Begin VB.CommandButton b_infor 
      BackColor       =   &H00C00000&
      Height          =   495
      Left            =   3480
      Picture         =   "frm_planacc.frx":1E4D
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Informes"
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton b_cancela 
      BackColor       =   &H00C00000&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      Picture         =   "frm_planacc.frx":228F
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Cancelar movimiento realizado"
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H00C00000&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      Picture         =   "frm_planacc.frx":26D1
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Grabar datos"
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton b_modif 
      BackColor       =   &H00C00000&
      Height          =   495
      Left            =   960
      Picture         =   "frm_planacc.frx":2B13
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Modificar datos de registro seleccionado"
      Top             =   4320
      Width           =   615
   End
   Begin VB.CommandButton b_nuevo 
      BackColor       =   &H00C00000&
      Height          =   495
      Left            =   120
      Picture         =   "frm_planacc.frx":2F55
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Ingresar nuevo registro"
      Top             =   4320
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000080&
      Caption         =   "Datos del plan de acción"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      Begin VB.TextBox t_plan 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   2040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Top             =   2400
         Width           =   7095
      End
      Begin MSMask.MaskEdBox mfecfin 
         Height          =   375
         Left            =   5520
         TabIndex        =   15
         Top             =   3840
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
      Begin VB.ComboBox Combo2 
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
         ItemData        =   "frm_planacc.frx":3397
         Left            =   2040
         List            =   "frm_planacc.frx":33A4
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   3840
         Width           =   3135
      End
      Begin VB.TextBox txt_detal 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   1200
         Width           =   7095
      End
      Begin VB.TextBox txt_encab 
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
         Left            =   2040
         MaxLength       =   60
         TabIndex        =   10
         Top             =   720
         Width           =   7095
      End
      Begin VB.TextBox txt_hora 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   8160
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
      Begin MSMask.MaskEdBox mfecha 
         Height          =   375
         Left            =   5280
         TabIndex        =   4
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   255
         Enabled         =   0   'False
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
      Begin VB.TextBox txt_nro 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   2040
         TabIndex        =   2
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label LABCARGO 
         Height          =   255
         Left            =   6600
         TabIndex        =   32
         Top             =   4080
         Visible         =   0   'False
         Width           =   3615
      End
      Begin VB.Label labid 
         Height          =   255
         Left            =   240
         TabIndex        =   27
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Plan de acción:"
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
         Left            =   120
         TabIndex        =   26
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Cierre:"
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
         Left            =   120
         TabIndex        =   13
         Top             =   3840
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Análisis de causa:"
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
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Título:"
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
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HORA:"
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
         Left            =   7080
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "FECHA:"
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
         Left            =   3960
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "NUMERO:"
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
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Label Label14 
      BackColor       =   &H0080FFFF&
      Caption         =   "Doble click para editar "
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
      Left            =   120
      TabIndex        =   22
      Top             =   7200
      Width           =   4215
   End
   Begin VB.Label labusuario 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1800
      TabIndex        =   8
      Top             =   7440
      Width           =   2535
   End
   Begin VB.Label Label5 
      Caption         =   "Usuario actual:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   7440
      Width           =   1695
   End
End
Attribute VB_Name = "frm_planacc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub b_buscafec_Click()
Dim Xm1, Xm2 As String
If WElusuario = "BDD" Or WElusuario = "BRUNO" Or WElusuario = "SPEREZ" Or WElusuario = "COMPUTOS" Then
   If chfec.Value = 1 Then
      Xm1 = InputBox("INGRESE DESDE QUE FECHA", "FECHA DESDE")
      Xm2 = InputBox("INGRESE HASTA QUE FECHA", "FECHA HASTA")
      If Xm1 <> "" And Xm2 <> "" Then
         data_accion.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 3 & " and cl_fnac >=#" & Format(Xm1, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(Xm2, "yyyy/mm/dd") & "# order by cl_fnac"
         data_accion.Refresh
      Else
         MsgBox "Faltó ingresar fechas"
         data_accion.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 3 & " order by cl_fnac"
         data_accion.Refresh
      End If
   Else
      If chcod.Value = 1 Then
         Xm1 = InputBox("INGRESE CODIGO A BUSCAR", "INGRESE CODIGO")
         If Xm1 <> "" Then
            data_accion.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 3 & " and estado >=" & Val(Xm1) & " order by estado"
            data_accion.Refresh
         Else
            MsgBox "Faltó ingresar CODIGO"
            data_accion.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 3 & " order by estado"
            data_accion.Refresh
         End If
      Else
         data_accion.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 3 & " order by cl_fnac"
         data_accion.Refresh
      End If
   End If
Else
   If chfec.Value = 1 Then
      Xm1 = InputBox("INGRESE DESDE QUE FECHA", "FECHA DESDE")
      Xm2 = InputBox("INGRESE HASTA QUE FECHA", "FECHA HASTA")
      If Xm1 <> "" And Xm2 <> "" Then
         data_accion.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 3 & " and cl_fnac >=#" & Format(Xm1, "yyyy/mm/dd") & "# and cl_fnac <=#" & Format(Xm2, "yyyy/mm/dd") & "# and cl_descpag ='" & WElusuario & "' order by cl_fnac"
         data_accion.Refresh
      Else
         MsgBox "Faltó ingresar fechas"
         data_accion.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 3 & " and cl_descpag ='" & WElusuario & "' order by cl_fnac"
         data_accion.Refresh
      End If
   Else
      If chcod.Value = 1 Then
         Xm1 = InputBox("INGRESE CODIGO A BUSCAR", "INGRESE CODIGO")
         If Xm1 <> "" Then
            data_accion.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 3 & " and estado >=" & Val(Xm1) & " and cl_descpag ='" & WElusuario & "' order by estado"
            data_accion.Refresh
         Else
            MsgBox "Faltó ingresar CODIGO"
            data_accion.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 3 & " and cl_descpag ='" & WElusuario & "' order by estado"
            data_accion.Refresh
         End If
      Else
         data_accion.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 3 & " and cl_descpag ='" & WElusuario & "' order by cl_fnac"
         data_accion.Refresh
      End If
   End If
End If
DBGrid1.SetFocus

End Sub

Private Sub b_cancela_Click()
'If XAlta = 1 Then
'   data_graba.Recordset.CancelUpdate
'End If
b_nuevo.Enabled = True
b_modif.Enabled = True
b_graba.Enabled = False
b_cancela.Enabled = False
b_buscafec.Enabled = True
b_infor.Enabled = True
DBGrid1.Enabled = True
borracamp
Frame1.Enabled = False

End Sub


Private Sub b_graba_Click()
Dim XXdest As Long
Dim Xelnro As Double
Xelnro = txt_nro.Text
XXdest = 0
If XAlta = 1 Then
   If Len(txt_encab.Text) > 5 Then
      If Len(txt_detal.Text) > 5 Then
         data_graba.Recordset.AddNew
         data_graba.Recordset("cl_etiquet") = 0
         data_graba.Recordset("cl_val2") = 7
         data_graba.Recordset("cl_codigo") = labid.Caption 'correlativo
         data_graba.Recordset("estado") = Xelnro
         data_graba.Recordset("cl_fnac") = mfecha.Text
         data_graba.Recordset("cl_nomcobr") = 3
         data_graba.Recordset("cl_ruc") = txt_hora.Text
'         data_cargo.Recordset.FindFirst "chofer ='" & List1.List(List1.ListIndex) & "'"
         data_graba.Recordset("cl_nom_sup") = Mid(WElusuario, 1, 25)
         data_cargo.Recordset.FindFirst "medico ='" & WElusuario & "'"
         If Not data_cargo.Recordset.NoMatch Then
            data_graba.Recordset("cl_desc2") = data_cargo.Recordset("medico")
         Else
            data_graba.Recordset("cl_desc2") = WElusuario
         End If
         data_graba.Recordset("cl_descpag") = labusuario.Caption
         data_graba.Recordset("cl_desc1") = txt_encab.Text
         data_graba.Recordset("info_debit") = txt_detal.Text
         If Combo2.ListIndex >= 0 Then
            data_graba.Recordset("cl_val1") = Combo2.ListIndex
         Else
            data_graba.Recordset("cl_val1") = -1
         End If
         If mfecfin.Text <> "__/__/____" Then
            data_graba.Recordset("cl_fultmov") = mfecfin.Text
         Else
         End If
         data_graba.Recordset("cl_codconv") = "X"
         data_graba.Recordset.Update
         If t_plan.Text <> "" Then
            labid.Caption = labid.Caption + 1
            data_graba.Recordset.AddNew
            data_graba.Recordset("cl_nrovend") = txt_nro.Text
            data_graba.Recordset("cl_etiquet") = 0
            data_graba.Recordset("estado") = 98
            data_graba.Recordset("cl_codigo") = labid.Caption
            data_graba.Recordset("info_debit") = t_plan.Text
            data_graba.Recordset.Update
         Else
            labid.Caption = labid.Caption + 1
            data_graba.Recordset.AddNew
            data_graba.Recordset("cl_nrovend") = labnro.Caption
            data_graba.Recordset("cl_etiquet") = 0
            data_graba.Recordset("estado") = 98
            data_graba.Recordset("cl_codigo") = labid.Caption
            data_graba.Recordset.Update
         End If
         b_nuevo.Enabled = True
         b_modif.Enabled = True
         b_graba.Enabled = False
         b_cancela.Enabled = False
         b_buscafec.Enabled = True
         b_infor.Enabled = True
         DBGrid1.Enabled = True
         Frame1.Enabled = False
         data_graba.Refresh
         data_accion.Refresh
         borracamp
         XAlta = 0
      Else
         MsgBox "Ingrese detalles"
      End If
   Else
      MsgBox "Ingrese título"
   End If
Else
   data_graba.Recordset.Edit
   
   data_graba.Recordset("cl_desc1") = txt_encab.Text
   data_graba.Recordset("info_debit") = txt_detal.Text
   If Combo2.ListIndex >= 0 Then
      data_graba.Recordset("cl_val1") = Combo2.ListIndex
   Else
      data_graba.Recordset("cl_val1") = -1
   End If
   If mfecfin.Text <> "__/__/____" Then
      data_graba.Recordset("cl_fultmov") = mfecfin.Text
   Else
   End If
   data_graba.Recordset.Update
   data_his2.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 3 & " and cl_nrovend =" & labnro.Caption & " and estado =" & 98
   data_his2.Refresh
   If data_his2.Recordset.RecordCount > 0 Then
      If IsNull(data_his2.Recordset("info_debit")) = False Then
         If data_his2.Recordset("info_debit") = t_plan.Text Then
         Else
            data_his2.Recordset.Edit
            data_his2.Recordset("info_debit") = t_plan.Text
            data_his2.Recordset.Update
         End If
      Else
         If t_plan.Text <> "" Then
            data_his2.Recordset.Edit
            data_his2.Recordset("info_debit") = t_plan.Text
            data_his2.Recordset.Update
         End If
      End If
   Else
'      labidd.Caption = labidd.Caption + 1
      data_his2.Recordset.AddNew
      data_his2.Recordset("cl_nrovend") = txt_nro.Text
      data_his2.Recordset("cl_etiquet") = 0
      data_his2.Recordset("estado") = 98
      data_his2.Recordset("cl_codigo") = labid.Caption
      If t_plan.Text <> "" Then
         data_his2.Recordset("info_debit") = t_plan.Text
      End If
      data_his2.Recordset.Update
   End If
   
   b_nuevo.Enabled = True
   b_modif.Enabled = True
   b_graba.Enabled = False
   b_cancela.Enabled = False
   b_buscafec.Enabled = True
   b_infor.Enabled = True
   DBGrid1.Enabled = True
   Frame1.Enabled = False
   data_graba.Refresh
   data_accion.Refresh
   borracamp
End If


End Sub

Private Sub b_histo_Click()
''frm_mejoraconsi.Show vbModal

End Sub

Private Sub b_infor_Click()
frm_infplanes.Show vbModal

End Sub

Private Sub b_modif_Click()
XAlta = 0
Frame1.Enabled = True
If Combo2.ListIndex >= 0 Then
   MsgBox "ATENCION! EL REGISTRO YA FUE CERRADO", vbInformation, "Mejora continua"
   Frame1.Enabled = False
Else
   b_nuevo.Enabled = False
   b_modif.Enabled = False
   b_graba.Enabled = True
   b_cancela.Enabled = True
   b_buscafec.Enabled = False
   b_infor.Enabled = False
   DBGrid1.Enabled = False
   borracamp
   data_graba.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 3 & " and estado =" & data_accion.Recordset("estado")
   data_graba.Refresh
   If data_graba.Recordset.RecordCount > 0 Then
      If IsNull(data_graba.Recordset("cl_val3")) = True Then
         Combo2.Enabled = False
         mfecfin.Enabled = False
      Else
         If data_graba.Recordset("cl_val3") = 1 Then
            Combo2.Enabled = True
            mfecfin.Enabled = True
         Else
            Combo2.Enabled = False
            mfecfin.Enabled = False
         End If
      End If
      igualaacc
   Else
      Frame1.Enabled = False
      b_nuevo.Enabled = True
      b_modif.Enabled = True
      b_graba.Enabled = False
      b_cancela.Enabled = False
      b_buscafec.Enabled = True
      b_infor.Enabled = True
      DBGrid1.Enabled = True
      Combo2.Enabled = False
      mfecfin.Enabled = False
   End If
   Combo2.Enabled = True
   mfecfin.Enabled = True
End If

End Sub

Private Sub b_nover_Click()
b_nuevo.Enabled = True
b_modif.Enabled = True
b_graba.Enabled = False
b_cancela.Enabled = False
b_buscafec.Enabled = True
b_infor.Enabled = True
DBGrid1.Enabled = True
txt_encab.Enabled = True
txt_detal.Enabled = True
t_plan.Enabled = True
Combo2.Enabled = True
mfecfin.Enabled = True
Check1.Enabled = True
b_ver.Enabled = True
b_nover.Enabled = False
Frame1.Enabled = False

End Sub

Private Sub b_nuevo_Click()
XAlta = 1
b_nuevo.Enabled = False
b_modif.Enabled = False
b_graba.Enabled = True
b_cancela.Enabled = True
b_buscafec.Enabled = False
b_infor.Enabled = False
DBGrid1.Enabled = False
Frame1.Enabled = True
borracamp
Data1.DatabaseName = App.Path & "\sapp.mdb"
Data1.RecordSource = "Select * from infor_sol order by cl_codigo"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveLast
   labid.Caption = Data1.Recordset("cl_codigo") + 1
Else
   labid.Caption = 1000
End If
txt_nro.Text = data_par.Recordset("prod_gral") + 1
mfecha.Text = Format(Date, "dd/mm/yyyy")
txt_hora.Text = Format(Time, "HH:mm")
Combo2.Enabled = False
mfecfin.Enabled = False

End Sub

Private Sub b_ver_Click()
b_nuevo.Enabled = False
b_modif.Enabled = False
b_graba.Enabled = False
b_cancela.Enabled = False
b_buscafec.Enabled = False
b_infor.Enabled = False
DBGrid1.Enabled = False
Frame1.Enabled = True
txt_encab.Enabled = False
txt_detal.Enabled = True
t_plan.Enabled = True
Combo2.Enabled = False
mfecfin.Enabled = False
Check1.Enabled = False
b_ver.Enabled = False
b_nover.Enabled = True


End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
   If WElusuario = "BDD" Or WElusuario = "BRUNO" Or WElusuario = "SPEREZ" Then
      data_accion.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 3 & " order by cl_fnac"
      data_accion.Refresh
   Else
      data_accion.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 3 & " and cl_descpag ='" & WElusuario & "' order by cl_fnac DESC"
      data_accion.Refresh
   End If
Else
   If WElusuario = "BDD" Or WElusuario = "BRUNO" Or WElusuario = "SPEREZ" Then
      data_accion.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 3 & " order by cl_fnac"
      data_accion.Refresh
   Else
      data_accion.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 3 & " and cl_descpag ='" & WElusuario & "' order by cl_fnac DESC"
      data_accion.Refresh
   End If
End If

End Sub



Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_graba.SetFocus
End If

End Sub

Private Sub Combo3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Combo2.Enabled = True Then
      Combo2.SetFocus
   Else
      b_graba.SetFocus
   End If
End If

End Sub

Private Sub DBGrid1_DblClick()
borracamp
igualaacc

End Sub

Private Sub Form_Load()
data_accion.DatabaseName = App.Path & "\sapp.mdb"
If WElusuario = "BDD" Or WElusuario = "BRUNO" Or WElusuario = "SPEREZ" Then
   data_accion.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 3 & " order by cl_fnac"
   data_accion.Refresh
Else
   data_accion.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 3 & " and cl_descpag ='" & WElusuario & "' order by cl_fnac DESC"
   data_accion.Refresh
End If

data_graba.DatabaseName = App.Path & "\sapp.mdb"
data_graba.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 3 & " order by estado"
data_graba.Refresh
data_his2.DatabaseName = App.Path & "\sapp.mdb"
data_his2.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 3 & " order by estado"
data_his2.Refresh

data_cargo.DatabaseName = App.Path & "\sapp.mdb"
data_cargo.RecordSource = "movil"
data_cargo.Refresh

data_par.DatabaseName = App.Path & "\parse.mdb"
data_par.RecordSource = "parsec0"
data_par.Refresh

labusuario.Caption = WElusuario

End Sub


Public Function borracamp()
txt_nro.Text = ""
mfecha.Text = "__/__/____"
txt_hora.Text = ""
txt_encab.Text = ""
txt_detal.Text = ""
mfecfin.Enabled = True
mfecfin.Text = "__/__/____"
mfecfin.Enabled = False
Combo2.Enabled = True
Combo2.ListIndex = -1
Combo2.Enabled = False
t_plan.Text = ""

End Function

Private Sub txt_encab_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_detal.SetFocus
End If

End Sub

Public Function igualaacc()
If data_accion.Recordset.RecordCount > 0 Then
    If IsNull(data_accion.Recordset("estado")) = False Then
       txt_nro.Text = data_accion.Recordset("estado")
    Else
       txt_nro.Text = 0
    End If
    If IsNull(data_accion.Recordset("cl_fnac")) = False Then
       mfecha.Text = data_accion.Recordset("cl_fnac")
    Else
       mfecha.Text = "__/__/____"
    End If
    If IsNull(data_accion.Recordset("cl_ruc")) = False Then
       txt_hora.Text = data_accion.Recordset("cl_ruc")
    Else
       txt_hora.Text = ""
    End If
    If IsNull(data_accion.Recordset("cl_desc2")) = False Then
       LABCARGO.Caption = data_accion.Recordset("cl_desc2")
    Else
       LABCARGO.Caption = WElusuario
    End If
    If IsNull(data_accion.Recordset("cl_desc1")) = False Then
       txt_encab.Text = data_accion.Recordset("cl_desc1")
    Else
       txt_encab.Text = ""
    End If
    If IsNull(data_accion.Recordset("info_debit")) = False Then
       txt_detal.Text = data_accion.Recordset("info_debit")
    Else
       txt_detal.Text = ""
    End If
    If IsNull(data_accion.Recordset("cl_val1")) = False Then
       Combo2.Enabled = True
       Combo2.ListIndex = data_accion.Recordset("cl_val1")
       Combo2.Enabled = False
    Else
       Combo2.Enabled = True
       Combo2.ListIndex = -1
       Combo2.Enabled = False
    End If
    If IsNull(data_accion.Recordset("cl_fultmov")) = False Then
       mfecfin.Enabled = True
       mfecfin.Text = Format(data_accion.Recordset("cl_fultmov"), "dd/mm/yyyy")
       mfecfin.Enabled = False
    Else
       mfecfin.Enabled = True
       mfecfin.Text = "__/__/____"
       mfecfin.Enabled = False
    End If
    data_his2.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 3 & " and cl_nrovend =" & txt_nro.Text & " and estado =" & 98
    data_his2.Refresh
    If data_his2.Recordset.RecordCount > 0 Then
       If IsNull(data_his2.Recordset("info_debit")) = False Then
          t_plan.Text = data_his2.Recordset("info_debit")
       Else
          t_plan.Text = ""
       End If
    Else
       t_plan.Text = ""
    End If

End If

End Function
