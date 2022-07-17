VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_actasrr 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registros de Actas de reunión"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9735
   Icon            =   "frm_actasrr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   9735
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   9000
      TabIndex        =   38
      Top             =   5160
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Data data_graba2 
      Caption         =   "data_graba2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5400
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data data_par 
      Caption         =   "data_par"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7440
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Buscar por título"
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
      TabIndex        =   34
      Top             =   5400
      Width           =   2895
   End
   Begin VB.CheckBox Check1 
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
      TabIndex        =   33
      Top             =   5040
      Width           =   2895
   End
   Begin VB.CommandButton b_nover 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   9000
      Picture         =   "frm_actasrr.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Cancelar la visualización del cuadro descripción."
      Top             =   7320
      Width           =   615
   End
   Begin VB.CommandButton b_ver 
      Height          =   495
      Left            =   7920
      Picture         =   "frm_actasrr.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   31
      ToolTipText     =   "Editar el cuadro DESCRIPCION para leer los datos ingresados."
      Top             =   7320
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
      Top             =   6960
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
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6720
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data data_graba 
      Caption         =   "data_graba"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton b_buscafec 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8040
      Picture         =   "frm_actasrr.frx":0F56
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5040
      Width           =   615
   End
   Begin VB.Data data_accion 
      Caption         =   "data_accion"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7200
      Visible         =   0   'False
      Width           =   2895
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_actasrr.frx":14E0
      Height          =   1695
      Left            =   120
      OleObjectBlob   =   "frm_actasrr.frx":14FA
      TabIndex        =   26
      Top             =   5640
      Width           =   9495
   End
   Begin VB.CommandButton b_infor 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3960
      Picture         =   "frm_actasrr.frx":2229
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Informes"
      Top             =   5040
      Width           =   615
   End
   Begin VB.CommandButton b_cancela 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   3000
      Picture         =   "frm_actasrr.frx":27B3
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Cancelar movimiento realizado"
      Top             =   5040
      Width           =   615
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2040
      Picture         =   "frm_actasrr.frx":2D3D
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Grabar datos"
      Top             =   5040
      Width           =   615
   End
   Begin VB.CommandButton b_modif 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1080
      Picture         =   "frm_actasrr.frx":32C7
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Modificar datos de registro seleccionado"
      Top             =   5040
      Width           =   615
   End
   Begin VB.CommandButton b_nuevo 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      Picture         =   "frm_actasrr.frx":3851
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Ingresar nuevo registro"
      Top             =   5040
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Datos del Acta"
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
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9495
      Begin VB.Data data_carga2 
         Caption         =   "data_carga2"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   7080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4680
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H0080FFFF&
         Caption         =   "Acta de reunión PRIVADA"
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
         TabIndex        =   35
         Top             =   4680
         Width           =   3735
      End
      Begin MSMask.MaskEdBox mfecfin 
         Height          =   375
         Left            =   5400
         TabIndex        =   20
         Top             =   4080
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
         ItemData        =   "frm_actasrr.frx":3DDB
         Left            =   2040
         List            =   "frm_actasrr.frx":3DE8
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   4080
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
         Height          =   1455
         Left            =   2040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   2520
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
         TabIndex        =   15
         Top             =   2040
         Width           =   7095
      End
      Begin VB.CommandButton b_elimin 
         Height          =   495
         Left            =   5160
         Picture         =   "frm_actasrr.frx":3E0F
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Elimina destinatario seleccionado"
         Top             =   1320
         Width           =   735
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   6120
         TabIndex        =   12
         Top             =   960
         Width           =   3015
      End
      Begin VB.CommandButton b_agreg 
         Height          =   495
         Left            =   5160
         Picture         =   "frm_actasrr.frx":4251
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Agrega..."
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
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
         ItemData        =   "frm_actasrr.frx":4693
         Left            =   2040
         List            =   "frm_actasrr.frx":4695
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   720
         Width           =   3015
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
      Begin VB.Label labpriv 
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   3480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label labnro2 
         Height          =   255
         Left            =   480
         TabIndex        =   36
         Top             =   1680
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label labid 
         Height          =   375
         Left            =   240
         TabIndex        =   30
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Jefes participantes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6120
         TabIndex        =   29
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C00000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   4200
         Width           =   1935
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C00000&
         Caption         =   "Descripción:"
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
         TabIndex        =   16
         Top             =   2520
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C00000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C00000&
         Caption         =   "Participantes:"
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
         TabIndex        =   9
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7080
         TabIndex        =   5
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3960
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
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
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1935
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
      TabIndex        =   27
      Top             =   7320
      Width           =   3495
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
      Left            =   2760
      TabIndex        =   8
      Top             =   7560
      Width           =   2895
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
      Top             =   7560
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   6480
      Picture         =   "frm_actasrr.frx":4697
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   2295
   End
End
Attribute VB_Name = "frm_actasrr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_agreg_Click()
Dim XX, Xban As Long
XX = 0
Xban = 0
If List1.ListCount >= 1 Then
   For XX = 1 To List1.ListCount
       List1.ListIndex = XX - 1
       If List1.List(List1.ListIndex) = Combo1.Text Then
          Xban = 1
       End If
   Next
Else
   Xban = 0
End If

If Combo1.ListIndex >= 0 And Xban <> 1 Then
   List1.AddItem Combo1.Text

End If

End Sub

Private Sub b_buscafec_Click()
Dim Xm1 As String
If Check1.value = 1 Then
   Xm1 = InputBox("INGRESE NUMERO DE ACTA A BUSCAR")
   If Xm1 <> "" Then
      data_accion.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 4 & " and cl_numero =" & Val(Xm1) & " and (cl_nom_sup ='" & "TODOS" & "' or cl_nom_sup ='" & WElusuario & "') and estado <>" & 97 & " order by estado"
      data_accion.Refresh
   Else
      MsgBox "Faltó ingresar número"
      data_accion.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 4 & " and (cl_nom_sup ='" & "TODOS" & "' or cl_nom_sup ='" & WElusuario & "') and estado <>" & 97 & " order by estado"
      data_accion.Refresh
   End If
Else
   If Check2.value = 1 Then
      Xm1 = InputBox("INGRESE TEXTO A BUSCAR EN EL TITULO")
      Xm1 = UCase(Xm1)
      If Xm1 <> "" Then
         data_accion.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 4 & " and cl_desc1 like '*" & Xm1 & "*' and estado <>" & 97 & " and (cl_nom_sup ='" & "TODOS" & "' or cl_nom_sup ='" & WElusuario & "') order by cl_fnac"
         data_accion.Refresh
      Else
         MsgBox "Faltó ingresar PALABRA"
         data_accion.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 4 & " and (cl_nom_sup ='" & "TODOS" & "' or cl_nom_sup ='" & WElusuario & "') and estado <>" & 97 & " order by estado"
         data_accion.Refresh
      End If
   Else
      data_accion.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 4 & " and (cl_nom_sup ='" & "TODOS" & "' or cl_nom_sup ='" & WElusuario & "') and estado <>" & 97 & " order by estado"
      data_accion.Refresh
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

Private Sub b_elimin_Click()
If List1.ListIndex >= 0 Then
   List1.RemoveItem List1.ListIndex
End If

End Sub

Private Sub b_graba_Click()
Dim XXdest As Long
Dim Xelnro As Double
Xelnro = labnro2.Caption
XXdest = 0
If labpriv.Caption = "" Then
   labpriv.Caption = 0
End If

If labpriv.Caption = 1 And Check3.value <> 1 Then
   MsgBox "ATENCIÓN!! SE CAMBIO EL ACTA DE PRIVADA A PUBLICA", vbCritical, "ACTAS"
End If
If XAlta = 1 Then
   If List1.ListCount >= 1 Then
      If Len(txt_encab.Text) > 5 Then
         If Len(txt_detal.Text) > 5 Then
            If Check3.value = 1 Then
               List1.ListIndex = 0
               For XXdest = 0 To List1.ListCount - 1
                   data_graba.Recordset.AddNew
                   data_cargo.RecordSource = "Select * from movil where chofer ='" & List1.List(List1.ListIndex) & "'"
                   data_cargo.Refresh
'                   data_cargo.Recordset.FindFirst "chofer ='" & List1.List(List1.ListIndex) & "'"
                   If data_cargo.Recordset.RecordCount > 0 Then
                      data_graba.Recordset("cl_nom_sup") = Mid(data_cargo.Recordset("medico"), 1, 25)
                   Else
                      data_graba.Recordset("cl_nom_sup") = WElusuario
                   End If
                   data_graba.Recordset("cl_etiquet") = 0
                   data_graba.Recordset("cl_val2") = 7
                   data_graba.Recordset("cl_codigo") = labid.Caption
                   data_graba.Recordset("estado") = Xelnro ' número real para busqueda
                   data_graba.Recordset("cl_fnac") = mfecha.Text
                   data_graba.Recordset("cl_ruc") = txt_hora.Text
                   data_graba.Recordset("cl_nomcobr") = 4
                   data_graba.Recordset("cl_numero") = txt_nro.Text
                   data_graba.Recordset("cl_zona") = Check3.value
                   data_graba.Recordset("cl_descpag") = labusuario.Caption
                   data_graba.Recordset("cl_desc2") = List1.List(List1.ListIndex)
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
                   data_graba.Recordset("cl_codconv") = "T"
                   data_graba.Recordset.Update
                   data_par.Recordset.Edit
                   data_par.Recordset("limite_dia") = data_par.Recordset("limite_dia") + 1
                   data_par.Recordset.Update
                   data_par.Refresh
                   Xelnro = Xelnro + 1
                   If labid.Caption <> "" Then
                      labid.Caption = labid.Caption + 1
                   End If
                   If List1.ListCount - 1 = List1.ListIndex Then
                   Else
                      List1.ListIndex = List1.ListIndex + 1
                   End If
               
               Next
            Else
               data_graba.Recordset.AddNew
               'data_cargo.Recordset.FindFirst "chofer ='" & List1.List(List1.ListIndex) & "'"
               'If Not data_cargo.Recordset.NoMatch Then
               '   data_graba.Recordset("cl_nom_sup") = Mid(data_cargo.Recordset("medico"), 1, 25)
               'Else
               data_graba.Recordset("cl_nom_sup") = "TODOS"
               data_graba.Recordset("cl_etiquet") = 0
               data_graba.Recordset("cl_val2") = 7
               data_graba.Recordset("cl_codigo") = labid.Caption
               data_graba.Recordset("estado") = Xelnro ' número real para busqueda
               data_graba.Recordset("cl_fnac") = mfecha.Text
               data_graba.Recordset("cl_ruc") = txt_hora.Text
               data_graba.Recordset("cl_nomcobr") = 4
               data_graba.Recordset("cl_numero") = txt_nro.Text
               data_graba.Recordset("cl_zona") = Check3.value
               data_graba.Recordset("cl_descpag") = labusuario.Caption
               data_cargo.Recordset.FindFirst "medico ='" & WElusuario & "'"
               If Not data_cargo.Recordset.NoMatch Then
                  data_graba.Recordset("cl_desc2") = data_cargo.Recordset("chofer")
               Else
                  data_graba.Recordset("cl_desc2") = WElusuario
               End If
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
               data_graba.Recordset("cl_codconv") = "T"
               data_graba.Recordset.Update
               data_par.Recordset.Edit
               data_par.Recordset("limite_dia") = data_par.Recordset("limite_dia") + 1
               data_par.Recordset.Update
               data_par.Refresh
               Xelnro = Xelnro + 1
               If labid.Caption <> "" Then
                  labid.Caption = labid.Caption + 1
               End If
               List1.ListIndex = 0
               For XXdest = 0 To List1.ListCount - 1
                   data_cargo.RecordSource = "Select * from movil where chofer ='" & List1.List(List1.ListIndex) & "'"
                   data_cargo.Refresh
'                   data_cargo.Recordset.FindFirst "chofer ='" & List1.List(List1.ListIndex) & "'"
                   If data_cargo.Recordset.RecordCount > 0 Then
                      If WElusuario = Mid(data_cargo.Recordset("medico"), 1, 25) Then
                         If List1.ListCount - 1 = List1.ListIndex Then
                         Else
                            List1.ListIndex = List1.ListIndex + 1
                         End If
                      Else
                         data_graba.Recordset.AddNew
                         data_graba.Recordset("cl_nom_sup") = WElusuario
                         data_graba.Recordset("cl_etiquet") = 0
                         data_graba.Recordset("cl_val2") = 7
                         data_graba.Recordset("cl_codigo") = labid.Caption
                         data_graba.Recordset("estado") = 97
                         data_graba.Recordset("cl_fnac") = mfecha.Text
                         data_graba.Recordset("cl_ruc") = txt_hora.Text
                         data_graba.Recordset("cl_nomcobr") = 4
                         data_graba.Recordset("cl_numero") = txt_nro.Text
                         data_graba.Recordset("cl_nom_sup") = Mid(data_cargo.Recordset("medico"), 1, 25)
                         data_graba.Recordset("cl_descpag") = labusuario.Caption
                         data_graba.Recordset("cl_desc2") = List1.List(List1.ListIndex)
                         data_graba.Recordset("cl_codconv") = "T"
                         data_graba.Recordset.Update
                         data_par.Recordset.Edit
                         data_par.Recordset("limite_dia") = data_par.Recordset("limite_dia") + 1
                         data_par.Recordset.Update
                         data_par.Refresh
                         Xelnro = Xelnro + 1
                         If labid.Caption <> "" Then
                            labid.Caption = labid.Caption + 1
                         End If
                         If List1.ListCount - 1 = List1.ListIndex Then
                         Else
                            List1.ListIndex = List1.ListIndex + 1
                         End If
                      End If
                   End If
               Next
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
      MsgBox "Ingrese al menos un destinatario"
   End If
Else
   If Check3.value = 1 Then
      List1.ListIndex = 0
      For XXdest = 0 To List1.ListCount - 1
          data_graba2.RecordSource = "Select * from infor_sol where cl_desc2 ='" & List1.List(List1.ListIndex) & "' and cl_numero =" & txt_nro.Text & " and cl_nomcobr =" & 4
          data_graba2.Refresh
          If data_graba2.Recordset.RecordCount > 0 Then
             data_graba2.Recordset.Edit
'             data_graba2.Recordset("cl_nom_sup") = WElusuario
             data_graba2.Recordset("cl_descpag") = labusuario.Caption
             data_graba2.Recordset("cl_desc2") = List1.List(List1.ListIndex)
             data_graba2.Recordset("cl_desc1") = txt_encab.Text
             data_graba2.Recordset("info_debit") = txt_detal.Text
             If Combo2.ListIndex >= 0 Then
                data_graba2.Recordset("cl_val1") = Combo2.ListIndex
             Else
                data_graba2.Recordset("cl_val1") = -1
             End If
             If mfecfin.Text <> "__/__/____" Then
                data_graba2.Recordset("cl_fultmov") = mfecfin.Text
             Else
                data_graba2.Recordset("cl_fultmov") = Null
             End If
             data_graba2.Recordset("cl_zona") = Check3.value
             data_graba2.Recordset.Update
             If List1.ListCount - 1 = List1.ListIndex Then
             Else
                List1.ListIndex = List1.ListIndex + 1
             End If
          Else
             data_graba2.Recordset.AddNew
             data_cargo.RecordSource = "Select * from movil where chofer ='" & List1.List(List1.ListIndex) & "'"
             data_cargo.Refresh
'             data_cargo.Recordset.FindFirst "chofer ='" & List1.List(List1.ListIndex) & "'"
             If data_cargo.Recordset.RecordCount > 0 Then
                data_graba2.Recordset("cl_nom_sup") = Mid(data_cargo.Recordset("medico"), 1, 25)
             Else
                data_graba2.Recordset("cl_nom_sup") = WElusuario
             End If
             data_graba2.Recordset("cl_etiquet") = 0
             data_graba2.Recordset("cl_val2") = 7
             data_graba2.Recordset("cl_codigo") = labid.Caption
             data_graba2.Recordset("estado") = data_par.Recordset("limite_dia") + 1 ' número real para busqueda
             data_graba2.Recordset("cl_fnac") = mfecha.Text
             data_graba2.Recordset("cl_ruc") = txt_hora.Text
             data_graba2.Recordset("cl_nomcobr") = 4
             data_graba2.Recordset("cl_numero") = txt_nro.Text
             data_graba2.Recordset("cl_zona") = Check3.value
             data_graba2.Recordset("cl_descpag") = labusuario.Caption
             data_graba2.Recordset("cl_desc2") = List1.List(List1.ListIndex)
             data_graba2.Recordset("cl_desc1") = txt_encab.Text
             data_graba2.Recordset("info_debit") = txt_detal.Text
             If Combo2.ListIndex >= 0 Then
                data_graba2.Recordset("cl_val1") = Combo2.ListIndex
             Else
                data_graba2.Recordset("cl_val1") = -1
             End If
             If mfecfin.Text <> "__/__/____" Then
                data_graba2.Recordset("cl_fultmov") = mfecfin.Text
             Else
                                           
             End If
             data_graba2.Recordset("cl_codconv") = "T"
             data_graba2.Recordset.Update
             If List1.ListCount - 1 = List1.ListIndex Then
             Else
                List1.ListIndex = List1.ListIndex + 1
             End If
             data_par.Recordset.Edit
             data_par.Recordset("limite_dia") = data_par.Recordset("limite_dia") + 1
             data_par.Recordset.Update
             data_par.Refresh
          End If
      Next
   Else
      data_graba.Recordset.Edit
'      List1.ListIndex = 0
      data_graba.Recordset("cl_nom_sup") = "TODOS"
'      data_graba.Recordset("cl_descpag") = labusuario.Caption
'      data_graba.Recordset("cl_desc2") = List1.List(List1.ListIndex)
        
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
         data_graba.Recordset("cl_fultmov") = Null
      End If
      data_graba.Recordset("cl_zona") = Check3.value
      data_graba.Recordset.Update
      List1.ListIndex = 0
      For XXdest = 0 To List1.ListCount - 1
          data_graba2.RecordSource = "Select * from infor_sol where cl_desc2 ='" & List1.List(List1.ListIndex) & "' and cl_numero =" & txt_nro.Text & " and cl_zona =" & 1
          data_graba2.Refresh
          If data_graba2.Recordset.RecordCount > 0 Then
             data_graba2.Recordset.Delete
          End If
          data_graba2.RecordSource = "Select * from infor_sol where cl_desc2 ='" & List1.List(List1.ListIndex) & "' and estado =" & 97 & " and cl_numero =" & txt_nro.Text
          data_graba2.Refresh
          If data_graba2.Recordset.RecordCount > 0 Then
          '''''data_cargo.Recordset.FindFirst "chofer ='" & List1.List(List1.ListIndex) & "'"
          Else
             data_graba2.RecordSource = "Select * from infor_sol where cl_desc2 ='" & List1.List(List1.ListIndex) & "' and cl_nom_sup ='" & "TODOS" & "' and cl_numero =" & txt_nro.Text
             data_graba2.Refresh
             If data_graba2.Recordset.RecordCount > 0 Then
             Else
                Data1.RecordSource = "Select * from infor_sol order by cl_codigo"
                Data1.Refresh
                If Data1.Recordset.RecordCount > 0 Then
                   Data1.Recordset.MoveLast
                   labid.Caption = Data1.Recordset("cl_codigo") + 1
                Else
                   labid.Caption = 1000
                End If
                data_graba2.Recordset.AddNew
                data_graba2.Recordset("cl_nom_sup") = WElusuario
                data_graba2.Recordset("cl_etiquet") = 0
                data_graba2.Recordset("cl_val2") = 7
                data_graba2.Recordset("cl_codigo") = labid.Caption
                data_graba2.Recordset("estado") = 97
                data_graba2.Recordset("cl_fnac") = mfecha.Text
                data_graba2.Recordset("cl_ruc") = txt_hora.Text
                data_graba2.Recordset("cl_nomcobr") = 4
                data_graba2.Recordset("cl_numero") = txt_nro.Text
                data_graba2.Recordset("cl_nom_sup") = Mid(data_cargo.Recordset("medico"), 1, 25)
                data_graba2.Recordset("cl_descpag") = labusuario.Caption
                data_graba2.Recordset("cl_desc2") = List1.List(List1.ListIndex)
                data_graba2.Recordset("cl_codconv") = "T"
                data_graba2.Recordset.Update
             End If
          End If
          If List1.ListCount - 1 = List1.ListIndex Then
          Else
             List1.ListIndex = List1.ListIndex + 1
          End If
      Next
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
labpriv.Caption = 0



End Sub



Private Sub b_infor_Click()
If WElusuario = "BDD" Or WElusuario = "SPEREZ" Or WElusuario = "JFERNAN" Or WElusuario = "BRUNO" Then
   frm_infactas.Show vbModal
Else
   MsgBox "Usuario no habilitado para informes"
End If

End Sub

Private Sub b_modif_Click()
labpriv.Caption = Check3.value

If WElusuario = "SPEREZ" Or WElusuario = "COMPUTOS" Or WElusuario = "BDD" Or WElusuario = "BRUNO" Then
   'If labusuario.Caption = data_accion.Recordset("cl_descpag") Then
    XAlta = 0
    Frame1.Enabled = True
    Combo2.Enabled = True
    mfecfin.Enabled = True
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
        data_graba.RecordSource = "Select * from infor_sol where cl_numero =" & data_accion.Recordset("cl_numero") & " and cl_nomcobr =" & 4 & " and cl_codigo =" & data_accion.Recordset("cl_codigo")
        data_graba.Refresh
        igualaacc
        Combo2.Enabled = True
        mfecfin.Enabled = True
    
    End If
Else
   MsgBox "Usuario no autorizado para modificación, solo puede editar el cuadro Descripción para ver"
'   Frame1.Enabled = True

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
Combo1.Enabled = True
List1.Enabled = True
txt_encab.Enabled = True
txt_detal.Enabled = True
Combo2.Enabled = True
mfecfin.Enabled = True
Check1.Enabled = True
b_ver.Enabled = True
b_nover.Enabled = False
'Check3.Enabled = False
Frame1.Enabled = False

End Sub

Private Sub b_nuevo_Click()
If WElusuario = "SPEREZ" Or WElusuario = "COMPUTOS" Or WElusuario = "BDD" Or WElusuario = "BRUNO" Then
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
'    Data1.DatabaseName = App.Path & "\sapp.mdb"
    Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
    Data1.RecordSource = "Select * from infor_sol order by cl_codigo DESC"
    Data1.Refresh
    If Data1.Recordset.RecordCount > 0 Then
'       Data1.Recordset.MoveLast
       labid.Caption = Data1.Recordset("cl_codigo") + 1
    Else
       labid.Caption = 1000
    End If
'    Data1.DatabaseName = App.Path & "\sapp.mdb"
    Data1.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 4 & " order by cl_numero DESC"
    Data1.Refresh
    If Data1.Recordset.RecordCount > 0 Then
'       Data1.Recordset.MoveLast
       If IsNull(Data1.Recordset("cl_numero")) = True Then
          txt_nro.Text = 1
       Else
          txt_nro.Text = Data1.Recordset("cl_numero") + 1
       End If
    Else
       txt_nro.Text = 1
    End If
    
    labnro2.Caption = data_par.Recordset("limite_dia") + 1
    mfecha.Text = Format(Date, "dd/mm/yyyy")
    txt_hora.Text = Format(Time, "HH:mm")
    Combo1.SetFocus
    Combo2.Enabled = False
    mfecfin.Enabled = False
Else
    MsgBox "Usuario no autorizado para crear registros"
End If

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
Combo1.Enabled = False
List1.Enabled = False
txt_encab.Enabled = False
txt_detal.Enabled = True
Combo2.Enabled = False
mfecfin.Enabled = False
Check1.Enabled = False
b_ver.Enabled = False
b_nover.Enabled = True
'Check3.Enabled = True


End Sub

Private Sub Check1_Click()
If Check1.value = 1 Then
   If Check2.value = 1 Then
      Check2.value = 0
   End If
End If

End Sub

Private Sub Check2_Click()
If Check2.value = 1 Then
   If Check1.value = 1 Then
      Check1.value = 0
   End If
End If

End Sub

Private Sub Combo1_Click()
b_agreg_Click

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

Private Sub Command1_Click()
Dim XXdest As Long
Frame1.Enabled = True
List1.AddItem "UNO"
List1.AddItem "DOS"
List1.AddItem "TRES"

List1.ListIndex = 0
For XXdest = 0 To List1.ListCount - 1
   MsgBox "es: " & List1.List(List1.ListIndex)
    
    If List1.ListCount - 1 = List1.ListIndex Then
   Else
       List1.ListIndex = List1.ListIndex + 1
   End If
   
Next

End Sub

Private Sub DBGrid1_DblClick()
borracamp
igualaacc

End Sub

Private Sub Form_Load()
Combo1.Clear
Combo1.AddItem "DIRECTOR GENERAL"
Combo1.AddItem "GERENTE GENERAL"
Combo1.AddItem "DIRECCION TECNICA"
Combo1.AddItem "SUB-DIREC.TECNICA"
Combo1.AddItem "GERENTE COMERCIAL"
Combo1.AddItem "JEFE DE MEDICOS DE MOVIL"
Combo1.AddItem "JEFE CHOFERES Y MANT."
Combo1.AddItem "JEFE TESORERIA/CONT."
Combo1.AddItem "JEFE C.COMPUTOS"
Combo1.AddItem "JEFE BASES Y ENF."
Combo1.AddItem "JEFE FARMACIA/ECONOMATO"
Combo1.AddItem "JEFE DESPACHO"
Combo1.AddItem "ENCARGADO METAS"
Combo1.AddItem "ASESOR GESTION CAP.HUMANO"

Combo1.ListIndex = -1
List1.Clear
data_accion.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_carga2.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_graba2.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_accion.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 4 & " and (cl_nom_sup ='" & "TODOS" & "' or cl_nom_sup ='" & WElusuario & "') and estado <>" & 97 & " order by cl_fnac DESC"
data_accion.Refresh

data_graba.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_graba.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 4 & " order by estado"
data_graba.Refresh
data_cargo.Connect = "odbc;dsn=" & Xconexrmt & ";"
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
Combo1.ListIndex = -1
List1.Clear
txt_encab.Text = ""
txt_detal.Text = ""
mfecfin.Enabled = True
mfecfin.Text = "__/__/____"
mfecfin.Enabled = False
Combo2.Enabled = True
Combo2.ListIndex = -1
Combo2.Enabled = False
Check3.value = 0
labnro2.Caption = ""

End Function

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub txt_detal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Combo2.Enabled = True Then
      Combo2.SetFocus
   End If
End If

End Sub

Private Sub txt_encab_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_detal.SetFocus
End If

End Sub

Public Function igualaacc()
If data_accion.Recordset.RecordCount > 0 Then
    If IsNull(data_accion.Recordset("estado")) = False Then
       labnro2.Caption = data_accion.Recordset("estado")
    Else
       labnro2.Caption = 0
    End If
    If IsNull(data_accion.Recordset("cl_numero")) = False Then
       txt_nro.Text = data_accion.Recordset("cl_numero")
    Else
       txt_nro.Text = 0
    End If
    If IsNull(data_accion.Recordset("cl_zona")) = False Then
       Check3.value = data_accion.Recordset("cl_zona")
    Else
       Check3.value = 0
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
    data_carga2.RecordSource = "Select * from infor_sol where cl_numero =" & data_accion.Recordset("cl_numero") & " and cl_nomcobr =" & 4 & " and estado =" & 97 & " order by cl_fnac DESC"
    data_carga2.Refresh
    If data_carga2.Recordset.RecordCount > 0 Then
       data_carga2.Recordset.MoveFirst
       Do While Not data_carga2.Recordset.EOF
          If IsNull(data_carga2.Recordset("cl_desc2")) = False Then
             List1.AddItem data_carga2.Recordset("cl_desc2")
          Else
             List1.AddItem "NO AGREGADO"
          End If
          data_carga2.Recordset.MoveNext
       Loop
       data_carga2.RecordSource = "Select * from infor_sol where cl_numero =" & data_accion.Recordset("cl_numero") & " and cl_nomcobr =" & 4 & " and estado <>" & 97 & " order by cl_fnac DESC"
       data_carga2.Refresh
       If data_carga2.Recordset.RecordCount > 0 Then
          data_carga2.Recordset.MoveFirst
          If IsNull(data_carga2.Recordset("cl_desc2")) = False Then
             List1.AddItem data_carga2.Recordset("cl_desc2")
          Else
             List1.AddItem "NO AGREGADO"
          End If
       End If
    Else
       data_carga2.RecordSource = "Select * from infor_sol where cl_numero =" & data_accion.Recordset("cl_numero") & " and cl_nomcobr =" & 4 & " order by cl_fnac DESC"
       data_carga2.Refresh
       If data_carga2.Recordset.RecordCount > 0 Then
          data_carga2.Recordset.MoveFirst
          Do While Not data_carga2.Recordset.EOF
             If IsNull(data_carga2.Recordset("cl_desc2")) = False Then
                List1.AddItem data_carga2.Recordset("cl_desc2")
             Else
                List1.AddItem "NO AGREGADO"
             End If
             data_carga2.Recordset.MoveNext
          Loop
       Else
          List1.AddItem data_accion.Recordset("cl_desc2")
       End If
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
    If IsNull(data_accion.Recordset("cl_zona")) = False Then
       Check3.value = data_accion.Recordset("cl_zona")
    Else
       Check3.value = 0
    End If
End If

End Function
