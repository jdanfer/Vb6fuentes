VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_planacc2 
   BackColor       =   &H00400040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planes de acción"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9735
   Icon            =   "frm_planacc2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   9735
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton b_nover 
      Enabled         =   0   'False
      Height          =   495
      Left            =   5280
      Picture         =   "frm_planacc2.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Cancelar la visualización del cuadro descripción."
      Top             =   7080
      Width           =   615
   End
   Begin VB.CommandButton b_ver 
      Height          =   495
      Left            =   4080
      Picture         =   "frm_planacc2.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Editar el cuadro DESCRIPCION para leer los datos ingresados."
      Top             =   7080
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
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6720
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
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ver solo acciones en proceso."
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
      Left            =   4920
      TabIndex        =   32
      Top             =   5160
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
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   7080
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
      Left            =   6360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton b_buscafec 
      BackColor       =   &H00FF8080&
      Height          =   615
      Left            =   8760
      Picture         =   "frm_planacc2.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   4680
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
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6600
      Visible         =   0   'False
      Width           =   2895
   End
   Begin MSMask.MaskEdBox mfecbusca 
      Height          =   375
      Left            =   6600
      TabIndex        =   28
      Top             =   4680
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
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_planacc2.frx":1108
      Height          =   1695
      Left            =   120
      OleObjectBlob   =   "frm_planacc2.frx":1122
      TabIndex        =   26
      Top             =   5400
      Width           =   9495
   End
   Begin VB.CommandButton b_infor 
      BackColor       =   &H00FF8080&
      Height          =   495
      Left            =   3960
      Picture         =   "frm_planacc2.frx":1E4D
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Informes"
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton b_cancela 
      BackColor       =   &H00FF8080&
      Enabled         =   0   'False
      Height          =   495
      Left            =   3000
      Picture         =   "frm_planacc2.frx":228F
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Cancelar movimiento realizado"
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H00FF8080&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2040
      Picture         =   "frm_planacc2.frx":26D1
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Grabar datos"
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton b_modif 
      BackColor       =   &H00FF8080&
      Height          =   495
      Left            =   1080
      Picture         =   "frm_planacc2.frx":2B13
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Modificar datos de registro seleccionado"
      Top             =   4560
      Width           =   615
   End
   Begin VB.CommandButton b_nuevo 
      BackColor       =   &H00FF8080&
      Height          =   495
      Left            =   120
      Picture         =   "frm_planacc2.frx":2F55
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Ingresar nuevo registro"
      Top             =   4560
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Datos de la acción solicitada"
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
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   9495
      Begin MSMask.MaskEdBox mfecfin 
         Height          =   375
         Left            =   5520
         TabIndex        =   20
         Top             =   3720
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
         ItemData        =   "frm_planacc2.frx":3397
         Left            =   2040
         List            =   "frm_planacc2.frx":33A4
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   3720
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
         Height          =   1215
         Left            =   2040
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   2400
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
         Picture         =   "frm_planacc2.frx":33CB
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
         Picture         =   "frm_planacc2.frx":380D
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
         ItemData        =   "frm_planacc2.frx":3C4F
         Left            =   2040
         List            =   "frm_planacc2.frx":3C51
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
      Begin VB.Label labid 
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Destinatario/s"
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
         Left            =   6120
         TabIndex        =   33
         Top             =   720
         Width           =   3015
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Conformidad:"
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
         TabIndex        =   18
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
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
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   2400
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
         TabIndex        =   14
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Dirigido a:"
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
      TabIndex        =   29
      Top             =   7080
      Width           =   3495
   End
   Begin VB.Label Label13 
      BackColor       =   &H0080FFFF&
      Caption         =   "Buscar fecha:"
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
      Left            =   4920
      TabIndex        =   27
      Top             =   4680
      Width           =   1575
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
      Left            =   2880
      TabIndex        =   8
      Top             =   0
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
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "frm_planacc2"
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
If mfecbusca.Text = "__/__/____" Then
Else
    If WElusuario = "BDD" Or WElusuario = "BRUNO" Or WElusuario = "SPEREZ" Then
       data_accion.RecordSource = "Select * from infor_sol where estado >=" & 70000 & " and cl_fnac =#" & Format(mfecbusca.Text, "yyyy/mm/dd") & "# order by estado"
       data_accion.Refresh
    Else
       data_accion.RecordSource = "Select * from infor_sol where estado >=" & 70000 & " and cl_fnac =#" & Format(mfecbusca.Text, "yyyy/mm/dd") & "# and (cl_descpag ='" & WElusuario & "' or cl_nom_sup ='" & WElusuario & "') order by estado"
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
Xelnro = txt_nro.Text
XXdest = 0
If XAlta = 1 Then
   If List1.ListCount >= 1 Then
      If Len(txt_encab.Text) > 5 Then
         If Len(txt_detal.Text) > 5 Then
            List1.ListIndex = 0
            For XXdest = 1 To List1.ListCount
                data_graba.Recordset.AddNew
                data_graba.Recordset("cl_etiquet") = 0
                data_graba.Recordset("cl_val2") = 7
                data_graba.Recordset("cl_codigo") = labid.Caption
                data_graba.Recordset("estado") = Xelnro
                data_graba.Recordset("cl_fnac") = mfecha.Text
                data_graba.Recordset("cl_ruc") = txt_hora.Text
                data_cargo.Recordset.FindFirst "chofer ='" & List1.List(List1.ListIndex) & "'"
                If Not data_cargo.Recordset.NoMatch Then
                   data_graba.Recordset("cl_nom_sup") = Mid(data_cargo.Recordset("medico"), 1, 25)
                Else
                   data_graba.Recordset("cl_nom_sup") = WElusuario
                End If
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
                If Option1.Value = True Then
                   data_graba.Recordset("cl_atrasop") = 1
                Else
                   If Option2.Value = True Then
                      data_graba.Recordset("cl_atrasop") = 2
                   Else
                      If Option3.Value = True Then
                         data_graba.Recordset("cl_atrasop") = 3
                      Else
                         data_graba.Recordset("cl_Atrasop") = 0
                      End If
                   End If
                End If
                If Combo3.ListIndex >= 0 Then
                   data_graba.Recordset("cl_grupo") = Combo3.ListIndex
                Else
                   data_graba.Recordset("cl_grupo") = -1
                End If
                data_graba.Recordset("cl_codconv") = "A"
                data_graba.Recordset.Update
                Xelnro = Xelnro + 1
                If labid.Caption <> "" Then
                   labid.Caption = labid.Caption + 1
                End If
                If List1.ListCount - 1 = List1.ListIndex Then
                Else
                   List1.ListIndex = List1.ListIndex + 1
                End If
            Next
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
   data_graba.Recordset.Edit
   List1.ListIndex = 0
   
   data_cargo.Recordset.FindFirst "chofer ='" & List1.List(List1.ListIndex) & "'"
   If Not data_cargo.Recordset.NoMatch Then
      data_graba.Recordset("cl_nom_sup") = Mid(data_cargo.Recordset("medico"), 1, 25)
   Else
      data_graba.Recordset("cl_nom_sup") = WElusuario
   End If
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
'      data_graba.Recordset("cl_fecing") = Date
   End If
   If Option1.Value = True Then
      data_graba.Recordset("cl_atrasop") = 1
   Else
      If Option2.Value = True Then
         data_graba.Recordset("cl_atrasop") = 2
      Else
         If Option3.Value = True Then
            data_graba.Recordset("cl_atrasop") = 3
         Else
            data_graba.Recordset("cl_Atrasop") = 0
         End If
      End If
   End If
   If Combo3.ListIndex >= 0 Then
      data_graba.Recordset("cl_grupo") = Combo3.ListIndex
   Else
      data_graba.Recordset("cl_grupo") = -1
   End If
   
   data_graba.Recordset.Update
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
frm_planaccing.Show vbModal

End Sub

Private Sub b_infor_Click()
If WElusuario = "BDD" Or WElusuario = "SPEREZ" Or WElusuario = "JFERNAN" Or WElusuario = "CALONSO" Or WElusuario = "SDOMINGUEZ" Then
   frm_infmejoras.Show vbModal
Else
   MsgBox "Usuario no habilitado para informes"
End If

End Sub

Private Sub b_modif_Click()
If WElusuario = "SPEREZ" Or WElusuario = "COMPUTOS" Then
   'If labusuario.Caption = data_accion.Recordset("cl_descpag") Then
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
         data_graba.RecordSource = "Select * from infor_sol where estado =" & data_accion.Recordset("estado")
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
         If WElusuario = "SPEREZ" Or WElusuario = "JFERNAN" Then
            Combo2.Enabled = True
            mfecfin.Enabled = True
         Else
            Combo2.Enabled = False
            mfecfin.Enabled = False
         End If
        'Else
        '    MsgBox "NO ES PROPIETARIO DE LA ACCION", vbCritical
        '    DBGrid1.SetFocus
        'End If
    End If
Else
   MsgBox "Usuario no autorizado para modificación, solo se habilita el cuadro Descripción para ver"
   Frame1.Enabled = True

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
Option1.Enabled = True
Option2.Enabled = True
Option3.Enabled = True
Combo3.Enabled = True
Combo2.Enabled = True
mfecfin.Enabled = True
Check1.Enabled = True
b_ver.Enabled = True
b_nover.Enabled = False
Frame1.Enabled = False

End Sub

Private Sub b_nuevo_Click()
If WElusuario = "SPEREZ" Or WElusuario = "COMPUTOS" Then
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
    If data_graba.Recordset.RecordCount > 0 Then
       data_graba.Recordset.MoveLast
       If data_graba.Recordset("estado") >= 70000 Then
          txt_nro.Text = data_graba.Recordset("estado") + 1
       Else
          txt_nro.Text = 70000
       End If
    Else
       txt_nro.Text = 70000
    End If
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
Option1.Enabled = False
Option2.Enabled = False
Option3.Enabled = False
Combo3.Enabled = False
Combo2.Enabled = False
mfecfin.Enabled = False
Check1.Enabled = False
b_ver.Enabled = False
b_nover.Enabled = True


End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
   If WElusuario = "BDD" Or WElusuario = "BRUNO" Or WElusuario = "SPEREZ" Then
      data_accion.RecordSource = "Select * from infor_sol where estado >=" & 70000 & " and cl_codconv ='" & "A" & "' order by cl_fnac DESC"
      data_accion.Refresh
   Else
      data_accion.RecordSource = "Select * from infor_sol where estado >=" & 70000 & " and cl_codconv ='" & "A" & "' and (cl_descpag ='" & WElusuario & "' or cl_nom_sup ='" & WElusuario & "') order by cl_fnac DESC"
      data_accion.Refresh
   End If
Else
   If WElusuario = "BDD" Or WElusuario = "BRUNO" Or WElusuario = "SPEREZ" Then
      data_accion.RecordSource = "Select * from infor_sol where estado >=" & 70000 & " order by cl_fnac DESC"
      data_accion.Refresh
   Else
      data_accion.RecordSource = "Select * from infor_sol where estado >=" & 70000 & " and (cl_descpag ='" & WElusuario & "' or cl_nom_sup ='" & WElusuario & "') order by cl_fnac DESC"
      data_accion.Refresh
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
Combo1.AddItem "JEFE PRUEBAS"
Combo1.ListIndex = -1
List1.Clear
data_accion.DatabaseName = App.Path & "\sapp.mdb"
If WElusuario = "BDD" Or WElusuario = "BRUNO" Or WElusuario = "SPEREZ" Then
   data_accion.RecordSource = "Select * from infor_sol where estado >=" & 70000 & " order by cl_fnac DESC"
   data_accion.Refresh
Else
   data_accion.RecordSource = "Select * from infor_sol where estado >=" & 70000 & " and (cl_descpag ='" & WElusuario & "' or cl_nom_sup ='" & WElusuario & "') order by cl_fnac DESC"
   data_accion.Refresh
End If

data_graba.DatabaseName = App.Path & "\sapp.mdb"
data_graba.RecordSource = "Select * from infor_sol order by estado"
data_graba.Refresh
data_cargo.DatabaseName = App.Path & "\sapp.mdb"
data_cargo.RecordSource = "movil"
data_cargo.Refresh

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
Option1.Value = False
Option2.Value = False
Option3.Value = False
Combo3.ListIndex = -1

End Function

Private Sub txt_detal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo3.SetFocus
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
       List1.AddItem data_accion.Recordset("cl_desc2")
    Else
       List1.Clear
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
    If IsNull(data_accion.Recordset("cl_atrasop")) = False Then
       If data_accion.Recordset("cl_atrasop") = 1 Then
          Option1.Value = True
       Else
          If data_accion.Recordset("cl_atrasop") = 2 Then
             Option2.Value = True
          Else
             If data_accion.Recordset("cl_atrasop") = 3 Then
                Option3.Value = True
             Else
                Option1.Value = False
                Option2.Value = False
                Option3.Value = False
             End If
          End If
       End If
    Else
       Option1.Value = False
       Option2.Value = False
       Option3.Value = False
    End If
    If IsNull(data_accion.Recordset("cl_grupo")) = False Then
       If data_accion.Recordset("cl_grupo") >= 0 Then
          Combo3.ListIndex = data_accion.Recordset("cl_grupo")
       Else
          Combo3.ListIndex = -1
       End If
    Else
       Combo3.ListIndex = -1
    End If
End If

End Function
