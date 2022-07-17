VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_solaudito 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud de auditoría"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   12090
   Icon            =   "frm_solaudito.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   12090
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_graba 
      Caption         =   "data_graba"
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
      Top             =   4920
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton b_buscafec 
      BackColor       =   &H0080FF80&
      Height          =   615
      Left            =   9480
      Picture         =   "frm_solaudito.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   4800
      Width           =   735
   End
   Begin VB.Data data_accion 
      Caption         =   "data_accion"
      Connect         =   "Access"
      DatabaseName    =   "C:\sappmys\sappmysql\sapp.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "env_soc"
      Top             =   4800
      Visible         =   0   'False
      Width           =   2895
   End
   Begin MSMask.MaskEdBox mfecbusca 
      Height          =   375
      Left            =   10320
      TabIndex        =   22
      Top             =   4560
      Width           =   1695
      _ExtentX        =   2990
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
      Bindings        =   "frm_solaudito.frx":0884
      Height          =   2055
      Left            =   120
      OleObjectBlob   =   "frm_solaudito.frx":089E
      TabIndex        =   20
      Top             =   5400
      Width           =   11895
   End
   Begin VB.CommandButton b_infor 
      BackColor       =   &H0080FF80&
      Height          =   735
      Left            =   4920
      Picture         =   "frm_solaudito.frx":15C9
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Informes"
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton b_cancela 
      BackColor       =   &H0080FF80&
      Enabled         =   0   'False
      Height          =   735
      Left            =   3720
      Picture         =   "frm_solaudito.frx":1A0B
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Cancelar movimiento realizado"
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H0080FF80&
      Enabled         =   0   'False
      Height          =   735
      Left            =   2520
      Picture         =   "frm_solaudito.frx":1E4D
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Grabar datos"
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton b_modif 
      BackColor       =   &H0080FF80&
      Height          =   735
      Left            =   1320
      Picture         =   "frm_solaudito.frx":228F
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Modificar datos de registro seleccionado"
      Top             =   4440
      Width           =   855
   End
   Begin VB.CommandButton b_nuevo 
      BackColor       =   &H0080FF80&
      Height          =   735
      Left            =   120
      Picture         =   "frm_solaudito.frx":26D1
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Ingresar nuevo registro"
      Top             =   4440
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Datos para la solicitud de auditoría"
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
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   11895
      Begin VB.TextBox txt_obs 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2040
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   29
         Top             =   3360
         Width           =   7695
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
         ItemData        =   "frm_solaudito.frx":2B13
         Left            =   2040
         List            =   "frm_solaudito.frx":2B1D
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txt_usuario 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9960
         TabIndex        =   25
         Top             =   360
         Width           =   1815
      End
      Begin MSMask.MaskEdBox mfecfin 
         Height          =   375
         Left            =   5160
         TabIndex        =   14
         Top             =   2760
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
      Begin VB.ComboBox Combo2 
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
         ItemData        =   "frm_solaudito.frx":2B3A
         Left            =   2040
         List            =   "frm_solaudito.frx":2B44
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2760
         Width           =   2895
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
         MaxLength       =   310
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   1560
         Width           =   7695
      End
      Begin VB.TextBox txt_hora 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   7560
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin MSMask.MaskEdBox mfecha 
         Height          =   375
         Left            =   4920
         TabIndex        =   4
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         ForeColor       =   255
         Enabled         =   0   'False
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
      Begin VB.TextBox txt_nro 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   2040
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Observaciones:"
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
         TabIndex        =   28
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Asistencia:"
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
         Left            =   120
         TabIndex        =   26
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Estado actual:"
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
         Left            =   120
         TabIndex        =   12
         Top             =   2760
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Descripción:"
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
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Usuario:"
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
         Left            =   8640
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HORA:"
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
         Left            =   6480
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "FECHA:"
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
         Left            =   3600
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "NUMERO:"
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
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFFF&
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
      TabIndex        =   23
      Top             =   7440
      Width           =   4095
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Buscar solicitudes por fecha:"
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
      Left            =   7320
      TabIndex        =   21
      Top             =   4560
      Width           =   2895
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
      Height          =   375
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
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "frm_solaudito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub b_buscafec_Click()
If mfecbusca.Text = "__/__/____" Then
Else
    If WElusuario = "BDD" Or WElusuario = "BRUNO" Or WElusuario = "SPEREZ" Or welsuario = "SDOMINGUEZ" Or WElusuario = "CALONSO" Or WElusuario = "COMPUTOS" Then
       data_accion.RecordSource = "Select * from env_soc where cl_codigo >=" & 80000 & " and cl_fnac =#" & Format(mfecbusca.Text, "yyyy/mm/dd") & "# order by cl_codigo"
       data_accion.Refresh
    Else
       data_accion.RecordSource = "Select * from env_soc where cl_codigo >=" & 80000 & " and cl_fnac =#" & Format(mfecbusca.Text, "yyyy/mm/dd") & "# and (cl_descpag ='" & WElusuario & "' or cl_nomvend ='" & WElusuario & "') order by cl_codigo"
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


Private Sub b_graba_Click()
Dim XXdest As Long
Dim Xelnro As Double
Xelnro = txt_nro.Text
XXdest = 0
If XAlta = 1 Then
   If Combo1.ListIndex >= 0 Then
      If txt_detal.Text <> "" Then
         data_graba.Recordset.AddNew
         data_graba.Recordset("cl_etiquet") = 0
         data_graba.Recordset("cl_cantdia") = 8
         data_graba.Recordset("cl_codigo") = Xelnro
         data_graba.Recordset("cl_fnac") = mfecha.Text
         data_graba.Recordset("cl_ruc") = txt_hora.Text
         data_graba.Recordset("cl_nomvend") = txt_usuario.Text
         data_graba.Recordset("cl_descpag") = labusuario.Caption
         If Combo1.ListIndex >= 0 Then
            data_graba.Recordset("cl_schqmn") = Combo1.ListIndex
            data_graba.Recordset("cl_tipclin") = Combo1.Text
         Else
            data_graba.Recordset("cl_schqmn") = 0
            data_graba.Recordset("cl_tipclin") = "AUDITORIA EXTERNA"
         End If
         If Len(txt_detal.Text) > 230 Then
            data_graba.Recordset("info_debit") = Mid(txt_detal.Text, 1, 130)
            data_graba.Recordset("cl_dircobr") = Mid(txt_detal.Text, 131, 100)
            data_graba.Recordset("cl_entre") = Mid(txt_detal.Text, 231, 80)
         Else
            If Len(txt_detal.Text) > 130 Then
               data_graba.Recordset("info_debit") = Mid(txt_detal.Text, 1, 130)
               data_graba.Recordset("cl_dircobr") = Mid(txt_detal.Text, 131, 100)
            Else
               data_graba.Recordset("info_debit") = txt_detal.Text
            End If
         End If
         If Combo2.ListIndex >= 0 Then
            data_graba.Recordset("cl_tipocli") = Combo2.ListIndex
         Else
            data_graba.Recordset("cl_tipocli") = 0
         End If
         If mfecfin.Text = "__/__/____" Then
         Else
            data_graba.Recordset("cl_fecing") = mfecfin.Text
         End If
         If Len(txt_obs.Text) > 0 Then
            If Len(txt_obs.Text) > 170 Then
               data_graba.Recordset("cl_direcci") = Mid(txt_obs.Text, 1, 80)
               data_graba.Recordset("cl_apellid") = Mid(txt_obs.Text, 81, 60)
               data_graba.Recordset("cl_nombre") = Mid(txt_obs.Text, 141, 30)
               data_graba.Recordset("cl_email") = Mid(txt_obs.Text, 171, 30)
            Else
               If Len(txt_obs.Text) > 140 Then
                  data_graba.Recordset("cl_direcci") = Mid(txt_obs.Text, 1, 80)
                  data_graba.Recordset("cl_apellid") = Mid(txt_obs.Text, 81, 60)
                  data_graba.Recordset("cl_nombre") = Mid(txt_obs.Text, 141, 30)
                  data_graba.Recordset("cl_email") = ""
               Else
                  If Len(txt_obs.Text) > 80 Then
                     data_graba.Recordset("cl_direcci") = Mid(txt_obs.Text, 1, 80)
                     data_graba.Recordset("cl_apellid") = Mid(txt_obs.Text, 81, 60)
                     data_graba.Recordset("cl_nombre") = ""
                     data_graba.Recordset("cl_email") = ""
                  Else
                     data_graba.Recordset("cl_direcci") = Mid(txt_obs.Text, 1, 80)
                     data_graba.Recordset("cl_apellid") = ""
                     data_graba.Recordset("cl_nombre") = ""
                     data_graba.Recordset("cl_email") = ""
                  End If
               End If
            End If
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
         XAlta = 0
      Else
         MsgBox "Ingrese texto en detalle"
      End If
   Else
      MsgBox "Ingrese tipo de solicitud"
   End If
Else
   data_graba.Recordset.Edit
   If Combo1.ListIndex >= 0 Then
      data_graba.Recordset("cl_schqmn") = Combo1.ListIndex
      data_graba.Recordset("cl_tipclin") = Combo1.Text
   Else
      data_graba.Recordset("cl_schqmn") = 0
      data_graba.Recordset("cl_tipclin") = "AUDITORIA EXTERNA"
   End If
   If Len(txt_detal.Text) > 230 Then
      data_graba.Recordset("info_debit") = Mid(txt_detal.Text, 1, 130)
      data_graba.Recordset("cl_dircobr") = Mid(txt_detal.Text, 131, 100)
      data_graba.Recordset("cl_entre") = Mid(txt_detal.Text, 231, 80)
   Else
      If Len(txt_detal.Text) > 130 Then
         data_graba.Recordset("info_debit") = Mid(txt_detal.Text, 1, 130)
         data_graba.Recordset("cl_dircobr") = Mid(txt_detal.Text, 131, 100)
      Else
         data_graba.Recordset("info_debit") = txt_detal.Text
      End If
   End If
   If Combo2.ListIndex >= 0 Then
      data_graba.Recordset("cl_tipocli") = Combo2.ListIndex
   Else
      data_graba.Recordset("cl_tipocli") = 0
   End If
   If mfecfin.Text = "__/__/____" Then
   Else
     data_graba.Recordset("cl_fecing") = mfecfin.Text
   End If
    If Len(txt_obs.Text) > 0 Then
       If Len(txt_obs.Text) > 170 Then
          data_graba.Recordset("cl_direcci") = Mid(txt_obs.Text, 1, 80)
          data_graba.Recordset("cl_apellid") = Mid(txt_obs.Text, 81, 60)
          data_graba.Recordset("cl_nombre") = Mid(txt_obs.Text, 141, 30)
          data_graba.Recordset("cl_email") = Mid(txt_obs.Text, 171, 30)
       Else
          If Len(txt_obs.Text) > 140 Then
             data_graba.Recordset("cl_direcci") = Mid(txt_obs.Text, 1, 80)
             data_graba.Recordset("cl_apellid") = Mid(txt_obs.Text, 81, 60)
             data_graba.Recordset("cl_nombre") = Mid(txt_obs.Text, 141, 30)
             data_graba.Recordset("cl_email") = ""
          Else
             If Len(txt_obs.Text) > 80 Then
                data_graba.Recordset("cl_direcci") = Mid(txt_obs.Text, 1, 80)
                data_graba.Recordset("cl_apellid") = Mid(txt_obs.Text, 81, 60)
                data_graba.Recordset("cl_nombre") = ""
                data_graba.Recordset("cl_email") = ""
             Else
                data_graba.Recordset("cl_direcci") = Mid(txt_obs.Text, 1, 80)
                data_graba.Recordset("cl_apellid") = ""
                data_graba.Recordset("cl_nombre") = ""
                data_graba.Recordset("cl_email") = ""
             End If
          End If
       End If
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

End Sub

Private Sub b_infor_Click()
If WElusuario = "BDD" Or WElusuario = "SPEREZ" Or WElusuario = "JFERNAN" Or WElusuario = "CALONSO" Or WElusuario = "SDOMINGUEZ" Then
   frm_infaudito.Show vbModal
Else
   MsgBox "Usuario no habilitado para informes"
End If

End Sub

Private Sub b_modif_Click()
'If labusuario.Caption = data_accion.Recordset("cl_descpag") Then
    XAlta = 0
    b_nuevo.Enabled = False
    b_modif.Enabled = False
    b_graba.Enabled = True
    b_cancela.Enabled = True
    b_buscafec.Enabled = False
    b_infor.Enabled = False
    DBGrid1.Enabled = False
    Frame1.Enabled = True
    borracamp
    data_graba.RecordSource = "Select * from env_soc where cl_codigo =" & data_accion.Recordset("cl_codigo")
    data_graba.Refresh
    If data_graba.Recordset.RecordCount > 0 Then
'       If IsNull(data_graba.Recordset("cl_cua_vto")) = True Then
'          Combo2.Enabled = False
'          mfecfin.Enabled = False
'       Else
'          If data_graba.Recordset("cl_cua_vto") = 1 Then
             Combo2.Enabled = True
             mfecfin.Enabled = True
'          Else
'             Combo2.Enabled = False
'             mfecfin.Enabled = False
'          End If
'       End If
       igualaacc
    Else
       Combo2.Enabled = True
       mfecfin.Enabled = True
       Frame1.Enabled = False
       b_nuevo.Enabled = True
       b_modif.Enabled = True
       b_graba.Enabled = False
       b_cancela.Enabled = False
       b_buscafec.Enabled = True
       b_infor.Enabled = True
       DBGrid1.Enabled = True
    End If
'Else
'    MsgBox "NO ES PROPIETARIO DE LA ACCION", vbCritical
'    DBGrid1.SetFocus
'End If

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
data_graba.Recordset.MoveLast
If data_graba.Recordset("cl_codigo") >= 80000 Then
   txt_nro.Text = data_graba.Recordset("cl_codigo") + 1
Else
   txt_nro.Text = 80000
End If
mfecha.Text = Format(Date, "dd/mm/yyyy")
txt_hora.Text = Format(Time, "HH:mm")
txt_usuario.Text = WElusuario
Combo1.SetFocus
Combo2.Enabled = True
mfecfin.Enabled = True

End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_detal.SetFocus
End If

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_graba.SetFocus
End If

End Sub


Private Sub DBGrid1_DblClick()
borracamp
igualaacc

End Sub

Private Sub Form_Load()
data_accion.Connect = "odbc;dsn=" & Xconexrmt & ";"
If WElusuario = "BDD" Or WElusuario = "BRUNO" Or WElusuario = "SPEREZ" Or WElusuario = "SDOMINGUEZ" Or WElusuario = "CALONSO" Or WElusuario = "COMPUTOS" Then
   data_accion.RecordSource = "Select * from env_soc where cl_codigo >=" & 80000 & " order by cl_fnac"
   data_accion.Refresh
Else
   data_accion.RecordSource = "Select * from env_soc where cl_codigo >=" & 80000 & " and (cl_descpag ='" & WElusuario & "' or cl_nomvend ='" & WElusuario & "') order by cl_fnac"
   data_accion.Refresh
End If

data_graba.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_graba.RecordSource = "Select * from env_soc order by cl_codigo"
data_graba.Refresh
'data_cargo.DatabaseName = App.Path & "\sapp.mdb"
'data_cargo.RecordSource = "movil"
'data_cargo.Refresh
data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"

labusuario.Caption = WElusuario

End Sub


Public Function borracamp()
txt_nro.Text = ""
mfecha.Text = "__/__/____"
txt_hora.Text = ""
Combo1.ListIndex = -1
txt_usuario.Text = ""
txt_detal.Text = ""
txt_obs.Text = ""
'mfecfin.Enabled = True
mfecfin.Text = "__/__/____"
'mfecfin.Enabled = False
'Combo2.Enabled = True
Combo2.ListIndex = -1
'Combo2.Enabled = False

End Function


Private Sub txt_detal_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
'   If Combo2.Enabled = True Then
      Combo2.SetFocus
'   Else
'      b_graba.SetFocus
'   End If
End If

End Sub

Private Sub txt_encab_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_detal.SetFocus
End If

End Sub

Public Function igualaacc()
If data_accion.Recordset.RecordCount > 0 Then
    If IsNull(data_accion.Recordset("cl_codigo")) = False Then
       txt_nro.Text = data_accion.Recordset("cl_codigo")
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
    If IsNull(data_accion.Recordset("cl_nomvend")) = False Then
       txt_usuario.Text = data_accion.Recordset("cl_nomvend")
    Else
       txt_usuario.Text = ""
    End If
    If IsNull(data_accion.Recordset("cl_schqmn")) = False Then
       Combo1.ListIndex = data_accion.Recordset("cl_schqmn")
    Else
       Combo1.ListIndex = 0
    End If
    If IsNull(data_accion.Recordset("info_debit")) = False Then
       txt_detal.Text = data_accion.Recordset("info_debit")
    Else
       txt_detal.Text = ""
    End If
    If IsNull(data_accion.Recordset("cl_dircobr")) = False Then
       txt_detal.Text = txt_detal.Text + " " + data_accion.Recordset("cl_dircobr")
    Else
'       txt_detal.Text = ""
    End If
    If IsNull(data_accion.Recordset("cl_entre")) = False Then
       txt_detal.Text = txt_detal.Text + " " + data_accion.Recordset("cl_entre")
    Else
'       txt_detal.Text = ""
    End If
    
    If IsNull(data_accion.Recordset("cl_tipocli")) = False Then
       Combo2.ListIndex = data_accion.Recordset("cl_tipocli")
    Else
       Combo2.ListIndex = -1
    End If
    If IsNull(data_accion.Recordset("cl_fecing")) = False Then
       mfecfin.Text = Format(data_accion.Recordset("cl_fecing"), "dd/mm/yyyy")
    Else
       mfecfin.Text = "__/__/____"
    End If
    If IsNull(data_accion.Recordset("cl_direcci")) = False Then
       txt_obs.Text = data_accion.Recordset("cl_direcci")
       If IsNull(data_accion.Recordset("cl_apellid")) = False Then
          txt_obs.Text = txt_obs.Text + data_accion.Recordset("cl_apellid")
          If IsNull(data_accion.Recordset("cl_nombre")) = False Then
             txt_obs.Text = txt_obs.Text + data_accion.Recordset("cl_nombre")
             If IsNull(data_accion.Recordset("cl_email")) = False Then
                txt_obs.Text = txt_obs.Text + data_accion.Recordset("cl_email")
             End If
          End If
       End If
    Else
       txt_obs.Text = ""
    End If
End If

End Function


