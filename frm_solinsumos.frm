VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frm_solinsumos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud de insumos informáticos"
   ClientHeight    =   6105
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9135
   Icon            =   "frm_solinsumos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6105
   ScaleWidth      =   9135
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_consped 
      Caption         =   "data_consped"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data data_stockbaja 
      Caption         =   "data_stockbaja"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_pedido 
      Caption         =   "data_pedido"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_stock 
      Caption         =   "data_stock"
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
      Top             =   3600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Control"
      Enabled         =   0   'False
      Height          =   735
      Left            =   240
      TabIndex        =   16
      Top             =   2520
      Width           =   8655
      Begin MSMask.MaskEdBox mfenv 
         Height          =   375
         Left            =   4320
         TabIndex        =   19
         Top             =   240
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Enviado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   240
         TabIndex        =   17
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lab_usuaenv 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   6360
         TabIndex        =   20
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha enviado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2760
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton b_cance 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   3120
      Picture         =   "frm_solinsumos.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Cancelar datos"
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2160
      Picture         =   "frm_solinsumos.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Grabar datos"
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton b_edita 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1200
      Picture         =   "frm_solinsumos.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Editar el registro seleccionado"
      Top             =   3240
      Width           =   615
   End
   Begin VB.CommandButton b_alta 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_solinsumos.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Crear nuevo registro"
      Top             =   3240
      Width           =   615
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_solinsumos.frx":1BB2
      Height          =   2175
      Left            =   240
      OleObjectBlob   =   "frm_solinsumos.frx":1BCC
      TabIndex        =   11
      Top             =   3840
      Width           =   8655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Datos para la solicitud"
      Enabled         =   0   'False
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8655
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "frm_solinsumos.frx":2DEB
         Height          =   660
         Left            =   2640
         TabIndex        =   8
         Top             =   960
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   1164
         _Version        =   393216
         Style           =   1
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_solinsumos.frx":2E04
         Left            =   5280
         List            =   "frm_solinsumos.frx":2E11
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox t_base 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin MSMask.MaskEdBox mfec 
         Height          =   375
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox t_cant 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Conformidad:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   3240
         TabIndex        =   22
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label labcod 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CANTIDAD:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "INSUMO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   2415
      End
      Begin VB.Label labusua 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   6240
         TabIndex        =   6
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "USUARIO:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   4800
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "BASE:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2880
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "FECHA:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   5760
      Picture         =   "frm_solinsumos.frx":2E38
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1575
   End
End
Attribute VB_Name = "frm_solinsumos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_alta_Click()

XAlta = 1
Frame1.Enabled = True
mfec.Text = Date
t_base.Text = frm_menu.data_parse.Recordset("base")
labusua.Caption = WElusuario
mfec.Enabled = False
t_base.Enabled = False
DBCombo1.SetFocus
DBGrid1.Enabled = False
b_alta.Enabled = False
b_graba.Enabled = True
b_edita.Enabled = False
b_cance.Enabled = True


End Sub

Private Sub b_cance_Click()
b_alta.Enabled = True
b_graba.Enabled = False
b_edita.Enabled = True
b_cance.Enabled = False
Frame2.Enabled = False
Frame1.Enabled = False

mfec.Text = "__/__/____"
t_base.Text = ""
labusua.Caption = ""
DBCombo1.Text = ""
labcod.Caption = ""
t_cant.Text = ""
Combo1.ListIndex = -1
DBGrid1.Enabled = True

End Sub

Private Sub b_edita_Click()
If mfec.Text <> "__/__/____" Then
    If data_pedido.Recordset("usuario") = WElusuario Then
       If IsNull(data_pedido.Recordset("envio")) = False Then
          If data_pedido.Recordset("envio") = 1 Then
             MsgBox "El pedido ya fue enviado, solo puede registrar conformidad."
             Frame1.Enabled = True
             mfec.Enabled = False
             t_base.Enabled = False
             t_cant.Enabled = False
             DBCombo1.Enabled = False
             Combo1.SetFocus
             DBGrid1.Enabled = False
             b_alta.Enabled = False
             b_graba.Enabled = True
             b_edita.Enabled = False
             b_cance.Enabled = True
          Else
            If WElusuario = "JFERNAN" Or WElusuario = "COMPUTOS" Or WElusuario = "JDIAZ" Or WElusuario = "GVELAZQUEZ" Then
               Frame2.Enabled = True
               Frame1.Enabled = True
               mfec.Enabled = False
               t_base.Enabled = False
               t_cant.SetFocus
            Else
               Frame2.Enabled = False
               Frame1.Enabled = True
               mfec.Enabled = False
               t_base.Enabled = False
               t_cant.SetFocus
            End If
            XAlta = 2
            b_alta.Enabled = False
            b_graba.Enabled = True
            b_edita.Enabled = False
            b_cance.Enabled = True
            DBGrid1.Enabled = False
          End If
       Else
            If WElusuario = "JFERNAN" Or WElusuario = "COMPUTOS" Or WElusuario = "JDIAZ" Or WElusuario = "GVELAZQUEZ" Then
               Frame2.Enabled = True
               Frame1.Enabled = True
               mfec.Enabled = False
               t_base.Enabled = False
               t_cant.SetFocus
            Else
               Frame2.Enabled = False
               Frame1.Enabled = True
               mfec.Enabled = False
               t_base.Enabled = False
               t_cant.SetFocus
            End If
            XAlta = 2
            b_alta.Enabled = False
            b_graba.Enabled = True
            b_edita.Enabled = False
            b_cance.Enabled = True
            DBGrid1.Enabled = False
       End If
    Else
        If WElusuario = "JFERNAN" Or WElusuario = "COMPUTOS" Or WElusuario = "JDIAZ" Or WElusuario = "GVELAZQUEZ" Then
            If WElusuario = "JFERNAN" Or WElusuario = "COMPUTOS" Or WElusuario = "JDIAZ" Or WElusuario = "GVELAZQUEZ" Then
               Frame2.Enabled = True
               Frame1.Enabled = True
               mfec.Enabled = False
               t_base.Enabled = False
               t_cant.SetFocus
            Else
               Frame2.Enabled = False
               Frame1.Enabled = True
               mfec.Enabled = False
               t_base.Enabled = False
               t_cant.SetFocus
            End If
            XAlta = 2
            b_alta.Enabled = False
            b_graba.Enabled = True
            b_edita.Enabled = False
            b_cance.Enabled = True
            DBGrid1.Enabled = False
        Else
            MsgBox "No es el usuario creador del pedido, no puede modificar"
        End If
    End If
Else
    MsgBox "No ha seleccionado registro"
End If
End Sub

Private Sub b_graba_Click()
'3197
On Error GoTo Algraba

If XAlta = 1 Then
   If t_cant.Text <> "" Then
      If DBCombo1.Text <> "" Then
         data_consped.RecordSource = "select * from Pedido_infor where usuario ='" & WElusuario & "' and envio in (1) and confor is null"
         data_consped.Refresh
         If data_consped.Recordset.RecordCount > 0 Then
            MsgBox "Tiene pedido sin registrar conformidad, verifique!"
         Else
             Data1.RecordSource = "select * from stock where descrip ='" & DBCombo1.Text & "' and grupo =" & 3
             Data1.Refresh
             If Data1.Recordset.RecordCount > 0 Then
                 data_pedido.Recordset.AddNew
                 data_pedido.Recordset("fecha") = mfec.Text
                 data_pedido.Recordset("hora") = Format(Time, "HH:mm")
                 data_pedido.Recordset("base") = t_base.Text
                 data_pedido.Recordset("usuario") = labusua.Caption
                 data_pedido.Recordset("ped_desc") = Mid(DBCombo1.Text, 1, 80)
                 data_pedido.Recordset("ped_cod") = Val(labcod.Caption)
                 data_pedido.Recordset("ped_cant") = t_cant.Text
                 If Combo1.ListIndex >= 0 Then
                    If Check1.Value = 1 Then
                       data_pedido.Recordset("confor") = Combo1.Text
                    Else
                       MsgBox "La conformidad no se graba porque no ha sido enviado"
                    End If
                 End If
                 data_pedido.Recordset.Update
                 data_pedido.Refresh
                 b_alta.Enabled = True
                 b_graba.Enabled = False
                 b_edita.Enabled = True
                 b_cance.Enabled = False
                 DBGrid1.Enabled = True
                mfec.Text = "__/__/____"
                t_base.Text = ""
                labusua.Caption = ""
                DBCombo1.Text = ""
                labcod.Caption = ""
                t_cant.Text = ""
                Combo1.ListIndex = -1
            Else
                MsgBox "No se encuentra el ítem, verifique", vbInformation
            End If
         End If
      Else
         MsgBox "Debe ingresar insumo"
      End If
   Else
      MsgBox "Debe ingresar cantidad"
   End If
   XAlta = 0
Else
   If t_cant.Text <> "" Then
      If DBCombo1.Text <> "" Then
         Data1.RecordSource = "select * from stock where descrip ='" & DBCombo1.Text & "' and grupo =" & 3
         Data1.Refresh
         If Data1.Recordset.RecordCount > 0 Then
            data_pedido.Recordset.Edit
            data_pedido.Recordset("ped_cant") = t_cant.Text
            If Combo1.ListIndex >= 0 Then
               If Check1.Value = 1 Then
                  data_pedido.Recordset("confor") = Combo1.Text
               Else
                  MsgBox "La conformidad no se graba porque no ha sido enviado"
               End If
            End If
            If Check1.Value = 1 Then
               If IsNull(data_pedido.Recordset("envio")) = False Then
                  MsgBox "El pedido ya ha sido enviado, no se bajará nuevamente del stock", vbInformation
               Else
                  data_pedido.Recordset("envio") = Check1.Value
                  data_pedido.Recordset("fecha_env") = mfenv.Text
                  data_pedido.Recordset("usuario_env") = WElusuario
                  data_stock.RecordSource = "Select * from stock where id =" & labcod.Caption
                  data_stock.Refresh
                  If data_stock.Recordset.RecordCount > 0 Then
                     data_stock.Recordset.Edit
                     data_stock.Recordset("actual") = data_stock.Recordset("actual") - t_cant.Text
                     data_stock.Recordset.Update
                  End If
               End If
            End If
            data_pedido.Recordset.Update
            data_pedido.Refresh
            b_alta.Enabled = True
            b_graba.Enabled = False
            b_edita.Enabled = True
            b_cance.Enabled = False
            DBGrid1.Enabled = True
            mfec.Text = "__/__/____"
            t_base.Text = ""
            labusua.Caption = ""
            DBCombo1.Text = ""
            labcod.Caption = ""
            t_cant.Text = ""
            Check1.Value = 0
            mfenv.Text = "__/__/____"
            lab_usuaenv.Caption = ""
            Combo1.ListIndex = -1
            XAlta = 0
         Else
            MsgBox "No se encuentra el ítem, verifique", vbInformation
         End If
      Else
         MsgBox "Debe ingresar insumo"
      End If
   Else
      MsgBox "Debe ingresar cantidad"
   End If

End If
   
Exit Sub

Algraba:
        If Err.Number = 3197 Then
           MsgBox "No se modificaron datos"
        Else
           MsgBox "Error: " & Err.Description
        End If
        
End Sub

Private Sub DBCombo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If DBCombo1.Text <> "" Then
      DBCombo1.ListField = "descrip"
      DBCombo1.BoundColumn = "descrip"
      If IsNumeric(DBCombo1.Text) Then
         If DBCombo1.Text <> "" Then
            data_stock.Recordset.FindFirst "id =" & DBCombo1.Text
         End If
         If Not data_stock.Recordset.NoMatch Then
            labcod.Caption = data_stock.Recordset("id")
            DBCombo1.Text = data_stock.Recordset("descrip")
            DBCombo1.Height = 500
            DBCombo1.ListField = ""
            DBCombo1.BoundColumn = ""
            t_cant.SetFocus
         Else
            data_stock.RecordSource = "select * from stock where id >=" & DBCombo1.Text & " and grupo =" & 3
            data_stock.Refresh
            DBCombo1.Height = 1400
         End If
      Else
         data_stock.Recordset.FindFirst "descrip ='" & DBCombo1.Text & "'"
         If Not data_stock.Recordset.NoMatch Then
            DBCombo1.Text = data_stock.Recordset("descrip")
            labcod.Caption = data_stock.Recordset("id")
            DBCombo1.Height = 500
            DBCombo1.ListField = ""
            DBCombo1.BoundColumn = ""
            t_cant.SetFocus
         Else
            data_stock.RecordSource = "select * from stock where descrip >='" & DBCombo1.Text & "' and grupo =" & 3 & " order by descrip"
            data_stock.Refresh
            DBCombo1.Height = 1400
         End If
      End If
   End If
End If

End Sub

Private Sub DBGrid1_DblClick()
mfec.Text = Format(data_pedido.Recordset("fecha"), "dd/mm/yyyy")
t_base.Text = data_pedido.Recordset("base")
labusua.Caption = data_pedido.Recordset("usuario")
DBCombo1.Text = data_pedido.Recordset("ped_desc")
labcod.Caption = data_pedido.Recordset("ped_cod")
t_cant.Text = data_pedido.Recordset("ped_cant")
If IsNull(data_pedido.Recordset("confor")) = False Then
   If data_pedido.Recordset("confor") = "Conforme" Then
      Combo1.ListIndex = 0
   Else
      If data_pedido.Recordset("confor") = "No Conforme" Then
         Combo1.ListIndex = 1
      Else
         If data_pedido.Recordset("confor") = "Con Demora" Then
            Combo1.ListIndex = 2
         Else
            Combo1.ListIndex = -1
         End If
      End If
   End If
Else
   Combo1.ListIndex = -1
End If
   
If IsNull(data_pedido.Recordset("envio")) = False Then
   Check1.Value = data_pedido.Recordset("envio")
   mfenv.Text = data_pedido.Recordset("fecha_env")
   lab_usuaenv.Caption = data_pedido.Recordset("usuario_env")
Else
   Check1.Value = 0
   mfenv.Text = "__/__/____"
   lab_usuaenv.Caption = ""
End If

End Sub

Private Sub Form_Load()
data_stock.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_stock.RecordSource = "select * from stock where grupo =" & 3
data_stock.Refresh

Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_consped.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_stockbaja.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_pedido.Connect = "odbc;dsn=" & Xconexrmt & ";"
If WElusuario = "JFERNAN" Or WElusuario = "COMPUTOS" Or WElusuario = "JDIAZ" Or WElusuario = "GVELAZQUEZ" Then
   data_pedido.RecordSource = "select * from Pedido_infor order by fecha DESC"
Else
   data_pedido.RecordSource = "select * from Pedido_infor where usuario ='" & WElusuario & "' order by fecha DESC"
End If

data_pedido.Refresh



End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Width = Me.Width
    .Height = Me.Height
End With

End Sub

Private Sub mfenv_GotFocus()
mfenv.Text = Format(Date, "dd/mm/yyyy")
lab_usuaenv.Caption = WElusuario

End Sub
