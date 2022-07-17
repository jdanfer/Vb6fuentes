VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Begin VB.Form frm_abmgrupos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grupos para Aranceles"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8970
   Icon            =   "frm_abmgrupos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   8970
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_consgpo 
      Caption         =   "data_consgpo"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6000
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data data_aran 
      Caption         =   "data_aran"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   5760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5760
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_abmgrupos.frx":058A
      Height          =   1695
      Left            =   240
      OleObjectBlob   =   "frm_abmgrupos.frx":05A4
      TabIndex        =   14
      Top             =   6360
      Width           =   8415
   End
   Begin VB.CommandButton b_inf 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4440
      Picture         =   "frm_abmgrupos.frx":0F8B
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5640
      Width           =   495
   End
   Begin VB.CommandButton b_elimina 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3600
      Picture         =   "frm_abmgrupos.frx":1515
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5640
      Width           =   495
   End
   Begin VB.CommandButton b_cancela 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      Picture         =   "frm_abmgrupos.frx":1A9F
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5640
      Width           =   495
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      Picture         =   "frm_abmgrupos.frx":2029
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5640
      Width           =   495
   End
   Begin VB.CommandButton b_edita 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1080
      Picture         =   "frm_abmgrupos.frx":25B3
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5640
      Width           =   495
   End
   Begin VB.CommandButton b_nuevo 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_abmgrupos.frx":2B3D
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5640
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos de Grupo"
      Enabled         =   0   'False
      Height          =   5415
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8535
      Begin VB.CommandButton b_borrt 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7800
         Picture         =   "frm_abmgrupos.frx":30C7
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Borrar todos los servicios de la lista"
         Top             =   4920
         Width           =   495
      End
      Begin VB.Data data_grupos 
         Caption         =   "data_grupos"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   3840
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   600
         Visible         =   0   'False
         Width           =   3135
      End
      Begin MSDBCtls.DBCombo DBCombo1 
         Bindings        =   "frm_abmgrupos.frx":3651
         Height          =   660
         Left            =   2160
         TabIndex        =   18
         Top             =   1320
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   1164
         _Version        =   393216
         Style           =   1
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Data data_servicios 
         Caption         =   "data_servicios"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   4200
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   240
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.TextBox t_obs 
         Height          =   855
         Left            =   2160
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   2040
         Width           =   5655
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "frm_abmgrupos.frx":366E
         Height          =   2415
         Left            =   240
         OleObjectBlob   =   "frm_abmgrupos.frx":3686
         TabIndex        =   15
         ToolTipText     =   "Los servicios que no están aquí se tomarán cómo particular. Con doble click puede eliminar de esta lista."
         Top             =   2880
         Width           =   7575
      End
      Begin VB.CommandButton b_del 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7920
         Picture         =   "frm_abmgrupos.frx":43BD
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Agregar TODOS los servicios"
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton b_add 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   7920
         Picture         =   "frm_abmgrupos.frx":4947
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Agregar servicio seleccionado"
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox t_desc 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         MaxLength       =   120
         TabIndex        =   4
         Top             =   840
         Width           =   5655
      End
      Begin VB.TextBox t_cod 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2160
         TabIndex        =   2
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   1680
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Observaciones:"
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
         Height          =   495
         Left            =   240
         TabIndex        =   17
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "Servicio:"
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
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Descripción:"
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
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Código:"
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
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   5400
      Picture         =   "frm_abmgrupos.frx":4ED1
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   2655
   End
End
Attribute VB_Name = "frm_abmgrupos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_add_Click()
'total = importe -(porcentaje*importe)

If t_cod.Text <> "" Then
   data_aran.Recordset.AddNew
   data_aran.Recordset("id_gpo") = t_cod.Text
   data_aran.Recordset("id_serv") = Val(Label5.Caption)
   data_aran.Recordset("desc_serv") = DBCombo1.Text
   data_aran.Recordset("prec_serv") = 0
   data_aran.Recordset("por_serv") = 100
   data_aran.Recordset.Update
   data_aran.RecordSource = "Select * from Aran_servicios where id_gpo =" & t_cod.Text
   data_aran.Refresh
End If

End Sub

Private Sub b_borrt_Click()
Dim Seguro As String
Seguro = MsgBox("Desea borrar el total de los servicios ingresados a este Grupo?", vbInformation + vbYesNo)
If Seguro = vbYes Then
    frm_abmgrupos.MousePointer = 11
    If t_cod.Text <> "" Then
       data_aran.RecordSource = "Select * from Aran_servicios where id_gpo =" & t_cod.Text
       data_aran.Refresh
       If data_aran.Recordset.RecordCount > 0 Then
          data_aran.Recordset.MoveFirst
          Do While Not data_aran.Recordset.EOF
             data_aran.Recordset.Delete
             data_aran.Recordset.MoveNext
          Loop
          data_aran.Refresh
       End If
    Else
       MsgBox "No hay grupo seleccionado", vbInformation
    End If
    frm_abmgrupos.MousePointer = 0
End If

End Sub

Private Sub b_cancela_Click()
t_cod.Text = ""
t_desc.Text = ""
t_obs.Text = ""
data_aran.RecordSource = "select * from Aran_servicios where id_gpo =" & 0
data_aran.Refresh


         b_nuevo.Enabled = True
         b_edita.Enabled = True
         b_graba.Enabled = False
         b_cancela.Enabled = False
         b_inf.Enabled = True
         b_elimina.Enabled = True
         DBGrid2.Enabled = False
         

End Sub

Private Sub b_del_Click()
Dim Xfami As String

If t_cod.Text <> "" Then
   Xfami = InputBox("Ingrese número de familia o 99 para TODOS")
   frm_abmgrupos.MousePointer = 11
   If Xfami = "99" Or Trim(Xfami) = "" Then
      data_servicios.RecordSource = "Select * from estudios where codest not in (999,30,31,881,882,990,991,992,993,994,995,996,997,8000) order by codest"
      data_servicios.Refresh
   Else
      data_servicios.RecordSource = "Select * from estudios where flia =" & Val(Xfami) & " and codest not in (999,30,31,881,882,990,991,992,993,994,995,996,997,8000) order by codest"
      data_servicios.Refresh
   End If
   data_servicios.Recordset.MoveFirst
   Do While Not data_servicios.Recordset.EOF
      If IsNull(data_servicios.Recordset("descrip")) = False Then
         If data_servicios.Recordset("descrip") <> "" Then
            data_aran.RecordSource = "Select * from Aran_servicios where id_gpo =" & t_cod.Text & " and id_serv =" & data_servicios.Recordset("codest")
            data_aran.Refresh
            If data_aran.Recordset.RecordCount > 0 Then
            Else
                data_aran.Recordset.AddNew
                data_aran.Recordset("id_gpo") = t_cod.Text
                data_aran.Recordset("id_serv") = data_servicios.Recordset("codest")
                data_aran.Recordset("desc_serv") = data_servicios.Recordset("descrip")
                data_aran.Recordset("prec_serv") = 0
                data_aran.Recordset("por_serv") = 100
                data_aran.Recordset.Update
            End If
         End If
      End If
      data_servicios.Recordset.MoveNext
   Loop
   data_aran.RecordSource = "Select * from Aran_servicios where id_gpo =" & t_cod.Text
   data_aran.Refresh
   frm_abmgrupos.MousePointer = 0
   
End If

End Sub

Private Sub b_eli_Click()

End Sub

Private Sub b_edita_Click()
If data_grupos.Recordset("id") = t_cod.Text Then
    b_nuevo.Enabled = False
    b_edita.Enabled = False
    b_graba.Enabled = True
    b_cancela.Enabled = True
    b_inf.Enabled = False
    b_elimina.Enabled = False
    DBGrid2.Enabled = True
    XAlta = 2
    Frame1.Enabled = True
    t_cod.Enabled = False
    t_desc.SetFocus
    
Else
   MsgBox "Seleccione en la lista de abajo el registro a editar", vbInformation
End If

End Sub

Private Sub b_elimina_Click()
Dim Xmensa As String
Xmensa = MsgBox("Desea borrar el registro seleccionado?", vbYesNo + vbInformation)
If Xmensa = vbYes Then
   If data_grupos.Recordset.RecordCount > 0 Then
      t_cod.Text = data_grupos.Recordset("id")
      data_grupos.Recordset.Delete
      data_aran.RecordSource = "Select * from Aran_servicios where id_gpo =" & t_cod.Text
      data_aran.Refresh
      If data_aran.Recordset.RecordCount > 0 Then
         data_aran.Recordset.MoveFirst
         Do While Not data_aran.Recordset.EOF
            data_aran.Recordset.Delete
            data_aran.Recordset.MoveNext
         Loop
      End If
      data_grupos.Refresh
      MsgBox "Registro eliminado.", vbInformation
   End If
Else

End If


End Sub

Private Sub b_graba_Click()

If XAlta = 1 Then
   If t_cod.Text <> "" And t_desc.Text <> "" Then
      data_consgpo.RecordSource = "select * from Aran_grupos where desc_gpo='" & t_desc.Text & "'"
      data_consgpo.Refresh
      If data_consgpo.Recordset.RecordCount > 0 Then
         MsgBox "Ya existe un grupo con este nombre", vbInformation
      Else
         data_aran.RecordSource = "Select * from Aran_servicios where id_gpo =" & t_cod.Text
         data_aran.Refresh
         If data_aran.Recordset.RecordCount > 0 Then
            data_grupos.Recordset.AddNew
            data_grupos.Recordset("desc_gpo") = t_desc.Text
            If t_obs.Text <> "" Then
               data_grupos.Recordset("obs_gpo") = t_obs.Text
            End If
            data_grupos.Recordset("fec_gpo") = Date
            data_grupos.Recordset("us_gpo") = WElusuario
            data_grupos.Recordset.Update
            data_grupos.Refresh
            b_nuevo.Enabled = True
            b_edita.Enabled = True
            b_graba.Enabled = False
            b_cancela.Enabled = False
            b_inf.Enabled = True
            b_elimina.Enabled = True
            DBGrid2.Enabled = False
         Else
            MsgBox "No hay servicios agregados para el grupo", vbInformation
         End If
      End If
      
   Else
      MsgBox "Falta descripción o código"
   End If

Else
   If t_cod.Text <> "" And t_desc.Text <> "" Then
      data_aran.RecordSource = "Select * from Aran_servicios where id_gpo =" & t_cod.Text
      data_aran.Refresh
      If data_aran.Recordset.RecordCount > 0 Then
         data_grupos.Recordset.Edit
         
         data_grupos.Recordset("desc_gpo") = t_desc.Text
         If t_obs.Text <> "" Then
            data_grupos.Recordset("obs_gpo") = t_obs.Text
         Else
            If IsNull(data_grupos.Recordset("obs_gpo")) = False Then
               data_grupos.Recordset("obs_gpo") = Null
            End If
         End If
         data_grupos.Recordset("fec_gpo") = Date
         data_grupos.Recordset("us_gpo") = WElusuario
         data_grupos.Recordset.Update
         data_grupos.Refresh
         b_nuevo.Enabled = True
         b_edita.Enabled = True
         b_graba.Enabled = False
         b_cancela.Enabled = False
         b_inf.Enabled = True
         b_elimina.Enabled = True
         DBGrid2.Enabled = False
      Else
         MsgBox "No hay servicios agregados para el grupo", vbInformation
      End If
   Else
      MsgBox "Falta descripción o código"
   End If

End If
End Sub

Private Sub b_inf_Click()
frm_infaran.Show vbModal

End Sub

Private Sub b_nuevo_Click()
XAlta = 1
Frame1.Enabled = True
t_cod.Text = ""
t_desc.Text = ""
DBCombo1.Text = ""
t_obs.Text = ""
If data_grupos.Recordset.RecordCount > 0 Then
   data_grupos.Recordset.MoveLast
   t_cod.Text = data_grupos.Recordset("id") + 1
Else
   t_cod.Text = 1
End If

t_cod.Enabled = False
data_aran.RecordSource = "Select * from Aran_servicios where id_gpo =" & t_cod.Text
data_aran.Refresh

t_desc.SetFocus
b_nuevo.Enabled = False
b_edita.Enabled = False
b_graba.Enabled = True
b_cancela.Enabled = True
b_inf.Enabled = False
b_elimina.Enabled = False
DBGrid2.Enabled = True

End Sub

Private Sub DBCombo1_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   If DBCombo1.Text <> "" Then
      DBCombo1.ListField = "DESCRIP"
      DBCombo1.BoundColumn = "DESCRIP"
      If IsNumeric(DBCombo1.Text) Then
         If DBCombo1.Text <> "" Then
            data_servicios.Recordset.FindFirst "codest =" & DBCombo1.Text
            If Not data_servicios.Recordset.NoMatch Then
               DBCombo1.Text = data_servicios.Recordset("descrip")
               Label5.Caption = data_servicios.Recordset("codest")
               DBCombo1.Height = 500
               DBCombo1.ListField = ""
               DBCombo1.BoundColumn = ""
               b_add.SetFocus
            Else
               data_servicios.RecordSource = "select * from estudios where codest >=" & DBCombo1.Text
               data_servicios.Refresh
               DBCombo1.Height = 1350
            End If
         Else
            data_servicios.RecordSource = "select * from estudios where descrip >='" & "A" & "'"
            data_servicios.Refresh
            DBCombo1.Height = 1350
         End If
      Else
         data_servicios.Recordset.FindFirst "descrip ='" & DBCombo1.Text & "'"
         If Not data_servicios.Recordset.NoMatch Then
            DBCombo1.Text = data_servicios.Recordset("descrip")
            Label5.Caption = data_servicios.Recordset("codest")
            DBCombo1.Height = 500
            DBCombo1.ListField = ""
            DBCombo1.BoundColumn = ""
            b_add.SetFocus
         Else
            data_servicios.RecordSource = "select * from estudios where descrip >='" & DBCombo1.Text & "'"
            data_servicios.Refresh
            DBCombo1.Height = 1350
         End If
      End If
   Else
      t_obs.SetFocus
   End If
End If



End Sub

Private Sub DBGrid1_DblClick()
If data_grupos.Recordset.RecordCount > 0 Then
    t_cod.Text = data_grupos.Recordset("id")
    If IsNull(data_grupos.Recordset("desc_gpo")) = False Then
       t_desc.Text = data_grupos.Recordset("desc_gpo")
    Else
       t_desc.Text = ""
    End If
    If IsNull(data_grupos.Recordset("obs_gpo")) = False Then
       t_obs.Text = data_grupos.Recordset("obs_gpo")
    Else
       t_obs.Text = ""
    End If
    data_aran.RecordSource = "Select * from Aran_servicios where id_gpo =" & t_cod.Text & " order by id_serv"
    data_aran.Refresh
End If

End Sub

Private Sub DBGrid2_DblClick()

data_aran.Recordset.Delete
data_aran.RecordSource = "Select * from Aran_servicios where id_gpo =" & t_cod.Text
data_aran.Refresh

End Sub

Private Sub Form_Load()
data_servicios.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_servicios.RecordSource = "estudios"
data_servicios.Refresh

data_grupos.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_grupos.RecordSource = "Aran_grupos"
data_grupos.Refresh

data_consgpo.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_grupos.RecordSource = "Aran_grupos"
'data_grupos.Refresh

data_aran.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_aran.RecordSource = "select * from Aran_servicios where id_gpo =" & 0 & " order by id_serv"
data_aran.Refresh


End Sub

Private Sub Form_Resize()
With Image1
     .Left = 0
     .Top = 0
     .Width = Me.Width
     .Height = Me.Height
End With

End Sub

Private Sub t_desc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   DBCombo1.SetFocus
End If

End Sub
