VERSION 5.00
Begin VB.Form frm_motivocance 
   BackColor       =   &H00400000&
   BorderStyle     =   0  'None
   Caption         =   "Motivos de cancelacion"
   ClientHeight    =   2730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   2730
   ScaleWidth      =   5880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5280
      Picture         =   "frm_motivocance.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Salir"
      Top             =   2160
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      Picture         =   "frm_motivocance.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Grabar datos"
      Top             =   2160
      Width           =   375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Ingrese motivo de CANCELACION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5415
      Begin VB.TextBox t_obs 
         Height          =   615
         Left            =   1320
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   1200
         Width           =   3735
      End
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         ItemData        =   "frm_motivocance.frx":0B14
         Left            =   120
         List            =   "frm_motivocance.frx":0B1E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   4935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Observaciones:"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frm_motivocance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Xind, Xcant As Integer
Dim Xcountt As Integer

If t_obs.Text <> "" And Combo1.ListIndex >= 0 Then
   For Xind = 1 To frm_especialistas.ListView1.ListItems.count
       frm_especialistas.ListView1.ListItems(Xind).Selected = True
       If frm_especialistas.ListView1.ListItems.Item(frm_especialistas.ListView1.SelectedItem.Index).Checked = True Then
          Xcant = Xcant + 1
       End If
   Next Xind
   Xind = 0
   If Xcant = 1 Then
      For Xind = 1 To frm_especialistas.ListView1.ListItems.count
          frm_especialistas.ListView1.ListItems(Xind).Selected = True
          If frm_especialistas.ListView1.ListItems.Item(frm_especialistas.ListView1.SelectedItem.Index).Checked = True Then
             Xnro = frm_especialistas.ListView1.ListItems(Xind).Text
             frm_especialistas.data_lista.RecordSource = "select * from t_fechas where fecha ='" & frm_especialistas.t_feccab.Text & "' and cod_cons =" & frm_especialistas.t_codcons.Text & " and nro =" & Xnro
             frm_especialistas.data_lista.Refresh
             If frm_especialistas.data_lista.Recordset.RecordCount > 0 Then
                frm_especialistas.data_borrados.RecordSource = "Select * from borrados where id <=" & 1
                frm_especialistas.data_borrados.Refresh
                frm_especialistas.data_borrados.Recordset.AddNew
                frm_especialistas.data_borrados.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
                frm_especialistas.data_borrados.Recordset("fecha_cons") = frm_especialistas.data_lista.Recordset("fecha")
                frm_especialistas.data_borrados.Recordset("medico") = frm_especialistas.data_lista.Recordset("nom_med")
                frm_especialistas.data_borrados.Recordset("local") = frm_especialistas.data_lista.Recordset("base")
                frm_especialistas.data_borrados.Recordset("hora") = frm_especialistas.data_lista.Recordset("hora_com")
                frm_especialistas.data_borrados.Recordset("hora_cons") = frm_especialistas.data_lista.Recordset("hora")
                frm_especialistas.data_borrados.Recordset("cedula") = frm_especialistas.data_lista.Recordset("ced_pac")
                frm_especialistas.data_borrados.Recordset("obs") = Combo1.Text
                frm_especialistas.data_borrados.Recordset("obs2") = t_obs.Text
                frm_especialistas.data_borrados.Recordset.Update
                                      
                frm_especialistas.data_lista.Recordset.Edit
                frm_especialistas.data_lista.Recordset("mat_pac") = Null
                frm_especialistas.data_lista.Recordset("nom_pac") = Null
                frm_especialistas.data_lista.Recordset("ced_pac") = Null
                frm_especialistas.data_lista.Recordset("convenio") = Null
                frm_especialistas.data_lista.Recordset("cel_pac") = Null
                frm_especialistas.data_lista.Recordset("tel_pac") = Null
                frm_especialistas.data_lista.Recordset("fec_nac") = Null
                frm_especialistas.data_lista.Recordset("hcsiono") = Null
                frm_especialistas.data_lista.Recordset("tipo_cons") = Null
                frm_especialistas.data_lista.Recordset("tipo_consd") = Null
                frm_especialistas.data_lista.Recordset("fec_anota") = Null
                frm_especialistas.data_lista.Recordset("hora_anota") = Null
                frm_especialistas.data_lista.Recordset("usua_anota") = Null
                frm_especialistas.data_lista.Recordset("edad") = Null
                frm_especialistas.data_lista.Recordset("usua_web") = Null
                frm_especialistas.data_lista.Recordset.Update
                frm_especialistas.t_mat.Text = ""
                frm_especialistas.t_ced.Text = ""
                frm_especialistas.t_codced.Text = ""
                frm_especialistas.t_nompac.Text = ""
                frm_especialistas.t_celu.Text = ""
                frm_especialistas.t_tellinea.Text = ""
                frm_especialistas.mfnac.Text = "__/__/____"
                frm_especialistas.t_conv.Text = ""
                frm_especialistas.cbotipcons.ListIndex = -1
                frm_especialistas.cbosino.ListIndex = -1

                frm_especialistas.data_lista.RecordSource = "select * from t_fechas where fecha ='" & frm_especialistas.data_cabfec.Recordset("fecha") & "' and cod_cons =" & frm_especialistas.data_cabfec.Recordset("cod_cons") & " order by nro"
                frm_especialistas.data_lista.Refresh
                Xcountt = 1
                If frm_especialistas.data_lista.Recordset.RecordCount > 0 Then
                   frm_especialistas.data_lista.Recordset.MoveFirst
                   frm_especialistas.ListView1.ListItems.Clear
                   Do While Not frm_especialistas.data_lista.Recordset.EOF
                      If IsNull(frm_especialistas.data_lista.Recordset("nro")) = False Then
                         frm_especialistas.ListView1.ListItems.Add Xcountt, , frm_especialistas.data_lista.Recordset("nro")
                      Else
                         frm_especialistas.ListView1.ListItems.Add Xcountt, , "0"
                      End If
                      If IsNull(frm_especialistas.data_lista.Recordset("hora")) = False Then
                         frm_especialistas.ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , frm_especialistas.data_lista.Recordset("hora")
                      Else
                         frm_especialistas.ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "00:00"
                      End If
                      If IsNull(frm_especialistas.data_lista.Recordset("ced_pac")) = False Then
                         frm_especialistas.ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , frm_especialistas.data_lista.Recordset("ced_pac")
                      Else
                         frm_especialistas.ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                      End If
                      If IsNull(frm_especialistas.data_lista.Recordset("nom_pac")) = False Then
                         frm_especialistas.ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , frm_especialistas.data_lista.Recordset("nom_pac")
                      Else
                         frm_especialistas.ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                      End If
                      If IsNull(frm_especialistas.data_lista.Recordset("convenio")) = False Then
                         frm_especialistas.ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , frm_especialistas.data_lista.Recordset("convenio")
                      Else
                         frm_especialistas.ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                      End If
                      If IsNull(frm_especialistas.data_lista.Recordset("cel_pac")) = False Then
                         frm_especialistas.ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , frm_especialistas.data_lista.Recordset("cel_pac")
                      Else
                         frm_especialistas.ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                      End If
                      If IsNull(frm_especialistas.data_lista.Recordset("tel_pac")) = False Then
                         frm_especialistas.ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , frm_especialistas.data_lista.Recordset("tel_pac")
                      Else
                         frm_especialistas.ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                      End If
                      If IsNull(frm_especialistas.data_lista.Recordset("tipo_consd")) = False Then
                         frm_especialistas.ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , frm_especialistas.data_lista.Recordset("tipo_consd")
                      Else
                         frm_especialistas.ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , " "
                      End If
                      If IsNull(frm_especialistas.data_lista.Recordset("usua_web")) = False Then
                         If frm_especialistas.data_lista.Recordset("usua_web") = "SI" Then
                            frm_especialistas.ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "WEB"
                         Else
                            frm_especialistas.ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "SAPP"
                         End If
                      Else
                         frm_especialistas.ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , "SAPP"
                      End If
                      frm_especialistas.data_lista.Recordset.MoveNext
                      Xcountt = Xcountt + 1
                   Loop
                End If
                MsgBox "La cosulta del paciente ha sido cancelada"
                Unload Me
             Else
                MsgBox "No se encuentra registro"
             End If
          End If
      Next
   Else
      MsgBox "Debe seleccionar UN SOLO REGISTRO"
   End If
End If

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

