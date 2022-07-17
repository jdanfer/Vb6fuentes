VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_moviles 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Móviles"
   ClientHeight    =   5460
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   8640
   Icon            =   "frm_moviles.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   8640
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport cr1 
      Left            =   3600
      Top             =   2400
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
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4200
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data data_enf 
      Caption         =   "data_enf"
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
      Top             =   5040
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data data_chof 
      Caption         =   "data_chof"
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
      Top             =   5160
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton b_elimina 
      BackColor       =   &H00C0E0FF&
      Height          =   735
      Left            =   5040
      Picture         =   "frm_moviles.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Elimina el registro del móvil seleccionado"
      Top             =   4320
      Width           =   855
   End
   Begin VB.Data data_med 
      Caption         =   "data_med"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4800
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.Data data_mov 
      Caption         =   "data_mov"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4800
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.CommandButton b_imprime 
      BackColor       =   &H00C0E0FF&
      Height          =   735
      Left            =   7440
      Picture         =   "frm_moviles.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton b_busca 
      BackColor       =   &H00C0E0FF&
      Height          =   735
      Left            =   6240
      Picture         =   "frm_moviles.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Buscar datos"
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton b_cance 
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      Height          =   735
      Left            =   3840
      Picture         =   "frm_moviles.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Cancelar la acción a realizar"
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      Height          =   735
      Left            =   2640
      Picture         =   "frm_moviles.frx":154A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Graba los datos"
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton b_modif 
      BackColor       =   &H00C0E0FF&
      Height          =   735
      Left            =   1440
      Picture         =   "frm_moviles.frx":198C
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Modificar datos"
      Top             =   4320
      Width           =   855
   End
   Begin VB.CommandButton b_nuevo 
      BackColor       =   &H00C0E0FF&
      Height          =   735
      Left            =   240
      Picture         =   "frm_moviles.frx":1DCE
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Agregar registro NUEVO"
      Top             =   4320
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Datos de móviles"
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
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.TextBox t_enf 
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
         Left            =   2280
         TabIndex        =   24
         Top             =   2430
         Width           =   4335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Buscar..."
         Height          =   375
         Left            =   6840
         TabIndex        =   23
         Top             =   2400
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar..."
         Height          =   375
         Left            =   6840
         TabIndex        =   22
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox t_chof 
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
         Left            =   2280
         TabIndex        =   18
         Top             =   1800
         Width           =   4335
      End
      Begin VB.TextBox t_base 
         Alignment       =   1  'Right Justify
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
         Left            =   5640
         TabIndex        =   16
         Top             =   3000
         Width           =   975
      End
      Begin MSMask.MaskEdBox mfec 
         Height          =   375
         Left            =   2280
         TabIndex        =   7
         Top             =   3000
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
      Begin VB.TextBox txt_codmed 
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
         Left            =   1320
         TabIndex        =   5
         Top             =   1320
         Visible         =   0   'False
         Width           =   855
      End
      Begin MSDBCtls.DBCombo dbmedic 
         Bindings        =   "frm_moviles.frx":2210
         Height          =   360
         Left            =   2280
         TabIndex        =   4
         Top             =   1080
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "MED_NOMBRE"
         BoundColumn     =   "MED_NOMBRE"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txt_nro 
         Alignment       =   1  'Right Justify
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
         Height          =   405
         Left            =   2280
         MaxLength       =   5
         TabIndex        =   2
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label6 
         Height          =   255
         Left            =   4320
         TabIndex        =   25
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label labenf 
         Height          =   255
         Left            =   960
         TabIndex        =   21
         Top             =   2760
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Enfermería:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label labchof 
         Height          =   255
         Left            =   960
         TabIndex        =   19
         Top             =   2160
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Chófer:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1920
         Width           =   2055
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Base Fact:"
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
         Left            =   4200
         TabIndex        =   15
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Fecha actualización"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   3120
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Médico:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Nro. MOVIL:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frm_moviles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub b_busca_Click()
frm_buscamovil.Show vbModal

End Sub

Private Sub b_cance_Click()
If XAlta = 1 Then
   data_mov.Recordset.CancelUpdate
End If
XAlta = 0
data_mov.Recordset.MoveFirst
habilitamov
igualacuadros

Frame1.Enabled = False

End Sub

Private Sub b_elimina_Click()
Dim Xrespumov As String
Xrespumov = MsgBox("Desea eliminar el registro del móvil?", vbYesNo, "Móviles")
If Xrespumov = vbYes Then
   data_mov.Recordset.Delete
   data_mov.Refresh
   borra_mov
   If data_mov.Recordset.RecordCount > 0 Then
      igualacuadros
   End If
End If

End Sub

Private Sub b_graba_Click()
If txt_nro.Text <> "" Then
   If txt_nro.Text > 0 Then
      If XAlta = 1 Then
         data_mov.Recordset("nroreg") = Label6.Caption
         data_mov.Recordset("movil") = txt_nro.Text
         If txt_codmed.Text = "" Then
            txt_codmed.Text = 0
         End If
         data_mov.Recordset("codmed") = txt_codmed.Text
         If dbmedic.Text = "" Then
         Else
            data_mov.Recordset("nommed") = dbmedic.Text
         End If
         If mfec.Text = "__/__/____" Then
         Else
            data_mov.Recordset("fecha_act") = mfec.Text
         End If
         data_mov.Recordset("hora_act") = Format(Time, "HH:mm")
         If t_base.Text = "" Then
            t_base.Text = 0
         End If
         data_mov.Recordset("ano") = t_base.Text
         If labchof.Caption = "" Then
            labchof.Caption = 0
         End If
         data_mov.Recordset("codchof") = labchof.Caption
         If t_chof.Text = "" Then
         Else
            data_mov.Recordset("nomchof") = t_chof.Text
         End If
         If labenf.Caption = "" Then
            labenf.Caption = 0
         End If
         data_mov.Recordset("codenf") = labenf.Caption
         If t_enf.Text = "" Then
         Else
            data_mov.Recordset("nomenf") = t_enf.Text
         End If
         data_mov.Recordset.Update
         data_mov.Refresh
         borra_mov
         XAlta = 0
      End If
      If XAlta = 2 Then
         data_mov.Recordset.Edit
         data_mov.Recordset("movil") = txt_nro.Text
         If txt_codmed.Text = "" Then
            txt_codmed.Text = 0
         End If
         data_mov.Recordset("codmed") = txt_codmed.Text
         If dbmedic.Text = "" Then
         Else
            data_mov.Recordset("nommed") = dbmedic.Text
         End If
         If mfec.Text = "__/__/____" Then
         Else
            data_mov.Recordset("fecha_act") = mfec.Text
         End If
         data_mov.Recordset("hora_act") = Format(Time, "HH:mm")
         If t_base.Text = "" Then
            t_base.Text = 0
         End If
         data_mov.Recordset("ano") = t_base.Text
         If labchof.Caption = "" Then
            labchof.Caption = 0
         End If
         data_mov.Recordset("codchof") = labchof.Caption
         If t_chof.Text = "" Then
         Else
            data_mov.Recordset("nomchof") = t_chof.Text
         End If
         If labenf.Caption = "" Then
            labenf.Caption = 0
         End If
         data_mov.Recordset("codenf") = labenf.Caption
         If t_enf.Text = "" Then
         Else
            data_mov.Recordset("nomenf") = t_enf.Text
         End If
         data_mov.Recordset.Update
         data_mov.Refresh
         XAlta = 0
      End If
      borra_mov
      habilitamov
      Frame1.Enabled = False
   Else
      MsgBox "No puede ser CERO el número de móvil", vbInformation, "Móviles"
   End If
Else
   MsgBox "No ingresó número de móvil", vbInformation, "Mensaje"

End If

End Sub

Private Sub b_imprime_Click()
If data_inf.Recordset.RecordCount > 0 Then
   data_inf.Recordset.MoveFirst
   Do While Not data_inf.Recordset.EOF
      data_inf.Recordset.Delete
      data_inf.Recordset.MoveNext
   Loop
   data_inf.Refresh
End If

If data_mov.Recordset.RecordCount > 0 Then
   data_mov.Recordset.MoveFirst
   Do While Not data_mov.Recordset.EOF
      data_inf.Recordset.AddNew
      data_inf.Recordset("cl_codigo") = data_mov.Recordset("movil")
      data_inf.Recordset("cl_apellid") = Mid(data_mov.Recordset("nommed"), 1, 30)
      data_inf.Recordset("cl_nombre") = Mid(data_mov.Recordset("nomchof"), 1, 30)
      data_inf.Recordset("cl_localid") = Mid(data_mov.Recordset("nomenf"), 1, 30)
      data_inf.Recordset("cl_fnac") = data_mov.Recordset("fecha_Act")
      data_inf.Recordset.Update
      data_mov.Recordset.MoveNext
   Loop
   data_inf.RecordSource = "Select * from infcli"
   data_inf.Refresh
   cr1.ReportFileName = App.Path & "\infmovile.rpt"
   cr1.ReportTitle = "INFORME DE MOVILES ACTUALIZADOS"
   cr1.Action = 1
End If

borra_mov

End Sub

Private Sub b_modif_Click()
deshabmov
XAlta = 2
Frame1.Enabled = True
txt_nro.SetFocus

End Sub

Private Sub b_nuevo_Click()
deshabmov
borra_mov
Frame1.Enabled = True
txt_nro.SetFocus
XAlta = 1
If data_mov.Recordset.RecordCount > 0 Then
   data_mov.Recordset.MoveLast
   Label6.Caption = data_mov.Recordset("nroreg") + 1
Else
   Label6.Caption = 1
End If
data_mov.Recordset.AddNew

End Sub

Private Sub Command1_Click()
frm_chofer.Show vbModal

End Sub

Private Sub Command2_Click()
frm_enferm.Show vbModal

End Sub

Private Sub dbmedic_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_chof.SetFocus
End If

End Sub

Private Sub dbmedic_LostFocus()
If IsNumeric(dbmedic.Text) = True Then
   If Val(dbmedic.Text) > 0 Then
      data_med.Recordset.FindFirst "med_cod =" & dbmedic.Text
      If Not data_med.Recordset.NoMatch Then
         dbmedic.Text = data_med.Recordset("med_nombre")
         txt_codmed.Text = data_med.Recordset("med_cod")
      Else
         MsgBox "No encontrado, busque por nombre", vbInformation, "Médicos"
         dbmedic.SetFocus
      End If
   Else
      data_med.Recordset.FindFirst "med_nombre ='" & dbmedic.Text & "'"
      If Not data_med.Recordset.NoMatch Then
         dbmedic.Text = data_med.Recordset("med_nombre")
         txt_codmed.Text = data_med.Recordset("med_cod")
      Else
         MsgBox "No encontrado, VERIFIQUE NOMBRE", vbInformation, "Médicos"
         dbmedic.SetFocus
      End If
   End If
Else
   If Len(dbmedic.Text) > 0 Then
      data_med.Recordset.FindFirst "med_nombre ='" & dbmedic.Text & "'"
      If Not data_med.Recordset.NoMatch Then
         dbmedic.Text = data_med.Recordset("med_nombre")
         txt_codmed.Text = data_med.Recordset("med_cod")
      Else
         MsgBox "No encontrado, VERIFIQUE NOMBRE", vbInformation, "Médicos"
         dbmedic.SetFocus
      End If
   End If
End If

End Sub





Private Sub Form_Load()

End Sub

Private Sub mfec_GotFocus()
mfec.Text = Date

End Sub

Private Sub mfec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_base.SetFocus
End If

End Sub




Private Sub t_base_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_graba.SetFocus
End If

End Sub

Private Sub t_chof_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_enf.SetFocus
End If

End Sub

Private Sub t_chof_LostFocus()
If t_chof.Text <> "" Then
   If IsNumeric(t_chof.Text) = True Then
      data_chof.Recordset.FindFirst "nromov =" & t_chof.Text
      If Not data_chof.Recordset.NoMatch Then
         t_chof.Text = data_chof.Recordset("chofer")
         labchof.Caption = data_chof.Recordset("nromov")
      Else
         t_chof.Text = ""
         labchof.Caption = ""
      End If
   Else
      data_chof.Recordset.FindFirst "chofer ='" & t_chof.Text & "'"
      If Not data_chof.Recordset.NoMatch Then
         t_chof.Text = data_chof.Recordset("chofer")
         labchof.Caption = data_chof.Recordset("nromov")
      Else
         t_chof.Text = ""
         labchof.Caption = ""
      End If
   End If
End If

End Sub

Private Sub t_enf_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfec.SetFocus
End If

End Sub

Private Sub t_enf_LostFocus()
If t_enf.Text <> "" Then
   If IsNumeric(t_enf.Text) = True Then
      data_enf.Recordset.FindFirst "id =" & t_enf.Text
      If Not data_enf.Recordset.NoMatch Then
         t_enf.Text = data_enf.Recordset("nomb")
         labenf.Caption = data_enf.Recordset("id")
      Else
         t_enf.Text = ""
         labenf.Caption = ""
      End If
   Else
      data_enf.Recordset.FindFirst "nomb ='" & t_enf.Text & "'"
      If Not data_enf.Recordset.NoMatch Then
         t_enf.Text = data_enf.Recordset("nomb")
         labenf.Caption = data_enf.Recordset("id")
      Else
         t_enf.Text = ""
         labenf.Caption = ""
      End If
   End If
End If

End Sub

Private Sub txt_nro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   dbmedic.SetFocus
End If

End Sub

Private Sub txt_nro_LostFocus()
If txt_nro.Text <> "" Then
Else
   MsgBox "No Ingresó móvil"
   txt_nro.SetFocus
End If

End Sub

Private Sub txt_seguro_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_graba.SetFocus
End If

End Sub

Public Sub habilitamov()
b_nuevo.Enabled = True
b_modif.Enabled = True
b_graba.Enabled = False
b_imprime.Enabled = True
b_busca.Enabled = True
b_cance.Enabled = False
b_elimina.Enabled = True

End Sub

Public Sub deshabmov()
b_nuevo.Enabled = False
b_modif.Enabled = False
b_graba.Enabled = True
b_imprime.Enabled = False
b_busca.Enabled = False
b_cance.Enabled = True
b_elimina.Enabled = False

End Sub

Public Sub igualacuadros()
txt_nro.Text = data_mov.Recordset("movil")
If IsNull(data_mov.Recordset("codmed")) = False Then
   txt_codmed.Text = data_mov.Recordset("codmed")
Else
   txt_codmed.Text = 0
End If
If IsNull(data_mov.Recordset("nommed")) = False Then
   dbmedic.Text = data_mov.Recordset("nommed")
Else
   dbmedic.Text = ""
End If
If IsNull(data_mov.Recordset("fecha_act")) = False Then
   mfec.Text = Format(data_mov.Recordset("fecha_act"), "dd/mm/yyyy")
Else
   mfec.Text = "__/__/____"
End If
If IsNull(data_mov.Recordset("ano")) = False Then
   t_base.Text = data_mov.Recordset("ano")
Else
   t_base.Text = 0
End If
If IsNull(data_mov.Recordset("codchof")) = False Then
   labchof.Caption = data_mov.Recordset("codchof")
Else
   labchof.Caption = 0
End If
If IsNull(data_mov.Recordset("nomchof")) = False Then
   t_chof.Text = data_mov.Recordset("nomchof")
Else
   t_chof.Text = ""
End If
If IsNull(data_mov.Recordset("codenf")) = False Then
   labenf.Caption = data_mov.Recordset("codenf")
Else
   labenf.Caption = 0
End If
If IsNull(data_mov.Recordset("nomenf")) = False Then
   t_enf.Text = data_mov.Recordset("nomenf")
Else
   t_enf.Text = ""
End If


End Sub

Public Function borra_mov()
txt_nro.Text = ""
txt_codmed.Text = ""
dbmedic.Text = ""
mfec.Text = "__/__/____"
t_base.Text = ""
labchof.Caption = ""
t_chof.Text = ""
labenf.Caption = ""
t_enf.Text = ""


End Function
