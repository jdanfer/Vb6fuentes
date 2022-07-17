VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_medicos 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Médicos"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7365
   Icon            =   "frm_medicos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   7365
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Se incluye en informes"
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
      Left            =   4080
      TabIndex        =   24
      Top             =   2280
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   4080
      TabIndex        =   23
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Médico de Cooperativa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3720
      TabIndex        =   22
      Top             =   2760
      Width           =   2415
   End
   Begin VB.Data data_parammed 
      Caption         =   "data_parammed"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2760
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_medhc 
      Caption         =   "data_medhc"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox t_codced 
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
      Height          =   285
      Left            =   3480
      TabIndex        =   21
      Top             =   2280
      Width           =   375
   End
   Begin VB.TextBox t_ced 
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
      Height          =   285
      Left            =   2160
      TabIndex        =   20
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Data data_buscamed 
      Caption         =   "data_buscamed"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.TextBox t_codhc 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      TabIndex        =   18
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txt_espec 
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
      Height          =   285
      Left            =   2160
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1320
      Width           =   2535
   End
   Begin VB.TextBox txt_tel 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2160
      MaxLength       =   30
      TabIndex        =   5
      Top             =   1800
      Width           =   2535
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   6240
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      RecordSource    =   "medicos"
      Top             =   1200
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_medicos.frx":0442
      Height          =   1695
      Left            =   120
      OleObjectBlob   =   "frm_medicos.frx":0459
      TabIndex        =   14
      Top             =   4560
      Width           =   6855
   End
   Begin VB.TextBox txt_bcob 
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
      Height          =   285
      Left            =   2280
      TabIndex        =   13
      Top             =   4200
      Width           =   3735
   End
   Begin VB.Data data_cob 
      Caption         =   "data_cob"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "medicos"
      Top             =   2400
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton bimp 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4320
      Picture         =   "frm_medicos.frx":0FE8
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Informes"
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton bbusca 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3480
      Picture         =   "frm_medicos.frx":1572
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Buscar"
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton bcance 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      Picture         =   "frm_medicos.frx":1AFC
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Cancelar acción"
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton bmodif 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1800
      Picture         =   "frm_medicos.frx":2086
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Modificar datos"
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton bgraba 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   960
      Picture         =   "frm_medicos.frx":2610
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Guardar datos"
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton bnuevo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      Picture         =   "frm_medicos.frx":2B9A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Nuevo registro"
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox txt_nomcob 
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
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   840
      Width           =   4695
   End
   Begin VB.TextBox txt_nrocob 
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
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "CEDULA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Código HC:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ESPECIALIDAD:"
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
      Left            =   240
      TabIndex        =   16
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Teléfonos:"
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
      Left            =   240
      TabIndex        =   15
      Top             =   1800
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Nombre a buscar:"
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
      TabIndex        =   12
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   7320
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   7320
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Nombre:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Cód.Médico:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin VB.Image Image1 
      Height          =   1695
      Left            =   5280
      Picture         =   "frm_medicos.frx":3124
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   1695
   End
End
Attribute VB_Name = "frm_medicos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bbusca_Click()
txt_bcob.Enabled = True
DBGrid1.Enabled = True
txt_bcob.SetFocus

End Sub

Private Sub bcance_Click()
If XAcnv = 1 Then
   data_cob.Recordset.CancelUpdate
   igualcob
   XAcnv = 0
   desh
Else
   igualcob
   XAcnv = 0
   desh
End If
bgraba.Enabled = False
bcance.Enabled = False
bmodif.Enabled = True
bbusca.Enabled = True
bimp.Enabled = True
bnuevo.Enabled = True

End Sub

Private Sub bgraba_Click()
On Error GoTo Algrab
Dim Lacedtexto As String

If txt_nrocob.Text <> "" And txt_nomcob.Text <> "" And t_ced.Text <> "" And t_codced.Text <> "" Then
   If txt_nrocob.Text <> 0 Then
         If XAcnv = 1 Then
            data_buscamed.RecordSource = "Select * from medicos where med_nombre ='" & txt_nomcob.Text & "'"
            data_buscamed.Refresh
            If data_buscamed.Recordset.RecordCount > 0 Then
               MsgBox "El médico ingresado ya existe, VERIFIQUE!!", vbCritical
            Else
               data_cob.Recordset("med_cod") = txt_nrocob.Text
               data_cob.Recordset("med_nombre") = txt_nomcob.Text
               data_cob.Recordset("med_esp") = txt_espec.Text
               data_cob.Recordset("med_socnom") = txt_tel.Text
               data_cob.Recordset("incentivo") = Check2.Value
               If t_codhc.Text <> "" Then
                  data_cob.Recordset("med_socnro") = t_codhc.Text
               End If
               data_cob.Recordset.Update
               data_medhc.RecordSource = "Select * from meta_tres where m_mat =" & txt_nrocob.Text
               data_medhc.Refresh
               If data_medhc.Recordset.RecordCount > 0 Then
                  data_medhc.Recordset.Edit
                  data_medhc.Recordset("m_nrofrm") = Trim(t_ced.Text) & Trim(t_codced.Text)
                  data_medhc.Recordset("m_codmed") = Check1.Value
                  data_medhc.Recordset.Update
               Else
                  data_parammed.Recordset.Edit
                  data_parammed.Recordset("nro_reg") = data_parammed.Recordset("nro_reg") + 1
                  data_parammed.Recordset.Update
                  data_parammed.Refresh
                  data_medhc.Recordset.AddNew
                  data_medhc.Recordset("id") = data_parammed.Recordset("nro_reg")
                  data_medhc.Recordset("m_fecha") = Date
                  data_medhc.Recordset("m_mat") = txt_nrocob.Text
                  data_medhc.Recordset("m_nrofrm") = Trim(t_ced.Text) & Trim(t_codced.Text)
                  data_medhc.Recordset("m_codmed") = Check1.Value
                  data_medhc.Recordset.Update
               End If
               XAcnv = 0
               Data1.Refresh
               bgraba.Enabled = False
               bcance.Enabled = False
               bmodif.Enabled = True
               bbusca.Enabled = True
               bimp.Enabled = True
               bnuevo.Enabled = True
               desh
            End If
         Else
'            data_cob.Recordset("med_cod") = txt_nrocob.Text
            data_cob.Recordset.Edit
            data_cob.Recordset("med_nombre") = txt_nomcob.Text
            data_cob.Recordset("med_esp") = txt_espec.Text
            data_cob.Recordset("med_socnom") = txt_tel.Text
            data_cob.Recordset("incentivo") = Check2.Value
            data_cob.Recordset.Update
            If t_codhc.Text <> "" Then
               If IsNull(data_cob.Recordset("med_socnro")) = False Then
                  If data_cob.Recordset("med_socnro") <> t_codhc.Text Then
                     data_cob.Recordset.Edit
                     data_cob.Recordset("med_socnro") = t_codhc.Text
                     data_cob.Recordset.Update
                  End If
               Else
                  data_cob.Recordset.Edit
                  data_cob.Recordset("med_socnro") = t_codhc.Text
                  data_cob.Recordset.Update
               End If
            End If
            data_medhc.RecordSource = "Select * from meta_tres where m_mat =" & txt_nrocob.Text
            data_medhc.Refresh
            If data_medhc.Recordset.RecordCount > 0 Then
               Lacedtexto = Trim(t_ced.Text) & Trim(t_codced.Text)
               If Trim(data_medhc.Recordset("m_nrofrm")) <> Trim(Lacedtexto) Then
                  data_medhc.Recordset.Edit
                  data_medhc.Recordset("m_nrofrm") = Trim(t_ced.Text) & Trim(t_codced.Text)
                  data_medhc.Recordset.Update
               End If
               If data_medhc.Recordset("m_codmed") <> Check1.Value Then
                  data_medhc.Recordset.Edit
                  data_medhc.Recordset("m_codmed") = Check1.Value
                  data_medhc.Recordset.Update
               End If
            Else
               data_parammed.Recordset.Edit
               data_parammed.Recordset("nro_reg") = data_parammed.Recordset("nro_reg") + 1
               data_parammed.Recordset.Update
               data_parammed.Refresh
               data_medhc.Recordset.AddNew
               data_medhc.Recordset("id") = data_parammed.Recordset("nro_reg")
               data_medhc.Recordset("m_fecha") = Date
               data_medhc.Recordset("m_mat") = txt_nrocob.Text
               data_medhc.Recordset("m_nrofrm") = Trim(t_ced.Text) & Trim(t_codced.Text)
               data_medhc.Recordset("m_codmed") = Check1.Value
               data_medhc.Recordset.Update
            End If
            
            XAcnv = 0
            Data1.Refresh
            bgraba.Enabled = False
            bcance.Enabled = False
            bmodif.Enabled = True
            bbusca.Enabled = True
            bimp.Enabled = True
            bnuevo.Enabled = True
            txt_nrocob.Enabled = True
            desh
         End If
   Else
      MsgBox "No ingresó médico", vbCritical, "Médicos"
      txt_nrocob.SetFocus
   End If
Else
   MsgBox "Faltan datos para poder grabar!", vbCritical, "Médicos"
   txt_nrocob.SetFocus
End If

Exit Sub

Algrab:
      If Err.Number = 3155 Then
         MsgBox "Error al grabar, verifique datos"
      Else
         MsgBox "Error al grabar, verifique datos"
      End If
      
End Sub

Private Sub bimp_Click()
'CrystalReport1.Action = 1

End Sub

Private Sub bmodif_Click()
If XWeltipoU = "ADMINISTRADOR" Then
    XAcnv = 0
    hab
    txt_nrocob.Enabled = False
    txt_nomcob.SetFocus
    bgraba.Enabled = True
    bcance.Enabled = True
    bmodif.Enabled = False
    bbusca.Enabled = False
    bimp.Enabled = False
    bnuevo.Enabled = False
Else
    MsgBox "No se permite modificar"
    Unload Me
    
End If

End Sub

Private Sub bnuevo_Click()
XAcnv = 1
hab
txt_nrocob.Text = ""
txt_nomcob.Text = ""
txt_tel.Text = ""
txt_espec.Text = ""
t_codhc.Text = ""
t_ced.Text = ""
t_codced.Text = ""
Check1.Value = 0
Check2.Value = 0
txt_nomcob.SetFocus
txt_nrocob.Enabled = False
bgraba.Enabled = True
bcance.Enabled = True
bmodif.Enabled = False
bbusca.Enabled = False
bimp.Enabled = False
bnuevo.Enabled = False
Data1.RecordSource = "Select * from medicos order by med_cod"
Data1.Refresh
Data1.Recordset.MoveLast
txt_nrocob.Text = Data1.Recordset("med_cod") + 1
data_cob.Recordset.AddNew

End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
On Error GoTo Albus

If KeyAscii = 13 Then
    If IsNull(data_cob.Recordset("med_cod")) = False Then
       txt_nrocob.Text = data_cob.Recordset("med_cod")
    Else
       txt_nrocob.Text = ""
    End If
    If IsNull(data_cob.Recordset("med_nombre")) = False Then
       txt_nomcob.Text = data_cob.Recordset("med_nombre")
    Else
       txt_nomcob.Text = ""
    End If
    If IsNull(data_cob.Recordset("incentivo")) = False Then
       Check2.Value = data_cob.Recordset("incentivo")
    Else
       Check2.Value = 0
    End If
    
    If IsNull(data_cob.Recordset("med_esp")) = False Then
       txt_espec.Text = data_cob.Recordset("med_esp")
    Else
       txt_espec.Text = ""
    End If
    If IsNull(data_cob.Recordset("med_socnom")) = False Then
       txt_tel.Text = data_cob.Recordset("med_socnom")
    Else
       txt_tel.Text = ""
    End If
    If IsNull(data_cob.Recordset("med_socnro")) = False Then
       t_codhc.Text = data_cob.Recordset("med_socnro")
    Else
       t_codhc.Text = ""
    End If

    If txt_nrocob.Text <> "" Then
       data_medhc.RecordSource = "Select * from meta_tres where m_mat =" & txt_nrocob.Text
       data_medhc.Refresh
       If data_medhc.Recordset.RecordCount > 0 Then
          If IsNull(data_medhc.Recordset("m_nrofrm")) = False Then
             If Len(Trim(data_medhc.Recordset("m_nrofrm"))) = 7 Then
                t_ced.Text = Mid(data_medhc.Recordset("m_nrofrm"), 1, 6)
                t_codced.Text = Mid(data_medhc.Recordset("m_nrofrm"), 7, 1)
             Else
                If Len(Trim(data_medhc.Recordset("m_nrofrm"))) = 8 Then
                   t_ced.Text = Mid(data_medhc.Recordset("m_nrofrm"), 1, 7)
                   t_codced.Text = Mid(data_medhc.Recordset("m_nrofrm"), 8, 1)
                Else
                   t_ced.Text = ""
                   t_codced.Text = ""
                End If
             End If
          Else
             t_ced.Text = ""
             t_codced.Text = ""
          End If
          If IsNull(data_medhc.Recordset("m_codmed")) = False Then
             Check1.Value = data_medhc.Recordset("m_codmed")
          End If
       Else
          t_ced.Text = ""
          t_codced.Text = ""
       End If
    End If


End If
txt_bcob.Enabled = False
DBGrid1.Enabled = False
bmodif.SetFocus

Exit Sub

Albus:
     If Err.Number = 3155 Then
        MsgBox "Error al buscar"
     Else
        MsgBox "Error al buscar"
     End If
     
End Sub

Private Sub Form_Initialize()
data_cob.Recordset.MoveLast
If IsNull(data_cob.Recordset("med_cod")) = False Then
   txt_nrocob.Text = data_cob.Recordset("med_cod")
Else
   txt_nrocob.Text = ""
End If
If IsNull(data_cob.Recordset("med_nombre")) = False Then
   txt_nomcob.Text = data_cob.Recordset("med_nombre")
Else
   txt_nomcob.Text = ""
End If
If IsNull(data_cob.Recordset("incentivo")) = False Then
   Check2.Value = data_cob.Recordset("incentivo")
Else
   Check2.Value = 0
End If
If IsNull(data_cob.Recordset("med_esp")) = False Then
   txt_espec.Text = data_cob.Recordset("med_esp")
Else
   txt_espec.Text = ""
End If
If IsNull(data_cob.Recordset("med_socnom")) = False Then
   txt_tel.Text = data_cob.Recordset("med_socnom")
Else
   txt_tel.Text = ""
End If
If IsNull(data_cob.Recordset("med_socnro")) = False Then
   t_codhc.Text = data_cob.Recordset("med_socnro")
Else
   t_codhc.Text = ""
End If
data_buscamed.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_medhc.Connect = "odbc;dsn=" & Xconexrmt & ";"
If txt_nrocob.Text <> "" Then
   data_medhc.RecordSource = "Select * from meta_tres where m_mat =" & txt_nrocob.Text
   data_medhc.Refresh
   If data_medhc.Recordset.RecordCount > 0 Then
      If IsNull(data_medhc.Recordset("m_nrofrm")) = False Then
         If Len(Trim(data_medhc.Recordset("m_nrofrm"))) = 7 Then
            t_ced.Text = Mid(data_medhc.Recordset("m_nrofrm"), 1, 6)
            t_codced.Text = Mid(data_medhc.Recordset("m_nrofrm"), 7, 1)
         Else
            If Len(Trim(data_medhc.Recordset("m_nrofrm"))) = 8 Then
               t_ced.Text = Mid(data_medhc.Recordset("m_nrofrm"), 1, 7)
               t_codced.Text = Mid(data_medhc.Recordset("m_nrofrm"), 8, 1)
            Else
               t_ced.Text = ""
               t_codced.Text = ""
            End If
         End If
      Else
         t_ced.Text = ""
         t_codced.Text = ""
      End If
      If IsNull(data_medhc.Recordset("m_codmed")) = False Then
         Check1.Value = data_medhc.Recordset("m_codmed")
      End If
   End If
End If

End Sub

Public Function hab()
txt_nrocob.Enabled = True
txt_nomcob.Enabled = True
txt_tel.Enabled = True
txt_espec.Enabled = True
t_codhc.Enabled = True
t_ced.Enabled = True
t_codced.Enabled = True
Check1.Enabled = True
Check2.Enabled = True
End Function

Public Function desh()
txt_nrocob.Enabled = False
txt_nomcob.Enabled = False
txt_tel.Enabled = False
txt_espec.Enabled = False
t_codhc.Enabled = False
t_ced.Enabled = False
t_codced.Enabled = False
Check1.Enabled = False
Check2.Enabled = False

End Function

Private Sub Form_Load()
data_cob.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
CrystalReport1.ReportFileName = App.path & "\medicos.rpt"

data_parammed.DatabaseName = App.path & "\paramhoras.mdb"
data_parammed.RecordSource = "parsec0"
data_parammed.Refresh

data_medhc.Connect = "odbc;dsn=" & Xconexrmt & ";"

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub t_ced_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_codced.SetFocus
End If

End Sub

Private Sub txt_bcob_Change()
data_cob.RecordSource = "select * from medicos where med_nombre >='" & txt_bcob.Text & "' order by med_nombre"
data_cob.Refresh

End Sub

Private Sub txt_bcob_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))
If KeyAscii = 13 Then
   KeyAscii = 0
   DBGrid1.SetFocus
End If

End Sub

Public Function igualcob()
data_cob.Recordset.MoveLast
If IsNull(data_cob.Recordset("med_cod")) = False Then
   txt_nrocob.Text = data_cob.Recordset("med_cod")
Else
   txt_nrocob.Text = ""
End If
If IsNull(data_cob.Recordset("med_nombre")) = False Then
   txt_nomcob.Text = data_cob.Recordset("med_nombre")
Else
   txt_nomcob.Text = ""
End If
If IsNull(data_cob.Recordset("incentivo")) = False Then
   Check2.Value = data_cob.Recordset("incentivo")
Else
   Check2.Value = 0
End If
If IsNull(data_cob.Recordset("med_esp")) = False Then
   txt_espec.Text = data_cob.Recordset("med_esp")
Else
   txt_espec.Text = ""
End If
If IsNull(data_cob.Recordset("med_socnom")) = False Then
   txt_tel.Text = data_cob.Recordset("med_socnom")
Else
   txt_tel.Text = ""
End If
If IsNull(data_cob.Recordset("med_socnro")) = False Then
   t_codhc.Text = data_cob.Recordset("med_socnro")
Else
   t_codhc.Text = ""
End If
If txt_nrocob.Text <> "" Then
   data_medhc.RecordSource = "Select * from meta_tres where m_mat =" & txt_nrocob.Text
   data_medhc.Refresh
   If data_medhc.Recordset.RecordCount > 0 Then
      If IsNull(data_medhc.Recordset("m_nrofrm")) = False Then
         If Len(Trim(data_medhc.Recordset("m_nrofrm"))) = 7 Then
            t_ced.Text = Mid(data_medhc.Recordset("m_nrofrm"), 1, 6)
            t_codced.Text = Mid(data_medhc.Recordset("m_nrofrm"), 7, 1)
         Else
            If Len(Trim(data_medhc.Recordset("m_nrofrm"))) = 8 Then
               t_ced.Text = Mid(data_medhc.Recordset("m_nrofrm"), 1, 7)
               t_codced.Text = Mid(data_medhc.Recordset("m_nrofrm"), 8, 1)
            Else
               t_ced.Text = ""
               t_codced.Text = ""
            End If
         End If
      Else
         t_ced.Text = ""
         t_codced.Text = ""
      End If
      If IsNull(data_medhc.Recordset("m_codmed")) = False Then
         Check1.Value = data_medhc.Recordset("m_codmed")
      End If
   Else
      t_ced.Text = ""
      t_codced.Text = ""
   End If
End If


End Function


Private Sub txt_espec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_tel.SetFocus
End If

End Sub

Private Sub txt_nomcob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_espec.SetFocus
End If

End Sub

Private Sub txt_nrocob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_nomcob.SetFocus
End If

End Sub

Private Sub txt_nrocob_LostFocus()
If XAcnv = 1 Then
   Data1.Recordset.FindFirst "med_cod =" & txt_nrocob.Text
   If Not Data1.Recordset.NoMatch Then
      MsgBox "Ya existe este número de médico", vbCritical, "Médicos"
   End If
End If

End Sub

Private Sub txt_tel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_ced.SetFocus
End If

End Sub
