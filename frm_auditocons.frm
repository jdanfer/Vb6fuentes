VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_auditocons 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Movimientos en la solicitud de auditoría"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10080
   Icon            =   "frm_auditocons.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   10080
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_histo 
      Caption         =   "data_histo"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data data_grabahis 
      Caption         =   "data_grabahis"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_auditocons.frx":0442
      Height          =   1575
      Left            =   120
      OleObjectBlob   =   "frm_auditocons.frx":045B
      TabIndex        =   15
      Top             =   4800
      Width           =   9735
   End
   Begin VB.CommandButton b_cancela 
      BackColor       =   &H00FF8080&
      Enabled         =   0   'False
      Height          =   735
      Left            =   3720
      Picture         =   "frm_auditocons.frx":0FDE
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Cancelar ingreso"
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H00FF8080&
      Enabled         =   0   'False
      Height          =   735
      Left            =   2520
      Picture         =   "frm_auditocons.frx":1420
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Grabar los datos ingresados"
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton b_modif 
      BackColor       =   &H00FF8080&
      Height          =   735
      Left            =   1320
      Picture         =   "frm_auditocons.frx":1862
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Modificar datos"
      Top             =   3960
      Width           =   855
   End
   Begin VB.CommandButton b_nuevo 
      BackColor       =   &H00FF8080&
      Height          =   735
      Left            =   120
      Picture         =   "frm_auditocons.frx":1CA4
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Nuevo registro"
      Top             =   3960
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C000&
      Caption         =   "Historial de movimientos de la solicitud"
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
      ForeColor       =   &H0080FF80&
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      Begin VB.TextBox txt_dethis 
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
         Left            =   2280
         MaxLength       =   130
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   2280
         Width           =   7095
      End
      Begin MSMask.MaskEdBox mhorahis 
         Height          =   375
         Left            =   6000
         TabIndex        =   9
         Top             =   1560
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfechis 
         Height          =   375
         Left            =   2280
         TabIndex        =   7
         Top             =   1560
         Width           =   1815
         _ExtentX        =   3201
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
      Begin VB.Label labus 
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   18
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   7080
         TabIndex        =   17
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Análisis:"
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
         TabIndex        =   10
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label7 
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
         Left            =   4560
         TabIndex        =   8
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label Label6 
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
         Left            =   240
         TabIndex        =   6
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         BorderWidth     =   3
         X1              =   0
         X2              =   9720
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label labtit 
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
         Left            =   2280
         TabIndex        =   4
         Top             =   960
         Width           =   7215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Socio:"
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
         TabIndex        =   3
         Top             =   960
         Width           =   1935
      End
      Begin VB.Label labnro 
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
         Left            =   2280
         TabIndex        =   2
         Top             =   480
         Width           =   1815
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
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Doble click para editar"
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
      TabIndex        =   16
      Top             =   6360
      Width           =   4455
   End
End
Attribute VB_Name = "frm_auditocons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_cancela_Click()
XAlta = 0
b_nuevo.Enabled = True
b_graba.Enabled = False
b_modif.Enabled = True
b_cancela.Enabled = False
DBGrid1.Enabled = True
mfechis.Text = "__/__/____"
mhorahis.Text = "__:__"
txt_dethis.Text = ""
Frame1.Enabled = False


End Sub

Private Sub b_graba_Click()
If XAlta = 1 Then
'   If Combo1.ListIndex <> -1 Then
      If txt_dethis.Text <> "" Then
         data_grabahis.Recordset.AddNew
         data_grabahis.Recordset("cl_etiquet") = 0
         data_grabahis.Recordset("cl_codigo") = 99
         data_grabahis.Recordset("cl_cantdia") = 8
         data_grabahis.Recordset("cl_cantpag") = labnro.Caption
         If mfechis.Text <> "__/__/____" Then
            data_grabahis.Recordset("cl_fultvta") = Format(mfechis.Text, "dd/mm/yyyy")
         End If
         data_grabahis.Recordset("cl_fax") = mhorahis.Text
         data_grabahis.Recordset("info_debit") = txt_dethis.Text
'         data_grabahis.Recordset("cl_dircobr") = txt_pasos.Text
         data_grabahis.Recordset("cl_nom_sup") = labus.Caption
         data_grabahis.Recordset.Update
         data_grabahis.Refresh
         data_histo.Refresh
         XAlta = 0
         b_nuevo.Enabled = True
         b_graba.Enabled = False
         b_modif.Enabled = True
         b_cancela.Enabled = False
         DBGrid1.Enabled = True
         mfechis.Text = "__/__/____"
         mhorahis.Text = "__:__"
         txt_dethis.Text = ""
         Frame1.Enabled = False
         DBGrid1.SetFocus
      Else
         MsgBox "Ingrese una descripción"
      End If
'   Else
'      MsgBox "Seleccione proceso"
'   End If
Else
   data_histo.Recordset.Edit
'   If Combo1.ListIndex <> -1 Then
'      data_histo.Recordset("cl_localid") = Combo1.Text
'   End If
'   data_histo.Recordset("cl_forpago") = Combo1.ListIndex
'   data_histo.Recordset("cl_fultvta") = mfechis.Text
'   data_histo.Recordset("cl_fax") = mhorahis.Text
   data_histo.Recordset("info_debit") = txt_dethis.Text
   data_histo.Recordset("cl_nom_sup") = labus.Caption
'   If txt_plazo.Text = "" Then
'      txt_plazo.Text = 0
'      labpla.Caption = Date
'   End If
'   data_histo.Recordset("cl_codced") = txt_plazo.Text
'   data_histo.Recordset("cl_fultmov") = labpla.Caption
   data_histo.Recordset.Update
   data_histo.Refresh
   data_grabahis.Refresh
   XAlta = 0
   mfechis.Enabled = True
   mhorahis.Enabled = True
   b_nuevo.Enabled = True
   b_graba.Enabled = False
   b_modif.Enabled = True
   b_cancela.Enabled = False
   DBGrid1.Enabled = True
   mfechis.Text = "__/__/____"
   mhorahis.Text = "__:__"
   txt_dethis.Text = ""
   Frame1.Enabled = False
   DBGrid1.SetFocus
   
End If


End Sub

Private Sub b_modif_Click()
'If WElusuario = frm_mejora.data_accion.Recordset("cl_nomvend") Then
    XAlta = 0
    Frame1.Enabled = True
    b_nuevo.Enabled = False
    b_graba.Enabled = True
    b_modif.Enabled = False
    b_cancela.Enabled = True
    DBGrid1.Enabled = False
    mfechis.Enabled = False
    mhorahis.Enabled = False
    
    txt_dethis.SetFocus
'Else
'    MsgBox "NO ES EL USUARIO PROPIETARIO DE LA ACCION", vbCritical
'    DBGrid1.SetFocus
'End If

End Sub

Private Sub b_nuevo_Click()
'If WElusuario = frm_mejora.data_accion.Recordset("cl_nomvend") Then
    Frame1.Enabled = True
    XAlta = 1
    b_nuevo.Enabled = False
    b_graba.Enabled = True
    b_modif.Enabled = False
    b_cancela.Enabled = True
    DBGrid1.Enabled = False
    mfechis.Text = "__/__/____"
    mhorahis.Text = "__:__"
    txt_dethis.Text = ""
    mfechis.Text = Format(Date, "dd/mm/yyyy")
    mhorahis.Text = Format(Time, "HH:mm")
'Else
'    MsgBox "NO ES EL USUARIO PROPIETARIO DE LA ACCION", vbCritical
'    DBGrid1.SetFocus
'End If

End Sub


Private Sub DBGrid1_DblClick()
If IsNull(data_histo.Recordset("cl_fultvta")) = False Then
   mfechis.Text = Format(data_histo.Recordset("cl_fultvta"), "dd/mm/yyyy")
Else
   mfechis.Text = "__/__/____"
End If
If IsNull(data_histo.Recordset("cl_fax")) = False Then
   mhorahis.Text = Format(data_histo.Recordset("cl_fax"), "HH:mm")
Else
   mhorahis.Text = "__:__"
End If
If IsNull(data_histo.Recordset("info_debit")) = False Then
   txt_dethis.Text = data_histo.Recordset("info_debit")
Else
   txt_dethis.Text = ""
End If


End Sub

Private Sub Form_Load()
labnro.Caption = frm_solaudito.txt_nro.Text
labtit.Caption = frm_solaudito.txt_nom.Text

If labnro.Caption = "" Then
   MsgBox "No existen registros"
Else
    data_histo.DatabaseName = App.Path & "\sapp.mdb"
    data_histo.RecordSource = "Select * from env_soc where cl_cantpag =" & labnro.Caption
    data_histo.Refresh
    data_grabahis.DatabaseName = App.Path & "\sapp.mdb"
    data_grabahis.RecordSource = "Select * from env_soc where cl_cantpag =" & labnro.Caption
    data_grabahis.Refresh
    labus.Caption = WElusuario
    If data_histo.Recordset.RecordCount > 0 Then
    Else
       MsgBox "No se han ingresado datos a esta solicitud", vbInformation
    End If
End If

End Sub

Private Sub mfechis_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mhorahis.SetFocus
End If

End Sub

Private Sub mhorahis_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_dethis.SetFocus
End If

End Sub


Private Sub txt_dethis_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_graba.SetFocus
End If

End Sub
