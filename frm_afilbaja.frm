VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_afilbaja 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Afiliación"
   ClientHeight    =   4035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6345
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4035
   ScaleWidth      =   6345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_cliaf 
      Caption         =   "data_cliaf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   2460
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancelar"
      Height          =   735
      Left            =   3120
      Picture         =   "frm_afilbaja.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aceptar"
      Height          =   735
      Left            =   360
      Picture         =   "frm_afilbaja.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3120
      Width           =   1815
   End
   Begin MSMask.MaskEdBox mfaf 
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   2400
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.TextBox t_cataf 
      Height          =   375
      Left            =   2520
      MaxLength       =   12
      TabIndex        =   6
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox t_nroaf 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   2520
      MaxLength       =   7
      TabIndex        =   0
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   0
      X2              =   6360
      Y1              =   3000
      Y2              =   3000
   End
   Begin VB.Label Label6 
      BackColor       =   &H00C00000&
      Caption         =   "Fecha Afiliación:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C00000&
      Caption         =   "Categoría:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C00000&
      Caption         =   "Número Afiliación:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label labnomaf 
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   5775
   End
   Begin VB.Label labmataf 
      BackColor       =   &H00C00000&
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "AFILIACION NUEVA A:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   4800
      Picture         =   "frm_afilbaja.frx":0B14
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "frm_afilbaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If t_nroaf.Text <> "" Then
   If t_cataf.Text <> "" Then
      If mfaf.Text <> "__/__/____" Then
         data_cliaf.Recordset.Edit
         data_cliaf.Recordset("cl_celular") = t_cataf.Text
         data_cliaf.Recordset("cl_tipocli") = Val(t_nroaf.Text)
         data_cliaf.Recordset("cl_fultvta") = Format(mfaf.Text, "yyyy-mm-dd")
         data_cliaf.Recordset.Update
         Unload Me
         frmquefac.Show vbModal
      End If
   End If
End If
      
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
'data_cliaf.DatabaseName = App.Path & "\sapp.mdb"
data_cliaf.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_cliaf.RecordSource = "clientes"
'data_cliaf.Refresh
data_cliaf.RecordSource = "Select * from clientes where cl_codigo =" & frmabm.txt_mat.Caption
data_cliaf.Refresh
'data_cliaf.Recordset.FindFirst "cl_codigo =" & frmabm.txt_mat.Caption
'If Not data_cliaf.Recordset.NoMatch Then
If data_cliaf.Recordset.RecordCount > 0 Then
   labmataf.Caption = data_cliaf.Recordset("cl_codigo")
   labnomaf.Caption = data_cliaf.Recordset("cl_apellid")
Else
   labmataf.Caption = 0
   labnomaf.Caption = "SIN NOM"
End If

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Width = Me.Width
    .Height = Me.Height
End With

End Sub

Private Sub mfaf_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command1.SetFocus
End If

End Sub

Private Sub t_cataf_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfaf.SetFocus
End If

End Sub

Private Sub t_nroaf_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_cataf.SetFocus
End If

End Sub
