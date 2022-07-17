VERSION 5.00
Begin VB.Form frm_facbaja 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   0  'None
   Caption         =   "Datos para facturar BAJA"
   ClientHeight    =   1800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7680
      Picture         =   "frm_facbaja.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cancelar facturación."
      Top             =   0
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0080FF80&
      Caption         =   "ACTIVO en otro convenio"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Afiliación NUEVA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   855
      Left            =   3360
      Picture         =   "frm_facbaja.frx":058A
      Stretch         =   -1  'True
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "frm_facbaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frmabm.t_queconv.Text = ""
Unload Me
frm_afilbaja.Show vbModal

End Sub

Private Sub Command2_Click()
Dim Xqueauto As String
Xqueauto = InputBox("Ingrese CONVENIO del paciente:", "Facturación")
If Xqueauto <> "" Then
   frmabm.t_queconv.Text = Mid(Trim(Xqueauto), 1, 6)
   frmabm.data_clientes.Recordset.Edit
   frmabm.data_clientes.Recordset("cl_celular") = Mid(Xqueauto, 1, 12)
   frmabm.data_clientes.Recordset("cl_tipocli") = 1
   frmabm.data_clientes.Recordset("cl_fultvta") = Date
   frmabm.data_clientes.Recordset.Update
   Unload Me
   frmquefac.Show vbModal
Else
   frmabm.t_queconv.Text = ""
   Unload Me
End If

End Sub

Private Sub Command3_Click()
frmabm.t_queconv.Text = ""
frmabm.btn_fact.Enabled = True
Unload Me

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub
