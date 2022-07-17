VERSION 5.00
Begin VB.Form frm_opszona 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   Caption         =   "Opciones de zona"
   ClientHeight    =   2205
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7110
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   7110
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   0
      Picture         =   "frm_opszona.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Aceptar"
      Top             =   1680
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      ItemData        =   "frm_opszona.frx":058A
      Left            =   2520
      List            =   "frm_opszona.frx":05BB
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Grupos de zonas:"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Opciones de zona"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7095
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   1440
      Picture         =   "frm_opszona.frx":0647
      Stretch         =   -1  'True
      Top             =   960
      Width           =   1815
   End
End
Attribute VB_Name = "frm_opszona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Wopszon = Combo1.ListIndex
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
