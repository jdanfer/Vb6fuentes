VERSION 5.00
Begin VB.Form frm_ordenanza 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "CANCELAR"
      Height          =   735
      Left            =   4080
      Picture         =   "frm_ordenanza.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2160
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "ACEPTAR"
      Height          =   735
      Left            =   480
      Picture         =   "frm_ordenanza.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2160
      Width           =   1455
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frm_ordenanza.frx":0B14
      Left            =   2280
      List            =   "frm_ordenanza.frx":0B1E
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1440
      Width           =   3975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Seleccione Opción:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   $"frm_ordenanza.frx":0B55
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
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6015
   End
End
Attribute VB_Name = "frm_ordenanza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'frm_largador.checkorden = 1
'frm_largador.t_ordnro.Text = Combo1.ListIndex
'frm_largador.t_ordtext.Text = Combo1.Text


Unload Me

End Sub

Private Sub Command2_Click()
'frm_largador.checkorden = 0
'frm_largador.t_ordnro.Text = ""
'frm_largador.t_ordtext.Text = ""


Unload Me

End Sub

