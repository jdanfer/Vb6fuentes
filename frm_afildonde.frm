VERSION 5.00
Begin VB.Form frm_afildonde 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Donde?"
   ClientHeight    =   2790
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7590
   Icon            =   "frm_afildonde.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2790
   ScaleWidth      =   7590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
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
      Left            =   7080
      Picture         =   "frm_afildonde.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cerrar sin seleccionar"
      Top             =   2400
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CONTINUAR"
      Height          =   615
      Left            =   2640
      Picture         =   "frm_afildonde.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2040
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   480
      ItemData        =   "frm_afildonde.frx":109E
      Left            =   3120
      List            =   "frm_afildonde.frx":10A0
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Caption         =   "Seleccione en que base realizará la factura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   6375
   End
End
Attribute VB_Name = "frm_afildonde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Combo1.ListIndex >= 0 Then
   frm_afilia.labfact.Caption = Combo1.Text
   Unload Me
   
Else
   MsgBox "Debe seleccionar una base"
   
End If
End Sub

Private Sub Command2_Click()
frm_afilia.labfact.Caption = ""

Unload Me

End Sub

Private Sub Form_Load()
Combo1.AddItem "1"
Combo1.AddItem "2"
Combo1.AddItem "3"
Combo1.AddItem "4"
Combo1.AddItem "6"
Combo1.AddItem "8"
Combo1.AddItem "11"
Combo1.AddItem "12"
Combo1.AddItem "13"
Combo1.AddItem "16"
Combo1.AddItem "17"
Combo1.AddItem "18"

End Sub
