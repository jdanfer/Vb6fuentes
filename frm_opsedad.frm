VERSION 5.00
Begin VB.Form frm_opsedad 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   Caption         =   "Opciones de edad"
   ClientHeight    =   3270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Incluir socios sin datos de Fecha de nacimiento"
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
      TabIndex        =   8
      Top             =   1200
      Width           =   5895
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frm_opsedad.frx":0000
      Left            =   3480
      List            =   "frm_opsedad.frx":0028
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   1800
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
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
      Left            =   240
      Picture         =   "frm_opsedad.frx":0090
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Aceptar"
      Top             =   2640
      Width           =   495
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Cumpleaños del mes de:"
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
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   3255
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Rango de edad a informar:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Value           =   -1  'True
      Width           =   3255
   End
   Begin VB.TextBox t_eh 
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
      Left            =   4560
      TabIndex        =   2
      Text            =   "110"
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox t_ed 
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
      Left            =   3480
      TabIndex        =   1
      Text            =   "0"
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Rango de 0 a 110 emite todos los socios ordenado por edad."
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
      TabIndex        =   7
      Top             =   840
      Width           =   5895
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      BorderWidth     =   3
      X1              =   0
      X2              =   6720
      Y1              =   2520
      Y2              =   2520
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF8080&
      Caption         =   "Opciones para edad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6735
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   2160
      Picture         =   "frm_opsedad.frx":061A
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   1575
   End
End
Attribute VB_Name = "frm_opsedad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1.value = True Then
   Wopsedd = Val(t_ed.Text)
   Wopsedh = Val(t_eh.Text)
   If Check1.value = 1 Then
      Wopsed = 3
   Else
      Wopsed = 1
   End If
Else
   If Option2.value = True Then
      Wopsed = 2
      Wopsedd = Combo1.ListIndex + 1
      Wopsedh = 0
   Else
      Wopsed = 0
      Wopsedd = 0
      Wopsedh = 0
   End If
End If
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

Private Sub t_ed_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_eh.SetFocus
End If

End Sub

Private Sub t_eh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command1.SetFocus
End If

End Sub
