VERSION 5.00
Begin VB.Form frm_opcionesing 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   0  'None
   Caption         =   "Opciones"
   ClientHeight    =   2070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   2070
   ScaleWidth      =   9135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ver borrados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5640
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Salir"
      Height          =   735
      Left            =   8160
      Picture         =   "frm_opcionesing.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Matrículas repetidas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Socios con más de una ficha"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
End
Attribute VB_Name = "frm_opcionesing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

frm_opcionesing.Hide
frm_depuraficha.Show vbModal

End Sub

Private Sub Command2_Click()
'frm_opcionesing.Hide
'frm_repet.Show vbModal

End Sub

Private Sub Command3_Click()
End

End Sub

Private Sub Command4_Click()
frm_listados.Show vbModal

End Sub
