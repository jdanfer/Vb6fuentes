VERSION 5.00
Begin VB.Form frm_creatablas 
   BackColor       =   &H00FF8080&
   Caption         =   "Crear Tablas"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   7905
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      Height          =   375
      Left            =   2640
      Picture         =   "frm_creatablas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   2640
      Picture         =   "frm_creatablas.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1920
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   120
      Width           =   1815
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Crear tabla"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nombres de campos"
      Height          =   375
      Left            =   2760
      TabIndex        =   5
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nombre de la tabla:"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Campos de la tabla:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frm_creatablas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub
