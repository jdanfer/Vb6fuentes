VERSION 5.00
Begin VB.Form frm_mensajedesp 
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "mensajedesp"
   ClientHeight    =   5040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8190
   FillColor       =   &H0000FFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   5040
   ScaleWidth      =   8190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7680
      Picture         =   "frm_mensajedesp.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cerrar"
      Top             =   4680
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Llamado con numeración repetida. El llamado igual se guardará. VERIFIQUE y avise a informática."
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   2415
      Left            =   480
      TabIndex        =   1
      Top             =   2040
      Width           =   7095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ATENCIÓN!!!"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   48.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1215
      Left            =   1080
      TabIndex        =   0
      Top             =   480
      Width           =   6135
   End
End
Attribute VB_Name = "frm_mensajedesp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

