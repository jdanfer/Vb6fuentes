VERSION 5.00
Begin VB.Form frmUserPass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuario y Contraseña"
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   855
   ClientWidth     =   3780
   Icon            =   "frmUserPass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   3780
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   480
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtPass 
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   8.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   1200
      PasswordChar    =   "="
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.TextBox txtUser 
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Contraseña:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   765
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Usuario:"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   285
      Width           =   855
   End
End
Attribute VB_Name = "frmUserPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
frmMain.chkAuth.Value = 0
Unload Me
End Sub

Private Sub cmdOK_Click()
frmMain.aUSUARIO = txtUser
frmMain.aCONTRASEÑA = txtPass
Unload Me
End Sub

