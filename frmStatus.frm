VERSION 5.00
Begin VB.Form frmStatus 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6105
   DrawStyle       =   5  'Transparent
   Icon            =   "frmStatus.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.TextBox txtStatus 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2415
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "frmStatus.frx":044A
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "frmStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
'txtStatus.Height = Me.Height - 400
'txtStatus.Width = Me.Width - 150
End Sub
