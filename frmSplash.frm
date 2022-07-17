VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00404040&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   3960
      Top             =   1920
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404000&
      Height          =   4050
      Left            =   150
      TabIndex        =   0
      Top             =   60
      Width           =   7080
      Begin VB.Image Image1 
         Height          =   1215
         Left            =   2400
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   2175
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Win XP/Vista/7/8/10"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   360
         Left            =   3960
         TabIndex        =   3
         Top             =   3120
         Width           =   2910
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Versión 05.2021v8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4710
         TabIndex        =   1
         Top             =   3600
         Width           =   2085
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sistema de Gestion para la salud"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   435
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   5655
      End
      Begin VB.Image Image2 
         Height          =   3855
         Left            =   120
         Picture         =   "frmSplash.frx":1D38
         Stretch         =   -1  'True
         Top             =   120
         Width           =   6855
      End
   End
   Begin VB.Image Image3 
      Height          =   2955
      Left            =   120
      Picture         =   "frmSplash.frx":2B09
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   3825
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
    frm_usuario.Show vbModal
End Sub

Private Sub Form_Resize()
With Image3
    .Left = 0
    .Top = 0
    .Width = Me.Width
    .Height = Me.Height
End With

End Sub

Private Sub Frame1_Click()
    Unload Me
    frm_usuario.Show vbModal
End Sub

Private Sub Timer1_Timer()
Unload Me
frm_usuario.Show vbModal

End Sub
