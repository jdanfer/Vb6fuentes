VERSION 5.00
Begin VB.Form frm_proresp 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Proceso de respaldos..."
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6450
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_proresp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   6000
      Left            =   5040
      Top             =   1800
   End
   Begin VB.CommandButton b_proc 
      Caption         =   "Procesar..."
      Height          =   615
      Left            =   2040
      TabIndex        =   0
      Top             =   2760
      Width           =   2655
   End
End
Attribute VB_Name = "frm_proresp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_proc_Click()
If Dir("C:\resp\respcont.zip") <> "" Then
   Kill "c:\resp\respcont.zip"
End If
If Dir("C:\resp\respsue.zip") <> "" Then
   Kill "c:\resp\respsue.zip"
End If

Shell (App.Path & "\pkzip -rp c:\resp\respcont.zip z:\wincont\*.*"), vbNormalFocus
Shell (App.Path & "\pkzip -rp c:\resp\respsue.zip c:\software\lidesu\*.*"), vbNormalFocus

End

End Sub

Private Sub Timer1_Timer()
b_proc_Click
Timer1.Enabled = False

End Sub
