VERSION 5.00
Begin VB.Form frm_obseval 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   0  'None
   Caption         =   "Observaciones"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   1320
      Picture         =   "frm_obseval.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Eliminar la observación seleccionada"
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   6480
      Picture         =   "frm_obseval.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancelar observaciones"
      Top             =   2160
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   120
      Picture         =   "frm_obseval.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guardar observaciones"
      Top             =   2160
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   1575
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   6975
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ingrese observaciones y presione el botón de visto o cancelar para no registrar las observaciones."
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
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   3120
      Picture         =   "frm_obseval.frx":109E
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   2535
   End
End
Attribute VB_Name = "frm_obseval"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Xquecol = 1 Then
   If Text1.Text <> "" Then
      frm_evalua1.Text1.Text = Text1.Text
   Else
      frm_evalua1.Text1.Text = ""
   End If
End If

If Xquecol = 2 Then
   If Text1.Text <> "" Then
      frm_evalua1.Text2.Text = Text1.Text
   Else
      frm_evalua1.Text2.Text = ""
   End If
End If

If Xquecol = 3 Then
   If Text1.Text <> "" Then
      frm_evalua1.Text3.Text = Text1.Text
   Else
      frm_evalua1.Text3.Text = ""
   End If
End If

If Xquecol = 4 Then
   If Text1.Text <> "" Then
      frm_evalua1.Text4.Text = Text1.Text
   Else
      frm_evalua1.Text4.Text = ""
   End If
End If

If Xquecol = 5 Then
   If Text1.Text <> "" Then
      frm_evalua1.Text5.Text = Text1.Text
   Else
      frm_evalua1.Text5.Text = ""
   End If
End If

If Xquecol = 6 Then
   If Text1.Text <> "" Then
      frm_evalua1.Text6.Text = Text1.Text
   Else
      frm_evalua1.Text6.Text = ""
   End If
End If

If Xquecol = 7 Then
   If Text1.Text <> "" Then
      frm_evalua1.Text7.Text = Text1.Text
   Else
      frm_evalua1.Text7.Text = ""
   End If
End If

If Xquecol = 8 Then
   If Text1.Text <> "" Then
      frm_evalua1.Text8.Text = Text1.Text
   Else
      frm_evalua1.Text8.Text = ""
   End If
End If

Unload Me


End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
Dim Xborraono As String
Xborraono = MsgBox("Desea borrar la observación de la evaluación?", vbInformation + vbYesNo)
If Xborraono = vbYes Then
   If Xquecol = 1 Then
      frm_evalua1.Text1.Text = ""
      Text1.Text = ""
   End If
   If Xquecol = 2 Then
      frm_evalua1.Text2.Text = ""
      Text1.Text = ""
   End If
   Unload Me
End If

End Sub

Private Sub Form_Load()
If Xquecol = 1 Then
   If frm_evalua1.Text1.Text <> "" Then
      Text1.Text = frm_evalua1.Text1.Text
   End If
End If

If Xquecol = 2 Then
   If frm_evalua1.Text2.Text <> "" Then
      Text1.Text = frm_evalua1.Text2.Text
   End If
End If

If Xquecol = 3 Then
   If frm_evalua1.Text3.Text <> "" Then
      Text1.Text = frm_evalua1.Text3.Text
   End If
End If

If Xquecol = 4 Then
   If frm_evalua1.Text4.Text <> "" Then
      Text1.Text = frm_evalua1.Text4.Text
   End If
End If

If Xquecol = 5 Then
   If frm_evalua1.Text5.Text <> "" Then
      Text1.Text = frm_evalua1.Text5.Text
   End If
End If

If Xquecol = 6 Then
   If frm_evalua1.Text6.Text <> "" Then
      Text1.Text = frm_evalua1.Text6.Text
   End If
End If

If Xquecol = 7 Then
   If frm_evalua1.Text7.Text <> "" Then
      Text1.Text = frm_evalua1.Text7.Text
   End If
End If

If Xquecol = 8 Then
   If frm_evalua1.Text8.Text <> "" Then
      Text1.Text = frm_evalua1.Text8.Text
   End If
End If


End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub
