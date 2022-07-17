VERSION 5.00
Begin VB.Form frm_calcula 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calcular importe total de servicios"
   ClientHeight    =   3000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6315
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_calcula.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   6315
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton b_igu 
      Caption         =   "="
      Height          =   495
      Left            =   5760
      TabIndex        =   11
      Top             =   480
      Width           =   495
   End
   Begin VB.CommandButton b_div 
      Caption         =   "/"
      Height          =   495
      Left            =   5160
      TabIndex        =   9
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton b_menos 
      Caption         =   "-"
      Height          =   495
      Left            =   4440
      TabIndex        =   8
      Top             =   840
      Width           =   495
   End
   Begin VB.CommandButton b_por 
      Caption         =   "*"
      Height          =   495
      Left            =   5160
      TabIndex        =   7
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton b_mas 
      Caption         =   "+"
      Height          =   495
      Left            =   4440
      TabIndex        =   6
      Top             =   240
      Width           =   495
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3960
      Picture         =   "frm_calcula.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2400
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0000FFFF&
      Height          =   495
      Left            =   2400
      Picture         =   "frm_calcula.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2400
      Width           =   615
   End
   Begin VB.TextBox t_pre 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label labs 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      Height          =   375
      Left            =   3840
      TabIndex        =   12
      Top             =   600
      Width           =   375
   End
   Begin VB.Label labtotdes 
      BackColor       =   &H0080FF80&
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Label labtot 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000FF00&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2400
      TabIndex        =   3
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H0080FF80&
      Caption         =   "TOTAL:"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FF80&
      Caption         =   "VALOR:"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
End
Attribute VB_Name = "frm_calcula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Xsigno As Integer
Public Xvalo1, Xvalo2 As Double

Private Sub b_div_Click()
labs.Caption = "/"
Xsigno = 4
If t_pre.Text <> "" Then
   If t_pre.Text = 0 Then
   Else
      Xvalo1 = CDbl(t_pre.Text)
      If labtot.Caption = "" Then
         labtot.Caption = t_pre.Text
      Else
         Xvalo2 = CDbl(labtot.Caption)
         If Xvalo2 = 0 Then
            labtot.Caption = 0
         Else
            labtot.Caption = Xvalo1 / Xvalo2
         End If
      End If
      labtot.Caption = Format(labtot.Caption, "Standard")
      labtotdes.Caption = labtotdes & " " & t_pre.Text
   End If
End If

End Sub

Private Sub b_igu_Click()
labs.Caption = ""
If t_pre.Text <> "" Then
   Xvalo1 = CDbl(t_pre.Text)
   If labtot.Caption = "" Then
      labtot.Caption = t_pre.Text
   Else
      If Xsigno = 2 Then
         Xvalo2 = CDbl(labtot.Caption)
         labtot.Caption = Xvalo1 * Xvalo2
      Else
         If Xsigno = 1 Then
            Xvalo2 = CDbl(labtot.Caption)
            labtot.Caption = Xvalo1 + Xvalo2
         Else
            If Xsigno = 3 Then
               Xvalo2 = CDbl(labtot.Caption)
               labtot.Caption = Xvalo1 - Xvalo2
            Else
               If Xsigno = 4 Then
                  Xvalo2 = CDbl(labtot.Caption)
                  If Xvalo2 = 0 Then
                     labtot.Caption = labtot.Caption
                  Else
                     labtot.Caption = Xvalo1 / Xvalo2
                  End If
               End If
            End If
         End If
      End If
   End If
   labtot.Caption = Format(labtot.Caption, "Standard")
   If Xsigno = 2 Then
      labtotdes.Caption = labtotdes & " * " & t_pre.Text
   Else
      If Xsigno = 1 Then
         labtotdes.Caption = labtotdes & " + " & t_pre.Text
      Else
         If Xsigno = 3 Then
            labtotdes.Caption = labtotdes & " - " & t_pre.Text
         Else
            If Xsigno = 4 Then
               labtotdes.Caption = labtotdes & " / " & t_pre.Text
            End If
'            labtotdes.Caption = labtotdes & " " & t_pre.Text
         End If
      End If
   End If
   t_pre.Text = 0
End If
Xsigno = 10

End Sub

Private Sub b_mas_Click()
labs.Caption = "+"
Xsigno = 1
If t_pre.Text <> "" Then
   If t_pre.Text = 0 Then
   Else
      Xvalo1 = CDbl(t_pre.Text)
      If labtot.Caption = "" Then
         labtot.Caption = t_pre.Text
      Else
         Xvalo2 = CDbl(labtot.Caption)
         labtot.Caption = Xvalo1 + Xvalo2
      End If
      labtot.Caption = Format(labtot.Caption, "Standard")
      labtotdes.Caption = labtotdes & " " & t_pre.Text
   End If
End If

End Sub

Private Sub b_menos_Click()
labs.Caption = "-"
Xsigno = 3
If t_pre.Text <> "" Then
   If t_pre.Text = 0 Then
   Else
      Xvalo1 = CDbl(t_pre.Text)
      If labtot.Caption = "" Then
         labtot.Caption = t_pre.Text
      Else
         Xvalo2 = CDbl(labtot.Caption)
         labtot.Caption = Xvalo1 - Xvalo2
      End If
      labtot.Caption = Format(labtot.Caption, "Standard")
      labtotdes.Caption = labtotdes & " " & t_pre.Text
   End If
End If

End Sub

Private Sub b_por_Click()
labs.Caption = "*"
Xsigno = 2
If t_pre.Text <> "" Then
   If t_pre.Text = 0 Then
   Else
      Xvalo1 = CDbl(t_pre.Text)
      If labtot.Caption = "" Then
         labtot.Caption = t_pre.Text
      Else
         Xvalo2 = CDbl(labtot.Caption)
         labtot.Caption = Xvalo1 * Xvalo2
      End If
      labtot.Caption = Format(labtot.Caption, "Standard")
      labtotdes.Caption = labtotdes & " " & t_pre.Text
   End If
End If

      
End Sub

Private Sub Command1_Click()
If labtot.Caption <> "" Then
   frm_factconve.t_imp.Text = Format(labtot.Caption, "Standard")
Else
'   frm_factconve.t_imp.Text = 0
End If
If labtotdes.Caption = "" Then
   frm_factconve.labcalc.Caption = ""
Else
   frm_factconve.labcalc.Caption = labtotdes.Caption
End If
Unload Me

End Sub

Private Sub Command2_Click()
t_pre.Text = 0
labtot.Caption = 0
labtotdes.Caption = ""
Xsigno = 0

Unload Me

End Sub

Private Sub Form_Load()

'If frm_factconve.t_preu.Text <> "" Then


'End If


End Sub



Private Sub t_pre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_por.SetFocus
End If

End Sub

Private Sub t_pre_LostFocus()
If t_pre.Text = "" Then
   t_pre.Text = 0
End If
t_pre.Text = Format(t_pre.Text, "Standard")


End Sub
