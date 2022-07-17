VERSION 5.00
Begin VB.Form frm_ops911 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos requeridos para SAME"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6030
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
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
      Left            =   5280
      Picture         =   "frm_ops911.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Salir"
      Top             =   4080
      Width           =   495
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
      Left            =   120
      Picture         =   "frm_ops911.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Procesar"
      Top             =   4080
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Datos para SAME (Cantidades)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5655
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frm_ops911.frx":0B14
         Left            =   2760
         List            =   "frm_ops911.frx":0B21
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   3480
         Width           =   2175
      End
      Begin VB.TextBox txt_ctras 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   12
         Top             =   3000
         Width           =   975
      End
      Begin VB.TextBox txt_cmuert 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   11
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox txt_cinter 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   10
         Top             =   2040
         Width           =   975
      End
      Begin VB.TextBox txt_calta 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   9
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txt_cher 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2760
         TabIndex        =   8
         Top             =   1080
         Width           =   975
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frm_ops911.frx":0B3E
         Left            =   2760
         List            =   "frm_ops911.frx":0B51
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C00000&
         Caption         =   "Area:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   3480
         Width           =   2655
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C00000&
         Caption         =   "Solicitado por:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
         Caption         =   "Muertos en traslado:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Muertos en escena:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "Internación:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Altas en zona:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Cantidad de heridos:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   2655
      End
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   2040
      Picture         =   "frm_ops911.frx":0B7C
      Stretch         =   -1  'True
      Top             =   4200
      Width           =   2415
   End
End
Attribute VB_Name = "frm_ops911"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Click()
If Combo1.ListIndex = 4 Then
   txt_cher.Enabled = False
   txt_calta.Enabled = False
   txt_cinter.Enabled = False
   txt_cmuert.Enabled = False
   txt_ctras.Enabled = False
Else
   txt_cher.Enabled = True
   txt_calta.Enabled = True
   txt_cinter.Enabled = True
   txt_cmuert.Enabled = True
   txt_ctras.Enabled = True
End If

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_cher.SetFocus
End If

End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command1.SetFocus
End If

End Sub

Private Sub Command1_Click()

If txt_ctras.Text = "" Then
   frm_largador.Label48.Caption = 0
Else
   frm_largador.Label48.Caption = txt_ctras.Text
End If

If Combo1.ListIndex = 4 Then
   frm_largador.Label45.Caption = InputBox("Ingrese número de llamado principal", "Llamado dependiente de:")
   frm_largador.Label46.Caption = 8
Else
   frm_largador.Label46.Caption = Combo1.ListIndex
End If
If txt_cher.Text = "" Then
   frm_largador.Label41.Caption = 0
Else
   frm_largador.Label41.Caption = txt_cher.Text
End If
If txt_calta.Text = "" Then
   frm_largador.Label42.Caption = 0
Else
   frm_largador.Label42.Caption = txt_calta.Text
End If
If txt_cinter.Text = "" Then
   frm_largador.Label43.Caption = 0
Else
   frm_largador.Label43.Caption = txt_cinter.Text
End If
If txt_cmuert.Text = "" Then
   frm_largador.Label44.Caption = 0
Else
   frm_largador.Label44.Caption = txt_cmuert.Text
End If
If Combo2.Text <> "" Then
   frm_largador.txt_quien.Text = Combo2.Text
Else
   frm_largador.txt_quien.Text = ""
End If
Unload Me

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
If frm_largador.Label46.Caption = 8 Then
   Combo1.ListIndex = 4
Else
   If frm_largador.Label46.Caption = 0 Or _
      frm_largador.Label46.Caption = 2 Or _
      frm_largador.Label46.Caption = 3 Or _
      frm_largador.Label46.Caption = 4 Then
      Combo1.ListIndex = frm_largador.Label46.Caption
   Else
      Combo1.ListIndex = -1
   End If
End If
If frm_largador.txt_quien.Text <> "" Then
   If frm_largador.txt_quien.Text = "Ruta" Then
      Combo2.ListIndex = 0
   End If
   If frm_largador.txt_quien.Text = "Urbano" Then
      Combo2.ListIndex = 1
   End If
   If frm_largador.txt_quien.Text = "Suburbano" Then
      Combo2.ListIndex = 2
   End If
End If
txt_cher.Text = frm_largador.Label41.Caption
txt_calta.Text = frm_largador.Label42.Caption
txt_cinter.Text = frm_largador.Label43.Caption
txt_cmuert.Text = frm_largador.Label44.Caption
If frm_largador.Label48.Caption = "" Then
   txt_ctras.Text = 0
Else
   txt_ctras.Text = frm_largador.Label48.Caption
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

Private Sub txt_calta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_cinter.SetFocus
End If

End Sub

Private Sub txt_cher_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_calta.SetFocus
End If

End Sub

Private Sub txt_cinter_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_cmuert.SetFocus
End If

End Sub

Private Sub txt_cmuert_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_ctras.SetFocus
End If

End Sub

Private Sub txt_ctras_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo2.SetFocus
End If

End Sub

Private Sub txt_quien_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command1.SetFocus
End If

End Sub
