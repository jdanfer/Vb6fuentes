VERSION 5.00
Begin VB.Form frm_cambiocont 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de contraseña"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   5715
   Icon            =   "frm_cambiocont.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   5715
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Guardar..."
      Height          =   495
      Left            =   2160
      Picture         =   "frm_cambiocont.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   1695
   End
   Begin VB.TextBox Text2 
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
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox Text1 
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
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   12
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Caption         =   "Repetir contraseña:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      Caption         =   "Nueva contraseña:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Usuario Actual:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   120
      Picture         =   "frm_cambiocont.frx":09CC
      Stretch         =   -1  'True
      Top             =   1560
      Width           =   1575
   End
End
Attribute VB_Name = "frm_cambiocont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Text1.Text = Text2.Text Then
   If Text1.Text = "" Then
      MsgBox "No ingresó contraseña nueva"
   Else
      Data1.Recordset.FindFirst "usuario ='" & Label2.Caption & "'"
      If Not Data1.Recordset.NoMatch Then
         If UCase(Data1.Recordset("clave")) = UCase(Text1.Text) Then
         Else
            Data1.Recordset.Edit
            Data1.Recordset("clave") = UCase(Text1.Text)
            Data1.Recordset.Update
         End If
         MsgBox "Se cambió la contraseña con EXITO!", vbInformation, "Usuarios"
      Else
         MsgBox "Usuario no encontrado"
      End If
   End If
End If
Unload Me

End Sub

Private Sub Form_Load()
Label2.Caption = WElusuario
'Data1.DatabaseName = App.Path & "\sapp.mdb"
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.RecordSource = "usuarios"
Data1.Refresh

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub
