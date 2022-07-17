VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_cancella 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   0  'None
   Caption         =   "Cancelar llamado"
   ClientHeight    =   3435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7185
   Icon            =   "frm_cancella.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.ComboBox text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "frm_cancella.frx":0442
      Left            =   2160
      List            =   "frm_cancella.frx":0479
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   1920
      Width           =   4695
   End
   Begin VB.CommandButton bcance 
      Caption         =   "CANCELAR"
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
      Left            =   4800
      TabIndex        =   7
      Top             =   2760
      Width           =   1455
   End
   Begin VB.CommandButton bacep 
      Caption         =   "ACEPTAR"
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
      Left            =   1200
      TabIndex        =   6
      Top             =   2760
      Width           =   1455
   End
   Begin MSMask.MaskEdBox mh 
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   1320
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "HH:mm"
      Mask            =   "##:##"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mf 
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   720
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C00000&
      Caption         =   "MOTIVO:"
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
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1920
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C00000&
      Caption         =   "HORA:"
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
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C00000&
      Caption         =   "FECHA:"
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
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "CANCELAR LLAMADO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   120
      Picture         =   "frm_cancella.frx":059B
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   975
   End
End
Attribute VB_Name = "frm_cancella"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bacep_Click()

If mf.Text <> "__/__/____" Then
   If mh.Text <> "__:__" Then
      If Text1.ListIndex >= 0 Then
         If frm_largador.txt_nro.Text <> "" Then
            Data1.RecordSource = "select * from llamado where nrolla =" & frm_largador.txt_nro.Text
            Data1.Refresh
            If Data1.Recordset.RecordCount > 0 Then
               Data1.Recordset.Edit
               Data1.Recordset("fec_cance") = Format(mf.Text, "dd/mm/yyyy")
               Data1.Recordset("hor_cance") = Format(mh.Text, "HH:mm")
               Data1.Recordset("cancela") = 1
               Data1.Recordset("pend") = 2
               Data1.Recordset("motcance") = Text1.Text
               Data1.Recordset("user_cance") = WElusuario
               Data1.Recordset("editando") = 1
               Data1.Recordset.Update
            Else
               MsgBox "No se encuentra el llamado"
            End If
         Else
            MsgBox "No se encuentra llamado, verifique"
         End If
      Else
         MsgBox "No seleccionó motivo. No se grabará!"
      End If
   Else
      MsgBox "No ingresó hora, no se grabara", vbCritical, "Mensaje"
   End If
Else
   MsgBox "No ingresó FECHA, no se grabara", vbCritical, "Mensaje"
End If
Unload Me

      
End Sub

Private Sub bcance_Click()
Unload Me

End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"

mh.Text = Format(Time, "HH:mm")
mf.Text = Date
mh.Enabled = False
mf.Enabled = False

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub mf_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mh.SetFocus
End If

End Sub

Private Sub mh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Text1.SetFocus
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   bacep.SetFocus
End If

End Sub
