VERSION 5.00
Begin VB.Form frm_obsenfadm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Observaciones (Administrador)"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8865
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_obsenfadm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   8865
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1440
      Visible         =   0   'False
      Width           =   1935
   End
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
      Height          =   375
      Left            =   8160
      Picture         =   "frm_obsenfadm.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   2760
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
      Height          =   375
      Left            =   240
      Picture         =   "frm_obsenfadm.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Guardar datos"
      Top             =   2760
      Width           =   495
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
      Height          =   1815
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   960
      Width           =   8415
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Observaciones:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   2895
   End
   Begin VB.Label Label3 
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
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   120
      Width           =   5415
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Acto Nro:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   4200
      Picture         =   "frm_obsenfadm.frx":109E
      Stretch         =   -1  'True
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "frm_obsenfadm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Data1.Recordset.RecordCount > 0 Then
   If Text1.Text <> "" Then
      If IsNull(Data1.Recordset("srv_obsadm")) = False Then
         If Data1.Recordset("srv_obsadm") <> Text1.Text Then
            Data1.Recordset.Edit
            Data1.Recordset("srv_obsadm") = Text1.Text
            Data1.Recordset.Update
            MsgBox "Registro grabado correctamente"
         End If
      Else
         Data1.Recordset.Edit
         Data1.Recordset("srv_obsadm") = Text1.Text
         Data1.Recordset.Update
         MsgBox "Registro grabado correctamente"
      
      End If
   Else
      If IsNull(Data1.Recordset("srv_obsadm")) = False Then
         Data1.Recordset.Edit
         Data1.Recordset("srv_obsadm") = Null
         Data1.Recordset.Update
         MsgBox "Registro grabado correctamente"
      End If
   End If
End If

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
Label2.Caption = frm_srvenferm.Label19.Caption
Label3.Caption = frm_srvenferm.t_nom.Text

Data1.Connect = "ODBC;DSN=sappespecial;"
Data1.RecordSource = "Select * from serv_enferm where id =" & Val(Label2.Caption)
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   If IsNull(Data1.Recordset("srv_obsadm")) = False Then
      Text1.Text = Data1.Recordset("srv_obsadm")
   Else
      Text1.Text = ""
   End If
Else
   Text1.Text = ""
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
