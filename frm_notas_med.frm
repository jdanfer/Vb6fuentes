VERSION 5.00
Begin VB.Form frm_notas_med 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notas"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7590
   Icon            =   "frm_notas_med.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   7590
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   6960
      Picture         =   "frm_notas_med.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   4200
      Width           =   375
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1200
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.TextBox t_notareg 
      Enabled         =   0   'False
      Height          =   2775
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   1800
      Width           =   6735
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6960
      Picture         =   "frm_notas_med.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   960
      Width           =   375
   End
   Begin VB.TextBox t_nota 
      Height          =   975
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   6735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Caption         =   "Notas registradas:"
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
      Left            =   240
      TabIndex        =   3
      Top             =   1560
      Width           =   6735
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Agregar nueva nota:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   6735
   End
End
Attribute VB_Name = "frm_notas_med"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Trim(t_nota.Text) <> "" Then
   t_notareg.Text = ""
   Data1.Recordset.AddNew
   Data1.Recordset("fecha") = Date
   Data1.Recordset("usuario") = WElusuario
   Data1.Recordset("matricula") = Val(frmabm.txt_mat.Caption)
   Data1.Recordset("nota") = t_nota.Text
   Data1.Recordset.Update
   Data1.RecordSource = "select * from notas_med where matricula =" & Val(frmabm.txt_mat.Caption) & " order by fecha"
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         If Trim(t_notareg.Text) <> "" Then
            t_notareg.Text = t_notareg.Text & vbCrLf & "Fecha:" & Format(Data1.Recordset("fecha"), "dd/mm/yyyy") & "--" & Data1.Recordset("nota") & "--->" & Data1.Recordset("usuario")
         Else
            t_notareg.Text = "Fecha:" & Data1.Recordset("fecha") & "--" & Data1.Recordset("nota") & "--->" & Data1.Recordset("usuario")
         End If
         Data1.Recordset.MoveNext
      Loop
   End If
   MsgBox "Registro grabado correctamente."
   t_nota.Text = ""
   
End If

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
If Trim(frmabm.txt_mat.Caption) <> "" Then
   Data1.RecordSource = "select * from notas_med where matricula =" & Val(frmabm.txt_mat.Caption) & " order by fecha"
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         If Trim(t_notareg.Text) <> "" Then
            t_notareg.Text = t_notareg.Text & vbCrLf & "Fecha:" & Format(Data1.Recordset("fecha"), "dd/mm/yyyy") & "--" & Data1.Recordset("nota") & "--->" & Data1.Recordset("usuario")
         Else
            t_notareg.Text = "Fecha:" & Format(Data1.Recordset("fecha"), "dd/mm/yyyy") & "--" & Data1.Recordset("nota") & "--->" & Data1.Recordset("usuario")
         End If
         Data1.Recordset.MoveNext
      Loop
   End If
Else
   MsgBox "No hay cliente seleccionado.", vbInformation
   Data1.RecordSource = "select * from notas_med where matricula =" & 0
   Data1.Refresh
   Command1.Enabled = False
End If

End Sub
