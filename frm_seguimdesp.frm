VERSION 5.00
Begin VB.Form frm_seguimdesp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Datos para el seguimiento"
   ClientHeight    =   3240
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6630
   Icon            =   "frm_seguimdesp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   6375
      Begin VB.OptionButton Option3 
         BackColor       =   &H00800000&
         Caption         =   "CON FACTORES DE RIESGO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   3240
         TabIndex        =   5
         Top             =   240
         Visible         =   0   'False
         Width           =   2895
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   6375
      Begin VB.OptionButton Option2 
         BackColor       =   &H00800000&
         Caption         =   "NO SINTOMÁTICO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   3240
         TabIndex        =   3
         Top             =   360
         Value           =   -1  'True
         Width           =   2895
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00800000&
         Caption         =   "SINTOMÁTICO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2895
      End
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2400
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2760
      Picture         =   "frm_seguimdesp.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Graba los datos y cierra la ventana"
      Top             =   2520
      Width           =   615
   End
End
Attribute VB_Name = "frm_seguimdesp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Eldia As Integer
Dim LaFechaProx As Date
On Error GoTo Nosepuede

If Data1.Recordset.RecordCount > 0 Then
   If Option1.Value = True Then
      Data1.Recordset.Edit
      LaFechaProx = Date + 3
      Eldia = Weekday(LaFechaProx, vbUseSystemDayOfWeek)
      If Eldia = 7 Then
         Data1.Recordset("prox_control") = Date + 4
      Else
         Data1.Recordset("prox_control") = Date + 3
      End If
      Data1.Recordset("op_seguim") = 1
      Data1.Recordset("editando") = 1
      Data1.Recordset.Update
   Else
      If Option2.Value = True Then
         If Option3.Value = True Then
            Data1.Recordset.Edit
            LaFechaProx = Date + 3
            Eldia = Weekday(LaFechaProx, vbUseSystemDayOfWeek)
            If Eldia = 7 Then
               Data1.Recordset("prox_control") = Date + 4
            Else
               Data1.Recordset("prox_control") = Date + 3
            End If
            Data1.Recordset("op_seguim") = 2
            Data1.Recordset("editando") = 1
            Data1.Recordset.Update
         Else
            Data1.Recordset.Edit
            LaFechaProx = Date + 4
            Eldia = Weekday(LaFechaProx, vbUseSystemDayOfWeek)
            If Eldia = 7 Then
               Data1.Recordset("prox_control") = Date + 5
            Else
               Data1.Recordset("prox_control") = Date + 4
            End If
            Data1.Recordset("op_seguim") = 3
            Data1.Recordset("editando") = 1
            Data1.Recordset.Update
         End If
      Else
         If IsNull(Data1.Recordset("prox_control")) = False Then
            Data1.Recordset.Edit
            Data1.Recordset("prox_control") = Null
            Data1.Recordset("editando") = 1
            Data1.Recordset.Update
         End If
         If IsNull(Data1.Recordset("prox_control")) = False Then
            Data1.Recordset.Edit
            Data1.Recordset("op_seguim") = Null
            Data1.Recordset("editando") = 1
            Data1.Recordset.Update
         End If
      End If
   End If
Else
   MsgBox "No se pueden grabar estos datos, verifique si ya grabó el llamado y vuelva a intentar.", vbCritical
   
End If
Unload Me

Exit Sub

Nosepuede:
        If Err.Number = 3155 Then
           MsgBox "Verifique datos en seguimiento covid", vbInformation
        Else
           MsgBox "Verifique datos, ya figuran en seguimiento covid", vbInformation
        End If
        Unload Me
        
        
End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.RecordSource = "select * from llamado where nrolla =" & frm_largador.txt_nro.Text
Data1.Refresh


End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
   Option3.Visible = False
End If

End Sub

Private Sub Option1_DblClick()
If Option1.Value = True Then
   Option1.Value = False
End If

End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
   Option3.Visible = True
Else
   Option3.Visible = False
End If

End Sub

Private Sub Option2_DblClick()
If Option2.Value = True Then
   Option2.Value = False
End If

End Sub

Private Sub Option3_DblClick()
If Option3.Value = True Then
   Option3.Value = False
End If

End Sub
