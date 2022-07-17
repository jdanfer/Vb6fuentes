VERSION 5.00
Begin VB.Form frm_repet 
   BackColor       =   &H00FF8080&
   BorderStyle     =   0  'None
   Caption         =   "Matrícula repetida"
   ClientHeight    =   4005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3000
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1200
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Text4 
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
      Left            =   6000
      TabIndex        =   8
      Top             =   840
      Width           =   375
   End
   Begin VB.TextBox Text3 
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
      Height          =   375
      Left            =   4200
      TabIndex        =   7
      Top             =   840
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Salir"
      Height          =   735
      Left            =   6360
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3120
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Procesar"
      Height          =   735
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3000
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   3
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "RECUERDE!! Que para realizar éste proceso, deberá ser el único usuario en el sistema."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   2160
      Width           =   6135
   End
   Begin VB.Label Label3 
      Caption         =   "Ingrese la cédula correspondiente:"
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
      TabIndex        =   6
      Top             =   840
      Width           =   3855
   End
   Begin VB.Label Label2 
      Caption         =   "Ingrese nuevo número de matrícula:"
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
      TabIndex        =   2
      Top             =   1680
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Ingrese número de matrícula repetida:"
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
      TabIndex        =   0
      Top             =   360
      Width           =   3855
   End
End
Attribute VB_Name = "frm_repet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Xlohace As String
Xlohace = ""
Xlohace = MsgBox("Desea realizar el proceso ?", vbCritical + vbYesNo, "SAPP")
If Xlohace = vbYes Then
   If Text1.Text = "" Then
   Else
      If Text2.Text = "" Then
      Else
         DoEvents
         Data1.RecordSource = "Select * from clientes where cl_codigo =" & Text1.Text & " and cl_cedula =" & Text3.Text & " and cl_codced =" & Text4.Text
         Data1.Refresh
         If Data1.Recordset.RecordCount > 0 Then
            Data1.Recordset.Edit
            Data1.Recordset("cl_codigo") = Text2.Text
            Data1.Recordset.Update
            Data2.Recordset.Edit
            Data2.Recordset("ultimo_soc") = Text2.Text
            Data2.Recordset.Update
            Text2.Enabled = True
            Text2.Text = Text2.Text + 1
            Text2.Enabled = False
            MsgBox "Proceso terminado"
            Text1.Text = ""
         End If
      End If
   End If
End If

      
End Sub

Private Sub Command2_Click()
End

End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\sapp.mdb"
Data2.DatabaseName = App.Path & "\parse.mdb"
Data2.RecordSource = "parsec0"
Data2.Refresh
Text2.Text = Data2.Recordset("ultimo_soc") + 1
Text2.Enabled = False

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Text3.SetFocus
End If

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Text4.SetFocus
End If

End Sub

