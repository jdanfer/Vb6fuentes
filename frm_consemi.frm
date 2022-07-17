VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_consemi 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultar Emisiones"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8880
   Icon            =   "frm_consemi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   8880
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox t_mat 
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
      Left            =   3000
      TabIndex        =   9
      Top             =   1440
      Width           =   2055
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
      Left            =   240
      Picture         =   "frm_consemi.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   6120
      Width           =   495
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_consemi.frx":09CC
      Height          =   4215
      Left            =   240
      OleObjectBlob   =   "frm_consemi.frx":09E0
      TabIndex        =   6
      Top             =   1920
      Width           =   8535
   End
   Begin VB.TextBox txt_cob 
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
      Left            =   3000
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Consultar..."
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
      Left            =   5880
      TabIndex        =   3
      Top             =   1440
      Width           =   1815
   End
   Begin VB.TextBox txt_a 
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
      Left            =   3600
      MaxLength       =   4
      TabIndex        =   2
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox txt_m 
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
      Left            =   3000
      MaxLength       =   2
      TabIndex        =   1
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Buscar por matrícula:"
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
      TabIndex        =   8
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "NUMERO DE COBRADOR:"
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
      TabIndex        =   4
      Top             =   840
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "MES/AÑO A CONSULTAR:"
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
      Top             =   240
      Width           =   2775
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   3120
      Picture         =   "frm_consemi.frx":276F
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   1335
   End
End
Attribute VB_Name = "frm_consemi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Nombre As String

frm_consemi.MousePointer = 11
Nombre = "emi"
If txt_m.Text > 9 Then
   Nombre = Nombre + Trim(txt_m.Text) + Mid(Trim(txt_a.Text), 3, 2)
Else
   Nombre = Nombre + "0" + Trim(txt_m.Text) + Mid(Trim(txt_a.Text), 3, 2)
End If
If txt_m.Text <> "" Then
   If txt_a.Text <> "" Then
      If txt_cob.Text <> "" Then
         If t_mat.Text <> "" Then
            Data1.RecordSource = "Select * from " & Nombre & " where cliente >=" & t_mat.Text & " order by cliente"
            Data1.Refresh
         Else
            Data1.RecordSource = "Select * from " & Nombre & " where nro_cobr =" & txt_cob.Text & " order by apellidos"
            Data1.Refresh
         End If
      Else
         If t_mat.Text <> "" Then
            Data1.RecordSource = "Select * from " & Nombre & " where cliente >=" & t_mat.Text & " order by cliente"
            Data1.Refresh
         Else
            Data1.RecordSource = "Select * from " & Nombre & " order by apellidos"
            Data1.Refresh
         End If
      End If
   End If
End If
frm_consemi.MousePointer = 0

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
txt_m.Text = Month(Date)
txt_a.Text = Year(Date)
Data1.DatabaseName = ""
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With


End Sub
