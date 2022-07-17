VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_aran 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aranceles de servicios por convenio"
   ClientHeight    =   8535
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9585
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_aran.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8535
   ScaleWidth      =   9585
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_graba 
      Caption         =   "data_graba"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4680
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data data_estudios 
      Caption         =   "data_estudios"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   6120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8160
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4680
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3480
      TabIndex        =   8
      Top             =   4920
      Width           =   3735
   End
   Begin MSDBGrid.DBGrid DBGrid2 
      Bindings        =   "frm_aran.frx":0442
      Height          =   2895
      Left            =   120
      OleObjectBlob   =   "frm_aran.frx":0456
      TabIndex        =   6
      Top             =   5280
      Width           =   9375
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4440
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8760
      Picture         =   "frm_aran.frx":0E35
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Salir"
      Top             =   4440
      Width           =   735
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_aran.frx":13BF
      Height          =   3855
      Left            =   120
      OleObjectBlob   =   "frm_aran.frx":13D3
      TabIndex        =   3
      Top             =   600
      Width           =   9375
   End
   Begin VB.TextBox t_busca 
      Height          =   375
      Left            =   5400
      TabIndex        =   2
      Top             =   120
      Width           =   4095
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      ItemData        =   "frm_aran.frx":2462
      Left            =   2520
      List            =   "frm_aran.frx":246C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Caption         =   "Doble click para agregar aranceles al convenio."
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   9
      Top             =   8160
      Width           =   5295
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      Caption         =   "Buscar convenio por descripción:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   5040
      Width           =   3375
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "%0 = Costo total sin descuento en precio de servicio."
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4440
      Width           =   5655
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "Buscar convenio por:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   7440
      Picture         =   "frm_aran.frx":2485
      Stretch         =   -1  'True
      Top             =   4920
      Width           =   1335
   End
End
Attribute VB_Name = "frm_aran"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub DBGrid2_DblClick()
Dim Xquehaceara, Xquehacedos As String
Xquehaceara = MsgBox("Desea agregar aranceles al convenio " & Data2.Recordset("cnv_codigo") & ":?", vbInformation + vbYesNo, "ARANCELES")
data_graba.RecordSource = "arancel"
data_graba.Refresh

If Xquehaceara = vbYes Then
   Data1.RecordSource = "Select * from arancel where ara_cnvcod ='" & Data2.Recordset("cnv_codigo") & "'"
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      Xquehacedos = MsgBox("Ya existen aranceles para este convenio, DESEA AGREGAR NUEVOS?", vbInformation + vbYesNo, "ARANCELES")
      If Xquehacedos = vbYes Then
        data_estudios.RecordSource = "Select * from estudios where flia in (1,6,9,10,13,14,11,3) order by codest"
        data_estudios.Refresh
        If data_estudios.Recordset.RecordCount > 0 Then
           data_estudios.Recordset.MoveFirst
        End If
        Do While Not data_estudios.Recordset.EOF
           Data1.RecordSource = "Select * from arancel where ara_cnvcod ='" & Data2.Recordset("cnv_codigo") & "' and ara_famnro =" & data_estudios.Recordset("codest")
           Data1.Refresh
           If Data1.Recordset.RecordCount > 0 Then
           Else
              data_graba.Recordset.AddNew
              data_graba.Recordset("ara_famnro") = data_estudios.Recordset("codest")
              data_graba.Recordset("ara_famnom") = Mid(data_estudios.Recordset("descrip"), 1, 40)
              data_graba.Recordset("ara_cnvcod") = Data2.Recordset("cnv_codigo")
              data_graba.Recordset("ara_cnvdes") = Mid(Data2.Recordset("cnv_desc"), 1, 20)
              data_graba.Recordset("ara_precio") = 0
              data_graba.Recordset("ara_porcen") = 0
              data_graba.Recordset.Update
           End If
           data_estudios.Recordset.MoveNext
        Loop
        DoEvents
        MsgBox "Proceso terminado, verifique en la lista arriba", vbInformation, "ARANCELES"
      End If
'      DBGrid2.SetFocus
   Else
      data_estudios.RecordSource = "Select * from estudios where flia in (1,6,9,10,13,14,11) order by codest"
      data_estudios.Refresh
      If data_estudios.Recordset.RecordCount > 0 Then
         data_estudios.Recordset.MoveFirst
      End If
      Do While Not data_estudios.Recordset.EOF
         data_graba.Recordset.AddNew
         data_graba.Recordset("ara_famnro") = data_estudios.Recordset("codest")
         data_graba.Recordset("ara_famnom") = Mid(data_estudios.Recordset("descrip"), 1, 40)
         data_graba.Recordset("ara_cnvcod") = Data2.Recordset("cnv_codigo")
         data_graba.Recordset("ara_cnvdes") = Mid(Data2.Recordset("cnv_desc"), 1, 20)
         data_graba.Recordset("ara_precio") = 0
         data_graba.Recordset("ara_porcen") = 0
         data_graba.Recordset.Update
         data_estudios.Recordset.MoveNext
      Loop
      MsgBox "Proceso terminado, verifique en la lista arriba", vbInformation, "ARANCELES"
      Combo1.SetFocus
   End If
   
End If


End Sub

Private Sub Form_Load()
'Data1.DatabaseName = App.Path & "\sapp.mdb"
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
'Data2.DatabaseName = App.Path & "\sapp.mdb"
Data2.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_estudios.DatabaseName = App.Path & "\sapp.mdb"
data_estudios.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_estudios.RecordSource = "estudios"
data_estudios.Refresh
'data_graba.DatabaseName = App.Path & "\sapp.mdb"
data_graba.Connect = "odbc;dsn=" & Xconexrmt & ";"

End Sub

Private Sub Form_Resize()
With Image1
     .Top = 0
     .Left = 0
     .Width = Me.Width
     .Height = Me.Height
End With

End Sub

Private Sub t_busca_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Combo1.ListIndex = 0 Then
      Data1.RecordSource = "Select * from arancel where ara_cnvcod ='" & t_busca.Text & "' order by ara_famnro"
      Data1.Refresh
   Else
      Data1.RecordSource = "Select top 200 * from arancel where ara_cnvdes >='" & t_busca.Text & "' order by ara_cnvdes, ara_famnro"
      Data1.Refresh
   End If
   DBGrid1.SetFocus
End If

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Data2.RecordSource = "Select * from convenio where cnv_desc >='" & Text1.Text & "' order by cnv_desc"
   Data2.Refresh
   DBGrid2.SetFocus
End If

End Sub
