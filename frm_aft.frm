VERSION 5.00
Begin VB.Form frm_aft 
   BackColor       =   &H00C0C000&
   BorderStyle     =   0  'None
   Caption         =   "AFT"
   ClientHeight    =   1950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   ScaleHeight     =   1950
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Visible         =   0   'False
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFFF&
      Height          =   495
      Left            =   5160
      Picture         =   "frm_aft.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cerrar ventana"
      Top             =   1440
      Width           =   615
   End
   Begin VB.TextBox t_cod 
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
      Left            =   2640
      MaxLength       =   8
      TabIndex        =   2
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackColor       =   &H0000FF00&
      Caption         =   "Ultimo entregado:"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   4
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H0080FFFF&
      Caption         =   "Ingrese CODIGO:"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H0080FFFF&
      Caption         =   "AUTORIZACION FINAL DE TRASLADO"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   4080
      Picture         =   "frm_aft.frx":058A
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   855
   End
End
Attribute VB_Name = "frm_aft"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command3_Click()
On Error GoTo AlGaAFT

If t_cod.Text <> "" Then
   XaftC = t_cod.Text
   Data1.RecordSource = "Select * from llamado where nrolla =" & frm_largador.txt_nro.Text
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      If IsNull(Data1.Recordset("aft")) = False Then
         If Data1.Recordset("aft") <> t_cod.Text Then
            Data1.Recordset.Edit
            Data1.Recordset("aft") = t_cod.Text
            Data1.Recordset("editando") = 1
            Data1.Recordset.Update
         End If
      Else
         Data1.Recordset.Edit
         Data1.Recordset("aft") = t_cod.Text
         Data1.Recordset("editando") = 1
         Data1.Recordset.Update
      End If
      frm_largador.Label40.Caption = "AFT:" & t_cod.Text
   Else
      MsgBox "Aún no se ha GUARDADO el llamado. Después de guardar puede ingresar el código", vbInformation
      frm_largador.Label40.Caption = ""
   End If
Else
   frm_largador.Label40.Caption = ""
   Data1.RecordSource = "Select * from llamado where nrolla =" & frm_largador.txt_nro.Text
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      If IsNull(Data1.Recordset("aft")) = False Then
         Data1.Recordset.Edit
         Data1.Recordset("aft") = Null
         Data1.Recordset("editando") = 1
         Data1.Recordset.Update
      End If
   End If
End If

Unload Me

Exit Sub

AlGaAFT:
      If Err.Number = 444 Then
         MsgBox "No se pudo grabar, comunique a informática ALGaAFT ERR:" & Err.Description
      Else
         MsgBox "Error al grabar, comunique a informática ALGaAFT ERR:" & Err.Number
      End If


End Sub

Private Sub Form_Load()
'Label3.Caption = ""
'Data1.DatabaseName = App.Path & "\sapp.mdb"
On Error GoTo AliniAFT
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
'Data1.RecordSource = "Select * from resplla where descol = not null order by descol DESC"
'Data1.Refresh
'If Data1.Recordset.RecordCount > 0 Then
'   Data1.Recordset.MoveFirst
'   Do While Data1.Recordset("descol") = 3
'      Data1.Recordset.MoveNext
'   Loop
   Label4.Caption = "Consultar"
'End If

Data1.RecordSource = "Select * from llamado where nrolla =" & frm_largador.txt_nro.Text
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   If IsNull(Data1.Recordset("aft")) = False Then
      t_cod.Text = Data1.Recordset("aft")
   Else
      t_cod.Text = ""
   End If
Else
   MsgBox "No ha GRABADO el llamado. Guarde primero y después registre el código.", vbInformation
   
End If

Exit Sub

AliniAFT:
      If Err.Number = 444 Then
         MsgBox "No se pudo grabar, comunique a informática INIAFT ERR:" & Err.Description
      Else
         MsgBox "Error al grabar, comunique a informática INIAFT ERR:" & Err.Number
      End If



End Sub

Private Sub Form_Resize()
With Image1
    .Top = 0
    .Left = 0
    .Width = Me.Width
    .Height = Me.Height
End With

End Sub

Private Sub t_cod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command3.SetFocus
End If

End Sub
