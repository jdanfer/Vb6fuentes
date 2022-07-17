VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_archenf 
   BackColor       =   &H00FF8080&
   Caption         =   "Subir archivos"
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8745
   Icon            =   "frm_archenf.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   8745
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   3600
      Top             =   1800
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3120
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3120
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   25000
      Left            =   1200
      Top             =   600
   End
   Begin VB.TextBox t_fec 
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.TextBox t_ced 
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   720
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Abrir"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salvar"
      Height          =   615
      Left            =   6840
      Picture         =   "frm_archenf.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      Height          =   615
      Left            =   6840
      Picture         =   "frm_archenf.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox t_id 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox t_nom 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   4215
   End
   Begin MSComDlg.CommonDialog cmm1 
      Left            =   3840
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Fecha:"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Cédula"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Archivo:"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "frm_archenf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pdfpath, pdfpath1 As String
Public pdffile As ADODB.Stream

Private Sub Command1_Click()
'With cmm1
'     .FileName = ""
'     .Filter = "PDF (*.pdf;) | *.pdf;"
'     .ShowOpen
'     If Len(.FileName) <> 0 Then
'        pdfpath = .FileName
'        pdfpath1 = .FileTitle
'        t_nom.Text = .FileTitle
'     End If
'     t_id.Text = 10
'End With

End Sub

Private Sub Command2_Click()
'Command2.Enabled = False
'Adodc1.RecordSource = "archs"
'Adodc1.Refresh

'If pdfpath <> "" Then
'   Data1.Recordset.Edit
'   Data1.Recordset("nro_informat") = Data1.Recordset("nro_informat") + 1
'   Data1.Recordset.Update
'   Adodc1.Recordset.AddNew
'   Set pdffile = New ADODB.Stream
'   pdffile.Type = adTypeBinary
'   pdffile.Open
'   pdffile.LoadFromFile pdfpath
'   Adodc1.Recordset.Fields("arch") = pdffile.Read
'   Adodc1.Recordset("id") = Data1.Recordset("nro_informat")
'   Adodc1.Recordset("nombre") = pdfpath1
'   Adodc1.Recordset("fecha") = Date
'   Adodc1.Recordset("cedula") = t_ced.Text
'   Adodc1.Recordset("fecha") = CDate(t_fec.Text)
'   Adodc1.Recordset.Update
 '  pdffile.Close
'   Set pdffile = Nothing
'   Kill "d:\laboratorios\" & t_nom.Text & ".pdf"
'   MsgBox "Guardado"
'Else
'   MsgBox "No hay archivo"
   
'End If
'Command2.Enabled = True

End Sub

Private Sub Command3_Click()
t_id.Text = Data1.Recordset("nroid")
't_id.Text = 3000009

'If Data2.Recordset("numero") = 1 Then
Data3.RecordSource = "Select * from arch_orden where idsrv =" & Data2.Recordset("numero")
'Else
'   Data3.RecordSource = "Select * from archs where id =" & t_id.Text
'End If
Data3.Refresh
If Data3.Recordset.RecordCount > 0 Then
   Set pdffile = New ADODB.Stream
   pdffile.Type = adTypeBinary
   pdffile.Open
   If IsNull(Data3.Recordset("archivo")) = False Then
      pdffile.Write Data3.Recordset("archivo").Value
      Dim pdfname As String
      Timer1.Enabled = True
      pdfname = "temporal"
      pdffile.SaveToFile "" & App.Path & "\laboratorio\" & pdfname & ".pdf", adSaveCreateOverWrite
      pdffile.Close
      Set pdffile = Nothing
'AcroRd32
'      Shell Data1.Recordset("desc") & " " & App.Path & "\laboratorio\temporal" & ".pdf", vbMaximizedFocus
'      Shell "c:\Program Files (x86)\Adobe\Reader 9.0\Reader\AcroRd32.exe " & App.Path & "\laboratorio\" & pdfname & ".pdf", vbMaximizedFocus

   Else
      MsgBox "no hay archivo"
   End If
Else
   MsgBox "No existe registro"
End If

End


End Sub

Private Sub Form_Load()

Data3.Connect = "ODBC;DSN=sappespecial;"

Data1.DatabaseName = App.Path & "\desc.mdb"
Data1.RecordSource = "desc"
Data1.Refresh

Data2.DatabaseName = App.Path & "\abrir.mdb"
Data2.RecordSource = "abrir"
Data2.Refresh

Command3_Click

End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False

End Sub
