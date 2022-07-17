VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00FF8080&
   Caption         =   "Ingreso de archivos al servidor"
   ClientHeight    =   3000
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8655
   Icon            =   "frm_ingotros.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   8655
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox t_id 
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Guardar"
      Height          =   735
      Left            =   7200
      Picture         =   "frm_ingotros.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      Height          =   735
      Left            =   7200
      Picture         =   "frm_ingotros.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   420
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   720
      Top             =   2520
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1085
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
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=sapparch"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "sapparch"
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
   Begin MSComDlg.CommonDialog cmm1 
      Left            =   4080
      Top             =   2160
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox t_fec 
      Height          =   375
      Left            =   2040
      TabIndex        =   5
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox t_ced 
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   720
      Width           =   1695
   End
   Begin VB.TextBox t_nom 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   120
      Width           =   6255
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080FFFF&
      Caption         =   "Para cédula 3584484-4 se debe ingresar: 35844844"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Fecha:"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Cedula:"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Archivo:"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pdfpath, pdfpath1 As String
Public pdffile As ADODB.Stream

Private Sub Command1_Click()
With cmm1
     .FileName = ""
     .Filter = "PDF (*.pdf;) | *.pdf;"
     .ShowOpen
     If Len(.FileName) <> 0 Then
        pdfpath = .FileName
        pdfpath1 = .FileTitle
        t_nom.Text = .FileTitle
     End If
'     t_id.Text = 10
End With
End Sub

Private Sub Command2_Click()
'Adodc1.ConnectionString = "ODBC;DSN=laboratorio;"
Command2.Enabled = False

If pdfpath <> "" Then
   Data1.Recordset.Edit
   Data1.Recordset("nro_material") = Data1.Recordset("nro_material") + 1
   Data1.Recordset.Update
   Adodc1.Recordset.AddNew
   Set pdffile = New ADODB.Stream
   pdffile.Type = adTypeBinary
   pdffile.Open
   pdffile.LoadFromFile pdfpath
   Adodc1.Recordset.Fields("arch") = pdffile.Read
   Adodc1.Recordset("id") = Data1.Recordset("nro_material")
   Adodc1.Recordset("nombre") = pdfpath1
   Adodc1.Recordset("fecha") = Date
   Adodc1.Recordset("cedula") = t_ced.Text
   Adodc1.Recordset("fecha") = CDate(t_fec.Text)
   Adodc1.Recordset.Update
   pdffile.Close
   Set pdffile = Nothing
   Kill "c:\laboratorios\" & t_nom.Text
   MsgBox "Guardado"
Else
   MsgBox "No hay archivo"
   
End If
Command2.Enabled = True
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path & "\paramb.mdb"
Data1.RecordSource = "paramb"
Data1.Refresh

Adodc1.RecordSource = "Select * from arcotro where id =" & 3000011
Adodc1.Refresh

End Sub

Private Sub t_fec_GotFocus()
t_fec.Text = Format(Date, "dd/mm/yyyy")

End Sub
