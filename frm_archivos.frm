VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_archivos 
   BackColor       =   &H00404040&
   Caption         =   "Subir archivos"
   ClientHeight    =   2790
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8745
   Icon            =   "frm_archivos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   8745
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_covid 
      Caption         =   "data_covid"
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
      Top             =   2280
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   4800
      Top             =   1200
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
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
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   2460
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   615
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1680
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
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1920
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
      Width           =   1815
   End
   Begin VB.TextBox t_ced 
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   720
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
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Abrir"
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Guardar"
      Height          =   615
      Left            =   7200
      Picture         =   "frm_archivos.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Buscar"
      Height          =   615
      Left            =   7200
      Picture         =   "frm_archivos.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   960
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
      Width           =   5775
   End
   Begin MSComDlg.CommonDialog cmm1 
      Left            =   3840
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   4680
      TabIndex        =   10
      Top             =   840
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Fecha:"
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Cédula"
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Archivo:"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2175
   End
End
Attribute VB_Name = "frm_archivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pdfpath, pdfpath1 As String
Public pdffile As ADODB.Stream
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub Command1_Click()
With cmm1
     .FileName = ""
     .Filter = "PDF (*.pdf;) | *.pdf;"
     .ShowOpen
     If Len(.FileName) <> 0 Then
        pdfpath = .FileName
        pdfpath1 = .FileTitle
        t_nom.Text = .FileTitle
        Label4.Caption = .FileName
     End If
     t_id.Text = 10
End With

End Sub

Private Sub Command2_Click()
'Adodc1.ConnectionString = "ODBC;DSN=laboratorio;"
Command2.Enabled = False
If Trim(t_ced.Text) <> "" Then
   Adodc1.RecordSource = "select * from archs where cedula ='" & t_ced.Text & "'"
Else
   Adodc1.RecordSource = "select * from archs where cedula ='" & "34805844" & "'"
End If
Adodc1.Refresh

If Trim(t_ced.Text) <> "" Then
    If pdfpath <> "" Then
       Data1.Recordset.Edit
       Data1.Recordset("nroid") = Data1.Recordset("nroid") + 1
       Data1.Recordset.Update
       Data1.Refresh
       Adodc1.Recordset.AddNew
       Set pdffile = New ADODB.Stream
       pdffile.Type = adTypeBinary
       pdffile.Open
       pdffile.LoadFromFile pdfpath
       Adodc1.Recordset.Fields("arch") = pdffile.Read
       Adodc1.Recordset("id") = Data1.Recordset("nroid")
       Adodc1.Recordset("nombre") = pdfpath1
       Adodc1.Recordset("fecha") = Date
       Adodc1.Recordset("cedula") = t_ced.Text
       Adodc1.Recordset("fecha") = CDate(t_fec.Text)
       Adodc1.Recordset("proveedor") = "Covid"
       Adodc1.Recordset.Update
       pdffile.Close
       Set pdffile = Nothing
       Kill Trim(Label4.Caption)
       MsgBox "Guardado"
       If Len(Trim(t_ced.Text)) = 7 Then
          data_covid.RecordSource = "select * from sol_hisopos where cedula =" & Val(Trim(Mid(t_ced.Text, 1, 6))) & " and si_result is null"
       Else
          data_covid.RecordSource = "select * from sol_hisopos where cedula =" & Val(Trim(Mid(t_ced.Text, 1, 7))) & " and si_result is null"
       End If
       data_covid.Refresh
       If data_covid.Recordset.RecordCount > 0 Then
          data_covid.Recordset.Edit
          data_covid.Recordset("si_result") = 1
          data_covid.Recordset("id_result") = Data1.Recordset("nroid")
          data_covid.Recordset("fecha_fact") = CDate(t_fec.Text)
          data_covid.Recordset("mot_cierre") = "Resultado"
          data_covid.Recordset.Update
       End If
    Else
       MsgBox "No hay archivo"
       
    End If
Else
    MsgBox "Ingrese cédula"
End If
Command2.Enabled = True

End Sub

Private Sub Command3_Click()
Dim Xclaveu As String
Dim Xced As String
Dim Xclaveok As Integer
Xclaveok = 0
t_id.Text = Data1.Recordset("nroid")
't_id.Text = 3000009
Dim pdfname As String
Data1.Recordset.Edit
Data1.Recordset("base") = 1
Data1.Recordset.Update
Data1.Refresh

If Data2.Recordset("numero") = 1 Then
   Data3.RecordSource = "Select * from arcotro where id =" & t_id.Text
Else
   Data3.RecordSource = "Select * from archs where id =" & t_id.Text
End If
Data3.Refresh
If Data3.Recordset.RecordCount > 0 Then
   Set pdffile = New ADODB.Stream
   pdffile.Type = adTypeBinary
   pdffile.Open
   If IsNull(Data3.Recordset("arch")) = False Then
      pdffile.Write Data3.Recordset("arch").Value
      Timer1.Enabled = True
      pdfname = "temporal"
      pdffile.SaveToFile "" & App.Path & "\laboratorio\" & pdfname & ".pdf", adSaveCreateOverWrite
      pdffile.Close
      Set pdffile = Nothing
   
      If IsNull(Data3.Recordset("nombre")) = False Then
         If Mid(Trim(Data3.Recordset("nombre")), 1, 6) = "result" Then
            If IsNull(Data3.Recordset("cedula")) = False Then
               Xced = Data3.Recordset("cedula")
            Else
               Xced = "0"
            End If
            MsgBox "Archivo protegido, Ingrese clave del usuario para continuar", vbCritical
            Xclaveu = InputBox("Ingrese clave")
            If Trim(Xclaveu) <> "" Then
               Data4.RecordSource = "Select * from passuser where ceduser ='" & Xced & "'"
               Data4.Refresh
               If Data4.Recordset.RecordCount > 0 Then
                  If Data4.Recordset("passuser") = Xclaveu Then
                     Xclaveok = 9
                  Else
                     Xclaveok = 0
                  End If
               Else
                  Xclaveok = 0
               End If
            Else
               Xclaveok = 0
            End If
            If Xclaveok = 0 Then
               MsgBox "Error en la clave, verifique!", vbCritical
               If Dir(App.Path & "\laboratorio\temporal.pdf") <> "" Then
                   Kill App.Path & "\laboratorio\temporal.pdf"
               End If
            
            Else
               ShellExecute Me.hwnd, "open", App.Path & "\laboratorio\temporal.pdf", "", "", 4
            
            End If
         Else
            ShellExecute Me.hwnd, "open", App.Path & "\laboratorio\temporal.pdf", "", "", 4
         
         End If
      Else
         ShellExecute Me.hwnd, "open", App.Path & "\laboratorio\temporal.pdf", "", "", 4
      
      End If
   Else
      MsgBox "no hay archivo"
   End If
Else
   MsgBox "No existe registro"
End If

End


End Sub

Private Sub Form_Load()

Data3.Connect = "ODBC;DSN=sapparch;"
Data4.Connect = "ODBC;DSN=sapparch;"

Data1.DatabaseName = App.Path & "\desc.mdb"
Data1.RecordSource = "desc"
Data1.Refresh

Data2.DatabaseName = App.Path & "\abrir.mdb"
Data2.RecordSource = "abrir"
Data2.Refresh

data_covid.Connect = "odbc;dsn=sappnew;"

Adodc1.ConnectionString = "dsn=sapparch"
'Command3_Click
t_fec.Text = Format(Date, "dd/mm/yyyy")

End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False

End Sub
