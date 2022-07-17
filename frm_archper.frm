VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_archper 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Archivos del usuario"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11580
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_archper.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   11580
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_desc 
      Caption         =   "data_desc"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   9480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4920
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSComDlg.CommonDialog cmm1 
      Left            =   5520
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7800
      Top             =   120
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      Connect         =   "DSN=sappper"
      OLEDBString     =   "DSN=sappper"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   7200
      TabIndex        =   15
      Top             =   5400
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.CommandButton b_guarvar 
      BackColor       =   &H0080FFFF&
      Height          =   495
      Left            =   8640
      Picture         =   "frm_archper.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton b_busvar 
      BackColor       =   &H0080FFFF&
      Height          =   495
      Left            =   7800
      Picture         =   "frm_archper.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton b_guarsan 
      BackColor       =   &H0080FFFF&
      Height          =   495
      Left            =   4920
      Picture         =   "frm_archper.frx":109E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton b_bussan 
      BackColor       =   &H0080FFFF&
      Height          =   495
      Left            =   4080
      Picture         =   "frm_archper.frx":1628
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton b_guarcto 
      BackColor       =   &H0080FFFF&
      Height          =   495
      Left            =   1200
      Picture         =   "frm_archper.frx":1BB2
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4800
      Width           =   495
   End
   Begin VB.CommandButton b_buscto 
      BackColor       =   &H0080FFFF&
      Height          =   495
      Left            =   360
      Picture         =   "frm_archper.frx":213C
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4800
      Width           =   495
   End
   Begin VB.ListBox List3 
      Height          =   4110
      Left            =   7800
      TabIndex        =   4
      Top             =   720
      Width           =   3015
   End
   Begin VB.ListBox List2 
      Height          =   4110
      Left            =   4080
      TabIndex        =   3
      Top             =   720
      Width           =   3135
   End
   Begin VB.ListBox List1 
      Height          =   4110
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label Label6 
      BackColor       =   &H0080C0FF&
      Caption         =   "Varios"
      Height          =   255
      Left            =   7800
      TabIndex        =   8
      Top             =   480
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080C0FF&
      Caption         =   "Sanciones, Observaciones"
      Height          =   255
      Left            =   4080
      TabIndex        =   7
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label4 
      BackColor       =   &H0080C0FF&
      Caption         =   "Contratos"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Haga doble click sobre el nombre en la lista para ver el documento."
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   5280
      Width           =   6855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   2
      Top             =   0
      Width           =   5055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Funcionario:"
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frm_archper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pdfpath, pdfpath1 As String
Public pdffile As ADODB.Stream
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long


Private Sub b_buscto_Click()
With cmm1
     .FileName = ""
     .Filter = "PDF (*.pdf;) | *.pdf;"
     .ShowOpen
     If Len(.FileName) <> 0 Then
        pdfpath = .FileName
        pdfpath1 = .FileTitle
        Text1.Text = .FileTitle
     End If
'     t_id.Text = 10
End With

End Sub

Private Sub b_bussan_Click()
With cmm1
     .FileName = ""
     .Filter = "PDF (*.pdf;) | *.pdf;"
     .ShowOpen
     If Len(.FileName) <> 0 Then
        pdfpath = .FileName
        pdfpath1 = .FileTitle
        Text1.Text = .FileTitle
     End If
'     t_id.Text = 10
End With

End Sub

Private Sub b_busvar_Click()
With cmm1
     .FileName = ""
     .Filter = "PDF (*.pdf;) | *.pdf;"
     .ShowOpen
     If Len(.FileName) <> 0 Then
        pdfpath = .FileName
        pdfpath1 = .FileTitle
        Text1.Text = .FileTitle
     End If
'     t_id.Text = 10
End With

End Sub

Private Sub b_guarcto_Click()
b_buscto.Enabled = False
b_guarcto.Enabled = False
Dim Xidarch As Integer

If Wxelnrocedev <> 0 Then
   Adodc1.RecordSource = "Select * from archcto order by id DESC"
   Adodc1.Refresh
   If Adodc1.Recordset.RecordCount > 0 Then
      Xidarch = Adodc1.Recordset("id") + 1
   Else
      Xidarch = 1
   End If
   If pdfpath <> "" Then
      Adodc1.Recordset.AddNew
      Adodc1.Recordset("id") = Xidarch
      Adodc1.Recordset("nombredoc") = Text1.Text
      Adodc1.Recordset("cedarch") = Wxelnrocedev
      Adodc1.Recordset("fecha") = Date
      Set pdffile = New ADODB.Stream
      pdffile.Type = adTypeBinary
      pdffile.Open
      pdffile.LoadFromFile pdfpath
      Adodc1.Recordset.Fields("arch") = pdffile.Read
      Adodc1.Recordset.Update
      pdffile.Close
      Set pdffile = Nothing
'        Kill "d:\laboratorios\" & t_nom.Text & ".pdf"
      MsgBox "Guardado"
   Else
      MsgBox "No hay archivo"
   End If
Else
   MsgBox "Seleccione un documento"
End If
b_buscto.Enabled = True
b_guarcto.Enabled = True

End Sub

Private Sub b_guarsan_Click()
b_bussan.Enabled = False
b_guarsan.Enabled = False
Dim Xidarch As Integer

If Wxelnrocedev <> 0 Then
   Adodc1.RecordSource = "Select * from archsanc order by id DESC"
   Adodc1.Refresh
   If Adodc1.Recordset.RecordCount > 0 Then
      Xidarch = Adodc1.Recordset("id") + 1
   Else
      Xidarch = 1
   End If
   If pdfpath <> "" Then
      Adodc1.Recordset.AddNew
      Adodc1.Recordset("id") = Xidarch
      Adodc1.Recordset("nombredoc") = Text1.Text
      Adodc1.Recordset("cedarch") = Wxelnrocedev
      Adodc1.Recordset("fecha") = Date
      Set pdffile = New ADODB.Stream
      pdffile.Type = adTypeBinary
      pdffile.Open
      pdffile.LoadFromFile pdfpath
      Adodc1.Recordset.Fields("arch") = pdffile.Read
      Adodc1.Recordset.Update
      pdffile.Close
      Set pdffile = Nothing
'        Kill "d:\laboratorios\" & t_nom.Text & ".pdf"
      MsgBox "Guardado"
   Else
      MsgBox "No hay archivo"
   End If
Else
   MsgBox "Seleccione un documento"
End If
b_bussan.Enabled = True
b_guarsan.Enabled = True


End Sub

Private Sub b_guarvar_Click()
b_busvar.Enabled = False
b_guarvar.Enabled = False
Dim Xidarch As Integer

If Wxelnrocedev <> 0 Then
   Adodc1.RecordSource = "Select * from archotros order by id DESC"
   Adodc1.Refresh
   If Adodc1.Recordset.RecordCount > 0 Then
      Xidarch = Adodc1.Recordset("id") + 1
   Else
      Xidarch = 1
   End If
   If pdfpath <> "" Then
      Adodc1.Recordset.AddNew
      Adodc1.Recordset("id") = Xidarch
      Adodc1.Recordset("nombredoc") = Text1.Text
      Adodc1.Recordset("cedarch") = Wxelnrocedev
      Adodc1.Recordset("fecha") = Date
      Set pdffile = New ADODB.Stream
      pdffile.Type = adTypeBinary
      pdffile.Open
      pdffile.LoadFromFile pdfpath
      Adodc1.Recordset.Fields("arch") = pdffile.Read
      Adodc1.Recordset.Update
      pdffile.Close
      Set pdffile = Nothing
'        Kill "d:\laboratorios\" & t_nom.Text & ".pdf"
      MsgBox "Guardado"
   Else
      MsgBox "No hay archivo"
   End If
Else
   MsgBox "Seleccione un documento"
End If
b_busvar.Enabled = True
b_guarvar.Enabled = True

End Sub

Private Sub Form_Load()
Label2.Caption = frm_abmper.t_nom1.Text & " " & frm_abmper.t_apel1.Text
Data1.Connect = "ODBC;DSN=sappper;"
Data1.RecordSource = "Select * from archcto where cedarch =" & Wxelnrocedev & " order by fecha"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Do While Not Data1.Recordset.EOF
      List1.AddItem Data1.Recordset("id") & "===" & Data1.Recordset("nombredoc")
      Data1.Recordset.MoveNext
   Loop
End If
data_desc.DatabaseName = App.Path & "\desc.mdb"
data_desc.RecordSource = "desc"
data_desc.Refresh

If WElusuario = "BRUNO" Or WElusuario = "GFERNANDEZ" Or WElusuario = "JFERNAN" Or WElusuario = "MCOSTA" Or _
   WElusuario = "SDOMINGUEZ" Or WElusuario = "SPEREZ" Or WElusuario = "DARIOH" Or WElusuario = "MARCELOM" Or _
   WElusuario = "ENRIQUE" Or WElusuario = "JGIMENEZ" Or WElusuario = "MCURBELO" Or WElusuario = "AGUILLEN" Then
Else
   b_buscto.Enabled = False
   b_guarcto.Enabled = False
   b_bussan.Enabled = False
   b_guarsan.Enabled = False
   b_busvar.Enabled = False
   b_guarvar.Enabled = False
   
End If
End Sub

Private Sub List1_DblClick()
Dim X, Xbandlab As Integer
Dim Xlac As String
Xlac = ""
Xbandlab = 0
frm_archper.MousePointer = 11
If Dir(App.Path & "\archivos\temporal.pdf") <> "" Then

   Kill App.Path & "\archivos\temporal.pdf"
End If

If List1.ListIndex >= 0 Then
   For X = 1 To Len(List1.List(List1.ListIndex))
       If Mid(List1.List(List1.ListIndex), X, 1) = "=" Then
          Xbandlab = 1
       Else
          If Xbandlab <> 1 Then
             Xlac = Xlac + Mid(List1.List(List1.ListIndex), X, 1)
          End If
       End If
   Next
   If Xlac <> "" Then
      Set pdffile = New ADODB.Stream
      pdffile.Type = adTypeBinary
      pdffile.Open
      Adodc1.RecordSource = "Select * from archcto where id =" & Xlac
      Adodc1.Refresh
      If IsNull(Adodc1.Recordset("arch")) = False Then
         pdffile.Write Adodc1.Recordset("arch").value
         Dim pdfname As String
         pdfname = "temporal"
         pdffile.SaveToFile "" & App.Path & "\archivos\" & pdfname & ".pdf", adSaveCreateOverWrite
         pdffile.Close
         Set pdffile = Nothing
         MsgBox "DOCUMENTO NRO:" & Xlac
'         Shell data_desc.Recordset("desc") & " " & App.Path & "\archivos\temporal" & ".pdf", vbMaximizedFocus
         frm_archper.MousePointer = 0
         ShellExecute Me.hwnd, "open", App.Path & "\archivos\temporal" & ".pdf", "", "", 4
      Else
         MsgBox "No hay archivo"
         pdffile.Close
         Set pdffile = Nothing
      End If
   Else
      frm_archper.MousePointer = 0
      MsgBox "No hay documento"
   End If
End If
frm_archper.MousePointer = 0

End Sub

Private Sub List2_DblClick()
Dim X, Xbandlab As Integer
Dim Xlac As String
Xlac = ""
Xbandlab = 0
frm_archper.MousePointer = 11
If Dir(App.Path & "\archivos\temporal.pdf") <> "" Then

   Kill App.Path & "\archivos\temporal.pdf"
End If

If List2.ListIndex >= 0 Then
   For X = 1 To Len(List2.List(List2.ListIndex))
       If Mid(List2.List(List2.ListIndex), X, 1) = "=" Then
          Xbandlab = 1
       Else
          If Xbandlab <> 1 Then
             Xlac = Xlac + Mid(List2.List(List2.ListIndex), X, 1)
          End If
       End If
   Next
   If Xlac <> "" Then
      Set pdffile = New ADODB.Stream
      pdffile.Type = adTypeBinary
      pdffile.Open
      Adodc1.RecordSource = "Select * from archsanc where id =" & Xlac
      Adodc1.Refresh
      If IsNull(Adodc1.Recordset("arch")) = False Then
         pdffile.Write Adodc1.Recordset("arch").value
         Dim pdfname As String
         pdfname = "temporal"
         pdffile.SaveToFile "" & App.Path & "\archivos\" & pdfname & ".pdf", adSaveCreateOverWrite
         pdffile.Close
         Set pdffile = Nothing
         MsgBox "DOCUMENTO NRO:" & Xlac
'         Shell data_desc.Recordset("desc") & " " & App.Path & "\archivos\temporal" & ".pdf", vbMaximizedFocus
         frm_archper.MousePointer = 0
         ShellExecute Me.hwnd, "open", App.Path & "\archivos\temporal" & ".pdf", "", "", 4
      Else
         MsgBox "No hay archivo"
         pdffile.Close
         Set pdffile = Nothing
      End If
   Else
      frm_archper.MousePointer = 0
      MsgBox "No hay documento"
   End If
End If
frm_archper.MousePointer = 0

End Sub

Private Sub List3_DblClick()
Dim X, Xbandlab As Integer
Dim Xlac As String
Xlac = ""
Xbandlab = 0
frm_archper.MousePointer = 11
If Dir(App.Path & "\archivos\temporal.pdf") <> "" Then

   Kill App.Path & "\archivos\temporal.pdf"
End If

If List3.ListIndex >= 0 Then
   For X = 1 To Len(List3.List(List3.ListIndex))
       If Mid(List3.List(List3.ListIndex), X, 1) = "=" Then
          Xbandlab = 1
       Else
          If Xbandlab <> 1 Then
             Xlac = Xlac + Mid(List3.List(List3.ListIndex), X, 1)
          End If
       End If
   Next
   If Xlac <> "" Then
      Set pdffile = New ADODB.Stream
      pdffile.Type = adTypeBinary
      pdffile.Open
      Adodc1.RecordSource = "Select * from archotros where id =" & Xlac
      Adodc1.Refresh
      If IsNull(Adodc1.Recordset("arch")) = False Then
         pdffile.Write Adodc1.Recordset("arch").value
         Dim pdfname As String
         pdfname = "temporal"
         pdffile.SaveToFile "" & App.Path & "\archivos\" & pdfname & ".pdf", adSaveCreateOverWrite
         pdffile.Close
         Set pdffile = Nothing
         MsgBox "DOCUMENTO NRO:" & Xlac
'         Shell data_desc.Recordset("desc") & " " & App.Path & "\archivos\temporal" & ".pdf", vbMaximizedFocus
         frm_archper.MousePointer = 0
         ShellExecute Me.hwnd, "open", App.Path & "\archivos\temporal" & ".pdf", "", "", 4
      Else
         MsgBox "No hay archivo"
         pdffile.Close
         Set pdffile = Nothing
      End If
   Else
      frm_archper.MousePointer = 0
      MsgBox "No hay documento"
   End If
End If
frm_archper.MousePointer = 0

End Sub

