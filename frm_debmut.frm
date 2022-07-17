VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_debmut 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Débitos mutuales"
   ClientHeight    =   5925
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7935
   Icon            =   "frm_debmut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   7935
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3840
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Excel 8.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5400
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Importar datos desde excel"
      Height          =   375
      Left            =   4560
      Picture         =   "frm_debmut.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Importar desde archivo ""debitos mutuales.xls"" guardado en carpeta Debitos"
      Top             =   120
      Width           =   3015
   End
   Begin VB.CommandButton b_imp 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      Picture         =   "frm_debmut.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Emitir a excel"
      Top             =   5280
      Width           =   495
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5520
      Visible         =   0   'False
      Width           =   4455
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_debmut.frx":109E
      Height          =   4815
      Left            =   120
      OleObjectBlob   =   "frm_debmut.frx":10B2
      TabIndex        =   2
      Top             =   480
      Width           =   7455
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00404040&
      Caption         =   "Buscar matrícula:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frm_debmut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_imp_Click()
Dim Xobjexel22 As Excel.Application
Dim Xlibexel22 As Excel.Workbook
Dim Xarchexel22 As New Excel.Worksheet
Dim Xlin, XCol As Integer
Dim Xtotreg, Xsub As Long
Dim Xarchtex As String
Dim Xlabrir3 As New Excel.Application
Data3.RecordSource = "select * from mutuales order by socio"
Data3.Refresh

frm_debmut.MousePointer = 11
b_imp.Enabled = False
If Data3.Recordset.RecordCount > 0 Then
   Data3.Recordset.MoveFirst
   Set Xobjexel22 = New Excel.Application
   Set Xlibexel22 = Xobjexel22.Workbooks.Add
   Set Xarchexel22 = Xlibexel22.Worksheets.Add
   Xarchexel22.Name = Trim("debmutuales")
   Xlibexel22.SaveAs ("C:\planillas\debitosMut.xls")
   Xarchtex = "C:\planillas\debitosMut.xls"
   Xlin = 1
   XCol = 1
   Xarchexel22.Cells(Xlin, XCol) = "DEPARTAMENTO TI SAPP S.A."
   Xlin = Xlin + 1
   XCol = XCol + 1
   Xarchexel22.Range("A1", "C3").Font.Size = 16
   Xarchexel22.Range("A" & Trim(str(Xlin)), "J" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
   Xarchexel22.Cells(Xlin, XCol) = "INFORME DE DEBITOS MUTUALES ---FECHA:" & Format(Date, "dd/mm/yyyy")
   XCol = 1
   Xlin = Xlin + 2
   Xnrocan = Xnrocan + Xlin
   Xarchexel22.Range("A" & Trim(str(Xlin)), "J" & Trim(str(Xlin))).Interior.color = RGB(58, 176, 218)
   Xarchexel22.Range("A" & Trim(str(Xlin))).ColumnWidth = 15
   Xarchexel22.Cells(Xlin, XCol) = "MATRICULA"
   XCol = XCol + 1
   Xarchexel22.Range("B" & Trim(str(Xlin))).ColumnWidth = 25
   Xarchexel22.Cells(Xlin, XCol) = "NOMBRE"
   XCol = XCol + 1
   Xarchexel22.Range("C" & Trim(str(Xlin))).ColumnWidth = 15
   Xarchexel22.Cells(Xlin, XCol) = "IMPORTE"
   XCol = XCol + 1
   Xarchexel22.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
   Xarchexel22.Cells(Xlin, XCol) = "RECIBO"
      
   Xlin = Xlin + 1
   XCol = 1
   Xsub = 0
   Do While Not Data3.Recordset.EOF
      Xarchexel22.Cells(Xlin, XCol) = Data3.Recordset("socio")
      XCol = XCol + 1
      Data4.RecordSource = "select * from clientes where cl_codigo =" & Data3.Recordset("socio")
      Data4.Refresh
      If Data4.Recordset.RecordCount > 0 Then
         Xarchexel22.Cells(Xlin, XCol) = Data4.Recordset("cl_apellid")
      Else
         Xarchexel22.Cells(Xlin, XCol) = "NN"
      End If
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = Data3.Recordset("importe_deuda")
      XCol = XCol + 1
      Xarchexel22.Cells(Xlin, XCol) = Data3.Recordset("recibo")
      Xsub = Xsub + Data3.Recordset("importe_deuda")
      Xlin = Xlin + 1
      XCol = 1
      Xtotreg = Xtotreg + 1
      Data3.Recordset.MoveNext
   Loop
   Xlin = Xlin + 1
   XCol = 1
   Xarchexel22.Cells(Xlin, XCol) = "Total Registros: " & Trim(str(Xtotreg)) & "----$." & Format(Xsub, "Standard")
   Xlin = Xlin + 1
   XCol = 1
   Xarchexel22.Cells(Xlin, XCol) = "FECHA DE EMISION:" & Format(Date, "dd/mm/yyyy")
   Xlibexel22.Save
   Xlibexel22.Close
   Xobjexel22.Quit
   Xlabrir3.Workbooks.Open Xarchtex, , False
   Xlabrir3.Visible = True
   Xlabrir3.WindowState = xlMaximized
   frm_debmut.MousePointer = 0
   MsgBox "Terminado"
Else
   frm_debmut.MousePointer = 0
   MsgBox "No hay registros"
End If
b_imp.Enabled = True

End Sub

Private Sub Command1_Click()
Dim ProcesarSi As String
Command1.Enabled = False
ProcesarSi = MsgBox("Desea borrar los datos actuales y cargar nuevos?", vbYesNo + vbInformation, "Débitos")
frm_debmut.MousePointer = 11
If ProcesarSi = vbYes Then
   If Data2.Recordset.RecordCount > 0 Then
      Data2.Recordset.MoveFirst
      Data3.RecordSource = "select * from mutuales where pendiente is null"
      Data3.Refresh
      If Data3.Recordset.RecordCount > 0 Then
         Data3.Recordset.MoveFirst
         Do While Not Data3.Recordset.EOF
            Data3.Recordset.Delete
            Data3.Recordset.MoveNext
         Loop
      End If
      Data1.RecordSource = "select * from mutuales order by socio"
      Data1.Refresh
      DoEvents
      Data3.RecordSource = "select * from mutuales"
      Data3.Refresh
      Do While Not Data2.Recordset.EOF
         If IsNull(Data2.Recordset("importe")) = False Then
            If Trim(str(Data2.Recordset("importe"))) <> "" Then
               If IsNull(Data2.Recordset("matricula_sapp")) = False Then
                  Data3.RecordSource = "select * from mutuales where socio =" & Data2.Recordset("matricula_sapp")
                  Data3.Refresh
                  If Data3.Recordset.RecordCount > 0 Then
                     Data3.Recordset.Edit
                     Data3.Recordset("importe_deuda") = Data3.Recordset("importe_deuda") + Data2.Recordset("importe")
                     Data3.Recordset.Update
                  Else
                     Data3.Recordset.AddNew
                     Data3.Recordset("socio") = Data2.Recordset("matricula_sapp")
                     Data3.Recordset("importe_deuda") = Data2.Recordset("importe")
                     Data3.Recordset("recibo") = Data2.Recordset("recibo")
                     Data3.Recordset("locked") = 0
                     Data3.Recordset.Update
                  End If
               End If
            End If
         End If
         Data2.Recordset.MoveNext
      Loop
   End If
   frm_debmut.MousePointer = 0
   Data1.RecordSource = "select * from mutuales order by socio"
   Data1.Refresh
   MsgBox "Proceso terminado. Se importaron correctamente los datos.", vbInformation
   
End If
Command1.Enabled = True
frm_debmut.MousePointer = 0

End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.RecordSource = "select * from mutuales order by socio"
Data1.Refresh

Data3.Connect = "odbc;dsn=" & Xconexrmt & ";"

Data4.Connect = "odbc;dsn=" & Xconexrmt & ";"


Data2.DatabaseName = "C:\debitos\debitos mutuales.xls"
Data2.RecordSource = "debitos$"
Data2.Refresh

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Trim(Text1.Text) <> "" Then
      Data1.RecordSource = "select * from mutuales where socio >=" & Text1.Text & " order by socio"
   Else
      Data1.RecordSource = "select * from mutuales order by socio"
   End If
   Data1.Refresh
   DBGrid1.SetFocus
End If

End Sub
