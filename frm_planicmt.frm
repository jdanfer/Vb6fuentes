VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_planicmt 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planilla mensual CMT"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6270
   Icon            =   "frm_planicmt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   6270
   StartUpPosition =   1  'CenterOwner
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
      Top             =   1920
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Procesar"
      Height          =   615
      Left            =   2160
      Picture         =   "frm_planicmt.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Datos para informe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5655
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   3480
         TabIndex        =   3
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox md 
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fechas:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frm_planicmt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command1.Enabled = False
Dim Xobjexel22 As Excel.Application
Dim Xlibexel22 As Excel.Workbook

Dim Xarchexel22 As New Excel.Worksheet
Dim Xqdia, Xcanxdia, Xlin, XCol As Integer
Dim Xtotreg As Long
Dim Xarchtex As String
Dim Textofecha As String

Dim Xlabrir3 As New Excel.Application
Dim Xlabancp As Integer

Xlin = 1
XCol = 1

If md.Text <> "__/__/____" Then
   If mh.Text <> "__/__/____" Then
      Data1.RecordSource = "select linmmdd.fecha,linmmdd.cod_prod,linmmdd.cod_cli,linmmdd.nom_cli,linmmdd.base,linmmdd.convenio,convenio.cnv_codigo," & _
      "convenio.cnv_grupo from linmmdd inner join convenio on linmmdd.convenio=convenio.cnv_codigo where " & _
      "linmmdd.fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and linmmdd.fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "# and linmmdd.cod_prod in (3,10018,10050,14005)"
      Data1.Refresh
      If Data1.Recordset.RecordCount > 0 Then
         Data1.Recordset.MoveFirst
         Set Xobjexel22 = New Excel.Application
         Set Xlibexel22 = Xobjexel22.Workbooks.Add
         Set Xarchexel22 = Xlibexel22.Worksheets.Add
         Xarchexel22.Name = Trim("CMT")
         Xlibexel22.SaveAs ("C:\planillas\CMT_" & Trim(str(Month(md.Text))) & Trim(str(Year(md.Text))) & ".xls")
         Xarchtex = "C:\planillas\CMT_" & Trim(str(Month(md.Text))) & Trim(str(Year(md.Text))) & ".xls"
         Xqdia = 0
         Xcanxdia = 0
         Xarchexel22.Cells(Xlin, XCol) = "DEPARTAMENTO TI SAPP S.A."
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel22.Range("A1", "C3").Font.Size = 16
         Xarchexel22.Cells(Xlin, XCol) = "PLANILLA DE CMT " & " DESDE: " & md.Text & " HASTA: " & mh.Text
         Xarchexel22.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel22.Range("A" & Trim(str(Xlin)), "E" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel22.Range("A" & Trim(str(Xlin))).ColumnWidth = 6
         Xarchexel22.Cells(Xlin, XCol) = "DIA"
         XCol = XCol + 1
         Xarchexel22.Range("B" & Trim(str(Xlin))).ColumnWidth = 6
         Xarchexel22.Cells(Xlin, XCol) = "MES"
         XCol = XCol + 1
         Xarchexel22.Range("C" & Trim(str(Xlin))).ColumnWidth = 6
         Xarchexel22.Cells(Xlin, XCol) = "AÑO"
         XCol = XCol + 1
         Xarchexel22.Range("D" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel22.Cells(Xlin, XCol) = "NOMBRE"
         XCol = XCol + 1
         Xarchexel22.Range("E" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel22.Cells(Xlin, XCol) = "ZONA"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not Data1.Recordset.EOF
            If IsNull(Data1.Recordset("cnv_grupo")) = False Then
               If Trim(Data1.Recordset("cnv_grupo")) <> "" Then
                  If Data1.Recordset("cnv_grupo") = "CPS" Or _
                     Data1.Recordset("cnv_grupo") = "CASH" Or _
                     Data1.Recordset("cnv_grupo") = "CASMU" Or _
                     Data1.Recordset("cnv_grupo") = "SEMM" Or _
                     Data1.Recordset("cnv_grupo") = "CAUTE" Or _
                     Data1.Recordset("cnv_grupo") = "911" Then
                     Xarchexel22.Cells(Xlin, XCol) = Trim(str(Day(Data1.Recordset("fecha"))))
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = Trim(str(Month(Data1.Recordset("fecha"))))
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = Trim(str(Year(Data1.Recordset("fecha"))))
                     XCol = XCol + 1
                     Xarchexel22.Cells(Xlin, XCol) = Data1.Recordset("nom_cli")
                     XCol = XCol + 1
                     If Data1.Recordset("base") = 1 Or Data1.Recordset("base") = 2 Or _
                        Data1.Recordset("base") = 3 Or Data1.Recordset("base") = 4 Or _
                        Data1.Recordset("base") = 18 Or Data1.Recordset("base") = 19 Then
                        Xarchexel22.Cells(Xlin, XCol) = "Zona: 1"
                     Else
                        Xarchexel22.Cells(Xlin, XCol) = "Zona: 2"
                     End If
                     
                     Xcanxdia = Xcanxdia + 1
                     Xlin = Xlin + 1
                     XCol = 1
                  End If
               Else
                  Xarchexel22.Cells(Xlin, XCol) = Trim(str(Day(Data1.Recordset("fecha"))))
                  XCol = XCol + 1
                  Xarchexel22.Cells(Xlin, XCol) = Trim(str(Month(Data1.Recordset("fecha"))))
                  XCol = XCol + 1
                  Xarchexel22.Cells(Xlin, XCol) = Trim(str(Year(Data1.Recordset("fecha"))))
                  XCol = XCol + 1
                  Xarchexel22.Cells(Xlin, XCol) = Data1.Recordset("nom_cli")
                  XCol = XCol + 1
                  If Data1.Recordset("base") = 1 Or Data1.Recordset("base") = 2 Or _
                     Data1.Recordset("base") = 3 Or Data1.Recordset("base") = 4 Or _
                     Data1.Recordset("base") = 18 Or Data1.Recordset("base") = 19 Then
                     Xarchexel22.Cells(Xlin, XCol) = "Zona: 1"
                  Else
                     Xarchexel22.Cells(Xlin, XCol) = "Zona: 2"
                  End If
                  
                  Xcanxdia = Xcanxdia + 1
                  Xlin = Xlin + 1
                  XCol = 1
               End If
            Else
               Xarchexel22.Cells(Xlin, XCol) = Trim(str(Day(Data1.Recordset("fecha"))))
               XCol = XCol + 1
               Xarchexel22.Cells(Xlin, XCol) = Trim(str(Month(Data1.Recordset("fecha"))))
               XCol = XCol + 1
               Xarchexel22.Cells(Xlin, XCol) = Trim(str(Year(Data1.Recordset("fecha"))))
               XCol = XCol + 1
               Xarchexel22.Cells(Xlin, XCol) = Data1.Recordset("nom_cli")
               XCol = XCol + 1
               If Data1.Recordset("base") = 1 Or Data1.Recordset("base") = 2 Or _
                  Data1.Recordset("base") = 3 Or Data1.Recordset("base") = 4 Or _
                  Data1.Recordset("base") = 18 Or Data1.Recordset("base") = 19 Then
                  Xarchexel22.Cells(Xlin, XCol) = 1
               Else
                  Xarchexel22.Cells(Xlin, XCol) = 2
               End If
               
               Xcanxdia = Xcanxdia + 1
               Xlin = Xlin + 1
               XCol = 1
            
            End If
            Data1.Recordset.MoveNext
         Loop
         Xlin = Xlin + 1
         XCol = 2
         MsgBox "Proceso terminado"
         Command1.Enabled = True
         Xarchexel22.Cells(Xlin, XCol) = "TOTAL DE REGISTROS:" & Xcanxdia
         Xlin = Xlin + 1
         XCol = 2
         Xlibexel22.Save
         Xlibexel22.Close
         Xobjexel22.Quit
         Xlabrir3.Workbooks.Open Xarchtex, , False
         Xlabrir3.Visible = True
         Xlabrir3.WindowState = xlMaximized
         
      End If
   End If
End If
Command1.Enabled = True


End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"

End Sub

Private Sub md_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mh.SetFocus
End If

End Sub
