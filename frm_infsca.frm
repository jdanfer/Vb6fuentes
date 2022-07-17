VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_infsca 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de Servicio control ambulatorio"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5205
   Icon            =   "frm_infsca.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   5205
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
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
      Height          =   300
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2520
      Picture         =   "frm_infsca.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Caption         =   "Datos de informe"
      Height          =   2535
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   4695
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_infsca.frx":0B14
         Left            =   1560
         List            =   "frm_infsca.frx":0B1E
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1200
         Width           =   2895
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   3000
         TabIndex        =   4
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
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
         Left            =   1560
         TabIndex        =   3
         Top             =   480
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
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Opción:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "FECHAS:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frm_infsca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'Data2.RecordSource = "infcli"
'Data2.Refresh
'If Data2.Recordset.RecordCount > 0 Then
'   Data2.Recordset.MoveFirst
'   Do While Not Data2.Recordset.EOF
'      Data2.Recordset.Delete
'      Data2.Recordset.MoveNext
'   Loop
'End If
Dim Xobjexel22 As Excel.Application
Dim Xlibexel22 As Excel.Workbook

Dim Xarchexel22 As New Excel.Worksheet
Dim Xqdia, Xcanxdia, Xlin, XCol As Integer
Dim Xtotreg As Long
Dim Xarchtex As String
Dim Textofecha As String

Dim Xlabrir3 As New Excel.Application
Dim Xlabancp As Integer
Dim Contar911 As Integer

On Error GoTo Quepasa

Xlin = 1
XCol = 1
Xtotreg = 0

If md.Text <> "__/__/____" Then
   If mh.Text <> "__/__/____" Then
      frm_infsca.MousePointer = 11
      If Combo1.ListIndex = 0 Then
         Data1.RecordSource = "select * from pendiente_sca where fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "#"
         Data1.Refresh
      Else
         If Combo1.ListIndex = 1 Then
            Data1.RecordSource = "select seguimiento_sca.fecha,seguimiento_sca.hora,seguimiento_sca.mat,seguimiento_sca.medicocod,seguimiento_sca.obs,seguimiento_sca.nro_ctrol,seguimiento_sca.fecha_prox,seguimiento_sca.fecha_alta,seguimiento_sca.id_seguimiento,us.id,us.nombre,us.apellidos " & _
            "from seguimiento_sca inner join us on seguimiento_sca.medicocod=us.id where seguimiento_sca.fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# and seguimiento_sca.fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "#"
            Data1.Refresh
         Else
            Data1.RecordSource = "select * from pendiente_sca where fecha <=#" & Format("01/01/2000", "yyyy/mm/dd") & "#"
            Data1.Refresh
        
         End If
      End If
      If Data1.Recordset.RecordCount > 0 Then
         If Combo1.ListIndex = 0 Then
            Set Xobjexel22 = New Excel.Application
            Set Xlibexel22 = Xobjexel22.Workbooks.Add
            Set Xarchexel22 = Xlibexel22.Worksheets.Add
            Xarchexel22.Name = "SCA"
            Xlibexel22.SaveAs ("C:\planillas\csa_" & Trim(str(Month(md.Text))) & Trim(str(Year(md.Text))) & ".xls")
            Xarchtex = "C:\planillas\csa_" & Trim(str(Month(md.Text))) & Trim(str(Year(md.Text))) & ".xls"
            Xqdia = 0
            Xcanxdia = 0
            Xarchexel22.Cells(Xlin, XCol) = "DEPARTAMENTO TI SAPP S.A."
            XCol = 6
            Xarchexel22.Cells(Xlin, XCol) = "FECHA: " & Format(Date, "dd/mm/yyyy")
            Xlin = Xlin + 1
            XCol = 2
            Xarchexel22.Range("A1", "C3").Font.Size = 16
            Xarchexel22.Cells(Xlin, XCol) = "PLANILLA DE SCA DESDE: " & md.Text & " HASTA: " & mh.Text
            Xarchexel22.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
            
            XCol = 1
            Xlin = Xlin + 2
            Xnrocan = Xnrocan + Xlin
            Xarchexel22.Range("A" & Trim(str(Xlin)), "G" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
            Xarchexel22.Range("A" & Trim(str(Xlin))).ColumnWidth = 12
            Xarchexel22.Cells(Xlin, XCol) = "FECHA"
            XCol = XCol + 1
            Xarchexel22.Range("B" & Trim(str(Xlin))).ColumnWidth = 10
            Xarchexel22.Cells(Xlin, XCol) = "HORA"
            XCol = XCol + 1
            Xarchexel22.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
            Xarchexel22.Cells(Xlin, XCol) = "MATRICULA"
            XCol = XCol + 1
            Xarchexel22.Range("D" & Trim(str(Xlin))).ColumnWidth = 35
            Xarchexel22.Cells(Xlin, XCol) = "NOMBRE"
            XCol = XCol + 1
            Xarchexel22.Range("E" & Trim(str(Xlin))).ColumnWidth = 6
            Xarchexel22.Cells(Xlin, XCol) = "BASE"
            XCol = XCol + 1
            Xarchexel22.Range("F" & Trim(str(Xlin))).ColumnWidth = 25
            Xarchexel22.Cells(Xlin, XCol) = "MEDICO"
            XCol = XCol + 1
            Xarchexel22.Range("G" & Trim(str(Xlin))).ColumnWidth = 12
            Xarchexel22.Cells(Xlin, XCol) = "FEC.CIERRE"
             
            Xlin = Xlin + 1
            XCol = 1
         
            Data1.Recordset.MoveFirst
            Do While Not Data1.Recordset.EOF
               Xarchexel22.Cells(Xlin, XCol) = "'" & Format(Data1.Recordset("fecha"), "dd/mm/yyyy")
               XCol = XCol + 1
               Xarchexel22.Cells(Xlin, XCol) = Data1.Recordset("hora")
               XCol = XCol + 1
               Xarchexel22.Cells(Xlin, XCol) = Data1.Recordset("mat")
               XCol = XCol + 1
               Xarchexel22.Cells(Xlin, XCol) = Data1.Recordset("nombre")
               XCol = XCol + 1
               Xarchexel22.Cells(Xlin, XCol) = Data1.Recordset("base")
               XCol = XCol + 1
               Xarchexel22.Cells(Xlin, XCol) = Data1.Recordset("mediconom")
               XCol = XCol + 1
               If IsNull(Data1.Recordset("fecha_cierre")) = False Then
                  Xarchexel22.Cells(Xlin, XCol) = "'" & Format(Data1.Recordset("fecha_cierre"), "dd/mm/yyyy")
               End If
               XCol = 1
               Data1.Recordset.MoveNext
               Xlin = Xlin + 1
               Xtotreg = Xtotreg + 1
            Loop
            Xlin = Xlin + 1
            XCol = 2
            Xarchexel22.Cells(Xlin, XCol) = "TOTAL DE REGISTROS:" & Xtotreg
            Xlibexel22.Save
            Xlibexel22.Close
            Xobjexel22.Quit
            Xlabrir3.Workbooks.Open Xarchtex, , False
            Xlabrir3.Visible = True
            Xlabrir3.WindowState = xlMaximized
            frm_infsca.MousePointer = 0
            MsgBox "Terminado"
         End If
               
         If Combo1.ListIndex = 1 Then
            Set Xobjexel22 = New Excel.Application
            Set Xlibexel22 = Xobjexel22.Workbooks.Add
            Set Xarchexel22 = Xlibexel22.Worksheets.Add
            Xarchexel22.Name = "SCA"
            Xlibexel22.SaveAs ("C:\planillas\csa_" & Trim(str(Month(md.Text))) & Trim(str(Year(md.Text))) & ".xls")
            Xarchtex = "C:\planillas\csa_" & Trim(str(Month(md.Text))) & Trim(str(Year(md.Text))) & ".xls"
            Xqdia = 0
            Xcanxdia = 0
            Xarchexel22.Cells(Xlin, XCol) = "DEPARTAMENTO TI SAPP S.A."
            XCol = 6
            Xarchexel22.Cells(Xlin, XCol) = "FECHA: " & Format(Date, "dd/mm/yyyy")
            Xlin = Xlin + 1
            XCol = 2
            Xarchexel22.Range("A1", "C3").Font.Size = 16
            Xarchexel22.Cells(Xlin, XCol) = "PLANILLA DE CONTROLES SCA DESDE: " & md.Text & " HASTA: " & mh.Text
            Xarchexel22.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
            
            XCol = 1
            Xlin = Xlin + 2
            Xnrocan = Xnrocan + Xlin
            Xarchexel22.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
            Xarchexel22.Range("A" & Trim(str(Xlin))).ColumnWidth = 12
            Xarchexel22.Cells(Xlin, XCol) = "FECHA"
            XCol = XCol + 1
            Xarchexel22.Range("B" & Trim(str(Xlin))).ColumnWidth = 10
            Xarchexel22.Cells(Xlin, XCol) = "HORA"
            XCol = XCol + 1
            Xarchexel22.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
            Xarchexel22.Cells(Xlin, XCol) = "MATRICULA"
            XCol = XCol + 1
            Xarchexel22.Range("D" & Trim(str(Xlin))).ColumnWidth = 35
            Xarchexel22.Cells(Xlin, XCol) = "NOMBRE"
            XCol = XCol + 1
            Xarchexel22.Range("E" & Trim(str(Xlin))).ColumnWidth = 6
            Xarchexel22.Cells(Xlin, XCol) = "BASE SCA"
            XCol = XCol + 1
            Xarchexel22.Range("F" & Trim(str(Xlin))).ColumnWidth = 25
            Xarchexel22.Cells(Xlin, XCol) = "MEDICO"
            XCol = XCol + 1
            Xarchexel22.Range("G" & Trim(str(Xlin))).ColumnWidth = 12
            Xarchexel22.Cells(Xlin, XCol) = "PROX.CTROL."
            XCol = XCol + 1
            Xarchexel22.Range("H" & Trim(str(Xlin))).ColumnWidth = 10
            Xarchexel22.Cells(Xlin, XCol) = "Nro.CTROL."
            XCol = XCol + 1
            Xarchexel22.Range("I" & Trim(str(Xlin))).ColumnWidth = 50
            Xarchexel22.Cells(Xlin, XCol) = "DETALLE"
             
            Xlin = Xlin + 1
            XCol = 1
         
            Data1.Recordset.MoveFirst
            Do While Not Data1.Recordset.EOF
               Xarchexel22.Cells(Xlin, XCol) = "'" & Format(Data1.Recordset("fecha"), "dd/mm/yyyy")
               XCol = XCol + 1
               Xarchexel22.Cells(Xlin, XCol) = Data1.Recordset("hora")
               XCol = XCol + 1
               Data2.RecordSource = "select * from pendiente_sca where id =" & Data1.Recordset("id_seguimiento")
               Data2.Refresh
               If Data2.Recordset.RecordCount > 0 Then
                  Data2.Recordset.MoveFirst
                  Xarchexel22.Cells(Xlin, XCol) = Data2.Recordset("mat")
                  XCol = XCol + 1
                  Xarchexel22.Cells(Xlin, XCol) = Data2.Recordset("nombre")
                  XCol = XCol + 1
                  Xarchexel22.Cells(Xlin, XCol) = Data2.Recordset("base")
                  XCol = XCol + 1
                  Xarchexel22.Cells(Xlin, XCol) = Data2.Recordset("mediconom")
                  XCol = XCol + 1
               Else
                  Xarchexel22.Cells(Xlin, XCol) = "0"
                  XCol = XCol + 1
                  Xarchexel22.Cells(Xlin, XCol) = "NN"
                  XCol = XCol + 1
                  Xarchexel22.Cells(Xlin, XCol) = "0"
                  XCol = XCol + 1
                  Xarchexel22.Cells(Xlin, XCol) = "Sin dato"
                  XCol = XCol + 1
               End If
               Xarchexel22.Cells(Xlin, XCol) = "'" & Format(Data1.Recordset("fecha_prox"), "dd/mm/yyyy")
               XCol = XCol + 1
               Xarchexel22.Cells(Xlin, XCol) = Data1.Recordset("nro_ctrol")
               XCol = XCol + 1
               Xarchexel22.Cells(Xlin, XCol) = Data1.Recordset("obs")
               XCol = 1
               Data1.Recordset.MoveNext
               Xlin = Xlin + 1
               Xtotreg = Xtotreg + 1
            Loop
            Xlin = Xlin + 1
            XCol = 2
            Xarchexel22.Cells(Xlin, XCol) = "TOTAL DE REGISTROS:" & Xtotreg
            Xlibexel22.Save
            Xlibexel22.Close
            Xobjexel22.Quit
            Xlabrir3.Workbooks.Open Xarchtex, , False
            Xlabrir3.Visible = True
            Xlabrir3.WindowState = xlMaximized
            frm_infsca.MousePointer = 0
            MsgBox "Terminado"
         End If
      Else
         frm_infsca.MousePointer = 0
         MsgBox "No hay registros"
      End If
   End If
End If
         
Exit Sub

Quepasa:
        If Err.Number = 3155 Then
           frm_infsca.MousePointer = 0
           MsgBox "Error al generar " & Err.Description
        Else
            frm_infsca.MousePointer = 0
           
           MsgBox "Error al generar " & Err.Description
        End If
         
End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"

Data2.Connect = "odbc;dsn=" & Xconexrmt & ";"


End Sub

Private Sub md_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mh.SetFocus
End If

End Sub
