VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_inffaccnv 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes facturación convenios"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4860
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_inffaccnv.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5670
   ScaleWidth      =   4860
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   2160
      Top             =   4440
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   582
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
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   375
      Left            =   1080
      Top             =   3720
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
      Caption         =   "data1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00FF0000&
      Caption         =   "Por fecha de comprobante"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   1320
      Width           =   3375
   End
   Begin VB.Data data_cab 
      Caption         =   "data_cab"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   1440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      ItemData        =   "frm_inffaccnv.frx":0442
      Left            =   2040
      List            =   "frm_inffaccnv.frx":044F
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Data data_cnv 
      Caption         =   "data_cnv"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4080
      Picture         =   "frm_inffaccnv.frx":046C
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Consultar cliente"
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox t_cli 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   2040
      TabIndex        =   11
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox t_cob 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Text            =   "0"
      Top             =   1800
      Width           =   975
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   2160
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
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
      Top             =   1560
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3720
      Picture         =   "frm_inffaccnv.frx":056E
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   5040
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_inffaccnv.frx":0AF8
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Procesar"
      Top             =   5040
      Width           =   615
   End
   Begin VB.OptionButton Option3 
      BackColor       =   &H00C00000&
      Caption         =   "Clientes Activos"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   4560
      Width           =   3375
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C00000&
      Caption         =   "Facturas pendientes"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   4080
      Width           =   3375
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C00000&
      Caption         =   "Facturas realizadas"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   3600
      Width           =   3375
   End
   Begin MSMask.MaskEdBox mfh 
      Height          =   375
      Left            =   2040
      TabIndex        =   2
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mfd 
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FF0000&
      Caption         =   "F.Pago:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FF0000&
      Caption         =   "CLIENTE:"
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   240
      TabIndex        =   10
      Top             =   2400
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FF0000&
      Caption         =   "COBRADOR:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Line Line2 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   0
      X2              =   4920
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      BorderWidth     =   3
      X1              =   0
      X2              =   4920
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FF0000&
      Caption         =   "FECHAS:"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   2760
      Picture         =   "frm_inffaccnv.frx":1082
      Stretch         =   -1  'True
      Top             =   3000
      Width           =   1695
   End
End
Attribute VB_Name = "frm_inffaccnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim Xdoc As Long
Dim Xtot, XtotalFact As Double
Xdoc = 0
Xtot = 0
XtotalFact = 0
Dim XCol, Xlin, Xnrocan, Xcolfija, Xcantsrv, Xcanttot As Long
Dim Xarchtex As String
Dim Xobjexel As Excel.Application
Dim Xlibexel As Excel.Workbook
Dim Xarchexel As New Excel.Worksheet
Dim Xlabrir As New Excel.Application
XCol = 1
Xlin = 1

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infvtas"
'   data_lin3.RecordSource = "Select * from indica_enfc where id =" & 50
'   data_lin3.Refresh

data_inf.RecordSource = "infvtas"
data_inf.Refresh
Command1.Enabled = False
Command2.Enabled = False
frm_inffaccnv.MousePointer = 11
Dim Xnrotipof As Integer

If mfd.Text = "__/__/____" Then
Else
   If mfh.Text = "__/__/____" Then
   Else
      If Option1.Value = True Then
         Data1.ConnectionString = "dsn=" & Xconexrmt
         If t_cli.Text <> "" Then
            If Check1.Value = 1 Then
               Data1.RecordSource = "Select * from linmmdd where cod_cli =" & t_cli.Text & " and fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' order by fecha"
               Data1.Refresh
            Else
               Data1.RecordSource = "Select * from linmmdd where cod_cli =" & t_cli.Text & " and realizada >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and realizada <='" & Format(mfh.Text, "yyyy-mm-dd") & "' order by realizada"
               Data1.Refresh
            End If
         Else
            If Combo1.ListIndex <= 0 Then
               If Check1.Value = 1 Then
                  Data1.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base in(101,102) order by fecha"
                  Data1.Refresh
               Else
                  Data1.RecordSource = "Select * from linmmdd where realizada >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and realizada <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base in(101,102) order by realizada"
                  Data1.Refresh
               End If
            Else
               If Combo1.ListIndex = 1 Then
                  If Check1.Value = 1 Then
                     Data1.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base in(101,102) and tipo ='" & "CREDITO" & "' order by fecha"
                     Data1.Refresh
                  Else
                     Data1.RecordSource = "Select * from linmmdd where realizada >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and realizada <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base in(101,102) and tipo ='" & "CREDITO" & "' order by realizada"
                     Data1.Refresh
                  End If
               Else
                  If Combo1.ListIndex = 2 Then
                     If Check1.Value = 1 Then
                        Data1.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base in(101,102) and tipo ='" & "CONTADO" & "' and tot_lin >" & 0 & " order by fecha"
                        Data1.Refresh
                     Else
                        Data1.RecordSource = "Select * from linmmdd where realizada >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and realizada <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base in(101,102) and tipo ='" & "CONTADO" & "' and tot_lin >" & 0 & " order by realizada"
                        Data1.Refresh
                     End If
                  Else
                     If Combo1.ListIndex = 3 Then
                        If Check1.Value = 1 Then
                           Data1.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base in(101,102) and pendiente in ('N','C') and tot_lin >" & 0 & " order by fecha"
                           Data1.Refresh
                        Else
                           Data1.RecordSource = "Select * from linmmdd where realizada >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and realizada <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base in(101,102) and pendiente in ('N','C') and tot_lin >" & 0 & " order by realizada"
                           Data1.Refresh
                        End If
                     Else
                        If Check1.Value = 1 Then
                           Data1.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base in(101,102) and tot_lin >" & 0 & " order by fecha"
                           Data1.Refresh
                        Else
                           Data1.RecordSource = "Select * from linmmdd where realizada >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and realizada <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base in(101,102) and tot_lin >" & 0 & " order by realizada"
                           Data1.Refresh
                        End If
                     End If
                  End If
               End If
            End If
         End If
         If Data1.Recordset.RecordCount > 0 Then
            Data1.Recordset.MoveFirst
            
            Set Xobjexel = New Excel.Application
            Set Xlibexel = Xobjexel.Workbooks.Add
            Set Xarchexel = Xlibexel.Worksheets.Add
            Xarchexel.Name = "FACTURACION CONVENIOS"
            Xlibexel.SaveAs ("C:\planillas\" & "Infvtas" & ".xls")
            Xarchtex = "C:\planillas\" & "Infvtas" & ".xls"
                        
            Xarchexel.Cells(Xlin, XCol) = "SAPP - CÓMPUTOS"
            XCol = 9
            Xarchexel.Cells(Xlin, XCol) = "FECHA:" & Format(Date, "dd/mm/yyyy")
            Xlin = Xlin + 1
            XCol = 2
            Xarchexel.Range("A1", "C3").Font.Size = 16
            If Check1.Value = 1 Then
               Xarchexel.Cells(Xlin, XCol) = "INFORME FACTURACIÓN CONVENIOS FECHA DE COMPROBANTE DESDE: " & mfd.Text & " HASTA: " & mfh.Text
            Else
               Xarchexel.Cells(Xlin, XCol) = "INFORME FACTURACIÓN CONVENIOS FECHA DE FIRMA DESDE: " & mfd.Text & " HASTA: " & mfh.Text
            End If
            Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
            XCol = 1
            Xlin = Xlin + 2
            Xnrocan = Xnrocan + Xlin
'            Xarchexel.Range("A4", "AD" & Trim(Str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
'            Xarchexel.Range("A4", "AD" & Trim(Str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
'            Xarchexel.Range("A4", "AD" & Trim(Str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
'            Xarchexel.Range("A4", "AD" & Trim(Str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
'            Xarchexel.Range("A4", "AD" & Trim(Str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
'            Xarchexel.Range("A4", "AD" & Trim(Str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
            Xarchexel.Range("A" & Trim(str(Xlin)), "AD" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
            Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 15
            If Check1.Value = 1 Then
               Xarchexel.Cells(Xlin, XCol) = "FECHA COMP."
            Else
               Xarchexel.Cells(Xlin, XCol) = "FECHA FIRMA"
            End If
            XCol = XCol + 1
            Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 15
            Xarchexel.Cells(Xlin, XCol) = "RUT"
            
            XCol = XCol + 1
            Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 15
            Xarchexel.Cells(Xlin, XCol) = "COD.CONV"
            
            XCol = XCol + 1
            Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 35
            Xarchexel.Cells(Xlin, XCol) = "NOMBRE/DENOMINACION"
            XCol = XCol + 1
            Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 12
            Xarchexel.Cells(Xlin, XCol) = "RUBRO CONT."
            XCol = XCol + 1
            Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 12
            Xarchexel.Cells(Xlin, XCol) = "NRO.FACTURA"
            XCol = XCol + 1
            Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 10
            Xarchexel.Cells(Xlin, XCol) = "MONEDA"
            XCol = XCol + 1
            Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 15
            Xarchexel.Cells(Xlin, XCol) = "TIPO DOC."
            XCol = XCol + 1
            Xarchexel.Range("I" & Trim(str(Xlin))).ColumnWidth = 15
            Xarchexel.Cells(Xlin, XCol) = "F.PAGO"
            XCol = XCol + 1
            Xarchexel.Range("J" & Trim(str(Xlin))).ColumnWidth = 15
            If Check1.Value = 1 Then
               Xarchexel.Cells(Xlin, XCol) = "FECHA FIRMA"
            Else
               Xarchexel.Cells(Xlin, XCol) = "FECHA COMP."
            End If
            XCol = XCol + 1
            Xarchexel.Range("K" & Trim(str(Xlin))).ColumnWidth = 35
            Xarchexel.Cells(Xlin, XCol) = "DETALLE"
            XCol = XCol + 1
            Xarchexel.Range("L" & Trim(str(Xlin))).ColumnWidth = 15
            Xarchexel.Cells(Xlin, XCol) = "IMPORTE"
            XCol = XCol + 1
            Xarchexel.Range("M" & Trim(str(Xlin))).ColumnWidth = 15
            Xarchexel.Cells(Xlin, XCol) = "IVA"
            XCol = XCol + 1
            Xarchexel.Range("N" & Trim(str(Xlin))).ColumnWidth = 15
            Xarchexel.Cells(Xlin, XCol) = "TOTAL"
            
            Xlin = Xlin + 1
            XCol = 1
            Do While Not Data1.Recordset.EOF
               If Check1.Value = 1 Then
                  Xarchexel.Cells(Xlin, XCol) = "'" & Format(Data1.Recordset("fecha"), "dd/mm/yyyy")
               Else
                  Xarchexel.Cells(Xlin, XCol) = "'" & Format(Data1.Recordset("realizada"), "dd/mm/yyyy")
               End If
               XCol = XCol + 1
               Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("ruc")
               XCol = XCol + 1
               Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("convenio")
               XCol = XCol + 1
               Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("nom_cli")
               XCol = XCol + 1
               Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("rub_cont")
               XCol = XCol + 1
               Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("factura")
               XCol = XCol + 1
               data_cab.RecordSource = "Select * from clirespl where cl_numero =" & Data1.Recordset("factura") & " and cl_codigo =" & Data1.Recordset("cod_cli")
               data_cab.Refresh
               If data_cab.Recordset.RecordCount > 0 Then
                  If IsNull(data_cab.Recordset("usu_baja")) = False Then
                     If data_cab.Recordset("usu_baja") = "USD" Then
                        Xarchexel.Cells(Xlin, XCol) = "U$s."
                        XCol = XCol + 1
                     Else
                        Xarchexel.Cells(Xlin, XCol) = "$."
                        XCol = XCol + 1
                     End If
                  Else
                     Xarchexel.Cells(Xlin, XCol) = "$."
                     XCol = XCol + 1
                  End If
               Else
                  Xarchexel.Cells(Xlin, XCol) = "$."
                  XCol = XCol + 1
               End If
               If IsNull(Data1.Recordset("pendiente")) = False Then
                  If Data1.Recordset("pendiente") = "F" Then
                     Xarchexel.Cells(Xlin, XCol) = "e-Factura"
                     XtotalFact = Data1.Recordset("tot_lin") + Data1.Recordset("valor_iva")
                     Xnrotipof = 9
                  Else
                     If Data1.Recordset("pendiente") = "T" Then
                        Xarchexel.Cells(Xlin, XCol) = "e-Ticket"
                        XtotalFact = Data1.Recordset("tot_lin") + Data1.Recordset("valor_iva")
                        Xnrotipof = 9
                     Else
                         If Data1.Recordset("pendiente") = "N" Then
                            Xarchexel.Cells(Xlin, XCol) = "NC e-Factura"
                            XtotalFact = Data1.Recordset("tot_lin") + Data1.Recordset("valor_iva")
                            Xnrotipof = 1
                         Else
                            If Data1.Recordset("pendiente") = "C" Then
                               Xarchexel.Cells(Xlin, XCol) = "NC e-Ticket"
                               XtotalFact = Data1.Recordset("tot_lin") + Data1.Recordset("valor_iva")
                               Xnrotipof = 1
                            Else
                               If Data1.Recordset("pendiente") = "A" Then
                                  Xarchexel.Cells(Xlin, XCol) = "ND e-Factura"
                                  XtotalFact = Data1.Recordset("tot_lin") + Data1.Recordset("valor_iva")
                                  Xnrotipof = 9
                               Else
                                  If Data1.Recordset("pendiente") = "B" Then
                                     Xarchexel.Cells(Xlin, XCol) = "ND e-Ticket"
                                     XtotalFact = Data1.Recordset("tot_lin") + Data1.Recordset("valor_iva")
                                     Xnrotipof = 9
                                  Else
                                     If Data1.Recordset("pendiente") = "X" Then
                                        Xarchexel.Cells(Xlin, XCol) = "Registro"
                                        XtotalFact = Data1.Recordset("tot_lin") + Data1.Recordset("valor_iva")
                                        Xnrotipof = 9
                                     Else
                                        If Data1.Recordset("pendiente") = "R" Then
                                           Xarchexel.Cells(Xlin, XCol) = "Registro"
                                           XtotalFact = Data1.Recordset("tot_lin") + Data1.Recordset("valor_iva")
                                           Xnrotipof = 9
                                        Else
                                           If Data1.Recordset("pendiente") = "Z" Then
                                              Xarchexel.Cells(Xlin, XCol) = "Registro"
                                              XtotalFact = Data1.Recordset("tot_lin") + Data1.Recordset("valor_iva")
                                              Xnrotipof = 9
                                           Else
                                              Xarchexel.Cells(Xlin, XCol) = "Registro"
                                              XtotalFact = Data1.Recordset("tot_lin") + Data1.Recordset("valor_iva")
                                              Xnrotipof = 9
                                           End If
                                        End If
                                     End If
                                  End If
                               End If
                            End If
                         End If
                     End If
                  End If
               Else
                  If Format(Data1.Recordset("fecha"), "yyyy/mm/dd") >= Format("01/07/2016", "yyyy/mm/dd") Then
                     'data_inf.Recordset("tipo") = "e-Factura " & Trim(Data1.Recordset("tipo"))
                     data_cab.RecordSource = "Select * from clirespl where cl_numero =" & Data1.Recordset("factura") & " and cl_codigo =" & Data1.Recordset("cod_cli")
                     data_cab.Refresh
                     If data_cab.Recordset.RecordCount > 0 Then
                        If data_cab.Recordset("cl_tipocli") = 111 Then
                           Xarchexel.Cells(Xlin, XCol) = "e-Factura"
                           XtotalFact = Data1.Recordset("tot_lin") + Data1.Recordset("valor_iva")
                           Xnrotipof = 9
                        Else
                           If data_cab.Recordset("cl_tipocli") = 101 Then
                              Xarchexel.Cells(Xlin, XCol) = "e-Ticket"
                              XtotalFact = Data1.Recordset("tot_lin") + Data1.Recordset("valor_iva")
                              Xnrotipof = 9
                           Else
                              If data_cab.Recordset("cl_tipocli") = 112 Then
                                 Xarchexel.Cells(Xlin, XCol) = "NC e-Factura"
                                 XtotalFact = Data1.Recordset("tot_lin") + Data1.Recordset("valor_iva")
                                 Xnrotipof = 1
                              Else
                                 If data_cab.Recordset("cl_tipocli") = 102 Then
                                    Xarchexel.Cells(Xlin, XCol) = "NC e-Ticket"
                                    XtotalFact = Data1.Recordset("tot_lin") + Data1.Recordset("valor_iva")
                                    Xnrotipof = 1
                                 Else
                                    If data_cab.Recordset("cl_tipocli") = 113 Then
                                       Xarchexel.Cells(Xlin, XCol) = "ND e-Factura"
                                       XtotalFact = Data1.Recordset("tot_lin") + Data1.Recordset("valor_iva")
                                       Xnrotipof = 9
                                    Else
                                       If data_cab.Recordset("cl_tipocli") = 103 Then
                                          Xarchexel.Cells(Xlin, XCol) = "ND e-Ticket"
                                          XtotalFact = Data1.Recordset("tot_lin") + Data1.Recordset("valor_iva")
                                          Xnrotipof = 9
                                       Else
                                          Xarchexel.Cells(Xlin, XCol) = "e-Factura"
                                          XtotalFact = Data1.Recordset("tot_lin") + Data1.Recordset("valor_iva")
                                          Xnrotipof = 9
                                       End If
                                    End If
                                 End If
                              End If
                           End If
                        End If
                     Else
                        XtotalFact = Data1.Recordset("tot_lin") + Data1.Recordset("valor_iva")
                        XtotalFact = "e-Factura"
                        Xnrotipof = 9
                     End If
                  Else
                     Xarchexel.Cells(Xlin, XCol) = Trim(Data1.Recordset("tipo"))
                     XtotalFact = Data1.Recordset("tot_lin")
                     Xnrotipof = 9
                  End If
               End If
               XCol = XCol + 1
               Xarchexel.Cells(Xlin, XCol) = Trim(Data1.Recordset("tipo"))
               XCol = XCol + 1
               If Check1.Value = 1 Then
                  Xarchexel.Cells(Xlin, XCol) = "'" & Format(Data1.Recordset("realizada"), "dd/mm/yyyy")
               Else
                  Xarchexel.Cells(Xlin, XCol) = "'" & Format(Data1.Recordset("fecha"), "dd/mm/yyyy")
               End If
               XCol = XCol + 1
               Adodc1.RecordSource = "Select * from indica_enfc where idhc =" & Data1.Recordset("factura") & " and in_dosis =" & 3
               Adodc1.Refresh
               If Adodc1.Recordset.RecordCount > 0 Then
                  Xarchexel.Cells(Xlin, XCol) = Adodc1.Recordset("in_obs")
               Else
                  Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("nom_prod")
               End If
               XCol = XCol + 1
               If Xnrotipof = 1 Then
                  Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("tot_lin") * -1
                  XCol = XCol + 1
                  Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("valor_iva") * -1
                  XCol = XCol + 1
                  Xarchexel.Cells(Xlin, XCol) = Format(XtotalFact, "Standard") * -1
               Else
                  Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("tot_lin")
                  XCol = XCol + 1
                  Xarchexel.Cells(Xlin, XCol) = Data1.Recordset("valor_iva")
                  XCol = XCol + 1
                  Xarchexel.Cells(Xlin, XCol) = Format(XtotalFact, "Standard")
               End If
               Data1.Recordset.MoveNext
               Xlin = Xlin + 1
               XCol = 1
            Loop
            Command1.Enabled = True
            Command2.Enabled = True
            frm_inffaccnv.MousePointer = 0
            MsgBox "Proceso terminado"
            Xlibexel.Save
            Xlibexel.Close
            Xobjexel.Quit
            Xlabrir.Workbooks.Open Xarchtex, , False
            Xlabrir.Visible = True
            Xlabrir.WindowState = xlMaximized

'            cr1.ReportFileName = App.Path & "\infvtasxcob.rpt"
'            cr1.ReportTitle = "INFORME DE FACTURAS REALIZADAS DESDE: " & mfd.Text & " HASTA: " & mfh.Text
'            cr1.Action = 1
         Else
            Xlibexel.Close
            Xobjexel.Quit
            Command1.Enabled = True
            Command2.Enabled = True
            frm_inffaccnv.MousePointer = 0
            MsgBox "No existen registros"
         End If
      End If
   
      If Option2.Value = True Then
         Data1.ConnectionString = "dsn=" & Xconexrmt
'         If t_cli.Text <> "" Then
'            Data1.RecordSource = "Select * from linmmdd where cod_cli =" & t_cli.Text & " and fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "'"
'            Data1.Refresh
'         Else
''            Data1.RecordSource = "Select * from linmmdd where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base in (101,102)"
'            Data1.Refresh
'         End If
         data_cnv.Connect = "odbc;dsn=" & Xconexrmt & ";"
         If Trim(t_cli.Text) <> "" Then
            data_cnv.RecordSource = "Select * from convenio where cnv_pmserv =" & 1 & " and cnv_fbaja is null and cnv_hasta >#" & Format(mfd.Text, "yyyy/mm/dd") & "# and cnv_cuenta =" & t_cli.Text
         Else
            data_cnv.RecordSource = "Select * from convenio where cnv_pmserv =" & 1 & " and cnv_fbaja is null and cnv_hasta >#" & Format(mfd.Text, "yyyy/mm/dd") & "#"
         End If
         data_cnv.Refresh
         If data_cnv.Recordset.RecordCount > 0 Then
            data_cnv.Recordset.MoveFirst
            Do While Not data_cnv.Recordset.EOF
               'data1.Recordset.FindFirst "cod_cli =" & data_cnv.Recordset("cnv_cuenta")
               Data1.RecordSource = "Select * from linmmdd where cod_cli =" & data_cnv.Recordset("cnv_cuenta") & " and fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and base in (101,102)"
               Data1.Refresh
               If Data1.Recordset.RecordCount > 0 Then
               Else
                  data_inf.Recordset.AddNew
                  data_inf.Recordset("fecha") = data_cnv.Recordset("cnv_desde")
                  data_inf.Recordset("cod_cli") = data_cnv.Recordset("cnv_cuenta")
                  data_inf.Recordset("nom_cli") = Mid(data_cnv.Recordset("cnv_desc"), 1, 30)
                  If IsNull(data_cnv.Recordset("cnv_paserv")) = False Then
                     data_inf.Recordset("base") = data_cnv.Recordset("cnv_paserv")
                  Else
                     data_inf.Recordset("base") = 0
                  End If
                  data_inf.Recordset("nom_prod") = data_cnv.Recordset("cnv_ruc")
                  data_inf.Recordset.Update
               End If
               data_cnv.Recordset.MoveNext
            Loop
            data_inf.RecordSource = "Select * from infvtas"
            data_inf.Refresh
            Command1.Enabled = True
            Command2.Enabled = True
            
            frm_inffaccnv.MousePointer = 0
            MsgBox "Proceso terminado"
            cr1.ReportFileName = App.path & "\infclifac.rpt"
            cr1.ReportTitle = "INFORME DE CLIENTES ACTIVOS CON FACTURAS PENDIENTES AL " & mfh.Text
            
            cr1.Action = 1
         Else
            frm_inffaccnv.MousePointer = 0
            MsgBox "No existen registros"
         End If
      End If
   
      If Option3.Value = True Then
         Data1.ConnectionString = "dsn=" & Xconexrmt
         Data1.RecordSource = "Select * from convenio where cnv_hasta >='" & Format(mfd.Text, "yyyy/mm/dd") & "' and cnv_pmserv =" & 1 & " and cnv_fbaja is null"
         Data1.Refresh
         If Data1.Recordset.RecordCount > 0 Then
            Data1.Recordset.MoveFirst
            Do While Not Data1.Recordset.EOF
               If IsNull(Data1.Recordset("cnv_pmserv")) = True Then
               Else
                  If Data1.Recordset("cnv_pmserv") = 1 Then
                     data_inf.Recordset.AddNew
                     data_inf.Recordset("fecha") = Data1.Recordset("cnv_desde")
                     data_inf.Recordset("cod_cli") = Data1.Recordset("cnv_cuenta")
                     data_inf.Recordset("nom_cli") = Mid(Data1.Recordset("cnv_desc"), 1, 30)
                     If IsNull(Data1.Recordset("cnv_paserv")) = False Then
                        data_inf.Recordset("base") = Data1.Recordset("cnv_paserv")
                     Else
                        data_inf.Recordset("base") = 0
                     End If
                     data_inf.Recordset("nom_prod") = Data1.Recordset("cnv_ruc")
                     data_inf.Recordset.Update
                  End If
               End If
               Data1.Recordset.MoveNext
            Loop
            data_inf.RecordSource = "Select * from infvtas"
            data_inf.Refresh
            Command1.Enabled = True
            Command2.Enabled = True
            
            frm_inffaccnv.MousePointer = 0
            MsgBox "Proceso terminado"
            cr1.ReportFileName = App.path & "\infclifac.rpt"
            cr1.ReportTitle = "INFORME DE CLIENTES ACTIVOS"
            cr1.Action = 1
         Else
            frm_inffaccnv.MousePointer = 0
            MsgBox "No existen registros"
         End If
      End If
   
   End If
End If
frm_inffaccnv.MousePointer = 0
      
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
frm_buscacnv.Show vbModal

End Sub

Private Sub Form_Load()
data_inf.DatabaseName = App.path & "\informes.mdb"
data_cab.Connect = "odbc;dsn=" & Xconexrmt & ";"
Adodc1.ConnectionString = "dsn=" & Xconexrmt

End Sub

Private Sub Form_Resize()
With Image1
     .Top = 0
     .Left = 0
     .Height = Me.Height
     .Width = Me.Width
End With

End Sub

Private Sub mfd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfh.SetFocus
End If

End Sub
