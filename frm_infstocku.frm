VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infstock 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de mantenimiento de stock"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6570
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_infstocku.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   6570
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   240
      TabIndex        =   15
      Top             =   5160
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin Crystal.CrystalReport cr2 
      Left            =   3000
      Top             =   4920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5760
      Picture         =   "frm_infstocku.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Salir"
      Top             =   5520
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_infstocku.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Procesar"
      Top             =   5520
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Opciones de informe"
      Height          =   4815
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6015
      Begin MSAdodcLib.Adodc data_item 
         Height          =   615
         Left            =   2760
         Top             =   2400
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
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
         Caption         =   "data_item"
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
      Begin VB.Data data_inf 
         Caption         =   "data_inf"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   2160
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2040
         Visible         =   0   'False
         Width           =   2655
      End
      Begin MSAdodcLib.Adodc data_sto 
         Height          =   375
         Left            =   3240
         Top             =   3600
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
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
         Caption         =   "data_sto"
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
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FF0000&
         Caption         =   "Ingreso de productos"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   4320
         Width           =   3855
      End
      Begin VB.TextBox T_CLI 
         Alignment       =   1  'Right Justify
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
         Left            =   3480
         TabIndex        =   13
         Top             =   1800
         Width           =   1575
      End
      Begin VB.TextBox t_cod 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3480
         TabIndex        =   10
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FF0000&
         Caption         =   "Productos por vencimiento."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3840
         Width           =   3855
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FF0000&
         Caption         =   "Total  de productos"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   3360
         Width           =   3855
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF0000&
         Caption         =   "Productos con stock en mínimo"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   2880
         Width           =   3855
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Gastos registrados."
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   2400
         Width           =   3855
      End
      Begin MSMask.MaskEdBox mfh 
         Height          =   375
         Left            =   4200
         TabIndex        =   3
         Top             =   480
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
         Left            =   2400
         TabIndex        =   2
         Top             =   480
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         BackColor       =   &H0000FFFF&
         Caption         =   "CODIGO DE CLIENTE (SOLO PARA OPCION GASTOS)"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1680
         Width           =   3375
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Código ITEM (vacío = TODOS)"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "Rango de fechas:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   1920
      Picture         =   "frm_infstocku.frx":0F56
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   2055
   End
End
Attribute VB_Name = "frm_infstock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frm_infstock.MousePointer = 11
Command2.Enabled = False
Command1.Enabled = False

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")
Dim Numdesde, Numale As Long
Dim Numhasta As Long

Numdesde = 100
Numhasta = 3400

Randomize
Numale = (Numdesde - Numhasta) * Rnd + Numhasta

Set RecInf = MiBaseact.OpenRecordset("Select * from infvtas")

MiBaseact.Execute "Delete * from infvtas"
data_inf.RecordSource = "infvtas"
data_inf.Refresh
pb.Visible = False
frm_infstock.MousePointer = 11

Dim Xobjexel As Excel.Application
Dim Xlibexel As Excel.Workbook
Dim Xarchexel As New Excel.Worksheet

Dim XCol, Xlin, Xnrocan, Xcolfija As Long
Dim Xarchtex As String
Dim Xlabrir As New Excel.Application
XCol = 1
Xlin = 1
Xnrocan = 1

If Option1.Value = True Then
   If mfd.Text <> "__/__/____" And mfh.Text <> "__/__/____" Then
      Set Xobjexel = New Excel.Application
      Set Xlibexel = Xobjexel.Workbooks.Add
      Set Xarchexel = Xlibexel.Worksheets.Add
      Xarchexel.Name = "ECONOMATO"
      Xlibexel.SaveAs ("C:\planillas\" & "Infeco" & Trim(str(Numale)) & ".xls")
      Xarchtex = "C:\planillas\" & "Infeco" & Trim(str(Numale)) & ".xls"
      
      If t_cod.Text <> "" Then
         If WElusuario = "JFERNAN" Or WElusuario = "COMPUTOS" Then
            data_sto.RecordSource = "select * from gastos where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and codprod =" & t_cod.Text & " order by codprod"
         Else
            data_sto.RecordSource = "select * from gastos where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and codprod =" & t_cod.Text & " order by codprod"
         End If
         data_sto.Refresh
      Else
         If t_cli.Text <> "" Then
            If WElusuario = "JFERNAN" Or WElusuario = "COMPUTOS" Then
               data_sto.RecordSource = "select gastos.fecha,gastos.codprod,gastos.descrip,gastos.codcli,gastos.nomcli,gastos.cant,stock.preuni,stock.grupo from gastos inner join stock on gastos.codprod=stock.id where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and codcli =" & t_cli.Text & " and stock.grupo in (3) order by codprod"
            Else
               data_sto.RecordSource = "select gastos.fecha,gastos.codprod,gastos.descrip,gastos.codcli,gastos.nomcli,gastos.cant,stock.preuni,stock.grupo from gastos inner join stock on gastos.codprod=stock.id where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and codcli =" & t_cli.Text & " and stock.grupo not in (3) order by codprod"
            End If
            data_sto.Refresh
         Else
            If WElusuario = "JFERNAN" Or WElusuario = "COMPUTOS" Then
               data_sto.RecordSource = "select gastos.fecha,gastos.codprod,gastos.descrip,gastos.codcli,gastos.nomcli,gastos.cant,gastos.prec,stock.preuni," & _
               "stock.grupo from gastos inner join stock on gastos.codprod=stock.id where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and stock.grupo in (3)"
            Else
               data_sto.RecordSource = "select gastos.fecha,gastos.codprod,gastos.descrip,gastos.codcli,gastos.nomcli,gastos.cant,stock.preuni,stock.grupo from gastos inner join stock on gastos.codprod=stock.id where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and stock.grupo not in (3)"
            End If
            data_sto.Refresh
         End If
      End If
      
      pb.Visible = True
      If data_sto.Recordset.RecordCount > 0 Then
         data_sto.Recordset.MoveLast
         pb.Max = data_sto.Recordset.RecordCount + 2
         pb.Value = 0
         data_sto.Recordset.MoveFirst
         Xarchexel.Cells(Xlin, XCol) = "SAPP - ECONOMATO"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         If Option1.Value = True Then
            Xarchexel.Cells(Xlin, XCol) = "INFORME " & Option1.Caption & " DESDE: " & mfd.Text & " HASTA: " & mfh.Text
         Else
            If Option2.Value = True Then
               Xarchexel.Cells(Xlin, XCol) = "INFORME ECONOMATO" & " DESDE: " & mfd.Text & " HASTA: " & mfh.Text
            End If
         End If
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "AD" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "AD" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "AD" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "AD" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "AD" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "AD" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "AD" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FECHA"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "COD.PROD"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "DESCRIPCION"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 5
         Xarchexel.Cells(Xlin, XCol) = "CLIENTE"
         XCol = XCol + 1
         Xarchexel.Cells(Xlin, XCol) = "NOMBRE CLIENTE"
         XCol = XCol + 1
         Xarchexel.Cells(Xlin, XCol) = "CANTIDAD"
         XCol = XCol + 1
         Xarchexel.Cells(Xlin, XCol) = "PRECIO U."
         
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_sto.Recordset.EOF
            Xarchexel.Cells(Xlin, XCol) = "'" & Format(data_sto.Recordset("fecha"), "dd/mm/yyyy")
            XCol = XCol + 1
            Xarchexel.Cells(Xlin, XCol) = data_sto.Recordset("codprod")
            XCol = XCol + 1
            Xarchexel.Cells(Xlin, XCol) = data_sto.Recordset("descrip")
            XCol = XCol + 1
            Xarchexel.Cells(Xlin, XCol) = data_sto.Recordset("codcli")
            XCol = XCol + 1
            Xarchexel.Cells(Xlin, XCol) = data_sto.Recordset("nomcli")
            XCol = XCol + 1
            Xarchexel.Cells(Xlin, XCol) = data_sto.Recordset("cant")
            XCol = XCol + 1
            
            If t_cod.Text <> "" Then
               data_item.RecordSource = "Select * from stock where id =" & data_sto.Recordset("codprod")
               data_item.Refresh
               If data_item.Recordset.RecordCount > 0 Then
                  Xarchexel.Cells(Xlin, XCol) = data_item.Recordset("preuni")
                  XCol = XCol + 1
'                  data_inf.Recordset.Edit
'                  data_inf.Recordset("tot_lin") = data_item.Recordset("preuni")
'                  data_inf.Recordset("costo_prod") = data_item.Recordset("preuni") * data_inf.Recordset("grupo")
'                  data_inf.Recordset.Update
               End If
            Else
               If t_cli.Text <> "" Then
                   data_item.RecordSource = "Select * from stock where id =" & data_sto.Recordset("codprod")
                   data_item.Refresh
                   If data_item.Recordset.RecordCount > 0 Then
                      Xarchexel.Cells(Xlin, XCol) = data_item.Recordset("preuni")
                      XCol = XCol + 1
                   End If
               Else
                    Xarchexel.Cells(Xlin, XCol) = data_sto.Recordset("preuni")
                    XCol = XCol + 1
               End If
            End If
            
            pb.Value = pb.Value + 1
            Xlin = Xlin + 1
            XCol = 1
            data_sto.Recordset.MoveNext
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
         
         data_inf.Refresh
         data_inf.RecordSource = "Select * from infvtas"
         data_inf.Refresh
         frm_infstock.MousePointer = 0
         MsgBox "Proceso terminado"
         data_inf.Refresh
'         cr2.ReportTitle = "INFORME DE GASTOS DESDE " & mfd.Text & " HASTA " & mfh.Text
'         cr2.ReportFileName = App.Path & "\infstog.rpt"
'         cr2.Action = 1
         pb.Visible = False
      Else
         frm_infstock.MousePointer = 0
          MsgBox "No existen registros"
          Xobjexel.Quit
      End If
   Else
         frm_infstock.MousePointer = 0
      MsgBox "No ingresó fechas", vbCritical
   End If
End If
If Option5.Value = True Then
   If mfd.Text <> "__/__/____" And mfh.Text <> "__/__/____" Then
      If t_cod.Text <> "" Then
         If WElusuario = "JFERNAN" Or WElusuario = "COMPUTOS" Then
            data_sto.RecordSource = "select * from lineascomp where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and codprod =" & t_cod.Text & " and grupo in (3) order by fecha"
         Else
            data_sto.RecordSource = "select * from lineascomp where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and codprod =" & t_cod.Text & " and grupo not in (3) order by fecha"
         End If
         data_sto.Refresh
      Else
         If WElusuario = "JFERNAN" Or WElusuario = "COMPUTOS" Then
            data_sto.RecordSource = "select * from lineascomp where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and grupo in (3) order by fecha"
         Else
            data_sto.RecordSource = "select * from lineascomp where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and grupo not in (3) order by fecha"
         End If
         data_sto.Refresh
      End If
      If data_sto.Recordset.RecordCount > 0 Then
         data_sto.Recordset.MoveFirst
         Do While Not data_sto.Recordset.EOF
            data_inf.Recordset.AddNew
            data_inf.Recordset("fecha") = data_sto.Recordset("fecha")
            data_inf.Recordset("cod_prod") = data_sto.Recordset("codprod")
            data_inf.Recordset("nom_prod") = Mid(data_sto.Recordset("coddesc"), 1, 50)
            data_inf.Recordset("tot_lin") = data_sto.Recordset("totprod")
            data_inf.Recordset("grupo") = data_sto.Recordset("cant")
            data_inf.Recordset.Update
            data_sto.Recordset.MoveNext
         Loop
         frm_infstock.MousePointer = 0
         MsgBox "Proceso terminado"
         cr2.ReportTitle = "INFORME DE COMPRAS DESDE " & mfd.Text & " HASTA " & mfh.Text
         cr2.ReportFileName = App.path & "\infstog.rpt"
         cr2.Action = 1
      Else
         frm_infstock.MousePointer = 0
          MsgBox "No existen registros"
      End If
   Else
         frm_infstock.MousePointer = 0
      MsgBox "No ingresó fechas", vbCritical
   End If
End If

If Option2.Value = True Then
   If t_cod.Text <> "" Then
      If WElusuario = "JFERNAN" Or WElusuario = "COMPUTOS" Then
         data_sto.RecordSource = "select * from stock where id =" & t_cod.Text & " and grupo in (3) order by descrip"
      Else
         data_sto.RecordSource = "select * from stock where id =" & t_cod.Text & " and grupo not in (3) order by descrip"
      End If
      data_sto.Refresh
   Else
      If WElusuario = "JFERNAN" Or WElusuario = "COMPUTOS" Then
         data_sto.RecordSource = "select * from stock where grupo in (3) order by descrip"
      Else
         data_sto.RecordSource = "select * from stock where grupo not in (3) order by descrip"
      End If
      data_sto.Refresh
   End If
   If data_sto.Recordset.RecordCount > 0 Then
      data_sto.Recordset.MoveFirst
      Do While Not data_sto.Recordset.EOF
         If data_sto.Recordset("actual") <= data_sto.Recordset("minimo") Then
            data_inf.Recordset.AddNew
            data_inf.Recordset("fecha") = data_sto.Recordset("ultact")
            data_inf.Recordset("cod_prod") = data_sto.Recordset("id")
            data_inf.Recordset("nom_prod") = Mid(data_sto.Recordset("descrip"), 1, 50)
            data_inf.Recordset("cod_cli") = data_sto.Recordset("actual")
            data_inf.Recordset("tot_lin") = data_sto.Recordset("minimo")
            data_inf.Recordset("factura") = data_sto.Recordset("basico")
            data_inf.Recordset("base") = data_sto.Recordset("basico") - data_sto.Recordset("actual")
            data_inf.Recordset.Update
         End If
         data_sto.Recordset.MoveNext
      Loop
         frm_infstock.MousePointer = 0
      MsgBox "Proceso terminado"
      cr2.ReportTitle = "INFORME DE STOCK EN MINIMO"
      cr2.ReportFileName = App.path & "\infstomin.rpt"
      cr2.Action = 1
   Else
         frm_infstock.MousePointer = 0
      MsgBox "No existen registros", vbCritical
   End If
End If

If Option3.Value = True Then
   If t_cod.Text <> "" Then
      data_sto.RecordSource = "select * from stock where id =" & t_cod.Text & " order by descrip,grupo"
      data_sto.Refresh
   Else
      If WElusuario = "JFERNAN" Or WElusuario = "COMPUTOS" Then
         data_sto.RecordSource = "select * from stock where grupo in (3) order by descrip,grupo"
      Else
         data_sto.RecordSource = "select * from stock where grupo not in (3) order by descrip,grupo"
      End If
      data_sto.Refresh
   End If
   If data_sto.Recordset.RecordCount > 0 Then
      data_sto.Recordset.MoveFirst
      Do While Not data_sto.Recordset.EOF
         data_inf.Recordset.AddNew
         data_inf.Recordset("fecha") = data_sto.Recordset("ultact")
         data_inf.Recordset("cod_prod") = data_sto.Recordset("id")
         data_inf.Recordset("nom_prod") = Mid(data_sto.Recordset("descrip"), 1, 50)
         data_inf.Recordset("cod_cli") = data_sto.Recordset("actual")
         data_inf.Recordset("tot_lin") = data_sto.Recordset("minimo")
         data_inf.Recordset("factura") = data_sto.Recordset("basico")
         data_inf.Recordset("base") = data_sto.Recordset("grupo")
         data_inf.Recordset("imp_iva") = data_sto.Recordset("preuni")
         data_inf.Recordset("nom_medic") = Mid(data_sto.Recordset("obs"), 1, 50)
         data_inf.Recordset.Update
         data_sto.Recordset.MoveNext
      Loop
         frm_infstock.MousePointer = 0
      MsgBox "Proceso terminado"
      cr2.ReportTitle = "INFORME DE TOTAL DE ITEMS POR GRUPOS"
      cr2.ReportFileName = App.path & "\infstotot.rpt"
      cr2.Action = 1
   Else
         frm_infstock.MousePointer = 0
      MsgBox "No existen registros", vbCritical
   End If
End If

If Option4.Value = True Then
   If t_cod.Text <> "" Then
      data_sto.RecordSource = "select * from lineascomp where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and codprod =" & t_cod.Text
      data_sto.Refresh
   Else
      If WElusuario = "JFERNAN" Or WElusuario = "COMPUTOS" Then
         data_sto.RecordSource = "select * from lineascomp where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and grupo in (3)"
      Else
         data_sto.RecordSource = "select * from lineascomp where fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and grupo not in (3)"
      End If
      data_sto.Refresh
   End If
   If data_sto.Recordset.RecordCount > 0 Then
      data_sto.Recordset.MoveFirst
      Do While Not data_sto.Recordset.EOF
         data_inf.Recordset.AddNew
         data_inf.Recordset("fecha") = data_sto.Recordset("fecha")
         data_inf.Recordset("cod_prod") = data_sto.Recordset("codprod")
         data_inf.Recordset("nom_prod") = Mid(data_sto.Recordset("coddesc"), 1, 50)
         data_inf.Recordset("cod_cli") = data_sto.Recordset("cant")
         data_inf.Recordset("base") = data_sto.Recordset("precuni")
         data_inf.Recordset.Update
         data_sto.Recordset.MoveNext
      Loop
         frm_infstock.MousePointer = 0
      MsgBox "Proceso terminado"
      cr2.ReportTitle = "INFORME DE ITEMS CON VENCIMIENTO"
      cr2.ReportFileName = App.path & "\infstoven.rpt"
      cr2.Action = 1
   Else
         frm_infstock.MousePointer = 0
      MsgBox "No existen registros", vbCritical
   End If
End If
frm_infstock.MousePointer = 0
Command1.Enabled = True
Command2.Enabled = True

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
'data_sto.DatabaseName = App.Path & "\" & Trim(Xlabdd)
'data_sto.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_sto.ConnectionString = "dsn=" & Xconexrmt
data_inf.DatabaseName = App.path & "\informes.mdb"
'data_inf.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"

'data_inf.RecordSource = "infvtas"
'data_inf.Refresh
data_item.ConnectionString = "dsn=" & Xconexrmt


End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub mfd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfh.SetFocus
End If

End Sub

Private Sub mfh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Option1.SetFocus
End If

End Sub

