VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infaccadm 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5385
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_infaccadm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   5385
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_llam 
      Caption         =   "data_llam"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2640
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc data_cli 
      Height          =   330
      Left            =   2160
      Top             =   2880
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
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
      Caption         =   "data_cli"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   1920
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3480
      Picture         =   "frm_infaccadm.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   960
      Picture         =   "frm_infaccadm.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Procesar"
      Top             =   2640
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Datos para informe"
      Height          =   2295
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   4815
      Begin MSAdodcLib.Adodc data1 
         Height          =   375
         Left            =   720
         Top             =   720
         Visible         =   0   'False
         Width           =   2175
         _ExtentX        =   3836
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSAdodcLib.Adodc data_lin 
         Height          =   375
         Left            =   2400
         Top             =   1680
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
         Caption         =   "data_lin"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FF8080&
         Caption         =   "Todos los registros"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1920
         Value           =   1  'Checked
         Width           =   3015
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "frm_infaccadm.frx":109E
         Left            =   1440
         List            =   "frm_infaccadm.frx":10B4
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1200
         Width           =   3135
      End
      Begin MSMask.MaskEdBox mfh 
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfd 
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
         Caption         =   "Opciones:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "Fechas:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   120
      Picture         =   "frm_infaccadm.frx":111F
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1335
   End
End
Attribute VB_Name = "frm_infaccadm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 0 Then
   MsgBox "Recuerde que ésta opción sólo se puede seleccionar si tiene abierto el módulo de ingresar acciones", vbExclamation

End If

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command1.SetFocus
End If

End Sub

Private Sub Command1_Click()

frm_infaccadm.MousePointer = 11

Command1.Enabled = False

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")
Dim XbaseCGal As Integer
Dim Lamatric As Long
Lamatric = 0
XbaseCGal = 0

MiBaseact.Execute "Delete * from infcli"

data_inf.RecordSource = "infcli"
data_inf.Refresh
Dim Xlamatacc As Long
Dim Xcantacc As Integer
Xlamatacc = 0
Xcantacc = 0
Dim Xlafecdesde As Date
Dim Xlafechasta As Date

Dim Xobjexelcar As Excel.Application
Dim Xlibexelcar As Excel.Workbook
Dim Xarchexelcar As New Excel.Worksheet
Dim Xlabrir3 As New Excel.Application
Dim BuscaCedDesp As String
BuscaCedDesp = ""
Dim XCol, Xlin, Xnrocan, Xcolfija As Long
Dim Xarchtex As String

If mfd.Text <> "__/__/____" And mfh.Text <> "__/__/____" Then
   If Combo1.ListIndex = 0 Then
      If Check1.Value = 1 Then
         Data1.RecordSource = "Select * from mant_sol where cl_fultpag >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fultpag <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and estado >" & 0 & " order by cl_fultpag"
         Data1.Refresh
      Else
         Data1.RecordSource = "Select * from mant_sol where cl_fultpag >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fultpag <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cl_zona =" & Wopsed & " and estado >" & 0 & " order by cl_fultpag"
         Data1.Refresh
      End If
   Else
      If Combo1.ListIndex = 1 Then 'Posibles pagos
         If Check1.Value = 1 Then
            Data1.RecordSource = "Select * from mant_sol where cl_fec1 >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fec1 <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and estado >" & 0 & " order by cl_fec1"
            Data1.Refresh
         Else
            Data1.RecordSource = "Select * from mant_sol where cl_fec1 >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fec1 <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cl_zona =" & Wopsed & " and estado >" & 0 & " order by cl_fec1"
            Data1.Refresh
         End If
      Else
         If Combo1.ListIndex = 2 Then 'Vtas sin cobrar
            If Check1.Value = 1 Then
'               Data1.RecordSource = "Select * from mant_sol where cl_val3 >=" & 0 & " and cl_fultpag >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fec1 <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and estado >" & 0 & " order by cl_fultpag"
               Data1.RecordSource = "Select deudas.fecha,deudas.cliente,deudas.nombre,deudas.importe,deudas.documento,deudas.tipodoc,deudas.fecha_pago,deudas.origen,mant_sol.cl_nro_sup,mant_sol.cl_ruc," & _
               "mant_sol.cl_atrasop,mant_sol.cl_descpag,mant_sol.cl_email,mant_sol.info_debit,mant_sol.cl_fec1,mant_sol.cl_desc2,mant_sol.cl_desc1" & _
               " from deudas left outer join mant_sol on deudas.documento=mant_sol.cl_nro_sup where tipodoc ='" & "CRE" & "' and fecha_pago is null and origen <='" & "Refinanc" & "' and fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' order by cliente"
               Data1.Refresh
            Else
               If Combo1.ListIndex = 3 Then
                  Data1.RecordSource = "Select Codigos_aut.fecha,Codigos_aut.usuario,Codigos_aut.codaut,Codigos_aut.modulo,Codigos_aut.usuario_caja,Codigos_aut.socio," & _
                  "clientes.cl_codigo,clientes.cl_apellid,clientes.cl_dpto" & _
                  " from Codigos_aut left outer join clientes on Codigos_aut.socio=clientes.cl_codigo where Codigos_aut.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and Codigos_aut.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "'"
                  Data1.Refresh
               Else
                   Data1.RecordSource = "Select deudas.fecha,deudas.cliente,deudas.nombre,deudas.documento,deudas.tipodoc,deudas.fecha_pago,deudas.origen,mant_sol.cl_nro_sup,mant_sol.cl_ruc," & _
                   "mant_sol.cl_atrasop,mant_sol.cl_descpag,mant_sol.cl_email,mant_sol.info_debit,mant_sol.cl_fec1,mant_sol.cl_desc2,mant_sol.cl_desc1" & _
                   " from deudas left outer join mant_sol on deudas.documento=mant_sol.cl_nro_sup where tipodoc ='" & "CRE" & "' and fecha_pago is null and origen <='" & "Refinanc" & "' and fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cl_zona =" & Wopsed & " order by cliente,cl_nro_sup"
    
    '               Data1.RecordSource = "deudas.fecha,deudas.cliente,deudas.nombre,deudas.documento,deudas.tipodoc,deudas.fecha_pago,deudas.origen,mant_sol.cl_ruc," & _
    '               "mant_sol.cl_atrasop,mant_sol.cl_descpag,mant_sol.email,mant_sol.info_debit,mant_sol.cl_fec1,mant_sol.desc2,mant_sol.desc1" & _
    '               " from deudas inner join mant_sol on deudas.documento=mant_sol.cl_nro_sup where tipodoc ='" & "CRE" & "' and fecha_pago is null and origen <='" & "Refinanc" & "' and fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cl_zona =" & Wopsed & " order by cliente,fecha"
    '               Data1.RecordSource = "Select * from deudas where tipodoc ='" & "CRE" & "' and fecha_pago is null and origen <='" & "Refinan" & "' and cliente =" & Val(cl_zona) & " order by fecha"
    '               Data1.RecordSource = "Select * from mant_sol where cl_val3 >=" & 0 & " and cl_fultpag >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fec1 <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cl_zona =" & Wopsed & " and estado >" & 0 & " order by cl_fultpag"
                   Data1.Refresh
               End If
            End If
         Else
            If Combo1.ListIndex = 3 Then
               Data1.RecordSource = "Select Codigos_aut.fecha,Codigos_aut.usuario,Codigos_aut.codaut,Codigos_aut.modulo,Codigos_aut.usuario_caja,Codigos_aut.socio," & _
               "clientes.cl_codigo,clientes.cl_apellid,clientes.cl_dpto" & _
               " from Codigos_aut left outer join clientes on Codigos_aut.socio=clientes.cl_codigo where Codigos_aut.fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and Codigos_aut.fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "'"
               Data1.Refresh
            Else
               If Combo1.ListIndex = 4 Then
                  Data1.RecordSource = "select Codigos_aut.fecha,Codigos_aut.usuario,Codigos_aut.codaut,Codigos_aut.socio,Codigos_aut.modulo,Codigos_aut.usuario_caja,Codigos_aut.mes_anio," & _
                  "Codigos_aut.tipo_deuda,Codigos_aut.cuando,Codigos_aut.contacto,Codigos_aut.observa,clientes.cl_codigo,clientes.cl_apellid,clientes.cl_codconv " & _
                  "from Codigos_aut inner join clientes on Codigos_aut.socio=clientes.cl_codigo where Codigos_aut.fecha >='" & Format(mfd.Text, "yyyy/mm/dd") & "' and Codigos_aut.fecha <='" & Format(mfh.Text, "yyyy/mm/dd") & "'"
                  Data1.Refresh
               Else
                  If Combo1.ListIndex = 5 Then
                     Data1.RecordSource = "select Codigos_aut.fecha,Codigos_aut.usuario,Codigos_aut.codaut,Codigos_aut.socio,Codigos_aut.modulo,Codigos_aut.usuario_caja,Codigos_aut.mes_anio,Codigos_aut.base," & _
                     "Codigos_aut.tipo_deuda,Codigos_aut.cuando,Codigos_aut.contacto,Codigos_aut.observa,clientes.cl_codigo,clientes.cl_apellid,clientes.cl_codconv " & _
                     "from Codigos_aut inner join clientes on Codigos_aut.socio=clientes.cl_codigo where Codigos_aut.fecha >='" & Format(mfd.Text, "yyyy/mm/dd") & "' and Codigos_aut.fecha <='" & Format(mfh.Text, "yyyy/mm/dd") & "' and Codigos_aut.codaut in ('C.GALICIA','URGENCIA CGAL') order by Codigos_aut.socio"
                     Data1.Refresh
                  Else
                     Data1.RecordSource = "Select deudas.fecha,deudas.cliente,deudas.nombre,deudas.documento,deudas.tipodoc,deudas.fecha_pago,deudas.origen,mant_sol.cl_nro_sup,mant_sol.cl_ruc," & _
                     "mant_sol.cl_atrasop,mant_sol.cl_descpag,mant_sol.cl_email,mant_sol.info_debit,mant_sol.cl_fec1,mant_sol.cl_desc2,mant_sol.cl_desc1" & _
                     " from deudas left outer join mant_sol on deudas.documento=mant_sol.cl_nro_sup where tipodoc ='" & "CRE" & "' and fecha_pago is null and origen <='" & "Refinanc" & "' and fecha >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cl_zona =" & Wopsed & " order by cliente,cl_nro_sup"
                     Data1.Refresh
                  End If
               End If
            End If
         End If
      End If
   End If
   
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      If Combo1.ListIndex = 4 Or Combo1.ListIndex = 5 Then
         If Combo1.ListIndex = 4 Then
              Data1.Recordset.MoveLast
              Xnrocan = Data1.Recordset.RecordCount + 5
              Data1.Recordset.MoveFirst
              Set Xobjexelcar = New Excel.Application
              Set Xlibexelcar = Xobjexelcar.Workbooks.Add
              Set Xarchexelcar = Xlibexelcar.Worksheets.Add
              Xlin = 1
              XCol = 1
              Xarchexelcar.Name = "autoriza"
              Xlibexelcar.SaveAs ("C:\planillas\autoriza.xls")
              Xarchtex = "C:\planillas\autoriza.xls"
              Xarchexelcar.Cells(Xlin, XCol) = "SAPP S.A.  -- DPTO.TI"
              Xlin = Xlin + 1
              XCol = 1
              Xarchexelcar.Range("A2", "C3").Font.Size = 16
              Xarchexelcar.Cells(Xlin, XCol) = "Informe de autorizaciones automáticas FECHAS:" & mfd.Text & " " & mfh.Text
              Xlin = Xlin + 1
              XCol = 1
              Xarchexelcar.Cells(Xlin, XCol) = "FECHA ACTUAL: " & Format(Date, "dd/mm/yyyy")
              XCol = 1
              Xlin = Xlin + 2
              Xarchexelcar.Range("A4", "O" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
              Xarchexelcar.Range("A4", "O" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
              Xarchexelcar.Range("A4", "O" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
              Xarchexelcar.Range("A4", "O" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
              Xarchexelcar.Range("A4", "O" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
              Xarchexelcar.Range("A4", "O" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
              Xarchexelcar.Range("A" & Trim(str(Xlin)), "O" & Trim(str(Xlin))).Interior.color = RGB(0, 160, 0)
              Xarchexelcar.Range("A" & Trim(str(Xlin))).ColumnWidth = 15
             
              Xarchexelcar.Cells(Xlin, XCol) = "FECHA"
              XCol = XCol + 1
              Xarchexelcar.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
              Xarchexelcar.Cells(Xlin, XCol) = "USUARIO"
              XCol = XCol + 1
              Xarchexelcar.Range("C" & Trim(str(Xlin))).ColumnWidth = 15
              Xarchexelcar.Cells(Xlin, XCol) = "MATRICULA"
              XCol = XCol + 1
              Xarchexelcar.Range("D" & Trim(str(Xlin))).ColumnWidth = 40
              Xarchexelcar.Cells(Xlin, XCol) = "NOMBRE"
              XCol = XCol + 1
              Xarchexelcar.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
              Xarchexelcar.Cells(Xlin, XCol) = "CONVENIO"
              XCol = XCol + 1
              Xarchexelcar.Range("F" & Trim(str(Xlin))).ColumnWidth = 15
              Xarchexelcar.Cells(Xlin, XCol) = "MODULO"
              XCol = XCol + 1
              Xarchexelcar.Range("G" & Trim(str(Xlin))).ColumnWidth = 15
              Xarchexelcar.Cells(Xlin, XCol) = "AUTORIZACION"
              XCol = XCol + 1
              Xarchexelcar.Range("H" & Trim(str(Xlin))).ColumnWidth = 15
              Xarchexelcar.Cells(Xlin, XCol) = "TIPO_DEUDA"
              XCol = XCol + 1
              Xarchexelcar.Range("I" & Trim(str(Xlin))).ColumnWidth = 15
              Xarchexelcar.Cells(Xlin, XCol) = "CUANDO PAGA"
              XCol = XCol + 1
              Xarchexelcar.Range("J" & Trim(str(Xlin))).ColumnWidth = 40
              Xarchexelcar.Cells(Xlin, XCol) = "CONTACTO"
              XCol = XCol + 1
              Xarchexelcar.Range("K" & Trim(str(Xlin))).ColumnWidth = 50
              Xarchexelcar.Cells(Xlin, XCol) = "OBSERVACIONES"
              XCol = XCol + 1
              Xarchexelcar.Range("L" & Trim(str(Xlin))).ColumnWidth = 10
              Xarchexelcar.Cells(Xlin, XCol) = "COSTO"
              XCol = XCol + 1
              Xarchexelcar.Range("M" & Trim(str(Xlin))).ColumnWidth = 10
              Xarchexelcar.Cells(Xlin, XCol) = "BOLETA"
              XCol = XCol + 1
              Xarchexelcar.Range("N" & Trim(str(Xlin))).ColumnWidth = 12
              Xarchexelcar.Cells(Xlin, XCol) = "TIPO_FACT"
                
              Xlin = Xlin + 1
              XCol = 1
              frm_infaccadm.MousePointer = 11
              Do While Not Data1.Recordset.EOF
                 If IsNull(Data1.Recordset("fecha")) = False Then
                    Xarchexelcar.Cells(Xlin, XCol) = "'" & Format(Data1.Recordset("fecha"), "dd/mm/yyyy")
                 Else
                    Xarchexelcar.Cells(Xlin, XCol) = "Sin fecha"
                 End If
                 XCol = XCol + 1
                 If IsNull(Data1.Recordset("usuario")) = False Then
                    Xarchexelcar.Cells(Xlin, XCol) = Data1.Recordset("usuario")
                 Else
                    Xarchexelcar.Cells(Xlin, XCol) = "s/d"
                 End If
                 XCol = XCol + 1
                 If IsNull(Data1.Recordset("socio")) = False Then
                    Xarchexelcar.Cells(Xlin, XCol) = Data1.Recordset("socio")
                 Else
                    Xarchexelcar.Cells(Xlin, XCol) = 0
                 End If
                 XCol = XCol + 1
                 If IsNull(Data1.Recordset("cl_apellid")) = False Then
                    Xarchexelcar.Cells(Xlin, XCol) = Data1.Recordset("cl_apellid")
                 Else
                    Xarchexelcar.Cells(Xlin, XCol) = "NN"
                 End If
                 XCol = XCol + 1
                 If IsNull(Data1.Recordset("cl_codconv")) = False Then
                    Xarchexelcar.Cells(Xlin, XCol) = Data1.Recordset("cl_codconv")
                 Else
                    Xarchexelcar.Cells(Xlin, XCol) = "NN"
                 End If
                 XCol = XCol + 1
                 If IsNull(Data1.Recordset("modulo")) = False Then
                    Xarchexelcar.Cells(Xlin, XCol) = Data1.Recordset("modulo")
                 Else
                    Xarchexelcar.Cells(Xlin, XCol) = "s/d"
                 End If
                 XCol = XCol + 1
                 If IsNull(Data1.Recordset("codaut")) = False Then
                    Xarchexelcar.Cells(Xlin, XCol) = Data1.Recordset("codaut")
                 Else
                    Xarchexelcar.Cells(Xlin, XCol) = "s/d"
                 End If
                 XCol = XCol + 1
                 If IsNull(Data1.Recordset("tipo_deuda")) = False Then
                    Xarchexelcar.Cells(Xlin, XCol) = Data1.Recordset("tipo_deuda")
                 Else
                    Xarchexelcar.Cells(Xlin, XCol) = "s/d"
                 End If
                 XCol = XCol + 1
                 If IsNull(Data1.Recordset("cuando")) = False Then
                    Xarchexelcar.Cells(Xlin, XCol) = Data1.Recordset("cuando")
                 Else
                    Xarchexelcar.Cells(Xlin, XCol) = "s/d"
                 End If
                 XCol = XCol + 1
                 If IsNull(Data1.Recordset("contacto")) = False Then
                    Xarchexelcar.Cells(Xlin, XCol) = Data1.Recordset("contacto")
                 Else
                    Xarchexelcar.Cells(Xlin, XCol) = "sin datos"
                 End If
                 XCol = XCol + 1
                 If IsNull(Data1.Recordset("observa")) = False Then
                    Xarchexelcar.Cells(Xlin, XCol) = Data1.Recordset("observa")
                 Else
                    Xarchexelcar.Cells(Xlin, XCol) = "sin datos"
                 End If
              
                 Data1.Recordset.MoveNext
                 Xlin = Xlin + 1
                 XCol = 1
              Loop
              frm_infaccadm.MousePointer = 0
            
              Xlibexelcar.Save
              Xlibexelcar.Close
              Xobjexelcar.Quit
              MsgBox "El archivo autoriza.xls ha sido guardado en la carpeta PLANILLAS del disco C", vbInformation
              Xlabrir3.Workbooks.Open Xarchtex, , False
              Xlabrir3.Visible = True
              Xlabrir3.WindowState = xlMaximized
         Else
              Data1.Recordset.MoveLast
              Xnrocan = Data1.Recordset.RecordCount + 5
              Data1.Recordset.MoveFirst
              Set Xobjexelcar = New Excel.Application
              Set Xlibexelcar = Xobjexelcar.Workbooks.Add
              Set Xarchexelcar = Xlibexelcar.Worksheets.Add
              Xlin = 1
              XCol = 1
              Xarchexelcar.Name = "autoriza"
              Xlibexelcar.SaveAs ("C:\planillas\autoriza.xls")
              Xarchtex = "C:\planillas\autoriza.xls"
              Xarchexelcar.Cells(Xlin, XCol) = "SAPP S.A.  -- DPTO.TI"
              Xlin = Xlin + 1
              XCol = 1
              Xarchexelcar.Range("A2", "C3").Font.Size = 16
              Xarchexelcar.Cells(Xlin, XCol) = "Informe de autorizaciones C.GALICIA FECHAS:" & mfd.Text & " " & mfh.Text
              Xlin = Xlin + 1
              XCol = 1
              Xarchexelcar.Cells(Xlin, XCol) = "FECHA ACTUAL: " & Format(Date, "dd/mm/yyyy")
              XCol = 1
              Xlin = Xlin + 2
              Xarchexelcar.Range("A4", "O" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
              Xarchexelcar.Range("A4", "O" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
              Xarchexelcar.Range("A4", "O" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
              Xarchexelcar.Range("A4", "O" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
              Xarchexelcar.Range("A4", "O" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
              Xarchexelcar.Range("A4", "O" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
              Xarchexelcar.Range("A" & Trim(str(Xlin)), "O" & Trim(str(Xlin))).Interior.color = RGB(0, 160, 0)
              Xarchexelcar.Range("A" & Trim(str(Xlin))).ColumnWidth = 15
             
              Xarchexelcar.Cells(Xlin, XCol) = "FECHA"
              XCol = XCol + 1
              Xarchexelcar.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
              Xarchexelcar.Cells(Xlin, XCol) = "USUARIO"
              XCol = XCol + 1
              Xarchexelcar.Range("C" & Trim(str(Xlin))).ColumnWidth = 15
              Xarchexelcar.Cells(Xlin, XCol) = "MATRICULA"
              XCol = XCol + 1
              Xarchexelcar.Range("D" & Trim(str(Xlin))).ColumnWidth = 40
              Xarchexelcar.Cells(Xlin, XCol) = "NOMBRE"
              XCol = XCol + 1
              Xarchexelcar.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
              Xarchexelcar.Cells(Xlin, XCol) = "CONVENIO"
              XCol = XCol + 1
              Xarchexelcar.Range("F" & Trim(str(Xlin))).ColumnWidth = 28
              Xarchexelcar.Cells(Xlin, XCol) = "SECTOR QUE AUTORIZA"
              XCol = XCol + 1
              Xarchexelcar.Range("G" & Trim(str(Xlin))).ColumnWidth = 23
              Xarchexelcar.Cells(Xlin, XCol) = "CONTACTO"
              XCol = XCol + 1
              Xarchexelcar.Range("H" & Trim(str(Xlin))).ColumnWidth = 15
              Xarchexelcar.Cells(Xlin, XCol) = "URGENCIA?"
              XCol = XCol + 1
              Xarchexelcar.Range("I" & Trim(str(Xlin))).ColumnWidth = 10
              Xarchexelcar.Cells(Xlin, XCol) = "BASE"
              XCol = XCol + 1
              Xarchexelcar.Range("J" & Trim(str(Xlin))).ColumnWidth = 22
              Xarchexelcar.Cells(Xlin, XCol) = "OBSERVACIONES"
                
              Xlin = Xlin + 1
              XCol = 1
              frm_infaccadm.MousePointer = 11
              Lamatric = 0
              Do While Not Data1.Recordset.EOF
                 If Lamatric = Data1.Recordset("socio") Then
                 Else
                    Xcantacc = Xcantacc + 1
                    If IsNull(Data1.Recordset("fecha")) = False Then
                       Xarchexelcar.Cells(Xlin, XCol) = "'" & Format(Data1.Recordset("fecha"), "dd/mm/yyyy")
                    Else
                       Xarchexelcar.Cells(Xlin, XCol) = "Sin fecha"
                    End If
                    XCol = XCol + 1
                    If IsNull(Data1.Recordset("usuario")) = False Then
                       Xarchexelcar.Cells(Xlin, XCol) = Data1.Recordset("usuario")
                    Else
                       Xarchexelcar.Cells(Xlin, XCol) = "s/d"
                    End If
                    XCol = XCol + 1
                    If IsNull(Data1.Recordset("socio")) = False Then
                       Xarchexelcar.Cells(Xlin, XCol) = Data1.Recordset("socio")
                    Else
                       Xarchexelcar.Cells(Xlin, XCol) = 0
                    End If
                    XCol = XCol + 1
                    If IsNull(Data1.Recordset("cl_apellid")) = False Then
                       Xarchexelcar.Cells(Xlin, XCol) = Data1.Recordset("cl_apellid")
                    Else
                       Xarchexelcar.Cells(Xlin, XCol) = "NN"
                    End If
                    XCol = XCol + 1
                    If IsNull(Data1.Recordset("cl_codconv")) = False Then
                       Xarchexelcar.Cells(Xlin, XCol) = Data1.Recordset("cl_codconv")
                    Else
                       Xarchexelcar.Cells(Xlin, XCol) = "NN"
                    End If
                    XCol = XCol + 1
                    XbaseCGal = Data1.Recordset("base")
                    If IsNull(Data1.Recordset("modulo")) = False Then
                       If XbaseCGal = 91 Or XbaseCGal = 92 Or XbaseCGal = 93 Then
                          Xarchexelcar.Cells(Xlin, XCol) = Data1.Recordset("modulo") & "/FARMACIA"
                       Else
                         If Data1.Recordset("modulo") = "FACTURACION" Then
                            Xarchexelcar.Cells(Xlin, XCol) = Data1.Recordset("modulo") & "/RECEPCION"
                         Else
                            Xarchexelcar.Cells(Xlin, XCol) = Data1.Recordset("modulo")
                         End If
                       End If
                    Else
                       Xarchexelcar.Cells(Xlin, XCol) = "s/d"
                    End If
                    XCol = XCol + 1
                    If Data1.Recordset("modulo") = "DESPACHO" Then
                       If IsNull(Data1.Recordset("contacto")) = False Then
                          If Trim(Data1.Recordset("contacto")) <> "" Then
                             Xarchexelcar.Cells(Xlin, XCol) = Data1.Recordset("contacto")
                          Else
                             If Len(Trim(Data1.Recordset("cuando"))) = 9 Then
                                BuscaCedDesp = Mid(Trim(Data1.Recordset("cuando")), 1, 7)
                             Else
                                If Len(Trim(Data1.Recordset("cuando"))) = 8 Then
                                   BuscaCedDesp = Mid(Trim(Data1.Recordset("cuando")), 1, 6)
                                Else
                                   BuscaCedDesp = "0"
                                End If
                             End If
                             If Val(BuscaCedDesp) > 0 Then
                                data_llam.RecordSource = "select * from llamado where ci =" & Val(BuscaCedDesp) & " and fecha =#" & Format(Data1.Recordset("fecha"), "yyyy/mm/dd") & "#"
                                data_llam.Refresh
                                If data_llam.Recordset.RecordCount > 0 Then
                                   If IsNull(data_llam.Recordset("telef")) = False Then
                                      Xarchexelcar.Cells(Xlin, XCol) = data_llam.Recordset("telef")
                                   Else
                                      Xarchexelcar.Cells(Xlin, XCol) = "Consultar Desp."
                                   End If
                                Else
                                   Xarchexelcar.Cells(Xlin, XCol) = "Consultar Desp."
                                End If
                             Else
                                Xarchexelcar.Cells(Xlin, XCol) = "Consultar Desp."
                             End If
                          End If
                       Else
                          If Len(Trim(Data1.Recordset("cuando"))) = 9 Then
                             BuscaCedDesp = Mid(Trim(Data1.Recordset("cuando")), 1, 7)
                          Else
                             If Len(Trim(Data1.Recordset("cuando"))) = 8 Then
                                BuscaCedDesp = Mid(Trim(Data1.Recordset("cuando")), 1, 6)
                             Else
                                BuscaCedDesp = "0"
                             End If
                          End If
                          If Val(BuscaCedDesp) > 0 Then
                             data_llam.RecordSource = "select * from llamado where ci =" & Val(BuscaCedDesp) & " and fecha =#" & Format(Data1.Recordset("fecha"), "yyyy/mm/dd") & "#"
                             data_llam.Refresh
                             If data_llam.Recordset.RecordCount > 0 Then
                                If IsNull(data_llam.Recordset("telef")) = False Then
                                   Xarchexelcar.Cells(Xlin, XCol) = data_llam.Recordset("telef")
                                Else
                                   Xarchexelcar.Cells(Xlin, XCol) = "Consultar Desp."
                                End If
                             Else
                                Xarchexelcar.Cells(Xlin, XCol) = "Consultar Desp."
                             End If
                          Else
                             Xarchexelcar.Cells(Xlin, XCol) = "Consultar Desp."
                          End If
                       End If
                    Else
                       If IsNull(Data1.Recordset("contacto")) = False Then
                          Xarchexelcar.Cells(Xlin, XCol) = Data1.Recordset("contacto")
                       Else
                          Xarchexelcar.Cells(Xlin, XCol) = "s/d"
                       End If
                    End If
                    XCol = XCol + 1
                    If IsNull(Data1.Recordset("codaut")) = False Then
                       If Data1.Recordset("codaut") = "C.GALICIA" Then
                          Xarchexelcar.Cells(Xlin, XCol) = "NO"
                       Else
                          Xarchexelcar.Cells(Xlin, XCol) = "SI"
                       End If
                    Else
                       Xarchexelcar.Cells(Xlin, XCol) = "s/d"
                    End If
                    XCol = XCol + 1
                    If IsNull(Data1.Recordset("base")) = False Then
                       Xarchexelcar.Cells(Xlin, XCol) = Data1.Recordset("base")
                    Else
                       Xarchexelcar.Cells(Xlin, XCol) = "s/d"
                    End If
                    Xlin = Xlin + 1
                    XCol = 1
                 
                 End If
                 Lamatric = Data1.Recordset("socio")
                 Data1.Recordset.MoveNext
              Loop
              Xlin = Xlin + 1
              XCol = 4
              Xarchexelcar.Cells(Xlin, XCol) = "TOTAL DE REGISTROS: " & Xcantacc
                              
              frm_infaccadm.MousePointer = 0
            
              Xlibexelcar.Save
              Xlibexelcar.Close
              Xobjexelcar.Quit
              MsgBox "El archivo autoriza.xls ha sido guardado en la carpeta PLANILLAS del disco C", vbInformation
              Xlabrir3.Workbooks.Open Xarchtex, , False
              Xlabrir3.Visible = True
              Xlabrir3.WindowState = xlMaximized
         
         End If
      Else
         Do While Not Data1.Recordset.EOF
            If Combo1.ListIndex = 2 Then
               If Xlamatacc = Data1.Recordset("cliente") Then
                  If IsNull(Data1.Recordset("cl_nro_sup")) = False Then
                     Xcantacc = Xcantacc + 1
                     If Xcantacc = 1 Then
                        data_inf.Recordset.AddNew
                        data_inf.Recordset("cl_codigo") = Data1.Recordset("cliente")
                        data_inf.Recordset("cl_apellid") = Data1.Recordset("nombre")
                        data_inf.Recordset("cl_codced") = Xcantacc
                        data_inf.Recordset("cl_fecing") = Data1.Recordset("fecha")
                        data_inf.Recordset("cl_ruc") = Data1.Recordset("cl_ruc")
                        data_inf.Recordset("cl_cedula") = Data1.Recordset("importe")
                        data_inf.Recordset("cl_nrovend") = Data1.Recordset("importe") / 1.1 * 0.1
                        data_inf.Recordset("cl_atrasop") = Data1.Recordset("cl_atrasop")
                        data_inf.Recordset("cl_descpag") = Data1.Recordset("cl_descpag")
                        Xlafecdesde = Data1.Recordset("fecha") - 14
                        Xlafechasta = Data1.Recordset("fecha") + 15
                        data_cli.RecordSource = "Select matric,telef,fecha from llamado where matric =" & Data1.Recordset("cliente") & " and fecha >='" & Format(Xlafecdesde, "yyyy-mm-dd") & "' and fecha <='" & Format(Xlafechasta, "yyyy-mm-dd") & "'"
                        data_cli.Refresh
                        If data_cli.Recordset.RecordCount > 0 Then
                           data_inf.Recordset("cl_direcci") = Mid(data_cli.Recordset("telef"), 1, 80) 'contacto
                        Else
                           data_inf.Recordset("cl_direcci") = Mid(Data1.Recordset("cl_desc1"), 1, 80)
                        End If
                        data_inf.Recordset.Update
                     Else
                        data_inf.RecordSource = "Select * from infcli where cl_codigo =" & Data1.Recordset("cliente") & " and cl_codced >=" & 1
                        data_inf.Refresh
                        If data_inf.Recordset.RecordCount > 0 Then
                           data_inf.Recordset.Edit
                           data_inf.Recordset("cl_codced") = data_inf.Recordset("cl_codced") + 1
                           data_inf.Recordset.Update
                        End If
                     End If
                  Else
                     data_inf.Recordset.AddNew
                     data_inf.Recordset("cl_codigo") = Data1.Recordset("cliente")
                     data_inf.Recordset("cl_apellid") = Data1.Recordset("nombre")
                     data_inf.Recordset("cl_codced") = 0
                     data_inf.Recordset("cl_fecing") = Data1.Recordset("fecha")
                     data_inf.Recordset("cl_ruc") = Data1.Recordset("cl_ruc")
                     data_inf.Recordset("cl_atrasop") = Data1.Recordset("cl_atrasop")
                     data_inf.Recordset("cl_descpag") = Data1.Recordset("cl_descpag")
                     data_inf.Recordset("cl_cedula") = Data1.Recordset("importe")
                     data_inf.Recordset("cl_nrovend") = Data1.Recordset("importe") / 1.1 * 0.1
                     data_inf.Recordset("info_debit") = Data1.Recordset("info_debit")
                     Xlafecdesde = Data1.Recordset("fecha") - 14
                     Xlafechasta = Data1.Recordset("fecha") + 15
                     data_cli.RecordSource = "Select matric,telef,fecha from llamado where matric =" & Data1.Recordset("cliente") & " and fecha >='" & Format(Xlafecdesde, "yyyy-mm-dd") & "' and fecha <='" & Format(Xlafechasta, "yyyy-mm-dd") & "'"
                     data_cli.Refresh
                     If data_cli.Recordset.RecordCount > 0 Then
                        data_inf.Recordset("cl_direcci") = Mid(data_cli.Recordset("telef"), 1, 80) 'contacto
                     Else
                        data_inf.Recordset("cl_direcci") = Mid(Data1.Recordset("cl_desc1"), 1, 80)
                     End If
                     data_inf.Recordset.Update
                  End If
                  Xlamatacc = Data1.Recordset("cliente")
               Else
                  Xcantacc = 0
                  Xlamatacc = Data1.Recordset("cliente")
                  Data1.Recordset.MovePrevious
               End If
            Else
               If Combo1.ListIndex = 3 Then
                  data_inf.Recordset.AddNew
                  data_inf.Recordset("cl_codigo") = Data1.Recordset("socio")
                  data_inf.Recordset("cl_apellid") = Data1.Recordset("cl_apellid")
                  data_inf.Recordset("cl_telefon") = Data1.Recordset("cl_dpto")
                  data_inf.Recordset("cl_fecing") = Data1.Recordset("fecha")
                  data_inf.Recordset("cl_ruc") = Mid(Data1.Recordset("codaut"), 1, 12)
                  data_inf.Recordset("cl_email") = Mid(Data1.Recordset("usuario"), 1, 30) 'quien autoriza
                  data_inf.Recordset("cl_nomcobr") = Data1.Recordset("usuario_caja")
                  data_inf.Recordset("cl_nombre") = Data1.Recordset("modulo")
                  data_inf.Recordset.Update
               Else
                   data_inf.Recordset.AddNew
                   data_inf.Recordset("cl_codigo") = Val(Data1.Recordset("cl_zona"))
                   If IsNull(Data1.Recordset("cl_zona")) = False Then
                      data_cli.RecordSource = "Select * from clientes where cl_codigo =" & Data1.Recordset("cl_zona")
                      data_cli.Refresh
                      If data_cli.Recordset.RecordCount > 0 Then
                         data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                         If IsNull(data_cli.Recordset("cl_dpto")) = False Then
                            data_inf.Recordset("cl_direcci") = Mid(data_cli.Recordset("cl_dpto"), 1, 30) 'contacto
                         Else
                            If IsNull(data_cli.Recordset("cl_telefon")) = False Then
                               data_inf.Recordset("cl_direcci") = Mid(data_cli.Recordset("cl_telefon"), 1, 30) 'contacto
                            End If
                         End If
                      Else
                         data_inf.Recordset("cl_apellid") = "NN"
                      End If
                   End If
                   data_inf.Recordset("cl_fecing") = Data1.Recordset("cl_fultpag")
               '         data_inf.Recordset("estado") = Data1.Recordset("estado")
                   data_inf.Recordset("cl_ruc") = Data1.Recordset("cl_ruc")
                   data_inf.Recordset("cl_atrasop") = Data1.Recordset("cl_atrasop")
                   data_inf.Recordset("cl_descpag") = Data1.Recordset("cl_descpag")
                   data_inf.Recordset("cl_nro_sup") = Data1.Recordset("cl_nro_sup") 'Nro de referencia
                   data_inf.Recordset("cl_email") = Data1.Recordset("cl_email") 'descrip referencia
                   data_inf.Recordset("info_debit") = Data1.Recordset("info_debit")
                   data_inf.Recordset("cl_fnac") = Data1.Recordset("cl_fec1") 'fecha pago
                   data_inf.Recordset("cl_nomcobr") = Mid(Data1.Recordset("cl_desc2"), 1, 25) ' forma de pago
                   If Combo1.ListIndex = 2 Then
                      Xlafecdesde = Data1.Recordset("fecha") - 14
                      Xlafechasta = Data1.Recordset("fecha") + 15
                       data_cli.RecordSource = "Select matric,telef,fecha from llamado where matric =" & Val(Data1.Recordset("cl_zona")) & " and fecha >='" & Format(Xlafecdesde, "yyyy-mm-dd") & "' and fecha <='" & Format(Xlafechasta, "yyyy-mm-dd") & "'"
                       data_cli.Refresh
                       If data_cli.Recordset.RecordCount > 0 Then
                          data_inf.Recordset("cl_direcci") = Mid(data_cli.Recordset("telef"), 1, 80) 'contacto
                       Else
                          data_inf.Recordset("cl_direcci") = Mid(Data1.Recordset("cl_desc1"), 1, 80)
                       End If
                   End If
                   
                   data_inf.Recordset.Update
               End If
            End If
            Data1.Recordset.MoveNext
         Loop
         data_inf.RecordSource = "Select * from infcli"
         data_inf.Refresh
         frm_infaccadm.MousePointer = 0
         Command1.Enabled = True
         If Combo1.ListIndex = 0 Then
            cr1.ReportFileName = App.path & "\infaccadm2.rpt"
            cr1.ReportTitle = "Informe de TODAS las acciones administrativas desde:" & mfd.Text & " HASTA:" & mfh.Text
         Else
            If Combo1.ListIndex = 2 Then
               cr1.ReportFileName = App.path & "\infaccadmnew.rpt"
               cr1.ReportTitle = "Informe Gestión de Cobranza DESDE:" & mfd.Text & " HASTA:" & mfh.Text
            Else
               If Combo1.ListIndex = 3 Then
                  cr1.ReportFileName = App.path & "\infcodautadm.rpt"
                  cr1.ReportTitle = "Informe de CODIGOS de autorización ADM. desde:" & mfd.Text & " HASTA:" & mfh.Text
               Else
                  cr1.ReportFileName = App.path & "\infaccadm.rpt"
                  cr1.ReportTitle = "Informe de acciones administrativas desde:" & mfd.Text & " HASTA:" & mfh.Text
               End If
            End If
         End If
         cr1.Action = 1
      End If
      
   End If
End If
frm_infaccadm.MousePointer = 0

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()


'data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cli.ConnectionString = "dsn=" & Xconexrmt
data_llam.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_lin.ConnectionString = "dsn=" & Xconexrmt
'data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.ConnectionString = "dsn=" & Xconexrmt
'data_gri.RecordSource = "Select * from mant_sol where estado >" & 0 & " order by estado"
'data_gri.Refresh
'data_inf.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"
data_inf.DatabaseName = App.path & "\informes.mdb"
data_inf.RecordSource = "infcli"
data_inf.Refresh

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
   Combo1.SetFocus
End If

End Sub
