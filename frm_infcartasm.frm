VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infcartasm 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de cartas mutuales"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7110
   Icon            =   "frm_infcartasm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   7110
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport cr2 
      Left            =   4440
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Nuevos"
      Height          =   495
      Left            =   2040
      TabIndex        =   10
      Top             =   3480
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSAdodcLib.Adodc data_cli2 
      Height          =   495
      Left            =   4320
      Top             =   2640
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
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
      Caption         =   "data_cli2"
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
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   960
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      Height          =   495
      Left            =   6360
      Picture         =   "frm_infcartasm.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Height          =   495
      Left            =   5280
      Picture         =   "frm_infcartasm.frx":0B14
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3240
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C00000&
      Caption         =   "Datos del informe"
      ForeColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6615
      Begin MSAdodcLib.Adodc data_cli 
         Height          =   375
         Left            =   1560
         Top             =   2400
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
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
         Caption         =   "data_cli"
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
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_infcartasm.frx":109E
         Left            =   1920
         List            =   "frm_infcartasm.frx":10BD
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1680
         Width           =   3495
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_infcartasm.frx":1157
         Left            =   1920
         List            =   "frm_infcartasm.frx":116D
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   960
         Width           =   3495
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   3840
         TabIndex        =   3
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
         Left            =   1920
         TabIndex        =   2
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         Caption         =   "Opción de informe:"
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
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Caption         =   "Mutualista:"
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
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "Rango de fechas:"
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
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   120
      Picture         =   "frm_infcartasm.frx":11AD
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   1695
   End
End
Attribute VB_Name = "frm_infcartasm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'On Error GoTo Quepasacartas

If Combo2.Text = "LLAMADOS" Or Combo2.Text = "POLICLINICAS" Or Combo2.Text = "RESUMEN DE CARTAS" Or Combo2.Text = "CARTAS FACTURADAS" Then
   Command3_Click
Else
    
    Dim Xobjexelcar As Excel.Application
    Dim Xlibexelcar As Excel.Workbook
    Dim Xarchexelcar As New Excel.Worksheet
    
    Dim Xlabrir3 As New Excel.Application
    
    Dim XCol, Xlin, Xnrocan, Xcolfija As Long
    Dim Xarchtex As String
    Dim Xtitmutual As String
        
    If Combo1.ListIndex = 0 Then
       Xtitmutual = "*CIRCULO CATOLICO DE OBREROS DEL URUGUAY*"
    End If
    If Combo1.ListIndex = 1 Then
       Xtitmutual = "*HOSPITAL EVANGELICO*"
    End If
    If Combo1.ListIndex = 2 Then
       Xtitmutual = "*SMI*"
    End If
    If Combo1.ListIndex = 3 Then
       Xtitmutual = "*UNIVERSAL*"
    End If
    If Combo1.ListIndex = 4 Then
       Xtitmutual = "*CASA DE GALICIA*"
    End If
    Xnrocan = 1
    
    If md.Text = "__/__/____" And mh.Text = "__/__/____" Then
       MsgBox "Debe ingresar Fechas"
    Else
       frm_infcartasm.MousePointer = 11
       
       Dim MiBaseact As Database
       Dim Unasesact As Workspace
       Set Unasesact = Workspaces(0)
       Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")
       MiBaseact.Execute "Delete * from infcli"
       data_inf.RecordSource = "infcli"
       data_inf.Refresh
       If Combo2.ListIndex = 3 Then
          data_cli.RecordSource = "Select * from clientes where saldo_chc2 =" & 1
          data_cli.Refresh
       Else
          If Combo2.ListIndex = 1 Then
             If Combo1.ListIndex = 0 Then
                data_cli.RecordSource = "select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cod_prod in (805)"
             Else
                If Combo1.ListIndex = 1 Then
                   data_cli.RecordSource = "select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cod_prod in (804)"
                Else
                   If Combo1.ListIndex = 2 Then
                      data_cli.RecordSource = "select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cod_prod in (802)"
                   Else
                      If Combo1.ListIndex = 3 Then
                         data_cli.RecordSource = "select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cod_prod in (803)"
                      Else
                         If Combo1.ListIndex = 4 Then
                            data_cli.RecordSource = "select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cod_prod in (806)"
                         Else
                            data_cli.RecordSource = "select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cod_prod in (802,803,804,805,806)"
                         End If
                      End If
                   End If
                End If
             End If
             data_cli.Refresh
          Else
             If Combo2.ListIndex = 4 Then
                data_cli.RecordSource = "Select * from clientes where fecha_reac >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha_reac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_decuota=" & 3
                data_cli.Refresh
             Else
                If Combo2.ListIndex = 0 Then
                   data_cli.RecordSource = "Select * from clientes where fecha_reac >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha_reac <='" & Format(mh.Text, "yyyy-mm-dd") & "' and cl_decuota=" & 1
                   data_cli.Refresh
                Else
                   data_cli.RecordSource = "Select * from clientes where fecha_reac >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha_reac <='" & Format(mh.Text, "yyyy-mm-dd") & "'"
                   data_cli.Refresh
                End If
             End If
          End If
       End If
       
       If data_cli.Recordset.RecordCount > 0 Then
          data_cli.Recordset.MoveFirst
          Do While Not data_cli.Recordset.EOF
             If Combo2.ListIndex = 1 Then
                data_inf.Recordset.AddNew
                data_inf.Recordset("cl_codigo") = data_cli.Recordset("cod_cli")
                data_inf.Recordset("cl_apellid") = data_cli.Recordset("nom_cli")
                data_inf.Recordset("cl_cedula") = data_cli.Recordset("ced_socio")
                data_inf.Recordset("cl_codced") = data_cli.Recordset("fact")
                data_inf.Recordset("fecha_reac") = data_cli.Recordset("fecha")
                data_inf.Recordset("saldo_cc") = data_cli.Recordset("factura")
                data_cli2.RecordSource = "Select * from clientes where cl_codigo =" & data_cli.Recordset("cod_cli")
                data_cli2.Refresh
                If data_cli2.Recordset.RecordCount > 0 Then
                   data_inf.Recordset("cl_fnac") = data_cli2.Recordset("cl_fnac")
                   data_inf.Recordset("cl_fecing") = data_cli2.Recordset("cl_fecing")
                   data_inf.Recordset("cl_direcci") = data_cli2.Recordset("cl_direcci")
                   data_inf.Recordset("cl_zona") = data_cli2.Recordset("cl_zona")
                   data_inf.Recordset("cl_telefon") = data_cli2.Recordset("cl_telefon")
                   data_inf.Recordset("cl_codconv") = data_cli2.Recordset("cl_codconv")
                   data_inf.Recordset("cl_socmnom") = data_cli2.Recordset("cl_socmnom")
                   data_inf.Recordset("cl_decuota") = data_cli2.Recordset("cl_decuota")
                   data_inf.Recordset("cl_dpto") = data_cli2.Recordset("cl_dpto")
                   data_inf.Recordset("cl_nombre") = data_cli2.Recordset("cl_nombre")
                End If
                data_cli2.RecordSource = "select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                data_cli2.Refresh
                If data_cli2.Recordset.RecordCount > 0 Then
                   data_inf.Recordset("cl_socmnom") = data_cli2.Recordset("cnv_grupo")
                End If
                data_inf.Recordset.Update
             Else
                data_inf.Recordset.AddNew
                data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                data_inf.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                data_inf.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                data_inf.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                data_inf.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                data_inf.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                data_inf.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                data_inf.Recordset("cl_socmnom") = data_cli.Recordset("cl_socmnom")
                data_inf.Recordset("fecha_reac") = data_cli.Recordset("fecha_reac")
                data_inf.Recordset("cl_decuota") = data_cli.Recordset("cl_decuota")
                data_inf.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                data_inf.Recordset("cl_nombre") = data_cli.Recordset("cl_nombre")
                If Combo2.ListIndex = 0 Then
                   data_cli2.RecordSource = "Select * from abmsocio where cl_codigo =" & data_cli.Recordset("cl_codigo") & " and cl_motivo ='" & "AVISO P/FIRMAR CARTA" & "'"
                   data_cli2.Refresh
                   If data_cli2.Recordset.RecordCount > 0 Then
                      data_cli2.Recordset.MoveLast
                      data_inf.Recordset("cl_cantdia") = data_cli2.Recordset.RecordCount
                   Else
                      data_inf.Recordset("cl_cantdia") = 0
                   End If
                   data_cli2.RecordSource = "Select * from linmmdd where cod_cli =" & data_cli.Recordset("cl_codigo") & " and fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and cod_prod not in (999,997,990,995,881,882,800,801,802,803,804,805,806,991,992,993,994,996,8000)"
                   data_cli2.Refresh
                   If data_cli2.Recordset.RecordCount > 0 Then
                      data_cli2.Recordset.MoveLast
                      data_inf.Recordset("cl_fultvta") = data_cli2.Recordset("fecha")
                   Else
                   
                   End If
                
                End If
                data_cli2.RecordSource = "select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                data_cli2.Refresh
                If data_cli2.Recordset.RecordCount > 0 Then
                   data_inf.Recordset("cl_socmnom") = data_cli2.Recordset("cnv_grupo")
                End If
                data_inf.Recordset.Update
             End If
             data_cli.Recordset.MoveNext
          Loop
          If data_inf.Recordset.RecordCount > 0 Then
             If Combo1.Text = "SMI" Then
                data_inf.Recordset.MoveFirst
                Do While Not data_inf.Recordset.EOF
                   If IsNull(data_inf.Recordset("cl_socmnom")) = False Then
                      If data_inf.Recordset("cl_socmnom") = Combo1.Text Or data_inf.Recordset("cl_socmnom") = "IMPASA" Then
                      Else
                         data_inf.Recordset.Delete
                      End If
                   Else
                      data_inf.Recordset.Delete
                   End If
                   data_inf.Recordset.MoveNext
                Loop
                data_inf.Refresh
             Else
                If Combo1.Text = "TODOS" Then
                Else
                   data_inf.Recordset.MoveFirst
                   Do While Not data_inf.Recordset.EOF
                      If IsNull(data_inf.Recordset("cl_socmnom")) = False Then
                         If data_inf.Recordset("cl_socmnom") = Combo1.Text Then
                         Else
                            data_inf.Recordset.Delete
                         End If
                      Else
                         data_inf.Recordset.Delete
                      End If
                      data_inf.Recordset.MoveNext
                   Loop
                   data_inf.Refresh
                End If
             End If
             If Combo2.ListIndex = 0 Then
                data_inf.Recordset.MoveFirst
                Do While Not data_inf.Recordset.EOF
                   If IsNull(data_inf.Recordset("cl_decuota")) = False Then
                      If data_inf.Recordset("cl_decuota") = 1 Then
                      Else
                         data_inf.Recordset.Delete
                      End If
                   Else
                      data_inf.Recordset.Delete
                   End If
                   data_inf.Recordset.MoveNext
                Loop
                data_inf.Refresh
             End If
             If Combo2.ListIndex = 2 Then ' Cartas realizadas
                data_inf.Recordset.MoveFirst
                Do While Not data_inf.Recordset.EOF
                   If IsNull(data_inf.Recordset("cl_decuota")) = False Then
                      If data_inf.Recordset("cl_decuota") = 2 Then
                      Else
                         data_inf.Recordset.Delete
                      End If
                   Else
                      data_inf.Recordset.Delete
                   End If
                   data_inf.Recordset.MoveNext
                Loop
                data_inf.Refresh
             End If
             If Combo2.ListIndex = 2 Then
                data_inf.Recordset.MoveFirst
                Set Xobjexelcar = New Excel.Application
                Set Xlibexelcar = Xobjexelcar.Workbooks.Add
                Set Xarchexelcar = Xlibexelcar.Worksheets.Add
                Xlin = 1
                XCol = 1
                Xarchexelcar.Name = "cartas"
                Xlibexelcar.SaveAs ("C:\planillas\cartas.xls")
                Xarchtex = "C:\planillas\cartas.xls"
                Xarchexelcar.Range("A1", "C3").Font.Size = 16
                Xarchexelcar.Cells(Xlin, XCol) = Xtitmutual
                Xlin = Xlin + 1
                XCol = 1
                Xarchexelcar.Cells(Xlin, XCol) = "SOLICITUD PARA INGRESAR AL PADRÓN DE SAPP"
                Xlin = Xlin + 1
                Xarchexelcar.Cells(Xlin, XCol) = "ENTREGADAS: " & mh.Text
    '            Xlin = Xlin + 1
    '            Xarchexelcar.Range("B" & Trim(Str(Xlin)), "I" & Trim(Str(Xlin))).Interior.color = RGB(0, 200, 200)
                XCol = 1
                Xlin = Xlin + 2
                If data_inf.Recordset.RecordCount > 0 Then
                   data_inf.RecordSource = "Select * from infcli order by fecha_reac"
                   data_inf.Refresh
                   data_inf.Recordset.MoveLast
                   data_inf.Recordset.MoveFirst
                End If
                Xnrocan = Xnrocan + data_inf.Recordset.RecordCount + Xlin
                
                Xarchexelcar.Range("A4", "H" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
                Xarchexelcar.Range("A4", "H" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
                Xarchexelcar.Range("A4", "H" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
                Xarchexelcar.Range("A4", "H" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
                Xarchexelcar.Range("A4", "H" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
                Xarchexelcar.Range("A4", "H" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
                Xarchexelcar.Range("A" & Trim(str(Xlin)), "H" & Trim(str(Xlin))).Interior.color = RGB(0, 160, 0)
                Xarchexelcar.Range("A" & Trim(str(Xlin))).ColumnWidth = 15
    
                Xarchexelcar.Cells(Xlin, XCol) = "FECHA"
                XCol = XCol + 1
                Xarchexelcar.Range("B" & Trim(str(Xlin))).ColumnWidth = 40
                Xarchexelcar.Cells(Xlin, XCol) = "NOMBRES"
                XCol = XCol + 1
                Xarchexelcar.Range("C" & Trim(str(Xlin))).ColumnWidth = 15
                Xarchexelcar.Cells(Xlin, XCol) = "DOCUMENTO"
                XCol = XCol + 1
                Xarchexelcar.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
                Xarchexelcar.Cells(Xlin, XCol) = "NACIMIENTO"
                XCol = XCol + 1
                Xarchexelcar.Range("E" & Trim(str(Xlin))).ColumnWidth = 43
                Xarchexelcar.Cells(Xlin, XCol) = "DOMICILIO"
                XCol = XCol + 1
                Xarchexelcar.Range("F" & Trim(str(Xlin))).ColumnWidth = 17
                Xarchexelcar.Cells(Xlin, XCol) = "ZONA"
                XCol = XCol + 1
                Xarchexelcar.Range("G" & Trim(str(Xlin))).ColumnWidth = 15
                Xarchexelcar.Cells(Xlin, XCol) = "TELEFONO"
                XCol = XCol + 1
                Xarchexelcar.Range("H" & Trim(str(Xlin))).ColumnWidth = 15
                Xarchexelcar.Cells(Xlin, XCol) = "TIPO AF."
                Xlin = Xlin + 1
                XCol = 1
                Do While Not data_inf.Recordset.EOF
                   Xarchexelcar.Cells(Xlin, XCol) = "'" & Format(data_inf.Recordset("fecha_reac"), "dd/mm/yyyy")
                   XCol = XCol + 1
                   Xarchexelcar.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
                   XCol = XCol + 1
                   Xarchexelcar.Cells(Xlin, XCol) = Trim(str(data_inf.Recordset("cl_cedula"))) & "-" & Trim(str(data_inf.Recordset("cl_codced")))
                   XCol = XCol + 1
                   If IsNull(data_inf.Recordset("cl_fnac")) = False Then
                      Xarchexelcar.Cells(Xlin, XCol) = "'" & Format(data_inf.Recordset("cl_fnac"), "dd/mm/yyyy")
                   Else
                      Xarchexelcar.Cells(Xlin, XCol) = "S/Fecha"
                   End If
                   XCol = XCol + 1
                   If IsNull(data_inf.Recordset("cl_direcci")) = False Then
                      If Trim(data_inf.Recordset("cl_direcci")) <> "" Then
                         Xarchexelcar.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
                      Else
                         Xarchexelcar.Cells(Xlin, XCol) = "Sin Datos"
                      End If
                   Else
                      Xarchexelcar.Cells(Xlin, XCol) = "Sin datos"
                   End If
                   XCol = XCol + 1
                   Xarchexelcar.Cells(Xlin, XCol) = data_inf.Recordset("cl_zona")
                   XCol = XCol + 1
                   Xarchexelcar.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
                   XCol = XCol + 1
                   Xarchexelcar.Cells(Xlin, XCol) = data_inf.Recordset("cl_nombre")
                   XCol = XCol + 1
              '     Xarchexelcar.Range("A" & Trim(Str(Xlin)), "H" & Trim(Str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
                   data_inf.Recordset.MoveNext
                   Xlin = Xlin + 1
                   XCol = 1
                Loop
                Xlibexelcar.Save
        '            Xlibexel.Application
                Xlibexelcar.Close
                Xobjexelcar.Quit
                MsgBox "El archivo ha sido guardado en la carpeta PLANILLAS del disco C", vbInformation
                
                Xlabrir3.Workbooks.Open Xarchtex, , False
                Xlabrir3.Visible = True
                Xlabrir3.WindowState = xlMaximized
    
    '                Shell frm_menu.data_usuac.Recordset("destino") & "excel.exe " & Xarchtex, vbMaximizedFocus
             End If
             If Combo2.ListIndex = 1 Then 'realizadas opcion 1
                data_inf.Recordset.MoveFirst
                Set Xobjexelcar = New Excel.Application
                Set Xlibexelcar = Xobjexelcar.Workbooks.Add
                Set Xarchexelcar = Xlibexelcar.Worksheets.Add
                Xlin = 1
                XCol = 1
                Xarchexelcar.Name = "cartas"
                Xlibexelcar.SaveAs ("C:\planillas\cartas.xls")
                Xarchtex = "C:\planillas\cartas.xls"
                Xarchexelcar.Range("A1", "C3").Font.Size = 16
                Xarchexelcar.Cells(Xlin, XCol) = Xtitmutual
                Xlin = Xlin + 1
                XCol = 1
                Xarchexelcar.Cells(Xlin, XCol) = "CARTAS REALIZADAS    Fecha Actual:" & Format(Date, "dd/mm/yyyy")
                Xlin = Xlin + 1
                Xarchexelcar.Cells(Xlin, XCol) = "FECHA: " & Format(md.Text, "dd/mm/yyyy") & " HASTA:" & Format(mh.Text, "dd/mm/yyyy")
    '            Xlin = Xlin + 1
    '            Xarchexelcar.Range("B" & Trim(Str(Xlin)), "I" & Trim(Str(Xlin))).Interior.color = RGB(0, 200, 200)
                XCol = 1
                Xlin = Xlin + 2
                If data_inf.Recordset.RecordCount > 0 Then
                   data_inf.RecordSource = "Select * from infcli order by fecha_reac"
                   data_inf.Refresh
                   data_inf.Recordset.MoveLast
                   data_inf.Recordset.MoveFirst
                End If
                Xnrocan = Xnrocan + data_inf.Recordset.RecordCount + Xlin
                Xarchexelcar.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
                Xarchexelcar.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
                Xarchexelcar.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
                Xarchexelcar.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
                Xarchexelcar.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
                Xarchexelcar.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
                Xarchexelcar.Range("A" & Trim(str(Xlin)), "J" & Trim(str(Xlin))).Interior.color = RGB(0, 160, 0)
                Xarchexelcar.Range("A" & Trim(str(Xlin))).ColumnWidth = 15
                Xarchexelcar.Cells(Xlin, XCol) = "FECHA"
                XCol = XCol + 1
                Xarchexelcar.Range("B" & Trim(str(Xlin))).ColumnWidth = 10
                Xarchexelcar.Cells(Xlin, XCol) = "CONVENIO"
                XCol = XCol + 1
                Xarchexelcar.Range("C" & Trim(str(Xlin))).ColumnWidth = 40
                Xarchexelcar.Cells(Xlin, XCol) = "NOMBRES"
                XCol = XCol + 1
                Xarchexelcar.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
                Xarchexelcar.Cells(Xlin, XCol) = "DOCUMENTO"
                XCol = XCol + 1
                Xarchexelcar.Range("E" & Trim(str(Xlin))).ColumnWidth = 15
                Xarchexelcar.Cells(Xlin, XCol) = "NACIMIENTO"
                XCol = XCol + 1
                Xarchexelcar.Range("F" & Trim(str(Xlin))).ColumnWidth = 43
                Xarchexelcar.Cells(Xlin, XCol) = "DOMICILIO"
                XCol = XCol + 1
                Xarchexelcar.Range("G" & Trim(str(Xlin))).ColumnWidth = 17
                Xarchexelcar.Cells(Xlin, XCol) = "ZONA"
                XCol = XCol + 1
                Xarchexelcar.Range("H" & Trim(str(Xlin))).ColumnWidth = 15
                Xarchexelcar.Cells(Xlin, XCol) = "TELEFONO"
                XCol = XCol + 1
                Xarchexelcar.Range("I" & Trim(str(Xlin))).ColumnWidth = 15
                Xarchexelcar.Cells(Xlin, XCol) = "TIPO AF."
                XCol = XCol + 1
                Xarchexelcar.Range("J" & Trim(str(Xlin))).ColumnWidth = 25
                Xarchexelcar.Cells(Xlin, XCol) = "PROMOTOR"
    
                Xlin = Xlin + 1
                XCol = 1
                Do While Not data_inf.Recordset.EOF
                   Xarchexelcar.Cells(Xlin, XCol) = "'" & Format(data_inf.Recordset("fecha_reac"), "dd/mm/yyyy")
                   XCol = XCol + 1
                   Xarchexelcar.Cells(Xlin, XCol) = data_inf.Recordset("cl_codconv")
                   XCol = XCol + 1
                   Xarchexelcar.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
                   XCol = XCol + 1
                   Xarchexelcar.Cells(Xlin, XCol) = Trim(str(data_inf.Recordset("cl_cedula"))) & "-" & Trim(str(data_inf.Recordset("cl_codced")))
                   XCol = XCol + 1
                   If IsNull(data_inf.Recordset("cl_fnac")) = False Then
                      Xarchexelcar.Cells(Xlin, XCol) = "'" & Format(data_inf.Recordset("cl_fnac"), "dd/mm/yyyy")
                   Else
                      Xarchexelcar.Cells(Xlin, XCol) = "S/Fecha"
                   End If
                   XCol = XCol + 1
                   If IsNull(data_inf.Recordset("cl_direcci")) = False Then
                      If Trim(data_inf.Recordset("cl_direcci")) <> "" Then
                         Xarchexelcar.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
                      Else
                         Xarchexelcar.Cells(Xlin, XCol) = "Sin Datos"
                      End If
                   Else
                      Xarchexelcar.Cells(Xlin, XCol) = "Sin datos"
                   End If
                   XCol = XCol + 1
                   Xarchexelcar.Cells(Xlin, XCol) = data_inf.Recordset("cl_zona")
                   XCol = XCol + 1
                   Xarchexelcar.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
                   XCol = XCol + 1
                   Xarchexelcar.Cells(Xlin, XCol) = data_inf.Recordset("cl_nombre")
                   XCol = XCol + 1
                   data_cli2.RecordSource = "select * from linmmdd_afil where factura =" & data_inf.Recordset("saldo_cc")
                   data_cli2.Refresh
                   If data_cli2.Recordset.RecordCount > 0 Then
                      Xarchexelcar.Cells(Xlin, XCol) = data_cli2.Recordset("nombre")
                   Else
                      Xarchexelcar.Cells(Xlin, XCol) = "Sin Datos"
                   End If
              '     Xarchexelcar.Range("A" & Trim(Str(Xlin)), "H" & Trim(Str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
                   data_inf.Recordset.MoveNext
                   Xlin = Xlin + 1
                   XCol = 1
                Loop
                Xlibexelcar.Save
        '            Xlibexel.Application
                Xlibexelcar.Close
                Xobjexelcar.Quit
                MsgBox "El archivo ha sido guardado en la carpeta PLANILLAS del disco C", vbInformation
                
                Xlabrir3.Workbooks.Open Xarchtex, , False
                Xlabrir3.Visible = True
                Xlabrir3.WindowState = xlMaximized
    
    '                Shell frm_menu.data_usuac.Recordset("destino") & "excel.exe " & Xarchtex, vbMaximizedFocus
             End If
             If Combo2.ListIndex = 0 Then 'Avisos de Firmar carta
                data_inf.Recordset.MoveFirst
                Set Xobjexelcar = New Excel.Application
                Set Xlibexelcar = Xobjexelcar.Workbooks.Add
                Set Xarchexelcar = Xlibexelcar.Worksheets.Add
                Xlin = 1
                XCol = 1
                Xarchexelcar.Name = "cartas"
                Xlibexelcar.SaveAs ("C:\planillas\cartas.xls")
                Xarchtex = "C:\planillas\cartas.xls"
                Xarchexelcar.Range("A1", "C3").Font.Size = 16
                Xarchexelcar.Cells(Xlin, XCol) = "SAPP S.A. Dpto. TI"
                Xlin = Xlin + 1
                XCol = 1
                Xarchexelcar.Cells(Xlin, XCol) = "AVISOS PARA FIRMAR CARTA"
                Xlin = Xlin + 1
                Xarchexelcar.Cells(Xlin, XCol) = "FECHA Actual: " & Format(Date, "dd/mm/yyyy")
    '            Xlin = Xlin + 1
    '            Xarchexelcar.Range("B" & Trim(Str(Xlin)), "I" & Trim(Str(Xlin))).Interior.color = RGB(0, 200, 200)
                XCol = 1
                Xlin = Xlin + 2
                If data_inf.Recordset.RecordCount > 0 Then
                   data_inf.RecordSource = "Select * from infcli order by fecha_reac"
                   data_inf.Refresh
                   data_inf.Recordset.MoveLast
                   data_inf.Recordset.MoveFirst
                End If
                Xnrocan = Xnrocan + data_inf.Recordset.RecordCount + Xlin
                
                Xarchexelcar.Range("A4", "L" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
                Xarchexelcar.Range("A4", "L" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
                Xarchexelcar.Range("A4", "L" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
                Xarchexelcar.Range("A4", "L" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
                Xarchexelcar.Range("A4", "L" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
                Xarchexelcar.Range("A4", "L" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
                Xarchexelcar.Range("A" & Trim(str(Xlin)), "L" & Trim(str(Xlin))).Interior.color = RGB(0, 160, 0)
                
                Xarchexelcar.Range("A" & Trim(str(Xlin))).ColumnWidth = 15
                Xarchexelcar.Cells(Xlin, XCol) = "FECHA"
                XCol = XCol + 1
                Xarchexelcar.Range("B" & Trim(str(Xlin))).ColumnWidth = 10
                Xarchexelcar.Cells(Xlin, XCol) = "CONVENIO"
                XCol = XCol + 1
                Xarchexelcar.Range("C" & Trim(str(Xlin))).ColumnWidth = 15
                Xarchexelcar.Cells(Xlin, XCol) = "GPO.MUTUAL"
                XCol = XCol + 1
                Xarchexelcar.Range("D" & Trim(str(Xlin))).ColumnWidth = 10
                Xarchexelcar.Cells(Xlin, XCol) = "MATRICULA"
                XCol = XCol + 1
                Xarchexelcar.Range("E" & Trim(str(Xlin))).ColumnWidth = 40
                Xarchexelcar.Cells(Xlin, XCol) = "NOMBRE"
                XCol = XCol + 1
                Xarchexelcar.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
                Xarchexelcar.Cells(Xlin, XCol) = "CEDULA"
                XCol = XCol + 1
                Xarchexelcar.Range("G" & Trim(str(Xlin))).ColumnWidth = 43
                Xarchexelcar.Cells(Xlin, XCol) = "DOMICILIO"
                XCol = XCol + 1
                Xarchexelcar.Range("H" & Trim(str(Xlin))).ColumnWidth = 17
                Xarchexelcar.Cells(Xlin, XCol) = "ZONA"
                XCol = XCol + 1
                Xarchexelcar.Range("I" & Trim(str(Xlin))).ColumnWidth = 15
                Xarchexelcar.Cells(Xlin, XCol) = "CELULAR"
                XCol = XCol + 1
                Xarchexelcar.Range("J" & Trim(str(Xlin))).ColumnWidth = 15
                Xarchexelcar.Cells(Xlin, XCol) = "TELEFONO"
                XCol = XCol + 1
                Xarchexelcar.Range("K" & Trim(str(Xlin))).ColumnWidth = 15
                Xarchexelcar.Cells(Xlin, XCol) = "AVISOS"
                XCol = XCol + 1
                Xarchexelcar.Range("L" & Trim(str(Xlin))).ColumnWidth = 15
                Xarchexelcar.Cells(Xlin, XCol) = "ULT.CONS"
                
                Xlin = Xlin + 1
                XCol = 1
                Do While Not data_inf.Recordset.EOF
                   Xarchexelcar.Cells(Xlin, XCol) = CDate(data_inf.Recordset("fecha_reac"))
                   XCol = XCol + 1
                   Xarchexelcar.Cells(Xlin, XCol) = data_inf.Recordset("cl_codconv")
                   XCol = XCol + 1
                   If data_inf.Recordset("cl_codconv") = "SMIN" Or data_inf.Recordset("cl_codconv") = "SMINA" Then
                      Xarchexelcar.Cells(Xlin, XCol) = "SMI"
                   Else
                      If data_inf.Recordset("cl_codconv") = "UNIVS" Or data_inf.Recordset("cl_codconv") = "UNNSAM" Then
                         Xarchexelcar.Cells(Xlin, XCol) = "UNIVERSAL"
                      Else
                         If data_inf.Recordset("cl_codconv") = "CCNOS" Or data_inf.Recordset("cl_codconv") = "CCNSAM" Then
                            Xarchexelcar.Cells(Xlin, XCol) = "CCOU"
                         Else
                            If data_inf.Recordset("cl_codconv") = "HEVANO" Or data_inf.Recordset("cl_codconv") = "EVNSAM" Then
                               Xarchexelcar.Cells(Xlin, XCol) = "H.EVANGELICO"
                            Else
                               If data_inf.Recordset("cl_codconv") = "GANOS" Or data_inf.Recordset("cl_codconv") = "CASANO" Or data_inf.Recordset("cl_codconv") = "CASNSA" Then
                                  Xarchexelcar.Cells(Xlin, XCol) = "CASA DE GALICIA"
                               Else
                                  Xarchexelcar.Cells(Xlin, XCol) = "SIN DATO"
                               End If
                            End If
                         End If
                      End If
                   End If
                   XCol = XCol + 1
                   Xarchexelcar.Cells(Xlin, XCol) = data_inf.Recordset("cl_codigo")
                   XCol = XCol + 1
                   Xarchexelcar.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
                   XCol = XCol + 1
                   Xarchexelcar.Cells(Xlin, XCol) = Trim(str(data_inf.Recordset("cl_cedula"))) & "-" & Trim(str(data_inf.Recordset("cl_codced")))
                   XCol = XCol + 1
                   If IsNull(data_inf.Recordset("cl_direcci")) = False Then
                      If Trim(data_inf.Recordset("cl_direcci")) <> "" Then
                         Xarchexelcar.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
                      Else
                         Xarchexelcar.Cells(Xlin, XCol) = "Sin Datos direcc"
                      End If
                   Else
                      Xarchexelcar.Cells(Xlin, XCol) = "Sin datos direcc"
                   End If
                   XCol = XCol + 1
                   If IsNull(data_inf.Recordset("cl_zona")) = False Then
                      Xarchexelcar.Cells(Xlin, XCol) = data_inf.Recordset("cl_zona")
                   Else
                      Xarchexelcar.Cells(Xlin, XCol) = "Sin Zona"
                   End If
                   XCol = XCol + 1
                   If IsNull(data_inf.Recordset("cl_dpto")) = False Then
                      Xarchexelcar.Cells(Xlin, XCol) = data_inf.Recordset("cl_dpto")
                   Else
                      Xarchexelcar.Cells(Xlin, XCol) = "Sin Celular"
                   End If
                   XCol = XCol + 1
                   If IsNull(data_inf.Recordset("cl_telefon")) = False Then
                      Xarchexelcar.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
                   Else
                      Xarchexelcar.Cells(Xlin, XCol) = "Sin Telef"
                   End If
                   XCol = XCol + 1
                   If IsNull(data_inf.Recordset("cl_cantdia")) = False Then
                      Xarchexelcar.Cells(Xlin, XCol) = data_inf.Recordset("cl_cantdia")
                   Else
                      Xarchexelcar.Cells(Xlin, XCol) = 0
                   End If
                   XCol = XCol + 1
                   If IsNull(data_inf.Recordset("cl_fultvta")) = False Then
                      Xarchexelcar.Cells(Xlin, XCol) = CDate(data_inf.Recordset("cl_fultvta"))
                   Else
                      Xarchexelcar.Cells(Xlin, XCol) = 0
                   End If
                   
                   XCol = XCol + 1
              '     Xarchexelcar.Range("A" & Trim(Str(Xlin)), "H" & Trim(Str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
                   data_inf.Recordset.MoveNext
                   Xlin = Xlin + 1
                   XCol = 1
                Loop
                Xlibexelcar.Save
        '            Xlibexel.Application
                Xlibexelcar.Close
                Xobjexelcar.Quit
                MsgBox "El archivo ha sido guardado en la carpeta PLANILLAS del disco C", vbInformation
                
                Xlabrir3.Workbooks.Open Xarchtex, , False
                Xlabrir3.Visible = True
                Xlabrir3.WindowState = xlMaximized
    
    '                Shell frm_menu.data_usuac.Recordset("destino") & "excel.exe " & Xarchtex, vbMaximizedFocus
             End If
          
          
          
          End If
          frm_infcartasm.MousePointer = 0
          MsgBox "Proceso terminado"
          If data_inf.Recordset.RecordCount > 0 Then
             data_inf.Refresh
             If Combo2.ListIndex = 0 Then
    '            cr1.ReportFileName = App.path & "\infcartasm.rpt"
    '            cr1.ReportTitle = "Informe de AVISOS para realizar carta mutual desde:" & md.Text & " hasta:" & mh.Text
    '            cr1.Action = 1
             Else
                If Combo2.ListIndex = 1 Then
    '               cr1.ReportFileName = App.path & "\infcartasm.rpt"
    '               cr1.ReportTitle = "Informe de CARTAS MUTUALES REALIZADAS desde:" & md.Text & " hasta:" & mh.Text
    '               cr1.Action = 1
                Else
                   If Combo2.ListIndex = 2 Then
                   Else
                      If Combo2.ListIndex = 4 Then
                         cr1.ReportFileName = App.path & "\infcartasm.rpt"
                         cr1.ReportTitle = "Informe de clientes que se negaron a firmar carta "
                         cr1.Action = 1
                      Else
                         cr1.ReportFileName = App.path & "\infcartasm.rpt"
                         cr1.ReportTitle = "Informe de clientes con SERVICIOS RESTRINGIDOS "
                         cr1.Action = 1
                      End If
                   End If
                End If
             End If
          End If
       Else
          frm_infcartasm.MousePointer = 0
          MsgBox "No existen registros con las fechas y mutualista seleccionada"
       End If
    End If

End If

'Exit Sub

'Quepasacartas:
'              If Err.Number = 3155 Then
'                 MsgBox "Hay un error en los datos, VERIFIQUE!!"
'                 If Combo2.ListIndex = 2 Then
'                    Xlibexelcar.Close
'                    Xobjexelcar.Quit
'                 End If
'              Else
'                 MsgBox "Error al crear el informe, verifique datos!"
'                 If Combo2.ListIndex = 2 Then
'                    Xlibexelcar.Close
'                    Xobjexelcar.Quit
'                 End If
'              End If
              
              
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
Dim Xobjexelcar As Excel.Application
Dim Xlibexelcar As Excel.Workbook
Dim Xarchexelcar As New Excel.Worksheet
Dim Lacedqueviene As String
Dim Xlabrir3 As New Excel.Application

Dim XCol, Xlin, Xnrocan, Xcolfija As Long
Dim Xarchtex As String
Dim Xtitmutual, Comofigura As String
Dim Xlamatque As Double
Xlamatque = 0
''SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','SMINA','UNNSAM','EVNSAM','CCNSAM','CASNSA')
frm_infcartasm.MousePointer = 11
       Dim MiBaseact As Database
       Dim Unasesact As Workspace
       Set Unasesact = Workspaces(0)
       Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")
       MiBaseact.Execute "Delete * from infcli"
       data_inf.RecordSource = "infcli"
       data_inf.Refresh

If Combo2.Text = "CARTAS FACTURADAS" Then
   If Combo1.Text = "CCOU" Then
      data_cli.RecordSource = "select linmmdd.fecha,linmmdd.convenio,linmmdd.cod_prod,linmmdd.nom_prod,linmmdd.base,linmmdd.cod_cli," & _
      "linmmdd.nom_cli,linmmdd.factura,clientes.cl_codigo,clientes.estado from linmmdd inner join " & _
      "clientes on linmmdd.cod_cli=clientes.cl_codigo where linmmdd.fecha >='" & Format(md.Text, "yyyy/mm/dd") & "' and linmmdd.fecha <='" & Format(mh.Text, "yyyy/mm/dd") & "' and linmmdd.cod_prod in (805) and clientes.estado not in (2,3)"
      data_cli.Refresh
   Else
      If Combo1.Text = "SMI" Then
         data_cli.RecordSource = "select linmmdd.fecha,linmmdd.convenio,linmmdd.cod_prod,linmmdd.nom_prod,linmmdd.base,linmmdd.cod_cli," & _
         "linmmdd.nom_cli,linmmdd.factura,clientes.cl_codigo,clientes.estado from linmmdd inner join " & _
         "clientes on linmmdd.cod_cli=clientes.cl_codigo where linmmdd.fecha >='" & Format(md.Text, "yyyy/mm/dd") & "' and linmmdd.fecha <='" & Format(mh.Text, "yyyy/mm/dd") & "' and linmmdd.cod_prod in (802) and clientes.estado not in (2,3)"
         data_cli.Refresh
      Else
         If Combo1.Text = "H.EVANGELICO" Then
            data_cli.RecordSource = "select linmmdd.fecha,linmmdd.convenio,linmmdd.cod_prod,linmmdd.nom_prod,linmmdd.base,linmmdd.cod_cli," & _
            "linmmdd.nom_cli,linmmdd.factura,clientes.cl_codigo,clientes.estado from linmmdd inner join " & _
            "clientes on linmmdd.cod_cli=clientes.cl_codigo where linmmdd.fecha >='" & Format(md.Text, "yyyy/mm/dd") & "' and linmmdd.fecha <='" & Format(mh.Text, "yyyy/mm/dd") & "' and linmmdd.cod_prod in (804) and clientes.estado not in (2,3)"
            data_cli.Refresh
         Else
            If Combo1.Text = "UNIVERSAL" Then
               data_cli.RecordSource = "select linmmdd.fecha,linmmdd.convenio,linmmdd.cod_prod,linmmdd.nom_prod,linmmdd.base,linmmdd.cod_cli," & _
               "linmmdd.nom_cli,linmmdd.factura,clientes.cl_codigo,clientes.estado from linmmdd inner join " & _
               "clientes on linmmdd.cod_cli=clientes.cl_codigo where linmmdd.fecha >='" & Format(md.Text, "yyyy/mm/dd") & "' and linmmdd.fecha <='" & Format(mh.Text, "yyyy/mm/dd") & "' and linmmdd.cod_prod in (803) and clientes.estado not in (2,3)"
               data_cli.Refresh
            Else
               If Combo1.Text = "CASA DE GALICIA" Then
                  data_cli.RecordSource = "select linmmdd.fecha,linmmdd.convenio,linmmdd.cod_prod,linmmdd.nom_prod,linmmdd.base,linmmdd.cod_cli," & _
                  "linmmdd.nom_cli,linmmdd.factura,clientes.cl_codigo,clientes.estado from linmmdd inner join " & _
                  "clientes on linmmdd.cod_cli=clientes.cl_codigo where linmmdd.fecha >='" & Format(md.Text, "yyyy/mm/dd") & "' and linmmdd.fecha <='" & Format(mh.Text, "yyyy/mm/dd") & "' and linmmdd.cod_prod in (806) and clientes.estado not in (2,3)"
                  data_cli.Refresh
               Else
                  data_cli.RecordSource = "select linmmdd.fecha,linmmdd.convenio,linmmdd.cod_prod,linmmdd.nom_prod,linmmdd.base,linmmdd.cod_cli," & _
                  "linmmdd.nom_cli,linmmdd.factura,clientes.cl_codigo,clientes.estado from linmmdd inner join " & _
                  "clientes on linmmdd.cod_cli=clientes.cl_codigo where linmmdd.fecha >='" & Format(md.Text, "yyyy/mm/dd") & "' and linmmdd.fecha <='" & Format(mh.Text, "yyyy/mm/dd") & "' and linmmdd.cod_prod in (802,803,804,805,806) and clientes.estado not in (2,3)"
                  data_cli.Refresh
               End If
            End If
         End If
      End If
   End If
   If data_cli.Recordset.RecordCount > 0 Then
      data_cli.Recordset.MoveFirst
      Do While Not data_cli.Recordset.EOF
         data_inf.Recordset.AddNew
         data_inf.Recordset("cl_fecing") = data_cli.Recordset("fecha")
         data_inf.Recordset("cl_codconv") = data_cli.Recordset("convenio")
         data_inf.Recordset("cl_codced") = data_cli.Recordset("cod_prod")
         data_inf.Recordset("cl_nombre") = data_cli.Recordset("nom_prod")
         data_inf.Recordset("cl_cedula") = data_cli.Recordset("base")
         data_inf.Recordset("cl_codigo") = data_cli.Recordset("cod_cli")
         data_inf.Recordset("cl_apellid") = data_cli.Recordset("nom_cli")
         data_cli2.RecordSource = "select * from linmmdd_afil where factura =" & data_cli.Recordset("factura")
         data_cli2.Refresh
         If data_cli2.Recordset.RecordCount > 0 Then
            If IsNull(data_cli2.Recordset("codfunc")) = False Then
                data_inf.Recordset("cl_nro_sup") = Val(data_cli2.Recordset("codfunc"))
                data_inf.Recordset("cl_nomvend") = data_cli2.Recordset("nombre")
            End If
         Else
            data_inf.Recordset("cl_nro_sup") = 0
            data_inf.Recordset("cl_nomvend") = "Sin Datos"
         End If
         data_inf.Recordset.Update
         data_cli.Recordset.MoveNext
      Loop
      frm_infcartasm.MousePointer = 0
      MsgBox "Proceso terminado"
      data_inf.RecordSource = "select * from infcli"
      data_inf.Refresh
      cr1.ReportFileName = App.path & "\infvtascartasp.rpt"
      cr1.ReportTitle = "Informe de cartas facturadas por promotor/mutual desde: " & md.Text & " hasta: " & mh.Text
      cr1.Action = 1
      
      cr2.ReportFileName = App.path & "\infvtascartas.rpt"
      cr2.ReportTitle = "Informe de cartas facturadas por promotor desde: " & md.Text & " hasta: " & mh.Text
      cr2.Action = 1
              
   End If
   
End If

If Combo2.Text = "LLAMADOS" Then
   If Combo1.Text = "CCOU" Then
      data_cli.RecordSource = "select cartasnosapp.fecha,cartasnosapp.usuario,cartasnosapp.cedula,cartasnosapp.nombre,cartasnosapp.codigo," & _
      "cartasnosapp.convenio,cartasnosapp.opcion,cartasnosapp.nrolla,llamado.nrolla,llamado.codmot,llamado.obsmot," & _
      "llamado.cancela,llamado.motmov,llamado.movilpas,llamado.telef,llamado.matric from cartasnosapp inner join llamado on " & _
      "cartasnosapp.nrolla=llamado.nrolla where cartasnosapp.fecha >='" & Format(md.Text, "yyyy/mm/dd") & "' and cartasnosapp.fecha <='" & Format(mh.Text, "yyyy/mm/dd") & "' and cartasnosapp.convenio in ('CCNOS','CCNSAM') and cartasnosapp.cedula is not null order by cartasnosapp.cedula,cartasnosapp.nombre,cartasnosapp.fecha"
      data_cli.Refresh
   Else
      If Combo1.Text = "SMI" Then
         data_cli.RecordSource = "select cartasnosapp.fecha,cartasnosapp.usuario,cartasnosapp.cedula,cartasnosapp.nombre,cartasnosapp.codigo," & _
         "cartasnosapp.convenio,cartasnosapp.opcion,cartasnosapp.nrolla,llamado.nrolla,llamado.codmot,llamado.obsmot," & _
         "llamado.cancela,llamado.motmov,llamado.movilpas,llamado.telef,llamado.matric from cartasnosapp inner join llamado on " & _
         "cartasnosapp.nrolla=llamado.nrolla where cartasnosapp.fecha >='" & Format(md.Text, "yyyy/mm/dd") & "' and cartasnosapp.fecha <='" & Format(mh.Text, "yyyy/mm/dd") & "' and cartasnosapp.convenio in ('SMIN','SMINA') and cartasnosapp.cedula is not null order by cartasnosapp.cedula,cartasnosapp.nombre,cartasnosapp.fecha"
         data_cli.Refresh
      Else
         If Combo1.Text = "H.EVANGELICO" Then
            data_cli.RecordSource = "select cartasnosapp.fecha,cartasnosapp.usuario,cartasnosapp.cedula,cartasnosapp.nombre,cartasnosapp.codigo," & _
            "cartasnosapp.convenio,cartasnosapp.opcion,cartasnosapp.nrolla,llamado.nrolla,llamado.codmot,llamado.obsmot," & _
            "llamado.cancela,llamado.motmov,llamado.movilpas,llamado.telef,llamado.matric from cartasnosapp inner join llamado on " & _
            "cartasnosapp.nrolla=llamado.nrolla where cartasnosapp.fecha >='" & Format(md.Text, "yyyy/mm/dd") & "' and cartasnosapp.fecha <='" & Format(mh.Text, "yyyy/mm/dd") & "' and cartasnosapp.convenio in ('HEVANO','EVNSAM') and cartasnosapp.cedula is not null order by cartasnosapp.cedula,cartasnosapp.fecha"
            data_cli.Refresh
         Else
            If Combo1.Text = "UNIVERSAL" Then
               data_cli.RecordSource = "select cartasnosapp.fecha,cartasnosapp.usuario,cartasnosapp.cedula,cartasnosapp.nombre,cartasnosapp.codigo," & _
               "cartasnosapp.convenio,cartasnosapp.opcion,cartasnosapp.nrolla,llamado.nrolla,llamado.codmot,llamado.obsmot," & _
               "llamado.cancela,llamado.motmov,llamado.movilpas,llamado.telef,llamado.matric from cartasnosapp inner join llamado on " & _
               "cartasnosapp.nrolla=llamado.nrolla where cartasnosapp.fecha >='" & Format(md.Text, "yyyy/mm/dd") & "' and cartasnosapp.fecha <='" & Format(mh.Text, "yyyy/mm/dd") & "' and cartasnosapp.convenio in ('UNIVS','UNNSAM') and cartasnosapp.cedula is not null order by cartasnosapp.cedula,cartasnosapp.fecha"
               data_cli.Refresh
            Else
               If Combo1.Text = "CASA DE GALICIA" Then
                  data_cli.RecordSource = "select cartasnosapp.fecha,cartasnosapp.usuario,cartasnosapp.cedula,cartasnosapp.nombre,cartasnosapp.codigo," & _
                  "cartasnosapp.convenio,cartasnosapp.opcion,cartasnosapp.nrolla,llamado.nrolla,llamado.codmot,llamado.obsmot," & _
                  "llamado.cancela,llamado.motmov,llamado.movilpas,llamado.telef,llamado.matric from cartasnosapp inner join llamado on " & _
                  "cartasnosapp.nrolla=llamado.nrolla where cartasnosapp.fecha >='" & Format(md.Text, "yyyy/mm/dd") & "' and cartasnosapp.fecha <='" & Format(mh.Text, "yyyy/mm/dd") & "' and cartasnosapp.convenio in ('GANOS','CASANO','CASNSA') and cartasnosapp.cedula is not null order by cartasnosapp.cedula,cartasnosapp.fecha"
                  data_cli.Refresh
               Else
                  data_cli.RecordSource = "select cartasnosapp.fecha,cartasnosapp.usuario,cartasnosapp.cedula,cartasnosapp.nombre,cartasnosapp.codigo," & _
                  "cartasnosapp.convenio,cartasnosapp.opcion,cartasnosapp.nrolla,llamado.nrolla,llamado.codmot,llamado.obsmot," & _
                  "llamado.cancela,llamado.motmov,llamado.movilpas,llamado.telef,llamado.matric from cartasnosapp inner join llamado on " & _
                  "cartasnosapp.nrolla=llamado.nrolla where cartasnosapp.fecha >='" & Format(md.Text, "yyyy/mm/dd") & "' and cartasnosapp.fecha <='" & Format(mh.Text, "yyyy/mm/dd") & "' and cartasnosapp.cedula is not null and cartasnosapp.convenio in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','SMINA','UNNSAM','EVNSAM','CCNSAM','CASNSA') order by cartasnosapp.cedula,cartasnosapp.fecha"
                  data_cli.Refresh
               End If
            End If
         End If
      End If
      
   End If
End If
'802 al 806 son cartas el 984 al 989 afiliaciones

If Combo2.Text = "POLICLINICAS" Then
   If Combo1.Text = "CCOU" Then
      data_cli.RecordSource = "select * from linmmdd where fecha >='" & Format(md.Text, "yyyy/mm/dd") & "' and fecha <='" & Format(mh.Text, "yyyy/mm/dd") & "' and convenio in ('CCNOS','CCNSAM') " & _
      "and cod_prod not in (800,802,803,804,805,806,984,985,986,987,989,991,992,993,994,995,996,997,999,8000) and base not in (19) order by cod_cli,fecha"
      data_cli.Refresh
   Else
      If Combo1.Text = "SMI" Then
         data_cli.RecordSource = "select * from linmmdd where fecha >='" & Format(md.Text, "yyyy/mm/dd") & "' and fecha <='" & Format(mh.Text, "yyyy/mm/dd") & "' and convenio in ('SMIN','SMINA') " & _
         "and cod_prod not in (800,802,803,804,805,806,984,985,986,987,989,991,992,993,994,995,996,997,999,8000) and base not in (19) order by cod_cli,fecha"
         data_cli.Refresh
      Else
         If Combo1.Text = "H.EVANGELICO" Then
            data_cli.RecordSource = "select * from linmmdd where fecha >='" & Format(md.Text, "yyyy/mm/dd") & "' and fecha <='" & Format(mh.Text, "yyyy/mm/dd") & "' and convenio in ('HEVANO','EVNSAM') " & _
            "and cod_prod not in (800,802,803,804,805,806,984,985,986,987,989,991,992,993,994,995,996,997,999,8000) and base not in (19) order by cod_cli,fecha"
            data_cli.Refresh
         Else
            If Combo1.Text = "UNIVERSAL" Then
               data_cli.RecordSource = "select * from linmmdd where fecha >='" & Format(md.Text, "yyyy/mm/dd") & "' and fecha <='" & Format(mh.Text, "yyyy/mm/dd") & "' and convenio in ('UNIVS','UNNSAM') " & _
               "and cod_prod not in (800,802,803,804,805,806,984,985,986,987,989,991,992,993,994,995,996,997,999,8000) and base not in (19) order by cod_cli,fecha"
               data_cli.Refresh
            Else
               If Combo1.Text = "CASA DE GALICIA" Then
                  data_cli.RecordSource = "select * from linmmdd where fecha >='" & Format(md.Text, "yyyy/mm/dd") & "' and fecha <='" & Format(mh.Text, "yyyy/mm/dd") & "' and convenio in ('GANOS','CASANO','CASNSA') " & _
                  "and cod_prod not in (800,802,803,804,805,806,984,985,986,987,989,991,992,993,994,995,996,997,999,8000) and base not in (19) order by cod_cli,fecha"
                  data_cli.Refresh
               Else
                  data_cli.RecordSource = "select * from linmmdd where fecha >='" & Format(md.Text, "yyyy/mm/dd") & "' and fecha <='" & Format(mh.Text, "yyyy/mm/dd") & "'" & _
                  "and convenio in ('SMIN','UNIVS','CCNOS','HEVANO','GANOS','CASANO','SMINA','UNNSAM','EVNSAM','CCNSAM','CASNSA') and cod_prod not in (800,802,803,804,805,806,984,985,986,987,989,991,992,993,994,995,996,997,999,8000) and base not in (19) order by cod_cli,fecha"
                  data_cli.Refresh
               End If
            End If
         End If
      End If
   End If
End If

If Combo2.Text = "CARTAS FACTURADAS" Then

Else
    If Combo2.Text = "RESUMEN DE CARTAS" Then
       Dim Xtotc1, Xtotc2, Xtotc3, Xtotc4, Xtotc5, XcartasF As Integer
       Xtotc1 = 0
       Xtotc2 = 0
       Xtotc3 = 0
       Xtotc4 = 0
       Xtotc5 = 0
       XcartasF = 0
       data_cli.RecordSource = "select * from clientes where cl_decuota in (1,2,3,4,5) and fecha_reac >='" & Format(md.Text, "yyyy/mm/dd") & "' and fecha_reac <='" & Format(mh.Text, "yyyy/mm/dd") & "'"
       data_cli.Refresh
    
       If data_cli.Recordset.RecordCount > 0 Then
          data_cli.Recordset.MoveLast
          Xnrocan = data_cli.Recordset.RecordCount + 15
          data_cli.Recordset.MoveFirst
          Set Xobjexelcar = New Excel.Application
          Set Xlibexelcar = Xobjexelcar.Workbooks.Add
          Set Xarchexelcar = Xlibexelcar.Worksheets.Add
          Xlin = 1
          XCol = 1
          Xarchexelcar.Name = "resumen"
          Xlibexelcar.SaveAs ("C:\planillas\resumen_cartas.xls")
          Xarchtex = "C:\planillas\resumen_cartas.xls"
          Xarchexelcar.Cells(Xlin, XCol) = "SAPP S.A.  -- DPTO.TI"
          Xlin = Xlin + 1
          XCol = 1
          Xarchexelcar.Range("A2", "C3").Font.Size = 16
          Xarchexelcar.Cells(Xlin, XCol) = "RESUMEN DE CONTROL CARTAS MUTUALES FECHAS:" & md.Text & " " & mh.Text
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
        
          Xarchexelcar.Cells(Xlin, XCol) = "FECHA_REG"
          XCol = XCol + 1
          Xarchexelcar.Range("B" & Trim(str(Xlin))).ColumnWidth = 15
          Xarchexelcar.Cells(Xlin, XCol) = "MATRICULA"
          XCol = XCol + 1
          Xarchexelcar.Range("C" & Trim(str(Xlin))).ColumnWidth = 40
          Xarchexelcar.Cells(Xlin, XCol) = "NOMBRE"
          XCol = XCol + 1
          Xarchexelcar.Range("D" & Trim(str(Xlin))).ColumnWidth = 12
          Xarchexelcar.Cells(Xlin, XCol) = "CONVENIO"
          XCol = XCol + 1
          Xarchexelcar.Range("E" & Trim(str(Xlin))).ColumnWidth = 20
          Xarchexelcar.Cells(Xlin, XCol) = "REGISTRO"
          XCol = XCol + 1
          Xarchexelcar.Range("F" & Trim(str(Xlin))).ColumnWidth = 20
          Xarchexelcar.Cells(Xlin, XCol) = "ZONA"
          XCol = XCol + 1
          Xarchexelcar.Range("G" & Trim(str(Xlin))).ColumnWidth = 15
          Xarchexelcar.Cells(Xlin, XCol) = "CELULAR"
          XCol = XCol + 1
          Xarchexelcar.Range("H" & Trim(str(Xlin))).ColumnWidth = 15
          Xarchexelcar.Cells(Xlin, XCol) = "TELEFONO"
          XCol = XCol + 1
          Xarchexelcar.Range("I" & Trim(str(Xlin))).ColumnWidth = 15
          Xarchexelcar.Cells(Xlin, XCol) = "CARTA_FACTURADA"
          Xlin = Xlin + 1
          XCol = 1
          Do While Not data_cli.Recordset.EOF
             If IsNull(data_cli.Recordset("fecha_reac")) = False Then
                Xarchexelcar.Cells(Xlin, XCol) = "'" & Format(data_cli.Recordset("fecha_reac"), "dd/mm/yyyy")
             Else
                Xarchexelcar.Cells(Xlin, XCol) = "Sin fecha"
             End If
             XCol = XCol + 1
             Xarchexelcar.Cells(Xlin, XCol) = data_cli.Recordset("cl_codigo")
             XCol = XCol + 1
             If IsNull(data_cli.Recordset("cl_apellid")) = False Then
                Xarchexelcar.Cells(Xlin, XCol) = data_cli.Recordset("cl_apellid")
             Else
                Xarchexelcar.Cells(Xlin, XCol) = "NN"
             End If
             XCol = XCol + 1
             If IsNull(data_cli.Recordset("cl_codconv")) = False Then
                Xarchexelcar.Cells(Xlin, XCol) = data_cli.Recordset("cl_codconv")
             Else
                Xarchexelcar.Cells(Xlin, XCol) = "NN"
             End If
             XCol = XCol + 1
             If data_cli.Recordset("cl_decuota") = 1 Then
                Xarchexelcar.Cells(Xlin, XCol) = "Aviso firmar carta"
                Xtotc1 = Xtotc1 + 1
             Else
                If data_cli.Recordset("cl_decuota") = 2 Then
                   Xarchexelcar.Cells(Xlin, XCol) = "Se recibe carta"
                   Xtotc2 = Xtotc2 + 1
                Else
                   If data_cli.Recordset("cl_decuota") = 3 Then
                      Xarchexelcar.Cells(Xlin, XCol) = "Se niega a firmar"
                      Xtotc3 = Xtotc3 + 1
                   Else
                      If data_cli.Recordset("cl_decuota") = 4 Then
                         Xarchexelcar.Cells(Xlin, XCol) = "Rechazada"
                         Xtotc4 = Xtotc4 + 1
                      Else
                         If data_cli.Recordset("cl_decuota") = 5 Then
                            Xarchexelcar.Cells(Xlin, XCol) = "Pendiente"
                            Xtotc5 = Xtotc5 + 1
                         End If
                      End If
                   End If
                End If
             End If
             XCol = XCol + 1
             If IsNull(data_cli.Recordset("cl_zona")) = False Then
                Xarchexelcar.Cells(Xlin, XCol) = data_cli.Recordset("cl_zona")
             Else
                Xarchexelcar.Cells(Xlin, XCol) = "Sin zona"
             End If
             XCol = XCol + 1
             If IsNull(data_cli.Recordset("cl_dpto")) = False Then
                Xarchexelcar.Cells(Xlin, XCol) = data_cli.Recordset("cl_dpto")
             Else
                Xarchexelcar.Cells(Xlin, XCol) = "Sin Cel"
             End If
             XCol = XCol + 1
             If IsNull(data_cli.Recordset("cl_telefon")) = False Then
                Xarchexelcar.Cells(Xlin, XCol) = data_cli.Recordset("cl_telefon")
             Else
                Xarchexelcar.Cells(Xlin, XCol) = "Sin Tel"
             End If
             XCol = XCol + 1
             data_cli2.RecordSource = "select * from linmmdd where cod_cli =" & data_cli.Recordset("cl_codigo") & " and cod_prod in (802,803,804,805,806)"
             data_cli2.Refresh
             If data_cli2.Recordset.RecordCount > 0 Then
                XcartasF = XcartasF + 1
                Xarchexelcar.Cells(Xlin, XCol) = "'" & Format(data_cli2.Recordset("fecha"), "dd/mm/yyyy")
             End If
             data_cli.Recordset.MoveNext
             Xlin = Xlin + 1
             XCol = 1
          Loop
          Xlin = Xlin + 1
          XCol = 3
          Xarchexelcar.Cells(Xlin, XCol) = "TOTAL AVISOS FIRMAR CARTA:"
          XCol = 4
          Xarchexelcar.Cells(Xlin, XCol) = Xtotc1
          Xlin = Xlin + 1
          XCol = 3
          Xarchexelcar.Cells(Xlin, XCol) = "TOTAL SE RECIBE CARTA:"
          XCol = 4
          Xarchexelcar.Cells(Xlin, XCol) = Xtotc2
          Xlin = Xlin + 1
          XCol = 3
          Xarchexelcar.Cells(Xlin, XCol) = "TOTAL NEGATIVAS:"
          XCol = 4
          Xarchexelcar.Cells(Xlin, XCol) = Xtotc3
          Xlin = Xlin + 1
          XCol = 3
          Xarchexelcar.Cells(Xlin, XCol) = "TOTAL RECHAZADAS:"
          XCol = 4
          Xarchexelcar.Cells(Xlin, XCol) = Xtotc4
          Xlin = Xlin + 1
          XCol = 3
          Xarchexelcar.Cells(Xlin, XCol) = "TOTAL PENDIENTES:"
          XCol = 4
          Xarchexelcar.Cells(Xlin, XCol) = Xtotc5
          Xlin = Xlin + 1
          XCol = 3
          Xarchexelcar.Cells(Xlin, XCol) = "FACTURADOS:"
          XCol = 4
          Xarchexelcar.Cells(Xlin, XCol) = XcartasF
          Xlin = Xlin + 1
          
          data_cli2.RecordSource = "select * from abmsocio where fecha >='" & Format(md.Text, "yyyy/mm/dd") & "' and fecha <='" & Format(mh.Text, "yyyy/mm/dd") & "' and cl_motivo ='" & "INGRESO A PADRON" & "'"
          data_cli2.Refresh
          If data_cli2.Recordset.RecordCount > 0 Then
             data_cli2.Recordset.MoveLast
             Xlin = Xlin + 1
             XCol = 3
             Xarchexelcar.Cells(Xlin, XCol) = "TOTAL INGRESOS A PADRÓN:"
             XCol = 4
             Xarchexelcar.Cells(Xlin, XCol) = data_cli2.Recordset.RecordCount
          End If
          
          data_cli2.RecordSource = "select * from linmmdd where fecha >='" & Format(md.Text, "yyyy/mm/dd") & "' and fecha <='" & Format(mh.Text, "yyyy/mm/dd") & "' and cod_prod in (802,803,804,805,806)"
          data_cli2.Refresh
          If data_cli2.Recordset.RecordCount > 0 Then
             data_cli2.Recordset.MoveLast
             Xlin = Xlin + 1
             XCol = 3
             Xarchexelcar.Cells(Xlin, XCol) = "TOTAL CARTAS FACTURADAS:"
             XCol = 4
             Xarchexelcar.Cells(Xlin, XCol) = data_cli2.Recordset.RecordCount
          End If
          
          frm_infcartasm.MousePointer = 0
          Xlibexelcar.Save
          Xlibexelcar.Close
          Xobjexelcar.Quit
          MsgBox "El archivo resumen_cartas.xls ha sido guardado en la carpeta PLANILLAS del disco C", vbInformation
          Xlabrir3.Workbooks.Open Xarchtex, , False
          Xlabrir3.Visible = True
          Xlabrir3.WindowState = xlMaximized
       Else
          frm_infcartasm.MousePointer = 0
          MsgBox "No hay registros"
       End If
    Else
        Comofigura = ""
        If data_cli.Recordset.RecordCount > 0 Then
           data_cli.Recordset.MoveLast
           Xnrocan = data_cli.Recordset.RecordCount + 4
           data_cli.Recordset.MoveFirst
           If Combo2.Text = "POLICLINICAS" Then
              Set Xobjexelcar = New Excel.Application
              Set Xlibexelcar = Xobjexelcar.Workbooks.Add
              Set Xarchexelcar = Xlibexelcar.Worksheets.Add
              Xlin = 1
              XCol = 1
              Xarchexelcar.Name = "controles"
              Xlibexelcar.SaveAs ("C:\planillas\control_cartas.xls")
              Xarchtex = "C:\planillas\control_cartas.xls"
              Xarchexelcar.Cells(Xlin, XCol) = "SAPP S.A.  -- DPTO.TI"
              Xlin = Xlin + 1
              XCol = 1
              Xarchexelcar.Range("A2", "C3").Font.Size = 16
              Xarchexelcar.Cells(Xlin, XCol) = Combo2.Text & " " & Combo1.Text & " NOSAPP FECHAS:" & md.Text & " " & mh.Text
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
              Xarchexelcar.Cells(Xlin, XCol) = "CEDULA"
              XCol = XCol + 1
              Xarchexelcar.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
              Xarchexelcar.Cells(Xlin, XCol) = "MATRICULA"
              XCol = XCol + 1
              Xarchexelcar.Range("E" & Trim(str(Xlin))).ColumnWidth = 40
              Xarchexelcar.Cells(Xlin, XCol) = "NOMBRE"
              XCol = XCol + 1
              Xarchexelcar.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
              Xarchexelcar.Cells(Xlin, XCol) = "CONVENIO"
              XCol = XCol + 1
              Xarchexelcar.Range("G" & Trim(str(Xlin))).ColumnWidth = 30
              Xarchexelcar.Cells(Xlin, XCol) = "SERVICIO"
              XCol = XCol + 1
              Xarchexelcar.Range("H" & Trim(str(Xlin))).ColumnWidth = 6
              Xarchexelcar.Cells(Xlin, XCol) = "BASE"
              XCol = XCol + 1
              Xarchexelcar.Range("I" & Trim(str(Xlin))).ColumnWidth = 20
              Xarchexelcar.Cells(Xlin, XCol) = "ZONA"
              XCol = XCol + 1
              Xarchexelcar.Range("J" & Trim(str(Xlin))).ColumnWidth = 20
              Xarchexelcar.Cells(Xlin, XCol) = "CARTA_FACTURADA"
              XCol = XCol + 1
              Xarchexelcar.Range("K" & Trim(str(Xlin))).ColumnWidth = 20
              Xarchexelcar.Cells(Xlin, XCol) = "PROMOTOR"
              XCol = XCol + 1
              Xarchexelcar.Range("L" & Trim(str(Xlin))).ColumnWidth = 25
              Xarchexelcar.Cells(Xlin, XCol) = "EN SAPP FIGURA"
              Xlin = Xlin + 1
              XCol = 1
              Xlamatque = data_cli.Recordset("cod_cli")
              
              Do While Not data_cli.Recordset.EOF
                 data_cli.Recordset.MoveNext
                 If data_cli.Recordset.EOF = True Then
                    data_cli.Recordset.MovePrevious
                    Xlamatque = 87
                 End If
                 If Xlamatque = data_cli.Recordset("cod_cli") Then
                    data_cli.Recordset.MovePrevious
                 Else
                    If Xlamatque = 87 Then
                    Else
                       data_cli.Recordset.MovePrevious
                    End If
                    If IsNull(data_cli.Recordset("fecha")) = False Then
                       Xarchexelcar.Cells(Xlin, XCol) = "'" & Format(data_cli.Recordset("fecha"), "dd/mm/yyyy")
                    Else
                       Xarchexelcar.Cells(Xlin, XCol) = "Sin fecha"
                    End If
                    XCol = XCol + 1
                    If IsNull(data_cli.Recordset("operador")) = False Then
                       Xarchexelcar.Cells(Xlin, XCol) = data_cli.Recordset("operador")
                    Else
                       Xarchexelcar.Cells(Xlin, XCol) = "s/d"
                    End If
                    XCol = XCol + 1
                    If IsNull(data_cli.Recordset("ced_socio")) = False Then
                       Xarchexelcar.Cells(Xlin, XCol) = Trim(str(data_cli.Recordset("ced_socio"))) & "-" & Trim(str(data_cli.Recordset("fact")))
                    Else
                       Xarchexelcar.Cells(Xlin, XCol) = "0-0"
                    End If
                    XCol = XCol + 1
                    If IsNull(data_cli.Recordset("cod_cli")) = False Then
                       Xarchexelcar.Cells(Xlin, XCol) = data_cli.Recordset("cod_cli")
                    Else
                       Xarchexelcar.Cells(Xlin, XCol) = 0
                    End If
                    XCol = XCol + 1
                    If IsNull(data_cli.Recordset("nom_cli")) = False Then
                       Xarchexelcar.Cells(Xlin, XCol) = data_cli.Recordset("nom_cli")
                    Else
                       Xarchexelcar.Cells(Xlin, XCol) = "NN"
                    End If
                    XCol = XCol + 1
                    If IsNull(data_cli.Recordset("convenio")) = False Then
                       Xarchexelcar.Cells(Xlin, XCol) = data_cli.Recordset("convenio")
                    End If
                    XCol = XCol + 1
                    If IsNull(data_cli.Recordset("nom_prod")) = False Then
                       Xarchexelcar.Cells(Xlin, XCol) = data_cli.Recordset("nom_prod")
                    End If
                    XCol = XCol + 1
                    If IsNull(data_cli.Recordset("base")) = False Then
                       Xarchexelcar.Cells(Xlin, XCol) = data_cli.Recordset("base")
                    Else
                       Xarchexelcar.Cells(Xlin, XCol) = 0
                    End If
                    XCol = XCol + 1
                    data_cli2.RecordSource = "select cl_codigo,estado,cl_codconv,cl_zona from clientes where cl_codigo=" & data_cli.Recordset("cod_cli")
                    data_cli2.Refresh
                    If data_cli2.Recordset.RecordCount > 0 Then
                       If IsNull(data_cli2.Recordset("cl_zona")) = False Then
                          Xarchexelcar.Cells(Xlin, XCol) = data_cli2.Recordset("cl_zona")
                       End If
                       If data_cli2.Recordset("estado") = 2 Or data_cli2.Recordset("estado") = 3 Then
                          Comofigura = "BAJA --" & data_cli2.Recordset("cl_codconv")
                       Else
                          Comofigura = "ACTIVO --" & data_cli2.Recordset("cl_codconv")
                       End If
                    Else
                       Comofigura = ""
                    End If
                    XCol = XCol + 1
                    data_cli2.RecordSource = "select cod_cli,cod_prod,fecha from linmmdd where fecha >='" & Format("2020-01-01", "yyyy-mm-dd") & "' and cod_cli =" & data_cli.Recordset("cod_cli") & " and cod_prod in (802,803,804,805,806) order by fecha DESC"
                    data_cli2.Refresh
                    If data_cli2.Recordset.RecordCount > 0 Then
                       Xarchexelcar.Cells(Xlin, XCol) = "'" & Format(data_cli2.Recordset("fecha"), "dd/mm/yyyy")
                    End If
                    XCol = XCol + 1
                    data_cli2.RecordSource = "select linmmdd.fecha,linmmdd.cod_cli,linmmdd.cod_prod,linmmdd.factura,linmmdd_afil.factura,linmmdd_afil.nombre from linmmdd inner join linmmdd_afil on linmmdd.factura=linmmdd_afil.factura where linmmdd.cod_cli =" & data_cli.Recordset("cod_cli") & " and linmmdd.cod_prod in (802,803,804,805,806)"
                    data_cli2.Refresh
                    If data_cli2.Recordset.RecordCount > 0 Then
                       If IsNull(data_cli2.Recordset("nombre")) = False Then
                          Xarchexelcar.Cells(Xlin, XCol) = data_cli2.Recordset("nombre")
                       Else
                          Xarchexelcar.Cells(Xlin, XCol) = "Sin promotor"
                       End If
                       XCol = XCol + 1
                    Else
                       XCol = XCol + 1
                    End If
                    If Trim(Comofigura) <> "" Then
                       Xarchexelcar.Cells(Xlin, XCol) = Comofigura
                    End If
                    Xlin = Xlin + 1
                    XCol = 1
                 End If
                 data_cli.Recordset.MoveNext
                 If Xlamatque = 87 Then
                 Else
                    Xlamatque = data_cli.Recordset("cod_cli")
                 End If
              Loop
              frm_infcartasm.MousePointer = 0
               
              Xlibexelcar.Save
              Xlibexelcar.Close
              Xobjexelcar.Quit
              MsgBox "El archivo control_cartas.xls ha sido guardado en la carpeta PLANILLAS del disco C", vbInformation
              Xlabrir3.Workbooks.Open Xarchtex, , False
              Xlabrir3.Visible = True
              Xlabrir3.WindowState = xlMaximized
           Else
               Set Xobjexelcar = New Excel.Application
               Set Xlibexelcar = Xobjexelcar.Workbooks.Add
               Set Xarchexelcar = Xlibexelcar.Worksheets.Add
               Xlin = 1
               XCol = 1
               Xarchexelcar.Name = "controles"
               Xlibexelcar.SaveAs ("C:\planillas\control_cartas.xls")
               Xarchtex = "C:\planillas\control_cartas.xls"
               Xarchexelcar.Cells(Xlin, XCol) = "SAPP S.A.  -- DPTO.TI"
               Xlin = Xlin + 1
               XCol = 1
               Xarchexelcar.Range("A2", "C3").Font.Size = 16
               Xarchexelcar.Cells(Xlin, XCol) = Combo2.Text & " " & Combo1.Text & " NOSAPP FECHAS:" & md.Text & " " & mh.Text
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
               Xarchexelcar.Cells(Xlin, XCol) = "DOCUMENTO"
               XCol = XCol + 1
               Xarchexelcar.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
               Xarchexelcar.Cells(Xlin, XCol) = "MATRICULA"
               XCol = XCol + 1
               Xarchexelcar.Range("E" & Trim(str(Xlin))).ColumnWidth = 40
               Xarchexelcar.Cells(Xlin, XCol) = "NOMBRE"
               XCol = XCol + 1
               Xarchexelcar.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
               Xarchexelcar.Cells(Xlin, XCol) = "CONVENIO"
               XCol = XCol + 1
               Xarchexelcar.Range("G" & Trim(str(Xlin))).ColumnWidth = 12
               Xarchexelcar.Cells(Xlin, XCol) = "REALIZA_CARTA"
               XCol = XCol + 1
               Xarchexelcar.Range("H" & Trim(str(Xlin))).ColumnWidth = 15
               Xarchexelcar.Cells(Xlin, XCol) = "LLAMADO Nro."
               XCol = XCol + 1
               Xarchexelcar.Range("I" & Trim(str(Xlin))).ColumnWidth = 5
               Xarchexelcar.Cells(Xlin, XCol) = "CLAVE"
               XCol = XCol + 1
               Xarchexelcar.Range("J" & Trim(str(Xlin))).ColumnWidth = 20
               Xarchexelcar.Cells(Xlin, XCol) = "MOTIVO"
               XCol = XCol + 1
               Xarchexelcar.Range("K" & Trim(str(Xlin))).ColumnWidth = 7
               Xarchexelcar.Cells(Xlin, XCol) = "MOVIL"
               XCol = XCol + 1
               Xarchexelcar.Range("L" & Trim(str(Xlin))).ColumnWidth = 15
               Xarchexelcar.Cells(Xlin, XCol) = "ZONA"
               XCol = XCol + 1
               Xarchexelcar.Range("M" & Trim(str(Xlin))).ColumnWidth = 18
               Xarchexelcar.Cells(Xlin, XCol) = "CARTA_FACTURADA"
               XCol = XCol + 1
               Xarchexelcar.Range("N" & Trim(str(Xlin))).ColumnWidth = 20
               Xarchexelcar.Cells(Xlin, XCol) = "PROMOTOR"
               XCol = XCol + 1
               Xarchexelcar.Range("O" & Trim(str(Xlin))).ColumnWidth = 25
               Xarchexelcar.Cells(Xlin, XCol) = "EN SAPP FIGURA" 'Ej: SI ACTIVO, NO
               
               Xlin = Xlin + 1
               XCol = 1
               If IsNull(data_cli.Recordset("cedula")) = False Then
                  Lacedqueviene = data_cli.Recordset("cedula")
               Else
                  Lacedqueviene = ""
               End If
               
               Do While Not data_cli.Recordset.EOF
                  data_cli.Recordset.MoveNext
                  If data_cli.Recordset.EOF = True Then
                     data_cli.Recordset.MovePrevious
                     Lacedqueviene = "fin"
                  End If
                  If Trim(Lacedqueviene) = Trim(data_cli.Recordset("cedula")) Then
                     data_cli.Recordset.MovePrevious
                  Else
                     If Lacedqueviene = "fin" Then
                     Else
                        data_cli.Recordset.MovePrevious
                     End If
                     If IsNull(data_cli.Recordset("fecha")) = False Then
                        Xarchexelcar.Cells(Xlin, XCol) = "'" & Format(data_cli.Recordset("fecha"), "dd/mm/yyyy")
                     Else
                        Xarchexelcar.Cells(Xlin, XCol) = "Sin fecha"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_cli.Recordset("usuario")) = False Then
                        Xarchexelcar.Cells(Xlin, XCol) = data_cli.Recordset("usuario")
                     Else
                        Xarchexelcar.Cells(Xlin, XCol) = "s/d"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_cli.Recordset("cedula")) = False Then
                        Xarchexelcar.Cells(Xlin, XCol) = data_cli.Recordset("cedula")
                     Else
                        Xarchexelcar.Cells(Xlin, XCol) = "Sin CED"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_cli.Recordset("matric")) = False Then
                        Xarchexelcar.Cells(Xlin, XCol) = data_cli.Recordset("matric")
                     Else
                        Xarchexelcar.Cells(Xlin, XCol) = 0
                     End If
                     XCol = XCol + 1
                     If IsNull(data_cli.Recordset("nombre")) = False Then
                        Xarchexelcar.Cells(Xlin, XCol) = data_cli.Recordset("nombre")
                     Else
                        data_cli2.RecordSource = "select * from llamado where nrolla =" & data_cli.Recordset("nrolla")
                        data_cli2.Refresh
                        If data_cli2.Recordset.RecordCount > 0 Then
                           If IsNull(data_cli2.Recordset("nombre")) = False Then
                              Xarchexelcar.Cells(Xlin, XCol) = data_cli2.Recordset("nombre")
                           Else
                              Xarchexelcar.Cells(Xlin, XCol) = "NN"
                           End If
                        Else
                           Xarchexelcar.Cells(Xlin, XCol) = "NN"
                        End If
                     End If
                     XCol = XCol + 1
                     If IsNull(data_cli.Recordset("convenio")) = False Then
                        Xarchexelcar.Cells(Xlin, XCol) = data_cli.Recordset("convenio")
                     Else
                        Xarchexelcar.Cells(Xlin, XCol) = 0
                     End If
                     XCol = XCol + 1
                     If IsNull(data_cli.Recordset("opcion")) = False Then
                        If IsNull(data_cli.Recordset("codigo")) = False Then
                           Xarchexelcar.Cells(Xlin, XCol) = data_cli.Recordset("opcion") & "-" & data_cli.Recordset("codigo")
                        Else
                           Xarchexelcar.Cells(Xlin, XCol) = data_cli.Recordset("opcion")
                        End If
                     Else
                        Xarchexelcar.Cells(Xlin, XCol) = "NO"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_cli.Recordset("nrolla")) = False Then
                        Xarchexelcar.Cells(Xlin, XCol) = data_cli.Recordset("nrolla")
                     Else
                        Xarchexelcar.Cells(Xlin, XCol) = 0
                     End If
                     XCol = XCol + 1
                     If IsNull(data_cli.Recordset("codmot")) = False Then
                        Xarchexelcar.Cells(Xlin, XCol) = data_cli.Recordset("codmot")
                     Else
                        Xarchexelcar.Cells(Xlin, XCol) = "V"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_cli.Recordset("obsmot")) = False Then
                        Xarchexelcar.Cells(Xlin, XCol) = data_cli.Recordset("obsmot")
                     Else
                        Xarchexelcar.Cells(Xlin, XCol) = "s/d"
                     End If
                     XCol = XCol + 1
                     If IsNull(data_cli.Recordset("movilpas")) = False Then
                        Xarchexelcar.Cells(Xlin, XCol) = data_cli.Recordset("movilpas")
                     Else
                        Xarchexelcar.Cells(Xlin, XCol) = 0
                     End If
                     XCol = XCol + 1
                     If IsNull(data_cli.Recordset("motmov")) = False Then
                        Xarchexelcar.Cells(Xlin, XCol) = data_cli.Recordset("motmov")
                     Else
                        Xarchexelcar.Cells(Xlin, XCol) = 0
                     End If
                     XCol = XCol + 1
                     If IsNull(data_cli.Recordset("matric")) = False Then
                        data_cli2.RecordSource = "select linmmdd.fecha,linmmdd.cod_cli,linmmdd.cod_prod,linmmdd.factura,linmmdd_afil.factura,linmmdd_afil.nombre from linmmdd inner join linmmdd_afil on linmmdd.factura=linmmdd_afil.factura where linmmdd.cod_cli =" & data_cli.Recordset("matric") & " and linmmdd.cod_prod in (802,803,804,805,806) order by fecha DESC"
                        data_cli2.Refresh
                        If data_cli2.Recordset.RecordCount > 0 Then
                           Xarchexelcar.Cells(Xlin, XCol) = "'" & Format(data_cli2.Recordset("fecha"), "dd/mm/yyyy")
                           XCol = XCol + 1
                           If IsNull(data_cli2.Recordset("nombre")) = False Then
                              Xarchexelcar.Cells(Xlin, XCol) = data_cli2.Recordset("nombre")
                           Else
                              Xarchexelcar.Cells(Xlin, XCol) = "Sin promotor"
                           End If
                           XCol = XCol + 1
                        Else
                           XCol = XCol + 2
                        End If
                        data_cli2.RecordSource = "select * from clientes where cl_codigo =" & data_cli.Recordset("matric")
                        data_cli2.Refresh
                        If data_cli2.Recordset.RecordCount > 0 Then
                           If data_cli2.Recordset("estado") = 2 Or data_cli2.Recordset("estado") = 3 Then
                              Xarchexelcar.Cells(Xlin, XCol) = "BAJA -Cat:" & data_cli2.Recordset("cl_codconv")
                           Else
                              Xarchexelcar.Cells(Xlin, XCol) = "ACTIVO -Cat:" & data_cli2.Recordset("cl_codconv")
                           End If
                           XCol = XCol + 1
                        Else
                           XCol = XCol + 1
                        End If
                     Else
                        XCol = XCol + 3
                     End If
                     Xlin = Xlin + 1
                     XCol = 1
                  End If
                  data_cli.Recordset.MoveNext
                  If Lacedqueviene = "fin" Then
                  Else
                     Lacedqueviene = data_cli.Recordset("cedula")
                  End If
               Loop
               frm_infcartasm.MousePointer = 0
               
               Xlibexelcar.Save
               Xlibexelcar.Close
               Xobjexelcar.Quit
               MsgBox "El archivo control_cartas.xls ha sido guardado en la carpeta PLANILLAS del disco C", vbInformation
               Xlabrir3.Workbooks.Open Xarchtex, , False
               Xlabrir3.Visible = True
               Xlabrir3.WindowState = xlMaximized
           End If
        Else
           MsgBox "No hay datos"
        End If
    End If
End If

frm_infcartasm.MousePointer = 0


End Sub

Private Sub Form_Load()

'data_inf.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"
'data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cli.ConnectionString = "dsn=" & Xconexrmt
data_cli2.ConnectionString = "dsn=" & Xconexrmt

data_inf.DatabaseName = App.path & "\informes.mdb"

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub md_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mh.SetFocus
End If

End Sub

Private Sub mh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo1.SetFocus
End If

End Sub
