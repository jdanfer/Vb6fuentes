VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_depuraficha 
   BackColor       =   &H00C0FFC0&
   Caption         =   "Proceso de eliminación de fichas"
   ClientHeight    =   4905
   ClientLeft      =   120
   ClientTop       =   420
   ClientWidth     =   8430
   Icon            =   "frm_depurafich.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   8430
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_parsec 
      Caption         =   "data_parsec"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data data_abm 
      Caption         =   "data_abm"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSAdodcLib.Adodc data2 
      Height          =   615
      Left            =   2760
      Top             =   1320
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "data2"
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
   Begin MSAdodcLib.Adodc data_datos 
      Height          =   855
      Left            =   3480
      Top             =   2040
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1508
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
      Caption         =   "data_datos"
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
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3000
      Visible         =   0   'False
      Width           =   3615
   End
   Begin MSComctlLib.ProgressBar pb1 
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   3360
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Procesar..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   360
      Picture         =   "frm_depurafich.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3960
      Width           =   2055
   End
   Begin VB.TextBox Text2 
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
      Left            =   3720
      TabIndex        =   5
      Top             =   1800
      Width           =   2415
   End
   Begin VB.TextBox Text1 
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
      Left            =   3720
      TabIndex        =   0
      Top             =   840
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H008080FF&
      Caption         =   "Salir"
      Height          =   495
      Left            =   7440
      Picture         =   "frm_depurafich.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4320
      Width           =   855
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Height          =   255
      Left            =   6600
      TabIndex        =   11
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H008080FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   10
      Top             =   2280
      Width           =   5775
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   1320
      Width           =   5775
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      Caption         =   $"frm_depurafich.frx":0F56
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   360
      TabIndex        =   6
      Top             =   2640
      Width           =   7935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Matrícula que se eliminará:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Matrícula continuará figurando:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   840
      Width           =   3255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   0
      X2              =   8400
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "ELIMINACION DE FICHAS INGRESADAS MAS DE 1 VEZ"
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
      Left            =   960
      TabIndex        =   1
      Top             =   120
      Width           =   6375
   End
End
Attribute VB_Name = "frm_depuraficha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End

End Sub

Private Sub Command2_Click()
Dim Xquehacer As String
Dim Xconvcod As String

On Error GoTo ErrDepurar

Xquehacer = ""

If Text1.Text <> "" And Text2.Text <> "" Then
    data_cli.RecordSource = "select * from clientes where cl_codigo =" & Text1.Text
    data_cli.Refresh
    If data_cli.Recordset.RecordCount > 0 Then
       data_cli.RecordSource = "select * from clientes where cl_codigo =" & Text2.Text
       data_cli.Refresh
       If data_cli.Recordset.RecordCount > 0 Then
           If Text1.Text > 0 Then
              If Text2.Text > 0 Then
                 Xquehacer = MsgBox("Está seguro de comenzar con el proceso?", vbCritical + vbYesNo, "Borrar una matrícula")
                 If Xquehacer = vbYes Then
                    DoEvents
                    frm_depuraficha.MousePointer = 11
                    Command2.Enabled = False
                    pb1.Value = 0
                    pb1.Max = 100
                    Data1.RecordSource = "infcli"
                    Data1.Refresh
                    'lineas de fact
                    data_cli.RecordSource = "Select * from clientes_history where cl_codigo =" & Text2.Text
                    data_cli.Refresh
                    If data_cli.Recordset.RecordCount > 0 Then
                       data_cli.Recordset.MoveFirst
                       Do While Not data_cli.Recordset.EOF
                          data_cli.Recordset.Delete
                          data_cli.Recordset.MoveNext
                       Loop
                       data_cli.Refresh
                    End If
                    
                    data_datos.RecordSource = "Select * from linmmdd where cod_cli =" & Text2.Text
                    data_datos.Refresh
                    If data_datos.Recordset.RecordCount > 0 Then
                       data_datos.Recordset.MoveLast
                       pb1.Max = pb1.Max + data_datos.Recordset.RecordCount
                       data_datos.Recordset.MoveFirst
                       DoEvents
                       Do While Not data_datos.Recordset.EOF
        '                  data_datos.Recordset.EditMode
                          data_datos.Recordset("cod_cli") = Text1.Text
                          data_datos.Recordset.Update
                          data_datos.Recordset.MoveNext
                          pb1.Value = pb1.Value + 1
                       Loop
                    End If
        'cabezal HCE
        
                     ' Cajas
                    data_datos.RecordSource = "Select * from caja where cod_socio =" & Text2.Text
                    data_datos.Refresh
                    If data_datos.Recordset.RecordCount > 0 Then
                       data_datos.Recordset.MoveLast
                       pb1.Max = pb1.Max + data_datos.Recordset.RecordCount
                       data_datos.Recordset.MoveFirst
                       DoEvents
                       Do While Not data_datos.Recordset.EOF
        '                  data_datos.Recordset.Edit
                          data_datos.Recordset("cod_socio") = Text1.Text
                          data_datos.Recordset.Update
                          data_datos.Recordset.MoveNext
                          pb1.Value = pb1.Value + 1
                       Loop
                    End If
                     ' Llamados
                    data_datos.RecordSource = "Select * from llamado where matric =" & Text2.Text
                    data_datos.Refresh
                    If data_datos.Recordset.RecordCount > 0 Then
                       data_datos.Recordset.MoveLast
                       pb1.Max = pb1.Max + data_datos.Recordset.RecordCount
                       data_datos.Recordset.MoveFirst
                       DoEvents
                       Do While Not data_datos.Recordset.EOF
        '                  data_datos.Recordset.Edit
                          data_datos.Recordset("matric") = Text1.Text
                          data_datos.Recordset.Update
                          data_datos.Recordset.MoveNext
                          pb1.Value = pb1.Value + 1
                       Loop
                    End If
                    ' Deudas
                    data_datos.RecordSource = "Select * from deudas where cliente =" & Text2.Text
                    data_datos.Refresh
                    If data_datos.Recordset.RecordCount > 0 Then
                       data_datos.Recordset.MoveLast
                       pb1.Max = pb1.Max + data_datos.Recordset.RecordCount
                       data_datos.Recordset.MoveFirst
                       DoEvents
                       Do While Not data_datos.Recordset.EOF
        '                  data_datos.Recordset.Edit
                          data_datos.Recordset("cliente") = Text1.Text
                          data_datos.Recordset.Update
                          data_datos.Recordset.MoveNext
                          pb1.Value = pb1.Value + 1
                       Loop
                    End If
                     ' Historial
                    data_datos.RecordSource = "Select * from abmsocio where cl_codigo =" & Text2.Text
                    data_datos.Refresh
                    If data_datos.Recordset.RecordCount > 0 Then
                       data_datos.Recordset.MoveLast
                       pb1.Max = pb1.Max + data_datos.Recordset.RecordCount
                       data_datos.Recordset.MoveFirst
                       DoEvents
                       Do While Not data_datos.Recordset.EOF
        '                  data_datos.Recordset.Edit
                          data_datos.Recordset("cl_codigo") = Text1.Text
                          data_datos.Recordset.Update
                          data_datos.Recordset.MoveNext
                          pb1.Value = pb1.Value + 1
                       Loop
                    End If
                    
                    Dim LaXmat, LaXmatnew As Double
                    LaXmat = CDbl(Text2.Text)
                    LaXmatnew = CDbl(Text1.Text)
                    
                    data2.RecordSource = "Select * from cabezal_hcdig where mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
        '                  data2.Recordset.Edit
                          data2.Recordset("mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    
                    data2.RecordSource = "Select * from cli_crmdeudas where base =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
        '                  data2.Recordset.Edit
                          data2.Recordset("base") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    
                    data2.RecordSource = "Select * from hc_antali where hc_mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
        '                  data2.Recordset.Edit
                          data2.Recordset("hc_mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    
                    data2.RecordSource = "Select * from hc_antalim where hc_mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
        '                  data2.Recordset.Edit
                          data2.Recordset("hc_mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    
                    data2.RecordSource = "Select * from hc_antamb where hc_mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
        '                  data2.Recordset.Edit
                          data2.Recordset("hc_mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    
                    data2.RecordSource = "Select * from hc_antfam where hc_mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
        '                  data2.Recordset.Edit
                          data2.Recordset("hc_mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    
                    data2.RecordSource = "Select * from hc_antgin where hc_mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
        '                  data2.Recordset.Edit
                          data2.Recordset("hc_mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    
                    data2.RecordSource = "Select * from hc_antinm where hc_mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
                          data2.Recordset("hc_mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    
                    data2.RecordSource = "Select * from hc_antinmu where hc_mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
                          data2.Recordset("hc_mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    
                    data2.RecordSource = "Select * from hc_antmad where hc_mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
                          data2.Recordset("hc_mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    
                    data2.RecordSource = "Select * from hc_antmad2 where hc_mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
                          data2.Recordset("hc_mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    data2.RecordSource = "Select * from hc_antper where hc_mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
                          data2.Recordset("hc_mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    
                    data2.RecordSource = "Select * from hc_antperi where hc_mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
                          data2.Recordset("hc_mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    
                    data2.RecordSource = "Select * from hc_antquir where hc_mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
                          data2.Recordset("hc_mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    data2.RecordSource = "Select * from hc_antsocioe where hc_mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
                          data2.Recordset("hc_mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    data2.RecordSource = "Select * from hc_ctrltemp where hc_mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
                          data2.Recordset("hc_mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    data2.RecordSource = "Select * from hc_ctroles where hc_mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
                          data2.Recordset("hc_mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    data2.RecordSource = "Select * from hc_ctrpa where hc_mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
                          data2.Recordset("hc_mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    data2.RecordSource = "Select * from hc_dieta where hc_mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
                          data2.Recordset("hc_mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    data2.RecordSource = "Select * from hc_examen where mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
                          data2.Recordset("mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    data2.RecordSource = "Select * from hc_lin where hc_mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
                          data2.Recordset("hc_mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    data2.RecordSource = "Select * from hc_mcyotro where hc_mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
                          data2.Recordset("hc_mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    data2.RecordSource = "Select * from hc_metascr where hc_mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
                          data2.Recordset("hc_mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    data2.RecordSource = "Select * from hc_metasr where hc_mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
                          data2.Recordset("hc_mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    data2.RecordSource = "Select * from hc_oculis where hc_mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
                          data2.Recordset("hc_mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    data2.RecordSource = "Select * from hc_odontol where hc_mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
                          data2.Recordset("hc_mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    data2.RecordSource = "Select * from hc_paracl where hc_mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
                          data2.Recordset("hc_mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    data2.RecordSource = "Select * from hc_prescrip where hc_mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
                          data2.Recordset("hc_mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    data2.RecordSource = "Select * from meta_dos where m_mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
                          data2.Recordset("m_mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
                    data2.RecordSource = "Select * from meta_uno where m_mat =" & LaXmat
                    data2.Refresh
                    If data2.Recordset.RecordCount > 0 Then
                       data2.Recordset.MoveFirst
                       Do While Not data2.Recordset.EOF
                          data2.Recordset("m_mat") = LaXmatnew
                          data2.Recordset.Update
                          data2.Recordset.MoveNext
                       Loop
                    End If
        'clirespl
                    data_datos.RecordSource = "Select * from clirespl where cl_codigo =" & Text2.Text
                    data_datos.Refresh
                    If data_datos.Recordset.RecordCount > 0 Then
        '               data_datos.Recordset.MoveLast
        '               pb1.Max = pb1.Max + data_datos.Recordset.RecordCount
                       data_datos.Recordset.MoveFirst
                       DoEvents
                       Do While Not data_datos.Recordset.EOF
        '                  data_datos.Recordset.Edit
                          data_datos.Recordset("cl_codigo") = Text1.Text
                          data_datos.Recordset.Update
                          data_datos.Recordset.MoveNext
        '                  pb1.Value = pb1.Value + 1
                       Loop
                    End If
                    
                    ' Borrar
                    
                    data_cli.RecordSource = "Select * from clientes where cl_codigo =" & Text2.Text
                    data_cli.Refresh
                    If data_cli.Recordset.RecordCount > 0 Then
                       Xconvcod = data_cli.Recordset("cl_codconv")
                       Data1.Recordset.AddNew
                       Data1.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                       Data1.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                       Data1.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                       Data1.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                       Data1.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                       Data1.Recordset("cl_nombre") = Label7.Caption
                       Data1.Recordset("cl_fultpag") = Date
                       Data1.Recordset.Update
                       
                       data_cli.Recordset.Delete
                    End If
                    data_abm.RecordSource = "Select * from abmsocio where cl_codigo =" & Text1.Text
                    data_abm.Refresh
                    data_abm.Recordset.AddNew
                    data_abm.Recordset("cl_codigo") = Text1.Text
                    data_abm.Recordset("cl_motivo") = "UNIFICA MAT."
                    data_abm.Recordset("desc") = "MODIF"
                    data_abm.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
                    data_abm.Recordset("hora") = Format(Time, "HH:mm")
                    data_abm.Recordset("usuario") = Label7.Caption
                    data_abm.Recordset("convenio") = Xconvcod
                    data_abm.Recordset("base") = data_parsec.Recordset("base")
                    data_abm.Recordset.Update
                    
                    pb1.Value = pb1.Value + 100
                    DoEvents
                    frm_depuraficha.MousePointer = 0
                    MsgBox "Proceso terminado!", vbInformation, "SAPP"
                    Text2.Text = ""
                    Text1.Text = ""
                    Command2.Enabled = True
                 End If
              Else
                 MsgBox "Ingrese una matrícula válida", vbCritical
                 Text2.SetFocus
              End If
           Else
              MsgBox "Ingrese una matrícula válida", vbCritical
              Text1.SetFocus
           End If
       Else
            MsgBox "ATENCION! La matrícula que ingresó para eliminar NO EXISTE, verifique en padrón!", vbCritical
       End If
    Else
       MsgBox "ATENCION! La matrícula que ingresó NO EXISTE, verifique en padrón!", vbCritical
    
    End If
Else
    MsgBox "Ingrese una matrícula válida", vbCritical
    Text1.SetFocus
End If

Exit Sub

ErrDepurar:
            If Err.Number = 3351 Then
               MsgBox "Hay un error en los datos, verifique!", vbCritical
            Else
               MsgBox "El socio contiene datos que no se pueden cambiar, pruebe invertir las matrículas", vbCritical
            End If

End Sub

Private Sub Form_Activate()
If App.PrevInstance = True Then
   MsgBox "La aplicación ya se encuentra abierta"
   End
Else
   Label7.Caption = frm_seg.Data1.Recordset("usuario")
   
End If

End Sub

Private Sub Form_Load()
data_cli.Connect = "ODBC;DSN=sappnew;"
data_abm.Connect = "ODBC;DSN=sappnew;"

'data_cli.DatabaseName = App.Path & "\sapp.mdb"
'data_datos.Connect = "ODBC;DSN=sapp;"
data_datos.ConnectionString = "dsn=sappnew"
data2.ConnectionString = "dsn=sappnew"
'data_datos.DatabaseName = App.Path & "\sapp.mdb"
Data1.DatabaseName = App.Path & "\borrado.mdb"
data_parsec.DatabaseName = App.Path & "\parse.mdb"
data_parsec.RecordSource = "parsec0"
data_parsec.Refresh

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Text2.SetFocus
End If

End Sub

Private Sub Text1_LostFocus()
If Text1.Text <> "" Then
   If Text1.Text > 0 Then
      data_cli.RecordSource = "Select * from clientes where cl_codigo =" & Text1.Text
      data_cli.Refresh
      If data_cli.Recordset.RecordCount > 0 Then
         Label5.Caption = data_cli.Recordset("cl_apellid")
      Else
         MsgBox "ATENCION:!! la matrícula ingresada no existe " & Text1.Text, vbCritical, "SAPP"
         Text1.Text = ""
         Label5.Caption = ""
      End If
   Else
      Label5.Caption = ""
   End If
Else
   Label5.Caption = ""
End If

End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command2.SetFocus
End If

End Sub

Private Sub Text2_LostFocus()
If Text2.Text <> "" Then
   If Text2.Text > 0 Then
      data_cli.RecordSource = "Select * from clientes where cl_codigo =" & Text2.Text
      data_cli.Refresh
      If data_cli.Recordset.RecordCount > 0 Then
         Label6.Caption = data_cli.Recordset("cl_apellid")
      Else
         MsgBox "ATENCION:!! la matrícula ingresada no existe " & Text2.Text, vbCritical, "SAPP"
'         Text2.Text = ""
'         Label6.Caption = ""
      End If
   Else
      Label6.Caption = ""
   End If
Else
   Label6.Caption = ""
End If

End Sub
