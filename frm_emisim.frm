VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_emisim 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Simulación de emisión"
   ClientHeight    =   2895
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7755
   Icon            =   "frm_emisim.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_simulan 
      Caption         =   "data_simulan"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data data_clipromo 
      Caption         =   "data_clipromo"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_promos 
      Caption         =   "data_promos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   7320
      Picture         =   "frm_emisim.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2520
      Width           =   495
   End
   Begin Crystal.CrystalReport crinfno 
      Left            =   3120
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_infno 
      Caption         =   "data_infno"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   5520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   360
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Data data_emision 
      Caption         =   "data_emision"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_infemis 
      Caption         =   "data_infemis"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   4680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1560
      Visible         =   0   'False
      Width           =   2295
   End
   Begin MSAdodcLib.Adodc data_cnv 
      Height          =   330
      Left            =   120
      Top             =   840
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
      Caption         =   "data_cnv"
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
   Begin MSAdodcLib.Adodc data_cli 
      Height          =   495
      Left            =   240
      Top             =   1680
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
   Begin VB.Data data_rectiq 
      Caption         =   "data_rectiq"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "EMITIQ"
      Top             =   1680
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox txt_nro5 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "3"
      Top             =   1200
      Visible         =   0   'False
      Width           =   615
   End
   Begin Crystal.CrystalReport crsim 
      Left            =   6960
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_ultrec2 
      Caption         =   "data_ultrec2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2460
   End
   Begin VB.Data data_ultrec 
      Caption         =   "data_ultrec"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2700
   End
   Begin VB.Data data_emitiq 
      Caption         =   "data_emitiq"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.Data data_usua 
      Caption         =   "data_usua"
      Connect         =   "Access"
      DatabaseName    =   "C:\WINDOWS\usapp.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "usuarioact"
      Top             =   2400
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_ultemiemi 
      Caption         =   "data_ultemiemi"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data data_ultemisim 
      Caption         =   "data_ultemisim"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   720
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   4800
      Top             =   2400
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2880
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   855
      Left            =   960
      TabIndex        =   0
      Top             =   960
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   1508
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFC0&
      Caption         =   "SIMULACION DE EMISION"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1080
      TabIndex        =   1
      Top             =   240
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   360
      Picture         =   "frm_emisim.frx":09CC
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   1695
   End
End
Attribute VB_Name = "frm_emisim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()

Label1.Caption = "PROCESO TERMINADO!!"
frm_emisim.MousePointer = 0
MsgBox "Proceso de SIMULACION finalizado", vbInformation, "Simulación"
'crsim.Action = 1

End Sub

Private Sub Command3_Click()
Unload Me

End Sub

Private Sub Form_Load()
data_ultrec2.DatabaseName = App.path & "\controles.mdb"
data_ultrec2.RecordSource = "nrosrec"
data_ultrec2.Refresh
data_ultrec.DatabaseName = App.path & "\controles.mdb"
data_ultrec.RecordSource = "ultnro"
data_ultrec.Refresh

'data_infemis.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"
data_infemis.DatabaseName = App.path & "\informes.mdb"
data_infno.DatabaseName = App.path & "\informes.mdb"

data_cnv.ConnectionString = "dsn=" & Xconexrmt
'data_cnv.RecordSource = "convenio"
'data_cnv.Refresh
data_ultemisim.DatabaseName = App.path & "\controles.mdb"
data_ultemisim.RecordSource = "ultsim"
data_ultemisim.Refresh
data_cli.ConnectionString = "dsn=" & Xconexrmt
data_promos.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_clipromo.Connect = "odbc;dsn=" & Xconexrmt & ";"

'data_cli.RecordSource = "Select clientes"
'data_cli.Refresh
data_ultemiemi.DatabaseName = App.path & "\controles.mdb"
data_ultemiemi.RecordSource = "ultemi"
data_ultemiemi.Refresh


data_emision.DatabaseName = App.path & "\simula.mdb"
data_simulan.DatabaseName = App.path & "\simulan.mdb"

'data_emision.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\simula.mdb"
'data_inf.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & Xlugar
'App.Path & "\informes.mdb"
'data_emision.ConnectionString = "dsn=sappfact"
'data_emision.RecordSource = "emisim"
'data_emision.Refresh

data_emitiq.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_emitiq.RecordSource = "emitiq"
data_emitiq.Refresh
crsim.ReportFileName = App.path & "\infsimulan.rpt"
data_rectiq.DatabaseName = App.path & "\env_tiq.mdb"
data_rectiq.RecordSource = "EMITIQ"
data_rectiq.Refresh


Dim Xdig, Xrut, Xtot, Xtot2, Xfactor, i As Integer
Dim Xerr As Integer
Xerr = 0
data_cnv.RecordSource = "Select * from convenio where cnv_emite ='" & "SI" & "' and cnv_ruc is not null"
data_cnv.Refresh
If data_cnv.Recordset.RecordCount > 0 Then
   data_cnv.Recordset.MoveFirst
   Do While Not data_cnv.Recordset.EOF
      i = 0
      If data_cnv.Recordset("cnv_ruc") <> "" Then
        If Len(Trim(data_cnv.Recordset("cnv_ruc"))) = 12 Then
           If IsNumeric(data_cnv.Recordset("cnv_ruc")) Then
              Xdig = Val(Mid(data_cnv.Recordset("cnv_ruc"), 12, 1))
              Xrut = Val(Mid(data_cnv.Recordset("cnv_ruc"), 1, 12))
              Xtot = 0
              Xfactor = 2
              For i = 1 To 11
                  If i = 1 Then
                     Xtot = Val(Mid(data_cnv.Recordset("cnv_ruc"), i, 1)) * 4
                     Xtot2 = Xtot2 + Xtot
                  End If
                  If i = 2 Then
                     Xtot = Val(Mid(data_cnv.Recordset("cnv_ruc"), i, 1)) * 3
                     Xtot2 = Xtot2 + Xtot
                  End If
                  If i = 3 Then
                     Xtot = Val(Mid(data_cnv.Recordset("cnv_ruc"), i, 1)) * 2
                     Xtot2 = Xtot2 + Xtot
                  End If
                  If i = 4 Then
                     Xtot = Val(Mid(data_cnv.Recordset("cnv_ruc"), i, 1)) * 9
                     Xtot2 = Xtot2 + Xtot
                  End If
                  If i = 5 Then
                     Xtot = Val(Mid(data_cnv.Recordset("cnv_ruc"), i, 1)) * 8
                     Xtot2 = Xtot2 + Xtot
                  End If
                  If i = 6 Then
                     Xtot = Val(Mid(data_cnv.Recordset("cnv_ruc"), i, 1)) * 7
                     Xtot2 = Xtot2 + Xtot
                  End If
                  If i = 7 Then
                     Xtot = Val(Mid(data_cnv.Recordset("cnv_ruc"), i, 1)) * 6
                     Xtot2 = Xtot2 + Xtot
                  End If
                  If i = 8 Then
                     Xtot = Val(Mid(data_cnv.Recordset("cnv_ruc"), i, 1)) * 5
                     Xtot2 = Xtot2 + Xtot
                  End If
                  If i = 9 Then
                     Xtot = Val(Mid(data_cnv.Recordset("cnv_ruc"), i, 1)) * 4
                     Xtot2 = Xtot2 + Xtot
                  End If
                  If i = 10 Then
                     Xtot = Val(Mid(data_cnv.Recordset("cnv_ruc"), i, 1)) * 3
                     Xtot2 = Xtot2 + Xtot
                  End If
                  If i = 11 Then
                     Xtot = Val(Mid(data_cnv.Recordset("cnv_ruc"), i, 1)) * 2
                     Xtot2 = Xtot2 + Xtot
                  End If
              Next
              Xtot = Xtot2 Mod 11
              If Xtot > 0 Then
                 Xtot = 11 - Xtot
              Else
                 Xdig = 0
              End If
              If Xtot = 11 Then
                 Xdig = 0
              Else
                 Xdig = Xtot
              End If
              If Xdig = Val(Mid(data_cnv.Recordset("cnv_ruc"), 12, 1)) Then
'                 Timer1.Enabled = True
              Else
                 MsgBox "El convenio " & data_cnv.Recordset("cnv_codigo") & " tiene un error en el RUT, debe modificar para poder generar", vbCritical
                 If WElusuario = "JFERNAN" Then
                    Xerr = 9
                 Else
                    Xerr = 9
                 End If
              End If
           Else
              MsgBox "El convenio " & data_cnv.Recordset("cnv_codigo") & " tiene un error en el RUT, debe modificar para poder generar", vbCritical
              If WElusuario = "JFERNAN" Then
                 Xerr = 9
              Else
                 Xerr = 9
              End If
           End If
        Else
           MsgBox "El convenio " & data_cnv.Recordset("cnv_codigo") & " tiene un error en el RUT, debe modificar para poder generar", vbCritical
           If WElusuario = "JFERNAN" Then
              Xerr = 9
           Else
              Xerr = 9
           End If
        End If
      End If
      Xtot2 = 0
      data_cnv.Recordset.MoveNext
   Loop
End If
If Xerr = 0 Then
   MsgBox "Se controlará socios sin cobrador...Aguarde!"
   Timer1.Enabled = True
Else
   MsgBox "Hay errores a modificar"
End If

''MsgBox "Terminado el inicio"

'data_cnv.RecordSource = "convenio"
'data_cnv.Refresh


End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub Timer1_Timer()
 Dim MiBase As Database
 Dim UnaSesion As Workspace
 Set UnaSesion = Workspaces(0)
 Dim Xfec As Date
 Dim Recemi As Recordset
 Dim Nomemi As String
 Dim Xmes As Integer
 Dim Xano As Integer
 Dim Xiva As Double
 Dim TOT As Double
 Dim CedPromo As String
 Dim Xsindescuento As Integer
 Xsindescuento = 0
 
 Command3.Enabled = False
' Xfec = Date + 25
 Xfec = Date + 15
 
 frm_emisim.MousePointer = 11
 Nomemi = "sim"
 Xmes = Month(Xfec)
 Xano = Year(Xfec)
 If Xmes < 10 Then
    Nomemi = Nomemi + "0" + Trim(str(Xmes)) + Mid(Trim(str(Xano)), 3, 2)
 Else
    Nomemi = Nomemi + Trim(str(Xmes)) + Mid(Trim(str(Xano)), 3, 2)
 End If
'Label1.Caption = "PROCESANDO " + Nomemi

Timer1.Enabled = False

MsgBox "Simulación de emisión: " + Nomemi, vbInformation, "Simulación"

Dim Xquepasa, Idpromos As Integer
Dim Descustr As String
Dim Totdescuento As Double
Dim Generaranual As Integer
Generaranual = 0

Totdescuento = 0
Idpromos = 0
Descustr = ""

Xquepasa = 0

data_cli.RecordSource = "Select * from clientes where cl_fecing >='" & Format("01/01/2016", "yyyy/mm/dd") & "' and cl_nrocobr =" & 0 & " and estado =" & 1 & _
" and cl_codconv not in ('SMIN','UCM','CCNOS','CUDNO','HEVANO','UNIVS','GANOS','CASANO','SEMM1','CASH','CCASMU','UDEMM','PART','APNORE','SEMM','LAGO','SUAT','ASIS','SMI4','CCSD','SECPOL','EVADPA','ESTE','CUDPD','911','911B','MSP','CONVE','TING','SMINR','UNIVNR','CCNRE','MP','CAUTE','CAAMEP','MUCAMT','CAMEPA','ASSIS','CERSEM','CASANR','HEVANR'" & _
",'COVAS','MUCATA','CASA4','CASA2','UNIDI','HEVAN','MUCAMS','SUMUNN','MUCAMP','BLUE','MEDNO','MUCAMM','711','HMIL','FUNDA','ASOCES','PLAYA','CNEF','REDSIE','IMPNO','RUSSO','4','PREFE1','CERANT','CASMU','RIMOS','CPS')"
'data_cli.RecordSource = "Select * from clientes where cl_nrocobr =" & 615 & " and cl_apellid <> '" & "" & "' and estado in (1,0)"
data_cli.Refresh

data_infno.RecordSource = "infcli"
data_infno.Refresh
If data_infno.Recordset.RecordCount > 0 Then
   data_infno.Recordset.MoveFirst
   Do While Not data_infno.Recordset.EOF
      data_infno.Recordset.Delete
      data_infno.Recordset.MoveNext
   Loop
End If
If data_cli.Recordset.RecordCount > 0 Then
   data_cli.Recordset.MoveFirst
   Do While Not data_cli.Recordset.EOF
      If data_cli.Recordset("cl_nrocobr") = 14 Or _
         data_cli.Recordset("cl_nrocobr") = 101 Or data_cli.Recordset("cl_nrocobr") = 102 Or _
         data_cli.Recordset("cl_nrocobr") = 110 Or data_cli.Recordset("cl_nrocobr") = 111 Or _
         data_cli.Recordset("cl_nrocobr") = 133 Or data_cli.Recordset("cl_nrocobr") = 144 Or _
         data_cli.Recordset("cl_nrocobr") = 222 Or data_cli.Recordset("cl_nrocobr") = 333 Or _
         data_cli.Recordset("cl_nrocobr") = 511 Or _
         data_cli.Recordset("cl_nrocobr") = 513 Or data_cli.Recordset("cl_nrocobr") = 515 Or _
         data_cli.Recordset("cl_nrocobr") = 516 Or _
         data_cli.Recordset("cl_nrocobr") = 555 Or data_cli.Recordset("cl_nrocobr") = 518 Or data_cli.Recordset("cl_nrocobr") = 15 Then
      Else
         data_cnv.RecordSource = "Select * from convenio where cnv_codigo ='" & Trim(data_cli.Recordset("cl_codconv")) & "'"
         data_cnv.Refresh
         If data_cnv.Recordset.RecordCount > 0 Then
            If data_cnv.Recordset("cnv_emite") = "SI" Then
               If IsNull(data_cnv.Recordset("cnv_fbaja")) = True Then
                  If data_cnv.Recordset("cnv_hasta") >= Date Then
                     If data_cnv.Recordset("cnv_cant_r") = 2 Then
                        data_infno.Recordset.AddNew
                        data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                        data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                        data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                        data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                        data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                        data_infno.Recordset("cl_nomconv") = data_cli.Recordset("cl_nomconv")
                        data_infno.Recordset.Update
                        Xquepasa = 9
                     End If
                  End If
               End If
            End If
         End If
      End If
      data_cli.Recordset.MoveNext
   Loop
End If

If Xquepasa = 9 Then
   data_infno.RecordSource = "Select * from infcli order by cl_fecing"
   data_infno.Refresh
   frm_emisim.MousePointer = 0
   MsgBox "Hay registros sin cobrador con categorías con emisión, VERIFIQUE!", vbCritical
   crinfno.ReportFileName = App.path & "\infno.rpt"
   crinfno.Action = 1
Else

    data_cli.RecordSource = "Select * from clientes where cl_nrocobr <>" & 0 & " and cl_apellid <> '" & "" & "' and estado in (1,0) " & _
    "and cl_codconv not in ('CCNOS','UNIVS','CCSP','CCSD','UNIDI','CASH','SMI4','SMIN','PART','CCASMU','CAUTE')"
    'data_cli.RecordSource = "Select * from clientes where cl_nrocobr =" & 615 & " and cl_apellid <> '" & "" & "' and estado in (1,0)"
    data_cli.Refresh
    data_cli.Recordset.MoveLast
    Xcount = data_cli.Recordset.RecordCount
    data_cli.Recordset.MoveFirst
    ProgressBar1.Max = Xcount
    Xcount = 0
    ProgressBar1.Value = 0
    'MsgBox "Terminada seleccion cli"
    
    Dim MiBaseact As Database
    Dim Unasesact As Workspace
    Set Unasesact = Workspaces(0)
    Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")
    
    MiBaseact.Execute "Delete * from infemis"
    
    data_infemis.RecordSource = "infemis"
    data_infemis.Refresh
    
    'MsgBox "Terminado borrado temporal"
    
    'If data_emision.Recordset.RecordCount > 0 Then
    '   data_emision.Recordset.MoveFirst
    '   Do While Not data_emision.Recordset.EOF
    '      data_emision.Recordset.Delete
    '      data_emision.Recordset.MoveNext
    '   Loop
    'End If
    
    Set MiBaseact = Unasesact.OpenDatabase(App.path & "\simulan.mdb")
    MiBaseact.Execute "Delete * from emisim where mes_emi =" & Xmes & " and anio_emi =" & Xano
    
    Set MiBaseact = Unasesact.OpenDatabase(App.path & "\simula.mdb")
    MiBaseact.Execute "Delete * from emisim"
    
    data_emision.RecordSource = "Select * from emisim"
    data_emision.Refresh
    
    '''MsgBox "Terminado borrado dos"
    
    Do While Not data_cli.Recordset.EOF
       If data_cli.Recordset("cl_nro_sup") = 4 Or data_cli.Recordset("cl_nrocobr") = 14 Or _
          data_cli.Recordset("cl_nrocobr") = 101 Or data_cli.Recordset("cl_nrocobr") = 102 Or _
          data_cli.Recordset("cl_nrocobr") = 110 Or data_cli.Recordset("cl_nrocobr") = 111 Or _
          data_cli.Recordset("cl_nrocobr") = 133 Or data_cli.Recordset("cl_nrocobr") = 144 Or _
          data_cli.Recordset("cl_nrocobr") = 222 Or data_cli.Recordset("cl_nrocobr") = 333 Or _
          data_cli.Recordset("cl_nrocobr") = 511 Or _
          data_cli.Recordset("cl_nrocobr") = 513 Or data_cli.Recordset("cl_nrocobr") = 515 Or _
          data_cli.Recordset("cl_nrocobr") = 516 Or _
          data_cli.Recordset("cl_nrocobr") = 555 Or data_cli.Recordset("cl_nrocobr") = 518 Or data_cli.Recordset("cl_nrocobr") = 15 Then
          data_cli.Recordset.MoveNext
       Else
            If IsNull(data_cli.Recordset("fecha_baja")) = False Then
               data_cli.Recordset.MoveNext
            Else
               If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                  data_cli.Recordset.MoveNext
               Else
                  If data_cli.Recordset("cl_codigo") <> 0 Then
                     If data_cli.Recordset("cl_nrocobr") <> "" Then
                         If data_cli.Recordset("cl_codigo") <> "" Then
    '                        data_cnv.Recordset.FindFirst "cnv_codigo = '" & Trim(data_cli.Recordset("cl_codconv")) & "'"
                            data_cnv.RecordSource = "Select * from convenio where cnv_codigo ='" & Trim(data_cli.Recordset("cl_codconv")) & "'"
                            data_cnv.Refresh
                            If data_cnv.Recordset.RecordCount > 0 Then
                               If data_cnv.Recordset("cnv_emite") = "SI" Then
                                  If IsNull(data_cnv.Recordset("cnv_fbaja")) = True Then
                                     If data_cnv.Recordset("cnv_hasta") >= Date Then
                                        If data_cnv.Recordset("cnv_cant_r") = 2 Then
                                           If IsNull(data_cli.Recordset("cl_nrocobr")) = False Then
                                              If IsNull(data_cli.Recordset("mesproxemi")) = False Then
                                                 If Val(data_cli.Recordset("mesproxemi")) = Val(Xmes) And Val(data_cli.Recordset("anoproxemi")) = Xano Then
                                                 Else
                                                    If Val(data_cli.Recordset("anoproxemi")) > Xano Then
                                                    Else
                                                       If Val(data_cli.Recordset("mesproxemi")) < Val(Xmes) Then
                                                          MsgBox "ATENCION! ANOTE: Socio: " & data_cli.Recordset("cl_codigo") & " DEBE VERIFICAR MES DE PRÓXIMA EMISIÓN!", vbCritical
                                                       End If
                                                    End If
                                                 End If
                                              Else
                                                 MsgBox "ATENCION! ANOTE: Socio: " & data_cli.Recordset("cl_codigo") & " DEBE VERIFICAR MES DE PRÓXIMA EMISIÓN!", vbCritical
                                              End If
                                             data_emision.Recordset.AddNew
                                             data_emision.Recordset("cod_cnv") = data_cli.Recordset("cl_codconv")
                                             data_emision.Recordset("nom_cnv") = data_cli.Recordset("cl_nomconv")
                                             data_emision.Recordset("tipocta") = "CC"
                                             data_emision.Recordset("cliente") = data_cli.Recordset("cl_codigo")
                                             data_emision.Recordset("apellidos") = data_cli.Recordset("cl_apellid")
                                             If data_cli.Recordset("cl_cedula") < 9999999.9 Then
                                                data_emision.Recordset("cedula") = Int(data_cli.Recordset("cl_cedula"))
                                                CedPromo = Trim(str(data_cli.Recordset("cl_cedula"))) & Trim(str(data_cli.Recordset("cl_codced")))
                                             End If
                                             
                                             data_emision.Recordset("cod") = data_cli.Recordset("cl_codced")
                                             data_emision.Recordset("fecha") = Date
                                             data_emision.Recordset("tipodoc") = "FAC"
                                             data_emision.Recordset("documento") = 0
                                             data_emision.Recordset("tipo") = "EMISION"
                                             data_emision.Recordset("importe") = data_cnv.Recordset("cnv_precio")
                                             data_emision.Recordset("debe_haber") = 1
                                             data_emision.Recordset("moneda") = data_cnv.Recordset("cnv_codmon")
                                             data_emision.Recordset("origen") = "Cuota " + Trim(str(Xmes)) + "/" + Trim(str(Xano))
                                             data_emision.Recordset("operador") = data_usua.Recordset("nombre")
                                             data_emision.Recordset("hora") = Format(Time, "HH:mm")
                                             data_emision.Recordset("dir_cli") = data_cli.Recordset("cl_direcci")
                                             data_emision.Recordset("loc_cli") = data_cli.Recordset("cl_zona")
                                             data_emision.Recordset("tel_cli") = data_cli.Recordset("cl_telefon")
                                             data_emision.Recordset("nro_superv") = 1
                                             data_emision.Recordset("nom_superv") = "SUPERVISOR GENERAL"
                                             data_emision.Recordset("nro_vende") = data_cli.Recordset("cl_nrovend")
                                             data_emision.Recordset("nom_vende") = data_cli.Recordset("cl_nomvend")
                                             data_emision.Recordset("grupo") = data_cli.Recordset("cl_grupo")
                                             data_emision.Recordset("numero") = 0
                                             data_emision.Recordset("zona") = data_cli.Recordset("cl_zona")
                                             data_emision.Recordset("nro_cobr") = data_cli.Recordset("cl_nrocobr")
                                             data_emision.Recordset("nom_cobr") = data_cli.Recordset("cl_nomcobr")
                                             data_emision.Recordset("mes") = Xmes
                                             data_emision.Recordset("ano") = Xano
                                             data_emision.Recordset("color_rec") = data_cnv.Recordset("cnv_colrec")
                                             If data_cli.Recordset("cl_fecing") <> "" Then
                                                data_emision.Recordset("fecha_ing") = Format(data_cli.Recordset("cl_fecing"), "dd/mm/yyyy")
                                             End If
                                             If data_cli.Recordset("cl_fnac") <> "" Then
                                                data_emision.Recordset("fecha_nac") = Format(data_cli.Recordset("cl_fnac"), "dd/mm/yyyy")
                                             End If
                                             data_emision.Recordset("tiquet") = 0
                                             data_emision.Recordset("iva") = 0
                                             data_emision.Recordset("deudas") = 0
                                             data_emision.Recordset("servi") = 0
                                             data_emision.Recordset("ap") = 0
                                             If IsNull(data_cli.Recordset("idpromos")) = False Then
                                                Idpromos = data_cli.Recordset("idpromos")
                                                If Idpromos > 0 Then
                                                   data_promos.RecordSource = "select * from promocion_gpo where id =" & data_cli.Recordset("idpromos")
                                                   data_promos.Refresh
                                                   If data_promos.Recordset.RecordCount > 0 Then
                                                      data_promos.Recordset.MoveFirst
                                                      If data_promos.Recordset("descu_imp") > 0 Then
                                                         Totdescuento = data_promos.Recordset("descu_imp")
                                                      Else
                                                         Descustr = "0." & data_promos.Recordset("descu_por")
                                                         Totdescuento = data_cnv.Recordset("cnv_precio") * CDbl(Descustr)
                                                      End If
                                                   End If
                                                End If
                                             Else
                                                Idpromos = 0
                                                Totdescuento = 0
                                             End If
                                             If Totdescuento > 0 Then
                                                If data_promos.Recordset("descrip") = "Pago anual" Then
                                                   If IsNull(data_cli.Recordset("mesproxemi")) = False Then
                                                      If IsNull(data_cli.Recordset("anoproxemi")) = False Then
                                                         If Val(data_cli.Recordset("mesproxemi")) = Val(Xmes) And Val(data_cli.Recordset("anoproxemi")) = Xano Then
                                                            Generaranual = 1
'                                                            MsgBox "Deberá realizar once nuevas entregas al socio: " & data_cli.Recordset("cl_codigo") & " por promo PAGO ANUAL", vbInformation
'                                                            MsgBox "RECUERDE! Generarle los timbres por socio", vbInformation
                                                         End If
                                                      End If
                                                   End If
                                                End If
                                                If data_promos.Recordset("descrip") = "Grupo de 3 o más" Then
                                                   If IsNull(data_cli.Recordset("cl_codruta")) = False Then
                                                      data_clipromo.RecordSource = "select * from clientes where cl_codruta =" & data_cli.Recordset("cl_codruta") & " and estado in (1)"
                                                      data_clipromo.Refresh
                                                      If data_clipromo.Recordset.RecordCount > 0 Then
                                                         data_clipromo.Recordset.MoveLast
                                                         If data_clipromo.Recordset.RecordCount < 2 Then
                                                            Xsindescuento = 1
                                                         Else
                                                            Xsindescuento = 0
                                                         End If
                                                      Else
                                                         Xsindescuento = 1
                                                      End If
                                                   Else
                                                      data_clipromo.RecordSource = "select * from clientes where cl_codruta =" & Val(CedPromo) & " and estado in (1)"
                                                      data_clipromo.Refresh
                                                      If data_clipromo.Recordset.RecordCount > 0 Then
                                                         data_clipromo.Recordset.MoveLast
                                                         If data_clipromo.Recordset.RecordCount < 2 Then
                                                            Xsindescuento = 1
                                                         Else
                                                            Xsindescuento = 0
                                                         End If
                                                      Else
                                                         Xsindescuento = 1
                                                      End If
                                                   End If
                                                End If
                                                If Xsindescuento = 1 Then
                                                   Totdescuento = 0
                                                   data_emision.Recordset("descimp") = 0
                                                   data_emision.Recordset("descpor") = 0
                                                Else
                                                   data_emision.Recordset("descimp") = -Totdescuento
                                                   data_emision.Recordset("descpor") = data_promos.Recordset("descu_por")
                                                   data_emision.Recordset("promo") = data_promos.Recordset("descrip")
                                                End If
                                                If data_cnv.Recordset("cnv_precio") > 0 Then
                                                   data_emision.Recordset("total") = data_cnv.Recordset("cnv_precio") - Totdescuento
                                                End If
                                             Else
                                                data_emision.Recordset("descimp") = -Totdescuento
                                                data_emision.Recordset("descpor") = 0
                                                data_emision.Recordset("total") = data_cnv.Recordset("cnv_precio")
                                             End If
                                             If IsNull(data_cli.Recordset("mesproxemi")) = False Then
                                                If IsNull(data_cli.Recordset("anoproxemi")) = False Then
                                                   If Val(data_cli.Recordset("mesproxemi")) = Val(Xmes) And Val(data_cli.Recordset("anoproxemi")) = Xano Then
                                                      data_emision.Recordset.Update
                                                   Else
                                                      data_emision.Recordset.CancelUpdate
                                                   End If
                                                Else
                                                   data_emision.Recordset.Update
                                                End If
                                             Else
                                                data_emision.Recordset.Update
                                             End If
                                             Descustr = ""
                                             Totdescuento = 0
                                             Xsindescuento = 0
                                             Idpromos = 0
                                             If Generaranual = 1 Then
                                                Generar_anual (data_cli.Recordset("cl_codigo"))
                                                Generaranual = 0
                                             End If
                                           End If
                                        Else
                                           If data_cli.Recordset("cl_codigo") = data_cnv.Recordset("cnv_cuenta") Then
                                              If IsNull(data_cli.Recordset("cl_nrocobr")) = False Then
                                                 If IsNull(data_cli.Recordset("mesproxemi")) = False Then
                                                    If Val(data_cli.Recordset("mesproxemi")) = Val(Xmes) And Val(data_cli.Recordset("anoproxemi")) = Xano Then
                                                    Else
                                                       If Val(data_cli.Recordset("anoproxemi")) > Xano Then
                                                       Else
                                                          If Val(data_cli.Recordset("mesproxemi")) < Val(Xmes) Then
                                                             MsgBox "ATENCION! ANOTE: Socio: " & data_cli.Recordset("cl_codigo") & " DEBE VERIFICAR MES DE PRÓXIMA EMISIÓN!", vbCritical
                                                          End If
                                                       End If
                                                    End If
                                                 Else
                                                    MsgBox "ATENCION! ANOTE: Socio: " & data_cli.Recordset("cl_codigo") & " DEBE VERIFICAR MES DE PRÓXIMA EMISIÓN!", vbCritical
                                                 End If
                                                 data_emision.Recordset.AddNew
                                                 data_emision.Recordset("cod_cnv") = data_cli.Recordset("cl_codconv")
                                                 data_emision.Recordset("nom_cnv") = data_cli.Recordset("cl_nomconv")
                                                 data_emision.Recordset("tipocta") = "CC"
                                                 data_emision.Recordset("cliente") = data_cli.Recordset("cl_codigo")
                                                 data_emision.Recordset("apellidos") = data_cli.Recordset("cl_apellid")
                                                 If data_cli.Recordset("cl_cedula") < 9999999.9 Then
                                                    data_emision.Recordset("cedula") = Int(data_cli.Recordset("cl_cedula"))
                                                    CedPromo = Trim(str(data_cli.Recordset("cl_cedula"))) & Trim(str(data_cli.Recordset("cl_codced")))
                                                 End If
                                                 data_emision.Recordset("cod") = data_cli.Recordset("cl_codced")
                                                 data_emision.Recordset("fecha") = Date
                                                 data_emision.Recordset("tipodoc") = "FAC"
                                                 data_emision.Recordset("documento") = 0
                                                 data_emision.Recordset("tipo") = "EMISION"
                                                 data_emision.Recordset("importe") = data_cnv.Recordset("cnv_precio")
                                                 data_emision.Recordset("debe_haber") = 1
                                                 data_emision.Recordset("moneda") = data_cnv.Recordset("cnv_codmon")
                                                 data_emision.Recordset("origen") = "Cuota " + Trim(str(Xmes)) + "/" + Trim(str(Xano))
                                                 data_emision.Recordset("operador") = data_usua.Recordset("nombre")
                                                 data_emision.Recordset("hora") = Format(Time, "HH:mm")
                                                 data_emision.Recordset("dir_cli") = data_cli.Recordset("cl_direcci")
                                                 data_emision.Recordset("loc_cli") = data_cli.Recordset("cl_zona")
                                                 data_emision.Recordset("tel_cli") = data_cli.Recordset("cl_telefon")
                                                 data_emision.Recordset("nro_superv") = 1
                                                 data_emision.Recordset("nom_superv") = "SUPERVISOR GENERAL"
                                                 data_emision.Recordset("nro_vende") = data_cli.Recordset("cl_nrovend")
                                                 data_emision.Recordset("nom_vende") = data_cli.Recordset("cl_nomvend")
                                                 data_emision.Recordset("grupo") = data_cli.Recordset("cl_grupo")
                                                 data_emision.Recordset("numero") = 0
                                                 data_emision.Recordset("zona") = data_cli.Recordset("cl_zona")
                                                 data_emision.Recordset("nro_cobr") = data_cli.Recordset("cl_nrocobr")
                                                 data_emision.Recordset("nom_cobr") = data_cli.Recordset("cl_nomcobr")
                                                 data_emision.Recordset("mes") = Xmes
                                                 data_emision.Recordset("ano") = Xano
                                                 data_emision.Recordset("color_rec") = data_cnv.Recordset("cnv_colrec")
                                                 If data_cli.Recordset("cl_fecing") <> "" Then
                                                    data_emision.Recordset("fecha_ing") = Format(data_cli.Recordset("cl_fecing"), "dd/mm/yyyy")
                                                 End If
                                                 If data_cli.Recordset("cl_fnac") <> "" Then
                                                    data_emision.Recordset("fecha_nac") = Format(data_cli.Recordset("cl_fnac"), "dd/mm/yyyy")
                                                 End If
                                                 data_emision.Recordset("tiquet") = 0
                                                 data_emision.Recordset("iva") = 0
                                                 data_emision.Recordset("deudas") = 0
                                                 data_emision.Recordset("servi") = 0
                                                 data_emision.Recordset("ap") = 0
                                                If IsNull(data_cli.Recordset("idpromos")) = False Then
                                                    Idpromos = data_cli.Recordset("idpromos")
                                                    If Idpromos > 0 Then
                                                      data_promos.RecordSource = "select * from promocion_gpo where id =" & data_cli.Recordset("idpromos")
                                                      data_promos.Refresh
                                                      If data_promos.Recordset.RecordCount > 0 Then
                                                         data_promos.Recordset.MoveFirst
                                                         If data_promos.Recordset("descu_imp") > 0 Then
                                                            Totdescuento = data_promos.Recordset("descu_imp")
                                                         Else
                                                            Descustr = "0." & data_promos.Recordset("descu_por")
                                                            Totdescuento = data_cnv.Recordset("cnv_precio") * CDbl(Descustr)
                                                         End If
                                                      End If
                                                   End If
                                                Else
                                                   Idpromos = 0
                                                   Totdescuento = 0
                                                End If
                                                If Totdescuento > 0 Then
                                                    If data_promos.Recordset("descrip") = "Pago anual" Then
                                                       If IsNull(data_cli.Recordset("mesproxemi")) = False Then
                                                          If IsNull(data_cli.Recordset("anoproxemi")) = False Then
                                                             If Val(data_cli.Recordset("mesproxemi")) = Val(Xmes) And Val(data_cli.Recordset("anoproxemi")) = Xano Then
                                                                Generaranual = 1
'                                                                MsgBox "Deberá realizar once nuevas entregas al socio: " & data_cli.Recordset("cl_codigo") & " por promo PAGO ANUAL", vbInformation
'                                                                MsgBox "RECUERDE! Generarle los timbres por socio", vbInformation
                                                             End If
                                                          End If
                                                       End If
                                                    End If
                                                   
                                                   If data_promos.Recordset("descrip") = "Grupo de 3 o más" Then
                                                      If IsNull(data_cli.Recordset("cl_codruta")) = False Then
                                                         data_clipromo.RecordSource = "select * from clientes where cl_codruta =" & data_cli.Recordset("cl_codruta") & " and estado in (1)"
                                                         data_clipromo.Refresh
                                                         If data_clipromo.Recordset.RecordCount > 0 Then
                                                            data_clipromo.Recordset.MoveLast
                                                            If data_clipromo.Recordset.RecordCount < 2 Then
                                                               Xsindescuento = 1
                                                            Else
                                                               Xsindescuento = 0
                                                            End If
                                                         Else
                                                            Xsindescuento = 1
                                                         End If
                                                      Else
                                                         data_clipromo.RecordSource = "select * from clientes where cl_codruta =" & Val(CedPromo) & " and estado in (1)"
                                                         data_clipromo.Refresh
                                                         If data_clipromo.Recordset.RecordCount > 0 Then
                                                            data_clipromo.Recordset.MoveLast
                                                            If data_clipromo.Recordset.RecordCount < 2 Then
                                                               Xsindescuento = 1
                                                            Else
                                                               Xsindescuento = 0
                                                            End If
                                                         Else
                                                            Xsindescuento = 1
                                                         End If
                                                      End If
                                                   End If
                                                   If Xsindescuento = 1 Then
                                                      Totdescuento = 0
                                                      data_emision.Recordset("descimp") = 0
                                                      data_emision.Recordset("descpor") = 0
                                                   Else
                                                      data_emision.Recordset("descimp") = -Totdescuento
                                                      data_emision.Recordset("descpor") = data_promos.Recordset("descu_por")
                                                      data_emision.Recordset("promo") = data_promos.Recordset("descrip")
                                                   End If
                                                   If data_cnv.Recordset("cnv_precio") > 0 Then
                                                      data_emision.Recordset("total") = data_cnv.Recordset("cnv_precio") - Totdescuento
                                                   End If
                                                Else
                                                   data_emision.Recordset("descimp") = -Totdescuento
                                                   data_emision.Recordset("descpor") = 0
                                                   data_emision.Recordset("total") = data_cnv.Recordset("cnv_precio")
                                                End If
                                                If IsNull(data_cli.Recordset("mesproxemi")) = False Then
                                                   If IsNull(data_cli.Recordset("anoproxemi")) = False Then
                                                      If Val(data_cli.Recordset("mesproxemi")) = Val(Xmes) And Val(data_cli.Recordset("anoproxemi")) = Xano Then
                                                         data_emision.Recordset.Update
                                                      Else
                                                         data_emision.Recordset.CancelUpdate
                                                      End If
                                                   Else
                                                      data_emision.Recordset.Update
                                                   End If
                                                Else
                                                   data_emision.Recordset.Update
                                                End If
                                                Descustr = ""
                                                Totdescuento = 0
                                                Idpromos = 0
                                                Xsindescuento = 0
                                                If Generaranual = 1 Then
                                                   Generar_anual (data_cli.Recordset("cl_codigo"))
                                                   Generaranual = 0
                                                End If
                                              End If
                                           End If
                                        End If
                                     End If
                                  End If
                               End If
                            End If
                         End If
                     End If
                  End If
                  data_cli.Recordset.MoveNext
               End If
            End If
       End If
       Xcount = Xcount + 1
       ProgressBar1.Value = ProgressBar1.Value + 1
       DoEvents
    Loop
    ' Proceso de timbres y deudas para emisión
    
    data_emision.Refresh
    data_emision.Recordset.MoveLast
    Xcount = data_emision.Recordset.RecordCount
    
    'ProgressBar1.Max = Xcount + data_emision.Recordset.RecordCount
    data_emision.Recordset.MoveFirst
    
''Deudas AP
    data_emitiq.RecordSource = "select * from convenio_tiquets where fecha_pago is null"
    data_emitiq.Refresh
    
    If data_emitiq.Recordset.RecordCount > 0 Then
       data_emitiq.Recordset.MoveLast
       Xcount = Xcount + data_emitiq.Recordset.RecordCount
       data_emitiq.Recordset.MoveFirst
    Else
       MsgBox "Atención!!!! no hay deudas de servicios AP.", vbCritical, "Emisión"
    '   ProgressBar1.Max = Xcount
    End If
    Dim Xdeu22, Xtotd22 As Double
    Xdeu22 = 0
    Xtotd22 = 0
    ProgressBar1.Max = ProgressBar1.Max + Xcount
    Do While Not data_emitiq.Recordset.EOF
       data_emision.RecordSource = "Select * from emisim where cod_cnv ='" & data_emitiq.Recordset("nom_grupo") & "'"
       data_emision.Refresh
       If data_emision.Recordset.RecordCount > 0 Then
          Xdeu22 = data_emitiq.Recordset("importe")
          data_emision.Recordset.Edit
          data_emision.Recordset("ap") = data_emision.Recordset("ap") + Xdeu22
          data_emision.Recordset("total") = data_emision.Recordset("total") + Xdeu22
          data_emision.Recordset.Update
       End If
       data_emitiq.Recordset.MoveNext
       Xcount = Xcount + 1
       ProgressBar1.Value = ProgressBar1.Value + 1
    Loop
    
''Deudas
    data_emitiq.RecordSource = "emitiq"
    data_emitiq.Refresh
    
    If data_emitiq.Recordset.RecordCount > 0 Then
       data_emitiq.Recordset.MoveLast
       Xcount = Xcount + data_emitiq.Recordset.RecordCount
       data_emitiq.Recordset.MoveFirst
    Else
       MsgBox "Atención!!!! no hay deudas para emisión, VERIFIQUE", vbCritical, "Emisión"
    '   ProgressBar1.Max = Xcount
    End If
    Dim Xdeu, Xtotd As Double
    Xdeu = 0
    Xtotd = 0
    ProgressBar1.Max = ProgressBar1.Max + Xcount
    ''''MsgBox "Comienza proceso timbres"
    Do While Not data_emitiq.Recordset.EOF
    '   data_emision.Recordset.FindFirst "cliente =" & data_emitiq.Recordset("mat")
       data_emision.RecordSource = "Select * from emisim where cliente =" & data_emitiq.Recordset("mat")
       data_emision.Refresh
       If data_emision.Recordset.RecordCount > 0 Then
          Xdeu = data_emision.Recordset("deudas") + data_emitiq.Recordset("imp")
          Xtotd = data_emision.Recordset("total") + data_emitiq.Recordset("imp")
          If data_emision.Recordset("deudas") <> Xdeu Then
             data_emision.Recordset.Edit
             data_emision.Recordset("deudas") = Xdeu
             data_emision.Recordset.Update
          End If
          If data_emision.Recordset("total") <> Xtotd Then
             data_emision.Recordset.Edit
             data_emision.Recordset("total") = Xtotd
             data_emision.Recordset.Update
          End If
       End If
       data_emitiq.Recordset.MoveNext
       Xcount = Xcount + 1
       ProgressBar1.Value = ProgressBar1.Value + 1
    Loop
    '''MsgBox "Terminado proceso timbres Uno"
    
    If data_rectiq.Recordset.RecordCount > 0 Then
       data_rectiq.Recordset.MoveLast
       ProgressBar1.Max = ProgressBar1.Max + data_rectiq.Recordset.RecordCount
       data_rectiq.Recordset.MoveFirst
    Else
       MsgBox "Atención!!!! no hay timbres para emisión, VERIFIQUE", vbCritical, "Emisión"
    End If
    
    Do While Not data_rectiq.Recordset.EOF
    '   data_emision.Recordset.FindFirst "cliente =" & data_rectiq.Recordset("mat")
       data_emision.RecordSource = "Select * from emisim where cliente =" & data_rectiq.Recordset("mat")
       data_emision.Refresh
       If data_emision.Recordset.RecordCount > 0 Then
          Xdeu = data_emision.Recordset("tiquet") + data_rectiq.Recordset("imp")
          Xtotd = data_emision.Recordset("total") + data_rectiq.Recordset("imp")
          If data_emision.Recordset("tiquet") <> Xdeu Then
             data_emision.Recordset.Edit
             data_emision.Recordset("tiquet") = Xdeu
             data_emision.Recordset.Update
          End If
          If Xtotd <> data_emision.Recordset("total") Then
             data_emision.Recordset.Edit
             data_emision.Recordset("total") = Xtotd
             data_emision.Recordset.Update
          End If
       End If
       data_rectiq.Recordset.MoveNext
       ProgressBar1.Value = ProgressBar1.Value + 1
    Loop
    '''MsgBox "Terminado timbres"
    data_emision.RecordSource = "emisim"
    data_emision.Refresh
    data_emision.Recordset.MoveFirst
    Do While Not data_emision.Recordset.EOF
    '   TOT = data_emision.Recordset("importe") + data_emision.Recordset("deudas")
       If data_emision.Recordset("ap") > 0 Then
          TOT = data_emision.Recordset("importe") + data_emision.Recordset("descimp") + data_emision.Recordset("ap")
       Else
          TOT = data_emision.Recordset("importe") + data_emision.Recordset("descimp")
       End If
       Xiva = TOT / 1.1
       Xiva = Xiva * 0.1
       TOT = data_emision.Recordset("importe") + data_emision.Recordset("ap") + data_emision.Recordset("deudas") + data_emision.Recordset("tiquet") + data_emision.Recordset("servi") + data_emision.Recordset("descimp")
    '   data_emision.Recordset.Edit
       If data_emision.Recordset("total") <> TOT Then
          data_emision.Recordset.Edit
          data_emision.Recordset("total") = TOT
          data_emision.Recordset.Update
       End If
       If data_emision.Recordset("iva") <> Xiva Then
          data_emision.Recordset.Edit
          data_emision.Recordset("iva") = Xiva
          data_emision.Recordset.Update
       End If
       data_emision.Recordset.MoveNext
       Xcount = Xcount + 1
       ProgressBar1.Value = ProgressBar1.Value + 1
    Loop
    data_emision.Recordset.MoveFirst
    '''MsgBox "Terminado emisión"
    Do While Not data_emision.Recordset.EOF
        data_infemis.Recordset.AddNew
        data_infemis.Recordset("cod_cnv") = data_emision.Recordset("cod_cnv")
        data_infemis.Recordset("nom_cnv") = data_emision.Recordset("nom_cnv")
        data_infemis.Recordset("cliente") = data_emision.Recordset("cliente")
        data_infemis.Recordset("apellidos") = data_emision.Recordset("apellidos")
        data_infemis.Recordset("documento") = data_emision.Recordset("documento")
        data_infemis.Recordset("importe") = data_emision.Recordset("importe")
        data_infemis.Recordset("nro_cobr") = data_emision.Recordset("nro_cobr")
        data_infemis.Recordset("nom_cobr") = data_emision.Recordset("nom_cobr")
        data_infemis.Recordset("mes") = data_emision.Recordset("mes")
        data_infemis.Recordset("ano") = data_emision.Recordset("ano")
        data_infemis.Recordset("color_rec") = data_emision.Recordset("color_rec")
        data_infemis.Recordset("tiquet") = data_emision.Recordset("tiquet")
        data_infemis.Recordset("deudas") = data_emision.Recordset("deudas")
        data_infemis.Recordset("iva") = data_emision.Recordset("iva")
        If IsNull(data_emision.Recordset("descimp")) = False Then
           data_infemis.Recordset("servi") = data_emision.Recordset("descimp")
        End If
        data_infemis.Recordset("total") = data_emision.Recordset("total")
        data_infemis.Recordset("ap") = data_emision.Recordset("ap")
        data_infemis.Recordset.Update
        data_emision.Recordset.MoveNext
    Loop
    
    data_simulan.RecordSource = "emisim"
    data_simulan.Refresh
    data_emision.Recordset.MoveFirst
    Do While Not data_emision.Recordset.EOF
       data_simulan.Recordset.AddNew
       data_simulan.Recordset("cod_cnv") = data_emision.Recordset("cod_cnv")
       data_simulan.Recordset("nom_cnv") = data_emision.Recordset("nom_cnv")
       data_simulan.Recordset("tipocta") = "CC"
       data_simulan.Recordset("cliente") = data_emision.Recordset("cliente")
       data_simulan.Recordset("apellidos") = data_emision.Recordset("apellidos")
       data_simulan.Recordset("cedula") = data_emision.Recordset("cedula")
       data_simulan.Recordset("cod") = data_emision.Recordset("cod")
       data_simulan.Recordset("fecha") = data_emision.Recordset("fecha")
       data_simulan.Recordset("tipodoc") = "FAC"
       data_simulan.Recordset("documento") = 0
       data_simulan.Recordset("tipo") = "EMISION"
       data_simulan.Recordset("importe") = data_emision.Recordset("importe")
       data_simulan.Recordset("debe_haber") = 1
       data_simulan.Recordset("moneda") = data_emision.Recordset("moneda")
       data_simulan.Recordset("origen") = data_emision.Recordset("origen")
       data_simulan.Recordset("operador") = data_emision.Recordset("operador")
       data_simulan.Recordset("hora") = data_emision.Recordset("hora")
       data_simulan.Recordset("dir_cli") = data_emision.Recordset("dir_cli")
       data_simulan.Recordset("loc_cli") = data_emision.Recordset("loc_cli")
       data_simulan.Recordset("tel_cli") = data_emision.Recordset("tel_cli")
       data_simulan.Recordset("nro_superv") = 1
       data_simulan.Recordset("nom_superv") = "SUPERVISOR GENERAL"
       data_simulan.Recordset("nro_vende") = data_emision.Recordset("nro_vende")
       data_simulan.Recordset("nom_vende") = data_emision.Recordset("nom_vende")
       data_simulan.Recordset("grupo") = data_emision.Recordset("grupo")
       data_simulan.Recordset("numero") = data_emision.Recordset("numero")
       data_simulan.Recordset("zona") = data_emision.Recordset("zona")
       data_simulan.Recordset("nro_cobr") = data_emision.Recordset("nro_cobr")
       data_simulan.Recordset("nom_cobr") = data_emision.Recordset("nom_cobr")
       data_simulan.Recordset("mes") = data_emision.Recordset("mes")
       data_simulan.Recordset("ano") = data_emision.Recordset("ano")
       data_simulan.Recordset("color_rec") = data_emision.Recordset("color_rec")
       data_simulan.Recordset("fecha_ing") = data_emision.Recordset("fecha_ing")
       data_simulan.Recordset("fecha_nac") = data_emision.Recordset("fecha_nac")
       data_simulan.Recordset("tiquet") = data_emision.Recordset("tiquet")
       data_simulan.Recordset("iva") = data_emision.Recordset("iva")
       data_simulan.Recordset("deudas") = data_emision.Recordset("deudas")
       data_simulan.Recordset("servi") = data_emision.Recordset("servi")
       data_simulan.Recordset("descimp") = data_emision.Recordset("descimp")
       data_simulan.Recordset("descpor") = data_emision.Recordset("descpor")
       data_simulan.Recordset("promo") = data_emision.Recordset("promo")
       data_simulan.Recordset("total") = data_emision.Recordset("total")
       data_simulan.Recordset("ap") = data_emision.Recordset("ap")
       data_simulan.Recordset("mes_emi") = Xmes
       data_simulan.Recordset("anio_emi") = Xano
       data_simulan.Recordset.Update
       data_emision.Recordset.MoveNext
    Loop
    Label1.Caption = "PROCESO TERMINADO!!"
    frm_emisim.MousePointer = 0
    data_infemis.RecordSource = "Select * from infemis"
    data_infemis.Refresh
    
    MsgBox "Proceso de SIMULACION finalizado", vbInformation, "Simulación"
    crsim.ReportTitle = "EMISION " + Trim(str(Xmes)) + "/" + Trim(str(Xano))
    crsim.Action = 1
    Timer1.Enabled = False
    frm_emisim.Hide
End If

End Sub

Public Sub Generar_anual(ByVal Xmatricula As Long)

Dim Xind As Integer
Xind = 0
Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset
Dim Xmesanual, Xanoanual As Integer
Dim Xfecanual As Date
Dim DescustrAnual As String
Dim TotdescuentoAnual As Double
DescustrAnual = ""
TotdescuentoAnual = 0

Xfecanual = Date + 15
'Xfecanual = Date + 25

Xmesanual = Month(Xfecanual)
Xanoanual = Year(Xfecanual)

If Xmesanual = 12 Then
   Xmesanual = 1
   Xanoanual = Xanoanual + 1
Else
   Xmesanual = Xmesanual + 1
   xanoanua = Xanoanual
End If

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from clientes where cl_codigo =" & Xmatricula
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With

If Xrecclii.RecordCount > 0 Then
   For Xind = 1 To 11
       data_emision.Recordset.AddNew
       data_emision.Recordset("cod_cnv") = Xrecclii("cl_codconv")
       data_emision.Recordset("nom_cnv") = Xrecclii("cl_nomconv")
       data_emision.Recordset("tipocta") = "CC"
       data_emision.Recordset("cliente") = Xrecclii("cl_codigo")
       data_emision.Recordset("apellidos") = Xrecclii("cl_apellid")
       If IsNull(Xrecclii("cl_cedula")) = False Then
          data_emision.Recordset("cedula") = Int(Xrecclii("cl_cedula"))
          data_emision.Recordset("cod") = Xrecclii("cl_codced")
       Else
          data_emision.Recordset("cedula") = 0
          data_emision.Recordset("cod") = 0
       End If
       data_emision.Recordset("fecha") = Date
       data_emision.Recordset("tipodoc") = "FAC"
       data_emision.Recordset("documento") = 0
       data_emision.Recordset("tipo") = "EMISION"
       data_emision.Recordset("importe") = data_cnv.Recordset("cnv_precio")
       data_emision.Recordset("debe_haber") = 1
       data_emision.Recordset("moneda") = data_cnv.Recordset("cnv_codmon")
       data_emision.Recordset("origen") = "Cuota " + Trim(str(Xmesanual)) + "/" + Trim(str(Xanoanual))
       data_emision.Recordset("operador") = WElusuario
       data_emision.Recordset("hora") = Format(Time, "HH:mm")
       data_emision.Recordset("dir_cli") = Xrecclii("cl_direcci")
       data_emision.Recordset("loc_cli") = Xrecclii("cl_zona")
       data_emision.Recordset("tel_cli") = Xrecclii("cl_telefon")
       data_emision.Recordset("nro_superv") = 1
       data_emision.Recordset("nom_superv") = "SUPERVISOR GENERAL"
       data_emision.Recordset("nro_vende") = Xrecclii("cl_nrovend")
       data_emision.Recordset("nom_vende") = Xrecclii("cl_nomvend")
       data_emision.Recordset("grupo") = Xrecclii("cl_grupo")
       data_emision.Recordset("numero") = 0
       data_emision.Recordset("zona") = Xrecclii("cl_zona")
       data_emision.Recordset("nro_cobr") = Xrecclii("cl_nrocobr")
       data_emision.Recordset("nom_cobr") = Xrecclii("cl_nomcobr")
       data_emision.Recordset("mes") = Xmesanual
       data_emision.Recordset("ano") = Xanoanual
       data_emision.Recordset("color_rec") = data_cnv.Recordset("cnv_colrec")
       If IsNull(Xrecclii("cl_fecing")) = False Then
          data_emision.Recordset("fecha_ing") = Format(Xrecclii("cl_fecing"), "dd/mm/yyyy")
       End If
       If IsNull(Xrecclii("cl_fnac")) = False Then
          data_emision.Recordset("fecha_nac") = Format(Xrecclii("cl_fnac"), "dd/mm/yyyy")
       End If
       data_emision.Recordset("tiquet") = 0
       data_emision.Recordset("iva") = 0
       data_emision.Recordset("deudas") = 0
       data_emision.Recordset("servi") = 0
       
       If data_promos.Recordset("descu_imp") > 0 Then
          TotdescuentoAnual = data_promos.Recordset("descu_imp")
       Else
          DescustrAnual = "0." & data_promos.Recordset("descu_por")
          TotdescuentoAnual = data_cnv.Recordset("cnv_precio") * CDbl(DescustrAnual)
       End If
       data_emision.Recordset("descimp") = -TotdescuentoAnual
       data_emision.Recordset("descpor") = data_promos.Recordset("descu_por")
       data_emision.Recordset("promo") = data_promos.Recordset("descrip")
       If data_cnv.Recordset("cnv_precio") > 0 Then
          data_emision.Recordset("total") = data_cnv.Recordset("cnv_precio") - TotdescuentoAnual
       End If
       data_emision.Recordset.Update
       DescustrAnual = ""
       TotdescuentoAnual = 0
       If Xmesanual = 12 Then
          Xmesanual = 1
          Xanoanual = Xanoanual + 1
       Else
          Xmesanual = Xmesanual + 1
          xanoanua = Xanoanual
       End If

   Next Xind
End If

Xrecclii.Close
ConbdSapp.Close

End Sub
