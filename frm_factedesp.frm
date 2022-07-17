VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_factedesp 
   Caption         =   "Form1"
   ClientHeight    =   3405
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8010
   Icon            =   "frm_factedesp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3405
   ScaleWidth      =   8010
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin MSAdodcLib.Adodc data_llamod 
      Height          =   495
      Left            =   120
      Top             =   2160
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
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
      Caption         =   "data_llamod"
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
   Begin MSAdodcLib.Adodc data_lincab 
      Height          =   495
      Left            =   4080
      Top             =   2640
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
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
      Caption         =   "data_lincab"
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
   Begin MSAdodcLib.Adodc data_lla 
      Height          =   735
      Left            =   4800
      Top             =   1800
      Visible         =   0   'False
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1296
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
      Caption         =   "data_lla"
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
   Begin VB.Data data_cablocal 
      Caption         =   "data_cablocal"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_erro 
      Caption         =   "data_erro"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   360
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Timer Timer1 
      Interval        =   35000
      Left            =   3360
      Top             =   1440
   End
   Begin VB.TextBox txt_fecha 
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   720
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label labtimbre 
      Height          =   255
      Left            =   2640
      TabIndex        =   25
      Top             =   3120
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label labzon 
      Height          =   255
      Left            =   360
      TabIndex        =   24
      Top             =   360
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label labdire 
      Height          =   375
      Left            =   1800
      TabIndex        =   23
      Top             =   1080
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Label labmovilpas 
      Height          =   495
      Left            =   7080
      TabIndex        =   22
      Top             =   2760
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label labcodzon 
      Height          =   375
      Left            =   3120
      TabIndex        =   21
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label labtick 
      Height          =   375
      Left            =   5520
      TabIndex        =   20
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label labnommed 
      Height          =   375
      Left            =   4800
      TabIndex        =   19
      Top             =   1680
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label labcodmed 
      Height          =   255
      Left            =   5400
      TabIndex        =   18
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label labcosto 
      Height          =   375
      Left            =   4320
      TabIndex        =   17
      Top             =   2520
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label labsexo 
      Height          =   255
      Left            =   6360
      TabIndex        =   16
      Top             =   2760
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label labnomcat 
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   2880
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Label labtelef 
      Height          =   255
      Left            =   3480
      TabIndex        =   14
      Top             =   2040
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label lablocal 
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   1920
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Label labnom 
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   1560
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label labcateg 
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label codfinal 
      Height          =   375
      Left            =   4320
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label labclave 
      Height          =   255
      Left            =   4440
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label labcodced 
      Height          =   375
      Left            =   2160
      TabIndex        =   8
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label labced 
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label labmatric 
      Height          =   375
      Left            =   2520
      TabIndex        =   6
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label labtrasla 
      Height          =   375
      Left            =   7200
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.Label labusuario 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label labhora 
      Height          =   495
      Left            =   7080
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label labnro 
      Height          =   255
      Left            =   4920
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
End
Attribute VB_Name = "frm_factedesp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Xcant As Integer

Private Sub Command1_Click()
Dim Xlafellama As Date
Dim Xconvcod As String
Timer1.Enabled = False
Xlafellama = CDate("20/08/2019")

On Error GoTo Cierroporeror

data_lla.RecordSource = "Select * from llamado where totend is null and pend =" & 2 & " and fecha >='" & Format(Xlafellama, "yyyy-mm-dd") & "' order by fecha limit 1"
'data_lla.RecordSource = "Select * from llamado where totend not in ('FACT','MAT','CED') and pend =" & 2 & " and nrolla >=" & 70160656 & " and nrolla <=" & 70160673
'70162343
'data_lla.RecordSource = "Select * from llamado where nrolla =" & 50176251
data_lla.Refresh
'70160656
If data_lla.Recordset.RecordCount > 0 Then
   data_lla.Recordset.MoveFirst
   If IsNull(data_lla.Recordset("timbre")) = False Then
      If data_lla.Recordset("timbre") = 1 Then
         labtimbre.Caption = "SI"
      Else
         labtimbre.Caption = "NO"
      End If
   Else
      labtimbre.Caption = "NO"
   End If
      If data_lla.Recordset("movilpas") = 99 Or data_lla.Recordset("categ") = "SAMCB" Or data_lla.Recordset("base") > 0 Or data_lla.Recordset("codzon") = 4 Or data_lla.Recordset("codzon") = 6 Or data_lla.Recordset("codzon") = 7 Or data_lla.Recordset("enfer") = 1 Then
         If data_lla.Recordset("codzon") = 6 Then
            If IsNull(data_lla.Recordset("mes")) = False Then
               If data_lla.Recordset("mes") > 90 Then
                    If IsNull(data_lla.Recordset("cancela")) = True Then
                       labnro.Caption = data_lla.Recordset("nrolla")
                       txt_fecha.Text = data_lla.Recordset("fecha")
                       labhora.Caption = data_lla.Recordset("hora")
                       labusuario.Caption = data_lla.Recordset("usuario")
                       If IsNull(data_lla.Recordset("trasla")) = False Then
                          labtrasla.Caption = data_lla.Recordset("trasla")
                       Else
                          labtrasla.Caption = 0
                       End If
                       If IsNull(data_lla.Recordset("motmov")) = False Then
                          labzon.Caption = Mid(data_lla.Recordset("motmov"), 1, 50)
                       Else
                          labzon.Caption = "S/Z"
                       End If
                       If IsNull(data_lla.Recordset("referen")) = False Then
                          labdire.Caption = Mid(data_lla.Recordset("referen"), 1, 70)
                       Else
                          labdire.Caption = "S/D"
                       End If
                       
                       If IsNull(data_lla.Recordset("ci")) = False Then
                          labced.Caption = data_lla.Recordset("ci")
                       Else
                          labced.Caption = 0
                       End If
                       If IsNull(data_lla.Recordset("codzon")) = False Then
                          labcodzon.Caption = data_lla.Recordset("codzon")
                       Else
                          labcodzon.Caption = 1
                       End If
                       If IsNull(data_lla.Recordset("movilpas")) = False Then
                          labmovilpas.Caption = data_lla.Recordset("movilpas")
                       Else
                          labmovilpas.Caption = 0
                       End If
                       If IsNull(data_lla.Recordset("categ")) = False Then
                          Xconvcod = data_lla.Recordset("categ")
                       Else
                          Xconvcod = "PART"
                       End If
                       
                       If IsNull(data_lla.Recordset("mes")) = False Then
                          If data_lla.Recordset("mes") > 10 Then
                             labcosto.Caption = data_lla.Recordset("mes")
                             If Xconvcod = "SA" Or Xconvcod = "SAF" Or Xconvcod = "CCOMS" Or Xconvcod = "CERCAS" Or Xconvcod = "SEGAM" Or _
                                Xconvcod = "EMERN" Or Xconvcod = "EMERG" Or Xconvcod = "EMERNE" Or Xconvcod = "CAAM" Or Xconvcod = "911" Or _
                                Xconvcod = "SAP" Or Xconvcod = "VIV19" Or Xconvcod = "VIV20" Or Xconvcod = "CAAMEP" Or Xconvcod = "911B" Or _
                                Xconvcod = "MSP" Or Xconvcod = "UDEMM" Or Xconvcod = "CERSEM" Or Xconvcod = "APNORE" Or Xconvcod = "CASH" Or _
                                Xconvcod = "SJ01" Then
                                labcosto.Caption = 0
                                data_lla.Recordset("mes") = 0
                                data_lla.Recordset.Update
                             Else
                                If Xconvcod = "MUCAFL" Or Xconvcod = "MUCAMA" Or Xconvcod = "MUCAMI" Or Xconvcod = "MUCAMM" Or Xconvcod = "MUCAMP" Or _
                                   Xconvcod = "MUCAMS" Or Xconvcod = "MUCAMT" Or Xconvcod = "MUCATA" Or Xconvcod = "SOLEME" Or Xconvcod = "911B" Or _
                                   Xconvcod = "CAAMEP" Or Xconvcod = "SOLAF" Or Xconvcod = "SOLAMB" Or Xconvcod = "SOC" Or Xconvcod = "911" Or _
                                   Xconvcod = "CASH" Or Xconvcod = "911B" Or Xconvcod = "CERSEM" Or Xconvcod = "CPS" Then
                                   labcosto.Caption = 0
                                   data_lla.Recordset("mes") = 0
                                   data_lla.Recordset.Update
                                End If
                             End If
                          Else
                             labcosto.Caption = 0
                          End If
                       Else
                          labcosto.Caption = 0
                       End If
                       data_llamod.RecordSource = "Select * from resplla where nro =" & labnro.Caption
                       data_llamod.Refresh
                       If data_llamod.Recordset.RecordCount > 0 Then
                          If IsNull(data_llamod.Recordset("mes")) = False Then
                             labcodced.Caption = Int(data_llamod.Recordset("mes"))
                          Else
                             labcodced.Caption = 0
                          End If
                       End If
                       If IsNull(data_lla.Recordset("codmot")) = False Then
                          labclave.Caption = data_lla.Recordset("codmot")
                       Else
                          labclave.Caption = "V"
                       End If
                       If IsNull(data_lla.Recordset("colormot")) = False Then
                          codfinal.Caption = data_lla.Recordset("colormot")
                       Else
                          codfinal.Caption = "V"
                       End If
                       If IsNull(data_lla.Recordset("categ")) = False Then
                          labcateg.Caption = data_lla.Recordset("categ")
                       Else
                          labcateg.Caption = "PART"
                       End If
                       If IsNull(data_lla.Recordset("nombre")) = False Then
                          labnom.Caption = data_lla.Recordset("nombre")
                       Else
                          labnom.Caption = "NN"
                       End If
                       If IsNull(data_lla.Recordset("motmov")) = False Then
                          lablocal.Caption = data_lla.Recordset("motmov")
                       Else
                          lablocal.Caption = "Sin Datos"
                       End If
                       If IsNull(data_lla.Recordset("telef")) = False Then
                          labtelef.Caption = data_lla.Recordset("telef")
                       Else
                          labtelef.Caption = "Sin Datos"
                       End If
                       If IsNull(data_lla.Recordset("nomcat")) = False Then
                          labnomcat.Caption = data_lla.Recordset("nomcat")
                       Else
                          labnomcat.Caption = "Sin Datos"
                       End If
                       If IsNull(data_lla.Recordset("hh")) = False Then
                          labsexo.Caption = data_lla.Recordset("hh")
                       Else
                          labsexo.Caption = 1
                       End If
                       If IsNull(data_lla.Recordset("codmed")) = False Then
                          labcodmed.Caption = data_lla.Recordset("codmed")
                       Else
                          labcodmed.Caption = 440
                       End If
                       If IsNull(data_lla.Recordset("matric")) = False Then
                          labmatric.Caption = data_lla.Recordset("matric")
                       Else
                          labmatric.Caption = 0
                       End If
                       If IsNull(data_lla.Recordset("nommed")) = False Then
                          labnommed.Caption = data_lla.Recordset("nommed")
                       Else
                          labnommed.Caption = "OTROS MEDICOS"
                       End If
                       If IsNull(data_lla.Recordset("nommed")) = False Then
                          labnommed.Caption = data_lla.Recordset("nommed")
                       Else
                          labnommed.Caption = "OTROS MEDICOS"
                       End If
                       If Format(data_lla.Recordset("fecha"), "yyyy-mm-dd") >= Format("08-07-2016", "yyyy-mm-dd") Then
                          If IsNull(data_lla.Recordset("realiza")) = False Then
                             labtick.Caption = data_lla.Recordset("realiza")
                          Else
                             labtick.Caption = 0
                          End If
                       Else
                          labtick.Caption = 0
                       End If
                       frm_despachofact2.Show
                       Unload frm_despachofact2
                    Else
                       If data_lla.Recordset("cancela") = 1 Then
                          If IsNull(data_lla.Recordset("totend")) = False Then
                             If data_lla.Recordset("totend") <> "" Then
                                If data_lla.Recordset("totend") = "FACT" Then
                                Else
                                   data_lla.Recordset("totend") = "FACT"
                                   data_lla.Recordset.Update
                                End If
                             Else
                                data_lla.Recordset("totend") = "FACT"
                                data_lla.Recordset.Update
                             End If
                          Else
                             data_lla.Recordset("totend") = "FACT"
                             data_lla.Recordset.Update
                          End If
                       Else
                       
                       End If
                    End If
               
               Else
                  If IsNull(data_lla.Recordset("totend")) = False Then
                     If data_lla.Recordset("totend") <> "" Then
                        If data_lla.Recordset("totend") = "FACT" Then
                        Else
                           data_lla.Recordset("totend") = "FACT"
                           data_lla.Recordset.Update
                        End If
                     Else
                        data_lla.Recordset("totend") = "FACT"
                        data_lla.Recordset.Update
                     End If
                  Else
                     data_lla.Recordset("totend") = "FACT"
                     data_lla.Recordset.Update
                  End If
               End If
            Else
               If IsNull(data_lla.Recordset("totend")) = False Then
                  If data_lla.Recordset("totend") <> "" Then
                     If data_lla.Recordset("totend") = "FACT" Then
                     Else
                        data_lla.Recordset("totend") = "FACT"
                        data_lla.Recordset.Update
                     End If
                  Else
                     data_lla.Recordset("totend") = "FACT"
                     data_lla.Recordset.Update
                  End If
               Else
                  data_lla.Recordset("totend") = "FACT"
                  data_lla.Recordset.Update
               End If
            End If
         Else
             If IsNull(data_lla.Recordset("totend")) = False Then
                If data_lla.Recordset("totend") <> "" Then
                   If data_lla.Recordset("totend") = "FACT" Then
                   Else
                      data_lla.Recordset("totend") = "FACT"
                      data_lla.Recordset.Update
                   End If
                Else
                   data_lla.Recordset("totend") = "FACT"
                   data_lla.Recordset.Update
                End If
             Else
                data_lla.Recordset("totend") = "FACT"
                data_lla.Recordset.Update
             End If
         End If
      Else
         If IsNull(data_lla.Recordset("cancela")) = True Then
            labnro.Caption = data_lla.Recordset("nrolla")
            txt_fecha.Text = data_lla.Recordset("fecha")
            labhora.Caption = data_lla.Recordset("hora")
            labusuario.Caption = data_lla.Recordset("usuario")
            If IsNull(data_lla.Recordset("trasla")) = False Then
               labtrasla.Caption = data_lla.Recordset("trasla")
            Else
               labtrasla.Caption = 0
            End If
            If IsNull(data_lla.Recordset("motmov")) = False Then
               labzon.Caption = Mid(data_lla.Recordset("motmov"), 1, 50)
            Else
               labzon.Caption = "S/Z"
            End If
            If IsNull(data_lla.Recordset("referen")) = False Then
               labdire.Caption = Mid(data_lla.Recordset("referen"), 1, 50)
            Else
               labdire.Caption = "S/D"
            End If
            
            If IsNull(data_lla.Recordset("ci")) = False Then
               labced.Caption = data_lla.Recordset("ci")
            Else
               labced.Caption = 0
            End If
            If IsNull(data_lla.Recordset("codzon")) = False Then
               labcodzon.Caption = data_lla.Recordset("codzon")
            Else
               labcodzon.Caption = 1
            End If
            If IsNull(data_lla.Recordset("movilpas")) = False Then
               labmovilpas.Caption = data_lla.Recordset("movilpas")
            Else
               labmovilpas.Caption = 0
            End If
            If IsNull(data_lla.Recordset("categ")) = False Then
               Xconvcod = data_lla.Recordset("categ")
            Else
               Xconvcod = "PART"
            End If
            
            If IsNull(data_lla.Recordset("mes")) = False Then
               If data_lla.Recordset("mes") > 10 Then
                  labcosto.Caption = data_lla.Recordset("mes")
                  If Xconvcod = "SA" Or Xconvcod = "SAF" Or Xconvcod = "CCOMS" Or Xconvcod = "CERCAS" Or Xconvcod = "SEGAM" Or _
                     Xconvcod = "EMERN" Or Xconvcod = "EMERG" Or Xconvcod = "EMERNE" Or Xconvcod = "CAAM" Or Xconvcod = "911" Or _
                     Xconvcod = "SAP" Or Xconvcod = "VIV19" Or Xconvcod = "VIV20" Or Xconvcod = "CAAMEP" Or Xconvcod = "911B" Or _
                     Xconvcod = "MSP" Or Xconvcod = "UDEMM" Or Xconvcod = "CERSEM" Or Xconvcod = "APNORE" Or Xconvcod = "CASH" Or _
                     Xconvcod = "SJ01" Then
                     labcosto.Caption = 0
                     data_lla.Recordset("mes") = 0
                     data_lla.Recordset.Update
                  Else
                     If Xconvcod = "MUCAFL" Or Xconvcod = "MUCAMA" Or Xconvcod = "MUCAMI" Or Xconvcod = "MUCAMM" Or Xconvcod = "MUCAMP" Or _
                        Xconvcod = "MUCAMS" Or Xconvcod = "MUCAMT" Or Xconvcod = "MUCATA" Or Xconvcod = "SOLEME" Or Xconvcod = "911B" Or _
                        Xconvcod = "CAAMEP" Or Xconvcod = "SOLAF" Or Xconvcod = "SOLAMB" Or Xconvcod = "SOC" Or Xconvcod = "911" Or _
                        Xconvcod = "CASH" Or Xconvcod = "911B" Or Xconvcod = "CERSEM" Or Xconvcod = "CPS" Then
                        labcosto.Caption = 0
                        data_lla.Recordset("mes") = 0
                        data_lla.Recordset.Update
                     End If
                  End If
               Else
                  labcosto.Caption = 0
               End If
            Else
               labcosto.Caption = 0
            End If
            data_llamod.RecordSource = "Select * from resplla where nro =" & labnro.Caption
            data_llamod.Refresh
            If data_llamod.Recordset.RecordCount > 0 Then
               If IsNull(data_llamod.Recordset("mes")) = False Then
                  labcodced.Caption = Int(data_llamod.Recordset("mes"))
               Else
                  labcodced.Caption = 0
               End If
            End If
            If IsNull(data_lla.Recordset("codmot")) = False Then
               labclave.Caption = data_lla.Recordset("codmot")
            Else
               labclave.Caption = "V"
            End If
            If IsNull(data_lla.Recordset("colormot")) = False Then
               codfinal.Caption = data_lla.Recordset("colormot")
            Else
               codfinal.Caption = "V"
            End If
            If IsNull(data_lla.Recordset("categ")) = False Then
               labcateg.Caption = data_lla.Recordset("categ")
            Else
               labcateg.Caption = "PART"
            End If
            If IsNull(data_lla.Recordset("nombre")) = False Then
               labnom.Caption = data_lla.Recordset("nombre")
            Else
               labnom.Caption = "NN"
            End If
            If IsNull(data_lla.Recordset("motmov")) = False Then
               lablocal.Caption = data_lla.Recordset("motmov")
            Else
               lablocal.Caption = "Sin Datos"
            End If
            If IsNull(data_lla.Recordset("telef")) = False Then
               labtelef.Caption = data_lla.Recordset("telef")
            Else
               labtelef.Caption = "Sin Datos"
            End If
            If IsNull(data_lla.Recordset("nomcat")) = False Then
               labnomcat.Caption = data_lla.Recordset("nomcat")
            Else
               labnomcat.Caption = "Sin Datos"
            End If
            If IsNull(data_lla.Recordset("hh")) = False Then
               labsexo.Caption = data_lla.Recordset("hh")
            Else
               labsexo.Caption = 1
            End If
            If IsNull(data_lla.Recordset("codmed")) = False Then
               labcodmed.Caption = data_lla.Recordset("codmed")
            Else
               labcodmed.Caption = 440
            End If
            If IsNull(data_lla.Recordset("matric")) = False Then
               labmatric.Caption = data_lla.Recordset("matric")
            Else
               labmatric.Caption = 0
            End If
            If IsNull(data_lla.Recordset("nommed")) = False Then
               labnommed.Caption = data_lla.Recordset("nommed")
            Else
               labnommed.Caption = "OTROS MEDICOS"
            End If
            If IsNull(data_lla.Recordset("nommed")) = False Then
               labnommed.Caption = data_lla.Recordset("nommed")
            Else
               labnommed.Caption = "OTROS MEDICOS"
            End If
            If Format(data_lla.Recordset("fecha"), "yyyy-mm-dd") >= Format("08-07-2016", "yyyy-mm-dd") Then
               If IsNull(data_lla.Recordset("realiza")) = False Then
                  labtick.Caption = data_lla.Recordset("realiza")
               Else
                  labtick.Caption = 0
               End If
            Else
               labtick.Caption = 0
            End If
            frm_despachofact2.Show
            Unload frm_despachofact2
         Else
            If data_lla.Recordset("cancela") = 1 Then
               If IsNull(data_lla.Recordset("totend")) = False Then
                  If data_lla.Recordset("totend") <> "" Then
                     If data_lla.Recordset("totend") = "FACT" Then
                     Else
                        data_lla.Recordset("totend") = "FACT"
                        data_lla.Recordset.Update
                     End If
                  Else
                     data_lla.Recordset("totend") = "FACT"
                     data_lla.Recordset.Update
                  End If
               Else
                  data_lla.Recordset("totend") = "FACT"
                  data_lla.Recordset.Update
               End If
            Else
            
            End If
         End If
      End If
End If
data_lla.Recordset.Close

Timer1.Enabled = True

Exit Sub

Cierroporeror:
              If Err.Number = 3155 Then
                 data_erro.Recordset.AddNew
                 data_erro.Recordset("id") = 1
                 data_erro.Recordset("fecha") = Date
                 data_erro.Recordset("hora") = Format(Time, "HH:mm")
                 data_erro.Recordset("nroerr") = Err.Number
                 data_erro.Recordset("desc") = "Al command1 de factedesp"
                 data_erro.Recordset.Update
              Else
                 data_erro.Recordset.AddNew
                 data_erro.Recordset("id") = 1
                 data_erro.Recordset("fecha") = Date
                 data_erro.Recordset("hora") = Format(Time, "HH:mm")
                 data_erro.Recordset("nroerr") = Err.Number
                 data_erro.Recordset("desc") = "Al command1 de factedesp"
                 data_erro.Recordset.Update
              End If
              End

End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
   MsgBox "Ya está abierto el programa. Ya está realizando facturación.", vbCritical
   End
End If

On Error GoTo xquepasaalini

'data_lla.Connect = "ODBC;DSN=sappnew;"
data_lla.ConnectionString = "dsn=sappnew"
'data_llamod.Connect = "ODBC;DSN=sappnew;"
data_llamod.ConnectionString = "dsn=sappnew"

data_erro.DatabaseName = App.Path & "\errores.mdb"
data_erro.RecordSource = "errores"
data_erro.Refresh

'data_lincab.Connect = "ODBC;DSN=sappnew;"
data_lincab.ConnectionString = "dsn=sappnew"
data_lincab.RecordSource = "select * from clirespl where cl_fnac ='" & Format(Date, "yyyy-mm-dd") & "'"
data_lincab.Refresh

data_cablocal.DatabaseName = App.Path & "\cablocal.mdb"
'data_cablocal.RecordSource = "cabezados"
'data_cablocal.Refresh

data_cablocal.RecordSource = "Select * from cabezados where cl_codced =" & 2
data_cablocal.Refresh

Exit Sub

xquepasaalini:
             If Err.Number = 3155 Then
                data_erro.Recordset.AddNew
                data_erro.Recordset("id") = 1
                data_erro.Recordset("fecha") = Date
                data_erro.Recordset("hora") = Format(Time, "HH:mm")
                data_erro.Recordset("nroerr") = Err.Number
                data_erro.Recordset("desc") = "Al iniciar Load"
                data_erro.Recordset.Update
             Else
                data_erro.Recordset.AddNew
                data_erro.Recordset("id") = 1
                data_erro.Recordset("fecha") = Date
                data_erro.Recordset("hora") = Format(Time, "HH:mm")
                data_erro.Recordset("nroerr") = Err.Number
                data_erro.Recordset("desc") = "Al iniciar Load"
                data_erro.Recordset.Update
             End If
             End

End Sub

Private Sub Timer1_Timer()

'Xcant = Xcant + 1
If Xcant >= 3 Then
   Xcant = 0
   Command1_Click
Else
   Xcant = Xcant + 1
End If

End Sub
