VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_afilfactura 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "Facturación"
   ClientHeight    =   3660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2280
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data data_altas 
      Caption         =   "data_altas"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   720
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data data_deudas 
      Caption         =   "data_deudas"
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
      Width           =   2415
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   3240
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      Destination     =   1
      WindowMaxButton =   0   'False
      WindowMinButton =   0   'False
      DiscardSavedData=   -1  'True
      WindowState     =   1
      PrintFileLinesPerPage=   60
      WindowShowNavigationCtls=   0   'False
      WindowShowCancelBtn=   0   'False
      WindowShowPrintBtn=   0   'False
      WindowShowExportBtn=   0   'False
      WindowShowZoomCtl=   0   'False
      WindowShowProgressCtls=   0   'False
   End
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   5520
      ScaleHeight     =   1635
      ScaleWidth      =   1755
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Data data_imagen 
      Caption         =   "data_imagen"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1080
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data data_facafil 
      Caption         =   "data_facafil"
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
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data data_lin3 
      Caption         =   "data_lin3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data data_lin 
      Caption         =   "data_lin"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data data_par 
      Caption         =   "data_par"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   480
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_caja 
      Caption         =   "data_caja"
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
      Top             =   120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_lin2 
      Caption         =   "data_lin2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_cabezal 
      Caption         =   "data_cabezal"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data data_cabeza2 
      Caption         =   "data_cabeza2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_temp 
      Caption         =   "data_temp"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data data_parse 
      Caption         =   "data_parse"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data data_cablocal 
      Caption         =   "data_cablocal"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   -120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   720
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
      Left            =   1560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   360
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_afilcons 
      Caption         =   "data_afilcons"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000005&
      Caption         =   "FACTURAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      Picture         =   "frm_afilfactura.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Label labvenceok 
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   2040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label labhasta 
      Height          =   255
      Left            =   2160
      TabIndex        =   26
      Top             =   1680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label labdesde 
      Height          =   255
      Left            =   120
      TabIndex        =   25
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label labautoriza 
      Height          =   255
      Left            =   2160
      TabIndex        =   24
      Top             =   1320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label labvence 
      Height          =   255
      Left            =   120
      TabIndex        =   23
      Top             =   1320
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label labnrolinea 
      Height          =   255
      Left            =   1800
      TabIndex        =   22
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lablinea 
      Height          =   255
      Left            =   1560
      TabIndex        =   21
      Top             =   3120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label labdescimp 
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   3240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label labdescporce 
      Height          =   255
      Left            =   240
      TabIndex        =   19
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   5400
      TabIndex        =   18
      Top             =   2520
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   2880
      TabIndex        =   17
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label4 
      Height          =   255
      Left            =   3720
      TabIndex        =   16
      Top             =   2160
      Visible         =   0   'False
      Width           =   4215
   End
   Begin VB.Label labnompromocion 
      Height          =   255
      Left            =   3960
      TabIndex        =   15
      Top             =   3360
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label labnrofact 
      Height          =   255
      Left            =   6360
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label labserie 
      Height          =   255
      Left            =   2760
      TabIndex        =   12
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label labporcedes 
      Height          =   375
      Left            =   6960
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label labivadescu 
      Height          =   375
      Left            =   6840
      TabIndex        =   10
      Top             =   1800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label t_pie 
      Height          =   615
      Left            =   3360
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.Label labsubtot 
      Height          =   375
      Left            =   5640
      TabIndex        =   8
      Top             =   3240
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label labiva 
      Height          =   375
      Left            =   5640
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label labdescu 
      Height          =   375
      Left            =   5640
      TabIndex        =   6
      Top             =   2280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label labapagar 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   1680
      Width           =   3255
   End
   Begin VB.Label labtotal 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total a pagar $:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Descuento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   960
      Width           =   5535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Afiliación $:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   360
      Width           =   2175
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   6840
      Picture         =   "frm_afilfactura.frx":058A
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1095
   End
End
Attribute VB_Name = "frm_afilfactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objPosCfe As PosCfe
Dim Miscr As Scripting.FileSystemObject
Dim xmlsc As TextStream
Dim Xlafac, Xlaserieref As String

Dim objUltimaSerieNumero As SerieNumeroCfe

Dim strUltimoGuid As String

Dim strIdTransaccionPos2000 As String

Public Xbb, Xconlin, Xtipodedocumento, Xcandelin As Integer
Public Xivvva, Xtot, Xsubt, Xivauno As Double

Private Sub Command1_Click()
Dim Xivauno As Double

Dim Xlf As Date
Dim Xelano, Xcandelin, XX2, Xcanlineasin As Integer
Dim Xlatasa, Xlatasa22 As Double
Dim XNombre As String
Dim Xmesdesde, Xaniodesde As Integer
Dim Xaa As Integer
Dim Xelivanuevo As Double
Dim Xelivadescu As Double
Dim Xsubtotalnew As Double
Dim Xtotcabezal As Double
Dim Xivacabezal As Double
Dim Xivaparalinea As Double
Xivaparalinea = 0
Xtotcabezal = 0
Xivacabezal = 0
Xcanlineasin = 0
Xsubtotalnew = 0
Xmesdesde = 0
Xaniodesde = 0
XNombre = ""
Xaa = 0
Xelivanuevo = 0
Xelivadescu = 0

Xlatasa = CDbl("10.000")
Xlatasa22 = CDbl("22.000")

Xcandelin = 0
Xelano = Year(Date) + 1
Command1.Enabled = False
If Xdeb = 12 Then
   data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(frm_afilpend.DBGrid1.TextMatrix(frm_afilpend.DBGrid1.RowSel, 1)) & " order by integra_nro"
Else
   data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(frm_afilia.labnro.Caption) & " order by integra_nro"
End If
data_afilcons.Refresh
If data_afilcons.Recordset.RecordCount > 0 Then
   data_afilcons.Recordset.MoveFirst
   Do While Not data_afilcons.Recordset.EOF
      If IsNull(data_afilcons.Recordset("matricula")) = True Then
         MsgBox "ATENCION!!, hay un registro sin datos en la ficha. Comunique a informática.", vbCritical
      End If
       Xcandelin = Xcandelin + 1
       Xcanlineasin = Xcanlineasin + 1
       Label6.Caption = Val(data_afilcons.Recordset("valorcuota"))
       data_temp.Recordset.AddNew
       data_temp.Recordset("obsp") = t_pie.Caption
       data_temp.Recordset("linea") = Xcandelin
       data_temp.Recordset("libro_rub") = "E-TICKET" ' tipo de documento (Ej.e-ticket)
       data_temp.Recordset("in_unid") = "INT1"
       data_temp.Recordset("in_mat") = 2 'gravado a tasa mínima
       data_temp.Recordset("cantidad") = 1
       data_temp.Recordset("tipo") = "CREDITO" 'contado/crédito
       data_temp.Recordset("realizada") = Format(Date, "dd/mm/yyyy")
       data_temp.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
       data_temp.Recordset("cod_cli") = data_afilcons.Recordset("matricula")
       XNombre = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("nom1")
       data_temp.Recordset("nom_cli") = Mid(XNombre, 1, 80)
       data_temp.Recordset("cod_prod") = 992
       data_temp.Recordset("nom_prod") = "AFILIACION"
       data_temp.Recordset("convenio") = data_afilcons.Recordset("categ")
       data_temp.Recordset("operador") = WElusuario
       data_temp.Recordset("hora") = Format(Time, "HH:mm")
       data_temp.Recordset("imp_timbre") = data_afilcons.Recordset("valorcuota") ' Val(labtotal.Caption)
       data_temp.Recordset("tot_lin") = data_afilcons.Recordset("valorcuota") ' Val(labtotal.Caption)
       data_temp.Recordset("base") = data_parse.Recordset("base")
       
       Xivaparalinea = data_afilcons.Recordset("valorcuota") * 0.1 / 1.1
       
       Xelivanuevo = data_afilcons.Recordset("importe_fin") * 0.1 / 1.1
       Xsubtotalnew = data_afilcons.Recordset("importe_fin") - Xelivanuevo
       
       data_temp.Recordset("pre_civa") = Format(Xivaparalinea, "Standard") ' CDbl(labiva.Caption)
       data_temp.Recordset("reg_cab") = 99
       data_temp.Recordset("servicio") = 0
       If Len(data_afilcons.Recordset("cedula")) = 7 Then
          data_temp.Recordset("ced_socio") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 1, 6))
          data_temp.Recordset("fact") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 7, 1))
       Else
          data_temp.Recordset("ced_socio") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 1, 7))
          data_temp.Recordset("fact") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 8, 1))
       End If
       data_temp.Recordset("moneda") = "UYU"
       data_temp.Recordset("nro_flia") = 8
       data_temp.Recordset("nom_flia") = "OTROS SERVICIOS"
       data_temp.Recordset("rub_cont") = data_parse.Recordset("srvcrd") 'rubro
       data_temp.Recordset("arancel") = data_afilcons.Recordset("valorcuota") 'Val(labtotal.Caption)
       data_temp.Recordset("precio_est") = data_afilcons.Recordset("valorcuota") 'Val(labtotal.Caption)
       data_temp.Recordset("imp_iva") = Format(Xivaparalinea, "Standard") 'CDbl(labiva.Caption)
       data_temp.Recordset.Update
       Xtotcabezal = data_afilcons.Recordset("importe_fin")
       Xivacabezal = Xelivanuevo
    '   If Xtotdescu > 0 Then
    '      labivadescu.Caption = Val(Xtotdescu) * 0.1 / 1.1
    '      labivadescu.Caption = Format(labivadescu.Caption, "Standard")
    '   Else
    '      labivadescu.Caption = ""
    '   End If
       
       
       If Label2.Caption = "Sin promoción" Then
          labdescporce.Caption = ""
          labdescimp.Caption = ""
       Else
          Xcandelin = Xcandelin + 1
          data_temp.Recordset.AddNew
          data_temp.Recordset("obsp") = t_pie.Caption
          data_temp.Recordset("linea") = Xcandelin
          data_temp.Recordset("libro_rub") = "E-TICKET" ' tipo de documento (Ej.e-ticket)
          data_temp.Recordset("in_unid") = "INT1"
          data_temp.Recordset("in_mat") = 2 'gravado a tasa mínima
          data_temp.Recordset("cantidad") = 1
          data_temp.Recordset("tipo") = "CREDITO" 'contado/crédito
          data_temp.Recordset("realizada") = Format(Date, "dd/mm/yyyy")
          data_temp.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
          data_temp.Recordset("cod_cli") = data_afilcons.Recordset("matricula")
          XNombre = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("nom1")
          data_temp.Recordset("nom_cli") = Mid(XNombre, 1, 80)
          data_temp.Recordset("cod_prod") = 883
          data_temp.Recordset("nom_prod") = "DESCUENTO " & Trim(str(data_afilcons.Recordset("desc_porce"))) & "%"
          data_temp.Recordset("convenio") = data_afilcons.Recordset("categ")
          data_temp.Recordset("operador") = WElusuario
          data_temp.Recordset("hora") = Format(Time, "HH:mm")
          data_temp.Recordset("imp_timbre") = data_afilcons.Recordset("desc_imp")
          data_temp.Recordset("tot_lin") = data_afilcons.Recordset("desc_imp")
          data_temp.Recordset("base") = data_parse.Recordset("base")
          Xelivadescu = data_afilcons.Recordset("desc_imp") * 0.1 / 1.1
          data_temp.Recordset("pre_civa") = Format(Xelivadescu, "Standard")
          data_temp.Recordset("reg_cab") = 99
          data_temp.Recordset("servicio") = 0
          If Len(data_afilcons.Recordset("cedula")) = 7 Then
             data_temp.Recordset("ced_socio") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 1, 6))
             data_temp.Recordset("fact") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 7, 1))
          Else
             data_temp.Recordset("ced_socio") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 1, 7))
             data_temp.Recordset("fact") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 8, 1))
          End If
          data_temp.Recordset("moneda") = "UYU"
          data_temp.Recordset("nro_flia") = 8
          data_temp.Recordset("nom_flia") = "OTROS SERVICIOS"
          data_temp.Recordset("rub_cont") = data_parse.Recordset("srvcnt") 'rubro
          data_temp.Recordset("arancel") = data_afilcons.Recordset("desc_imp")
          data_temp.Recordset("precio_est") = data_afilcons.Recordset("desc_imp")
          data_temp.Recordset("imp_iva") = Format(Xelivadescu, "Standard")
          data_temp.Recordset.Update
          labdescporce.Caption = data_afilcons.Recordset("desc_porce")
          labdescimp.Caption = data_afilcons.Recordset("desc_imp")
       End If
       If labnompromocion.Caption = "Pago anual" Then
       
          If Day(Date) >= 25 Then
             If Month(Date) = 12 Or Month(Date) = 11 Then
                If Month(Date) = 11 Then
                   Xmesdesde = 1
                   Xaniodesde = Year(Date) + 1
                Else
                   Xmesdesde = 2
                   Xaniodesde = Year(Date) + 1
                End If
             Else
                Xmesdesde = Month(Date) + 2
                Xaniodesde = Year(Date)
             End If
          Else
             If Month(Date) = 12 Then
                Xmesdesde = 1
                Xaniodesde = Year(Date) + 1
             Else
                Xmesdesde = Month(Date) + 1
                Xaniodesde = Year(Date)
             End If
          End If
          
          For Xaa = 1 To 11
              Xcandelin = Xcandelin + 1
              Xcanlineasin = Xcanlineasin + 1
              data_temp.Recordset.AddNew
              data_temp.Recordset("obsp") = t_pie.Caption
              data_temp.Recordset("linea") = Xcandelin
              data_temp.Recordset("libro_rub") = "E-TICKET" ' tipo de documento (Ej.e-ticket)
              data_temp.Recordset("in_unid") = "INT1"
              data_temp.Recordset("in_mat") = 2 'gravado a tasa mínima
              data_temp.Recordset("cantidad") = 1
              data_temp.Recordset("tipo") = "CREDITO" 'contado/crédito
              data_temp.Recordset("realizada") = Format(Date, "dd/mm/yyyy")
              data_temp.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
              data_temp.Recordset("cod_cli") = data_afilcons.Recordset("matricula")
              XNombre = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("nom1")
              data_temp.Recordset("nom_cli") = Mid(XNombre, 1, 80)
              data_temp.Recordset("cod_prod") = 999
              data_temp.Recordset("nom_prod") = "PAGO DE CUOTA " & Trim(str(Xmesdesde)) & "/" & Trim(str(Xaniodesde))
              data_temp.Recordset("mes_paga") = Xmesdesde
              data_temp.Recordset("ano_paga") = Xaniodesde
              data_temp.Recordset("convenio") = data_afilcons.Recordset("categ")
              data_temp.Recordset("operador") = WElusuario
              data_temp.Recordset("hora") = Format(Time, "HH:mm")
              data_temp.Recordset("imp_timbre") = data_afilcons.Recordset("valorcuota")
              data_temp.Recordset("tot_lin") = data_afilcons.Recordset("valorcuota")
              data_temp.Recordset("base") = data_parse.Recordset("base")
              Xelivanuevo = 0
              Xsubtotalnew = 0
              Xivaparalinea = data_afilcons.Recordset("valorcuota") * 0.1 / 1.1
              
              Xelivanuevo = data_afilcons.Recordset("importe_fin") * 0.1 / 1.1
              Xsubtotalnew = data_afilcons.Recordset("importe_fin") - Xelivanuevo
              data_temp.Recordset("pre_civa") = Format(Xivaparalinea, "Standard")
       
              Xtotcabezal = Xtotcabezal + data_afilcons.Recordset("importe_fin")
              Xivacabezal = Xivacabezal + Xelivanuevo
              
              data_temp.Recordset("reg_cab") = 99
              data_temp.Recordset("servicio") = 0
              If Len(data_afilcons.Recordset("cedula")) = 7 Then
                 data_temp.Recordset("ced_socio") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 1, 6))
                 data_temp.Recordset("fact") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 7, 1))
              Else
                 data_temp.Recordset("ced_socio") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 1, 7))
                 data_temp.Recordset("fact") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 8, 1))
              End If
              data_temp.Recordset("moneda") = "UYU"
              data_temp.Recordset("nro_flia") = 8
              data_temp.Recordset("nom_flia") = "OTROS SERVICIOS"
              data_temp.Recordset("rub_cont") = data_parse.Recordset("srvcnt") 'rubro
              data_temp.Recordset("arancel") = data_afilcons.Recordset("valorcuota")
              data_temp.Recordset("precio_est") = data_afilcons.Recordset("valorcuota")
              data_temp.Recordset("imp_iva") = Format(Xivaparalinea, "Standard")
              data_temp.Recordset.Update
                                
              Xcandelin = Xcandelin + 1
              data_temp.Recordset.AddNew
              data_temp.Recordset("obsp") = t_pie.Caption
              data_temp.Recordset("linea") = Xcandelin
              data_temp.Recordset("libro_rub") = "E-TICKET" ' tipo de documento (Ej.e-ticket)
              data_temp.Recordset("in_unid") = "INT1"
              data_temp.Recordset("in_mat") = 2 'gravado a tasa mínima
              data_temp.Recordset("cantidad") = 1
              data_temp.Recordset("tipo") = "CREDITO" 'contado/crédito
              data_temp.Recordset("realizada") = Format(Date, "dd/mm/yyyy")
              data_temp.Recordset("fecha") = Format(Date, "dd/mm/yyyy")
              data_temp.Recordset("cod_cli") = data_afilcons.Recordset("matricula")
              XNombre = data_afilcons.Recordset("ape1") & " " & data_afilcons.Recordset("nom1")
              data_temp.Recordset("nom_cli") = Mid(XNombre, 1, 80)
              data_temp.Recordset("cod_prod") = 883
              data_temp.Recordset("nom_prod") = "DESCUENTO " & Trim(str(data_afilcons.Recordset("desc_porce"))) & "%"
              data_temp.Recordset("convenio") = data_afilcons.Recordset("categ")
              data_temp.Recordset("operador") = WElusuario
              data_temp.Recordset("hora") = Format(Time, "HH:mm")
              data_temp.Recordset("imp_timbre") = data_afilcons.Recordset("desc_imp")
              data_temp.Recordset("tot_lin") = data_afilcons.Recordset("desc_imp")
              data_temp.Recordset("base") = data_parse.Recordset("base")
              Xelivadescu = data_afilcons.Recordset("desc_imp") * 0.1 / 1.1
              data_temp.Recordset("pre_civa") = Format(Xelivadescu, "Standard")
              data_temp.Recordset("reg_cab") = 99
              data_temp.Recordset("servicio") = 0
              If Len(data_afilcons.Recordset("cedula")) = 7 Then
                 data_temp.Recordset("ced_socio") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 1, 6))
                 data_temp.Recordset("fact") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 7, 1))
              Else
                 data_temp.Recordset("ced_socio") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 1, 7))
                 data_temp.Recordset("fact") = Val(Mid(Trim(str(data_afilcons.Recordset("cedula"))), 8, 1))
              End If
              data_temp.Recordset("moneda") = "UYU"
              data_temp.Recordset("nro_flia") = 8
              data_temp.Recordset("nom_flia") = "OTROS SERVICIOS"
              data_temp.Recordset("rub_cont") = data_parse.Recordset("srvcnt") 'rubro
              data_temp.Recordset("arancel") = data_afilcons.Recordset("desc_imp")
              data_temp.Recordset("precio_est") = data_afilcons.Recordset("desc_imp")
              data_temp.Recordset("imp_iva") = Format(Xelivadescu, "Standard")
              data_temp.Recordset.Update
          
              If Xmesdesde = 12 Then
                 Xaniodesde = Xaniodesde + 1
                 Xmesdesde = 1
              Else
                 Xaniodesde = Xaniodesde
                 Xmesdesde = Xmesdesde + 1
              End If
          Next
            
       End If
       
       data_temp.Refresh
       data_temp.Recordset.MoveLast
       data_temp.Recordset.MoveFirst
       data_cabeza2.Recordset.AddNew
       data_cabeza2.Recordset("cl_tipcli") = "1.0"
    '   If Label7.Caption = "E-TICKET" Then
       data_cabeza2.Recordset("cl_tipocli") = 101
       data_cabeza2.Recordset("cl_telefon") = "e-Ticket"
       data_cabeza2.Recordset("cl_fnac") = Format(Date, "dd/mm/yyyy")
       data_cabeza2.Recordset("cl_nrovend") = 1
       data_cabeza2.Recordset("cl_forpago") = 2 'contado
       data_cabeza2.Recordset("cl_celular") = "CREDITO"
       data_cabeza2.Recordset("fecha_modi") = Format(Date, "dd/mm/yyyy")
       data_cabeza2.Recordset("cl_diacobr") = Trim(str(data_par.Recordset("ruc")))
       data_cabeza2.Recordset("cl_nrotarj") = data_par.Recordset("nombre")
       data_cabeza2.Recordset("cl_tjemi_n") = data_par.Recordset("nombre")
       data_cabeza2.Recordset("cl_tjemi_c") = data_par.Recordset("codsuc")
       data_cabeza2.Recordset("cl_referen") = data_par.Recordset("domic")
       data_cabeza2.Recordset("tit_tarj") = data_par.Recordset("ciudad")
       data_cabeza2.Recordset("cl_nomconv") = data_par.Recordset("dpto")
        'receptor
       data_cabeza2.Recordset("cl_nro_sup") = 4
       data_cabeza2.Recordset("hora_baja") = "UY"
       data_cabeza2.Recordset("cl_nom_sup") = data_afilcons.Recordset("matricula")
       data_cabeza2.Recordset("cl_nombre") = "CONSUMO FINAL"
          
          'data_cabezal.Recordset("cl_medflia") = "nro doc para extranjeros"
       data_cabeza2.Recordset("info_debit") = Mid(XNombre, 1, 80)
       data_cabeza2.Recordset("cl_direcci") = Mid(data_afilcons.Recordset("direc1"), 1, 30)
       data_cabeza2.Recordset("cl_zona") = Mid(data_afilcons.Recordset("nomzona"), 1, 25)
       data_cabeza2.Recordset("cl_localid") = "URUGUAY" 'opcional
       data_cabeza2.Recordset("cl_codigo") = data_afilcons.Recordset("matricula")
       data_cabeza2.Recordset("usu_baja") = "UYU"
       
       data_cabeza2.Recordset("cl_atrasop") = Xlatasa
       data_cabeza2.Recordset("cl_decuota") = Xlatasa22
       data_cabeza2.Recordset("saldo_cc2") = 0 'iva básico
       data_cabeza2.Recordset("saldo_doc") = Format(Xtotcabezal, "Standard")
       data_cabeza2.Recordset("cl_grupo") = data_temp.Recordset.RecordCount
       data_cabeza2.Recordset("saldo_chc") = Format(Xtotcabezal, "Standard")
       data_cabeza2.Recordset.Update
       data_cabeza2.Refresh
       data_cabeza2.Recordset.MoveFirst
'       If labnompromocion.Caption = "Pago anual" Then
'          data_cabeza2.Recordset.Edit
'          data_cabeza2.Recordset("saldo_doc2") = data_cabeza2.Recordset("saldo_doc2") * 12
'          data_cabeza2.Recordset("saldo_cc") = data_cabeza2.Recordset("saldo_cc") * 12
'          data_cabeza2.Recordset("saldo_doc") = data_cabeza2.Recordset("saldo_doc") * 12
'          data_cabeza2.Recordset("cl_grupo") = 12
'          data_cabeza2.Recordset("saldo_chc") = data_cabeza2.Recordset("saldo_chc") * 12
'          data_cabeza2.Recordset.Update
'       End If
       
       Xivacabezal = data_cabeza2.Recordset("saldo_doc") * 0.1 / 1.1
       data_cabeza2.Recordset.Edit
       data_cabeza2.Recordset("saldo_cc") = Format(Xivacabezal, "Standard")
       data_cabeza2.Recordset("saldo_doc2") = Format(Xtotcabezal, "Standard") - Format(Xivacabezal, "Standard")
       data_cabeza2.Recordset.Update

       labserie.Caption = ""
       labnrofact.Caption = ""
       lablinea.Caption = Trim(str(Xcanlineasin))
       E_ticket
'       labserie.Caption = "A"
'       labnrofact.Caption = data_parse.Recordset("nrorec") + 1
'       data_parse.Recordset.Edit
'       data_parse.Recordset("nrorec") = data_parse.Recordset("nrorec") + 1
'       data_parse.Recordset.Update
'       data_parse.Refresh
       
       Xcanlineasin = 0
       
       data_cabeza2.Refresh
       If labnrofact.Caption <> "" Then
          If data_cabeza2.Recordset.RecordCount > 0 Then
             data_cabeza2.Recordset.MoveFirst
          End If
           
          data_afilcons.Recordset.MoveNext
          Xcandelin = 0
       Else
           MsgBox "No se pudo generar el documento.", vbCritical
           data_afilcons.Recordset.MoveNext
       End If
   Loop
   
   Unload Me
Else
   MsgBox "No se encuentra afiliación para facturar.", vbInformation
   Unload Me
End If
        

End Sub

Private Sub Form_Load()
Dim Xtotal, Xapagar, Xtotdescu As Double
Xtotal = 0
Xapagar = 0
Xtotdescu = 0

data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"

data_afilcons.Connect = "odbc;dsn=" & Xconexrmt & ";"
If Xdeb = 12 Then
   data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(frm_afilpend.DBGrid1.TextMatrix(frm_afilpend.DBGrid1.RowSel, 1))
Else
   data_afilcons.RecordSource = "select * from afiliaciones_new where afilia_nro =" & Val(frm_afilia.labnro.Caption)
End If
data_afilcons.Refresh
If data_afilcons.Recordset.RecordCount > 0 Then
   data_afilcons.Recordset.MoveFirst
   If IsNull(data_afilcons.Recordset("codpromo")) = False Then
      Consulta_promos
      labporcedes.Caption = data_afilcons.Recordset("desc_porce")
   Else
      Label2.Caption = "Sin promoción"
      labporcedes.Caption = ""
   End If
   Do While Not data_afilcons.Recordset.EOF
      Xtotal = Xtotal + data_afilcons.Recordset("valorcuota")
      If IsNull(data_afilcons.Recordset("desc_imp")) = False Then
         If data_afilcons.Recordset("desc_imp") > 0 Then
            Xapagar = Xapagar + data_afilcons.Recordset("importe_fin")
            Xtotdescu = Xtotdescu + data_afilcons.Recordset("desc_imp")
         Else
            Xapagar = Xtotal
            Xtotdescu = Xtotdescu
         End If
      Else
         Xapagar = Xtotal
         Xtotdescu = Xtotdescu
      End If
      data_afilcons.Recordset.MoveNext
   Loop
   labtotal.Caption = Val(Xtotal)
   labapagar.Caption = Val(Xapagar)
   labdescu.Caption = Val(Xtotdescu)
   labiva.Caption = Val(Xapagar) * 0.1 / 1.1
   labiva.Caption = Format(labiva.Caption, "Standard")
   labsubtot.Caption = Val(Xapagar) - CDbl(labiva.Caption)
   If Xtotdescu > 0 Then
      labivadescu.Caption = Val(Xtotdescu) * 0.1 / 1.1
      labivadescu.Caption = Format(labivadescu.Caption, "Standard")
   Else
      labivadescu.Caption = ""
   End If
   data_afilcons.Recordset.MoveFirst
Else
   labtotal.Caption = ""
   labapagar.Caption = ""
   labdescu.Caption = ""
   labiva.Caption = ""
   labsubtot.Caption = ""
'   Command1.Enabled = False
   
End If

data_erro.DatabaseName = App.path & "\errores.mdb"
data_erro.RecordSource = "errores"
data_erro.Refresh

'data_eror.DatabaseName = App.path & "\erores.mdb"
'data_eror.RecordSource = "erores"
'data_eror.Refresh

data_parse.DatabaseName = App.path & "\PARSE.mdb"
data_parse.RecordSource = "parsec0"
data_parse.Refresh

data_temp.DatabaseName = App.path & "\factura.mdb"

data_cabeza2.DatabaseName = App.path & "\factura.mdb"
data_cabeza2.RecordSource = "cabezados"
data_cabeza2.Refresh
If data_cabeza2.Recordset.RecordCount > 0 Then
   data_cabeza2.Recordset.MoveFirst
   Do While Not data_cabeza2.Recordset.EOF
      data_cabeza2.Recordset.Delete
      data_cabeza2.Recordset.MoveNext
   Loop
End If
data_cabeza2.Refresh

data_deudas.Connect = "ODBC;DSN=" & Xconexrmt & ";"
       
data_cabezal.Connect = "ODBC;DSN=" & Xconexrmt & ";"
data_cabezal.RecordSource = "Select * from clirespl where cl_codigo =" & 25048
data_cabezal.Refresh
    
data_lin2.Connect = "ODBC;DSN=" & Xconexrmt & ";"
data_lin.Connect = "ODBC;DSN=" & Xconexrmt & ";"
data_lin3.Connect = "ODBC;DSN=" & Xconexrmt & ";"
data_facafil.Connect = "ODBC;DSN=" & Xconexrmt & ";"
    
data_caja.Connect = "ODBC;DSN=" & Xconexrmt & ";"
data_caja.RecordSource = "select * from caja where fecha >=#" & Format(Date, "yyyy/mm/dd") & "#"
data_caja.Refresh
           
data_par.Connect = "ODBC;DSN=sappfact;"
data_par.RecordSource = "paramsapp"
data_par.Refresh
    
t_pie.Caption = "Caja de Usuario: " & WElusuario & " FECHA:" & Format(Date, "dd/mm/yyyy")
data_temp.RecordSource = "lineas"
data_temp.Refresh
If data_temp.Recordset.RecordCount > 0 Then
   data_temp.Recordset.MoveFirst
   Do While Not data_temp.Recordset.EOF
      data_temp.Recordset.Delete
      data_temp.Recordset.MoveNext
   Loop
End If
data_temp.Refresh

data_imagen.DatabaseName = App.path & "\imagen.mdb"
data_imagen.RecordSource = "qr"
data_imagen.Refresh


End Sub
Public Sub Consulta_promos()
Dim Xsqlpromo, Xporcent As String
Dim Xrecclii As New ADODB.Recordset
Dim ValorDesc As Double

ConectarBD
ConbdSapp.Open
ValorDesc = 0
Xporcent = ""
Xsqlpromo = "Select * from promocion_gpo where id =" & data_afilcons.Recordset("codpromo")
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Xrecclii.MoveFirst
   labnompromocion.Caption = Xrecclii("descrip")
   Label2.Caption = "Descuento " & Trim(str(Xrecclii("descu_por"))) & "% " & Xrecclii("descrip")
   Label4.Caption = Xrecclii("descrip")
   Label5.Caption = Xrecclii("descu_por")
Else
   labnompromocion.Caption = ""
   Label2.Caption = "Sin Promoción"
   Label4.Caption = ""
   Label5.Caption = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Sub


Public Sub E_ticket()

Dim strIdTransac As String

On Error GoTo Cierrosieser3

Dim Xlaslineas As Integer
Dim XtotSinDesc As Double
XtotSinDesc = 0
Xlaslineas = 1
Set objPosCfe = New PosCfe
    
Dim objresultado As Resultado
    
data_temp.RecordSource = "select * from lineas where cod_prod not in (883)"
data_temp.Refresh
If data_temp.Recordset.RecordCount > 0 Then
   data_temp.Recordset.MoveFirst
End If

If IsNull(data_parse.Recordset("contacto")) = False Then
   Set objresultado = objPosCfe.Inicializar("SAPP0001", data_parse.Recordset("contacto"), vbNullString)
'   Set objresultado = objPosCfe.Inicializar("SAPP-105", data_parse.Recordset("contacto"), vbNullString)
Else
   Set objresultado = objPosCfe.Inicializar("SAPP0001", "SAPP-006", vbNullString)
End If
    
'data_temp.Recordset.MoveFirst
Dim strMensaje As String
strMensaje = "No se pudo inicializar el POS"
If objresultado Is Nothing Then
   MsgBox strMensaje
   Exit Sub
End If

If Not objresultado.OperacionExitosa Then
   If objresultado.Mensaje <> vbNullString Then strMensaje = strMensaje & ": " & objresultado.Mensaje
      MsgBox strMensaje
      Exit Sub
End If
strIdTransac = objPosCfe.CrearGuid
    
    'estado de la conexión
If Not EstaInicializado() Then Exit Sub

   Dim objresultado22 As ResultadoConsultaConexion
   Set objresultado22 = objPosCfe.ObtenerEstadoConexion

   Dim strMensaje22 As String
   strMensaje22 = "No se pudo consultar el estado de la conexión"

   If objresultado22 Is Nothing Then
      MsgBox strMensaje22
      Exit Sub
   End If

   If Not objresultado22.OperacionExitosa Then
      If objresultado22.Mensaje <> vbNullString Then strMensaje22 = strMensaje22 & ": " & objresultado22.Mensaje
         MsgBox strMensaje22
        Exit Sub
   End If

If Not EstaInicializado() Then Exit Sub
    
    Dim objCfe As CFE
    Set objCfe = New CFE

    Dim objCf As ClassFactory

    Set objCf = New ClassFactory
       
    Set objCfe.ETck = New ETck
    With objCfe.ETck.Encabezado.IdDoc
        .TipoCFE = objCf.EnumConverter.IdDoc_Tck_TipoCFEFromString(Trim(str(data_cabeza2.Recordset("cl_tipocli"))))
        .FchEmis.SetDate Year(data_temp.Recordset("fecha")), Month(data_temp.Recordset("fecha")), Day(data_temp.Recordset("fecha"))
        .IsValidMntBruto = True
        .MntBruto = IdDoc_Tck_MntBruto_1
        If data_cabeza2.Recordset("cl_forpago") = 1 Then
           .FmaPago = IdDoc_Tck_FmaPago_1
        Else
           .FmaPago = IdDoc_Tck_FmaPago_2
        End If
    End With
    With objCfe.ETck.Encabezado.Emisor
        .RUCEmisor = data_par.Recordset("ruc")
        .RznSoc = data_par.Recordset("nomc")
        .CdgDGISucur.FromString Trim(str(data_par.Recordset("codsuc")))
        .DomFiscal = data_par.Recordset("domic")
        .Ciudad = data_par.Recordset("ciudad")
        .Departamento = data_par.Recordset("dpto")
    End With
    Set objCfe.ETck.Encabezado.Receptor = New Receptor_Tck
    Set objCfe.ETck.Encabezado.Receptor.Receptor_Tck_Choice = New Receptor_Tck_Choice
    With objCfe.ETck.Encabezado.Receptor
        .TipoDocRecep = DocType_4
        .CodPaisRecep = CodPaisType_UY
        .Receptor_Tck_Choice.DocRecepExt = data_cabeza2.Recordset("cl_nom_sup")
        .RznSocRecep = data_cabeza2.Recordset("info_debit")
        .DirRecep = data_cabeza2.Recordset("cl_direcci")
        .CiudadRecep = data_cabeza2.Recordset("cl_zona")
    End With
    With objCfe.ETck.Encabezado.Totales
        .TpoMoneda = objCf.EnumConverter.TipMonTypeFromString(data_cabeza2.Recordset("usu_baja"))
        .IsValidTpoCambio = True
        .TpoCambio.FromString "1"
        .IsValidMntNetoIvaTasaMin = True
        .IsValidMntIVATasaMin = True
        .MntNetoIvaTasaMin.FromString Format(data_cabeza2.Recordset("saldo_doc2"), "0.00")
        .IVATasaMin = TasaIVAType_10FullStop000
        .MntIVATasaMin.FromString Format(data_cabeza2.Recordset("saldo_cc"), "0.00")
        .CantLinDet.FromString Trim(lablinea.Caption)
        .MntTotal.FromString Format(data_cabeza2.Recordset("saldo_doc"), "0.00")
        .MntPagar.FromString Format(data_cabeza2.Recordset("saldo_chc"), "0.00")
    End With
    Do While Not data_temp.Recordset.EOF
       With objCfe.ETck.Detalle.Item.AddNew
          labnrolinea.Caption = Trim(str(Int(Xlaslineas)))
             If Label2.Caption = "Sin promoción" Then
                .NroLinDet.FromString Trim(labnrolinea.Caption)
                .IndFact = objCf.EnumConverter.Item_Det_Fact_IndFactFromString(Trim(str(data_temp.Recordset("in_mat"))))
                .NomItem = data_temp.Recordset("nom_prod")
                .cantidad.FromString Trim(str(data_temp.Recordset("cantidad")))
                .UniMed = "N/A"
                .PrecioUnitario.FromString Format(data_temp.Recordset("imp_timbre"), "0.00")
                .MontoItem.FromString Format(data_temp.Recordset("tot_lin"), "0.00")
             Else
                XtotSinDesc = data_temp.Recordset("imp_timbre") - Val(labdescimp.Caption)
                XtotSinDesc = Int(XtotSinDesc)
                .NroLinDet.FromString Trim(labnrolinea.Caption)
                .IndFact = objCf.EnumConverter.Item_Det_Fact_IndFactFromString(Trim(str(data_temp.Recordset("in_mat"))))
                .NomItem = data_temp.Recordset("nom_prod")
                .cantidad.FromString Trim(str(data_temp.Recordset("cantidad")))
                .UniMed = "N/A"
                .PrecioUnitario.FromString Format(data_temp.Recordset("imp_timbre"), "0.00") 'Total sin descuento
                .IsValidDescuentoMonto = True
                .IsValidDescuentoPct = True
                .DescuentoPct.FromString Format(labdescporce.Caption, "0")
                .DescuentoMonto.FromString Format(labdescimp.Caption, "0.00") 'monto del descuento
                .MontoItem.FromString Format(XtotSinDesc, "0.00") 'total a pagar sin descuento
             End If
       End With
       Xlaslineas = Xlaslineas + 1
       data_temp.Recordset.MoveNext
       
    Loop
    Dim s As String
    s = objCfe.ToXml(True, XmlFormatting_Indented)

    Dim strGuid As String
    strGuid = objPosCfe.CrearGuid()
    Dim objResultadoCfe As ResultadoCfe
    Set objResultadoCfe = objPosCfe.FirmarYEnviarCfe(strGuid, objCfe)
    
    Set objUltimaSerieNumero = Nothing
    DesplegarInfoEstadoCfe "No se pudo firmar el CFE", objResultadoCfe
    If Not objUltimaSerieNumero Is Nothing Then _
        ' cmdFirmarNc.Enabled = True
'       MsgBox "firmar NC"
    End If
    Fin_eticket
Exit Sub


Cierrosieser3:
                If Err.Number = 3155 Then
                   MsgBox "ERROR: al generar"
                   End
                Else
                   MsgBox "ERROR: al generar e-ticket"
                   End
                End If

End Sub

Public Sub Fin_eticket()
Dim Xlatasa, Xlatasa22 As Double
Dim Xcandelin, Xdondeelerr As Integer
Dim Xelivadeudas As Double

On Error GoTo Algrabaretic

Xelivadeudas = 0

Xlatasa = CDbl("10.000")
Xlatasa22 = CDbl("22.000")
'labserie.Caption = "A"
'labnrofact.Caption = 100011
data_temp.RecordSource = "select * from lineas"
data_temp.Refresh
If data_temp.Recordset.RecordCount > 0 Then
   data_temp.Recordset.MoveFirst
End If
Xcandelin = 0
If data_temp.Recordset.RecordCount > 0 Then
      
   data_lin.RecordSource = "select * from linmmdd where cod_cli =" & data_temp.Recordset("cod_cli")
   data_lin.Refresh
   data_lin2.RecordSource = "Select * from hc_torax"
   data_lin2.Refresh
   data_lin3.RecordSource = "Select * from indica_enfc"
   data_lin3.Refresh
      
   Do While Not data_temp.Recordset.EOF
      Xcandelin = Xcandelin + 1
      data_lin.Recordset.AddNew
      If IsNull(data_temp.Recordset("linea")) = False Then
         data_lin.Recordset("linea") = data_temp.Recordset("linea")
      Else
         data_lin.Recordset("linea") = 1
      End If
      If labnrofact.Caption <> "" Then
         If IsNumeric(labnrofact.Caption) = True Then
            data_lin.Recordset("factura") = Val(labnrofact.Caption)
         Else
            data_lin.Recordset("factura") = Val(labnrofact.Caption)
         End If
      Else
         data_lin.Recordset("factura") = 0
      End If
      data_lin.Recordset("tipo") = data_temp.Recordset("tipo")
      data_lin.Recordset("realizada") = Format(data_temp.Recordset("realizada"), "dd/mm/yyyy")
      data_lin.Recordset("fecha") = Format(data_temp.Recordset("fecha"), "dd/mm/yyyy")
      data_lin.Recordset("cod_cli") = data_temp.Recordset("cod_cli")
      data_lin.Recordset("nom_cli") = Mid(data_temp.Recordset("nom_cli"), 1, 30)
      data_lin.Recordset("convenio") = data_temp.Recordset("convenio")
      data_lin.Recordset("cod_prod") = data_temp.Recordset("cod_prod")
      data_lin.Recordset("nom_prod") = Mid(data_temp.Recordset("nom_prod"), 1, 50)
      data_lin.Recordset("operador") = data_temp.Recordset("operador")
      data_lin.Recordset("hora") = data_temp.Recordset("hora")
      If IsNull(data_temp.Recordset("imp_timbre")) = False Then
         If data_temp.Recordset("cod_prod") = 883 Then
            data_lin.Recordset("imp_timbre") = -data_temp.Recordset("imp_timbre")
         Else
            data_lin.Recordset("imp_timbre") = Format(data_temp.Recordset("imp_timbre"), "Standard") ' sub total de la línea
         End If
      Else
         data_lin.Recordset("imp_timbre") = 0
      End If
      If data_temp.Recordset("cod_prod") = 883 Then
         data_lin.Recordset("tot_lin") = -data_temp.Recordset("tot_lin") ' total de la linea de la factura
         data_lin.Recordset("precio_est") = -data_temp.Recordset("precio_est")
         data_lin.Recordset("imp_iva") = -data_temp.Recordset("imp_iva")
         data_lin.Recordset("pre_civa") = -data_temp.Recordset("pre_civa")
         data_lin.Recordset("valor_iva") = -data_temp.Recordset("pre_civa")
      Else
         data_lin.Recordset("tot_lin") = data_temp.Recordset("tot_lin") ' total de la linea de la factura
         data_lin.Recordset("precio_est") = data_temp.Recordset("precio_est")
         data_lin.Recordset("imp_iva") = data_temp.Recordset("imp_iva")
         data_lin.Recordset("pre_civa") = data_temp.Recordset("pre_civa")
         data_lin.Recordset("valor_iva") = data_temp.Recordset("pre_civa")
      End If
      data_lin.Recordset("base") = data_temp.Recordset("base")
      data_lin.Recordset("nom_med_a") = data_temp.Recordset("nom_med_a")
      data_lin.Recordset("rub_cont") = data_temp.Recordset("rub_cont")
      data_lin.Recordset("nom_flia") = data_temp.Recordset("nom_flia")
      data_lin.Recordset("reg_cab") = data_temp.Recordset("reg_cab") '=99
      data_lin.Recordset("servicio") = data_temp.Recordset("servicio")
      If IsNull(data_temp.Recordset("ced_socio")) = False Then
         data_lin.Recordset("ced_socio") = data_temp.Recordset("ced_socio")
      Else
         data_lin.Recordset("ced_socio") = 0
      End If
      data_lin.Recordset("fact") = data_temp.Recordset("fact") 'codced
      data_lin.Recordset("moneda") = data_temp.Recordset("moneda")
      data_lin.Recordset("nro_flia") = data_temp.Recordset("nro_flia")
      data_lin.Recordset("rub_cont") = data_temp.Recordset("rub_cont")
      data_lin.Recordset("arancel") = data_temp.Recordset("arancel")
      data_lin.Recordset("nro_med_a") = data_temp.Recordset("nro_med_a")
      If labserie.Caption <> "" Then
         data_lin.Recordset("moneda") = labserie.Caption
      Else
         data_lin.Recordset("moneda") = "A"
      End If
      data_lin.Recordset("tipo_mov") = Trim(str("2"))
      data_lin.Recordset("pendiente") = "T"
      data_lin.Recordset("mes_paga") = data_temp.Recordset("mes_paga")
      data_lin.Recordset("ano_paga") = data_temp.Recordset("ano_paga")
      data_lin.Recordset.Update
      
      
      data_temp.Recordset.MoveNext
   Loop
   data_temp.Recordset.MoveFirst
   
   data_facafil.RecordSource = "select * from linmmdd_afil where fecha =#" & Format(Date, "yyyy/mm/dd") & "#"
   data_facafil.Refresh
   data_facafil.Recordset.AddNew
   data_facafil.Recordset("fecha") = data_temp.Recordset("fecha")
   data_facafil.Recordset("factura") = Val(labnrofact.Caption)
   data_facafil.Recordset("codfunc") = data_afilcons.Recordset("codvende")
   data_facafil.Recordset("nombre") = Devuelve_vende()
   data_facafil.Recordset.Update
'------
   data_temp.RecordSource = "Select * from lineas"
   data_temp.Refresh
   If data_temp.Recordset.RecordCount > 0 Then
      data_temp.Recordset.MoveFirst
      Do While Not data_temp.Recordset.EOF
         data_temp.Recordset.Edit
         data_temp.Recordset("cl_nrotarj") = data_cabeza2.Recordset("cl_nrotarj")
         data_temp.Recordset("cl_referen") = data_cabeza2.Recordset("cl_referen")
         data_temp.Recordset("cl_tjemi_c") = data_cabeza2.Recordset("cl_tjemi_c")
         data_temp.Recordset("cl_diacobr") = data_cabeza2.Recordset("cl_diacobr")
         data_temp.Recordset("cl_telefon") = data_cabeza2.Recordset("cl_telefon")
         data_temp.Recordset("qr") = data_cabeza2.Recordset("qr")
         data_temp.Recordset("cl_fax") = data_cabeza2.Recordset("cl_fax")
         data_temp.Recordset("cl_socmnro") = data_cabeza2.Recordset("cl_socmnro")
         data_temp.Recordset("cl_numero") = data_cabeza2.Recordset("cl_numero")
         data_temp.Recordset("cl_celular") = data_cabeza2.Recordset("cl_celular")
         If IsNull(data_cabeza2.Recordset("cl_fnac")) = False Then
            data_temp.Recordset("cl_fnac") = Format(data_cabeza2.Recordset("cl_fnac"), "dd/mm/yyyy")
         End If
         data_temp.Recordset("usu_baja") = data_cabeza2.Recordset("usu_baja")
         data_temp.Recordset("info_debit") = data_cabeza2.Recordset("info_debit")
         data_temp.Recordset("cl_nrocobr") = data_cabeza2.Recordset("cl_nrocobr")
         data_temp.Recordset("cl_medflia") = data_cabeza2.Recordset("cl_medflia")
         data_temp.Recordset("hora_baja") = data_cabeza2.Recordset("hora_baja")
         data_temp.Recordset("cl_nomcobr") = data_cabeza2.Recordset("cl_nomcobr")
         data_temp.Recordset("cl_nom_sup") = data_cabeza2.Recordset("cl_nom_sup")
         data_temp.Recordset("saldo_cc") = data_cabeza2.Recordset("saldo_cc")
         data_temp.Recordset("saldo_doc") = data_cabeza2.Recordset("saldo_doc")
         If IsNull(data_cabeza2.Recordset("cl_fultpag")) = False Then
            data_temp.Recordset("cl_fultpag") = Format(data_cabeza2.Recordset("cl_fultpag"), "dd/mm/yyyy")
         End If
         data_temp.Recordset("cl_nombre") = data_cabeza2.Recordset("cl_nombre")
         data_temp.Recordset.Update
         data_temp.Recordset.MoveNext
      Loop
   End If
   
   data_cabezal.Recordset.AddNew
    '           data_cabezal.Recordset("id") = 1
   data_cabezal.Recordset("cl_tipcli") = "1.0"
   data_cabezal.Recordset("cl_tipocli") = data_cabeza2.Recordset("cl_tipocli")
   data_cabezal.Recordset("cl_socmnro") = labserie.Caption
   data_cabezal.Recordset("cl_numero") = Val(labnrofact.Caption)
   data_cabezal.Recordset("cl_fnac") = data_cabeza2.Recordset("cl_fnac")
   data_cabezal.Recordset("fecha_reac") = data_cabeza2.Recordset("fecha_reac")
   data_cabezal.Recordset("cl_tj_venc") = data_cabeza2.Recordset("cl_tj_venc")
   data_cabezal.Recordset("cl_nrovend") = data_cabeza2.Recordset("cl_nrovend")
   data_cabezal.Recordset("cl_forpago") = data_cabeza2.Recordset("cl_forpago")
   data_cabezal.Recordset("cl_celular") = data_cabeza2.Recordset("cl_celular") 'descripcion f.pago
   data_cabezal.Recordset("fecha_modi") = data_cabeza2.Recordset("fecha_modi")
   data_cabezal.Recordset("cl_diacobr") = data_cabeza2.Recordset("cl_diacobr")
   data_cabezal.Recordset("cl_nrotarj") = data_cabeza2.Recordset("cl_nrotarj")
   data_cabezal.Recordset("cl_tjemi_n") = data_cabeza2.Recordset("cl_tjemi_n")
   data_cabezal.Recordset("cl_tjemi_c") = data_cabeza2.Recordset("cl_tjemi_c")
   data_cabezal.Recordset("cl_referen") = data_cabeza2.Recordset("cl_referen")
   data_cabezal.Recordset("tit_tarj") = data_cabeza2.Recordset("tit_tarj")
   data_cabezal.Recordset("cl_nomconv") = data_cabeza2.Recordset("cl_nomconv")
    'receptor
   data_cabezal.Recordset("cl_nro_sup") = data_cabeza2.Recordset("cl_nro_sup")
   data_cabezal.Recordset("hora_baja") = data_cabeza2.Recordset("hora_baja")
   data_cabezal.Recordset("cl_nom_sup") = data_cabeza2.Recordset("cl_nom_sup")
        'data_cabezal.Recordset("cl_medflia") = "nro doc para extranjeros"
   data_cabezal.Recordset("info_debit") = data_cabeza2.Recordset("info_debit")
   data_cabezal.Recordset("cl_direcci") = data_cabeza2.Recordset("cl_direcci")
   data_cabezal.Recordset("cl_zona") = data_cabeza2.Recordset("cl_zona")
    'data_cabezal.Recordset("cl_email") = frm_convenios.t_dpto.Text 'opcional
   data_cabezal.Recordset("cl_localid") = data_cabeza2.Recordset("cl_localid") 'opcional
   data_cabezal.Recordset("cl_codigo") = data_cabeza2.Recordset("cl_codigo")
   data_cabezal.Recordset("usu_baja") = data_cabeza2.Recordset("usu_baja") 'moneda
   data_cabezal.Recordset("saldo_chc2") = data_cabeza2.Recordset("saldo_chc2") 'valor dolar
   data_cabezal.Recordset("saldo_cc") = data_cabeza2.Recordset("saldo_cc")  'iva minimo
   data_cabezal.Recordset("saldo_cc2") = data_cabeza2.Recordset("saldo_cc2") 'iva básico
   data_cabezal.Recordset("cl_atrasoa") = data_cabeza2.Recordset("cl_atrasoa") 'subtot iva 22
   data_cabezal.Recordset("cl_cedula") = data_cabeza2.Recordset("cl_cedula") 'subtot iva cero
   data_cabezal.Recordset("saldo_doc2") = data_cabeza2.Recordset("saldo_doc2")
   data_cabezal.Recordset("cl_atrasop") = data_cabeza2.Recordset("cl_atrasop")
   data_cabezal.Recordset("cl_decuota") = data_cabeza2.Recordset("cl_decuota")
   data_cabezal.Recordset("saldo_doc") = data_cabeza2.Recordset("saldo_doc")
   data_cabezal.Recordset("cl_grupo") = data_cabeza2.Recordset("cl_grupo")
   data_cabezal.Recordset("saldo_chc") = data_cabeza2.Recordset("saldo_chc")
   data_cabezal.Recordset("cl_telefon") = data_cabeza2.Recordset("cl_telefon")
   data_cabezal.Recordset("cl_nombre") = data_cabeza2.Recordset("cl_nombre")
   data_cabezal.Recordset("cl_cuopaga") = data_cabeza2.Recordset("cl_cuopaga")
   data_cabezal.Recordset("codmotbaja") = data_cabeza2.Recordset("codmotbaja")
   data_cabezal.Recordset("ultanopmut") = data_cabeza2.Recordset("ultanopmut")
   data_cabezal.Recordset("cl_fultvta") = data_cabeza2.Recordset("cl_fultvta")
   data_cabezal.Recordset("cl_entre") = data_cabeza2.Recordset("cl_entre")
   data_cabezal.Recordset("codmotbaja") = data_cabeza2.Recordset("codmotbaja")
   data_cabezal.Recordset("ultanopmut") = data_cabeza2.Recordset("ultanopmut")
   data_cabezal.Recordset("cl_fultvta") = data_cabeza2.Recordset("cl_fultvta")
   data_cabezal.Recordset("cl_entre") = data_cabeza2.Recordset("cl_entre")
   data_cabezal.Recordset("cl_fultpag") = data_cabeza2.Recordset("cl_fultpag")
   data_cabezal.Recordset("cl_ultmesp") = data_cabeza2.Recordset("cl_ultmesp")
   data_cabezal.Recordset("cl_nomvend") = data_cabeza2.Recordset("cl_nomvend")
   data_cabezal.Recordset("cl_fax") = data_cabeza2.Recordset("cl_fax")
   data_cabezal.Recordset.Update
    'fin de cabezal
   If data_temp.Recordset.RecordCount > 0 Then
      data_temp.Recordset.MoveFirst
      data_lin2.RecordSource = "Select * from deudas where cliente =" & data_temp.Recordset("cod_cli")
      data_lin2.Refresh
      Do While Not data_temp.Recordset.EOF
         If data_temp.Recordset("cod_prod") = 883 Then
         Else
            data_lin2.Recordset.AddNew
            data_lin2.Recordset("cliente") = data_temp.Recordset("cod_cli")
            data_lin2.Recordset("nombre") = Mid(data_temp.Recordset("nom_cli"), 1, 70)
            data_lin2.Recordset("cod_cnv") = data_temp.Recordset("convenio")
            data_lin3.RecordSource = "Select * from convenio where cnv_codigo ='" & data_temp.Recordset("convenio") & "'"
            data_lin3.Refresh
            If data_lin3.Recordset.RecordCount > 0 Then
               data_lin2.Recordset("nom_cnv") = Mid(data_lin3.Recordset("cnv_desc"), 1, 25)
            End If
            data_lin2.Recordset("tipocta") = labserie.Caption
            data_lin2.Recordset("fecha") = data_temp.Recordset("fecha")
            data_lin2.Recordset("tipodoc") = "CRE"
            data_lin2.Recordset("documento") = Val(labnrofact.Caption)
            data_lin2.Recordset("importe") = Val(Label6.Caption)
            data_lin2.Recordset("moneda") = 1
            data_lin2.Recordset("origen") = "E-TICKET NRO." & Trim(labserie.Caption) & " " & Trim(str(Val(labnrofact.Caption)))
            data_lin2.Recordset("nro_superv") = 30
            data_lin2.Recordset("nro_vende") = 1
            If data_temp.Recordset("cod_prod") = 992 Then
               data_lin2.Recordset("mes") = 0
               data_lin2.Recordset("ano") = 0
            Else
               data_lin2.Recordset("mes") = data_temp.Recordset("mes_paga")
               data_lin2.Recordset("ano") = data_temp.Recordset("ano_paga")
            End If
            data_lin2.Recordset("estado_cta") = 1
            data_lin2.Recordset("tiquet") = 0
            data_lin2.Recordset("deudas") = 0
            data_lin2.Recordset("total") = Val(data_afilcons.Recordset("importe_fin"))
            data_lin2.Recordset("servi") = 0
            Xelivadeudas = Val(data_afilcons.Recordset("importe_fin")) * 0.1 / 1.1
            
            data_lin2.Recordset("iva") = Format(Xelivadeudas, "Standard")
            
            If Label2.Caption = "Sin promoción" Then
               data_lin2.Recordset("descimp") = 0
               data_lin2.Recordset("descpor") = 0
            Else
               data_lin2.Recordset("promo") = Label4.Caption
               data_lin2.Recordset("descimp") = -Val(data_afilcons.Recordset("desc_imp"))
               data_lin2.Recordset("descpor") = Val(Label5.Caption)
            End If
            data_lin2.Recordset.Update
         End If
         data_temp.Recordset.MoveNext
      Loop
   End If
   
   If IsNull(data_afilcons.Recordset("sifact")) = False Then
      If data_afilcons.Recordset("sifact") <> 1 Then
         data_afilcons.Recordset.Edit
         data_afilcons.Recordset("sifact") = 1
         data_afilcons.Recordset.Update
      End If
   Else
      data_afilcons.Recordset.Edit
      data_afilcons.Recordset("sifact") = 1
      data_afilcons.Recordset.Update
   End If
   data_imagen.Recordset.AddNew
   data_imagen.Recordset("fecha") = Date
   data_imagen.Recordset("nrofact") = Val(labnrofact.Caption)
   data_imagen.Recordset("serie") = labserie.Caption
   Picture1.Picture = LoadPicture(App.path & "\qr.bmp")
   data_imagen.Recordset.Update
   data_imagen.Refresh

   data_imagen.RecordSource = "Select * from qr where nrofact =" & Val(labnrofact.Caption) & " and serie ='" & labserie.Caption & "'"
   data_imagen.Refresh
   If data_imagen.Recordset.RecordCount > 0 Then
      data_imagen.Recordset.MoveFirst
      data_temp.RecordSource = "Select * from lineas"
      data_temp.Refresh
      If data_temp.Recordset.RecordCount > 0 Then
         data_temp.Recordset.MoveFirst
         Do While Not data_temp.Recordset.EOF
            data_temp.Recordset.Edit
            data_temp.Recordset("qr") = data_imagen.Recordset("qr")
            If data_temp.Recordset("cod_prod") = 883 Then
               If data_temp.Recordset("tot_lin") > 0 Then
                  data_temp.Recordset("tot_lin") = -data_temp.Recordset("tot_lin")
               End If
            End If
            data_temp.Recordset.Update
            data_temp.Recordset.MoveNext
         Loop
      End If
   Else
      MsgBox "No se encontró la imágen QR", vbInformation
   End If
   data_temp.RecordSource = "select * from lineas"
   data_temp.Refresh
   If data_temp.Recordset.RecordCount > 0 Then
      data_temp.Recordset.MoveFirst
   End If
   cr1.ReportFileName = App.path & "\infticksapp3.rpt"
   cr1.Action = 1

   data_cabeza2.Refresh
   If data_cabeza2.Recordset.RecordCount > 0 Then
      data_cabeza2.Recordset.MoveFirst
      Do While Not data_cabeza2.Recordset.EOF
         data_cabeza2.Recordset.Delete
         data_cabeza2.Recordset.MoveNext
      Loop
   End If
   data_cabeza2.Refresh

   data_temp.Refresh
   If data_temp.Recordset.RecordCount > 0 Then
      data_temp.Recordset.MoveFirst
      Do While Not data_temp.Recordset.EOF
         data_temp.Recordset.Delete
         data_temp.Recordset.MoveNext
      Loop
   End If
   data_temp.Refresh

Else
   MsgBox "No hay líneas de facturación"
End If
'terminado

Exit Sub


Algrabaretic:
            If Err.Number = 3155 Then
               MsgBox "ERROR: al guardar registro.", vbCritical
               End
            Else
               MsgBox "ERROR: al guardar e-ticket", vbCritical
               End
            End If

End Sub

Private Function EstaInicializado() As Boolean

    EstaInicializado = False

    If objPosCfe Is Nothing Or Not objPosCfe.Inicializado Then
        MsgBox "Debe inicializar el POS"
        Set objPosCfe = Nothing
        Exit Function
    End If

    EstaInicializado = True
End Function

Private Sub DesplegarInfoEstadoCfe(Mensaje As String, ResultadoCfe As ResultadoCfe)

'On Error GoTo Xxquepasaalenv

    If ResultadoCfe Is Nothing Then
        MsgBox Mensaje
        data_erro.Recordset.AddNew
        data_erro.Recordset("id") = 19
        data_erro.Recordset("fecha") = Date
        data_erro.Recordset("hora") = Format(Time, "HH:mm")
        data_erro.Recordset("nroerr") = 1
        data_erro.Recordset("desc") = Mid(Trim(Mensaje), 1, 130)
        data_erro.Recordset.Update
        Exit Sub
    End If

    If Not ResultadoCfe.OperacionEjecutada Or ResultadoCfe.EstadoCfe Is Nothing Then
        If ResultadoCfe.Mensaje <> vbNullString Then Mensaje = Mensaje & ": " & ResultadoCfe.Mensaje
        MsgBox Mensaje
        data_erro.Recordset.AddNew
        data_erro.Recordset("id") = 19
        data_erro.Recordset("fecha") = Date
        data_erro.Recordset("hora") = Format(Time, "HH:mm")
        data_erro.Recordset("nroerr") = 2
        data_erro.Recordset("desc") = Mid(Trim(Mensaje), 1, 130)
        data_erro.Recordset.Update
        
        Exit Sub
    End If

    If ResultadoCfe.EstadoCfe.Error Then
        Mensaje = Mensaje & ", ocurrió un error"
        If ResultadoCfe.EstadoCfe.Mensaje <> vbNullString Then _
            Mensaje = Mensaje & ": " & ResultadoCfe.EstadoCfe.Mensaje
        MsgBox Mensaje
        data_erro.Recordset.AddNew
        data_erro.Recordset("id") = 19
        data_erro.Recordset("fecha") = Date
        data_erro.Recordset("hora") = Format(Time, "HH:mm")
        data_erro.Recordset("nroerr") = 3
        data_erro.Recordset("desc") = Mid(Trim(Mensaje), 1, 130)
        data_erro.Recordset.Update
        
        Exit Sub
    End If

    If ResultadoCfe.EstadoCfe.SerieNumeroCfe Is Nothing Then
        MsgBox "El CFE no trae número de folio, no se puede terminar la factura"
        data_erro.Recordset.AddNew
        data_erro.Recordset("id") = 19
        data_erro.Recordset("fecha") = Date
        data_erro.Recordset("hora") = Format(Time, "HH:mm")
        data_erro.Recordset("nroerr") = 4
        data_erro.Recordset("desc") = "El CFE no trae número de folio"
        data_erro.Recordset.Update
        
        Exit Sub
    End If

    If ResultadoCfe.EstadoCfe.DatosCae Is Nothing Then
        MsgBox "El CFE no trae datos del CAE, no se puede terminar la factura"
        data_erro.Recordset.AddNew
        data_erro.Recordset("id") = 19
        data_erro.Recordset("fecha") = Date
        data_erro.Recordset("hora") = Format(Time, "HH:mm")
        data_erro.Recordset("nroerr") = 5
        data_erro.Recordset("desc") = "CFE no trae datos del CAE"
        data_erro.Recordset.Update
        
        Exit Sub
    End If

    If (CInt(ResultadoCfe.EstadoCfe.SerieNumeroCfe.TipoCFE) < 200) Then
        Dim strFile As String
        strFile = App.path & "\qr.bmp"
        Dim objresultado As Resultado
        Set objresultado = objPosCfe.GenerarQr(ResultadoCfe.EstadoCfe.DatosQr, 100, strFile)

        Dim strMensaje As String
        strMensaje = "No se pudo generar el QR"

        If objresultado Is Nothing Then
            MsgBox strMensaje
            data_erro.Recordset.AddNew
            data_erro.Recordset("id") = 19
            data_erro.Recordset("fecha") = Date
            data_erro.Recordset("hora") = Format(Time, "HH:mm")
            data_erro.Recordset("nroerr") = 6
            data_erro.Recordset("desc") = Mid(Trim(strMensaje), 1, 130)
            data_erro.Recordset.Update
            
            Exit Sub
        End If

        If Not objresultado.OperacionExitosa Then
            If objresultado.Mensaje <> vbNullString Then strMensaje = strMensaje & ": " & objresultado.Mensaje
            MsgBox strMensaje
            data_erro.Recordset.AddNew
            data_erro.Recordset("id") = 19
            data_erro.Recordset("fecha") = Date
            data_erro.Recordset("hora") = Format(Time, "HH:mm")
            data_erro.Recordset("nroerr") = 7
            data_erro.Recordset("desc") = Mid(Trim(strMensaje), 1, 130)
            data_erro.Recordset.Update
            
            Exit Sub
        End If

'        imgQr.Picture = LoadPicture(strFile)
    End If
    If Val(ResultadoCfe.EstadoCfe.CodigoRespuesta) = 0 Or _
       Val(ResultadoCfe.EstadoCfe.CodigoRespuesta) = 11 Then
       labserie.Caption = ResultadoCfe.EstadoCfe.SerieNumeroCfe.serie
       labnrofact.Caption = ResultadoCfe.EstadoCfe.SerieNumeroCfe.Numero
       
       data_cabeza2.Recordset.Edit
       data_cabeza2.Recordset("cl_socmnro") = labserie.Caption
       data_cabeza2.Recordset("cl_numero") = Val(labnrofact.Caption)
       labvence.Caption = CStr(ResultadoCfe.EstadoCfe.DatosCae.Vencimiento)
       labautoriza.Caption = CStr(ResultadoCfe.EstadoCfe.DatosCae.Autorizacion)
       labdesde.Caption = labserie.Caption & " " & CStr(ResultadoCfe.EstadoCfe.DatosCae.NumeroDesde) & " - " & labserie.Caption & " " & CStr(ResultadoCfe.EstadoCfe.DatosCae.NumeroHasta)
       labhasta.Caption = CStr(ResultadoCfe.EstadoCfe.CodigoSeguridad)
       If Len(labvence.Caption) = 8 Then
          labvenceok.Caption = Mid(labvence.Caption, 7, 2) & "/" & Mid(labvence.Caption, 5, 2) & "/" & Mid(labvence.Caption, 1, 4)
       Else
          labvenceok.Caption = "31/12/2016"
       End If
       If labvenceok.Caption <> "" Then
          data_cabeza2.Recordset("cl_fultpag") = CDate(labvenceok.Caption)
       Else
          data_cabeza2.Recordset("cl_fultpag") = CDate("01/01/2018")
       End If
       If labautoriza.Caption <> "" Then
          data_cabeza2.Recordset("cl_nrocobr") = Val(labautoriza.Caption)
       Else
          data_cabeza2.Recordset("cl_nrocobr") = 0
       End If
       data_cabeza2.Recordset("cl_medflia") = Trim(labdesde.Caption)
       data_cabeza2.Recordset("cl_fax") = Trim(labhasta.Caption)
       data_cabeza2.Recordset.Update
       
       Dim objResultado44 As ResultadoObtenerQr
       Set objResultado44 = objPosCfe.ObtenerQr(ResultadoCfe.EstadoCfe.DatosQr, 100)
       
       Dim strFile2 As String
       strFile2 = App.path & "\qr.bmp"
       Dim f As Long
       f = FreeFile()
       Open strFile2 For Binary As #f
       Put #f, , objResultado44.ImagenQr
       Close #f

        Set objUltimaSerieNumero = Nothing
 
        If Not objUltimaSerieNumero Is Nothing Then _
        
        End If
    
    
    Else
       data_erro.Recordset.AddNew
       data_erro.Recordset("id") = ResultadoCfe.EstadoCfe.CodigoRespuesta
       data_erro.Recordset("fecha") = Date
       data_erro.Recordset("hora") = Format(Time, "HH:mm")
       data_erro.Recordset("nroerr") = 7
       data_erro.Recordset("desc") = "ERROR EN EL RESULTADO CFE"
       data_erro.Recordset.Update
       MsgBox "Error al terminar la factura, verifique datos!", vbCritical
       End
    End If
       

    strUltimoGuid = ResultadoCfe.EstadoCfe.Guid
    Set objUltimaSerieNumero = ResultadoCfe.EstadoCfe.SerieNumeroCfe

    'cmdConsultaXguid.Enabled = True
    'cmdConsultaXnumero.Enabled = True

End Sub


Public Function Devuelve_vende() As String

Dim Xsqlpromo As String
Dim Xrecclii As New ADODB.Recordset

ConectarBD
ConbdSapp.Open
             
Xsqlpromo = "Select * from vende_func where idfunc =" & data_afilcons.Recordset("codvende")
With Xrecclii
    .CursorLocation = adUseClient
    .CursorType = adOpenKeyset
    .LockType = adLockOptimistic
    .Open Xsqlpromo, ConbdSapp, , , adCmdText
End With
If Xrecclii.RecordCount > 0 Then
   Devuelve_vende = Xrecclii("nombre")
Else
   Devuelve_vende = ""
End If

Xrecclii.Close
ConbdSapp.Close

End Function

Private Sub Form_Resize()
With Image1
   .Top = 0
   .Left = 0
   .Width = Me.Width
   .Height = Me.Height
End With

End Sub

Public Sub Crear_Mod()
Dim Cl_apellid, Cl_entre, Cl_dir, VenceT, MutAfil, VendeAfil As String
Dim Xmatnew As Long
Xmatnew = 0
Cl_apellid = ""
Cl_entre = ""
VenceT = ""
Cl_dir = ""
MutAfil = ""
VendeAfil = ""

If IsNull(data_afilcons.Recordset("matricula")) = False Then
   data_cli.RecordSource = "select * from clientes where cl_codigo =" & data_afilcons.Recordset("matricula")
Else
   data_cli.RecordSource = "select * from clientes where cl_codigo =" & 0
End If
data_cli.Refresh

If data_cli.Recordset.RecordCount > 0 And IsNull(data_afilcons.Recordset("matricula")) = False Then
   If IsNull(data_cli.Recordset("fecha_baja")) = False Then
      data_cli.Recordset.Edit
      data_cli.Recordset("fecha_baja") = Null
      data_cli.Recordset.Update
   End If
   If IsNull(data_cli.Recordset("estado")) = False Then
      If data_cli.Recordset("estado") <> 1 Then
         data_cli.Recordset.Edit
         data_cli.Recordset("estado") = 1
         data_cli.Recordset.Update
      End If
   Else
      data_cli.Recordset.Edit
      data_cli.Recordset("estado") = 1
      data_cli.Recordset.Update
   End If
   If IsNull(data_afilcons.Recordset("catreal")) = False Then
      If data_cli.Recordset("cl_codconv") <> data_afilcons.Recordset("catreal") Then
         data_cli.Recordset.Edit
         data_cli.Recordset("cl_codconv") = data_afilcons.Recordset("catreal")
         data_cli.Recordset("cl_nomconv") = data_afilcons.Recordset("catrealdes")
         data_cli.Recordset.Update
      End If
   End If
   
End If

End Sub
