VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_debitos 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proceso débitos automáticos"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6630
   Icon            =   "frm_debitos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   6630
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog cmm1 
      Left            =   3120
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
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
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc data_emi 
      Height          =   375
      Left            =   3600
      Top             =   120
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
      DataSourceName  =   "sappnew"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "data_emi"
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
   Begin MSAdodcLib.Adodc data1 
      Height          =   375
      Left            =   4680
      Top             =   4200
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      DataSourceName  =   "sappnew"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "data1"
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
      Height          =   330
      Left            =   4560
      Top             =   4440
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      DataSourceName  =   "sappnew"
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
   Begin Crystal.CrystalReport CR1 
      Left            =   6120
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "FoxPro 2.6;"
      DatabaseName    =   "C:\debitos"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DEVAUT"
      Top             =   4920
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Datos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6135
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   1320
         Top             =   2280
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "Adodc2"
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
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   2760
         Top             =   2520
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
      Begin VB.Data data_redpagos 
         Caption         =   "data_redpagos"
         Connect         =   "Excel 8.0;"
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
         Top             =   3360
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.Data data_deudas 
         Caption         =   "data_deudas"
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
         Top             =   360
         Visible         =   0   'False
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   2280
         Visible         =   0   'False
         Width           =   5775
      End
      Begin VB.CommandButton b_mut 
         BackColor       =   &H00808080&
         Caption         =   "Cargar mutuales"
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
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2760
         Width           =   1695
      End
      Begin VB.CommandButton b_devol 
         BackColor       =   &H00808080&
         Caption         =   "Procesar Devoluciones"
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
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3840
         Width           =   1695
      End
      Begin VB.Data data_mutuales 
         Caption         =   "data_mutuales"
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
         Top             =   1440
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   4200
         TabIndex        =   11
         Top             =   1560
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSAdodcLib.Adodc data_arq 
         Height          =   330
         Left            =   1920
         Top             =   2400
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
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
         Caption         =   "data_arq"
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
      Begin VB.TextBox t_codced 
         Height          =   285
         Left            =   4200
         TabIndex        =   10
         Top             =   1200
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton b_sal 
         BackColor       =   &H00808080&
         Caption         =   "SALIR"
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
         Left            =   3600
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3840
         Width           =   1695
      End
      Begin VB.CommandButton b_proc 
         BackColor       =   &H00808080&
         Caption         =   "Procesar Débitos"
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
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2760
         Width           =   1695
      End
      Begin MSMask.MaskEdBox mfec 
         Height          =   375
         Left            =   2160
         TabIndex        =   7
         Top             =   1800
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
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
      Begin VB.TextBox txt_ano 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         MaxLength       =   4
         TabIndex        =   5
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox txt_mes 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         MaxLength       =   2
         TabIndex        =   4
         Top             =   1080
         Width           =   615
      End
      Begin VB.ComboBox cbotarj 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "frm_debitos.frx":0442
         Left            =   2160
         List            =   "frm_debitos.frx":0458
         TabIndex        =   1
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "FECHA DÉBITO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
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
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "MES/AÑO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "TARJETA:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   120
      Picture         =   "frm_debitos.frx":048A
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   735
   End
End
Attribute VB_Name = "frm_debitos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim txtpath, txtpath1 As String
Public txtfile As ADODB.Stream
Private Sub b_devol_Click()
Dim Linea, Recibo, Noestan As String
Dim Xind, Xnrorec, XlineaCant, XlineaBrou, XlineaBrouL, Xrecdesde, Xrechasta As Long
Dim ConfirmaDev, Letrabrou, XtextoMaster As String
Dim Xpipe, Xpipeprimero As Integer
XtextoMaster = ""
Xpipe = 0
Xpipeprimero = 0
ConfirmaDev = ""
Xind = 0
XlineaCant = 0
Linea = ""
Recibo = ""
Xnrorec = 0
Text1.Text = ""
Letrabrou = ""
Noestan = ""
XlineaBrou = 0
XlineaBrouL = 0
Xrecdesde = 0
Xrechasta = 0
If cbotarj.ListIndex >= 0 Then
    If cbotarj.ListIndex = 5 Then 'REDPAGOS
       Dim Xdesea As String
       Xdesea = MsgBox("Verifique que esté guardado el archivo redpagos.xls en la carpeta DEBITOS del disco C.", vbYesNo + vbInformation, "Débitos")
       frm_debitos.MousePointer = 11
       data_redpagos.DatabaseName = "C:\debitos\redpagos.xls"
       data_redpagos.RecordSource = "redpagos$"
       data_redpagos.Refresh
       If Xdesea = vbYes Then
            If data_redpagos.Recordset.RecordCount > 0 Then
               data_redpagos.Recordset.MoveFirst
               Do While Not data_redpagos.Recordset.EOF
                  data_emi.RecordSource = "select * from arqueo where nrorec =" & data_redpagos.Recordset("factura") & " and cob =" & 221
                  data_emi.Refresh
                  If data_emi.Recordset.RecordCount > 0 Then
                     If data_emi.Recordset("arqueo") <> "C" Then
                        data_emi.Recordset("arqueo") = "C"
                        data_emi.Recordset("fecha") = data_redpagos.Recordset("fecha")
                        data_emi.Recordset("usuar") = WElusuario
                        data_emi.Recordset.Update
                     End If
                  End If
                  data_emi.RecordSource = "select * from deudas where documento =" & data_redpagos.Recordset("factura") & " and nro_cobr =" & 221
                  data_emi.Refresh
                  If data_emi.Recordset.RecordCount > 0 Then
                     If IsNull(data_emi.Recordset("fecha_pago")) = False Then
                     Else
                        data_emi.Recordset("fecha_pago") = data_redpagos.Recordset("fecha")
                        data_emi.Recordset.Update
                     End If
                  End If
                  data_redpagos.Recordset.MoveNext
               Loop
               frm_debitos.MousePointer = 0
               MsgBox "Proceso terminado.", vbInformation
                 'Al generar el arqueo ingresar cómo pendiente los cobradores redpagos y tarjetas
            End If
       End If
    Else
        With cmm1
             .FileName = ""
             .Filter = "TXT (*.txt;) | *.txt;"
             .ShowOpen
             If Len(.FileName) <> 0 Then
                txtpath = .FileName
                txtpath1 = .FileTitle
                Text1.Text = .FileName
             End If
        '     t_id.Text = 10
        End With
        
        If Text1.Text <> "" Then
           ConfirmaDev = MsgBox("Desea procesar las devoluciones de " & cbotarj.Text & "?", vbInformation + vbYesNo, "Débitos")
           If ConfirmaDev = vbYes Then
              frm_debitos.MousePointer = 11
              Open Text1.Text For Input As #1
              If cbotarj.ListIndex = 3 Then 'brou ---ok 607
                 data_emi.RecordSource = "select * from arqueo where cob =" & 607 & " and arqueo in ('P','E')"
                 data_emi.Refresh
                 If data_emi.Recordset.RecordCount > 0 Then
                    data_emi.Recordset.MoveFirst
                    Do While Not data_emi.Recordset.EOF
                       data_emi.Recordset("fecha") = Date
                       data_emi.Recordset("arqueo") = "C"
                       data_emi.Recordset("usuar") = WElusuario
                       data_emi.Recordset.Update
                       data_emi.Recordset.MoveNext
                    Loop
                 End If
                 Do While Not EOF(1)
                    Line Input #1, Linea
                    XlineaBrou = 187
                    XlineaBrouL = 2
                    Xrecdesde = 129
                    Xrechasta = 140
                    For Xind = 1 To Len(Linea)
                        If Mid(Linea, Xind, 1) = "N" Then
                           XlineaBrou = Xind + 129
                           Recibo = Mid(Linea, XlineaBrou, 10)
                           Xnrorec = Val(Recibo)
                           '' ok acá seguir
                           data_emi.RecordSource = "select * from arqueo where cob=" & 607 & " and nrorec =" & Xnrorec
                           data_emi.Refresh
                           If data_emi.Recordset.RecordCount > 0 Then
                              If data_emi.Recordset("arqueo") <> "P" Then
                                 data_emi.Recordset("arqueo") = "P"
                                 data_emi.Recordset.Update
                              End If
                           Else
                              data_emi.RecordSource = "select * from mutuales where recibo =" & Xnrorec
                              data_emi.Refresh
                              If data_emi.Recordset.RecordCount > 0 Then
                                 MsgBox "ATENCION! ANOTE: Pendiente de pago recibo mutual " & Trim(str(Xnrorec)) & " $." & Format(data_emi.Recordset("importe_deuda"), "Standard")
                              End If
                           End If
                        End If
                    Next Xind
                    data_emi.RecordSource = "select * from arqueo where cob =" & 607 & " and arqueo in ('C')"
                    data_emi.Refresh
                    If data_emi.Recordset.RecordCount > 0 Then
                       data_emi.Recordset.MoveFirst
                       Do While Not data_emi.Recordset.EOF
                          data_deudas.RecordSource = "select * from deudas where documento =" & data_emi.Recordset("nrorec") & " and fecha_pago is null"
                          data_deudas.Refresh
                          If data_deudas.Recordset.RecordCount > 0 Then
                             data_deudas.Recordset.Edit
                             data_deudas.Recordset("fecha_pago") = Date
                             data_deudas.Recordset.Update
                          End If
                          data_emi.Recordset.MoveNext
                       Loop
                    End If
                    Recibo = ""
                 Loop
                 frm_debitos.MousePointer = 0
        '         If Trim(Noestan) <> "" Then
                 MsgBox "Proceso terminado.", vbExclamation
        '         End If
            '1181115
              End If
              
              If cbotarj.ListIndex = 1 Then 'master ---ok 683
                 data_emi.RecordSource = "select * from arqueo where cob =" & 683 & " and arqueo not in ('B','D','C')"
                 data_emi.Refresh
                 If data_emi.Recordset.RecordCount > 0 Then
                    data_emi.Recordset.MoveFirst
                    Do While Not data_emi.Recordset.EOF
                       data_emi.Recordset("fecha") = Date
                       data_emi.Recordset("arqueo") = "C"
                       data_emi.Recordset("usuar") = WElusuario
                       data_emi.Recordset.Update
                       data_emi.Recordset.MoveNext
                    Loop
                 End If
                 Do While Not EOF(1)
                    Line Input #1, Linea
                    XlineaCant = 0
                    For Xind = 1 To Len(Linea)
                        If Trim(Mid(Linea, Xind, 1)) = "|" Then
                           Xpipe = Xpipe + 1
                           If XlineaCant = 0 Then
                              XlineaCant = 9
                           End If
                        Else
                           If XlineaCant = 0 Then
                              Recibo = Recibo & Mid(Linea, Xind, 1)
                           End If
                        End If
                        If Xpipe = 7 Then
                           If Xpipeprimero = 0 Then
                           Else
                              XtextoMaster = XtextoMaster & Mid(Linea, Xind, 1)
                           End If
                           Xpipeprimero = Xpipeprimero + 1
                        End If
                    Next Xind
                    If Mid(XtextoMaster, 1, 23) = "Observado-Modifique el " Then
                       Xnrorec = 0
                    Else
                       Xnrorec = Val(Recibo)
                    End If
                    XlineaCant = 0
                    XtextoMaster = ""
                    Xpipe = 0
                    Xpipeprimero = 0
                    data_emi.RecordSource = "select * from arqueo where cob=" & 683 & " and nrorec =" & Xnrorec & " and arqueo in ('C')"
                    data_emi.Refresh
                    If data_emi.Recordset.RecordCount > 0 Then
                       data_emi.Recordset("arqueo") = "P"
                       data_emi.Recordset("fecha") = Date
                       data_emi.Recordset("usuar") = WElusuario
                       data_emi.Recordset.Update
                    Else
                       data_emi.RecordSource = "select * from mutuales where recibo =" & Xnrorec
                       data_emi.Refresh
                       If data_emi.Recordset.RecordCount > 0 Then
                          MsgBox "ATENCION! ANOTE: Pendiente de pago recibo mutual " & Trim(str(Xnrorec)) & " $." & Format(data_emi.Recordset("importe_deuda"), "Standard")
                       End If
                    End If
                    Recibo = ""
                 Loop
                 data_emi.RecordSource = "select * from arqueo where cob=" & 683 & " and arqueo in ('C')"
                 data_emi.Refresh
                 If data_emi.Recordset.RecordCount > 0 Then
                    data_emi.Recordset.MoveFirst
                    Do While Not data_emi.Recordset.EOF
                       data_deudas.RecordSource = "select * from deudas where documento =" & data_emi.Recordset("nrorec") & " and nro_cobr =" & 683 & " and fecha_pago is null"
                       data_deudas.Refresh
                       If data_deudas.Recordset.RecordCount > 0 Then
                          data_deudas.Recordset.Edit
                          data_deudas.Recordset("fecha_pago") = Date
                          data_deudas.Recordset.Update
                       End If
                       data_emi.Recordset.MoveNext
                    Loop
                 End If
                 
                 frm_debitos.MousePointer = 0
        '         If Trim(Noestan) <> "" Then
                 MsgBox "Proceso terminado.", vbExclamation
        '         End If
            '1181115
              End If
              
              If cbotarj.ListIndex = 2 Then 'CABAL ---ok 673
                 data_emi.RecordSource = "select * from arqueo where cob =" & 673 & " and arqueo not in ('C')"
                 data_emi.Refresh
                 If data_emi.Recordset.RecordCount > 0 Then
                    data_emi.Recordset.MoveFirst
                    Do While Not data_emi.Recordset.EOF
                       data_emi.Recordset("fecha") = Date
                       data_emi.Recordset("arqueo") = "C"
                       data_emi.Recordset("usuar") = WElusuario
                       data_emi.Recordset.Update
                       data_deudas.RecordSource = "select * from deudas where documento =" & data_emi.Recordset("nrorec") & " and nro_cobr =" & 673 & " and fecha_pago is null"
                       data_deudas.Refresh
                       If data_deudas.Recordset.RecordCount > 0 Then
                          data_deudas.Recordset.Edit
                          data_deudas.Recordset("fecha_pago") = Date
                          data_deudas.Recordset.Update
                       End If
                       data_emi.Recordset.MoveNext
                    Loop
                 End If
                 Do While Not EOF(1)
                    XlineaCant = XlineaCant + 1
                    Line Input #1, Linea
                    Recibo = ""
                    If XlineaCant >= 17 Then
                       For Xind = 1 To Len(Linea)
                           If Trim(Linea) <> "" Then
                              If Xind >= 88 Then
                                 If Xind <= 101 Then
                                    Recibo = Recibo & Mid(Linea, Xind, 1)
                                 End If
                              End If
                           End If
                       Next Xind
                    End If
                    If Trim(Recibo) <> "" Then
                        Xnrorec = Val(Recibo)
                        data_emi.RecordSource = "select * from arqueo where cob=" & 673 & " and nrorec =" & Xnrorec
                        data_emi.Refresh
                        If data_emi.Recordset.RecordCount > 0 Then
                           data_emi.Recordset("fecha") = Date
                           data_emi.Recordset("arqueo") = "P"
                           data_emi.Recordset.Update
                        Else
                           data_emi.RecordSource = "select * from mutuales where recibo =" & Xnrorec
                           data_emi.Refresh
                           If data_emi.Recordset.RecordCount > 0 Then
                              MsgBox "ATENCION! ANOTE: Pendiente de pago recibo mutual " & Trim(str(Xnrorec)) & " $." & Format(data_emi.Recordset("importe_deuda"), "Standard")
                           End If
                        End If
                        data_emi.RecordSource = "select * from deudas where documento =" & Xnrorec & " and nro_cobr =" & 673 & " and fecha_pago is not null"
                        data_emi.Refresh
                        If data_emi.Recordset.RecordCount > 0 Then
                           Adodc2.RecordSource = "select * from linmmdd where cod_cli =" & data_emi.Recordset("cliente") & " and cod_prod in (999) and mes_paga =" & data_emi.Recordset("mes") & " and ano_paga =" & data_emi.Recordset("ano")
                           Adodc2.Refresh
                           If Adodc2.Recordset.RecordCount > 0 Then
                           Else
                              data_emi.Recordset("fecha_pago") = Null
                              data_emi.Recordset.Update
                           End If
                        End If
                    End If
                 Loop
                 frm_debitos.MousePointer = 0
                 MsgBox "Proceso terminado.", vbExclamation
              End If
              
              If cbotarj.ListIndex = 0 Then 'VISA ---ok 514
                 data_emi.RecordSource = "select * from arqueo where cob =" & 514 & " and arqueo not in ('C','D','B')"
                 data_emi.Refresh
                 If data_emi.Recordset.RecordCount > 0 Then
                    data_emi.Recordset.MoveFirst
                    Do While Not data_emi.Recordset.EOF
                       data_emi.Recordset("fecha") = Date
                       data_emi.Recordset("arqueo") = "C"
                       data_emi.Recordset("usuar") = WElusuario
                       data_emi.Recordset.Update
                       data_emi.Recordset.MoveNext
                    Loop
                    data_emi.RecordSource = "select * from deudas where nro_cobr =" & 514 & " and fecha_pago is null"
                    data_emi.Refresh
                    If data_emi.Recordset.RecordCount > 0 Then
                       data_emi.Recordset.MoveFirst
                       Do While Not data_emi.Recordset.EOF
                          Adodc1.RecordSource = "select * from arqueo where arqueo in ('C','D') and nrorec =" & data_emi.Recordset("documento") & " and cob =" & 514
                          Adodc1.Refresh
                          If Adodc1.Recordset.RecordCount > 0 Then
                             data_emi.Recordset("fecha_pago") = Date
                             data_emi.Recordset.Update
                          End If
                          data_emi.Recordset.MoveNext
                       Loop
                    End If
                 End If
                 Do While Not EOF(1)
                    XlineaCant = XlineaCant + 1
                    Line Input #1, Linea
                    Recibo = ""
                    For Xind = 1 To Len(Linea)
                        If Trim(Linea) <> "" Then
                           If Xind >= 83 Then
                              If Xind <= 91 Then
                                 Recibo = Recibo & Mid(Linea, Xind, 1)
                              End If
                           End If
                        End If
                    Next Xind
                    If Trim(Recibo) <> "" Then
                       Xnrorec = Val(Recibo)
                       data_emi.RecordSource = "select * from arqueo where cob=" & 514 & " and matricula =" & Xnrorec & " and arqueo in ('C') order by fecha DESC"
                       data_emi.Refresh
                       If data_emi.Recordset.RecordCount > 0 Then
                          data_emi.Recordset("fecha") = Date
                          data_emi.Recordset("arqueo") = "P"
                          data_emi.Recordset.Update
                       Else
                          data_emi.RecordSource = "select * from mutuales where recibo =" & Xnrorec
                          data_emi.Refresh
                          If data_emi.Recordset.RecordCount > 0 Then
                             MsgBox "ATENCION! ANOTE: Pendiente de pago recibo mutual " & Trim(str(Xnrorec)) & " $." & Format(data_emi.Recordset("importe_deuda"), "Standard")
                          End If
                       End If
                       data_emi.RecordSource = "select * from deudas where cliente =" & Xnrorec & " and nro_cobr =" & 514 & " and fecha_pago is not null order by fecha DESC"
                       data_emi.Refresh
                       If data_emi.Recordset.RecordCount > 0 Then
                          Adodc2.RecordSource = "select * from linmmdd where cod_cli =" & data_emi.Recordset("cliente") & " and cod_prod in (999) and mes_paga =" & data_emi.Recordset("mes") & " and ano_paga =" & data_emi.Recordset("ano")
                          Adodc2.Refresh
                          If Adodc2.Recordset.RecordCount > 0 Then
                          Else
                             data_emi.Recordset("fecha_pago") = Null
                             data_emi.Recordset.Update
                          End If
                       End If
                    End If
                 Loop
                 frm_debitos.MousePointer = 0
                 MsgBox "Proceso terminado.", vbExclamation
              End If
              
              If cbotarj.ListIndex = 4 Then 'OCA ---ok 690
                 data_emi.RecordSource = "select * from arqueo where cob =" & 690 & " and arqueo not in ('C','D','B')"
                 data_emi.Refresh
                 If data_emi.Recordset.RecordCount > 0 Then
                    data_emi.Recordset.MoveFirst
                    Do While Not data_emi.Recordset.EOF
                       data_emi.Recordset("fecha") = Date
                       data_emi.Recordset("arqueo") = "C"
                       data_emi.Recordset("usuar") = WElusuario
                       data_emi.Recordset.Update
                       data_emi.Recordset.MoveNext
                    Loop
                    data_emi.RecordSource = "select * from deudas where nro_cobr =" & 690 & " and fecha_pago is null"
                    data_emi.Refresh
                    If data_emi.Recordset.RecordCount > 0 Then
                       data_emi.Recordset.MoveFirst
                       Do While Not data_emi.Recordset.EOF
                          Adodc1.RecordSource = "select * from arqueo where arqueo in ('C','D') and nrorec =" & data_emi.Recordset("documento") & " and cob =" & 690
                          Adodc1.Refresh
                          If Adodc1.Recordset.RecordCount > 0 Then
                             data_emi.Recordset("fecha_pago") = Date
                             data_emi.Recordset.Update
                          End If
                          data_emi.Recordset.MoveNext
                       Loop
                    End If
                 End If
                 Do While Not EOF(1)
                    XlineaCant = XlineaCant + 1
                    Line Input #1, Linea
                    Recibo = ""
                    For Xind = 1 To Len(Linea)
                        If Trim(Linea) <> "" Then
                           If Xind >= 76 Then
                              If Xind <= 95 Then
                                 Recibo = Recibo & Mid(Linea, Xind, 1)
                              End If
                           End If
                        End If
                    Next Xind
                    If Trim(Recibo) <> "" Then
                       Xnrorec = Val(Recibo)
                       data_emi.RecordSource = "select * from arqueo where cob=" & 690 & " and nrorec =" & Xnrorec & " and arqueo in ('C')"
                       data_emi.Refresh
                       If data_emi.Recordset.RecordCount > 0 Then
                          data_emi.Recordset("fecha") = Date
                          data_emi.Recordset("arqueo") = "P"
                          data_emi.Recordset.Update
                       Else
                          data_emi.RecordSource = "select * from mutuales where recibo =" & Xnrorec
                          data_emi.Refresh
                          If data_emi.Recordset.RecordCount > 0 Then
                             MsgBox "ATENCION! ANOTE: Pendiente de pago recibo mutual " & Trim(str(Xnrorec)) & " $." & Format(data_emi.Recordset("importe_deuda"), "Standard")
                          End If
                       End If
                       data_emi.RecordSource = "select * from deudas where documento =" & Xnrorec & " and nro_cobr =" & 690 & " and fecha_pago is not null"
                       data_emi.Refresh
                       If data_emi.Recordset.RecordCount > 0 Then
                          Adodc2.RecordSource = "select * from linmmdd where cod_cli =" & data_emi.Recordset("cliente") & " and cod_prod in (999) and mes_paga =" & data_emi.Recordset("mes") & " and ano_paga =" & data_emi.Recordset("ano")
                          Adodc2.Refresh
                          If Adodc2.Recordset.RecordCount > 0 Then
                          Else
                             data_emi.Recordset("fecha_pago") = Null
                             data_emi.Recordset.Update
                          End If
                       End If
                    End If
                 Loop
                 frm_debitos.MousePointer = 0
                 MsgBox "Proceso terminado.", vbExclamation
              End If
              Close #1
           End If
        End If
    End If
Else
   MsgBox "Seleccione tarjeta a procesar.", vbExclamation
End If

End Sub

Private Sub b_mut_Click()
frm_debmut.Show vbModal

End Sub

Private Sub b_proc_Click()
Dim Xlin, Xtarj, Xfec, XCantex, Xnommes, Xnomimp, Xnomced, Xmesmast, Xmescabal, Xnommescab As String
Dim XP, Xcant As Integer
Dim XImp, Xtot, Xtardelf, Ximparqueo As Double
Dim Xtitulo As String
Dim XNombre, Xnomarq As String
Dim Xventar As String
Dim Xmatstr, Xnrorecstr, Ximpgrastr, Ximpgrastr2 As String
Dim Ximpgrav As Double
Dim Xelcoddevuelto As Integer


Xcant = 0

Dim Xmes, Xano As Integer

XNombre = "emi"
Xnomarq = "arq"
If txt_mes.Text > 9 Then
   XNombre = XNombre + Trim(txt_mes.Text) + Mid(Trim(txt_ano.Text), 3, 2)
   Xmes = txt_mes.Text - 1
   Xano = txt_ano.Text
   Xnomarq = Xnomarq + Trim(str(Xmes)) + Mid(Trim(str(Xano)), 3, 2)
Else
   XNombre = XNombre + "0" + Trim(txt_mes.Text) + Mid(Trim(txt_ano.Text), 3, 2)
   If txt_mes.Text = 1 Then
      Xmes = 12
      Xano = txt_ano.Text - 1
   Else
      Xmes = txt_mes.Text - 1
      Xano = txt_ano.Text
   End If
   If Xmes > 9 Then
      Xnomarq = Xnomarq + Trim(str(Xmes)) + Mid(Trim(str(Xano)), 3, 2)
   Else
      Xnomarq = Xnomarq + "0" + Trim(str(Xmes)) + Mid(Trim(str(Xano)), 3, 2)
   End If
End If
If Day(Date) <= 11 Then
   Xnomarq = "arqueo"
End If

frm_debitos.MousePointer = 11
b_proc.Enabled = False
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infcli"
data_inf.RecordSource = "infcli"
data_inf.Refresh


If mfec.Text <> "__/__/____" Then
   Xfec = Trim(Mid(mfec.Text, 7, 4))
   Xfec = Xfec + Trim(Mid(mfec.Text, 4, 2))
   Xfec = Xfec + Trim(Mid(mfec.Text, 1, 2))
   Xmesmast = Trim(Mid(mfec.Text, 4, 2))
   Xmesmast = Xmesmast + "/" + Trim(Mid(mfec.Text, 9, 2))
   Xnommescab = Trim(Mid(mfec.Text, 4, 2))
   Xnommescab = Xnommescab + Trim(Mid(mfec.Text, 9, 2))
   Xmescabal = Trim(Mid(mfec.Text, 1, 2))
   Xmescabal = Xmescabal + Trim(Mid(mfec.Text, 4, 2))
   Xmescabal = Xmescabal + Trim(Mid(mfec.Text, 9, 2))
End If
Dim Ximpivag As Double
Dim Ximpivagstr As String

If Month(mfec.Text) = 1 Then
   Xnommes = "ENE/" + Trim(Mid(mfec.Text, 9, 2))
End If
If Month(mfec.Text) = 2 Then
   Xnommes = "FEB/" + Trim(Mid(mfec.Text, 9, 2))
End If
If Month(mfec.Text) = 3 Then
   Xnommes = "MAR/" + Trim(Mid(mfec.Text, 9, 2))
End If
If Month(mfec.Text) = 4 Then
   Xnommes = "ABR/" + Trim(Mid(mfec.Text, 9, 2))
End If
If Month(mfec.Text) = 5 Then
   Xnommes = "MAY/" + Trim(Mid(mfec.Text, 9, 2))
End If
If Month(mfec.Text) = 6 Then
   Xnommes = "JUN/" + Trim(Mid(mfec.Text, 9, 2))
End If
If Month(mfec.Text) = 7 Then
   Xnommes = "JUL/" + Trim(Mid(mfec.Text, 9, 2))
End If
If Month(mfec.Text) = 8 Then
   Xnommes = "AGO/" + Trim(Mid(mfec.Text, 9, 2))
End If
If Month(mfec.Text) = 9 Then
   Xnommes = "SET/" + Trim(Mid(mfec.Text, 9, 2))
End If
If Month(mfec.Text) = 10 Then
   Xnommes = "OCT/" + Trim(Mid(mfec.Text, 9, 2))
End If
If Month(mfec.Text) = 11 Then
   Xnommes = "NOV/" + Trim(Mid(mfec.Text, 9, 2))
End If
If Month(mfec.Text) = 12 Then
   Xnommes = "DIC/" + Trim(Mid(mfec.Text, 9, 2))
End If
Dim control As String

If cbotarj.ListIndex = 0 Then 'visa ---okkk
   Proceso_visa
End If
If cbotarj.ListIndex = 1 Then 'master --ok
   Proceso_master
   data_inf.RecordSource = "select * from infcli"
   data_inf.Refresh
   Xtitulo = "DEBITOS MASTERCARD CORRESPONDIENTES MES : " & txt_mes.Text & "/" & txt_ano.Text
   cr1.ReportFileName = App.path & "\infdebitos.rpt"
   cr1.ReportTitle = Xtitulo
   cr1.Action = 1
   MsgBox "Proceso terminado!!", vbInformation, "Mensaje"
End If
'8 a 20

If cbotarj.ListIndex = 2 Then ''' CABAL --ok
   Dim Ximpcabal, Xivacabal As Double
   Ximpcabal = 0
   data_cli.RecordSource = "Select * from clientes where cl_nrocobr =" & 673
   data_cli.Refresh
   If data_cli.Recordset.RecordCount > 0 Then
      data_cli.Recordset.MoveFirst
      Open "c:\debitos\LISTAGRA.ROU" For Output As #1
      Do While Not data_cli.Recordset.EOF
         XImp = 0
         If IsNull(data_cli.Recordset("cl_nrotarj")) = False Then
            If Len(Trim(data_cli.Recordset("cl_nrotarj"))) = 16 Then
               If IsNull(data_cli.Recordset("estado")) = False Then
                  If data_cli.Recordset("estado") = 1 Or data_cli.Recordset("estado") = 0 Then
                     data_emi.RecordSource = "select * from deudas where cliente =" & data_cli.Recordset("cl_codigo") & " and fecha_pago is null and fecha >='" & Format("01/10/2021", "yyyy/mm/dd") & "' and nro_cobr =" & 673 & " and mes >" & 0
                     data_emi.Refresh
                     If data_emi.Recordset.RecordCount > 0 Then
                        data_emi.Recordset.MoveFirst
                        Do While Not data_emi.Recordset.EOF
                             Xlin = "CBCU28443915006N"
                             Xtarj = Trim(data_cli.Recordset("cl_nrotarj"))
                             Xlin = Xlin + Trim(Xtarj)
                             Call controlced(data_cli.Recordset("ci_tarj"), data_cli.Recordset("codcitarj"))
                             If Int(Val(t_codced.Text)) = Int(data_cli.Recordset("codcitarj")) Then
                             Else
                                MsgBox "Atención: error en número de cédula, verifique matrícula: " + Trim(str(data_cli.Recordset("cl_codigo"))), vbCritical, "Mensaje"
                                End
                             End If
                             If IsNull(data_cli.Recordset("ci_tarj")) = False Then
                                Xnomced = Trim(str(Int(data_cli.Recordset("ci_tarj"))))
                             Else
                                Xnomced = "0"
                             End If
                             If IsNull(data_cli.Recordset("codcitarj")) = False Then
                                Xnomced = Xnomced + Trim(str(data_cli.Recordset("codcitarj")))
                             Else
                                Xnomced = Xnomced + "0"
                             End If
                             If Len(Trim(Xnomced)) = 9 Then
                                Xlin = Xlin + Trim(Xnomced)
                             End If
                             If Len(Trim(Xnomced)) = 8 Then
                                Xlin = Xlin + "0" + Trim(Xnomced)
                             End If
                             If Len(Trim(Xnomced)) = 7 Then
                                Xlin = Xlin + "00" + Trim(Xnomced)
                             End If
                             If Len(Trim(Xnomced)) = 6 Then
                                Xlin = Xlin + "000" + Trim(Xnomced)
                             End If
                             If Len(Trim(Xnomced)) = 5 Then
                                Xlin = Xlin + "0000" + Trim(Xnomced)
                             End If
        '                     data_emi.Recordset.FindFirst "cliente =" & data_cli.Recordset("cl_codigo")
        '                     data_emi.RecordSource = "select * from " & XNombre & " where cliente =" & data_cli.Recordset("cl_codigo")
                             XImp = data_emi.Recordset("total")
                             If Len(Trim(str(Int(XImp)))) = 1 Then
                                Xnomimp = "00000000" + Trim(str(Int(XImp))) + "00"
                             End If
                             If Len(Trim(str(Int(XImp)))) = 2 Then
                                Xnomimp = "0000000" + Trim(str(Int(XImp))) + "00"
                             End If
                             If Len(Trim(str(Int(XImp)))) = 3 Then
                                Xnomimp = "000000" + Trim(str(Int(XImp))) + "00"
                             End If
                             If Len(Trim(str(Int(XImp)))) = 4 Then
                                Xnomimp = "00000" + Trim(str(Int(XImp))) + "00"
                             End If
                             If Len(Trim(str(Int(XImp)))) = 5 Then
                                Xnomimp = "0000" + Trim(str(Int(XImp))) + "00"
                             End If
                             If Len(Trim(str(Int(XImp)))) = 6 Then
                                Xnomimp = "000" + Trim(str(Int(XImp))) + "00"
                             End If
                             Xlin = Xlin + Trim(Xnomimp) + Trim(Xmescabal)
                             Xlin = Xlin + Trim(Xnommescab)
                             Xnrorecstr = ""
                                If Len(Trim(str(Int(data_emi.Recordset("documento"))))) = 1 Then
                                   Xnrorecstr = "00000000000" & Trim(str(Int(data_emi.Recordset("documento"))))
                                End If
                                If Len(Trim(str(Int(data_emi.Recordset("documento"))))) = 2 Then
                                   Xnrorecstr = "0000000000" & Trim(str(Int(data_emi.Recordset("documento"))))
                                End If
                                If Len(Trim(str(Int(data_emi.Recordset("documento"))))) = 3 Then
                                   Xnrorecstr = "000000000" & Trim(str(Int(data_emi.Recordset("documento"))))
                                End If
                                If Len(Trim(str(Int(data_emi.Recordset("documento"))))) = 4 Then
                                   Xnrorecstr = "00000000" & Trim(str(Int(data_emi.Recordset("documento"))))
                                End If
                                If Len(Trim(str(Int(data_emi.Recordset("documento"))))) = 5 Then
                                   Xnrorecstr = "0000000" & Trim(str(Int(data_emi.Recordset("documento"))))
                                End If
                                If Len(Trim(str(Int(data_emi.Recordset("documento"))))) = 6 Then
                                   Xnrorecstr = "000000" & Trim(str(Int(data_emi.Recordset("documento"))))
                                End If
                                If Len(Trim(str(Int(data_emi.Recordset("documento"))))) = 7 Then
                                   Xnrorecstr = "00000" & Trim(str(Int(data_emi.Recordset("documento"))))
                                End If
                                If Len(Trim(str(Int(data_emi.Recordset("documento"))))) = 8 Then
                                   Xnrorecstr = "0000" & Trim(str(Int(data_emi.Recordset("documento"))))
                                End If
                                If Xnrorecstr = "" Then
                                   Xnrorecstr = "000000000000"
                                End If
                             Xlin = Xlin + Trim(Xnrorecstr) + "1" & "000000000000000"
        '                        Ximpcabal = data_emi.Recordset("importe") + data_emi.Recordset("deudas")
        '                        Ximpgrav = data_emi.Recordset("importe") + data_emi.Recordset("deudas")
                             Ximpcabal = XImp
                             Ximpgrav = XImp
                             Ximpgrav = Ximpgrav / 1.1
                             Ximpgrav = Ximpgrav * 0.1
                             Ximpivag = Ximpgrav
                             Xivacabal = Ximpcabal - Ximpgrav
                             Ximpgrav = Xivacabal
                             Ximpivag = Ximpivag * 0.01
                             Ximpgrastr = Format(Ximpgrav, "Standard")
                             Ximpivag = 0
                             Ximpivagstr = Format(Ximpivag, "Standard")
                             If Len(Trim(Ximpgrastr)) = 5 Then
                                Ximpgrastr2 = "00000000000" & Mid(Trim(Ximpgrastr), 1, 2)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 4, 2)
                             End If
                             If Len(Trim(Ximpgrastr)) = 6 Then
                                Ximpgrastr2 = "0000000000" & Mid(Trim(Ximpgrastr), 1, 3)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 5, 2)
                             End If
                             If Len(Trim(Ximpgrastr)) = 8 Then
                                Ximpgrastr2 = "000000000" & Mid(Trim(Ximpgrastr), 1, 1)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 3, 3)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 7, 2)
                             End If
                             If Len(Trim(Ximpgrastr)) = 9 Then
                                Ximpgrastr2 = "00000000" & Mid(Trim(Ximpgrastr), 1, 2)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 4, 3)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 8, 2)
                             End If
                             If Len(Trim(Ximpgrastr)) < 5 Then
                                Ximpgrastr2 = "000000000000000"
                             End If
                             Xlin = Xlin + Trim(Ximpgrastr2)
                             If Len(Trim(Ximpivagstr)) = 3 Then
                                Ximpgrastr2 = "0000000000000" & Mid(Trim(Ximpivagstr), 2, 2)
                             End If
                             If Len(Trim(Ximpivagstr)) = 4 Then
                                Ximpgrastr2 = "000000000000" & Mid(Trim(Ximpivagstr), 1, 1)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpivagstr), 3, 2)
                             End If
                             If Len(Trim(Ximpivagstr)) = 5 Then
                                Ximpgrastr2 = "00000000000" & Mid(Trim(Ximpivagstr), 1, 2)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpivagstr), 4, 2)
                             End If
                             If Len(Trim(Ximpivagstr)) = 6 Then
                                Ximpgrastr2 = "0000000000" & Mid(Trim(Ximpivagstr), 1, 3)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpivagstr), 5, 2)
                             End If
                             If Len(Trim(Ximpivagstr)) = 8 Then
                                Ximpgrastr2 = "000000000" & Mid(Trim(Ximpivagstr), 1, 1)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpivagstr), 3, 3)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpivagstr), 7, 2)
                             End If
                             If Ximpivagstr = "" Then
                                Ximpgrastr2 = "000000000000000"
                             End If
                             Xlin = Xlin + Trim(Ximpgrastr2)
                             
                             Xlin = Xlin + "        "
                             Print #1, Xlin
                             data_inf.Recordset.AddNew
                             data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                             data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                             data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                             data_inf.Recordset("cl_hon_pes") = XImp
                             data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                             data_inf.Recordset("cl_nrotarj") = data_cli.Recordset("cl_nrotarj")
                             data_inf.Recordset.Update
                             data_emi.Recordset.MoveNext
                        Loop
                  
                        data_mutuales.RecordSource = "select * from mutuales where socio =" & data_cli.Recordset("cl_codigo") & " and importe_deuda >=" & 0
                        data_mutuales.Refresh
                        If data_mutuales.Recordset.RecordCount > 0 Then
                           data_mutuales.Recordset.MoveFirst
                           Do While Not data_mutuales.Recordset.EOF
                           
                             Xlin = "CBCU28443915006N"
                             Xtarj = Trim(data_cli.Recordset("cl_nrotarj"))
                             Xlin = Xlin + Trim(Xtarj)
                             Call controlced(data_cli.Recordset("ci_tarj"), data_cli.Recordset("codcitarj"))
                             If Int(Val(t_codced.Text)) = Int(data_cli.Recordset("codcitarj")) Then
                             Else
                                MsgBox "Atención: error en número de cédula, verifique matrícula: " + Trim(str(data_cli.Recordset("cl_codigo"))), vbCritical, "Mensaje"
                                End
                             End If
                             If IsNull(data_cli.Recordset("ci_tarj")) = False Then
                                Xnomced = Trim(str(Int(data_cli.Recordset("ci_tarj"))))
                             Else
                                Xnomced = "0"
                             End If
                             If IsNull(data_cli.Recordset("codcitarj")) = False Then
                                Xnomced = Xnomced + Trim(str(data_cli.Recordset("codcitarj")))
                             Else
                                Xnomced = Xnomced + "0"
                             End If
                             If Len(Trim(Xnomced)) = 9 Then
                                Xlin = Xlin + Trim(Xnomced)
                             End If
                             If Len(Trim(Xnomced)) = 8 Then
                                Xlin = Xlin + "0" + Trim(Xnomced)
                             End If
                             If Len(Trim(Xnomced)) = 7 Then
                                Xlin = Xlin + "00" + Trim(Xnomced)
                             End If
                             If Len(Trim(Xnomced)) = 6 Then
                                Xlin = Xlin + "000" + Trim(Xnomced)
                             End If
                             If Len(Trim(Xnomced)) = 5 Then
                                Xlin = Xlin + "0000" + Trim(Xnomced)
                             End If
        '                     data_emi.Recordset.FindFirst "cliente =" & data_cli.Recordset("cl_codigo")
        '                     data_emi.RecordSource = "select * from " & XNombre & " where cliente =" & data_cli.Recordset("cl_codigo")
                             XImp = data_mutuales.Recordset("importe_deuda")
                             If Len(Trim(str(Int(XImp)))) = 1 Then
                                Xnomimp = "00000000" + Trim(str(Int(XImp))) + "00"
                             End If
                             If Len(Trim(str(Int(XImp)))) = 2 Then
                                Xnomimp = "0000000" + Trim(str(Int(XImp))) + "00"
                             End If
                             If Len(Trim(str(Int(XImp)))) = 3 Then
                                Xnomimp = "000000" + Trim(str(Int(XImp))) + "00"
                             End If
                             If Len(Trim(str(Int(XImp)))) = 4 Then
                                Xnomimp = "00000" + Trim(str(Int(XImp))) + "00"
                             End If
                             If Len(Trim(str(Int(XImp)))) = 5 Then
                                Xnomimp = "0000" + Trim(str(Int(XImp))) + "00"
                             End If
                             If Len(Trim(str(Int(XImp)))) = 6 Then
                                Xnomimp = "000" + Trim(str(Int(XImp))) + "00"
                             End If
                             Xlin = Xlin + Trim(Xnomimp) + Trim(Xmescabal)
                             Xlin = Xlin + Trim(Xnommescab)
                             Xnrorecstr = ""
                                If Len(Trim(str(Int(data_mutuales.Recordset("recibo"))))) = 1 Then
                                   Xnrorecstr = "00000000000" & Trim(str(Int(data_mutuales.Recordset("recibo"))))
                                End If
                                If Len(Trim(str(Int(data_mutuales.Recordset("recibo"))))) = 2 Then
                                   Xnrorecstr = "0000000000" & Trim(str(Int(data_mutuales.Recordset("recibo"))))
                                End If
                                If Len(Trim(str(Int(data_mutuales.Recordset("recibo"))))) = 3 Then
                                   Xnrorecstr = "000000000" & Trim(str(Int(data_mutuales.Recordset("recibo"))))
                                End If
                                If Len(Trim(str(Int(data_mutuales.Recordset("recibo"))))) = 4 Then
                                   Xnrorecstr = "00000000" & Trim(str(Int(data_mutuales.Recordset("recibo"))))
                                End If
                                If Len(Trim(str(Int(data_mutuales.Recordset("recibo"))))) = 5 Then
                                   Xnrorecstr = "0000000" & Trim(str(Int(data_mutuales.Recordset("recibo"))))
                                End If
                                If Len(Trim(str(Int(data_mutuales.Recordset("recibo"))))) = 6 Then
                                   Xnrorecstr = "000000" & Trim(str(Int(data_mutuales.Recordset("recibo"))))
                                End If
                                If Len(Trim(str(Int(data_mutuales.Recordset("recibo"))))) = 7 Then
                                   Xnrorecstr = "00000" & Trim(str(Int(data_mutuales.Recordset("recibo"))))
                                End If
                                If Len(Trim(str(Int(data_mutuales.Recordset("recibo"))))) = 8 Then
                                   Xnrorecstr = "0000" & Trim(str(Int(data_mutuales.Recordset("recibo"))))
                                End If
                                If Xnrorecstr = "" Then
                                   Xnrorecstr = "000000000000"
                                End If
                             Xlin = Xlin + Trim(Xnrorecstr) + "1" & "000000000000000"
        '                        Ximpcabal = data_emi.Recordset("importe") + data_emi.Recordset("deudas")
        '                        Ximpgrav = data_emi.Recordset("importe") + data_emi.Recordset("deudas")
                             Ximpcabal = XImp
                             Ximpgrav = XImp
                             Ximpgrav = Ximpgrav / 1.1
                             Ximpgrav = Ximpgrav * 0.1
                             Ximpivag = Ximpgrav
                             Xivacabal = Ximpcabal - Ximpgrav
                             Ximpgrav = Xivacabal
                             Ximpivag = Ximpivag * 0.01
                             Ximpgrastr = Format(Ximpgrav, "Standard")
                             Ximpivag = 0
                             Ximpivagstr = Format(Ximpivag, "Standard")
                             If Len(Trim(Ximpgrastr)) = 5 Then
                                Ximpgrastr2 = "00000000000" & Mid(Trim(Ximpgrastr), 1, 2)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 4, 2)
                             End If
                             If Len(Trim(Ximpgrastr)) = 6 Then
                                Ximpgrastr2 = "0000000000" & Mid(Trim(Ximpgrastr), 1, 3)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 5, 2)
                             End If
                             If Len(Trim(Ximpgrastr)) = 8 Then
                                Ximpgrastr2 = "000000000" & Mid(Trim(Ximpgrastr), 1, 1)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 3, 3)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 7, 2)
                             End If
                             If Len(Trim(Ximpgrastr)) = 9 Then
                                Ximpgrastr2 = "00000000" & Mid(Trim(Ximpgrastr), 1, 2)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 4, 3)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 8, 2)
                             End If
                             If Len(Trim(Ximpgrastr)) < 5 Then
                                Ximpgrastr2 = "000000000000000"
                             End If
                             Xlin = Xlin + Trim(Ximpgrastr2)
                             If Len(Trim(Ximpivagstr)) = 3 Then
                                Ximpgrastr2 = "0000000000000" & Mid(Trim(Ximpivagstr), 2, 2)
                             End If
                             If Len(Trim(Ximpivagstr)) = 4 Then
                                Ximpgrastr2 = "000000000000" & Mid(Trim(Ximpivagstr), 1, 1)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpivagstr), 3, 2)
                             End If
                             If Len(Trim(Ximpivagstr)) = 5 Then
                                Ximpgrastr2 = "00000000000" & Mid(Trim(Ximpivagstr), 1, 2)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpivagstr), 4, 2)
                             End If
                             If Len(Trim(Ximpivagstr)) = 6 Then
                                Ximpgrastr2 = "0000000000" & Mid(Trim(Ximpivagstr), 1, 3)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpivagstr), 5, 2)
                             End If
                             If Len(Trim(Ximpivagstr)) = 8 Then
                                Ximpgrastr2 = "000000000" & Mid(Trim(Ximpivagstr), 1, 1)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpivagstr), 3, 3)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpivagstr), 7, 2)
                             End If
                             If Ximpivagstr = "" Then
                                Ximpgrastr2 = "000000000000000"
                             End If
                             Xlin = Xlin + Trim(Ximpgrastr2)
                             
                             Xlin = Xlin + "        "
                             Print #1, Xlin
                             data_inf.Recordset.AddNew
                             data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                             data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                             data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                             data_inf.Recordset("cl_hon_pes") = XImp
                             data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                             data_inf.Recordset("cl_nrotarj") = data_cli.Recordset("cl_nrotarj")
                             data_inf.Recordset.Update
                             data_mutuales.Recordset.MoveNext
                           Loop
                        
                        End If
                  
                     Else
                        'solo mutual
                        data_mutuales.RecordSource = "select * from mutuales where socio =" & data_cli.Recordset("cl_codigo") & " and importe_deuda >=" & 0
                        data_mutuales.Refresh
                        If data_mutuales.Recordset.RecordCount > 0 Then
                           data_mutuales.Recordset.MoveFirst
                           Do While Not data_mutuales.Recordset.EOF
                           
                             Xlin = "CBCU28443915006N"
                             Xtarj = Trim(data_cli.Recordset("cl_nrotarj"))
                             Xlin = Xlin + Trim(Xtarj)
                             Call controlced(data_cli.Recordset("ci_tarj"), data_cli.Recordset("codcitarj"))
                             If Int(Val(t_codced.Text)) = Int(data_cli.Recordset("codcitarj")) Then
                             Else
                                MsgBox "Atención: error en número de cédula, verifique matrícula: " + Trim(str(data_cli.Recordset("cl_codigo"))), vbCritical, "Mensaje"
                                End
                             End If
                             If IsNull(data_cli.Recordset("ci_tarj")) = False Then
                                Xnomced = Trim(str(Int(data_cli.Recordset("ci_tarj"))))
                             Else
                                Xnomced = "0"
                             End If
                             If IsNull(data_cli.Recordset("codcitarj")) = False Then
                                Xnomced = Xnomced + Trim(str(data_cli.Recordset("codcitarj")))
                             Else
                                Xnomced = Xnomced + "0"
                             End If
                             If Len(Trim(Xnomced)) = 9 Then
                                Xlin = Xlin + Trim(Xnomced)
                             End If
                             If Len(Trim(Xnomced)) = 8 Then
                                Xlin = Xlin + "0" + Trim(Xnomced)
                             End If
                             If Len(Trim(Xnomced)) = 7 Then
                                Xlin = Xlin + "00" + Trim(Xnomced)
                             End If
                             If Len(Trim(Xnomced)) = 6 Then
                                Xlin = Xlin + "000" + Trim(Xnomced)
                             End If
                             If Len(Trim(Xnomced)) = 5 Then
                                Xlin = Xlin + "0000" + Trim(Xnomced)
                             End If
        '                     data_emi.Recordset.FindFirst "cliente =" & data_cli.Recordset("cl_codigo")
        '                     data_emi.RecordSource = "select * from " & XNombre & " where cliente =" & data_cli.Recordset("cl_codigo")
                             XImp = data_mutuales.Recordset("importe_deuda")
                             If Len(Trim(str(Int(XImp)))) = 1 Then
                                Xnomimp = "00000000" + Trim(str(Int(XImp))) + "00"
                             End If
                             If Len(Trim(str(Int(XImp)))) = 2 Then
                                Xnomimp = "0000000" + Trim(str(Int(XImp))) + "00"
                             End If
                             If Len(Trim(str(Int(XImp)))) = 3 Then
                                Xnomimp = "000000" + Trim(str(Int(XImp))) + "00"
                             End If
                             If Len(Trim(str(Int(XImp)))) = 4 Then
                                Xnomimp = "00000" + Trim(str(Int(XImp))) + "00"
                             End If
                             If Len(Trim(str(Int(XImp)))) = 5 Then
                                Xnomimp = "0000" + Trim(str(Int(XImp))) + "00"
                             End If
                             If Len(Trim(str(Int(XImp)))) = 6 Then
                                Xnomimp = "000" + Trim(str(Int(XImp))) + "00"
                             End If
                             Xlin = Xlin + Trim(Xnomimp) + Trim(Xmescabal)
                             Xlin = Xlin + Trim(Xnommescab)
                             Xnrorecstr = ""
                                If Len(Trim(str(Int(data_mutuales.Recordset("recibo"))))) = 1 Then
                                   Xnrorecstr = "00000000000" & Trim(str(Int(data_mutuales.Recordset("recibo"))))
                                End If
                                If Len(Trim(str(Int(data_mutuales.Recordset("recibo"))))) = 2 Then
                                   Xnrorecstr = "0000000000" & Trim(str(Int(data_mutuales.Recordset("recibo"))))
                                End If
                                If Len(Trim(str(Int(data_mutuales.Recordset("recibo"))))) = 3 Then
                                   Xnrorecstr = "000000000" & Trim(str(Int(data_mutuales.Recordset("recibo"))))
                                End If
                                If Len(Trim(str(Int(data_mutuales.Recordset("recibo"))))) = 4 Then
                                   Xnrorecstr = "00000000" & Trim(str(Int(data_mutuales.Recordset("recibo"))))
                                End If
                                If Len(Trim(str(Int(data_mutuales.Recordset("recibo"))))) = 5 Then
                                   Xnrorecstr = "0000000" & Trim(str(Int(data_mutuales.Recordset("recibo"))))
                                End If
                                If Len(Trim(str(Int(data_mutuales.Recordset("recibo"))))) = 6 Then
                                   Xnrorecstr = "000000" & Trim(str(Int(data_mutuales.Recordset("recibo"))))
                                End If
                                If Len(Trim(str(Int(data_mutuales.Recordset("recibo"))))) = 7 Then
                                   Xnrorecstr = "00000" & Trim(str(Int(data_mutuales.Recordset("recibo"))))
                                End If
                                If Len(Trim(str(Int(data_mutuales.Recordset("recibo"))))) = 8 Then
                                   Xnrorecstr = "0000" & Trim(str(Int(data_mutuales.Recordset("recibo"))))
                                End If
                                If Xnrorecstr = "" Then
                                   Xnrorecstr = "000000000000"
                                End If
                             Xlin = Xlin + Trim(Xnrorecstr) + "1" & "000000000000000"
        '                        Ximpcabal = data_emi.Recordset("importe") + data_emi.Recordset("deudas")
        '                        Ximpgrav = data_emi.Recordset("importe") + data_emi.Recordset("deudas")
                             Ximpcabal = XImp
                             Ximpgrav = XImp
                             Ximpgrav = Ximpgrav / 1.1
                             Ximpgrav = Ximpgrav * 0.1
                             Ximpivag = Ximpgrav
                             Xivacabal = Ximpcabal - Ximpgrav
                             Ximpgrav = Xivacabal
                             Ximpivag = Ximpivag * 0.01
                             Ximpgrastr = Format(Ximpgrav, "Standard")
                             Ximpivag = 0
                             Ximpivagstr = Format(Ximpivag, "Standard")
                             If Len(Trim(Ximpgrastr)) = 5 Then
                                Ximpgrastr2 = "00000000000" & Mid(Trim(Ximpgrastr), 1, 2)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 4, 2)
                             End If
                             If Len(Trim(Ximpgrastr)) = 6 Then
                                Ximpgrastr2 = "0000000000" & Mid(Trim(Ximpgrastr), 1, 3)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 5, 2)
                             End If
                             If Len(Trim(Ximpgrastr)) = 8 Then
                                Ximpgrastr2 = "000000000" & Mid(Trim(Ximpgrastr), 1, 1)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 3, 3)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 7, 2)
                             End If
                             If Len(Trim(Ximpgrastr)) = 9 Then
                                Ximpgrastr2 = "00000000" & Mid(Trim(Ximpgrastr), 1, 2)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 4, 3)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 8, 2)
                             End If
                             If Len(Trim(Ximpgrastr)) < 5 Then
                                Ximpgrastr2 = "000000000000000"
                             End If
                             Xlin = Xlin + Trim(Ximpgrastr2)
                             If Len(Trim(Ximpivagstr)) = 3 Then
                                Ximpgrastr2 = "0000000000000" & Mid(Trim(Ximpivagstr), 2, 2)
                             End If
                             If Len(Trim(Ximpivagstr)) = 4 Then
                                Ximpgrastr2 = "000000000000" & Mid(Trim(Ximpivagstr), 1, 1)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpivagstr), 3, 2)
                             End If
                             If Len(Trim(Ximpivagstr)) = 5 Then
                                Ximpgrastr2 = "00000000000" & Mid(Trim(Ximpivagstr), 1, 2)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpivagstr), 4, 2)
                             End If
                             If Len(Trim(Ximpivagstr)) = 6 Then
                                Ximpgrastr2 = "0000000000" & Mid(Trim(Ximpivagstr), 1, 3)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpivagstr), 5, 2)
                             End If
                             If Len(Trim(Ximpivagstr)) = 8 Then
                                Ximpgrastr2 = "000000000" & Mid(Trim(Ximpivagstr), 1, 1)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpivagstr), 3, 3)
                                Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpivagstr), 7, 2)
                             End If
                             If Ximpivagstr = "" Then
                                Ximpgrastr2 = "000000000000000"
                             End If
                             Xlin = Xlin + Trim(Ximpgrastr2)
                             
                             Xlin = Xlin + "        "
                             Print #1, Xlin
                             data_inf.Recordset.AddNew
                             data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                             data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                             data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                             data_inf.Recordset("cl_hon_pes") = XImp
                             data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                             data_inf.Recordset("cl_nrotarj") = data_cli.Recordset("cl_nrotarj")
                             data_inf.Recordset.Update
                             data_mutuales.Recordset.MoveNext
                           Loop
                        
                        End If
                     End If
                  End If
               Else
                  MsgBox "Verifique socio:" & data_cli.Recordset("cl_codigo") & " NO SE INCLUYE"
               End If
            End If
         End If
         data_cli.Recordset.MoveNext
      Loop
      Close #1
      data_inf.RecordSource = "select * from infcli"
      data_inf.Refresh
      Xtitulo = "DEBITOS CABAL CORRESPONDIENTES MES : " & txt_mes.Text & "/" & txt_ano.Text
      cr1.ReportFileName = App.path & "\infdebitos.rpt"
      cr1.ReportTitle = Xtitulo
      cr1.Action = 1
      MsgBox "Proceso terminado", vbInformation, "Mensaje"
   End If
End If
If cbotarj.ListIndex = 3 Then 'brou ---ok
   Proceso_brou
   data_inf.RecordSource = "select * from infcli"
   data_inf.Refresh
   Xtitulo = "DEBITOS BROU CORRESPONDIENTES MES : " & txt_mes.Text & "/" & txt_ano.Text
   cr1.ReportFileName = App.path & "\infdebitos.rpt"
   cr1.ReportTitle = Xtitulo
   cr1.Action = 1
   MsgBox "Proceso terminado!!", vbInformation, "Mensaje"

End If
If cbotarj.ListIndex = 4 Then 'oca
   Proceso_oca
End If

If cbotarj.ListIndex = 5 Then 'redpagos ---okkk
   Proceso_redpagos
End If

frm_debitos.MousePointer = 0
b_proc.Enabled = True

End Sub

Private Sub b_sal_Click()
Unload Me


End Sub

Private Sub cbotarj_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_mes.SetFocus
End If

End Sub


Private Sub Form_Load()
mfec.Text = Format(Date, "dd/mm/yyyy")
txt_mes.Text = Month(Date)
txt_ano.Text = Year(Date)
Data1.ConnectionString = "dsn=" & Xconexrmt
data_cli.ConnectionString = "dsn=" & Xconexrmt
data_emi.ConnectionString = "dsn=" & Xconexrmt
data_arq.ConnectionString = "dsn=" & Xconexrmt
data_mutuales.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_deudas.Connect = "odbc;dsn=" & Xconexrmt & ";"
Adodc1.ConnectionString = "dsn=" & Xconexrmt
Adodc2.ConnectionString = "dsn=" & Xconexrmt

'data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_inf.DatabaseName = App.path & "\informes.mdb"
'data_inf.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"
data_inf.RecordSource = "infcli"
data_inf.Refresh
'data_emi.DatabaseName = ""
'data_emi.Connect = "odbc;dsn=" & Xconexrmt & ";"

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub mfec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_proc.SetFocus
End If

End Sub

Private Sub txt_ano_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfec.SetFocus
End If

End Sub

Private Sub txt_mes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_ano.SetFocus
End If

End Sub

Public Function controlced(ByVal Nced As Long, ByVal Ncced As Integer) As Integer
Dim Xpond, Xn1, Xn2, Xn3, Xn4, Xn5, Xn6, Xn7, Xtot As Long
Dim Xcedtex, Xtottex As String
Dim Xced1, Xced2, Xced3, Xced4, Xced5, Xced6, Xced7, Xlargo As Long
Xn1 = 2
Xn2 = 9
Xn3 = 8
Xn4 = 7
Xn5 = 6
Xn6 = 3
Xn7 = 4
Xpond = 10
'If IsNumeric(txt_ced2.Text) = False Then
'   txt_ced2.Text = 0
'Else
   Xcedtex = Trim(str(Nced))
   Xlargo = Len(Xcedtex)
   If Xlargo = 6 Then
      Xcedtex = "0" & Trim(Xcedtex)
   End If
   Xced1 = Val(Mid(Trim(Xcedtex), 1, 1))
   Xced2 = Val(Mid(Xcedtex, 2, 1))
   Xced3 = Val(Mid(Xcedtex, 3, 1))
   Xced4 = Val(Mid(Xcedtex, 4, 1))
   Xced5 = Val(Mid(Xcedtex, 5, 1))
   Xced6 = Val(Mid(Xcedtex, 6, 1))
   Xced7 = Val(Mid(Xcedtex, 7, 1))
   Xced1 = Xced1 * Xn1
   Xced2 = Xced2 * Xn2
   Xced3 = Xced3 * Xn3
   Xced4 = Xced4 * Xn4
   Xced5 = Xced5 * Xn5
   Xced6 = Xced6 * Xn6
   Xced7 = Xced7 * Xn7
   Xtot = Xced1 + Xced2 + Xced3 + Xced4 + Xced5 + Xced6 + Xced7
   If Len(Trim(str(Xtot))) = 1 Then
      Xtottex = "0000" & Trim(str(Xtot))
   End If
   If Len(Trim(str(Xtot))) = 2 Then
      Xtottex = "000" & Trim(str(Xtot))
   End If
   If Len(Trim(str(Xtot))) = 3 Then
      Xtottex = "00" & Trim(str(Xtot))
   End If
   If Len(Trim(str(Xtot))) = 4 Then
      Xtottex = "0" & Trim(str(Xtot))
   End If
   Xtot = Val(Mid(Xtottex, 5, 1))
   If Xtot <> 0 Then
      Xtot = Xpond - Xtot
   Else
      Xtot = 0
   End If
   If Xtot = Ncced Then
      t_codced.Text = Ncced
   Else
      t_codced.Text = 99
   End If
'      If txt_ced.Text <> 0 Then


End Function

Function EsCCValido(sTarjeta As String) As Boolean
Dim iPeso As Integer
Dim iDigito As Integer
Dim iSuma As Integer
Dim iContador As Integer
Dim sNuevaTarjeta As String
Dim cCaracter As String * 1

    iPeso = 0
    iDigito = 0
    iSuma = 0

    'Reemplazar cualquier no digito por una cadena vacía
    For iContador = 1 To Len(sTarjeta)
        cCaracter = Mid(sTarjeta, iContador, 1)
        If IsNumeric(cCaracter) Then
            sNuevaTarjeta = sNuevaTarjeta & cCaracter
        End If
    Next iContador

    ' Si es 0 devolver Falso
    If sNuevaTarjeta = 0 Then
        EsCCValido = False
        Exit Function
    End If

    ' Si el número de dígitos es par el primer peso es 2, de lo
    ' contrario es 1
    If (Len(sNuevaTarjeta) Mod 2) = 0 Then
        iPeso = 2
    Else
        iPeso = 1
    End If

    For iContador = 1 To Len(sNuevaTarjeta)
        iDigito = Mid(sNuevaTarjeta, iContador, 1) * iPeso
        If iDigito > 9 Then
           iDigito = iDigito - 9
        End If
        iSuma = iSuma + iDigito
        
        ' Cambiar peso para el siguiente dígito
        If iPeso = 2 Then
            iPeso = 1
        Else
            iPeso = 2
        End If
    Next iContador

    ' Devolver verdadero si la suma es divisible por 10
    If (iSuma Mod 10) = 0 Then
        EsCCValido = True
    Else
        EsCCValido = False
    End If

End Function

Public Sub Proceso_visa()
Dim Xlin, Xtarj, Xfec, XCantex, Xnommes, Xnomimp, Xnomced, Xmesmast, Xmescabal, Xnommescab As String
Dim XP, Xcant As Integer
Dim XImp, Xtot, Xtardelf, Ximparqueo As Double
Dim Xtitulo As String
Dim XNombre, Xnomarq As String
Dim Xventar As String
Dim Xmatstr, Xnrorecstr, Ximpgrastr, Ximpgrastr2 As String
Dim Ximpgrav As Double
Dim Xelcoddevuelto As Integer


Xcant = 0

Dim Xmes, Xano As Integer

XNombre = "emi"
Xnomarq = "arq"
If txt_mes.Text > 9 Then
   XNombre = XNombre + Trim(txt_mes.Text) + Mid(Trim(txt_ano.Text), 3, 2)
   Xmes = txt_mes.Text - 1
   Xano = txt_ano.Text
   Xnomarq = Xnomarq + Trim(str(Xmes)) + Mid(Trim(str(Xano)), 3, 2)
Else
   XNombre = XNombre + "0" + Trim(txt_mes.Text) + Mid(Trim(txt_ano.Text), 3, 2)
   If txt_mes.Text = 1 Then
      Xmes = 12
      Xano = txt_ano.Text - 1
   Else
      Xmes = txt_mes.Text - 1
      Xano = txt_ano.Text
   End If
   If Xmes > 9 Then
      Xnomarq = Xnomarq + Trim(str(Xmes)) + Mid(Trim(str(Xano)), 3, 2)
   Else
      Xnomarq = Xnomarq + "0" + Trim(str(Xmes)) + Mid(Trim(str(Xano)), 3, 2)
   End If
End If

If mfec.Text <> "__/__/____" Then
   Xfec = Trim(Mid(mfec.Text, 7, 4))
   Xfec = Xfec + Trim(Mid(mfec.Text, 4, 2))
   Xfec = Xfec + Trim(Mid(mfec.Text, 1, 2))
   Xmesmast = Trim(Mid(mfec.Text, 4, 2))
   Xmesmast = Xmesmast + "/" + Trim(Mid(mfec.Text, 9, 2))
   Xnommescab = Trim(Mid(mfec.Text, 4, 2))
   Xnommescab = Xnommescab + Trim(Mid(mfec.Text, 9, 2))
   Xmescabal = Trim(Mid(mfec.Text, 1, 2))
   Xmescabal = Xmescabal + Trim(Mid(mfec.Text, 4, 2))
   Xmescabal = Xmescabal + Trim(Mid(mfec.Text, 9, 2))
End If
Dim Ximpivag As Double
Dim Ximpivagstr As String

If Month(mfec.Text) = 1 Then
   Xnommes = "ENE/" + Trim(Mid(mfec.Text, 9, 2))
End If
If Month(mfec.Text) = 2 Then
   Xnommes = "FEB/" + Trim(Mid(mfec.Text, 9, 2))
End If
If Month(mfec.Text) = 3 Then
   Xnommes = "MAR/" + Trim(Mid(mfec.Text, 9, 2))
End If
If Month(mfec.Text) = 4 Then
   Xnommes = "ABR/" + Trim(Mid(mfec.Text, 9, 2))
End If
If Month(mfec.Text) = 5 Then
   Xnommes = "MAY/" + Trim(Mid(mfec.Text, 9, 2))
End If
If Month(mfec.Text) = 6 Then
   Xnommes = "JUN/" + Trim(Mid(mfec.Text, 9, 2))
End If
If Month(mfec.Text) = 7 Then
   Xnommes = "JUL/" + Trim(Mid(mfec.Text, 9, 2))
End If
If Month(mfec.Text) = 8 Then
   Xnommes = "AGO/" + Trim(Mid(mfec.Text, 9, 2))
End If
If Month(mfec.Text) = 9 Then
   Xnommes = "SET/" + Trim(Mid(mfec.Text, 9, 2))
End If
If Month(mfec.Text) = 10 Then
   Xnommes = "OCT/" + Trim(Mid(mfec.Text, 9, 2))
End If
If Month(mfec.Text) = 11 Then
   Xnommes = "NOV/" + Trim(Mid(mfec.Text, 9, 2))
End If
If Month(mfec.Text) = 12 Then
   Xnommes = "DIC/" + Trim(Mid(mfec.Text, 9, 2))
End If
Dim control As String

data_cli.RecordSource = "Select * from clientes where cl_nrocobr =" & 514 & " and estado in (1) and fecha_baja is null"
data_cli.Refresh
If data_cli.Recordset.RecordCount > 0 Then
   data_cli.Recordset.MoveFirst
   Open "c:\debitos\Visanet1.txt" For Output As #1
   Xtot = 0
   Do While Not data_cli.Recordset.EOF
      Xmatstr = ""
      If Len(Trim(data_cli.Recordset("cl_nrotarj"))) = 16 Then
         data_emi.RecordSource = "Select * from deudas where cliente =" & data_cli.Recordset("cl_codigo") & " and fecha_pago is null and nro_cobr in (514) and fecha >='" & Format("01/08/2021", "yyyy-mm-dd") & "' and mes >" & 0
         data_emi.Refresh
         If data_emi.Recordset.RecordCount > 0 Then
            data_emi.Recordset.MoveFirst
            Do While Not data_emi.Recordset.EOF
               Xtarj = Trim(data_cli.Recordset("cl_nrotarj"))
               Xcant = Xcant + 1
               If Len(Trim(str(Xcant))) = 1 Then
                  XCantex = "000000" + Trim(str(Xcant))
               End If
               If Len(Trim(str(Xcant))) = 2 Then
                  XCantex = "00000" + Trim(str(Xcant))
               End If
               If Len(Trim(str(Xcant))) = 3 Then
                  XCantex = "0000" + Trim(str(Xcant))
               End If
               If Len(Trim(str(Xcant))) = 4 Then
                  XCantex = "000" + Trim(str(Xcant))
               End If
               If Len(Trim(str(Int(data_cli.Recordset("cl_codigo"))))) = 2 Then
                  Xmatstr = "00000000" & Trim(str(Int(data_cli.Recordset("cl_codigo"))))
               End If
               If Len(Trim(str(Int(data_cli.Recordset("cl_codigo"))))) = 3 Then
                  Xmatstr = "0000000" & Trim(str(Int(data_cli.Recordset("cl_codigo"))))
               End If
               If Len(Trim(str(Int(data_cli.Recordset("cl_codigo"))))) = 4 Then
                  Xmatstr = "000000" & Trim(str(Int(data_cli.Recordset("cl_codigo"))))
               End If
               If Len(Trim(str(Int(data_cli.Recordset("cl_codigo"))))) = 5 Then
                  Xmatstr = "00000" & Trim(str(Int(data_cli.Recordset("cl_codigo"))))
               End If
               If Len(Trim(str(Int(data_cli.Recordset("cl_codigo"))))) = 6 Then
                  Xmatstr = "0000" & Trim(str(Int(data_cli.Recordset("cl_codigo"))))
               End If
               If Len(Trim(str(Int(data_cli.Recordset("cl_codigo"))))) = 7 Then
                  Xmatstr = "000" & Trim(str(Int(data_cli.Recordset("cl_codigo"))))
               End If
               If Len(Trim(str(Int(data_cli.Recordset("cl_codigo"))))) = 8 Then
                  Xmatstr = "00" & Trim(str(Int(data_cli.Recordset("cl_codigo"))))
               End If
               If Len(Trim(str(Int(data_cli.Recordset("cl_codigo"))))) = 9 Then
                  Xmatstr = "0" & Trim(str(Int(data_cli.Recordset("cl_codigo"))))
               End If
               If Len(Trim(str(Int(data_cli.Recordset("cl_codigo"))))) = 10 Then
                  Xmatstr = Trim(str(Int(data_cli.Recordset("cl_codigo"))))
               End If
               If Xmatstr = "" Then
                  Xmatstr = "0000000000"
               End If
               XImp = Int(data_emi.Recordset("total"))
               If Len(Trim(str(Int(XImp)))) = 1 Then
                  Xnomimp = "000000000000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 2 Then
                  Xnomimp = "00000000000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 3 Then
                  Xnomimp = "0000000000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 4 Then
                  Xnomimp = "000000000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 5 Then
                  Xnomimp = "00000000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 6 Then
                  Xnomimp = "0000000" + Trim(str(Int(XImp))) + "00"
               End If
               Xtot = Xtot + XImp
               If IsNull(data_cli.Recordset("cl_tj_venc")) = False Then
                  Xventar = Trim(Mid(CStr(data_cli.Recordset("cl_tj_venc")), 4, 2))
                  Xventar = Trim(Xventar) & Trim(Mid(CStr(data_cli.Recordset("cl_tj_venc")), 9, 2))
               Else
                  Xventar = "1201"
               End If
               Xlin = "10000001" + Trim(XCantex) + "02001024" + "001" + Trim(Xfec) + "085801009005000" + Trim(Xtarj) + Trim(Xventar) + Trim(Xfec) + Trim(Xnomimp) + "000000" + Trim(Xnommes) + Trim(Xmatstr) & "6" & "                       "
               Xnrorecstr = ""
               If Len(Trim(str(Int(data_emi.Recordset("documento"))))) = 1 Then
                  Xnrorecstr = "000000" & Trim(str(Int(data_emi.Recordset("documento"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("documento"))))) = 2 Then
                  Xnrorecstr = "00000" & Trim(str(Int(data_emi.Recordset("documento"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("documento"))))) = 3 Then
                  Xnrorecstr = "0000" & Trim(str(Int(data_emi.Recordset("documento"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("documento"))))) = 4 Then
                  Xnrorecstr = "000" & Trim(str(Int(data_emi.Recordset("documento"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("documento"))))) = 5 Then
                  Xnrorecstr = "00" & Trim(str(Int(data_emi.Recordset("documento"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("documento"))))) = 6 Then
                  Xnrorecstr = "0" & Trim(str(Int(data_emi.Recordset("documento"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("documento"))))) = 7 Then
                  Xnrorecstr = Trim(str(Int(data_emi.Recordset("documento"))))
               End If
               If Xnrorecstr = "" Then
                  Xnrorecstr = "0000000"
               End If
'               Ximpgrav = data_emi.Recordset("importe") + data_emi.Recordset("deudas")
               Ximpgrav = data_emi.Recordset("total")
'               Ximpgrav = Ximpgrav + Ximparqueo
               Ximpgrav = Ximpgrav / 1.1
               Ximpgrastr = Format(Ximpgrav, "Standard")
               If Len(Trim(Ximpgrastr)) = 5 Then
                  Ximpgrastr2 = "00000000000" & Mid(Trim(Ximpgrastr), 1, 2)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 4, 2)
               End If
               If Len(Trim(Ximpgrastr)) = 6 Then
                  Ximpgrastr2 = "0000000000" & Mid(Trim(Ximpgrastr), 1, 3)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 5, 2)
               End If
               If Len(Trim(Ximpgrastr)) = 8 Then
                  Ximpgrastr2 = "000000000" & Mid(Trim(Ximpgrastr), 1, 1)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 3, 3)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 7, 2)
               End If
               If Len(Trim(Ximpgrastr)) = 9 Then
                  Ximpgrastr2 = "00000000" & Mid(Trim(Ximpgrastr), 1, 2)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 4, 3)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 8, 2)
               End If
               If Len(Trim(Ximpgrastr)) < 5 Then
                  Ximpgrastr2 = "000000000000000"
               End If
               Xlin = "10000001" + Trim(XCantex) + "02001024" + "001" + Trim(Xfec) + "085801009005000" + Trim(Xtarj) + Trim(Xventar) + Trim(Xfec) + Trim(Xnomimp) + "000000" + Trim(Xnommes) + Trim(Xmatstr) & "6" & "  " & Trim(Xnrorecstr) & Trim(Ximpgrastr2) & "                                                             "
               Print #1, Xlin
               data_inf.Recordset.AddNew
               data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
               data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
               data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
               data_inf.Recordset("cl_hon_pes") = XImp
               data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
               data_inf.Recordset("cl_nrotarj") = data_cli.Recordset("cl_nrotarj")
               data_inf.Recordset.Update
               data_emi.Recordset.MoveNext
            Loop
         End If
         data_emi.RecordSource = "Select * from mutuales where socio =" & data_cli.Recordset("cl_codigo") & " and importe_deuda >" & 0
         data_emi.Refresh
         If data_emi.Recordset.RecordCount > 0 Then
            data_emi.Recordset.MoveFirst
            Do While Not data_emi.Recordset.EOF
               Xtarj = Trim(data_cli.Recordset("cl_nrotarj"))
               Xcant = Xcant + 1
               If Len(Trim(str(Xcant))) = 1 Then
                  XCantex = "000000" + Trim(str(Xcant))
               End If
               If Len(Trim(str(Xcant))) = 2 Then
                  XCantex = "00000" + Trim(str(Xcant))
               End If
               If Len(Trim(str(Xcant))) = 3 Then
                  XCantex = "0000" + Trim(str(Xcant))
               End If
               If Len(Trim(str(Xcant))) = 4 Then
                  XCantex = "000" + Trim(str(Xcant))
               End If
               If Len(Trim(str(Int(data_cli.Recordset("cl_codigo"))))) = 2 Then
                  Xmatstr = "00000000" & Trim(str(Int(data_cli.Recordset("cl_codigo"))))
               End If
               If Len(Trim(str(Int(data_cli.Recordset("cl_codigo"))))) = 3 Then
                  Xmatstr = "0000000" & Trim(str(Int(data_cli.Recordset("cl_codigo"))))
               End If
               If Len(Trim(str(Int(data_cli.Recordset("cl_codigo"))))) = 4 Then
                  Xmatstr = "000000" & Trim(str(Int(data_cli.Recordset("cl_codigo"))))
               End If
               If Len(Trim(str(Int(data_cli.Recordset("cl_codigo"))))) = 5 Then
                  Xmatstr = "00000" & Trim(str(Int(data_cli.Recordset("cl_codigo"))))
               End If
               If Len(Trim(str(Int(data_cli.Recordset("cl_codigo"))))) = 6 Then
                  Xmatstr = "0000" & Trim(str(Int(data_cli.Recordset("cl_codigo"))))
               End If
               If Len(Trim(str(Int(data_cli.Recordset("cl_codigo"))))) = 7 Then
                  Xmatstr = "000" & Trim(str(Int(data_cli.Recordset("cl_codigo"))))
               End If
               If Len(Trim(str(Int(data_cli.Recordset("cl_codigo"))))) = 8 Then
                  Xmatstr = "00" & Trim(str(Int(data_cli.Recordset("cl_codigo"))))
               End If
               If Len(Trim(str(Int(data_cli.Recordset("cl_codigo"))))) = 9 Then
                  Xmatstr = "0" & Trim(str(Int(data_cli.Recordset("cl_codigo"))))
               End If
               If Len(Trim(str(Int(data_cli.Recordset("cl_codigo"))))) = 10 Then
                  Xmatstr = Trim(str(Int(data_cli.Recordset("cl_codigo"))))
               End If
               If Xmatstr = "" Then
                  Xmatstr = "0000000000"
               End If
               XImp = data_emi.Recordset("importe_deuda")
               If Len(Trim(str(Int(XImp)))) = 1 Then
                  Xnomimp = "000000000000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 2 Then
                  Xnomimp = "00000000000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 3 Then
                  Xnomimp = "0000000000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 4 Then
                  Xnomimp = "000000000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 5 Then
                  Xnomimp = "00000000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 6 Then
                  Xnomimp = "0000000" + Trim(str(Int(XImp))) + "00"
               End If
               Xtot = Xtot + Int(XImp)
               If IsNull(data_cli.Recordset("cl_tj_venc")) = False Then
                  Xventar = Trim(Mid(CStr(data_cli.Recordset("cl_tj_venc")), 4, 2))
                  Xventar = Trim(Xventar) & Trim(Mid(CStr(data_cli.Recordset("cl_tj_venc")), 9, 2))
               Else
                  Xventar = "1201"
               End If
               Xlin = "10000001" + Trim(XCantex) + "02001024" + "001" + Trim(Xfec) + "085801009005000" + Trim(Xtarj) + Trim(Xventar) + Trim(Xfec) + Trim(Xnomimp) + "000000" + Trim(Xnommes) + Trim(Xmatstr) & "6" & "                       "
               Xnrorecstr = ""
               If Len(Trim(str(Int(data_emi.Recordset("recibo"))))) = 1 Then
                  Xnrorecstr = "000000" & Trim(str(Int(data_emi.Recordset("recibo"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("recibo"))))) = 2 Then
                  Xnrorecstr = "00000" & Trim(str(Int(data_emi.Recordset("recibo"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("recibo"))))) = 3 Then
                  Xnrorecstr = "0000" & Trim(str(Int(data_emi.Recordset("recibo"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("recibo"))))) = 4 Then
                  Xnrorecstr = "000" & Trim(str(Int(data_emi.Recordset("recibo"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("recibo"))))) = 5 Then
                  Xnrorecstr = "00" & Trim(str(Int(data_emi.Recordset("recibo"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("recibo"))))) = 6 Then
                  Xnrorecstr = "0" & Trim(str(Int(data_emi.Recordset("recibo"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("recibo"))))) = 7 Then
                  Xnrorecstr = Trim(str(Int(data_emi.Recordset("recibo"))))
               End If
               If Xnrorecstr = "" Then
                  Xnrorecstr = "0000000"
               End If
               Ximpgrav = data_emi.Recordset("importe_deuda")
               Ximpgrav = Ximpgrav + Ximparqueo
               Ximpgrav = Ximpgrav / 1.1
               Ximpgrastr = Format(Ximpgrav, "Standard")
               If Len(Trim(Ximpgrastr)) = 5 Then
                  Ximpgrastr2 = "00000000000" & Mid(Trim(Ximpgrastr), 1, 2)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 4, 2)
               End If
               If Len(Trim(Ximpgrastr)) = 6 Then
                  Ximpgrastr2 = "0000000000" & Mid(Trim(Ximpgrastr), 1, 3)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 5, 2)
               End If
               If Len(Trim(Ximpgrastr)) = 8 Then
                  Ximpgrastr2 = "000000000" & Mid(Trim(Ximpgrastr), 1, 1)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 3, 3)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 7, 2)
               End If
               If Len(Trim(Ximpgrastr)) = 9 Then
                  Ximpgrastr2 = "00000000" & Mid(Trim(Ximpgrastr), 1, 2)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 4, 3)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 8, 2)
               End If
               If Len(Trim(Ximpgrastr)) < 5 Then
                  Ximpgrastr2 = "000000000000000"
               End If
               Xlin = "10000001" + Trim(XCantex) + "02001024" + "001" + Trim(Xfec) + "085801009005000" + Trim(Xtarj) + Trim(Xventar) + Trim(Xfec) + Trim(Xnomimp) + "000000" + Trim(Xnommes) + Trim(Xmatstr) & "6" & "  " & Trim(Xnrorecstr) & Trim(Ximpgrastr2) & "                                                             "
               Print #1, Xlin
               data_inf.Recordset.AddNew
               data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
               data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
               data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
               data_inf.Recordset("cl_hon_pes") = XImp
               data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
               data_inf.Recordset("cl_nrotarj") = data_cli.Recordset("cl_nrotarj")
               data_inf.Recordset.Update
               data_emi.Recordset.MoveNext
            Loop
         End If
      Else
         MsgBox "Verifique socio " & data_cli.Recordset("cl_codigo"), vbCritical
         End
      End If
      data_cli.Recordset.MoveNext
      Ximparqueo = 0
   Loop
   If Len(Trim(str(Int(Xtot)))) = 1 Then
      Xnomimp = "000000000000" + Trim(str(Int(Xtot))) + "00"
   End If
   If Len(Trim(str(Int(Xtot)))) = 2 Then
      Xnomimp = "00000000000" + Trim(str(Int(Xtot))) + "00"
   End If
   If Len(Trim(str(Int(Xtot)))) = 3 Then
      Xnomimp = "0000000000" + Trim(str(Int(Xtot))) + "00"
   End If
   If Len(Trim(str(Int(Xtot)))) = 4 Then
      Xnomimp = "000000000" + Trim(str(Int(Xtot))) + "00"
   End If
   If Len(Trim(str(Int(Xtot)))) = 5 Then
      Xnomimp = "00000000" + Trim(str(Int(Xtot))) + "00"
   End If
   If Len(Trim(str(Int(Xtot)))) = 6 Then
      Xnomimp = "0000000" + Trim(str(Int(Xtot))) + "00"
   End If
   If Len(Trim(str(Int(Xtot)))) = 7 Then
      Xnomimp = "000000" + Trim(str(Int(Xtot))) + "00"
   End If
    
   Xlin = "T02001024001" + Trim(Xfec) + "00000000"
   Xlin = Xlin + Trim(XCantex)
   Xlin = Xlin + Trim(Xnomimp)
   Xlin = Xlin + "0000000000000000000000000000000000000000000000000000AUTOMATICO                "
   Print #1, Xlin
   Open "c:\debitos\Visanet2.txt" For Output As #2
   Xlin = "10000001020010240001" + Trim(Xfec) + "085801009005" + Trim(XCantex) + Trim(Xnomimp) & "                                                                  "
   Print #2, Xlin
   Xlin = "T02001024001" + Trim(Xfec) + "00000000"
   Xlin = Xlin + Trim(XCantex)
   Xlin = Xlin + Trim(Xnomimp)
      
   Xlin = Xlin + "0000000000000000000000000000000000000000000000000000AUTOMATICO                "
   Print #2, Xlin
   data_inf.RecordSource = "select * from infcli"
   data_inf.Refresh
   Xtitulo = "DEBITOS VISA CORRESPONDIENTES MES : " & txt_mes.Text & "/" & txt_ano.Text
   cr1.ReportFileName = App.path & "\infdebitos.rpt"
   cr1.ReportTitle = Xtitulo
   cr1.Action = 1
    
    MsgBox "Proceso terminado", vbInformation, "Mensaje"
    Close #1
    Close #2
End If


End Sub

Public Sub Proceso_brou()
   Dim Micadbrou As String
   Dim XLacedbrou As String
   Dim Xmiimp As Long
   Dim Xtotregs, Xtotimpbr, Xtotregsm, Xtotimpbrm As Long
   Dim Xtotimpg, Xtotimpgm, Ximparqueo As Double
   Ximparqueo = 0
   Xtotimpg = 0
   Open "c:\debitos\BROU-" & Trim(str(txt_mes.Text)) & "-" & Trim(str(txt_ano.Text)) & ".txt" For Output As #1
'   Open "c:\debitos\sapp399.txt" For Output As #2
   data_cli.RecordSource = "Select * from clientes where cl_nrocobr =" & 607 & " and estado in (1) and fecha_baja is null"
   data_cli.Refresh
   If data_cli.Recordset.RecordCount > 0 Then
      data_cli.Recordset.MoveFirst
'''cedula titual---verificar que cargue cedula titular tarjeta
      Do While Not data_cli.Recordset.EOF
         data_emi.RecordSource = "Select * from deudas where cliente =" & data_cli.Recordset("cl_codigo") & " and fecha_pago is null and nro_cobr in (607) and fecha >='" & Format("01/08/2021", "yyyy-mm-dd") & "' and mes >" & 0
         data_emi.Refresh
         If data_emi.Recordset.RecordCount > 0 Then
            data_emi.Recordset.MoveFirst
            Do While Not data_emi.Recordset.EOF
               Xmiimp = data_emi.Recordset("total")
               
               Micadbrou = "1 00100"
               Micadbrou = Micadbrou + Mid(Trim(str(Year(mfec.Text))), 3, 2)
               If Month(mfec.Text) < 10 Then
                  Micadbrou = Micadbrou + "0" + Trim(str(Month(mfec.Text)))
               Else
                  Micadbrou = Micadbrou + Trim(str(Month(mfec.Text)))
               End If
               If Day(mfec.Text) < 10 Then
                  Micadbrou = Micadbrou + "0" + Trim(str(Day(mfec.Text)))
               Else
                  Micadbrou = Micadbrou + Trim(str(Day(mfec.Text)))
               End If
'               If Len(Trim(str(data_cli.Recordset("ci_tarj")))) > 7 Then
               If Len(Trim(str(data_cli.Recordset("cl_cedula")))) > 7 Then
                  MsgBox "Atención: error en número de cédula, verifique matrícula: " + Trim(str(data_cli.Recordset("cl_codigo"))), vbCritical, "Mensaje"
                  End
               Else
                  If Len(Trim(str(data_cli.Recordset("cl_cedula")))) < 4 Then
                     MsgBox "IMPOSIBLE CONTINUAR!!Atención!!: error en número de cédula, verifique matrícula: " + Trim(str(data_cli.Recordset("cl_codigo"))), vbCritical, "Mensaje"
                     End
                  End If
               End If
               If Len(Trim(str(data_cli.Recordset("cl_cedula")))) = 7 Then
'                  Micadbrou = Micadbrou + Trim(str(data_cli.Recordset("ci_tarj"))) + Trim(str(data_cli.Recordset("codcitarj")))
                  Micadbrou = Micadbrou + Trim(str(data_cli.Recordset("cl_cedula"))) + Trim(str(data_cli.Recordset("cl_codced")))
               End If
               If Len(Trim(str(data_cli.Recordset("cl_cedula")))) = 6 Then
                  Micadbrou = Micadbrou + "0" + Trim(str(data_cli.Recordset("cl_cedula"))) + Trim(str(data_cli.Recordset("cl_codced")))
               End If
               If Len(Trim(str(data_cli.Recordset("cl_cedula")))) = 5 Then
                  Micadbrou = Micadbrou + "00" + Trim(str(data_cli.Recordset("cl_cedula"))) + Trim(str(data_cli.Recordset("cl_codced")))
               End If
               Micadbrou = Micadbrou + "0000000"
               Micadbrou = Micadbrou + "000386"
               Xtotregs = Xtotregs + 1
               Xtotimpbr = Xtotimpbr + Xmiimp
               Micadbrou = Micadbrou + "98A00000000000"
               Micadbrou = Micadbrou + Mid(Trim(str(Year(mfec.Text))), 3, 2)
               If Month(mfec.Text) < 10 Then
                  Micadbrou = Micadbrou + "0" + Trim(str(Month(mfec.Text)))
               Else
                  Micadbrou = Micadbrou + Trim(str(Month(mfec.Text)))
               End If
               If Len(Trim(str(Int(Xmiimp)))) = 1 Then
                  Micadbrou = Micadbrou + "000000000000" + Trim(str(Int(Xmiimp))) + "00"
               End If
               If Len(Trim(str(Int(Xmiimp)))) = 2 Then
                  Micadbrou = Micadbrou + "00000000000" + Trim(str(Int(Xmiimp))) + "00"
               End If
               If Len(Trim(str(Int(Xmiimp)))) = 3 Then
                  Micadbrou = Micadbrou + "0000000000" + Trim(str(Int(Xmiimp))) + "00"
               End If
               If Len(Trim(str(Int(Xmiimp)))) = 4 Then
                  Micadbrou = Micadbrou + "000000000" + Trim(str(Int(Xmiimp))) + "00"
               End If
               If Len(Trim(str(Int(Xmiimp)))) = 5 Then
                  Micadbrou = Micadbrou + "00000000" + Trim(str(Int(Xmiimp))) + "00"
               End If
               If Len(Trim(str(Int(Xmiimp)))) = 6 Then
                  Micadbrou = Micadbrou + "0000000" + Trim(str(Int(Xmiimp))) + "00"
               End If
               If Len(Trim(str(Int(Xmiimp)))) = 7 Then
                  Micadbrou = Micadbrou + "000000" + Trim(str(Int(Xmiimp))) + "00"
               End If
               Micadbrou = Micadbrou + "0000000000000"
               Micadbrou = Micadbrou + "                                                "
               Xnrorecstr = ""
               If Len(Trim(str(Int(data_emi.Recordset("documento"))))) = 1 Then
                  Xnrorecstr = "00000000000" & Trim(str(Int(data_emi.Recordset("documento"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("documento"))))) = 2 Then
                  Xnrorecstr = "0000000000" & Trim(str(Int(data_emi.Recordset("documento"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("documento"))))) = 3 Then
                  Xnrorecstr = "000000000" & Trim(str(Int(data_emi.Recordset("documento"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("documento"))))) = 4 Then
                  Xnrorecstr = "00000000" & Trim(str(Int(data_emi.Recordset("documento"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("documento"))))) = 5 Then
                  Xnrorecstr = "0000000" & Trim(str(Int(data_emi.Recordset("documento"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("documento"))))) = 6 Then
                  Xnrorecstr = "000000" & Trim(str(Int(data_emi.Recordset("documento"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("documento"))))) = 7 Then
                  Xnrorecstr = "00000" & Trim(str(Int(data_emi.Recordset("documento"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("documento"))))) = 8 Then
                  Xnrorecstr = "0000" & Trim(str(Int(data_emi.Recordset("documento"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("documento"))))) = 9 Then
                  Xnrorecstr = "000" & Trim(str(Int(data_emi.Recordset("documento"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("documento"))))) = 10 Then
                  Xnrorecstr = "00" & Trim(str(Int(data_emi.Recordset("documento"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("documento"))))) = 11 Then
                  Xnrorecstr = "0" & Trim(str(Int(data_emi.Recordset("documento"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("documento"))))) = 12 Then
                  Xnrorecstr = Trim(str(Int(data_emi.Recordset("documento"))))
               End If
               If Xnrorecstr = "" Then
                  Xnrorecstr = "000000000000"
               End If
               Micadbrou = Micadbrou + Trim(Xnrorecstr) & "1"
               Ximpgrav = data_emi.Recordset("importe") + data_emi.Recordset("deudas")
               Ximpgrav = Ximpgrav + Ximparqueo
               Ximpgrav = Ximpgrav / 1.1
               Ximpgrav = Format(Ximpgrav, "Standard")
               Xtotimpg = Xtotimpg + Ximpgrav
               Ximpgrastr = Format(Ximpgrav, "Standard")
               If Len(Trim(Ximpgrastr)) = 5 Then
                  Ximpgrastr2 = "00000000000" & Mid(Trim(Ximpgrastr), 1, 2)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 4, 2)
               End If
               If Len(Trim(Ximpgrastr)) = 6 Then
                  Ximpgrastr2 = "0000000000" & Mid(Trim(Ximpgrastr), 1, 3)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 5, 2)
               End If
               If Len(Trim(Ximpgrastr)) = 8 Then
                  Ximpgrastr2 = "000000000" & Mid(Trim(Ximpgrastr), 1, 1)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 3, 3)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 7, 2)
               End If
               If Len(Trim(Ximpgrastr)) = 9 Then
                  Ximpgrastr2 = "00000000" & Mid(Trim(Ximpgrastr), 1, 2)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 4, 3)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 8, 2)
               End If
               If Len(Trim(Ximpgrastr)) < 5 Then
                  Ximpgrastr2 = "000000000000000"
               End If
               Micadbrou = Micadbrou + "000000000000000" & Trim(Ximpgrastr2) & "000000000000000"
               Print #1, Micadbrou
               data_inf.Recordset.AddNew
               data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
               data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
               data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
               data_inf.Recordset("cl_hon_pes") = Xmiimp
               data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
               data_inf.Recordset("cl_nrotarj") = data_cli.Recordset("cl_cedula")
               data_inf.Recordset.Update
               data_emi.Recordset.MoveNext
               Ximparqueo = 0
            Loop
         End If
''mutuales
         data_emi.RecordSource = "Select * from mutuales where socio =" & data_cli.Recordset("cl_codigo") & " and importe_deuda >" & 0
         data_emi.Refresh
         If data_emi.Recordset.RecordCount > 0 Then
            data_emi.Recordset.MoveFirst
            Do While Not data_emi.Recordset.EOF
               Xmiimp = data_emi.Recordset("importe_deuda")
               Micadbrou = "1 00100"
               Micadbrou = Micadbrou + Mid(Trim(str(Year(mfec.Text))), 3, 2)
               If Month(mfec.Text) < 10 Then
                  Micadbrou = Micadbrou + "0" + Trim(str(Month(mfec.Text)))
               Else
                  Micadbrou = Micadbrou + Trim(str(Month(mfec.Text)))
               End If
               If Day(mfec.Text) < 10 Then
                  Micadbrou = Micadbrou + "0" + Trim(str(Day(mfec.Text)))
               Else
                  Micadbrou = Micadbrou + Trim(str(Day(mfec.Text)))
               End If
               If Len(Trim(str(data_cli.Recordset("ci_tarj")))) > 7 Then
                  MsgBox "Atención: error en número de cédula, verifique matrícula: " + Trim(str(data_cli.Recordset("cl_codigo"))), vbCritical, "Mensaje"
                  End
               Else
                  If Len(Trim(str(data_cli.Recordset("ci_tarj")))) < 4 Then
                     MsgBox "IMPOSIBLE CONTINUAR!!Atención!!: error en número de cédula, verifique matrícula: " + Trim(str(data_cli.Recordset("cl_codigo"))), vbCritical, "Mensaje"
                     End
                  End If
               End If
               If Len(Trim(str(data_cli.Recordset("ci_tarj")))) = 7 Then
                  Micadbrou = Micadbrou + Trim(str(data_cli.Recordset("ci_tarj"))) + Trim(str(data_cli.Recordset("codcitarj")))
               End If
               If Len(Trim(str(data_cli.Recordset("ci_tarj")))) = 6 Then
                  Micadbrou = Micadbrou + "0" + Trim(str(data_cli.Recordset("ci_tarj"))) + Trim(str(data_cli.Recordset("codcitarj")))
               End If
               If Len(Trim(str(data_cli.Recordset("ci_tarj")))) = 5 Then
                  Micadbrou = Micadbrou + "00" + Trim(str(data_cli.Recordset("ci_tarj"))) + Trim(str(data_cli.Recordset("codcitarj")))
               End If
               Micadbrou = Micadbrou + "0000000"
               Micadbrou = Micadbrou + "000386"
               Xtotregs = Xtotregs + 1
               Xtotimpbr = Xtotimpbr + Xmiimp
               Micadbrou = Micadbrou + "98A00000000000"
               Micadbrou = Micadbrou + Mid(Trim(str(Year(mfec.Text))), 3, 2)
               If Month(mfec.Text) < 10 Then
                  Micadbrou = Micadbrou + "0" + Trim(str(Month(mfec.Text)))
               Else
                  Micadbrou = Micadbrou + Trim(str(Month(mfec.Text)))
               End If
               If Len(Trim(str(Int(Xmiimp)))) = 1 Then
                  Micadbrou = Micadbrou + "000000000000" + Trim(str(Int(Xmiimp))) + "00"
               End If
               If Len(Trim(str(Int(Xmiimp)))) = 2 Then
                  Micadbrou = Micadbrou + "00000000000" + Trim(str(Int(Xmiimp))) + "00"
               End If
               If Len(Trim(str(Int(Xmiimp)))) = 3 Then
                  Micadbrou = Micadbrou + "0000000000" + Trim(str(Int(Xmiimp))) + "00"
               End If
               If Len(Trim(str(Int(Xmiimp)))) = 4 Then
                  Micadbrou = Micadbrou + "000000000" + Trim(str(Int(Xmiimp))) + "00"
               End If
               If Len(Trim(str(Int(Xmiimp)))) = 5 Then
                  Micadbrou = Micadbrou + "00000000" + Trim(str(Int(Xmiimp))) + "00"
               End If
               If Len(Trim(str(Int(Xmiimp)))) = 6 Then
                  Micadbrou = Micadbrou + "0000000" + Trim(str(Int(Xmiimp))) + "00"
               End If
               If Len(Trim(str(Int(Xmiimp)))) = 7 Then
                  Micadbrou = Micadbrou + "000000" + Trim(str(Int(Xmiimp))) + "00"
               End If
               Micadbrou = Micadbrou + "0000000000000"
               Micadbrou = Micadbrou + "                                                "
               Xnrorecstr = ""
               If Len(Trim(str(Int(data_emi.Recordset("recibo"))))) = 1 Then
                  Xnrorecstr = "00000000000" & Trim(str(Int(data_emi.Recordset("recibo"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("recibo"))))) = 2 Then
                  Xnrorecstr = "0000000000" & Trim(str(Int(data_emi.Recordset("recibo"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("recibo"))))) = 3 Then
                  Xnrorecstr = "000000000" & Trim(str(Int(data_emi.Recordset("recibo"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("recibo"))))) = 4 Then
                  Xnrorecstr = "00000000" & Trim(str(Int(data_emi.Recordset("recibo"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("recibo"))))) = 5 Then
                  Xnrorecstr = "0000000" & Trim(str(Int(data_emi.Recordset("recibo"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("recibo"))))) = 6 Then
                  Xnrorecstr = "000000" & Trim(str(Int(data_emi.Recordset("recibo"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("recibo"))))) = 7 Then
                  Xnrorecstr = "00000" & Trim(str(Int(data_emi.Recordset("recibo"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("recibo"))))) = 8 Then
                  Xnrorecstr = "0000" & Trim(str(Int(data_emi.Recordset("recibo"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("recibo"))))) = 9 Then
                  Xnrorecstr = "000" & Trim(str(Int(data_emi.Recordset("recibo"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("recibo"))))) = 10 Then
                  Xnrorecstr = "00" & Trim(str(Int(data_emi.Recordset("recibo"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("recibo"))))) = 11 Then
                  Xnrorecstr = "0" & Trim(str(Int(data_emi.Recordset("recibo"))))
               End If
               If Len(Trim(str(Int(data_emi.Recordset("recibo"))))) = 12 Then
                  Xnrorecstr = Trim(str(Int(data_emi.Recordset("recibo"))))
               End If
               If Xnrorecstr = "" Then
                  Xnrorecstr = "000000000000"
               End If
               Micadbrou = Micadbrou + Trim(Xnrorecstr) & "1"
               Ximpgrav = data_emi.Recordset("importe_deuda")
               Ximpgrav = Ximpgrav + Ximparqueo
               Ximpgrav = Ximpgrav / 1.1
               Ximpgrav = Format(Ximpgrav, "Standard")
               Xtotimpg = Xtotimpg + Ximpgrav
               Ximpgrastr = Format(Ximpgrav, "Standard")
               If Len(Trim(Ximpgrastr)) = 5 Then
                  Ximpgrastr2 = "00000000000" & Mid(Trim(Ximpgrastr), 1, 2)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 4, 2)
               End If
               If Len(Trim(Ximpgrastr)) = 6 Then
                  Ximpgrastr2 = "0000000000" & Mid(Trim(Ximpgrastr), 1, 3)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 5, 2)
               End If
               If Len(Trim(Ximpgrastr)) = 8 Then
                  Ximpgrastr2 = "000000000" & Mid(Trim(Ximpgrastr), 1, 1)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 3, 3)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 7, 2)
               End If
               If Len(Trim(Ximpgrastr)) = 9 Then
                  Ximpgrastr2 = "00000000" & Mid(Trim(Ximpgrastr), 1, 2)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 4, 3)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 8, 2)
               End If
               If Len(Trim(Ximpgrastr)) < 5 Then
                  Ximpgrastr2 = "000000000000000"
               End If
               Micadbrou = Micadbrou + "000000000000000" & Trim(Ximpgrastr2) & "000000000000000"
               Print #1, Micadbrou
               data_inf.Recordset.AddNew
               data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
               data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
               data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
               data_inf.Recordset("cl_hon_pes") = Xmiimp
               data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
               data_inf.Recordset("cl_nrotarj") = data_cli.Recordset("cl_nrotarj")
               data_inf.Recordset.Update
               data_emi.Recordset.MoveNext
               Ximparqueo = 0
            Loop
         End If
         data_cli.Recordset.MoveNext
      Loop
      Micadbrou = "2 00100"
      Micadbrou = Micadbrou + Mid(Trim(str(Year(mfec.Text))), 3, 2)
      If Month(mfec.Text) < 10 Then
         Micadbrou = Micadbrou + "0" + Trim(str(Month(mfec.Text)))
      Else
         Micadbrou = Micadbrou + Trim(str(Month(mfec.Text)))
      End If
      If Day(mfec.Text) < 10 Then
         Micadbrou = Micadbrou + "0" + Trim(str(Day(mfec.Text)))
      Else
         Micadbrou = Micadbrou + Trim(str(Day(mfec.Text)))
      End If
      If Len(Trim(str(Int(Xtotregs)))) = 1 Then
         Micadbrou = Micadbrou + "00000" + Trim(str(Int(Xtotregs)))
      End If
      If Len(Trim(str(Int(Xtotregs)))) = 2 Then
         Micadbrou = Micadbrou + "0000" + Trim(str(Int(Xtotregs)))
      End If
      If Len(Trim(str(Int(Xtotregs)))) = 3 Then
         Micadbrou = Micadbrou + "000" + Trim(str(Int(Xtotregs)))
      End If
      If Len(Trim(str(Int(Xtotimpbr)))) = 1 Then
         Micadbrou = Micadbrou + "000000000000000" + Trim(str(Int(Xtotimpbr))) + "00"
      End If
      If Len(Trim(str(Int(Xtotimpbr)))) = 2 Then
         Micadbrou = Micadbrou + "00000000000000" + Trim(str(Int(Xtotimpbr))) + "00"
      End If
      If Len(Trim(str(Int(Xtotimpbr)))) = 3 Then
         Micadbrou = Micadbrou + "0000000000000" + Trim(str(Int(Xtotimpbr))) + "00"
      End If
      If Len(Trim(str(Int(Xtotimpbr)))) = 4 Then
         Micadbrou = Micadbrou + "000000000000" + Trim(str(Int(Xtotimpbr))) + "00"
      End If
      If Len(Trim(str(Int(Xtotimpbr)))) = 5 Then
         Micadbrou = Micadbrou + "00000000000" + Trim(str(Int(Xtotimpbr))) + "00"
      End If
      If Len(Trim(str(Int(Xtotimpbr)))) = 6 Then
         Micadbrou = Micadbrou + "0000000000" + Trim(str(Int(Xtotimpbr))) + "00"
      End If
      If Len(Trim(str(Int(Xtotimpbr)))) = 7 Then
         Micadbrou = Micadbrou + "000000000" + Trim(str(Int(Xtotimpbr))) + "00"
      End If
      Micadbrou = Micadbrou + "000000"
      Micadbrou = Micadbrou + "000000000000000000"
      Micadbrou = Micadbrou + "000000"
      Micadbrou = Micadbrou + "000000000000000000"
      Micadbrou = Micadbrou + "0000000000000000" & "0000000000000000" & "           " & "000000000000000000"
      Ximpgrastr = Format(Xtotimpg, "Standard")
      If Len(Trim(Ximpgrastr)) = 5 Then
         Ximpgrastr2 = "00000000000000" & Mid(Trim(Ximpgrastr), 1, 2)
         Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 4, 2)
      End If
      If Len(Trim(Ximpgrastr)) = 6 Then
         Ximpgrastr2 = "0000000000000" & Mid(Trim(Ximpgrastr), 1, 3)
         Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 5, 2)
      End If
      If Len(Trim(Ximpgrastr)) = 8 Then
         Ximpgrastr2 = "000000000000" & Mid(Trim(Ximpgrastr), 1, 1)
         Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 3, 3)
         Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 7, 2)
      End If
      If Len(Trim(Ximpgrastr)) = 9 Then
         Ximpgrastr2 = "00000000000" & Mid(Trim(Ximpgrastr), 1, 2)
         Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 4, 3)
         Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 8, 2)
      End If
      If Len(Trim(Ximpgrastr)) = 10 Then
         Ximpgrastr2 = "0000000000" & Mid(Trim(Ximpgrastr), 1, 3)
         Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 5, 3)
         Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 9, 2)
      End If
      If Len(Trim(Ximpgrastr)) < 5 Then
         Ximpgrastr2 = "000000000000000000"
      End If
      Micadbrou = Micadbrou + Trim(Ximpgrastr2) & "000000000000000000" & "0000"
      Print #1, Micadbrou;
'''''aquí termina el archivo 1
   Else
      MsgBox "No hay registros."
   End If
   Close #1

End Sub

Public Sub Proceso_master()
Dim Xcant As Integer
Dim Ximpgrastr, Xlin, Xtarj, Xfec, XCantex, Xnommes, Xnomimp, Xnomced, Xmesmast, Xmescabal, Xnommescab As String
Dim Ximpgrav, Xtotimpg As Double
Ximpgrav = 0
Xtotimpg = 0

If mfec.Text <> "__/__/____" Then
   Xfec = Trim(Mid(mfec.Text, 7, 4))
   Xfec = Xfec + Trim(Mid(mfec.Text, 4, 2))
   Xfec = Xfec + Trim(Mid(mfec.Text, 1, 2))
   Xmesmast = Trim(Mid(mfec.Text, 4, 2))
   Xmesmast = Xmesmast + "/" + Trim(Mid(mfec.Text, 9, 2))
   Xnommescab = Trim(Mid(mfec.Text, 4, 2))
   Xnommescab = Xnommescab + Trim(Mid(mfec.Text, 9, 2))
   Xmescabal = Trim(Mid(mfec.Text, 1, 2))
   Xmescabal = Xmescabal + Trim(Mid(mfec.Text, 4, 2))
   Xmescabal = Xmescabal + Trim(Mid(mfec.Text, 9, 2))
End If

Xlin = ""
data_cli.RecordSource = "Select * from clientes where cl_nrocobr =" & 683 & " and estado in (1) and fecha_baja is null" 'MASTER
data_cli.Refresh
If data_cli.Recordset.RecordCount > 0 Then
   Open "c:\debitos\DA168D.txt" For Output As #1
   data_cli.Recordset.MoveFirst
   Xcant = 0
   Do While Not data_cli.Recordset.EOF
      If Len(Trim(data_cli.Recordset("cl_nrotarj"))) = 16 Then
         data_emi.RecordSource = "Select * from deudas where cliente =" & data_cli.Recordset("cl_codigo") & " and fecha_pago is null and nro_cobr in (683) and fecha >='" & Format("01/08/2021", "yyyy-mm-dd") & "' and mes >" & 0
         data_emi.Refresh
         If data_emi.Recordset.RecordCount > 0 Then
            data_emi.Recordset.MoveFirst
            Do While Not data_emi.Recordset.EOF
               Xlin = "091643042"
               Xnomced = Trim(Int(data_emi.Recordset("documento")))
'               Xnomced = Xnomced + Trim(data_cli.Recordset("codcitarj"))
               Xtarj = Trim(data_cli.Recordset("cl_nrotarj"))
               Xlin = Xlin + Trim(Xtarj)
               Xcant = Xcant + 1
               If Len(Trim(Xnomced)) = 11 Then
                  Xlin = Xlin + "0" + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 10 Then
                  Xlin = Xlin + "00" + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 9 Then
                  Xlin = Xlin + "000" + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 8 Then
                  Xlin = Xlin + "0000" + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 7 Then
                  Xlin = Xlin + "00000" + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 6 Then
                  Xlin = Xlin + "000000" + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 5 Then
                  Xlin = Xlin + "0000000" + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 4 Then
                  Xlin = Xlin + "00000000" + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 3 Then
                  Xlin = Xlin + "000000000" + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 2 Then
                  Xlin = Xlin + "0000000000" + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 1 Then
                  Xlin = Xlin + "00000000000" + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 0 Then
                  Xlin = Xlin + "000000000000" + Trim(Xnomced)
               End If

               Xlin = Xlin + "00199901"
               XImp = data_emi.Recordset("total")
               If Len(Trim(str(Int(XImp)))) = 1 Then
                  Xnomimp = "00000000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 2 Then
                  Xnomimp = "0000000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 3 Then
                  Xnomimp = "000000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 4 Then
                  Xnomimp = "00000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 5 Then
                  Xnomimp = "0000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 6 Then
                  Xnomimp = "000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 7 Then
                  Xnomimp = "00" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 8 Then
                  Xnomimp = "0" + Trim(str(Int(XImp))) + "00"
               End If
               Xlin = Xlin + Trim(Xnomimp) + Trim(Xmesmast)
               Xlin = Xlin + "       "
               Xlin = Xlin + "                                        1"
               Ximpgrav = data_emi.Recordset("total")
               Ximpgrav = Ximpgrav / 1.1
               Ximpgrav = Format(Ximpgrav, "Standard")
               Ximpgrastr = Format(Ximpgrav, "Standard")
               If Len(Trim(Ximpgrastr)) = 5 Then
                  Ximpgrastr2 = "0000000" & Mid(Trim(Ximpgrastr), 1, 2)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 4, 2)
               End If
               If Len(Trim(Ximpgrastr)) = 6 Then
                  Ximpgrastr2 = "000000" & Mid(Trim(Ximpgrastr), 1, 3)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 5, 2)
               End If
               If Len(Trim(Ximpgrastr)) = 8 Then
                  Ximpgrastr2 = "00000" & Mid(Trim(Ximpgrastr), 1, 1)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 3, 3)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 7, 2)
               End If
               If Len(Trim(Ximpgrastr)) = 9 Then
                  Ximpgrastr2 = "0000" & Mid(Trim(Ximpgrastr), 1, 2)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 4, 3)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 8, 2)
               End If
               If Len(Trim(Ximpgrastr)) < 5 Then
                  Ximpgrastr2 = "000000000000"
               End If
               Xlin = Xlin + Trim(Ximpgrastr2)
               Ximpgrastr = ""
               Ximpgrastr = Trim(str(Int(data_emi.Recordset("documento"))))
               If Len(Trim(Ximpgrastr)) = 4 Then
                  Xlin = Xlin + "0000000000000000" + Trim(Ximpgrastr)
               End If
               If Len(Trim(Ximpgrastr)) = 5 Then
                  Xlin = Xlin + "000000000000000" + Trim(Ximpgrastr)
               End If
               If Len(Trim(Ximpgrastr)) = 6 Then
                  Xlin = Xlin + "00000000000000" + Trim(Ximpgrastr)
               End If
               If Len(Trim(Ximpgrastr)) = 7 Then
                  Xlin = Xlin + "0000000000000" + Trim(Ximpgrastr)
               End If
               If Len(Trim(Ximpgrastr)) = 8 Then
                  Xlin = Xlin + "000000000000" + Trim(Ximpgrastr)
               End If
               If Len(Trim(Ximpgrastr)) = 9 Then
                  Xlin = Xlin + "00000000000" + Trim(Ximpgrastr)
               End If
               If Len(Trim(Ximpgrastr)) = 10 Then
                  Xlin = Xlin + "0000000000" + Trim(Ximpgrastr)
               End If
               If Len(Trim(Ximpgrastr)) = 11 Then
                  Xlin = Xlin + "000000000" + Trim(Ximpgrastr)
               End If
               Xlin = Xlin + "                                                            "
               Xtot = Xtot + XImp
               data_inf.Recordset.AddNew
               data_inf.Recordset("info_debit") = Xlin
               data_inf.Recordset("cl_hon_pes") = XImp
               data_inf.Recordset.Update
               data_inf.Recordset.AddNew
               data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
               data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
               data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
               data_inf.Recordset("cl_hon_pes") = XImp
               data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
               data_inf.Recordset("cl_nrotarj") = data_cli.Recordset("cl_nrotarj")
               data_inf.Recordset("info_debit") = "NO"
               data_inf.Recordset.Update
               data_emi.Recordset.MoveNext
            Loop
         End If
         data_emi.RecordSource = "Select * from mutuales where socio =" & data_cli.Recordset("cl_codigo") & " and importe_deuda >" & 0
         data_emi.Refresh
         If data_emi.Recordset.RecordCount > 0 Then
            data_emi.Recordset.MoveFirst
            Do While Not data_emi.Recordset.EOF
               Xlin = "091643042"
               Xnomced = Trim(Int(data_emi.Recordset("recibo")))
'               Xnomced = Xnomced + Trim(data_cli.Recordset("codcitarj"))
               Xtarj = Trim(data_cli.Recordset("cl_nrotarj"))
               Xlin = Xlin + Trim(Xtarj)
               Xcant = Xcant + 1
               If Len(Trim(Xnomced)) = 11 Then
                  Xlin = Xlin + "0" + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 10 Then
                  Xlin = Xlin + "00" + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 9 Then
                  Xlin = Xlin + "000" + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 8 Then
                  Xlin = Xlin + "0000" + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 7 Then
                  Xlin = Xlin + "00000" + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 6 Then
                  Xlin = Xlin + "000000" + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 5 Then
                  Xlin = Xlin + "0000000" + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 4 Then
                  Xlin = Xlin + "00000000" + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 3 Then
                  Xlin = Xlin + "000000000" + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 2 Then
                  Xlin = Xlin + "0000000000" + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 1 Then
                  Xlin = Xlin + "00000000000" + Trim(Xnomced)
               End If
               Xlin = Xlin + "00199901"
               XImp = data_emi.Recordset("importe_deuda")
               If Len(Trim(str(Int(XImp)))) = 1 Then
                  Xnomimp = "00000000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 2 Then
                  Xnomimp = "0000000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 3 Then
                  Xnomimp = "000000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 4 Then
                  Xnomimp = "00000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 5 Then
                  Xnomimp = "0000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 6 Then
                  Xnomimp = "000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 7 Then
                  Xnomimp = "00" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 8 Then
                  Xnomimp = "0" + Trim(str(Int(XImp))) + "00"
               End If
               Xlin = Xlin + Trim(Xnomimp) + Trim(Xmesmast)
               Xlin = Xlin + "       "
               Xlin = Xlin + "                                        1"
               Ximpgrav = data_emi.Recordset("importe_deuda")
               Ximpgrav = Ximpgrav / 1.1
               Ximpgrav = Format(Ximpgrav, "Standard")
               Ximpgrastr = Format(Ximpgrav, "Standard")
               If Len(Trim(Ximpgrastr)) = 5 Then
                  Ximpgrastr2 = "0000000" & Mid(Trim(Ximpgrastr), 1, 2)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 4, 2)
               End If
               If Len(Trim(Ximpgrastr)) = 6 Then
                  Ximpgrastr2 = "000000" & Mid(Trim(Ximpgrastr), 1, 3)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 5, 2)
               End If
               If Len(Trim(Ximpgrastr)) = 8 Then
                  Ximpgrastr2 = "00000" & Mid(Trim(Ximpgrastr), 1, 1)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 3, 3)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 7, 2)
               End If
               If Len(Trim(Ximpgrastr)) = 9 Then
                  Ximpgrastr2 = "0000" & Mid(Trim(Ximpgrastr), 1, 2)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 4, 3)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 8, 2)
               End If
               If Len(Trim(Ximpgrastr)) < 5 Then
                  Ximpgrastr2 = "00000000000"
               End If
               Xlin = Xlin + Trim(Ximpgrastr2)
               Ximpgrastr = ""
               Ximpgrastr = Trim(str(Int(data_emi.Recordset("recibo"))))
               If Len(Trim(Ximpgrastr)) = 1 Then
                  Xlin = Xlin + "0000000000000000000" + Trim(Ximpgrastr)
               End If
               If Len(Trim(Ximpgrastr)) = 2 Then
                  Xlin = Xlin + "000000000000000000" + Trim(Ximpgrastr)
               End If
               If Len(Trim(Ximpgrastr)) = 3 Then
                  Xlin = Xlin + "00000000000000000" + Trim(Ximpgrastr)
               End If
               If Len(Trim(Ximpgrastr)) = 4 Then
                  Xlin = Xlin + "0000000000000000" + Trim(Ximpgrastr)
               End If
               If Len(Trim(Ximpgrastr)) = 5 Then
                  Xlin = Xlin + "000000000000000" + Trim(Ximpgrastr)
               End If
               If Len(Trim(Ximpgrastr)) = 6 Then
                  Xlin = Xlin + "00000000000000" + Trim(Ximpgrastr)
               End If
               If Len(Trim(Ximpgrastr)) = 7 Then
                  Xlin = Xlin + "0000000000000" + Trim(Ximpgrastr)
               End If
               If Len(Trim(Ximpgrastr)) = 8 Then
                  Xlin = Xlin + "000000000000" + Trim(Ximpgrastr)
               End If
               If Len(Trim(Ximpgrastr)) = 9 Then
                  Xlin = Xlin + "00000000000" + Trim(Ximpgrastr)
               End If
               If Len(Trim(Ximpgrastr)) = 10 Then
                  Xlin = Xlin + "0000000000" + Trim(Ximpgrastr)
               End If
               If Len(Trim(Ximpgrastr)) = 11 Then
                  Xlin = Xlin + "000000000" + Trim(Ximpgrastr)
               End If
               Xlin = Xlin + "                                                            "
               Xtot = Xtot + XImp
               data_inf.Recordset.AddNew
               data_inf.Recordset("info_debit") = Xlin
               data_inf.Recordset("cl_hon_pes") = XImp
               data_inf.Recordset.Update
               data_inf.Recordset.AddNew
               data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
               data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
               data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
               data_inf.Recordset("cl_hon_pes") = XImp
               data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
               data_inf.Recordset("cl_nrotarj") = data_cli.Recordset("cl_nrotarj")
               data_inf.Recordset("info_debit") = "NO"
               data_inf.Recordset.Update
               data_emi.Recordset.MoveNext
            Loop
         End If
      Else
         MsgBox "Error en una tarjeta:" & data_cli.Recordset("cl_codigo")
         End
      End If
      data_cli.Recordset.MoveNext
      Ximparqueo = 0
   Loop
   Xlin = "091643041"
   Xnommes = Trim(Mid(mfec.Text, 1, 2))
   Xnommes = Xnommes + Trim(Mid(mfec.Text, 4, 2))
   Xnommes = Xnommes + Trim(Mid(mfec.Text, 9, 2))
   Xlin = Xlin + Trim(Xnommes)
   If Len(Trim(str(Xcant))) = 1 Then
      XCantex = "000000" + Trim(str(Xcant))
   End If
   If Len(Trim(str(Xcant))) = 2 Then
      XCantex = "00000" + Trim(str(Xcant))
   End If
   If Len(Trim(str(Xcant))) = 3 Then
      XCantex = "0000" + Trim(str(Xcant))
   End If
   If Len(Trim(str(Xcant))) = 4 Then
      XCantex = "000" + Trim(str(Xcant))
   End If
   Xlin = Xlin + Trim(XCantex)
   If Len(Trim(str(Int(Xtot)))) = 1 Then
      Xnomimp = "000000000000" + Trim(str(Int(Xtot))) + "00"
   End If
   If Len(Trim(str(Int(Xtot)))) = 2 Then
      Xnomimp = "00000000000" + Trim(str(Int(Xtot))) + "00"
   End If
   If Len(Trim(str(Int(Xtot)))) = 3 Then
      Xnomimp = "0000000000" + Trim(str(Int(Xtot))) + "00"
   End If
   If Len(Trim(str(Int(Xtot)))) = 4 Then
      Xnomimp = "000000000" + Trim(str(Int(Xtot))) + "00"
   End If
   If Len(Trim(str(Int(Xtot)))) = 5 Then
      Xnomimp = "00000000" + Trim(str(Int(Xtot))) + "00"
   End If
   If Len(Trim(str(Int(Xtot)))) = 6 Then
      Xnomimp = "0000000" + Trim(str(Int(Xtot))) + "00"
   End If
   Xlin = Xlin + Trim(Xnomimp)
   Xlin = Xlin + "                                                   "
   Xlin = Xlin + "                                                  "
   Xlin = Xlin + "                                                  "
   Xlin = Xlin + "            "
   Print #1, Xlin
   data_inf.Refresh
   data_inf.Recordset.MoveFirst
   Do While Not data_inf.Recordset.EOF
      If IsNull(data_inf.Recordset("info_debit")) = False Then
         If data_inf.Recordset("info_debit") <> "" Then
            If Len(Trim(data_inf.Recordset("info_debit"))) > 55 Then
               Print #1, data_inf.Recordset("info_debit")
            End If
          End If
      End If
      data_inf.Recordset.MoveNext
   Loop
   Close #1
   data_inf.Refresh
   If data_inf.Recordset.RecordCount > 0 Then
      data_inf.Recordset.MoveFirst
      Do While Not data_inf.Recordset.EOF
         If data_inf.Recordset("info_debit") = "NO" Then
         Else
            data_inf.Recordset.Delete
         End If
         data_inf.Recordset.MoveNext
      Loop
   End If
End If

End Sub

Public Sub Proceso_redpagos()
Dim Xlineat As String
Dim Xobjexel22 As Excel.Application
Dim Xlibexel22 As Excel.Workbook
Dim Xarchexel22 As New Excel.Worksheet
Dim Xlin, XCol As Integer
Dim Xtotreg, Xsub As Long
Dim Xarchtex As String
Dim Xlabrir3 As New Excel.Application
Dim Xmat As Long
Dim TextodeFecha As String
TextodeFecha = ""
If Month(Date) > 9 Then
   If Day(Date) > 9 Then
      TextodeFecha = Trim(str(Year(Date))) & Trim(str(Month(Date))) & Trim(str(Day(Date)))
   Else
      TextodeFecha = Trim(str(Year(Date))) & Trim(str(Month(Date))) & "0" & Trim(str(Day(Date)))
   End If
Else
   If Day(Date) > 9 Then
      TextodeFecha = Trim(str(Year(Date))) & "0" & Trim(str(Month(Date))) & Trim(str(Day(Date)))
   Else
      TextodeFecha = Trim(str(Year(Date))) & "0" & Trim(str(Month(Date))) & "0" & Trim(str(Day(Date)))
   End If
End If
Xmat = 0
Xlineat = ""
Xtotreg = 0
'data_arq.Connect = "odbc;dsn=sappnew;"
'data_arq.ConnectionString = "dsn="
data_arq.RecordSource = "select deudas.nro_cobr,deudas.nombre,deudas.fecha_pago,deudas.ano,deudas.mes,deudas.total,deudas.fecha,deudas.documento,deudas.servi,deudas.cliente," & _
"deudas.fecha_pago,clientes.cl_codigo,clientes.cl_cedula,clientes.cl_codced,clientes.estado from deudas inner join clientes on deudas.cliente=clientes.cl_codigo" & _
" where clientes.estado in (1) and deudas.fecha_pago is null and deudas.nro_cobr in (221) order by deudas.cliente,deudas.fecha"
data_arq.Refresh
If data_arq.Recordset.RecordCount > 0 Then
   data_arq.Recordset.MoveFirst
   Xlin = 1
   XCol = 1
   MsgBox "El archivo se guardará en la carpeta Debitos del disco C", vbInformation
   
   Open "c:\debitos\RedPagos-" & TextodeFecha & ".csv" For Output As #1
   
   
   'Set Xobjexel22 = New Excel.Application
   'Set Xlibexel22 = Xobjexel22.Workbooks.Add
   'Set Xarchexel22 = Xlibexel22.Worksheets.Add
   'Xarchexel22.Name = Trim("Sapp")
   'Xlibexel22.SaveAs ("C:\debitos\RedPagos-" & TextodeFecha & ".csv")
   'Xarchtex = "C:\debitos\RedPagos-" & TextodeFecha & ".csv"
   Xmat = data_arq.Recordset("cliente")
   Do While Not data_arq.Recordset.EOF
      Xlineat = Trim(str(data_arq.Recordset("ano"))) & "," & Trim(str(data_arq.Recordset("mes"))) & "," & Trim(str(Xtotreg)) & "," & Trim(str(data_arq.Recordset("cl_cedula"))) & Trim(str(data_arq.Recordset("cl_codced"))) & "," & _
      data_arq.Recordset("nombre") & ",0," & Trim(str(Val(data_arq.Recordset("total")))) & "," & Format(data_arq.Recordset("fecha"), "dd/mm/yyyy") & "," & _
      Format(data_arq.Recordset("fecha"), "dd/mm/yyyy") & "," & Trim(str(data_arq.Recordset("documento"))) & "," & Trim(str(1)) & "," & Format(data_arq.Recordset("servi"), "###0.00")
      Xmat = data_arq.Recordset("cliente")
      data_arq.Recordset.MoveNext
      If data_arq.Recordset.EOF = True Then
      Else
         If Xmat = data_arq.Recordset("cliente") Then
            Xtotreg = Xtotreg + 1
         Else
            Xtotreg = 0
         End If
      End If
      Print #1, Xlineat
      'Xarchexel22.Cells(Xlin, XCol) = Xlineat
      Xlin = Xlin + 1
      Xlineat = ""
   Loop
   Close #1
   
'   Xlibexel22.Save
'   Xlibexel22.Close
'   Xobjexel22.Quit
'   Xlabrir3.Workbooks.Open Xarchtex, , False
'   Xlabrir3.Visible = True
'   Xlabrir3.WindowState = xlMaximized
   MsgBox "Proceso terminado. El archivo está en la carpeta débitos con el nombre:RedPagos-" & TextodeFecha & ".csv"
   
End If

End Sub

Public Sub Proceso_oca()
Dim Xcant As Integer
Dim Xlin, Xtarj, Xfec, XCantex, Xnommes, Xnomimp, Xnomced, Xmesmast, Xmescabal, Xnommescab, CliOcaStr As String
Dim CliOca As Long
Dim XImp, Xtot, Xtardelf, Ximparqueo As Double
Dim Ximpgrav As Double
Dim Xmatstr, Xnrorecstr, Ximpgrastr, Ximpgrastr2 As String

XImp = 0
CliOca = 0
CliOcaStr = ""
If mfec.Text <> "__/__/____" Then
   Xfec = Trim(Mid(mfec.Text, 7, 4))
   Xfec = Xfec + Trim(Mid(mfec.Text, 4, 2))
   Xfec = Xfec + Trim(Mid(mfec.Text, 1, 2))
   Xmesmast = Trim(Mid(mfec.Text, 4, 2))
   Xmesmast = Xmesmast + "/" + Trim(Mid(mfec.Text, 9, 2))
   Xnommescab = Trim(Mid(mfec.Text, 4, 2))
   Xnommescab = Xnommescab + Trim(Mid(mfec.Text, 9, 2))
   Xmescabal = Trim(Mid(mfec.Text, 1, 2))
   Xmescabal = Xmescabal + Trim(Mid(mfec.Text, 4, 2))
   Xmescabal = Xmescabal + Trim(Mid(mfec.Text, 9, 2))
End If

Xlin = ""
data_cli.RecordSource = "Select * from clientes where cl_nrocobr =" & 690 & " and estado in (1) and fecha_baja is null and ci_tarj >" & 0 'OCA
data_cli.Refresh
If data_cli.Recordset.RecordCount > 0 Then
   Open "c:\debitos\OCA-" & Trim(str(Month(Date))) & "-" & Trim(str(Year(Date))) & ".txt" For Output As #1
   data_cli.Recordset.MoveFirst
   Xcant = 0
   Do While Not data_cli.Recordset.EOF
      If Len(Trim(str(data_cli.Recordset("ci_tarj")))) >= 5 Then
         data_emi.RecordSource = "Select * from deudas where cliente =" & data_cli.Recordset("cl_codigo") & " and fecha_pago is null and nro_cobr in (690) and fecha >='" & Format("01/07/2021", "yyyy-mm-dd") & "' and mes >" & 0
         data_emi.Refresh
         If data_emi.Recordset.RecordCount > 0 Then
            data_emi.Recordset.MoveFirst
            Do While Not data_emi.Recordset.EOF
               Xcant = Xcant + 1
               CliOca = Int(data_cli.Recordset("cl_codigo"))
               CliOcaStr = Trim(str(CliOca))
               If Len(Trim(CliOcaStr)) = 12 Then
                  Xlin = "                                      " + Trim(CliOcaStr)
               End If
               If Len(Trim(CliOcaStr)) = 11 Then
                  Xlin = "                                       " + Trim(CliOcaStr)
               End If
               If Len(Trim(CliOcaStr)) = 10 Then
                  Xlin = "                                        " + Trim(CliOcaStr)
               End If
               If Len(Trim(CliOcaStr)) = 9 Then
                  Xlin = "                                         " + Trim(CliOcaStr)
               End If
               If Len(Trim(CliOcaStr)) = 8 Then
                  Xlin = "                                          " + Trim(CliOcaStr)
               End If
               If Len(Trim(CliOcaStr)) = 7 Then
                  Xlin = "                                           " + Trim(CliOcaStr)
               End If
               If Len(Trim(CliOcaStr)) = 6 Then
                  Xlin = "                                            " + Trim(CliOcaStr)
               End If
               If Len(Trim(CliOcaStr)) = 5 Then
                  Xlin = "                                             " + Trim(CliOcaStr)
               End If
               If Len(Trim(CliOcaStr)) = 4 Then
                  Xlin = "                                              " + Trim(CliOcaStr)
               End If
               If Len(Trim(CliOcaStr)) = 3 Then
                  Xlin = "                                               " + Trim(CliOcaStr)
               End If
               If Len(Trim(CliOcaStr)) = 2 Then
                  Xlin = "                                                " + Trim(CliOcaStr)
               End If
               If Len(Trim(CliOcaStr)) = 1 Then
                  Xlin = "                                                 " + Trim(CliOcaStr)
               End If
               Xnomced = Trim(Int(data_cli.Recordset("ci_tarj")))
'               Xnomced = Xnomced + Trim(data_cli.Recordset("codcitarj"))
               If Len(Trim(Xnomced)) = 4 Then
                  Xlin = Xlin + "      " + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 5 Then
                  Xlin = Xlin + "     " + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 6 Then
                  Xlin = Xlin + "    " + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 7 Then
                  Xlin = Xlin + "   " + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 8 Then
                  Xlin = Xlin + "  " + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 9 Then
                  Xlin = Xlin + " " + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 10 Then
                  Xlin = Xlin + Trim(Xnomced)
               End If
               Xlin = Xlin + " 858"
               XImp = data_emi.Recordset("total")
               If Len(Trim(str(Int(XImp)))) = 1 Then
                  Xnomimp = "0000000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 2 Then
                  Xnomimp = "000000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 3 Then
                  Xnomimp = "00000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 4 Then
                  Xnomimp = "0000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 5 Then
                  Xnomimp = "000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 6 Then
                  Xnomimp = "00" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 7 Then
                  Xnomimp = "0" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 8 Then
                  Xnomimp = Trim(str(Int(XImp))) + "00"
               End If
               Xlin = Xlin + Trim(Xnomimp) + "1"
               Xdescrecibo = Trim(str(Int(data_emi.Recordset("documento"))))
               If Len(Trim(str(Int(Xdescrecibo)))) = 3 Then
                  Xlin = Xlin + "                 " + Trim(Xdescrecibo)
               Else
                  If Len(Trim(str(Int(Xdescrecibo)))) = 4 Then
                     Xlin = Xlin + "                " + Trim(Xdescrecibo)
                  Else
                     If Len(Trim(str(Int(Xdescrecibo)))) = 5 Then
                        Xlin = Xlin + "               " + Trim(Xdescrecibo)
                     Else
                        If Len(Trim(str(Int(Xdescrecibo)))) = 6 Then
                           Xlin = Xlin + "              " + Trim(Xdescrecibo)
                        Else
                           If Len(Trim(str(Int(Xdescrecibo)))) = 7 Then
                              Xlin = Xlin + "             " + Trim(Xdescrecibo)
                           Else
                              If Len(Trim(str(Int(Xdescrecibo)))) = 8 Then
                                 Xlin = Xlin + "            " + Trim(Xdescrecibo)
                              Else
                                 If Len(Trim(str(Int(Xdescrecibo)))) = 9 Then
                                    Xlin = Xlin + "           " + Trim(Xdescrecibo)
                                 Else
                                    If Len(Trim(str(Int(Xdescrecibo)))) = 10 Then
                                       Xlin = Xlin + "           " + Trim(Xdescrecibo)
                                    Else
                                       Xlin = Xlin + "                9999"
                                    End If
                                 End If
                              End If
                           End If
                        End If
                     End If
                  End If
               End If
               Ximpgrav = data_emi.Recordset("total")
               Ximpgrav = Ximpgrav / 1.1
               Ximpgrav = Format(Ximpgrav, "Standard")
               Ximpgrastr = Format(Ximpgrav, "Standard")
               If Len(Trim(Ximpgrastr)) = 5 Then
                  Ximpgrastr2 = "     " & Mid(Trim(Ximpgrastr), 1, 2)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 4, 2)
               End If
               If Len(Trim(Ximpgrastr)) = 6 Then
                  Ximpgrastr2 = "    " & Mid(Trim(Ximpgrastr), 1, 3)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 5, 2)
               End If
               If Len(Trim(Ximpgrastr)) = 8 Then
                  Ximpgrastr2 = "   " & Mid(Trim(Ximpgrastr), 1, 1)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 3, 3)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 7, 2)
               End If
               If Len(Trim(Ximpgrastr)) = 9 Then
                  Ximpgrastr2 = "  " & Mid(Trim(Ximpgrastr), 1, 2)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 4, 3)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 8, 2)
               End If
               Xlin = Xlin + Ximpgrastr2 & "        0"
               Print #1, Xlin
               data_inf.Recordset.AddNew
               data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
               data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
               data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
               data_inf.Recordset("cl_hon_pes") = XImp
               data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
               data_inf.Recordset("cl_nrotarj") = data_cli.Recordset("cl_nrotarj")
               data_inf.Recordset.Update
                
               data_emi.Recordset.MoveNext
            Loop
         End If
         data_emi.RecordSource = "Select * from mutuales where socio =" & data_cli.Recordset("cl_codigo") & " and importe_deuda >" & 0
         data_emi.Refresh
         If data_emi.Recordset.RecordCount > 0 Then
            data_emi.Recordset.MoveFirst
            Do While Not data_emi.Recordset.EOF
               Xcant = Xcant + 1
               CliOca = Int(data_cli.Recordset("cl_codigo"))
               CliOcaStr = Trim(str(CliOca))
               If Len(Trim(CliOcaStr)) = 12 Then
                  Xlin = "                                      " + Trim(CliOcaStr)
               End If
               If Len(Trim(CliOcaStr)) = 11 Then
                  Xlin = "                                       " + Trim(CliOcaStr)
               End If
               If Len(Trim(CliOcaStr)) = 10 Then
                  Xlin = "                                        " + Trim(CliOcaStr)
               End If
               If Len(Trim(CliOcaStr)) = 9 Then
                  Xlin = "                                         " + Trim(CliOcaStr)
               End If
               If Len(Trim(CliOcaStr)) = 8 Then
                  Xlin = "                                          " + Trim(CliOcaStr)
               End If
               If Len(Trim(CliOcaStr)) = 7 Then
                  Xlin = "                                           " + Trim(CliOcaStr)
               End If
               If Len(Trim(CliOcaStr)) = 6 Then
                  Xlin = "                                            " + Trim(CliOcaStr)
               End If
               If Len(Trim(CliOcaStr)) = 5 Then
                  Xlin = "                                             " + Trim(CliOcaStr)
               End If
               If Len(Trim(CliOcaStr)) = 4 Then
                  Xlin = "                                              " + Trim(CliOcaStr)
               End If
               If Len(Trim(CliOcaStr)) = 3 Then
                  Xlin = "                                               " + Trim(CliOcaStr)
               End If
               If Len(Trim(CliOcaStr)) = 2 Then
                  Xlin = "                                                " + Trim(CliOcaStr)
               End If
               If Len(Trim(CliOcaStr)) = 1 Then
                  Xlin = "                                                 " + Trim(CliOcaStr)
               End If
               Xnomced = Trim(Int(data_cli.Recordset("ci_tarj")))
'               Xnomced = Xnomced + Trim(data_cli.Recordset("codcitarj"))
               If Len(Trim(Xnomced)) = 4 Then
                  Xlin = Xlin + "      " + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 5 Then
                  Xlin = Xlin + "     " + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 6 Then
                  Xlin = Xlin + "    " + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 7 Then
                  Xlin = Xlin + "   " + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 8 Then
                  Xlin = Xlin + "  " + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 9 Then
                  Xlin = Xlin + " " + Trim(Xnomced)
               End If
               If Len(Trim(Xnomced)) = 10 Then
                  Xlin = Xlin + Trim(Xnomced)
               End If
               Xlin = Xlin + " 858"
               XImp = data_emi.Recordset("importe_deuda")
               If Len(Trim(str(Int(XImp)))) = 1 Then
                  Xnomimp = "0000000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 2 Then
                  Xnomimp = "000000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 3 Then
                  Xnomimp = "00000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 4 Then
                  Xnomimp = "0000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 5 Then
                  Xnomimp = "000" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 6 Then
                  Xnomimp = "00" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 7 Then
                  Xnomimp = "0" + Trim(str(Int(XImp))) + "00"
               End If
               If Len(Trim(str(Int(XImp)))) = 8 Then
                  Xnomimp = Trim(str(Int(XImp))) + "00"
               End If
               Xlin = Xlin + Trim(Xnomimp) + "1"
               Xdescrecibo = Trim(str(Int(data_emi.Recordset("recibo"))))
               If Len(Trim(str(Int(Xdescrecibo)))) = 3 Then
                  Xlin = Xlin + "                 " + Trim(Xdescrecibo)
               Else
                  If Len(Trim(str(Int(Xdescrecibo)))) = 4 Then
                     Xlin = Xlin + "                " + Trim(Xdescrecibo)
                  Else
                     If Len(Trim(str(Int(Xdescrecibo)))) = 5 Then
                        Xlin = Xlin + "               " + Trim(Xdescrecibo)
                     Else
                        If Len(Trim(str(Int(Xdescrecibo)))) = 6 Then
                           Xlin = Xlin + "              " + Trim(Xdescrecibo)
                        Else
                           If Len(Trim(str(Int(Xdescrecibo)))) = 7 Then
                              Xlin = Xlin + "             " + Trim(Xdescrecibo)
                           Else
                              If Len(Trim(str(Int(Xdescrecibo)))) = 8 Then
                                 Xlin = Xlin + "            " + Trim(Xdescrecibo)
                              Else
                                 If Len(Trim(str(Int(Xdescrecibo)))) = 9 Then
                                    Xlin = Xlin + "           " + Trim(Xdescrecibo)
                                 Else
                                    If Len(Trim(str(Int(Xdescrecibo)))) = 10 Then
                                       Xlin = Xlin + "           " + Trim(Xdescrecibo)
                                    Else
                                       Xlin = Xlin + "                9999"
                                    End If
                                 End If
                              End If
                           End If
                        End If
                     End If
                  End If
               End If
               Ximpgrav = data_emi.Recordset("importe_deuda")
               Ximpgrav = Ximpgrav / 1.1
               Ximpgrav = Format(Ximpgrav, "Standard")
               Ximpgrastr = Format(Ximpgrav, "Standard")
               If Len(Trim(Ximpgrastr)) = 5 Then
                  Ximpgrastr2 = "     " & Mid(Trim(Ximpgrastr), 1, 2)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 4, 2)
               End If
               If Len(Trim(Ximpgrastr)) = 6 Then
                  Ximpgrastr2 = "    " & Mid(Trim(Ximpgrastr), 1, 3)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 5, 2)
               End If
               If Len(Trim(Ximpgrastr)) = 8 Then
                  Ximpgrastr2 = "   " & Mid(Trim(Ximpgrastr), 1, 1)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 3, 3)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 7, 2)
               End If
               If Len(Trim(Ximpgrastr)) = 9 Then
                  Ximpgrastr2 = "  " & Mid(Trim(Ximpgrastr), 1, 2)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 4, 3)
                  Ximpgrastr2 = Ximpgrastr2 & Mid(Trim(Ximpgrastr), 8, 2)
               End If
               Xlin = Xlin + Ximpgrastr2 & "        0"
               Print #1, Xlin
               data_inf.Recordset.AddNew
               data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
               data_inf.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
               data_inf.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
               data_inf.Recordset("cl_hon_pes") = XImp
               data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
               data_inf.Recordset("cl_nrotarj") = data_cli.Recordset("cl_nrotarj")
               data_inf.Recordset.Update
               
               data_emi.Recordset.MoveNext
            Loop
         End If
      Else
         MsgBox "Error en una cédula:" & data_cli.Recordset("cl_codigo")
         End
      End If
      data_cli.Recordset.MoveNext
      Ximparqueo = 0
   Loop
   Close #1
   data_inf.RecordSource = "select * from infcli"
   data_inf.Refresh
   Xtitulo = "DEBITOS OCA CORRESPONDIENTES MES : " & txt_mes.Text & "/" & txt_ano.Text
   cr1.ReportFileName = App.path & "\infdebitos.rpt"
   cr1.ReportTitle = Xtitulo
   cr1.Action = 1
    
   MsgBox "Proceso terminado", vbInformation, "Mensaje"

End If

End Sub
