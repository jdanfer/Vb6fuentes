VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_pasamem 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pasar datos al memory"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   Icon            =   "frm_pasamem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6945
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data_lin 
      Height          =   375
      Left            =   1080
      Top             =   120
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
      Caption         =   "data_lin"
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
   Begin MSAdodcLib.Adodc data_caja 
      Height          =   375
      Left            =   3840
      Top             =   2520
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
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
      Caption         =   "data_caja"
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
   Begin VB.Data ctradm 
      Caption         =   "ctradm"
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
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   615
      Left            =   3000
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   4440
      TabIndex        =   11
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Data data_ca2 
      Caption         =   "data_ca2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Data data_cont 
      Caption         =   "data_cont"
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
      Top             =   3120
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6120
      Picture         =   "frm_pasamem.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Salir"
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Picture         =   "frm_pasamem.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Aceptar"
      Top             =   3360
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Datos a procesar"
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
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6375
      Begin MSAdodcLib.Adodc data_cnv 
         Height          =   375
         Left            =   3240
         Top             =   240
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
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
      Begin MSAdodcLib.Adodc data1 
         Height          =   375
         Left            =   480
         Top             =   840
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3413
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
      Begin MSAdodcLib.Adodc data_cab 
         Height          =   375
         Left            =   360
         Top             =   240
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
         Caption         =   "data_cab"
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
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Ventas Crédito"
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
         Left            =   240
         TabIndex        =   10
         Top             =   2520
         Width           =   2655
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FF0000&
         Caption         =   "Cajas tesorería"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2040
         Width           =   2655
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
         Height          =   375
         Left            =   5280
         TabIndex        =   6
         Text            =   "0"
         Top             =   1680
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FF0000&
         Caption         =   "Seleccionar Base"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3480
         TabIndex        =   5
         Top             =   1440
         Width           =   2655
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FF0000&
         Caption         =   "Todas las bases"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   2655
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   4200
         TabIndex        =   3
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
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
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "FECHAS:"
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
         TabIndex        =   1
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      Height          =   1935
      Left            =   240
      Picture         =   "frm_pasamem.frx":0F56
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   2055
   End
End
Attribute VB_Name = "frm_pasamem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim Xlibro As String
Dim Arch As String
Dim XImp As Double
Dim XIVA, Xtimbre As Double
Dim Xusu, Xtexobs As String
Dim Ctacaja As String
Dim mes, ano, dia, xbase, xnumrub As Long
Dim Xdeb, Xhab, Xqm, Xqa As String
Dim Xbander As Integer
Dim Xelstrin As String

Xbander = 0

If Text1.Text = 101 Or Text1.Text = 102 Then
   Command4_Click
Else
    frm_pasamem.MousePointer = 11
    If Check1.Value = 1 Then
       data_caja.ConnectionString = "dsn=" & Xconexrmt
       data_cont.Connect = "odbc;dsn=" & Xconexrmt & ";"
       data_cont.RecordSource = "rubteso"
       data_cont.Refresh
    Else
       data_caja.ConnectionString = "dsn=" & Xconexrmt
       data_cont.DatabaseName = App.path & "\doccont.mdb"
       data_cont.RecordSource = "doc_cont"
       data_cont.Refresh
    End If
'    data_lin.RecordSource = "linmmdd"
'    data_lin.Refresh
    Arch = "IM"
    mes = Month(md.Text)
    ano = Year(md.Text)
    If mes < 10 Then
       Arch = Arch + Mid(Trim(str(ano)), 3, 2) + Trim("0") + Trim(str(mes)) + "01.txt"
       Xqm = Trim("0") + Trim(str(mes))
       Xqa = Mid(Trim(str(ano)), 3, 2)
    Else
       Arch = Arch + Mid(Trim(str(ano)), 3, 2) + Trim(str(mes)) + "01.txt"
       Xqm = Trim(str(mes))
       Xqa = Mid(Trim(str(ano)), 3, 2)
    End If
    
    If Check1.Value = 1 Then
        If md.Text <> "__/__/____" Then
           If mh.Text <> "__/__/____" Then
              If Option1.Value = True Then
                 data_caja.RecordSource = "Select * from tesorero where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' order by usuario,fecha,cod_rub"
                 data_caja.Refresh
              Else
                 data_caja.RecordSource = "Select * from tesorero where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' order by usuario,fecha,cod_rub"
                 data_caja.Refresh
              End If
'Shell("C:\Program Files\LibreOffice 4\program\scalc.exe
              If Dir("Z:\Memory\Conty\VARTESO" & "\" & Trim(Arch)) <> "" Then
                 Kill ("Z:\Memory\Conty\VARTESO" & "\" & Trim(Arch))
              End If
              Open "Z:\Memory\Conty\VARTESO" & "\" & Trim(Arch) For Output As #1
              If data_caja.Recordset.RecordCount > 0 Then
                 data_caja.Recordset.MoveFirst
                 Xusu = data_caja.Recordset("usuario")
                 dia = Day(data_caja.Recordset("fecha"))
                 xnumrub = data_caja.Recordset("cod_rub")
                 Print #1, "Dia, Debe, Haber,Concepto, Moneda,  Total, CodigoIVA, IVA, Cotizacion, Libro"
                 Do While Not data_caja.Recordset.EOF
                    If Xusu = data_caja.Recordset("usuario") And dia = Day(data_caja.Recordset("fecha")) Then
                       If data_caja.Recordset("moneda") = 2 Then
                          XImp = XImp + data_caja.Recordset("saldou")
                       Else
                          XImp = XImp + data_caja.Recordset("monto")
                       End If
                       xnumrub = data_caja.Recordset("cod_rub")
                       Xusu = data_caja.Recordset("usuario")
                       dia = Day(data_caja.Recordset("fecha"))
                       If data_caja.Recordset("concep") = "E" Then
                          Xlibro = "I"
                       Else
                          Xlibro = "E"
                       End If
                       If dia < 10 Then
'                          Print #1, " " + Trim(Str(dia))
                       Else
'                          Print #1, Trim(Str(dia))
                       End If
                       Xelstrin = Trim(str(dia)) & ","
'                       Print #1, Trim(Str(data_caja.Recordset("cod_debe")))
                       Xelstrin = Xelstrin & Trim(str(data_caja.Recordset("cod_debe"))) & ","
'                       Print #1, Trim(Str(data_caja.Recordset("cod_haber")))
                       Xelstrin = Xelstrin & Trim(str(data_caja.Recordset("cod_haber"))) & ","
'                       Print #1, Mid(Trim(data_caja.Recordset("obs")), 1, 30)
                       Xelstrin = Xelstrin & Mid(Trim(data_caja.Recordset("obs")), 1, 30) & ","
                       If data_caja.Recordset("moneda") = 2 Then
                          'Print #1, "1"
                          Xelstrin = Xelstrin & "1,"
                       Else
                          'Print #1, "0"
                          Xelstrin = Xelstrin & "0,"
                       End If
                       Xelstrin = Xelstrin & Format(XImp, "######0.00") & ","
                       XIVA = data_caja.Recordset("iva")
'                       Print #1, Trim(Str(XIVA))
                       Xelstrin = Xelstrin & Trim(str(XIVA)) & ","
                       
                       XIVA = data_caja.Recordset("impiva")
                       Xelstrin = Xelstrin & Format(XIVA, "######0.00") & ","
                          Xelstrin = Xelstrin & "0.000" & ","
'''                       End If
                       Xelstrin = Xelstrin & Trim(Xlibro)
'                       Print #1, Trim(Xlibro)
    '                   data_caja.Recordset.MoveNext
                       Print #1, Xelstrin
                       xnumrub = data_caja.Recordset("cod_rub")
                       Xusu = data_caja.Recordset("usuario")
                       dia = Day(data_caja.Recordset("fecha"))
                       XImp = 0
                       XIVA = 0
                       data_caja.Recordset.MoveNext
                    Else
                       xnumrub = data_caja.Recordset("cod_rub")
                       Xusu = data_caja.Recordset("usuario")
                       dia = Day(data_caja.Recordset("fecha"))
                       XImp = 0
                       XIVA = 0
                    End If
                 Loop
              Else
                 MsgBox "No existen registros para procesar", vbInformation, "Mensaje"
              End If
           End If
        End If
        Close #1
        MsgBox "Proceso terminado. El archivo quedó guardado en VARTESO del MEMORY del disco C", vbInformation, "Mensaje"
    Else
        If Check2.Value = 1 Then 'crédito
           Command3_Click
        Else
            If md.Text <> "__/__/____" Then
               If mh.Text <> "__/__/____" Then
                  If Option1.Value = True Then
                     data_caja.RecordSource = "Select * from caja where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and base <=" & 100 & " order by base,fecha,usuario,numero"
                     data_caja.Refresh
                  Else
                     data_caja.RecordSource = "Select * from caja where base =" & Text1.Text & " And fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' order by base,fecha,usuario,numero"
                     data_caja.Refresh
                  End If
                  If Dir("C:\Cajas Memory" & "\" & Trim(Arch)) <> "" Then
                     Kill ("C:\Cajas Memory" & "\" & Trim(Arch))
                  End If
                  Open "C:\Cajas Memory" & "\" & Trim(Arch) For Output As #1
'                  data1.RecordSource = "clientes"
'                  data1.Refresh
                  Print #1, "Dia, Debe, Haber,Concepto, Moneda,  Total, CodigoIVA, IVA, Cotizacion, Libro"
                  If data_caja.Recordset.RecordCount > 0 Then
                     data_ca2.DatabaseName = App.path & "\caja.mdb"
                     data_ca2.RecordSource = "caja"
                     data_ca2.Refresh
                     If data_ca2.Recordset.RecordCount > 0 Then
                        data_ca2.Recordset.MoveFirst
                        Do While Not data_ca2.Recordset.EOF
                           data_ca2.Recordset.Delete
                           data_ca2.Recordset.MoveNext
                        Loop
                     End If
                     data_caja.Recordset.MoveFirst
                     Do While Not data_caja.Recordset.EOF
                        data_ca2.Recordset.AddNew
                        data_ca2.Recordset("fecha") = data_caja.Recordset("fecha")
                        data_ca2.Recordset("hora") = data_caja.Recordset("hora")
                        data_ca2.Recordset("base") = data_caja.Recordset("base")
                        data_ca2.Recordset("numero") = data_caja.Recordset("numero")
                        data_ca2.Recordset("moneda") = data_caja.Recordset("moneda")
                        data_ca2.Recordset("nombre") = data_caja.Recordset("nombre")
                        data_ca2.Recordset("movimiento") = data_caja.Recordset("movimiento")
                        data_ca2.Recordset("imp_fact") = data_caja.Recordset("imp_fact")
                        data_ca2.Recordset("imp_iva") = data_caja.Recordset("imp_iva")
                        data_ca2.Recordset("opiva") = data_caja.Recordset("opiva")
                        data_ca2.Recordset("documento") = data_caja.Recordset("documento")
                        data_ca2.Recordset("observ") = data_caja.Recordset("observ")
                        data_ca2.Recordset("saldo") = data_caja.Recordset("saldo")
                        data_ca2.Recordset("usuario") = data_caja.Recordset("usuario")
                        data_ca2.Recordset("saldo_user") = data_caja.Recordset("saldo_user")
                        data_ca2.Recordset("cod_serv") = data_caja.Recordset("cod_serv")
                        data_ca2.Recordset.Update
                        data_caja.Recordset.MoveNext
                     Loop
                      If data_ca2.Recordset.RecordCount > 0 Then
                         data_ca2.Recordset.MoveFirst
                         Do While Not data_ca2.Recordset.EOF
                            If data_ca2.Recordset("cod_serv") = 80011 Then
                               data_ca2.Recordset.Edit
                               data_ca2.Recordset("numero") = 1140525
                               data_ca2.Recordset("nombre") = "Reserva Sicología"
                               data_ca2.Recordset("opiva") = 0
                               data_ca2.Recordset("imp_iva") = 0
                               data_ca2.Recordset.Update
                            End If
                            If data_ca2.Recordset("cod_serv") = 80012 Then
                               data_ca2.Recordset.Edit
                               data_ca2.Recordset("numero") = 1140522
                               data_ca2.Recordset("nombre") = "Reserva Nutricionista"
                               data_ca2.Recordset("opiva") = 0
                               data_ca2.Recordset("imp_iva") = 0
                               data_ca2.Recordset.Update
                            End If
                            If data_ca2.Recordset("cod_serv") = 80013 Then
                               data_ca2.Recordset.Edit
                               data_ca2.Recordset("numero") = 1140521
                               data_ca2.Recordset("nombre") = "Reserva Odontología"
                               data_ca2.Recordset("opiva") = 0
                               data_ca2.Recordset("imp_iva") = 0
                               data_ca2.Recordset.Update
                            End If
                            If data_ca2.Recordset("cod_serv") = 80014 Then
                               data_ca2.Recordset.Edit
                               data_ca2.Recordset("numero") = 1140524
                               data_ca2.Recordset("nombre") = "Reserva Carné de Salud"
                               data_ca2.Recordset("opiva") = 0
                               data_ca2.Recordset("imp_iva") = 0
                               data_ca2.Recordset.Update
                            End If
                            If data_ca2.Recordset("cod_serv") = 80015 Then
                               data_ca2.Recordset.Edit
                               data_ca2.Recordset("numero") = 1140523
                               data_ca2.Recordset("nombre") = "Reserva Fisioterapia"
                               data_ca2.Recordset("opiva") = 0
                               data_ca2.Recordset("imp_iva") = 0
                               data_ca2.Recordset.Update
                            End If
                            If data_ca2.Recordset("cod_serv") = 80016 Then
                               data_ca2.Recordset.Edit
                               data_ca2.Recordset("numero") = 1140526
                               data_ca2.Recordset("nombre") = "Reserva Psiquiatría"
                               data_ca2.Recordset("opiva") = 0
                               data_ca2.Recordset("imp_iva") = 0
                               data_ca2.Recordset.Update
                            End If
                            If data_ca2.Recordset("cod_serv") = 995 Then
                               If data_ca2.Recordset("numero") = 211587 Then 'RX
                                  data_ca2.Recordset.Edit
                                  data_ca2.Recordset("numero") = 211587
                                  data_ca2.Recordset("nombre") = "TIMBRE RX"
                                  data_ca2.Recordset("opiva") = 0
                                  data_ca2.Recordset("imp_iva") = 0
                                  data_ca2.Recordset.Update
                               Else
                                  If data_ca2.Recordset("numero") = 211332 Then 'FERTILAB
                                     data_ca2.Recordset.Edit
                                     data_ca2.Recordset("numero") = 211332
                                     data_ca2.Recordset("nombre") = "TIMBRE FERTILAB"
                                     data_ca2.Recordset("opiva") = 0
                                     data_ca2.Recordset("imp_iva") = 0
                                     data_ca2.Recordset.Update
                                  Else
                                     data_ca2.Recordset.Edit
                                     data_ca2.Recordset("numero") = 213076
                                     data_ca2.Recordset("nombre") = "OBL.TIMBRE PROF"
                                     data_ca2.Recordset("opiva") = 0
                                     data_ca2.Recordset("imp_iva") = 0
                                     data_ca2.Recordset.Update
                                  End If
                               End If
                            Else
                               If data_ca2.Recordset("cod_serv") = 993 Or _
                                  data_ca2.Recordset("cod_serv") = 994 Then
                                  data_ca2.Recordset.Edit
                                  data_ca2.Recordset("numero") = 112022
                                  data_ca2.Recordset("nombre") = "COBRANZA BASE"
                                  data_ca2.Recordset("opiva") = 0
                                  data_ca2.Recordset("imp_iva") = 0
                                  data_ca2.Recordset.Update
                               Else
                                  If data_ca2.Recordset("cod_serv") = 60108 Then
                                     data_ca2.Recordset.Edit
                                     data_ca2.Recordset("numero") = 211302
                                     data_ca2.Recordset("nombre") = "M.UNIVERSAL"
                                     data_ca2.Recordset("opiva") = 0
                                     data_ca2.Recordset("imp_iva") = 0
                                     data_ca2.Recordset.Update
                                  Else
                                     If data_ca2.Recordset("cod_serv") = 60106 Then
                                        data_ca2.Recordset.Edit
                                        data_ca2.Recordset("numero") = 211397
                                        data_ca2.Recordset("nombre") = "M.SMI"
                                        data_ca2.Recordset("opiva") = 0
                                        data_ca2.Recordset("imp_iva") = 0
                                        data_ca2.Recordset.Update
                                     Else
'                                        If data_ca2.Recordset("cod_serv") = 60103 Then
'                                           data_ca2.Recordset.Edit
'                                           data_ca2.Recordset("numero") = 211358
'                                           data_ca2.Recordset("nombre") = "M.CCOU"
'                                           data_ca2.Recordset("opiva") = 0
'                                           data_ca2.Recordset("imp_iva") = 0
'                                           data_ca2.Recordset.Update
'                                        Else
                                           If data_ca2.Recordset("cod_serv") = 992 Then
                                              data_ca2.Recordset.Edit
                                              data_ca2.Recordset("numero") = 513007
                                              data_ca2.Recordset("nombre") = "AFILIACIONES"
                                              data_ca2.Recordset("opiva") = 0
                                              data_ca2.Recordset("imp_iva") = 0
                                              data_ca2.Recordset.Update
                                           Else
                                              If data_ca2.Recordset("cod_serv") = 60109 Then
                                                 data_ca2.Recordset.Edit
                                                 data_ca2.Recordset("numero") = 211372
                                                 data_ca2.Recordset("nombre") = "M.CGALICIA"
                                                 data_ca2.Recordset("opiva") = 0
                                                 data_ca2.Recordset("imp_iva") = 0
                                                 data_ca2.Recordset.Update
                                              End If
                                           End If
                                        'End If
                                     End If
                                  End If
                               End If
                            End If
                            If data_ca2.Recordset("cod_serv") = 996 Then
                               data_ca2.Recordset.Edit
                               data_ca2.Recordset("numero") = 211473
                               data_ca2.Recordset("nombre") = "PAGO A CUENTA"
                               data_ca2.Recordset("opiva") = 0
                               data_ca2.Recordset("imp_iva") = 0
                               data_ca2.Recordset.Update
                            End If
                            
                            data_ca2.Recordset.MoveNext
                         Loop
                         data_ca2.Recordset.MoveFirst
                         xbase = data_ca2.Recordset("base")
                         dia = Day(data_ca2.Recordset("fecha"))
                         xnumrub = data_ca2.Recordset("numero")
                         Xusu = data_ca2.Recordset("usuario")
                         If IsNull(data_ca2.Recordset("observ")) = False Then
                            Xtexobs = Mid(data_ca2.Recordset("observ"), 1, 10)
                         Else
                            Xtexobs = "SIN DATOS"
                         End If
                         Do While Not data_ca2.Recordset.EOF
                            If data_ca2.Recordset("numero") = 111111 Or _
                               data_ca2.Recordset("numero") = 111132 Or _
                               data_ca2.Recordset("numero") = 111133 Or _
                               data_ca2.Recordset("numero") = 111112 Or _
                               data_ca2.Recordset("numero") = 111113 Or _
                               data_ca2.Recordset("numero") = 111114 Or _
                               data_ca2.Recordset("numero") = 111116 Or _
                               data_ca2.Recordset("numero") = 111118 Or _
                               data_ca2.Recordset("numero") = 111191 Or _
                               data_ca2.Recordset("numero") = 111110 Or data_ca2.Recordset("numero") = 111200 Or _
                               data_ca2.Recordset("numero") = 111117 Or data_ca2.Recordset("numero") = 111115 Or _
                               data_ca2.Recordset("numero") = 111136 Or data_ca2.Recordset("numero") = 111198 Or _
                               data_ca2.Recordset("numero") = 111125 Or data_ca2.Recordset("numero") = 111202 Or _
                               data_ca2.Recordset("numero") = 111192 Or _
                               data_ca2.Recordset("numero") = 111188 Or data_ca2.Recordset("numero") = 111199 Then
                               
                               If Val(XImp) <> 0 Then
                                  data_cont.Recordset.FindFirst "cta_caja =" & xnumrub & " And base ='" & Trim(str(xbase)) & "'"
                                
                                  If Not data_cont.Recordset.NoMatch Then
                                     Xlibro = data_cont.Recordset("libro")
                                     Xelstrin = Trim(str(dia)) & ","
                                     Xelstrin = Xelstrin & data_cont.Recordset("ctadebe") & ","
                                     Xelstrin = Xelstrin & data_cont.Recordset("ctahaber") & ","
                                     If xnumrub = 114002 Or _
                                        xnumrub = 513007 Then
                                        Xelstrin = Xelstrin & Mid(Trim(data_cont.Recordset("tipo")), 1, 15) & ","
                                     Else
                                        If xnumrub = 513012 Or _
                                           xnumrub = 422008 Then
                                           Xelstrin = Xelstrin & Mid(Trim(data_cont.Recordset("tipo")), 1, 20) + " " + Mid(Trim(Xusu), 1, 9) & ","
                                        Else
                                           Xelstrin = Xelstrin & Mid(Trim(data_cont.Recordset("tipo")), 1, 30) & ","
                                        End If
                                     End If
                                     Xelstrin = Xelstrin & "0,"
                                     Xelstrin = Xelstrin & Format(XImp, "######0.00") & ","
                                   
                                   If xnumrub = 422034 Or _
                                      xnumrub = 421013 Or _
                                      xnumrub = 211362 Or _
                                      xnumrub = 422008 Or _
                                      xnumrub = 111005 Then
                                      XIVA = 0
                                      Xelstrin = Xelstrin & "0," & "0.00," & "0.000," & Trim(Xlibro)
'                                      Print #1, Xelstrin

                                   Else
'sobrante, provisorios,
                                      If xnumrub = 513012 Or _
                                         xnumrub = 213041 Or _
                                         xnumrub = 211408 Or _
                                         xnumrub = 211423 Or _
                                         xnumrub = 211421 Or _
                                         xnumrub = 211428 Or _
                                         xnumrub = 213076 Or _
                                         xnumrub = 114003 Or xnumrub = 114006 Or _
                                         xnumrub = 114002 Or xnumrub = 211372 Or _
                                         xnumrub = 111001 Or xnumrub = 211473 Or _
                                         xnumrub = 211302 Or _
                                         xnumrub = 211397 Or xnumrub = 211358 Then
                                         XIVA = 0
                                         Xelstrin = Xelstrin & "0," & "0.00," & "0.000," & Trim(Xlibro)
                                         Print #1, Xelstrin
                                      Else
                                          If xnumrub = 422005 Or _
                                             xnumrub = 422004 Or _
                                             xnumrub = 422007 Then
                                             Xelstrin = Xelstrin & "2,"
                                             XIVA = XImp / 1.22
                                             XIVA = XIVA * 0.22
                                             Xelstrin = Xelstrin & Format(XIVA, "######0.00") & ","
                                             Xelstrin = Xelstrin & "0.000," & Trim(Xlibro)
                                             Print #1, Xelstrin
                                          Else
                                              Xelstrin = Xelstrin & "3,"
                                              XIVA = XImp / 1.1
                                              XIVA = XIVA * 0.1
                                              Xelstrin = Xelstrin & Format(XIVA, "######0.00") & ","
                                              Xelstrin = Xelstrin & "0.000,"
                                              Xelstrin = Xelstrin & Trim(Xlibro)
                                              Print #1, Xelstrin
                                          End If
                                      End If
                                   End If
                                End If
                               End If
                               data_ca2.Recordset.MoveNext
                               If data_ca2.Recordset.EOF = False Then
                                  xnumrub = data_ca2.Recordset("numero")
                                  xbase = data_ca2.Recordset("base")
                                  dia = Day(data_ca2.Recordset("fecha"))
                                  Xusu = Mid(data_ca2.Recordset("usuario"), 1, 10)
                                  If IsNull(data_ca2.Recordset("observ")) = False Then
                                     Xtexobs = Mid(data_ca2.Recordset("observ"), 1, 10)
                                  Else
                                     Xtexobs = "SIN DATOS"
                                  End If
                               End If
                               XImp = 0
                               XIVA = 0
                            Else
                                If xbase = data_ca2.Recordset("base") And dia = Day(data_ca2.Recordset("fecha")) And Xusu = data_ca2.Recordset("usuario") Then
                                   If xnumrub = data_ca2.Recordset("numero") Then
                                      XImp = XImp + data_ca2.Recordset("imp_fact")
                                      If xbase = 1 Then
                                         Ctacaja = 5100101
                                      End If
                                      If xbase = 2 Then
                                         Ctacaja = 5100201
                                      End If
                                      If xbase = 3 Then
                                         Ctacaja = 5100301
                                      End If
                                      If xbase = 4 Then
                                         Ctacaja = 5100401
                                      End If
                                      If xbase = 6 Then
                                         Ctacaja = 5100601
                                      End If
                                      If xbase = 8 Then
                                         Ctacaja = 5100801
                                      End If
                                      If xbase = 9 Then
                                         Ctacaja = 5100901
                                      End If
                                      If xbase = 10 Then
                                         Ctacaja = 5101001
                                      End If
                                      If xbase = 11 Then
                                         Ctacaja = 5101101
                                      End If
                                      If xbase = 12 Then
                                         Ctacaja = 5101201
                                      End If
                                      If xbase = 13 Then
                                         Ctacaja = 5101301
                                      End If
                                      If xbase = 15 Then
                                         Ctacaja = 5101501
                                      End If
                                      If xbase = 16 Then
                                         Ctacaja = 5101601
                                      End If
                                      If xbase = 91 Then
                                         Ctacaja = 5101601
                                      End If
                                      If xbase = 17 Then
                                         Ctacaja = 5101701
                                      End If
                                      If xbase = 18 Then
                                         Ctacaja = 5101801
                                      End If
                                      If xbase = 92 Then
                                         Ctacaja = 5101801
                                      End If
                                      If xbase = 93 Then
                                         Ctacaja = 5101701
                                      End If
                                      data_ca2.Recordset.MoveNext 'final del arch
                                      If data_ca2.Recordset.EOF = True Then
                                      
                                        data_cont.Recordset.FindFirst "cta_caja =" & xnumrub & " And base ='" & Trim(str(xbase)) & "'"
                                        If Not data_cont.Recordset.NoMatch Then
                                           Xlibro = data_cont.Recordset("libro")
                                           Xelstrin = Trim(str(dia)) & ","
                                           If xnumrub = 213076 Then
                                              Xelstrin = Xelstrin & data_cont.Recordset("ctadebe") & ","
                                           Else
                                              Xelstrin = Xelstrin & data_cont.Recordset("ctadebe") & ","
                                           End If
                                           Xelstrin = Xelstrin & data_cont.Recordset("ctahaber") & ","
                                           If xnumrub = 114002 Or _
                                              xnumrub = 513007 Then
                                              Xelstrin = Xelstrin & Mid(Trim(data_cont.Recordset("tipo")), 1, 15) & ","
                                           Else
                                              If xnumrub = 513012 Or _
                                                 xnumrub = 422008 Then
                                                 Xelstrin = Xelstrin & Mid(Trim(data_cont.Recordset("tipo")), 1, 20) + " " + Mid(Trim(Xusu), 1, 9) & ","
                                              Else
                                                 Xelstrin = Xelstrin & Mid(Trim(data_cont.Recordset("tipo")), 1, 30) & ","
                                              End If
                                           End If
                                           Xelstrin = Xelstrin & "0,"
                                           Xelstrin = Xelstrin & Format(XImp, "######0.00") & ","
                                           If xnumrub = 422034 Or _
                                              xnumrub = 421013 Or _
                                              xnumrub = 112022 Or _
                                              xnumrub = 211362 Or _
                                              xnumrub = 112001 Or _
                                              xnumrub = 112004 Or _
                                              xnumrub = 112006 Or _
                                              xnumrub = 112011 Or _
                                              xnumrub = 112021 Or _
                                              xnumrub = 112042 Or _
                                              xnumrub = 112113 Or _
                                              xnumrub = 112110 Or _
                                              xnumrub = 422008 Or _
                                              xnumrub = 112044 Or _
                                              xnumrub = 112117 Or _
                                              xnumrub = 111005 Then
                                              XIVA = 0
                                              Xelstrin = Xelstrin & "0," & "0.00," & "0.000," & Trim(Xlibro)
                                              Print #1, Xelstrin
                                           Else
                                              If xnumrub = 513012 Or xnumrub = 113012 Or _
                                                 xnumrub = 213041 Or xnumrub = 113013 Or _
                                                 xnumrub = 211408 Or xnumrub = 113014 Or _
                                                 xnumrub = 211423 Or xnumrub = 113025 Or _
                                                 xnumrub = 211421 Or xnumrub = 113370 Or _
                                                 xnumrub = 211428 Or xnumrub = 113369 Or _
                                                 xnumrub = 213076 Or xnumrub = 113038 Or _
                                                 xnumrub = 114003 Or xnumrub = 114006 Or _
                                                 xnumrub = 114002 Or _
                                                 xnumrub = 111001 Or xnumrub = 211473 Or _
                                                 xnumrub = 113003 Or xnumrub = 211372 Or _
                                                 xnumrub = 211302 Or _
                                                 xnumrub = 211397 Or xnumrub = 211358 Then
                                                 XIVA = 0
                                                  Xelstrin = Xelstrin & "0," & "0.00," & "0.000," & Trim(Xlibro)
                                                 Print #1, Xelstrin
                                              Else
                                                  If xnumrub = 422005 Or _
                                                     xnumrub = 422004 Or _
                                                     xnumrub = 422007 Then
                                                     Xelstrin = Xelstrin & "2,"
                                                     XIVA = XImp / 1.22
                                                     XIVA = XIVA * 0.22
                                                     Xelstrin = Xelstrin & Format(XIVA, "######0.00") & ","
                                                     Xelstrin = Xelstrin & "0.000," & Trim(Xlibro)
                                                     Print #1, Xelstrin
                                                  Else
                                                      Xelstrin = Xelstrin & "3,"
                                                      XIVA = XImp / 1.1
                                                      XIVA = XIVA * 0.1
                                                      Xelstrin = Xelstrin & Format(XIVA, "######0.00") & ","
                                                      Xelstrin = Xelstrin & "0.000,"
                                                      Xelstrin = Xelstrin & Trim(Xlibro)
                                                      Print #1, Xelstrin
                                                  End If
                                              End If
                                           End If
                                        Else
                                           MsgBox "Atención: NO se encuentra RUBRO: " & xnumrub & " BASE: " & xbase, vbInformation, "Mensaje"
                                           End
                                        End If
                                      Else
                                        
                                      End If
                                   Else
                                      If xbase = 1 Then
                                         Ctacaja = 5100101
                                      End If
                                      If xbase = 2 Then
                                         Ctacaja = 5100201
                                      End If
                                      If xbase = 3 Then
                                         Ctacaja = 5100301
                                      End If
                                      If xbase = 4 Then
                                         Ctacaja = 5100401
                                      End If
                                      If xbase = 6 Then
                                         Ctacaja = 5100601
                                      End If
                                      If xbase = 8 Then
                                         Ctacaja = 5100801
                                      End If
                                      If xbase = 9 Then
                                         Ctacaja = 5100901
                                      End If
                                      If xbase = 10 Then
                                         Ctacaja = 5101001
                                      End If
                                      If xbase = 11 Then
                                         Ctacaja = 5101101
                                      End If
                                      If xbase = 12 Then
                                         Ctacaja = 5101201
                                      End If
                                      If xbase = 13 Then
                                         Ctacaja = 5101301
                                      End If
                                      If xbase = 15 Then
                                         Ctacaja = 5101501
                                      End If
                                      If xbase = 16 Then
                                         Ctacaja = 5101601
                                      End If
                                      If xbase = 91 Then
                                         Ctacaja = 5101601
                                      End If
                                      If xbase = 17 Then
                                         Ctacaja = 5101701
                                      End If
                                      If xbase = 18 Then
                                         Ctacaja = 5101801
                                      End If
                                      If xbase = 92 Then
                                         Ctacaja = 5101801
                                      End If
                                      If xbase = 93 Then
                                         Ctacaja = 5101701
                                      End If
                                      
                                      data_cont.Recordset.FindFirst "cta_caja =" & xnumrub & " And base ='" & Trim(str(xbase)) & "'"
                                      If Not data_cont.Recordset.NoMatch Then
                                         Xlibro = data_cont.Recordset("libro")
                                         Xelstrin = Trim(str(dia)) & ","
                                         If xnumrub = 213076 Then
                                            Xelstrin = Xelstrin & data_cont.Recordset("ctadebe") & ","
                                         Else
                                            Xelstrin = Xelstrin & data_cont.Recordset("ctadebe") & ","
                                         End If
                                         Xelstrin = Xelstrin & data_cont.Recordset("ctahaber") & ","
                                         If xnumrub = 114002 Or _
                                            xnumrub = 513007 Then
                                            Xelstrin = Xelstrin & Mid(Trim(data_cont.Recordset("tipo")), 1, 15) & ","
                                         Else
                                            If xnumrub = 513012 Or _
                                               xnumrub = 422008 Then
                                               Xelstrin = Xelstrin & Mid(Trim(data_cont.Recordset("tipo")), 1, 20) + " " + Mid(Trim(Xusu), 1, 9) & ","
                                            Else
                                               Xelstrin = Xelstrin & Mid(Trim(data_cont.Recordset("tipo")), 1, 30) & ","
                                            End If
                                         End If
                                         Xelstrin = Xelstrin & "0,"
                                         Xelstrin = Xelstrin & Format(XImp, "######0.00") & ","
                                         If xnumrub = 422034 Or _
                                            xnumrub = 421013 Or _
                                            xnumrub = 112022 Or _
                                            xnumrub = 211362 Or _
                                            xnumrub = 112001 Or _
                                            xnumrub = 112004 Or _
                                            xnumrub = 112006 Or _
                                            xnumrub = 112011 Or _
                                            xnumrub = 112021 Or _
                                            xnumrub = 112042 Or _
                                            xnumrub = 112113 Or _
                                            xnumrub = 112110 Or _
                                            xnumrub = 422008 Or _
                                            xnumrub = 112044 Or _
                                            xnumrub = 112117 Or _
                                            xnumrub = 111005 Then
                                            XIVA = 0
                                            Xelstrin = Xelstrin & "0," & "0.00," & "0.000," & Trim(Xlibro)
                                            Print #1, Xelstrin
                                         Else
                                            If xnumrub = 513012 Or xnumrub = 113370 Or _
                                               xnumrub = 213041 Or xnumrub = 113369 Or _
                                               xnumrub = 211408 Or xnumrub = 113038 Or _
                                               xnumrub = 211423 Or xnumrub = 113025 Or _
                                               xnumrub = 211421 Or xnumrub = 113014 Or _
                                               xnumrub = 211428 Or xnumrub = 113013 Or _
                                               xnumrub = 213076 Or xnumrub = 113012 Or _
                                               xnumrub = 114003 Or xnumrub = 114006 Or _
                                               xnumrub = 114002 Or _
                                               xnumrub = 111001 Or xnumrub = 211473 Or _
                                               xnumrub = 113003 Or xnumrub = 211372 Or _
                                               xnumrub = 211302 Or _
                                               xnumrub = 211397 Or xnumrub = 211358 Then
                                               XIVA = 0
                                                Xelstrin = Xelstrin & "0," & "0.00," & "0.000," & Trim(Xlibro)
                                               Print #1, Xelstrin
                                            Else
                                                If xnumrub = 422005 Or _
                                                   xnumrub = 422004 Or _
                                                   xnumrub = 422007 Then
                                                   Xelstrin = Xelstrin & "2,"
                                                   XIVA = XImp / 1.22
                                                   XIVA = XIVA * 0.22
                                                   Xelstrin = Xelstrin & Format(XIVA, "######0.00") & ","
                                                   Xelstrin = Xelstrin & "0.000," & Trim(Xlibro)
                                                   Print #1, Xelstrin
                                                Else
                                                    Xelstrin = Xelstrin & "3,"
                                                    XIVA = XImp / 1.1
                                                    XIVA = XIVA * 0.1
                                                    Xelstrin = Xelstrin & Format(XIVA, "######0.00") & ","
                                                    Xelstrin = Xelstrin & "0.000,"
                                                    Xelstrin = Xelstrin & Trim(Xlibro)
                                                    If XImp = 0 Then
                                                    Else
                                                       Print #1, Xelstrin
                                                    End If
                                                End If
                                            End If
                                         End If
                                      Else
                                         MsgBox "Atención: NO se encuentra RUBRO: " & xnumrub & " BASE: " & xbase, vbInformation, "Mensaje"
                                         End
                                      End If
                                      
                                      xnumrub = data_ca2.Recordset("numero")
                                      xbase = data_ca2.Recordset("base")
                                      dia = Day(data_ca2.Recordset("fecha"))
                                      Xusu = data_ca2.Recordset("usuario")
                                      If IsNull(data_ca2.Recordset("observ")) = False Then
                                         Xtexobs = Mid(data_ca2.Recordset("observ"), 1, 10)
                                      Else
                                         Xtexobs = "SIN DATOS"
                                      End If
                                      XImp = 0
                                      XIVA = 0
                                   End If
                                Else
                                    If xbase = 1 Then
                                       Ctacaja = 5100101
                                    End If
                                    If xbase = 2 Then
                                       Ctacaja = 5100201
                                    End If
                                    If xbase = 3 Then
                                       Ctacaja = 5100301
                                    End If
                                    If xbase = 4 Then
                                       Ctacaja = 5100401
                                    End If
                                    If xbase = 6 Then
                                       Ctacaja = 5100601
                                    End If
                                    If xbase = 8 Then
                                       Ctacaja = 5100801
                                    End If
                                    If xbase = 9 Then
                                       Ctacaja = 5100901
                                    End If
                                    If xbase = 10 Then
                                       Ctacaja = 5101001
                                    End If
                                    If xbase = 11 Then
                                       Ctacaja = 5101101
                                    End If
                                    If xbase = 12 Then
                                       Ctacaja = 5101201
                                    End If
                                    If xbase = 13 Then
                                       Ctacaja = 5101301
                                    End If
                                    If xbase = 15 Then
                                       Ctacaja = 5101501
                                    End If
                                    If xbase = 16 Then
                                       Ctacaja = 5101601
                                    End If
                                    If xbase = 91 Then
                                       Ctacaja = 5101601
                                    End If
                                    If xbase = 17 Then
                                       Ctacaja = 5101701
                                    End If
                                    If xbase = 18 Then
                                       Ctacaja = 5101801
                                    End If
                                    If xbase = 92 Then
                                       Ctacaja = 5101801
                                    End If
                                    If xbase = 93 Then
                                       Ctacaja = 5101701
                                    End If
                                   
                                   data_cont.Recordset.FindFirst "cta_caja =" & xnumrub & " And base ='" & Trim(str(xbase)) & "'"
                                   If Not data_cont.Recordset.NoMatch Then
                                      Xlibro = data_cont.Recordset("libro")
                                      Xelstrin = Trim(str(dia)) & ","
                                      If xnumrub = 213076 Then
                                         Xelstrin = Xelstrin & data_cont.Recordset("ctadebe") & ","
                                      Else
                                         Xelstrin = Xelstrin & data_cont.Recordset("ctadebe") & ","
                                      End If
                                      Xelstrin = Xelstrin & data_cont.Recordset("ctahaber") & ","
                                      If xnumrub = 114002 Or _
                                         xnumrub = 513007 Then
                                         Xelstrin = Xelstrin & Mid(Trim(data_cont.Recordset("tipo")), 1, 15) & ","
                                      Else
                                         If xnumrub = 513012 Or _
                                            xnumrub = 422008 Then
                                            Xelstrin = Xelstrin & Mid(Trim(data_cont.Recordset("tipo")), 1, 20) + " " + Mid(Trim(Xusu), 1, 9) & ","
                                         Else
                                            Xelstrin = Xelstrin & Mid(Trim(data_cont.Recordset("tipo")), 1, 30) & ","
                                         End If
                                      End If
                                      Xelstrin = Xelstrin & "0,"
                                      Xelstrin = Xelstrin & Format(XImp, "######0.00") & ","
                                      If xnumrub = 422034 Or _
                                         xnumrub = 421013 Or _
                                         xnumrub = 112022 Or _
                                         xnumrub = 211362 Or _
                                         xnumrub = 112001 Or _
                                         xnumrub = 112004 Or _
                                         xnumrub = 112006 Or _
                                         xnumrub = 112011 Or _
                                         xnumrub = 112021 Or _
                                         xnumrub = 112042 Or _
                                         xnumrub = 112113 Or _
                                         xnumrub = 112110 Or _
                                         xnumrub = 422008 Or _
                                         xnumrub = 112044 Or _
                                         xnumrub = 112117 Or _
                                         xnumrub = 111005 Then
                                         XIVA = 0
                                         Xelstrin = Xelstrin & "0," & "0.00," & "0.000," & Trim(Xlibro)
                                         Print #1, Xelstrin
                                      Else
                                         If xnumrub = 513012 Or xnumrub = 113370 Or _
                                            xnumrub = 213041 Or xnumrub = 113369 Or _
                                            xnumrub = 211408 Or xnumrub = 113038 Or _
                                            xnumrub = 211423 Or xnumrub = 113025 Or _
                                            xnumrub = 211421 Or xnumrub = 113014 Or _
                                            xnumrub = 211428 Or xnumrub = 113013 Or _
                                            xnumrub = 213076 Or xnumrub = 113012 Or _
                                            xnumrub = 114003 Or xnumrub = 114006 Or _
                                            xnumrub = 114002 Or _
                                            xnumrub = 111001 Or xnumrub = 211473 Or _
                                            xnumrub = 113003 Or xnumrub = 211372 Or _
                                            xnumrub = 211302 Or _
                                            xnumrub = 211397 Or xnumrub = 211358 Then
                                            XIVA = 0
                                            Xelstrin = Xelstrin & "0," & "0.00," & "0.000," & Trim(Xlibro)
                                            Print #1, Xelstrin
                                         Else
                                            If xnumrub = 422005 Or _
                                               xnumrub = 422004 Or _
                                               xnumrub = 422007 Then
                                               XIVA = XImp / 1.22
                                               XIVA = XIVA * 0.22
                                               Xelstrin = Xelstrin & "2," & Format(XIVA, "######0.00") & ","
                                               Xelstrin = Xelstrin & "0.000," & Trim(Xlibro)
                                               Print #1, Xelstrin
                                            Else
                                                Xelstrin = Xelstrin & "3,"
                                                XIVA = XImp / 1.1
                                                XIVA = XIVA * 0.1
                                                Xelstrin = Xelstrin & Format(XIVA, "######0.00") & ","
                                                 Xelstrin = Xelstrin & "0.000," & Trim(Xlibro)
                                                 If XImp = 0 Then
                                                 Else
                                                    Print #1, Xelstrin
                                                 End If
                                            End If
                                         End If
                                      End If
                                   Else
                                      MsgBox "Atención: NO se encuentra RUBRO: " & xnumrub & " BASE: " & xbase, vbInformation, "Mensaje"
                                      End
                                   End If
                '                data_caja.Recordset.MoveNext
                                    xnumrub = data_ca2.Recordset("numero")
                                    xbase = data_ca2.Recordset("base")
                                    dia = Day(data_ca2.Recordset("fecha"))
                                    Xusu = data_ca2.Recordset("usuario")
                                    If IsNull(data_ca2.Recordset("observ")) = False Then
                                       Xtexobs = Mid(data_ca2.Recordset("observ"), 1, 10)
                                    Else
                                       Xtexobs = "SIN DATOS"
                                    End If
                                    XImp = 0
                                    XIVA = 0
                                End If
                            End If
                         Loop
                      Else
                         MsgBox "No existen registros con esta fecha"
                      End If
                  End If
               End If
            End If
            Close #1
'            If WElusuario = "CDEMORAES" Or WElusuario = "PAOLA" Or WElusuario = "FOSORIO" Or WElusuario = "GBOTTA" Then
'               Shell ("TXT2CFWU.EXE " & Arch & " x:\memory\conty\cajas " & Xqm & " " & Xqa), vbMaximizedFocus
'            Else
'               Shell ("TXT2CFWU.EXE " & Arch & " c:\memory\conty\cajas " & Xqm & " " & Xqa), vbMaximizedFocus
'            End If
            MsgBox "Proceso terminado. El archivo quedó guardado en CAJAS MEMORY del disco C", vbInformation, "Mensaje"
        End If
    End If
End If
frm_pasamem.MousePointer = 0

End Sub

Private Sub Command2_Click()
Dim M As Long
'Shell ("c:\txt2cfwu.exe " & Arch & " c:\wincont\cajas\" & Arch), vbNormalFocus
Unload Me

End Sub

Private Sub Command3_Click()
Dim Arch1 As String
Dim mes2, ano2 As Long
Dim Xelstrin2 As String

Arch1 = "IM"
mes2 = Month(md.Text)
ano2 = Year(md.Text)
If mes2 < 10 Then
   Arch1 = Arch1 + Mid(Trim(str(ano2)), 3, 2) + Trim("0") + Trim(str(mes2)) + "01.txt"
   Xqm = Trim("0") + Trim(str(mes2))
   Xqa = Mid(Trim(str(ano2)), 3, 2)

Else
   Arch1 = Arch1 + Mid(Trim(str(ano2)), 3, 2) + Trim(str(mes2)) + "01.txt"
   Xqm = Trim(str(mes2))
   Xqa = Mid(Trim(str(ano2)), 3, 2)
End If
        If md.Text <> "__/__/____" Then
           If mh.Text <> "__/__/____" Then
              data_caja.RecordSource = "Select * from linmmdd where base <=" & 100 & " and fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' And fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' And tipo ='" & "CREDITO" & "' order by base,fecha"
              data_caja.Refresh
              If Dir("C:\Cajas Memory" & "\" & Trim(Arch1)) <> "" Then
                 Kill ("C:\Cajas Memory" & "\" & Trim(Arch1))
              End If
              Open "C:\Cajas Memory" & "\" & Trim(Arch1) For Output As #1
'              data1.RecordSource = "clientes"
'              data1.Refresh
              If data_caja.Recordset.RecordCount > 0 Then
                 Print #1, "Dia, Debe, Haber,Concepto, Moneda,  Total, CodigoIVA, IVA, Cotizacion, Libro"
                 data_ca2.DatabaseName = App.path & "\caja.mdb"
                 data_ca2.RecordSource = "caja"
                 data_ca2.Refresh
                 If data_ca2.Recordset.RecordCount > 0 Then
                    data_ca2.Recordset.MoveFirst
                    Do While Not data_ca2.Recordset.EOF
                       data_ca2.Recordset.Delete
                       data_ca2.Recordset.MoveNext
                    Loop
                 End If
                 data_caja.Recordset.MoveFirst
                 Do While Not data_caja.Recordset.EOF
                    data_ca2.Recordset.AddNew
                    data_ca2.Recordset("fecha") = data_caja.Recordset("fecha")
                    data_ca2.Recordset("hora") = data_caja.Recordset("hora")
                    data_ca2.Recordset("base") = data_caja.Recordset("base")
                    data_ca2.Recordset("numero") = data_caja.Recordset("rub_cont")
                    data_ca2.Recordset("moneda") = 1
                    data_ca2.Recordset("nombre") = data_caja.Recordset("rub_nomb")
                    data_ca2.Recordset("movimiento") = "INGRESO"
                    If IsNull(data_caja.Recordset("pendiente")) = False Then
                       If data_caja.Recordset("pendiente") = "N" Or data_caja.Recordset("pendiente") = "C" Or data_caja.Recordset("pendiente") = "R" Then
                          data_ca2.Recordset("imp_fact") = data_caja.Recordset("tot_lin") * -1
                          data_ca2.Recordset("imp_iva") = data_caja.Recordset("imp_iva") * -1
                       Else
                          data_ca2.Recordset("imp_fact") = data_caja.Recordset("tot_lin")
                          data_ca2.Recordset("imp_iva") = data_caja.Recordset("imp_iva")
                       End If
                    Else
                       data_cab.RecordSource = "Select * from clirespl where cl_numero =" & data_caja.Recordset("factura") & " and cl_codigo =" & data_caja.Recordset("cod_cli")
                       data_cab.Refresh
                       If data_cab.Recordset.RecordCount > 0 Then
                          If IsNull(data_cab.Recordset("cl_telefon")) = False Then
                             If data_cab.Recordset("cl_telefon") = "NC E-TICKET" Or data_cab.Recordset("cl_telefon") = "NC E-FACTURA" Then
                                data_ca2.Recordset("imp_fact") = data_caja.Recordset("tot_lin") * -1
                                data_ca2.Recordset("imp_iva") = data_caja.Recordset("imp_iva") * -1
                             Else
                                data_ca2.Recordset("imp_fact") = data_caja.Recordset("tot_lin")
                                data_ca2.Recordset("imp_iva") = data_caja.Recordset("imp_iva")
                             End If
                          Else
                             data_ca2.Recordset("imp_fact") = data_caja.Recordset("tot_lin")
                             data_ca2.Recordset("imp_iva") = data_caja.Recordset("imp_iva")
                          End If
                       Else
                          data_ca2.Recordset("imp_fact") = data_caja.Recordset("tot_lin")
                          data_ca2.Recordset("imp_iva") = data_caja.Recordset("imp_iva")
                       End If
                    End If
'                    data_ca2.Recordset("imp_fact") = data_caja.Recordset("tot_lin")
'                    data_ca2.Recordset("imp_iva") = data_caja.Recordset("imp_iva")
                    data_ca2.Recordset("opiva") = 1
                    data_ca2.Recordset("documento") = data_caja.Recordset("factura")
                    data_ca2.Recordset("observ") = data_caja.Recordset("nom_prod")
                    data_ca2.Recordset("saldo") = 0
                    data_ca2.Recordset("usuario") = data_caja.Recordset("operador")
                    data_ca2.Recordset("saldo_user") = 0
                    data_ca2.Recordset.Update
                    data_caja.Recordset.MoveNext
                 Loop
                  If data_ca2.Recordset.RecordCount > 0 Then
                     data_ca2.Recordset.MoveFirst
                     xbase = data_ca2.Recordset("base")
                     dia = Day(data_ca2.Recordset("fecha"))
                     xnumrub = data_ca2.Recordset("numero")
                     Xusu = Mid(data_ca2.Recordset("usuario"), 1, 10)
                     If IsNull(data_ca2.Recordset("observ")) = False Then
                        Xtexobs = Mid(data_ca2.Recordset("observ"), 1, 10)
                     Else
                        Xtexobs = "SIN DATOS"
                     End If
                     Do While Not data_ca2.Recordset.EOF
                        If data_ca2.Recordset("numero") = 111111 Or _
                           data_ca2.Recordset("numero") = 111132 Or _
                           data_ca2.Recordset("numero") = 111133 Or _
                           data_ca2.Recordset("numero") = 111112 Or _
                           data_ca2.Recordset("numero") = 111113 Or _
                           data_ca2.Recordset("numero") = 111114 Or _
                           data_ca2.Recordset("numero") = 111116 Or _
                           data_ca2.Recordset("numero") = 111118 Or _
                           data_ca2.Recordset("numero") = 111191 Or data_ca2.Recordset("numero") = 111200 Or _
                           data_ca2.Recordset("numero") = 111110 Or data_ca2.Recordset("numero") = 111115 Or _
                           data_ca2.Recordset("numero") = 111117 Or data_ca2.Recordset("numero") = 111198 Or _
                           data_ca2.Recordset("numero") = 111136 Or data_ca2.Recordset("numero") = 111202 Or _
                           data_ca2.Recordset("numero") = 111125 Or _
                           data_ca2.Recordset("numero") = 111192 Or _
                           data_ca2.Recordset("numero") = 111188 Or data_ca2.Recordset("numero") = 111199 Then
                           If Val(XImp) <> 0 Then
                            data_cont.Recordset.FindFirst "cta_caja =" & xnumrub & " And base ='" & Trim(str(xbase)) & "'"
                            If Not data_cont.Recordset.NoMatch Then
                               Xlibro = data_cont.Recordset("libro")
'                               If dia < 10 Then
'                                  Print #1, " " + Trim(Str(dia))
'                               Else
'                                  Print #1, Trim(Str(dia))
'                               End If
                               Xelstrin2 = Trim(str(dia)) & ","
                               If xnumrub = 213076 Then
'                                  Print #1, data_cont.Recordset("ctadebe")
                                  Xelstrin2 = Xelstrin2 & data_cont.Recordset("ctadebe") & ","
                               Else
'                                  Print #1, data_cont.Recordset("ctadebe")
                                  Xelstrin2 = Xelstrin2 & data_cont.Recordset("ctadebe") & ","
                               End If
'                               Print #1, data_cont.Recordset("ctahaber")
                               Xelstrin2 = Xelstrin2 & data_cont.Recordset("ctahaber") & ","
                               If xnumrub = 114002 Or _
                                  xnumrub = 513007 Then
'                                  Print #1, Mid(Trim(data_cont.Recordset("tipo")), 1, 15)
                                  Xelstrin2 = Xelstrin2 & Mid(Trim(data_cont.Recordset("tipo")), 1, 15) & ","
                               Else
                                  If xnumrub = 513012 Or _
                                     xnumrub = 422008 Then
'                                     Print #1, Mid(Trim(data_cont.Recordset("tipo")), 1, 20) + " " + Mid(Trim(Xusu), 1, 9)
                                     Xelstrin2 = Xelstrin2 & Mid(Trim(data_cont.Recordset("tipo")), 1, 20) + " " + Mid(Trim(Xusu), 1, 9) & ","
                                  Else
'                                     Print #1, Mid(Trim(data_cont.Recordset("tipo")), 1, 30)
                                     Xelstrin2 = Xelstrin2 & Mid(Trim(data_cont.Recordset("tipo")), 1, 30) & ","
                                  End If
                               End If
'                               Print #1, "0"
                               Xelstrin2 = Xelstrin2 & "0" & ","
                               Xelstrin2 = Xelstrin2 & Format(XImp, "######0.00") & ","
                               If xnumrub = 422034 Or _
                                  xnumrub = 421013 Or _
                                  xnumrub = 112022 Or _
                                  xnumrub = 211362 Or _
                                  xnumrub = 112001 Or _
                                  xnumrub = 112004 Or _
                                  xnumrub = 112006 Or _
                                  xnumrub = 112011 Or _
                                  xnumrub = 112021 Or _
                                  xnumrub = 112042 Or _
                                  xnumrub = 112113 Or _
                                  xnumrub = 112110 Or _
                                  xnumrub = 422008 Or _
                                  xnumrub = 112044 Or _
                                  xnumrub = 113003 Or xnumrub = 112117 Or _
                                  xnumrub = 111005 Then
                                  XIVA = 0
'                                  Print #1, "0"
'                                  Print #1, "      0.00"
'                                  Print #1, "  0.000000"
'                                  Print #1, Trim(Xlibro)
                                  Xelstrin2 = Xelstrin2 & "0," & "0.00," & "0.000," & Trim(Xlibro)
                                  Print #1, Xelstrin2
                               Else
                                  If xnumrub = 513012 Or _
                                     xnumrub = 213041 Or _
                                     xnumrub = 211408 Or _
                                     xnumrub = 211423 Or _
                                     xnumrub = 211421 Or _
                                     xnumrub = 211428 Or _
                                     xnumrub = 213076 Or _
                                     xnumrub = 114003 Or _
                                     xnumrub = 114002 Or _
                                     xnumrub = 111001 Or _
                                     xnumrub = 211302 Or _
                                     xnumrub = 211397 Or xnumrub = 211358 Then
                                     XIVA = 0
'                                     Print #1, "0"
'                                     Print #1, "      0.00"
'                                     Print #1, "  0.000000"
'                                     Print #1, Trim(Xlibro)
                                     Xelstrin2 = Xelstrin2 & "0," & "0.00," & "0.000," & Trim(Xlibro)
                                     Print #1, Xelstrin2
                                  Else
                                      If xnumrub = 422005 Or _
                                         xnumrub = 422004 Or _
                                         xnumrub = 422007 Then
                                         'Print #1, "2"
                                         Xelstrin2 = Xelstrin2 & "2,"
                                         XIVA = XImp / 1.22
                                         XIVA = XIVA * 0.22
                                         Xelstrin2 = Xelstrin2 & Format(XIVA, "######0.00") & ","
'                                         Print #1, "  0.000000"
                                         Xelstrin2 = Xelstrin2 & "0.000,"
                                         Xelstrin2 = Xelstrin2 & Trim(Xlibro)
                                         Print #1, Xelstrin2
'                                         Print #1, Trim(Xlibro)
                                      Else
                                          'Print #1, "3"
                                          Xelstrin2 = Xelstrin2 & "3,"
                                          XIVA = XImp / 1.1
                                          XIVA = XIVA * 0.1
                                          Xelstrin2 = Xelstrin2 & Format(XIVA, "######0.00") & ","
                                          Xelstrin2 = Xelstrin2 & "0.000,"
                                          Xelstrin2 = Xelstrin2 & Trim(Xlibro)
                                          Print #1, Xelstrin2
'                                          Print #1, Trim(Xlibro)
                                      End If
                                  End If
                               End If
                            End If
                           End If
                           data_ca2.Recordset.MoveNext
                           If data_ca2.Recordset.EOF = False Then
                              xnumrub = data_ca2.Recordset("numero")
                              xbase = data_ca2.Recordset("base")
                              dia = Day(data_ca2.Recordset("fecha"))
                              Xusu = Mid(data_ca2.Recordset("usuario"), 1, 10)
                              If IsNull(data_ca2.Recordset("observ")) = False Then
                                 Xtexobs = Mid(data_ca2.Recordset("observ"), 1, 10)
                              Else
                                 Xtexobs = "SIN DATOS"
                              End If
                           End If
                           XImp = 0
                           XIVA = 0
                        Else
                            If xbase = data_ca2.Recordset("base") And dia = Day(data_ca2.Recordset("fecha")) Then
                               If xnumrub = data_ca2.Recordset("numero") Then
                                  XImp = XImp + data_ca2.Recordset("imp_fact")
                                  If xbase = 1 Then
                                     Ctacaja = 5100101
                                  End If
                                  If xbase = 2 Then
                                     Ctacaja = 5100201
                                  End If
                                  If xbase = 3 Then
                                     Ctacaja = 5100301
                                  End If
                                  If xbase = 4 Then
                                     Ctacaja = 5100401
                                  End If
                                  If xbase = 6 Then
                                     Ctacaja = 5100601
                                  End If
                                  If xbase = 8 Then
                                     Ctacaja = 5100801
                                  End If
                                  If xbase = 9 Then
                                     Ctacaja = 5100901
                                  End If
                                  If xbase = 10 Then
                                     Ctacaja = 5101001
                                  End If
                                  If xbase = 11 Then
                                     Ctacaja = 5101101
                                  End If
                                  If xbase = 12 Then
                                     Ctacaja = 5101201
                                  End If
                                  If xbase = 13 Then
                                     Ctacaja = 5101301
                                  End If
                                  If xbase = 15 Then
                                     Ctacaja = 5101501
                                  End If
                                  If xbase = 16 Then
                                     Ctacaja = 5101601
                                  End If
                                  If xbase = 91 Then
                                     Ctacaja = 5101601
                                  End If
                                  If xbase = 17 Then
                                     Ctacaja = 5101701
                                  End If
                                  If xbase = 18 Then
                                     Ctacaja = 5101801
                                  End If
                                  If xbase = 92 Then
                                     Ctacaja = 5101801
                                  End If
                                  If xbase = 93 Then
                                     Ctacaja = 5101701
                                  End If
                                  data_ca2.Recordset.MoveNext
                               Else
                                  If xbase = 1 Then
                                     Ctacaja = 5100101
                                  End If
                                  If xbase = 2 Then
                                     Ctacaja = 5100201
                                  End If
                                  If xbase = 3 Then
                                     Ctacaja = 5100301
                                  End If
                                  If xbase = 4 Then
                                     Ctacaja = 5100401
                                  End If
                                  If xbase = 6 Then
                                     Ctacaja = 5100601
                                  End If
                                  If xbase = 8 Then
                                     Ctacaja = 5100801
                                  End If
                                  If xbase = 9 Then
                                     Ctacaja = 5100901
                                  End If
                                  If xbase = 10 Then
                                     Ctacaja = 5101001
                                  End If
                                  If xbase = 11 Then
                                     Ctacaja = 5101101
                                  End If
                                  If xbase = 12 Then
                                     Ctacaja = 5101201
                                  End If
                                  If xbase = 13 Then
                                     Ctacaja = 5101301
                                  End If
                                  If xbase = 15 Then
                                     Ctacaja = 5101501
                                  End If
                                  If xbase = 16 Then
                                     Ctacaja = 5101601
                                  End If
                                  If xbase = 91 Then
                                     Ctacaja = 5101601
                                  End If
                                  If xbase = 17 Then
                                     Ctacaja = 5101701
                                  End If
                                  If xbase = 18 Then
                                     Ctacaja = 5101801
                                  End If
                                  If xbase = 92 Then
                                     Ctacaja = 5101801
                                  End If
                                  If xbase = 93 Then
                                     Ctacaja = 5101701
                                  End If
                                  data_cont.Recordset.FindFirst "cta_caja =" & xnumrub & " And base ='" & Trim(str(xbase)) & "'"
                                  If Not data_cont.Recordset.NoMatch Then
                                     Xlibro = data_cont.Recordset("libro")
                                     If dia < 10 Then
                                        'Print #1, " " + Trim(Str(dia))
                                        Xelstrin2 = Trim(str(dia)) & ","
                                     Else
                                        'Print #1, Trim(Str(dia))
                                        Xelstrin2 = Trim(str(dia)) & ","
                                     End If
                                     If xnumrub = 213076 Then
                                        'Print #1, data_cont.Recordset("ctadebe")
                                        Xelstrin2 = Xelstrin2 & data_cont.Recordset("ctadebe") & ","
                                     Else
                                        'Print #1, data_cont.Recordset("ctadebe")
                                        Xelstrin2 = Xelstrin2 & data_cont.Recordset("ctadebe") & ","
                                     End If
                                     'Print #1, data_cont.Recordset("ctahaber")
                                     Xelstrin2 = Xelstrin2 & data_cont.Recordset("ctahaber") & ","
                                     If xnumrub = 114002 Or _
                                        xnumrub = 513007 Then
                                        'Print #1, Mid(Trim(data_cont.Recordset("tipo")), 1, 15)
                                        Xelstrin2 = Xelstrin2 & Mid(Trim(data_cont.Recordset("tipo")), 1, 15) & ","
                                     Else
                                        If xnumrub = 513012 Or _
                                           xnumrub = 422008 Then
                                           'Print #1, Mid(Trim(data_cont.Recordset("tipo")), 1, 20) + " " + Mid(Trim(Xusu), 1, 9)
                                           Xelstrin2 = Xelstrin2 & Mid(Trim(data_cont.Recordset("tipo")), 1, 20) + " " + Mid(Trim(Xusu), 1, 9) & ","
                                        Else
                                           'Print #1, Mid(Trim(data_cont.Recordset("tipo")), 1, 30)
                                           Xelstrin2 = Xelstrin2 & Mid(Trim(data_cont.Recordset("tipo")), 1, 30) & ","
                                        End If
                                     End If
                                     'Print #1, "0"
                                     Xelstrin2 = Xelstrin2 & "0,"
                                     Xelstrin2 = Xelstrin2 & Format(XImp, "######0.00") & ","
                                     'If Len(Trim(Str(Int(XImp)))) >= 7 Then
                                     '   Print #1, Format(XImp, "######0.00")
                                     'End If
                                     If xnumrub = 422034 Or _
                                        xnumrub = 421013 Or _
                                        xnumrub = 112022 Or _
                                        xnumrub = 211362 Or _
                                        xnumrub = 112001 Or _
                                        xnumrub = 112004 Or _
                                        xnumrub = 112006 Or _
                                        xnumrub = 112011 Or _
                                        xnumrub = 112021 Or _
                                        xnumrub = 112042 Or _
                                        xnumrub = 112113 Or _
                                        xnumrub = 112110 Or _
                                        xnumrub = 422008 Or _
                                        xnumrub = 112044 Or xnumrub = 112117 Or _
                                        xnumrub = 111005 Then
                                        XIVA = 0
                                        'Print #1, "0"
                                        'Print #1, "      0.00"
                                        'Print #1, "  0.000000"
                                        'Print #1, Trim(Xlibro)
                                        Xelstrin2 = Xelstrin2 & "0," & "0.00," & "0.000," & Trim(Xlibro)
                                        Print #1, Xelstrin2
                                     Else
                                        If xnumrub = 513012 Or _
                                           xnumrub = 213041 Or _
                                           xnumrub = 211408 Or _
                                           xnumrub = 211423 Or _
                                           xnumrub = 211421 Or _
                                           xnumrub = 211428 Or _
                                           xnumrub = 213076 Or _
                                           xnumrub = 114003 Or _
                                           xnumrub = 114002 Or _
                                           xnumrub = 111001 Or _
                                           xnumrub = 113003 Or _
                                           xnumrub = 211302 Or _
                                           xnumrub = 211397 Or xnumrub = 211358 Then
                                           XIVA = 0
                                           'Print #1, "0"
                                           'Print #1, "      0.00"
                                           'Print #1, "  0.000000"
                                           'Print #1, Trim(Xlibro)
                                           Xelstrin2 = Xelstrin2 & "0," & "0.00," & "0.000," & Trim(Xlibro)
                                           Print #1, Xelstrin2
                                        Else
                                            If xnumrub = 422005 Or _
                                               xnumrub = 422004 Or _
                                               xnumrub = 422007 Then
                                               'Print #1, "2"
                                               Xelstrin2 = Xelstrin2 & "2" & ","
                                               XIVA = XImp / 1.22
                                               XIVA = XIVA * 0.22
                                               Xelstrin2 = Xelstrin2 & Format(XIVA, "######0.00") & ","
                                               'If Len(Trim(Str(Int(XIVA)))) >= 7 Then
                                               '   Print #1, Format(XIVA, "######0.00")
                                               'End If
                                               'Print #1, "  0.000000"
                                               Xelstrin2 = Xelstrin2 & "0.000" & ","
                                               'Print #1, Trim(Xlibro)
                                               Xelstrin2 = Xelstrin2 & Trim(Xlibro)
                                               Print #1, Xelstrin2
                                            Else
                                                'Print #1, "3"
                                                Xelstrin2 = Xelstrin2 & "3" & ","
                                                XIVA = XImp / 1.1
                                                XIVA = XIVA * 0.1
                                                Xelstrin2 = Xelstrin2 & Format(XIVA, "######0.00") & ","
                                                'If Len(Trim(Str(Int(XIVA)))) >= 7 Then
                                                '   Print #1, Format(XIVA, "######0.00")
                                                'End If
                                                'Print #1, "  0.000000"
                                                Xelstrin2 = Xelstrin2 & "0.00,"
                                                'Print #1, Trim(Xlibro)
                                                Xelstrin2 = Xelstrin2 & Trim(Xlibro)
                                                Print #1, Xelstrin2
                                            End If
                                        End If
                                     End If
                                  Else
                                     MsgBox "Atención: NO se encuentra RUBRO: " & xnumrub & " BASE: " & xbase, vbInformation, "Mensaje"
                                     End
                                  End If
            '                      data_caja.Recordset.MoveNext
                                  xnumrub = data_ca2.Recordset("numero")
                                  xbase = data_ca2.Recordset("base")
                                  dia = Day(data_ca2.Recordset("fecha"))
                                  Xusu = Mid(data_ca2.Recordset("usuario"), 1, 10)
                                  If IsNull(data_ca2.Recordset("observ")) = False Then
                                     Xtexobs = Mid(data_ca2.Recordset("observ"), 1, 10)
                                  Else
                                     Xtexobs = "SIN DATOS"
                                  End If
                                  XImp = 0
                                  XIVA = 0
                               End If
                            Else
                                If xbase = 1 Then
                                   Ctacaja = 5100101
                                End If
                                If xbase = 2 Then
                                   Ctacaja = 5100201
                                End If
                                If xbase = 3 Then
                                   Ctacaja = 5100301
                                End If
                                If xbase = 4 Then
                                   Ctacaja = 5100401
                                End If
                                If xbase = 6 Then
                                   Ctacaja = 5100601
                                End If
                                If xbase = 8 Then
                                   Ctacaja = 5100801
                                End If
                                If xbase = 9 Then
                                   Ctacaja = 5100901
                                End If
                                If xbase = 10 Then
                                   Ctacaja = 5101001
                                End If
                                If xbase = 11 Then
                                   Ctacaja = 5101101
                                End If
                                If xbase = 12 Then
                                   Ctacaja = 5101201
                                End If
                                If xbase = 13 Then
                                   Ctacaja = 5101301
                                End If
                                If xbase = 15 Then
                                   Ctacaja = 5101501
                                End If
                                If xbase = 16 Then
                                   Ctacaja = 5101601
                                End If
                                If xbase = 91 Then
                                   Ctacaja = 5101601
                                End If
                                If xbase = 17 Then
                                   Ctacaja = 5101701
                                End If
                                If xbase = 18 Then
                                   Ctacaja = 5101801
                                End If
                                If xbase = 92 Then
                                   Ctacaja = 5101801
                                End If
                                If xbase = 93 Then
                                   Ctacaja = 5101701
                                End If
                               
                               data_cont.Recordset.FindFirst "cta_caja =" & xnumrub & " And base ='" & Trim(str(xbase)) & "'"
                               If Not data_cont.Recordset.NoMatch Then
                                  Xlibro = data_cont.Recordset("libro")
'                                  If dia < 10 Then
'                                     Print #1, " " + Trim(Str(dia))
'                                  Else
'                                     Print #1, Trim(Str(dia))
'                                  End If
                                  Xelstrin2 = Trim(str(dia)) & ","
                                  If xnumrub = 213076 Then
'                                     Print #1, data_cont.Recordset("ctadebe")
                                     Xelstrin2 = Xelstrin2 & data_cont.Recordset("ctadebe") & ","
                                  Else
'                                     Print #1, data_cont.Recordset("ctadebe")
                                     Xelstrin2 = Xelstrin2 & data_cont.Recordset("ctadebe") & ","
                                  End If
                                  'Print #1, data_cont.Recordset("ctahaber")
                                  Xelstrin2 = Xelstrin2 & data_cont.Recordset("ctahaber") & ","
                                  If xnumrub = 114002 Or _
                                     xnumrub = 513007 Then
                                     'Print #1, Mid(Trim(data_cont.Recordset("tipo")), 1, 15)
                                     Xelstrin2 = Xelstrin2 & Mid(Trim(data_cont.Recordset("tipo")), 1, 15) & ","
                                  Else
                                     If xnumrub = 513012 Or _
                                        xnumrub = 422008 Then
                                        'Print #1, Mid(Trim(data_cont.Recordset("tipo")), 1, 20) + " " + Mid(Trim(Xusu), 1, 9)
                                        Xelstrin2 = Xelstrin2 & Mid(Trim(data_cont.Recordset("tipo")), 1, 20) + " " + Mid(Trim(Xusu), 1, 9) & ","
                                     Else
                                        'Print #1, Mid(Trim(data_cont.Recordset("tipo")), 1, 30)
                                        Xelstrin2 = Xelstrin2 & Mid(Trim(data_cont.Recordset("tipo")), 1, 30) & ","
                                     End If
                                  End If
                                  'Print #1, "0"
                                  Xelstrin2 = Xelstrin2 & "0,"
                                  Xelstrin2 = Xelstrin2 & Format(XImp, "######0.00") & ","
                                  'If Len(Trim(Str(Int(XImp)))) >= 7 Then
                                  '   Print #1, Format(XImp, "######0.00")
                                  'End If
                                  If xnumrub = 422034 Or _
                                     xnumrub = 421013 Or _
                                     xnumrub = 112022 Or _
                                     xnumrub = 211362 Or _
                                     xnumrub = 112001 Or _
                                     xnumrub = 112004 Or _
                                     xnumrub = 112006 Or _
                                     xnumrub = 112011 Or _
                                     xnumrub = 112021 Or _
                                     xnumrub = 112042 Or _
                                     xnumrub = 112113 Or _
                                     xnumrub = 112110 Or _
                                     xnumrub = 422008 Or _
                                     xnumrub = 112044 Or xnumrub = 112117 Or _
                                     xnumrub = 111005 Then
                                     XIVA = 0
                                     'Print #1, "0"
                                     'Print #1, "      0.00"
                                     'Print #1, "  0.000000"
                                     'Print #1, Trim(Xlibro)
                                     Xelstrin2 = Xelstrin2 & "0," & "0.00," & "0.000," & Trim(Xlibro)
                                     Print #1, Xelstrin2
                                  Else
                                     If xnumrub = 513012 Or _
                                        xnumrub = 213041 Or _
                                        xnumrub = 211408 Or _
                                        xnumrub = 211423 Or _
                                        xnumrub = 211421 Or _
                                        xnumrub = 211428 Or _
                                        xnumrub = 213076 Or _
                                        xnumrub = 114003 Or _
                                        xnumrub = 114002 Or _
                                        xnumrub = 111001 Or _
                                        xnumrub = 113003 Or _
                                        xnumrub = 211302 Or _
                                        xnumrub = 211397 Or xnumrub = 211358 Then
                                        XIVA = 0
                                        Xelstrin2 = Xelstrin2 & "0," & "0.00," & "0.000," & Trim(Xlibro)
                                        Print #1, Xelstrin2
                                     Else
                                        If xnumrub = 422005 Or _
                                           xnumrub = 422004 Or _
                                           xnumrub = 422007 Then
                                           'Print #1, "2"
                                           Xelstrin2 = Xelstrin2 & "2" & ","
                                           XIVA = XImp / 1.22
                                           XIVA = XIVA * 0.22
                                           Xelstrin2 = Xelstrin2 & Format(XIVA, "######0.00") & ","
'                                           If Len(Trim(Str(Int(XIVA)))) >= 7 Then
'                                              Print #1, Format(XIVA, "######0.00")
'                                           End If
                                           'Print #1, "  0.000000"
                                           Xelstrin2 = Xelstrin2 & "0.000" & ","
                                           'Print #1, Trim(Xlibro)
                                           Xelstrin2 = Xelstrin2 & Trim(Xlibro)
                                           Print #1, Xelstrin2
                                        Else
                                            'Print #1, "3"
                                            Xelstrin2 = Xelstrin2 & "3,"
                                            XIVA = XImp / 1.1
                                            XIVA = XIVA * 0.1
                                            Xelstrin2 = Xelstrin2 & Format(XIVA, "######0.00") & ","
                                            'If Len(Trim(Str(Int(XIVA)))) >= 7 Then
                                            '   Print #1, Format(XIVA, "######0.00")
                                            'End If
                                            'Print #1, "  0.000000"
                                            Xelstrin2 = Xelstrin2 & "0.000,"
                                            'Print #1, Trim(Xlibro)
                                            Xelstrin2 = Xelstrin2 & Trim(Xlibro)
                                            Print #1, Xelstrin2
                                        End If
                                     End If
                                  End If
                               Else
                                  MsgBox "Atención: NO se encuentra RUBRO: " & xnumrub & " BASE: " & xbase, vbInformation, "Mensaje"
                                  End
                               End If
            '                data_caja.Recordset.MoveNext
                                xnumrub = data_ca2.Recordset("numero")
                                xbase = data_ca2.Recordset("base")
                                dia = Day(data_ca2.Recordset("fecha"))
                                Xusu = Mid(data_ca2.Recordset("usuario"), 1, 10)
                                If IsNull(data_ca2.Recordset("observ")) = False Then
                                   Xtexobs = Mid(data_ca2.Recordset("observ"), 1, 10)
                                Else
                                   Xtexobs = "SIN DATOS"
                                End If
                                XImp = 0
                                XIVA = 0
                            End If
                        End If
                     Loop
                  Else
                     MsgBox "No existen registros con esta fecha"
                  End If
              End If
           End If
        End If
        Close #1
        'If WElusuario = "CDEMORAES" Or WElusuario = "PAOLA" Or WElusuario = "FOSORIO" Then
        '   Shell ("TXT2CFWU.EXE " & Arch1 & " x:\memory\conty\cajas " & Xqm & " " & Xqa), vbMaximizedFocus
        'Else
        '   Shell ("TXT2CFWU.EXE " & Arch1 & " c:\memory\conty\cajas " & Xqm & " " & Xqa), vbMaximizedFocus
        'End If
        MsgBox "Proceso terminado. El archivo quedó guardado en CAJAS MEMORY del disco C", vbInformation, "Mensaje"

End Sub

Private Sub Command4_Click()
Dim Xlibro As String
Dim Arch As String
Dim XImp As Double
Dim XIVA, Xtimbre, XIVA2 As Double
Dim Xusu, Xtexobs As String
Dim Ctacaja As String
Dim mes, ano, dia, xbase, xnumrub, XNfac As Long
Dim Xdeb, Xhab, Xqm, Xqa As String
Dim Xbander As Integer
Dim Xelstrin, Xelrut As String
Xelstrin = ""
Xbander = 0
Xelrut = ""
XNfac = 0
frm_pasamem.MousePointer = 11
'data_caja.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_caja.ConnectionString = "dsn=" & Xconexrmt
data_cont.DatabaseName = App.path & "\doccont.mdb"
data_cont.RecordSource = "doc_cont"
data_cont.Refresh

'data_lin.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_lin.RecordSource = "Select * from linmmdd where fecha >='" & Format(md.Text, "yyyy-mm-dd") & "' and fecha <='" & Format(mh.Text, "yyyy-mm-dd") & "' and base =" & Text1.Text & " order by fecha,factura"
data_lin.Refresh

Arch = "IM"
mes = Month(md.Text)
ano = Year(md.Text)
If mes < 10 Then
   Arch = Arch + Mid(Trim(str(ano)), 3, 2) + Trim("0") + Trim(str(mes)) + "01.txt"
   Xqm = Trim("0") + Trim(str(mes))
   Xqa = Mid(Trim(str(ano)), 3, 2)
Else
   Arch = Arch + Mid(Trim(str(ano)), 3, 2) + Trim(str(mes)) + "01.txt"
   Xqm = Trim(str(mes))
   Xqa = Mid(Trim(str(ano)), 3, 2)
End If

If md.Text <> "__/__/____" Then
   If mh.Text <> "__/__/____" Then
      If Dir("C:\Cajas Memory" & "\" & Trim(Arch)) <> "" Then
         Kill ("C:\Cajas Memory" & "\" & Trim(Arch))
      End If
      Open "C:\Cajas Memory" & "\" & Trim(Arch) For Output As #1
      If data_lin.Recordset.RecordCount > 0 Then
         data_lin.Recordset.MoveFirst
         Xusu = data_lin.Recordset("operador")
         dia = Day(data_lin.Recordset("fecha"))
         xnumrub = data_lin.Recordset("rub_cont")
         XNfac = data_lin.Recordset("factura")
         Print #1, "Dia, Debe, Haber,Concepto,Ruc, Moneda,  Total, CodigoIVA, IVA, Cotizacion, Libro"
         Do While Not data_lin.Recordset.EOF
'            If dia = Day(data_lin.Recordset("fecha")) Then
            If XNfac = data_lin.Recordset("factura") Then
               If IsNull(data_lin.Recordset("pendiente")) = False Then
                  If data_lin.Recordset("pendiente") = "N" Or data_lin.Recordset("pendiente") = "C" Or data_lin.Recordset("pendiente") = "R" Then
                     XImp = XImp - data_lin.Recordset("tot_lin") - data_lin.Recordset("valor_iva")
                     XIVA = XIVA - data_lin.Recordset("valor_iva")
                      'data_lin.Recordset("imp_fact") = data_caja.Recordset("tot_lin") * -1
                      'data_ca2.Recordset("imp_iva") = data_caja.Recordset("imp_iva") * -1
                  Else
                     XImp = XImp + data_lin.Recordset("tot_lin") + data_lin.Recordset("valor_iva")
                     XIVA = XIVA + data_lin.Recordset("valor_iva")
'                      data_ca2.Recordset("imp_fact") = data_caja.Recordset("tot_lin")
'                      data_ca2.Recordset("imp_iva") = data_caja.Recordset("imp_iva")
                  End If
               Else
                  data_cab.RecordSource = "Select * from clirespl where cl_numero =" & data_lin.Recordset("factura") & " and cl_codigo =" & data_lin.Recordset("cod_cli")
                  data_cab.Refresh
                  If data_cab.Recordset.RecordCount > 0 Then
                     If IsNull(data_cab.Recordset("cl_telefon")) = False Then
                        If data_cab.Recordset("cl_telefon") = "NC E-TICKET" Or data_cab.Recordset("cl_telefon") = "NC E-FACTURA" Then
                           XImp = XImp - data_lin.Recordset("tot_lin") - data_lin.Recordset("valor_iva")
                           XIVA = XIVA - data_lin.Recordset("valor_iva")
                        Else
                           XImp = XImp + data_lin.Recordset("tot_lin") + data_lin.Recordset("valor_iva")
                           XIVA = XIVA + data_lin.Recordset("valor_iva")
                        End If
                     Else
                        XImp = XImp + data_lin.Recordset("tot_lin") + data_lin.Recordset("valor_iva")
                        XIVA = XIVA + data_lin.Recordset("valor_iva")
                     End If
                  Else
                     XImp = XImp + data_lin.Recordset("tot_lin") + data_lin.Recordset("valor_iva")
                     XIVA = XIVA + data_lin.Recordset("valor_iva")
                  End If
               End If
               If data_lin.Recordset("costo_prod") = 1 Then
                  If IsNull(data_lin.Recordset("pendiente")) = False Then
                     If data_lin.Recordset("pendiente") = "N" Or data_lin.Recordset("pendiente") = "C" Or data_lin.Recordset("pendiente") = "R" Then
                        XIVA = data_lin.Recordset("tot_lin") * 0.22
                        XImp = XImp - data_lin.Recordset("tot_lin") - XIVA
                        XIVA = XIVA
                     Else
                        XIVA = data_lin.Recordset("tot_lin") * 0.22
                        XImp = data_lin.Recordset("tot_lin") + XIVA
                     End If
                  Else
                     XIVA = data_lin.Recordset("tot_lin") * 0.22
                     XImp = data_lin.Recordset("tot_lin") + XIVA
                  End If
               End If
               xnumrub = data_lin.Recordset("rub_cont")
               Xusu = data_lin.Recordset("operador")
               dia = Day(data_lin.Recordset("fecha"))
               XNfac = data_lin.Recordset("factura")
               data_lin.Recordset.MoveNext
            Else
                data_lin.Recordset.MovePrevious
                If IsNull(data_lin.Recordset("ruc")) = False Then
                   Xelrut = data_lin.Recordset("ruc")
                Else
                   Xelrut = ""
                End If
                xnumrub = data_lin.Recordset("rub_cont")
                Xusu = data_lin.Recordset("operador")
                dia = Day(data_lin.Recordset("fecha"))
                data_cnv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("convenio") & "'"
                data_cnv.Refresh
                If data_cnv.Recordset.RecordCount > 0 Then
                   Xlibro = data_cont.Recordset("libro")
                   If dia < 10 Then
                      Xelstrin = Trim(str(dia)) & ","
                   Else
                      Xelstrin = Trim(str(dia)) & ","
                   End If
                   If IsNull(data_cnv.Recordset("cnv_uapago")) = False Then
                      Xelstrin = Xelstrin & data_cnv.Recordset("cnv_uapago") & ","
                   Else
                      Xelstrin = Xelstrin & "0" & ","
                   End If
                   Xelstrin = Xelstrin & data_lin.Recordset("rub_cont") & ","
                   Xelstrin = Xelstrin & "F." & Trim(str(data_lin.Recordset("factura"))) & " " & data_lin.Recordset("nom_cli") & ","
                   If Len(Trim(Xelrut)) > 2 Then
                      Xelstrin = Xelstrin & Trim(Xelrut) & ","
                   Else
                      Xelstrin = Xelstrin & ","
                   End If
                   Xelstrin = Xelstrin & "0,"
                   Xelstrin = Xelstrin & Format(XImp, "######0.00") & ","
                   If IsNull(data_lin.Recordset("costo_prod")) = False Then
                      XIVA2 = data_lin.Recordset("costo_prod")
                      If XIVA2 = 0 Then
                         XIVA2 = 3
                      Else
                         If XIVA2 = 1 Then
                            XIVA2 = 4
                         Else
                            If XIVA2 = 2 Then
                               XIVA2 = 0
                            Else
                               XIVA2 = 3
                            End If
                         End If
                      End If
                   Else
                       XIVA2 = 3
                   End If
                   Xelstrin = Xelstrin & Trim(str(XIVA2)) & ","
                   Xelstrin = Xelstrin & Format(XIVA, "######0.00") & "," & "0.000" & ","
                   Xelstrin = Xelstrin & Trim(Xlibro)
'                   Print #1, Trim(Xlibro)
                   Print #1, Xelstrin
    '                   data_caja.Recordset.MoveNext
                
                Else
                   MsgBox "No se encontró CONVENIO", vbInformation, "Mensaje"
                   Unload Me
                End If
                XImp = 0
                XIVA = 0
                data_lin.Recordset.MoveNext
                xnumrub = data_lin.Recordset("rub_cont")
                Xusu = data_lin.Recordset("operador")
                dia = Day(data_lin.Recordset("fecha"))
                XNfac = data_lin.Recordset("factura")
            
            End If
         Loop
         data_lin.Recordset.MovePrevious
         If IsNull(data_lin.Recordset("ruc")) = False Then
            Xelrut = data_lin.Recordset("ruc")
         Else
            Xelrut = ""
         End If
         xnumrub = data_lin.Recordset("rub_cont")
         Xusu = data_lin.Recordset("operador")
         dia = Day(data_lin.Recordset("fecha"))
         data_cnv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_lin.Recordset("convenio") & "'"
         data_cnv.Refresh
         If data_cnv.Recordset.RecordCount > 0 Then
            Xlibro = data_cont.Recordset("libro")
            If dia < 10 Then
               Xelstrin = Trim(str(dia)) & ","
            Else
               Xelstrin = Trim(str(dia)) & ","
            End If
            If IsNull(data_cnv.Recordset("cnv_uapago")) = False Then
               Xelstrin = Xelstrin & data_cnv.Recordset("cnv_uapago") & ","
            Else
               Xelstrin = Xelstrin & "0" & ","
            End If
            Xelstrin = Xelstrin & data_lin.Recordset("rub_cont") & ","
            Xelstrin = Xelstrin & "F." & Trim(str(data_lin.Recordset("factura"))) & " " & data_lin.Recordset("nom_cli") & ","
            If Len(Trim(Xelrut)) > 2 Then
               Xelstrin = Xelstrin & Trim(Xelrut) & ","
            Else
               Xelstrin = Xelstrin & ","
            End If
            Xelstrin = Xelstrin & "0,"
            Xelstrin = Xelstrin & Format(XImp, "######0.00") & ","
            If IsNull(data_lin.Recordset("costo_prod")) = False Then
               XIVA2 = data_lin.Recordset("costo_prod")
               If XIVA2 = 0 Then
                  XIVA2 = 3
               Else
                  If XIVA2 = 1 Then
                     XIVA2 = 4
                  Else
                     If XIVA2 = 2 Then
                        XIVA2 = 0
                     Else
                        XIVA2 = 3
                     End If
                  End If
               End If
            Else
                XIVA2 = 3
            End If
            Xelstrin = Xelstrin & Trim(str(XIVA2)) & ","
            Xelstrin = Xelstrin & Format(XIVA, "######0.00") & "," & "0.000" & ","
            Xelstrin = Xelstrin & Trim(Xlibro)
'                   Print #1, Trim(Xlibro)
            Print #1, Xelstrin
         Else
            MsgBox "No se encontró convenio.", vbInformation, "Mensaje"
            Unload Me
         End If
         XImp = 0
         XIVA = 0
         Close #1
         MsgBox "Proceso terminado. El archivo quedó guardado en CAJAS MEMORY del disco C", vbInformation, "Mensaje"
      Else
         MsgBox "No existen registros"
      End If
   End If
End If

End Sub

Private Sub Form_Load()
'data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.ConnectionString = "dsn=" & Xconexrmt
'data_cab.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cnv.ConnectionString = "dsn=" & Xconexrmt
data_cab.ConnectionString = "dsn=" & Xconexrmt
data_lin.ConnectionString = "dsn=" & Xconexrmt
ctradm.DatabaseName = App.path & "\ctradm.mdb"
ctradm.RecordSource = "ctradm"
ctradm.Refresh

End Sub

Private Sub Form_Resize()
With Image1
    .Top = 0
    .Left = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub md_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mh.SetFocus
End If

End Sub

Private Sub Option1_Click()
Text1.Visible = False

End Sub

Private Sub Option2_Click()
Text1.Visible = True

End Sub
