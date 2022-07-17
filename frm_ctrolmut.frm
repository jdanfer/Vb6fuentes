VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_ctrolmut 
   BackColor       =   &H0000C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de socios mutuales"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7380
   Icon            =   "frm_ctrolmut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   7380
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_paramnew 
      Caption         =   "data_paramnew"
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
      Top             =   3120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 8.0;"
      DatabaseName    =   "C:\mutuales\ccou.xls"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3600
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   "socios$"
      Top             =   2640
      Visible         =   0   'False
      Width           =   2775
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   3000
      Visible         =   0   'False
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data data_mut 
      Caption         =   "data_mut"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4440
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2655
   End
   Begin MSAdodcLib.Adodc data_conv 
      Height          =   375
      Left            =   3600
      Top             =   2760
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
      Caption         =   "data_conv"
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
   Begin VB.CheckBox Check1 
      BackColor       =   &H0080FFFF&
      Caption         =   "1727 SOLO CLAVE 1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   2640
      Visible         =   0   'False
      Width           =   3015
   End
   Begin Crystal.CrystalReport cr3 
      Left            =   4080
      Top             =   3840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport cr2 
      Left            =   3840
      Top             =   3360
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
      Height          =   300
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3960
      Visible         =   0   'False
      Width           =   3375
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   3240
      Top             =   3480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
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
      Left            =   4560
      Picture         =   "frm_ctrolmut.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3600
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Procesar..."
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
      Left            =   600
      Picture         =   "frm_ctrolmut.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Datos para control"
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
      Height          =   2295
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   6615
      Begin MSAdodcLib.Adodc data_cli 
         Height          =   375
         Left            =   720
         Top             =   1080
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
         Connect         =   "dsn=sappnew"
         OLEDBString     =   "dsn=sappnew"
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
      Begin MSAdodcLib.Adodc data_lin 
         Height          =   330
         Left            =   3720
         Top             =   240
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
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "frm_ctrolmut.frx":0F56
         Left            =   2280
         List            =   "frm_ctrolmut.frx":0F84
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   4095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00404040&
         Caption         =   "LOS CONTROLES SE REALIZAN MEDIANTE EL NÚMERO DE CÉDULA DE IDENTIDAD."
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
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   6135
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Mutualista:"
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
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cash"
      Height          =   495
      Left            =   5760
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   5760
      TabIndex        =   6
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Procesando Altas/Modif"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2640
      TabIndex        =   10
      Top             =   2640
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   120
      Picture         =   "frm_ctrolmut.frx":100E
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   1095
   End
End
Attribute VB_Name = "frm_ctrolmut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frm_ctrolmut.MousePointer = 11
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

Dim Xobjexel As Excel.Application
Dim Xlibexel As Excel.Workbook
Dim Xarchexel As New Excel.Worksheet

Dim XCol, Xlin, Xnrocan, Xcolfija As Long
Dim Xarchtex As String
Dim Xlabrir As New Excel.Application

On Error GoTo Quepasamut

MiBaseact.Execute "Delete * from infcli"
data_inf.RecordSource = "infcli"
data_inf.Refresh

Set MiBaseact = Unasesact.OpenDatabase(App.path & "\mutuales.mdb")
MiBaseact.Execute "Delete * from infno"
data_infno.DatabaseName = App.path & "\mutuales.mdb"
data_infno.RecordSource = "infno"
data_infno.Refresh

Dim Xlafectex As String
Dim Xcedeva As Long
Dim Xccedeva As Integer

If Combo1.Text = "CASA DE GALICIA" Then 'OK
   data_mut.RecordSource = "cgal"
   data_mut.Refresh
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         data_mut.Recordset.Delete
         data_mut.Recordset.MoveNext
      Loop
   End If
   Data1.DatabaseName = "C:\mutuales\cgal.xls"
   Data1.RecordSource = "socios$"
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         data_mut.Recordset.AddNew
         data_mut.Recordset("ced") = Data1.Recordset("ced")
         data_mut.Recordset("mat") = Data1.Recordset("mat")
         data_mut.Recordset("nom1") = Data1.Recordset("nom1")
         data_mut.Recordset("nom2") = Data1.Recordset("nom2")
         data_mut.Recordset("ape1") = Data1.Recordset("ape1")
         data_mut.Recordset("ape2") = Data1.Recordset("ape2")
         If IsNull(Data1.Recordset("fnac")) = False Then
            data_mut.Recordset("fnac") = Data1.Recordset("fnac")
         End If
         If IsNull(Data1.Recordset("categ")) = False Then
            data_mut.Recordset("categ") = Mid(Data1.Recordset("categ"), 1, 255)
         End If
         If IsNull(Data1.Recordset("domicilio")) = False Then
            data_mut.Recordset("domicilio") = Mid(Data1.Recordset("domicilio"), 1, 255)
         End If
         If IsNull(Data1.Recordset("telefono")) = False Then
            data_mut.Recordset("telefono") = Mid(Data1.Recordset("telefono"), 1, 255)
         End If
         If IsNull(Data1.Recordset("celular")) = False Then
            data_mut.Recordset("celular") = Mid(Data1.Recordset("celular"), 1, 255)
         End If
         If IsNull(Data1.Recordset("correo")) = False Then
            data_mut.Recordset("correo") = Mid(Data1.Recordset("correo"), 1, 255)
         End If
         If IsNull(Data1.Recordset("fecing")) = False Then
            data_mut.Recordset("fecing") = Data1.Recordset("fecing")
         End If
         data_mut.Recordset.Update
         Data1.Recordset.MoveNext
      Loop
   End If
   data_mut.Refresh
   Label3.Visible = True
   Label3.Caption = "Procesando Altas/Modif"
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveLast
      DoEvents
      pb.Visible = True
      pb.Max = data_mut.Recordset.RecordCount
      pb.Value = 0
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         If IsNull(data_mut.Recordset("ced")) = False Then
            Xcedeva = data_mut.Recordset("ced")
            data_mut.Recordset.Edit
            data_mut.Recordset("cednum") = Xcedeva
            data_mut.Recordset.Update
         Else
            Xcedeva = 0
         End If
         If Xcedeva > 0 Then
'            data_cli.Recordset.FindFirst "cl_cedula =" & Xcedeva
            data_cli.RecordSource = "Select * from clientes where cl_cedula =" & Xcedeva
            data_cli.Refresh
            If data_cli.Recordset.RecordCount > 0 Then
               data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
               data_conv.Refresh
               If data_conv.Recordset.RecordCount > 0 Then
                  If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                     If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                        data_inf.Recordset.AddNew
                        If IsNull(data_mut.Recordset("fnac")) = False Then
                           data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                        End If
                        data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced")))
                        If IsNull(data_mut.Recordset("domicilio")) = False Then
                           data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                        End If
                        If IsNull(data_mut.Recordset("categ")) = False Then
                           data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                        End If
                        If IsNull(data_mut.Recordset("celular")) = False Then
                           data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                        End If
                        If IsNull(data_mut.Recordset("telefono")) = False Then
                           data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                        End If
                        If IsNull(data_mut.Recordset("ape2")) = False Then
                           If IsNull(data_mut.Recordset("nom2")) = False Then
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                           Else
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                           End If
                        Else
                           If IsNull(data_mut.Recordset("nom2")) = False Then
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                           Else
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                           End If
                        End If
                        If IsNull(data_mut.Recordset("correo")) = False Then
                           data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                        End If
                        data_inf.Recordset("cl_nombre") = "REACTIVAR"
                        data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                        data_inf.Recordset.Update
                     Else
                        If data_conv.Recordset("cnv_codigo") = "GANOS" Or data_conv.Recordset("cnv_codigo") = "CASANO" Or _
                           data_conv.Recordset("cnv_codigo") = "CASANR" Or data_conv.Recordset("cnv_codigo") = "CASNSA" Then
                            data_inf.Recordset.AddNew
                            If IsNull(data_mut.Recordset("fnac")) = False Then
                               data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                            End If
                            data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced")))
                            If IsNull(data_mut.Recordset("domicilio")) = False Then
                               data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                            End If
                            If IsNull(data_mut.Recordset("categ")) = False Then
                               data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                            End If
                            If IsNull(data_mut.Recordset("celular")) = False Then
                               data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                            End If
                            If IsNull(data_mut.Recordset("telefono")) = False Then
                               data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                            End If
                            If IsNull(data_mut.Recordset("ape2")) = False Then
                               If IsNull(data_mut.Recordset("nom2")) = False Then
                                  data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                               Else
                                  data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                               End If
                            Else
                               If IsNull(data_mut.Recordset("nom2")) = False Then
                                  data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                               Else
                                  data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                               End If
                            End If
                            If IsNull(data_mut.Recordset("correo")) = False Then
                               data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                            End If
                            data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                            data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                            data_inf.Recordset.Update
                        End If
                     End If
                  Else
                     data_inf.Recordset.AddNew
                     If IsNull(data_mut.Recordset("fnac")) = False Then
                        data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                     End If
                     data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced")))
                     If IsNull(data_mut.Recordset("domicilio")) = False Then
                        data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                     End If
                     If IsNull(data_mut.Recordset("categ")) = False Then
                        data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                     End If
                     If IsNull(data_mut.Recordset("celular")) = False Then
                        data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                     End If
                     If IsNull(data_mut.Recordset("telefono")) = False Then
                        data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                     End If
                     If IsNull(data_mut.Recordset("ape2")) = False Then
                        If IsNull(data_mut.Recordset("nom2")) = False Then
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                        Else
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                        End If
                     Else
                        If IsNull(data_mut.Recordset("nom2")) = False Then
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                        Else
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                        End If
                     End If
                     If IsNull(data_mut.Recordset("correo")) = False Then
                        data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                     End If
                     If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                        data_inf.Recordset("cl_nombre") = "CONV.INCORRECTO BAJ"
                     Else
                        data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                     End If
                     data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_inf.Recordset.Update
                  End If
               Else
                  data_inf.Recordset.AddNew
                  If IsNull(data_mut.Recordset("fnac")) = False Then
                     data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                  End If
                  data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced")))
                  If IsNull(data_mut.Recordset("domicilio")) = False Then
                     data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                  End If
                  If IsNull(data_mut.Recordset("categ")) = False Then
                     data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                  End If
                  If IsNull(data_mut.Recordset("celular")) = False Then
                     data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                  End If
                  If IsNull(data_mut.Recordset("telefono")) = False Then
                     data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                  End If
                  If IsNull(data_mut.Recordset("ape2")) = False Then
                     If IsNull(data_mut.Recordset("nom2")) = False Then
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                     Else
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                     End If
                  Else
                     If IsNull(data_mut.Recordset("nom2")) = False Then
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                     Else
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                     End If
                  End If
                  If IsNull(data_mut.Recordset("correo")) = False Then
                     data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                  End If
                  If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                     data_inf.Recordset("cl_nombre") = "CONV.INCORRECTO BAJA"
                  Else
                     data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                  End If
                  data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                  data_inf.Recordset.Update
               End If
            Else
               data_inf.Recordset.AddNew
               If IsNull(data_mut.Recordset("fnac")) = False Then
                  data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
               End If
               data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced")))
               If IsNull(data_mut.Recordset("domicilio")) = False Then
                  data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
               End If
               If IsNull(data_mut.Recordset("categ")) = False Then
                  data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
               End If
               If IsNull(data_mut.Recordset("celular")) = False Then
                  data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
               End If
               If IsNull(data_mut.Recordset("telefono")) = False Then
                  data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
               End If
               If IsNull(data_mut.Recordset("ape2")) = False Then
                  If IsNull(data_mut.Recordset("nom2")) = False Then
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                  Else
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                  End If
               Else
                  If IsNull(data_mut.Recordset("nom2")) = False Then
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                  Else
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                  End If
               End If
               If IsNull(data_mut.Recordset("correo")) = False Then
                  data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
               End If
               data_inf.Recordset("cl_nombre") = "NO ESTA EN P.SAPP"
               data_inf.Recordset.Update
            End If
         End If
         data_mut.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      DoEvents
      data_cli.RecordSource = "Select * from clientes where estado <>" & 2 & " And estado <>" & 3 & " and cl_codconv not in ('UDEMM','PART','SMIN','UNIVS','EMERN','CASH','MSP','CAAMEP','UCM','CCNOS','HEVANO','HEVAN')"
      data_cli.Refresh
      data_cli.Recordset.MoveLast
      data_cli.Recordset.MoveFirst
      pb.Max = pb.Max + data_cli.Recordset.RecordCount
      data_mut.Refresh
      Label3.Caption = "Procesando BAJAS..."
      DoEvents
      Do While Not data_cli.Recordset.EOF
         If IsNull(data_cli.Recordset("cl_codconv")) = False Then
            If IsNull(data_cli.Recordset("cl_cedula")) = False Then
               If data_cli.Recordset("cl_cedula") > 0 Then
                  If data_cli.Recordset("cl_codconv") = "GANOS" Or data_cli.Recordset("cl_codconv") = "CASANO" Or _
                     data_cli.Recordset("cl_codconv") = "CASANR" Or data_cli.Recordset("cl_codconv") = "CASNSA" Then
                  Else
                     data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                     data_conv.Refresh
                     If data_conv.Recordset.RecordCount > 0 Then
                        If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                           data_mut.RecordSource = "Select * from cgal where cednum =" & Int(data_cli.Recordset("cl_cedula"))
                           data_mut.Refresh
                           If data_mut.Recordset.RecordCount > 0 Then
                           Else
                              data_infno.Recordset.AddNew
                              data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                              data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                              data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                              data_infno.Recordset("cl_nombre") = "BAJA"
                              data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                              data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                              data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                              data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                              data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                              data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                              data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                              data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                              data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                              data_infno.Recordset.Update
                           End If
                        End If
                     End If
                  End If
               Else
                  If data_cli.Recordset("cl_codconv") = "GANOS" Or data_cli.Recordset("cl_codconv") = "CASANO" Or _
                     data_cli.Recordset("cl_codconv") = "CASANR" Or data_cli.Recordset("cl_codconv") = "CASNSA" Then
                  Else
                  
                    data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                    data_conv.Refresh
                    If data_conv.Recordset.RecordCount > 0 Then
                       If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                          data_infno.Recordset.AddNew
                          data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                          data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                          data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                          data_infno.Recordset("cl_nombre") = "BAJA"
                          data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                          data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                          data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                          data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                          data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                          data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                          data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                          data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                          data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                          data_infno.Recordset.Update
                       End If
                    End If
                  End If
               End If
            Else
               If data_cli.Recordset("cl_codconv") = "GANOS" Or data_cli.Recordset("cl_codconv") = "CASANO" Or _
                  data_cli.Recordset("cl_codconv") = "CASANR" Or data_cli.Recordset("cl_codconv") = "CASNSA" Then
               Else
                    data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                    data_conv.Refresh
                    If data_conv.Recordset.RecordCount > 0 Then
                       If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                          data_infno.Recordset.AddNew
                          data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                          data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                          data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                          data_infno.Recordset("cl_nombre") = "BAJA"
                          data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                          data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                          data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                          data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                          data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                          data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                          data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                          data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                          data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                          data_infno.Recordset.Update
                       End If
                    End If
                End If
            End If
         End If
         data_cli.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      Label3.Visible = False
      Label3.Caption = ""
      DoEvents
      
      data_inf.RecordSource = "Select * from infcli where cl_nombre ='" & "NO ESTA EN P.SAPP" & "' order by cl_apellid"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "ALTAS"
         Xlibexel.SaveAs ("C:\planillas\" & "CGalicia-Altas" & ".xls")
         Xarchtex = "C:\planillas\CGalicia-Altas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE ALTAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_inf.Recordset.EOF
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_celular")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_celular")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_fnac")) = False Then
               Xarchexel.Cells(Xlin, XCol) = CDate(data_inf.Recordset("cl_fnac"))
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nom_sup")
            Else
               Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("info_debit")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("info_debit")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
            End If
            data_inf.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      Else
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "ALTAS"
         Xlibexel.SaveAs ("C:\planillas\" & "CGalicia-Altas" & ".xls")
         Xarchtex = "C:\planillas\CGalicia-Altas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE ALTAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      
      End If
      
      data_inf.RecordSource = "Select * from infcli where cl_nombre not in ('NO ESTA EN P.SAPP','ACTIVO') order by cl_apellid"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1
         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         
         Xarchexel.Name = "MODIF"
         Xlibexel.SaveAs ("C:\planillas\CGalicia-Mod.xls")
         Xarchtex = "C:\planillas\CGalicia-Mod.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE MODIFICACIONES MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "J" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 13
         Xarchexel.Cells(Xlin, XCol) = "MODIFICACION"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("I" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("J" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_inf.Recordset.EOF
            If IsNull(data_inf.Recordset("cl_nombre")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nombre")
            Else
               Xarchexel.Cells(Xlin, XCol) = "MODIF"
            End If
            XCol = XCol + 1
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_codigo")
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_apellid")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
            Else
               Xarchexel.Cells(Xlin, XCol) = "NN"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_celular")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_celular")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_fnac")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_fnac")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nom_sup")
            Else
               Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("info_debit")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("info_debit")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
            End If
            data_inf.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      End If
      
      data_infno.RecordSource = "Select * from infno where cl_nombre in ('BAJA') order by cl_apellid"
      data_infno.Refresh
      If data_infno.Recordset.RecordCount > 0 Then
         data_infno.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "BAJAS"
         Xlibexel.SaveAs ("C:\planillas\CGalicia-Bajas.xls")
         Xarchtex = "C:\planillas\CGalicia-Bajas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE BAJAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRES"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_infno.Recordset.EOF
            If IsNull(data_infno.Recordset("cl_codigo")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_codigo")
            Else
               Xarchexel.Cells(Xlin, XCol) = "0"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_apellid")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_apellid")
            Else
               Xarchexel.Cells(Xlin, XCol) = "NN"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_cedula")) = False Then
               If IsNull(data_infno.Recordset("cl_codced")) = False Then
                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_infno.Recordset("cl_cedula"))) & "-" & Trim(str(data_infno.Recordset("cl_codced")))
               Else
                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_infno.Recordset("cl_cedula"))) & "-0"
               End If
            Else
               Xarchexel.Cells(Xlin, XCol) = "0-0"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_codconv")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_codconv")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_zona")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_zona")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_direcci")
            End If
            data_infno.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      End If
      MsgBox "Proceso terminado"
   End If
End If

If Combo1.Text = "H.EVANGELICO" Then 'ok
   data_mut.RecordSource = "evang"
   data_mut.Refresh
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         data_mut.Recordset.Delete
         data_mut.Recordset.MoveNext
      Loop
   End If
   Data1.DatabaseName = "C:\mutuales\evang.xls"
   Data1.RecordSource = "socios$"
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         data_mut.Recordset.AddNew
         data_mut.Recordset("ced") = Data1.Recordset("ced")
         data_mut.Recordset("mat") = Data1.Recordset("mat")
         data_mut.Recordset("nom1") = Data1.Recordset("nom1")
         data_mut.Recordset("nom2") = Data1.Recordset("nom2")
         data_mut.Recordset("ape1") = Data1.Recordset("ape1")
         data_mut.Recordset("ape2") = Data1.Recordset("ape2")
         If IsNull(Data1.Recordset("fnac")) = False Then
            data_mut.Recordset("fnac") = Data1.Recordset("fnac")
         End If
         data_mut.Recordset("sexo") = Data1.Recordset("sexo")
         If IsNull(Data1.Recordset("categ")) = False Then
            data_mut.Recordset("categ") = Mid(Data1.Recordset("categ"), 1, 255)
         End If
         If IsNull(Data1.Recordset("domicilio")) = False Then
            data_mut.Recordset("domicilio") = Mid(Data1.Recordset("domicilio"), 1, 255)
         End If
         If IsNull(Data1.Recordset("telefono")) = False Then
            data_mut.Recordset("telefono") = Mid(Data1.Recordset("telefono"), 1, 255)
         End If
         If IsNull(Data1.Recordset("celular")) = False Then
            data_mut.Recordset("celular") = Mid(Data1.Recordset("celular"), 1, 255)
         End If
         If IsNull(Data1.Recordset("correo")) = False Then
            data_mut.Recordset("correo") = Mid(Data1.Recordset("correo"), 1, 255)
         End If
         If IsNull(Data1.Recordset("fecing")) = False Then
            data_mut.Recordset("fecing") = Data1.Recordset("fecing")
         End If
         data_mut.Recordset.Update
         Data1.Recordset.MoveNext
      Loop
   End If
   data_mut.Refresh
   Label3.Visible = True
   Label3.Caption = "Procesando Altas/Modif"
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveLast
      DoEvents
      pb.Visible = True
      pb.Max = data_mut.Recordset.RecordCount
      pb.Value = 0
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         If IsNull(data_mut.Recordset("ced")) = False Then
            Xcedeva = Val(data_mut.Recordset("ced"))
            data_mut.Recordset.Edit
            data_mut.Recordset("cednum") = Xcedeva
            data_mut.Recordset.Update
         Else
            Xcedeva = 0
         End If
         If Xcedeva > 0 Then
'            data_cli.Recordset.FindFirst "cl_cedula =" & Xcedeva
            data_cli.RecordSource = "Select * from clientes where cl_cedula =" & Xcedeva
            data_cli.Refresh
            If data_cli.Recordset.RecordCount > 0 Then
               data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
               data_conv.Refresh
               If data_conv.Recordset.RecordCount > 0 Then
                  If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                     If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                        data_inf.Recordset.AddNew
                        If IsNull(data_mut.Recordset("fnac")) = False Then
                           Xlafectex = Trim(Mid(data_mut.Recordset("fnac"), 1, 10))
                        Else
                           Xlafectex = ""
                        End If
                        If Xlafectex <> "" Then
                           data_inf.Recordset("cl_fnac") = CDate(Xlafectex)
                        End If
                        data_inf.Recordset("cl_celular") = data_mut.Recordset("ced")
                        If IsNull(data_mut.Recordset("domicilio")) = False Then
                           data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                        End If
                        If IsNull(data_mut.Recordset("categ")) = False Then
                           data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                        End If
                        If IsNull(data_mut.Recordset("celular")) = False Then
                           data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                        End If
                        If IsNull(data_mut.Recordset("telefono")) = False Then
                           data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                        End If
                        If IsNull(data_mut.Recordset("ape2")) = False Then
                           If IsNull(data_mut.Recordset("nom2")) = False Then
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                           Else
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                           End If
                        Else
                           If IsNull(data_mut.Recordset("nom2")) = False Then
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                           Else
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                           End If
                        End If
                        If IsNull(data_mut.Recordset("correo")) = False Then
                           data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                        End If
                        data_inf.Recordset("cl_nombre") = "REACTIVAR"
                        data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                        data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                        
                        data_inf.Recordset.Update
                     Else
                        If data_conv.Recordset("cnv_codigo") = "HEVANO" Or data_conv.Recordset("cnv_codigo") = "HEVAN" Or _
                           data_conv.Recordset("cnv_codigo") = "HEVANR" Or data_conv.Recordset("cnv_codigo") = "EVNSAM" Then
                            data_inf.Recordset.AddNew
                            If IsNull(data_mut.Recordset("fnac")) = False Then
                               Xlafectex = Trim(Mid(data_mut.Recordset("fnac"), 1, 10))
                            Else
                               Xlafectex = ""
                            End If
                            If Xlafectex <> "" Then
                               data_inf.Recordset("cl_fnac") = CDate(Xlafectex)
                            End If
                            data_inf.Recordset("cl_celular") = data_mut.Recordset("ced")
                            If IsNull(data_mut.Recordset("domicilio")) = False Then
                               data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                            End If
                            If IsNull(data_mut.Recordset("categ")) = False Then
                               data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                            End If
                            If IsNull(data_mut.Recordset("celular")) = False Then
                               data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                            End If
                            If IsNull(data_mut.Recordset("telefono")) = False Then
                               data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                            End If
                            If IsNull(data_mut.Recordset("ape2")) = False Then
                               If IsNull(data_mut.Recordset("nom2")) = False Then
                                  data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                               Else
                                  data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                               End If
                            Else
                               If IsNull(data_mut.Recordset("nom2")) = False Then
                                  data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                               Else
                                  data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                               End If
                            End If
                            If IsNull(data_mut.Recordset("correo")) = False Then
                               data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                            End If
                            data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                            data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                            data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                            data_inf.Recordset.Update
                        End If
                     End If
                  Else
                     data_inf.Recordset.AddNew
                     If IsNull(data_mut.Recordset("fnac")) = False Then
                        Xlafectex = Trim(Mid(data_mut.Recordset("fnac"), 1, 10))
                     Else
                        Xlafectex = ""
                     End If
                     If Xlafectex <> "" Then
                        data_inf.Recordset("cl_fnac") = CDate(Xlafectex)
                     End If
                     data_inf.Recordset("cl_celular") = data_mut.Recordset("ced")
                     If IsNull(data_mut.Recordset("domicilio")) = False Then
                        data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                     End If
                     If IsNull(data_mut.Recordset("categ")) = False Then
                        data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                     End If
                     If IsNull(data_mut.Recordset("celular")) = False Then
                        data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                     End If
                     If IsNull(data_mut.Recordset("telefono")) = False Then
                        data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                     End If
                     If IsNull(data_mut.Recordset("ape2")) = False Then
                        If IsNull(data_mut.Recordset("nom2")) = False Then
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                        Else
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                        End If
                     Else
                        If IsNull(data_mut.Recordset("nom2")) = False Then
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                        Else
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                        End If
                     End If
                     If IsNull(data_mut.Recordset("correo")) = False Then
                        data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                     End If
                     If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                        data_inf.Recordset("cl_nombre") = "CONV.INCORRECTO BAJ"
                     Else
                        data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                     End If
                     data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                     data_inf.Recordset.Update
                  End If
               Else
                  data_inf.Recordset.AddNew
                  If IsNull(data_mut.Recordset("fnac")) = False Then
                     Xlafectex = Trim(Mid(data_mut.Recordset("fnac"), 1, 10))
                  Else
                     Xlafectex = ""
                  End If
                  If Xlafectex <> "" Then
                     data_inf.Recordset("cl_fnac") = CDate(Xlafectex)
                  End If
                  data_inf.Recordset("cl_celular") = data_mut.Recordset("ced")
                  If IsNull(data_mut.Recordset("domicilio")) = False Then
                     data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                  End If
                  If IsNull(data_mut.Recordset("categ")) = False Then
                     data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                  End If
                  If IsNull(data_mut.Recordset("celular")) = False Then
                     data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                  End If
                  If IsNull(data_mut.Recordset("telefono")) = False Then
                     data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                  End If
                  If IsNull(data_mut.Recordset("ape2")) = False Then
                     If IsNull(data_mut.Recordset("nom2")) = False Then
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                     Else
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                     End If
                  Else
                     If IsNull(data_mut.Recordset("nom2")) = False Then
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                     Else
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                     End If
                  End If
                  If IsNull(data_mut.Recordset("correo")) = False Then
                     data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                  End If
                  If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                     data_inf.Recordset("cl_nombre") = "CONV.INCORRECTO BAJA"
                  Else
                     data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                  End If
                  data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                  data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                  data_inf.Recordset.Update
               End If
            Else
               data_inf.Recordset.AddNew
               If IsNull(data_mut.Recordset("fnac")) = False Then
                  Xlafectex = Trim(Mid(data_mut.Recordset("fnac"), 1, 10))
               Else
                  Xlafectex = ""
               End If
               If Xlafectex <> "" Then
                  data_inf.Recordset("cl_fnac") = CDate(Xlafectex)
               End If
               data_inf.Recordset("cl_celular") = data_mut.Recordset("ced")
               If IsNull(data_mut.Recordset("domicilio")) = False Then
                  data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
               End If
               If IsNull(data_mut.Recordset("categ")) = False Then
                  data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
               End If
               If IsNull(data_mut.Recordset("celular")) = False Then
                  data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
               End If
               If IsNull(data_mut.Recordset("telefono")) = False Then
                  data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
               End If
               If IsNull(data_mut.Recordset("ape2")) = False Then
                  If IsNull(data_mut.Recordset("nom2")) = False Then
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                  Else
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                  End If
               Else
                  If IsNull(data_mut.Recordset("nom2")) = False Then
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                  Else
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                  End If
               End If
               If IsNull(data_mut.Recordset("correo")) = False Then
                  data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
               End If
               data_inf.Recordset("cl_nombre") = "NO ESTA EN P.SAPP"
               data_inf.Recordset.Update
            End If
         End If
         data_mut.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      DoEvents
      data_cli.RecordSource = "Select * from clientes where estado <>" & 2 & " And estado <>" & 3 & " and cl_codconv not in ('UDEMM','PART','SMIN','UNIVS','EMERN','CASH','MSP','CAAMEP','UCM','CCNOS')"
      data_cli.Refresh
      data_cli.Recordset.MoveLast
      data_cli.Recordset.MoveFirst
      pb.Max = pb.Max + data_cli.Recordset.RecordCount
      data_mut.Refresh
      Label3.Caption = "Procesando BAJAS..."
      DoEvents
      Do While Not data_cli.Recordset.EOF
         If IsNull(data_cli.Recordset("cl_codconv")) = False Then
            If IsNull(data_cli.Recordset("cl_cedula")) = False Then
               If data_cli.Recordset("cl_cedula") > 0 Then
                  If data_cli.Recordset("cl_codconv") = "HEVANO" Or data_cli.Recordset("cl_codconv") = "HEVAN" Or _
                     data_cli.Recordset("cl_codconv") = "HEVANR" Or data_cli.Recordset("cl_codconv") = "EVNSAM" Then
                  Else
'                     data_conv.Recordset.FindFirst "cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                     data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                     data_conv.Refresh
                     If data_conv.Recordset.RecordCount > 0 Then
                        If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                           data_mut.RecordSource = "Select * from evang where cednum =" & Int(data_cli.Recordset("cl_cedula"))
                           data_mut.Refresh
'                           data_mut.Recordset.FindFirst "cednum =" & Int(data_cli.Recordset("cl_cedula"))
                           If data_mut.Recordset.RecordCount > 0 Then
                           Else
                              data_infno.Recordset.AddNew
                              data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                              data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                              data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                              data_infno.Recordset("cl_nombre") = "BAJA"
                              data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                              data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                              data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                              data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                              data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                              data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                              data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                              data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                              data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                              data_infno.Recordset.Update
                           End If
                        End If
                     End If
                  End If
               Else
                  data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                  data_conv.Refresh
                  If data_conv.Recordset.RecordCount > 0 Then
                     If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                        data_infno.Recordset.AddNew
                        data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                        data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                        data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                        data_infno.Recordset("cl_nombre") = "BAJA"
                        data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                        data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                        data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                        data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                        data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                        data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                        data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                        data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                        data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                        data_infno.Recordset.Update
                     End If
                  End If
               End If
            Else
               data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
               data_conv.Refresh
               If data_conv.Recordset.RecordCount > 0 Then
                  If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                     data_infno.Recordset.AddNew
                     data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                     data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                     data_infno.Recordset("cl_nombre") = "BAJA"
                     data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                     data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                     data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                     data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                     data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                     data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                     data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                     data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                     data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                     data_infno.Recordset.Update
                  End If
               End If
            End If
         End If
         data_cli.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      Label3.Visible = False
      Label3.Caption = ""
      DoEvents
      
      data_inf.RecordSource = "Select * from infcli where cl_nombre ='" & "NO ESTA EN P.SAPP" & "' order by cl_apellid"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "ALTAS"
         Xlibexel.SaveAs ("C:\planillas\" & "EVANG-Altas" & ".xls")
         Xarchtex = "C:\planillas\EVANG-Altas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE ALTAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_inf.Recordset.EOF
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_celular")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_celular")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_fnac")) = False Then
               Xarchexel.Cells(Xlin, XCol) = CDate(data_inf.Recordset("cl_fnac"))
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nom_sup")
            Else
               Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("info_debit")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("info_debit")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
            End If
            data_inf.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      Else
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "ALTAS"
         Xlibexel.SaveAs ("C:\planillas\" & "EVANG-Altas" & ".xls")
         Xarchtex = "C:\planillas\EVANG-Altas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE ALTAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      
      End If
      
      data_inf.RecordSource = "Select * from infcli where cl_nombre not in ('NO ESTA EN P.SAPP','ACTIVO') order by cl_apellid"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1
         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "MODIF"
         Xlibexel.SaveAs ("C:\planillas\EVANG-Mod.xls")
         Xarchtex = "C:\planillas\EVANG-Mod.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE MODIFICACIONES MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "J" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 20
         Xarchexel.Cells(Xlin, XCol) = "MODIFICACION"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("I" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("J" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("K" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_inf.Recordset.EOF
            If IsNull(data_inf.Recordset("cl_nombre")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nombre")
            Else
               Xarchexel.Cells(Xlin, XCol) = "MODIF"
            End If
            XCol = XCol + 1
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_codconv")
            
            XCol = XCol + 1
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_codigo")
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_apellid")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
            Else
               Xarchexel.Cells(Xlin, XCol) = "NN"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_celular")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_celular")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_fnac")) = False Then
               Xarchexel.Cells(Xlin, XCol) = CDate(data_inf.Recordset("cl_fnac"))
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nom_sup")
            Else
               Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("info_debit")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("info_debit")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
            End If
            data_inf.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      End If
      
      data_infno.RecordSource = "Select * from infno where cl_nombre in ('BAJA') order by cl_apellid"
      data_infno.Refresh
      If data_infno.Recordset.RecordCount > 0 Then
         data_infno.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "BAJAS"
         Xlibexel.SaveAs ("C:\planillas\EVANG-Bajas.xls")
         Xarchtex = "C:\planillas\EVANG-Bajas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE BAJAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRES"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_infno.Recordset.EOF
            If IsNull(data_infno.Recordset("cl_codigo")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_codigo")
            Else
               Xarchexel.Cells(Xlin, XCol) = "0"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_apellid")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_apellid")
            Else
               Xarchexel.Cells(Xlin, XCol) = "NN"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_cedula")) = False Then
               If IsNull(data_infno.Recordset("cl_codced")) = False Then
                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_infno.Recordset("cl_cedula"))) & "-" & Trim(str(data_infno.Recordset("cl_codced")))
               Else
                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_infno.Recordset("cl_cedula"))) & "-0"
               End If
            Else
               Xarchexel.Cells(Xlin, XCol) = "0-0"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_codconv")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_codconv")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_zona")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_zona")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_direcci")
            End If
            data_infno.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      End If
      MsgBox "Proceso terminado"
   End If
End If

If Combo1.Text = "CPS" Then 'OK
   Dim Xcedcps As String
   data_mut.RecordSource = "cps"
   data_mut.Refresh
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         data_mut.Recordset.Delete
         data_mut.Recordset.MoveNext
      Loop
   End If
   Data1.DatabaseName = "C:\mutuales\cps.xls"
   Data1.RecordSource = "socios$"
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         data_mut.Recordset.AddNew
         data_mut.Recordset("ced") = Data1.Recordset("ced")
         data_mut.Recordset("mat") = Data1.Recordset("mat")
         data_mut.Recordset("nom1") = Data1.Recordset("nom1")
         data_mut.Recordset("nom2") = Data1.Recordset("nom2")
         data_mut.Recordset("ape1") = Data1.Recordset("ape1")
         data_mut.Recordset("ape2") = Data1.Recordset("ape2")
         If IsNull(Data1.Recordset("fnac")) = False Then
            data_mut.Recordset("fnac") = Data1.Recordset("fnac")
         End If
         If IsNull(Data1.Recordset("categ")) = False Then
            data_mut.Recordset("categ") = Mid(Data1.Recordset("categ"), 1, 255)
         End If
         If IsNull(Data1.Recordset("domicilio")) = False Then
            data_mut.Recordset("domicilio") = Mid(Data1.Recordset("domicilio"), 1, 255)
         End If
         If IsNull(Data1.Recordset("telefono")) = False Then
            data_mut.Recordset("telefono") = Mid(Data1.Recordset("telefono"), 1, 255)
         End If
         If IsNull(Data1.Recordset("celular")) = False Then
            data_mut.Recordset("celular") = Mid(Data1.Recordset("celular"), 1, 255)
         End If
         If IsNull(Data1.Recordset("correo")) = False Then
            data_mut.Recordset("correo") = Mid(Data1.Recordset("correo"), 1, 255)
         End If
         If IsNull(Data1.Recordset("fecing")) = False Then
            data_mut.Recordset("fecing") = Data1.Recordset("fecing")
         End If
         data_mut.Recordset.Update
         Data1.Recordset.MoveNext
      Loop
   End If
   data_mut.Refresh
   Label3.Visible = True
   Label3.Caption = "Procesando Altas/Modif"
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveLast
      DoEvents
      pb.Visible = True
      pb.Max = data_mut.Recordset.RecordCount
      pb.Value = 0
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         If IsNull(data_mut.Recordset("ced")) = False Then
            Xcedeva = data_mut.Recordset("ced")
            If Len(data_mut.Recordset("ced")) = 7 Then
               Xcedcps = Mid(Trim(str(data_mut.Recordset("ced"))), 1, 6)
            Else
               Xcedcps = Mid(Trim(str(data_mut.Recordset("ced"))), 1, 7)
            End If
            data_mut.Recordset.Edit
            data_mut.Recordset("cednum") = Val(Xcedcps)
            data_mut.Recordset.Update
            Xcedeva = Val(Xcedcps)
         Else
            Xcedeva = 0
         End If
         If Xcedeva > 0 Then
'            data_cli.Recordset.FindFirst "cl_cedula =" & Xcedeva
            data_cli.RecordSource = "Select * from clientes where cl_cedula =" & Xcedeva
            data_cli.Refresh
            If data_cli.Recordset.RecordCount > 0 Then
               data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
               data_conv.Refresh
               If data_conv.Recordset.RecordCount > 0 Then
                  If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                     If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                        data_inf.Recordset.AddNew
                        data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                        data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced")))
                        If IsNull(data_mut.Recordset("domicilio")) = False Then
                           data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                        End If
                        If IsNull(data_mut.Recordset("categ")) = False Then
                           data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                        End If
                        If IsNull(data_mut.Recordset("celular")) = False Then
                           data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                        End If
                        If IsNull(data_mut.Recordset("telefono")) = False Then
                           data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                        End If
                        If IsNull(data_mut.Recordset("ape2")) = False Then
                           If IsNull(data_mut.Recordset("nom2")) = False Then
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                           Else
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                           End If
                        Else
                           If IsNull(data_mut.Recordset("nom2")) = False Then
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                           Else
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                           End If
                        End If
                        If IsNull(data_mut.Recordset("correo")) = False Then
                           data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                        End If
                        data_inf.Recordset("cl_nombre") = "REACTIVAR"
                        data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                        data_inf.Recordset.Update
                     Else
                     
                     End If
                  Else
                     data_inf.Recordset.AddNew
                     data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                     data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced")))
                     If IsNull(data_mut.Recordset("domicilio")) = False Then
                        data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                     End If
                     If IsNull(data_mut.Recordset("categ")) = False Then
                        data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                     End If
                     If IsNull(data_mut.Recordset("celular")) = False Then
                        data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                     End If
                     If IsNull(data_mut.Recordset("telefono")) = False Then
                        data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                     End If
                     If IsNull(data_mut.Recordset("ape2")) = False Then
                        If IsNull(data_mut.Recordset("nom2")) = False Then
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                        Else
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                        End If
                     Else
                        If IsNull(data_mut.Recordset("nom2")) = False Then
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                        Else
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                        End If
                     End If
                     If IsNull(data_mut.Recordset("correo")) = False Then
                        data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                     End If
                     If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                        data_inf.Recordset("cl_nombre") = "CONV.INCORRECTO BAJ"
                     Else
                        data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                     End If
                     data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_inf.Recordset.Update
                  End If
               Else
                  data_inf.Recordset.AddNew
                  data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                  data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced")))
                  If IsNull(data_mut.Recordset("domicilio")) = False Then
                     data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                  End If
                  If IsNull(data_mut.Recordset("categ")) = False Then
                     data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                  End If
                  If IsNull(data_mut.Recordset("celular")) = False Then
                     data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                  End If
                  If IsNull(data_mut.Recordset("telefono")) = False Then
                     data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                  End If
                  If IsNull(data_mut.Recordset("ape2")) = False Then
                     If IsNull(data_mut.Recordset("nom2")) = False Then
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                     Else
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                     End If
                  Else
                     If IsNull(data_mut.Recordset("nom2")) = False Then
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                     Else
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                     End If
                  End If
                  If IsNull(data_mut.Recordset("correo")) = False Then
                     data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                  End If
                  If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                     data_inf.Recordset("cl_nombre") = "CONV.INCORRECTO BAJA"
                  Else
                     data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                  End If
                  data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                  data_inf.Recordset.Update
               End If
            Else
               data_inf.Recordset.AddNew
               data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
               data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced")))
               If IsNull(data_mut.Recordset("domicilio")) = False Then
                  data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
               End If
               If IsNull(data_mut.Recordset("categ")) = False Then
                  data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
               End If
               If IsNull(data_mut.Recordset("celular")) = False Then
                  data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
               End If
               If IsNull(data_mut.Recordset("telefono")) = False Then
                  data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
               End If
               If IsNull(data_mut.Recordset("ape2")) = False Then
                  If IsNull(data_mut.Recordset("nom2")) = False Then
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                  Else
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                  End If
               Else
                  If IsNull(data_mut.Recordset("nom2")) = False Then
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                  Else
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                  End If
               End If
               If IsNull(data_mut.Recordset("correo")) = False Then
                  data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
               End If
               data_inf.Recordset("cl_nombre") = "NO ESTA EN P.SAPP"
               data_inf.Recordset.Update
            End If
         End If
         data_mut.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      DoEvents
      data_cli.RecordSource = "Select * from clientes where estado <>" & 2 & " And estado <>" & 3 & " and cl_codconv in ('CPS','CPSSA')"
      data_cli.Refresh
      data_cli.Recordset.MoveLast
      data_cli.Recordset.MoveFirst
      pb.Max = pb.Max + data_cli.Recordset.RecordCount
      data_mut.Refresh
      Label3.Caption = "Procesando BAJAS..."
      DoEvents
      Do While Not data_cli.Recordset.EOF
         If IsNull(data_cli.Recordset("cl_codconv")) = False Then
            If IsNull(data_cli.Recordset("cl_cedula")) = False Then
               If data_cli.Recordset("cl_cedula") > 0 Then
                  'If data_cli.Recordset("cl_codconv") = "HEVANO" Or data_cli.Recordset("cl_codconv") = "HEVAN" Or _
                  '   data_cli.Recordset("cl_codconv") = "HEVANR" Or data_cli.Recordset("cl_codconv") = "EVNSAM" Then
                  'Else
'                     data_conv.Recordset.FindFirst "cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                     data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                     data_conv.Refresh
                     If data_conv.Recordset.RecordCount > 0 Then
                        If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                           data_mut.RecordSource = "Select * from cps where cednum =" & Int(data_cli.Recordset("cl_cedula"))
                           data_mut.Refresh
'                           data_mut.Recordset.FindFirst "cednum =" & Int(data_cli.Recordset("cl_cedula"))
                           If data_mut.Recordset.RecordCount > 0 Then
                           Else
                              data_infno.Recordset.AddNew
                              data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                              data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                              data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                              data_infno.Recordset("cl_nombre") = "BAJA"
                              data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                              data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                              data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                              data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                              data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                              data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                              data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                              data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                              data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                              data_infno.Recordset.Update
                           End If
                        End If
                     End If
                  'End If
               Else
                  data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                  data_conv.Refresh
                  If data_conv.Recordset.RecordCount > 0 Then
                     If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                        data_infno.Recordset.AddNew
                        data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                        data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                        data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                        data_infno.Recordset("cl_nombre") = "BAJA"
                        data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                        data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                        data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                        data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                        data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                        data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                        data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                        data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                        data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                        data_infno.Recordset.Update
                     End If
                  End If
               End If
            Else
               data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
               data_conv.Refresh
               If data_conv.Recordset.RecordCount > 0 Then
                  If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                     data_infno.Recordset.AddNew
                     data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                     data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                     data_infno.Recordset("cl_nombre") = "BAJA"
                     data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                     data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                     data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                     data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                     data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                     data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                     data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                     data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                     data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                     data_infno.Recordset.Update
                  End If
               End If
            End If
         End If
         data_cli.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      Label3.Visible = False
      Label3.Caption = ""
      DoEvents

      data_inf.RecordSource = "Select * from infcli where cl_nombre ='" & "NO ESTA EN P.SAPP" & "' order by cl_apellid"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "ALTAS"
         Xlibexel.SaveAs ("C:\planillas\" & "CPS-Altas" & ".xls")
         Xarchtex = "C:\planillas\CPS-Altas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE ALTAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_inf.Recordset.EOF
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_celular")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_celular")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_fnac")) = False Then
               Xarchexel.Cells(Xlin, XCol) = CDate(data_inf.Recordset("cl_fnac"))
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nom_sup")
            Else
               Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("info_debit")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("info_debit")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
            End If
            data_inf.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      Else
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "ALTAS"
         Xlibexel.SaveAs ("C:\planillas\" & "CPS-Altas" & ".xls")
         Xarchtex = "C:\planillas\CPS-Altas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE ALTAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      
      End If
      
      data_inf.RecordSource = "Select * from infcli where cl_nombre not in ('NO ESTA EN P.SAPP','ACTIVO') order by cl_apellid"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1
         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "MODIF"
         Xlibexel.SaveAs ("C:\planillas\CPS-Mod.xls")
         Xarchtex = "C:\planillas\CPS-Mod.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE MODIFICACIONES MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "J" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 13
         Xarchexel.Cells(Xlin, XCol) = "MODIFICACION"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("I" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("J" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_inf.Recordset.EOF
            If IsNull(data_inf.Recordset("cl_nombre")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nombre")
            Else
               Xarchexel.Cells(Xlin, XCol) = "MODIF"
            End If
            XCol = XCol + 1
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_codigo")
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_apellid")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
            Else
               Xarchexel.Cells(Xlin, XCol) = "NN"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_celular")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_celular")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_fnac")) = False Then
               Xarchexel.Cells(Xlin, XCol) = CDate(data_inf.Recordset("cl_fnac"))
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nom_sup")
            Else
               Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("info_debit")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("info_debit")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
            End If
            data_inf.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      End If
      
      data_infno.RecordSource = "Select * from infno where cl_nombre in ('BAJA') order by cl_apellid"
      data_infno.Refresh
      If data_infno.Recordset.RecordCount > 0 Then
         data_infno.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "BAJAS"
         Xlibexel.SaveAs ("C:\planillas\CPS-Bajas.xls")
         Xarchtex = "C:\planillas\CPS-Bajas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE BAJAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRES"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_infno.Recordset.EOF
            If IsNull(data_infno.Recordset("cl_codigo")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_codigo")
            Else
               Xarchexel.Cells(Xlin, XCol) = "0"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_apellid")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_apellid")
            Else
               Xarchexel.Cells(Xlin, XCol) = "NN"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_cedula")) = False Then
               If IsNull(data_infno.Recordset("cl_codced")) = False Then
                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_infno.Recordset("cl_cedula"))) & "-" & Trim(str(data_infno.Recordset("cl_codced")))
               Else
                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_infno.Recordset("cl_cedula"))) & "-0"
               End If
            Else
               Xarchexel.Cells(Xlin, XCol) = "0-0"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_codconv")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_codconv")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_zona")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_zona")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_direcci")
            End If
            data_infno.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      End If
      MsgBox "Proceso terminado"
   End If

End If

'                  If data_cli.Recordset("cl_codconv") = "CCNOS" Or data_cli.Recordset("cl_codconv") = "CCNRE" Or _
'                     data_cli.Recordset("cl_codconv") = "CCNSAM" Then

If Combo1.Text = "CCOU" Then 'OK
   data_mut.RecordSource = "ccou"
   data_mut.Refresh
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         data_mut.Recordset.Delete
         data_mut.Recordset.MoveNext
      Loop
   End If
   Data1.DatabaseName = "C:\mutuales\ccou.xls"
   Data1.RecordSource = "socios$"
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         data_mut.Recordset.AddNew
         data_mut.Recordset("ced") = Data1.Recordset("ced")
         data_mut.Recordset("dv") = Data1.Recordset("dv")
         data_mut.Recordset("nom1") = Data1.Recordset("nom1")
         data_mut.Recordset("nom2") = Data1.Recordset("nom2")
         data_mut.Recordset("ape1") = Data1.Recordset("ape1")
         data_mut.Recordset("ape2") = Data1.Recordset("ape2")
         data_mut.Recordset("fnac") = Data1.Recordset("fnac")
         data_mut.Recordset("categ") = Mid(Data1.Recordset("categ"), 1, 255)
         data_mut.Recordset("domicilio") = Mid(Data1.Recordset("domicilio"), 1, 255)
         data_mut.Recordset("telefono") = Mid(Data1.Recordset("telefono"), 1, 255)
         data_mut.Recordset("celular") = Mid(Data1.Recordset("celular"), 1, 255)
         data_mut.Recordset("correo") = Mid(Data1.Recordset("correo"), 1, 255)
         data_mut.Recordset("fecing") = Data1.Recordset("fecing")
         data_mut.Recordset.Update
         Data1.Recordset.MoveNext
      Loop
   End If
   data_mut.Refresh
   Label3.Visible = True
   Label3.Caption = "Procesando Altas/Modif"
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveLast
      DoEvents
      pb.Visible = True
      pb.Max = data_mut.Recordset.RecordCount
      pb.Value = 0
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         If IsNull(data_mut.Recordset("ced")) = False Then
            Xcedeva = Val(data_mut.Recordset("ced"))
            data_mut.Recordset.Edit
            data_mut.Recordset("cednum") = Xcedeva
            data_mut.Recordset.Update
         Else
            Xcedeva = 0
         End If
         If Xcedeva > 0 Then
            data_cli.RecordSource = "Select * from clientes where cl_cedula =" & Xcedeva
            data_cli.Refresh
            If data_cli.Recordset.RecordCount > 0 Then
               data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
               data_conv.Refresh
               If data_conv.Recordset.RecordCount > 0 Then
                  If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                     If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                        data_inf.Recordset.AddNew
                        If IsNull(data_mut.Recordset("fnac")) = False Then
                           data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                        End If
                        data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced"))) & "-" & Trim(str(data_mut.Recordset("dv")))
                        If IsNull(data_mut.Recordset("domicilio")) = False Then
                           data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                        End If
                        If IsNull(data_mut.Recordset("categ")) = False Then
                           data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                        End If
                        If IsNull(data_mut.Recordset("celular")) = False Then
                           data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                        End If
                        If IsNull(data_mut.Recordset("telefono")) = False Then
                           data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                        End If
                        If IsNull(data_mut.Recordset("ape2")) = False Then
                           If IsNull(data_mut.Recordset("nom2")) = False Then
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                           Else
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                           End If
                        Else
                           If IsNull(data_mut.Recordset("nom2")) = False Then
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                           Else
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                           End If
                        End If
                        If IsNull(data_mut.Recordset("correo")) = False Then
                           data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                        End If
                        data_inf.Recordset("cl_nombre") = "REACTIVAR"
                        data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                        data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                        
                        data_inf.Recordset.Update
                     Else
                        If data_conv.Recordset("cnv_codigo") = "CCNOS" Or data_conv.Recordset("cnv_codigo") = "CCNRE" Or _
                           data_conv.Recordset("cnv_codigo") = "CCNSAM" Then
                            data_inf.Recordset.AddNew
                            If IsNull(data_mut.Recordset("fnac")) = False Then
                               data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                            End If
                            data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced"))) & "-" & Trim(str(data_mut.Recordset("dv")))
                            If IsNull(data_mut.Recordset("domicilio")) = False Then
                               data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                            End If
                            If IsNull(data_mut.Recordset("categ")) = False Then
                               data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                            End If
                            If IsNull(data_mut.Recordset("celular")) = False Then
                               data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                            End If
                            If IsNull(data_mut.Recordset("telefono")) = False Then
                               data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                            End If
                            If IsNull(data_mut.Recordset("ape2")) = False Then
                               If IsNull(data_mut.Recordset("nom2")) = False Then
                                  data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                               Else
                                  data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                               End If
                            Else
                               If IsNull(data_mut.Recordset("nom2")) = False Then
                                  data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                               Else
                                  data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                               End If
                            End If
                            If IsNull(data_mut.Recordset("correo")) = False Then
                               data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                            End If
                            data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                            data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                            data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                            data_inf.Recordset.Update
                        End If
                     End If
                  Else
                     data_inf.Recordset.AddNew
                     If IsNull(data_mut.Recordset("fnac")) = False Then
                        data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                     End If
                     data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced"))) & "-" & Trim(str(data_mut.Recordset("dv")))
                     If IsNull(data_mut.Recordset("domicilio")) = False Then
                        data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                     End If
                     If IsNull(data_mut.Recordset("categ")) = False Then
                        data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                     End If
                     If IsNull(data_mut.Recordset("celular")) = False Then
                        data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                     End If
                     If IsNull(data_mut.Recordset("telefono")) = False Then
                        data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                     End If
                     If IsNull(data_mut.Recordset("ape2")) = False Then
                        If IsNull(data_mut.Recordset("nom2")) = False Then
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                        Else
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                        End If
                     Else
                        If IsNull(data_mut.Recordset("nom2")) = False Then
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                        Else
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                        End If
                     End If
                     If IsNull(data_mut.Recordset("correo")) = False Then
                        data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                     End If
                     If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                        data_inf.Recordset("cl_nombre") = "CONV.INCORRECTO BAJ"
                     Else
                        data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                     End If
                     data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                     
                     data_inf.Recordset.Update
                  End If
               Else
                  data_inf.Recordset.AddNew
                  If IsNull(data_mut.Recordset("fnac")) = False Then
                     data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                  End If
                  data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced"))) & "-" & Trim(str(data_mut.Recordset("dv")))
                  If IsNull(data_mut.Recordset("domicilio")) = False Then
                     data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                  End If
                  If IsNull(data_mut.Recordset("categ")) = False Then
                     data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                  End If
                  If IsNull(data_mut.Recordset("celular")) = False Then
                     data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                  End If
                  If IsNull(data_mut.Recordset("telefono")) = False Then
                     data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                  End If
                  If IsNull(data_mut.Recordset("ape2")) = False Then
                     If IsNull(data_mut.Recordset("nom2")) = False Then
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                     Else
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                     End If
                  Else
                     If IsNull(data_mut.Recordset("nom2")) = False Then
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                     Else
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                     End If
                  End If
                  If IsNull(data_mut.Recordset("correo")) = False Then
                     data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                  End If
                  If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                     data_inf.Recordset("cl_nombre") = "CONV.INCORRECTO BAJA"
                  Else
                     data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                  End If
                  data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                  data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                  
                  data_inf.Recordset.Update
               End If
            Else
               data_inf.Recordset.AddNew
               If IsNull(data_mut.Recordset("fnac")) = False Then
                  data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
               End If
               data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced"))) & "-" & Trim(str(data_mut.Recordset("dv")))
               If IsNull(data_mut.Recordset("domicilio")) = False Then
                  data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
               End If
               If IsNull(data_mut.Recordset("categ")) = False Then
                  data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
               End If
               If IsNull(data_mut.Recordset("celular")) = False Then
                  data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
               End If
               If IsNull(data_mut.Recordset("telefono")) = False Then
                  data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
               End If
               If IsNull(data_mut.Recordset("ape2")) = False Then
                  If IsNull(data_mut.Recordset("nom2")) = False Then
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                  Else
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                  End If
               Else
                  If IsNull(data_mut.Recordset("nom2")) = False Then
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                  Else
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                  End If
               End If
               If IsNull(data_mut.Recordset("correo")) = False Then
                  data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
               End If
               data_inf.Recordset("cl_nombre") = "NO ESTA EN P.SAPP"
               data_inf.Recordset.Update
            End If
         End If
         data_mut.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      DoEvents
      data_cli.RecordSource = "Select * from clientes where estado <>" & 2 & " And estado <>" & 3 & " and cl_codconv not in ('UDEMM','PART','SMIN','UNIVS','EMERN','CASH','MSP','CAAMEP','UCM','CCNOS','CCFSJ','CCFSJA','CCFAT','CCFATA')"
      data_cli.Refresh
      data_cli.Recordset.MoveLast
      data_cli.Recordset.MoveFirst
      pb.Max = pb.Max + data_cli.Recordset.RecordCount
      data_mut.Refresh
      Label3.Caption = "Procesando BAJAS..."
      DoEvents
      Do While Not data_cli.Recordset.EOF
         If IsNull(data_cli.Recordset("cl_codconv")) = False Then
            If IsNull(data_cli.Recordset("cl_cedula")) = False Then
               If data_cli.Recordset("cl_cedula") > 0 Then
                  If data_cli.Recordset("cl_codconv") = "CCNOS" Or data_cli.Recordset("cl_codconv") = "CCNRE" Or _
                     data_cli.Recordset("cl_codconv") = "CCNSAM" Or data_cli.Recordset("cl_codconv") = "CCFSJ" Or _
                     data_cli.Recordset("cl_codconv") = "CCFSJA" Or data_cli.Recordset("cl_codconv") = "CCFAT" Or _
                     data_cli.Recordset("cl_codconv") = "CCFATA" Then
                  Else
'                     data_conv.Recordset.FindFirst "cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                     data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                     data_conv.Refresh
                     If data_conv.Recordset.RecordCount > 0 Then
                        If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                           data_mut.RecordSource = "Select * from ccou where cednum =" & Int(data_cli.Recordset("cl_cedula"))
                           data_mut.Refresh
'                           data_mut.Recordset.FindFirst "cednum =" & Int(data_cli.Recordset("cl_cedula"))
                           If data_mut.Recordset.RecordCount > 0 Then
                           Else
                              data_infno.Recordset.AddNew
                              data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                              data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                              data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                              data_infno.Recordset("cl_nombre") = "BAJA"
                              data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                              data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                              data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                              data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                              data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                              data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                              data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                              data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                              data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                              data_infno.Recordset.Update
                           End If
                        End If
                     End If
                  End If
               Else
                  data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                  data_conv.Refresh
                  If data_conv.Recordset.RecordCount > 0 Then
                     If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                        data_infno.Recordset.AddNew
                        data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                        data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                        data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                        data_infno.Recordset("cl_nombre") = "BAJA"
                        data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                        data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                        data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                        data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                        data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                        data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                        data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                        data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                        data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                        data_infno.Recordset.Update
                     End If
                  End If
               End If
            Else
               data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
               data_conv.Refresh
               If data_conv.Recordset.RecordCount > 0 Then
                  If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                     data_infno.Recordset.AddNew
                     data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                     data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                     data_infno.Recordset("cl_nombre") = "BAJA"
                     data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                     data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                     data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                     data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                     data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                     data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                     data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                     data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                     data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                     data_infno.Recordset.Update
                  End If
               End If
            End If
         End If
         data_cli.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      Label3.Visible = False
      Label3.Caption = ""
      DoEvents
      
      data_inf.RecordSource = "Select * from infcli where cl_nombre ='" & "NO ESTA EN P.SAPP" & "' order by cl_apellid"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "ALTAS"
         Xlibexel.SaveAs ("C:\planillas\" & "CCOU-Altas" & ".xls")
         Xarchtex = "C:\planillas\CCOU-Altas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE ALTAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_inf.Recordset.EOF
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_celular")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_celular")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_fnac")) = False Then
               Xarchexel.Cells(Xlin, XCol) = CDate(data_inf.Recordset("cl_fnac"))
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nom_sup")
            Else
               Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("info_debit")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("info_debit")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
            End If
            data_inf.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      Else
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "ALTAS"
         Xlibexel.SaveAs ("C:\planillas\" & "CCOU-Altas" & ".xls")
         Xarchtex = "C:\planillas\CCOU-Altas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE ALTAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      
      End If
      
      data_inf.RecordSource = "Select * from infcli where cl_nombre not in ('NO ESTA EN P.SAPP','ACTIVO','BAJA') order by cl_apellid"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1
         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "MODIF"
         Xlibexel.SaveAs ("C:\planillas\CCOU-Mod.xls")
         Xarchtex = "C:\planillas\CCOU-Mod.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE MODIFICACIONES MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "J" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 20
         Xarchexel.Cells(Xlin, XCol) = "MODIFICACION"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("I" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("J" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("K" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_inf.Recordset.EOF
            If IsNull(data_inf.Recordset("cl_nombre")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nombre")
            Else
               Xarchexel.Cells(Xlin, XCol) = "MODIF"
            End If
            XCol = XCol + 1
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_codconv")
            
            XCol = XCol + 1
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_codigo")
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_apellid")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
            Else
               Xarchexel.Cells(Xlin, XCol) = "NN"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_celular")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_celular")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_fnac")) = False Then
               Xarchexel.Cells(Xlin, XCol) = CDate(data_inf.Recordset("cl_fnac"))
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nom_sup")
            Else
               Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("info_debit")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("info_debit")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
            End If
            data_inf.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      End If
      
      data_infno.RecordSource = "Select * from infno where cl_nombre in ('BAJA') order by cl_apellid"
      data_infno.Refresh
      If data_infno.Recordset.RecordCount > 0 Then
         data_infno.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "BAJAS"
         Xlibexel.SaveAs ("C:\planillas\CCOU-Bajas.xls")
         Xarchtex = "C:\planillas\CCOU-Bajas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE BAJAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRES"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_infno.Recordset.EOF
            If IsNull(data_infno.Recordset("cl_codigo")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_codigo")
            Else
               Xarchexel.Cells(Xlin, XCol) = "0"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_apellid")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_apellid")
            Else
               Xarchexel.Cells(Xlin, XCol) = "NN"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_cedula")) = False Then
               If IsNull(data_infno.Recordset("cl_codced")) = False Then
                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_infno.Recordset("cl_cedula"))) & "-" & Trim(str(data_infno.Recordset("cl_codced")))
               Else
                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_infno.Recordset("cl_cedula"))) & "-0"
               End If
            Else
               Xarchexel.Cells(Xlin, XCol) = "0-0"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_codconv")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_codconv")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_zona")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_zona")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_direcci")
            End If
            data_infno.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      End If
      MsgBox "Proceso terminado"
   End If


End If

If Combo1.Text = "SEMM" Then 'OK
   Dim Xcedsem As String
   data_mut.RecordSource = "semm"
   data_mut.Refresh
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         data_mut.Recordset.Delete
         data_mut.Recordset.MoveNext
      Loop
   End If
   Data1.DatabaseName = "C:\mutuales\semm.xls"
   Data1.RecordSource = "socios$"
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         If IsNull(Data1.Recordset("nom1")) = False Then
            If Trim(Data1.Recordset("nom1")) <> "" Then
                data_mut.Recordset.AddNew
                data_mut.Recordset("ced") = Data1.Recordset("ced")
                If IsNull(Data1.Recordset("fnac")) = False Then
                   data_mut.Recordset("fnac") = Data1.Recordset("fnac")
                End If
                data_mut.Recordset("nom1") = Data1.Recordset("nom1")
                data_mut.Recordset("nom2") = Data1.Recordset("nom2")
                data_mut.Recordset("ape1") = Data1.Recordset("ape1")
                data_mut.Recordset("ape2") = Data1.Recordset("ape2")
                If IsNull(Data1.Recordset("categ")) = False Then
                   data_mut.Recordset("categ") = Mid(Data1.Recordset("categ"), 1, 255)
                End If
                If IsNull(Data1.Recordset("domicilio")) = False Then
                   data_mut.Recordset("domicilio") = Mid(Data1.Recordset("domicilio"), 1, 255)
                End If
                If IsNull(Data1.Recordset("telefono")) = False Then
                   data_mut.Recordset("telefono") = Mid(Data1.Recordset("telefono"), 1, 255)
                End If
                If IsNull(Data1.Recordset("celular")) = False Then
                   data_mut.Recordset("celular") = Mid(Data1.Recordset("celular"), 1, 255)
                End If
                If IsNull(Data1.Recordset("correo")) = False Then
                   data_mut.Recordset("correo") = Mid(Data1.Recordset("correo"), 1, 255)
                End If
                data_mut.Recordset("fecing") = Data1.Recordset("fecing")
                data_mut.Recordset.Update
            End If
         End If
         Data1.Recordset.MoveNext
      Loop
   End If
   
   data_mut.RecordSource = "semm"
   data_mut.Refresh
   Label3.Visible = True
   
   Label3.Caption = "Procesando Altas/Modif"
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveLast
      DoEvents
      pb.Visible = True
      pb.Max = data_mut.Recordset.RecordCount
      pb.Value = 0
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         If IsNull(data_mut.Recordset("ced")) = False Then
            Xcedeva = data_mut.Recordset("ced")
            If Len(data_mut.Recordset("ced")) = 7 Then
               Xcedsem = Mid(Trim(str(data_mut.Recordset("ced"))), 1, 6)
            Else
               Xcedsem = Mid(Trim(str(data_mut.Recordset("ced"))), 1, 7)
            End If
            data_mut.Recordset.Edit
            data_mut.Recordset("cednum") = Val(Xcedsem)
            data_mut.Recordset.Update
            Xcedeva = Val(Xcedsem)
         Else
            Xcedeva = 0
         End If
         If Xcedeva > 0 Then
'            data_cli.Recordset.FindFirst "cl_cedula =" & Xcedeva
            data_cli.RecordSource = "Select * from clientes where cl_cedula =" & Xcedeva
            data_cli.Refresh
            If data_cli.Recordset.RecordCount > 0 Then
               data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
               data_conv.Refresh
               If data_conv.Recordset.RecordCount > 0 Then
                  If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                     If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                        data_inf.Recordset.AddNew
                        If IsNull(data_mut.Recordset("fnac")) = False Then
                           data_inf.Recordset("cl_fnac") = Format(data_mut.Recordset("fnac"), "dd/mm/yyyy")
                        End If
                        data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced")))
                        If IsNull(data_mut.Recordset("domicilio")) = False Then
                           data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                        End If
                        If IsNull(data_mut.Recordset("categ")) = False Then
                           data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                        End If
                        If IsNull(data_mut.Recordset("celular")) = False Then
                           data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                        End If
                        If IsNull(data_mut.Recordset("telefono")) = False Then
                           data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                        End If
                        If IsNull(data_mut.Recordset("ape2")) = False Then
                           If IsNull(data_mut.Recordset("nom2")) = False Then
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                           Else
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                           End If
                        Else
                           If IsNull(data_mut.Recordset("nom2")) = False Then
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                           Else
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                           End If
                        End If
                        If IsNull(data_mut.Recordset("correo")) = False Then
                           data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                        End If
                        data_inf.Recordset("cl_nombre") = "REACTIVAR"
                        data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                        data_inf.Recordset.Update
                     Else
                        If data_conv.Recordset("cnv_codigo") = "SEMM" Then
                            data_inf.Recordset.AddNew
                            If IsNull(data_mut.Recordset("fnac")) = False Then
                               data_inf.Recordset("cl_fnac") = Format(data_mut.Recordset("fnac"), "dd/mm/yyyy")
                            End If
                            data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced")))
                            If IsNull(data_mut.Recordset("domicilio")) = False Then
                               data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                            End If
                            If IsNull(data_mut.Recordset("categ")) = False Then
                               data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                            End If
                            If IsNull(data_mut.Recordset("celular")) = False Then
                               data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                            End If
                            If IsNull(data_mut.Recordset("telefono")) = False Then
                               data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                            End If
                            If IsNull(data_mut.Recordset("ape2")) = False Then
                               If IsNull(data_mut.Recordset("nom2")) = False Then
                                  data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                               Else
                                  data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                               End If
                            Else
                               If IsNull(data_mut.Recordset("nom2")) = False Then
                                  data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                               Else
                                  data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                               End If
                            End If
                            If IsNull(data_mut.Recordset("correo")) = False Then
                               data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                            End If
                            data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                            data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                            data_inf.Recordset.Update
                        End If
                     End If
                  Else
                     data_inf.Recordset.AddNew
                     If IsNull(data_mut.Recordset("fnac")) = False Then
                        data_inf.Recordset("cl_fnac") = Format(data_mut.Recordset("fnac"), "dd/mm/yyyy")
                     End If
                     data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced")))
                     If IsNull(data_mut.Recordset("domicilio")) = False Then
                        data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                     End If
                     If IsNull(data_mut.Recordset("categ")) = False Then
                        data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                     End If
                     If IsNull(data_mut.Recordset("celular")) = False Then
                        data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                     End If
                     If IsNull(data_mut.Recordset("telefono")) = False Then
                        data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                     End If
                     If IsNull(data_mut.Recordset("ape2")) = False Then
                        If IsNull(data_mut.Recordset("nom2")) = False Then
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                        Else
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                        End If
                     Else
                        If IsNull(data_mut.Recordset("nom2")) = False Then
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                        Else
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                        End If
                     End If
                     If IsNull(data_mut.Recordset("correo")) = False Then
                        data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                     End If
                     If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                        data_inf.Recordset("cl_nombre") = "CONV.INCORRECTO BAJ"
                     Else
                        data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                     End If
                     data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_inf.Recordset.Update
                  End If
               Else
                  data_inf.Recordset.AddNew
                  If IsNull(data_mut.Recordset("fnac")) = False Then
                     data_inf.Recordset("cl_fnac") = Format(data_mut.Recordset("fnac"), "dd/mm/yyyy")
                  End If
                  data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced")))
                  If IsNull(data_mut.Recordset("domicilio")) = False Then
                     data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                  End If
                  If IsNull(data_mut.Recordset("categ")) = False Then
                     data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                  End If
                  If IsNull(data_mut.Recordset("celular")) = False Then
                     data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                  End If
                  If IsNull(data_mut.Recordset("telefono")) = False Then
                     data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                  End If
                  If IsNull(data_mut.Recordset("ape2")) = False Then
                     If IsNull(data_mut.Recordset("nom2")) = False Then
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                     Else
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                     End If
                  Else
                     If IsNull(data_mut.Recordset("nom2")) = False Then
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                     Else
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                     End If
                  End If
                  If IsNull(data_mut.Recordset("correo")) = False Then
                     data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                  End If
                  If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                     data_inf.Recordset("cl_nombre") = "CONV.INCORRECTO BAJA"
                  Else
                     data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                  End If
                  data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                  data_inf.Recordset.Update
               End If
            Else
               data_inf.Recordset.AddNew
               If IsNull(data_mut.Recordset("fnac")) = False Then
                  data_inf.Recordset("cl_fnac") = Format(data_mut.Recordset("fnac"), "dd/mm/yyyy")
               End If
               data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced")))
               If IsNull(data_mut.Recordset("domicilio")) = False Then
                  data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
               End If
               If IsNull(data_mut.Recordset("categ")) = False Then
                  data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
               End If
               If IsNull(data_mut.Recordset("celular")) = False Then
                  data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
               End If
               If IsNull(data_mut.Recordset("telefono")) = False Then
                  data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
               End If
               If IsNull(data_mut.Recordset("ape2")) = False Then
                  If IsNull(data_mut.Recordset("nom2")) = False Then
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                  Else
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                  End If
               Else
                  If IsNull(data_mut.Recordset("nom2")) = False Then
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                  Else
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                  End If
               End If
               If IsNull(data_mut.Recordset("correo")) = False Then
                  data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
               End If
               data_inf.Recordset("cl_nombre") = "NO ESTA EN P.SAPP"
               data_inf.Recordset.Update
            End If
         End If
         data_mut.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      DoEvents
      data_cli.RecordSource = "Select * from clientes where estado <>" & 2 & " And estado <>" & 3 & " and cl_codconv in ('SEMM1')"
      data_cli.Refresh
      data_cli.Recordset.MoveLast
      data_cli.Recordset.MoveFirst
      pb.Max = pb.Max + data_cli.Recordset.RecordCount
      data_mut.Refresh
      Label3.Caption = "Procesando BAJAS..."
      DoEvents
      Do While Not data_cli.Recordset.EOF
         If IsNull(data_cli.Recordset("cl_codconv")) = False Then
            If IsNull(data_cli.Recordset("cl_cedula")) = False Then
               If data_cli.Recordset("cl_cedula") > 0 Then
                  'If data_cli.Recordset("cl_codconv") = "HEVANO" Or data_cli.Recordset("cl_codconv") = "HEVAN" Or _
                  '   data_cli.Recordset("cl_codconv") = "HEVANR" Or data_cli.Recordset("cl_codconv") = "EVNSAM" Then
                  'Else
'                     data_conv.Recordset.FindFirst "cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                     data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                     data_conv.Refresh
                     If data_conv.Recordset.RecordCount > 0 Then
                        If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                           data_mut.RecordSource = "Select * from semm where cednum =" & Int(data_cli.Recordset("cl_cedula"))
                           data_mut.Refresh
'                           data_mut.Recordset.FindFirst "cednum =" & Int(data_cli.Recordset("cl_cedula"))
                           If data_mut.Recordset.RecordCount > 0 Then
                           Else
                              data_infno.Recordset.AddNew
                              data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                              data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                              data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                              data_infno.Recordset("cl_nombre") = "BAJA"
                              data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                              data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                              data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                              data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                              data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                              data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                              data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                              data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                              data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                              data_infno.Recordset.Update
                           End If
                        End If
                     End If
                  'End If
               Else
                  data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                  data_conv.Refresh
                  If data_conv.Recordset.RecordCount > 0 Then
                     If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                        data_infno.Recordset.AddNew
                        data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                        data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                        data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                        data_infno.Recordset("cl_nombre") = "BAJA"
                        data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                        data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                        data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                        data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                        data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                        data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                        data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                        data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                        data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                        data_infno.Recordset.Update
                     End If
                  End If
               End If
            Else
               data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
               data_conv.Refresh
               If data_conv.Recordset.RecordCount > 0 Then
                  If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                     data_infno.Recordset.AddNew
                     data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                     data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                     data_infno.Recordset("cl_nombre") = "BAJA"
                     data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                     data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                     data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                     data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                     data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                     data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                     data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                     data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                     data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                     data_infno.Recordset.Update
                  End If
               End If
            End If
         End If
         data_cli.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      Label3.Visible = False
      Label3.Caption = ""
      DoEvents

      data_inf.RecordSource = "Select * from infcli where cl_nombre ='" & "NO ESTA EN P.SAPP" & "' order by cl_apellid"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "ALTAS"
         Xlibexel.SaveAs ("C:\planillas\" & "SEMM-Altas" & ".xls")
         Xarchtex = "C:\planillas\SEMM-Altas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE ALTAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_inf.Recordset.EOF
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_celular")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_celular")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_fnac")) = False Then
               Xarchexel.Cells(Xlin, XCol) = CDate(data_inf.Recordset("cl_fnac"))
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nom_sup")
            Else
               Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("info_debit")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("info_debit")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
            End If
            data_inf.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      Else
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "ALTAS"
         Xlibexel.SaveAs ("C:\planillas\" & "SEMM-Altas" & ".xls")
         Xarchtex = "C:\planillas\SEMM-Altas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE ALTAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      
      End If
      
      data_inf.RecordSource = "Select * from infcli where cl_nombre not in ('NO ESTA EN P.SAPP','ACTIVO') order by cl_apellid"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1
         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "MODIF"
         Xlibexel.SaveAs ("C:\planillas\SEMM-Mod.xls")
         Xarchtex = "C:\planillas\SEMM-Mod.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE MODIFICACIONES MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "J" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 13
         Xarchexel.Cells(Xlin, XCol) = "MODIFICACION"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("I" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("J" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_inf.Recordset.EOF
            If IsNull(data_inf.Recordset("cl_nombre")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nombre")
            Else
               Xarchexel.Cells(Xlin, XCol) = "MODIF"
            End If
            XCol = XCol + 1
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_codigo")
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_apellid")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
            Else
               Xarchexel.Cells(Xlin, XCol) = "NN"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_celular")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_celular")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_fnac")) = False Then
               Xarchexel.Cells(Xlin, XCol) = CDate(data_inf.Recordset("cl_fnac"))
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nom_sup")
            Else
               Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("info_debit")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("info_debit")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
            End If
            data_inf.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      End If
      
      data_infno.RecordSource = "Select * from infno where cl_nombre in ('BAJA') order by cl_apellid"
      data_infno.Refresh
      If data_infno.Recordset.RecordCount > 0 Then
         data_infno.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "BAJAS"
         Xlibexel.SaveAs ("C:\planillas\SEMM-Bajas.xls")
         Xarchtex = "C:\planillas\SEMM-Bajas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE BAJAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRES"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_infno.Recordset.EOF
            If IsNull(data_infno.Recordset("cl_codigo")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_codigo")
            Else
               Xarchexel.Cells(Xlin, XCol) = "0"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_apellid")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_apellid")
            Else
               Xarchexel.Cells(Xlin, XCol) = "NN"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_cedula")) = False Then
               If IsNull(data_infno.Recordset("cl_codced")) = False Then
                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_infno.Recordset("cl_cedula"))) & "-" & Trim(str(data_infno.Recordset("cl_codced")))
               Else
                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_infno.Recordset("cl_cedula"))) & "-0"
               End If
            Else
               Xarchexel.Cells(Xlin, XCol) = "0-0"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_codconv")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_codconv")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_zona")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_zona")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_direcci")
            End If
            data_infno.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      End If
      MsgBox "Proceso terminado"
   End If
   
End If
If Combo1.Text = "RET.MILITARES" Then
   Command3.Visible = True
   Command3_Click
   Command3.Visible = False
End If
If Combo1.Text = "SMI" Or Combo1.Text = "CASMU" Then
   Command3.Visible = True
   Command3_Click
   Command3.Visible = False
End If
If Combo1.Text = "CASH" Or Combo1.Text = "UNIVERSAL" Then
   Command4.Visible = True
   Command4_Click
   Command4.Visible = False
End If
If Combo1.Text = "BLUE CROSS" Then
   BlueCross
End If
If Combo1.Text = "CAUTE" Then
   Caute
End If
If Combo1.Text = "SEGURO AMERICANO" Then
   Seguro
End If
If Combo1.Text = "SUMMUM" Then
   Summum
End If

If Combo1.Text = "CCOU SJ" Then
   CcouSJ
End If
If Combo1.Text = "CCOU ATL" Then
   CcouAtl
End If


frm_ctrolmut.MousePointer = 0

Exit Sub

Quepasamut:
            If Err.Number = 53 Then
               frm_ctrolmut.MousePointer = 0
               MsgBox "Error en proceso " & Err.Description
            Else
               frm_ctrolmut.MousePointer = 0
               MsgBox "Error en proceso :" & Err.Description
            End If
            
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
Dim Xcedeva As Long

Dim Xobjexel As Excel.Application
Dim Xlibexel As Excel.Workbook
Dim Xarchexel As New Excel.Worksheet

Dim XCol, Xlin, Xnrocan, Xcolfija As Long
Dim Xarchtex As String
Dim Xlabrir As New Excel.Application

frm_ctrolmut.MousePointer = 11

If Combo1.Text = "CASMU" Then 'OK
   data_mut.RecordSource = "casmu"
   data_mut.Refresh
'   data_cli.RecordSource = "Select * from clientes where estado <>" & 2 & " And estado <>" & 3
'   data_cli.RecordSource = "Select * from clientes"
'   data_cli.Refresh
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         If IsNull(data_mut.Recordset("ced")) = False Then
            If data_mut.Recordset("ced") <> "" Then
               If Len(data_mut.Recordset("ced")) = 9 Then
                  Xccedeva = Val(Mid(Trim(data_mut.Recordset("ced")), 9, 1))
               Else
                  If Len(data_mut.Recordset("ced")) = 8 Then
                     Xccedeva = Val(Mid(Trim(data_mut.Recordset("ced")), 8, 1))
                  Else
                     Xccedeva = 88
                  End If
               End If
            Else
               Xccedeva = 88
            End If
         Else
            Xccedeva = 88
         End If
         data_mut.Recordset.Edit
         data_mut.Recordset("codced") = Xccedeva
         data_mut.Recordset.Update
         data_mut.Recordset.MoveNext
      Loop
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         If IsNull(data_mut.Recordset("ced")) = False Then
            Xcedeva = Val(data_mut.Recordset("ced"))
         Else
            Xcedeva = 0
         End If
         If Xcedeva > 0 Then
            data_cli.RecordSource = "Select * from clientes where cl_cedula =" & Xcedeva
            data_cli.Refresh
            If data_cli.Recordset.RecordCount > 0 Then
               data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
               data_conv.Refresh
               If data_conv.Recordset.RecordCount > 0 Then
                  If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                     If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                        data_inf.Recordset.AddNew
                         data_inf.Recordset("cl_celular") = data_mut.Recordset("ced")
    '                     data_inf.Recordset("cl_celular") = Trim(Str(data_mut.Recordset("ced"))) & "-" & Trim(Str(data_mut.Recordset("co")))
                         data_inf.Recordset("cl_apellid") = data_mut.Recordset("apel1") + " " + data_mut.Recordset("nom1")
                         data_inf.Recordset("cl_nombre") = "REACTIVAR"
                         data_inf.Recordset.Update
                     End If
                  Else
                     data_inf.Recordset.AddNew
                     data_inf.Recordset("cl_celular") = data_mut.Recordset("ced")
 '                     data_inf.Recordset("cl_celular") = Trim(Str(data_mut.Recordset("ced"))) & "-" & Trim(Str(data_mut.Recordset("co")))
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("apel1") + " " + data_mut.Recordset("nom1")
                     If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                        data_inf.Recordset("cl_nombre") = "CONV.INCORRECTO BAJ"
                     Else
                        data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                     End If
                     data_inf.Recordset.Update
                  End If
               Else
                  data_inf.Recordset.AddNew
                  data_inf.Recordset("cl_celular") = data_mut.Recordset("ced")
    '             data_inf.Recordset("cl_celular") = Trim(Str(data_mut.Recordset("ced"))) & "-" & Trim(Str(data_mut.Recordset("co")))
                  data_inf.Recordset("cl_apellid") = data_mut.Recordset("apel1") + " " + data_mut.Recordset("nom1")
                  If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                     data_inf.Recordset("cl_nombre") = "CONV.INCORRECTO BAJA"
                  Else
                     data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                  End If
                  data_inf.Recordset.Update
               End If
            Else
               data_inf.Recordset.AddNew
'               data_inf.Recordset("cl_celular") = Trim(Str(data_mut.Recordset("ced"))) & "-" & Trim(Str(data_mut.Recordset("co")))
               data_inf.Recordset("cl_celular") = data_mut.Recordset("ced")
    '          data_inf.Recordset("cl_celular") = Trim(Str(data_mut.Recordset("ced"))) & "-" & Trim(Str(data_mut.Recordset("co")))
               data_inf.Recordset("cl_apellid") = data_mut.Recordset("apel1") + " " + data_mut.Recordset("nom1")
'               If IsNull(data_mut.Recordset("sede")) = False Then
'                  data_inf.Recordset("cl_localid") = data_mut.Recordset("sede")
'               End If
               data_inf.Recordset("cl_nombre") = "NO ESTA EN P.SAPP"
               data_inf.Recordset.Update
            End If
         Else
            data_inf.Recordset.AddNew
            data_inf.Recordset("cl_celular") = data_mut.Recordset("ced")
    '       data_inf.Recordset("cl_celular") = Trim(Str(data_mut.Recordset("ced"))) & "-" & Trim(Str(data_mut.Recordset("co")))
            data_inf.Recordset("cl_apellid") = data_mut.Recordset("apel1") + " " + data_mut.Recordset("nom1")
'            If IsNull(data_mut.Recordset("sede")) = False Then
'               data_inf.Recordset("cl_localid") = data_mut.Recordset("sede")
'            End If
            data_inf.Recordset("cl_nombre") = "CI.0 EN casmu"
            data_inf.Recordset.Update
         End If
         data_mut.Recordset.Edit
         data_mut.Recordset("cednum") = Xcedeva
         data_mut.Recordset.Update
         data_mut.Recordset.MoveNext
      Loop
      data_cli.RecordSource = "Select * from clientes where estado <>" & 2 & " And estado <>" & 3
      data_cli.Refresh
      data_cli.Recordset.MoveFirst
      Do While Not data_cli.Recordset.EOF
         If IsNull(data_cli.Recordset("cl_codconv")) = False Then
            If IsNull(data_cli.Recordset("cl_cedula")) = False Then
               If data_cli.Recordset("cl_cedula") > 0 Then
                  data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                  data_conv.Refresh
                  If data_conv.Recordset.RecordCount > 0 Then
                     If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                        data_mut.RecordSource = "Select * from casmu where cednum =" & Int(data_cli.Recordset("cl_cedula"))
                        data_mut.Refresh
                        If data_mut.Recordset.RecordCount <= 0 Then
'''                        data_mut.Recordset.FindFirst "cednum =" & Int(data_cli.Recordset("cl_cedula"))
'''                        If data_mut.Recordset.NoMatch Then
                           data_infno.Recordset.AddNew
                           data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                           data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                           data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                           data_infno.Recordset("cl_nombre") = "BAJA"
                           data_infno.Recordset.Update
                        End If
                     End If
                  End If
               Else
                  data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                  data_conv.Refresh
                  If data_conv.Recordset.RecordCount > 0 Then
                     If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                        data_infno.Recordset.AddNew
                        data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                        data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                        data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                        data_infno.Recordset("cl_nombre") = "BAJA"
                        data_infno.Recordset.Update
                     End If
                  End If
               End If
            Else
               data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
               data_conv.Refresh
               If data_conv.Recordset.RecordCount > 0 Then
                  If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                     data_infno.Recordset.AddNew
                     data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                     data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                     data_infno.Recordset("cl_nombre") = "BAJA"
                     data_infno.Recordset.Update
                  End If
               End If
            End If
         End If
         data_cli.Recordset.MoveNext
      Loop
      
      data_inf.RecordSource = "Select * from infcli order by cl_apellid"
      data_inf.Refresh
      cr2.ReportTitle = "MUTUALISTA: CASMU"
      cr2.ReportFileName = App.path & "\infctrolmut2.rpt"
      cr2.Action = 1
   
      data_infno.RecordSource = "Select * from infno order by cl_apellid"
      data_infno.Refresh
      cr3.ReportTitle = "MUTUALISTA: CASMU"
      cr3.ReportFileName = App.path & "\infctrolmutb.rpt"
      cr3.Action = 1
   
      data_inf.RecordSource = "Select * from infcli order by cl_apellid"
      data_inf.Refresh
      cr1.ReportTitle = "MUTUALISTA: CASMU"
      cr1.ReportFileName = App.path & "\infctrolmut.rpt"
      cr1.Action = 1
   End If
End If

If Combo1.Text = "SMI" Then 'OK
   data_mut.RecordSource = "smi"
   data_mut.Refresh
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         data_mut.Recordset.Delete
         data_mut.Recordset.MoveNext
      Loop
   End If
   Data1.DatabaseName = "C:\mutuales\smi.xls"
   Data1.RecordSource = "socios$"
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         data_mut.Recordset.AddNew
         data_mut.Recordset("ced") = Data1.Recordset("ced")
         data_mut.Recordset("dv") = Data1.Recordset("dv")
         data_mut.Recordset("nombre") = Data1.Recordset("nombre")
         data_mut.Recordset("sexo") = Data1.Recordset("sexo")
         If IsNull(Data1.Recordset("fnac")) = False Then
            data_mut.Recordset("fnac") = Data1.Recordset("fnac")
         End If
         If IsNull(Data1.Recordset("categ")) = False Then
            data_mut.Recordset("categ") = Mid(Data1.Recordset("categ"), 1, 255)
         End If
         If IsNull(Data1.Recordset("domicilio")) = False Then
            data_mut.Recordset("domicilio") = Mid(Data1.Recordset("domicilio"), 1, 255)
         End If
         If IsNull(Data1.Recordset("telefono")) = False Then
            data_mut.Recordset("telefono") = Mid(Data1.Recordset("telefono"), 1, 255)
         End If
         If IsNull(Data1.Recordset("celular")) = False Then
            data_mut.Recordset("celular") = Mid(Data1.Recordset("celular"), 1, 255)
         End If
         If IsNull(Data1.Recordset("correo")) = False Then
            data_mut.Recordset("correo") = Mid(Data1.Recordset("correo"), 1, 255)
         End If
         data_mut.Recordset.Update
         Data1.Recordset.MoveNext
      Loop
   End If
   data_mut.Refresh
   Label3.Visible = True
   Label3.Caption = "Procesando Altas/Modif"
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveLast
      DoEvents
      pb.Visible = True
      pb.Max = data_mut.Recordset.RecordCount
      pb.Value = 0
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         If IsNull(data_mut.Recordset("ced")) = False Then
            Xcedeva = data_mut.Recordset("ced")
            data_mut.Recordset.Edit
            data_mut.Recordset("cednum") = Xcedeva
            data_mut.Recordset.Update
         Else
            Xcedeva = 0
         End If
         If Xcedeva > 0 Then
            data_cli.RecordSource = "Select * from clientes where cl_cedula =" & Xcedeva
            data_cli.Refresh
            If data_cli.Recordset.RecordCount > 0 Then
               If data_cli.Recordset("cl_codconv") = "SMIN" Or data_cli.Recordset("cl_codconv") = "SMINR" Or _
                  data_cli.Recordset("cl_codconv") = "SMINA" Then
                  data_inf.Recordset.AddNew
                  If IsNull(data_mut.Recordset("fnac")) = False Then
                     data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                  End If
                  data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced"))) & "-" & Trim(str(data_mut.Recordset("dv")))
                  If IsNull(data_mut.Recordset("domicilio")) = False Then
                     data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                  End If
                  If IsNull(data_mut.Recordset("categ")) = False Then
                     data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                  End If
                  If IsNull(data_mut.Recordset("celular")) = False Then
                     data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                  End If
                  If IsNull(data_mut.Recordset("telefono")) = False Then
                     data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                  End If
                  data_inf.Recordset("cl_apellid") = Mid(data_mut.Recordset("nombre"), 1, 60)
                  If IsNull(data_mut.Recordset("correo")) = False Then
                     data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                  End If
                  data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                  data_conv.RecordSource = "select * from conves_mut where nombre ='" & data_mut.Recordset("categ") & "'"
                  data_conv.Refresh
                  If data_conv.Recordset.RecordCount > 0 Then
                     data_inf.Recordset("cl_nomvend") = data_conv.Recordset("codsapp")
                  Else
                     data_inf.Recordset("cl_nomvend") = "NO ENCONTRADO"
                  End If
                  data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                  data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                  
                  data_inf.Recordset.Update
               Else
                    data_conv.RecordSource = "Select * from conves_mut where codsapp ='" & data_cli.Recordset("cl_codconv") & "' and nombre ='" & data_mut.Recordset("categ") & "'"
                    data_conv.Refresh
                    If data_conv.Recordset.RecordCount > 0 Then
                          If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                             data_inf.Recordset.AddNew
                             If IsNull(data_mut.Recordset("fnac")) = False Then
                                data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                             End If
                             data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced"))) & "-" & Trim(str(data_mut.Recordset("dv")))
                             If IsNull(data_mut.Recordset("domicilio")) = False Then
                                data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                             End If
                             If IsNull(data_mut.Recordset("categ")) = False Then
                                data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                             End If
                             If IsNull(data_mut.Recordset("celular")) = False Then
                                data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                             End If
                             If IsNull(data_mut.Recordset("telefono")) = False Then
                                data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                             End If
                             data_inf.Recordset("cl_apellid") = Mid(data_mut.Recordset("nombre"), 1, 60)
                             If IsNull(data_mut.Recordset("correo")) = False Then
                                data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                             End If
                             data_inf.Recordset("cl_nombre") = "REACTIVAR"
                             data_inf.Recordset("cl_nomvend") = "CORRECTO"
                             data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                             data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                             data_inf.Recordset.Update
                          End If
                    Else
                       data_inf.Recordset.AddNew
                       If IsNull(data_mut.Recordset("fnac")) = False Then
                          data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                       End If
                       data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced"))) & "-" & Trim(str(data_mut.Recordset("dv")))
                       If IsNull(data_mut.Recordset("domicilio")) = False Then
                          data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                       End If
                       If IsNull(data_mut.Recordset("categ")) = False Then
                          data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                       End If
                       If IsNull(data_mut.Recordset("celular")) = False Then
                          data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                       End If
                       If IsNull(data_mut.Recordset("telefono")) = False Then
                          data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                       End If
                       data_inf.Recordset("cl_apellid") = Mid(data_mut.Recordset("nombre"), 1, 60)
                       If IsNull(data_mut.Recordset("correo")) = False Then
                          data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                       End If
                       If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                          data_inf.Recordset("cl_nombre") = "CONV.INCORRECTO BAJA"
                       Else
                          data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                       End If
                       data_conv.RecordSource = "select * from conves_mut where nombre ='" & data_mut.Recordset("categ") & "'"
                       data_conv.Refresh
                       If data_conv.Recordset.RecordCount > 0 Then
                          data_inf.Recordset("cl_nomvend") = data_conv.Recordset("codsapp")
                       Else
                          data_inf.Recordset("cl_nomvend") = "NO ENCONTRADO"
                       End If
                       data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                       data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                       data_inf.Recordset.Update
                    End If
               End If
            Else
               data_inf.Recordset.AddNew
               If IsNull(data_mut.Recordset("fnac")) = False Then
                  data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
               End If
               data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced"))) & "-" & Trim(str(data_mut.Recordset("dv")))
               If IsNull(data_mut.Recordset("domicilio")) = False Then
                  data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
               End If
               If IsNull(data_mut.Recordset("categ")) = False Then
                  data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
               End If
               If IsNull(data_mut.Recordset("celular")) = False Then
                  data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
               End If
               If IsNull(data_mut.Recordset("telefono")) = False Then
                  data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
               End If
               data_inf.Recordset("cl_apellid") = Mid(data_mut.Recordset("nombre"), 1, 60)
               If IsNull(data_mut.Recordset("correo")) = False Then
                  data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
               End If
               data_inf.Recordset("cl_nombre") = "NO ESTA EN P.SAPP"
               data_conv.RecordSource = "select * from conves_mut where nombre ='" & data_mut.Recordset("categ") & "'"
               data_conv.Refresh
               If data_conv.Recordset.RecordCount > 0 Then
                  data_inf.Recordset("cl_nomvend") = data_conv.Recordset("codsapp")
               Else
                  data_inf.Recordset("cl_nomvend") = "NO ENCONTRADO"
               End If
               data_inf.Recordset.Update
            End If
         End If
         data_mut.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      DoEvents
      data_cli.RecordSource = "Select * from clientes where estado <>" & 2 & " And estado <>" & 3 & " and cl_codconv not in ('UDEMM','PART','UNIVS','HEVAN','EMERN','CASH','MSP','CAAMEP','UCM','CCNOS','HEVANO')"
      data_cli.Refresh
      data_cli.Recordset.MoveLast
      data_cli.Recordset.MoveFirst
      pb.Max = pb.Max + data_cli.Recordset.RecordCount
      data_mut.Refresh
      Label3.Caption = "Procesando BAJAS..."
      DoEvents
      Do While Not data_cli.Recordset.EOF
         If IsNull(data_cli.Recordset("cl_codconv")) = False Then
            If IsNull(data_cli.Recordset("cl_cedula")) = False Then
               If data_cli.Recordset("cl_cedula") > 0 Then
                  If data_cli.Recordset("cl_codconv") = "SMIN" Or data_cli.Recordset("cl_codconv") = "SMINR" Or _
                     data_cli.Recordset("cl_codconv") = "SMINA" Then
                  Else
'                     data_conv.Recordset.FindFirst "cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                     data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                     data_conv.Refresh
                     If data_conv.Recordset.RecordCount > 0 Then
                        If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                           data_mut.RecordSource = "Select * from smi where cednum =" & Int(data_cli.Recordset("cl_cedula"))
                           data_mut.Refresh
'                           data_mut.Recordset.FindFirst "cednum =" & Int(data_cli.Recordset("cl_cedula"))
                           If data_mut.Recordset.RecordCount > 0 Then
                           Else
                              data_infno.Recordset.AddNew
                              data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                              data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                              data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                              data_infno.Recordset("cl_nombre") = "BAJA"
                              data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                              data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                              data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                              data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                              data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                              data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                              data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                              data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                              data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                              data_infno.Recordset.Update
                           End If
                        End If
                     End If
                  End If
               Else
                  data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                  data_conv.Refresh
                  If data_conv.Recordset.RecordCount > 0 Then
                     If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                        data_infno.Recordset.AddNew
                        data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                        data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                        data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                        data_infno.Recordset("cl_nombre") = "BAJA"
                        data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                        data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                        data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                        data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                        data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                        data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                        data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                        data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                        data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                        data_infno.Recordset.Update
                     End If
                  End If
               End If
            Else
               data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
               data_conv.Refresh
               If data_conv.Recordset.RecordCount > 0 Then
                  If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                     data_infno.Recordset.AddNew
                     data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                     data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                     data_infno.Recordset("cl_nombre") = "BAJA"
                     data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                     data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                     data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                     data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                     data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                     data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                     data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                     data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                     data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                     data_infno.Recordset.Update
                  End If
               End If
            End If
         End If
         data_cli.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      Label3.Visible = False
      Label3.Caption = ""
      DoEvents
      
      data_inf.RecordSource = "Select * from infcli where cl_nombre ='" & "NO ESTA EN P.SAPP" & "' order by cl_apellid"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "ALTAS"
         Xlibexel.SaveAs ("C:\planillas\" & "SMI-Altas" & ".xls")
         Xarchtex = "C:\planillas\SMI-Altas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE ALTAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONV.A INGRESAR"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("I" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_inf.Recordset.EOF
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_nomvend")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nomvend")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_celular")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_celular")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_fnac")) = False Then
               Xarchexel.Cells(Xlin, XCol) = CDate(data_inf.Recordset("cl_fnac"))
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nom_sup")
            Else
               Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("info_debit")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("info_debit")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
            End If
            data_inf.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      Else
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "ALTAS"
         Xlibexel.SaveAs ("C:\planillas\" & "SMI-Altas" & ".xls")
         Xarchtex = "C:\planillas\SMI-Altas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE ALTAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONV.A INGRESAR"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("I" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      
      End If
      
      data_inf.RecordSource = "Select * from infcli where cl_nombre not in ('NO ESTA EN P.SAPP','ACTIVO') order by cl_apellid"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1
         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "MODIF"
         Xlibexel.SaveAs ("C:\planillas\SMI-Mod.xls")
         Xarchtex = "C:\planillas\SMI-Mod.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE MODIFICACIONES MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "J" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 20
         Xarchexel.Cells(Xlin, XCol) = "MODIFICACION"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONV.ACTUAL"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONV.A INGRESAR"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("I" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("J" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("K" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("L" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_inf.Recordset.EOF
            If IsNull(data_inf.Recordset("cl_nombre")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nombre")
            Else
               Xarchexel.Cells(Xlin, XCol) = "MODIF"
            End If
            XCol = XCol + 1
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_codconv")
            
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_nomvend")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nomvend")
            End If
            XCol = XCol + 1
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_codigo")
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_apellid")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
            Else
               Xarchexel.Cells(Xlin, XCol) = "NN"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_celular")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_celular")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_fnac")) = False Then
               Xarchexel.Cells(Xlin, XCol) = CDate(data_inf.Recordset("cl_fnac"))
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nom_sup")
            Else
               Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("info_debit")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("info_debit")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
            End If
            data_inf.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      End If
      
      data_infno.RecordSource = "Select * from infno where cl_nombre in ('BAJA') order by cl_apellid"
      data_infno.Refresh
      If data_infno.Recordset.RecordCount > 0 Then
         data_infno.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "BAJAS"
         Xlibexel.SaveAs ("C:\planillas\SMI-Bajas.xls")
         Xarchtex = "C:\planillas\SMI-Bajas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE BAJAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRES"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_infno.Recordset.EOF
            If IsNull(data_infno.Recordset("cl_codigo")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_codigo")
            Else
               Xarchexel.Cells(Xlin, XCol) = "0"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_apellid")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_apellid")
            Else
               Xarchexel.Cells(Xlin, XCol) = "NN"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_cedula")) = False Then
               If IsNull(data_infno.Recordset("cl_codced")) = False Then
                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_infno.Recordset("cl_cedula"))) & "-" & Trim(str(data_infno.Recordset("cl_codced")))
               Else
                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_infno.Recordset("cl_cedula"))) & "-0"
               End If
            Else
               Xarchexel.Cells(Xlin, XCol) = "0-0"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_codconv")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_codconv")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_zona")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_zona")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_direcci")
            End If
            data_infno.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      End If
      MsgBox "Proceso terminado"
   End If

End If
frm_ctrolmut.MousePointer = 0

End Sub

Private Sub Command4_Click()
frm_ctrolmut.MousePointer = 11
'CASH
Dim Xcedeva As Long

Dim Xobjexel As Excel.Application
Dim Xlibexel As Excel.Workbook
Dim Xarchexel As New Excel.Worksheet

Dim XCol, Xlin, Xnrocan, Xcolfija As Long
Dim Xarchtex As String
Dim Xlabrir As New Excel.Application
Dim XsiAltas As String

If Combo1.Text = "CASH" Then 'ok
   data_mut.RecordSource = "cash"
   data_mut.Refresh
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         data_mut.Recordset.Delete
         data_mut.Recordset.MoveNext
      Loop
   End If
   Data1.DatabaseName = "C:\mutuales\cash.xls"
   Data1.RecordSource = "socios$"
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         data_mut.Recordset.AddNew
         data_mut.Recordset("ced") = Data1.Recordset("ced")
         data_mut.Recordset("nombre") = Data1.Recordset("nombre")
         data_mut.Recordset("zona") = Data1.Recordset("zona")
         If IsNull(Data1.Recordset("fnac")) = False Then
            data_mut.Recordset("fnac") = Data1.Recordset("fnac")
         End If
         If IsNull(Data1.Recordset("categ")) = False Then
            data_mut.Recordset("categ") = Mid(Data1.Recordset("categ"), 1, 255)
         End If
         If IsNull(Data1.Recordset("domicilio")) = False Then
            data_mut.Recordset("domicilio") = Mid(Data1.Recordset("domicilio"), 1, 255)
         End If
         If IsNull(Data1.Recordset("celular")) = False Then
            data_mut.Recordset("celular") = Mid(Data1.Recordset("celular"), 1, 255)
         End If
         If IsNull(Data1.Recordset("correo")) = False Then
            data_mut.Recordset("correo") = Mid(Data1.Recordset("correo"), 1, 255)
         End If
         data_mut.Recordset.Update
         Data1.Recordset.MoveNext
      Loop
   End If
   data_mut.Refresh
   Label3.Visible = True
   Label3.Caption = "Procesando Altas/Modif"
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveLast
      DoEvents
      pb.Visible = True
      pb.Max = data_mut.Recordset.RecordCount
      pb.Value = 0
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         If IsNull(data_mut.Recordset("ced")) = False Then
            Xcedeva = Val(data_mut.Recordset("ced"))
            data_mut.Recordset.Edit
            data_mut.Recordset("cednum") = Xcedeva
            data_mut.Recordset.Update
         Else
            Xcedeva = 0
         End If
         If Xcedeva > 0 Then
'            data_cli.Recordset.FindFirst "cl_cedula =" & Xcedeva
            data_cli.RecordSource = "Select * from clientes where cl_cedula =" & Xcedeva
            data_cli.Refresh
            If data_cli.Recordset.RecordCount > 0 Then
               data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
               data_conv.Refresh
               If data_conv.Recordset.RecordCount > 0 Then
                  If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                     If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                        data_inf.Recordset.AddNew
                        data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                        data_inf.Recordset("cl_celular") = data_mut.Recordset("ced")
                        If IsNull(data_mut.Recordset("domicilio")) = False Then
                           data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                        End If
                        If IsNull(data_mut.Recordset("categ")) = False Then
                           data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                        End If
                        If IsNull(data_mut.Recordset("celular")) = False Then
                           data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("celular"))
                        End If
                        If IsNull(data_mut.Recordset("zona")) = False Then
                           data_inf.Recordset("cl_telefon") = Trim(Mid(data_mut.Recordset("zona"), 1, 20))
                        End If
                        data_inf.Recordset("cl_apellid") = Mid(data_mut.Recordset("nombre"), 1, 60)
                        If IsNull(data_mut.Recordset("correo")) = False Then
                           data_inf.Recordset("cl_entre") = Trim(data_mut.Recordset("correo"))
                        End If
                        data_inf.Recordset("cl_nombre") = "REACTIVAR"
                        data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                        data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                        data_inf.Recordset.Update
                     End If
                  Else
                     data_inf.Recordset.AddNew
                    data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                    data_inf.Recordset("cl_celular") = data_mut.Recordset("ced")
                    If IsNull(data_mut.Recordset("domicilio")) = False Then
                       data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                    End If
                    If IsNull(data_mut.Recordset("categ")) = False Then
                       data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                    End If
                    If IsNull(data_mut.Recordset("celular")) = False Then
                       data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("celular"))
                    End If
                    If IsNull(data_mut.Recordset("zona")) = False Then
                       data_inf.Recordset("cl_telefon") = Trim(Mid(data_mut.Recordset("zona"), 1, 20))
                    End If
                    data_inf.Recordset("cl_apellid") = Mid(data_mut.Recordset("nombre"), 1, 60)
                    If IsNull(data_mut.Recordset("correo")) = False Then
                       data_inf.Recordset("cl_entre") = Trim(data_mut.Recordset("correo"))
                    End If
                     If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                        data_inf.Recordset("cl_nombre") = "CONV.INCORRECTO BAJ"
                     Else
                        data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                     End If
                     data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                     data_inf.Recordset.Update
                  End If
               Else
                  data_inf.Recordset.AddNew
                  data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                  data_inf.Recordset("cl_celular") = data_mut.Recordset("ced")
                  If IsNull(data_mut.Recordset("domicilio")) = False Then
                     data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                  End If
                  If IsNull(data_mut.Recordset("categ")) = False Then
                     data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                  End If
                  If IsNull(data_mut.Recordset("celular")) = False Then
                     data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("celular"))
                  End If
                  If IsNull(data_mut.Recordset("zona")) = False Then
                     data_inf.Recordset("cl_telefon") = Trim(Mid(data_mut.Recordset("zona"), 1, 20))
                  End If
                  data_inf.Recordset("cl_apellid") = Mid(data_mut.Recordset("nombre"), 1, 60)
                  If IsNull(data_mut.Recordset("correo")) = False Then
                     data_inf.Recordset("cl_entre") = Trim(data_mut.Recordset("correo"))
                  End If
                  If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                     data_inf.Recordset("cl_nombre") = "CONV.INCORRECTO BAJA"
                  Else
                     data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                  End If
                  data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                  data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                  data_inf.Recordset.Update
               End If
            Else
               data_inf.Recordset.AddNew
                data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                data_inf.Recordset("cl_celular") = data_mut.Recordset("ced")
                If IsNull(data_mut.Recordset("domicilio")) = False Then
                   data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                End If
                If IsNull(data_mut.Recordset("categ")) = False Then
                   data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                End If
                If IsNull(data_mut.Recordset("celular")) = False Then
                   data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("celular"))
                End If
                If IsNull(data_mut.Recordset("zona")) = False Then
                   data_inf.Recordset("cl_telefon") = Trim(Mid(data_mut.Recordset("zona"), 1, 20))
                End If
                data_inf.Recordset("cl_apellid") = Mid(data_mut.Recordset("nombre"), 1, 60)
                If IsNull(data_mut.Recordset("correo")) = False Then
                   data_inf.Recordset("cl_entre") = Trim(data_mut.Recordset("correo"))
                End If
               
               data_inf.Recordset("cl_nombre") = "NO ESTA EN P.SAPP"
               data_inf.Recordset.Update
            End If
         End If
         data_mut.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      DoEvents
      
      data_cli.RecordSource = "Select * from clientes where estado <>" & 2 & " And estado <>" & 3 & " and cl_codconv in ('CASH')"
      data_cli.Refresh
      data_cli.Recordset.MoveLast
      data_cli.Recordset.MoveFirst
      pb.Max = pb.Max + data_cli.Recordset.RecordCount
      data_mut.Refresh
      Label3.Caption = "Procesando BAJAS..."
      DoEvents
      Do While Not data_cli.Recordset.EOF
         If IsNull(data_cli.Recordset("cl_codconv")) = False Then
            If IsNull(data_cli.Recordset("cl_cedula")) = False Then
               If data_cli.Recordset("cl_cedula") > 0 Then
                  'If data_cli.Recordset("cl_codconv") = "UNIVS" Or data_cli.Recordset("cl_codconv") = "UNIVNR" Or _
                  '   data_cli.Recordset("cl_codconv") = "UNNSAM" Then
                  'Else
'                     data_conv.Recordset.FindFirst "cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                     'data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                     'data_conv.Refresh
                     'If data_conv.Recordset.RecordCount > 0 Then
                        'If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                           data_mut.RecordSource = "Select * from cash where cednum =" & Int(data_cli.Recordset("cl_cedula"))
                           data_mut.Refresh
'                           data_mut.Recordset.FindFirst "cednum =" & Int(data_cli.Recordset("cl_cedula"))
                           If data_mut.Recordset.RecordCount > 0 Then
                           Else
                              data_infno.Recordset.AddNew
                              data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                              data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                              data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                              data_infno.Recordset("cl_nombre") = "BAJA"
                              data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                              data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                              data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                              data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                              data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                              data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                              data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                              data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                              data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                              data_infno.Recordset.Update
                           End If
                        'End If
                     'End If
                  'End If
               Else
                  'data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                  'data_conv.Refresh
                  'If data_conv.Recordset.RecordCount > 0 Then
                  '   If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                        data_infno.Recordset.AddNew
                        data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                        data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                        data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                        data_infno.Recordset("cl_nombre") = "BAJA"
                        data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                        data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                        data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                        data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                        data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                        data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                        data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                        data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                        data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                        data_infno.Recordset.Update
                     'End If
                  'End If
               End If
            Else
               'data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
               'data_conv.Refresh
               'If data_conv.Recordset.RecordCount > 0 Then
               '   If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                     data_infno.Recordset.AddNew
                     data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                     data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                     data_infno.Recordset("cl_nombre") = "BAJA"
                     data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                     data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                     data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                     data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                     data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                     data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                     data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                     data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                     data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                     data_infno.Recordset.Update
                '  End If
               'End If
            End If
         End If
         data_cli.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      Label3.Visible = False
      Label3.Caption = ""
      DoEvents
      
      data_inf.RecordSource = "Select * from infcli where cl_nombre ='" & "NO ESTA EN P.SAPP" & "' order by cl_apellid"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "ALTAS"
         Xlibexel.SaveAs ("C:\planillas\" & "Cash-Altas" & ".xls")
         Xarchtex = "C:\planillas\Cash-Altas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE ALTAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 20
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_inf.Recordset.EOF
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_celular")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_celular")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_fnac")) = False Then
               Xarchexel.Cells(Xlin, XCol) = CDate(data_inf.Recordset("cl_fnac"))
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nom_sup")
            Else
               Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("info_debit")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("info_debit")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_entre")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_entre")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
            End If
            data_inf.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      Else
         XCol = 1
         Xlin = 1
         Xnrocan = 1
         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "ALTAS"
         Xlibexel.SaveAs ("C:\planillas\" & "Cash-Altas" & ".xls")
         Xarchtex = "C:\planillas\Cash-Altas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE ALTAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 20
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      
      End If
      
      data_inf.RecordSource = "Select * from infcli where cl_nombre not in ('NO ESTA EN P.SAPP','ACTIVO') order by cl_apellid"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1
         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "MODIF"
         Xlibexel.SaveAs ("C:\planillas\Cash-Mod.xls")
         Xarchtex = "C:\planillas\Cash-Mod.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE MODIFICACIONES MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "J" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 20
         Xarchexel.Cells(Xlin, XCol) = "MODIFICACION"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 20
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("I" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("J" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("K" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_inf.Recordset.EOF
            If IsNull(data_inf.Recordset("cl_nombre")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nombre")
            Else
               Xarchexel.Cells(Xlin, XCol) = "MODIF"
            End If
            XCol = XCol + 1
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_codconv")
            XCol = XCol + 1
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_codigo")
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_apellid")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
            Else
               Xarchexel.Cells(Xlin, XCol) = "NN"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_celular")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_celular")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_fnac")) = False Then
               Xarchexel.Cells(Xlin, XCol) = CDate(data_inf.Recordset("cl_fnac"))
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nom_sup")
            Else
               Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("info_debit")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("info_debit")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_entre")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_entre")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
            End If
            data_inf.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      End If
      
      data_infno.RecordSource = "Select * from infno where cl_nombre in ('BAJA') order by cl_apellid"
      data_infno.Refresh
      If data_infno.Recordset.RecordCount > 0 Then
         Dim SiDarBajas As String
         
         data_infno.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "BAJAS"
         Xlibexel.SaveAs ("C:\planillas\Cash-Bajas.xls")
         Xarchtex = "C:\planillas\Cash-Bajas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE BAJAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRES"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         SiDarBajas = MsgBox("Desea dar Bajas automáticas en el sistema?", vbInformation + vbYesNo)
         
         Do While Not data_infno.Recordset.EOF
            If IsNull(data_infno.Recordset("cl_codigo")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_codigo")
            Else
               Xarchexel.Cells(Xlin, XCol) = "0"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_apellid")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_apellid")
            Else
               Xarchexel.Cells(Xlin, XCol) = "NN"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_cedula")) = False Then
               If IsNull(data_infno.Recordset("cl_codced")) = False Then
                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_infno.Recordset("cl_cedula"))) & "-" & Trim(str(data_infno.Recordset("cl_codced")))
               Else
                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_infno.Recordset("cl_cedula"))) & "-0"
               End If
            Else
               Xarchexel.Cells(Xlin, XCol) = "0-0"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_codconv")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_codconv")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_zona")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_zona")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_direcci")
            End If
            If SiDarBajas = vbYes Then
               data_cli.RecordSource = "select * from clientes where cl_codigo =" & data_infno.Recordset("cl_codigo")
               data_cli.Refresh
               If data_cli.Recordset.RecordCount > 0 Then
                  If IsNull(data_cli.Recordset("fecha_baja")) = True Then
                     data_cli.Recordset("fecha_baja") = Date
                     data_cli.Recordset("estado") = 2
                     data_cli.Recordset.Update
                     data_cli.RecordSource = "select * from abmsocio where cl_codigo =" & data_infno.Recordset("cl_codigo")
                     data_cli.Refresh
                     data_cli.Recordset.AddNew
                     data_cli.Recordset("usuario") = WElusuario
                     data_cli.Recordset("fecha") = Date
                     data_cli.Recordset("hora") = Format(Time, "HH:mm")
                     data_cli.Recordset("cl_codigo") = data_infno.Recordset("cl_codigo")
                     data_cli.Recordset("desc") = "BAJA"
                     data_cli.Recordset("cl_motivo") = "SIN DATOS"
                     data_cli.Recordset("convenio") = data_infno.Recordset("cl_codconv")
                     data_cli.Recordset("base") = frm_menu.data_parse.Recordset("base")
                     data_cli.Recordset.Update
                  End If
               End If
            End If
            data_infno.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      End If
      Dim XlacedCas As Long
      Dim XeldvCas As Integer
      Dim Xlanuevamat As Long
      Dim Xconteo As Integer
      Dim Xladire As String
   
      XsiAltas = MsgBox("Confirma que desea agregar las ALTAS en forma automática al sistema?", vbInformation + vbYesNo)
      If XsiAltas = vbYes Then
         frm_ctrolmut.MousePointer = 11
         data_inf.RecordSource = "Select * from infcli where cl_nombre ='" & "NO ESTA EN P.SAPP" & "' order by cl_apellid"
         data_inf.Refresh
         Xconteo = 0
         If data_inf.Recordset.RecordCount > 0 Then
            data_inf.Recordset.MoveFirst
            Do While Not data_inf.Recordset.EOF
               If Len(Trim(data_inf.Recordset("cl_celular"))) = 9 Then
                  XlacedCas = Val(Mid(Trim(data_inf.Recordset("cl_celular")), 1, 7))
                  XeldvCas = Val(Mid(Trim(data_inf.Recordset("cl_celular")), 9, 1))
               Else
                  If Len(Trim(data_inf.Recordset("cl_celular"))) = 8 Then
                     XlacedCas = Val(Mid(Trim(data_inf.Recordset("cl_celular")), 1, 6))
                     XeldvCas = Val(Mid(Trim(data_inf.Recordset("cl_celular")), 8, 1))
                  Else
                     XlacedCas = Val(Mid(Trim(data_inf.Recordset("cl_celular")), 1, 7))
                     XeldvCas = Val(Mid(Trim(data_inf.Recordset("cl_celular")), 9, 1))
                  End If
               End If
               data_cli.RecordSource = "select * from clientes where cl_cedula =" & XlacedCas
               data_cli.Refresh
               If data_cli.Recordset.RecordCount > 0 Then
               
               Else
                  Xlanuevamat = data_paramnew.Recordset("p_matric") + 1
                  data_paramnew.Recordset.Edit
                  data_paramnew.Recordset("p_matric") = data_paramnew.Recordset("p_matric") + 1
                  data_paramnew.Recordset.Update
                  data_paramnew.Refresh
                  
                  data_cli.Recordset.AddNew
                  data_cli.Recordset("cl_codigo") = Xlanuevamat
                  data_cli.Recordset("estado") = 1
                  data_cli.Recordset("cl_codconv") = "CASH"
                  data_cli.Recordset("cl_nomconv") = "CONVENIO COOPERATIVA CASH"
                  data_cli.Recordset("cl_apellid") = data_inf.Recordset("cl_apellid")
                  Xladire = Trim(data_inf.Recordset("cl_direcci") & " " & data_inf.Recordset("cl_telefon"))
                  data_cli.Recordset("cl_direcci") = Mid(Xladire, 1, 80)
                  If IsNull(data_inf.Recordset("info_debit")) = False Then
                     data_cli.Recordset("cl_telefon") = Mid(data_inf.Recordset("info_debit"), 1, 20)
                     data_cli.Recordset("cl_entre") = Mid(data_inf.Recordset("info_debit"), 1, 80)
                  End If
                  If IsNull(data_inf.Recordset("cl_fnac")) = False Then
                     data_cli.Recordset("cl_fnac") = data_inf.Recordset("cl_fnac")
                  End If
                  data_cli.Recordset("cl_cedula") = XlacedCas
                  data_cli.Recordset("cl_codced") = XeldvCas
                  data_cli.Recordset("cl_nrovend") = 785
                  data_cli.Recordset("cl_nomvend") = "CASH"
                  data_cli.Recordset("cl_fecing") = Date
                  data_cli.Recordset("cl_forpago") = 1
                  data_cli.Recordset("cl_descpag") = "Abono Mensual"
                  data_cli.Recordset("cl_sexo") = 1
                  data_cli.Recordset("fecha_sys") = Date
                  data_cli.Recordset.Update
                  Xconteo = Xconteo + 1
               End If
               data_inf.Recordset.MoveNext
            Loop
            frm_ctrolmut.MousePointer = 0
            MsgBox "Proceso de altas automáticas terminado. Se ingresaron " & Xconteo & " registros.", vbExclamation
         End If
      End If
      Dim XsiReactivar As String
      XsiReactivar = MsgBox("Confirma que desea REACTIVAR en forma automática?", vbInformation + vbYesNo)
      If XsiReactivar = vbYes Then
         frm_ctrolmut.MousePointer = 11
         Xconteo = 0
         data_inf.RecordSource = "Select * from infcli where cl_nombre ='" & "REACTIVAR" & "' order by cl_apellid"
         data_inf.Refresh
         Xconteo = 0
         If data_inf.Recordset.RecordCount > 0 Then
            data_inf.Recordset.MoveFirst
            Do While Not data_inf.Recordset.EOF
               If Len(Trim(data_inf.Recordset("cl_celular"))) = 9 Then
                  XlacedCas = Val(Mid(Trim(data_inf.Recordset("cl_celular")), 1, 7))
                  XeldvCas = Val(Mid(Trim(data_inf.Recordset("cl_celular")), 9, 1))
               Else
                  If Len(Trim(data_inf.Recordset("cl_celular"))) = 8 Then
                     XlacedCas = Val(Mid(Trim(data_inf.Recordset("cl_celular")), 1, 6))
                     XeldvCas = Val(Mid(Trim(data_inf.Recordset("cl_celular")), 8, 1))
                  Else
                     XlacedCas = Val(Mid(Trim(data_inf.Recordset("cl_celular")), 1, 7))
                     XeldvCas = Val(Mid(Trim(data_inf.Recordset("cl_celular")), 9, 1))
                  End If
               End If
               data_cli.RecordSource = "select * from clientes where cl_cedula =" & XlacedCas
               data_cli.Refresh
               If data_cli.Recordset.RecordCount > 0 Then
                  data_cli.Recordset("estado") = 1
                  data_cli.Recordset("fecha_baja") = Null
                  data_cli.Recordset("cl_codconv") = "CASH"
                  data_cli.Recordset("cl_nomconv") = "CONVENIO COOPERATIVA CASH"
                  Xladire = Trim(data_inf.Recordset("cl_direcci") & " " & data_inf.Recordset("cl_telefon"))
                  data_cli.Recordset("cl_direcci") = Mid(Xladire, 1, 80)
                  If IsNull(data_inf.Recordset("info_debit")) = False Then
                     data_cli.Recordset("cl_telefon") = Mid(data_inf.Recordset("info_debit"), 1, 20)
                     data_cli.Recordset("cl_entre") = Mid(data_inf.Recordset("info_debit"), 1, 80)
                  End If
                  data_cli.Recordset("cl_nrovend") = 785
                  data_cli.Recordset("cl_nomvend") = "CASH"
                  data_cli.Recordset("cl_fecing") = Date
                  data_cli.Recordset("cl_forpago") = 1
                  data_cli.Recordset("cl_descpag") = "Abono Mensual"
                  data_cli.Recordset("fecha_modi") = Date
                  data_cli.Recordset.Update
                  Xconteo = Xconteo + 1
               End If
               data_inf.Recordset.MoveNext
            Loop
            frm_ctrolmut.MousePointer = 0
            MsgBox "Proceso de modificaciones terminado. Se reactivaron " & Xconteo & " registros.", vbExclamation
         End If
                  
         frm_ctrolmut.MousePointer = 0
         MsgBox "Proceso terminado"
      End If
   End If
End If

''Universal
If Combo1.Text = "UNIVERSAL" Then 'ok
   data_mut.RecordSource = "univ"
   data_mut.Refresh
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         data_mut.Recordset.Delete
         data_mut.Recordset.MoveNext
      Loop
   End If
   Data1.DatabaseName = "C:\mutuales\univ.xls"
   Data1.RecordSource = "socios$"
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         data_mut.Recordset.AddNew
         data_mut.Recordset("ced") = Data1.Recordset("ced")
         data_mut.Recordset("mat") = Data1.Recordset("mat")
         data_mut.Recordset("dv") = Data1.Recordset("dv")
         data_mut.Recordset("nombre") = Data1.Recordset("nombre")
         data_mut.Recordset("sexo") = Data1.Recordset("sexo")
         If IsNull(Data1.Recordset("fnac")) = False Then
            data_mut.Recordset("fnac") = Data1.Recordset("fnac")
         End If
         If IsNull(Data1.Recordset("categ")) = False Then
            data_mut.Recordset("categ") = Mid(Data1.Recordset("categ"), 1, 255)
         End If
         If IsNull(Data1.Recordset("domicilio")) = False Then
            data_mut.Recordset("domicilio") = Mid(Data1.Recordset("domicilio"), 1, 255)
         End If
         If IsNull(Data1.Recordset("telefono")) = False Then
            data_mut.Recordset("telefono") = Mid(Data1.Recordset("telefono"), 1, 255)
         End If
         If IsNull(Data1.Recordset("celular")) = False Then
            data_mut.Recordset("celular") = Mid(Data1.Recordset("celular"), 1, 255)
         End If
         If IsNull(Data1.Recordset("correo")) = False Then
            data_mut.Recordset("correo") = Mid(Data1.Recordset("correo"), 1, 255)
         End If
         data_mut.Recordset.Update
         Data1.Recordset.MoveNext
      Loop
   End If
   data_mut.Refresh
   Label3.Visible = True
   Label3.Caption = "Procesando Altas/Modif"
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveLast
      DoEvents
      pb.Visible = True
      pb.Max = data_mut.Recordset.RecordCount
      pb.Value = 0
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         If IsNull(data_mut.Recordset("ced")) = False Then
            Xcedeva = data_mut.Recordset("ced")
            data_mut.Recordset.Edit
            data_mut.Recordset("cednum") = Xcedeva
            data_mut.Recordset.Update
         Else
            Xcedeva = 0
         End If
         If Xcedeva > 0 Then
            data_cli.RecordSource = "Select * from clientes where cl_cedula =" & Xcedeva
            data_cli.Refresh
            If data_cli.Recordset.RecordCount > 0 Then
               data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
               data_conv.Refresh
               If data_conv.Recordset.RecordCount > 0 Then
                  If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                     If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                        data_inf.Recordset.AddNew
                        data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                        data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced"))) & "-" & Trim(str(data_mut.Recordset("dv")))
                        If IsNull(data_mut.Recordset("domicilio")) = False Then
                           data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                        End If
                        If IsNull(data_mut.Recordset("categ")) = False Then
                           data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                        End If
                        If IsNull(data_mut.Recordset("celular")) = False Then
                           data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                        End If
                        If IsNull(data_mut.Recordset("telefono")) = False Then
                           data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                        End If
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("nombre")
                        If IsNull(data_mut.Recordset("correo")) = False Then
                           data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                        End If
                        data_inf.Recordset("cl_nombre") = "REACTIVAR"
                        data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                        data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                        data_inf.Recordset.Update
                     Else
                        If data_conv.Recordset("cnv_codigo") = "UNIVS" Or data_conv.Recordset("cnv_codigo") = "UNIVNR" Or _
                           data_conv.Recordset("cnv_codigo") = "UNNSAM" Then
                            data_inf.Recordset.AddNew
                            data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                            data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced"))) & "-" & Trim(str(data_mut.Recordset("dv")))
                            If IsNull(data_mut.Recordset("domicilio")) = False Then
                               data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                            End If
                            If IsNull(data_mut.Recordset("categ")) = False Then
                               data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                            End If
                            If IsNull(data_mut.Recordset("celular")) = False Then
                               data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                            End If
                            If IsNull(data_mut.Recordset("telefono")) = False Then
                               data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                            End If
                            data_inf.Recordset("cl_apellid") = data_mut.Recordset("nombre")
                            If IsNull(data_mut.Recordset("correo")) = False Then
                               data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                            End If
                            data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                            data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                            data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                            data_inf.Recordset.Update
                        End If
                     End If
                  Else
                     data_inf.Recordset.AddNew
                     data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                     data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced"))) & "-" & Trim(str(data_mut.Recordset("dv")))
                     If IsNull(data_mut.Recordset("domicilio")) = False Then
                        data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                     End If
                     If IsNull(data_mut.Recordset("categ")) = False Then
                        data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                     End If
                     If IsNull(data_mut.Recordset("celular")) = False Then
                        data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                     End If
                     If IsNull(data_mut.Recordset("telefono")) = False Then
                        data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                     End If
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("nombre")
                     If IsNull(data_mut.Recordset("correo")) = False Then
                        data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                     End If
                     If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                        data_inf.Recordset("cl_nombre") = "CONV.INCORRECTO BAJ"
                     Else
                        data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                     End If
                     data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                     data_inf.Recordset.Update
                  End If
               Else
                  data_inf.Recordset.AddNew
                  data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                  data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced"))) & "-" & Trim(str(data_mut.Recordset("dv")))
                  If IsNull(data_mut.Recordset("domicilio")) = False Then
                     data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                  End If
                  If IsNull(data_mut.Recordset("categ")) = False Then
                     data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                  End If
                  If IsNull(data_mut.Recordset("celular")) = False Then
                     data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                  End If
                  If IsNull(data_mut.Recordset("telefono")) = False Then
                     data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                  End If
                  data_inf.Recordset("cl_apellid") = data_mut.Recordset("nombre")
                  If IsNull(data_mut.Recordset("correo")) = False Then
                     data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                  End If
                  If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                     data_inf.Recordset("cl_nombre") = "CONV.INCORRECTO BAJA"
                  Else
                     data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                  End If
                  data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                  data_inf.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                  data_inf.Recordset.Update
               End If
            Else
               data_inf.Recordset.AddNew
               data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
               data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced"))) & "-" & Trim(str(data_mut.Recordset("dv")))
               If IsNull(data_mut.Recordset("domicilio")) = False Then
                  data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
               End If
               If IsNull(data_mut.Recordset("categ")) = False Then
                  data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
               End If
               If IsNull(data_mut.Recordset("celular")) = False Then
                  data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
               End If
               If IsNull(data_mut.Recordset("telefono")) = False Then
                  data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
               End If
               data_inf.Recordset("cl_apellid") = data_mut.Recordset("nombre")
               If IsNull(data_mut.Recordset("correo")) = False Then
                  data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
               End If
               data_inf.Recordset("cl_nombre") = "NO ESTA EN P.SAPP"
               data_inf.Recordset.Update
            End If
         End If
         data_mut.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      DoEvents
      data_cli.RecordSource = "Select * from clientes where estado <>" & 2 & " And estado <>" & 3 & " and cl_codconv not in ('UDEMM','PART','SMIN','HEVAN','EMERN','CASH','MSP','CAAMEP','UCM','CCNOS','HEVANO')"
      data_cli.Refresh
      data_cli.Recordset.MoveLast
      data_cli.Recordset.MoveFirst
      pb.Max = pb.Max + data_cli.Recordset.RecordCount
      data_mut.Refresh
      Label3.Caption = "Procesando BAJAS..."
      DoEvents
      Do While Not data_cli.Recordset.EOF
         If IsNull(data_cli.Recordset("cl_codconv")) = False Then
            If IsNull(data_cli.Recordset("cl_cedula")) = False Then
               If data_cli.Recordset("cl_cedula") > 0 Then
                  If data_cli.Recordset("cl_codconv") = "UNIVS" Or data_cli.Recordset("cl_codconv") = "UNIVNR" Or _
                     data_cli.Recordset("cl_codconv") = "UNNSAM" Then
                  Else
'                     data_conv.Recordset.FindFirst "cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                     data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                     data_conv.Refresh
                     If data_conv.Recordset.RecordCount > 0 Then
                        If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                           data_mut.RecordSource = "Select * from univ where cednum =" & Int(data_cli.Recordset("cl_cedula"))
                           data_mut.Refresh
'                           data_mut.Recordset.FindFirst "cednum =" & Int(data_cli.Recordset("cl_cedula"))
                           If data_mut.Recordset.RecordCount > 0 Then
                           Else
                              data_infno.Recordset.AddNew
                              data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                              data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                              data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                              data_infno.Recordset("cl_nombre") = "BAJA"
                              data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                              data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                              data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                              data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                              data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                              data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                              data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                              data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                              data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                              data_infno.Recordset.Update
                           End If
                        End If
                     End If
                  End If
               Else
                  data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                  data_conv.Refresh
                  If data_conv.Recordset.RecordCount > 0 Then
                     If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                        data_infno.Recordset.AddNew
                        data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                        data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                        data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                        data_infno.Recordset("cl_nombre") = "BAJA"
                        data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                        data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                        data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                        data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                        data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                        data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                        data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                        data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                        data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                        data_infno.Recordset.Update
                     End If
                  End If
               End If
            Else
               data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
               data_conv.Refresh
               If data_conv.Recordset.RecordCount > 0 Then
                  If data_conv.Recordset("cnv_grupo") = Combo1.Text Then
                     data_infno.Recordset.AddNew
                     data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                     data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                     data_infno.Recordset("cl_nombre") = "BAJA"
                     data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                     data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                     data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                     data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                     data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                     data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                     data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                     data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                     data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                     data_infno.Recordset.Update
                  End If
               End If
            End If
         End If
         data_cli.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      Label3.Visible = False
      Label3.Caption = ""
      DoEvents
      
      data_inf.RecordSource = "Select * from infcli where cl_nombre ='" & "NO ESTA EN P.SAPP" & "' order by cl_apellid"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "ALTAS"
         Xlibexel.SaveAs ("C:\planillas\" & "Univ-Altas" & ".xls")
         Xarchtex = "C:\planillas\Univ-Altas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE ALTAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_inf.Recordset.EOF
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_celular")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_celular")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_fnac")) = False Then
               Xarchexel.Cells(Xlin, XCol) = CDate(data_inf.Recordset("cl_fnac"))
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nom_sup")
            Else
               Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("info_debit")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("info_debit")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
            End If
            data_inf.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      Else
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "ALTAS"
         Xlibexel.SaveAs ("C:\planillas\" & "Univ-Altas" & ".xls")
         Xarchtex = "C:\planillas\Univ-Altas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE ALTAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      
      End If
      
      data_inf.RecordSource = "Select * from infcli where cl_nombre not in ('NO ESTA EN P.SAPP','ACTIVO') order by cl_apellid"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1
         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "MODIF"
         Xlibexel.SaveAs ("C:\planillas\Univ-Mod.xls")
         Xarchtex = "C:\planillas\Univ-Mod.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE MODIFICACIONES MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "J" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 20
         Xarchexel.Cells(Xlin, XCol) = "MODIFICACION"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("I" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("J" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("K" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_inf.Recordset.EOF
            If IsNull(data_inf.Recordset("cl_nombre")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nombre")
            Else
               Xarchexel.Cells(Xlin, XCol) = "MODIF"
            End If
            XCol = XCol + 1
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_codconv")
            XCol = XCol + 1
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_codigo")
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_apellid")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
            Else
               Xarchexel.Cells(Xlin, XCol) = "NN"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_celular")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_celular")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_fnac")) = False Then
               Xarchexel.Cells(Xlin, XCol) = CDate(data_inf.Recordset("cl_fnac"))
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nom_sup")
            Else
               Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("info_debit")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("info_debit")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
            End If
            data_inf.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      End If
      
      data_infno.RecordSource = "Select * from infno where cl_nombre in ('BAJA') order by cl_apellid"
      data_infno.Refresh
      If data_infno.Recordset.RecordCount > 0 Then
         data_infno.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "BAJAS"
         Xlibexel.SaveAs ("C:\planillas\Univ-Bajas.xls")
         Xarchtex = "C:\planillas\Univ-Bajas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE BAJAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRES"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_infno.Recordset.EOF
            If IsNull(data_infno.Recordset("cl_codigo")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_codigo")
            Else
               Xarchexel.Cells(Xlin, XCol) = "0"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_apellid")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_apellid")
            Else
               Xarchexel.Cells(Xlin, XCol) = "NN"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_cedula")) = False Then
               If IsNull(data_infno.Recordset("cl_codced")) = False Then
                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_infno.Recordset("cl_cedula"))) & "-" & Trim(str(data_infno.Recordset("cl_codced")))
               Else
                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_infno.Recordset("cl_cedula"))) & "-0"
               End If
            Else
               Xarchexel.Cells(Xlin, XCol) = "0-0"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_codconv")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_codconv")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_zona")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_zona")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_direcci")
            End If
            data_infno.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      End If
   End If
End If
frm_ctrolmut.MousePointer = 0

MsgBox "Proceso terminado"

End Sub

Private Sub Form_Load()
'data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_cli.RecordSource = "clientes"
'data_cli.Refresh
'data_mut.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\mutuales.mdb"
data_mut.DatabaseName = App.path & "\mutuales.mdb"
data_lin.ConnectionString = "dsn=" & Xconexrmt
'data_lin.RecordSource = "linmmdd"
'data_lin.Refresh
'data_inf.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"
data_inf.DatabaseName = App.path & "\informes.mdb"
data_inf.RecordSource = "infcli"
data_inf.Refresh
data_conv.ConnectionString = "dsn=" & Xconexrmt

data_paramnew.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_paramnew.RecordSource = "param_gral"
data_paramnew.Refresh

'data_conv.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_conv.RecordSource = "convenio"
'data_conv.Refresh


End Sub

Private Sub Form_Resize()
With Image1
     .Left = 0
     .Top = 0
     .Height = Me.Height
     .Width = Me.Width
End With

End Sub

Public Sub BlueCross()
Dim Xcedeva As Long
Dim Xcedblue, Xfnac As String

Dim Xobjexel As Excel.Application
Dim Xlibexel As Excel.Workbook
Dim Xarchexel As New Excel.Worksheet

Dim XCol, Xlin, Xnrocan, Xcolfija As Long
Dim Xarchtex As String
Dim Xlabrir As New Excel.Application

If Combo1.Text = "BLUE CROSS" Then 'OK
   data_mut.RecordSource = "blue"
   data_mut.Refresh
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         data_mut.Recordset.Delete
         data_mut.Recordset.MoveNext
      Loop
   End If
   Data1.DatabaseName = "C:\mutuales\bluecross.xls"
   Data1.RecordSource = "socios$"
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         data_mut.Recordset.AddNew
         If IsNull(Data1.Recordset("ced")) = False Then
            data_mut.Recordset("ced") = Val(Data1.Recordset("ced"))
         Else
            data_mut.Recordset("ced") = 0
         End If
         If IsNull(Data1.Recordset("nom1")) = False Then
            data_mut.Recordset("nom1") = Data1.Recordset("nom1")
         End If
         If IsNull(Data1.Recordset("fnac")) = False Then
            Xfnac = Trim(Mid(Data1.Recordset("fnac"), 7, 2)) & "/" & Trim(Mid(Data1.Recordset("fnac"), 5, 2)) & "/" & Trim(Mid(Data1.Recordset("fnac"), 1, 4))
            data_mut.Recordset("fnac") = Format(Xfnac, "dd/mm/yyyy")
         End If
         If IsNull(Data1.Recordset("domicilio")) = False Then
            data_mut.Recordset("domicilio") = Mid(Data1.Recordset("domicilio"), 1, 255)
         End If
         If IsNull(Data1.Recordset("zona")) = False Then
            data_mut.Recordset("zona") = Mid(Data1.Recordset("zona"), 1, 50)
         End If
         If IsNull(Data1.Recordset("telefono")) = False Then
            data_mut.Recordset("telefono") = Mid(Data1.Recordset("telefono"), 1, 255)
         End If
         If IsNull(Data1.Recordset("celular")) = False Then
            data_mut.Recordset("celular") = Mid(Data1.Recordset("celular"), 1, 255)
         End If
         If IsNull(Data1.Recordset("correo")) = False Then
            data_mut.Recordset("correo") = Mid(Data1.Recordset("correo"), 1, 255)
         End If
         data_mut.Recordset.Update
         Data1.Recordset.MoveNext
      Loop
   End If
   data_mut.Refresh
   
   Data1.RecordSource = "socios2$"
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         data_mut.Recordset.AddNew
         Xcedblue = Trim(str(Data1.Recordset("ced"))) & Trim(str(Data1.Recordset("dv")))
         data_mut.Recordset("ced") = Val(Xcedblue)
         If IsNull(Data1.Recordset("nom2")) = False Then
            If IsNull(Data1.Recordset("ape2")) = False Then
               data_mut.Recordset("nom1") = Data1.Recordset("ape1") & " " & Data1.Recordset("ape2") & " " & Data1.Recordset("nom1") & " " & Data1.Recordset("nom2")
            Else
               data_mut.Recordset("nom1") = Data1.Recordset("ape1") & " " & Data1.Recordset("nom1") & " " & Data1.Recordset("nom2")
            End If
         Else
            If IsNull(Data1.Recordset("ape2")) = False Then
               data_mut.Recordset("nom1") = Data1.Recordset("ape1") & " " & Data1.Recordset("ape2") & " " & Data1.Recordset("nom1")
            Else
               data_mut.Recordset("nom1") = Data1.Recordset("ape1") & " " & Data1.Recordset("nom1")
            End If
         End If
         If IsNull(Data1.Recordset("domicilio")) = False Then
            data_mut.Recordset("domicilio") = Mid(Data1.Recordset("domicilio"), 1, 255)
         End If
         If IsNull(Data1.Recordset("zona")) = False Then
            data_mut.Recordset("zona") = Mid(Data1.Recordset("zona"), 1, 50)
         End If
         If IsNull(Data1.Recordset("telefono")) = False Then
            data_mut.Recordset("telefono") = Mid(Data1.Recordset("telefono"), 1, 255)
         End If
         If IsNull(Data1.Recordset("celular")) = False Then
            data_mut.Recordset("celular") = Mid(Data1.Recordset("celular"), 1, 255)
         End If
         If IsNull(Data1.Recordset("correo")) = False Then
            data_mut.Recordset("correo") = Mid(Data1.Recordset("correo"), 1, 255)
         End If
         data_mut.Recordset.Update
         Data1.Recordset.MoveNext
      Loop
   End If
   data_mut.Refresh
   Label3.Visible = True
   Label3.Caption = "Procesando Altas/Modif"
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveLast
      DoEvents
      pb.Visible = True
      pb.Max = data_mut.Recordset.RecordCount
      pb.Value = 0
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         If IsNull(data_mut.Recordset("ced")) = False Then
            Xcedeva = data_mut.Recordset("ced")
            If Len(data_mut.Recordset("ced")) = 7 Then
               Xcedblue = Mid(Trim(str(data_mut.Recordset("ced"))), 1, 6)
            Else
               Xcedblue = Mid(Trim(str(data_mut.Recordset("ced"))), 1, 7)
            End If
            data_mut.Recordset.Edit
            data_mut.Recordset("cednum") = Val(Xcedblue)
            data_mut.Recordset.Update
            Xcedeva = Val(Xcedblue)
         Else
            Xcedeva = 0
         End If
         If Xcedeva > 0 Then
         
'            data_cli.Recordset.FindFirst "cl_cedula =" & Xcedeva
            data_cli.RecordSource = "Select * from clientes where cl_cedula =" & Xcedeva
            data_cli.Refresh
            If data_cli.Recordset.RecordCount > 0 Then
               data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
               data_conv.Refresh
               If data_conv.Recordset.RecordCount > 0 Then
                  If data_conv.Recordset("cnv_codigo") = "BLUE" Then
                     If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                        data_inf.Recordset.AddNew
                        If IsNull(data_mut.Recordset("fnac")) = False Then
                           data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                        End If
                        data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced")))
                        If IsNull(data_mut.Recordset("domicilio")) = False Then
                           data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                        End If
                        If IsNull(data_mut.Recordset("celular")) = False Then
                           data_inf.Recordset("cl_dpto") = Mid(Trim(data_mut.Recordset("celular")), 1, 12)
                        End If
                        If IsNull(data_mut.Recordset("telefono")) = False Then
                           data_inf.Recordset("cl_telefon") = Mid(Trim(data_mut.Recordset("telefono")), 1, 20)
                        End If
                        If IsNull(data_mut.Recordset("ape2")) = False Then
                           If IsNull(data_mut.Recordset("nom2")) = False Then
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                           Else
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                           End If
                        Else
                           If IsNull(data_mut.Recordset("nom2")) = False Then
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                           Else
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                           End If
                        End If
                        If IsNull(data_mut.Recordset("correo")) = False Then
                           data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                        End If
                        data_inf.Recordset("cl_nombre") = "REACTIVAR"
                        data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                        data_inf.Recordset("cl_zona") = data_mut.Recordset("zona")
                        data_inf.Recordset.Update
                     End If
                  Else
                     data_inf.Recordset.AddNew
                     If IsNull(data_mut.Recordset("fnac")) = False Then
                        data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                     End If
                     data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced")))
                     If IsNull(data_mut.Recordset("domicilio")) = False Then
                        data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                     End If
                     If IsNull(data_mut.Recordset("celular")) = False Then
                        data_inf.Recordset("cl_dpto") = Mid(Trim(data_mut.Recordset("celular")), 1, 12)
                     End If
                     If IsNull(data_mut.Recordset("telefono")) = False Then
                        data_inf.Recordset("cl_telefon") = Mid(Trim(data_mut.Recordset("telefono")), 1, 20)
                     End If
                     If IsNull(data_mut.Recordset("ape2")) = False Then
                        If IsNull(data_mut.Recordset("nom2")) = False Then
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                        Else
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                        End If
                     Else
                        If IsNull(data_mut.Recordset("nom2")) = False Then
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                        Else
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                        End If
                     End If
                     If IsNull(data_mut.Recordset("correo")) = False Then
                        data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                     End If
                     If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                        data_inf.Recordset("cl_nombre") = "CONV.INCORRECTO BAJ"
                     Else
                        data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                     End If
                     data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_inf.Recordset("cl_zona") = data_mut.Recordset("zona")
                     data_inf.Recordset.Update
                  End If
               End If
            Else
               data_inf.Recordset.AddNew
               If IsNull(data_mut.Recordset("fnac")) = False Then
                  data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
               End If
               data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced")))
               If IsNull(data_mut.Recordset("domicilio")) = False Then
                  data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
               End If
               If IsNull(data_mut.Recordset("celular")) = False Then
                  data_inf.Recordset("cl_dpto") = Mid(Trim(data_mut.Recordset("celular")), 1, 12)
               End If
               If IsNull(data_mut.Recordset("telefono")) = False Then
                  data_inf.Recordset("cl_telefon") = Mid(Trim(data_mut.Recordset("telefono")), 1, 20)
               End If
               If IsNull(data_mut.Recordset("ape2")) = False Then
                  If IsNull(data_mut.Recordset("nom2")) = False Then
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                  Else
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                  End If
               Else
                  If IsNull(data_mut.Recordset("nom2")) = False Then
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                  Else
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                  End If
               End If
               If IsNull(data_mut.Recordset("correo")) = False Then
                  data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
               End If
               data_inf.Recordset("cl_nombre") = "NO ESTA EN P.SAPP"
               data_inf.Recordset("cl_zona") = data_mut.Recordset("zona")
               data_inf.Recordset.Update
            End If
         End If
         data_mut.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      DoEvents
      data_cli.RecordSource = "Select * from clientes where estado not in (2,3) and cl_codconv in ('BLUE')"
      data_cli.Refresh
      data_cli.Recordset.MoveLast
      data_cli.Recordset.MoveFirst
      pb.Max = pb.Max + data_cli.Recordset.RecordCount
      data_mut.Refresh
      Label3.Caption = "Procesando BAJAS..."
      DoEvents
      Do While Not data_cli.Recordset.EOF
         If IsNull(data_cli.Recordset("cl_codconv")) = False Then
            If IsNull(data_cli.Recordset("cl_cedula")) = False Then
               If data_cli.Recordset("cl_cedula") > 0 Then
                  data_mut.RecordSource = "Select * from blue where cednum =" & Int(data_cli.Recordset("cl_cedula"))
                  data_mut.Refresh
                  If data_mut.Recordset.RecordCount > 0 Then
                  Else
                     data_infno.Recordset.AddNew
                     data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                     data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                     data_infno.Recordset("cl_nombre") = "BAJA"
                     data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                     data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                     data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                     data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                     data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                     data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                     data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                     data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                     data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                     data_infno.Recordset.Update
                  End If
               Else
                   data_infno.Recordset.AddNew
                   data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                   data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                   data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                   data_infno.Recordset("cl_nombre") = "BAJA"
                   data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                   data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                   data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                   data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                   data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                   data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                   data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                   data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                   data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                   data_infno.Recordset.Update
               End If
            End If
         End If
         data_cli.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      Label3.Visible = False
      Label3.Caption = ""
      DoEvents

      data_inf.RecordSource = "Select * from infcli where cl_nombre ='" & "NO ESTA EN P.SAPP" & "' order by cl_apellid"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "ALTAS"
         Xlibexel.SaveAs ("C:\planillas\" & "BlueCross-Altas" & ".xls")
         Xarchtex = "C:\planillas\BlueCross-Altas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE ALTAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_inf.Recordset.EOF
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_celular")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_celular")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_fnac")) = False Then
               Xarchexel.Cells(Xlin, XCol) = CDate(data_inf.Recordset("cl_fnac"))
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_zona")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_zona")
            Else
               Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("info_debit")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("info_debit")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
            End If
            data_inf.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      Else
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "ALTAS"
         Xlibexel.SaveAs ("C:\planillas\" & "BlueCross-Altas" & ".xls")
         Xarchtex = "C:\planillas\BlueCross-Altas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE ALTAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      
      End If
      
      data_inf.RecordSource = "Select * from infcli where cl_nombre not in ('NO ESTA EN P.SAPP','ACTIVO') order by cl_apellid"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1
         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         
         Xarchexel.Name = "MODIF"
         Xlibexel.SaveAs ("C:\planillas\BlueCross-Mod.xls")
         Xarchtex = "C:\planillas\BlueCross-Mod.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE MODIFICACIONES MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "J" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 13
         Xarchexel.Cells(Xlin, XCol) = "MODIFICACION"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("I" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("J" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_inf.Recordset.EOF
            If IsNull(data_inf.Recordset("cl_nombre")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nombre")
            Else
               Xarchexel.Cells(Xlin, XCol) = "MODIF"
            End If
            XCol = XCol + 1
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_codigo")
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_apellid")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
            Else
               Xarchexel.Cells(Xlin, XCol) = "NN"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_celular")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_celular")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_fnac")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_fnac")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_zona")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_zona")
            Else
               Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("info_debit")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("info_debit")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
            End If
            data_inf.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      End If
      
      data_infno.RecordSource = "Select * from infno where cl_nombre in ('BAJA') order by cl_apellid"
      data_infno.Refresh
      If data_infno.Recordset.RecordCount > 0 Then
         data_infno.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "BAJAS"
         Xlibexel.SaveAs ("C:\planillas\BlueCross-Bajas.xls")
         Xarchtex = "C:\planillas\BlueCross-Bajas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE BAJAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRES"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_infno.Recordset.EOF
            If IsNull(data_infno.Recordset("cl_codigo")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_codigo")
            Else
               Xarchexel.Cells(Xlin, XCol) = "0"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_apellid")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_apellid")
            Else
               Xarchexel.Cells(Xlin, XCol) = "NN"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_cedula")) = False Then
               If IsNull(data_infno.Recordset("cl_codced")) = False Then
                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_infno.Recordset("cl_cedula"))) & "-" & Trim(str(data_infno.Recordset("cl_codced")))
               Else
                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_infno.Recordset("cl_cedula"))) & "-0"
               End If
            Else
               Xarchexel.Cells(Xlin, XCol) = "0-0"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_codconv")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_codconv")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_zona")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_zona")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_direcci")
            End If
            data_infno.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      End If
      MsgBox "Proceso terminado"
   End If
End If


End Sub

Public Sub Caute()
Dim Xcedeva As Long
Dim Xcedblue, Xfnac As String

Dim Xobjexel As Excel.Application
Dim Xlibexel As Excel.Workbook
Dim Xarchexel As New Excel.Worksheet

Dim XCol, Xlin, Xnrocan, Xcolfija As Long
Dim Xarchtex As String
Dim Xlabrir As New Excel.Application
Dim XsiAltas As String

If Combo1.Text = "CAUTE" Then 'OK
   data_mut.RecordSource = "caute"
   data_mut.Refresh
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         data_mut.Recordset.Delete
         data_mut.Recordset.MoveNext
      Loop
   End If
   Data1.DatabaseName = "C:\mutuales\caute.xls"
   Data1.RecordSource = "socios$"
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         data_mut.Recordset.AddNew
         If IsNull(Data1.Recordset("ced")) = False Then
            data_mut.Recordset("ced") = Data1.Recordset("ced")
         Else
            data_mut.Recordset("ced") = 0
         End If
         If IsNull(Data1.Recordset("mat")) = False Then
            data_mut.Recordset("mat") = Data1.Recordset("mat")
         End If
         If IsNull(Data1.Recordset("nom1")) = False Then
            data_mut.Recordset("nom1") = Data1.Recordset("nom1")
         End If
         If IsNull(Data1.Recordset("ape1")) = False Then
            data_mut.Recordset("ape1") = Data1.Recordset("ape1")
         End If
         If IsNull(Data1.Recordset("fnac")) = False Then
            data_mut.Recordset("fnac") = Format(Data1.Recordset("fnac"), "dd/mm/yyyy")
         End If
         If IsNull(Data1.Recordset("domicilio")) = False Then
            data_mut.Recordset("domicilio") = Mid(Data1.Recordset("domicilio"), 1, 255)
         End If
         If IsNull(Data1.Recordset("telefono")) = False Then
            data_mut.Recordset("telefono") = Mid(Data1.Recordset("telefono"), 1, 255)
         End If
         If IsNull(Data1.Recordset("celular")) = False Then
            data_mut.Recordset("celular") = Mid(Data1.Recordset("celular"), 1, 255)
         End If
         If IsNull(Data1.Recordset("correo")) = False Then
            data_mut.Recordset("correo") = Mid(Data1.Recordset("correo"), 1, 255)
         End If
         data_mut.Recordset.Update
         Data1.Recordset.MoveNext
      Loop
   End If
   data_mut.Refresh
   
   Label3.Visible = True
   Label3.Caption = "Procesando Altas/Modif"
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveLast
      DoEvents
      pb.Visible = True
      pb.Max = data_mut.Recordset.RecordCount
      pb.Value = 0
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         If IsNull(data_mut.Recordset("ced")) = False Then
            If Len(data_mut.Recordset("ced")) = 7 Then
               Xcedblue = Mid(Trim(str(data_mut.Recordset("ced"))), 1, 6)
            Else
               If Len(data_mut.Recordset("ced")) = 8 Then
                  Xcedblue = Mid(Trim(str(data_mut.Recordset("ced"))), 1, 7)
               Else
                  Xcedblue = Mid(Trim(str(data_mut.Recordset("ced"))), 1, 5)
               End If
            End If
            data_mut.Recordset.Edit
            data_mut.Recordset("cednum") = Val(Xcedblue)
            data_mut.Recordset.Update
            Xcedeva = Val(Xcedblue)
         Else
            Xcedeva = 0
         End If
         If Xcedeva > 0 Then
         
'            data_cli.Recordset.FindFirst "cl_cedula =" & Xcedeva
            data_cli.RecordSource = "Select * from clientes where cl_cedula =" & Xcedeva
            data_cli.Refresh
            If data_cli.Recordset.RecordCount > 0 Then
               data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
               data_conv.Refresh
               If data_conv.Recordset.RecordCount > 0 Then
                  If data_conv.Recordset("cnv_codigo") = "CAUTE" Then
                     If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                        data_inf.Recordset.AddNew
                        If IsNull(data_mut.Recordset("fnac")) = False Then
                           data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                        End If
                        data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced")))
                        If IsNull(data_mut.Recordset("domicilio")) = False Then
                           data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                        End If
                        If IsNull(data_mut.Recordset("celular")) = False Then
                           data_inf.Recordset("cl_dpto") = Mid(Trim(data_mut.Recordset("celular")), 1, 12)
                        End If
                        If IsNull(data_mut.Recordset("telefono")) = False Then
                           data_inf.Recordset("cl_telefon") = Mid(Trim(data_mut.Recordset("telefono")), 1, 20)
                        End If
                        If IsNull(data_mut.Recordset("ape2")) = False Then
                           If IsNull(data_mut.Recordset("nom2")) = False Then
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                           Else
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                           End If
                        Else
                           If IsNull(data_mut.Recordset("nom2")) = False Then
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                           Else
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                           End If
                        End If
                        If IsNull(data_mut.Recordset("correo")) = False Then
                           data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                        End If
                        data_inf.Recordset("cl_nombre") = "REACTIVAR"
                        data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                        data_inf.Recordset.Update
                     End If
                  Else
                     data_inf.Recordset.AddNew
                     If IsNull(data_mut.Recordset("fnac")) = False Then
                        data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                     End If
                     data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced")))
                     If IsNull(data_mut.Recordset("domicilio")) = False Then
                        data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                     End If
                     If IsNull(data_mut.Recordset("celular")) = False Then
                        data_inf.Recordset("cl_dpto") = Mid(Trim(data_mut.Recordset("celular")), 1, 12)
                     End If
                     If IsNull(data_mut.Recordset("telefono")) = False Then
                        data_inf.Recordset("cl_telefon") = Mid(Trim(data_mut.Recordset("telefono")), 1, 20)
                     End If
                     If IsNull(data_mut.Recordset("ape2")) = False Then
                        If IsNull(data_mut.Recordset("nom2")) = False Then
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                        Else
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                        End If
                     Else
                        If IsNull(data_mut.Recordset("nom2")) = False Then
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                        Else
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                        End If
                     End If
                     If IsNull(data_mut.Recordset("correo")) = False Then
                        data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                     End If
                     If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                        data_inf.Recordset("cl_nombre") = "CONV.INCORRECTO BAJ"
                     Else
                        data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                     End If
                     data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_inf.Recordset.Update
                  End If
               End If
            Else
               data_inf.Recordset.AddNew
               If IsNull(data_mut.Recordset("fnac")) = False Then
                  data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
               End If
               data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced")))
               If IsNull(data_mut.Recordset("domicilio")) = False Then
                  data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
               End If
               If IsNull(data_mut.Recordset("celular")) = False Then
                  data_inf.Recordset("cl_dpto") = Mid(Trim(data_mut.Recordset("celular")), 1, 12)
               End If
               If IsNull(data_mut.Recordset("telefono")) = False Then
                  data_inf.Recordset("cl_telefon") = Mid(Trim(data_mut.Recordset("telefono")), 1, 20)
               End If
               If IsNull(data_mut.Recordset("ape2")) = False Then
                  If IsNull(data_mut.Recordset("nom2")) = False Then
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                  Else
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                  End If
               Else
                  If IsNull(data_mut.Recordset("nom2")) = False Then
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                  Else
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                  End If
               End If
               If IsNull(data_mut.Recordset("correo")) = False Then
                  data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
               End If
               data_inf.Recordset("cl_nombre") = "NO ESTA EN P.SAPP"
               data_inf.Recordset.Update
            End If
         End If
         data_mut.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      DoEvents
      data_cli.RecordSource = "Select * from clientes where estado not in (2,3) and cl_codconv in ('CAUTE')"
      data_cli.Refresh
      data_cli.Recordset.MoveLast
      data_cli.Recordset.MoveFirst
      pb.Max = pb.Max + data_cli.Recordset.RecordCount
      data_mut.Refresh
      Label3.Caption = "Procesando BAJAS..."
      DoEvents
      Do While Not data_cli.Recordset.EOF
         If IsNull(data_cli.Recordset("cl_codconv")) = False Then
            If IsNull(data_cli.Recordset("cl_cedula")) = False Then
               If data_cli.Recordset("cl_cedula") > 0 Then
                  data_mut.RecordSource = "Select * from caute where cednum =" & Int(data_cli.Recordset("cl_cedula"))
                  data_mut.Refresh
                  If data_mut.Recordset.RecordCount > 0 Then
                  Else
                     data_infno.Recordset.AddNew
                     data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                     data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                     data_infno.Recordset("cl_nombre") = "BAJA"
                     data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                     data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                     data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                     data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                     data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                     data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                     data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                     data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                     data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                     data_infno.Recordset.Update
                  End If
               Else
                   data_infno.Recordset.AddNew
                   data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                   data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                   data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                   data_infno.Recordset("cl_nombre") = "BAJA"
                   data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                   data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                   data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                   data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                   data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                   data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                   data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                   data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                   data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                   data_infno.Recordset.Update
               End If
            End If
         End If
         data_cli.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      Label3.Visible = False
      Label3.Caption = ""
      DoEvents

      data_inf.RecordSource = "Select * from infcli where cl_nombre ='" & "NO ESTA EN P.SAPP" & "' order by cl_apellid"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "ALTAS"
         Xlibexel.SaveAs ("C:\planillas\" & "CAUTE-Altas" & ".xls")
         Xarchtex = "C:\planillas\CAUTE-Altas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE ALTAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_inf.Recordset.EOF
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_celular")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_celular")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_fnac")) = False Then
               Xarchexel.Cells(Xlin, XCol) = CDate(data_inf.Recordset("cl_fnac"))
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_zona")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_zona")
            Else
               Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("info_debit")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("info_debit")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
            End If
            data_inf.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      Else
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "ALTAS"
         Xlibexel.SaveAs ("C:\planillas\" & "CAUTE-Altas" & ".xls")
         Xarchtex = "C:\planillas\CAUTE-Altas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE ALTAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      
      End If
      
      data_inf.RecordSource = "Select * from infcli where cl_nombre not in ('NO ESTA EN P.SAPP','ACTIVO') order by cl_apellid"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1
         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         
         Xarchexel.Name = "MODIF"
         Xlibexel.SaveAs ("C:\planillas\CAUTE-Mod.xls")
         Xarchtex = "C:\planillas\CAUTE-Mod.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE MODIFICACIONES MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "J" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 13
         Xarchexel.Cells(Xlin, XCol) = "MODIFICACION"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("I" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("J" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_inf.Recordset.EOF
            If IsNull(data_inf.Recordset("cl_nombre")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nombre")
            Else
               Xarchexel.Cells(Xlin, XCol) = "MODIF"
            End If
            XCol = XCol + 1
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_codigo")
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_apellid")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
            Else
               Xarchexel.Cells(Xlin, XCol) = "NN"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_celular")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_celular")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_fnac")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_fnac")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_zona")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_zona")
            Else
               Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("info_debit")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("info_debit")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
            End If
            data_inf.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      End If
      
      data_infno.RecordSource = "Select * from infno where cl_nombre in ('BAJA') order by cl_apellid"
      data_infno.Refresh
      If data_infno.Recordset.RecordCount > 0 Then
         data_infno.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "BAJAS"
         Xlibexel.SaveAs ("C:\planillas\CAUTE-Bajas.xls")
         Xarchtex = "C:\planillas\CAUTE-Bajas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE BAJAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRES"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_infno.Recordset.EOF
            If IsNull(data_infno.Recordset("cl_codigo")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_codigo")
            Else
               Xarchexel.Cells(Xlin, XCol) = "0"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_apellid")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_apellid")
            Else
               Xarchexel.Cells(Xlin, XCol) = "NN"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_cedula")) = False Then
               If IsNull(data_infno.Recordset("cl_codced")) = False Then
                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_infno.Recordset("cl_cedula"))) & "-" & Trim(str(data_infno.Recordset("cl_codced")))
               Else
                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_infno.Recordset("cl_cedula"))) & "-0"
               End If
            Else
               Xarchexel.Cells(Xlin, XCol) = "0-0"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_codconv")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_codconv")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_zona")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_zona")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_direcci")
            End If
            data_infno.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      End If
      
      XsiAltas = MsgBox("Confirma que desea agregar las ALTAS en forma automática al sistema?", vbInformation + vbYesNo)
      If XsiAltas = vbYes Then
         Dim XlacedCas As Long
         Dim XeldvCas As Integer
         Dim Xlanuevamat As Long
         Dim Xconteo As Integer
         Dim Xladire As String
         frm_ctrolmut.MousePointer = 11
         data_inf.RecordSource = "Select * from infcli where cl_nombre ='" & "NO ESTA EN P.SAPP" & "' order by cl_apellid"
         data_inf.Refresh
         Xconteo = 0
         If data_inf.Recordset.RecordCount > 0 Then
            data_inf.Recordset.MoveFirst
            Do While Not data_inf.Recordset.EOF
               If Len(Trim(data_inf.Recordset("cl_celular"))) = 8 Then
                  XlacedCas = Val(Mid(Trim(data_inf.Recordset("cl_celular")), 1, 7))
                  XeldvCas = Val(Mid(Trim(data_inf.Recordset("cl_celular")), 8, 1))
               Else
                  If Len(Trim(data_inf.Recordset("cl_celular"))) = 7 Then
                     XlacedCas = Val(Mid(Trim(data_inf.Recordset("cl_celular")), 1, 6))
                     XeldvCas = Val(Mid(Trim(data_inf.Recordset("cl_celular")), 7, 1))
                  Else
                     XlacedCas = Val(Mid(Trim(data_inf.Recordset("cl_celular")), 1, 7))
                     XeldvCas = Val(Mid(Trim(data_inf.Recordset("cl_celular")), 8, 1))
                  End If
               End If
               data_cli.RecordSource = "select * from clientes where cl_cedula =" & XlacedCas
               data_cli.Refresh
               If data_cli.Recordset.RecordCount > 0 Then
               
               Else
                  Xlanuevamat = data_paramnew.Recordset("p_matric") + 1
                  data_paramnew.Recordset.Edit
                  data_paramnew.Recordset("p_matric") = data_paramnew.Recordset("p_matric") + 1
                  data_paramnew.Recordset.Update
                  data_paramnew.Refresh
                  
                  data_cli.Recordset.AddNew
                  data_cli.Recordset("cl_codigo") = Xlanuevamat
                  data_cli.Recordset("estado") = 1
                  data_cli.Recordset("cl_codconv") = "CAUTE"
                  data_cli.Recordset("cl_nomconv") = "CAUTE ANTEL"
                  data_cli.Recordset("cl_apellid") = data_inf.Recordset("cl_apellid")
                  If IsNull(data_inf.Recordset("cl_fnac")) = False Then
                     data_cli.Recordset("cl_fnac") = data_inf.Recordset("cl_fnac")
                  End If
                  data_cli.Recordset("cl_cedula") = XlacedCas
                  data_cli.Recordset("cl_codced") = XeldvCas
                  data_cli.Recordset("cl_nrovend") = 799
                  data_cli.Recordset("cl_nomvend") = "*TODOS"
                  data_cli.Recordset("cl_fecing") = Date
                  data_cli.Recordset("cl_forpago") = 1
                  data_cli.Recordset("cl_descpag") = "Abono Mensual"
                  data_cli.Recordset("cl_sexo") = 1
                  data_cli.Recordset("fecha_sys") = Date
                  data_cli.Recordset.Update
                  Xconteo = Xconteo + 1
               End If
               data_inf.Recordset.MoveNext
            Loop
            frm_ctrolmut.MousePointer = 0
            MsgBox "Proceso de altas automáticas terminado. Se ingresaron " & Xconteo & " registros.", vbExclamation
         End If
      End If
      MsgBox "Proceso terminado"
   End If
End If

End Sub

Public Sub Summum()
Dim Xcedeva As Long
Dim Xcedblue, Xfnac As String

Dim Xobjexel As Excel.Application
Dim Xlibexel As Excel.Workbook
Dim Xarchexel As New Excel.Worksheet

Dim XCol, Xlin, Xnrocan, Xcolfija As Long
Dim Xarchtex As String
Dim Xlabrir As New Excel.Application

If Combo1.Text = "SUMMUM" Then 'OK
   data_mut.RecordSource = "summum"
   data_mut.Refresh
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         data_mut.Recordset.Delete
         data_mut.Recordset.MoveNext
      Loop
   End If
   Data1.DatabaseName = "C:\mutuales\summum.xls"
   Data1.RecordSource = "socios$"
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         data_mut.Recordset.AddNew
         If IsNull(Data1.Recordset("ced")) = False Then
            data_mut.Recordset("ced") = Data1.Recordset("ced")
         Else
            data_mut.Recordset("ced") = 0
         End If
         If IsNull(Data1.Recordset("nom1")) = False Then
            data_mut.Recordset("nom1") = Data1.Recordset("nom1")
         End If
         If IsNull(Data1.Recordset("nom2")) = False Then
            data_mut.Recordset("nom2") = Data1.Recordset("nom2")
         End If
         If IsNull(Data1.Recordset("ape1")) = False Then
            data_mut.Recordset("ape1") = Data1.Recordset("ape1")
         End If
         If IsNull(Data1.Recordset("ape2")) = False Then
            data_mut.Recordset("ape2") = Data1.Recordset("ape2")
         End If
         If IsNull(Data1.Recordset("fnac")) = False Then
            data_mut.Recordset("fnac") = Format(Data1.Recordset("fnac"), "dd/mm/yyyy")
         End If
         If IsNull(Data1.Recordset("zona")) = False Then
            data_mut.Recordset("zona") = Data1.Recordset("zona")
         End If
         If IsNull(Data1.Recordset("domicilio")) = False Then
            If IsNull(Data1.Recordset("domicilio2")) = False Then
               data_mut.Recordset("domicilio") = Mid(Data1.Recordset("domicilio"), 1, 100) & " " & Mid(Data1.Recordset("domicilio2"), 1, 100)
            Else
               data_mut.Recordset("domicilio") = Mid(Data1.Recordset("domicilio"), 1, 200)
            End If
         End If
         If IsNull(Data1.Recordset("telefono")) = False Then
            data_mut.Recordset("telefono") = Mid(Data1.Recordset("telefono"), 1, 255)
         End If
         If IsNull(Data1.Recordset("celular")) = False Then
            data_mut.Recordset("celular") = Mid(Data1.Recordset("celular"), 1, 255)
         End If
         If IsNull(Data1.Recordset("correo")) = False Then
            data_mut.Recordset("correo") = Mid(Data1.Recordset("correo"), 1, 255)
         End If
         data_mut.Recordset.Update
         Data1.Recordset.MoveNext
      Loop
   End If
   data_mut.Refresh
   
   Label3.Visible = True
   Label3.Caption = "Procesando Altas/Modif"
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveLast
      DoEvents
      pb.Visible = True
      pb.Max = data_mut.Recordset.RecordCount
      pb.Value = 0
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         If IsNull(data_mut.Recordset("ced")) = False Then
            If Len(data_mut.Recordset("ced")) = 7 Then
               Xcedblue = Mid(Trim(str(data_mut.Recordset("ced"))), 1, 6)
            Else
               If Len(data_mut.Recordset("ced")) = 8 Then
                  Xcedblue = Mid(Trim(str(data_mut.Recordset("ced"))), 1, 7)
               Else
                  Xcedblue = Mid(Trim(str(data_mut.Recordset("ced"))), 1, 5)
               End If
            End If
            data_mut.Recordset.Edit
            data_mut.Recordset("cednum") = Val(Xcedblue)
            data_mut.Recordset.Update
            Xcedeva = Val(Xcedblue)
         Else
            Xcedeva = 0
         End If
         If Xcedeva > 0 Then
         
'            data_cli.Recordset.FindFirst "cl_cedula =" & Xcedeva
            data_cli.RecordSource = "Select * from clientes where cl_cedula =" & Xcedeva
            data_cli.Refresh
            If data_cli.Recordset.RecordCount > 0 Then
               data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
               data_conv.Refresh
               If data_conv.Recordset.RecordCount > 0 Then
                  If data_conv.Recordset("cnv_codigo") = "SUMMUM" Then
                     If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                        data_inf.Recordset.AddNew
                        If IsNull(data_mut.Recordset("fnac")) = False Then
                           data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                        End If
                        data_inf.Recordset("cl_celular") = Trim(data_mut.Recordset("ced"))
                        If IsNull(data_mut.Recordset("domicilio")) = False Then
                           data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                        End If
                        If IsNull(data_mut.Recordset("celular")) = False Then
                           data_inf.Recordset("cl_dpto") = Mid(Trim(data_mut.Recordset("celular")), 1, 12)
                        End If
                        If IsNull(data_mut.Recordset("telefono")) = False Then
                           data_inf.Recordset("cl_telefon") = Mid(Trim(data_mut.Recordset("telefono")), 1, 20)
                        End If
                        If IsNull(data_mut.Recordset("ape2")) = False Then
                           If IsNull(data_mut.Recordset("nom2")) = False Then
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                           Else
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                           End If
                        Else
                           If IsNull(data_mut.Recordset("nom2")) = False Then
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                           Else
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                           End If
                        End If
                        If IsNull(data_mut.Recordset("zona")) = False Then
                           data_inf.Recordset("cl_zona") = Trim(data_mut.Recordset("zona"))
                        End If
                        If IsNull(data_mut.Recordset("correo")) = False Then
                           data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                        End If
                        data_inf.Recordset("cl_nombre") = "REACTIVAR"
                        data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                        data_inf.Recordset.Update
                     End If
                  Else
                     data_inf.Recordset.AddNew
                     If IsNull(data_mut.Recordset("fnac")) = False Then
                        data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                     End If
                     data_inf.Recordset("cl_celular") = Trim(data_mut.Recordset("ced"))
                     If IsNull(data_mut.Recordset("domicilio")) = False Then
                        data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                     End If
                     If IsNull(data_mut.Recordset("celular")) = False Then
                        data_inf.Recordset("cl_dpto") = Mid(Trim(data_mut.Recordset("celular")), 1, 12)
                     End If
                     If IsNull(data_mut.Recordset("telefono")) = False Then
                        data_inf.Recordset("cl_telefon") = Mid(Trim(data_mut.Recordset("telefono")), 1, 20)
                     End If
                     If IsNull(data_mut.Recordset("ape2")) = False Then
                        If IsNull(data_mut.Recordset("nom2")) = False Then
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                        Else
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                        End If
                     Else
                        If IsNull(data_mut.Recordset("nom2")) = False Then
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                        Else
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                        End If
                     End If
                     If IsNull(data_mut.Recordset("correo")) = False Then
                        data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                     End If
                     If IsNull(data_mut.Recordset("zona")) = False Then
                        data_inf.Recordset("cl_zona") = Trim(data_mut.Recordset("zona"))
                     End If
                     If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                        data_inf.Recordset("cl_nombre") = "CONV.INCORRECTO BAJ"
                     Else
                        data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                     End If
                     data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_inf.Recordset.Update
                  End If
               End If
            Else
               data_inf.Recordset.AddNew
               If IsNull(data_mut.Recordset("fnac")) = False Then
                  data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
               End If
               data_inf.Recordset("cl_celular") = Trim(data_mut.Recordset("ced"))
               If IsNull(data_mut.Recordset("domicilio")) = False Then
                  data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
               End If
               If IsNull(data_mut.Recordset("celular")) = False Then
                  data_inf.Recordset("cl_dpto") = Mid(Trim(data_mut.Recordset("celular")), 1, 12)
               End If
               If IsNull(data_mut.Recordset("telefono")) = False Then
                  data_inf.Recordset("cl_telefon") = Mid(Trim(data_mut.Recordset("telefono")), 1, 20)
               End If
               If IsNull(data_mut.Recordset("ape2")) = False Then
                  If IsNull(data_mut.Recordset("nom2")) = False Then
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                  Else
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                  End If
               Else
                  If IsNull(data_mut.Recordset("nom2")) = False Then
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                  Else
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                  End If
               End If
               If IsNull(data_mut.Recordset("correo")) = False Then
                  data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
               End If
               If IsNull(data_mut.Recordset("zona")) = False Then
                  data_inf.Recordset("cl_zona") = Trim(data_mut.Recordset("zona"))
               End If
               data_inf.Recordset("cl_nombre") = "NO ESTA EN P.SAPP"
               data_inf.Recordset.Update
            End If
         End If
         data_mut.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      DoEvents
      data_cli.RecordSource = "Select * from clientes where estado not in (2,3) and cl_codconv in ('SUMMUM')"
      data_cli.Refresh
      data_cli.Recordset.MoveLast
      data_cli.Recordset.MoveFirst
      pb.Max = pb.Max + data_cli.Recordset.RecordCount
      data_mut.Refresh
      Label3.Caption = "Procesando BAJAS..."
      DoEvents
      Do While Not data_cli.Recordset.EOF
         If IsNull(data_cli.Recordset("cl_codconv")) = False Then
            If IsNull(data_cli.Recordset("cl_cedula")) = False Then
               If data_cli.Recordset("cl_cedula") > 0 Then
                  data_mut.RecordSource = "Select * from summum where cednum =" & Int(data_cli.Recordset("cl_cedula"))
                  data_mut.Refresh
                  If data_mut.Recordset.RecordCount > 0 Then
                  Else
                     data_infno.Recordset.AddNew
                     data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                     data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                     data_infno.Recordset("cl_nombre") = "BAJA"
                     data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                     data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                     data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                     data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                     data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                     data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                     data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                     data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                     data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                     data_infno.Recordset.Update
                  End If
               Else
                   data_infno.Recordset.AddNew
                   data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                   data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                   data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                   data_infno.Recordset("cl_nombre") = "BAJA"
                   data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                   data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                   data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                   data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                   data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                   data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                   data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                   data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                   data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                   data_infno.Recordset.Update
               End If
            End If
         End If
         data_cli.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      Label3.Visible = False
      Label3.Caption = ""
      DoEvents

      data_inf.RecordSource = "Select * from infcli where cl_nombre ='" & "NO ESTA EN P.SAPP" & "' order by cl_apellid"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "ALTAS"
         Xlibexel.SaveAs ("C:\planillas\" & "Summum-Altas" & ".xls")
         Xarchtex = "C:\planillas\Summum-Altas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE ALTAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_inf.Recordset.EOF
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_celular")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_celular")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_fnac")) = False Then
               Xarchexel.Cells(Xlin, XCol) = CDate(data_inf.Recordset("cl_fnac"))
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_zona")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_zona")
            Else
               Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("info_debit")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("info_debit")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
            End If
            data_inf.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      Else
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "ALTAS"
         Xlibexel.SaveAs ("C:\planillas\" & "Summum-Altas" & ".xls")
         Xarchtex = "C:\planillas\Summum-Altas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE ALTAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      
      End If
      
      data_inf.RecordSource = "Select * from infcli where cl_nombre not in ('NO ESTA EN P.SAPP','ACTIVO') order by cl_apellid"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1
         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         
         Xarchexel.Name = "MODIF"
         Xlibexel.SaveAs ("C:\planillas\Summum-Mod.xls")
         Xarchtex = "C:\planillas\Summum-Mod.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE MODIFICACIONES MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "J" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 13
         Xarchexel.Cells(Xlin, XCol) = "MODIFICACION"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("I" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("J" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_inf.Recordset.EOF
            If IsNull(data_inf.Recordset("cl_nombre")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nombre")
            Else
               Xarchexel.Cells(Xlin, XCol) = "MODIF"
            End If
            XCol = XCol + 1
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_codigo")
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_apellid")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
            Else
               Xarchexel.Cells(Xlin, XCol) = "NN"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_celular")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_celular")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_fnac")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_fnac")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_zona")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_zona")
            Else
               Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("info_debit")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("info_debit")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
            End If
            data_inf.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      End If
      
      data_infno.RecordSource = "Select * from infno where cl_nombre in ('BAJA') order by cl_apellid"
      data_infno.Refresh
      If data_infno.Recordset.RecordCount > 0 Then
         data_infno.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "BAJAS"
         Xlibexel.SaveAs ("C:\planillas\Summum-Bajas.xls")
         Xarchtex = "C:\planillas\Summum-Bajas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE BAJAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRES"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_infno.Recordset.EOF
            If IsNull(data_infno.Recordset("cl_codigo")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_codigo")
            Else
               Xarchexel.Cells(Xlin, XCol) = "0"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_apellid")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_apellid")
            Else
               Xarchexel.Cells(Xlin, XCol) = "NN"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_cedula")) = False Then
               If IsNull(data_infno.Recordset("cl_codced")) = False Then
                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_infno.Recordset("cl_cedula"))) & "-" & Trim(str(data_infno.Recordset("cl_codced")))
               Else
                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_infno.Recordset("cl_cedula"))) & "-0"
               End If
            Else
               Xarchexel.Cells(Xlin, XCol) = "0-0"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_codconv")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_codconv")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_zona")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_zona")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_direcci")
            End If
            data_infno.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      End If
      MsgBox "Proceso terminado"
   End If
End If

End Sub

Public Sub Seguro()
Dim Xcedeva As Long
Dim Xcedblue, Xfnac As String

Dim Xobjexel As Excel.Application
Dim Xlibexel As Excel.Workbook
Dim Xarchexel As New Excel.Worksheet

Dim XCol, Xlin, Xnrocan, Xcolfija As Long
Dim Xarchtex As String
Dim Xlabrir As New Excel.Application

If Combo1.Text = "SEGURO AMERICANO" Then 'OK
   data_mut.RecordSource = "seguro"
   data_mut.Refresh
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         data_mut.Recordset.Delete
         data_mut.Recordset.MoveNext
      Loop
   End If
   Data1.DatabaseName = "C:\mutuales\seguro.xls"
   Data1.RecordSource = "socios$"
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         data_mut.Recordset.AddNew
         If IsNull(Data1.Recordset("ced")) = False Then
            data_mut.Recordset("ced") = Val(Data1.Recordset("ced"))
         Else
            data_mut.Recordset("ced") = 0
         End If
         If IsNull(Data1.Recordset("nom1")) = False Then
            data_mut.Recordset("nom1") = Data1.Recordset("nom1")
         End If
         If IsNull(Data1.Recordset("nom2")) = False Then
            data_mut.Recordset("nom2") = Data1.Recordset("nom2")
         End If
         If IsNull(Data1.Recordset("ape1")) = False Then
            data_mut.Recordset("ape1") = Data1.Recordset("ape1")
         End If
         If IsNull(Data1.Recordset("ape2")) = False Then
            data_mut.Recordset("ape2") = Data1.Recordset("ape2")
         End If
         If IsNull(Data1.Recordset("fnac")) = False Then
            data_mut.Recordset("fnac") = Format(Data1.Recordset("fnac"), "dd/mm/yyyy")
         End If
         If IsNull(Data1.Recordset("fecing")) = False Then
            data_mut.Recordset("fecing") = Format(Data1.Recordset("fecing"), "dd/mm/yyyy")
         End If
         If IsNull(Data1.Recordset("zona")) = False Then
            data_mut.Recordset("zona") = Data1.Recordset("zona")
         End If
         If IsNull(Data1.Recordset("domicilio")) = False Then
            data_mut.Recordset("domicilio") = Mid(Data1.Recordset("domicilio"), 1, 255)
         End If
         If IsNull(Data1.Recordset("telefono")) = False Then
            data_mut.Recordset("telefono") = Mid(Data1.Recordset("telefono"), 1, 255)
         End If
         If IsNull(Data1.Recordset("celular")) = False Then
            data_mut.Recordset("celular") = Mid(Data1.Recordset("celular"), 1, 255)
         End If
         If IsNull(Data1.Recordset("correo")) = False Then
            data_mut.Recordset("correo") = Mid(Data1.Recordset("correo"), 1, 255)
         End If
         data_mut.Recordset.Update
         Data1.Recordset.MoveNext
      Loop
   End If
   data_mut.Refresh
   
   Label3.Visible = True
   Label3.Caption = "Procesando Altas/Modif"
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveLast
      DoEvents
      pb.Visible = True
      pb.Max = data_mut.Recordset.RecordCount
      pb.Value = 0
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         If IsNull(data_mut.Recordset("ced")) = False Then
            If Len(data_mut.Recordset("ced")) = 7 Then
               Xcedblue = Mid(Trim(str(data_mut.Recordset("ced"))), 1, 6)
            Else
               If Len(data_mut.Recordset("ced")) = 8 Then
                  Xcedblue = Mid(Trim(str(data_mut.Recordset("ced"))), 1, 7)
               Else
                  Xcedblue = Mid(Trim(str(data_mut.Recordset("ced"))), 1, 5)
               End If
            End If
            data_mut.Recordset.Edit
            data_mut.Recordset("cednum") = Val(Xcedblue)
            data_mut.Recordset.Update
            Xcedeva = Val(Xcedblue)
         Else
            Xcedeva = 0
         End If
         If Xcedeva > 0 Then
         
'            data_cli.Recordset.FindFirst "cl_cedula =" & Xcedeva
            data_cli.RecordSource = "Select * from clientes where cl_cedula =" & Xcedeva
            data_cli.Refresh
            If data_cli.Recordset.RecordCount > 0 Then
               data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
               data_conv.Refresh
               If data_conv.Recordset.RecordCount > 0 Then
                  If data_conv.Recordset("cnv_codigo") = "SEGAM" Then
                     If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                        data_inf.Recordset.AddNew
                        If IsNull(data_mut.Recordset("fnac")) = False Then
                           data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                        End If
                        data_inf.Recordset("cl_celular") = Trim(data_mut.Recordset("ced"))
                        If IsNull(data_mut.Recordset("domicilio")) = False Then
                           data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                        End If
                        If IsNull(data_mut.Recordset("celular")) = False Then
                           data_inf.Recordset("cl_dpto") = Mid(Trim(data_mut.Recordset("celular")), 1, 12)
                        End If
                        If IsNull(data_mut.Recordset("telefono")) = False Then
                           data_inf.Recordset("cl_telefon") = Mid(Trim(data_mut.Recordset("telefono")), 1, 20)
                        End If
                        If IsNull(data_mut.Recordset("ape2")) = False Then
                           If IsNull(data_mut.Recordset("nom2")) = False Then
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                           Else
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                           End If
                        Else
                           If IsNull(data_mut.Recordset("nom2")) = False Then
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                           Else
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                           End If
                        End If
                        If IsNull(data_mut.Recordset("zona")) = False Then
                           data_inf.Recordset("cl_zona") = Trim(data_mut.Recordset("zona"))
                        End If
                        If IsNull(data_mut.Recordset("correo")) = False Then
                           data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                        End If
                        data_inf.Recordset("cl_nombre") = "REACTIVAR"
                        data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                        data_inf.Recordset.Update
                     End If
                  Else
                     data_inf.Recordset.AddNew
                     If IsNull(data_mut.Recordset("fnac")) = False Then
                        data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                     End If
                     data_inf.Recordset("cl_celular") = Trim(data_mut.Recordset("ced"))
                     If IsNull(data_mut.Recordset("domicilio")) = False Then
                        data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                     End If
                     If IsNull(data_mut.Recordset("celular")) = False Then
                        data_inf.Recordset("cl_dpto") = Mid(Trim(data_mut.Recordset("celular")), 1, 12)
                     End If
                     If IsNull(data_mut.Recordset("telefono")) = False Then
                        data_inf.Recordset("cl_telefon") = Mid(Trim(data_mut.Recordset("telefono")), 1, 20)
                     End If
                     If IsNull(data_mut.Recordset("ape2")) = False Then
                        If IsNull(data_mut.Recordset("nom2")) = False Then
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                        Else
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                        End If
                     Else
                        If IsNull(data_mut.Recordset("nom2")) = False Then
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                        Else
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                        End If
                     End If
                     If IsNull(data_mut.Recordset("correo")) = False Then
                        data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                     End If
                     If IsNull(data_mut.Recordset("zona")) = False Then
                        data_inf.Recordset("cl_zona") = Trim(data_mut.Recordset("zona"))
                     End If
                     If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                        data_inf.Recordset("cl_nombre") = "CONV.INCORRECTO BAJ"
                     Else
                        data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                     End If
                     data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_inf.Recordset.Update
                  End If
               End If
            Else
               data_inf.Recordset.AddNew
               If IsNull(data_mut.Recordset("fnac")) = False Then
                  data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
               End If
               data_inf.Recordset("cl_celular") = Trim(data_mut.Recordset("ced"))
               If IsNull(data_mut.Recordset("domicilio")) = False Then
                  data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
               End If
               If IsNull(data_mut.Recordset("celular")) = False Then
                  data_inf.Recordset("cl_dpto") = Mid(Trim(data_mut.Recordset("celular")), 1, 12)
               End If
               If IsNull(data_mut.Recordset("telefono")) = False Then
                  data_inf.Recordset("cl_telefon") = Mid(Trim(data_mut.Recordset("telefono")), 1, 20)
               End If
               If IsNull(data_mut.Recordset("ape2")) = False Then
                  If IsNull(data_mut.Recordset("nom2")) = False Then
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                  Else
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                  End If
               Else
                  If IsNull(data_mut.Recordset("nom2")) = False Then
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                  Else
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                  End If
               End If
               If IsNull(data_mut.Recordset("correo")) = False Then
                  data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
               End If
               If IsNull(data_mut.Recordset("zona")) = False Then
                  data_inf.Recordset("cl_zona") = Trim(data_mut.Recordset("zona"))
               End If
               data_inf.Recordset("cl_nombre") = "NO ESTA EN P.SAPP"
               data_inf.Recordset.Update
            End If
         End If
         data_mut.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      DoEvents
      data_cli.RecordSource = "Select * from clientes where estado not in (2,3) and cl_codconv in ('SEGAM')"
      data_cli.Refresh
      data_cli.Recordset.MoveLast
      data_cli.Recordset.MoveFirst
      pb.Max = pb.Max + data_cli.Recordset.RecordCount
      data_mut.Refresh
      Label3.Caption = "Procesando BAJAS..."
      DoEvents
      Do While Not data_cli.Recordset.EOF
         If IsNull(data_cli.Recordset("cl_codconv")) = False Then
            If IsNull(data_cli.Recordset("cl_cedula")) = False Then
               If data_cli.Recordset("cl_cedula") > 0 Then
                  data_mut.RecordSource = "Select * from seguro where cednum =" & Int(data_cli.Recordset("cl_cedula"))
                  data_mut.Refresh
                  If data_mut.Recordset.RecordCount > 0 Then
                  Else
                     data_infno.Recordset.AddNew
                     data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                     data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                     data_infno.Recordset("cl_nombre") = "BAJA"
                     data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                     data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                     data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                     data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                     data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                     data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                     data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                     data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                     data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                     data_infno.Recordset.Update
                  End If
               Else
                   data_infno.Recordset.AddNew
                   data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                   data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                   data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                   data_infno.Recordset("cl_nombre") = "BAJA"
                   data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                   data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                   data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                   data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                   data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                   data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                   data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                   data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                   data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                   data_infno.Recordset.Update
               End If
            End If
         End If
         data_cli.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      Label3.Visible = False
      Label3.Caption = ""
      DoEvents

      data_inf.RecordSource = "Select * from infcli where cl_nombre ='" & "NO ESTA EN P.SAPP" & "' order by cl_apellid"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "ALTAS"
         Xlibexel.SaveAs ("C:\planillas\" & "SEGAM-Altas" & ".xls")
         Xarchtex = "C:\planillas\SEGAM-Altas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE ALTAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_inf.Recordset.EOF
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_celular")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_celular")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_fnac")) = False Then
               Xarchexel.Cells(Xlin, XCol) = CDate(data_inf.Recordset("cl_fnac"))
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_zona")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_zona")
            Else
               Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("info_debit")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("info_debit")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
            End If
            data_inf.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      Else
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "ALTAS"
         Xlibexel.SaveAs ("C:\planillas\" & "SEGAM-Altas" & ".xls")
         Xarchtex = "C:\planillas\SEGAM-Altas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE ALTAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      
      End If
      
      data_inf.RecordSource = "Select * from infcli where cl_nombre not in ('NO ESTA EN P.SAPP','ACTIVO') order by cl_apellid"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1
         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         
         Xarchexel.Name = "MODIF"
         Xlibexel.SaveAs ("C:\planillas\SEGAM-Mod.xls")
         Xarchtex = "C:\planillas\SEGAM-Mod.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE MODIFICACIONES MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "J" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 13
         Xarchexel.Cells(Xlin, XCol) = "MODIFICACION"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("I" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("J" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_inf.Recordset.EOF
            If IsNull(data_inf.Recordset("cl_nombre")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nombre")
            Else
               Xarchexel.Cells(Xlin, XCol) = "MODIF"
            End If
            XCol = XCol + 1
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_codigo")
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_apellid")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
            Else
               Xarchexel.Cells(Xlin, XCol) = "NN"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_celular")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_celular")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_fnac")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_fnac")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_zona")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_zona")
            Else
               Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("info_debit")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("info_debit")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
            End If
            data_inf.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      End If
      
      data_infno.RecordSource = "Select * from infno where cl_nombre in ('BAJA') order by cl_apellid"
      data_infno.Refresh
      If data_infno.Recordset.RecordCount > 0 Then
         data_infno.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "BAJAS"
         Xlibexel.SaveAs ("C:\planillas\SEGAM-Bajas.xls")
         Xarchtex = "C:\planillas\SEGAM-Bajas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE BAJAS MUTUALISTA: " & Combo1.Text & " FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRES"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_infno.Recordset.EOF
            If IsNull(data_infno.Recordset("cl_codigo")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_codigo")
            Else
               Xarchexel.Cells(Xlin, XCol) = "0"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_apellid")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_apellid")
            Else
               Xarchexel.Cells(Xlin, XCol) = "NN"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_cedula")) = False Then
               If IsNull(data_infno.Recordset("cl_codced")) = False Then
                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_infno.Recordset("cl_cedula"))) & "-" & Trim(str(data_infno.Recordset("cl_codced")))
               Else
                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_infno.Recordset("cl_cedula"))) & "-0"
               End If
            Else
               Xarchexel.Cells(Xlin, XCol) = "0-0"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_codconv")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_codconv")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_zona")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_zona")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_direcci")
            End If
            data_infno.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      End If
      MsgBox "Proceso terminado"
   End If
End If

End Sub

Public Sub CcouSJ()
Dim Xcedeva As Long
Dim Xcedblue, Xfnac As String

Dim Xobjexel As Excel.Application
Dim Xlibexel As Excel.Workbook
Dim Xarchexel As New Excel.Worksheet

Dim XCol, Xlin, Xnrocan, Xcolfija As Long
Dim Xarchtex As String
Dim Xlabrir As New Excel.Application

Dim Xpond, Xn1, Xn2, Xn3, Xn4, Xn5, Xn6, Xn7, Xtot, Xlacedu As Long
Dim Xcedtex, Xtottex As String
Dim Xced1, Xced2, Xced3, Xced4, Xced5, Xced6, Xced7, Xlargo As Long


If Combo1.Text = "CCOU SJ" Then 'OK
   data_mut.RecordSource = "ccou"
   data_mut.Refresh
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         data_mut.Recordset.Delete
         data_mut.Recordset.MoveNext
      Loop
   End If
   Data1.DatabaseName = "C:\mutuales\ccousj.xls"
   Data1.RecordSource = "socios$"
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         data_mut.Recordset.AddNew
         data_mut.Recordset("ced") = Data1.Recordset("ced")
         data_mut.Recordset("nom1") = Data1.Recordset("nom1")
         If IsNull(Data1.Recordset("nom2")) = False Then
            data_mut.Recordset("nom2") = Data1.Recordset("nom2")
         End If
         data_mut.Recordset("ape1") = Data1.Recordset("ape1")
         If IsNull(Data1.Recordset("ape2")) = False Then
            data_mut.Recordset("ape2") = Data1.Recordset("ape2")
         End If
         data_mut.Recordset("fnac") = Data1.Recordset("fnac")
         data_mut.Recordset("categ") = Mid(Data1.Recordset("categ"), 1, 255)
         data_mut.Recordset("domicilio") = Mid(Data1.Recordset("domicilio"), 1, 255)
         data_mut.Recordset("telefono") = Mid(Data1.Recordset("telefono"), 1, 255)
         data_mut.Recordset("celular") = Mid(Data1.Recordset("celular"), 1, 255)
         data_mut.Recordset("correo") = Mid(Data1.Recordset("correo"), 1, 255)
         data_mut.Recordset("fecing") = Data1.Recordset("fecing")
         data_mut.Recordset.Update
         Data1.Recordset.MoveNext
      Loop
   End If
   data_mut.Refresh
   Xn1 = 2
   Xn2 = 9
   Xn3 = 8
   Xn4 = 7
   Xn5 = 6
   Xn6 = 3
   Xn7 = 4
   Xpond = 10
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         Xcedtex = Trim(data_mut.Recordset("ced"))
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
         data_mut.Recordset.Edit
         data_mut.Recordset("dv") = Xtot
         data_mut.Recordset.Update
         data_mut.Recordset.MoveNext
      Loop
   End If
   Label3.Visible = True
   Label3.Caption = "Procesando Altas/Modif"
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveFirst
      DoEvents
      pb.Visible = True
      pb.Max = data_mut.Recordset.RecordCount
      pb.Value = 0
      Do While Not data_mut.Recordset.EOF
         If IsNull(data_mut.Recordset("ced")) = False Then
            Xcedeva = Val(data_mut.Recordset("ced"))
            data_mut.Recordset.Edit
            data_mut.Recordset("cednum") = Xcedeva
            data_mut.Recordset.Update
         Else
            Xcedeva = 0
         End If
         If Xcedeva > 0 Then
            data_cli.RecordSource = "Select * from clientes where cl_cedula =" & Xcedeva
            data_cli.Refresh
            If data_cli.Recordset.RecordCount > 0 Then
               data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
               data_conv.Refresh
               If data_conv.Recordset.RecordCount > 0 Then
                  If data_conv.Recordset("cnv_grupo") = "CCOU" Then
                     If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                        data_inf.Recordset.AddNew
                        If IsNull(data_mut.Recordset("fnac")) = False Then
                           data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                        End If
                        data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced"))) & "-" & Trim(str(data_mut.Recordset("dv")))
                        If IsNull(data_mut.Recordset("domicilio")) = False Then
                           data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                        End If
                        If IsNull(data_mut.Recordset("categ")) = False Then
                           data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                        End If
                        If IsNull(data_mut.Recordset("celular")) = False Then
                           data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                        End If
                        If IsNull(data_mut.Recordset("telefono")) = False Then
                           data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                        End If
                        If IsNull(data_mut.Recordset("ape2")) = False Then
                           If IsNull(data_mut.Recordset("nom2")) = False Then
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                           Else
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                           End If
                        Else
                           If IsNull(data_mut.Recordset("nom2")) = False Then
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                           Else
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                           End If
                        End If
                        If IsNull(data_mut.Recordset("correo")) = False Then
                           data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                        End If
                        data_inf.Recordset("cl_nombre") = "REACTIVAR"
                        data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                        data_inf.Recordset.Update
                     Else
                        If data_conv.Recordset("cnv_codigo") = "CCNOS" Or data_conv.Recordset("cnv_codigo") = "CCNRE" Or _
                           data_conv.Recordset("cnv_codigo") = "CCNSAM" Then
                            data_inf.Recordset.AddNew
                            If IsNull(data_mut.Recordset("fnac")) = False Then
                               data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                            End If
                            data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced"))) & "-" & Trim(str(data_mut.Recordset("dv")))
                            If IsNull(data_mut.Recordset("domicilio")) = False Then
                               data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                            End If
                            If IsNull(data_mut.Recordset("categ")) = False Then
                               data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                            End If
                            If IsNull(data_mut.Recordset("celular")) = False Then
                               data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                            End If
                            If IsNull(data_mut.Recordset("telefono")) = False Then
                               data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                            End If
                            If IsNull(data_mut.Recordset("ape2")) = False Then
                               If IsNull(data_mut.Recordset("nom2")) = False Then
                                  data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                               Else
                                  data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                               End If
                            Else
                               If IsNull(data_mut.Recordset("nom2")) = False Then
                                  data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                               Else
                                  data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                               End If
                            End If
                            If IsNull(data_mut.Recordset("correo")) = False Then
                               data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                            End If
                            data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                            data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                            data_inf.Recordset.Update
                        End If
                     End If
                  Else
                     data_inf.Recordset.AddNew
                     If IsNull(data_mut.Recordset("fnac")) = False Then
                        data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                     End If
                     data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced"))) & "-" & Trim(str(data_mut.Recordset("dv")))
                     If IsNull(data_mut.Recordset("domicilio")) = False Then
                        data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                     End If
                     If IsNull(data_mut.Recordset("categ")) = False Then
                        data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                     End If
                     If IsNull(data_mut.Recordset("celular")) = False Then
                        data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                     End If
                     If IsNull(data_mut.Recordset("telefono")) = False Then
                        data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                     End If
                     If IsNull(data_mut.Recordset("ape2")) = False Then
                        If IsNull(data_mut.Recordset("nom2")) = False Then
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                        Else
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                        End If
                     Else
                        If IsNull(data_mut.Recordset("nom2")) = False Then
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                        Else
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                        End If
                     End If
                     If IsNull(data_mut.Recordset("correo")) = False Then
                        data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                     End If
                     If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                        data_inf.Recordset("cl_nombre") = "CONV.INCORRECTO BAJ"
                     Else
                        data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                     End If
                     data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_inf.Recordset.Update
                  End If
               Else
                  data_inf.Recordset.AddNew
                  If IsNull(data_mut.Recordset("fnac")) = False Then
                     data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                  End If
                  data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced"))) & "-" & Trim(str(data_mut.Recordset("dv")))
                  If IsNull(data_mut.Recordset("domicilio")) = False Then
                     data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                  End If
                  If IsNull(data_mut.Recordset("categ")) = False Then
                     data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                  End If
                  If IsNull(data_mut.Recordset("celular")) = False Then
                     data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                  End If
                  If IsNull(data_mut.Recordset("telefono")) = False Then
                     data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                  End If
                  If IsNull(data_mut.Recordset("ape2")) = False Then
                     If IsNull(data_mut.Recordset("nom2")) = False Then
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                     Else
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                     End If
                  Else
                     If IsNull(data_mut.Recordset("nom2")) = False Then
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                     Else
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                     End If
                  End If
                  If IsNull(data_mut.Recordset("correo")) = False Then
                     data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                  End If
                  If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                     data_inf.Recordset("cl_nombre") = "CONV.INCORRECTO BAJA"
                  Else
                     data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                  End If
                  data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                  data_inf.Recordset.Update
               End If
            Else
               data_inf.Recordset.AddNew
               If IsNull(data_mut.Recordset("fnac")) = False Then
                  data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
               End If
               data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced"))) & "-" & Trim(str(data_mut.Recordset("dv")))
               If IsNull(data_mut.Recordset("domicilio")) = False Then
                  data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
               End If
               If IsNull(data_mut.Recordset("categ")) = False Then
                  data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
               End If
               If IsNull(data_mut.Recordset("celular")) = False Then
                  data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
               End If
               If IsNull(data_mut.Recordset("telefono")) = False Then
                  data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
               End If
               If IsNull(data_mut.Recordset("ape2")) = False Then
                  If IsNull(data_mut.Recordset("nom2")) = False Then
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                  Else
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                  End If
               Else
                  If IsNull(data_mut.Recordset("nom2")) = False Then
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                  Else
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                  End If
               End If
               If IsNull(data_mut.Recordset("correo")) = False Then
                  data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
               End If
               data_inf.Recordset("cl_nombre") = "NO ESTA EN P.SAPP"
               data_inf.Recordset.Update
            End If
         End If
         data_mut.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      DoEvents
      data_cli.RecordSource = "Select * from clientes where estado <>" & 2 & " And estado <>" & 3 & " and cl_codconv in ('CCFSJ','CCFSJA')"
      data_cli.Refresh
      data_cli.Recordset.MoveLast
      data_cli.Recordset.MoveFirst
      pb.Max = pb.Max + data_cli.Recordset.RecordCount
      data_mut.Refresh
      Label3.Caption = "Procesando BAJAS..."
      DoEvents
      Do While Not data_cli.Recordset.EOF
         If IsNull(data_cli.Recordset("cl_codconv")) = False Then
            If IsNull(data_cli.Recordset("cl_cedula")) = False Then
               If data_cli.Recordset("cl_cedula") > 0 Then
                  If data_cli.Recordset("cl_codconv") = "CCNOS" Or data_cli.Recordset("cl_codconv") = "CCNRE" Or _
                     data_cli.Recordset("cl_codconv") = "CCNSAM" Then
                  Else
'                     data_conv.Recordset.FindFirst "cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                     data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                     data_conv.Refresh
                     If data_conv.Recordset.RecordCount > 0 Then
                        If data_conv.Recordset("cnv_grupo") = "CCOU" Then
                           data_mut.RecordSource = "Select * from ccou where cednum =" & Int(data_cli.Recordset("cl_cedula"))
                           data_mut.Refresh
'                           data_mut.Recordset.FindFirst "cednum =" & Int(data_cli.Recordset("cl_cedula"))
                           If data_mut.Recordset.RecordCount > 0 Then
                           Else
                              data_infno.Recordset.AddNew
                              data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                              data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                              data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                              data_infno.Recordset("cl_nombre") = "BAJA"
                              data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                              data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                              data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                              data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                              data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                              data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                              data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                              data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                              data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                              data_infno.Recordset.Update
                           End If
                        End If
                     End If
                  End If
               Else
                  data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                  data_conv.Refresh
                  If data_conv.Recordset.RecordCount > 0 Then
                     If data_conv.Recordset("cnv_grupo") = "CCOU" Then
                        data_infno.Recordset.AddNew
                        data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                        data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                        data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                        data_infno.Recordset("cl_nombre") = "BAJA"
                        data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                        data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                        data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                        data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                        data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                        data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                        data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                        data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                        data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                        data_infno.Recordset.Update
                     End If
                  End If
               End If
            Else
               data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
               data_conv.Refresh
               If data_conv.Recordset.RecordCount > 0 Then
                  If data_conv.Recordset("cnv_grupo") = "CCOU" Then
                     data_infno.Recordset.AddNew
                     data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                     data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                     data_infno.Recordset("cl_nombre") = "BAJA"
                     data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                     data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                     data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                     data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                     data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                     data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                     data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                     data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                     data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                     data_infno.Recordset.Update
                  End If
               End If
            End If
         End If
         data_cli.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      Label3.Visible = False
      Label3.Caption = ""
      DoEvents
      
      data_inf.RecordSource = "Select * from infcli where cl_nombre ='" & "NO ESTA EN P.SAPP" & "' order by cl_apellid"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "ALTAS"
         Xlibexel.SaveAs ("C:\planillas\" & "CCOU-Altas" & ".xls")
         Xarchtex = "C:\planillas\CCOU-Altas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE ALTAS MUTUALISTA: " & Combo1.Text & " SAN JACINTO FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_inf.Recordset.EOF
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_celular")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_celular")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_fnac")) = False Then
               Xarchexel.Cells(Xlin, XCol) = CDate(data_inf.Recordset("cl_fnac"))
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nom_sup")
            Else
               Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("info_debit")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("info_debit")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
            End If
            data_inf.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      Else
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "ALTAS"
         Xlibexel.SaveAs ("C:\planillas\" & "CCOU-Altas" & ".xls")
         Xarchtex = "C:\planillas\CCOU-Altas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE ALTAS MUTUALISTA: " & Combo1.Text & " SAN JACINTO FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      
      End If
      
      data_inf.RecordSource = "Select * from infcli where cl_nombre not in ('NO ESTA EN P.SAPP','ACTIVO') order by cl_apellid"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1
         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "MODIF"
         Xlibexel.SaveAs ("C:\planillas\CCOU-Mod.xls")
         Xarchtex = "C:\planillas\CCOU-Mod.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE MODIFICACIONES MUTUALISTA: " & Combo1.Text & " SAN JACINTO FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "J" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 13
         Xarchexel.Cells(Xlin, XCol) = "MODIFICACION"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("I" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("J" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_inf.Recordset.EOF
            If IsNull(data_inf.Recordset("cl_nombre")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nombre")
            Else
               Xarchexel.Cells(Xlin, XCol) = "MODIF"
            End If
            XCol = XCol + 1
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_codigo")
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_apellid")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
            Else
               Xarchexel.Cells(Xlin, XCol) = "NN"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_celular")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_celular")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_fnac")) = False Then
               Xarchexel.Cells(Xlin, XCol) = CDate(data_inf.Recordset("cl_fnac"))
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nom_sup")
            Else
               Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("info_debit")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("info_debit")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
            End If
            data_inf.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      End If
      
      data_infno.RecordSource = "Select * from infno where cl_nombre in ('BAJA') order by cl_apellid"
      data_infno.Refresh
      If data_infno.Recordset.RecordCount > 0 Then
         data_infno.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "BAJAS"
         Xlibexel.SaveAs ("C:\planillas\CCOU-Bajas.xls")
         Xarchtex = "C:\planillas\CCOU-Bajas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE BAJAS MUTUALISTA: " & Combo1.Text & " SAN JACINTO FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRES"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_infno.Recordset.EOF
            If IsNull(data_infno.Recordset("cl_codigo")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_codigo")
            Else
               Xarchexel.Cells(Xlin, XCol) = "0"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_apellid")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_apellid")
            Else
               Xarchexel.Cells(Xlin, XCol) = "NN"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_cedula")) = False Then
               If IsNull(data_infno.Recordset("cl_codced")) = False Then
                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_infno.Recordset("cl_cedula"))) & "-" & Trim(str(data_infno.Recordset("cl_codced")))
               Else
                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_infno.Recordset("cl_cedula"))) & "-0"
               End If
            Else
               Xarchexel.Cells(Xlin, XCol) = "0-0"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_codconv")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_codconv")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_zona")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_zona")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_direcci")
            End If
            data_infno.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      End If
   End If
   MsgBox "Proceso terminado"


End If

End Sub
Public Sub CcouAtl()
Dim Xcedeva As Long
Dim Xcedblue, Xfnac As String

Dim Xobjexel As Excel.Application
Dim Xlibexel As Excel.Workbook
Dim Xarchexel As New Excel.Worksheet

Dim XCol, Xlin, Xnrocan, Xcolfija As Long
Dim Xarchtex As String
Dim Xlabrir As New Excel.Application

Dim Xpond, Xn1, Xn2, Xn3, Xn4, Xn5, Xn6, Xn7, Xtot, Xlacedu As Long
Dim Xcedtex, Xtottex As String
Dim Xced1, Xced2, Xced3, Xced4, Xced5, Xced6, Xced7, Xlargo As Long

'CCFAT
'CCFATA

If Combo1.Text = "CCOU ATL" Then 'OK
   data_mut.RecordSource = "ccou"
   data_mut.Refresh
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveFirst
      Do While Not data_mut.Recordset.EOF
         data_mut.Recordset.Delete
         data_mut.Recordset.MoveNext
      Loop
   End If
   Data1.DatabaseName = "C:\mutuales\ccouatl.xls"
   Data1.RecordSource = "socios$"
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      Do While Not Data1.Recordset.EOF
         data_mut.Recordset.AddNew
         data_mut.Recordset("ced") = Data1.Recordset("ced")
         data_mut.Recordset("dv") = Data1.Recordset("dv")
         data_mut.Recordset("nom1") = Data1.Recordset("nom1")
         If IsNull(Data1.Recordset("nom2")) = False Then
            data_mut.Recordset("nom2") = Data1.Recordset("nom2")
         End If
         data_mut.Recordset("ape1") = Data1.Recordset("ape1")
         If IsNull(Data1.Recordset("ape2")) = False Then
            data_mut.Recordset("ape2") = Data1.Recordset("ape2")
         End If
         data_mut.Recordset("fnac") = Data1.Recordset("fnac")
         data_mut.Recordset("categ") = Mid(Data1.Recordset("categ"), 1, 255)
         data_mut.Recordset("domicilio") = Mid(Data1.Recordset("domicilio"), 1, 255)
         data_mut.Recordset("telefono") = Mid(Data1.Recordset("telefono"), 1, 255)
         data_mut.Recordset("celular") = Mid(Data1.Recordset("celular"), 1, 255)
         data_mut.Recordset("correo") = Mid(Data1.Recordset("correo"), 1, 255)
         data_mut.Recordset("fecing") = Data1.Recordset("fecing")
         data_mut.Recordset.Update
         Data1.Recordset.MoveNext
      Loop
   End If
   data_mut.Refresh
   Label3.Visible = True
   Label3.Caption = "Procesando Altas/Modif"
   If data_mut.Recordset.RecordCount > 0 Then
      data_mut.Recordset.MoveLast
      data_mut.Recordset.MoveFirst
      DoEvents
      pb.Visible = True
      pb.Max = data_mut.Recordset.RecordCount
      pb.Value = 0
      Do While Not data_mut.Recordset.EOF
         If IsNull(data_mut.Recordset("ced")) = False Then
            Xcedeva = Val(data_mut.Recordset("ced"))
            data_mut.Recordset.Edit
            data_mut.Recordset("cednum") = Xcedeva
            data_mut.Recordset.Update
         Else
            Xcedeva = 0
         End If
         If Xcedeva > 0 Then
            data_cli.RecordSource = "Select * from clientes where cl_cedula =" & Xcedeva
            data_cli.Refresh
            If data_cli.Recordset.RecordCount > 0 Then
               data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
               data_conv.Refresh
               If data_conv.Recordset.RecordCount > 0 Then
                  If data_conv.Recordset("cnv_grupo") = "CCOU" Then
                     If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                        data_inf.Recordset.AddNew
                        If IsNull(data_mut.Recordset("fnac")) = False Then
                           data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                        End If
                        data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced"))) & "-" & Trim(str(data_mut.Recordset("dv")))
                        If IsNull(data_mut.Recordset("domicilio")) = False Then
                           data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                        End If
                        If IsNull(data_mut.Recordset("categ")) = False Then
                           data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                        End If
                        If IsNull(data_mut.Recordset("celular")) = False Then
                           data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                        End If
                        If IsNull(data_mut.Recordset("telefono")) = False Then
                           data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                        End If
                        If IsNull(data_mut.Recordset("ape2")) = False Then
                           If IsNull(data_mut.Recordset("nom2")) = False Then
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                           Else
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                           End If
                        Else
                           If IsNull(data_mut.Recordset("nom2")) = False Then
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                           Else
                              data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                           End If
                        End If
                        If IsNull(data_mut.Recordset("correo")) = False Then
                           data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                        End If
                        data_inf.Recordset("cl_nombre") = "REACTIVAR"
                        data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                        data_inf.Recordset.Update
                     Else
                        If data_conv.Recordset("cnv_codigo") = "CCNOS" Or data_conv.Recordset("cnv_codigo") = "CCNRE" Or _
                           data_conv.Recordset("cnv_codigo") = "CCNSAM" Then
                            data_inf.Recordset.AddNew
                            If IsNull(data_mut.Recordset("fnac")) = False Then
                               data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                            End If
                            data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced"))) & "-" & Trim(str(data_mut.Recordset("dv")))
                            If IsNull(data_mut.Recordset("domicilio")) = False Then
                               data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                            End If
                            If IsNull(data_mut.Recordset("categ")) = False Then
                               data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                            End If
                            If IsNull(data_mut.Recordset("celular")) = False Then
                               data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                            End If
                            If IsNull(data_mut.Recordset("telefono")) = False Then
                               data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                            End If
                            If IsNull(data_mut.Recordset("ape2")) = False Then
                               If IsNull(data_mut.Recordset("nom2")) = False Then
                                  data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                               Else
                                  data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                               End If
                            Else
                               If IsNull(data_mut.Recordset("nom2")) = False Then
                                  data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                               Else
                                  data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                               End If
                            End If
                            If IsNull(data_mut.Recordset("correo")) = False Then
                               data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                            End If
                            data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                            data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                            data_inf.Recordset.Update
                        End If
                     End If
                  Else
                     data_inf.Recordset.AddNew
                     If IsNull(data_mut.Recordset("fnac")) = False Then
                        data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                     End If
                     data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced"))) & "-" & Trim(str(data_mut.Recordset("dv")))
                     If IsNull(data_mut.Recordset("domicilio")) = False Then
                        data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                     End If
                     If IsNull(data_mut.Recordset("categ")) = False Then
                        data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                     End If
                     If IsNull(data_mut.Recordset("celular")) = False Then
                        data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                     End If
                     If IsNull(data_mut.Recordset("telefono")) = False Then
                        data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                     End If
                     If IsNull(data_mut.Recordset("ape2")) = False Then
                        If IsNull(data_mut.Recordset("nom2")) = False Then
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                        Else
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                        End If
                     Else
                        If IsNull(data_mut.Recordset("nom2")) = False Then
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                        Else
                           data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                        End If
                     End If
                     If IsNull(data_mut.Recordset("correo")) = False Then
                        data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                     End If
                     If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                        data_inf.Recordset("cl_nombre") = "CONV.INCORRECTO BAJ"
                     Else
                        data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                     End If
                     data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_inf.Recordset.Update
                  End If
               Else
                  data_inf.Recordset.AddNew
                  If IsNull(data_mut.Recordset("fnac")) = False Then
                     data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
                  End If
                  data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced"))) & "-" & Trim(str(data_mut.Recordset("dv")))
                  If IsNull(data_mut.Recordset("domicilio")) = False Then
                     data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
                  End If
                  If IsNull(data_mut.Recordset("categ")) = False Then
                     data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
                  End If
                  If IsNull(data_mut.Recordset("celular")) = False Then
                     data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
                  End If
                  If IsNull(data_mut.Recordset("telefono")) = False Then
                     data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
                  End If
                  If IsNull(data_mut.Recordset("ape2")) = False Then
                     If IsNull(data_mut.Recordset("nom2")) = False Then
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                     Else
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                     End If
                  Else
                     If IsNull(data_mut.Recordset("nom2")) = False Then
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                     Else
                        data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                     End If
                  End If
                  If IsNull(data_mut.Recordset("correo")) = False Then
                     data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
                  End If
                  If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                     data_inf.Recordset("cl_nombre") = "CONV.INCORRECTO BAJA"
                  Else
                     data_inf.Recordset("cl_nombre") = "CONVENIO INCORRECTO"
                  End If
                  data_inf.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                  data_inf.Recordset.Update
               End If
            Else
               data_inf.Recordset.AddNew
               If IsNull(data_mut.Recordset("fnac")) = False Then
                  data_inf.Recordset("cl_fnac") = data_mut.Recordset("fnac")
               End If
               data_inf.Recordset("cl_celular") = Trim(str(data_mut.Recordset("ced"))) & "-" & Trim(str(data_mut.Recordset("dv")))
               If IsNull(data_mut.Recordset("domicilio")) = False Then
                  data_inf.Recordset("cl_direcci") = Trim(Mid(data_mut.Recordset("domicilio"), 1, 80))
               End If
               If IsNull(data_mut.Recordset("categ")) = False Then
                  data_inf.Recordset("cl_nom_sup") = Trim(Mid(data_mut.Recordset("categ"), 1, 25))
               End If
               If IsNull(data_mut.Recordset("celular")) = False Then
                  data_inf.Recordset("cl_dpto") = Trim(data_mut.Recordset("celular"))
               End If
               If IsNull(data_mut.Recordset("telefono")) = False Then
                  data_inf.Recordset("cl_telefon") = Trim(data_mut.Recordset("telefono"))
               End If
               If IsNull(data_mut.Recordset("ape2")) = False Then
                  If IsNull(data_mut.Recordset("nom2")) = False Then
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1") & " " & data_mut.Recordset("nom2")
                  Else
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("ape2") + " " + data_mut.Recordset("nom1")
                  End If
               Else
                  If IsNull(data_mut.Recordset("nom2")) = False Then
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1") & " " + data_mut.Recordset("nom2")
                  Else
                     data_inf.Recordset("cl_apellid") = data_mut.Recordset("ape1") + " " + data_mut.Recordset("nom1")
                  End If
               End If
               If IsNull(data_mut.Recordset("correo")) = False Then
                  data_inf.Recordset("info_debit") = Trim(data_mut.Recordset("correo"))
               End If
               data_inf.Recordset("cl_nombre") = "NO ESTA EN P.SAPP"
               data_inf.Recordset.Update
            End If
         End If
         data_mut.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      DoEvents
      data_cli.RecordSource = "Select * from clientes where estado <>" & 2 & " And estado <>" & 3 & " and cl_codconv in ('CCFAT','CCFATA')"
      data_cli.Refresh
      data_cli.Recordset.MoveLast
      data_cli.Recordset.MoveFirst
      pb.Max = pb.Max + data_cli.Recordset.RecordCount
      data_mut.Refresh
      Label3.Caption = "Procesando BAJAS..."
      DoEvents
      Do While Not data_cli.Recordset.EOF
         If IsNull(data_cli.Recordset("cl_codconv")) = False Then
            If IsNull(data_cli.Recordset("cl_cedula")) = False Then
               If data_cli.Recordset("cl_cedula") > 0 Then
                  If data_cli.Recordset("cl_codconv") = "CCNOS" Or data_cli.Recordset("cl_codconv") = "CCNRE" Or _
                     data_cli.Recordset("cl_codconv") = "CCNSAM" Then
                  Else
'                     data_conv.Recordset.FindFirst "cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                     data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                     data_conv.Refresh
                     If data_conv.Recordset.RecordCount > 0 Then
                        If data_conv.Recordset("cnv_grupo") = "CCOU" Then
                           data_mut.RecordSource = "Select * from ccou where cednum =" & Int(data_cli.Recordset("cl_cedula"))
                           data_mut.Refresh
'                           data_mut.Recordset.FindFirst "cednum =" & Int(data_cli.Recordset("cl_cedula"))
                           If data_mut.Recordset.RecordCount > 0 Then
                           Else
                              data_infno.Recordset.AddNew
                              data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                              data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                              data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                              data_infno.Recordset("cl_nombre") = "BAJA"
                              data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                              data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                              data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                              data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                              data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                              data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                              data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                              data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                              data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                              data_infno.Recordset.Update
                           End If
                        End If
                     End If
                  End If
               Else
                  data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
                  data_conv.Refresh
                  If data_conv.Recordset.RecordCount > 0 Then
                     If data_conv.Recordset("cnv_grupo") = "CCOU" Then
                        data_infno.Recordset.AddNew
                        data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                        data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                        data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                        data_infno.Recordset("cl_nombre") = "BAJA"
                        data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                        data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                        data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                        data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                        data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                        data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                        data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                        data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                        data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                        data_infno.Recordset.Update
                     End If
                  End If
               End If
            Else
               data_conv.RecordSource = "Select * from convenio where cnv_codigo ='" & data_cli.Recordset("cl_codconv") & "'"
               data_conv.Refresh
               If data_conv.Recordset.RecordCount > 0 Then
                  If data_conv.Recordset("cnv_grupo") = "CCOU" Then
                     data_infno.Recordset.AddNew
                     data_infno.Recordset("cl_codigo") = data_cli.Recordset("cl_codigo")
                     data_infno.Recordset("cl_apellid") = data_cli.Recordset("cl_apellid")
                     data_infno.Recordset("cl_codconv") = data_cli.Recordset("cl_codconv")
                     data_infno.Recordset("cl_nombre") = "BAJA"
                     data_infno.Recordset("cl_cedula") = data_cli.Recordset("cl_cedula")
                     data_infno.Recordset("cl_codced") = data_cli.Recordset("cl_codced")
                     data_infno.Recordset("cl_telefon") = data_cli.Recordset("cl_telefon")
                     data_infno.Recordset("cl_dpto") = data_cli.Recordset("cl_dpto")
                     data_infno.Recordset("cl_zona") = data_cli.Recordset("cl_zona")
                     data_infno.Recordset("cl_fnac") = data_cli.Recordset("cl_fnac")
                     data_infno.Recordset("cl_direcci") = data_cli.Recordset("cl_direcci")
                     data_infno.Recordset("cl_fecing") = data_cli.Recordset("cl_fecing")
                     data_infno.Recordset("cl_nrocobr") = data_cli.Recordset("cl_nrocobr")
                     data_infno.Recordset.Update
                  End If
               End If
            End If
         End If
         data_cli.Recordset.MoveNext
         pb.Value = pb.Value + 1
      Loop
      Label3.Visible = False
      Label3.Caption = ""
      DoEvents
      
      data_inf.RecordSource = "Select * from infcli where cl_nombre ='" & "NO ESTA EN P.SAPP" & "' order by cl_apellid"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "ALTAS"
         Xlibexel.SaveAs ("C:\planillas\" & "CCOU-Altas" & ".xls")
         Xarchtex = "C:\planillas\CCOU-Altas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE ALTAS MUTUALISTA: " & Combo1.Text & " ATLANTIDA FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_inf.Recordset.EOF
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_celular")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_celular")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_fnac")) = False Then
               Xarchexel.Cells(Xlin, XCol) = CDate(data_inf.Recordset("cl_fnac"))
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nom_sup")
            Else
               Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("info_debit")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("info_debit")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
            End If
            data_inf.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      Else
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "ALTAS"
         Xlibexel.SaveAs ("C:\planillas\" & "CCOU-Altas" & ".xls")
         Xarchtex = "C:\planillas\CCOU-Altas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE ALTAS MUTUALISTA: " & Combo1.Text & " ATLANTIDA FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      
      End If
      
      data_inf.RecordSource = "Select * from infcli where cl_nombre not in ('NO ESTA EN P.SAPP','ACTIVO') order by cl_apellid"
      data_inf.Refresh
      If data_inf.Recordset.RecordCount > 0 Then
         data_inf.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1
         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "MODIF"
         Xlibexel.SaveAs ("C:\planillas\CCOU-Mod.xls")
         Xarchtex = "C:\planillas\CCOU-Mod.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE MODIFICACIONES MUTUALISTA: " & Combo1.Text & " ATLANTIDA FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "J" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "J" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 13
         Xarchexel.Cells(Xlin, XCol) = "MODIFICACION"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRE"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 12
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "FEC.NAC."
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("I" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "CORREO ELECTRONICO"
         XCol = XCol + 1
         Xarchexel.Range("J" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_inf.Recordset.EOF
            If IsNull(data_inf.Recordset("cl_nombre")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nombre")
            Else
               Xarchexel.Cells(Xlin, XCol) = "MODIF"
            End If
            XCol = XCol + 1
            Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_codigo")
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_apellid")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_apellid")
            Else
               Xarchexel.Cells(Xlin, XCol) = "NN"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_celular")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_celular")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_fnac")) = False Then
               Xarchexel.Cells(Xlin, XCol) = CDate(data_inf.Recordset("cl_fnac"))
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_nom_sup")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_nom_sup")
            Else
               Xarchexel.Cells(Xlin, XCol) = "Sin Datos"
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("info_debit")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("info_debit")
            End If
            XCol = XCol + 1
            If IsNull(data_inf.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_inf.Recordset("cl_direcci")
            End If
            data_inf.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      End If
      
      data_infno.RecordSource = "Select * from infno where cl_nombre in ('BAJA') order by cl_apellid"
      data_infno.Refresh
      If data_infno.Recordset.RecordCount > 0 Then
         data_infno.Recordset.MoveFirst
         XCol = 1
         Xlin = 1
         Xnrocan = 1

         Set Xobjexel = New Excel.Application
         Set Xlibexel = Xobjexel.Workbooks.Add
         Set Xarchexel = Xlibexel.Worksheets.Add
         Xarchexel.Name = "BAJAS"
         Xlibexel.SaveAs ("C:\planillas\CCOU-Bajas.xls")
         Xarchtex = "C:\planillas\CCOU-Bajas.xls"
         Xarchexel.Cells(Xlin, XCol) = "SAPP - DPTO.TI"
         Xlin = Xlin + 1
         XCol = XCol + 1
         Xarchexel.Range("A1", "C3").Font.Size = 16
         Xarchexel.Cells(Xlin, XCol) = "INFORME DE BAJAS MUTUALISTA: " & Combo1.Text & " ATLANTIDA FECHA: " & Date
         Xarchexel.Range("B" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(0, 200, 200)
         XCol = 1
         Xlin = Xlin + 2
         Xnrocan = Xnrocan + Xlin
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideHorizontal).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlInsideVertical).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeBottom).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeTop).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeLeft).LineStyle = xlContinuous
         Xarchexel.Range("A4", "I" & Trim(str(Xnrocan))).Borders(xlEdgeRight).LineStyle = xlContinuous
         Xarchexel.Range("A" & Trim(str(Xlin)), "I" & Trim(str(Xlin))).Interior.color = RGB(215, 120, 120)
         Xarchexel.Range("A" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "MATRICULA"
         XCol = XCol + 1
         Xarchexel.Range("B" & Trim(str(Xlin))).ColumnWidth = 35
         Xarchexel.Cells(Xlin, XCol) = "APELLIDO/NOMBRES"
         XCol = XCol + 1
         Xarchexel.Range("C" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CEDULA"
         XCol = XCol + 1
         Xarchexel.Range("D" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CONVENIO"
         XCol = XCol + 1
         Xarchexel.Range("E" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "CELULAR"
         XCol = XCol + 1
         Xarchexel.Range("F" & Trim(str(Xlin))).ColumnWidth = 10
         Xarchexel.Cells(Xlin, XCol) = "TELEFONO"
         XCol = XCol + 1
         Xarchexel.Range("G" & Trim(str(Xlin))).ColumnWidth = 15
         Xarchexel.Cells(Xlin, XCol) = "ZONA"
         XCol = XCol + 1
         Xarchexel.Range("H" & Trim(str(Xlin))).ColumnWidth = 45
         Xarchexel.Cells(Xlin, XCol) = "DIRECCION"
         Xlin = Xlin + 1
         XCol = 1
         Do While Not data_infno.Recordset.EOF
            If IsNull(data_infno.Recordset("cl_codigo")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_codigo")
            Else
               Xarchexel.Cells(Xlin, XCol) = "0"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_apellid")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_apellid")
            Else
               Xarchexel.Cells(Xlin, XCol) = "NN"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_cedula")) = False Then
               If IsNull(data_infno.Recordset("cl_codced")) = False Then
                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_infno.Recordset("cl_cedula"))) & "-" & Trim(str(data_infno.Recordset("cl_codced")))
               Else
                  Xarchexel.Cells(Xlin, XCol) = Trim(str(data_infno.Recordset("cl_cedula"))) & "-0"
               End If
            Else
               Xarchexel.Cells(Xlin, XCol) = "0-0"
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_codconv")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_codconv")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_dpto")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_dpto")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_telefon")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_telefon")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_zona")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_zona")
            End If
            XCol = XCol + 1
            If IsNull(data_infno.Recordset("cl_direcci")) = False Then
               Xarchexel.Cells(Xlin, XCol) = data_infno.Recordset("cl_direcci")
            End If
            data_infno.Recordset.MoveNext
            Xlin = Xlin + 1
            XCol = 1
         Loop
         Xlibexel.Save
         Xlibexel.Close
         Xobjexel.Quit
         Xlabrir.Workbooks.Open Xarchtex, , False
         Xlabrir.Visible = True
         Xlabrir.WindowState = xlMaximized
      End If
   End If
   MsgBox "Proceso terminado"


End If

End Sub

