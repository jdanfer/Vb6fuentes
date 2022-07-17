VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_matester 
   BackColor       =   &H00C0C000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Envío y Recepción material estéril"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9255
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_matester.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_buscamat 
      Caption         =   "data_buscamat"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5160
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_matbd 
      Caption         =   "data_matbd"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6840
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   6360
      Top             =   1800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton b_busca 
      BackColor       =   &H0080FFFF&
      Height          =   375
      Left            =   2880
      Picture         =   "frm_matester.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5520
      Width           =   375
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      Top             =   6600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data data_gri 
      Caption         =   "data_gri"
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
      RecordSource    =   "mant_sol"
      Top             =   5400
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSMask.MaskEdBox mfb 
      Height          =   375
      Left            =   1440
      TabIndex        =   18
      Top             =   5520
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_matester.frx":0B14
      Height          =   2055
      Left            =   120
      OleObjectBlob   =   "frm_matester.frx":0B2B
      TabIndex        =   17
      Top             =   5880
      Width           =   9015
   End
   Begin VB.CommandButton b_imp 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   2520
      Picture         =   "frm_matester.frx":1BA6
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5040
      Width           =   375
   End
   Begin VB.CommandButton b_cance 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   1920
      Picture         =   "frm_matester.frx":2130
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   5040
      Width           =   375
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H00C0FFFF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   1320
      Picture         =   "frm_matester.frx":26BA
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5040
      Width           =   375
   End
   Begin VB.CommandButton b_edita 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   720
      Picture         =   "frm_matester.frx":2C44
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5040
      Width           =   375
   End
   Begin VB.CommandButton b_nuevo 
      BackColor       =   &H00C0FFFF&
      Height          =   375
      Left            =   120
      Picture         =   "frm_matester.frx":31CE
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   5040
      Width           =   375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Datos"
      Enabled         =   0   'False
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9015
      Begin MSAdodcLib.Adodc data_mat 
         Height          =   375
         Left            =   1560
         Top             =   4080
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
         Caption         =   "data_mat"
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
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   3000
         Picture         =   "frm_matester.frx":3758
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Agregar a la lista de materiales"
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox t_cant 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1920
         TabIndex        =   37
         Top             =   1200
         Width           =   735
      End
      Begin VB.ComboBox cbomat 
         BackColor       =   &H00C0FFFF&
         Height          =   360
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   720
         Width           =   6975
      End
      Begin VB.TextBox t_b 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1920
         TabIndex        =   33
         Top             =   2400
         Width           =   1095
      End
      Begin MSMask.MaskEdBox mfctrol 
         Height          =   375
         Left            =   1920
         TabIndex        =   31
         Top             =   3360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H0000FFFF&
         Caption         =   "Control final"
         Enabled         =   0   'False
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   3360
         Width           =   1815
      End
      Begin VB.TextBox t_usret 
         Height          =   375
         Left            =   4920
         TabIndex        =   29
         Top             =   2880
         Width           =   1695
      End
      Begin MSMask.MaskEdBox mhret 
         Height          =   375
         Left            =   4200
         TabIndex        =   28
         Top             =   2880
         Width           =   735
         _ExtentX        =   1296
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   5
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfret 
         Height          =   375
         Left            =   1920
         TabIndex        =   26
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H0080FFFF&
         Caption         =   "Retorno"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   2880
         Width           =   1815
      End
      Begin VB.TextBox labusurec 
         Height          =   375
         Left            =   7200
         MaxLength       =   20
         TabIndex        =   24
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox t_base 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   5640
         TabIndex        =   21
         Top             =   240
         Width           =   615
      End
      Begin MSMask.MaskEdBox mhr 
         Height          =   375
         Left            =   6240
         TabIndex        =   10
         Top             =   2400
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   5
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   3840
         TabIndex        =   8
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   5
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfr 
         Height          =   375
         Left            =   4320
         TabIndex        =   7
         Top             =   2400
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Recibe"
         Height          =   375
         Left            =   3120
         TabIndex        =   6
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox t_material 
         Height          =   840
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1560
         Width           =   6975
      End
      Begin MSMask.MaskEdBox mf 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   240
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cantidad:"
         Height          =   375
         Left            =   120
         TabIndex        =   36
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   2055
         Left            =   6600
         Stretch         =   -1  'True
         Top             =   2880
         Width           =   2295
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Observaciones:"
         Height          =   375
         Left            =   120
         TabIndex        =   34
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Base o Móvil:"
         Height          =   375
         Left            =   120
         TabIndex        =   32
         Top             =   2400
         Width           =   1815
      End
      Begin VB.Label Label8 
         BackColor       =   &H0080FFFF&
         Caption         =   "Hora:"
         Height          =   375
         Left            =   3720
         TabIndex        =   27
         Top             =   2880
         Width           =   615
      End
      Begin VB.Label Label7 
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   6360
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Base:"
         Height          =   375
         Left            =   4920
         TabIndex        =   20
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Hora:"
         Height          =   375
         Left            =   5640
         TabIndex        =   9
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Material:"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Hora:"
         Height          =   375
         Left            =   3240
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Fecha:"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Label labusuario 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   6720
      TabIndex        =   19
      Top             =   5520
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FFFF&
      Caption         =   "Buscar fecha:"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Image Image2 
      Height          =   1095
      Left            =   6600
      Picture         =   "frm_matester.frx":3CE2
      Stretch         =   -1  'True
      Top             =   5280
      Width           =   1935
   End
End
Attribute VB_Name = "frm_matester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_busca_Click()
If mfb.Text <> "__/__/____" Then
   data_gri.RecordSource = "Select * from mant_sol where estado =" & -1 & " and cl_fultpag >=#" & Format(mfb.Text, "yyyy/mm/dd") & "# order by cl_fultpag DESC"
   data_gri.Refresh
   DBGrid1.SetFocus
End If

End Sub

Private Sub b_cance_Click()
      borrar_dat
      b_nuevo.Enabled = True
      b_edita.Enabled = True
      b_graba.Enabled = False
      b_cance.Enabled = False
      b_imp.Enabled = True
      b_busca.Enabled = True
      Frame1.Enabled = False
      XAlta = 0

End Sub

Private Sub b_edita_Click()
If Label7.Caption <> "" Then
   data_mat.RecordSource = "Select * from mant_sol where cl_nro_sup =" & Label7.Caption & " and estado =" & -1
'   data_mat.Recordset.FindFirst "cl_nro_sup =" & Label7.Caption & " and estado =" & -1
   data_mat.Refresh
   If data_mat.Recordset.RecordCount > 0 Then
      borrar_dat
      iguala_dat
      b_nuevo.Enabled = False
      b_edita.Enabled = False
      b_graba.Enabled = True
      b_cance.Enabled = True
      b_imp.Enabled = False
      b_busca.Enabled = False
      Frame1.Enabled = True
      t_material.SetFocus
      XAlta = 0
   End If
End If

End Sub

Private Sub b_graba_Click()
If XAlta = 1 Then
   If t_material.Text <> "" And mf.Text <> "__/__/____" Then
      data_mat.Recordset.AddNew
      data_mat.Recordset("cl_codigo") = Label7.Caption
      data_mat.Recordset("cl_fultpag") = Format(mf.Text, "dd/mm/yyyy")
      data_mat.Recordset("cl_nro_sup") = Label7.Caption
      If mh.Text <> "__:__" Then
         data_mat.Recordset("cl_fax") = mh.Text
      Else
         data_mat.Recordset("cl_fax") = Format(Time, "HH:mm")
      End If
      data_mat.Recordset("cl_val1") = t_base.Text
      data_mat.Recordset("info_debit") = t_material.Text
      data_mat.Recordset("cl_val2") = Check1.value
      If mfr.Text <> "__/__/____" Then
         data_mat.Recordset("cl_fec1") = Format(mfr.Text, "dd/mm/yyyy")
      End If
      If mhr.Text <> "__:__" Then
         data_mat.Recordset("cl_ruc") = mhr.Text
      End If
      If labusuario.Caption <> "" Then
         data_mat.Recordset("cl_nom_sup") = labusuario.Caption
      Else
         data_mat.Recordset("cl_nom_sup") = WElusuario
      End If
      If labusurec.Text <> "" Then
         data_mat.Recordset("cl_descpag") = labusurec.Text
      End If
      data_mat.Recordset("cl_nrovend") = Check2.value
      If mfret.Text <> "__/__/____" Then
         data_mat.Recordset("cl_fultmov") = Format(mfret.Text, "dd/mm/yyyy")
      End If
      If mhret.Text <> "__:__" Then
         data_mat.Recordset("cl_codconv") = Format(mhret.Text, "HH:mm")
      End If
      If t_usret.Text <> "" Then
         data_mat.Recordset("cl_desc2") = t_usret.Text
      End If
      data_mat.Recordset("cl_atrasoa") = Check3.value
      If mfctrol.Text <> "__/__/____" Then
         data_mat.Recordset("cl_fec2") = Format(mfctrol.Text, "dd/mm/yyyy")
      End If
      data_mat.Recordset("estado") = -1
      If t_b.Text <> "" Then
         data_mat.Recordset("cl_val3") = t_b.Text
      End If
      data_mat.Recordset.Update
      data_mat.Refresh
      data_gri.Refresh
      borrar_dat
      b_nuevo.Enabled = True
      b_edita.Enabled = True
      b_graba.Enabled = False
      b_cance.Enabled = False
      b_imp.Enabled = True
      b_busca.Enabled = True
      Frame1.Enabled = False
      XAlta = 0
   End If
Else
'   data_mat.Recordset.EditMode
   data_mat.Recordset("info_debit") = t_material.Text
   data_mat.Recordset("cl_val2") = Check1.value
   If mfr.Text <> "__/__/____" Then
      data_mat.Recordset("cl_fec1") = Format(mfr.Text, "dd/mm/yyyy")
   End If
   If mhr.Text <> "__:__" Then
      data_mat.Recordset("cl_ruc") = mhr.Text
   End If
   If labusuario.Caption <> "" Then
      data_mat.Recordset("cl_nom_sup") = labusuario.Caption
   Else
      data_mat.Recordset("cl_nom_sup") = WElusuario
   End If
   If labusurec.Text <> "" Then
      data_mat.Recordset("cl_descpag") = labusurec.Text
   End If
   data_mat.Recordset("cl_nrovend") = Check2.value
   If mfret.Text <> "__/__/____" Then
      data_mat.Recordset("cl_fultmov") = Format(mfret.Text, "dd/mm/yyyy")
   Else
      If IsNull(data_mat.Recordset("cl_fultmov")) = False Then
         data_mat.Recordset("cl_fultmov") = Null
      End If
   End If
   If mhret.Text <> "__:__" Then
      data_mat.Recordset("cl_codconv") = Format(mhret.Text, "HH:mm")
   Else
      If IsNull(data_mat.Recordset("cl_codconv")) = False Then
         data_mat.Recordset("cl_codconv") = Null
      End If
   End If
   If t_usret.Text <> "" Then
      data_mat.Recordset("cl_desc2") = t_usret.Text
   End If
   data_mat.Recordset("cl_atrasoa") = Check3.value
   If mfctrol.Text <> "__/__/____" Then
      data_mat.Recordset("cl_fec2") = Format(mfctrol.Text, "dd/mm/yyyy")
   Else
      If IsNull(data_mat.Recordset("cl_fec2")) = False Then
         data_mat.Recordset("cl_fec2") = Null
      End If
   End If
   If t_b.Text <> "" Then
      data_mat.Recordset("cl_val3") = t_b.Text
   End If
   data_mat.Recordset.Update
   data_mat.Refresh
   data_gri.Refresh
   borrar_dat
   b_nuevo.Enabled = True
   b_edita.Enabled = True
   b_graba.Enabled = False
   b_cance.Enabled = False
   b_imp.Enabled = True
   b_busca.Enabled = True
   Frame1.Enabled = False
   XAlta = 0
 
End If

End Sub

Private Sub b_imp_Click()
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.Path & "\informes.mdb")

MiBaseact.Execute "Delete * from infcli"

data_inf.RecordSource = "infcli"
data_inf.Refresh

Dim Xdes, Xhas As String
Xdes = InputBox("Ingrese FECHA DESDE:")
Xhas = InputBox("Ingrese FECHA HASTA:")
If Xdes <> "" And Xhas <> "" Then
   data_gri.RecordSource = "Select * from mant_sol where estado =" & -1 & " and cl_fultpag >=#" & Format(Xdes, "yyyy/mm/dd") & "# and cl_fultpag <=#" & Format(Xhas, "yyyy/mm/dd") & "# order by cl_fultpag"
   data_gri.Refresh
   If data_gri.Recordset.RecordCount > 0 Then
      data_gri.Recordset.MoveFirst
      Do While Not data_gri.Recordset.EOF
         data_inf.Recordset.AddNew
         data_inf.Recordset("cl_fecing") = data_gri.Recordset("cl_fultpag")
         data_inf.Recordset("cl_fax") = data_gri.Recordset("cl_fax")
         data_inf.Recordset("cl_nom_sup") = data_gri.Recordset("cl_nom_sup")
         data_inf.Recordset("info_debit") = data_gri.Recordset("info_debit")
         data_inf.Recordset("cl_fnac") = data_gri.Recordset("cl_fec1")
         data_inf.Recordset("cl_codigo") = data_gri.Recordset("cl_val1") 'base
         data_inf.Recordset("cl_fultmov") = data_gri.Recordset("cl_fultmov") ' fecha retorno
         data_inf.Recordset("cl_nomconv") = Mid(data_gri.Recordset("cl_desc2"), 1, 25) 'usuario retorno
         data_inf.Recordset("cl_fultpag") = data_gri.Recordset("cl_fec2") 'fecha control
         data_inf.Recordset("cl_codced") = data_gri.Recordset("cl_val3")
         data_inf.Recordset.Update
         data_gri.Recordset.MoveNext
      Loop
      MsgBox "Terminado"
      cr1.ReportFileName = App.Path & "\infmatest.rpt"
      cr1.ReportTitle = "Material a esterilizar desde: " & Format(Xdes, "dd/mm/yyyy") & " HASTA:" & Format(Xhas, "dd/mm/yyyy")
      cr1.Action = 1
      
   End If

End If

End Sub

Private Sub b_nuevo_Click()
Label7.Caption = Data1.Recordset("nro_material") + 1
b_nuevo.Enabled = False
b_edita.Enabled = False
b_graba.Enabled = True
b_cance.Enabled = True
b_imp.Enabled = False
b_busca.Enabled = False
borrar_dat

Frame1.Enabled = True
t_material.SetFocus
mf.Text = Date
mh.Text = Format(Time, "HH:mm")
t_base.Text = frm_menu.data_parse.Recordset("base")
labusuario.Caption = WElusuario
XAlta = 1
Data1.Recordset.Edit
Data1.Recordset("nro_material") = Data1.Recordset("nro_material") + 1
Data1.Recordset.Update

End Sub

Private Sub cbomat_Click()
data_buscamat.RecordSource = "Select * from material where descrip ='" & cbomat.Text & "'"
data_buscamat.Refresh
If data_buscamat.Recordset.RecordCount > 0 Then
   Image1.Picture = LoadPicture(App.Path & "\picest\" & Trim(Str(data_buscamat.Recordset("id"))) & ".jpg")
Else
   Image1.Picture = LoadPicture(App.Path & "\fotos\sinfoto.jpg")
End If


End Sub

Private Sub cbomat_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_material.SetFocus
End If

End Sub

Private Sub Check1_Click()
If Check1.value = 1 Then
   labusurec.Text = WElusuario
Else
   labusurec.Text = ""
End If

End Sub

Private Sub Check2_Click()
If Check2.value = 1 Then
   t_usret.Text = WElusuario
Else
   t_usret.Text = ""
End If

End Sub

Private Sub Check3_Click()
If Check3.value = 1 Then
   mfctrol.Text = Format(Date, "dd/mm/yyyy")
End If

End Sub

Private Sub Command1_Click()
If t_material.Text = "" Then
   If cbomat.ListIndex >= 0 Then
      If t_cant.Text <> "" Then
         t_material.Text = cbomat.Text & " Cantidad: " & t_cant.Text
      Else
         MsgBox "Ingrese cantidad"
      End If
   Else
      MsgBox "Seleccione Material"
   End If
Else
   If cbomat.ListIndex >= 0 Then
      If t_cant.Text <> "" Then
         t_material.Text = t_material.Text & vbNewLine & cbomat.Text & " Cantidad: " & t_cant.Text
      Else
         MsgBox "Ingrese cantidad"
      End If
   Else
      MsgBox "Seleccione Material"
   End If

End If

End Sub

Private Sub DBGrid1_DblClick()
data_mat.RecordSource = "Select * from mant_sol where cl_nro_sup =" & data_gri.Recordset("cl_nro_sup") & " and estado =" & -1
'data_mat.Recordset.FindFirst "cl_nro_sup =" & data_gri.Recordset("cl_nro_sup") & " and estado =" & -1
data_mat.Refresh
If data_mat.Recordset.RecordCount > 0 Then
   data_mat.Recordset.MoveFirst
   borrar_dat
   iguala_dat
Else
   borrar_dat
End If

End Sub

Private Sub Form_Load()
Dim Xfec As Date
Xfec = Date - 130
data_mat.ConnectionString = "dsn=" & Xconexrmt
data_mat.RecordSource = "Select * from mant_sol where estado =" & -1
data_mat.Refresh

Data1.DatabaseName = App.Path & "\paramb.mdb"
Data1.RecordSource = "paramb"
Data1.Refresh

data_matbd.DatabaseName = App.Path & "\material.mdb"
data_matbd.RecordSource = "Select * from material order by descrip"
data_matbd.Refresh

data_buscamat.DatabaseName = App.Path & "\material.mdb"

data_gri.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_gri.RecordSource = "Select * from mant_sol where estado =" & -1 & " and cl_fultpag >=#" & Format(Xfec, "yyyy/mm/dd") & "# order by cl_fultpag DESC"
data_gri.Refresh

data_inf.DatabaseName = App.Path & "\informes.mdb"
'data_inf.RecordSource = "infcli"
'data_inf.Refresh
cbomat.Clear

If data_matbd.Recordset.RecordCount > 0 Then
   data_matbd.Recordset.MoveFirst
   Do While Not data_matbd.Recordset.EOF
      cbomat.AddItem data_matbd.Recordset("descrip")
      data_matbd.Recordset.MoveNext
   Loop
End If
If WElusuario = "JFERNAN" Or WElusuario = "MCURBELO" Or WElusuario = "CLOVRECICH" Or WElusuario = "CRISTINA" Or WElusuario = "SALVAREZ" Or _
   WElusuario = "VPRESTE" Or WElusuario = "GGARRETA" Then
   Check3.Enabled = True
   mfctrol.Enabled = True
Else
   Check3.Enabled = False
   mfctrol.Enabled = False

End If
End Sub

Public Sub borrar_dat()
mf.Text = "__/__/____"
mh.Text = "__:__"
t_base.Text = ""
t_material.Text = ""
labusuario.Caption = ""
Check1.value = 0
mfr.Text = "__/__/____"
mhr.Text = "__:__"
labusurec.Text = ""
Check2.value = 0
mfret.Text = "__/__/____"
mhret.Text = "__:__"
t_usret.Text = ""
Check3.value = 0
mfctrol.Text = "__/__/____"
t_b.Text = ""
t_cant.Text = ""
cbomat.ListIndex = -1

End Sub

Public Sub iguala_dat()
If IsNull(data_mat.Recordset("cl_nro_sup")) = False Then
   Label7.Caption = data_mat.Recordset("cl_nro_sup")
Else
   Label7.Caption = 1
End If
If IsNull(data_mat.Recordset("cl_fultpag")) = False Then
   mf.Text = Format(data_mat.Recordset("cl_fultpag"), "dd/mm/yyyy")
Else
   mf.Text = "__/__/____"
End If
If IsNull(data_mat.Recordset("cl_fax")) = False Then
   mh.Text = Format(data_mat.Recordset("cl_fax"), "HH:mm")
Else
   mh.Text = "__:__"
End If
If IsNull(data_mat.Recordset("cl_val1")) = False Then
   t_base.Text = data_mat.Recordset("cl_val1")
Else
   t_base.Text = ""
End If
If IsNull(data_mat.Recordset("info_debit")) = False Then
   t_material.Text = data_mat.Recordset("info_debit")
Else
   t_material.Text = ""
End If
If IsNull(data_mat.Recordset("cl_val2")) = False Then
   Check1.value = data_mat.Recordset("cl_val2")
Else
   Check1.value = 0
End If
If IsNull(data_mat.Recordset("cl_fec1")) = False Then
   mfr.Text = Format(data_mat.Recordset("cl_fec1"), "dd/mm/yyyy")
Else
   mfr.Text = "__/__/____"
End If
If IsNull(data_mat.Recordset("cl_ruc")) = False Then
   mhr.Text = Format(data_mat.Recordset("cl_ruc"), "HH:mm")
Else
   mhr.Text = "__:__"
End If
If IsNull(data_mat.Recordset("cl_descpag")) = False Then
   labusurec.Text = data_mat.Recordset("cl_descpag")
Else
   labusurec.Text = ""
End If
If IsNull(data_mat.Recordset("cl_nom_sup")) = False Then
   labusuario.Caption = data_mat.Recordset("cl_nom_sup")
Else
   labusuario.Caption = ""
End If
If IsNull(data_mat.Recordset("cl_nrovend")) = False Then
   Check2.value = data_mat.Recordset("cl_nrovend")
Else
   Check2.value = 0
End If
If IsNull(data_mat.Recordset("cl_fultmov")) = False Then
   mfret.Text = Format(data_mat.Recordset("cl_fultmov"), "dd/mm/yyyy")
Else
   mfret.Text = "__/__/____"
End If
If IsNull(data_mat.Recordset("cl_codconv")) = False Then
   mhret.Text = Format(data_mat.Recordset("cl_codconv"), "HH:mm")
Else
   mhret.Text = "__:__"
End If
If IsNull(data_mat.Recordset("cl_desc2")) = False Then
   t_usret.Text = data_mat.Recordset("cl_desc2")
Else
   t_usret.Text = ""
End If
If IsNull(data_mat.Recordset("cl_atrasoa")) = False Then
   Check3.value = data_mat.Recordset("cl_atrasoa")
Else
   Check3.value = 0
End If
If IsNull(data_mat.Recordset("cl_fec2")) = False Then
   mfctrol.Text = Format(data_mat.Recordset("cl_fec2"), "dd/mm/yyyy")
Else
   mfctrol.Text = "__/__/____"
End If
If IsNull(data_mat.Recordset("cl_val3")) = False Then
   t_b.Text = data_mat.Recordset("cl_val3")
Else
   t_b.Text = ""
End If


End Sub

Private Sub Form_Resize()
With Image2
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub mfctrol_GotFocus()
If mfctrol.Text = "__/__/____" Then
   mfctrol.Text = Format(Date, "dd/mm/yyyy")
End If

End Sub

Private Sub mfr_GotFocus()
If mfr.Text = "__/__/____" Then
   mfr.Text = Format(Date, "dd/mm/yyyy")
End If

End Sub

Private Sub mfret_GotFocus()
If mfret.Text = "__/__/____" Then
   mfret.Text = Format(Date, "dd/mm/yyyy")
End If

End Sub

Private Sub mhr_GotFocus()
If mhr.Text = "__:__" Then
   mhr.Text = Format(Time, "HH:mm")
End If

End Sub

Private Sub mhret_GotFocus()
If mhret.Text = "__:__" Then
   mhret.Text = Format(Time, "HH:mm")
End If

End Sub

Private Sub t_b_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfr.SetFocus
End If

End Sub

Private Sub t_cant_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_material.SetFocus
End If

End Sub
