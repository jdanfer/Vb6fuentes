VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_mejoracons 
   BackColor       =   &H0080FF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultar y/o modificar datos de la acción"
   ClientHeight    =   8820
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   10080
   Icon            =   "frm_mejoracons.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   10080
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   9360
      Picture         =   "frm_mejoracons.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Nuevo registro"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3480
      Picture         =   "frm_mejoracons.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Editar cuadro de análisis y descripción"
      Top             =   6000
      Width           =   495
   End
   Begin VB.Data data_his2 
      Caption         =   "data_his2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6720
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data data_histo 
      Caption         =   "data_histo"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6720
      Visible         =   0   'False
      Width           =   3135
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_mejoracons.frx":0E0E
      Height          =   1935
      Left            =   120
      OleObjectBlob   =   "frm_mejoracons.frx":0E27
      TabIndex        =   18
      Top             =   6600
      Width           =   9735
   End
   Begin VB.CommandButton b_cancela 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      Picture         =   "frm_mejoracons.frx":1B52
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Cancelar acción"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      Picture         =   "frm_mejoracons.frx":20DC
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Guardar registro"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton b_modif 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   960
      Picture         =   "frm_mejoracons.frx":2666
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Editar registro"
      Top             =   6000
      Width           =   495
   End
   Begin VB.CommandButton b_nuevo 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      Picture         =   "frm_mejoracons.frx":2BF0
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Nuevo registro"
      Top             =   6000
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Historial de movimientos de la acción"
      Enabled         =   0   'False
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
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9735
      Begin MSAdodcLib.Adodc data1 
         Height          =   375
         Left            =   4200
         Top             =   4200
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
      Begin MSAdodcLib.Adodc data_grabahis 
         Height          =   375
         Left            =   3360
         Top             =   3720
         Visible         =   0   'False
         Width           =   2775
         _ExtentX        =   4895
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
         Caption         =   "data_grabahis"
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
      Begin VB.Data data_param 
         Caption         =   "data_param"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   6360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   5040
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.TextBox txt_anali 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   2520
         Width           =   7335
      End
      Begin MSMask.MaskEdBox mfecter 
         Height          =   375
         Left            =   3840
         TabIndex        =   25
         Top             =   5400
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Proceso terminado"
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
         Left            =   240
         TabIndex        =   24
         Top             =   5400
         Width           =   3255
      End
      Begin VB.TextBox txt_plazo 
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
         Left            =   2280
         TabIndex        =   8
         Top             =   4920
         Width           =   1215
      End
      Begin VB.TextBox txt_dethis 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   2280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   3600
         Width           =   7335
      End
      Begin MSMask.MaskEdBox mhorahis 
         Height          =   375
         Left            =   6000
         TabIndex        =   12
         Top             =   1560
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfechis 
         Height          =   375
         Left            =   2280
         TabIndex        =   10
         Top             =   1560
         Width           =   1815
         _ExtentX        =   3201
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
         ItemData        =   "frm_mejoracons.frx":317A
         Left            =   2280
         List            =   "frm_mejoracons.frx":318D
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   2040
         Width           =   3735
      End
      Begin VB.Label labid2da 
         Height          =   375
         Left            =   6000
         TabIndex        =   31
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label labidd 
         Height          =   375
         Left            =   4560
         TabIndex        =   28
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Análisis de Causas:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   26
         Top             =   2520
         Width           =   2055
      End
      Begin VB.Label labpla 
         BackColor       =   &H00C0FFC0&
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
         Left            =   3840
         TabIndex        =   23
         Top             =   4920
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Plazo (días):"
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
         Left            =   240
         TabIndex        =   22
         Top             =   4920
         Width           =   2055
      End
      Begin VB.Label labus 
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
         Left            =   7080
         TabIndex        =   21
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label2 
         Caption         =   "Usuario actual:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7080
         TabIndex        =   20
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Descripción de la acción."
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
         Left            =   240
         TabIndex        =   13
         Top             =   3600
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "HORA:"
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
         Left            =   4560
         TabIndex        =   11
         Top             =   1560
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "FECHA:"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Acción:"
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
         Left            =   240
         TabIndex        =   5
         Top             =   2040
         Width           =   2055
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000A&
         BorderWidth     =   3
         X1              =   0
         X2              =   9720
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Label labtit 
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
         Left            =   2280
         TabIndex        =   4
         Top             =   960
         Width           =   7215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "TITULO:"
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
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   2055
      End
      Begin VB.Label labnro 
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
         Left            =   2280
         TabIndex        =   2
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "NUMERO:"
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
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Doble click para editar"
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
      Left            =   120
      TabIndex        =   19
      Top             =   8520
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   6240
      Picture         =   "frm_mejoracons.frx":31CC
      Stretch         =   -1  'True
      Top             =   6240
      Width           =   2415
   End
End
Attribute VB_Name = "frm_mejoracons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_cancela_Click()
XAlta = 0
Combo1.ListIndex = 0
b_nuevo.Enabled = True
b_graba.Enabled = False
b_modif.Enabled = True
b_cancela.Enabled = False
DBGrid1.Enabled = True
mfechis.Text = "__/__/____"
mhorahis.Text = "__:__"
txt_dethis.Text = ""
txt_anali.Text = ""
txt_plazo.Text = ""
labpla.Caption = ""
Check1.value = 0
mfecter.Text = "__/__/____"
Frame1.Enabled = False
Command1.Enabled = True


End Sub

Private Sub b_graba_Click()
'On Error GoTo queerres

If XAlta = 1 Then
   If Combo1.ListIndex <> -1 Then
      If txt_dethis.Text <> "" Then
         data_grabahis.Recordset.AddNew
         data_grabahis.Recordset("cl_etiquet") = 0
         data_grabahis.Recordset("estado") = 99
         data_grabahis.Recordset("cl_codigo") = labidd.Caption
         If Combo1.ListIndex <> -1 Then
            data_grabahis.Recordset("cl_descpag") = Combo1.Text
         End If
         data_grabahis.Recordset("cl_val2") = 7
         data_grabahis.Recordset("cl_nro_sup") = Combo1.ListIndex
         data_grabahis.Recordset("cl_nrovend") = labnro.Caption
         data_grabahis.Recordset("cl_nomcobr") = 1
         If mfechis.Text <> "__/__/____" Then
            data_grabahis.Recordset("cl_fultpag") = Format(mfechis.Text, "dd/mm/yyyy")
         End If
         data_grabahis.Recordset("cl_fax") = mhorahis.Text
         data_grabahis.Recordset("info_debit") = txt_dethis.Text
         data_grabahis.Recordset("cl_nom_sup") = labus.Caption
         If txt_plazo.Text = "" Then
            txt_plazo.Text = 0
            labpla.Caption = Date
         End If
         data_grabahis.Recordset("cl_atrasoa") = txt_plazo.Text
         If labpla.Caption <> "" Then
            data_grabahis.Recordset("cl_fec1") = Format(labpla.Caption, "dd/mm/yyyy")
         End If
         data_grabahis.Recordset("cl_val3") = Check1.value
         If mfecter.Text <> "__/__/____" Then
            data_grabahis.Recordset("cl_fec2") = mfecter.Text
         End If
         data_grabahis.Recordset.Update
         data_grabahis.Refresh
         data_histo.Refresh
         If txt_anali.Text <> "" Then
            labidd.Caption = data_param.Recordset("nro_mam2") + 1
            data_param.Recordset.Edit
            data_param.Recordset("nro_mam2") = data_param.Recordset("nro_mam2") + 1
            data_param.Recordset.Update
            
            data_grabahis.Recordset.AddNew
            data_grabahis.Recordset("cl_nrovend") = labnro.Caption
            data_grabahis.Recordset("cl_etiquet") = 0
            data_grabahis.Recordset("estado") = 98
            data_grabahis.Recordset("cl_codigo") = labidd.Caption
            data_grabahis.Recordset("info_debit") = txt_anali.Text
            data_grabahis.Recordset("cl_nomcobr") = 1
            data_grabahis.Recordset.Update
         Else
            labidd.Caption = data_param.Recordset("nro_mam2") + 1
            data_param.Recordset.Edit
            data_param.Recordset("nro_mam2") = data_param.Recordset("nro_mam2") + 1
            data_param.Recordset.Update
            
            data_grabahis.Recordset.AddNew
            data_grabahis.Recordset("cl_nrovend") = labnro.Caption
            data_grabahis.Recordset("cl_etiquet") = 0
            data_grabahis.Recordset("estado") = 98
            data_grabahis.Recordset("cl_codigo") = labidd.Caption
            data_grabahis.Recordset("cl_nomcobr") = 1
            data_grabahis.Recordset.Update
         End If
         data_grabahis.Refresh
         data_histo.Refresh
         If Check1.value = 1 Then
            Data1.RecordSource = "Select * from infor_sol where estado =" & labnro.Caption & " and cl_nomcobr =" & 1
            Data1.Refresh
            If Data1.Recordset.RecordCount > 0 Then
               If IsNull(Data1.Recordset("cl_val3")) = False Then
                  If Data1.Recordset("cl_val3") = Check1.value Then
                  Else
'                    data1.Recordset.Edit
                    Data1.Recordset("cl_val3") = Check1.value
                    Data1.Recordset.Update
                  End If
               End If
               If IsNull(Data1.Recordset("cl_codconv")) = False Then
                  If Data1.Recordset("cl_codconv") = "C" Then
                  Else
'                     data1.Recordset.Edit
                     Data1.Recordset("cl_codconv") = "C"
                     Data1.Recordset.Update
                  End If
               End If
            End If
         Else
            Data1.RecordSource = "Select * from infor_sol where estado =" & labnro.Caption & " and cl_nomcobr =" & 1
            Data1.Refresh
'            data1.Recordset.FindFirst "estado =" & labnro.Caption & " and cl_nomcobr =" & 1
            If Data1.Recordset.RecordCount > 0 Then
               If IsNull(Data1.Recordset("cl_val3")) = False Then
                  If Data1.Recordset("cl_val3") = Check1.value Then
                  Else
'                    data1.Recordset.Edit
                    Data1.Recordset("cl_val3") = Check1.value
                    Data1.Recordset.Update
                  End If
               End If
               If IsNull(Data1.Recordset("cl_codconv")) = False Then
                  If Data1.Recordset("cl_codconv") = Null Then
                  Else
'                     data1.Recordset.Edit
                     Data1.Recordset("cl_codconv") = Null
                     Data1.Recordset.Update
                  End If
               End If
            End If
         
         End If
         XAlta = 0
         Combo1.ListIndex = -1
         b_nuevo.Enabled = True
         b_graba.Enabled = False
         b_modif.Enabled = True
         b_cancela.Enabled = False
         DBGrid1.Enabled = True
         mfechis.Text = "__/__/____"
         mhorahis.Text = "__:__"
         txt_dethis.Text = ""
         txt_anali.Text = ""
         txt_plazo.Text = ""
         labpla.Caption = ""
         Check1.value = 0
         mfecter.Text = "__/__/____"
         Frame1.Enabled = False
         DBGrid1.SetFocus
      Else
         MsgBox "Ingrese una descripción del análisis"
      End If
   Else
      MsgBox "Seleccione proceso"
   End If
Else
   data_histo.Recordset.Edit
   If txt_plazo.Text = "" Then
      data_histo.Recordset("cl_atrasoa") = 0
      data_histo.Recordset("cl_fec1") = Null
   Else
      data_histo.Recordset("cl_atrasoa") = txt_plazo.Text
      If labpla.Caption <> "" Then
         data_histo.Recordset("cl_fec1") = Format(labpla.Caption, "dd/mm/yyyy")
      Else
         data_histo.Recordset("cl_fec1") = Null
      End If
   End If
   If Combo1.ListIndex <> -1 Then
      data_histo.Recordset("cl_descpag") = Combo1.Text
   Else
      data_histo.Recordset("cl_descpag") = Null
   End If
   data_histo.Recordset("cl_nro_sup") = Combo1.ListIndex
   data_histo.Recordset("info_debit") = txt_dethis.Text
   data_histo.Recordset("cl_nom_sup") = labus.Caption
    data_histo.Recordset("cl_val3") = Check1.value
    If mfecter.Text <> "__/__/____" Then
       data_histo.Recordset("cl_fec2") = mfecter.Text
    Else
       data_histo.Recordset("cl_fec2") = Null
    End If
   data_histo.Recordset.Update
   data_histo.Refresh
   data_grabahis.Refresh
   If Check1.value = 1 Then
      Data1.RecordSource = "Select * from infor_sol where estado =" & labnro.Caption & " and cl_nomcobr =" & 1
      Data1.Refresh
      If Data1.Recordset.RecordCount > 0 Then
         If IsNull(Data1.Recordset("cl_val3")) = False Then
            If Data1.Recordset("cl_val3") = Check1.value Then
            Else
'               data1.Recordset.Edit
               Data1.Recordset("cl_val3") = Check1.value
               Data1.Recordset.Update
            End If
         End If
         If IsNull(Data1.Recordset("cl_codconv")) = False Then
            If Data1.Recordset("cl_codconv") = "C" Then
            Else
'               data1.Recordset.Edit
               Data1.Recordset("cl_codconv") = "C"
               Data1.Recordset.Update
            End If
         End If
      End If
   Else
      Data1.RecordSource = "Select * from infor_sol where estado =" & labnro.Caption & " and cl_nomcobr =" & 1
      Data1.Refresh
      If Data1.Recordset.RecordCount > 0 Then
         If IsNull(Data1.Recordset("cl_val3")) = False Then
            If Data1.Recordset("cl_val3") = Check1.value Then
            Else
'              data1.Recordset.Edit
              Data1.Recordset("cl_val3") = Check1.value
              Data1.Recordset.Update
            End If
         End If
         If IsNull(Data1.Recordset("cl_codconv")) = False Then
            If Data1.Recordset("cl_codconv") = Null Then
            Else
'               data1.Recordset.Edit
               Data1.Recordset("cl_codconv") = Null
               Data1.Recordset.Update
            End If
         End If
      End If
   End If
   
   data_his2.Connect = "odbc;dsn=" & Xconexrmt & ";"
'   data_his2.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 1 & " order by cl_codigo"
'   data_his2.Refresh
   
   labid2da.Caption = data_param.Recordset("nro_mam2") + 1
   data_param.Recordset.Edit
   data_param.Recordset("nro_mam2") = data_param.Recordset("nro_mam2") + 1
   data_param.Recordset.Update
   
   
'   If data_his2.Recordset.RecordCount > 0 Then
'      data_his2.Recordset.MoveLast
'      labidd.Caption = data_his2.Recordset("cl_codigo") + 1
'   End If
   data_his2.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 1 & " and cl_nrovend =" & labnro.Caption & " and estado =" & 98
   data_his2.Refresh
   If data_his2.Recordset.RecordCount > 0 Then
      If IsNull(data_his2.Recordset("info_debit")) = False Then
         If data_his2.Recordset("info_debit") = txt_anali.Text Then
         Else
            If txt_anali.Text = "" Then
               data_his2.Recordset.Edit
               data_his2.Recordset("info_debit") = Null
               data_his2.Recordset.Update
            Else
               data_his2.Recordset.Edit
               data_his2.Recordset("info_debit") = txt_anali.Text
               data_his2.Recordset.Update
            End If
         End If
      Else
         If txt_anali.Text <> "" Then
            data_his2.Recordset.Edit
            data_his2.Recordset("info_debit") = txt_anali.Text
            data_his2.Recordset.Update
         End If
      End If
   Else
'      labidd.Caption = labidd.Caption + 1
      data_his2.Recordset.AddNew
      data_his2.Recordset("cl_nrovend") = labnro.Caption
      data_his2.Recordset("cl_etiquet") = 0
      data_his2.Recordset("estado") = 98
      data_his2.Recordset("cl_codigo") = labid2da.Caption
      data_his2.Recordset("info_debit") = txt_anali.Text
      data_his2.Recordset("cl_nomcobr") = 1
      data_his2.Recordset.Update
   End If
   data_grabahis.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 1 & " and cl_nrovend =" & labnro.Caption & " and estado =" & 99
   data_grabahis.Refresh
   XAlta = 0
   mfechis.Enabled = True
   mhorahis.Enabled = True
   Combo1.Enabled = True
   txt_plazo.Enabled = True
   Combo1.ListIndex = -1
   b_nuevo.Enabled = True
   b_graba.Enabled = False
   b_modif.Enabled = True
   b_cancela.Enabled = False
   DBGrid1.Enabled = True
   mfechis.Text = "__/__/____"
   mhorahis.Text = "__:__"
   txt_dethis.Text = ""
   txt_anali.Text = ""
   txt_plazo.Text = ""
   labpla.Caption = ""
   Check1.value = 0
   mfecter.Text = "__/__/____"
   Frame1.Enabled = False
   DBGrid1.SetFocus
   
End If

'Exit Sub
'queerres:
'     If Err.Number = 3197 Then
'        MsgBox "Atención! No hay modificaciones para GRABAR, VERIFIQUE o cancele la acción", vbInformation, "SAPP"
'     Else
'        MsgBox "Atención! Verifique los datos " & Err.Description
'     End If

End Sub

Private Sub b_modif_Click()

frm_mejora.Frame1.Enabled = True
If Combo1.ListIndex < 0 And txt_anali.Text = "" And txt_dethis.Text = "" Then
   MsgBox "NO HAY DATOS SELECCIONADOS PARA MODIFICAR, VERIFIQUE!", vbCritical
   frm_mejora.Frame1.Enabled = False
Else
    If frm_mejora.Combo2.ListIndex >= 0 And WElusuario <> "COMPUTOS" Then
       MsgBox "ATENCION! EL REGISTRO YA FUE CERRADO", vbInformation, "Mejora continua"
       frm_mejora.Frame1.Enabled = False
    Else
    
        If WElusuario = frm_mejora.data_accion.Recordset("cl_nom_sup") Or WElusuario = "SPEREZ" Or WElusuario = "SDOMINGUEZ" Or WElusuario = "AGUILLEN" Or WElusuario = "COMPUTOS" Or WElusuario = "MEUGENIA" Then
            XAlta = 0
            Frame1.Enabled = True
            b_nuevo.Enabled = False
            b_graba.Enabled = True
            b_modif.Enabled = False
            b_cancela.Enabled = True
            DBGrid1.Enabled = False
            mfechis.Enabled = False
            mhorahis.Enabled = False
        '    Combo1.Enabled = False
        '    txt_plazo.Enabled = False
            txt_dethis.SetFocus
        Else
            MsgBox "NO ES EL USUARIO PROPIETARIO DE LA ACCION", vbCritical
            DBGrid1.SetFocus
        End If
    End If
End If

End Sub

Private Sub b_nuevo_Click()
If data_histo.Recordset.RecordCount >= 1 Then
   MsgBox "Ya existe acción ingresada, realice modificación.", vbInformation
Else
    frm_mejora.Frame1.Enabled = True
    If frm_mejora.Combo2.ListIndex >= 0 Then
       MsgBox "ATENCION! EL REGISTRO YA FUE CERRADO", vbInformation, "Mejora continua"
       frm_mejora.Frame1.Enabled = False
    Else
       frm_mejora.Frame1.Enabled = False
    '   frm_mejoracons.Show vbModal
    
        Data1.ConnectionString = "dsn=" & Xconexrmt
        Data1.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 1 & " order by cl_codigo"
        Data1.Refresh
    '    If Data1.Recordset.RecordCount > 0 Then
    '       Data1.Recordset.MoveLast
        labidd.Caption = data_param.Recordset("nro_mam2") + 1
        data_param.Recordset.Edit
        data_param.Recordset("nro_mam2") = data_param.Recordset("nro_mam2") + 1
        data_param.Recordset.Update
        
    '    Else
    '       labidd.Caption = 10001
    '    End If
        
        If WElusuario = frm_mejora.data_accion.Recordset("cl_nom_sup") Or WElusuario = "SPEREZ" Or WElusuario = "SDOMINGUEZ" Or WElusuario = "AGUILLEN" Or WElusuario = "DARIOH" Or WElusuario = "MEUGENIA" Then
            Frame1.Enabled = True
            XAlta = 1
            Combo1.ListIndex = 0
            b_nuevo.Enabled = False
            b_graba.Enabled = True
            b_modif.Enabled = False
            b_cancela.Enabled = True
            DBGrid1.Enabled = False
            mfechis.Text = "__/__/____"
            mhorahis.Text = "__:__"
            txt_dethis.Text = ""
            txt_anali.Text = ""
            txt_plazo.Text = ""
            labpla.Caption = ""
            Check1.value = 0
            mfecter.Text = "__/__/____"
            Combo1.SetFocus
            Combo1.ListIndex = 0
            mfechis.Text = Format(Date, "dd/mm/yyyy")
            mhorahis.Text = Format(Time, "HH:mm")
        Else
            MsgBox "NO ES EL USUARIO PROPIETARIO DE LA ACCION", vbCritical
            DBGrid1.SetFocus
        End If
    End If
End If

End Sub

Private Sub Check1_Click()
If Check1.value = 1 Then
   mfecter.Text = Date
Else
   mfecter.Text = "__/__/____"
End If

End Sub

Private Sub Combo1_Click()
If Combo1.ListIndex = 2 Then
   MsgBox "No puede seleccionar ésta acción"
   Combo1.ListIndex = 0
End If

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_anali.SetFocus
End If

End Sub

Private Sub Command1_Click()
b_nuevo.Enabled = False
b_modif.Enabled = False
b_graba.Enabled = False
b_cancela.Enabled = True
DBGrid1.Enabled = False
Frame1.Enabled = True
Combo1.Enabled = False
mfechis.Enabled = False
mhorahis.Enabled = False
txt_anali.Enabled = True
txt_dethis.Enabled = True
txt_plazo.Enabled = False
Check1.Enabled = False
mfecter.Enabled = False
txt_dethis.SetFocus
Command1.Enabled = False

End Sub

Private Sub Command2_Click()
Dim Xloborra, Xlaclave As String
Xlaclave = InputBox("Ingrese clave de confirmación")

If WElusuario = "SPEREZ" And Xlaclave = WxclaveU Then
   Xloborra = MsgBox("Borrar?" & data_histo.Recordset("cl_codigo"), vbInformation + vbYesNo)
   If Xloborra = vbYes Then
      data_histo.Recordset.Delete
      data_histo.Refresh
      data_his2.Recordset.Delete
      data_his2.Refresh
      Unload Me
   End If
Else
   If WElusuario = "AGUILLEN" And UCase(Xlaclave) = WxclaveU Then
      Xloborra = MsgBox("Borrar?" & data_histo.Recordset("cl_codigo"), vbInformation + vbYesNo)
      If Xloborra = vbYes Then
         data_histo.Recordset.Delete
         data_histo.Refresh
         data_his2.Recordset.Delete
         data_his2.Refresh
         Unload Me
      End If
   Else
      MsgBox "Usuario no autorizado"
   End If
End If

End Sub

Private Sub DBGrid1_DblClick()
If IsNull(data_histo.Recordset("cl_nro_sup")) = False Then
   Combo1.ListIndex = data_histo.Recordset("cl_nro_sup")
Else
   Combo1.ListIndex = -1
End If
If IsNull(data_histo.Recordset("cl_fultpag")) = False Then
   mfechis.Text = Format(data_histo.Recordset("cl_fultpag"), "dd/mm/yyyy")
Else
   mfechis.Text = "__/__/____"
End If
If IsNull(data_histo.Recordset("cl_fax")) = False Then
   mhorahis.Text = Format(data_histo.Recordset("cl_fax"), "HH:mm")
Else
   mhorahis.Text = "__:__"
End If
If IsNull(data_histo.Recordset("info_debit")) = False Then
   txt_dethis.Text = data_histo.Recordset("info_debit")
Else
   txt_dethis.Text = ""
End If
If IsNull(data_histo.Recordset("cl_atrasoa")) = False Then
   txt_plazo.Text = data_histo.Recordset("cl_atrasoa")
Else
   txt_plazo.Text = ""
End If
If IsNull(data_histo.Recordset("cl_fec1")) = False Then
   labpla.Caption = data_histo.Recordset("cl_fec1")
Else
   labpla.Caption = ""
End If
If IsNull(data_histo.Recordset("cl_val3")) = False Then
   Check1.value = data_histo.Recordset("cl_val3")
Else
   Check1.value = 0
End If
If IsNull(data_histo.Recordset("cl_fec2")) = False Then
   If data_histo.Recordset("cl_fec2") <> "" Then
      mfecter.Text = data_histo.Recordset("cl_fec2")
   End If
Else
   mfecter.Text = "__/__/____"
End If
   
'data_his2.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 1 & " and cl_nrovend =" & labnro.Caption & " and estado =" & 98 & " and cl_codigo =" & data_histo.Recordset("cl_codigo")
data_his2.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 1 & " and cl_nrovend =" & labnro.Caption & " and estado =" & 98

data_his2.Refresh
If data_his2.Recordset.RecordCount > 0 Then
   If IsNull(data_his2.Recordset("info_debit")) = False Then
      txt_anali.Text = data_his2.Recordset("info_debit")
   Else
      txt_anali.Text = ""
   End If
Else
   txt_anali.Text = ""
End If


End Sub

Private Sub Form_Load()
labnro.Caption = frm_mejora.txt_nro.Text
labtit.Caption = frm_mejora.txt_encab.Text
If labnro.Caption = "" Then
   MsgBox "No existen registros"
Else
    data_histo.Connect = "odbc;dsn=" & Xconexrmt & ";"
    data_histo.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 1 & " and cl_nrovend =" & labnro.Caption & " and estado =" & 99
    data_histo.Refresh
    data_grabahis.ConnectionString = "dsn=" & Xconexrmt
    data_grabahis.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 1 & " and cl_nrovend =" & labnro.Caption & " and estado =" & 99
    data_grabahis.Refresh
    data_his2.Connect = "odbc;dsn=" & Xconexrmt & ";"
    data_his2.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 1 & " and cl_nrovend =" & labnro.Caption & " and estado =" & 98
    data_his2.Refresh
    Data1.ConnectionString = "dsn=" & Xconexrmt
    Data1.RecordSource = "Select * from infor_sol where cl_nomcobr =" & 1
    Data1.Refresh
    labus.Caption = WElusuario
    If data_histo.Recordset.RecordCount > 0 Then
    Else
       MsgBox "No se han ingresado movimientos para ésta mejora", vbInformation
    End If
End If

data_param.DatabaseName = App.Path & "\paramam.mdb"
data_param.RecordSource = "parsec0"
data_param.Refresh


End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub mfechis_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mhorahis.SetFocus
End If

End Sub

Private Sub mhorahis_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_dethis.SetFocus
End If

End Sub

Private Sub txt_anali_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_dethis.SetFocus
End If

End Sub

Private Sub txt_dethis_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   b_graba.SetFocus
End If

End Sub



Private Sub txt_plazo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_graba.SetFocus
End If

End Sub

Private Sub txt_plazo_LostFocus()
If txt_plazo.Text <> "" Then
   labpla.Caption = Date + txt_plazo.Text
Else
   txt_plazo.Text = 0
   labpla.Caption = Date
End If

End Sub
