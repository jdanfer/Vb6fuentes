VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_aumentos 
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aumento de Convenios y Estudios"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9300
   Icon            =   "frm_aumentos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   9300
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView ListView1 
      Height          =   1455
      Left            =   4440
      TabIndex        =   16
      Top             =   3480
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   2566
      View            =   3
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "COD"
         Object.Width           =   952
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "NOMBRE"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.Data data_fam 
      Caption         =   "data_fam"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data data_respara 
      Caption         =   "data_respara"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5520
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Data data_aran 
      Caption         =   "data_aran"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6000
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data data_prec 
      Caption         =   "data_prec"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "cnv_prec"
      Top             =   5760
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_ant 
      Caption         =   "data_ant"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_act 
      Caption         =   "Data_act"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton b_can 
      BackColor       =   &H00FFFFFF&
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
      Height          =   735
      Left            =   7080
      Picture         =   "frm_aumentos.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton b_acep 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ACEPTAR"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      Picture         =   "frm_aumentos.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5280
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00404040&
      Caption         =   "Datos para aumento"
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
      Height          =   4815
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   8775
      Begin VB.CheckBox Check4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aumentar solo servicios AP"
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
         Left            =   4920
         TabIndex        =   19
         Top             =   2040
         Width           =   3495
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H0080FFFF&
         Caption         =   "No aumentar aranceles"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   17
         Top             =   3840
         Width           =   2535
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aumento de servicios por FAMILIA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   315
         Left            =   240
         TabIndex        =   14
         ToolTipText     =   "Cuando se aumentan los servicios también aumentan los aranceles"
         Top             =   3240
         Width           =   3975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   495
         Left            =   3960
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin MSMask.MaskEdBox mdesde 
         Height          =   375
         Left            =   2280
         TabIndex        =   12
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin MSAdodcLib.Adodc adohist 
         Height          =   330
         Left            =   4800
         Top             =   120
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
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
         Caption         =   "adohist"
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
      Begin VB.TextBox mpor 
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
         Left            =   2280
         TabIndex        =   1
         ToolTipText     =   "IMPORTANTE! Tenga en cuenta que el separador de decimales en su PC debe estar configurado la COMA y NO el PUNTO"
         Top             =   960
         Width           =   1455
      End
      Begin VB.TextBox t_col 
         Alignment       =   2  'Center
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
         Left            =   3480
         MaxLength       =   1
         TabIndex        =   10
         Top             =   2040
         Width           =   735
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aumentar por COLOR:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   2040
         Width           =   3255
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aumentar SOLO San Jacinto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   4680
         TabIndex        =   8
         Top             =   2040
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aumento de Servicios (TODOS)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   4440
         TabIndex        =   5
         ToolTipText     =   "Cuando se aumentan los servicios también aumentan los aranceles"
         Top             =   2760
         Width           =   3975
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aumento de Convenios"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   2760
         Width           =   3975
      End
      Begin VB.Label Label5 
         BackColor       =   &H0080FFFF&
         Caption         =   "La separación de decimales puede ser una coma o un punto (dependiendo de la configuración que tenga la PC)"
         Height          =   615
         Left            =   240
         TabIndex        =   18
         ToolTipText     =   "Ante dudas de configuración, consultar con TI"
         Top             =   1320
         Width           =   3495
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Desmarque los que no aumentarán."
         Height          =   735
         Left            =   2880
         TabIndex        =   15
         Top             =   3960
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha de comienzo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ejemplo: para un aumento de 5,4% ingrese 1,054"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   3840
         TabIndex        =   3
         Top             =   960
         Width           =   4095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Aumento:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   2055
      End
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   120
      Picture         =   "frm_aumentos.frx":0F56
      Stretch         =   -1  'True
      Top             =   3960
      Width           =   1095
   End
End
Attribute VB_Name = "frm_aumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_acep_Click()
Dim XXR As String
Dim Mifec As Date
Dim Mivalor As Double
Dim Xcompfecha As Date
Dim Xfecfin As Date
Dim Xnrofliaa As Integer
Dim x As Integer

On Error GoTo Quepasaaumento

Xnrofliaa = 0

Xfecfin = CDate(mdesde.Text) - 1

If Option1.Value = True Then
   b_acep.Enabled = False
   XXR = MsgBox("Se cambiaran los precios de Convenios, Desea Continuar?", vbExclamation + vbOKCancel, "Mensaje")
   MsgBox "Se Guardará un respaldo en ESTYCONVANT.MDB", vbInformation, "Mensaje"
   
   If XXR = vbOK Then
      frm_aumentos.MousePointer = 11
      
      data_act.RecordSource = "convenio"
      data_act.Refresh
      data_ant.RecordSource = "convant"
      data_ant.Refresh
      If Check4.Value = 1 Then
      
      Else
         If data_ant.Recordset.RecordCount > 0 Then
            data_ant.Recordset.MoveFirst
            Do While Not data_ant.Recordset.EOF
               data_ant.Recordset.Delete
               data_ant.Recordset.MoveNext
            Loop
         End If
         data_act.Recordset.MoveFirst
         Do While Not data_act.Recordset.EOF
            data_ant.Recordset.AddNew
            data_ant.Recordset("cnv_codigo") = data_act.Recordset("cnv_codigo")
            data_ant.Recordset("cnv_desc") = Mid(data_act.Recordset("cnv_desc"), 1, 65)
            data_ant.Recordset("cnv_precio") = data_act.Recordset("cnv_precio")
            data_ant.Recordset.Update
            data_act.Recordset.MoveNext
         Loop
         frm_aumentos.MousePointer = 0
         MsgBox "Se guardó un respaldo del archivo de convenios anterior...", vbInformation, "Mensaje"
         frm_aumentos.MousePointer = 11
      End If
      data_act.RecordSource = "select * from convenio where cnv_hasta >=#" & Format(Date, "yyyy/mm/dd") & "#"
      data_act.Refresh
      data_act.Recordset.MoveFirst
      Do While Not data_act.Recordset.EOF
         If Check2.Value = 1 Then
            If t_col.Text <> "" Then
               If IsNull(data_act.Recordset("cnv_colrec")) = False Then
                  If Trim(data_act.Recordset("cnv_colrec")) = Trim(UCase(t_col.Text)) Then
                     If IsNull(data_act.Recordset("cnv_precio")) = False Then
                        If data_act.Recordset("cnv_precio") <> 0 Then
                           data_act.Recordset.Edit
                           Mivalor = data_act.Recordset("cnv_precio") * mpor.Text
                           data_act.Recordset("cnv_precio") = Round(Mivalor)
                           data_act.Recordset.Update
                           data_prec.RecordSource = "Select * from cnv_prec where cnv_codigo ='" & data_act.Recordset("cnv_codigo") & "' order by cnv_codigo,cnv_desde,cnv_hasta"
                           data_prec.Refresh
                           If data_prec.Recordset.RecordCount > 0 Then
                              data_prec.Recordset.MoveLast
                              Xcompfecha = Date - 1
                              If data_prec.Recordset("cnv_hasta") = Xcompfecha Then
                                 data_prec.Recordset.Edit
                                 data_prec.Recordset("cnv_hasta") = Date - 2
                                 data_prec.Recordset.Update
                              Else
                                 data_prec.Recordset.Edit
                                 data_prec.Recordset("cnv_hasta") = Date - 1
                                 data_prec.Recordset.Update
                              End If
                              data_prec.Recordset.AddNew
                              data_prec.Recordset("cnv_codigo") = data_act.Recordset("cnv_codigo")
                              data_prec.Recordset("cnv_desde") = CDate(mdesde.Text)
                              Mifec = Date + 365
                              data_prec.Recordset("cnv_hasta") = Mifec
                              data_prec.Recordset("precio") = Round(data_act.Recordset("cnv_precio"))
                              data_prec.Recordset("moneda") = 1
                              data_prec.Recordset.Update
                           Else
                              data_prec.Recordset.AddNew
                              data_prec.Recordset("cnv_codigo") = data_act.Recordset("cnv_codigo")
                              data_prec.Recordset("cnv_desde") = CDate(mdesde.Text)
                              Mifec = Date + 365
                              data_prec.Recordset("cnv_hasta") = Mifec
                              data_prec.Recordset("precio") = Val(data_act.Recordset("cnv_precio"))
                              data_prec.Recordset("moneda") = 1
                              data_prec.Recordset.Update
                           End If
                        End If
                     End If
                     data_act.Recordset.MoveNext
                  Else
                     data_act.Recordset.MoveNext
                  End If
               Else
                  data_act.Recordset.MoveNext
               End If
            Else
               data_act.Recordset.MoveNext
            End If
         Else
            If Check4.Value = 1 Then
               If IsNull(data_act.Recordset("cnv_preccons")) = False Then
                  If data_act.Recordset("cnv_preccons") > 0 Then
                     data_act.Recordset.Edit
                     Mivalor = data_act.Recordset("cnv_preccons") * mpor.Text
                     data_act.Recordset("cnv_preccons") = Round(Mivalor)
                     data_act.Recordset.Update
                  End If
               End If
            Else
                If IsNull(data_act.Recordset("cnv_precio")) = False Then
                   If data_act.Recordset("cnv_precio") <> 0 Then
                      data_act.Recordset.Edit
                      Mivalor = data_act.Recordset("cnv_precio") * mpor.Text
                      data_act.Recordset("cnv_precio") = Round(Mivalor)
                      data_act.Recordset.Update
                      data_prec.RecordSource = "Select * from cnv_prec where cnv_codigo ='" & data_act.Recordset("cnv_codigo") & "' order by cnv_codigo,cnv_desde,cnv_hasta"
                      data_prec.Refresh
                      If data_prec.Recordset.RecordCount > 0 Then
                         data_prec.Recordset.MoveLast
                         data_prec.Recordset.Edit
                         data_prec.Recordset("cnv_hasta") = Date - 1
                         data_prec.Recordset.Update
                         data_prec.Recordset.AddNew
                         data_prec.Recordset("cnv_codigo") = data_act.Recordset("cnv_codigo")
                         data_prec.Recordset("cnv_desde") = CDate(mdesde.Text)
                         Mifec = Date + 365
                         data_prec.Recordset("cnv_hasta") = Mifec
                         data_prec.Recordset("precio") = Round(data_act.Recordset("cnv_precio"))
                         data_prec.Recordset("moneda") = 1
                         data_prec.Recordset.Update
                      Else
                         data_prec.Recordset.AddNew
                         data_prec.Recordset("cnv_codigo") = data_act.Recordset("cnv_codigo")
                         data_prec.Recordset("cnv_desde") = CDate(mdesde.Text)
                         Mifec = Date + 365
                         data_prec.Recordset("cnv_hasta") = Mifec
                         data_prec.Recordset("precio") = Val(data_act.Recordset("cnv_precio"))
                         data_prec.Recordset("moneda") = 1
                         data_prec.Recordset.Update
                      End If
                   End If
                End If
            End If
            data_act.Recordset.MoveNext
         End If
      Loop
      frm_aumentos.MousePointer = 0
      MsgBox "Proceso terminado", vbInformation, "Mensaje"
   End If
End If
If Option2.Value = True Then
   b_acep.Enabled = False
   XXR = MsgBox("Se cambiaran los precios de Servicios, Desea Continuar?", vbExclamation + vbOKCancel, "Mensaje")
   If XXR = vbOK Then
      frm_aumentos.MousePointer = 11
      
      data_act.RecordSource = "estudios"
      data_act.Refresh
      data_ant.RecordSource = "estudant"
      data_ant.Refresh
      If data_ant.Recordset.RecordCount > 0 Then
         data_ant.Recordset.MoveFirst
         Do While Not data_ant.Recordset.EOF
            data_ant.Recordset.Delete
            data_ant.Recordset.MoveNext
         Loop
      End If
      data_act.Recordset.MoveFirst
      Do While Not data_act.Recordset.EOF
         data_ant.Recordset.AddNew
         data_ant.Recordset("codest") = data_act.Recordset("codest")
         data_ant.Recordset("descrip") = data_act.Recordset("descrip")
         data_ant.Recordset("cons") = data_act.Recordset("cons")
         data_ant.Recordset("uc") = data_act.Recordset("uc")
         data_ant.Recordset("ucfh") = data_act.Recordset("ucfh")
         data_ant.Recordset("part") = data_act.Recordset("part")
         data_ant.Recordset.Update
         data_act.Recordset.MoveNext
      Loop
      frm_aumentos.MousePointer = 0
      MsgBox "Se guardó un respaldo del archivo anterior de estudios en ESTYCONVANT.MDB.", vbInformation, "Mensaje"
      frm_aumentos.MousePointer = 11
      data_act.Recordset.MoveFirst
      Do While Not data_act.Recordset.EOF
         If data_act.Recordset("codest") >= 20041 And data_act.Recordset("codest") <= 20119 Then
         Else
            If data_act.Recordset("flia") = 13 Or data_act.Recordset("flia") = 6 Then
            Else
               If data_act.Recordset("codest") = 990 Or data_act.Recordset("codest") = 995 Then
               Else
                  If IsNull(data_act.Recordset("cons")) = False Then
                     If data_act.Recordset("cons") <> 0 Then
                        data_act.Recordset.Edit
                        Mivalor = data_act.Recordset("cons") * mpor.Text
                        data_act.Recordset("cons") = Round(Mivalor)
                        data_act.Recordset.Update
                     End If
                  End If
                  If IsNull(data_act.Recordset("uc")) = False Then
                     If data_act.Recordset("uc") <> 0 Then
                        Mivalor = data_act.Recordset("uc") * mpor.Text
                        data_act.Recordset.Edit
                        data_act.Recordset("uc") = Round(Mivalor)
                        data_act.Recordset.Update
                     End If
                  End If
                  If IsNull(data_act.Recordset("part")) = False Then
                     If data_act.Recordset("part") <> 0 Then
                        data_act.Recordset.Edit
                        Mivalor = data_act.Recordset("part") * mpor.Text
                        data_act.Recordset("part") = Round(Mivalor)
                        data_act.Recordset.Update
                     End If
                  End If
                  If IsNull(data_act.Recordset("ucfh")) = False Then
                     If data_act.Recordset("ucfh") <> 0 Then
                        Mivalor = data_act.Recordset("ucfh") * mpor.Text
                        data_act.Recordset.Edit
                        data_act.Recordset("ucfh") = Round(Mivalor)
                        data_act.Recordset.Update
                     End If
                  End If
               End If
            End If
         End If
         data_act.Recordset.MoveNext
      Loop
      If Check3.Value = 1 Then
      Else
         data_respara.RecordSource = "arancel"
         data_respara.Refresh
         data_aran.RecordSource = "Select * from Aran_servicios where prec_serv >" & 0
         data_aran.Refresh
         If data_aran.Recordset.RecordCount > 0 Then
            data_aran.Recordset.MoveFirst
            Do While Not data_aran.Recordset.EOF
               data_respara.Recordset.AddNew
               data_respara.Recordset("ara_famnro") = data_aran.Recordset("id_gpo")
               data_respara.Recordset("ara_cnvcod") = Trim(str(data_aran.Recordset("id_serv")))
               data_respara.Recordset("ara_precio") = data_aran.Recordset("prec_serv")
               data_respara.Recordset.Update
               data_aran.Recordset.MoveNext
            Loop
            data_aran.Recordset.MoveFirst
            Do While Not data_aran.Recordset.EOF
               data_aran.Recordset.Edit
               Mivalor = data_aran.Recordset("ara_precio") * mpor.Text
               data_aran.Recordset("ara_precio") = Round(Mivalor)
               data_aran.Recordset.Update
               data_aran.Recordset.MoveNext
            Loop
         End If
      End If
      frm_aumentos.MousePointer = 0
      MsgBox "Proceso terminado", vbInformation, "Mensaje"
   End If
End If

If Option3.Value = True Then
   b_acep.Enabled = False
   XXR = MsgBox("Se cambiarán precios de Servicios de FLIA. seleccionadas, Desea Continuar?", vbExclamation + vbOKCancel, "Mensaje")
   If XXR = vbOK Then
      frm_aumentos.MousePointer = 11
      data_respara.RecordSource = "arancel"
      data_respara.Refresh
      data_aran.RecordSource = "Select * from Aran_servicios where prec_serv >" & 0
      data_aran.Refresh
      If data_aran.Recordset.RecordCount > 0 Then
         data_aran.Recordset.MoveFirst
         Do While Not data_aran.Recordset.EOF
            data_respara.Recordset.AddNew
            data_respara.Recordset("ara_famnro") = data_aran.Recordset("id_gpo")
            data_respara.Recordset("ara_cnvcod") = Trim(str(data_aran.Recordset("id_serv")))
            data_respara.Recordset("ara_precio") = data_aran.Recordset("prec_serv")
            data_respara.Recordset.Update
            data_aran.Recordset.MoveNext
         Loop
      End If
           
      data_act.RecordSource = "estudios"
      data_act.Refresh
      data_ant.RecordSource = "estudant"
      data_ant.Refresh
      If data_ant.Recordset.RecordCount > 0 Then
         data_ant.Recordset.MoveFirst
         Do While Not data_ant.Recordset.EOF
            data_ant.Recordset.Delete
            data_ant.Recordset.MoveNext
         Loop
      End If
      data_act.Recordset.MoveFirst
      Do While Not data_act.Recordset.EOF
         data_ant.Recordset.AddNew
         data_ant.Recordset("codest") = data_act.Recordset("codest")
         data_ant.Recordset("descrip") = data_act.Recordset("descrip")
         data_ant.Recordset("cons") = data_act.Recordset("cons")
         data_ant.Recordset("uc") = data_act.Recordset("uc")
         data_ant.Recordset("ucfh") = data_act.Recordset("ucfh")
         data_ant.Recordset("part") = data_act.Recordset("part")
         data_ant.Recordset.Update
         data_act.Recordset.MoveNext
      Loop
      frm_aumentos.MousePointer = 0
      MsgBox "Se guardó un respaldo del archivo anterior de precios en ESTYCONVANT.MDB", vbInformation, "Mensaje"
      frm_aumentos.MousePointer = 11
      data_act.Recordset.MoveFirst
      Do While Not data_act.Recordset.EOF
         For x = 1 To ListView1.ListItems.count
             ListView1.ListItems(x).Selected = True
             If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
                Xnrofliaa = Val(ListView1.ListItems.Item(x).Text)
                If data_act.Recordset("flia") = Xnrofliaa Then
                   If data_act.Recordset("codest") = 990 Or data_act.Recordset("codest") = 995 Then
                   Else
                      If IsNull(data_act.Recordset("cons")) = False Then
                         If data_act.Recordset("cons") <> 0 Then
                            data_act.Recordset.Edit
                            Mivalor = data_act.Recordset("cons") * mpor.Text
                            data_act.Recordset("cons") = Round(Mivalor)
                            data_act.Recordset.Update
                         End If
                      End If
                      If IsNull(data_act.Recordset("uc")) = False Then
                         If data_act.Recordset("uc") <> 0 Then
                            Mivalor = data_act.Recordset("uc") * mpor.Text
                            data_act.Recordset.Edit
                            data_act.Recordset("uc") = Round(Mivalor)
                            data_act.Recordset.Update
                         End If
                      End If
                      If IsNull(data_act.Recordset("part")) = False Then
                         If data_act.Recordset("part") <> 0 Then
                            data_act.Recordset.Edit
                            Mivalor = data_act.Recordset("part") * mpor.Text
                            data_act.Recordset("part") = Round(Mivalor)
                            data_act.Recordset.Update
                         End If
                      End If
                      If IsNull(data_act.Recordset("ucfh")) = False Then
                         If data_act.Recordset("ucfh") <> 0 Then
                            Mivalor = data_act.Recordset("ucfh") * mpor.Text
                            data_act.Recordset.Edit
                            data_act.Recordset("ucfh") = Round(Mivalor)
                            data_act.Recordset.Update
                         End If
                      End If
                      If Check3.Value = 1 Then
                      Else
                         data_aran.RecordSource = "Select * from Aran_servicios where id_serv =" & data_act.Recordset("codest") & " and prec_serv >" & 0
                         data_aran.Refresh
                         If data_aran.Recordset.RecordCount > 0 Then
                            data_aran.Recordset.MoveFirst
                            Do While Not data_aran.Recordset.EOF
                               If IsNull(data_aran.Recordset("prec_serv")) = False Then
                                   Mivalor = data_aran.Recordset("prec_serv") * mpor.Text
                                   If data_aran.Recordset("prec_serv") <> Round(Mivalor) Then
                                      data_aran.Recordset.Edit
                                      data_aran.Recordset("prec_serv") = Round(Mivalor)
                                      data_aran.Recordset.Update
                                   End If
                               End If
                               data_aran.Recordset.MoveNext
                            Loop
                         End If
                      End If
                   End If
                End If
             End If
         Next x
         x = 0
         data_act.Recordset.MoveNext
      Loop
      frm_aumentos.MousePointer = 0
      MsgBox "Proceso terminado", vbInformation, "Mensaje"
   End If
End If

Exit Sub

Quepasaaumento:
               If Err.Number = 53 Then
                  frm_aumentos.MousePointer = 0
                  MsgBox "ERROR:" & Err.Description & " " & Err.Number
               Else
                  frm_aumentos.MousePointer = 0
                  MsgBox "ERROR!!" & Err.Description
               End If

End Sub

Private Sub b_can_Click()
Unload Me

End Sub

Private Sub Command1_Click()
Dim mhasta As Date
''''Cargo por primera vez el historial
mhasta = CDate(mdesde.Text) + 365
adohist.RecordSource = "Select * from hist_conve"
adohist.Refresh

Command1.Enabled = False
data_act.RecordSource = "Select * from convenio where cnv_precio >" & 0 & " and cnv_emite ='" & "SI" & "' and cnv_hasta >=#" & Format(Date, "yyyy/mm/dd") & "#"
data_act.Refresh
If data_act.Recordset.RecordCount > 0 Then
   data_act.Recordset.MoveFirst
   Do While Not data_act.Recordset.EOF
      If IsNull(data_act.Recordset("cnv_colrec")) = False Then
         If data_act.Recordset("cnv_colrec") = "M" Or data_act.Recordset("cnv_colrec") = "R" Or data_act.Recordset("cnv_colrec") = "A" Or data_act.Recordset("cnv_colrec") = "V" Then
            adohist.Recordset.AddNew
            adohist.Recordset("h_codconv") = data_act.Recordset("cnv_codigo")
            adohist.Recordset("h_importe") = data_act.Recordset("cnv_precio")
            adohist.Recordset("h_fecha1") = Format(mdesde.Text, "yyyy-mm-dd")
            adohist.Recordset("h_fecha2") = Format(mhasta, "yyyy-mm-dd")
            adohist.Recordset("fecha") = Format(Date, "yyyy-mm-dd")
            adohist.Recordset("hora") = Format(Time, "HH:mm")
            adohist.Recordset("usuario") = WElusuario
            adohist.Recordset.Update
         End If
      End If
      data_act.Recordset.MoveNext
   Loop
   MsgBox "Terminado"
Else
   MsgBox "No hay registros"
End If
Command1.Enabled = True

End Sub


Private Sub Form_Load()
Dim x As Integer
Dim Xcount As Long
Xcount = 1
data_act.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_ant.DatabaseName = App.path & "\estyconvant.mdb"
data_prec.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_respara.DatabaseName = App.path & "\arancel.mdb"
data_aran.Connect = "odbc;dsn=" & Xconexrmt & ";"
adohist.ConnectionString = "dsn=" & Xconexrmt
mdesde.Text = Format(Date, "dd/mm/yyyy")
data_fam.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_fam.RecordSource = "Select * from familias order by fam_numero"
data_fam.Refresh

If data_fam.Recordset.RecordCount > 0 Then
   ListView1.ListItems.Clear
   data_fam.Recordset.MoveFirst
   Do While Not data_fam.Recordset.EOF
      If IsNull(data_fam.Recordset("fam_numero")) = False Then
         ListView1.ListItems.Add Xcount, , data_fam.Recordset("fam_numero")
      Else
         ListView1.ListItems.Add Xcount, , "0"
      End If
      If IsNull(data_fam.Recordset("fam_nombre")) = False Then
         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , data_fam.Recordset("fam_nombre")
      Else
         ListView1.ListItems.Item(Xcount).ListSubItems.Add , , " "
      End If
      data_fam.Recordset.MoveNext
      Xcount = Xcount + 1
   Loop
End If
For x = 1 To ListView1.ListItems.count
    ListView1.ListItems(x).Selected = True
    If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
    Else
       ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True
    End If
Next x

End Sub

Private Sub Form_Resize()
With Image1
    .Top = 0
    .Left = 0
    .Width = Me.Width
    .Height = Me.Height
End With

End Sub

Private Sub mpor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_col.SetFocus
End If

End Sub

Private Sub mpor_LostFocus()
If mpor.Text <> "" Then
   mpor.Text = Format(mpor.Text, "#.###")
End If

End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
   Label4.Visible = False
   ListView1.Visible = False
Else

End If

End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
   Label4.Visible = False
   ListView1.Visible = False
Else

End If

End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
   Label4.Visible = True
   ListView1.Visible = True
Else
   Label4.Visible = False
   ListView1.Visible = False

End If
End Sub
