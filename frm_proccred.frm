VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_proccred 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Procesar ventas crédito a emisión"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6705
   Icon            =   "frm_proccred.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   6705
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport crr 
      Left            =   3120
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ProgressBar pbb 
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   2280
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Data data_conv 
      Caption         =   "data_conv"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1680
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Data data_cli 
      Caption         =   "data_cli"
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
      Width           =   2415
   End
   Begin VB.Data data_ser 
      Caption         =   "data_ser"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1080
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data data_lin 
      Caption         =   "data_lin"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2040
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      Picture         =   "frm_proccred.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Salir"
      Top             =   2760
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      Picture         =   "frm_proccred.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Procesar"
      Top             =   2760
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Datos a procesar..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6015
      Begin MSMask.MaskEdBox mh 
         Height          =   375
         Left            =   4200
         TabIndex        =   3
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         Left            =   2520
         TabIndex        =   2
         Top             =   480
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
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
         BackColor       =   &H00C0FFC0&
         Caption         =   "Rango de Fechas:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   2640
      Picture         =   "frm_proccred.frx":0F56
      Stretch         =   -1  'True
      Top             =   2160
      Width           =   2535
   End
End
Attribute VB_Name = "frm_proccred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
frm_proccred.MousePointer = 11
Command1.Enabled = False
Command2.Enabled = False
If md.Text <> "__/__/____" Then
   If mh.Text <> "__/__/____" Then
      If data_ser.Recordset.RecordCount > 0 Then
         data_ser.Recordset.MoveFirst
         Do While Not data_ser.Recordset.EOF
            data_ser.Recordset.Delete
            data_ser.Recordset.MoveNext
         Loop
      End If
      data_lin.RecordSource = "select * from linmmdd where tipo ='" & "CREDITO" & "' And fecha >=#" & Format(md.Text, "yyyy/mm/dd") & "# And fecha <=#" & Format(mh.Text, "yyyy/mm/dd") & "#"
      data_lin.Refresh
      If data_lin.Recordset.RecordCount > 0 Then
         data_lin.Recordset.MoveLast
         pbb.Max = data_lin.Recordset.RecordCount
         data_lin.Recordset.MoveFirst
         Do While Not data_lin.Recordset.EOF
            data_conv.Recordset.FindFirst "cnv_codigo ='" & data_lin.Recordset("convenio") & "'"
            If Not data_conv.Recordset.NoMatch Then
               If IsNull(data_conv.Recordset("cnv_emite")) = False Then
                  If IsNull(data_conv.Recordset("cnv_colrec")) = False Then
                     If data_conv.Recordset("cnv_emite") = "SI" Then
                        If data_conv.Recordset("cnv_colrec") <> "" Then
                           data_cli.RecordSource = "Select * from clientes where cl_codigo =" & data_lin.Recordset("cod_cli")
                           If data_cli.Recordset.RecordCount > 0 Then
                              If IsNull(data_cli.Recordset("estado")) = True Then
                              Else
                                 If data_cli.Recordset("estado") = 2 Or data_cli.Recordset("estado") = 3 Then
                                 Else
                                    data_ser.Recordset.AddNew
                                    data_ser.Recordset("mat") = data_lin.Recordset("cod_cli")
                                    data_ser.Recordset("nombre") = data_lin.Recordset("nom_cli")
                                    data_ser.Recordset("imp") = data_lin.Recordset("tot_lin")
                                    data_ser.Recordset("fecha") = data_lin.Recordset("fecha")
                                    data_ser.Recordset("cob") = data_cli.Recordset("cl_nrocobr")
                                    data_ser.Recordset.Update
                                 End If
                              End If
                           End If
                        End If
                     End If
                  End If
               End If
            End If
            data_lin.Recordset.MoveNext
            pbb.value = pbb.value + 1
         Loop
         frm_proccred.MousePointer = 0
         crr.ReportTitle = "VENTAS CREDITO INCLUIDAS EN EMISION FECHA: " & md.Text & " AL " & mh.Text
         crr.ReportFileName = App.Path & "\infcredemi.rpt"
         crr.Action = 1
      End If
   End If
End If
frm_proccred.MousePointer = 0
Command1.Enabled = False
Command2.Enabled = False
pbb.value = 0
   
End Sub

Private Sub Form_Load()
data_ser.DatabaseName = App.Path & "\env_tiq.mdb"
data_ser.RecordSource = "emiserv"
data_ser.Refresh
data_lin.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_conv.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_conv.RecordSource = "convenio"
data_conv.Refresh
data_cli.Connect = "odbc;dsn=" & Xconexrmt & ";"
'data_cli.RecordSource = "clientes"
'data_cli.Refresh

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub md_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mh.SetFocus
End If

End Sub

Private Sub mh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command1.SetFocus
End If

End Sub
