VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_accat 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acciones registradas"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   11835
   Icon            =   "frm_accat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   11835
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   8160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3480
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4560
      TabIndex        =   16
      Top             =   3000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   10800
      TabIndex        =   15
      Top             =   5880
      Width           =   975
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   6000
      TabIndex        =   14
      Top             =   3000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   8760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2880
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   3000
      Picture         =   "frm_accat.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2040
      Picture         =   "frm_accat.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1080
      Picture         =   "frm_accat.frx":0F56
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      Picture         =   "frm_accat.frx":14E0
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2880
      Width           =   615
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_accat.frx":1A6A
      Height          =   2415
      Left            =   120
      OleObjectBlob   =   "frm_accat.frx":1A7E
      TabIndex        =   1
      Top             =   3480
      Width           =   11655
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Datos de la acción tomada"
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
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      Begin VB.TextBox txt_det 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   2400
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   960
         Width           =   8535
      End
      Begin VB.TextBox txt_us 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8880
         TabIndex        =   7
         Top             =   360
         Width           =   2055
      End
      Begin MSMask.MaskEdBox mhora 
         Height          =   375
         Left            =   6000
         TabIndex        =   5
         Top             =   360
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:mm:ss"
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfec 
         Height          =   375
         Left            =   2400
         TabIndex        =   3
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
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
      Begin VB.Label Label5 
         BackColor       =   &H00FF0000&
         Caption         =   "Detalle:"
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
         Height          =   855
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FF0000&
         Caption         =   "Usuario:"
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
         Height          =   375
         Left            =   7440
         TabIndex        =   6
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FF0000&
         Caption         =   "Hora:"
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
         Height          =   375
         Left            =   4680
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "Fecha:"
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
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   8640
      Picture         =   "frm_accat.frx":27A9
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1215
   End
End
Attribute VB_Name = "frm_accat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Frame1.Enabled = True
mfec.Text = Date
mhora.Text = Format(Time, "HH:mm:ss")
txt_us.Text = WElusuario
txt_det.SetFocus
XAlta = 1

End Sub

Private Sub Command2_Click()
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = True
Command4.Enabled = True
Frame1.Enabled = True
txt_det.SetFocus

XAlta = 0

End Sub

Private Sub Command3_Click()

If XAlta = 1 Then
   If txt_det.Text <> "" Then
    Text1.Text = Data2.Recordset("nro_accadm") + 1
    Data2.Recordset.Edit
    Data2.Recordset("nro_accadm") = Text1.Text
    Data2.Recordset.Update
    Data2.Refresh
    Data1.Recordset.AddNew
    Data1.Recordset("at_nroseg") = Text3.Text
    Data1.Recordset("at_fecseg") = mfec.Text
    Data1.Recordset("at_horseg") = mhora.Text
    Data1.Recordset("at_ususeg") = txt_us.Text
    Data1.Recordset("at_detseg") = txt_det.Text
    Data1.Recordset("at_nronro") = Text1.Text
    Data1.Recordset.Update
    Data1.Refresh
    XAlta = 0
    Text1.Text = Text1.Text + 1
  Else
    MsgBox "Debe ingresar detalles.", vbInformation
  End If
Else
   Data1.Recordset.Edit
   Data1.Recordset("at_nroseg") = Text3.Text
   Data1.Recordset("at_fecseg") = mfec.Text
   Data1.Recordset("at_horseg") = mhora.Text
   Data1.Recordset("at_ususeg") = txt_us.Text
   Data1.Recordset("at_detseg") = txt_det.Text
   Data1.Recordset.Update
   Data1.Refresh
   XAlta = 0
End If
Data1.RecordSource = "Select * from seguirat where at_nroseg =" & Text3.Text & " order by at_nronro"
Data1.Refresh
mfec.Text = "__/__/____"
mhora.Text = "__:__:__"
txt_us.Text = ""
txt_det.Text = ""
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Frame1.Enabled = False

   
End Sub

Private Sub Command4_Click()
mfec.Text = "__/__/____"
mhora.Text = "__:__:__"
txt_us.Text = ""
txt_det.Text = ""
Command1.Enabled = True
Command2.Enabled = True
Command3.Enabled = False
Command4.Enabled = False
Frame1.Enabled = False

End Sub

Private Sub Command5_Click()
Unload Me

End Sub

Private Sub DBGrid1_DblClick()
mfec.Text = Data1.Recordset("at_fecseg")
mhora.Text = Data1.Recordset("at_horseg")
txt_us.Text = Data1.Recordset("at_ususeg")
If IsNull(Data1.Recordset("at_detseg")) = False Then
   txt_det.Text = Data1.Recordset("at_detseg")
Else
   txt_det.Text = ""
End If

End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"

Data2.DatabaseName = App.path & "\paramb.mdb"
Data2.RecordSource = "paramb"
Data2.Refresh

'Data1.Connect = "ODBC;DSN=sappat;"
Data1.RecordSource = "Select * from seguirat where at_nroseg =" & frm_atsocio.txt_nro.Text & " order by at_nronro"
Data1.Refresh
Text3.Text = frm_atsocio.txt_nro.Text

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub
