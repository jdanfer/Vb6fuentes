VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_cobr 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cobradores"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7365
   Icon            =   "frm_cobr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   7365
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   480
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data data_arq 
      Caption         =   "data_arq"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   6240
      Top             =   3720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      Top             =   3600
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_cobr.frx":0442
      Height          =   2415
      Left            =   120
      OleObjectBlob   =   "frm_cobr.frx":0459
      TabIndex        =   14
      Top             =   3840
      Width           =   7095
   End
   Begin VB.TextBox txt_bcob 
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
      Height          =   285
      Left            =   2280
      TabIndex        =   13
      Top             =   3480
      Width           =   3735
   End
   Begin VB.Data data_cob 
      Caption         =   "data_cob"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3240
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.CommandButton bimp 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4320
      Picture         =   "frm_cobr.frx":0E3C
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Informes"
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton bbusca 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3480
      Picture         =   "frm_cobr.frx":13C6
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Buscar"
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton bcance 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      Picture         =   "frm_cobr.frx":1950
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Cancelar acción"
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton bmodif 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1800
      Picture         =   "frm_cobr.frx":1EDA
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Editar"
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton bgraba 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   960
      Picture         =   "frm_cobr.frx":2464
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Guardar datos"
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton bnuevo 
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
      Left            =   120
      Picture         =   "frm_cobr.frx":29EE
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Nuevo registro"
      Top             =   2520
      Width           =   495
   End
   Begin MSMask.MaskEdBox fecing 
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      Enabled         =   0   'False
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
   Begin VB.TextBox txt_nomcob 
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
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   720
      Width           =   4695
   End
   Begin VB.TextBox txt_nrocob 
      Enabled         =   0   'False
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
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label labtotpp 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3720
      TabIndex        =   22
      Top             =   2040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label10 
      BackColor       =   &H00FFFFFF&
      Caption         =   "$."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3240
      TabIndex        =   21
      Top             =   2040
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label labtotrecp 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2160
      TabIndex        =   20
      Top             =   2040
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pendiente:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   19
      Top             =   2040
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label labtotpc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3720
      TabIndex        =   18
      Top             =   1560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "$."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   3240
      TabIndex        =   17
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label labtotrecc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2160
      TabIndex        =   16
      Top             =   1560
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total Cobrado:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   240
      TabIndex        =   15
      Top             =   1560
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C00000&
      Caption         =   "Nombre a buscar:"
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
      Left            =   120
      TabIndex        =   12
      Top             =   3480
      Width           =   2175
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   7320
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   7320
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FECHA INGRESO:"
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
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nombre:"
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
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nro.Cobrador:"
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
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   2910
      Left            =   4200
      Picture         =   "frm_cobr.frx":2F78
      Top             =   1080
      Width           =   3885
   End
End
Attribute VB_Name = "frm_cobr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bbusca_Click()
txt_bcob.Enabled = True
txt_bcob.SetFocus

End Sub

Private Sub bcance_Click()
If XAcnv = 1 Then
   data_cob.Recordset.CancelUpdate
   igualcob
   XAcnv = 0
   desh
Else
   igualcob
   XAcnv = 0
   desh
End If
bgraba.Enabled = False
bcance.Enabled = False
bmodif.Enabled = True
bbusca.Enabled = True
bimp.Enabled = True
bnuevo.Enabled = True

End Sub

Private Sub bgraba_Click()
If txt_nrocob.Text <> "" Then
   If txt_nrocob.Text <> 0 Then
         If XAcnv = 1 Then
            data_cob.Recordset("cb_numero") = txt_nrocob.Text
            data_cob.Recordset("cb_nombre") = txt_nomcob.Text
            If fecing.Text <> "__/__/____" Then
               data_cob.Recordset("cb_fch_ing") = Format(fecing.Text, "dd/mm/yyyy")
            End If
            data_cob.Recordset("cb_recatra") = 0
            data_cob.Recordset.Update
            Data1.Refresh
            XAcnv = 0
            bgraba.Enabled = False
            bcance.Enabled = False
            bmodif.Enabled = True
            bbusca.Enabled = True
            bimp.Enabled = True
            bnuevo.Enabled = True
            desh
         Else
            data_cob.Recordset.Edit
            data_cob.Recordset("cb_numero") = txt_nrocob.Text
            data_cob.Recordset("cb_nombre") = txt_nomcob.Text
            If fecing.Text <> "__/__/____" Then
               data_cob.Recordset("cb_fch_ing") = Format(fecing.Text, "dd/mm/yyyy")
            End If
            data_cob.Recordset.Update
            XAcnv = 0
            Data1.Refresh
            bgraba.Enabled = False
            bcance.Enabled = False
            bmodif.Enabled = True
            bbusca.Enabled = True
            bimp.Enabled = True
            bnuevo.Enabled = True
            desh
         End If
   Else
      MsgBox "No ingresó cobrador", vbCritical, "Cobradores"
      txt_nrocob.SetFocus
   End If
Else
   MsgBox "No ingresó cobrador", vbCritical, "Cobradores"
   txt_nrocob.SetFocus
End If

End Sub

Private Sub bimp_Click()
    Dim MiBaseact As Database
    Dim Unasesact As Workspace
    Set Unasesact = Workspaces(0)
    Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")
    
    MiBaseact.Execute "Delete * from infvtas"
   
   Data2.RecordSource = "infvtas"
   Data2.Refresh
   
   If data_cob.Recordset.RecordCount > 0 Then
      data_cob.Recordset.MoveFirst
      Do While Not data_cob.Recordset.EOF
         Data2.Recordset.AddNew
         Data2.Recordset("cod_cli") = data_cob.Recordset("cb_numero")
         Data2.Recordset("nom_cli") = data_cob.Recordset("cb_nombre")
         Data2.Recordset.Update
         data_cob.Recordset.MoveNext
      Loop
   End If
   Data2.RecordSource = "select * from infvtas order by cod_cli"
   Data2.Refresh
   
CrystalReport1.ReportFileName = App.path & "\cobradores.rpt"
CrystalReport1.Action = 1

End Sub

Private Sub bmodif_Click()
XAcnv = 0
hab
txt_nrocob.SetFocus
bgraba.Enabled = True
bcance.Enabled = True
bmodif.Enabled = False
bbusca.Enabled = False
bimp.Enabled = False
bnuevo.Enabled = False
txt_bcob.Enabled = False
DBGrid1.Enabled = False

End Sub

Private Sub bnuevo_Click()
XAcnv = 1
hab
txt_nrocob.Text = ""
txt_nomcob.Text = ""
fecing.Text = "__/__/____"
txt_nrocob.SetFocus
bgraba.Enabled = True
bcance.Enabled = True
bmodif.Enabled = False
bbusca.Enabled = False
bimp.Enabled = False
bnuevo.Enabled = False
data_cob.Recordset.AddNew

End Sub

Private Sub DBGrid1_DblClick()
frm_cobr.MousePointer = 11
If IsNull(data_cob.Recordset("cb_numero")) = False Then
   txt_nrocob.Text = data_cob.Recordset("cb_numero")
Else
   txt_nrocob.Text = ""
End If
If IsNull(data_cob.Recordset("cb_nombre")) = False Then
   txt_nomcob.Text = data_cob.Recordset("cb_nombre")
Else
   txt_nomcob.Text = ""
End If
If IsNull(data_cob.Recordset("cb_fch_ing")) = False Then
   fecing.Text = Format(data_cob.Recordset("cb_fch_ing"), "dd/mm/yyyy")
Else
   fecing.Text = "__/__/____"
End If

'If IsNull(data_cob.Recordset("cb_numero")) = False Then
'   Dim Xtotporcob As Double
'   Xtotporcob = 0
'   data_arq.RecordSource = "Select * from arqueo where cob =" & data_cob.Recordset("cb_numero") & " and arqueo ='" & "C" & "'"
'   data_arq.Refresh
'   If data_arq.Recordset.RecordCount > 0 Then
'      data_arq.Recordset.MoveFirst
 '     Do While Not data_arq.Recordset.EOF
'         Xtotporcob = Xtotporcob + data_arq.Recordset("total")
'         data_arq.Recordset.MoveNext
'      Loop
'      labtotrecc.Caption = data_arq.Recordset.RecordCount
'      labtotpc.Caption = Format(Xtotporcob, "Standard")
'   Else
'      labtotrecc.Caption = 0
'      labtotpc.Caption = 0
'   End If
'   Xtotporcob = 0
'   data_arq.RecordSource = "Select * from arqueo where cob =" & data_cob.Recordset("cb_numero") & " and arqueo in ('P','E')"
'   data_arq.Refresh
'   If data_arq.Recordset.RecordCount > 0 Then
'      data_arq.Recordset.MoveFirst
'      Do While Not data_arq.Recordset.EOF
'         Xtotporcob = Xtotporcob + data_arq.Recordset("total")
'         data_arq.Recordset.MoveNext
'      Loop
'      labtotrecc.Caption = data_arq.Recordset.RecordCount
'      labtotpc.Caption = Format(Xtotporcob, "Standard")
'   Else
'      labtotrecp.Caption = 0
'      labtotpp.Caption = 0
'   End If
'Else
'   labtotrecc.Caption = 0
 '  labtotpc.Caption = 0
'   labtotrecp.Caption = 0
'   labtotpp.Caption = 0
'End If
frm_cobr.MousePointer = 0

End Sub

Private Sub fecing_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   bgraba.SetFocus
End If

End Sub

Private Sub Form_Initialize()
data_cob.Recordset.MoveLast
If IsNull(data_cob.Recordset("cb_numero")) = False Then
   txt_nrocob.Text = data_cob.Recordset("cb_numero")
Else
   txt_nrocob.Text = ""
End If
If IsNull(data_cob.Recordset("cb_nombre")) = False Then
   txt_nomcob.Text = data_cob.Recordset("cb_nombre")
Else
   txt_nomcob.Text = ""
End If
If IsNull(data_cob.Recordset("cb_fch_ing")) = False Then
   fecing.Text = Format(data_cob.Recordset("cb_fch_ing"), "dd/mm/yyyy")
Else
   fecing.Text = "__/__/____"
End If
data_arq.Connect = "odbc;dsn=" & Xconexrmt & ";"
'If IsNull(data_cob.Recordset("cb_numero")) = False Then
'   Dim Xtotporcob As Double
'   Xtotporcob = 0
'   data_arq.RecordSource = "Select * from arqueo where cob =" & data_cob.Recordset("cb_numero") & " and arqueo ='" & "C" & "'"
'   data_arq.Refresh
'   If data_arq.Recordset.RecordCount > 0 Then
'      data_arq.Recordset.MoveFirst
'      Do While Not data_arq.Recordset.EOF
'         Xtotporcob = Xtotporcob + data_arq.Recordset("total")
'         data_arq.Recordset.MoveNext
'      Loop
'      labtotrecc.Caption = data_arq.Recordset.RecordCount
'      labtotpc.Caption = Format(Xtotporcob, "Standard")
'   Else
'      labtotrecc.Caption = 0
'      labtotpc.Caption = 0
'   End If
'   Xtotporcob = 0
'   data_arq.RecordSource = "Select * from arqueo where cob =" & data_cob.Recordset("cb_numero") & " and arqueo in ('P','E')"
'   data_arq.Refresh
'   If data_arq.Recordset.RecordCount > 0 Then
'      data_arq.Recordset.MoveFirst
'      Do While Not data_arq.Recordset.EOF
'         Xtotporcob = Xtotporcob + data_arq.Recordset("total")
'         data_arq.Recordset.MoveNext
'      Loop
'      labtotrecc.Caption = data_arq.Recordset.RecordCount
'      labtotpc.Caption = Format(Xtotporcob, "Standard")
'   Else
'      labtotrecp.Caption = 0
'      labtotpp.Caption = 0
'   End If
'Else
'   labtotrecc.Caption = 0
'   labtotpc.Caption = 0
'   labtotrecp.Caption = 0
'   labtotpp.Caption = 0
'End If
Data2.DatabaseName = App.path & "\informes.mdb"

End Sub

Public Function hab()
txt_nrocob.Enabled = True
txt_nomcob.Enabled = True
fecing.Enabled = True

End Function

Public Function desh()
txt_nrocob.Enabled = False
txt_nomcob.Enabled = False
fecing.Enabled = False

End Function

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.RecordSource = "Select * from cobrador where cb_recatra <>" & 2
Data1.Refresh
data_cob.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_cob.RecordSource = "Select * from cobrador where cb_recatra <>" & 2
data_cob.Refresh

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub txt_bcob_Change()
data_cob.RecordSource = "select * from cobrador where cb_nombre >='" & txt_bcob.Text & "' and cb_recatra <>" & 2 & " order by cb_nombre"
data_cob.Refresh

End Sub

Private Sub txt_bcob_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   KeyAscii = 0
   DBGrid1.SetFocus
End If

End Sub

Public Function igualcob()
data_cob.Recordset.MoveLast
If IsNull(data_cob.Recordset("cb_numero")) = False Then
   txt_nrocob.Text = data_cob.Recordset("cb_numero")
Else
   txt_nrocob.Text = ""
End If
If IsNull(data_cob.Recordset("cb_nombre")) = False Then
   txt_nomcob.Text = data_cob.Recordset("cb_nombre")
Else
   txt_nomcob.Text = ""
End If
If IsNull(data_cob.Recordset("cb_fch_ing")) = False Then
   fecing.Text = Format(data_cob.Recordset("cb_fch_ing"), "dd/mm/yyyy")
Else
   fecing.Text = "__/__/____"
End If

End Function

Private Sub txt_nomcob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   fecing.SetFocus
End If

End Sub

Private Sub txt_nrocob_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_nomcob.SetFocus
End If

End Sub

Private Sub txt_nrocob_LostFocus()
If XAcnv = 1 Then
   Data1.Recordset.FindFirst "cb_numero =" & txt_nrocob.Text & " and cb_recatra <>" & 2
   If Not Data1.Recordset.NoMatch Then
      MsgBox "Ya existe este número de cobrador", vbCritical, "Cobrador"
   End If
End If

End Sub
