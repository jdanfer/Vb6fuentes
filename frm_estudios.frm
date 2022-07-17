VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{FAEEE763-117E-101B-8933-08002B2F4F5A}#1.1#0"; "DBLIST32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_estudios 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estudios"
   ClientHeight    =   8250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7830
   Icon            =   "frm_estudios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   7830
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_servis 
      Caption         =   "data_servis"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7920
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data data_gpos 
      Caption         =   "data_gpos"
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
      Top             =   7560
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7080
      Picture         =   "frm_estudios.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Borra los datos en la tabla de respaldos y actualiza."
      Top             =   7680
      Width           =   615
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5400
      Visible         =   0   'False
      Width           =   2655
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   6960
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_buscaest 
      Caption         =   "data_buscaest"
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
      Top             =   5160
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Data data_flia 
      Caption         =   "data_flia"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5280
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4440
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data data_est 
      Caption         =   "data_est"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_estudios.frx":0884
      Height          =   2055
      Left            =   120
      OleObjectBlob   =   "frm_estudios.frx":08A0
      TabIndex        =   18
      Top             =   5640
      Width           =   7575
   End
   Begin VB.TextBox txt_busca 
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
      Left            =   2520
      TabIndex        =   17
      Top             =   5280
      Width           =   3495
   End
   Begin VB.CommandButton bimp 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4200
      Picture         =   "frm_estudios.frx":1287
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Informes"
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton bbusca 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3480
      Picture         =   "frm_estudios.frx":1811
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Buscar datos"
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton bcance 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      Picture         =   "frm_estudios.frx":1D9B
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Cancelar acciòn"
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton bmodif 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1800
      Picture         =   "frm_estudios.frx":2325
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Editar registro"
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton bgraba 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   960
      Picture         =   "frm_estudios.frx":28AF
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Grabar registro"
      Top             =   4440
      Width           =   495
   End
   Begin VB.CommandButton bnuevo 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      Picture         =   "frm_estudios.frx":2E39
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Nuevo registro"
      Top             =   4440
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Datos de los servicios"
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
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.CheckBox chrecibo 
         BackColor       =   &H00C00000&
         Caption         =   "Registra como RECIBO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   615
         Left            =   120
         TabIndex        =   25
         Top             =   3360
         Width           =   2175
      End
      Begin VB.CheckBox chsindeuda 
         BackColor       =   &H00C00000&
         Caption         =   "Sin control de deuda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5040
         TabIndex        =   24
         Top             =   360
         Width           =   2415
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1695
         Left            =   2280
         TabIndex        =   23
         Top             =   2520
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   2990
         View            =   3
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
            Text            =   "Código"
            Object.Width           =   1129
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.TextBox txt_part 
         Alignment       =   1  'Right Justify
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
         Left            =   5880
         TabIndex        =   21
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox txt_nroflia 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2280
         TabIndex        =   6
         Top             =   1440
         Width           =   855
      End
      Begin VB.TextBox txt_precdo 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   9
         Top             =   2040
         Width           =   1575
      End
      Begin MSDBCtls.DBCombo dbcboflia 
         Bindings        =   "frm_estudios.frx":33C3
         DataSource      =   "data_flia"
         Height          =   360
         Left            =   3120
         TabIndex        =   7
         Top             =   1440
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   635
         _Version        =   393216
         ListField       =   "FAM_NOMBRE"
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox txt_desc 
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
         TabIndex        =   4
         Top             =   960
         Width           =   5175
      End
      Begin VB.TextBox txt_codest 
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
         TabIndex        =   2
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
         Caption         =   "Grupos:"
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
         Height          =   615
         Left            =   120
         TabIndex        =   22
         Top             =   2520
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C00000&
         Caption         =   "$. PARTICULAR:"
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
         Left            =   4080
         TabIndex        =   20
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "$. Socios:"
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
         TabIndex        =   8
         Top             =   2040
         Width           =   2175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "Familia:"
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
         TabIndex        =   5
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Descripción:"
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
         TabIndex        =   3
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Cód.de Estudio:"
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
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C00000&
      Caption         =   "Buscar por descripción:"
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
      TabIndex        =   16
      Top             =   5280
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   7920
      Y1              =   5040
      Y2              =   5040
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   4560
      Picture         =   "frm_estudios.frx":33DB
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   975
   End
End
Attribute VB_Name = "frm_estudios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bbusca_Click()
DBGrid1.Enabled = False
bnuevo.Enabled = False
bgraba.Enabled = False
bcance.Enabled = False
bmodif.Enabled = False
bimp.Enabled = False
DBGrid1.Enabled = True
TXT_BUSCA.Enabled = True
TXT_BUSCA.SetFocus

End Sub

Private Sub bcance_Click()
If XAcnv = 1 Then
   data_est.Recordset.CancelUpdate
   bnuevo.Enabled = True
   bgraba.Enabled = False
   bcance.Enabled = False
   bmodif.Enabled = True
   bimp.Enabled = True
   bbusca.Enabled = True
   Frame1.Enabled = False
   data_est.Recordset.MoveLast
   ListView1.ListItems.Clear
   iguala
   XAcnv = 0
Else
   bnuevo.Enabled = True
   bgraba.Enabled = False
   bcance.Enabled = False
   bmodif.Enabled = True
   bimp.Enabled = True
   bbusca.Enabled = True
   Frame1.Enabled = False
'   data_est.Recordset.MoveLast
   ListView1.ListItems.Clear
   iguala
   XAcnv = 0
End If
End Sub

Private Sub bgraba_Click()
Dim Xind As Integer
On Error GoTo Nohaydatos

If XAcnv = 1 Then
   If txt_codest.Text <> "" Then
      data_buscaest.Recordset.FindFirst "codest =" & txt_codest.Text
      If data_buscaest.Recordset.NoMatch Then
         data_est.Recordset("codest") = txt_codest.Text
         data_est.Recordset("descrip") = txt_desc.Text
         data_est.Recordset("flia") = txt_nroflia.Text
         data_est.Recordset("nomflia") = dbcboflia.Text
         data_est.Recordset("moneda") = 1
         data_est.Recordset("cons") = txt_precdo.Text
         data_est.Recordset("uc") = txt_precdo.Text
         data_est.Recordset("part") = txt_part.Text
         data_est.Recordset("ucfh") = txt_part.Text
         data_est.Recordset("sin_deuda") = chsindeuda.Value
         data_est.Recordset("es_recibo") = chrecibo.Value
         data_est.Recordset.Update
         For Xind = 1 To ListView1.ListItems.count
             ListView1.ListItems(Xind).Selected = True
             If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
               data_servis.RecordSource = "Select * from Aran_servicios where id_serv =" & txt_codest.Text & " and id_gpo=" & Val(ListView1.ListItems.Item(ListView1.SelectedItem.index).Text)
               data_servis.Refresh
               If data_servis.Recordset.RecordCount > 0 Then
               Else
                  data_servis.Recordset.AddNew
                  data_servis.Recordset("id_gpo") = Val(ListView1.ListItems.Item(ListView1.SelectedItem.index).Text)
                  data_servis.Recordset("id_serv") = txt_codest.Text
                  data_servis.Recordset("desc_serv") = txt_desc.Text
                  data_servis.Recordset("prec_serv") = txt_precdo.Text
                  data_servis.Recordset("por_serv") = 0
                  data_servis.Recordset.Update
               End If
            End If
         Next Xind
         XAcnv = 0
         Frame1.Enabled = False
         bnuevo.Enabled = True
         bgraba.Enabled = False
         bcance.Enabled = False
         bmodif.Enabled = True
         bimp.Enabled = True
         bbusca.Enabled = True
      Else
         MsgBox "Ya existe éste código de estudio,VERIFIQUE!!", vbCritical, "Estudios"
         txt_codest.SetFocus
      End If
   Else
      MsgBox "No ingresó código de estudio", vbCritical, "Estudios"
      txt_codest.SetFocus
   End If
Else
   If txt_codest.Text <> "" Then
      If data_est.Recordset("codest") <> txt_codest.Text Then
         data_est.Recordset.Edit
         data_est.Recordset("codest") = txt_codest.Text
         data_est.Recordset.Update
      End If
      If data_est.Recordset("descrip") <> txt_desc.Text Then
         data_est.Recordset.Edit
         data_est.Recordset("descrip") = txt_desc.Text
         data_est.Recordset.Update
      End If
      If data_est.Recordset("flia") <> txt_nroflia.Text Then
         data_est.Recordset.Edit
         data_est.Recordset("flia") = txt_nroflia.Text
         data_est.Recordset.Update
      End If
      If data_est.Recordset("nomflia") <> dbcboflia.Text Then
         data_est.Recordset.Edit
         data_est.Recordset("nomflia") = dbcboflia.Text
         data_est.Recordset.Update
      End If
      If data_est.Recordset("cons") <> txt_precdo.Text Then
         data_est.Recordset.Edit
         data_est.Recordset("cons") = txt_precdo.Text
         data_est.Recordset.Update
      End If
      If data_est.Recordset("uc") <> txt_precdo.Text Then
         data_est.Recordset.Edit
         data_est.Recordset("uc") = txt_precdo.Text
         data_est.Recordset.Update
      End If
      If data_est.Recordset("part") <> txt_part.Text Then
         data_est.Recordset.Edit
         data_est.Recordset("part") = txt_part.Text
         data_est.Recordset.Update
      End If
      If data_est.Recordset("ucfh") <> txt_part.Text Then
         data_est.Recordset.Edit
         data_est.Recordset("ucfh") = txt_part.Text
         data_est.Recordset.Update
      End If
      If data_est.Recordset("sin_deuda") <> chsindeuda.Value Then
         data_est.Recordset.Edit
         data_est.Recordset("sin_deuda") = chsindeuda.Value
         data_est.Recordset.Update
      End If
      If data_est.Recordset("es_recibo") <> chrecibo.Value Then
         data_est.Recordset.Edit
         data_est.Recordset("es_recibo") = chrecibo.Value
         data_est.Recordset.Update
      End If
      For Xind = 1 To ListView1.ListItems.count
          ListView1.ListItems(Xind).Selected = True
          If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
             data_servis.RecordSource = "Select * from Aran_servicios where id_serv =" & txt_codest.Text & " and id_gpo=" & Val(ListView1.ListItems.Item(ListView1.SelectedItem.index).Text)
             data_servis.Refresh
             If data_servis.Recordset.RecordCount > 0 Then
                If data_servis.Recordset("prec_serv") <> txt_precdo.Text Then
                    data_servis.Recordset.Edit
                    data_servis.Recordset("id_serv") = txt_codest.Text
                    data_servis.Recordset("prec_serv") = txt_precdo.Text
                    data_servis.Recordset.Update
                End If
             Else
                data_servis.Recordset.AddNew
                data_servis.Recordset("id_gpo") = Val(ListView1.ListItems.Item(ListView1.SelectedItem.index).Text)
                data_servis.Recordset("id_serv") = txt_codest.Text
                data_servis.Recordset("desc_serv") = txt_desc.Text
                data_servis.Recordset("prec_serv") = txt_precdo.Text
                data_servis.Recordset("por_serv") = 0
                data_servis.Recordset.Update
             End If
          End If
      Next Xind
      XAcnv = 0
      Frame1.Enabled = False
      bnuevo.Enabled = True
      bgraba.Enabled = False
      bcance.Enabled = False
      bmodif.Enabled = True
      bimp.Enabled = True
      bbusca.Enabled = True
   Else
      MsgBox "Verifique el código de estudio", vbCritical, "Estudios"
      txt_codest.SetFocus
   End If
End If

Exit Sub

Nohaydatos:
            If Err.Number = 3150 Then
               MsgBox "No hay datos para grabar.", vbInformation
            Else
               MsgBox "No hay datos para grabar. Verifique!", vbInformation
            End If
            
End Sub

Private Sub bimp_Click()
Dim Xqueestimp As String
Xqueestimp = MsgBox("Desea imprimir TODOS los servicios?", vbExclamation + vbYesNo, "Opciones de impresión")
frm_estudios.MousePointer = 11
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infcli"

data_inf.RecordSource = "infcli"
data_inf.Refresh

If Xqueestimp = vbYes Then
   data_est.Recordset.MoveFirst
   DoEvents
   Do While Not data_est.Recordset.EOF
      If data_est.Recordset("codest") <> 991 Then
         If data_est.Recordset("codest") >= 13005 And data_est.Recordset("codest") <= 13008 Then
         Else
            If data_est.Recordset("codest") = 13020 Or data_est.Recordset("codest") = 13030 Or data_est.Recordset("codest") = 13033 Or _
               data_est.Recordset("codest") = 13031 Or data_est.Recordset("codest") = 13032 Then
            Else
               If data_est.Recordset("codest") >= 15001 And data_est.Recordset("codest") <= 15004 Then
               Else
                  data_inf.Recordset.AddNew
                  data_inf.Recordset("cl_codigo") = data_est.Recordset("codest")
                  data_inf.Recordset("cl_apellid") = data_est.Recordset("descrip")
                  data_inf.Recordset("cl_cedula") = data_est.Recordset("cons")
                  data_inf.Recordset.Update
               End If
            End If
         End If
      End If
      data_est.Recordset.MoveNext
   Loop
   CrystalReport1.ReportTitle = "INFORME DE TODOS LOS ESTUDIOS REGISTRADOS "
   CrystalReport1.Action = 1

Else
   MsgBox "Emita el informe desde la opción de listados de estudios por familia", vbInformation, "Informes"
   
End If
frm_estudios.MousePointer = 0

End Sub

Private Sub bmodif_Click()
XAcnv = 0
data_est.Recordset.FindFirst "codest =" & txt_codest.Text
If Not data_est.Recordset.NoMatch Then
    Frame1.Enabled = True
    bnuevo.Enabled = False
    bgraba.Enabled = True
    bcance.Enabled = True
    bmodif.Enabled = False
    bimp.Enabled = False
    bbusca.Enabled = False
    iguala
    txt_codest.Enabled = True
    txt_desc.SetFocus
'    txt_codest.SetFocus
End If

End Sub

Private Sub bnuevo_Click()
XAcnv = 1
Frame1.Enabled = True
bnuevo.Enabled = False
bgraba.Enabled = True
bcance.Enabled = True
bmodif.Enabled = False
bimp.Enabled = False
bbusca.Enabled = False
data_est.Recordset.AddNew
txt_codest.Text = ""
txt_desc.Text = ""
txt_nroflia.Text = ""
dbcboflia.Text = ""
txt_precdo.Text = 0
txt_part.Text = 0
chsindeuda.Value = 0
chrecibo.Value = 0
If txt_codest.Enabled = False Then
   txt_codest.Enabled = True
End If
Dim Xcountt As Integer
Xcountt = 1
ListView1.ListItems.Clear

txt_codest.SetFocus
data_gpos.RecordSource = "select * from Aran_grupos"
data_gpos.Refresh
If data_gpos.Recordset.RecordCount > 0 Then
   data_gpos.Recordset.MoveFirst
   Do While Not data_gpos.Recordset.EOF
      If IsNull(data_gpos.Recordset("id")) = False Then
         ListView1.ListItems.Add Xcountt, , data_gpos.Recordset("id")
      Else
         ListView1.ListItems.Add Xcountt, , "0"
      End If
      If IsNull(data_gpos.Recordset("desc_gpo")) = False Then
         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_gpos.Recordset("desc_gpo")
      Else
         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
      End If
      ListView1.ListItems(Xcountt).Selected = True
      ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True
      data_gpos.Recordset.MoveNext
      Xcountt = Xcountt + 1
      
   Loop
End If
'Select * from Aran_servicios where id_gpo =" & t_cod.Text
'                ListView1.ListItems(Xind).Selected = True
'                If ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True Then
'                   Xcant = Xcant + 1
'                End If

End Sub

Private Sub cbomon_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   bgraba.SetFocus
End If

End Sub

Private Sub Command1_Click()
If WElusuario = "COMPUTOS" Then
    Data1.DatabaseName = App.path & "\estresp.mdb"
    Data1.RecordSource = "estudios"
    Data1.Refresh
    If Data1.Recordset.RecordCount > 0 Then
       Data1.Recordset.MoveFirst
       Do While Not Data1.Recordset.EOF
          Data1.Recordset.Delete
          Data1.Recordset.MoveNext
       Loop
    End If
    If data_est.Recordset.RecordCount > 0 Then
       data_est.Recordset.MoveFirst
       Do While Not data_est.Recordset.EOF
          Data1.Recordset.AddNew
          Data1.Recordset("codest") = data_est.Recordset("codest")
          Data1.Recordset("descrip") = data_est.Recordset("descrip")
          Data1.Recordset("flia") = data_est.Recordset("flia")
          Data1.Recordset("nomflia") = data_est.Recordset("nomflia")
          Data1.Recordset("moneda") = data_est.Recordset("moneda")
          Data1.Recordset("cons") = data_est.Recordset("cons")
          Data1.Recordset("uc") = data_est.Recordset("uc")
          Data1.Recordset("part") = data_est.Recordset("part")
          Data1.Recordset("ucfh") = data_est.Recordset("ucfh")
          Data1.Recordset.Update
          data_est.Recordset.MoveNext
       Loop
    End If
    MsgBox "Proceso terminado"
Else
    MsgBox "Solo usuario administrador"
End If

End Sub

Private Sub dbcboflia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_precdo.SetFocus
End If

End Sub

Private Sub dbcboflia_LostFocus()
data_flia.Recordset.FindFirst "fam_nombre = '" & dbcboflia.Text & "'"
If Not data_flia.Recordset.NoMatch Then
   txt_nroflia.Text = data_flia.Recordset("fam_numero")
Else
   MsgBox "No se encuentra FAMILIA", vbCritical, "Mensaje"
   txt_precdo.SetFocus
End If

End Sub

Private Sub DBGrid1_DblClick()
   data_est.Recordset.FindFirst "codest =" & data_buscaest.Recordset("codest")
   If Not data_est.Recordset.NoMatch Then
      iguala
      DBGrid1.Enabled = False
      bnuevo.Enabled = True
      bgraba.Enabled = False
      bcance.Enabled = False
      bmodif.Enabled = True
      bimp.Enabled = True
      TXT_BUSCA.Enabled = False
      bmodif.SetFocus
   End If


End Sub

Private Sub DBGrid1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
End If
   
End Sub

Private Sub Form_Initialize()
Dim Xcountt, Xind As Integer
Xcountt = 1
data_est.Recordset.MoveLast
If IsNull(data_est.Recordset("codest")) = False Then
   txt_codest.Text = data_est.Recordset("codest")
Else
   txt_codest.Text = ""
End If
If IsNull(data_est.Recordset("descrip")) = False Then
   txt_desc.Text = data_est.Recordset("descrip")
Else
   txt_codest.Text = ""
End If
If IsNull(data_est.Recordset("flia")) = False Then
   txt_nroflia.Text = data_est.Recordset("flia")
Else
   txt_nroflia.Text = ""
End If
If IsNull(data_est.Recordset("nomflia")) = False Then
   dbcboflia.Text = data_est.Recordset("nomflia")
Else
   dbcboflia.Text = ""
End If
If IsNull(data_est.Recordset("sin_deuda")) = False Then
   chsindeuda.Value = data_est.Recordset("sin_deuda")
Else
   chsindeuda.Value = 0
End If
If IsNull(data_est.Recordset("es_recibo")) = False Then
   chrecibo.Value = data_est.Recordset("es_recibo")
Else
   chrecibo.Value = 0
End If
If IsNull(data_est.Recordset("cons")) = False Then
   txt_precdo.Text = data_est.Recordset("cons")
Else
   txt_precdo.Text = ""
End If
If IsNull(data_est.Recordset("part")) = False Then
   txt_part.Text = data_est.Recordset("part")
Else
   txt_part.Text = 0
End If
   
data_gpos.RecordSource = "select * from Aran_grupos"
data_gpos.Refresh
If data_gpos.Recordset.RecordCount > 0 Then
   data_gpos.Recordset.MoveFirst
   Do While Not data_gpos.Recordset.EOF
      If IsNull(data_gpos.Recordset("id")) = False Then
         ListView1.ListItems.Add Xcountt, , data_gpos.Recordset("id")
      Else
         ListView1.ListItems.Add Xcountt, , "0"
      End If
      If IsNull(data_gpos.Recordset("desc_gpo")) = False Then
         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_gpos.Recordset("desc_gpo")
      Else
         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
      End If
'      ListView1.ListItems(Xcountt).Selected = True
'      ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True
      data_gpos.Recordset.MoveNext
      Xcountt = Xcountt + 1
      
   Loop
End If
If txt_codest.Text <> "" Then
   For Xind = 1 To ListView1.ListItems.count
       ListView1.ListItems(Xind).Selected = True
       data_servis.RecordSource = "Select * from Aran_servicios where id_serv =" & txt_codest.Text & " and id_gpo=" & Val(ListView1.ListItems.Item(ListView1.SelectedItem.index).Text)
       data_servis.Refresh
       If data_servis.Recordset.RecordCount > 0 Then
          ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True
       Else
          ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = False
       End If
    Next Xind
End If

'Select * from Aran_servicios where id_gpo =" & t_cod.Text

End Sub

Private Sub Form_Load()
data_gpos.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_servis.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_est.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_est.RecordSource = "estudios"
data_est.Refresh
data_buscaest.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_buscaest.RecordSource = "estudios"
data_buscaest.Refresh
data_flia.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_flia.RecordSource = "familias"
data_flia.Refresh
CrystalReport1.ReportFileName = App.path & "\estudios.rpt"
data_inf.DatabaseName = App.path & "\informes.mdb"
'data_inf.RecordSource = "infcli"
'data_inf.Refresh

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub txt_busca_Change()
data_buscaest.RecordSource = "select * from estudios where descrip >='" & TXT_BUSCA.Text & "' order by descrip"
data_buscaest.Refresh

End Sub

Private Sub TXT_BUSCA_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   DBGrid1.SetFocus
End If

End Sub

Public Function iguala()
Dim Xcountt, Xind As Integer
Xcountt = 1
Xind = 0
If IsNull(data_est.Recordset("codest")) = False Then
   txt_codest.Text = data_est.Recordset("codest")
Else
   txt_codest.Text = ""
End If
If IsNull(data_est.Recordset("descrip")) = False Then
   txt_desc.Text = data_est.Recordset("descrip")
Else
   txt_codest.Text = ""
End If
If IsNull(data_est.Recordset("flia")) = False Then
   txt_nroflia.Text = data_est.Recordset("flia")
Else
   txt_nroflia.Text = ""
End If
If IsNull(data_est.Recordset("sin_deuda")) = False Then
   chsindeuda.Value = data_est.Recordset("sin_deuda")
Else
   chsindeuda.Value = 0
End If
If IsNull(data_est.Recordset("es_recibo")) = False Then
   chrecibo.Value = data_est.Recordset("es_recibo")
Else
   chrecibo.Value = 0
End If
If IsNull(data_est.Recordset("nomflia")) = False Then
   dbcboflia.Text = data_est.Recordset("nomflia")
Else
   dbcboflia.Text = ""
End If
If IsNull(data_est.Recordset("cons")) = False Then
   txt_precdo.Text = data_est.Recordset("cons")
Else
   txt_precdo.Text = ""
End If
If IsNull(data_est.Recordset("part")) = False Then
   txt_part.Text = data_est.Recordset("part")
Else
   txt_part.Text = 0
End If
ListView1.ListItems.Clear

data_gpos.RecordSource = "select * from Aran_grupos"
data_gpos.Refresh
If data_gpos.Recordset.RecordCount > 0 Then
   data_gpos.Recordset.MoveFirst
   Do While Not data_gpos.Recordset.EOF
      If IsNull(data_gpos.Recordset("id")) = False Then
         ListView1.ListItems.Add Xcountt, , data_gpos.Recordset("id")
      Else
         ListView1.ListItems.Add Xcountt, , "0"
      End If
      If IsNull(data_gpos.Recordset("desc_gpo")) = False Then
         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , data_gpos.Recordset("desc_gpo")
      Else
         ListView1.ListItems.Item(Xcountt).ListSubItems.Add , , ""
      End If
'      ListView1.ListItems(Xcountt).Selected = True
'      ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True
      data_gpos.Recordset.MoveNext
      Xcountt = Xcountt + 1
      
   Loop
End If
If txt_codest.Text <> "" Then
   For Xind = 1 To ListView1.ListItems.count
       ListView1.ListItems(Xind).Selected = True
       data_servis.RecordSource = "Select * from Aran_servicios where id_serv =" & txt_codest.Text & " and id_gpo=" & Val(ListView1.ListItems.Item(ListView1.SelectedItem.index).Text)
       data_servis.Refresh
       If data_servis.Recordset.RecordCount > 0 Then
          ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = True
       Else
          ListView1.ListItems.Item(ListView1.SelectedItem.index).Checked = False
       End If
    Next Xind
End If


End Function

Private Sub txt_codest_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_desc.SetFocus
End If

End Sub

Private Sub txt_desc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_nroflia.SetFocus
End If

End Sub

Private Sub txt_nroflia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   dbcboflia.SetFocus
End If

End Sub

Private Sub txt_nroflia_LostFocus()
If txt_nroflia.Text <> "" Then
   data_flia.Recordset.FindFirst "fam_numero =" & txt_nroflia.Text
   If Not data_flia.Recordset.NoMatch Then
      dbcboflia.Text = data_flia.Recordset("fam_nombre")
      txt_precdo.SetFocus
   Else
      dbcboflia.SetFocus
   End If
Else
   dbcboflia.SetFocus
End If

End Sub



Private Sub txt_part_LostFocus()
If txt_part.Text = "" Then
   txt_part.Text = 0
End If

End Sub




Private Sub txt_precdo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_part.SetFocus
End If

End Sub

Private Sub txt_precdo_LostFocus()
If txt_precdo.Text = "" Then
   txt_precdo.Text = 0
End If

End Sub

