VERSION 5.00
Begin VB.Form frm_espec 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Horarios especialistas"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7140
   Icon            =   "frm_espec.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   7140
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_cance 
      Caption         =   "data_cance"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5040
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data data_lis 
      Caption         =   "data_lis"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton b_rec 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Borrar Datos..."
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
      Height          =   615
      Left            =   4920
      MouseIcon       =   "frm_espec.frx":0442
      MousePointer    =   99  'Custom
      Picture         =   "frm_espec.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Borra listas y fechas ingresadas del especialista seleccionado"
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Data data_buscod 
      Caption         =   "data_buscod"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5160
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.CommandButton b_bustodo 
      BackColor       =   &H00C0C0FF&
      Caption         =   "BUSCAR..."
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
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4800
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Datos de Horarios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   21
      Top             =   5400
      Width           =   6855
      Begin VB.CommandButton b_imp 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Imprimir"
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
         Left            =   4920
         Picture         =   "frm_espec.frx":0B8E
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton b_ing 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Ingresar Fechas"
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
         Left            =   2520
         Picture         =   "frm_espec.frx":0FD0
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton b_cons 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Consulta Fechas"
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
         Left            =   120
         Picture         =   "frm_espec.frx":1412
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Datos de Especialista"
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
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
      Begin VB.Data data_medicos 
         Caption         =   "data_medicos"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   4080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   720
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox t_cmed 
         Height          =   285
         Left            =   240
         TabIndex        =   34
         Top             =   1680
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ComboBox cbomed 
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
         Height          =   360
         Left            =   1440
         TabIndex        =   33
         Top             =   1440
         Width           =   5295
      End
      Begin VB.TextBox t_mh 
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
         Left            =   2640
         TabIndex        =   31
         Top             =   2760
         Width           =   495
      End
      Begin VB.TextBox t_hh 
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
         TabIndex        =   30
         Top             =   2760
         Width           =   495
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "C/DOS"
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
         Height          =   255
         Left            =   5160
         TabIndex        =   28
         Top             =   2520
         Width           =   1215
      End
      Begin VB.CommandButton b_alta 
         Height          =   615
         Left            =   240
         Picture         =   "frm_espec.frx":1B54
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3960
         Width           =   735
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   300
         Left            =   2880
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "PARSEC0"
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Data data_espec 
         Caption         =   "data_espec"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   420
         Left            =   -1080
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   4080
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Data data_med 
         Caption         =   "data_med"
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
         RecordSource    =   "medicos"
         Top             =   3600
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.CommandButton b_busca 
         Height          =   615
         Left            =   4560
         Picture         =   "frm_espec.frx":1F96
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Eliminar registro"
         Top             =   3960
         Width           =   735
      End
      Begin VB.CommandButton b_canc 
         Enabled         =   0   'False
         Height          =   615
         Left            =   3480
         Picture         =   "frm_espec.frx":23D8
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3960
         Width           =   735
      End
      Begin VB.CommandButton b_graba 
         Enabled         =   0   'False
         Height          =   615
         Left            =   2400
         Picture         =   "frm_espec.frx":281A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   3960
         Width           =   735
      End
      Begin VB.CommandButton b_modif 
         Height          =   615
         Left            =   1320
         Picture         =   "frm_espec.frx":2C5C
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   3960
         Width           =   735
      End
      Begin VB.TextBox txt_espera 
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
         Left            =   4680
         TabIndex        =   15
         Top             =   3240
         Width           =   615
      End
      Begin VB.TextBox txt_cantp 
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
         Left            =   2400
         TabIndex        =   13
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox txt_mmpp 
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
         Left            =   4560
         TabIndex        =   11
         Top             =   2280
         Width           =   615
      End
      Begin VB.TextBox txt_mm 
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
         Left            =   2640
         MaxLength       =   2
         TabIndex        =   9
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox txt_hh 
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
         MaxLength       =   2
         TabIndex        =   8
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox txt_desc 
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
         Left            =   2640
         MaxLength       =   50
         TabIndex        =   5
         Top             =   960
         Width           =   4095
      End
      Begin VB.TextBox txt_cod 
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
         Left            =   1440
         MaxLength       =   6
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txt_base 
         Alignment       =   2  'Center
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
         Height          =   375
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0E0FF&
         Caption         =   "MEDICO:"
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
         TabIndex        =   32
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0E0FF&
         Caption         =   "HORA TERMINA:"
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
         TabIndex        =   29
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0E0FF&
         Caption         =   "MINUTOS"
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
         Left            =   5160
         TabIndex        =   26
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   2
         X1              =   0
         X2              =   6840
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0E0FF&
         Caption         =   "ESPERA:"
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
         Left            =   3480
         TabIndex        =   14
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0E0FF&
         Caption         =   "CANT. de PACIENTES:"
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
         TabIndex        =   12
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0E0FF&
         Caption         =   "C/PACIENTE"
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
         Left            =   3360
         TabIndex        =   10
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0E0FF&
         Caption         =   "HORA COMIENZO"
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
         TabIndex        =   7
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "HORARIOS Y CANTIDAD DE PACIENTES:"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   4575
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "CODIGO:"
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
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "BASE:"
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
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm_espec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub b_acep_Click()
   MsgBox "Proceso finalizado", vbInformation, "Mensaje"

End Sub

Private Sub b_alta_Click()
XAlta = 1
txt_base.Enabled = True
txt_cod.Enabled = True
txt_desc.Enabled = True
txt_hh.Enabled = True
txt_mm.Enabled = True
txt_mmpp.Enabled = True
txt_cantp.Enabled = True
txt_espera.Enabled = True
t_hh.Enabled = True
t_mh.Enabled = True
cbomed.Enabled = True

Check1.Enabled = True
borraesp
b_alta.Enabled = False
b_bustodo.Enabled = False
b_modif.Enabled = False
b_graba.Enabled = True
b_canc.Enabled = True
b_busca.Enabled = False
Frame2.Enabled = False
txt_base.SetFocus
data_espec.Recordset.AddNew

End Sub

Private Sub b_busca_Click()
Dim XRES As String
If XWeltipoU <> "USUARIOS" Then
   XRES = MsgBox("Desea Borrar??", vbCritical + vbYesNo, "Mensaje")
   If XRES = vbYes Then
      data_espec.Recordset.Delete
      data_espec.Refresh
      igualar
   End If
Else
   MsgBox "Usuario no autorizado"
End If

End Sub

Private Sub b_bustodo_Click()
frm_busespe.Show vbModal

End Sub

Private Sub b_canc_Click()
If XAlta = 1 Then
   data_espec.Recordset.CancelUpdate
   XAlta = 0
   borraesp
   If data_espec.Recordset.RecordCount > 0 Then
      data_espec.Recordset.MoveLast
   End If
   igualar
   b_alta.Enabled = True
   b_bustodo.Enabled = True
   b_modif.Enabled = True
   b_graba.Enabled = False
   b_canc.Enabled = False
   b_busca.Enabled = True
   Frame2.Enabled = True
   b_bustodo.SetFocus
   txt_base.Enabled = False
   txt_cod.Enabled = False
   txt_desc.Enabled = False
   txt_hh.Enabled = False
   t_hh.Enabled = False
   t_mh.Enabled = False
   txt_mm.Enabled = False
   txt_mmpp.Enabled = False
   txt_cantp.Enabled = False
   txt_espera.Enabled = False
   cbomed.Enabled = False
   Check1.Enabled = False
Else
   XAlta = 0
   borraesp
   If data_espec.Recordset.RecordCount > 0 Then
      data_espec.Recordset.MoveLast
   End If
   igualar
   b_alta.Enabled = True
   b_bustodo.Enabled = True
   b_modif.Enabled = True
   b_graba.Enabled = False
   b_canc.Enabled = False
   b_busca.Enabled = True
   Frame2.Enabled = True
   b_bustodo.SetFocus
   txt_base.Enabled = False
   txt_cod.Enabled = False
   txt_desc.Enabled = False
   txt_hh.Enabled = False
   txt_mm.Enabled = False
   txt_mmpp.Enabled = False
   txt_cantp.Enabled = False
   txt_espera.Enabled = False
   t_hh.Enabled = False
   t_mh.Enabled = False
   cbomed.Enabled = False
   
   Check1.Enabled = False
End If
End Sub

Private Sub b_cons_Click()
frm_consfechas.Show vbModal

End Sub

Private Sub b_graba_Click()
If t_mh.Text = "0" Then
   t_mh.Text = "00"
End If
If txt_mm.Text = "0" Then
   txt_mm.Text = "00"
End If

If t_hh.Text <> "" And t_mh.Text <> "" Then
   If t_hh.Text <> "0" And t_mh.Text <> "0" Then
        Dim Xh1, Xh2, Xm1, Xm2, Xtotmh, Xtotmm As Long
        Xh1 = txt_hh.Text
        Xm1 = txt_mm.Text
        Xh2 = t_hh.Text
        Xm2 = t_mh.Text
        Xtotmh = Xh2 - Xh1
        If Xtotmh < 0 Then
           Xtotmh = 60
        Else
           Xtotmh = Xtotmh * 60
        End If
        Xtotmm = Xm2 - Xm1
        If Xtotmm < 0 Then
           Xtotmm = Xtotmm + 60
        End If
        Xtotmm = Xtotmm + Xtotmh
        If txt_mmpp.Text = "" Then
           txt_mmpp.Text = 15
        End If
        txt_cantp.Text = Xtotmm / txt_mmpp.Text
        txt_cantp.Text = Int(txt_cantp.Text)
   End If
End If
If XAlta = 1 Then
   If txt_base.Text <> "" Then
      If txt_cod.Text <> "" Then
         data_buscod.Recordset.FindFirst "hora ='" & txt_cod.Text & "'"
         If Not data_buscod.Recordset.NoMatch Then
            MsgBox "Ya existe CODIGO", vbCritical, "Mensaje"
            txt_cod.SetFocus
         Else
             data_espec.Recordset("base") = txt_base.Text
             data_espec.Recordset("hora") = txt_cod.Text
             data_espec.Recordset("nom_medic") = Mid(txt_desc.Text, 1, 50)
             data_espec.Recordset("convenio") = txt_hh.Text
             data_espec.Recordset("moneda") = txt_mm.Text
             data_espec.Recordset("cod_medic") = txt_mmpp.Text
             data_espec.Recordset("imp_iva") = txt_cantp.Text
             data_espec.Recordset("mes_paga") = txt_espera.Text
             data_espec.Recordset("reg_cab") = Check1.value
             data_espec.Recordset("factura") = t_hh.Text
             data_espec.Recordset("cod_cli") = t_mh.Text
             If cbomed.Text <> "" Then
                data_medicos.RecordSource = "Select * from medicos where med_nombre ='" & cbomed.Text & "'"
                data_medicos.Refresh
                If data_medicos.Recordset.RecordCount > 0 Then
                   t_cmed.Text = data_medicos.Recordset("med_cod")
                Else
                   t_cmed.Text = 0
                   cbomed.Text = ""
                End If
             Else
                t_cmed.Text = 0
             End If
             data_espec.Recordset("cod_prod") = t_cmed.Text
             data_espec.Recordset.Update
                
             XAlta = 0
'''             borraesp
             data_espec.Refresh
             data_espec.Recordset.FindFirst "hora ='" & txt_cod.Text & "' And base =" & txt_base.Text
'             If data_espec.Recordset.RecordCount > 0 Then
'                data_espec.Recordset.MoveLast
'             End If
             borraesp
             igualar
             b_alta.Enabled = True
             b_bustodo.Enabled = True
             b_modif.Enabled = True
             b_graba.Enabled = False
             b_canc.Enabled = False
             b_busca.Enabled = True
             Frame2.Enabled = True
             b_bustodo.SetFocus
            txt_base.Enabled = False
            txt_cod.Enabled = False
            txt_desc.Enabled = False
            txt_hh.Enabled = False
            txt_mm.Enabled = False
            txt_mmpp.Enabled = False
            txt_cantp.Enabled = False
            txt_espera.Enabled = False
            t_hh.Enabled = False
            t_mh.Enabled = False
            cbomed.Enabled = False
            Check1.Enabled = False
         End If
      End If
   End If
Else
   If txt_base.Text <> "" Then
      If txt_cod.Text <> "" Then
         data_espec.Recordset.Edit
         data_espec.Recordset("base") = txt_base.Text
         data_espec.Recordset("hora") = txt_cod.Text
         data_espec.Recordset("nom_medic") = Mid(txt_desc.Text, 1, 75)
         data_espec.Recordset("convenio") = txt_hh.Text
         data_espec.Recordset("moneda") = txt_mm.Text
         data_espec.Recordset("cod_medic") = txt_mmpp.Text
         data_espec.Recordset("imp_iva") = txt_cantp.Text
         data_espec.Recordset("mes_paga") = txt_espera.Text
         data_espec.Recordset("reg_cab") = Check1.value
         data_espec.Recordset("factura") = t_hh.Text
         data_espec.Recordset("cod_cli") = t_mh.Text
         If cbomed.ListIndex >= 0 Then
            data_medicos.RecordSource = "Select * from medicos where med_nombre ='" & cbomed.Text & "'"
            data_medicos.Refresh
            If data_medicos.Recordset.RecordCount > 0 Then
               t_cmed.Text = data_medicos.Recordset("med_cod")
            Else
               t_cmed.Text = 0
               cbomed.Text = ""
            End If
         Else
            t_cmed.Text = 0
         End If
         data_espec.Recordset("cod_prod") = t_cmed.Text
         data_espec.Recordset.Update
         XAlta = 0
         borraesp
         If data_espec.Recordset.RecordCount > 0 Then
            data_espec.Recordset.MoveLast
         End If
         igualar
         b_alta.Enabled = True
         b_bustodo.Enabled = True
         b_modif.Enabled = True
         b_graba.Enabled = False
         b_canc.Enabled = False
         b_busca.Enabled = True
         Frame2.Enabled = True
         b_bustodo.SetFocus
        txt_base.Enabled = False
        txt_cod.Enabled = False
        txt_desc.Enabled = False
        txt_hh.Enabled = False
        txt_mm.Enabled = False
        txt_mmpp.Enabled = False
        txt_cantp.Enabled = False
        txt_espera.Enabled = False
        t_hh.Enabled = False
        t_mh.Enabled = False
        cbomed.Enabled = False
        Check1.Enabled = False
      End If
   End If
End If
End Sub

Private Sub b_imp_Click()
frm_fechasesp.Show vbModal

End Sub

Private Sub b_ing_Click()
WCodesp = txt_cod.Text
WNomesp = txt_desc.Text
WBase = txt_base.Text

frm_creaesp.Show vbModal

End Sub

Private Sub b_modif_Click()
XAlta = 0
txt_base.Enabled = True
txt_cod.Enabled = False
txt_desc.Enabled = True
txt_hh.Enabled = True
txt_mm.Enabled = True
txt_mmpp.Enabled = True
txt_cantp.Enabled = True
txt_espera.Enabled = True
t_hh.Enabled = True
cbomed.Enabled = True
t_mh.Enabled = True
b_alta.Enabled = False
b_bustodo.Enabled = False
b_modif.Enabled = False
b_graba.Enabled = True
b_canc.Enabled = True
b_busca.Enabled = False
Frame2.Enabled = False
Check1.Enabled = True

End Sub

Private Sub b_rec_Click()
Dim Xquefec As String
Dim Xrespborr, Xquerespes As String
Dim Baseaborrar As Database
Dim Espacioborra As Workspace
Dim Xlafecborra As String

If WElusuario = "CLAUDIA2" Or WElusuario = "CLAUDIA" Or WElusuario = "SDOMINGUEZ" Or WElusuario = "AACUÑA" Or WElusuario = "MIKAELA" Or WElusuario = "MCURBELO" Then
    Set Espacioborra = Workspaces(0)
    Set Baseaborrar = Espacioborra.OpenDatabase(App.Path & "\sapp.mdb")
    b_rec.Enabled = False
    Xquerespes = MsgBox("Desea BORRAR SOLO UNA FECHA??", vbExclamation + vbYesNoCancel, "Especialistas")
    If Xquerespes = vbYes Then
       xlafechaborra = InputBox("Ingrese fecha a BORRAR:(Ej.01/10/2009)", "Fechas de especialistas")
       If xlafechaborra <> "" Then
          frm_espec.MousePointer = 11
          Baseaborrar.Execute "Delete * from lista where fecha =#" & Format(xlafechaborra, "yyyy/mm/dd") & "# and base =" & txt_base.Text & " and cod ='" & txt_cod.Text & "'"
          Baseaborrar.Execute "Delete * from fechasesp where fecha =#" & Format(xlafechaborra, "yyyy/mm/dd") & "# and base =" & txt_base.Text & " and cod ='" & txt_cod.Text & "'"
          Baseaborrar.Execute "Delete * from mant_sol where cl_fnac =#" & Format(xlafechaborra, "yyyy/mm/dd") & "# and cl_grupo =" & txt_base.Text & " and cl_fax ='" & txt_cod.Text & "'"
       
          frm_espec.MousePointer = 0
          MsgBox "Fecha ELIMINADA!!"
       End If
    Else
       If Xquerespes = vbNo Then
          Xquefec = InputBox("Ingrese hasta que fecha desea borrar?(Ej.01/04/2008):", "Borrar listas")
          If Xquefec <> "" Then
             Xrespborr = MsgBox("Está seguro de borrar?", vbExclamation + vbYesNo, "Mensaje")
             If Xrespborr = vbYes Then
                frm_espec.MousePointer = 11
                Baseaborrar.Execute "Delete * from lista where fecha <=#" & Format(Xquefec, "yyyy/mm/dd") & "#"
                Baseaborrar.Execute "Delete * from fechasesp where fecha <=#" & Format(Xquefec, "yyyy/mm/dd") & "#"
                frm_espec.MousePointer = 0
                MsgBox "Proceso terminado"
             End If
          End If
       End If
    End If
    b_rec.Enabled = True
Else
    MsgBox "Usuario sin permiso"
End If

End Sub

Private Sub Form_Activate()
b_bustodo.SetFocus

End Sub

Public Function borraesp()
txt_base.Text = ""
txt_cod.Text = ""
txt_desc.Text = ""
txt_hh.Text = ""
txt_mm.Text = ""
txt_mmpp.Text = ""
txt_cantp.Text = ""
txt_espera.Text = ""
t_hh.Text = ""
t_mh.Text = ""
cbomed.Text = ""
Check1.value = 0

End Function

Public Function igualar()
If data_espec.Recordset.RecordCount > 0 Then
   txt_base.Text = data_espec.Recordset("base")
   txt_cod.Text = data_espec.Recordset("hora")
   txt_desc.Text = data_espec.Recordset("nom_medic")
   txt_hh.Text = data_espec.Recordset("convenio")
   txt_mm.Text = data_espec.Recordset("moneda")
   txt_mmpp.Text = Int(data_espec.Recordset("cod_medic"))
   txt_cantp.Text = Int(data_espec.Recordset("imp_iva"))
   txt_espera.Text = Int(data_espec.Recordset("mes_paga"))
   If IsNull(data_espec.Recordset("factura")) = False Then
      t_hh.Text = Int(data_espec.Recordset("factura"))
   Else
      t_hh.Text = 0
   End If
   If IsNull(data_espec.Recordset("cod_cli")) = False Then
      t_mh.Text = Int(data_espec.Recordset("cod_cli"))
   Else
      t_mh.Text = 0
   End If
   Check1.value = data_espec.Recordset("reg_cab")
   If IsNull(data_espec.Recordset("cod_prod")) = False Then
      If data_espec.Recordset("cod_prod") <> 0 Then
         data_medicos.RecordSource = "Select * from medicos where med_cod =" & data_espec.Recordset("cod_prod")
         data_medicos.Refresh
         If data_medicos.Recordset.RecordCount > 0 Then
            cbomed.Text = data_medicos.Recordset("med_nombre")
            t_cmed.Text = data_medicos.Recordset("med_cod")
         Else
            cbomed.Text = ""
            t_cmed.Text = 0
         End If
      Else
         cbomed.Text = ""
         t_cmed.Text = 0
      End If
   Else
      cbomed.Text = ""
      t_cmed.Text = 0
   End If
End If

End Function

Private Sub Form_Load()
data_med.DatabaseName = App.Path & "\sapp.mdb"
data_espec.DatabaseName = App.Path & "\sapp.mdb"
Data1.DatabaseName = App.Path & "\parse.mdb"
Data1.RecordSource = "parsec0"
Data1.Refresh
data_lis.DatabaseName = App.Path & "\sapp.mdb"
data_medicos.DatabaseName = App.Path & "\sapp.mdb"
data_medicos.RecordSource = "Select * from medicos order by med_nombre"
data_medicos.Refresh
If data_medicos.Recordset.RecordCount > 0 Then
   data_medicos.Recordset.MoveFirst
   Do While Not data_medicos.Recordset.EOF
      cbomed.AddItem data_medicos.Recordset("med_nombre")
      data_medicos.Recordset.MoveNext
   Loop
End If

data_buscod.DatabaseName = App.Path & "\sapp.mdb"
data_buscod.RecordSource = "lineas"
data_buscod.Refresh

If WElusuario = "JFERNAN" Or WElusuario = "CLAUDIA" Or WElusuario = "GFERNANDEZ" Or WElusuario = "MARIAROSA" Or WElusuario = "MPEREZ" Or WElusuario = "AACUÑA" Or _
   WElusuario = "MCURBELO" Or WElusuario = "JONATHAN" Or WElusuario = "SDOMINGUEZ" Or WElusuario = "MSANCHEZ" Or WElusuario = "GUSTAVO" Or WElusuario = "MIKAELA" Then
   Frame1.Enabled = True
   b_rec.Enabled = True
   data_espec.RecordSource = "Select * from lineas order by base"
   data_espec.Refresh
Else
   Frame1.Enabled = False
   b_rec.Enabled = False
   data_espec.RecordSource = "Select * from lineas where base =" & Data1.Recordset("base")
   data_espec.Refresh
End If
If data_espec.Recordset.RecordCount > 0 Then
   If IsNull(data_espec.Recordset("Base")) = True Then
      txt_base.Text = ""
   Else
      txt_base.Text = data_espec.Recordset("base")
   End If
   txt_cod.Text = data_espec.Recordset("hora")
   txt_desc.Text = data_espec.Recordset("nom_medic")
   txt_hh.Text = data_espec.Recordset("convenio")
   txt_mm.Text = data_espec.Recordset("moneda")
   txt_mmpp.Text = Int(data_espec.Recordset("cod_medic"))
   txt_cantp.Text = Int(data_espec.Recordset("imp_iva"))
   txt_espera.Text = Int(data_espec.Recordset("mes_paga"))
   If IsNull(data_espec.Recordset("cod_prod")) = False Then
      If data_espec.Recordset("cod_prod") <> 0 Then
         data_medicos.RecordSource = "Select * from medicos where med_cod =" & data_espec.Recordset("cod_prod")
         data_medicos.Refresh
         If data_medicos.Recordset.RecordCount > 0 Then
            cbomed.Text = data_medicos.Recordset("med_nombre")
            t_cmed.Text = data_medicos.Recordset("med_cod")
         Else
            cbomed.Text = ""
            t_cmed.Text = 0
         End If
      Else
         cbomed.Text = ""
         t_cmed.Text = 0
      End If
   Else
      cbomed.Text = ""
      t_cmed.Text = 0
   End If
   If IsNull(data_espec.Recordset("reg_cab")) = False Then
      If data_espec.Recordset("reg_cab") = 1 Then
         Check1.value = 1
      Else
         Check1.value = 0
      End If
   Else
      Check1.value = 0
   End If
End If
Dim Xfeccan As Date
Xfeccan = Date
data_cance.DatabaseName = App.Path & "\sapp.mdb"
data_cance.RecordSource = "Select * from fechasesp where fecha >=#" & Format(Xfeccan, "yyyy/mm/dd") & "# and codmed >=" & 0 & " and base =" & frm_menu.data_parse.Recordset("base")
data_cance.Refresh
If data_cance.Recordset.RecordCount > 0 Then
   MsgBox "EXISTE CONSULTA CANCELADA PARA ÉSTA BASE: " & data_cance.Recordset("desc") & " VERIFIQUE CON EL ADMINISTRADOR DE LAS FECHAS", vbInformation
End If


End Sub

Private Sub Option1_Click()
md.Visible = True
mh.Visible = True
Label9.Visible = False

End Sub

Private Sub Option2_Click()
md.Visible = False
mh.Visible = False
Label9.Visible = True

End Sub

Private Sub t_hh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_mh.SetFocus
End If

End Sub

Private Sub t_mh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_mmpp.SetFocus
End If

End Sub

Private Sub txt_base_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_cod.SetFocus
End If

End Sub

Private Sub txt_cantp_KeyPress(KeyAscii As Integer)
If txt_cantp.Text = "" Then
Else
   If KeyAscii = 13 Then
      txt_espera.SetFocus
   End If
End If

End Sub

Private Sub txt_cantp_LostFocus()
If txt_cantp.Text = "" Then
   txt_cantp.Text = 0
End If

End Sub

Private Sub txt_cod_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   txt_desc.SetFocus
End If

End Sub

Private Sub txt_desc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_hh.SetFocus
End If

End Sub

Private Sub txt_espera_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_graba.SetFocus
End If

End Sub

Private Sub txt_espera_LostFocus()
If txt_espera.Text = "" Then
   txt_espera.Text = 0
End If

End Sub

Private Sub txt_hh_KeyPress(KeyAscii As Integer)
If txt_hh.Text = "" Then
Else
   If KeyAscii = 13 Then
      txt_mm.SetFocus
   End If
End If

End Sub

Private Sub txt_hh_LostFocus()
If txt_hh.Text = "" Then
   MsgBox "Ingrese Hora"
   txt_hh.SetFocus
End If

End Sub

Private Sub txt_mm_KeyPress(KeyAscii As Integer)
If txt_mm.Text = "" Then
Else
   If KeyAscii = 13 Then
      t_hh.SetFocus
   End If
End If
End Sub

Private Sub txt_mm_LostFocus()
If txt_mm.Text = "" Then
   MsgBox "Ingrese MINUTOS"
   txt_mm.SetFocus
End If

End Sub

Private Sub txt_mmpp_KeyPress(KeyAscii As Integer)
If txt_mmpp.Text = "" Then
Else
   If KeyAscii = 13 Then
      txt_cantp.SetFocus
   End If
End If
End Sub

Private Sub txt_mmpp_LostFocus()
If txt_mmpp.Text = "" Then
   MsgBox "Ingrese DATO"
   txt_mmpp.SetFocus
End If

End Sub
