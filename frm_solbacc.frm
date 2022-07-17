VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_solbacc 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Acciones al registro de solicitud de baja"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8880
   Icon            =   "frm_solbacc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   8880
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_paradm 
      Caption         =   "data_paradm"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   7200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Data data_accadm 
      Caption         =   "data_accadm"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   3600
      Visible         =   0   'False
      Width           =   2895
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
      Top             =   3480
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.CommandButton b_alta 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_solbacc.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Grabar nueva acción"
      Top             =   3360
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos de la acción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3015
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   8415
      Begin VB.TextBox t_accion 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   1080
         Width           =   8055
      End
      Begin MSMask.MaskEdBox mhora 
         Height          =   375
         Left            =   4440
         TabIndex        =   5
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "HH:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mf 
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VB.Label Label5 
         BackColor       =   &H00FF0000&
         Caption         =   "Detalles de la acción"
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
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label labusu 
         BackColor       =   &H00C00000&
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
         Left            =   6480
         TabIndex        =   7
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "Usuario:"
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
         Left            =   5520
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Hora:"
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
         Left            =   3600
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Fecha:"
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
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Label labcod 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5880
      TabIndex        =   10
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frm_solbacc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_alta_Click()
Dim Confir As String
Dim ConfAdm As String

If t_accion.Text = "" Then
   MsgBox "No hay datos para grabar", vbCritical
Else
   Confir = MsgBox("Confirma que desea agregar esta acción?", vbInformation + vbYesNo, "Acciones")
   If Confir = vbYes Then
      Data1.Recordset.AddNew
      Data1.Recordset("idid") = labcod.Caption
      Data1.Recordset("fecha") = mf.Text
      Data1.Recordset("hora") = mhora.Text
      Data1.Recordset("usuario") = labusu.Caption
      Data1.Recordset("accion") = t_accion.Text
      Data1.Recordset.Update
      Data1.RecordSource = "select * from solbaja_acc where idid =" & labcod.Caption
      Data1.Refresh
      frm_solicitudbaja.t_accion.Text = ""
      If Data1.Recordset.RecordCount > 0 Then
         Data1.Recordset.MoveFirst
         Do While Not Data1.Recordset.EOF
            frm_solicitudbaja.t_accion.Text = frm_solicitudbaja.t_accion.Text & Data1.Recordset("fecha") & "--" & Data1.Recordset("hora") & "--" & "--" & Data1.Recordset("usuario") & "--" & Data1.Recordset("accion")
            frm_solicitudbaja.t_accion.Text = frm_solicitudbaja.t_accion.Text & Chr(13) & Chr(10) & "-----------------------------------------------------" & Chr(13) & Chr(10)
            Data1.Recordset.MoveNext
         Loop
      End If
         
      ConfAdm = MsgBox("Registro grabado. Desea agregar ésta acción al historial administrativo?", vbInformation + vbYesNo, "Acciones")
      If ConfAdm = vbYes Then
         data_paradm.Recordset.Edit
         data_paradm.Recordset("nro_accadm") = data_paradm.Recordset("nro_accadm") + 1
         data_paradm.Recordset.Update
    
         data_accadm.Recordset.AddNew
         data_accadm.Recordset("cl_fultpag") = Format(mf.Text, "dd/mm/yyyy")
         data_accadm.Recordset("estado") = data_paradm.Recordset("nro_accadm")
         data_accadm.Recordset("cl_codigo") = data_paradm.Recordset("nro_accadm")
         data_accadm.Recordset("cl_ruc") = Format(Time, "HH:mm")
         data_accadm.Recordset("cl_atrasop") = frm_solicitudbaja.t_base.Text
         data_accadm.Recordset("cl_descpag") = WElusuario
         data_accadm.Recordset("cl_zona") = frm_solicitudbaja.t_mat.Text
         data_accadm.Recordset("cl_nro_sup") = 0
         data_accadm.Recordset("info_debit") = t_accion.Text
         data_accadm.Recordset("cl_desc1") = frm_solicitudbaja.t_cel.Text
         data_accadm.Recordset.Update
      End If
      
      Unload Me
   End If
End If

End Sub

Private Sub Form_Load()
mf.Text = Date
mhora.Text = Format(Time, "HH:mm")
labusu.Caption = WElusuario
labcod.Caption = frm_solicitudbaja.t_cod.Text

Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.RecordSource = "select * from solbaja_acc where idid =" & labcod.Caption
Data1.Refresh

data_accadm.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_accadm.RecordSource = "Select * from mant_sol where cl_codigo =" & 30468 & " and estado is not null"
data_accadm.Refresh
data_paradm.DatabaseName = App.path & "\paramb.mdb"
data_paradm.RecordSource = "paramb"
data_paradm.Refresh

End Sub
