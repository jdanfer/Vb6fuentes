VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_ctrolstock 
   BackColor       =   &H0080C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de productos y stock"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9045
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_ctrolstock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   9045
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton b_borra 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3480
      Picture         =   "frm_ctrolstock.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Eliminar el registro seleccionado."
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton b_impr 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5160
      Picture         =   "frm_ctrolstock.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Menú de informes por pantalla o impresora de stock."
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton b_busca 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4320
      Picture         =   "frm_ctrolstock.frx":0F56
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Buscar datos"
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton b_cance 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2640
      Picture         =   "frm_ctrolstock.frx":14E0
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Cancelar grabado de registro."
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton b_graba 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1800
      Picture         =   "frm_ctrolstock.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Grabar los datos."
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton b_modif 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   960
      Picture         =   "frm_ctrolstock.frx":1FF4
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Modificar datos del registro seleccionado"
      Top             =   5040
      Width           =   495
   End
   Begin VB.CommandButton b_alta 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   120
      Picture         =   "frm_ctrolstock.frx":257E
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Agregar nuevo registro "
      Top             =   5040
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Datos del producto y stocks"
      Enabled         =   0   'False
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   8775
      Begin VB.TextBox t_alerta 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   7320
         TabIndex        =   31
         Top             =   3480
         Width           =   1335
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   495
         Left            =   2160
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   360
         Visible         =   0   'False
         Width           =   2175
      End
      Begin MSMask.MaskEdBox mfvence 
         Height          =   375
         Left            =   2280
         TabIndex        =   10
         Top             =   3480
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "frm_ctrolstock.frx":2B08
         Left            =   2280
         List            =   "frm_ctrolstock.frx":2B18
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   2880
         Width           =   3255
      End
      Begin VB.TextBox t_prec 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   7440
         TabIndex        =   8
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox t_obs 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2280
         MaxLength       =   200
         TabIndex        =   11
         Top             =   4080
         Width           =   6375
      End
      Begin VB.TextBox t_act 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2280
         TabIndex        =   7
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox t_bas 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   7440
         TabIndex        =   6
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox t_min 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2280
         TabIndex        =   5
         Top             =   1800
         Width           =   1335
      End
      Begin MSMask.MaskEdBox mfultact 
         Height          =   375
         Left            =   6960
         TabIndex        =   4
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfing 
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   1320
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox t_desc 
         Height          =   375
         Left            =   2280
         MaxLength       =   70
         TabIndex        =   2
         Top             =   840
         Width           =   6375
      End
      Begin VB.TextBox t_cod 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   2280
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C00000&
         Caption         =   "Alerta de consumo:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4920
         TabIndex        =   30
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C00000&
         Caption         =   "F. Vencimiento:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   28
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C00000&
         Caption         =   "Grupo:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   27
         Top             =   2880
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C00000&
         Caption         =   "Precio Unitario:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4920
         TabIndex        =   25
         Top             =   2280
         Width           =   2535
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C00000&
         Caption         =   "Observaciones:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   19
         Top             =   4080
         Width           =   2055
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C00000&
         Caption         =   "Stock ACTUAL:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C00000&
         Caption         =   "Stock BASICO:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4920
         TabIndex        =   17
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
         Caption         =   "Stock MINIMO:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   1800
         Width           =   2055
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "Ult. actualización:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4920
         TabIndex        =   15
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "Fecha ingreso:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   14
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Descripción:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "Código:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   6240
      Picture         =   "frm_ctrolstock.frx":2B46
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   1935
   End
End
Attribute VB_Name = "frm_ctrolstock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_alta_Click()
XAlta = 1
Frame1.Enabled = True
limpiast
Data1.RecordSource = "Select * from stock order by id DESC"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   t_cod.Text = Data1.Recordset("id") + 1
Else
   t_cod.Text = InputBox("Ingrese número de código de comienzo")
End If
t_desc.SetFocus
botondes

End Sub

Private Sub b_borra_Click()
Dim Xborsion As String
Xborsion = ""
If t_cod.Text = "" Then
Else
   Data1.RecordSource = "Select * from stock where id =" & t_cod.Text
   Data1.Refresh
   If Data1.Recordset.RecordCount > 0 Then
      Xborsion = MsgBox("Desea borrar el registro seleccionado?", vbInformation + vbYesNo, "Stock")
      If Xborsion = vbYes Then
         Data1.Recordset.Delete
         Data1.RecordSource = "stock"
         Data1.Refresh
         limpiast
      End If
   End If
End If

End Sub

Private Sub b_busca_Click()
frm_busstock.Show vbModal

End Sub

Private Sub b_cance_Click()
'If XAlta = 1 Then
'   Data1.Recordset.CancelUpdate
'End If
Frame1.Enabled = True
limpiast
Frame1.Enabled = False
botonhab

End Sub

Private Sub b_graba_Click()
Dim Xedita As Integer
Xedita = 0


If t_cod.Text <> "" Then
   If t_cod.Text > 0 Then
      If XAlta = 1 Then
         Data1.Recordset.AddNew
         Data1.Recordset("id") = t_cod.Text
         Data1.Recordset("descrip") = UCase(t_desc.Text)
         If t_min.Text = "" Then
            t_min.Text = 0
         End If
         Data1.Recordset("minimo") = t_min.Text
         If t_bas.Text = "" Then
            t_bas.Text = 0
         End If
         Data1.Recordset("basico") = t_bas.Text
         If t_act.Text = "" Then
            t_act.Text = 0
         End If
         Data1.Recordset("actual") = t_act.Text
         If t_prec.Text = "" Then
            t_prec.Text = 0
         End If
         Data1.Recordset("preuni") = t_prec.Text
         If mfing.Text = "__/__/____" Then
         Else
            Data1.Recordset("ingreso") = mfing.Text
         End If
         If mfultact.Text = "__/__/____" Then
         Else
            Data1.Recordset("ultact") = mfultact.Text
         End If
         If mfvence.Text = "__/__/____" Then
         Else
            Data1.Recordset("vence") = mfvence.Text
         End If
         Data1.Recordset("grupo") = Combo1.ListIndex
         Data1.Recordset("obs") = t_obs.Text
         If t_alerta.Text = "" Then
            Data1.Recordset("alerta") = 0
         Else
            Data1.Recordset("alerta") = t_alerta.Text
         End If
         Data1.Recordset.Update
      Else
'         data1.Recordset.Edit
'         Data1.Recordset("id") = t_cod.Text
         If Trim(Data1.Recordset("descrip")) <> Trim(UCase(t_desc.Text)) Then
            ConectarBD
            ConbdSapp.Open
            ConbdSapp.Execute "Update stock set descrip ='" & t_desc.Text & "' where id =" & t_cod.Text
            ConbdSapp.Close
                        
'            Data1.Recordset.Edit
'            Data1.Recordset("descrip") = Trim(UCase(t_desc.Text))
'            Data1.Recordset.Update
         End If
         If t_min.Text = "" Then
            t_min.Text = 0
         End If
         If CDbl(Data1.Recordset("minimo")) <> CDbl(t_min.Text) Then
            ConectarBD
            ConbdSapp.Open
            ConbdSapp.Execute "Update stock set minimo =" & t_min.Text & " where id =" & t_cod.Text
            ConbdSapp.Close
'            Data1.Recordset.Edit
'            Data1.Recordset("minimo") = t_min.Text
'            Data1.Recordset.Update
         End If
         If t_bas.Text = "" Then
            t_bas.Text = 0
         End If
         If CDbl(Data1.Recordset("basico")) <> CDbl(t_bas.Text) Then
            ConectarBD
            ConbdSapp.Open
            ConbdSapp.Execute "Update stock set basico =" & t_bas.Text & " where id =" & t_cod.Text
            ConbdSapp.Close
            
'            Data1.Recordset.Edit
'            Data1.Recordset("basico") = t_bas.Text
'            Data1.Recordset.Update
         End If
         If t_act.Text = "" Then
            t_act.Text = 0
         End If
         If CDbl(Data1.Recordset("actual")) <> CDbl(t_act.Text) Then
            ConectarBD
            ConbdSapp.Open
            ConbdSapp.Execute "Update stock set actual =" & t_act.Text & " where id =" & t_cod.Text
            ConbdSapp.Close
         End If
         If t_prec.Text = "" Then
            t_prec.Text = 0
         End If
         If CDbl(Data1.Recordset("preuni")) <> CDbl(t_prec.Text) Then
            ConectarBD
            ConbdSapp.Open
            ConbdSapp.Execute "Update stock set preuni =" & t_prec.Text & " where id =" & t_cod.Text
            ConbdSapp.Close
         End If
         If t_alerta.Text = "" Then
            t_alerta.Text = 0
         End If
         If IsNull(Data1.Recordset("alerta")) = False Then
            If CDbl(Data1.Recordset("alerta")) <> CDbl(t_alerta.Text) Then
               ConectarBD
               ConbdSapp.Open
               ConbdSapp.Execute "Update stock set alerta =" & t_alerta.Text & " where id =" & t_cod.Text
               ConbdSapp.Close
            End If
         Else
            ConectarBD
            ConbdSapp.Open
            ConbdSapp.Execute "Update stock set alerta =" & t_alerta.Text & " where id =" & t_cod.Text
            ConbdSapp.Close
         End If
         If mfing.Text = "__/__/____" Then
         Else
            If Format(Data1.Recordset("ingreso"), "yyyy/mm/dd") <> Format(mfing.Text, "yyyy/mm/dd") Then
               ConectarBD
               ConbdSapp.Open
               ConbdSapp.Execute "Update stock set ingreso ='" & Format(mfing.Text, "yyyy-mm-dd") & "' where id =" & t_cod.Text
               ConbdSapp.Close
               
'               Data1.Recordset.Edit
'               Data1.Recordset("ingreso") = mfing.Text
'               Data1.Recordset.Update
            End If
         End If
         If mfultact.Text = "__/__/____" Then
         Else
            If Format(Data1.Recordset("ultact"), "yyyy/mm/dd") <> Format(mfultact.Text, "yyyy/mm/dd") Then
               ConectarBD
               ConbdSapp.Open
               ConbdSapp.Execute "Update stock set ultact ='" & Format(mfultact.Text, "yyyy-mm-dd") & "' where id =" & t_cod.Text
               ConbdSapp.Close
               
'               Data1.Recordset.Edit
'               Data1.Recordset("ultact") = mfultact.Text
'               Data1.Recordset.Update
            End If
         End If
         If mfvence.Text = "__/__/____" Then
         Else
            If Format(Data1.Recordset("vence"), "yyyy/mm/dd") <> Format(mfvence.Text, "yyyy/mm/dd") Then
               ConectarBD
               ConbdSapp.Open
               ConbdSapp.Execute "Update stock set vence ='" & Format(mfvence.Text, "yyyy-mm-dd") & "' where id =" & t_cod.Text
               ConbdSapp.Close
               
'               Data1.Recordset.Edit
'               Data1.Recordset("vence") = mfvence.Text
'               Data1.Recordset.Update
            End If
         End If
         
         If Data1.Recordset("grupo") <> Combo1.ListIndex Then
            ConectarBD
            ConbdSapp.Open
            ConbdSapp.Execute "Update stock set grupo =" & Combo1.ListIndex & " where id =" & t_cod.Text
            ConbdSapp.Close
            
'            Data1.Recordset.Edit
'            Data1.Recordset("grupo") = Combo1.ListIndex
'            Data1.Recordset.Update
         End If
         If IsNull(Data1.Recordset("obs")) = False Then
            If Data1.Recordset("obs") <> t_obs.Text Then
               ConectarBD
               ConbdSapp.Open
               ConbdSapp.Execute "Update stock set obs ='" & t_obs.Text & "' where id =" & t_cod.Text
               ConbdSapp.Close
            End If
         Else
            If t_obs.Text <> "" Then
               ConectarBD
               ConbdSapp.Open
               ConbdSapp.Execute "Update stock set obs ='" & t_obs.Text & "' where id =" & t_cod.Text
               ConbdSapp.Close
            End If
            
'            Data1.Recordset.Edit
'            Data1.Recordset("obs") = t_obs.Text
'            Data1.Recordset.Update
         End If
      End If
   End If
End If
Frame1.Enabled = False
botonhab

End Sub

Private Sub b_impr_Click()
'Unload Me
frm_infstock.Show vbModal

End Sub

Private Sub b_modif_Click()
XAlta = 2
Frame1.Enabled = True
If t_cod.Text <> "" Then
   Data1.RecordSource = "Select * from stock where id =" & t_cod.Text
   Data1.Refresh
   t_desc.SetFocus
   t_cod.Enabled = False
   botondes
Else
   MsgBox "Hay error en la búsqueda, verifique!"
End If

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfvence.SetFocus
End If

End Sub



Private Sub Form_Load()
'Data1.DatabaseName = App.Path & "\" & Trim(Xlabdd)
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.RecordSource = "select * from stock where id =" & 100
Data1.Refresh

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub mfing_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfultact.SetFocus
End If

End Sub

Private Sub mfultact_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_min.SetFocus
End If

End Sub

Private Sub mfvence_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_obs.SetFocus
End If

End Sub

Private Sub t_act_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_prec.SetFocus
End If

End Sub

Private Sub t_bas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_act.SetFocus
End If

End Sub

Private Sub t_cod_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
   t_desc.SetFocus
End If

End Sub

Private Sub t_desc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(chr(KeyAscii)))
If KeyAscii = 13 Then
   mfing.SetFocus
End If

End Sub

Private Sub t_min_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_bas.SetFocus
End If

End Sub

Private Sub t_obs_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   b_graba.SetFocus
End If

End Sub

Private Sub t_prec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo1.SetFocus
End If

End Sub

Public Function limpiast()
t_cod.Text = 0
t_desc.Text = ""
mfing.Text = "__/__/____"
mfultact.Text = "__/__/____"
t_min.Text = ""
t_bas.Text = ""
t_act.Text = ""
t_prec.Text = ""
Combo1.ListIndex = 0
mfvence.Text = "__/__/____"
t_obs.Text = ""

End Function

Public Function botondes()
b_alta.Enabled = False
b_modif.Enabled = False
b_graba.Enabled = True
b_cance.Enabled = True
b_borra.Enabled = False
b_busca.Enabled = False
b_impr.Enabled = False
b_borra.Enabled = False

End Function

Public Function botonhab()
b_alta.Enabled = True
b_modif.Enabled = True
b_graba.Enabled = False
b_cance.Enabled = False
b_borra.Enabled = True
b_busca.Enabled = True
b_impr.Enabled = True
b_borra.Enabled = True

End Function
