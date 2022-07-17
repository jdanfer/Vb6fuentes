VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frm_pendllap 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pendientes PEDIATRIA"
   ClientHeight    =   10305
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   15000
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_pendllap.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10305
   ScaleWidth      =   15000
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cbocodfin 
      Height          =   360
      ItemData        =   "frm_pendllap.frx":0442
      Left            =   11040
      List            =   "frm_pendllap.frx":044F
      Style           =   2  'Dropdown List
      TabIndex        =   23
      Top             =   7560
      Width           =   1815
   End
   Begin VB.Data data_med 
      Caption         =   "data_med"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox t_diag 
      Height          =   375
      Left            =   8640
      MaxLength       =   70
      TabIndex        =   21
      Top             =   7080
      Width           =   4215
   End
   Begin VB.TextBox t_mov 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   8640
      TabIndex        =   19
      Top             =   7560
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   360
      Left            =   8640
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   6600
      Width           =   4215
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   13080
      Picture         =   "frm_pendllap.frx":046A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Informes"
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Pasar llamado"
      Height          =   495
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Data data_azul 
      Caption         =   "data_azul"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7680
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FF80&
      Caption         =   "Cerrar llamado"
      Height          =   495
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Data Data4 
      Caption         =   "Data4"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   11400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8280
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
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
      Top             =   8040
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   13800
      Top             =   7680
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   10200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8280
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   8040
      Visible         =   0   'False
      Width           =   3060
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_pendllap.frx":09F4
      Height          =   5415
      Left            =   120
      OleObjectBlob   =   "frm_pendllap.frx":0A08
      TabIndex        =   1
      Top             =   960
      Width           =   14655
   End
   Begin MSComctlLib.TabStrip tabver 
      Height          =   6255
      Left            =   0
      TabIndex        =   0
      Top             =   240
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   11033
      MultiRow        =   -1  'True
      Style           =   2
      Separators      =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "PENDIENTES"
            Key             =   "a"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "EN CURSO"
            Key             =   "b"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label15 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cod.Final:"
      Height          =   375
      Left            =   9960
      TabIndex        =   22
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Label Label14 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Diagnóstico:"
      Height          =   375
      Left            =   7320
      TabIndex        =   20
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label Label13 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Móvil:"
      Height          =   375
      Left            =   7320
      TabIndex        =   18
      Top             =   7560
      Width           =   1335
   End
   Begin VB.Label Label12 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Médico"
      Height          =   255
      Left            =   7320
      TabIndex        =   16
      Top             =   6600
      Width           =   1335
   End
   Begin VB.Label Label11 
      BackColor       =   &H0080FFFF&
      Caption         =   "Datos opcionales para cierre de llamado:"
      ForeColor       =   &H00C00000&
      Height          =   735
      Left            =   5520
      TabIndex        =   15
      Top             =   6600
      Width           =   1695
   End
   Begin VB.Label Label10 
      Height          =   495
      Left            =   2400
      TabIndex        =   11
      Top             =   9600
      Width           =   10935
   End
   Begin VB.Label Label9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "MOTIVO:"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   9600
      Width           =   2295
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   0
      X2              =   17640
      Y1              =   10200
      Y2              =   10200
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      X1              =   0
      X2              =   17640
      Y1              =   8040
      Y2              =   8040
   End
   Begin VB.Label Label8 
      Height          =   255
      Left            =   7440
      TabIndex        =   9
      Top             =   9240
      Width           =   5895
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Convenio:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   8
      Top             =   9240
      Width           =   1815
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   9240
      Width           =   3015
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "TELEFONO:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   9240
      Width           =   2295
   End
   Begin VB.Label Label4 
      Height          =   615
      Left            =   2400
      TabIndex        =   5
      Top             =   8520
      Width           =   10935
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "DIRECCION:"
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   4
      Top             =   8520
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Más datos del llamado seleccionado:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   8280
      Width           =   4815
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "F1 = Cambia de Pendientes a En Curso"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   7200
      Width           =   4815
   End
   Begin VB.Image Image1 
      Height          =   2055
      Left            =   13200
      Picture         =   "frm_pendllap.frx":279F
      Stretch         =   -1  'True
      Top             =   7560
      Width           =   1455
   End
End
Attribute VB_Name = "frm_pendllap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub tab_Click()


End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_diag.SetFocus
End If

End Sub

Private Sub Command1_Click()
'3075
On Error GoTo Verqueer

data_azul.RecordSource = "Select * from llamado where nrolla =" & DBGrid1.Columns(13)
data_azul.Refresh
If data_azul.Recordset.RecordCount > 0 Then
   data_azul.Recordset.Edit
   data_azul.Recordset("pend") = 2
   data_azul.Recordset("fec_rea") = Format(Date, "dd/mm/yyyy")
   data_azul.Recordset("hor_rea") = Format(Time, "HH:mm")
   If Combo1.ListIndex >= 0 Then
      data_med.Recordset.FindFirst "med_nombre ='" & Combo1.Text & "'"
      If Not data_med.Recordset.NoMatch Then
         data_azul.Recordset("codmed") = data_med.Recordset("med_cod")
         data_azul.Recordset("nommed") = data_med.Recordset("med_nombre")
      End If
   End If
   If t_diag.Text <> "" Then
      data_azul.Recordset("diag") = t_diag.Text
   End If
   If t_mov.Text <> "" Then
      data_azul.Recordset("movilpas") = t_mov.Text
   End If
   If cbocodfin.ListIndex = 0 Then
      data_azul.Recordset("colormot") = "V"
   Else
      If cbocodfin.ListIndex = 1 Then
         data_azul.Recordset("colormot") = "A"
      Else
         If cbocodfin.ListIndex = 2 Then
            data_azul.Recordset("colormot") = "R"
         Else
            data_azul.Recordset("colormot") = "V"
         End If
      End If
   End If
   data_azul.Recordset("editando") = 1
   data_azul.Recordset.Update
   MsgBox "Llamado cerrado"
    If tabver.SelectedItem.index = 1 Then
       DBGrid1.BackColor = &HFFC0C0
       Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2016", "yyyy/mm/dd") & "# and pend not in (1,2) and codmot ='" & "Z" & "' order by mm,nrolla"
       Data1.Refresh
    Else
       If tabver.SelectedItem.index = 2 Then
          DBGrid1.BackColor = &HC0C0FF
          Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2016", "yyyy/mm/dd") & "# and pend in (1) and codmot ='" & "Z" & "' order by mm,nrolla"
          Data1.Refresh
       End If
    End If
End If

Exit Sub

Verqueer:
         If Err.Number = 3075 Then
            MsgBox "Seleccione un llamado"
         Else
            MsgBox "Verifique si seleccionó el llamado"
         End If
         
   
End Sub

Private Sub Command2_Click()
On Error GoTo Verqueer2

data_azul.RecordSource = "Select * from llamado where nrolla =" & DBGrid1.Columns(13)
data_azul.Refresh
If data_azul.Recordset.RecordCount > 0 Then
   data_azul.Recordset.Edit
   data_azul.Recordset("pend") = 1
   data_azul.Recordset("fecpas") = Format(Date, "dd/mm/yyyy")
   data_azul.Recordset("horpas") = Format(Time, "HH:mm")
   data_azul.Recordset("editando") = 1
   data_azul.Recordset.Update
   MsgBox "Llamado Pasado"
    If tabver.SelectedItem.index = 1 Then
       DBGrid1.BackColor = &HFFC0C0
       Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2016", "yyyy/mm/dd") & "# and pend not in (1,2) and codmot ='" & "Z" & "' order by mm,nrolla"
       Data1.Refresh
    Else
       If tabver.SelectedItem.index = 2 Then
          DBGrid1.BackColor = &HC0C0FF
          Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2016", "yyyy/mm/dd") & "# and pend in (1) and codmot ='" & "Z" & "' order by mm,nrolla"
          Data1.Refresh
       End If
    End If
  
End If
Exit Sub

Verqueer2:
         If Err.Number = 3075 Then
            MsgBox "Seleccione un llamado"
         Else
            MsgBox "Verifique si seleccionó el llamado"
         End If
          
End Sub

Private Sub Command3_Click()
frm_infazul.Show vbModal

End Sub

Private Sub DBGrid1_DblClick()
'MsgBox "FAVOR SELECCIONAR PRESIONANDO LA TECLA ENTER. GRACIAS!", vbCritical, "DESPACHO"
'MsgBox "ES: " & DBGrid1.Columns(13)
On Error GoTo Quepasoalabrir2

'frm_largador.data_lla.Refresh

'frm_largador.data_lla.Recordset.FindFirst "nrolla =" & DBGrid1.Columns(13)
frm_largador.data_lla.RecordSource = "Select * from llamado where nrolla =" & DBGrid1.Columns(13)
frm_largador.data_lla.Refresh

'If Not frm_largador.data_lla.Recordset.NoMatch Then
If frm_largador.data_lla.Recordset.RecordCount > 0 Then
    Data1.Recordset.FindFirst "nrolla =" & DBGrid1.Columns(13)
    If Not Data1.Recordset.NoMatch Then
        frm_largador.txt_nro.Text = Data1.Recordset("nrolla")
        frm_largador.mfecha.Text = Format(Data1.Recordset("fecha"), "dd/mm/yyyy")
        frm_largador.txt_hora.Text = Format(Data1.Recordset("hora"), "HH:mm")
        frm_largador.txt_usua.Text = Data1.Recordset("usuario")
        If IsNull(Data1.Recordset("activo")) = False Then
           frm_largador.Label3.Caption = Format(Data1.Recordset("activo"), "HH:mm:ss")
        Else
           frm_largador.Label3.Caption = "00:00:00"
        End If
        If IsNull(Data1.Recordset("nomodif")) = False Then
           frm_largador.Check1.Value = Data1.Recordset("nomodif")
        Else
           frm_largador.Check1.Value = 0
        End If
        If IsNull(Data1.Recordset("matric")) = False Then
           frm_largador.txt_mat.Text = Data1.Recordset("matric")
        Else
           frm_largador.txt_mat.Text = 0
        End If
        If IsNull(Data1.Recordset("nombre")) = False Then
           frm_largador.txt_nomb.Text = Data1.Recordset("nombre")
        Else
           frm_largador.txt_nomb.Text = ""
        End If
        If IsNull(Data1.Recordset("edad")) = False Then
           frm_largador.txt_edad.Text = Data1.Recordset("edad")
        Else
           frm_largador.txt_edad.Text = ""
        End If
        If IsNull(Data1.Recordset("mes")) = False Then
           If Data1.Recordset("mes") > 10 Then
              frm_largador.txt_costo.Text = Data1.Recordset("mes")
           Else
              frm_largador.txt_costo.Text = 0
           End If
        Else
           frm_largador.txt_costo.Text = 0
        End If
        If IsNull(Data1.Recordset("timbre")) = False Then
           frm_largador.cbotimbre.ListIndex = Data1.Recordset("timbre")
        Else
           frm_largador.cbotimbre.ListIndex = -1
        End If
        If IsNull(Data1.Recordset("valor_timbre")) = False Then
           frm_largador.t_timbre.Text = Data1.Recordset("valor_timbre")
        Else
           frm_largador.t_timbre.Text = ""
        End If
        If IsNull(Data1.Recordset("aft")) = False Then
           frm_largador.Label40.Caption = "AFT:" & Data1.Recordset("aft")
        Else
           frm_largador.Label40.Caption = ""
        End If
        If Format(Data1.Recordset("fecha"), "yyyy/mm/dd") >= Format("2016/12/01", "yyyy/mm/dd") Then
           If IsNull(Data1.Recordset("realiza")) = False Then
              frm_largador.chtmut.Value = Data1.Recordset("realiza")
           Else
              frm_largador.chtmut.Value = 0
           End If
        Else
           frm_largador.chtmut.Value = 0
        End If
        If IsNull(Data1.Recordset("ano")) = False Then
           If Data1.Recordset("ano") > 10 Then
              frm_largador.txt_boleta.Text = Data1.Recordset("ano")
           Else
              frm_largador.txt_boleta.Text = 0
           End If
        Else
           frm_largador.txt_boleta.Text = 0
        End If
        If IsNull(Data1.Recordset("unied")) = False Then
           If Data1.Recordset("unied") = 3 Then
              frm_largador.cboed.ListIndex = 0
           Else
              If Data1.Recordset("unied") = 2 Then
                 frm_largador.cboed.ListIndex = 1
              Else
                 If Data1.Recordset("unied") = 1 Then
                    frm_largador.cboed.ListIndex = 2
                 Else
                    frm_largador.cboed.ListIndex = 0
                 End If
              End If
           End If
        Else
           frm_largador.cboed.ListIndex = 0
        End If
        If IsNull(Data1.Recordset("categ")) = False Then
           frm_largador.txt_cat.Text = Data1.Recordset("categ")
        Else
           frm_largador.txt_cat.Text = ""
        End If
        If IsNull(Data1.Recordset("nomcat")) = False Then
           frm_largador.txt_nomcat.Text = Data1.Recordset("nomcat")
        Else
           frm_largador.txt_nomcat.Text = ""
        End If
        If IsNull(Data1.Recordset("ci")) = False Then
           frm_largador.txt_ced.Text = Int(Data1.Recordset("ci"))
        Else
           frm_largador.txt_ced.Text = 0
        End If
        If IsNull(Data1.Recordset("telef")) = False Then
           frm_largador.txt_tel.Text = Data1.Recordset("telef")
        Else
           frm_largador.txt_tel.Text = ""
        End If
        If IsNull(Data1.Recordset("codzon")) = False Then
           If Data1.Recordset("codzon") = 2 Then
              frm_largador.cbozona.ListIndex = 1
           Else
              If Data1.Recordset("codzon") = 3 Then
                 frm_largador.cbozona.ListIndex = 2
              Else
                 If Data1.Recordset("codzon") = 4 Then
                    frm_largador.cbozona.ListIndex = 3
                 Else
                    If Data1.Recordset("codzon") = 5 Then
                       frm_largador.cbozona.ListIndex = 4
                    Else
                       If Data1.Recordset("codzon") = 6 Then
                          frm_largador.cbozona.ListIndex = 5
                       Else
                          If Data1.Recordset("codzon") = 7 Then
                             frm_largador.cbozona.ListIndex = 6
                          Else
                             frm_largador.cbozona.ListIndex = 0
                          End If
                       End If
                    End If
                 End If
              End If
           End If
        Else
           frm_largador.cbozona.ListIndex = 0
        End If
        If IsNull(Data1.Recordset("base")) = False Then
           frm_largador.cbobase.Text = Data1.Recordset("base")
        Else
           frm_largador.cbobase.Text = 0
        End If
        If IsNull(Data1.Recordset("referen")) = False Then
           frm_largador.txt_direc.Text = Data1.Recordset("referen")
        Else
           frm_largador.txt_direc.Text = ""
        End If
        If IsNull(Data1.Recordset("motcon")) = False Then
           frm_largador.txt_ante.Text = Data1.Recordset("motcon")
        Else
           frm_largador.txt_ante.Text = ""
        End If
        If IsNull(Data1.Recordset("obs")) = False Then
           frm_largador.txt_obs.Text = Data1.Recordset("obs")
        Else
           frm_largador.txt_obs.Text = ""
        End If
        If IsNull(Data1.Recordset("obsmot")) = False Then
           frm_largador.txt_mot.Text = Data1.Recordset("obsmot")
        Else
           frm_largador.txt_mot.Text = ""
        End If
        If IsNull(Data1.Recordset("codmot")) = False Then
           If Data1.Recordset("codmot") = "R" Then
              frm_largador.cbocolor.ListIndex = 2
           Else
              If Data1.Recordset("codmot") = "A" Then
                 frm_largador.cbocolor.ListIndex = 1
              Else
                 If Data1.Recordset("codmot") = "C" Then
                    frm_largador.cbocolor.ListIndex = 3
                 Else
                    If Data1.Recordset("codmot") = "Z" Then
                       frm_largador.cbocolor.ListIndex = 4
                    Else
                       If Data1.Recordset("codmot") = "N" Then
                          frm_largador.cbocolor.ListIndex = 5
                       Else
                          frm_largador.cbocolor.ListIndex = 0
                       End If
                    End If
                 End If
              End If
           End If
        Else
           frm_largador.cbocolor.ListIndex = 0
        End If
        If IsNull(Data1.Recordset("motmov")) = True Then
           frm_largador.txt_locali.Text = ""
        Else
           frm_largador.txt_locali.Text = Data1.Recordset("motmov")
        End If
        If frm_largador.cbocolor.Text = "VERDE" Then
           frm_largador.cbocolor.BackColor = &HC000&
        Else
           If frm_largador.cbocolor.Text = "ROJO" Then
              frm_largador.cbocolor.BackColor = &HFF&
           Else
              If frm_largador.cbocolor.Text = "AMARILLO" Then
                 frm_largador.cbocolor.BackColor = &HFFFF&
              Else
                 If frm_largador.cbocolor.Text = "CELESTE" Then
                    frm_largador.cbocolor.BackColor = &HFFFF00
                 Else
                    If frm_largador.cbocolor.Text = "AZUL" Then
                       frm_largador.cbocolor.BackColor = &HC00000
                    Else
                       If frm_largador.cbocolor.Text = "NEGRO" Then
                          frm_largador.cbocolor.BackColor = &H80000006
                       Else
                          frm_largador.cbocolor.BackColor = &HFFFFFF
                       End If
                    End If
                 End If
              End If
           End If
        End If
        
        Data3.RecordSource = "Select * from resplla where nro =" & DBGrid1.Columns(13)
        Data3.Refresh
        If Data3.Recordset.RecordCount > 0 Then
           If IsNull(Data3.Recordset("telef")) = False Then
              If Data3.Recordset("telef") = "RECIBO" Then
                 frm_largador.Combo2.ListIndex = 0
              Else
                 If Data3.Recordset("telef") = "CONFORME" Then
                    frm_largador.Combo2.ListIndex = 1
                 Else
                    frm_largador.Combo2.ListIndex = -1
                 End If
              End If
           Else
              frm_largador.Combo2.ListIndex = -1
           End If
           
           If IsNull(Data3.Recordset("mes")) = False Then
              frm_largador.t_codced.Text = Int(Data3.Recordset("mes"))
           Else
              frm_largador.t_codced.Text = 0
           End If
           If IsNull(Data3.Recordset("pasado")) = False Then
              frm_largador.Check4.Value = Data3.Recordset("pasado")
           Else
              frm_largador.Check4.Value = 0
           End If
        Else
           frm_largador.Check4.Value = 0
        End If
        
        Unload Me
    Else
        MsgBox "No se encontró el registro, VERIFIQUE!!", vbCritical
    End If
Else
    MsgBox "Error en la búsqueda, seleccione nuevamente " & Data1.Recordset("nrolla"), vbInformation, "Mensaje"
    DBGrid1.SetFocus
End If
Exit Sub

Quepasoalabrir2:
            If Err.Number > 0 Then
               MsgBox " ERROR: " & str(Err.Number) & " Vuelva a intentar abrir"
               Unload Me
            End If
          
End Sub

Private Sub DBGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo Quepasaalabrir2

If KeyCode = vbKeyReturn Then
   DBGrid1_DblClick

End If

Exit Sub

Quepasaalabrir2:
                If Err.Number > 0 Then
                   MsgBox "Error al abrir, reintente o cierre ésta pantalla y vuelva a abrir", vbInformation
                End If
                
End Sub

Private Sub DBGrid1_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Data1.Recordset.RecordCount > 0 Then
    If IsNull(Data1.Recordset("direcc")) = False Then
       If IsNull(Data1.Recordset("referen")) = False Then
          Label4.Caption = Data1.Recordset("referen")
       Else
          Label4.Caption = Data1.Recordset("direcc")
       End If
       If IsNull(Data1.Recordset("telef")) = False Then
          Label6.Caption = Data1.Recordset("telef")
       Else
          Label6.Caption = ""
       End If
       If IsNull(Data1.Recordset("nomcat")) = False Then
          Label8.Caption = Data1.Recordset("nomcat")
       Else
          Label8.Caption = ""
       End If
       If IsNull(Data1.Recordset("motcon")) = False Then
          If IsNull(Data1.Recordset("obsmot")) = False Then
             Label10.Caption = Data1.Recordset("motcon") + " " + Data1.Recordset("obsmot")
          Else
             Label10.Caption = Data1.Recordset("motcon")
          End If
       Else
          Label10.Caption = ""
       End If
       
    End If
End If

End Sub


Private Sub Form_Deactivate()
'Timer1.Enabled = False


End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyF1 Then
   If tabver.SelectedItem.index = 1 Then
      tabver.Tabs(2).Selected = True
   Else
      tabver.Tabs(1).Selected = True
   End If
End If

End Sub

Private Sub Form_Load()
Dim Xlaf As Date
Xlaf = Date - 15
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_azul.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_med.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_med.RecordSource = "Select * from medicos order by med_nombre"
data_med.Refresh
If data_med.Recordset.RecordCount > 0 Then
   data_med.Recordset.MoveFirst
   Do While Not data_med.Recordset.EOF
      Combo1.AddItem data_med.Recordset("med_nombre")
      data_med.Recordset.MoveNext
   Loop
End If
cbocodfin.ListIndex = 0

'Data1.RecordSource = "llamado"
'Data1.Refresh
tabver.Refresh
'tabver.SelectedItem.Selected = True
If tabver.SelectedItem.index = 1 Then
   DBGrid1.BackColor = &HFFC0C0
   Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2016", "yyyy/mm/dd") & "# and pend not in (1,2,6) and codmot ='" & "Z" & "' order by mm,nrolla"
   Data1.Refresh
Else
   If tabver.SelectedItem.index = 2 Then
      DBGrid1.BackColor = &HC0C0FF
      Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2016", "yyyy/mm/dd") & "# and pend in (1,6) and codmot ='" & "Z" & "' order by mm,nrolla"
      Data1.Refresh
   End If
End If
Data2.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data3.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data4.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data4.RecordSource = "medicos"
Data4.Refresh

End Sub

Private Sub Form_Resize()
With Image1
     .Left = 0
     .Top = 0
     .Height = Me.Height
     .Width = Me.Width
End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Timer1.Enabled = False

End Sub

Private Sub t_diag_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_mov.SetFocus
End If

End Sub

Private Sub t_mov_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command1.SetFocus
End If

End Sub

Private Sub tabver_Click()
Dim Xlaf As Date
Xlaf = Date - 15

If tabver.SelectedItem.index = 1 Then
   DBGrid1.BackColor = &HFFC0C0
   Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2016", "yyyy/mm/dd") & "# and pend not in (1,2,6) and codmot ='" & "Z" & "' order by mm,nrolla"
   Data1.Refresh
Else
   If tabver.SelectedItem.index = 2 Then
      DBGrid1.BackColor = &HC0C0FF
      Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2016", "yyyy/mm/dd") & "# and pend in (1,6) and codmot ='" & "Z" & "' order by mm,nrolla"
      Data1.Refresh
   End If
End If

End Sub

Private Sub tabver_GotFocus()
If tabver.SelectedItem.index = 1 Then
   DBGrid1.BackColor = &HFFC0C0
   Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2016", "yyyy/mm/dd") & "# and pend not in (1,2,6) and codmot ='" & "Z" & "' order by mm,nrolla"
   Data1.Refresh
Else
   If tabver.SelectedItem.index = 2 Then
      DBGrid1.BackColor = &HC0C0FF
      Data1.RecordSource = "Select * from llamado where fecha >=#" & Format("01/07/2016", "yyyy/mm/dd") & "# and pend in (1,6) and codmot ='" & "Z" & "' order by mm,nrolla"
      Data1.Refresh
   End If
End If

End Sub

Private Sub tabver_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   DBGrid1.SetFocus
End If

End Sub


Private Sub Timer1_Timer()
'Data1.Refresh

End Sub
