VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm_cmtresult 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "CMT Resultados"
   ClientHeight    =   3990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7995
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   7995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_par 
      Caption         =   "data_par"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1920
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.Data data_medsapp 
      Caption         =   "data_medsapp"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2160
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data data_conshc 
      Caption         =   "data_conshc"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   2520
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Data data_hce 
      Caption         =   "data_hce"
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
      Top             =   3360
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      Top             =   2880
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   7200
      Picture         =   "frm_cmtresult.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir sin grabar"
      Top             =   3360
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   5880
      Picture         =   "frm_cmtresult.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Grabar datos y salir."
      Top             =   3360
      Width           =   495
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   1440
      Width           =   4455
   End
   Begin MSMask.MaskEdBox mhora 
      Height          =   375
      Left            =   6720
      TabIndex        =   4
      Top             =   840
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
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
   Begin MSMask.MaskEdBox mfec 
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   840
      Width           =   1575
      _ExtentX        =   2778
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
   Begin VB.Label labid 
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00404040&
      Caption         =   "En suma:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404040&
      Caption         =   "Fecha y hora de cierre:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "frm_cmtresult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim XcedDoc, Xhcesi As String

If mfec.Text <> "__/__/____" Then
   If mhora.Text <> "__:__" Then
      If Trim(Text1.Text) <> "" Then
         Data1.Recordset.Edit
         Data1.Recordset("obs_cmt") = Text1.Text
         Data1.Recordset("fecha_cierre") = mfec.Text
         Data1.Recordset("hora_cierre") = mhora.Text
         Data1.Recordset("nom_usu") = WElusuario
         Data1.Recordset.Update
         
         XcedDoc = ""
         Xhcesi = MsgBox("Desea crear Historia Clínica para este paciente con estos datos?", vbInformation + vbYesNo)
         XcedDoc = InputBox("Ingrese su número de cédula completo (Ejemplo: para CI:1234567-8, ingresar: 12345678)")
         
         If Xhcesi = vbYes And XcedDoc <> "" Then
            data_conshc.RecordSource = "select * from us where documento ='" & Trim(XcedDoc) & "'"
            data_conshc.Refresh
            If data_conshc.Recordset.RecordCount > 0 Then
               data_hce.RecordSource = "select * from cabezal_hc where cb_mat =" & frm_pendlla.Data6.Recordset("matricula")
               data_hce.Refresh
               If data_hce.Recordset.RecordCount > 0 Then
                  data_hce.RecordSource = "select * from cabezal_hcdig where mat =" & frm_pendlla.Data6.Recordset("matricula")
                  data_hce.Refresh
                  data_par.Recordset.Edit
                  data_par.Recordset("p_hc") = data_par.Recordset("p_hc") + 1
                  data_par.Recordset.Update
                  data_hce.Recordset.AddNew
                  data_hce.Recordset("id") = data_par.Recordset("p_hc")
                  data_hce.Recordset("hc_nro") = data_par.Recordset("p_hc")
                  data_hce.Recordset("mat") = frm_pendlla.Data6.Recordset("matricula")
                  data_hce.Recordset("cednum") = frm_pendlla.Data6.Recordset("cedula")
                  data_hce.Recordset("cedtext") = Trim(str(frm_pendlla.Data6.Recordset("cedula"))) & Trim(str(frm_pendlla.Data6.Recordset("codced")))
                  data_hce.Recordset("codced") = frm_pendlla.Data6.Recordset("codced")
                  data_hce.Recordset("fecha") = Format(Date, "dd-mm-yyyy")
                  data_hce.Recordset("hora") = Format(Time, "HH:mm:ss")
                  data_hce.Recordset("codigo") = 3
                  data_hce.Recordset("tipo_cons") = 9
                  data_hce.Recordset("tipo_consd") = "Orientación Telefónica"
                  data_hce.Recordset("hc_base") = frm_menu.data_parse.Recordset("base")
                  data_hce.Recordset("hc_codmed") = data_conshc.Recordset("id")
                  data_hce.Recordset("hc_nommed") = data_conshc.Recordset("nombre") & " " & data_conshc.Recordset("apellidos")
                  data_hce.Recordset("hc_cpmed") = data_conshc.Recordset("cp")
                  
                  data_hce.Recordset.Update
      
                  data_hce.RecordSource = "Select * from hc_mcyotro where id =" & 529
                  data_hce.Refresh
                  data_hce.Recordset.AddNew
                  data_hce.Recordset("id") = data_par.Recordset("p_hc")
                  data_hce.Recordset("hc_nro") = data_par.Recordset("p_hc")
                  data_hce.Recordset("hc_mat") = frm_pendlla.Data6.Recordset("matricula")
                  data_hce.Recordset("fecha") = Format(Date, "dd-mm-yyyy")
                  data_hce.Recordset("hora") = Format(Time, "HH:mm")
                  data_hce.Recordset("hc_mc") = "Orientación telefónica"
                  data_hce.Recordset("hc_otros") = Text1.Text
                  data_hce.Recordset.Update

                  data_hce.RecordSource = "Select * from cli_crmdeudas where nrofact =" & data_par.Recordset("p_hc")
                  data_hce.Refresh
                  data_hce.Recordset.AddNew
                  data_hce.Recordset("id") = data_par.Recordset("p_hc")
                  data_hce.Recordset("base") = frm_pendlla.Data6.Recordset("matricula")
                  data_hce.Recordset("nrofact") = data_par.Recordset("p_hc")
                  data_hce.Recordset("obs") = "registro de orientación clínica por vía telefónica"
                  data_hce.Recordset("usuario") = "Z719"
                  data_hce.Recordset("forma_pago") = 1
                  data_hce.Recordset("var1n") = 3
                  data_hce.Recordset.Update
   
                  data_hce.RecordSource = "Select * from cabezal_hcdig where id =" & data_par.Recordset("p_hc") & " and mat =" & frm_pendlla.Data6.Recordset("matricula")
                  data_hce.Refresh
                  If data_hce.Recordset.RecordCount > 0 Then
                     If IsNull(data_hce.Recordset("hc_fin")) = True Then
                        data_hce.Recordset.Edit
                        data_hce.Recordset("hc_fin") = 5
                        data_hce.Recordset.Update
                     End If
                  End If
          
                  MsgBox "HC creada correctamente", vbInformation
               Else
                  MsgBox "No se encuentra registro de ficha en HCE, deberá crearla manualmente.", vbInformation
               End If
            Else
               MsgBox "No se encuentra número de cédula del médico.", vbInformation
            End If
         End If
         
         MsgBox "Datos de CMT grabados correctamente.", vbInformation
         
         Unload Me
      Else
         MsgBox "Ingrese texto."
      End If
   Else
      MsgBox "Ingrese hora"
   End If
Else
   MsgBox "Ingrese fecha"
End If

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
data_hce.Connect = "odbc;dsn=sappnew;"
data_conshc.Connect = "odbc;dsn=sappnew;"
data_medsapp.Connect = "odbc;dsn=sappnew;"

data_par.Connect = "odbc;dsn=sappnew;"
data_par.RecordSource = "param_gral"
data_par.Refresh

Data1.Connect = "odbc;dsn=sappnew;"
Data1.RecordSource = "select * from sol_hisopos where id =" & frm_pendlla.Data6.Recordset("id")
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   labid.Caption = Data1.Recordset("id")
   Label1.Caption = Data1.Recordset("nombre")
   If IsNull(Data1.Recordset("fecha_cierre")) = False Then
      mfec.Text = Format(Data1.Recordset("fecha_cierre"), "dd/mm/yyyy")
   Else
      mfec.Text = Format(Date, "dd/mm/yyyy")
   End If
   If IsNull(Data1.Recordset("hora_cierre")) = False Then
      mhora.Text = Format(Data1.Recordset("hora_cierre"), "HH:mm")
   Else
      mhora.Text = Format(Time, "HH:mm")
   End If
   If IsNull(Data1.Recordset("obs_cmt")) = False Then
      Text1.Text = Data1.Recordset("obs_cmt")
   Else
      Text1.Text = ""
   End If
Else
   MsgBox "Error al seleccionar. Cierre la ventana y vuelva a ingresar.", vbCritical
   
End If

End Sub
