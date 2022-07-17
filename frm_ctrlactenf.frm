VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_ctrlactenf 
   BackColor       =   &H0080C0FF&
   Caption         =   "Cumplimiento de actos de enfermería"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10935
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_ctrlactenf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   10935
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Actualizar"
      Height          =   375
      Left            =   9120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5040
      Width           =   1575
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_ctrlactenf.frx":0442
      Height          =   2295
      Left            =   120
      OleObjectBlob   =   "frm_ctrlactenf.frx":0456
      TabIndex        =   10
      Top             =   2760
      Width           =   10575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1800
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Data data_est 
      Caption         =   "data_est"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   -240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1800
      Visible         =   0   'False
      Width           =   3015
   End
   Begin MSMask.MaskEdBox mhor 
      Height          =   375
      Left            =   5160
      TabIndex        =   8
      Top             =   1320
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   8
      Format          =   "HH:mm:ss"
      Mask            =   "##:##:##"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Cancelar"
      Height          =   495
      Left            =   6960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2160
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Registrar..."
      Height          =   495
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   1935
   End
   Begin MSMask.MaskEdBox mfec 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   1320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      MaxLength       =   10
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   600
      Width           =   6015
   End
   Begin VB.Label Label7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Doble click para seleccionar."
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   5040
      Width           =   4695
   End
   Begin VB.Label Label6 
      Height          =   375
      Left            =   6240
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   3
      X1              =   0
      X2              =   10920
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "HORA:"
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      Caption         =   "FECHA"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ACTO:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "SOCIO:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   1335
      Left            =   9000
      Picture         =   "frm_ctrlactenf.frx":14CD
      Stretch         =   -1  'True
      Top             =   1200
      Width           =   1335
   End
End
Attribute VB_Name = "frm_ctrlactenf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command1.Enabled = False
frm_ctrlactenf.MousePointer = 11
If mhor.Text <> "__:__:__" Then
   If Format(Data1.Recordset("hora"), "HH:mm") > Format(mhor.Text, "HH:mm") Then
      MsgBox "La hora de cierre es menor que la hora de registro, verifique!", vbInformation
   Else
      Data1.Recordset.Edit
      Data1.Recordset("servicio") = 1
      Data1.Recordset("nom_superv") = mhor.Text
      Data1.Recordset.Update
      Data1.Refresh
   End If
   mhor.Text = "__:__:__"
   mfec.Text = "__/__/____"
   Label8.Caption = ""
Else
   frm_ctrlactenf.MousePointer = 0
   MsgBox "Ingrese hora"
End If
frm_ctrlactenf.MousePointer = 0
Command1.Enabled = True

End Sub

Private Sub Command2_Click()
Unload Me

End Sub



Private Sub Command3_Click()
Dim Xfeccc As Date
Xfeccc = CDate("20/07/2019")
Label2.Caption = frmabm.txt_apellid.Text
Label6.Caption = frmabm.txt_mat.Caption
frm_ctrlactenf.MousePointer = 11
'Data1.DatabaseName = App.Path & "\sapp.mdb"
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
If WElusuario = "MCURBELO" Then
   Data1.RecordSource = "Select * from linmmdd where fecha >=#" & Format(Xfeccc, "yyyy/mm/dd") & "# and cod_prod in (20001,20017,20003,30048,20043,20042,20053,20065,20070,20085,20091,20099,20106,20074,20048,20051,20063,20083,20097) and servicio in(0) order by fecha"
Else
   Data1.RecordSource = "Select * from linmmdd where fecha >=#" & Format(Xfeccc, "yyyy/mm/dd") & "# and cod_prod in (20001,20017,20003,30048,20043,20042,20053,20065,20070,20085,20091,20099,20106,20074,20048,20051,20063,20083,20097) and base =" & frm_menu.data_parse.Recordset("base") & " and servicio in(0) order by fecha"
End If
Data1.Refresh
frm_ctrlactenf.MousePointer = 0

End Sub

Private Sub DBGrid1_DblClick()
If IsNull(Data1.Recordset("nom_prod")) = False Then
   Label8.Caption = Data1.Recordset("nom_prod")
   Label2.Caption = Data1.Recordset("nom_cli")
   If IsNull(Data1.Recordset("servicio")) = False Then
      If Data1.Recordset("servicio") = 1 Then
         MsgBox "Este acto de enfermería YA FUE CERRADO"
         mhor.Text = "__:__:__"
      Else
         mfec.Text = Data1.Recordset("fecha")
         mhor.SetFocus
         mhor.Text = Format(Time, "HH:mm:ss")
      End If
   Else
      mfec.Text = Data1.Recordset("fecha")
      mhor.SetFocus
      mhor.Text = Format(Time, "HH:mm:ss")
   End If
         
Else
   Label8.Caption = ""
   Label2.Caption = ""
End If

End Sub

Private Sub Form_Load()
Dim Xfeccc As Date
Xfeccc = CDate("20/07/2019")
Label2.Caption = frmabm.txt_apellid.Text
Label6.Caption = frmabm.txt_mat.Caption
'Data1.DatabaseName = App.Path & "\sapp.mdb"
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
If WElusuario = "MCURBELO" Then
   Data1.RecordSource = "Select * from linmmdd where fecha >=#" & Format(Xfeccc, "yyyy/mm/dd") & "# and cod_prod in (20001,20017,20003,30048,20043,20042,20053,20065,20070,20085,20091,20099,20106,20074,20048,20051,20063,20083,20097) and servicio in(0) order by fecha"
Else
   Data1.RecordSource = "Select * from linmmdd where fecha >=#" & Format(Xfeccc, "yyyy/mm/dd") & "# and cod_prod in (20001,20017,20003,30048,20043,20042,20053,20065,20070,20085,20091,20099,20106,20074,20048,20051,20063,20083,20097) and base =" & frm_menu.data_parse.Recordset("base") & " and servicio in(0) order by fecha"
End If
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
