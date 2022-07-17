VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infencu 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de encuestas"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5520
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_infencu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   5520
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   3240
      Top             =   3720
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "data1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   2040
      Top             =   1680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4560
      Picture         =   "frm_infencu.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Salir"
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      Picture         =   "frm_infencu.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Procesar"
      Top             =   3600
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Caption         =   "Datos de informe"
      Height          =   3375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4815
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C00000&
         Caption         =   "Especialistas"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   3000
         Width           =   2055
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Consultar encuestas ingresadas"
         Height          =   495
         Left            =   240
         Picture         =   "frm_infencu.frx":0F56
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1920
         Width           =   4335
      End
      Begin VB.Data data_inf 
         Caption         =   "data_inf"
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
         Top             =   1440
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C00000&
         Caption         =   "Solo Policlínica"
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   240
         TabIndex        =   9
         Top             =   2640
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C00000&
         Caption         =   "Solo Domicilio"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2520
         TabIndex        =   8
         Top             =   2640
         Width           =   2055
      End
      Begin VB.ComboBox Combo1 
         Height          =   360
         ItemData        =   "frm_infencu.frx":14E0
         Left            =   1560
         List            =   "frm_infencu.frx":14F6
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   3015
      End
      Begin MSMask.MaskEdBox mfh 
         Height          =   375
         Left            =   3120
         TabIndex        =   3
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mfd 
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C00000&
         Caption         =   "Encuesta:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C00000&
         Caption         =   "FECHAS:"
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Image Image1 
      Height          =   735
      Left            =   1080
      Picture         =   "frm_infencu.frx":1545
      Stretch         =   -1  'True
      Top             =   3120
      Width           =   1695
   End
End
Attribute VB_Name = "frm_infencu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
   If Check2.Value = 1 Then
      Check2.Value = 0
   End If
End If

End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
   If Check1.Value = 1 Then
      Check1.Value = 0
   End If
End If

End Sub

Private Sub Command1_Click()
Dim Xo, Xtotp7 As Integer
Dim Xresul, Xtotr As Double
Dim Xres1, Xres2, Xres3, Xres4, Xres5, Xres6, Xres7, Xtotgp As Long
Xresul = 0
If Combo1.ListIndex = 2 Then
   Xo = 6
Else
   Xo = 6
End If
Xtotp7 = 0

Command1.Enabled = False

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infcli"
data_inf.RecordSource = "infcli"
data_inf.Refresh

If mfd.Text = "__/__/____" Then
   MsgBox "Ingrese FECHA"
Else
   If Combo1.ListIndex = 0 Then
      Data1.RecordSource = "Select * from rrhh_sol where cl_fnac >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mfh.Text, "yyyy-mm-dd") & "' order by cl_fnac"
      Data1.Refresh
   Else
      If Combo1.ListIndex = 1 Then
         Data1.RecordSource = "Select * from rrhh_sol where cl_fnac >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cl_atrasoa =" & 0 & " order by cl_fnac"
         Data1.Refresh
      Else
         If Combo1.ListIndex = 2 Then
            Data1.RecordSource = "Select * from rrhh_sol where cl_fnac >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cl_atrasoa =" & 1 & " order by cl_fnac"
            Data1.Refresh
         Else
            If Combo1.ListIndex = 3 Then
               Data1.RecordSource = "Select * from rrhh_sol where cl_fnac >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cl_atrasoa =" & 2 & " order by cl_fnac"
               Data1.Refresh
            Else
               If Combo1.ListIndex = 4 Then
                  If Check2.Value = 1 Then
                     Data1.RecordSource = "Select * from rrhh_sol where cl_fnac >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cl_atrasoa =" & 0 & " order by cl_fnac"
                     Data1.Refresh
                  Else
                     If Check1.Value = 1 Then
                        Data1.RecordSource = "Select * from rrhh_sol where cl_fnac >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cl_atrasoa =" & 1 & " order by cl_fnac"
                        Data1.Refresh
                     Else
                        If Check3.Value = 1 Then
                           Data1.RecordSource = "Select * from rrhh_sol where cl_fnac >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cl_atrasoa =" & 3 & " order by cl_fnac"
                           Data1.Refresh
                        Else
                           Data1.RecordSource = "Select * from rrhh_sol where cl_fnac >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cl_atrasoa in (0,1,3) order by cl_fnac"
                           Data1.Refresh
                        End If
                     End If
                  End If
               Else
                  If Combo1.ListIndex = 5 Then
                     Data1.RecordSource = "Select * from rrhh_sol where cl_fnac >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cl_atrasoa =" & 3 & " order by cl_fnac"
                     Data1.Refresh
                  Else
                     If Check2.Value = 1 Then
                        Data1.RecordSource = "Select * from rrhh_sol where cl_fnac >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cl_atrasoa =" & 0 & " order by cl_fnac"
                        Data1.Refresh
                     Else
                        If Check1.Value = 1 Then
                           Data1.RecordSource = "Select * from rrhh_sol where cl_fnac >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mfh.Text, "yyyy-mm-dd") & "' and cl_atrasoa =" & 1 & " order by cl_fnac"
                           Data1.Refresh
                        Else
                           Data1.RecordSource = "Select * from rrhh_sol where cl_fnac >='" & Format(mfd.Text, "yyyy-mm-dd") & "' and cl_fnac <='" & Format(mfh.Text, "yyyy-mm-dd") & "' order by cl_fnac"
                           Data1.Refresh
                        End If
                     End If
                  End If
               End If
            End If
         End If
      End If
   End If
   If Data1.Recordset.RecordCount > 0 Then
      Data1.Recordset.MoveFirst
      If Combo1.ListIndex = 4 Then
        data_inf.Recordset.AddNew
        data_inf.Recordset("cl_codigo") = 1
        data_inf.Recordset("cl_direcci") = "Pregunta Uno"
        data_inf.Recordset.Update
        data_inf.Recordset.AddNew
        data_inf.Recordset("cl_codigo") = 2
        data_inf.Recordset("cl_direcci") = "Pregunta Dos"
        data_inf.Recordset.Update
        data_inf.Recordset.AddNew
        data_inf.Recordset("cl_codigo") = 3
        data_inf.Recordset("cl_direcci") = "Pregunta Tres"
        data_inf.Recordset.Update
        data_inf.Recordset.AddNew
        data_inf.Recordset("cl_codigo") = 4
        data_inf.Recordset("cl_direcci") = "Pregunta Cuatro"
        data_inf.Recordset.Update
        data_inf.Recordset.AddNew
        data_inf.Recordset("cl_codigo") = 5
        data_inf.Recordset("cl_direcci") = "Pregunta Cinco"
        data_inf.Recordset.Update
        data_inf.Recordset.AddNew
        data_inf.Recordset("cl_codigo") = 6
        data_inf.Recordset("cl_direcci") = "Pregunta Seis"
        data_inf.Recordset.Update
        data_inf.Recordset.AddNew
        data_inf.Recordset("cl_codigo") = 7
        data_inf.Recordset("cl_direcci") = "Pregunta Siete"
        data_inf.Recordset.Update
      End If
      Do While Not Data1.Recordset.EOF
         If Combo1.ListIndex = 4 Then
            Xtotgp = Xtotgp + 1
'1
            If Data1.Recordset("cl_numero") = 0 Then
               Xres1 = Xres1 + 5
            End If
            If Data1.Recordset("cl_numero") = 1 Then
               Xres1 = Xres1 + 4
            End If
            If Data1.Recordset("cl_numero") = 2 Then
               Xres1 = Xres1 + 3
            End If
            If Data1.Recordset("cl_numero") = 3 Then
               Xres1 = Xres1 + 2
            End If
            If Data1.Recordset("cl_numero") = 4 Then
               Xres1 = Xres1 + 1
            End If
'2
            If Data1.Recordset("cl_zona") = 0 Then
               Xres2 = Xres2 + 5
            End If
            If Data1.Recordset("cl_zona") = 1 Then
               Xres2 = Xres2 + 4
            End If
            If Data1.Recordset("cl_zona") = 2 Then
               Xres2 = Xres2 + 3
            End If
            If Data1.Recordset("cl_zona") = 3 Then
               Xres2 = Xres2 + 2
            End If
            If Data1.Recordset("cl_zona") = 4 Then
               Xres2 = Xres2 + 1
            End If
'3
            If Data1.Recordset("cl_nomcobr") = 0 Then
               Xres3 = Xres3 + 5
            End If
            If Data1.Recordset("cl_nomcobr") = 1 Then
               Xres3 = Xres3 + 4
            End If
            If Data1.Recordset("cl_nomcobr") = 2 Then
               Xres3 = Xres3 + 3
            End If
            If Data1.Recordset("cl_nomcobr") = 3 Then
               Xres3 = Xres3 + 2
            End If
            If Data1.Recordset("cl_nomcobr") = 4 Then
               Xres3 = Xres3 + 1
            End If
'4
            If Data1.Recordset("cl_val1") = 0 Then
               Xres4 = Xres4 + 5
            End If
            If Data1.Recordset("cl_val1") = 1 Then
               Xres4 = Xres4 + 4
            End If
            If Data1.Recordset("cl_val1") = 2 Then
               Xres4 = Xres4 + 3
            End If
            If Data1.Recordset("cl_val1") = 3 Then
               Xres4 = Xres4 + 2
            End If
            If Data1.Recordset("cl_val1") = 4 Then
               Xres4 = Xres4 + 1
            End If
'5
            If Data1.Recordset("cl_val2") = 0 Then
               Xres5 = Xres5 + 5
            End If
            If Data1.Recordset("cl_val2") = 1 Then
               Xres5 = Xres5 + 4
            End If
            If Data1.Recordset("cl_val2") = 2 Then
               Xres5 = Xres5 + 3
            End If
            If Data1.Recordset("cl_val2") = 3 Then
               Xres5 = Xres5 + 2
            End If
            If Data1.Recordset("cl_val2") = 4 Then
               Xres5 = Xres5 + 1
            End If
'6
            If Data1.Recordset("cl_val3") = 0 Then
               Xres6 = Xres6 + 5
            End If
            If Data1.Recordset("cl_val3") = 1 Then
               Xres6 = Xres6 + 4
            End If
            If Data1.Recordset("cl_val3") = 2 Then
               Xres6 = Xres6 + 3
            End If
            If Data1.Recordset("cl_val3") = 3 Then
               Xres6 = Xres6 + 2
            End If
            If Data1.Recordset("cl_val3") = 4 Then
               Xres6 = Xres6 + 1
            End If
         
            If IsNull(Data1.Recordset("cl_etiquet")) = False Then
               If Data1.Recordset("cl_etiquet") = 0 Then
                  Xres7 = Xres7 + 5
                  Xtotp7 = Xtotp7 + 1
               End If
               If Data1.Recordset("cl_etiquet") = 1 Then
                  Xres7 = Xres7 + 4
                  Xtotp7 = Xtotp7 + 1
               End If
               If Data1.Recordset("cl_etiquet") = 2 Then
                  Xres7 = Xres7 + 3
                  Xtotp7 = Xtotp7 + 1
               End If
               If Data1.Recordset("cl_etiquet") = 3 Then
                  Xres7 = Xres7 + 2
                  Xtotp7 = Xtotp7 + 1
               End If
               If Data1.Recordset("cl_etiquet") = 4 Then
                  Xres7 = Xres7 + 1
                  Xtotp7 = Xtotp7 + 1
               End If
            
            Else
               Xres7 = 0
            End If
            Data1.Recordset.MoveNext
         Else
            data_inf.Recordset.AddNew
            data_inf.Recordset("cl_codigo") = Data1.Recordset("cl_nrovend")
            data_inf.Recordset("cl_apellid") = Data1.Recordset("cl_desc1")
            data_inf.Recordset("cl_fnac") = Data1.Recordset("cl_fnac")
            data_inf.Recordset("cl_nombre") = Data1.Recordset("cl_descpag")
            data_inf.Recordset("cl_direcci") = Data1.Recordset("cl_desc2")
            If Data1.Recordset("cl_numero") = 0 Then
               Xresul = Xresul + 5
            End If
            If Data1.Recordset("cl_numero") = 1 Then
               Xresul = Xresul + 4
            End If
            If Data1.Recordset("cl_numero") = 2 Then
               Xresul = Xresul + 3
            End If
            If Data1.Recordset("cl_numero") = 3 Then
               Xresul = Xresul + 2
            End If
            If Data1.Recordset("cl_numero") = 4 Then
               Xresul = Xresul + 1
            End If
            
            If Data1.Recordset("cl_zona") = 0 Then
               Xresul = Xresul + 5
            End If
            If Data1.Recordset("cl_zona") = 1 Then
               Xresul = Xresul + 4
            End If
            If Data1.Recordset("cl_zona") = 2 Then
               Xresul = Xresul + 3
            End If
            If Data1.Recordset("cl_zona") = 3 Then
               Xresul = Xresul + 2
            End If
            If Data1.Recordset("cl_zona") = 4 Then
               Xresul = Xresul + 1
            End If
            
            If Data1.Recordset("cl_nomcobr") = 0 Then
               Xresul = Xresul + 5
            End If
            If Data1.Recordset("cl_nomcobr") = 1 Then
               Xresul = Xresul + 4
            End If
            If Data1.Recordset("cl_nomcobr") = 2 Then
               Xresul = Xresul + 3
            End If
            If Data1.Recordset("cl_nomcobr") = 3 Then
               Xresul = Xresul + 2
            End If
            If Data1.Recordset("cl_nomcobr") = 4 Then
               Xresul = Xresul + 1
            End If
            
            If Data1.Recordset("cl_val1") = 0 Then
               Xresul = Xresul + 5
            End If
            If Data1.Recordset("cl_val1") = 1 Then
               Xresul = Xresul + 4
            End If
            If Data1.Recordset("cl_val1") = 2 Then
               Xresul = Xresul + 3
            End If
            If Data1.Recordset("cl_val1") = 3 Then
               Xresul = Xresul + 2
            End If
            If Data1.Recordset("cl_val1") = 4 Then
               Xresul = Xresul + 1
            End If
            
            If Data1.Recordset("cl_val2") = 0 Then
               Xresul = Xresul + 5
            End If
            If Data1.Recordset("cl_val2") = 1 Then
               Xresul = Xresul + 4
            End If
            If Data1.Recordset("cl_val2") = 2 Then
               Xresul = Xresul + 3
            End If
            If Data1.Recordset("cl_val2") = 3 Then
               Xresul = Xresul + 2
            End If
            If Data1.Recordset("cl_val2") = 4 Then
               Xresul = Xresul + 1
            End If
            
            If Data1.Recordset("cl_val3") = 0 Then
               Xresul = Xresul + 5
            End If
            If Data1.Recordset("cl_val3") = 1 Then
               Xresul = Xresul + 4
            End If
            If Data1.Recordset("cl_val3") = 2 Then
               Xresul = Xresul + 3
            End If
            If Data1.Recordset("cl_val3") = 3 Then
               Xresul = Xresul + 2
            End If
            If Data1.Recordset("cl_val3") = 4 Then
               Xresul = Xresul + 1
            End If
            
            If Xo = 7 Then
               If Data1.Recordset("cl_etiquet") = 0 Then
                  Xresul = Xresul + 5
               End If
               If Data1.Recordset("cl_etiquet") = 1 Then
                  Xresul = Xresul + 4
               End If
               If Data1.Recordset("cl_etiquet") = 2 Then
                  Xresul = Xresul + 3
               End If
               If Data1.Recordset("cl_etiquet") = 3 Then
                  Xresul = Xresul + 2
               End If
               If Data1.Recordset("cl_etiquet") = 4 Then
                  Xresul = Xresul + 1
               End If
               
               Xtotp7 = Xtotp7 + 1
            End If
            Xtotr = Xresul / Xo
            Xtotr = Xtotr
            If Xo = 7 Then
               data_inf.Recordset("cl_codced") = Round(Xtotp7, 0)
            Else
               data_inf.Recordset("cl_codced") = Round(Xtotr, 0)
            End If
            data_inf.Recordset("cl_cedula") = Xtotr
            data_inf.Recordset.Update
            Xresul = 0
            Data1.Recordset.MoveNext
         End If
      Loop
      If Combo1.ListIndex = 4 Then
         data_inf.Recordset.MoveFirst
         Do While Not data_inf.Recordset.EOF
            data_inf.Recordset.Edit
            data_inf.Recordset("cl_cedula") = Xres1
            data_inf.Recordset("cl_codced") = Xtotgp
            data_inf.Recordset.Update
            data_inf.Recordset.MoveNext
            
            data_inf.Recordset.Edit
            data_inf.Recordset("cl_cedula") = Xres2
            data_inf.Recordset("cl_codced") = Xtotgp
            data_inf.Recordset.Update
            data_inf.Recordset.MoveNext
            
            data_inf.Recordset.Edit
            data_inf.Recordset("cl_cedula") = Xres3
            data_inf.Recordset("cl_codced") = Xtotgp
            data_inf.Recordset.Update
            data_inf.Recordset.MoveNext
            
            data_inf.Recordset.Edit
            data_inf.Recordset("cl_cedula") = Xres4
            data_inf.Recordset("cl_codced") = Xtotgp
            data_inf.Recordset.Update
            data_inf.Recordset.MoveNext
            
            data_inf.Recordset.Edit
            data_inf.Recordset("cl_cedula") = Xres5
            data_inf.Recordset("cl_codced") = Xtotgp
            data_inf.Recordset.Update
            data_inf.Recordset.MoveNext
            
            data_inf.Recordset.Edit
            data_inf.Recordset("cl_cedula") = Xres6
            data_inf.Recordset("cl_codced") = Xtotgp
            data_inf.Recordset.Update
            data_inf.Recordset.MoveNext
            
            data_inf.Recordset.Edit
            data_inf.Recordset("cl_cedula") = Xres7
            data_inf.Recordset("cl_codced") = Xtotp7
            data_inf.Recordset.Update
            data_inf.Recordset.MoveNext
            
         Loop
      End If
      data_inf.RecordSource = "Select * from infcli order by cl_fnac"
      data_inf.Refresh
      If Combo1.ListIndex = 4 Then
         cr1.ReportFileName = App.path & "\infencup.rpt"
      Else
         cr1.ReportFileName = App.path & "\infencu.rpt"
      End If
      cr1.DiscardSavedData = True
      If Check1.Value = 1 Then
         cr1.ReportTitle = "ENCUESTAS REALIZADAS EN DOMICILIO" & Combo1.Text & " DESDE: " & mfd.Text
      Else
         If Check2.Value = 1 Then
            cr1.ReportTitle = "ENCUESTAS REALIZADAS EN POLICLÍNICA" & Combo1.Text & " DESDE: " & mfd.Text
         Else
            cr1.ReportTitle = "ENCUESTAS REALIZADAS EN " & Combo1.Text & " DESDE: " & mfd.Text
         End If
      End If
      cr1.Action = 1
      
   Else
      MsgBox "No existen registros"
   End If
End If
Command1.Enabled = True

End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
XAlta = 85
frm_buscaencu.Show vbModal

End Sub

Private Sub Form_Load()
'data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data1.ConnectionString = "dsn=" & Xconexrmt

data_inf.DatabaseName = App.path & "\informes.mdb"
'data_inf.ConnectionString = "provider=Microsoft.jet.oledb.3.51; data Source =" & App.Path & "\informes.mdb"

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub mfd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mfh.SetFocus
End If

End Sub

Private Sub mfh_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Combo1.SetFocus
End If

End Sub
