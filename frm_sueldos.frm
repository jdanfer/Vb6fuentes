VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_sueldos 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sueldos al BROU"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   7455
   Icon            =   "frm_sueldos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   7455
   StartUpPosition =   1  'CenterOwner
   Begin Crystal.CrystalReport cr1 
      Left            =   6720
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Data data_inf 
      Caption         =   "data_inf"
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
      Top             =   600
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "tarjbrou"
      Top             =   2760
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "brou"
      Top             =   840
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton bcarga 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cargar Datos"
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
      Left            =   240
      MouseIcon       =   "frm_sueldos.frx":0442
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   720
      Width           =   1935
   End
   Begin VB.CommandButton bsalir 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Terminar"
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
      Left            =   5040
      Picture         =   "frm_sueldos.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4560
      Width           =   1935
   End
   Begin VB.CommandButton bproc 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Procesar disquete"
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
      Left            =   360
      Picture         =   "frm_sueldos.frx":0CD6
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4560
      Width           =   1815
   End
   Begin VB.CommandButton bbus 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   3720
      Picture         =   "frm_sueldos.frx":1260
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Buscar"
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton bcan 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2880
      Picture         =   "frm_sueldos.frx":17EA
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Cancelar acción"
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton bgrab 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2040
      Picture         =   "frm_sueldos.frx":1D74
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Guardar datos"
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton bmod 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1200
      Picture         =   "frm_sueldos.frx":22FE
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Modificar registro"
      Top             =   3840
      Width           =   495
   End
   Begin VB.CommandButton bnue 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   360
      Picture         =   "frm_sueldos.frx":2888
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Nuevo registro"
      Top             =   3840
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Datos de sueldos"
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
      ForeColor       =   &H00FF0000&
      Height          =   2055
      Left            =   360
      TabIndex        =   5
      Top             =   1440
      Width           =   6735
      Begin VB.TextBox timp 
         Alignment       =   1  'Right Justify
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   14346
            SubFormatType   =   1
         EndProperty
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
         Left            =   1920
         TabIndex        =   17
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txt_cod 
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
         Left            =   3480
         MaxLength       =   1
         TabIndex        =   8
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox txt_ced 
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
         Height          =   375
         Left            =   1920
         MaxLength       =   7
         TabIndex        =   7
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C00000&
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
         Left            =   240
         TabIndex        =   19
         Top             =   960
         Width           =   5415
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C00000&
         Caption         =   "IMPORTE:"
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
         Left            =   240
         TabIndex        =   9
         Top             =   1440
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C00000&
         Caption         =   "CEDULA:"
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
         Left            =   240
         TabIndex        =   6
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.TextBox txt_a 
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
      Left            =   6240
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox txt_m 
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
      Left            =   5640
      MaxLength       =   2
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
   Begin MSMask.MaskEdBox mfec 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      _Version        =   393216
      ForeColor       =   255
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
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
   Begin VB.Line Line3 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   0
      X2              =   7440
      Y1              =   3720
      Y2              =   3720
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   0
      X2              =   7440
      Y1              =   4440
      Y2              =   4440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      X1              =   0
      X2              =   7440
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "MES/AÑO:"
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
      Left            =   4200
      TabIndex        =   2
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "FECHA CREDITO:"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Image Image1 
      Height          =   1455
      Left            =   4680
      Picture         =   "frm_sueldos.frx":2E12
      Stretch         =   -1  'True
      Top             =   3840
      Width           =   1935
   End
End
Attribute VB_Name = "frm_sueldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub bbus_Click()
frm_bussue.Show vbModal

End Sub

Private Sub bcan_Click()
If XAlta = 1 Then
   XAlta = 0
   Frame1.Enabled = False
   bgrab.Enabled = False
   bcan.Enabled = False
   bmod.Enabled = True
   bbus.Enabled = True
   bnue.Enabled = True
   borrasu
Else
   XAlta = 0
   Frame1.Enabled = False
   bgrab.Enabled = False
   bcan.Enabled = False
   bmod.Enabled = True
   bbus.Enabled = True
   bnue.Enabled = True
   borrasu
End If

End Sub

Private Sub bcarga_Click()
If mfec.Text <> "__/__/____" Then
   If txt_m.Text <> "" Then
      If txt_a.Text <> "" Then
         If txt_m.Text <> 0 Then
            If txt_a.Text <> 0 Then
               data1.RecordSource = "select * from brou where fecha =#" & Format(mfec.Text, "yyyy/mm/dd") & "# order by ced"
               data1.Refresh
               bcarga.Enabled = False
               mfec.Enabled = False
               txt_m.Enabled = False
               txt_a.Enabled = False
               bnue.Enabled = True
               bmod.Enabled = True
               bbus.Enabled = True
               bnue.SetFocus
            End If
         End If
      End If
   End If
End If

End Sub

Private Sub bgrab_Click()
If txt_ced.Text <> "" Then
   If timp.Text <> "" Then
      If XAlta = 1 Then
         data1.Recordset.FindFirst "ced =" & txt_ced.Text
         If data1.Recordset.NoMatch Then
            data1.Recordset.AddNew
            data1.Recordset("fecha") = Format(mfec.Text, "dd/mm/yyyy")
            data1.Recordset("ced") = txt_ced.Text
            data1.Recordset("codver") = txt_cod.Text
            data1.Recordset("mes") = txt_m.Text
            data1.Recordset("ano") = txt_a.Text
            data1.Recordset("importe") = timp.Text
            data1.Recordset("comis") = 0
            data1.Recordset("nom1") = data2.Recordset("nom1")
            data1.Recordset("nom2") = data2.Recordset("nom2")
            data1.Recordset("ape1") = data2.Recordset("ape1")
            data1.Recordset("ape2") = data2.Recordset("ape2")
            data1.Recordset("total") = 0
            data1.Recordset.Update
            XAlta = 0
            borrasu
            Frame1.Enabled = False
            bgrab.Enabled = False
            bcan.Enabled = False
            bmod.Enabled = True
            bbus.Enabled = True
            bnue.Enabled = True
            bnue.SetFocus
         Else
            MsgBox "Ya existe, VERIFIQUE!!", vbCritical, "Mensaje"
            txt_ced.SetFocus
         End If
      Else
         data2.Recordset.Edit
         data1.Recordset("fecha") = Format(mfec.Text, "dd/mm/yyyy")
         data1.Recordset("ced") = txt_ced.Text
         data1.Recordset("codver") = txt_cod.Text
         data1.Recordset("mes") = txt_m.Text
         data1.Recordset("ano") = txt_a.Text
         data1.Recordset("importe") = timp.Text
         data1.Recordset("comis") = 0
         data1.Recordset("nom1") = data2.Recordset("nom1")
         data1.Recordset("nom2") = data2.Recordset("nom2")
         data1.Recordset("ape1") = data2.Recordset("ape1")
         data1.Recordset("ape2") = data2.Recordset("ape2")
         data1.Recordset("total") = 0
         data1.Recordset.Update
         XAlta = 0
         borrasu
         Frame1.Enabled = False
         bgrab.Enabled = False
         bcan.Enabled = False
         bmod.Enabled = True
         bbus.Enabled = True
         bnue.Enabled = True
         bnue.SetFocus
      End If
   End If
End If

End Sub

Private Sub bmod_Click()
XAlta = 0
Frame1.Enabled = True
txt_ced.SetFocus
bgrab.Enabled = True
bcan.Enabled = True
bmod.Enabled = False
bbus.Enabled = False
bnue.Enabled = False

End Sub

Private Sub bnue_Click()
XAlta = 1
Frame1.Enabled = True
txt_ced.SetFocus
bgrab.Enabled = True
bcan.Enabled = True
bmod.Enabled = False
bbus.Enabled = False
bnue.Enabled = False
borrasu

End Sub

Private Sub bproc_Click()
Dim Xcadsue As String
Dim Xelimp As Long
Dim Xtotreg, Xtotimpp As Long
frm_sueldos.MousePointer = 11
bproc.Enabled = False

Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.Path & "\informes.mdb")

MiBaseact.Execute "Delete * from infcli"

data_inf.RecordSource = "infcli"
data_inf.Refresh

If mfec.Text <> "__/__/____" Then
   data1.RecordSource = "Select * from brou where fecha =#" & Format(mfec.Text, "yyyy/mm/dd") & "# and importe >" & 0 & " order by ced"
   data1.Refresh
   If data1.Recordset.RecordCount > 0 Then
      Open App.Path & "\sueldos.txt" For Output As #1
      Do While Not data1.Recordset.EOF
        Xcadsue = "1 001SJ"
        Xcadsue = Xcadsue + Mid(Trim(Str(Year(mfec.Text))), 3, 2)
        If Month(mfec.Text) < 10 Then
           Xcadsue = Xcadsue + "0" + Trim(Str(Month(mfec.Text)))
        Else
           Xcadsue = Xcadsue + Trim(Str(Month(mfec.Text)))
        End If
        If Day(mfec.Text) < 10 Then
           Xcadsue = Xcadsue + "0" + Trim(Str(Day(mfec.Text)))
        Else
           Xcadsue = Xcadsue + Trim(Str(Day(mfec.Text)))
        End If
        If Len(Trim(Str(data1.Recordset("ced")))) > 7 Then
           MsgBox "Atención: error en número de cédula, verifique : " + Trim(Str(data1.Recordset("ced"))), vbCritical, "Mensaje"
        End If
        If Len(Trim(Str(data1.Recordset("ced")))) = 7 Then
           Xcadsue = Xcadsue + Trim(Str(data1.Recordset("ced"))) + Trim(Str(data1.Recordset("codver"))) + "0000000"
        End If
        If Len(Trim(Str(data1.Recordset("ced")))) = 6 Then
           Xcadsue = Xcadsue + "0" + Trim(Str(data1.Recordset("ced"))) + Trim(Str(data1.Recordset("codver"))) + "0000000"
        End If
        If Len(Trim(Str(data1.Recordset("ced")))) = 5 Then
           Xcadsue = Xcadsue + "00" + Trim(Str(data1.Recordset("ced"))) + Trim(Str(data1.Recordset("codver"))) + "0000000"
        End If
        Xcadsue = Xcadsue + "00038598A00000000000"
        Xcadsue = Xcadsue + Mid(Trim(Str(Year(mfec.Text))), 3, 2)
        If Month(mfec.Text) < 10 Then
           Xcadsue = Xcadsue + "0" + Trim(Str(Month(mfec.Text)))
        Else
           Xcadsue = Xcadsue + Trim(Str(Month(mfec.Text)))
        End If
        Xelimp = data1.Recordset("importe")
        If Len(Trim(Str(Int(Xelimp)))) = 1 Then
           Xcadsue = Xcadsue + "000000000000" + Trim(Str(Int(Xelimp))) + "00"
        End If
        If Len(Trim(Str(Int(Xelimp)))) = 2 Then
           Xcadsue = Xcadsue + "00000000000" + Trim(Str(Int(Xelimp))) + "00"
        End If
        If Len(Trim(Str(Int(Xelimp)))) = 3 Then
           Xcadsue = Xcadsue + "0000000000" + Trim(Str(Int(Xelimp))) + "00"
        End If
        If Len(Trim(Str(Int(Xelimp)))) = 4 Then
           Xcadsue = Xcadsue + "000000000" + Trim(Str(Int(Xelimp))) + "00"
        End If
        If Len(Trim(Str(Int(Xelimp)))) = 5 Then
           Xcadsue = Xcadsue + "00000000" + Trim(Str(Int(Xelimp))) + "00"
        End If
        If Len(Trim(Str(Int(Xelimp)))) = 6 Then
           Xcadsue = Xcadsue + "0000000" + Trim(Str(Int(Xelimp))) + "00"
        End If
        If Len(Trim(Str(Int(Xelimp)))) = 7 Then
           Xcadsue = Xcadsue + "000000" + Trim(Str(Int(Xelimp))) + "00"
        End If
        Xtotreg = Xtotreg + 1
        Xtotimpp = Xtotimpp + Xelimp
        Xcadsue = Xcadsue + "0000000000000"
        Xcadsue = Xcadsue + "                        "
        Xcadsue = Xcadsue + "                        "
        Print #1, Xcadsue
        data_inf.Recordset.AddNew
        data_inf.Recordset("cl_apellid") = data1.Recordset("ape1")
        data_inf.Recordset("cl_nombre") = data1.Recordset("nom1")
        data_inf.Recordset("cl_cedula") = data1.Recordset("ced")
        data_inf.Recordset("saldo_cc") = Xelimp
        data_inf.Recordset.Update
        
        data1.Recordset.MoveNext
      Loop
      Xcadsue = "20001SJ"
      Xcadsue = Xcadsue + Mid(Trim(Str(Year(mfec.Text))), 3, 2)
      If Month(mfec.Text) < 10 Then
         Xcadsue = Xcadsue + "0" + Trim(Str(Month(mfec.Text)))
      Else
         Xcadsue = Xcadsue + Trim(Str(Month(mfec.Text)))
      End If
      If Day(mfec.Text) < 10 Then
         Xcadsue = Xcadsue + "0" + Trim(Str(Day(mfec.Text)))
      Else
         Xcadsue = Xcadsue + Trim(Str(Day(mfec.Text)))
      End If
      If Len(Trim(Str(Int(Xtotreg)))) = 1 Then
         Xcadsue = Xcadsue + "00000" + Trim(Str(Int(Xtotreg)))
      End If
      If Len(Trim(Str(Int(Xtotreg)))) = 2 Then
         Xcadsue = Xcadsue + "0000" + Trim(Str(Int(Xtotreg)))
      End If
      If Len(Trim(Str(Int(Xtotreg)))) = 3 Then
         Xcadsue = Xcadsue + "000" + Trim(Str(Int(Xtotreg)))
      End If
      If Len(Trim(Str(Int(Xtotimpp)))) = 1 Then
         Xcadsue = Xcadsue + "000000000000000" + Trim(Str(Int(Xtotimpp))) + "00"
      End If
      If Len(Trim(Str(Int(Xtotimpp)))) = 2 Then
         Xcadsue = Xcadsue + "00000000000000" + Trim(Str(Int(Xtotimpp))) + "00"
      End If
      If Len(Trim(Str(Int(Xtotimpp)))) = 3 Then
         Xcadsue = Xcadsue + "0000000000000" + Trim(Str(Int(Xtotimpp))) + "00"
      End If
      If Len(Trim(Str(Int(Xtotimpp)))) = 4 Then
         Xcadsue = Xcadsue + "000000000000" + Trim(Str(Int(Xtotimpp))) + "00"
      End If
      If Len(Trim(Str(Int(Xtotimpp)))) = 5 Then
         Xcadsue = Xcadsue + "00000000000" + Trim(Str(Int(Xtotimpp))) + "00"
      End If
      If Len(Trim(Str(Int(Xtotimpp)))) = 6 Then
         Xcadsue = Xcadsue + "0000000000" + Trim(Str(Int(Xtotimpp))) + "00"
      End If
      If Len(Trim(Str(Int(Xtotimpp)))) = 7 Then
         Xcadsue = Xcadsue + "000000000" + Trim(Str(Int(Xtotimpp))) + "00"
      End If
      If Len(Trim(Str(Int(Xtotimpp)))) = 8 Then
         Xcadsue = Xcadsue + "00000000" + Trim(Str(Int(Xtotimpp))) + "00"
      End If
      Xcadsue = Xcadsue + "000000"
      Xcadsue = Xcadsue + "000000000000000000"
      Xcadsue = Xcadsue + "000000"
      Xcadsue = Xcadsue + "000000000000000000"
      Xcadsue = Xcadsue + "0000000000000000"
      Xcadsue = Xcadsue + "                           "
      Print #1, Xcadsue;
   End If
   Close #1
   MsgBox "EL ARCHIVO SE GUARDO EN C:\SAPPMYS\SAPPMYSQL\ARCHIVOS\", vbInformation, "Proceso de sueldos"
   FileCopy App.Path & "\sueldos.txt", "c:\sappmys\sappmysql\archivos\sueldos.txt"
   cr1.ReportFileName = App.Path & "\infsueldos.rpt"
   cr1.ReportTitle = "FECHA DE PROCESO DE SUELDOS:  " & mfec.Text
   cr1.Action = 1
   
End If
frm_sueldos.MousePointer = 0
bproc.Enabled = True


End Sub

Private Sub bsalir_Click()
Unload Me

End Sub

Private Sub Form_Load()
data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
data2.Connect = "odbc;dsn=" & Xconexrmt & ";"
data_inf.DatabaseName = App.Path & "\informes.mdb"
data_inf.RecordSource = "infcli"
data_inf.Refresh

End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub mfec_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_m.SetFocus
End If

End Sub

Private Sub timp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   bgrab.SetFocus
End If

End Sub

Private Sub timp_LostFocus()
If timp.Text <> "" Then
   timp.Text = Round(timp.Text)
   timp.Text = Format(timp.Text, "Standard")
End If

End Sub

Private Sub txt_a_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   bcarga.SetFocus
End If

End Sub

Private Sub txt_ced_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   timp.SetFocus
End If

End Sub

Private Sub txt_ced_LostFocus()
If txt_ced.Text <> "" Then
   data2.Recordset.FindFirst "cedula =" & txt_ced.Text
   If Not data2.Recordset.NoMatch Then
      txt_cod.Text = data2.Recordset("codver")
      Label5.Caption = data2.Recordset("nom1")
      Label5.Caption = Label5.Caption + " " + data2.Recordset("ape1")
   Else
      MsgBox "No se encontró funcionario", vbCritical, "Mensaje"
      txt_ced.SetFocus
   End If
End If

End Sub

Private Sub txt_m_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_a.SetFocus
End If

End Sub

Public Function borrasu()
txt_ced.Text = ""
txt_cod.Text = ""
Label5.Caption = ""
timp.Text = ""

End Function
