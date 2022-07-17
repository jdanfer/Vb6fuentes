VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frm_infconves 
   BackColor       =   &H00FF8080&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe de convenios"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6780
   Icon            =   "frm_infconves.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   330
      Left            =   1200
      Top             =   2760
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
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
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   2640
      Top             =   1320
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   2640
      Top             =   2760
      Visible         =   0   'False
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   873
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
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_infconves.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Procesar informe"
      Top             =   2880
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Datos para el informe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   6255
      Begin VB.CheckBox Check4 
         Caption         =   "NO EMITEN/NO FACTURAN"
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         Top             =   2160
         Width           =   2535
      End
      Begin VB.CheckBox Check3 
         Caption         =   "Agregar socios por convenio"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   2160
         Width           =   2535
      End
      Begin MSAdodcLib.Adodc adocli 
         Height          =   330
         Left            =   240
         Top             =   360
         Visible         =   0   'False
         Width           =   2655
         _ExtentX        =   4683
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
         Caption         =   "adocli"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Solo convenios vencidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   7
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Excluir convenios vencidos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Data Data1 
         Caption         =   "Data1"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   240
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Resumen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3360
         TabIndex        =   4
         Top             =   1680
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Detalle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1680
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frm_infconves.frx":0B14
         Left            =   1680
         List            =   "frm_infconves.frx":0B24
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   4215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FF0000&
         Caption         =   "Informe de:"
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
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   3240
      Picture         =   "frm_infconves.frx":0B79
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   3015
   End
End
Attribute VB_Name = "frm_infconves"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check2.Value = 1 Then
   Check2.Value = 0
End If

End Sub

Private Sub Check2_Click()
If Check1.Value = 1 Then
   Check1.Value = 0
End If

End Sub

Private Sub Check4_Click()
If Check4.Value = 1 Then
   If Combo1.ListIndex = 2 Then
   Else
      MsgBox "Debe seleccionar TODOS los convenios", vbCritical
   End If
End If

End Sub

Private Sub Command1_Click()
Dim MiBaseact As Database
Dim Unasesact As Workspace
Set Unasesact = Workspaces(0)
Set MiBaseact = Unasesact.OpenDatabase(App.path & "\informes.mdb")

MiBaseact.Execute "Delete * from infcli"
Data1.RecordSource = "infcli"
Data1.Refresh
Command1.Enabled = False
frm_infconves.MousePointer = 11
If Combo1.ListIndex = 0 Then
   If Check1.Value = 1 Then
      Adodc1.RecordSource = "Select * from convenio where cnv_hasta >='" & Format(Date, "yyyy/mm/dd") & "' and cnv_emite ='" & "SI" & "' and cnv_fbaja is null"
   Else
      If Check2.Value = 1 Then
         Adodc1.RecordSource = "Select * from convenio where cnv_hasta <'" & Format(Date, "yyyy/mm/dd") & "' and cnv_emite ='" & "SI" & "' and cnv_fbaja is null"
      Else
         Adodc1.RecordSource = "Select * from convenio where cnv_emite ='" & "SI" & "' and cnv_fbaja is null"
      End If
   End If
   
   Adodc1.Refresh
   
   If Adodc1.Recordset.RecordCount > 0 Then
      Adodc1.Recordset.MoveFirst
      Do While Not Adodc1.Recordset.EOF
         If IsNull(Adodc1.Recordset("cnv_colrec")) = False Then
            If Adodc1.Recordset("cnv_colrec") = "M" Or Adodc1.Recordset("cnv_colrec") = "R" Or Adodc1.Recordset("cnv_colrec") = "A" Or Adodc1.Recordset("cnv_colrec") = "V" Then
               If Adodc1.Recordset("cnv_cant_r") = 2 Then 'varios recibos
                  Data1.Recordset.AddNew
                  Data1.Recordset("cl_codconv") = Adodc1.Recordset("cnv_codigo")
                  Data1.Recordset("cl_apellid") = Mid(Adodc1.Recordset("cnv_desc"), 1, 60)
                  Data1.Recordset("cl_dpto") = Adodc1.Recordset("cnv_colrec")
                  Data1.Recordset("cl_fnac") = Adodc1.Recordset("cnv_desde")
                  Data1.Recordset("cl_fultmov") = Adodc1.Recordset("cnv_hasta")
                  Data1.Recordset("cl_ruc") = Adodc1.Recordset("cnv_ruc")
                  If IsNull(Adodc1.Recordset("cnv_fbaja")) = False Then
                     Data1.Recordset("cl_fecing") = Adodc1.Recordset("cnv_fbaja")
                  End If
                  Data1.Recordset("saldo_cc") = Adodc1.Recordset("cnv_precio")
                  If Adodc1.Recordset("cnv_precio") > 0 Then
                     Adodc2.RecordSource = "Select * from cnv_prec where cnv_codigo ='" & Adodc1.Recordset("cnv_codigo") & "' order by cnv_desde DESC"
                     Adodc2.Refresh
                     If Adodc2.Recordset.RecordCount > 1 Then
                        Adodc2.Recordset.MoveFirst
                        Adodc2.Recordset.MoveNext
                        Data1.Recordset("saldo_cc2") = Adodc2.Recordset("precio")
                     End If
                  End If
                  Data1.Recordset("cl_codced") = Adodc1.Recordset("cnv_cant_r")
                  Data1.Recordset("cl_fax") = Adodc1.Recordset("cnv_alta")
                  Data1.Recordset("cl_codigo") = Adodc1.Recordset("cnv_cuenta")
                  Data1.Recordset("cl_nomcobr") = "Varios e-ticket"
                  adocli.RecordSource = "select * from clientes where estado =" & 1 & " and cl_codconv ='" & Adodc1.Recordset("cnv_codigo") & "'"
                  adocli.Refresh
                  If adocli.Recordset.RecordCount > 0 Then
                     adocli.Recordset.MoveLast
                     Data1.Recordset("cl_cedula") = adocli.Recordset.RecordCount
                  Else
                     Data1.Recordset("cl_cedula") = 0
                  End If
                  
                  Data1.Recordset.Update
                  If Adodc1.Recordset("cnv_pmserv") = 1 Then 'se realiza facturación
             
                  End If
               Else
                  If Adodc1.Recordset("cnv_cant_r") = 1 Then 'un solo recibo
                     Data1.Recordset.AddNew
                     Data1.Recordset("cl_codconv") = Adodc1.Recordset("cnv_codigo")
                     Data1.Recordset("cl_apellid") = Mid(Adodc1.Recordset("cnv_desc"), 1, 60)
                     Data1.Recordset("cl_dpto") = Adodc1.Recordset("cnv_colrec")
                     Data1.Recordset("cl_fnac") = Adodc1.Recordset("cnv_desde")
                     Data1.Recordset("cl_fultmov") = Adodc1.Recordset("cnv_hasta")
                     Data1.Recordset("cl_ruc") = Adodc1.Recordset("cnv_ruc")
                     If IsNull(Adodc1.Recordset("cnv_fbaja")) = False Then
                        Data1.Recordset("cl_fecing") = Adodc1.Recordset("cnv_fbaja")
                     End If
                     Data1.Recordset("saldo_cc") = Adodc1.Recordset("cnv_precio")
                     If Adodc1.Recordset("cnv_precio") > 0 Then
                        Adodc2.RecordSource = "Select * from cnv_prec where cnv_codigo ='" & Adodc1.Recordset("cnv_codigo") & "' order by cnv_desde DESC"
                        Adodc2.Refresh
                        If Adodc2.Recordset.RecordCount > 1 Then
                           Adodc2.Recordset.MoveFirst
                           Adodc2.Recordset.MoveNext
                           Data1.Recordset("saldo_cc2") = Adodc2.Recordset("precio")
                        End If
                     End If
                     Data1.Recordset("cl_codced") = Adodc1.Recordset("cnv_cant_r")
                     Data1.Recordset("cl_fax") = Adodc1.Recordset("cnv_alta")
                     Data1.Recordset("cl_codigo") = Adodc1.Recordset("cnv_cuenta")
                     Data1.Recordset("cl_nomcobr") = "A.Protegida"
                     If Adodc1.Recordset("cnv_pmserv") = 1 Then 'se realiza facturación
                        Data1.Recordset("cl_nomcobr") = "A.P. se factura en Dpto."
                     End If
                     adocli.RecordSource = "select * from clientes where estado =" & 1 & " and cl_codconv ='" & Adodc1.Recordset("cnv_codigo") & "'"
                     adocli.Refresh
                     If adocli.Recordset.RecordCount > 0 Then
                        adocli.Recordset.MoveLast
                        Data1.Recordset("cl_cedula") = adocli.Recordset.RecordCount
                     Else
                        Data1.Recordset("cl_cedula") = 0
                     End If
                      
                     Data1.Recordset.Update
                  End If
               End If
            End If
         End If
         Adodc1.Recordset.MoveNext
      Loop
      frm_infconves.MousePointer = 0
      Data1.RecordSource = "Select * from infcli"
      Data1.Refresh
      cr1.ReportFileName = App.path & "\infconves.rpt"
      cr1.ReportTitle = "Informe de CONVENIOS con emisión SAPP"
      cr1.Action = 1
      
   Else
      MsgBox "No existen registros"
   End If
End If
If Combo1.ListIndex = 1 Then
   If Check1.Value = 1 Then
      If Check4.Value = 1 Then
         Adodc1.RecordSource = "Select * from convenio where cnv_hasta >='" & Format(Date, "yyyy/mm/dd") & "' and (cnv_pmserv is null or cnv_pmserv in (0)) and cnv_fbaja is null"
      Else
         Adodc1.RecordSource = "Select * from convenio where cnv_hasta >='" & Format(Date, "yyyy/mm/dd") & "' and cnv_fbaja is null"
      End If
   Else
      If Check2.Value = 1 Then
         Adodc1.RecordSource = "Select * from convenio where cnv_hasta <'" & Format(Date, "yyyy/mm/dd") & "' and cnv_fbaja is null"
      Else
         Adodc1.RecordSource = "Select * from convenio where cnv_fbaja is null"
      End If
   End If
   
   Adodc1.Refresh
   
   If Adodc1.Recordset.RecordCount > 0 Then
      Adodc1.Recordset.MoveFirst
      Do While Not Adodc1.Recordset.EOF
         If IsNull(Adodc1.Recordset("cnv_colrec")) = False Then
            If Adodc1.Recordset("cnv_colrec") = "M" Or Adodc1.Recordset("cnv_colrec") = "R" Or Adodc1.Recordset("cnv_colrec") = "A" Or Adodc1.Recordset("cnv_colrec") = "V" Then
               If Adodc1.Recordset("cnv_cant_r") = 1 Then 'un recibos
                  Data1.Recordset.AddNew
                  Data1.Recordset("cl_codconv") = Adodc1.Recordset("cnv_codigo")
                  Data1.Recordset("cl_apellid") = Mid(Adodc1.Recordset("cnv_desc"), 1, 60)
                  Data1.Recordset("cl_dpto") = Adodc1.Recordset("cnv_colrec")
                  Data1.Recordset("cl_fnac") = Adodc1.Recordset("cnv_desde")
                  Data1.Recordset("cl_fultmov") = Adodc1.Recordset("cnv_hasta")
                  Data1.Recordset("cl_ruc") = Adodc1.Recordset("cnv_ruc")
                  Data1.Recordset("saldo_cc") = Adodc1.Recordset("cnv_precio")
                  If IsNull(Adodc1.Recordset("cnv_fbaja")) = False Then
                     Data1.Recordset("cl_fecing") = Adodc1.Recordset("cnv_fbaja")
                  End If
                  If Adodc1.Recordset("cnv_precio") > 0 Then
                     Adodc2.RecordSource = "Select * from cnv_prec where cnv_codigo ='" & Adodc1.Recordset("cnv_codigo") & "' order by cnv_desde DESC"
                     Adodc2.Refresh
                     If Adodc2.Recordset.RecordCount > 1 Then
                        Adodc2.Recordset.MoveFirst
                        Adodc2.Recordset.MoveNext
                        Data1.Recordset("saldo_cc2") = Adodc2.Recordset("precio")
                     End If
                  End If
                  Data1.Recordset("cl_codced") = Adodc1.Recordset("cnv_cant_r")
                  Data1.Recordset("cl_fax") = Adodc1.Recordset("cnv_alta")
                  Data1.Recordset("cl_codigo") = Adodc1.Recordset("cnv_cuenta")
                  Data1.Recordset("cl_nomcobr") = "A.P. Emisión"
                  If Adodc1.Recordset("cnv_pmserv") = 1 Then 'se realiza facturación
                     Data1.Recordset("cl_nomcobr") = "A.P. Facturación"
             
                  End If
                  
                  Data1.Recordset.Update
               End If
            End If
         End If
         Adodc1.Recordset.MoveNext
      Loop
      frm_infconves.MousePointer = 0
      Data1.RecordSource = "Select * from infcli"
      Data1.Refresh
      cr1.ReportFileName = App.path & "\infconves.rpt"
      cr1.ReportTitle = "Informe de CONVENIOS con un solo recibo de emisión (A.P.)"
      cr1.Action = 1
   
   Else
      MsgBox "No existen registros"
   End If
End If
If Combo1.ListIndex = 2 Then
   If Check1.Value = 1 Then
      Adodc1.RecordSource = "Select * from convenio where cnv_hasta >='" & Format(Date, "yyyy/mm/dd") & "' and cnv_fbaja is null"
   Else
      If Check2.Value = 1 Then
         Adodc1.RecordSource = "Select * from convenio where cnv_hasta <'" & Format(Date, "yyyy/mm/dd") & "' and cnv_fbaja is null"
      Else
         If Check4.Value = 1 Then
            Adodc1.RecordSource = "Select * from convenio where cnv_codigo is not null and cnv_fbaja is null and (cnv_emite is null or cnv_emite ='" & "NO" & "') and (cnv_pmserv is null or cnv_pmserv in (-1,0))"
         Else
            Adodc1.RecordSource = "Select * from convenio where cnv_codigo is not null and cnv_fbaja is null"
         End If
      End If
   End If
   
   Adodc1.Refresh
   
   If Adodc1.Recordset.RecordCount > 0 Then
      Adodc1.Recordset.MoveFirst
      Do While Not Adodc1.Recordset.EOF
         Data1.Recordset.AddNew
         Data1.Recordset("cl_codconv") = Adodc1.Recordset("cnv_codigo")
         Data1.Recordset("cl_apellid") = Mid(Adodc1.Recordset("cnv_desc"), 1, 60)
         Data1.Recordset("cl_dpto") = Adodc1.Recordset("cnv_colrec")
         Data1.Recordset("cl_fnac") = Adodc1.Recordset("cnv_desde")
         Data1.Recordset("cl_fultmov") = Adodc1.Recordset("cnv_hasta")
         Data1.Recordset("cl_ruc") = Adodc1.Recordset("cnv_ruc")
         If IsNull(Adodc1.Recordset("cnv_fbaja")) = False Then
            Data1.Recordset("cl_fecing") = Adodc1.Recordset("cnv_fbaja")
         End If
         Data1.Recordset("saldo_cc") = Adodc1.Recordset("cnv_precio")
'         If Adodc1.Recordset("cnv_precio") > 0 Then
'            Adodc2.RecordSource = "Select * from cnv_prec where cnv_codigo ='" & Adodc1.Recordset("cnv_codigo") & "' order by cnv_desde DESC"
'            Adodc2.Refresh
'            If Adodc2.Recordset.RecordCount > 1 Then
'               Adodc2.Recordset.MoveFirst
'               Adodc2.Recordset.MoveNext
'               Data1.Recordset("saldo_cc2") = Adodc2.Recordset("precio")
'            End If
'         End If

         Data1.Recordset("saldo_cc2") = 0
         adocli.RecordSource = "select * from clientes where estado =" & 1 & " and cl_codconv ='" & Adodc1.Recordset("cnv_codigo") & "'"
         adocli.Refresh
         If adocli.Recordset.RecordCount > 0 Then
            adocli.Recordset.MoveLast
            Data1.Recordset("cl_cedula") = adocli.Recordset.RecordCount
         Else
            Data1.Recordset("cl_cedula") = 0
         End If
         Data1.Recordset("cl_codced") = Adodc1.Recordset("cnv_cant_r")
         Data1.Recordset("cl_fax") = Adodc1.Recordset("cnv_alta")
         Data1.Recordset("cl_codigo") = Adodc1.Recordset("cnv_cuenta")
         If IsNull(Adodc1.Recordset("cnv_cant_r")) = False Then
            If Adodc1.Recordset("cnv_cant_r") = 2 Then
               Data1.Recordset("cl_nomcobr") = "Varios e-ticket"
            Else
               If Adodc1.Recordset("cnv_cant_r") = 1 Then 'un solo recibo
                  Data1.Recordset("cl_nomcobr") = "A.P. Emisión"
               Else
                  Data1.Recordset("cl_nomcobr") = "S/D"
               End If
            End If
         End If
         Data1.Recordset.Update
         Adodc1.Recordset.MoveNext
      Loop
      frm_infconves.MousePointer = 0
      Data1.RecordSource = "Select * from infcli"
      Data1.Refresh
      cr1.ReportFileName = App.path & "\infconves.rpt"
      cr1.ReportTitle = "Informe de TODOS los CONVENIOS"
      cr1.Action = 1
   
   Else
      MsgBox "No existen registros"
   End If
End If
If Combo1.ListIndex = 3 Then
   Adodc1.RecordSource = "Select * from convenio where cnv_fbaja is not null order by cnv_fbaja"
   Adodc1.Refresh
   If Adodc1.Recordset.RecordCount > 0 Then
      Adodc1.Recordset.MoveFirst
      Do While Not Adodc1.Recordset.EOF
         If IsNull(Adodc1.Recordset("cnv_colrec")) = False Then
            If Adodc1.Recordset("cnv_colrec") = "M" Or Adodc1.Recordset("cnv_colrec") = "R" Or Adodc1.Recordset("cnv_colrec") = "A" Or Adodc1.Recordset("cnv_colrec") = "V" Then
               If Adodc1.Recordset("cnv_cant_r") = 2 Then 'varios recibos
                  Data1.Recordset.AddNew
                  Data1.Recordset("cl_codconv") = Adodc1.Recordset("cnv_codigo")
                  Data1.Recordset("cl_apellid") = Mid(Adodc1.Recordset("cnv_desc"), 1, 60)
                  Data1.Recordset("cl_dpto") = Adodc1.Recordset("cnv_colrec")
                  Data1.Recordset("cl_fnac") = Adodc1.Recordset("cnv_desde")
                  Data1.Recordset("cl_fultmov") = Adodc1.Recordset("cnv_hasta")
                  Data1.Recordset("cl_ruc") = Adodc1.Recordset("cnv_ruc")
                  If IsNull(Adodc1.Recordset("cnv_fbaja")) = False Then
                     Data1.Recordset("cl_fecing") = Adodc1.Recordset("cnv_fbaja")
                  End If
                  Data1.Recordset("saldo_cc") = Adodc1.Recordset("cnv_precio")
                  If Adodc1.Recordset("cnv_precio") > 0 Then
                     Adodc2.RecordSource = "Select * from cnv_prec where cnv_codigo ='" & Adodc1.Recordset("cnv_codigo") & "' order by cnv_desde DESC"
                     Adodc2.Refresh
                     If Adodc2.Recordset.RecordCount > 1 Then
                        Adodc2.Recordset.MoveFirst
                        Adodc2.Recordset.MoveNext
                        Data1.Recordset("saldo_cc2") = Adodc2.Recordset("precio")
                     End If
                  End If
                  Data1.Recordset("cl_codced") = Adodc1.Recordset("cnv_cant_r")
                  Data1.Recordset("cl_fax") = Adodc1.Recordset("cnv_alta")
                  Data1.Recordset("cl_codigo") = Adodc1.Recordset("cnv_cuenta")
                  Data1.Recordset("cl_nomcobr") = "Varios e-ticket"
                  adocli.RecordSource = "select * from clientes where estado =" & 1 & " and cl_codconv ='" & Adodc1.Recordset("cnv_codigo") & "'"
                  adocli.Refresh
                  If adocli.Recordset.RecordCount > 0 Then
                     adocli.Recordset.MoveLast
                     Data1.Recordset("cl_cedula") = adocli.Recordset.RecordCount
                  Else
                     Data1.Recordset("cl_cedula") = 0
                  End If
                  Data1.Recordset.Update
                  If Adodc1.Recordset("cnv_pmserv") = 1 Then 'se realiza facturación
             
                  End If
               Else
                  If Adodc1.Recordset("cnv_cant_r") = 1 Then 'un solo recibo
                     Data1.Recordset.AddNew
                     Data1.Recordset("cl_codconv") = Adodc1.Recordset("cnv_codigo")
                     Data1.Recordset("cl_apellid") = Mid(Adodc1.Recordset("cnv_desc"), 1, 60)
                     Data1.Recordset("cl_dpto") = Adodc1.Recordset("cnv_colrec")
                     Data1.Recordset("cl_fnac") = Adodc1.Recordset("cnv_desde")
                     Data1.Recordset("cl_fultmov") = Adodc1.Recordset("cnv_hasta")
                     Data1.Recordset("cl_ruc") = Adodc1.Recordset("cnv_ruc")
                     If IsNull(Adodc1.Recordset("cnv_fbaja")) = False Then
                        Data1.Recordset("cl_fecing") = Adodc1.Recordset("cnv_fbaja")
                     End If
                     Data1.Recordset("saldo_cc") = Adodc1.Recordset("cnv_precio")
                     If Adodc1.Recordset("cnv_precio") > 0 Then
                        Adodc2.RecordSource = "Select * from cnv_prec where cnv_codigo ='" & Adodc1.Recordset("cnv_codigo") & "' order by cnv_desde DESC"
                        Adodc2.Refresh
                        If Adodc2.Recordset.RecordCount > 1 Then
                           Adodc2.Recordset.MoveFirst
                           Adodc2.Recordset.MoveNext
                           Data1.Recordset("saldo_cc2") = Adodc2.Recordset("precio")
                        End If
                     End If
                     Data1.Recordset("cl_codced") = Adodc1.Recordset("cnv_cant_r")
                     Data1.Recordset("cl_fax") = Adodc1.Recordset("cnv_alta")
                     Data1.Recordset("cl_codigo") = Adodc1.Recordset("cnv_cuenta")
                     Data1.Recordset("cl_nomcobr") = "A.Protegida"
                     If Adodc1.Recordset("cnv_pmserv") = 1 Then 'se realiza facturación
                        Data1.Recordset("cl_nomcobr") = "A.P. se factura en Dpto."
                     End If
                     adocli.RecordSource = "select * from clientes where estado =" & 1 & " and cl_codconv ='" & Adodc1.Recordset("cnv_codigo") & "'"
                     adocli.Refresh
                     If adocli.Recordset.RecordCount > 0 Then
                        adocli.Recordset.MoveLast
                        Data1.Recordset("cl_cedula") = adocli.Recordset.RecordCount
                     Else
                        Data1.Recordset("cl_cedula") = 0
                     End If
                     Data1.Recordset.Update
                  End If
               End If
            End If
         End If
         Adodc1.Recordset.MoveNext
      Loop
      frm_infconves.MousePointer = 0
      Data1.RecordSource = "Select * from infcli"
      Data1.Refresh
      cr1.ReportFileName = App.path & "\infconvesb.rpt"
      cr1.ReportTitle = "Informe de CONVENIOS con FECHA de BAJA registrada."
      cr1.Action = 1
      
   Else
      MsgBox "No existen registros"
   End If
End If
   
   
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.path & "\informes.mdb"

Adodc1.ConnectionString = "dsn=" & Xconexrmt
Adodc2.ConnectionString = "dsn=" & Xconexrmt
adocli.ConnectionString = "dsn=" & Xconexrmt


End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub
