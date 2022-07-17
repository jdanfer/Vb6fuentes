VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frm_prestamo 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Préstamos BROU"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8250
   Icon            =   "frm_prestamo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   8250
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_abmp 
      Caption         =   "data_abmp"
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
      Top             =   6360
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Historial de movimientos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   66
      Top             =   6600
      Width           =   2055
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
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
      Width           =   3015
   End
   Begin Crystal.CrystalReport cr1 
      Left            =   6120
      Top             =   6480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      DiscardSavedData=   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
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
      Left            =   4440
      Picture         =   "frm_prestamo.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   55
      ToolTipText     =   "Informes"
      Top             =   5880
      Width           =   495
   End
   Begin VB.CommandButton bbuscar 
      BackColor       =   &H00FFFFFF&
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
      Left            =   5280
      Picture         =   "frm_prestamo.frx":09CC
      Style           =   1  'Graphical
      TabIndex        =   54
      ToolTipText     =   "Buscar"
      Top             =   5880
      Width           =   495
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5280
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton bp 
      BackColor       =   &H0080C0FF&
      Caption         =   "Procesar..."
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
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   5880
      Width           =   1215
   End
   Begin VB.CommandButton be 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3600
      Picture         =   "frm_prestamo.frx":0F56
      Style           =   1  'Graphical
      TabIndex        =   51
      ToolTipText     =   "Eliminar registro"
      Top             =   5880
      Width           =   495
   End
   Begin VB.CommandButton bc 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   2760
      Picture         =   "frm_prestamo.frx":14E0
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Cancelar acción"
      Top             =   5880
      Width           =   495
   End
   Begin VB.CommandButton bg 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   495
      Left            =   1920
      Picture         =   "frm_prestamo.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Grabar"
      Top             =   5880
      Width           =   495
   End
   Begin VB.CommandButton bm 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1080
      Picture         =   "frm_prestamo.frx":1FF4
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Editar registro"
      Top             =   5880
      Width           =   495
   End
   Begin VB.CommandButton bn 
      BackColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_prestamo.frx":257E
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Alta de registro"
      Top             =   5880
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Datos del personal"
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
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
      Begin VB.TextBox txt_nomc 
         Height          =   285
         Left            =   2040
         MaxLength       =   30
         TabIndex        =   64
         Top             =   3720
         Width           =   3735
      End
      Begin VB.TextBox txt_codcedc 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7200
         MaxLength       =   1
         TabIndex        =   62
         Top             =   3360
         Width           =   375
      End
      Begin VB.TextBox txt_cedc 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5760
         MaxLength       =   7
         TabIndex        =   61
         Top             =   3360
         Width           =   1455
      End
      Begin VB.TextBox txt_estciv 
         Height          =   285
         Left            =   2040
         MaxLength       =   1
         TabIndex        =   59
         ToolTipText     =   "1=Soltero, 2=Casado, 3=Separado, 4= Viudo, 5=Divorciado, 6=Concubino"
         Top             =   3360
         Width           =   615
      End
      Begin VB.TextBox txt_tel 
         Height          =   285
         Left            =   5280
         MaxLength       =   15
         TabIndex        =   57
         Top             =   2280
         Width           =   2295
      End
      Begin VB.TextBox ttoth 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6000
         TabIndex        =   46
         Top             =   5160
         Width           =   1575
      End
      Begin VB.TextBox timps 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   44
         Top             =   5160
         Width           =   1935
      End
      Begin VB.TextBox tretleg 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6000
         TabIndex        =   42
         Top             =   4800
         Width           =   1575
      End
      Begin VB.TextBox tgtia 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   40
         Top             =   4800
         Width           =   1935
      End
      Begin VB.TextBox tret 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6000
         TabIndex        =   38
         Top             =   4440
         Width           =   1575
      End
      Begin VB.TextBox tcuo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   36
         Top             =   4440
         Width           =   1935
      End
      Begin VB.TextBox tcanhd 
         Height          =   285
         Left            =   6960
         TabIndex        =   34
         Top             =   4080
         Width           =   615
      End
      Begin VB.TextBox thd 
         Height          =   285
         Left            =   4560
         TabIndex        =   32
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox tmj 
         Height          =   285
         Left            =   2040
         TabIndex        =   30
         Top             =   4080
         Width           =   735
      End
      Begin VB.TextBox tcarg 
         Height          =   285
         Left            =   6720
         TabIndex        =   28
         Top             =   3000
         Width           =   855
      End
      Begin VB.TextBox tden 
         Height          =   285
         Left            =   2040
         TabIndex        =   26
         Top             =   3000
         Width           =   3135
      End
      Begin MSMask.MaskEdBox ming 
         Height          =   255
         Left            =   2040
         TabIndex        =   24
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
      Begin MSMask.MaskEdBox mnac 
         Height          =   255
         Left            =   6240
         TabIndex        =   22
         Top             =   2640
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
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
      Begin VB.TextBox tcodpos 
         Height          =   285
         Left            =   2040
         TabIndex        =   20
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox tloc 
         Height          =   285
         Left            =   5280
         TabIndex        =   18
         Top             =   1920
         Width           =   2295
      End
      Begin VB.TextBox tcoddep 
         Height          =   285
         Left            =   4560
         TabIndex        =   17
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox tnp 
         Height          =   285
         Left            =   2040
         TabIndex        =   15
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox tdir 
         Height          =   285
         Left            =   2040
         TabIndex        =   13
         Top             =   1560
         Width           =   5535
      End
      Begin VB.TextBox tape2 
         Height          =   285
         Left            =   4920
         TabIndex        =   11
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox tape1 
         Height          =   285
         Left            =   2040
         TabIndex        =   10
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox tn2 
         Height          =   285
         Left            =   4920
         TabIndex        =   8
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox tn1 
         Height          =   285
         Left            =   2040
         TabIndex        =   7
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox tsex 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6960
         TabIndex        =   5
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox tcod 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3720
         MaxLength       =   1
         TabIndex        =   3
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox tced 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2040
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label26 
         BackColor       =   &H00C0FFC0&
         Caption         =   "NOMBRE CONYUGE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   63
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label Label25 
         BackColor       =   &H00C0FFC0&
         Caption         =   "CEDULA CONYUGE:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3600
         TabIndex        =   60
         Top             =   3360
         Width           =   2175
      End
      Begin VB.Label Label24 
         BackColor       =   &H00C0FFC0&
         Caption         =   "ESTADO CIVIL:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label Label23 
         BackColor       =   &H00C0FFC0&
         Caption         =   "TELEF. PART."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   56
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label22 
         BackColor       =   &H00C0FFFF&
         Caption         =   "1=MASC, 2=FEM."
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
         Left            =   5880
         TabIndex        =   53
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label21 
         BackColor       =   &H00C0FFC0&
         Caption         =   "TOT.HABERES:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   45
         Top             =   5160
         Width           =   1695
      End
      Begin VB.Label Label20 
         BackColor       =   &H00C0FFC0&
         Caption         =   "IMP. SUELDO:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   5160
         Width           =   1815
      End
      Begin VB.Label Label19 
         BackColor       =   &H00C0FFC0&
         Caption         =   "RET.LEGALES:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   41
         Top             =   4800
         Width           =   1695
      End
      Begin VB.Label Label18 
         BackColor       =   &H00C0FFC0&
         Caption         =   "GTIA. ALQUILER:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   4800
         Width           =   1815
      End
      Begin VB.Label Label17 
         BackColor       =   &H00C0FFC0&
         Caption         =   "RET.JUDICIAL:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   37
         Top             =   4440
         Width           =   1695
      End
      Begin VB.Label Label16 
         BackColor       =   &H00C0FFC0&
         Caption         =   "CUOTA SUGERIDA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   35
         Top             =   4440
         Width           =   1815
      End
      Begin VB.Label Label15 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Cant. Hs./Días:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5640
         TabIndex        =   33
         Top             =   4080
         Width           =   1335
      End
      Begin VB.Label Label14 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Hora/Diario:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   31
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label13 
         BackColor       =   &H00C0FFC0&
         Caption         =   "MES/JORNAL:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   4080
         Width           =   1815
      End
      Begin VB.Label Label12 
         BackColor       =   &H00C0FFC0&
         Caption         =   "C.CARGO:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5520
         TabIndex        =   27
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "DENOMINACION:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label10 
         BackColor       =   &H00C0FFC0&
         Caption         =   "FECHA ING."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   2640
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0FFC0&
         Caption         =   "FECHA NAC."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   21
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0FFC0&
         Caption         =   "COD.POSTAL:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2280
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Cod.Dpto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3240
         TabIndex        =   16
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Nro.PUERTA:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   1920
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "DIRECCION:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0FFC0&
         Caption         =   "APELLIDOS:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "NOMBRES:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "SEXO:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5880
         TabIndex        =   4
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "CEDULA:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
   End
   Begin VB.OLE OLE1 
      AutoActivate    =   1  'GetFocus
      Height          =   735
      Left            =   3000
      TabIndex        =   65
      Top             =   6120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      X1              =   0
      X2              =   8280
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Image Image1 
      Height          =   1095
      Left            =   4200
      Picture         =   "frm_prestamo.frx":2B08
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   1575
   End
End
Attribute VB_Name = "frm_prestamo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bbuscar_Click()
frm_busperso.Show vbModal

End Sub

Private Sub bc_Click()
If XAlta = 1 Then
   Data1.Recordset.CancelUpdate
   Data1.Refresh
End If
bn.Enabled = True
bm.Enabled = True
bg.Enabled = False
bc.Enabled = False
be.Enabled = True
bbuscar.Enabled = True
bp.Enabled = True
veodatos
Frame1.Enabled = False

End Sub

Private Sub be_Click()
Dim Xmenbor As String
Xmenbor = MsgBox("Desea borrar el registro seleccionado?", vbExclamation + vbYesNo, "Mensaje")
If Xmenbor = vbYes Then
   Data1.Recordset.FindFirst "cedula =" & tced.Text
   If Not Data1.Recordset.NoMatch Then
      Data1.Recordset.Delete
      data_abmp.Recordset.AddNew
      data_abmp.Recordset("cedula") = tced.Text
      data_abmp.Recordset("codver") = tcod.Text
      data_abmp.Recordset("fecha") = Date
      data_abmp.Recordset("hora") = Format(Time, "HH:mm:ss")
      data_abmp.Recordset("usuario") = WElusuario
      data_abmp.Recordset("desc") = "BORRA REGISTRO"
      data_abmp.Recordset.Update
      borrarcam
   End If
End If


End Sub

Private Sub bg_Click()
If XAlta = 1 Then
   Data1.Recordset("cedula") = tced.Text
   Data1.Recordset("codver") = tcod.Text
   Data1.Recordset("sexo") = tsex.Text
   Data1.Recordset("nom1") = tn1.Text
   Data1.Recordset("nom2") = tn2.Text
   Data1.Recordset("ape1") = tape1.Text
   Data1.Recordset("ape2") = tape2.Text
   Data1.Recordset("calle") = tdir.Text
   If tnp.Text = "" Then
      Data1.Recordset("nropuerta") = "S/N"
   Else
      Data1.Recordset("nropuerta") = tnp.Text
   End If
   Data1.Recordset("coddep") = tcoddep.Text
   Data1.Recordset("localid") = tloc.Text
   If tcodpos.Text = "" Then
      Data1.Recordset("codpos") = 9100
   Else
      Data1.Recordset("codpos") = tcodpos.Text
   End If
   If txt_tel.Text = "" Then
   Else
      Data1.Recordset("telef") = txt_tel.Text
   End If
   If txt_estciv.Text = "" Then
      Data1.Recordset("estciv") = 1
   Else
      Data1.Recordset("estciv") = txt_estciv.Text
   End If
   If txt_cedc.Text = "" Then
      txt_cedc.Text = 0
      txt_codcedc.Text = 0
   End If
   Data1.Recordset("cedc") = txt_cedc.Text
   Data1.Recordset("codcedc") = txt_codcedc.Text
   Data1.Recordset("nomc") = txt_nomc.Text
   Data1.Recordset("fecnac") = Format(mnac.Text, "dd/mm/yyyy")
   Data1.Recordset("fecing") = Format(ming.Text, "dd/mm/yyyy")
   Data1.Recordset("desccar") = tden.Text
   Data1.Recordset("caraccar") = tcarg.Text
   Data1.Recordset("mj") = tmj.Text
   Data1.Recordset("hd") = thd.Text
   If tcanhd.Text = "" Then
      tcanhd.Text = 0
   End If
   Data1.Recordset("canths") = tcanhd.Text
   If tcuo.Text = "" Then
      tcuo.Text = 0
   End If
   Data1.Recordset("cuosug") = tcuo.Text
   If tret.Text = "" Then
      tret.Text = 0
   End If
   Data1.Recordset("retjud") = tret.Text
   If tgtia.Text = "" Then
      tgtia.Text = 0
   End If
   Data1.Recordset("alqui") = tgtia.Text
   Data1.Recordset("retleg") = tretleg.Text
   Data1.Recordset("impsue") = timps.Text
   Data1.Recordset("tothab") = ttoth.Text
   Data1.Recordset.Update
   XAlta = 0
   data_abmp.Recordset.AddNew
   data_abmp.Recordset("cedula") = tced.Text
   data_abmp.Recordset("codver") = tcod.Text
   data_abmp.Recordset("sexo") = tsex.Text
   data_abmp.Recordset("nom1") = tn1.Text
   data_abmp.Recordset("nom2") = tn2.Text
   data_abmp.Recordset("ape1") = tape1.Text
   data_abmp.Recordset("ape2") = tape2.Text
   data_abmp.Recordset("calle") = tdir.Text
   If tnp.Text = "" Then
      data_abmp.Recordset("nropuerta") = "S/N"
   Else
      data_abmp.Recordset("nropuerta") = tnp.Text
   End If
   data_abmp.Recordset("coddep") = tcoddep.Text
   data_abmp.Recordset("localid") = tloc.Text
   If tcodpos.Text = "" Then
      data_abmp.Recordset("codpos") = 9100
   Else
      data_abmp.Recordset("codpos") = tcodpos.Text
   End If
   If txt_tel.Text = "" Then
   Else
      data_abmp.Recordset("telef") = txt_tel.Text
   End If
   If txt_estciv.Text = "" Then
      data_abmp.Recordset("estciv") = 1
   Else
      data_abmp.Recordset("estciv") = txt_estciv.Text
   End If
   If txt_cedc.Text = "" Then
      txt_cedc.Text = 0
      txt_codcedc.Text = 0
   End If
   data_abmp.Recordset("cedc") = txt_cedc.Text
   data_abmp.Recordset("codcedc") = txt_codcedc.Text
   data_abmp.Recordset("nomc") = txt_nomc.Text
   data_abmp.Recordset("fecnac") = Format(mnac.Text, "dd/mm/yyyy")
   data_abmp.Recordset("fecing") = Format(ming.Text, "dd/mm/yyyy")
   data_abmp.Recordset("desccar") = tden.Text
   data_abmp.Recordset("caraccar") = tcarg.Text
   data_abmp.Recordset("mj") = tmj.Text
   data_abmp.Recordset("hd") = thd.Text
   If tcanhd.Text = "" Then
      tcanhd.Text = 0
   End If
   data_abmp.Recordset("canths") = tcanhd.Text
   If tcuo.Text = "" Then
      tcuo.Text = 0
   End If
   data_abmp.Recordset("cuosug") = tcuo.Text
   If tret.Text = "" Then
      tret.Text = 0
   End If
   data_abmp.Recordset("retjud") = tret.Text
   If tgtia.Text = "" Then
      tgtia.Text = 0
   End If
   data_abmp.Recordset("alqui") = tgtia.Text
   data_abmp.Recordset("retleg") = tretleg.Text
   data_abmp.Recordset("impsue") = timps.Text
   data_abmp.Recordset("tothab") = ttoth.Text
   data_abmp.Recordset("fecha") = Date
   data_abmp.Recordset("hora") = Format(Time, "HH:mm:ss")
   data_abmp.Recordset("usuario") = WElusuario
   data_abmp.Recordset("desc") = "CREA REGISTRO"
   data_abmp.Recordset.Update
Else
   Data1.Recordset.Edit
   Data1.Recordset("cedula") = tced.Text
   Data1.Recordset("codver") = tcod.Text
   Data1.Recordset("sexo") = tsex.Text
   Data1.Recordset("nom1") = tn1.Text
   Data1.Recordset("nom2") = tn2.Text
   Data1.Recordset("ape1") = tape1.Text
   Data1.Recordset("ape2") = tape2.Text
   Data1.Recordset("calle") = tdir.Text
   If tnp.Text = "" Then
      Data1.Recordset("nropuerta") = "S/N"
   Else
      Data1.Recordset("nropuerta") = tnp.Text
   End If
   Data1.Recordset("coddep") = tcoddep.Text
   Data1.Recordset("localid") = tloc.Text
   If tcodpos.Text = "" Then
      Data1.Recordset("codpos") = 9100
   Else
      Data1.Recordset("codpos") = tcodpos.Text
   End If
   If txt_tel.Text = "" Then
   Else
      Data1.Recordset("telef") = txt_tel.Text
   End If
   If txt_estciv.Text = "" Then
      Data1.Recordset("estciv") = 1
   Else
      Data1.Recordset("estciv") = txt_estciv.Text
   End If
   If txt_cedc.Text = "" Then
      txt_cedc.Text = 0
      txt_codcedc.Text = 0
   End If
   Data1.Recordset("cedc") = txt_cedc.Text
   Data1.Recordset("codcedc") = txt_codcedc.Text
   Data1.Recordset("nomc") = txt_nomc.Text
   Data1.Recordset("fecnac") = Format(mnac.Text, "dd/mm/yyyy")
   Data1.Recordset("fecing") = Format(ming.Text, "dd/mm/yyyy")
   Data1.Recordset("desccar") = tden.Text
   Data1.Recordset("caraccar") = tcarg.Text
   Data1.Recordset("mj") = tmj.Text
   Data1.Recordset("hd") = thd.Text
   If tcanhd.Text = "" Then
      tcanhd.Text = 0
   End If
   Data1.Recordset("canths") = tcanhd.Text
   If tcuo.Text = "" Then
      tcuo.Text = 0
   End If
   Data1.Recordset("cuosug") = tcuo.Text
   If tret.Text = "" Then
      tret.Text = 0
   End If
   Data1.Recordset("retjud") = tret.Text
   If tgtia.Text = "" Then
      tgtia.Text = 0
   End If
   Data1.Recordset("alqui") = tgtia.Text
   Data1.Recordset("retleg") = tretleg.Text
   Data1.Recordset("impsue") = timps.Text
   Data1.Recordset("tothab") = ttoth.Text
   Data1.Recordset.Update
   XAlta = 0
   data_abmp.Recordset.AddNew
   data_abmp.Recordset("cedula") = tced.Text
   data_abmp.Recordset("codver") = tcod.Text
   data_abmp.Recordset("sexo") = tsex.Text
   data_abmp.Recordset("nom1") = tn1.Text
   data_abmp.Recordset("nom2") = tn2.Text
   data_abmp.Recordset("ape1") = tape1.Text
   data_abmp.Recordset("ape2") = tape2.Text
   data_abmp.Recordset("calle") = tdir.Text
   If tnp.Text = "" Then
      data_abmp.Recordset("nropuerta") = "S/N"
   Else
      data_abmp.Recordset("nropuerta") = tnp.Text
   End If
   data_abmp.Recordset("coddep") = tcoddep.Text
   data_abmp.Recordset("localid") = tloc.Text
   If tcodpos.Text = "" Then
      data_abmp.Recordset("codpos") = 9100
   Else
      data_abmp.Recordset("codpos") = tcodpos.Text
   End If
   If txt_tel.Text = "" Then
   Else
      data_abmp.Recordset("telef") = txt_tel.Text
   End If
   If txt_estciv.Text = "" Then
      data_abmp.Recordset("estciv") = 1
   Else
      data_abmp.Recordset("estciv") = txt_estciv.Text
   End If
   If txt_cedc.Text = "" Then
      txt_cedc.Text = 0
      txt_codcedc.Text = 0
   End If
   data_abmp.Recordset("cedc") = txt_cedc.Text
   data_abmp.Recordset("codcedc") = txt_codcedc.Text
   data_abmp.Recordset("nomc") = txt_nomc.Text
   data_abmp.Recordset("fecnac") = Format(mnac.Text, "dd/mm/yyyy")
   data_abmp.Recordset("fecing") = Format(ming.Text, "dd/mm/yyyy")
   data_abmp.Recordset("desccar") = tden.Text
   data_abmp.Recordset("caraccar") = tcarg.Text
   data_abmp.Recordset("mj") = tmj.Text
   data_abmp.Recordset("hd") = thd.Text
   If tcanhd.Text = "" Then
      tcanhd.Text = 0
   End If
   data_abmp.Recordset("canths") = tcanhd.Text
   If tcuo.Text = "" Then
      tcuo.Text = 0
   End If
   data_abmp.Recordset("cuosug") = tcuo.Text
   If tret.Text = "" Then
      tret.Text = 0
   End If
   data_abmp.Recordset("retjud") = tret.Text
   If tgtia.Text = "" Then
      tgtia.Text = 0
   End If
   data_abmp.Recordset("alqui") = tgtia.Text
   data_abmp.Recordset("retleg") = tretleg.Text
   data_abmp.Recordset("impsue") = timps.Text
   data_abmp.Recordset("tothab") = ttoth.Text
   data_abmp.Recordset("fecha") = Date
   data_abmp.Recordset("hora") = Format(Time, "HH:mm:ss")
   data_abmp.Recordset("usuario") = WElusuario
   data_abmp.Recordset("desc") = "MODIFICACION"
   data_abmp.Recordset.Update
   data_abmp.Refresh

End If
bn.Enabled = True
bm.Enabled = True
bg.Enabled = False
bc.Enabled = False
be.Enabled = True
bbuscar.Enabled = True
bp.Enabled = True
Frame1.Enabled = False

End Sub

Private Sub bm_Click()
Frame1.Enabled = True
XAlta = 0
'borrarcam
Data1.Recordset.FindFirst "cedula =" & tced.Text
If Not Data1.Recordset.NoMatch Then
   veodatos
    tced.SetFocus
    bn.Enabled = False
    bm.Enabled = False
    bg.Enabled = True
    bc.Enabled = True
    be.Enabled = False
    bbuscar.Enabled = False
    bp.Enabled = False
Else
    MsgBox "Error de cédula, verifique", vbCritical, "Mensaje"
    End
End If

End Sub

Private Sub bn_Click()
Frame1.Enabled = True
borrarcam
tced.SetFocus
XAlta = 1
bn.Enabled = False
bm.Enabled = False
bg.Enabled = True
bc.Enabled = True
be.Enabled = False
bbuscar.Enabled = False
bp.Enabled = False
Data1.Recordset.AddNew

End Sub

Private Sub bp_Click()
Dim Micadena As String
Dim XLaced As String
frm_prestamo.MousePointer = 11
Dim xcuenta As Integer
Data1.RecordSource = "select * from prestamo"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Data1.Recordset.MoveFirst
   Open App.path & "\hab5398.txt" For Output As #1
   Do While Not Data1.Recordset.EOF
      If Len(Trim(str(Data1.Recordset("cedula")))) < 7 Then
         Micadena = "0" + Trim(str(Data1.Recordset("cedula"))) + Trim(str(Data1.Recordset("codver")))
      Else
         Micadena = Trim(str(Data1.Recordset("cedula"))) + Trim(str(Data1.Recordset("codver")))
      End If
      XLaced = Trim(Micadena)
      Micadena = Micadena + Mid(Data1.Recordset("nom1"), 1, 30)
      xcuenta = Len(Data1.Recordset("nom1"))
      xcuenta = xcuenta + 1
      For xcuenta = xcuenta To 30
          Micadena = Micadena + " "
      Next
      If IsNull(Data1.Recordset("nom2")) = False Then
         Micadena = Micadena + Mid(Data1.Recordset("nom2"), 1, 30)
         xcuenta = Len(Data1.Recordset("nom2"))
         xcuenta = xcuenta + 1
         For xcuenta = xcuenta To 30
             Micadena = Micadena + " "
         Next
      Else
         Micadena = Micadena + "                              "
      End If
      Micadena = Micadena + Mid(Data1.Recordset("ape1"), 1, 30)
      xcuenta = Len(Data1.Recordset("ape1"))
      xcuenta = xcuenta + 1
      For xcuenta = xcuenta To 30
          Micadena = Micadena + " "
      Next
      If IsNull(Data1.Recordset("ape2")) = False Then
         Micadena = Micadena + Mid(Data1.Recordset("ape2"), 1, 30)
         xcuenta = Len(Data1.Recordset("ape2"))
         xcuenta = xcuenta + 1
         For xcuenta = xcuenta To 30
            Micadena = Micadena + " "
         Next
      Else
         Micadena = Micadena + "                              "
      End If
      Micadena = Micadena + Mid(Data1.Recordset("calle"), 1, 30)
      xcuenta = Len(Data1.Recordset("calle"))
      xcuenta = xcuenta + 1
      For xcuenta = xcuenta To 30
          Micadena = Micadena + " "
      Next
      If IsNull(Data1.Recordset("nropuerta")) = False Then
         Micadena = Micadena + Mid(Data1.Recordset("nropuerta"), 1, 5)
         xcuenta = Len(Data1.Recordset("nropuerta"))
         xcuenta = xcuenta + 1
         For xcuenta = xcuenta To 5
            Micadena = Micadena + " "
         Next
      Else
         Micadena = Micadena + "     "
      End If
      Micadena = Micadena + "                              " '30 espacios nroapto
      If Len(Trim(str(Data1.Recordset("coddep")))) < 2 Then
         Micadena = Micadena + "0" + Trim(str(Data1.Recordset("coddep")))
      Else
         Micadena = Micadena + Trim(str(Data1.Recordset("coddep")))
      End If
      If IsNull(Data1.Recordset("localid")) = False Then
         Micadena = Micadena + Mid(Data1.Recordset("localid"), 1, 20)
         xcuenta = Len(Data1.Recordset("localid"))
         xcuenta = xcuenta + 1
         For xcuenta = xcuenta To 20
            Micadena = Micadena + " "
         Next
      Else
         Micadena = Micadena + "                    "
      End If
      If Len(Trim(str(Data1.Recordset("codpos")))) = 5 Then
         Micadena = Micadena + Trim(Data1.Recordset("codpos"))
      Else
         If Len(Trim(str(Data1.Recordset("codpos")))) = 4 Then
            Micadena = Micadena + "0" + Trim(Data1.Recordset("codpos"))
         Else
            Micadena = Micadena + "00000"
         End If
      End If
      If Day(Data1.Recordset("fecnac")) < 10 Then
         Micadena = Micadena + "0" + Trim(str(Day(Data1.Recordset("fecnac"))))
      Else
         Micadena = Micadena + Trim(str(Day(Data1.Recordset("fecnac"))))
      End If
      If Month(Data1.Recordset("fecnac")) < 10 Then
         Micadena = Micadena + "0" + Trim(str(Month(Data1.Recordset("fecnac"))))
      Else
         Micadena = Micadena + Trim(str(Month(Data1.Recordset("fecnac"))))
      End If
      Micadena = Micadena + Trim(str(Year(Data1.Recordset("fecnac"))))
      Micadena = Micadena + "05398"
      Micadena = Micadena + Trim(XLaced) + "00"
      Micadena = Micadena + Mid(Data1.Recordset("desccar"), 1, 30)
      xcuenta = Len(Data1.Recordset("desccar"))
      xcuenta = xcuenta + 1
      For xcuenta = xcuenta To 30
          Micadena = Micadena + " "
      Next
      Micadena = Micadena + "0" + Trim(str(Data1.Recordset("caraccar")))
      Micadena = Micadena + Trim(Data1.Recordset("mj"))
      If Day(Data1.Recordset("fecing")) < 10 Then
         Micadena = Micadena + "0" + Trim(str(Day(Data1.Recordset("fecing"))))
      Else
         Micadena = Micadena + Trim(str(Day(Data1.Recordset("fecing"))))
      End If
      If Month(Data1.Recordset("fecing")) < 10 Then
         Micadena = Micadena + "0" + Trim(str(Month(Data1.Recordset("fecing"))))
      Else
         Micadena = Micadena + Trim(str(Month(Data1.Recordset("fecing"))))
      End If
      Micadena = Micadena + Mid(Trim(str(Year(Data1.Recordset("fecing")))), 3, 2)
      Micadena = Micadena + "000000"
      Micadena = Micadena + Trim(Data1.Recordset("hd"))
      If Len(Trim(str(Data1.Recordset("canths")))) = 1 Then
         Micadena = Micadena + "00" + Trim(str(Data1.Recordset("canths")))
      End If
      If Len(Trim(str(Data1.Recordset("canths")))) = 2 Then
         Micadena = Micadena + "0" + Trim(str(Data1.Recordset("canths")))
      End If
      If Len(Trim(str(Data1.Recordset("canths")))) = 3 Then
         Micadena = Micadena + Trim(str(Data1.Recordset("canths")))
      End If
      If Len(Trim(str(Int(Data1.Recordset("cuosug"))))) = 1 Then
         Micadena = Micadena + "000000000000" + Trim(str(Int(Data1.Recordset("cuosug")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("cuosug"))))) = 2 Then
         Micadena = Micadena + "00000000000" + Trim(str(Int(Data1.Recordset("cuosug")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("cuosug"))))) = 3 Then
         Micadena = Micadena + "0000000000" + Trim(str(Int(Data1.Recordset("cuosug")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("cuosug"))))) = 4 Then
         Micadena = Micadena + "000000000" + Trim(str(Int(Data1.Recordset("cuosug")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("cuosug"))))) = 5 Then
         Micadena = Micadena + "00000000" + Trim(str(Int(Data1.Recordset("cuosug")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("cuosug"))))) = 6 Then
         Micadena = Micadena + "0000000" + Trim(str(Int(Data1.Recordset("cuosug")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("cuosug"))))) = 7 Then
         Micadena = Micadena + "000000" + Trim(str(Int(Data1.Recordset("cuosug")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("cuosug"))))) = 8 Then
         Micadena = Micadena + "00000" + Trim(str(Int(Data1.Recordset("cuosug")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("cuosug"))))) = 9 Then
         Micadena = Micadena + "0000" + Trim(str(Int(Data1.Recordset("cuosug")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("retjud"))))) = 1 Then
         Micadena = Micadena + "000000000000" + Trim(str(Int(Data1.Recordset("retjud")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("retjud"))))) = 2 Then
         Micadena = Micadena + "00000000000" + Trim(str(Int(Data1.Recordset("retjud")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("retjud"))))) = 3 Then
         Micadena = Micadena + "0000000000" + Trim(str(Int(Data1.Recordset("retjud")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("retjud"))))) = 4 Then
         Micadena = Micadena + "000000000" + Trim(str(Int(Data1.Recordset("retjud")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("retjud"))))) = 5 Then
         Micadena = Micadena + "00000000" + Trim(str(Int(Data1.Recordset("retjud")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("retjud"))))) = 6 Then
         Micadena = Micadena + "0000000" + Trim(str(Int(Data1.Recordset("retjud")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("retjud"))))) = 7 Then
         Micadena = Micadena + "000000" + Trim(str(Int(Data1.Recordset("retjud")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("retjud"))))) = 8 Then
         Micadena = Micadena + "00000" + Trim(str(Int(Data1.Recordset("retjud")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("retjud"))))) = 9 Then
         Micadena = Micadena + "0000" + Trim(str(Int(Data1.Recordset("retjud")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("alqui"))))) = 1 Then
         Micadena = Micadena + "000000000000" + Trim(str(Int(Data1.Recordset("alqui")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("alqui"))))) = 2 Then
         Micadena = Micadena + "00000000000" + Trim(str(Int(Data1.Recordset("alqui")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("alqui"))))) = 3 Then
         Micadena = Micadena + "0000000000" + Trim(str(Int(Data1.Recordset("alqui")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("alqui"))))) = 4 Then
         Micadena = Micadena + "000000000" + Trim(str(Int(Data1.Recordset("alqui")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("alqui"))))) = 5 Then
         Micadena = Micadena + "00000000" + Trim(str(Int(Data1.Recordset("alqui")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("alqui"))))) = 6 Then
         Micadena = Micadena + "0000000" + Trim(str(Int(Data1.Recordset("alqui")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("alqui"))))) = 7 Then
         Micadena = Micadena + "000000" + Trim(str(Int(Data1.Recordset("alqui")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("alqui"))))) = 8 Then
         Micadena = Micadena + "00000" + Trim(str(Int(Data1.Recordset("alqui")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("alqui"))))) = 9 Then
         Micadena = Micadena + "0000" + Trim(str(Int(Data1.Recordset("alqui")))) + "00"
      End If
      Micadena = Micadena + "000000000000000"
      If Len(Trim(str(Int(Data1.Recordset("retleg"))))) = 1 Then
         Micadena = Micadena + "000000000000" + Trim(str(Int(Data1.Recordset("retleg")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("retleg"))))) = 2 Then
         Micadena = Micadena + "00000000000" + Trim(str(Int(Data1.Recordset("retleg")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("retleg"))))) = 3 Then
         Micadena = Micadena + "0000000000" + Trim(str(Int(Data1.Recordset("retleg")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("retleg"))))) = 4 Then
         Micadena = Micadena + "000000000" + Trim(str(Int(Data1.Recordset("retleg")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("retleg"))))) = 5 Then
         Micadena = Micadena + "00000000" + Trim(str(Int(Data1.Recordset("retleg")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("retleg"))))) = 6 Then
         Micadena = Micadena + "0000000" + Trim(str(Int(Data1.Recordset("retleg")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("retleg"))))) = 7 Then
         Micadena = Micadena + "000000" + Trim(str(Int(Data1.Recordset("retleg")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("retleg"))))) = 8 Then
         Micadena = Micadena + "00000" + Trim(str(Int(Data1.Recordset("retleg")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("retleg"))))) = 9 Then
         Micadena = Micadena + "0000" + Trim(str(Int(Data1.Recordset("retleg")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("impsue"))))) = 1 Then
         Micadena = Micadena + "000000000000" + Trim(str(Int(Data1.Recordset("impsue")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("impsue"))))) = 2 Then
         Micadena = Micadena + "00000000000" + Trim(str(Int(Data1.Recordset("impsue")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("impsue"))))) = 3 Then
         Micadena = Micadena + "0000000000" + Trim(str(Int(Data1.Recordset("impsue")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("impsue"))))) = 4 Then
         Micadena = Micadena + "000000000" + Trim(str(Int(Data1.Recordset("impsue")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("impsue"))))) = 5 Then
         Micadena = Micadena + "00000000" + Trim(str(Int(Data1.Recordset("impsue")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("impsue"))))) = 6 Then
         Micadena = Micadena + "0000000" + Trim(str(Int(Data1.Recordset("impsue")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("impsue"))))) = 7 Then
         Micadena = Micadena + "000000" + Trim(str(Int(Data1.Recordset("impsue")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("impsue"))))) = 8 Then
         Micadena = Micadena + "00000" + Trim(str(Int(Data1.Recordset("impsue")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("impsue"))))) = 9 Then
         Micadena = Micadena + "0000" + Trim(str(Int(Data1.Recordset("impsue")))) + "00"
      End If
      Micadena = Micadena + Trim(str(Data1.Recordset("sexo")))
      If Len(Trim(str(Int(Data1.Recordset("tothab"))))) = 1 Then
         Micadena = Micadena + "000000000000" + Trim(str(Int(Data1.Recordset("tothab")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("tothab"))))) = 2 Then
         Micadena = Micadena + "00000000000" + Trim(str(Int(Data1.Recordset("tothab")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("tothab"))))) = 3 Then
         Micadena = Micadena + "0000000000" + Trim(str(Int(Data1.Recordset("tothab")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("tothab"))))) = 4 Then
         Micadena = Micadena + "000000000" + Trim(str(Int(Data1.Recordset("tothab")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("tothab"))))) = 5 Then
         Micadena = Micadena + "00000000" + Trim(str(Int(Data1.Recordset("tothab")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("tothab"))))) = 6 Then
         Micadena = Micadena + "0000000" + Trim(str(Int(Data1.Recordset("tothab")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("tothab"))))) = 7 Then
         Micadena = Micadena + "000000" + Trim(str(Int(Data1.Recordset("tothab")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("tothab"))))) = 8 Then
         Micadena = Micadena + "00000" + Trim(str(Int(Data1.Recordset("tothab")))) + "00"
      End If
      If Len(Trim(str(Int(Data1.Recordset("tothab"))))) = 9 Then
         Micadena = Micadena + "0000" + Trim(str(Int(Data1.Recordset("tothab")))) + "00"
      End If
      If IsNull(Data1.Recordset("telef")) = False Then
         Micadena = Micadena + Mid(Data1.Recordset("telef"), 1, 15)
         xcuenta = Len(Data1.Recordset("telef"))
         xcuenta = xcuenta + 1
         For xcuenta = xcuenta To 15
             Micadena = Micadena + " "
         Next
      Else
         Micadena = Micadena + "               "
      End If
      If IsNull(Data1.Recordset("estciv")) = False Then
         Micadena = Micadena + Trim(str(Data1.Recordset("estciv")))
      Else
         Micadena = Micadena + "1"
      End If
      If IsNull(Data1.Recordset("cedc")) = False Then
         If Data1.Recordset("cedc") > 0 Then
            If Len(Trim(str(Data1.Recordset("cedula")))) < 7 Then
               Micadena = Micadena + "0" + Trim(str(Data1.Recordset("cedc"))) + Trim(str(Data1.Recordset("codcedc")))
            Else
               Micadena = Micadena + Trim(str(Data1.Recordset("cedc"))) + Trim(str(Data1.Recordset("codcedc")))
            End If
         Else
            Micadena = Micadena + "00000000"
         End If
      Else
         Micadena = Micadena + "00000000"
      End If
      If IsNull(Data1.Recordset("nomc")) = False Then
         Micadena = Micadena + Mid(Data1.Recordset("nomc"), 1, 30)
         xcuenta = Len(Data1.Recordset("nomc"))
         xcuenta = xcuenta + 1
         For xcuenta = xcuenta To 30
             Micadena = Micadena + " "
         Next
      Else
         Micadena = Micadena + "                              "
      End If

      Print #1, Micadena
      Data1.Recordset.MoveNext
   Loop
   Close #1
End If
frm_prestamo.MousePointer = 0
MsgBox "Proceso terminado, se mostrará el archivo guardado cómo: hab5398"
data_abmp.Recordset.AddNew
data_abmp.Recordset("cedula") = 0
data_abmp.Recordset("codver") = 0
data_abmp.Recordset("fecha") = Date
data_abmp.Recordset("hora") = Format(Time, "HH:mm:ss")
data_abmp.Recordset("usuario") = WElusuario
data_abmp.Recordset("desc") = "GENERA ARCHIVO"
data_abmp.Recordset.Update

OLE1.SourceDoc = App.path & "\hab5398.txt"
OLE1.Action = 1
OLE1.DoVerb (-1)
OLE1.Close

End Sub

Private Sub Command1_Click()
Command1.Enabled = False
frm_prestamo.MousePointer = 11
If Data2.Recordset.RecordCount > 0 Then
   Data2.Recordset.MoveFirst
   Do While Not Data2.Recordset.EOF
      Data2.Recordset.Delete
      Data2.Recordset.MoveNext
   Loop
End If
Data1.RecordSource = "select * from prestamo"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Do While Not Data1.Recordset.EOF
      Data2.Recordset.AddNew
      Data2.Recordset("cl_cedula") = Data1.Recordset("cedula")
      Data2.Recordset("cl_codced") = Data1.Recordset("codver")
      Data2.Recordset("cl_apellid") = Data1.Recordset("ape1") + " " + Data1.Recordset("nom1")
      Data2.Recordset("cl_direcci") = Data1.Recordset("calle")
      Data2.Recordset("cl_localid") = Data1.Recordset("localid")
      Data2.Recordset("cl_fnac") = Data1.Recordset("fecnac")
      Data2.Recordset("cl_nombre") = Data1.Recordset("desccar")
      Data2.Recordset("cl_fecing") = Data1.Recordset("fecing")
      Data2.Recordset("cl_nrovend") = Data1.Recordset("canths")
      Data2.Recordset("cl_atrasoa") = Data1.Recordset("cuosug")
      Data2.Recordset("cl_atrasop") = Data1.Recordset("retjud")
      Data2.Recordset("cl_cantpag") = Data1.Recordset("alqui")
      Data2.Recordset("cl_cantdia") = Data1.Recordset("retleg")
      Data2.Recordset("cl_pri_vto") = Data1.Recordset("impsue")
      Data2.Recordset("cl_seg_vto") = Data1.Recordset("tothab")
      Data2.Recordset.Update
      Data1.Recordset.MoveNext
   Loop
   frm_prestamo.MousePointer = 0
   Data2.RecordSource = "infcli"
   Data2.Refresh
   cr1.ReportFileName = App.path & "\infpres.rpt"
   cr1.Action = 1
   
End If
frm_prestamo.MousePointer = 0
Command1.Enabled = True

End Sub

Private Sub Command2_Click()
'MsgBox "Opción no habilitada"
''frm_verpres.Show vbModal

End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.path & "\prestamo.mdb"
Data1.RecordSource = "prestamo"
Data1.Refresh
Data2.DatabaseName = App.path & "\informes.mdb"
Data2.RecordSource = "infcli"
Data2.Refresh
data_abmp.DatabaseName = App.path & "\abmpres.mdb"
data_abmp.RecordSource = "prestamo"
data_abmp.Refresh

veodatos

End Sub

Public Function veodatos()
tced.Text = Data1.Recordset("cedula")
tcod.Text = Data1.Recordset("codver")
tsex.Text = Data1.Recordset("sexo")
tn1.Text = Data1.Recordset("nom1")
If IsNull(Data1.Recordset("nom2")) = False Then
   tn2.Text = Data1.Recordset("nom2")
Else
   tn2.Text = ""
End If
tape1.Text = Data1.Recordset("ape1")
If IsNull(Data1.Recordset("ape2")) = False Then
   tape2.Text = Data1.Recordset("ape2")
Else
   tape2.Text = ""
End If
tdir.Text = Data1.Recordset("calle")
If IsNull(Data1.Recordset("nropuerta")) = False Then
   tnp.Text = Data1.Recordset("nropuerta")
Else
   tnp.Text = ""
End If
tcoddep.Text = Data1.Recordset("coddep")
tloc.Text = Data1.Recordset("localid")
tcodpos.Text = Data1.Recordset("codpos")
If IsNull(Data1.Recordset("telef")) = False Then
   txt_tel.Text = Data1.Recordset("telef")
Else
   txt_tel.Text = ""
End If
If IsNull(Data1.Recordset("estciv")) = False Then
   txt_estciv.Text = Data1.Recordset("estciv")
Else
   txt_estciv.Text = 1
End If
If IsNull(Data1.Recordset("cedc")) = False Then
   txt_cedc.Text = Data1.Recordset("cedc")
Else
   txt_cedc.Text = 0
End If
If IsNull(Data1.Recordset("codcedc")) = False Then
   txt_codcedc.Text = Data1.Recordset("codcedc")
Else
   txt_codcedc.Text = 0
End If
If IsNull(Data1.Recordset("nomc")) = False Then
   txt_nomc.Text = Data1.Recordset("nomc")
Else
   txt_nomc.Text = ""
End If

mnac.Text = Format(Data1.Recordset("fecnac"), "dd/mm/yyyy")
ming.Text = Format(Data1.Recordset("fecing"), "dd/mm/yyyy")
tden.Text = Data1.Recordset("desccar")
tcarg.Text = Data1.Recordset("caraccar")
tmj.Text = Data1.Recordset("mj")
thd.Text = Data1.Recordset("hd")
tcanhd.Text = Data1.Recordset("canths")
tcuo.Text = Format(Data1.Recordset("cuosug"), "Standard")
tret.Text = Format(Data1.Recordset("retjud"), "Standard")
tgtia.Text = Format(Data1.Recordset("alqui"), "Standard")
tretleg.Text = Format(Data1.Recordset("retleg"), "Standard")
timps.Text = Format(Data1.Recordset("impsue"), "Standard")
ttoth.Text = Format(Data1.Recordset("tothab"), "Standard")


End Function

Public Function borrarcam()
tced.Text = ""
tcod.Text = ""
tsex.Text = ""
tn1.Text = ""
tn2.Text = ""
tape1.Text = ""
tape2.Text = ""
tdir.Text = ""
tnp.Text = ""
tcoddep.Text = ""
tloc.Text = ""
tcodpos.Text = ""
mnac.Text = "__/__/____"
ming.Text = "__/__/____"
tden.Text = ""
tcarg.Text = ""
tmj.Text = ""
thd.Text = ""
tcanhd.Text = ""
tcuo.Text = ""
tret.Text = ""
tgtia.Text = ""
tretleg.Text = ""
timps.Text = ""
ttoth.Text = ""
txt_tel.Text = ""
txt_estciv.Text = ""
txt_cedc.Text = ""
txt_codcedc.Text = ""
txt_nomc.Text = ""

End Function

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub

Private Sub ming_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   mnac.SetFocus
End If

End Sub

Private Sub mnac_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   tden.SetFocus
End If

End Sub

Private Sub tape1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   tape2.SetFocus
End If

End Sub

Private Sub tape2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   tdir.SetFocus
End If

End Sub

Private Sub tcanhd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   tcuo.SetFocus
End If

End Sub

Private Sub tcarg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_estciv.SetFocus
End If

End Sub

Private Sub tced_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   tcod.SetFocus
End If

End Sub

Private Sub tcod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   tn1.SetFocus
End If

End Sub

Private Sub tcoddep_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   tloc.SetFocus
End If

End Sub

Private Sub tcodpos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_tel.SetFocus
End If

End Sub

Private Sub tcuo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   tret.SetFocus
End If

End Sub

Private Sub tcuo_LostFocus()
If tcuo.Text <> "" Then
   tcuo.Text = Format(tcuo.Text, "Standard")
End If


End Sub

Private Sub tden_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   tcarg.SetFocus
End If

End Sub

Private Sub tdir_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   tnp.SetFocus
End If

End Sub

Private Sub tgtia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   tretleg.SetFocus
End If

End Sub

Private Sub tgtia_LostFocus()
If tgtia.Text <> "" Then
   tgtia.Text = Format(tgtia.Text, "Standard")
End If

End Sub

Private Sub thd_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   tcanhd.SetFocus
End If

End Sub

Private Sub timps_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   ttoth.SetFocus
End If

End Sub

Private Sub timps_LostFocus()
If timps.Text <> "" Then
   timps.Text = Format(timps.Text, "Standard")
End If

End Sub

Private Sub tloc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   tcodpos.SetFocus
End If

End Sub

Private Sub tmj_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   thd.SetFocus
End If

End Sub

Private Sub tn1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   tn2.SetFocus
End If

End Sub

Private Sub tn2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   tape1.SetFocus
End If

End Sub

Private Sub tnp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   tcoddep.SetFocus
End If

End Sub

Private Sub tret_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   tgtia.SetFocus
End If

End Sub

Private Sub tret_LostFocus()
If tret.Text <> "" Then
   tret.Text = Format(tret.Text, "Standard")
End If

End Sub

Private Sub tretleg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   timps.SetFocus
End If

End Sub

Private Sub tretleg_LostFocus()
If tretleg.Text <> "" Then
   tretleg.Text = Format(tretleg.Text, "Standard")
End If

End Sub

Private Sub ttoth_LostFocus()
If ttoth.Text <> "" Then
   ttoth.Text = Format(ttoth.Text, "Standard")
End If

End Sub

Private Sub txt_cedc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_codcedc.SetFocus
End If

End Sub

Private Sub txt_codcedc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_nomc.SetFocus
End If

End Sub

Private Sub txt_estciv_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txt_cedc.SetFocus
End If

End Sub

Private Sub txt_nomc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   tmj.SetFocus
End If

End Sub

Private Sub txt_tel_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   ming.SetFocus
End If

End Sub
