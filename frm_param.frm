VERSION 5.00
Begin VB.Form frm_param 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros de la empresa"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9495
   Icon            =   "frm_param.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   9495
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data data_ecg 
      Caption         =   "data_ecg"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   2880
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Data data_ui 
      Caption         =   "data_ui"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   3960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5640
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data data_paramhs 
      Caption         =   "data_paramhs"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2160
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7560
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_par2 
      Caption         =   "data_par2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7800
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data data_larga 
      Caption         =   "data_larga"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   960
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6480
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data data_nte 
      Caption         =   "data_nte"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   6840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   6120
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Data data_sur 
      Caption         =   "data_sur"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   5040
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   5640
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data data_par 
      Caption         =   "data_par"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   840
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7920
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Height          =   495
      Left            =   240
      Picture         =   "frm_param.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   7920
      Width           =   495
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Caption         =   "Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7695
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   9015
      Begin VB.Data Data_eval 
         Caption         =   "Data_eval"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   3960
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   6360
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox t_evalua 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   60
         Top             =   7080
         Width           =   1695
      End
      Begin VB.TextBox t_rubcuo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   58
         Top             =   7080
         Width           =   1695
      End
      Begin VB.TextBox t_nrofactcnv 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   56
         Top             =   6600
         Width           =   1695
      End
      Begin VB.TextBox t_codfac 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         TabIndex        =   54
         Top             =   6600
         Width           =   1695
      End
      Begin VB.TextBox t_dolar 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   7080
         TabIndex        =   52
         Top             =   6120
         Width           =   1695
      End
      Begin VB.TextBox t_nropedeco2 
         Height          =   405
         Left            =   2640
         TabIndex        =   50
         Top             =   6000
         Width           =   1695
      End
      Begin VB.Data data_recmys 
         Caption         =   "data_recmys"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   1440
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   600
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox t_nrorecnew 
         Height          =   375
         Left            =   7080
         TabIndex        =   49
         Top             =   5640
         Width           =   1695
      End
      Begin VB.TextBox t_nropedeco 
         Height          =   375
         Left            =   2640
         TabIndex        =   47
         Top             =   5640
         Width           =   1695
      End
      Begin VB.TextBox t_ui 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   7080
         TabIndex        =   45
         Top             =   5160
         Width           =   1695
      End
      Begin VB.TextBox txt_horarios 
         Height          =   375
         Left            =   2640
         TabIndex        =   43
         Top             =   5160
         Width           =   1695
      End
      Begin VB.TextBox t_nromam 
         Height          =   375
         Left            =   7080
         TabIndex        =   41
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Data data_mam 
         Caption         =   "data_mam"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   3960
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox t_accmam 
         Height          =   375
         Left            =   2640
         TabIndex        =   39
         Top             =   4680
         Width           =   1695
      End
      Begin VB.TextBox t_ecg 
         Height          =   375
         Left            =   7080
         TabIndex        =   37
         Top             =   4200
         Width           =   1695
      End
      Begin VB.TextBox t_nroaccadm 
         Height          =   375
         Left            =   2640
         TabIndex        =   35
         ToolTipText     =   "SOLO PARA PC ADMINISTRACION"
         Top             =   4200
         Width           =   1695
      End
      Begin VB.TextBox t_nrollam 
         Height          =   375
         Left            =   7080
         TabIndex        =   32
         ToolTipText     =   "SOLO PARA PC DESPACHO"
         Top             =   3240
         Width           =   1695
      End
      Begin VB.TextBox t_nrollaml 
         Height          =   375
         Left            =   7080
         TabIndex        =   31
         ToolTipText     =   "SOLO PARA PC DESPACHO"
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox t_nrollamn 
         Height          =   375
         Left            =   7080
         TabIndex        =   30
         ToolTipText     =   "SOLO PARA PC DESPACHO"
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox t_nrollams 
         Height          =   375
         Left            =   7080
         TabIndex        =   29
         ToolTipText     =   "SOLO PARA PC DESPACHO"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox t_nrorec 
         Height          =   375
         Left            =   7080
         TabIndex        =   28
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox t_rubcred 
         Height          =   375
         Left            =   7080
         TabIndex        =   27
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox t_rubcdo 
         Height          =   375
         Left            =   7080
         TabIndex        =   20
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox t_nromatenf 
         Height          =   375
         Left            =   7080
         TabIndex        =   18
         Top             =   3720
         Width           =   1695
      End
      Begin VB.TextBox t_nromamind 
         Height          =   375
         Left            =   2640
         TabIndex        =   17
         Top             =   3720
         Width           =   1695
      End
      Begin VB.TextBox t_nropedpsoc 
         Height          =   375
         Left            =   2640
         TabIndex        =   16
         Top             =   3240
         Width           =   1695
      End
      Begin VB.TextBox t_nropedrrhh 
         Height          =   375
         Left            =   2640
         TabIndex        =   15
         Top             =   2760
         Width           =   1695
      End
      Begin VB.TextBox t_nropedmant 
         Height          =   375
         Left            =   2640
         TabIndex        =   14
         Top             =   2280
         Width           =   1695
      End
      Begin VB.TextBox t_nropedinf 
         Height          =   375
         Left            =   2640
         TabIndex        =   13
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox t_nrosoc 
         Height          =   375
         Left            =   2640
         TabIndex        =   12
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox t_nrofaccdo 
         Height          =   375
         Left            =   2640
         TabIndex        =   11
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox t_nrobase 
         Height          =   375
         Left            =   2640
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label29 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nro. Evaluación:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   4680
         TabIndex        =   59
         Top             =   7080
         Width           =   2415
      End
      Begin VB.Label Label28 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Rubro CUOTAS:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   57
         Top             =   7080
         Width           =   2415
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nro.Fact.Convenios:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   4680
         TabIndex        =   55
         Top             =   6600
         Width           =   2415
      End
      Begin VB.Label Label26 
         Caption         =   "Código FACTURACIÓN:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   53
         Top             =   6600
         Width           =   2415
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Valor U$s."
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4680
         TabIndex        =   51
         Top             =   6120
         Width           =   2415
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nro.Recibo nuevo:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4680
         TabIndex        =   48
         Top             =   5640
         Width           =   2415
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nro.Pedido Economato:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   240
         TabIndex        =   46
         Top             =   5640
         Width           =   2415
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFFFF&
         Caption         =   "VALOR DE U.I."
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4680
         TabIndex        =   44
         Top             =   5160
         Width           =   2415
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nro.Horarios:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   42
         Top             =   5160
         Width           =   2415
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nro. MAM:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4680
         TabIndex        =   40
         Top             =   4680
         Width           =   2415
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nro.Acciones MAM:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   38
         Top             =   4680
         Width           =   2415
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ultimo Mat.Enf."
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4680
         TabIndex        =   36
         Top             =   4200
         Width           =   2415
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nro Acciones ADM:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   34
         Top             =   4200
         Width           =   2415
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ultimo llamado general"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4680
         TabIndex        =   26
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ultimo llamado largador"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4680
         TabIndex        =   25
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ultimo llamado NORTE"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4680
         TabIndex        =   24
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ultimo llamado SUR"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4680
         TabIndex        =   23
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ultimo Número de RECIBO:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4680
         TabIndex        =   22
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Rubro Crédito:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4680
         TabIndex        =   21
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Rubro Contado:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4680
         TabIndex        =   19
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ultimo ECG:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   4680
         TabIndex        =   9
         Top             =   3720
         Width           =   2415
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ultimo reg. MAM Ind."
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   3720
         Width           =   2415
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ultimo pedido P.Social"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   3240
         Width           =   2415
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ultimo pedido RRHH"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   2760
         Width           =   2415
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ultimo pedido mant."
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   2280
         Width           =   2415
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ultimo pedido informat."
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ultimo nro de socio:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ultima factura contado/credito"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nro de BASE:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2415
      End
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   5760
      Picture         =   "frm_param.frx":0B14
      Stretch         =   -1  'True
      Top             =   6600
      Width           =   1335
   End
End
Attribute VB_Name = "frm_param"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'On Error GoTo Versihayerr2

data_par.Recordset.Edit
data_par.Recordset("base") = t_nrobase.Text
data_par.Recordset("contado") = t_nrofaccdo.Text
data_par.Recordset("ultimo_soc") = t_nrosoc.Text
data_par.Recordset("musada") = t_nropedinf.Text
data_par.Recordset("cajame") = t_nropedmant.Text
data_par.Recordset("mnueva") = t_nropedrrhh.Text
data_par.Recordset("compromiso") = t_nropedpsoc.Text
data_par.Recordset("ctacontado") = t_nromamind.Text
data_par.Recordset("srvcnt") = t_rubcdo.Text
data_par.Recordset("srvcrd") = t_rubcred.Text
data_par.Recordset("nrorec") = t_nrorec.Text
data_par.Recordset("tasaintmn") = t_nrollam.Text
data_par.Recordset("rub_cuotas") = t_nropedeco.Text
data_par.Recordset("notadev") = t_nropedeco2.Text
data_par.Recordset("contacto") = t_codfac.Text
data_par.Recordset("cbrcuo") = t_rubcuo.Text
data_par.Recordset.Update

data_par2.Recordset.Edit
data_par2.Recordset("nro_accadm") = t_nroaccadm.Text
data_par2.Recordset("nro_material") = t_nromatenf.Text
data_par2.Recordset("nro_informat") = t_ecg.Text
data_par2.Recordset.Update

Data_eval.Recordset.Edit
Data_eval.Recordset("nro_eval") = t_evalua.Text
Data_eval.Recordset("base") = t_nrobase.Text
Data_eval.Recordset.Update


If data_recmys.Recordset("nro_rec") = t_nrorecnew.Text Then
Else
   data_recmys.Recordset.Edit
   data_recmys.Recordset("nro_rec") = t_nrorecnew.Text
   data_recmys.Recordset.Update
End If
If t_nrofactcnv.Text <> "" Then
   If IsNull(data_recmys.Recordset("varios")) = False Then
      If t_nrofactcnv.Text = data_recmys.Recordset("varios") Then
      Else
         data_recmys.Recordset.Edit
         data_recmys.Recordset("varios") = t_nrofactcnv.Text
         data_recmys.Recordset.Update
      End If
   Else
      data_recmys.Recordset.Edit
      data_recmys.Recordset("varios") = t_nrofactcnv.Text
      data_recmys.Recordset.Update
   End If
End If

data_mam.Recordset.Edit
data_mam.Recordset("base") = t_nrobase.Text
data_mam.Recordset("nro_mam2") = t_accmam.Text
data_mam.Recordset("nro_mam1") = t_nromam.Text
data_mam.Recordset.Update

data_paramhs.Recordset.Edit
data_paramhs.Recordset("base") = t_nrobase.Text
data_paramhs.Recordset("nro_reg") = txt_horarios.Text
data_paramhs.Recordset.Update

data_sur.Recordset.Edit
data_sur.Recordset("tasaintmn") = t_nrollams.Text
data_sur.Recordset.Update

data_nte.Recordset.Edit
data_nte.Recordset("tasaintmn") = t_nrollamn.Text
data_nte.Recordset.Update

data_larga.Recordset.Edit
data_larga.Recordset("tasaintmn") = t_nrollaml.Text
data_larga.Recordset.Update

If t_ui.Text <> "" Then
   If data_ui.Recordset("descrip") <> t_ui.Text Then
        data_ui.Recordset.Edit
        data_ui.Recordset("descrip") = t_ui.Text
        data_ui.Recordset("fecha") = Date
        data_ui.Recordset.Update
   End If
End If
If t_dolar.Text <> "" Then
   If data_ui.Recordset("hora") <> t_dolar.Text Then
        data_ui.Recordset.Edit
        data_ui.Recordset("hora") = t_dolar.Text
        data_ui.Recordset("fecha") = Date
        data_ui.Recordset.Update
   End If
End If

MsgBox "Datos grabados correctamente"


'Exit Sub

'Versihayerr2:
'            If Err.Number = 3155 Then
'               MsgBox "Error al grabar algún dato en nulo"
'            Else
'               MsgBox "Error al grabar algún campo en nulo"
'            End If

End Sub

Private Sub Form_Load()
'On Error GoTo Versihayerr

data_par.DatabaseName = App.path & "\parse.mdb"
data_par.RecordSource = "parsec0"
data_par.Refresh

data_sur.DatabaseName = App.path & "\sur\parse.mdb"
data_sur.RecordSource = "parsec0"
data_sur.Refresh

data_nte.DatabaseName = App.path & "\norte\parse.mdb"
data_nte.RecordSource = "parsec0"
data_nte.Refresh

data_larga.DatabaseName = App.path & "\largador\parse.mdb"
data_larga.RecordSource = "parsec0"
data_larga.Refresh

data_paramhs.DatabaseName = App.path & "\paramhoras.mdb"
data_paramhs.RecordSource = "parsec0"
data_paramhs.Refresh

Data_eval.DatabaseName = App.path & "\parevalu.mdb"
Data_eval.RecordSource = "parevalu"
Data_eval.Refresh

data_recmys.Connect = "odbc;dsn=sappfact;"
data_recmys.RecordSource = "paramsapp"
data_recmys.Refresh

If IsNull(data_recmys.Recordset("varios")) = False Then
   t_nrofactcnv.Text = data_recmys.Recordset("varios")
End If

data_ui.Connect = "ODBC;DSN=sappnew;"
data_ui.RecordSource = "hc_frecresp"
data_ui.Refresh
If data_ui.Recordset.RecordCount > 0 Then
   t_ui.Text = data_ui.Recordset("descrip")
   t_dolar.Text = data_ui.Recordset("hora")
Else
   data_ui.Recordset.AddNew
   data_ui.Recordset("descrip") = t_ui.Text
   data_ui.Recordset("fecha") = Date
   data_ui.Recordset.Update
   data_ui.Refresh
   t_ui.Text = data_ui.Recordset("descrip")
   t_dolar.Text = 0
End If

data_par2.DatabaseName = App.path & "\paramb.mdb"
data_par2.RecordSource = "paramb"
data_par2.Refresh

data_mam.DatabaseName = App.path & "\paramam.mdb"
data_mam.RecordSource = "parsec0"
data_mam.Refresh

t_accmam.Text = data_mam.Recordset("nro_mam2")
t_nromam.Text = data_mam.Recordset("nro_mam1")

t_evalua.Text = Data_eval.Recordset("nro_eval")

t_nrorecnew.Text = data_recmys.Recordset("nro_rec")

t_nropedeco.Text = data_par.Recordset("rub_cuotas")
t_nropedeco2.Text = data_par.Recordset("notadev")
If IsNull(data_par.Recordset("contacto")) = False Then
   t_codfac.Text = data_par.Recordset("contacto")
Else
   t_codfac.Text = ""
End If

t_nrobase.Text = data_par.Recordset("base")
t_nrofaccdo.Text = data_par.Recordset("contado")
t_nrosoc.Text = data_par.Recordset("ultimo_soc")
t_nropedinf.Text = data_par.Recordset("musada")
t_nropedmant.Text = data_par.Recordset("cajame")
t_nropedrrhh.Text = data_par.Recordset("mnueva")
t_nropedpsoc.Text = data_par.Recordset("compromiso")
t_nromamind.Text = data_par.Recordset("ctacontado")
t_nroaccadm.Text = data_par2.Recordset("nro_accadm")
t_rubcdo.Text = data_par.Recordset("srvcnt")
t_rubcred.Text = data_par.Recordset("srvcrd")
t_rubcuo.Text = data_par.Recordset("cbrcuo")
t_nrorec.Text = data_par.Recordset("nrorec")
t_nrollams.Text = data_sur.Recordset("tasaintmn")
t_nrollamn.Text = data_nte.Recordset("tasaintmn")
t_nrollaml.Text = data_larga.Recordset("tasaintmn")
t_nrollam.Text = data_par.Recordset("tasaintmn")
t_nromatenf.Text = data_par2.Recordset("nro_material")
t_ecg.Text = data_par2.Recordset("nro_informat")
txt_horarios.Text = data_paramhs.Recordset("nro_reg")


'Exit Sub

'Versihayerr:
'            If Err.Number = 3155 Then
'               MsgBox "Error al cargar algún dato en nulo"
'            Else
'               MsgBox "Error al cargar algún  campo en nulo"
'            End If
            
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Form_Resize()
With Image1
     .Left = 0
     .Top = 0
     .Height = Me.Height
     .Width = Me.Width
End With

End Sub

