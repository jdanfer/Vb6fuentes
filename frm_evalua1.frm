VERSION 5.00
Begin VB.Form frm_evalua1 
   BackColor       =   &H00FF0000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Evaluación de Personal"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14400
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frm_evalua1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   14400
   StartUpPosition =   2  'CenterScreen
   Begin VB.Data data_per 
      Caption         =   "data_per"
      Connect         =   "odbc;dsn=sappper"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   7920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   7320
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton b_firma 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Firmar evaluación"
      Height          =   495
      Left            =   12120
      Picture         =   "frm_evalua1.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   8160
      Width           =   2175
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
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
      Top             =   720
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FF8080&
      Height          =   6495
      Left            =   13320
      TabIndex        =   32
      Top             =   1680
      Width           =   975
      Begin VB.CommandButton b_cierra 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         Picture         =   "frm_evalua1.frx":0E54
         Style           =   1  'Graphical
         TabIndex        =   65
         ToolTipText     =   "Cerrar Evaluación."
         Top             =   3360
         Width           =   495
      End
      Begin VB.CommandButton b_imp 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         Picture         =   "frm_evalua1.frx":13DE
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Informe de evaluación"
         Top             =   2520
         Width           =   495
      End
      Begin VB.CommandButton b_cance 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   495
         Left            =   240
         Picture         =   "frm_evalua1.frx":1968
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   1800
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.CommandButton b_graba 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   495
         Left            =   240
         Picture         =   "frm_evalua1.frx":1EF2
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton b_edita 
         BackColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         Picture         =   "frm_evalua1.frx":247C
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Comenzar evaluación"
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label4 
         BackColor       =   &H0080FF80&
         Caption         =   "Inicio"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   64
         Top             =   120
         Width           =   495
      End
   End
   Begin VB.Data data_periodo 
      Caption         =   "data_periodo"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   11520
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data data_texto 
      Caption         =   "data_texto"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   8760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   840
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Data data_titulos 
      Caption         =   "data_titulos"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   495
      Left            =   10320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   600
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FF8080&
      Caption         =   "Evaluación"
      Enabled         =   0   'False
      Height          =   6495
      Left            =   120
      TabIndex        =   9
      Top             =   1680
      Width           =   13215
      Begin VB.TextBox Text8 
         Height          =   510
         Left            =   1800
         TabIndex        =   63
         Top             =   5280
         Visible         =   0   'False
         Width           =   6615
      End
      Begin VB.TextBox Text7 
         Height          =   495
         Left            =   1800
         TabIndex        =   62
         Top             =   4800
         Visible         =   0   'False
         Width           =   6615
      End
      Begin VB.TextBox Text6 
         Height          =   495
         Left            =   1800
         TabIndex        =   61
         Top             =   4200
         Visible         =   0   'False
         Width           =   6615
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   1800
         TabIndex        =   60
         Top             =   3600
         Visible         =   0   'False
         Width           =   6615
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   1800
         TabIndex        =   59
         Top             =   3000
         Visible         =   0   'False
         Width           =   6615
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   1800
         TabIndex        =   58
         Top             =   2400
         Visible         =   0   'False
         Width           =   6615
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   1800
         TabIndex        =   57
         Top             =   1800
         Visible         =   0   'False
         Width           =   6615
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   1800
         TabIndex        =   56
         Top             =   1200
         Visible         =   0   'False
         Width           =   6615
      End
      Begin VB.CommandButton b_8 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   12600
         Picture         =   "frm_evalua1.frx":2A06
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Escribir observaciones para ésta pregunta"
         Top             =   5880
         Width           =   495
      End
      Begin VB.CommandButton b_7 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   12600
         Picture         =   "frm_evalua1.frx":2F90
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Escribir observaciones para ésta pregunta"
         Top             =   5160
         Width           =   495
      End
      Begin VB.CommandButton b_6 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   12600
         Picture         =   "frm_evalua1.frx":351A
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Escribir observaciones para ésta pregunta"
         Top             =   4440
         Width           =   495
      End
      Begin VB.CommandButton b_5 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   12600
         Picture         =   "frm_evalua1.frx":3AA4
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Escribir observaciones para ésta pregunta"
         Top             =   3720
         Width           =   495
      End
      Begin VB.CommandButton b_4 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   12600
         Picture         =   "frm_evalua1.frx":402E
         Style           =   1  'Graphical
         TabIndex        =   51
         ToolTipText     =   "Escribir observaciones para ésta pregunta"
         Top             =   3000
         Width           =   495
      End
      Begin VB.CommandButton b_3 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   12600
         Picture         =   "frm_evalua1.frx":45B8
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Escribir observaciones para ésta pregunta"
         Top             =   2280
         Width           =   495
      End
      Begin VB.CommandButton b_2 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   12600
         Picture         =   "frm_evalua1.frx":4B42
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Escribir observaciones para ésta pregunta"
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton b_1 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   12600
         Picture         =   "frm_evalua1.frx":50CC
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Escribir observaciones para ésta pregunta"
         Top             =   840
         Width           =   495
      End
      Begin VB.TextBox t_8 
         Height          =   390
         Left            =   1920
         TabIndex        =   43
         Top             =   5160
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox t_7 
         Height          =   390
         Left            =   1560
         TabIndex        =   42
         Top             =   4560
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox t_6 
         Height          =   390
         Left            =   1320
         TabIndex        =   41
         Top             =   3960
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox t_5 
         Height          =   390
         Left            =   1200
         TabIndex        =   40
         Top             =   3360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox t_4 
         Height          =   390
         Left            =   840
         TabIndex        =   39
         Top             =   2760
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox t_3 
         Height          =   390
         Left            =   840
         TabIndex        =   38
         Top             =   2160
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox t_2 
         Height          =   390
         Left            =   840
         TabIndex        =   37
         Top             =   1560
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox t_1 
         Height          =   390
         Left            =   720
         TabIndex        =   36
         Top             =   960
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox cbop8 
         Height          =   390
         Left            =   11520
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   5880
         Width           =   855
      End
      Begin VB.ComboBox cbop7 
         Height          =   390
         Left            =   11520
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   5160
         Width           =   855
      End
      Begin VB.ComboBox cbop6 
         Height          =   390
         Left            =   11520
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   4440
         Width           =   855
      End
      Begin VB.ComboBox cbop5 
         Height          =   390
         Left            =   11520
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   3720
         Width           =   855
      End
      Begin VB.ComboBox cbop4 
         Height          =   390
         Left            =   11520
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   3000
         Width           =   855
      End
      Begin VB.ComboBox cbop3 
         Height          =   390
         Left            =   11520
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   2280
         Width           =   855
      End
      Begin VB.ComboBox cbop2 
         Height          =   390
         Left            =   11520
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   1560
         Width           =   855
      End
      Begin VB.ComboBox cbop1 
         Height          =   390
         ItemData        =   "frm_evalua1.frx":5656
         Left            =   11520
         List            =   "frm_evalua1.frx":5658
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   840
         Width           =   855
      End
      Begin VB.CommandButton b_select 
         Height          =   495
         Left            =   9360
         Picture         =   "frm_evalua1.frx":565A
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox t_id 
         Height          =   390
         Left            =   10440
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cbotitulos 
         Height          =   390
         ItemData        =   "frm_evalua1.frx":5BE4
         Left            =   1680
         List            =   "frm_evalua1.frx":5BE6
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   360
         Width           =   7455
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H00808000&
         Caption         =   "Puntaje"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   11400
         TabIndex        =   21
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lab8 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   120
         TabIndex        =   20
         Top             =   5880
         Width           =   11295
      End
      Begin VB.Label lab7 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   120
         TabIndex        =   19
         Top             =   5160
         Width           =   11295
      End
      Begin VB.Label lab6 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   120
         TabIndex        =   18
         Top             =   4440
         Width           =   11295
      End
      Begin VB.Label lab5 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   120
         TabIndex        =   17
         Top             =   3720
         Width           =   11295
      End
      Begin VB.Label lab4 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   120
         TabIndex        =   16
         Top             =   3000
         Width           =   11295
      End
      Begin VB.Label lab3 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   11295
      End
      Begin VB.Label lab2 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   11295
      End
      Begin VB.Label lab1 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   11295
      End
      Begin VB.Label Label11 
         BackColor       =   &H00808000&
         Caption         =   "Título:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF8080&
      Caption         =   "Datos del funcionario"
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   14175
      Begin VB.ComboBox cboperio 
         Height          =   390
         ItemData        =   "frm_evalua1.frx":5BE8
         Left            =   9360
         List            =   "frm_evalua1.frx":5BEA
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   720
         Width           =   2175
      End
      Begin VB.Data data_evalua 
         Caption         =   "data_evalua"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   10680
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   -120
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label labfec 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   8
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label7 
         BackColor       =   &H00808000&
         Caption         =   "Período:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7080
         TabIndex        =   7
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label6 
         BackColor       =   &H00808000&
         Caption         =   "Fecha Actual:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label labnomj 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   9360
         TabIndex        =   5
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label labnome 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2640
         TabIndex        =   4
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00808000&
         Caption         =   "Nombre del evaluador:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7080
         TabIndex        =   3
         Top             =   360
         Width           =   2295
      End
      Begin VB.Label Label2 
         BackColor       =   &H00808000&
         Caption         =   "Nombre del Empleado:"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Label labnota 
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   120
      TabIndex        =   47
      Top             =   8160
      Width           =   11775
   End
   Begin VB.Label labnro 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   11760
      TabIndex        =   25
      Top             =   720
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   "Evaluación de Desempeño"
      BeginProperty Font 
         Name            =   "Baskerville Old Face"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   0
      Width           =   9135
   End
   Begin VB.Image Image1 
      Height          =   1575
      Left            =   600
      Picture         =   "frm_evalua1.frx":5BEC
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "frm_evalua1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Wxtitu As Integer

Private Sub b_alta_Click()



End Sub

Private Sub b_1_Click()
Xquecol = 1
frm_obseval.Show vbModal


End Sub

Private Sub b_2_Click()
Xquecol = 2
frm_obseval.Show vbModal

End Sub

Private Sub b_3_Click()
Xquecol = 3
frm_obseval.Show vbModal

End Sub

Private Sub b_4_Click()
Xquecol = 4
frm_obseval.Show vbModal

End Sub

Private Sub b_5_Click()
Xquecol = 5
frm_obseval.Show vbModal

End Sub

Private Sub b_6_Click()
Xquecol = 6
frm_obseval.Show vbModal

End Sub

Private Sub b_7_Click()
Xquecol = 7
frm_obseval.Show vbModal

End Sub

Private Sub b_8_Click()
Xquecol = 8
frm_obseval.Show vbModal

End Sub

Private Sub b_cance_Click()
cbop1.ListIndex = -1
cbop2.ListIndex = -1
cbop3.ListIndex = -1
cbop4.ListIndex = -1
cbop5.ListIndex = -1
cbop6.ListIndex = -1
cbop7.ListIndex = -1
cbop8.ListIndex = -1
cbotitulos.ListIndex = -1

b_graba.Enabled = False
b_cance.Enabled = False
Frame2.Enabled = False
b_firma.Enabled = True
b_edita.Enabled = True
b_imp.Enabled = True

End Sub

Private Sub b_cierra_Click()
Dim Xfirmaono, Xclaveeva As String
Dim Xcanfirma As Integer

On Error GoTo Errevalfirma

Xcanfirma = 0

b_firma.Enabled = False
b_cierra.Enabled = False

Xfirmaono = MsgBox("Desea cerrar la evaluación de: " & labnome.Caption & " del período: " & cboperio.Text & "?", vbInformation + vbYesNo)
If Xfirmaono = vbYes Then
   Xclaveeva = InputBox("INGRESE SU CLAVE DEL SISTEMA PARA FIRMAR")
   If Xclaveeva = "" Then
      MsgBox "No ingresó clave, no se guardará la información"
   Else
      If WxclaveU = Xclaveeva Then
         frm_evalua1.MousePointer = 11
         data_evalua.RecordSource = "Select * from evaluas where idempl =" & Wxelnrocedev & " and periodo ='" & cboperio.Text & "'"
         data_evalua.Refresh
         If data_evalua.Recordset.RecordCount > 0 Then
            data_evalua.Recordset.MoveLast
            If data_evalua.Recordset.RecordCount >= 64 Then
               data_evalua.Recordset.MoveFirst
               Do While Not data_evalua.Recordset.EOF
                  If IsNull(data_evalua.Recordset("cierre")) = False Then
                     If data_evalua.Recordset("cierre") <> "SI" Then
                        data_evalua.Recordset.Edit
                        data_evalua.Recordset("cierre") = "SI"
                        data_evalua.Recordset("fecha_cierre") = Date
                        data_evalua.Recordset("hora") = Format(Time, "HH:mm")
                        data_evalua.Recordset.Update
                        Xcanfirma = Xcanfirma + 1
                     End If
                  Else
                     data_evalua.Recordset.Edit
                     data_evalua.Recordset("cierre") = "SI"
                     data_evalua.Recordset("fecha_cierre") = Date
                     data_evalua.Recordset("hora") = Format(Time, "HH:mm")
                     data_evalua.Recordset.Update
                     Xcanfirma = Xcanfirma + 1
                  End If
                  data_evalua.Recordset.MoveNext
               Loop
               frm_evalua1.MousePointer = 0
               If Xcanfirma = 0 Then
                  MsgBox "LA EVALUACIÓN YA ESTABA CERRADA!!", vbInformation
               Else
                  MsgBox "Evaluación cerrada CORRECTAMENTE!", vbInformation
               End If
            Else
               frm_evalua1.MousePointer = 0
               MsgBox "No están completas todas las preguntas, VERIFIQUE!"
            End If
         Else
            frm_evalua1.MousePointer = 0
            MsgBox "No se encuentran contestadas todas las preguntas, no se puede cerrar. VERIFIQUE!!", vbInformation
         End If
      Else
         MsgBox "CLAVE INCORRECTA! REINTENTE.", vbExclamation
      End If
   End If
End If

b_firma.Enabled = True
b_cierra.Enabled = True

Exit Sub

Errevalfirma:
             If Err.Number = 3155 Then
                MsgBox "Error al firmar"
                b_firma.Enabled = True
             Else
                MsgBox "Error al firmar, verifique si completó todas las preguntas"
                b_firma.Enabled = True
             End If

End Sub

Private Sub b_edita_Click()

b_graba.Enabled = True
b_cance.Enabled = True
Frame2.Enabled = True
b_firma.Enabled = False
b_imp.Enabled = False

'data_evalua.RecordSource = "evaluas"
'data_evalua.Refresh
cbotitulos.SetFocus

End Sub

Private Sub b_firma_Click()
Dim Xfirmaono, Xclaveeva As String
Dim Xcanfirma As Integer

On Error GoTo Errevalfirma

Xcanfirma = 0

b_firma.Enabled = False

Xfirmaono = MsgBox("Desea firmar la evaluación de: " & labnome.Caption & " del período: " & cboperio.Text & "?", vbInformation + vbYesNo)
If Xfirmaono = vbYes Then
   Xclaveeva = InputBox("INGRESE SU CLAVE DEL SISTEMA PARA FIRMAR")
   If Xclaveeva = "" Then
      MsgBox "No ingresó clave, no se guardará la información"
   Else
      If WxclaveU = Xclaveeva Then
         data_evalua.RecordSource = "Select * from evaluas where idempl =" & Wxelnrocedev & " and periodo ='" & cboperio.Text & "' and idjefe =" & Wxeljefeid
         data_evalua.Refresh
         If data_evalua.Recordset.RecordCount > 0 Then
            data_evalua.Recordset.MoveLast
            If data_evalua.Recordset.RecordCount >= 32 Then
               data_evalua.Recordset.MoveFirst
               Do While Not data_evalua.Recordset.EOF
                  If IsNull(data_evalua.Recordset("firma")) = False Then
                     If data_evalua.Recordset("firma") <> 5 Then
                        data_evalua.Recordset.Edit
                        data_evalua.Recordset("firma") = 5
                        data_evalua.Recordset("usuario") = WElusuario
                        data_evalua.Recordset.Update
                        Xcanfirma = Xcanfirma + 1
                     End If
                  Else
                     data_evalua.Recordset.Edit
                     data_evalua.Recordset("firma") = 5
                     data_evalua.Recordset("usuario") = WElusuario
                     data_evalua.Recordset.Update
                     Xcanfirma = Xcanfirma + 1
                  End If
                  data_evalua.Recordset.MoveNext
               Loop
               If Xcanfirma = 0 Then
                  MsgBox "LA EVALUACIÓN YA FUE FIRMADA ANTERIORMENTE!!", vbCritical
               Else
                  MsgBox "Evaluación firmada CORRECTAMENTE!", vbInformation
               End If
            Else
               MsgBox "No están completas todas las preguntas, VERIFIQUE!"
            End If
         Else
            MsgBox "No se encuentran evaluaciones ingresadas, VERIFIQUE!!"
         End If
      Else
         MsgBox "CLAVE INCORRECTA! REINTENTE.", vbExclamation
      End If
   End If
End If

b_firma.Enabled = True

Exit Sub

Errevalfirma:
             If Err.Number = 3155 Then
                MsgBox "Error al firmar"
                b_firma.Enabled = True
             Else
                MsgBox "Error al firmar, verifique si completó todas las preguntas"
                b_firma.Enabled = True
             End If
             
End Sub

Private Sub b_graba_Click()
Dim x, Xsigrabo As Integer

On Error GoTo Errevalgrab2

Xsigrabo = 0

b_graba.Enabled = False

If cbop1.ListIndex >= 0 And cbop2.ListIndex >= 0 And cbop3.ListIndex >= 0 And _
   cbop4.ListIndex >= 0 And cbop5.ListIndex >= 0 And cbop6.ListIndex >= 0 And _
   cbop7.ListIndex >= 0 And cbop8.ListIndex >= 0 Then

   data_evalua.RecordSource = "Select * from evaluas where idempl =" & Wxelnrocedev & " and periodo ='" & cboperio.Text & "' and idtitulo =" & Wxtitu & " and idjefe =" & Wxeljefeid & " and id2 =" & Wxelnroid2 & " order by idpregun"
   data_evalua.Refresh
   If data_evalua.Recordset.RecordCount > 0 Then
      data_evalua.Recordset.MoveLast
      data_evalua.Recordset.MoveFirst
      If IsNull(data_evalua.Recordset("firma")) = True Then
         For x = 1 To 8
            data_evalua.Recordset.Edit
            If IsNull(data_evalua.Recordset("fechamod")) = False Then
               If Format(data_evalua.Recordset("fechamod"), "dd-mm-yyyy") <> Format(Date, "dd-mm-yyyy") Then
                  data_evalua.Recordset("fechamod") = Format(Date, "dd-mm-yyyy")
                  Xsigrabo = 8
               End If
            Else
               data_evalua.Recordset("fechamod") = Format(Date, "dd-mm-yyyy")
               Xsigrabo = 8
            End If
            If data_evalua.Recordset("idjefe") <> Wxeljefeid Then
               data_evalua.Recordset("idjefe") = Wxeljefeid
               Xsigrabo = 8
            End If
            If x = 1 Then
               If data_evalua.Recordset("idpregun") <> t_1.Text Then
                  data_evalua.Recordset("idpregun") = t_1.Text
                  Xsigrabo = 8
               End If
               If Trim(Text1.Text) <> "" Then
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     If Trim(data_evalua.Recordset("obs")) <> Trim(Text1.Text) Then
                        data_evalua.Recordset("obs") = Text1.Text
                        Xsigrabo = 8
                     End If
                  Else
                     If Trim(Text1.Text) <> "" Then
                        data_evalua.Recordset("obs") = Text1.Text
                        Xsigrabo = 8
                     End If
                  End If
               Else
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     data_evalua.Recordset("obs") = Null
                     Xsigrabo = 8
                  End If
               End If
            End If
            If x = 2 Then
               If data_evalua.Recordset("idpregun") <> t_2.Text Then
                  data_evalua.Recordset("idpregun") = t_2.Text
                  Xsigrabo = 8
               End If
               If Trim(Text2.Text) <> "" Then
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     If Trim(data_evalua.Recordset("obs")) <> Trim(Text2.Text) Then
                        data_evalua.Recordset("obs") = Text2.Text
                        Xsigrabo = 8
                     End If
                  Else
                     If Trim(Text2.Text) <> "" Then
                        data_evalua.Recordset("obs") = Text2.Text
                        Xsigrabo = 8
                     End If
                  End If
               Else
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     data_evalua.Recordset("obs") = Null
                     Xsigrabo = 8
                  End If
               End If
            End If
            If x = 3 Then
               If data_evalua.Recordset("idpregun") <> t_3.Text Then
                  data_evalua.Recordset("idpregun") = t_3.Text
                  Xsigrabo = 8
               End If
               If Trim(Text3.Text) <> "" Then
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     If Trim(data_evalua.Recordset("obs")) <> Trim(Text3.Text) Then
                        data_evalua.Recordset("obs") = Text3.Text
                        Xsigrabo = 8
                     End If
                  Else
                     If Trim(Text3.Text) <> "" Then
                        data_evalua.Recordset("obs") = Text3.Text
                        Xsigrabo = 8
                     End If
                  End If
               Else
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     data_evalua.Recordset("obs") = Null
                     Xsigrabo = 8
                  End If
               End If
            End If
            If x = 4 Then
               If data_evalua.Recordset("idpregun") <> t_4.Text Then
                  data_evalua.Recordset("idpregun") = t_4.Text
                  Xsigrabo = 8
               End If
               If Trim(Text4.Text) <> "" Then
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     If Trim(data_evalua.Recordset("obs")) <> Trim(Text4.Text) Then
                        data_evalua.Recordset("obs") = Text4.Text
                        Xsigrabo = 8
                     End If
                  Else
                     If Trim(Text4.Text) <> "" Then
                        data_evalua.Recordset("obs") = Text4.Text
                        Xsigrabo = 8
                     End If
                  End If
               Else
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     data_evalua.Recordset("obs") = Null
                     Xsigrabo = 8
                  End If
               End If
            End If
            If x = 5 Then
               If data_evalua.Recordset("idpregun") <> t_5.Text Then
                  data_evalua.Recordset("idpregun") = t_5.Text
                  Xsigrabo = 8
               End If
               If Trim(Text5.Text) <> "" Then
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     If Trim(data_evalua.Recordset("obs")) <> Trim(Text5.Text) Then
                        data_evalua.Recordset("obs") = Text5.Text
                        Xsigrabo = 8
                     End If
                  Else
                     If Trim(Text5.Text) <> "" Then
                        data_evalua.Recordset("obs") = Text5.Text
                        Xsigrabo = 8
                     End If
                  End If
               Else
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     data_evalua.Recordset("obs") = Null
                     Xsigrabo = 8
                  End If
               End If
            End If
            If x = 6 Then
               If data_evalua.Recordset("idpregun") <> t_6.Text Then
                  data_evalua.Recordset("idpregun") = t_6.Text
                  Xsigrabo = 8
               End If
               If Trim(Text6.Text) <> "" Then
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     If Trim(data_evalua.Recordset("obs")) <> Trim(Text6.Text) Then
                        data_evalua.Recordset("obs") = Text6.Text
                        Xsigrabo = 8
                     End If
                  Else
                     If Trim(Text6.Text) <> "" Then
                        data_evalua.Recordset("obs") = Text6.Text
                        Xsigrabo = 8
                     End If
                  End If
               Else
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     data_evalua.Recordset("obs") = Null
                     Xsigrabo = 8
                  End If
               End If
            End If
            If x = 7 Then
               If data_evalua.Recordset("idpregun") <> t_7.Text Then
                  data_evalua.Recordset("idpregun") = t_7.Text
                  Xsigrabo = 8
               End If
               If Trim(Text7.Text) <> "" Then
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     If Trim(data_evalua.Recordset("obs")) <> Trim(Text7.Text) Then
                        data_evalua.Recordset("obs") = Text7.Text
                        Xsigrabo = 8
                     End If
                  Else
                     If Trim(Text7.Text) Then
                        data_evalua.Recordset("obs") = Text7.Text
                        Xsigrabo = 8
                     End If
                  End If
               Else
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     data_evalua.Recordset("obs") = Null
                     Xsigrabo = 8
                  End If
               End If
            End If
            If x = 8 Then
               If data_evalua.Recordset("idpregun") <> t_8.Text Then
                  data_evalua.Recordset("idpregun") = t_8.Text
                  Xsigrabo = 8
               End If
               If Trim(Text8.Text) <> "" Then
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     If Trim(data_evalua.Recordset("obs")) <> Trim(Text8.Text) Then
                        data_evalua.Recordset("obs") = Text8.Text
                        Xsigrabo = 8
                     End If
                  Else
                     If Trim(Text8.Text) <> "" Then
                        data_evalua.Recordset("obs") = Text8.Text
                        Xsigrabo = 8
                     End If
                  End If
               Else
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     data_evalua.Recordset("obs") = Null
                     Xsigrabo = 8
                  End If
               End If
            End If
            If data_evalua.Recordset("periodo") <> cboperio.Text Then
               data_evalua.Recordset("periodo") = cboperio.Text
               Xsigrabo = 8
            End If
            If x = 1 Then
               If data_evalua.Recordset("puntos") <> Val(cbop1.Text) Then
                  data_evalua.Recordset("puntos") = Val(cbop1.Text)
                  Xsigrabo = 8
               End If
            End If
            If x = 2 Then
               If data_evalua.Recordset("puntos") <> Val(cbop2.Text) Then
                  data_evalua.Recordset("puntos") = Val(cbop2.Text)
                  Xsigrabo = 8
               End If
            End If
            If x = 3 Then
               If data_evalua.Recordset("puntos") <> Val(cbop3.Text) Then
                  data_evalua.Recordset("puntos") = Val(cbop3.Text)
                  Xsigrabo = 8
               End If
            End If
            If x = 4 Then
               If data_evalua.Recordset("puntos") <> Val(cbop4.Text) Then
                  data_evalua.Recordset("puntos") = Val(cbop4.Text)
                  Xsigrabo = 8
               End If
            End If
            If x = 5 Then
               If data_evalua.Recordset("puntos") <> Val(cbop5.Text) Then
                  data_evalua.Recordset("puntos") = Val(cbop5.Text)
                  Xsigrabo = 8
               End If
            End If
            If x = 6 Then
               If data_evalua.Recordset("puntos") <> Val(cbop6.Text) Then
                  data_evalua.Recordset("puntos") = Val(cbop6.Text)
                  Xsigrabo = 8
               End If
            End If
            If x = 7 Then
               If data_evalua.Recordset("puntos") <> Val(cbop7.Text) Then
                  data_evalua.Recordset("puntos") = Val(cbop7.Text)
                  Xsigrabo = 8
               End If
            End If
            If x = 8 Then
               If data_evalua.Recordset("puntos") <> Val(cbop8.Text) Then
                  data_evalua.Recordset("puntos") = Val(cbop8.Text)
                  Xsigrabo = 8
               End If
            End If
            If Xsigrabo = 8 Then
               data_evalua.Recordset.Update
               Xsigrabo = 0
            Else
               data_evalua.Recordset.CancelUpdate
            End If
            data_evalua.Recordset.MoveNext
        Next
        cbop1.ListIndex = -1
        cbop2.ListIndex = -1
        cbop3.ListIndex = -1
        cbop4.ListIndex = -1
        cbop5.ListIndex = -1
        cbop6.ListIndex = -1
        cbop7.ListIndex = -1
        cbop8.ListIndex = -1
        cbotitulos.ListIndex = -1
        b_graba.Enabled = False
        b_cance.Enabled = False
        b_edita.Enabled = True
        Frame2.Enabled = False
        b_imp.Enabled = True
        b_firma.Enabled = True
     Else
        MsgBox "Solo se modificarán Observaciones porque la evaluación está firmada"
        For x = 1 To 8
            data_evalua.Recordset.Edit
            If x = 1 Then
               If Text1.Text <> "" Then
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     If data_evalua.Recordset("obs") <> Text1.Text Then
                        data_evalua.Recordset("obs") = Text1.Text
                        Xsigrabo = 8
                     End If
                  Else
                     data_evalua.Recordset("obs") = Text1.Text
                     Xsigrabo = 8
                  End If
               Else
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     data_evalua.Recordset("obs") = Null
                     Xsigrabo = 8
                  End If
               End If
            End If
            If x = 2 Then
               If Text2.Text <> "" Then
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     If data_evalua.Recordset("obs") <> Text2.Text Then
                        data_evalua.Recordset("obs") = Text2.Text
                        Xsigrabo = 8
                     End If
                  Else
                     data_evalua.Recordset("obs") = Text2.Text
                     Xsigrabo = 8
                  End If
               Else
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     data_evalua.Recordset("obs") = Null
                     Xsigrabo = 8
                  End If
               End If
            End If
            If x = 3 Then
               If Text3.Text <> "" Then
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     If data_evalua.Recordset("obs") <> Text3.Text Then
                        data_evalua.Recordset("obs") = Text3.Text
                        Xsigrabo = 8
                     End If
                  Else
                     data_evalua.Recordset("obs") = Text3.Text
                     Xsigrabo = 8
                  End If
               Else
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     data_evalua.Recordset("obs") = Null
                     Xsigrabo = 8
                  End If
               End If
            End If
            If x = 4 Then
               If Text4.Text <> "" Then
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     If data_evalua.Recordset("obs") <> Text4.Text Then
                        data_evalua.Recordset("obs") = Text4.Text
                        Xsigrabo = 8
                     End If
                  Else
                     data_evalua.Recordset("obs") = Text4.Text
                     Xsigrabo = 8
                  End If
               Else
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     data_evalua.Recordset("obs") = Null
                     Xsigrabo = 8
                  End If
               End If
            End If
            If x = 5 Then
               If Text5.Text <> "" Then
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     If data_evalua.Recordset("obs") <> Text5.Text Then
                        data_evalua.Recordset("obs") = Text5.Text
                        Xsigrabo = 8
                     End If
                  Else
                     data_evalua.Recordset("obs") = Text5.Text
                     Xsigrabo = 8
                  End If
               Else
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     data_evalua.Recordset("obs") = Null
                     Xsigrabo = 8
                  End If
               End If
            End If
            If x = 6 Then
               If Text6.Text <> "" Then
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     If data_evalua.Recordset("obs") <> Text6.Text Then
                        data_evalua.Recordset("obs") = Text6.Text
                        Xsigrabo = 8
                     End If
                  Else
                     data_evalua.Recordset("obs") = Text6.Text
                     Xsigrabo = 8
                  End If
               Else
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     data_evalua.Recordset("obs") = Null
                     Xsigrabo = 8
                  End If
               End If
            End If
            If x = 7 Then
               If Text7.Text <> "" Then
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     If data_evalua.Recordset("obs") <> Text7.Text Then
                        data_evalua.Recordset("obs") = Text7.Text
                        Xsigrabo = 8
                     End If
                  Else
                     data_evalua.Recordset("obs") = Text7.Text
                     Xsigrabo = 8
                  End If
               Else
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     data_evalua.Recordset("obs") = Null
                     Xsigrabo = 8
                  End If
               End If
            End If
            If x = 8 Then
               If Text8.Text <> "" Then
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     If data_evalua.Recordset("obs") <> Text8.Text Then
                        data_evalua.Recordset("obs") = Text8.Text
                        Xsigrabo = 8
                     End If
                  Else
                     data_evalua.Recordset("obs") = Text8.Text
                     Xsigrabo = 8
                  End If
               Else
                  If IsNull(data_evalua.Recordset("obs")) = False Then
                     data_evalua.Recordset("obs") = Null
                     Xsigrabo = 8
                  End If
               End If
            End If
            If Xsigrabo = 8 Then
               data_evalua.Recordset.Update
               Xsigrabo = 0
            Else
               data_evalua.Recordset.CancelUpdate
            End If
            data_evalua.Recordset.MoveNext
        Next
        cbop1.ListIndex = -1
        cbop2.ListIndex = -1
        cbop3.ListIndex = -1
        cbop4.ListIndex = -1
        cbop5.ListIndex = -1
        cbop6.ListIndex = -1
        cbop7.ListIndex = -1
        cbop8.ListIndex = -1
        cbotitulos.ListIndex = -1
        b_graba.Enabled = False
        b_cance.Enabled = False
        b_edita.Enabled = True
        Frame2.Enabled = False
        b_imp.Enabled = True
        b_firma.Enabled = True
        
'        MsgBox "La evaluación ya está firmada, no es posible modificar", vbInformation
     End If
   Else
      For x = 1 To 8
          data_evalua.Recordset.AddNew
          data_evalua.Recordset("id") = Data1.Recordset("nro_eval") + 1
          Data1.Recordset.Edit
          Data1.Recordset("nro_eval") = Data1.Recordset("nro_eval") + 1
          Data1.Recordset.Update
          data_evalua.Recordset("fecha") = Format(labfec.Caption, "dd-mm-yyyy")
          data_evalua.Recordset("fechamod") = Format(Date, "dd-mm-yyyy")
          data_evalua.Recordset("idempl") = Wxelnrocedev
          data_evalua.Recordset("idjefe") = Wxeljefeid
          data_evalua.Recordset("idtitulo") = Wxtitu
          data_evalua.Recordset("titulo") = cbotitulos.Text
          data_evalua.Recordset("id2") = Wxelnroid2
          If x = 1 Then
             data_evalua.Recordset("idpregun") = t_1.Text
             If Text1.Text <> "" Then
                data_evalua.Recordset("obs") = Text1.Text
             End If
          End If
          If x = 2 Then
             data_evalua.Recordset("idpregun") = t_2.Text
             If Text2.Text <> "" Then
                data_evalua.Recordset("obs") = Text2.Text
             End If
          End If
          If x = 3 Then
             data_evalua.Recordset("idpregun") = t_3.Text
             If Text3.Text <> "" Then
                data_evalua.Recordset("obs") = Text3.Text
             End If
          End If
          If x = 4 Then
             data_evalua.Recordset("idpregun") = t_4.Text
             If Text4.Text <> "" Then
                data_evalua.Recordset("obs") = Text4.Text
             End If
          End If
          If x = 5 Then
             data_evalua.Recordset("idpregun") = t_5.Text
             If Text5.Text <> "" Then
                data_evalua.Recordset("obs") = Text5.Text
             End If
          End If
          If x = 6 Then
             data_evalua.Recordset("idpregun") = t_6.Text
             If Text6.Text <> "" Then
                data_evalua.Recordset("obs") = Text6.Text
             End If
          End If
          If x = 7 Then
             data_evalua.Recordset("idpregun") = t_7.Text
             If Text7.Text <> "" Then
                data_evalua.Recordset("obs") = Text7.Text
             End If
          End If
          If x = 8 Then
             data_evalua.Recordset("idpregun") = t_8.Text
             If Text8.Text <> "" Then
                data_evalua.Recordset("obs") = Text8.Text
             End If
          End If
          data_evalua.Recordset("periodo") = cboperio.Text
          If x = 1 Then
             data_evalua.Recordset("puntos") = Val(cbop1.Text)
          End If
          If x = 2 Then
             data_evalua.Recordset("puntos") = Val(cbop2.Text)
          End If
          If x = 3 Then
             data_evalua.Recordset("puntos") = Val(cbop3.Text)
          End If
          If x = 4 Then
             data_evalua.Recordset("puntos") = Val(cbop4.Text)
          End If
          If x = 5 Then
             data_evalua.Recordset("puntos") = Val(cbop5.Text)
          End If
          If x = 6 Then
             data_evalua.Recordset("puntos") = Val(cbop6.Text)
          End If
          If x = 7 Then
             data_evalua.Recordset("puntos") = Val(cbop7.Text)
          End If
          If x = 8 Then
             data_evalua.Recordset("puntos") = Val(cbop8.Text)
          End If
          data_evalua.Recordset.Update
      Next
      cbop1.ListIndex = -1
      cbop2.ListIndex = -1
      cbop3.ListIndex = -1
      cbop4.ListIndex = -1
      cbop5.ListIndex = -1
      cbop6.ListIndex = -1
      cbop7.ListIndex = -1
      cbop8.ListIndex = -1
      cbotitulos.ListIndex = -1
      b_graba.Enabled = False
      b_cance.Enabled = False
      b_edita.Enabled = True
      b_imp.Enabled = True
      Frame2.Enabled = False
      b_firma.Enabled = True
   
   End If
Else
   MsgBox "Hay preguntas sin seleccionar puntaje, verifique!!", vbInformation
   
End If
b_firma.Enabled = True
   
b_graba.Enabled = True

Exit Sub

Errevalgrab2:
             If Err.Number = 3155 Then
                MsgBox "Error al grabar " & Err.Number & " " & Err.Description
               b_graba.Enabled = True
                Unload Me
             Else
                MsgBox "Error al grabar, verifique puntajes " & Err.Number & " " & Err.Description
                b_graba.Enabled = True
                Unload Me
             End If
             

End Sub

Private Sub b_imp_Click()
If WElusuario = "BRUNO" Or WElusuario = "GFERNANDEZ" Or WElusuario = "JFERNAN" Or WElusuario = "MCOSTA" Or _
   WElusuario = "SDOMINGUEZ" Or WElusuario = "SPEREZ" Or WElusuario = "DARIOH" Or WElusuario = "MARCELOM" Or XWeltipoU = "USUARIOS ADM" Or _
   WElusuario = "ENRIQUE" Or WElusuario = "BDD" Or WElusuario = "AGUILLEN" Or XWeltipoU = "ADMINISTRADOR" Or XWeltipoU = "USUARIOS DESP" Then
   frm_infeval.Show vbModal
Else

End If

End Sub

Private Sub b_select_Click()

b_select.Enabled = False
cbop1.ListIndex = -1
cbop2.ListIndex = -1
cbop3.ListIndex = -1
cbop4.ListIndex = -1
cbop5.ListIndex = -1
cbop6.ListIndex = -1
cbop7.ListIndex = -1
cbop8.ListIndex = -1
Dim Xpun As Integer
Xpun = 0
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""

data_titulos.RecordSource = "Select * from titulos where descrip ='" & cbotitulos.Text & "'"
data_titulos.Refresh
If data_titulos.Recordset.RecordCount > 0 Then
   Wxtitu = data_titulos.Recordset("id")
   data_texto.RecordSource = "Select * from textos where codtitulo =" & data_titulos.Recordset("id") & " and codcargo =" & Wxquepreg & " order by pregunta"
   data_texto.Refresh
   If data_texto.Recordset.RecordCount > 0 Then
      data_texto.Recordset.MoveFirst
      Do While Not data_texto.Recordset.EOF
         If data_texto.Recordset("pregunta") = 1 Then
            lab1.Caption = data_texto.Recordset("descrip")
            t_1.Text = 1
         End If
         If data_texto.Recordset("pregunta") = 2 Then
            lab2.Caption = data_texto.Recordset("descrip")
            t_2.Text = 2
         End If
         If data_texto.Recordset("pregunta") = 3 Then
            lab3.Caption = data_texto.Recordset("descrip")
            t_3.Text = 3
         End If
         If data_texto.Recordset("pregunta") = 4 Then
            lab4.Caption = data_texto.Recordset("descrip")
            t_4.Text = 4
         End If
         If data_texto.Recordset("pregunta") = 5 Then
            lab5.Caption = data_texto.Recordset("descrip")
            t_5.Text = 5
         End If
         If data_texto.Recordset("pregunta") = 6 Then
            lab6.Caption = data_texto.Recordset("descrip")
            t_6.Text = 6
         End If
         If data_texto.Recordset("pregunta") = 7 Then
            lab7.Caption = data_texto.Recordset("descrip")
            t_7.Text = 7
         End If
         If data_texto.Recordset("pregunta") = 8 Then
            lab8.Caption = data_texto.Recordset("descrip")
            t_8.Text = 8
         End If
       
         If data_texto.Recordset("pregunta") = 9 Then
            lab1.Caption = data_texto.Recordset("descrip")
            t_1.Text = 9
         End If
         If data_texto.Recordset("pregunta") = 10 Then
            lab2.Caption = data_texto.Recordset("descrip")
            t_2.Text = 10
         End If
         If data_texto.Recordset("pregunta") = 11 Then
            lab3.Caption = data_texto.Recordset("descrip")
            t_3.Text = 11
         End If
         If data_texto.Recordset("pregunta") = 12 Then
            lab4.Caption = data_texto.Recordset("descrip")
            t_4.Text = 12
         End If
         If data_texto.Recordset("pregunta") = 13 Then
            lab5.Caption = data_texto.Recordset("descrip")
            t_5.Text = 13
         End If
         If data_texto.Recordset("pregunta") = 14 Then
            lab6.Caption = data_texto.Recordset("descrip")
            t_6.Text = 14
         End If
         If data_texto.Recordset("pregunta") = 15 Then
            lab7.Caption = data_texto.Recordset("descrip")
            t_7.Text = 15
         End If
         If data_texto.Recordset("pregunta") = 16 Then
            lab8.Caption = data_texto.Recordset("descrip")
            t_8.Text = 16
         End If
       
         If data_texto.Recordset("pregunta") = 17 Then
            lab1.Caption = data_texto.Recordset("descrip")
            t_1.Text = 17
         End If
         If data_texto.Recordset("pregunta") = 18 Then
            lab2.Caption = data_texto.Recordset("descrip")
            t_2.Text = 18
         End If
         If data_texto.Recordset("pregunta") = 19 Then
            lab3.Caption = data_texto.Recordset("descrip")
            t_3.Text = 19
         End If
         If data_texto.Recordset("pregunta") = 20 Then
            lab4.Caption = data_texto.Recordset("descrip")
            t_4.Text = 20
         End If
         If data_texto.Recordset("pregunta") = 21 Then
            lab5.Caption = data_texto.Recordset("descrip")
            t_5.Text = 21
         End If
         If data_texto.Recordset("pregunta") = 22 Then
            lab6.Caption = data_texto.Recordset("descrip")
            t_6.Text = 22
         End If
         If data_texto.Recordset("pregunta") = 23 Then
            lab7.Caption = data_texto.Recordset("descrip")
            t_7.Text = 23
         End If
         If data_texto.Recordset("pregunta") = 24 Then
            lab8.Caption = data_texto.Recordset("descrip")
            t_8.Text = 24
         End If
       
         If data_texto.Recordset("pregunta") = 25 Then
            lab1.Caption = data_texto.Recordset("descrip")
            t_1.Text = 25
         End If
         If data_texto.Recordset("pregunta") = 26 Then
            lab2.Caption = data_texto.Recordset("descrip")
            t_2.Text = 26
         End If
         If data_texto.Recordset("pregunta") = 27 Then
            lab3.Caption = data_texto.Recordset("descrip")
            t_3.Text = 27
         End If
         If data_texto.Recordset("pregunta") = 28 Then
            lab4.Caption = data_texto.Recordset("descrip")
            t_4.Text = 28
         End If
         If data_texto.Recordset("pregunta") = 29 Then
            lab5.Caption = data_texto.Recordset("descrip")
            t_5.Text = 29
         End If
         If data_texto.Recordset("pregunta") = 30 Then
            lab6.Caption = data_texto.Recordset("descrip")
            t_6.Text = 30
         End If
         If data_texto.Recordset("pregunta") = 31 Then
            lab7.Caption = data_texto.Recordset("descrip")
            t_7.Text = 31
         End If
         If data_texto.Recordset("pregunta") = 32 Then
            lab8.Caption = data_texto.Recordset("descrip")
            t_8.Text = 32
         End If
       'aca ini
         If data_texto.Recordset("pregunta") = 33 Then
            lab1.Caption = data_texto.Recordset("descrip")
            t_1.Text = 33
         End If
         If data_texto.Recordset("pregunta") = 34 Then
            lab2.Caption = data_texto.Recordset("descrip")
            t_2.Text = 34
         End If
         If data_texto.Recordset("pregunta") = 35 Then
            lab3.Caption = data_texto.Recordset("descrip")
            t_3.Text = 35
         End If
         If data_texto.Recordset("pregunta") = 36 Then
            lab4.Caption = data_texto.Recordset("descrip")
            t_4.Text = 36
         End If
         If data_texto.Recordset("pregunta") = 37 Then
            lab5.Caption = data_texto.Recordset("descrip")
            t_5.Text = 37
         End If
         If data_texto.Recordset("pregunta") = 38 Then
            lab6.Caption = data_texto.Recordset("descrip")
            t_6.Text = 38
         End If
         If data_texto.Recordset("pregunta") = 39 Then
            lab7.Caption = data_texto.Recordset("descrip")
            t_7.Text = 39
         End If
         If data_texto.Recordset("pregunta") = 40 Then
            lab8.Caption = data_texto.Recordset("descrip")
            t_8.Text = 40
         End If
''aca fin
         data_texto.Recordset.MoveNext
      Loop
      If frm_abmper.Combo2.Text = 2016 Then
         data_evalua.RecordSource = "Select * from evaluas where idempl =" & Wxelnrocedev & " and periodo ='" & cboperio.Text & "' and idtitulo =" & Wxtitu & " and idjefe =" & Wxeljefeid & " and id2 =" & Wxelnroid2 & " order by idpregun"
         data_evalua.Refresh
      Else
         data_evalua.RecordSource = "Select * from evaluas where idempl =" & Wxelnrocedev & " and periodo ='" & cboperio.Text & "' and idtitulo =" & Wxtitu & " and idjefe =" & Wxeljefeid & " and id2 =" & Wxelnroid2 & " order by idpregun"
         data_evalua.Refresh
         If data_evalua.Recordset.RecordCount > 0 Then
            If IsNull(data_evalua.Recordset("cierre")) = False Then
               If data_evalua.Recordset("cierre") = "SI" Then
                  MsgBox "Evaluación CERRADA. No se puede modificar.", vbExclamation
                  b_graba.Enabled = False
               End If
            Else
               b_graba.Enabled = True
            End If
         Else
            b_graba.Enabled = True
         End If
      End If
      If data_evalua.Recordset.RecordCount > 0 Then
         Xpun = 1
         Do While Not data_evalua.Recordset.EOF
            If Xpun = 1 Then
               If data_evalua.Recordset("puntos") = 1 Then
                  cbop1.ListIndex = 0
               End If
               If data_evalua.Recordset("puntos") = 2 Then
                  cbop1.ListIndex = 1
               End If
               If data_evalua.Recordset("puntos") = 3 Then
                  cbop1.ListIndex = 2
               End If
               If data_evalua.Recordset("puntos") = 4 Then
                  cbop1.ListIndex = 3
               End If
               If IsNull(data_evalua.Recordset("obs")) = False Then
                  Text1.Text = data_evalua.Recordset("obs")
               Else
                  Text1.Text = ""
               End If
            End If
            If Xpun = 2 Then
               If data_evalua.Recordset("puntos") = 1 Then
                  cbop2.ListIndex = 0
               End If
               If data_evalua.Recordset("puntos") = 2 Then
                  cbop2.ListIndex = 1
               End If
               If data_evalua.Recordset("puntos") = 3 Then
                  cbop2.ListIndex = 2
               End If
               If data_evalua.Recordset("puntos") = 4 Then
                  cbop2.ListIndex = 3
               End If
               If IsNull(data_evalua.Recordset("obs")) = False Then
                  Text2.Text = data_evalua.Recordset("obs")
               Else
                  Text2.Text = ""
               End If
            End If
            If Xpun = 3 Then
               If data_evalua.Recordset("puntos") = 1 Then
                  cbop3.ListIndex = 0
               End If
               If data_evalua.Recordset("puntos") = 2 Then
                  cbop3.ListIndex = 1
               End If
               If data_evalua.Recordset("puntos") = 3 Then
                  cbop3.ListIndex = 2
               End If
               If data_evalua.Recordset("puntos") = 4 Then
                  cbop3.ListIndex = 3
               End If
               If IsNull(data_evalua.Recordset("obs")) = False Then
                  Text3.Text = data_evalua.Recordset("obs")
               Else
                  Text3.Text = ""
               End If
            End If
            If Xpun = 4 Then
               If data_evalua.Recordset("puntos") = 1 Then
                  cbop4.ListIndex = 0
               End If
               If data_evalua.Recordset("puntos") = 2 Then
                  cbop4.ListIndex = 1
               End If
               If data_evalua.Recordset("puntos") = 3 Then
                  cbop4.ListIndex = 2
               End If
               If data_evalua.Recordset("puntos") = 4 Then
                  cbop4.ListIndex = 3
               End If
               If IsNull(data_evalua.Recordset("obs")) = False Then
                  Text4.Text = data_evalua.Recordset("obs")
               Else
                  Text4.Text = ""
               End If
            End If
            If Xpun = 5 Then
               If data_evalua.Recordset("puntos") = 1 Then
                  cbop5.ListIndex = 0
               End If
               If data_evalua.Recordset("puntos") = 2 Then
                  cbop5.ListIndex = 1
               End If
               If data_evalua.Recordset("puntos") = 3 Then
                  cbop5.ListIndex = 2
               End If
               If data_evalua.Recordset("puntos") = 4 Then
                  cbop5.ListIndex = 3
               End If
               If IsNull(data_evalua.Recordset("obs")) = False Then
                  Text5.Text = data_evalua.Recordset("obs")
               Else
                  Text5.Text = ""
               End If
            End If
            If Xpun = 6 Then
               If data_evalua.Recordset("puntos") = 1 Then
                  cbop6.ListIndex = 0
               End If
               If data_evalua.Recordset("puntos") = 2 Then
                  cbop6.ListIndex = 1
               End If
               If data_evalua.Recordset("puntos") = 3 Then
                  cbop6.ListIndex = 2
               End If
               If data_evalua.Recordset("puntos") = 4 Then
                  cbop6.ListIndex = 3
               End If
               If IsNull(data_evalua.Recordset("obs")) = False Then
                  Text6.Text = data_evalua.Recordset("obs")
               Else
                  Text6.Text = ""
               End If
            End If
            If Xpun = 7 Then
               If data_evalua.Recordset("puntos") = 1 Then
                  cbop7.ListIndex = 0
               End If
               If data_evalua.Recordset("puntos") = 2 Then
                  cbop7.ListIndex = 1
               End If
               If data_evalua.Recordset("puntos") = 3 Then
                  cbop7.ListIndex = 2
               End If
               If data_evalua.Recordset("puntos") = 4 Then
                  cbop7.ListIndex = 3
               End If
               If IsNull(data_evalua.Recordset("obs")) = False Then
                  Text7.Text = data_evalua.Recordset("obs")
               Else
                  Text7.Text = ""
               End If
            End If
            If Xpun = 8 Then
               If data_evalua.Recordset("puntos") = 1 Then
                  cbop8.ListIndex = 0
               End If
               If data_evalua.Recordset("puntos") = 2 Then
                  cbop8.ListIndex = 1
               End If
               If data_evalua.Recordset("puntos") = 3 Then
                  cbop8.ListIndex = 2
               End If
               If data_evalua.Recordset("puntos") = 4 Then
                  cbop8.ListIndex = 3
               End If
               If IsNull(data_evalua.Recordset("obs")) = False Then
                  Text8.Text = data_evalua.Recordset("obs")
               Else
                  Text8.Text = ""
               End If
            End If
            data_evalua.Recordset.MoveNext
            Xpun = Xpun + 1
         Loop
      End If
   Else
      MsgBox "No existen preguntas para éste título/cargo"
      
   End If

Else
   MsgBox "No se encuentra título"
End If
b_imp.Enabled = True
b_firma.Enabled = True
b_select.Enabled = True

End Sub

Private Sub cbop1_Click()
If cbop1.ListIndex = 0 Then
   labnota.Caption = "1-->Bajo lo esperado. --- Desempeño no cumple con las expectativas en uno o más objetivos de su responsabilidad -uno o más de los objetivos no fueron cumplidos-. Se debe trazar un plan de desarrollo para mejoras con PLAZOS y MONITOREOS"
Else
   If cbop1.ListIndex = 1 Then
      labnota.Caption = "2--> Cumple lo esperado. --- Desempeño cumple con las expectativas de su descripción de cargo y responsabilidades asignadas. El trabajador necesita constante apoyo para la mejora de su rendimiento."
   Else
      If cbop1.ListIndex = 2 Then
         labnota.Caption = "3--> Logra lo esperado. --- Desempeño cumple con las expectativas de su descripción de cargo y responsabilidades asignadas. Es un aporte constante para su jefatura y pares."
      Else
         If cbop1.ListIndex = 3 Then
            labnota.Caption = "4--> Sobre lo esperado. --- Desempeño excelente en cada uno de sus objetivos y tareas a ejecutar , es un funcionario ejemplo y destacado para sus jefaturas y pares."
         Else
            labnota.Caption = ""
         End If
      End If
   End If
End If

   
End Sub

Private Sub cbop2_Click()
If cbop2.ListIndex = 0 Then
   labnota.Caption = "1-->Bajo lo esperado. --- Desempeño no cumple con las expectativas en uno o más objetivos de su responsabilidad -uno o más de los objetivos no fueron cumplidos-. Se debe trazar un plan de desarrollo para mejoras con PLAZOS y MONITOREOS"
Else
   If cbop2.ListIndex = 1 Then
      labnota.Caption = "2--> Cumple lo esperado. --- Desempeño cumple con las expectativas de su descripción de cargo y responsabilidades asignadas. El trabajador necesita constante apoyo para la mejora de su rendimiento."
   Else
      If cbop2.ListIndex = 2 Then
         labnota.Caption = "3--> Logra lo esperado. --- Desempeño cumple con las expectativas de su descripción de cargo y responsabilidades asignadas. Es un aporte constante para su jefatura y pares."
      Else
         If cbop2.ListIndex = 3 Then
            labnota.Caption = "4--> Sobre lo esperado. --- Desempeño excelente en cada uno de sus objetivos y tareas a ejecutar , es un funcionario ejemplo y destacado para sus jefaturas y pares."
         Else
            labnota.Caption = ""
         End If
      End If
   End If
End If

End Sub

Private Sub cbop3_Click()
If cbop3.ListIndex = 0 Then
   labnota.Caption = "1-->Bajo lo esperado. --- Desempeño no cumple con las expectativas en uno o más objetivos de su responsabilidad -uno o más de los objetivos no fueron cumplidos-. Se debe trazar un plan de desarrollo para mejoras con PLAZOS y MONITOREOS"
Else
   If cbop3.ListIndex = 1 Then
      labnota.Caption = "2--> Cumple lo esperado. --- Desempeño cumple con las expectativas de su descripción de cargo y responsabilidades asignadas. El trabajador necesita constante apoyo para la mejora de su rendimiento."
   Else
      If cbop3.ListIndex = 2 Then
         labnota.Caption = "3--> Logra lo esperado. --- Desempeño cumple con las expectativas de su descripción de cargo y responsabilidades asignadas. Es un aporte constante para su jefatura y pares."
      Else
         If cbop3.ListIndex = 3 Then
            labnota.Caption = "4--> Sobre lo esperado. --- Desempeño excelente en cada uno de sus objetivos y tareas a ejecutar , es un funcionario ejemplo y destacado para sus jefaturas y pares."
         Else
            labnota.Caption = ""
         End If
      End If
   End If
End If

End Sub

Private Sub cbop4_Click()
If cbop4.ListIndex = 0 Then
   labnota.Caption = "1-->Bajo lo esperado. --- Desempeño no cumple con las expectativas en uno o más objetivos de su responsabilidad -uno o más de los objetivos no fueron cumplidos-. Se debe trazar un plan de desarrollo para mejoras con PLAZOS y MONITOREOS"
Else
   If cbop4.ListIndex = 1 Then
      labnota.Caption = "2--> Cumple lo esperado. --- Desempeño cumple con las expectativas de su descripción de cargo y responsabilidades asignadas. El trabajador necesita constante apoyo para la mejora de su rendimiento."
   Else
      If cbop4.ListIndex = 2 Then
         labnota.Caption = "3--> Logra lo esperado. --- Desempeño cumple con las expectativas de su descripción de cargo y responsabilidades asignadas. Es un aporte constante para su jefatura y pares."
      Else
         If cbop4.ListIndex = 3 Then
            labnota.Caption = "4--> Sobre lo esperado. --- Desempeño excelente en cada uno de sus objetivos y tareas a ejecutar , es un funcionario ejemplo y destacado para sus jefaturas y pares."
         Else
            labnota.Caption = ""
         End If
      End If
   End If
End If

End Sub

Private Sub cbop5_Click()
If cbop5.ListIndex = 0 Then
   labnota.Caption = "1-->Bajo lo esperado. --- Desempeño no cumple con las expectativas en uno o más objetivos de su responsabilidad -uno o más de los objetivos no fueron cumplidos-. Se debe trazar un plan de desarrollo para mejoras con PLAZOS y MONITOREOS"
Else
   If cbop5.ListIndex = 1 Then
      labnota.Caption = "2--> Cumple lo esperado. --- Desempeño cumple con las expectativas de su descripción de cargo y responsabilidades asignadas. El trabajador necesita constante apoyo para la mejora de su rendimiento."
   Else
      If cbop5.ListIndex = 2 Then
         labnota.Caption = "3--> Logra lo esperado. --- Desempeño cumple con las expectativas de su descripción de cargo y responsabilidades asignadas. Es un aporte constante para su jefatura y pares."
      Else
         If cbop5.ListIndex = 3 Then
            labnota.Caption = "4--> Sobre lo esperado. --- Desempeño excelente en cada uno de sus objetivos y tareas a ejecutar , es un funcionario ejemplo y destacado para sus jefaturas y pares."
         Else
            labnota.Caption = ""
         End If
      End If
   End If
End If

End Sub

Private Sub cbop6_Click()
If cbop6.ListIndex = 0 Then
   labnota.Caption = "1-->Bajo lo esperado. --- Desempeño no cumple con las expectativas en uno o más objetivos de su responsabilidad -uno o más de los objetivos no fueron cumplidos-. Se debe trazar un plan de desarrollo para mejoras con PLAZOS y MONITOREOS"
Else
   If cbop6.ListIndex = 1 Then
      labnota.Caption = "2--> Cumple lo esperado. --- Desempeño cumple con las expectativas de su descripción de cargo y responsabilidades asignadas. El trabajador necesita constante apoyo para la mejora de su rendimiento."
   Else
      If cbop6.ListIndex = 2 Then
         labnota.Caption = "3--> Logra lo esperado. --- Desempeño cumple con las expectativas de su descripción de cargo y responsabilidades asignadas. Es un aporte constante para su jefatura y pares."
      Else
         If cbop6.ListIndex = 3 Then
            labnota.Caption = "4--> Sobre lo esperado. --- Desempeño excelente en cada uno de sus objetivos y tareas a ejecutar , es un funcionario ejemplo y destacado para sus jefaturas y pares."
         Else
            labnota.Caption = ""
         End If
      End If
   End If
End If

End Sub

Private Sub cbop7_Click()
If cbop7.ListIndex = 0 Then
   labnota.Caption = "1-->Bajo lo esperado. --- Desempeño no cumple con las expectativas en uno o más objetivos de su responsabilidad -uno o más de los objetivos no fueron cumplidos-. Se debe trazar un plan de desarrollo para mejoras con PLAZOS y MONITOREOS"
Else
   If cbop7.ListIndex = 1 Then
      labnota.Caption = "2--> Cumple lo esperado. --- Desempeño cumple con las expectativas de su descripción de cargo y responsabilidades asignadas. El trabajador necesita constante apoyo para la mejora de su rendimiento."
   Else
      If cbop7.ListIndex = 2 Then
         labnota.Caption = "3--> Logra lo esperado. --- Desempeño cumple con las expectativas de su descripción de cargo y responsabilidades asignadas. Es un aporte constante para su jefatura y pares."
      Else
         If cbop7.ListIndex = 3 Then
            labnota.Caption = "4--> Sobre lo esperado. --- Desempeño excelente en cada uno de sus objetivos y tareas a ejecutar , es un funcionario ejemplo y destacado para sus jefaturas y pares."
         Else
            labnota.Caption = ""
         End If
      End If
   End If
End If

End Sub

Private Sub cbop8_Click()
If cbop8.ListIndex = 0 Then
   labnota.Caption = "1-->Bajo lo esperado. --- Desempeño no cumple con las expectativas en uno o más objetivos de su responsabilidad -uno o más de los objetivos no fueron cumplidos-. Se debe trazar un plan de desarrollo para mejoras con PLAZOS y MONITOREOS"
Else
   If cbop8.ListIndex = 1 Then
      labnota.Caption = "2--> Cumple lo esperado. --- Desempeño cumple con las expectativas de su descripción de cargo y responsabilidades asignadas. El trabajador necesita constante apoyo para la mejora de su rendimiento."
   Else
      If cbop8.ListIndex = 2 Then
         labnota.Caption = "3--> Logra lo esperado. --- Desempeño cumple con las expectativas de su descripción de cargo y responsabilidades asignadas. Es un aporte constante para su jefatura y pares."
      Else
         If cbop8.ListIndex = 3 Then
            labnota.Caption = "4--> Sobre lo esperado. --- Desempeño excelente en cada uno de sus objetivos y tareas a ejecutar , es un funcionario ejemplo y destacado para sus jefaturas y pares."
         Else
            labnota.Caption = ""
         End If
      End If
   End If
End If

End Sub

Private Sub Command3_Click()
Xquecol = 3

End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.path & "\parevalu.mdb"
Data1.RecordSource = "parevalu"
Data1.Refresh

cbop1.AddItem "1"
cbop1.AddItem "2"
cbop1.AddItem "3"
cbop1.AddItem "4"

cbop2.AddItem "1"
cbop2.AddItem "2"
cbop2.AddItem "3"
cbop2.AddItem "4"

cbop3.AddItem "1"
cbop3.AddItem "2"
cbop3.AddItem "3"
cbop3.AddItem "4"

cbop4.AddItem "1"
cbop4.AddItem "2"
cbop4.AddItem "3"
cbop4.AddItem "4"

cbop5.AddItem "1"
cbop5.AddItem "2"
cbop5.AddItem "3"
cbop5.AddItem "4"

cbop6.AddItem "1"
cbop6.AddItem "2"
cbop6.AddItem "3"
cbop6.AddItem "4"

cbop7.AddItem "1"
cbop7.AddItem "2"
cbop7.AddItem "3"
cbop7.AddItem "4"

cbop8.AddItem "1"
cbop8.AddItem "2"
cbop8.AddItem "3"
cbop8.AddItem "4"

If frm_abmper.Combo2.Text = "2016" Then
   data_titulos.Connect = "ODBC;DSN=eval2015;"
   data_evalua.Connect = "ODBC;DSN=eval2015;"
        
   data_texto.Connect = "ODBC;DSN=eval2015;"
   data_periodo.Connect = "ODBC;DSN=eval2015;"
   data_periodo.RecordSource = "periodo"
   data_periodo.Refresh
Else
   data_titulos.Connect = "ODBC;DSN=sappper;"
   data_evalua.Connect = "ODBC;DSN=sappper;"
   data_texto.Connect = "ODBC;DSN=sappper;"
   data_periodo.Connect = "ODBC;DSN=sappper;"
   data_periodo.RecordSource = "periodo"
   data_periodo.Refresh
End If

If data_periodo.Recordset.RecordCount > 0 Then
   data_periodo.Recordset.MoveFirst
'   Label8.Caption = data_periodo.Recordset("descrip")
   Do While Not data_periodo.Recordset.EOF
      cboperio.AddItem data_periodo.Recordset("descrip")
      data_periodo.Recordset.MoveNext
   Loop
   cboperio.ListIndex = 0
Else
'   Label8.Caption = ""
End If

data_titulos.RecordSource = "Select * from titulos order by id"
data_titulos.Refresh
If data_titulos.Recordset.RecordCount > 0 Then
   data_titulos.Recordset.MoveFirst
   Do While Not data_titulos.Recordset.EOF
      cbotitulos.AddItem data_titulos.Recordset("descrip")
      data_titulos.Recordset.MoveNext
   Loop
End If
labnome.Caption = frm_abmper.data_buscap.Recordset("nom1") & " " & frm_abmper.data_buscap.Recordset("ape1")


If XWquecargo = 1 Then
   labnomj.Caption = frm_abmper.t_nom1.Text & " " & frm_abmper.t_apel1.Text
   b_1.Visible = False
   b_2.Visible = False
   b_3.Visible = False
   b_4.Visible = False
   b_5.Visible = False
   b_6.Visible = False
   b_7.Visible = False
   b_8.Visible = False
Else
   data_per.RecordSource = "Select * from personas where id =" & Wxeljefeid
   data_per.Refresh
   b_1.Visible = True
   b_2.Visible = True
   b_3.Visible = True
   b_4.Visible = True
   b_5.Visible = True
   b_6.Visible = True
   b_7.Visible = True
   b_8.Visible = True
   If data_per.Recordset.RecordCount > 0 Then
      labnomj.Caption = data_per.Recordset("nom1") & " " & data_per.Recordset("ape1")
   Else
      MsgBox "No existe funcionario, verifique con el administrador"
      Unload Me
      
   End If

End If

labfec.Caption = Date
If labnome.Caption = labnomj.Caption Then
   If WElusuario = "JFERNAN" Then
      b_cierra.Enabled = True
   Else
      b_cierra.Enabled = False
   End If
Else
   b_cierra.Enabled = True
End If


End Sub

Private Sub Form_Resize()
With Image1
    .Left = 0
    .Top = 0
    .Height = Me.Height
    .Width = Me.Width
End With

End Sub
