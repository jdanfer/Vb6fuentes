VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frm_consultos 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "frm_consultos"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4665
   Icon            =   "frm_consultos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   4665
   StartUpPosition =   1  'CenterOwner
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frm_consultos.frx":058A
      Height          =   3135
      Left            =   240
      OleObjectBlob   =   "frm_consultos.frx":059E
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "odbc;dsn=sappespecial;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "cons"
      Top             =   3240
      Visible         =   0   'False
      Width           =   1815
   End
End
Attribute VB_Name = "frm_consultos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
