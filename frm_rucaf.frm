VERSION 5.00
Begin VB.Form frm_rucaf 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Información a RUCAF"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6750
   Icon            =   "frm_rucaf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   6750
   StartUpPosition =   1  'CenterOwner
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   3480
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   1320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4320
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox t_anio 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3840
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox t_mes 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Procesar"
      Height          =   615
      Left            =   360
      Picture         =   "frm_rucaf.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   1815
      Left            =   2760
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   495
      Left            =   5520
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   495
      Left            =   3720
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   2040
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Aguarde, procesando....10 minutos aprox."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00800000&
      Caption         =   "Mes y Año a enviar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   360
      TabIndex        =   7
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "frm_rucaf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim objDOM As MSXML2.DOMDocument30
Dim objModule As MSXML2.DOMDocument30
Set objDOM = New MSXML2.DOMDocument30
Dim objRootElem As MSXML2.IXMLDOMElement
Set objModule = New MSXML2.DOMDocument30
'Set XMLInstruccion = objDOM.createProcessingInstruction("xml", "version='1.0' encoding='ISO-8859-1'")
'objDOM.appendChild XMLInstruccion
Dim Xnomemi, Xarch As String
Dim Xlargo, Xlargo2, Xlargo3, XcantReg As Integer
Xnomemi = "emi"
Xarch = Trim(t_anio.Text)

If t_mes.Text <> "" Then
   If t_mes.Text < 10 Then
      Xnomemi = Xnomemi & "0" & Trim(Str(t_mes.Text))
      Xarch = Xarch & "0" & Trim(Str(t_mes.Text))
   Else
      Xnomemi = Xnomemi & Trim(Str(t_mes.Text))
      Xarch = Xarch & Trim(Str(t_mes.Text))
   End If
End If
If t_anio.Text <> "" Then
   Xnomemi = Xnomemi & Mid(Trim(Str(t_anio.Text)), 3, 2)
End If
Xarch = Xarch & "_303"
Data1.RecordSource = "Select * from " & Xnomemi & " where nro_cobr not in (5,11)"
Data1.Refresh
If Data1.Recordset.RecordCount > 0 Then
   Label2.Visible = True
   Data1.Recordset.MoveFirst
   DoEvents
   Text1.Text = ""
   Text1.Text = "<?xml version=""1.0"" encoding=""ISO-8859-1""?>"
   Text1.Text = Text1.Text & vbCrLf & "<Afiliados xmlns=""RUCAF"">"
   Open App.Path & "\" & Xarch & ".xml" For Output As #1
   XcantReg = 0
   Do While Not Data1.Recordset.EOF
      Data3.RecordSource = "Select * from convenio where cnv_codigo ='" & Data1.Recordset("cod_cnv") & "'"
      Data3.Refresh
      If Data3.Recordset.RecordCount > 0 Then
         If Data3.Recordset("cnv_cant_r") = 2 Then
            XcantReg = XcantReg + 1
            If IsNull(Data1.Recordset("cedula")) = False Then
               If Data1.Recordset("cedula") > 99000 Then
                  Text1.Text = Text1.Text & "<Afiliados.Afiliado>"
                  Text1.Text = Text1.Text & vbCrLf & "<TipoDocumento>1</TipoDocumento>"
                  If IsNull(Data1.Recordset("cedula")) = False Then
                     Text1.Text = Text1.Text & vbCrLf & "<NroDocumento>" & Data1.Recordset("cedula") & "</NroDocumento>"
                     Text1.Text = Text1.Text & vbCrLf & "<CI>" & Data1.Recordset("cedula") & "</CI>"
                  Else
                     Text1.Text = Text1.Text & vbCrLf & "<NroDocumento>0</NroDocumento>"
                     Text1.Text = Text1.Text & vbCrLf & "<CI>0</CI>"
                  End If
                  Data2.RecordSource = "Select * from clientes where cl_codigo =" & Data1.Recordset("cliente")
                  Data2.Refresh
                  If Data2.Recordset.RecordCount > 0 Then
                     If IsNull(Data2.Recordset("cl_sexo")) = False Then
                        If Data2.Recordset("cl_sexo") = 1 Then
                           Text1.Text = Text1.Text & vbCrLf & "<Sexo>M</Sexo>"
                        Else
                           Text1.Text = Text1.Text & vbCrLf & "<Sexo>F</Sexo>"
                        End If
                     Else
                        Text1.Text = Text1.Text & vbCrLf & "<Sexo>F</Sexo>"
                     End If
                  Else
                     Text1.Text = Text1.Text & vbCrLf & "<Sexo>F</Sexo>"
                  End If
                  If IsNull(Data1.Recordset("fecha_nac")) = False Then
                     Text1.Text = Text1.Text & vbCrLf & "<FchNac>" & Format(Data1.Recordset("fecha_nac"), "yyyy-mm-dd") & "</FchNac>"
                  Else
                     Text1.Text = Text1.Text & vbCrLf & "<FchNac>1900-01-01</FchNac>"
                  End If
                  If IsNull(Data1.Recordset("apellidos")) = False Then
                     Xlargo = Len(Trim(Data1.Recordset("apellidos")))
                     Xlargo2 = Xlargo / 2
                     Xlargo3 = Xlargo2 + 1
                     Text1.Text = Text1.Text & vbCrLf & "<NomPri>" & Mid(Data1.Recordset("apellidos"), Xlargo3, Xlargo2) & "</NomPri>"
                     Text1.Text = Text1.Text & vbCrLf & "<ApePri>" & Mid(Data1.Recordset("apellidos"), 1, Xlargo2) & "</ApePri>"
                  Else
                     Text1.Text = Text1.Text & vbCrLf & "<NomPri>NN</NomPri>"
                     Text1.Text = Text1.Text & vbCrLf & "<ApePri>NN</ApePri>"
                  End If
                  Text1.Text = Text1.Text & vbCrLf & "<PaisCodDir>UY</PaisCodDir>"
                  Text1.Text = Text1.Text & vbCrLf & "<DptoCodDir>3</DptoCodDir>"
                  If Data2.Recordset.RecordCount > 0 Then
                    If IsNull(Data2.Recordset("cl_grupo")) = False Then
                       If Data2.Recordset("cl_grupo") >= 101 And Data2.Recordset("cl_grupo") <= 104 Then
                          Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>627</LocCodDir>"
                       Else
                          If Data2.Recordset("cl_grupo") >= 201 And Data2.Recordset("cl_grupo") <= 209 Then
                             Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>724</LocCodDir>"
                          Else
                             If Data2.Recordset("cl_grupo") >= 300 And Data2.Recordset("cl_grupo") <= 322 Then
                                Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>729</LocCodDir>"
                             Else
                                If Data2.Recordset("cl_grupo") >= 400 And Data2.Recordset("cl_grupo") <= 419 Then
                                   Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>621</LocCodDir>"
                                Else
                                   If Data2.Recordset("cl_grupo") >= 500 And Data2.Recordset("cl_grupo") <= 501 Then
                                      Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>630</LocCodDir>"
                                   Else
                                      If Data2.Recordset("cl_grupo") = 630 Then
                                         Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>521</LocCodDir>"
                                      Else
                                         If Data2.Recordset("cl_grupo") = 650 Or Data2.Recordset("cl_grupo") = 800 Then
                                            Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>524</LocCodDir>"
                                         Else
                                            If Data2.Recordset("cl_grupo") >= 670 And Data2.Recordset("cl_grupo") <= 679 Then
                                               Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>528</LocCodDir>"
                                            Else
                                               If Data2.Recordset("cl_grupo") >= 600 And Data2.Recordset("cl_grupo") <= 640 Then
                                                  Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>322</LocCodDir>"
                                               Else
                                                  If Data2.Recordset("cl_grupo") >= 700 And Data2.Recordset("cl_grupo") <= 722 Then
                                                     Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>834</LocCodDir>"
                                                  Else
                                                      If Data2.Recordset("cl_grupo") >= 801 And Data2.Recordset("cl_grupo") <= 815 Then
                                                         Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>631</LocCodDir>"
                                                      Else
                                                         Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>627</LocCodDir>"
                                                      End If
                                                  End If
                                               End If
                                            End If
                                         End If
                                      End If
                                   End If
                                End If
                             End If
                          End If
                       End If
                    End If
                  Else
                    Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>627</LocCodDir>"
                  End If
                  Text1.Text = Text1.Text & vbCrLf & "<Instituciones>"
                  Text1.Text = Text1.Text & vbCrLf & "<Afiliados.Afiliado.InstitucionItem>"
                  If IsNull(Data1.Recordset("fecha_ing")) = False Then
                     Text1.Text = Text1.Text & vbCrLf & "<InstFchReg>" & Format(Data1.Recordset("fecha_ing"), "yyyy-mm-dd") & "</InstFchReg>"
                  Else
                     Text1.Text = Text1.Text & vbCrLf & "<InstFchReg>" & Format(Date, "yyyy-mm-dd") & "</InstFchReg>"
                  End If
                  Text1.Text = Text1.Text & vbCrLf & "<InstCod>303</InstCod>"
                  Text1.Text = Text1.Text & vbCrLf & "<TpoCobCod>300</TpoCobCod>"
                  Text1.Text = Text1.Text & vbCrLf & "<InstDptoCod>3</InstDptoCod>"
                  Text1.Text = Text1.Text & vbCrLf & "</Afiliados.Afiliado.InstitucionItem>"
                  Text1.Text = Text1.Text & vbCrLf & "</Instituciones>"
                  Text1.Text = Text1.Text & vbCrLf & "</Afiliados.Afiliado>"
               Else
                  If Data1.Recordset("cedula") > 0 Then
                    Text1.Text = Text1.Text & "<Afiliados.Afiliado>"
                    Text1.Text = Text1.Text & vbCrLf & "<TipoDocumento>1</TipoDocumento>"
                    If IsNull(Data1.Recordset("cedula")) = False Then
                       Text1.Text = Text1.Text & vbCrLf & "<NroDocumento>0</NroDocumento>"
                       Text1.Text = Text1.Text & vbCrLf & "<CI>0</CI>"
                    Else
                       Text1.Text = Text1.Text & vbCrLf & "<NroDocumento>0</NroDocumento>"
                       Text1.Text = Text1.Text & vbCrLf & "<CI>0</CI>"
                    End If
                    Data2.RecordSource = "Select * from clientes where cl_codigo =" & Data1.Recordset("cliente")
                    Data2.Refresh
                    If Data2.Recordset.RecordCount > 0 Then
                       If IsNull(Data2.Recordset("cl_sexo")) = False Then
                          If Data2.Recordset("cl_sexo") = 1 Then
                             Text1.Text = Text1.Text & vbCrLf & "<Sexo>M</Sexo>"
                          Else
                             Text1.Text = Text1.Text & vbCrLf & "<Sexo>F</Sexo>"
                          End If
                       Else
                          Text1.Text = Text1.Text & vbCrLf & "<Sexo>F</Sexo>"
                       End If
                    Else
                       Text1.Text = Text1.Text & vbCrLf & "<Sexo>F</Sexo>"
                    End If
                    If IsNull(Data1.Recordset("fecha_nac")) = False Then
                       Text1.Text = Text1.Text & vbCrLf & "<FchNac>" & Format(Data1.Recordset("fecha_nac"), "yyyy-mm-dd") & "</FchNac>"
                    Else
                       Text1.Text = Text1.Text & vbCrLf & "<FchNac>1900-01-01</FchNac>"
                    End If
                    If IsNull(Data1.Recordset("apellidos")) = False Then
                       Xlargo = Len(Trim(Data1.Recordset("apellidos")))
                       Xlargo2 = Xlargo / 2
                       Xlargo3 = Xlargo2 + 1
                       Text1.Text = Text1.Text & vbCrLf & "<NomPri>" & Mid(Data1.Recordset("apellidos"), Xlargo3, Xlargo2) & "</NomPri>"
                       Text1.Text = Text1.Text & vbCrLf & "<ApePri>" & Mid(Data1.Recordset("apellidos"), 1, Xlargo2) & "</ApePri>"
                    Else
                       Text1.Text = Text1.Text & vbCrLf & "<NomPri>NN</NomPri>"
                       Text1.Text = Text1.Text & vbCrLf & "<ApePri>NN</ApePri>"
                    End If
                    Text1.Text = Text1.Text & vbCrLf & "<PaisCodDir>UY</PaisCodDir>"
                    Text1.Text = Text1.Text & vbCrLf & "<DptoCodDir>16</DptoCodDir>"
                    If Data2.Recordset.RecordCount > 0 Then
                      If IsNull(Data2.Recordset("cl_grupo")) = False Then
                         If Data2.Recordset("cl_grupo") >= 101 And Data2.Recordset("cl_grupo") <= 104 Then
                            Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>627</LocCodDir>"
                         Else
                            If Data2.Recordset("cl_grupo") >= 201 And Data2.Recordset("cl_grupo") <= 209 Then
                               Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>724</LocCodDir>"
                            Else
                               If Data2.Recordset("cl_grupo") >= 300 And Data2.Recordset("cl_grupo") <= 322 Then
                                  Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>729</LocCodDir>"
                               Else
                                  If Data2.Recordset("cl_grupo") >= 400 And Data2.Recordset("cl_grupo") <= 419 Then
                                     Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>621</LocCodDir>"
                                  Else
                                     If Data2.Recordset("cl_grupo") >= 500 And Data2.Recordset("cl_grupo") <= 501 Then
                                        Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>630</LocCodDir>"
                                     Else
                                        If Data2.Recordset("cl_grupo") = 630 Then
                                           Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>521</LocCodDir>"
                                        Else
                                           If Data2.Recordset("cl_grupo") = 650 Or Data2.Recordset("cl_grupo") = 800 Then
                                              Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>524</LocCodDir>"
                                           Else
                                              If Data2.Recordset("cl_grupo") >= 670 And Data2.Recordset("cl_grupo") <= 679 Then
                                                 Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>528</LocCodDir>"
                                              Else
                                                 If Data2.Recordset("cl_grupo") >= 600 And Data2.Recordset("cl_grupo") <= 640 Then
                                                    Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>322</LocCodDir>"
                                                 Else
                                                    If Data2.Recordset("cl_grupo") >= 700 And Data2.Recordset("cl_grupo") <= 722 Then
                                                       Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>834</LocCodDir>"
                                                    Else
                                                        If Data2.Recordset("cl_grupo") >= 801 And Data2.Recordset("cl_grupo") <= 815 Then
                                                           Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>631</LocCodDir>"
                                                        Else
                                                           Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>627</LocCodDir>"
                                                        End If
                                                    End If
                                                 End If
                                              End If
                                           End If
                                        End If
                                     End If
                                  End If
                               End If
                            End If
                         End If
                      End If
                    Else
                      Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>627</LocCodDir>"
                    End If
                    Text1.Text = Text1.Text & vbCrLf & "<Instituciones>"
                    Text1.Text = Text1.Text & vbCrLf & "<Afiliados.Afiliado.InstitucionItem>"
                    If IsNull(Data1.Recordset("fecha_ing")) = False Then
                       Text1.Text = Text1.Text & vbCrLf & "<InstFchReg>" & Format(Data1.Recordset("fecha_ing"), "yyyy-mm-dd") & "</InstFchReg>"
                    Else
                       Text1.Text = Text1.Text & vbCrLf & "<InstFchReg>" & Format(Date, "yyyy-mm-dd") & "</InstFchReg>"
                    End If
                    Text1.Text = Text1.Text & vbCrLf & "<InstCod>303</InstCod>"
                    Text1.Text = Text1.Text & vbCrLf & "<TpoCobCod>300</TpoCobCod>"
                    Text1.Text = Text1.Text & vbCrLf & "<InstDptoCod>16</InstDptoCod>"
                    Text1.Text = Text1.Text & vbCrLf & "</Afiliados.Afiliado.InstitucionItem>"
                    Text1.Text = Text1.Text & vbCrLf & "</Instituciones>"
                    Text1.Text = Text1.Text & vbCrLf & "</Afiliados.Afiliado>"
                 End If
               End If
            Else
              Text1.Text = Text1.Text & "<Afiliados.Afiliado>"
              Text1.Text = Text1.Text & vbCrLf & "<TipoDocumento>1</TipoDocumento>"
              If IsNull(Data1.Recordset("cedula")) = False Then
                 Text1.Text = Text1.Text & vbCrLf & "<NroDocumento>0</NroDocumento>"
                 Text1.Text = Text1.Text & vbCrLf & "<CI>0</CI>"
              Else
                 Text1.Text = Text1.Text & vbCrLf & "<NroDocumento>0</NroDocumento>"
                 Text1.Text = Text1.Text & vbCrLf & "<CI>0</CI>"
              End If
              Data2.RecordSource = "Select * from clientes where cl_codigo =" & Data1.Recordset("cliente")
              Data2.Refresh
              If Data2.Recordset.RecordCount > 0 Then
                 If IsNull(Data2.Recordset("cl_sexo")) = False Then
                    If Data2.Recordset("cl_sexo") = 1 Then
                       Text1.Text = Text1.Text & vbCrLf & "<Sexo>M</Sexo>"
                    Else
                       Text1.Text = Text1.Text & vbCrLf & "<Sexo>F</Sexo>"
                    End If
                 Else
                    Text1.Text = Text1.Text & vbCrLf & "<Sexo>F</Sexo>"
                 End If
              Else
                 Text1.Text = Text1.Text & vbCrLf & "<Sexo>F</Sexo>"
              End If
              If IsNull(Data1.Recordset("fecha_nac")) = False Then
                 Text1.Text = Text1.Text & vbCrLf & "<FchNac>" & Format(Data1.Recordset("fecha_nac"), "yyyy-mm-dd") & "</FchNac>"
              Else
                 Text1.Text = Text1.Text & vbCrLf & "<FchNac>1900-01-01</FchNac>"
              End If
              If IsNull(Data1.Recordset("apellidos")) = False Then
                 Xlargo = Len(Trim(Data1.Recordset("apellidos")))
                 Xlargo2 = Xlargo / 2
                 Xlargo3 = Xlargo2 + 1
                 Text1.Text = Text1.Text & vbCrLf & "<NomPri>" & Mid(Data1.Recordset("apellidos"), Xlargo3, Xlargo2) & "</NomPri>"
                 Text1.Text = Text1.Text & vbCrLf & "<ApePri>" & Mid(Data1.Recordset("apellidos"), 1, Xlargo2) & "</ApePri>"
              Else
                 Text1.Text = Text1.Text & vbCrLf & "<NomPri>NN</NomPri>"
                 Text1.Text = Text1.Text & vbCrLf & "<ApePri>NN</ApePri>"
              End If
              Text1.Text = Text1.Text & vbCrLf & "<PaisCodDir>UY</PaisCodDir>"
              Text1.Text = Text1.Text & vbCrLf & "<DptoCodDir>16</DptoCodDir>"
                If Data2.Recordset.RecordCount > 0 Then
                  If IsNull(Data2.Recordset("cl_grupo")) = False Then
                     If Data2.Recordset("cl_grupo") >= 101 And Data2.Recordset("cl_grupo") <= 104 Then
                        Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>627</LocCodDir>"
                     Else
                        If Data2.Recordset("cl_grupo") >= 201 And Data2.Recordset("cl_grupo") <= 209 Then
                           Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>724</LocCodDir>"
                        Else
                           If Data2.Recordset("cl_grupo") >= 300 And Data2.Recordset("cl_grupo") <= 322 Then
                              Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>729</LocCodDir>"
                           Else
                              If Data2.Recordset("cl_grupo") >= 400 And Data2.Recordset("cl_grupo") <= 419 Then
                                 Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>621</LocCodDir>"
                              Else
                                 If Data2.Recordset("cl_grupo") >= 500 And Data2.Recordset("cl_grupo") <= 501 Then
                                    Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>630</LocCodDir>"
                                 Else
                                    If Data2.Recordset("cl_grupo") = 630 Then
                                       Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>521</LocCodDir>"
                                    Else
                                       If Data2.Recordset("cl_grupo") = 650 Or Data2.Recordset("cl_grupo") = 800 Then
                                          Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>524</LocCodDir>"
                                       Else
                                          If Data2.Recordset("cl_grupo") >= 670 And Data2.Recordset("cl_grupo") <= 679 Then
                                             Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>528</LocCodDir>"
                                          Else
                                             If Data2.Recordset("cl_grupo") >= 600 And Data2.Recordset("cl_grupo") <= 640 Then
                                                Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>322</LocCodDir>"
                                             Else
                                                If Data2.Recordset("cl_grupo") >= 700 And Data2.Recordset("cl_grupo") <= 722 Then
                                                   Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>834</LocCodDir>"
                                                Else
                                                    If Data2.Recordset("cl_grupo") >= 801 And Data2.Recordset("cl_grupo") <= 815 Then
                                                       Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>631</LocCodDir>"
                                                    Else
                                                       Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>627</LocCodDir>"
                                                    End If
                                                End If
                                             End If
                                          End If
                                       End If
                                    End If
                                 End If
                              End If
                           End If
                        End If
                     End If
                  End If
                Else
                  Text1.Text = Text1.Text & vbCrLf & "<LocCodDir>627</LocCodDir>"
                End If
              Text1.Text = Text1.Text & vbCrLf & "<Instituciones>"
              Text1.Text = Text1.Text & vbCrLf & "<Afiliados.Afiliado.InstitucionItem>"
              If IsNull(Data1.Recordset("fecha_ing")) = False Then
                 Text1.Text = Text1.Text & vbCrLf & "<InstFchReg>" & Format(Data1.Recordset("fecha_ing"), "yyyy-mm-dd") & "</InstFchReg>"
              Else
                 Text1.Text = Text1.Text & vbCrLf & "<InstFchReg>" & Format(Date, "yyyy-mm-dd") & "</InstFchReg>"
              End If
              Text1.Text = Text1.Text & vbCrLf & "<InstCod>303</InstCod>"
              Text1.Text = Text1.Text & vbCrLf & "<TpoCobCod>300</TpoCobCod>"
              Text1.Text = Text1.Text & vbCrLf & "<InstDptoCod>16</InstDptoCod>"
              Text1.Text = Text1.Text & vbCrLf & "</Afiliados.Afiliado.InstitucionItem>"
              Text1.Text = Text1.Text & vbCrLf & "</Instituciones>"
              Text1.Text = Text1.Text & vbCrLf & "</Afiliados.Afiliado>"
            End If
         End If
      End If
      Print #1, Text1.Text
      Text1.Text = ""
      Data1.Recordset.MoveNext
   Loop
   Text1.Text = ""
   Text1.Text = "</Afiliados>"
   Label2.Visible = False
   Print #1, Text1.Text
   Close #1
   objDOM.Load (App.Path & "\" & Xarch & ".xml")
'   objDOM.Save (App.Path & "\hc.xml")
End If
MsgBox "Terminado. Total de Registros procesados:" & XcantReg & ". El archivo fue guardado en la carpeta del sistema SAPP.", vbInformation

End Sub

Private Sub Command2_Click()
Dim objDOM As MSXML2.DOMDocument
Dim objModule As MSXML2.DOMDocument
Set objDOM = New MSXML2.DOMDocument
Dim objRootElem As MSXML2.IXMLDOMElement
Set objModule = New MSXML2.DOMDocument
Set XMLInstruccion = objDOM.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
objDOM.appendChild XMLInstruccion
 
         Set objRootElem = objDOM.createElement("Afiliados")
 
         objRootElem.setAttribute "xmlns", "MSPRUCAF"
         objDOM.appendChild objRootElem
         objDOM.selectSingleNode("//Afiliados").appendChild objDOM.createElement("Afiliados.Afiliado")
         objDOM.selectSingleNode("//Afiliados/Afiliados.Afiliado").appendChild objDOM.createElement("Tipo")
         objDOM.selectSingleNode("//Afiliados/Afiliados.Afiliado/Tipo").Text = "1"
         objDOM.Save (App.Path & "\hc.xml")
 
'  Set objRootElem = objDom.createElement("BCE:Balanza")
 
'         objRootElem.setAttribute "xmlns:BCE", "www.sat.gob.mx/esquemas/ContabilidadE/1_1/BalanzaComprobacion"
'         objRootElem.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
'         objRootElem.setAttribute "xsi:schemaLocation", "www.sat.gob.mx/esquemas/ContabilidadE/1_1/BalanzaComprobacion http://www.sat.gob.mx/esquemas/ContabilidadE/1_1/BalanzaComprobacion/BalanzaComprobacion_1_1.xsd"
'         objDom.appendChild objRootElem
         Call XMLHeader
 
         Fila = 1
 
         With MSHFlexGrid1
              Do While .TextMatrix(Fila, 0) <> vbNullString
 
                'create new element and attribute by calling to CreateXMLBalanza function.
                objModule.loadXML (CreateXMLBalanza(.TextMatrix(Fila, 0), .TextMatrix(Fila, 1), .TextMatrix(Fila, 2), .TextMatrix(Fila, 3), .TextMatrix(Fila, 4)))
 
                Set objNode = objModule.FirstChild
                objDOM.documentElement.appendChild objNode
 
                Fila = Fila + 1
 
              Loop
 
              'Saves the xml document in c:\temp directory
                 NombreBalanza = "XXAX010101X01" + Combo2.Text + Combo1.Text + "B" + Combo3.Text + ".XML"
              objDOM.Save (NombreBalanza)
              MsgBox " SE HA GENERADO EXISTOSAMENTE EL ARCHIVO EN: " & NombreBalanza, vbInformation, "GENERAR BALANZA XML"
 
         End With
End Sub
 
Public Function XMLHeader()
Dim objDOM As MSXML2.DOMDocument
Dim objRootElem As MSXML2.IXMLDOMElement
Dim objDocAttribute As IXMLDOMAttribute

         Set objDocAttribute = objDOM.createAttribute("Version")
         objDocAttribute.nodeValue = "1.1"
         objRootElem.setAttributeNode objDocAttribute
 
         Set objDocAttribute = objDOM.createAttribute("RFC")
         objDocAttribute.nodeValue = "XXAX010101X01"
         objRootElem.setAttributeNode objDocAttribute
 
         Set objDocAttribute = objDOM.createAttribute("Mes")
         objDocAttribute.nodeValue = FrmBalanzaXML.Combo1.Text
         objRootElem.setAttributeNode objDocAttribute
 
         Set objDocAttribute = objDOM.createAttribute("Anio")
         objDocAttribute.nodeValue = FrmBalanzaXML.Combo2.Text
         objRootElem.setAttributeNode objDocAttribute
 
         If OptionXML = 1 Then
            Set objDocAttribute = objDOM.createAttribute("TipoEnvio")
            objDocAttribute.nodeValue = FrmBalanzaXML.Combo3.Text
            objRootElem.setAttributeNode objDocAttribute
 
            Set objDocAttribute = objDOM.createAttribute("FechaModBal")
            objDocAttribute.nodeValue = "2015-04-30" 'FrmBalanzaXML.Combo3.Text
            objRootElem.setAttributeNode objDocAttribute
 
         End If
 
End Function
Public Function CreateXMLBalanza(strNumCta As String, strSaldoIni As String, _
    strDebe As String, strHaber As String, strSaldoFin As String) As String
 
   Set objDomBal = New DOMDocument
 
   Set objRootElem = objDomBal.createElement("Ctas")
   objDomBal.appendChild objRootElem
 
   Set objMbrAttribute = objDomBal.createAttribute("NumCta")
   objMbrAttribute.nodeValue = strNumCta
   objRootElem.setAttributeNode objMbrAttribute
 
   Set objMbrAttribute = objDomBal.createAttribute("SaldoIni")
   objMbrAttribute.nodeValue = strSaldoIni
   objRootElem.setAttributeNode objMbrAttribute
 
   Set objMbrAttribute = objDomBal.createAttribute("Debe")
   objMbrAttribute.nodeValue = strDebe
   objRootElem.setAttributeNode objMbrAttribute
 
   Set objMbrAttribute = objDomBal.createAttribute("Haber")
   objMbrAttribute.nodeValue = strHaber
   objRootElem.setAttributeNode objMbrAttribute
 
   Set objMbrAttribute = objDomBal.createAttribute("SaldoFin")
   objMbrAttribute.nodeValue = strSaldoFin
   objRootElem.setAttributeNode objMbrAttribute
 
   CreateXMLBalanza = objDomBal.XML
 
End Function

Private Sub Command3_Click()
Dim objDOM As New MSXML2.DOMDocument30
   Dim objNode As MSXML2.IXMLDOMNode
   Dim objPerson As MSXML2.IXMLDOMNode
   Dim objPerson2 As MSXML2.IXMLDOMNode
   
   Dim objGrandChildNode As MSXML2.IXMLDOMNode
   Dim objGrandChildNode2 As MSXML2.IXMLDOMNode
   
   Dim objAttribute As MSXML2.IXMLDOMAttribute
   Dim objElement As MSXML2.IXMLDOMElement
   
   ' Create the main xml node
'   Set objNode = objDOM.createNode(NODE_PROCESSING_INSTRUCTION, "xml", "version='1.0' encoding='ISO-8859-1'")
'   objDOM.appendChild objNode
Set XMLInstruccion = objDOM.createProcessingInstruction("xml", "version='1.0' encoding='ISO-8859-1'")
objDOM.appendChild XMLInstruccion
   
 
   '      Set objRootElem = objDOM.createElement("Afiliados")
 
   '      objRootElem.setAttribute "xmlns", "MSPRUCAF"
   '      objDOM.appendChild objRootElem
   
   
   Set objNode = objDOM.createNode(NODE_ELEMENT, "Afiliados", "")

   
   
   Dim i As Integer
   
   For i = 1 To 4
        Set objPerson = objDOM.createNode(NODE_ELEMENT, "Afiliados.Afiliado", "")
        objNode.appendChild objPerson
        
        Set objGrandChildNode = objDOM.createNode(NODE_ELEMENT, "age", "")
        objGrandChildNode.Text = "some data"
        objPerson.appendChild objGrandChildNode
        
        Set objGrandChildNode = objDOM.createNode(NODE_ELEMENT, "name", "")
        objGrandChildNode.Text = "some data"
        objPerson.appendChild objGrandChildNode
        
        Set objGrandChildNode = objDOM.createNode(NODE_ELEMENT, "address", "")
        objGrandChildNode.Text = "some data"
        objPerson.appendChild objGrandChildNode
        
        
   
   Next i

   ''Set objDetail = objDOM.createNode(NODE_ELEMENT, "detail", "")
   ''objNode.appendChild objDetail
  
   objDOM.appendChild objNode
   Set objNode = Nothing
   
   MsgBox objDOM.XML
   objDOM.Save (App.Path & "\hc.xml")
   
   Set objDOM = Nothing
End Sub

Private Sub Command4_Click()
Dim objDOM As MSXML2.DOMDocument30
Dim objModule As MSXML2.DOMDocument30
Set objDOM = New MSXML2.DOMDocument30
Dim objRootElem As MSXML2.IXMLDOMElement
Set objModule = New MSXML2.DOMDocument30
Set XMLInstruccion = objDOM.createProcessingInstruction("xml", "version='1.0' encoding='ISO-8859-1'")
objDOM.appendChild XMLInstruccion
 
'         Set objRootElem = objDOM.createElement("Afiliados")
 
'         objRootElem.setAttribute "xmlns", "MSPRUCAF"
'         objDOM.appendChild objRootElem
         Text1.Text = "<?xml version=""1.0"" encoding=""ISO-8859-1""?>"
         Text1.Text = Text1.Text & "<Afiliados xmlns=""RUCAF"">"
         Text1.Text = Text1.Text & "<Afiliados.Afiliado>"
         Text1.Text = Text1.Text & vbCrLf & "<Tipo>1</Tipo>"
         Text1.Text = Text1.Text & vbCrLf & "<Cedula>348484</Cedula>"
         Text1.Text = Text1.Text & vbCrLf & "</Afiliados.Afiliado>"
         
         Text1.Text = Text1.Text & vbCrLf & "<Afiliados.Afiliado>"
         Text1.Text = Text1.Text & vbCrLf & "<Tipo>1</Tipo>"
         Text1.Text = Text1.Text & vbCrLf & "<Cedula>5555</Cedula>"
         Text1.Text = Text1.Text & vbCrLf & "</Afiliados.Afiliado>"
         Text1.Text = Text1.Text & "</Afiliados>"
         Open App.Path & "\mit.xml" For Output As #1
         Print #1, Text1.Text

         Close #1
         objDOM.Load (App.Path & "\mit.xml")

'         objDOM.selectSingleNode("//Afiliados").appendChild objDOM.createElement("Afiliados.Afiliado")
         
'         objDOM.selectSingleNode("//Afiliados/Afiliados.Afiliado").appendChild objDOM.createNode(NODE_ELEMENT, "Tipo", "")
'         objDOM.selectSingleNode("//Afiliados/Afiliados.Afiliado/Tipo").Text = "1"
'         objDOM.selectSingleNode("//Afiliados/Afiliados.Afiliado").appendChild objDOM.createNode(NODE_ELEMENT, "TipoDocumento", "")
'         objDOM.selectSingleNode("//Afiliados/Afiliados.Afiliado/TipoDocumento").Text = "1"
'         objDOM.selectSingleNode("//Afiliados/Afiliados.Afiliado").appendChild objDOM.createNode(NODE_ELEMENT, "NroDocumento", "")
'         objDOM.selectSingleNode("//Afiliados/Afiliados.Afiliado/NroDocumento").Text = "34805844"
'
'         objDOM.selectSingleNode("//Afiliados").appendChild objDOM.createElement("Afiliados.Afiliado")
'         objDOM.selectSingleNode("//Afiliados/Afiliados.Afiliado").appendChild objDOM.createNode(NODE_ELEMENT, "Tipo", "")
'         objDOM.selectSingleNode("//Afiliados/Afiliados.Afiliado/Tipo").Text = "1"
'         objDOM.selectSingleNode("//Afiliados/Afiliados.Afiliado").appendChild objDOM.createNode(NODE_ELEMENT, "TipoDocumento", "")
'         objDOM.selectSingleNode("//Afiliados/Afiliados.Afiliado/TipoDocumento").Text = "1"
'         objDOM.selectSingleNode("//Afiliados/Afiliados.Afiliado").appendChild objDOM.createNode(NODE_ELEMENT, "NroDocumento", "")
'         objDOM.selectSingleNode("//Afiliados/Afiliados.Afiliado/NroDocumento").Text = "8888"
         

         objDOM.Save (App.Path & "\hc.xml")

End Sub

Private Sub Form_Load()
Data1.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data2.Connect = "odbc;dsn=" & Xconexrmt & ";"
Data3.Connect = "odbc;dsn=" & Xconexrmt & ";"


End Sub

Private Sub t_anio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Command1.SetFocus
End If

End Sub

Private Sub t_mes_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   t_anio.SetFocus
End If

End Sub
