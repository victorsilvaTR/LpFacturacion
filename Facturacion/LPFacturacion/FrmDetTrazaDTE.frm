VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmDetTrazaDTE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle Estado DTE"
   ClientHeight    =   10125
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10125
   ScaleWidth      =   9870
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_CopyLinkTraza 
      Caption         =   "Copiar enlace a Traza"
      Height          =   1035
      Left            =   8400
      Picture         =   "FrmDetTrazaDTE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Frame Frame6 
      Caption         =   "Avisos DTE anteriores u otros avisos"
      Height          =   1395
      Left            =   180
      TabIndex        =   19
      Top             =   8400
      Width           =   7935
      Begin MSFlexGridLib.MSFlexGrid Gr_Avisos 
         Height          =   1035
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   1826
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   " Intercambio "
      Height          =   1395
      Left            =   180
      TabIndex        =   16
      Top             =   6840
      Width           =   7935
      Begin MSFlexGridLib.MSFlexGrid Gr_Intercam 
         Height          =   1035
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   1826
         _Version        =   393216
      End
   End
   Begin VB.CommandButton Bt_VerTraza 
      Caption         =   "Ver Detalle Completo"
      Height          =   1035
      Left            =   8400
      Picture         =   "FrmDetTrazaDTE.frx":05C5
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2340
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   180
      TabIndex        =   15
      Top             =   1260
      Width           =   7935
      Begin VB.Label Lb_EventoFinal 
         Caption         =   "Evento Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   2040
         TabIndex        =   18
         Top             =   300
         Width           =   3855
      End
      Begin VB.Label Label3 
         Caption         =   "Evento Final:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   300
         Width           =   1635
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "  SII "
      Height          =   2175
      Left            =   180
      TabIndex        =   14
      Top             =   4560
      Width           =   7935
      Begin MSFlexGridLib.MSFlexGrid Gr_SII 
         Height          =   1815
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   3201
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Acepta "
      Height          =   2175
      Left            =   180
      TabIndex        =   13
      Top             =   2220
      Width           =   7935
      Begin MSFlexGridLib.MSFlexGrid Gr_Acepta 
         Height          =   1815
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   3201
         _Version        =   393216
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   180
      Picture         =   "FrmDetTrazaDTE.frx":0B8A
      ScaleHeight     =   675
      ScaleWidth      =   615
      TabIndex        =   12
      Top             =   300
      Width           =   615
   End
   Begin VB.CommandButton Bt_Cerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   8400
      TabIndex        =   6
      Top             =   360
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   1140
      TabIndex        =   7
      Top             =   240
      Width           =   6975
      Begin VB.TextBox Tx_Folio 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   240
         Width           =   1275
      End
      Begin VB.TextBox Tx_TipoDoc 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   3315
      End
      Begin VB.Label Label2 
         Caption         =   "Folio:"
         Height          =   255
         Left            =   4920
         TabIndex        =   10
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Documento:"
         Height          =   255
         Left            =   300
         TabIndex        =   8
         Top             =   300
         Width           =   1035
      End
   End
End
Attribute VB_Name = "FrmDetTrazaDTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const C_TIPO = 0
Const C_ESTADO = 1
Const C_FECHA = 2
Const C_LNKRESENVIOSII = 3
Const C_VERRESENVIOSII = 4
Const C_VERDET = 5

Const NCOLS = C_VERDET

Dim lTraza As AcpTraza_t
Dim lIdDTE As Long
Dim lidEstado As Integer, lTxtEstado As String
Friend Function FView(ByVal IdDTE As Long, Tr As AcpTraza_t, idEstado As Integer, TxtEstado As String)

   lIdDTE = IdDTE
   lTraza = Tr
   
   Me.Show vbModal
   
   idEstado = lidEstado
   TxtEstado = lTxtEstado
   
End Function

Private Sub bt_Cerrar_Click()
   Unload Me
End Sub

Private Sub Bt_CopyLinkTraza_Click()
   Dim Traza As String

   If lTraza.Url = "" Then
      MsgBox1 "No se encuentra disponible el detalle de estado del DTE.", vbInformation
      Exit Sub
   End If
   
   'Traza = ReplaceStr(lTraza.Url, "/v01/", "/traza/")
   Traza = Trim(ReplaceStr(lTraza.UrlData, "", ""))
   
   Call SetClipText(Traza)
   
End Sub

Private Sub Bt_VerTraza_Click()
   Dim Traza As String

   If lTraza.Url = "" Then
      MsgBox1 "No se encuentra disponible el detalle de estado del DTE.", vbInformation
      Exit Sub
   End If

   'Traza = ReplaceStr(lTraza.Url, "/v01/", "/traza/")
   Traza = Trim(ReplaceStr(lTraza.UrlData, "", ""))
   
   DoEvents
      
   Call ShellExecute(Me.hWnd, "open", Traza, "", "", 1)
   DoEvents
      

End Sub

Private Sub Form_Load()
   
   Call SetUpGrid(Gr_Acepta)
   Call SetUpGrid(Gr_SII)
   Call SetUpGrid(Gr_Intercam)
   Call SetUpGrid(Gr_Avisos)
   
   Call LoadAll
   
End Sub

Private Sub SetUpGrid(Gr As MSFlexGrid)
   Dim i As Integer

   Gr.Cols = NCOLS + 1
   Call FGrSetup(Gr)
   
   Gr.ColWidth(C_TIPO) = 2000
   Gr.ColWidth(C_ESTADO) = 1300
   Gr.ColWidth(C_FECHA) = 1700
   Gr.ColWidth(C_LNKRESENVIOSII) = 0
   Gr.ColWidth(C_VERRESENVIOSII) = 1000
   Gr.ColWidth(C_VERDET) = 1300

   Gr.ColAlignment(C_FECHA) = flexAlignRightCenter
   Gr.ColAlignment(C_VERRESENVIOSII) = flexAlignCenterCenter
   Gr.ColAlignment(C_VERDET) = flexAlignCenterCenter
   
   Gr.TextMatrix(0, C_TIPO) = "Evento"
   Gr.TextMatrix(0, C_ESTADO) = "Resultado"
   Gr.TextMatrix(0, C_FECHA) = "Fecha - Hora"
   Gr.TextMatrix(0, C_VERRESENVIOSII) = "Ver detalle"
   Gr.TextMatrix(0, C_VERDET) = "Observaciones"
   
   Call FGrVRows(Gr, 1)
   
   
End Sub
Private Sub LoadAll()
   Dim Q1 As String
   Dim Rs As Recordset
   Dim i As Integer
   Dim EventoFinal As String
   Dim LlegoAlSII As Boolean
   Dim HayError As Boolean
   Dim EstadoDTE As Integer
   Dim EstadoDTESII As Integer
   Dim Msg As String, Msg1 As String
   Dim Idx As Integer, Idx1 As Integer
   Dim href1 As String, href2 As String

   If lIdDTE > 0 Then
      Q1 = "SELECT TipoLib, TipoDoc, Folio, CodDocSII FROM DTE WHERE IdDTE= " & lIdDTE & " AND IdEmpresa = " & gEmpresa.Id
      Set Rs = OpenRs(DbMain, Q1)
      
'      Tx_TipoDoc = gTipoDoc(GetTipoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")))).Diminutivo & " - " & gTipoDoc(GetTipoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")))).Nombre & " (" & vFld(Rs("CodDocSII")) & ")"
      
      If vFld(Rs("TipoLib")) = LIB_OTROS And vFld(Rs("TipoDoc")) = TIPODOC_GUIADESPACHO Then
         Tx_TipoDoc = gTipoDocDTE(IDXTIPODOCDTE_GUIADESPACHO).Diminutivo & " - " & gTipoDocDTE(IDXTIPODOCDTE_GUIADESPACHO).Nombre & " (" & vFld(Rs("CodDocSII")) & ")"
      Else
         Tx_TipoDoc = gTipoDoc(GetTipoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")))).Diminutivo & " - " & gTipoDoc(GetTipoDoc(vFld(Rs("TipoLib")), vFld(Rs("TipoDoc")))).Nombre & " (" & vFld(Rs("CodDocSII")) & ")"
      End If

      
      Tx_Folio = vFld(Rs("Folio"))
     
      Call CloseRs(Rs)
      
   End If

   Gr_Acepta.rows = lTraza.nAcepta + Gr_Acepta.FixedRows
   EventoFinal = ""
   For i = 0 To lTraza.nAcepta - 1
      Gr_Acepta.TextMatrix(i + Gr_Acepta.FixedRows, C_TIPO) = lTraza.Acepta(i).Tipo
      Gr_Acepta.TextMatrix(i + Gr_Acepta.FixedRows, C_ESTADO) = StrEstadoTraza(lTraza.Acepta(i).Estado)
      If lTraza.Acepta(i).Estado = ET_ERR Then
         Call FGrForeColor(Gr_Acepta, i + Gr_Acepta.FixedRows, C_ESTADO, vbRed)
         HayError = True
      End If
      Gr_Acepta.TextMatrix(i + Gr_Acepta.FixedRows, C_FECHA) = Format(lTraza.Acepta(i).Fecha, "dd mmm yy hh:mm:ss")
      If lTraza.Acepta(i).Obs <> "" Then
         Gr_Acepta.Row = i + Gr_Acepta.FixedRows
         Gr_Acepta.Col = C_VERDET
         Gr_Acepta.CellPictureAlignment = flexAlignCenterCenter
         Set Gr_Acepta.CellPicture = FrmMain.Pc_Lupa2
      End If
'      Gr_SII.TextMatrix(i + Gr_SII.FixedRows, C_VERDET) = IIf(lTraza.Acepta(i).Obs <> "", "Ver", "")

      EventoFinal = lTraza.Acepta(i).Tipo
   Next i
   
   Gr_SII.rows = lTraza.nSII + Gr_SII.FixedRows
   EventoFinal = ""
   For i = 0 To lTraza.nSII - 1
      Gr_SII.TextMatrix(i + Gr_SII.FixedRows, C_TIPO) = lTraza.SII(i).Tipo
      Gr_SII.TextMatrix(i + Gr_SII.FixedRows, C_ESTADO) = StrEstadoTraza(lTraza.SII(i).Estado)
      If lTraza.SII(i).Estado = ET_ERR Then
         Call FGrForeColor(Gr_SII, i + Gr_SII.FixedRows, C_ESTADO, vbRed)
         HayError = True
      End If
      
      Gr_SII.TextMatrix(i + Gr_SII.FixedRows, C_FECHA) = Format(lTraza.SII(i).Fecha, "dd mmm yy hh:mm:ss")
      Gr_SII.TextMatrix(i + Gr_SII.FixedRows, C_LNKRESENVIOSII) = lTraza.SII(i).UrlTipo
      
      If lTraza.SII(i).UrlTipo <> "" Then
         Gr_SII.Row = i + Gr_SII.FixedRows
         Gr_SII.Col = C_VERRESENVIOSII
         Gr_SII.CellPictureAlignment = flexAlignCenterCenter
         Set Gr_SII.CellPicture = FrmMain.Pc_Lupa2
      End If
'      Gr_SII.TextMatrix(i + Gr_SII.FixedRows, C_VERRESENVIOSII) = IIf(lTraza.SII(i).UrlTipo <> "", "Ver", "")
      
      If lTraza.SII(i).Obs <> "" Then
         Gr_SII.Row = i + Gr_SII.FixedRows
         Gr_SII.Col = C_VERDET
         Gr_SII.CellPictureAlignment = flexAlignCenterCenter
         Set Gr_SII.CellPicture = FrmMain.Pc_Lupa2
      End If
'      Gr_SII.TextMatrix(i + Gr_SII.FixedRows, C_VERDET) = IIf(lTraza.SII(i).Obs <> "", "Ver", "")
   
      ' 24 feb 2021: se pone este IF porque el Aceptado no viene al final
      If EventoFinal <> gTxtEstadoDTESII(EDTESII_ACEPTADO) And InStr(1, EventoFinal, gTxtEstadoDTESII(EDTESII_REPARO), vbTextCompare) <= 0 Then
         EventoFinal = lTraza.SII(i).Tipo
      End If
'      If EventoFinal <> gTxtEstadoDTESII(EDTESII_ACEPTADO) And InStr(1, EventoFinal, gTxtEstadoDTESII(EDTESII_REPARO), vbTextCompare) <= 0 Then
'         EventoFinal = lTraza.SII(i).Tipo
'      End If
      
      LlegoAlSII = True
   Next i
   
   Lb_EventoFinal = EventoFinal
   If HayError Then
      Lb_EventoFinal.ForeColor = vbRed
   End If
   
   EstadoDTESII = 0
   EstadoDTE = EDTE_ENVIADO
   Call AddLog("Paso 5", 2)
   If LlegoAlSII Then
      For i = 1 To MAX_ESTADODTESII
         'If InStr(1, EventoFinal, gTxtEstadoDTESII(i), vbTextCompare) > 0 Then
         If InStr(1, EventoFinal, gEstadoDTESII(i), vbTextCompare) > 0 Then
            
            EstadoDTESII = i
            
'            If EventoFinal = gTxtEstadoDTESII(EDTESII_PROCESADO) Then
'               EstadoDTE = EDTE_PROCESADO
'            ElseIf EventoFinal = gTxtEstadoDTESII(EDTESII_ACEPTADO) Or InStr(1, EventoFinal, gTxtEstadoDTESII(EDTESII_REPARO), vbTextCompare) Then
'               EstadoDTE = EDTE_EMITIDO
'            ElseIf EventoFinal = gTxtEstadoDTESII(EDTESII_RECHAZADO) Then
'               EstadoDTE = EDTE_ERROR
'            End If
            If EventoFinal = gTxtEstadoDTESII(EDTESII_PROCESADO) Then
               EstadoDTE = EDTE_PROCESADO
            ElseIf EventoFinal = gEstadoDTESII(EDTESII_ACEPTADO) Or EventoFinal = gEstadoDTESII(EDTESII_PAGADO) Or EventoFinal = gEstadoDTESII(EDTESII_ENVIADO) Or InStr(1, EventoFinal, gTxtEstadoDTESII(EDTESII_REPARO), vbTextCompare) Then
               EstadoDTE = EDTE_EMITIDO
            ElseIf EventoFinal = gTxtEstadoDTESII(EDTESII_RECHAZADO) Then
               EstadoDTE = EDTE_ERROR
            End If
            
            Exit For
         End If
      Next i
      
      If EstadoDTESII > 0 Then
         Q1 = "UPDATE DTE SET "
         If EstadoDTE <> 0 Then   'cambió de estado, entonces grabamos
            lTxtEstado = gEstadoDTE(EstadoDTE)
            lidEstado = EstadoDTE  ' 24 feb 2021
            Q1 = Q1 & " IdEstado = " & EstadoDTE & ","
         End If
         Q1 = Q1 & " IdEstadoSII = " & EstadoDTESII & " WHERE IdDTE = " & lIdDTE & " AND IdEmpresa = " & gEmpresa.Id
         Call ExecSQL(DbMain, Q1)
      End If
   End If

   Gr_Intercam.rows = lTraza.nIntercam + Gr_Intercam.FixedRows
   For i = 0 To lTraza.nIntercam - 1
      Gr_Intercam.TextMatrix(i + Gr_Intercam.FixedRows, C_TIPO) = lTraza.Intercam(i).Tipo
      Gr_Intercam.TextMatrix(i + Gr_Intercam.FixedRows, C_ESTADO) = StrEstadoTraza(lTraza.Intercam(i).Estado)
      Gr_Intercam.TextMatrix(i + Gr_Intercam.FixedRows, C_FECHA) = Format(lTraza.Intercam(i).Fecha, "dd mmm yy hh:mm:ss")
   Next i


   Gr_Avisos.rows = lTraza.nAvisos + Gr_Avisos.FixedRows
   For i = 0 To lTraza.nAvisos - 1
      Gr_Avisos.TextMatrix(i + Gr_Avisos.FixedRows, C_TIPO) = lTraza.Avisos(i).Tipo
      Gr_Avisos.TextMatrix(i + Gr_Avisos.FixedRows, C_ESTADO) = StrEstadoTraza(lTraza.Avisos(i).Estado)
      Gr_Avisos.TextMatrix(i + Gr_Avisos.FixedRows, C_FECHA) = Format(lTraza.Avisos(i).Fecha, "dd mmm yy hh:mm:ss")
      If lTraza.Avisos(i).Obs <> "" Then
         Gr_Avisos.Row = i + Gr_Avisos.FixedRows
         Gr_Avisos.Col = C_VERDET
         Gr_Avisos.CellPictureAlignment = flexAlignCenterCenter
         Set Gr_Avisos.CellPicture = FrmMain.Pc_Lupa2
         
         Msg = lTraza.Avisos(i).Obs
         Idx = InStr(Msg, "http://")
         Idx1 = InStr(Msg, """>http://")
         If Idx > 0 And Idx1 > 0 Then
            href1 = Mid(Msg, Idx, Idx1 - Idx)
            href2 = Mid(Msg, Idx1 + 2)
            If href1 <> "" Or href2 <> "" Then
               Gr_Avisos.Row = i + Gr_Avisos.FixedRows
               Gr_Avisos.Col = C_VERRESENVIOSII
               Gr_Avisos.CellPictureAlignment = flexAlignCenterCenter
               Set Gr_Avisos.CellPicture = FrmMain.Pc_Lupa2

               Gr_Avisos.TextMatrix(i + Gr_Avisos.FixedRows, C_LNKRESENVIOSII) = IIf(href2 <> "", href2, href1)
            End If
         End If

      End If

   Next i

   'If EstadoDTE = EDTE_EMITIDO And EstadoDTESII > 0 Then
   If EstadoDTESII > 0 Then
      lTxtEstado = lTxtEstado & "/" & gDesEstadoDTESII(EstadoDTESII)
   End If

   Call FGrVRows(Gr_Acepta)
   Call FGrVRows(Gr_SII)
   Call FGrVRows(Gr_Intercam)
   Call FGrVRows(Gr_Avisos)


End Sub
Private Sub Gr_Acepta_DblClick()

   Call GrDblClick(Gr_Acepta, lTraza.Acepta)

End Sub


Private Sub Gr_Avisos_DblClick()
   
   Call GrDblClick(Gr_Avisos, lTraza.Avisos)

End Sub

Private Sub Gr_SII_DblClick()

   Call GrDblClick(Gr_SII, lTraza.SII)
   
End Sub

Friend Sub GrDblClick(Grid As MSFlexGrid, DetTraza() As DetTraza_t)
   Dim Row As Integer
   Dim Col As Integer
   Dim Msg As String, Msg1 As String
   Dim Idx As Integer, Idx1 As Integer
   Dim href1 As String, href2 As String
   
   Row = Grid.Row
   Col = Grid.Col
   
   If Col = C_VERRESENVIOSII Then
      If Grid.TextMatrix(Row, C_LNKRESENVIOSII) <> "" Then
         DoEvents
      
         Call ShellExecute(Me.hWnd, "open", Grid.TextMatrix(Row, C_LNKRESENVIOSII), "", "", 1)
         DoEvents
      Else
         MsgBox1 "No hay detalle disponible.", vbInformation
      End If
      
   ElseIf Col = C_VERDET Then
      Msg = DetTraza(Row - Grid.FixedRows).Obs
      
      If Msg <> "" Then
         
'         Idx = InStr(Msg, "<a href=")
'         If Idx > 0 Then
'            Msg1 = Left(Msg, Idx - 1)
'            Msg1 = ReplaceStr(Msg1, "<br>", vbCrLf)
'            Msg1 = ReplaceStr(Msg1, "<br/>", vbCrLf)
'         End If
'         Idx = InStr(Msg, "http://")
'         Idx1 = InStr(Msg, """>http://")
'         If Idx > 0 And Idx1 > 0 Then
'            href1 = Mid(Msg, Idx, Idx1 - Idx)
'            href2 = Mid(Msg, Idx1 + 2)
'         End If
'         MsgBox1 "Observaciones:" & vbCrLf & vbCrLf & Msg1, vbInformation
            MsgBox1 "Observaciones:" & vbCrLf & vbCrLf & Msg, vbInformation
      Else
         MsgBox1 "No hay observaciones disponibles.", vbInformation
      End If
   
   End If
   
End Sub
