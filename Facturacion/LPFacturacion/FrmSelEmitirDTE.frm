VERSION 5.00
Begin VB.Form FrmSelEmitirDTE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Emitir Documento Tributario Electrónico (DTE)"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   7905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_Cancel 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6120
      TabIndex        =   12
      Top             =   1020
      Width           =   1215
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "Emitir"
      Height          =   375
      Left            =   6120
      TabIndex        =   11
      Top             =   540
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   240
      Picture         =   "FrmSelEmitirDTE.frx":0000
      ScaleHeight     =   570
      ScaleWidth      =   450
      TabIndex        =   13
      Top             =   540
      Width           =   450
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionar Tipo de Documento Electrónico"
      Height          =   3675
      Left            =   960
      TabIndex        =   0
      Top             =   420
      Width           =   4455
      Begin VB.OptionButton Op_TipoDoc 
         Caption         =   "Option Guia Despacho"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   1020
         Width           =   3735
      End
      Begin VB.OptionButton Op_TipoDoc 
         Caption         =   "Option1"
         Height          =   195
         Index           =   9
         Left            =   360
         TabIndex        =   10
         Top             =   3120
         Width           =   3735
      End
      Begin VB.OptionButton Op_TipoDoc 
         Caption         =   "Option1"
         Height          =   195
         Index           =   8
         Left            =   360
         TabIndex        =   9
         Top             =   2820
         Width           =   3735
      End
      Begin VB.OptionButton Op_TipoDoc 
         Caption         =   "Option1"
         Height          =   195
         Index           =   7
         Left            =   360
         TabIndex        =   8
         Top             =   2520
         Width           =   3735
      End
      Begin VB.OptionButton Op_TipoDoc 
         Caption         =   "Option1"
         Height          =   195
         Index           =   6
         Left            =   360
         TabIndex        =   7
         Top             =   2220
         Width           =   3735
      End
      Begin VB.OptionButton Op_TipoDoc 
         Caption         =   "Option1"
         Height          =   195
         Index           =   5
         Left            =   360
         TabIndex        =   6
         Top             =   1920
         Width           =   3735
      End
      Begin VB.OptionButton Op_TipoDoc 
         Caption         =   "Option1"
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   5
         Top             =   1620
         Width           =   3735
      End
      Begin VB.OptionButton Op_TipoDoc 
         Caption         =   "Option1"
         Height          =   195
         Index           =   3
         Left            =   360
         TabIndex        =   4
         Top             =   1320
         Width           =   3735
      End
      Begin VB.OptionButton Op_TipoDoc 
         Caption         =   "Option1"
         Height          =   195
         Index           =   2
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   3735
      End
      Begin VB.OptionButton Op_TipoDoc 
         Caption         =   "Option1"
         Height          =   195
         Index           =   1
         Left            =   360
         TabIndex        =   1
         Top             =   420
         Width           =   3735
      End
   End
End
Attribute VB_Name = "FrmSelEmitirDTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const IDX_GUIA = 100

Private Sub bt_Cancel_Click()
   Unload Me
End Sub

Private Sub bt_OK_Click()
   Dim Frm As FrmDTE, FrmAdm As FrmAdmDTE
   Dim FrmTRef As FrmTipoRef
   Dim i As Integer, OptSel As Integer, Rc As Integer, IdDTERef As Long
   Dim TipoNota As String
   Dim EstadoUltDTE As Integer
   Dim Q1 As String
   Dim Rs As Recordset
   Dim DiminutivoDocRef As String
   Dim EsNotaCredDebFactCompra As Boolean
   
   
   'si es NC vemos si selecciona un documento asociado
   For i = 0 To UBound(gTipoDocDTE)
   
      If gTipoDocDTE(i).IdxTipoDoc = 0 And gTipoDocDTE(i).CodDocDTESII = "" Then
         Exit For
      End If
      
      If Op_TipoDoc(i).Value <> 0 Then
           
         'vemos si el último documento, emitido del mismo tipo, ya está en estado Emitido, si no mamdamos una advertencia
         Q1 = "SELECT IdEstado FROM DTE "
         Q1 = Q1 & " WHERE TipoLib = " & gTipoDocDTE(i).TipoLib & " AND TipoDoc = " & gTipoDocDTE(i).TipoDoc & " AND CodDocSII = '" & gTipoDocDTE(i).CodDocDTESII & "'"
         Q1 = Q1 & " ORDER BY IdDTE desc"
         
         Set Rs = OpenRs(DbMain, Q1)
         
         If Not Rs.EOF Then
            EstadoUltDTE = vFld(Rs(0))
         End If
         
         Call CloseRs(Rs)
         
         If EstadoUltDTE > 0 And EstadoUltDTE <> EDTE_EMITIDO Then
            If MsgBox1("El último DTE de este tipo enviado al SII, aún no se encuentra en estado Emitido." & vbCrLf & vbCrLf & "Por lo tanto el SII podría reutilizar el folio." & vbCrLf & vbCrLf & "¿Desea revisar y actualizar el estado del DTE, ingresando a la opción 'DTE Emitidos', antes de continuar?", vbYesNo + vbQuestion) = vbYes Then
               Set FrmAdm = New FrmAdmDTE
               Rc = FrmAdm.FView()
               Set FrmAdm = Nothing
               Exit Sub
            End If
         End If

   
         Rc = vbOK
         If gTipoDocDTE(i).Diminutivo = "NCV" Or gTipoDocDTE(i).Diminutivo = "NDV" Or gTipoDocDTE(i).Diminutivo = "NCE" Or gTipoDocDTE(i).Diminutivo = "NDE" Then
            Set FrmTRef = New FrmTipoRef
            Rc = FrmTRef.FSelect(gTipoDocDTE(i).Diminutivo, OptSel, EsNotaCredDebFactCompra)
            Set Frm = Nothing
         
            If Rc = vbOK Then
               MsgBox1 "Seleccione el documento de referencia.", vbInformation
               Set FrmAdm = New FrmAdmDTE
               Rc = FrmAdm.FSelect(gTipoDocDTE(i).Diminutivo, OptSel, IdDTERef, DiminutivoDocRef, EsNotaCredDebFactCompra)
               Set FrmAdm = Nothing
               
               If Rc = vbCancel Then
               
                  'vemos si desea crear una NC o ND desde cero, sin doc de referencia
                  If gTipoDocDTE(i).Diminutivo = "NDV" Or gTipoDocDTE(i).Diminutivo = "NDE" Then  'nota de débito de venta o exportación
                     TipoNota = "Nota de Débito"
                  Else
                     TipoNota = "Nota de Crédito"
                  End If
                  
                  If MsgBox1("¿Desea crear una " & TipoNota & " desde cero, sin utilizar alguno de los documentos almacenados en el sistema?", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
                     Rc = vbOK
                  Else
                     Exit Sub
                  End If
                  
               End If
               
            Else
               Exit Sub
               
            End If
            
         End If
         
         If Rc = vbOK Then
            Set Frm = New FrmDTE
            Call Frm.FNew(gTipoDocDTE(i).IdxTipoDoc, OptSel, IdDTERef, gTipoDocDTE(i).CodDocDTESII = CODDOCDTESII_GUIADESPACHO, DiminutivoDocRef, EsNotaCredDebFactCompra)
            Set Frm = Nothing
            
         Else

         End If
      End If
   Next i
End Sub

Private Sub Form_Load()
   Dim i As Integer
   Dim DocsNoHabilitados As Boolean
      
   For i = 0 To UBound(gTipoDocDTE)
   
      If gTipoDocDTE(i).CodDocDTESII = "" Then
         Exit For
      End If
      
      Op_TipoDoc(i).Caption = gTipoDocDTE(i).Nombre & " Electrónica"
      
'      If gTipoDocDTE(i).Diminutivo <> "GDE" And (gTipoDocDTE(i).Diminutivo = "LFV" Or DocsNoHabilitados) Then  'por ahora no se habilitan los docs desde Liquidación Factura en adelante
'      If gTipoDocDTE(i).Diminutivo = "LFV" Or gTipoDocDTE(i).Diminutivo = "NCE" Or gTipoDocDTE(i).Diminutivo = "NDE" Then   'por ahora no se habilitan
      If gTipoDocDTE(i).Diminutivo = "LFV" Then    'por ahora no se habilitan
         DocsNoHabilitados = True
         Op_TipoDoc(i).Enabled = False
      End If
   Next i
      
   Op_TipoDoc(1).Value = 1   'seleccionamos el primer documento (Factura de Venta)
      
End Sub

Private Sub Op_TipoDoc_DblClick(Index As Integer)
   Call bt_OK_Click
End Sub
