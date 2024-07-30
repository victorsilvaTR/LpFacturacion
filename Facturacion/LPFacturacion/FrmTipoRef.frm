VERSION 5.00
Begin VB.Form FrmTipoRef 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generar Nota de Crédito"
   ClientHeight    =   4140
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8670
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   8670
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Fr_EsNotaCredDebFactCompra 
      Caption         =   "Tipo de Documento"
      Height          =   735
      Left            =   1380
      TabIndex        =   7
      Top             =   2820
      Width           =   5115
      Begin VB.CheckBox Ch_EsNotaCredDebFactCompra 
         Caption         =   "Nota de Crédito de Factura de Compra"
         Height          =   255
         Left            =   300
         TabIndex        =   8
         Top             =   300
         Width           =   3795
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   420
      Picture         =   "FrmTipoRef.frx":0000
      ScaleHeight     =   615
      ScaleWidth      =   555
      TabIndex        =   6
      Top             =   600
      Width           =   555
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6900
      TabIndex        =   5
      Top             =   1080
      Width           =   1275
   End
   Begin VB.CommandButton Bt_Aceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   6900
      TabIndex        =   4
      Top             =   600
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   1380
      TabIndex        =   0
      Top             =   540
      Width           =   5115
      Begin VB.OptionButton Op_Ref 
         Caption         =   "Generar Nota de Crédito para Corregir Textos"
         Height          =   375
         Index           =   2
         Left            =   300
         TabIndex        =   3
         Top             =   780
         Width           =   4035
      End
      Begin VB.OptionButton Op_Ref 
         Caption         =   "Generar Nota de Crédito para Corregir Montos"
         Height          =   375
         Index           =   3
         Left            =   300
         TabIndex        =   2
         Top             =   1140
         Width           =   4035
      End
      Begin VB.OptionButton Op_Ref 
         Caption         =   "Generar Nota de Crédito de Anulación"
         Height          =   375
         Index           =   1
         Left            =   300
         TabIndex        =   1
         Top             =   420
         Width           =   4035
      End
   End
End
Attribute VB_Name = "FrmTipoRef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const SM_HEIGT = 3450

Dim lRc As Integer
Dim lOptSel As Integer
Dim lDimTipoDoc As String
Dim lEsNotaCredDebFactCompra As Boolean

Public Function FSelect(ByVal DimTipoDoc As String, OptSel As Integer, EsNotaCredDebFactCompra As Boolean) As Integer
   lDimTipoDoc = DimTipoDoc
   
   Me.Show vbModal
   FSelect = lRc
   OptSel = lOptSel
   EsNotaCredDebFactCompra = lEsNotaCredDebFactCompra
   
End Function
Private Sub Bt_Aceptar_Click()
   
   lRc = vbOK

   Unload Me
End Sub

Private Sub Bt_Cancelar_Click()
   
   lOptSel = 0
   lRc = vbCancel
  
   Unload Me
End Sub

Private Sub Ch_EsNotaCredDebFactCompra_Click()

   lEsNotaCredDebFactCompra = IIf(Ch_EsNotaCredDebFactCompra <> 0, True, False)
   
End Sub

Private Sub Form_Load()

   If lDimTipoDoc = "NDV" Or lDimTipoDoc = "NDE" Then   'nota de débito de venta o exportación
      Ch_EsNotaCredDebFactCompra.Caption = "Nota de Débito de Factura de Compra"
      
      Me.Caption = "Generar Nota de Débito"
      Op_Ref(REF_ANULA).Caption = "Generar Nota de Débito que elimina una Nota de Crédito en la referencia en forma completa"
      Op_Ref(REF_CORRIGETEXTO).Visible = False
      Op_Ref(REF_CORRIGEMONTOS).Caption = "Generar Nota de Débito para Corregir Montos"
      
   End If
   
   If lDimTipoDoc = "NCE" Or lDimTipoDoc = "NDE" Then
      Fr_EsNotaCredDebFactCompra.Visible = False
      Me.Height = SM_HEIGT
   End If
   
End Sub

Private Sub Op_Ref_Click(Index As Integer)
   lOptSel = Index
End Sub

Private Sub Op_Ref_DblClick(Index As Integer)
   lOptSel = Index
   Call Bt_Aceptar_Click
End Sub
