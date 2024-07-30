VERSION 5.00
Begin VB.Form FrmDatosFactExp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Otros Antecedentes Factura de Exportación"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10995
   Icon            =   "FrmDatosFactExp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_EdEquivalencias 
      Caption         =   "Equi&valencias"
      Height          =   795
      Left            =   9360
      Picture         =   "FrmDatosFactExp.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   5460
      Width           =   1275
   End
   Begin VB.CommandButton bt_Cancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      CausesValidation=   0   'False
      Height          =   315
      Left            =   9360
      TabIndex        =   12
      Top             =   1080
      Width           =   1275
   End
   Begin VB.CommandButton bt_OK 
      Caption         =   "Aceptar"
      Height          =   315
      Left            =   9360
      TabIndex        =   11
      Top             =   720
      Width           =   1275
   End
   Begin VB.Frame Frame3 
      Caption         =   "Contenido del Documento"
      Height          =   1095
      Left            =   1140
      TabIndex        =   25
      Top             =   6780
      Width           =   7755
      Begin VB.TextBox Tx_TotalBultos 
         Height          =   315
         Left            =   2640
         MaxLength       =   14
         TabIndex        =   10
         Top             =   420
         Width           =   1515
      End
      Begin VB.Label Label1 
         Caption         =   "Total Bultos:"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   26
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Antecedentes de Documento"
      Height          =   1455
      Left            =   1200
      TabIndex        =   22
      Top             =   5100
      Width           =   7695
      Begin VB.CommandButton Bt_ConvMoneda 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6360
         Picture         =   "FrmDatosFactExp.frx":064E
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Convertir moneda"
         Top             =   780
         Width           =   375
      End
      Begin VB.CommandButton Bt_Equivalencia 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6780
         Picture         =   "FrmDatosFactExp.frx":0AD6
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Equivalencias"
         Top             =   780
         Width           =   375
      End
      Begin VB.TextBox Tx_TipoCambio 
         Height          =   315
         Left            =   2640
         MaxLength       =   8
         TabIndex        =   9
         Top             =   840
         Width           =   1515
      End
      Begin VB.ComboBox Cb_TipoMoneda 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   420
         Width           =   4515
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de cambio:"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   24
         Top             =   900
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo de moneda:"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   23
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   300
      Picture         =   "FrmDatosFactExp.frx":0F26
      ScaleHeight     =   615
      ScaleWidth      =   555
      TabIndex        =   17
      Top             =   600
      Width           =   555
   End
   Begin VB.Frame Frame1 
      Caption         =   "Despacho de Exportación"
      Height          =   4275
      Left            =   1200
      TabIndex        =   16
      Top             =   600
      Width           =   7695
      Begin VB.ComboBox Cb_ModVenta 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2340
         Width           =   4515
      End
      Begin VB.ComboBox Cb_ClausulaVenta 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2760
         Width           =   4515
      End
      Begin VB.TextBox Tx_TotClausulaVenta 
         Height          =   315
         Left            =   2640
         MaxLength       =   15
         TabIndex        =   6
         Top             =   3180
         Width           =   1515
      End
      Begin VB.ComboBox Cb_ViaTransporte 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   3600
         Width           =   4515
      End
      Begin VB.ComboBox Cb_PuertoDesembarque 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1800
         Width           =   4515
      End
      Begin VB.ComboBox Cb_PuertoEmbarque 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1380
         Width           =   4515
      End
      Begin VB.ComboBox Cb_Pais 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   960
         Width           =   4515
      End
      Begin VB.ComboBox Cb_IndServicio 
         Height          =   315
         Left            =   2640
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   420
         Width           =   4515
      End
      Begin VB.Label Label1 
         Caption         =   "Modalidad de Venta:"
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   30
         Top             =   2400
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Cláusula de Venta:"
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   29
         Top             =   2820
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Total Cláusula de Venta:"
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   28
         Top             =   3240
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Vía de Transporte:"
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   27
         Top             =   3660
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Puerto de Desembarque:"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   21
         Top             =   1860
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Puerto de Embarque:"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   20
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "País Destino:"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   19
         Top             =   1020
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "Indicador Servicio:"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   18
         Top             =   480
         Width           =   2175
      End
   End
End
Attribute VB_Name = "FrmDatosFactExp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lDTEFactExp As DTEFactExp_t
Dim lRc As Integer

Friend Function FEdit(DTEFactExp As DTEFactExp_t)

   lDTEFactExp = DTEFactExp
   
   Me.Show vbModal
   
   DTEFactExp = lDTEFactExp
   
End Function

Private Sub Bt_Cancelar_Click()
   lRc = vbCancel
   Unload Me
End Sub

Private Sub Bt_EdEquivalencias_Click()
   Dim Frm As FrmEquivalencias
   
   Set Frm = New FrmEquivalencias
   Call Frm.FEdit(0)
   Set Frm = Nothing

End Sub

Private Sub bt_OK_Click()

   If Not Valida() Then
      Exit Sub
   End If
   
   Call SaveData
   
   lRc = vbOK
   Unload Me
End Sub
Private Function LoadData()

   Call CbSelItem(Cb_IndServicio, Val(lDTEFactExp.CodIndServicio))
   Call CbSelItem(Cb_Pais, Val(lDTEFactExp.CodPais))
   Call CbSelItem(Cb_PuertoEmbarque, Val(lDTEFactExp.CodPuertoEmbarque))
   Call CbSelItem(Cb_PuertoDesembarque, Val(lDTEFactExp.CodPuertoDesembarque))
   Call CbSelItem(Cb_TipoMoneda, Val(lDTEFactExp.CodMoneda))
   Call CbSelItem(Cb_ModVenta, Val(lDTEFactExp.CodModVenta))
   Call CbSelItem(Cb_ClausulaVenta, Val(lDTEFactExp.CodClausulaVenta))
   Call CbSelItem(Cb_ViaTransporte, Val(lDTEFactExp.CodViaTransporte))
   Tx_TipoCambio = IIf(lDTEFactExp.TipoCambio <> 0, Format(lDTEFactExp.TipoCambio, DBLFMT2), "")
   Tx_TotalBultos = IIf(lDTEFactExp.TotalBultos <> 0, Format(lDTEFactExp.TotalBultos, NUMFMT), "")
   Tx_TotClausulaVenta = IIf(lDTEFactExp.TotClausulaVenta <> 0, Format(lDTEFactExp.TotClausulaVenta, DBLFMT2), "")
   
   

End Function
Private Function SaveData()

   lDTEFactExp.CodIndServicio = CbItemData(Cb_IndServicio)
   lDTEFactExp.CodPais = CbItemData(Cb_Pais)
   lDTEFactExp.CodPuertoEmbarque = CbItemData(Cb_PuertoEmbarque)
   lDTEFactExp.CodPuertoDesembarque = CbItemData(Cb_PuertoDesembarque)
   lDTEFactExp.CodMoneda = Right("00" & CbItemData(Cb_TipoMoneda), 3)
   lDTEFactExp.CodModVenta = CbItemData(Cb_ModVenta)
   lDTEFactExp.CodClausulaVenta = CbItemData(Cb_ClausulaVenta)
   lDTEFactExp.CodViaTransporte = Right("0" & CbItemData(Cb_ViaTransporte), 2)
   lDTEFactExp.TipoCambio = vFmt(Tx_TipoCambio)
   lDTEFactExp.TotalBultos = vFmt(Tx_TotalBultos)
   lDTEFactExp.TotClausulaVenta = vFmt(Tx_TotClausulaVenta)
   

  
End Function
Private Function FillCb()
   Dim i As Integer
   Dim Q1 As String

   'Indicador de Servicio
   Call CbAddItem(Cb_IndServicio, "", 0)
   For i = 1 To UBound(gIndServicio)
      Call CbAddItem(Cb_IndServicio, gIndServicio(i).Nombre, gIndServicio(i).Codigo)
   Next i

   'Modalidad de Venta
   Call CbAddItem(Cb_ModVenta, "", 0)
   For i = 1 To UBound(gModVenta)
      Call CbAddItem(Cb_ModVenta, gModVenta(i).Nombre, gModVenta(i).Codigo)
   Next i

   'Vía de Transporte
   Call CbAddItem(Cb_ViaTransporte, "", 0)
   For i = 1 To UBound(gViaTransporte)
      Call CbAddItem(Cb_ViaTransporte, gViaTransporte(i).Nombre, gViaTransporte(i).Codigo)
   Next i

   'Pais
   Call CbAddItem(Cb_Pais, "", 0)
   Q1 = "SELECT Nombre, Codigo FROM Paises ORDER BY Nombre"
   Call FillCombo(Cb_Pais, DbMain, Q1, 0, True)
   
   'Puerto Embarque
   Call CbAddItem(Cb_PuertoEmbarque, "", 0)
   Q1 = "SELECT Nombre, Codigo FROM Puertos ORDER BY Nombre"
   Call FillCombo(Cb_PuertoEmbarque, DbMain, Q1, 0, True)
   
   'Puerto DesEmbarque
   Call CbAddItem(Cb_PuertoDesembarque, "", 0)
   Q1 = "SELECT Nombre, Codigo FROM Puertos ORDER BY Nombre"
   Call FillCombo(Cb_PuertoDesembarque, DbMain, Q1, 0, True)
   
   'Moneda
   Call CbAddItem(Cb_TipoMoneda, "", 0)
   Q1 = "SELECT Descrip, CodAduana FROM Monedas WHERE NOT CodAduana IS NULL AND CodAduana <> ' ' ORDER BY Descrip"
   Call FillCombo(Cb_TipoMoneda, DbMain, Q1, 0, True)
   
   'Clausula de Venta
   Call CbAddItem(Cb_ClausulaVenta, "", 0)
   Q1 = "SELECT Sigla + ' - ' + Nombre As Descrip, Codigo FROM ClauCompraVenta ORDER BY Sigla"
   Call FillCombo(Cb_ClausulaVenta, DbMain, Q1, 0, False)
   
End Function

Function Valida() As Boolean
   Dim ValidaDatosExp As Boolean
   Dim CodServ As Long

   Valida = False
   
   ValidaDatosExp = True
   
   
   ValidaDatosExp = ValidaDatosExp And CbItemData(Cb_IndServicio) <> 0
   If Not ValidaDatosExp Then
      MsgBox1 "Falta definir el Indicador de Servicio.", vbExclamation
      Exit Function
   End If
   
   ValidaDatosExp = ValidaDatosExp And CbItemData(Cb_Pais) <> 0
   If Not ValidaDatosExp Then
      MsgBox1 "Falta definir el País de Destino.", vbExclamation
      Exit Function
   End If
   
   CodServ = CbItemData(Cb_IndServicio)
   
   If CodServ = Val(INDSERV_MERCADERIAS) Or CodServ = Val(INDSERV_TRANSPTERRESTRE) Then
      ValidaDatosExp = ValidaDatosExp And CbItemData(Cb_PuertoEmbarque) <> 0
      If Not ValidaDatosExp Then
         MsgBox1 "Sí Indicador de Servicio es Mercaderías o Servicio de Transporte Terrestre Internacional, debe indicar el Puerto de Embarque.", vbExclamation
         Exit Function
      End If
      
      ValidaDatosExp = ValidaDatosExp And CbItemData(Cb_PuertoDesembarque) <> 0
      If Not ValidaDatosExp Then
         MsgBox1 "Sí Indicador de Servicio es Mercaderías o Servicio de Transporte Terrestre Internacional, debe indicar el Puerto de Desembarque.", vbExclamation
         Exit Function
      End If
      
      ValidaDatosExp = ValidaDatosExp And CbItemData(Cb_ModVenta) <> 0
      If Not ValidaDatosExp Then
         MsgBox1 "Sí Indicador de Servicio es Mercaderías o Servicio de Transporte Terrestre Internacional, debe indicar la Modalidad de Venta.", vbExclamation
         Exit Function
      End If
      
      ValidaDatosExp = ValidaDatosExp And CbItemData(Cb_ClausulaVenta) <> 0
      If Not ValidaDatosExp Then
         MsgBox1 "Sí Indicador de Servicio es Mercaderías o Servicio de Transporte Terrestre Internacional, debe indicar la Cláusula de Venta.", vbExclamation
         Exit Function
      End If
      
      ValidaDatosExp = ValidaDatosExp And vFmt(Tx_TotClausulaVenta) <> 0
      If Not ValidaDatosExp Then
         MsgBox1 "Sí Indicador de Servicio es Mercaderías o Servicio de Transporte Terrestre Internacional, debe indicar el Total de la Cláusula de Venta.", vbExclamation
         Exit Function
      End If
      
      ValidaDatosExp = ValidaDatosExp And CbItemData(Cb_ViaTransporte) <> 0
      If Not ValidaDatosExp Then
         MsgBox1 "Sí Indicador de Servicio es Mercaderías o Servicio de Transporte Terrestre Internacional, debe indicar la Vía de Transporte.", vbExclamation
         Exit Function
      End If
            
      ValidaDatosExp = ValidaDatosExp And vFmt(Tx_TotalBultos) <> 0
      If Not ValidaDatosExp Then
         MsgBox1 "Sí Indicador de Servicio es Mercaderías o Servicio de Transporte Terrestre Internacional, debe indicar el Total de Bultos.", vbExclamation
         Exit Function
      End If
            
   End If

   ValidaDatosExp = ValidaDatosExp And CbItemData(Cb_TipoMoneda) <> 0
   If Not ValidaDatosExp Then
      MsgBox1 "Falta definir el Tipo de Moneda.", vbExclamation
      Exit Function
   End If
   
   ValidaDatosExp = ValidaDatosExp And vFmt(Tx_TipoCambio) <> 0
   If Not ValidaDatosExp Then
      MsgBox1 "Falta definir el Tipo de Cambio a Pesos.", vbExclamation
      Exit Function
   End If
   
   Valida = True
   
End Function



Private Sub Cb_IndServicio_Click()
   Dim CodServ As Long

   CodServ = CbItemData(Cb_IndServicio)
   
   If CodServ = 0 Or CodServ = Val(INDSERV_MERCADERIAS) Or CodServ = Val(INDSERV_TRANSPTERRESTRE) Then
      Cb_PuertoEmbarque.Enabled = True
      Cb_PuertoDesembarque.Enabled = True
      Cb_ModVenta.Enabled = True
      Cb_ClausulaVenta.Enabled = True
      Call SetRO(Tx_TotClausulaVenta, False)
      Cb_ViaTransporte.Enabled = True
      Call SetRO(Tx_TotalBultos, False)
    
   Else
      Cb_PuertoEmbarque.Enabled = False
      Cb_PuertoEmbarque.ListIndex = 0
      Cb_PuertoDesembarque.Enabled = False
      Cb_PuertoDesembarque.ListIndex = 0
      Cb_ModVenta.Enabled = False
      Cb_ModVenta.ListIndex = 0
      Cb_ClausulaVenta.Enabled = False
      Cb_ClausulaVenta.ListIndex = 0
      Call SetRO(Tx_TotClausulaVenta, True)
      Tx_TotClausulaVenta = ""
      Cb_ViaTransporte.Enabled = False
      Cb_ViaTransporte.ListIndex = 0
      Call SetRO(Tx_TotalBultos, True)
      Tx_TotalBultos = ""
   End If
   
End Sub

Private Sub Cb_TipoMoneda_Click()

   If LCase(Cb_TipoMoneda) = "pesos" Then
      Tx_TipoCambio = Format(1, DBLFMT2)
      Call SetRO(Tx_TipoCambio, True)
   Else
      Tx_TipoCambio = ""
      Call SetRO(Tx_TipoCambio, False)
   End If
   
End Sub

Private Sub Form_Load()
   Call FillCb
   Call LoadData
End Sub

Private Sub Tx_NBultos_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)
End Sub

Private Sub Tx_NBultos_LostFocus()
   Tx_TotalBultos = IIf(vFmt(Tx_TotalBultos), Format(vFmt(Tx_TotalBultos), NUMFMT), "")
   
End Sub

Private Sub Tx_TipoCambio_KeyPress(KeyAscii As Integer)
   Call KeyDecPos(KeyAscii)

End Sub

Private Sub Tx_TipoCambio_LostFocus()
   Tx_TipoCambio = IIf(vFmt(Tx_TipoCambio) <> 0, Format(vFmt(Tx_TipoCambio), DBLFMT2), "")

End Sub

Private Sub Tx_TotClauVenta_KeyPress(KeyAscii As Integer)
   Call KeyDecPos(KeyAscii)

End Sub

Private Sub Tx_TotClauVenta_LostFocus()
   Tx_TotClausulaVenta = IIf(vFmt(Tx_TotClausulaVenta) <> 0, Format(vFmt(Tx_TotClausulaVenta), DBLFMT2), "")

End Sub
Private Sub Bt_ConvMoneda_Click()
   Dim Frm As FrmConverMoneda
   Dim Valor As Double
      
   Set Frm = New FrmConverMoneda
   Call Frm.FView(Valor)
      
   Set Frm = Nothing

End Sub
Private Sub bt_Equivalencia_Click()
   Dim Frm As FrmEquivalencias
   Dim Valor As Double
   
   Set Frm = New FrmEquivalencias
   Call Frm.FSelect(0, Valor, True)
   
   If Valor > 0 Then
      Tx_TipoCambio = Format(Valor, DBLFMT2)
   Else
      Tx_TipoCambio = 0
   End If
   
   Set Frm = Nothing

End Sub


Private Sub Tx_TotClausulaVenta_KeyPress(KeyAscii As Integer)
   Call KeyNumPos(KeyAscii)

End Sub

Private Sub Tx_TotClausulaVenta_LostFocus()
   Tx_TotClausulaVenta = Format(vFmt(Tx_TotClausulaVenta), NUMFMT)
End Sub
