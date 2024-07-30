VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmImportDocs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importacion de Documetos Contables"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   7215
      Left            =   -360
      TabIndex        =   0
      Top             =   -240
      Width           =   10335
      Begin VB.Frame Frame2 
         Caption         =   "Seleccione Filtros a Importar"
         Height          =   3735
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   6255
         Begin VB.ComboBox Cb_Desde 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   480
            Width           =   1455
         End
         Begin VB.ComboBox Cb_Hasta 
            Height          =   315
            Left            =   3600
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   480
            Width           =   1335
         End
         Begin VB.CommandButton bt_Import 
            Caption         =   "Importar"
            Height          =   315
            Left            =   1680
            TabIndex        =   7
            Top             =   1560
            Width           =   2355
         End
         Begin VB.CommandButton bt_Cancel 
            Cancel          =   -1  'True
            Caption         =   "Cancelar"
            Height          =   315
            Left            =   1680
            TabIndex        =   6
            Top             =   2160
            Width           =   2355
         End
         Begin VB.ComboBox Cb_Tipo 
            Height          =   315
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   960
            Width           =   3855
         End
         Begin MSComctlLib.ProgressBar PgrBar 
            Height          =   225
            Left            =   120
            TabIndex        =   8
            Top             =   3000
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   397
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Archivo:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   9
            Top             =   3360
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   1080
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha desde:"
            Height          =   195
            Index           =   4
            Left            =   120
            TabIndex        =   3
            Top             =   540
            Width           =   975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "hasta:"
            Height          =   195
            Index           =   3
            Left            =   2880
            TabIndex        =   2
            Top             =   540
            Width           =   435
         End
      End
   End
End
Attribute VB_Name = "FrmImportDocs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub bt_Cancel_Click()
Unload Me
End Sub

Private Sub bt_Import_Click()
Dim p As Integer
p = 0
If CbItemData(Cb_Tipo) = 2 Or CbItemData(Cb_Tipo) = 1 Then
    p = 1
    Call Import_DTERecibidosAceptaNew
    Call SetIniString(gCfgFile, "Import-" & gEmpresa.Rut, "FDteRec", CLng(Int(Now))) ' 2 ago 2019: para saber cuando leyo por última vez
End If

If CbItemData(Cb_Tipo) = 3 Or CbItemData(Cb_Tipo) = 1 Then
    p = 1
    Call Import_DTEEmitidosAceptaNew
End If

If p <> 0 Then
    MsgBox1 "Proceso de importación finalizado.", vbInformation
    PgrBar.Value = 0
    Label1(1).Caption = "Proceso Finalizado"
End If

End Sub

Private Function Import_DTERecibidosAceptaNew() As Boolean
   Dim Files() As String
   Dim i As Integer
   Dim nFiles As Integer, Rc As Integer
   Dim desdeIng, HastaIng As Long
   Dim FechaFile As Long
   Dim nombreFile As String

   Rc = AcpListArchFCompra(Val(gEmpresa.Rut), Files, nFiles)
   
   If Rc <> 0 Then
      MsgBox1 "Error al intentar importar archivos.", vbExclamation
      Exit Function
   End If
   
   If nFiles <= 0 Then
      MsgBox1 "No se encontraron archivos para importar.", vbExclamation
      Exit Function
   End If
      
   lPathLogImp = W.AppPath & "\Log"
   If MkDirect(lPathLogImp) Then
      MsgErr lPathLogImp
      Exit Function
   End If
   
   Err.Clear
      
   PgrBar.Max = nFiles
   PgrBar.Value = 0
   For i = 0 To nFiles - 1
      If Files(i) = "" Then
         Exit For
      End If
      
      nombreFile = Files(i)
      Label1(1).Caption = "Archivos Recibidos " & i + 1 & " De " & nFiles & " Nombre : " & nombreFile 'Files(i)
      'desdeIng = IIf(Tx_FechaDesde <> "", GetTxDate(Tx_FechaDesde), 0)
      'HastaIng = IIf(Tx_FechaHasta <> "", GetTxDate(Tx_FechaHasta), 100000)
      desdeIng = VFmtDate(DateSerial(Year(Date), CbItemData(Cb_Desde), 1))
      HastaIng = VFmtDate(DateSerial(IIf(CbItemData(Cb_Hasta) <> 12, Year(Date), Year(Date) + 1), IIf(CbItemData(Cb_Hasta) <> 12, CbItemData(Cb_Hasta) + 1, 1), 1))
      FechaFile = VFmtDate(DateSerial(CInt(Mid(Files(i), 1, 4)), CInt(Mid(Files(i), 5, 2)), 1))
      'Sleep (5000)
      If FechaFile >= desdeIng And FechaFile <= HastaIng Then
        Call Import_DTERecFile(Files(i))
      End If
      
      PgrBar.Value = i + 1
   Next i
   
   'MsgBox1 "Proceso de importación finalizado.", vbInformation
   
   Call SetIniString(gCfgFile, "Import-" & gEmpresa.Rut, "FDteRec", CLng(Int(Now))) ' 2 ago 2019: para saber cuando leyo por última vez
   
   
End Function

Private Function Import_DTEEmitidosAceptaNew() As Boolean
   Dim Files() As String
   Dim i As Integer
   Dim nFiles As Integer, Rc As Integer
   Dim desdeIng, HastaIng As Long
   Dim FechaFile As Long
   Dim nombreFile As String

   Rc = AcpListArchFVenta(Val(gEmpresa.Rut), Files, nFiles)
   
   If Rc <> 0 Then
      MsgBox1 "Error al intentar importar archivos.", vbExclamation
      Exit Function
   End If
   
   If nFiles <= 0 Then
      MsgBox1 "No se encontraron archivos para importar.", vbExclamation
      Exit Function
   End If
      
   lPathLogImp = W.AppPath & "\Log"
   If MkDirect(lPathLogImp) Then
      MsgErr lPathLogImp
      Exit Function
   End If
   
   Err.Clear
      
   PgrBar.Max = nFiles
   PgrBar.Value = 0
      
   For i = 0 To nFiles - 1
      If Files(i) = "" Then
         Exit For
      End If
      
      nombreFile = Files(i)
      Label1(1).Caption = "Archivos Emitidos " & i + 1 & " De " & nFiles & " Nombre : " & nombreFile 'Files(i)
      'Lb_Archivo = Files(i)
'      desdeIng = IIf(Tx_FechaDesde <> "", GetTxDate(Tx_FechaDesde), 0)
'      HastaIng = IIf(Tx_FechaHasta <> "", GetTxDate(Tx_FechaHasta), 100000)
      desdeIng = VFmtDate(DateSerial(Year(Date), CbItemData(Cb_Desde), 1))
      HastaIng = VFmtDate(DateSerial(IIf(CbItemData(Cb_Hasta) <> 12, Year(Date), Year(Date) + 1), IIf(CbItemData(Cb_Hasta) <> 12, CbItemData(Cb_Hasta) + 1, 1), 1))
      FechaFile = VFmtDate(DateSerial(CInt(Mid(Files(i), 1, 4)), CInt(Mid(Files(i), 5, 2)), 1))
      'Sleep (5000)
      If FechaFile >= desdeIng And FechaFile <= HastaIng Then
        Call Import_DTEEmiFile(Files(i))
      End If
      
      PgrBar.Value = i + 1
   Next i
   
   'MsgBox1 "Proceso de importación finalizado.", vbInformation
   
   Call SetIniString(gCfgFile, "Import-" & gEmpresa.Rut, "FDteRec", CLng(Int(Now))) ' 2 ago 2019: para saber cuando leyo por última vez
   
   
End Function

Private Sub Form_Load()

Call AddItem(Me.Cb_Tipo, "Ambos", 1, False)
Call AddItem(Me.Cb_Tipo, "Compras", 2, True)
Call AddItem(Me.Cb_Tipo, "Ventas", 3, False)

Call CbFillMes(Cb_Desde, 1)
Call CbFillMes(Cb_Hasta, 1)

End Sub

Public Function Import_DTERecFile(ByVal FName As String) As Boolean
   Dim Q1 As String, Rs As Recordset
   Dim i As Integer, j As Integer, l As Integer
   Dim p As Long, pr As Long
   Dim FNameLogImp As String
   Dim Fd As Long
   Dim Sep As String
   Dim Buf As String, FileBuf As String
   Dim Txt As String
   Dim NDocsOK As Integer
   Dim RazonSocial As String, NotValidRut As Boolean
   Dim Rc As Integer
   Dim CodDocDTESII As String
   Dim RutEmisor As String, RutReceptor As String
   Dim TipoLib As Integer, TipoDoc As Integer
   Dim NumDoc As String
   Dim IdEntidad As Long
   Dim FechaPublicacion As Long
   Dim FechaEmision As Long
   Dim MontoAfecto As Double
   Dim MontoExento As Double
   Dim MontoIVA As Double
   Dim MontoOtrosImp As Double
   Dim MontoTotal As Double
   Dim Uri As String, CondPago As Integer
   Dim FechaVenc As Long, FechaCesion As Long, FechaRecSII As Long
   Dim IdDTERec As Long
   Dim DocErr As Boolean
   Dim Aux As String, TxtImpuestos As String
   Dim MsgDocsOK As String
   Dim TipoLibDoc As String
   Dim NDocNuevos As Long
   
   TipoLib = LIB_COMPRAS
       
   Import_DTERecFile = False   'error
   On Error Resume Next
      
   FNameLogImp = lPathLogImp & "\ImpDTERecAcepta_" & Format(Now, "yyyymmdd") & ".log"
   
   On Error Resume Next
   
   'leemos el archivo a un buffer
   Rc = AcpArchFCompra(Val(gEmpresa.Rut), FName, FileBuf)
   
   If Rc <> 0 Or FileBuf = "" Then
      MsgBox1 "Error al leer el archico: " & FName, vbExclamation + vbOKOnly
      Exit Function
   End If
   
   Debug.Print "Importando " & FName
   
   Sep = "|"
   NDocsOK = 0
   pr = 1    'posición del registro en el archivo
         
   'Campos: tipo|tipo_documento|folio|emisor|razon_social_emisor|receptor|publicacion|emision|monto_neto|monto_exento|monto_iva|monto_total|impuestos|estado_acepta|estado_sii|estado_intercambio|informacion_intercambio|uri|73|referencias|mensaje_nar|uri_nar|uri_arm|fecha_arm|condicion_pago|fecha_vencimiento|estado_cesion|url_correo_cesion|fecha_cesion|fecha_recepcion_sii|estado_reclamo_mercaderia|fecha_reclamo_mercaderia|estado_reclamo_contenido|fecha_reclamo_contenido|estado_nar|fecha_nar|mensaje_nar
         
   Do
      DocErr = False
      'leemos el registro
      Buf = NextField2(FileBuf, pr, vbCrLf)
      If Buf = "" Then
         Exit Do
      End If
         
      l = l + 1
               
      Buf = Trim(Buf)
      
      If Buf = "" Then
         GoTo NextRec
      ElseIf l = 1 Then ' Primer registro con nombres de campos
         GoTo NextRec
      End If
      
      p = 1
   
      'ahora leemos los documentos y los insertamos uno por uno
      
      'Tipo
      CodDocDTESII = Trim(NextField2(Buf, p, Sep))
      If CodDocDTESII = "" Then
         Call AddLogImp(FNameLogImp, FName, l, "Falta el tipo de documento.")
         DocErr = True
      End If
      
      If Val(CodDocDTESII) = CODDOCDTESII_GUIADESPACHO Then
         TipoLib = LIB_OTROS
         TipoDoc = TIPODOC_GUIADESPACHO
      Else
         TipoLib = LIB_COMPRAS
         TipoDoc = GetTipoDocFromCodDocDTESII(TipoLib, CodDocDTESII)
      End If
      
      'Tipo Doc en letras
      Aux = NextField2(Buf, p, Sep)
      
      'Folio
      NumDoc = Trim(NextField2(Buf, p, Sep))
'      If NumDoc = "388724" Or NumDoc = "47755" Or NumDoc = "178558" Or NumDoc = "23003" Or NumDoc = "7728343" Or NumDoc = "211660" Or NumDoc = "47748" Then
'        NumDoc = NumDoc
'      End If
      If NumDoc = "" Then
         Call AddLogImp(FNameLogImp, FName, l, "Falta el número de documento.")
         DocErr = True
      End If
      
      'RUT Emisor - Proveedor
      RutEmisor = Trim(NextField2(Buf, p, Sep))
      If Not ValidRut(RutEmisor) Then
         Call AddLogImp(FNameLogImp, FName, l, "RUT inválido.")
         DocErr = True
      End If
           
      
      'Razón Social
      RazonSocial = Utf8Ansi(Trim(NextField2(Buf, p, Sep))) ' 30 jul 2019 - pam: se agrega Utf8Ansi( )
      If RazonSocial = "" Then
         Call AddLogImp(FNameLogImp, FName, l, "Falta razón social.")
         DocErr = True
      Else
         'Creamos la entidad si no está
         IdEntidad = GetIdEntidad(RutEmisor, RazonSocial, NotValidRut)
         
         If IdEntidad = 0 Then
            Rc = AddEntidad(RutEmisor, RazonSocial, IdEntidad)
         End If
      End If
                    
      If RazonSocial = "" Then
         MsgBeep vbExclamation
      End If
      
      RutEmisor = vFmtCID(RutEmisor)
      
      'RUT Receptor
      RutReceptor = Trim(NextField2(Buf, p, Sep))
      If Not ValidRut(RutReceptor) Then
         Call AddLogImp(FNameLogImp, FName, l, "RUT inválido.")
         DocErr = True
      End If

      RutReceptor = vFmtCID(RutReceptor)
      
      'Fecha Publicación
      FechaPublicacion = Int(GetDate(Trim(NextField2(Buf, p, Sep)), "ymd"))
      
      'Fecha Docto (Fecha Emision)
      FechaEmision = Int(GetDate(Trim(NextField2(Buf, p, Sep)), "ymd"))
      
      'Monto Neto;Monto Exento;Monto IVA
      MontoAfecto = vFmt(NextField2(Buf, p, Sep))
      MontoExento = vFmt(NextField2(Buf, p, Sep))
      MontoIVA = vFmt(NextField2(Buf, p, Sep))
      
      'Monto Total
      MontoTotal = vFmt(NextField2(Buf, p, Sep))
      MontoOtrosImp = MontoTotal - MontoAfecto - MontoExento - MontoIVA
         
      'TxtImpuestos
      TxtImpuestos = Trim(NextField2(Buf, p, Sep))
      
      'Estado_Acepta
      Aux = Trim(NextField2(Buf, p, Sep))
      
      'Estado_SII
      Aux = Trim(NextField2(Buf, p, Sep))
      'If StrComp(Aux, "ACEPTADO_POR_EL_SII", vbTextCompare) <> 0 Then
      If StrComp(Aux, "ACEPTADO_POR_EL_SII", vbTextCompare) <> 0 And StrComp(Aux, "ACEPTADO", vbTextCompare) <> 0 Then
         Debug.Print "ImportDTE: No aceptado por el SII: NumDoc=" & NumDoc
         DocErr = True
      End If
      
      'estado_intercambio|informacion_intercambio
      Aux = Trim(NextField2(Buf, p, Sep))
      Aux = Trim(NextField2(Buf, p, Sep))
      
      'uri
      Uri = Trim(NextField2(Buf, p, Sep))
      
      'referencias| mensaje_nar| uri_nar|  uri_arm|  fecha_arm
      Aux = Trim(NextField2(Buf, p, Sep))
      Aux = Trim(NextField2(Buf, p, Sep))
      Aux = Trim(NextField2(Buf, p, Sep))
      Aux = Trim(NextField2(Buf, p, Sep))
      Aux = Trim(NextField2(Buf, p, Sep))

      'Condición de pago
      Aux = Utf8Ansi(Trim(NextField2(Buf, p, Sep))) ' 30 jul 2019 - pam: se agrega Utf8Ansi( )
      If StrComp(Aux, "Crédito", vbTextCompare) = 0 Then
         CondPago = FP_CREDITO
      ElseIf StrComp(Aux, "Contado", vbTextCompare) = 0 Then
         CondPago = FP_CONTADO
      Else
         CondPago = FP_SINCOSTO
      End If

      'Fecha vencimiento
      FechaVenc = Int(GetDate(Trim(NextField2(Buf, p, Sep))))

      'estado cesión| url correo cesión
      Aux = Trim(NextField2(Buf, p, Sep))

      'Fecha cesión| Fecha Recep SII
      FechaCesion = Int(GetDate(Trim(NextField2(Buf, p, Sep)), "ymd"))
      FechaRecSII = Int(GetDate(Trim(NextField2(Buf, p, Sep)), "ymd"))

      'estado_reclamo_mercaderia|  fecha_reclamo_mercaderia|   estado_reclamo_contenido|   fecha_reclamo_contenido| estado_nar|  fecha_nar|   mensaje_nar
      Aux = Trim(NextField2(Buf, p, Sep))
      Aux = Trim(NextField2(Buf, p, Sep))
      Aux = Trim(NextField2(Buf, p, Sep))
      Aux = Trim(NextField2(Buf, p, Sep))
      Aux = Trim(NextField2(Buf, p, Sep))
      Aux = Trim(NextField2(Buf, p, Sep))
      Aux = Trim(NextField2(Buf, p, Sep))


      If Not DocErr Then
      
         'vemos si el doc existe
         Q1 = "SELECT IdDTE, IdEntidad FROM DTERecibidos WHERE IdEmpresa = " & gEmpresa.Id
         Q1 = Q1 & " AND RUTEmisor = '" & RutEmisor & "'"
         Q1 = Q1 & " AND CodDocSII = '" & CodDocDTESII & "'"
         Q1 = Q1 & " AND Folio = " & Val(NumDoc)
         
         Set Rs = OpenRs(DbMain, Q1)
         If Not Rs.EOF Then ' ya existe el doc
         
            ' Se revisa si cambio idIdentidad es diferente, de lo contrario no se ve en el reporte
            If vFld(Rs("idEntidad")) <> IdEntidad Then
               Q1 = "UPDATE DTERecibidos SET IdEntidad=" & IdEntidad & " WHERE IdEmpresa = " & gEmpresa.Id & " And IdDTE=" & vFld(Rs("IdDTE"))
            Else
               Q1 = ""
            End If
               
            Call CloseRs(Rs)
                
            If Q1 <> "" Then
               Call ExecSQL(DbMain, Q1)
            End If
         
'            Call AddLog("ImportDTE: el documento '" & FmtCID(RutEmisor) & "|" & CodDocDTESII & "|" & Val(NumDoc) & "' ya existe en la base, no se importa.")
            Call AddLogImp(FNameLogImp, FName, l, "El documento '" & FmtCID(RutEmisor) & "|" & CodDocDTESII & "|" & Val(NumDoc) & "' ya existe en la base, no se importa.")
            GoTo NextRec
         Else
            IdDTERec = TbAddNew(DbMain, "DTERecibidos", "IdDTE", "IdEmpresa")
         End If
         Call CloseRs(Rs)
         
         'no existe, lo agregamos
      
         
         Q1 = "UPDATE DTERecibidos SET "
         Q1 = Q1 & "  IdEmpresa = " & gEmpresa.Id
         Q1 = Q1 & ", TipoDoc = " & TipoDoc
         Q1 = Q1 & ", TipoLib = " & TipoLib
         Q1 = Q1 & ", CodDocSII = '" & CodDocDTESII & "'"
         Q1 = Q1 & ", Folio = " & Val(NumDoc)
         Q1 = Q1 & ", RUTEmisor = '" & RutEmisor & "'"
         Q1 = Q1 & ", RazonSocial = '" & RazonSocial & "'"
         Q1 = Q1 & ", IdEntidad = " & IdEntidad
         Q1 = Q1 & ", RUTReceptor = '" & RutReceptor & "'"
         Q1 = Q1 & ", FPublicacion = " & FechaPublicacion
         Q1 = Q1 & ", FEmision = " & FechaEmision
         Q1 = Q1 & ", Neto = " & MontoAfecto
         Q1 = Q1 & ", Exento = " & MontoExento
         Q1 = Q1 & ", IVA = " & MontoIVA
         Q1 = Q1 & ", Total = " & MontoTotal
         Q1 = Q1 & ", Impuestos = " & MontoOtrosImp
         'Q1 = Q1 & ", TxtDetImpuestos = '" & TxtImpuestos & "'"
         Q1 = Q1 & ", TxtDetImpuestos = '" & Replace(TxtImpuestos, Chr(39), "") & "'"
         Q1 = Q1 & ", UrlDTE = '" & Uri & "'"
         Q1 = Q1 & ", FormaPago = " & CondPago
         Q1 = Q1 & ", FVenc = " & FechaVenc
         Q1 = Q1 & ", Fcesion = " & FechaCesion
         Q1 = Q1 & ", FRecepSII = " & FechaRecSII
         
         Q1 = Q1 & " WHERE IdDTE =" & IdDTERec
         Rc = ExecSQL(DbMain, Q1)
   
         NDocNuevos = NDocNuevos + 1
         
      End If

NextRec:

   Loop
      
   
EndFnc:

   If NDocsOK > 1 Then
      MsgDocsOK = "Se importaron " & NDocsOK & " documentos nuevos."
   ElseIf NDocsOK = 1 Then
      MsgDocsOK = "Se importó un documento nuevo."
   Else
      MsgDocsOK = "No se importaron documentos nuevos."
   End If
      
   Call AddLogImp(FNameLogImp, FName, l, MsgDocsOK)
   
   Import_DTERecFile = False   'error
   

End Function

Public Function Import_DTEEmiFile(ByVal FName As String) As Boolean
   Dim Q1 As String, Rs As Recordset
   Dim i As Integer, j As Integer, l As Integer
   Dim p As Long, pr As Long
   Dim FNameLogImp As String
   Dim Fd As Long
   Dim Sep As String
   Dim Buf As String, FileBuf As String
   Dim Txt As String
   Dim NDocsOK As Integer
   Dim RazonSocial As String, NotValidRut As Boolean
   Dim Rc As Integer
   Dim CodDocDTESII As String
   Dim RutEmisor As String, RutReceptor As String
   Dim TipoLib As Integer, TipoDoc As Integer
   Dim NumDoc As String
   Dim IdEntidad As Long
   Dim FechaPublicacion As Long
   Dim FechaEmision As Long
   Dim MontoAfecto As Double
   Dim MontoExento As Double
   Dim MontoIVA As Double
   Dim MontoOtrosImp As Double
   Dim MontoTotal As Double
   Dim Uri As String, CondPago As Integer
   Dim FechaVenc As Long, FechaCesion As Long, FechaRecSII As Long
   Dim IdDTERec As Long
   Dim DocErr As Boolean
   Dim Aux As String, TxtImpuestos As String
   Dim MsgDocsOK As String
   Dim TipoLibDoc As String
   Dim NDocNuevos As Long
   Dim idEstado As Integer
   Dim EstadoSII As Integer
   
   TipoLib = LIB_VENTAS
       
   Import_DTEEmiFile = False   'error
   On Error Resume Next
      
   FNameLogImp = lPathLogImp & "\ImpDTERecAcepta_" & Format(Now, "yyyymmdd") & ".log"
   
   On Error Resume Next
   
   'leemos el archivo a un buffer
   Rc = AcpArchFVenta(Val(gEmpresa.Rut), FName, FileBuf)
   
   If Rc <> 0 Or FileBuf = "" Then
      MsgBox1 "Error al leer el archico: " & FName, vbExclamation + vbOKOnly
      Exit Function
   End If
   
   Debug.Print "Importando " & FName
   
   Sep = "|"
   NDocsOK = 0
   pr = 1    'posición del registro en el archivo
         
   'Campos: tipo|tipo_documento|folio|emisor|razon_social_emisor|receptor|publicacion|emision|monto_neto|monto_exento|monto_iva|monto_total|impuestos|estado_acepta|estado_sii|estado_intercambio|informacion_intercambio|uri|73|referencias|mensaje_nar|uri_nar|uri_arm|fecha_arm|condicion_pago|fecha_vencimiento|estado_cesion|url_correo_cesion|fecha_cesion|fecha_recepcion_sii|estado_reclamo_mercaderia|fecha_reclamo_mercaderia|estado_reclamo_contenido|fecha_reclamo_contenido|estado_nar|fecha_nar|mensaje_nar
         
   Do
      DocErr = False
      'leemos el registro
      Buf = NextField2(FileBuf, pr, vbCrLf)
      If Buf = "" Then
         Exit Do
      End If
         
      l = l + 1
               
      Buf = Trim(Buf)
      
      If Buf = "" Then
         GoTo NextRec
      ElseIf l = 1 Then ' Primer registro con nombres de campos
         GoTo NextRec
      End If
      
      p = 1
   
      'ahora leemos los documentos y los insertamos uno por uno
      
      'Tipo Doc en letras
      Aux = NextField2(Buf, p, Sep)
      
      'Tipo
      CodDocDTESII = Trim(NextField2(Buf, p, Sep))
      If CodDocDTESII = "" Then
         Call AddLogImp(FNameLogImp, FName, l, "Falta el tipo de documento.")
         DocErr = True
      End If
      
      
      If Val(CodDocDTESII) = CODDOCDTESII_GUIADESPACHO Then
         TipoLib = LIB_OTROS
         TipoDoc = TIPODOC_GUIADESPACHO
      Else
         TipoLib = LIB_VENTAS
         TipoDoc = GetTipoDocFromCodDocDTESII(TipoLib, CodDocDTESII)
      End If
      
      
      
      'Folio
      NumDoc = Trim(NextField2(Buf, p, Sep))
      If NumDoc = "" Then
         Call AddLogImp(FNameLogImp, FName, l, "Falta el número de documento.")
         DocErr = True
      End If
      
      'RUT Emisor - Proveedor
      RutEmisor = Trim(NextField2(Buf, p, Sep))
      If Not ValidRut(RutEmisor) Then
         Call AddLogImp(FNameLogImp, FName, l, "RUT inválido.")
         DocErr = True
      End If
      
      
      RutEmisor = vFmtCID(RutEmisor)
      
      'RUT Receptor
      RutReceptor = Trim(NextField2(Buf, p, Sep))
      If Not ValidRut(RutReceptor) Then
         Call AddLogImp(FNameLogImp, FName, l, "RUT inválido.")
         DocErr = True
      End If

      
      
'      'Razón Social
      RazonSocial = Utf8Ansi(Trim(NextField2(Buf, p, Sep))) ' 30 jul 2019 - pam: se agrega Utf8Ansi( )
      If RazonSocial = "" Then
         Call AddLogImp(FNameLogImp, FName, l, "Falta razón social.")
         DocErr = True
      Else
         'Creamos la entidad si no está
         IdEntidad = GetIdEntidad(RutReceptor, RazonSocial, NotValidRut)

         If IdEntidad = 0 Then
            Rc = AddEntidad(RutReceptor, RazonSocial, IdEntidad)
         End If
      End If
      
      RutReceptor = vFmtCID(RutReceptor)

      If RazonSocial = "" Then
         MsgBeep vbExclamation
      End If
      
      'Fecha Publicación
      FechaPublicacion = Int(GetDate(Trim(NextField2(Buf, p, Sep)), "ymd"))
      
      'Fecha Docto (Fecha Emision)
      FechaEmision = Int(GetDate(Trim(NextField2(Buf, p, Sep)), "ymd"))
      
      'Monto Neto;Monto Exento;Monto IVA
      MontoTotal = vFmt(NextField2(Buf, p, Sep))
      MontoAfecto = vFmt(NextField2(Buf, p, Sep))
      Uri = Trim(NextField2(Buf, p, Sep))
      MontoExento = vFmt(NextField2(Buf, p, Sep))
      MontoIVA = vFmt(NextField2(Buf, p, Sep))
      
      'Monto Total
      'MontoTotal = vFmt(NextField2(Buf, p, Sep))
      'MontoOtrosImp = MontoTotal - MontoAfecto - MontoExento - MontoIVA
         
      'Estado_SII
      Aux = Trim(NextField2(Buf, p, Sep))
      idEstado = EDTE_ENVIADO
      'If StrComp(Aux, "ACEPTADO_POR_EL_SII", vbTextCompare) <> 0 Then
      If StrComp(Aux, "ACEPTADO_POR_EL_SII", vbTextCompare) <> 0 And StrComp(Aux, "ACEPTADO", vbTextCompare) <> 0 Then
         Debug.Print "ImportDTE: No aceptado por el SII: NumDoc=" & NumDoc
         DocErr = True
         EstadoSII = EDTESII_RECHAZADO
      Else
         EstadoSII = EDTESII_ACEPTADO
      End If
      
      
      'TxtImpuestos
      TxtImpuestos = Trim(NextField2(Buf, p, Sep))
      
      'Estado_Acepta
      Aux = Trim(NextField2(Buf, p, Sep))
      

      
      'estado_intercambio|informacion_intercambio
      Aux = Trim(NextField2(Buf, p, Sep))
      Aux = Trim(NextField2(Buf, p, Sep))
      
      'uri
      'Uri = Trim(NextField2(Buf, p, Sep))
      
      'referencias| mensaje_nar| uri_nar|  uri_arm|  fecha_arm
      Aux = Trim(NextField2(Buf, p, Sep))
      Aux = Trim(NextField2(Buf, p, Sep))
      Aux = Trim(NextField2(Buf, p, Sep))
      Aux = Trim(NextField2(Buf, p, Sep))
      Aux = Trim(NextField2(Buf, p, Sep))

      'Condición de pago
      Aux = Utf8Ansi(Trim(NextField2(Buf, p, Sep))) ' 30 jul 2019 - pam: se agrega Utf8Ansi( )
      If StrComp(Aux, "Crédito", vbTextCompare) = 0 Then
         CondPago = FP_CREDITO
      ElseIf StrComp(Aux, "Contado", vbTextCompare) = 0 Then
         CondPago = FP_CONTADO
      Else
         CondPago = FP_SINCOSTO
      End If

      'Fecha vencimiento
      FechaVenc = Int(GetDate(Trim(NextField2(Buf, p, Sep))))

      'estado cesión| url correo cesión
      Aux = Trim(NextField2(Buf, p, Sep))

      'Fecha cesión| Fecha Recep SII
      FechaCesion = Int(GetDate(Trim(NextField2(Buf, p, Sep)), "ymd"))
      FechaRecSII = Int(GetDate(Trim(NextField2(Buf, p, Sep)), "ymd"))

      'estado_reclamo_mercaderia|  fecha_reclamo_mercaderia|   estado_reclamo_contenido|   fecha_reclamo_contenido| estado_nar|  fecha_nar|   mensaje_nar
      Aux = Trim(NextField2(Buf, p, Sep))
      Aux = Trim(NextField2(Buf, p, Sep))
      Aux = Trim(NextField2(Buf, p, Sep))
      Aux = Trim(NextField2(Buf, p, Sep))
      Aux = Trim(NextField2(Buf, p, Sep))
      Aux = Trim(NextField2(Buf, p, Sep))
      Aux = Trim(NextField2(Buf, p, Sep))


      If Not DocErr Then
      
         'vemos si el doc existe
         Q1 = "SELECT IdDTE, IdEntidad FROM DTE WHERE IdEmpresa = " & gEmpresa.Id
         Q1 = Q1 & " AND RUT = '" & RutReceptor & "'"
         Q1 = Q1 & " AND CodDocSII = '" & CodDocDTESII & "'"
         Q1 = Q1 & " AND Folio = " & Val(NumDoc)
         
         Set Rs = OpenRs(DbMain, Q1)
         If Not Rs.EOF Then ' ya existe el doc
         
            ' Se revisa si cambio idIdentidad es diferente, de lo contrario no se ve en el reporte
            If vFld(Rs("idEntidad")) <> IdEntidad Then
               Q1 = "UPDATE DTE SET IdEntidad=" & IdEntidad & " WHERE IdEmpresa = " & gEmpresa.Id & " And IdDTE=" & vFld(Rs("IdDTE"))
            Else
               Q1 = ""
            End If
               
            Call CloseRs(Rs)
                
            If Q1 <> "" Then
               Call ExecSQL(DbMain, Q1)
            End If
         
'            Call AddLog("ImportDTE: el documento '" & FmtCID(RutEmisor) & "|" & CodDocDTESII & "|" & Val(NumDoc) & "' ya existe en la base, no se importa.")
            Call AddLogImp(FNameLogImp, FName, l, "El documento '" & FmtCID(RutEmisor) & "|" & CodDocDTESII & "|" & Val(NumDoc) & "' ya existe en la base, no se importa.")
            GoTo NextRec
         Else
            IdDTERec = TbAddNew(DbMain, "DTE", "IdDTE", "IdEmpresa")
         End If
         Call CloseRs(Rs)
         
         'no existe, lo agregamos
      
         
         Q1 = "UPDATE DTE SET "
         Q1 = Q1 & "  IdEmpresa = " & gEmpresa.Id
         Q1 = Q1 & ", TipoDoc = " & TipoDoc
         Q1 = Q1 & ", TipoLib = " & TipoLib
         Q1 = Q1 & ", CodDocSII = '" & CodDocDTESII & "'"
         Q1 = Q1 & ", Folio = " & Val(NumDoc)
         Q1 = Q1 & ", Fecha = " & FechaEmision
         Q1 = Q1 & ", IdEntidad = " & IdEntidad
         Q1 = Q1 & ", RUT = '" & RutReceptor & "'"
         Q1 = Q1 & ", SubTotal = " & MontoAfecto
         Q1 = Q1 & ", Neto = " & MontoAfecto
         Q1 = Q1 & ", IVA = " & MontoIVA
         Q1 = Q1 & ", Total = " & MontoTotal
         'Q1 = Q1 & ", Total = " & MontoTotal
         Q1 = Q1 & ", IdUsuario =  1 " '& MontoExento
         Q1 = Q1 & ", FechaCreacion = " & FechaEmision
         Q1 = Q1 & ", IdEstado = " & idEstado
         Q1 = Q1 & ", IdEstadoSII = " & EstadoSII
         'Q1 = Q1 & ", Impuestos = " & MontoOtrosImp
         'Q1 = Q1 & ", TxtDetImpuestos = '" & TxtImpuestos & "'"
         'Q1 = Q1 & ", TxtDetImpuestos = '" & Replace(TxtImpuestos, Chr(39), "") & "'"
         Q1 = Q1 & ", UrlDTE = '" & Uri & "'"
         'Q1 = Q1 & ", FormaPago = " & CondPago
         Q1 = Q1 & ", FechaVenc = " & FechaVenc
         'Q1 = Q1 & ", Fcesion = " & FechaCesion
         'Q1 = Q1 & ", FRecepSII = " & FechaRecSII
         
         Q1 = Q1 & " WHERE IdDTE =" & IdDTERec
         Rc = ExecSQL(DbMain, Q1)
   
         NDocNuevos = NDocNuevos + 1
         
      End If

NextRec:

   Loop
      
   
EndFnc:

   If NDocsOK > 1 Then
      MsgDocsOK = "Se importaron " & NDocsOK & " documentos nuevos."
   ElseIf NDocsOK = 1 Then
      MsgDocsOK = "Se importó un documento nuevo."
   Else
      MsgDocsOK = "No se importaron documentos nuevos."
   End If
      
   Call AddLogImp(FNameLogImp, FName, l, MsgDocsOK)
   
   Import_DTEEmiFile = False   'error
   

End Function

