VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmImpDTERecibidos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar DTEs Recibidos"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9690
   Icon            =   "FrmImpDTERecibidos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   300
      TabIndex        =   2
      Top             =   360
      Width           =   8895
      Begin MSComctlLib.ProgressBar PgrBar 
         Height          =   230
         Left            =   240
         TabIndex        =   3
         Top             =   420
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Archivo:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   780
         Width           =   585
      End
      Begin VB.Label Lb_Archivo 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1260
         TabIndex        =   4
         Top             =   780
         Width           =   45
      End
   End
   Begin VB.CommandButton Bt_Importar 
      Caption         =   "Importar"
      Height          =   375
      Left            =   6240
      TabIndex        =   1
      Top             =   1980
      Width           =   1395
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   7800
      TabIndex        =   0
      Top             =   1980
      Width           =   1395
   End
End
Attribute VB_Name = "FrmImpDTERecibidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lPathLogImp As String

Public Function Import_DTERecibidosAcepta() As Boolean
   Dim Files() As String
   Dim i As Integer
   Dim nFiles As Integer, Rc As Integer

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
      
   For i = 0 To nFiles - 1
      If Files(i) = "" Then
         Exit For
      End If
      
      Lb_Archivo = Files(i)
            
      Call Import_DTERecFile(Files(i))
      
      PgrBar.Value = i + 1
   Next i
   
   MsgBox1 "Proceso de importación finalizado.", vbInformation
   
   Call SetIniString(gCfgFile, "Import-" & gEmpresa.Rut, "FDteRec", CLng(Int(Now))) ' 2 ago 2019: para saber cuando leyo por última vez
   
   
End Function

Public Function Import_DTERecFile(ByVal fname As String) As Boolean
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
   Rc = AcpArchFCompra(Val(gEmpresa.Rut), fname, FileBuf)
   
   If Rc <> 0 Or FileBuf = "" Then
      MsgBox1 "Error al leer el archico: " & fname, vbExclamation + vbOKOnly
      Exit Function
   End If
   
   Debug.Print "Importando " & fname
   
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
         Call AddLogImp(FNameLogImp, fname, l, "Falta el tipo de documento.")
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
      If NumDoc = "" Then
         Call AddLogImp(FNameLogImp, fname, l, "Falta el número de documento.")
         DocErr = True
      End If
      
      'RUT Emisor - Proveedor
      RutEmisor = Trim(NextField2(Buf, p, Sep))
      If Not ValidRut(RutEmisor) Then
         Call AddLogImp(FNameLogImp, fname, l, "RUT inválido.")
         DocErr = True
      End If
           
      
      'Razón Social
      RazonSocial = Utf8Ansi(Trim(NextField2(Buf, p, Sep))) ' 30 jul 2019 - pam: se agrega Utf8Ansi( )
      If RazonSocial = "" Then
         Call AddLogImp(FNameLogImp, fname, l, "Falta razón social.")
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
         Call AddLogImp(FNameLogImp, fname, l, "RUT inválido.")
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
         Q1 = "SELECT IdDTE, IdEntidad FROM DTERecibidos WHERE IdEmpresa = " & gEmpresa.id
         Q1 = Q1 & " AND RUTEmisor = '" & RutEmisor & "'"
         Q1 = Q1 & " AND CodDocSII = '" & CodDocDTESII & "'"
         Q1 = Q1 & " AND Folio = " & Val(NumDoc)
         
         Set Rs = OpenRs(DbMain, Q1)
         If Not Rs.EOF Then ' ya existe el doc
         
            ' Se revisa si cambio idIdentidad es diferente, de lo contrario no se ve en el reporte
            If vFld(Rs("idEntidad")) <> IdEntidad Then
               Q1 = "UPDATE DTERecibidos SET IdEntidad=" & IdEntidad & " WHERE IdEmpresa = " & gEmpresa.id & " And IdDTE=" & vFld(Rs("IdDTE"))
            Else
               Q1 = ""
            End If
               
            Call CloseRs(Rs)
                
            If Q1 <> "" Then
               Call ExecSQL(DbMain, Q1)
            End If
         
'            Call AddLog("ImportDTE: el documento '" & FmtCID(RutEmisor) & "|" & CodDocDTESII & "|" & Val(NumDoc) & "' ya existe en la base, no se importa.")
            Call AddLogImp(FNameLogImp, fname, l, "El documento '" & FmtCID(RutEmisor) & "|" & CodDocDTESII & "|" & Val(NumDoc) & "' ya existe en la base, no se importa.")
            GoTo NextRec
         End If
         Call CloseRs(Rs)
         
         'no existe, lo agregamos
      
         IdDTERec = TbAddNew(DbMain, "DTERecibidos", "IdDTE", "IdEmpresa")
         Q1 = "UPDATE DTERecibidos SET "
         Q1 = Q1 & "  IdEmpresa = " & gEmpresa.id
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
      
   Call AddLogImp(FNameLogImp, fname, l, MsgDocsOK)
   
   Import_DTERecFile = False   'error
   

End Function

Private Sub Bt_Cancelar_Click()
   Unload Me
End Sub

Private Sub Bt_Importar_Click()

   Call Import_DTERecibidosAcepta
   Unload Me

End Sub

Private Sub Form_Load()
   Debug.Print "A importar"
End Sub
