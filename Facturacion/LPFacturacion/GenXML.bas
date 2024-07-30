Attribute VB_Name = "GenXML"
   Option Explicit

Const xVERSION = "1.0"

Public Const FMTDATEDTE = "yyyy-mm-dd"
Public Const FMTNUMDTE = "0"

' ** Funcion que crea el Json de un DTE

Public Function GenJsonDTE(DTE As DTE_t) As String
   Dim xDoc As MSXML2.DOMDocument
   Dim xIDTE As MSXML2.IXMLDOMNode
   Dim xDTE As MSXML2.IXMLDOMNode
   Dim xNode As MSXML2.IXMLDOMNode, xChofer As MSXML2.IXMLDOMNode
   Dim xEncab As MSXML2.IXMLDOMNode
   Dim xDetalle As MSXML2.IXMLDOMNode
   Dim xRef As MSXML2.IXMLDOMNode
   Dim xDescto As MSXML2.IXMLDOMNode
   Dim Folio As String
   Dim IdDoc As String
   Dim objElement As MSXML2.IXMLDOMElement
   Dim i As Integer
   Dim xCodItem As MSXML2.IXMLDOMNode, xImpAdic As MSXML2.IXMLDOMNode
   Dim xSubTotInfo As MSXML2.IXMLDOMNode
   Dim XmlDTE As String
   Dim Idx As Integer
   Dim TotDescIva As Long
   
   TotDescIva = 0
   Folio = ""
   If gConectData.Proveedor = PROV_LP Then
      Folio = "1"
   End If
   
   IdDoc = "F" & Folio & "T" & DTE.CodDocSII
   
   Call AddDebug("GenJsonDTE: idDoc=" & IdDoc)
   
   Set xDoc = New MSXML2.DOMDocument
   Set xNode = xDoc.createProcessingInstruction("xml", "version='1.0'")
   xDoc.appendChild xNode
   Set xNode = Nothing
      
   Set xNode = xAddTagWithAttr(xDoc, xDoc, "DTE", "", "version", "1.0")

   If gConectData.Proveedor = PROV_LP Then
      If DTE.EsExport Then
         Set xDTE = xAddTagWithAttr(xDoc, xNode, "Exportaciones", "", "ID", IdDoc)
      Else
         Set xDTE = xAddTagWithAttr(xDoc, xNode, "Documento", "", "ID", IdDoc)
      End If
   Else   'Acepta
      Set objElement = xNode
      objElement.setAttributeNode xAddAttrib(xDoc, xNode, "xmlns", "http://www.sii.cl/SiiDte")
      objElement.setAttributeNode xAddAttrib(xDoc, xNode, "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
      
      If DTE.EsExport Then
         Set xDTE = xAddTag(xDoc, xNode, "Exportaciones")
      Else
         Set xDTE = xAddTag(xDoc, xNode, "Documento")
      End If
   End If
   
   Set xEncab = xAddTag(xDoc, xDTE, "Encabezado")
   Set xNode = xAddTag(xDoc, xEncab, "IdDoc")
      Call xAddTag(xDoc, xNode, "TipoDTE", DTE.CodDocSII)
      Call xAddTag(xDoc, xNode, "Folio", Folio)
      Call xAddTag(xDoc, xNode, "FchEmis", Format(DTE.Fecha, FMTDATEDTE))
      If DTE.CodDocSII = CODDOCDTESII_GUIADESPACHO Then
         Call xAddTag(xDoc, xNode, "TipoDespacho", DTE.TipoDespacho)
         Call xAddTag(xDoc, xNode, "IndTraslado", DTE.Traslado)
      End If
      If DTE.EsExport Then
         Call xAddTag(xDoc, xNode, "IndServicio", DTE.FactExp.CodIndServicio)
      End If
      If DTE.FormaDePago > 0 Then
         Call xAddTag(xDoc, xNode, "FmaPago", DTE.FormaDePago)
      End If
      If DTE.FechaVenc > 0 Then   'tiene que ir después de tipo despacho. El orden afecta
         Call xAddTag(xDoc, xNode, "FchVenc", Format(DTE.FechaVenc, FMTDATEDTE))
      End If
   
   Set xNode = xAddTag(xDoc, xEncab, "Emisor")
      Call xAddTag(xDoc, xNode, "RUTEmisor", gEmpresa.Rut & "-" & DV_Rut(gEmpresa.Rut))
      Call xAddTag(xDoc, xNode, "RznSoc", Left(Ansi2XmlTxt(gEmpresa.RazonSocial), 100))
      Call xAddTag(xDoc, xNode, "GiroEmis", Left(Ansi2XmlTxt(gEmpresa.Giro), 80))
      Call xAddTag(xDoc, xNode, "Acteco", Left(Ansi2XmlTxt(gEmpresa.CodActEcono), 6))
      Call xAddTag(xDoc, xNode, "DirOrigen", Left(Ansi2XmlTxt(gEmpresa.Direccion), 60))
      Call xAddTag(xDoc, xNode, "CmnaOrigen", Left(Ansi2XmlTxt(gEmpresa.Comuna), 20))
      Call xAddTag(xDoc, xNode, "CiudadOrigen", Left(Ansi2XmlTxt(gEmpresa.Ciudad), 20))
   
   Set xNode = xAddTag(xDoc, xEncab, "Receptor")
      If DTE.NotValidRut Then
         Call xAddTag(xDoc, xNode, "RUTRecep", ENTIMP_RUT)
         Call xAddTag(xDoc, xNode, "RznSocRecep", ENTIMP_RSOCIAL)
      Else
         Call xAddTag(xDoc, xNode, "RUTRecep", DTE.Rut & "-" & DV_Rut(DTE.Rut))
         Call xAddTag(xDoc, xNode, "RznSocRecep", Left(Ansi2XmlTxt(DTE.RazonSocial), 100))
      End If
      Call xAddTag(xDoc, xNode, "GiroRecep", Left(Ansi2XmlTxt(DTE.Giro), 40))
      Call xAddTag(xDoc, xNode, "Contacto", Left(Ansi2XmlTxt(DTE.Contacto), 80))
      Call xAddTag(xDoc, xNode, "DirRecep", Left(Ansi2XmlTxt(DTE.Direccion), 70))
      Call xAddTag(xDoc, xNode, "CmnaRecep", Left(Ansi2XmlTxt(DTE.Comuna), 20))
      Call xAddTag(xDoc, xNode, "CiudadRecep", Left(Ansi2XmlTxt(DTE.Ciudad), 20))


   If DTE.EsExport Or DTE.EsGuiaDesp Then
      Set xNode = xAddTag(xDoc, xEncab, "Transporte")
      
      If DTE.EsGuiaDesp Then
         If DTE.GuiaDesp.Patente <> "" Then
            Call xAddTag(xDoc, xNode, "Patente", Left(Ansi2XmlTxt(Trim(DTE.GuiaDesp.Patente)), 8))
         End If
                  
         If (DTE.GuiaDesp.RutChofer <> "" Or DTE.GuiaDesp.NombreChofer <> "") And (DTE.TipoDespacho = GD_DESPEMICLI Or DTE.TipoDespacho = GD_DESPEMIOTRO) Then
'            Call xAddTag(xDoc, xNode, "RUTTrans", DTE.GuiaDesp.RutChofer & "-" & DV_Rut(DTE.GuiaDesp.RutChofer))
'            Call xAddTag(xDoc, xNode, "RUTChofer", DTE.GuiaDesp.RutChofer & "-" & DV_Rut(DTE.GuiaDesp.RutChofer))   'Si usamos RUTChofer el SII reclama al validar el archivo
'            Call xAddTag(xDoc, xNode, "NombreChofer", Left(Ansi2XmlTxt(DTE.GuiaDesp.NombreChofer), 30))             'se cambia por mensaje de error de validación del SII
            
            Set xChofer = xAddTag(xDoc, xNode, "Chofer") ' 10 may 2019 - pam: se agrega un nivel
            
            If DTE.GuiaDesp.RutChofer <> "" Then   ' primero debe ser el RUT
               Call xAddTag(xDoc, xChofer, "RUTChofer", DTE.GuiaDesp.RutChofer & "-" & DV_Rut(DTE.GuiaDesp.RutChofer))   'Si usamos RUTChofer el SII reclama al validar el archivo
            End If
            
            If DTE.GuiaDesp.NombreChofer <> "" Then
               Call xAddTag(xDoc, xChofer, "NombreChofer", Left(Ansi2XmlTxt(DTE.GuiaDesp.NombreChofer), 30))
            End If
            
         End If
         
      End If
      
      If DTE.EsExport Then
         Call xAddTag(xDoc, xNode, "CiudadDest", Left(Ansi2XmlTxt(DTE.Ciudad), 20))
         Call xAddTag(xDoc, xNode, "CodModVenta", DTE.FactExp.CodModVenta)
         Call xAddTag(xDoc, xNode, "CodClauVenta", DTE.FactExp.CodClausulaVenta)
         Call xAddTag(xDoc, xNode, "TotClauVenta", DTE.FactExp.TotClausulaVenta)
         Call xAddTag(xDoc, xNode, "CodViaTransp", DTE.FactExp.CodViaTransporte)
         Call xAddTag(xDoc, xNode, "CodPtoEmbarque", DTE.FactExp.CodPuertoEmbarque)
         Call xAddTag(xDoc, xNode, "CodPtoDesemb", DTE.FactExp.CodPuertoEmbarque)
         Call xAddTag(xDoc, xNode, "TotBultos", DTE.FactExp.TotalBultos)
      End If
      
   End If
   
   TotDescIva = 0
   Set xNode = xAddTag(xDoc, xEncab, "Totales")
      Call xAddTag(xDoc, xNode, "MntNeto", DTE.Neto)
      Call xAddTag(xDoc, xNode, "MntExe", DTE.Exento)
      If DTE.TasaIVA > 0 Then
         Call xAddTag(xDoc, xNode, "TasaIVA", DTE.TasaIVA * 100)
         Call xAddTag(xDoc, xNode, "IVA", DTE.Iva)
      End If
      
      If DTE.ImpAdic(0).IdImpAdic > 0 And DTE.ImpAdic(0).IdImpAdic <> 8 Then   'hay impuestos adicionales
         For i = 0 To UBound(DTE.ImpAdic)
            If DTE.ImpAdic(i).IdImpAdic > 0 Then
               Set xImpAdic = xAddTag(xDoc, xNode, "ImptoReten")
                  Call xAddTag(xDoc, xImpAdic, "TipoImp", DTE.ImpAdic(i).IdImpAdicSII)
                  Call xAddTag(xDoc, xImpAdic, "TasaImp", ReplaceStr(Format(DTE.ImpAdic(i).TasaImpAdic, DBLFMT2), ",", "."))
                  Call xAddTag(xDoc, xImpAdic, "MontoImp", DTE.ImpAdic(i).MontoImpAdic)
            Else
               Exit For
            End If
         Next i
      ElseIf DTE.ImpAdic(0).IdImpAdic > 0 And DTE.ImpAdic(0).IdImpAdic = 8 Then   'hay impuestos adicionales
         For i = 0 To UBound(DTE.ImpAdic)
            If DTE.ImpAdic(i).IdImpAdic > 0 Then
'               Set xImpAdic = xAddTag(xDoc, xNode, "Rebaja")
'                  Call xAddTag(xDoc, xImpAdic, "TipoReb", DTE.ImpAdic(i).IdImpAdic)
'                  Call xAddTag(xDoc, xImpAdic, "DescReb", DTE.ImpAdic(i).DescImpAdic)
'                  Call xAddTag(xDoc, xImpAdic, "TasaReb", ReplaceStr(Format(DTE.ImpAdic(i).TasaImpAdic, DBLFMT2), ",", "."))
'                  Call xAddTag(xDoc, xImpAdic, "MontoReb", DTE.ImpAdic(i).MontoImpAdic)
                  TotDescIva = DTE.ImpAdic(i).MontoImpAdic
            Else
               Exit For
            End If
         Next i
      End If
      
      Call xAddTag(xDoc, xNode, "CredEC", TotDescIva)
      Call xAddTag(xDoc, xNode, "MntTotal", DTE.Total)
      
   If DTE.EsExport Then
      Set xNode = xAddTag(xDoc, xEncab, "OtraMoneda")
         Call xAddTag(xDoc, xNode, "TpoMoneda", DTE.FactExp.CodMoneda)
         Call xAddTag(xDoc, xNode, "TpoCambio", DTE.FactExp.TipoCambio)
   End If
            
   'Detalle
   For i = 0 To UBound(DTE.DetDTE)
      
      If DTE.DetDTE(i).Producto = "" And i > 0 Then   'tiene que ir al menos una linea de detalle
         Exit For
      End If
      
      Set xDetalle = xAddTag(xDoc, xDTE, "Detalle")
         
         Call xAddTag(xDoc, xDetalle, "NroLinDet", i + 1)
         
         If DTE.DetDTE(i).Cantidad > 0 Then
            Set xCodItem = xAddTag(xDoc, xDetalle, "CdgItem")
               Call xAddTag(xDoc, xCodItem, "TpoCodigo", DTE.DetDTE(i).TipoCod)
               Call xAddTag(xDoc, xCodItem, "VlrCodigo", DTE.DetDTE(i).CodProd)
            If DTE.DetDTE(i).EsExento Then
               Call xAddTag(xDoc, xDetalle, "IndExe", 1)
            End If
            Call xAddTag(xDoc, xDetalle, "NmbItem", Left(Ansi2XmlTxt(DTE.DetDTE(i).Producto), 80))
            Call xAddTag(xDoc, xDetalle, "DscItem", Left(Ansi2XmlTxt(DTE.DetDTE(i).Descrip), 1000))
            
            ' Call xAddTag(xDoc, xDetalle, "QtyItem", Format(DTE.DetDTE(i).Cantidad, FMTNUMDTE))
            ' 13 oct 2017 - pam: permite cantidad con decimales
            Call xAddTag(xDoc, xDetalle, "QtyItem", ReplaceStr(Round(DTE.DetDTE(i).Cantidad, 2), ",", "."))
            Call xAddTag(xDoc, xDetalle, "UnmdItem", Left(Ansi2XmlTxt(DTE.DetDTE(i).UMedida), 4))
                        
            Call xAddTag(xDoc, xDetalle, "PrcItem", Format(DTE.DetDTE(i).Precio, FMTNUMDTE))
            If DTE.DetDTE(i).PjeDescto > 0 Then
               Call xAddTag(xDoc, xDetalle, "DescuentoPct", ReplaceStr(Format(DTE.DetDTE(i).PjeDescto, DBLFMT2), ",", "."))
               Call xAddTag(xDoc, xDetalle, "DescuentoMonto", Format(DTE.DetDTE(i).MontoDescto, FMTNUMDTE))
            End If
            If DTE.DetDTE(i).CodImpAdicSII <> "" And DTE.DetDTE(i).CodImpAdicSII <> "0" And DTE.DetDTE(i).CodImpAdicSII <> "8" Then
               Call xAddTag(xDoc, xDetalle, "CodImpAdic", DTE.DetDTE(i).CodImpAdicSII)
            End If
            Call xAddTag(xDoc, xDetalle, "MontoItem", Format(DTE.DetDTE(i).SubTotal, FMTNUMDTE))
         
         Else   'línea de detalle en blanco o 0
            Call xAddTag(xDoc, xDetalle, "NmbItem", Ansi2XmlTxt(DTE.DetDTE(i).Producto))
            Call xAddTag(xDoc, xDetalle, "DscItem", Ansi2XmlTxt(DTE.DetDTE(i).Descrip))
            Call xAddTag(xDoc, xDetalle, "MontoItem", Format(DTE.DetDTE(i).SubTotal, FMTNUMDTE))
         End If
   
   Next i
   
   
   'Subtotales Informativos (utilizados para entregar el detalle de los impuestos adicionales
   For i = 0 To UBound(DTE.ImpAdic)
      If DTE.ImpAdic(i).IdImpAdic > 0 Then
         Set xSubTotInfo = xAddTag(xDoc, xDTE, "SubTotInfo")
            
            Call xAddTag(xDoc, xSubTotInfo, "NroSTI", i + 1)
            
            Call xAddTag(xDoc, xSubTotInfo, "GlosaSTI", Left(Ansi2XmlTxt(DTE.ImpAdic(i).DescImpAdic), 40))
            Call xAddTag(xDoc, xSubTotInfo, "SubTotNetoSTI", DTE.ImpAdic(i).NetoImpAdic)
            Call xAddTag(xDoc, xSubTotInfo, "SubTotAdicSTI", DTE.ImpAdic(i).MontoImpAdic)
            Call xAddTag(xDoc, xSubTotInfo, "ValSubtotSTI", DTE.ImpAdic(i).NetoImpAdic)
      Else
         Exit For
      End If
   Next i
   
   
'   For i = 0 To UBound(DTE.DetDTE)
'
'      If DTE.DetDTE(i).IdImpAdic <> 0 Then
'
'         Set xSubTotInfo = xAddTag(xDoc, xDTE, "SubTotInfo")
'
'            Call xAddTag(xDoc, xSubTotInfo, "NroSTI", i + 1)
'
'            Call xAddTag(xDoc, xSubTotInfo, "GlosaSTI", Left(Ansi2XmlTxt(DTE.DetDTE(i).DescImpAdic), 40))
'            Call xAddTag(xDoc, xSubTotInfo, "SubTotNetoSTI", DTE.DetDTE(i).SubTotal)
'            Call xAddTag(xDoc, xSubTotInfo, "SubTotAdicSTI", DTE.DetDTE(i).MontoImpAdic)
'            Call xAddTag(xDoc, xSubTotInfo, "ValSubtotSTI", DTE.DetDTE(i).SubTotal)
'      End If
'
'   Next i
   
   
  
   'Descuento Global
   If DTE.DesctoGlobal > 0 Then
         
      i = 1
   
      Set xDescto = xAddTag(xDoc, xDTE, "DscRcgGlobal")
         Call xAddTag(xDoc, xDescto, "NroLinDR", i)
         Call xAddTag(xDoc, xDescto, "TpoMov", "D")
         Call xAddTag(xDoc, xDescto, "GlosaDR", "")
         Call xAddTag(xDoc, xDescto, "TpoValor", "%")
         Call xAddTag(xDoc, xDescto, "ValorDR", DTE.DesctoGlobal)
      
   End If
   
   'referencias
   For i = 0 To UBound(DTE.Referencia)
      
      If DTE.Referencia(i).IdTipoDocRef = 0 Then
         Exit For
      End If
      
      Set xRef = xAddTag(xDoc, xDTE, "Referencia")
         
         Call xAddTag(xDoc, xRef, "NroLinRef", i + 1)
         Call xAddTag(xDoc, xRef, "TpoDocRef", DTE.Referencia(i).CodDocRefSII)
         Call xAddTag(xDoc, xRef, "FolioRef", DTE.Referencia(i).FolioRef)
         Call xAddTag(xDoc, xRef, "FchRef", Format(DTE.Referencia(i).FechaRef, FMTDATEDTE))
         If DTE.Referencia(i).CodRefSII > 0 Then
            Call xAddTag(xDoc, xRef, "CodRef", DTE.Referencia(i).CodRefSII)
         End If
         Call xAddTag(xDoc, xRef, "RazonRef", Ansi2XmlTxt(DTE.Referencia(i).RazonReferencia))
   
   Next i

   If gConectData.Proveedor = PROV_LP Then
      Call xAddTag(xDoc, xDTE, "TmstFirma", 0)
   End If
   
   If W.InDesign Or gDebug > 0 Then
'      xDoc.Save ("D:\Temp\kk.xml")
      Call xDoc.Save(W.AppPath & "\Log\logdte.xml")
   End If
   
   
   XmlDTE = xDoc.xml
   If gConectData.Proveedor = PROV_LP Then
      'eliminamos posibles newlines
      XmlDTE = ReplaceStr(XmlDTE, vbCr, "")
      XmlDTE = ReplaceStr(XmlDTE, vbLf, "")
   Else
      'eliminamos atributo vacío del tag Documento
      XmlDTE = ReplaceStr(XmlDTE, "xmlns=""""", "")
      Idx = InStr(XmlDTE, "<Encabezado>")
      If Idx > 0 Then
         Idx = Idx + Len("<Encabezado>")
         
         XmlDTE = ReplaceStrStartAt(XmlDTE, """", "&quot;", Idx)
         XmlDTE = ReplaceStrStartAt(XmlDTE, "'", "&apos;", Idx)
      End If
   End If
   
   GenJsonDTE = XmlDTE
   
   Set xDoc = Nothing
   
End Function

' ** Funcion que crea el XML de un DTE

Public Function GenXMLDTE(DTE As DTE_t) As String
   Dim xDoc As MSXML2.DOMDocument
   Dim xIDTE As MSXML2.IXMLDOMNode
   Dim xDTE As MSXML2.IXMLDOMNode
   Dim xNode As MSXML2.IXMLDOMNode, xChofer As MSXML2.IXMLDOMNode
   Dim xEncab As MSXML2.IXMLDOMNode
   Dim xDetalle As MSXML2.IXMLDOMNode
   Dim xRef As MSXML2.IXMLDOMNode
   Dim xDescto As MSXML2.IXMLDOMNode
   Dim xObse As MSXML2.IXMLDOMNode
   Dim Folio As String
   Dim IdDoc As String
   Dim objElement As MSXML2.IXMLDOMElement
   Dim i As Integer
   Dim xCodItem As MSXML2.IXMLDOMNode, xImpAdic As MSXML2.IXMLDOMNode
   Dim xSubTotInfo As MSXML2.IXMLDOMNode
   Dim XmlDTE As String
   Dim Idx As Integer
   Dim TotDescIva As Long
   
   
   TotDescIva = 0
   Folio = ""
   If gConectData.Proveedor = PROV_LP Then
      Folio = "1"
   End If
   
   IdDoc = "F" & Folio & "T" & DTE.CodDocSII
   
   Call AddDebug("GenXMLDTE: idDoc=" & IdDoc)
   
   Set xDoc = New MSXML2.DOMDocument
   Set xNode = xDoc.createProcessingInstruction("xml", "version='1.0'")
   xDoc.appendChild xNode
   Set xNode = Nothing
      
   Set xNode = xAddTagWithAttr(xDoc, xDoc, "DTE", "", "version", "1.0")

   If gConectData.Proveedor = PROV_LP Then
      If DTE.EsExport Then
         Set xDTE = xAddTagWithAttr(xDoc, xNode, "Exportaciones", "", "ID", IdDoc)
      Else
         Set xDTE = xAddTagWithAttr(xDoc, xNode, "Documento", "", "ID", IdDoc)
      End If
   Else   'Acepta
      Set objElement = xNode
      objElement.setAttributeNode xAddAttrib(xDoc, xNode, "xmlns", "http://www.sii.cl/SiiDte")
      objElement.setAttributeNode xAddAttrib(xDoc, xNode, "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
      
      If DTE.EsExport Then
         Set xDTE = xAddTag(xDoc, xNode, "Exportaciones")
      Else
         Set xDTE = xAddTag(xDoc, xNode, "Documento")
      End If
   End If
   
   Set xEncab = xAddTag(xDoc, xDTE, "Encabezado")
   Set xNode = xAddTag(xDoc, xEncab, "IdDoc")
      Call xAddTag(xDoc, xNode, "TipoDTE", DTE.CodDocSII)
      Call xAddTag(xDoc, xNode, "Folio", Folio)
      Call xAddTag(xDoc, xNode, "FchEmis", Format(DTE.Fecha, FMTDATEDTE))
      
      If DTE.CodDocSII = CODDOCDTESII_GUIADESPACHO Then
         Call xAddTag(xDoc, xNode, "TipoDespacho", DTE.TipoDespacho)
         Call xAddTag(xDoc, xNode, "IndTraslado", DTE.Traslado)
      End If
      If DTE.EsExport Then
         Call xAddTag(xDoc, xNode, "IndServicio", DTE.FactExp.CodIndServicio)
      End If
      If DTE.FormaDePago > 0 Then
         Call xAddTag(xDoc, xNode, "FmaPago", DTE.FormaDePago)
      End If
      If DTE.FechaVenc > 0 Then   'tiene que ir después de tipo despacho. El orden afecta
         Call xAddTag(xDoc, xNode, "FchVenc", Format(DTE.FechaVenc, FMTDATEDTE))
      End If
      If DTE.DetFormaPago > 0 Then
         Call xAddTag(xDoc, xNode, "DetFormaPago", DTE.TextDetFormaPago)
      End If
   
   Set xNode = xAddTag(xDoc, xEncab, "Emisor")
      Call xAddTag(xDoc, xNode, "RUTEmisor", gEmpresa.Rut & "-" & DV_Rut(gEmpresa.Rut))
      Call xAddTag(xDoc, xNode, "RznSoc", Left(Ansi2XmlTxt(gEmpresa.RazonSocial), 100))
      Call xAddTag(xDoc, xNode, "GiroEmis", Left(Ansi2XmlTxt(gEmpresa.Giro), 80))
      Call xAddTag(xDoc, xNode, "Acteco", Left(Ansi2XmlTxt(gEmpresa.CodActEcono), 6))
      Call xAddTag(xDoc, xNode, "DirOrigen", Left(Ansi2XmlTxt(gEmpresa.Direccion), 60))
      Call xAddTag(xDoc, xNode, "CmnaOrigen", Left(Ansi2XmlTxt(gEmpresa.Comuna), 20))
      Call xAddTag(xDoc, xNode, "CiudadOrigen", Left(Ansi2XmlTxt(gEmpresa.Ciudad), 20))
      If DTE.Vendedor > 0 Then
        Call xAddTag(xDoc, xNode, "Vendedor", Left(Ansi2XmlTxt(DTE.TextVendedor), 20))
      End If
   
   Set xNode = xAddTag(xDoc, xEncab, "Receptor")
      If DTE.NotValidRut Then
         Call xAddTag(xDoc, xNode, "RUTRecep", ENTIMP_RUT)
         Call xAddTag(xDoc, xNode, "RznSocRecep", ENTIMP_RSOCIAL)
      Else
         Call xAddTag(xDoc, xNode, "RUTRecep", DTE.Rut & "-" & DV_Rut(DTE.Rut))
         Call xAddTag(xDoc, xNode, "RznSocRecep", Left(Ansi2XmlTxt(DTE.RazonSocial), 100))
      End If
      Call xAddTag(xDoc, xNode, "GiroRecep", Left(Ansi2XmlTxt(DTE.Giro), 40))
      Call xAddTag(xDoc, xNode, "Contacto", Left(Ansi2XmlTxt(DTE.Contacto), 80))
      Call xAddTag(xDoc, xNode, "DirRecep", Left(Ansi2XmlTxt(DTE.Direccion), 70))
      Call xAddTag(xDoc, xNode, "CmnaRecep", Left(Ansi2XmlTxt(DTE.Comuna), 20))
      Call xAddTag(xDoc, xNode, "CiudadRecep", Left(Ansi2XmlTxt(DTE.Ciudad), 20))
      Call xAddTag(xDoc, xNode, "MailReceptor", Left(Ansi2XmlTxt(DTE.MailReceptor), 80))


   If DTE.EsExport Or DTE.EsGuiaDesp Then
      Set xNode = xAddTag(xDoc, xEncab, "Transporte")
      
      If DTE.EsGuiaDesp Then
         If DTE.GuiaDesp.Patente <> "" Then
            Call xAddTag(xDoc, xNode, "Patente", Left(Ansi2XmlTxt(Trim(DTE.GuiaDesp.Patente)), 8))
         End If
                  
         If (DTE.GuiaDesp.RutChofer <> "" Or DTE.GuiaDesp.NombreChofer <> "") And (DTE.TipoDespacho = GD_DESPEMICLI Or DTE.TipoDespacho = GD_DESPEMIOTRO) Then
'            Call xAddTag(xDoc, xNode, "RUTTrans", DTE.GuiaDesp.RutChofer & "-" & DV_Rut(DTE.GuiaDesp.RutChofer))
'            Call xAddTag(xDoc, xNode, "RUTChofer", DTE.GuiaDesp.RutChofer & "-" & DV_Rut(DTE.GuiaDesp.RutChofer))   'Si usamos RUTChofer el SII reclama al validar el archivo
'            Call xAddTag(xDoc, xNode, "NombreChofer", Left(Ansi2XmlTxt(DTE.GuiaDesp.NombreChofer), 30))             'se cambia por mensaje de error de validación del SII
            
            Set xChofer = xAddTag(xDoc, xNode, "Chofer") ' 10 may 2019 - pam: se agrega un nivel
            
            If DTE.GuiaDesp.RutChofer <> "" Then   ' primero debe ser el RUT
               Call xAddTag(xDoc, xChofer, "RUTChofer", DTE.GuiaDesp.RutChofer & "-" & DV_Rut(DTE.GuiaDesp.RutChofer))   'Si usamos RUTChofer el SII reclama al validar el archivo
            End If
            
            If DTE.GuiaDesp.NombreChofer <> "" Then
               Call xAddTag(xDoc, xChofer, "NombreChofer", Left(Ansi2XmlTxt(DTE.GuiaDesp.NombreChofer), 30))
            End If
            
         End If
         
      End If
      
      If DTE.EsExport Then
         Call xAddTag(xDoc, xNode, "CiudadDest", Left(Ansi2XmlTxt(DTE.Ciudad), 20))
         Call xAddTag(xDoc, xNode, "CodModVenta", DTE.FactExp.CodModVenta)
         Call xAddTag(xDoc, xNode, "CodClauVenta", DTE.FactExp.CodClausulaVenta)
         Call xAddTag(xDoc, xNode, "TotClauVenta", DTE.FactExp.TotClausulaVenta)
         Call xAddTag(xDoc, xNode, "CodViaTransp", DTE.FactExp.CodViaTransporte)
         Call xAddTag(xDoc, xNode, "CodPtoEmbarque", DTE.FactExp.CodPuertoEmbarque)
         Call xAddTag(xDoc, xNode, "CodPtoDesemb", DTE.FactExp.CodPuertoEmbarque)
         Call xAddTag(xDoc, xNode, "TotBultos", DTE.FactExp.TotalBultos)
      End If
      
   End If
 
   TotDescIva = 0
   Set xNode = xAddTag(xDoc, xEncab, "Totales")
      Call xAddTag(xDoc, xNode, "MntNeto", DTE.Neto)
      Call xAddTag(xDoc, xNode, "MntExe", DTE.Exento)
      If DTE.TasaIVA > 0 Then
         Call xAddTag(xDoc, xNode, "TasaIVA", DTE.TasaIVA * 100)
         Call xAddTag(xDoc, xNode, "IVA", DTE.Iva)
      End If
      
      If DTE.ImpAdic(0).IdImpAdic > 0 And DTE.ImpAdic(0).IdImpAdic <> 8 Then   'hay impuestos adicionales
         For i = 0 To UBound(DTE.ImpAdic)
            If DTE.ImpAdic(i).IdImpAdic > 0 Then
               Set xImpAdic = xAddTag(xDoc, xNode, "ImptoReten")
                  Call xAddTag(xDoc, xImpAdic, "TipoImp", DTE.ImpAdic(i).IdImpAdicSII)
                  Call xAddTag(xDoc, xImpAdic, "TasaImp", ReplaceStr(Format(DTE.ImpAdic(i).TasaImpAdic, DBLFMT2), ",", "."))
                  Call xAddTag(xDoc, xImpAdic, "MontoImp", DTE.ImpAdic(i).MontoImpAdic)
                  
                  'Call xAddTag(xDoc, xNode, "GrntDep", 0)
                  
            Else
               Exit For
            End If
         Next i
       ElseIf DTE.ImpAdic(0).IdImpAdic > 0 And DTE.ImpAdic(0).IdImpAdic = 8 Then   'hay impuestos adicionales
         For i = 0 To UBound(DTE.ImpAdic)
            If DTE.ImpAdic(i).IdImpAdic > 0 Then
'               Set xImpAdic = xAddTag(xDoc, xNode, "Rebaja")
'                  Call xAddTag(xDoc, xImpAdic, "TipoReb", DTE.ImpAdic(i).IdImpAdic)
'                  Call xAddTag(xDoc, xImpAdic, "DescReb", DTE.ImpAdic(i).DescImpAdic)
'                  Call xAddTag(xDoc, xImpAdic, "TasaReb", ReplaceStr(Format(DTE.ImpAdic(i).TasaImpAdic, DBLFMT2), ",", "."))
'                  Call xAddTag(xDoc, xImpAdic, "MontoReb", DTE.ImpAdic(i).MontoImpAdic)
                  TotDescIva = DTE.ImpAdic(i).MontoImpAdic
                  Call xAddTag(xDoc, xNode, "CredEC", TotDescIva)
            Else
               Exit For
            End If
         Next i
      End If
      
      
      
      Call xAddTag(xDoc, xNode, "MntTotal", DTE.Total)
      
      
   If DTE.EsExport Then
      Set xNode = xAddTag(xDoc, xEncab, "OtraMoneda")
         Call xAddTag(xDoc, xNode, "TpoMoneda", DTE.FactExp.CodMoneda)
         Call xAddTag(xDoc, xNode, "TpoCambio", DTE.FactExp.TipoCambio)
   End If
            
   'Detalle
   For i = 0 To UBound(DTE.DetDTE)
      
      If DTE.DetDTE(i).Producto = "" And i > 0 Then   'tiene que ir al menos una linea de detalle
         Exit For
      End If
      
      Set xDetalle = xAddTag(xDoc, xDTE, "Detalle")
         
         Call xAddTag(xDoc, xDetalle, "NroLinDet", i + 1)
         
         If DTE.DetDTE(i).Cantidad > 0 Then
            Set xCodItem = xAddTag(xDoc, xDetalle, "CdgItem")
               Call xAddTag(xDoc, xCodItem, "TpoCodigo", DTE.DetDTE(i).TipoCod)
               Call xAddTag(xDoc, xCodItem, "VlrCodigo", DTE.DetDTE(i).CodProd)
            If DTE.DetDTE(i).EsExento Then
               Call xAddTag(xDoc, xDetalle, "IndExe", 1)
            End If
            Call xAddTag(xDoc, xDetalle, "NmbItem", Left(Ansi2XmlTxt(DTE.DetDTE(i).Producto), 80))
            Call xAddTag(xDoc, xDetalle, "DscItem", Left(Ansi2XmlTxt(DTE.DetDTE(i).Descrip), 1000))
            
            ' Call xAddTag(xDoc, xDetalle, "QtyItem", Format(DTE.DetDTE(i).Cantidad, FMTNUMDTE))
            ' 13 oct 2017 - pam: permite cantidad con decimales
            Call xAddTag(xDoc, xDetalle, "QtyItem", ReplaceStr(Round(DTE.DetDTE(i).Cantidad, 2), ",", "."))
            Call xAddTag(xDoc, xDetalle, "UnmdItem", Left(Ansi2XmlTxt(DTE.DetDTE(i).UMedida), 4))
                        
            Call xAddTag(xDoc, xDetalle, "PrcItem", Format(DTE.DetDTE(i).Precio, FMTNUMDTE))
            If DTE.DetDTE(i).PjeDescto > 0 Then
               Call xAddTag(xDoc, xDetalle, "DescuentoPct", ReplaceStr(Format(DTE.DetDTE(i).PjeDescto, DBLFMT2), ",", "."))
               Call xAddTag(xDoc, xDetalle, "DescuentoMonto", Format(DTE.DetDTE(i).MontoDescto, FMTNUMDTE))
            End If
            If DTE.DetDTE(i).CodImpAdicSII <> "" And DTE.DetDTE(i).CodImpAdicSII <> "0" And DTE.DetDTE(i).CodImpAdicSII <> "126" Then
               Call xAddTag(xDoc, xDetalle, "CodImpAdic", DTE.DetDTE(i).CodImpAdicSII)
            End If
            Call xAddTag(xDoc, xDetalle, "MontoItem", Format(DTE.DetDTE(i).SubTotal, FMTNUMDTE))
         
         Else   'línea de detalle en blanco o 0
            Call xAddTag(xDoc, xDetalle, "NmbItem", Ansi2XmlTxt(DTE.DetDTE(i).Producto))
            Call xAddTag(xDoc, xDetalle, "DscItem", Ansi2XmlTxt(DTE.DetDTE(i).Descrip))
            Call xAddTag(xDoc, xDetalle, "MontoItem", Format(DTE.DetDTE(i).SubTotal, FMTNUMDTE))
         End If
   
   Next i
   
   
   'Subtotales Informativos (utilizados para entregar el detalle de los impuestos adicionales
   For i = 0 To UBound(DTE.ImpAdic)
      If DTE.ImpAdic(i).IdImpAdic > 0 Then
         Set xSubTotInfo = xAddTag(xDoc, xDTE, "SubTotInfo")
            
            Call xAddTag(xDoc, xSubTotInfo, "NroSTI", i + 1)
            
            Call xAddTag(xDoc, xSubTotInfo, "GlosaSTI", Left(Ansi2XmlTxt(DTE.ImpAdic(i).DescImpAdic), 40))
            Call xAddTag(xDoc, xSubTotInfo, "SubTotNetoSTI", DTE.ImpAdic(i).NetoImpAdic)
            Call xAddTag(xDoc, xSubTotInfo, "SubTotAdicSTI", DTE.ImpAdic(i).MontoImpAdic)
            Call xAddTag(xDoc, xSubTotInfo, "ValSubtotSTI", DTE.ImpAdic(i).NetoImpAdic)
      Else
         Exit For
      End If
   Next i
   
   
'   For i = 0 To UBound(DTE.DetDTE)
'
'      If DTE.DetDTE(i).IdImpAdic <> 0 Then
'
'         Set xSubTotInfo = xAddTag(xDoc, xDTE, "SubTotInfo")
'
'            Call xAddTag(xDoc, xSubTotInfo, "NroSTI", i + 1)
'
'            Call xAddTag(xDoc, xSubTotInfo, "GlosaSTI", Left(Ansi2XmlTxt(DTE.DetDTE(i).DescImpAdic), 40))
'            Call xAddTag(xDoc, xSubTotInfo, "SubTotNetoSTI", DTE.DetDTE(i).SubTotal)
'            Call xAddTag(xDoc, xSubTotInfo, "SubTotAdicSTI", DTE.DetDTE(i).MontoImpAdic)
'            Call xAddTag(xDoc, xSubTotInfo, "ValSubtotSTI", DTE.DetDTE(i).SubTotal)
'      End If
'
'   Next i
   
   
  
   'Descuento Global
   If DTE.DesctoGlobal > 0 Then
         
      i = 1
   
      Set xDescto = xAddTag(xDoc, xDTE, "DscRcgGlobal")
         Call xAddTag(xDoc, xDescto, "NroLinDR", i)
         Call xAddTag(xDoc, xDescto, "TpoMov", "D")
         Call xAddTag(xDoc, xDescto, "GlosaDR", "")
         Call xAddTag(xDoc, xDescto, "TpoValor", "%")
         Call xAddTag(xDoc, xDescto, "ValorDR", DTE.DesctoGlobal)
      
   End If
   
   'referencias
   For i = 0 To UBound(DTE.Referencia)
      
      If DTE.Referencia(i).IdTipoDocRef = 0 Then
         Exit For
      End If
      
      Set xRef = xAddTag(xDoc, xDTE, "Referencia")
         
         Call xAddTag(xDoc, xRef, "NroLinRef", i + 1)
         Call xAddTag(xDoc, xRef, "TpoDocRef", DTE.Referencia(i).CodDocRefSII)
         Call xAddTag(xDoc, xRef, "FolioRef", DTE.Referencia(i).FolioRef)
         Call xAddTag(xDoc, xRef, "FchRef", Format(DTE.Referencia(i).FechaRef, FMTDATEDTE))
         If DTE.Referencia(i).CodRefSII > 0 Then
            Call xAddTag(xDoc, xRef, "CodRef", DTE.Referencia(i).CodRefSII)
         End If
         Call xAddTag(xDoc, xRef, "RazonRef", Ansi2XmlTxt(DTE.Referencia(i).RazonReferencia))
   
   Next i
   
   Set xObse = xAddTag(xDoc, xDTE, "Observaciones")
   Call xAddTag(xDoc, xObse, "Observaciones", DTE.Observaciones)

   If gConectData.Proveedor = PROV_LP Then
      Call xAddTag(xDoc, xDTE, "TmstFirma", 0)
   End If
   
   If W.InDesign Or gDebug > 0 Then
'      xDoc.Save ("D:\Temp\kk.xml")
      Call xDoc.Save(W.AppPath & "\Log\logdte.xml")
   End If
   
   
   XmlDTE = xDoc.xml
   If gConectData.Proveedor = PROV_LP Then
      'eliminamos posibles newlines
      XmlDTE = ReplaceStr(XmlDTE, vbCr, "")
      XmlDTE = ReplaceStr(XmlDTE, vbLf, "")
   Else
      'eliminamos atributo vacío del tag Documento
      XmlDTE = ReplaceStr(XmlDTE, "xmlns=""""", "")
      Idx = InStr(XmlDTE, "<Encabezado>")
      If Idx > 0 Then
         Idx = Idx + Len("<Encabezado>")
         
         XmlDTE = ReplaceStrStartAt(XmlDTE, """", "&quot;", Idx)
         XmlDTE = ReplaceStrStartAt(XmlDTE, "'", "&apos;", Idx)
      End If
   End If
   
   Dim objNode As Object
   Dim bTmp()   As Byte
   
   bTmp() = StrConv(XmlDTE, vbFromUnicode)
   Set xDoc = New MSXML2.DOMDocument
  Set objNode = xDoc.createElement("b64")

  objNode.DataType = "bin.base64"
  objNode.nodeTypedValue = bTmp()
  Dim base64 As String
  base64 = objNode.Text
  'base64 = ReplaceStr(base64, vbCr, "")
  'base64 = ReplaceStr(base64, vbLf, "")
  base64 = Replace(base64, vbCr, "")
  base64 = Replace(base64, vbLf, "")
  
'     Dim myMSXML As Object
'   Set myMSXML = CreateObject("Microsoft.XmlHttp")
'myMSXML.Open "POST", URL_TECNOBACKCERT & "v1/emitir/xml", False
'myMSXML.setRequestHeader "x-api-key", API_KEY_TECNOBACK
'myMSXML.setRequestHeader "Content-Type", "application/json"
'Dim Req As String
'Req = ""
'Call MkJSon(Req, "session_id", "001d27702d95bab5808683dc0ebd4eaaf1140190953")
'Call MkJSon(Req, "rut_signer", "9199548-4")
'Call MkJSon(Req, "encoding_xml", "UTF-8")
'Call MkJSon(Req, "xml_b64", base64)
'myMSXML.send "{ " & Mid(Req, 2) & " } "
'MsgBox myMSXML.responseText
  
   'GenXMLDTE = XmlDTE
   GenXMLDTE = base64
   
   Set xDoc = Nothing
   
End Function

Public Function Ansi2XmlTxt(ByVal Buf As String)
   Dim s As String

   Buf = ReplaceStr(Buf, vbTab, " ")
   Buf = ReplaceStr(Buf, vbCrLf, " ")
   Buf = ReplaceStr(Buf, vbCr, " ")
   Buf = ReplaceStr(Buf, vbLf, " ")

   If gConectData.Proveedor = PROV_LP Then
      Ansi2XmlTxt = HtmlEscape2(Buf)
   Else ' Acepta
      Ansi2XmlTxt = AsciiToLatin1(Buf)
   End If
   
End Function
