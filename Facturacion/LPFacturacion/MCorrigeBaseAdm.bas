Attribute VB_Name = "McorrigeBaseAdm"
Option Explicit


Private lDbVer As Integer
Private lUpdOK As Boolean

' Para hacer manteciones a ciertas tablas con manejo de versión
Public Sub CorrigeBaseAdm()
   
   On Error Resume Next
   
   lDbVer = 0
   lUpdOK = True
   
   If Not CorrigeBaseAdm_V1() Then
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V2() Then
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V3() Then               'agregada 13 sep 2018
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V4() Then               'agregada 12 mar 2019
      Exit Sub
   End If
      
   If Not CorrigeBaseAdm_V5() Then               'agregada 29 ago 2019
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V6() Then               'agregada 1 julio 2021
      Exit Sub
   End If
   
   If Not CorrigeBaseAdm_V7() Then               'agregada 30 OCT FPR
      Exit Sub
   End If

      
'   If lDbVer > 2 Then
'      MsgBox1 "¡ ATENCION !" & vbCrLf & vbCrLf & "La base de datos corresponde a una versión posterior de este programa." & vbCrLf & "Debe actualizar el programa antes de continuar, de lo contrario podría dañar la información..", vbCritical
'      Call CloseDb(DbMain)
'      End
'   End If

End Sub

Public Function CorrigeBaseAdm_V7() As Boolean   'agregada 29 sep 2019
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 7 -----------------------------------

   If lDbVer = 7 And lUpdOK = True Then
      
      'SI NO EXISTEN LOS INSERTA
      
      Q1 = "INSERT INTO REGIONES (CODIGO, COMUNA) "
      Q1 = Q1 & " SELECT DISTINCT  16,'CHILLAN' FROM REGIONES WHERE NOT EXISTS (SELECT 1 FROM REGIONES WHERE COMUNA = 'CHILLAN')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO REGIONES (CODIGO, COMUNA) "
      Q1 = Q1 & " SELECT DISTINCT  16,'BULNES' FROM REGIONES WHERE NOT EXISTS (SELECT 1 FROM REGIONES WHERE COMUNA = 'BULNES')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO REGIONES (CODIGO, COMUNA) "
      Q1 = Q1 & " SELECT DISTINCT  16,'CHILLAN VIEJO' FROM REGIONES WHERE NOT EXISTS (SELECT 1 FROM REGIONES WHERE COMUNA = 'CHILLAN VIEJO')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO REGIONES (CODIGO, COMUNA) "
      Q1 = Q1 & " SELECT DISTINCT  16,'EL CARMEN' FROM REGIONES WHERE NOT EXISTS (SELECT 1 FROM REGIONES WHERE COMUNA = 'EL CARMEN')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO REGIONES (CODIGO, COMUNA) "
      Q1 = Q1 & " SELECT DISTINCT  16,'PEMUCO' FROM REGIONES WHERE NOT EXISTS (SELECT 1 FROM REGIONES WHERE COMUNA = 'PEMUCO')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO REGIONES (CODIGO, COMUNA) "
      Q1 = Q1 & " SELECT DISTINCT  16,'PINTO' FROM REGIONES WHERE NOT EXISTS (SELECT 1 FROM REGIONES WHERE COMUNA = 'PINTO')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO REGIONES (CODIGO, COMUNA) "
      Q1 = Q1 & " SELECT DISTINCT  16,'QUILLON' FROM REGIONES WHERE NOT EXISTS (SELECT 1 FROM REGIONES WHERE COMUNA = 'QUILLON')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO REGIONES (CODIGO, COMUNA) "
      Q1 = Q1 & " SELECT DISTINCT  16,'SAN IGNACIO' FROM REGIONES WHERE NOT EXISTS (SELECT 1 FROM REGIONES WHERE COMUNA = 'SAN IGNACIO')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO REGIONES (CODIGO, COMUNA) "
      Q1 = Q1 & " SELECT DISTINCT  16,'YUNGAY' FROM REGIONES WHERE NOT EXISTS (SELECT 1 FROM REGIONES WHERE COMUNA = 'YUNGAY')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO REGIONES (CODIGO, COMUNA) "
      Q1 = Q1 & " SELECT DISTINCT  16,'QUIRIHUE' FROM REGIONES WHERE NOT EXISTS (SELECT 1 FROM REGIONES WHERE COMUNA = 'QUIRIHUE')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO REGIONES (CODIGO, COMUNA) "
      Q1 = Q1 & " SELECT DISTINCT  16,'COBQUECURA' FROM REGIONES WHERE NOT EXISTS (SELECT 1 FROM REGIONES WHERE COMUNA = 'COBQUECURA')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO REGIONES (CODIGO, COMUNA) "
      Q1 = Q1 & " SELECT DISTINCT  16,'COELEMU' FROM REGIONES WHERE NOT EXISTS (SELECT 1 FROM REGIONES WHERE COMUNA = 'COELEMU')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO REGIONES (CODIGO, COMUNA) "
      Q1 = Q1 & " SELECT DISTINCT  16,'NINHUE' FROM REGIONES WHERE NOT EXISTS (SELECT 1 FROM REGIONES WHERE COMUNA = 'NINHUE')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO REGIONES (CODIGO, COMUNA) "
      Q1 = Q1 & " SELECT DISTINCT  16,'PORTEZUELO' FROM REGIONES WHERE NOT EXISTS (SELECT 1 FROM REGIONES WHERE COMUNA = 'PORTEZUELO')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO REGIONES (CODIGO, COMUNA) "
      Q1 = Q1 & " SELECT DISTINCT  16,'RANQUIL' FROM REGIONES WHERE NOT EXISTS (SELECT 1 FROM REGIONES WHERE COMUNA = 'RANQUIL')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO REGIONES (CODIGO, COMUNA) "
      Q1 = Q1 & " SELECT DISTINCT  16,'TREHUACO' FROM REGIONES WHERE NOT EXISTS (SELECT 1 FROM REGIONES WHERE COMUNA = 'TREHUACO')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO REGIONES (CODIGO, COMUNA) "
      Q1 = Q1 & " SELECT DISTINCT  16,'SAN CARLOS' FROM REGIONES WHERE NOT EXISTS (SELECT 1 FROM REGIONES WHERE COMUNA = 'SAN CARLOS')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO REGIONES (CODIGO, COMUNA) "
      Q1 = Q1 & " SELECT DISTINCT  16,'COIHUECO' FROM REGIONES WHERE NOT EXISTS (SELECT 1 FROM REGIONES WHERE COMUNA = 'COIHUECO')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO REGIONES (CODIGO, COMUNA) "
      Q1 = Q1 & " SELECT DISTINCT  16,'ÑIQUEN' FROM REGIONES WHERE NOT EXISTS (SELECT 1 FROM REGIONES WHERE COMUNA = 'ÑIQUEN')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO REGIONES (CODIGO, COMUNA) "
      Q1 = Q1 & " SELECT DISTINCT  16,'SAN FABIAN' FROM REGIONES WHERE NOT EXISTS (SELECT 1 FROM REGIONES WHERE COMUNA = 'SAN FABIAN')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO REGIONES (CODIGO, COMUNA) "
      Q1 = Q1 & " SELECT DISTINCT  16,'SAN NICOLAS' FROM REGIONES WHERE NOT EXISTS (SELECT 1 FROM REGIONES WHERE COMUNA = 'SAN NICOLAS')"
      Call ExecSQL(DbMain, Q1)
      
      
      'SI EXISTEN LOS UPDATEA
      
      Q1 = "UPDATE REGIONES SET CODIGO = 16 WHERE  COMUNA = 'CHILLAN'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE REGIONES SET CODIGO = 16 WHERE  COMUNA = 'BULNES'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE REGIONES SET CODIGO = 16 WHERE  COMUNA = 'CHILLAN VIEJO'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE REGIONES SET CODIGO = 16 WHERE  COMUNA = 'EL CARMEN'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE REGIONES SET CODIGO = 16 WHERE  COMUNA = 'PEMUCO'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE REGIONES SET CODIGO = 16 WHERE  COMUNA = 'PINTO'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE REGIONES SET CODIGO = 16 WHERE  COMUNA = 'QUILLON'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE REGIONES SET CODIGO = 16 WHERE  COMUNA = 'SAN IGNACIO'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE REGIONES SET CODIGO = 16 WHERE  COMUNA = 'YUNGAY'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE REGIONES SET CODIGO = 16 WHERE  COMUNA = 'QUIRIHUE'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE REGIONES SET CODIGO = 16 WHERE  COMUNA = 'COBQUECURA'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE REGIONES SET CODIGO = 16 WHERE  COMUNA = 'COELEMU'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE REGIONES SET CODIGO = 16 WHERE  COMUNA = 'NINHUE'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE REGIONES SET CODIGO = 16 WHERE  COMUNA = 'PORTEZUELO'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE REGIONES SET CODIGO = 16 WHERE  COMUNA = 'RANQUIL'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE REGIONES SET CODIGO = 16 WHERE  COMUNA = 'TREHUACO'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE REGIONES SET CODIGO = 16 WHERE  COMUNA = 'SAN CARLOS'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE REGIONES SET CODIGO = 16 WHERE  COMUNA = 'COIHUECO'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE REGIONES SET CODIGO = 16 WHERE  COMUNA = 'ÑIQUEN'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE REGIONES SET CODIGO = 16 WHERE  COMUNA = 'SAN FABIAN'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE REGIONES SET CODIGO = 16 WHERE  COMUNA = 'SAN NICOLAS'"
      Call ExecSQL(DbMain, Q1)
      
      If lUpdOK Then
         lDbVer = 8
         Q1 = "UPDATE Param SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBaseAdm_V7 = lUpdOK

End Function

Public Function CorrigeBaseAdm_V6() As Boolean   'agregada 1 jul 2021
   Dim Tbl As TableDef
   Dim fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset
   Dim Q1 As String, Q2 As String


   On Error Resume Next

   '--------------------- Versión 6 -----------------------------------

   If lDbVer = 6 And lUpdOK = True Then
   
      'Agregamos campo TipoIVARetenido a tabla TipoValor
      Set Tbl = DbMain.TableDefs("TipoValor")

      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("TipoIVARetenido", dbByte)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoValor.TipoIVARetenido", vbExclamation
         lUpdOK = False
      End If
      
      'Dejamos TAsa en NULL donde vale 0
      Q1 = "UPDATE TipoValor SET Tasa = ' ' WHERE Tasa = 0"
      Call ExecSQL(DbMain, Q1)
      
      'actualizamos códigos SII de algunos impuestos adicionales
      Q1 = "UPDATE TipoValor SET CodSIIDTE = '15' WHERE Codigo = " & LIBVENTAS_IVARETTOT
      Call ExecSQL(DbMain, Q1)
     
      Q1 = "UPDATE TipoValor SET CodSIIDTE = '14' WHERE Codigo = " & LIBVENTAS_RETMARGENCOM
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET CodSIIDTE = '271' WHERE Codigo = " & LIBVENTAS_ILABEDANALCAZUCAR
      Call ExecSQL(DbMain, Q1)

     
      'agregamos nuevos impuestos adicionales para Ventas
      Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, "
      Q1 = Q1 & " Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto, TipoIVARetenido)"
      Q1 = Q1 & " VALUES(" & LIB_VENTAS & ","

      Q2 = Q1 & LIBVENTAS_IVA_ANTICIP_FAENACARNE & ", 'IVA Anticip. Faenam. Carne', ' ', ' ', 0, ' ', ',3,4,5,', "
      Q2 = Q2 & "'IVA Anticip.','Faenamiento Carne',' ', 23, 5, 0, '17', 'IVA Anticipado Faenamiento Carne', 0)"

      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETPARCIAL_LEGUMBRES & ", 'IVA Ret. Parcial Legumbres', ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Parcial','Legumbres',' ', 24, ' ', 0, '30', 'IVA Retenido Parcial Legumbres', " & IVARET_PARCIAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETTOTAL_LEGUMBRES & ", 'IVA Ret. Total Legumbres' , ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Total','Legumbres',' ', 25, 100, 0, '301', 'IVA Retenido Total Legumbres', " & IVARET_TOTAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETTOTAL_SILVESTRES & ", 'IVA Ret. Total Silvestres' , ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Total','Silvestres',' ', 26, 100, 0, '31', 'IVA Retenido Total Silvestres', " & IVARET_TOTAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETPARCIAL_GANADO & ", 'IVA Ret. Parcial Ganado' , ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Parcial','Ganado',' ', 27, ' ', 0, '32', 'IVA Retenido Parcial Ganado', " & IVARET_PARCIAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETTOTAL_GANADO & ", 'IVA Ret. Total Ganado' , ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Total','Ganado',' ', 28, 100, 0, '321', 'IVA Retenido Total Ganado', " & IVARET_TOTAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETPARCIAL_MADERA & ", 'IVA Ret. Parcial Madera' , ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Parcial','Madera',' ', 29, ' ', 0, '33', 'IVA Retenido Parcial Madera', " & IVARET_PARCIAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETTOTAL_MADERA & ", 'IVA Ret. Total Madera' , ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Total','Madera',' ', 30, 100, 0, '331', 'IVA Retenido Total Madera', " & IVARET_TOTAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETPARCIAL_TRIGO & ", 'IVA Ret. Parcial Trigo' , ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Parcial','Trigo',' ', 31, ' ', 0, '34', 'IVA Retenido Parcial Trigo', " & IVARET_PARCIAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETTOTAL_TRIGO & ", 'IVA Ret. Total Trigo' , ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Total','Trigo',' ', 32, 100, 0, '341', 'IVA Retenido Total Trigo', " & IVARET_TOTAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETPARCIAL_ARROZ & ", 'IVA Ret. Parcial Arroz' , ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Parcial','Arroz',' ', 33, ' ', 0, '36', 'IVA Retenido Parcial Arroz', " & IVARET_PARCIAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETTOTAL_ARROZ & ", 'IVA Ret. Total Arroz' , ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Total','Arroz',' ', 34, 100, 0, '361', 'IVA Retenido Total Arroz', " & IVARET_TOTAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETPARCIAL_HIDROBIOLOGICAS & ", 'IVA Ret. Parcial Hidrobiológicas', ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Parcial','Hidrobiológicas',' ', 35, ' ', 0, '37', 'IVA Retenido Parcial Hidrobiológica',  " & IVARET_PARCIAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETTOTAL_HIDROBIOLÓGICAS & ", 'IVA Ret. Total Hidrobiológicas' , ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Total','Hidrobiológicas',' ', 36, 100, 0, '371', 'IVA Retenido Total Hidrobiológica', " & IVARET_TOTAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETTOTAL_CHATARRA & ", 'IVA Ret. Total Chatarra', ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Total','Chatarra',' ', 37, 100, 0, '38', 'IVA Retenido Total Chatarrae', " & IVARET_TOTAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETTOTAL_PPA & ", 'IVA Ret. Total PPA', ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Total','PPA',' ', 38, 100, 0, '39', 'IVA Retenido Total PPA', " & IVARET_TOTAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETTOTAL_CARTONES & ", 'IVA Ret. Total Cartones', ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Total','Cartones',' ', 39, 100, 0, '47', 'IVA Retenido Total Cartones', " & IVARET_TOTAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETPARCIAL_BERRIES & ", 'IVA Ret. Parcial Berries', ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Parcial','Berries',' ', 40, ' ', 0, '48', 'IVA Retenido Parcial Berries', " & IVARET_PARCIAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RETTOTAL_BERRIES & ", 'IVA Ret. Total Berries', ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'IVA Ret. Total','Berries',' ', 41, 100, 0, '481', 'IVA Retenido Total Berries', " & IVARET_TOTAL & ")"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_FACT_COMPRA_SIN_RET & ", 'Fact. compra sin Retención', ' ', ' ', 0, ' ', ',3,4,5,',"
      Q2 = Q2 & "'Factura de compra','sin Retención',' ', 42, 0, 0, '49', 'Factura de compra sin Retención', 0)"
      
      Call ExecSQL(DbMain, Q2)
      
      Q2 = Q1 & LIBVENTAS_IVA_RET_FACT_INICIO & ", 'IVA Retenido Factura de Inicio', ' ', ' ', 0, ' ', ',17,',"
      Q2 = Q2 & "'IVA Ret.','Fact. Inicio',' ', 43, 100, 0, '60', 'IVA Retenido Factura de Inicio', " & IVARET_TOTAL & ")"

      Call ExecSQL(DbMain, Q2)
      
      
      Q1 = "UPDATE TipoValor SET TipoIVARetenido = " & IVARET_TOTAL & " WHERE TipoLib = " & LIB_VENTAS & " AND Codigo IN ( " & LIBVENTAS_IVARETTOT & "," & LIBVENTAS_IVAADQCONSTINMUEBLES & ") "
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET TipoIVARetenido = " & IVARET_PARCIAL & " WHERE  TipoLib = " & LIB_VENTAS & " AND Codigo IN ( " & LIBVENTAS_IVARETPARC & ") "
      Call ExecSQL(DbMain, Q1)
            
       'Tipo IVA Retenido en Libro de Compras
       
      Q1 = "UPDATE TipoValor SET TipoIVARetenido = " & IVARET_PARCIAL & " WHERE TipoLib = " & LIB_COMPRAS
      Q1 = Q1 & " AND Codigo >=  " & LIBCOMPRAS_IVARETPARCTRIGO & " AND Codigo <= " & LIBCOMPRAS_IVARETPARCFAMBPASAS
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET TipoIVARetenido = " & IVARET_TOTAL & " WHERE TipoLib = " & LIB_COMPRAS
      Q1 = Q1 & " AND ((Codigo >=  " & LIBCOMPRAS_IVARETTOTCHATARRA & " AND Codigo <= " & LIBCOMPRAS_IVARETTOTCARTONES & ")"
      Q1 = Q1 & " OR Codigo =  " & LIBCOMPRAS_IVARETORO & ")"
      Call ExecSQL(DbMain, Q1)
     
     'Tipo IVA Retenido en Libro de Compras
      
      Q1 = "UPDATE TipoValor SET TipoIVARetenido = " & IVARET_TOTAL & " WHERE TipoLib = " & LIB_COMPRAS
      Q1 = Q1 & " AND Codigo = " & LIBCOMPRAS_IVARETTOT
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE TipoValor SET TipoIVARetenido = " & IVARET_PARCIAL & " WHERE TipoLib = " & LIB_COMPRAS
      Q1 = Q1 & " AND Codigo = " & LIBCOMPRAS_IVARETPARC
      Call ExecSQL(DbMain, Q1)
      
      'Insertamos impuestos específicos Diesel y Gasolina en libro de Ventas
      Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
      Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & LIBVENTAS_IMPDIESEL & ", 'Impuesto Específico Diesel', ' ', ' ', 0, ' ', ',1,3,4,', 'Imto. Esp.','Diesel',' ', 44, 100, 0, '28', 'Impuesto Específico Diesel')"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO TipoValor (TipoLib, Codigo, Valor, Diminutivo, Atributo, Multiple, CodF29, TipoDoc, Tit1, Tit2, CodImpSII, Orden, Tasa, EsRecuperable, CodSIIDTE, TitCompleto)"
      Q1 = Q1 & " VALUES(" & LIB_VENTAS & "," & LIBVENTAS_IMPGASOLINA & ", 'Impuesto Específico Gasolina', ' ', ' ', 0, ' ', ',1,3,4,', 'Imto. Esp.','Gasolina',' ', 45, 100, 0, '35', 'Impuesto Específico Gasolina')"
      Call ExecSQL(DbMain, Q1)
      
      If lUpdOK Then
         lDbVer = 7
         Q1 = "UPDATE Param SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   

   End If

   CorrigeBaseAdm_V6 = lUpdOK

End Function

Public Function CorrigeBaseAdm_V5() As Boolean   'agregada 29 ago 2019
   Dim Tbl As TableDef
   Dim fld As Field
   Dim Rc As Long
   Dim DbComun As String
   Dim Rs As Recordset
   Dim Q1 As String


   On Error Resume Next

   '--------------------- Versión 5 -----------------------------------

   If lDbVer = 5 And lUpdOK = True Then
   
      'Agregamos campo OldCodigo a CodActiv, esto por el cambio de codificación del SII a partir de Nov 2018
      Set Tbl = DbMain.TableDefs("CodActiv")
     
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("OldCodigo", dbText, 10)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "CodActiv.OldCodigo", vbExclamation
         lUpdOK = False
      End If
                 
      'achicamos campo codigo, tenía tamaño 255, lo cual es ridículo
      Call AlterField(DbMain, "CodActiv", "Codigo", dbText, 10)
     
      'agregamos los nuevos códigos de actividad económica del SII validos desde nov 2018
      Call UpdateCodActiv2018
      
      If lUpdOK Then
         lDbVer = 6
         Q1 = "UPDATE Param SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   

   End If

   CorrigeBaseAdm_V5 = lUpdOK

End Function

Public Function CorrigeBaseAdm_V4() As Boolean           'agregada 12 mar 2019
   Dim Tbl As TableDef
   Dim fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
   Dim Q1 As String
  

   On Error Resume Next
   
   
   '--------------------- Versión 4 -----------------------------------
   If lDbVer = 4 And lUpdOK = True Then
   
      'agregamos campo EsFijo a TipoDocRef
      
      Set Tbl = DbMain.TableDefs("TipoDocRef")
      
      Tbl.Fields.Append Tbl.CreateField("EsFijo", dbBoolean)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "TipoDocRef.EsFijo", vbExclamation
         lUpdOK = False
      End If
      
      Q1 = "UPDATE TipoDocRef SET EsFijo = -1"
      Call ExecSQL(DbMain, Q1)
      
      'agregamos tabla TipoVehiculo
      Set Tbl = New TableDef
      Tbl.Name = "TipoVehiculo"
      
      Err.Clear
      Set fld = Tbl.CreateField("Id", dbLong)
      fld.Attributes = dbAutoIncrField ' Autonumber
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "TipoVehiculo.Id", vbExclamation
         lUpdOK = False
      End If
                        
      Err.Clear
      Set fld = Tbl.CreateField("Codigo", dbText, 4)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "TipoVehiculo.Codigo", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("Nombre", dbText, 80)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "TipoVehiculo.Nombre", vbExclamation
         lUpdOK = False
      End If
      
      DbMain.TableDefs.Append Tbl
      If Err = 0 Then
         DbMain.TableDefs.Refresh
         
         Q1 = "CREATE UNIQUE INDEX Idx ON TipoVehiculo (Id) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)
         
         Q1 = "CREATE UNIQUE INDEX IdxCod ON TipoVehiculo (Codigo)"
         Rc = ExecSQL(DbMain, Q1, False)
         
      ElseIf Err <> 3010 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Tabla TipoVehiculo", vbExclamation
         lUpdOK = False

      End If

      'agregamos tipos de vehículos
      Q1 = "INSERT INTO TipoVehiculo (Codigo, Nombre) VALUES('1', 'Camión')"
      Call ExecSQL(DbMain, Q1)
      Q1 = "INSERT INTO TipoVehiculo (Codigo, Nombre) VALUES('2', 'Tractocamión')"
      Call ExecSQL(DbMain, Q1)
      Q1 = "INSERT INTO TipoVehiculo (Codigo, Nombre) VALUES('3', 'Furgón')"
      Call ExecSQL(DbMain, Q1)
      Q1 = "INSERT INTO TipoVehiculo (Codigo, Nombre) VALUES('4', 'Camioneta 5 Jeep')"
      Call ExecSQL(DbMain, Q1)
      Q1 = "INSERT INTO TipoVehiculo (Codigo, Nombre) VALUES('6', 'Cámara frigorífica')"
      Call ExecSQL(DbMain, Q1)
      Q1 = "INSERT INTO TipoVehiculo (Codigo, Nombre) VALUES('7', 'Furgón Térmico con unidad de frío')"
      Call ExecSQL(DbMain, Q1)
      Q1 = "INSERT INTO TipoVehiculo (Codigo, Nombre) VALUES('8', 'Generador para semiremolque de contenido frigorífico')"
      Call ExecSQL(DbMain, Q1)
      Q1 = "INSERT INTO TipoVehiculo (Codigo, Nombre) VALUES('9', 'Motor de bomba de semiremolque estanque y silo')"
      Call ExecSQL(DbMain, Q1)
      Q1 = "INSERT INTO TipoVehiculo (Codigo, Nombre) VALUES('10', 'Motor de semiremolque con grúa lateral para contenedor')"
      Call ExecSQL(DbMain, Q1)
      Q1 = "INSERT INTO TipoVehiculo (Codigo, Nombre) VALUES('11', 'Otro')"
      Call ExecSQL(DbMain, Q1)
        
      
      If lUpdOK Then
         lDbVer = 5
         Q1 = "UPDATE Param SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseAdm_V4 = lUpdOK

End Function
Public Function CorrigeBaseAdm_V3() As Boolean           'agregada 13 sep 2018
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
  

   On Error Resume Next
   
   
   '--------------------- Versión 3 -----------------------------------
   If lDbVer = 3 And lUpdOK = True Then
   
      Err.Clear
      
       'agregamos tabla Paises
      Set Tbl = New TableDef
      Tbl.Name = "Paises"
      
      Err.Clear
      Set fld = Tbl.CreateField("Id", dbLong)
      fld.Attributes = dbAutoIncrField ' Autonumber
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Paises.Id", vbExclamation
         lUpdOK = False
      End If
                        
      Err.Clear
      Set fld = Tbl.CreateField("Codigo", dbText, 4)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Paises.Codigo", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("Nombre", dbText, 100)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Paises.Nombre", vbExclamation
         lUpdOK = False
      End If
      
      DbMain.TableDefs.Append Tbl
      If Err = 0 Then
         DbMain.TableDefs.Refresh
         
         Q1 = "CREATE UNIQUE INDEX Idx ON Paises (Id) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)
         
         Q1 = "CREATE UNIQUE INDEX IdxCod ON Paises (Codigo)"
         Rc = ExecSQL(DbMain, Q1, False)
         
      ElseIf Err <> 3010 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Tabla Paises", vbExclamation
         lUpdOK = False

      End If
   
    
    
       'agregamos tabla Puertos
      Set Tbl = New TableDef
      Tbl.Name = "Puertos"
      
      Err.Clear
      Set fld = Tbl.CreateField("Id", dbLong)
      fld.Attributes = dbAutoIncrField ' Autonumber
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Puertos.Id", vbExclamation
         lUpdOK = False
      End If
                        
      Err.Clear
      Set fld = Tbl.CreateField("IdPais", dbLong)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Puertos.CodPais", vbExclamation
         lUpdOK = False
      End If
      
      Err.Clear
      Set fld = Tbl.CreateField("Codigo", dbText, 4)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Puertos.Codigo", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("Nombre", dbText, 100)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Puertos.Nombre", vbExclamation
         lUpdOK = False
      End If
      
      DbMain.TableDefs.Append Tbl
      If Err = 0 Then
         DbMain.TableDefs.Refresh
         
         Q1 = "CREATE UNIQUE INDEX Idx ON Puertos (Id) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)
         
         Q1 = "CREATE UNIQUE INDEX IdxCod ON Puertos (CodPais, Codigo)"
         Rc = ExecSQL(DbMain, Q1, False)
         
      ElseIf Err <> 3010 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Tabla Puertos", vbExclamation
         lUpdOK = False

      End If
   
      'agregamos tabla ClauCompraventa (Cláusula de Compraventa)
      Set Tbl = New TableDef
      Tbl.Name = "ClauCompraventa"
      
      Err.Clear
      Set fld = Tbl.CreateField("Id", dbLong)
      fld.Attributes = dbAutoIncrField ' Autonumber
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "ClauCompraventa.Id", vbExclamation
         lUpdOK = False
      End If
                        
      Err.Clear
      Set fld = Tbl.CreateField("Codigo", dbText, 2)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "ClauCompraventa.Codigo", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("Sigla", dbText, 5)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "ClauCompraventa.Sigla", vbExclamation
         lUpdOK = False
      End If
      
      Err.Clear
      Set fld = Tbl.CreateField("Nombre", dbText, 40)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "ClauCompraventa.Nombre", vbExclamation
         lUpdOK = False
      End If
      
      Err.Clear
      Set fld = Tbl.CreateField("EsFijo", dbBoolean)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "ClauCompraventa.EsFijo", vbExclamation
         lUpdOK = False
      End If
      
      DbMain.TableDefs.Append Tbl
      If Err = 0 Then
         DbMain.TableDefs.Refresh
         
         Q1 = "CREATE UNIQUE INDEX Idx ON ClauCompraVenta (Id) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)
         
         Q1 = "CREATE UNIQUE INDEX IdxCod ON ClauCompraVenta (Codigo, Sigla)"
         Rc = ExecSQL(DbMain, Q1, False)
         
      ElseIf Err <> 3010 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Tabla ClauCompraventa", vbExclamation
         lUpdOK = False

      End If
  
      'Llenamos la tabla ClauCompraVenta
      Q1 = "INSERT INTO ClauCompraventa( Codigo, Sigla, Nombre, EsFijo) VALUES( '1', 'CIF', 'Costos, Seguro Y Flete', -1 )"
      Call ExecSQL(DbMain, Q1)
      Q1 = "INSERT INTO ClauCompraventa( Codigo, Sigla, Nombre, EsFijo) VALUES( '2', 'CFR', 'Costos Y Flete', -1 )"
      Call ExecSQL(DbMain, Q1)
      Q1 = "INSERT INTO ClauCompraventa( Codigo, Sigla, Nombre, EsFijo) VALUES( '3', 'EXW', 'En Fábrica', -1 )"
      Call ExecSQL(DbMain, Q1)
      Q1 = "INSERT INTO ClauCompraventa( Codigo, Sigla, Nombre, EsFijo) VALUES( '4', 'FAS', 'Franco Al Costado Del Buque', -1 )"
      Call ExecSQL(DbMain, Q1)
      Q1 = "INSERT INTO ClauCompraventa( Codigo, Sigla, Nombre, EsFijo) VALUES( '5', 'FOB', 'Franco a Bordo', -1 )"
      Call ExecSQL(DbMain, Q1)
      Q1 = "INSERT INTO ClauCompraventa( Codigo, Sigla, Nombre, EsFijo) VALUES( '6', 'S/CL', 'Sin Cláusula De Compraventa', -1 )"
      Call ExecSQL(DbMain, Q1)
      Q1 = "INSERT INTO ClauCompraventa( Codigo, Sigla, Nombre, EsFijo) VALUES( '9', 'DDP', 'Entregadas Derechos Pagados', -1 )"
      Call ExecSQL(DbMain, Q1)
      Q1 = "INSERT INTO ClauCompraventa( Codigo, Sigla, Nombre, EsFijo) VALUES( '10', 'FCA', 'Franco Transportista', -1 )"
      Call ExecSQL(DbMain, Q1)
      Q1 = "INSERT INTO ClauCompraventa( Codigo, Sigla, Nombre, EsFijo) VALUES( '11', 'CPT', 'Transporte Pagado Hasta', -1 )"
      Call ExecSQL(DbMain, Q1)
      Q1 = "INSERT INTO ClauCompraventa( Codigo, Sigla, Nombre, EsFijo) VALUES( '12', 'CIP', 'Transporte y Seguro Pagado Hasta', -1 )"
      Call ExecSQL(DbMain, Q1)
      Q1 = "INSERT INTO ClauCompraventa( Codigo, Sigla, Nombre, EsFijo) VALUES( '17', 'DAT', 'Entregadas En Puerto Destino', -1 )"
      Call ExecSQL(DbMain, Q1)
      Q1 = "INSERT INTO ClauCompraventa( Codigo, Sigla, Nombre, EsFijo) VALUES( '18', 'DAP', 'Entregadas En Lugar Convenido', -1 )"
      Call ExecSQL(DbMain, Q1)
 
    
      'Actualizamos tabla de Monedas
      Call AlterField(DbMain, "Monedas", "Descrip", dbText, 100)
      
      'agregamos campo CodAduana a Monedas
      Set Tbl = DbMain.TableDefs("Monedas")

      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("CodAduana", dbText, 3)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "Monedas.CodAduana", vbExclamation
         lUpdOK = False
      End If
      Err.Clear
      
      'agregamos campo EsFijo a Monedas
      Tbl.Fields.Append Tbl.CreateField("EsFijo", dbBoolean)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "Monedas.EsFijo", vbExclamation
         lUpdOK = False
      End If
      
      Q1 = "CREATE UNIQUE INDEX IdxCod ON Monedas (CodAduana)"
      Rc = ExecSQL(DbMain, Q1, False)
         
      Q1 = "UPDATE Monedas SET CodAduana = '200', EsFijo = -1 WHERE Descrip = 'Pesos'"
      Call ExecSQL(DbMain, Q1)

      Q1 = "UPDATE Monedas SET CodAduana = '013', Descrip = 'Dólar USA', EsFijo = -1  WHERE Descrip = 'Dólar'"
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Monedas SET EsFijo = -1  WHERE Simbolo = 'UF' OR Simbolo = 'UTM'"
      Call ExecSQL(DbMain, Q1)
    
      If lUpdOK Then
         lDbVer = 4
         Q1 = "UPDATE Param SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseAdm_V3 = lUpdOK

End Function
Public Function CorrigeBaseAdm_V2() As Boolean
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
  

   On Error Resume Next
   
   
   '--------------------- Versión 2 -----------------------------------
   If lDbVer = 2 And lUpdOK = True Then
   
      Err.Clear
      
      'agregamos Tasa a ILA anacl con alto azucar
      Call AlterField(DbMain, "TipoValor", "CodSIIDTE", dbText, 3)
      
      Q1 = "UPDATE TipoValor SET Tasa = 18, CodSIIDTE = '271' WHERE TipoLib = " & LIB_VENTAS & " AND Codigo = " & LIBVENTAS_ILABEDANALCAZUCAR
      Call ExecSQL(DbMain, Q1)
    
      If lUpdOK Then
         lDbVer = 3
         Q1 = "UPDATE Param SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseAdm_V2 = lUpdOK

End Function

Public Function CorrigeBaseAdm_V1() As Boolean
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
  

   On Error Resume Next
   
   lUpdOK = True

   Q1 = "SELECT Valor FROM Param WHERE Tipo = 'DBVER'"
   Set Rs = OpenRs(DbMain, Q1)
   If Rs Is Nothing Then
      MsgBox1 "La base de datos está corrupta o es muy antigua.", vbCritical
      Call CloseDb(DbMain)
      End
   End If
      
   If Rs.EOF Then
      Call CloseRs(Rs)
      Q1 = "INSERT INTO Param (Tipo, Codigo, Valor) VALUES ('DBVER', 0, '0')"
      Call ExecSQL(DbMain, Q1)
      lDbVer = 1
   Else
      lDbVer = Val(vFld(Rs("Valor")))
      If lDbVer = 0 Then
         lDbVer = 1
      End If
   End If

   Call CloseRs(Rs)
  
   
   '--------------------- Versión 1 -----------------------------------
   If lDbVer = 1 And lUpdOK = True Then
   
      Err.Clear
      
      Q1 = "CREATE TABLE Monedas (IdMoneda integer, Descrip char(50), Simbolo char(10), DecInf Single, DecVenta Single, Caracteristica integer)"
      Rc = ExecSQL(DbMain, Q1)
      
      Q1 = "CREATE UNIQUE INDEX IdxMoneda ON Monedas (IdMoneda) WITH PRIMARY"
      Rc = ExecSQL(DbMain, Q1)
              
      Q1 = "INSERT INTO Monedas (IdMoneda, Descrip, Simbolo, DecInf, DecVenta, Caracteristica ) "
      Q1 = Q1 & " VALUES (1, 'Pesos', '$', 0, 0, 0 )"
      Rc = ExecSQL(DbMain, Q1)
              
      Q1 = "INSERT INTO Monedas (IdMoneda, Descrip, Simbolo, DecInf, DecVenta, Caracteristica ) "
      Q1 = Q1 & " VALUES (2, 'Dólar', 'US$', 2, 0, 3 )"
      Rc = ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO Monedas (IdMoneda, Descrip, Simbolo, DecInf, DecVenta, Caracteristica ) "
      Q1 = Q1 & " VALUES (3, 'Unidad de Fomento', 'UF', 3, 2, 3 )"
      Rc = ExecSQL(DbMain, Q1)
      
      Q1 = "INSERT INTO Monedas (IdMoneda, Descrip, Simbolo, DecInf, DecVenta, Caracteristica ) "
      Q1 = Q1 & " VALUES (4, 'Unidad Tributaria Mensual', 'UTM', 3, 0, 2 )"
      Rc = ExecSQL(DbMain, Q1)
     
      Q1 = "CREATE TABLE Equivalencia (IdMoneda long, Fecha long, Valor double)"
      Rc = ExecSQL(DbMain, Q1)
      
      Q1 = "CREATE INDEX IdxEquiv ON Equivalencia (IdMoneda)"
      Rc = ExecSQL(DbMain, Q1)
     
     
     
      If lUpdOK Then
         lDbVer = 2
         Q1 = "UPDATE Param SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBaseAdm_V1 = lUpdOK

End Function

Private Sub UpdateCodActiv2018()
   Dim Q1 As String
   
   'eliminamos los registros de la versión 1 y 2 ya que ahora no se usan
   Q1 = "DELETE * FROM CodActiv WHERE Version < 3"
   Call ExecSQL(DbMain, Q1)
   
   'insertamos los nuevos registros con versión 3
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011101', 'Cultivo de trigo', 3, '011111')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011102', 'Cultivo de maíz', 3, '011112')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011103', 'Cultivo de avena', 3, '011113')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011200', 'Cultivo de arroz', 3, '011114')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011104', 'Cultivo de cebada', 3, '011115')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011105', 'Cultivo de otros cereales (excepto trigo, maíz, avena y cebada)', 3, '011119')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011902', 'Cultivos forrajeros en praderas mejoradas o sembradas; cultivos suplementarios forrajeros', 3, '011121')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011902', 'Cultivos forrajeros en praderas mejoradas o sembradas; cultivos suplementarios forrajeros', 3, '011122')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011106', 'Cultivo de porotos', 3, '011131')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011107', 'Cultivo de lupino', 3, '011132')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011108', 'Cultivo de otras legumbres (excepto porotos y lupino)', 3, '011139')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011301', 'Cultivo de papas', 3, '011141')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011302', 'Cultivo de camotes', 3, '011142')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011303', 'Cultivo de otros tubérculos (excepto papas y camotes)', 3, '011149')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011109', 'Cultivo de semillas de raps', 3, '011151')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011110', 'Cultivo de semillas de maravilla (girasol)', 3, '011152')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012600', 'Cultivo de frutos oleaginosos (incluye el cultivo de aceitunas)', 3, '011159')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011111', 'Cultivo de semillas de cereales, legumbres y oleaginosas (excepto semillas de raps y maravilla)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011111', 'Cultivo de semillas de cereales, legumbres y oleaginosas (excepto semillas de raps y maravilla)', 3, '011160')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011304', 'Cultivo de remolacha azucarera', 3, '011191')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011500', 'Cultivo de tabaco', 3, '011192')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('016300', 'Actividades poscosecha', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011600', 'Cultivo de plantas de fibra', 3, '011193')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012900', 'Cultivo de otras plantas perennes', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012802', 'Cultivo de plantas aromáticas, medicinales y farmacéuticas', 3, '011194')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012900', 'Cultivo de otras plantas perennes', 3, '011199')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011400', 'Cultivo de caña de azúcar', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012802', 'Cultivo de plantas aromáticas, medicinales y farmacéuticas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011306', 'Cultivo de hortalizas y melones', 3, '011211')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012600', 'Cultivo de frutos oleaginosos (incluye el cultivo de aceitunas)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012802', 'Cultivo de plantas aromáticas, medicinales y farmacéuticas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011306', 'Cultivo de hortalizas y melones', 3, '011212')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011306', 'Cultivo de hortalizas y melones', 3, '011213')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011901', 'Cultivo de flores', 3, '011220')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('013000', 'Cultivo de plantas vivas incluida la producción en viveros (excepto viveros forestales)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011305', 'Cultivo de semillas de hortalizas', 3, '011230')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011903', 'Cultivos de semillas de flores; cultivo de semillas de plantas forrajeras', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012501', 'Cultivo de semillas de frutas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('013000', 'Cultivo de plantas vivas incluida la producción en viveros (excepto viveros forestales)', 3, '011240')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('011306', 'Cultivo de hortalizas y melones', 3, '011250')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012900', 'Cultivo de otras plantas perennes', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('023000', 'Recolección de productos forestales distintos de la madera', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012111', 'Cultivo de uva destinada a la producción de pisco y aguardiente', 3, '011311')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012112', 'Cultivo de uva destinada a la producción de vino', 3, '011312')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('110200', 'Elaboración de vinos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012120', 'Cultivo de uva para mesa', 3, '011313')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012400', 'Cultivo de frutas de pepita y de hueso', 3, '011321')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012200', 'Cultivo de frutas tropicales y subtropicales (incluye el cultivo de paltas)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012300', 'Cultivo de cítricos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012502', 'Cultivo de otros frutos y nueces de árboles y arbustos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012600', 'Cultivo de frutos oleaginosos (incluye el cultivo de aceitunas)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012502', 'Cultivo de otros frutos y nueces de árboles y arbustos', 3, '011322')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('023000', 'Recolección de productos forestales distintos de la madera', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012700', 'Cultivo de plantas con las que se preparan bebidas (incluye el cultivo de café, té y mate)', 3, '011330')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012801', 'Cultivo de especias', 3, '011340')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('014101', 'Cría de ganado bovino para la producción lechera', 3, '012111')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('014102', 'Cría de ganado bovino para la producción de carne o como ganado reproductor', 3, '012112')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('014410', 'Cría de ovejas (ovinos)', 3, '012120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('014200', 'Cría de caballos y otros equinos', 3, '012130')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('014500', 'Cría de cerdos', 3, '012210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('014601', 'Cría de aves de corral para la producción de carne', 3, '012221')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('014602', 'Cría de aves de corral para la producción de huevos', 3, '012222')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('014909', 'Cría de otros animales n.c.p.', 3, '012223')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('014909', 'Cría de otros animales n.c.p.', 3, '012230')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('014901', 'Apicultura', 3, '012240')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('014909', 'Cría de otros animales n.c.p.', 3, '012250')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('032200', 'Acuicultura de agua dulce', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('014909', 'Cría de otros animales n.c.p.', 3, '012290')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('014300', 'Cría de llamas, alpacas, vicuñas, guanacos y otros camélidos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('014420', 'Cría de cabras (caprinos)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('032130', 'Reproducción y cría de moluscos, crustáceos y gusanos marinos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('032200', 'Acuicultura de agua dulce', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('015000', 'Cultivo de productos agrícolas en combinación con la cría de animales (explotación mixta)', 3, '013000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('016300', 'Actividades poscosecha', 3, '014011')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('016300', 'Actividades poscosecha', 3, '014012')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('016400', 'Tratamiento de semillas para propagación', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('016100', 'Actividades de apoyo a la agricultura', 3, '014013')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('016100', 'Actividades de apoyo a la agricultura', 3, '014014')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('016100', 'Actividades de apoyo a la agricultura', 3, '014015')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('016100', 'Actividades de apoyo a la agricultura', 3, '014019')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('813000', 'Actividades de paisajismo, servicios de jardinería y servicios conexos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('960901', 'Servicios de adiestramiento, guardería, peluquería, paseo de mascotas (excepto act. veterinarias)', 3, '014021')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('016200', 'Actividades de apoyo a la ganadería', 3, '014022')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('017000', 'Caza ordinaria y mediante trampas y actividades de servicios conexas', 3, '015010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('017000', 'Caza ordinaria y mediante trampas y actividades de servicios conexas', 3, '015090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('949909', 'Actividades de otras asociaciones n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('022000', 'Extracción de madera', 3, '020010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('021002', 'Silvicultura y otras actividades forestales (excepto explotación de viveros forestales)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('023000', 'Recolección de productos forestales distintos de la madera', 3, '020020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('021001', 'Explotación de viveros forestales', 3, '020030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('012900', 'Cultivo de otras plantas perennes', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('024001', 'Servicios de forestación a cambio de una retribución o por contrata', 3, '020041')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('024002', 'Servicios de corta de madera a cambio de una retribución o por contrata', 3, '020042')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('024003', 'Servicios de extinción y prevención de incendios forestales', 3, '020043')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('024009', 'Otros servicios de apoyo a la silvicultura n.c.p.', 3, '020049')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('032200', 'Acuicultura de agua dulce', 3, '051010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('032110', 'Cultivo y crianza de peces marinos', 3, '051020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('032120', 'Cultivo, reproducción y manejo de algas marinas', 3, '051030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('032130', 'Reproducción y cría de moluscos, crustáceos y gusanos marinos', 3, '051040')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('032140', 'Servicios relacionados con la acuicultura marina', 3, '051090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('032200', 'Acuicultura de agua dulce', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('031110', 'Pesca marítima industrial, excepto de barcos factoría', 3, '052010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('102050', 'Actividades de elaboración y conservación de pescado, realizadas en barcos factoría', 3, '052020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('031120', 'Pesca marítima artesanal', 3, '052030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('031200', 'Pesca de agua dulce', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('031130', 'Recolección y extracción de productos marinos', 3, '052040')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('031140', 'Servicios relacionados con la pesca marítima', 3, '052050')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('031200', 'Pesca de agua dulce', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('051000', 'Extracción de carbón de piedra', 3, '100000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('052000', 'Extracción de lignito', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('089200', 'Extracción de turba', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('192000', 'Fabricación de productos de la refinación del petróleo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('061000', 'Extracción de petróleo crudo', 3, '111000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('062000', 'Extracción de gas natural', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('091001', 'Actividades de apoyo para la extracción de petróleo y gas natural prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('091001', 'Actividades de apoyo para la extracción de petróleo y gas natural prestados por empresas', 3, '112000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('072100', 'Extracción de minerales de uranio y torio', 3, '120000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('071000', 'Extracción de minerales de hierro', 3, '131000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('072910', 'Extracción de oro y plata', 3, '132010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('072991', 'Extracción de zinc y plomo', 3, '132020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('072992', 'Extracción de manganeso', 3, '132030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('072999', 'Extracción de otros minerales metalíferos no ferrosos n.c.p. (excepto zinc, plomo y manganeso)', 3, '132090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('040000', 'Extracción y procesamiento de cobre', 3, '133000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('081000', 'Extracción de piedra, arena y arcilla', 3, '141000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('089190', 'Extracción de minerales para la fabricación de abonos y productos químicos n.c.p.', 3, '142100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('089300', 'Extracción de sal', 3, '142200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('089110', 'Extracción y procesamiento de litio', 3, '142300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('089190', 'Extracción de minerales para la fabricación de abonos y productos químicos n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('089900', 'Explotación de otras minas y canteras n.c.p.', 3, '142900')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('081000', 'Extracción de piedra, arena y arcilla', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('101011', 'Explotación de mataderos de bovinos, ovinos, equinos, caprinos, porcinos y camélidos', 3, '151110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('521001', 'Explotación de frigoríficos para almacenamiento y depósito', 3, '151120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('101019', 'Explotación de mataderos de aves y de otros tipos de animales n.c.p.', 3, '151130')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('101020', 'Elaboración y conservación de carne y productos cárnicos', 3, '151140')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('102010', 'Producción de harina de pescado', 3, '151210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('102030', 'Elaboración y conservación de otros pescados, en plantas en tierra (excepto barcos factoría)', 3, '151221')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('102020', 'Elaboración y conservación de salmónidos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('102040', 'Elaboración y conservación de crustáceos, moluscos y otros productos acuáticos, en plantas en tierra', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107500', 'Elaboración de comidas y platos preparados envasados, rotulados y con información nutricional', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('102020', 'Elaboración y conservación de salmónidos', 3, '151222')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('102030', 'Elaboración y conservación de otros pescados, en plantas en tierra (excepto barcos factoría)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('102040', 'Elaboración y conservación de crustáceos, moluscos y otros productos acuáticos, en plantas en tierra', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107500', 'Elaboración de comidas y platos preparados envasados, rotulados y con información nutricional', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('102020', 'Elaboración y conservación de salmónidos', 3, '151223')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('102030', 'Elaboración y conservación de otros pescados, en plantas en tierra (excepto barcos factoría)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('102040', 'Elaboración y conservación de crustáceos, moluscos y otros productos acuáticos, en plantas en tierra', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('102060', 'Elaboración y procesamiento de algas', 3, '151230')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('103000', 'Elaboración y conservación de frutas, legumbres y hortalizas', 3, '151300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107500', 'Elaboración de comidas y platos preparados envasados, rotulados y con información nutricional', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('104000', 'Elaboración de aceites y grasas de origen vegetal y animal (excepto elaboración de mantequilla)', 3, '151410')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('104000', 'Elaboración de aceites y grasas de origen vegetal y animal (excepto elaboración de mantequilla)', 3, '151420')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('104000', 'Elaboración de aceites y grasas de origen vegetal y animal (excepto elaboración de mantequilla)', 3, '151430')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('105000', 'Elaboración de productos lácteos', 3, '152010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('105000', 'Elaboración de productos lácteos', 3, '152020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('105000', 'Elaboración de productos lácteos', 3, '152030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('106101', 'Molienda de trigo: producción de harina, sémola y gránulos', 3, '153110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('106102', 'Molienda de arroz; producción de harina de arroz', 3, '153120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('106109', 'Elaboración de otros productos de molinería n.c.p.', 3, '153190')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('106200', 'Elaboración de almidones y productos derivados del almidón', 3, '153210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('106200', 'Elaboración de almidones y productos derivados del almidón', 3, '153220')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('108000', 'Elaboración de piensos preparados para animales', 3, '153300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107100', 'Elaboración de productos de panadería y pastelería', 3, '154110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107100', 'Elaboración de productos de panadería y pastelería', 3, '154120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107200', 'Elaboración de azúcar', 3, '154200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107300', 'Elaboración de cacao, chocolate y de productos de confitería', 3, '154310')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107300', 'Elaboración de cacao, chocolate y de productos de confitería', 3, '154320')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107400', 'Elaboración de macarrones, fideos, alcuzcuz y productos farináceos similares', 3, '154400')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107500', 'Elaboración de comidas y platos preparados envasados, rotulados y con información nutricional', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107901', 'Elaboración de té, café, mate e infusiones de hierbas', 3, '154910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107902', 'Elaboración de levaduras naturales o artificiales', 3, '154920')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107903', 'Elaboración de vinagres, mostazas, mayonesas y condimentos en general', 3, '154930')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107909', 'Elaboración de otros productos alimenticios n.c.p.', 3, '154990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('103000', 'Elaboración y conservación de frutas, legumbres y hortalizas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107100', 'Elaboración de productos de panadería y pastelería', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107500', 'Elaboración de comidas y platos preparados envasados, rotulados y con información nutricional', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('110110', 'Elaboración de pisco (industrias pisqueras)', 3, '155110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('110120', 'Destilación, rectificación y mezclas de bebidas alcohólicas; excepto pisco', 3, '155120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('201109', 'Fabricación de otras sustancias químicas básicas n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('110200', 'Elaboración de vinos', 3, '155200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('110300', 'Elaboración de bebidas malteadas y de malta', 3, '155300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('110401', 'Elaboración de bebidas no alcohólicas', 3, '155410')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('110402', 'Producción de aguas minerales y otras aguas embotelladas', 3, '155420')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('353002', 'Elaboración de hielo (excepto fabricación de hielo seco)', 3, '155430')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('120001', 'Elaboración de cigarros y cigarrillos', 3, '160010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('120009', 'Elaboración de otros productos de tabaco n.c.p.', 3, '160090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('131200', 'Tejedura de productos textiles', 3, '171100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('131100', 'Preparación e hilatura de fibras textiles', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('131300', 'Acabado de productos textiles', 3, '171200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('952900', 'Reparación de otros efectos personales y enseres domésticos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('139200', 'Fabricación de artículos confeccionados de materiales textiles, excepto prendas de vestir', 3, '172100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('325009', 'Fabricación de instrumentos y materiales médicos, oftalmológicos y odontológicos n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331900', 'Reparación de otros tipos de equipo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('139300', 'Fabricación de tapices y alfombras', 3, '172200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('139400', 'Fabricación de cuerdas, cordeles, bramantes y redes', 3, '172300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331900', 'Reparación de otros tipos de equipo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('139900', 'Fabricación de otros productos textiles n.c.p.', 3, '172910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('131300', 'Acabado de productos textiles', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('139900', 'Fabricación de otros productos textiles n.c.p.', 3, '172990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('170900', 'Fabricación de otros artículos de papel y cartón', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('143000', 'Fabricación de artículos de punto y ganchillo', 3, '173000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('139100', 'Fabricación de tejidos de punto y ganchillo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('141001', 'Fabricación de prendas de vestir de materiales textiles y similares', 3, '181010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('131300', 'Acabado de productos textiles', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('141002', 'Fabricación de prendas de vestir de cuero natural o artificial', 3, '181020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('141003', 'Fabricación de accesorios de vestir', 3, '181030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('141004', 'Fabricación de ropa de trabajo', 3, '181040')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('131300', 'Acabado de productos textiles', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('329000', 'Otras industrias manufactureras n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('151100', 'Curtido y adobo de cueros; adobo y teñido de pieles', 3, '182000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('142000', 'Fabricación de artículos de piel', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('151100', 'Curtido y adobo de cueros; adobo y teñido de pieles', 3, '191100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('151200', 'Fabricación de maletas, bolsos y artículos similares, artículos de talabartería y guarnicionería', 3, '191200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('329000', 'Otras industrias manufactureras n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('152000', 'Fabricación de calzado', 3, '192000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('151200', 'Fabricación de maletas, bolsos y artículos similares, artículos de talabartería y guarnicionería', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('162900', 'Fabricación de otros productos de madera, de artículos de corcho, paja y materiales trenzables', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('221900', 'Fabricación de otros productos de caucho', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('222000', 'Fabricación de productos de plástico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('323000', 'Fabricación de artículos de deporte', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('161000', 'Aserrado y acepilladura de madera', 3, '201000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('162100', 'Fabricación de hojas de madera para enchapado y tableros a base de madera', 3, '202100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('162200', 'Fabricación de partes y piezas de carpintería para edificios y construcciones', 3, '202200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('162300', 'Fabricación de recipientes de madera', 3, '202300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331900', 'Reparación de otros tipos de equipo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('162900', 'Fabricación de otros productos de madera, de artículos de corcho, paja y materiales trenzables', 3, '202900')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('329000', 'Otras industrias manufactureras n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331900', 'Reparación de otros tipos de equipo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('170110', 'Fabricación de celulosa y otras pastas de madera', 3, '210110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('170190', 'Fabricación de papel y cartón para su posterior uso industrial n.c.p.', 3, '210121')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('170190', 'Fabricación de papel y cartón para su posterior uso industrial n.c.p.', 3, '210129')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('170200', 'Fabricación de papel y cartón ondulado y de envases de papel y cartón', 3, '210200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('170900', 'Fabricación de otros artículos de papel y cartón', 3, '210900')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('181109', 'Otras actividades de impresión n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('222000', 'Fabricación de productos de plástico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('581100', 'Edición de libros', 3, '221101')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('581200', 'Edición de directorios y listas de correo', 3, '221109')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('329000', 'Otras industrias manufactureras n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('592000', 'Actividades de grabación de sonido y edición de música', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('581300', 'Edición de diarios, revistas y otras publicaciones periódicas', 3, '221200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('592000', 'Actividades de grabación de sonido y edición de música', 3, '221300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('581900', 'Otras actividades de edición', 3, '221900')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('581300', 'Edición de diarios, revistas y otras publicaciones periódicas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('181101', 'Impresión de libros', 3, '222101')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('181109', 'Otras actividades de impresión n.c.p.', 3, '222109')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('170900', 'Fabricación de otros artículos de papel y cartón', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('181200', 'Actividades de servicios relacionadas con la impresión', 3, '222200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('182000', 'Reproducción de grabaciones', 3, '223000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('191000', 'Fabricación de productos de hornos de coque', 3, '231000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('192000', 'Fabricación de productos de la refinación del petróleo', 3, '232000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('201109', 'Fabricación de otras sustancias químicas básicas n.c.p.', 3, '233000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('210000', 'Fabricación de productos farmacéuticos, sustancias químicas medicinales y productos botánicos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('242009', 'Fabricación de productos primarios de metales preciosos y de otros metales no ferrosos n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('381200', 'Recogida de desechos peligrosos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('382200', 'Tratamiento y eliminación de desechos peligrosos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('201101', 'Fabricación de carbón vegetal (excepto activado); fabricación de briquetas de carbón vegetal', 3, '241110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('201109', 'Fabricación de otras sustancias químicas básicas n.c.p.', 3, '241190')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('191000', 'Fabricación de productos de hornos de coque', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('201200', 'Fabricación de abonos y compuestos de nitrógeno', 3, '241200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('382100', 'Tratamiento y eliminación de desechos no peligrosos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('201300', 'Fabricación de plásticos y caucho sintético en formas primarias', 3, '241300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('202100', 'Fabricación de plaguicidas y otros productos químicos de uso agropecuario', 3, '242100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('202200', 'Fabricación de pinturas, barnices y productos de revestimiento, tintas de imprenta y masillas', 3, '242200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('210000', 'Fabricación de productos farmacéuticos, sustancias químicas medicinales y productos botánicos', 3, '242300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('325009', 'Fabricación de instrumentos y materiales médicos, oftalmológicos y odontológicos n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('202300', 'Fabricación de jabones y detergentes, preparados para limpiar, perfumes y preparados de tocador', 3, '242400')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('202901', 'Fabricación de explosivos y productos pirotécnicos', 3, '242910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('202909', 'Fabricación de otros productos químicos n.c.p.', 3, '242990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('107909', 'Elaboración de otros productos alimenticios n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('201109', 'Fabricación de otras sustancias químicas básicas n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('261000', 'Fabricación de componentes y tableros electrónicos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('268000', 'Fabricación de soportes magnéticos y ópticos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281700', 'Fabricación de maquinaria y equipo de oficina (excepto computadores y equipo periférico)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('203000', 'Fabricación de fibras artificiales', 3, '243000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('221100', 'Fabricación de cubiertas y cámaras de caucho; recauchutado y renovación de cubiertas de caucho', 3, '251110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('221100', 'Fabricación de cubiertas y cámaras de caucho; recauchutado y renovación de cubiertas de caucho', 3, '251120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('221900', 'Fabricación de otros productos de caucho', 3, '251900')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('222000', 'Fabricación de productos de plástico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281200', 'Fabricación de equipo de propulsión de fluidos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('329000', 'Otras industrias manufactureras n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331900', 'Reparación de otros tipos de equipo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('222000', 'Fabricación de productos de plástico', 3, '252010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('222000', 'Fabricación de productos de plástico', 3, '252020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('222000', 'Fabricación de productos de plástico', 3, '252090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('261000', 'Fabricación de componentes y tableros electrónicos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('273300', 'Fabricación de dispositivos de cableado', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('329000', 'Otras industrias manufactureras n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331400', 'Reparación de equipo eléctrico (excepto reparación de equipo y enseres domésticos)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331900', 'Reparación de otros tipos de equipo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('231001', 'Fabricación de vidrio plano', 3, '261010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('231002', 'Fabricación de vidrio hueco', 3, '261020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('231003', 'Fabricación de fibras de vidrio', 3, '261030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('231009', 'Fabricación de productos de vidrio n.c.p.', 3, '261090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331900', 'Reparación de otros tipos de equipo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('239300', 'Fabricación de otros productos de porcelana y de cerámica', 3, '269101')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('239300', 'Fabricación de otros productos de porcelana y de cerámica', 3, '269109')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('239200', 'Fabricación de materiales de construcción de arcilla', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('239100', 'Fabricación de productos refractarios', 3, '269200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('239200', 'Fabricación de materiales de construcción de arcilla', 3, '269300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('239400', 'Fabricación de cemento, cal y yeso', 3, '269400')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('239500', 'Fabricación de artículos de hormigón, cemento y yeso', 3, '269510')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('239500', 'Fabricación de artículos de hormigón, cemento y yeso', 3, '269520')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('239500', 'Fabricación de artículos de hormigón, cemento y yeso', 3, '269530')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('239500', 'Fabricación de artículos de hormigón, cemento y yeso', 3, '269590')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('239600', 'Corte, talla y acabado de la piedra', 3, '269600')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('239900', 'Fabricación de otros productos minerales no metálicos n.c.p.', 3, '269910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('239900', 'Fabricación de otros productos minerales no metálicos n.c.p.', 3, '269990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('131200', 'Tejedura de productos textiles', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('329000', 'Otras industrias manufactureras n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331900', 'Reparación de otros tipos de equipo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('241000', 'Industrias básicas de hierro y acero', 3, '271000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('243100', 'Fundición de hierro y acero', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('242001', 'Fabricación de productos primarios de cobre', 3, '272010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('242002', 'Fabricación de productos primarios de aluminio', 3, '272020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('242009', 'Fabricación de productos primarios de metales preciosos y de otros metales no ferrosos n.c.p.', 3, '272090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('243100', 'Fundición de hierro y acero', 3, '273100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('243200', 'Fundición de metales no ferrosos', 3, '273200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('251100', 'Fabricación de productos metálicos para uso estructural', 3, '281100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331100', 'Reparación de productos elaborados de metal', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('251201', 'Fabricación de recipientes de metal para gases comprimidos o licuados', 3, '281211')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('251209', 'Fabricación de tanques, depósitos y recipientes de metal n.c.p.', 3, '281219')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331100', 'Reparación de productos elaborados de metal', 3, '281280')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('251300', 'Fabricación de generadores de vapor, excepto calderas de agua caliente para calefacción central', 3, '281310')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331100', 'Reparación de productos elaborados de metal', 3, '281380')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('259100', 'Forja, prensado, estampado y laminado de metales; pulvimetalurgia', 3, '289100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('259200', 'Tratamiento y revestimiento de metales; maquinado', 3, '289200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('016200', 'Actividades de apoyo a la ganadería', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('181109', 'Otras actividades de impresión n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331100', 'Reparación de productos elaborados de metal', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('952900', 'Reparación de otros efectos personales y enseres domésticos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('259300', 'Fabricación de artículos de cuchillería, herramientas de mano y artículos de ferretería', 3, '289310')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331100', 'Reparación de productos elaborados de metal', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('259300', 'Fabricación de artículos de cuchillería, herramientas de mano y artículos de ferretería', 3, '289320')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331100', 'Reparación de productos elaborados de metal', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('259900', 'Fabricación de otros productos elaborados de metal n.c.p.', 3, '289910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('259900', 'Fabricación de otros productos elaborados de metal n.c.p.', 3, '289990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281700', 'Fabricación de maquinaria y equipo de oficina (excepto computadores y equipo periférico)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('329000', 'Otras industrias manufactureras n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331100', 'Reparación de productos elaborados de metal', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281100', 'Fabricación de motores y turbinas, excepto para aeronaves, vehículos automotores y motocicletas', 3, '291110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331209', 'Reparación de otro tipo de maquinaria y equipos industriales n.c.p.', 3, '291180')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281300', 'Fabricación de otras bombas, compresores, grifos y válvulas', 3, '291210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281100', 'Fabricación de motores y turbinas, excepto para aeronaves, vehículos automotores y motocicletas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281200', 'Fabricación de equipo de propulsión de fluidos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331209', 'Reparación de otro tipo de maquinaria y equipos industriales n.c.p.', 3, '291280')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281400', 'Fabricación de cojinetes, engranajes, trenes de engranajes y piezas de transmisión', 3, '291310')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281200', 'Fabricación de equipo de propulsión de fluidos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331209', 'Reparación de otro tipo de maquinaria y equipos industriales n.c.p.', 3, '291380')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281500', 'Fabricación de hornos, calderas y quemadores', 3, '291410')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331209', 'Reparación de otro tipo de maquinaria y equipos industriales n.c.p.', 3, '291480')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281600', 'Fabricación de equipo de elevación y manipulación', 3, '291510')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331209', 'Reparación de otro tipo de maquinaria y equipos industriales n.c.p.', 3, '291580')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281900', 'Fabricación de otros tipos de maquinaria de uso general', 3, '291910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('325009', 'Fabricación de instrumentos y materiales médicos, oftalmológicos y odontológicos n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331209', 'Reparación de otro tipo de maquinaria y equipos industriales n.c.p.', 3, '291980')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282100', 'Fabricación de maquinaria agropecuaria y forestal', 3, '292110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331201', 'Reparación de maquinaria agropecuaria y forestal', 3, '292180')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282200', 'Fabricación de maquinaria para la conformación de metales y de máquinas herramienta', 3, '292210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('279000', 'Fabricación de otros tipos de equipo eléctrico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281800', 'Fabricación de herramientas de mano motorizadas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281900', 'Fabricación de otros tipos de maquinaria de uso general', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331209', 'Reparación de otro tipo de maquinaria y equipos industriales n.c.p.', 3, '292280')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331400', 'Reparación de equipo eléctrico (excepto reparación de equipo y enseres domésticos)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282300', 'Fabricación de maquinaria metalúrgica', 3, '292310')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331202', 'Reparación de maquinaria metalúrgica, para la minería, extracción de petróleo y para la construcción', 3, '292380')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282400', 'Fabricación de maquinaria para la explotación de minas y canteras y para obras de construcción', 3, '292411')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282400', 'Fabricación de maquinaria para la explotación de minas y canteras y para obras de construcción', 3, '292412')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331202', 'Reparación de maquinaria metalúrgica, para la minería, extracción de petróleo y para la construcción', 3, '292480')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282500', 'Fabricación de maquinaria para la elaboración de alimentos, bebidas y tabaco', 3, '292510')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331203', 'Reparación de maquinaria para la elaboración de alimentos, bebidas y tabaco', 3, '292580')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282600', 'Fabricación de maquinaria para la elaboración de productos textiles, prendas de vestir y cueros', 3, '292610')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331204', 'Reparación de maquinaria para producir textiles, prendas de vestir, artículos de cuero y calzado', 3, '292680')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('252000', 'Fabricación de armas y municiones', 3, '292710')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('303000', 'Fabricación de aeronaves, naves espaciales y maquinaria conexa', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('304000', 'Fabricación de vehículos militares de combate', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331100', 'Reparación de productos elaborados de metal', 3, '292780')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282900', 'Fabricación de otros tipos de maquinaria de uso especial', 3, '292910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('259300', 'Fabricación de artículos de cuchillería, herramientas de mano y artículos de ferretería', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('279000', 'Fabricación de otros tipos de equipo eléctrico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282600', 'Fabricación de maquinaria para la elaboración de productos textiles, prendas de vestir y cueros', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331209', 'Reparación de otro tipo de maquinaria y equipos industriales n.c.p.', 3, '292980')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331100', 'Reparación de productos elaborados de metal', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331400', 'Reparación de equipo eléctrico (excepto reparación de equipo y enseres domésticos)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('275000', 'Fabricación de aparatos de uso doméstico', 3, '293000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281500', 'Fabricación de hornos, calderas y quemadores', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281900', 'Fabricación de otros tipos de maquinaria de uso general', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('262000', 'Fabricación de computadores y equipo periférico', 3, '300010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('261000', 'Fabricación de componentes y tableros electrónicos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281700', 'Fabricación de maquinaria y equipo de oficina (excepto computadores y equipo periférico)', 3, '300020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('271000', 'Fabricación de motores, generadores y transformadores eléctricos, aparatos de distribución y control', 3, '311010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('261000', 'Fabricación de componentes y tableros electrónicos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('279000', 'Fabricación de otros tipos de equipo eléctrico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281100', 'Fabricación de motores y turbinas, excepto para aeronaves, vehículos automotores y motocicletas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331400', 'Reparación de equipo eléctrico (excepto reparación de equipo y enseres domésticos)', 3, '311080')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331209', 'Reparación de otro tipo de maquinaria y equipos industriales n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('271000', 'Fabricación de motores, generadores y transformadores eléctricos, aparatos de distribución y control', 3, '312010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('261000', 'Fabricación de componentes y tableros electrónicos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('273300', 'Fabricación de dispositivos de cableado', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('279000', 'Fabricación de otros tipos de equipo eléctrico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331400', 'Reparación de equipo eléctrico (excepto reparación de equipo y enseres domésticos)', 3, '312080')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('273200', 'Fabricación de otros hilos y cables eléctricos', 3, '313000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('261000', 'Fabricación de componentes y tableros electrónicos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('273100', 'Fabricación de cables de fibra óptica', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('279000', 'Fabricación de otros tipos de equipo eléctrico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331400', 'Reparación de equipo eléctrico (excepto reparación de equipo y enseres domésticos)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('272000', 'Fabricación de pilas, baterías y acumuladores', 3, '314000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331400', 'Reparación de equipo eléctrico (excepto reparación de equipo y enseres domésticos)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('274000', 'Fabricación de equipo eléctrico de iluminación', 3, '315010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('279000', 'Fabricación de otros tipos de equipo eléctrico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331400', 'Reparación de equipo eléctrico (excepto reparación de equipo y enseres domésticos)', 3, '315080')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('279000', 'Fabricación de otros tipos de equipo eléctrico', 3, '319010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('259900', 'Fabricación de otros productos elaborados de metal n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('263000', 'Fabricación de equipo de comunicaciones', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('265100', 'Fabricación de equipo de medición, prueba, navegación y control', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('273300', 'Fabricación de dispositivos de cableado', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('274000', 'Fabricación de equipo eléctrico de iluminación', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282200', 'Fabricación de maquinaria para la conformación de metales y de máquinas herramienta', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('293000', 'Fabricación de partes, piezas y accesorios para vehículos automotores', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('302000', 'Fabricación de locomotoras y material rodante', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331400', 'Reparación de equipo eléctrico (excepto reparación de equipo y enseres domésticos)', 3, '319080')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331309', 'Reparación de otros equipos electrónicos y ópticos n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('261000', 'Fabricación de componentes y tableros electrónicos', 3, '321010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('279000', 'Fabricación de otros tipos de equipo eléctrico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331400', 'Reparación de equipo eléctrico (excepto reparación de equipo y enseres domésticos)', 3, '321080')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('263000', 'Fabricación de equipo de comunicaciones', 3, '322010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('265100', 'Fabricación de equipo de medición, prueba, navegación y control', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('951200', 'Reparación de equipo de comunicaciones (incluye la reparación teléfonos celulares)', 3, '322080')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331309', 'Reparación de otros equipos electrónicos y ópticos n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('264000', 'Fabricación de aparatos electrónicos de consumo', 3, '323000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('261000', 'Fabricación de componentes y tableros electrónicos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('263000', 'Fabricación de equipo de comunicaciones', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('267000', 'Fabricación de instrumentos ópticos y equipo fotográfico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281700', 'Fabricación de maquinaria y equipo de oficina (excepto computadores y equipo periférico)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331209', 'Reparación de otro tipo de maquinaria y equipos industriales n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331309', 'Reparación de otros equipos electrónicos y ópticos n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('952100', 'Reparación de aparatos electrónicos de consumo (incluye aparatos de televisión y radio)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('325009', 'Fabricación de instrumentos y materiales médicos, oftalmológicos y odontológicos n.c.p.', 3, '331110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('266000', 'Fabricación de equipo de irradiación y equipo electrónico de uso médico y terapéutico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('329000', 'Otras industrias manufactureras n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('325001', 'Actividades de laboratorios dentales', 3, '331120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331309', 'Reparación de otros equipos electrónicos y ópticos n.c.p.', 3, '331180')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331900', 'Reparación de otros tipos de equipo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('265100', 'Fabricación de equipo de medición, prueba, navegación y control', 3, '331210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('267000', 'Fabricación de instrumentos ópticos y equipo fotográfico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281900', 'Fabricación de otros tipos de maquinaria de uso general', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282900', 'Fabricación de otros tipos de maquinaria de uso especial', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('325009', 'Fabricación de instrumentos y materiales médicos, oftalmológicos y odontológicos n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331301', 'Reparación de equipo de medición, prueba, navegación y control', 3, '331280')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331900', 'Reparación de otros tipos de equipo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('265100', 'Fabricación de equipo de medición, prueba, navegación y control', 3, '331310')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331309', 'Reparación de otros equipos electrónicos y ópticos n.c.p.', 3, '331380')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('325009', 'Fabricación de instrumentos y materiales médicos, oftalmológicos y odontológicos n.c.p.', 3, '332010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('267000', 'Fabricación de instrumentos ópticos y equipo fotográfico', 3, '332020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('273100', 'Fabricación de cables de fibra óptica', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282900', 'Fabricación de otros tipos de maquinaria de uso especial', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331309', 'Reparación de otros equipos electrónicos y ópticos n.c.p.', 3, '332080')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('265200', 'Fabricación de relojes', 3, '333000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('321100', 'Fabricación de joyas y artículos conexos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('321200', 'Fabricación de bisutería y artículos conexos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331900', 'Reparación de otros tipos de equipo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('291000', 'Fabricación de vehículos automotores', 3, '341000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('292000', 'Fabricación de carrocerías para vehículos automotores; fabricación de remolques y semirremolques', 3, '342000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331100', 'Reparación de productos elaborados de metal', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('293000', 'Fabricación de partes, piezas y accesorios para vehículos automotores', 3, '343000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('139200', 'Fabricación de artículos confeccionados de materiales textiles, excepto prendas de vestir', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281100', 'Fabricación de motores y turbinas, excepto para aeronaves, vehículos automotores y motocicletas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('301100', 'Construcción de buques, embarcaciones menores y estructuras flotantes', 3, '351110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331501', 'Reparación de buques, embarcaciones menores y estructuras flotantes', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('301100', 'Construcción de buques, embarcaciones menores y estructuras flotantes', 3, '351120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331501', 'Reparación de buques, embarcaciones menores y estructuras flotantes', 3, '351180')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('301200', 'Construcción de embarcaciones de recreo y de deporte', 3, '351210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331501', 'Reparación de buques, embarcaciones menores y estructuras flotantes', 3, '351280')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('302000', 'Fabricación de locomotoras y material rodante', 3, '352000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331509', 'Reparación de otros equipos de transporte n.c.p., excepto vehículos automotores', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('303000', 'Fabricación de aeronaves, naves espaciales y maquinaria conexa', 3, '353010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281100', 'Fabricación de motores y turbinas, excepto para aeronaves, vehículos automotores y motocicletas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282900', 'Fabricación de otros tipos de maquinaria de uso especial', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331502', 'Reparación de aeronaves y naves espaciales', 3, '353080')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('309100', 'Fabricación de motocicletas', 3, '359100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281100', 'Fabricación de motores y turbinas, excepto para aeronaves, vehículos automotores y motocicletas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('309200', 'Fabricación de bicicletas y de sillas de ruedas', 3, '359200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('309900', 'Fabricación de otros tipos de equipo de transporte n.c.p.', 3, '359900')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281600', 'Fabricación de equipo de elevación y manipulación', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('310009', 'Fabricación de colchones; fabricación de otros muebles n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331209', 'Reparación de otro tipo de maquinaria y equipos industriales n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331509', 'Reparación de otros equipos de transporte n.c.p., excepto vehículos automotores', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('310001', 'Fabricación de muebles principalmente de madera', 3, '361010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('310009', 'Fabricación de colchones; fabricación de otros muebles n.c.p.', 3, '361020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('221900', 'Fabricación de otros productos de caucho', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('222000', 'Fabricación de productos de plástico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('281700', 'Fabricación de maquinaria y equipo de oficina (excepto computadores y equipo periférico)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('293000', 'Fabricación de partes, piezas y accesorios para vehículos automotores', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('301100', 'Construcción de buques, embarcaciones menores y estructuras flotantes', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('302000', 'Fabricación de locomotoras y material rodante', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('303000', 'Fabricación de aeronaves, naves espaciales y maquinaria conexa', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('952400', 'Reparación de muebles y accesorios domésticos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('321100', 'Fabricación de joyas y artículos conexos', 3, '369100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('322000', 'Fabricación de instrumentos musicales', 3, '369200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331900', 'Reparación de otros tipos de equipo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('323000', 'Fabricación de artículos de deporte', 3, '369300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('329000', 'Otras industrias manufactureras n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('324000', 'Fabricación de juegos y juguetes', 3, '369400')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('264000', 'Fabricación de aparatos electrónicos de consumo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282900', 'Fabricación de otros tipos de maquinaria de uso especial', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331209', 'Reparación de otro tipo de maquinaria y equipos industriales n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331900', 'Reparación de otros tipos de equipo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('329000', 'Otras industrias manufactureras n.c.p.', 3, '369910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('329000', 'Otras industrias manufactureras n.c.p.', 3, '369920')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('202909', 'Fabricación de otros productos químicos n.c.p.', 3, '369930')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('329000', 'Otras industrias manufactureras n.c.p.', 3, '369990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('139900', 'Fabricación de otros productos textiles n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('151200', 'Fabricación de maletas, bolsos y artículos similares, artículos de talabartería y guarnicionería', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('162900', 'Fabricación de otros productos de madera, de artículos de corcho, paja y materiales trenzables', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('170900', 'Fabricación de otros artículos de papel y cartón', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('221900', 'Fabricación de otros productos de caucho', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('222000', 'Fabricación de productos de plástico', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('259900', 'Fabricación de otros productos elaborados de metal n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('282900', 'Fabricación de otros tipos de maquinaria de uso especial', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('309200', 'Fabricación de bicicletas y de sillas de ruedas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('321200', 'Fabricación de bisutería y artículos conexos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331209', 'Reparación de otro tipo de maquinaria y equipos industriales n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('383001', 'Recuperación y reciclamiento de desperdicios y desechos metálicos', 3, '371000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('383002', 'Recuperación y reciclamiento de papel', 3, '372010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('383003', 'Recuperación y reciclamiento de vidrio', 3, '372020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('383009', 'Recuperación y reciclamiento de otros desperdicios y desechos n.c.p.', 3, '372090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('351011', 'Generación de energía eléctrica en centrales hidroeléctricas', 3, '401011')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('351012', 'Generación de energía eléctrica en centrales termoeléctricas', 3, '401012')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('351012', 'Generación de energía eléctrica en centrales termoeléctricas', 3, '401013')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('351019', 'Generación de energía eléctrica en otras centrales n.c.p.', 3, '401019')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('351020', 'Transmisión de energía eléctrica', 3, '401020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('351030', 'Distribución de energía eléctrica', 3, '401030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('352020', 'Fabricación de gas; distribución de combustibles gaseosos por tubería, excepto regasificación de GNL', 3, '402000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('352010', 'Regasificación de Gas Natural Licuado (GNL)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('353001', 'Suministro de vapor y de aire acondicionado', 3, '403000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('360000', 'Captación, tratamiento y distribución de agua', 3, '410000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('431200', 'Preparación del terreno', 3, '451010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('390000', 'Actividades de descontaminación y otros servicios de gestión de desechos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('431100', 'Demolición', 3, '451020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('410010', 'Construcción de edificios para uso residencial', 3, '452010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('410020', 'Construcción de edificios para uso no residencial', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('439000', 'Otras actividades especializadas de construcción', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('421000', 'Construcción de carreteras y líneas de ferrocarril', 3, '452020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('422000', 'Construcción de proyectos de servicio público', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('429000', 'Construcción de otras obras de ingeniería civil', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('439000', 'Otras actividades especializadas de construcción', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('432100', 'Instalaciones eléctricas', 3, '453000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('432200', 'Instalaciones de gasfitería, calefacción y aire acondicionado', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('432900', 'Otras instalaciones para obras de construcción', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('433000', 'Terminación y acabado de edificios', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('439000', 'Otras actividades especializadas de construcción', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('433000', 'Terminación y acabado de edificios', 3, '454000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('332000', 'Instalación de maquinaria y equipos industriales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('432900', 'Otras instalaciones para obras de construcción', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('439000', 'Otras actividades especializadas de construcción', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('439000', 'Otras actividades especializadas de construcción', 3, '455000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('451001', 'Venta al por mayor de vehículos automotores', 3, '501010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('451002', 'Venta al por menor de vehículos automotores nuevos o usados (incluye compraventa)', 3, '501020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('452001', 'Servicio de lavado de vehículos automotores', 3, '502010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522190', 'Actividades de servicios vinculadas al transporte terrestre n.c.p.', 3, '502020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('452002', 'Mantenimiento y reparación de vehículos automotores', 3, '502080')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('453000', 'Venta de partes, piezas y accesorios para vehículos automotores', 3, '503000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('454001', 'Venta de motocicletas', 3, '504010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('454002', 'Venta de partes, piezas y accesorios de motocicletas', 3, '504020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('454003', 'Mantenimiento y reparación de motocicletas', 3, '504080')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('473000', 'Venta al por menor de combustibles para vehículos automotores en comercios especializados', 3, '505000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('461001', 'Corretaje al por mayor de productos agrícolas', 3, '511010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('461002', 'Corretaje al por mayor de ganado', 3, '511020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('461009', 'Otros tipos de corretajes o remates al por mayor n.c.p.', 3, '511030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('619090', 'Otras actividades de telecomunicaciones n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('462020', 'Venta al por mayor de animales vivos', 3, '512110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('462090', 'Venta al por mayor de otras materias primas agropecuarias n.c.p.', 3, '512120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('462010', 'Venta al por mayor de materias primas agrícolas', 3, '512130')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('462090', 'Venta al por mayor de otras materias primas agropecuarias n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('463011', 'Venta al por mayor de frutas y verduras', 3, '512210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('463012', 'Venta al por mayor de carne y productos cárnicos', 3, '512220')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('463013', 'Venta al por mayor de productos del mar (pescados, mariscos y algas)', 3, '512230')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('463020', 'Venta al por mayor de bebidas alcohólicas y no alcohólicas', 3, '512240')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('463014', 'Venta al por mayor de productos de confitería', 3, '512250')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('463030', 'Venta al por mayor de tabaco', 3, '512260')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('463019', 'Venta al por mayor de huevos, lácteos, abarrotes y de otros alimentos n.c.p.', 3, '512290')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('464100', 'Venta al por mayor de productos textiles, prendas de vestir y calzado', 3, '513100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('464901', 'Venta al por mayor de muebles, excepto muebles de oficina', 3, '513910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('464902', 'Venta al por mayor de artículos eléctricos y electrónicos para el hogar', 3, '513920')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('464903', 'Venta al por mayor de artículos de perfumería, de tocador y cosméticos', 3, '513930')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('464904', 'Venta al por mayor de artículos de papelería y escritorio', 3, '513940')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('464905', 'Venta al por mayor de libros', 3, '513951')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('464906', 'Venta al por mayor de diarios y revistas', 3, '513952')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('464907', 'Venta al por mayor de productos farmacéuticos y medicinales', 3, '513960')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('464908', 'Venta al por mayor de instrumentos científicos y quirúrgicos', 3, '513970')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('464909', 'Venta al por mayor de otros enseres domésticos n.c.p.', 3, '513990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('466302', 'Venta al por mayor de materiales de construcción, artículos de ferretería, gasfitería y calefacción', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('466100', 'Venta al por mayor de combustibles sólidos, líquidos y gaseosos y productos conexos', 3, '514110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('466100', 'Venta al por mayor de combustibles sólidos, líquidos y gaseosos y productos conexos', 3, '514120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('466100', 'Venta al por mayor de combustibles sólidos, líquidos y gaseosos y productos conexos', 3, '514130')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('466100', 'Venta al por mayor de combustibles sólidos, líquidos y gaseosos y productos conexos', 3, '514140')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('466200', 'Venta al por mayor de metales y minerales metalíferos', 3, '514200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('466301', 'Venta al por mayor de madera en bruto y productos primarios de la elaboración de madera', 3, '514310')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('466302', 'Venta al por mayor de materiales de construcción, artículos de ferretería, gasfitería y calefacción', 3, '514320')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('466901', 'Venta al por mayor de productos químicos', 3, '514910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('466902', 'Venta al por mayor de desechos metálicos (chatarra)', 3, '514920')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('464908', 'Venta al por mayor de instrumentos científicos y quirúrgicos', 3, '514930')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('466909', 'Venta al por mayor de desperdicios, desechos y otros productos n.c.p.', 3, '514990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('465300', 'Venta al por mayor de maquinaria, equipo y materiales agropecuarios', 3, '515001')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('465901', 'Venta al por mayor de maquinaria metalúrgica, para la minería, extracción de petróleo y construcción', 3, '515002')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('465901', 'Venta al por mayor de maquinaria metalúrgica, para la minería, extracción de petróleo y construcción', 3, '515003')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('465901', 'Venta al por mayor de maquinaria metalúrgica, para la minería, extracción de petróleo y construcción', 3, '515004')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('465902', 'Venta al por mayor de maquinaria para la elaboración de alimentos, bebidas y tabaco', 3, '515005')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('465903', 'Venta al por mayor de maquinaria para la industria textil, del cuero y del calzado', 3, '515006')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('465100', 'Venta al por mayor de computadores, equipo periférico y programas informáticos', 3, '515007')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('465200', 'Venta al por mayor de equipo, partes y piezas electrónicos y de telecomunicaciones', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('465904', 'Venta al por mayor de maquinaria y equipo de oficina; venta al por mayor de muebles de oficina', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('465905', 'Venta al por mayor de equipo de transporte(excepto vehículos automotores, motocicletas y bicicletas)', 3, '515008')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('465909', 'Venta al por mayor de otros tipos de maquinaria y equipo n.c.p.', 3, '515009')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('469000', 'Venta al por mayor no especializada', 3, '519000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('471100', 'Venta al por menor en comercios de alimentos, bebidas o tabaco (supermercados e hipermercados)', 3, '521111')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('472101', 'Venta al por menor de alimentos en comercios especializados (almacenes pequeños y minimarket)', 3, '521112')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('471100', 'Venta al por menor en comercios de alimentos, bebidas o tabaco (supermercados e hipermercados)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('472101', 'Venta al por menor de alimentos en comercios especializados (almacenes pequeños y minimarket)', 3, '521120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('471990', 'Otras actividades de venta al por menor en comercios no especializados n.c.p.', 3, '521200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('471910', 'Venta al por menor en comercios de vestuario y productos para el hogar (grandes tiendas)', 3, '521300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477399', 'Venta al por menor de otros productos en comercios especializados n.c.p.', 3, '521900')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('472200', 'Venta al por menor de bebidas alcohólicas y no alcohólicas en comercios especializados (botillerías)', 3, '522010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('472102', 'Venta al por menor en comercios especializados de carne y productos cárnicos', 3, '522020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('472103', 'Venta al por menor en comercios especializados de frutas y verduras (verdulerías)', 3, '522030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('472104', 'Venta al por menor en comercios especializados de pescado, mariscos y productos conexos', 3, '522040')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('472105', 'Venta al por menor en comercios especializados de productos de panadería y pastelería', 3, '522050')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477391', 'Venta al por menor de alimento y accesorios para mascotas en comercios especializados', 3, '522060')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('472109', 'Venta al por menor en comercios especializados de huevos, confites y productos alimenticios n.c.p.', 3, '522070')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('472109', 'Venta al por menor en comercios especializados de huevos, confites y productos alimenticios n.c.p.', 3, '522090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('472300', 'Venta al por menor de tabaco y productos de tabaco en comercios especializados', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477201', 'Venta al por menor de productos farmacéuticos y medicinales en comercios especializados', 3, '523111')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477201', 'Venta al por menor de productos farmacéuticos y medicinales en comercios especializados', 3, '523112')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477201', 'Venta al por menor de productos farmacéuticos y medicinales en comercios especializados', 3, '523120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477202', 'Venta al por menor de artículos ortopédicos en comercios especializados', 3, '523130')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477203', 'Venta al por menor de artículos de perfumería, de tocador y cosméticos en comercios especializados', 3, '523140')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477101', 'Venta al por menor de calzado en comercios especializados', 3, '523210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477102', 'Venta al por menor de prendas y accesorios de vestir en comercios especializados', 3, '523220')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('475100', 'Venta al por menor de telas, lanas, hilos y similares en comercios especializados', 3, '523230')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477103', 'Venta al por menor de carteras, maletas y otros accesorios de viaje en comercios especializados', 3, '523240')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477102', 'Venta al por menor de prendas y accesorios de vestir en comercios especializados', 3, '523250')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('475100', 'Venta al por menor de telas, lanas, hilos y similares en comercios especializados', 3, '523290')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('475300', 'Venta al por menor de tapices, alfombras y cubrimientos para paredes y pisos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('475909', 'Venta al por menor de aparatos eléctricos, textiles para el hogar y otros enseres domésticos n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477103', 'Venta al por menor de carteras, maletas y otros accesorios de viaje en comercios especializados', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('475909', 'Venta al por menor de aparatos eléctricos, textiles para el hogar y otros enseres domésticos n.c.p.', 3, '523310')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('474200', 'Venta al por menor de equipo de sonido y de video en comercios especializados', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('475909', 'Venta al por menor de aparatos eléctricos, textiles para el hogar y otros enseres domésticos n.c.p.', 3, '523320')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('475901', 'Venta al por menor de muebles y colchones en comercios especializados', 3, '523330')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('475902', 'Venta al por menor de instrumentos musicales en comercios especializados', 3, '523340')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('476200', 'Venta al por menor de grabaciones de música y de video en comercios especializados', 3, '523350')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('475909', 'Venta al por menor de aparatos eléctricos, textiles para el hogar y otros enseres domésticos n.c.p.', 3, '523360')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('475909', 'Venta al por menor de aparatos eléctricos, textiles para el hogar y otros enseres domésticos n.c.p.', 3, '523390')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('475201', 'Venta al por menor de artículos de ferretería y materiales de construcción', 3, '523410')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('475202', 'Venta al por menor de pinturas, barnices y lacas en comercios especializados', 3, '523420')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('475203', 'Venta al por menor de productos de vidrio en comercios especializados', 3, '523430')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477399', 'Venta al por menor de otros productos en comercios especializados n.c.p.', 3, '523911')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477393', 'Venta al por menor de artículos ópticos en comercios especializados', 3, '523912')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('476400', 'Venta al por menor de juegos y juguetes en comercios especializados', 3, '523921')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('474100', 'Venta al por menor de computadores, equipo periférico, programas informáticos y equipo de telecom.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('476101', 'Venta al por menor de libros en comercios especializados', 3, '523922')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('476102', 'Venta al por menor de diarios y revistas en comercios especializados', 3, '523923')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('476103', 'Venta al por menor de artículos de papelería y escritorio en comercios especializados', 3, '523924')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('474100', 'Venta al por menor de computadores, equipo periférico, programas informáticos y equipo de telecom.', 3, '523930')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('476301', 'Venta al por menor de artículos de caza y pesca en comercios especializados', 3, '523941')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477392', 'Venta al por menor de armas y municiones en comercios especializados', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('476302', 'Venta al por menor de bicicletas y sus repuestos en comercios especializados', 3, '523942')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('476309', 'Venta al por menor de otros artículos y equipos de deporte n.c.p.', 3, '523943')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477394', 'Venta al por menor de artículos de joyería, bisutería y relojería en comercios especializados', 3, '523950')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477310', 'Venta al por menor de gas licuado en bombonas (cilindros) en comercios especializados', 3, '523961')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477395', 'Venta al por menor de carbón, leña y otros combustibles de uso doméstico en comercios especializados', 3, '523969')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477396', 'Venta al por menor de recuerdos, artesanías y artículos religiosos en comercios especializados', 3, '523991')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477397', 'Venta al por menor de flores, plantas, arboles, semillas y abonos en comercios especializados', 3, '523992')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477398', 'Venta al por menor de mascotas en comercios especializados', 3, '523993')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477391', 'Venta al por menor de alimento y accesorios para mascotas en comercios especializados', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477399', 'Venta al por menor de otros productos en comercios especializados n.c.p.', 3, '523999')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('475300', 'Venta al por menor de tapices, alfombras y cubrimientos para paredes y pisos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477401', 'Venta al por menor de antigüedades en comercios', 3, '524010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477402', 'Venta al por menor de ropa usada en comercios', 3, '524020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('477409', 'Venta al por menor de otros artículos de segunda mano en comercios n.c.p.', 3, '524090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('649209', 'Otras actividades de concesión de crédito n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('479100', 'Venta al por menor por correo, por Internet y vía telefónica', 3, '525110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('479100', 'Venta al por menor por correo, por Internet y vía telefónica', 3, '525120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('479100', 'Venta al por menor por correo, por Internet y vía telefónica', 3, '525130')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('478100', 'Venta al por menor de alimentos, bebidas y tabaco en puestos de venta y mercados (incluye ferias)', 3, '525200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('478200', 'Venta al por menor de productos textiles, prendas de vestir y calzado en puestos de venta y mercados', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('478900', 'Venta al por menor de otros productos en puestos de venta y mercados (incluye ferias)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('479901', 'Venta al por menor realizada por independientes en la locomoción colectiva (Ley 20.388)', 3, '525911')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('479909', 'Otras actividades de venta por menor no realizadas en comercios, puestos de venta o mercados n.c.p.', 3, '525919')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('479902', 'Venta al por menor mediante maquinas expendedoras', 3, '525920')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('479903', 'Venta al por menor por comisionistas (no dependientes de comercios)', 3, '525930')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('479909', 'Otras actividades de venta por menor no realizadas en comercios, puestos de venta o mercados n.c.p.', 3, '525990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('479100', 'Venta al por menor por correo, por Internet y vía telefónica', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('952300', 'Reparación de calzado y de artículos de cuero', 3, '526010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('952200', 'Reparación de aparatos de uso doméstico, equipo doméstico y de jardinería', 3, '526020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331309', 'Reparación de otros equipos electrónicos y ópticos n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('952100', 'Reparación de aparatos electrónicos de consumo (incluye aparatos de televisión y radio)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('952900', 'Reparación de otros efectos personales y enseres domésticos', 3, '526030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('952900', 'Reparación de otros efectos personales y enseres domésticos', 3, '526090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('802000', 'Actividades de servicios de sistemas de seguridad (incluye servicios de cerrajería)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('951200', 'Reparación de equipo de comunicaciones (incluye la reparación teléfonos celulares)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('952400', 'Reparación de muebles y accesorios domésticos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('551001', 'Actividades de hoteles', 3, '551010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('551002', 'Actividades de moteles', 3, '551020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('559001', 'Actividades de residenciales para estudiantes y trabajadores', 3, '551030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('551003', 'Actividades de residenciales para turistas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('559009', 'Otras actividades de alojamiento n.c.p.', 3, '551090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('551009', 'Otras actividades de alojamiento para turistas n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('552000', 'Actividades de camping y de parques para casas rodantes', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('561000', 'Actividades de restaurantes y de servicio móvil de comidas', 3, '552010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('561000', 'Actividades de restaurantes y de servicio móvil de comidas', 3, '552020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('563009', 'Otras actividades de servicio de bebidas n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('562900', 'Suministro industrial de comidas por encargo; concesión de servicios de alimentación', 3, '552030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('562900', 'Suministro industrial de comidas por encargo; concesión de servicios de alimentación', 3, '552040')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('562100', 'Suministro de comidas por encargo (Servicios de banquetería)', 3, '552050')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('562900', 'Suministro industrial de comidas por encargo; concesión de servicios de alimentación', 3, '552090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('561000', 'Actividades de restaurantes y de servicio móvil de comidas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('563009', 'Otras actividades de servicio de bebidas n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('491100', 'Transporte interurbano de pasajeros por ferrocarril', 3, '601001')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522190', 'Actividades de servicios vinculadas al transporte terrestre n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('491200', 'Transporte de carga por ferrocarril', 3, '601002')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522190', 'Actividades de servicios vinculadas al transporte terrestre n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('492110', 'Transporte urbano y suburbano de pasajeros vía metro y metrotren', 3, '602110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('492120', 'Transporte urbano y suburbano de pasajeros vía locomoción colectiva', 3, '602120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('492250', 'Transporte de pasajeros en buses interurbanos', 3, '602130')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('492130', 'Transporte de pasajeros vía taxi colectivo', 3, '602140')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('492210', 'Servicios de transporte de escolares', 3, '602150')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('492220', 'Servicios de transporte de trabajadores', 3, '602160')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('492290', 'Otras actividades de transporte de pasajeros por vía terrestre n.c.p.', 3, '602190')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('492190', 'Otras actividades de transporte urbano y suburbano de pasajeros por vía terrestre n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('492230', 'Servicios de transporte de pasajeros en taxis libres y radiotaxis', 3, '602210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('492240', 'Servicios de transporte a turistas', 3, '602220')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('492290', 'Otras actividades de transporte de pasajeros por vía terrestre n.c.p.', 3, '602230')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('492290', 'Otras actividades de transporte de pasajeros por vía terrestre n.c.p.', 3, '602290')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('492300', 'Transporte de carga por carretera', 3, '602300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('493090', 'Otras actividades de transporte por tuberías n.c.p.', 3, '603000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('493010', 'Transporte por oleoductos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('493020', 'Transporte por gasoductos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('501100', 'Transporte de pasajeros marítimo y de cabotaje', 3, '611001')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('501200', 'Transporte de carga marítimo y de cabotaje', 3, '611002')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('502100', 'Transporte de pasajeros por vías de navegación interiores', 3, '612001')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('502200', 'Transporte de carga por vías de navegación interiores', 3, '612002')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('511000', 'Transporte de pasajeros por vía aérea', 3, '621010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('512000', 'Transporte de carga por vía aérea', 3, '621020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('511000', 'Transporte de pasajeros por vía aérea', 3, '622001')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('512000', 'Transporte de carga por vía aérea', 3, '622002')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522400', 'Manipulación de la carga', 3, '630100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('521009', 'Otros servicios de almacenamiento y depósito n.c.p.', 3, '630200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522110', 'Explotación de terminales terrestres de pasajeros', 3, '630310')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522120', 'Explotación de estacionamientos de vehículos automotores y parquímetros', 3, '630320')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522200', 'Actividades de servicios vinculadas al transporte acuático', 3, '630330')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522300', 'Actividades de servicios vinculadas al transporte aéreo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522130', 'Servicios prestados por concesionarios de carreteras', 3, '630340')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522190', 'Actividades de servicios vinculadas al transporte terrestre n.c.p.', 3, '630390')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331509', 'Reparación de otros equipos de transporte n.c.p., excepto vehículos automotores', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522200', 'Actividades de servicios vinculadas al transporte acuático', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522300', 'Actividades de servicios vinculadas al transporte aéreo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('791100', 'Actividades de agencias de viajes', 3, '630400')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('791200', 'Actividades de operadores turísticos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('799000', 'Otros servicios de reservas y actividades conexas (incluye venta de entradas para teatro, y otros)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522910', 'Agencias de aduanas', 3, '630910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522990', 'Otras actividades de apoyo al transporte n.c.p.', 3, '630920')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('522920', 'Agencias de naves', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('531000', 'Actividades postales', 3, '641100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('821900', 'Fotocopiado, preparación de documentos y otras actividades especializadas de apoyo de oficina', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('532000', 'Actividades de mensajería', 3, '641200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('611010', 'Telefonía fija', 3, '642010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('612010', 'Telefonía móvil celular', 3, '642020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('613010', 'Telefonía móvil satelital', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('611020', 'Telefonía larga distancia', 3, '642030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('613020', 'Televisión de pago satelital', 3, '642040')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('611030', 'Televisión de pago por cable', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('612030', 'Televisión de pago inalámbrica', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('611090', 'Otros servicios de telecomunicaciones alámbricas n.c.p.', 3, '642050')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('612090', 'Otros servicios de telecomunicaciones inalámbricas n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('613090', 'Otros servicios de telecomunicaciones por satélite n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('619010', 'Centros de llamados y centros de acceso a Internet', 3, '642061')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('619010', 'Centros de llamados y centros de acceso a Internet', 3, '642062')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('619090', 'Otras actividades de telecomunicaciones n.c.p.', 3, '642090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('611090', 'Otros servicios de telecomunicaciones alámbricas n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('612020', 'Radiocomunicaciones móviles', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('612090', 'Otros servicios de telecomunicaciones inalámbricas n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('613090', 'Otros servicios de telecomunicaciones por satélite n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('641100', 'Banca central', 3, '651100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('641910', 'Actividades bancarias', 3, '651910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('649201', 'Financieras', 3, '651920')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('641990', 'Otros tipos de intermediación monetaria n.c.p.', 3, '651990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('649100', 'Leasing financiero', 3, '659110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('649100', 'Leasing financiero', 3, '659120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('649209', 'Otras actividades de concesión de crédito n.c.p.', 3, '659210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('649202', 'Actividades de crédito prendario', 3, '659220')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('649900', 'Otras actividades de servicios financieros, excepto las de seguros y fondos de pensiones n.c.p.', 3, '659231')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('661201', 'Actividades de securitizadoras', 3, '659232')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('649209', 'Otras actividades de concesión de crédito n.c.p.', 3, '659290')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('663091', 'Administradoras de fondos de inversión', 3, '659911')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('663092', 'Administradoras de fondos mutuos', 3, '659912')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('663093', 'Administradoras de fices (fondos de inversión de capital extranjero)', 3, '659913')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('663094', 'Administradoras de fondos para la vivienda', 3, '659914')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('663099', 'Administradoras de fondos para otros fines n.c.p.', 3, '659915')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('643000', 'Fondos y sociedades de inversión y entidades financieras similares', 3, '659920')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('642000', 'Actividades de sociedades de cartera', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('774000', 'Arrendamiento de propiedad intelectual y similares, excepto obras protegidas por derechos de autor', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('949909', 'Actividades de otras asociaciones n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('651100', 'Seguros de vida', 3, '660101')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('652000', 'Reaseguros', 3, '660102')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('663010', 'Administradoras de Fondos de Pensiones (AFP)', 3, '660200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('653000', 'Fondos de pensiones', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('651210', 'Seguros generales, excepto actividades de Isapres', 3, '660301')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('651100', 'Seguros de vida', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('652000', 'Reaseguros', 3, '660302')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('651220', 'Actividades de Isapres', 3, '660400')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('661100', 'Administración de mercados financieros', 3, '671100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('661202', 'Corredores de bolsa', 3, '671210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('661203', 'Agentes de valores', 3, '671220')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('661209', 'Otros servicios de corretaje de valores y commodities n.c.p.', 3, '671290')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('661901', 'Actividades de cámaras de compensación', 3, '671910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('661902', 'Administración de tarjetas de crédito', 3, '671921')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('661903', 'Empresas de asesoría y consultoría en inversión financiera; sociedades de apoyo al giro', 3, '671929')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('661904', 'Actividades de clasificadoras de riesgo', 3, '671930')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('661204', 'Actividades de casas de cambio y operadores de divisa', 3, '671940')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('661909', 'Otras actividades auxiliares de las actividades de servicios financieros n.c.p.', 3, '671990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('641990', 'Otros tipos de intermediación monetaria n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('662200', 'Actividades de agentes y corredores de seguros', 3, '672010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('662100', 'Evaluación de riesgos y daños (incluye actividades de liquidadores de seguros)', 3, '672020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('662900', 'Otras actividades auxiliares de las actividades de seguros y fondos de pensiones', 3, '672090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('681011', 'Alquiler de bienes inmuebles amoblados o con equipos y maquinarias', 3, '701001')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('681012', 'Compra, venta y alquiler (excepto amoblados) de inmuebles', 3, '701009')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('429000', 'Construcción de otras obras de ingeniería civil', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('681020', 'Servicios imputados de alquiler de viviendas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('682000', 'Actividades inmobiliarias realizadas a cambio de una retribución o por contrata', 3, '702000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('811000', 'Actividades combinadas de apoyo a instalaciones', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('771000', 'Alquiler de vehículos automotores sin chofer', 3, '711101')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('773001', 'Alquiler de equipos de transporte sin operario, excepto vehículos automotores', 3, '711102')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('771000', 'Alquiler de vehículos automotores sin chofer', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('773001', 'Alquiler de equipos de transporte sin operario, excepto vehículos automotores', 3, '711200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('773001', 'Alquiler de equipos de transporte sin operario, excepto vehículos automotores', 3, '711300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('773002', 'Alquiler de maquinaria y equipo agropecuario, forestal, de construcción e ing. civil, sin operarios', 3, '712100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('773002', 'Alquiler de maquinaria y equipo agropecuario, forestal, de construcción e ing. civil, sin operarios', 3, '712200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('773003', 'Alquiler de maquinaria y equipo de oficina, sin operarios (sin servicio administrativo)', 3, '712300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('773009', 'Alquiler de otros tipos de maquinarias y equipos sin operario n.c.p.', 3, '712900')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('772100', 'Alquiler y arrendamiento de equipo recreativo y deportivo', 3, '713010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('772900', 'Alquiler de otros efectos personales y enseres domésticos (incluye mobiliario para eventos)', 3, '713020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('772200', 'Alquiler de cintas de video y discos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('772900', 'Alquiler de otros efectos personales y enseres domésticos (incluye mobiliario para eventos)', 3, '713030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('773009', 'Alquiler de otros tipos de maquinarias y equipos sin operario n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('772900', 'Alquiler de otros efectos personales y enseres domésticos (incluye mobiliario para eventos)', 3, '713090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('772100', 'Alquiler y arrendamiento de equipo recreativo y deportivo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('773009', 'Alquiler de otros tipos de maquinarias y equipos sin operario n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('620100', 'Actividades de programación informática', 3, '722000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('631100', 'Procesamiento de datos, hospedaje y actividades conexas', 3, '724000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('581100', 'Edición de libros', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('581200', 'Edición de directorios y listas de correo', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('581300', 'Edición de diarios, revistas y otras publicaciones periódicas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('581900', 'Otras actividades de edición', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('582000', 'Edición de programas informáticos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('592000', 'Actividades de grabación de sonido y edición de música', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('601000', 'Transmisiones de radio', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('602000', 'Programación y transmisiones de televisión', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('620200', 'Actividades de consultoría de informática y de gestión de instalaciones informáticas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('631200', 'Portales web', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('951100', 'Reparación de computadores y equipo periférico', 3, '725000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('331209', 'Reparación de otro tipo de maquinaria y equipos industriales n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('620200', 'Actividades de consultoría de informática y de gestión de instalaciones informáticas', 3, '726000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('620900', 'Otras actividades de tecnología de la información y de servicios informáticos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('721000', 'Investigaciones y desarrollo experimental en el campo de las ciencias naturales y la ingeniería', 3, '731000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('722000', 'Investigaciones y desarrollo experimental en el campo de las ciencias sociales y las humanidades', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('722000', 'Investigaciones y desarrollo experimental en el campo de las ciencias sociales y las humanidades', 3, '732000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('691001', 'Servicios de asesoramiento y representación jurídica', 3, '741110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('691002', 'Servicio notarial', 3, '741120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('691003', 'Conservador de bienes raíces', 3, '741130')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('691004', 'Receptores judiciales', 3, '741140')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('691009', 'Servicios de arbitraje; síndicos de quiebra y peritos judiciales; otras actividades jurídicas n.c.p.', 3, '741190')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('692000', 'Actividades de contabilidad, teneduría de libros y auditoría; consultoría fiscal', 3, '741200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('732000', 'Estudios de mercado y encuestas de opinión pública', 3, '741300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('702000', 'Actividades de consultoría de gestión', 3, '741400')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('701000', 'Actividades de oficinas principales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('855000', 'Actividades de apoyo a la enseñanza', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('711001', 'Servicios de arquitectura (diseño de edificios, dibujo de planos de construcción, entre otros)', 3, '742110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099001', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por empresas', 3, '742121')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('091001', 'Actividades de apoyo para la extracción de petróleo y gas natural prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('711002', 'Empresas de servicios de ingeniería y actividades conexas de consultoría técnica', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('099002', 'Actividades de apoyo para la explotación de otras minas y canteras prestados por profesionales', 3, '742122')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('091002', 'Actividades de apoyo para la extracción de petróleo y gas natural prestados por profesionales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('711003', 'Servicios profesionales de ingeniería y actividades conexas de consultoría técnica', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('711002', 'Empresas de servicios de ingeniería y actividades conexas de consultoría técnica', 3, '742131')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('711003', 'Servicios profesionales de ingeniería y actividades conexas de consultoría técnica', 3, '742132')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('711002', 'Empresas de servicios de ingeniería y actividades conexas de consultoría técnica', 3, '742141')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('741009', 'Otras actividades especializadas de diseño n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('711003', 'Servicios profesionales de ingeniería y actividades conexas de consultoría técnica', 3, '742142')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('741009', 'Otras actividades especializadas de diseño n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('749009', 'Otras actividades profesionales, científicas y técnicas n.c.p.', 3, '742190')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('711003', 'Servicios profesionales de ingeniería y actividades conexas de consultoría técnica', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('712001', 'Actividades de plantas de revisión técnica para vehículos automotores', 3, '742210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('712009', 'Otros servicios de ensayos y análisis técnicos (excepto actividades de plantas de revisión técnica)', 3, '742290')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('731001', 'Servicios de publicidad prestados por empresas', 3, '743001')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('731002', 'Servicios de publicidad prestados por profesionales', 3, '743002')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('783000', 'Otras actividades de dotación de recursos humanos', 3, '749110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('782000', 'Actividades de agencias de empleo temporal (incluye empresas de servicios transitorios)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('781000', 'Actividades de agencias de empleo', 3, '749190')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('803000', 'Actividades de investigación (incluye actividades de investigadores y detectives privados)', 3, '749210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('801001', 'Servicios de seguridad privada prestados por empresas', 3, '749221')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('802000', 'Actividades de servicios de sistemas de seguridad (incluye servicios de cerrajería)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('801002', 'Servicio de transporte de valores en vehículos blindados', 3, '749222')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('801003', 'Servicios de seguridad privada prestados por independientes', 3, '749229')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('812100', 'Limpieza general de edificios', 3, '749310')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('812909', 'Otras actividades de limpieza de edificios e instalaciones industriales n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('812901', 'Desratización, desinfección y exterminio de plagas no agrícolas', 3, '749320')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('742001', 'Servicios de revelado, impresión y ampliación de fotografías', 3, '749401')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('742002', 'Servicios y actividades de fotografía', 3, '749402')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('742003', 'Servicios personales de fotografía', 3, '749409')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('829200', 'Actividades de envasado y empaquetado', 3, '749500')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('829110', 'Actividades de agencias de cobro', 3, '749911')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('829120', 'Actividades de agencias de calificación crediticia', 3, '749912')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('749001', 'Asesoría y gestión en la compra o venta de pequeñas y medianas empresas', 3, '749913')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('741001', 'Actividades de diseño de vestuario', 3, '749921')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('741002', 'Actividades de diseño y decoración de interiores', 3, '749922')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('741009', 'Otras actividades especializadas de diseño n.c.p.', 3, '749929')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('821900', 'Fotocopiado, preparación de documentos y otras actividades especializadas de apoyo de oficina', 3, '749931')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('821100', 'Actividades combinadas de servicios administrativos de oficina', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('749003', 'Servicios personales de traducción e interpretación', 3, '749932')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('749002', 'Servicios de traducción e interpretación prestados por empresas', 3, '749933')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('821900', 'Fotocopiado, preparación de documentos y otras actividades especializadas de apoyo de oficina', 3, '749934')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('749004', 'Actividades de agencias y agentes de representación de actores, deportistas y otras figuras públicas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('829900', 'Otras actividades de servicios de apoyo a las empresas n.c.p.', 3, '749950')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('823000', 'Organización de convenciones y exposiciones comerciales', 3, '749961')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('823000', 'Organización de convenciones y exposiciones comerciales', 3, '749962')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('822000', 'Actividades de call-center', 3, '749970')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('829900', 'Otras actividades de servicios de apoyo a las empresas n.c.p.', 3, '749990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('639900', 'Otras actividades de servicios de información n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('731001', 'Servicios de publicidad prestados por empresas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('742002', 'Servicios y actividades de fotografía', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('855000', 'Actividades de apoyo a la enseñanza', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('841100', 'Actividades de la administración pública en general', 3, '751110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('681011', 'Alquiler de bienes inmuebles amoblados o con equipos y maquinarias', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('681012', 'Compra, venta y alquiler (excepto amoblados) de inmuebles', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('682000', 'Actividades inmobiliarias realizadas a cambio de una retribución o por contrata', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('799000', 'Otros servicios de reservas y actividades conexas (incluye venta de entradas para teatro, y otros)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('841200', 'Regulación de las actividades de organismos que prestan servicios sanitarios, educativos, culturales', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('841300', 'Regulación y facilitación de la actividad económica', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('910100', 'Actividades de bibliotecas y archivos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('841100', 'Actividades de la administración pública en general', 3, '751120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('842300', 'Actividades de mantenimiento del orden público y de seguridad', 3, '751200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('841100', 'Actividades de la administración pública en general', 3, '751300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('842100', 'Relaciones exteriores', 3, '752100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('889000', 'Otras actividades de asistencia social sin alojamiento', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('842200', 'Actividades de defensa', 3, '752200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('842300', 'Actividades de mantenimiento del orden público y de seguridad', 3, '752300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('712009', 'Otros servicios de ensayos y análisis técnicos (excepto actividades de plantas de revisión técnica)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('843090', 'Otros planes de seguridad social de afiliación obligatoria n.c.p.', 3, '753010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('843010', 'Fondo Nacional de Salud (FONASA)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('843020', 'Instituto de Previsión Social (IPS)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('649203', 'Cajas de compensación', 3, '753020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('843090', 'Otros planes de seguridad social de afiliación obligatoria n.c.p.', 3, '753090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('850021', 'Enseñanza preescolar privada', 3, '801010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('850011', 'Enseñanza preescolar pública', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('850022', 'Enseñanza primaria, secundaria científico humanista y técnico profesional privada', 3, '801020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('850012', 'Enseñanza primaria, secundaria científico humanista y técnico profesional pública', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('850022', 'Enseñanza primaria, secundaria científico humanista y técnico profesional privada', 3, '802100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('850012', 'Enseñanza primaria, secundaria científico humanista y técnico profesional pública', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('850022', 'Enseñanza primaria, secundaria científico humanista y técnico profesional privada', 3, '802200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('850012', 'Enseñanza primaria, secundaria científico humanista y técnico profesional pública', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('853120', 'Enseñanza superior en universidades privadas', 3, '803010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('853110', 'Enseñanza superior en universidades públicas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('853201', 'Enseñanza superior en institutos profesionales', 3, '803020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('853202', 'Enseñanza superior en centros de formación técnica', 3, '803030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('850022', 'Enseñanza primaria, secundaria científico humanista y técnico profesional privada', 3, '809010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('850012', 'Enseñanza primaria, secundaria científico humanista y técnico profesional pública', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('854901', 'Enseñanza preuniversitaria', 3, '809020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('854909', 'Otros tipos de enseñanza n.c.p.', 3, '809030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('854909', 'Otros tipos de enseñanza n.c.p.', 3, '809041')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('854902', 'Servicios personales de educación', 3, '809049')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('861020', 'Actividades de hospitales y clínicas privadas', 3, '851110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('861010', 'Actividades de hospitales y clínicas públicas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('861020', 'Actividades de hospitales y clínicas privadas', 3, '851120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('861010', 'Actividades de hospitales y clínicas públicas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('862031', 'Servicios de médicos prestados de forma independiente', 3, '851211')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('862021', 'Centros médicos privados (establecimientos de atención ambulatoria)', 3, '851212')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('862010', 'Actividades de centros de salud municipalizados (servicios de salud pública)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('862032', 'Servicios de odontólogos prestados de forma independiente', 3, '851221')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('862022', 'Centros de atención odontológica privados (establecimientos de atención ambulatoria)', 3, '851222')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('869010', 'Actividades de laboratorios clínicos y bancos de sangre', 3, '851910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('869092', 'Servicios prestados de forma independiente por otros profesionales de la salud', 3, '851920')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('869091', 'Otros servicios de atención de la salud humana prestados por empresas', 3, '851990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('871000', 'Actividades de atención de enfermería en instituciones', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('872000', 'Actividades de atención en instituciones para personas con discapacidad mental y toxicómanos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('873000', 'Actividades de atención en instituciones para personas de edad y personas con discapacidad física', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('750001', 'Actividades de clínicas veterinarias', 3, '852010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('750002', 'Actividades de veterinarios, técnicos y otro personal auxiliar, prestados de forma independiente', 3, '852021')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('750002', 'Actividades de veterinarios, técnicos y otro personal auxiliar, prestados de forma independiente', 3, '852029')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('879000', 'Otras actividades de atención en instituciones', 3, '853100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('872000', 'Actividades de atención en instituciones para personas con discapacidad mental y toxicómanos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('873000', 'Actividades de atención en instituciones para personas de edad y personas con discapacidad física', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('889000', 'Otras actividades de asistencia social sin alojamiento', 3, '853200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('561000', 'Actividades de restaurantes y de servicio móvil de comidas', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('855000', 'Actividades de apoyo a la enseñanza', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('881000', 'Actividades de asistencia social sin alojamiento para personas de edad y personas con discapacidad', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('382100', 'Tratamiento y eliminación de desechos no peligrosos', 3, '900010')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('382200', 'Tratamiento y eliminación de desechos peligrosos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('812909', 'Otras actividades de limpieza de edificios e instalaciones industriales n.c.p.', 3, '900020')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('813000', 'Actividades de paisajismo, servicios de jardinería y servicios conexos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('381100', 'Recogida de desechos no peligrosos', 3, '900030')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('381200', 'Recogida de desechos peligrosos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('382100', 'Tratamiento y eliminación de desechos no peligrosos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('382200', 'Tratamiento y eliminación de desechos peligrosos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('370000', 'Evacuación y tratamiento de aguas servidas', 3, '900040')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('370000', 'Evacuación y tratamiento de aguas servidas', 3, '900050')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('370000', 'Evacuación y tratamiento de aguas servidas', 3, '900090')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('390000', 'Actividades de descontaminación y otros servicios de gestión de desechos', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('941100', 'Actividades de asociaciones empresariales y de empleadores', 3, '911100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('941200', 'Actividades de asociaciones profesionales', 3, '911210')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('941200', 'Actividades de asociaciones profesionales', 3, '911290')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('942000', 'Actividades de sindicatos', 3, '912000')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('949100', 'Actividades de organizaciones religiosas', 3, '919100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('949200', 'Actividades de organizaciones políticas', 3, '919200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('949901', 'Actividades de centros de madres', 3, '919910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('889000', 'Otras actividades de asistencia social sin alojamiento', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('949902', 'Actividades de clubes sociales', 3, '919920')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('949903', 'Fundaciones y corporaciones; asociaciones que promueven actividades culturales o recreativas', 3, '919930')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('949909', 'Actividades de otras asociaciones n.c.p.', 3, '919990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('591100', 'Actividades de producción de películas cinematográficas, videos y programas de televisión', 3, '921110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('591200', 'Actividades de postproducción de películas cinematográficas, videos y programas de televisión', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('592000', 'Actividades de grabación de sonido y edición de música', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('591300', 'Actividades de distribución de películas cinematográficas, videos y programas de televisión', 3, '921120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('591400', 'Actividades de exhibición de películas cinematográficas y cintas de video', 3, '921200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('602000', 'Programación y transmisiones de televisión', 3, '921310')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('591100', 'Actividades de producción de películas cinematográficas, videos y programas de televisión', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('601000', 'Transmisiones de radio', 3, '921320')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('592000', 'Actividades de grabación de sonido y edición de música', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('900001', 'Servicios de producción de obras de teatro, conciertos, espectáculos de danza, otras prod. escénicas', 3, '921411')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('900009', 'Otras actividades creativas, artísticas y de entretenimiento n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('900001', 'Servicios de producción de obras de teatro, conciertos, espectáculos de danza, otras prod. escénicas', 3, '921419')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('900009', 'Otras actividades creativas, artísticas y de entretenimiento n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('900002', 'Actividades artísticas realizadas por bandas de música, compañías de teatro, circenses y similares', 3, '921420')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('900009', 'Otras actividades creativas, artísticas y de entretenimiento n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('900003', 'Actividades de artistas realizadas de forma independiente: actores, músicos, escritores, entre otros', 3, '921430')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('900009', 'Otras actividades creativas, artísticas y de entretenimiento n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('799000', 'Otros servicios de reservas y actividades conexas (incluye venta de entradas para teatro, y otros)', 3, '921490')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('854200', 'Enseñanza cultural', 3, '921911')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('563001', 'Actividades de discotecas y cabaret (night club), con predominio del servicio de bebidas', 3, '921912')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('932909', 'Otras actividades de esparcimiento y recreativas n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('932100', 'Actividades de parques de atracciones y parques temáticos', 3, '921920')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('932909', 'Otras actividades de esparcimiento y recreativas n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('900001', 'Servicios de producción de obras de teatro, conciertos, espectáculos de danza, otras prod. escénicas', 3, '921930')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('932909', 'Otras actividades de esparcimiento y recreativas n.c.p.', 3, '921990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('799000', 'Otros servicios de reservas y actividades conexas (incluye venta de entradas para teatro, y otros)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('639100', 'Actividades de agencias de noticias', 3, '922001')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('900004', 'Servicios prestados por periodistas independientes', 3, '922002')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('742002', 'Servicios y actividades de fotografía', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('910100', 'Actividades de bibliotecas y archivos', 3, '923100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('591200', 'Actividades de postproducción de películas cinematográficas, videos y programas de televisión', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('910200', 'Actividades de museos, gestión de lugares y edificios históricos', 3, '923200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('910300', 'Actividades de jardines botánicos, zoológicos y reservas naturales', 3, '923300')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('931109', 'Gestión de otras instalaciones deportivas n.c.p.', 3, '924110')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('492290', 'Otras actividades de transporte de pasajeros por vía terrestre n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('931909', 'Otras actividades deportivas n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('932909', 'Otras actividades de esparcimiento y recreativas n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('931209', 'Actividades de otros clubes deportivos n.c.p.', 3, '924120')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('931201', 'Actividades de clubes de fútbol amateur y profesional', 3, '924131')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('931201', 'Actividades de clubes de fútbol amateur y profesional', 3, '924132')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('931101', 'Hipódromos', 3, '924140')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('931909', 'Otras actividades deportivas n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('931901', 'Promoción y organización de competencias deportivas', 3, '924150')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('854100', 'Enseñanza deportiva y recreativa', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('854100', 'Enseñanza deportiva y recreativa', 3, '924160')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('931909', 'Otras actividades deportivas n.c.p.', 3, '924190')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('799000', 'Otros servicios de reservas y actividades conexas (incluye venta de entradas para teatro, y otros)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('920090', 'Otras actividades de juegos de azar y apuestas n.c.p.', 3, '924910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('920010', 'Actividades de casinos de juegos', 3, '924920')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('932901', 'Gestión de salas de pool; gestión (explotación) de juegos electrónicos', 3, '924930')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('931102', 'Gestión de salas de billar; gestión de salas de bolos (bowling)', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('781000', 'Actividades de agencias de empleo', 3, '924940')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('932909', 'Otras actividades de esparcimiento y recreativas n.c.p.', 3, '924990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('592000', 'Actividades de grabación de sonido y edición de música', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('931909', 'Otras actividades deportivas n.c.p.', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('960100', 'Lavado y limpieza, incluida la limpieza en seco, de productos textiles y de piel', 3, '930100')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('960200', 'Peluquería y otros tratamientos de belleza', 3, '930200')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('960310', 'Servicios funerarios', 3, '930310')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('960320', 'Servicios de cementerios', 3, '930320')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('960310', 'Servicios funerarios', 3, '930330')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('960310', 'Servicios funerarios', 3, '930390')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('960320', 'Servicios de cementerios', 3, '')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('960902', 'Actividades de salones de masajes, baños turcos, saunas, servicio de baños públicos', 3, '930910')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('960909', 'Otras actividades de servicios personales n.c.p.', 3, '930990')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('970000', 'Actividades de los hogares como empleadores de personal doméstico', 3, '950001')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('949904', 'Consejo de administración de edificios y condominios', 3, '950002')"
   Call ExecSQL(DbMain, Q1)
   Q1 = "INSERT INTO CodActiv(Codigo, Descrip, Version, OldCodigo) VALUES ('990000', 'Actividades de organizaciones y órganos extraterritoriales', 3, '990000')"
   Call ExecSQL(DbMain, Q1)
      
End Sub

