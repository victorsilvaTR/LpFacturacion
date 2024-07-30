Attribute VB_Name = "MCorrigeBase"
Option Explicit

Private lDbVer As Integer
Private lUpdOK As Boolean
' Para hacer manteciones a ciertas tablas con manejo de versión
Public Sub CorrigeBase()
   
   On Error Resume Next
   
   lDbVer = 0
   lUpdOK = True
   
   If Not CorrigeBase_V1() Then
      Exit Sub
   End If
   
   If Not CorrigeBase_V2() Then        'agregada el 12 de oct 2016
      Exit Sub
   End If

   If Not CorrigeBase_V3() Then        'agregada el 14 de nov 2016
      Exit Sub
   End If

   If Not CorrigeBase_V4() Then        'agregada el 20 dic 2016
      Exit Sub
   End If


   If Not CorrigeBase_V5() Then        'agregada el 11 ene 2017
      Exit Sub
   End If

   If Not CorrigeBase_V6() Then        'agregada el 30 mar 2017
      Exit Sub
   End If

   If Not CorrigeBase_V7() Then        'agregada el 11 may 2017
      Exit Sub
   End If

   If Not CorrigeBase_V8() Then        'agregada el 25 jul 2018
      Exit Sub
   End If

   If Not CorrigeBase_V9() Then        'agregada el 8 ago 2018
      Exit Sub
   End If

   If Not CorrigeBase_V10() Then        'agregada el 11 sept 2018
      Exit Sub
   End If

   If Not CorrigeBase_V11() Then        'agregada el 26 feb 2019
      Exit Sub
   End If

   If Not CorrigeBase_V12() Then        'agregada el 8 mar 2019
      Exit Sub
   End If

   If Not CorrigeBase_V13() Then        'agregada el 13 mar 2019
      Exit Sub
   End If

   If Not CorrigeBase_V14() Then        'agregada el 10 abr 2019
      Exit Sub
   End If

   If Not CorrigeBase_V15() Then        'agregada el 15 abr 2019
      Exit Sub
   End If

   If Not CorrigeBase_V16() Then        'agregada el 6 may 2019
      Exit Sub
   End If

   If Not CorrigeBase_V17() Then        'agregada el 6 jun 2019
      Exit Sub
   End If

   If Not CorrigeBase_V18() Then        'agregada el 29 ago 2019
      Exit Sub
   End If

   If Not CorrigeBase_V19() Then        'agregada el 29 sep 2019
      Exit Sub
   End If
   
   If Not CorrigeBase_V20() Then        'agregada el 09 de ene 2023
      Exit Sub
   End If
   
   If Not CorrigeBase_V21() Then        'agregada el 09 de ene 2023
      Exit Sub
   End If

   If lDbVer > 22 Then
      MsgBox1 "¡ ATENCION !" & vbCrLf & vbCrLf & "La base de datos corresponde a una versión posterior de este programa." & vbCrLf & "Debe actualizar el programa antes de continuar, de lo contrario podría dañar la información..", vbCritical
      Call CloseDb(DbMain)
      End
   End If

End Sub

Public Function CorrigeBase_V21() As Boolean        'agregada el 10 abr 2019
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
  
   On Error Resume Next

   '--------------------- Versión 21 -----------------------------------

   If lDbVer = 21 Then
   
     Err.Clear
      
      'agrandemos el campo NotValidRut a DTE
      Set Tbl = DbMain.TableDefs("DTE")

      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("DetFormaPago", dbLong)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "DTE.DetFormaPago", vbExclamation
         lUpdOK = False
      End If
      
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("Vendedor", dbLong)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "DTE.Vendedor", vbExclamation
         lUpdOK = False
      End If
   
   
            
      Err.Clear
      
      'agregamos tabla DTERecibidos
      Set Tbl = New TableDef
      Tbl.Name = "DetFormaPago"
      
      Err.Clear
      Set fld = Tbl.CreateField("Id", dbLong)
      fld.Attributes = dbAutoIncrField ' Autonumber
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DetFormaPago.Id", vbExclamation
         lUpdOK = False
      End If
                        
                  
      Err.Clear
      Set fld = Tbl.CreateField("Descripcion", dbText, 100)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DetFormaPago.Descripcion", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("FormaPago", dbLong)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DetFormaPago.FormaPago", vbExclamation
         lUpdOK = False
      End If
      
      Err.Clear
      Set fld = Tbl.CreateField("Estado", dbLong)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DetFormaPago.Estado", vbExclamation
         lUpdOK = False
      End If
                  
      DbMain.TableDefs.Append Tbl
      If Err = 0 Then
         DbMain.TableDefs.Refresh
         
         Q1 = "CREATE UNIQUE INDEX Idi ON DetFormaPago (Id) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)
         
         Q1 = "CREATE UNIQUE INDEX IdxDFM ON DetFormaPago (Id, FormaPago, Estado)"
         Rc = ExecSQL(DbMain, Q1, False)
         
      ElseIf Err <> 3010 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Tabla DetFormaPago", vbExclamation
         lUpdOK = False

      End If
      
      
      Err.Clear
      
      'agregamos tabla Vendedor
      Set Tbl = New TableDef
      Tbl.Name = "Vendedor"
      
      Err.Clear
      Set fld = Tbl.CreateField("Rut", dbText, 12)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Vendedor.Rut", vbExclamation
         lUpdOK = False
      End If
                        
                  
      Err.Clear
      Set fld = Tbl.CreateField("Nombre", dbText, 100)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Vendedor.Nombre", vbExclamation
         lUpdOK = False
      End If
      
      Err.Clear
      Set fld = Tbl.CreateField("Codigo", dbLong)
      'fld.Attributes = dbAutoIncrField ' Autonumber
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Vendedor.Codigo", vbExclamation
         lUpdOK = False
      End If
      
      Err.Clear
      Set fld = Tbl.CreateField("Estado", dbLong)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Vendedor.Estado", vbExclamation
         lUpdOK = False
      End If
                  
      DbMain.TableDefs.Append Tbl
      If Err = 0 Then
         DbMain.TableDefs.Refresh
         
         Q1 = "CREATE UNIQUE INDEX Idv ON Vendedor (Rut) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)
         
         Q1 = "CREATE UNIQUE INDEX IdxVEN ON Vendedor (Rut, Codigo, Estado)"
         Rc = ExecSQL(DbMain, Q1, False)
         
      ElseIf Err <> 3010 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Tabla DetFormaPago", vbExclamation
         lUpdOK = False

      End If
      
           
      If lUpdOK Then
         lDbVer = 22
         Q1 = "UPDATE ParamEmpDTE SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V21 = lUpdOK

End Function

Public Function CorrigeBase_V20() As Boolean   'agregada 29 sep 2019
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 20 -----------------------------------

   If lDbVer = 20 And lUpdOK = True Then
   
      ' se updatea la tasa para la rebaja del IVA
      Q1 = "Update TipoValor Set Tasa = 65 Where Codigo = 8 And Atributo = 'SINUSO'"
      Call ExecSQL(DbMain, Q1, W.InDesign)
   
      
      If lUpdOK Then
         lDbVer = 21
         Q1 = "UPDATE ParamEmpDTE SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V20 = lUpdOK

End Function

Public Function CorrigeBase_V19() As Boolean   'agregada 29 sep 2019
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 19 -----------------------------------

   If lDbVer = 19 And lUpdOK = True Then
   
      ' El código DIN es definido por el SII y hay clientes que agregagon la empresa DIN con otro RUT
      Q1 = "UPDATE Entidades SET Codigo='DIN_SA' WHERE Rut<>'55555555' And Codigo='DIN'"
      Call ExecSQL(DbMain, Q1, W.InDesign)
   
      'agregamos entidad especial fija para Factura de Exportación, por s
      Q1 = "INSERT INTO Entidades( IdEmpresa, RUT, Codigo, Nombre, Ciudad, Clasif" & ENT_CLIENTE & ")VALUES(" & gEmpresa.Id & ", '" & RUT_DEFEXPORT & "', 'DIN', 'DIN', 'Santiago', 1)"
      Call ExecSQL(DbMain, Q1, W.InDesign)
      
      If lUpdOK Then
         lDbVer = 20
         Q1 = "UPDATE ParamEmpDTE SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V19 = lUpdOK

End Function

Public Function CorrigeBase_V18() As Boolean   'agregada 29 ago 2019
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim fld As Field
   Dim Rc As Long
   Dim DbComun As String


   On Error Resume Next

   '--------------------- Versión 18 -----------------------------------

   If lDbVer = 18 And lUpdOK = True Then

      'Actualizamos los códigos de actividad economica de acuerdo a nueva codificación SII de Nov 2018
'      Q1 = "UPDATE Empresa LEFT JOIN CodActiv ON Empresa.CodActEconom = CodActiv.OldCodigo SET Empresa.CodActEconom = CodActiv.Codigo, ActEconom = 0"
'      Call ExecSQL(DbMain, Q1)

      MsgBox1 "ATENCIÓN: El SII cambió la codificación de Actividades Económicas." & vbCrLf & vbCrLf & "Hemos actualizado el Código de Actividad Económica de la empresa seleccionada, de acuerdo a la nueva codificación entregada por el SII, respetando el esquema de conversión entregado por el mismo SII." & vbCrLf & vbCrLf & "Sin embargo, debe revisar que el código asignado se ajuste a la actividad de la empresa.", vbInformation + vbOKOnly
      
      If lUpdOK Then
         lDbVer = 19
         Q1 = "UPDATE ParamEmpDTE SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If

   End If

   CorrigeBase_V18 = lUpdOK

End Function

Public Function CorrigeBase_V17() As Boolean        'agregada el 7 jun 2019
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
  
   On Error Resume Next

   '--------------------- Versión 17 -----------------------------------

   If lDbVer = 17 Then
            
      Err.Clear
      
      'actualizamos el IdEntidad en el DTE cuando este no calza con el RUT asociado al DTE
      
      Q1 = "UPDATE DTE INNER JOIN Entidades ON DTE.Rut = Entidades.RUT AND DTE.IdEmpresa = Entidades.IdEMpresa SET "
      Q1 = Q1 & " DTE.IdEntidad = Entidades.IdEntidad "
      Q1 = Q1 & " WHERE DTE.IdEmpresa = " & gEmpresa.Id
      Call ExecSQL(DbMain, Q1, W.InDesign)
      
      If lUpdOK Then
         lDbVer = 18
         Q1 = "UPDATE ParamEmpDTE SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V17 = lUpdOK

End Function

Public Function CorrigeBase_V16() As Boolean        'agregada el 6 may 2019
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
  
   On Error Resume Next

   '--------------------- Versión 16 -----------------------------------

   If lDbVer = 16 Then
            
      Err.Clear
      
      'cambiamos nombre de campo Id por IdDTE en tabla DTERecibidos
      Set Tbl = DbMain.TableDefs("DTERecibidos")
             
      Err.Clear
      Tbl.Fields("Id").Name = "IdDTE"
      Tbl.Fields.Refresh
      
      Q1 = "DROP INDEX Idx ON DTERecibidos "
      Rc = ExecSQL(DbMain, Q1, False)

      Q1 = "CREATE UNIQUE INDEX Idx ON DTERecibidos (IdDTE) WITH PRIMARY"
      Rc = ExecSQL(DbMain, Q1, False)

      
      Set Tbl = Nothing

      If lUpdOK Then
         lDbVer = 17
         Q1 = "UPDATE ParamEmpDTE SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V16 = lUpdOK


End Function

Public Function CorrigeBase_V15() As Boolean        'agregada el 15 abr 2019
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
  
   On Error Resume Next

   '--------------------- Versión 15 -----------------------------------

   If lDbVer = 15 Then
            
      Err.Clear
      
      Q1 = "DROP INDEX IdxDTE ON DTERecibidos "
      Rc = ExecSQL(DbMain, Q1, False)

      Q1 = "CREATE UNIQUE INDEX IdxDTE ON DTERecibidos (IdEmpresa, CodDocSII, RUTEmisor, Folio)"
      Rc = ExecSQL(DbMain, Q1, False)

      'cambiamos nombre de campo Id por IdDTE en tabla DTERecibidos
      Set Tbl = DbMain.TableDefs("DTERecibidos")
             
      Err.Clear
      Tbl.Fields("Id").Name = "IdDTE"
      Tbl.Fields.Refresh
            
      Set Tbl = Nothing

    

      If lUpdOK Then
         lDbVer = 16
         Q1 = "UPDATE ParamEmpDTE SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V15 = lUpdOK


End Function

Public Function CorrigeBase_V14() As Boolean        'agregada el 10 abr 2019
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
  
   On Error Resume Next

   '--------------------- Versión 14 -----------------------------------

   If lDbVer = 14 Then
            
      Err.Clear
      
      'agregamos tabla DTERecibidos
      Set Tbl = New TableDef
      Tbl.Name = "DTERecibidos"
      
      Err.Clear
      Set fld = Tbl.CreateField("Id", dbLong)
      fld.Attributes = dbAutoIncrField ' Autonumber
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTERecibidos.Id", vbExclamation
         lUpdOK = False
      End If
                        
      Err.Clear
      Set fld = Tbl.CreateField("IdEmpresa", dbLong)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTERecibidos.IdEmpresa", vbExclamation
         lUpdOK = False
      End If
                        
      Err.Clear
      Set fld = Tbl.CreateField("TipoDoc", dbInteger)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTERecibidos.TipoDoc", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("TipoLib", dbInteger)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTERecibidos.TipoLib", vbExclamation
         lUpdOK = False
      End If
      
      Err.Clear
      Set fld = Tbl.CreateField("CodDocSII", dbText, 3)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTERecibidos.CodDocSII", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("Folio", dbLong)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTERecibidos.Folio", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("RUTEmisor", dbText, 12)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTERecibidos.RUTEmisor", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("RazonSocial", dbText, 100)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTERecibidos.RazonSocial", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("IdEntidad", dbLong)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTERecibidos.IdEntidad", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("RUTReceptor", dbText, 12)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTERecibidos.RUTReceptor", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("FPublicacion", dbDouble)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTERecibidos.FPublicacion", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("FEmision", dbLong)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTERecibidos.FEmision", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("Neto", dbDouble)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTERecibidos.Neto", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("Exento", dbDouble)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTERecibidos.Exento", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("IVA", dbDouble)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTERecibidos.IVA", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("Total", dbDouble)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTERecibidos.Total", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("Impuestos", dbDouble)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTERecibidos.Impuestos", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("TxtDetImpuestos", dbText, 250)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTERecibidos.TxtDetImpuestos", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("UrlDTE", dbText, 250)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTERecibidos.TxtDetImpuestos", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("FormaPago", dbInteger)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTERecibidos.TxtDetImpuestos", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("FVenc", dbLong)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTERecibidos.FVenc", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("FCesion", dbLong)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTERecibidos.FCesion", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("FRecepSII", dbLong)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTERecibidos.FRecepSII", vbExclamation
         lUpdOK = False
      End If
            
      DbMain.TableDefs.Append Tbl
      If Err = 0 Then
         DbMain.TableDefs.Refresh
         
         Q1 = "CREATE UNIQUE INDEX Idx ON DTERecibidos (Id) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)
         
         Q1 = "CREATE UNIQUE INDEX IdxDTE ON DTERecibidos (IdEmpresa, CodDocSII, RUT, Folio)"
         Rc = ExecSQL(DbMain, Q1, False)
         
         Q1 = "CREATE INDEX IdxTipo ON DTERecibidos (IdEmpresa, TipoLib, TipoDoc)"
         Rc = ExecSQL(DbMain, Q1, False)
         
         Q1 = "CREATE INDEX IdxFecha ON DTERecibidos (IdEmpresa, FEmision)"
         Rc = ExecSQL(DbMain, Q1, False)
         
      ElseIf Err <> 3010 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Tabla DTERecibidos", vbExclamation
         lUpdOK = False

      End If
      
           
      If lUpdOK Then
         lDbVer = 15
         Q1 = "UPDATE ParamEmpDTE SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V14 = lUpdOK

End Function


Public Function CorrigeBase_V13() As Boolean        'agregada el 13 mar 2019
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
  
   On Error Resume Next

   '--------------------- Versión 13 -----------------------------------

   If lDbVer = 13 Then
            
      Err.Clear
      
      'agrandemos el campo NotValidRut a DTE
      Set Tbl = DbMain.TableDefs("DTE")

      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("NotValidRUT", dbBoolean)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "DTE.NotValidRUT", vbExclamation
         lUpdOK = False
      End If
      
      'agrandemos el campo EsProducto a Productos
      Set Tbl = DbMain.TableDefs("Productos")

      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("EsProducto", dbBoolean)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "Productos.EsProducto", vbExclamation
         lUpdOK = False
      End If
      
      'actualizamos todos los productos con EsProducto = SI
      Q1 = "UPDATE Productos SET EsProducto = -1 "
      Call ExecSQL(DbMain, Q1)
            
      'agregamos tabla Vehiculos
      Set Tbl = New TableDef
      Tbl.Name = "Vehiculos"
      
      Err.Clear
      Set fld = Tbl.CreateField("Id", dbLong)
      fld.Attributes = dbAutoIncrField ' Autonumber
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Vehiculos.Id", vbExclamation
         lUpdOK = False
      End If
                        
      Err.Clear
      Set fld = Tbl.CreateField("IdEmpresa", dbLong)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Vehiculos.IdEmpresa", vbExclamation
         lUpdOK = False
      End If
                        
      Err.Clear
      Set fld = Tbl.CreateField("Patente", dbText, 8)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Vehiculos.Patente", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("IdTipoVehiculo", dbLong)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Vehiculos.IdTipoVehiculo", vbExclamation
         lUpdOK = False
      End If
      
      Err.Clear
      Set fld = Tbl.CreateField("Descrip", dbText, 80)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Vehiculos.Descrip", vbExclamation
         lUpdOK = False
      End If
      
      DbMain.TableDefs.Append Tbl
      If Err = 0 Then
         DbMain.TableDefs.Refresh
         
         Q1 = "CREATE UNIQUE INDEX Idx ON Vehiculos (Id) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)
         
         Q1 = "CREATE UNIQUE INDEX IdxPat ON Vehiculos (IdEmpresa, Patente)"
         Rc = ExecSQL(DbMain, Q1, False)
         
         Q1 = "CREATE INDEX IdxTipo ON Vehiculos (IdEmpresa, IdTipoVehiculo)"
         Rc = ExecSQL(DbMain, Q1, False)
         
      ElseIf Err <> 3010 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Tabla Vehiculos", vbExclamation
         lUpdOK = False

      End If
      
      'cambiamos nombre de campo en tabla DTEGuiaDesp: pasa de RutTrans a RutChofer
      Set Tbl = DbMain.TableDefs("DTEGuiaDesp")

      Err.Clear
      Tbl.Fields("RutTrans").Name = "RutChofer"
      Tbl.Fields.Refresh
      
      'Agregamos campo "NombreChofer"
      Err.Clear
      Set fld = Tbl.CreateField("NombreChofer", dbText, 30)
      Tbl.Fields.Append fld
      
      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTEGuiaDesp.NombreChofer", vbExclamation
         lUpdOK = False
      End If
      
      Set Tbl = Nothing
      
      
      'agregamos tabla Conductores
      Set Tbl = New TableDef
      Tbl.Name = "Conductores"
      
      Err.Clear
      Set fld = Tbl.CreateField("Id", dbLong)
      fld.Attributes = dbAutoIncrField ' Autonumber
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Conductores.Id", vbExclamation
         lUpdOK = False
      End If
                        
      Err.Clear
      Set fld = Tbl.CreateField("IdEmpresa", dbLong)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Conductores.IdEmpresa", vbExclamation
         lUpdOK = False
      End If
                        
      Err.Clear
      Set fld = Tbl.CreateField("RUTChofer", dbText, 12)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Conductores.RUTChofer", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("NombreChofer", dbText, 30)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Conductores.NombreChofer", vbExclamation
         lUpdOK = False
      End If
            
      DbMain.TableDefs.Append Tbl
      If Err = 0 Then
         DbMain.TableDefs.Refresh
         
         Q1 = "CREATE UNIQUE INDEX Idx ON Conductores (Id) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)
         
         Q1 = "CREATE UNIQUE INDEX IdxPat ON Conductores (IdEmpresa, RUTChofer)"
         Rc = ExecSQL(DbMain, Q1, False)
                  
      ElseIf Err <> 3010 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Tabla Conductores", vbExclamation
         lUpdOK = False

      End If
           
      If lUpdOK Then
         lDbVer = 14
         Q1 = "UPDATE ParamEmpDTE SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V13 = lUpdOK

End Function


Public Function CorrigeBase_V12() As Boolean        'agregada el 8 mar 2019
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
  
   On Error Resume Next

   '--------------------- Versión 12 -----------------------------------

   If lDbVer = 12 Then
            
      Err.Clear
      
      'actualizamos clasificación de entidad especial fija para Factura de Exportación
      Q1 = "UPDATE Entidades SET Clasif" & ENT_CLIENTE & " = 1 WHERE Codigo = '" & ENTIMP_RSOCIAL & "'"  'es el mismo apara importación y exportación
      Call ExecSQL(DbMain, Q1)
         

      If lUpdOK Then
         lDbVer = 13
         Q1 = "UPDATE ParamEmpDTE SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V12 = lUpdOK

End Function

Public Function CorrigeBase_V11() As Boolean        'agregada el 26 feb 2019
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
  
   On Error Resume Next

   '--------------------- Versión 11 -----------------------------------

   If lDbVer = 11 Then
            
      Err.Clear
         
      'agregamos tabla DTEGuiaDesp con los antecedentes de una factura de Guía de Despacho
      Set Tbl = New TableDef
      Tbl.Name = "DTEGuiaDesp"
      
      Err.Clear
      Set fld = Tbl.CreateField("IdDTEGuiaDesp", dbLong)
      fld.Attributes = dbAutoIncrField ' Autonumber
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTEGuiaDesp.IdDDTEGuiaDesp", vbExclamation
         lUpdOK = False
      End If
                        
      Err.Clear
      Set fld = Tbl.CreateField("IdDTE", dbLong)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTEGuiaDesp.IdDTE", vbExclamation
         lUpdOK = False
      End If
                        
      Err.Clear
      Set fld = Tbl.CreateField("IdEmpresa", dbLong)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTEGuiaDesp.IdEmpresa", vbExclamation
         lUpdOK = False
      End If
                        
      Err.Clear
      Set fld = Tbl.CreateField("Patente", dbText, 8)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTEGuiaDesp.Patente", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("RutTrans", dbText, 12)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTEGuiaDesp.RutTrans", vbExclamation
         lUpdOK = False
      End If
                        
      
      DbMain.TableDefs.Append Tbl
      If Err = 0 Then
         DbMain.TableDefs.Refresh
         
         Q1 = "CREATE UNIQUE INDEX Idx ON DTEGuiaDesp (IdDTEGuiaDesp) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)
         
         Q1 = "CREATE UNIQUE INDEX IdxDTE ON DTEGuiaDesp (IdDTE)"
         Rc = ExecSQL(DbMain, Q1, False)
                  
      ElseIf Err <> 3010 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Tabla DTEGuiaDesp", vbExclamation
         lUpdOK = False

      End If

      If lUpdOK Then
         lDbVer = 12
         Q1 = "UPDATE ParamEmpDTE SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
      
   End If
   
   CorrigeBase_V11 = lUpdOK

End Function


Public Function CorrigeBase_V10() As Boolean        'agregada el 11 sept. 2018
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
  
   On Error Resume Next

   '--------------------- Versión 10 -----------------------------------

   If lDbVer = 10 Then
            
      Err.Clear
      
      'agregamos entidad especial fija para Factura de Exportación
      Q1 = "INSERT INTO Entidades( IdEmpresa, RUT, Codigo, Nombre, Ciudad, Clasif" & ENT_CLIENTE & ")VALUES(" & gEmpresa.Id & ", '" & RUT_DEFEXPORT & "', 'DIN', 'DIN', 'Santiago', 1)"
      Call ExecSQL(DbMain, Q1, W.InDesign)
         
      'agregamos tabla DTEFactExp con los antecedentes de una factura de exportación
      Set Tbl = New TableDef
      Tbl.Name = "DTEFactExp"
      
      Err.Clear
      Set fld = Tbl.CreateField("IdDTEFactExp", dbLong)
      fld.Attributes = dbAutoIncrField ' Autonumber
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTEFactExp.IdDTEFactExp", vbExclamation
         lUpdOK = False
      End If
                        
      Err.Clear
      Set fld = Tbl.CreateField("IdDTE", dbLong)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTEFactExp.IdDTE", vbExclamation
         lUpdOK = False
      End If
                        
      Err.Clear
      Set fld = Tbl.CreateField("IdEmpresa", dbLong)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTEFactExp.IdEmpresa", vbExclamation
         lUpdOK = False
      End If
                        
      Err.Clear
      Set fld = Tbl.CreateField("CodIndServicio", dbText, 2)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTEFactExp.CodIndServicio", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("CodPais", dbText, 4)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTEFactExp.CodPais", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("CodPuertoEmbarque", dbText, 4)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTEFactExp.CodPuertoEmbarque", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("CodPuertoDesembarque", dbText, 4)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTEFactExp.CodPuertoDesembarque", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("CodMoneda", dbText, 3)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTEFactExp.CodMoneda", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("TipoCambioPesos", dbSingle)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTEFactExp.TipoCambioPesos", vbExclamation
         lUpdOK = False
      End If
                        
      Err.Clear
      Set fld = Tbl.CreateField("CodModVenta", dbText, 2)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTEFactExp.CodModVenta", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("CodClauCompraVenta", dbText, 2)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTEFactExp.CodClauCompraVenta", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("CodViaTransporte", dbText, 2)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTEFactExp.CodViaTransporte", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("TotalBultos", dbDouble)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTEFactExp.TotalBultos", vbExclamation
         lUpdOK = False
      End If
            
      Err.Clear
      Set fld = Tbl.CreateField("TotalClauVenta", dbDouble)
      Tbl.Fields.Append fld

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "DTEFactExp.TotalClauVenta", vbExclamation
         lUpdOK = False
      End If
            
      
      DbMain.TableDefs.Append Tbl
      If Err = 0 Then
         DbMain.TableDefs.Refresh
         
         Q1 = "CREATE UNIQUE INDEX Idx ON DTEFactExp (IdDTEFactExp) WITH PRIMARY"
         Rc = ExecSQL(DbMain, Q1, False)
         
         Q1 = "CREATE UNIQUE INDEX IdxDTE ON DTEFactExp (IdDTE)"
         Rc = ExecSQL(DbMain, Q1, False)
         
      ElseIf Err <> 3010 Then ' ya existe
         MsgBox1 "Error " & Err & ", " & Error & vbLf & "Tabla DTEFactExp", vbExclamation
         lUpdOK = False

      End If
      
      If lUpdOK Then
         lDbVer = 11
         Q1 = "UPDATE ParamEmpDTE SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   

   End If

   CorrigeBase_V10 = lUpdOK

End Function


Public Function CorrigeBase_V9() As Boolean        'agregada el 9 ago 2018
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
  
   On Error Resume Next

   '--------------------- Versión 9 -----------------------------------

   If lDbVer = 9 Then
            
      Err.Clear
      
      'agrandemos el campo Email de la entidad
'      Call AlterField(DbMain, "Entidades", "Email", dbText, 100)
      
      'agregamos campo IdEmpresa a ParamEmpDTE
      Set Tbl = DbMain.TableDefs("ParamEmpDTE")

      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("IdEmpresa", dbLong)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "ParamEmpDTE.IdEmpresa", vbExclamation
         lUpdOK = False
      End If
      
      Q1 = "UPDATE ParamEmpDTE SET IdEmpresa = " & gEmpresa.Id
      Call ExecSQL(DbMain, Q1)
      
      Q1 = "UPDATE Contactos SET IdEmpresa = " & gEmpresa.Id
      Call ExecSQL(DbMain, Q1)

      'agregamos campo ObsDTE a DTE
      Set Tbl = DbMain.TableDefs("DTE")

      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("ObsDTE", dbText, 100)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "DTE.ObsDTE", vbExclamation
         lUpdOK = False
      End If
      If lUpdOK Then
         lDbVer = 10
         Q1 = "UPDATE ParamEmpDTE SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V9 = lUpdOK

End Function


Public Function CorrigeBase_V8() As Boolean        'agregada el 25 jul 2018
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
  
   On Error Resume Next

   '--------------------- Versión 8 -----------------------------------

   If lDbVer = 8 Then
            
      Err.Clear
      
      'agrandamos campos de Entidades
'      Call AlterField(DbMain, "Entidades", "Nombre", dbText, 100)
'      Call AlterField(DbMain, "Entidades", "Giro", dbText, 80)
      
      If lUpdOK Then
         lDbVer = 9
         Q1 = "UPDATE ParamEmpDTE SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V8 = lUpdOK

End Function


Public Function CorrigeBase_V7() As Boolean        'agregada el 11 may 2017
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
  
   On Error Resume Next

   '--------------------- Versión 7 -----------------------------------

   If lDbVer = 7 Then
            
      Err.Clear
      
      Set Tbl = DbMain.TableDefs("Entidades")

      'agregamos campo EntRelacionada a Entidades
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("EntRelacionada", dbBoolean)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "Entidades.EntRelacionada", vbExclamation
         lUpdOK = False
      End If

      If lUpdOK Then
         lDbVer = 8
         Q1 = "UPDATE ParamEmpDTE SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V7 = lUpdOK

End Function


Public Function CorrigeBase_V6() As Boolean        'agregada el 30 mar 2017
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
  
   On Error Resume Next

   '--------------------- Versión 6 -----------------------------------

   If lDbVer = 6 Then
            
      Err.Clear
      
      Set Tbl = DbMain.TableDefs("Empresa")

      'agregamos campo ObsDTE
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("ObsDTE", dbText, 100)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "Empresa.ObsDTE", vbExclamation
         lUpdOK = False
      End If

      If lUpdOK Then
         lDbVer = 7
         Q1 = "UPDATE ParamEmpDTE SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V6 = lUpdOK

End Function


Public Function CorrigeBase_V5() As Boolean        'agregada el 11 ene 2017
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
  
   On Error Resume Next

   '--------------------- Versión 5 -----------------------------------

   If lDbVer = 5 Then
            
      Err.Clear
      
      Set Tbl = DbMain.TableDefs("DTE")

      'agregamos campo FechaVenc
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("FechaVenc", dbLong)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "DTE.FechaVenc", vbExclamation
         lUpdOK = False
      End If

      'agregamos campo Traslado
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("FormaDePago", dbByte)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "DTE.FormaDePago", vbExclamation
         lUpdOK = False
      End If

      If lUpdOK Then
         lDbVer = 6
         Q1 = "UPDATE ParamEmpDTE SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V5 = lUpdOK

End Function


Public Function CorrigeBase_V4() As Boolean        'agregada el 20 dic 2016
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
  
   On Error Resume Next

   '--------------------- Versión 4 -----------------------------------

   If lDbVer = 4 Then
            
      Err.Clear
      
      Set Tbl = DbMain.TableDefs("DTE")

      'agregamos campo TipoDespacho
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("TipoDespacho", dbInteger)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "DTE.TipoDespacho", vbExclamation
         lUpdOK = False
      End If

      'agregamos campo Traslado
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("Traslado", dbInteger)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "DTE.Traslado", vbExclamation
         lUpdOK = False
      End If

      If lUpdOK Then
         lDbVer = 5
         Q1 = "UPDATE ParamEmpDTE SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V4 = lUpdOK

End Function


Public Function CorrigeBase_V3() As Boolean        'agregada el 14 de nov 2016
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
  
   On Error Resume Next

   '--------------------- Versión 3 -----------------------------------

   If lDbVer = 3 Then
            
      Err.Clear
      
      Set Tbl = DbMain.TableDefs("DTE")

      'agregamos campo UrlDTE
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("UrlDTE", dbText, 250)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "DTE.UrlDTE", vbExclamation
         lUpdOK = False
      End If

      If lUpdOK Then
         lDbVer = 4
         Q1 = "UPDATE ParamEmpDTE SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V3 = lUpdOK

End Function


Public Function CorrigeBase_V2() As Boolean        'agregada el 12 de oct 2016
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
  
   On Error Resume Next

   '--------------------- Versión 2 -----------------------------------

   If lDbVer = 2 Then
            
      Err.Clear
      
      Set Tbl = DbMain.TableDefs("DetDTE")

      'agregamos campo Exento
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("MontoImpAdic", dbDouble)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "DetDTE.MontoImpAdic", vbExclamation
         lUpdOK = False
      End If

      If lUpdOK Then
         lDbVer = 3
         Q1 = "UPDATE ParamEmpDTE SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V2 = lUpdOK

End Function



Public Function CorrigeBase_V1() As Boolean
   Dim Q1 As String
   Dim Tbl As TableDef
   Dim fld As Field
   Dim Rc As Integer
   Dim Rs As Recordset
  
   On Error Resume Next

   Q1 = "SELECT Valor FROM ParamEmpDTE WHERE Tipo = 'DBVER'"
   Set Rs = OpenRs(DbMain, Q1)
   If Rs.EOF Then
      Call CloseRs(Rs)
      Q1 = "INSERT INTO ParamEmpDTE (Tipo, Codigo, Valor) VALUES ('DBVER', 0, '0')"
      Call ExecSQL(DbMain, Q1)
      lDbVer = 1
   Else
      lDbVer = Val(vFld(Rs("Valor")))
   End If

   Call CloseRs(Rs)
   
   '--------------------- Versión 1 -----------------------------------

   If lDbVer = 1 Then
            
      Err.Clear
      
      Set Tbl = DbMain.TableDefs("DTE")

      'agregamos campo Exento
      Err.Clear
      Tbl.Fields.Append Tbl.CreateField("Exento", dbDouble)

      If Err = 0 Then
         Tbl.Fields.Refresh
      ElseIf Err <> 3191 Then ' ya existe
         MsgBeep vbExclamation
         MsgBox "Error " & Err & ", " & Error & vbLf & "DTE.Exento", vbExclamation
         lUpdOK = False
      End If

      If lUpdOK Then
         lDbVer = 2
         Q1 = "UPDATE ParamEmpDTE SET Valor=" & lDbVer & " WHERE Tipo='DBVER'"
         Call ExecSQL(DbMain, Q1)
      End If
   
   End If
   
   CorrigeBase_V1 = lUpdOK

End Function
