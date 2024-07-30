Attribute VB_Name = "LPFactMain"
Option Explicit

Public Sub Main()
   Dim Rc As Integer, i As Integer, bDemo As Boolean
   Dim BoolIniEmpresa As Boolean
   Dim Msg As String, key As Long, PKey As Long
   Dim Q1 As String
   Dim Rs As Recordset
   Dim Usr As String
   Dim n As Integer
   Dim DbName As String
      
   Call PamInit
   Call PamRandomize

'   If W.InDesign And W.PcName = "PAM-PC" Then
'      Call AcpTest
'   End If


   ' pam: Nueva Instancia... despues de PamInit y PamRandomize
   key = GenInstanceKey()
   PKey = Val(GetCmdParam("i"))
   ' MsgBox "Key=" & Key & " ? " & PKey
   gNuevaInstancia = (key = PKey)

   gDebug = GetDebug()
   
   Call ChkSystem(True)

  'no se permite más de un usuario en un mismo equipo para evitar que algunos usuarios multipliquen sus licencias utilizando concección remota con Terminal Server
   If App.PrevInstance And gNuevaInstancia = False Then
      MsgBox "Esta aplicación ya se está ejecutando." & Chr(10) & "Use Alt+Tab hasta encontrarla", vbExclamation
      End
   End If

   Call InitFactComun

   Debug.Print "&" & Hex(FwVersion("", 0))
   If FwVersion("", 0) >= &H20004 Then ' *** por ahora
      Call FwInit("", 8725387) ' permite que el DLL funcione
   End If

   gDbPath = GetCmdParam("DbPath")
   If gDbPath = "" Then
      gDbPath = W.AppPath & "\Datos"
      If APP_DEMO Then
         gDbPath = W.AppPath & "\Datos" & "Demo"
      End If
   Else
      gDbPath = ReplaceStr(gDbPath, "%AppPath%", W.AppPath)
   End If
   Call AddLog("Main: gDbPath=[" & gDbPath & "]", 1)
   

   gImportPath = W.AppPath & "\Importar"
   gExportPath = W.AppPath & "\Exportar"
  
   gPdfDTE = W.AppPath & "\PdfDTE"
   
   On Error Resume Next
   MkDir gDbPath & "\Empresas"
'   MkDir gDbPath & "\Importar"
'   MkDir gDbPath & "\Exportar"
   MkDir gPdfDTE
   
'   RmDir (gDbPath & "\Importar")
'   RmDir (gDbPath & "\Exportar")
'
'   MkDir gImportPath
   MkDir gExportPath
   MkDir W.AppPath & "\Log"
   
   ' Verificación de Inscripción del equipo
   gLicFile = gDbPath & "\Empresas\InfoFact.cfg"
   gAppCode.Demo = True ' por defecto
   
   ' Esquema nuevo
   Call InscribPC  ' para poder ejecutar
   Call CheckInscPC  ' Nueva inscripción
      
   If APP_DEMO Then
      gAppCode.Demo = True
   End If
   
   If gAppCode.Demo Then
      gAppCode.NivProd = VER_DEMO
   End If
   
   If gAppCode.Demo Then
      Call AddLog("Version DEMO - " & APP_DEMO)
   End If

    
   gEmpresa.Rut = ""
   gEmpresa.Ano = 0
   
   DbName = gDbPath & "\" & BD_COMUNDTE
   
   If Not ExistFile(DbName) Then
      Call CopyFile(gDbPath & "\" & BD_COMUNDTEVACIA, DbName)
   End If
   
   If OpenDbAdmFact() = False Then
      End
   End If
      
   'si es APP_DEMO y la base no es de demo, pa' fuera para no dañar los datos con CorrigeBase
   
   If APP_DEMO Then
      
      'tiene más de 3 empersas con RUT distinto de 1, 2, 3
      
      Q1 = "SELECT Count(*) As N FROM Empresas WHERE RUT NOT IN ('1','2','3')"
      Set Rs = OpenRs(DbMain, Q1)
      
      If Not Rs.EOF Then
      
         If vFld(Rs("N")) > 0 Then
            MsgBox1 "La base de datos NO corresponde a la DEMO de LP Facturación." & vbCrLf & vbCrLf & gDbPath, vbCritical
            Call CloseRs(Rs)
            Call CloseDb(DbMain)
            End
         End If
         
      End If
      
      Call CloseRs(Rs)
      
   End If
      
   Call CorrigeBaseAdm
   
   Call AddLog("Main: a InscribPC", 2)
   
   If ContRegisterPc("", gCantLicencias) = False Then
      Call CloseDb(DbMain)
      End
   End If
   
   Usr = ContRegisteredUsr()
   
   Call AddDebug("Main: ContRegisteredUsr: '" & Usr & "'")
   
'   Q1 = "SELECT Pid FROM PcUsr WHERE PC = '" & ParaSQL(W.PcName) & "' AND Usr = '" & ParaSQL(W.UserName) & "'"
'   Set Rs = OpenRs(DbMain, Q1)
'
'   If Not Rs.EOF Then
'      Call AddDebug("Main: SELECT Pid: " & vFld(Rs("Pid")))
'   Else
'      Call AddDebug("Main: SELECT Pid: NULL")
'   End If
'
'   Call CloseRs(Rs)
      
      
   Call ReadOficina
   
   'Call CheckInscPC
   
   Call AddLog("Main: a FrmStart.show", 2)
   
   FrmStart.Show vbModeless
   DoEvents
   
   If gAppCode.Demo Then
      gAppCode.NivProd = VER_DEMO ' en modo DEMO mostramos todo lo del producto
      MsgBox1 "Este programa no está registrado y funcionará en modo DEMO." & vbCrLf & "Para registrarlo utilice la opción Ayuda>>Solicitar/Ingresar código de licencia.", vbInformation
   End If
   
   Sleep 500
      
   ' ****** 19 ago 2013 ************
   bDemo = gAppCode.Demo
   Call ReadPrimerUso   ' obtiene o registra el primero uso de la versión actual del programa
   
   If W.InDesign = False And APP_DEMO = False Then
   
      If FwChkActive(0) <> vbYes Then
   '      Call CloseDb(DbMain)
   '      End
      End If
   
   End If
   
   Call AddLog("Main: DM: " & bDemo & " => " & gAppCode.Demo & " - " & APP_DEMO)
   
   If bDemo <> gAppCode.Demo Then   ' por si paso de No demo a Si demo

      Call AddLog("Main: paso de No demo a Si demo")
      Call CloseDb(DbMain)

      Call AddLog("Main: a OpenDbAdmFact 2")
      If OpenDbAdmFact() = False Then
         Call AddLog("Main: falló OpenDbAdmFact 2")
         End
      End If
   End If
      
   Call AddLog("Main: a SetDbPath")
      
   Call SetDbPathFact(FrmStart.Drive1) ' se absolutiza el gDbPath
   Call AddLog("Main: de SetDbPath ==> gDbPath=[" & gDbPath & "]")
   
   gHRPath = GetCmdParam("HR")
   Call AddLog("Main: de GetCmdParam")
   If gHRPath = "" Then
      i = rInStr(gDbPath, "\")
      If i Then ' asumimos que viene al final viene "\Datos"
         gHRPath = Left(gDbPath, i) & ".."
      End If
      ' gHRPath = W.AppPath & "\.."
   End If
      
   Call AddDebug("Main: a IdUser")
   
gHRPath = "3C3F786D6C2076657273696F6E3D22312E30223F3E0D0A3C4454452076657273696F6E3D22312E302220786D6C6E733D22687474703A2F2F7777772E7369692E636C2F5369694474652220786D6C6E733A7873693D22687474703A2F2F7777772E77332E6F72672F323030312F584D4C536368656D612D696E7374616E6365223E3C446F63756D656E746F203E3C456E636162657A61646F3E3C4964446F633E3C5469706F4454453E35323C2F5469706F4454453E3C466F6C696F2F3E3C466368456D69733E323031392D30322D31383C2F466368456D69733E3C46636856656E633E323031392D30332D31373C2F46636856656E633E3C5469706F446573706163686F3E313C2F5469706F446573706163686F3E3C496E64547261736C61646F3E363C2F496E64547261736C61646F3E3C2F4964446F633E3C456D69736F723E3C525554456D69736F723E37383038393830302D333C2F525554456D"

gHRPath = gHRPath & "69736F723E3C527A6E536F633E525543414E545520532E412E3C2F527A6E536F633E3C4769726F456D69733E434F4E535452554343494F4E3C2F4769726F456D69733E3C41637465636F3E3435323031303C2F41637465636F3E3C4469724F726967656E3E50414E414D45524943414E4120535552204B4D203638303C2F4469724F726967656E3E3C436D6E614F726967656E3E5041445245204C41532043415341533C2F436D6E614F726967656E3E3C4369756461644F726967656E"

gHRPath = gHRPath & "3E54454D55434F3C2F4369756461644F726967656E3E3C2F456D69736F723E3C5265636570746F723E3C52555452656365703E31373835393133312D323C2F52555452656365703E3C527A6E536F6352656365703E5041424C4F2053414E444F56414C204D4F52414C45533C2F527A6E536F6352656365703E3C4769726F52656365703E504152544943554C41523C2F4769726F52656365703E3C436F6E746163746F2F3E3C44697252656365703E50414E414D45524943414E4120535552204B4D20363C2F44697252656365703E3C436D6E6152656365703E5041445245204C41532043415341533C2F436D6E6152656365703E3C43697564616452656365703E54454D55434F3C2F43697564616452656365703E3C2F5265636570746F7"

gHRPath = gHRPath & "23E3C546F74616C65733E3C4D6E744E65746F3E333030303C2F4D6E744E65746F3E3C4D6E744578653E303C2F4D6E744578653E3C546173614956413E31393C2F546173614956413E3C4956413E3537303C2F4956413E3C4D6E74546F74616C3E333537303C2F4D6E74546F74616C3E3C2F546F74616C65733E3C2F456E636162657A61646F3E3C446574616C6C653E3C4E726F4C696E4465743E313C2F4E726F4C696E4465743E3C4364674974656D3E3C54706F436F6469676F2F3E3C566C72436F6469676F3E5445524D4F54454D3C2F566C72436F6469676F3E3C2F4364674974656D3E3C4E6D624974656D3E455053203135204B4720353020"

gHRPath = gHRPath & "4D4D3C2F4E6D624974656D3E3C4473634974656D3E4D454449444120504C414E4348412031303030583530304D4D202D20554E49444144204445204D4544494441204D323C2F4473634974656D3E3C5174794974656D3E323C2F5174794974656D3E3C5072634974656D3E313530303C2F5072634974656D3E3C4D6F6E746F4974656D3E333030303C2F4D6F6E746F4974656D3E3C2F446574616C6C653E3C5265666572656E6369613E3C4E726F4C696E5265663E313C2F4E726F4C696E5265663E3C54706F446F635265663E3830313C2F54706F446F635265663E3C466F6C696F5265663E313C2F466F6C696F5265663E3C4663685265663E323031392D30322D31353C2F4663685265663E3C52617A6F6E5265662F3E3C2F5265666572656E6369613E3C2F446F63756D656E746F3E3C2F4454453E0D0A3C4461746F41646A756E746F206E6F6D6272653D224D61696C5265636570746F72223E7465726D6F74656D40727563616E74752E636C3C2F4461746F41646A756E746F3E3C4461746F7341646A756E746F733E3C4E6F6D62726544413E4D61696C5F5265636570746F723C2F4E6F6D62726544413E3C56616C6F7244413E7465726D6F746"

gHRPath = gHRPath & "56D40727563616E74752E636C3C2F56616C6F7244413E3C2F4461746F7341646A756E746F733E3C4461746F7341646A756E746F733E3C4E6F6D62726544413E5375626A6563745F4D61696C3C2F4E6F6D62726544413E3C56616C6F7244413E456E76696F20446F63756D656E746F20456C656374726F6E69636F3C2F56616C6F7244413E3C2F4461746F7341646A756E746F733E3C4461746F7341646A756E746F733E3C4E6F6D62726544413E4D61696C5F456D69736F723C2F4E6F6D62726544413E3C56616C6F7244413E6661637475726163696F6E40727563616E74752E636C3C2F56616C6F7244413E3C2F4461746F7341646A756E746F733E"
   
   Debug.Print Hex2Str(gHRPath)
   

   
   
   FrmIdUser.Show vbModal
   Call AddDebug("Main: después de IdUser, gUsuario.Rc=" & gUsuario.Rc & ", dbg=" & gDebug)
   
   If gUsuario.Rc = vbCancel Then
      Call ContUnregisterPc(2)
      Call CloseDb(DbMain)
      End
   End If
   
   'inicializamos arreglos básicos constantes y leemos archivo Ini
   Call AddDebug("Main: a IniLPFactura")
   Call IniLPFactura
   
   'creamos clases de impresión de grillas
   Call AddLog("Main: a CreatePrtFormats")
   Call CreatePrtFormats

   Q1 = "SELECT Count(*) as N FROM Empresas"
   If gAppCode.Demo Then
      Q1 = Q1 & " WHERE RUT IN ('1','2','3')" ' 10 mar 2021: se agrega este where para no mostrar una lista vacía
   End If
   Set Rs = OpenRs(DbMain, Q1)
   If Not Rs.EOF Then
      n = vFld(Rs("N"))
   End If
   Call CloseRs(Rs)
   
   Call AddLog("Main: a BoolIniEmpresa")
   If n > 0 Then
      
      'Mostramos la pantalla de selección de empresas según usuario
      BoolIniEmpresa = False
      
      Do While BoolIniEmpresa = False
      
'         If FrmSelEmpresas.FSelect() = vbCancel Then

'            Call ContUnregisterPc(3)
'            Call CloseDb(DbMain)
'            End
'         End If
         
         If FrmSelEmpresas.FSelect() = vbOK Then
            'Cerramos la DB LexContab
            Call CloseDb(DbMain)
   
            Call AddLog("Main: FrmSelEmpresas OK. A IniEmpresa")
               
            'Se abre la base de datos de la empresa y se inicializan sus datos básicos
            BoolIniEmpresa = IniEmpresa()
            
            Call AddLog("Main: IniEmpresa RC=" & BoolIniEmpresa)
            If BoolIniEmpresa = False Then
               If OpenDbAdmFact() = False Then
                  End
               End If
               
               'seteamos los datos de la empresa para clase de impresion de grillas
               Call SetPrtData
               
            End If
         Else
            Exit Do
         End If
      Loop
      
   End If
   
   Call AddDebug("Main: pasamos Loop IniEmpresa", 1)
   
   On Error GoTo 0
   
   FrmMain.Show vbModeless
'   Form1.Show vbModeless
   
   DoEvents
   Unload FrmStart
   
End Sub
'Revisar PAM
Private Sub ReadPrimerUso()
   Dim Q1 As String, Rs As Recordset, Rc As Long
   
   If gAppCode.Demo Then
      Exit Sub
   End If
   
   ' Primer uso de la version actual del programa
   Q1 = "SELECT Valor FROM ParamDTE WHERE Tipo='FUVER' And Codigo=" & W.FVersion
   Set Rs = OpenRs(DbMain, Q1)
   Q1 = ""
   If Rs.EOF Then
      gAppCode.FUsoVersion = Int(Now)
      Q1 = "INSERT INTO ParamDTE (Tipo, Codigo, Valor ) VALUES ( 'FUVER', " & W.FVersion & ", '" & gAppCode.FUsoVersion & "' )"
   Else
      gAppCode.FUsoVersion = Val(vFld(Rs("Valor")))
      
      If gAppCode.FUsoVersion <= 0 Or gAppCode.FUsoVersion > Int(Now) Then
         gAppCode.FUsoVersion = Int(Now)
         Q1 = "UPDATE ParamDTE  Set Valor='" & gAppCode.FUsoVersion & "' WHERE Tipo='FUVER' And Codigo=" & W.FVersion
      End If
   End If
   Call CloseRs(Rs)
   
   If Q1 <> "" Then
      Rc = ExecSQL(DbMain, Q1)
   End If
   
End Sub


Private Sub InitFactComun()
   Dim i As Integer
      
   gLPFactura = "T" & "h" & "o" & "ms" & "on" & " R" & "eu" & "t" & "er" & "s" & " F" & "ac" & "tu" & "ra" & "c" & "ió" & "n"

   App.HelpFile = W.AppPath & "\TRFactura.hlp"
   
   On Error Resume Next
   MkDir ("C:\TReuters")
   gIniFile = "C:\TReuters\TRFactura.ini"
   On Error GoTo 0
   
   gCfgFile = W.AppPath & "\TRFactura.cfg"
   gAdmUser = "administ"
   gValidRut = True

   gAppCode.Code = FAIRFACT_CODE
   gAppCode.Name = APP_NAME
   gAppCode.Title = App.Title
   gAppCode.TVerif = 1 ' LexContab
   gAppCode.emailSop = "soporte@thomsonreuters.com"
   gAppCode.emailInfo = "soporte@thomsonreuters.com"
   gAppCode.Contacto = "ThomsonReuters"
   gAppCode.TxInsc1 = "Gracias por probar nuestro producto. Si usted desea adquirirlo, por favor contáctese con Legal Publishing a los teléfonos (56-2) 510 5100, (56) 600 700 8000."
   gAppCode.TxInsc2 = "Para obtener el Código de Usuario: utilice el botón [Solicitud de Codigo de Usuario] o utilice el botón [Copiar datos] y luego péguelos en un email dirigido a " & gAppCode.emailSop & "."
   gAppCode.IniFile = gIniFile
   gAppCode.CfgFile = gCfgFile
   
   gAppCode.NivDef = VER_5EMP ' el más limitado
   
   ' pam: 13 dic 2010
   i = 0
   gAppCode.Nivel(i).Id = VER_ILIM
   gAppCode.Nivel(i).Desc = "Sin límite de empresas"
   
   i = i + 1
   gAppCode.Nivel(i).Id = VER_5EMP
   gAppCode.Nivel(i).Desc = "Máximo cinco empresas"
   
   i = i + 1
   gAppCode.Nivel(i).Id = 0
   gAppCode.Nivel(i).Desc = ""   ' fin de la lista
   
   
   'Call GetExtInfo("html", gHtmExt)
   Call GetExtInfo(".html", gHtmExt) 'PS le agregué ., porque no reconocia el iexplorer.exe
   
   Call FindPrinter(GetIniString(gIniFile, "Config", "Printer"), True)
   
   On Error Resume Next
   
End Sub

