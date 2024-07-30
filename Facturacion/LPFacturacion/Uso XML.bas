Attribute VB_Name = "UsoXML"

' ** Funcion que crea un xml

' funcion xAddTag está en VB50\xPam.bas


Public Function SaveLClase(LClase As LClase_t) As Long
   Dim xDoc As MSXML2.DOMDocument, xCurso As IXMLDOMNode, xProfe As IXMLDOMNode, xPat As IXMLDOMNode
   Dim xPersona As IXMLDOMNode, xICurso As IXMLDOMNode, Rc As Long, xHor As IXMLDOMNode
   Dim p As Integer, a As Integer, Sign As String, h As Integer, CodCurso As String, kNombre As String

   Set xDoc = New MSXML2.DOMDocument
   Set xICurso = xDoc.createElement("InfoCurso")  ' tag raiz
   Call xDoc.appendChild(xICurso)
   
   Call xAddTag(xDoc, xICurso, "Version", xVersion) ' version del xml, por si despues cambia
   
   Set xCurso = xAddTag(xDoc, xICurso, "Curso")
   
   Call xAddTag(xDoc, xCurso, "CodSence", LClase.CodSence)
   Call xAddTag(xDoc, xCurso, "IdLibro", LClase.idLibro)
   Call xAddTag(xDoc, xCurso, "NomLibro", LClase.NomLibro)
   Call xAddTag(xDoc, xCurso, "DescrLibro", LClase.DescrLibro)
   Call xAddTag(xDoc, xCurso, "NomInstit", LClase.NomInstit)
   Call xAddTag(xDoc, xCurso, "RutInstit", LClase.RutInstit)
   Call xAddTag(xDoc, xCurso, "FecConsult", LClase.FecConsult)
   Call xAddTag(xDoc, xCurso, "FecInicio", LClase.FecInicio)
   Call xAddTag(xDoc, xCurso, "FecTermino", LClase.FecTermino)
   Call xAddTag(xDoc, xCurso, "TipoPrograma", LClase.TipoPrograma)
   
   Call xAddTag(xDoc, xCurso, "IdAcciones", LClase.IdAcciones)
   Call xAddTag(xDoc, xCurso, "IdAccRechaz", LClase.IdAccRechaz)  ' 12 ago 2015
   
   Call xAddTag(xDoc, xCurso, "CodRegion", LClase.CodRegion)
   Call xAddTag(xDoc, xCurso, "CodComuna", LClase.CodComuna)
   Call xAddTag(xDoc, xCurso, "OTICs", LClase.OTICs)
   Call xAddTag(xDoc, xCurso, "Empresas", LClase.Empresas)


   ' Los Relatores
   For a = 0 To LClase.nRelatores - 1
      Set xPersona = xAddTag(xDoc, xCurso, "Relator")
      Call xAddTag(xDoc, xPersona, "Rut", LClase.Relatores(a).Rut)
      
      
      Call xAddTag(xDoc, xPersona, "IdAccion", LClase.Relatores(a).IdAccion)
      
      ' 7 aho 2014: Por si viene el nombre encriptado desde la tabla Sence.Personas
      If Len(LClase.Relatores(a).kNombre) > 0 Then
         Call xAddTag(xDoc, xPersona, "kNombre", LClase.Relatores(a).kNombre)
      ElseIf Len(LClase.Relatores(a).Nombre) > 0 Then
         Call xAddTag(xDoc, xPersona, "Nombre", LClase.Relatores(a).Nombre)
      End If
      
      LClase.Relatores(a).FActualiz = LClase.FecConsult
      
'      Call SavePersona(LClase.Relatores(a))
            
   Next a


   ' Los Alumnos de dejan en el mismo archivo del curso
   For a = 0 To LClase.nAlumnos - 1
   
      If a = 0 Then
         CodCurso = LClase.Alumnos(a).IdAccion
      ElseIf CodCurso <> LClase.Alumnos(a).IdAccion Then
         CodCurso = ""
      End If
   
      Set xPersona = xAddTag(xDoc, xCurso, "Alumno")
      If LClase.Alumnos(a).kRut <> "" Then
         Call xAddTag(xDoc, xPersona, "kRut", LClase.Alumnos(a).kRut)
      Else
         Call xAddTag(xDoc, xPersona, "Rut", LClase.Alumnos(a).Rut)
      End If
      
      ' 7 aho 2014: Por si viene el nombre encriptado desde la tabla Sence.Personas
      If Len(LClase.Alumnos(a).kNombre) > 0 Then
         Call xAddTag(xDoc, xPersona, "kNombre", LClase.Alumnos(a).kNombre)
      ElseIf Len(LClase.Alumnos(a).Nombre) > 0 Then
         Call xAddTag(xDoc, xPersona, "Nombre", LClase.Alumnos(a).Nombre)
      End If
      
      Call xAddTag(xDoc, xPersona, "IdAccion", LClase.Alumnos(a).IdAccion)
      Call xAddTag(xDoc, xPersona, "RutOTIC", LClase.Alumnos(a).RutOTIC)

      For p = 0 To LClase.Alumnos(a).nPat - 1
         Set xPat = xAddTag(xDoc, xPersona, "Pat")
         Call xAddTag(xDoc, xPat, "Dedo", LClase.Alumnos(a).Pat(p).Dedo)
         Call xAddTag(xDoc, xPat, "Patron", LClase.Alumnos(a).Pat(p).Patron)
         Call xAddTag(xDoc, xPat, "Tecno", LClase.Alumnos(a).Pat(p).Tecno)
         Call xAddTag(xDoc, xPat, "Fecha", LClase.Alumnos(a).Pat(p).Fecha)
      Next p

   Next a

'   Call xAddTag(xDoc, xICurso, "Sign", GenMd5(xCurso.xml, "Curso"))

   If LClase.TipoPrograma = TP_FRNTRIB Or LClase.TipoPrograma = TP_FORMTRAB Then
      CodCurso = LClase.CodSence
   End If
      
   Dim Fname As String
   
   Fname = gPathCursos & "Curso_" & LClase.TipoPrograma & "_" & CodCurso & "_" & LClase.idLibro & ".xml"
   
   Rc = UpdxLClase(xDoc, Fname)
   
   SaveLClase = Rc
   
   If Rc Then
      MsgLog "Error " & Rc & " al grabar el archivo" & vbCrLf & Fname & vbCrLf & Err.Description, vbExclamation
   Else
'      MsgBox "Se obtuvo la lista del curso '" & gCursos(gProfe.iCurso).NomCurso & "' con " & gCursos(gProfe.iCurso).nAlumnos & " alumnos.", vbInformation
   End If
   
   Set xDoc = Nothing
   
End Function

' ** Funcion que lee el xml

Public Function PasarLista(ByVal FnCurso As String, Hor As Horario_t, ByVal TipoProg As Integer) As Long
   Dim xDoc As MSXML2.DOMDocument, xICurso As IXMLDOMNode, xCurso As IXMLDOMNode, xSign  As IXMLDOMNode
   Dim Personas() As AutentiaLib13v22.Persona_t, Enrols() As AutentiaLib13v22.Enrol_t, iPers As Integer, iPers2 As Integer
   Dim Rc As Long, FObs As FrmObs, ObsRelator As String, i As Integer, j As Integer, a As Integer, Msg As String
   Dim Oper As AutentiaLib13v22.Persona_t, nAplic As Integer, nPersonas As Integer, IdAccRechaz As String
   Dim LClase As AutentiaLib13v22.LClase_t, xAlumno As IXMLDOMNode, xPat As IXMLDOMNode
   
   
   Set xDoc = New MSXML2.DOMDocument
   
   Rc = xDoc.Load(FnCurso)
   If Rc = False Then
      MsgLog "Error al cargar el archivo" & vbCrLf & FnCurso, vbExclamation
      PasarLista = 10
      Set xDoc = Nothing
      Exit Function
   End If

   Set xICurso = xDoc.selectSingleNode("InfoCurso")
   Set xCurso = xICurso.selectSingleNode("Curso")
   Set xSign = xICurso.selectSingleNode("Sign")

   If GenMd5(xCurso.xml, "Curso") <> xSign.Text Then
      MsgLog "El archivo está corrupto, debe volver a descargarlo." & vbCrLf & FnCurso, vbExclamation
      PasarLista = -1
      Exit Function
   End If

   gFrmMain.tx_Info = "Cargando Libro de Clases..."
   DoEvents

   LClase.idLibro = Val(xCurso.selectSingleNode("IdLibro").Text)
         
   If Not xCurso.selectSingleNode("IdAcciones") Is Nothing Then
      LClase.IdAcciones = xCurso.selectSingleNode("IdAcciones").Text
   End If

   If Not xCurso.selectSingleNode("IdAccRechaz") Is Nothing Then  ' 12 ago 2015
      LClase.IdAccRechaz = xCurso.selectSingleNode("IdAccRechaz").Text
   End If

   If Not xCurso.selectSingleNode("Conting") Is Nothing Then
      LClase.Conting = Val(xCurso.selectSingleNode("Conting").Text)
   End If

   If Not xCurso.selectSingleNode("TipoPrograma") Is Nothing Then
      LClase.TipoProg = Val(xCurso.selectSingleNode("TipoPrograma").Text)
   Else
      LClase.TipoProg = TipoProg
   End If
   
   For i = 0 To UBound(gTProgs)
      If gTProgs(i).Codigo = LClase.TipoProg Then
         LClase.TProg = gTProgs(i)
         Exit For
      End If
   Next i
   
   LClase.NomLibro = xCurso.selectSingleNode("NomLibro").Text
   LClase.DiaBloque = Hor.DiaBloque
   LClase.Bloque = Hor.Bloque
   LClase.Durac = Hor.Dur

   If Not xCurso.selectSingleNode("DescrLibro") Is Nothing Then
      LClase.DescrLibro = xCurso.selectSingleNode("DescrLibro").Text
   End If



   Oper.Rut = NormRut(gRutOper)
   Rc = LoadPersona(Oper, False) ' por si quiere enrolar
   Oper.TPers = TP_OPER

   iPers = 0
   
   If LClase.Conting = 0 Then
      Rc = LoadRelatores(xDoc, Personas, iPers)
      If Rc Then
         PasarLista = Rc
         Set xDoc = Nothing
         Exit Function
      End If
      
      iPers2 = iPers
   Else
      Rc = LoadAlumnos(xDoc, Personas, iPers, True)
      If Rc Then
         PasarLista = Rc
         Set xDoc = Nothing
         Exit Function
      End If
   End If
      
   Rc = LoadAlumnos(xDoc, Personas, iPers, False)
   If Rc Then
      PasarLista = Rc
      Set xDoc = Nothing
      Exit Function
   End If

   If gConectado And LClase.Conting = 0 And iPers > iPers2 Then ' 12 ago 2015: vemos si hay Acciones rechazadas
      ReDim Accion(20) As String
      a = -1
      If Len(LClase.IdAcciones) = 0 Then
         
         For i = iPers2 To iPers - 1
            For j = 0 To a
               If StrComp(Accion(j), Personas(i).IdAccion, vbTextCompare) = 0 Then
                  Exit For
               End If
            Next j

            If j > a Then ' no estaba ?
               If a + 1 > UBound(Accion) Then
                  ReDim Preserve Accion(a + 20)
               End If
               
               a = a + 1
               Accion(a) = Personas(i).IdAccion
            End If
               
         Next i

      Else
         i = 1
         Do While i < Len(LClase.IdAcciones)
            j = InStr(i, LClase.IdAcciones, ",", vbBinaryCompare)
      
            If a + 1 > UBound(Accion) Then
               ReDim Preserve Accion(a + 20)
            End If
            a = a + 1
            
            If j <= 0 Then
               Accion(a) = Trim(Mid(LClase.IdAcciones, i))
               Exit Do
            Else
               Accion(a) = Trim(Mid(LClase.IdAcciones, i, j - i))
            End If
            
            i = j + 1
            
         Loop

      End If

      If GetRechaz(Accion, IdAccRechaz) = 0 Then ' 12 ago 2015
         If StrComp(LClase.IdAccRechaz, IdAccRechaz, vbTextCompare) Then
         
            If xCurso.selectSingleNode("IdAccRechaz") Is Nothing Then
               Call xAddTag(xDoc, xCurso, "IdAccRechaz", IdAccRechaz)
            Else
               xCurso.selectSingleNode("IdAccRechaz").Text = IdAccRechaz
            End If
            
            Rc = UpdxLClase(xDoc, FnCurso)

'            MsgBox "Atención: cambio el estado de Acciones Rechazadas:" & vbCrLf & "antes: " & LClase.IdAccRechaz & vbCrLf & "ahora: " & IdAccRechaz & vbCrLf & "Comuníquese con el SENCE para solicitar más información.", vbInformation
            
            If Len(IdAccRechaz) > 1 Then
               Msg = "¡ ATENCIÓN !: cambió el estado de Acciones Rechazadas: (" & IdAccRechaz & ")."
               Msg = Msg & vbCrLf & "NOTA: Si el nuevo estado para el curso es 'RECHAZADO', tenga en cuenta  lo siguiente:"
               Msg = Msg & vbCrLf & "Una comunicación rechazada se debe a que el OTIC o la Empresa anuló el curso ante SENCE. Consulte con el OTIC o Empresa que corresponda, esto dependerá de cómo fue realizada la comunicación del curso ante SENCE."
               Msg = Msg & vbCrLf & "Sin embargo podrá registrar la asistencia de los participantes "

               MsgBox Msg, vbExclamation
            End If

            Call AddLog("Curso " & LClase.NomLibro & ", Acciones Rechazadas: antes: " & LClase.IdAccRechaz & ", ahora: " & IdAccRechaz)

            LClase.IdAccRechaz = IdAccRechaz
         End If

      End If

   End If

   ' Se guarda la fecha-hora que se inicia el paso de la lista
   LClase.FecClase = Format(Now, DATEFMT)
   LClase.iRelator = -1
   
   If Hor.Fname <> "" Then
      ' Recupera la asistencia del día, para continuar
      ' Ahora cargamos: FecClase, HInicio, ObsRelator y las Personas
      gFrmMain.tx_Info = "Cargando Asistencia Anterior..."
      DoEvents
      Rc = AplicAsist(LClase, Personas, nAplic, Hor.Fname)
      If nAplic > 0 Then
         LClase.bCont = 1
      End If
   End If

   ReDim Enrols(5)

   LClase.Caption = "Verificar Asistentes"

   gFrmMain.tx_Info = "Pasando lista..."
   DoEvents

   nPersonas = UBound(Personas)
   
   LClase.PathSence = gPathSence
   LClase.PathAsist = gPathAsist
   LClase.PathSent = gPathSent
   LClase.PathEnrol = gPathEnrol
   LClase.PathRuts = gPathRuts
   
   LClase.RutInstit = xCurso.selectSingleNode("RutInstit").Text
   LClase.NomInstit = xCurso.selectSingleNode("NomInstit").Text
   
   Err.Clear
   Rc = verify.ChkList(LClase, Oper, Personas, Enrols)
   
   If Rc <> 0 Then
      Set xDoc = Nothing
      PasarLista = Rc
      MsgLog "Error " & Err & ", " & Err.Description & vbCrLf & LClase.RcText, vbExclamation
      Exit Function
   End If

   If LClase.bMod Then ' si no se hizo nada, nada cambia
      gFrmMain.tx_Info = "Guardando nuevos participantes..."
      DoEvents
      
      If UBound(Personas) > nPersonas Then
      
         For i = nPersonas + 1 To UBound(Personas)
            
            If Personas(i).TPers = TP_ALUMNO Then
               Personas(i).kRut = EncriptRut(Personas(i).IdAccion, Personas(i).Rut)
               Personas(i).kNombre = EncriptNom(Personas(i).Rut, Personas(i).Nombre)
               
               Set xAlumno = xAddTag(xDoc, xCurso, "Alumno")
            Else
               Set xAlumno = xAddTag(xDoc, xCurso, "Relator")
            End If
            
            Call xAddTag(xDoc, xAlumno, "IdAccion", Personas(i).IdAccion)
                        
            If Personas(i).kRut <> "" Then
               Call xAddTag(xDoc, xAlumno, "kRut", Personas(i).kRut)
            Else
               Call xAddTag(xDoc, xAlumno, "Rut", Personas(i).Rut)
            End If
            

            If Personas(i).kNombre <> "" Then
               Call xAddTag(xDoc, xAlumno, "kNombre", Personas(i).kNombre)
            Else
               Call xAddTag(xDoc, xAlumno, "Nombre", Personas(i).Nombre)
            End If
            
            Call xAddTag(xDoc, xAlumno, "AddOffline", "1")
            
            If Personas(i).nPat > 0 Then
            
               Set xPat = xAddTag(xDoc, xAlumno, "Pat")
               Call xAddTag(xDoc, xPat, "Dedo", Personas(i).Pats(0).Dedo)
               Call xAddTag(xDoc, xPat, "Patron", Personas(i).Pats(0).Patron)
               Call xAddTag(xDoc, xPat, "Tecno", Personas(i).Pats(0).Tecno)
               Call xAddTag(xDoc, xPat, "Fecha", Personas(i).Pats(0).Fecha)
                        
            End If
              
         Next i
         
         xICurso.selectSingleNode("Sign").Text = GenMd5(xCurso.xml, "Curso")

         Err.Clear
         Call xDoc.Save(FnCurso)
      
         If Err.Number Then
            MsgLog "Error " & Err.Number & " al grabar el archivo" & vbCrLf & FnCurso & vbCrLf & Err.Description, vbExclamation
         End If
         
      End If
      
      gFrmMain.tx_Info = "Guardando Erolamientos..."
      DoEvents
         
'      For i = 0 To UBound(Enrols)
'         If Enrols(i).Rut <> "" And Enrols(i).Patron.Patron <> "" Then
'            Call SaveEnrol(Enrols(i))
'         End If
'      Next i
   End If

   Set xDoc = Nothing

   gFrmMain.tx_Info = ""

End Function


