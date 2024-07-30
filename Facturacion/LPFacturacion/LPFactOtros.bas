Attribute VB_Name = "LPFactOtros"
Option Explicit


Public Function ImportarEntidades(ByVal fname As String)
   Dim Fd As Long, Rc As Long, Q1 As String, Buf As String, l As Long, p As Long
   Dim QI As String, Aux As String, Rs As Recordset, i As Integer, r As Integer
   Dim AuxRut As String, NotValidRut As Boolean
   
   Dim RutEnt As String, CodEnt As String, NomEnt As String, DirEnt As String
   Dim RegEnt As Integer, ComuEnt As Integer, CiuEnt As String, TelEnt As String, FaxEnt As String
   Dim CodActEconEnt As String, DirPostEnt As String, ComuPostEnt As String, emailEnt As String
   Dim UrlEnt As String, ObsEnt As String, TipoEnt(MAX_ENTCLASIF) As Byte
   Dim Giro As String
   Dim SinClasif As Boolean
   Dim EsSupermercado As Boolean
   
   Fd = FreeFile
   Open fname For Input As #Fd
   If Err Then
      MsgErr fname
      ImportarEntidades = -Err
      Exit Function
   End If

   QI = "INSERT INTO Entidades (IdEmpresa, Rut, Codigo, Nombre, Direccion, Region, Comuna, Ciudad, Telefonos, Fax, Giro, DomPostal, ComPostal, Email, Web, Estado, Obs, Clasif0, Clasif1, Clasif2, Clasif3, Clasif4, Clasif5, EsSupermercado, NotValidRut)"
   Q1 = Q1 & "( VALUES (" & gEmpresa.Id ' 15 feb 2020
   
   Do Until EOF(Fd)
      Line Input #Fd, Buf
      l = l + 1
      'Debug.Print l & ")" & Buf
         
      p = 1
      Buf = Trim(Buf)
      
      If Buf = "" Then
         GoTo NextRec
      ElseIf l = 1 And InStr(1, Buf, "Nombre", vbTextCompare) Then
         GoTo NextRec
      End If
      
      NotValidRut = False
      AuxRut = Trim(NextField2(Buf, p))
      RutEnt = vFmtCID(AuxRut)
      If RutEnt = "0" Then
         RutEnt = AuxRut
         NotValidRut = True
      End If
      
      CodEnt = Trim(NextField2(Buf, p))
      NomEnt = Trim(NextField2(Buf, p))
            
      If RutEnt = "" Then
         If MsgBox1("Línea " & l & ": Falta el RUT de la entidad", vbExclamation + vbOKCancel) = vbCancel Then
            Exit Do
         End If
         GoTo NextRec
      End If
      
      If CodEnt = "" Then
         If MsgBox1("Línea " & l & ": Falta el código de la entidad", vbExclamation + vbOKCancel) = vbCancel Then
            Exit Do
         End If
         GoTo NextRec
      End If
      
      Q1 = "SELECT idEntidad FROM Entidades WHERE RUT='" & RutEnt & "' OR Codigo='" & CodEnt & "'" & " AND IdEmpresa = " & gEmpresa.Id
      Set Rs = OpenRs(DbMain, Q1)
      i = Rs.EOF
      Call CloseRs(Rs)
      If i = 0 Then
         If MsgBox1("Línea " & l & ": La entidad '" & NomEnt & "' (RUT=" & RutEnt & ", Código=" & CodEnt & ") ya existe.", vbExclamation + vbOKCancel) = vbCancel Then
            Exit Do
         End If
         GoTo NextRec
      End If
      
      DirEnt = Trim(NextField2(Buf, p))
      
      Aux = Trim(NextField2(Buf, p))
      If Aux <> "" Then
         Q1 = "SELECT Id, Codigo FROM Regiones WHERE Comuna='" & UCase(Aux) & "'"
         Set Rs = OpenRs(DbMain, Q1)
         If Rs.EOF = False Then
            RegEnt = vFld(Rs("Codigo"))
            ComuEnt = vFld(Rs("id"))
         Else
            If MsgBox1("Línea " & l & ": No se encontró la comuna '" & Aux & "' en la tabla de comunas.", vbExclamation + vbOKCancel) = vbCancel Then
               Exit Do
            End If
            RegEnt = -1
            ComuEnt = -1
         End If
         Call CloseRs(Rs)
      Else
         RegEnt = -1
         ComuEnt = -1
      End If

      CiuEnt = Trim(NextField2(Buf, p))
      TelEnt = Trim(NextField2(Buf, p))
      FaxEnt = Trim(NextField2(Buf, p))
'      CodActEconEnt = Trim(NextField2(Buf, p))
'      If CodActEconEnt <> "" Then
'         Q1 = "SELECT Codigo FROM CodActiv WHERE Codigo='" & CodActEconEnt & "'"
'         Set Rs = OpenRs(DbMain, Q1)
'
'         If Rs.EOF Then
'            MsgBox1 "Línea " & l & ": No se encontró la actividad económica '" & CodActEconEnt & "' en la tabla de actividades.", vbExclamation
'            CodActEconEnt = ""
'         End If
'         Call CloseRs(Rs)
'      End If
      Giro = Trim(NextField2(Buf, p))
      DirPostEnt = Trim(NextField2(Buf, p))
      ComuPostEnt = Trim(NextField2(Buf, p))
      emailEnt = Trim(NextField2(Buf, p))
      UrlEnt = Trim(NextField2(Buf, p))
      ObsEnt = Trim(NextField2(Buf, p))
      SinClasif = True

      'clasificación de la entidad
      For i = 0 To MAX_ENTCLASIF
         
         Aux = LCase(Trim(NextField2(Buf, p)))
         TipoEnt(i) = Abs(Aux = "x" Or Val(Aux) <> 0)
         If TipoEnt(i) <> 0 Then
            SinClasif = False
         End If
      Next i
      
      If SinClasif Then
         TipoEnt(0) = 1
      End If

     'Es supermercado?
      Aux = LCase(Trim(NextField2(Buf, p)))
      EsSupermercado = Abs(Aux = "x" Or Val(Aux) <> 0)

      Q1 = ",'" & RutEnt & "'"
      Q1 = Q1 & ",'" & CodEnt & "'"
      Q1 = Q1 & ",'" & NomEnt & "'"
      Q1 = Q1 & ",'" & DirEnt & "'"
      Q1 = Q1 & "," & RegEnt
      Q1 = Q1 & "," & ComuEnt
      Q1 = Q1 & ",'" & CiuEnt & "'"
      Q1 = Q1 & ",'" & TelEnt & "'"
      Q1 = Q1 & ",'" & FaxEnt & "'"
      Q1 = Q1 & ",'" & Giro & "'"
      Q1 = Q1 & ",'" & DirPostEnt & "'"
      Q1 = Q1 & ",'" & ComuPostEnt & "'"
      Q1 = Q1 & ",'" & emailEnt & "'"
      Q1 = Q1 & ",'" & UrlEnt & "'"
      Q1 = Q1 & "," & EE_ACTIVO
      Q1 = Q1 & ",'" & ObsEnt & "'"
      For i = 0 To MAX_ENTCLASIF
         Q1 = Q1 & "," & TipoEnt(i)
      Next i
      Q1 = Q1 & "," & Abs(EsSupermercado)
      Q1 = Q1 & "," & IIf(NotValidRut <> 0, 1, 0)
      Q1 = Q1 & " )"
      
      Debug.Print Q1

      Rc = ExecSQL(DbMain, QI & Q1)
      r = r + 1

NextRec:
   Loop

   Close #Fd

   ImportarEntidades = r

End Function

