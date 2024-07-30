VERSION 5.00
Begin VB.Form FrmImportEntTxt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar Entidades desde un Archivo de Texto"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Seleccionar Archivo"
      Height          =   1095
      Left            =   1320
      TabIndex        =   4
      Top             =   480
      Width           =   9315
      Begin VB.CommandButton Bt_Browse 
         Height          =   495
         Left            =   7800
         Picture         =   "FrmImportEntTxt.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Tx_FName 
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   420
         Width           =   7455
      End
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   360
      Picture         =   "FrmImportEntTxt.frx":056A
      ScaleHeight     =   615
      ScaleWidth      =   555
      TabIndex        =   3
      Top             =   540
      Width           =   555
   End
   Begin VB.CommandButton Bt_Importar 
      Caption         =   "Importar"
      Height          =   375
      Left            =   7680
      TabIndex        =   2
      Top             =   1800
      Width           =   1395
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   9240
      TabIndex        =   1
      Top             =   1800
      Width           =   1395
   End
   Begin VB.CommandButton Bt_ViewFmt 
      Caption         =   "Ver Formato"
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   1800
      Width           =   1395
   End
End
Attribute VB_Name = "FrmImportEntTxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lFName As String

Private Sub Bt_Cancelar_Click()
   Unload Me
End Sub

Private Sub Bt_Importar_Click()
   Me.MousePointer = vbHourglass
   Bt_Importar.Enabled = False
   DoEvents
   Call ImportarEntidades(lFName)
   Bt_Importar.Enabled = True
   Me.MousePointer = vbDefault
   
End Sub
Private Sub Bt_ViewFmt_Click()
   Dim Frm As FrmFmtImport
   
   Set Frm = New FrmFmtImport
   Call Frm.FViewEntidad
   Set Frm = Nothing

End Sub

Private Sub Bt_Browse_Click()

   FrmMain.Cm_ComDlg.CancelError = True
   FrmMain.Cm_ComDlg.Filename = ""
   FrmMain.Cm_ComDlg.InitDir = gImportPath
   FrmMain.Cm_ComDlg.Filter = "Archivos de Texto (*.txt)|*.txt"
   FrmMain.Cm_ComDlg.DialogTitle = "Seleccionar Archivo de Importación"
   FrmMain.Cm_ComDlg.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
 
   On Error Resume Next
   FrmMain.Cm_ComDlg.ShowOpen
   
   If Err = cdlCancel Then
      Exit Sub
   ElseIf Err Then
      MsgBox1 "Error " & Err & ", " & Error & NL & FrmMain.Cm_ComDlg.Filename, vbExclamation
      Exit Sub
   End If

   If FrmMain.Cm_ComDlg.Filename = "" Then
      Exit Sub
   End If
   Err.Clear
   
   lFName = FrmMain.Cm_ComDlg.Filename
   
   Tx_FName = lFName
   
   DoEvents
      
End Sub

Public Function ImportarEntidades(ByVal FName As String)
   Dim Fd As Long, Rc As Long, Q1 As String, Buf As String, l As Long, p As Long
   Dim QI As String, Aux As String, Rs As Recordset, i As Integer, r As Integer
   Dim AuxRut As String, NotValidRut As Boolean
   
   Dim RutEnt As String, CodEnt As String, NomEnt As String, DirEnt As String
   Dim RegEnt As Integer, ComuEnt As Integer, CiuEnt As String, TelEnt As String, FaxEnt As String
   Dim CodActEconEnt As String, DirPostEnt As String, ComuPostEnt As String, emailEnt As String
   Dim UrlEnt As String, ObsEnt As String, TipoEnt(MAX_ENTCLASIF) As Byte
   Dim Giro As String
   Dim SinClasif As Boolean
   Dim EsSupermercado As Integer
   
   Fd = FreeFile
   Open FName For Input As #Fd
   If Err Then
      MsgErr FName
      ImportarEntidades = -Err
      Exit Function
   End If

   QI = "INSERT INTO Entidades (Rut, Codigo, Nombre, Direccion, Region, Comuna, Ciudad, Telefonos, Fax, Giro, DomPostal, ComPostal, Email, Web, Estado, Obs, Clasif0, Clasif1, Clasif2, Clasif3, Clasif4, Clasif5, EsSupermercado, NotValidRut) VALUES"
   
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
         If MsgBox1("Línea " & l & ": Falta el RUT de la entidad." & vbCrLf & "¿Desea continuar?", vbExclamation Or vbYesNo) <> vbYes Then
            Exit Do
         End If
         GoTo NextRec
      End If
      
      If CodEnt = "" Then
         If MsgBox1("Línea " & l & ": Falta el código de la entidad." & vbCrLf & "¿Desea continuar?", vbExclamation Or vbYesNo) <> vbYes Then
            Exit Do
         End If
         GoTo NextRec
      End If
      
      Q1 = "SELECT idEntidad FROM Entidades WHERE RUT='" & RutEnt & "' OR Codigo='" & CodEnt & "'" & " AND IdEmpresa = " & gEmpresa.Id
      Set Rs = OpenRs(DbMain, Q1)
      i = Rs.EOF
      Call CloseRs(Rs)
      If i = 0 Then
         If MsgBox1("Línea " & l & ": La entidad '" & NomEnt & "' (RUT=" & RutEnt & ", Código=" & CodEnt & ") ya existe, no será incluído." & vbCrLf & "¿Desea continuar?", vbExclamation Or vbYesNo) <> vbYes Then
            Exit Do
         End If
         GoTo NextRec
      End If
      
      DirEnt = Trim(NextField2(Buf, p))
      
      Aux = Trim(NextField2(Buf, p))
      If Aux <> "" Then
         Q1 = "SELECT Id, Codigo FROM Regiones WHERE Comuna='" & ParaSQL(UCase(Aux)) & "'"
         Set Rs = OpenRs(DbMain, Q1)
         If Rs.EOF = False Then
            RegEnt = vFld(Rs("Codigo"))
            ComuEnt = vFld(Rs("id"))
         Else
            Call CloseRs(Rs)
            If MsgBox1("Línea " & l & ": No se encontró la comuna '" & Aux & "' en la tabla de comunas, no será asignada." & vbCrLf & "¿Desea continuar", vbExclamation Or vbYesNo) <> vbYes Then
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
      If emailEnt <> "" And ValidEmail(emailEnt) = False Then
         If MsgBox1("Línea " & l & ": El mail '" & emailEnt & "' es inválido, no será asignado." & vbCrLf & "¿Desea continuar", vbExclamation Or vbYesNo) <> vbYes Then
            Exit Do
         End If
         emailEnt = ""
      End If
           
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

      Q1 = " ( "
      Q1 = Q1 & "'" & RutEnt & "'"
      Q1 = Q1 & ",'" & ParaSQL(CodEnt) & "'"
      Q1 = Q1 & ",'" & ParaSQL(NomEnt) & "'"
      Q1 = Q1 & ",'" & ParaSQL(DirEnt) & "'"
      Q1 = Q1 & "," & RegEnt
      Q1 = Q1 & "," & ComuEnt
      Q1 = Q1 & ",'" & ParaSQL(CiuEnt) & "'"
      Q1 = Q1 & ",'" & ParaSQL(TelEnt) & "'"
      Q1 = Q1 & ",'" & ParaSQL(FaxEnt) & "'"
      Q1 = Q1 & ",'" & ParaSQL(Giro) & "'"
      Q1 = Q1 & ",'" & ParaSQL(DirPostEnt) & "'"
      Q1 = Q1 & ",'" & ParaSQL(ComuPostEnt) & "'"
      Q1 = Q1 & ",'" & ParaSQL(emailEnt) & "'"
      Q1 = Q1 & ",'" & ParaSQL(UrlEnt) & "'"
      Q1 = Q1 & "," & EE_ACTIVO
      Q1 = Q1 & ",'" & ParaSQL(ObsEnt) & "'"
      For i = 0 To MAX_ENTCLASIF
         Q1 = Q1 & "," & TipoEnt(i)
      Next i
      Q1 = Q1 & "," & EsSupermercado
      Q1 = Q1 & "," & IIf(NotValidRut <> 0, 1, 0)
      Q1 = Q1 & " )"
      
      Debug.Print Q1

      Rc = ExecSQL(DbMain, QI & Q1)
      r = r + 1

NextRec:
   Loop

   Close #Fd

   ImportarEntidades = r
   
   MsgBox1 "Se importaron " & r & " Entidades.", vbInformation

End Function

