VERSION 5.00
Begin VB.Form FrmImportProd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Importar Productos y Servicios"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   11430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Bt_ViewFmt 
      Caption         =   "Ver Formato"
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1860
      Width           =   1395
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   9480
      TabIndex        =   4
      Top             =   1860
      Width           =   1395
   End
   Begin VB.CommandButton Bt_Importar 
      Caption         =   "Importar"
      Height          =   375
      Left            =   7920
      TabIndex        =   3
      Top             =   1860
      Width           =   1395
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   600
      Picture         =   "FrmImportProd.frx":0000
      ScaleHeight     =   615
      ScaleWidth      =   555
      TabIndex        =   6
      Top             =   600
      Width           =   555
   End
   Begin VB.Frame Frame2 
      Caption         =   "Seleccionar Archivo"
      Height          =   1095
      Left            =   1560
      TabIndex        =   5
      Top             =   540
      Width           =   9315
      Begin VB.TextBox Tx_FName 
         Height          =   315
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   420
         Width           =   7455
      End
      Begin VB.CommandButton Bt_Browse 
         Height          =   495
         Left            =   7800
         Picture         =   "FrmImportProd.frx":058D
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "FrmImportProd"
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

   If lFName = "" Then
      MsgBox1 "Debe seleccionar un archivo usando el botón Buscar.", vbExclamation
      Exit Sub
   End If
   
   Me.MousePointer = vbHourglass
   Call ImportarProd(lFName)
   Me.MousePointer = vbDefault
   
End Sub

Private Sub Bt_ViewFmt_Click()
   Dim Frm As FrmFmtImport
   
   Set Frm = New FrmFmtImport
   Call Frm.FViewProducto
   Set Frm = Nothing
   
End Sub

Private Sub Bt_Browse_Click()

   FrmMain.Cm_ComDlg.CancelError = True
   FrmMain.Cm_ComDlg.FileName = ""
   FrmMain.Cm_ComDlg.InitDir = gImportPath
   FrmMain.Cm_ComDlg.Filter = "Archivos de Texto (*.txt)|*.txt"
   FrmMain.Cm_ComDlg.DialogTitle = "Seleccionar Archivo de Importación"
   FrmMain.Cm_ComDlg.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
 
   On Error Resume Next
   FrmMain.Cm_ComDlg.ShowOpen
   
   If Err = cdlCancel Then
      Exit Sub
   ElseIf Err Then
      MsgBox1 "Error " & Err & ", " & Error & NL & FrmMain.Cm_ComDlg.FileName, vbExclamation
      Exit Sub
   End If

   If FrmMain.Cm_ComDlg.FileName = "" Then
      Exit Sub
   End If
   Err.Clear
   
   lFName = FrmMain.Cm_ComDlg.FileName
   
   Tx_FName = lFName
   
   DoEvents
      
End Sub


Public Function ImportarProd(ByVal FName As String)
   Dim Fd As Long, Rc As Long, Q1 As String, Buf As String, l As Long, p As Long
   Dim QI As String, Aux As String, Rs As Recordset, i As Integer, r As Integer, YaExiste As Boolean
   Dim TipoCod As String, Codigo As String, Producto As String, UMedida As String, Precio As Double, EsProd As String, Obs As String
   
   If FName = "" Then
      Exit Function
   End If
   
   Fd = FreeFile
   Open FName For Input As #Fd
   If Err Then
      MsgErr FName
      ImportarProd = -Err
      Exit Function
   End If

   QI = "INSERT INTO Productos (IdEmpresa, TipoCod, CodProd, Producto, UMedida, Precio, EsProducto, Obs ) VALUES (" & gEmpresa.Id
   
   Do Until EOF(Fd)
      Line Input #Fd, Buf
      l = l + 1
      'Debug.Print l & ")" & Buf
         
      p = 1
      Buf = Trim(Buf)
      
      If Buf = "" Then
         GoTo NextRec
      ElseIf l = 1 And InStr(1, Buf, "Producto", vbTextCompare) Then
         GoTo NextRec
      End If
            
      TipoCod = Trim(NextField2(Buf, p))
      Codigo = Trim(NextField2(Buf, p))
                  
      Producto = Trim(NextField2(Buf, p))
      UMedida = Trim(NextField2(Buf, p))
      Precio = vFmt(NextField2(Buf, p))
      EsProd = NextField2(Buf, p)
      Obs = NextField2(Buf, p)
      
      If Producto = "" Then
         If MsgBox1("Línea " & l & ": Falta ingresar el nombre del producto o servicio.", vbExclamation + vbOKCancel) = vbCancel Then
            Exit Do
         End If
         GoTo NextRec
      End If

      If Precio = 0 Then
         If MsgBox1("Línea " & l & ": Falta ingresar el precio del producto o servicio.", vbExclamation + vbOKCancel) = vbCancel Then
            Exit Do
         End If
         GoTo NextRec
      End If
      
      EsProd = UCase(EsProd)
      
      If EsProd <> "SI" And EsProd <> "NO" Then
         If MsgBox1("Línea " & l & ": Debe indicar si la información ingresada corresponde a un producto o a un servicio.", vbExclamation + vbOKCancel) = vbCancel Then
            Exit Do
         End If
         GoTo NextRec
      End If
      
      If TipoCod <> "" And Codigo <> "" Then
         Q1 = "SELECT IdProducto FROM Productos WHERE TipoCod = '" & TipoCod & "' AND CodProd = '" & Codigo & "' AND IdEmpresa = " & gEmpresa.Id
         Set Rs = OpenRs(DbMain, Q1)
         
         If Not Rs.EOF Then
            YaExiste = True
         End If
         Call CloseRs(Rs)
                 
      ElseIf Codigo <> "" Then
         Q1 = "SELECT IdProducto FROM Productos WHERE (TipoCod = ' ' OR TipoCod IS NULL) AND CodProd = '" & Codigo & "' AND IdEmpresa = " & gEmpresa.Id
         Set Rs = OpenRs(DbMain, Q1)
         
         If Not Rs.EOF Then
            YaExiste = True
         End If
         Call CloseRs(Rs)
         
      Else
         Q1 = "SELECT IdProducto FROM Productos WHERE Producto = '" & Producto & "' AND IdEmpresa = " & gEmpresa.Id
         Set Rs = OpenRs(DbMain, Q1)
         
         If Not Rs.EOF Then
            YaExiste = True
         End If
         Call CloseRs(Rs)

      End If
      
      If YaExiste Then
         If MsgBox1("Línea " & l & ": Este producto o servicio ya existe.", vbExclamation + vbOKCancel) = vbCancel Then
            Exit Do
         End If
         GoTo NextRec
      End If
         
            
      Q1 = QI & ",'" & TipoCod & "'"
      Q1 = Q1 & ",'" & Codigo & "'"
      Q1 = Q1 & ",'" & Producto & "'"
      Q1 = Q1 & ",'" & UMedida & "'"
      Q1 = Q1 & "," & str(Precio)
      Q1 = Q1 & "," & ValSiNo(EsProd)
      Q1 = Q1 & ",'" & ParaSQL(Left(Obs, 255)) & "'"
      Q1 = Q1 & " )"
      
      Debug.Print Q1

      Rc = ExecSQL(DbMain, Q1)
      r = r + 1

NextRec:
   Loop

   Close #Fd

   MsgBox1 "Se importaron " & r & " productos.", vbInformation
   
   ImportarProd = r

End Function

