VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmMantVendedor 
   Caption         =   "Mantenedor de Vendedores"
   ClientHeight    =   6675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7530
   LinkTopic       =   "Form1"
   ScaleHeight     =   6675
   ScaleWidth      =   7530
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   6855
      Left            =   -120
      TabIndex        =   0
      Top             =   -120
      Width           =   8535
      Begin VB.Frame Frame2 
         Height          =   6135
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   6975
         Begin VB.ComboBox Cb_Estado 
            Height          =   315
            Left            =   840
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1920
            Width           =   1695
         End
         Begin VB.CommandButton Bt_Cerrar 
            Caption         =   "Cerrar"
            Height          =   315
            Left            =   5040
            TabIndex        =   10
            Top             =   1920
            Width           =   1095
         End
         Begin VB.CommandButton Bt_Save 
            Caption         =   "Guardar"
            Height          =   840
            Left            =   5040
            Picture         =   "FrmMantVendedor.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Grabar cambios componente seleccionada"
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton Bt_Limpiar 
            Caption         =   "Limpiar"
            Height          =   315
            Left            =   5040
            TabIndex        =   8
            Top             =   1440
            Width           =   1095
         End
         Begin VB.TextBox Tx_Nombre 
            Height          =   315
            Left            =   840
            TabIndex        =   4
            ToolTipText     =   "Ingrese cualquier parte de la Razón Social"
            Top             =   1200
            Width           =   3705
         End
         Begin VB.TextBox Tx_RUT 
            Height          =   315
            Left            =   840
            MaxLength       =   12
            TabIndex        =   3
            Top             =   480
            Width           =   1395
         End
         Begin VB.TextBox Tx_Codigo 
            Height          =   315
            Left            =   3120
            TabIndex        =   2
            Top             =   480
            Width           =   1395
         End
         Begin MSFlexGridLib.MSFlexGrid Grid 
            Height          =   3075
            Left            =   240
            TabIndex        =   11
            Top             =   2760
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   5424
            _Version        =   393216
            Cols            =   4
            FixedCols       =   3
            AllowUserResizing=   1
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   2000
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nombre: "
            Height          =   195
            Index           =   3
            Left            =   120
            TabIndex        =   7
            Top             =   1260
            Width           =   645
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Rut:"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   6
            Top             =   540
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Codigo:"
            Height          =   195
            Index           =   4
            Left            =   2400
            TabIndex        =   5
            Top             =   540
            Width           =   540
         End
      End
   End
End
Attribute VB_Name = "FrmMantVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const C_RUT = 0
Const C_CODIGO = 1
Const C_NOMBRE = 2
Const C_IDESTADO = 3
Const C_ESTADO = 4

Const NCOLS = C_ESTADO

Private Sub Bt_Cerrar_Click()
Unload Me
End Sub

Private Sub Bt_Limpiar_Click()
Me.Tx_Codigo.Text = ""
Me.Tx_Nombre.Text = ""
Me.Tx_RUT.Text = ""
End Sub

Private Sub Bt_Save_Click()
  Dim Q1 As String
  Dim Rs As Recordset
  
  If Not Valida Then
   Exit Sub
  End If


  Q1 = " SELECT Rut, Codigo, Nombre, Estado FROM Vendedor WHERE Rut = '" & vFmtRut(Trim(Me.Tx_RUT.Text)) & "'"

   
  Set Rs = OpenRs(DbMain, Q1)
  
  If Rs.EOF = False Then
        Q1 = " UPDATE Vendedor "
        Q1 = Q1 & " SET Codigo = " & Trim(Me.Tx_Codigo.Text) & ", "
        Q1 = Q1 & " Nombre = '" & Me.Tx_Nombre.Text & "', "
        Q1 = Q1 & " Estado = " & CbItemData(Me.Cb_Estado)
        Q1 = Q1 & " WHERE Rut = '" & vFmtRut(Trim(Me.Tx_RUT.Text)) & "'"
        Call ExecSQL(DbMain, Q1)
  Else
        
        Q1 = " INSERT INTO  Vendedor (Rut, Codigo, Nombre, Estado) Values('" & vFmtRut(Trim(Me.Tx_RUT.Text)) & "'," & Trim(Me.Tx_Codigo.Text) & ",'" & Me.Tx_Nombre.Text & "', " & CbItemData(Me.Cb_Estado) & ") "
        Call ExecSQL(DbMain, Q1)
  End If

   Call CloseRs(Rs)
   
   Call LoadAll
End Sub

Private Sub Form_Load()
Call CargaCbEstado
Call SetUpGrid
Call LoadAll
End Sub
Private Sub Grid_DblClick()
Dim Col As Integer
   Dim Row As Integer
         
   Row = Grid.MouseRow
   Col = Grid.MouseCol

   If Row < Grid.FixedRows Then
      Exit Sub
   End If
   
   If Grid.TextMatrix(Row, C_RUT) <> "" Then
        Me.Tx_RUT.Text = Grid.TextMatrix(Row, C_RUT)
        Me.Tx_Codigo.Text = Grid.TextMatrix(Row, C_CODIGO)
        Me.Tx_Nombre.Text = Grid.TextMatrix(Row, C_NOMBRE)
        Cb_Estado.ListIndex = Grid.TextMatrix(Row, C_IDESTADO)
   End If
End Sub



Private Sub Tx_Codigo_KeyPress(KeyAscii As Integer)
Call KeyNum(KeyAscii)
End Sub
Private Sub Tx_Codigo_Validate(Cancel As Boolean)
Call ValidaCodigo
End Sub

Private Sub Tx_Nombre_KeyPress(KeyAscii As Integer)
Call KeyUpper(KeyAscii)
End Sub

Private Sub Tx_RUT_KeyPress(KeyAscii As Integer)
Call KeyRut(KeyAscii)
End Sub

Private Sub Tx_RUT_LostFocus()
Dim AuxRut As String
   
   AuxRut = FmtCID(vFmtCID(Tx_RUT))
   If AuxRut <> "0-0" Then
      Tx_RUT = AuxRut
   End If
End Sub

Private Sub Tx_RUT_Validate(Cancel As Boolean)
   If Tx_RUT = "" Then
      Exit Sub
   End If
   
   If Not ValidRut(Tx_RUT) Then
      MsgBox1 "Rut No es Valido", vbExclamation
      Cancel = True
      Exit Sub
   End If

End Sub

Private Sub CargaCbEstado()

    For i = 0 To UBound(gEstado)
       Call CbAddItem(Cb_Estado, gEstado(i), i)
    Next i
    Cb_Estado.ListIndex = ES_ACTIVO
End Sub

Private Sub SetUpGrid()

   Grid.Cols = NCOLS + 1

   Call FGrSetup(Grid, True)
   Grid.FixedCols = 0
 
   Grid.ColWidth(C_RUT) = 1200
   Grid.ColWidth(C_CODIGO) = 1000
   Grid.ColWidth(C_NOMBRE) = 3000
   Grid.ColWidth(C_IDESTADO) = 0
   Grid.ColWidth(C_ESTADO) = 800


   
   Grid.ColAlignment(C_RUT) = flexAlignLeftCenter
   Grid.ColAlignment(C_CODIGO) = flexAlignLeftCenter
   Grid.ColAlignment(C_NOMBRE) = flexAlignLeftCenter
   Grid.ColAlignment(C_ESTADO) = flexAlignLeftCenter
   
   Grid.TextMatrix(0, C_RUT) = "Rut"
   Grid.TextMatrix(0, C_CODIGO) = "Codigo"
   Grid.TextMatrix(0, C_NOMBRE) = "Nombre"
   Grid.TextMatrix(0, C_ESTADO) = "Estado"

   
   Call FGrVRows(Grid, 1)
End Sub

Private Sub LoadAll()
  Dim Q1 As String
  Dim Rs As Recordset


  Q1 = " SELECT Rut, Codigo, Nombre, Estado FROM Vendedor "

   
  Set Rs = OpenRs(DbMain, Q1)
   
   
   Grid.rows = Grid.FixedRows
   i = Grid.rows
   
   Grid.Redraw = False
   
   Do While Not Rs.EOF
   
      Grid.rows = Grid.rows + 1
      
      Grid.TextMatrix(i, C_RUT) = FmtRut(vFld(Rs("Rut")))
      Grid.TextMatrix(i, C_CODIGO) = vFld(Rs("Codigo"))
      Grid.TextMatrix(i, C_NOMBRE) = vFld(Rs("Nombre"))
      Grid.TextMatrix(i, C_IDESTADO) = vFld(Rs("Estado"))
      Grid.TextMatrix(i, C_ESTADO) = gEstado(vFld(Rs("Estado")))
      
      
      i = i + 1
      Rs.MoveNext
      
   Loop
   
   Call CloseRs(Rs)
   
   Call FGrVRows(Grid, 1)
   
   Grid.TopRow = Grid.FixedRows
   
   'Marco la columna Ordenada
      
   Grid.Row = 0
   Grid.Col = C_ESTADO
   Set Grid.CellPicture = FrmMain.Pc_Flecha
   
   Grid.Redraw = True
End Sub

Private Function Valida() As Boolean
   Valida = True
   
   If Tx_RUT.Text = "" Then
      MsgBox1 "Favor ingresar un Rut", vbExclamation
      Valida = False
      Exit Function
   End If
   
   If Not ValidRut(Tx_RUT) Then
      MsgBox1 "Rut No es Valido", vbExclamation
      Valida = False
      Exit Function
   End If
   
   If Me.Tx_Codigo.Text = "" Then
      MsgBox1 "Favor ingresar un Codigo", vbExclamation
      Valida = False
      Exit Function
   End If
   
   If Not ValidaCodigo() Then
      Valida = False
      Exit Function
   End If


End Function

Private Function ValidaCodigo() As Boolean
  Dim Q1 As String
  Dim Rs As Recordset
  ValidaCodigo = True
  
  If Trim(Me.Tx_Codigo.Text) = "" Then
    Exit Function
  End If
  
  Q1 = " SELECT Rut, Codigo, Nombre, Estado FROM Vendedor WHERE Codigo = " & Trim(Me.Tx_Codigo.Text) & ""

   
  Set Rs = OpenRs(DbMain, Q1)
  If Rs.EOF = False Then
        If Trim(Me.Tx_RUT.Text) <> "" Then
            If vFld(Rs("Rut")) <> vFmtRut(Me.Tx_RUT.Text) Then
                MsgBox1 "Este codigo ya existe para otro vendedor, Favor cambiar", vbExclamation
                ValidaCodigo = False
            End If
        End If
  End If

   Call CloseRs(Rs)

End Function


