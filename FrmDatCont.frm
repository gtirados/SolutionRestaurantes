VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmDatCont 
   ClientHeight    =   4575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7680
   Icon            =   "FrmDatCont.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox textovar 
      Height          =   285
      Left            =   4440
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3240
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid grid_comi 
      Height          =   3615
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   6376
      _Version        =   327680
      Cols            =   5
   End
   Begin MSFlexGridLib.MSFlexGrid Gridtablas 
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   6376
      _Version        =   327680
      Cols            =   5
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Retornar"
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdagregar 
      Caption         =   "&Agregar"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label lblmensa 
      Caption         =   "F2 = Editar                                           DELETE = Eliminar"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   4095
   End
End
Attribute VB_Name = "FrmDatCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wBAN As Integer
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset

Private Sub cmdagregar_Click()

Dim WCODART As Integer
Dim wvalor
Dim ultimo As Integer
Dim Item As Integer
Dim Wflag
If TEXTOVAR.Visible Then
  MsgBox "Esta Usted. Editanto ....!", 48, Pub_Titulo
  Exit Sub
End If
If Left(CmdAgregar.Caption, 2) = "&G" Then
    Item = 0
    Wflag = ""
    For fila = 2 To grid_comi.Rows - 1
       If Wflag = "A" Then
          MsgBox "Verificar , El Ultimo rango debe ser parecido a unico ", 48, Pub_Titulo
          Exit Sub
       End If
       If Val(grid_comi.TextMatrix(fila, 1)) <> 0 And Val(grid_comi.TextMatrix(fila, 2)) = 0 Then
         Wflag = "A"
       ElseIf Val(grid_comi.TextMatrix(fila, 1)) > Val(grid_comi.TextMatrix(fila, 2)) Then
         MsgBox "Verificar , Dia Inicial no debe ser Mayor que la Final ", 48, Pub_Titulo
         Exit Sub
       End If
       If Not Val(grid_comi.TextMatrix(fila, 1)) >= Item Then
          MsgBox "Verificar , Dia Inicial  debe estar en Orden ascendente ", 48, Pub_Titulo
          Exit Sub
       End If
       Item = Val(grid_comi.TextMatrix(fila, 1))
    Next fila
    On Error GoTo ESCAPA
    CN.Execute "Begin Transaction", rdExecDirect
    SQ_OPER = 2
    LEER_TAB_LLAVE
    Do Until tab_mayor.EOF
      tab_mayor.Delete
      tab_mayor.MoveNext
    Loop
    ultimo = -1
    For fila = 2 To grid_comi.Rows - 1
        ultimo = ultimo + 1
        tab_mayor.AddNew
        tab_mayor!tab_codcia = PUB_CODCIA
        tab_mayor!tab_tipreg = PUB_TIPREG
        tab_mayor!TAB_NUMTAB = ultimo
        tab_mayor!TAB_NOMCORTO = Trim(grid_comi.TextMatrix(fila, 3))
        tab_mayor!TAB_NOMLARGO = Trim(grid_comi.TextMatrix(fila, 7))
        tab_mayor!TAB_CODART = 0
        tab_mayor!TAB_CONTAB2 = Trim(FrmDatCont.grid_comi.TextMatrix(fila, 1))
        tab_mayor.Update
    Next fila
    CN.Execute "Commit Transaction", rdExecDirect
    On Error GoTo 0
    Screen.MousePointer = 0
    CmdAgregar.Enabled = False
    Exit Sub
End If
If PUB_TIPREG = 199 Then MsgBox "Solo hasta 4 tipos de bloqueos"
If PUB_TIPREG = 199 Then If Gridtablas.Rows >= 5 Then Exit Sub
wBAN = 1
Mensaje = "Ingrese la Descripción  :"
titulo = "Datos Necesario..."
valorpred = " "
wvalor = InputBox(Mensaje, titulo, valorpred)
If wvalor = "" Then
 wBAN = 0
 Exit Sub
End If
If Trim(wvalor) = "" Then
  MsgBox "Invalido Nombre .. Intente Nuevamente ..", 48, Pub_Titulo
  wBAN = 0
  Exit Sub
End If
PS_REP01(0) = PUB_CODCIA
PS_REP01(1) = PUB_TIPREG
PS_REP01(2) = wvalor
PS_REP01(3) = -1
If PUB_TIPREG = 123 Then
  PS_REP01(4) = PUB_CODART
End If
llave_rep01.Requery
If Not llave_rep01.EOF Then
  MsgBox "Descripción Existe. Intente Nuevamente ..", 48, Pub_Titulo
  wBAN = 0
  Exit Sub
End If
CmdAgregar.Enabled = False
Command1.Enabled = False
WCODART = 0
If PUB_TIPREG = 123 Then
 WCODART = PUB_CODART
End If
SQ_OPER = 2
LEER_TAB_LLAVE
tab_mayor.MoveLast
If tab_mayor.EOF Then
  ultimo = 0
Else
  ultimo = tab_mayor!TAB_NUMTAB
End If
ultimo = ultimo + 1
On Error GoTo SALE
If PUB_TIPREG = 50 Then
  wBAN = 1
  Mensaje = "Ingrese su Codigo de la SUNAT   :"
  titulo = "Datos Necesario..."
  valorpred = " "
  WVALOR2 = InputBox(Mensaje, titulo, valorpred)
  If WVALOR2 = "" Then
   wBAN = 0
   Exit Sub
  End If
  If Trim(WVALOR2) = "" Then
    MsgBox "Codigo Invalido .. Intente Nuevamente ..", 48, Pub_Titulo
    wBAN = 0
    Exit Sub
  End If
  ultimo = Val(WVALOR2)
End If

tab_mayor.AddNew
tab_mayor!tab_codcia = PUB_CODCIA
tab_mayor!tab_tipreg = PUB_TIPREG
tab_mayor!TAB_NUMTAB = ultimo
tab_mayor!TAB_NOMLARGO = Left(wvalor, 40)
tab_mayor!TAB_NOMCORTO = Left(wvalor, 10)
tab_mayor!TAB_CODART = WCODART
tab_mayor.Update
Item = Val(FrmDatCont.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 0))
Gridtablas.Rows = Gridtablas.Rows + 1
FrmDatCont.Gridtablas.RowHeight(Gridtablas.Rows - 1) = 280
FrmDatCont.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 0) = ultimo
FrmDatCont.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 1) = Left(wvalor, 40)
FrmDatCont.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 2) = PUB_CODCIA
FrmDatCont.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 3) = PUB_TIPREG
FrmDatCont.Gridtablas.TextMatrix(Gridtablas.Rows - 1, 4) = ultimo

CmdAgregar.Enabled = True
Command1.Enabled = True
wBAN = 0
Exit Sub
ESCAPA:
    CN.Execute "Rollback Transaction", rdExecDirect
    MsgBox Err.Number & "  " & Err.Description & "  Intente Nuevamente."
    wBAN = 0
    CmdAgregar.Enabled = True
    Command1.Enabled = True

Exit Sub
SALE:
If Err.Number = 40002 Then
  MsgBox "Codigo Ya Existe.", 48, Pub_Titulo
Else
  MsgBox "Intente Nunevamente.", 48, Pub_Titulo
End If
CmdAgregar.Enabled = True
Command1.Enabled = True

  wBAN = 0
End Sub

Private Sub Command1_Click()
If CmdAgregar.Enabled And Left(CmdAgregar.Caption, 2) = "&G" Then
   pub_mensaje = "Los Datos han cambiado desea Grabarlos ... ?"
   Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
   If Pub_Respuesta = vbNo Then
     GoTo fin
   End If
   cmdagregar_Click
End If
fin:
Unload FrmDatCont
End Sub

Private Sub Form_Activate()
If PUB_TIPREG = 333 Then
    CmdAgregar.Enabled = False
    FrmDatCont.Gridtablas.Clear
    FrmDatCont.Gridtablas.Visible = False
    lblmensa.Caption = "[F2] = Modificar        [Insert]= Agregar      [DEL] = Eliminar "
    CmdAgregar.Caption = "&Grabar"
    wBAN = 0
    FrmDatCont.grid_comi.Clear
    FrmDatCont.grid_comi.Cols = 8
    FrmDatCont.grid_comi.Rows = 3
    FrmDatCont.grid_comi.FixedRows = 2
    FrmDatCont.grid_comi.ColWidth(0) = 400
    FrmDatCont.grid_comi.ColWidth(1) = 800
    FrmDatCont.grid_comi.ColWidth(2) = 800
    FrmDatCont.grid_comi.ColWidth(3) = 1000
    FrmDatCont.grid_comi.ColWidth(4) = 1
    FrmDatCont.grid_comi.ColWidth(5) = 1
    FrmDatCont.grid_comi.ColWidth(6) = 1
    FrmDatCont.grid_comi.TextMatrix(0, 0) = "Item"
    FrmDatCont.grid_comi.TextMatrix(0, 1) = "Cuenta"
    FrmDatCont.grid_comi.TextMatrix(1, 1) = "Destino"
    FrmDatCont.grid_comi.TextMatrix(0, 2) = "Debe(D) /"
    FrmDatCont.grid_comi.TextMatrix(1, 2) = "Haber(H)"
    FrmDatCont.grid_comi.TextMatrix(0, 3) = "  (%)   "
    FrmDatCont.grid_comi.TextMatrix(1, 3) = "Particip."
    FrmDatCont.grid_comi.ColAlignment(1) = 7
    FrmDatCont.grid_comi.ColAlignment(2) = 7
    FrmDatCont.grid_comi.ColAlignment(3) = 7
    fila = 1
    SQ_OPER = 2
    If PUB_TIPREG = 123 Then
      SQ_OPER = 3
    End If
    LEER_TAB_LLAVE
    Do Until tab_mayor.EOF
     fila = fila + 1
     FrmDatCont.grid_comi.Rows = FrmDatCont.grid_comi.Rows + 1
     FrmDatCont.grid_comi.RowHeight(fila) = 280
     FrmDatCont.grid_comi.TextMatrix(fila, 0) = Nulo_Valor0(tab_mayor!TAB_NUMTAB)
     FrmDatCont.grid_comi.TextMatrix(fila, 1) = Format(tab_mayor!TAB_contable2, "0")
     FrmDatCont.grid_comi.TextMatrix(fila, 2) = "D"
     FrmDatCont.grid_comi.TextMatrix(fila, 3) = Format(tab_mayor!TAB_NOMCORTO, "0.00")
     FrmDatCont.grid_comi.TextMatrix(fila, 4) = Nulo_Valor0(tab_mayor!tab_codcia)
     FrmDatCont.grid_comi.TextMatrix(fila, 5) = Nulo_Valor0(tab_mayor!tab_tipreg)
     FrmDatCont.grid_comi.TextMatrix(fila, 6) = Trim(tab_mayor!TAB_NUMTAB)
     FrmDatCont.grid_comi.TextMatrix(fila, 7) = Trim(tab_mayor!TAB_NOMLARGO)
     tab_mayor.MoveNext
    Loop
    FrmDatCont.grid_comi.Rows = FrmDatCont.grid_comi.Rows - 1
    FrmDatCont.grid_comi.Visible = True
    FrmDatCont.grid_comi.SetFocus
Else
  
End If

End Sub

Private Sub Form_Load()
CenterMe FrmDatCont
If PUB_TIPREG = 123 Then
  pub_cadena = "SELECT * FROM TABLAS WHERE TAB_CODCIA = ? AND TAB_TIPREG = ? AND TAB_NOMLARGO = ? AND TAB_NUMTAB <> ? AND TAB_CODART = ? ORDER BY TAB_CODCIA"
Else
  pub_cadena = "SELECT * FROM TABLAS WHERE TAB_CODCIA = ? AND TAB_TIPREG = ? AND TAB_NOMLARGO = ? AND TAB_NUMTAB <> ?  ORDER BY TAB_CODCIA"
  
End If
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01.rdoParameters(0) = " "
PS_REP01.rdoParameters(1) = 0
PS_REP01.rdoParameters(2) = " "
PS_REP01.rdoParameters(3) = 0
If PUB_TIPREG = 123 Then
 PS_REP01.rdoParameters(4) = 0
End If

Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

End Sub

Public Sub ElGrid_Click(wsGrid As MSFlexGrid, wsTexto As TextBox)
wsTexto.Left = wsGrid.Left + wsGrid.CellLeft
wsTexto.Width = wsGrid.CellWidth
wsTexto.Top = wsGrid.Top + wsGrid.CellTop
wsTexto.Tag = wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)
wsTexto.Text = wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)
wsTexto.Visible = False

End Sub
Public Sub ElGrid_EnterCell(wsGrid As MSFlexGrid, wsTexto As TextBox, Optional Bloq1, Optional Bloq2, Optional Bloq3, Optional Bloq4, Optional Bloq5)
wsGrid.CellFontBold = True
wsTexto.Left = wsGrid.Left + wsGrid.CellLeft
wsTexto.Width = wsGrid.CellWidth
wsTexto.Top = wsGrid.Top + wsGrid.CellTop
wsTexto.Tag = wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)
wsTexto.Text = wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)
Flag_Bloq = ""
If Not IsMissing(Bloq1) Then
  If wsGrid.COL = Bloq1 Then
    Flag_Bloq = "A"
  End If
End If
If Not IsMissing(Bloq2) Then
  If wsGrid.COL = Bloq2 Then
    Flag_Bloq = "A"
  End If
End If
If Not IsMissing(Bloq3) Then
  If wsGrid.COL = Bloq3 Then
    Flag_Bloq = "A"
  End If
End If
If Not IsMissing(Bloq4) Then
  If wsGrid.COL = Bloq4 Then
    Flag_Bloq = "A"
  End If
End If
If Not IsMissing(Bloq5) Then
  If wsGrid.COL = Bloq5 Then
    Flag_Bloq = "A"
  End If
End If

End Sub
Public Sub ElGrid_KeyDown(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyCode)
Flag_F2 = ""
If Flag_Bloq = "A" Then
  wsKeyCode = 0
  Exit Sub
End If

If wsKeyCode <> 113 Then
 Exit Sub
End If
If wsTexto.Visible = False Then
  Flag_F2 = "A"
  ElGrid_DblClick wsGrid, wsTexto
End If
End Sub
Public Sub ElGrid_KeyPress(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyAscii, Optional SaltaCol)
If wsKeyAscii = 27 Then
 Exit Sub
End If
If wsKeyAscii = 9 Or wsKeyAscii = 13 Then
  If Not IsMissing(SaltaCol) Then
    If wsGrid.COL = SaltaCol And wsGrid.Row <> wsGrid.Rows - 1 Then
       wsGrid.Row = wsGrid.Row + 1
       wsGrid.COL = wsGrid.FixedCols
       Exit Sub
    End If
  End If
  If wsGrid.COL <> wsGrid.Cols - 1 Then
    wsGrid.COL = wsGrid.COL + 1
  End If
  Exit Sub
End If
If Flag_Bloq = "A" Then
 wsKeyAscii = 0
 Exit Sub
End If

Dim cade
wsTexto.FontBold = True
wsTexto.ForeColor = QBColor(1)

wsTexto.Visible = True
If wsTexto.Enabled = True And wsTexto.Visible = True Then
   wsTexto.SetFocus
   wsTexto.SelStart = 0
   wsTexto.SelLength = Len(wsTexto)
End If
cade = Chr(wsKeyAscii)
SendKeys cade, True

End Sub
Private Sub ElGrid_LeaveCell(wsGrid As MSFlexGrid, wsTexto As TextBox)
If Flag_Consis = "A" Then
 wsTexto.FontBold = True
 wsTexto.ForeColor = QBColor(12)
 wsTexto.Visible = True
 If wsTexto.Enabled = True And wsTexto.Visible = True Then
   wsTexto.SetFocus
   wsTexto.SelStart = 0
   wsTexto.SelLength = Len(wsTexto)
 End If
 Exit Sub
End If
If Left(Trim(wsGrid.Text), 1) = "-" Then
 wsGrid.CellForeColor = QBColor(12)
Else
 wsGrid.CellForeColor = QBColor(0)
End If

wsGrid.CellFontBold = False
End Sub
Private Sub ElGrid_DblClick(wsGrid As MSFlexGrid, wsTexto As TextBox)
If Flag_Bloq = "A" Then
  Exit Sub
End If
wsTexto.FontBold = True
wsTexto.ForeColor = QBColor(12)
wsTexto.Visible = True
If wsTexto.Enabled = True And wsTexto.Visible = True Then
   wsTexto.SetFocus
   wsTexto.SelStart = 0
   wsTexto.SelLength = Len(wsTexto)
End If
End Sub
Private Sub ElGrid_GotFocus(wsGrid As MSFlexGrid, wsTexto As TextBox)
ElGrid_Click wsGrid, wsTexto
End Sub
Private Sub TEXTO_LosFocus(wsGrid As MSFlexGrid, wsTexto As TextBox)
ElGrid_Click wsGrid, wsTexto
End Sub
Public Sub TEXTO_KeyDown(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyCode As Integer, Optional SaltaCol)
If wsKeyCode = 46 Then
  TEXTO_KeyPress wsGrid, wsTexto, 46 '  wsKeyCode = 46
End If
If wsKeyCode = 40 Or wsKeyCode = 37 Or wsKeyCode = 39 Or wsKeyCode = 38 Then
 If Flag_F2 = "A" Then
   Exit Sub
 End If
 If Flag_Consis = "A" Then
   wsTexto.SetFocus
   wsTexto.SelStart = 0
   wsTexto.SelLength = Len(wsTexto)
   Beep
   Exit Sub
 End If
 If Gridtablas.Visible = False Then
   If grid_comi.COL <= 2 Then
    wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL) = Format(Val(wsTexto.Text), "0")
   Else
    wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL) = Format(Val(wsTexto.Text), "0.00")
   End If
   CmdAgregar.Enabled = True
 Else
   wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL) = wsTexto.Text
 End If
 wsTexto.Visible = False
 If wsKeyCode = 40 Then ' ABAJO
  If wsGrid.Row <> wsGrid.Rows - 1 Then
     wsGrid.Row = wsGrid.Row + 1
  End If
 End If
 If wsKeyCode = 38 Then ' arriba
  If wsGrid.Row <> wsGrid.FixedRows Then
     wsGrid.Row = wsGrid.Row - 1
  End If
 End If
 If wsKeyCode = 37 Then ' isquierda
  If wsGrid.COL <> wsGrid.FixedCols Then
   '  wsGrid.Col = wsGrid.Col - 1
  End If
 End If
 If wsKeyCode = 39 Then ' derecha
  If Not IsMissing(SaltaCol) Then
      If wsGrid.COL = SaltaCol And wsGrid.Row <> wsGrid.Rows - 1 Then
   '     wsGrid.Row = wsGrid.Row + 1
   '     wsGrid.Col = wsGrid.FixedCols
        GoTo wsfinal
      End If
  End If
  If wsGrid.COL <> wsGrid.Cols - 1 Then
  '   wsGrid.Col = wsGrid.Col + 1
  End If
 End If
wsfinal:
 wsTexto.FontBold = False
 wsTexto.ForeColor = QBColor(0)
 wsTexto.Text = ""
 wsGrid.SetFocus
End If

End Sub

Public Sub TEXTO_KeyPress(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyAscii As Integer, Optional SaltaCol, Optional ConsisCol1, Optional ConsisVal1, Optional ConsisCol2, Optional ConsisVal2, Optional ConsisCol3, Optional ConsisVal3, Optional ConsisCol4, Optional ConsisVal4, Optional ConsisCol5, Optional ConsisVal5, Optional ConsisCol6, Optional ConsisVal6, Optional ConsisCol7, Optional ConsisVal7, Optional ConsisCol8, Optional ConsisVal8, Optional ConsisCol9, Optional ConsisVal9, Optional ConsisCol10, Optional ConsisVal10)
If wsKeyAscii = 13 Or wsKeyAscii = 9 Then
  Flag_F2 = ""
  TEXTO_KeyDown wsGrid, wsTexto, 39, SaltaCol
  Exit Sub
End If
If wsKeyAscii = 27 Then
  ElGrid_Click wsGrid, wsTexto
  wsGrid.SetFocus
End If
If Not IsMissing(ConsisCol1) And Not IsMissing(ConsisVal1) Then
  If wsGrid.COL = ConsisCol1 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal1, ConsisCol1
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol2) And Not IsMissing(ConsisVal2) Then
  If wsGrid.COL = ConsisCol2 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal2, ConsisCol2
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol3) And Not IsMissing(ConsisVal3) Then
  If wsGrid.COL = ConsisCol3 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal3, ConsisCol3
   Exit Sub
  End If
End If

End Sub

Private Sub gridtablas_Click()
ElGrid_Click Gridtablas, TEXTOVAR
End Sub

Private Sub gridtablas_DblClick()
ElGrid_DblClick Gridtablas, TEXTOVAR
End Sub

Private Sub gridtablas_EnterCell()
ElGrid_EnterCell Gridtablas, TEXTOVAR, 2, 3, 4
End Sub

Private Sub gridtablas_GotFocus()
ElGrid_GotFocus Gridtablas, TEXTOVAR
End Sub

Private Sub gridtablas_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 And Gridtablas.COL = 1 Then
 Exit Sub
End If
If Gridtablas.Row = 0 Then
 Exit Sub
End If
If KeyCode = 45 Then
 cmdagregar_Click
 Exit Sub
End If
If KeyCode = 46 Then
    pub_mensaje = " Eliminar    : " & Trim(Gridtablas.TextMatrix(Gridtablas.Row, 1)) & " ¿Desea Continuar... ?"
    Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
    If Pub_Respuesta = vbNo Then
      Gridtablas.SetFocus
      Exit Sub
    End If
    SQ_OPER = 1
    PUB_CODCIA = FrmDatCont.Gridtablas.TextMatrix(Gridtablas.Row, 2)
    PUB_TIPREG = FrmDatCont.Gridtablas.TextMatrix(Gridtablas.Row, 3)
    PUB_NUMTAB = FrmDatCont.Gridtablas.TextMatrix(Gridtablas.Row, 4)
    LEER_TAB_LLAVE
    If tab_llave.EOF Then
      Exit Sub
    End If
    tab_llave.Delete
    If Gridtablas.Rows = 2 Then
     Gridtablas.Rows = Gridtablas.Rows - 1
    Else
    Gridtablas.RemoveItem Gridtablas.Row
    End If
    Gridtablas.SetFocus
   
End If

ElGrid_KeyDown Gridtablas, TEXTOVAR, KeyCode

End Sub

Private Sub gridtablas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Gridtablas.COL = 1 Then
 Exit Sub
End If
If KeyAscii = 27 Then
 Command1_Click
 Exit Sub
End If
End Sub

Private Sub gridtablas_LeaveCell()
ElGrid_LeaveCell Gridtablas, TEXTOVAR
End Sub

Private Sub gridtablas_Scroll()
If Gridtablas.RowIsVisible(Gridtablas.Row) = False Then
   TEXTOVAR.Visible = False
   Exit Sub
Else
  TEXTOVAR.Visible = False
End If

TEXTOVAR.Left = Gridtablas.Left + Gridtablas.CellLeft ' gridtablas.CellLeft  'Frame4.Left + gridtablas.CellLeft
TEXTOVAR.Width = Gridtablas.CellWidth
TEXTOVAR.Top = Gridtablas.Top + Gridtablas.CellTop ' Frame4.Top + gridtablas.Top + gridtablas.CellTop
TEXTOVAR.DataField = Gridtablas.COL
TEXTOVAR.Tag = Gridtablas.Row
TEXTOVAR.Text = Gridtablas.TextMatrix(Gridtablas.Row, Gridtablas.COL)

End Sub
Private Sub textovar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 PS_REP01(0) = PUB_CODCIA
 PS_REP01(1) = PUB_TIPREG
 PS_REP01(2) = TEXTOVAR.Text
 If Val(FrmDatCont.Gridtablas.TextMatrix(Gridtablas.Row, 4)) <> 0 Then
  PS_REP01(3) = Val(FrmDatCont.Gridtablas.TextMatrix(Gridtablas.Row, 4))
 Else
  PS_REP01(3) = -1
 End If
 llave_rep01.Requery
 If Not llave_rep01.EOF Then
   MsgBox "Descripción Existe. Intente Nuevamente ..", 48, Pub_Titulo
   Azul TEXTOVAR, TEXTOVAR
   Exit Sub
 End If
End If
If Gridtablas.Visible = False Then
 If KeyAscii = 13 And grid_comi.COL = 1 And Val(TEXTOVAR.Text) > 365 Then
   Azul TEXTOVAR, TEXTOVAR
   Exit Sub
 End If
 If KeyAscii = 13 And grid_comi.COL = 2 And Val(TEXTOVAR.Text) > 365 Then
   Azul TEXTOVAR, TEXTOVAR
   Exit Sub
 End If
 If KeyAscii = 13 And grid_comi.COL = 3 And Val(TEXTOVAR.Text) > 100 Then
   Azul TEXTOVAR, TEXTOVAR
   Exit Sub
 End If
 TEXTO_KeyPress grid_comi, TEXTOVAR, KeyAscii, 3, 1, 1, 2, 1, 3, 2
 Exit Sub
End If
TEXTO_KeyPress Gridtablas, TEXTOVAR, KeyAscii, 2, 5, 2, 6, 2
If KeyAscii = 13 Then
    If Gridtablas.Visible Then
    
     SQ_OPER = 1
     PUB_CODCIA = FrmDatCont.Gridtablas.TextMatrix(Gridtablas.Row, 2)
     PUB_TIPREG = FrmDatCont.Gridtablas.TextMatrix(Gridtablas.Row, 3)
     PUB_NUMTAB = FrmDatCont.Gridtablas.TextMatrix(Gridtablas.Row, 4)
     LEER_TAB_LLAVE
     If tab_llave.EOF Then
       Exit Sub
     End If
     tab_llave.Edit
     tab_llave!TAB_NOMLARGO = FrmDatCont.Gridtablas.TextMatrix(Gridtablas.Row, 1)
     If PUB_TIPREG = 222 Or PUB_TIPREG = 122 Then
      tab_llave!TAB_NOMCORTO = FrmDatCont.Gridtablas.TextMatrix(Gridtablas.Row, 5)
      tab_llave!TAB_contable2 = FrmDatCont.Gridtablas.TextMatrix(Gridtablas.Row, 6)
     End If
     tab_llave.Update
    ElseIf grid_comi.Visible Then
    
    End If
End If

End Sub
Private Sub textovar_LostFocus()
If Gridtablas.Visible Then
 TEXTO_LosFocus Gridtablas, TEXTOVAR
Else
 TEXTO_LosFocus grid_comi, TEXTOVAR
End If
End Sub


' aqui enpi

Private Sub grid_comi_Click()
If grid_comi.Rows < 3 Then
 Exit Sub
End If
ElGrid_Click grid_comi, TEXTOVAR
End Sub

Private Sub grid_comi_DblClick()
If grid_comi.Rows < 3 Then
 Exit Sub
End If
ElGrid_DblClick grid_comi, TEXTOVAR
End Sub

Private Sub grid_comi_EnterCell()
If grid_comi.Rows < 3 Then
 Exit Sub
End If
ElGrid_EnterCell grid_comi, TEXTOVAR
End Sub

Private Sub grid_comi_GotFocus()
ElGrid_GotFocus grid_comi, TEXTOVAR
End Sub

Private Sub grid_comi_KeyDown(KeyCode As Integer, Shift As Integer)
If grid_comi.Rows < 3 And KeyCode = 113 Then
 Exit Sub
End If
If grid_comi.Rows < 3 And KeyCode = 46 Then
 Exit Sub
End If

If KeyCode = 45 Then
'Insetar nuevo
 If grid_comi.TextMatrix(FrmDatCont.grid_comi.Rows - 1, 1) = "0" And grid_comi.TextMatrix(FrmDatCont.grid_comi.Rows - 1, 2) = "0" And grid_comi.TextMatrix(FrmDatCont.grid_comi.Rows - 1, 3) = "0.00" Then
  MsgBox "No procede, se Agrego y esta en blanco. ", 48, Pub_Titulo
  Exit Sub
 End If
 FrmDatCont.grid_comi.Rows = FrmDatCont.grid_comi.Rows + 1
 FrmDatCont.grid_comi.RowHeight(FrmDatCont.grid_comi.Rows - 1) = 280
 FrmDatCont.grid_comi.TextMatrix(FrmDatCont.grid_comi.Rows - 1, 1) = "0"
 FrmDatCont.grid_comi.TextMatrix(FrmDatCont.grid_comi.Rows - 1, 2) = "0"
 FrmDatCont.grid_comi.TextMatrix(FrmDatCont.grid_comi.Rows - 1, 3) = "0.00"
 CmdAgregar.Enabled = True
 Exit Sub
End If
If KeyCode = 46 Then
   pub_mensaje = " Eliminar    : " & Trim(grid_comi.TextMatrix(grid_comi.Row, 1)) & " ¿Desea Continuar... ?"
   Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
   If Pub_Respuesta = vbNo Then
      grid_comi.SetFocus
      Exit Sub
   End If
    If grid_comi.Rows = 3 Then
     grid_comi.Rows = grid_comi.Rows - 1
    Else
    grid_comi.RemoveItem grid_comi.Row
    End If
    CmdAgregar.Enabled = True
    grid_comi.SetFocus
End If

ElGrid_KeyDown grid_comi, TEXTOVAR, KeyCode

End Sub

Private Sub grid_comi_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And grid_comi.COL = 1 Then
 Exit Sub
End If
If KeyAscii = 27 Then
 Command1_Click
 Exit Sub
End If
End Sub

Private Sub grid_comi_LeaveCell()
ElGrid_LeaveCell grid_comi, TEXTOVAR
End Sub

Private Sub grid_comi_Scroll()
If grid_comi.RowIsVisible(grid_comi.Row) = False Then
   TEXTOVAR.Visible = False
   Exit Sub
Else
  TEXTOVAR.Visible = False
End If

TEXTOVAR.Left = grid_comi.Left + grid_comi.CellLeft ' grid_comi.CellLeft  'Frame4.Left + grid_comi.CellLeft
TEXTOVAR.Width = grid_comi.CellWidth
TEXTOVAR.Top = grid_comi.Top + grid_comi.CellTop ' Frame4.Top + grid_comi.Top + grid_comi.CellTop
TEXTOVAR.DataField = grid_comi.COL
TEXTOVAR.Tag = grid_comi.Row
TEXTOVAR.Text = grid_comi.TextMatrix(grid_comi.Row, grid_comi.COL)

End Sub




Private Sub Consistencias(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyAscii As Integer, Optional ConsisVal, Optional ConsisCol)
  Static valor
  Dim car As String
  Flag_Consis = ""
  If ConsisVal = 2 Then ' NUMEROS CON DECIMALES
    car = Chr$(wsKeyAscii)
    car = UCase$(Chr$(wsKeyAscii))
    wsKeyAscii = Asc(car)
    If wsKeyAscii = 45 Then
      If wsTexto.Text <> "" Then
         Beep
         wsKeyAscii = 0
         Exit Sub
      End If
      Flag_Consis = "A"
    End If
    If wsKeyAscii = 46 Then
      If InStr(1, wsTexto.Text, ".") <> 0 Then
        Beep
        wsKeyAscii = 0
        Exit Sub
      End If
    End If
    If car < "0" Or car > "9" Then
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And car <> "." And car <> "-" Then
          wsKeyAscii = 0
          Beep
          Exit Sub
        End If
    End If
  ElseIf ConsisVal = 1 Then ' NUMEROS ENTEROS
    car = Chr$(wsKeyAscii)
    car = UCase$(Chr$(wsKeyAscii))
    wsKeyAscii = Asc(car)
    If car < "0" Or car > "9" Then
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And wsKeyAscii <> 32 And car <> "-" Then
          wsKeyAscii = 0
          Beep
        End If
      End If
  End If

End Sub

