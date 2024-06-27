VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frmformulacion 
   Caption         =   "Formulación"
   ClientHeight    =   4905
   ClientLeft      =   60
   ClientTop       =   1620
   ClientWidth     =   9480
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   9480
   Begin VB.TextBox txtvar 
      Height          =   285
      Left            =   4560
      TabIndex        =   17
      Top             =   3720
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Articulos para Procesos"
      Height          =   4215
      Left            =   6790
      TabIndex        =   2
      Top             =   120
      Width           =   2775
      Begin VB.ListBox LISARTI 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3840
         Left            =   40
         TabIndex        =   14
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1455
      Left            =   3000
      TabIndex        =   7
      Top             =   120
      Width           =   3495
      Begin VB.CommandButton grabaorden 
         Caption         =   "Grabar &Orden"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   19
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Calcular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtPeso 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txttotal 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Total Peso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Total  ( S/.)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   10
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Ce&rrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   6
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Frame FrameP 
      Caption         =   "Lista de Articulos para Procesos"
      Height          =   2535
      Left            =   0
      TabIndex        =   3
      Top             =   1800
      Width           =   6855
      Begin VB.CommandButton cmdbajar 
         Caption         =   "Bajar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6200
         TabIndex        =   16
         Top             =   1200
         Width           =   600
      End
      Begin VB.CommandButton cmdsubir 
         Caption         =   "Subir"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6200
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   720
         Width           =   600
      End
      Begin MSFlexGridLib.MSFlexGrid gridpro 
         Height          =   2175
         Left            =   0
         TabIndex        =   18
         Top             =   240
         Width           =   6090
         _ExtentX        =   10742
         _ExtentY        =   3836
         _Version        =   327680
         AllowBigSelection=   0   'False
         FocusRect       =   2
         HighLight       =   2
         AllowUserResizing=   1
      End
      Begin VB.Label Label4 
         Caption         =   "Orden"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6220
         TabIndex        =   20
         Top             =   360
         Width           =   550
      End
      Begin VB.Label momento 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Top             =   840
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Procesos :"
      Height          =   1215
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   2895
      Begin VB.CommandButton Command3 
         Caption         =   "Crear &Procesos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   13
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox LISPROCESOS 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Editar Procesos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "Frmformulacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cod As String
Dim kPASE As Boolean
Dim Flag_Consis As String * 1
Dim Flag_F2 As String * 1
Dim Flag_Bloq As String * 1
Dim Flag_Inicial As String * 1
Dim Flag_Change  As String * 1
Dim PS_REP02 As rdoQuery
Dim llave_rep02 As rdoResultset

Private Sub cmdbajar_Click()
If gridpro.Enabled And gridpro.Rows > 1 Then
 ORDENA_GRID gridpro, 1
 grabaorden.Enabled = True
 gridpro.SetFocus
End If
End Sub

Private Sub cmdsubir_Click()
If gridpro.Enabled And gridpro.Rows > 1 Then
 ORDENA_GRID gridpro, -1
 grabaorden.Enabled = True
 gridpro.SetFocus
End If
End Sub

Private Sub Command1_Click()
If kPASE Then
  SUMTOTAL
  Command1.SetFocus
End If
End Sub

Private Sub Command2_Click()
Unload Frmformulacion
End Sub

Private Sub Command3_Click()
Load FrmProcesos
FrmProcesos.Show 1
'LISTA_PROCESOS
End Sub


Private Sub Form_Load()
Dim cade
Flag_Consis = ""
Flag_F2 = ""
Flag_Bloq = "A"
Flag_Inicial = "A"
Flag_Change = ""
pub_cadena = "SELECT * FROM PROCESOS WHERE PRO_CODCIA = ? AND PRO_CODART = ? AND PRO_CODPRO = ? AND PRO_SEC = ? ORDER BY PRO_SEC"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurValues)
kPASE = False
LISTA_PROCESOS
LISTA_ARTI
gridpro.Enabled = False


End Sub

Public Sub LISTA_ARTI()
Dim ca, WF
Dim loc_arti As rdoResultset
cad = "SELECT * FROM ARTI WHERE ART_KEY > 0 AND ART_CODCIA = '" & LK_CODCIA & "' ORDER BY ART_NOMBRE"
Set loc_arti = CN.OpenResultset(cad, rdOpenKeyset, rdConcurValues)
loc_arti.Requery

Do Until loc_arti.EOF
  LISARTI.AddItem loc_arti!ART_NOMBRE & String(30, " ") & loc_arti!ART_KEY
  loc_arti.MoveNext
Loop

End Sub
Public Sub LISTA_PROCESOS()
SQ_OPER = 2
PUB_CODCIA = LK_CODCIA
PUB_TIPREG = 888
LEER_TAB_LLAVE
LISPROCESOS.Clear
Do Until tab_mayor.EOF
  LISPROCESOS.AddItem tab_mayor!TAB_NOMLARGO & String(30, " ") & tab_mayor!tab_numtab
tab_mayor.MoveNext
Loop


End Sub


Private Sub grabaorden_Click()
 Screen.MousePointer = 11
 PS_REP02(0) = LK_CODCIA
 For fila = 1 To gridpro.Rows - 1
   PS_REP02(1) = Val(gridpro.TextMatrix(fila, 8))
   PS_REP02(2) = Val(gridpro.TextMatrix(fila, 7))
   PS_REP02(3) = Val(gridpro.TextMatrix(fila, 6))
   llave_rep02.Requery
   If llave_rep02.EOF Then
     MsgBox "Error Intentar , Solo se han Grabado los " & fila & " Primeros", 48, Pub_Titulo
     Exit Sub
   End If
   llave_rep02.Edit
   llave_rep02!PRO_ORDEN = fila
   llave_rep02.Update
 Next fila
 LISPROCESOS_Click
 Screen.MousePointer = 0
 grabaorden.Enabled = False
End Sub

Private Sub gridpro_Click()
If gridpro.Rows <> 1 Then
 ElGrid_Click gridpro, txtvar
End If
End Sub

Private Sub gridpro_EnterCell()
If gridpro.Rows <> 1 Then
ElGrid_EnterCell gridpro, txtvar, 1, 3, 4, 5
End If
End Sub

Private Sub gridpro_GotFocus()
If gridpro.Rows <> 1 Then
 ElGrid_GotFocus gridpro, txtvar
End If
End Sub

Private Sub gridpro_KeyDown(KeyCode As Integer, Shift As Integer)
If gridpro.Rows = 1 Then
 Exit Sub
End If
If KeyCode = 46 Then
 Dim WSEC As Integer
 Dim WCODPRO As Integer
 Dim WCODART As Integer
 pub_mensaje = "Eliminar este Articulo del Proceso " & Trim(Left(LISPROCESOS.text, 30)) & " ..?"
 Pub_Respuesta = MsgBox(pub_mensaje, vbInformation + vbYesNo + vbDefaultButton2, Pub_Titulo)
 If Pub_Respuesta = vbNo Then
   Screen.MousePointer = 0
   Exit Sub
 End If
 Screen.MousePointer = 11
 WSEC = Val(gridpro.TextMatrix(gridpro.Row, 6))
 WCODPRO = Val(gridpro.TextMatrix(gridpro.Row, 7))
 WCODART = Val(gridpro.TextMatrix(gridpro.Row, 8))
 PS_REP02(0) = LK_CODCIA
 PS_REP02(1) = WCODART
 PS_REP02(2) = WCODPRO
 PS_REP02(3) = WSEC
 llave_rep02.Requery
 If llave_rep02.EOF Then
  Screen.MousePointer = 0
  MsgBox "Algun Error Intentente nuevamente...", 48, Pub_Titulo
  Exit Sub
 End If
 llave_rep02.Delete
 If gridpro.Row = 1 Then
   gridpro.Rows = 1
 Else
  gridpro.RemoveItem gridpro.Row
 End If
 Screen.MousePointer = 0
 Exit Sub
End If
ElGrid_KeyDown gridpro, txtvar, KeyCode
End Sub

Private Sub gridpro_KeyPress(KeyAscii As Integer)
If gridpro.Rows = 1 Then
 Exit Sub
End If
If KeyAscii = 13 And gridpro.Col = 5 Then
 Exit Sub
End If
ElGrid_KeyPress gridpro, txtvar, KeyAscii
End Sub

Private Sub gridpro_LeaveCell()
If gridpro.Rows <> 1 Then
ElGrid_LeaveCell gridpro, txtvar
End If
End Sub

Private Sub LISARTI_DblClick()
LISARTI_KeyPress 13
End Sub

Private Sub LISARTI_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
  Exit Sub
End If
Dim PSPROC_MAYOR2 As rdoQuery
Dim proc_mayor2   As rdoResultset
Dim WSEC
Static WORDEN
Dim wCODARTI
wCODARTI = Val(Trim(Right(LISARTI.text, 8)))
If wCODARTI = 0 Or kPASE = False Then
  MsgBox "Debe Seleccionar Proceso ..", 48, Pub_Titulo
  LISPROCESOS.SetFocus
  Exit Sub
End If
PUB_CODPRO = Val(Right(LISPROCESOS.text, 8))
Screen.MousePointer = 11
pub_cadena = "SELECT * FROM PROCESOS WHERE PRO_CODCIA = ? AND PRO_CODPRO = ? ORDER BY PRO_CODCIA, PRO_SEC"
Set PSPROC_MAYOR2 = CN.CreateQuery("", pub_cadena)
Set proc_mayor2 = PSPROC_MAYOR2.OpenResultset(rdOpenKeyset, rdConcurValues)
PSPROC_MAYOR2(0) = LK_CODCIA
PSPROC_MAYOR2(1) = PUB_CODPRO
proc_mayor2.Requery
Do Until proc_mayor2.EOF
 If Val(proc_mayor2!PRO_CODART) = wCODARTI Then
   Screen.MousePointer = 0
   MsgBox " YA Existe en Proceso. ", 48, Pub_Titulo
   Exit Sub
 End If
 proc_mayor2.MoveNext
Loop
SQ_OPER = 1
pu_codcia = LK_CODCIA
PUB_KEY = wCODARTI
LEER_ART_LLAVE
If art_LLAVE.EOF Or Val(wCODARTI) = 1 Then
   Screen.MousePointer = 0
   MsgBox "Error de Articulo  ...", 48, Pub_Titulo
   Exit Sub
End If
SQ_OPER = 1
pu_codcia = LK_CODCIA
PUB_CODART = wCODARTI
LEER_ARM_LLAVE
If arm_llave.EOF Then
   Screen.MousePointer = 0
   MsgBox "Error en stock  ...", 48, Pub_Titulo
   Exit Sub
End If
proc_mayor2.MoveFirst
If proc_mayor2.EOF Then
  WSEC = 1
  WORDEN = 1
Else
  proc_mayor2.MoveLast
  WSEC = proc_mayor2!PRO_SEC
  WORDEN = proc_mayor2!PRO_ORDEN
  WORDEN = WORDEN + 1
  WSEC = WSEC + 1
End If
proc_mayor2.AddNew
proc_mayor2!pro_codpro = PUB_CODPRO
proc_mayor2!PRO_CODART = Val(wCODARTI)
proc_mayor2!PRO_CODCIA = LK_CODCIA
proc_mayor2!PRO_SEC = WSEC
proc_mayor2!PRO_ORDEN = WORDEN
proc_mayor2.Update
gridpro.Rows = gridpro.Rows + 1
gridpro.RowHeight(gridpro.Rows - 1) = 285
gridpro.TextMatrix(gridpro.Rows - 1, 0) = gridpro.Rows - 1
gridpro.TextMatrix(gridpro.Rows - 1, 1) = Left(LISARTI.text, 30)
gridpro.TextMatrix(gridpro.Rows - 1, 2) = 0
gridpro.TextMatrix(gridpro.Rows - 1, 3) = art_LLAVE!ART_COSPRO
gridpro.TextMatrix(gridpro.Rows - 1, 4) = 0
gridpro.TextMatrix(gridpro.Rows - 1, 5) = arm_llave!arm_stock
gridpro.TextMatrix(gridpro.Rows - 1, 6) = WSEC
gridpro.TextMatrix(gridpro.Rows - 1, 7) = PUB_CODPRO
gridpro.TextMatrix(gridpro.Rows - 1, 8) = Val(wCODARTI)

Screen.MousePointer = 0
End Sub

Private Sub LISPROCESOS_Click()
On Error GoTo SALE
cod = Trim(Right(LISPROCESOS.text, 8))
If cod = "" Then
  MsgBox "Tiene que Seleccionar un Proceso ..", 48, Pub_Titulo
  LISPROCESOS.SetFocus
  Exit Sub
End If
DoEvents
momento.Caption = "Un Momento .."
DoEvents
LLENA_PRO Val(cod)
Frmformulacion.momento.Caption = ""
gridpro.Enabled = True
cmdsubir.Enabled = True
cmdbajar.Enabled = True
kPASE = True
SUMTOTAL
Exit Sub
SALE:
MsgBox Err.Description, 48, Pub_Titulo
Unload Frmformulacion
Exit Sub
End Sub

Private Sub LISPROCESOS_GotFocus()
'Frmformulacion.DBGrid1.Visible = False
cmdsubir.Enabled = False
cmdbajar.Enabled = False
kPASE = False
End Sub


Public Sub SUMTOTAL()
On Error GoTo SALE
Dim SUMPESO
Dim SUMTOTAL
SUMPESO = 0
SUMTOTAL = 0
Frmformulacion.gridpro.Visible = False
DoEvents
momento.Caption = "Un Momento .."
DoEvents
For fila = 1 To Frmformulacion.gridpro.Rows - 1
 SUMPESO = SUMPESO + Val(Frmformulacion.gridpro.TextMatrix(fila, 2))
 SUMTOTAL = SUMTOTAL + Val(Frmformulacion.gridpro.TextMatrix(fila, 4))
Next fila
txtPeso.text = Format(SUMPESO, "Currency")
txttotal.text = Format(SUMTOTAL, "Currency")
Frmformulacion.gridpro.Visible = True
momento.Caption = ""
If Frmformulacion.gridpro.Rows > 0 Then
  gridpro.SetFocus
End If

Exit Sub
SALE:
If Err.Number = 40503 Then
  'Frmformulacion.DBGrid1.Visible = True
  momento.Caption = ""
  MsgBox "Datos Incorrectos Verificar ...!!!", 48, Pub_Titulo
  DBGrid1.SetFocus
Else
  MsgBox Err.Description, 48, Pub_Titulo
  'Frmformulacion.MSRDC1.Cancel
  'Frmformulacion.DBGrid1.Refresh
  Unload Frmformulacion
  
End If

End Sub

Public Sub LLENA_PRO(cod As Integer)
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset
gridpro.Clear
gridpro.Cols = 9
gridpro.ColWidth(0) = 300
gridpro.ColWidth(1) = 2000
gridpro.ColAlignment(1) = 1
gridpro.ColWidth(2) = 700
gridpro.ColWidth(3) = 700
gridpro.ColWidth(4) = 700
gridpro.ColWidth(5) = 1000
gridpro.ColWidth(6) = 1
gridpro.ColWidth(7) = 1
gridpro.ColWidth(8) = 1
gridpro.TextMatrix(0, 0) = "It."
gridpro.TextMatrix(0, 1) = "Articulo"
gridpro.TextMatrix(0, 2) = "Cantidad"
gridpro.TextMatrix(0, 3) = "CosPro"
gridpro.TextMatrix(0, 4) = "Sub-Total"
gridpro.TextMatrix(0, 5) = "Stock"

csele = "PROCESOS.PRO_CODPRO = " & cod & " AND PROCESOS.PRO_CODCIA = '" & LK_CODCIA & "' ORDER BY PROCESOS.PRO_ORDEN"
cadetodo = " SELECT ARTI.ART_NOMBRE, PROCESOS.PRO_SEC, PROCESOS.PRO_CODART, PROCESOS.PRO_CODPRO, PROCESOS.PRO_CANTIDAD, ARTI.ART_COSPRO, PROCESOS.PRO_SUBTOTAL,ARTICULO.ARM_STOCK, PROCESOS.PRO_ORDEN, PROCESOS.PRO_SEC From  PROCESOS, ARTICULO, ARTI  Where  PROCESOS.PRO_CODART = ARTICULO.ARM_CODART  AND  PROCESOS.PRO_CODCIA = ARTICULO.ARM_CODCIA   AND PROCESOS.PRO_CODART = ARTI.ART_KEY  AND PROCESOS.PRO_CODCIA = ARTI.ART_CODCIA  AND " & csele
Set PS_REP01 = CN.CreateQuery("", cadetodo)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
llave_rep01.Requery
gridpro.Rows = 1
If llave_rep01.EOF Then
 MsgBox " NO Existe Datos ..", 48, Pub_Titulo
 Exit Sub
End If
gridpro.Rows = 2
fila = 0
Do Until llave_rep01.EOF
    fila = fila + 1
    gridpro.Rows = fila + 1
    gridpro.RowHeight(fila) = 285
    gridpro.TextMatrix(fila, 0) = fila
    gridpro.TextMatrix(fila, 1) = llave_rep01!ART_NOMBRE
    gridpro.TextMatrix(fila, 2) = Nulo_Valor0(llave_rep01!PRO_CANTIDAD)
    gridpro.TextMatrix(fila, 3) = llave_rep01!ART_COSPRO
    gridpro.TextMatrix(fila, 4) = Nulo_Valor0(llave_rep01!PRO_SUBTOTAL)
    gridpro.TextMatrix(fila, 5) = llave_rep01!arm_stock
    gridpro.TextMatrix(fila, 6) = llave_rep01!PRO_SEC
    gridpro.TextMatrix(fila, 7) = llave_rep01!pro_codpro
    gridpro.TextMatrix(fila, 8) = llave_rep01!PRO_CODART
    llave_rep01.MoveNext
Loop

End Sub

Public Sub ElGrid_Click(wsGrid As MSFlexGrid, wsTexto As TextBox)
If wsGrid.CellWidth < 0 Then
 Exit Sub
End If
wsTexto.Left = wsGrid.Left + wsGrid.CellLeft
wsTexto.Width = wsGrid.CellWidth
wsTexto.Top = wsGrid.Top + wsGrid.CellTop + FrameP.Top
wsTexto.Tag = wsGrid.TextMatrix(wsGrid.Row, wsGrid.Col)
wsTexto.text = wsGrid.TextMatrix(wsGrid.Row, wsGrid.Col)
wsTexto.Visible = False

End Sub
Public Sub ElGrid_EnterCell(wsGrid As MSFlexGrid, wsTexto As TextBox, Optional Bloq1, Optional Bloq2, Optional Bloq3, Optional Bloq4, Optional Bloq5)
If wsGrid.CellWidth < 0 Then
Exit Sub
End If
'wsGrid.CellFontBold = True
wsGrid.CellBackColor = QBColor(1)
wsGrid.CellForeColor = QBColor(15)

wsTexto.Left = wsGrid.Left + wsGrid.CellLeft
wsTexto.Width = wsGrid.CellWidth
wsTexto.Top = wsGrid.Top + wsGrid.CellTop + FrameP.Top
wsTexto.Tag = wsGrid.TextMatrix(wsGrid.Row, wsGrid.Col)
wsTexto.text = wsGrid.TextMatrix(wsGrid.Row, wsGrid.Col)
Flag_Bloq = ""
If Not IsMissing(Bloq1) Then
  If wsGrid.Col = Bloq1 Then
    Flag_Bloq = "A"
  End If
End If
If Not IsMissing(Bloq2) Then
  If wsGrid.Col = Bloq2 Then
    Flag_Bloq = "A"
  End If
End If
If Not IsMissing(Bloq3) Then
  If wsGrid.Col = Bloq3 Then
    Flag_Bloq = "A"
  End If
End If
If Not IsMissing(Bloq4) Then
  If wsGrid.Col = Bloq4 Then
    Flag_Bloq = "A"
  End If
End If
If Not IsMissing(Bloq5) Then
  If wsGrid.Col = Bloq5 Then
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
    If wsGrid.Col = SaltaCol And wsGrid.Row <> wsGrid.Rows - 1 Then
       wsGrid.Row = wsGrid.Row + 1
       wsGrid.Col = wsGrid.FixedCols
       Exit Sub
    End If
  End If
  If wsGrid.Col <> wsGrid.Cols - 1 Then
    If wsGrid.Col <> 6 Then
     wsGrid.Col = wsGrid.Col + 1
    End If
  End If
  Exit Sub
End If
If Flag_Bloq = "A" Then
 wsKeyAscii = 0
 Exit Sub
End If

Dim cade
'wsTexto.FontBold = True
wsTexto.ForeColor = QBColor(1)
wsTexto.text = ""
wsTexto.Visible = True
cade = UCase(Chr(wsKeyAscii))
'wsTexto.text = cade
If wsTexto.Enabled = True And wsTexto.Visible = True Then
   wsTexto.SetFocus
   wsTexto.SelStart = 0
   wsTexto.SelLength = Len(wsTexto)
End If
Flag_Change = "A"
'cade = Chr(wsKeyAscii)
SendKeys cade, True
wsTexto.SelStart = Len(wsTexto)

End Sub
Private Sub ElGrid_LeaveCell(wsGrid As MSFlexGrid, wsTexto As TextBox)
If Flag_Consis = "A" Then
 'wsTexto.FontBold = True
 wsTexto.ForeColor = QBColor(12)
 wsTexto.Visible = True
 If wsTexto.Enabled = True And wsTexto.Visible = True Then
   wsTexto.SetFocus
   wsTexto.SelStart = 0
   wsTexto.SelLength = Len(wsTexto)
 End If
 Exit Sub
End If
If Left(Trim(wsGrid.text), 1) = "-" Then
 wsGrid.CellForeColor = QBColor(12)
 wsGrid.CellBackColor = QBColor(15)
Else
 wsGrid.CellBackColor = QBColor(15)
 wsGrid.CellForeColor = QBColor(0)
End If
'wsGrid.CellFontBold = False
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
 Flag_Change = ""
 If wsGrid.Col <> 0 Then
  wsGrid.TextMatrix(wsGrid.Row, wsGrid.Col) = Format(wsTexto.text, "0.0000")
 End If
 gridpro.TextMatrix(wsGrid.Row, 4) = Format(Val(wsTexto.text) * Val(gridpro.TextMatrix(wsGrid.Row, 3)), "0.00")
 Screen.MousePointer = 11
 GRABA_DATO Val(wsGrid.TextMatrix(wsGrid.Row, wsGrid.Col)), Val(gridpro.TextMatrix(wsGrid.Row, 4))
 Screen.MousePointer = 0
 Flag_Bloq = ""
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
  If wsGrid.Col <> wsGrid.FixedCols Then
     If wsGrid.Col = 3 Then
        wsGrid.Col = wsGrid.Col - 2
     Else
        wsGrid.Col = wsGrid.Col - 1
     End If
  End If
 End If
 If wsKeyCode = 39 Then ' derecha
  If Not IsMissing(SaltaCol) Then
     If wsGrid.Col = SaltaCol Then
        If wsGrid.Row <> wsGrid.Rows - 1 Then
          wsGrid.Row = wsGrid.Row + 1
          wsGrid.Col = wsGrid.FixedCols
          GoTo wsfinal
        ElseIf wsGrid.Row = wsGrid.Rows - 1 And wsGrid.Col = wsGrid.Cols - 1 Then
         If Trim(wsGrid.TextMatrix(wsGrid.Row, 0)) <> "" And Val(wsGrid.TextMatrix(wsGrid.Row, 1)) <> 0 And Val(wsGrid.TextMatrix(wsGrid.Row, 3)) <> 0 Then
          ' wsGrid.Rows = wsGrid.Rows + 1
           'wsGrid.Row = wsGrid.Row + 1
         '  wsGrid.Col = wsGrid.FixedCols
         '  GoTo wsfinal
         Else
           wsGrid.Col = wsGrid.FixedCols
          GoTo wsfinal
         End If
        End If
     ElseIf wsGrid.Row = wsGrid.Rows - 1 And wsGrid.Col = wsGrid.Cols - 1 Then
        wsGrid.Col = wsGrid.FixedCols
        GoTo wsfinal
     ElseIf wsGrid.Col = 2 And wsGrid.Row <> wsGrid.Rows - 1 Then
       wsGrid.Row = wsGrid.Row + 1
       Exit Sub
     End If
  End If
  If wsGrid.Col <> wsGrid.Cols - 1 Then
     If wsGrid.Col = 1 Then
      wsGrid.Col = wsGrid.Col + 2
     ElseIf wsGrid.Col >= 13 Then
'      wsGrid.Row = wsGrid.Row + 1
      wsGrid.Col = wsGrid.FixedCols
       GoTo wsfinal
     Else
      wsGrid.Col = wsGrid.Col + 1
     End If
  End If
 End If
wsfinal:
 wsTexto.FontBold = False
 wsTexto.ForeColor = QBColor(0)
 wsTexto.text = ""
 wsGrid.SetFocus
End If
'Exit Sub

End Sub

Public Sub TEXTO_KeyPress(wsGrid As MSFlexGrid, wsTexto As TextBox, wsKeyAscii As Integer, Optional SaltaCol, Optional ConsisCol1, Optional ConsisVal1, Optional ConsisCol2, Optional ConsisVal2, Optional ConsisCol3, Optional ConsisVal3, Optional ConsisCol4, Optional ConsisVal4, Optional ConsisCol5, Optional ConsisVal5, Optional ConsisCol6, Optional ConsisVal6, Optional ConsisCol7, Optional ConsisVal7, Optional ConsisCol8, Optional ConsisVal8, Optional ConsisCol9, Optional ConsisVal9, Optional ConsisCol10, Optional ConsisVal10)
If wsKeyAscii = 13 Or wsKeyAscii = 9 Then
  Flag_F2 = ""
  TEXTO_KeyDown wsGrid, wsTexto, 39, SaltaCol
  Exit Sub
End If
If wsKeyAscii = 27 Then
  ElGrid_Click wsGrid, wsTexto
  Flag_Change = "A"
  wsGrid.SetFocus
End If
If Not IsMissing(ConsisCol1) And Not IsMissing(ConsisVal1) Then
  If wsGrid.Col = ConsisCol1 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal1, ConsisCol1
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol2) And Not IsMissing(ConsisVal2) Then
  If wsGrid.Col = ConsisCol2 Then '
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal2, ConsisCol2
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol3) And Not IsMissing(ConsisVal3) Then
  If wsGrid.Col = ConsisCol3 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal3, ConsisCol3
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol4) And Not IsMissing(ConsisVal4) Then
  If wsGrid.Col = ConsisCol4 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal4, ConsisCol4
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol5) And Not IsMissing(ConsisVal5) Then
  If wsGrid.Col = ConsisCol5 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal5, ConsisCol5
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol6) And Not IsMissing(ConsisVal6) Then
  If wsGrid.Col = ConsisCol6 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal6, ConsisCol6
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol7) And Not IsMissing(ConsisVal7) Then
  If wsGrid.Col = ConsisCol7 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal7, ConsisCol7
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol8) And Not IsMissing(ConsisVal8) Then
  If wsGrid.Col = ConsisCol8 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal8, ConsisCol8
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol9) And Not IsMissing(ConsisVal9) Then
  If wsGrid.Col = ConsisCol9 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal9, ConsisCol9
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol10) And Not IsMissing(ConsisVal10) Then
  If wsGrid.Col = ConsisCol10 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal10, ConsisCol10
   Exit Sub
  End If
End If

End Sub

Private Sub txtvar_KeyDown(KeyCode As Integer, Shift As Integer)
TEXTO_KeyDown gridpro, txtvar, KeyCode
End Sub

Private Sub txtvar_KeyPress(KeyAscii As Integer)
TEXTO_KeyPress gridpro, txtvar, KeyAscii, 1, 2, 2
End Sub

Private Sub txtvar_LostFocus()
TEXTO_LosFocus gridpro, txtvar
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
      If wsTexto.text <> "" Then
         Beep
         wsKeyAscii = 0
         Exit Sub
      End If
      Flag_Consis = "A"
    End If
    If wsKeyAscii = 46 Then
      If InStr(1, wsTexto.text, ".") <> 0 Then
        Beep
        wsKeyAscii = 0
        Exit Sub
      End If
    End If
    If car < "0" Or car > "9" Then
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And wsKeyAscii <> 32 And car <> "." And car <> "-" Then
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

Public Sub ORDENA_GRID(wgrid As MSFlexGrid, wvalor As Integer)
Dim Datos1() As String
Dim Datos2() As String
'wvalor solo 1  Bajar uno
'wvalor solo -1  Subir uno
If wvalor = 1 Then
  If wgrid.Row = wgrid.Rows - 1 Then
    Exit Sub
  End If
ElseIf wvalor = -1 Then
  If wgrid.Row = 1 Then
   Exit Sub
  End If
End If
Dim tempo1 As Integer
Dim tempo2 As Integer
Dim texto1 As String
Dim texto2 As String
ReDim Datos1(wgrid.Cols - 1) As String
ReDim Datos2(wgrid.Cols - 1) As String
tempo1 = wgrid.Row
tempo2 = wgrid.Row + (1 * wvalor)
For i = 0 To wgrid.Cols - 1
 Datos1(i) = wgrid.TextMatrix(wgrid.Row, i)
Next i
wgrid.Row = wgrid.Row + (1 * wvalor)
For i = 0 To wgrid.Cols - 1
 Datos2(i) = wgrid.TextMatrix(wgrid.Row, i)
Next i
wgrid.RemoveItem tempo2
wgrid.AddItem texto1, tempo2
For i = 0 To wgrid.Cols - 1
 wgrid.TextMatrix(tempo2, i) = Datos1(i)
Next i
wgrid.RowHeight(tempo2) = 285
'Selected(tempo2) = True

wgrid.Row = tempo1
wgrid.RemoveItem tempo2
wgrid.AddItem texto2, tempo1
For i = 0 To wgrid.Cols - 1
 wgrid.TextMatrix(tempo1, i) = Datos2(i)
Next i
wgrid.RowHeight(tempo1) = 285
wgrid.Row = tempo2
'lista.Selected(tempo2) = True


End Sub


Public Sub GRABA_DATO(wcantidad As Currency, WSUBTOTAL As Currency)
Dim WSEC As Integer
Dim WCODPRO As Integer
Dim WCODART As Integer
WSEC = Val(gridpro.TextMatrix(gridpro.Row, 6))
WCODPRO = Val(gridpro.TextMatrix(gridpro.Row, 7))
WCODART = Val(gridpro.TextMatrix(gridpro.Row, 8))
PS_REP02(0) = LK_CODCIA
PS_REP02(1) = WCODART
PS_REP02(2) = WCODPRO
PS_REP02(3) = WSEC
llave_rep02.Requery
If llave_rep02.EOF Then
 Exit Sub
End If
llave_rep02.Edit
llave_rep02!PRO_CANTIDAD = wcantidad
llave_rep02!PRO_SUBTOTAL = WSUBTOTAL
llave_rep02.Update

End Sub
