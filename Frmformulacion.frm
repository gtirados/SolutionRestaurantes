VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "Comctl32.ocx"
Begin VB.Form Frmformulacion 
   Caption         =   "Formulación"
   ClientHeight    =   5220
   ClientLeft      =   60
   ClientTop       =   1620
   ClientWidth     =   10950
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   10950
   Begin ComctlLib.ListView ListView1 
      Height          =   375
      Left            =   840
      TabIndex        =   21
      Top             =   4560
      Visible         =   0   'False
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   661
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.TextBox txtvar 
      Height          =   285
      Left            =   5160
      TabIndex        =   17
      Top             =   3600
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Caption         =   "Articulos para Procesos"
      Height          =   4575
      Left            =   6790
      TabIndex        =   7
      Top             =   0
      Width           =   4095
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
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label5 
         Caption         =   "[Enter] = Seleccionar         [F3] = Buscar"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   4080
         Width           =   1935
         WordWrap        =   -1  'True
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
      Left            =   4440
      TabIndex        =   5
      Top             =   4680
      Width           =   1695
   End
   Begin VB.Frame FrameP 
      Caption         =   "Lista de Insumos para Prod. Termiado"
      Height          =   2775
      Left            =   0
      TabIndex        =   8
      Top             =   1800
      Width           =   6855
      Begin VB.CommandButton cmdbajar 
         Caption         =   "Bajar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6200
         TabIndex        =   3
         Top             =   1200
         Width           =   600
      End
      Begin VB.CommandButton cmdsubir 
         Caption         =   "Subir"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6200
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   720
         Width           =   600
      End
      Begin MSFlexGridLib.MSFlexGrid gridpro 
         Height          =   2175
         Left            =   0
         TabIndex        =   1
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
         TabIndex        =   19
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
         TabIndex        =   10
         Top             =   840
         Width           =   4095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Producto :"
      Height          =   1335
      Left            =   0
      TabIndex        =   6
      Top             =   240
      Width           =   6615
      Begin VB.TextBox i_codart2 
         Height          =   285
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label i_nomarti 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   6255
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1575
      Left            =   3240
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
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
         TabIndex        =   18
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
         TabIndex        =   16
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   840
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Producto Terminado"
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
      TabIndex        =   9
      Top             =   0
      Width           =   2655
   End
End
Attribute VB_Name = "Frmformulacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VAR_ACTIVAR As Integer
Dim VAR
Dim loc_key As Integer
Dim cod As String
Dim kPASE As Boolean
Dim Flag_Consis As String * 1
Dim Flag_F2 As String * 1
Dim Flag_Bloq As String * 1
Dim Flag_Inicial As String * 1
Dim Flag_Change  As String * 1
Dim PS_REP02 As rdoQuery
Dim llave_rep02 As rdoResultset
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset





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

End Sub

Private Sub Form_Load()
loc_key = 0
Dim cade
pub_cadena = "SELECT * FROM PROCESOS WHERE PRO_CODCIA = ? AND PRO_CODART = ? AND PRO_CODPRO = ? AND PRO_SEC = ? ORDER BY PRO_SEC"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA=? ORDER BY ART_NOMBRE"
Set PSREP01 = CN.CreateQuery("", pub_cadena)
PSREP01(0) = LK_CODCIA
Set llave_rep01 = PSREP01.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)

kPASE = False
LISTA_ARTI
gridpro.Enabled = False
Screen.MousePointer = 0


End Sub

Public Sub LISTA_ARTI()
Dim ca, WF
Dim loc_arti As rdoResultset
cad = "SELECT * FROM ARTI WHERE ART_KEY > 0 AND ART_CODCIA = '" & LK_CODCIA & "' ORDER BY ART_NOMBRE"
Set loc_arti = CN.OpenResultset(cad, rdOpenKeyset, rdConcurValues)
loc_arti.Requery

Do Until loc_arti.EOF
  If loc_arti!art_familia = 1 Then
  LISARTI.AddItem loc_arti!ART_NOMBRE & String(60, " ") & loc_arti!ART_KEY
  End If
  loc_arti.MoveNext
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
' LISPROCESOS_Click
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
 Dim Wsec As Integer
 Dim WCODPRO As Integer
 Dim WCODART As Integer
 pub_mensaje = "Eliminar este Articulo del Proceso " & Trim(i_nomarti.Caption) & " ..?"
 Pub_Respuesta = MsgBox(pub_mensaje, vbInformation + vbYesNo + vbDefaultButton2, Pub_Titulo)
 If Pub_Respuesta = vbNo Then
   Screen.MousePointer = 0
   Exit Sub
 End If
 Screen.MousePointer = 11
 Wsec = Val(gridpro.TextMatrix(gridpro.Row, 6))
 WCODPRO = Val(gridpro.TextMatrix(gridpro.Row, 7))
 WCODART = Val(gridpro.TextMatrix(gridpro.Row, 8))
 PS_REP02(0) = LK_CODCIA
 PS_REP02(1) = WCODART
 PS_REP02(2) = WCODPRO
 PS_REP02(3) = Wsec
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
If KeyAscii = 13 And gridpro.COL = 5 Then
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
Dim Wsec
Static WORDEN
Dim wCODARTI

wCODARTI = Val(Trim(Right(LISARTI.Text, 8)))
If wCODARTI = 0 Or kPASE = False Then
  MsgBox "Debe Seleccionar Proceso ..", 48, Pub_Titulo
   i_codart2.SetFocus
  Exit Sub
End If
PUB_CODPRO = Val(Trim(Right(i_nomarti.Caption, 10)))
Screen.MousePointer = 11

pub_cadena = "SELECT * FROM PROCESOS WHERE PRO_CODCIA = ? AND PRO_CODPRO = ? ORDER BY PRO_CODCIA, PRO_SEC"
Set PSPROC_MAYOR2 = CN.CreateQuery("", pub_cadena)
PSPROC_MAYOR2(0) = 0
PSPROC_MAYOR2(1) = 0
Set proc_mayor2 = PSPROC_MAYOR2.OpenResultset(rdOpenKeyset, rdConcurValues)
PSPROC_MAYOR2(0) = LK_CODCIA
PSPROC_MAYOR2(1) = PUB_CODPRO
proc_mayor2.Requery
Do Until proc_mayor2.EOF
 If Val(proc_mayor2!PRO_CODART) = wCODARTI Then
   Screen.MousePointer = 0
   MsgBox " YA Existe en la lista. ", 48, Pub_Titulo
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
  Wsec = 1
Else
  proc_mayor2.MoveLast
  Wsec = proc_mayor2!PRO_SEC
  WORDEN = proc_mayor2!PRO_ORDEN
  WORDEN = WORDEN + 1
  Wsec = Wsec + 1
End If
proc_mayor2.AddNew
proc_mayor2!pro_codpro = PUB_CODPRO
proc_mayor2!PRO_CODART = Val(wCODARTI)
proc_mayor2!PRO_CODCIA = LK_CODCIA
proc_mayor2!PRO_SEC = Wsec
proc_mayor2!PRO_ORDEN = WORDEN
proc_mayor2.Update
gridpro.Rows = gridpro.Rows + 1
gridpro.RowHeight(gridpro.Rows - 1) = 285
gridpro.TextMatrix(gridpro.Rows - 1, 0) = gridpro.Rows - 1
gridpro.TextMatrix(gridpro.Rows - 1, 1) = Left(LISARTI.Text, 30)
gridpro.TextMatrix(gridpro.Rows - 1, 2) = 0
gridpro.TextMatrix(gridpro.Rows - 1, 3) = arm_llave!ARM_COSPRO
gridpro.TextMatrix(gridpro.Rows - 1, 4) = 0
gridpro.TextMatrix(gridpro.Rows - 1, 5) = arm_llave!ARM_STOCK
gridpro.TextMatrix(gridpro.Rows - 1, 6) = Wsec
gridpro.TextMatrix(gridpro.Rows - 1, 7) = PUB_CODPRO
gridpro.TextMatrix(gridpro.Rows - 1, 8) = Val(wCODARTI)

Screen.MousePointer = 0
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
txtPeso.Text = Format(SUMPESO, "Currency")
txttotal.Text = Format(SUMTOTAL, "Currency")
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
cadetodo = " SELECT ARTI.ART_NOMBRE, PROCESOS.PRO_SEC, PROCESOS.PRO_CODART, PROCESOS.PRO_CODPRO, PROCESOS.PRO_CANTIDAD, ARTI.ART_COSPRO, PROCESOS.PRO_SUBTOTAL,ARTICULO.ARM_STOCK, ARTICULO.ARM_COSPRO, PROCESOS.PRO_ORDEN, PROCESOS.PRO_SEC From  PROCESOS, ARTICULO, ARTI  Where  PROCESOS.PRO_CODART = ARTICULO.ARM_CODART  AND  PROCESOS.PRO_CODCIA = ARTICULO.ARM_CODCIA   AND PROCESOS.PRO_CODART = ARTI.ART_KEY  AND PROCESOS.PRO_CODCIA = ARTI.ART_CODCIA  AND " & csele
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
    gridpro.TextMatrix(fila, 3) = llave_rep01!ARM_COSPRO
    gridpro.TextMatrix(fila, 4) = Nulo_Valor0(llave_rep01!PRO_SUBTOTAL)
    gridpro.TextMatrix(fila, 5) = llave_rep01!ARM_STOCK
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
wsTexto.Tag = wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)
wsTexto.Text = wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)
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
    If wsGrid.COL <> 6 Then
     wsGrid.COL = wsGrid.COL + 1
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
wsTexto.Text = ""
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
If Left(Trim(wsGrid.Text), 1) = "-" Then
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
 If wsGrid.COL <> 0 Then
  wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL) = Format(wsTexto.Text, "0.0000")
 End If
 gridpro.TextMatrix(wsGrid.Row, 4) = Format(Val(wsTexto.Text) * Val(gridpro.TextMatrix(wsGrid.Row, 3)), "0.00")
 Screen.MousePointer = 11
 GRABA_DATO Val(wsGrid.TextMatrix(wsGrid.Row, wsGrid.COL)), Val(gridpro.TextMatrix(wsGrid.Row, 4))
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
  If wsGrid.COL <> wsGrid.FixedCols Then
     If wsGrid.COL = 3 Then
        wsGrid.COL = wsGrid.COL - 2
     Else
        wsGrid.COL = wsGrid.COL - 1
     End If
  End If
 End If
 If wsKeyCode = 39 Then ' derecha
  If Not IsMissing(SaltaCol) Then
     If wsGrid.COL = SaltaCol Then
        If wsGrid.Row <> wsGrid.Rows - 1 Then
          wsGrid.Row = wsGrid.Row + 1
          wsGrid.COL = wsGrid.FixedCols
          GoTo wsfinal
        ElseIf wsGrid.Row = wsGrid.Rows - 1 And wsGrid.COL = wsGrid.Cols - 1 Then
         If Trim(wsGrid.TextMatrix(wsGrid.Row, 0)) <> "" And Val(wsGrid.TextMatrix(wsGrid.Row, 1)) <> 0 And Val(wsGrid.TextMatrix(wsGrid.Row, 3)) <> 0 Then
          ' wsGrid.Rows = wsGrid.Rows + 1
           'wsGrid.Row = wsGrid.Row + 1
         '  wsGrid.Col = wsGrid.FixedCols
         '  GoTo wsfinal
         Else
           wsGrid.COL = wsGrid.FixedCols
          GoTo wsfinal
         End If
        End If
     ElseIf wsGrid.Row = wsGrid.Rows - 1 And wsGrid.COL = wsGrid.Cols - 1 Then
        wsGrid.COL = wsGrid.FixedCols
        GoTo wsfinal
     ElseIf wsGrid.COL = 2 And wsGrid.Row <> wsGrid.Rows - 1 Then
       wsGrid.Row = wsGrid.Row + 1
       Exit Sub
     End If
  End If
  If wsGrid.COL <> wsGrid.Cols - 1 Then
     If wsGrid.COL = 1 Then
      wsGrid.COL = wsGrid.COL + 2
     ElseIf wsGrid.COL >= 13 Then
'      wsGrid.Row = wsGrid.Row + 1
      wsGrid.COL = wsGrid.FixedCols
       GoTo wsfinal
     Else
      wsGrid.COL = wsGrid.COL + 1
     End If
  End If
 End If
wsfinal:
 wsTexto.FontBold = False
 wsTexto.ForeColor = QBColor(0)
 wsTexto.Text = ""
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
  If wsGrid.COL = ConsisCol1 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal1, ConsisCol1
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol2) And Not IsMissing(ConsisVal2) Then
  If wsGrid.COL = ConsisCol2 Then '
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
If Not IsMissing(ConsisCol4) And Not IsMissing(ConsisVal4) Then
  If wsGrid.COL = ConsisCol4 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal4, ConsisCol4
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol5) And Not IsMissing(ConsisVal5) Then
  If wsGrid.COL = ConsisCol5 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal5, ConsisCol5
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol6) And Not IsMissing(ConsisVal6) Then
  If wsGrid.COL = ConsisCol6 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal6, ConsisCol6
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol7) And Not IsMissing(ConsisVal7) Then
  If wsGrid.COL = ConsisCol7 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal7, ConsisCol7
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol8) And Not IsMissing(ConsisVal8) Then
  If wsGrid.COL = ConsisCol8 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal8, ConsisCol8
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol9) And Not IsMissing(ConsisVal9) Then
  If wsGrid.COL = ConsisCol9 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal9, ConsisCol9
   Exit Sub
  End If
End If
If Not IsMissing(ConsisCol10) And Not IsMissing(ConsisVal10) Then
  If wsGrid.COL = ConsisCol10 Then
   Consistencias wsGrid, wsTexto, wsKeyAscii, ConsisVal10, ConsisCol10
   Exit Sub
  End If
End If

End Sub

Private Sub LISARTI_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> 114 Then Exit Sub
Dim wbus  As String
wbus = InputBox("Buscar :", "Busqueda de Insumos...", "")
If wbus = "" Then Exit Sub
LISARTI.Visible = False
For fila = 0 To LISARTI.ListCount - 1
  LISARTI.ListIndex = fila
  If UCase(Trim(Left(LISARTI.Text, 40))) Like "*" & UCase(wbus) & "*" Then
    LISARTI.Visible = True
    LISARTI.SetFocus
    Exit Sub
  End If
Next fila
LISARTI.Visible = True
LISARTI.SetFocus
MsgBox "NO encontrado.", 48, Pub_Titulo
LISARTI.SetFocus


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
Dim Wsec As Integer
Dim WCODPRO As Integer
Dim WCODART As Double
Wsec = Val(gridpro.TextMatrix(gridpro.Row, 6))
WCODPRO = Val(gridpro.TextMatrix(gridpro.Row, 7))
WCODART = Val(gridpro.TextMatrix(gridpro.Row, 8))
PS_REP02(0) = LK_CODCIA
PS_REP02(1) = WCODART
PS_REP02(2) = WCODPRO
PS_REP02(3) = Wsec
llave_rep02.Requery
If llave_rep02.EOF Then
 Exit Sub
End If
llave_rep02.Edit
llave_rep02!PRO_CANTIDAD = wcantidad
llave_rep02!PRO_SUBTOTAL = WSUBTOTAL
llave_rep02.Update

End Sub
Private Sub i_codart2_Change()
If i_codart2.Text = "" Then
 i_nomarti.Caption = ""
  VAR_ACTIVAR = 0
End If
gridpro.Clear
gridpro.Cols = 9
gridpro.Rows = 2

End Sub

Private Sub i_codart2_GotFocus()
cmdsubir.Enabled = False
cmdbajar.Enabled = False
'kPASE = False
'Azul i_codart2, i_codart2
'i_codart2.text = ""
'i_nomarti.Caption = ""
End Sub
Private Sub i_codart2_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If Not ListView1.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And i_codart2.Text = "" Then
  loc_key = 1
  Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > ListView1.ListItems.count Then loc_key = ListView1.ListItems.count
 GoTo POSICION
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo POSICION
End If
GoTo fin
POSICION:
'  KeyCode = 0
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
  DoEvents
  i_codart2.SelStart = Len(i_codart2.Text)
  DoEvents
fin:

End Sub
Private Sub i_codart2_KeyPress(KeyAscii As Integer)
Dim WCOD_ORIGINAL As Currency
Dim valor As String
Dim tf As Integer
Dim i, car
Dim itmFound As ListItem
car = Chr(KeyAscii)
KeyAscii = Asc(UCase(car))
If KeyAscii = 27 Then
 ListView1.Visible = False
 i_codart2.Text = ""
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
VAR_ACTIVAR = 0
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  PUB_KEY = 0
Else
 On Error GoTo mucho
 PUB_KEY = Val(i_codart2.Text)
 On Error GoTo 0
 If Len(i_codart2.Text) = 0 Then
    Exit Sub
 End If
 If IsNumeric(i_codart2.Text) = False Then
   PUB_KEY = 0
 End If
End If

If PUB_KEY <> 0 Then
    SQ_OPER = 1
    PUB_KEY = i_codart2.Text
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    If art_LLAVE.EOF Then
       MsgBox "Codigo NO Existe.", 48, Pub_Titulo
       Azul i_codart2, i_codart2
       GoTo fin
    End If
    If Val(art_LLAVE!art_familia) <> 1 Then
       MsgBox "Producto no es Producto Terminado", 48, Pub_Titulo
       Exit Sub
     End If

    WCOD_ORIGINAL = art_LLAVE!ART_KEY
    i_nomarti.Caption = Trim(art_LLAVE!ART_NOMBRE) & String(180, " ") & WCOD_ORIGINAL
    ListView1.Visible = False
    LLAMA_INSU WCOD_ORIGINAL
    'Pantalla.SetFocus
    Exit Sub
Else
  If ListView1.Visible = False And VAR_ACTIVAR <> 99 And i_codart2.Text <> "" And LK_FLAG_ORIGINAL <> "A" And LK_FLAG_ALTERNO = "A" Then
IR_ALTERNO:
     SQ_OPER = 3
     pu_alterno = i_codart2.Text
     pu_codcia = LK_CODCIA
     LEER_ART_LLAVE
     If art_llave_alt.EOF Then
       MsgBox "Codigo No Existe ...", 48, Pub_Titulo
       Azul i_codart2, i_codart2
       Exit Sub
     End If
     If Val(art_llave_alt!art_familia) <> 4 Then
       MsgBox "Producto no es Producto Terminado", 48, Pub_Titulo
       Exit Sub
     End If
     WCOD_ORIGINAL = art_llave_alt!ART_KEY
     i_nomarti.Caption = Trim(art_llave_alt!ART_NOMBRE) & String(180, " ") & WCOD_ORIGINAL
     ListView1.Visible = False
     LLAMA_INSU WCOD_ORIGINAL
     Exit Sub
  Else
    If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
     Exit Sub
    End If
    valor = UCase(ListView1.ListItems.Item(loc_key).Text)
    If Trim(UCase(i_codart2.Text)) = Left(valor, Len(Trim(i_codart2.Text))) And Len(Trim(i_codart2.Text)) <> 0 Then
      If VAR_ACTIVAR = 0 And LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
        i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key))
        GoTo IR_ALTERNO
      End If
      If VAR_ACTIVAR <> 99 Then
       i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
      Else
       i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key))
      End If
      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
       PUB_KEY = Val(ListView1.ListItems.Item(loc_key).SubItems(1))
      Else
       PUB_KEY = i_codart2.Text
      End If
      LEER_ART_LLAVE
      VAR_ACTIVAR = 0
      If art_LLAVE.EOF Then
        MsgBox "Codigo No Existe ...", 48, Pub_Titulo
        Azul i_codart2, i_codart2
        Exit Sub
      End If
      If Val(art_LLAVE!art_familia) <> 1 Then
       MsgBox "Producto no es Producto Terminado", 48, Pub_Titulo
       Exit Sub
      End If
      WCOD_ORIGINAL = art_LLAVE!ART_KEY
      i_nomarti.Caption = Trim(art_LLAVE!ART_NOMBRE) & String(180, " ") & WCOD_ORIGINAL
      i_codart2.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
      ListView1.Visible = False
      LLAMA_INSU WCOD_ORIGINAL
      Exit Sub
    Else
      Exit Sub
    End If
    
  End If
End If
dale:
ListView1.Visible = False
fin:
mucho:
Exit Sub
ERROR_CODIGO:
MsgBox "Codigo NO Valido .... ", 48, Pub_Titulo
Azul i_codart2, i_codart2
  

End Sub

Private Sub i_codart2_KeyUp(KeyCode As Integer, Shift As Integer)
Dim VAR
If KeyCode = 13 Then Exit Sub
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
   If Len(i_codart2.Text) = 0 Or i_codart2.Text = "" Then
    ListView1.Visible = False
    Exit Sub
  End If
  If i_codart2.Text = "*" And KeyCode = 106 Then
   VAR_ACTIVAR = 99
   Exit Sub
  ElseIf i_codart2.Text = "" Then
   VAR_ACTIVAR = 0
   Exit Sub
  End If
  If VAR_ACTIVAR <> 99 Then
    Exit Sub
  End If
  If Left(i_codart2.Text, 1) = "*" Then
   i_codart2.Text = Mid(i_codart2.Text, 2, Len(i_codart2.Text))
   i_codart2.SelStart = Len(i_codart2.Text)
  End If
Else
 If Len(i_codart2.Text) = 0 Or IsNumeric(i_codart2.Text) = True Then
   ListView1.Visible = False
   Exit Sub
 End If
End If
If ListView1.Visible = False And KeyCode <> 13 Or Len(i_codart2.Text) = 1 Then
    VAR = Asc(i_codart2.Text)
    VAR = VAR + 1
    If VAR = 33 Or VAR = 91 Then
       VAR = "ZZZZZZZZ"
    Else
       VAR = Chr(VAR)
    End If
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
      numarchi = 3
      archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO FROM ARTI WHERE  ART_FAMILIA = 4 AND ART_CODCIA = '" & LK_CODCIA & "' AND ART_ALTERNO BETWEEN '" & i_codart2.Text & "' AND  '" & VAR & "' ORDER BY ART_ALTERNO"
    Else
      numarchi = 0
      archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO FROM ARTI WHERE  ART_FAMILIA = 4 AND  ART_CODCIA = '" & LK_CODCIA & "' AND ART_NOMBRE BETWEEN '" & i_codart2.Text & "' AND  '" & VAR & "' ORDER BY ART_NOMBRE"
    End If
    PROC_LISVIEW ListView1
    loc_key = 0
    If ListView1.Visible Then
     loc_key = 1
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As ListItem    ' Variable FoundItem.
If ListView1.Visible Then
  Set itmFound = ListView1.FindItem(LTrim(i_codart2.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > ListView1.ListItems.count Then
      ListView1.ListItems.Item(ListView1.ListItems.count).EnsureVisible
   Else
     ListView1.ListItems.Item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  Exit Sub
End If



End Sub



Private Sub i_codart2_LostFocus()
ListView1.Visible = False
End Sub


Public Sub LLAMA_INSU(WCODI As Currency)
On Error GoTo SALE
cod = WCODI
If cod = 0 Then
  MsgBox "Tiene que Seleccionar un Proceso ..", 48, Pub_Titulo
  i_codart2.SetFocus
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
'Unload Frmformulacion
Exit Sub

End Sub
