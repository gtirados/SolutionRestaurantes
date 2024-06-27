VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmcocina 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Cocina en Tiempo Real"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13785
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7995
   ScaleWidth      =   13785
   WindowState     =   2  'Maximized
   Begin VB.OptionButton Opcion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Automatico"
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   16
      Top             =   960
      Width           =   1455
   End
   Begin VB.OptionButton Opcion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Manual"
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdcerrar 
      Caption         =   "Cerrar"
      Height          =   735
      Left            =   10920
      TabIndex        =   12
      Top             =   360
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Actualizar"
      Height          =   495
      Left            =   5520
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   4920
   End
   Begin VB.CommandButton cmdcom 
      Caption         =   "Comenzar"
      Height          =   495
      Left            =   7560
      TabIndex        =   7
      Top             =   600
      Width           =   1935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox txts 
      Height          =   375
      Left            =   5520
      TabIndex        =   5
      Top             =   720
      Width           =   495
   End
   Begin MSFlexGridLib.MSFlexGrid grid2 
      Height          =   7575
      Left            =   9720
      TabIndex        =   4
      Top             =   1680
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   13361
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer tiempo 
      Interval        =   1000
      Left            =   120
      Top             =   5520
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   7575
      Left            =   120
      TabIndex        =   2
      Top             =   1680
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   13361
      _Version        =   393216
      AllowBigSelection=   0   'False
      GridLinesFixed  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox fecha2 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox fecha1 
      Height          =   375
      Left            =   600
      TabIndex        =   15
      Top             =   720
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.Timer tmrInsumos 
      Interval        =   1000
      Left            =   9840
      Top             =   120
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "TOTALES POR ATENDER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   9720
      TabIndex        =   14
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "PLATOS POR DESPACHAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   1320
      Width           =   8535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Segundos"
      Height          =   255
      Index           =   1
      Left            =   6360
      TabIndex        =   11
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Opcion de Ejecucion"
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   10
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lbltiempo 
      Height          =   735
      Left            =   2400
      TabIndex        =   9
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label lblmensa 
      Height          =   615
      Left            =   1320
      TabIndex        =   8
      Top             =   4560
      Width           =   4935
   End
End
Attribute VB_Name = "frmcocina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset
Dim PS_REP02 As rdoQuery
Dim llave_rep02 As rdoResultset
Dim PS_CON01 As rdoQuery
Dim llave_con01 As rdoResultset
Dim PS_CON02 As rdoQuery
Dim ps_con04 As rdoQuery
Dim llave_con02 As rdoResultset
Dim llave_con04 As rdoResultset
Dim PS_CON03  As rdoQuery
Dim llave_con03  As rdoResultset
Dim wsFECHA1
Dim wsFECHA2


Public Function fechas() As Boolean
fechas = True
If Right(fecha1.Text, 2) = "__" Then
     wsFECHA1 = Left(fecha1.Text, 8)
Else
     wsFECHA1 = Trim(fecha1.Text)
End If
If Right(fecha2.Text, 2) = "__" Then
     wsFECHA2 = Left(fecha2.Text, 8)
Else
     wsFECHA2 = Trim(fecha2.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalida ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If Not IsDate(wsFECHA2) Then
 MsgBox "Fecha Invalida ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 MsgBox "Fecha Invalida ..", 48, Pub_Titulo
 GoTo CANCELA
End If

Exit Function
CANCELA:
 fechas = False
End Function


Public Sub ACTUALIZA_GRID()
cabe
cabe1

End Sub




Private Sub cmdCerrar_Click()
Unload frmcocina
End Sub

Private Sub cmdcom_Click()
If Left(cmdcom.Caption, 2) = "&D" Then
    cmdcom.Caption = "Comenzar"
    Timer1.Enabled = False
    tiempo.Enabled = False
    lbltiempo.Caption = ""
    VScroll1.Enabled = True
    txts.Enabled = True
    Opcion(0).Enabled = True
    Opcion(1).Enabled = True
    fecha1.Enabled = True
    fecha2.Enabled = True
Else
   If val(txts.Text) <= 0 Then
    MsgBox "Verificar Tiempo.", 48, Pub_Titulo
    Exit Sub
   End If
   If fechas() = False Then
     Exit Sub
   End If
   fecha1.Enabled = False
   fecha2.Enabled = False
   Opcion(0).Enabled = False
   Opcion(1).Enabled = False
   VScroll1.Enabled = False
   txts.Enabled = False
   DoEvents
   cmdcom.Caption = "&Detener"
   Timer1.Interval = val(txts.Text) * 1000
   Timer1.Enabled = True
   lbltiempo.Caption = txts.Text
   tiempo.Enabled = True
End If
End Sub

Private Sub Command1_Click()

Dim wstotal As Currency
Dim WCONTADO As Currency
Dim WCREDITO As Currency


ACTUALIZA_GRID

cabe
If Opcion(0).Value = True Then
  If fechas() = False Then
   Exit Sub
  End If
End If
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = wsFECHA1
PS_REP01(2) = wsFECHA2
llave_rep01.Requery
If llave_rep01.EOF Then
  MsgBox "No Existe Datos", 48, Pub_Titulo
  GoTo fin
End If


PS_CON01(0) = LK_CODCIA
PS_CON01(1) = wsFECHA1
PS_CON01(2) = wsFECHA2
PS_CON01(3) = llave_rep01!PED_CODart




Fila = 0
grid1.Rows = 1

Do Until llave_rep01.EOF
  Fila = Fila + 1

  grid1.Rows = grid1.Rows + 1
    
  llave_con01.Requery
  
  SQ_OPER = 1
    PUB_KEY = Trim(llave_rep01!PED_CODart)
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    'co_llave!CO_CANATEND = art_LLAVE!art_familia

  grid1.TextMatrix(Fila, 0) = Trim(art_LLAVE!art_nombre)
  
  'grid1.TextMatrix(fila, 0) = Trim(llave_rep01!PED_CODart)
  grid1.TextMatrix(Fila, 1) = Trim(llave_rep01!PED_CODart)
  grid1.TextMatrix(Fila, 2) = Format(llave_rep01!PED_CANTIDAD, "###,###,##0")
  grid1.TextMatrix(Fila, 3) = Format(llave_rep01!PED_CODCLIE, "0")
  grid1.TextMatrix(Fila, 4) = Format(llave_rep01!PED_CODCLIE, "0")
  grid1.TextMatrix(Fila, 5) = Format(llave_rep01!PED_CODVEN, "0")
  grid1.TextMatrix(Fila, 6) = Format(llave_rep01!PED_HORA, "0.00")
  grid1.TextMatrix(Fila, 8) = Trim(llave_rep01!PED_CODCIA)
  grid1.TextMatrix(Fila, 9) = Trim(llave_rep01!PED_numser)
  grid1.TextMatrix(Fila, 10) = Trim(llave_rep01!PED_NUMFAC)
  grid1.TextMatrix(Fila, 11) = Trim(llave_rep01!PED_numsec)
  grid1.TextMatrix(Fila, 12) = Trim(llave_rep01!PED_OFERTA)
  llave_rep01.MoveNext
Loop
  Fila = Fila + 1
  grid1.Rows = grid1.Rows + 1
  
  
  grid1.RowHeight(Fila) = 250

'grid1.SetFocus

cabe1

If Opcion(0).Value = True Then
  If fechas() = False Then
   Exit Sub
  End If
End If
PS_REP02(0) = LK_CODCIA
PS_REP02(1) = wsFECHA1
PS_REP02(2) = wsFECHA2
llave_rep02.Requery
If llave_rep02.EOF Then
  MsgBox "No Existe Datos", 48, Pub_Titulo
  GoTo fin
End If


PS_REP01(0) = LK_CODCIA
llave_rep01.Requery
If llave_rep01.EOF Then
  MsgBox "No Existe Datos", 48, Pub_Titulo
  GoTo fin
End If


PS_CON02(0) = LK_CODCIA
PS_CON02(1) = wsFECHA1
PS_CON02(2) = wsFECHA2



Fila = 0
grid2.Rows = 1

Do Until llave_rep02.EOF
  Fila = Fila + 1


  grid2.Rows = grid2.Rows + 1
  
  PS_CON02(3) = llave_rep02!PED_CODart
    
  llave_con02.Requery
  
    SQ_OPER = 1
    PUB_KEY = Trim(llave_rep02!PED_CODart)
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    

  grid2.TextMatrix(Fila, 0) = Trim(art_LLAVE!art_nombre)
  
  'grid2.TextMatrix(fila, 0) = Trim(llave_rep02!PED_CODart) ' & " " & llave_rep01!CO_NOMART)
  grid2.TextMatrix(Fila, 1) = Format(llave_con02!Total, "###,###,##0")
  
  
  llave_rep02.MoveNext
Loop
  Fila = Fila + 1
  grid2.Rows = grid2.Rows + 1
  
  
  grid2.RowHeight(Fila) = 250

grid2.SetFocus


Exit Sub
fin:

End Sub

Private Sub Form_Load()
pub_cadena = "SELECT * FROM PEDIDOS WHERE PED_CODCIA = ? and PED_fecha >= ? and PED_fecha <= ? and PED_FAMILIA = 1 AND PED_CANATEN = 0 AND PED_ESTADO = 'N' "
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = 0
PS_REP01(2) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)


pub_cadena = "SELECT DISTINCT PED_CODART  FROM PEDIDOS WHERE PED_CODCIA = ? and PED_fecha >= ? and PED_fecha <= ? and PED_FAMILIA = 1 AND PED_CANATEN = 0 AND PED_ESTADO = 'N'"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = 0
PS_REP02(1) = 0
PS_REP02(2) = 0
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)




'pub_cadena = "SELECT SUM((CO_CANTIDAD/CO_EQUIV))AS TOTAL FROM COCINA WHERE CO_CODCIA = ? AND CO_FECHA >= ? AND CO_FECHA <= ? AND  CO_CODART = ? AND  CO_ESTADO <> 'E'  "
pub_cadena = "SELECT PED_CANTIDAD,PED_CODCLIE,PED_HORA,PED_CODVEN FROM PEDIDOS WHERE PED_CODCIA = ? AND PED_FECHA >= ? AND PED_FECHA <= ? AND  PED_CODART = ? AND  PED_ESTADO = 'N' and PED_FAMILIA = 1 AND PED_CANATEN = 0 ORDER BY PED_HORA ASC"
Set PS_CON01 = CN.CreateQuery("", pub_cadena)
PS_CON01(0) = 0
PS_CON01(1) = 0
PS_CON01(2) = 0
PS_CON01(3) = 0
Set llave_con01 = PS_CON01.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)

'pub_cadena = "SELECT SUM(FAR_PRECIO*(FAR_CANTIDAD/FAR_EQUIV))AS TOTAL FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND FAR_CODVEN = ? AND FAR_DESCTO = 0 AND  FAR_ESTADO <> 'E' AND FAR_TIPMOV = 10 AND FAR_SIGNO_CAR = 1 AND FAR_DIAS <> 0 AND FAR_MONEDA = '" & Left(moneda.Text, 1) & "'"
'pub_cadena = "SELECT SUM((CO_CANTIDAD/CO_EQUIV))AS TOTAL FROM COCINA WHERE CO_CODCIA = ? AND CO_FECHA >= ? AND CO_FECHA <= ? AND  CO_CODART = ? AND  CO_ESTADO <> 'E'  "

pub_cadena = "SELECT sum(PED_CANTIDAD) as total  FROM PEDIDOS WHERE PED_CODCIA = ? AND PED_FECHA >= ? AND PED_FECHA <= ?  AND  PED_CODART = ? AND PED_ESTADO = 'N' and PED_FAMILIA = 1 AND PED_CANATEN = 0"
Set PS_CON02 = CN.CreateQuery("", pub_cadena)
PS_CON02(0) = 0
PS_CON02(1) = 0
PS_CON02(2) = 0
PS_CON02(3) = 0

Set llave_con02 = PS_CON02.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)



 grid1.Cols = 13
 grid1.ColAlignment(0) = 1
 grid1.ColWidth(0) = 3800
 grid1.ColWidth(1) = 0
 grid1.ColWidth(2) = 500
 grid1.ColWidth(3) = 500
 grid1.ColWidth(4) = 500
 grid1.ColWidth(5) = 500
 grid1.ColWidth(6) = 1600
 grid1.ColWidth(7) = 0
 grid1.ColWidth(8) = 0
 grid1.ColWidth(9) = 0
 grid1.ColWidth(10) = 600
 grid1.ColWidth(11) = 100
 grid1.ColWidth(12) = 2000
 grid2.Cols = 2
 grid2.ColAlignment(0) = 1
 grid2.ColWidth(0) = 4200
 grid2.ColWidth(1) = 800
 
 
 fecha1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 fecha1.Mask = "##/##/####"
 fecha2.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 fecha2.Mask = "##/##/####"
 tiempo.Enabled = False
 cabe
 cabe1
 
'moneda.ListIndex = 0
End Sub


Public Sub cabe()
 grid1.Clear
 grid1.TextMatrix(0, 0) = "PLATO"
 grid1.TextMatrix(0, 1) = "CODIGO"
 grid1.TextMatrix(0, 2) = "CANT."
 grid1.TextMatrix(0, 3) = "MES."
 grid1.TextMatrix(0, 4) = "NEW."
 grid1.TextMatrix(0, 5) = "MOZ."
 grid1.TextMatrix(0, 6) = "HORA"
 grid1.TextMatrix(0, 7) = ""
 grid1.TextMatrix(0, 8) = "CIA"
 grid1.TextMatrix(0, 9) = "SERIE"
 grid1.TextMatrix(0, 10) = "NUMERO"
 grid1.TextMatrix(0, 11) = "SEC"
 grid1.TextMatrix(0, 12) = "DETALLE"
 grid1.RowHeight(0) = 350
 'grid1.RowHeight(1) = 350
End Sub

Public Sub cabe1()
 grid2.Clear
 grid2.TextMatrix(0, 0) = "PLATO"
 grid2.TextMatrix(0, 1) = "CANT."
 grid2.RowHeight(0) = 250
 grid2.RowHeight(1) = 250
End Sub

Private Sub grid1_Click()
If Trim(grid1.TextMatrix(grid1.Row, 7)) = "" Then
            grid1.TextMatrix(grid1.Row, 7) = "X"
            BackColorRow grid1.Row, grid1, &HC0C0FF
        Else
            grid1.TextMatrix(grid1.Row, 7) = " "
            BackColorRow grid1.Row, grid1, &HFFFFFF
        End If
End Sub

Private Sub Grid1_DblClick()
Dim WC
Dim a, WF As Integer
Dim tf, t, tC
Dim SALE As Boolean
Dim i As Integer
Dim NroRows As Integer
'If KeyCode = 46 Then
If grid1.Rows <= 0 Then Exit Sub
If grid1.Rows <= 1 Then
    pub_mensaje = MsgBox("Confirma despacho? ", vbYesNo + vbExclamation + vbDefaultButton1, Pub_Titulo)
    If pub_mensaje = vbNo Then
      grid1.SetFocus
      Exit Sub
    End If
    cabe
Else
   pub_mensaje = MsgBox("Confirma despacho? ", vbYesNo + vbExclamation + vbDefaultButton1, Pub_Titulo)
   If pub_mensaje = vbNo Then
      grid1.SetFocus
     Exit Sub
   Else
   
   'If Trim(grid1.TextMatrix(grid1.Row, 6)) = "" Then
   '         grid1.TextMatrix(grid1.Row, 6) = "X"
   'End If
   '
  ' pub_cadena = "UPDATE COCINA SET CO_ATENDIDO = 'A' WHERE CO_CODCIA = '" & LK_CODCIA & "' AND CO_NUMSER =  " & Trim(grid1.TextMatrix(grid1.Row, 8)) & "  AND CO_NUMFAC = " & Trim(grid1.TextMatrix(grid1.Row, 9)) & " and co_codart = " & Trim(grid1.TextMatrix(grid1.Row, 1)) & ""
  ' CN.Execute pub_cadena, rdExecDirect
   '  grid1.RowHeight(grid1.Row) = 1
   NroRows = grid1.Rows - 1
   For i = NroRows To 1 Step -1
        If grid1.TextMatrix(i, 7) = "X" Then
           ' grid1.RemoveItem (i + 1)
            pub_cadena = "UPDATE PEDIDOS SET PED_CANATEN = PED_CANTIDAD WHERE PED_CODCIA = '" & LK_CODCIA & "' AND PED_NUMSER =  " & Trim(grid1.TextMatrix(grid1.Row, 9)) & "  AND PED_NUMFAC = " & Trim(grid1.TextMatrix(grid1.Row, 10)) & " and PED_codart = " & Trim(grid1.TextMatrix(grid1.Row, 1)) & " and PED_NUMSEC = " & Trim(grid1.TextMatrix(grid1.Row, 11)) & ""
            CN.Execute pub_cadena, rdExecDirect
            'NroRows = NroRows - 1
        End If
   Next i
   
   grid1.Row = grid1.Row
   grid1.Refresh
   Command1_Click
   
   grid1.SetFocus
   End If
End If
'End If
'grid1.SetFocus
Exit Sub
End Sub

Private Sub grid1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 32 Then
        If Trim(grid1.TextMatrix(grid1.Row, 7)) = "" Then
            grid1.TextMatrix(grid1.Row, 7) = "X"
            BackColorRow grid1.Row, grid1, &HC0C0FF
        Else
            grid1.TextMatrix(grid1.Row, 7) = " "
            BackColorRow grid1.Row, grid1, &HFFFFFF
        End If
    End If
End Sub

Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
Dim WC
Dim a, WF As Integer
Dim tf, t, tC
Dim SALE As Boolean
Dim i As Integer
Dim NroRows As Integer
If KeyCode = 46 Then
'If grid1.Rows <= 0 Then Exit Sub
'If grid1.Rows <= 1 Then
'    pub_mensaje = MsgBox("Confirma despacho? ", vbYesNo + vbExclamation + vbDefaultButton1, Pub_Titulo)
'    If pub_mensaje = vbNo Then
'      grid1.SetFocus
'      Exit Sub
'    End If
'    cabe
'Else
'   pub_mensaje = MsgBox("Confirma despacho? ", vbYesNo + vbExclamation + vbDefaultButton1, Pub_Titulo)
'   If pub_mensaje = vbNo Then
'      grid1.SetFocus
'     Exit Sub
'   Else
   
   'If Trim(grid1.TextMatrix(grid1.Row, 6)) = "" Then
   '         grid1.TextMatrix(grid1.Row, 6) = "X"
   'End If
   '
  ' pub_cadena = "UPDATE COCINA SET CO_ATENDIDO = 'A' WHERE CO_CODCIA = '" & LK_CODCIA & "' AND CO_NUMSER =  " & Trim(grid1.TextMatrix(grid1.Row, 8)) & "  AND CO_NUMFAC = " & Trim(grid1.TextMatrix(grid1.Row, 9)) & " and co_codart = " & Trim(grid1.TextMatrix(grid1.Row, 1)) & ""
  ' CN.Execute pub_cadena, rdExecDirect
   '  grid1.RowHeight(grid1.Row) = 1
   NroRows = grid1.Rows - 1
   For i = NroRows To 1 Step -1
'        If grid1.TextMatrix(i, 7) = "X" Then
'           ' grid1.RemoveItem (i + 1)
'            pub_cadena = "UPDATE COCINA SET CO_ATENDIDO = 'A' WHERE CO_CODCIA = '" & LK_CODCIA & "' AND CO_NUMSER =  " & Trim(grid1.TextMatrix(grid1.Row, 9)) & "  AND CO_NUMFAC = " & Trim(grid1.TextMatrix(grid1.Row, 10)) & " and co_codart = " & Trim(grid1.TextMatrix(grid1.Row, 1)) & ""
'            CN.Execute pub_cadena, rdExecDirect
'            'NroRows = NroRows - 1
'        End If
   Next i
   
   grid1.Row = grid1.Row
   grid1.Refresh
   Command1_Click
   
   grid1.SetFocus
   End If
'End If
'End If
'grid1.SetFocus
Exit Sub

End Sub

Private Sub opcion_Click(Index As Integer)
If Index = 0 Then
    VScroll1.Visible = False
    txts.Visible = False
    cmdcom.Visible = False
    Command1.Visible = True
ElseIf Index = 1 Then
    VScroll1.Visible = True
    txts.Visible = True
    txts.Text = "05"
    cmdcom.Visible = True
    Command1.Visible = False
End If
End Sub

Private Sub tiempo_Timer()
lbltiempo.Caption = Format(val(lbltiempo.Caption) - 1, "00")
If val(lbltiempo.Caption) = 0 Then
  lbltiempo.Caption = txts.Text
End If
End Sub

Private Sub Timer1_Timer()
DoEvents
lblmensa.Caption = "Actualizando Información. . ."
DoEvents
Command1_Click
lblmensa.Caption = ""
DoEvents
End Sub

Private Sub tmrInsumos_Timer()
pub_cadena = "select ar.art_nombre,a.arm_stock,ar.art_stock_min " & _
"from articulo a inner join arti ar on " & _
"a.arm_codcia = ar.art_codcia And a.arm_codart = ar.art_key " & _
"where ar.art_flag_stock='M' and a.arm_stock <= ar.art_stock_min " & _
"and ar.art_codcia = ? "

'pub_cadena = "SELECT sum(PED_CANTIDAD) as total  FROM PEDIDOS WHERE PED_CODCIA = ? AND PED_FECHA >= ? AND PED_FECHA <= ?  AND  PED_CODART = ? AND PED_ESTADO <> 'E' and PED_FAMILIA = 2 AND PED_CANATEN = 0"
Set ps_con04 = CN.CreateQuery("", pub_cadena)
'Set PS_CON02 = CN.CreateQuery("", pub_cadena)
ps_con04(0) = LK_CODCIA


'Set llave_con04 = PS_CON02.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)
Set llave_con04 = ps_con04.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)

If Not llave_con04.EOF Then
    If oFrmIns Is Nothing Then
        Set oFrmIns = New frmInsumos
        Set oFrmIns.oData = llave_con04
        oFrmIns.Show
    End If
'    Set frmInsumos.oData = llave_con04
'    frmInsumos.Show
End If
End Sub

Private Sub txts_Change()
If val(txts.Text) > 0 And val(txts.Text) < 66 Then VScroll1.Value = val(txts.Text)
End Sub

Private Sub txts_KeyPress(KeyAscii As Integer)
If KeyAscii = 48 Then
  KeyAscii = 0
  Exit Sub
End If
SOLO_ENTERO KeyAscii
End Sub

Private Sub txts_LostFocus()
If val(txts.Text) > 65 Then
  txts.Text = 65
End If
End Sub

Private Sub VScroll1_Change()
 txts.Text = VScroll1.Value
End Sub
