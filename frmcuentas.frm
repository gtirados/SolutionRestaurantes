VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmcuentas 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Control de Cuentas"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6195
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   234
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   413
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdprint 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Imprimir"
      Height          =   1095
      Left            =   9240
      Picture         =   "frmcuentas.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.OptionButton Opcion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Automatico"
      Height          =   255
      Index           =   1
      Left            =   3360
      TabIndex        =   12
      Top             =   960
      Width           =   1455
   End
   Begin VB.OptionButton Opcion 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Manual"
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   0
      Top             =   600
      Width           =   1455
   End
   Begin VB.CommandButton cmdcerrar 
      Caption         =   "Cerrar"
      Height          =   735
      Left            =   11400
      TabIndex        =   9
      Top             =   480
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Actualizar"
      Height          =   495
      Left            =   5040
      TabIndex        =   1
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
      Left            =   6960
      TabIndex        =   4
      Top             =   600
      Width           =   1935
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   720
      Width           =   255
   End
   Begin VB.TextBox txts 
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
   Begin VB.Timer tiempo 
      Interval        =   1000
      Left            =   120
      Top             =   5520
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   7575
      Left            =   6240
      TabIndex        =   14
      Top             =   1800
      Width           =   8775
      _ExtentX        =   15478
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
   Begin MSFlexGridLib.MSFlexGrid grid2 
      Height          =   7575
      Left            =   240
      TabIndex        =   15
      Top             =   1800
      Width           =   6015
      _ExtentX        =   10610
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
   Begin MSMask.MaskEdBox fecha1 
      Height          =   375
      Left            =   240
      TabIndex        =   16
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox fecha2 
      Height          =   375
      Left            =   1800
      TabIndex        =   13
      Top             =   600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   393216
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "DETALLE DE CUENTA"
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
      Left            =   5400
      TabIndex        =   11
      Top             =   1320
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "CUENTA POR MESA"
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
      TabIndex        =   10
      Top             =   1320
      Width           =   8535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Segundos"
      Height          =   255
      Index           =   1
      Left            =   5760
      TabIndex        =   8
      Top             =   720
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Opcion de Ejecucion"
      Height          =   255
      Index           =   0
      Left            =   3360
      TabIndex        =   7
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label lbltiempo 
      Height          =   735
      Left            =   2400
      TabIndex        =   6
      Top             =   7680
      Width           =   1335
   End
   Begin VB.Label lblmensa 
      Height          =   615
      Left            =   1320
      TabIndex        =   5
      Top             =   4560
      Visible         =   0   'False
      Width           =   4935
   End
End
Attribute VB_Name = "frmcuentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset
Dim PS_REP02 As rdoQuery
Dim llave_rep02 As rdoResultset
Dim PS_REP03 As rdoQuery
Dim llave_rep03 As rdoResultset
Dim PS_REP04 As rdoQuery
Dim llave_rep04 As rdoResultset
Dim PS_CON01 As rdoQuery
Dim llave_con01 As rdoResultset
Dim PS_CON02 As rdoQuery
Dim llave_con02 As rdoResultset
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

Public Function REP_CONSUL() As Integer
Dim WMONEDA As String * 1
Dim wser As String * 3
Dim WSRUTA As String
Dim indice As Integer
Dim wm As Integer
Dim llave_rep01 As rdoResultset
Dim PS_REP01 As rdoQuery
Dim i As Integer
Dim valor
Dim loc_xl As Object
Dim loc_codtra As Integer
Dim wRuta As String
Dim WSNUMDOC As String
Dim numero_device As Integer
'If LK_EMP = "HER" Then
'  wRuta = "C:\ADMIN\STANDAR\"
'Else
LOC_TIPMOV = 201
If LK_EMP_PTO = "A" Then
  wRuta = PUB_RUTA_OTRO & "PTOVTA\"
Else
  wRuta = PUB_RUTA_OTRO
End If
'If Left(moneda.Text, 1) = "S" Then
 WMONEDA = "S"
'Else
' WMONEDA = "D"
'End If

'End If
  If Trim(Nulo_Valors(par_llave!PAR_DEVICE_FBG)) <> "" Then
     'numero_device = 0
     'Reportes.PrinterName = Printers(numero_device).DeviceName
     'Reportes.PrinterDriver = Printers(numero_device).DriverName '"RASDD.DLL"
     'Reportes.PrinterPort = Printers(numero_device).Port
  End If
    FORM_COT.Reportes.Connect = PUB_ODBC
    FORM_COT.Reportes.Destination = crptToWindow  '= crptToPrinter
    FORM_COT.Reportes.WindowLeft = 2
    FORM_COT.Reportes.WindowTop = 70
    FORM_COT.Reportes.WindowWidth = 635
    FORM_COT.Reportes.WindowHeight = 390
    FORM_COT.Reportes.Formulas(1) = ""
    PUB_NETO = val(txttotal.Text)
    PU_NUMSER = val((tserie.Text))
    PU_NUMFAC = val((txtdoc.Text))
    
'    pub_cadena = "SELECT * FROM PEDIDOS WHERE PED_CODCIA = ? AND PED_TIPMOV = 201  ORDER BY  PED_NUMFAC DESC "
'    Set PSTEMP_MAYOR = CN.CreateQuery("", pub_cadena)
'    PSTEMP_MAYOR(0) = LK_CODCIA
'    PSTEMP_MAYOR.MaxRows = 1
'    Set temp_mayor = PSTEMP_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)
'
'    PSTEMP_MAYOR(0) = LK_CODCIA
'    temp_mayor.Requery
'
'    If temp_mayor.EOF Then
'    PU_NUMFAC = Val((txtdoc.Text))
'    Else
'    PU_NUMFAC = Val(Nulo_Valor0(temp_mayor!PED_NUMFAC))
'    End If
    FORM_COT.Reportes.Formulas(1) = ""
    FORM_COT.Reportes.Formulas(1) = "SON_EFECTIVO=  ' " & CONVER_LETRAS(PUB_NETO, WMONEDA) & "'"
    FORM_COT.Reportes.WindowTitle = "PEDIDO:" & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "00000000")
    FORM_COT.Reportes.ReportFileName = wRuta + "comanda.RPT"
    pub_cadena = "{PEDIDOS.PED_TIPMOV} = " & LOC_TIPMOV & " AND {PEDIDOS.PED_CODCIA} = '" & LK_CODCIA & "' AND  {PEDIDOS.PED_NUMSER} = '" & PU_NUMSER & "' AND {PEDIDOS.PED_NUMFAC} = " & PU_NUMFAC
    FORM_COT.Reportes.SelectionFormula = pub_cadena
    On Error GoTo accion
    FORM_COT.Reportes.Action = 1
  On Error GoTo 0
Exit Function
accion:
 MsgBox Err.Description
 MsgBox "Intente Nuevamente, la impresion de Modo manual", 48, Pub_Titulo
 Exit Function

End Function



Public Sub ACTUALIZA_GRID()
cabe
cabe1

End Sub




Private Sub cmdCerrar_Click()
Unload frmcuentas
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

Private Sub cmdprint_Click()
Dim intRpta As Integer
intRpta = MsgBox("Desea el imprimir esta Comanda?", vbQuestion + vbYesNo)
If intRpta = vbYes Then
    Call REP_CONSUL
 Else
 cmdCerrar_Click
End If
End Sub

Private Sub Command1_Click()


'pub_cadena = "SELECT SUM(CAR_IMPORTE)AS DEUDA, CAR_CODCLIE FROM CARTERA WHERE CAR_CODCIA = '02' AND CAR_CP = 'C' AND CAR_TIPDOC <> 'CH' AND CAR_IMPORTE <> 0  GROUP BY CAR_CODCLIE "
'Set PS_CON03 = CN.CreateQuery("", pub_cadena)
'Set llave_con03 = PS_CON03.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

'Do Until llave_con03.EOF
'  MsgBox llave_con03!CAR_CODCLIE & "    " & llave_con03!DEUDA
' llave_con03.MoveNext
'Loop


'Exit Sub

Dim wstotal As Currency
Dim WCONTADO As Currency
Dim WCREDITO As Currency

'pub_cadena = "SELECT * FROM COCINA WHERE CO_CODCIA = ? AND CO_FECHA >= ? AND CO_FECHA <= ? AND  CO_CODART = ? AND  CO_ESTADO <> 'E' ORDER BY CO_HORA ASC"
'Set PS_CON01 = CN.CreateQuery("", pub_cadena)
'PS_CON01(0) = 0
'PS_CON01(1) = LK_FECHA_DIA
'PS_CON01(2) = LK_FECHA_DIA
'PS_CON01(3) = 0
'Set llave_con01 = PS_CON01.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)

ACTUALIZA_GRID

'cabe
'grid1.Visible = False
'If Opcion(0).Value = True Then
'  If fechas() = False Then
'   Exit Sub
'  End If
'End If
'PS_REP01(0) = LK_CODCIA
'PS_REP01(1) = wsFECHA1
'PS_REP01(2) = wsFECHA2
'llave_rep01.Requery
'If llave_rep01.EOF Then
'  MsgBox "No Existe Datos", 48, Pub_Titulo
'  GoTo fin
'End If
'
'
'PS_CON01(0) = LK_CODCIA
'PS_CON01(1) = wsFECHA1
'PS_CON01(2) = wsFECHA2
'PS_CON01(3) = llave_rep01!CO_codart
'
'
''PS_CON02(0) = LK_CODCIA
''PS_CON02(1) = wsFECHA1
''PS_CON02(2) = wsFECHA2
'
'fila = 0
'grid1.Rows = 1
'
'Do Until llave_rep01.EOF
'  fila = fila + 1
'
'  grid1.Rows = grid1.Rows + 1
'
'  llave_con01.Requery
'
'  grid1.TextMatrix(fila, 0) = Trim(llave_rep01!CO_NOMART)
'  grid1.TextMatrix(fila, 1) = Trim(llave_rep01!CO_codart)
'  grid1.TextMatrix(fila, 2) = Format(llave_rep01!CO_CANTIDAD, "###,###,##0")
'  grid1.TextMatrix(fila, 3) = Format(llave_rep01!CO_mesa, "0")
'  grid1.TextMatrix(fila, 4) = Format(llave_rep01!CO_MESANEW, "0")
'  grid1.TextMatrix(fila, 5) = Format(llave_rep01!CO_CODVEN, "0")
'  grid1.TextMatrix(fila, 6) = Format(llave_rep01!CO_HORA, "0.00")
'  grid1.TextMatrix(fila, 8) = Trim(llave_rep01!CO_CODCIA)
'  grid1.TextMatrix(fila, 9) = Trim(llave_rep01!CO_numser)
'  grid1.TextMatrix(fila, 10) = Trim(llave_rep01!CO_numfac)
'  llave_rep01.MoveNext
'Loop
'  fila = fila + 1
'  grid1.Rows = grid1.Rows + 1
'
'
'  grid1.RowHeight(fila) = 250

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


'PS_REP01(0) = LK_CODCIA
'llave_rep01.Requery
'If llave_rep01.EOF Then
'  MsgBox "No Existe Datos", 48, Pub_Titulo
'  GoTo fin
'End If






fila = 0
grid2.Rows = 1

Do Until llave_rep02.EOF
  fila = fila + 1


  grid2.Rows = grid2.Rows + 1
  PS_CON02(0) = LK_CODCIA
  PS_CON02(1) = wsFECHA1
  PS_CON02(2) = wsFECHA2
  PS_CON02(3) = llave_rep02!PED_NUMFAC
    
  llave_con02.Requery
  
  pub_cadena = "SELECT * FROM PEDIDOS WHERE PED_CODCIA = ? and PED_FECHA >= ? and PED_FECHA <= ?  and PED_NUMFAC = ? AND PED_APROBADO <> 'S' ORDER BY PED_NUMFAC ASC"
Set PS_REP03 = CN.CreateQuery("", pub_cadena)
PS_REP03(0) = LK_CODCIA
PS_REP03(1) = wsFECHA1
PS_REP03(2) = wsFECHA2
PS_REP03(3) = llave_rep02!PED_NUMFAC
Set llave_rep03 = PS_REP03.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)
  llave_rep03.Requery
  

  
  grid2.TextMatrix(fila, 0) = Trim(llave_rep03!PED_CODCLIE) ' & " " & llave_rep01!CO_NOMART
  grid2.TextMatrix(fila, 1) = Format(llave_con02!total, "0.00")
  
  
'pub_cadena = "SELECT * FROM COCINA WHERE CO_CODCIA = ? and co_fecha >= ? and co_fecha <= ?  and co_mesa = ? AND CO_PAGO <> 'S' ORDER BY CO_NUMFAC ASC"
'Set PS_REP03 = CN.CreateQuery("", pub_cadena)
'PS_REP03(0) = LK_CODCIA
'PS_REP03(1) = wsFECHA1
'PS_REP03(2) = wsFECHA2
'PS_REP03(3) = llave_rep02!CO_MESA
'Set llave_rep03 = PS_REP03.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)
'
'
'
'  llave_rep03.Requery
  
  grid2.TextMatrix(fila, 2) = Trim(llave_rep03!PED_NUMFAC)
  grid2.TextMatrix(fila, 3) = " "
  
  
  llave_rep02.MoveNext
Loop
  fila = fila + 1
  grid2.Rows = grid2.Rows + 1
  
  
  grid2.RowHeight(fila) = 250

grid2.SetFocus


Exit Sub
fin:

End Sub

Private Sub Form_Load()
pub_cadena = "SELECT * FROM PEDIDOS WHERE PED_CODCIA = ? and PED_FECHA >= ? and PED_FECHA <= ?  AND PED_APROBADO <> 'S' ORDER BY PED_NUMFAC ASC "
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = 0
PS_REP01(2) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)


pub_cadena = "SELECT DISTINCT PED_NUMFAC,PED_CODCLIE FROM PEDIDOS WHERE PED_CODCIA = ? and PED_FECHA >= ? and PED_FECHA <= ?  AND PED_APROBADO <> 'S' ORDER BY PED_CODCLIE"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = 0
PS_REP02(1) = 0
PS_REP02(2) = 0
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)

pub_cadena = "SELECT * FROM PEDIDOS WHERE PED_CODCIA = ? and PED_FECHA >= ? and PED_FECHA <= ?  and PED_NUMFAC = ? AND PED_APROBADO <> 'S' ORDER BY PED_NUMFAC ASC"
Set PS_REP03 = CN.CreateQuery("", pub_cadena)
PS_REP03(0) = 0
PS_REP03(1) = 0
PS_REP03(2) = 0
PS_REP03(3) = 0
Set llave_rep03 = PS_REP03.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)


'pub_cadena = "SELECT SUM((CO_CANTIDAD/CO_EQUIV))AS TOTAL FROM COCINA WHERE CO_CODCIA = ? AND CO_FECHA >= ? AND CO_FECHA <= ? AND  CO_CODART = ? AND  CO_ESTADO <> 'E'  "
'pub_cadena = "SELECT CO_CANTIDAD,CO_MESA,CO_HORA,CO_CODVEN FROM COCINA WHERE CO_CODCIA = ? AND CO_FECHA >= ? AND CO_FECHA <= ? AND  CO_CODART = ? AND  CO_ESTADO <> 'E'  AND CO_PAGO <> 'S' ORDER BY CO_NUMFAC ASC"
'Set PS_CON01 = CN.CreateQuery("", pub_cadena)
'PS_CON01(0) = 0
'PS_CON01(1) = 0
'PS_CON01(2) = 0
'PS_CON01(3) = 0
'Set llave_con01 = PS_CON01.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)

'pub_cadena = "SELECT SUM(FAR_PRECIO*(FAR_CANTIDAD/FAR_EQUIV))AS TOTAL FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND FAR_CODVEN = ? AND FAR_DESCTO = 0 AND  FAR_ESTADO <> 'E' AND FAR_TIPMOV = 10 AND FAR_SIGNO_CAR = 1 AND FAR_DIAS <> 0 AND FAR_MONEDA = '" & Left(moneda.Text, 1) & "'"
'pub_cadena = "SELECT SUM((CO_CANTIDAD/CO_EQUIV))AS TOTAL FROM COCINA WHERE CO_CODCIA = ? AND CO_FECHA >= ? AND CO_FECHA <= ? AND  CO_CODART = ? AND  CO_ESTADO <> 'E'  "

pub_cadena = "SELECT sum(PED_CANTIDAD*PED_PRECIO) as total FROM PEDIDOS WHERE PED_CODCIA = ? AND PED_FECHA >= ? AND PED_FECHA <= ?  AND  PED_NUMFAC = ? AND PED_ESTADO <> 'E'  AND PED_APROBADO <> 'S'"
Set PS_CON02 = CN.CreateQuery("", pub_cadena)
PS_CON02(0) = 0
PS_CON02(1) = 0
PS_CON02(2) = 0
PS_CON02(3) = 0

Set llave_con02 = PS_CON02.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)



 grid1.Cols = 11
 grid1.ColAlignment(0) = 1
 grid1.ColWidth(0) = 3500
 grid1.ColWidth(1) = 0
 grid1.ColWidth(2) = 1000
 grid1.ColWidth(3) = 1000
 grid1.ColWidth(4) = 1000
 grid1.ColWidth(5) = 800
 grid1.ColWidth(6) = 1500
 grid1.ColWidth(7) = 600
 grid1.ColWidth(8) = 0
 grid1.ColWidth(9) = 0
 grid1.ColWidth(10) = 600
 grid1.Visible = False
 grid2.Cols = 4
 grid2.ColAlignment(0) = 1
 grid2.ColWidth(0) = 3200
 grid2.ColWidth(1) = 1200
 grid2.ColWidth(2) = 1200
 grid2.ColWidth(3) = 0
 
 
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
 grid1.TextMatrix(0, 0) = "Plato"
 grid1.TextMatrix(0, 1) = "Cod."
 grid1.TextMatrix(0, 2) = "Cant."
 grid1.TextMatrix(0, 3) = "P.Unit."
 grid1.TextMatrix(0, 4) = "Total"
 grid1.TextMatrix(0, 5) = "Estado"
 grid1.TextMatrix(0, 6) = "Hora"
 grid1.TextMatrix(0, 7) = "Mozo"
 grid1.TextMatrix(0, 8) = "Cia"
 grid1.TextMatrix(0, 9) = "Ser"
 grid1.TextMatrix(0, 10) = "Num"
 grid1.RowHeight(0) = 250
 'grid1.Rowwidth(1) = 50
End Sub

Public Sub cabe1()
 grid2.Clear
 grid2.TextMatrix(0, 0) = "Mesa"
 grid2.TextMatrix(0, 1) = "S/."
 grid2.TextMatrix(0, 2) = "Coman."
 grid2.TextMatrix(0, 3) = " "
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
'Dim WC
'Dim a, WF As Integer
'Dim tf, t, tC
'Dim SALE As Boolean
'Dim i As Integer
'Dim NroRows As Integer
''If KeyCode = 46 Then
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
'
'   'If Trim(grid1.TextMatrix(grid1.Row, 6)) = "" Then
'   '         grid1.TextMatrix(grid1.Row, 6) = "X"
'   'End If
'   '
'  ' pub_cadena = "UPDATE COCINA SET CO_ATENDIDO = 'A' WHERE CO_CODCIA = '" & LK_CODCIA & "' AND CO_NUMSER =  " & Trim(grid1.TextMatrix(grid1.Row, 8)) & "  AND CO_NUMFAC = " & Trim(grid1.TextMatrix(grid1.Row, 9)) & " and co_codart = " & Trim(grid1.TextMatrix(grid1.Row, 1)) & ""
'  ' CN.Execute pub_cadena, rdExecDirect
'   '  grid1.RowHeight(grid1.Row) = 1
'   NroRows = grid1.Rows - 1
'   For i = NroRows To 1 Step -1
'        If grid1.TextMatrix(i, 7) = "X" Then
'           ' grid1.RemoveItem (i + 1)
'            pub_cadena = "UPDATE COCINA SET CO_ATENDIDO = 'A' WHERE CO_CODCIA = '" & LK_CODCIA & "' AND CO_NUMSER =  " & Trim(grid1.TextMatrix(grid1.Row, 9)) & "  AND CO_NUMFAC = " & Trim(grid1.TextMatrix(grid1.Row, 10)) & " and co_codart = " & Trim(grid1.TextMatrix(grid1.Row, 1)) & ""
'            CN.Execute pub_cadena, rdExecDirect
'            'NroRows = NroRows - 1
'        End If
'   Next i
'
'   grid1.Row = grid1.Row
'   grid1.Refresh
'   Command1_Click
'
'   grid1.SetFocus
'   End If
'End If
''End If
''grid1.SetFocus
'Exit Sub
End Sub

Private Sub grid1_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = 32 Then
'        If Trim(grid1.TextMatrix(grid1.Row, 7)) = "" Then
'            grid1.TextMatrix(grid1.Row, 7) = "X"
'            BackColorRow grid1.Row, grid1, &HC0C0FF
'        Else
'            grid1.TextMatrix(grid1.Row, 7) = " "
'            BackColorRow grid1.Row, grid1, &HFFFFFF
'        End If
'    End If
End Sub

Private Sub grid1_KeyUp(KeyCode As Integer, Shift As Integer)
'Dim WC
'Dim a, WF As Integer
'Dim tf, t, tC
'Dim SALE As Boolean
'Dim i As Integer
'Dim NroRows As Integer
'If KeyCode = 46 Then
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
'
'   'If Trim(grid1.TextMatrix(grid1.Row, 6)) = "" Then
'   '         grid1.TextMatrix(grid1.Row, 6) = "X"
'   'End If
'   '
'  ' pub_cadena = "UPDATE COCINA SET CO_ATENDIDO = 'A' WHERE CO_CODCIA = '" & LK_CODCIA & "' AND CO_NUMSER =  " & Trim(grid1.TextMatrix(grid1.Row, 8)) & "  AND CO_NUMFAC = " & Trim(grid1.TextMatrix(grid1.Row, 9)) & " and co_codart = " & Trim(grid1.TextMatrix(grid1.Row, 1)) & ""
'  ' CN.Execute pub_cadena, rdExecDirect
'   '  grid1.RowHeight(grid1.Row) = 1
'   NroRows = grid1.Rows - 1
'   For i = NroRows To 1 Step -1
'        If grid1.TextMatrix(i, 7) = "X" Then
'           ' grid1.RemoveItem (i + 1)
'            pub_cadena = "UPDATE COCINA SET CO_ATENDIDO = 'A' WHERE CO_CODCIA = '" & LK_CODCIA & "' AND CO_NUMSER =  " & Trim(grid1.TextMatrix(grid1.Row, 9)) & "  AND CO_NUMFAC = " & Trim(grid1.TextMatrix(grid1.Row, 10)) & " and co_codart = " & Trim(grid1.TextMatrix(grid1.Row, 1)) & ""
'            CN.Execute pub_cadena, rdExecDirect
'            'NroRows = NroRows - 1
'        End If
'   Next i
'
'   grid1.Row = grid1.Row
'   grid1.Refresh
'   Command1_Click
'
'   grid1.SetFocus
'   End If
'End If
'End If
''grid1.SetFocus
'Exit Sub

End Sub

Private Sub grid2_Click()

cabe
grid1.Visible = True
If Opcion(0).Value = True Then
  If fechas() = False Then
   Exit Sub
  End If
End If
'PS_REP02(0) = LK_CODCIA
'PS_REP02(1) = wsFECHA1
'PS_REP02(2) = wsFECHA2
'llave_rep01.Requery
'If llave_rep01.EOF Then
'  MsgBox "No Existe Datos", 48, Pub_Titulo
'  'GoTo fin
'End If

pub_cadena = "SELECT * FROM PEDIDOS WHERE PED_CODCIA = ? and PED_FECHA >= ? and PED_FECHA <= ?  and PED_NUMFAC = ? AND PED_APROBADO <> 'S' ORDER BY PED_NUMFAC ASC"
Set PS_REP04 = CN.CreateQuery("", pub_cadena)
PS_REP04(0) = 0
PS_REP04(1) = 0
PS_REP04(2) = 0
PS_REP04(3) = 0
Set llave_rep04 = PS_REP04.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)


PS_REP04(0) = LK_CODCIA
PS_REP04(1) = wsFECHA1
PS_REP04(2) = wsFECHA2
If grid2.TextMatrix(grid2.Row, 2) = "" Then
Exit Sub
Else
PS_REP04(3) = grid2.TextMatrix(grid2.Row, 2)
End If
llave_rep04.Requery
'
''PS_CON02(0) = LK_CODCIA
''PS_CON02(1) = wsFECHA1
''PS_CON02(2) = wsFECHA2




fila = 0
grid1.Rows = 1

Do Until llave_rep04.EOF
  fila = fila + 1

  grid1.Rows = grid1.Rows + 1

  SQ_OPER = 1
    PUB_KEY = Trim(llave_rep04!PED_CODart)
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    'co_llave!CO_CANATEND = art_LLAVE!art_familia
    
   

  grid1.TextMatrix(fila, 0) = Trim(art_LLAVE!art_nombre)
  grid1.TextMatrix(fila, 1) = Trim(llave_rep04!PED_CODart)
  grid1.TextMatrix(fila, 2) = Format(llave_rep04!PED_CANTIDAD, "0")
  grid1.TextMatrix(fila, 3) = Format(llave_rep04!PED_PRECIO, "0.00")
  grid1.TextMatrix(fila, 4) = Format((Format(llave_rep04!PED_PRECIO, "0.00") * Format(llave_rep04!PED_CANTIDAD, "0.00")), "0.00")
  If Trim(llave_rep04!PED_ESTADO) = "E" Then
  grid1.TextMatrix(fila, 5) = "Anul."
  BackColorRow grid1.Row, grid1, &HC0C0FF
  Else
  grid1.TextMatrix(fila, 5) = "Norm."
  BackColorRow grid1.Row, grid1, &HFFFFFF
  End If
  grid1.TextMatrix(fila, 6) = Format(llave_rep04!PED_HORA, "0.00")
  grid1.TextMatrix(fila, 7) = Format(llave_rep04!PED_CODVEN, "0")
  grid1.TextMatrix(fila, 8) = Trim(llave_rep04!PED_CODCIA)
  grid1.TextMatrix(fila, 9) = Trim(llave_rep04!PED_numser)
  grid1.TextMatrix(fila, 10) = Trim(llave_rep04!PED_NUMFAC)
  llave_rep04.MoveNext
Loop
  fila = fila + 1
  grid1.Rows = grid1.Rows + 1


  grid1.RowHeight(fila) = 250
  'grid1.Rowwidth(fila) = 250

grid2.SetFocus
End Sub

Private Sub grid2_DblClick()
Dim WC
Dim a, WF As Integer
Dim tf, t, tC
Dim SALE As Boolean
Dim i As Integer
Dim NroRows As Integer
'If KeyCode = 46 Then
If grid2.Rows <= 0 Then Exit Sub
If grid2.Rows <= 1 Then
    pub_mensaje = MsgBox("Confirma pago? ", vbYesNo + vbExclamation + vbDefaultButton1, Pub_Titulo)
    If pub_mensaje = vbNo Then
      grid2.SetFocus
      Exit Sub
    End If
    cabe
Else
   pub_mensaje = MsgBox("Confirma pago? ", vbYesNo + vbExclamation + vbDefaultButton1, Pub_Titulo)
   If pub_mensaje = vbNo Then
      grid2.SetFocus
     Exit Sub
   Else
   
   If Trim(grid2.TextMatrix(grid2.Row, 3)) = "" Then
            grid2.TextMatrix(grid2.Row, 3) = "X"
   End If
   '
  ' pub_cadena = "UPDATE COCINA SET CO_ATENDIDO = 'A' WHERE CO_CODCIA = '" & LK_CODCIA & "' AND CO_NUMSER =  " & Trim(grid2.TextMatrix(grid2.Row, 8)) & "  AND CO_NUMFAC = " & Trim(grid2.TextMatrix(grid2.Row, 9)) & " and co_codart = " & Trim(grid2.TextMatrix(grid2.Row, 1)) & ""
  ' CN.Execute pub_cadena, rdExecDirect
   '  grid2.RowHeight(grid2.Row) = 1
   NroRows = grid2.Rows - 1
   For i = NroRows To 1 Step -1
        If grid2.TextMatrix(i, 3) = "X" Then
           ' grid2.RemoveItem (i + 1)
            pub_cadena = "UPDATE PEDIDOS SET PED_APROBADO = 'S' WHERE PED_CODCIA = '" & LK_CODCIA & "' AND PED_NUMSER =  100  AND PED_NUMFAC = " & Trim(grid2.TextMatrix(grid2.Row, 2)) & ""
            CN.Execute pub_cadena, rdExecDirect
            'NroRows = NroRows - 1
        End If
   Next i
   
   grid2.Row = grid2.Row
   grid2.Refresh
   Command1_Click
   
   grid2.SetFocus
   End If
End If
'End If
'grid2.SetFocus
Exit Sub
End Sub

Private Sub grid2_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 32 Then
        If Trim(grid2.TextMatrix(grid2.Row, 2)) = "" Then
            grid2.TextMatrix(grid1.Row, 2) = "X"
            BackColorRow grid2.Row, grid2, &HC0C0FF
        Else
            grid2.TextMatrix(grid1.Row, 2) = " "
            BackColorRow grid2.Row, grid2, &HFFFFFF
        End If
    End If
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
