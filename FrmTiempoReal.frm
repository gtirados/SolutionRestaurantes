VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmTiempoReal 
   Caption         =   "Venta Acumulada "
   ClientHeight    =   6990
   ClientLeft      =   270
   ClientTop       =   1170
   ClientWidth     =   9015
   Icon            =   "FrmTiempoReal.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   9015
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tiempo 
      Interval        =   1000
      Left            =   120
      Top             =   5280
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   4800
   End
   Begin VB.Frame fra1 
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin VB.ComboBox moneda 
         Height          =   315
         ItemData        =   "FrmTiempoReal.frx":0442
         Left            =   1440
         List            =   "FrmTiempoReal.frx":044C
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton cmdcom 
         Caption         =   "Comenzar"
         Height          =   375
         Left            =   7080
         TabIndex        =   13
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Mostrar"
         Height          =   375
         Left            =   5400
         TabIndex        =   12
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox txts 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4680
         MaxLength       =   2
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   375
         Left            =   5160
         Max             =   1
         Min             =   65
         TabIndex        =   9
         Top             =   360
         Value           =   65
         Width           =   255
      End
      Begin VB.OptionButton opcion 
         Caption         =   "Automatico"
         Height          =   255
         Index           =   1
         Left            =   3240
         TabIndex        =   3
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton opcion 
         Caption         =   "Manual"
         Height          =   255
         Index           =   0
         Left            =   3240
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin MSMask.MaskEdBox fecha1 
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox fecha2 
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Moneda :"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lbl1 
         Caption         =   "Segundos"
         Height          =   255
         Index           =   1
         Left            =   5760
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lbl2 
         Caption         =   "Rango de Fecha :"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   1455
      End
      Begin VB.Label lbl1 
         Caption         =   "Opcion de Ejecución"
         Height          =   255
         Index           =   0
         Left            =   3000
         TabIndex        =   4
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Frame fra2 
      Caption         =   "Ventas acumuladas por Vendedor"
      Height          =   5895
      Left            =   0
      TabIndex        =   1
      Top             =   1080
      Width           =   8895
      Begin VB.CommandButton cmdcerrar 
         Caption         =   "Ce&rrar"
         Height          =   375
         Left            =   7200
         TabIndex        =   15
         Top             =   5400
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid grid1 
         Height          =   5055
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   8916
         _Version        =   393216
         Rows            =   4
         FixedRows       =   3
         FixedCols       =   0
      End
      Begin VB.Label lbltiempo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   16
         Top             =   3840
         Width           =   495
      End
      Begin VB.Label lblmensa 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   480
         TabIndex        =   14
         Top             =   2040
         Width           =   3255
      End
   End
End
Attribute VB_Name = "FrmTiempoReal"
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
Dim llave_con02 As rdoResultset
Dim PS_CON03  As rdoQuery
Dim llave_con03  As rdoResultset
Dim wsFECHA1
Dim wsFECHA2


Private Sub cmdCerrar_Click()
Unload FrmTiempoReal
End Sub

Private Sub cmdcom_Click()
If Left(cmdcom.Caption, 2) = "&D" Then
    cmdcom.Caption = "Comenzar"
    Timer1.Enabled = False
    tiempo.Enabled = False
    lblTiempo.Caption = ""
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
   lblTiempo.Caption = txts.Text
   tiempo.Enabled = True
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

cabe
If Opcion(0).Value = True Then
  If fechas() = False Then
   Exit Sub
  End If
End If
PS_REP01(0) = LK_CODCIA
llave_rep01.Requery
If llave_rep01.EOF Then
  MsgBox "No Existe Datos", 48, Pub_Titulo
  GoTo fin
End If


PS_CON01(0) = LK_CODCIA
PS_CON01(1) = wsFECHA1
PS_CON01(2) = wsFECHA2


PS_CON02(0) = LK_CODCIA
PS_CON02(1) = wsFECHA1
PS_CON02(2) = wsFECHA2

fila = 2
grid1.rows = 3
WCONTADO = 0
WCREDITO = 0
Do Until llave_rep01.EOF
  fila = fila + 1
  wstotal = 0
  grid1.rows = grid1.rows + 1
  grid1.TextMatrix(fila, 0) = Trim(Format(llave_rep01!VEM_CODVEN, "000") & " " & llave_rep01!VEM_NOMBRE)
  PS_CON01(3) = llave_rep01!VEM_CODVEN
  llave_con01.Requery
  grid1.TextMatrix(fila, 1) = Format(llave_con01!Total, "###,###,##0.00")
  wstotal = wstotal + Nulo_Valor0(llave_con01!Total)
  WCONTADO = WCONTADO + Nulo_Valor0(llave_con01!Total)
  PS_CON02(3) = llave_rep01!VEM_CODVEN
  llave_con02.Requery
  grid1.TextMatrix(fila, 3) = Format(llave_con02!Total, "###,###,##0.00")
  wstotal = wstotal + Nulo_Valor0(llave_con02!Total)
  WCREDITO = WCREDITO + Nulo_Valor0(llave_con02!Total)
  grid1.TextMatrix(fila, 2) = Format((llave_con01!Total * 100) / wstotal, "#,##0.00")
  grid1.TextMatrix(fila, 4) = Format((llave_con02!Total * 100) / wstotal, "#,##0.00")
  grid1.TextMatrix(fila, 5) = Format(wstotal, "#,###,##0.00")
  llave_rep01.MoveNext
Loop
  fila = fila + 1
  grid1.rows = grid1.rows + 1
  grid1.TextMatrix(fila, 1) = Format(WCONTADO, "###,###,##0.00")
  grid1.TextMatrix(fila, 3) = Format(WCREDITO, "###,###,##0.00")
  If (WCONTADO + WCREDITO) <> 0 Then
    grid1.TextMatrix(fila, 2) = Format((WCONTADO * 100) / (WCONTADO + WCREDITO), "#,##0.00")
  End If
  If (WCONTADO + WCREDITO) <> 0 Then
    grid1.TextMatrix(fila, 4) = Format((WCREDITO * 100) / (WCONTADO + WCREDITO), "#,##0.00")
  End If
  grid1.RowHeight(fila) = 350
  grid1.TextMatrix(fila, 5) = Format(WCONTADO + WCREDITO, "###,##0.00")
  If Left(moneda.Text, 1) = "S" Then
    grid1.TextMatrix(fila, 0) = "Total S/. - "
  Else
    grid1.TextMatrix(fila, 0) = "Total US$.- "
  End If
  
  
  grid1.TextMatrix(2, 1) = grid1.TextMatrix(fila, 1)
  grid1.TextMatrix(2, 3) = grid1.TextMatrix(fila, 3)
  grid1.TextMatrix(2, 2) = grid1.TextMatrix(fila, 2)
  grid1.TextMatrix(2, 4) = grid1.TextMatrix(fila, 4)
  grid1.TextMatrix(2, 5) = grid1.TextMatrix(fila, 5)
grid1.SetFocus

Exit Sub
fin:
End Sub

Private Sub Form_Load()

pub_cadena = "SELECT * FROM VEMAEST WHERE VEM_CODCIA = ? "
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)

 grid1.Cols = 6
 grid1.ColAlignment(0) = 1
 grid1.ColWidth(0) = 2500
 grid1.ColWidth(1) = 1000
 grid1.ColWidth(2) = 900
 grid1.ColWidth(3) = 1000
 grid1.ColWidth(4) = 900
 grid1.ColWidth(5) = 1200
 fecha1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 fecha1.Mask = "##/##/####"
 fecha2.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 fecha2.Mask = "##/##/####"
 tiempo.Enabled = False
cabe
moneda.ListIndex = 0
End Sub

Public Sub cabe()
 grid1.Clear
 grid1.TextMatrix(0, 0) = "Vendedor"
 grid1.TextMatrix(0, 1) = "  Ventas"
 grid1.TextMatrix(0, 2) = "  "
 grid1.TextMatrix(0, 3) = "  Ventas"
 grid1.TextMatrix(0, 4) = "    "
 grid1.TextMatrix(0, 5) = " Total"
 
 grid1.TextMatrix(1, 0) = ""
 grid1.TextMatrix(1, 1) = " al Contado"
 grid1.TextMatrix(1, 2) = "  ( % )"
 grid1.TextMatrix(1, 3) = " al Credito"
 grid1.TextMatrix(1, 4) = "  ( % )"
 grid1.TextMatrix(1, 5) = ""
 If Left(moneda.Text, 1) = "S" Then
   grid1.TextMatrix(2, 0) = "Totales  S/.-  :"
 Else
   grid1.TextMatrix(2, 0) = "Totales US$.-  :"
 End If
 grid1.RowHeight(2) = 350
End Sub

Public Sub ACTUALIZA_GRID()
cabe


End Sub

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
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If Not IsDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If

Exit Function
CANCELA:
 fechas = False
End Function

Private Sub moneda_Click()
pub_cadena = "SELECT SUM(FAR_PRECIO*(FAR_CANTIDAD/FAR_EQUIV))AS TOTAL FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND  FAR_CODVEN = ? AND FAR_DESCTO = 0 AND  FAR_ESTADO <> 'E' AND FAR_TIPMOV = 10 and FAR_DIAS = 0  AND FAR_MONEDA = '" & Left(moneda.Text, 1) & "'" ' AND (FAR_SIGNO_CAR = 1 OR FAR_SIGNO_CAR = 0) "
Set PS_CON01 = CN.CreateQuery("", pub_cadena)
PS_CON01(0) = 0
PS_CON01(1) = LK_FECHA_DIA
PS_CON01(2) = LK_FECHA_DIA
PS_CON01(3) = 0
Set llave_con01 = PS_CON01.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)

pub_cadena = "SELECT SUM(FAR_PRECIO*(FAR_CANTIDAD/FAR_EQUIV))AS TOTAL FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND FAR_CODVEN = ? AND FAR_DESCTO = 0 AND  FAR_ESTADO <> 'E' AND FAR_TIPMOV = 10 AND FAR_SIGNO_CAR = 1 AND FAR_DIAS <> 0 AND FAR_MONEDA = '" & Left(moneda.Text, 1) & "'"
Set PS_CON02 = CN.CreateQuery("", pub_cadena)
PS_CON02(0) = 0
PS_CON02(1) = 0
PS_CON02(2) = 0
PS_CON02(3) = 0
Set llave_con02 = PS_CON02.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)

End Sub

Private Sub opcion_Click(index As Integer)
If index = 0 Then
    VScroll1.Visible = False
    txts.Visible = False
    cmdcom.Visible = False
    Command1.Visible = True
ElseIf index = 1 Then
    VScroll1.Visible = True
    txts.Visible = True
    txts.Text = "05"
    cmdcom.Visible = True
    Command1.Visible = False
End If
End Sub

Private Sub tiempo_Timer()
lblTiempo.Caption = Format(val(lblTiempo.Caption) - 1, "00")
If val(lblTiempo.Caption) = 0 Then
  lblTiempo.Caption = txts.Text
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
