VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "Comctl32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form FrmTranf 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Transferencias a Procesos"
   ClientHeight    =   6555
   ClientLeft      =   1500
   ClientTop       =   1140
   ClientWidth     =   11130
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "FRMTRANF.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6555
   ScaleWidth      =   11130
   Tag             =   "55"
   WindowState     =   2  'Maximized
   Begin VB.Frame f1 
      Caption         =   "Datos:"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   10815
      Begin VB.TextBox txtcli 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   240
         TabIndex        =   17
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox tserie 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8040
         TabIndex        =   16
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtdoc 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8640
         TabIndex        =   15
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lcodart 
         Caption         =   "Nombre / Razon Social:"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   21
         Tag             =   "9999"
         Top             =   240
         Width           =   2925
      End
      Begin VB.Label lblcli 
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1800
         TabIndex        =   20
         Top             =   480
         Width           =   4695
      End
      Begin VB.Label lcodart 
         Caption         =   "Nº. Doc O/P"
         Height          =   255
         Index           =   0
         Left            =   8760
         TabIndex        =   19
         Tag             =   "9999"
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label lcodart 
         Caption         =   "Serie"
         Height          =   255
         Index           =   3
         Left            =   8040
         TabIndex        =   18
         Tag             =   "9999"
         Top             =   360
         Width           =   525
      End
   End
   Begin VB.Frame ESTADO 
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   5535
      Left            =   4680
      TabIndex        =   1
      Tag             =   "100"
      Top             =   960
      Width           =   6015
      Begin MSFlexGridLib.MSFlexGrid grid_fac 
         Height          =   4815
         Left            =   120
         TabIndex        =   0
         Tag             =   "9999"
         Top             =   360
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   8493
         _Version        =   327680
         Rows            =   3
         FixedRows       =   2
         FocusRect       =   2
         HighLight       =   2
         GridLines       =   2
         AllowUserResizing=   3
      End
      Begin VB.Label Label3 
         Caption         =   "Lista de Articulos :"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Tag             =   "9999"
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label momen 
         Caption         =   "Un Momento ..."
         Height          =   375
         Left            =   3360
         TabIndex        =   4
         Top             =   1560
         Width           =   1575
      End
   End
   Begin VB.Frame fraped 
      Height          =   2415
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   4575
      Begin MSFlexGridLib.MSFlexGrid gridped 
         Height          =   2055
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   3625
         _Version        =   327680
      End
   End
   Begin VB.Frame frapt 
      Height          =   3615
      Left            =   120
      TabIndex        =   12
      Top             =   3360
      Width           =   4575
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   3255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   5741
         _Version        =   327680
      End
   End
   Begin VB.CommandButton cmdconsulta 
      Caption         =   "Co&nsultas"
      Height          =   375
      Left            =   6240
      Picture         =   "FRMTRANF.frx":0442
      TabIndex        =   9
      Top             =   6600
      Width           =   1575
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   495
      Left            =   7200
      TabIndex        =   8
      Top             =   6840
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
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
   Begin VB.CommandButton cmdIngreso 
      Caption         =   "&Ingreso"
      Height          =   375
      Left            =   4920
      Picture         =   "FRMTRANF.frx":058C
      TabIndex        =   7
      Top             =   6600
      Width           =   1695
   End
   Begin VB.CommandButton SALIR 
      Caption         =   "Ce&rrar"
      Height          =   375
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   6600
      Width           =   1095
   End
   Begin VB.CommandButton cancelar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   5
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   6600
      Width           =   1095
   End
   Begin ComctlLib.ProgressBar Barra 
      Height          =   315
      Left            =   3360
      TabIndex        =   3
      Tag             =   "0"
      Top             =   6840
      Visible         =   0   'False
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   556
      _Version        =   327682
      Appearance      =   0
      Min             =   77
      Max             =   91
   End
   Begin Crystal.CrystalReport Reportes 
      Left            =   360
      Top             =   7080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
End
Attribute VB_Name = "FrmTranf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim VAR_ACTIVAR As Integer
Dim WCOD_ORIGINAL As Currency
Dim WPASA As Boolean
Dim WSELE As String * 1
Dim llave1
Dim loc_key
Dim fila As Integer
Dim ws_bruto_d, ws_bruto_h As Currency
Dim SUM_D As Currency
Dim SUM_H As Currency
Dim PSTEMP_LLAVE As rdoQuery
Dim temp_llave As rdoResultset
Dim WMODO As String * 1
Dim LOC_ITEM As Integer
Dim cop_llave As rdoResultset
Dim PSCOP_LLAVE As rdoQuery
Dim LOC_CANCELA As Integer
Dim PSTEMP_MAYOR As rdoQuery
Dim temp_mayor As rdoResultset
Dim temporal
Dim wfila_act As Integer
Dim loc_ini As Integer
Dim loc_fin  As Integer
Dim Wsec As Integer
Dim PSLOC_WARTI As rdoQuery
Dim llave_sum_arti   As rdoResultset
Dim PRE_ETIQUETA(5) As String * 20
Dim LOC_TIPMOV As Integer

Option Explicit

Private Sub cancelar_Click()

WMODO = ""
cmdIngreso.Caption = "&Ingreso"
f1.Enabled = False
ESTADO.Enabled = False
'PB.Visible = False
fila = 0
SUM_D = 0
SUM_H = 0
LIMPIA_DATOS
CABE_MAN
cmdIngreso.Enabled = True
'grid_fac.SetFocus

End Sub

Private Sub cmdconsulta_Click()

cmdIngreso.Enabled = False
tserie.Locked = False
txtdoc.Locked = False
tserie.Enabled = True
txtdoc.Enabled = True


f1.Enabled = True
tserie.Text = "100"
tserie.Locked = True
Azul txtdoc, txtdoc

End Sub


Private Sub cmdIngreso_Click()
Dim ws_tot_debe, ws_tot_haber As Currency
Dim er As rdoError
Dim pub_mensaje As String
Const ingre = 2
Const modif = 1
Dim N As Integer
Dim LOC_SALDO_CAR As Currency
Dim FLAG As Boolean
Dim pub_mensaje_err As String
Dim WS_NRO_MOV, ws_nro_voucher As Long
Dim w_dh  As String

If Left(cmdIngreso.Caption, 2) = "&G" Then
If Trim(txtcli.Text) = "" Then
  MsgBox "Nombre del Cliente ", 48, Pub_Titulo
  txtcli.SetFocus
  Exit Sub
End If
If Val(txttotal.Text) <= 0 Then
  MsgBox "Ingrese Datos ", 48, Pub_Titulo
  grid_fac.SetFocus
  Exit Sub
End If

If grid_fac.Rows = 3 Then
 If grid_fac.TextMatrix(2, 0) = "" Then
   MsgBox "Ingrese Datos de Productos ", 48, Pub_Titulo
   grid_fac.SetFocus
   Exit Sub
 End If
End If

Barra.Visible = False
suma_grid
For fila = 2 To grid_fac.Rows - 1
 If grid_fac.TextMatrix(fila, 1) <> "" Then
  If Val(grid_fac.TextMatrix(fila, 2)) <= 0 Then
    MsgBox "Verificar, cantidad en cero o menor. - " & grid_fac.TextMatrix(fila, 1) & " : " & grid_fac.TextMatrix(fila, 0), 48, Pub_Titulo
    grid_fac.SetFocus
    GoTo fin
  End If
  If Val(grid_fac.TextMatrix(fila, 4)) = 0 Then
    MsgBox "Verificar hay algun precio en 0 .", 48, Pub_Titulo
    grid_fac.SetFocus
    GoTo fin
  End If
End If
Next fila
Screen.MousePointer = 11
DoEvents
Barra.Visible = True
DoEvents
Barra.Min = 0
Barra.Max = fila
Barra.Value = 0
exito = True
Barra.Value = 1
GoSub ACT1
Call REP_CONSUL
fila = 1
SUM_D = 0
SUM_H = 0
CABE_MAN
LIMPIA_DATOS
fila = 0
'cancelar.SetFocus
CABE_MAN
Barra.Visible = False
cmdIngreso.Caption = "&Ingreso"

GoTo fin

ACT1:

fila = 1
FLAG = False
WS_NRO_MOV = 0
fila = 2
Do While FLAG = False
   If Trim(grid_fac.TextMatrix(fila, 1)) = "" Then GoTo pasa
    ' grabo todo
   temp_llave.AddNew
   temp_llave!PED_CODCIA = LK_CODCIA
   temp_llave!PED_FECHA = LK_FECHA_DIA
   temp_llave!PED_NUMSER = Trim(tserie.Text)
   temp_llave!PED_NUMFAC = Val(txtdoc.Text)
   temp_llave!PED_NUMSEC = WS_NRO_MOV
   temp_llave!PED_CANTIDAD = Val(grid_fac.TextMatrix(fila, 2))
   temp_llave!PED_PRECIO = Val(grid_fac.TextMatrix(fila, 4))
   temp_llave!PED_CODUSU = LK_CODUSU
   temp_llave!PED_IGV = Val(txtigv.Text)
   temp_llave!PED_BRUTO = Val(txtvalorv.Text)
   temp_llave!PED_ESTADO = "N"
   temp_llave!PED_CODUSU = LK_CODUSU
   temp_llave!PED_CODART = Val(grid_fac.TextMatrix(fila, 10))
   temp_llave!PED_UNIDAD = Trim(grid_fac.TextMatrix(fila, 3))
   temp_llave!PED_EQUIV = Val(grid_fac.TextMatrix(fila, 12))
   temp_llave!PED_NOMCLIE = Trim(FrmTranf.lblcli.Caption)
   'temp_llave!PED_RUCCLIE = Trim(txtruc.Text) ' Trim(fbg.Text)
   temp_llave!PED_CODCLIE = Val(txtcli.Text)
   temp_llave!PED_TIPMOV = 201
   temp_llave!PED_HORA = Format(Now, "hh:mm:ss AMPM")
   temp_llave!PED_DESCTO = Val(grid_fac.TextMatrix(fila, 5))
   'temp_llave!PED_MONEDA = Left(Trim(moneda.Text), 1)
   'temp_llave!PED_CONTACTO = txtatte.Text
   'temp_llave!PED_FORMA = Trim(forma.Text)
   'temp_llave!PED_TIEMPO = Trim(tiempo.Text)
   'temp_llave!PED_OFERTA = Trim(oferta.Text)
   temp_llave!PED_SUBTOTAL = Val(grid_fac.TextMatrix(fila, 6))
   temp_llave.Update
pasa:
   fila = fila + 1
   WS_NRO_MOV = WS_NRO_MOV + 1
   If fila >= FrmTranf.grid_fac.Rows Then
      FLAG = True
   End If
  
Loop

Return
Screen.MousePointer = 1
Exit Sub
End If
' cuando pulsa Ingreso
Dim wser As String
Dim wnumfac As String

cmdIngreso.Caption = "&Grabar / Imprimir"
f1.Enabled = True
ESTADO.Enabled = True
LIMPIA_DATOS
CABE_MAN
WMODO = "I"
PSTEMP_MAYOR(0) = LK_CODCIA
temp_mayor.Requery
If temp_mayor.EOF Then
 wser = 100
 wnumfac = 1
Else
 wser = Nulo_Valors(temp_mayor!PED_NUMSER)
 wnumfac = Val(Nulo_Valor0(temp_mayor!PED_NUMFAC)) + 1
End If
tserie.Text = wser
txtdoc.Text = wnumfac

grid_fac.Rows = grid_fac.Rows + 1
grid_fac.RowHeight(grid_fac.Rows - 1) = 285
grid_fac.TextMatrix(grid_fac.Rows - 1, 0) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 1) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 2) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 3) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 4) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 5) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 6) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 7) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 8) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 9) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 10) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 11) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 12) = ""

grid_fac.TextMatrix(grid_fac.Rows - 1, 14) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 15) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 16) = ""
grid_fac.TextMatrix(grid_fac.Rows - 1, 17) = ""
txtcli.SetFocus
Exit Sub

Error_fatal:
    pub_mensaje = "Se ha producido un error " & "al abrir la conexión:" & Err & " - " & Error & vbCr
    For Each er In rdoErrors
        pub_mensaje = pub_mensaje & er.Description & ":" & er.Number & vbCr
        MsgBox pub_mensaje
    Next er
    CN.Execute "Rollback Transaction", rdExecDirect
'    Resume AbandonCn
Exit Sub

errorr:
 MsgBox pub_mensaje_err, 48, Pub_Titulo
fin:
Screen.MousePointer = 0
Exit Sub
SALE:
If Err.Number = 6 Then
  MsgBox "Verficar Importe.", 48, Pub_Titulo
  If TEXTOVAR.Visible Then Azul3 TEXTOVAR, TEXTOVAR
  FrmTranf.Barra.Visible = False
  Screen.MousePointer = 0
  grid_fac.SetFocus
Else
  MsgBox Err.Description, 48, Pub_Titulo
End If

End Sub


Private Sub Form_Load()
'On Error GoTo SALE
Wsec = 0
LOC_CANCELA = 0
fila = 0
wfila_act = 0
WSELE = ""
Dim ws_indice As Integer
Dim cade
WMODO = ""
Dim PSPRO_V As rdoQuery
Dim PRO_V As rdoResultset


pub_cadena = "SELECT * FROM PEDIDOS WHERE PED_CODCIA = ? AND PED_TIPMOV = ? AND PED_NUMSER = ? and PED_NUMFAC = ?  ORDER BY PED_NUMSEC"
Set PSLOC_WARTI = CN.CreateQuery("", pub_cadena)
PSLOC_WARTI(0) = 0
PSLOC_WARTI(1) = 0
PSLOC_WARTI(2) = 0
PSLOC_WARTI(3) = 0
Set llave_sum_arti = PSLOC_WARTI.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM PEDIDOS WHERE PED_CODCIA = ? AND PED_TIPMOV = 201  ORDER BY  PED_NUMFAC DESC "
Set PSTEMP_MAYOR = CN.CreateQuery("", pub_cadena)
PSTEMP_MAYOR(0) = LK_CODCIA
PSTEMP_MAYOR.MaxRows = 1
Set temp_mayor = PSTEMP_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM PEDIDOS WHERE  PED_TIPMOV = 201 ORDER BY PED_CODCIA"
Set PSTEMP_LLAVE = CN.CreateQuery("", pub_cadena)
Set temp_llave = PSTEMP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
temp_llave.Requery

fila = 0
DoEvents
LIMPIA_DATOS
CABE_MAN
SQ_OPER = 2
PUB_TIPREG = 45
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
fila = 0
Do Until tab_mayor.EOF
PRE_ETIQUETA(fila) = Trim(tab_mayor!tab_nomlargo)
fila = fila + 1
tab_mayor.MoveNext
Loop

Exit Sub
SALE:
MsgBox "Depurar: " & Err.Description, 48, Pub_Titulo
Resume Next
End Sub

Private Sub grid_fac_EnterCell()
TEXTOVAR.Visible = False
TEXTOVAR.Text = Trim(grid_fac.TextMatrix(grid_fac.Row, grid_fac.COL))
TEXTOVAR.Left = grid_fac.Left + grid_fac.CellLeft
TEXTOVAR.Width = grid_fac.CellWidth
TEXTOVAR.Height = grid_fac.CellHeight
TEXTOVAR.Top = ESTADO.Top + grid_fac.Top + grid_fac.CellTop - 1560 '480
If grid_fac.COL = 1 Then
 If Trim(grid_fac.TextMatrix(grid_fac.Row, 1)) = "" Then
  TEXTOVAR.Visible = True
  TEXTOVAR.SetFocus
 End If
End If
If Trim(grid_fac.TextMatrix(grid_fac.Row, 1)) <> "" Then
 stock.Caption = Format(Val(grid_fac.TextMatrix(grid_fac.Row, 15)) / Val(grid_fac.TextMatrix(grid_fac.Row, 12)), "0.00")
 unid.Caption = grid_fac.TextMatrix(grid_fac.Row, 16)
 nomarti.Caption = grid_fac.TextMatrix(grid_fac.Row, 0)
Else
 stock.Caption = ""
 unid.Caption = ""
 nomarti.Caption = ""
End If

End Sub

Private Sub grid_fac_KeyPress(KeyAscii As Integer)
Dim a As Integer
Dim t, WC
Dim wprecios As String * 12
Static CONS
Dim wactivo As Integer
If KeyAscii <> 13 Then Exit Sub
If grid_fac.Rows <= 1 Then Exit Sub
'If grid_fac.COL = 1 Then Exit Sub
If grid_fac.COL >= 6 Then Exit Sub

If grid_fac.COL = 2 Then
 If Trim(grid_fac.TextMatrix(grid_fac.Row, 1)) = "" Then
    grid_fac.SetFocus
    Exit Sub
 End If
End If
If grid_fac.COL = 3 Then
    If Trim(grid_fac.TextMatrix(grid_fac.Row, 1)) = "" Then
     grid_fac.SetFocus
     Exit Sub
    End If
'    UNIDAD.Left = grid_fac.Left + grid_fac.CellLeft
'    UNIDAD.Width = grid_fac.CellWidth
'    UNIDAD.Top = ESTADO.Top + grid_fac.Top + grid_fac.CellTop - 1550 '480
    SQ_OPER = 2
    pu_codcia = LK_CODCIA
    PUB_CODART = Val(grid_fac.TextMatrix(grid_fac.Row, 10))
    LEER_PRE_LLAVE
'    UNIDAD.Clear
'    UNIDAD.Visible = True
    wactivo = 0
    Do Until pre_mayor.EOF
     'UNIDAD.AddItem Trim(pre_mayor!pre_UNIDAD) & String(30, " ") & pre_mayor!pre_secuencia
     'If pre_mayor!PRE_FLAG_UNIDAD = "A" Then
     '  wactivo = pre_mayor.AbsolutePosition - 1
     'End If
     pre_mayor.MoveNext
    Loop
    On Error GoTo pasa
'    UNIDAD.ListIndex = wactivo
    grid_fac.TextMatrix(grid_fac.Row, 13) = wactivo
    On Error GoTo 0
'    UNIDAD.Visible = True
'    UNIDAD.SetFocus
    SendKeys "%{up}"
     Exit Sub
End If
If grid_fac.COL = 4 Then
    If Trim(grid_fac.TextMatrix(grid_fac.Row, 1)) = "" Then
      grid_fac.SetFocus
      Exit Sub
    End If
'    PRECIOS.Left = grid_fac.Left + grid_fac.CellLeft
'    PRECIOS.Width = grid_fac.CellWidth + 600
'    PRECIOS.Top = ESTADO.Top + grid_fac.Top + grid_fac.CellTop - 1550 '480
    SQ_OPER = 2
    pu_codcia = LK_CODCIA
    PUB_CODART = Val(grid_fac.TextMatrix(grid_fac.Row, 10))
    LEER_PRE_LLAVE
'    PRECIOS.Clear
    On Error GoTo pasa
    
    On Error GoTo 0
    
    SendKeys "%{up}"
     Exit Sub
End If
If grid_fac.COL = 5 Then
 If Trim(grid_fac.TextMatrix(grid_fac.Row, 1)) = "" Then
    grid_fac.SetFocus
    Exit Sub
 End If
End If


'    TEXTOVAR.Left = grid_fac.Left + grid_fac.CellLeft
'    TEXTOVAR.Width = grid_fac.CellWidth
'    TEXTOVAR.Height = grid_fac.CellHeight
'    TEXTOVAR.Top = ESTADO.Top + grid_fac.Top + grid_fac.CellTop - 1200 '480
'    TEXTOVAR.Text = grid_fac.TextMatrix(grid_fac.Row, grid_fac.COL)
'    wfila_act = grid_fac.Row
Exit Sub
pasa:
Resume Next
End Sub

Private Sub grid_fac_KeyUp(KeyCode As Integer, Shift As Integer)
Dim WC
Dim a, WF As Integer
Dim tf, t, tC
Dim SALE As Boolean

If KeyCode = 46 Then
If grid_fac.Rows <= 2 Then Exit Sub
If grid_fac.Rows <= 3 Then
    pub_mensaje = MsgBox("Quitar el Producto para la Orden de Compra ", vbYesNo + vbExclamation + vbDefaultButton2, Pub_Titulo)
    If pub_mensaje = vbNo Then
      grid_fac.SetFocus
      Exit Sub
    End If
    CABE_MAN
Else
   pub_mensaje = MsgBox("Quitar el Producto para la Orden de Compra ", vbYesNo + vbExclamation + vbDefaultButton2, Pub_Titulo)
   If pub_mensaje = vbNo Then
     grid_fac.SetFocus
     Exit Sub
   Else
   '  grid_fac.RowHeight(grid_fac.Row) = 1
   grid_fac.RemoveItem (grid_fac.Row)
   grid_fac.Row = grid_fac.Row
   grid_fac.Refresh
   suma_grid
   grid_fac.SetFocus
   End If
End If
End If
'grid_fac.SetFocus
Exit Sub



End Sub



Private Sub ListView1_DblClick()
' loc_key = ListView1.SelectedItem.Index
' TEXTOVAR.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
' TEXTOVAR_KeyPress 13
End Sub

Private Sub ListView1_GotFocus()
'If loc_key <> 0 Then
' Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
' ListView1.ListItems.Item(loc_key).Selected = True
' ListView1.ListItems.Item(loc_key).EnsureVisible
'End If

End Sub

Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)
'If loc_key <> 0 Then
' loc_key = ListView1.SelectedItem.Index
' TEXTOVAR.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
'End If
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
'If KeyAscii = 27 Then
' ListView1.Visible = False
' TEXTOVAR.Text = ""
' TEXTOVAR.SetFocus
' Exit Sub
'End If
'If KeyAscii <> 13 Then
' Exit Sub
'End If
'ListView1_DblClick
End Sub

Private Sub ListView1_LostFocus()
ListView1.Visible = False
End Sub



Private Sub PRECIOS_GotFocus()
grid_fac.TextMatrix(grid_fac.Row, 13) = Format(Val(Mid(PRECIOS.Text, 10, Len(Trim(PRECIOS.Text)) - 10)), "0.00")
grid_fac.TextMatrix(grid_fac.Row, 14) = Val(Right(PRECIOS.Text, 3))
End Sub

Private Sub PRECIOS_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 PRECIOS.Visible = False
 grid_fac.SetFocus
End If
If KeyAscii <> 13 Then Exit Sub
'SQ_OPER = 1
'pu_codcia = LK_CODCIA
'PUB_CODART = Val(grid_fac.TextMatrix(grid_fac.Row, 10))
'PUB_SECUEN = Val(Right(UNIDAD.Text, 4))
'LEER_PRE_LLAVE

grid_fac.TextMatrix(grid_fac.Row, 4) = Format(Val(Mid(PRECIOS.Text, 10, Len(Trim(PRECIOS.Text)) - 10)), "0.00")
grid_fac.TextMatrix(grid_fac.Row, 13) = Format(Val(Mid(PRECIOS.Text, 10, Len(Trim(PRECIOS.Text)) - 10)), "0.00")
grid_fac.TextMatrix(grid_fac.Row, 14) = Val(Right(PRECIOS.Text, 3))

PRECIOS.Visible = False
suma_grid
grid_fac.COL = 5
grid_fac_KeyPress 13


End Sub

Private Sub PRECIOS_KeyUp(KeyCode As Integer, Shift As Integer)
Dim ww As String
Dim wpre As Currency
If KeyCode = 45 Then
'seIf grid_fac.COL = 6 Then
ww = InputBox("Digite Precios :", "Ingreso de Precio", "0")
wpre = Val(ww)
grid_fac.TextMatrix(grid_fac.Row, 4) = wpre
If Val(grid_fac.TextMatrix(grid_fac.Row, 13)) <> 0 Then
  grid_fac.TextMatrix(grid_fac.Row, 5) = redondea((Val(grid_fac.TextMatrix(grid_fac.Row, 13)) - wpre) * 100 / Val(Val(grid_fac.TextMatrix(grid_fac.Row, 13))))
End If
PRECIOS.Visible = False
suma_grid
'grid_fac.Visible = True
grid_fac.SetFocus


End If
End Sub

Private Sub salir_Click()
Unload FrmTranf
End Sub


Public Sub LIMPIA_DATOS()
grid_fac.Enabled = True
lblcli.Caption = ""

f1.Enabled = True
txtcli.Text = ""


tserie.Text = ""
txtdoc.Text = ""
grid_fac.Clear

txtigv.Text = ""
txtvalorv.Text = ""
txttotal.Text = ""
TEXTOVAR.Visible = False
stock.Caption = ""
unid.Caption = ""
nomarti.Caption = ""




End Sub

Public Sub CABE_MAN()
grid_fac.Cols = 18
grid_fac.Rows = 2
grid_fac.Clear
fila = 0
grid_fac.ColWidth(0) = 2400 ' nombre arti
grid_fac.ColWidth(1) = 900 ' codigo arti
grid_fac.ColWidth(2) = 800 ' cantidad
grid_fac.ColWidth(3) = 900 ' unidad
grid_fac.ColWidth(4) = 800 ' precio
grid_fac.ColWidth(5) = 700 ' decto.
grid_fac.ColWidth(6) = 1000 ' sub total
grid_fac.ColWidth(7) = 800 ' peso
grid_fac.ColWidth(8) = 0
grid_fac.ColWidth(9) = 0
grid_fac.ColWidth(10) = 0 '  COD ORIGINAL
grid_fac.ColWidth(11) = 0 '  PRE_SECUENCIA
grid_fac.ColWidth(12) = 0 '  PRE_EQUIV
grid_fac.ColWidth(13) = 0 '  PRE_PRECIO COLOCADO
grid_fac.ColWidth(14) = 0 '  numero de PRE_PRECIO
grid_fac.ColWidth(15) = 0 '  numero de arm_stock
grid_fac.ColWidth(16) = 0 '  numero de pre_unidad
grid_fac.ColWidth(17) = 0

grid_fac.TextMatrix(0, 0) = "Articulo"
grid_fac.TextMatrix(0, 1) = "Codigo"
grid_fac.TextMatrix(0, 2) = "Cantidad"
grid_fac.TextMatrix(0, 3) = "Unidad"
grid_fac.TextMatrix(0, 4) = "Precios"
grid_fac.TextMatrix(0, 5) = "Dscto"
grid_fac.TextMatrix(1, 5) = "  (%)"
grid_fac.TextMatrix(0, 6) = "Sub Total"
grid_fac.TextMatrix(0, 7) = "Peso(Kg)"
grid_fac.TextMatrix(0, 8) = ""
grid_fac.TextMatrix(0, 9) = ""
grid_fac.RowHeight(1) = 320

End Sub
Public Sub suma_grid()
On Error GoTo SALE
Dim WF As Integer
WF = 2
Dim fx As Integer
Dim wcantid As Currency
Dim wpeso As Currency
fx = 1
SUM_H = 0
SUM_D = 0
wcantid = 0
Do While fx = 1
    'If Left(grid_fac.TextMatrix(WF, 0), 1) <> "T" Then
      SUM_D = SUM_D + Val(grid_fac.TextMatrix(WF, 4))
      SUM_H = SUM_H + Val(Val(grid_fac.TextMatrix(WF, 2)) * Val(grid_fac.TextMatrix(WF, 4)))
      wcantid = wcantid + Val(grid_fac.TextMatrix(WF, 2))
      wpeso = wpeso + Val(grid_fac.TextMatrix(WF, 7))
      grid_fac.TextMatrix(WF, 6) = Format(Val(grid_fac.TextMatrix(WF, 2)) * Val(grid_fac.TextMatrix(WF, 4)), "0.00")
    'End If
    WF = WF + 1
    If WF = grid_fac.Rows Then
        fx = 0
    Else
        If Trim(grid_fac.TextMatrix(WF, 0)) = "" Then fx = 0
    End If
Loop
   fila = WF - 1
   grid_fac.TextMatrix(1, 0) = "Totales = "
   grid_fac.TextMatrix(1, 6) = Format(SUM_H, "####0.00")
   grid_fac.TextMatrix(1, 2) = Format(wcantid, "####0.00")
   grid_fac.TextMatrix(1, 7) = Format(wpeso, "####0.00")
   
   txttotal.Text = Format(SUM_H, "#####0.00")
   txtvalorv.Text = Format((SUM_H / ((100 + LK_IGV) / 100)), "####0.00")
   txtigv.Text = Format(SUM_H - Val(txtvalorv.Text), "#####0.00")
  
Exit Sub
SALE:
cancelar_Click
'MsgBox "Verficar Importe.", 48, Pub_Titulo
'Resume Next
'If TEXTOVAR.Visible Then Azul3 TEXTOVAR, TEXTOVAR
End Sub
Public Sub suma_subtotal()
If WMODO = "I" Then Exit Sub

Dim WF As Integer
Dim WFIN As Integer
Dim WINI As Integer

Dim fx As Integer
Exit Sub
End Sub

Private Sub Consistencias(wsGrid As MSFlexGrid, wsTexto As RichTextBox, wsKeyAscii As Integer)
  Static valor
  Dim car As String
 ' NUMEROS CON DECIMALES
    car = Chr$(wsKeyAscii)
    car = UCase$(Chr$(wsKeyAscii))
    wsKeyAscii = Asc(car)
    If wsKeyAscii = 45 Then
      If wsTexto.Text <> "" Then
         Beep
         wsKeyAscii = 0
         Exit Sub
      End If
    End If
    If wsKeyAscii = 46 Then
      If InStr(1, wsTexto.Text, ".") <> 0 Then
        Beep
        wsKeyAscii = 0
        Exit Sub
      End If
    End If
    
    If car < "0" Or car > "9" Then
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And car <> "." Then
          wsKeyAscii = 0
          Beep
          Exit Sub
        End If
    End If

End Sub

Public Sub CABE_ING()
grid_fac.Cols = 6
grid_fac.Rows = 3
grid_fac.Clear
grid_fac.MergeCells = 4
grid_fac.MergeCol(0) = True
grid_fac.MergeCol(1) = True
grid_fac.MergeCol(2) = True
grid_fac.MergeCol(3) = True
grid_fac.MergeCol(4) = False
grid_fac.MergeCol(5) = False
grid_fac.MergeRow(2) = False
grid_fac.RowHeight(0) = 285
grid_fac.RowHeight(1) = 285
grid_fac.RowHeight(2) = 285

fila = 0
grid_fac.ColWidth(0) = 400
grid_fac.ColWidth(1) = 1400
grid_fac.ColWidth(2) = 2500
grid_fac.ColWidth(3) = 0
grid_fac.ColWidth(4) = 1500
grid_fac.ColWidth(5) = 1500

grid_fac.TextMatrix(0, 0) = "Item"
grid_fac.TextMatrix(0, 1) = "Cuenta"
grid_fac.TextMatrix(0, 2) = "Descripcion"
grid_fac.TextMatrix(0, 3) = "Glosa"
grid_fac.TextMatrix(0, 4) = "Debe"
grid_fac.TextMatrix(0, 5) = "Haber"
grid_fac.TextMatrix(1, 0) = "Item"
grid_fac.TextMatrix(1, 1) = "Cuenta"
grid_fac.TextMatrix(1, 2) = "Descripcion"
grid_fac.TextMatrix(1, 3) = "Glosa"

'grid_fac.MergeCol
'grid_fac.MergeRow(2) = True



End Sub

Private Sub textovar_Change()
If Not TEXTOVAR.Visible Then Exit Sub
If grid_fac.COL = 5 Then
 grid_fac.TextMatrix(grid_fac.Row, 4) = redondea(Val(grid_fac.TextMatrix(grid_fac.Row, 13)) * (100 - Val(TEXTOVAR.Text)) / 100)
End If
If grid_fac.COL = 1 Then
    grid_fac.TextMatrix(grid_fac.Row, 0) = ""
    grid_fac.TextMatrix(grid_fac.Row, 0) = ""
    grid_fac.TextMatrix(grid_fac.Row, 1) = ""
    grid_fac.TextMatrix(grid_fac.Row, 2) = ""
    grid_fac.TextMatrix(grid_fac.Row, 3) = ""
    grid_fac.TextMatrix(grid_fac.Row, 4) = ""
    grid_fac.TextMatrix(grid_fac.Row, 5) = ""
    grid_fac.TextMatrix(grid_fac.Row, 6) = ""
    grid_fac.TextMatrix(grid_fac.Row, 7) = ""
    grid_fac.TextMatrix(grid_fac.Row, 8) = ""
    grid_fac.TextMatrix(grid_fac.Row, 9) = ""
    grid_fac.TextMatrix(grid_fac.Row, 10) = ""
    grid_fac.TextMatrix(grid_fac.Row, 11) = ""
    grid_fac.TextMatrix(grid_fac.Row, 12) = ""
    grid_fac.TextMatrix(grid_fac.Row, 14) = ""
    grid_fac.TextMatrix(grid_fac.Row, 15) = ""
    grid_fac.TextMatrix(grid_fac.Row, 16) = ""
    grid_fac.TextMatrix(grid_fac.Row, 17) = ""
    grid_fac.Text = TEXTOVAR.Text
    stock.Caption = ""
    unid.Caption = ""
    nomarti.Caption = ""
    suma_grid
Else
 If grid_fac.COL = 2 Then
  grid_fac.Text = TEXTOVAR.Text
 Else
  grid_fac.Text = Format(TEXTOVAR.Text, "0.00")
 End If
 suma_grid
 suma_subtotal
End If
End Sub

Private Sub TEXTOVAR_GotFocus()
'temporal = grid_fac.TextMatrix(grid_fac.Row, grid_fac.COL)
End Sub

Private Sub textovar_KeyDown(KeyCode As Integer, Shift As Integer)

' busca arti
If Not ListView1.Visible Then
If KeyCode = 40 Then  ' flecha abajo
  If grid_fac.Row = grid_fac.Rows - 1 Then Exit Sub
  If Trim(grid_fac.Text) <> "" Then Exit Sub
  grid_fac.Row = grid_fac.Row + 1
  grid_fac.SetFocus
  Exit Sub
End If
If KeyCode = 38 Then
 If Trim(grid_fac.Text) <> "" Then Exit Sub
 grid_fac.Row = grid_fac.Row - 1
 grid_fac.SetFocus
 Exit Sub
End If
If KeyCode = 39 Then
If Trim(grid_fac.Text) <> "" Then Exit Sub
 grid_fac.COL = grid_fac.COL + 1
 grid_fac.SetFocus
 Exit Sub
End If
End If
If grid_fac.COL <> 1 Then Exit Sub
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If Not ListView1.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And TEXTOVAR.Text = "" Then
  loc_key = 1
  Set ListView1.SelectedItem = ListView1.ListItems(loc_key)
'  LISTVIEW1.Visible = False
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
  TEXTOVAR.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
  DoEvents
  TEXTOVAR.SelStart = Len(TEXTOVAR.Text)
  DoEvents
fin:

End Sub

Private Sub textovar_KeyPress(KeyAscii As Integer)
'SOLO_DECIMAL TEXTOVAR, KeyAscii
If KeyAscii = 27 Then
  If TEXTOVAR.Text = "" Then
    TEXTOVAR.Visible = False
    grid_fac.SetFocus
    Exit Sub
  End If
  TEXTOVAR.Text = "" ' temporal
  'TEXTOVAR.Visible = False
  'grid_fac.SetFocus
  ListView1.Visible = False
  Exit Sub
End If
If grid_fac.COL = 2 Or grid_fac.COL = 4 Or grid_fac.COL = 5 Then Consistencias grid_fac, TEXTOVAR, KeyAscii
If KeyAscii <> 13 Then Exit Sub

If grid_fac.COL = 2 Then
 TEXTOVAR.Visible = False
 If Val(arm_llave!ARM_STOCK) - Val(grid_fac.TextMatrix(grid_fac.Row, 2)) <= 0 Then
      MsgBox "Stock es :" & Format(arm_llave!ARM_STOCK, "0.00") & "  /  Aplicando la cantidad : " & Format(Val(arm_llave!ARM_STOCK) - Val(grid_fac.TextMatrix(grid_fac.Row, 2)), "0.00"), 48, Pub_Titulo
 End If
 grid_fac.COL = 3
 If Trim(grid_fac.Text) <> "" Then
   grid_fac.SetFocus
   Exit Sub
 End If
 grid_fac_KeyPress 13
 Exit Sub
End If
If grid_fac.COL = 6 Then
' grid_fac.TextMatrix(grid_fac.Row, 6) = textovar.Text
' suma_grid
' textovar.Visible = False
' grid_fac.SetFocus
' Exit Sub
End If
If grid_fac.COL = 5 Then
 TEXTOVAR.Visible = False
 If Trim(grid_fac.TextMatrix(grid_fac.Rows - 1, 1)) <> "" Then
   grid_fac.Rows = grid_fac.Rows + 1
   grid_fac.RowHeight(grid_fac.Rows - 1) = 285
   grid_fac.Row = grid_fac.Rows - 1
 Else
  If grid_fac.Row < grid_fac.Rows - 1 Then
     grid_fac.Row = grid_fac.Row + 1
  End If
 End If
 grid_fac.COL = 1
 TEXTOVAR.Visible = True
 TEXTOVAR.SetFocus
 Exit Sub
End If




If grid_fac.COL <> 1 Then Exit Sub

Dim valor As String
Dim tf As Integer
Dim i, car
Dim itmFound As ListItem
car = Chr(KeyAscii)
KeyAscii = Asc(UCase(car))
If KeyAscii = 27 Then
 ListView1.Visible = False
 TEXTOVAR.Text = ""
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
VAR_ACTIVAR = 0
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  PUB_KEY = 0
Else
 On Error GoTo mucho
 PUB_KEY = Val(TEXTOVAR.Text)
 On Error GoTo 0
 If Len(TEXTOVAR.Text) = 0 Then
    Exit Sub
 End If
 If IsNumeric(TEXTOVAR.Text) = False Then
   PUB_KEY = 0
 End If
End If

If PUB_KEY <> 0 Then
    SQ_OPER = 1
    PUB_KEY = TEXTOVAR.Text
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    If art_LLAVE.EOF Then
       MsgBox "Codigo NO Existe.", 48, Pub_Titulo
       Azul3 TEXTOVAR, TEXTOVAR
       GoTo fin
    End If
    If art_LLAVE!art_flag_stock <> "M" Then
       MsgBox "Producto no es Mercaderia.", 48, Pub_Titulo
       Azul3 TEXTOVAR, TEXTOVAR
       GoTo fin
    End If
    WCOD_ORIGINAL = art_LLAVE!ART_KEY
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    PUB_CODART = WCOD_ORIGINAL
    LEER_ARM_LLAVE
    SQ_OPER = 1
    pu_codcia = LK_CODCIA
    PUB_SECUEN = 0
    LEER_PRE_LLAVE
    grid_fac.TextMatrix(grid_fac.Row, 16) = pre_llave!pre_UNIDAD
    grid_fac.TextMatrix(grid_fac.Row, 15) = arm_llave!ART_stock
    grid_fac.TextMatrix(grid_fac.Row, 12) = pre_llave!pre_equiv
    grid_fac.TextMatrix(grid_fac.Row, 11) = pre_llave!pre_secuencia
    grid_fac.TextMatrix(grid_fac.Row, 0) = art_LLAVE!ART_NOMBRE
    grid_fac.TextMatrix(grid_fac.Row, 10) = art_LLAVE!ART_KEY
    ListView1.Visible = False
    TEXTOVAR.Visible = False
    grid_fac.COL = 2
    If Trim(grid_fac.Text) <> "" Then
      grid_fac.SetFocus
      Exit Sub
    End If
    TEXTOVAR.Visible = True
    TEXTOVAR.SetFocus
    Exit Sub
Else
  If ListView1.Visible = False And VAR_ACTIVAR <> 99 And TEXTOVAR.Text <> "" And LK_FLAG_ORIGINAL <> "A" And LK_FLAG_ALTERNO = "A" Then
IR_ALTERNO:
     SQ_OPER = 3
     pu_alterno = TEXTOVAR.Text
     pu_codcia = LK_CODCIA
     LEER_ART_LLAVE
     If art_llave_alt.EOF Then
       MsgBox "Codigo No Existe ...", 48, Pub_Titulo
       Azul3 TEXTOVAR, TEXTOVAR
       Exit Sub
     End If
     If art_llave_alt!art_flag_stock <> "M" Then
       MsgBox "Producto no es Mercaderia.", 48, Pub_Titulo
       Azul3 TEXTOVAR, TEXTOVAR
       GoTo fin
     End If
     ListView1.Visible = False
     WCOD_ORIGINAL = art_llave_alt!ART_KEY
     SQ_OPER = 1
     pu_codcia = LK_CODCIA
     PUB_CODART = WCOD_ORIGINAL
     LEER_ARM_LLAVE
     SQ_OPER = 1
     pu_codcia = LK_CODCIA
     PUB_SECUEN = 0
     LEER_PRE_LLAVE
     grid_fac.TextMatrix(grid_fac.Row, 16) = pre_llave!pre_UNIDAD
     grid_fac.TextMatrix(grid_fac.Row, 15) = arm_llave!ARM_STOCK
     grid_fac.TextMatrix(grid_fac.Row, 12) = pre_llave!pre_equiv
     grid_fac.TextMatrix(grid_fac.Row, 11) = pre_llave!pre_secuencia
    
     grid_fac.TextMatrix(grid_fac.Row, 0) = art_llave_alt!ART_NOMBRE
     grid_fac.TextMatrix(grid_fac.Row, 10) = art_llave_alt!ART_KEY
     TEXTOVAR.Visible = False
     ListView1.Visible = False
     grid_fac.COL = 2
     If Trim(grid_fac.Text) <> "" Then
       grid_fac.SetFocus
       Exit Sub
     End If
     TEXTOVAR.Visible = True
     Azul3 TEXTOVAR, TEXTOVAR
     Exit Sub
  Else
    If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
     Exit Sub
    End If
    valor = UCase(ListView1.ListItems.Item(loc_key).Text)
    If Trim(UCase(TEXTOVAR.Text)) = Left(valor, Len(Trim(TEXTOVAR.Text))) And Len(Trim(TEXTOVAR.Text)) <> 0 Then
      If VAR_ACTIVAR = 0 And LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
        TEXTOVAR.Text = Trim(ListView1.ListItems.Item(loc_key))
        GoTo IR_ALTERNO
      End If
      If VAR_ACTIVAR <> 99 Then
       TEXTOVAR.Text = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
      Else
       TEXTOVAR.Text = Trim(ListView1.ListItems.Item(loc_key))
      End If
      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
       PUB_KEY = Val(ListView1.ListItems.Item(loc_key).SubItems(1))
      Else
       PUB_KEY = TEXTOVAR.Text
      End If
      LEER_ART_LLAVE
      VAR_ACTIVAR = 0
      If art_LLAVE.EOF Then
        MsgBox "Codigo No Existe ...", 48, Pub_Titulo
        Azul3 TEXTOVAR, TEXTOVAR
        Exit Sub
      End If
      If art_LLAVE!art_flag_stock <> "M" Then
       MsgBox "Producto no es Mercaderia.", 48, Pub_Titulo
       Azul3 TEXTOVAR, TEXTOVAR
       GoTo fin
      End If
      WCOD_ORIGINAL = art_LLAVE!ART_KEY
      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      PUB_CODART = WCOD_ORIGINAL
      LEER_ARM_LLAVE
      SQ_OPER = 1
      pu_codcia = LK_CODCIA
      PUB_SECUEN = 0
      LEER_PRE_LLAVE
      grid_fac.TextMatrix(grid_fac.Row, 16) = pre_llave!pre_UNIDAD
      grid_fac.TextMatrix(grid_fac.Row, 15) = arm_llave!ART_stock
      grid_fac.TextMatrix(grid_fac.Row, 12) = pre_llave!pre_equiv
      grid_fac.TextMatrix(grid_fac.Row, 11) = pre_llave!pre_secuencia
      ListView1.Visible = False
      grid_fac.TextMatrix(grid_fac.Row, 0) = art_LLAVE!ART_NOMBRE
      grid_fac.TextMatrix(grid_fac.Row, 10) = art_LLAVE!ART_KEY
      grid_fac.COL = 2
      If Trim(grid_fac.Text) <> "" Then
        grid_fac.SetFocus
        Exit Sub
      End If
      TEXTOVAR.Visible = True
      TEXTOVAR.SetFocus
     
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
Azul3 TEXTOVAR, TEXTOVAR
  

Exit Sub

End Sub

Private Sub textovar_KeyUp(KeyCode As Integer, Shift As Integer)
If grid_fac.COL <> 1 Then Exit Sub
' busca arti
Dim VAR
If KeyCode = 13 Then Exit Sub
If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
  If Len(TEXTOVAR.Text) = 0 Or Trim(TEXTOVAR.Text) = "" Then
    ListView1.Visible = False
    Exit Sub
  End If
  If TEXTOVAR.Text = "*" And KeyCode = 106 Then
   VAR_ACTIVAR = 99
   Exit Sub
  ElseIf TEXTOVAR.Text = "" Then
   VAR_ACTIVAR = 0
   Exit Sub
  End If
  If VAR_ACTIVAR <> 99 Then
    Exit Sub
  End If
  If Left(TEXTOVAR.Text, 1) = "*" Then
   TEXTOVAR.Text = Mid(TEXTOVAR.Text, 2, Len(TEXTOVAR.Text))
   TEXTOVAR.SelStart = Len(TEXTOVAR.Text)
  End If
Else
 If Len(TEXTOVAR.Text) = 0 Or IsNumeric(TEXTOVAR.Text) = True Then
   ListView1.Visible = False
   Exit Sub
 End If
End If
If ListView1.Visible = False And KeyCode <> 13 Or Len(TEXTOVAR.Text) = 1 Then
    VAR = Asc(TEXTOVAR.Text)
    VAR = VAR + 1
    If VAR = 33 Or VAR = 91 Then
       VAR = "ZZZZZZZZ"
    Else
       VAR = Chr(VAR)
    End If
    If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
      numarchi = 3
      archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO FROM ARTI WHERE  ART_CODCIA = '" & LK_CODCIA & "' AND ART_CALIDAD = 1 AND ART_FLAG_STOCK = 'M' AND ART_ALTERNO BETWEEN '" & TEXTOVAR.Text & "' AND  '" & VAR & "' ORDER BY ART_ALTERNO"
    Else
      numarchi = 0
      archi = "SELECT ART_KEY, ART_CODCIA, ART_NOMBRE, ART_ALTERNO FROM ARTI WHERE  ART_CODCIA = '" & LK_CODCIA & "' AND ART_CALIDAD = 1 AND ART_FLAG_STOCK = 'M' AND ART_NOMBRE BETWEEN '" & TEXTOVAR.Text & "' AND  '" & VAR & "' ORDER BY ART_NOMBRE"
    End If
   ' If Len(TEXTOVAR.text) > 1 And ListView1.ListItems.count = 0 Then
   ' Else
     PROC_LISVIEW ListView1
   ' End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As ListItem    ' Variable FoundItem.
If ListView1.Visible Then
  Set itmFound = ListView1.FindItem(LTrim(TEXTOVAR.Text), lvwText, , lvwPartial)
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

Private Sub textovar_LostFocus()
'TEXTOVAR.Visible = False
'If TEXTOVAR.Visible Then
'   TEXTOVAR.Visible = False
'   grid_fac.Row = wfila_act
'   grid_fac.SetFocus
   Exit Sub
'End If

End Sub

Public Sub LLENADOS(cont As ListBox, tip As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = tip
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    cont.AddItem " "
    Do Until tab_mayor.EOF
        cont.AddItem tab_mayor!tab_nomlargo & String(60, " ") & tab_mayor!TAB_NUMTAB
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
End Sub

Private Sub txtdoc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 PUB_NUMSER = Val(tserie.Text)
 PUB_NUMFAC = Val(txtdoc.Text)
 LLENA_DOCU
End If
End Sub

Private Sub UNIDAD_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
 UNIDAD.Visible = False
 grid_fac.SetFocus
End If


If KeyAscii <> 13 Then Exit Sub
SQ_OPER = 1
pu_codcia = LK_CODCIA
PUB_CODART = Val(grid_fac.TextMatrix(grid_fac.Row, 10))
PUB_SECUEN = Val(Right(UNIDAD.Text, 4))
LEER_PRE_LLAVE
grid_fac.TextMatrix(grid_fac.Row, 3) = Trim(Left(UNIDAD.Text, 12))
grid_fac.TextMatrix(grid_fac.Row, 4) = "" 'Format(Val(grid_fac.TextMatrix(grid_fac.Row, 11)) / Val(grid_fac.TextMatrix(grid_fac.Row, 17)), "0.00")
grid_fac.TextMatrix(grid_fac.Row, 7) = redondea(Nulo_Valor0(pre_llave!pre_peso) * Val(grid_fac.TextMatrix(grid_fac.Row, 2)))
grid_fac.TextMatrix(grid_fac.Row, 11) = pre_llave!pre_secuencia
grid_fac.TextMatrix(grid_fac.Row, 12) = pre_llave!pre_equiv
grid_fac.TextMatrix(grid_fac.Row, 16) = pre_llave!pre_UNIDAD
stock.Caption = Format(Val(grid_fac.TextMatrix(grid_fac.Row, 15)) / Val(grid_fac.TextMatrix(grid_fac.Row, 12)), "0.00")
unid.Caption = grid_fac.TextMatrix(grid_fac.Row, 16)
nomarti.Caption = grid_fac.TextMatrix(grid_fac.Row, 0)

UNIDAD.Visible = False
suma_grid
grid_fac.COL = 4
grid_fac_KeyPress 13

End Sub
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
 WMONEDA = "S"


'End If
  If Trim(Nulo_Valors(par_llave!PAR_DEVICE_FBG)) <> "" Then
     numero_device = 0
     Reportes.PrinterName = Printers(numero_device).DeviceName
     Reportes.PrinterDriver = Printers(numero_device).DriverName '"RASDD.DLL"
     Reportes.PrinterPort = Printers(numero_device).Port
  End If

    FrmTranf.Reportes.Connect = PUB_ODBC
    FrmTranf.Reportes.Destination = crptToWindow  '= crptToPrinter
    FrmTranf.Reportes.WindowLeft = 2
    FrmTranf.Reportes.WindowTop = 70
    FrmTranf.Reportes.WindowWidth = 635
    FrmTranf.Reportes.WindowHeight = 390
    FrmTranf.Reportes.Formulas(1) = ""
    PUB_NETO = Val(txttotal.Text)
    PU_NUMSER = Val((tserie.Text))
    PU_NUMFAC = Val((txtdoc.Text))
    FrmTranf.Reportes.Formulas(1) = ""
    FrmTranf.Reportes.Formulas(1) = "SON_EFECTIVO=  ' " & CONVER_LETRAS(PUB_NETO, WMONEDA) & "'"
    FrmTranf.Reportes.WindowTitle = "GUIA DE COTIZACION  :" & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "00000000")
    FrmTranf.Reportes.ReportFileName = wRuta + "COTI.RPT"
    pub_cadena = "{PEDIDOS.PED_TIPMOV} = " & LOC_TIPMOV & " AND {PEDIDOS.PED_CODCIA} = '" & LK_CODCIA & "' AND  {PEDIDOS.PED_NUMSER} = '" & PU_NUMSER & "' AND {PEDIDOS.PED_NUMFAC} = " & PU_NUMFAC
    FrmTranf.Reportes.SelectionFormula = pub_cadena
    On Error GoTo accion
    FrmTranf.Reportes.Action = 1
    On Error GoTo 0
Exit Function
accion:
 MsgBox Err.Description
 MsgBox "Intente Nuevamente, la impresion de Modo manual", 48, Pub_Titulo
 Exit Function

End Function

Private Sub txtcli_GotFocus()
 Azul txtcli, txtcli
End Sub

Private Sub txtcli_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If Not ListView1.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And txtcli.Text = "" Then
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
  ListView1.ListItems.Item(loc_key).Selected = True
  ListView1.ListItems.Item(loc_key).EnsureVisible
  txtcli.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
  txtcli.SelStart = Len(txtcli.Text)
fin:

End Sub
Private Sub txtcli_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer
Dim i
Dim itmFound As ListItem
On Error GoTo SALCODI

If KeyAscii = 27 Then
 txtcli.Text = ""
 lblcli.Caption = ""
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
On Error GoTo CODI_ERR

If pu_codclie <> 0 And IsNumeric(txtcli.Text) = True Then
   If Len(Trim(txtcli.Text)) = LK_DIG_RUC Then ' LONG DEL RUC
        pu_cp = "C"
        PUB_RUC = Trim(txtcli.Text)
        SQ_OPER = 4
        pu_codcia = LK_CODCIA
        LEER_CLI_LLAVE
        If cli_ruc.EOF Then
           MsgBox "R.U.C. No Existe ", 48, Pub_Titulo
           Exit Sub
        End If
        txtcli.Text = cli_ruc!cli_codclie
   End If
   SQ_OPER = 1
   On Error GoTo mucho
   pu_codcia = LK_CODCIA
   pu_cp = "C"
   pu_codclie = Val(txtcli.Text)
   LEER_CLI_LLAVE
   On Error GoTo 0
   If cli_llave.EOF Then
     Azul txtcli, txtcli
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     txtcli.SetFocus
     GoTo fin
   End If
   ListView1.Visible = False
   txtcli.Text = cli_llave!cli_codclie
   FrmTranf.lblcli.Caption = cli_llave!cli_nombre
   

   Screen.MousePointer = 0
Else
   If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
     Exit Sub
   End If
   valor = UCase(ListView1.ListItems.Item(loc_key).Text)
   If Trim(UCase(txtcli.Text)) = Left(valor, Len(Trim(txtcli.Text))) Then
   Else
      Exit Sub
   End If
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   pu_cp = "C"
   pu_codclie = Val(ListView1.ListItems.Item(loc_key).SubItems(1))
   LEER_CLI_LLAVE
   On Error GoTo 0
   If cli_llave.EOF Then
     Azul txtcli, txtcli
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     txtcli.SetFocus
     GoTo fin
   End If
   ListView1.Visible = False
   txtcli.Text = cli_llave!cli_codclie
   FrmTranf.lblcli.Caption = cli_llave!cli_nombre
   

End If
dale:
ListView1.Visible = False
fin:
mucho:
CODI_ERR:
Exit Sub
SALCODI:
MsgBox Err.Description & " Intente Nuevamente ", 48, Pub_Titulo
Unload frmCLI
End Sub

Private Sub txtcli_KeyUp(KeyCode As Integer, Shift As Integer)
Dim NADA
Dim VAR
If Len(txtcli.Text) = 0 Or IsNumeric(txtcli.Text) = True Then
   ListView1.Visible = False
   Exit Sub
End If
If ListView1.Visible = False And KeyCode <> 13 Or Len(txtcli.Text) = 1 Then
    If txtcli.Text = "" Then txtcli.Text = " "
    VAR = Asc(txtcli.Text)
    VAR = VAR + 1
    NADA = VAR
    If VAR = 33 Or VAR = 91 Then
       VAR = "ZZZZZZZZ"
    Else
       VAR = Chr(VAR)
    End If
    numarchi = 1
    archi = "SELECT CLI_CODCLIE, CLI_CODCIA, CLI_CP, CLI_NOMBRE,CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM FROM CLIENTES WHERE  CLI_CP = 'C' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & txtcli.Text & "' AND  '" & VAR & "' ORDER BY CLI_NOMBRE"
    PROC_LISVIEW ListView1
    loc_key = 1
    If NADA = 33 Or NADA = 91 Then
      If ListView1.Visible = False Then
        loc_key = 0
        MsgBox "No existe Datos ...", 48, Pub_Titulo
        txtcli.Text = ""
      End If
    End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As ListItem    ' Variable FoundItem.
If ListView1.Visible Then
  Set itmFound = ListView1.FindItem(LTrim(txtcli.Text), lvwText, , lvwPartial)
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
  End If
  Exit Sub
End If
End Sub


Public Sub LLENA_DOCU()
Dim MON As String
LIMPIA_DATOS
CABE_MAN

PSLOC_WARTI(0) = LK_CODCIA
PSLOC_WARTI(1) = 201
PSLOC_WARTI(2) = PUB_NUMSER
PSLOC_WARTI(3) = PUB_NUMFAC
llave_sum_arti.Requery
If llave_sum_arti.EOF Then
  tserie.Text = PUB_NUMSER
  txtdoc.Text = PUB_NUMFAC
  MsgBox "No Existe Cotización.", 48, Pub_Titulo
  Exit Sub
End If
txtigv.Text = llave_sum_arti!PED_IGV
txtvalorv.Text = llave_sum_arti!PED_BRUTO
txttotal.Text = Format(llave_sum_arti!PED_IGV + llave_sum_arti!PED_BRUTO, "0.00")
FrmTranf.lblcli.Caption = llave_sum_arti!PED_NOMCLIE

txtcli.Text = llave_sum_arti!PED_CODCLIE

If llave_sum_arti!PED_MONEDA = "S" Then

 i_moneda.Caption = "S/."
 grid_fac.TextMatrix(1, 4) = "S/."
Else
 i_moneda.Caption = "US$."
 grid_fac.TextMatrix(1, 4) = "US$."
End If
tserie.Text = PUB_NUMSER
txtdoc.Text = PUB_NUMFAC


'llave_sum_arti!PED_HORA

fila = 2
Do Until llave_sum_arti.EOF
   SQ_OPER = 1
   PUB_KEY = llave_sum_arti!PED_CODART
   pu_codcia = LK_CODCIA
   LEER_ART_LLAVE
   grid_fac.Rows = grid_fac.Rows + 1
   grid_fac.RowHeight(grid_fac.Rows - 1) = 285
   grid_fac.TextMatrix(fila, 1) = art_LLAVE!ART_ALTERNO
   grid_fac.TextMatrix(fila, 0) = art_LLAVE!ART_NOMBRE
   
   grid_fac.TextMatrix(fila, 2) = llave_sum_arti!PED_CANTIDAD
   grid_fac.TextMatrix(fila, 4) = llave_sum_arti!PED_PRECIO
   grid_fac.TextMatrix(fila, 10) = llave_sum_arti!PED_CODART
   grid_fac.TextMatrix(fila, 3) = llave_sum_arti!PED_UNIDAD
   grid_fac.TextMatrix(fila, 12) = llave_sum_arti!PED_EQUIV
   grid_fac.TextMatrix(fila, 5) = llave_sum_arti!PED_DESCTO
   fila = fila + 1
   llave_sum_arti.MoveNext
Loop
suma_grid
grid_fac.Enabled = False
f1.Enabled = False


Azul txtdoc, txtdoc
End Sub
