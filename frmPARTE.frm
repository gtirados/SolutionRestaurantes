VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "Comctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form frmPARTE 
   Caption         =   "Parte Diario Grifero"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6720
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5205
   ScaleWidth      =   6720
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdeliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   3000
      TabIndex        =   47
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   6000
      TabIndex        =   46
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton CmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   240
      TabIndex        =   45
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdmostrar 
      Caption         =   "Mostar &Parte"
      Height          =   320
      Left            =   480
      TabIndex        =   44
      Top             =   600
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos :"
      Height          =   975
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9375
      Begin VB.TextBox t_numfac 
         Height          =   300
         Left            =   960
         TabIndex        =   0
         Top             =   200
         Width           =   1095
      End
      Begin VB.TextBox t_isla 
         Height          =   285
         Left            =   6960
         TabIndex        =   4
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox t_turno 
         Height          =   285
         Left            =   3600
         TabIndex        =   3
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox t_grifero 
         Height          =   285
         Left            =   3600
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
      Begin MSMask.MaskEdBox t_fecha 
         Height          =   285
         Left            =   6960
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   327680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label lblisla 
         Height          =   255
         Left            =   7680
         TabIndex        =   19
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblturno 
         Height          =   255
         Left            =   4320
         TabIndex        =   18
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label lblparte 
         Caption         =   "Nº Parte :"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
      Begin VB.Label lblgrifero 
         Height          =   255
         Left            =   4320
         TabIndex        =   14
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label lblparte 
         Caption         =   "Isla :"
         Height          =   255
         Index           =   3
         Left            =   6360
         TabIndex        =   13
         Top             =   600
         Width           =   495
      End
      Begin VB.Label lblparte 
         Caption         =   "Turno :"
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   12
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lblparte 
         Caption         =   "Fecha :"
         Height          =   255
         Index           =   1
         Left            =   6360
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblparte 
         Caption         =   "Grifero :"
         Height          =   255
         Index           =   0
         Left            =   2880
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdcerrar 
      Caption         =   "Ce&rrar"
      Height          =   375
      Left            =   8040
      TabIndex        =   8
      Top             =   5280
      Width           =   1215
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   495
      Left            =   6840
      TabIndex        =   17
      Top             =   5760
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
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame fgalones 
      Caption         =   "Lectura de Gagoles :"
      Height          =   2295
      Left            =   0
      TabIndex        =   5
      Top             =   960
      Width           =   9375
      Begin VB.TextBox textovar 
         Height          =   285
         Left            =   2520
         MaxLength       =   9
         TabIndex        =   16
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid grid_lec 
         Height          =   2055
         Left            =   120
         TabIndex        =   7
         Top             =   195
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   3625
         _Version        =   327680
         Rows            =   3
         FixedRows       =   2
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   0
      TabIndex        =   6
      Top             =   3240
      Width           =   9375
      Begin VB.CommandButton impd 
         Caption         =   "Imp. Detalle "
         Height          =   375
         Left            =   7680
         TabIndex        =   50
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmdAct 
         Caption         =   "Actualizar Valores de P/D."
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
         Left            =   6360
         TabIndex        =   48
         Top             =   1560
         Width           =   2895
      End
      Begin VB.TextBox t_vales 
         Height          =   285
         Left            =   4800
         TabIndex        =   42
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox t_entrega 
         Height          =   285
         Left            =   7800
         TabIndex        =   40
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox t_falsob 
         Height          =   285
         Left            =   4800
         TabIndex        =   38
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox t_voucher 
         Height          =   285
         Left            =   4800
         TabIndex        =   36
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox t_cheques 
         Height          =   285
         Left            =   4800
         TabIndex        =   34
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox t_desfacont 
         Height          =   285
         Left            =   4800
         TabIndex        =   32
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox t_desodcont 
         Height          =   285
         Left            =   1800
         TabIndex        =   30
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox t_desfacred 
         Height          =   285
         Left            =   1800
         TabIndex        =   28
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox t_desodcred 
         Height          =   285
         Left            =   1800
         TabIndex        =   26
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox t_credito 
         Height          =   285
         Left            =   1800
         TabIndex        =   24
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox t_efectivo 
         Height          =   285
         Left            =   1800
         TabIndex        =   22
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox t_totalgen 
         Height          =   285
         Left            =   7800
         TabIndex        =   20
         Top             =   480
         Width           =   1215
      End
      Begin ComctlLib.ProgressBar pb 
         Height          =   135
         Left            =   6360
         TabIndex        =   49
         Top             =   1440
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   238
         _Version        =   327682
         Appearance      =   0
      End
      Begin VB.Label LTOT 
         Caption         =   "Faltante/Sobrante:"
         Height          =   255
         Index           =   11
         Left            =   3120
         TabIndex        =   43
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label LTOT 
         Caption         =   "Entrega Total :"
         Height          =   255
         Index           =   10
         Left            =   6120
         TabIndex        =   41
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label LTOT 
         Caption         =   "Vales "
         Height          =   255
         Index           =   9
         Left            =   3120
         TabIndex        =   39
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label LTOT 
         Caption         =   "Voucher:"
         Height          =   255
         Index           =   8
         Left            =   3120
         TabIndex        =   37
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label LTOT 
         Caption         =   "Cheques :"
         Height          =   255
         Index           =   7
         Left            =   3120
         TabIndex        =   35
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label LTOT 
         Caption         =   "Dscto. C.Fact.Contado"
         Height          =   255
         Index           =   6
         Left            =   3120
         TabIndex        =   33
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label LTOT 
         Caption         =   "Dscto. C.O.D. Contado"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   31
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label LTOT 
         Caption         =   "Dscto.C.Fact. Credito:"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   29
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label LTOT 
         Caption         =   "Dscto. C.O.D. Credito:"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   27
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label LTOT 
         Caption         =   "Credito:"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   25
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label l_efectivo 
         Caption         =   "Efectivo :"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   120
         Width           =   1695
      End
      Begin VB.Label LTOT 
         Caption         =   "Total General :"
         Height          =   255
         Index           =   0
         Left            =   6120
         TabIndex        =   21
         Top             =   480
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport Reportes 
      Left            =   120
      Top             =   6000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
End
Attribute VB_Name = "frmPARTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim pag_llave As rdoResultset
Dim PSPAG_LLAVE As rdoQuery
Dim loc_key As Integer
Dim PSMA_LLAVE As rdoQuery
Dim man_llave As rdoResultset
Dim pag_busca As rdoResultset
Dim PSPAG_BUSCA As rdoQuery
Dim FLAG_NUEVO As String * 1
Dim tempo_isla As Integer
Dim PSACT_LLAVE As rdoQuery
Dim actparte_llave As rdoResultset




Private Sub cancelar_Click()
FLAG_NUEVO = ""
LIMPIA
OBBLO
CmdModificar.Caption = "&Modificar"
cmdmostrar.Enabled = True
CmdModificar.Enabled = False
cmdEliminar.Enabled = False
t_numfac.Text = ""
grid_lec.Cols = 2
t_numfac.SetFocus
End Sub

Private Sub cmdAct_Click()
Dim x_Credito As Currency
Dim x_Efectivo As Currency
Dim x_Desfacont As Currency
Dim x_Desodcont As Currency
Dim x_Desfacred As Currency
Dim x_Desodcred As Currency
Dim x_Voucher As Currency
Dim x_Vales As Currency

pub_cadena = "SELECT  FAR_CANTIDAD, FAR_CODART, FAR_DESCTO, FAR_NUM_LOTE, FAR_SUBTOTAL FROM FACART WHERE  FAR_CODCIA = ? AND FAR_FECHA = ? AND FAR_ISLA = ? AND FAR_TURNO = ? and FAR_ESTADO <> 'E' AND FAR_ESTADO <> 'M' ORDER BY FAR_CODART "
Set PSACT_LLAVE = CN.CreateQuery("", pub_cadena)
PSACT_LLAVE(0) = 0
PSACT_LLAVE(1) = LK_FECHA_DIA
PSACT_LLAVE(2) = 0
PSACT_LLAVE(3) = 0
Set actparte_llave = PSACT_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
PSACT_LLAVE(0) = LK_CODCIA
PSACT_LLAVE(1) = t_fecha.Text
PSACT_LLAVE(2) = Val(t_isla.Text)
PSACT_LLAVE(3) = Val(t_turno.Text)
actparte_llave.Requery
If actparte_llave.EOF Then
  MsgBox "No Existe Documentos para este Parte.", 48, Pub_Titulo
  Exit Sub
End If
PB.Min = 0
PB.Max = actparte_llave.RowCount
PB.Value = 0
PB.Visible = True
x_Credito = 0
x_Efectivo = 0
x_Desfacont = 0
x_Desodcont = 0
x_Desfacred = 0
x_Desodcred = 0
x_Voucher = 0
x_Vales = 0

Do Until actparte_llave.EOF
'  If Val(actparte_llave!far_DESCTO) <> 0 Then Stop
    PB.Value = PB.Value + 1
    If actparte_llave!far_NUM_LOTE = 0 Then
      'x_Efectivo = x_Efectivo + actparte_llave!far_limcre_ANT
      x_Desfacont = x_Desfacont + actparte_llave!FAR_DESCTO
    ElseIf actparte_llave!far_NUM_LOTE = 15 Then
      'x_Efectivo = x_Efectivo + actparte_llave!far_limcre_ANT
      x_Desodcont = x_Desodcont + actparte_llave!FAR_DESCTO
    ElseIf actparte_llave!far_NUM_LOTE = 1 Then
      'x_Desfacred = 0
    ElseIf actparte_llave!far_NUM_LOTE = 20 Then
      x_Credito = x_Credito + actparte_llave!FAR_SUBTOTAL
      x_Desodcred = x_Desodcred + actparte_llave!FAR_DESCTO
    ElseIf actparte_llave!far_NUM_LOTE = 5 Then
      x_Voucher = x_Voucher + actparte_llave!FAR_SUBTOTAL
      x_Desfacont = x_Desfacont + actparte_llave!FAR_DESCTO
    ElseIf actparte_llave!far_NUM_LOTE = 25 Then
      x_Vales = x_Vales + actparte_llave!FAR_SUBTOTAL
    End If
  actparte_llave.MoveNext
Loop

t_credito.Text = Format(x_Credito, "0.00")
t_efectivo.Text = Format(x_Efectivo, "0.00")
t_desfacont.Text = Format(x_Desfacont, "0.00")
t_desodcont.Text = Format(x_Desodcont, "0.00")
t_desfacred.Text = Format(x_Desfacred, "0.00")
t_desodcred.Text = Format(x_Desodcred, "0.00")
t_voucher.Text = Format(x_Voucher, "0.00")
t_vales.Text = Format(x_Vales, "0.00")
calcula2
PB.Visible = False
End Sub

Private Sub cmdcerrar_Click()
Unload frmPARTE
End Sub

Private Sub cmdEliminar_Click()
pag_llave.Requery
If pag_llave.EOF Then Exit Sub
pub_mensaje = "Desea Eliminar Parte Diario de Grifero  Nº: " & Trim(t_numfac.Text)
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If
Do Until pag_llave.EOF
   pag_llave.Delete
 pag_llave.MoveNext
Loop
cancelar_Click
End Sub

Private Sub CmdModificar_Click()
If Left(CmdModificar.Caption, 2) = "&M" Then
    cmdEliminar.Enabled = False
    CmdModificar.Caption = "&Grabar"
    OBDES
    t_grifero.Locked = False
    t_grifero.BackColor = QBColor(7)
    t_isla.Locked = True
    t_isla.BackColor = QBColor(7)
    t_turno.Locked = True
    t_turno.BackColor = QBColor(7)
    t_fecha.Enabled = False
    t_grifero.SetFocus
Else
 CmdModificar.Caption = "&Modificar"
 cmdEliminar.Enabled = True
 GRABA_PAG
 
 PSMA_LLAVE(0) = LK_CODCIA
 PSMA_LLAVE(1) = 2105
 PSMA_LLAVE(2) = Val(t_isla.Text)
 man_llave.Requery
 Do Until man_llave.EOF
  man_llave.Edit
  man_llave!tab_nomlargo = Trim(grid_lec.TextMatrix(2, man_llave.AbsolutePosition))
  man_llave.Update
  man_llave.MoveNext
 Loop
 OBBLO
 LIMPIA
 FLAG_NUEVO = ""
 cmdmostrar.Enabled = True
 t_numfac.SetFocus
End If



End Sub

Private Sub cmdmostrar_Click()
If Val(t_numfac.Text) = 0 Then Exit Sub
PSPAG_LLAVE(0) = LK_CODCIA
PSPAG_LLAVE(1) = Val(t_numfac.Text)
pag_llave.Requery
If pag_llave.EOF Then
  pub_mensaje = "No Existe Num Parte . Desea Agregar uno Nuevo?"
  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
  If Pub_Respuesta = vbNo Then
     Azul t_numfac, t_numfac
     Exit Sub
  End If
   cmdmostrar.Enabled = False
   FLAG_NUEVO = "A"
   OBDES
   CmdModificar.Caption = "&Grabar Parte"
   CmdModificar.Enabled = True
   t_grifero.SetFocus
   Exit Sub
End If
 
PSMA_LLAVE(0) = LK_CODCIA
PSMA_LLAVE(1) = 2105
PSMA_LLAVE(2) = Val(pag_llave!PAG_ISLA)
man_llave.Requery
If pag_llave.RowCount <> man_llave.RowCount Then
  MsgBox "Ojo.... Se a agregado mas mangueras en esta isla ", 48, Pub_Titulo
End If


grid_lec.Clear
grid_lec.Cols = pag_llave.RowCount + 1
grid_lec.Rows = 8
cabe
t_isla.Text = pag_llave!PAG_ISLA
PUB_CODVEN = Val(t_isla.Text)
pu_codcia = LK_CODCIA
SQ_OPER = 1
LEER_VEN_LLAVE
If Not ven_llave.EOF Then
  lblisla.Caption = ven_llave!VEM_NOMBRE
End If
t_turno.Text = pag_llave!pag_turno
SQ_OPER = 1
PUB_TIPREG = 2102
PUB_NUMTAB = Val(t_turno.Text)
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
If Not tab_llave.EOF Then
  lblturno.Caption = tab_llave!tab_nomlargo
End If
t_grifero.Text = pag_llave!PAG_CODRES
SQ_OPER = 1
PUB_TIPREG = 2103
PUB_NUMTAB = Val(t_grifero.Text)
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
If Not tab_llave.EOF Then
  lblgrifero.Caption = tab_llave!tab_nomlargo
End If
t_fecha.Text = pag_llave!pag_fecha
Do Until pag_llave.EOF
  grid_lec.ColWidth(pag_llave.AbsolutePosition) = 1100
  grid_lec.TextMatrix(0, pag_llave.AbsolutePosition) = "(" & pag_llave!PAG_SURTIDOR & ")"
  SQ_OPER = 1
  pu_codcia = LK_CODCIA
  PUB_KEY = Val(pag_llave!PAG_CODART)
  LEER_ART_LLAVE
  If art_LLAVE.EOF Then MsgBox "Producto no Existe Codigo: " & PUB_KEY, 48, Pub_Titulo
  grid_lec.TextMatrix(1, pag_llave.AbsolutePosition) = art_LLAVE!ART_NOMBRE
  grid_lec.TextMatrix(2, pag_llave.AbsolutePosition) = pag_llave!PAG_LEC_CIERRE
  grid_lec.TextMatrix(3, pag_llave.AbsolutePosition) = pag_llave!PAG_LEC_INICIO
  grid_lec.TextMatrix(4, pag_llave.AbsolutePosition) = 0
  grid_lec.TextMatrix(5, pag_llave.AbsolutePosition) = pag_llave!PAG_PRECIO
  grid_lec.TextMatrix(7, pag_llave.AbsolutePosition) = PUB_KEY
  t_efectivo.Text = pag_llave!PAG_EFECTIVO
  t_credito.Text = pag_llave!PAG_CREDITO
  t_desodcred.Text = pag_llave!PAG_DESODCRED
  t_desfacred.Text = pag_llave!PAG_DESFACRED
  t_desodcont.Text = pag_llave!PAG_DESODCONT
  t_desfacont.Text = pag_llave!PAG_DESFACONT
  t_cheques.Text = pag_llave!PAG_CHEQUE
  t_vales.Text = pag_llave!PAG_VALE
  t_voucher.Text = pag_llave!PAG_VOUCHER
  t_falsob.Text = pag_llave!PAG_FALSOB
  t_entrega.Text = pag_llave!PAG_ENTREGA
  t_totalgen.Text = pag_llave!PAG_TOTAL

  pag_llave.MoveNext
Loop
calcula
calcula2
cmdEliminar.Enabled = True
CmdModificar.Enabled = True
cmdmostrar.Enabled = False
CmdModificar.SetFocus
Exit Sub


End Sub

Private Sub Form_Load()

pub_cadena = "SELECT * FROM TABLAS WHERE TAB_CODCIA = ? AND TAB_TIPREG = ?  AND TAB_CONTABLE2 = ? ORDER BY TAB_NUMTAB "
Set PSMA_LLAVE = CN.CreateQuery("", pub_cadena)
PSMA_LLAVE(0) = 0
PSMA_LLAVE(1) = 0
PSMA_LLAVE(2) = 0
Set man_llave = PSMA_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM PARGRI WHERE PAG_CODCIA = ?   AND PAG_NUMPARTE = ? ORDER BY PAG_NUMPARTE"
Set PSPAG_LLAVE = CN.CreateQuery("", pub_cadena)
PSPAG_LLAVE(0) = 0
PSPAG_LLAVE(1) = 0
Set pag_llave = PSPAG_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT PAG_ISLA, PAG_TURNO, PAG_FECHA, PAG_NUMPARTE  FROM PARGRI WHERE PAG_CODCIA = ? AND PAG_ISLA = ? AND  PAG_TURNO = ? AND PAG_FECHA = ?  ORDER BY PAG_NUMPARTE"
Set PSPAG_BUSCA = CN.CreateQuery("", pub_cadena)
PSPAG_BUSCA(0) = 0
PSPAG_BUSCA(1) = 0
PSPAG_BUSCA(2) = 0
PSPAG_BUSCA(3) = LK_FECHA_DIA
Set pag_busca = PSPAG_BUSCA.OpenResultset(rdOpenKeyset, rdConcurValues)


t_fecha.Visible = True
t_fecha.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
t_fecha.Mask = "##/##/####"
OBBLO
FLAG_NUEVO = ""
cmdEliminar.Enabled = False
CmdModificar.Enabled = False
End Sub

Public Sub cabe()
grid_lec.Rows = 8
grid_lec.ColWidth(0) = 1300
grid_lec.TextMatrix(0, 0) = "MANGUERA    :"
grid_lec.TextMatrix(1, 0) = "PRODUCTO    :"
grid_lec.TextMatrix(2, 0) = "AL CIERRE   :"
grid_lec.TextMatrix(3, 0) = "AL INICIO   :"
grid_lec.TextMatrix(4, 0) = "TOTAL       :"
grid_lec.TextMatrix(5, 0) = "PRECIO COMB.:"
grid_lec.TextMatrix(6, 0) = "EMTREGA     :"
'grid_lec.RowHeight(0) = 285
'grid_lec.RowHeight(1) = 285
grid_lec.RowHeight(2) = 260
grid_lec.RowHeight(3) = 260
grid_lec.RowHeight(4) = 285
grid_lec.RowHeight(5) = 260
grid_lec.RowHeight(6) = 260
grid_lec.RowHeight(7) = 0

End Sub

Private Sub grid_lec_EnterCell()
If grid_lec.CellHeight < 0 Then Exit Sub
textovar.Visible = False
textovar.Text = Trim(grid_lec.TextMatrix(grid_lec.Row, grid_lec.COL))
textovar.Left = grid_lec.Left + grid_lec.CellLeft
textovar.Width = grid_lec.CellWidth
textovar.Height = grid_lec.CellHeight
textovar.Top = grid_lec.Top + grid_lec.CellTop  ' ESTADO.Top + grid_lec.Top + grid_lec.CellTop - 1560 '480
If grid_lec.COL = 1 Then
 If Trim(grid_lec.TextMatrix(grid_lec.Row, 1)) = "" Then
   If Trim(grid_lec.TextMatrix(grid_lec.Row, 0)) <> "" Then
    textovar.Visible = True
    textovar.SetFocus
   End If
 End If
End If
If grid_lec.COL = 0 Then
 If Trim(grid_lec.TextMatrix(grid_lec.Row, 0)) = "" Then
  textovar.Visible = True
  textovar.SetFocus
 End If
End If


End Sub



Private Sub grid_lec_KeyPress(KeyAscii As Integer)
Dim a As Integer
Dim t, wc
Dim wprecios As String * 12
Static CONS
Dim wactivo As Integer
If KeyAscii <> 13 Then Exit Sub
If grid_lec.Rows < 1 Then Exit Sub
If grid_lec.Row <> 2 Then Exit Sub
    textovar.Visible = True
    Azul textovar, textovar
    textovar.SetFocus
Exit Sub
pasa:
Resume Next

End Sub

Private Sub grid_lec_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wc
Dim a, WF As Integer
Dim tf, t, tC
Dim SALE As Boolean
If KeyCode = 114 Then
  Azul t_efectivo, t_efectivo
  Exit Sub
End If
If KeyCode = 45 Then
   If Trim(grid_lec.TextMatrix(grid_lec.Rows - 1, 2)) = "" Then Exit Sub
   grid_lec.Rows = grid_lec.Rows + 1
   grid_lec.COL = 0
   grid_lec.RowHeight(grid_lec.Rows - 1) = 285
   grid_lec.Row = grid_lec.Rows - 1
   Exit Sub
End If

If KeyCode = 46 Then
If grid_lec.Rows <= 2 Then Exit Sub
If grid_lec.Rows <= 3 Then
    pub_mensaje = MsgBox("Quitar el Producto para la Orden de Compra ", vbYesNo + vbExclamation + vbDefaultButton2, Pub_Titulo)
    If pub_mensaje = vbNo Then
      grid_lec.SetFocus
      Exit Sub
    End If
Else
   pub_mensaje = MsgBox("Quitar el Surtidor ? ", vbYesNo + vbExclamation + vbDefaultButton2, Pub_Titulo)
   If pub_mensaje = vbNo Then
     grid_lec.SetFocus
     Exit Sub
   Else
   '  grid_lec.RowHeight(grid_lec.Row) = 1
   grid_lec.RemoveItem (grid_lec.Row)
   grid_lec.Row = grid_lec.Row
   grid_lec.Refresh
   grid_lec.SetFocus
   End If
End If
End If
'grid_lec.SetFocus
Exit Sub



End Sub

Private Sub impd_Click()
  If Val(t_isla.Text) = 0 Then Exit Sub
  If Val(t_turno.Text) = 0 Then Exit Sub
  
  Dim Wfile
  Wfile = "DETPAR.RPT"
  Reportes.Connect = PUB_ODBC
  Reportes.ReportFileName = PUB_RUTA_OTRO & Wfile
  Reportes.WindowTitle = "Reporte :  " & "Dettale de Parte Diario Grifero" & " - Archivo:(" & Wfile & ")"
  Reportes.Destination = crptToWindow
  Reportes.WindowLeft = 2
  Reportes.WindowTop = 70
  Reportes.WindowWidth = 635
  Reportes.WindowHeight = 390
  Reportes.Formulas(0) = "CIA=  '" & Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)) & "'"
  Reportes.Formulas(1) = "FECHAS=  '" & Format(t_fecha.Text, "dd/mm/yyyy") & "'"
  DIA = Day(t_fecha.Text)
  MES = Month(t_fecha.Text)
  ANO = Year(t_fecha.Text)
  'Reportes.Formulas(50) = "FECHADIA= Date ( " & ANO & "," & MES & "," & DIA & ")"
  pub_cadena = "{FACART.FAR_TURNO} = " & Trim(t_turno.Text) & " AND {FACART.FAR_ISLA} = " & Trim(t_isla.Text) & " AND {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_FECHA_COMPRA}  = Date ( " & ANO & "," & MES & "," & DIA & ")"
  'Debug.Print pub_cadena
  Reportes.SelectionFormula = pub_cadena
 ' Debug.Print pub_cadena
  Reportes.Action = 1
  Screen.MousePointer = 0
  
  
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then ListView1.Visible = False
End Sub

Private Sub ListView1_LostFocus()
ListView1.Visible = False

End Sub

Private Sub t_cheques_Change()
calcula2
End Sub

Private Sub t_cheques_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL t_cheques, KeyAscii
If KeyAscii = 13 Then
  Azul t_voucher, t_voucher
End If

End Sub

Private Sub t_credito_Change()
calcula2
End Sub

Private Sub t_credito_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL t_credito, KeyAscii
If KeyAscii = 13 Then
  Azul t_desodcred, t_desodcred
End If

End Sub

Private Sub t_desfacont_Change()
calcula2
End Sub

Private Sub t_desfacont_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL t_desfacont, KeyAscii
If KeyAscii = 13 Then
  Azul t_cheques, t_cheques
End If

End Sub

Private Sub t_desfacred_Change()
calcula2
End Sub

Private Sub t_desfacred_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL t_desfacred, KeyAscii
If KeyAscii = 13 Then
  Azul t_desodcont, t_desodcont
End If

End Sub

Private Sub t_desodcont_Change()
calcula2
End Sub

Private Sub t_desodcont_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL t_desodcont, KeyAscii
If KeyAscii = 13 Then
  Azul t_desfacont, t_desfacont
End If

End Sub

Private Sub t_desodcred_Change()
calcula2
End Sub

Private Sub t_desodcred_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL t_desodcred, KeyAscii
If KeyAscii = 13 Then
  Azul t_desfacred, t_desfacred
End If

End Sub

Private Sub t_efectivo_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL t_efectivo, KeyAscii
If KeyAscii = 13 Then
  Azul t_credito, t_credito
End If
End Sub

Private Sub t_entrega_Change()
calcula2
End Sub

Private Sub t_entrega_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL t_entrega, KeyAscii
If KeyAscii = 13 Then
   Azul t_totalgen, t_totalgen
End If

End Sub

Private Sub t_falsob_Change()
calcula2
End Sub

Private Sub t_falsob_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL t_falsob, KeyAscii
If KeyAscii = 13 Then
  Azul t_entrega, t_entrega
End If

End Sub

Private Sub t_fecha_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
Azul t_isla, t_isla
End Sub

Private Sub t_grifero_Change()
lblgrifero.Caption = ""
End Sub

Private Sub t_grifero_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If Not ListView1.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And t_grifero.Text = "" Then
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
  t_grifero.Text = Trim(ListView1.ListItems.Item(loc_key).Text) & " "
  DoEvents
  t_grifero.SelStart = Len(t_grifero.Text)
  DoEvents
fin:

End Sub

Private Sub t_grifero_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  t_grifero.Text = ""
  lblgrifero.Caption = ""
  ListView1.Visible = False
  Exit Sub
End If
If KeyAscii <> 13 Then Exit Sub
If Val(t_grifero.Text) <> 0 And IsNumeric(t_grifero.Text) = True Then
   loc_key = 0
   SQ_OPER = 1
   PUB_TIPREG = 2103
   PUB_NUMTAB = Val(t_grifero.Text)
   PUB_CODCIA = LK_CODCIA
   LEER_TAB_LLAVE
   If tab_llave.EOF Then
     Azul t_grifero, t_grifero
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     t_grifero.SetFocus
     GoTo fin
   End If
   t_grifero.Text = PUB_NUMTAB
   lblgrifero.Caption = Trim(tab_llave!tab_nomlargo)
   ListView1.Visible = False
   Azul t_turno, t_turno
   ListView1.Visible = False
   Screen.MousePointer = 0
   Exit Sub
Else
   If loc_key > ListView1.ListItems.count Or loc_key = 0 Then
     Exit Sub
   End If
   valor = UCase(ListView1.ListItems.Item(loc_key).Text)
   If Trim(UCase(t_grifero.Text)) = Left(valor, Len(Trim(t_grifero.Text))) Then
   Else
      Exit Sub
   End If
   SQ_OPER = 1
   PUB_TIPREG = 2103
   PUB_NUMTAB = Trim(ListView1.ListItems.Item(loc_key).SubItems(1))
   PUB_CODCIA = LK_CODCIA
   LEER_TAB_LLAVE
   If tab_llave.EOF Then
     Azul t_grifero, t_grifero
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     t_grifero.SetFocus
     GoTo fin
   End If
   t_grifero.Text = PUB_NUMTAB
   lblgrifero.Caption = Trim(tab_llave!tab_nomlargo)
   ListView1.Visible = False
   Azul t_turno, t_turno
End If

fin:


End Sub

Private Sub t_grifero_KeyUp(KeyCode As Integer, Shift As Integer)
Dim VAR
If KeyCode = 27 Then Exit Sub
If Len(t_grifero.Text) = 0 Or t_grifero.Locked = True Or IsNumeric(t_grifero.Text) = True Then
   ListView1.Visible = False
   Exit Sub
End If
If ListView1.Visible = False And KeyCode <> 13 Or Len(t_grifero.Text) = 1 Then
    VAR = Asc(t_grifero.Text)
    VAR = VAR + 1
    If VAR = 33 Or VAR = 91 Then
       VAR = "ZZZZZZZZ"
    Else
       VAR = Chr(VAR)
    End If
    numarchi = 99
    archi = "SELECT * FROM TABLAS WHERE  TAB_CODCIA = '" & LK_CODCIA & "' AND TAB_TIPREG = 2103  AND TAB_NOMLARGO BETWEEN '" & t_grifero.Text & "' AND  '" & VAR & "' ORDER BY TAB_NOMLARGO"
    LISTA_TABLAS ListView1, 1000
    loc_key = 1
    If ListView1.Visible = False Then
        loc_key = 0
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
  Set itmFound = ListView1.FindItem(LTrim(t_grifero.Text), lvwText, , lvwPartial)
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

Private Sub t_isla_Change()
lblisla.Caption = ""
End Sub

Private Sub t_isla_GotFocus()
'tempo_isla = Val(t_isla.Text)
End Sub

Private Sub t_isla_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then Exit Sub
If t_isla.Locked Then
    grid_lec.SetFocus
    grid_lec.COL = 1
    grid_lec.Row = 1
    grid_lec.Row = 2
    Exit Sub
End If
PUB_CODVEN = Val(t_isla.Text)
pu_codcia = LK_CODCIA
SQ_OPER = 1
LEER_VEN_LLAVE
On Error GoTo 0
If ven_llave.EOF Then
  Azul t_isla, t_isla
  MsgBox "Isla NO Existe ...", 48, Pub_Titulo
  Exit Sub
End If

PSPAG_BUSCA(0) = LK_CODCIA
PSPAG_BUSCA(1) = Val(t_isla.Text)
PSPAG_BUSCA(2) = Val(t_turno.Text)
PSPAG_BUSCA(3) = t_fecha.Text
pag_busca.Requery
If Not pag_busca.EOF Then
 MsgBox "Datos Existen, Corresponden a :" & Chr(13) & "N° Parte: " & pag_busca!pag_numparte & Chr(13) & "Fecha: " & Format(pag_busca!pag_fecha, "dd/mm/yyyy") & Chr(13) & "Isla: " & pag_busca!PAG_ISLA & Chr(13) & "Turno: " & pag_busca!pag_turno, 48, Pub_Titulo
 Azul t_numfac, t_numfac
 Exit Sub
End If

GoSub NUEVO
t_isla.Text = ven_llave!vem_codven
lblisla.Caption = ven_llave!VEM_NOMBRE
grid_lec.SetFocus
grid_lec.COL = 1
grid_lec.Row = 1
grid_lec.Row = 2
Exit Sub

NUEVO:
PSMA_LLAVE(0) = LK_CODCIA
PSMA_LLAVE(1) = 2105
PSMA_LLAVE(2) = Val(t_isla.Text)
man_llave.Requery
If man_llave.EOF Then
  MsgBox "Debe Definir Mangueras para la Isla: " & t_isla.Text, 48, Pub_Titulo
  Exit Sub
End If
grid_lec.Clear
grid_lec.Cols = man_llave.RowCount + 1
grid_lec.Rows = 8
cabe
Do Until man_llave.EOF
  grid_lec.ColWidth(man_llave.AbsolutePosition) = 1100
  grid_lec.TextMatrix(0, man_llave.AbsolutePosition) = "(" & man_llave!TAB_NUMTAB & ")"
  SQ_OPER = 1
  pu_codcia = LK_CODCIA
  PUB_KEY = Val(man_llave!TAB_CODART)
  LEER_ART_LLAVE
  If art_LLAVE.EOF Then MsgBox "Producto no Existe Codigo: " & PUB_KEY, 48, Pub_Titulo
  PUB_CODART = PUB_KEY
  PUB_SECUEN = 0
  SQ_OPER = 1
  LEER_PRE_LLAVE
  grid_lec.TextMatrix(1, man_llave.AbsolutePosition) = art_LLAVE!ART_NOMBRE
  grid_lec.TextMatrix(0, man_llave.AbsolutePosition) = "(" & man_llave!TAB_NUMTAB & ")"
  grid_lec.TextMatrix(3, man_llave.AbsolutePosition) = Trim(man_llave!tab_nomlargo)
  grid_lec.TextMatrix(4, man_llave.AbsolutePosition) = 0
  grid_lec.TextMatrix(5, man_llave.AbsolutePosition) = pre_llave!PRE_PRE1
  grid_lec.TextMatrix(7, man_llave.AbsolutePosition) = PUB_KEY
  man_llave.MoveNext
Loop

Return

End Sub

Private Sub t_isla_LostFocus()
'If tempo_isla <> Val(t_isla.Text) Then
' t_isla_KeyPress 13
'End If
End Sub

Private Sub t_numfac_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii <> 13 Then Exit Sub
If cmdmostrar.Enabled Then
    cmdmostrar_Click
Else
    Azul t_grifero, t_grifero
End If

End Sub

Private Sub t_totalgen_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If CmdModificar.Enabled Then CmdModificar.SetFocus
End If

End Sub

Private Sub t_turno_Change()
lblturno.Caption = ""
End Sub

Private Sub t_turno_KeyPress(KeyAscii As Integer)

If KeyAscii <> 13 Then Exit Sub
   SQ_OPER = 1
   PUB_TIPREG = 2102
   PUB_NUMTAB = Val(t_turno.Text)
   PUB_CODCIA = LK_CODCIA
   LEER_TAB_LLAVE
   If tab_llave.EOF Then
     Azul t_turno, t_turno
     MsgBox "Turno NO Existe...", 48, Pub_Titulo
     t_turno.SetFocus
     Exit Sub
   End If
   t_turno.Text = PUB_NUMTAB
   lblturno.Caption = tab_llave!tab_nomlargo
   If t_fecha.Enabled Then
     Azul2 t_fecha, t_fecha
   Else
     Azul t_isla, t_isla
   End If
End Sub

Private Sub t_vales_Change()
calcula2
End Sub

Private Sub t_vales_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL t_vales, KeyAscii
If KeyAscii = 13 Then
  Azul t_falsob, t_falsob
End If

End Sub

Private Sub t_voucher_Change()
calcula2
End Sub

Private Sub t_voucher_KeyPress(KeyAscii As Integer)
SOLO_DECIMAL t_voucher, KeyAscii
If KeyAscii = 13 Then
  Azul t_vales, t_vales
End If

End Sub

Private Sub textovar_Change()
If Not textovar.Visible Then Exit Sub
'If grid_lec.COL = 1 Then
    'grid_lec.TextMatrix(grid_lec.Row, 1) = ""
    'grid_lec.TextMatrix(grid_lec.Row, 2) = ""
    'grid_lec.TextMatrix(grid_lec.Row, 3) = ""
'Else
' If grid_lec.COL = 2 Then
'  grid_lec.Text = textovar.Text
' Else
  grid_lec.Text = Format(textovar.Text, "0.00")
' End If
'End If
End Sub

Private Sub textovar_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then  ' flecha abajo
  If grid_lec.Row = grid_lec.Rows - 1 Then Exit Sub
  If Trim(grid_lec.Text) <> "" Then Exit Sub
  grid_lec.Row = grid_lec.Row + 1
  grid_lec.SetFocus
  Exit Sub
End If
If KeyCode = 38 Then
 If Trim(grid_lec.Text) <> "" Then Exit Sub
 grid_lec.Row = grid_lec.Row - 1
 grid_lec.SetFocus
 Exit Sub
End If
If KeyCode = 39 Then
If Trim(grid_lec.Text) <> "" Then Exit Sub
 grid_lec.COL = grid_lec.COL + 1
 grid_lec.SetFocus
 Exit Sub
End If

 Exit Sub
fin:


End Sub

Private Sub textovar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  'If textovar.Text = "" Then
    textovar.Visible = False
    grid_lec.SetFocus
    Exit Sub
  'End If
  'textovar.Text = "" ' temporal
  Exit Sub
End If
'If grid_lec.COL = 2 Or grid_lec.COL = 4 Or grid_lec.COL = 5 Then Consistencias grid_lec, textovar, KeyAscii
If KeyAscii <> 13 Then Exit Sub
If grid_lec.Row = 2 And Trim(textovar.Text) <> "" Then
  If Val(grid_lec.TextMatrix(3, grid_lec.COL)) > Val(textovar.Text) Then
    MsgBox "EL Nº de Cierre no puede ser menor que el inicial", 48, Pub_Titulo
    Azul textovar, textovar
    Exit Sub
  End If
End If
textovar.Visible = False
textovar.Visible = False
If grid_lec.COL < grid_lec.Cols - 1 Then
     grid_lec.COL = grid_lec.COL + 1
End If
calcula
calcula2

Exit Sub


End Sub


Public Sub LIMPIA()
t_isla.Text = ""
t_turno.Text = ""
t_grifero.Text = ""
t_fecha.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
t_efectivo.Text = ""
t_credito.Text = ""
t_desodcred.Text = ""
t_desfacred.Text = ""
t_desodcont.Text = ""
t_desfacont.Text = ""
t_cheques.Text = ""
t_voucher.Text = ""
t_vales.Text = ""
t_falsob.Text = ""
t_entrega.Text = ""
t_totalgen.Text = ""
t_efectivo.Text = ""
t_falsob.Text = ""
grid_lec.Clear
cabe
textovar.Text = ""
textovar.Visible = False

End Sub

Public Sub calcula()
Dim cant As Currency
Dim entrega As Currency
entrega = 0
For fila = 1 To grid_lec.Cols - 1
 cant = Val(grid_lec.TextMatrix(2, fila)) - Val(grid_lec.TextMatrix(3, fila))
 If cant < 0 Then GoTo dale
 grid_lec.TextMatrix(4, fila) = Format(cant, "0.00")
 grid_lec.TextMatrix(6, fila) = Format(cant * Val(grid_lec.TextMatrix(5, fila)), "0.00")
 entrega = entrega + Val(grid_lec.TextMatrix(6, fila))
dale:
Next fila
t_totalgen.Text = Format(entrega, "0.00")
End Sub

Public Sub OBBLO()
t_grifero.Locked = True
t_grifero.BackColor = QBColor(7)
t_isla.Locked = True
t_isla.BackColor = QBColor(7)
t_turno.Locked = True
t_turno.BackColor = QBColor(7)

t_fecha.Enabled = False

t_efectivo.Locked = True
t_efectivo.BackColor = QBColor(7)

t_credito.Locked = True
t_credito.BackColor = QBColor(7)

t_desodcred.Locked = True
t_desodcred.BackColor = QBColor(7)

t_desfacred.Locked = True
t_desfacred.BackColor = QBColor(7)

t_desodcont.Locked = True
t_desodcont.BackColor = QBColor(7)

t_desfacont.Locked = True
t_desfacont.BackColor = QBColor(7)

t_cheques.Locked = True
t_cheques.BackColor = QBColor(7)

t_voucher.Locked = True
t_voucher.BackColor = QBColor(7)

t_vales.Locked = True
t_vales.BackColor = QBColor(7)

t_falsob.Locked = True
t_falsob.BackColor = QBColor(7)

t_entrega.Locked = True
t_entrega.BackColor = QBColor(7)

t_totalgen.Locked = True
t_totalgen.BackColor = QBColor(7)

t_efectivo.Locked = True
t_efectivo.BackColor = QBColor(7)
grid_lec.Enabled = False
cmdAct.Enabled = False
End Sub
Public Sub OBDES()
t_grifero.Locked = False
t_grifero.BackColor = QBColor(15)
t_isla.Locked = False
t_isla.BackColor = QBColor(15)
t_turno.Locked = False
t_turno.BackColor = QBColor(15)

t_fecha.Enabled = True

t_efectivo.Locked = False
t_efectivo.BackColor = QBColor(15)

t_credito.Locked = False
t_credito.BackColor = QBColor(15)

t_desodcred.Locked = False
t_desodcred.BackColor = QBColor(15)

t_desfacred.Locked = False
t_desfacred.BackColor = QBColor(15)

t_desodcont.Locked = False
t_desodcont.BackColor = QBColor(15)

t_desfacont.Locked = False
t_desfacont.BackColor = QBColor(15)

t_cheques.Locked = False
t_cheques.BackColor = QBColor(15)

t_voucher.Locked = False
t_voucher.BackColor = QBColor(15)

t_vales.Locked = False
t_vales.BackColor = QBColor(15)

t_falsob.Locked = False
t_falsob.BackColor = QBColor(15)

t_entrega.Locked = False
t_entrega.BackColor = QBColor(15)

t_totalgen.Locked = False
t_totalgen.BackColor = QBColor(15)

t_efectivo.Locked = False
t_efectivo.BackColor = QBColor(15)
grid_lec.Enabled = True
cmdAct.Enabled = True
End Sub



Public Sub GRABA_PAG()
If FLAG_NUEVO = "A" Then GoTo NUEVO_PARTE
pag_llave.Requery
If pag_llave.EOF Then
   Exit Sub
End If
t_fecha.Text = pag_llave!pag_fecha
Do Until pag_llave.EOF
  pag_llave.Edit
  pag_llave!PAG_LEC_CIERRE = Val(grid_lec.TextMatrix(2, pag_llave.AbsolutePosition))
  pag_llave!PAG_LEC_INICIO = Val(grid_lec.TextMatrix(3, pag_llave.AbsolutePosition))
  pag_llave!PAG_PRECIO = Val(grid_lec.TextMatrix(5, pag_llave.AbsolutePosition))
  pag_llave!PAG_EFECTIVO = Val(t_efectivo.Text)
  pag_llave!PAG_CREDITO = Val(t_credito.Text)
  pag_llave!PAG_DESODCRED = Val(t_desodcred.Text)
  pag_llave!PAG_DESFACRED = Val(t_desfacred.Text)
  pag_llave!PAG_DESODCONT = Val(t_desodcont.Text)
  pag_llave!PAG_DESFACONT = Val(t_desfacont.Text)
  pag_llave!PAG_CHEQUE = Val(t_cheques.Text)
  pag_llave!PAG_VALE = Val(t_vales.Text)
  pag_llave!PAG_VOUCHER = Val(t_voucher.Text)
  pag_llave!PAG_FALSOB = Val(t_falsob.Text)
  pag_llave!PAG_ENTREGA = Val(t_entrega.Text)
  pag_llave!PAG_TOTAL = Val(t_totalgen.Text)
  pag_llave.Update
  pag_llave.MoveNext
Loop
Exit Sub
NUEVO_PARTE:

PSMA_LLAVE(0) = LK_CODCIA
PSMA_LLAVE(1) = 2105
PSMA_LLAVE(2) = Val(t_isla.Text)
man_llave.Requery
If man_llave.EOF Then
  MsgBox "Debe Definir Mangueras para la Isla: " & t_isla.Text, 48, Pub_Titulo
  Exit Sub
End If
Do Until man_llave.EOF
  pag_llave.AddNew
  pag_llave!PAG_CODCIA = LK_CODCIA
  pag_llave!PAG_SURTIDOR = man_llave!TAB_NUMTAB
  pag_llave!pag_fecha = t_fecha.Text
  pag_llave!PAG_ISLA = Val(t_isla.Text)
  pag_llave!pag_turno = Val(t_turno.Text)
  pag_llave!PAG_CODRES = Val(t_grifero.Text)
  pag_llave!pag_numparte = Val(t_numfac.Text)
  pag_llave!PAG_CODART = Val(grid_lec.TextMatrix(7, man_llave.AbsolutePosition))
  pag_llave!PAG_LEC_CIERRE = Val(grid_lec.TextMatrix(2, man_llave.AbsolutePosition))
  pag_llave!PAG_LEC_INICIO = Val(grid_lec.TextMatrix(3, man_llave.AbsolutePosition))
  pag_llave!PAG_PRECIO = Val(grid_lec.TextMatrix(5, man_llave.AbsolutePosition))
  pag_llave!PAG_EFECTIVO = Val(t_efectivo.Text)
  pag_llave!PAG_CREDITO = Val(t_credito.Text)
  pag_llave!PAG_DESODCRED = Val(t_desodcred.Text)
  pag_llave!PAG_DESFACRED = Val(t_desfacred.Text)
  pag_llave!PAG_DESODCONT = Val(t_desodcont.Text)
  pag_llave!PAG_DESFACONT = Val(t_desfacont.Text)
  pag_llave!PAG_CHEQUE = Val(t_cheques.Text)
  pag_llave!PAG_VALE = Val(t_vales.Text)
  pag_llave!PAG_VOUCHER = Val(t_voucher.Text)
  pag_llave!PAG_FALSOB = Val(t_falsob.Text)
  pag_llave!PAG_ENTREGA = Val(t_entrega.Text)
  pag_llave!PAG_TOTAL = Val(t_totalgen.Text)
  pag_llave.Update
  man_llave.MoveNext
Loop

End Sub

Public Sub calcula2()
If Left(CmdModificar.Caption, 2) <> "&G" Then Exit Sub
t_efectivo.Text = Format(Val(t_totalgen.Text) - Val(t_credito.Text) - Val(t_desodcred.Text) - Val(t_desfacred.Text) - Val(t_desodcont.Text) - Val(t_desfacont.Text) - Val(t_cheques.Text) - Val(t_voucher.Text) - Val(t_vales.Text) + Val(t_falsob.Text), "0.00")
l_efectivo.Caption = "Efectivo -  " & t_efectivo.Text
t_efectivo.Text = Format(Val(t_efectivo.Text), "0.0")
't_entrega.Text = Format(Val(t_efectivo.Text) + Val(t_credito.Text) + Val(t_desodcred.Text) + Val(t_desfacred.Text) + Val(t_desodcont.Text) + Val(t_desfacont.Text) + Val(t_cheques.Text) + Val(t_voucher.Text) + Val(t_vales.Text), "0.00")
't_falsob.Text = Format(Val(t_efectivo.Text) - Val(t_totalgen.Text), "0.00")
If Val(t_falsob.Text) < 0 Then
  t_falsob.ForeColor = QBColor(12)
Else
  t_falsob.ForeColor = QBColor(0)
End If
End Sub
