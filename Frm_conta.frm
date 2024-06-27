VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FORM_CONTA 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Ingreso de Vouchers"
   ClientHeight    =   4890
   ClientLeft      =   3900
   ClientTop       =   3150
   ClientWidth     =   6600
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
   Icon            =   "Frm_Mayor.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4890
   ScaleWidth      =   6600
   Tag             =   "55"
   WindowState     =   2  'Maximized
   Begin MSMask.MaskEdBox i_fecha 
      Height          =   375
      Left            =   840
      TabIndex        =   22
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      _Version        =   327680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Frame MSTCOM 
      Caption         =   "Cuentas   ... "
      Height          =   3615
      Left            =   5040
      TabIndex        =   19
      Top             =   600
      Visible         =   0   'False
      Width           =   4335
      Begin VB.TextBox LEIDO 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   240
         TabIndex        =   20
         Top             =   240
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid grid1 
         Height          =   2775
         Left            =   240
         TabIndex        =   21
         Top             =   720
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   4895
         _Version        =   327680
         FixedCols       =   0
      End
   End
   Begin VB.TextBox i_voucher 
      Height          =   375
      Left            =   6000
      TabIndex        =   1
      Text            =   "i_voucher"
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox i_tipdoc 
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Text            =   "i_tipdoc"
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton SALIR 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5280
      TabIndex        =   12
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cancelar 
      BackColor       =   &H00000040&
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3840
      TabIndex        =   11
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton grabar 
      BackColor       =   &H00FF00FF&
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   2280
      TabIndex        =   10
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   5400
      Width           =   1095
   End
   Begin ComctlLib.ProgressBar Barra 
      Height          =   255
      Left            =   1800
      TabIndex        =   14
      Tag             =   "0"
      Top             =   4560
      Visible         =   0   'False
      Width           =   3675
      _ExtentX        =   6482
      _ExtentY        =   450
      _Version        =   327680
      Appearance      =   1
      MouseIcon       =   "Frm_Mayor.frx":0442
      Min             =   77
      Max             =   91
   End
   Begin VB.Frame ESTADO 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Estado :  "
      ForeColor       =   &H00000000&
      Height          =   4845
      Left            =   0
      TabIndex        =   7
      Tag             =   "100"
      Top             =   480
      Width           =   9495
      Begin VB.ComboBox i_d_h 
         Height          =   315
         ItemData        =   "Frm_Mayor.frx":045E
         Left            =   5640
         List            =   "Frm_Mayor.frx":0468
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3720
         Width           =   585
      End
      Begin VB.TextBox i_glosa 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "9999"
         Text            =   "i_glosa"
         Top             =   3720
         Width           =   3975
      End
      Begin VB.TextBox i_cuenta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         MaxLength       =   10
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "9999"
         Top             =   3720
         Width           =   1215
      End
      Begin VB.TextBox i_importe 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6360
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "9999"
         Top             =   3720
         Width           =   1215
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   5520
         TabIndex        =   25
         Top             =   4080
         Width           =   3855
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminación"
            Height          =   375
            Left            =   1920
            TabIndex        =   27
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton cmdActualizar 
            Caption         =   "&Modificar"
            Height          =   375
            Left            =   240
            TabIndex        =   26
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "&Consultar"
         Height          =   375
         Left            =   2040
         TabIndex        =   24
         Top             =   4440
         Width           =   1695
      End
      Begin VB.CommandButton cmdIngreso 
         Caption         =   "&Ingreso"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   4440
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid grid_fac 
         Height          =   3255
         Left            =   120
         TabIndex        =   6
         Tag             =   "9999"
         Top             =   240
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   5741
         _Version        =   327680
      End
      Begin VB.Label Lcuenta 
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   240
         TabIndex        =   29
         Top             =   4080
         Width           =   3735
      End
      Begin VB.Label Label1 
         Caption         =   "D/H"
         Height          =   255
         Left            =   5640
         TabIndex        =   15
         Tag             =   "9999"
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Glosa:"
         Height          =   195
         Left            =   2640
         TabIndex        =   13
         Tag             =   "9999"
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label Lprecio 
         Caption         =   "Importe:"
         Height          =   255
         Left            =   6360
         TabIndex        =   9
         Tag             =   "9999"
         Top             =   3480
         Width           =   855
      End
      Begin VB.Label LCODART 
         Caption         =   "Cuenta:"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Tag             =   "9999"
         Top             =   3480
         Width           =   645
      End
      Begin VB.Label momen 
         Caption         =   "Un Momento ..."
         Height          =   375
         Left            =   3360
         TabIndex        =   28
         Top             =   1560
         Width           =   1575
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Fecha:"
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   120
      Width           =   855
   End
   Begin VB.Label L2 
      Caption         =   "Voucher"
      Height          =   255
      Left            =   5160
      TabIndex        =   17
      Top             =   120
      Width           =   855
   End
   Begin VB.Label L1 
      Caption         =   "Tipo Doc:"
      Height          =   255
      Left            =   2760
      TabIndex        =   16
      Top             =   120
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   1
      Tag             =   "9999"
      X1              =   -480
      X2              =   9480
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   0
      Tag             =   "9999"
      X1              =   0
      X2              =   9480
      Y1              =   5400
      Y2              =   5400
   End
End
Attribute VB_Name = "FORM_CONTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public fila As Integer
Dim WPASA As Boolean
Dim WSELE As String * 1
Dim llave1
Dim filax As Integer
Dim ws_bruto_d, ws_bruto_h As Currency
Dim SUM_D As Currency
Dim SUM_H As Currency
Dim PSTEMP_LLAVE As rdoQuery
Dim temp_llave As rdoResultset
Dim WMODO As String * 1
Dim LOC_ITEM As Integer

Option Explicit


Private Sub cancelar_Click()
ESTADO.Caption = "Estado : "
'cancela_todo
filax = 0
SUM_D = 0
SUM_H = 0
LIMPIA_DATOS
FORM_CONTA.Lcuenta.Caption = ""
grid_fac.SetFocus
SendKeys "^{HOME}", True
'i_fecha.SetFocus
'CABE
End Sub


Private Sub cmdActualizar_Click()
Dim a
On Error GoTo SALE
If Left(cmdActualizar.Caption, 2) = "&M" Then
    ESTADO.Caption = "Estado :   < MODIFICANDO >"
    cmdActualizar.Caption = "&Actualizar"
    cmdEliminar.Enabled = False
    grid_fac.SetFocus
Else
  If WSELE = "N" Or Trim(WSELE) = "" Or LOC_ITEM = 0 Then
    MsgBox "Seleccione , para poder Actualizar..", 48, WS_TITULO
    Exit Sub
  End If
  'sn_mensaje = " ¿Confirma Actualizacion .. ?"
  'ws_respuesta = MsgBox(sn_mensaje, WS_ESTILO, WS_TITULO)
  'If ws_respuesta = vbNo Then
  '  GoTo fin
  'End If
  Barra.Visible = True
  Barra.Min = 0
  Barra.Max = 10
  Barra.Value = 1
  cov_llave.Edit
  Barra.Value = 2
  cov_llave!COV_CODCTA = i_cuenta.text
  grid_fac.TextMatrix(LOC_ITEM, 1) = i_cuenta.text
  Barra.Value = 3
  cov_llave!COV_GLOSA = i_glosa.text
  grid_fac.TextMatrix(LOC_ITEM, 3) = i_glosa.text
  Barra.Value = 4
  cov_llave!COV_IMPORTE = i_importe.text
  Barra.Value = 5
  cov_llave!COV_DH = i_d_h.text
  Barra.Value = 6
  cov_llave.Update
  If i_d_h.text = "D" Then
    grid_fac.TextMatrix(LOC_ITEM, 4) = i_importe.text
    grid_fac.TextMatrix(LOC_ITEM, 5) = ""
  Else
    grid_fac.TextMatrix(LOC_ITEM, 5) = i_importe.text
    grid_fac.TextMatrix(LOC_ITEM, 4) = ""
  End If
  Barra.Value = 7
'  cmdConsultar_Click
  i_cuenta.text = ""
  Lcuenta.Caption = ""
  i_glosa.text = ""
  i_importe.text = ""
  i_d_h.ListIndex = -1
  i_importe.text = ""
  grid_fac.Row = LOC_ITEM
  For a = 1 To grid_fac.Cols - 1
        grid_fac.Col = a
        grid_fac.CellBackColor = vbWhite
  Next a
  Barra.Value = 8
  'cmdActualizar.Caption = "&Modificar"
 ' cmdEliminar.Enabled = True
  Barra.Value = 10
  Barra.Visible = False
  grid_fac.Col = 1
  grid_fac.Row = LOC_ITEM
  LOC_ITEM = 0
  grid_fac.SetFocus
End If
Exit Sub
fin:
Exit Sub
SALE:
MsgBox Err.Description
End Sub

Private Sub cmdConsultar_Click()
ESTADO.Caption = "Estado :   < CONSULTANDO >"
WMODO = "C"
CABE_MAN
i_cuenta.text = ""
i_glosa.text = ""
i_importe.text = ""
i_d_h.ListIndex = -1

i_tipdoc.Locked = True
i_voucher.Locked = True
i_tipdoc.BackColor = QBColor(7)
i_voucher.BackColor = QBColor(7)
FORM_CONTA.i_fecha.Enabled = True
FORM_CONTA.i_d_h.Enabled = True
FORM_CONTA.i_tipdoc.Enabled = False
FORM_CONTA.i_voucher.Enabled = False
FORM_CONTA.i_cuenta.Enabled = True
FORM_CONTA.i_glosa.Enabled = True
FORM_CONTA.i_importe.Enabled = True

cmdIngreso.Enabled = False
cmdConsultar.Enabled = True
cmdActualizar.Caption = "&Modificar"
cmdActualizar.Enabled = False
cmdEliminar.Enabled = False
cmdEliminar.Caption = "&Eliminación"
grabar.Enabled = False
WPASA = False
i_fecha.SetFocus
End Sub

Private Sub cmdEliminar_Click()
If Left(cmdEliminar.Caption, 12) = "&Eliminación" Then
    WSELE = ""
    ESTADO.Caption = "Estado :   < ELIMINADO >"
    cmdEliminar.Caption = "&Eliminar Todas"
    cmdActualizar.Enabled = False
    grid_fac.SetFocus
Else
   Barra.Visible = True
   Barra.Min = 0
   Barra.Max = 10
   Barra.Value = 1
   WSELE = ""
   fila = 0
   Barra.Value = 3
   Do Until fila = grid_fac.Rows - 1
      fila = fila + 1
      grid_fac.Col = 6
      grid_fac.Row = fila
      If grid_fac.CellBackColor = vbRed Then
         PSCOV_LLAVE.rdoParameters(0) = LK_CODCIA
         PSCOV_LLAVE.rdoParameters(1) = Val(grid_fac.text)
         cov_llave.Requery
         If cov_llave.EOF Then
            MsgBox "ERROR GRAVE..."
         End If
      cov_llave.Delete
      WSELE = "S"
      End If
  Loop
  Barra.Value = 6
  If WSELE = "N" Or Trim(WSELE) = "" Then
    MsgBox "Seleccione  , para poder Eliminar..", 48, WS_TITULO
    Exit Sub
  End If
  Barra.Value = 7
  'ESTADO.Caption = "Estado :   < ELIMINADO >"
  cmdEliminar.Caption = "&Eliminación"
  cmdActualizar.Enabled = False
  cmdEliminar.Enabled = False
  cmdConsultar_Click
  Barra.Value = 10
  Barra.Visible = False
  
End If
End Sub

Private Sub cmdIngreso_Click()
ESTADO.Caption = "Estado :   < Ingreso de VOUCHER >"
WMODO = "I"
i_tipdoc.Locked = False
i_voucher.Locked = False
i_tipdoc.BackColor = QBColor(15)
i_voucher.BackColor = QBColor(15)
FORM_CONTA.i_fecha.Enabled = True
FORM_CONTA.i_d_h.Enabled = True
FORM_CONTA.i_tipdoc.Enabled = True
FORM_CONTA.i_voucher.Enabled = True
FORM_CONTA.i_cuenta.Enabled = True
FORM_CONTA.i_glosa.Enabled = True
FORM_CONTA.i_importe.Enabled = True

cmdConsultar.Enabled = False
cmdActualizar.Enabled = False
cmdEliminar.Enabled = False
cmdConsultar.Enabled = False
grabar.Enabled = False
WPASA = True
CABE_ING
filax = 0
SUM_D = 0
SUM_H = 0

i_fecha.SetFocus
End Sub



Private Sub Form_Load()
filax = 0
grid1.Rows = 50
grid1.Cols = 3
grid1.Rows = 2
grid1.ColAlignment(0) = 1
grid1.ColAlignment(1) = 1
grid1.ColAlignment(2) = 1
grid1.ColWidth(0) = 1300
grid1.ColWidth(1) = 2300
grid1.ColWidth(2) = 1

WSELE = ""
Dim ws_indice As Integer
Dim cade
WMODO = ""
cade = "SELECT * FROM COMOV WHERE COV_CODCIA = ?  ORDER BY COV_NRO_MOV "
Set PSTEMP_LLAVE = CN.CreateQuery("", cade)
Set temp_llave = PSTEMP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

Dim cadena As String
'Dim ws_indice As Integer

cadena = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND COV_FECHA_voucher >=?    ORDER BY COV_CODCIA, COV_FECHA_voucher, COV_NRO_MOV"
Set PSCOV_MAYOR = CN.CreateQuery("", cadena)
Set cov_mayor = PSCOV_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)

cadena = "SELECT * FROM COMOV WHERE COV_CODCIA = ?  AND COV_NRO_MOV=?   ORDER BY COV_CODCIA, COV_FECHA_voucher, COV_NRO_MOV"
Set PSCOV_LLAVE = CN.CreateQuery("", cadena)
Set cov_llave = PSCOV_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)


filax = 0
DoEvents
LIMPIA_DATOS

i_fecha.Mask = "##/##/##"
i_fecha.text = Format(LK_FECHA_DIA, "dd/mm/yy")
End Sub

Private Sub Form_Terminate()
'FORMGEN.TRANS.SetFocus
End Sub

Private Sub grabar_Click()

'On Error GoTo Error_fatal
Dim er As rdoError
Dim msg As String
Const ingre = 2
Const modif = 1
Dim N As Integer
Dim LOC_SALDO_CAR As Currency
Dim WS_DESCTO1, WS_DESCTO2 As Currency
Dim WS_CANTIDAD As Currency
Dim WS_NUMSEC As Integer
Dim count As Integer
Dim contador As Integer
Dim cadena As String
Dim zona_nombre As String
Dim subzona_nombre  As String
Dim WS_IMPRESION_LET As Long
Dim WS_CHENUM As Currency
Dim WS_NUM_OPER  As Integer
Dim ws_tot_debe As Currency
Dim ws_tot_haber As Currency
Dim WS_DATOS As String * 2
Dim WS_FLAG As Integer
Dim WS_TRANSITO As String * 1
Dim WS_FLAG2 As Integer
Dim FLAG As Boolean
Dim msg_err As String
Dim WS_NRO_MOV As Integer
Dim w_dh  As String

If Not IsDate(i_fecha.text) Then
  MsgBox "Fecha Erreda... Verificar .! ", 48, WS_TITULO
  i_fecha.SetFocus
  GoTo fin
End If
If Trim(i_tipdoc.text) = "" Then
  MsgBox "Ingrese Texto al Tipo de Documento...", 48, WS_TITULO
  i_tipdoc.SetFocus
  GoTo fin
End If
If Trim(i_voucher.text) = "" Or Val(i_voucher.text) = 0 Then
  MsgBox "Ingrese Nro. de  VOUCHER ", 48, WS_TITULO
  i_voucher.SetFocus
  GoTo fin
End If

FLAG = False
grid_fac.Col = 1
fila = 1
ws_tot_haber = 0
ws_tot_debe = 0
FORM_CONTA.grid_fac.Row = fila
Do While FLAG = False
   If fila = 1 And Trim(grid_fac.TextMatrix(fila, 0)) = "" Then
     MsgBox "INGRESE VOUCHERS ...."
     i_fecha.SetFocus
     GoTo fin
   End If
   grid_fac.Row = fila
   If grid_fac.CellBackColor = vbScrollBars Then
     GoTo wotro
   End If
   If Trim(grid_fac.TextMatrix(fila, 4)) <> "" Then
        w_dh = "D"
        PUB_IMPORTE = grid_fac.TextMatrix(fila, 4)
        ws_tot_debe = ws_tot_debe + PUB_IMPORTE
   ElseIf Trim(grid_fac.TextMatrix(fila, 5)) <> "" Then
        w_dh = "H"
        PUB_IMPORTE = grid_fac.TextMatrix(fila, 5)
        ws_tot_haber = ws_tot_haber + PUB_IMPORTE
   End If
wotro:
   fila = fila + 1
   FORM_CONTA.grid_fac.Row = fila
   If Left(FORM_CONTA.grid_fac.TextMatrix(fila + 1, 3), 1) = "F" Then
      FLAG = True
   End If

Loop
If ws_tot_debe = 0 And ws_tot_haber = 0 Then
  MsgBox "ingrese los Nuevos Items ..", 48, WS_TITULO
  Barra.Visible = False
  i_cuenta.SetFocus
  GoTo fin
End If
If ws_tot_debe <> ws_tot_haber Then
   MsgBox " Revisa...NO CUADRA  == >  Debe  : " & Format(ws_tot_debe, "##,###0.00") & "   Haber : " & Format(ws_tot_haber, "##,###0.00"), 48, WS_TITULO
   Barra.Visible = False
   i_cuenta.SetFocus
   GoTo fin
End If

Screen.MousePointer = 11
DoEvents
Barra.Visible = True
DoEvents
Barra.Min = 0
Barra.Max = 10
exito = True
Barra.Value = 2

GoSub ACT1
filax = 0
SUM_D = 0
SUM_H = 0
Barra.Value = 10

FORM_CONTA.i_d_h.ListIndex = -1
FORM_CONTA.i_tipdoc = ""
FORM_CONTA.i_voucher = ""
FORM_CONTA.grid_fac.Clear
FORM_CONTA.i_cuenta = ""
FORM_CONTA.i_glosa = ""
FORM_CONTA.i_importe = ""
WPASA = True
CABE_ING
grid_fac.SetFocus
SendKeys "^{HOME}", True
cmdIngreso.SetFocus
If exito = False Then
   Barra.Visible = False
   MsgBox msg_err
    GoTo errorr
End If
Barra.Value = 10
N = 105
i_importe.text = ""
Barra.Visible = False
GoTo fin

ACT1:

PSTEMP_LLAVE(0) = LK_CODCIA
temp_llave.Requery
If temp_llave.EOF Then
 WS_NRO_MOV = 0
Else
 temp_llave.MoveLast
 WS_NRO_MOV = temp_llave!COV_NRO_MOV
End If
fila = 1
WS_NRO_MOV = WS_NRO_MOV + 1
PUB_NUMSER = Val(i_voucher.text)
PUB_FECHA = i_fecha.text
FLAG = False
Do While FLAG = False
   FORM_CONTA.grid_fac.Row = fila
   If grid_fac.CellBackColor = vbScrollBars Then
     GoTo wotro2
   End If
   PUB_CONCEPTO = Trim(grid_fac.TextMatrix(fila, 3))
   PUB_CUENTA = Trim(grid_fac.TextMatrix(fila, 1))
   If Trim(grid_fac.TextMatrix(fila, 4)) <> "" Then
        w_dh = "D"
        PUB_IMPORTE = grid_fac.TextMatrix(fila, 4)
   ElseIf Trim(grid_fac.TextMatrix(fila, 5)) <> "" Then
        w_dh = "H"
        PUB_IMPORTE = grid_fac.TextMatrix(fila, 5)
   End If
       
    ' grabo todo
   cov_llave.AddNew
   cov_llave!COV_NRO_MOV = WS_NRO_MOV
   cov_llave!COV_CODCIA = LK_CODCIA
   cov_llave!cov_nro_voucher = PUB_NUMSER
   cov_llave!cov_fecha_voucher = PUB_FECHA
   cov_llave!COV_GLOSA = PUB_CONCEPTO
   cov_llave!COV_FECHA_PROC = LK_FECHA_DIA
   cov_llave!COV_CODCTA = PUB_CUENTA
   cov_llave!COV_DH = w_dh
   cov_llave!COV_IMPORTE = PUB_IMPORTE
   cov_llave!COV_ESTADO = "0"
   cov_llave!COV_CODUSU = LK_CODUSU
   cov_llave!COV_FLAG_AUTOMATICA = "0"
   cov_llave.Update
wotro2:
   fila = fila + 1
   WS_NRO_MOV = WS_NRO_MOV + 1
   FORM_CONTA.grid_fac.Row = fila
   If Left(FORM_CONTA.grid_fac.TextMatrix(fila + 1, 3), 1) = "F" Then
      FLAG = True
   End If
   
Loop
Return

ACT7:
all_menor.AddNew
all_menor!all_numoper = PUB_NUM_OPER_XXX

'For Each Control In FORMGEN.Controls
'    If Control.Tag < 100 And Control.Tag > 0 And Control.Visible = True Then
'       contador = Control.Tag
'       msg_err =  all.Recordset.Fields(contador)
'       all(contador) = Control.text
'    End If
'Next Control

all_menor!ALL_CODCIA = LK_CODCIA
all_menor!all_codtra = LK_CODTRA
all_menor!all_flag_ext = Nulo_Valors(tra_llave!TRA_FLAG_EXT)
all_menor!all_codclie = PUB_CODCLIE
all_menor!ALL_CODART = PUB_CODART
all_menor!all_importe_amort = PUB_IMPORTE_AMORT
all_menor!ALL_INTADE = PUB_INTADE
all_menor!ALL_INTVEN = PUB_INTVEN
all_menor!all_codusu = LK_CODUSU
all_menor!all_FBG = PUB_FBG
all_menor!ALL_CODVEN = PUB_CODVEN
all_menor!all_importe = PUB_IMPORTE
all_menor!ALL_GASTOS_FIJOS = PUB_GASTOS_FIJOS
all_menor!all_codcia_r = PUB_CODCIA_R
all_menor!ALL_NUMDOC_R = PUB_NUMDOC_R
all_menor!ALL_NUMDOC = PUB_NUMDOC
all_menor!all_cp = PUB_CP
all_menor!all_tipdoc = PUB_TIPDOC
all_menor!ALL_SALDO_CAR = LOC_SALDO_CAR
all_menor!ALL_NUMFAC_C = PUB_NUMFAC_C
all_menor!ALL_NUMSER_C = PUB_NUMSER_C
all_menor!all_codban = PUB_CODBAN
all_menor!all_concepto = PUB_CONCEPTO
all_menor!ALL_CHENUM = PUB_CHENUM
all_menor!ALL_FECHA_DIA = LK_FECHA_DIA '**
all_menor!ALL_FECHA_VCTO = PUB_FECHA_VCTO '**
all_menor!ALL_CANTIDAD = PUB_CANTIDAD
all_menor!all_numser = PUB_NUMSER
all_menor!all_numfac = PUB_NUMFAC
all_menor!ALL_NETO = PUB_NETO  '**IMPORTE
all_menor!ALL_BRUTO = PUB_SUBTOTAL
all_menor!ALL_TIPMOV = PUB_TIPMOV
all_menor!ALL_IMPTO = PUB_IMPTO
all_menor!ALL_DESCTO = PUB_DESCTO
all_menor!ALL_INTERES = PUB_INTERES
all_menor!ALL_PORDESCTO1 = PUB_PORDESCTO1
all_menor!ALL_PORDESCTO2 = PUB_PORDESCTO2
all_menor!ALL_GASTOS = PUB_GASTOS
all_menor!ALL_PRECIO = PUB_PRECIO2
all_menor!ALL_SITUACION_ACT = PUB_SITUACION_ACT
all_menor!ALL_SITUACION_ANT = PUB_SITUACION_ANT
all_menor!ALL_GRUPO_ANT = PUB_GRUPO_ANT
all_menor!ALL_GRUPO_ACT = PUB_GRUPO_ACT

all_menor!ALL_DIAS_A = pub_diasA
all_menor!ALL_DIAS_V = pub_diasV
all_menor!ALL_FLAG_SUSPE = PUB_FLAG_SUSPE
all_menor!ALL_PARCIAL = PUB_PARCIAL
all_menor!ALL_UNIFICADA = PUB_UNIFICADA
all_menor!all_secuencia = PUB_SECUENCIA
all_menor!all_SIGNO_CAR = pub_signo_car
all_menor!ALL_SIGNO_CAJA = pub_signo_caja
all_menor!ALL_SIGNO_LOT = pub_signo_lot

all_menor!ALL_SIGNO_CCM = pub_signo_ccm
all_menor!all_sIGNO_ARM = pub_signo_arm

all_menor!ALL_TASA_VENTA = PUB_TASA_VENTA
all_menor!ALL_NUM_INI = PUB_NUM_INI
all_menor!ALL_CHENUM = PUB_CHENUM
all_menor!ALL_CHESEC = PUB_CHESEC
all_menor!ALL_CHESER = PUB_CHESER

'all_menor!ALL_FLAG_ARM = ws_flag_arm
'all_menor!ALL_FLAG_che = ws_flag_che
'all_menor!ALL_FLAG_FAR = ws_flag_far
'all_menor!ALL_FLAG_CAR = ws_flag_car
'all_menor!ALL_FLAG_PAR = ws_flag_par
'all_menor!ALL_FLAG_CCM = ws_flag_ccm
'all_menor!ALL_FLAG_LOT = ws_flag_lot

'all_menor!all_codtra_ext = Left(FORMGEN.TRANS.text, 4)
'all_menor!ALL_CODALI = PUB_CODALI
'all_menor!ALL_FECHA_LOTE = PUB_FECHA_LOTE
'all_menor!ALL_NUM_LOTE = PUB_NUM_OPER_XXX


all_menor.Update

Return



  

Screen.MousePointer = 1

GoTo fin



Exit Sub
Error_fatal:
    msg = "Se ha producido un error " & "al abrir la conexión:" & Err & " - " & Error & vbCr
    For Each er In rdoErrors
        msg = msg & er.Description & ":" & er.Number & vbCr
        MsgBox msg
    Next er

    CN.Execute "Rollback Transaction", rdExecDirect

    
'    Resume AbandonCn
Exit Sub
errorr:
 MsgBox msg_err, 48, WS_TITULO

fin:
Screen.MousePointer = 0


End Sub

Private Sub grid_fac_KeyPress(KeyAscii As Integer)
Dim a As Integer
Dim t, wc
Static CONS
If KeyAscii <> 13 Then
   GoTo fin
End If
If Trim(grid_fac.TextMatrix(grid_fac.Row, 0)) = "totales" Or Trim(grid_fac.TextMatrix(grid_fac.Row, 0)) = "" Then
 Exit Sub
End If
If Left(cmdActualizar.Caption, 2) <> "&A" Then
 Exit Sub
End If

fila = grid_fac.Row
wc = grid_fac.Col
i_cuenta.text = grid_fac.TextMatrix(fila, 1)
i_glosa.text = grid_fac.TextMatrix(fila, 3)
If grid_fac.TextMatrix(fila, 4) <> "" Then
   i_d_h.ListIndex = 0
   i_importe.text = grid_fac.TextMatrix(fila, 4)
Else
   i_d_h.ListIndex = 1
   i_importe.text = grid_fac.TextMatrix(fila, 5)
End If
   If CONS <> 0 Then
     grid_fac.Row = CONS
     For a = 1 To grid_fac.Cols - 1
       grid_fac.Col = a
       grid_fac.CellBackColor = vbWhite
     Next a
   End If
  grid_fac.Row = fila
  LOC_ITEM = fila
   If grid_fac.CellBackColor = vbScrollBars Then
      For a = 1 To grid_fac.Cols - 1
        grid_fac.Col = a
        grid_fac.CellBackColor = vbWhite
      Next a
      WSELE = "N"
    Else
      For a = 1 To grid_fac.Cols - 1
        grid_fac.Col = a
        grid_fac.CellBackColor = vbScrollBars
      Next a
      WSELE = "S"
    End If
 grid_fac.Col = wc
 CONS = grid_fac.Row
 PSCOV_LLAVE.rdoParameters(0) = LK_CODCIA
 PSCOV_LLAVE.rdoParameters(1) = grid_fac.TextMatrix(fila, 6)
 cov_llave.Requery
 i_cuenta.SetFocus

fin:
End Sub

Private Sub grid_fac_KeyUp(KeyCode As Integer, Shift As Integer)
Dim wc
Dim a, WF As Integer
Dim tf, t, tC
Dim SALE As Boolean
tC = grid_fac.Col
tf = grid_fac.Row
If KeyCode = 46 Then
If Trim(grid_fac.TextMatrix(tf, 0)) = "" Then
  Exit Sub
End If
grid_fac.Col = 1
If grid_fac.CellBackColor = vbScrollBars Then
  grid_fac.Col = tC
  grid_fac.Row = tf
  Exit Sub
End If

 If cmdActualizar.Enabled = False And cmdEliminar.Enabled = False Then
     WF = 1
     FORM_CONTA.grid_fac.Row = WF
     SALE = False
     Do While SALE = False
       If WF = 1 And Trim(grid_fac.TextMatrix(WF, 0)) = "" Then
         Exit Sub
       End If
        WF = WF + 1
        FORM_CONTA.grid_fac.Row = WF
        If Left(FORM_CONTA.grid_fac.TextMatrix(WF + 1, 3), 1) = "F" Then
            If Trim(grid_fac.TextMatrix(tf, 4)) <> "" Then
                PUB_IMPORTE = grid_fac.TextMatrix(tf, 4)
                SUM_D = SUM_D - PUB_IMPORTE
                grid_fac.TextMatrix(WF + 1, 4) = Format(SUM_D, "###,##0.00")
            End If
            If Trim(grid_fac.TextMatrix(tf, 5)) <> "" Then
                PUB_IMPORTE = grid_fac.TextMatrix(tf, 5)
                SUM_H = SUM_H - PUB_IMPORTE
                grid_fac.TextMatrix(WF + 1, 5) = Format(SUM_H, "###,##0.00")
            End If
            SALE = True
        End If
     Loop
     grid_fac.Row = tf
     For a = 1 To grid_fac.Cols - 1
        grid_fac.Col = a
        grid_fac.CellBackColor = vbScrollBars
     Next a
     grid_fac.Col = tC
     grid_fac.Row = tf
 End If
End If
If KeyCode <> 32 Then
   GoTo fin
End If
If Trim(grid_fac.TextMatrix(grid_fac.Row, 0)) = "totales" Then
 Exit Sub
End If
If Trim(grid_fac.TextMatrix(grid_fac.Row, 0)) = "" Then
 Exit Sub
End If
If Left(cmdEliminar.Caption, 9) <> "&Eliminar" Then
 Exit Sub
End If

t = grid_fac.Row
wc = grid_fac.Col
    If fila <> 0 Then
       For a = 1 To grid_fac.Cols - 1
        grid_fac.Col = a
        grid_fac.Row = fila
        grid_fac.CellBackColor = vbWhite
       Next a
       fila = 0
    End If
    grid_fac.Row = t
    If grid_fac.CellBackColor = vbRed Then
      For a = 1 To grid_fac.Cols - 1
        grid_fac.Col = a
        grid_fac.CellBackColor = vbWhite
      Next a
    Else
      For a = 1 To grid_fac.Cols - 1
        grid_fac.Col = a
        grid_fac.CellBackColor = vbRed
      Next a
    End If
grid_fac.Col = wc
fin:
End Sub

Private Sub GRID1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    MSTCOM.Visible = False
    i_cuenta.SetFocus
    KeyCode = 0
    Exit Sub
End If
If KeyCode = 13 Then
    If Left(grid1.TextMatrix(grid1.Row, 0), 1) = "" Then
        Exit Sub
    End If
    i_cuenta.text = Trim(grid1.TextMatrix(grid1.Row, 0))
    Lcuenta.Caption = Trim(grid1.TextMatrix(grid1.Row, 1))
    SQ_OPER = 1
    PUB_CUENTA = i_cuenta.text
    PUB_CODCIA = LK_CODCIA
    LEER_COM_LLAVE

    MSTCOM.Visible = False
    i_glosa.SetFocus
End If


End Sub


Private Sub i_cuenta_GotFocus()
Azul i_cuenta, i_cuenta
Lcuenta.Caption = ""
End Sub

Private Sub i_d_h_KeyPress(KeyAscii As Integer)
Azul i_importe, i_importe
End Sub

Private Sub i_fecha_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
  GoTo fin
End If
If IsDate(i_fecha.text) = False Then
   MsgBox "Fecha Invalida ...", 48, WS_TITULO
   i_fecha.SetFocus
   i_fecha.SelStart = 0
   i_fecha.SelLength = Len(i_fecha.text)
   GoTo fin
End If
If CDate(i_fecha.text) < 1900 Then
   MsgBox "Fecha NO Procede ...", 48, WS_TITULO
   i_fecha.SetFocus
   i_fecha.SelStart = 0
   i_fecha.SelLength = Len(i_fecha.text)
   GoTo fin
End If
If Left(cmdActualizar.Caption, 2) = "&A" Then
   MsgBox "No procede, Debes cambiar de Estado a CONSULTANDO ", 48, WS_TITULO
   grid_fac.SetFocus
   Exit Sub
End If
If Left(cmdEliminar.Caption, 9) = "&Eliminar" Then
   MsgBox "No procede, Debes cambiar de Estado a CONSULTANDO ", 48, WS_TITULO
   grid_fac.SetFocus
   Exit Sub
End If

If WMODO = "C" Then
 GoTo SIGUE
End If
 i_tipdoc.SetFocus
 Exit Sub

SIGUE:
' ' OTRO CASO
Dim WS_SALDO As Currency
Dim Tit As String
Dim i As Integer
Dim success%
Dim con_cuenta As String * 1
If KeyAscii <> 13 Then
  GoTo fin
End If
con_cuenta = "N"
If Trim(i_cuenta.text) <> "" Then
 SQ_OPER = 1
 PUB_CUENTA = i_cuenta.text
 PUB_CODCIA = LK_CODCIA
 LEER_COM_LLAVE
 If com_llave.EOF Then
   MsgBox "Su Codigo de Cuenta NO Existe ...", 48, WS_TITULO
   i_cuenta.SetFocus
 End If
 con_cuenta = "S"
End If

Screen.MousePointer = 11
grid_fac.Visible = False
DoEvents
'success% = SetWindowPos(FrmProcesa.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
DoEvents
Dim ws_tot_debe, ws_tot_haber As Currency
Dim ws_fecha  As String
Dim ws_voucher As Currency

PSCOV_MAYOR.rdoParameters(0) = LK_CODCIA
PSCOV_MAYOR.rdoParameters(1) = i_fecha.text
cov_mayor.Requery
fila = 0
cov_mayor.MoveFirst
If cov_mayor.EOF = True Then
   FrmProcesa.Hide
   MsgBox "No hay Registros...."
   Screen.MousePointer = 0
   grid_fac.Visible = True
   GoTo fin
End If
ws_tot_debe = 0
ws_tot_haber = 0
Tit = grid_fac.FormatString
'grid_fac.Clear
grid_fac.FormatString = Tit
ws_fecha = cov_mayor!cov_fecha_voucher
ws_voucher = cov_mayor!cov_nro_voucher
Do Until cov_mayor.EOF
       If con_cuenta = "N" Then
        If ws_fecha <> cov_mayor!cov_fecha_voucher Or ws_voucher <> cov_mayor!cov_nro_voucher Then
          fila = fila + 1
          grid_fac.Rows = fila + 1
          grid_fac.TextMatrix(fila, 0) = "totales "
          grid_fac.TextMatrix(fila, 1) = ""
          grid_fac.TextMatrix(fila, 2) = ""
          grid_fac.TextMatrix(fila, 3) = ""
          grid_fac.TextMatrix(fila, 4) = Format(ws_tot_debe, "##,##0.00")
          grid_fac.TextMatrix(fila, 5) = Format(ws_tot_haber, "##,##0.00")
          ws_tot_debe = 0
          ws_tot_haber = 0
        End If
       End If
       If Trim(cov_mayor!COV_CODCTA) <> Trim(i_cuenta.text) And con_cuenta = "S" Then
         GoTo otro
       End If
       fila = fila + 1
       grid_fac.Rows = fila + 1
       grid_fac.Row = fila
       grid_fac.TextMatrix(fila, 0) = Format(cov_mayor!cov_fecha_voucher, "dd/mm/yy")
       grid_fac.TextMatrix(fila, 1) = Trim(cov_mayor!COV_CODCTA)
       grid_fac.TextMatrix(fila, 2) = cov_mayor!cov_nro_voucher
       grid_fac.TextMatrix(fila, 3) = Trim(cov_mayor!COV_GLOSA)
       If cov_mayor!COV_DH = "D" Then
          grid_fac.TextMatrix(fila, 4) = Nulo_Valor0(cov_mayor!COV_IMPORTE)
          grid_fac.TextMatrix(fila, 5) = ""
          ws_tot_debe = ws_tot_debe + Nulo_Valor0(cov_mayor!COV_IMPORTE)
       ElseIf cov_mayor!COV_DH = "H" Then
             grid_fac.TextMatrix(fila, 4) = ""
             grid_fac.TextMatrix(fila, 5) = Nulo_Valor0(cov_mayor!COV_IMPORTE)
             ws_tot_haber = ws_tot_haber + Nulo_Valor0(cov_mayor!COV_IMPORTE)
       End If
       grid_fac.TextMatrix(fila, 6) = cov_mayor!COV_NRO_MOV
       
otro:
        ws_fecha = cov_mayor!cov_fecha_voucher
        ws_voucher = cov_mayor!cov_nro_voucher
        cov_mayor.MoveNext
Loop
 fila = fila + 1
 grid_fac.Rows = fila + 1
 grid_fac.Row = fila
 grid_fac.TextMatrix(fila, 0) = "totales "
 grid_fac.TextMatrix(fila, 1) = ""
 grid_fac.TextMatrix(fila, 2) = ""
 grid_fac.TextMatrix(fila, 3) = ""
 grid_fac.TextMatrix(fila, 4) = Format(ws_tot_debe, "##,##0.00")
 grid_fac.TextMatrix(fila, 5) = Format(ws_tot_haber, "##,##0.00")
 ws_tot_debe = 0
 ws_tot_haber = 0
 cmdActualizar.Enabled = True
 cmdEliminar.Enabled = True
 grid_fac.Visible = True
 grid_fac.Col = 1
 grid_fac.Row = 1
 grid_fac.SetFocus
 Screen.MousePointer = 0

fin:

End Sub

Private Sub i_glosa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    i_d_h.SetFocus
    SendKeys "%{UP}"
End If

End Sub

Private Sub i_importe_KeyPress(KeyAscii As Integer)
Dim valor As Currency
Dim subtotal As Currency
Dim tf As Integer
If KeyAscii <> 13 Then
   GoTo fin
End If
If WPASA = False Then
 If cmdActualizar.Enabled Then
   cmdActualizar.SetFocus
 End If
 
 Exit Sub
End If
If WMODO = "C" Then
 Exit Sub
End If
grabar.Enabled = True
If com_llave.EOF = True Then
   MsgBox "Primero seleccione Cuenta... ", 48, WS_TITULO
   GoTo fin
Else
   If com_llave!COM_CUENTA <> Val(i_cuenta.text) Then
   MsgBox "Primero seleccione Articulo.. ", 48, WS_TITULO
   GoTo fin
   End If
End If
If Not decimal1(i_importe.text) Then
   Azul i_importe, i_importe
   GoTo fin
End If
If i_importe.text = "" Then
   i_importe.SetFocus
   GoTo fin
End If
If i_d_h.text = "" Then
   i_d_h.SetFocus
   SendKeys "%{UP}"
   GoTo fin
End If

If numfilas = -1 Then
   fila = 0
End If

salta:
   
'FORMGEN.Frame1.Visible = False
'   CABE_ING
   filax = filax + 1
   grid_fac.Row = filax
   grid_fac.Col = 0
   grid_fac.CellAlignment = 7
   grid_fac.text = filax
   grid_fac.Col = 1
   grid_fac.CellAlignment = 1
   grid_fac.text = com_llave!COM_CUENTA
   grid_fac.Col = 2
   grid_fac.text = com_llave!COM_DESCRIPCION
   grid_fac.Col = 3
   grid_fac.text = i_glosa.text
   If Trim(i_d_h.text) = "D" Then
     grid_fac.Col = 5
     grid_fac.text = ""
     grid_fac.Col = 4
     grid_fac.text = ""
     grid_fac.CellAlignment = 7
     grid_fac.text = Format(i_importe.text, "##,###,##0.00")
     SUM_D = SUM_D + Val(i_importe.text)
   ElseIf Trim(i_d_h.text) = "H" Then
      grid_fac.Col = 4
      grid_fac.text = ""
      grid_fac.Col = 5
      grid_fac.text = ""
      grid_fac.CellAlignment = 7
      grid_fac.text = Format(i_importe.text, "##,###,##0.00")
      SUM_H = SUM_H + Val(i_importe.text)
   End If
  
   i_importe.text = ""
   i_d_h.ListIndex = -1
   grid_fac.Row = filax + 1
   grid_fac.TextMatrix(filax + 1, 3) = ""
   grid_fac.TextMatrix(filax + 1, 4) = ""
   grid_fac.TextMatrix(filax + 1, 5) = ""
   grid_fac.Row = filax + 2
   grid_fac.Col = 3
   grid_fac.text = "FIN  == > "
   grid_fac.Col = 4
   grid_fac.CellAlignment = 7
   grid_fac.text = Format(SUM_D, "##,##0.00")
   grid_fac.Col = 5
   grid_fac.CellAlignment = 7
   grid_fac.text = Format(SUM_H, "##,##0.00")
   If filax > 10 Then
      grid_fac.SetFocus
      SendKeys "{HOME}", True
      SendKeys "{DOWN}", True
      SendKeys "{UP 6}", True
  End If
 i_cuenta.SetFocus
'If Val(I_BRUTO_D.text) = Val(I_BRUTO_H.text) Then
'   MsgBox "*******  OK   ********", 48, WS_TITULO
'   grabar.SetFocus
'End If
fin:


End Sub


Private Sub i_tipdoc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If
i_voucher.SetFocus

fin:

End Sub

Private Sub i_cuenta_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer

If KeyAscii <> 13 Then
   GoTo fin
End If

            
 If i_cuenta.text = "" Then
  llave1 = ""
  LEIDO.text = ""
  Call PRO_CUENTAS("")
  MSTCOM.Visible = True
  FORM_CONTA.LEIDO.SetFocus
  Exit Sub
 Else
    SQ_OPER = 1
    PUB_CUENTA = i_cuenta.text
    PUB_CODCIA = LK_CODCIA
    LEER_COM_LLAVE
    If com_llave.EOF Then
       MsgBox "CUENTA NO EXISTE ...", 48, WS_TITULO
       i_cuenta.SetFocus
       GoTo fin
    Else
       Lcuenta.Caption = Trim(com_llave!COM_DESCRIPCION)
       i_glosa.SetFocus
    End If
    
 End If
i_glosa.SetFocus
 

fin:
End Sub



Private Sub i_voucher_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 i_cuenta.SetFocus
End If
End Sub

Private Sub LEIDO_Change()
Call PRO_CUENTAS("")
grid1.Col = 0
grid1.Row = 1
grid1.ColSel = grid1.Cols - 1
grid1.RowSel = 1

End Sub

Private Sub LEIDO_GotFocus()
grid1.Col = 0
grid1.Row = 1
grid1.ColSel = grid1.Cols - 1
grid1.RowSel = 1


End Sub

Private Sub LEIDO_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    MSTCOM.Visible = False
    ' 0
    i_cuenta.SetFocus
    Exit Sub
End If
If KeyAscii = 13 Then
   GRID1_KeyUp 13, 0
End If

End Sub

Private Sub LEIDO_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    grid1.Col = 1
    grid1.Row = 1
    grid1.SetFocus
End If


End Sub

Private Sub salir_Click()
Unload FORM_CONTA
End Sub


Public Sub LIMPIA_DATOS()
i_tipdoc.BackColor = QBColor(15)
i_voucher.BackColor = QBColor(15)
'FORM_CONTA.i_fecha = ""
cmdActualizar.Caption = "&Modificar"
cmdEliminar.Caption = "&Eliminacion"
i_fecha.text = Format(LK_FECHA_DIA, "dd/mm/yy")
FORM_CONTA.i_fecha.Enabled = False
FORM_CONTA.i_d_h.ListIndex = -1
FORM_CONTA.i_d_h.Enabled = False
FORM_CONTA.i_tipdoc = ""
FORM_CONTA.i_tipdoc.Enabled = False
FORM_CONTA.i_voucher = ""
FORM_CONTA.i_voucher.Enabled = False
FORM_CONTA.grid_fac.Clear
FORM_CONTA.i_cuenta = ""
FORM_CONTA.i_cuenta.Enabled = False
FORM_CONTA.i_glosa = ""
FORM_CONTA.i_glosa.Enabled = False
FORM_CONTA.i_importe = ""
FORM_CONTA.i_importe.Enabled = False
cmdActualizar.Enabled = False
cmdEliminar.Enabled = False
grabar.Enabled = False
cmdIngreso.Enabled = True
cmdConsultar.Enabled = True
WPASA = False
End Sub

Public Sub CABE_MAN()
grid_fac.FontName = "Sans Serif"
grid_fac.FontSize = 7
grid_fac.Cols = 7
grid_fac.Rows = 100
grid_fac.Clear
grid_fac.ColWidth(0) = 800
grid_fac.ColWidth(1) = 1000
grid_fac.ColWidth(2) = 1000
grid_fac.ColWidth(3) = 3000
grid_fac.ColWidth(4) = 1200
grid_fac.ColWidth(5) = 1200

grid_fac.TextMatrix(0, 0) = "Fecha"
grid_fac.TextMatrix(0, 1) = "Cuenta"
grid_fac.TextMatrix(0, 2) = "Voucher"
grid_fac.TextMatrix(0, 3) = "Glosa"
grid_fac.TextMatrix(0, 4) = "Debe"
grid_fac.TextMatrix(0, 5) = "Haber"

End Sub
Public Sub CABE_ING()
grid_fac.FontName = "Sans Serif"
grid_fac.FontSize = 7
grid_fac.Cols = 6
grid_fac.Rows = 100
grid_fac.Clear
grid_fac.ColWidth(0) = 300
grid_fac.ColWidth(1) = 1000
grid_fac.ColWidth(2) = 1800
grid_fac.ColWidth(3) = 2200
grid_fac.ColWidth(4) = 1200
grid_fac.ColWidth(5) = 1200

grid_fac.TextMatrix(0, 0) = "Item"
grid_fac.TextMatrix(0, 1) = "Cuenta"
grid_fac.TextMatrix(0, 2) = "Descripcion"
grid_fac.TextMatrix(0, 3) = "Glosa"
grid_fac.TextMatrix(0, 4) = "Debe"
grid_fac.TextMatrix(0, 5) = "Haber"

End Sub

Public Sub LISTA_CUENTAS()

grid1.TextMatrix(0, 0) = " Cuenta "
grid1.TextMatrix(0, 1) = " Descripción "

End Sub

Public Function PRO_CUENTAS(archi As String)
Dim TEMPO As String
Dim fila As Integer
Dim NUMCAMPO As Integer
Dim wBusca As String
Dim OJO As String * 1
wBusca = "SELECT * FROM COMAEST WHERE COM_DESCRIPCION >= ? AND COM_CODCIA = ? ORDER BY COM_DESCRIPCION"

If UNICO <> wBusca Then
   Set PSX = CN.CreateQuery("", wBusca)
End If
PSX.rdoParameters(0) = LEIDO.text
PSX.rdoParameters(1) = LK_CODCIA

If UNICO <> wBusca Then
   Set X = PSX.OpenResultset(rdOpenKeyset)
End If
If UNICO = wBusca Then
   X.Requery
   If X.RowCount > 0 Then
      X.MoveFirst
   End If
End If
If X.EOF = True Then
   grid1.Clear
   grid1.Visible = True
   grid1.Row = 1
   grid1.Col = 1
   grid1.text = "No hay registros"
   llave1 = ""
   'LEIDO.Locked = True
   Exit Function
End If

If X.rdoColumns(0) = llave1 Then
    Exit Function
End If
If X.RowCount > 0 Then
   llave1 = X.rdoColumns(0)
End If

UNICO = wBusca
grid1.Rows = 2
fila = 0
grid1.Clear
LISTA_CUENTAS
LEIDO.Locked = False

Do Until X.EOF Or fila = 50
    fila = fila + 1
    grid1.TextMatrix(fila, 0) = Nulo_Valors(X.rdoColumns(1))
    grid1.TextMatrix(fila, 1) = Nulo_Valors(X.rdoColumns(2))
    grid1.TextMatrix(fila, 2) = X.rdoColumns(7)
    X.MoveNext
    grid1.Rows = grid1.Rows + 1
Loop
grid1.TextMatrix(fila + 1, 1) = "                * * *    END    * * * "
grid1.Row = 1
grid1.Col = 1
grid1.Visible = True
If LEIDO.Enabled And LEIDO.Visible Then
    LEIDO.SetFocus
End If


End Function

