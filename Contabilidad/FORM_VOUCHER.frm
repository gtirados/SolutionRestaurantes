VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FORM_VOUCHER 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Ingreso de Vouchers"
   ClientHeight    =   6495
   ClientLeft      =   75
   ClientTop       =   1380
   ClientWidth     =   9480
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
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6495
   ScaleWidth      =   9480
   Tag             =   "55"
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Eliminar 
      BackColor       =   &H00FF00FF&
      Caption         =   "&Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   16
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.TextBox i_fecha 
      Height          =   375
      Left            =   1080
      TabIndex        =   13
      Text            =   "i_fecha"
      Top             =   240
      Width           =   1335
   End
   Begin VB.CommandButton SALIR 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   8
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton cancelar 
      BackColor       =   &H00000040&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Actualizar 
      BackColor       =   &H00FF00FF&
      Caption         =   "&Actualizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   5280
      Width           =   1095
   End
   Begin ComctlLib.ProgressBar Barra 
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Tag             =   "0"
      Top             =   4800
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   327680
      Appearance      =   1
      MouseIcon       =   "FORM_VOUCHER.frx":0000
      Min             =   77
      Max             =   91
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Items"
      ForeColor       =   &H00000000&
      Height          =   3765
      Left            =   120
      TabIndex        =   0
      Tag             =   "100"
      Top             =   720
      Width           =   8535
      Begin VB.ComboBox i_d_h 
         Height          =   315
         ItemData        =   "FORM_VOUCHER.frx":001C
         Left            =   5400
         List            =   "FORM_VOUCHER.frx":0026
         TabIndex        =   15
         Text            =   "i_d_h"
         Top             =   3240
         Width           =   615
      End
      Begin MSFlexGridLib.MSFlexGrid grid_fac 
         Height          =   2775
         Left            =   0
         TabIndex        =   9
         Tag             =   "9999"
         Top             =   240
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   4895
         _Version        =   327680
         Cols            =   7
         FormatString    =   $"FORM_VOUCHER.frx":0030
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
         Height          =   405
         Left            =   1560
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "9999"
         Text            =   "i_glosa"
         Top             =   3240
         Width           =   3615
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
         Height          =   435
         Left            =   120
         MaxLength       =   10
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "9999"
         Top             =   3240
         Width           =   1215
      End
      Begin VB.TextBox i_importe 
         Alignment       =   1  'Right Justify
         Height          =   405
         Left            =   6120
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "9999"
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "D/H"
         Height          =   255
         Left            =   5400
         TabIndex        =   12
         Tag             =   "9999"
         Top             =   3000
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Glosa:"
         Height          =   195
         Left            =   2640
         TabIndex        =   10
         Tag             =   "9999"
         Top             =   3000
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Lprecio 
         Caption         =   "Importe:"
         Height          =   255
         Left            =   6480
         TabIndex        =   4
         Tag             =   "9999"
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label LCODART 
         Caption         =   "Cuenta:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Tag             =   "9999"
         Top             =   3000
         Width           =   645
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Fecha:"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   240
      Width           =   855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   1
      Tag             =   "9999"
      X1              =   -480
      X2              =   9480
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   0
      Tag             =   "9999"
      X1              =   0
      X2              =   9480
      Y1              =   5160
      Y2              =   5160
   End
End
Attribute VB_Name = "FORM_VOUCHER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Public fila As Integer
Dim ws_bruto_d, ws_bruto_h As Currency
Option Explicit

Private Sub boton_autorizacion_Click()
grid_autorizacion.Visible = True
grid_autorizacion.SetFocus
End Sub

Private Sub TOMA_DATOS()
'INICIALIZA_INPUT
Dim pos2 As Integer
Dim CARAC As String
PUB_NUM_OPER = Val(FORMGEN.i_num_oper.text)

PUB_CODCLIE = Val(FORMGEN.i_codcli.text)
PUB_CANT_CHEQ = Val(FORMGEN.i_cant_cheq.text)
PUB_NUM_INI = Val(FORMGEN.i_num_ini.text)
PUB_CODVEN = Val(FORMGEN.i_codven.text)
PUB_IMPORTE = Val(FORMGEN.i_importe.text)
PUB_NETO = Val(FORMGEN.i_neto.text)
PUB_SUBTOTAL = Val(FORMGEN.i_subtotal.text)

PUB_IMPTO = Val(FORMGEN.i_impto.text)
PUB_DESCTO = Val(FORMGEN.i_descto.text)
PUB_INTERES = Val(FORMGEN.i_interes.text)
PUB_GASTOS = Val(FORMGEN.i_gastos.text)
PUB_CANTIDAD = Val(FORMGEN.i_cantidad.text)

PUB_SERDOC = 0 'Val(FORMGEN.i_serdoc.text)

PUB_NUMSER = Val(FORMGEN.i_numser.text)
PUB_CODALI = FORMGEN.i_codali.text

PUB_FECHA_LOTE = FORMGEN.i_fecha_lote.text
If IsDate(PUB_FECHA_LOTE) = False Then
   PUB_FECHA_LOTE = #1/1/1900#
End If
PUB_NUM_LOTE = Val(FORMGEN.i_num_lote.text)

PUB_NUMFAC = Val(FORMGEN.i_numfac.text)

PUB_NUMSER_C = Val(FORMGEN.i_numser_c.text)
PUB_NUMFAC_C = Val(FORMGEN.i_numfac_c.text)

PUB_USUARIO = RTrim(FORMGEN.i_CODUSU.text)
If FORMGEN.i_nat_jur.Visible = True Then
   PUB_NAT_JUR = Left(FORMGEN.i_nat_jur.text, 1)
End If
PUB_GRUPO_ACT = Val(FORMGEN.i_grupo_act.text)
PUB_GRUPO_ANT = Val(FORMGEN.i_grupo_ant.text)


PUB_TIPO_BLOQ_act1 = Left(FORMGEN.i_tipo_bloq_act1.text, 1)
PUB_TIPO_BLOQ_act2 = Left(FORMGEN.i_tipo_bloq_act2.text, 1)
PUB_TIPO_BLOQ_act3 = Left(FORMGEN.i_tipo_bloq_act3.text, 1)
PUB_TIPO_BLOQ_act4 = Left(FORMGEN.i_tipo_bloq_act4.text, 1)
PUB_TIPO_BLOQ_ant1 = Left(FORMGEN.i_tipo_bloq_ant1.text, 1)
PUB_TIPO_BLOQ_ant2 = Left(FORMGEN.i_tipo_bloq_ant2.text, 1)
PUB_TIPO_BLOQ_ant3 = Left(FORMGEN.i_tipo_bloq_ant3.text, 1)
PUB_TIPO_BLOQ_ant4 = Left(FORMGEN.i_tipo_bloq_ant4.text, 1)

PUB_LIMCRE_ACT = Val(FORMGEN.i_limcre.text)
PUB_LIMCRE_ANT = Val(FORMGEN.i_limcre_ant.text)
PUB_CODBAN = Val(FORMGEN.i_codban.text)
pub_dias = Val(FORMGEN.i_dias.text)
PUB_TASA_ADE = Val(FORMGEN.i_tasa_ade.text)
PUB_PORDESCTO1 = Val(FORMGEN.i_pordescto1.text)
PUB_CONCEPTO = FORMGEN.i_concepto.text
PUB_CHENUM = Val(FORMGEN.i_chenum.text)
PUB_CHESER = FORMGEN.i_cheser.text
PUB_FBG = FORMGEN.i_fbg.text

PUB_CHESEC = 0
PUB_NUMGUIA = Val(FORMGEN.i_numguia.text)
PUB_TASA_VEN = Val(FORMGEN.i_tasa_ven.text)
PUB_IMPORTE_AMORT = Val(FORMGEN.i_importe_amort.text)
PUB_INTVEN = Val(FORMGEN.i_intven.text)
PUB_GASTOS_FIJOS = Val(FORMGEN.i_gastos_fijos.text)
PUB_GASTOS_NOT = Val(FORMGEN.i_gastos_not.text)
PUB_CODCIA_R = Right(FORMGEN.i_cias.text, 2)
PUB_NOMCIA = Left(FORMGEN.i_cias.text, 20)
'POS1 = InStr(1, FORMGEN.i_2403.text, "@", 1)
'POS1 = POS1 + 1
'CARAC = Mid(FORMGEN.i_2403.text, POS1)
'PUB_TIPMOV = Val(CARAC)
'PUB_NUMTAB = PUB_TIPMOV
'PUB_TIPREG = 4
'SQ_OPER = 1
'LEER_TAB_LLAVE
'If Not tab_llave.EOF Then
'If Not IsNull(tab_llave!tab_cp) Then
'   PUB_CP = tab_llave!tab_cp
'End If
'End If
PUB_NUMDOC_R = Val(FORMGEN.i_numdoc_r.text)
PUB_NUMDOC = Val(FORMGEN.i_numdoc.text)

PUB_CODART = Val(FORMGEN.i_codart.text)
PUB_PRECIO = Val(FORMGEN.i_precio.text)
PUB_PORDESCTO2 = Val(FORMGEN.i_pordescto2.text)
PUB_INTADE = Val(FORMGEN.i_intade.text)
PUB_AJUSTE = Val(FORMGEN.i_ajuste.text)
PUB_TOTAL = Val(FORMGEN.i_total_pago.text)
'TEXTOX(m) = FORMGEN.i_total_pago.text
If IsDate(FORMGEN.i_fecha_vcto.text) Then
    PUB_FECHA_VCTO = CDate(FORMGEN.i_fecha_vcto.text)
Else
    PUB_FECHA_VCTO = 0
End If
'PUB_saldocar = FORMGEN.i_saldocar.text
pub_diasA = Val(FORMGEN.i_diasA.text)
pub_diasV = Val(FORMGEN.i_diasV.text)
PUB_TASA_VENTA = Val(FORMGEN.i_tasa_venta.text)
If FORMGEN.i_def.ListCount = 0 Then
   PUB_SECUENCIA = 0
Else
   pos2 = InStr(1, FORMGEN.i_def.text, ".", 1)
   CARAC = Mid(FORMGEN.i_def.text, 1, pos2)
   PUB_SECUENCIA = Val(CARAC)
End If

End Sub







Private Sub Boton_Compras_Click()
Dim Tit As String
Dim WS_IMPORTE As Currency
FORMGEN.grid_far.Visible = True
Tit = FORMGEN.grid_far.FormatString
FORMGEN.grid_far.Clear
FORMGEN.grid_far.FormatString = Tit

FORMGEN.grid_far.SetFocus
PUB_FECHA = FORMGEN.i_fecha_vcto.text
PUB_CODCLIE = Val(FORMGEN.i_codcli.text)

SQ_OPER = 2
LEER_FAR_LLAVE

fila = 0
If far_codcli.EOF = True Then
   FORMGEN.grid_far.Rows = 2
   FORMGEN.grid_far.Row = 1
   FORMGEN.grid_far.Col = 1
   FORMGEN.grid_far.text = "No hay registros"

End If

Do Until far_codcli.EOF Or fila = 90
    ws_respuesta = vbYes
    If far_codcli!FAR_NUMSEC <> 1 Then
       ws_respuesta = vbNo
    End If

       
    If ws_respuesta = vbYes Then
       fila = fila + 1
       FORMGEN.grid_far.Rows = fila + 1
       FORMGEN.grid_far.Row = fila
       FORMGEN.grid_far.Col = 0
       FORMGEN.grid_far.text = ""
       FORMGEN.grid_far.Col = 1
       FORMGEN.grid_far.text = far_codcli!FAR_FECHA
       FORMGEN.grid_far.Col = 2
       FORMGEN.grid_far.text = far_codcli!far_numser
       FORMGEN.grid_far.Col = 3
       FORMGEN.grid_far.text = far_codcli!FAR_NUMFAC
       FORMGEN.grid_far.Col = 4
       FORMGEN.grid_far.text = far_codcli!far_numser_c
       FORMGEN.grid_far.Col = 5
       FORMGEN.grid_far.text = far_codcli!far_numfac_c
       FORMGEN.grid_far.Col = 6
       FORMGEN.grid_far.text = far_codcli!FAR_NUMGUIA
      
       FORMGEN.grid_far.Col = 7
       WS_IMPORTE = Val(far_codcli!far_bruto) + Val(far_codcli!far_gastos) - Val(far_codcli!far_descto) + Val(far_codcli!far_impto)
       FORMGEN.grid_far.text = WS_IMPORTE
       FORMGEN.grid_far.Col = 8
       FORMGEN.grid_far.text = far_codcli!FAR_CP
       FORMGEN.grid_far.Col = 9
       FORMGEN.grid_far.text = far_codcli!FAR_CODCLIE
       FORMGEN.grid_far.Col = 10
       FORMGEN.grid_far.text = far_codcli!far_tipmov
       FORMGEN.grid_far.Col = 11
       FORMGEN.grid_far.text = far_codcli!FAR_CODCIA
     
       
    End If
    far_codcli.MoveNext
Loop

FORMGEN.grid_far.Row = 1
FORMGEN.grid_far.Col = 1

End Sub

Private Sub Boton_Letras_Click()
   If FORMGEN.i_codcli.text <> "" Then
      PROCESA_GRID_LETRA
   End If

End Sub

Private Sub Actualizar_Click()
sn_mensaje = " ¿Confirma Actualizacion .. ?"
ws_respuesta = MsgBox(sn_mensaje, WS_ESTILO, WS_TITULO)
If ws_respuesta = vbNo Then
   GoTo fin
End If


cov_llave.Edit
cov_llave!COV_CUENTA = i_cuenta.text
cov_llave!COV_GLOSA = i_glosa.text
cov_llave!COV_IMPORTE = i_importe.text
cov_llave!COV_DH = i_d_h.text
cov_llave.Update
fin:
End Sub

Private Sub cancelar_Click()
fila = 0
Do Until fila = grid_fac.Rows
  fila = fila + 1
  gird_fac.Row = fila
  If grid_fac.CellBackColor <> vbWhite Then
     For a = 0 To grid_fac.Cols - 1
         grid_fac.Col = a
         grid_fac.CellBackColor = vbWhite
     Next a
  End If
Loop

End Sub
Private Sub Cheques_Click()
Dim WS_SALDO As Currency
Dim Tit As String
Dim success%
Screen.MousePointer = 11
PUB_CODBAN = Val(FORMGEN.i_codban.text)

SQ_OPER = 1
LEER_CCM_LLAVE
If ccm_llave.EOF Then
   Screen.MousePointer = 0
   MsgBox "Banco no existe..."
   GoTo fin
End If

fila = 0
PUB_FECHA = #1/1/1900#
SQ_OPER = 4
LEER_CHE_LLAVE
If che_repo.EOF = True Then
   Screen.MousePointer = 11
   MsgBox "No hay estado de cuenta..."
   GoTo fin
End If
success% = SetWindowPos(FrmProcesa.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
FrmProcesa.Show
MUESTRA_PROCESO "Un Momento ..."
Screen.MousePointer = 11
che_repo.MoveLast
WS_SALDO = ccm_llave!CCM_SALDO
FORMGEN.grid_che.Visible = False
Tit = FORMGEN.grid_che.FormatString
grid_che.Clear
FORMGEN.grid_che.FormatString = Tit
Do Until che_repo.BOF
       fila = fila + 1
       FORMGEN.grid_che.Rows = fila + 1
       FORMGEN.grid_che.Row = fila
       FORMGEN.grid_che.Col = 0
       FORMGEN.grid_che.text = fila & ".-"
       FORMGEN.grid_che.Col = 1
       FORMGEN.grid_che.text = che_repo!che_fecha
       If che_repo!che_estado = "E" Or che_repo!CHE_CODTRA = 1111 Then
          FORMGEN.grid_che.CellBackColor = vb3DLight
       End If
       FORMGEN.grid_che.Col = 2
       FORMGEN.grid_che.ColAlignment(2) = 0
       If che_repo!CHE_CHENUM = 0 Then
          FORMGEN.grid_che.text = " "
       Else
          FORMGEN.grid_che.text = che_repo!CHE_CHESER & "-" & che_repo!CHE_CHENUM & "-" & che_repo!CHE_CHESEC
       End If
       If che_repo!che_estado = "E" Or che_repo!CHE_CODTRA = 1111 Then
          FORMGEN.grid_che.CellBackColor = vb3DLight
       End If
       FORMGEN.grid_che.Col = 3
       FORMGEN.grid_che.text = Nulo_Valors(che_repo!che_abreviado)
       If che_repo!che_estado = "E" Or che_repo!CHE_CODTRA = 1111 Then
          FORMGEN.grid_che.CellBackColor = vb3DLight
       End If
       FORMGEN.grid_che.Col = 4
       If che_repo!che_estado = "E" Or che_repo!CHE_CODTRA = 1111 Then
          FORMGEN.grid_che.CellBackColor = vb3DLight
       End If
       If che_repo!CHE_SIGNO_CCM = 1 Then
          FORMGEN.grid_che.text = Format(che_repo!che_importe, "CURRENCY")
          FORMGEN.grid_che.Col = 5
          FORMGEN.grid_che.text = " "
       Else
          FORMGEN.grid_che.text = " "
          FORMGEN.grid_che.Col = 5
          FORMGEN.grid_che.text = Format(che_repo!che_importe, "CURRENCY")
       End If
        If che_repo!che_estado = "E" Or che_repo!CHE_CODTRA = 1111 Then
           FORMGEN.grid_che.CellBackColor = vb3DLight
        End If
       FORMGEN.grid_che.Col = 6
       FORMGEN.grid_che.text = Format(che_repo!che_saldo, "CURRENCY")
       If WS_SALDO <> che_repo!che_saldo Then
          MsgBox "Avisar a Computo...hay diferencia... " & che_repo!CHE_CHENUM & che_repo!CHE_CHESEC
       End If
       WS_SALDO = WS_SALDO - che_repo!CHE_SIGNO_CCM * che_repo!che_importe
        If che_repo!che_estado = "E" Or che_repo!CHE_CODTRA = 1111 Then
        FORMGEN.grid_che.CellBackColor = vb3DLight
        End If
       
       FORMGEN.grid_che.Col = 7
       FORMGEN.grid_che.ColAlignment(7) = 0

       FORMGEN.grid_che.text = Mid(che_repo!che_CONCEPTO, 1, 80)
       If che_repo!che_estado = "E" Or che_repo!CHE_CODTRA = 1111 Then
          FORMGEN.grid_che.CellBackColor = vb3DLight
       End If
       FORMGEN.grid_che.Col = 8
       FORMGEN.grid_che.ColAlignment(8) = 0

       FORMGEN.grid_che.text = Mid(che_repo!che_CONCEPTO, 81, 80)
       If che_repo!che_estado = "E" Or che_repo!CHE_CODTRA = 1111 Then
          FORMGEN.grid_che.CellBackColor = vb3DLight
       End If
       FORMGEN.grid_che.Col = 9
       FORMGEN.grid_che.ColAlignment(9) = 0

       FORMGEN.grid_che.text = Mid(che_repo!che_CONCEPTO, 161, 90)
       If che_repo!che_estado = "E" Or che_repo!CHE_CODTRA = 1111 Then
          FORMGEN.grid_che.CellBackColor = vb3DLight
       End If

       che_repo.MovePrevious
Loop
FrmProcesa.Hide
Screen.MousePointer = 0
FORMGEN.grid_che.Visible = True
FORMGEN.grid_che.SetFocus

fin:

End Sub




Private Sub Combo1_Change()

End Sub

Private Sub Diario_Click()

End Sub

Private Sub Diario_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

NUMERO = FORMGEN.Diario.TabIndex
avanza_campo
fin:

End Sub

Private Sub Eliminar_Click()
fila = 0
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
      End If
  Loop
  grid_fac.Clear
End Sub

Private Sub Form_Click()
FORMGEN.WhatsThisMode
End Sub

Private Sub Form_Load()
Dim cadena As String
Dim ws_indice As Integer

cadena = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND COV_FECHA_voucher >=?    ORDER BY COV_CODCIA, COV_FECHA_voucher, COV_NRO_MOV"
Set PSCOV_MAYOR = CN.CreateQuery("", cadena)
Set cov_mayor = PSCOV_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)

cadena = "SELECT * FROM COMOV WHERE COV_CODCIA = ?  AND COV_NRO_MOV=?   ORDER BY COV_CODCIA, COV_FECHA_voucher, COV_NRO_MOV"
Set PSCOV_LLAVE = CN.CreateQuery("", cadena)
Set cov_llave = PSCOV_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

'grid_fac.FontName = "Sans Serif"
'grid_fac.FontSize = 7

grid_fac.Cols = 6
grid_fac.Rows = 50
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
Dim WS_DATOS As String * 2
Dim WS_FLAG As Integer
Dim WS_TRANSITO As String * 1
Dim WS_FLAG2 As Integer
Dim subtotal As Currency
Dim WS_BRUTO2 As Currency
Dim ws_diferencia As Currency
Dim WS_CORRELATIVO As Double
Dim neto As Currency
Dim fx As Integer
Dim FLAG As Boolean
Dim msg_err As String
Dim ws_flag_arm As Integer
Dim ws_flag_che As Integer
Dim ws_flag_far As Integer
Dim ws_flag_car As Integer
Dim ws_flag_par As Integer
Dim ws_flag_ccm As Integer
Dim ws_flag_ven As Integer
Dim ws_flag_lot As Integer
Dim ws_tot_CANTIDAD As Currency
Dim WS_UNIDAD   As Currency
Screen.MousePointer = 11
WS_NUMSEC = 0


exito = True
TOMA_DATOS
CAPTURA_DATOS


PASA:

If PUB_SUBTOTAL <> 0 Then
If WS_BRUTO <> PUB_SUBTOTAL Then
msg = " ¿Total Bruto no coincide ...Desea Continuar ?"
ws_respuesta = MsgBox(msg, WS_ESTILO, WS_TITULO)
If ws_respuesta = vbNo Then
   GoTo fin
End If
End If
End If


CN.Execute "Begin Transaction", rdExecDirect
cadena = "SELECT * FROM CONTROLL"
Set con_llave = CN.OpenResultset(cadena, rdOpenKeyset, rdConcurLock)


Barra.Visible = True
Barra.Min = 62
Barra.Max = 90

Barra.Value = 62

N = 62
Do Until Val(tra_llave(N)) = 0 Or N = 76 Or exito = False
NUMERO = tra_llave(N)
On NUMERO GoSub CON1, CON2, CON3, CON4, CON5, CON6, CON7, CON8, CON9, CON10, CON11, CON12, CON13, CON14, CON15, CON16, CON17, CON18, CON19, CON20, CON21, CON22
Barra.Value = N

N = N + 1
Loop
If exito = False Then
   con_llave.Close
   Barra.Visible = False

   CN.Execute "Rollback Transaction", rdExecDirect
   MsgBox msg_err
   PUB_TIPDOC = Nulo_Valors(def_llave!def_tipdoc)
   GoTo errorr
End If


PUB_NUM_OPER = 32000
SQ_OPER = 2
LEER_ALL_LLAVE
If all_menor.EOF = False Then
   all_menor.MoveLast
   PUB_NUM_OPER_XXX = all_menor!all_numoper
Else
   PUB_NUM_OPER_XXX = 0
End If
PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1



N = 77
Do Until Val(tra_llave(N)) = 0 Or N = 91
NUMERO = tra_llave(N)
 On NUMERO GoSub ACT1, ACT2, ACT3, ACT4, ACT5, ACT6, ACT7, ACT8, ACT9, ACT10, ACT11, ACT12, ACT13, ACT14, ACT15, ACT16, ACT17, ACT18
Barra.Value = N
N = N + 1
Loop
Barra.Value = 90
N = 105
Do Until Val(tra_llave(N)) = 0 Or N = 107
NUMERO = tra_llave(N)
 On NUMERO GoSub REP1, REP2, REP3

N = N + 1
Loop

con_llave.Close
CN.Execute "Commit Transaction", rdExecDirect


i_importe.text = ""
cancela_todo
pasa_def
Barra.Visible = False

' AQUI INICIALIZO VARIABLES QUE NO SE CEREAN  AL ENTRAR AL COMMAND1.CLICK
WS_BRUTO = 0
filax = 0
If FORMGEN.i_importe.Visible = True And (PUB_IMPORTE = 0 Or PUB_CHESEC <> 0) Then
   FORMGEN.i_importe.SetFocus
Else
   FORMGEN.Controls(WS_INDICE_RETORNO).SetFocus
End If

GoTo fin

CON1:

SQ_OPER = 1
If PUB_CODCLIE = 0 Then
   msg_err = "Cliente Invalido ..."
   FORMGEN.i_codcli.SetFocus
   exito = False
End If
PUB_CODCIA = LK_CODCIA
LEER_CLI_LLAVE
If cli_llave.EOF = True Then
   msg_err = "Cliente/Proveedor no existe...."
   FORMGEN.i_codcli.SetFocus
   exito = False
End If
Return

CON2:
If PUB_CODART = 0 Then
   msg_err = "Articulo Invalido ..."
   FORMGEN.i_codart.SetFocus
   exito = False
End If

PSART_LLAVE.rdoParameters(0) = PUB_CODART
art_llave.Requery
If art_llave.EOF = True Then
   msg_err = "ARTICULO NO EXISTE...."
   FORMGEN.i_codart.SetFocus
   exito = False
End If
Return


CON3:
SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
LEER_VEN_LLAVE
If ven_llave.EOF Then
   msg_err = "vendedor no existe... "
   FORMGEN.i_codven.SetFocus
   exito = False

End If
If WS_BRUTO = 0 Then
   msg_err = "Importe errado... "
   exito = False
End If
Return
CON4:
   Select Case PUB_FBG
   Case "G"
      If PU_NUMFAC >= ven_llave!VEM_I_NUM_G And PU_NUMFAC <= ven_llave!VEM_F_NUM_G Then
      Else
         PU_NUMFAC = ven_llave!VEM_I_NUM_G
      End If
   Case "B"
      If PU_NUMFAC >= ven_llave!VEM_I_NUM_B And PU_NUMFAC <= ven_llave!VEM_F_NUM_B Then
      Else
         PU_NUMFAC = ven_llave!VEM_I_NUM_B
      End If
   Case "F"
      If PU_NUMFAC >= ven_llave!VEM_I_NUM_F And PU_NUMFAC <= ven_llave!VEM_F_NUM_F Then
      Else
         PU_NUMFAC = ven_llave!VEM_I_NUM_F
      End If
   Case Else
      MsgBox "Numero incorrecto ...Revisar talonario de vendedor"
   FORMGEN.i_fbg.SetFocus
   End Select

Return

CON5:
SQ_OPER = 1
PUB_CODCIAL = LK_CODCIA
LEER_CAR_LLAVE
If car_llave.EOF = False Then
   msg_err = "DOCUMENTO YA EXISTE..."
   exito = False
   
End If
Return

CON6:
If IsDate(PUB_FECHA_LOTE) = False Then
   msg_err = "Falta seleccionar Lote ..."
   FORMGEN.i_codali.SetFocus
   exito = False
End If


   SQ_OPER = 1
   LEER_LOT_LLAVE
   If lot_llave.EOF Then
      msg_err = "Lote no existe..."
      FORMGEN.i_codali.SetFocus
      exito = False
      Return
   End If
   
   ws_diferencia = Val(lot_llave!LOT_PESO) + Val(lot_llave!LOT_PESO_venta)

   If ws_diferencia < 0 Or ws_diferencia = 0 Then
      msg_err = "No hay stock para venta de este lote..."
      FORMGEN.i_codali.SetFocus
      exito = False
   End If
   
   
   Return

CON7:
   SQ_OPER = 1
   LEER_CCM_LLAVE
   If ccm_llave.EOF Then
      msg_err = "banco no existe..."
      FORMGEN.i_codban.SetFocus
      exito = False
   End If
   Return
   
   
CON8:
Return

SQ_OPER = 1
PUB_CODCIA = PUB_CODCIA_R
LEER_PAR_LLAVE
If par_llave.EOF Then
   msg_err = "CIA NO EXISTE ..."
   End
End If
If par_llave!par_flag_cierre > 7 Then
   msg_err = "!!! Punto de Venta Receptor YA CERRO ..."
   exito = False
End If



Return

CON9:
SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
LEER_PAR_LLAVE
If par_llave.EOF Then
   msg_err = "CIA NO EXISTE ..."
   End
End If

If par_llave!par_flag_cierre > 7 Then
   msg_err = "!!! PTO. DE VENTA YA CERRO ..."
   exito = False
End If

Return

CON10:
count = 1
fila = 0
WS_FLAG2 = 0
ws_respuesta = vbYes
Do While ws_respuesta = vbYes
 WS_FLAG = 0
 FORMGEN.gridl.Row = count
  
 FORMGEN.gridl.Col = 0
 If FORMGEN.gridl.text = "@" Then
    WS_FLAG = 1
 End If
 
 FORMGEN.gridl.Col = 6
 If WS_FLAG = 1 And WS_FLAG2 = 0 Then
    PUB_FECHA_VCTO = FORMGEN.gridl.text
    WS_FLAG2 = 1
 End If
  
 
 If WS_FLAG = 1 Then
 If PUB_FECHA_VCTO = FORMGEN.gridl.text Then
   fila = fila + 1
 Else
   msg_err = "!!! NO PROCEDE LETRA CON DISTINTA FECHA ..."
   exito = False
   Return
 End If
 End If
 
 
 
 count = count + 1
 FORMGEN.gridl.Row = count
 FORMGEN.gridl.Col = 12
 PUB_NUMDOC = Val(FORMGEN.gridl.text)
 If FORMGEN.gridl.Rows = FORMGEN.gridl.Row + 1 Then
    ws_respuesta = vbNo
 End If
 Loop

If fila = 1 Then
 msg_err = "!!! NO PROCEDE UNA SOLA LETRA ..."
 exito = False
 Return
End If
If fila = 0 Then
 msg_err = "!!! NO PROCEDE ..."
 exito = False
 Return
End If

Return

CON11:
'FORMGEN.i_tasa_ade.text = gen!GEN_TASA_LEG_ADEL
'FORMGEN.i_tasa_ven.text = gen!GEN_TASA_LEG_ADEL

PUB_GASTOS_FIJOS = FORMGEN.i_gastos_fijos.text
PUB_GASTOS_NOT = FORMGEN.i_gastos_not.text
pub_diasV = FORMGEN.i_diasV.text
PUB_FECHA_VCTO = FORMGEN.i_fecha_vcto.text
pub_diasA = FORMGEN.i_diasA.text
PUB_INTADE = FORMGEN.i_intade.text
PUB_INTVEN = FORMGEN.i_intven.text
PUB_IMPORTE_AMORT = FORMGEN.i_importe_amort.text
pub_total_liq = PUB_IMPORTE_AMORT + PUB_INTVEN + PUB_INTADE + PUB_GASTOS_FIJOS + PUB_GASTOS_NOT
If i_importe.text <> pub_total_liq Then
   msg_err = "hay algun error...."
   exito = False
   Return
End If

Return

CON12:
If FORMGEN.i_numfac.ListIndex = -1 Then
   msg_err = "No ha seleccionado Numero de Documento ..."
   exito = False
   Return
End If

SQ_OPER = 1
PU_TIPMOV = PUB_TIPMOV
PU_NUMSER = PUB_NUMSER
PU_NUMFAC = PUB_NUMFAC
PU_CODCIA = PUB_CODCIA
PU_FBG = PUB_FBG
PU_NUMSEC = 1
LEER_FAR_LLAVE
If far_llave.EOF = False Then
   msg_err = "Documento ya Procesado..."
   exito = False
   Return
End If


Return

CON13:
' REVISAR ESTO .....
'If Val(FORMGEN.i_numfac_c.text) = 0 Then
'   msg_err = "Falta indicar N. factura"
'   exito = False
'   Return
'End If

SQ_OPER = 1
PU_TIPMOV = PUB_TIPMOV
PU_NUMSER = PUB_NUMSER
PU_NUMFAC = PUB_NUMFAC
PU_CODCIA = PUB_CODCIA
PU_FBG = PUB_FBG
PU_NUMSEC = 1
LEER_FAR_LLAVE
If far_llave.EOF = True Then
   msg_err = "No existe Documento ..."
   exito = False
   Return
End If

Return

CON14:
'COMPRAS

If PUB_IMPTO <> 0 Then
   If PUB_NUMFAC_C = 0 Then
      msg_err = "Revisar Factura ...con Impto..."
      exito = False
   End If
End If

If PUB_IMPTO = 0 Then
   If PUB_NUMFAC_C <> 0 Then
      msg_err = "Revisar Factura ...con Impto..."
      exito = False
   End If
End If

SQ_OPER = 1
LEER_CAR_LLAVE
If car_llave.EOF = False Then
   msg_err = "CARTERA YA EXISTE..."
   exito = False
End If
Return

CON15:
'Verifica si es un cheque ya trabajado.. que existe en archivo...
   PUB_CHESEC = 99
   SQ_OPER = 1
   LEER_CHE_LLAVE
   If che_llave.EOF Then
       msg_err = "Numero de Cheque errado ...."
       exito = False
       Return
   End If
che_llave.MoveLast
' ?????? REVISAR ESTO EL IGUAL ???
If che_llave!che_fecha = #1/1/1900# And che_llave!che_fecha <> LK_FECHA_DIA Then
   msg_err = "Cheque ya procesado...otro dia..."
   exito = False
   Return
End If

If che_llave!che_fecha > #1/1/1900# Then
   PUB_CHESEC = che_llave!CHE_CHESEC + 1
   Return
End If



' ese cheque es el que sigue en la secuencia ????
' lectura de todos los cheques de fecha 0 y serie xxx

   PUB_CHESEC = 0

   SQ_OPER = 3
   PUB_FECHA = 0
   LEER_CHE_LLAVE
   If che_menor.EOF Then
      msg_err = "Numero de Serie errado ...."
      exito = False
      Return
   End If
che_menor.MoveFirst
   
If PUB_CHENUM = che_menor!CHE_CHENUM Then
Else
'   muestra_cheques... CAMBIAR A CHE_MENOR
   msg_err = "No corresponde numero de Cheque ...?"
   exito = False
   Return
End If

   
Return

CON16:
SQ_OPER = 1
LEER_TAB_LLAVE



Return

CON17:

If LK_CODTRA <> 2401 Then
   WS_DESCTO = PUB_DESCTO
   WS_IGV = PUB_IMPTO
   Return
End If

WS_DESCTO1 = WS_BRUTO * Val(FORMGEN.i_pordescto1) / 100
WS_DESCTO1 = redondea(WS_DESCTO1)
WS_DESCTO2 = (WS_BRUTO - WS_DESCTO1) * Val(FORMGEN.i_pordescto2) / 100
WS_DESCTO2 = redondea(WS_DESCTO2)
WS_DESCTO2 = WS_DESCTO1 + WS_DESCTO2
WS_IGV = 0
If PUB_FBG = "B" Or PUB_FBG = "F" Then
   WS_IGV = LK_IGV
End If
If LK_CODTRA = 2401 Then
   WS_IMPORTE = WS_BRUTO - WS_DESCTO2
   WS_IGV = WS_IMPORTE * WS_IGV / 100
   WS_IGV = redondea(WS_IGV)
End If
   
   WS_NETO = WS_BRUTO - WS_DESCTO2 + WS_IGV + PUB_GASTOS
   WS_IMPORTE = PUB_NETO + PUB_AJUSTE - WS_NETO
   WS_IMPORTE = Format(WS_IMPORTE, "CURRENCY")
   If WS_IMPORTE <> 0 Then
       msg_err = "Hay diferencia en datos... Revisar" & NL & "total bruto=" & WS_BRUTO & NL & "Descto=" & WS_DESCTO & NL & "Igv=" & WS_IGV & NL & "Total Neto=" & WS_NETO
       exito = False
       Return
   End If


If PUB_FBG = "B" Or PUB_FBG = "F" Then
   If PUB_IMPTO = 0 Then
       msg_err = "Falta el Igv.."
       exito = False
       Return
   End If
End If

Return

CON18:
If PUB_CP = " " Then
   Return
End If

SQ_OPER = 1
LEER_CAR_LLAVE
If car_llave.EOF = True Then
   msg_err = "DOCUMENTO EN CARTERA NO EXISTE..."
   exito = False
   Return
End If

If FORMGEN.imp_orig.Visible = False Then
   msg_err = "Falta seleccionar el documento a operar ..."
   exito = False
   Return
End If

If LK_CODTRA = 2710 Then
If MOSTRAR.Caption = "LIQ. OK " Then
Else
   msg_err = "Liquidacion aun no cuadra...."
   exito = False
End If
End If
' Para operaciones de cartera con cheques
If FORMGEN.i_importe.Visible = True And FORMGEN.i_importe_amort.Visible = False Then
   PUB_IMPORTE_AMORT = PUB_IMPORTE
End If
If FORMGEN.i_importe.Visible = False And FORMGEN.i_importe_amort.Visible = True Then
   PUB_IMPORTE = PUB_IMPORTE_AMORT
End If

If PUB_IMPORTE_AMORT > car_llave!CAR_IMPORTE Then
   msg_err = "Importe no puede ser mayor...que el documento...."
   exito = False
End If
'If PUB_IMPORTE_AMORT = 0 Then
'   msg_err = "Importe debe ser mayor que 0 ..."
'   exito = False
'End If



Return

CON19:
If pub_signo_car = 0 Then
   Return
End If

SQ_OPER = 3
PUB_CODCIA = LK_CODCIA
PUB_CODCIAL = LK_CODCIA

LEER_CAR_LLAVE
If car_menor.EOF = True Then
   PUB_NUMDOC = 1
Else
   car_menor.MoveLast
   PUB_NUMDOC = car_menor!car_numdoc + 1
End If

Return

CON20:
contador = 0
Do Until contador = PUB_CANT_CHEQ
PUB_CHESEC = 0
PUB_CHENUM = contador + PUB_NUM_INI
contador = contador + 1
SQ_OPER = 1
LEER_CHE_LLAVE
If Not che_llave.EOF Then
       msg_err = "Numero de Serie y Cheque ya existe..." & PUB_CHESER & "-" & PUB_CHENUM
       exito = False
       Return
End If
Loop
Return

CON21:
SQ_OPER = 1
LEER_ALL_LLAVE
If all_llave.EOF Then
   msg_err = "Numero de Operacion Incorrecto...."
   exito = False
   Return
End If

If all_llave!all_flag_ext = "E" Then
   msg_err = "Operacion ya Extornada........"
   exito = False
   Return
End If

If all_llave!ALL_CODCIA <> LK_CODCIA Then
   msg_err = "Cia. no coincide..."
   exito = False
   Return
End If




PUB_NUM_OPER = all_llave!all_numoper
PUB_NUM_OPER_EXT = all_llave!all_numoper

PUB_CODCIA = all_llave!ALL_CODCIA
PUB_CODCIAL = all_llave!ALL_CODCIA
PUB_CODCLIE = all_llave!all_codclie
PUB_NUMDOC = Nulo_Valor0(all_llave!ALL_NUMDOC)
PUB_CODART = all_llave!ALL_CODART
PUB_IMPORTE_AMORT = all_llave!all_importe_amort
PUB_IMPORTE = all_llave!all_importe
PUB_CP = all_llave!all_cp
PUB_TIPDOC = all_llave!all_tipdoc
PUB_TIPMOV = all_llave!ALL_TIPMOV
PUB_CODBAN = all_llave!all_codban
PUB_CHENUM = all_llave!ALL_CHENUM
PUB_CHESER = Nulo_Valor0(all_llave!ALL_CHESER)
PUB_CHESEC = Nulo_Valor0(all_llave!ALL_CHESEC)
PUB_SERDOC = Nulo_Valor0(all_llave!ALL_SERDOC)
PUB_NUMDOC = Nulo_Valor0(all_llave!ALL_NUMDOC)

PUB_CANTIDAD = all_llave!ALL_CANTIDAD
PUB_NUMSER = all_llave!all_numser
PUB_FBG = all_llave!all_FBG

PUB_NUMFAC = all_llave!all_numfac
PUB_SECUENCIA = all_llave!all_secuencia
PUB_NUM_INI = all_llave!ALL_NUM_INI
pub_signo_ccm = Nulo_Valor0(all_llave!ALL_SIGNO_CCM) * -1
pub_signo_car = Nulo_Valor0(all_llave!all_SIGNO_CAR) * -1
pub_signo_arm = Nulo_Valor0(all_llave!all_sIGNO_ARM) * -1
pub_signo_caja = Nulo_Valor0(all_llave!ALL_SIGNO_CAJA) * -1
pub_signo_lot = Nulo_Valor0(all_llave!ALL_SIGNO_LOT) * -1
PUB_CONCEPTO = "Extorno - " & all_llave!all_concepto

PUB_FECHA_LOTE = all_llave!ALL_FECHA_LOTE
PUB_CODALI = all_llave!ALL_CODALI
PUB_NUM_LOTE = all_llave!ALL_NUM_LOTE

If pub_signo_ccm <> 0 Then
   SQ_OPER = 1
   LEER_CCM_LLAVE
   If ccm_llave.EOF Then
      msg_err = "banco no existe..."
      exito = False
   End If
End If

If pub_signo_car <> 0 Then
SQ_OPER = 1
PUB_CODCIAL = LK_CODCIA
LEER_CAR_LLAVE
If car_llave.EOF = True Then
   msg_err = "DOCUMENTO NO EXISTE..."
   exito = False
End If
End If

If pub_signo_arm = 0 Then
   Return
End If
PUB_NUMSER_C = PUB_NUMSER
PUB_NUMFAC_C = PUB_NUMFAC




PU_NUMSER = PUB_NUMSER
PU_NUMFAC = PUB_NUMFAC
PU_CODCIA = LK_CODCIA
PU_TIPMOV = PUB_TIPMOV
PU_FBG = PUB_FBG
'FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_NUMSER = ? AND FAR_FBG=? AND FAR_NUMFAC = ? AND FAR_NUMSEC = ? ORDER BY FAR_TIPMOV, FAR_CODCIA, FAR_NUMSER, FAR_FBG, FAR_NUMFAC"
SQ_OPER = 1
LEER_FAR_LLAVE
If far_llave.EOF = True Then
   MsgBox "!!! NO HAY ESE DOCUMENTO ..."
   exito = False
   Return
End If
If far_llave!FAR_CODCIA <> LK_CODCIA Then
   MsgBox "!!! NO TE CORRESPONDE..."
   exito = False
   Return
End If
WS_DATOS = "SI"
fila = 0
Do Until WS_DATOS = "NO"
   fila = fila + 1
   grid_fac.Row = fila
   grid_fac.Col = 1
   grid_fac.text = far_llave!far_codart
   grid_fac.Col = 3
   grid_fac.text = far_llave!FAR_UNIDAD
   grid_fac.Col = 8
   grid_fac.text = far_llave!FAR_CANTIDAD
   grid_fac.Col = 9
   grid_fac.text = far_llave!far_precio
   grid_fac.Col = 10
   grid_fac.text = " "
'   WS_NUMSEC = far_llave!FAR_NUMSEC
   far_llave.MoveNext
   If far_llave.EOF = False Then
   If far_llave!FAR_NUMFAC = PU_NUMFAC And far_llave!far_tipmov = PUB_TIPMOV And far_llave!FAR_CODCIA = LK_CODCIA Then
      Print " "
   Else
      WS_DATOS = "NO"
   End If
   End If
   If far_llave.EOF = True Then
      WS_DATOS = "NO"
   End If
   
   
Loop
   fila = fila + 1
   FORMGEN.grid_fac.Row = fila
   FORMGEN.grid_fac.Col = 1
   FORMGEN.grid_fac.text = "FIN"


Return
CON22:
SQ_OPER = 1
PUB_CODCIAL = LK_CODCIA
LEER_CAR_LLAVE
If car_llave.EOF = True Then
   msg_err = "DOCUMENTO NO EXISTE..."
   exito = False
   Return
End If

If car_llave!car_fecha_ingr <> LK_FECHA_DIA Then
   msg_err = "NO PUEDE MODIFICAR DOCS... DE OTRO DIA..."
   exito = False
   Return
End If
   
Return




GoTo fin

ACT1:

If pub_signo_arm = 0 Then
   Return
End If

WS_NUMSEC = 0
fila = 1
WS_BRUTO2 = 0
FLAG = False
Do While FLAG = False
   FORMGEN.grid_fac.Row = fila
   FORMGEN.grid_fac.Col = 1
   If Val(FORMGEN.grid_fac.text) = 0 Then
      GoTo OTROMAS
   End If
   
   
   PUB_CODART = FORMGEN.grid_fac.text

   FORMGEN.grid_fac.Col = 8
   PUB_CANTIDAD = Val(FORMGEN.grid_fac.text)
   
   FORMGEN.grid_fac.Col = 3
   WS_UNIDAD = Val(FORMGEN.grid_fac.text)
   
   FORMGEN.grid_fac.Col = 9
   PUB_PRECIO2 = Val(FORMGEN.grid_fac.text)
   
   FORMGEN.grid_fac.Col = 12
   PUB_COSPRO = Val(FORMGEN.grid_fac.text)
   
   
   FORMGEN.grid_fac.Col = 10
   PUB_AUTKEY = Val(FORMGEN.grid_fac.text)
   
   If PUB_AUTKEY = 0 Then
      GoTo SIGUE
   End If
   
   SQ_OPER = 1
   LEER_AUT_LLAVE
   If aut_llave.EOF Then
      msg_err = "ERROR EN AUTORIZACION ..."
      GoTo errorr
   End If
   
   If PUB_AUTKEY <> 0 Then
      aut_llave.Edit
      aut_llave!AUT_NUM_OPER = PUB_NUM_OPER_XXX
      aut_llave!AUT_CANTIDAD = PUB_CANTIDAD
      aut_llave!AUT_estado = "1"
      aut_llave.Update
'      SENDMAIL "alancosme@almacen", "AUTORIZACIONES", "PASE AUTORIZADO.... EJECUTADA"
   End If
   

SIGUE:
If LK_CODTRA <> 2401 Then
   GoTo SIGUE2
End If

FORMGEN.grid_fac.Col = 11
PUB_PRECIO_CASH = Val(FORMGEN.grid_fac.text)
If FORMGEN.i_fbg = "F" Or FORMGEN.i_fbg = "B" Then
   PUB_PRECIO_CASH = PUB_PRECIO_CASH / (1 + LK_IGV / 100)
   PUB_PRECIO_CASH = redondea(PUB_PRECIO_CASH)
End If
subtotal = PUB_CANTIDAD * PUB_PRECIO_CASH
subtotal = redondea(subtotal)
WS_BRUTO2 = WS_BRUTO2 + subtotal

SIGUE2:

   PUB_CODCIA = LK_CODCIA
   SQ_OPER = 1
   LEER_ARM_LLAVE
   If arm_llave.EOF Then
      msg_err = "ARTICULO NO EXISTE ..."
'      GoTo errorr
   End If
   
   

      contador = contador + 1
      arm_llave.Edit
      arm_llave!ARM_STOCK = Val(arm_llave!ARM_STOCK) + PUB_CANTIDAD * pub_signo_arm
      If pub_signo_arm = -1 Then
         arm_llave!ARM_SALIDAS = Val(arm_llave!ARM_SALIDAS) + PUB_CANTIDAD
      Else
      If pub_signo_arm = 1 Then
         arm_llave!ARM_INGRESOS = arm_llave!ARM_INGRESOS + PUB_CANTIDAD
      End If
      End If
      arm_llave.Update
      ws_flag_arm = modif

      far_llave.AddNew
      far_llave!far_tipmov = PUB_TIPMOV
      far_llave!FAR_CODCIA = LK_CODCIA
      far_llave!far_numser = PUB_NUMSER
      far_llave!FAR_NUMFAC = PUB_NUMFAC
      WS_NUMSEC = WS_NUMSEC + 1
      far_llave!FAR_NUMSEC = WS_NUMSEC
      far_llave!far_codart = arm_llave!ARM_CODART
      far_llave!FAR_CANTIDAD = PUB_CANTIDAD
      ws_tot_CANTIDAD = PUB_CANTIDAD + ws_tot_CANTIDAD
      
      far_llave!FAR_UNIDAD = WS_UNIDAD
      
      far_llave!FAR_LOTE = PUB_CODALI
      far_llave!FAR_FECHA_LOTE = LK_FECHA_DIA
      far_llave!FAR_NUM_LOTE = PUB_NUM_OPER_XXX
      

      far_llave!FAR_SIGNO_ARM = pub_signo_arm
      far_llave!far_SIGNO_LOT = pub_signo_lot
      
      far_llave!FAR_CODCLIE = PUB_CODCLIE
      far_llave!FAR_CP = PUB_CP
      If far_llave!far_tipmov = 1 Then
         far_llave!far_transito = "T"
       Else
         far_llave!far_transito = " "
       End If
       
      far_llave!far_estado = tra_llave!TRA_FLAG_EXT

      far_llave!far_stock = arm_llave!ARM_STOCK
      far_llave!far_cospro = PUB_COSPRO
      far_llave!far_precio = PUB_PRECIO2
      far_llave!FAR_FBG = PUB_FBG
      far_llave!far_impto = PUB_IMPTO
      far_llave!far_descto = PUB_DESCTO
      far_llave!far_gastos = PUB_GASTOS
      far_llave!far_bruto = PUB_SUBTOTAL
      far_llave!far_INTERES = PUB_INTERES
      far_llave!FAR_NUMDOC = PUB_NUMDOC
      far_llave!FAR_NUMGUIA = PUB_NUMGUIA
      far_llave!FAR_pordeSCTO1 = PUB_PORDESCTO1
      far_llave!FAR_pordescto2 = PUB_PORDESCTO2
      far_llave!FAR_AJUSTE = PUB_AJUSTE
      far_llave!FAR_DIAS = 0
      far_llave!FAR_OTRA_CIA = PUB_CODCIA_R
      far_llave!FAR_FECHA = LK_FECHA_DIA
      far_llave!far_numser_c = PUB_NUMSER_C
      far_llave!far_numfac_c = PUB_NUMFAC_C
      far_llave!FAR_NUMOPER = PUB_NUM_OPER_XXX
      far_llave!FAR_PRECIO_NETO = 0
      far_llave!FAR_CONCEPTO = PUB_CONCEPTO
      
      far_llave.Update
      ws_flag_far = ingre
OTROMAS:
   fila = fila + 1
   FORMGEN.grid_fac.Row = fila
   FORMGEN.grid_fac.Col = 1
   If FORMGEN.grid_fac.text = "FIN" Then
      FLAG = True
   End If
   
Loop

FORMGEN.i_interes.text = Val(FORMGEN.i_subtotal.text) - WS_BRUTO2
FORMGEN.i_subtotal.text = Val(FORMGEN.i_subtotal.text) - Val(FORMGEN.i_interes.text)
PUB_SUBTOTAL = FORMGEN.i_subtotal.text
PUB_INTERES = FORMGEN.i_interes.text




If PUB_TASA_VENTA = 0 Then
   GoTo salta
End If

If PUB_TASA_ORIG = PUB_TASA_VENTA Then
   GoTo salta
End If

      PUB_AUTKEY = PUB_AUTKEY2
      SQ_OPER = 1
      LEER_AUT_LLAVE
      If aut_llave.EOF Then
         msg_err = "ERROR EN AUTORIZACION ..."
         GoTo errorr
      End If

   
  
      aut_llave.Edit
      aut_llave!AUT_NUM_OPER = PUB_NUM_OPER_XXX
      aut_llave!AUT_estado = "1"
      aut_llave.Update
      SENDMAIL "alancosme@almacen", "AUTORIZACIONES", "PASE AUTORIZADO.... EJECUTADA"
  

salta:

    

If PUB_TIPMOV <> 2 Then
   Return
End If


   PU_NUMSER = 0
   PU_NUMSEC = 1
   PU_NUMFAC = Val(FORMGEN.i_numdoc_r.text)
   PU_CODCIA = PUB_CODCIA_R
   PU_TIPMOV = 1
   SQ_OPER = 1
   LEER_FAR_LLAVE
   If far_llave.EOF = True Then
      msg_err = "!!! ERROR EN ACTUALIZACION ..."
      exito = False
   End
   End If
   
WS_DATOS = "SI"
Do Until WS_DATOS = "NO"
   far_llave.Edit
   far_llave!far_transito = ""
   far_llave.Update
   far_llave.MoveNext
   If far_llave.EOF = False Then
   If far_llave!FAR_NUMFAC = PU_NUMFAC And far_llave!far_tipmov = 1 And far_llave!FAR_OTRA_CIA = LK_CODCIA Then
      Print " "
   Else
      WS_DATOS = "NO"
   End If
   End If
   If far_llave.EOF = True Then
      WS_DATOS = "NO"
   End If
   
   
Loop





Return

ACT2:

'If pub_signo_lot = 0 Then
'   Return
'End If
'SI EXOTNRO ES PARA MODIFICAR...'

      lot_llave.Edit
      lot_llave!LOT_PESO_venta = Val(lot_llave!LOT_PESO_venta) + PUB_CANTIDAD * pub_signo_lot
      lot_llave.Update
      ws_flag_lot = modif

Return

ACT3:
If pub_signo_car = 0 Then
   Return
End If

car_llave.AddNew
car_llave!car_codclie = PUB_CODCLIE
car_llave!car_codcia = LK_CODCIA
car_llave!car_numguia = PUB_NUMGUIA
car_llave!CAR_TIPDOC = PUB_TIPDOC
car_llave!car_cp = PUB_CP
car_llave!car_serdoc = PUB_SERDOC
car_llave!car_numdoc = PUB_NUMDOC
car_llave!car_fecha_ingr = LK_FECHA_DIA
car_llave!CAR_fecha_vcto = PUB_FECHA_VCTO
car_llave!car_situacion = PUB_SITUACION_ACT
If PUB_IMPTO <> 0 Then
   car_llave!CAR_NAT_JUR = "J"
Else
   car_llave!CAR_NAT_JUR = "N"
End If

car_llave!CAR_NUM_REN = 0
car_llave!car_concepto = PUB_CONCEPTO

car_llave!car_int_8dias = 0
car_llave!car_GRUPO = 0
If PUB_TIPMOV <> 0 Then
   car_llave!CAR_IMPORTE = PUB_NETO
Else
   car_llave!CAR_IMPORTE = PUB_IMPORTE_AMORT
   LOC_SALDO_CAR = PUB_IMPORTE_AMORT
End If

car_llave!CAR_CODTRA = LK_CODTRA
car_llave!CAR_SIGNO_CAR = pub_signo_car

car_llave.Update
ws_flag_car = ingre
Return

ACT4:


Return

ACT5:
If pub_signo_ccm = 0 Then
   Return
End If

ccm_llave.Edit
ccm_llave!CCM_SALDO = ccm_llave!CCM_SALDO + PUB_IMPORTE * pub_signo_ccm
ccm_llave.Update
ws_flag_ccm = modif

Return

ACT6:
'REVISAR....

'If LK_CODTRA = 1111 Then
'   If pub_signo_lot = 0 Then
'      Return
'   End If
'   If PUB_FLAG_LOT = ingre Then
'   Else
'   IF
'      lot_llave.Delete
'      Return
'   End If
'End If

lot_llave.AddNew
lot_llave!LOT_CODCIA = LK_CODCIA
lot_llave!LOT_CODALI = PUB_CODALI
lot_llave!LOT_FECHA = LK_FECHA_DIA
lot_llave!LOT_num_operac = PUB_NUM_OPER_XXX
lot_llave!LOT_DESCRI = FORMGEN.i_concepto.text
lot_llave!LOT_PESO = ws_tot_CANTIDAD
lot_llave!LOT_PESO_venta = 0
lot_llave!LOT_PRECIO = redondea(WS_BRUTO / ws_tot_CANTIDAD)
far_llave!far_SIGNO_LOT = pub_signo_lot
far_llave!FAR_SIGNO_ARM = pub_signo_arm

'desboradmineot... xxxxxxxxxxxxxxxxxxxxxxxxx
lot_llave.Update
ws_flag_lot = ingre

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

all_menor!ALL_FLAG_ARM = ws_flag_arm
all_menor!ALL_FLAG_che = ws_flag_che
all_menor!ALL_FLAG_FAR = ws_flag_far
all_menor!ALL_FLAG_CAR = ws_flag_car
all_menor!ALL_FLAG_PAR = ws_flag_par
all_menor!ALL_FLAG_CCM = ws_flag_ccm
all_menor!ALL_FLAG_LOT = ws_flag_lot

all_menor!all_codtra_ext = Left(FORMGEN.TRANS.text, 4)
all_menor!ALL_CODALI = PUB_CODALI
all_menor!ALL_FECHA_LOTE = PUB_FECHA_LOTE
all_menor!ALL_NUM_LOTE = PUB_NUM_OPER_XXX


all_menor.Update

Return


ACT8:
If pub_signo_car = 0 Then
   Return
End If


car_llave.Edit
car_llave!CAR_IMPORTE = car_llave!CAR_IMPORTE + PUB_IMPORTE_AMORT * pub_signo_car
LOC_SALDO_CAR = car_llave!CAR_IMPORTE
If FORMGEN.i_fecha_vcto.Visible = True Then
   car_llave!CAR_fecha_vcto = PUB_FECHA_VCTO
End If
If FORMGEN.i_concepto.Visible = True Then
   car_llave!car_concepto = PUB_CONCEPTO
End If

If PUB_IMPORTE_AMORT <> 0 And LK_CODTRA <> 1111 Then
   car_llave!CAR_NUM_REN = car_llave!CAR_NUM_REN + 1
End If

If pub_signo_car <> 0 Then
If car_llave!CAR_IMPORTE = 0 And car_llave!car_int_8dias = 0 Then
   car_llave!car_situacion = 9
End If
End If

If FORMGEN.i_grupo_act.Visible = True Then
   car_llave!car_GRUPO = PUB_GRUPO_ACT
End If

car_llave.Update
ws_flag_car = modif
Return

ACT9:
PUB_NUM = 0
PUB_AUTKEY = 30000
SQ_OPER = 3
LEER_AUT_LLAVE

If aut_menor.EOF = False Then
   aut_menor.MoveLast
   PUB_NUM = aut_menor!aut_key
End If
If PUB_CODART > 1 Then
   DETERMINA_PRECIO
End If
PUB_NUM = PUB_NUM + 1
aut_menor.AddNew
aut_menor!AUT_NUM_OPER = PUB_NUM_OPER_XXX
aut_menor!aut_key = PUB_NUM
aut_menor!AUT_CODUSU = LK_CODUSU
aut_menor!AUT_codtra = LK_CODTRA
aut_menor!aut_codusu_final = PUB_USUARIO
aut_menor!aut_precio = PUB_PRECIO
aut_menor!AUT_SERIE = PUB_NUMSER
aut_menor!aut_numfac = PUB_NUMSER
aut_menor!AUT_estado = ""
aut_menor!AUT_FECHA_INGR = LK_FECHA_DIA
aut_menor!AUT_HORA_INGR = Time
aut_menor!aut_codart = PUB_CODART
aut_menor!AUT_CANTIDAD = PUB_CANTIDAD
aut_menor!AUT_precio_cash = PUB_PRECIO_CASH
aut_menor!AUT_codclie = PUB_CODCLIE

aut_menor!AUT_dias = pub_dias
aut_menor!AUT_TASA_VENTA = PUB_TASA_VENTA
aut_menor.Update

Return

ACT10:
ws_respuesta = vbYes
count = 1
Do While ws_respuesta = vbYes
 WS_FLAG = 0
 FORMGEN.gridl.Row = count
 FORMGEN.gridl.Col = 6
 If IsDate(FORMGEN.gridl.text) = True Then
 If PUB_FECHA_VCTO <> FORMGEN.gridl.text Then
    WS_FLAG = 1
 End If
 End If
 FORMGEN.gridl.Col = 0
 If FORMGEN.gridl.text <> "@" Then
    WS_FLAG = 1
 End If
 
 If WS_FLAG = 0 Then
    FORMGEN.gridl.Col = 15
    PUB_NUMDOC = FORMGEN.gridl.text
    FORMGEN.gridl.Col = 11
    PUB_TIPDOC = FORMGEN.gridl.text
    FORMGEN.gridl.Col = 12
    PUB_CP = FORMGEN.gridl.text
    FORMGEN.gridl.Col = 13
    PUB_CODCLIE = FORMGEN.gridl.text
    FORMGEN.gridl.Col = 14
    PUB_SERDOC = FORMGEN.gridl.text

    FORMGEN.gridl.Col = 1
    PUB_CODCIAL = FORMGEN.gridl.text

    SQ_OPER = 1
    LEER_CAR_LLAVE
    If car_llave.EOF = True Then
       msg_err = "DOCUMENTO NO EXISTE..."
    End
    End If
 End If
 If WS_FLAG = 0 Then
    WS_BRUTO = car_llave!CAR_IMPORTE + WS_BRUTO
    WS_BRUTO2 = car_llave!car_int_8dias + WS_BRUTO2
    car_llave.Edit
    car_llave!car_situacion = 9
    car_llave.Update
 End If
 count = count + 1
 FORMGEN.gridl.Row = count
 FORMGEN.gridl.Col = 12
 PUB_NUMDOC = Val(FORMGEN.gridl.text)
  If FORMGEN.gridl.Rows = FORMGEN.gridl.Row + 1 Then
    ws_respuesta = vbNo
  End If


 Loop
 
car_llave.AddNew
car_llave!car_codclie = PUB_CODCLIE
car_llave!car_codcia = LK_CODCIA
car_llave!CAR_TIPDOC = PUB_TIPDOC
car_llave!car_cp = PUB_CP
par_llave.Edit 'xxxxxxxxx
If PUB_TIPDOC = "LE" Then
   par_llave!PAR_NUM_LETRA = par_llave!PAR_NUM_LETRA + 1
   car_llave!car_numdoc = par_llave!PAR_NUM_LETRA
Else
If PUB_TIPDOC = "FA" Then
   par_llave!PAR_NUM_FACTURA = par_llave!PAR_NUM_FACTURA + 1
   car_llave!car_numdoc = par_llave!PAR_NUM_FACTURA
End If
End If
par_llave.Update

car_llave!car_fecha_ingr = LK_FECHA_DIA
car_llave!CAR_fecha_vcto = PUB_FECHA_VCTO
car_llave!car_fecha_vcto_orig = PUB_FECHA_VCTO
car_llave!car_situacion = 2
car_llave!CAR_NAT_JUR = "N"
'car_llave!CAR_FECHA_ULT_RENOV = 0
'car_llave!CAR_FECHA_ULT_RENOV_BAK = 0
car_llave!CAR_NUM_REN = 0
'car_llave!CAR_FECHA_PROT = 0
car_llave!car_int_8dias = WS_BRUTO2
car_llave!car_GRUPO = 0
car_llave!CAR_IMPORTE = WS_BRUTO
LOC_SALDO_CAR = car_llave!CAR_IMPORTE
'car_llave!CAR_IMPORTE_INI = WS_BRUTO
car_llave.Update

 
Return

ACT11:
If pub_signo_ccm = 0 Then
   Return
End If
If PUB_CHENUM = 0 Or LK_CODTRA = 1111 Or PUB_CHESEC <> 0 Then
   che_llave.AddNew
   che_llave!CHE_CODBAN = PUB_CODBAN
   che_llave!CHE_CODCIA = PUB_CODCIA
   che_llave!CHE_CHESER = PUB_CHESER
   che_llave!CHE_CHESEC = PUB_CHESEC
   che_llave!CHE_CHENUM = PUB_CHENUM
   che_llave!che_importe = PUB_IMPORTE
   che_llave!che_abreviado = PUB_ABREVIADO
   che_llave!che_fecha = LK_FECHA_DIA
   che_llave!CHE_CODUSU = LK_CODUSU
   che_llave!che_CONCEPTO = PUB_CONCEPTO
   che_llave!CHE_NUMOPER = PUB_NUM_OPER_XXX
   che_llave!che_saldo = ccm_llave!CCM_SALDO
   che_llave!CHE_SIGNO_CCM = pub_signo_ccm
   che_llave!CHE_CODTRA = LK_CODTRA
   che_llave!che_chenum_ext = PUB_CHENUM_EXT
   che_llave!che_estado = " "
   che_llave.Update
   ws_flag_che = ingre
 Else
   che_llave.Edit
   che_llave!che_importe = PUB_IMPORTE
   che_llave!che_fecha = LK_FECHA_DIA
   che_llave!CHE_CODUSU = LK_CODUSU
   che_llave!che_CONCEPTO = PUB_CONCEPTO
   che_llave!che_abreviado = PUB_ABREVIADO
   che_llave!CHE_NUMOPER = PUB_NUM_OPER_XXX
   che_llave!che_saldo = ccm_llave!CCM_SALDO
   che_llave!CHE_SIGNO_CCM = pub_signo_ccm

   che_llave.Update
   ws_flag_che = modif
End If

Return
ACT12:
cli_llave.Edit
cli_llave!cli_tipo_bloq1 = PUB_TIPO_BLOQ_act1
cli_llave!cli_tipo_bloq2 = PUB_TIPO_BLOQ_act2
cli_llave!cli_tipo_bloq3 = PUB_TIPO_BLOQ_act3
cli_llave!cli_tipo_bloq4 = PUB_TIPO_BLOQ_act4

If cli_llave!CLI_limcre <> PUB_LIMCRE_ACT Then
   cli_llave!CLI_fecha_aprob = LK_FECHA_DIA
End If
cli_llave!CLI_limcre = PUB_LIMCRE_ACT
cli_llave.Update
Return

ACT13:
      
      far_llave.AddNew
      far_llave!far_tipmov = PUB_TIPMOV
      far_llave!FAR_CODCIA = LK_CODCIA
      far_llave!far_numser = 0
      far_llave!FAR_NUMFAC = PUB_NUMFAC
      far_llave!FAR_NUMSEC = PUB_NUMSER '900
      far_llave!far_codart = 0
      far_llave!FAR_CANTIDAD = PUB_CANTIDAD
'      far_llave!FAR_SIGNO = 0
      far_llave!FAR_CODCLIE = PUB_CODCLIE
      far_llave!FAR_CP = PUB_CP
      far_llave!far_transito = " "
      far_llave!far_stock = 0
      far_llave!far_cospro = 0
      far_llave!far_precio = PUB_PRECIO
      far_llave!FAR_FBG = " "
      far_llave!far_impto = 0
      far_llave!far_descto = 0
      far_llave!far_gastos = 0
      far_llave!far_bruto = 0
      far_llave!FAR_NUMDOC = 0
      far_llave!FAR_NUMGUIA = 0
      far_llave!FAR_pordeSCTO1 = 0
      far_llave!FAR_pordescto2 = 0
      far_llave!FAR_AJUSTE = 0
      far_llave!FAR_DIAS = 0
      far_llave!FAR_OTRA_CIA = " "
      far_llave!FAR_FECHA = LK_FECHA_DIA
      far_llave!far_numser_c = 0
      far_llave!far_numfac_c = 0
      far_llave!FAR_NUMOPER = PUB_NUM_OPER_XXX
      far_llave!far_tipo_bloq_act1 = PUB_TIPO_BLOQ_act1
      far_llave!far_tipo_bloq_act2 = PUB_TIPO_BLOQ_act2
      far_llave!far_tipo_bloq_act3 = PUB_TIPO_BLOQ_act3
      far_llave!far_tipo_bloq_act4 = PUB_TIPO_BLOQ_act4
      far_llave!far_tipo_bloq_ant1 = PUB_TIPO_BLOQ_ant1
      far_llave!far_tipo_bloq_ant2 = PUB_TIPO_BLOQ_ant2
      far_llave!far_tipo_bloq_ant3 = PUB_TIPO_BLOQ_ant3
      far_llave!far_tipo_bloq_ant4 = PUB_TIPO_BLOQ_ant4
      far_llave!far_limcre_act = PUB_LIMCRE_ACT
      far_llave!far_limcre_ant = PUB_LIMCRE_ANT
      
      far_llave!FAR_LOTE = PUB_CODALI
      far_llave!FAR_FECHA_LOTE = PUB_FECHA_LOTE
      far_llave!FAR_NUM_LOTE = PUB_NUM_LOTE
      far_llave!FAR_CONCEPTO = PUB_CONCEPTO
      far_llave!far_SIGNO_LOT = pub_signo_lot
      far_llave!FAR_SIGNO_ARM = pub_signo_arm
  
      far_llave.Update

Return
ACT14:
far_llave.Edit
far_llave!far_numfac_c = PUB_NUMFAC_C
far_llave!far_numser_c = PUB_NUMSER_C
far_llave.Update
Boton_Compras_Click
Return

ACT15:
Return

ACT16:
If PUB_TIPMOV = 0 Then
   Return
End If
SQ_OPER = 3
PU_TIPMOV = PUB_TIPMOV
PU_CODCIA = PUB_CODCIA
PU_NUMSER = 999
PUB_NUMSER = 999
LEER_FAR_LLAVE
If far_menor.EOF = False Then
   far_menor.MoveLast
End If
PU_NUMFAC = 1
If Not far_menor.EOF Then
   PU_NUMFAC = far_menor!FAR_NUMFAC + 1
End If
PUB_NUMFAC = PU_NUMFAC
far_llave.MoveFirst
fila = 0
'pasa 2 veces, pero no importa
' SE CAMBIA EL NUMFAC POR EL 888
Do Until far_llave.EOF
   far_llave.Edit
   far_llave!far_numser = 888
   far_llave!FAR_NUMFAC = PU_NUMFAC
   far_llave.Update
   far_llave.MoveNext
Loop

Return

ACT17:
contador = 0
Do Until contador = PUB_CANT_CHEQ
che_llave.AddNew
che_llave!CHE_CODBAN = PUB_CODBAN
che_llave!CHE_CODCIA = LK_CODCIA
che_llave!CHE_CHESER = PUB_CHESER
che_llave!CHE_CHESEC = 0
che_llave!CHE_CHENUM = contador + PUB_NUM_INI
contador = contador + 1
che_llave!che_importe = 0
che_llave!che_fecha = 0
che_llave!CHE_CODUSU = LK_CODUSU
che_llave!che_CONCEPTO = " "
che_llave!CHE_NUMOPER = contador + 30000
che_llave!che_saldo = 0
che_llave!CHE_SIGNO_CCM = 0
che_llave!che_estado = " "
che_llave!che_abreviado = ""
che_llave!CHE_CODTRA = LK_CODTRA
che_llave!che_chenum_ext = 0

che_llave.Update
Loop
Return

ACT18:
PUB_NUM_OPER = PUB_NUM_OPER_EXT
SQ_OPER = 1
LEER_ALL_LLAVE
If all_llave.EOF Then
   msg_err = "Numero de Operacion Incorrecto...."
   exito = False
   Return
End If

all_llave.Edit
all_llave!all_flag_ext = "E"
all_llave.Update

If pub_signo_ccm <> 0 Then
   SQ_OPER = 2
   PUB_FECHA = LK_FECHA_DIA
   LEER_CHE_LLAVE
   If che_oper.EOF = True Then
      msg_err = "Error en Bancos..no existe tal cuenta..."
      GoTo errorr
   End If
   che_oper.Edit
   che_oper!che_estado = "E"
   che_oper.Update
   PUB_CHESEC = che_oper!CHE_CHESEC + 50
   PUB_CHENUM_EXT = PUB_CHENUM
   PUB_CHENUM = 0
   PUB_ABREVIADO = "EXT"
End If


Return

REP1:
Unload frmeditor

'*** Imprime la Transacción a un Archivo
Open LK_CODUSU For Output As #1
Cabe_Trans
Printer.FontSize = 6
Printer.FontName = "Arial"
Printer.FontBold = False
If FORMGEN.Frame4.Visible = True Then
   Repo_Grid
End If

Close #1
'** CARGA EL EDITOR **
Screen.MousePointer = 0
frmeditor.Show 1
Load frmeditor
Return

REP2:

Return

REP3:

PUB_TIPREG = 20
PUB_NUMTAB = cli_llave!CLI_CASA_ZONA
SQ_OPER = 1
LEER_TAB_LLAVE
If tab_llave.EOF = True Then
   msg_err = "ERROR EN ZONA.."
End If
'zona_nombre = tab_llave!TAB_NOMLARGO

PUB_TIPREG = 30
PUB_NUMTAB = cli_llave!CLI_CASA_SUBZONA
SQ_OPER = 1
LEER_TAB_LLAVE
'subzona_nombre = tab_llave!TAB_NOMLARGO


'*** REPORTE DE LETRA ***
IMP_LETRA CStr(WS_IMPRESION_LET), CStr(PUB_FECHA_VCTO), CStr(PUB_NETO), CStr(PUB_FECHA), CStr(PUB_FECHA_VCTO), CStr(PUB_NETO), cli_llave!CLI_NOMBRE, Nulo_Valors(cli_llave!CLI_CASA_DIREC), cli_llave!CLI_CASA_NUM, zona_nombre, subzona_nombre

Return
   

Screen.MousePointer = 1
'CLEAR_GRID
'frame4.Visible = False
'Frame1.Visible = False






GoTo fin





Error_fatal:
    msg = "Se ha producido un error " & "al abrir la conexión:" & Err & " - " & Error & vbCr
    For Each er In rdoErrors
        msg = msg & er.Description & ":" & er.Number & vbCr
        MsgBox msg
    Next er

    CN.Execute "Rollback Transaction", rdExecDirect

    
'    Resume AbandonCn
errorr:

fin:
Screen.MousePointer = 0



End Sub

Private Sub Grid_all_LostFocus()
FORMGEN.Grid_all.Visible = False

End Sub

Private Sub Grid_all_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If LK_CODTRA <> 1111 Then
   GoTo fin
End If
If Button <> 2 Then
   GoTo fin
End If



   Grid_all.Col = 10
   If Grid_all.text = "E" Then
      MsgBox "Ya Extornado..."
      GoTo fin
   End If
   
WS_CONTADOR = 1
Do Until WS_CONTADOR > Grid_all.Cols - 1
   Grid_all.Col = WS_CONTADOR
   FORMGEN.Grid_all.CellBackColor = vbRed
   WS_CONTADOR = WS_CONTADOR + 1
Loop

   
   
   
PUB_MENSAJE = MsgBox("Esta seguro de Extornar", 36, WS_TITULO)

If PUB_MENSAJE = vbNo Then
WS_CONTADOR = 1
Do Until WS_CONTADOR > Grid_all.Cols - 1
   Grid_all.Col = WS_CONTADOR
   FORMGEN.Grid_all.CellBackColor = vbWhite
   WS_CONTADOR = WS_CONTADOR + 1
Loop
GoTo fin
End If



Grid_all.Col = 0
FORMGEN.i_num_oper.text = Grid_all.text
NUMERO = FORMGEN.Diario.TabIndex
avanza_campo

fin:

End Sub

Private Sub grid_autorizacion_GotFocus()
Dim cambiar_color As Boolean
Dim Tit As String
Tit = FORMGEN.grid_autorizacion.FormatString
grid_autorizacion.Clear
FORMGEN.grid_autorizacion.FormatString = Tit

PUB_AUTKEY = 30000
SQ_OPER = 3
LEER_AUT_LLAVE
If aut_menor.EOF = True Then
   MsgBox "Lo siento ... no hay autorizaciones"
   grid_autorizacion.Visible = False
   GoTo fin
End If
aut_menor.MoveLast
fila = 1
Do Until aut_menor.BOF Or fila = 20
cambiar_color = False
If Nulo_Valors(aut_menor!AUT_estado) = "1" Then
   cambiar_color = True
End If

grid_autorizacion.Row = fila
grid_autorizacion.Col = 1
If cambiar_color = True Then
   FORMGEN.grid_autorizacion.CellBackColor = vb3DLight
End If
grid_autorizacion.text = aut_menor!aut_key
grid_autorizacion.Col = 2
If cambiar_color = True Then
   FORMGEN.grid_autorizacion.CellBackColor = vb3DLight
End If

grid_autorizacion.text = aut_menor!aut_codusu_final
grid_autorizacion.Col = 3
If cambiar_color = True Then
   FORMGEN.grid_autorizacion.CellBackColor = vb3DLight
End If


grid_autorizacion.text = aut_menor!aut_codart
grid_autorizacion.Col = 4
If cambiar_color = True Then
   FORMGEN.grid_autorizacion.CellBackColor = vb3DLight
End If


grid_autorizacion.text = aut_menor!aut_precio
grid_autorizacion.Col = 5
If cambiar_color = True Then
   FORMGEN.grid_autorizacion.CellBackColor = vb3DLight
End If


grid_autorizacion.text = aut_menor!aut_numfac
grid_autorizacion.Col = 6
If cambiar_color = True Then
   FORMGEN.grid_autorizacion.CellBackColor = vb3DLight
End If

grid_autorizacion.text = aut_menor!AUT_CANTIDAD
grid_autorizacion.Col = 7
If cambiar_color = True Then
   FORMGEN.grid_autorizacion.CellBackColor = vb3DLight
End If

grid_autorizacion.text = aut_menor!AUT_TASA_VENTA

grid_autorizacion.Col = 8
If cambiar_color = True Then
   FORMGEN.grid_autorizacion.CellBackColor = vb3DLight
End If
grid_autorizacion.text = aut_menor!AUT_dias

grid_autorizacion.Col = 9
If cambiar_color = True Then
   FORMGEN.grid_autorizacion.CellBackColor = vb3DLight
End If
grid_autorizacion.text = Nulo_Valor0(aut_menor!AUT_precio_cash)


grid_autorizacion.Col = 10
If cambiar_color = True Then
   FORMGEN.grid_autorizacion.CellBackColor = vb3DLight
End If
grid_autorizacion.text = Nulo_Valors(aut_menor!AUT_estado)

grid_autorizacion.Col = 11
If cambiar_color = True Then
   FORMGEN.grid_autorizacion.CellBackColor = vb3DLight
End If
grid_autorizacion.text = aut_menor!AUT_codtra

grid_autorizacion.Col = 12
If cambiar_color = True Then
   FORMGEN.grid_autorizacion.CellBackColor = vb3DLight
End If
grid_autorizacion.text = aut_menor!AUT_FECHA_INGR

grid_autorizacion.Col = 13
If cambiar_color = True Then
   FORMGEN.grid_autorizacion.CellBackColor = vb3DLight
End If
grid_autorizacion.text = Nulo_Valor0(aut_menor!AUT_HORA_INGR)

aut_menor.MovePrevious
fila = fila + 1
Loop
fin:
End Sub

Private Sub grid_autorizacion_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> 13 Then
   GoTo fin
End If

If boton_autorizacion.Visible = True Then
   GoTo fin
End If


grid_autorizacion.Col = 1
If Val(FORMGEN.grid_autorizacion.text) = 0 Then
      GoTo fin
End If


grid_autorizacion.Col = 10
If FORMGEN.grid_autorizacion.text = "1" Then
   MsgBox "Autorizacion ya Ejecutada..."
   GoTo fin
End If

grid_autorizacion.Col = 12
If (FORMGEN.grid_autorizacion.text) <> LK_FECHA_DIA Then
   MsgBox "Autorizacion ya Pasada de fecha.."
   GoTo fin
End If


PUB_AUTKEY2 = 0
grid_autorizacion.Col = 11
If FORMGEN.i_dias.Enabled = False And Val(FORMGEN.grid_autorizacion.text) = 2003 Then
   FORMGEN.i_autorizacion.Visible = True
   FORMGEN.Autoriz.Visible = True
   grid_autorizacion.Col = 1

   FORMGEN.i_autorizacion.text = grid_autorizacion.text
   FORMGEN.i_codart2.SetFocus
   GoTo fin
End If

If FORMGEN.i_dias.Enabled = False Then
   GoTo fin
End If


grid_autorizacion.Col = 1
PUB_AUTKEY = Val(FORMGEN.grid_autorizacion.text)
PUB_AUTKEY2 = PUB_AUTKEY
grid_autorizacion.Col = 11
If Val(FORMGEN.grid_autorizacion.text) <> 2004 Then
   GoTo fin
End If




SQ_OPER = 1

LEER_AUT_LLAVE
If aut_llave.EOF = True Then
   MsgBox "Numero de Autorizacion errado"
   GoTo fin
End If


FORMGEN.i_tasa_venta.text = aut_llave!AUT_TASA_VENTA


FORMGEN.i_tipdoc.Enabled = False
FORMGEN.i_dias.Enabled = False
FORMGEN.i_fecha_vcto.Enabled = False
FORMGEN.i_pordescto1.Enabled = False
FORMGEN.i_pordescto2.Enabled = False

FORMGEN.i_image_llave.Visible = False
FORMGEN.i_subtotal.SetFocus

fin:

End Sub

Private Sub grid_autorizacion_LostFocus()
grid_autorizacion.Visible = False
End Sub

Private Sub Grid_che_LostFocus()
grid_che.Visible = False
End Sub



Private Sub Grid_LOT_KeyPress(KeyAscii As Integer)
FORMGEN.Grid_lot.Col = 5
FORMGEN.i_num_lote = FORMGEN.Grid_lot.text
FORMGEN.Grid_lot.Col = 1
FORMGEN.i_fecha_lote.text = FORMGEN.Grid_lot.text
FORMGEN.Grid_lot.Col = 4
FORMGEN.i_precio.text = FORMGEN.Grid_lot.text
FORMGEN.Grid_lot.Col = 2
FORMGEN.i_cantidad.text = FORMGEN.Grid_lot.text
FORMGEN.Grid_lot.Visible = False

FORMGEN.i_cantidad.SetFocus



End Sub

Private Sub Grid_lot_LostFocus()
Grid_lot.Visible = False
End Sub

Private Sub i_cheser_LostFocus()
   PUB_CHESEC = 0
   PUB_CODBAN = FORMGEN.i_codban.text
   PUB_CHESER = FORMGEN.i_cheser.text
   SQ_OPER = 3
   PUB_FECHA = 0
   LEER_CHE_LLAVE
   If che_menor.EOF = False Then
      FORMGEN.i_chenum.text = che_menor!CHE_CHENUM
   End If

End Sub

Private Sub Lotes_Click()
Dim WS_SALDO As Currency
Dim Tit As String
Dim success%

FORMGEN.Grid_lot.Top = 2500
FORMGEN.Grid_lot.Left = 300
FORMGEN.Grid_lot.ColWidth(0) = 1800
FORMGEN.Grid_lot.ColWidth(1) = 2000
FORMGEN.Grid_lot.ColWidth(2) = 1500
FORMGEN.Grid_lot.ColWidth(4) = 1500
FORMGEN.Grid_lot.ColWidth(5) = 1500

FORMGEN.Grid_lot.Width = 8400
FORMGEN.Grid_lot.Height = 2500



Screen.MousePointer = 11
fila = 0
PUB_FECHA_LOTE = #1/1/1900#
PUB_CODALI = FORMGEN.i_codali.text
PUB_CODCIA = LK_CODCIA
SQ_OPER = 2
LEER_LOT_LLAVE
If lot_mayor.EOF = True Then
   Screen.MousePointer = 11
   MsgBox "No hay lotes..."
   GoTo fin
End If
success% = SetWindowPos(FrmProcesa.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
FrmProcesa.Show
MUESTRA_PROCESO "Un Momento ..."
Screen.MousePointer = 11
lot_mayor.MoveLast
FORMGEN.Grid_lot.Visible = False
Tit = FORMGEN.Grid_lot.FormatString
Grid_lot.Clear
FORMGEN.Grid_lot.FormatString = Tit
Do Until lot_mayor.BOF
       fila = fila + 1
       FORMGEN.Grid_lot.Rows = fila + 1
       FORMGEN.Grid_lot.Row = fila
       FORMGEN.Grid_lot.Col = 0
       FORMGEN.Grid_lot.text = lot_mayor!LOT_CODALI
       FORMGEN.Grid_lot.Col = 1
       FORMGEN.Grid_lot.text = lot_mayor!LOT_FECHA
       FORMGEN.Grid_lot.Col = 2
       FORMGEN.Grid_lot.text = lot_mayor!LOT_PESO
       FORMGEN.Grid_lot.Col = 3
       FORMGEN.Grid_lot.text = lot_mayor!LOT_PESO_venta
       FORMGEN.Grid_lot.Col = 4
       FORMGEN.Grid_lot.text = lot_mayor!LOT_PRECIO
       FORMGEN.Grid_lot.Col = 5
       FORMGEN.Grid_lot.text = lot_mayor!LOT_num_operac
       
       lot_mayor.MovePrevious
Loop
FrmProcesa.Hide
Screen.MousePointer = 0
FORMGEN.Grid_lot.Visible = True
FORMGEN.Grid_lot.SetFocus

fin:


End Sub

Private Sub GRID1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
  FORMGEN.Frame1.Visible = False
End If
If KeyCode = 13 Then
   Print ""
Else
   GoTo fin
End If

'If numarchi = 1 Then
'   FORMGEN.grid1.Col = 4
'   PUB_CODCLIE = Val(FORMGEN.grid1.text)
'   If PUB_CODCLIE <> 0 Then
'      FORMGEN.i_codcli.text = PUB_CODCLIE
'      NUMERO = FORMGEN.i_codcli.TabIndex
'      SQ_OPER = 1
'      PUB_CODCIA = LK_CODCIA
'      LEER_CLI_LLAVE
'      GoTo FIN
'   End If
'End If

Select Case numarchi
Case 1
 FORMGEN.Grid1.Col = 1
 FORMGEN.i_nomcli.Caption = FORMGEN.Grid1.text
 FORMGEN.Grid1.Col = 2
 FORMGEN.i_codcli.text = FORMGEN.Grid1.text
 PUB_CODCLIE = Val(FORMGEN.i_codcli.text)
 NUMERO = FORMGEN.i_codcli.TabIndex
 TEXTOX(FORMGEN.i_codcli.TabIndex) = FORMGEN.i_codcli.text
 ETIQUETAX(FORMGEN.i_codcli.TabIndex) = FORMGEN.LABELGEN(FORMGEN.i_codcli.TabIndex).Caption
 NOMBREX(FORMGEN.i_codcli.TabIndex) = FORMGEN.i_nomcli.Caption
 SQ_OPER = 1
 PUB_CODCIA = LK_CODCIA
 LEER_CLI_LLAVE
 vuelca_datos
 Case 2
 FORMGEN.Grid1.Col = 1
 FORMGEN.i_nomven.Caption = FORMGEN.Grid1.text
 FORMGEN.Grid1.Col = 2
 FORMGEN.i_codven.text = FORMGEN.Grid1.text
 PUB_CODVEN = FORMGEN.i_codven.text
 NUMERO = FORMGEN.i_codven.TabIndex
 TEXTOX(FORMGEN.i_codven.TabIndex) = FORMGEN.i_codven.text
 ETIQUETAX(FORMGEN.i_codven.TabIndex) = FORMGEN.LABELGEN(FORMGEN.i_codven.TabIndex).Caption
 NOMBREX(FORMGEN.i_codven.TabIndex) = FORMGEN.i_nomven.Caption
 
 Case 3
 FORMGEN.Grid1.Col = 2
 FORMGEN.TRANS.text = FORMGEN.Grid1.text
 NUMERO = 0
 Case 4

 FORMGEN.Grid1.Col = 2
 FORMGEN.i_codart2.text = FORMGEN.Grid1.text
 SQ_OPER = 1
 PUB_KEY = Val(FORMGEN.i_codart2.text)
 LEER_ART_LLAVE
 If art_llave.EOF Then
    MsgBox "ARTICULO NO EXISTE..."
    FORMGEN.LEIDO.SetFocus
    GoTo fin
 End If
  
 i_nomart2.Caption = art_llave!art_nombre
 If FORMGEN.i_codart.Visible = True Then
    FORMGEN.i_codart.text = PUB_KEY
    FORMGEN.i_nomart.Caption = art_llave!art_nombre
    NUMERO = FORMGEN.i_codart.TabIndex
    GoTo SIGUE
 End If
 PUB_COSPRO = art_llave!ART_COSPRO
 BUSCAR_ARM
 GoTo fin
 Case 7
 FORMGEN.Grid1.Col = 1
 FORMGEN.i_nomban.Caption = FORMGEN.Grid1.text
 FORMGEN.Grid1.Col = 2
 FORMGEN.i_codban.text = FORMGEN.Grid1.text
 PUB_CODBAN = Val(FORMGEN.i_codban.text)
 NUMERO = FORMGEN.i_codban.TabIndex

 
 
End Select
SIGUE:
avanza_campo
FORMGEN.Frame1.Visible = False

fin:
End Sub
Private Sub grid_fac_KeyPress(KeyAscii As Integer)
Dim a As Integer
If KeyAscii <> 13 Then
   GoTo fin
End If
 
fila = grid_fac.Row
i_cuenta.text = grid_fac.TextMatrix(fila, 1)
i_glosa.text = grid_fac.TextMatrix(fila, 3)
If grid_fac.TextMatrix(fila, 3) = "D" Then
   i_d_h.ListIndex = 0
   i_importe.text = grid_fac.TextMatrix(fila, 4)
Else
   i_d_h.ListIndex = 1
   i_importe.text = grid_fac.TextMatrix(fila, 5)
End If


   If grid_fac.CellBackColor = vbRed Then
      For a = 0 To grid_fac.Cols - 1
        grid_fac.Col = a
        grid_fac.CellBackColor = vbWhite
      Next a
    Else
      For a = 0 To grid_fac.Cols - 1
        grid_fac.Col = a
        grid_fac.CellBackColor = vbRed
      Next a
    End If

PSCOV_LLAVE.rdoParameters(0) = LK_CODCIA
PSCOV_LLAVE.rdoParameters(1) = grid_fac.TextMatrix(fila, 6)
cov_llave.Requery


fin:
End Sub

Private Sub grid_fac_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode <> 32 Then
   GoTo fin
End If
Dim a As Integer
   If grid_fac.CellBackColor = vbRed Then
      For a = 0 To grid_fac.Cols - 1
        grid_fac.Col = a
        grid_fac.CellBackColor = vbWhite
      Next a
    Else
      For a = 0 To grid_fac.Cols - 1
        grid_fac.Col = a
        grid_fac.CellBackColor = vbRed
      Next a
    End If
fin:
End Sub

Private Sub grid_far_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If
 FORMGEN.grid_far.Col = 1
 If FORMGEN.grid_far.text = "" Then
    GoTo fin
 End If
 If Not IsDate(FORMGEN.grid_far.text) Then
    GoTo fin
 End If
 
 FORMGEN.grid_far.Col = 10
 PUB_TIPMOV = FORMGEN.grid_far.text
 FORMGEN.grid_far.Col = 11
 PUB_CODCIA = FORMGEN.grid_far.text
 FORMGEN.grid_far.Col = 2
 PUB_NUMSER = FORMGEN.grid_far.text
 FORMGEN.grid_far.Col = 3
 PUB_NUMFAC = FORMGEN.grid_far.text
 FORMGEN.grid_far.Col = 6
 PUB_NUMGUIA = FORMGEN.grid_far.text

 PUB_NUMSEC = 1
 
 FORMGEN.i_numfac.Clear
 FORMGEN.i_numfac.AddItem PUB_NUMFAC
 FORMGEN.i_numfac.ListIndex = 0

 FORMGEN.i_numser.text = PUB_NUMSER
 FORMGEN.i_numguia.text = PUB_NUMGUIA
 
 
 PU_TIPMOV = PUB_TIPMOV
 PU_CODCIA = PUB_CODCIA
 PU_NUMSER = PUB_NUMSER
 PU_NUMFAC = PUB_NUMFAC
 PU_NUMSEC = PUB_NUMSEC
 
 SQ_OPER = 1
 LEER_FAR_LLAVE
 If far_llave.EOF = True Then
    MsgBox "DOCUMENTO NO EXISTE..."
   End
 End If
 NUMERO = FORMGEN.grid_far.TabIndex
 avanza_campo
  
fin:


End Sub

Private Sub gridl_KeyPress(KeyAscii As Integer)
Dim WS_IND2, WS_IND3 As Integer
If KeyAscii <> 13 Then
   GoTo fin
End If


 FORMGEN.gridl.Col = 1
 If FORMGEN.gridl.text = "" Then
    GoTo fin
 End If
 If Not IsNumeric(FORMGEN.gridl.text) Then
    GoTo fin
 End If
 
FORMGEN.gridl.Col = 2
If Len(Trim(PUB_TIPDOC)) <> 0 Then
If PUB_TIPDOC = FORMGEN.gridl.text Then
Else
   MsgBox "Tipo de documento no corresponde ..."
   GoTo fin
End If
End If
 

 FORMGEN.gridl.Col = 14
 PUB_SERDOC = FORMGEN.gridl.text
 FORMGEN.gridl.Col = 15
 PUB_NUMDOC = FORMGEN.gridl.text
 
 FORMGEN.i_numdoc.text = PUB_NUMDOC
 FORMGEN.gridl.Col = 2
' estos datos no deben cambiar ...se toman del defcont....
 PUB_TIPDOC = FORMGEN.gridl.text
 FORMGEN.gridl.Col = 12
 PUB_CP = FORMGEN.gridl.text
 FORMGEN.gridl.Col = 13
 PUB_CODCLIE = FORMGEN.gridl.text
 FORMGEN.gridl.Col = 1
 PUB_CODCIAL = FORMGEN.gridl.text

 SQ_OPER = 1
 LEER_CAR_LLAVE
 If car_llave.EOF = True Then
    MsgBox "DOCUMENTO NO EXISTE..."
   End
 End If



vuelca_datos_sit
FORMGEN.i_fecha_vcto.text = car_llave!CAR_fecha_vcto
FORMGEN.i_dias.text = DateDiff("d", car_llave!CAR_fecha_vcto, LK_FECHA_DIA)
If FORMGEN.i_dias.text > 33000 Then
   FORMGEN.i_dias.text = 0
End If
 
FORMGEN.imp_orig.text = car_llave!CAR_IMPORTE
FORMGEN.fecha_orig.text = car_llave!CAR_fecha_vcto
FORMGEN.i_concepto.text = RTrim(car_llave!car_concepto)

 
grupol.Visible = True
grupol.Top = 4500
grupol.Left = 100
 
NUMERO = FORMGEN.Boton_Letras.TabIndex
avanza_campo
  
  
fin:

End Sub

Private Sub GRIDL_LostFocus()
FORMGEN.gridl.Visible = False
End Sub

Private Sub i_cant_cheq_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If


NUMERO = FORMGEN.i_cant_cheq.TabIndex

avanza_campo
fin:

End Sub

Private Sub i_cantidad_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If


NUMERO = FORMGEN.i_cantidad.TabIndex

avanza_campo
fin:

End Sub

Private Sub i_chenum_GotFocus()

Azul i_chenum, i_chenum
End Sub

Private Sub i_chenum_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If


NUMERO = FORMGEN.i_chenum.TabIndex

avanza_campo
fin:


End Sub

Private Sub i_cheser_GotFocus()
Azul i_cheser, i_cheser
End Sub

Private Sub i_cheser_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If


NUMERO = FORMGEN.i_cheser.TabIndex

avanza_campo
fin:

End Sub

Private Sub i_cias_GotFocus()
MUESTRA_PUNTOS
FORMGEN.i_cias.ListIndex = 0

End Sub


Private Sub i_cias_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If
PUB_CODCIA_R = Right(i_cias.text, 2)
PUB_NOMCIA = Left(i_cias.text, 20)

NUMERO = FORMGEN.i_cias.TabIndex
TEXTOX(FORMGEN.i_cias.TabIndex) = FORMGEN.i_cias.text
ETIQUETAX(FORMGEN.i_cias.TabIndex) = FORMGEN.LABELGEN(FORMGEN.i_cias.TabIndex).Caption
NOMBREX(FORMGEN.i_cias.TabIndex) = " "

avanza_campo
fin:


End Sub


Private Sub i_codali_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If
NUMERO = FORMGEN.i_codali.TabIndex

avanza_campo
fin:

End Sub

Private Sub i_codart_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer
PUB_AUTKEY = 0
If KeyAscii <> 13 Then
     GoTo fin
  End If
tf = decimal1(FORMGEN.i_codart)
If tf = False Then
      MsgBox "DATO DEBE SER NUMERICO ", 64
      GoTo fin
   End If
            
 If i_codart.text = "" Then
   UNICO = ""
   numarchi = 4
   archi = "SELECT * FROM ARTI WHERE ART_NOMBRE >= ? ORDER BY ART_NOMBRE"
   PROCESA_GRID (archi)
 Else
    SQ_OPER = 1
    PUB_KEY = FORMGEN.i_codart.text
    LEER_ART_LLAVE
    If art_llave.EOF Then
       MsgBox "ARTI NO EXISTE ..."
       GoTo fin
    Else
       FORMGEN.i_nomart.Caption = art_llave(2)
       NUMERO = FORMGEN.i_codart.TabIndex
       avanza_campo
    End If
 End If

fin:
End Sub
Private Sub i_codart2_GotFocus()

Dim WS_DATOS As String

Dim subtotal As Currency
If LK_CODTRA <> 2403 Then
   GoTo fin
End If



If PUB_TIPMOV <> 2 Then
   GoTo fin
End If

   PU_NUMSER = 0
   PU_NUMSEC = 1
   PU_NUMFAC = Val(FORMGEN.i_numdoc_r)
   PU_CODCIA = PUB_CODCIA_R
   PU_TIPMOV = 1
   SQ_OPER = 1
   LEER_FAR_LLAVE
   If far_llave.EOF = True Then
   MsgBox "!!! NO HAY ESE DOCUMENTO EN TRANSITO..."
   exito = False
   GoTo fin
   End If
   
   
   
   If far_llave!far_transito = "T" Then
   Else
   MsgBox "!!!FALTA LA MARCA DE TRANSITO..."
   exito = False
   GoTo fin
   End If
   
   If far_llave!FAR_OTRA_CIA <> LK_CODCIA Then
   MsgBox "!!! NO TE CORRESPONDE..."
   exito = False
   GoTo fin
   End If
FORMGEN.Frame1.Visible = False
WS_DATOS = "SI"
fila = 0
Do Until WS_DATOS = "NO"
   fila = fila + 1
'   WS_ULT_FILA = fila
   grid_fac.Row = fila
   grid_fac.Col = 1
   grid_fac.text = far_llave!far_codart
   grid_fac.Col = 2
   PUB_KEY = far_llave!far_codart
   SQ_OPER = 1
   LEER_ART_LLAVE
   If art_llave.EOF Then
      grid_fac.text = "Desconocido"
   Else
      grid_fac.text = art_llave!art_nombre
   End If
   PUB_COSPRO = art_llave!ART_COSPRO
   BUSCAR_ARM
   
   grid_fac.Col = 3
   grid_fac.text = Format(far_llave!FAR_STOCK_REF, "###,###,##0.00")
   grid_fac.ColAlignment(3) = 1

   grid_fac.Col = 4
   grid_fac.text = Format(far_llave!FAR_CANTIDAD, "###,###,##0.00")
   grid_fac.ColAlignment(4) = 1
   grid_fac.Col = 5
   grid_fac.ColAlignment(5) = 1
   grid_fac.text = Format(far_llave!far_precio, "###,###,##0.00")
   
   subtotal = far_llave!FAR_CANTIDAD * far_llave!far_precio
'   PUB_SUBTOTAL = PUB_SUBTOTAL2 + subtotal
   grid_fac.Col = 6
   grid_fac.ColAlignment(6) = 1
   grid_fac.text = Format$(subtotal, "###,###,##0.00")
   grid_fac.Col = 7
   grid_fac.ColAlignment(7) = 1
   grid_fac.text = "I"
   grid_fac.Col = 8
   grid_fac.text = far_llave!FAR_CANTIDAD
   grid_fac.Col = 3
   grid_fac.text = far_llave!FAR_CANTIDAD_REF

   grid_fac.Col = 9
   grid_fac.text = far_llave!far_precio
   grid_fac.HighLight = True
   far_llave.MoveNext
   'REVISAR PORQUE GRABA 3 EN VEZ DE 3.3 EN PRECIO
   If far_llave.EOF = False Then
   If far_llave!FAR_NUMFAC = PU_NUMFAC And far_llave!far_tipmov = 1 And far_llave!FAR_OTRA_CIA = LK_CODCIA Then
      Print " "
   Else
      WS_DATOS = "NO"
   End If
   End If
   If far_llave.EOF = True Then
      WS_DATOS = "NO"
   End If
   
   
Loop
   FORMGEN.i_cant.Locked = True
   FORMGEN.i_codart2.Locked = True
   FORMGEN.grabar.SetFocus
fin:








End Sub

Private Sub i_codban_GotFocus()
Azul i_codban, i_codban
End Sub

Private Sub i_codban_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

If FORMGEN.i_codban.text <> "" Then
   PUB_CODBAN = Val(FORMGEN.i_codban.text)
   SQ_OPER = 1
   LEER_CCM_LLAVE
   If ccm_llave.EOF Then
      MsgBox "banco NO EXISTE ...", 48, WS_TITULO
      FORMGEN.i_codban.SetFocus
      GoTo fin
   Else
      FORMGEN.i_nomban.Caption = ccm_llave(2)
      NUMERO = FORMGEN.i_codban.TabIndex
      avanza_campo
      
   End If
Else
   numarchi = 7
   UNICO = ""
   archi = "SELECT * FROM CCMAEST WHERE CCM_NOMBRE > ? ORDER BY CCM_NOMBRE"
   PROCESA_GRID (archi)
  
End If

  




fin:


End Sub
Private Sub i_codcli_GotFocus()
Azul FORMGEN.i_codcli, i_codcli
' aqui salta de codclie........
If PUB_CP = "C" Or PUB_CP = "P" Then
Else
   NUMERO = FORMGEN.i_codcli.TabIndex
   avanza_campo
End If

End Sub
Private Sub i_codcli_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

If ENTERO(FORMGEN.i_codcli.text) = False Then
   Azul FORMGEN.i_codcli, FORMGEN.i_codcli
   GoTo fin
End If

PUB_CODCLIE = Val(FORMGEN.i_codcli.text)
If PUB_CODCLIE <> 0 Then
   SQ_OPER = 1
   PUB_CODCIA = LK_CODCIA
   LEER_CLI_LLAVE
   If cli_llave.EOF Then
    Azul FORMGEN.i_codcli, FORMGEN.i_codcli
    MsgBox "REGISTRO NO EXISTE ...", 48, WS_TITULO
    FORMGEN.i_codcli.SetFocus
    GoTo fin
   Else
      FORMGEN.i_nomcli.Caption = cli_llave(2)
      NUMERO = FORMGEN.i_codcli.TabIndex
      avanza_campo
      vuelca_datos
   End If
End If

If PUB_CODCLIE = 0 Then
   UNICO = ""
   numarchi = 1
   archi = "SELECT * FROM CLIENTES WHERE CLI_NOMBRE >= ? ORDER BY CLI_NOMBRE"
   PROCESA_GRID (archi)
End If
        
fin:

End Sub



Private Sub i_codven_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

If ENTERO(FORMGEN.i_codven.text) = False Then
   Azul FORMGEN.i_codven, FORMGEN.i_codven
   GoTo fin
End If

PUB_CODVEN = Val(FORMGEN.i_codven.text)
If PUB_CODVEN <> 0 Then
   SQ_OPER = 1
   LEER_VEN_LLAVE
   If ven_llave.EOF Then
      Azul FORMGEN.i_codven, FORMGEN.i_codven
      MsgBox "VENDEDOR NO EXISTE ...", 48, WS_TITULO
      FORMGEN.i_codven.SetFocus
      GoTo fin
   Else
      FORMGEN.i_nomven.Caption = ven_llave(2)
      NUMERO = FORMGEN.i_codven.TabIndex
      avanza_campo
   End If
Else
   numarchi = 2
   UNICO = ""
   archi = "SELECT * FROM VEMAEST WHERE VEM_NOMBRE > ? ORDER BY VEM_NOMBRE"
   PROCESA_GRID (archi)
  
End If



fin:



End Sub



Private Sub i_concepto_GotFocus()
Azul i_concepto, i_concepto
End Sub

Private Sub i_concepto_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If


NUMERO = FORMGEN.i_concepto.TabIndex

avanza_campo
fin:

End Sub
Private Sub i_concepto_LostFocus()
Dim longitud As Integer



'longitud = Len(i_concepto.text)
'If Mid(i_concepto.text, longitud - 1, 1) = "" And Mid(i_concepto.text, longitud, 1) = "" Then
'Else
'If Asc(Mid(i_concepto.text, longitud - 1, 1)) = 13 And Asc(Mid(i_concepto.text, longitud, 1)) = 10 Then
'   i_concepto.text = Left(i_concepto.text, longitud - 2)
'End If
'End If

'MsgBox Asc(Mid(i_concepto.text, longitud - 1, 1))
'MsgBox Asc(Mid(i_concepto.text, longitud, 1))

End Sub

Private Sub i_def_GotFocus()
Dim POS1 As Integer
Dim CARAC As String
If FORMGEN.TRANS.text = "" Then
   GoTo fin
End If


fin:
End Sub

Private Sub i_DEF_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If
NUMERO = TABLA_TAG(tra_llave(2))
WS_INDICE_RETORNO = NUMERO
FORMGEN.Controls(NUMERO).SetFocus

fin:
End Sub


Private Sub i_def_LostFocus()
Dim POS1, CARAC

'Aqui se lee la contabilidad
POS1 = InStr(1, FORMGEN.i_def.List(i_def.ListIndex), ".", 1)
POS1 = POS1 - 1
CARAC = Mid(FORMGEN.i_def.List(i_def.ListIndex), 1, POS1)
PUB_SECUENCIA = Val(CARAC)



SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
LEER_DEF_LLAVE
If def_llave.EOF Then
   MsgBox "No existe Definicion en contabilidad "
   FORMGEN.TRANS.SetFocus
   GoTo fin
End If
pasa_def
If FORMGEN.Frame4.Visible = True Then
   If Val(Nulo_Valor0(def_llave!def_relacion_cant)) = 0 Then
      FORMGEN.i_unidad.Visible = False
      FORMGEN.lcant.Visible = False
   Else
      FORMGEN.i_unidad.Visible = True
      FORMGEN.lcant.Visible = True
   End If
End If

fin:
End Sub

Private Sub i_descto_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If
PUB_DESCTO = Val(FORMGEN.i_descto.text)
FORMGEN.i_impto.SetFocus

fin:

End Sub

Private Sub i_dias_Click()
If MOSTRAR.Visible = True Then
   MOSTRAR.Caption = "< ?  > "
End If

End Sub


Private Sub i_dias_GotFocus()
'SendKeys "%{DOWN}"
End Sub

Private Sub i_dias_KeyPress(KeyAscii As Integer)
Dim FECHA_DIA As Date
Dim fecha_vcto As Variant
Dim msg As String
If KeyAscii <> 13 Then
   GoTo fin
End If
'If ENTERO(FORMGEN.i_dias.text) = False Then
'   Azul FORMGEN.i_dias, FORMGEN.i_dias
'   GoTo FIN
'End If
If LK_CODTRA = 2401 Then
If Val(i_dias.text) = 0 Or Val(i_dias.text) = 8 Or Val(i_dias.text) = 15 Or Val(i_dias.text) = 30 Then
  
  Else
     sn_mensaje = " ¿Esta seguro del numero de Dias... ?"
     ws_respuesta = MsgBox(sn_mensaje, WS_ESTILO, WS_TITULO)
  If ws_respuesta = vbNo Then   ' El usuario eligió
     i_dias.SetFocus
  End If
 
End If
End If

NUMERO = FORMGEN.i_dias.TabIndex
avanza_campo
fin:

End Sub
Private Sub i_dias_LostFocus()

If LK_CODTRA = 2401 Then
   FORMGEN.i_tasa_venta.text = Val(FORMGEN.i_dias.text) * gen!gen_tasa_venta / 30
   FORMGEN.i_tasa_venta.text = redondea(FORMGEN.i_tasa_venta.text)
   PUB_TASA_ORIG = FORMGEN.i_tasa_venta.text
End If
FORMGEN.i_fecha_vcto.text = Str(DateAdd("d", Val(FORMGEN.i_dias.text), LK_FECHA_DIA))
End Sub

Private Sub i_fbg_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

NUMERO = FORMGEN.i_fbg.TabIndex

avanza_campo
fin:

End Sub

Private Sub i_fbg_LostFocus()
'If Not ven_llave.EOF Then
'If i_fbg = "F" Then
'   PUB_NUMSER = ven_llave!VEM_SERIE_F
'Else
'If i_fbg = "B" Then
'   PUB_NUMSER = ven_llave!VEM_SERIE_B
'Else
'If i_fbg = "G" Then
'   PUB_NUMSER = ven_llave!VEM_SERIE_GUIA
'End If
'End If
'End If
'End If

End Sub

Private Sub i_fecha_vcto_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If
If IsDate(i_fecha_vcto) = False Then
   MsgBox "FECHA NO VALIDA ...", 48, WS_TITULO
   GoTo fin
End If

NUMERO = FORMGEN.i_fecha_vcto.TabIndex
avanza_campo
fin:

End Sub
Private Sub i_gastos_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

PUB_GASTOS = Val(FORMGEN.i_gastos.text)
FORMGEN.i_descto.SetFocus

fin:


End Sub


Private Sub i_gastos_not_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

If decimal1(i_gastos_not.text) = False Then
   MsgBox "DATO DEBE SER NUMERICO ", 64, WS_TITULO
   GoTo fin
End If
NUMERO = FORMGEN.i_gastos_not.TabIndex
avanza_campo
fin:

End Sub


Private Sub i_grupo_act_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

NUMERO = FORMGEN.i_grupo_act.TabIndex
avanza_campo
fin:

End Sub

Private Sub i_image_llave_Click()

If LK_CODTRA = 2401 Then
Else
   GoTo otro
End If


grid_autorizacion.Visible = True
grid_autorizacion.SetFocus

otro:
End Sub

Private Sub i_importe_amort_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

If decimal1(i_importe_amort.text) = False Then
   MsgBox "DATO DEBE SER NUMERICO ", 64, WS_TITULO
   GoTo fin
End If
NUMERO = FORMGEN.i_importe_amort.TabIndex
avanza_campo
fin:

End Sub


Private Sub i_cant_Change()

End Sub

Private Sub i_fecha_KeyPress(KeyAscii As Integer)
Dim WS_SALDO As Currency
Dim Tit As String
Dim i As Integer
Dim success%

If KeyAscii <> 13 Then
  GoTo fin
End If


Screen.MousePointer = 11
success% = SetWindowPos(FrmProcesa.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
DoEvents
Dim ws_tot_debe, ws_tot_haber As Currency
Dim ws_fecha  As String


PSCOV_MAYOR.rdoParameters(0) = LK_CODCIA
PSCOV_MAYOR.rdoParameters(1) = i_fecha.text
cov_mayor.Requery
fila = 0
cov_mayor.MoveFirst
If cov_mayor.EOF = True Then
   FrmProcesa.Hide
   MsgBox "No hay Registros...."
   Screen.MousePointer = 0
   GoTo fin
End If
Tit = grid_fac.FormatString
grid_fac.Clear
grid_fac.FormatString = Tit
ws_fecha = cov_mayor!cov_fecha_voucher
Do Until cov_mayor.EOF
       If ws_fecha <> cov_mayor!cov_fecha_voucher Then
          fila = fila + 1
          grid_fac.Rows = fila + 1
          grid_fac.Row = fila
          grid_fac.Col = 0
          grid_fac.text = "totales "
          grid_fac.Col = 1
          grid_fac.text = " "
          grid_fac.Col = 2
          grid_fac.text = " "
          grid_fac.Col = 3
          grid_fac.text = " "
          
          grid_fac.Col = 4
          grid_fac.text = ws_tot_debe
          grid_fac.Col = 5
          grid_fac.text = ws_tot_haber
          ws_tot_debe = 0
          ws_tot_haber = 0
       End If
           


       fila = fila + 1
       grid_fac.Rows = fila + 1
       grid_fac.Row = fila
       grid_fac.Col = 0
       grid_fac.text = cov_mayor!cov_fecha_voucher
       grid_fac.Col = 1
       grid_fac.text = cov_mayor!COV_CUENTA
       grid_fac.Col = 2
       grid_fac.text = cov_mayor!COV_NRO_VOUCHER
       grid_fac.Col = 3
       grid_fac.text = cov_mayor!COV_GLOSA
       If cov_mayor!COV_DH = "D" Then
          grid_fac.Col = 4
          grid_fac.text = cov_mayor!COV_IMPORTE
          grid_fac.Col = 5
          grid_fac.text = ""
          ws_tot_debe = ws_tot_debe + cov_mayor!COV_IMPORTE
       Else
          If cov_mayor!COV_DH = "H" Then
             grid_fac.Col = 4
             grid_fac.text = ""
             grid_fac.Col = 5
             grid_fac.text = cov_mayor!COV_IMPORTE
             ws_tot_haber = ws_tot_haber + cov_mayor!COV_IMPORTE
          End If
       End If
       
       grid_fac.Col = 6
       grid_fac.text = cov_mayor!COV_NRO_MOV
       
otro:
        ws_fecha = cov_mayor!cov_fecha_voucher
        cov_mayor.MoveNext
Loop
          fila = fila + 1
          grid_fac.Rows = fila + 1
          grid_fac.Row = fila
          grid_fac.Col = 0
          grid_fac.text = "totales "
          grid_fac.Col = 1
          grid_fac.text = ""
          grid_fac.Col = 2
          grid_fac.text = ""
          grid_fac.Col = 3
          grid_fac.text = ""
          grid_fac.Col = 4
          grid_fac.text = ws_tot_debe
          grid_fac.Col = 5
          grid_fac.text = ws_tot_haber
          ws_tot_debe = 0
          ws_tot_haber = 0



Screen.MousePointer = 0

fin:
End Sub

Private Sub i_importe_KeyPress(KeyAscii As Integer)
Dim valor As Currency
Dim subtotal As Currency
Dim tf As Integer
If KeyAscii <> 13 Then
   GoTo fin
End If

PUB_CODART = com_llave!COM_CUENTA

If com_llave.EOF = True Then
   MsgBox "Primero seleccione Cuenta... ", 48, WS_TITULO
   GoTo fin
Else
   If com_llave!COM_CUENTA <> Val(i_cuenta.text) Then
   MsgBox "Primero seleccione Articulo.. ", 48, WS_TITULO
   GoTo fin
   End If
End If

If i_importe.text = "" Then
   GoTo fin
End If


Static fila As Integer
If numfilas = -1 Then
   fila = 0
End If

salta:
   
FORMGEN.Frame1.Visible = False
  
   
   filax = filax + 1
   grid_fac.Cols = 10
   grid_fac.Rows = 20

   grid_fac.Row = filax
   grid_fac.Col = 1
   grid_fac.text = com_llave!COM_CUENTA
   grid_fac.Col = 2
   grid_fac.text = com_llave!COM_DESCRIPCION
   grid_fac.Col = 3
   grid_fac.text = i_glosa.text
   grid_fac.ColAlignment(3) = 1
   grid_fac.Col = 4
   grid_fac.text = Format(i_importe.text, "##,###,##0.00")
   grid_fac.ColAlignment(4) = 1
   
'   If i_d_h.text = "D" Then
'      ws_bruto_d = ws_bruto_d + Val(i_importe.text)
'      I_BRUTO_D.text = ws_bruto_d
'   Else
'      ws_bruto_h = ws_bruto_h + Val(i_importe.text)
'      I_BRUTO_H.text = ws_bruto_h
'   End If
   
   grid_fac.Col = 5
   grid_fac.text = i_d_h.text
   grid_fac.Col = 6
   grid_fac.text = i_importe.text

   If filax > 6 Then
      FORMGEN.grid_fac.SetFocus
      SendKeys "{HOME}", True
      SendKeys "{DOWN}", True
      SendKeys "{UP 6}", True
   End If
   i_cuenta.SetFocus
   i_importe.text = ""
   i_d_h.text = ""
   
   grid_fac.Row = filax + 1
   grid_fac.Col = 1
   grid_fac.text = "FIN"
   grid_fac.Col = 6
   grid_fac.ColAlignment(6) = 1
   grid_fac.text = Format$(WS_BRUTO, "###,###,##0.00")

'If Val(I_BRUTO_D.text) = Val(I_BRUTO_H.text) Then
'   MsgBox "*******  OK   ********", 48, WS_TITULO
'   grabar.SetFocus
'End If

fin:



End Sub


Private Sub i_impto_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If
PUB_IMPTO = Val(FORMGEN.i_impto.text)
FORMGEN.i_neto.SetFocus

fin:

End Sub


Private Sub i_gastos_fijos_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

If decimal1(i_gastos_fijos.text) = False Then
   MsgBox "DATO DEBE SER NUMERICO ", 64, WS_TITULO
   GoTo fin
End If
NUMERO = FORMGEN.i_gastos_fijos.TabIndex
avanza_campo
fin:

End Sub

Private Sub i_intade_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

If decimal1(i_intade.text) = False Then
   MsgBox "DATO DEBE SER NUMERICO ", 64, WS_TITULO
   GoTo fin
End If
NUMERO = FORMGEN.i_intade.TabIndex
avanza_campo
fin:

End Sub


Private Sub i_intven_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

If decimal1(FORMGEN.i_intven.text) = False Then
   MsgBox "DATO DEBE SER NUMERICO ", 64, WS_TITULO
   GoTo fin
End If

NUMERO = FORMGEN.i_intven.TabIndex
avanza_campo
fin:




End Sub


Private Sub i_limcre_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

NUMERO = FORMGEN.i_limcre.TabIndex
avanza_campo
fin:
End Sub

Private Sub i_nat_jur_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

NUMERO = FORMGEN.i_nat_jur.TabIndex

avanza_campo
fin:

End Sub

Private Sub i_neto_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

Dim neto As Currency

neto = Val(FORMGEN.i_subtotal) + Val(FORMGEN.i_interes) + Val(i_gastos.text) - Val(FORMGEN.i_descto) + Val(FORMGEN.i_impto)
If neto <> Val(FORMGEN.i_neto.text) Then
MsgBox "ESTA SEGURA DE LO DIGITADO..LA SUMA NO CUADRA...PERO SE RESPETA LO QUE DICE EL DOCUMENTO...)", 48, WS_TITULO
sn_mensaje = " ¿Desea Corregir datos .. ?"
ws_respuesta = MsgBox(sn_mensaje, WS_ESTILO, WS_TITULO)
If ws_respuesta = vbYes Then
   GoTo fin
End If
End If
If Val(FORMGEN.i_impto.text) = 0 Then
   GoTo salta
End If

WS_IMPORTE = Val(FORMGEN.i_subtotal) + Val(FORMGEN.i_interes) + Val(i_gastos.text) - Val(FORMGEN.i_descto)

salta:
FORMGEN.i_subtotal.Enabled = False
FORMGEN.i_gastos.Enabled = False
FORMGEN.i_descto.Enabled = False
FORMGEN.i_impto.Enabled = False
FORMGEN.i_neto.Enabled = False

FORMGEN.i_tipdoc.Enabled = False
FORMGEN.i_dias.Enabled = False
FORMGEN.i_fecha_vcto.Enabled = False
FORMGEN.i_pordescto1.Enabled = False
FORMGEN.i_pordescto2.Enabled = False

FORMGEN.i_codart2.SetFocus

fin:

End Sub



Private Sub i_num_ini_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If


NUMERO = FORMGEN.i_num_ini.TabIndex

avanza_campo
fin:

End Sub

Private Sub i_num_oper_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

NUMERO = FORMGEN.i_num_oper.TabIndex
avanza_campo
fin:

End Sub

Private Sub i_numdoc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If


NUMERO = FORMGEN.i_numdoc.TabIndex
avanza_campo
fin:

End Sub


Private Sub I_NUMDOC_R_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

NUMERO = FORMGEN.i_numdoc_r.TabIndex
avanza_campo


fin:

End Sub

Private Sub i_numfac_c_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

NUMERO = FORMGEN.i_numfac_c.TabIndex

avanza_campo
fin:

End Sub

Private Sub i_numfac_GotFocus()
If PUB_TIPMOV <> 0 Then
   PUB_FBG = FORMGEN.i_fbg.text
   llena_numfac
End If
End Sub

Private Sub i_numfac_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If


NUMERO = FORMGEN.i_numfac.TabIndex

avanza_campo
fin:
End Sub



Private Sub i_numguia_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

If decimal1(i_numguia.text) = False Then
   MsgBox "DATO DEBE SER NUMERICO ", 64, WS_TITULO

   GoTo fin
End If

NUMERO = FORMGEN.i_numguia.TabIndex
avanza_campo
fin:

End Sub

Private Sub i_numser_c_GotFocus()
   Azul i_numser, i_numser
End Sub

Private Sub i_numser_c_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If
NUMERO = FORMGEN.i_numser_c.TabIndex
avanza_campo
fin:

End Sub

Private Sub i_numser_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If


NUMERO = FORMGEN.i_numser.TabIndex

avanza_campo
fin:

End Sub


Private Sub i_numser_LostFocus()


FORMGEN.i_numfac.ListIndex = -1

End Sub

Private Sub i_numser_r_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

FORMGEN.i_numser_r.SetFocus

fin:

End Sub

Private Sub i_pordescto1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

NUMERO = FORMGEN.i_pordescto1.TabIndex
avanza_campo
fin:

End Sub


Private Sub i_pordescto2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

NUMERO = FORMGEN.i_pordescto2.TabIndex
avanza_campo
fin:

End Sub


Private Sub i_precio_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If
NUMERO = FORMGEN.i_precio.TabIndex
avanza_campo
fin:

End Sub


Private Sub i_precio2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If
If FORMGEN.Frame4.Visible = True Then
   If FORMGEN.i_unidad.Visible = True Then
      FORMGEN.i_unidad.SetFocus
   Else
      FORMGEN.i_cant.SetFocus
   End If
End If



fin:

End Sub

Private Sub i_subtotal_GotFocus()
Beep
Beep
If FORMGEN.i_dias.Enabled = True Then
   FORMGEN.i_image_llave.Visible = True
End If
End Sub

Private Sub i_subtotal_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If
If LK_CODTRA = 2403 Then
   FORMGEN.i_codart2.SetFocus
End If
If LK_CODTRA = 2406 Then
   FORMGEN.i_codart2.SetFocus
End If

If Val(FORMGEN.i_subtotal.text) = 0 Then
   GoTo fin
End If
If LK_CODTRA = 2401 Then
   FORMGEN.i_descto.SetFocus
Else
   FORMGEN.i_gastos.SetFocus
End If



fin:

End Sub
Private Sub i_subtotal_LostFocus()
PUB_SUBTOTAL_BAK = Val(FORMGEN.i_subtotal.text)
End Sub

Private Sub i_tasa_ade_Change()
If FORMGEN.i_tasa_ade.text = "" Then
   FORMGEN.i_tasa_ade.text = gen!gen_tasa_leg_adel
End If

If MOSTRAR.Visible = True Then
   MOSTRAR.Caption = "< ?  > "
End If
End Sub

Private Sub i_tasa_ade_GotFocus()
Azul FORMGEN.i_tasa_ade, FORMGEN.i_tasa_ade
If MOSTRAR.Visible = True Then
   MOSTRAR.Caption = "< ?  > "
End If

If MOSTRAR.Visible = True Then
   MOSTRAR.Caption = "< ?  > "
End If
End Sub


Private Sub i_tasa_ade_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If
If decimal1(i_tasa_ade.text) = False Then
   MsgBox "DATO DEBE SER NUMERICO ", 64, WS_TITULO
   GoTo fin
End If

NUMERO = FORMGEN.i_tasa_ade.TabIndex
avanza_campo
fin:
End Sub

End Sub

Private Sub i_tasa_ven_GotFocus()
Azul FORMGEN.i_tasa_ven, FORMGEN.i_tasa_ven

If MOSTRAR.Visible = True Then
   MOSTRAR.Caption = "< ?  > "
End If

End Sub


Private Sub i_tasa_ven_KeyPress(KeyAscii As Integer)

If KeyAscii <> 13 Then
   GoTo fin
End If

If decimal1(i_tasa_ven.text) = False Then
   MsgBox "DATO DEBE SER NUMERICO ", 64, WS_TITULO
   GoTo fin
End If

salta:
NUMERO = FORMGEN.i_tasa_ven.TabIndex
avanza_campo
fin:

End Sub



Private Sub i_tasa_venta_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

NUMERO = FORMGEN.i_tasa_venta.TabIndex

avanza_campo
fin:

End Sub

Private Sub i_tipdoc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

NUMERO = FORMGEN.i_tipdoc.TabIndex

avanza_campo
fin:

End Sub

Private Sub i_codusu_GotFocus()
'SendKeys "%{DOWN}"
End Sub


Private Sub i_unidad_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

FORMGEN.i_unidad.SetFocus

fin:

End Sub

Private Sub Image2_Click()
If numarchi <> 1 Then
   MsgBox "Busqueda para Clientes"
   GoTo fin
End If
   

FORMGEN.LEIDO2.Visible = True
FORMGEN.Image2.Visible = False
FORMGEN.LEIDO2.text = ""
FORMGEN.LEIDO2.SetFocus

fin:

End Sub

Private Sub i_codusu_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

NUMERO = FORMGEN.i_CODUSU.TabIndex
avanza_campo
fin:

End Sub


Private Sub ImgAyuda_Click()
Load FrmAyuda
FrmAyuda.Show
End Sub

Private Sub LEIDO_Change()

Dim DATO As String
'VAR1 = "SELECT * FROM ARTI WHERE ART_NOMBRE > ? ORDER BY ART_NOMBRE"

DATO = procesa(archi)

End Sub

Private Sub LEIDO_Click()
FORMGEN.Grid1.Row = 1
FORMGEN.Grid1.Col = 1

End Sub


Private Sub LEIDO_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If

FORMGEN.Grid1.Row = 1
FORMGEN.Grid1.SetFocus

Select Case numarchi
Case 1
 FORMGEN.Grid1.Col = 1
 FORMGEN.i_nomcli.Caption = FORMGEN.Grid1.text
 FORMGEN.Grid1.Col = 2
 FORMGEN.i_codcli.text = FORMGEN.Grid1.text
 PUB_CODCLIE = Val(FORMGEN.i_codcli.text)
 NUMERO = FORMGEN.i_codcli.TabIndex
 TEXTOX(FORMGEN.i_codcli.TabIndex) = FORMGEN.i_codcli.text
 ETIQUETAX(FORMGEN.i_codcli.TabIndex) = FORMGEN.LABELGEN(FORMGEN.i_codcli.TabIndex).Caption
 NOMBREX(FORMGEN.i_codcli.TabIndex) = FORMGEN.i_nomcli.Caption
 SQ_OPER = 1
 PUB_CODCIA = LK_CODCIA
 LEER_CLI_LLAVE
 vuelca_datos
 Case 2
 FORMGEN.Grid1.Col = 1
 FORMGEN.i_nomven.Caption = FORMGEN.Grid1.text
 FORMGEN.Grid1.Col = 2
 FORMGEN.i_codven.text = FORMGEN.Grid1.text
 PUB_CODVEN = FORMGEN.i_codven.text
 NUMERO = FORMGEN.i_codven.TabIndex
 TEXTOX(FORMGEN.i_codven.TabIndex) = FORMGEN.i_codven.text
 ETIQUETAX(FORMGEN.i_codven.TabIndex) = FORMGEN.LABELGEN(FORMGEN.i_codven.TabIndex).Caption
 NOMBREX(FORMGEN.i_codven.TabIndex) = FORMGEN.i_nomven.Caption
 
 Case 3
 FORMGEN.Grid1.Col = 2
 FORMGEN.TRANS.text = FORMGEN.Grid1.text
 NUMERO = 0
 Case 4

 FORMGEN.Grid1.Col = 2
 FORMGEN.i_codart2.text = FORMGEN.Grid1.text
 SQ_OPER = 1
 PUB_KEY = Val(FORMGEN.i_codart2.text)
 LEER_ART_LLAVE
 If art_llave.EOF Then
    MsgBox "ARTICULO NO EXISTE..."
    FORMGEN.LEIDO.SetFocus
    GoTo fin
 End If
  
 i_nomart2.Caption = art_llave!art_nombre
 If FORMGEN.i_codart.Visible = True Then
    FORMGEN.i_codart.text = PUB_KEY
    FORMGEN.i_nomart.Caption = art_llave!art_nombre
    NUMERO = FORMGEN.i_codart.TabIndex
    GoTo SIGUE
 End If
 
 PUB_COSPRO = art_llave!ART_COSPRO
 BUSCAR_ARM
 GoTo fin
 Case 7
 FORMGEN.Grid1.Col = 1
 FORMGEN.i_nomban.Caption = FORMGEN.Grid1.text
 FORMGEN.Grid1.Col = 2
 FORMGEN.i_codban.text = FORMGEN.Grid1.text
 PUB_CODBAN = Val(FORMGEN.i_codban.text)
 NUMERO = FORMGEN.i_codban.TabIndex

 
 
End Select
SIGUE:
avanza_campo
FORMGEN.Frame1.Visible = False


fin:
End Sub

Private Sub LEIDO_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
  FORMGEN.Frame1.Visible = False
End If


otro:
If KeyCode = 38 Or KeyCode = 40 Then
   Print " "
Else
   GoTo FINAL
End If

FORMGEN.Grid1.SetFocus

FINAL:

End Sub


Private Sub i_cuenta_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer

If KeyAscii <> 13 Then
   GoTo fin
End If

            
 If i_cuenta.text = "" Then
   FORMGEN.Frame4.Visible = True
   UNICO = ""
   
   archi = "SELECT * FROM COMAEST WHERE COM_DESCRIPCION >= ? ORDER BY COM_DESCRIPCION"
   PROCESA_GRID (archi)
 Else
    SQ_OPER = 1
    PUB_CUENTA = i_cuenta.text
    PUB_CODCIA = LK_CODCIA
    
    LEER_COM_LLAVE
    If com_llave.EOF Then
       MsgBox "CUENTA NO EXISTE ...", 48, WS_TITULO
       GoTo fin
    Else
       i_glosa.SetFocus
    End If
 End If
 
 

fin:
End Sub


Private Sub LEIDO2_Click()
LEIDO2.text = ""
End Sub

Private Sub LEIDO2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If
If numarchi <> 1 Then
   MsgBox "Busqueda para Clientes"
   GoTo fin
End If


Dim DATO As String

DATO = alta_vista_nombre(FORMGEN.Grid1, FORMGEN.LEIDO2.text, archi)
fin:
End Sub
Private Sub LEIDO2_KeyUp(KeyCode As Integer, Shift As Integer)
'If KeyCode = 38 Or KeyCode = 40 Then
'   FORMGEN.grid1.SetFocus
'End If



End Sub
Private Sub LEIDO2_LostFocus()
FORMGEN.LEIDO2.Visible = False
FORMGEN.Image2.Visible = True
End Sub

Private Sub LisTransa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  FORMGEN.TRANS.text = Trim(Left(LisTransa.text, 7))
  FORMGEN.LisTransa.Visible = False
  FORMGEN.TRANS.SetFocus
ElseIf KeyAscii = 27 Then
  KeyAscii = 0
  FORMGEN.LisTransa.Visible = False
  FORMGEN.TRANS.SetFocus
End If

End Sub

Private Sub LisTransa_LostFocus()
  FORMGEN.LisTransa.Visible = False
  FORMGEN.TRANS.SetFocus
End Sub

Private Sub mostrar_Click()
Dim ws_tot_pago As Currency
Dim WS_INTADE As Currency
Dim WS_DIF As Currency

Dim FF As Integer
Dim fx As Integer

If car_llave.EOF = True Then
   MsgBox "Seleccionar letra...", 48, WS_TITULO
   GoTo fin
End If



If decimal1(i_importe.text) = False Then
   MsgBox "importe errado... ", 48, WS_TITULO
   GoTo fin
End If


FORMGEN.i_gastos_fijos.text = redondea(Val(car_llave!car_int_8dias))
FORMGEN.i_gastos_not.text = redondea(Nulo_Valor0(car_llave!car_gastos_not))

FORMGEN.i_diasV.text = DateDiff("d", car_llave!CAR_fecha_vcto, LK_FECHA_DIA)

If FORMGEN.i_diasV.text < 0 Then
   FORMGEN.i_diasV.text = 0
End If


FORMGEN.i_diasA.text = DateDiff("d", LK_FECHA_DIA, FORMGEN.i_fecha_vcto.text)

If car_llave!CAR_fecha_vcto > LK_FECHA_DIA Then
   FORMGEN.i_diasA.text = 0
End If

FORMGEN.i_intven.text = redondea(car_llave!CAR_IMPORTE * Val(FORMGEN.i_diasV.text) * Val(FORMGEN.i_tasa_ven.text) / 100)

WS_IMPORTE = 100 * (Val(FORMGEN.i_importe.text) - Val(FORMGEN.i_gastos_fijos.text) - Val(FORMGEN.i_gastos_not.text) - Val(FORMGEN.i_intven.text)) - (car_llave!CAR_IMPORTE * Val(FORMGEN.i_tasa_ade.text) * FORMGEN.i_diasA.text)
WS_IMPORTE = WS_IMPORTE / (100 - Val(FORMGEN.i_diasA.text) * Val(FORMGEN.i_tasa_ade.text))
WS_IMPORTE = redondea(WS_IMPORTE)
If FORMGEN.i_importe.text = "" Then
   WS_IMPORTE = car_llave!CAR_IMPORTE
   GoTo otro
End If

Do Until ws_tot_pago = Val(FORMGEN.i_importe.text)
   If fx = 0 Then
      WS_IMPORTE = WS_IMPORTE + 0.01
   Else
      WS_IMPORTE = WS_IMPORTE - 0.01
   End If
   WS_INTADE = ((car_llave!CAR_IMPORTE - WS_IMPORTE) * Val(FORMGEN.i_tasa_ade.text) * Val(FORMGEN.i_diasA.text)) / 100
   If WS_INTADE < 0 Then
      WS_INTADE = 0
   End If
   
   WS_INTADE = redondea(WS_INTADE)
   ws_tot_pago = WS_IMPORTE + Val(FORMGEN.i_gastos_fijos.text) + Val(FORMGEN.i_gastos_not.text) + Val(FORMGEN.i_intven.text) + WS_INTADE
   If ws_tot_pago > Val(FORMGEN.i_importe.text) Then
      fx = 1
   Else
      fx = 0
   End If
   
   If FF = 0 Then
      WS_DIF = ws_tot_pago - Val(FORMGEN.i_importe.text)
      WS_IMPORTE = WS_IMPORTE - WS_DIF
      WS_IMPORTE = redondea(WS_IMPORTE)
      FF = 1
   End If
   
   
Loop

If WS_IMPORTE > car_llave!CAR_IMPORTE Then
   MsgBox "DEMASIADO PARA ESTA LETRA..."
   FORMGEN.i_importe_amort.text = 0
   FORMGEN.i_intade.text = 0
   FORMGEN.i_codcli.SetFocus
   GoTo fin
End If
If WS_IMPORTE < 0 Then
   MsgBox "Aumentar monto a pagar..."
   FORMGEN.i_importe_amort.text = 0
   FORMGEN.i_intade.text = 0
   FORMGEN.i_codcli.SetFocus
   GoTo fin
End If

If WS_INTADE < 0 Then
   MsgBox "Aumentar monto a pagar... "
   FORMGEN.i_importe_amort.text = 0
   FORMGEN.i_intade.text = 0
   FORMGEN.i_codcli.SetFocus
   GoTo fin
End If
If Val(FORMGEN.i_intven.text) < 0 Then
   MsgBox "Aumentar monto a pagar... "
   FORMGEN.i_importe_amort.text = 0
   FORMGEN.i_intade.text = 0
   FORMGEN.i_codcli.SetFocus
   GoTo fin
End If

otro:
FORMGEN.i_importe_amort.text = redondea(WS_IMPORTE)
FORMGEN.i_intade.text = redondea(WS_INTADE)
i_total_pago.text = redondea(WS_IMPORTE + Val(FORMGEN.i_gastos_fijos.text) + Val(FORMGEN.i_gastos_not.text) + Val(FORMGEN.i_intven.text) + Val(FORMGEN.i_intade.text))
MOSTRAR.Caption = "LIQ. OK "

PASA:
i_saldocar.text = car_llave!CAR_IMPORTE - i_importe_amort.text
If i_saldocar.text < 0 Then
   MsgBox "importe de amortizacion errado...", 48, WS_TITULO
   GoTo fin
End If


'NUMERO = FORMGEN.MOSTRAR.TabIndex
'avanza_campo
grabar.SetFocus
fin:

End Sub

Private Sub salir_Click()
Unload FORM_VOUCHER
End Sub

Private Sub Trans_KeyPress(KeyAscii As Integer)
Dim car As String
Dim SN As String
Dim Control As Object
Dim PRIMER As Integer
Dim POS1, pos2, DIF As Integer
Dim CARAC, CARAC2 As String
Dim i, J, WS_SECUENCIA As Integer
INICIO:

If KeyAscii = 27 Then
  Exit Sub
End If
If KeyAscii = 13 Then
  If TRANS.text = "" Then
    TRANS.SetFocus
    Exit Sub
  End If
End If
If KeyAscii <> 13 Then
   GoTo SALIR
End If
If Trim(TRANS.text) = "5555" Then
  Load FrmActualiza
  FrmActualiza.Show 1
  GoTo SALIR
End If


If tra_llave.EOF Then
   GoTo saltarin
End If

'If Int(FORMGEN.TRANS.text) <> 0 Then
'If Int(FORMGEN.TRANS.text) = tra_llave(0) Then
'   GoTo SALIR
'End If
'End If

cancela_todo
' desaparece a todos los campos actuales
nn = 2
m_ind = 0
Do Until Val(tra_llave(nn)) = 0 Or nn = 62
         m_ind = m_ind + 1
         FORMGEN.LABELGEN(m_ind).Visible = False
         NUMERO = TABLA_TAG(tra_llave(nn))
         If TypeOf FORMGEN.Controls(NUMERO) Is TextBox Then
            FORMGEN.Controls(NUMERO).text = ""
         End If
         FORMGEN.Controls(NUMERO).Visible = False
nn = nn + 4
Loop
saltarin:

f = decimal1(FORMGEN.TRANS.text)
If f = False Then
   MsgBox "DATO DEBE SER NUMERICO ", 64, WS_TITULO
   GoTo SALIR
End If
PUB_CODTRA = Val(Left(FORMGEN.TRANS.text, 4))
LK_CODTRA = PUB_CODTRA
If FORMGEN.TRANS.text <> "" Then
   SQ_OPER = 1
   LEER_TRA_LLAVE
   If tra_llave.EOF Then
      MsgBox "TRANSACCION NO EXISTE...", 48, WS_TITULO
      Azul TRANS, TRANS
      GoTo SALIR
   Else
      LK_NOMTRA = tra_llave(1)
      LK_CODTRA = tra_llave(0)
      FORMGEN.nomtra.Caption = LK_NOMTRA
   End If
Else
   UNICO = ""
   numarchi = 3
   archi = "SELECT * FROM TRANSACCION WHERE TRA_DESCRIPCION  >= ? ORDER BY TRA_DESCRIPCION"
   PROCESA_GRID (archi)
   SQ_OPER = 1
   LEER_TRA_LLAVE
End If

SN = "N"
i = 1
Do Until tra_llave(92 + i) = 0 Or SN = "S" Or i = 10
   J = 1
   Do Until lk_GRUPOS(J) = 0 Or SN = "S" Or J = 10
      If lk_GRUPOS(J) = tra_llave(92 + i) Then
         SN = "S"
      End If
   J = J + 1
   Loop
i = i + 1
Loop

'If SN = "S" Then
'   GoTo CACHETE
'End If


   J = 1
   Do Until lk_CODTRAS(J) = "" Or SN = "Y" Or J = 10
      If Left(lk_CODTRAS(J), 4) = tra_llave(0) Then
         SN = "Y"
         Exit Do
      End If
   J = J + 1
   Loop


If SN = "N" Then
   MsgBox "No Tiene los Derechos Asignados"
   GoTo SALIR
End If



gen.MoveFirst
PASO:

FORMGEN.nomtra.Visible = True
'FORMGEN.i_descto.Locked = False
'FORMGEN.i_gastos.Locked = False
'FORMGEN.i_impto.Locked = False
'FORMGEN.i_neto.Locked = False
'FORMGEN.i_subtotal.Locked = False

CARAC = FORMGEN.TRANS.text
'FORMGEN.Refresh
DoEvents
LLENA_CAMPOS2

If LK_CODTRA = 1401 Then
   PUB_NUMSER = Nulo_Valor0(par_llave!PAR_SER_KARDEX)
   FORMGEN.i_numser.text = PUB_NUMSER
End If


FORMGEN.TRANS.text = CARAC
gridl.Visible = False
grid_che.Visible = False

NUMERO = TABLA_TAG(tra_llave(2))
WS_INDICE_RETORNO = NUMERO

FORMGEN.i_def.Visible = True
i_def.Clear

If SN = "S" Then
SQ_OPER = 2
PUB_CODTRA = LK_CODTRA
LEER_DEF_LLAVE
Do Until def_mayor.EOF
   FORMGEN.i_def.AddItem def_mayor!DEF_SECUENCIA & ".-" & def_mayor!DEF_DESCRIPCION
def_mayor.MoveNext
Loop
GoTo TODO
End If


CARAC2 = "N"
POS1 = 5
pos2 = 99
Do Until CARAC2 = "S" Or pos2 = 0
   POS1 = POS1 + 1
   pos2 = InStr(POS1, lk_CODTRAS(J), ".", 1)
   If pos2 = 0 Then
      pos2 = POS1 + 3
      CARAC2 = "S"
   End If
   DIF = pos2 - POS1
   CARAC = Mid(lk_CODTRAS(J), POS1, DIF)
   PUB_SECUENCIA = Val(CARAC)
   SQ_OPER = 1
   PUB_CODTRA = LK_CODTRA
   LEER_DEF_LLAVE
   FORMGEN.i_def.AddItem def_llave!DEF_SECUENCIA & ".-" & def_llave!DEF_DESCRIPCION
POS1 = pos2
Loop
   
TODO:
POS1 = InStr(1, FORMGEN.TRANS.text, ".", 1)
POS1 = POS1 + 1
If POS1 = 1 Then
   PUB_SECUENCIA = 0
Else
   CARAC = Mid(FORMGEN.TRANS.text, POS1)
   PUB_SECUENCIA = Val(CARAC)
End If
If PUB_SECUENCIA = 0 Then
   i = 0
   GoTo listo
End If
SN = "N"
i = 0

Do Until SN = "S" Or i > FORMGEN.i_def.ListCount - 1
   POS1 = InStr(1, FORMGEN.i_def.List(i), ".", 1)
   POS1 = POS1 - 1
   CARAC = Mid(FORMGEN.i_def.List(i), 1, POS1)

   If PUB_SECUENCIA = Val(CARAC) Then
      SN = "S"
      Exit Do
   End If
i = i + 1
Loop


If SN = "N" Then
   MsgBox "No Tiene los Derechos Asignados "
   GoTo SALIR
End If


listo:
SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
LEER_DEF_LLAVE
If def_llave.EOF Then
   MsgBox "No existe Definicion en contabilidad "
   FORMGEN.TRANS.SetFocus
   GoTo SALIR
End If
pasa_def
'*********    5 ULTIMOS MOVIMIENTOS **********
'If Trim(LK_CODTRA) <> "215" Then
'  Diario.Left = 7320
'  Diario.Top = 5280
'  Diario.Visible = True
'End If

FORMGEN.i_def.SetFocus

FORMGEN.i_def.ListIndex = i
If POS1 = 1 Then
SendKeys "%{DOWN}"
Else
FORMGEN.Controls(WS_INDICE_RETORNO).SetFocus
End If

If FORMGEN.i_def.ListCount = 1 Then
   FORMGEN.Controls(WS_INDICE_RETORNO).SetFocus
End If

SALIR:
End Sub

Private Sub TRANS_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then
 FORMGEN.LisTransa.Left = 980
 FORMGEN.LisTransa.Top = 0
 FORMGEN.LisTransa.Width = 3250
 FORMGEN.LisTransa.Height = 1000
 FORMGEN.LisTransa.Visible = True
 FORMGEN.LisTransa.SetFocus
End If

End Sub

Private Sub Text1_Change()

End Sub
