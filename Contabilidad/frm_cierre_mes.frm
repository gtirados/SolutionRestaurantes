VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frm_mes_cierre 
   BackColor       =   &H00FAEFDA&
   Caption         =   "Cambio de Periodo en Contabilidad"
   ClientHeight    =   4125
   ClientLeft      =   1560
   ClientTop       =   1635
   ClientWidth     =   9240
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
   Icon            =   "frm_cierre_mes.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4125
   ScaleWidth      =   9240
   Tag             =   "55"
   Begin VB.CommandButton Cierre 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Cierre del Periodo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3240
      Picture         =   "frm_cierre_mes.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   2880
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Cierre de cuentas de Gastos e Ingresos y Saldos Intermediarios"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      Picture         =   "frm_cierre_mes.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   9
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   2880
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Cierre de Cuentas de Balances"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5280
      Picture         =   "frm_cierre_mes.frx":06D6
      Style           =   1  'Graphical
      TabIndex        =   8
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   2880
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Frame framensa 
      BackColor       =   &H00FAEFDA&
      Height          =   2295
      Left            =   45
      TabIndex        =   2
      Top             =   75
      Width           =   6495
      Begin ComctlLib.ProgressBar Barra 
         Height          =   285
         Left            =   840
         TabIndex        =   3
         Tag             =   "0"
         Top             =   1800
         Visible         =   0   'False
         Width           =   4395
         _ExtentX        =   7752
         _ExtentY        =   503
         _Version        =   327682
         Appearance      =   0
      End
      Begin VB.Label lblmomento 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   1560
         Width           =   4335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Periodo  Actual :"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lfecha2 
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         Height          =   375
         Left            =   2880
         TabIndex        =   6
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lfecha1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.CommandButton SALIR 
      Caption         =   "Ce&rrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   7920
      Picture         =   "frm_cierre_mes.frx":0820
      Style           =   1  'Graphical
      TabIndex        =   0
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label ayu 
      BackColor       =   &H00FAEFDA&
      Caption         =   $"frm_cierre_mes.frx":096A
      Height          =   2175
      Left            =   6720
      TabIndex        =   10
      Top             =   240
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   1
      Tag             =   "9999"
      X1              =   0
      X2              =   9960
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   0
      Tag             =   "9999"
      X1              =   120
      X2              =   9600
      Y1              =   2760
      Y2              =   2760
   End
End
Attribute VB_Name = "frm_mes_cierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WPASA As Boolean
Dim WSELE As String * 1
Dim llave1
Dim ws_bruto_d, ws_bruto_h As Currency
Dim SUM_D As Currency
Dim SUM_H As Currency
Dim PSCOV_CUENTA As rdoQuery
Dim cov_cuenta As rdoResultset
Dim PSCOM_CUENTA As rdoQuery
Dim com_cuenta As rdoResultset
Dim PSCOP_LLAVE As rdoQuery
Dim cop_llave As rdoResultset
Dim PSCOH_MAYOR As rdoQuery
Dim coh_mayor As rdoResultset
Dim PS_CUENTAX As rdoQuery
Dim cov_cuentax As rdoResultset

Option Explicit


Private Sub Command1_Click()
Dim cuentas As rdoResultset
Dim PSCUENTAS As rdoQuery
Dim comovs   As rdoResultset
Dim PSCOMOVS As rdoQuery
Dim suma_d As Currency
Dim suma_h As Currency
Dim ws_nro_voucher  As Integer
Dim ws_nro_sec As Integer
Dim wcta  As String
Dim wdh  As String * 1
Dim wcta_clientes As Currency
Dim WS_ESTADO As String * 1
Dim wscta  As String
Dim SUMA_SALDOS_D As Currency
Dim SUMA_SALDOS_H As Currency
Dim CTA_CIERRE As String
Dim ULTIMO_OPER As Integer

If LK_EMP_PTO = "A" Then
PSCOV_VOUCHER(0) = "00"
Else
PSCOV_VOUCHER(0) = LK_CODCIA
End If
PUB_CAL_INI = #1/1/2000#
PUB_CAL_FIN = #1/1/2000#
PSCOV_VOUCHER(1) = PUB_CAL_INI
PSCOV_VOUCHER(2) = PUB_CAL_FIN
cov_voucher.Requery
If cov_voucher.EOF Then
 ws_nro_voucher = 0
Else
 cov_voucher.MoveLast
 ws_nro_voucher = cov_voucher!COV_NRO_VOUCHER
End If
ULTIMO_OPER = Nulo_Valor0(cop_llave!COP_ULT_OPER)

pub_cadena = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND COV_CODCTA = ? AND (COV_FECHA_VOUCHER >= '" & Format(PUB_CAL_FIN, "yyyy/mm/dd") & "' AND COV_FECHA_VOUCHER <= '" & Format(PUB_CAL_FIN, "yyyy/mm/dd") & "')  AND COV_NRO_MES = " & LK_NRO_MES & "  ORDER BY COV_CODCTA"
Set PSCOMOVS = CN.CreateQuery("", pub_cadena)
PSCOMOVS(0) = 0
PSCOMOVS(1) = 0
Set comovs = PSCOMOVS.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND (COM_TIPO_CTA  >= 1 AND COM_TIPO_CTA <= 5)  AND COM_NIVEL = 3 ORDER BY COM_CUENTA"
Set PSCUENTAS = CN.CreateQuery("", pub_cadena)
PSCUENTAS(0) = 0
Set cuentas = PSCUENTAS.OpenResultset(rdOpenKeyset, rdConcurValues)

CN.Execute "DELETE COMOV WHERE COV_FLAG_AUTOMATICA = 'Z' AND COV_FECHA_VOUCHER = '" & Format(PUB_CAL_FIN, "yyyy/mm/dd") & "'  ", rdExecDirect
pub_mensaje = "Cuentas anteriores han sido removidas de ," & Trim(Command1.Caption) & " desea Generar el Cierre ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If

' Cierre la cta 69
PSCUENTAS(0) = LK_CODCIA

cuentas.Requery
Barra.Visible = True
DoEvents
Barra.Value = 0
Barra.Min = 0
Barra.Max = cuentas.RowCount

ws_nro_voucher = ws_nro_voucher + 1
ws_nro_sec = 0
Do Until cuentas.EOF
 Barra.Value = Barra.Value + 1
 wscta = cuentas!com_cuenta
' If Left(wscta, 2) = "20" Then Stop
 GoSub COMOV
 If (Val(cuentas!COM_DEB_ANO) + suma_d) = 0 And (Val(cuentas!COM_HAB_ANO) + suma_h) = 0 Then GoTo SALTA1
 If ((Val(cuentas!COM_DEB_ANO) + suma_d) * cuentas!com_signo_d) + ((Val(cuentas!COM_HAB_ANO) + suma_h) * cuentas!com_signo_h) = 0 Then GoTo SALTA1
 If (Val(cuentas!COM_DEB_ANO) + suma_d) > (Val(cuentas!COM_HAB_ANO) + suma_h) Then
' If (Val(cuentas!COM_DEB_ANO) + suma_d) > (Val(cuentas!COM_HAB_ANO) + suma_h) Then
   'wcta_clientes = ((Val(cuentas!COM_DEB_ANO) + suma_d) * cuentas!COM_SIGNO_D) + ((Val(cuentas!COM_HAB_ANO) + suma_h) * cuentas!COM_SIGNO_H)
   wcta_clientes = (Val(cuentas!COM_DEB_ANO) + suma_d) - (Val(cuentas!COM_HAB_ANO) + suma_h)
   WS_ESTADO = "Z"
   wcta = cuentas!com_cuenta
   wdh = "H"
   GoSub GRABA
 Else
   wcta_clientes = (Val(cuentas!COM_HAB_ANO) + suma_h) - (Val(cuentas!COM_DEB_ANO) + suma_d)
   If Left(Trim(cuentas!com_cuenta), 2) = "39" Then
     wcta_clientes = Abs(wcta_clientes)
   End If
   WS_ESTADO = "Z"
   wcta = cuentas!com_cuenta
   wdh = "D"
   GoSub GRABA
 End If
SALTA1:
 cuentas.MoveNext
Loop
Barra.Visible = False

MsgBox "Proceso Terminado Procesar la Mayorización", 48, Pub_Titulo
Exit Sub
COMOV:
PSCOMOVS(0) = LK_CODCIA
PSCOMOVS(1) = wscta
comovs.Requery
suma_d = 0
suma_h = 0
Do Until comovs.EOF
If comovs!COV_DH = "D" Then
  suma_d = suma_d + comovs!COV_IMPORTE
Else
  suma_h = suma_h + comovs!COV_IMPORTE
End If

comovs.MoveNext
Loop
Return

GRABA:
     ws_nro_sec = ws_nro_sec + 1
     cov_voucher.AddNew
     If LK_EMP_PTO = "A" Then
       cov_voucher!COV_CODCIA = "00"
     Else
       cov_voucher!COV_CODCIA = LK_CODCIA
     End If
     cov_voucher!COV_FECHA_VOUCHER = PUB_CAL_FIN
     cov_voucher!COV_NRO_MOV = ws_nro_sec
     cov_voucher!COV_NRO_VOUCHER = ws_nro_voucher
     cov_voucher!COV_CODCTA = wcta
     cov_voucher!COV_DH = wdh
     cov_voucher!COV_IMPORTE = wcta_clientes
     cov_voucher!COV_ESTADO = " "
     cov_voucher!COV_CODUSU = LK_CODCIA
     cov_voucher!cov_flag_automatica = WS_ESTADO
     cov_voucher!COV_glosa = "Cierre de Cuentas"
     cov_voucher!COV_FECHA_doc = LK_FECHA_DIA
     cov_voucher.Update

Return



End Sub

Private Sub Command2_Click()
Dim cuentas As rdoResultset
Dim PSCUENTAS As rdoQuery
Dim comovs   As rdoResultset
Dim PSCOMOVS As rdoQuery
Dim suma_d As Currency
Dim suma_h As Currency
Dim ws_nro_voucher  As Integer
Dim ws_nro_sec As Integer
Dim wcta  As String
Dim wdh  As String * 1
Dim wcta_clientes As Currency
Dim WS_ESTADO As String * 1
Dim wscta  As String
Dim SUMA_SALDOS_D As Currency
Dim SUMA_SALDOS_H As Currency
Dim CTA_CIERRE As String

If LK_EMP_PTO = "A" Then
PSCOV_VOUCHER(0) = "00"
Else
PSCOV_VOUCHER(0) = LK_CODCIA
End If
PUB_CAL_INI = #1/1/2000#
PUB_CAL_FIN = #1/1/2000#
PSCOV_VOUCHER(1) = PUB_CAL_INI
PSCOV_VOUCHER(2) = PUB_CAL_FIN
cov_voucher.Requery
If cov_voucher.EOF Then
 ws_nro_voucher = 0
Else
 cov_voucher.MoveLast
 ws_nro_voucher = cov_voucher!COV_NRO_VOUCHER
End If

pub_cadena = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND COV_CODCTA = ? AND cov_flag_automatica = 'W' AND COV_NRO_MES = " & LK_NRO_MES & " ORDER BY COV_CODCTA"
Set PSCOMOVS = CN.CreateQuery("", pub_cadena)
PSCOMOVS(0) = 0
PSCOMOVS(1) = 0
Set comovs = PSCOMOVS.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND (COM_CUENTA >= ? AND COM_CUENTA < ? OR COM_CUENTA >= ? AND COM_CUENTA < ?) AND COM_NIVEL = 3 ORDER BY COM_CUENTA"
Set PSCUENTAS = CN.CreateQuery("", pub_cadena)
PSCUENTAS(0) = 0
PSCUENTAS(1) = 0
PSCUENTAS(2) = 0
PSCUENTAS(3) = 0
PSCUENTAS(4) = 0
Set cuentas = PSCUENTAS.OpenResultset(rdOpenKeyset, rdConcurValues)

CN.Execute "DELETE COMOV WHERE (COV_FLAG_AUTOMATICA = 'Z' OR COV_FLAG_AUTOMATICA = 'W') AND COV_FECHA_VOUCHER = '" & Format(PUB_CAL_FIN, "yyyy/mm/dd") & "'  ", rdExecDirect
pub_mensaje = "Cuentas anteriores han sido removidas , desea Generar el Cierre ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If
Barra.Visible = True
Barra.Value = 0
Barra.Min = 0
Barra.Max = 14

' Cierre la cta 69
PSCUENTAS(0) = LK_CODCIA
PSCUENTAS(1) = "69"
PSCUENTAS(2) = "70"
PSCUENTAS(3) = "69"
PSCUENTAS(4) = "70"

cuentas.Requery
ws_nro_voucher = ws_nro_voucher + 1
cop_llave.Edit
cop_llave!COP_ULT_OPER = ws_nro_voucher
cop_llave.Update
ws_nro_sec = 0
Do Until cuentas.EOF
 wscta = cuentas!com_cuenta
 GoSub COMOV
 If (Val(cuentas!COM_DEB_ANO) + suma_d) = 0 And (Val(cuentas!COM_HAB_ANO) + suma_h) = 0 Then GoTo SALTA1
 If (Val(cuentas!COM_DEB_ANO) + suma_d) > (Val(cuentas!COM_HAB_ANO) + suma_h) Then
   wcta_clientes = (Val(cuentas!COM_DEB_ANO) + suma_d) - (Val(cuentas!COM_HAB_ANO) + suma_h)
   WS_ESTADO = "W"
   wcta = cuentas!com_cuenta
   wdh = "H"
   GoSub GRABA
   wcta = cuentas!com_cuenta_cierre
   wdh = "D"
   GoSub GRABA
 Else
   wcta_clientes = (Val(cuentas!COM_DEB_ANO) + suma_d) - (Val(cuentas!COM_HAB_ANO) + suma_h)
   WS_ESTADO = "W"
   wcta = cuentas!com_cuenta
   wdh = "H"
   GoSub GRABA
   wcta = cuentas!com_cuenta_cierre
   wdh = "D"
   GoSub GRABA
 End If
SALTA1:
 cuentas.MoveNext
Loop

Barra.Value = Barra.Value + 1
 
' cierre de la clase 9
ws_nro_voucher = ws_nro_voucher + 1
ws_nro_sec = 0
PSCUENTAS(0) = LK_CODCIA
PSCUENTAS(1) = "90"
PSCUENTAS(2) = "99"
PSCUENTAS(3) = "90"
PSCUENTAS(4) = "99"

suma_d = 0
suma_h = 0
cuentas.Requery
SUMA_SALDOS_D = 0
SUMA_SALDOS_H = 0
CTA_CIERRE = cuentas!com_cuenta_cierre
Do Until cuentas.EOF
 'wscta = cuentas!com_cuenta
 'GoSub COMOV
 If (Val(cuentas!COM_DEB_ANO) + suma_d) = 0 And (Val(cuentas!COM_HAB_ANO) + suma_h) = 0 Then GoTo SALTA2
 If (Val(cuentas!COM_DEB_ANO) + suma_d) > (Val(cuentas!COM_HAB_ANO) + suma_h) Then
   wcta_clientes = (Val(cuentas!COM_DEB_ANO) + suma_d) - (Val(cuentas!COM_HAB_ANO) + suma_h)
   WS_ESTADO = "W"
   wcta = cuentas!com_cuenta
   wdh = "H"
   GoSub GRABA
   SUMA_SALDOS_H = SUMA_SALDOS_H + wcta_clientes
 Else
   wcta_clientes = (Val(cuentas!COM_DEB_ANO) + suma_d) - (Val(cuentas!COM_HAB_ANO) + suma_h)
   WS_ESTADO = "W"
   wcta = cuentas!com_cuenta
   wdh = "D"
   GoSub GRABA
   SUMA_SALDOS_D = SUMA_SALDOS_D + wcta_clientes
 End If
SALTA2:
 cuentas.MoveNext
Loop
wcta_clientes = SUMA_SALDOS_H
WS_ESTADO = "W"
wcta = CTA_CIERRE
wdh = "D"
GoSub GRABA

Barra.Value = Barra.Value + 1
' cierre de la 60 y 61 con 80
ws_nro_voucher = ws_nro_voucher + 1
ws_nro_sec = 0
PSCUENTAS(0) = LK_CODCIA
PSCUENTAS(1) = "60"
PSCUENTAS(2) = "62"
PSCUENTAS(3) = "60"
PSCUENTAS(4) = "62"

cuentas.Requery
suma_d = 0
suma_h = 0

SUMA_SALDOS_D = 0
SUMA_SALDOS_H = 0
CTA_CIERRE = cuentas!com_cuenta_cierre
Do Until cuentas.EOF
 wscta = cuentas!com_cuenta
 GoSub COMOV
 If (Val(cuentas!COM_DEB_ANO) + suma_d) = 0 And (Val(cuentas!COM_HAB_ANO) + suma_h) = 0 Then GoTo SALTA3
 If (Val(cuentas!COM_DEB_ANO) + suma_d) > (Val(cuentas!COM_HAB_ANO) + suma_h) Then
   wcta_clientes = (Val(cuentas!COM_DEB_ANO) + suma_d) - (Val(cuentas!COM_HAB_ANO) + suma_h)
   WS_ESTADO = "W"
   wcta = cuentas!com_cuenta
   wdh = "H"
   GoSub GRABA
   SUMA_SALDOS_H = SUMA_SALDOS_H + wcta_clientes
 Else
   wcta_clientes = (Val(cuentas!COM_HAB_ANO) + suma_h) - (Val(cuentas!COM_DEB_ANO) + suma_d)
   WS_ESTADO = "W"
   wcta = cuentas!com_cuenta
   wdh = "D"
   GoSub GRABA
   SUMA_SALDOS_D = SUMA_SALDOS_D + wcta_clientes
 End If
SALTA3:
 cuentas.MoveNext
Loop
wcta_clientes = SUMA_SALDOS_H - SUMA_SALDOS_D
WS_ESTADO = "W"
wcta = CTA_CIERRE
wdh = "D"
GoSub GRABA

Barra.Value = Barra.Value + 1
' cierre de la cta 70 con 80
ws_nro_voucher = ws_nro_voucher + 1
ws_nro_sec = 0
PSCUENTAS(0) = LK_CODCIA
PSCUENTAS(1) = "70"
PSCUENTAS(2) = "71"
PSCUENTAS(3) = "70"
PSCUENTAS(4) = "71"

cuentas.Requery
suma_d = 0
suma_h = 0
SUMA_SALDOS_D = 0
SUMA_SALDOS_H = 0
CTA_CIERRE = cuentas!com_cuenta_cierre
Do Until cuentas.EOF
 wscta = cuentas!com_cuenta
 GoSub COMOV
 If (Val(cuentas!COM_DEB_ANO) + suma_d) = 0 And (Val(cuentas!COM_HAB_ANO) + suma_h) = 0 Then GoTo SALTA4
 If (Val(cuentas!COM_DEB_ANO) + suma_d) > (Val(cuentas!COM_HAB_ANO) + suma_h) Then
   wcta_clientes = (Val(cuentas!COM_DEB_ANO) + suma_d) - (Val(cuentas!COM_HAB_ANO) + suma_h)
   WS_ESTADO = "W"
   wcta = cuentas!com_cuenta
   wdh = "H"
   GoSub GRABA
   SUMA_SALDOS_H = SUMA_SALDOS_H + wcta_clientes
 Else
   wcta_clientes = (Val(cuentas!COM_HAB_ANO) + suma_h) - (Val(cuentas!COM_DEB_ANO) + suma_d)
   WS_ESTADO = "W"
   wcta = cuentas!com_cuenta
   wdh = "D"
   GoSub GRABA
   SUMA_SALDOS_D = SUMA_SALDOS_D + wcta_clientes
 End If
SALTA4:
 cuentas.MoveNext
Loop
wcta_clientes = SUMA_SALDOS_D
WS_ESTADO = "W"
wcta = CTA_CIERRE
wdh = "H"
GoSub GRABA


Barra.Value = Barra.Value + 1
' cierre de la cta 80
ws_nro_voucher = ws_nro_voucher + 1
ws_nro_sec = 0
PSCUENTAS(0) = LK_CODCIA
PSCUENTAS(1) = "80"
PSCUENTAS(2) = "81"
PSCUENTAS(3) = "80"
PSCUENTAS(4) = "81"

cuentas.Requery
suma_d = 0
suma_h = 0
SUMA_SALDOS_D = 0
SUMA_SALDOS_H = 0
CTA_CIERRE = cuentas!com_cuenta_cierre
Do Until cuentas.EOF
 wscta = cuentas!com_cuenta
 GoSub COMOV
 If (Val(cuentas!COM_DEB_ANO) + suma_d) = 0 And (Val(cuentas!COM_HAB_ANO) + suma_h) = 0 Then GoTo SALTA5
 If (Val(cuentas!COM_DEB_ANO) + suma_d) > (Val(cuentas!COM_HAB_ANO) + suma_h) Then
   wcta_clientes = (Val(cuentas!COM_DEB_ANO) + suma_d) - (Val(cuentas!COM_HAB_ANO) + suma_h)
   WS_ESTADO = "W"
   wcta = cuentas!com_cuenta
   wdh = "D" 'corregir
   GoSub GRABA
   SUMA_SALDOS_H = SUMA_SALDOS_H + wcta_clientes
 Else
   wcta_clientes = (Val(cuentas!COM_HAB_ANO) + suma_h) - (Val(cuentas!COM_DEB_ANO) + suma_d)
   WS_ESTADO = "W"
   wcta = cuentas!com_cuenta
   wdh = "D"
   GoSub GRABA
   SUMA_SALDOS_D = SUMA_SALDOS_D + wcta_clientes
 End If
SALTA5:
 cuentas.MoveNext
Loop
wcta_clientes = SUMA_SALDOS_D
WS_ESTADO = "W"
wcta = CTA_CIERRE
wdh = "H"
GoSub GRABA


Barra.Value = Barra.Value + 1
' cierre de la cta 63
ws_nro_voucher = ws_nro_voucher + 1
ws_nro_sec = 0
PSCUENTAS(0) = LK_CODCIA
PSCUENTAS(1) = "63"
PSCUENTAS(2) = "64"
PSCUENTAS(3) = "63"
PSCUENTAS(4) = "64"

cuentas.Requery
suma_d = 0
suma_h = 0
SUMA_SALDOS_D = 0
SUMA_SALDOS_H = 0
CTA_CIERRE = cuentas!com_cuenta_cierre
Do Until cuentas.EOF
 wscta = cuentas!com_cuenta
 GoSub COMOV
 If (Val(cuentas!COM_DEB_ANO) + suma_d) = 0 And (Val(cuentas!COM_HAB_ANO) + suma_h) = 0 Then GoTo SALTA6
 If (Val(cuentas!COM_DEB_ANO) + suma_d) > (Val(cuentas!COM_HAB_ANO) + suma_h) Then
   wcta_clientes = (Val(cuentas!COM_DEB_ANO) + suma_d) - (Val(cuentas!COM_HAB_ANO) + suma_h)
   WS_ESTADO = "W"
   wcta = cuentas!com_cuenta
   wdh = "H"
   GoSub GRABA
   SUMA_SALDOS_H = SUMA_SALDOS_H + wcta_clientes
 Else
   wcta_clientes = (Val(cuentas!COM_HAB_ANO) + suma_h) - (Val(cuentas!COM_DEB_ANO) + suma_d)
   WS_ESTADO = "W"
   wcta = cuentas!com_cuenta
   wdh = "D"
   GoSub GRABA
   SUMA_SALDOS_D = SUMA_SALDOS_D + wcta_clientes
 End If
SALTA6:
 cuentas.MoveNext
Loop
wcta_clientes = SUMA_SALDOS_H
WS_ESTADO = "W"
wcta = CTA_CIERRE
wdh = "D"
GoSub GRABA


Barra.Value = Barra.Value + 1
' cierre de la cta 82
ws_nro_voucher = ws_nro_voucher + 1
ws_nro_sec = 0
PSCUENTAS(0) = LK_CODCIA
PSCUENTAS(1) = "82"
PSCUENTAS(2) = "83"
PSCUENTAS(3) = "82"
PSCUENTAS(4) = "83"

cuentas.Requery
suma_d = 0
suma_h = 0

SUMA_SALDOS_D = 0
SUMA_SALDOS_H = 0
CTA_CIERRE = cuentas!com_cuenta_cierre
Do Until cuentas.EOF
 wscta = cuentas!com_cuenta
 GoSub COMOV
 If (Val(cuentas!COM_DEB_ANO) + suma_d) = 0 And (Val(cuentas!COM_HAB_ANO) + suma_h) = 0 Then GoTo SALTA7
 If (Val(cuentas!COM_DEB_ANO) + suma_d) > (Val(cuentas!COM_HAB_ANO) + suma_h) Then
   wcta_clientes = (Val(cuentas!COM_DEB_ANO) + suma_d) - (Val(cuentas!COM_HAB_ANO) + suma_h)
   WS_ESTADO = "W"
   wcta = cuentas!com_cuenta
   wdh = "H"
   GoSub GRABA
   SUMA_SALDOS_H = SUMA_SALDOS_H + wcta_clientes
 Else
   wcta_clientes = (Val(cuentas!COM_HAB_ANO) + suma_h) - (Val(cuentas!COM_DEB_ANO) + suma_d)
   WS_ESTADO = "W"
   wcta = cuentas!com_cuenta
   wdh = "D"
   GoSub GRABA
   SUMA_SALDOS_D = SUMA_SALDOS_D + wcta_clientes
 End If
SALTA7:
 cuentas.MoveNext
Loop
wcta_clientes = SUMA_SALDOS_D
WS_ESTADO = "W"
wcta = CTA_CIERRE
wdh = "H"
GoSub GRABA



Barra.Value = Barra.Value + 1
' cierre de la cta 62 y 64
ws_nro_voucher = ws_nro_voucher + 1
ws_nro_sec = 0
PSCUENTAS(0) = LK_CODCIA
PSCUENTAS(1) = "62"
PSCUENTAS(2) = "63"
PSCUENTAS(3) = "64"
PSCUENTAS(4) = "65"

cuentas.Requery
suma_d = 0
suma_h = 0

SUMA_SALDOS_D = 0
SUMA_SALDOS_H = 0
CTA_CIERRE = cuentas!com_cuenta_cierre
Do Until cuentas.EOF
 wscta = cuentas!com_cuenta
 GoSub COMOV
 If (Val(cuentas!COM_DEB_ANO) + suma_d) = 0 And (Val(cuentas!COM_HAB_ANO) + suma_h) = 0 Then GoTo SALTA8
 If (Val(cuentas!COM_DEB_ANO) + suma_d) > (Val(cuentas!COM_HAB_ANO) + suma_h) Then
   wcta_clientes = (Val(cuentas!COM_DEB_ANO) + suma_d) - (Val(cuentas!COM_HAB_ANO) + suma_h)
   WS_ESTADO = "W"
   wcta = cuentas!com_cuenta
   wdh = "H"
   GoSub GRABA
   SUMA_SALDOS_H = SUMA_SALDOS_H + wcta_clientes
 Else
   wcta_clientes = (Val(cuentas!COM_HAB_ANO) + suma_h) - (Val(cuentas!COM_DEB_ANO) + suma_d)
   WS_ESTADO = "W"
   wcta = cuentas!com_cuenta
   wdh = "D"
   GoSub GRABA
   SUMA_SALDOS_D = SUMA_SALDOS_D + wcta_clientes
 End If
SALTA8:
 cuentas.MoveNext
Loop
wcta_clientes = SUMA_SALDOS_H
WS_ESTADO = "W"
wcta = CTA_CIERRE
wdh = "D"
GoSub GRABA


Barra.Value = Barra.Value + 1
' cierre de la cta 83
ws_nro_voucher = ws_nro_voucher + 1
ws_nro_sec = 0
PSCUENTAS(0) = LK_CODCIA
PSCUENTAS(1) = "83"
PSCUENTAS(2) = "84"
PSCUENTAS(3) = "83"
PSCUENTAS(4) = "84"

cuentas.Requery
suma_d = 0
suma_h = 0

SUMA_SALDOS_D = 0
SUMA_SALDOS_H = 0
CTA_CIERRE = cuentas!com_cuenta_cierre
Do Until cuentas.EOF
 wscta = cuentas!com_cuenta
 GoSub COMOV
 If (Val(cuentas!COM_DEB_ANO) + suma_d) = 0 And (Val(cuentas!COM_HAB_ANO) + suma_h) = 0 Then GoTo SALTA9
 If (Val(cuentas!COM_DEB_ANO) + suma_d) > (Val(cuentas!COM_HAB_ANO) + suma_h) Then
   wcta_clientes = (Val(cuentas!COM_DEB_ANO) + suma_d) - (Val(cuentas!COM_HAB_ANO) + suma_h)
   WS_ESTADO = "W"
   wcta = cuentas!com_cuenta
   wdh = "H"
   GoSub GRABA
   SUMA_SALDOS_H = SUMA_SALDOS_H + wcta_clientes
 Else
   wcta_clientes = (Val(cuentas!COM_HAB_ANO) + suma_h) - (Val(cuentas!COM_DEB_ANO) + suma_d)
   WS_ESTADO = "W"
   wcta = cuentas!com_cuenta
   wdh = "D"
   GoSub GRABA
   SUMA_SALDOS_D = SUMA_SALDOS_D + wcta_clientes
 End If
SALTA9:
 cuentas.MoveNext
Loop
wcta_clientes = SUMA_SALDOS_H
WS_ESTADO = "W"
wcta = CTA_CIERRE
wdh = "D"
GoSub GRABA



Barra.Value = Barra.Value + 1
' cierre de la cta 65 y 68
ws_nro_voucher = ws_nro_voucher + 1
ws_nro_sec = 0
PSCUENTAS(0) = LK_CODCIA
PSCUENTAS(1) = "65"
PSCUENTAS(2) = "66"
PSCUENTAS(3) = "68"
PSCUENTAS(4) = "69"

cuentas.Requery
suma_d = 0
suma_h = 0

SUMA_SALDOS_D = 0
SUMA_SALDOS_H = 0
CTA_CIERRE = cuentas!com_cuenta_cierre
Do Until cuentas.EOF
 wscta = cuentas!com_cuenta
 GoSub COMOV
 If (Val(cuentas!COM_DEB_ANO) + suma_d) = 0 And (Val(cuentas!COM_HAB_ANO) + suma_h) = 0 Then GoTo SALTA10
 If (Val(cuentas!COM_DEB_ANO) + suma_d) > (Val(cuentas!COM_HAB_ANO) + suma_h) Then
   wcta_clientes = (Val(cuentas!COM_DEB_ANO) + suma_d) - (Val(cuentas!COM_HAB_ANO) + suma_h)
   WS_ESTADO = "W"
   wcta = cuentas!com_cuenta
   wdh = "H"
   GoSub GRABA
   SUMA_SALDOS_H = SUMA_SALDOS_H + wcta_clientes
 Else
   wcta_clientes = (Val(cuentas!COM_HAB_ANO) + suma_h) - (Val(cuentas!COM_DEB_ANO) + suma_d)
   WS_ESTADO = "W"
   wcta = cuentas!com_cuenta
   wdh = "D"
   GoSub GRABA
   SUMA_SALDOS_D = SUMA_SALDOS_D + wcta_clientes
 End If
SALTA10:
 cuentas.MoveNext
Loop
wcta_clientes = SUMA_SALDOS_H
WS_ESTADO = "W"
wcta = CTA_CIERRE
wdh = "D"
GoSub GRABA


Barra.Value = Barra.Value + 1
' cierre de la cta 84
ws_nro_voucher = ws_nro_voucher + 1
ws_nro_sec = 0
PSCUENTAS(0) = LK_CODCIA
PSCUENTAS(1) = "84"
PSCUENTAS(2) = "85"
PSCUENTAS(3) = "84"
PSCUENTAS(4) = "85"

cuentas.Requery
suma_d = 0
suma_h = 0

SUMA_SALDOS_D = 0
SUMA_SALDOS_H = 0
CTA_CIERRE = cuentas!com_cuenta_cierre
Do Until cuentas.EOF
 wscta = cuentas!com_cuenta
 GoSub COMOV
 If (Val(cuentas!COM_DEB_ANO) + suma_d) = 0 And (Val(cuentas!COM_HAB_ANO) + suma_h) = 0 Then GoTo SALTA11
 If (Val(cuentas!COM_DEB_ANO) + suma_d) > (Val(cuentas!COM_HAB_ANO) + suma_h) Then
   wcta_clientes = (Val(cuentas!COM_DEB_ANO) + suma_d) - (Val(cuentas!COM_HAB_ANO) + suma_h)
   WS_ESTADO = "W"
   wcta = cuentas!com_cuenta
   wdh = "H"
   GoSub GRABA
   SUMA_SALDOS_H = SUMA_SALDOS_H + wcta_clientes
 Else
   wcta_clientes = (Val(cuentas!COM_HAB_ANO) + suma_h) - (Val(cuentas!COM_DEB_ANO) + suma_d)
   WS_ESTADO = "W"
   wcta = cuentas!com_cuenta
   wdh = "D"
   GoSub GRABA
   SUMA_SALDOS_D = SUMA_SALDOS_D + wcta_clientes
 End If
SALTA11:
 cuentas.MoveNext
Loop
wcta_clientes = SUMA_SALDOS_H
WS_ESTADO = "W"
wcta = CTA_CIERRE
wdh = "D"
GoSub GRABA



Barra.Value = Barra.Value + 1
' cierre de la cta 66 y 67
ws_nro_voucher = ws_nro_voucher + 1
ws_nro_sec = 0
PSCUENTAS(0) = LK_CODCIA
PSCUENTAS(1) = "66"
PSCUENTAS(2) = "68"
PSCUENTAS(3) = "66"
PSCUENTAS(4) = "68"

cuentas.Requery
suma_d = 0
suma_h = 0

SUMA_SALDOS_D = 0
SUMA_SALDOS_H = 0
CTA_CIERRE = cuentas!com_cuenta_cierre
Do Until cuentas.EOF
 wscta = cuentas!com_cuenta
 GoSub COMOV
 If (Val(cuentas!COM_DEB_ANO) + suma_d) = 0 And (Val(cuentas!COM_HAB_ANO) + suma_h) = 0 Then GoTo SALTA12
 If (Val(cuentas!COM_DEB_ANO) + suma_d) > (Val(cuentas!COM_HAB_ANO) + suma_h) Then
   wcta_clientes = (Val(cuentas!COM_DEB_ANO) + suma_d) - (Val(cuentas!COM_HAB_ANO) + suma_h)
   WS_ESTADO = "W"
   wcta = cuentas!com_cuenta
   wdh = "H"
   GoSub GRABA
   SUMA_SALDOS_H = SUMA_SALDOS_H + wcta_clientes
 Else
   wcta_clientes = (Val(cuentas!COM_HAB_ANO) + suma_h) - (Val(cuentas!COM_DEB_ANO) + suma_d)
   WS_ESTADO = "W"
   wcta = cuentas!com_cuenta
   wdh = "D"
   GoSub GRABA
   SUMA_SALDOS_D = SUMA_SALDOS_D + wcta_clientes
 End If
SALTA12:
 cuentas.MoveNext
Loop
wcta_clientes = SUMA_SALDOS_H
WS_ESTADO = "W"
wcta = CTA_CIERRE
wdh = "D"
GoSub GRABA


' cierre de la cta 77
ws_nro_voucher = ws_nro_voucher + 1
ws_nro_sec = 0
PSCUENTAS(0) = LK_CODCIA
PSCUENTAS(1) = "76"
PSCUENTAS(2) = "78"
PSCUENTAS(3) = "77"
PSCUENTAS(4) = "78"

cuentas.Requery
suma_d = 0
suma_h = 0

SUMA_SALDOS_D = 0
SUMA_SALDOS_H = 0
CTA_CIERRE = cuentas!com_cuenta_cierre
Do Until cuentas.EOF
 wscta = cuentas!com_cuenta
 GoSub COMOV
 If (Val(cuentas!COM_DEB_ANO) + suma_d) = 0 And (Val(cuentas!COM_HAB_ANO) + suma_h) = 0 Then GoTo SALTA13
 If (Val(cuentas!COM_DEB_ANO) + suma_d) > (Val(cuentas!COM_HAB_ANO) + suma_h) Then
   wcta_clientes = (Val(cuentas!COM_DEB_ANO) + suma_d) - (Val(cuentas!COM_HAB_ANO) + suma_h)
   WS_ESTADO = "W"
   wcta = cuentas!com_cuenta
   wdh = "H"
   GoSub GRABA
   SUMA_SALDOS_H = SUMA_SALDOS_H + wcta_clientes
 Else
   wcta_clientes = (Val(cuentas!COM_HAB_ANO) + suma_h) - (Val(cuentas!COM_DEB_ANO) + suma_d)
   WS_ESTADO = "W"
   wcta = cuentas!com_cuenta
   wdh = "D"
   GoSub GRABA
   SUMA_SALDOS_D = SUMA_SALDOS_D + wcta_clientes
 End If
SALTA13:
 cuentas.MoveNext
Loop
wcta_clientes = SUMA_SALDOS_D
WS_ESTADO = "W"
wcta = CTA_CIERRE
wdh = "H"
GoSub GRABA


Barra.Value = Barra.Value + 1
' cierre de la cta 85
ws_nro_voucher = ws_nro_voucher + 1
ws_nro_sec = 0
PSCUENTAS(0) = LK_CODCIA
PSCUENTAS(1) = "85"
PSCUENTAS(2) = "86"
PSCUENTAS(3) = "85"
PSCUENTAS(4) = "86"

cuentas.Requery
suma_d = 0
suma_h = 0

SUMA_SALDOS_D = 0
SUMA_SALDOS_H = 0
CTA_CIERRE = cuentas!com_cuenta_cierre
Do Until cuentas.EOF
 wscta = cuentas!com_cuenta
 GoSub COMOV
 If (Val(cuentas!COM_DEB_ANO) + suma_d) = 0 And (Val(cuentas!COM_HAB_ANO) + suma_h) = 0 Then GoTo SALTA14
 If (Val(cuentas!COM_DEB_ANO) + suma_d) > (Val(cuentas!COM_HAB_ANO) + suma_h) Then
   wcta_clientes = (Val(cuentas!COM_DEB_ANO) + suma_d) - (Val(cuentas!COM_HAB_ANO) + suma_h)
   WS_ESTADO = "W"
   wcta = cuentas!com_cuenta
   wdh = "H"
   GoSub GRABA
   SUMA_SALDOS_H = SUMA_SALDOS_H + wcta_clientes
 Else
   wcta_clientes = (Val(cuentas!COM_HAB_ANO) + suma_h) - (Val(cuentas!COM_DEB_ANO) + suma_d)
   WS_ESTADO = "W"
   wcta = cuentas!com_cuenta
   wdh = "D"
   GoSub GRABA
   SUMA_SALDOS_D = SUMA_SALDOS_D + wcta_clientes
 End If
SALTA14:
 cuentas.MoveNext
Loop
wcta_clientes = SUMA_SALDOS_H
WS_ESTADO = "W"
wcta = CTA_CIERRE
wdh = "D"
GoSub GRABA
Barra.Value = Barra.Value + 1
Barra.Visible = False
MsgBox "Proceso Terminado, Realizar los Asientos en el diario de la Distribución de Utilidades ", 48, Pub_Titulo
Exit Sub
COMOV:
PSCOMOVS(0) = LK_CODCIA
PSCOMOVS(1) = wscta
comovs.Requery
suma_d = 0
suma_h = 0
Do Until comovs.EOF
If comovs!COV_DH = "D" Then
  suma_d = suma_d + comovs!COV_IMPORTE
Else
  suma_h = suma_h + comovs!COV_IMPORTE
End If
comovs.MoveNext
Loop
Return

GRABA:
     ws_nro_sec = ws_nro_sec + 1
     cov_voucher.AddNew
     If LK_EMP_PTO = "A" Then
       cov_voucher!COV_CODCIA = "00"
     Else
       cov_voucher!COV_CODCIA = LK_CODCIA
     End If
     cov_voucher!COV_FECHA_VOUCHER = PUB_CAL_FIN
     cov_voucher!COV_NRO_MOV = ws_nro_sec
     cov_voucher!COV_NRO_VOUCHER = ws_nro_voucher
     cov_voucher!COV_CODCTA = wcta
     cov_voucher!COV_DH = wdh
     cov_voucher!COV_IMPORTE = wcta_clientes
     cov_voucher!COV_ESTADO = " "
     cov_voucher!COV_CODUSU = LK_CODCIA
     cov_voucher!cov_flag_automatica = "W"
     cov_voucher!COV_glosa = "Cierre de Cuentas"
     cov_voucher!COV_FECHA_doc = LK_FECHA_DIA
     cov_voucher.Update

Return



End Sub

Private Sub Form_Load()
CenterMe frm_mes_cierre
WSELE = ""
Dim ws_indice As Integer
Dim cade
cade = "SELECT * FROM COMOV WHERE COV_CODCIA = ?  AND COV_FECHA_VOUCHER >= ? AND COV_FECHA_VOUCHER <= ? AND COV_NRO_MES = " & LK_NRO_MES & "  ORDER BY COV_FECHA_VOUCHER, COV_NRO_MOV"
Set PSCOV_CUENTA = CN.CreateQuery("", cade)
PSCOV_CUENTA(0) = 0
PSCOV_CUENTA(1) = LK_FECHA_DIA
PSCOV_CUENTA(2) = LK_FECHA_DIA
Set cov_cuenta = PSCOV_CUENTA.OpenResultset(rdOpenKeyset, rdConcurValues)

cade = "SELECT * FROM COMAEST WHERE COM_CODCIA = ?  ORDER BY COM_CUENTA"
Set PSCOM_CUENTA = CN.CreateQuery("", cade)
PSCOM_CUENTA(0) = 0
Set com_cuenta = PSCOM_CUENTA.OpenResultset(rdOpenKeyset, rdConcurValues)

cade = "SELECT * FROM COPARAM WHERE COP_CODCIA = ?"
Set PSCOP_LLAVE = CN.CreateQuery("", cade)
PSCOP_LLAVE(0) = 0
Set cop_llave = PSCOP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

cade = "SELECT * FROM COHMAEST WHERE COH_CODCIA = ?"
Set PSCOH_MAYOR = CN.CreateQuery("", cade)
PSCOH_MAYOR(0) = 0
Set coh_mayor = PSCOH_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)

PSCOP_LLAVE.rdoParameters(0) = LK_CODCIA

cop_llave.Requery
If cop_llave.EOF Then
   MsgBox "Definir Parametros en Contabilidad ...", 48, Pub_Titulo
   Exit Sub
End If

lfecha1.Caption = LK_FECHA_COP1
lfecha2.Caption = LK_FECHA_COP2
If CDate(LK_FECHA_COP1) = #1/1/2000# Then
 Command1.Visible = True
 Command2.Visible = True
 ayu.Visible = True
 framensa.Left = 120
 Cierre.Left = 120
End If


End Sub

Private Sub Form_Terminate()
'FORMGEN.TRANS.SetFocus
End Sub
Private Sub Cierre_Click()
Dim WS_ESTADO As String * 1
Dim wsssaldo As Currency
Dim wfecha As String
Dim ws_tot_debe, ws_tot_haber As Currency
Dim ws_fin As Integer
Dim ws_deb_ano As Currency
Dim wcta_clientes As Currency
Dim ws_hab_ano As Currency
Dim ws_nro_voucher As Integer
Dim wcta As String * 12
Dim wdh As String * 1
Dim WS_NRO_MOV As Integer
Dim WS_CUENTA As String * 12
Dim WS_SALDO As Currency
Dim ws_tot_saldo As Currency
Dim CONTADOR As Integer
Dim ws_anomes As Integer
Dim ws_nro_sec As Integer
Dim ws_fecha As Date
Dim WFECHA_COMI As Date
Dim WFECHA_FIM As Date
Dim PS_CUENTAX As rdoQuery
Dim cov_cuentax As rdoResultset
Dim PS_CUENTA2 As rdoQuery
Dim cov_cuenta2 As rdoResultset
Dim PS_CUENTA3 As rdoQuery
Dim cov_cuenta3 As rdoResultset

Dim PSCOMOVS As rdoQuery
Dim comovs As rdoResultset

Dim qw_saldo80 As Currency
Dim w_saldo_ant As Currency
Dim w_d As Currency
Dim w_h As Currency

If cop_llave!COP_FLAG_MAYORIZACION <> "M" Then
   MsgBox "Falta Mayorizar...", 48, Pub_Titulo
   Exit Sub
End If


PSCOP_LLAVE.rdoParameters(0) = LK_CODCIA
cop_llave.Requery
If cop_llave.EOF Then
   MsgBox "Error Grave ..."
   GoTo fin
End If

If par_llave!PAR_CONTABILIDAD = "A" Then
If cop_llave!COP_FECHA_GENCONTAB <> LK_FECHA_COP2 Then
   MsgBox "Revise Datos...Ultima fecha a procesar debe ser " & LK_FECHA_COP2, 48, Pub_Titulo
   MsgBox "La ultima vez proceso fue el dia " & cop_llave!COP_FECHA_GENCONTAB, 48, Pub_Titulo
   GoTo fin
End If
End If


PSCOV_CUENTA.rdoParameters(0) = LK_CODCIA
PSCOV_CUENTA.rdoParameters(1) = LK_FECHA_COP1
PSCOV_CUENTA.rdoParameters(2) = LK_FECHA_COP2

cov_cuenta.Requery
If cov_cuenta.EOF Then
   pub_mensaje = MsgBox("Ojo no hay movimientos Esta seguro de Cerrar ?", 36)
   If pub_mensaje = vbNo Then
     GoTo fin
   Else
     GoTo continua
   End If
End If


continua:
pub_mensaje = "Continuar con el cierre del periodo?..."
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
  GoTo fin
End If
If DatePart("m", LK_FECHA_COP2) = DatePart("m", PUB_CAL_INI) Then
Else
  GoTo saltito1
End If
lblmomento.Caption = "Verificando existencia de cuentas de cierre..."
SQ_OPER = 2
pu_codcia = LK_CODCIA
PUB_CUENTA = " "
PUB_CODCIA = LK_CODCIA
LEER_COM_LLAVE

Barra.Min = 0
Barra.Max = com_mayor.RowCount
CONTADOR = 0
PSCOH_MAYOR.rdoParameters(0) = LK_CODCIA
coh_mayor.Requery
com_mayor.MoveFirst
Do Until com_mayor.EOF
   If Val(cop_llave!cop_nivel_max) = Val(com_mayor!com_nivel) Then
      If com_mayor!com_tipo_cta >= 1 And com_mayor!com_nivel <= 5 Then
      Else
         ws_deb_ano = Val(com_mayor!COM_DEB_ANO) + Val(com_mayor!COM_DEB_MES)
         ws_hab_ano = Val(com_mayor!COM_HAB_ANO) + Val(com_mayor!COM_HAB_MES)
         If ws_deb_ano <> 0 Or ws_hab_ano <> 0 Then
            SQ_OPER = 1
            PUB_CUENTA = com_llave!com_cuenta_cierre
            PUB_CODCIA = LK_CODCIA
            LEER_COM_LLAVE
            If com_llave.EOF Then
               MsgBox "cuenta cierre no existe,.." & PUB_CUENTA & "...." & com_mayor!com_cuenta
               GoTo fin
            End If
         End If
      End If
   End If
    com_mayor.MoveNext
Loop

saltito1:
SQ_OPER = 1
PUB_CAL_INI = DateAdd("d", 1, LK_FECHA_COP2)
If DatePart("m", PUB_CAL_INI) = 1 And DatePart("d", PUB_CAL_INI) = 1 Then
   MsgBox "Ojo Nuevo Periodo será el 01 de enero..." & Chr(13) & "Este periodo se usa para los asientos de ajustes...", 48, Pub_Titulo
   wfecha = PUB_CAL_INI
   GoTo SALTITO
End If
otra_vez:
wfecha = InputBox("La Fecha de Inicio de periodo es : " & PUB_CAL_INI & " , ahora ingrese la fecha proxima de Cierre  :")
If IsDate(wfecha) = False Then
   MsgBox "Fecha Invalida.....Reintente", 48, Pub_Titulo
   GoTo otra_vez
Else
   If CDate(wfecha) <= LK_FECHA_COP2 Then
   MsgBox "Fecha debe ser mayor que " & LK_FECHA_COP1, 48, Pub_Titulo
   GoTo otra_vez
   End If
End If
SALTITO:
PUB_CAL_FIN = wfecha
On Error GoTo error_graba
CN.Execute "Begin Transaction", rdExecDirect
pub_cadena = "SELECT * FROM CONTROLL"
Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)

lblmomento.Caption = "Guardando Información del periodo actual...."
DoEvents
Barra.Visible = True
DoEvents
SQ_OPER = 2
pu_codcia = LK_CODCIA
PUB_CUENTA = " "
PUB_CODCIA = LK_CODCIA
LEER_COM_LLAVE

Barra.Min = 0
'Barra.Max = cov_cuenta.RowCount
Barra.Max = com_mayor.RowCount
CONTADOR = 0
PSCOH_MAYOR.rdoParameters(0) = LK_CODCIA
coh_mayor.Requery
com_mayor.MoveFirst
Do Until com_mayor.EOF
'  If Trim(com_mayor!com_cuenta) > "90" Then Stop
   coh_mayor.AddNew
   coh_mayor!coh_codcia = com_mayor!COM_CODCIA
   coh_mayor!COH_FECHA_PROCESO = LK_FECHA_COP1
   coh_mayor!COH_FECHA_PROCESO2 = LK_FECHA_COP2
   coh_mayor!coH_cuenta = com_mayor!com_cuenta
   coh_mayor!coH_DESCRIPCION = com_mayor!com_DESCRIPCION
   coh_mayor!coh_nivel = com_mayor!com_nivel
   coh_mayor!coh_cuenta_sup = com_mayor!com_cuenta_sup
   coh_mayor!COH_TIPO_CTA = com_mayor!com_tipo_cta
   coh_mayor!COH_DEB_ANO = com_mayor!COM_DEB_ANO
   coh_mayor!COH_HAB_ANO = com_mayor!COM_HAB_ANO
   coh_mayor!COH_DEB_MES = com_mayor!COM_DEB_MES
   coh_mayor!COH_HAB_MES = com_mayor!COM_HAB_MES
   coh_mayor!COH_SIGNO_D = com_mayor!com_signo_d
   coh_mayor!COH_SIGNO_H = com_mayor!com_signo_h
   coh_mayor.Update
   com_mayor.MoveNext
   CONTADOR = CONTADOR + 1
   DoEvents
   Barra.Value = CONTADOR
Loop


lblmomento.Caption = "Inicializando periodo ...."
DoEvents
Cierre.Enabled = False
SALIR.Enabled = False

'PSCOM_CUENTA.rdoParameters(0) = LK_CODCIA
'com_cuenta.Requery

'If com_cuenta.EOF = True Then
'   MsgBox "No hay Plan contable..."
'   GoTo fin
'End If

Barra.Visible = True
Barra.Min = 0
Barra.Max = com_mayor.RowCount
CONTADOR = 0
'Label1.Caption = "Verificando cuadre de Cuentas..."
com_mayor.MoveFirst
Do Until com_mayor.EOF
   com_mayor.Edit
   com_mayor!COM_DEB_ANO = Val(com_mayor!COM_DEB_ANO) + Val(com_mayor!COM_DEB_MES)
   com_mayor!COM_HAB_ANO = Val(com_mayor!COM_HAB_ANO) + Val(com_mayor!COM_HAB_MES)
   com_mayor!COM_DEB_MES = 0
   com_mayor!COM_HAB_MES = 0
   com_mayor.Update
   WS_SALDO = Val(com_mayor!COM_DEB_ANO) + Val(com_mayor!COM_HAB_ANO)
   ws_tot_saldo = ws_tot_saldo + WS_SALDO
   com_mayor.MoveNext
   CONTADOR = CONTADOR + 1
   DoEvents
   Barra.Value = CONTADOR
Loop
If WS_SALDO <> 0 Then
   MsgBox "Comprobación fallo .... Diferencia=" & WS_SALDO
   GoTo fin
End If
lblmomento.Caption = "Generando fechas del nuevo periodo...."
DoEvents
   

lblmomento.Caption = ""
DoEvents
'If DatePart("m", LK_FECHA_COP2) = DatePart("m", PUB_CAL_INI) Then GoTo pasa

pasa:
lblmomento.Caption = "Terminado proceso de cierre...."
DoEvents
CONTADOR = 0
pasito:
WFECHA_COMI = LK_FECHA_COP1
WFECHA_FIM = LK_FECHA_COP2

cop_llave.Edit
'LK_FECHA_COP1 = PUB_CAL_INI
'LK_FECHA_COP2 = PUB_CAL_FIN
cop_llave!COP_FLAG_MAYORIZACION = " "
cop_llave.Update
lblmomento.Caption = ""
Barra.Visible = False
DoEvents

'If PUB_CAL_INI = PUB_CAL_FIN And DatePart("d", PUB_CAL_INI) = 1 And DatePart("m", PUB_CAL_INI) = 1 Then
'Else
'   GoTo termina
'End If
If WFECHA_COMI = WFECHA_COMI And DatePart("d", WFECHA_COMI) = 1 And DatePart("m", WFECHA_COMI) = 1 Then
Else
   GoTo termina
End If

' Inicializar el comaest
SQ_OPER = 2
pu_codcia = LK_CODCIA
PUB_CUENTA = " "
PUB_CODCIA = LK_CODCIA
LEER_COM_LLAVE
com_mayor.Requery
Barra.Visible = True
Barra.Min = 0
Barra.Max = com_mayor.RowCount
Barra.Value = 0
Do Until com_mayor.EOF
   Barra.Value = Barra.Value + 1
   com_mayor.Edit
   com_mayor!COM_DEB_ANO = 0
   com_mayor!COM_HAB_ANO = 0
   com_mayor!COM_DEB_MES = 0
   com_mayor!COM_HAB_MES = 0
   com_mayor.Update
   com_mayor.MoveNext
Loop

PUB_CAL_FIN = #1/1/2000#
PUB_CAL_INI = #1/2/2000#
If LK_EMP_PTO = "A" Then
PSCOV_VOUCHER(0) = "00"
Else
PSCOV_VOUCHER(0) = LK_CODCIA
End If
PSCOV_VOUCHER(1) = PUB_CAL_FIN
PSCOV_VOUCHER(2) = PUB_CAL_FIN
cov_voucher.Requery
If cov_voucher.EOF Then
 ws_nro_voucher = 0
Else
 cov_voucher.MoveLast
 ws_nro_voucher = cov_voucher!COV_NRO_VOUCHER
End If
ws_nro_sec = 0

pub_cadena = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND COV_NRO_VOUCHER >= ? AND (COV_FECHA_VOUCHER >= '" & Format(PUB_CAL_FIN, "yyyy/mm/dd") & "' AND COV_FECHA_VOUCHER <= '" & Format(PUB_CAL_FIN, "yyyy/mm/dd") & "')  AND COV_NRO_MES = " & LK_NRO_MES & " ORDER BY COV_CODCTA"
Set PSCOMOVS = CN.CreateQuery("", pub_cadena)
PSCOMOVS(0) = 0
PSCOMOVS(1) = 0
Set comovs = PSCOMOVS.OpenResultset(rdOpenKeyset, rdConcurValues)

'Cierre la cta 69
PSCOMOVS(0) = LK_CODCIA
PSCOMOVS(1) = ws_nro_voucher
comovs.Requery
Do Until comovs.EOF
  wcta = comovs!COV_CODCTA
 If Trim(comovs!COV_DH) = "D" Then
   wdh = "H"
 Else
   wdh = "D"
 End If
 wcta_clientes = comovs!COV_IMPORTE
 GoSub GRABA
 comovs.MoveNext
Loop
lblmomento.Caption = "Realizando Cierre de Cuentas ...."
DoEvents
Barra.Visible = True
Barra.Min = 0

termina:

CN.Execute "Commit Transaction", rdExecDirect
con_llave.Close
MsgBox "Proceso terminado...Ok", 48, Pub_Titulo

End

GRABA:
     ws_nro_sec = ws_nro_sec + 1
     cov_voucher.AddNew
     If LK_EMP_PTO = "A" Then
       cov_voucher!COV_CODCIA = "00"
     Else
       cov_voucher!COV_CODCIA = LK_CODCIA
     End If
     cov_voucher!COV_FECHA_VOUCHER = PUB_CAL_INI
     cov_voucher!COV_NRO_MOV = ws_nro_sec
     cov_voucher!COV_NRO_VOUCHER = ws_nro_voucher
     cov_voucher!COV_CODCTA = wcta
     cov_voucher!COV_DH = wdh
     cov_voucher!COV_IMPORTE = wcta_clientes
     cov_voucher!COV_ESTADO = " "
     cov_voucher!COV_CODUSU = LK_CODCIA
     cov_voucher!cov_flag_automatica = " "
     cov_voucher!COV_glosa = "Balance Inicial"
     cov_voucher!COV_FECHA_doc = LK_FECHA_DIA
     cov_voucher.Update

Return

fin:

Unload frm_mes_cierre
Unload FORM_CONTA
Exit Sub
error_graba:
MsgBox Err.Description
Resume Next
con_llave.Close
CN.Execute "Rollback Transaction", rdExecDirect


End Sub

Private Sub salir_Click()
Unload frm_mes_cierre
End Sub



