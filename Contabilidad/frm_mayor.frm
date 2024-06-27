VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frm_mayoriz 
   BackColor       =   &H00FAEFDA&
   Caption         =   "Proceso de Mayorización"
   ClientHeight    =   2580
   ClientLeft      =   1560
   ClientTop       =   1575
   ClientWidth     =   7740
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
   Icon            =   "frm_mayor.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2580
   ScaleWidth      =   7740
   Tag             =   "55"
   Begin VB.Frame Frame3 
      BackColor       =   &H00FAEFDA&
      Height          =   2235
      Left            =   6135
      TabIndex        =   10
      Top             =   15
      Width           =   1530
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
         Height          =   690
         Left            =   240
         Picture         =   "frm_mayor.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   3
         Tag             =   "9999"
         Top             =   1350
         Width           =   1095
      End
      Begin VB.CommandButton Mayorizar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Mayorizar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   240
         Picture         =   "frm_mayor.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Tag             =   "9999"
         Top             =   330
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FAEFDA&
      Height          =   690
      Left            =   30
      TabIndex        =   7
      Top             =   1560
      Width           =   6015
      Begin VB.TextBox fecha1 
         Height          =   330
         Left            =   945
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   195
         Width           =   1335
      End
      Begin VB.TextBox fecha2 
         Height          =   285
         Left            =   4170
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   225
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   3255
         TabIndex        =   8
         Top             =   225
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FAEFDA&
      Height          =   1530
      Left            =   30
      TabIndex        =   4
      Top             =   15
      Width           =   6015
      Begin ComctlLib.ProgressBar Barra 
         Height          =   240
         Left            =   150
         TabIndex        =   6
         Tag             =   "0"
         Top             =   1155
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   423
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion de Proceso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   225
         Width           =   5655
      End
   End
   Begin VB.Label lblbarraos 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Solution - Gestion Comercial"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   0
      TabIndex        =   11
      Top             =   2280
      Width           =   7770
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   1
      Tag             =   "9999"
      X1              =   0
      X2              =   9960
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00404040&
      BorderWidth     =   2
      Index           =   0
      Tag             =   "9999"
      X1              =   -120
      X2              =   9360
      Y1              =   2640
      Y2              =   2640
   End
End
Attribute VB_Name = "frm_mayoriz"
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
Dim PSCOX_CUENTA As rdoQuery
Dim PSCOM_CUENTA As rdoQuery
Dim PSallx_llave As rdoQuery

Dim PSTEMP_LLAVE As rdoQuery
Dim temp_llave As rdoResultset

Dim pr_llave As rdoResultset
Dim PSPR_LLAVE As rdoQuery

Dim PSCOM_CUENTA_SUP As rdoQuery
Dim cox_cuenta As rdoResultset
Dim allx_llave As rdoResultset
Dim com_cuenta As rdoResultset
Dim com_cuenta_sup As rdoResultset
Dim co2_cuenta As rdoResultset
Dim PSCO2_CUENTA As rdoQuery
Dim PSCOV_CUENTA As rdoQuery
Dim cov_cuenta As rdoResultset
Dim PSCOV_CUENTA2 As rdoQuery
Dim cov_cuenta2  As rdoResultset

Dim PSDOCSAL As rdoQuery
Dim DOCSAL As rdoResultset
Dim PSMOVICON1 As rdoQuery
Dim MOVICONT1 As rdoResultset
Dim PSMOVICON2 As rdoQuery
Dim MOVICONT2 As rdoResultset
Dim PSDOCSAL2 As rdoQuery
Dim DOCSAL2 As rdoResultset

Option Explicit


Private Sub Form_Load()
WSELE = ""
Dim ws_fecha
Dim ws_indice As Integer
Dim cade
'Frame1.Visible = False


'cade = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND COV_FECHA_VOUCHER>=? AND COV_NRO_MOV=? AND COV_NRO_MES = " & LK_NRO_MES & " "
'Set PSCO2_CUENTA = CN.CreateQuery("", cade)
'PSCO2_CUENTA(0) = 0
'PSCO2_CUENTA(1) = LK_FECHA_DIA
'Set co2_cuenta = PSCO2_CUENTA.OpenResultset(rdOpenKeyset, rdConcurValues)


pub_cadena = "SELECT MOV_TIPO_CAMBIO, MOV_FLAG_TC, MOV_FLAG_DES,MOV_CP, MOV_MONEDA, MOV_CODCLIE, MOV_DH, MOV_NUMFAC, MOV_IMPORTE, MOV_FECHA_EMI, MOV_CODCTA  FROM MOVICONT WHERE MOV_CODCIA = ? AND MOV_TIPMOV = ? AND (MOV_FECHA >= ? AND MOV_FECHA <= ?) AND MOV_NRO_MES = " & LK_NRO_MES & "  ORDER BY MOV_FLAG_DES, MOV_CODCTA , MOV_DH"
Set PSPR_LLAVE = CN.CreateQuery("", pub_cadena)
PSPR_LLAVE(0) = 0
PSPR_LLAVE(1) = 0
PSPR_LLAVE(2) = LK_FECHA_DIA
PSPR_LLAVE(3) = LK_FECHA_DIA
Set pr_llave = PSPR_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT *  FROM MOVICONT WHERE MOV_CODCIA = ? AND (MOV_FECHA >= ? AND MOV_FECHA <=?) AND MOV_NRO_MES = ?  AND MOV_TIPMOV = ?  ORDER BY MOV_NRO_VOUCHER DESC "
Set PSTEMP_LLAVE = CN.CreateQuery("", pub_cadena)
PSTEMP_LLAVE(0) = 0
PSTEMP_LLAVE(1) = LK_FECHA_DIA
PSTEMP_LLAVE(2) = LK_FECHA_DIA
PSTEMP_LLAVE(3) = 0
PSTEMP_LLAVE(4) = 0
Set temp_llave = PSTEMP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

cade = "SELECT * FROM COMAEST WHERE COM_CODCIA = ?  ORDER BY COM_CUENTA"
Set PSCOM_CUENTA = CN.CreateQuery("", cade)
PSCOM_CUENTA(0) = 0
Set com_cuenta = PSCOM_CUENTA.OpenResultset(rdOpenKeyset, rdConcurValues)

cade = "SELECT * FROM COMAEST WHERE COM_CODCIA = ?  AND ( COM_NIVEL = ? OR COM_NIVEL= ? ) ORDER BY COM_CUENTA"
Set PSCOM_CUENTA_SUP = CN.CreateQuery("", cade)
PSCOM_CUENTA_SUP(0) = 0
PSCOM_CUENTA_SUP(1) = 0
PSCOM_CUENTA_SUP(2) = 0
Set com_cuenta_sup = PSCOM_CUENTA_SUP.OpenResultset(rdOpenKeyset, rdConcurValues)

cade = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND COV_FECHA_VOUCHER >= ? AND COV_FECHA_VOUCHER <= ? AND COV_NRO_MES = " & LK_NRO_MES & " ORDER BY COV_CODCTA, COV_FECHA_VOUCHER, COV_NRO_MOV"
Set PSCOV_CUENTA = CN.CreateQuery("", cade)
PSCOV_CUENTA(0) = 0
PSCOV_CUENTA(1) = LK_FECHA_DIA
PSCOV_CUENTA(2) = LK_FECHA_DIA
Set cov_cuenta = PSCOV_CUENTA.OpenResultset(rdOpenKeyset, rdConcurValues)

'cade = "SELECT * FROM COMOV WHERE COV_CODCIA = '" & LK_CODCIA & "' AND COV_NRO_MES = " & LK_NRO_MES & " "
'Set PSCOV_CUENTA2 = CN.CreateQuery("", cade)
'Set cov_cuenta2 = PSCOV_CUENTA2.OpenResultset(rdOpenKeyset, rdConcurValues)

cade = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND COV_FECHA_VOUCHER >= ? AND COV_FECHA_VOUCHER <= ? AND COV_NRO_MES = " & LK_NRO_MES & " ORDER BY COV_NRO_MOV " 'COV_FECHA_VOUCHER, COV_NRO_MOV"
Set PSCOX_CUENTA = CN.CreateQuery("", cade)
PSCOX_CUENTA(0) = 0
PSCOX_CUENTA(1) = LK_FECHA_DIA
PSCOX_CUENTA(2) = LK_FECHA_DIA
Set cox_cuenta = PSCOX_CUENTA.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM ALLOG WHERE   ALL_CODCIA = ? AND ALL_FECHA_DIA >= ? AND ALL_FECHA_DIA <= ? ORDER BY ALL_FECHA_DIA, ALL_NUMOPER "
Set PSallx_llave = CN.CreateQuery("", pub_cadena)
PSallx_llave(0) = 0
PSallx_llave(1) = LK_FECHA_DIA
PSallx_llave(2) = LK_FECHA_DIA
Set allx_llave = PSallx_llave.OpenResultset(rdOpenKeyset, rdConcurValues)


fecha1.Text = LK_FECHA_COP1
fecha2.Text = LK_FECHA_COP2

'**********************************
pub_cadena = "SELECT * FROM DOCSAL WHERE DOC_CODCIA = ? AND DOC_PERIODO = ? AND DOC_SERIE = ? AND DOC_NUMFAC = ? ORDER BY DOC_PERIODO"
Set PSDOCSAL = CN.CreateQuery("", pub_cadena)
PSDOCSAL(0) = ""
PSDOCSAL(1) = 0
PSDOCSAL(2) = 0
PSDOCSAL(3) = 0
Set DOCSAL = PSDOCSAL.OpenResultset(rdOpenKeyset, rdConcurValues)


End Sub

Private Sub Form_Terminate()
Unload frm_mayoriz
End Sub

Private Sub Mayorizar_Click()
' procesa destinos
PRO_DESTINOS
pasa_diario
frm_mayoriz.PROMARY
MayorizarDoc
End Sub

Private Sub salir_Click()
Unload frm_mayoriz
End Sub

Public Sub PROMARY()
Dim WW_FIN, ws_fin As Integer
Dim ws_parcial As Currency

Dim ww_ff As Integer
Dim ww_fff As Integer
Dim WS_SUMA As Currency
Dim ws_tot_debe As Currency
Dim ws_tot_haber As Currency
Dim ws_niv1, ws_niv2 As Integer
Dim WS_CUENTA As String * 12
Dim saldo_i As Currency
Dim saldo_d As Currency
Dim saldo_h As Currency
Dim CONTADOR As Integer
Dim FLAG, fx
Dim NIVEL_ACTUAL As Integer
Dim ffecha1, ffecha2
Dim WS_ENTRO As Integer
Dim WS_NRO_MOV As Currency
Dim ws_nro_mov_ult As Currency
Dim ws_nro_voucher As Currency
Dim WS_DD As Currency
Dim WS_HH As Currency
Dim ws_glosa As String
Dim ws_utilidad As Currency
Dim ws_fecha_proc As Date
Dim WS_IMPORTE As Currency
Dim ws_codusu As String
Dim WS_ESTADO As String * 1
Dim WS_MIRA As Double
Dim WS_SALDO  As Currency
Dim WS_SAL_INICIAL As Currency
Dim ws_dh As String * 1
Dim WS_DH1 As String * 1
Dim wpub_cadena As String
Dim total As Currency
Dim saldo As Currency
Dim CARAC As String * 1
Dim cade As String
Dim WS_NIVEL As Integer
Dim i, J As Integer
Dim TAB_CUENTAS(20) As String * 12
Dim WS_SIGNO_D, WS_SIGNO_H As Integer
Dim WS_CUENTA1
Dim wCOM_NIVEL(6) As Integer
Dim NIVEL_MAX As Integer

cop_llave.Requery

If Not cop_llave.EOF Then
For i = 1 To 6
  If cop_llave.rdoColumns(i) <> 0 Then
     wCOM_NIVEL(i) = cop_llave.rdoColumns(i)
     NIVEL_MAX = i
  End If
Next i
End If

PSallx_llave(0) = LK_CODCIA
PSallx_llave(1) = LK_FECHA_COP1
PSallx_llave(2) = LK_FECHA_COP2

allx_llave.Requery

allx_llave.MoveLast
If Not allx_llave.EOF Then
If par_llave!PAR_CONTABILIDAD = "A" Then
'If cop_llave!COP_ULT_OPER = allx_llave!ALL_NUMOPER And cop_llave!COP_FECHA_GENCONTAB = allx_llave!ALL_FECHA_DIA Then
'Else
'   MsgBox "Se ha adicionado movimientos...Reintente el Pase a Contabilidad"
'End If
End If
End If
PSCOX_CUENTA.rdoParameters(0) = LK_CODCIA
PSCOX_CUENTA.rdoParameters(1) = CDate(fecha1.Text)
PSCOX_CUENTA.rdoParameters(2) = CDate(fecha2.Text)
cox_cuenta.Requery
If cox_cuenta.EOF Then
   MsgBox "No hay movimientos para mayorizar...."
   GoTo cierra
   GoTo fin
End If


'PSCO2_CUENTA.rdoParameters(0) = LK_CODCIA
'PSCO2_CUENTA.rdoParameters(1) = CDate(fecha2.Text)
'PSCO2_CUENTA.rdoParameters(2) = 10000
'co2_cuenta.Requery
'If Not co2_cuenta.EOF Then
'   co2_cuenta.Delete
'End If
'co2_cuenta.Requery

Frame1.Visible = True

Mayorizar.Enabled = False
SALIR.Enabled = False

Barra.Visible = True
'****PENDIENTE DE LOS AÑOS  *****   COS_NRO_ANO = " & WANO & " AND
Label1.Caption = "Inicializando saldos Debe y Haber "
wpub_cadena = "UPDATE COMAEST SET COM_DEB_MES = 0, COM_HAB_MES = 0 WHERE COM_CODCIA = '" & LK_CODCIA & "'"
CN.Execute wpub_cadena, rdExecDirect

If LK_NRO_MES = 0 Then
 wpub_cadena = "UPDATE COMSAL SET COS_DEB00 = 0, COS_HAB00 = 0 WHERE COS_CODCIA = '" & LK_CODCIA & "' AND COS_NRO_ANO = " & Format(LK_FECHA_COP1, "yyyy")
End If
If LK_NRO_MES = 1 Then
 wpub_cadena = "UPDATE COMSAL SET COS_DEB01 = 0, COS_HAB01 = 0 WHERE COS_CODCIA = '" & LK_CODCIA & "' AND COS_NRO_ANO = " & Format(LK_FECHA_COP1, "yyyy")
End If
If LK_NRO_MES = 2 Then
 wpub_cadena = "UPDATE COMSAL SET COS_DEB02 = 0, COS_HAB02 = 0 WHERE COS_CODCIA = '" & LK_CODCIA & "' AND COS_NRO_ANO = " & Format(LK_FECHA_COP1, "yyyy")
End If
If LK_NRO_MES = 3 Then
 wpub_cadena = "UPDATE COMSAL SET COS_DEB03 = 0, COS_HAB03 = 0 WHERE COS_CODCIA = '" & LK_CODCIA & "' AND COS_NRO_ANO = " & Format(LK_FECHA_COP1, "yyyy")
End If
If LK_NRO_MES = 4 Then
 wpub_cadena = "UPDATE COMSAL SET COS_DEB04 = 0, COS_HAB04 = 0 WHERE COS_CODCIA = '" & LK_CODCIA & "' AND COS_NRO_ANO = " & Format(LK_FECHA_COP1, "yyyy")
End If
If LK_NRO_MES = 5 Then
 wpub_cadena = "UPDATE COMSAL SET COS_DEB05 = 0, COS_HAB05 = 0 WHERE COS_CODCIA = '" & LK_CODCIA & "' AND COS_NRO_ANO = " & Format(LK_FECHA_COP1, "yyyy")
End If
If LK_NRO_MES = 6 Then
 wpub_cadena = "UPDATE COMSAL SET COS_DEB06 = 0, COS_HAB06 = 0 WHERE COS_CODCIA = '" & LK_CODCIA & "' AND COS_NRO_ANO = " & Format(LK_FECHA_COP1, "yyyy")
End If
If LK_NRO_MES = 7 Then
 wpub_cadena = "UPDATE COMSAL SET COS_DEB07 = 0, COS_HAB07 = 0 WHERE COS_CODCIA = '" & LK_CODCIA & "' AND COS_NRO_ANO = " & Format(LK_FECHA_COP1, "yyyy")
End If
If LK_NRO_MES = 8 Then
 wpub_cadena = "UPDATE COMSAL SET COS_DEB08 = 0, COS_HAB08 = 0 WHERE COS_CODCIA = '" & LK_CODCIA & "' AND COS_NRO_ANO = " & Format(LK_FECHA_COP1, "yyyy")
End If
If LK_NRO_MES = 9 Then
 wpub_cadena = "UPDATE COMSAL SET COS_DEB09 = 0, COS_HAB09 = 0 WHERE COS_CODCIA = '" & LK_CODCIA & "' AND COS_NRO_ANO = " & Format(LK_FECHA_COP1, "yyyy")
End If
If LK_NRO_MES = 10 Then
 wpub_cadena = "UPDATE COMSAL SET COS_DEB10 = 0, COS_HAB10 = 0 WHERE COS_CODCIA = '" & LK_CODCIA & "' AND COS_NRO_ANO = " & Format(LK_FECHA_COP1, "yyyy")
End If
If LK_NRO_MES = 11 Then
 wpub_cadena = "UPDATE COMSAL SET COS_DEB11 = 0, COS_HAB11 = 0 WHERE COS_CODCIA = '" & LK_CODCIA & "' AND COS_NRO_ANO = " & Format(LK_FECHA_COP1, "yyyy")
End If
If LK_NRO_MES = 12 Then
 wpub_cadena = "UPDATE COMSAL SET COS_DEB12 = 0, COS_HAB12 = 0 WHERE COS_CODCIA = '" & LK_CODCIA & "' AND COS_NRO_ANO = " & Format(LK_FECHA_COP1, "yyyy")
End If
CN.Execute wpub_cadena, rdExecDirect


  
cox_cuenta.Requery
cox_cuenta.MoveLast
WS_NRO_MOV = cox_cuenta!COV_NRO_MOV


'If Month(LK_FECHA_COP1) = Month(Nulo_Valor0(cop_llave!COP_FECHA_AYER)) Then
   GoTo PASA2
'End If
' asientos de cierre de periodo

WS_DD = 0
WS_HH = 0
Pub_Respuesta = MsgBox("Ojo.. se procede al cierre del mes ", Pub_Estilo)
If Pub_Respuesta = vbNo Then
   GoTo fin
End If

    PUB_TIPREG = 77
    PUB_CODCIA = LK_CODCIA
    SQ_OPER = 2
    i = 0
    LEER_TAB_LLAVE
    Do Until tab_mayor.EOF
       If Val(Mid(tab_mayor!tab_nomlargo, 1, 2)) <> 0 And Mid(tab_mayor!tab_nomlargo, 1, 2) <> WS_CUENTA1 Then
          i = i + 1
          TAB_CUENTAS(i) = Mid(tab_mayor!tab_nomlargo, 1, 2)
          WS_CUENTA1 = Mid(tab_mayor!tab_nomlargo, 1, 2)
       End If
       tab_mayor.MoveNext
    Loop
    i = i + 1
'    TAB_CUENTAS(I) = "89"
    


'PSCOM_CUENTA.rdoParameters(0) = LK_CODCIA
'com_cuenta.Requery
    
    
 '   Do Until com_cuenta.EOF
 '   If com_cuenta!com_cuenta = "89          " Then
 '      ws_utilidad = com_cuenta!COM_SAL_INICIAL
 '   End If''

'      FLAG = 0
'      J = 1
'      Do Until I < J Or FLAG = 1
'         If (Trim(Mid(com_cuenta!com_cuenta, 1, 2))) = Trim(TAB_CUENTAS(J)) And cop_llave!COP_NIVEL_AFECTACION = Val(com_cuenta!com_NIVEL) Then
'            FLAG = 1
'         End If
'         J = J + 1
'      Loop
      
'      If FLAG = 1 And com_cuenta!COM_SAL_INICIAL <> 0 Then
'            cox_cuenta.AddNew
'            WS_NRO_MOV = WS_NRO_MOV + 1
'            cox_cuenta!COV_NRO_MOV = WS_NRO_MOV
'            cox_cuenta!COV_CODCTA = com_cuenta!com_cuenta
'            If com_cuenta!COM_SAL_INICIAL < 0 Then
'               If com_cuenta!com_SIGNO_D > 0 Then
'                  cox_cuenta!coV_DH = "D"
'                  cox_cuenta!COV_IMPORTE = com_cuenta!COM_SAL_INICIAL * -1
'                  WS_DD = WS_DD + cox_cuenta!COV_IMPORTE
'               Else
'                  cox_cuenta!coV_DH = "H"
'                  cox_cuenta!COV_IMPORTE = com_cuenta!COM_SAL_INICIAL * -1
'                  WS_HH = WS_HH + cox_cuenta!COV_IMPORTE
'               End If
'            Else
'               If com_cuenta!com_SIGNO_D < 0 Then
'                  cox_cuenta!coV_DH = "D"
'                  cox_cuenta!COV_IMPORTE = com_cuenta!COM_SAL_INICIAL
'                  WS_DD = WS_DD + cox_cuenta!COV_IMPORTE
'               Else
'                  cox_cuenta!coV_DH = "H"
'                  cox_cuenta!COV_IMPORTE = com_cuenta!COM_SAL_INICIAL
'                  WS_HH = WS_HH + cox_cuenta!COV_IMPORTE
'               End If
'           End If
'
'            cox_cuenta!cov_nro_voucher = 999
'            cox_cuenta!COV_GLOSA = "Movimiento de Cierre"
'            cox_cuenta!COV_FECHA_PROC = LK_FECHA_COP1
'            cox_cuenta!COV_CODUSU = LK_CODUSU
'            cox_cuenta!cov_flag_automatica = "C"
'            cox_cuenta!COV_CODCIA = LK_CODCIA
'            cox_cuenta!cov_FECHA_VOUCHER = LK_FECHA_COP1
'            cox_cuenta!COV_ESTADO = " "
'            cox_cuenta.Update
'         End If
'         com_cuenta.MoveNext
'
'       Loop
'            cox_cuenta.AddNew
'            WS_NRO_MOV = WS_NRO_MOV + 1
'            cox_cuenta!COV_NRO_MOV = WS_NRO_MOV
'            If ws_utilidad > 0 Then
'               cox_cuenta!coV_DH = "D"
'               cox_cuenta!COV_IMPORTE = ws_utilidad
'               cox_cuenta!COV_CODCTA = 59101
'            Else
'               cox_cuenta!coV_DH = "H"
'               cox_cuenta!COV_IMPORTE = Abs(ws_utilidad)
'               cox_cuenta!COV_CODCTA = 59101
'            End If
  
'            cox_cuenta!cov_nro_voucher = 999
'            cox_cuenta!COV_GLOSA = "Movimiento de Cierre"
'            cox_cuenta!COV_FECHA_PROC = LK_FECHA_COP1
'            cox_cuenta!COV_CODUSU = LK_CODUSU
'            cox_cuenta!cov_flag_automatica = "C"
'            cox_cuenta!COV_CODCIA = LK_CODCIA
'            cox_cuenta!cov_FECHA_VOUCHER = LK_FECHA_COP1
'            cox_cuenta!COV_ESTADO = " "
 '           cox_cuenta.Update

            
            


PASA2:
PSCOV_CUENTA.rdoParameters(0) = LK_CODCIA
PSCOV_CUENTA.rdoParameters(1) = fecha1.Text
PSCOV_CUENTA.rdoParameters(2) = fecha2.Text
cov_cuenta.Requery
Barra.Min = 0
Barra.Max = cov_cuenta.RowCount
CONTADOR = 0
Label1.Caption = "Verificando Cuadre de Vouchers"
Do Until cov_cuenta.EOF
      If cov_cuenta!COV_DH = "D" Then
         ws_tot_debe = Val(ws_tot_debe) + Nulo_Valor0(cov_cuenta!COV_IMPORTE)
      Else
         If cov_cuenta!COV_DH = "H" Then
            ws_tot_haber = Val(ws_tot_haber) + Nulo_Valor0(cov_cuenta!COV_IMPORTE)
         End If
      End If
   CONTADOR = CONTADOR + 1
   DoEvents
   Barra.Value = CONTADOR
   cov_cuenta.MoveNext
Loop

If ws_tot_debe <> ws_tot_haber Then
   ws_tot_debe = ws_tot_debe - ws_tot_haber
   MsgBox "Movimiento no Cuadrado...Revise por Voucher" & "DIFERENCIA=" & ws_tot_debe
'   GoTo fin
End If

cov_cuenta.MoveFirst

Barra.Min = 0
Barra.Max = cov_cuenta.RowCount
CONTADOR = 0


Label1.Caption = "Sumando Cuentas Contables..."
DoEvents
cov_cuenta.MoveFirst
Do Until cov_cuenta.EOF
      ws_tot_debe = 0
      ws_tot_haber = 0
      WS_FLAG = 0
      WS_CUENTA = cov_cuenta!COV_CODCTA
'      If Trim(WS_CUENTA) = "96109" Then Stop
      Do Until cov_cuenta.EOF Or WS_FLAG = 1
      If cov_cuenta!COV_DH = "D" Then
         ws_tot_debe = ws_tot_debe + Nulo_Valor0(cov_cuenta!COV_IMPORTE)
      Else
      If cov_cuenta!COV_DH = "H" Then
         ws_tot_haber = ws_tot_haber + Nulo_Valor0(cov_cuenta!COV_IMPORTE)
      End If
      End If
      WS_CUENTA = cov_cuenta!COV_CODCTA
      cov_cuenta.MoveNext
      CONTADOR = CONTADOR + 1
      DoEvents
      Barra.Value = CONTADOR
      If Not cov_cuenta.EOF Then
         If cov_cuenta!COV_CODCTA <> WS_CUENTA Then
            WS_FLAG = 1
         End If
      End If
      Loop
      SQ_OPER = 1
      PUB_CUENTA = WS_CUENTA
      PUB_CODCIA = LK_CODCIA
      LEER_COM_LLAVE
      If com_llave.EOF Then
         MsgBox "Error no existe cuenta contable... " & WS_CUENTA
      Else
      com_llave.Edit
      com_llave!COM_DEB_MES = ws_tot_debe '+ com_llave!COM_DEB_MES
      com_llave!COM_HAB_MES = ws_tot_haber '+ com_llave!COM_HAB_MES
      com_llave.Update
      ACT_SALDO_MES PUB_CUENTA, ws_tot_debe, ws_tot_haber
      
      End If
Loop

'ESTE PROCEDIMIENTO ACTUALIZA EL COMAEST Y COMSAL
MAYORIZO:
Label1.Caption = "Mayorizando . . ."
DoEvents
Barra.Value = 0
Barra.Min = 0
Barra.Max = 3
ws_niv1 = NIVEL_MAX
Do Until ws_niv1 = 1
'   barra.Value = barra.Value + 1
   DoEvents
   ws_niv2 = ws_niv1 - 1
   PSCOM_CUENTA_SUP.rdoParameters(0) = LK_CODCIA
   PSCOM_CUENTA_SUP.rdoParameters(1) = ws_niv1
   PSCOM_CUENTA_SUP.rdoParameters(2) = ws_niv2
   com_cuenta_sup.Requery
   saldo_i = 0
   saldo_d = 0
   saldo_h = 0
   WS_CUENTA = com_cuenta_sup!com_cuenta
   Barra.Min = 0
   Barra.Max = com_cuenta_sup.RowCount
   Do Until com_cuenta_sup.EOF
      Barra.Value = com_cuenta_sup.AbsolutePosition
'AGREGADO POR MI=======================================
      saldo_d = 0
      saldo_h = 0
      WS_CUENTA = com_cuenta_sup!com_cuenta
      archi = "SELECT SUM(COM_HAB_MES) AS COM_HAB_MES, SUM(COM_DEB_MES) AS COM_DEB_MES FROM COMAEST WHERE COM_CODCIA = '" & LK_CODCIA & "' AND COM_CUENTA LIKE '" & Trim(WS_CUENTA) & "%' AND COM_NIVEL = " & ws_niv1
      Set PSX = CN.CreateQuery("", archi)
      Set X = PSX.OpenResultset(rdOpenKeyset)
      X.Requery
      If Not X.EOF Then
        If Not (IsNull(X!COM_DEB_MES) And IsNull(X!COM_HAB_MES)) Then
            saldo_d = X!COM_DEB_MES
            saldo_h = X!COM_HAB_MES
            SQ_OPER = 1
            PUB_CUENTA = WS_CUENTA
            PUB_CODCIA = LK_CODCIA
            LEER_COM_LLAVE
            com_llave.Edit
            com_llave!COM_DEB_MES = saldo_d
            com_llave!COM_HAB_MES = saldo_h
            com_llave.Update
            ACT_SALDO_MES PUB_CUENTA, saldo_d, saldo_h
        End If
      End If
      
'=======================================================
'BLOQUEADO POR MIC
'      If Val(com_cuenta_sup!com_nivel) = ws_niv1 Then
'        saldo_d = com_cuenta_sup!COM_DEB_MES + saldo_d
'        saldo_h = com_cuenta_sup!COM_HAB_MES + saldo_h

'      Else
'        SQ_OPER = 1
'        PUB_CUENTA = WS_CUENTA
''        If Left(PUB_CUENTA, 2) = "97" Then Stop
'        PUB_CODCIA = LK_CODCIA
'        LEER_COM_LLAVE
'        com_llave.Edit
'        com_llave!COM_DEB_MES = saldo_d
'        com_llave!COM_HAB_MES = saldo_h
'        com_llave.Update
'        ACT_SALDO_MES PUB_CUENTA, saldo_d, saldo_h
'        WS_CUENTA = com_cuenta_sup!com_cuenta
'        saldo_d = 0
'        saldo_h = 0
'      End If
'=======================================================
      com_cuenta_sup.MoveNext
    Loop
    ws_niv1 = ws_niv1 - 1
Loop

    PUB_TIPREG = 77
    PUB_CODCIA = LK_CODCIA
    SQ_OPER = 2
    LEER_TAB_LLAVE
    Do Until tab_mayor.EOF
       CARAC = Mid(tab_mayor!tab_nomlargo, 3, 1)
       If CARAC = "," Then
          PUB_CUENTA = Mid(tab_mayor!tab_nomlargo, 1, 2)
          PUB_CODCIA = LK_CODCIA
          SQ_OPER = 1
          LEER_COM_LLAVE
          If com_llave.EOF Then
             MsgBox "Corregir Estado de Perdias y Ganancias, Cuenta no Existe = " & Mid(tab_mayor!tab_nomlargo, 1, 2)
             saldo = 0
          Else
            saldo = (Val(com_llave!COM_DEB_MES) + Val(Nulo_Valor0(com_llave!COM_DEB_ANO))) * com_llave!com_signo_d + (Val(Nulo_Valor0(com_llave!COM_HAB_MES)) + Val(Nulo_Valor0(com_llave!COM_HAB_ANO))) * com_llave!com_signo_h
            total = total + (saldo * tab_mayor!TAB_CODART)
          End If
       End If
       tab_mayor.MoveNext
    Loop
    
'      SQ_OPER = 1
'      PUB_CUENTA = "89"
'      LEER_COM_LLAVE
'      If com_llave.EOF Then
'         MsgBox "Crear cuenta de Resultados..89 ..."
'         GoTo fin
'      End If
'      saldo = com_llave!COM_SAL_INICIAL + com_llave!COM_DEB_MES * com_llave!com_SIGNO_D + com_llave!COM_HAB_MES * com_llave!com_SIGNO_H
'      saldo = saldo - total
'      If saldo = 0 Then
'         GoTo PASA
'      End If
       
 '  ws_tot_debe = 0
 '  ws_tot_haber = 0
 '  co2_cuenta.AddNew
 '  co2_cuenta!COV_NRO_MOV = 10000
 '  co2_cuenta!COV_CODCIA = LK_CODCIA
 '  co2_cuenta!cov_nro_voucher = 1000
 '  co2_cuenta!cov_FECHA_VOUCHER = LK_FECHA_COP2
 '  co2_cuenta!COV_GLOSA = " "
 '  co2_cuenta!COV_FECHA_doc = LK_FECHA_COP2
 '  If saldo < 0 Then
 '     PUB_CUENTA = cop_llave!cop_cuenta_azul
 '     co2_cuenta!coV_DH = "D"
 '     co2_cuenta!COV_IMPORTE = saldo * -1
 '     ws_tot_debe = saldo * -1
 '  Else
 '     PUB_CUENTA = cop_llave!cop_cuenta_rojo
 '     co2_cuenta!coV_DH = "H"
 '     co2_cuenta!COV_IMPORTE = saldo
 '     ws_tot_haber = saldo
 '  End If
      
 '  co2_cuenta!COV_CODCTA = PUB_CUENTA
 '  co2_cuenta!COV_ESTADO = "0"
 '  co2_cuenta!COV_CODUSU = LK_CODUSU
 '  co2_cuenta!cov_flag_automatica = "P"
 '  co2_cuenta.Update
    
    
    
  ' ws_fin = 0
  ' Do Until ws_fin = 1
  '    SQ_OPER = 1
  '    LEER_COM_LLAVE
  '    com_llave.Edit
  '    com_llave!COM_DEB_MES = ws_tot_debe + com_llave!COM_DEB_MES
  '    com_llave!COM_HAB_MES = ws_tot_haber + com_llave!COM_HAB_MES
  '    com_llave.Update
  '    If com_llave!com_NIVEL = 1 Then
  '       ws_fin = 1
  '    Else
  '       PUB_CUENTA = com_llave!com_cuenta_sup
  '    End If
  ' Loop
   
pasa:
PSCOM_CUENTA.rdoParameters(0) = LK_CODCIA
com_cuenta.Requery

Dim ws_tot_saldo(6) As Currency
ws_tot_saldo(1) = 0
ws_tot_saldo(2) = 0
ws_tot_saldo(3) = 0
ws_tot_saldo(4) = 0
ws_tot_saldo(5) = 0
ws_tot_saldo(6) = 0
ws_tot_debe = 0


Barra.Min = 0
Barra.Max = com_cuenta.RowCount
Label1.Caption = "Verificando Cuentas Contables"
DoEvents
WS_ENTRO = 0
CONTADOR = 0
WS_NIVEL = 0
   Do Until WW_FIN = 1
      WS_SIGNO_D = com_cuenta!com_signo_d
      WS_SIGNO_H = com_cuenta!com_signo_h
      Do Until WS_NIVEL = 1 Or WW_FIN = 1
         WS_NIVEL = com_cuenta!com_nivel
         ws_fin = 0
         Do Until ws_fin = 1 Or WW_FIN = 1
            If com_cuenta!com_signo_d <> WS_SIGNO_D Or com_cuenta!com_signo_h <> WS_SIGNO_H Then
               MsgBox "Signo se ajustaran...Repocesar" & com_cuenta!com_cuenta
               com_cuenta.Edit
               com_cuenta!com_signo_d = WS_SIGNO_D
               com_cuenta!com_signo_h = WS_SIGNO_H
               com_cuenta.Update
               WS_ENTRO = 1
            End If
            WS_SALDO = (Val(Nulo_Valor0(com_cuenta!COM_HAB_MES)) + Val(Nulo_Valor0(com_cuenta!COM_HAB_ANO))) * com_cuenta!com_signo_h + (Val(Nulo_Valor0(com_cuenta!COM_DEB_MES)) + Val(Nulo_Valor0(com_cuenta!COM_DEB_MES))) * com_cuenta!com_signo_d
            ws_tot_debe = WS_SALDO + ws_tot_debe
            WS_CUENTA = com_cuenta!com_cuenta
            com_cuenta.MoveNext
            If Not com_cuenta.EOF Then
               If com_cuenta!com_nivel <> WS_NIVEL Then
                  ws_fin = 1
               End If
               If WS_NIVEL = 1 And com_cuenta!com_nivel = 1 Then
                  ws_fin = 1
               End If
            Else
               ws_fin = 1
               WW_FIN = 1
            End If
            
            CONTADOR = CONTADOR + 1
            Barra.Value = CONTADOR
         Loop
         If WS_ENTRO = 1 Then
            GoTo fin
         End If

         ws_tot_saldo(WS_NIVEL) = ws_tot_debe + ws_tot_saldo(WS_NIVEL)
         If WW_FIN = 0 Then
            WS_NIVEL = com_cuenta!com_nivel
         End If
         ws_tot_debe = 0
      Loop
    
      
      ws_tot_debe = ws_tot_saldo(1)
      i = 1
'      Do Until i > 6
'      If ws_tot_saldo(i) <> 0 Then
'         If ws_tot_saldo(i) <> ws_tot_debe Then
'            MsgBox "OJO AVISAR A COMPUTO ...(MAÑANA...)" & WS_CUENTA
'         End If
'         ws_tot_saldo(i) = 0
'      End If
'      i = i + 1
'      Loop
      ws_tot_debe = 0
      WS_NIVEL = 0
   Loop



'cov_cuenta.Requery
'Barra.Min = 0
'Barra.Max = cov_cuenta.RowCount
'CONTADOR = 0
'Label1.Caption = "Correlativo de Saldos..."
'WS_CUENTA = cov_cuenta!COV_CODCTA
'Do Until cov_cuenta.EOF
'   If Trim(WS_CUENTA) <> Trim(cov_cuenta!COV_CODCTA) Then
'      SQ_OPER = 1
'      PUB_CUENTA = cov_cuenta!COV_CODCTA
'      LEER_COM_LLAVE
'      cov_cuenta.Edit
'      cov_cuenta!COV_SALDO = com_llave!COM_SAL_INICIAL
'      cov_cuenta.Update
'   End If
'   CONTADOR = CONTADOR + 1
'   DoEvents
'   Barra.Value = CONTADOR
'   WS_CUENTA = cov_cuenta!COV_CODCTA
'   cov_cuenta.MoveNext
'Loop


cierra:
cop_llave.Requery
cop_llave.Edit
cop_llave!COP_FLAG_MAYORIZACION = "M"
'cop_llave!COP_FECHA_AYER = fecha1.text
LK_FECHA_COP2 = fecha2.Text
cop_llave.Update
ACT_MESES (1)


Exit Sub
fin:
Unload frm_mayoriz


End Sub
Public Sub PRO_DESTINOS()
Dim WS_NRO_MOV As Integer
Dim ws_fecha_voucher As Date
Dim ws_nro_voucher As Currency
Dim ww_fff As Integer
Dim ws_por As Currency
Dim ws_dh As String
Dim ws_codusu As String
Dim wc_importe As Currency
Dim ws_fecha_proc As Date
Dim ws_glosa As String
Dim WS_SUMA As Currency
Dim ww_ff As Integer
Dim ws_vou As Integer
Dim CONTADOR As Integer
Dim wpub_cadena As String
Dim ffecha1 As String
Dim ffecha2 As String

Dim cade As String
Dim WS_NRO As Currency
Dim TEMPO_VAR
Dim PSCOV_CUENTA2 As rdoQuery
Dim cox_cuenta  As rdoResultset
Dim WS_CUENTA As String
Dim ws_parcial As Currency
Dim ws_dh_cov As String * 1
Dim tt_importe As Currency
Dim tempo_tipmov As Integer

Dim wcadena As String
Dim wvalor  As String * 1

Dim imp_total As Currency
Dim imp_des1 As Currency
Dim imp_des2 As Currency
Dim imp_des3 As Currency
Dim imp_des4 As Currency
Dim imp_des5 As Currency

' CHEQUE DE MES CERRADO 0 ABIERTO
WS_NRO = 0
wcadena = ""
SQ_OPER = 2
PUB_TIPREG = 155
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
If tab_mayor.EOF Then
  MsgBox "Crear tab_tipreg = 155 para seguridas", 48, Pub_Titulo
Else
    Do Until tab_mayor.EOF
      If tab_mayor!TAB_NUMTAB = Val(Format(LK_FECHA_COP1, "yyyy")) Then
        wcadena = Trim(tab_mayor!tab_nomlargo)
      End If
      tab_mayor.MoveNext
    Loop
    wvalor = Mid(wcadena, LK_NRO_MES + 1, 1)
    If wvalor = "1" Then
       MsgBox "<<<< Mes CERRADO Operaciones >>>>", vbCritical, Pub_Titulo
       Exit Sub
    End If
End If

'cade = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND COV_FECHA_VOUCHER >= ? AND COV_FECHA_VOUCHER <= ?  AND COV_NRO_MES = " & LK_NRO_MES & " ORDER BY COV_CODCTA, COV_FECHA_VOUCHER, COV_NRO_MOV"
'Set PSCOV_CUENTA2 = CN.CreateQuery("", cade)
'PSCOV_CUENTA2(0) = 0
'PSCOV_CUENTA2(1) = LK_FECHA_DIA
'PSCOV_CUENTA2(2) = LK_FECHA_DIA
'Set cox_cuenta = PSCOV_CUENTA2.OpenResultset(rdOpenKeyset, rdConcurValues)

cade = "SELECT * FROM MOVICONT WHERE MOV_CODCIA = ? AND MOV_FECHA >= ? AND MOV_FECHA <= ?  AND MOV_NRO_MES = " & LK_NRO_MES & " ORDER BY MOV_TIPMOV, MOV_DH, MOV_CODCTA, MOV_FECHA, MOV_NRO_MOV"
Set PSCOV_CUENTA2 = CN.CreateQuery("", cade)
PSCOV_CUENTA2(0) = 0
PSCOV_CUENTA2(1) = LK_FECHA_DIA
PSCOV_CUENTA2(2) = LK_FECHA_DIA
Set cox_cuenta = PSCOV_CUENTA2.OpenResultset(rdOpenKeyset, rdConcurValues)

ffecha1 = Format(LK_FECHA_COP1, "dd/mm/yyyy")
ffecha2 = Format(LK_FECHA_COP2, "dd/mm/yyyy")
wpub_cadena = "DELETE MOVICONT  WHERE (MOV_FLAG_DES = 'A') AND MOV_CODCIA = '" & LK_CODCIA & "' AND MOV_FECHA >=  '" & ffecha1 & "' AND MOV_FECHA <=  '" & ffecha2 & "'"
CN.Execute wpub_cadena, rdExecDirect

'pub_mensaje = "Procesoar los Destinos ¿Desea Continuar... ?"
'Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
'If Pub_Respuesta = vbNo Then
'    Exit Sub
'End If
Frame1.Visible = True
DoEvents

Label1.Caption = "Eliminando Movimientos Automaticos"
Label1.Caption = ""


PSCOV_CUENTA2.rdoParameters(0) = LK_CODCIA
PSCOV_CUENTA2.rdoParameters(1) = LK_FECHA_COP1
PSCOV_CUENTA2.rdoParameters(2) = LK_FECHA_COP2
cox_cuenta.Requery


PSTEMP_LLAVE(0) = LK_CODCIA
PSTEMP_LLAVE(1) = LK_FECHA_COP1
PSTEMP_LLAVE(2) = LK_FECHA_COP2
PSTEMP_LLAVE(3) = LK_NRO_MES

'barra.Max = cox_cuenta.RowCount
CONTADOR = 0
Label1.Caption = "Creando Movimientos Automaticos"
If Not cox_cuenta.EOF Then
 Barra.Max = cox_cuenta.RowCount
End If
Barra.Min = 0
Barra.Value = 0
Barra.Visible = True
DoEvents
tempo_tipmov = -1
ws_vou = -99 'cox_cuenta!MOV_nro_voucher
Do Until cox_cuenta.EOF
'     If cox_cuenta!MOV_TIPMOV = 3 Then Stop
      Barra.Value = Barra.Value + 1
      SQ_OPER = 1
      PUB_CUENTA = cox_cuenta!MOV_CODCTA
      PUB_CODCIA = LK_CODCIA
      LEER_COM_LLAVE
      If com_llave.EOF Then
         MsgBox "Revisar el diario ...Cuenta : " & PUB_CUENTA
         GoTo fin
      End If
      ww_ff = 0
      If com_llave!COM_POR_AUTOM_D <> 0 Then ww_ff = 1
      If com_llave!COM_POR_AUTOM_D2 <> 0 Then ww_ff = 2
      If com_llave!COM_POR_AUTOM_D3 <> 0 Then ww_ff = 3
      If com_llave!COM_POR_AUTOM_D4 <> 0 Then ww_ff = 4
      If com_llave!COM_POR_AUTOM_D5 <> 0 Then ww_ff = 5
      If Left(cox_cuenta!MOV_CODCTA, 1) = "9" And TEMPO_VAR = 1 Then
'        MsgBox ".."
      End If
      If Left(cox_cuenta!MOV_CODCTA, 1) = "9" Then
        TEMPO_VAR = 1
      End If
'      WS_NRO_MOV = 0
      
      WS_SUMA = Val(Nulo_Valors(com_llave!com_cuenta_AUTOM_D)) + Val(Nulo_Valors(com_llave!COM_CUENTA_AUTOM_D2)) + Val(Nulo_Valors(com_llave!COM_CUENTA_AUTOM_D3)) + Val(Nulo_Valors(com_llave!COM_CUENTA_AUTOM_D4)) + Val(Nulo_Valors(com_llave!COM_CUENTA_AUTOM_D5))
          
      If WS_SUMA <> 0 Then
         WS_SUMA = 0
         'ws_nro_voucher = cox_cuenta!COV_NRO_VOUCHER
         WS_NRO = cox_cuenta!MOV_NRO_VOUCHER
         ws_glosa = cox_cuenta!MOV_GLOSA
         ws_fecha_proc = cox_cuenta!MOV_FECHA
         wc_importe = cox_cuenta!MOV_IMPORTE
         tt_importe = wc_importe
         ws_codusu = cox_cuenta!MOV_CODUSU
         
         If cox_cuenta!MOV_DH = "H" And (Left(cox_cuenta!MOV_CODCTA, 1) = "6" Or Left(cox_cuenta!MOV_CODCTA, 1) = "9") Then
            ws_dh = "H"
         Else
            ws_dh = "D"
         End If
         imp_des1 = 0
         imp_des2 = 0
         imp_des3 = 0
         imp_des4 = 0
         imp_des5 = 0
         WS_CUENTA = com_llave!com_cuenta_AUTOM_D
         ws_dh_cov = cox_cuenta!MOV_DH
         If Trim(com_llave!com_cuenta_AUTOM_D) <> "" Then
            ws_por = com_llave!COM_POR_AUTOM_D
            imp_des1 = Format(wc_importe * ws_por / 100, "0.00")
         End If

         If Trim(com_llave!COM_CUENTA_AUTOM_D2) <> "" Then
            ws_por = com_llave!COM_POR_AUTOM_D2
            imp_des2 = Format(wc_importe * ws_por / 100, "0.00")
         End If
         If Trim(com_llave!COM_CUENTA_AUTOM_D3) <> "" Then
            ws_por = com_llave!COM_POR_AUTOM_D3
            imp_des3 = Format(wc_importe * ws_por / 100, "0.00")
         End If
         If Trim(com_llave!COM_CUENTA_AUTOM_D4) <> "" Then
            ws_por = com_llave!COM_POR_AUTOM_D4
            imp_des4 = Format(wc_importe * ws_por / 100, "0.00")
         End If
         If Trim(com_llave!COM_CUENTA_AUTOM_D5) <> "" Then
            ws_por = com_llave!COM_POR_AUTOM_D5
            imp_des5 = Format(wc_importe * ws_por / 100, "0.00")
         End If
         
         imp_total = imp_des1 + imp_des2 + imp_des3 + imp_des4 + imp_des5
         If Val(wc_importe) <> Val(imp_total) Then
          imp_des1 = Val(imp_des1) + (Val(wc_importe) - Val(imp_total))
         End If
         
         WS_CUENTA = com_llave!com_cuenta_AUTOM_D
         If Trim(com_llave!com_cuenta_AUTOM_D) <> "" Then
            ww_fff = 1
            ws_parcial = imp_des1
            GoSub graba_autom
         End If
         WS_CUENTA = com_llave!COM_CUENTA_AUTOM_D2
         If Trim(com_llave!COM_CUENTA_AUTOM_D2) <> "" Then
            ww_fff = 2
            ws_parcial = imp_des2
            GoSub graba_autom
         End If
         WS_CUENTA = com_llave!COM_CUENTA_AUTOM_D3
         If Trim(com_llave!COM_CUENTA_AUTOM_D3) <> "" Then
            ww_fff = 3
            ws_parcial = imp_des3
            GoSub graba_autom
         End If
         WS_CUENTA = com_llave!COM_CUENTA_AUTOM_D4
         If Trim(com_llave!COM_CUENTA_AUTOM_D4) <> "" Then
            ww_fff = 4
            ws_parcial = imp_des4
            GoSub graba_autom
         End If
         WS_CUENTA = com_llave!COM_CUENTA_AUTOM_D5
         If Trim(com_llave!COM_CUENTA_AUTOM_D5) <> "" Then
            ww_fff = 5
            ws_parcial = imp_des5
            GoSub graba_autom
         End If
         
         
         If Trim(com_llave!com_cuenta_AUTO_H) <> "" Then
            WS_CUENTA = com_llave!com_cuenta_AUTO_H
            WS_SUMA = 0
            ws_por = 100
            ws_parcial = wc_importe
            If cox_cuenta!MOV_DH = "H" And (Left(cox_cuenta!MOV_CODCTA, 1) = "6" Or Left(cox_cuenta!MOV_CODCTA, 1) = "9") Then
               ws_dh = "D" ' Vws_dh = cox_cuenta!COV_DH
            Else
               ws_dh = "H"
            End If
            GoSub graba_autom
         End If
      End If
   CONTADOR = CONTADOR + 1
   DoEvents
'   barra.Value = CONTADOR
   cox_cuenta.MoveNext
Loop
 Barra.Visible = False
 Label1.Caption = ""
 cop_llave.Requery
 cop_llave.Edit
 cop_llave!cop_FLAG_DES = "A"
 cop_llave.Update
 ACT_MESES (0)
' MsgBox "PROCESO TERMINADO...", 48, Pub_Titulo
 
 
 

Exit Sub

graba_autom:
           If ww_ff <> 0 Then
               If ws_por = 0 Then Return
           End If
           'ws_parcial = Format(wc_importe * ws_por / 100, "0.00")
           If ws_parcial = 0 Then Return
            
           If tempo_tipmov <> cox_cuenta!MOV_TIPMOV Then
             PSTEMP_LLAVE(4) = cox_cuenta!MOV_TIPMOV
             temp_llave.Requery
             If temp_llave.EOF Then
                ws_nro_voucher = 0
                ws_fecha_voucher = LK_FECHA_COP1
             Else
           '     temp_llave.MoveLast
                ws_nro_voucher = temp_llave!MOV_NRO_VOUCHER
                ws_fecha_voucher = temp_llave!MOV_FECHA
            End If
            tempo_tipmov = cox_cuenta!MOV_TIPMOV
            WS_NRO_MOV = 0
            ws_nro_voucher = ws_nro_voucher + 1
            ws_vou = cox_cuenta!MOV_NRO_VOUCHER
          Else
           If cox_cuenta!MOV_NRO_VOUCHER <> ws_vou Then
                ws_nro_voucher = ws_nro_voucher + 1
                ws_vou = cox_cuenta!MOV_NRO_VOUCHER
            End If
          End If
          
      
            temp_llave.AddNew
            temp_llave!MOV_NRO_VOUCHER = Val(WS_NRO) + 0.1 ' ws_nro_voucher
            temp_llave!MOV_FECHA = ws_fecha_voucher 'ws_fecha_proc
            WS_NRO_MOV = WS_NRO_MOV + 1
            temp_llave!MOV_NRO_MOV = WS_NRO_MOV
            temp_llave!MOV_TIPMOV = cox_cuenta!MOV_TIPMOV
            If tempo_tipmov = 1 Then
                temp_llave!MOV_GLOSA = "Destino por las Compras "
            ElseIf tempo_tipmov = 2 Then
                temp_llave!MOV_GLOSA = "Destino por las Ventas "
            ElseIf tempo_tipmov = 3 Then
                If cox_cuenta!MOV_PLANTILLA = 100 Then
                  temp_llave!MOV_GLOSA = "Destino por los Ingresos de Caja"
                Else
                    temp_llave!MOV_GLOSA = "Destino por los Egresos de Caja"
                End If
            Else
                temp_llave!MOV_GLOSA = "Destino por " & Trim(ws_glosa)
            End If
            temp_llave!MOV_MONEDA = "S"
            temp_llave!MOV_SUNAT = "00"
            temp_llave!MOV_serie = 0
            temp_llave!MOV_numfac = 0
            temp_llave!MOV_codclie = 0
            temp_llave!MOV_CP = ""
            temp_llave!MOV_FBG = " "
            temp_llave!MOV_MARCA = ""
            temp_llave!MOV_fecha_EMI = ws_fecha_voucher 'ws_fecha_proc
            temp_llave!MOV_serie_c = 0
            temp_llave!MOV_numfac_c = 0
            temp_llave!MOV_FBG_C = " "
            temp_llave!MOV_PLANTILLA = cox_cuenta!MOV_PLANTILLA
            temp_llave!MOV_FLAG_TC = " "
            temp_llave!MOV_TIPO_CAMBIO = 0
            temp_llave!MOV_CODUSU = LK_CODUSU
            temp_llave!MOV_CODCTA = WS_CUENTA
             'If Trim(WS_CUENTA) = "" Then Stop
            temp_llave!MOV_DH = ws_dh
            
            temp_llave!MOV_DETALLE = Trim(temp_llave!MOV_GLOSA) ' "Destino Cta.: " & Trim(com_llave!com_cuenta) & " " & ws_glosa
            temp_llave!MOV_FLAG_DES = "A"
            WS_SUMA = WS_SUMA + ws_parcial
            temp_llave!MOV_IMPORTE = ws_parcial
            temp_llave!MOV_CODUSU = ws_codusu
            temp_llave!MOV_CODCIA = LK_CODCIA
            temp_llave!MOV_nro_MES = LK_NRO_MES
            temp_llave!MOV_PERIODO = Format(LK_FECHA_COP2, "yyyy")
            temp_llave.Update
   TEMPO_VAR = 0
 Return
Exit Sub
fin:


End Sub


Public Sub pasa_diario()
Dim wscodcia As String
Dim wcta As String
Dim wc_codcta As String
Dim wdh As String * 1
Dim WSGLOSA As String
Dim ws_nro_voucher As Currency
Dim LOC_PROCESO As Integer
Dim wpub_cadena As String
Dim CUENTA_PRO As Integer
Dim kl_voucher  As Integer
Dim WS_NRO_MOV  As Integer
Dim wc_importe As Currency
Dim wc_sum_debe  As Currency
Dim wc_sum_haber  As Currency
Dim WSUM_GEN_D As Currency
Dim WSUM_GEN_H As Currency
Dim Wflag As String * 1
Dim wfecha1  As String
Dim wfecha2  As String
Dim wcadena As String
Dim wvalor  As String * 1
Dim sTmpTipMov As String
' CHEQUE DE MES CERRADO 0 ABIERTO
    wcadena = ""
    SQ_OPER = 2
    PUB_TIPREG = 155
    PUB_CODCIA = LK_CODCIA
    LEER_TAB_LLAVE
    If tab_mayor.EOF Then
      MsgBox "Crear tab_tipreg = 155 para seguridas", 48, Pub_Titulo
    Else
        Do Until tab_mayor.EOF
          If tab_mayor!TAB_NUMTAB = Val(Format(LK_FECHA_COP1, "yyyy")) Then
            wcadena = Trim(tab_mayor!tab_nomlargo)
          End If
          tab_mayor.MoveNext
        Loop
        wvalor = Mid(wcadena, LK_NRO_MES + 1, 1)
        If wvalor = "1" Then
           MsgBox "<<<< Mes CERRADO Operaciones >>>>", vbCritical, Pub_Titulo
           Exit Sub
        End If
    End If

'pub_mensaje = "Confirmar el Proceso de Pase al Diario General ?"
'Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
'If Pub_Respuesta = vbNo Then
'   Exit Sub
'End If

'cmdproce.Enabled = False
'proce.Enabled = False
'cmdrepo.Enabled = False
'cmdcli.Enabled = False
    DoEvents
    Label1.Caption = "Procesando ..."
    DoEvents
    CUENTA_PRO = 0
    wpub_cadena = ""
'If LK_EMP = "PIU" Then
    wfecha1 = Format(LK_FECHA_COP1, "dd/mm/yyyy")
    wfecha2 = Format(LK_FECHA_COP2, "dd/mm/yyyy")
'Else
' wfecha1 = Format(LK_FECHA_COP1, "yyyy/mm/dd")
' wfecha2 = Format(LK_FECHA_COP2, "yyyy/mm/dd")
'End If
'wpub_cadena = "DELETE COMOV  WHERE COV_FLAG_AUTOMATICA = '" & Trim(LOC_PROCESO) & "'  AND COV_CODCIA = '" & LK_CODCIA & "' AND COV_FECHA_VOUCHER >=  '" & wfecha1 & "'  AND COV_FECHA_VOUCHER <=  '" & wfecha2 & "' and COV_NRO_MES = " & LK_NRO_MES

    wpub_cadena = "DELETE COMOV  WHERE COV_CODCIA = '" & LK_CODCIA & "' AND COV_FECHA_VOUCHER >=  '" & wfecha1 & "'  AND COV_FECHA_VOUCHER <=  '" & wfecha2 & "' and COV_NRO_MES = " & LK_NRO_MES
    CN.Execute wpub_cadena, rdExecDirect
    'LOC_PROCESO = 1
PROCESO_SIGUE:
    DoEvents
    
    WSUM_GEN_H = 0
    WSUM_GEN_D = 0
    wc_sum_debe = 0
    wc_sum_haber = 0
    wc_importe = 0
    CUENTA_PRO = CUENTA_PRO + 1
    LOC_PROCESO = CUENTA_PRO ' Val(Left(proce.Text, 3))
    
    SQ_OPER = 1
    PUB_CODCIA = "00"
    PUB_TIPREG = 150
    PUB_NUMTAB = LOC_PROCESO
    LEER_TAB_LLAVE
    If Not tab_llave.EOF Then
        Label1.Caption = "Procesando ..." & tab_llave("TAB_NOMLARGO")
        sTmpTipMov = tab_llave("TAB_NOMLARGO")
    End If
    
    Label1.Visible = True
    Barra.Visible = True
    PSPR_LLAVE(0) = LK_CODCIA
    PSPR_LLAVE(1) = LOC_PROCESO
    PSPR_LLAVE(2) = LK_FECHA_COP1
    PSPR_LLAVE(3) = LK_FECHA_COP2
    pr_llave.Requery
    If pr_llave.EOF Then
      'Label1.Visible = False
      Barra.Visible = False
      cop_llave.Requery
      cop_llave.Edit
      If LOC_PROCESO = 1 Then
         cop_llave!cop_FLAG_REGC = "A"
      ElseIf LOC_PROCESO = 2 Then
         cop_llave!cop_FLAG_REGV = "A"
      ElseIf LOC_PROCESO = 3 Then
         cop_llave!cop_FLAG_CAJA = "A"
      End If
      cop_llave!cop_FLAG_DES = " "
      cop_llave!COP_FLAG_MAYORIZACION = " "
      cop_llave.Update
      'MsgBox "No existen movimientos ", 48, Pub_Titulo
    GoTo ava
    '  Exit Sub
    End If

    Wflag = ""
    PSCOV_VOUCHER(0) = LK_CODCIA
    PSCOV_VOUCHER(1) = LK_FECHA_COP1
    PSCOV_VOUCHER(2) = LK_FECHA_COP2
    'Debug.Print PSCOV_VOUCHER.SQL
    cov_voucher.Requery
    If cov_voucher.EOF Then
     ws_nro_voucher = 0
    Else
     cov_voucher.MoveLast
     ws_nro_voucher = cov_voucher!COV_NRO_VOUCHER
    End If
    ws_nro_voucher = ws_nro_voucher + 1
    
    Barra.Max = pr_llave.RowCount
    Barra.Min = 0
    Barra.Value = 0
    wc_codcta = pr_llave!MOV_CODCTA

    WSGLOSA = "Por " & sTmpTipMov & " : " & Format(LK_FECHA_COP2, "mmmm")
    

WS_NRO_MOV = 0
kl_voucher = 0
Do Until pr_llave.EOF
  Barra.Value = Barra.Value + 1
  If Trim(wc_codcta) <> Trim(pr_llave!MOV_CODCTA) Then
'  If Trim(wc_codcta) = "40110" Then Stop
'   If wc_sum_debe <> wc_sum_haber Then Stop
    wc_importe = wc_sum_debe
    wdh = "D"
    GoSub GRABA
    wdh = "H"
    wc_importe = wc_sum_haber
    GoSub GRABA
    wc_sum_debe = 0
    wc_importe = 0
    wc_sum_haber = 0
    wc_codcta = Trim(pr_llave!MOV_CODCTA)
  End If
  If kl_voucher = 0 And Trim(pr_llave!MOV_FLAG_DES) = "A" Then
    If WSUM_GEN_H <> WSUM_GEN_D Then
      If WSUM_GEN_D > WSUM_GEN_H Then
       wdh = "H"
       wc_importe = WSUM_GEN_D - WSUM_GEN_H
      Else
       wdh = "D"
       wc_importe = WSUM_GEN_H - WSUM_GEN_D
      End If
      wcta = "389004"
      GoSub GRABA
      'If wdh = "H" Then
      '  wdh = "D"
      '  wcta = "389004"
      '  GoSub GRABA
      '  wdh = "H"
      '  wcta = "75910"
      '  GoSub GRABA
      'Else
      '  wdh = "H"
      '  wcta = "389004"
      '  GoSub GRABA
      '  wdh = "D"
      '  wcta = "94599"
      '  GoSub GRABA
     'End If
     WSUM_GEN_D = 0
     WSUM_GEN_H = 0
    End If
    ws_nro_voucher = ws_nro_voucher + 1
    kl_voucher = 1
  End If
  wc_importe = Val(pr_llave!MOV_IMPORTE)
  If pr_llave!MOV_MONEDA = "D" And Trim(pr_llave!MOV_FLAG_TC) = "" Then
     SQ_OPER = 1
     PUB_CAL_INI = pr_llave!MOV_fecha_EMI
     PUB_CAL_FIN = pr_llave!MOV_fecha_EMI
     PUB_CODCIA = LK_CODCIA
     LEER_CAL_LLAVE
     If Nulo_Valor0(cal_llave!cal_tipo_cambio) = 0 Then
       MsgBox "Falta tipo de cambio"
       Exit Sub
     End If
     wc_importe = Format((wc_importe * Val(cal_llave!cal_tipo_cambio)), "0.00")
  ElseIf Trim(pr_llave!MOV_MONEDA) = "D" And Trim(pr_llave!MOV_FLAG_TC) = "A" Then
      wc_importe = Format((wc_importe * Val(pr_llave!MOV_TIPO_CAMBIO)), "0.00")
  End If
  If pr_llave!MOV_DH = "D" Then
    wc_sum_debe = wc_sum_debe + wc_importe
    WSUM_GEN_D = WSUM_GEN_D + wc_importe
  Else
    wc_sum_haber = wc_sum_haber + wc_importe
    WSUM_GEN_H = WSUM_GEN_H + wc_importe
  End If
  wdh = pr_llave!MOV_DH
  wcta = pr_llave!MOV_CODCTA
  DoEvents
  pr_llave.MoveNext
  DoEvents
Loop
' siempre hay data el ultimo registro
wc_importe = wc_sum_debe
wdh = "D"
GoSub GRABA
wdh = "H"
wc_importe = wc_sum_haber
GoSub GRABA
If WSUM_GEN_H <> WSUM_GEN_D Then
   If WSUM_GEN_D > WSUM_GEN_H Then
      wdh = "H"
      wc_importe = WSUM_GEN_D - WSUM_GEN_H
   Else
      wdh = "D"
      wc_importe = WSUM_GEN_H - WSUM_GEN_D
   End If
   wcta = "389004"
   GoSub GRABA
   'If wdh = "H" Then
   '  wdh = "D"
   '  wcta = "389004"
   '  GoSub GRABA
   '  wdh = "H"
   '  wcta = "75910"
   '  GoSub GRABA
   'Else
   '  wdh = "H"
   '  wcta = "389004"
   '  GoSub GRABA
   '  wdh = "D"
   '  wcta = "94599"
   '  GoSub GRABA
   'End If
   
End If

 
wc_sum_debe = 0
wc_importe = 0
wc_sum_haber = 0
ava:
If CUENTA_PRO <> 19 Then 'mic cambiado por numero de movimiento =6
  'proce.ListIndex = CUENTA_PRO - 1
  GoTo PROCESO_SIGUE:
End If

' reporte a Medida
'cop_llave.Requery
'cop_llave.Edit
'cop_llave!cop_FLAG_DES = " "
'cop_llave!COP_FLAG_MAYORIZACION = " "
'Select Case Val(LOC_PROCESO)
'Case 1
'   REGISTROS 1
'   cop_llave!cop_FLAG_REGC = "A"
'Case 2
'   REGISTROS 2
 '  cop_llave!cop_FLAG_REGV = "A"
'Case 3
'   CAJABANCOS 3
'   cop_llave!cop_FLAG_CAJA = "A"
'Case 4

'End Select
'cop_llave.Update
'Barra.Visible = False
'Label1.Visible = False

ACT_MESES (0)
'MsgBox "Proceso Terminado", 48, Pub_Titulo
'Unload FrmProce
Exit Sub
GRABA:
   If wc_importe = 0 Then GoTo PASACOV
    cov_voucher.AddNew
    cov_voucher!COV_NRO_MOV = WS_NRO_MOV
    cov_voucher!COV_CODCIA = LK_CODCIA
    If LK_EMP_PTO = "A" Then
     cov_voucher!COV_CODCIA = wscodcia
    End If
    cov_voucher!COV_NUMTAB = 0
    cov_voucher!COV_NRO_VOUCHER = ws_nro_voucher
    cov_voucher!COV_FECHA_VOUCHER = LK_FECHA_COP2
    cov_voucher!COV_glosa = WSGLOSA
    cov_voucher!COV_FECHA_doc = LK_FECHA_DIA
    cov_voucher!COV_CODCTA = wcta
    cov_voucher!COV_DH = wdh
    cov_voucher!COV_IMPORTE = wc_importe
    cov_voucher!COV_ESTADO = "0"
    cov_voucher!COV_CODUSU = LK_CODUSU
    cov_voucher!cov_flag_automatica = Trim(LOC_PROCESO)
    cov_voucher!cov_nro_mes = LK_NRO_MES
    cov_voucher!COV_TIPMOV = LOC_PROCESO 'agregado por mic
    If wc_importe <> cov_voucher!COV_IMPORTE Then Stop
    cov_voucher.Update
    
    WS_NRO_MOV = WS_NRO_MOV + 1
PASACOV:
Return

End Sub
'************************
Private Sub MayorizarDoc()
Dim SQL As String
Dim iMes As Integer
Dim sMes As String
Dim periodo As Integer
Dim iCount As Integer
    Label1.Caption = "Procesando Saldo de Documentos"
    iMes = LK_NRO_MES
    sMes = Format(iMes, "00")
    periodo = Format(LK_FECHA_COP2, "yyyy")
    SQL = "UPDATE Docsal SET Doc_Deb" & sMes & " = 0, Doc_Hab" & sMes & " = 0, Doc_Flag = 'C' WHERE Doc_codcia = '" & LK_CODCIA & "' AND Doc_Periodo = " & CStr(periodo)
    CN.Execute SQL, rdExecDirect

    pub_cadena = "SELECT MOV_NRO_MES, SUM(MOV_IMPORTE) AS MOV_IMPORTE, MOV_SERIE, MOV_NUMFAC, MOV_CODCTA, MOV_TIPMOV, MOV_DH, MOV_RUC, MOV_CODCLIE, MOV_OPC, MOV_EXONERADO FROM MOVICONT " & _
                 "WHERE MOV_CODCIA = ? AND MOV_PERIODO = ? AND MOV_NRO_MES = ? AND MOV_SERIE <> 0 AND MOV_NUMFAC <> 0 " & _
                 "GROUP BY MOV_NRO_MES, MOV_SERIE, MOV_NUMFAC, MOV_DH, MOV_RUC, MOV_CODCLIE, MOV_CODCTA, MOV_TIPMOV, MOV_OPC, MOV_EXONERADO" 'ORDER BY MOV_PERIODO

    Set PSMOVICON1 = CN.CreateQuery("", pub_cadena)
    PSMOVICON1(0) = LK_CODCIA
    PSMOVICON1(1) = periodo
    PSMOVICON1(2) = iMes
    Barra.Visible = True
    Set MOVICONT1 = PSMOVICON1.OpenResultset(rdOpenKeyset, rdConcurValues)
    MOVICONT1.Requery
    iCount = 0
    If Not MOVICONT1.EOF Then Barra.Max = MOVICONT1.RowCount
    Do While Not MOVICONT1.EOF
        iCount = iCount + 1
        Barra.Value = iCount
        PSDOCSAL(0) = LK_CODCIA 'DOC_CODCIA
        PSDOCSAL(1) = Format(LK_FECHA_COP2, "yyyy") 'DOC_PERIODO
        PSDOCSAL(2) = MOVICONT1!MOV_serie
        PSDOCSAL(3) = MOVICONT1!MOV_numfac
        DOCSAL.Requery
        If DOCSAL.EOF Then
            SQL = "INSERT INTO DOCSAL(Doc_Flag, DOC_CODCIA, DOC_CODCLIE, DOC_CUENTA, DOC_TIPDOC, DOC_SERIE, DOC_NUMFAC, DOC_PERIODO, DOC_RUC, DOC_OPC, DOC_EXONERACION, DOC_MES)"
            SQL = SQL & "VALUES ('V', '" & LK_CODCIA & "'," & MOVICONT1!MOV_codclie & ",'" & Trim(MOVICONT1!MOV_CODCTA) & "'," & MOVICONT1!MOV_TIPMOV & ","
            SQL = SQL & MOVICONT1!MOV_serie & "," & MOVICONT1!MOV_numfac & "," & Format(LK_FECHA_COP2, "yyyy") & ",'" & Trim(MOVICONT1!MOV_RUC) & "','" & MOVICONT1!MOV_OPC & "','"
            SQL = SQL & MOVICONT1!MOV_EXONERADO & "','" & iMes & "')"
            CN.Execute SQL, rdExecDirect
        End If
            
        SQL = "UPDATE DOCSAL SET DOC_MES = " & CStr(iMes) & ", DOC_FECHA_EMI = '" & LK_FECHA_COP2 & "', DOC_FECHA = '" & LK_FECHA_COP2 & "', "
        If MOVICONT1!MOV_DH = "D" Then
            SQL = SQL & "DOC_DEB" & sMes & " = " & CStr(MOVICONT1!MOV_IMPORTE) & " "
        ElseIf MOVICONT1!MOV_DH = "H" Then
            SQL = SQL & "DOC_HAB" & sMes & " = " & CStr(MOVICONT1!MOV_IMPORTE) & " "
        End If
        SQL = SQL + "WHERE DOC_CODCIA = '" & LK_CODCIA & "' AND DOC_PERIODO = " & CStr(periodo) & " AND DOC_SERIE = " & CStr(MOVICONT1!MOV_serie) & " AND DOC_NUMFAC = " & CStr(MOVICONT1!MOV_numfac)
        CN.Execute SQL, rdExecDirect
        MOVICONT1.MoveNext
    Loop
    Barra.Value = 0
    Barra.Visible = False
    MsgBox "Proceso terminado...Ok", 48, Pub_Titulo
    
    Unload Me
End Sub
