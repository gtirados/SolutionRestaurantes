VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frm_mayoriz 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Proceso de Mayorización"
   ClientHeight    =   4050
   ClientLeft      =   1560
   ClientTop       =   1575
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
   Icon            =   "frm_mayor.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4050
   ScaleWidth      =   6600
   Tag             =   "55"
   Begin VB.TextBox fecha2 
      Height          =   375
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2040
      Width           =   1335
   End
   Begin VB.TextBox fecha1 
      Height          =   375
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   6015
      Begin ComctlLib.ProgressBar Barra 
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Tag             =   "0"
         Top             =   1080
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   503
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   5655
      End
   End
   Begin VB.CommandButton Mayorizar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "&Mayorizar"
      Height          =   855
      Left            =   1440
      Picture         =   "frm_mayor.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton SALIR 
      Caption         =   "Ce&rrar"
      Height          =   855
      Left            =   3600
      Picture         =   "frm_mayor.frx":058C
      Style           =   1  'Graphical
      TabIndex        =   0
      Tag             =   "9999"
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Hasta:"
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   7
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Desde :"
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   2040
      Width           =   855
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

Option Explicit


Private Sub Form_Load()
WSELE = ""
Dim ws_fecha
Dim ws_indice As Integer
Dim cade
Frame1.Visible = False


'cade = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND COV_FECHA_VOUCHER>=? AND COV_NRO_MOV=? AND COV_NRO_MES = " & LK_NRO_MES & " "
'Set PSCO2_CUENTA = CN.CreateQuery("", cade)
'PSCO2_CUENTA(0) = 0
'PSCO2_CUENTA(1) = LK_FECHA_DIA
'Set co2_cuenta = PSCO2_CUENTA.OpenResultset(rdOpenKeyset, rdConcurValues)


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

cade = "SELECT COV_CODCTA, COV_DH , COV_IMPORTE FROM COMOV WHERE COV_CODCIA = ? AND COV_FECHA_VOUCHER >= ? AND COV_FECHA_VOUCHER <= ? AND COV_NRO_MES = " & LK_NRO_MES & " ORDER BY COV_CODCTA, COV_FECHA_VOUCHER, COV_NRO_MOV"

cade = "SELECT COV_CODCTA, COV_DH,  SUM(COV_IMPORTE) AS TOT FROM COMOV WHERE COV_CODCIA = ? AND COV_FECHA_VOUCHER >= ? AND COV_FECHA_VOUCHER <= ? AND COV_NRO_MES = " & LK_NRO_MES & "  GROUP BY COV_CODCTA, COV_DH"
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


FECHA1.Text = LK_FECHA_COP1
fecha2.Text = LK_FECHA_COP2


End Sub

Private Sub Form_Terminate()
Unload frm_mayoriz
End Sub

Private Sub Mayorizar_Click()
frm_mayoriz.PROMARY
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
PSCOX_CUENTA.rdoParameters(1) = CDate(FECHA1.Text)
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
    PUB_CODCIA = "03"
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

            
            




PSCOV_CUENTA.rdoParameters(0) = LK_CODCIA
PSCOV_CUENTA.rdoParameters(1) = FECHA1.Text
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

PASA2:

PSCOV_CUENTA.rdoParameters(0) = LK_CODCIA
PSCOV_CUENTA.rdoParameters(1) = FECHA1.Text
PSCOV_CUENTA.rdoParameters(2) = fecha2.Text
cov_cuenta.Requery
'cov_cuenta.MoveFirst
PUB_CUENTA = cov_cuenta!COV_CODCTA
'NEW_DH = cov_cuenta!COV_DH
Barra.Min = 0
Barra.Value = 0
Barra.Max = cov_cuenta.RowCount

CONTADOR = 0


Label1.Caption = "Sumando Cuentas Contables..."
DoEvents
'cov_cuenta.MoveFirst

Do Until cov_cuenta.EOF
      Barra.Value = Barra.Value + 1
      ws_tot_debe = 0
      ws_tot_haber = 0
      If cov_cuenta!COV_DH = "D" Then
        ws_tot_debe = cov_cuenta!TOT
      Else
        ws_tot_haber = cov_cuenta!TOT
      End If
      PUB_CUENTA = cov_cuenta!COV_CODCTA
      cov_cuenta.MoveNext
      If cov_cuenta.EOF Then
         ACT_SALDO_MES PUB_CUENTA, ws_tot_debe, ws_tot_haber
         Exit Do
      End If
      If PUB_CUENTA = cov_cuenta!COV_CODCTA Then
            If cov_cuenta!COV_DH = "D" Then
              ws_tot_debe = cov_cuenta!TOT
            Else
              ws_tot_haber = cov_cuenta!TOT
            End If
      Else
          cov_cuenta.MovePrevious
      End If
      ACT_SALDO_MES PUB_CUENTA, ws_tot_debe, ws_tot_haber
      cov_cuenta.MoveNext
Loop



MAYORIZO:
Label1.Caption = "Mayorizando . . ."
DoEvents
Barra.Value = 0
Barra.Min = 0
Barra.Max = 3
ws_niv1 = NIVEL_MAX
Do Until ws_niv1 = 1
'   barra.Value = barra.Value + 1
   'cade = "SELECT * FROM COMAEST WHERE
   'COM_CODCIA = ?  AND ( COM_NIVEL = ? OR COM_NIVEL= ? ) ORDER BY COM_CUENTA"
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
      If Val(com_cuenta_sup!com_nivel) = ws_niv1 Then
        saldo_d = com_cuenta_sup!COM_DEB_MES + saldo_d
        saldo_h = com_cuenta_sup!COM_HAB_MES + saldo_h
      Else
        SQ_OPER = 1
        PUB_CUENTA = WS_CUENTA
'        If Left(PUB_CUENTA, 2) = "97" Then Stop
        PUB_CODCIA = LK_CODCIA
        LEER_COM_LLAVE
        com_llave.Edit
        com_llave!COM_DEB_MES = saldo_d
        com_llave!COM_HAB_MES = saldo_h
        com_llave.Update
        ACT_SALDO_MES PUB_CUENTA, saldo_d, saldo_h
        WS_CUENTA = com_cuenta_sup!com_cuenta
        saldo_d = 0
        saldo_h = 0
      End If
      
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
            saldo = (Val(com_llave!COM_DEB_MES) + Val(com_llave!COM_DEB_ANO)) * com_llave!com_signo_d + (Val(com_llave!COM_HAB_MES) + Val(com_llave!COM_HAB_ANO)) * com_llave!com_signo_h
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
            WS_SALDO = (Val(com_cuenta!COM_HAB_MES) + Val(com_cuenta!COM_HAB_ANO)) * com_cuenta!com_signo_h + (Val(com_cuenta!COM_DEB_MES) + Val(com_cuenta!COM_DEB_MES)) * com_cuenta!com_signo_d
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
Label1.Caption = ""
Barra.Visible = False
MsgBox "Proceso terminado...Ok", 48, Pub_Titulo


fin:
Unload frm_mayoriz


End Sub
