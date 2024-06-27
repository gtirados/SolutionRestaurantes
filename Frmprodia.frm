VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form PRODIA 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Actualización de Fechas"
   ClientHeight    =   5400
   ClientLeft      =   300
   ClientTop       =   1770
   ClientWidth     =   3885
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Frmprodia.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   3885
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Situación de Operaciones de la Compañia: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1215
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   3615
      Begin VB.PictureBox poperativo 
         AutoSize        =   -1  'True
         Height          =   540
         Left            =   2760
         Picture         =   "Frmprodia.frx":0442
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.PictureBox pbloqueado 
         AutoSize        =   -1  'True
         Height          =   540
         Left            =   2760
         Picture         =   "Frmprodia.frx":0884
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   540
      End
      Begin VB.OptionButton option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Operativo."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   2295
      End
      Begin VB.OptionButton option1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Bloqueado."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   11
         Top             =   720
         Width           =   2415
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cerrar Operaciones del Día."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   720
      Picture         =   "Frmprodia.frx":0CC6
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3840
      Width           =   2415
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ce&rrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1200
      Picture         =   "Frmprodia.frx":1108
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4560
      Width           =   1455
   End
   Begin ComctlLib.ProgressBar ProgBar 
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   3360
      Visible         =   0   'False
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   344
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   3615
      Begin VB.ListBox EMP 
         BackColor       =   &H008B4914&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   840
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   7
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cierre de Compañia(s) :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   2295
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   3480
      Top             =   3840
   End
   Begin VB.Label lblbarraos 
      Alignment       =   2  'Center
      BackColor       =   &H008B4914&
      Caption         =   "Solution for Business"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   5160
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Día:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   615
   End
   Begin VB.Label LblFecha 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label POR 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cerrando Operaciones Diarias..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   2
      Left            =   600
      TabIndex        =   4
      Top             =   3600
      Visible         =   0   'False
      Width           =   2640
   End
   Begin VB.Label lblcierre 
      AutoSize        =   -1  'True
      BackColor       =   &H00808000&
      BackStyle       =   0  'Transparent
      Caption         =   "CIERRE DIARIO DE OPERACIONES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   255
      TabIndex        =   0
      Top             =   60
      Width           =   3255
   End
End
Attribute VB_Name = "PRODIA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WS_SALDO As Currency
Dim WS_SALDO_D As Currency
Dim PSPRE_MAYOR2  As rdoQuery
Dim PSMES_MAYOR2  As rdoQuery
Dim pre_mayor2 As rdoResultset
Dim mes_mayor2 As rdoResultset
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset
Dim stock_llave As rdoResultset
Dim PSST_LLAVE As rdoQuery
Dim chedef_llave As rdoResultset
Dim PSCHE_DEF  As rdoQuery






Public Sub REPO_CAJA_GEN2(ww_codcia As String)
Dim ww_moneda
WS_SALDO = 0
WS_SALDO_D = 0

PUB_FECHA = LK_FECHA_DIA
pu_codcia = ww_codcia
SQ_OPER = 1
LEER_ALL_LLAVE
If all_llave.EOF Then Exit Sub
PUB_CODCIA = ww_codcia
LEER_PAR_LLAVE


   WS_SALDO = Nulo_Valor0(par_llave!PAR_SALDO_CAJA_ayer)
   WS_SALDO_D = Nulo_Valor0(par_llave!PAR_SALDO_CAJA_D_ayer)

Do Until all_llave.EOF
   If all_llave!ALL_SIGNO_CAJA = 0 Then GoTo OTRO
   
   
   If all_llave!all_flag_ext = "E" Then GoTo OTRO
   
   If all_llave!ALL_SIGNO_CAR = 0 And all_llave!ALL_tipmov = 0 Then
      WS_IMPORTE = all_llave!ALL_IMPORTE
   Else
      WS_IMPORTE = all_llave!ALL_IMPORTE_AMORT
   End If
   
   
   If Trim(all_llave!ALL_moneda_ccm) <> " " And val(all_llave!all_codban) <> 0 Then
      ww_moneda = all_llave!ALL_moneda_ccm
   ElseIf Trim(all_llave!ALL_MONEDA_CLI) <> " " And val(all_llave!ALL_CODCLIE) <> 0 Then
      ww_moneda = all_llave!ALL_MONEDA_CLI
   ElseIf Trim(all_llave!ALL_MONEDA_CAJA) <> " " Then
      ww_moneda = all_llave!ALL_MONEDA_CAJA
   End If
   
   If ww_moneda = "S" Then
   If all_llave!ALL_SIGNO_CAJA = 1 Then
      WS_SALDO = WS_SALDO + WS_IMPORTE
   Else
      WS_SALDO = WS_SALDO - WS_IMPORTE
   End If
   End If
   If ww_moneda = "D" Then
   If all_llave!ALL_SIGNO_CAJA = 1 Then
      WS_SALDO_D = WS_SALDO_D + WS_IMPORTE
   Else
      WS_SALDO_D = WS_SALDO_D - WS_IMPORTE
   End If
   End If
OTRO:
  all_llave.MoveNext
  Loop
  
par_llave.Edit
par_llave!PAR_SALDO_CAJA_HOY = WS_SALDO
par_llave!PAR_SALDO_CAJA_D_HOY = WS_SALDO_D
par_llave.Update



End Sub

Private Sub Command1_Click()
Dim CONTADOR As Long
Dim ww_ult_oper As Integer
Dim WS_SALDO_S As Currency
Dim WS_SALDO_D As Currency

Dim WS_CODTRA As Integer
Dim WW_FECHA As Date
Dim WS_SALDO2 As Currency
Dim WS_BLOQ1, WS_BLOQ2 As String
Dim ws_saldo_caa As Currency
Dim WS_MONEDA As String * 1
Dim ww_dias As Integer
Dim wcodven As Integer
Dim WDOCU As String
Dim xcuenta As Integer
Dim ws_codcia As String * 2

Dim CajaAyerSTmp As Currency
Dim CajaAyerDTmp As Currency

ww_dias = 0
wcodven = 0
ws_saldo_caa = 0

pub_cadena = "SELECT * FROM MESAS WHERE MES_ESTADO <> 'L' AND MES_CODCIA = ?  "
Set PSMES_MAYOR2 = CN.CreateQuery("", pub_cadena)
PSMES_MAYOR2(0) = LK_CODCIA
Set mes_mayor2 = PSMES_MAYOR2.OpenResultset(rdOpenKeyset, rdConcurValues)

If Not mes_mayor2.EOF Then
MsgBox "Faltan procesar algunas mesas, por favor procesarlas antes de proceder"
 Exit Sub
End If


LimpiaParametros oCmdEjec
oCmdEjec.CommandText = "SP_DELIVERY_VERIFICA"
 oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@CODCIA", adChar, adParamInput, 2, LK_CODCIA)
oCmdEjec.Parameters.Append oCmdEjec.CreateParameter("@FECHA", adDBTimeStamp, adParamInput, , LK_FECHA_DIA)
                
Set oRSfp = oCmdEjec.Execute
If Not oRSfp.EOF Then
    If CBool(oRSfp!Dato) Then
        MsgBox "Hay Deliverys sin Liquidar." & vbCrLf & oRSfp!pedidos, vbInformation, Pub_Titulo
        Exit Sub
    End If
End If


If GEN!gen_cierre_todas = 0 Then
    pub_mensaje = "CIERRE DEL DIA DE LA EMPRESA : " & Trim(par_llave!PAR_NOMBRE) & " ¿Desea Continuar... ?"
Else
    pub_mensaje = "CIERRE DEL DIA DE EMPRESA(S) !!! ...   ¿Desea Continuar... ?"
End If
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If
If par_llave!PAR_FECHA_DIA <> LK_FECHA_DIA Then
   MsgBox "!La Fecha del Sistema ha cambiado, ya se cerro el día de la compañia ..!, el sistema se cerrará ", 48, Pub_Titulo
   End
End If

'If Nulo_Valor0(GEN!GEN_TASA_VENTA) = 99 Then   ' quitado por GTS para no correr cancelacion automatica de cheques
 '  CANCEL_CH
'End If


WDOCU = "Inicio de Operaciones "
Dim COS, CAA, PRE, CCMM As rdoResultset
Dim PSCAR, PSCOS, PSCCMM As rdoQuery

pub_cadena = "SELECT FAR_COSTEO_REAL FROM FACART WHERE FAR_CODCIA = ? AND FAR_COSTEO_REAL = 'A'"
Set PSCOS = CN.CreateQuery("", pub_cadena)
PSCOS(0) = LK_CODCIA
Set COS = PSCOS.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM PRECIOS "
Set PRE = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)
pub_cadena = "SELECT * FROM CCMAEST"
Set CCMM = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM CARTERA WHERE CAR_CODCIA = '" & LK_CODCIA & "' AND CAR_IMPORTE > 0 AND CAR_CP='C'  AND  CAR_FECHA_VCTO_ORIG <= ? "
Set PSCAR = CN.CreateQuery("", pub_cadena)
PSCAR(0) = LK_FECHA_DIA
Set car = PSCAR.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = ? "
Set PSPRE_MAYOR2 = CN.CreateQuery("", pub_cadena)
PSPRE_MAYOR2(0) = LK_CODCIA
Set pre_mayor2 = PSPRE_MAYOR2.OpenResultset(rdOpenKeyset, rdConcurValues)


Timer1.Enabled = False
lblcierre.Visible = True
Command1.Enabled = False

ProgBar.Visible = True
DoEvents
'POR(0).Visible = True
DoEvents
'POR(1).Visible = True
DoEvents
POR(2).Visible = True
DoEvents
CONTADOR = 0
WS_BLOQ1 = ""
WS_BLOQ2 = ""

'On Error GoTo SALIR
CN.Execute "BEGIN TRANSACTION", rdExecDirect

CCMM.Requery
Do Until CCMM.EOF
    CCMM.Edit
    CCMM!CCM_SAL_ANTERIOR = CCMM!CCM_SALDO
    CCMM.Update
    CCMM.MoveNext
Loop
If LK_EMP_PTO = "A" Then
    xcuenta = 1
    For fila = 1 To 30
      ws_codcia = Mid(Trim(GEN!gen_ART_CIAS), xcuenta, 2)
      If Trim(ws_codcia) = "" Then Exit For
      GoSub cierra_cia
      PS_REP01(0) = ws_codcia
      llave_rep01.Requery
      GoSub PASA_POR_CIAS
      xcuenta = xcuenta + 2
    Next fila
Else
     ws_codcia = LK_CODCIA
     GoSub cierra_cia
     PS_REP01(0) = LK_CODCIA
     llave_rep01.Requery
     GoSub PASA_POR_CIAS
End If

CN.Execute "commit TRANSACTION", rdExecDirect
 
MsgBox "Proceso de Cierre Terminado Satisfactoriamente... Ahora el Sistema se Cerrara...", 48, Pub_Titulo
End
Exit Sub


PASA_POR_CIAS:
   llave_rep01.Requery
   SQ_OPER = 1
   PUB_CODCIA = llave_rep01!PAR_CODCIA
   PUB_CAL_INI = llave_rep01!PAR_FECHA_DIA
   PUB_CAL_FIN = DateAdd("m", 1, llave_rep01!PAR_FECHA_DIA)
   LEER_CAL_LLAVE 1
   If cal_llave.EOF Then
      MsgBox "NO PUEDE SER.... ERROR GRAVE EN CAL"
      GoTo SALIDA_ERROR
   End If
   
   cal_llave.Edit
   cal_llave!CAL_INDICE = 3
   cal_llave.Update
   CONTADOR = 1
   cal_llave.MoveNext
   Do Until cal_llave!CAL_LABORABLE = "S" Or cal_llave.EOF
      CONTADOR = CONTADOR + 1
      cal_llave.MoveNext
   Loop
   If cal_llave.EOF Then
      MsgBox "Falta las fechas .... "
      GoTo SALIDA_ERROR
   End If
   cal_llave.Edit
   cal_llave!CAL_INDICE = 1
   cal_llave.Update
   GoSub manda_numero
   WS_CODTRA = 9999
   SQ_OPER = 1
   PUB_NUMTAB = 0
   PUB_CODCIA = llave_rep01!PAR_CODCIA
   PUB_TIPREG = 1000
   LEER_TAB_LLAVE
   WS_SALDO = 0
   'ICA
   If LK_FLAG_GRIFO = "A" Or LK_EMP = "3AA" Or LK_EMP = "PIU" Or LK_EMP = "PAR" Then 'Or LK_EMP = "HER"
   Else
    'WS_SALDO = tab_llave!TAB_contable2
   End If
    '*
    llave_rep01.Requery
    llave_rep01.Edit
    'llave_rep01!par_flag_cierre = 0
    CajaAyerDTmp = llave_rep01!PAR_SALDO_CAJA_D_HOY
    CajaAyerSTmp = llave_rep01!PAR_SALDO_CAJA_HOY
    llave_rep01!PAR_SALDO_CAJA_HOY = cAJA("S")   '* Nulo_Valor0(llave_rep01!par_saldo_caja_hoy)
    llave_rep01!PAR_SALDO_CAJA_D_HOY = cAJA("D")   '*Nulo_Valor0(llave_rep01!PAR_SALDO_CAJA_D_HOY)
    llave_rep01!PAR_FECHA_DIA = cal_llave!CAL_FECHA
    llave_rep01.Update

   
   
   GoSub GRABA_ALLOG
   PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
   GoSub PROCESA_CAR
   ProgBar.Min = 0
   CONTADOR = 0
   PSPRE_MAYOR2.rdoParameters(0) = llave_rep01!PAR_CODCIA
   pre_mayor2.Requery
   CONTADOR = 0
   If Not pre_mayor2.EOF Then ProgBar.max = pre_mayor2.RowCount
   Do Until pre_mayor2.EOF
         pre_mayor2.Edit
         pre_mayor2!PRE_cosTO_ant = Nulo_Valor0(pre_mayor2!PRE_COSTO)
         pre_mayor2.Update
         pre_mayor2.MoveNext
         CONTADOR = CONTADOR + 1
         ProgBar.Value = CONTADOR
         DoEvents
   Loop
   If LK_FLAG_GRIFO <> "A" Then GoTo salta_grifo
   SQ_OPER = 2
   PUB_KEY = 0
   pu_codcia = LK_CODCIA
   LEER_ART_LLAVE
   Do Until art_mayor.EOF
     art_mayor.Edit
     art_mayor!art_cash = val(Nulo_Valor0(art_mayor!art_cash)) + val(Nulo_Valor0(art_mayor!art_margen))
     art_mayor.Update
     art_mayor.MoveNext
   Loop
salta_grifo:
    llave_rep01.Requery
    llave_rep01.Edit
    llave_rep01!par_flag_cierre = 0
    llave_rep01!PAR_SALDO_CAJA_ayer = CajaAyerSTmp '* Nulo_Valor0(llave_rep01!par_saldo_caja_hoy)
    llave_rep01!PAR_SALDO_CAJA_D_ayer = CajaAyerDTmp  '*Nulo_Valor0(llave_rep01!PAR_SALDO_CAJA_D_HOY)
    llave_rep01!PAR_FECHA_DIA = cal_llave!CAL_FECHA
    llave_rep01.Update

Return



PROCESA_CAR:
If Nulo_Valor0(par_llave!PAR_DIAS_LARGE) = 0 Then Return

WW_FECHA = DateAdd("d", cal_llave!CAL_FECHA, Nulo_Valor0(par_llave!PAR_DIAS_LARGE) * -1)

PSCAR.rdoParameters(0) = WW_FECHA
car.Requery
Do Until car.EOF
   SQ_OPER = 1
   pu_codclie = car!CAR_CODCLIE
   pu_codcia = car!car_codcia
   pu_cp = "C"
   LEER_CLI_LLAVE
   ww_dias = DateDiff("d", car!car_fecha_vcto_orig, cal_llave!CAL_FECHA)
   If car!car_NUMFAC <> 0 Then
     WDOCU = car!car_FBG & " / " & Format(car!car_NUMSER, "000") & " - " & Format(car!car_NUMFAC, "00000000")
   Else
     WDOCU = car!car_TIPDOC
     If car!car_TIPDOC = "CH" Then
        WDOCU = car!car_TIPDOC & " / " & Format(car!car_NUM_CHEQUE, "00000000")
     End If
   End If
   wcodven = car!CAR_codven
   ws_saldo_caa = car!car_importe
   If Nulo_Valors(cli_llave!CLI_TIPO_BLOQ1) <> "1" Then
      WS_BLOQ1 = cli_llave!CLI_TIPO_BLOQ1 & Nulo_Valors(cli_llave!CLI_TIPO_BLOQ2) & (cli_llave!CLI_TIPO_BLOQ3) & Nulo_Valors(cli_llave!CLI_TIPO_BLOQ4)
      cli_llave.Edit
      cli_llave!CLI_TIPO_BLOQ1 = "1"
      cli_llave.Update
      WS_BLOQ2 = cli_llave!CLI_TIPO_BLOQ1 & Nulo_Valors(cli_llave!CLI_TIPO_BLOQ2) & (cli_llave!CLI_TIPO_BLOQ3) & Nulo_Valors(cli_llave!CLI_TIPO_BLOQ4)
      WS_CODTRA = 2582
      GoSub GRABA_ALLOG
      PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
   End If
   
   car.MoveNext
Loop

Return


manda_numero:
SQ_OPER = 2
PUB_FECHA = cal_llave!CAL_FECHA
pu_codcia = llave_rep01!PAR_CODCIA
LEER_ALL_LLAVE
If all_menor.EOF = False Then
   PUB_NUM_OPER_XXX = all_menor!ALL_NUMOPER
Else
   PUB_NUM_OPER_XXX = 0
End If
PUB_NUM_OPER_XXX = PUB_NUM_OPER_XXX + 1
Return

cierra_cia:
      PS_REP01(0) = ws_codcia
      llave_rep01.Requery
    If llave_rep01!par_flag_cierre <> 9 Then
       MsgBox "!!! Falta realizar Cierre de Operaciones ..." & llave_rep01!PAR_CODCIA & "-" & llave_rep01!PAR_NOMBRE
       GoTo SALIDA_ERROR
       Exit Sub
    End If
'    If llave_rep01!par_flag_costos <> 9 Then
'       MsgBox "!!! Falta procesar Costos  del Dia ...En " & llave_rep01!PAR_CODCIA & " - " & Trim(llave_rep01!PAR_NOMBRE), 48, Pub_Titulo
'       GoTo SALIDA_ERROR
'    End If
    
    If LK_EMP = "HER" Or LK_FLAG_GRIFO = "A" Or LK_EMP = "3AA" Or LK_EMP = "PIU" Or LK_EMP = "PAR" Or (LK_ICA = "A") Then
       GoTo PASE_CAJA
    End If
    If LK_FLAG_SOS = "A" Then GoTo PASE_CAJA
    SQ_OPER = 1
    pu_codcia = llave_rep01!PAR_CODCIA
    PUB_FECHA = LK_FECHA_DIA
    LEER_ALL_LLAVE
    all_llave.MoveLast
    If all_llave.EOF Then
       MsgBox "El Sistema Esta Cerrando sin Movimientos ... !!!", 48, Pub_Titulo
       GoTo PASE_CAJA
    End If
    SQ_OPER = 1
    PUB_NUMTAB = 0
    PUB_CODCIA = llave_rep01!PAR_CODCIA
    PUB_TIPREG = 1000
    LEER_TAB_LLAVE
    If tab_llave.EOF Then
       MsgBox "!!! Falta procesar Caja Soles  del Dia ...En " & llave_rep01!PAR_CODCIA & " - " & Trim(llave_rep01!PAR_NOMBRE), 48, Pub_Titulo
       GoTo SALIDA_ERROR
       Exit Sub
    Else
      If ((val(all_llave!ALL_NUMOPER) <> val(tab_llave!tab_NOMLARGO)) Or (CDate(tab_llave!tab_nomcorto) <> LK_FECHA_DIA)) And val(tab_llave!tab_NOMLARGO) <> 0 Then
       MsgBox "!!! Falta procesar Caja Soles   del Dia ...Ir a Contabilidad " & llave_rep01!PAR_CODCIA & " - " & Trim(llave_rep01!PAR_NOMBRE), 48, Pub_Titulo
       GoTo SALIDA_ERROR
       Exit Sub
      Else
        WS_SALDO_S = tab_llave!TAB_contable2
      End If
    End If
    
    SQ_OPER = 1
    PUB_NUMTAB = 0
    PUB_CODCIA = llave_rep01!PAR_CODCIA
    PUB_TIPREG = 1001
    LEER_TAB_LLAVE
    If tab_llave.EOF Then
       MsgBox "!!! Falta procesar Caja  Dollares (U$$) del Dia ...En " & llave_rep01!PAR_CODCIA & " - " & Trim(llave_rep01!PAR_NOMBRE), 48, Pub_Titulo
       GoTo SALIDA_ERROR
       Exit Sub
    Else
      If (val(all_llave!ALL_NUMOPER) <> val(tab_llave!tab_NOMLARGO) Or tab_llave!tab_nomcorto <> LK_FECHA_DIA) And val(tab_llave!tab_NOMLARGO) <> 0 Then
       MsgBox "!!! Falta procesar Caja  Dollares (U$$) del Dia ...En " & llave_rep01!PAR_CODCIA & " - " & Trim(llave_rep01!PAR_NOMBRE), 48, Pub_Titulo
       GoTo SALIDA_ERROR
       Exit Sub
      Else
      WS_SALDO_D = tab_llave!TAB_contable2
      End If
    End If
    
    llave_rep01.Edit
    llave_rep01!PAR_SALDO_CAJA_HOY = WS_SALDO_S
    llave_rep01!PAR_SALDO_CAJA_D_HOY = WS_SALDO_D
    llave_rep01.Update
    
PASE_CAJA:
' AGREGE ACV
GoTo PASO
    PSCOS.rdoParameters(0) = llave_rep01!PAR_CODCIA
    COS.Requery
    Do Until COS.EOF
        COS.Edit
        COS!FAR_COSTEO_REAL = " "
        COS.Update
        COS.MoveNext
    Loop
PASO:
Return


GRABA_ALLOG:
all_llave.AddNew
all_llave!ALL_NUMOPER = PUB_NUM_OPER_XXX

all_llave!all_CODCIA = llave_rep01!PAR_CODCIA
all_llave!ALL_CODTRA = WS_CODTRA
all_llave!all_flag_ext = "E"
all_llave!ALL_CODCLIE = pu_codclie
all_llave!ALL_CODART = 0
all_llave!ALL_IMPORTE_AMORT = 0
all_llave!all_codusu = LK_CODUSU
all_llave!ALL_FBG = ""
all_llave!ALL_CODVEN = wcodven
all_llave!ALL_IMPORTE = llave_rep01!PAR_SALDO_CAJA_HOY
all_llave!ALL_IMPORTE_DOLL = llave_rep01!PAR_SALDO_CAJA_D_HOY
all_llave!ALL_NUMDOC = 0
all_llave!ALL_CP = pu_cp
all_llave!ALL_TIPDOC = ""
all_llave!all_numfac_c = 0
all_llave!all_numser_c = 0
all_llave!all_codban = 0
all_llave!all_concepto = WDOCU
all_llave!all_chenum = 0
all_llave!ALL_FECHA_DIA = cal_llave!CAL_FECHA
all_llave!ALL_FECHA_SUNAT = cal_llave!CAL_FECHA
all_llave!ALL_FECHA_VCTO = cal_llave!CAL_FECHA
all_llave!ALL_CANTIDAD = ww_dias
all_llave!ALL_NUMSER = 0
all_llave!all_numfac = 0
all_llave!all_neto = 0
all_llave!ALL_BRUTO = ws_saldo_caa
all_llave!ALL_tipmov = 0
all_llave!ALL_IMPTO = 0
all_llave!ALL_flete = 0
all_llave!ALL_HORA = Now
all_llave!ALL_DESCTO = 0
all_llave!ALL_GASTOS = 0
all_llave!ALL_PRECIO = 0
all_llave!ALL_MONEDA_CLI = ""
all_llave!ALL_moneda_ccm = ""
all_llave!ALL_MONEDA_CAJA = ""
all_llave!all_SECUENCIA = 0
all_llave!ALL_SIGNO_CAR = 0
all_llave!ALL_SIGNO_CAJA = 0
all_llave!ALL_SIGNO_CCM = 0
all_llave!all_sIGNO_ARM = 0
all_llave!all_chenum = 0
all_llave!ALL_CHESEC = 0
all_llave!ALL_CHESER = 0
all_llave!ALL_SUBTRA = ""
all_llave!ALL_TIPO_BLOQ_ACT = WS_BLOQ2
all_llave!ALL_TIPO_BLOQ_ANT = WS_BLOQ1
all_llave!all_codtra_ext = 0
all_llave!ALL_TIPO_CAMBIO = 0
all_llave!ALL_RUC = 0
all_llave.Update
Return

SALIDA_ERROR:
SALIR:
CN.Execute "ROLLBACK TRANSACTION", rdExecDirect
Unload PRODIA
Exit Sub
fin:
MsgBox Err.Description
Resume Next
End Sub

Private Sub Command4_Click()
Unload PRODIA
End Sub

Private Sub Form_Activate()
Dim xcuenta  As Integer
Dim ws_codcia As Integer
If LK_EMP_PTO = "A" Then
   If LK_CODCIA <> "00" Then
     Screen.MousePointer = 0
     MsgBox "Se Encuentra en punto de Venta. El Cierre del día se ejecuta en la Compañia Central ", 48, Pub_Titulo
     Unload PRODIA
     Exit Sub
   End If
'   Label3.Caption = "ALMACEN Y PUNTOS DE VENTAS"
   Frame3 = "Estado de las Operaciones de  Empresas es: "
   option1(0).Caption = "D i s p o n i b l e s"
   option1(1).Caption = "C e r r a d a s"
   xcuenta = 1
   For fila = 1 To 30
      ws_codcia = Mid(Trim(GEN!gen_ART_CIAS), xcuenta, 2)
      If Trim(ws_codcia) = "" Then Exit For
      PS_REP01(0) = ws_codcia
      llave_rep01.Requery
      EMP.AddItem Trim(llave_rep01!PAR_CODCIA) + " " + Trim(llave_rep01!PAR_NOMBRE)
      xcuenta = xcuenta + 2
   Next fila
   lblFecha.Caption = Format(LK_FECHA_DIA, "dddd, d mmmm yyyy")
   EMP.ListIndex = 0
   SQ_OPER = 1
   PUB_CODCIA = LK_CODCIA
   LEER_PAR_LLAVE
   If par_llave!par_flag_cierre = 9 Then
      option1(1).Value = True
      option1(1).ForeColor = QBColor(12)
      option1(0).ForeColor = QBColor(0)
   Else
      option1(0).Value = True
      option1(0).ForeColor = QBColor(2)
      option1(1).ForeColor = QBColor(0)
   End If
Else
'   Label3.Caption = "C O M P A Ñ I A"
   SQ_OPER = 1
   PUB_CODCIA = LK_CODCIA
   LEER_PAR_LLAVE
   EMP.AddItem Trim(par_llave!PAR_CODCIA) + " " + Trim(par_llave!PAR_NOMBRE)
   lblFecha.Caption = Format(LK_FECHA_DIA, "dddd, dd Mmmm yyyy")
   EMP.ListIndex = 0
   If par_llave!par_flag_cierre = 9 Then
      option1(1).Value = True
      option1(1).ForeColor = QBColor(12)
      option1(0).ForeColor = QBColor(0)
      pbloqueado.Visible = True
      poperativo.Visible = False
   Else
      option1(0).Value = True
      option1(0).ForeColor = QBColor(2)
      option1(1).ForeColor = QBColor(0)
      pbloqueado.Visible = False
      poperativo.Visible = True
   End If
End If

If option1(0).Visible Then
'  Option1(0).SetFocus
End If

End Sub

Private Sub Form_DblClick()
CANCEL_CH
End Sub

Private Sub Form_Load()
Dim ws_codcia As String * 2
Dim xcuenta As Integer
pub_cadena = ""
If LK_EMP_PTO = "A" Then
 pub_cadena = "SELECT FAR_TRANSITO FROM facart WHERE Far_TRANSITO = ? AND FAR_TRANSITO<>'E'"
 Set PSFFF_LLAVE = CN.CreateQuery("", pub_cadena)
 PSFFF_LLAVE(0) = " "
 Set FFF_LLAVE = PSFFF_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
 PSFFF_LLAVE(0) = " "
 FFF_LLAVE.Requery
 If Not FFF_LLAVE.EOF Then
    Screen.MousePointer = 0
    MsgBox "!!!! Hay mercaderia en transito ...Recepcionar en Cia: " & FFF_LLAVE!far_otra_cia
    Unload PRODIA
    Exit Sub
 End If
End If
If LK_FLAG_GRIFO = "A" Then
 pub_cadena = "SELECT * FROM TABLAS WHERE TAB_CODCIA = ? AND TAB_TIPREG = ?  AND TAB_NOMCORTO = ? AND TAB_CODCLIE = ? AND TAB_CODART = ?  ORDER BY TAB_NUMTAB "
 Set PSST_LLAVE = CN.CreateQuery("", pub_cadena)
 PSST_LLAVE(0) = 0
 PSST_LLAVE(1) = 0
 PSST_LLAVE(2) = 0
 PSST_LLAVE(3) = 0
 PSST_LLAVE(4) = 0
 Set stock_llave = PSST_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
End If


'Frame3 = "Estado de Operaciones :"
'Option1(0).Caption = "Di s p o n i b l e"
'Option1(1).Caption = "C e r r a d o"
If LK_EMP_PTO = "A" Then
    xcuenta = 1
    For fila = 1 To 30
       ws_codcia = Mid(Trim(GEN!gen_ART_CIAS), xcuenta, 2)
       If Trim(ws_codcia) = "" Then Exit For
        pub_cadena = pub_cadena + " FFF_CODCIA = '" & ws_codcia & "' OR "
       xcuenta = xcuenta + 2
    Next fila
    If pub_cadena = "" Then
       MsgBox "!!!! Verificar esta Activado puntos de Ventas pero no existe declaracion de Cias!!!!", 48, Pub_Titulo
       Unload PRODIA
       Exit Sub
    End If
End If

Timer1.Enabled = True

'Dim chedef_llave As rdoResultset
'Dim PSCHE_DEF  As rdoQuery''

'pub_cadena = "SELECT  CHE_ESTADO FROM CHEQUES WHERE CHE_FECHA_COBRO = ? AND CHE_ESTADO = 'T' ORDER BY CHE_FECHA_COBRO"
'Set PS_REP01 = CN.CreateQuery("", pub_cadena)
'PS_REP01(0) = LK_CODCIA
'Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT  * FROM PARGEN WHERE PAR_CODCIA = ?  order by par_codcia"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = LK_CODCIA
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)

Exit Sub

If Nulo_Valor0(GEN!gen_cierre_todas) = 1 Or Nulo_Valor0(GEN!gen_cierre_todas) = 3 Then
   Frame2.Visible = False
   'Frame1.Left = Frame1.Left - 1500
Else
   lblFecha.Caption = "Fecha Actual : " & Format(LK_FECHA_DIA, "dddd, d mmmm yyyy")
'   Label3.Caption = Trim(par_llave!PAR_NOMBRE)
   SQ_OPER = 1
   PUB_CODCIA = LK_CODCIA
   LEER_PAR_LLAVE
   If par_llave!par_flag_cierre = 9 Then
      option1(1).Value = True
   Else
      option1(0).Value = True
   End If
End If

If Nulo_Valor0(GEN!gen_cierre_todas) = 1 Or Nulo_Valor0(GEN!gen_cierre_todas) = 3 Then
End If

End Sub


Private Sub Option1_Click(Index As Integer)
If option1(0).Visible = False Then GoTo fin

For fila = 0 To EMP.ListCount - 1
   EMP.ListIndex = fila
   If Trim(Left(EMP.Text, 2)) = "" Then
     MsgBox "Seleccione de la lista una Compañia.", 48, Pub_Titulo
     EMP.SetFocus
     Exit Sub
   End If
   PS_REP01(0) = Left(EMP.Text, 2)
   llave_rep01.Requery
   llave_rep01.Edit
   If Index = 1 Then
      llave_rep01!par_flag_cierre = 9
      option1(1).Value = True
      option1(1).ForeColor = QBColor(12)
      option1(0).ForeColor = QBColor(0)
      poperativo.Visible = False
      pbloqueado.Visible = True
   Else
      llave_rep01!par_flag_cierre = 0
      option1(0).Value = True
      option1(1).ForeColor = QBColor(0)
      option1(0).ForeColor = QBColor(2)
      poperativo.Visible = True
      pbloqueado.Visible = False
   End If
  llave_rep01.Update
Next fila
EMP.ListIndex = 0
Command1.SetFocus
Exit Sub
   PS_REP01(0) = Left(EMP.Text, 2)
   llave_rep01.Requery
   If llave_rep01!par_flag_cierre = 9 Then
      option1(1).Value = True
      option1(1).ForeColor = QBColor(12)
      option1(0).ForeColor = QBColor(0)
   Else
      option1(0).Value = True
      option1(1).ForeColor = QBColor(0)
      option1(0).ForeColor = QBColor(12)
   End If

Exit Sub
fin:
End Sub


Private Sub Timer1_Timer()
lblcierre.Visible = Not lblcierre.Visible
End Sub

Private Function cAJA(ByVal moneda As String) As Currency
Dim wsFECHA1 As String
Dim WS_MONEDA As String
Dim ww_moneda As String
Dim WS_SALDO_ING As Currency
Dim WS_SALDO_SAL As Currency


wsFECHA1 = LK_FECHA_DIA

WS_SALDO = 0
WS_MONEDA = moneda

PUB_FECHA = wsFECHA1
SQ_OPER = 1
pu_codcia = LK_CODCIA
LEER_ALL_LLAVE
If all_llave.EOF Then GoTo VAMOS

ws_conta = 0
If WS_MONEDA = "S" Then
     WS_SALDO = all_llave!ALL_IMPORTE
Else
     WS_SALDO = all_llave!ALL_IMPORTE_DOLL
End If
'MsgBox all_llave!all_codtra
all_llave.MoveNext
Do Until all_llave.EOF
   If all_llave!ALL_SIGNO_CAJA = 0 Then GoTo OTRO
   If all_llave!all_CODCIA <> LK_CODCIA Then GoTo OTRO
   If LK_EMP = "HER" And all_llave!ALL_CODTRA = 2727 And all_llave!ALL_SIGNO_CAR = 1 Then GoTo OTRO
   If all_llave!all_flag_ext = "E" Then GoTo OTRO
   If (all_llave!ALL_SIGNO_CAR = 0 And all_llave!ALL_tipmov = 0) Or (all_llave("all_codtra") = 2725 Or all_llave("all_codtra") = 5360 Or all_llave("all_codtra") = 2770 Or all_llave("all_codtra") = 2774) Then   'agregado gts para 5360 2770 y 2774
      WS_IMPORTE = all_llave!ALL_IMPORTE
   Else
      WS_IMPORTE = all_llave!ALL_IMPORTE_AMORT
   End If
   'bloqueado por mic para diro
'   If all_llave!ALL_SIGNO_CAR <> 0 And all_llave!ALL_tipmov = 0 And LK_EMP = "HER" Then
'      WS_IMPORTE = all_llave!ALL_IMPORTE
'   End If
  
   If Trim(all_llave!ALL_moneda_ccm) <> " " And val(all_llave!all_codban) <> 0 Then
      ww_moneda = all_llave!ALL_moneda_ccm
   ElseIf Trim(all_llave!ALL_MONEDA_CLI) <> " " And val(all_llave!ALL_CODCLIE) <> 0 Then
      ww_moneda = all_llave!ALL_MONEDA_CLI
   ElseIf Trim(all_llave!ALL_MONEDA_CAJA) <> " " Then
      ww_moneda = all_llave!ALL_MONEDA_CAJA
   End If
   If ww_moneda <> WS_MONEDA Then GoTo OTRO

   If all_llave!ALL_SIGNO_CAJA = 1 Then
      WS_SALDO = WS_SALDO + WS_IMPORTE
      WS_SALDO_ING = WS_SALDO_ING + WS_IMPORTE
   Else
      WS_SALDO = WS_SALDO - WS_IMPORTE
      WS_SALDO_SAL = WS_SALDO_SAL + WS_IMPORTE
   End If
   ws_conta = ws_conta + 1
   
OTRO:
  PUB_NUM_OPER = all_llave!ALL_NUMOPER
  all_llave.MoveNext
Loop
   If LK_FECHA_DIA = wsFECHA1 Then
      If WS_MONEDA = "S" Then
         PUB_TIPREG = 1000
      Else
         PUB_TIPREG = 1001
      End If
      SQ_OPER = 1
      PUB_NUMTAB = 0
      PUB_CODCIA = LK_CODCIA
      LEER_TAB_LLAVE
      If tab_llave.EOF Then
         tab_llave.AddNew
         tab_llave!TAB_NUMTAB = 0
         tab_llave!TAB_CODCIA = LK_CODCIA
         If WS_MONEDA = "S" Then
            tab_llave!TAB_TIPREG = 1000
         Else
            tab_llave!TAB_TIPREG = 1001
         End If
      Else
         tab_llave.Edit
      End If
      tab_llave!tab_NOMLARGO = PUB_NUM_OPER
      tab_llave!tab_nomcorto = Format(LK_FECHA_DIA, "dd/mm/yyyy")
      tab_llave!TAB_contable2 = WS_SALDO
      tab_llave.Update
   End If
   cAJA = WS_SALDO
VAMOS:

Screen.MousePointer = 0

Exit Function
CANCELA:
      Screen.MousePointer = 0
    Exit Function
FINTODO:
    MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
    Screen.MousePointer = 0
    Exit Function

End Function
