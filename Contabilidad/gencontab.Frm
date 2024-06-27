VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form GenContab 
   BackColor       =   &H00FAEFDA&
   Caption         =   "Pasa a Contabilidad ..."
   ClientHeight    =   2400
   ClientLeft      =   1935
   ClientTop       =   2220
   ClientWidth     =   6570
   Icon            =   "gencontab.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2400
   ScaleWidth      =   6570
   Begin VB.Frame Frame1 
      BackColor       =   &H00FAEFDA&
      Height          =   2025
      Index           =   1
      Left            =   4890
      TabIndex        =   1
      Top             =   15
      Width           =   1620
      Begin VB.CommandButton Command1 
         Caption         =   "Transferir"
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
         Left            =   210
         Picture         =   "gencontab.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   255
         Width           =   1155
      End
      Begin VB.CommandButton salir 
         Caption         =   "Salir"
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
         Left            =   225
         Picture         =   "gencontab.frx":058C
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1155
         Width           =   1155
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FAEFDA&
      Height          =   2025
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   15
      Width           =   4770
      Begin ComctlLib.ProgressBar barra 
         Height          =   240
         Left            =   165
         TabIndex        =   2
         Top             =   1470
         Visible         =   0   'False
         Width           =   4425
         _ExtentX        =   7805
         _ExtentY        =   423
         _Version        =   327682
         Appearance      =   1
      End
      Begin MSMask.MaskEdBox texto2 
         Height          =   285
         Left            =   2640
         TabIndex        =   3
         Top             =   795
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
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
      Begin MSMask.MaskEdBox texto1 
         Height          =   285
         Left            =   720
         TabIndex        =   4
         Top             =   795
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   855
         TabIndex        =   6
         Top             =   405
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2865
         TabIndex        =   5
         Top             =   405
         Width           =   690
      End
   End
   Begin Crystal.CrystalReport Reportes 
      Left            =   0
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lblbarraos 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Solution - Gestión Comercial"
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
      Height          =   330
      Left            =   0
      TabIndex        =   9
      Top             =   2070
      Width           =   6600
   End
End
Attribute VB_Name = "GenContab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xl As Object

Private Sub Command1_Click()
Dim DIA
Dim DIA1
Dim MES
Dim MES1
Dim ano
Dim ANO1
Dim II As Integer
Dim sum_debe As Currency
Dim sum_haber As Currency
Dim AWQ_TIPO_CAMBIO As Currency
Dim WS_CORRELATIVO As Long
Dim N, M, J, JJ As Integer
Dim indice As Integer
Dim tab_saldos(30) As Currency
Dim NIVEL_MAX As Integer
Dim ww_codcia As String * 2
Dim WS_TOTAL_1 As Double
Dim WS_NRO_MOV As Double
Dim ww_ult_oper As Integer
Dim WS_CUENTA As String * 12
Dim ws_importe2 As Currency
Dim longi As Integer
Dim ws_dh As String * 1
Dim WS_IMPORTE, WS_DIF As Currency
Dim ws_tot_pollos As Currency
Dim WS_TOT_IMPORTE As Currency
Dim ws_total As Currency
Dim ws_total_cantidad As Currency
Dim WS_PRECIO As Currency
Dim ws_glosa As String
Dim PSTEMP_LLAVE As rdoQuery
Dim PSTEMP2_LLAVE As rdoQuery
Dim PSTEMP3_LLAVE As rdoQuery
Dim PSTEMP4_LLAVE As rdoQuery
Dim PSallx_llave As rdoQuery

Dim temp2_llave As rdoResultset
Dim temp3_llave As rdoResultset
Dim temp4_llave As rdoResultset
Dim allx_llave As rdoResultset

Dim PSMOV_VOU As rdoQuery
Dim VOU_MOV As rdoResultset

Dim temp_llave As rdoResultset
Dim cop_llave As rdoResultset
Dim PSCOP_LLAVE As rdoQuery
Dim ctb_llave As rdoResultset
Dim PSCTB_LLAVE As rdoQuery
Dim cov_mayor As rdoResultset
Dim PSCOV_MAYOR As rdoQuery
Dim WS_CUENTA_CLAVE As String * 12
Dim WW_CUENTA(12) As String * 12
Dim WW_DH(12) As String * 1
Dim WW_CAMPO(12) As Integer
Dim fecha, fecha2, CONTADOR
Dim ws_codcia As String * 2
Dim wdiaaa As Date
Dim F1 As Integer
Dim count_asientos As Integer
Dim Def_asientos As Integer

Dim wdif_redondeo As Currency
Dim wdif_tc As Currency
Dim t_ctatc As String * 12
Dim modo_tc As String * 1
Dim ws_voucher As Integer
Dim FLAG_CAMBIO As Integer
Dim TipoMovi As Integer
'=======================
Dim iTipoDocSunat As Integer

ws_voucher = 0

Command1.Enabled = False
Pub_Titulo = " "
pub_mensaje = "Confirma pase de datos a  Contabilidad?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If
' original de alan pub_cadena = "SELECT * FROM ALLOG WHERE   ALL_CODCIA = ? AND ALL_FECHA_DIA >= ? AND ALL_FECHA_DIA <= ? AND ALL_CODTRA <> 1401 AND ALL_CODTRA <> 2401 AND ALL_CODTRA <> 9999  ORDER BY ALL_FECHA_DIA, ALL_NUMOPER "
pub_cadena = "SELECT * FROM ALLOG WHERE ALL_FLAG_EXT <> 'E' AND ALL_CODCIA = ? AND ALL_FECHA_DIA >= ? AND ALL_FECHA_DIA <= ? AND ALL_CODTRA <> 9999  ORDER BY ALL_FECHA_DIA, ALL_NUMOPER " 'AGREGUE MIC ALL_ESTADO <> 'E'
Set PSallx_llave = CN.CreateQuery("", pub_cadena)
PSallx_llave(0) = ""
PSallx_llave(1) = Date
PSallx_llave(2) = Date

Set allx_llave = PSallx_llave.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM MOVICONT WHERE MOV_CODCIA = ? AND MOV_FECHA >=? AND MOV_FECHA <= ? AND MOV_NRO_MES = ? ORDER BY MOV_FECHA , MOV_NRO_VOUCHER "
Set PSTEMP_LLAVE = CN.CreateQuery("", pub_cadena)
PSTEMP_LLAVE(0) = ""
PSTEMP_LLAVE(1) = Date
PSTEMP_LLAVE(2) = Date
PSTEMP_LLAVE(3) = 0
Set temp_llave = PSTEMP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM COMAEST WHERE COM_CODCIA = ?  AND COM_CUENTA >= ? AND COM_NIVEL= ? ORDER BY COM_CUENTA"
Set PSTEMP2_LLAVE = CN.CreateQuery("", pub_cadena)
PSTEMP2_LLAVE(0) = ""
PSTEMP2_LLAVE(1) = ""
PSTEMP2_LLAVE(2) = 0
Set temp2_llave = PSTEMP2_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM ARTI WHERE ART_CODCIA = ?  AND ART_NOMBRE  = ? "
Set PSTEMP4_LLAVE = CN.CreateQuery("", pub_cadena)
PSTEMP4_LLAVE(0) = ""
PSTEMP4_LLAVE(1) = ""
Set temp4_llave = PSTEMP4_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM COPARAM WHERE COP_CODCIA = ?"
Set PSCOP_LLAVE = CN.CreateQuery("", pub_cadena)
PSCOP_LLAVE(0) = ""
Set cop_llave = PSCOP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
modo_tc = ""

If LK_EMP_PTO = "A" Then
   ww_codcia = "00"
Else
   ww_codcia = LK_CODCIA
End If


PSCOP_LLAVE.rdoParameters(0) = ww_codcia
cop_llave.Requery


If Not cop_llave.EOF Then
    For J = 1 To 6
      If cop_llave.rdoColumns(J) <> 0 Then
         NIVEL_MAX = J
      End If
    Next J
End If

pub_cadena = "SELECT * FROM facart WHERE FAR_ESTADO <> 'E' AND FAR_CODCIA=? AND FAR_FECHA=? AND FAR_NUMOPER=?  ORDER BY FAR_CODCIA, FAR_NUMOPER, FAR_NUMSEC" 'AGREGUE FAR_ESTADO <> 'E' MIC
Set PSCTB_LLAVE = CN.CreateQuery("", pub_cadena)
PSCTB_LLAVE(0) = ""
PSCTB_LLAVE(1) = Date
PSCTB_LLAVE(2) = 0
Set ctb_llave = PSCTB_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT Max(MOV_NRO_VOUCHER) as NroVoucher FROM MOVICONT WHERE MOV_CODCIA = ? AND (MOV_FECHA >= ? AND MOV_FECHA <=?) AND MOV_NRO_MES = ? AND MOV_TIPMOV = ?"
Set PSMOV_VOU = CN.CreateQuery("", pub_cadena)
PSMOV_VOU(0) = 0
PSMOV_VOU(1) = LK_FECHA_DIA
PSMOV_VOU(2) = LK_FECHA_DIA
PSMOV_VOU(3) = 0
PSMOV_VOU(4) = 0
Set VOU_MOV = PSMOV_VOU.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


'If LK_FECHA_COP2 = cop_llave!COP_FECHA_gencontab Then
'   pub_mensaje = " !!! ADVERTENCIA !!! ... Ojo Falta cerrar el dia de contabilidad. Aun esta con fecha de :" & LK_FECHA_COP1
'   Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
'   GoTo fin
'End If
If Not IsDate(texto1.Text) Then
   MsgBox "Verificar Fechas ", 48, Pub_Titulo
   Exit Sub
End If
If Not IsDate(texto2.Text) Then
   MsgBox "Verificar Fechas ", 48, Pub_Titulo
   Exit Sub
End If

LK_FECHA_COP1 = texto1.Text
LK_FECHA_COP2 = texto2.Text
   
barra.Visible = True
wdif_tc = 0
wdif_redondeo = 0

fecha = Format(LK_FECHA_COP1, "dd/mm/yyyy")
fecha2 = Format(LK_FECHA_COP2, "dd/mm/yyyy")


pub_cadena = "DELETE MOVICONT WHERE MOV_CODCIA = '" & LK_CODCIA & "' AND (MOV_FECHA >=  '" & fecha & "' AND MOV_FECHA <=  '" & fecha2 & "') AND  MOV_NRO_MES = " & Format(texto1.Text, "mm") 'quitado por mic MOV_TIPMOV = 3 AND
CN.Execute pub_cadena, rdExecDirect

PSTEMP_LLAVE(0) = ww_codcia
PSTEMP_LLAVE(1) = texto1.Text 'LK_FECHA_COP1
PSTEMP_LLAVE(2) = texto2.Text ' LK_FECHA_COP2
PSTEMP_LLAVE(3) = Format(texto2.Text, "mm")
temp_llave.Requery

If temp_llave.EOF Then
  ws_voucher = 0
Else
  temp_llave.MoveLast
  ws_voucher = temp_llave!MOV_NRO_VOUCHER
End If

ws_total = 0
Command1.Enabled = False

PSallx_llave(0) = LK_CODCIA
PSallx_llave(1) = texto1.Text  'LK_FECHA_COP1
PSallx_llave(2) = texto2.Text ' LK_FECHA_COP2

allx_llave.Requery
If allx_llave.EOF Then
   barra.Visible = False
   Command1.Enabled = True
   MsgBox "No hay registros ..."
   Exit Sub
End If
'GoSub WEXCEL
'xl.Cells(1, 1) = Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption))
'xl.Cells(2, 1) = "' DEL " & Format(LK_FECHA_COP1, "dd/mm/yyyy") & " AL " & Format(LK_FECHA_COP2, "dd/mm/yyyy")
F1 = 7

CONTADOR = 0
barra.Min = 0
barra.Max = allx_llave.RowCount
wdiaaa = #5/2/2000#
WS_NRO_MOV = 0
FLAG_CAMBIO = Nulo_Valor0(allx_llave!ALL_NUMOPER2)

Do Until allx_llave.EOF
    WS_NRO_MOV = 0
    If allx_llave!ALL_codtra = 9999 Then
       FLAG_CAMBIO = Nulo_Valor0(allx_llave!ALL_NUMOPER2)
       GoTo OTRO
    End If
    If allx_llave!ALL_NUMOPER2 <> FLAG_CAMBIO Then
      FLAG_CAMBIO = Nulo_Valor0(allx_llave!ALL_NUMOPER2)
    End If
    CONTADOR = CONTADOR + 1
    barra.Value = CONTADOR
    DoEvents
    ww_ult_oper = allx_llave!ALL_NUMOPER
    If allx_llave!all_flag_ext = "E" Then
       GoTo OTRO
    End If
   
    ' POR QUITAR TEMPORALMENTE
    If allx_llave!ALL_codtra = 5310 And LK_EMP = "HER" Then
        If Format(allx_llave!ALL_FECHA_DIA, "dd/mm/yyyy") = Format(wdiaaa, "dd/mm/yyyy") Then
          GoTo OTRO
        End If
    End If
    'If allx_llave!ALL_CODTRA = 2748 Then GoTo OTRO
    'If allx_llave!ALL_codtra = 1401 Then GoTo OTRO
    'If allx_llave!ALL_codtra = 2412 Then GoTo OTRO
    'If allx_llave!ALL_codtra = 2410 Then GoTo OTRO
    
    
    If allx_llave!ALL_codtra = 1401 Then
        TipoMovi = 1 ' Registro de compras
    ElseIf allx_llave!ALL_codtra = 2401 Then
        TipoMovi = 2 ' Registro de ventas
    ElseIf allx_llave!ALL_codtra = 2720 Or allx_llave!ALL_codtra = 2725 Or allx_llave!ALL_codtra = 5310 Then
        TipoMovi = 3 ' Ingresos de fondos
    ElseIf allx_llave!ALL_codtra = 2738 Or allx_llave!ALL_codtra = 5360 Or allx_llave!ALL_codtra = 2735 Then
        TipoMovi = 4 ' Egresos de fondos
    End If
    
    If allx_llave!ALL_codtra = 2409 Then GoTo OTRO
    If allx_llave!ALL_codtra = 2407 Then GoTo OTRO
    If allx_llave!ALL_codtra = 2211 Then GoTo OTRO
    If allx_llave!ALL_codtra = 2103 Then GoTo OTRO
    If allx_llave!ALL_codtra = 2741 Then GoTo OTRO
    If allx_llave!ALL_codtra = 2580 Then GoTo OTRO
    If allx_llave!ALL_codtra = 1122 Then GoTo OTRO
    If allx_llave!ALL_codtra = 5715 Then GoTo OTRO
    If allx_llave!ALL_codtra = 2403 Then GoTo OTRO
    If allx_llave!ALL_codtra = 5714 Then
       If allx_llave!ALL_CODBAN = 1 Then GoTo OTRO
    End If
    If allx_llave!ALL_codtra = 1455 Then
       If allx_llave!ALL_SIGNO_CAR <> -1 Then GoTo OTRO
    End If
    If allx_llave!ALL_codtra = 2401 Then
      If allx_llave!ALL_tipdoc = "CH" Then GoTo OTRO
      'If allx_llave!ALL_SIGNO_CAR <> 0 Then GoTo OTRO
    End If
   
    t_ctatc = "10101"
    SQ_OPER = 1
    PUB_CODTRA = allx_llave!ALL_codtra
    PUB_SECUENCIA = allx_llave!all_SECUENCIA
    PUB_CODCIA = ww_codcia
    LEER_CNT_LLAVE
    If cnt_llave.EOF Then
       MsgBox "Definir Sub Transaccion Contable  en : " & "TRANS: " & PUB_CODTRA & " Secuencia: " & PUB_SECUENCIA
       GoTo OTRO
    End If
    Def_asientos = 0
    For II = 3 To 36 Step 3
       If Trim(cnt_llave(II)) <> "" Then
           Def_asientos = Def_asientos + 1
       End If
    Next II
   
    count_asientos = 0
    sum_debe = 0
    sum_haber = 0

    ' PARA EL TIPO DE CAMBIO
      AWQ_TIPO_CAMBIO = 1
      Print allx_llave!ALL_codtra
      If allx_llave!ALL_moneda_CAJA = "D" Then
         'PUB_CAL_INI = allx_llave!ALL_fecha_ANT ' allx_llave!ALL_fecha_dia
         'PUB_CAL_FIN = allx_llave!ALL_fecha_ANT ' allx_llave!ALL_fecha_dia
         'PUB_CODCIA = LK_CODCIA
         'LEER_CAL_LLAVE 1
         'If Nulo_Valor0(cal_llave!cal_tipo_cambio) = 0 Then
           AWQ_TIPO_CAMBIO = 1
         'Else
         '  AWQ_TIPO_CAMBIO = cal_llave!cal_tipo_cambio
         'End If
         'If AWQ_TIPO_CAMBIO <= 0 Then
         '    MsgBox "Definir Tipo de Cambios para el Periodo Actual. Dia : " & allx_llave!ALL_FECHA_DIA & " (en el Calendario del Sistema)", 48, Pub_Titulo
         '    MsgBox "Falta Tipo de Cambio.... "
         '    GoTo fin
         'End If
      End If
   
      N = 1
      M = 3
      Do Until N > 11
             WW_CUENTA(N) = Nulo_Valors(cnt_llave(M))
             If allx_llave!ALL_codtra = 5318 Then 'AGREGE ACV
                If allx_llave!ALL_SIGNO_CCM = 1 Then
                  WW_DH(N) = "D"
                ElseIf allx_llave!ALL_SIGNO_CCM = -1 Then
                  WW_DH(N) = "H"
                End If
             Else
               WW_DH(N) = Nulo_Valors(cnt_llave(M + 1))
             End If
             WW_CAMPO(N) = Val(Nulo_Valor0(cnt_llave(M + 2)))
             
             If (WW_CUENTA(N) = "CLIENTES    " Or WW_CUENTA(N) = "CLIENTES2   " Or Right(Trim(WW_CUENTA(N)), 2) = "XX") And allx_llave!ALL_SIGNO_CAR <> 0 Then
                SQ_OPER = 1
                pu_cp = allx_llave!ALL_cp
                pu_codclie = allx_llave!ALL_CODCLIE
                pu_codcia = LK_CODCIA
                LEER_CLI_LLAVE
                If cli_llave.EOF Then
                   MsgBox "OJO REVISAR CODIGO DE CLIENTES..." & PUB_CODCLIE
                Else
                   If WW_CUENTA(N) = "CLIENTES    " Then
                      WW_CUENTA(N) = Nulo_Valors(cli_llave!CLI_CUENTA_CONTAB)
                   ElseIf Trim(WW_CUENTA(N)) = "CLIENTES2" Then
                      WW_CUENTA(N) = Nulo_Valors(cli_llave!CLI_CUENTA_CONTAB2)
                   End If
                   If Trim(WW_CUENTA(N)) <> "" And Trim(allx_llave!ALL_tipdoc) = "LE" Then
                       WW_CUENTA(N) = Mid(Trim(WW_CUENTA(N)), 1, 2) + "3" + Mid(Trim(WW_CUENTA(N)), 4, Len(Trim(WW_CUENTA(N))))
                   End If
                   t_ctatc = WW_CUENTA(N)
                   If Trim(Nulo_Valors(cli_llave!CLI_CUENTA_CONTAB)) = "" Then MsgBox "Falta cuenta contable (12)..." & pu_codclie
                End If
             End If
        
             If WW_CUENTA(N) = "BANCOS      " And allx_llave!ALL_SIGNO_CCM <> 0 Then
                SQ_OPER = 1
                pu_cp = allx_llave!ALL_cp
                PUB_CODBAN = allx_llave!ALL_CODBAN
                pu_codcia = LK_CODCIA
                LEER_CCM_LLAVE
                If ccm_llave.EOF Then
                   MsgBox "OJO REVISAR CODIGO DE BANCO..." & PUB_CODBAN
                Else
                   WW_CUENTA(N) = Nulo_Valors(ccm_llave!CCM_CUENTA_CONTAB)
                   If Trim(WW_CUENTA(N)) = "" Then
                     MsgBox "Definir Cuenta Contable a : " & ccm_llave!CCM_CODBAN & " " & ccm_llave!CCM_nombre
                   End If
                End If
             End If
             If Right(Trim(WW_CUENTA(N)), 2) = "XX" Then
                longi = Len(Trim(WW_CUENTA(N)))
                WW_CUENTA(N) = Left(WW_CUENTA(N), longi - 2) & Right(Trim(cli_llave!CLI_CUENTA_CONTAB), 2)
                If Trim(cli_llave!CLI_CUENTA_CONTAB) = "" Then WW_CUENTA(N) = ""
             End If
             N = N + 1
             M = M + 3
       Loop
  
    
  ' AQUI FALTA EL PRIMERO ES CERO PERO EL SEGUNDO ..SI TIENE DATOS
    If WW_CUENTA(1) <> "FACART      " Then
       N = 1
      GoTo automatico
    End If
   
    PSCTB_LLAVE.rdoParameters(0) = LK_CODCIA
    PSCTB_LLAVE.rdoParameters(1) = allx_llave!ALL_FECHA_DIA
    PSCTB_LLAVE.rdoParameters(2) = allx_llave!ALL_NUMOPER
    ctb_llave.Requery
    ws_total = 0
    ws_total_cantidad = 0
    Do Until ctb_llave.EOF
        SQ_OPER = 1
        PUB_KEY = ctb_llave!far_codart
        pu_codcia = LK_CODCIA
        LEER_ART_LLAVE
        If art_LLAVE.EOF Then
           MsgBox "ERROR...GRAVE EN ARTI"
           End
        End If
      
         ws_dh = WW_DH(1)
         WS_PRECIO = ctb_llave!FAR_COSPRO
         WS_CUENTA = Nulo_Valors(art_LLAVE!art_cuenta_contab)
         WS_IMPORTE = WS_PRECIO * ctb_llave!FAR_CANTIDAD
      
         SQ_OPER = 1
         PUB_CUENTA = WS_CUENTA
         PUB_CODCIA = ww_codcia
         LEER_COM_LLAVE
         If com_llave.EOF Then
            MsgBox "OJO..NO EXISTE CUENTA CONTABLE..REVISAR ARTICULO:" & PUB_KEY & "  " & art_LLAVE!ART_NOMBRE
            GoTo fin
         End If
            
        WS_IMPORTE = redondea(WS_IMPORTE)
        ws_total = WS_IMPORTE + ws_total
        ws_total_cantidad = ctb_llave!FAR_CANTIDAD + ws_total_cantidad
        GoSub REGISTRA
        ctb_llave.MoveNext
      Loop
      
      
      If allx_llave!ALL_codtra = 1401 Then
         WS_NETO = Val(allx_llave!all_neto) + Val(allx_llave!ALL_flete) - ws_total
         If WS_NETO <> 0 Then
            WS_CUENTA = Nulo_Valors(cop_llave!cop_cta_red_FAVOR)
            WS_IMPORTE = WS_NETO
            ws_total = ws_total + WS_NETO
            GoSub REGISTRA
         End If
     End If
     
     

N = 2
automatico:
'*****************
    PSMOV_VOU.rdoParameters(0) = LK_CODCIA
    PSMOV_VOU.rdoParameters(1) = fecha
    PSMOV_VOU.rdoParameters(2) = fecha2
    PSMOV_VOU.rdoParameters(3) = LK_NRO_MES
    PSMOV_VOU.rdoParameters(4) = TipoMovi
    VOU_MOV.Requery
    If VOU_MOV.EOF Then
       ws_voucher = 0
    Else
       ws_voucher = IIf(IsNull(VOU_MOV!NroVoucher), 0, VOU_MOV!NroVoucher)
    End If
    ws_voucher = ws_voucher + 1
'*****************
      Do Until N > 11
        If Left(WW_CUENTA(N), 1) = "P" Then GoTo automatico3
        If Val(WW_CUENTA(N)) <> 0 Then
            WS_CUENTA = WW_CUENTA(N)
            ws_dh = WW_DH(N)
            indice = WW_CAMPO(N)
            If indice = 0 Then
               WS_IMPORTE = ws_total
            Else
                If par_llave!PAR_MONEDA_CON = "S" Then
                    'If allx_llave!ALL_MONEDA_CLI = "S" Then
                    GoSub soles
                Else
                    GoSub dollares
                End If
            End If
            GoSub REGISTRA
        End If
        N = N + 1
     Loop
      ' AGREGE ACV
     wdif_tc = 0
     wdif_redondeo = 0
     If allx_llave!ALL_SIGNO_CAR <> 0 And indice <> 8 Then
       If allx_llave!ALL_moneda_CLI = "D" Then
          If allx_llave!ALL_SIGNO_CAJA <> 0 Then
            'PUB_CAL_INI = allx_llave!ALL_fecha_ANT ' allx_llave!ALL_fecha_dia
            'PUB_CAL_FIN = allx_llave!ALL_fecha_ANT ' allx_llave!ALL_fecha_dia
            'PUB_CODCIA = LK_CODCIA
            'LEER_CAL_LLAVE
            'If Nulo_Valor0(cal_llave!cal_tipo_cambio) = 0 Then
            '   MsgBox "Definir Tipo de Cambios para el Periodo Actual. Dia : " & allx_llave!ALL_FECHA_DIA & " (en el Calendario del Sistema)", 48, Pub_Titulo
            '  GoTo fin
            'End If
            If allx_llave!ALL_moneda_CAJA = "S" Then
              WS_IMPORTE = redondea(allx_llave(indice) * allx_llave!ALL_tipo_cambio)
              wdif_tc = Val(Format(allx_llave!all_Importe_AMORT * cal_llave!cal_tipo_cambio, "0.00")) - Val(Format(allx_llave!all_Importe_AMORT * allx_llave!ALL_tipo_cambio, "0.00"))
              wdif_redondeo = (WS_IMPORTE - allx_llave!ALL_IMPORTE)
            ElseIf allx_llave!ALL_moneda_CAJA = "D" Then
              WS_IMPORTE = redondea(allx_llave(indice) * allx_llave!ALL_tipo_cambio)
              wdif_tc = Val(Format(allx_llave!all_Importe_AMORT * cal_llave!cal_tipo_cambio, "0.00")) - Val(Format(allx_llave!all_Importe_AMORT * allx_llave!ALL_tipo_cambio, "0.00"))
              wdif_redondeo = 0
            End If
          Else
           If allx_llave!ALL_moneda_CAJA = "S" Then
              WS_IMPORTE = redondea(allx_llave(indice)) ' * allx_llave!ALL_tipo_cambio)
              'wdif_tc = Val(Format(allx_llave!all_Importe_AMORT * allx_llave!ALL_tipo_cambio, "0.00")) - Val(Format(allx_llave!all_Importe_AMORT * cal_llave!cal_tipo_cambio, "0.00"))
              wdif_redondeo = allx_llave!ALL_IMPORTE - (Val(Format(allx_llave!all_Importe_AMORT * allx_llave!ALL_tipo_cambio, "0.00")))
            ElseIf allx_llave!ALL_moneda_CAJA = "D" Then
              'WS_IMPORTE = redondea(allx_llave(indice) * allx_llave!ALL_tipo_cambio)
              'wdif_tc = Val(Format(allx_llave!all_Importe_AMORT * cal_llave!cal_tipo_cambio, "0.00")) - Val(Format(allx_llave!all_Importe_AMORT * allx_llave!ALL_tipo_cambio, "0.00"))
              'wdif_redondeo = 0
            End If
          End If
       End If
    End If
     
     GoTo OTRO
     
automatico3:
'*****************
    PSMOV_VOU.rdoParameters(0) = LK_CODCIA
    PSMOV_VOU.rdoParameters(1) = fecha
    PSMOV_VOU.rdoParameters(2) = fecha2
    PSMOV_VOU.rdoParameters(3) = LK_NRO_MES
    PSMOV_VOU.rdoParameters(4) = TipoMovi
    VOU_MOV.Requery
    If VOU_MOV.EOF Then
       ws_voucher = 0
    Else
       ws_voucher = IIf(IsNull(VOU_MOV!NroVoucher), 0, VOU_MOV!NroVoucher)
    End If
    ws_voucher = ws_voucher + 1
'*****************
    PSTEMP2_LLAVE.rdoParameters(0) = LK_CODCIA
    PSTEMP2_LLAVE.rdoParameters(1) = Trim(cop_llave!COP_CUENTA_PRORRATEO)
    PSTEMP2_LLAVE.rdoParameters(2) = NIVEL_MAX
    WS_TOT_IMPORTE = 0
    J = 0
    ws_dh = WW_DH(2)
    indice = WW_CAMPO(2)
    If indice = 0 Then
       ws_importe2 = ws_total
    Else
       ws_importe2 = all_llave(indice)
    End If
    
    temp2_llave.Requery
    Do Until temp2_llave.EOF
       If Trim(cop_llave!COP_CUENTA_PRORRATEO) = Left(temp2_llave!com_cuenta, Len(Trim(cop_llave!COP_CUENTA_PRORRATEO))) Then
            If Mid(Trim(WW_CUENTA(2)), 2, 10) = Right(Trim(temp2_llave!com_cuenta), Len(Mid(Trim(WW_CUENTA(2)), 2, 10))) Then
                J = J + 1
                If J = JJ Then
                   WS_IMPORTE = ws_importe2 - WS_TOT_IMPORTE
                Else
                   WS_IMPORTE = redondea(ws_importe2 * tab_saldos(J) / ws_tot_pollos)
                   WS_TOT_IMPORTE = WS_TOT_IMPORTE + WS_IMPORTE
                End If
                WS_CUENTA = temp2_llave!com_cuenta
                GoSub REGISTRA
            End If
       End If
       temp2_llave.MoveNext
    Loop
    
    
OTRO:
   
   allx_llave.MoveNext
      
   Loop
cop_llave.Edit
cop_llave!COP_ULT_OPER = ww_ult_oper
cop_llave!COP_FECHA_GENCONTAB = LK_FECHA_DIA
cop_llave!COP_FLAG_PASE = "A"
cop_llave!cop_FLAG_CAJA = " "
cop_llave!COP_FLAG_MAYORIZACION = " "
cop_llave!cop_FLAG_DES = " "

cop_llave.Update


'xl.DisplayAlerts = False
'xl.Worksheets(1).Protect PUB_CLAVE
'xl.Application.Visible = True
'Set xl = Nothing
Screen.MousePointer = 0
MsgBox "Proceso terminado....ok ", 48, Pub_Titulo

'{MOVICONT.MOV_CODCIA} = "10" and
'{MOVICONT.MOV_FECHA} = Date (2000, 09, 06)
barra.Visible = False
Command1.Enabled = True
Reportes.Connect = PUB_ODBC
Reportes.WindowTitle = "Reporte :  Voucher"
Reportes.Destination = crptToWindow
Reportes.WindowLeft = 2
Reportes.WindowTop = 70
Reportes.WindowWidth = 635
Reportes.WindowHeight = 390
Reportes.Formulas(0) = "CIA=  '" & Mid(MDIForm1.TXTCIA.Caption, 4 + 8, Len(MDIForm1.TXTCIA.Caption)) & "'"
Reportes.ReportFileName = PUB_RUTA_OTRO & "vou.rpt"

DIA = Day(LK_FECHA_COP1)
MES = Month(LK_FECHA_COP1)
ano = Year(LK_FECHA_COP1)
DIA1 = Day(LK_FECHA_COP2)
MES1 = Month(LK_FECHA_COP2)
ANO1 = Year(LK_FECHA_COP2)
pub_cadena = "{MOVICONT.MOV_CODCIA} = '" & LK_CODCIA & "' AND {MOVICONT.MOV_FECHA} >= Date ( " & ano & "," & MES & "," & DIA & ") AND {MOVICONT.MOV_FECHA} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
Reportes.SelectionFormula = pub_cadena
On Error GoTo MIRA
Reportes.Action = 1
On Error GoTo 0

GoTo fin
REGISTRA:
    
'      If allx_llave!ALL_NUMOPER2 = 53 Then Stop
        If WS_IMPORTE <> 0 And Val(WS_CUENTA) <> 0 Then
            temp_llave.AddNew
            count_asientos = count_asientos + 1
            If ws_dh = "D" Then
               sum_debe = sum_debe + WS_IMPORTE
            ElseIf ws_dh = "H" Then
               sum_haber = sum_haber + WS_IMPORTE
            End If
            WS_NRO_MOV = WS_NRO_MOV + 1
            temp_llave!MOV_NRO_MOV = WS_NRO_MOV
            If LK_EMP_PTO = "A" Then
               temp_llave!MOV_CODCIA = "00"
            Else
               temp_llave!MOV_CODCIA = allx_llave!all_codCIA
            End If
            temp_llave!MOV_nro_MES = Format(texto1.Text, "mm")
            temp_llave!MOV_CODCTA = WS_CUENTA
            temp_llave!MOV_DH = ws_dh
            'AGRGEGADO POR MIC SOLO PARA VENTAS IGV
            If WS_CUENTA = 701001 Or WS_CUENTA = 601001 Then
                WS_IMPORTE = WS_IMPORTE / (1 + LK_IGV / 100)
            ElseIf WS_CUENTA = 401001 Then
                WS_IMPORTE = WS_IMPORTE - (WS_IMPORTE / (1 + LK_IGV / 100))
            End If
            WS_IMPORTE = Format(WS_IMPORTE, "#0.00")
            temp_llave!MOV_IMPORTE = WS_IMPORTE
            temp_llave!MOV_FECHA = allx_llave!ALL_FECHA_DIA
            temp_llave!MOV_fecha_EMI = allx_llave!ALL_FECHA_SUNAT
            temp_llave!MOV_PERIODO = Format(allx_llave!ALL_FECHA_DIA, "yyyy")
            temp_llave!MOV_GLOSA = allx_llave!ALL_autocon
            temp_llave!MOV_NRO_VOUCHER = ws_voucher
            temp_llave!MOV_TIPMOV = TipoMovi
            temp_llave!MOV_GLOSA = " "
            If allx_llave!ALL_SIGNO_CAJA = 0 Then
              temp_llave!MOV_MONEDA = allx_llave!ALL_moneda_CLI
            Else
              temp_llave!MOV_MONEDA = allx_llave!ALL_moneda_CAJA
            End If
            temp_llave!MOV_SUNAT = allx_llave!ALL_CODSUNAT
            temp_llave!MOV_FBG = allx_llave!ALL_tipdoc
            '!!!!!!!!!!!!!! V E R I F I C A R !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            'TODO ESTO SE BLOQUEO POR PRUEBA
'            If Val(allx_llave!ALL_CODCLIE) <> 0 And allx_llave!ALL_SIGNO_CCM = 0 And allx_llave!ALL_NUMFAC_C <> 0 Then 'agregue and allx_llave!ALL_NUMFAC_C <> 0 para pasar mic
'              temp_llave!MOV_serie = allx_llave!ALL_NUMSER_C
'              temp_llave!MOV_numfac = allx_llave!ALL_NUMFAC_C
'            End If
'            If allx_llave!ALL_CHENUM <> 0 And allx_llave!ALL_SIGNO_CAR = 0 Then
'              temp_llave!MOV_FBG = "CH"
'              temp_llave!MOV_serie = Val(allx_llave!ALL_CHESER)
'              temp_llave!MOV_numfac = allx_llave!ALL_CHENUM
'            End If
'            If allx_llave!ALL_NUMFAC_C = 0 Then  ' quite allx_llave!ALL_SIGNO_CAJA agregue allx_llave!ALL_NUMFAC_C para pasar
'              temp_llave!MOV_FBG = allx_llave!ALL_FBG
'              temp_llave!MOV_serie = allx_llave!ALL_NUMSER
'              temp_llave!MOV_numfac = allx_llave!ALL_NUMFAC
'            End If
            'aGREGUE PO LA PARTE BLOQUEADA
            If allx_llave!ALL_CHENUM <> 0 And allx_llave!ALL_SIGNO_CAR = 0 Then
              temp_llave!MOV_FBG = "CH"
              temp_llave!MOV_serie = Val(allx_llave!ALL_CHESER)
              temp_llave!MOV_numfac = allx_llave!ALL_CHENUM
            Else
                temp_llave!MOV_serie_c = allx_llave!ALL_NUMSER_C
                temp_llave!MOV_numfac_c = allx_llave!ALL_NUMFAC_C
                temp_llave!MOV_serie = allx_llave!ALL_NUMSER
                temp_llave!MOV_numfac = allx_llave!ALL_NUMFAC
            End If
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            temp_llave!MOV_codclie = allx_llave!ALL_CODCLIE
            temp_llave!MOV_CP = allx_llave!ALL_cp
            
            temp_llave!MOV_MARCA = "X"
            temp_llave!MOV_DETALLE = allx_llave!ALL_autocon
            temp_llave!MOV_FBG_C = " "
            temp_llave!MOV_PLANTILLA = 0
            temp_llave!MOV_FLAG_TC = ""
            temp_llave!MOV_TIPO_CAMBIO = allx_llave!ALL_tipo_cambio
            temp_llave!MOV_FLAG_DES = " "
            temp_llave!MOV_CODUSU = LK_CODUSU
            temp_llave!mov_vou2 = allx_llave!ALL_codtra & allx_llave!ALL_NUMOPER
        
           temp_llave.Update
      End If
      
      Return
      
soles:
'   If allx_llave!ALL_CODTRA = 2735 Then Stop
      WS_IMPORTE = allx_llave(indice)
      If allx_llave!ALL_codtra = 1455 Then GoTo sal
      If allx_llave!ALL_SIGNO_CAJA <> 0 And indice = 8 Then
         If allx_llave!ALL_moneda_CAJA = "D" Then WS_IMPORTE = redondea(allx_llave(indice) * allx_llave!ALL_tipo_cambio)
      End If
      If allx_llave!ALL_SIGNO_CAJA <> 0 And indice = 7 And allx_llave!ALL_codtra = 2401 Then
         If allx_llave!ALL_moneda_CLI = "D" Then WS_IMPORTE = redondea(allx_llave(indice) * allx_llave!ALL_tipo_cambio)
      End If
      If allx_llave!ALL_SIGNO_CAR <> 0 And indice = 7 Then
         If allx_llave!ALL_moneda_CLI = "D" Then
           If allx_llave!ALL_SIGNO_CAJA <> 0 Or allx_llave!ALL_codtra = 1455 Then
             'PUB_CAL_INI = allx_llave!ALL_fecha_ANT ' allx_llave!ALL_fecha_dia
             'PUB_CAL_FIN = allx_llave!ALL_fecha_ANT ' allx_llave!ALL_fecha_dia
             'PUB_CODCIA = LK_CODCIA
             'LEER_CAL_LLAVE
             'If Nulo_Valor0(cal_llave!cal_tipo_cambio) = 0 Then
             '   MsgBox "Definir Tipo de Cambios para el Periodo Actual. Dia : " & allx_llave!ALL_FECHA_DIA & " (en el Calendario del Sistema)", 48, Pub_Titulo
             '  GoTo fin
             'End If
            ' ACV
              WS_IMPORTE = redondea(allx_llave(indice)) '* cal_llave!cal_tipo_cambio)  ' allx_llave!ALL_tipo_cambio)
            Else
              If allx_llave!ALL_moneda_CLI <> "S" Then
              WS_IMPORTE = redondea(allx_llave(indice)) ' * allx_llave!ALL_tipo_cambio)  ' allx_llave!ALL_tipo_cambio)
              Print WS_CUENTA & ws_dh
              
              'WS_IMPORTE = redondea(allx_llave(indice) * allx_llave!ALL_tipo_cambio)  ' allx_llave!ALL_tipo_cambio)
              End If
            End If
         End If
      End If
      If allx_llave!ALL_codtra = 2735 Then
        If allx_llave!ALL_SIGNO_CCM <> 0 And allx_llave!ALL_SIGNO_CAR = 0 And ws_dh = "D" Then
          WS_IMPORTE = 0
        End If
      End If
 '     Print WS_CUENTA
      If allx_llave!ALL_SIGNO_CCM <> 0 And indice = 8 Then
         If allx_llave!ALL_moneda_CCM = "D" Then WS_IMPORTE = redondea(allx_llave(indice)) ' * allx_llave!ALL_tipo_cambio)
      End If
'      MsgBox WS_IMPORTE
sal:
Return
      
dollares:
      WS_IMPORTE = allx_llave(indice)

      If allx_llave!ALL_SIGNO_CAR <> 0 Then
         If allx_llave!ALL_moneda_CLI = "S" Then
            WS_IMPORTE = allx_llave(indice) / allx_llave!ALL_tipo_cambio
         Else
            WS_IMPORTE = allx_llave(indice)
         End If
      End If
      If allx_llave!ALL_SIGNO_CCM <> 0 Then
         If allx_llave!ALL_moneda_CCM = "S" Then
            WS_IMPORTE = allx_llave(indice) / 1 ' allx_llave!all_tipo_cambio ' ACV
         Else
            WS_IMPORTE = allx_llave(indice)
         End If
      End If
         
Return

WEXCEL:
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  xl.Workbooks.Open CONS_ADMIN & "CONTABILIDAD\RESPASE.xls", 0, True, 4, PUB_CLAVE, PUB_CLAVE

Return
Exit Sub
MIRA:
Resume Next

fin:
'Unload GenContab
End Sub

Private Sub Form_Load()
    CenterMe GenContab
    If cop_llave.EOF Then
      MsgBox "Definir Parametros en Contabilidad... ", 48, Pub_Titulo
      Exit Sub
    End If
    texto1.Text = Format(LK_FECHA_COP1, "dd/mm/yyyy")
    texto1.Mask = "##/##/####"
    texto2.Text = Format(LK_FECHA_COP2, "dd/mm/yyyy")
    texto2.Mask = "##/##/####"
End Sub

Private Sub salir_Click()
    Unload GenContab
End Sub


