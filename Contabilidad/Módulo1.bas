Attribute VB_Name = "Módulo1"
Public PUB_DSN As String
Public CN As rdoConnection
Public CONn As String
Public EN As rdoEnvironment
'Dim espaciot As Workspace
Public ODBCRUTA As String
Public PSX As rdoQuery
Public PSTRA As rdoQuery
Public PSVEN As rdoQuery
Public PSCOS_LLAVE As rdoQuery
Public cos_llave As rdoResultset
Public PSFAR_LLAVE2 As rdoQuery
Public PSFAR_CODCLIE As rdoQuery

Public PSCNT_LLAVE As rdoQuery
Public PSCNT_MAYOR As rdoQuery
Public PSPRE_MAYOR As rdoQuery
Public PSPRE_LLAVE As rdoQuery

Public PSCAL_LLAVE As rdoQuery
Public PSFER_LLAVE As rdoQuery
Public PSUSU_LLAVE As rdoQuery

Public cal_llave As rdoResultset
Public fer_llave As rdoResultset
Public PUB_CAL_INI As Date
Public PUB_CAL_FIN As Date
Public PUB_CAL_ANO As Integer
Public PSCOV_VOUCHER  As rdoQuery

Public PSCOP_LLAVE As rdoQuery
Public cop_llave  As rdoResultset

Public PSPAR_MULTI As rdoQuery
Public par_multi   As rdoResultset

Public PSFFF_LLAVE As rdoQuery
Public FFF_LLAVE As rdoResultset
Public cov_voucher  As rdoResultset

Public calen As rdoResultset

Public PSPER_BUSCA As rdoQuery
Public per_busca As rdoResultset


Public PSZON_LLAVE As rdoQuery
Public far_llave2 As rdoResultset
Public ven As rdoResultset
Public cont As rdoResultset
Public tra As rdoResultset
Public GEN As rdoResultset
Public arm As rdoResultset
Public ccm As rdoResultset
Public par As rdoResultset
Public car As rdoResultset
Public PRE As rdoResultset
Public CLI As rdoResultset
Public all As rdoResultset
Public usu As rdoResultset
Public aut As rdoResultset
Public che As rdoResultset
Public gru As rdoResultset
Public lis_tra As rdoResultset
Public aut_llave As rdoResultset
Public aut_menor As rdoResultset
Public cli_llave As rdoResultset
Public cli_mayor As rdoResultset
Public cli_mayor2 As rdoResultset
Public ven_llave As rdoResultset
Public tra_llave As rdoResultset
Public tra_menu As rdoResultset
Public art_LLAVE As rdoResultset
Public art_LLAVE10 As rdoResultset
Public art_mayor As rdoResultset
Public arm_llave As rdoResultset
Public ccm_llave As rdoResultset
Public ccm_mayor As rdoResultset
Public ccm_mayor2 As rdoResultset
Public far_llave As rdoResultset
Public far_menor As rdoResultset
Public far_menor2 As rdoResultset
Public far_menor3 As rdoResultset
Public proc_mayor As rdoResultset
Public cnt_mayor As rdoResultset
Public pre_mayor As rdoResultset
Public pre_llave As rdoResultset
Public PSTAB_MENOR As rdoQuery
Public tab_menor As rdoResultset


Public far_codcli As rdoResultset
Public usu_llave As rdoResultset

Public com_llave As rdoResultset
Public com_mayor As rdoResultset
Public PSCOM_LLAVE As rdoQuery
Public PSCOM_MAYOR As rdoQuery
Public PUB_CUENTA As String
Public cnt_llave As rdoResultset
Public con_llave As rdoResultset
Public par_llave As rdoResultset
Public car_llave As rdoResultset
Public caa_histo As rdoResultset
Public car_mayor As rdoResultset
Public car_menor As rdoResultset
Public car_far As rdoResultset

Public pro_llave As rdoResultset
Public all_llave As rdoResultset
Public all_menor As rdoResultset
Public Gen_llave As rdoResultset
Public tab_llave As rdoResultset
Public tab_mayor As rdoResultset
Public SUT_MAYOR As rdoResultset
Public SUT_LLAVE As rdoResultset
Public cov_llave As rdoResultset
Public cov_mayor As rdoResultset
Public che_menor As rdoResultset
Public che_oper As rdoResultset
Public che_repo As rdoResultset
Public che_llave As rdoResultset
Public Che_mayor As rdoResultset
Public che_movi As rdoResultset
Public caa_LLAVE As rdoResultset

Public zon_llave As rdoResultset
Public X As rdoResultset
Public SQ_OPER As Integer
Public sq_keybuff As String
Public archi As String
Public numarchi As Integer
Public UNICO As String

Public PSART_LLAVE_ALT As rdoQuery
Public art_llave_alt As rdoResultset

Public PSAUT_LLAVE As rdoQuery
Public PSAUT_MENOR As rdoQuery
Public PSPAR_LLAVE As rdoQuery
Public PSCLI_LLAVE As rdoQuery
Public PSCON_LLAVE As rdoQuery

Public PSCLI_MAYOR As rdoQuery

Public PSCAA_HISTO As rdoQuery
Public PSCLI_MAYOR2 As rdoQuery
Public PSVEN_LLAVE As rdoQuery
Public PSTRA_LLAVE As rdoQuery
Public PSTRA_MENU As rdoQuery
Public PSART_LLAVE As rdoQuery
Public PSART_LLAVE10 As rdoQuery
Public PSART_MAYOR As rdoQuery
Public PSARM_LLAVE As rdoQuery
Public PSCCM_LLAVE As rdoQuery
Public PSCCM_MAYOR As rdoQuery
Public PSCCM_MAYOR2 As rdoQuery
Public PSFAR_LLAVE As rdoQuery
Public PSFAR_MENOR As rdoQuery
Public PSFAR_MENOR2 As rdoQuery
Public PSFAR_MENOR3 As rdoQuery
Public PSPROC_MAYOR As rdoQuery
Public PSCAR_FAR As rdoQuery
Public PSCAA_LLAVE As rdoQuery

Public PSCAR_LLAVE As rdoQuery
Public PSCAR_MENOR As rdoQuery
Public PSALL_LLAVE As rdoQuery
Public PSALL_MENOR As rdoQuery
Public PSCAR_MAYOR As rdoQuery
Public PSTAB_LLAVE As rdoQuery
Public PSTAB_MAYOR As rdoQuery
Public PSPRO_LLAVE As rdoQuery
Public PSSUT_LLAVE As rdoQuery
Public PSSUT_MAYOR As rdoQuery
Public PSCOV_LLAVE As rdoQuery
Public PSCOV_MAYOR As rdoQuery
Public PSCHE_MENOR As rdoQuery
Public PSCHE_LLAVE As rdoQuery
Public PSCHE_OPER As rdoQuery
Public PSCHE_MAYOR As rdoQuery
Public PSCHE_MOVI As rdoQuery
Public PS_PAR As rdoQuery
Public PS_GEN As rdoQuery
Public PSCHE_REPO As rdoQuery
Public LLAVE As rdoQuery
Public cli_ruc As rdoResultset
Public PSCLI_RUC As rdoQuery
Public pr_numfac As rdoResultset
Public PSPR_NUMFAC As rdoQuery

 

Public cli_llave10 As rdoResultset
Public PSCLI_LLAVE10  As rdoQuery

Public cls_llave As rdoResultset
Public PSCLS_LLAVE As rdoQuery
Public RS_CC As rdoResultset
Public PSCC As rdoQuery
Public numfilas As Integer
Public f As Boolean
Public ws_fecha_dia As Date
Public WS_TALON As String * 1

Public Sub MUESTRA_USUario()
FORMGEN.i_CODUSU.Clear
usu.Requery
usu.MoveFirst
FORMGEN.i_CODUSU.AddItem ""
Do Until usu.EOF
  FORMGEN.i_CODUSU.AddItem usu!usu_key ' & "      " & par!PAR_CODCIA
  usu.MoveNext
Loop
End Sub


Public Sub LEER_CLI_LLAVE()
Select Case SQ_OPER
Case 1
PSCLI_LLAVE.rdoParameters(0) = pu_cp
PSCLI_LLAVE.rdoParameters(1) = pu_codclie
PSCLI_LLAVE.rdoParameters(2) = pu_codcia
cli_llave.Requery
GoTo salida

Case 2
PSCLI_MAYOR.rdoParameters(0) = pu_cp
PSCLI_MAYOR.rdoParameters(1) = pu_codclie
PSCLI_MAYOR.rdoParameters(2) = pu_codcia
cli_mayor.Requery
GoTo salida
Case 3
GoTo salida
Case 4
 PSCLI_RUC.rdoParameters(0) = pu_cp
 PSCLI_RUC.rdoParameters(1) = PUB_RUC
 PSCLI_RUC.rdoParameters(2) = pu_codcia
 cli_ruc.Requery
GoTo salida
Case 5
 PSCLS_LLAVE.rdoParameters(0) = pu_cp
 PSCLS_LLAVE.rdoParameters(1) = pu_codclie
 PSCLS_LLAVE.rdoParameters(2) = pu_codcia
 cls_llave.Requery
GoTo salida
Case 10
PSCLI_LLAVE10.rdoParameters(0) = pu_cp
PSCLI_LLAVE10.rdoParameters(1) = pu_codclie
PSCLI_LLAVE10.rdoParameters(2) = pu_codcia
cli_llave10.Requery
GoTo salida


End Select


salida:

End Sub

Public Sub LEER_VEN_LLAVE()
Select Case SQ_OPER
Case 1
PSVEN_LLAVE.rdoParameters(0) = pu_codcia
PSVEN_LLAVE.rdoParameters(1) = PUB_CODVEN

GoTo COMUN

Case 2
PSVEN_MAYOR.rdoParameters(0) = sq_keybuff
GoTo COMUN

End Select


COMUN:


ven_llave.Requery

End Sub

Public Sub CONEXION_GEN()
  Dim bdata As String
' On Error GoTo ALGUN_ERROR
  Dim wAcceso As String
  Dim success%
  Dim iStatusBarWidth As Integer
  Dim Srutas As String
  Dim ws_color As Integer
  wdsn = "dsn_datos"
  'wdsn = "dd"
  Splash.Label1.Caption = "Iniciando Conexión a Base de Datos..."
  DoEvents
  PUB_DSN = UCase(wdsn)
  ws_color = 3
  Srutas = CONS_ADMIN & "SONIDOS\Splash.WAV"
  iStatusBarWidth = 4075
  Screen.MousePointer = vbHourglass
  DoEvents
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  NL = Chr(13) & Chr(10)
  Set EN = rdoEnvironments(0)
  wAcceso = "159357852456"
  If Dir("C:\sysload.txt", vbArchive) = "" Then
     bdata = "bdatos"
     GoTo salta_archi
  End If
  Open "C:\sysLoad.txt" For Input As #1
  Do While Not EOF(1)
      Line Input #1, bdata
   Loop
  Close #1
  
  Open "C:\sysLoad.txt" For Output As #1
  Print #1, "bdatos"
  Close #1
  bdata = Trim(bdata)
salta_archi:
 wAcceso = "anteromariano"
  CONn$ = "dsn=" & wdsn & ";uid=sa;pwd=" & wAcceso & ";database=" & bdata & " ;"
  'CONn$ = "dsn=" & wdsn & ";uid=sa;pwd=" & wAcceso & ";database=bdatos;"
  DoEvents
  Set CN = EN.OpenConnection(" ", False, False, CONn$)
  CN.QueryTimeout = 90
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  DoEvents
  Call PlaySound(Srutas, 1, 1) 'Archivos de Sonidos
  Splash.Label1.Caption = "Conexión establecida..."
  
  pub_cadena = "SELECT * FROM GENERAL WHERE GEN_KEY <> ? ORDER BY GEN_KEY"
  Set PS_GEN = CN.CreateQuery("", pub_cadena)
  PS_GEN(0) = 0
  Set GEN = PS_GEN.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.lblEmpresa.Caption = Trim(GEN!GEN_NOMBRE)
  'If Val(GEN!gen_id) = 1 Then
  '   Splash.lblanexo.Visible = True
  '   DoEvents
  'End If
  DoEvents
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  DoEvents
  pub_cadena = "SELECT * FROM calendario WHERE CAL_CODCIA = ? AND CAL_FECHA >= ? AND CAL_FECHA <= ?  ORDER BY CAL_FECHA "
  Set PSCAL_LLAVE = CN.CreateQuery("", pub_cadena)
  PSCAL_LLAVE(0) = 0
  PSCAL_LLAVE(1) = LK_FECHA_DIA
  PSCAL_LLAVE(2) = LK_FECHA_DIA
  Set cal_llave = PSCAL_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
    
  pub_cadena = "SELECT * FROM CLIENTES WHERE CLI_CP=? AND CLI_CODCLIE  = ? AND CLI_CODCIA = ? ORDER BY CLI_CP ,CLI_CODCLIE"
  Set PSCLI_LLAVE = CN.CreateQuery("", pub_cadena)
  PSCLI_LLAVE(0) = 0
  PSCLI_LLAVE(1) = 0
  PSCLI_LLAVE(2) = 0
  Set cli_llave = PSCLI_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  
  pub_cadena = "SELECT CLI_NOMBRE, CLI_RUC_ESPOSO FROM CLIENTES WHERE CLI_CP=? AND CLI_CODCLIE  = ? AND CLI_CODCIA = ? ORDER BY CLI_CP ,CLI_CODCLIE"
  Set PSCLI_LLAVE10 = CN.CreateQuery("", pub_cadena)
  PSCLI_LLAVE10(0) = 0
  PSCLI_LLAVE10(1) = 0
  PSCLI_LLAVE10(2) = 0
  Set cli_llave10 = PSCLI_LLAVE10.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.Label1.Caption = "Iniciando Maestros de Archivos..."
  DoEvents
  pub_cadena = "SELECT * FROM CLISAL WHERE CLS_CP = ? AND CLS_CODCLIE  = ? AND CLS_CODCIA = ? ORDER BY CLS_CP"
  Set PSCLS_LLAVE = CN.CreateQuery("", pub_cadena)
  PSCLS_LLAVE(0) = 0
  PSCLS_LLAVE(1) = 0
  PSCLS_LLAVE(2) = 0
  Set cls_llave = PSCLS_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  
  
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  pub_cadena = "SELECT  CLI_CODCLIE FROM CLIENTES WHERE CLI_CP=? AND CLI_RUC_ESPOSO = ? AND CLI_CODCIA = ? "
  Set PSCLI_RUC = CN.CreateQuery("", pub_cadena)
  PSCLI_RUC(0) = ""
  PSCLI_RUC(1) = ""
  PSCLI_RUC(2) = ""
  Set cli_ruc = PSCLI_RUC.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)

  pub_cadena = "SELECT * FROM CLIENTES WHERE CLI_CP=? AND CLI_CODCLIE  >= ? AND CLI_CODCIA = ? ORDER BY CLI_CP ,CLI_CODCLIE"
  Set PSCLI_MAYOR = CN.CreateQuery("", pub_cadena)
  PSCLI_MAYOR(0) = 0
  PSCLI_MAYOR(1) = 0
  PSCLI_MAYOR(2) = 0
  Set cli_mayor = PSCLI_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  'pub_cadena = "SELECT * FROM PROCESOS WHERE PRO_CODCIA=? AND PRO_CODPRO=? ORDER BY PRO_CODCIA, PRO_CODPRO, PRO_ORDEN"
  'Set PSPROC_MAYOR = CN.CreateQuery("", pub_cadena)
  'PSPROC_MAYOR(0) = 0
  'PSPROC_MAYOR(1) = 0
  'Set proc_mayor = PSPROC_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)
  
  
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  Splash.Label1.Caption = "Iniciando Relaciones de D.B. ..."
  
  DoEvents
  'pub_cadena = "SELECT * FROM VEMAEST WHERE VEM_CODCIA  = ? AND VEM_CODVEN = ? ORDER BY VEM_CODCIA, VEM_CODVEN"
  'Set PSVEN_LLAVE = CN.CreateQuery("", pub_cadena)
  'PSVEN_LLAVE(0) = 0
  'PSVEN_LLAVE(1) = 0
  'Set ven_llave = PSVEN_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100

  pub_cadena = "SELECT * FROM TRANSACCION WHERE TRA_KEY = ? ORDER BY TRA_KEY"
  Set PSTRA_LLAVE = CN.CreateQuery("", pub_cadena)
  PSTRA_LLAVE(0) = 0
  Set tra_llave = PSTRA_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  pub_cadena = "SELECT * FROM TRANSACCION WHERE TRA_KEY >= ? AND TRA_FLAG_ACTIVO = 'A'  ORDER BY TRA_DESCRIPCION"
  Set PSTRA_MENU = CN.CreateQuery("", pub_cadena)
  PSTRA_MENU(0) = 0
  Set tra_menu = PSTRA_MENU.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  Splash.Label1.Caption = "Cargando Libros Contables..."
  DoEvents
  pub_cadena = "SELECT PER_INDEX, [Nombres y Apellidos] FROM listapersonal WHERE PER_INDEX = ? "
'  Set PSPER_BUSCA = CN.CreateQuery("", pub_cadena)
'  PSPER_BUSCA(0) = 0
 ' Set per_busca = PSPER_BUSCA.OpenResultset(rdOpenKeyset, rdConcurValues)

  'pub_cadena = "SELECT * FROM ARTI WHERE ART_KEY = ? AND ART_CODCIA = ? ORDER BY ART_CODCIA, ART_KEY"
  'Set PSART_LLAVE = CN.CreateQuery("", pub_cadena)
  'PSART_LLAVE(0) = 0
  'PSART_LLAVE(1) = 0
  'Set art_LLAVE = PSART_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  Splash.Label1.Caption = "Iniciando Componentes de Relación..."
  DoEvents
  'pub_cadena = "SELECT * FROM ARTI WHERE ART_KEY = ? AND ART_CODCIA = ? ORDER BY ART_CODCIA, ART_KEY"
  'Set PSART_LLAVE10 = CN.CreateQuery("", pub_cadena)
  'PSART_LLAVE10(0) = 0
  'PSART_LLAVE10(1) = 0
  'Set art_LLAVE10 = PSART_LLAVE10.OpenResultset(rdOpenKeyset, rdConcurValues)
  
  'pub_cadena = "SELECT * FROM ARTI WHERE ART_KEY >= ? AND ART_CODCIA=? ORDER BY ART_CODCIA, ART_KEY"
  'Set PSART_MAYOR = CN.CreateQuery("", pub_cadena)
  'PSART_MAYOR(0) = 0
  'PSART_MAYOR(1) = 0
  'Set art_mayor = PSART_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  'pub_cadena = "SELECT * FROM ARTI WHERE ART_ALTERNO = ? AND ART_CODCIA = ? ORDER BY ART_CODCIA, ART_ALTERNO"
  'Set PSART_LLAVE_ALT = CN.CreateQuery("", pub_cadena)
  'PSART_LLAVE_ALT(0) = 0
  'PSART_LLAVE_ALT(1) = 0
  'Set art_llave_alt = PSART_LLAVE_ALT.OpenResultset(rdOpenKeyset, rdConcurValues)
  DoEvents
  'pub_cadena = "SELECT * FROM ARTICULO WHERE ARM_CODART = ? AND ARM_CODCIA = ? ORDER BY ARM_CODART, ARM_CODCIA"
  'Set PSARM_LLAVE = CN.CreateQuery("", pub_cadena)
  'PSARM_LLAVE(0) = 0
  'PSARM_LLAVE(1) = 0
  'Set arm_llave = PSARM_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  'pub_cadena = "SELECT * FROM SUB_TRANSA WHERE SUT_CODTRA = ? AND SUT_SECUENCIA = ? ORDER BY SUT_CODTRA, SUT_SECUENCIA"
  'Set PSSUT_LLAVE = CN.CreateQuery("", pub_cadena)
  'PSSUT_LLAVE(0) = 0
  'PSSUT_LLAVE(1) = 0
  'Set SUT_LLAVE = PSSUT_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100

  pub_cadena = "SELECT * FROM CONTABILIDAD WHERE CNT_CODCIA= ? AND CNT_CODTRA = ? AND CNT_SECUENCIA = ? ORDER BY CNT_CODTRA, CNT_SECUENCIA"
  Set PSCNT_LLAVE = CN.CreateQuery("", pub_cadena)
  PSCNT_LLAVE(0) = 0
  PSCNT_LLAVE(1) = 0
  PSCNT_LLAVE(2) = 0
  Set cnt_llave = PSCNT_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  Splash.Label1.Caption = "Iniciando Archivos del Sistema Integrado..."
  DoEvents
  'pub_cadena = "SELECT * FROM CONTABILIDAD WHERE CNT_CODCIA= ? AND CNT_CODTRA = ?  ORDER BY CNT_CODTRA, CNT_SECUENCIA"
  'Set PSCNT_MAYOR = CN.CreateQuery("", pub_cadena)
  'PSCNT_MAYOR(0) = 0
  'PSCNT_MAYOR(1) = 0
  'Set cnt_mayor = PSCNT_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100

  'pub_cadena = "SELECT * FROM SUB_TRANSA WHERE SUT_CODTRA = ?  ORDER BY SUT_CODTRA, SUT_SECUENCIA"
  'Set PSSUT_MAYOR = CN.CreateQuery("", pub_cadena)
  'PSSUT_MAYOR(0) = 0
  'Set SUT_MAYOR = PSSUT_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  pub_cadena = "SELECT * FROM CCMAEST WHERE CCM_CODBAN = ? AND CCM_CODCIA = ? ORDER BY CCM_CODBAN, CCM_CODCIA "
  Set PSCCM_LLAVE = CN.CreateQuery("", pub_cadena)
  PSCCM_LLAVE(0) = 0
  PSCCM_LLAVE(1) = 0
  Set ccm_llave = PSCCM_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100

  'pub_cadena = "SELECT * FROM CCMAEST WHERE CCM_CODCIA = ? AND CCM_CODBAN > ?   ORDER BY CCM_CODBAN"
  'Set PSCCM_MAYOR = CN.CreateQuery("", pub_cadena)
  'PSCCM_MAYOR(0) = 0
  'PSCCM_MAYOR(1) = 0
  'Set ccm_mayor = PSCCM_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  'pub_cadena = "SELECT * FROM CCMAEST WHERE CCM_CODCIA = ? AND CCM_CODBAN > ?   ORDER BY CCM_codban"
  'Set PSCCM_MAYOR2 = CN.CreateQuery("", pub_cadena)
  'PSCCM_MAYOR2(0) = 0
  'PSCCM_MAYOR2(1) = 0
  'Set ccm_mayor2 = PSCCM_MAYOR2.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.Label1.Caption = "Iniciando Archivos de Movimientos Integrado..."
  DoEvents
  
  'pub_cadena = "SELECT * FROM facart WHERE FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_NUMSER = ? AND FAR_FBG=? AND FAR_NUMFAC = ?  ORDER BY FAR_TIPMOV, FAR_CODCIA, FAR_NUMSER, FAR_FBG, FAR_NUMFAC, FAR_NUMSEC"
  'Set PSFAR_LLAVE = CN.CreateQuery("", pub_cadena)
  'PSFAR_LLAVE(0) = 0
  'PSFAR_LLAVE(1) = 0
  'PSFAR_LLAVE(2) = 0
  'PSFAR_LLAVE(3) = 0
  'PSFAR_LLAVE(4) = 0
  'Set far_llave = PSFAR_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
   
  'pub_cadena = "SELECT * FROM facart WHERE FAR_CP = ? AND FAR_CODCLIE = ? AND FAR_FECHA >= ? ORDER BY FAR_CP, far_CODCLIE, FAR_FECHA, FAR_NUMSER, FAR_NUMFAC"
 ' Set PSFAR_CODCLIE = CN.CreateQuery("", pub_cadena)
 ' PSFAR_CODCLIE(0) = 0
 ' PSFAR_CODCLIE(1) = 0
 ' PSFAR_CODCLIE(2) = LK_FECHA_DIA
 ' Set far_codcli = PSFAR_CODCLIE.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  pub_cadena = "SELECT MOV_CODCIA FROM MOVICONT WHERE MOV_CODCIA = ? AND MOV_TIPMOV = ? AND (MOV_FECHA >= ? AND MOV_FECHA <= ?) AND MOV_FBG = ? AND MOV_SERIE = ? AND MOV_NUMFAC = ? "
  Set PSPR_NUMFAC = CN.CreateQuery("", pub_cadena)
  PSPR_NUMFAC.MaxRows = 1
  PSPR_NUMFAC(0) = 0
  PSPR_NUMFAC(1) = 0
  PSPR_NUMFAC(2) = LK_FECHA_DIA
  PSPR_NUMFAC(3) = LK_FECHA_DIA
  PSPR_NUMFAC(4) = 0
  PSPR_NUMFAC(5) = 0
  PSPR_NUMFAC(6) = 0
  Set pr_numfac = PSPR_NUMFAC.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)


  'pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = ? AND PRE_CODART = ?  AND PRE_SECUENCIA = ? ORDER BY PRE_SECUENCIA"
  'Set PSPRE_LLAVE = CN.CreateQuery("", pub_cadena)
  'PSPRE_LLAVE(0) = 0
  'PSPRE_LLAVE(1) = 0
  'PSPRE_LLAVE(2) = 0
  'Set pre_llave = PSPRE_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  'pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = ? AND PRE_CODART = ?  ORDER BY PRE_EQUIV"
 ' Set PSPRE_MAYOR = CN.CreateQuery("", pub_cadena)
 ' PSPRE_MAYOR(0) = 0
 ' PSPRE_MAYOR(1) = 0
 ' Set pre_mayor = PSPRE_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  'pub_cadena = "SELECT FAR_NUMFAC FROM FACART WHERE FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_FBG = ? AND FAR_NUMSER = ? ORDER BY FAR_TIPMOV, FAR_CODCIA, FAR_FBG , FAR_NUMSER, FAR_NUMFAC DESC"
  'Set PSFAR_MENOR = CN.CreateQuery("", pub_cadena)
  'PSFAR_MENOR(0) = 0
  'PSFAR_MENOR(1) = 0
  'PSFAR_MENOR(2) = 0
  'PSFAR_MENOR(3) = 0
  'PSFAR_MENOR.MaxRows = 1
  'Set far_menor = PSFAR_MENOR.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
     
  'pub_cadena = "SELECT * FROM facart WHERE FAR_TIPMOV = ? AND FAR_CODCIA = ? AND FAR_FBG = ? AND FAR_NUMSER=? AND FAR_FECHA = ? ORDER BY FAR_TIPMOV, FAR_CODCIA, FAR_FBG , FAR_NUMSER, FAR_NUMFAC"
  'Set PSFAR_MENOR2 = CN.CreateQuery("", pub_cadena)
  'PSFAR_MENOR2(0) = 0
  'PSFAR_MENOR2(1) = 0
  'PSFAR_MENOR2(2) = 0
  'PSFAR_MENOR2(3) = 0
  'PSFAR_MENOR2(4) = LK_FECHA_DIA
  'Set far_menor2 = PSFAR_MENOR2.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  'pub_cadena = "SELECT * FROM facart WHERE FAR_FECHA = ? AND FAR_NUMOPER = ? AND FAR_CODCIA = ? AND FAR_ESTADO <> 'E' ORDER BY FAR_CODCIA, FAR_FECHA, FAR_NUMOPER, FAR_NUMSEC"
  'Set PSFAR_MENOR3 = CN.CreateQuery("", pub_cadena)
  'PSFAR_MENOR3(0) = LK_FECHA_DIA
  'PSFAR_MENOR3(1) = 0
  'PSFAR_MENOR3(2) = 0
  'Set far_menor3 = PSFAR_MENOR3.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  Splash.Label1.Caption = "Cargando Modulos para Integracion del Sistema ..."
  DoEvents
  'pub_cadena = "SELECT * FROM CHEQUES WHERE CHE_CODBAN= ? AND CHE_CODCIA = ?  AND CHE_CHESER=?  AND CHE_FECHA = ? ORDER BY CHE_CODBAN, CHE_CODCIA, CHE_CHESER, CHE_CHENUM "
  'Set PSCHE_MENOR = CN.CreateQuery("", pub_cadena)
  'PSCHE_MENOR(0) = 0
  'PSCHE_MENOR(1) = 0
  'PSCHE_MENOR(2) = 0
  'PSCHE_MENOR(3) = LK_FECHA_DIA
  'Set che_menor = PSCHE_MENOR.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  'pub_cadena = "SELECT * FROM CHEQUES WHERE CHE_NUMOPER= ? AND CHE_FECHA=? ORDER BY CHE_FECHA,CHE_NUMOPER"
  'Set PSCHE_OPER = CN.CreateQuery("", pub_cadena)
  'PSCHE_OPER(0) = 0
  'PSCHE_OPER(1) = LK_FECHA_DIA
  'Set che_oper = PSCHE_OPER.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100

  'pub_cadena = "SELECT * FROM CHEQUES WHERE CHE_CODBAN = ? AND CHE_CODCIA=?  AND CHE_FECHA >= ? ORDER BY CHE_CODBAN, CHE_CODCIA, CHE_FECHA, CHE_NUMOPER"
  'Set PSCHE_REPO = CN.CreateQuery("", pub_cadena)
  'PSCHE_REPO(0) = 0
  'PSCHE_REPO(1) = 0
  'PSCHE_REPO(2) = LK_FECHA_DIA
  'Set che_repo = PSCHE_REPO.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  'pub_cadena = "SELECT * FROM CHEQUES WHERE CHE_CODBAN= ? AND CHE_CODCIA = ? AND CHE_CHESER=?  AND CHE_CHENUM=? AND CHE_CHESEC <=? ORDER BY CHE_CODBAN, CHE_CODCIA, CHE_CHESER, CHE_CHENUM, CHE_CHESEC"
  'Set PSCHE_LLAVE = CN.CreateQuery("", pub_cadena)
  'PSCHE_LLAVE(0) = 0
  'PSCHE_LLAVE(1) = 0
  'PSCHE_LLAVE(2) = 0
  'PSCHE_LLAVE(3) = 0
  'PSCHE_LLAVE(4) = 0
  'Set che_llave = PSCHE_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  DoEvents
  
  'pub_cadena = "SELECT * FROM PARGEN WHERE PAR_CODCIA = ?  ORDER BY PAR_CODCIA "
  'Set PSPAR_LLAVE = CN.CreateQuery("", pub_cadena)
  'PSPAR_LLAVE(0) = 0
  'Set par_llave = PSPAR_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100

  'pub_cadena = "SELECT * FROM CARTERA WHERE CAR_CP = ? AND CAR_CODCLIE = ? AND CAR_CODCIA =? AND CAR_SERDOC = ? AND CAR_NUMDOC =? AND CAR_TIPDOC= ?  ORDER BY CAR_CP , CAR_CODCLIE, CAR_CODCIA, CAR_SERDOC, CAR_NUMDOC, CAR_TIPDOC "
  'Set PSCAR_LLAVE = CN.CreateQuery("", pub_cadena)
  'PSCAR_LLAVE(0) = 0
  'PSCAR_LLAVE(1) = 0
  'PSCAR_LLAVE(2) = 0
  'PSCAR_LLAVE(3) = 0
  'PSCAR_LLAVE(4) = 0
  'PSCAR_LLAVE(5) = 0
  'Set car_llave = PSCAR_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  'pub_cadena = "SELECT * FROM CARTERA WHERE  CAR_CODCIA = ? AND CAR_CP = ? AND CAR_CODCLIE = ?  AND CAR_IMPORTE <> 0 ORDER BY CAR_CODCIA, CAR_CP, CAR_CODCLIE, CAR_TIPDOC, CAR_NUMGUIA , CAR_FECHA_INGR, CAR_PRECIO"
  'Set PSCAR_MAYOR = CN.CreateQuery("", pub_cadena)
  'PSCAR_MAYOR(0) = 0
  'PSCAR_MAYOR(1) = 0
  'PSCAR_MAYOR(2) = 0
  'Set car_mayor = PSCAR_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  
  'pub_cadena = "SELECT * FROM CARTERA WHERE CAR_CODCIA =? AND CAR_CP = ?  AND CAR_TIPDOC=? AND CAR_SERDOC = ?  ORDER BY CAR_CODCIA, CAR_CP ,CAR_TIPDOC, CAR_SERDOC, CAR_NUMDOC DESC"
  'Set PSCAR_MENOR = CN.CreateQuery("", pub_cadena)
  'PSCAR_MENOR(0) = 0
  'PSCAR_MENOR(1) = 0
  'PSCAR_MENOR(2) = 0
  'PSCAR_MENOR(3) = 0
  'PSCAR_MENOR.MaxRows = 1
  'Set car_menor = PSCAR_MENOR.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
 ' pub_cadena = "SELECT * FROM CARTERA WHERE CAR_CODCIA =? AND CAR_FBG = ?  AND CAR_TIPDOC=? AND CAR_NUMSER = ?  AND CAR_NUMFAC = ? "
 ' Set PSCAR_FAR = CN.CreateQuery("", pub_cadena)
 ' PSCAR_FAR(0) = 0
 ' PSCAR_FAR(1) = 0
 ' PSCAR_FAR(2) = 0
 ' PSCAR_FAR(3) = 0
 ' PSCAR_FAR(4) = 0
 ' Set car_far = PSCAR_FAR.OpenResultset(rdOpenKeyset, rdConcurValues)
'  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  
  'pub_cadena = "SELECT * FROM CARACU WHERE CAA_CP=? AND CAA_CODCLIE = ? AND CAA_CODCIA=? AND CAA_TIPDOC=? AND CAA_FECHA=? AND CAA_NUM_OPER=? ORDER BY CAA_FECHA"
  'Set PSCAA_HISTO = CN.CreateQuery("", pub_cadena)
  'PSCAA_HISTO(0) = 0
  'PSCAA_HISTO(1) = 0
  'PSCAA_HISTO(2) = 0
  'PSCAA_HISTO(3) = 0
  'PSCAA_HISTO(4) = LK_FECHA_DIA
  'PSCAA_HISTO(5) = 0
  'Set caa_histo = PSCAA_HISTO.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100


  pub_cadena = "SELECT * FROM COMAEST WHERE COM_CUENTA = ? AND COM_CODCIA = ? ORDER BY COM_CUENTA, COM_CODCIA "
  Set PSCOM_LLAVE = CN.CreateQuery("", pub_cadena)
  PSCOM_LLAVE(0) = 0
  PSCOM_LLAVE(1) = 0
  Set com_llave = PSCOM_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  pub_cadena = "SELECT * FROM COMSAL WHERE COS_CUENTA = ? AND COS_CODCIA = ? AND COS_NRO_ANO = ? "
  Set PSCOS_LLAVE = CN.CreateQuery("", pub_cadena)
  PSCOS_LLAVE(0) = 0
  PSCOS_LLAVE(1) = 0
  PSCOS_LLAVE(2) = 0
  Set cos_llave = PSCOS_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
    
  pub_cadena = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND COM_CUENTA > ? ORDER BY COM_CODCIA, COM_CUENTA"
  Set PSCOM_MAYOR = CN.CreateQuery("", pub_cadena)
  PSCOM_MAYOR(0) = 0
  PSCOM_MAYOR(1) = 0
  Set com_mayor = PSCOM_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  pub_cadena = "SELECT * FROM TABLAS WHERE TAB_TIPREG = ? AND TAB_CODCIA = ? ORDER BY TAB_CODCIA,TAB_TIPREG, TAB_NUMTAB"
  Set PSTAB_MAYOR = CN.CreateQuery("", pub_cadena)
  PSTAB_MAYOR(0) = 0
  PSTAB_MAYOR(1) = 0
  Set tab_mayor = PSTAB_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)

  pub_cadena = "SELECT * FROM TABLAS WHERE TAB_TIPREG = ? AND TAB_NUMTAB = ? AND TAB_CODCIA = ? ORDER BY TAB_CODCIA,TAB_TIPREG, TAB_NUMTAB"
  Set PSTAB_LLAVE = CN.CreateQuery("", pub_cadena)
  PSTAB_LLAVE(0) = 0
  PSTAB_LLAVE(1) = 0
  PSTAB_LLAVE(2) = 0
  Set tab_llave = PSTAB_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  
  pub_cadena = "SELECT * FROM TABLAS WHERE TAB_TIPREG = ? AND  TAB_CODCIA = ? AND TAB_CODART = ? ORDER BY TAB_CODCIA,TAB_TIPREG, TAB_NUMTAB"
  Set PSTAB_MENOR = CN.CreateQuery("", pub_cadena)
  PSTAB_MENOR(0) = 0
  PSTAB_MENOR(1) = 0
  PSTAB_MENOR(2) = 0
  Set tab_menor = PSTAB_MENOR.OpenResultset(rdOpenKeyset, rdConcurValues)

  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  DoEvents
  
  'pub_cadena = "SELECT * FROM ALLOG WHERE   ALL_CODCIA = ? AND ALL_FECHA_DIA = ?  ORDER BY ALL_NUMOPER "
  'Set PSALL_LLAVE = CN.CreateQuery("", pub_cadena)
  'PSALL_LLAVE(0) = 0
  'PSALL_LLAVE(1) = LK_FECHA_DIA
  'Set all_llave = PSALL_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  'pub_cadena = "SELECT ALL_NUMOPER FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_DIA = ?   ORDER BY ALL_NUMOPER DESC "
  'Set PSALL_MENOR = CN.CreateQuery("", pub_cadena)
  'PSALL_MENOR(0) = 0
  'PSALL_MENOR(1) = LK_FECHA_DIA
  'PSALL_MENOR.MaxRows = 1
'  Set all_menor = PSALL_MENOR.OpenResultset(rdOpenKeyset, rdConcurValues)

  
  pub_cadena = "SELECT * FROM PARGEN WHERE PAR_CODCIA = ?  ORDER BY PAR_CODCIA"
  Set PSPAR_LLAVE = CN.CreateQuery("", pub_cadena)
  PSPAR_LLAVE(0) = 0
  Set par_llave = PSPAR_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  
  pub_cadena = "SELECT PAR_CODCIA, PAR_NOMBRE, PAR_NOMBRE_CORTO FROM PARGEN WHERE PAR_CODCIA = ?  ORDER BY PAR_CODCIA "
  Set PSPAR_MULTI = CN.CreateQuery("", pub_cadena)
  PSPAR_MULTI(0) = ""
  Set par_multi = PSPAR_MULTI.OpenResultset(rdOpenKeyset, rdConcurValues)
  
  
  'pub_cadena = "SELECT * FROM AUTORIZACION WHERE AUT_CODCIA = ? and AUT_KEY  = ?  ORDER BY AUT_KEY , aut_secuencia"
  'Set PSAUT_LLAVE = CN.CreateQuery("", pub_cadena)
  'PSAUT_LLAVE(0) = 0
  'PSAUT_LLAVE(1) = 0
  'Set aut_llave = PSAUT_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  DoEvents
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  'pub_cadena = "SELECT * FROM AUTORIZACION WHERE AUT_CODCIA= ? AND AUT_KEY  < ? and AUT_FECHA >= ? ORDER BY AUT_KEY"
  'Set PSAUT_MENOR = CN.CreateQuery("", pub_cadena)
  'PSAUT_MENOR(0) = 0
  'PSAUT_MENOR(1) = 0
  'PSAUT_MENOR(2) = LK_FECHA_DIA
  'Set aut_menor = PSAUT_MENOR.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100

  
  pub_cadena = "SELECT * FROM PARGEN WHERE PAR_CODCIA <> ? ORDER BY PAR_CODCIA"
  Set PS_PAR = CN.CreateQuery("", pub_cadena)
  PS_PAR(0) = 0
  Set par = PS_PAR.OpenResultset(rdOpenKeyset, rdConcurValues)

  pub_cadena = "SELECT * FROM TRANSACCION WHERE TRA_FLAG_ACTIVO = 'A' AND TRA_KEY <= 8000 ORDER BY TRA_KEY"
  Set lis_tra = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurReadOnly) ', rdConcurLock)

  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  Splash.Label1.Caption = "Iniciando Maestros de Usuarios..."
  DoEvents
  
  pub_cadena = "SELECT * FROM usuarios WHERE USU_KEY = ?  ORDER BY USU_KEY"
  Set PSUSU_LLAVE = CN.CreateQuery("", pub_cadena)
  PSUSU_LLAVE(0) = 0
  Set usu_llave = PSUSU_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  
  'pub_cadena = "SELECT * FROM facart WHERE FAR_TIPMOV = ? AND FAR_NUMSER = ? AND FAR_NUMFAC = ?  ORDER BY  FAR_TIPMOV, FAR_NUMSER, FAR_NUMFAC"
  'Set PSFAR_LLAVE2 = CN.CreateQuery("", pub_cadena)
  'PSFAR_LLAVE2(0) = 0
  'PSFAR_LLAVE2(1) = 0
  'PSFAR_LLAVE2(2) = 0
  'DoEvents
  'Set far_llave2 = PSFAR_LLAVE2.OpenResultset(rdOpenKeyset, rdConcurValues)
  
  'pub_cadena = "SELECT * FROM ARTI WHERE ART_NOMBRE >= ?  ORDER BY ART_NOMBRE"
  'Set PSX = CN.CreateQuery("", pub_cadena)
  'PSX(0) = 0
  'Set X = PSX.OpenResultset(rdOpenKeyset, rdConcurValues)
  
  
  'pub_cadena = "SELECT * FROM CCMAEST WHERE CCM_CODBAN = ? AND CCM_CODCIA = ?  ORDER BY CCM_CODBAN, CCM_CODCIA"
  'Set PSCCM_LLAVE = CN.CreateQuery("", pub_cadena)
  'PSCCM_LLAVE(0) = 0
  'PSCCM_LLAVE(1) = 0
  'Set ccm_llave = PSCCM_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  
  pub_cadena = "SELECT * FROM COPARAM WHERE COP_CODCIA = ? "
  Set PSCOP_LLAVE = CN.CreateQuery("", pub_cadena)
  PSCOP_LLAVE(0) = 0
  Set cop_llave = PSCOP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

  

'  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  'Splash.rctStatusBar.Value = 4560 ' 46
  'MsgBox Splash.rctStatusBar.Value
  DoEvents
  
  pub_cadena = "SELECT * FROM USUARIOS ORDER BY usu_key"
  Set usu = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)
'  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  cad = "SELECT * FROM GRUPOS ORDER BY GRU_NOMBRE"
  Set gru = CN.OpenResultset(cad, rdOpenKeyset, rdConcurValues)
'  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  DoEvents
  Splash.Label1.Caption = "Terminando las Conexiones..."
  DoEvents
   
'  pub_cadena = "SELECT * FROM CONTROLL  ORDER BY CON_KEY"
'  Set PSCON_LLAVE = CN.CreateQuery("", pub_cadena)
'  PSCON_LLAVE.RowsetSize = 1
'  Set con_llave = PSCON_LLAVE.OpenResultset(rdOpenKeyset, rdConcurLock)
Exit Sub
ALGUN_ERROR:
 MsgBox "Verificar si esta en la Red de WINDOWS    ... Detalle : " & Err.Description, 48, Pub_Titulo
 End
End Sub




Public Sub LEER_TRA_LLAVE()

Select Case SQ_OPER
Case 1
PSTRA_LLAVE.rdoParameters(0) = PUB_CODTRA
GoTo COMUN

Case 2
  'PSTRA_MAYOR.rdoParameters(0) = sq_keybuff
  'GoTo COMUN
   PSTRA_MENU.rdoParameters(0) = PUB_INICIO
   tra_menu.Requery
    Exit Sub
End Select


COMUN:
tra_llave.Requery

End Sub


Public Sub LEER_ART_LLAVE()
If LK_EMP_PTO = "A" Then
  pu_codcia = "00"
End If
Select Case SQ_OPER
Case 1
  PSART_LLAVE.rdoParameters(0) = PUB_KEY
  PSART_LLAVE.rdoParameters(1) = pu_codcia
GoTo COMUN

Case 2
    PSART_MAYOR.rdoParameters(0) = PUB_KEY
    PSART_MAYOR.rdoParameters(1) = pu_codcia
    art_mayor.Requery
    Exit Sub
Case 3
  PSART_LLAVE_ALT.rdoParameters(0) = pu_alterno
  PSART_LLAVE_ALT.rdoParameters(1) = pu_codcia
  art_llave_alt.Requery
  Exit Sub
Case 10
  PSART_LLAVE10.rdoParameters(0) = PUB_KEY
  PSART_LLAVE10.rdoParameters(1) = pu_codcia
  art_LLAVE10.Requery
  Exit Sub
End Select

COMUN:
art_LLAVE.Requery

End Sub

Public Sub LEER_ARM_LLAVE()
Select Case SQ_OPER
Case 1
PSARM_LLAVE.rdoParameters(0) = PUB_CODART
PSARM_LLAVE.rdoParameters(1) = pu_codcia

GoTo COMUN

Case 2
PSARM_MAYOR.rdoParameters(0) = sq_keybuff
GoTo COMUN

End Select

COMUN:
arm_llave.Requery

End Sub
Public Sub LEER_CAL_LLAVE(Optional tc)
Select Case SQ_OPER
Case 1
PUB_CODCIA = "00"
If Not IsMissing(tc) Then
   If tc = 1 Then PUB_CODCIA = LK_CODCIA
End If
PSCAL_LLAVE.rdoParameters(0) = PUB_CODCIA
PSCAL_LLAVE.rdoParameters(1) = PUB_CAL_INI
PSCAL_LLAVE.rdoParameters(2) = PUB_CAL_FIN
cal_llave.Requery
End Select

salida:

End Sub
Public Sub LEER_CCM_LLAVE()
Select Case SQ_OPER
Case 1
PSCCM_LLAVE.rdoParameters(0) = PUB_CODBAN
PSCCM_LLAVE.rdoParameters(1) = pu_codcia

GoTo COMUN

Case 2
PSCCM_MAYOR.rdoParameters(0) = pu_codcia
PSCCM_MAYOR.rdoParameters(1) = PUB_CODBAN
ccm_mayor.Requery
Exit Sub

Case 3
PSCCM_MAYOR2.rdoParameters(0) = pu_codcia
PSCCM_MAYOR2.rdoParameters(1) = PUB_CODBAN
ccm_mayor2.Requery

End Select


COMUN:
ccm_llave.Requery

End Sub
Public Sub LEER_COM_LLAVE()
Dim wscodcia   As String * 2
'wscodcia = LK_CODCIA
wscodcia = PUB_CODCIA
If LK_EMP_PTO = "A" Then
 wscodcia = "00"
End If
Select Case SQ_OPER
Case 1
PSCOM_LLAVE.rdoParameters(0) = PUB_CUENTA
PSCOM_LLAVE.rdoParameters(1) = wscodcia

GoTo COMUN

Case 2
 PSCOM_MAYOR.rdoParameters(0) = wscodcia
 PSCOM_MAYOR.rdoParameters(1) = PUB_CUENTA
 com_mayor.Requery
 Exit Sub

Case 3

LEER_OTRO:
 PSCOS_LLAVE.rdoParameters(0) = PUB_CUENTA
 PSCOS_LLAVE.rdoParameters(1) = wscodcia
 PSCOS_LLAVE.rdoParameters(2) = Format(LK_FECHA_COP1, "yyyy")
  ' Print LK_FECHA_COP2
 cos_llave.Requery
 If cos_llave.EOF Then
    cos_llave.AddNew
    cos_llave!COS_CUENTA = PUB_CUENTA
    cos_llave!COS_CODCIA = wscodcia
    cos_llave!COS_NRO_ANO = Format(LK_FECHA_COP1, "yyyy")
    cos_llave.Update
    GoTo LEER_OTRO
 End If
Exit Sub

End Select


COMUN:
com_llave.Requery

End Sub

Public Sub LEER_PAR_LLAVE()
Select Case SQ_OPER
Case 1
PSPAR_LLAVE.rdoParameters(0) = PUB_CODCIA

GoTo COMUN

End Select


COMUN:
par_llave.Requery

End Sub

Public Sub LEER_CAR_LLAVE()

Select Case SQ_OPER
Case 1
PSCAR_LLAVE.rdoParameters(0) = pu_cp
PSCAR_LLAVE.rdoParameters(1) = pu_codclie
PSCAR_LLAVE.rdoParameters(2) = pu_codcia
PSCAR_LLAVE.rdoParameters(3) = PUB_SERDOC
PSCAR_LLAVE.rdoParameters(4) = PUB_NUMDOC
PSCAR_LLAVE.rdoParameters(5) = PUB_TIPDOC
car_llave.Requery
GoTo SALIR


Case 2
PSCAR_MAYOR.rdoParameters(0) = pu_codcia
PSCAR_MAYOR.rdoParameters(1) = pu_cp
PSCAR_MAYOR.rdoParameters(2) = pu_codclie
car_mayor.Requery
GoTo SALIR

Case 3
PSCAR_MENOR.rdoParameters(0) = pu_codcia
PSCAR_MENOR.rdoParameters(1) = pu_cp
PSCAR_MENOR.rdoParameters(2) = PUB_TIPDOC
PSCAR_MENOR.rdoParameters(3) = PUB_SERDOC
car_menor.Requery
GoTo SALIR

Case 4
PSCAR_FAR.rdoParameters(0) = pu_codcia
PSCAR_FAR.rdoParameters(1) = PU_FBG
PSCAR_FAR.rdoParameters(2) = PUB_TIPDOC
PSCAR_FAR.rdoParameters(3) = PUB_NUMSER
PSCAR_FAR.rdoParameters(4) = PUB_NUMFAC
car_far.Requery
GoTo SALIR

End Select



SALIR:
End Sub

Public Sub LEER_CAA_LLAVE()

Select Case SQ_OPER
Case 1
PSCAA_HISTO.rdoParameters(0) = pu_cp
PSCAA_HISTO.rdoParameters(1) = pu_codclie
PSCAA_HISTO.rdoParameters(2) = pu_codcia
PSCAA_HISTO.rdoParameters(3) = PUB_TIPDOC
PSCAA_HISTO.rdoParameters(4) = PUB_FECHA
PSCAA_HISTO.rdoParameters(5) = PUB_NUM_OPER
caa_histo.Requery

End Select

End Sub

Public Sub LEER_PRE_LLAVE()
If LK_EMP_PTO = "A" Then
  pu_codcia = "00"
End If

Select Case SQ_OPER
Case 1
PSPRE_LLAVE.rdoParameters(0) = pu_codcia
PSPRE_LLAVE.rdoParameters(1) = PUB_CODART
PSPRE_LLAVE.rdoParameters(2) = PUB_SECUEN
pre_llave.Requery
Case 2
PSPRE_MAYOR.rdoParameters(0) = pu_codcia
PSPRE_MAYOR.rdoParameters(1) = PUB_CODART
pre_mayor.Requery

End Select
End Sub

Public Sub LEER_FAR_LLAVE()

Select Case SQ_OPER
Case 1
PSFAR_LLAVE.rdoParameters(0) = PU_TIPMOV
PSFAR_LLAVE.rdoParameters(1) = pu_codcia
PSFAR_LLAVE.rdoParameters(2) = PU_NUMSER
PSFAR_LLAVE.rdoParameters(3) = PU_FBG
PSFAR_LLAVE.rdoParameters(4) = PU_NUMFAC
GoTo FARLLAVE

Case 2
PSFAR_CODCLIE.rdoParameters(0) = pu_cp
PSFAR_CODCLIE.rdoParameters(1) = pu_codclie
PSFAR_CODCLIE.rdoParameters(2) = PUB_FECHA

GoTo FARcodclie

Case 3
PSFAR_MENOR.rdoParameters(0) = PU_TIPMOV
PSFAR_MENOR.rdoParameters(1) = pu_codcia
PSFAR_MENOR.rdoParameters(2) = PU_FBG
PSFAR_MENOR.rdoParameters(3) = PU_NUMSER
far_menor.Requery
GoTo SALIR

Case 4
PSFAR_MENOR2.rdoParameters(0) = PU_TIPMOV
PSFAR_MENOR2.rdoParameters(1) = pu_codcia
PSFAR_MENOR2.rdoParameters(2) = PU_FBG
PSFAR_MENOR2.rdoParameters(3) = PU_NUMSER
PSFAR_MENOR2.rdoParameters(4) = pu_fecha
far_menor2.Requery
GoTo SALIR
Case 5
PSFAR_MENOR3.rdoParameters(0) = PUB_FECHA
PSFAR_MENOR3.rdoParameters(1) = PUB_NUM_OPER_EXT
PSFAR_MENOR3.rdoParameters(2) = LK_CODCIA
far_menor3.Requery
GoTo SALIR

End Select


FARLLAVE:
far_llave.Requery
GoTo SALIR

FARcodclie:
far_codcli.Requery
GoTo SALIR




SALIR:
End Sub
Public Sub LEER_CHE_LLAVE()
Select Case SQ_OPER
Case 1
PSCHE_LLAVE(0) = PUB_CODBAN
PSCHE_LLAVE(1) = LK_CODCIA
PSCHE_LLAVE(2) = PUB_CHESER
PSCHE_LLAVE(3) = PUB_CHENUM
PSCHE_LLAVE(4) = PUB_CHESEC
GoTo CHELLAVE
Case 2
PSCHE_OPER(0) = PUB_NUM_OPER
PSCHE_OPER(1) = PUB_FECHA
che_oper.Requery

GoTo SALIR

Case 3
PSCHE_MENOR(0) = PUB_CODBAN
PSCHE_MENOR(1) = LK_CODCIA
PSCHE_MENOR(2) = PUB_CHESER
PSCHE_MENOR(3) = PUB_FECHA
GoTo CHEMENOR
Case 4
  PSCHE_REPO.rdoParameters(0) = PUB_CODBAN
  PSCHE_REPO.rdoParameters(1) = LK_CODCIA
  PSCHE_REPO.rdoParameters(2) = PUB_FECHA
  che_repo.Requery
  GoTo SALIR

End Select


CHELLAVE:
che_llave.Requery
GoTo SALIR

CHEMENOR:
che_menor.Requery
GoTo SALIR

SALIR:
End Sub


Public Sub LEER_ALL_LLAVE()

Select Case SQ_OPER
Case 1
PSALL_LLAVE.rdoParameters(0) = pu_codcia
PSALL_LLAVE.rdoParameters(1) = PUB_FECHA
GoTo ALLLLAVE

Case 2
PSALL_MENOR.rdoParameters(0) = pu_codcia
PSALL_MENOR.rdoParameters(1) = PUB_FECHA
all_menor.Requery
GoTo SALIR


End Select


ALLLLAVE:
all_llave.Requery
GoTo SALIR



SALIR:

End Sub

Public Sub LEER_TAB_LLAVE()
Select Case SQ_OPER
Case 1
PSTAB_LLAVE.rdoParameters(0) = PUB_TIPREG
PSTAB_LLAVE.rdoParameters(1) = PUB_NUMTAB
PSTAB_LLAVE.rdoParameters(2) = PUB_CODCIA
GoTo LLAVE

Case 2
PSTAB_MAYOR.rdoParameters(0) = PUB_TIPREG
PSTAB_MAYOR.rdoParameters(1) = PUB_CODCIA
GoTo mayor
Case 3
PSTAB_MENOR.rdoParameters(0) = PUB_TIPREG
PSTAB_MENOR.rdoParameters(1) = PUB_CODCIA
PSTAB_MENOR.rdoParameters(2) = PUB_CODART
tab_menor.Requery
GoTo fin
End Select

LLAVE:
tab_llave.Requery
GoTo fin

mayor:
tab_mayor.Requery


fin:
End Sub
Public Sub LEER_ZON_LLAVE()
'  PSZON_LLAVE.rdoParameters(0) = PUB_TIPZON
'  PSZON_LLAVE.rdoParameters(1) = PUB_NUMZON
'  zon_llave.Requery
End Sub

Public Function ENTERO(TEXTO As String) As Boolean
Dim LARGO As Integer
Dim i, X As Integer
Dim DIG As Integer
LARGO = Len(TEXTO)
i = LARGO
ENTERO = True
Do Until i = 0
   DIG = Asc(Mid(TEXTO, i, 1))
   If (DIG > 47 And DIG < 58) Then
       X = 0
   Else
       ENTERO = False
       Exit Do
   End If
   i = i - 1
  
   Loop

End Function
Public Sub LEER_AUT_LLAVE()
Select Case SQ_OPER
Case 1
PSAUT_LLAVE.rdoParameters(0) = pu_codcia
PSAUT_LLAVE.rdoParameters(1) = pub_autkey
GoTo COMUN

Case 3
PSAUT_MENOR.rdoParameters(0) = pu_codcia
PSAUT_MENOR.rdoParameters(1) = pub_autkey
PSAUT_MENOR.rdoParameters(2) = PUB_FECHA
aut_menor.Requery
GoTo salida

End Select

COMUN:
aut_llave.Requery

salida:
End Sub

Public Sub LEER_PROC_LLAVE()
Select Case SQ_OPER
Case 2
PSPROC_MAYOR.rdoParameters(0) = PUB_CODCIA
PSPROC_MAYOR.rdoParameters(1) = PUB_CODPRO
proc_mayor.Requery

End Select


fin:
End Sub


Public Sub LEER_SUT_LLAVE()
Select Case SQ_OPER
Case 1
PSSUT_LLAVE.rdoParameters(0) = PUB_CODTRA
PSSUT_LLAVE.rdoParameters(1) = PUB_SECUENCIA
GoTo COMUN

Case 2
PSSUT_MAYOR.rdoParameters(0) = PUB_CODTRA
GoTo COMUN2

End Select


COMUN:
SUT_LLAVE.Requery

GoTo fin

COMUN2:
SUT_MAYOR.Requery
GoTo fin


fin:
End Sub
Public Sub LEER_CNT_LLAVE()
Select Case SQ_OPER
Case 1
PSCNT_LLAVE.rdoParameters(0) = PUB_CODCIA
PSCNT_LLAVE.rdoParameters(1) = PUB_CODTRA
PSCNT_LLAVE.rdoParameters(2) = PUB_SECUENCIA
cnt_llave.Requery
GoTo fin

Case 2
PSCNT_MAYOR.rdoParameters(0) = PUB_CODCIA
PSCNT_MAYOR.rdoParameters(1) = PUB_CODTRA
cnt_mayor.Requery
GoTo fin

End Select

fin:
End Sub


Public Sub LEER_COV_LLAVE()
Select Case SQ_OPER
Case 1
PSCOV_LLAVE.rdoParameters(1) = PUB_CODCONT
PSCOV_LLAVE.rdoParameters(0) = PUB_CODCIA

GoTo COMUN

Case 2
PSCOV_MAYOR.rdoParameters(0) = PUB_CODCIA
GoTo COMUN2

End Select

COMUN:
cov_llave.Requery
GoTo fin

COMUN2:
cov_mayor.Requery

fin:
End Sub

