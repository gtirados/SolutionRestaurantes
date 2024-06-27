Attribute VB_Name = "Módulo1"
'******************************
Public mItem As MSComctlLib.ListItem
Public rsT As New ADODB.Recordset
Public cnT As New ADODB.Connection
Public msg As String
'******************************
Public PUB_DSN As String
Public CN As rdoConnection
Public CONn As String
Public EN As rdoEnvironment
'Dim espaciot As Workspace
Public ODBCRUTA As String
Public PSX As rdoQuery
Public PSTRA As rdoQuery
Public PSVEN As rdoQuery

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
Public PSFFF_LLAVE As rdoQuery
Public FFF_LLAVE As rdoResultset

Public pac_llave  As rdoResultset
Public PSPAC_LLAVE As rdoQuery

Public cov_voucher  As rdoResultset

Public calen As rdoResultset

Public PSSER_LLAVE As rdoQuery
Public ser_llave  As rdoResultset


Public PSZON_LLAVE As rdoQuery
Public far_llave2 As rdoResultset
Public ven As rdoResultset
Public cont As rdoResultset
Public tra As rdoResultset
Public gen As rdoResultset
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
Public che_mayor As rdoResultset
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


Public Sub CONEXION_GEN()
' On Error GoTo ALGUN_ERROR
  Dim success%
  Dim iStatusBarWidth As Integer
  Dim Srutas As String
  Dim ws_color As Integer
  Dim wAcceso As String
  wdsn = "dsn_datos"
  'wdsn = "dsn_datos2"
  'wdsn = "dd"
  PUB_DSN = UCase(wdsn)
  ws_color = 3
  Srutas = "C:\ADMIN\SONIDOS\Splash.WAV"
  iStatusBarWidth = 4075
  Screen.MousePointer = vbHourglass
  DoEvents
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  NL = Chr(13) & Chr(10)
  Set EN = rdoEnvironments(0)
  'CONn$ = "dsn=" & wdsn & ";uid=abel;pwd=abel;database=bdatos;"
  'CONn$ = "dsn=" & wdsn & ";uid=jorge;pwd=jjj99;database=bdatos;"
  
  
  wAcceso = "anteromariano"
  CONn$ = "dsn=" & wdsn & ";uid=sa;pwd=" & wAcceso & ";database=bdatos;"
  DoEvents
  Set CN = EN.OpenConnection(" ", False, False, CONn$)
  CN.QueryTimeout = 90
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  DoEvents
  Call PlaySound(Srutas, 1, 1) 'Archivos de Sonidos
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  DoEvents
  pub_cadena = "SELECT * FROM calendario WHERE CAL_CODCIA = ? AND CAL_FECHA >= ? AND CAL_FECHA <= ?  ORDER BY CAL_FECHA "
  Set PSCAL_LLAVE = CN.CreateQuery("", pub_cadena)
  PSCAL_LLAVE(0) = 0
  PSCAL_LLAVE(1) = LK_FECHA_DIA
  PSCAL_LLAVE(2) = LK_FECHA_DIA
  Set cal_llave = PSCAL_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
    
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100

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
  
  pub_cadena = "SELECT * FROM ARTI WHERE ART_KEY = ? AND ART_CODCIA = ? ORDER BY ART_CODCIA, ART_KEY"
  Set PSART_LLAVE = CN.CreateQuery("", pub_cadena)
  PSART_LLAVE(0) = 0
  PSART_LLAVE(1) = 0
  DoEvents
  Set art_LLAVE = PSART_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  pub_cadena = "SELECT * FROM ARTI WHERE ART_KEY = ? AND ART_CODCIA = ? ORDER BY ART_CODCIA, ART_KEY"
  Set PSART_LLAVE10 = CN.CreateQuery("", pub_cadena)
  PSART_LLAVE10(0) = 0
  PSART_LLAVE10(1) = 0
  Set art_LLAVE10 = PSART_LLAVE10.OpenResultset(rdOpenKeyset, rdConcurValues)
  
  pub_cadena = "SELECT * FROM ARTI WHERE ART_KEY >= ? AND ART_CODCIA=? ORDER BY ART_CODCIA, ART_KEY"
  Set PSART_MAYOR = CN.CreateQuery("", pub_cadena)
  PSART_MAYOR(0) = 0
  PSART_MAYOR(1) = 0
  Set art_mayor = PSART_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  pub_cadena = "SELECT * FROM ARTI WHERE ART_ALTERNO = ? AND ART_CODCIA = ? ORDER BY ART_CODCIA, ART_ALTERNO"
  Set PSART_LLAVE_ALT = CN.CreateQuery("", pub_cadena)
  PSART_LLAVE_ALT(0) = 0
  PSART_LLAVE_ALT(1) = 0
  DoEvents
  Set art_llave_alt = PSART_LLAVE_ALT.OpenResultset(rdOpenKeyset, rdConcurValues)

  pub_cadena = "SELECT * FROM ARTICULO WHERE ARM_CODART = ? AND ARM_CODCIA = ? ORDER BY ARM_CODART, ARM_CODCIA"
  Set PSARM_LLAVE = CN.CreateQuery("", pub_cadena)
  PSARM_LLAVE(0) = 0
  PSARM_LLAVE(1) = 0
  Set arm_llave = PSARM_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  pub_cadena = "SELECT * FROM SUB_TRANSA WHERE SUT_CODTRA = ? AND SUT_SECUENCIA = ? ORDER BY SUT_CODTRA, SUT_SECUENCIA"
  Set PSSUT_LLAVE = CN.CreateQuery("", pub_cadena)
  PSSUT_LLAVE(0) = 0
  PSSUT_LLAVE(1) = 0
  Set SUT_LLAVE = PSSUT_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100

  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100

  pub_cadena = "SELECT * FROM SUB_TRANSA WHERE SUT_CODTRA = ?  ORDER BY SUT_CODTRA, SUT_SECUENCIA"
  Set PSSUT_MAYOR = CN.CreateQuery("", pub_cadena)
  PSSUT_MAYOR(0) = 0
  Set SUT_MAYOR = PSSUT_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = ? AND PRE_CODART = ?  AND PRE_SECUENCIA = ? ORDER BY PRE_SECUENCIA"
  Set PSPRE_LLAVE = CN.CreateQuery("", pub_cadena)
  PSPRE_LLAVE(0) = 0
  PSPRE_LLAVE(1) = 0
  PSPRE_LLAVE(2) = 0
  Set pre_llave = PSPRE_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  pub_cadena = "SELECT * FROM PRECIOS WHERE PRE_CODCIA = ? AND PRE_CODART = ?  ORDER BY PRE_EQUIV"
  Set PSPRE_MAYOR = CN.CreateQuery("", pub_cadena)
  PSPRE_MAYOR(0) = 0
  PSPRE_MAYOR(1) = 0
  Set pre_mayor = PSPRE_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100

  pub_cadena = "SELECT * FROM COMAEST WHERE COM_CUENTA = ? AND COM_CODCIA = ? ORDER BY COM_CUENTA, COM_CODCIA "
  Set PSCOM_LLAVE = CN.CreateQuery("", pub_cadena)
  PSCOM_LLAVE(0) = 0
  PSCOM_LLAVE(1) = 0
  Set com_llave = PSCOM_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
    
  pub_cadena = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND COM_CUENTA > ? ORDER BY COM_CODCIA, COM_CUENTA"
  Set PSCOM_MAYOR = CN.CreateQuery("", pub_cadena)
  PSCOM_MAYOR(0) = 0
  PSCOM_MAYOR(1) = 0
  Set com_mayor = PSCOM_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)
  
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  DoEvents
  

  pub_cadena = "SELECT * FROM PARGEN WHERE PAR_CODCIA = ?  ORDER BY PAR_CODCIA "
  Set PSPAR_LLAVE = CN.CreateQuery("", pub_cadena)
  PSPAR_LLAVE(0) = 0
  Set par_llave = PSPAR_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  pub_cadena = "SELECT * FROM CONTABILIDAD WHERE CNT_CODCIA= ? AND CNT_CODTRA = ? AND CNT_SECUENCIA = ? ORDER BY CNT_CODTRA, CNT_SECUENCIA"
  Set PSCNT_LLAVE = CN.CreateQuery("", pub_cadena)
  PSCNT_LLAVE(0) = 0
  PSCNT_LLAVE(1) = 0
  PSCNT_LLAVE(2) = 0
  Set cnt_llave = PSCNT_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  
  pub_cadena = "SELECT * FROM CONTABILIDAD WHERE CNT_CODCIA= ? AND CNT_CODTRA = ?  ORDER BY CNT_CODTRA, CNT_SECUENCIA"
  Set PSCNT_MAYOR = CN.CreateQuery("", pub_cadena)
  PSCNT_MAYOR(0) = 0
  PSCNT_MAYOR(1) = 0
  Set cnt_mayor = PSCNT_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)
  
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
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
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  pub_cadena = "SELECT * FROM PARGEN WHERE PAR_CODCIA = ?  ORDER BY PAR_CODCIA"
  Set PSPAR_LLAVE = CN.CreateQuery("", pub_cadena)
  PSPAR_LLAVE(0) = 0
  Set par_llave = PSPAR_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  pub_cadena = "SELECT * FROM PARARC WHERE PAC_CODCIA = ?  ORDER BY PAC_CODCIA"
  Set PSPAC_LLAVE = CN.CreateQuery("", pub_cadena)
  PSPAC_LLAVE(0) = ""
  Set pac_llave = PSPAC_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  
  'pub_cadena = "SELECT * FROM SERIES WHERE SER_CODCIA = ? AND SER_FBG = ? AND SER_NUMERO = ? ORDER BY SER_CODCIA , SER_FBG, SER_NUMERO"
  'Set PSSER_LLAVE = CN.CreateQuery("", pub_cadena)
  'PSSER_LLAVE(0) = 0
  'PSSER_LLAVE(1) = 0
  'PSSER_LLAVE(2) = 0
  'Set ser_llave = PSSER_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  
  
  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100

  
  pub_cadena = "SELECT * FROM PARGEN WHERE PAR_CODCIA <> ? ORDER BY PAR_CODCIA"
  Set PS_PAR = CN.CreateQuery("", pub_cadena)
  PS_PAR(0) = 0
  Set par = PS_PAR.OpenResultset(rdOpenKeyset, rdConcurValues)

  pub_cadena = "SELECT * FROM TRANSACCION WHERE TRA_FLAG_ACTIVO = 'A' AND TRA_KEY <= 8000 ORDER BY TRA_KEY"
  Set lis_tra = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurReadOnly) ', rdConcurLock)

  Splash.rctStatusBar.Value = Splash.rctStatusBar.Value + 100
  
  pub_cadena = "SELECT * FROM usuarios WHERE USU_KEY = ?  ORDER BY USU_KEY"
  Set PSUSU_LLAVE = CN.CreateQuery("", pub_cadena)
  PSUSU_LLAVE(0) = 0
  Set usu_llave = PSUSU_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
  
  
  'pub_cadena = "SELECT * FROM ARTI WHERE ART_NOMBRE >= ?  ORDER BY ART_NOMBRE"
  'Set PSX = CN.CreateQuery("", pub_cadena)
  'Set X = PSX.OpenResultset(rdOpenKeyset, rdConcurValues)
  

  DoEvents
  pub_cadena = "SELECT * FROM GENERAL WHERE GEN_KEY <> ? ORDER BY GEN_KEY"
  Set PS_GEN = CN.CreateQuery("", pub_cadena)
  PS_GEN(0) = 0
  Set gen = PS_GEN.OpenResultset(rdOpenKeyset, rdConcurValues)
  
  pub_cadena = "SELECT * FROM USUARIOS ORDER BY usu_key"
  Set usu = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurValues) ' rdConcurReadOnly) ', rdConcurLock)

  cad = "SELECT * FROM GRUPOS ORDER BY GRU_NOMBRE"
  Set gru = CN.OpenResultset(cad, rdOpenKeyset, rdConcurValues)

  DoEvents
  Splash.Label1.Caption = "... Conexión Establecida ..."
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


Public Sub LEER_CAL_LLAVE(Optional TC)
Select Case SQ_OPER
Case 1
PUB_CODCIA = "00"
If Not IsMissing(TC) Then
   If TC = 1 Then PUB_CODCIA = LK_CODCIA
End If

PSCAL_LLAVE.rdoParameters(0) = PUB_CODCIA
PSCAL_LLAVE.rdoParameters(1) = PUB_CAL_INI
PSCAL_LLAVE.rdoParameters(2) = PUB_CAL_FIN
cal_llave.Requery
End Select


salida:

End Sub
Public Sub LEER_PAR_LLAVE()
Select Case SQ_OPER
Case 1
 PSPAR_LLAVE.rdoParameters(0) = PUB_CODCIA
 GoTo COMUN
Case 2
 PSPAC_LLAVE.rdoParameters(0) = PUB_CODCIA
 pac_llave.Requery
 Exit Sub
End Select


COMUN:
par_llave.Requery

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
GoTo FIN
End Select

LLAVE:
tab_llave.Requery
GoTo FIN

mayor:
tab_mayor.Requery


FIN:
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


FIN:
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

GoTo FIN

COMUN2:
SUT_MAYOR.Requery
GoTo FIN


FIN:
End Sub
Public Sub LEER_CNT_LLAVE()
Select Case SQ_OPER
Case 1
PSCNT_LLAVE.rdoParameters(0) = PUB_CODCIA
PSCNT_LLAVE.rdoParameters(1) = PUB_CODTRA
PSCNT_LLAVE.rdoParameters(2) = PUB_SECUENCIA
cnt_llave.Requery
GoTo FIN

Case 2
PSCNT_MAYOR.rdoParameters(0) = PUB_CODCIA
PSCNT_MAYOR.rdoParameters(1) = PUB_CODTRA
cnt_mayor.Requery
GoTo FIN

End Select

FIN:
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

Public Sub LEER_COM_LLAVE()
Dim wscodcia   As String * 2
wscodcia = LK_CODCIA
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

End Select


COMUN:
com_llave.Requery

End Sub


Public Sub LEER_SER_LLAVE()
PSSER_LLAVE(0) = LK_CODCIA
PSSER_LLAVE(1) = PUB_FBG
PSSER_LLAVE(2) = PUB_NUMSER
ser_llave.Requery
End Sub

