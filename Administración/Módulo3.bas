Attribute VB_Name = "Módulo3"
Option Explicit
Public FILAX As Integer
Public LKCHEK As Boolean
Public LK_DIRECCION As String
Public LK_EMP_PTO As String * 1
Public PUB_CLAVE As String
Public LK_TIPO_CAMBIO As Currency
Public LK_MONEDA As String * 1
Public LK_USU_STOCK As String * 1
Public LK_FLAG_LIMITE As String * 1
Public WR_line As Integer
Public ws_conta As Integer
Public Const WHORA = #12:00:00 PM#
Public PUB_RUTA_REPORTE As String
Public PUB_RUTA_OTRO As String
Public CONST_DSN As String
Public CONST_SERVER As String
Public CONST_UID As String
Public CONST_PWD As String
Public PUB_ODBC As String
Public PUB_LINEAS As Integer
Public PUB_FLAG As Integer
Public PUB_NUM As Long
Public pub_signo_ccm As Integer
Public pub_signo_car As Integer
Public pub_signo_caja As Integer
Public pub_signo_arm As Integer
Public pub_signo_ped As Integer
Public pub_ojo As String * 1
Public PUB_FBG As String * 1
Public PASA As Integer
Public TABLA_TAG(300)
Public PUB_CP As String * 1
Public PUB_CV As String * 1
Public PUB_CODCLIE As Long
Public PUB_CHENUM As Long
Public pu_codcia As String * 2
Public pu_cp As String * 1
Public pu_codclie As Currency
Public PUB_CHENUM_EXT As Long
Public PUB_SITUACION_ACT As Integer
Public PUB_SITUACION_ANT As Integer
Public PUB_SD As String * 1
Public PUB_CANT_CHEQ As Integer
Public PUB_NUM_INI As Currency
Public PUB_CHESER As String * 3
Public PUB_CHESEC As Integer
Public pub_autkey As Long
Public PUB_LIMCRE_ACT As Currency
Public PUB_PRECIO2_ORIG As Currency
Public PUB_LIMCRE_ANT As Currency
Public PUB_TIPO_BLOQ_act1 As String * 1
Public PUB_TIPO_BLOQ_act2 As String * 1
Public PUB_TIPO_BLOQ_act3 As String * 1
Public PUB_TIPO_BLOQ_act4 As String * 1
Public PUB_TIPO_BLOQ_ant1 As String * 1
Public PUB_TIPO_BLOQ_ant2 As String * 1
Public PUB_TIPO_BLOQ_ant3 As String * 1
Public PUB_TIPO_BLOQ_ant4 As String * 1
Public PUB_SECUEN As Integer
Public PUB_INICIO  As Integer
Public PUB_CODCIA As String * 2
Public PUB_CODCIAL As String * 2
Public PUB_CODUSU As String
Public PUB_NOMCIA As String
Public PUB_TIPDOC As String * 2
Public PUB_NUMKAR As Long
Public PUB_NUMDOC As Long
Public PUB_NUMGUIA As Long
Public PUB_CODART As Long
Public PUB_PEDSER As Integer
Public PUB_PEDSEC As Integer
Public PUB_PEDFAC As Long
Public PUB_IMPORTE As Currency
Public PUB_IMPORTE_AMORT As Currency
Public pub_numplan As Double
Public pub_diasA As Long
Public pub_dias As Integer
Public PUB_TIPMOV As Integer
Public pub_total_2455 As Currency
Public PUB_TIPMOV_REF As Integer
Public PUB_IMPTO2 As Currency
Public PUB_BRUTO2 As Currency


Public PUB_CODCONT As String * 12

Public PUB_DS As String * 1
Public PUB_NOMBRE_BANCO As String * 30
Public PUB_NUM_CHEQUE As String * 12

Public PUB_CODBAN As Integer 'Integer
Public PUB_CONCEPTO As String
Public PUB_FLAG_VENCIDO As Integer
Public PUB_FECHA As Date
Public PUB_FECHA_INGR As Date
Public PUB_FECHA_VCTO As Date
Public PUB_NUMSER As Integer
Public PUB_NUMFAC As Long
Public PUB_NUMSER_C As Integer
Public PUB_NUMFAC_C As Long
Public PUB_NOMART As String
Public PUB_SERDOC As Integer
Public PUB_NETO As Currency
Public PUB_TOTAL As Currency
Public PUB_IMPTO As Currency
Public PUB_FLETE As Currency
Public PUB_SUBTOTAL As Currency
Public PUB_SUBTOTAL2 As Currency
Public PUB_SUBTOTAL_BAK As Currency
Public pub_deuda As Currency

Public PUB_DESCTO As Currency
Public PUB_GASTOS As Currency
Public PUB_PRECIO As Currency
Public PUB_PRECIO2 As Currency
Public PUB_COSPRO As Currency
Public PUB_CANTIDAD As Currency
Public PUB_UNIDAD As Currency
Public PUB_JABAS As Integer
Public PUB_CODPRO As Integer
Public PUB_NOMPRO As String * 50
Public PUB_ESPESOR As Integer
Public PUB_LINEA As String * 20
Public PUB_CALIDAD As Integer
Public PUB_SECUENCIA As Integer

Public PUB_CODVEN As Integer
Public PUB_KEY As Long
Public PUB_CODTRA As Integer
Public PUB_IS As String * 1
Public PUB_VF As Boolean
Public PUB_CODIGO As Long
Public PUB_NUM_OPER As Integer
Public PUB_NUM_OPER_EXT As Integer
Public PUB_NUM_OPER_XXX As Integer

Public PUB_NUMTAB As Integer
Public PUB_TIPREG As Integer
Public PUB_USUARIO As String
Public PUB_TIPZON As Integer
Public PUB_NUMZON As Integer
Public PUB_ABREVIADO As String * 5
Public PUB_PAG1 As Integer
Public PUB_PAG2 As Integer
Public PUB_PAG3 As Integer
Public PUB_PAG4 As Integer
Public PUB_PAG5 As Integer
Public PUB_PAG6 As Integer

Public PUB_PAGX As Integer
Public WR_PAG As Integer
Public PUB_CODALI As String * 6


Public LK_PRECIO As String * 1
Public LK_CODUSU As String
Public LK_CODCIA As String * 2
Public LK_CIA_REF As String * 2
Public LK_IGV As Currency
Public LK_FECHA_DIA As Date
Public LK_NOMCORTO As String
Public LK_OPERADOR As String
Public LK_TERMINAL As Integer
Public LK_NIVUSU As Integer
Public LK_PRINTER As Boolean
Public LK_CODTRA As Integer
Public LK_NOMTRA As String
Public LK_DIASEM As Integer
Public LK_FECHA_SGTE As Date
Public LK_DIA_LETRAS As String
Public LK_MES_LETRAS As String
Public LK_FECHA_AYER As Date
Public LK_TABLA As String
Public lk_GRUPOS(10) As Integer
Public lk_TRANSA(10) As Integer
Public lk_CODTRAS(10) As String * 20
Public lk_OTROS() As String * 2
Public lk_OTROS_Count As Integer
Public LK_COBRADOR As Integer
Public LK_RELACION_STOCK As Currency
Public LK_FLAG_FACTURACION As String * 1
Public LK_FLAG_ALTERNO As String * 1
Public LK_FLAG_ORIGINAL As String * 1
Public pu_alterno As String * 10
Public LK_EMP As String * 3
Public LK_FLAG_CALCULO As String * 1
Public LK_FAC_IMP As String * 1
Public WCABE As Integer

Public NUM_CONTAB(99) As Currency

Public TEXTOX(20) As String
Public NOMBREX(20) As String
Public ETIQUETAX(20) As String
Public WS_IMPORTE As Currency
Public WS_NETO As Currency
Public WS_DESCTO As Currency
Public WS_IMPTO As Currency
Public WS_IGV As Currency
Public WS_BRUTO As Currency
Public SUB_CANT As Currency
Public SUB_FLETE As Currency
Public SUB_JABAS As Currency
Public SUB_UNIDAD As Currency
Public PU_TIPMOV As Integer
Public WS_LETRA_ACTIVA As Boolean
Public PU_NUMFAC As Currency
Public PU_NUMSER As Integer
Public PU_FBG As String * 1
Public pu_fecha As Date

Public OP_FORM As String * 1
Public FRM_STATUS As String * 1 '1 ES ACTIVO 0 DESACTIVO
Public Tab_Clave As Integer
Public ACEPTA As String * 1
Public wTABLA As String
Public CAMPOS As Integer
Public Posi_Reg As Integer
Public Cta_Add As Integer
Public ws_pub_mensaje As String
Public NL As String
Public fila As Integer
Public WStop As Boolean
Declare Function SetWindowPos Lib "User32" (ByVal h&, ByVal hb&, ByVal X&, ByVal Y&, ByVal cx&, ByVal cy&, ByVal f&) As Long
Declare Function FindWindow Lib "User32" (ByVal lpClassName As Any, ByVal lpWindowName As Any) As Long
Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Const FLAGS = 1 Or 2
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Pub_Titulo As String
Public Const Pub_Estilo = vbYesNo + vbQuestion + vbDefaultButton2
Public Pub_Respuesta As Integer
Public pub_mensaje As String
Public pub_cadena As String
Public Codi_Grupo As Integer 'permiso
Public Nom_User As String  'permisos
Public WS_FLAG As Integer
Public WS_CONTADOR As Integer
Public NRO_CONTROL As Integer
Public WS_INDICE_RETORNO As Integer
Public tab_precioss(5) As String * 7

Public Sub CenterMe(frmForm As Form)
  frmForm.Left = (Screen.Width - frmForm.Width) / 2
  frmForm.Top = (Screen.Height - frmForm.Height) / 2
End Sub


Public Function CONVER_LETRAS(NUMERO_BASE As Currency) As String
Dim NUM(16) As String * 9
Dim DECEN(10) As String * 9
Dim CENTEN(10) As String * 9
Dim VECTOR(5) As String * 3
Dim NUMERO As String
Dim WS_DEC As String
Dim RESTA As Currency
Dim ENTERO As Double
Dim LETRAS(120) As String * 1
Dim wa As String
Dim i, N, t As Integer
Dim C As Integer
Dim DU As Integer
Dim d As Integer
Dim u As Integer
Dim cdu As Currency
NUM(1) = "UN "
NUM(2) = "DOS"
NUM(3) = "TRES"
NUM(4) = "CUATRO"
NUM(5) = "CINCO"
NUM(6) = "SEIS"
NUM(7) = "SIETE"
NUM(8) = "OCHO"
NUM(9) = "NUEVE"
NUM(10) = "DIEZ"
NUM(11) = "ONCE"
NUM(12) = "DOCE"
NUM(13) = "TRECE"
NUM(14) = "CATORCE"
NUM(15) = "QUINCE"

DECEN(1) = "DIEZ"
DECEN(2) = "VEI"
DECEN(3) = "TREI"
DECEN(4) = "CUARE"
DECEN(5) = "CINCUE"
DECEN(6) = "SESE"
DECEN(7) = "SETE"
DECEN(8) = "OCHE"
DECEN(9) = "NOVE"

CENTEN(1) = "CIEN"
CENTEN(2) = "DOS"
CENTEN(3) = "TRES"
CENTEN(4) = "CUATRO"
CENTEN(5) = "QUINI"
CENTEN(6) = "SEIS"
CENTEN(7) = "SETE"
CENTEN(8) = "OCHO"
CENTEN(9) = "NOVE"

'*** PARTE DECIMAL ******
ENTERO = Int(NUMERO_BASE)
RESTA = NUMERO_BASE - ENTERO
NUMERO = NUMERO_BASE - RESTA
NUMERO = Int(NUMERO)
NUMERO = Format(NUMERO, "000000000000")
WS_DEC = RESTA * 100
WS_DEC = "y " & Format(WS_DEC, "00") & "/100  NUEVOS SOLES"
VECTOR(1) = Mid(NUMERO, 1, 3)
VECTOR(2) = Mid(NUMERO, 4, 3)
VECTOR(3) = Mid(NUMERO, 7, 3)
VECTOR(4) = Mid(NUMERO, 10, 3)
pub_cadena = ""
For i = 1 To 4
    t = 0
    N = 1
    cdu = Val(VECTOR(i))
    C = Int(cdu / 100)
    DU = cdu - (C * 100)
    d = Int(DU / 10)
    u = DU - (d * 10)
    If cdu > 99 Then
        wa = Trim(CENTEN(C))
        pub_cadena = pub_cadena + wa
        If C > 1 Then
            If C = 5 Then
                wa = "ENTOS "
                pub_cadena = pub_cadena + wa
            Else
                wa = "CIENTOS "
                pub_cadena = pub_cadena + wa
            End If
        Else
            If DU = 0 Then
                wa = " "
                pub_cadena = pub_cadena + wa
            Else
                wa = "TO "
                pub_cadena = pub_cadena + wa
            End If
        End If
    End If
    If DU > 0 And DU <> 20 Then
        If DU > 19 Then
            wa = Trim(DECEN(d))
            pub_cadena = pub_cadena + wa
            If u = 0 Then
                wa = "NTA"
                pub_cadena = pub_cadena + wa
            Else
                wa = "NTI"
                pub_cadena = pub_cadena + wa
                wa = Trim(NUM(u))
                pub_cadena = pub_cadena + wa
            End If
        Else
            If DU < 16 Then
                wa = " " & Trim(NUM(DU))
                pub_cadena = pub_cadena + wa
            Else
                wa = "DIECI"
                pub_cadena = pub_cadena + wa
                wa = Trim(NUM(u))
                pub_cadena = pub_cadena + wa
            End If
        End If

    End If
    If DU = 20 Then
        wa = "VEINTE"
        pub_cadena = pub_cadena + wa
    End If
    t = t + cdu
    If cdu <> u And i = 1 Then
        wa = " MIL "
        pub_cadena = pub_cadena + wa
    End If
    If t <> 0 And i = 2 Then
        If cdu = 1 And t = 1 Then
            wa = " MILLON "
            pub_cadena = pub_cadena + wa
        Else
            wa = "MILLONES "
            pub_cadena = pub_cadena + wa
        End If
    End If
    If cdu <> 0 And i = 3 Then
        wa = " MIL "
        pub_cadena = pub_cadena + wa
    End If
    If i = 4 Then
        wa = " " & WS_DEC
        pub_cadena = pub_cadena + wa
    End If

Next

If Left(pub_cadena, 2) = " y" Then
 CONVER_LETRAS = "CERO " & pub_cadena
Else
 CONVER_LETRAS = pub_cadena
End If
End Function

Public Function NUM_NEGA(Optional valor) As String
Dim temp1 As String
If Val(valor) < 0 Then
 temp1 = Format(valor, "##,###,##0.00")
 temp1 = Mid(temp1, 2, Len(temp1))
 temp1 = "(" & temp1 & ")"
 NUM_NEGA = temp1
Else
 NUM_NEGA = Format(valor, "##,###,##0.00")
 NUM_NEGA = NUM_NEGA + " "
End If
End Function
Public Sub PROC_LISVIEW(LV1 As ListView, Optional wmax)
On Error GoTo SALE
Dim wmaximo As Integer
Dim itmX As ListItem
If Not IsMissing(wmax) Then wmaximo = wmax Else wmaximo = 1000
Set PSX = CN.CreateQuery("", archi)
Set X = PSX.OpenResultset(rdOpenKeyset)
X.Requery
LV1.ListItems.Clear
LV1.ColumnHeaders.Clear
If X.EOF Then LV1.Visible = False: Exit Sub
LV1.Top = 1800
LV1.Left = 3000
LV1.Width = 6500
LV1.Height = 3200
LV1.Visible = True
If numarchi = 3 Then ' para codigos alternos
 LV1.ColumnHeaders.Add 1, , "Alterno", 1000
 LV1.ColumnHeaders.Add 2, , "Descripción", 4000
 LV1.ColumnHeaders.Add 3, , "Original", 0
ElseIf numarchi = 1 Then
 LV1.ColumnHeaders.Add 1, , "Descripción", 2500
 LV1.ColumnHeaders.Add 2, , "Cod.", 600
 LV1.ColumnHeaders.Add 3, , "Dirección", 4000
Else
 LV1.ColumnHeaders.Add 1, , "Descripción", 4000
 LV1.ColumnHeaders.Add 2, , "Cod.", 400
End If
Do Until X.EOF Or X.AbsolutePosition - 1 >= wmaximo
   If numarchi = 1 Or numarchi = 3 Then Set itmX = LV1.ListItems.Add(, , Trim(CStr(X.rdoColumns(3)))) Else Set itmX = LV1.ListItems.Add(, , Trim(CStr(X.rdoColumns(2))))
   If numarchi = 3 Then itmX.SubItems(1) = Trim(CStr(X.rdoColumns(2))) Else: itmX.SubItems(1) = Trim(CStr(X.rdoColumns(0)))
   If numarchi = 3 Then itmX.SubItems(2) = Trim(CStr(X.rdoColumns(0)))
   If numarchi = 1 Then itmX.SubItems(2) = Trim(CStr(X.rdoColumns(4))) + " # " + Trim(CStr(X.rdoColumns(6)))
   itmX.Tag = X.AbsolutePosition
   X.MoveNext
Loop
LV1.ToolTipText = "Encontrados : " & itmX.Tag & "/" & X.RowCount & " Muestra un Maximo de: " & wmaximo
Exit Sub
SALE:
If Err.Number = 40002 Then Exit Sub Else MsgBox Err.Description, 48, Pub_Titulo
End Sub

Public Sub ACTUALIZA_CIA(WCAMBIO_CIA As String)
LK_CODCIA = WCAMBIO_CIA
SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
LEER_PAR_LLAVE
If par_llave.EOF Then
  MsgBox "NO Existe Compañia Avisar !!!! Posible Error .. Intente Nuevamente ....", 48, Pub_Titulo
  End
  Exit Sub
End If
 If Nulo_Valors(par_llave!PAR_FLAG_ALTERNO) = "A" Then
   LK_FLAG_ORIGINAL = " "
   LK_FLAG_ALTERNO = "A"
 Else
   LK_FLAG_ORIGINAL = "A"
   LK_FLAG_ALTERNO = " "
   MDIForm1.Toolbar1.Buttons.Item(13).Enabled = False
 End If
 If LK_FLAG_ORIGINAL <> "A" Then
  MDIForm1.Toolbar1.Buttons.Item(13).Enabled = True
  If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
   MDIForm1.Toolbar1.Buttons.Item(13).ToolTipText = "Modo Alterno"
   DoEvents
   MDIForm1.Toolbar1.Buttons.Item(13).Image = 5
   DoEvents
  Else
   MDIForm1.Toolbar1.Buttons.Item(13).ToolTipText = "Modo Original"
   DoEvents
   MDIForm1.Toolbar1.Buttons.Item(13).Image = 40
   DoEvents
  End If
 End If
LK_EMP = Nulo_Valors(par_llave!PAR_EMPRESA)
LK_MONEDA = Nulo_Valors(par_llave!PAR_MONEDA_fac)
If Trim(LK_MONEDA) = "" Then MsgBox "DEFINIR LA MONEDA DE COMPAÑIA", 48, Pub_Titulo
LK_FLAG_FACTURACION = Nulo_Valors(par_llave!PAR_FLAG_FACTURACION)
LK_FLAG_CALCULO = Nulo_Valors(par_llave!PAR_FLAG_CALCULO)
If IsNull(par_llave!PAR_FECHA_DIA) Then
  MsgBox "URGENTE!!!. Esta Compañia No Tiene Definida la Fecha de Trabajo. Verificar!!!", 48, Pub_Titulo
  MDIForm1.StatusBar1.Panels(4).text = " V A C I O "
  LK_FECHA_DIA = #1/1/1900#
Else
  LK_FECHA_DIA = par_llave!PAR_FECHA_DIA
  MDIForm1.StatusBar1.Panels(4).text = "Fecha de Empresa : " & Format(LK_FECHA_DIA, "dd, mmmm yyyy")
End If
LK_EMP_PTO = Nulo_Valors(par_llave!PAR_PTO_VTA)
If LK_EMP = "3AA" Then
  MDIForm1.menuTit1.Caption = "&Maestros"
Else
  MDIForm1.menuTit1.Caption = "&Mantenimientos"
End If

If Format(LK_FECHA_DIA, "dd/mm/yyyy") = "30/12/1899" Then
 MsgBox "Generar el Calendario de Esta Empresa.", 48, Pub_Titulo
End If

End Sub
Public Function BAN_LINE(VAR As String) As String
Dim TEMP As String * 15
Dim N1 As Integer
Dim N2 As Integer
N1 = InStr(1, VAR, " ") - 1
N2 = Len(VAR) - N1
VAR = String(N2, "    ") + Left(VAR, N1)
BAN_LINE = VAR
End Function


Public Function BUSCA_ETIQUETA(WNUMTAB As Integer) As String
SQ_OPER = 1
PUB_TIPREG = 300
PUB_CODCIA = LK_CODCIA
PUB_NUMTAB = WNUMTAB
LEER_TAB_LLAVE
If tab_llave.EOF Then
   BUSCA_ETIQUETA = "XXXXXX"
Else
   BUSCA_ETIQUETA = Trim(tab_llave!TAB_NOMLARGO)
End If
End Function

Public Sub Permisos()
Dim W1 As String * 2
Dim i, wPosF, WPosV, cuenta As Integer
Dim wc As Integer
Dim SAL As Boolean
Dim cade As String
Dim WNUM As Integer
Dim f As Integer
Dim a As Integer
Dim wAcceso(6) As String * 40
On Error GoTo sigue
Screen.MousePointer = 11
usu.MoveFirst
Do Until usu.EOF
If Trim(usu!usu_key) = LK_CODUSU Then
    wAcceso(1) = Nulo_Valors(usu!usu_menu1)
    wAcceso(2) = Nulo_Valors(usu!usu_menu2)
    wAcceso(3) = Nulo_Valors(usu!usu_menu3)
    wAcceso(4) = Nulo_Valors(usu!usu_menu4)
    wAcceso(5) = Nulo_Valors(usu!usu_menu5)
    wAcceso(6) = Nulo_Valors(usu!usu_menu6)
    Exit Do
End If
usu.MoveNext
Loop

For wc = 1 To 6
    DoEvents
    WNUM = 0
    wPosF = 0
    WPosV = 0
    cuenta = 0
    WPosV = Len(wAcceso(wc))
    cade = Trim(wAcceso(wc))
    cuenta = 0
    wPosF = 1
    a = 0
    'If wc = 4 Then
    ' For i = 0 To 8
    '   MDIForm1.SubmenuTit1.Item(i).Enabled = False
    ' Next i
    'End If
    
    For i = 1 To Len(cade)
        If Mid(cade, i, 1) = "." Then
            a = a + 1
        End If
    Next i
    
    Do Until cuenta = a
        cuenta = cuenta + 1
        DoEvents
        wPosF = InStr(wPosF, cade, ".", 1) + 1
        DoEvents
        WNUM = Mid(cade, wPosF, 2)
        If Right(WNUM, 1) = "." Then
            WNUM = Left(WNUM, 2)
            wPosF = wPosF - 1
        End If
        Select Case wc
            Case 1
             MDIForm1.SubmenuTit1.Item(WNUM).Enabled = True
            Case 2
            MDIForm1.SubmenuTit2.Item(WNUM).Enabled = True
            Case 3
            MDIForm1.submenutit3.Item(WNUM).Enabled = True
            Case 4
            MDIForm1.SubmenuTit4.Item(WNUM).Enabled = True
            Case 5
            MDIForm1.submenutit5.Item(WNUM).Enabled = True
            Case 6
            MDIForm1.SubmenuTit6.Item(WNUM).Enabled = True
            Case 7
            'MDIForm1.SubmenuTit7.Item(WNUM).Enabled = True
            Case 8
            'MDIForm1.SubmenuTit8.Item(WNUM).Enabled = True
       End Select
    Loop
Next wc
MDIForm1.menuTit1.Enabled = True
MDIForm1.menuTit2.Enabled = True
MDIForm1.menutit3.Enabled = True
MDIForm1.menutit4.Enabled = True
MDIForm1.menutit5.Enabled = True
MDIForm1.menutit6.Enabled = True
MDIForm1.menuAyuda.Enabled = True

If InStr(1, wAcceso(1), ".0") <> 0 Then
    MDIForm1.Toolbar1.Buttons(4).Enabled = True
End If
If InStr(1, wAcceso(1), ".1.") <> 0 Then
    MDIForm1.Toolbar1.Buttons(6).Enabled = True
End If
If InStr(1, wAcceso(1), ".3") <> 0 Then
    MDIForm1.Toolbar1.Buttons(7).Enabled = True
End If
If InStr(1, wAcceso(1), ".8") <> 0 Then
    MDIForm1.Toolbar1.Buttons(2).Enabled = True
End If
If InStr(1, wAcceso(1), ".12") <> 0 Then
    MDIForm1.Toolbar1.Buttons(12).Enabled = True
End If
If InStr(1, wAcceso(1), ".15") <> 0 Then
    MDIForm1.Toolbar1.Buttons(16).Enabled = True
End If
If InStr(1, wAcceso(2), ".0") <> 0 Then
    MDIForm1.Toolbar1.Buttons(1).Enabled = True
End If
If InStr(1, wAcceso(6), ".1") <> 0 Then
    MDIForm1.Toolbar1.Buttons(9).Enabled = True
    MDIForm1.Toolbar1.Buttons(10).Enabled = True
    MDIForm1.Toolbar1.Buttons(11).Enabled = True
End If
If InStr(1, wAcceso(6), ".2.") <> 0 Or InStr(1, wAcceso(6), ".2") <> 0 Then
    MDIForm1.Toolbar1.Buttons(5).Enabled = True
End If
'MDIForm1.Toolbar1.Buttons(13).Enabled = True
MDIForm1.Toolbar1.Buttons(14).Enabled = True
MDIForm1.Toolbar1.Buttons(16).Enabled = True
Screen.MousePointer = 0
Exit Sub
sigue:
If Err.Number = 340 Then
    Resume Next
End If
End Sub
Public Sub NOTPermisos()
On Error GoTo sigue
Dim S As Integer
' Quita los permisos de Reportes
For fila = 0 To MDIForm1.menuAlm.Count - 1
 If fila <> 0 Then
   Unload MDIForm1.menuAlm(fila)
 Else
   MDIForm1.menuAlm(fila).Caption = ""
 End If
Next fila
For fila = 0 To MDIForm1.menuVent.Count - 1
 If fila <> 0 Then
  Unload MDIForm1.menuVent(fila)
 Else
   MDIForm1.menuVent(fila).Caption = ""
 End If
Next fila
For fila = 0 To MDIForm1.menuMoli.Count - 1
 If fila <> 0 Then
  Unload MDIForm1.menuMoli(fila)
 Else
   MDIForm1.menuMoli(fila).Caption = ""
 End If
Next fila
For fila = 0 To MDIForm1.MenuContab.Count - 1
 If fila <> 0 Then
  Unload MDIForm1.MenuContab(fila)
 Else
   MDIForm1.MenuContab(fila).Caption = ""
 End If
Next fila
For fila = 0 To MDIForm1.Menucomp.Count - 1
 If fila <> 0 Then
   Unload MDIForm1.Menucomp(fila)
 Else
   MDIForm1.Menucomp(fila).Caption = ""
 End If
Next fila
For fila = 0 To MDIForm1.menudis1.Count - 1
 If fila <> 0 Then
   Unload MDIForm1.menudis1(fila)
 Else
   MDIForm1.menudis1(fila).Caption = ""
 End If
Next fila
For fila = 0 To MDIForm1.menudis2.Count - 1
 If fila <> 0 Then
   Unload MDIForm1.menudis2(fila)
 Else
   MDIForm1.menudis2(fila).Caption = ""
 End If
Next fila
For fila = 0 To MDIForm1.menudis3.Count - 1
 If fila <> 0 Then
   Unload MDIForm1.menudis3(fila)
 Else
   MDIForm1.menudis3(fila).Caption = ""
 End If
Next fila

For fila = 0 To MDIForm1.SubmenuTit4.Count - 1
  If MDIForm1.SubmenuTit4(fila).Visible Then
    MDIForm1.SubmenuTit4(fila).Visible = False
 End If
Next fila

MDIForm1.Toolbar1.Buttons(1).Enabled = False
MDIForm1.Toolbar1.Buttons(2).Enabled = False
MDIForm1.Toolbar1.Buttons(4).Enabled = False
MDIForm1.Toolbar1.Buttons(5).Enabled = False
MDIForm1.Toolbar1.Buttons(6).Enabled = False
MDIForm1.Toolbar1.Buttons(7).Enabled = False
MDIForm1.Toolbar1.Buttons(9).Enabled = False
MDIForm1.Toolbar1.Buttons(10).Enabled = False
MDIForm1.Toolbar1.Buttons(11).Enabled = False
MDIForm1.Toolbar1.Buttons(12).Enabled = False
MDIForm1.Toolbar1.Buttons(13).Enabled = False
MDIForm1.Toolbar1.Buttons(14).Enabled = False
MDIForm1.Toolbar1.Buttons(16).Enabled = False

MDIForm1.menuAyuda.Enabled = False
MDIForm1.menuTit1.Enabled = False
MDIForm1.menuTit2.Enabled = False
MDIForm1.menutit3.Enabled = False
MDIForm1.menutit4.Enabled = False
MDIForm1.menutit5.Enabled = False
MDIForm1.menutit6.Enabled = False

For S = 0 To MDIForm1.SubmenuTit1.Count - 1
    If Not MDIForm1.SubmenuTit1.Item(S).Caption = "-" Then
        MDIForm1.SubmenuTit1.Item(S).Enabled = False
    End If
Next S
For S = 0 To MDIForm1.SubmenuTit2.Count - 1
   If Not MDIForm1.SubmenuTit2.Item(S).Caption = "-" Then
    MDIForm1.SubmenuTit2.Item(S).Enabled = False
   End If
Next S
For S = 0 To MDIForm1.submenutit3.Count - 1
   If Not MDIForm1.submenutit3.Item(S).Caption = "-" Then
    MDIForm1.submenutit3.Item(S).Enabled = False
   End If
Next S
For S = 0 To MDIForm1.SubmenuTit4.Count - 1
   If Not MDIForm1.SubmenuTit4.Item(S).Caption = "-" Then
    MDIForm1.SubmenuTit4.Item(S).Enabled = False
   End If
Next S
For S = 0 To MDIForm1.submenutit5.Count - 1
   If Not MDIForm1.submenutit5.Item(S).Caption = "-" Then
    MDIForm1.submenutit5.Item(S).Enabled = False
   End If
Next S
For S = 0 To MDIForm1.SubmenuTit6.Count - 1
   If Not MDIForm1.SubmenuTit6.Item(S).Caption = "-" Then
    MDIForm1.SubmenuTit6.Item(S).Enabled = False
   End If
Next S


Exit Sub
sigue:
'MsgBox Err.Description
Resume Next

End Sub


Public Sub LLENA_LISTRANSA(wlista As ListBox)
Dim SN As String
 Dim i As Integer
 Dim J As Integer
    lis_tra.Requery
    wlista.Clear
    wlista.Visible = False
    Do Until lis_tra.EOF
     If gen!gen_bloqueo = "A" And lis_tra!TRA_KEY = 1409 Then GoTo PASA:
        SN = "N"
        i = 0
        Do Until SN = "S" Or i = 10
            i = i + 1
            J = 0
            Do Until SN = "S" Or J = 10
                J = J + 1
                If lk_GRUPOS(J) = lis_tra(92 + i) And lk_GRUPOS(J) <> 0 Then
                    SN = "S"
                 End If
            Loop
        Loop
        If SN = "N" And LK_CODUSU <> "ADMIN" Then
            J = 1
            Do Until lk_CODTRAS(J) = "" Or SN = "Y" Or J = 10
               If Left(lk_CODTRAS(J), 4) = lis_tra(0) Then
                  SN = "Y"
                  Exit Do
                End If
                J = J + 1
            Loop
            If SN = "N" And LK_CODUSU <> "ADMIN" Then
            Else
              wlista.AddItem lis_tra!TRA_KEY & "   " & UCase(lis_tra!tra_descripcion)
            End If
        Else
            wlista.AddItem lis_tra!TRA_KEY & "   " & UCase(lis_tra!tra_descripcion)
        End If
PASA:
        lis_tra.MoveNext
    Loop
    wlista.Visible = True
End Sub

