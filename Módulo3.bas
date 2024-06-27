Attribute VB_Name = "Módulo3"
Option Explicit
Public FILAX As Integer
Public LKCHEK As Boolean
Public PUB_TIPDES As Integer
Public PUB_LISDES As Integer
Public LK_FLAG_CAMBIAR As String * 1
Public LK_DIRECCION As String
Public LK_EMP_PTO As String * 1
Public PUB_CLAVE As String
Public LK_TIPO_CAMBIO As Double
Public LK_MONEDA As String * 1
Public PUB_VISITA As Integer
Public pub_flag_cambio As Integer
Public LK_USU_STOCK As String * 1
Public LK_USU_CUENTA As String * 1
Public LK_USU_IMPRIME As String * 1
Public LK_USU_PASSWORD As String * 1
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
Public CONST_CIACENTRAL As String
Public PUB_ODBC As String
Public PUB_LINEAS As Integer
Public PUB_FLAG As Integer
Public PUB_NUM As Long
Public pub_signo_ccm As Integer
Public pub_signo_car As Integer
Public pub_signo_caja As Integer
Public pub_signo_arm As Integer
Public pub_signo_ped As Integer
Public PUB_SERGUIA As Integer
Public pub_ojo As String * 1
Public PUB_FBG As String * 1
Public pasa As Integer
Public TABLA_TAG(1000)
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
Public PUB_ISLA  As Integer
Public PUB_TURNO As Integer


Public PUB_CODCONT As String * 12

Public PUB_DS As String * 1
Public PUB_NOMBRE_BANCO As String * 30
Public PUB_NUM_CHEQUE As String * 12

Public PUB_CODBAN As Integer 'Integer
Public PUB_CONCEPTO As String
Public PUB_FLAG_VENCIDO As Integer
Public PUB_FECHA As Date
Public PUB_FECHA1 As Date ' Modificado agregado para Reporte de Caja 20042004
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
Public PUB_CODCHOFER As Integer 'AGREGADO PARA CHOFER POR MIC
Public PUB_CODVEN As Integer
Public PUB_KEY As Long
Public PUB_CODTRA As Integer
Public PUB_IS As String * 1
Public PUB_VF As Boolean
Public PUB_CODIGO As Long
Public PUB_NUM_OPER As Integer
Public PUB_NUM_OPER_EXT As Integer
Public PUB_NUM_OPER_XXX As Integer

Public PUB_NUMTAB As Currency
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
Public pu_alterno As String * 50
Public LK_EMP As String * 3
Public LK_FLAG_CALCULO As String * 1
Public LK_FAC_IMP As String * 1
Public LK_CONTROLA_STOCK As String * 1
Public LK_PASA_BOLETAS As String * 1


Public LK_DIG_RUC As Integer

Public WCABE As Integer

Public NUM_CONTAB(99) As Currency

Public TEXTOX(20) As String
Public NOMBREX(20) As String
Public ETIQUETAX(20) As String
Public WS_IMPORTE As Currency
Public WS_NETO As Currency
Public WS_DESCTO As Currency
Public WS_IMPTO As Currency
Public ws_igv As Currency
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
Public PU_FBG2 As String * 1
Public pu_fecha As Date
Public PUB_SO As String * 1

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
Public fila As Long
Public WStop As Boolean
Declare Function SetWindowPos Lib "User32" (ByVal h&, ByVal hb&, ByVal x&, ByVal y&, ByVal cx&, ByVal cy&, ByVal f&) As Long
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
Public tab_precioss(6) As String * 7
Public LK_DEVICE_FBG As String * 3
Public LK_FLAG_GRIFO As String * 1
Public LK_FLAG_EXED As String * 1
Public LK_FLAG_SOS As String * 1
Public LK_ACTIVA As String * 1
Public PUB_ACU_TOTAL As Currency
Public PUB_FLAG_VENCIDO_VISTA As Integer
Public pub_mensaje_err As String
Public ww_respuesta  As Integer
Public loc_doucumento As String
Public par_MOV_NRO_MOV As Integer

'VARIABLES DE RELACION CONTABLE ONLYONT
   Public PSMOV_LLAVE As rdoQuery
   Public mov_llave As rdoResultset
      Public par_MOV_CODCIA As String * 2
   Public par_MOV_NRO_VOUCHER As Integer
   Public par_MOV_TIPMOV As Integer
   Public par_MOV_FECHA As Date
   Public par_MOV_GLOSA As String
   Public par_MOV_MONEDA As String * 1
   Public par_MOV_CODCTA As String * 12
   Public par_MOV_DH As String * 1
   Public par_MOV_IMPORTE As Currency
   Public par_MOV_DETALLE As String
   Public par_MOV_nro_MES As Integer
   Public par_MOV_fecha_EMI As Date
   Public par_MOV_PLANTILLA As Integer
   Public par_MOV_FLAG_DES As String * 1
   Public PUB_FLAG_VENCIDO_CC As Integer
   Public PUB_FLAG_DOC As Integer
   Public PSMANO_CODI As rdoQuery
   Public mano_CODI As rdoResultset
Public PUB_TIPOPRINT As Integer
   
   Public Sub CentrarFormulario(objFpadre As Form, objFhijo As Form)
    objFhijo.Left = (objFpadre.ScaleWidth - objFhijo.Width) / 2
    objFhijo.Top = (objFpadre.ScaleHeight - objFhijo.Height) / 2
   End Sub

Public Sub pasa_def()

PUB_CP = Nulo_Valors(SUT_LLAVE!SUT_cp)
pub_signo_ccm = SUT_LLAVE!SUT_signo_ccm
pub_signo_car = Nulo_Valor0(SUT_LLAVE!SUT_SIGNO_CAR)
pub_signo_arm = Nulo_Valor0(SUT_LLAVE!SUT_SIGNO_ARM)
pub_signo_caja = Nulo_Valor0(SUT_LLAVE!SUT_signo_caja)

PUB_TIPMOV = Nulo_Valor0(SUT_LLAVE!SUT_TIPMOV)
PUB_TIPMOV_REF = Nulo_Valor0(SUT_LLAVE!SUT_TIPMOV_REF)
PUB_TIPDOC = Nulo_Valors(SUT_LLAVE!sut_TIPDOC)
PUB_ABREVIADO = Nulo_Valors(SUT_LLAVE!SUT_abreviado)
PUB_CODPRO = Nulo_Valor0(SUT_LLAVE!SUT_codpro)
PUB_CODALI = PUB_ABREVIADO

If PUB_CODPRO > 0 Then
   SQ_OPER = 1
   PUB_TIPREG = 888
   PUB_NUMTAB = PUB_CODPRO
   PUB_CODCIA = LK_CODCIA
   LEER_TAB_LLAVE
   If Not tab_llave.EOF Then PUB_CODALI = tab_llave!tab_nomcorto
End If


End Sub
Public Sub ACT_FORMGEN()
 Dim i
 Dim c As Integer
 Dim TIPREG
 '''*****TEMPORAL *******''
    PUB_TIPREG = 999
    'PUB_TIPREG = 777
    PUB_CODCIA = "00"
    SQ_OPER = 2
    LEER_TAB_LLAVE
     Do Until tab_mayor.EOF
        tab_mayor.Delete
        tab_mayor.MoveNext
     Loop
 FORMGEN.Print " "
 FORMGEN.Print " "
 FORMGEN.Print "Actualizando...."
 c = 0
 For i = 0 To FORMGEN.Controls.count - 1
     FORMGEN.Print "                " & FORMGEN.Controls(i).Name
       If Not (FORMGEN.Controls(i).Tag = 0 Or FORMGEN.Controls(i).Tag > 400) Then
'   MsgBox FORMGEN.Controls(i).Name
        tab_mayor.AddNew
        tab_mayor!TAB_CODCIA = "00"
        tab_mayor!TAB_TIPREG = PUB_TIPREG
        tab_mayor!TAB_NUMTAB = FORMGEN.Controls(i).Tag
        
        tab_mayor!tab_nomcorto = val(i)
        tab_mayor!tab_NOMLARGO = Trim(FORMGEN.Controls(i).Name)
        tab_mayor.Update
        c = c + 1
      End If
 Next i
FORMGEN.Cls
FORMGEN.Print " "
FORMGEN.Print " "
FORMGEN.Print " Proceso Terminado..."
MsgBox "Cantidad de Controles Actualizados....  : " & c & "  O K !!", 48, Pub_Titulo
FORMGEN.Cls

End Sub
Public Function CONVER_LETRAS(NUMERO_BASE As Currency, WMONEDA As String) As String
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
Dim c As Integer
Dim DU As Integer
Dim D As Integer
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
If WMONEDA = "S" Then
  WS_DEC = "y " & Format(WS_DEC, "00") & "/100 SOLES"
Else
  WS_DEC = "y " & Format(WS_DEC, "00") & "/100 DOLARES AMERICANOS"
End If
VECTOR(1) = Mid(NUMERO, 1, 3)
VECTOR(2) = Mid(NUMERO, 4, 3)
VECTOR(3) = Mid(NUMERO, 7, 3)
VECTOR(4) = Mid(NUMERO, 10, 3)
pub_cadena = ""
For i = 1 To 4
    t = 0
    N = 1
    cdu = val(VECTOR(i))
    c = Int(cdu / 100)
    DU = cdu - (c * 100)
    D = Int(DU / 10)
    u = DU - (D * 10)
    If cdu > 99 Then
        wa = Trim(CENTEN(c))
        pub_cadena = pub_cadena + wa
        If c > 1 Then
            If c = 5 Then
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
            wa = Trim(DECEN(D))
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
If val(valor) < 0 Then
 temp1 = Format(valor, "##,###,##0.00")
 temp1 = Mid(temp1, 2, Len(temp1))
 temp1 = "(" & temp1 & ")"
 NUM_NEGA = temp1
Else
 NUM_NEGA = Format(valor, "##,###,##0.00")
 NUM_NEGA = NUM_NEGA + " "
End If
End Function
Public Sub PROC_LISVIEW(LV1 As Object, Optional wmax) ' ListView
On Error GoTo SALE
Dim wmaximo As Integer
Dim itmX As Object 'ListItem
If Not IsMissing(wmax) Then wmaximo = wmax Else wmaximo = 3000
Set PSX = CN.CreateQuery("", archi)
Set x = PSX.OpenResultset(rdOpenForwardOnly)
x.Requery
LV1.ListItems.Clear
LV1.ColumnHeaders.Clear
If x.EOF Then LV1.Visible = False: Exit Sub
LV1.Top = 3000
'LV1.Left = 3300
'LV1.Width = 6500
LV1.Height = 4000
LV1.Width = 12000
LV1.Left = 100
'LV1.Visible = True
If numarchi = 3 Then ' para codigos alternos
 LV1.ColumnHeaders.Add 1, , "Alterno", 980
 LV1.ColumnHeaders.Add 2, , "Descripción", 4000
 LV1.ColumnHeaders.Add 3, , "Original", 0
 ''If LK_EMP = "3AA" Then LV1.ColumnHeaders.Add 4, , "Stock", 1000
 LV1.ColumnHeaders.Add 4, , "Stock", 1000
 LV1.ColumnHeaders.Add 5, , "P.Menor", 1000
 LV1.ColumnHeaders.Add 6, , "P.Mayor", 1000
 LV1.ColumnHeaders.Add 7, , "Marca", 0
 LV1.ColumnHeaders.Add 8, , "Color", 0
  
ElseIf numarchi = 1 Then
 LV1.ColumnHeaders.Add 1, , "Descripción", 3800
 LV1.Width = 11500
 LV1.Left = 300
 LV1.ColumnHeaders.Add 2, , "Cod.", 600
 LV1.ColumnHeaders.Add 3, , "Dirección", 3800
 LV1.ColumnHeaders.Add 4, , "Zona", 1500
ElseIf numarchi = 0 Then
 LV1.ColumnHeaders.Add 1, , "Descripción", 4200
 LV1.ColumnHeaders.Add 2, , "Cod.", 0
 LV1.ColumnHeaders.Add 3, , "Codigo", 1000
 LV1.ColumnHeaders.Add 4, , "P.Menor", 900
 LV1.ColumnHeaders.Add 5, , "P.Mayor", 900
 LV1.ColumnHeaders.Add 6, , "Stock", 900
 LV1.ColumnHeaders.Add 7, , "Marca", 0
 LV1.ColumnHeaders.Add 8, , "Color", 0
 
Else
 LV1.ColumnHeaders.Add 1, , "Descripción", 4200
 LV1.ColumnHeaders.Add 2, , "Cod.", 1000
End If
Do Until x.EOF Or x.AbsolutePosition - 1 >= wmaximo
   If numarchi = 1 Or numarchi = 3 Then Set itmX = LV1.ListItems.Add(, , (Trim(CStr(x.rdoColumns(3))))) Else: Set itmX = LV1.ListItems.Add(, , Trim(CStr(x.rdoColumns(2))))
   If numarchi = 3 Then itmX.SubItems(1) = Trim(CStr(x.rdoColumns(2))) Else: itmX.SubItems(1) = Trim(CStr(x.rdoColumns(0)))
   If numarchi = 0 Or numarchi = 3 Then
    '' itmX.SubItems(2) = Trim(CStr(X.rdoColumns(0)))
      If numarchi = 3 Then
        itmX.SubItems(3) = Format(x.rdoColumns(4) / x.rdoColumns(5), "0.000")
        itmX.SubItems(4) = Trim(CStr(x.rdoColumns(6)))
        itmX.SubItems(5) = Trim(CStr(x.rdoColumns(7)))
        itmX.SubItems(6) = Trim(CStr(x.rdoColumns(8)))  'marca
        itmX.SubItems(7) = Trim(CStr(x.rdoColumns(9)))  'color
      Else
        'itmX.SubItems(2) = Trim(X.rdoColumns(6))
        'itmX.SubItems(3) = Trim(X.rdoColumns(7))
        'itmX.SubItems(4) = Format(X.rdoColumns(4) / X.rdoColumns(5), "0.000")
        'itmX.SubItems(5) = Trim(X.rdoColumns(3))
        ' Inicio Modificado
        itmX.SubItems(2) = Trim(x.rdoColumns(3))
        itmX.SubItems(3) = Trim(x.rdoColumns(8))
        itmX.SubItems(4) = Trim(x.rdoColumns(9))
        itmX.SubItems(5) = Format(x.rdoColumns(4) / x.rdoColumns(5), "0.000")
       ' itmX.SubItems(6) = Trim(X.rdoColumns(11))
       ' itmX.SubItems(7) = Trim(X.rdoColumns(10))
        
        
        
      End If
   End If
   If numarchi = 1 Then itmX.SubItems(2) = Trim(CStr(Nulo_Valors(x.rdoColumns(4)))) + " # " + Trim(CStr(x.rdoColumns(6))):  itmX.SubItems(3) = Trim(CStr(x!tab_NOMLARGO))

   itmX.Tag = x.AbsolutePosition
   x.MoveNext
Loop
LV1.ToolTipText = "Encontrados : " & itmX.Tag & "/" & x.RowCount & " Muestra un Maximo de: " & wmaximo
LV1.Visible = True
'DoEvents
Exit Sub
SALE:
MsgBox Err.Description
'Resume Next
If Err.Number = 40002 Then Exit Sub Else MsgBox Err.Description, 48, Pub_Titulo

End Sub

Public Sub BORRA_FIELDS(NUM As Integer, ARRAYY As Variant)
Dim x As Integer
For x = 1 To NUM
 ARRAYY(x).Text = ""
Next

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
   MDIForm1.Toolbar1.Buttons.Item(14).Enabled = False
 End If
 If LK_FLAG_ORIGINAL <> "A" Then
  MDIForm1.Toolbar1.Buttons.Item(14).Enabled = True
  If LK_FLAG_ALTERNO = "A" And LK_FLAG_ORIGINAL <> "A" Then
   MDIForm1.Toolbar1.Buttons.Item(12).ToolTipText = "Busqueda sub-codigo de articulo."
   DoEvents
   MDIForm1.Toolbar1.Buttons.Item(12).Image = 5
   DoEvents
  Else
   MDIForm1.Toolbar1.Buttons.Item(12).ToolTipText = "Busqueda nombre de articulo."
   DoEvents
   MDIForm1.Toolbar1.Buttons.Item(12).Image = 18
   DoEvents
  End If
 End If
LK_EMP = Nulo_Valors(par_llave!PAR_EMPRESA)
LK_MONEDA = Nulo_Valors(par_llave!PAR_MONEDA_FAC)
If Trim(LK_MONEDA) = "" Then MsgBox "DEFINIR LA MONEDA DE COMPAÑIA", 48, Pub_Titulo
LK_FLAG_FACTURACION = Nulo_Valors(par_llave!PAR_FLAG_FACTURACION)
LK_FLAG_CALCULO = Nulo_Valors(par_llave!PAR_FLAG_CALCULO)
LK_FLAG_EXED = Nulo_Valors(par_llave!PAR_FLAG_EXED)

If IsNull(par_llave!PAR_FECHA_DIA) Then
  MsgBox "URGENTE!!!. Esta Compañia No Tiene Definida la Fecha de Trabajo. Verificar!!!", 48, Pub_Titulo
  MDIForm1.stb_EB.Panels("date").Text = Date
  LK_FECHA_DIA = Date
Else
  LK_FECHA_DIA = par_llave!PAR_FECHA_DIA
  MDIForm1.stb_EB.Panels("date").Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
End If
LK_EMP_PTO = Nulo_Valors(par_llave!PAR_PTO_VTA)
LK_FLAG_GRIFO = Nulo_Valors(par_llave!PAR_FLAG_GRIFO)
If LK_FLAG_GRIFO = "A" Then
  MDIForm1.menugrifo.Visible = True
  MDIForm1.menugrifo.Enabled = True
Else
  MDIForm1.menugrifo.Visible = False
  MDIForm1.menugrifo.Enabled = False
End If

If Format(LK_FECHA_DIA, "dd/mm/yyyy") = "30/12/1899" Then
 MsgBox "Generar el Calendario de Esta Empresa.", 48, Pub_Titulo
End If
LK_CAJERO = ""
If LK_EMP = "3AA" Then
  MDIForm1.menuTit1.Caption = "&Maestros"
  LK_CAJERO = ""
Else
  MDIForm1.menuTit1.Caption = "&Mantenimientos"
End If
PSCOP_LLAVE.rdoParameters(0) = par_llave!PAR_CIACON
cop_llave.Requery
LK_NIVEL_MAX = 3
If cop_llave.EOF Then
Else
 LK_NIVEL_MAX = cop_llave!cop_nivel_max
End If
'FrmParGen.txtcolor.Text =
On Error GoTo ecolor
If val(Nulo_Valors(par_llave!PAR_COLOR)) <> 0 Then
 MDIForm1.lblciagen.BackColor = QBColor(Left(Nulo_Valors(par_llave!PAR_COLOR), 2))
 MDIForm1.lblciagen.ForeColor = QBColor(Mid(Nulo_Valors(par_llave!PAR_COLOR), 3, 4))
Else
 MDIForm1.lblciagen.BackColor = QBColor(15)
 MDIForm1.lblciagen.ForeColor = QBColor(0)
End If
cop_llave.Requery

SQ_OPER = 2
PUB_CODCIA = LK_CODCIA
LEER_PAR_LLAVE
If pac_llave.EOF Then
  
End If

Exit Sub
ecolor:
' MsgBox "Verificar los codigos de color en Solution", 48, Pub_Titulo
End Sub
Public Function CAR_TOT_CPX(WCP As String, wcodcia As String, wcodclie As Double) As Double
Dim SUM_IMPORTE As Double
    PUB_FLAG_VENCIDO = 0
    SUM_IMPORTE = 0
    SQ_OPER = 2
    pu_codcia = wcodcia
    pu_codclie = wcodclie
    pu_cp = WCP
    pub_cadena = "Guias: "
    LEER_CAR_LLAVE
    Do Until car_mayor.EOF
       If car_mayor!car_importe > 0 Then
          If DateDiff("d", car_mayor!car_fecha_vcto, LK_FECHA_DIA) > Nulo_Valor0(par_llave!PAR_DIAS_VENC) Then PUB_FLAG_VENCIDO = 1
       End If
       SUM_IMPORTE = SUM_IMPORTE + (car_mayor!car_importe)
       car_mayor.MoveNext
    Loop
    
    CAR_TOT_CPX = SUM_IMPORTE
End Function
Public Function CAR_TOT_CPX2(WCP As String, wcodcia As String, wcodclie As Double) As Double
Dim SUM_IMPORTE As Double
Dim SUM_CC As Double
    SUM_CC = 0
    PUB_FLAG_DOC = 0
    PUB_FLAG_VENCIDO = 0
    PUB_FLAG_VENCIDO_VISTA = 0
    SUM_IMPORTE = 0
    SQ_OPER = 2
    pu_codcia = wcodcia
    pu_codclie = wcodclie
    pu_cp = WCP
    pub_cadena = "Guias: "
    pub_mensaje = ""
    LEER_CAR_LLAVE
    Do Until car_mayor.EOF
       If LK_FLAG_SOS = "A" And car_mayor!CAR_FLAG_SO <> "A" Then GoTo pasa
       If car_mayor!car_importe > 0 Then
          If DateDiff("d", car_mayor!car_fecha_vcto_orig, LK_FECHA_DIA) > Nulo_Valor0(par_llave!PAR_DIAS_VENC) Then PUB_FLAG_VENCIDO = 1
          If DateDiff("d", car_mayor!car_fecha_vcto_orig, LK_FECHA_DIA) >= 0 Or DateDiff("d", car_mayor!car_fecha_vcto_orig, LK_FECHA_DIA) < Nulo_Valor0(par_llave!PAR_DIAS_VENC) Then PUB_FLAG_VENCIDO_VISTA = 1
       End If
       If car_mayor!car_TIPDOC <> "CH" Then
        If car_mayor!car_importe > 0 Then
             PUB_FLAG_DOC = PUB_FLAG_DOC + 1
        End If
        If car_mayor!car_TIPDOC <> "CC" Then

        If car_mayor!CAR_MONEDA = "D" Then
            PUB_CAL_INI = car_mayor!CAR_FECHA_SUNAT
            PUB_CAL_FIN = car_mayor!CAR_FECHA_SUNAT
            PUB_CODCIA = LK_CODCIA
            SQ_OPER = 1
            LEER_CAL_LLAVE
            If Not cal_llave.EOF Then
              SUM_IMPORTE = SUM_IMPORTE + redondea(car_mayor!car_importe * Nulo_Valor0(cal_llave!cal_tipo_cambio))
            Else
              SUM_IMPORTE = SUM_IMPORTE + (car_mayor!car_importe)
            End If
         Else
            SUM_IMPORTE = SUM_IMPORTE + (car_mayor!car_importe)
         End If
        Else
           SUM_CC = SUM_CC + (car_mayor!car_importe)
        End If
        End If
        pub_mensaje = pub_mensaje & car_mayor!car_TIPDOC & "   - " & Nulo_Valors(car_mayor!car_FBG) & "/. " & car_mayor!car_NUMSER & " - " & car_mayor!car_NUMFAC & "       F.Vcto: " & car_mayor!car_fecha_vcto & "        " & car_mayor!CAR_MONEDA & "/." & Format(car_mayor!car_importe, "0.00") & "     -V: " & car_mayor!CAR_codven & Chr(13) & Chr(13)
pasa:
       car_mayor.MoveNext
    Loop
    If PUB_FLAG_VENCIDO = 1 Then PUB_FLAG_VENCIDO_VISTA = 0
    CAR_TOT_CPX2 = SUM_IMPORTE
    pub_mensaje = pub_mensaje & "TOTAL CREDITO =  " & SUM_IMPORTE & Chr(13) & "TOTAL CONTADO REPARTO =  " & SUM_CC
    
End Function
Public Function CAR_TOT_CC(WCP As String, wcodcia As String, wcodclie As Double) As Double
Dim SUM_IMPORTE As Double
    PUB_FLAG_VENCIDO_CC = 0
    PUB_FLAG_VENCIDO = 0
    PUB_FLAG_VENCIDO_VISTA = 0
    SUM_IMPORTE = 0
    SQ_OPER = 2
    pu_codcia = wcodcia
    pu_codclie = wcodclie
    pu_cp = WCP
    pub_cadena = "Guias: "
    pub_mensaje = ""
    LEER_CAR_LLAVE
    Do Until car_mayor.EOF
       If car_mayor!car_TIPDOC <> "CC" Then GoTo pasa
         If car_mayor!CAR_MONEDA = "D" Then
            PUB_CAL_INI = car_mayor!CAR_FECHA_SUNAT
            PUB_CAL_FIN = car_mayor!CAR_FECHA_SUNAT
            PUB_CODCIA = LK_CODCIA
            SQ_OPER = 1
            LEER_CAL_LLAVE
            If Not cal_llave.EOF Then
              SUM_IMPORTE = SUM_IMPORTE + redondea(car_mayor!car_importe * Nulo_Valor0(cal_llave!cal_tipo_cambio))
            Else
              SUM_IMPORTE = SUM_IMPORTE + (car_mayor!car_importe)
            End If
         Else
            SUM_IMPORTE = SUM_IMPORTE + (car_mayor!car_importe)
         End If
       If car_mayor!car_importe > 0 Then
          If DateDiff("d", car_mayor!car_fecha_vcto, LK_FECHA_DIA) > Nulo_Valor0(par_llave!PAR_DIAS_VENC_CC) Then PUB_FLAG_VENCIDO_CC = 1
       End If
       pub_mensaje = pub_mensaje & car_mayor!car_TIPDOC & "   - " & Nulo_Valors(car_mayor!car_FBG) & "/. " & car_mayor!car_NUMSER & " - " & car_mayor!car_NUMFAC & "       F.Vcto: " & car_mayor!car_fecha_vcto & "        " & car_mayor!CAR_MONEDA & "/." & Format(car_mayor!car_importe, "0.00") & "     -V: " & car_mayor!CAR_codven & Chr(13) & Chr(13)
       
pasa:
       car_mayor.MoveNext
    Loop
    CAR_TOT_CC = SUM_IMPORTE
    pub_mensaje = pub_mensaje & "TOTAL  =  " & SUM_IMPORTE
    
End Function

Public Function BAN_LINE(VAR As String) As String
Dim TEMP As String * 15
Dim N1 As Integer
Dim N2 As Integer
N1 = InStr(1, VAR, " ") - 1
N2 = Len(VAR) - N1
VAR = String(N2, "    ") + Left(VAR, N1)
BAN_LINE = VAR
End Function


Public Sub LLENA_LISTRANSA(wlista As ListBox, wtra As Integer)
 Dim SN As String
 Dim i As Integer
 Dim j As Integer
    lis_tra.Requery
    wlista.Clear
    wlista.Visible = False
    Do Until lis_tra.EOF
        If lis_tra!TRA_KEY = 1409 And LK_FLAG_SOS = "A" Then GoTo otro2
        If wtra = 1 Then
          If lis_tra!TRA_KEY = 2107 Or lis_tra!TRA_KEY = 2105 Or lis_tra!TRA_KEY = 2101 Then GoTo otro2
        ElseIf wtra = 2 Then
           If lis_tra!TRA_KEY <> 2101 And lis_tra!TRA_KEY <> 2105 And lis_tra!TRA_KEY <> 2107 Then GoTo otro2
        End If
       ' If LK_FLAG_GRIFO = "A" Then
       '     If lis_tra!tra_key <> 2101 And lis_tra!tra_key <> 2105 Then GoTo otro2
       ' Else
       '     If lis_tra!tra_key = 2107 Then GoTo otro2
       ' End If
        SN = "N"
        i = 0
        Do Until SN = "S" Or i = 10
            i = i + 1
            j = 0
            Do Until SN = "S" Or j = 10
                j = j + 1
                If lk_GRUPOS(j) = lis_tra(92 + i) And lk_GRUPOS(j) <> 0 Then
                    SN = "S"
                 End If
            Loop
        Loop
        If SN = "N" And LK_CODUSU <> "ADMIN" Then
            j = 1
            Do Until lk_CODTRAS(j) = "" Or SN = "Y" Or j = 10
               If Left(lk_CODTRAS(j), 4) = lis_tra(0) Then
                  SN = "Y"
                  Exit Do
                End If
                j = j + 1
            Loop
            If SN = "N" And LK_CODUSU <> "ADMIN" Then
            Else
              wlista.AddItem lis_tra!TRA_KEY & "   " & UCase(lis_tra!TRA_DESCRIPCION)
            End If
        Else
            wlista.AddItem lis_tra!TRA_KEY & "   " & UCase(lis_tra!TRA_DESCRIPCION)
        End If
otro2:
        lis_tra.MoveNext
    Loop
    wlista.Visible = True
End Sub

Public Function BUSCA_ETIQUETA(WNUMTAB As Integer) As String
SQ_OPER = 1
PUB_TIPREG = 300
PUB_CODCIA = LK_CODCIA
PUB_NUMTAB = WNUMTAB
LEER_TAB_LLAVE
If tab_llave.EOF Then
   BUSCA_ETIQUETA = "XXXXXX"
Else
   BUSCA_ETIQUETA = Trim(tab_llave!tab_NOMLARGO)
End If
End Function

Public Sub Permisos()
Dim W1 As String * 2
Dim i, wPosF, WPosV, cuenta As Integer
Dim WC As Integer
Dim sal As Boolean
Dim cade As String
Dim WNUM As Integer
Dim f As Integer
Dim a As Integer
Dim wAcceso(7) As String * 40
On Error GoTo sigue
Screen.MousePointer = 11
usu.MoveFirst
Do Until usu.EOF
If Trim(usu!USU_KEY) = LK_CODUSU Then
    wAcceso(1) = Nulo_Valors(usu!usu_menu1)
    wAcceso(2) = Nulo_Valors(usu!usu_menu2)
    wAcceso(3) = Nulo_Valors(usu!usu_menu3)
    wAcceso(4) = Nulo_Valors(usu!usu_menu4)
    wAcceso(5) = Nulo_Valors(usu!usu_menu5)
    wAcceso(6) = Nulo_Valors(usu!usu_menu6)
    wAcceso(7) = Nulo_Valors(usu!usu_menu7)
    Exit Do
End If
usu.MoveNext
Loop

For WC = 1 To 7
    DoEvents
    WNUM = 0
    wPosF = 0
    WPosV = 0
    cuenta = 0
    WPosV = Len(wAcceso(WC))
    cade = Trim(wAcceso(WC))
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
        Select Case WC
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
            MDIForm1.SubmenuTit7.Item(WNUM).Enabled = True
            Case 8
            'MDIForm1.SubmenuTit8.Item(WNUM).Enabled = True
       End Select
    Loop
Next WC
MDIForm1.menuTit1.Enabled = True
MDIForm1.menuTit2.Enabled = True
MDIForm1.menutit3.Enabled = True
MDIForm1.menutit4.Enabled = True
MDIForm1.menutit5.Enabled = True
MDIForm1.menutit6.Enabled = True
MDIForm1.menutit7.Enabled = True
MDIForm1.menuAyuda.Enabled = True

If InStr(1, wAcceso(1), ".0") <> 0 Then
    MDIForm1.Toolbar1.Buttons(3).Enabled = True
End If
If InStr(1, wAcceso(1), ".1.") <> 0 Then
    MDIForm1.Toolbar1.Buttons(5).Enabled = True
End If
If InStr(1, wAcceso(1), ".3") <> 0 Then
    MDIForm1.Toolbar1.Buttons(6).Enabled = True
End If
If InStr(1, wAcceso(1), ".12") <> 0 Then
    MDIForm1.Toolbar1.Buttons(10).Enabled = True
End If
If InStr(1, wAcceso(2), ".0") <> 0 Then
    MDIForm1.Toolbar1.Buttons(1).Enabled = True
End If
If InStr(1, wAcceso(6), ".0") <> 0 Then
    MDIForm1.Toolbar1.Buttons(7).Enabled = True
End If
If InStr(1, wAcceso(6), ".1") <> 0 Then
    MDIForm1.Toolbar1.Buttons(9).Enabled = True
End If
If InStr(1, wAcceso(6), ".2") <> 0 Then
    MDIForm1.Toolbar1.Buttons(8).Enabled = True
End If
If InStr(1, wAcceso(6), ".4") <> 0 Then
    MDIForm1.Toolbar1.Buttons(11).Enabled = True
End If
If InStr(1, wAcceso(6), ".6") <> 0 Then
    MDIForm1.Toolbar1.Buttons(4).Enabled = True
End If
MDIForm1.Toolbar1.Buttons(12).Enabled = True
MDIForm1.Toolbar1.Buttons(13).Enabled = True
MDIForm1.Toolbar1.Buttons(14).Enabled = True
MDIForm1.Toolbar1.Buttons(15).Enabled = True
MDIForm1.Toolbar1.Buttons(16).Enabled = True
MDIForm1.Toolbar1.Buttons(18).Enabled = True

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
For fila = 0 To MDIForm1.menuAlm.count - 1
 If fila <> 0 Then
   Unload MDIForm1.menuAlm(fila)
 Else
   MDIForm1.menuAlm(fila).Caption = ""
 End If
Next fila
For fila = 0 To MDIForm1.menuVent.count - 1
 If fila <> 0 Then
  Unload MDIForm1.menuVent(fila)
 Else
   MDIForm1.menuVent(fila).Caption = ""
 End If
Next fila
For fila = 0 To MDIForm1.menuMoli.count - 1
 If fila <> 0 Then
  Unload MDIForm1.menuMoli(fila)
 Else
   MDIForm1.menuMoli(fila).Caption = ""
 End If
Next fila
For fila = 0 To MDIForm1.MenuContab.count - 1
 If fila <> 0 Then
  Unload MDIForm1.MenuContab(fila)
 Else
   MDIForm1.MenuContab(fila).Caption = ""
 End If
Next fila
For fila = 0 To MDIForm1.Menucomp.count - 1
 If fila <> 0 Then
   Unload MDIForm1.Menucomp(fila)
 Else
   MDIForm1.Menucomp(fila).Caption = ""
 End If
Next fila
For fila = 0 To MDIForm1.menudis1.count - 1
 If fila <> 0 Then
   Unload MDIForm1.menudis1(fila)
 Else
   MDIForm1.menudis1(fila).Caption = ""
 End If
Next fila
For fila = 0 To MDIForm1.menudis2.count - 1
 If fila <> 0 Then
   Unload MDIForm1.menudis2(fila)
 Else
   MDIForm1.menudis2(fila).Caption = ""
 End If
Next fila
For fila = 0 To MDIForm1.menudis3.count - 1
 If fila <> 0 Then
   Unload MDIForm1.menudis3(fila)
 Else
   MDIForm1.menudis3(fila).Caption = ""
 End If
Next fila

For fila = 0 To MDIForm1.SubmenuTit4.count - 1
  If MDIForm1.SubmenuTit4(fila).Visible Then
    MDIForm1.SubmenuTit4(fila).Visible = False
 End If
Next fila

MDIForm1.Toolbar1.Buttons(3).Enabled = False
MDIForm1.Toolbar1.Buttons(5).Enabled = False
MDIForm1.Toolbar1.Buttons(6).Enabled = False
MDIForm1.Toolbar1.Buttons(10).Enabled = False
MDIForm1.Toolbar1.Buttons(1).Enabled = False
MDIForm1.Toolbar1.Buttons(7).Enabled = False
MDIForm1.Toolbar1.Buttons(9).Enabled = False
MDIForm1.Toolbar1.Buttons(8).Enabled = False
MDIForm1.Toolbar1.Buttons(11).Enabled = False
MDIForm1.Toolbar1.Buttons(4).Enabled = False
MDIForm1.Toolbar1.Buttons(12).Enabled = False
MDIForm1.Toolbar1.Buttons(13).Enabled = False
MDIForm1.Toolbar1.Buttons(14).Enabled = False
MDIForm1.Toolbar1.Buttons(15).Enabled = False
MDIForm1.Toolbar1.Buttons(16).Enabled = False
MDIForm1.Toolbar1.Buttons(18).Enabled = False


MDIForm1.menuAyuda.Enabled = False
MDIForm1.menuTit1.Enabled = False
MDIForm1.menuTit2.Enabled = False
MDIForm1.menutit3.Enabled = False
MDIForm1.menutit4.Enabled = False
MDIForm1.menutit5.Enabled = False
MDIForm1.menutit6.Enabled = False
MDIForm1.menutit7.Enabled = False

For S = 0 To MDIForm1.SubmenuTit1.count - 1
    If Not MDIForm1.SubmenuTit1.Item(S).Caption = "-" Then
        MDIForm1.SubmenuTit1.Item(S).Enabled = False
    End If
Next S
For S = 0 To MDIForm1.SubmenuTit2.count - 1
   If Not MDIForm1.SubmenuTit2.Item(S).Caption = "-" Then
    MDIForm1.SubmenuTit2.Item(S).Enabled = False
   End If
Next S
For S = 0 To MDIForm1.submenutit3.count - 1
   If Not MDIForm1.submenutit3.Item(S).Caption = "-" Then
    MDIForm1.submenutit3.Item(S).Enabled = False
   End If
Next S
For S = 0 To MDIForm1.SubmenuTit4.count - 1
   If Not MDIForm1.SubmenuTit4.Item(S).Caption = "-" Then
    MDIForm1.SubmenuTit4.Item(S).Enabled = False
   End If
Next S
For S = 0 To MDIForm1.submenutit5.count - 1
   If Not MDIForm1.submenutit5.Item(S).Caption = "-" Then
    MDIForm1.submenutit5.Item(S).Enabled = False
   End If
Next S
For S = 0 To MDIForm1.SubmenuTit6.count - 1
   If Not MDIForm1.SubmenuTit6.Item(S).Caption = "-" Then
    MDIForm1.SubmenuTit6.Item(S).Enabled = False
   End If
Next S
For S = 0 To MDIForm1.SubmenuTit7.count - 1
   If Not MDIForm1.SubmenuTit7.Item(S).Caption = "-" Then
    MDIForm1.SubmenuTit7.Item(S).Enabled = False
   End If
Next S

Exit Sub
sigue:
'MsgBox Err.Description
Resume Next

End Sub

Public Function REP_TRANSAC() As Integer
Dim WSRUTA As String
Dim indice As Integer
Dim wm As Integer
Dim llave_rep01 As rdoResultset
Dim PS_REP01 As rdoQuery
Dim i As Integer
Dim valor
Dim loc_xl As Object
Dim wRuta As String
Dim WNUM As Integer
Dim WDEVI As String * 3
Dim WMONE As String * 1
pub_cadena = "SELECT * FROM REP_TRANSA WHERE REP_CODTRA = ? ORDER BY REP_CODTRA"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurValues)
PS_REP01(0) = tra_llave!TRA_KEY
llave_rep01.Requery
If llave_rep01.EOF Then
  REP_TRANSAC = -1
  Exit Function
End If
If Trim(Nulo_Valors(llave_rep01!REP_ACTIVO)) <> "A" Then
  REP_TRANSAC = -1
  Exit Function
End If
If LK_FAC_IMP = "A" Then
  REP_TRANSAC = -1
  Exit Function
End If
'pub_mensaje = "Desea Imprimir Ahora... ?"
'Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
'If Pub_Respuesta = vbNo Then
'  REP_TRANSAC = -1
'  Exit Function
'End If
wRuta = PUB_RUTA_OTRO
SQ_OPER = 2
PUB_CODCIA = LK_CODCIA
PUB_CODVEN = val(FORMGEN.i_codven.Text)
LEER_PAR_LLAVE
If pac_llave.EOF Then
   MsgBox "No se ha definido archivos de Impresión", 48, Pub_Titulo
   Exit Function
End If


If llave_rep01!REP_CODTRA = 2401 Or llave_rep01!REP_CODTRA = 2414 Then
    FORMGEN.Reportes.Connect = PUB_ODBC
    WDEVI = Nulo_Valors(LK_DEVICE_FBG)
    If Trim(WDEVI) <> "" Then
      If PUB_FBG = "F" Then
        WNUM = Mid(WDEVI, 1, 1)
      ElseIf PUB_FBG = "B" Then
        WNUM = Mid(WDEVI, 2, 1)
      ElseIf PUB_FBG = "G" Then
        WNUM = Mid(WDEVI, 3, 1)
      End If
      On Error GoTo SALIMP
      FORMGEN.Reportes.PrinterName = Printers(WNUM).DeviceName
      FORMGEN.Reportes.PrinterDriver = Printers(WNUM).DriverName '"RASDD.DLL"
      FORMGEN.Reportes.PrinterPort = Printers(WNUM).Port
      
    End If

    If Trim(Nulo_Valors(llave_rep01!REP_IMP)) <> "A" Then
        FORMGEN.Reportes.WindowLeft = 2
        FORMGEN.Reportes.WindowTop = 70
        FORMGEN.Reportes.WindowWidth = 635
        FORMGEN.Reportes.WindowHeight = 390
        FORMGEN.Reportes.Destination = crptToWindow
        REP_TRANSAC = 0
    Else
        FORMGEN.Reportes.Destination = crptToPrinter
        REP_TRANSAC = 1
    End If
    FORMGEN.Reportes.Formulas(0) = ""
    FORMGEN.Reportes.Formulas(1) = ""
    FORMGEN.Reportes.Formulas(2) = ""
    FORMGEN.Reportes.Formulas(3) = ""
    FORMGEN.Reportes.Formulas(4) = ""
    If llave_rep01!REP_CODTRA = 2414 Then
      pub_cadena = "{FACART.FAR_ESTADO} <> 'E' AND {FACART.FAR_TIPMOV} = 93 and {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_NUMFAC} = " & PU_NUMFAC
      FORMGEN.Reportes.ReportFileName = wRuta & "CAMBP.RPT"
      GoTo pasa_todo
    End If
    If Trim(FORMGEN.i_moneda.Text) = "$" Then
      WMONE = "D"
    Else
      WMONE = "S"
    End If
   
    FORMGEN.Reportes.Formulas(1) = "SON=  '" & CONVER_LETRAS(PUB_NETO, WMONE) & "'"
    SQ_OPER = 1 'Dato de Operacion al leer_far_llave --JC
    LEER_FAR_LLAVE 'Leo de facart --JC
    If Trim(FORMGEN.i_fbg.Text) = "B" Then
      FORMGEN.Reportes.WindowTitle = "BOLETA  :" & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "00000000")
      'CONDICION PARA IMPRIMIR: SI ES PARA CONSUMO VA A OTRO REPORTE
      'SINO ES EL NOMBRE DEL DATO DE LA COLUMNA PAC_ARCHI_B DE LA TABLA PARARC--JC
        If far_llave!far_JABAS = 0 Then 'por detalle
            FORMGEN.Reportes.ReportFileName = wRuta + Trim(pac_llave!PAC_ARCHI_B)
        Else
            FORMGEN.Reportes.ReportFileName = wRuta + "BOLCONSUMO.rpt"
        End If
        '**********************************
      'FORMGEN.Reportes.ReportFileName = wRuta & Trim(pac_llave!PAC_ARCHI_B) ' "CLIBOL.RPT"
    ElseIf Trim(FORMGEN.i_fbg.Text) = "F" Then
      FORMGEN.Reportes.WindowTitle = "FACTURA : " & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "00000000")
      'CONDICION PARA IMPRIMIR: SI ES PARA CONSUMO VA A OTRO REPORTE
        'SINO ES EL NOMBRE DEL DATO DE LA COLUMNA PAC_ARCHI_F DE LA TABLA PARARC--JC
        If far_llave!far_JABAS = 0 Then 'por detalle
            FORMGEN.Reportes.ReportFileName = wRuta + Trim(pac_llave!PAC_ARCHI_F)
        Else
            FORMGEN.Reportes.ReportFileName = wRuta + "FACCONSUMO.rpt"
        End If
        '**********************************
      'FORMGEN.Reportes.ReportFileName = wRuta & Trim(pac_llave!PAC_ARCHI_F) ' "CLIFAC.RPT"
    ElseIf Trim(FORMGEN.i_fbg.Text) = "G" Then
      FORMGEN.Reportes.WindowTitle = " GUIA   : " & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "00000000")
      FORMGEN.Reportes.ReportFileName = wRuta & Trim(pac_llave!PAC_ARCHI_G) ' "CLIGUI.RPT"
    ElseIf Trim(FORMGEN.i_fbg.Text) = "P" Then
      FORMGEN.Reportes.WindowTitle = "PEDIDO : " & Format(PU_NUMSER, "000") & " - " & Format(PU_NUMFAC, "00000000")
      'CONDICION PARA IMPRIMIR: SI ES PARA CONSUMO O DETALLE
        If far_llave!far_JABAS = 0 Then 'por detalle
            FORMGEN.Reportes.ReportFileName = wRuta + "PEDDETALLE.rpt"
        Else
            FORMGEN.Reportes.ReportFileName = wRuta + "PEDCONSUMO.rpt"
        End If
        '**********************************
    End If
    pub_cadena = "{FACART.FAR_ESTADO} <> 'E' AND {FACART.FAR_TIPMOV} = 10 and {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' and {FACART.FAR_FBG} = '" & Trim(FORMGEN.i_fbg.Text) & "' AND {FACART.FAR_NUMSER}= '" & PU_NUMSER & "' AND {FACART.FAR_NUMFAC} = " & PU_NUMFAC
    'If LK_EMP = "PIU" Then
pasa_todo:
         FORMGEN.Reportes.SelectionFormula = pub_cadena
         On Error GoTo accion
         DoEvents
         FORMGEN.Reportes.Action = 1
         DoEvents
        ' On Error GoTo 0
      ' If LK_CODTRA = 2401 And Val(FORMGEN.i_numguia.Text) > 0 Then
       '  FORMGEN.Reportes.WindowTitle = "GUIA DE VENTA  " & FORMGEN.Reportes.WindowTitle
       '  FORMGEN.Reportes.ReportFileName = wRuta & "FACGUIA.RPT"
       '  pub_mensaje = "Ahora Desea Imprimir la " & Trim(FORMGEN.Reportes.WindowTitle) & "   ¿ Continuar... ?"
       '  Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
       '  If Pub_Respuesta = vbYes Then
       '    FORMGEN.Reportes.SelectionFormula = pub_cadena
       '    On Error GoTo accion
       '    DoEvents
       '    FORMGEN.Reportes.Action = 1
       '    DoEvents
       '    On Error GoTo 0
       '  End If
      'End If
Exit Function

accion:
On Error GoTo err_final
'MsgBox Err.Number & Err.Description
 MsgBox "Se produjo un conflicto Nro:" & Err.Number & Chr(13) & "El Sistema Intentará mostrar la impresión por pantalla", vbInformation, Pub_Titulo
 'MsgBox "Intente Nuevamente, la impresion de Modo manual", 48, Pub_Titulo
 If Trim(Nulo_Valors(llave_rep01!REP_IMP)) <> "A" Then
 Else
   FORMGEN.Reportes.WindowLeft = 2
   FORMGEN.Reportes.WindowTop = 70
   FORMGEN.Reportes.WindowWidth = 635
   FORMGEN.Reportes.WindowHeight = 390
   FORMGEN.Reportes.Destination = crptToWindow
   FORMGEN.Reportes.Action = 1
   REP_TRANSAC = 0
 End If
 Exit Function
End If
On Error GoTo 0
'Reporte de Excel
If LK_EMP = "CAM" Then
  GoTo VOUCHER
End If



WSRUTA = Trim(llave_rep01!REP_RUTA)
If loc_xl Is Nothing Then
    Set loc_xl = CreateObject("Excel.Application")
End If

loc_xl.Workbooks.Open WSRUTA, 0, True, 4, PUB_CLAVE, PUB_CLAVE
fila = 2
wm = 1
Do Until val(tra_llave(fila)) = 0 Or fila = 62
  wm = wm + 1
  If Trim(Nulo_Valors(llave_rep01(wm))) <> "" Then
      indice = TABLA_TAG(tra_llave(fila))
      If tra_llave(fila) = 19 Then
         valor = Trim(FORMGEN.i_nomban)
      ElseIf tra_llave(fila) = 4 Then
         valor = Trim(FORMGEN.i_nomCLI)
      ElseIf tra_llave(fila) = 70 Then
         valor = Trim(FORMGEN.i_nomban2)
      ElseIf tra_llave(fila) = 41 Then
         valor = Trim(FORMGEN.i_nomven)
      Else
       On Error GoTo sigue
        valor = FORMGEN.Controls(indice).Text
       On Error GoTo 0
      End If
      For i = 1 To loc_xl.Names.count
         If Right(loc_xl.Names(i).Name, 3) = Trim(Format(tra_llave(fila), "000")) Then
           loc_xl.APPLICATION.Range(loc_xl.APPLICATION.Names(i).Value).Value = valor
           Exit For
         End If
      Next i
  End If
  
fila = fila + 4

Loop
If Trim(Nulo_Valors(llave_rep01!REP_CIA)) <> "" Then
  valor = Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia")))
  For i = 1 To loc_xl.Names.count
    If loc_xl.Names(i).Name = llave_rep01!REP_CIA Then
      loc_xl.APPLICATION.Range(loc_xl.APPLICATION.Names(i).Value).Value = valor
      Exit For
    End If
  Next i
End If
If Trim(Nulo_Valors(llave_rep01!REP_DIA)) <> "" Then
  valor = Format(LK_FECHA_DIA, "dd/mm/yyyy")
  For i = 1 To loc_xl.Names.count
    If loc_xl.Names(i).Name = llave_rep01!REP_DIA Then
      loc_xl.APPLICATION.Range(loc_xl.APPLICATION.Names(i).Value).Value = valor
      Exit For
    End If
  Next i
End If
If Trim(Nulo_Valors(llave_rep01!REP_USU)) <> "A" Then
  valor = LK_CODUSU
  For i = 1 To loc_xl.Names.count
    If loc_xl.Names(i).Name = llave_rep01!REP_USU Then
      loc_xl.APPLICATION.Range(loc_xl.APPLICATION.Names(i).Value).Value = valor
      Exit For
    End If
  Next i
End If
If Trim(Nulo_Valors(llave_rep01!REP_TRA)) <> "A" Then
  valor = tra_llave!TRA_KEY
  For i = 1 To loc_xl.Names.count
    If loc_xl.Names(i).Name = llave_rep01!REP_TRA Then
      loc_xl.APPLICATION.Range(loc_xl.APPLICATION.Names(i).Value).Value = valor
      Exit For
    End If
  Next i
End If
loc_xl.DisplayAlerts = False
If Trim(Nulo_Valors(llave_rep01!REP_IMP)) <> "A" Then
 loc_xl.APPLICATION.Visible = True
 loc_xl.ActiveWindow.Activate
 REP_TRANSAC = 0
Else
  loc_xl.APPLICATION.Quit
  loc_xl.Worksheets.PrintOut
  REP_TRANSAC = 1
End If
Set loc_xl = Nothing
Exit Function

VOUCHER:


Exit Function




sigue:
valor = ""
SALIMP:
Resume Next
Exit Function
err_final:

End Function

Public Sub OPEN_LOG(wMensaje As String)
Dim RUTA As String
RUTA = PUB_RUTA_OTRO & "WSLOG.txt"
On Error GoTo pasa
Kill RUTA
On Error GoTo 0
On Error GoTo SALE
Open RUTA For Output As #1
Print #1, wMensaje
Print #1, "Fecha:" & Format(LK_FECHA_DIA, "dd/mm/yyyy") & " Hora: "; Format(Now, "hh:mm:ss AMPM")
Exit Sub
SALE:
    MsgBox Err.Number & Err.Description, 48, Pub_Titulo
 Close #1
 Exit Sub
pasa:
 Resume Next
End Sub
Public Sub CLOSE_LOG()
Dim RUTA As String
On Error GoTo SALE
'OBSERVACION : RUTA = PUB_RUTA_OTRO & "WSLOG.txt"
Print #1, "Fin de Proceso."
Close #1
Exit Sub
SALE:
    MsgBox Err.Number & Err.Description, 48, Pub_Titulo
End Sub

Public Sub WRITE_LOG(wWRITE As String)
Dim RUTA As String
On Error GoTo SALE
' OBSERVACION: RUTA = PUB_RUTA_OTRO & "WSLOG.txt"
Print #1, wWRITE
Exit Sub
SALE:
    MsgBox Err.Number & Err.Description, 48, Pub_Titulo
End Sub
Public Sub MOSTRAR_LOG()
Dim RUTA As String
Dim WL As Object
On Error GoTo SALE
RUTA = PUB_RUTA_OTRO & "WSLOG.txt"

On Error GoTo 0
If WL Is Nothing Then
    Set WL = CreateObject("word.Application")
End If
On Error GoTo no_existe
WL.APPLICATION.WindowState = 1
WL.Documents.Open FileName:=RUTA
WL.APPLICATION.Visible = True
Set WL = Nothing
On Error GoTo 0
Exit Sub


no_existe:
'Call Shell("C:\ADMIN\HERTISA\WSLOG.TXT", 1)
Exit Sub
SALE:
    MsgBox Err.Number & Err.Description, 48, Pub_Titulo
End Sub

Public Sub LISTA_TABLAS(LV1 As ListView, Optional wmax)
On Error GoTo SALE
Dim wmaximo As Integer
Dim itmX As ListItem
If Not IsMissing(wmax) Then wmaximo = wmax Else wmaximo = 1000
Set PSX = CN.CreateQuery("", archi)
Set x = PSX.OpenResultset(rdOpenForwardOnly)
x.Requery
LV1.ListItems.Clear
LV1.ColumnHeaders.Clear
If x.EOF Then LV1.Visible = False: Exit Sub
LV1.Top = 1800
LV1.Left = 3000
LV1.Width = 6500
LV1.Height = 3200
LV1.Visible = True
If numarchi = 99 Then ' para codigos alternos
 LV1.ColumnHeaders.Add 1, , "Descripción", 4000
 LV1.ColumnHeaders.Add 2, , "Cod.", 400
End If
Do Until x.EOF Or x.AbsolutePosition - 1 >= wmaximo
   If numarchi = 99 Then
     Set itmX = LV1.ListItems.Add(, , Trim(CStr(x.rdoColumns(3))))
          itmX.SubItems(1) = Trim(CStr(x.rdoColumns(2)))
   End If
   itmX.Tag = x.AbsolutePosition
   x.MoveNext
Loop
LV1.ToolTipText = "Encontrados : " & itmX.Tag & "/" & x.RowCount & " Muestra un Maximo de: " & wmaximo
Exit Sub
SALE:
If Err.Number = 40002 Then Exit Sub Else MsgBox Err.Description, 48, Pub_Titulo
End Sub

Public Sub ACT_TIPO_CAMBIO()
'LK_TIPO_CAMBIO = Format(LK_TIPO_CAMBIO, "0.0000")
Dim WC
WC = InputBox("Tipo de Cambio para - Fecha: " + Format(LK_FECHA_DIA, "dd/mm/yyyy") + " es :", "Actualizar Tipo de Cambio ", Format(LK_TIPO_CAMBIO, "0.0000000"))
If WC = "" Then Exit Sub
If val(WC) <= 0 Then
   MsgBox " NO Procede...", 48, Pub_Titulo
   Exit Sub
End If
LK_TIPO_CAMBIO = Format(val(WC), "0.0000000")
GEN.Edit
GEN!gen_tipo_cambio = LK_TIPO_CAMBIO
GEN.Update
GEN.Requery
'MDIForm1.StatusBar1.Panels(3).Text = "T.C.= S/. " + Format(LK_TIPO_CAMBIO, "0.0000")
End Sub
Public Function JALAR_TC(wfecha As Date, WES As Integer) As Currency
PUB_CAL_INI = wfecha
PUB_CAL_FIN = wfecha
pu_codcia = LK_CODCIA
PUB_CODCIA = LK_CODCIA
SQ_OPER = 1
LEER_CAL_LLAVE
If cal_llave.EOF Then
  JALAR_TC = 0
  Exit Function
End If
If val(WES) = 3 Then
 If IsNull(cal_llave!cal_tipo_cambio) Then
    JALAR_TC = 0
    Exit Function
 End If
 JALAR_TC = cal_llave!cal_tipo_cambio
End If

If val(WES) = 4 Then
 If IsNull(cal_llave!CAL_TC_MERCA) Then
    JALAR_TC = 0
    Exit Function
 End If
 JALAR_TC = cal_llave!CAL_TC_MERCA
End If


If val(WES) = 1 Then
 If IsNull(cal_llave!cal_tc_ingre) Then
    JALAR_TC = 0
    Exit Function
 End If
 JALAR_TC = cal_llave!cal_tc_ingre
End If
If val(WES) = -1 Then
 If IsNull(cal_llave!cal_tc_salid) Then
    JALAR_TC = 0
    Exit Function
 End If
 JALAR_TC = cal_llave!cal_tc_salid
End If

End Function

Public Sub PROC_TRANSA(LV1 As Object, Optional wmax)
On Error GoTo SALE
Dim wmaximo As Integer
Dim itmX As Object
If Not IsMissing(wmax) Then wmaximo = wmax Else wmaximo = 1000
Set PSX = CN.CreateQuery("", archi)
Set x = PSX.OpenResultset(rdOpenForwardOnly)
x.Requery
LV1.ListItems.Clear
LV1.ColumnHeaders.Clear
If x.EOF Then LV1.Visible = False: Exit Sub
LV1.Top = 3380
LV1.Height = 2900
LV1.Width = 10100
LV1.Left = 10
If numarchi = 9 Then
 LV1.ColumnHeaders.Add 1, , "Operaciones", 6200
 LV1.ColumnHeaders.Add 2, , "Codigo.", 1000
 LV1.ColumnHeaders.Add 3, , "Interno", 0
End If
Do Until x.EOF Or x.AbsolutePosition - 1 >= wmaximo
  If numarchi = 9 Then
    Set itmX = LV1.ListItems.Add(, , Trim(CStr(x.rdoColumns(1))))
    itmX.SubItems(1) = Trim(CStr(x.rdoColumns(2)))
   ' itmX.SubItems(3) = Format(X.rdoColumns(4), "0.00")
    itmX.SubItems(2) = Format(x.rdoColumns(0), "0.00")
   '      itmX.SubItems(3) = Format(X.rdoColumns(3), "0.00")
  End If
  itmX.Tag = x.AbsolutePosition
  x.MoveNext
Loop
LV1.ToolTipText = "Encontrados : " & itmX.Tag & "/" & x.RowCount & " Muestra un Maximo de: " & wmaximo
LV1.Visible = True
'DoEvents
Exit Sub
SALE:
MsgBox Err.Description
'Resume Next
If Err.Number = 40002 Then Exit Sub Else MsgBox Err.Description, 48, Pub_Titulo

End Sub




Public Function CHE_BLOQ_MES(WTIPMOV As Integer) As Integer
Dim wcadena
Dim wvalor As String
CHE_BLOQ_MES = 0
SQ_OPER = 2
PUB_TIPREG = 155
PUB_CODCIA = cop_llave!cop_codcia
LEER_TAB_LLAVE
If tab_mayor.EOF Then
  MsgBox "Crear tab_tipreg = 155 para seguridad de meses.", 48, Pub_Titulo
  CHE_BLOQ_MES = 1
Else
    Do Until tab_mayor.EOF
      If tab_mayor!TAB_NUMTAB = val(Format(cop_llave!cop_fecha_proceso, "yyyy")) Then
        wcadena = Trim(tab_mayor!tab_NOMLARGO)
      End If
      tab_mayor.MoveNext
    Loop
    wvalor = Mid(wcadena, val(Format(cop_llave!cop_fecha_proceso, "mm")) + 1, 1)
    If wvalor = "1" Then
       'MsgBox "<<<< Mes CERRADO Operaciones >>>>", vbCritical, Pub_Titulo
       CHE_BLOQ_MES = 1
       Exit Function
    End If
End If

PSMOV_LLAVE(0) = cop_llave!cop_codcia
PSMOV_LLAVE(1) = cop_llave!cop_fecha_proceso
PSMOV_LLAVE(2) = cop_llave!cop_fecha_proceso2
PSMOV_LLAVE(3) = WTIPMOV
PSMOV_LLAVE(4) = cop_llave!cop_nro_mes
mov_llave.Requery
par_MOV_CODCIA = cop_llave!cop_codcia
If Not mov_llave.EOF Then
   CHE_BLOQ_MES = 9
End If


End Function

Public Sub ASIENTO_MOVICONT(EE As Excel.APPLICATION, WTIPMOV As Integer)
On Error GoTo SALE
Dim FLAGX As String * 1
Dim correla As Integer

FLAGX = ""
correla = 4

PSCOP_LLAVE.rdoParameters(0) = LK_CODCIA
cop_llave.Requery

  
par_MOV_NRO_VOUCHER = 0
par_MOV_NRO_MOV = 0

par_MOV_TIPMOV = WTIPMOV
par_MOV_FECHA = cop_llave!cop_fecha_proceso2
par_MOV_fecha_EMI = cop_llave!cop_fecha_proceso2
par_MOV_nro_MES = val(cop_llave!cop_nro_mes)
par_MOV_MONEDA = "S"
par_MOV_FLAG_DES = " "
If par_MOV_TIPMOV = 1 Then
 par_MOV_DETALLE = "Por las Compras del periodo"
 par_MOV_PLANTILLA = 1
ElseIf par_MOV_TIPMOV = 2 Then
 par_MOV_DETALLE = "Por las Ventas del periodo"
 par_MOV_PLANTILLA = 1
ElseIf par_MOV_TIPMOV = 3 Then
 par_MOV_DETALLE = "Resúmen Egresos caja "
 par_MOV_PLANTILLA = 126
ElseIf par_MOV_TIPMOV = 4 Then
 par_MOV_DETALLE = "Por los Asientos varios "
 par_MOV_PLANTILLA = 1
End If
par_MOV_GLOSA = par_MOV_DETALLE
' llave para el nro de asiento
PSMOV_LLAVE(0) = cop_llave!cop_codcia
PSMOV_LLAVE(1) = cop_llave!cop_fecha_proceso
PSMOV_LLAVE(2) = cop_llave!cop_fecha_proceso2
PSMOV_LLAVE(3) = WTIPMOV
PSMOV_LLAVE(4) = cop_llave!cop_nro_mes
mov_llave.Requery
par_MOV_CODCIA = cop_llave!cop_codcia
If mov_llave.EOF Then
 par_MOV_NRO_VOUCHER = 0
Else
 par_MOV_NRO_VOUCHER = mov_llave!MOV_NRO_VOUCHER
End If
par_MOV_NRO_VOUCHER = par_MOV_NRO_VOUCHER + 1
'EE.Application.Visible = True
Do Until FLAGX = "A"
   correla = correla + 1
   If val(EE.Cells(correla, 1)) = 0 Then
     If UCase(Left(Trim(EE.Cells(correla, 1)), 5)) = "TOTAL" And val(EE.Cells(correla + 1, 1)) <> 0 Then
        par_MOV_NRO_MOV = 0
        par_MOV_NRO_VOUCHER = par_MOV_NRO_VOUCHER + 1
        par_MOV_DETALLE = "Resúmen Ingresos de Caja "
        par_MOV_PLANTILLA = 100
        correla = correla + 1
      Else
        Exit Do
      End If
   End If
   par_MOV_CODCTA = Format(EE.Cells(correla, 1), "##########")
   par_MOV_DH = Format(EE.Cells(correla, 5), "#")
   par_MOV_IMPORTE = val(redondea(EE.Cells(correla, 3))) + val(redondea(EE.Cells(correla, 4)))
   par_MOV_CODCIA = Format(cop_llave!cop_codcia, "00")
   ' CONSISTENCIAS DE NEGATIVOS
   If par_MOV_IMPORTE < 0 Then
      par_MOV_IMPORTE = Abs(par_MOV_IMPORTE)
      If par_MOV_DH = "D" Then
        par_MOV_DH = "H"
      Else
        par_MOV_DH = "D"
      End If
   End If
   mov_llave.AddNew
   mov_llave!MOV_NRO_MOV = par_MOV_NRO_MOV
   mov_llave!MOV_CODCIA = par_MOV_CODCIA
   mov_llave!MOV_NRO_VOUCHER = par_MOV_NRO_VOUCHER
   mov_llave!MOV_TIPMOV = par_MOV_TIPMOV
   mov_llave!MOV_FECHA = par_MOV_FECHA
   mov_llave!MOV_GLOSA = par_MOV_GLOSA
   mov_llave!MOV_MONEDA = par_MOV_MONEDA
   mov_llave!MOV_CODCTA = par_MOV_CODCTA
   mov_llave!MOV_DH = par_MOV_DH
   mov_llave!MOV_IMPORTE = par_MOV_IMPORTE
   mov_llave!MOV_SUNAT = "00"
   mov_llave!MOV_serie = 0
   mov_llave!MOV_numfac = 0
   mov_llave!MOV_codclie = 0
   mov_llave!MOV_CP = " "
   mov_llave!MOV_FBG = " "
   mov_llave!MOV_MARCA = "X"
   mov_llave!MOV_DETALLE = par_MOV_DETALLE
   mov_llave!MOV_FBG_C = " "
   mov_llave!MOV_NUMFAC_C = 0
   mov_llave!MOV_SERIE_C = 0
   mov_llave!MOV_nro_MES = par_MOV_nro_MES
   mov_llave!MOV_fecha_EMI = par_MOV_fecha_EMI
   mov_llave!MOV_PLANTILLA = par_MOV_PLANTILLA
   mov_llave!MOV_FLAG_TC = ""
   mov_llave!MOV_TIPO_CAMBIO = 1
   mov_llave!MOV_FLAG_DES = par_MOV_FLAG_DES
   mov_llave!MOV_CODUSU = LK_CODUSU
   mov_llave.Update
   par_MOV_NRO_MOV = par_MOV_NRO_MOV + 1
Loop

Exit Sub
SALE:
MsgBox Err.Description
Resume Next
End Sub

Public Sub CANCEL_CH()
Dim WS_TOT As Currency
Dim WS_IMPORTE_AMORT As Currency
Dim PSCH_LLAVE As rdoQuery
Dim ch_llave As rdoResultset
Dim WNUMERO As String
Dim wser
Dim wFAC
Dim WFBG
Dim PSFA_LLAVE As rdoQuery
Dim fa_llave As rdoResultset

pub_cadena = "SELECT CAR_FBG ,CAR_NUMSER, CAR_NUMFAC, car_fecha_vcto ,CAR_IMP_INI, CAR_SERDOC, CAR_NUMDOC,CAR_CP,CAR_CODCLIE, CAR_TIPDOC , CAR_FBG, CAR_NUMFAC_C, CAR_NUMSER_C, CAR_IMPORTE  FROM CARTERA WHERE CAR_CODCIA = ? AND (CAR_TIPDOC = ? OR CAR_TIPDOC = 'CC') AND CAR_CODCLIE = ? AND CAR_FBG = ? AND CAR_NUMSER = ? AND CAR_NUMFAC = ? AND CAR_IMPORTE <> 0  "
Set PSFA_LLAVE = CN.CreateQuery("", pub_cadena)
PSFA_LLAVE(0) = 0
PSFA_LLAVE(1) = 0
PSFA_LLAVE(2) = 0
PSFA_LLAVE(3) = 0
PSFA_LLAVE(4) = 0
PSFA_LLAVE(5) = 0

Set fa_llave = PSFA_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT CAR_NUM_CHEQUE, CAR_FECHA_VCTO, CAR_IMP_INI, CAR_NUMSER, CAR_NUMFAC, CAR_CONCEPTO, CAR_SERDOC, CAR_NUMDOC,CAR_CP,CAR_CODCLIE, CAR_TIPDOC , CAR_FBG, CAR_NUMFAC_C, CAR_NUMSER_C, CAR_IMPORTE  FROM CARTERA WHERE CAR_CODCIA = ? AND CAR_TIPDOC = ? AND CAR_IMPORTE <> 0  "
Set PSCH_LLAVE = CN.CreateQuery("", pub_cadena)
PSCH_LLAVE(0) = 0
PSCH_LLAVE(1) = ""
Set ch_llave = PSCH_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
PSCH_LLAVE(0) = LK_CODCIA
PSCH_LLAVE(1) = "CH"
ch_llave.Requery

SQ_OPER = 2
pu_codcia = LK_CODCIA
PUB_FECHA = LK_FECHA_DIA
LEER_ALL_LLAVE
pub_numoper = 0
If all_menor.EOF Then
  pub_numoper = 1 + 30000
Else
  pub_numoper = all_menor!ALL_NUMOPER + 30000
End If
PRODIA.ProgBar.Visible = True
' ch_llave
If Not ch_llave.EOF Then PRODIA.ProgBar.max = ch_llave.RowCount
PRODIA.ProgBar.Min = 0
PRODIA.ProgBar.Value = 0
PRODIA.POR(2).Caption = "Liquidando Cheques Pendientes..."
Do Until ch_llave.EOF
      PRODIA.ProgBar.Value = PRODIA.ProgBar.Value + 1
      PSFA_LLAVE(0) = LK_CODCIA
      PSFA_LLAVE(1) = "FA"
      WNUMERO = Mid(ch_llave!car_concepto, 4, Len(Trim(ch_llave!car_concepto)))
      wFAC = Mid(WNUMERO, InStr(1, WNUMERO, "-") + 1, Len(WNUMERO))
      wser = Mid(WNUMERO, 1, InStr(1, WNUMERO, "-") - 1)
      WFBG = Left(ch_llave!car_concepto, 1)
      PUB_CODCLIE = ch_llave!CAR_CODCLIE
      PSFA_LLAVE(2) = ch_llave!CAR_CODCLIE
      PSFA_LLAVE(3) = WFBG
      PSFA_LLAVE(4) = wser
      PSFA_LLAVE(5) = wFAC
      fa_llave.Requery
      If fa_llave.EOF Then
        MsgBox "Cheque no Aplicado a ningun Documento " & Chr(13) & " OJO  - A N O T A R :" & Chr(13) & "Nº.Ch/: " & ch_llave!car_NUM_CHEQUE & Chr(13) & "Nº Documento : " & Trim(ch_llave!car_concepto) & Chr(13) & "Importe : " & Format(ch_llave!car_importe, "##,##0.00"), 48, Pub_Titulo
        GoTo bien
      End If
      PUB_SERDOC = fa_llave!car_SERDOC
      PUB_NUMDOC = fa_llave!car_NUMDOC
      PUB_CP = fa_llave!CAR_cp
      PUB_TIPDOC = fa_llave!car_TIPDOC
      PUB_FECHA_VCTO = fa_llave!car_fecha_vcto
      PUB_CONCEPTO = "Liq. CH/. " & ch_llave!car_NUM_CHEQUE & "- automatica."
      PUB_NUMFAC_C = fa_llave!CAR_NUMFAC_C
      PUB_NUMSER_C = fa_llave!CAR_NUMSER_C
      PUB_NUMSER = fa_llave!car_NUMSER
      PUB_NUMFAC = fa_llave!car_NUMFAC
      PUB_CHENUM = 0
      WS_TOT = fa_llave!CAR_IMP_INI
      PUB_FBG = fa_llave!car_FBG
      WS_IMPORTE_AMORT = val((ch_llave!car_importe))
      pub_signo_car = -1
      fa_llave.Edit
      fa_llave!car_importe = Format(val(fa_llave!car_importe) + val((ch_llave!car_importe)), "0.00")
      fa_llave.Update
      pub_numoper = pub_numoper + 1
      GoSub REGISTRAR
      pub_signo_car = 1
      PUB_SERDOC = ch_llave!car_SERDOC
      PUB_NUMDOC = ch_llave!car_NUMDOC
      PUB_CP = ch_llave!CAR_cp
      PUB_TIPDOC = ch_llave!car_TIPDOC
      PUB_FECHA_VCTO = ch_llave!car_fecha_vcto
      PUB_CONCEPTO = "Liquidacion automatica"
      PUB_NUMFAC_C = ch_llave!CAR_NUMFAC_C
      PUB_NUMSER_C = ch_llave!CAR_NUMSER_C
      PUB_CHENUM = val(ch_llave!car_NUM_CHEQUE)
      WS_TOT = ch_llave!CAR_IMP_INI
      PUB_FBG = ch_llave!car_FBG
      ch_llave.Edit
      ch_llave!car_importe = 0
      ch_llave.Update
      pub_numoper = pub_numoper + 1
      GoSub REGISTRAR
bien:
    ch_llave.MoveNext
Loop
PRODIA.POR(2).Caption = "Cerrando Operaciones Diarias..."

Exit Sub
REGISTRAR:
caa_histo.AddNew
caa_histo!CAA_CODCLIE = PUB_CODCLIE
caa_histo!caa_codcia = LK_CODCIA
caa_histo!CAA_TIPDOC = PUB_TIPDOC
caa_histo!CAA_CP = PUB_CP
caa_histo!CAA_NUM_OPER = pub_numoper
caa_histo!caa_INTVEN = 0
caa_histo!caa_DIASV = 0
caa_histo!caa_DIASA = 0
caa_histo!caa_tasav = 0
caa_histo!caa_TIPO_CAMBIO = LK_TIPO_CAMBIO
caa_histo!caa_serdoc = PUB_SERDOC
caa_histo!CAA_NUMDOC = PUB_NUMDOC
caa_histo!CAA_FECHA = LK_FECHA_DIA
caa_histo!CAA_FECHA_VCTO = PUB_FECHA_VCTO
caa_histo!caa_situacion = 0
caa_histo!caa_concepto = PUB_CONCEPTO
caa_histo!CAA_IMPORTE = Abs(WS_IMPORTE_AMORT) * pub_signo_car
caa_histo!CAA_TOTAL = Abs(WS_TOT) * pub_signo_car
caa_histo!CAA_SALDO = 0 'Nulo_Valor0(cli_llave!cli_SALDO)
caa_histo!caa_SALDO_car = 0
caa_histo!CAA_SIGNO_CAJA = 0
caa_histo!CAA_SIGNO_CAJA_REAL = 0
caa_histo!CAA_SIGNO_CAR = pub_signo_car
caa_histo!CAA_TIPMOV = 0
caa_histo!CAA_hora = Now
caa_histo!CAA_CODUSU = LK_CODUSU
caa_histo!CAA_ESTADO = "N"
caa_histo!CAa_NOMBRE = ""
caa_histo!CAA_NUMPLAN = 0
caa_histo!CAA_FECHA_COBRO = LK_FECHA_DIA
caa_histo!CAa_NUM_CHEQUE = PUB_CHENUM
caa_histo!CAa_numser = PUB_NUMSER
caa_histo!CAa_numfac = PUB_NUMFAC
caa_histo!caa_numser_c = PUB_NUMSER_C
caa_histo!caa_numfac_c = PUB_NUMFAC_C
caa_histo!CAA_NOTA = ""
caa_histo!CAA_FBG = PUB_FBG
caa_histo!CAA_CODVEN = 0
caa_histo!CAa_numGUIA = 0
caa_histo!CAa_SERGUIA = 0
caa_histo!caa_situacion = 0
caa_histo!caa_FLAG_SO = " "
caa_histo!caa_signo_ccm = 0
caa_histo!caa_codban = 0
caa_histo!caa_codTRA = 2727
caa_histo!CAA_RECIBO = 0
caa_histo!CAA_SERIE = 0
caa_histo!caa_situacion = " "
caa_histo.Update

Return


End Sub
