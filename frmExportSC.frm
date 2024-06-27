VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmExportSC 
   BackColor       =   &H00808080&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Proceso de exportación de operaciones de la Gestión Comercial"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7755
   Icon            =   "frmExportSC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox txtReporte 
      Height          =   4455
      Left            =   0
      TabIndex        =   11
      Top             =   2250
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   7858
      _Version        =   393217
      TextRTF         =   $"frmExportSC.frx":000C
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00FFFFFF&
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
      Height          =   855
      Left            =   6465
      Picture         =   "frmExportSC.frx":008E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   975
      Width           =   1155
   End
   Begin VB.CommandButton cmdTransferir 
      BackColor       =   &H00FFFFFF&
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
      Height          =   855
      Left            =   5220
      Picture         =   "frmExportSC.frx":0904
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   960
      Width           =   1155
   End
   Begin ComctlLib.ProgressBar Barra 
      Height          =   240
      Left            =   30
      TabIndex        =   4
      Tag             =   "0"
      Top             =   6705
      Width           =   7680
      _ExtentX        =   13547
      _ExtentY        =   423
      _Version        =   327682
      Appearance      =   0
   End
   Begin MSMask.MaskEdBox fecha2 
      Height          =   285
      Left            =   4770
      TabIndex        =   9
      Top             =   165
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
   Begin MSMask.MaskEdBox fecha1 
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Top             =   150
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
      Caption         =   "Seleccione el intervalo de Fechas que desea Exportar a movicont luego haga click en transferir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   585
      Index           =   1
      Left            =   135
      TabIndex        =   6
      Top             =   795
      Width           =   3855
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ultimos Movimientos a Importar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Index           =   0
      Left            =   2130
      TabIndex        =   5
      Top             =   1980
      Width           =   3150
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
      Left            =   4170
      TabIndex        =   2
      Top             =   210
      Width           =   540
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
      Left            =   540
      TabIndex        =   1
      Top             =   195
      Width           =   615
   End
   Begin VB.Label lblbarraos 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Solution - Gestion Contable"
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
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   6930
      Width           =   7770
   End
   Begin VB.Label lblbarraos 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   660
      Index           =   1
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   7740
   End
End
Attribute VB_Name = "frmExportSC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iDia As Integer
Dim iDia1 As Integer
Dim iMes As Integer
Dim iMes1 As Integer
Dim iAno As Integer
Dim iAno1 As Integer

Dim PsMovicont As rdoQuery
Dim RsMovicont As rdoResultset
Dim PsAllog As rdoQuery
Dim RsAllog As rdoResultset
Dim PsComaest As rdoQuery
Dim RsComaest As rdoResultset
Dim PSMOV_VOU As rdoQuery
Dim VOU_MOV As rdoResultset
Dim TipoMovimiento As Integer

Private Sub cmdSalir_Click()
    Set frmExportSC = Nothing
    Unload Me
End Sub

Private Sub cmdTransferir_Click()
Dim CodTra As Integer
Dim Opcion As Integer
Dim cuenta As String
Dim CuentaTmp As String
Dim DH As String
Dim DHTmp As String
Dim Campo As Integer
Dim CampoTmp As Integer
Dim Importe As Double
Dim iRecorrido As Integer
Dim TipoCambio As Double
Dim iSecuencia As Integer
Dim NumSer As String
Dim NumFac As Long
Dim NumVoucher As Integer
Dim iBarra As Integer
Dim sFBG As String

On Error GoTo Handler
    cmdTransferir.Enabled = False
    cmdSalir.Enabled = False
    PsAllog.rdoParameters(0) = LK_CODCIA
    PsAllog.rdoParameters(1) = FECHA1.Text
    PsAllog.rdoParameters(2) = fecha2.Text
    
    txtReporte.Text = "===================================================================================="
    txtReporte.Text = txtReporte.Text & vbCrLf & "         REPORTE DE OPERACIONES EXPORTADAS DEL MODULO DE GESTION COMERCIAL"
    txtReporte.Text = txtReporte.Text & vbCrLf & "===================================================================================="
    
    pub_cadena = "SELECT * FROM CONTROLL"
    CN.Execute "Begin Transaction", rdExecDirect
    Set con_llave = CN.OpenResultset(pub_cadena, rdOpenKeyset, rdConcurLock)
        
    archi = "SELECT * FROM contabilidad WHERE cnt_codcia='" & LK_CODCIA & "'"
    Set PSX = CN.CreateQuery("", archi)
    Set X = PSX.OpenResultset(rdOpenKeyset)
    X.Requery
    Do While Not X.EOF
        DoEvents
        CodTra = X("cnt_codtra")
        Opcion = X("cnt_secuencia")
        
        If CodTra = 1401 Then
            TipoMovimiento = 1 ' Registro de compras
        ElseIf CodTra = 1455 Or CodTra = 5335 Then
            TipoMovimiento = 9 ' Canje de Documentos
        ElseIf CodTra = 2401 Or CodTra = 2402 Then
            TipoMovimiento = 2 ' Registro de ventas
        ElseIf CodTra = 2720 Or CodTra = 2725 Or CodTra = 5310 Or CodTra = 5335 Then
            TipoMovimiento = 3 ' Ingresos de fondos
        ElseIf CodTra = 2738 Or CodTra = 5360 Or CodTra = 2735 Then
            TipoMovimiento = 4 ' Egresos de fondos
        ElseIf CodTra = 2741 Then
            TipoMovimiento = 17 'Obligaciones diversas
        ElseIf CodTra = 2748 Then
            TipoMovimiento = 18 'tranferencia
        End If
        
        PsAllog.rdoParameters(3) = CodTra
        PsAllog.rdoParameters(4) = Opcion
        RsAllog.Requery
        
        If RsAllog.RowCount > 0 Then Barra.Max = RsAllog.RowCount + 1
        
        Do While Not RsAllog.EOF
            iBarra = iBarra + 1
            Barra.Value = iBarra
            NumVoucher = NroVoucher
            iSecuencia = 0
            NumSer = ""
            NumFac = 0
            sFBG = ""
            If (RsAllog!ALL_codtra = 2735 Or RsAllog!ALL_codtra = 2748) And RsAllog!ALL_SIGNO_CCM = 0 Then
                GoTo OtroRegistro
            End If
            If (RsAllog!ALL_codtra = 5318) And RsAllog!ALL_SIGNO_CCM = -1 Then
                GoTo OtroRegistro
            End If
            If (RsAllog!ALL_codtra = 2770) And RsAllog!ALL_SIGNO_CCM <> 1 Then
                GoTo OtroRegistro
            End If
            If RsAllog!ALL_codtra = 1455 Then
                GoTo OtroRegistro
            End If
            NumSer = " "
            If RsAllog!ALL_codtra = 1401 Or RsAllog!ALL_codtra = 2720 Or RsAllog!ALL_codtra = 2725 Or RsAllog!ALL_codtra = 5310 Or RsAllog!ALL_codtra = 2741 Or RsAllog!ALL_codtra = 2748 Or RsAllog!ALL_codtra = 2770 Or RsAllog!ALL_codtra = 2735 Then
                NumSer = RsAllog!ALL_NUMSER_C
                NumFac = RsAllog!ALL_NUMFAC_C
            End If
            If RsAllog!ALL_CHENUM <> 0 And (RsAllog!ALL_codtra = 2738 Or RsAllog!ALL_codtra = 5360 Or RsAllog!ALL_codtra = 5318) Then
                NumSer = Val(RsAllog!ALL_CHESER)
                NumFac = RsAllog!ALL_CHENUM
            End If
            If RsAllog!ALL_codtra = 2401 Or RsAllog!ALL_codtra = 2402 Then  ' quite RsAllog!ALL_SIGNO_CAJA agregue RsAllog!ALL_NUMFAC_C para pasar
                NumSer = RsAllog!ALL_NUMSER
                NumFac = RsAllog!ALL_NUMFAC
            End If
            
            txtReporte.Text = txtReporte.Text & vbCrLf & "Transaccion = " & RsAllog("all_codtra") & "        NroDocumento : " & Trim(NumSer) & " - " & NumFac & vbCrLf
            
            If RsAllog!ALL_FBG = "G" Then GoTo OtroRegistro
            
            PsMovicont.rdoParameters(0) = LK_CODCIA
            PsMovicont.rdoParameters(1) = RsAllog!ALL_FECHA_SUNAT
            PsMovicont.rdoParameters(2) = Format(RsAllog!ALL_FECHA_SUNAT, "mm")
            PsMovicont.rdoParameters(3) = Val(NumSer)
            PsMovicont.rdoParameters(4) = NumFac
            PsMovicont.rdoParameters(5) = RsAllog!ALL_cp
            PsMovicont.rdoParameters(6) = RsAllog!ALL_CODCLIE
            PsMovicont.rdoParameters(7) = RsAllog!ALL_FBG
            PsMovicont.rdoParameters(8) = RsAllog!ALL_codtra
            PsMovicont.rdoParameters(9) = RsAllog!all_SECUENCIA
            RsMovicont.Requery
            Do While Not RsMovicont.EOF
                NumVoucher = RsMovicont("mov_nro_voucher")
                RsMovicont.Delete
                RsMovicont.MoveNext
            Loop
            
            SQ_OPER = 1
            PUB_CAL_INI = RsAllog!ALL_FECHA_DIA
            PUB_CAL_FIN = RsAllog!ALL_FECHA_DIA
            PUB_CODCIA = LK_CODCIA
            LEER_CAL_LLAVE
            If Nulo_Valor0(cal_llave!cal_tipo_cambio) = 0 Then
                MsgBox "Definir Tipo de Cambios para el Periodo Actual. Dia : " & RsAllog!ALL_FECHA_DIA & " (en el Calendario del Sistema)", 48, Pub_Titulo
                'GoTo fin
            Else
                TipoCambio = cal_llave!cal_tipo_cambio
            End If
            
            For iRecorrido = 0 To 11
                cuenta = Trim(X(3 * iRecorrido + 3))
                DH = X(3 * iRecorrido + 4)
                Campo = X(3 * iRecorrido + 5)
                If (Trim(CuentaTmp) = "" Or Trim(DHTmp) = "" Or CampoTmp = 0) And iRecorrido <> 0 Then GoTo OtraCuenta
                If cuenta <> CuentaTmp And iRecorrido <> 0 Then
                    cuenta = CuentaTmp
                    DH = DHTmp
                    Campo = CampoTmp
                    Importe = RsAllog(Campo)
                    CuentaTmp = Trim(X(3 * iRecorrido + 3))
                    DHTmp = X(3 * iRecorrido + 4)
                    CampoTmp = X(3 * iRecorrido + 5)
                    GoSub Procesar
                Else
                    Importe = Importe + RsAllog(Campo)
                    CuentaTmp = Trim(X(3 * iRecorrido + 3))
                    DHTmp = X(3 * iRecorrido + 4)
                    CampoTmp = X(3 * iRecorrido + 5)
                    GoTo OtraCuenta
                End If
                
                'GoSub Procesar
OtraCuenta:
            Next iRecorrido
            
            GoTo OtroRegistro
            
Procesar:
            iSecuencia = iSecuencia + 1
            If RsAllog!ALL_moneda_CAJA = "D" Then
                Importe = Importe * TipoCambio
            Else
                Importe = Importe
            End If
            
            RsMovicont.AddNew
            RsMovicont!MOV_CODCIA = LK_CODCIA
            RsMovicont!MOV_nro_MES = Format(FECHA1.Text, "mm")
            RsMovicont!MOV_NRO_VOUCHER = NumVoucher
            RsMovicont!MOV_NRO_MOV = iSecuencia
            RsMovicont!MOV_TIPMOV = TipoMovimiento
            RsMovicont!MOV_PERIODO = Format(RsAllog!ALL_FECHA_DIA, "yyyy")
            RsMovicont!MOV_FECHA = RsAllog!ALL_FECHA_DIA
            
            If Trim(cuenta) = "CLIENTES" Or Trim(cuenta) = "CLIENTES2" Then GoSub CtaClientes
            If Trim(cuenta) = "BANCOS" Or Trim(cuenta) = "BANCOS2" Then GoSub CtaBancos
            If Trim(cuenta) = "BANCOS" Or Trim(cuenta) = "BANCOS2" Then MsgBox "Bancos"
            
            RsMovicont!MOV_CODCTA = cuenta
            RsMovicont!MOV_DH = DH
            RsMovicont!MOV_IMPORTE = Importe
            RsMovicont!MOV_GLOSA = Trim(RsAllog!ALL_autocon)
            If RsAllog!ALL_SIGNO_CAJA = 0 Then
                RsMovicont!MOV_MONEDA = RsAllog!ALL_moneda_CLI
            Else
                RsMovicont!MOV_MONEDA = RsAllog!ALL_moneda_CAJA
            End If
            RsMovicont!MOV_serie = NumSer
            RsMovicont!MOV_numfac = NumFac
            If RsAllog!ALL_CHENUM <> 0 And RsAllog!ALL_SIGNO_CAR = 0 Then
                RsMovicont!MOV_FBG = "CH"
            Else
                RsMovicont!MOV_FBG = RsAllog!ALL_FBG
                RsMovicont!MOV_serie_c = RsAllog!ALL_NUMSER_C
                RsMovicont!MOV_numfac_c = RsAllog!ALL_NUMFAC_C
            End If
            
            RsMovicont!MOV_SUNAT = RsAllog!ALL_CODSUNAT
            RsMovicont!MOV_codclie = RsAllog!ALL_CODCLIE
            RsMovicont!MOV_CP = RsAllog!ALL_cp
            
            RsMovicont!MOV_MARCA = "X"
            RsMovicont!MOV_DETALLE = ""
            RsMovicont!MOV_FBG_C = ""
            
            RsMovicont!MOV_fecha_EMI = RsAllog!ALL_FECHA_SUNAT
            RsMovicont!MOV_PLANTILLA = 0
            RsMovicont!MOV_FLAG_TC = ""
            RsMovicont!MOV_TIPO_CAMBIO = RsAllog!ALL_tipo_cambio
            RsMovicont!MOV_FLAG_DES = ""
            RsMovicont!MOV_CODUSU = LK_CODUSU
            'rsMovicont!MOV_VOU2=
            RsMovicont!MOV_RUC = RsAllog!ALL_RUC
            RsMovicont!MOV_OPC = 0
            RsMovicont!MOV_EXONERADO = 1
            RsMovicont!MOV_CC = ""
            RsMovicont!MOV_CODTRA = CodTra
            RsMovicont!MOV_OPERACION = Opcion
            RsMovicont.Update
            
            txtReporte.Text = txtReporte.Text & "  Importe = " & Importe & "    Cuenta = " & cuenta & "     DH = " & DH & vbCrLf
            
            Importe = 0
            Return
OtroRegistro:
            RsAllog.MoveNext
        Loop
        Barra.Value = 0
        iBarra = 0
        X.MoveNext
    Loop
    CN.Execute "Commit Transaction", rdExecDirect
    con_llave.Close
    GoTo Termino
    
CtaClientes:
    If RsAllog!ALL_SIGNO_CAR <> 0 Then
        SQ_OPER = 1
        pu_cp = RsAllog!ALL_cp
        pu_codclie = RsAllog!ALL_CODCLIE
        pu_codcia = LK_CODCIA
        LEER_CLI_LLAVE
        If cli_llave.EOF Then
           MsgBox "OJO REVISAR CODIGO DE CLIENTES..." & PUB_CODCLIE
        Else
           If Trim(cuenta) = "CLIENTES" Then
              cuenta = Nulo_Valors(cli_llave!CLI_CUENTA_CONTAB)
           ElseIf Trim(cuenta) = "CLIENTES2" Then
              cuenta = Nulo_Valors(cli_llave!CLI_CUENTA_CONTAB2)
           End If
           If Trim(cuenta) = "" Then MsgBox "Falta cuenta contable (12)..." & pu_codclie
        End If
     End If
     Return

CtaBancos:
' Or _
        ((Trim(cuenta) = "BANCOS" Or Trim(cuenta) = "BANCOS2") And RsAllog!ALL_SIGNO_CCM = 0 And RsAllog!ALL_codtra = 2748)
     If (Trim(cuenta) = "BANCOS" Or Trim(cuenta) = "BANCOS2") And _
        ((RsAllog!ALL_SIGNO_CCM <> 0 And (RsAllog!ALL_codtra = 2735 Or RsAllog!ALL_codtra = 2748)) Or _
         (RsAllog!ALL_SIGNO_CCM = 1 And (RsAllog!ALL_codtra = 2770 Or RsAllog!ALL_codtra = 5318))) Then
        SQ_OPER = 1
        pu_cp = RsAllog!ALL_cp
        PUB_CODBAN = RsAllog!ALL_CODBAN
        pu_codcia = LK_CODCIA
        LEER_CCM_LLAVE
        If ccm_llave.EOF Then
           MsgBox "OJO REVISAR CODIGO DE BANCO..." & PUB_CODBAN
        Else
            If Trim(cuenta) = "BANCOS" Then
                cuenta = Nulo_Valors(ccm_llave!CCM_CUENTA_CONTAB2)
            ElseIf Trim(cuenta) = "BANCOS2" Then
                cuenta = Nulo_Valors(ccm_llave!CCM_CUENTA_CONTAB)
            End If
           If Trim(cuenta) = "" Then
             MsgBox "Definir Cuenta Contable a : " & ccm_llave!CCM_CODBAN & " " & ccm_llave!CCM_nombre
           End If
        End If
     End If
    Return
    
Termino:
    cmdTransferir.Enabled = True
     cmdSalir.Enabled = True
    Exit Sub

Handler:
     MsgBox Err.Description, vbCritical, Pub_Titulo
     con_llave.Close
     CN.Execute "Rollback Transaction", rdExecDirect
     cmdTransferir.Enabled = True
     cmdSalir.Enabled = True
     Barra.Value = 0
End Sub

Private Sub Form_Load()
    CenterMe Me
    If cop_llave.EOF Then
      MsgBox "Definir Parametros en Contabilidad... ", 48, Pub_Titulo
      Exit Sub
    End If
    FECHA1.Text = Format(LK_FECHA_COP1, "dd/mm/yyyy")
    FECHA1.Mask = "##/##/####"
    fecha2.Text = Format(LK_FECHA_COP2, "dd/mm/yyyy")
    fecha2.Mask = "##/##/####"
    
    pub_cadena = "SELECT Max(MOV_NRO_VOUCHER) as NroVoucher FROM MOVICONT WHERE MOV_CODCIA = ? AND MOV_NRO_MES = ? AND MOV_TIPMOV = ? AND MOV_PERIODO = ? "
    Set PSMOV_VOU = CN.CreateQuery("", pub_cadena)
    PSMOV_VOU(0) = ""
    PSMOV_VOU(1) = 0
    PSMOV_VOU(2) = 0
    PSMOV_VOU(3) = 0
    Set VOU_MOV = PSMOV_VOU.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

    pub_cadena = "SELECT * FROM MOVICONT WHERE MOV_CODCIA = ? AND MOV_FECHA =? AND  MOV_NRO_MES = ? AND MOV_SERIE = ? AND MOV_NUMFAC = ? AND MOV_CP = ? AND MOV_CODCLIE = ? AND MOV_FBG = ? AND MOV_CODTRA = ? AND MOV_OPERACION = ? "
    Set PsMovicont = CN.CreateQuery("", pub_cadena)
    PsMovicont(0) = ""
    PsMovicont(1) = Date
    PsMovicont(2) = 0
    PsMovicont(3) = 0
    PsMovicont(4) = 0
    PsMovicont(5) = ""
    PsMovicont(6) = 0
    PsMovicont(7) = ""
    PsMovicont(8) = 0
    PsMovicont(9) = 0
    Set RsMovicont = PsMovicont.OpenResultset(rdOpenKeyset, rdConcurValues)
    
    pub_cadena = "SELECT * FROM ALLOG WHERE ALL_FLAG_EXT <> 'E' AND ALL_CODCIA = ? AND ALL_FECHA_DIA >= ? AND ALL_FECHA_DIA <= ? AND ALL_CODTRA = ? AND ALL_SECUENCIA = ?  ORDER BY ALL_FECHA_DIA, ALL_NUMOPER " 'AGREGUE MIC ALL_ESTADO <> 'E'
    Set PsAllog = CN.CreateQuery("", pub_cadena)
    PsAllog(0) = ""
    PsAllog(1) = Date
    PsAllog(2) = Date
    PsAllog(3) = 0
    PsAllog(4) = 0
    Set RsAllog = PsAllog.OpenResultset(rdOpenKeyset, rdConcurValues)
    
    pub_cadena = "SELECT * FROM COMAEST WHERE COM_CODCIA = ?  AND COM_CUENTA >= ? AND COM_NIVEL= ? ORDER BY COM_CUENTA"
    Set PsComaest = CN.CreateQuery("", pub_cadena)
    PsComaest(0) = ""
    PsComaest(1) = ""
    PsComaest(2) = 0
    Set RsComaest = PsComaest.OpenResultset(rdOpenKeyset, rdConcurValues)

End Sub
Private Function NroVoucher() As Integer
    PSMOV_VOU.rdoParameters(0) = LK_CODCIA
    PSMOV_VOU.rdoParameters(1) = LK_NRO_MES
    PSMOV_VOU.rdoParameters(2) = TipoMovimiento
    PSMOV_VOU.rdoParameters(3) = Format(RsAllog!ALL_FECHA_SUNAT, "yyyy")
    VOU_MOV.Requery
    If VOU_MOV.EOF Then
       NroVoucher = 1
    Else
       NroVoucher = IIf(IsNull(VOU_MOV!NroVoucher), 0, VOU_MOV!NroVoucher) + 1
    End If
    
End Function
