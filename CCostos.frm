VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmcospro 
   Caption         =   "Tipo de Cambio y Costeo de Articulos"
   ClientHeight    =   6990
   ClientLeft      =   2565
   ClientTop       =   1920
   ClientWidth     =   5580
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   5580
   Begin VB.CommandButton Command2 
      Caption         =   "Ce&rrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo de Cambios:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   5295
      Begin VB.TextBox fechafinal 
         Height          =   285
         Left            =   3240
         TabIndex        =   14
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin RichTextLib.RichTextBox TEXTOVAR 
         Height          =   375
         Left            =   1680
         TabIndex        =   8
         Top             =   1200
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   393217
         BackColor       =   12632064
         BorderStyle     =   0
         MultiLine       =   0   'False
         TextRTF         =   $"CCostos.frx":0000
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid gridigv 
         Height          =   3015
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "[Enter] = para Editar"
         Top             =   840
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   5318
         _Version        =   393216
         Rows            =   3
         FixedRows       =   2
         BackColorBkg    =   8421376
      End
      Begin MSMask.MaskEdBox txtfecha 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         Caption         =   "hasta"
         Height          =   255
         Left            =   2640
         TabIndex        =   15
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Digitar [enter] para modificar valores:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   "Mostar del :"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   240
      Top             =   6240
   End
   Begin VB.Frame Frame1 
      Caption         =   "Proceso de Costeo de Articulo:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   4
      Top             =   3960
      Width           =   5295
      Begin VB.TextBox TLINEA 
         Height          =   285
         Left            =   4680
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.TextBox TART 
         Height          =   288
         Left            =   3480
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   972
      End
      Begin VB.CheckBox CHEART 
         Caption         =   "Solo un Articulo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
      Begin ComctlLib.ProgressBar ProgBar 
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Visible         =   0   'False
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   0
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Costear Articulos"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Label LART 
         Caption         =   "Codigo Alterno:"
         Height          =   255
         Left            =   2280
         TabIndex        =   13
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label POR 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "Procesando.."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Visible         =   0   'False
         Width           =   4815
      End
   End
   Begin VB.Label lblbarraos 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "OSBusiness"
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
      TabIndex        =   17
      Top             =   6720
      Width           =   5655
   End
End
Attribute VB_Name = "Frmcospro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim temporal
Dim temfecha

Private Sub Command1_Click()
Dim WLINEA As Integer
Dim WCOSPRO_SUP As Currency
Dim WCODART As Currency
Dim FLAG_93 As String * 1
Dim wfec As String
Dim WS_SALDO As Currency
Dim WTEXTO As String
Dim CONTADOR As Integer
Dim wk_equiv As Currency
Dim indice As Integer
Dim WS_TIPO_CAMBIO As Double
Dim far_cambiox  As rdoResultset
Dim PSFAR_CAMBIO As rdoQuery
Dim PSFAR_COSTO As rdoQuery
Dim Far_Cost As rdoResultset
Dim ultimo_costo, ultimo_cambio As Currency
Dim PSFAR_COSTO2 As rdoQuery
Dim Far_Cost2 As rdoResultset
Dim Far_envio As rdoResultset
Dim PSFAR_ENVIO As rdoQuery
Dim PSFAR_COMPR As rdoQuery
Dim Far_Compr As rdoResultset
Dim PSFAR_COMPR2 As rdoQuery
Dim Far_Compr2 As rdoResultset
Dim WPASA_PRECIO As Currency
Dim WPASA_COSPRO_ANT As Currency
Dim ws_cospro As Currency
Dim ws_valor_ant As Currency
Dim ws_compra As Currency
Dim ws_ajuste, ws_perdida, WS_AJUSTE2 As Currency
Dim WS_NOMART As String * 22
Dim WR_IMPORTE1 As String * 12
Dim WR_IMPORTE2 As String * 12
Dim WR_IMPORTE3 As String * 12
Dim WR_IMPORTE4 As String * 12
Dim WR_IMPORTE5 As String * 12
Dim WR_IMPORTE6 As String * 12
Dim WR_IMPORTE7 As String * 12
Dim WR_IMPORTE8 As String * 12
Dim WS_VALOR1, ws_valor2, ws_Valor3, ws_valor4, ws_valor5, ws_valor6, ws_valor7 As Currency
Dim wCODARTI As Currency
Dim flagcito As String * 1
Dim WMONEDA As String

Dim PSFAR_OTRO As rdoQuery
Dim Far_otro As rdoResultset
If fechafinal.Visible Then
  If Not IsDate(fechafinal.Text) Then
    MsgBox "Fecha no procede.", 48, Pub_Titulo
    Exit Sub
   End If
End If
WCODART = 0
If CHEART.Value = 1 Then
    SQ_OPER = 3
    pu_alterno = Trim(TART.Text)
    pu_codcia = LK_CODCIA
    LEER_ART_LLAVE
    art_llave_alt.Requery
    If art_llave_alt.EOF Then
       MsgBox "Codigo no Existe / No Procede ", 48, Pub_Titulo
       Exit Sub
    End If
    WCODART = art_llave_alt!art_KEY
End If

If Not IsDate(txtfecha.Text) Then
  MsgBox "Fecha no Prcede. ", 48, Pub_Titulo
  Exit Sub
End If

If gridigv.Visible = False Then
  MsgBox "Determine fecha inicial y final . Ademas pulse <Enter> en la fecha final", 48, Pub_Titulo
  Exit Sub
End If

pub_mensaje = "Proceso de Costeo de Articulos de la Fecha : " & Chr(13) & " " & Format(txtfecha.Text, "dd/mm/yyyy") & " al " & Format(LK_FECHA_DIA, "dd/mm/yyyy") & Chr(13) & "¿ Continuar? ..."
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If
If Not CONSIS_TC(CDate(txtfecha.Text)) Then
  Exit Sub
End If


ProgBar.Visible = True
DoEvents
'POR(0).Visible = True
DoEvents
'POR(1).Visible = True
DoEvents
POR(2).Visible = True
DoEvents
ProgBar.Value = 0
ProgBar.Min = 0
POR(2).Caption = "Iniciando Proceso. . ."
ultimo_costo = 0
ultimo_cambio = 0
' proceso de consistencia de fechas con t.c.
FLAG_93 = ""
DoEvents
Command1.Enabled = False

pub_cadena = "SELECT FAR_COSTEO, FAR_COSTEO_REAL, FAR_PRECIO FROM FACART WHERE FAR_CODCIA = ? AND FAR_CODART = ? AND FAR_NUMGUIA = ? AND FAR_TIPMOV = 101  AND FAR_ESTADO<>'E' "
Set PSFAR_ENVIO = CN.CreateQuery("", pub_cadena)
PSFAR_ENVIO(0) = 0
PSFAR_ENVIO(1) = 0
PSFAR_ENVIO(2) = 0
Set Far_envio = PSFAR_ENVIO.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT FAR_CANTIDAD,FAR_COSTEO, FAR_COSTEO_REAL, FAR_PRECIO FROM FACART WHERE FAR_CODCIA = ? AND FAR_NUMSER = ? AND FAR_NUMFAC = ? AND FAR_TIPMOV = 93 AND FAR_SIGNO_ARM = 1 AND FAR_ESTADO <> 'E'  "
Set PSFAR_CAMBIO = CN.CreateQuery("", pub_cadena)
PSFAR_CAMBIO(0) = 0
PSFAR_CAMBIO(1) = 0
PSFAR_CAMBIO(2) = 0
Set far_cambiox = PSFAR_CAMBIO.OpenResultset(rdOpenKeyset, rdConcurValues)


pub_cadena = "SELECT FAR_CANTIDAD,FAR_COSTEO, FAR_COSTEO_REAL, FAR_PRECIO FROM FACART WHERE FAR_CODCIA = ? AND FAR_NUMSER = ? AND FAR_NUMFAC = ? AND FAR_TIPMOV = 93 AND FAR_SIGNO_ARM = 1 "
Set PSFAR_OTRO = CN.CreateQuery("", pub_cadena)
PSFAR_OTRO(0) = 0
PSFAR_OTRO(1) = 0
PSFAR_OTRO(2) = 0
Set Far_otro = PSFAR_OTRO.OpenResultset(rdOpenKeyset, rdConcurValues)



'pub_cadena = "SELECT FAR_CODART, SUM(FAR_CANTIDAD * FAR_SIGNO_ARM)  FROM FACART WHERE FAR_CODCIA = ? AND FAR_CODART = ? AND FAR_FECHA_COMPRA >= ?  AND FAR_ESTADO<>'E'   GROUP BY FAR_CODART"
'Set PSFAR_COSTO2 = CN.CreateQuery("", pub_cadena)
'PSFAR_COSTO2(0) = LK_CODCIA
'PSFAR_COSTO2(1) = 0
'PSFAR_COSTO2(2) = LK_FECHA_DIA
'Set Far_Cost2 = PSFAR_COSTO2.OpenResultset(rdOpenForwardOnly, rdConcurReadOnly)


'pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCIA = ? AND FAR_CODART = ? AND FAR_FECHA_COMPRA >= ? AND  FAR_FECHA_COMPRA <= ?   and far_estado<>'E' ORDER BY FAR_CODCIA, FAR_FECHA_COMPRA, FAR_SIGNO_ARM DESC , FAR_NUMOPER2"
pub_cadena = "SELECT FAR_NUM_LOTE,FAR_PRECIO_NETO,FAR_COSPRO_ANT,FAR_COSPRO_SUP, FAR_STOCK , FAR_TIPO_CAMBIO, FAR_COSTEO_REAL, FAR_NUMGUIA, FAR_NUMSER, FAR_NUMFAC, FAR_OTRA_CIA, FAR_CODART, FAR_COSPRO, FAR_FLETE, FAR_BRUTO, FAR_DESCTO, FAR_GASTOS, FAR_TOT_DESCTO, FAR_SIGNO_ARM, FAR_CANTIDAD, FAR_TIPMOV, FAR_MONEDA, FAR_FECHA_COMPRA, FAR_COSTEO, FAR_ESTADO, FAR_PRECIO, FAR_EQUIV FROM FACART WHERE FAR_CODCIA = ? AND FAR_CODART = ? AND FAR_FECHA_COMPRA >= ? AND  FAR_FECHA_COMPRA <= ?   and far_estado <>'E' ORDER BY FAR_CODCIA, FAR_FECHA_COMPRA, FAR_SIGNO_ARM DESC , FAR_NUMOPER2"
Set PSFAR_COSTO = CN.CreateQuery("", pub_cadena)
PSFAR_COSTO(0) = LK_CODCIA
PSFAR_COSTO(1) = 0
PSFAR_COSTO(2) = LK_FECHA_DIA
PSFAR_COSTO(3) = LK_FECHA_DIA
Set Far_Cost = PSFAR_COSTO.OpenResultset(rdOpenKeyset, rdConcurValues)

'¡DistinctRow
'DISTINCTROW

pub_cadena = "SELECT Distinct FAR_CODART FROM FACART WHERE FAR_CODCIA= ? AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ?  AND  ( FAR_COSTEO_REAL = 'A' OR FAR_COSTEO = 'A' )  AND FAR_ESTADO <> 'E'  ORDER BY FAR_CODART"
If CHEART.Value = 1 Then
  pub_cadena = "SELECT Distinct FAR_CODART FROM FACART WHERE FAR_CODCIA= ? AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ? AND FAR_CODART= " & WCODART & "  AND  ( FAR_COSTEO_REAL = 'A'  OR FAR_FECHA_COMPRA <> FAR_FECHA  OR FAR_MONEDA= 'D' )  AND FAR_ESTADO<>'E'  ORDER BY FAR_CODART"
End If
If Val(TLINEA.Text) <> 0 Then
  WLINEA = Val(TLINEA.Text)
  ' PRUEBA
  pub_cadena = "SELECT Distinct FAR_CODART FROM FACART, ARTI WHERE (ART_CODCIA = FAR_CODCIA) AND (FAR_CODART = ART_KEY) AND ART_LINEA = " & WLINEA & " AND FAR_CODCIA= ? AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ? AND ( FAR_COSTEO_REAL = 'A'  OR FAR_FECHA_COMPRA <> FAR_FECHA  OR FAR_MONEDA= 'D' )  AND FAR_ESTADO<>'E'  ORDER BY FAR_CODART"
End If
' TEMPORAL
''''pub_cadena = "SELECT Distinct FAR_CODART FROM FACART, ARTI WHERE (ART_CODCIA = FAR_CODCIA) AND (FAR_CODART = ART_KEY) AND  FAR_TIPMOV = 6 AND FAR_NUM_LOTE <> 10 AND FAR_CODCIA= ? AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ? AND ( FAR_COSTEO_REAL = 'A'  OR FAR_FECHA_COMPRA <> FAR_FECHA  OR FAR_MONEDA= 'D' )  AND FAR_ESTADO<>'E'  ORDER BY FAR_CODART"
Set PSFAR_COMPR = CN.CreateQuery("", pub_cadena)
PSFAR_COMPR(0) = LK_CODCIA
PSFAR_COMPR(1) = LK_FECHA_DIA
PSFAR_COMPR(2) = LK_FECHA_DIA
Set Far_Compr = PSFAR_COMPR.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
PSFAR_COMPR.rdoParameters(0) = LK_CODCIA
PSFAR_COMPR.rdoParameters(1) = CDate(txtfecha.Text)
If fechafinal.Visible Then
 PSFAR_COMPR.rdoParameters(2) = fechafinal.Text
Else
 PSFAR_COMPR.rdoParameters(2) = LK_FECHA_DIA
End If


WTEXTO = ""
CALCULA_93:

Far_Compr.Requery
ProgBar.Min = 0
ProgBar.Value = 0
If Not Far_Compr.EOF Then ProgBar.Max = Far_Compr.RowCount ''+ 2

'Do Until Far_Compr.EOF
'F1 = F1 + 1
'xl.Cells(F1, 1) = Far_Compr!far_codart

'Loop
'
If Far_Compr.EOF Then
  GoTo pasa
   MsgBox "No Movimientos para Costear....", 48, Pub_Titulo
    GoTo flag_avanza
   Unload Frmcospro
   Unload FORMGEN
   Exit Sub
End If
'OPEN_LOG "Proceso de Costeo"

POR(2).Caption = "Actualizando Costos. . ."
DoEvents
wCODARTI = Far_Compr!far_codart
flagcito = "A"
WPASA_COSPRO_ANT = 0
WCOSPRO_SUP = 0
Do Until Far_Compr.EOF
' If Far_Compr.AbsolutePosition < 469 Then GoTo OTRO_ART
   ProgBar.Value = Far_Compr.AbsolutePosition
   POR(2).Caption = "Costeando... " & ProgBar.Value & " de " & ProgBar.Max
   DoEvents
   'PSFAR_COSTO2.rdoParameters(0) = LK_CODCIA
   'PSFAR_COSTO2.rdoParameters(1) = Far_Compr!far_codart
   'PSFAR_COSTO2.rdoParameters(2) = CDate(txtfecha.Text)
   'Far_Cost2.Requery
   PSFAR_COSTO.rdoParameters(0) = LK_CODCIA
   PSFAR_COSTO.rdoParameters(1) = Far_Compr!far_codart
   PSFAR_COSTO.rdoParameters(2) = CDate(txtfecha.Text)
   If fechafinal.Visible Then
     PSFAR_COSTO.rdoParameters(3) = fechafinal.Text
   Else
     PSFAR_COSTO.rdoParameters(3) = LK_FECHA_DIA
   End If


   Far_Cost.Requery
   If Far_Cost.EOF Then
      MsgBox "Ver Mensaje : entre tablas facart : " & Far_Compr!far_codart
      End
   End If
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   PUB_KEY = Far_Cost!far_codart
   LEER_ART_LLAVE
   If art_LLAVE.EOF Then
      MsgBox "ERROR GRAVE... "
      End
   End If
   SQ_OPER = 1
   pu_codcia = LK_CODCIA
   PUB_CODART = Far_Cost!far_codart
 '   If PUB_CODART = 17484 Then Stop
   LEER_ARM_LLAVE
   If arm_llave.EOF Then
      MsgBox "ERROR GRAVE...arm "
      End
   End If

   WS_NOMART = art_LLAVE!art_nombre
   WPASA_COSPRO_ANT = Nulo_Valor0(Far_Cost!FAR_COSPRO_ANT)
   
   If LK_EMP = "HER" Or LK_EMP = "PIU" Then
       ws_cospro = Far_Cost!FAR_COSPRO
   End If
   
   If CDate(txtfecha.Text) = GEN!GEN_FECHA_REF Then
      If Far_Cost!far_signo_arm = -1 Then
        WS_SALDO = (0 + (Far_Cost!far_signo_arm * Far_Cost!far_cantidad) * -1)
      Else
        WS_SALDO = (0 + ((Far_Cost!far_signo_arm * Far_Cost!far_cantidad) * -1)) '' + arm_llave!ARM_STOCK - Far_Cost2(1)
      End If
      If Far_Cost!FAR_costeo = "A" And Far_Cost!far_estado <> "E" And Far_Cost!far_signo_arm = 1 Then
      Else
      WS_SALDO = 0
      End If
   Else
      If Far_Cost!far_signo_arm = -1 Then
        WS_SALDO = (Nulo_Valor0(Far_Cost!FAR_STOCK) + ((Far_Cost!far_signo_arm * Far_Cost!far_cantidad) * -1))
      Else
        WS_SALDO = (Nulo_Valor0(Far_Cost!FAR_STOCK) + ((Far_Cost!far_signo_arm * Far_Cost!far_cantidad) * -1))
      End If
   End If
   If WS_SALDO <= 0 Then WS_SALDO = 0
 '  Far_Cost.Requery
   ProgBar.Value = 0
   ProgBar.Max = Far_Cost.RowCount
   If Far_Cost!FAR_costeo = "A" And Far_Cost!far_estado <> "E" And Far_Cost!far_signo_arm = 1 And LK_EMP <> "PIU" Then
      WCOSPRO_SUP = Nulo_Valor0(Far_Cost!FAR_COSPRO_SUP)
      ws_cospro = Nulo_Valor0(Far_Cost!FAR_COSPRO_SUP)
   Else
      WCOSPRO_SUP = Nulo_Valor0(Far_Cost!FAR_COSPRO)
      ws_cospro = Nulo_Valor0(Far_Cost!FAR_COSPRO)
   End If
   Do Until Far_Cost.EOF
      ProgBar.Value = Far_Cost.AbsolutePosition
      
      DoEvents
      WS_TIPO_CAMBIO = 1
      WS_SALDO = WS_SALDO + (Far_Cost!far_signo_arm * Far_Cost!far_cantidad)
      If art_LLAVE!ART_MONEDA = "D" And Far_Cost!FAR_MONEDA = "D" Then
      Else
        If Far_Cost!FAR_MONEDA = "D" Then WS_TIPO_CAMBIO = JALAR(Far_Cost!FAR_fecha_compra, Far_Cost!FAR_fecha_compra)
      End If
      Far_Cost.Edit
      If Far_Cost!FAR_costeo = "A" And Far_Cost!far_estado <> "E" And Far_Cost!far_signo_arm = 1 Then
      ' es temporal para regulizar costos
        ''If Far_Cost!FAR_TIPMOV = 6 And Far_Cost!FAR_PRECIO <> 0 And Far_Cost!FAR_NUM_LOTE <> 10 Then
        '' Far_Cost!FAR_PRECIO = ws_cospro
        '' GoTo NADA
        ''End If
         WPASA_COSPRO_ANT = ws_cospro
         wk_equiv = Far_Cost!FAR_equiv
         'If Far_Cost!FAR_TIPMOV = 101 And Far_Cost!far_cantidad = 100 And Far_Cost!far_codart = 79119 Then Stop
         'WS_VALOR1 = (WS_SALDO - Far_Cost!FAR_CANTIDAD)
         'ws_valor2 = ws_cospro
         ws_valor_ant = (WS_SALDO - Far_Cost!far_cantidad * wk_equiv) * ws_cospro
         'ws_Valor3 = ws_valor_ant
         WS_IGV = 0
         ws_compra = WS_TIPO_CAMBIO * Far_Cost!far_cantidad * Far_Cost!FAR_PRECIO / wk_equiv
         ws_ajuste = 0
         If LK_EMP = "3AA" Then
           WS_AJUSTE2 = Nulo_Valor0(Far_Cost!FAR_FLETE)
         Else
           WS_AJUSTE2 = Nulo_Valor0(Far_Cost!FAR_FLETE) * WS_TIPO_CAMBIO
         End If
         WS_AJUSTE2 = WS_AJUSTE2 - (Nulo_Valor0(Far_Cost!FAR_DESCTO) * WS_TIPO_CAMBIO)
         ws_valor4 = Far_Cost!far_cantidad
         ws_compra = (ws_compra + ws_ajuste) + WS_AJUSTE2
         'ws_valor5 = ws_compra
         Far_Cost!far_precio_neto = ws_compra
         ultimo_costo = Far_Cost!FAR_PRECIO
         ultimo_cambio = Far_Cost!FAR_tipo_cambio
         If ws_valor_ant < 0 And LK_EMP <> "3AA" Then
            ws_valor_ant = 0
         End If
         'ws_valor6 = (ws_valor_ant + ws_compra)
         If WS_SALDO <> 0 Then
            ws_cospro = (ws_valor_ant + ws_compra) / WS_SALDO
         Else
            ws_cospro = (ws_valor_ant + ws_compra) / Far_Cost!far_cantidad
         End If
         'ws_valor7 = ws_cospro
         If Far_Cost!FAR_COSPRO <> ws_cospro Then
'           Print Far_Cost!FAR_fecha_compra
'           Print Far_Cost!far_codart
          If Far_Cost!FAR_fecha_compra <> "20/06/01" Then
              If Nulo_Valors(art_LLAVE!ART_MONEDA) = "S" Then
                WMONEDA = "S/. "
              Else
                WMONEDA = "US$. "
             End If
           End If
            WTEXTO = WTEXTO + Trim(art_LLAVE!art_nombre) & Chr(13) & "Costo Anterior: " & WMONEDA & Format(Far_Cost!FAR_COSPRO, "0.0000") & Chr(13) & "Nuevo Costo: " & WMONEDA & Format(ws_cospro, "0.0000") & Chr(13)
         End If
         If LK_EMP = "HER" Then
          '  ws_cospro = Val(Format((Far_Cost!FAR_PRECIO / Far_Cost!FAR_equiv) * WS_TIPO_CAMBIO, "0.0000"))
         End If
      End If
NADA:
      Far_Cost!FAR_COSPRO_SUP = WCOSPRO_SUP
      Far_Cost!FAR_COSPRO = ws_cospro
      Far_Cost!FAR_COSPRO_ANT = WPASA_COSPRO_ANT
      WCOSPRO_SUP = Far_Cost!FAR_COSPRO
      If Far_Cost!FAR_TIPMOV = 100 Then
        Far_Cost!FAR_PRECIO = ws_cospro
      End If
      If Far_Cost!FAR_TIPMOV = 93 Then
        If Far_Cost!far_signo_arm = -1 Then
            Far_Cost!FAR_PRECIO = ws_cospro
            PSFAR_CAMBIO(0) = LK_CODCIA
            PSFAR_CAMBIO(1) = Far_Cost!far_numser
            PSFAR_CAMBIO(2) = Far_Cost!far_numfac
            far_cambiox.Requery
            If Not far_cambiox.EOF Then
                far_cambiox.Edit
                far_cambiox!FAR_PRECIO = Format(redondea(Far_Cost!far_cantidad * Far_Cost!FAR_COSPRO) / far_cambiox!far_cantidad, "0.0000")
                far_cambiox!FAR_costeo = "A"
                far_cambiox!FAR_COSTEO_REAL = "A"
                far_cambiox.Update
            End If
          End If
      End If
      If Far_Cost!FAR_TIPMOV = 100 Then
        'Far_Cost!FAR_PRECIO = ws_cospro
        PSFAR_ENVIO(0) = Far_Cost!far_otra_cia
        PSFAR_ENVIO(1) = Far_Cost!far_codart
        PSFAR_ENVIO(2) = Far_Cost!far_numguia
        Far_envio.Requery
        Do Until Far_envio.EOF
            Far_envio.Edit
            Far_envio!FAR_PRECIO = Far_Cost!FAR_COSPRO
            Far_envio!FAR_costeo = "A"
            Far_envio!FAR_COSTEO_REAL = "A"
            Far_envio.Update
        Far_envio.MoveNext
        Loop
      End If
      If WS_TIPO_CAMBIO <> 1 Then Far_Cost!FAR_tipo_cambio = WS_TIPO_CAMBIO
      Far_Cost!FAR_STOCK = WS_SALDO
   
      Far_Cost.Update
      Far_Cost.MoveNext
Loop
        POR(2).Caption = "Actualizando Costos de Articulo. . ."
        arm_llave.Edit
        arm_llave!arm_cospro = ws_cospro
        arm_llave!ARM_COSTO_ULT = redondea(ultimo_costo)
        arm_llave.Update
        indice = 0
        SQ_OPER = 2
        PUB_CODART = art_LLAVE!art_KEY
        pu_codcia = LK_CODCIA
        LEER_PRE_LLAVE
        If pre_mayor.EOF Then MsgBox "Error en Precios..."
            Do Until pre_mayor.EOF
              pre_mayor.Edit
              If ws_valor4 > 0 Then
                  pre_mayor!PRE_COSTO = Format(ws_compra * pre_mayor!PRE_EQUIV / ws_valor4, "0.0000")
              Else
                  pre_mayor!PRE_COSTO = Nulo_Valor0(pre_mayor!PRE_cosTO_ant)
              End If

              pre_mayor.Update
              pre_mayor.MoveNext
           Loop

OTRO_ART:
    ProgBar.Max = Far_Compr.RowCount
    Far_Compr.MoveNext
Loop
pasa:

If FLAG_93 <> "A" Then
 FLAG_93 = "A"
 pub_cadena = "SELECT Distinct FAR_CODART FROM FACART WHERE FAR_CODCIA= ? AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ? AND ((FAR_TIPMOV = 93 AND FAR_SIGNO_ARM = 1) OR FAR_TIPMOV = 100)  AND (FAR_COSTEO_REAL = 'A'  OR FAR_FECHA_COMPRA <> FAR_FECHA  OR FAR_MONEDA= 'D')  AND FAR_ESTADO <>'E'  ORDER BY FAR_CODART" ' , FAR_CODCIA,FAR_FECHA_COMPRA"
 If CHEART.Value = 1 Then
   pub_cadena = "SELECT Distinct FAR_CODART FROM FACART WHERE FAR_CODART =  " & WCODART & " AND FAR_CODCIA= ? AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ?  AND ((FAR_TIPMOV = 93 AND FAR_SIGNO_ARM = 1) OR FAR_TIPMOV = 100)  AND (FAR_COSTEO_REAL = 'A' )  AND FAR_ESTADO<>'E'  ORDER BY FAR_CODART" ' , FAR_CODCIA,FAR_FECHA_COMPRA"
 End If
 If Val(TLINEA.Text) <> 0 Then
  WLINEA = Val(TLINEA.Text)
  ' PRUEBA
  pub_cadena = "SELECT Distinct FAR_CODART FROM FACART, ARTI WHERE (ART_CODCIA = FAR_CODCIA) AND (FAR_CODART = ART_KEY) AND ART_LINEA = " & WLINEA & "  AND FAR_TIPMOV = 93 AND FAR_SIGNO_ARM = 1 AND  FAR_CODCIA= ? AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ? AND ( FAR_COSTEO_REAL = 'A'  OR FAR_FECHA_COMPRA <> FAR_FECHA  OR FAR_MONEDA= 'D' )  AND FAR_ESTADO<>'E'  ORDER BY FAR_CODART"
 End If


 Set PSFAR_COMPR = CN.CreateQuery("", pub_cadena)
 PSFAR_COMPR(0) = LK_CODCIA
 PSFAR_COMPR(1) = LK_FECHA_DIA
 PSFAR_COMPR(2) = LK_FECHA_DIA
 Set Far_Compr = PSFAR_COMPR.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
 PSFAR_COMPR.rdoParameters(0) = LK_CODCIA
 PSFAR_COMPR.rdoParameters(1) = CDate(txtfecha.Text)
 If fechafinal.Visible Then
   PSFAR_COMPR.rdoParameters(2) = fechafinal.Text
 Else
   PSFAR_COMPR.rdoParameters(2) = LK_FECHA_DIA
 End If
 GoTo CALCULA_93
End If



'ProgBar.Value = ProgBar.Value + 1
'CLOSE_LOG

POR(2).Caption = "Terminando Proceso. . ."
DoEvents
flag_avanza:
PUB_CODCIA = LK_CODCIA
SQ_OPER = 1
LEER_PAR_LLAVE
par_llave.Edit
par_llave!par_flag_cierre = 8
par_llave!par_flag_costos = 9
par_llave.Update
''ProgBar.Value = ProgBar.Value + 1
salida:
CHEART.Value = 0
If Trim(WTEXTO) <> "" Then
 MsgBox " P r o c e s o   d e   C o s t e o " & Chr(13) & WTEXTO, vbInformation, Pub_Titulo
End If
POR(2).Caption = ""
MsgBox "Proceso Terminado.", vbInformation, Pub_Titulo
Command1.Enabled = True
ProgBar.Visible = False
'Unload Costos
'Unload FORMGEN
SALE:
Exit Sub

PROCESO_RECECPION:
If LK_EMP_PTO = "A" Then
  If LK_CODCIA <> "00" Then
    SQ_OPER = 1
    PUB_CODCIA = "00"
    LEER_PAR_LLAVE
    If par_llave!par_flag_costos <> "9" Then
       MsgBox "Costear Primero en Cia Principal ..."
       Unload Frmcospro
       GoTo fin
    End If
End If
End If

pub_cadena = "SELECT * FROM FFFART WHERE FFF_CODCIA=?  ORDER BY FFF_CODCIA"
Set PSFAR_COSTO = CN.CreateQuery("", pub_cadena)
PSFAR_COSTO(0) = LK_CODCIA
Set Far_Cost = PSFAR_COSTO.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT FAR_COSPRO, FAR_PRECIO, FAR_COSTEO, FAR_COSTEO_REAL FROM FACART WHERE FAR_TIPMOV =? AND FAR_CODCIA= ? AND FAR_NUMSER = ? AND FAR_FBG = ? AND FAR_NUMFAC = ? AND FAR_NUMSEC = ?   ORDER BY FAR_CODCIA"
Set PSFAR_COMPR = CN.CreateQuery("", pub_cadena)
PSFAR_COMPR(0) = 0
PSFAR_COMPR(1) = LK_CODCIA
PSFAR_COMPR(2) = 0
PSFAR_COMPR(3) = " "
PSFAR_COMPR(4) = 0
PSFAR_COMPR(5) = 0
Set Far_Compr = PSFAR_COMPR.OpenResultset(rdOpenKeyset, rdConcurValues)

PSFAR_COSTO(0) = LK_CODCIA
Far_Cost.Requery
    
Do Until Far_Cost.EOF
    PSFAR_COMPR(0) = Far_Cost!FFF_TIPMOV_R
    PSFAR_COMPR(1) = Far_Cost!FFF_codcia_R
    PSFAR_COMPR(2) = Far_Cost!FFF_numser_R
    PSFAR_COMPR(3) = Far_Cost!FFF_fbg_R
    PSFAR_COMPR(4) = Far_Cost!FFF_NUMFAC_R
    PSFAR_COMPR(5) = Far_Cost!FFF_NUMSEC_R
    Far_Compr.Requery
    If Far_Compr.EOF Then
      MsgBox "Registro No puede Costear: TIPMOV:" & Far_Cost!FFF_TIPMOV_R & " CODCIA:" & Far_Cost!FFF_codcia_R & " NUMSER:" & Far_Cost!FFF_numser_R & " FBG:" & Far_Cost!FFF_fbg_R & " NUMFAC:" & Far_Cost!FFF_NUMFAC_R & " NUMSEC:" & Far_Cost!FFF_NUMSEC_R, 48, Pub_Titulo
      GoTo SIGUE_R
    End If
    WPASA_PRECIO = Far_Compr!FAR_COSPRO ' JALO SU COSTO PROMEDIO
    
    PSFAR_COMPR(0) = Far_Cost!FFF_TIPMOV
    PSFAR_COMPR(1) = Far_Cost!FFF_codcia
    PSFAR_COMPR(2) = Val(Far_Cost!FFF_numser)
    PSFAR_COMPR(3) = Far_Cost!FFF_fbg
    PSFAR_COMPR(4) = Far_Cost!FFF_NUMFAC
    PSFAR_COMPR(5) = Far_Cost!FFF_NUMSEC
    Far_Compr.Requery
    If Far_Compr.EOF Then
      MsgBox "Registro No puede Costear: TIPMOV:" & Far_Cost!FFF_TIPMOV & " CODCIA:" & Far_Cost!FFF_codcia & " NUMSER:" & Far_Cost!FFF_numser & " FBG:" & Far_Cost!FFF_fbg & " NUMFAC:" & Far_Cost!FFF_NUMFAC & " NUMSEC:" & Far_Cost!FFF_NUMSEC, 48, Pub_Titulo
      GoTo SIGUE_R
    End If
    Far_Compr.Edit
    Far_Compr!FAR_PRECIO = WPASA_PRECIO ' ASIGNO EL COSTO AL PRECIO
    Far_Compr!FAR_costeo = "A"
    Far_Compr!FAR_COSTEO_REAL = "A"
    Far_Compr.Update
    
SIGUE_R:
    Far_Cost.MoveNext
Loop

Return
fin:
ProgBar.Visible = False
POR(0).Visible = False
POR(1).Visible = False
POR(2).Visible = False
Command1.Enabled = True
End Sub

Private Sub Command2_Click()
Unload Frmcospro
End Sub


Private Sub CHEART_Click()
If CHEART.Value = 1 Then
  TART.Visible = True
  LART.Visible = True
  TART.SetFocus
Else
  TART.Visible = False
  LART.Visible = False
End If
End Sub

Private Sub Form_Activate()
If txtfecha.Visible Then
 txtfecha_KeyPress 13
 Azul2 txtfecha, txtfecha
End If
End Sub

Private Sub Form_Load()
CenterMe Frmcospro
If LK_CODUSU = "ADMIN" Then TLINEA.Visible = True
'LblFecha.Caption = Format(LK_FECHA_DIA, "dddd, d mmmm yyyy")
'empresa.Caption = Trim(par_llave!PAR_NOMBRE)
txtfecha.Text = "01/" & Format(DatePart("m", LK_FECHA_DIA), "00") & "/" & DatePart("yyyy", LK_FECHA_DIA)
txtfecha.Mask = "##/##/####"
txtfecha.Visible = True
txtfecha.TabIndex = 0
Muestra_tc txtfecha.Text
temfecha = LK_FECHA_DIA
'If LK_CODUSU = "ADMIN" Then
fechafinal.Visible = True
fechafinal.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
'End If
End Sub
Private Sub gridiGV_KeyPress(KeyAscii As Integer)
Dim a As Integer
Dim t, WC
Static CONS
If KeyAscii <> 13 Then Exit Sub

'If Trim(gridigv.TextMatrix(gridigv.Row, 9)) <> "8" Then
'  If Trim(gridigv.TextMatrix(gridigv.Row, 0)) = "" Then Exit Sub
'  If Trim(gridigv.TextMatrix(gridigv.Row, 1)) <> "" And gridigv.Col = 2 Or gridigv.Col = 3 Then GoTo leer
'  If Trim(gridigv.TextMatrix(gridigv.Row, 8)) <> "0" Then Exit Sub
'End If


'If gridigv.Col = 1 And WMODO = "I" Then
'   a = Val(gridigv.TextMatrix(gridigv.Row - 1, 0))
'   a = a + 1
'  gridigv.TextMatrix(gridigv.Row, 0) = a
'End If
'If WMODO = "I" Or WMODO = "C" Then
    TEXTOVAR.Left = gridigv.Left + gridigv.CellLeft
    TEXTOVAR.Width = gridigv.CellWidth
    TEXTOVAR.Height = gridigv.CellHeight
    TEXTOVAR.Top = gridigv.Top + gridigv.CellTop
    TEXTOVAR.Text = gridigv.TextMatrix(gridigv.Row, gridigv.Col)
    TEXTOVAR.Visible = True
    Azul3 TEXTOVAR, TEXTOVAR
    TEXTOVAR.SetFocus
'End If
End Sub

Private Sub gridiGV_KeyUp(KeyCode As Integer, Shift As Integer)
Dim WC
Dim a, WF As Integer
Dim tf, t, tC
Dim SALE As Boolean
Dim Wsec

'If WMODO = "C" Then Exit Sub

'If cop_llave!COP_FLAG_MAYORIZACION = "M" Then
 'MsgBox "Ojo estaba Mayorizado..."
'End If


If Left(gridigv.TextMatrix(gridigv.Row, 0), 2) <> "MA" Then Exit Sub
 If KeyCode = 32 Then
  'If WMODO <> "C" Then Exit Sub
  tC = gridigv.Col
  For fila = 1 To gridigv.Cols - 1
      gridigv.Col = fila
      If gridigv.CellBackColor = QBColor(12) Then
         gridigv.CellBackColor = QBColor(15)
         gridigv.TextMatrix(gridigv.Row, 9) = "9"
      Else
         gridigv.CellBackColor = QBColor(12)
         gridigv.TextMatrix(gridigv.Row, 9) = "-1"
      End If
  Next fila
  gridigv.Col = tC
  gridigv.SetFocus
  Exit Sub
End If
If KeyCode = 45 Then
    Wsec = Wsec + 1
    If Trim(gridigv.TextMatrix(gridigv.Row + 1, 11)) = "8" Then
         Exit Sub
    Else
      If Trim(gridigv.TextMatrix(gridigv.Row + 1, 0)) = "T" Then Exit Sub
    End If
    If Val(gridigv.TextMatrix(gridigv.Row, 4)) = 0 And Val(gridigv.TextMatrix(gridigv.Row, 5)) = 0 Then Exit Sub
    gridigv.AddItem "", gridigv.Row + 1
    gridigv.TextMatrix(gridigv.Row + 1, 0) = "MAN. " & Format(gridigv.TextMatrix(gridigv.Row, 10), "dd/mm/yyyy")
    gridigv.TextMatrix(gridigv.Row + 1, 6) = Wsec
    gridigv.TextMatrix(gridigv.Row + 1, 8) = gridigv.TextMatrix(gridigv.Row, 8)
    gridigv.TextMatrix(gridigv.Row + 1, 3) = gridigv.TextMatrix(gridigv.Row, 3)
    gridigv.TextMatrix(gridigv.Row + 1, 7) = gridigv.TextMatrix(gridigv.Row, 7)
    gridigv.TextMatrix(gridigv.Row + 1, 10) = gridigv.TextMatrix(gridigv.Row, 10)
    gridigv.TextMatrix(gridigv.Row + 1, 11) = "8"
    gridigv.Row = gridigv.Row + 1
    gridigv.Col = 1
    gridigv.SetFocus
End If
Exit Sub
If KeyCode = 46 Then
If gridigv.Rows <= 3 Then
Else
   pub_mensaje = MsgBox("Desea Quitar el Item de la Cuenta : " & Trim(gridigv.TextMatrix(gridigv.Row, 1)), vbYesNo + vbExclamation + vbDefaultButton2, Pub_Titulo)
   If pub_mensaje = vbNo Then
     gridigv.SetFocus
     Exit Sub
   Else
     gridigv.RowHeight(gridigv.Row) = 1
     gridigv.Row = gridigv.Row + 1
    
   'gridiGV.RemoveItem (gridiGV.Row)
   'gridiGV.Refresh
   gridigv.SetFocus
   End If
End If
End If
'gridiGV.SetFocus
Exit Sub



End Sub

Private Sub gridigv_Scroll()
TEXTOVAR.Visible = False
End Sub


Private Sub textovar_Change()
gridigv.Text = Format(TEXTOVAR.Text, "0.0000")
End Sub

Private Sub TEXTOVAR_GotFocus()
 temporal = gridigv.TextMatrix(gridigv.Row, gridigv.Col)
End Sub

Private Sub textovar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  TEXTOVAR.Text = temporal
  TEXTOVAR.Visible = False
  gridigv.SetFocus
  Exit Sub
End If
If gridigv.Col = 1 Then Consistencias gridigv, TEXTOVAR, KeyAscii
If gridigv.Col = 4 Then Consistencias gridigv, TEXTOVAR, KeyAscii
If gridigv.Col = 5 Or gridigv.Col = 6 Then Consistencias gridigv, TEXTOVAR, KeyAscii
If KeyAscii <> 13 Then
   GoTo fin
End If
If gridigv.Col = 1 Or gridigv.Col = 4 Then
  If Val(TEXTOVAR.Text) > 99 Then
    Azul3 TEXTOVAR, TEXTOVAR
    Exit Sub
  End If
End If

PUB_CAL_INI = gridigv.TextMatrix(gridigv.Row, 2)
PUB_CAL_FIN = gridigv.TextMatrix(gridigv.Row, 2)
pu_codcia = LK_CODCIA
SQ_OPER = 1
PUB_CODCIA = LK_CODCIA
LEER_CAL_LLAVE
cal_llave.Edit
If gridigv.Col = 4 Then
   cal_llave!cal_tipo_cambio = Val(TEXTOVAR.Text)
End If
If gridigv.Col = 1 Then
   cal_llave!CAL_TC_MERCA = Val(TEXTOVAR.Text)
   If Format(LK_FECHA_DIA, "dd/mm/yyyy") = Format(gridigv.TextMatrix(gridigv.Row, 0), "dd/mm/yyyy") Then
      LK_TIPO_CAMBIO = Val(TEXTOVAR.Text)
      'MDIForm1.StatusBar1.Panels(3).Text = "T.C.= S/. " + Format(LK_TIPO_CAMBIO, "0.0000")
   End If
End If
If gridigv.Col = 5 Then
   cal_llave!cal_tc_ingre = Val(TEXTOVAR.Text)
End If
If gridigv.Col = 6 Then
   cal_llave!cal_tc_salid = Val(TEXTOVAR.Text)
End If

cal_llave.Update
If gridigv.Row >= gridigv.Rows - 1 Then
Else
  gridigv.Row = gridigv.Row + 1
End If
gridigv.SetFocus
TEXTOVAR.Visible = False

fin:

End Sub


Private Sub Timer1_Timer()
'lblcierre.Visible = Not lblcierre.Visible
End Sub

Public Sub Muestra_tc(wfecha_ini As Date)
Dim wdiaI, wdiaF As String
Dim wmesM As String
gridigv.Clear
gridigv.Cols = 7
gridigv.Rows = 2
gridigv.ColWidth(0) = 1000
gridigv.ColWidth(1) = 900
gridigv.ColWidth(2) = 0
gridigv.ColWidth(3) = 0
gridigv.ColWidth(4) = 900
gridigv.ColWidth(5) = 900
gridigv.ColWidth(6) = 900

gridigv.TextMatrix(0, 0) = "Fecha"
gridigv.TextMatrix(0, 1) = "Publicación"
gridigv.TextMatrix(1, 0) = "-"
gridigv.Row = 1
gridigv.Col = 1

gridigv.TextMatrix(1, 1) = "Compra"
gridigv.CellForeColor = QBColor(9)
gridigv.CellFontBold = True
gridigv.CellAlignment = 4
gridigv.TextMatrix(0, 4) = "---------"
gridigv.Col = 4
gridigv.TextMatrix(1, 4) = "Venta"
gridigv.CellForeColor = QBColor(4)
gridigv.CellFontBold = True
gridigv.CellAlignment = 4

gridigv.Col = 5
gridigv.TextMatrix(0, 5) = "Vigente"
gridigv.TextMatrix(1, 5) = "Ingresos"
gridigv.CellForeColor = QBColor(6)
gridigv.CellFontBold = True
gridigv.CellAlignment = 4
gridigv.Col = 6
gridigv.TextMatrix(0, 6) = "--------"
gridigv.TextMatrix(1, 6) = "Egresos"
gridigv.CellForeColor = QBColor(6)
gridigv.CellFontBold = True
gridigv.CellAlignment = 4



PUB_CAL_INI = wfecha_ini ' wdiaI & "/" & wmesM & "/" & PUB_CAL_ANO
PUB_CAL_FIN = LK_FECHA_DIA 'wdiaF & "/" & wmesM & "/" & PUB_CAL_ANO
pu_codcia = LK_CODCIA
PUB_CODCIA = LK_CODCIA
SQ_OPER = 1
LEER_CAL_LLAVE
cal_llave.MoveFirst
fila = 1
gridigv.Rows = 2
Do Until cal_llave.EOF
  fila = fila + 1
  gridigv.Rows = gridigv.Rows + 1
  gridigv.RowHeight(gridigv.Rows - 1) = 285
  gridigv.TextMatrix(fila, 0) = Format(cal_llave!CAL_FECHA, "dd/mm/yyyy")
  gridigv.TextMatrix(fila, 1) = Format(cal_llave!CAL_TC_MERCA, "0.0000")
  gridigv.TextMatrix(fila, 2) = Format(cal_llave!CAL_FECHA, "dd/mm/yyyy")
  gridigv.TextMatrix(fila, 4) = Format(cal_llave!cal_tipo_cambio, "0.0000")
  gridigv.TextMatrix(fila, 5) = Format(cal_llave!cal_tc_ingre, "0.0000")
  gridigv.TextMatrix(fila, 6) = Format(cal_llave!cal_tc_salid, "0.0000")
  cal_llave.MoveNext
Loop
gridigv.Visible = True
gridigv.Col = 1
If gridigv.Rows > 2 Then gridigv.Row = 2
If gridigv.Visible Then gridigv.SetFocus

End Sub

Public Function CONSIS_TC(wfecha_ini As Date) As Boolean
Dim WRES As String
PUB_CAL_INI = wfecha_ini ' wdiaI & "/" & wmesM & "/" & PUB_CAL_ANO
PUB_CAL_FIN = LK_FECHA_DIA
pu_codcia = LK_CODCIA
PUB_CODCIA = LK_CODCIA
SQ_OPER = 1
LEER_CAL_LLAVE
WRES = "Las Siguientes Fechas no tienen tipo de cambio:" & Chr(13)
fila = 0
Do Until cal_llave.EOF
If Nulo_Valor0(cal_llave!cal_tipo_cambio) = 0 And cal_llave!CAL_LABORABLE = "S" Then
  WRES = WRES + Format(cal_llave!CAL_FECHA, "dd/mm/yyyy") + " "
  fila = 1
End If
cal_llave.MoveNext
Loop
WRES = WRES + "." + Chr(13) + "Consulte su tabla de tipo de cambios"
If fila = 1 Then
  MsgBox WRES, 48, Pub_Titulo
  Muestra_tc wfecha_ini
  CONSIS_TC = False
Else
   CONSIS_TC = True
End If

End Function
Private Sub Consistencias(wsGrid As MSFlexGrid, wsTexto As RichTextBox, wsKeyAscii As Integer)
  Static VALOR
  Dim car As String
 ' NUMEROS CON DECIMALES
    car = Chr$(wsKeyAscii)
    car = UCase$(Chr$(wsKeyAscii))
    wsKeyAscii = Asc(car)
    If wsKeyAscii = 45 Then
      If wsTexto.Text <> "" Then
         Beep
         wsKeyAscii = 0
         Exit Sub
      End If
    End If
    If wsKeyAscii = 46 Then
      If InStr(1, wsTexto.Text, ".") <> 0 Then
        Beep
        wsKeyAscii = 0
        Exit Sub
      End If
    End If
    
    If car < "0" Or car > "9" Then
      If wsKeyAscii <> 8 And wsKeyAscii <> 13 And car <> "." Then
          wsKeyAscii = 0
          Beep
          Exit Sub
        End If
    End If

End Sub

Private Sub txtfecha_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 And IsDate(txtfecha.Text) = True Then
   Muestra_tc txtfecha.Text
Else
   gridigv.Clear
End If
End Sub

Public Function JALAR(wfecha_ini As Date, wfecha_fin As Date) As Currency
PUB_CAL_INI = wfecha_ini
PUB_CAL_FIN = wfecha_fin
pu_codcia = LK_CODCIA
PUB_CODCIA = LK_CODCIA
SQ_OPER = 1
LEER_CAL_LLAVE
JALAR = cal_llave!cal_tipo_cambio
End Function

Private Sub txtfecha_LostFocus()
If txtfecha.Text = "" Then Exit Sub
If Not IsDate(txtfecha.Text) Then
  MsgBox "Fecha no procede. ", 48, Pub_Titulo
  Azul2 txtfecha, txtfecha
  Exit Sub
End If
If txtfecha.Text > LK_FECHA_DIA Then
  MsgBox "Fecha no Puede ser mayor a la del dia.", 48, Pub_Titulo
  Azul2 txtfecha, txtfecha
  Exit Sub
End If


End Sub


