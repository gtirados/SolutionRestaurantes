VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "Comctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "CRYSTL32.OCX"
Begin VB.Form FrmImpC1 
   Caption         =   "Reportes"
   ClientHeight    =   4380
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8610
   Icon            =   "FrmImpC1.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   8610
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmcuentas 
      Caption         =   "Opciones"
      Height          =   1095
      Left            =   120
      TabIndex        =   26
      Top             =   1200
      Visible         =   0   'False
      Width           =   5295
      Begin VB.TextBox a_cta2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   28
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox a_cta1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1320
         TabIndex        =   27
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "al "
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   30
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Cuentas del "
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   29
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.TextBox txtnivel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      MaxLength       =   1
      TabIndex        =   24
      Top             =   720
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.CheckBox periodos 
      Caption         =   "Acumular Periodos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   23
      Top             =   120
      Width           =   2655
   End
   Begin VB.ComboBox FECHA1 
      Height          =   315
      Left            =   5880
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   240
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CheckBox che1 
      Caption         =   "Incrementar a Diario"
      Height          =   195
      Left            =   480
      TabIndex        =   20
      Top             =   1200
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   4215
      Begin VB.Label lblreporte 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   4020
      End
   End
   Begin VB.Frame fcontab 
      Height          =   2775
      Left            =   3240
      TabIndex        =   10
      Top             =   600
      Visible         =   0   'False
      Width           =   5295
      Begin VB.OptionButton opnivel 
         Caption         =   "Nivel 6"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   19
         Top             =   2400
         Width           =   855
      End
      Begin VB.OptionButton opnivel 
         Caption         =   "Nivel 5"
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   17
         Top             =   2040
         Width           =   855
      End
      Begin VB.OptionButton opnivel 
         Caption         =   "Nivel 4"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   1680
         Width           =   855
      End
      Begin VB.OptionButton opnivel 
         Caption         =   "Nivel 3"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   855
      End
      Begin VB.OptionButton opnivel 
         Caption         =   "Nivel 2"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton opnivel 
         Caption         =   "Nivel 1"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   855
      End
      Begin VB.OptionButton opnivel 
         Caption         =   "Balance"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.ListBox listacta 
         Height          =   2310
         Left            =   1200
         Style           =   1  'Checkbox
         TabIndex        =   11
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label lblmensa 
         Caption         =   "Balance Pricipal"
         Height          =   255
         Left            =   2160
         TabIndex        =   18
         Top             =   1200
         Width           =   1215
      End
   End
   Begin VB.CommandButton pantalla 
      Caption         =   "Por &Pantalla .."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cerrar 
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
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   3960
      Width           =   1095
   End
   Begin ComctlLib.ProgressBar ProgBar 
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   3480
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
   Begin Crystal.CrystalReport Reportes 
      Left            =   10
      Top             =   3600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
   Begin MSMask.MaskEdBox txtCampo2 
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327680
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
   Begin MSMask.MaskEdBox txtCampo1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   327680
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
   Begin VB.Label lblnivel 
      Caption         =   "Nivel :"
      Height          =   255
      Left            =   360
      TabIndex        =   25
      Top             =   720
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblperiodos 
      Caption         =   "Periodo  :"
      Height          =   255
      Left            =   5880
      TabIndex        =   22
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblcampo2 
      AutoSize        =   -1  'True
      Caption         =   "Campo1"
      Height          =   195
      Left            =   1560
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label lblcampo1 
      AutoSize        =   -1  'True
      Caption         =   "Campo1"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1440
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label lblProceso 
      Alignment       =   2  'Center
      Caption         =   "Procesando ..."
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   3720
      Visible         =   0   'False
      Width           =   3375
   End
End
Attribute VB_Name = "FrmImpC1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LOC_FECHA_ULT As Date
Dim REP_FECHA1
Dim REP_FECHA2
Dim xl As Object
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset
Dim PS_REP02 As rdoQuery
Dim llave_rep02 As rdoResultset
Dim PS_REP03 As rdoQuery
Dim llave_rep03 As rdoResultset
Dim PS_REP04 As rdoQuery
Dim llave_rep04 As rdoResultset
Dim wranF, wran1, wran2, WPAS
Dim C1 As Integer
Dim F1 As Integer
Dim xcuenta As Integer
Dim i As Integer
Dim Mensaje, titulo, valorpred As String
Dim Wfile  As String
Dim WFORM  As String
Dim wCOM_NIVEL(6) As Integer
Dim NIVEL_MAX  As Integer
Dim PSCTA1 As rdoQuery
Dim loc_cta1 As rdoResultset
Dim PSCOH_LLAVE As rdoQuery
Dim coh_llave As rdoResultset


Private Sub a_cta1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   Azul a_cta2, a_cta2
End If
End Sub

Private Sub a_cta2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If pantalla.Enabled Then pantalla.SetFocus
End If

End Sub

Private Sub cerrar_Click()
Unload FrmImpC1
End Sub

Private Sub Form_Load()
CenterMe FrmImpC1
pub_cadena = "SELECT COV_FECHA_VOUCHER FROM COMOV WHERE COV_CODCIA = ? AND COV_FECHA_VOUCHER >= ? AND COV_FECHA_VOUCHER <= ?   ORDER BY COV_FECHA_VOUCHER DESC"
Set PS_REP04 = CN.CreateQuery("", pub_cadena)
PS_REP04.MaxRows = 1
PS_REP04(0) = 0
PS_REP04(1) = LK_FECHA_DIA
PS_REP04(2) = LK_FECHA_DIA
Set llave_rep04 = PS_REP04.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
If cop_llave.EOF Then
   MsgBox "Definir Parametros en Contabilidad .. NO Procede ", 48, Pub_Titulo
   Exit Sub
End If
PS_REP04(0) = LK_CODCIA
PS_REP04(1) = LK_FECHA_COP1
PS_REP04(2) = LK_FECHA_COP2
llave_rep04.Requery
If Not llave_rep04.EOF Then
    LOC_FECHA_ULT = llave_rep04!COV_FECHA_VOUCHER
Else
    LOC_FECHA_ULT = LK_FECHA_COP1
End If

Screen.MousePointer = 11
If tra_llave.EOF Then
   Screen.MousePointer = 0
   Exit Sub
End If
Screen.MousePointer = 0
Wfile = Trim(tra_llave(3))
WFORM = Trim(tra_llave(7))
lblreporte.Caption = Trim(tra_llave(1))
LLENA_COPARAN
If Wfile = "CTA_HISTORICO" Or Wfile = "LIBRO_MAYOR" Or Wfile = "BALANCE" Then
  lblcampo1.Visible = True
  lblcampo1.Caption = "Periodo del : " & Format(LK_FECHA_COP1, "dd/mm/yyyy") & " al  : " & Format(LK_FECHA_COP2, "dd/mm/yyyy")
  pub_cadena = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND COM_NIVEL = ? ORDER BY COM_CUENTA "
  Set PSCTA1 = CN.CreateQuery("", pub_cadena)
  PSCTA1(0) = 0
  PSCTA1(1) = 0
  Set loc_cta1 = PSCTA1.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
  fcontab.Visible = True
  opnivel(0).Value = True
  For fila = 0 To opnivel.Count - 1
   If fila >= NIVEL_MAX + 1 Then
     opnivel(fila).Enabled = False
   End If
  Next fila
  If Wfile = "CTA_HISTORICO" Or Wfile = "LIBRO_MAYOR" Then
    fcontab.Visible = True
    opnivel(0).Value = True
    For fila = 0 To opnivel.Count - 1
      If fila >= NIVEL_MAX + 1 Then
        opnivel(fila).Enabled = False
      End If
    Next fila
    opnivel(0).Enabled = False
    opnivel(1).Value = True
    lblcampo1.Caption = "Fecha Inicial : "
    lblcampo1.Visible = True
    txtCampo1.Text = Format(LK_FECHA_COP1, "dd/mm/yyyy")
    txtCampo2.Text = Format(LK_FECHA_COP2, "dd/mm/yyyy")
    txtCampo1.Mask = "##/##/####"
    txtCampo1.Visible = True
    lblcampo2.Caption = "Fecha Final: "
    lblcampo2.Visible = True
    txtCampo2.Mask = "##/##/####"
    txtCampo2.Visible = True
    lblcampo2.Visible = True
    lblcampo2.Caption = "Fecha Final"
  End If
End If
If Wfile = "BAL_COMPRO" Then
 lblnivel.Visible = True
 txtnivel.Visible = True
 txtnivel.Text = "1"
 lblcampo1.Visible = True
 lblcampo1.Caption = "Periodo del : " & Format(LK_FECHA_COP1, "dd/mm/yyyy") & " al  : " & Format(LK_FECHA_COP2, "dd/mm/yyyy")
End If
If Wfile = "A_CUENTAS" Then
  frmcuentas.Visible = True
  a_cta1.TabIndex = 0
End If

If Wfile = "ANALISIS" Then
  lblcampo1.Caption = "Fecha Inicial : "
  lblcampo1.Visible = True
  txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
  txtCampo1.Mask = "##/##/####"
  txtCampo1.Visible = True
  txtCampo1.TabIndex = 0
End If
If Wfile = "LIBRO_DIARIO" Or Wfile = "LIBRO_RESTO" Or Wfile = "LIBRO_CAJA" Then
  lblcampo1.Caption = "Fecha Inicial : "
  lblcampo1.Visible = True
  txtCampo1.Text = Format(LK_FECHA_COP1, "dd/mm/yyyy")
  txtCampo2.Text = Format(LK_FECHA_COP2, "dd/mm/yyyy")
  txtCampo1.Mask = "##/##/####"
  txtCampo1.Visible = True
  lblcampo2.Caption = "Fecha Final: "
  lblcampo2.Visible = True
  txtCampo2.Mask = "##/##/####"
  txtCampo2.Visible = True
  If LK_EMP <> "PIU" Then
    If Wfile <> "LIBRO_DIARIO" Then che1.Visible = True
  End If
  If Wfile = "LIBRO_DIARIO" Then
   lblcampo2.Visible = True
   lblcampo2.Caption = "DIARIO -  Periodo del : " & Format(LK_FECHA_COP1, "dd/mm/yyyy") & " al  : " & Format(LK_FECHA_COP2, "dd/mm/yyyy")
  End If
End If
If Wfile = "LIBRO_CAJA" Then
  che1.Visible = True
'  lblcampo1.Visible = True
'  lblcampo1.Caption = "Periodo del : " & Format(LK_FECHA_COP1, "dd/mm/yyyy") & " al  : " & Format(LK_FECHA_COP2, "dd/mm/yyyy")
End If
Exit Sub




End Sub

Private Sub listacta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If pantalla.Enabled Then pantalla.SetFocus
End If
End Sub

Private Sub listacta_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 113 Then
 For fila = 0 To listacta.ListCount - 1
   listacta.Selected(fila) = True
 Next fila
End If
If KeyCode = 114 Then
 For fila = 0 To listacta.ListCount - 1
   listacta.Selected(fila) = False
 Next fila
End If

End Sub

Private Sub opnivel_Click(Index As Integer)
If Index = 0 Then
  listacta.Visible = False
Else
  LLENA_CTA Index
End If
End Sub

Private Sub Pantalla_Click()
Dim wsFECHA1
Dim wsFECHA2
'On Error GoTo SALE

If Wfile = "BALANCE" Then
 Call BALANCE
End If
If Wfile = "BAL_COMPRO" Then
 Call BAL_COMPRO
End If
If Wfile = "ANALISIS" Then
  Call ANALISIS
End If
If Wfile = "LIBRO_CAJA" Then
 Call LIBRO_CAJA
End If
If Wfile = "LIBRO_RESTO" Then
 Call LIBRO_RESTO
End If
If Wfile = "LIBRO_DIARIO" Then
 Call LIBRO_DIARIO
End If
If Wfile = "LIBRO_MAYOR" Then
 Call LIBRO_MAYOR
End If
If Wfile = "CTA_HISTORICO" Then
Call CTA_HISTORICO
End If
If Wfile = "A_CUENTAS" Then
 Call A_CUENTAS
End If

If Wfile = "ESTADO1" Then
  POWER_REPORT 77
End If
If Wfile = "ESTADO2" Then
  POWER_REPORT 78
End If

Exit Sub
sale:
ProgBar.Visible = False
lblProceso.Visible = False
pantalla.Enabled = True
cerrar.Enabled = True
MsgBox Err.Description + "Intente Nuevamente.", 48, Pub_Titulo
End Sub

Private Sub periodos_Click()
If periodos.Value = 1 Then
  'lblperiodos.Visible = True
  'FECHA1.Visible = True
  'LLENA_PERIODOS
Else
  'lblperiodos.Visible = False
  'FECHA1.Visible = False
End If
End Sub

Private Sub txtcampo1_GotFocus()
'Azul txtCampo1, txtCampo1
End Sub

Private Sub txtCampo1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
 Exit Sub
End If

If txtCampo2.Visible Then
 If Not IsDate(txtCampo2) Then
   txtCampo2.Text = Format(txtCampo1.Text, "dd/mm/yyyy")
 End If
 Azul2 txtCampo2, txtCampo2
Else
pantalla.SetFocus
End If
 

End Sub

Private Sub txtcampo2_GotFocus()
'Azul txtCampo2, txtCampo2
End Sub

Private Sub txtcampo2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
 Exit Sub
End If
If pantalla.Enabled Then
   pantalla.SetFocus
End If

End Sub
Public Sub LLENADOS(cont As ListBox, tip As Integer)
Dim CONTA As Integer
    CONTA = -1
    PUB_TIPREG = tip
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
'    cont.AddItem " "
    Do Until tab_mayor.EOF
        cont.AddItem tab_mayor!tab_nomlargo & String(60, " ") & tab_mayor!TAB_NUMTAB
        CONTA = CONTA + 1
        tab_mayor.MoveNext
    Loop
End Sub
Public Sub LLENADOS_COMBO(cont As ComboBox, tip As Integer)
    PUB_TIPREG = tip
    SQ_OPER = 2
    LEER_TAB_LLAVE
    cont.ToolTipText = "TAB_TIPREG = " & tip
    cont.Clear
    Do Until tab_mayor.EOF
        cont.AddItem tab_mayor!tab_nomlargo & String(60, " ") & tab_mayor!TAB_NUMTAB
        tab_mayor.MoveNext
    Loop
End Sub


Public Sub BALANCE()
'On Error GoTo FINTODO
Dim CT_RESULTADO As String
Dim WCAMBIA
Dim ws_clave As String
Dim WSFECHA As Date
Dim wcodven As Integer
Dim wvalor
Dim Wche As Integer
Dim wkSELECT As String
Dim wsfile As String
Dim F2 As Integer
Dim saldos As Currency
Dim SALDO_TOTAL As Currency
Dim Wflag As String * 1
Dim WCOL1 As Integer
Dim WCOL2 As Integer
Dim SALDO_COL1 As Currency
Dim SALDO_COL2 As Currency
Dim wsaldo_resultado As Currency
Dim wtipcta
Dim CARAC As String
Dim saldo As Currency
Dim total As Currency
Dim wfi As Integer
If periodos.Value = 1 Then
' If Trim(fecha1.Text) = "" Then
'   MsgBox " Seleccione un periodo de la lista.", 48, Pub_Titulo
'   fecha1.SetFocus
'   Exit Sub
' End If
End If
CT_RESULTADO = "89"
wsfile = ""
pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImpC1.lblProceso.Caption = "Activando Reporte... un Momento ."
DoEvents

If opnivel(0).Value = True Then
Else
  GoTo AUXILIAR
  Exit Sub
End If

'If periodos.Value = 1 Then
' pub_cadena = "SELECT * FROM COHMAEST WHERE (COH_FECHA_PROCESO >= ? AND COH_FECHA_PROCESO2 <= ? )AND COH_CODCIA = ? AND COH_NIVEL = ? AND ( COH_TIPO_CTA = ? OR COH_TIPO_CTA = ? )  ORDER BY COH_CODCIA, COH_TIPO_CTA , COH_CUENTA"
'Else
 pub_cadena = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND COM_NIVEL = ? AND ( COM_TIPO_CTA = ? OR COM_TIPO_CTA = ? )  ORDER BY COM_CODCIA, COM_TIPO_CTA , COM_CUENTA"
'End If
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
'If periodos.Value = 1 Then
'  PS_REP01(0) = LK_FECHA_DIA
'  PS_REP01(1) = LK_FECHA_DIA
'  PS_REP01(2) = 0
'  PS_REP01(3) = 0
'  PS_REP01(4) = 0
'  PS_REP01(5) = 0
'Else
  PS_REP01(0) = 0
  PS_REP01(1) = 0
  PS_REP01(2) = 0
  PS_REP01(3) = 0
'End If
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
DoEvents

'*** VERFICA SI HAY DATOS , O ESTAN CORRECTOS
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = 1
PS_REP01(2) = 1
PS_REP01(3) = 3
llave_rep01.Requery
If llave_rep01.EOF = True Then
   MsgBox "!!! NO EXISTEN Datos ...", 48, Pub_Titulo
   GoTo CANCELA
End If
' Proceso de Resultado del Ejercicio variable a devolver "saldo"
GoSub RESULTADO
wsaldo_resultado = saldo
ws_clave = PUB_CLAVE
FrmImpC1.ProgBar.Min = 0
FrmImpC1.ProgBar.Max = llave_rep01.RowCount
FrmImpC1.ProgBar.Value = 0
FrmImpC1.lblProceso.Visible = True
FrmImpC1.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL

'GoSub LETRAS

FrmImpC1.ProgBar.Visible = True
DoEvents
xcuenta = 0
F2 = 5
F1 = 5  'Fila Inicial
C1 = 1
FrmImpC1.lblProceso.Caption = "Procesando . . .  un Momento ."
DoEvents
fila = 0
WCAMBIA = llave_rep01!com_tipo_cta
SALDO_TOTAL = 0
saldos = 0
Wflag = ""
Do Until llave_rep01.EOF
   FrmImpC1.ProgBar.Value = FrmImpC1.ProgBar.Value + 1
   wtipcta = llave_rep01!com_tipo_cta
   If WCAMBIA <> wtipcta Then
     Wflag = "A"
     WCAMBIA = llave_rep01!com_tipo_cta
     F1 = F1 + 1
     xl.Cells(F1, C1 + 1) = "TOTAL ACTIVO CORRIENTE = "
     xl.Cells(F1, C1 + 2) = saldos
     xl.Cells(F1, C1 + 2).Borders.Item(xlEdgeTop).LineStyle = 1
     SALDO_TOTAL = SALDO_TOTAL + saldos
     F1 = F1 + 1
     saldos = 0
   End If
   F1 = F1 + 1
   xl.Cells(F1, C1) = Trim(llave_rep01!com_cuenta)
   xl.Cells(F1, C1 + 1) = Trim(llave_rep01!com_DESCRIPCION)
   JALA_SALDO llave_rep01!com_cuenta, periodos.Value
   'saldos = saldos + ((Val(llave_rep01!COM_DEB_ANO) + Val(llave_rep01!COM_DEB_MES)) * llave_rep01!com_SIGNO_D) + ((Val(llave_rep01!COM_HAB_ANO) + Val(llave_rep01!COM_HAB_MES)) * llave_rep01!com_SIGNO_H)
   saldos = saldos + ((PUB_IMPORTE_DEB) * llave_rep01!com_signo_d) + ((PUB_IMPORTE_HAB) * llave_rep01!com_signo_h)
   xl.Cells(F1, C1 + 2) = ((PUB_IMPORTE_DEB) * llave_rep01!com_signo_d) + ((PUB_IMPORTE_HAB) * llave_rep01!com_signo_h)
   llave_rep01.MoveNext
Loop
F1 = F1 + 1
If Wflag = "A" Then
  xl.Cells(F1, C1 + 1) = "TOTAL ACTIVO NO CORRIENTE = "
Else
  xl.Cells(F1, C1 + 1) = "TOTAL ACTIVO CORRIENTE = "
End If
xl.Cells(F1, C1 + 2) = saldos
xl.Cells(F1, C1 + 2).Borders.Item(xlEdgeTop).LineStyle = 1

SALDO_TOTAL = SALDO_TOTAL + saldos
F1 = F1 + 2
WCOL1 = F1
SALDO_COL1 = SALDO_TOTAL
     
C1 = 5
F1 = 5
PS_REP01(2) = 2
PS_REP01(3) = 4
llave_rep01.Requery
If llave_rep01.EOF = True Then
   MsgBox "!!! NO EXISTEN Datos ...", 48, Pub_Titulo
   GoTo CANCELA
End If
SALDO_TOTAL = 0
saldos = 0
WCAMBIA = llave_rep01!com_tipo_cta
Wflag = ""
Do Until llave_rep01.EOF
   wtipcta = llave_rep01!com_tipo_cta
   If WCAMBIA <> wtipcta Then
     Wflag = "A"
     WCAMBIA = llave_rep01!com_tipo_cta
     F1 = F1 + 1
     xl.Cells(F1, C1 + 1) = "TOTAL PASIVO CORRIENTE = "
     xl.Cells(F1, C1 + 2) = saldos
     xl.Cells(F1, C1 + 2).Borders.Item(xlEdgeTop).LineStyle = 1
     SALDO_TOTAL = SALDO_TOTAL + saldos
     F1 = F1 + 1
     saldos = 0
   End If
   F1 = F1 + 1
   xl.Cells(F1, C1) = Trim(llave_rep01!com_cuenta)
   xl.Cells(F1, C1 + 1) = Trim(llave_rep01!com_DESCRIPCION)
   JALA_SALDO llave_rep01!com_cuenta, periodos.Value
   'saldos = saldos + ((Val(llave_rep01!COM_DEB_ANO) + Val(llave_rep01!COM_DEB_MES)) * llave_rep01!com_SIGNO_D) + ((Val(llave_rep01!COM_HAB_ANO) + Val(llave_rep01!COM_HAB_MES)) * llave_rep01!com_SIGNO_H)
   saldos = saldos + ((PUB_IMPORTE_DEB) * llave_rep01!com_signo_d) + ((PUB_IMPORTE_HAB) * llave_rep01!com_signo_h)
   xl.Cells(F1, C1 + 2) = ((PUB_IMPORTE_DEB) * llave_rep01!com_signo_d) + ((PUB_IMPORTE_HAB) * llave_rep01!com_signo_h)
   llave_rep01.MoveNext
Loop
F1 = F1 + 1
If Wflag = "A" Then
  xl.Cells(F1, C1 + 1) = "TOTAL PASIVO NO CORRIENTE = "
Else
  xl.Cells(F1, C1 + 1) = "TOTAL PASIVO CORRIENTE = "
End If
xl.Cells(F1, C1 + 2) = saldos
xl.Cells(F1, C1 + 2).Borders.Item(xlEdgeTop).LineStyle = 1
SALDO_TOTAL = SALDO_TOTAL + saldos
F1 = F1 + 1
xl.Cells(F1, C1 + 1) = "TOTAL PASIVO = "
xl.Cells(F1, C1 + 2) = SALDO_TOTAL
xl.Cells(F1, C1 + 2).Borders.Item(xlEdgeTop).LineStyle = 1

C1 = 5
'F1 = 5
F1 = F1 + 1
PS_REP01(2) = 5
PS_REP01(3) = 5
llave_rep01.Requery
If llave_rep01.EOF = True Then
   MsgBox "!!! NO EXISTEN Datos ...", 48, Pub_Titulo
   GoTo CANCELA
End If
saldos = 0
Do Until llave_rep01.EOF
'   FrmImpC1.ProgBar.Value = FrmImpC1.ProgBar.Value + 1
  F1 = F1 + 1
  xl.Cells(F1, C1) = Trim(llave_rep01!com_cuenta)
  xl.Cells(F1, C1 + 1) = Trim(llave_rep01!com_DESCRIPCION)
  If Trim(llave_rep01!com_cuenta) = CT_RESULTADO Then
    xl.Cells(F1, C1 + 2) = wsaldo_resultado
    saldos = saldos + wsaldo_resultado
  Else
   JALA_SALDO llave_rep01!com_cuenta, periodos.Value
'   saldos = saldos + ((PUB_IMPORTE_DEB) * llave_rep01!com_SIGNO_D) + ((PUB_IMPORTE_HAB) * llave_rep01!com_SIGNO_H)
   xl.Cells(F1, C1 + 2) = ((PUB_IMPORTE_DEB) * llave_rep01!com_signo_d) + ((PUB_IMPORTE_HAB) * llave_rep01!com_signo_h)
   saldos = saldos + ((PUB_IMPORTE_DEB) * llave_rep01!com_signo_d) + ((PUB_IMPORTE_HAB) * llave_rep01!com_signo_h)
  End If
   llave_rep01.MoveNext
Loop
F1 = F1 + 1


SQ_OPER = 1
PUB_CUENTA = "89"
PUB_CODCIA = LK_CODCIA
LEER_COM_LLAVE
If com_llave.EOF Then
   MsgBox "Cuenta Contable : " & PUB_CUENTA
End If
xl.Cells(F1, C1) = Trim(com_llave!com_cuenta)
xl.Cells(F1, C1 + 1) = Trim(com_llave!com_DESCRIPCION)
xl.Cells(F1, C1 + 2) = wsaldo_resultado
saldos = saldos + wsaldo_resultado


F1 = F1 + 1
xl.Cells(F1, C1 + 1) = "TOTAL PATRIMONIO = "
xl.Cells(F1, C1 + 2) = saldos
xl.Cells(F1, C1 + 2).Borders.Item(xlEdgeTop).LineStyle = 1
SALDO_TOTAL = SALDO_TOTAL + saldos
SALDO_COL2 = SALDO_TOTAL
F1 = F1 + 2
WCOL2 = F1
If WCOL1 > WCOL2 Then
 F1 = WCOL1
ElseIf WCOL1 = WCOL2 Then
 F1 = WCOL1
ElseIf WCOL1 < WCOL2 Then
 F1 = WCOL2
End If
C1 = 1
xl.Cells(F1, C1 + 1) = "TOTAL ACTIVO = "
xl.Cells(F1, C1 + 2) = SALDO_COL1
xl.Cells(F1, C1 + 2).Borders.Item(xlEdgeTop).LineStyle = 1
xl.Cells(F1, C1 + 2).Borders.Item(xlEdgeBottom).LineStyle = 1
C1 = 5
xl.Cells(F1, C1 + 1) = "TOTAL PASIVO Y PATRIMONIO = "
xl.Cells(F1, C1 + 2) = SALDO_COL2
xl.Cells(F1, C1 + 2).Borders.Item(xlEdgeTop).LineStyle = 1
xl.Cells(F1, C1 + 2).Borders.Item(xlEdgeBottom).LineStyle = 1

xl.Cells(1, 1) = Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption))
xl.Cells(2, 1) = "B A L A N C E     A L  " & Format(LOC_FECHA_ULT, "dd/mm/yyyy")
xl.Cells(3, 1) = "MONEDA : N U E V O S   S O L E S "

If SALDO_COL1 <> SALDO_COL2 Then
   MsgBox "Balance NO Cuadra por = " & Format(Abs(SALDO_COL1 - SALDO_COL2), "0.00"), 48, Pub_Titulo
End If


  FrmImpC1.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
'  xl.Worksheets(1).Range(wranF).Font.Name = "Draft 17cpi"
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.Application.Visible = True
'  xl.Workbooks(1).Save
 ' xl.Application.Visible = True
  DoEvents
  FrmImpC1.lblProceso.Visible = False
  FrmImpC1.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImpC1.pantalla.Enabled = True
  FrmImpC1.cerrar.Enabled = True
  FrmImpC1.pantalla.Caption = "Por &Pantalla"
  FrmImpC1.lblProceso.Visible = False
Exit Sub


AUXILIAR: ' LIBROS AUXILIARES

For fila = 0 To 5
  If opnivel(fila).Value Then
    WCOL1 = fila + 1
    Exit For
  End If
Next
Dim wscta1, wscta2 As String

pub_cadena = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND COM_CUENTA >=  ? AND COM_CUENTA < ? AND COM_NIVEL = ? ORDER BY COM_CODCIA, COM_CUENTA"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = 0
PS_REP01(2) = 0
PS_REP01(3) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

FrmImpC1.lblProceso.Visible = True
FrmImpC1.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
GoSub WEXCEL
ws_clave = PUB_CLAVE

F1 = 5  'Fila Inicial
C1 = 1
For fila = 0 To listacta.ListCount - 1
  listacta.ListIndex = fila
  If listacta.Selected(fila) Then
    wscta1 = Val(Left(listacta.Text, 6))
    wscta2 = Val(Left(listacta.Text, 6)) + 1
    If WCOL1 > NIVEL_MAX Then
      WCOL1 = NIVEL_MAX
    End If
    GoSub OTRA_CTA
    F1 = F1 + 2
  End If
Next fila

  xl.Cells(1, 1) = Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption))
  xl.Cells(2, 1) = "A U X I L I A R   C U E N T A  "
  xl.Cells(3, 1) = "'" & Format(LOC_FECHA_ULT, "dd/mm/yyyy")
  FrmImpC1.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.Application.Visible = True
  FrmImpC1.lblProceso.Visible = False
  FrmImpC1.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImpC1.pantalla.Enabled = True
  FrmImpC1.cerrar.Enabled = True
  FrmImpC1.pantalla.Caption = "Por &Pantalla"
  FrmImpC1.lblProceso.Visible = False
Exit Sub

OTRA_CTA:
    PS_REP01(0) = LK_CODCIA
    PS_REP01(1) = wscta1
    PS_REP01(2) = wscta2
    PS_REP01(3) = WCOL1
    llave_rep01.Requery
    If llave_rep01.EOF = True Then
       GoTo sigue_cta
    End If
    
    FrmImpC1.ProgBar.Min = 0
    FrmImpC1.ProgBar.Max = llave_rep01.RowCount
    FrmImpC1.ProgBar.Value = 0
    DoEvents
    FrmImpC1.ProgBar.Visible = True
    DoEvents
    xcuenta = 0
    FrmImpC1.lblProceso.Caption = "Procesando . . .  un Momento ."
    DoEvents
    SALDO_TOTAL = 0
    saldos = 0
    Wflag = ""
    xl.Cells(F1, C1 + 1) = Trim(listacta.Text)
    Do Until llave_rep01.EOF
       FrmImpC1.ProgBar.Value = FrmImpC1.ProgBar.Value + 1
       F1 = F1 + 1
       xl.Cells(F1, C1) = Trim(llave_rep01!com_cuenta)
       xl.Cells(F1, C1 + 1) = Trim(llave_rep01!com_DESCRIPCION)
       JALA_SALDO llave_rep01!com_cuenta, periodos.Value
       saldos = saldos + ((PUB_IMPORTE_DEB) * llave_rep01!com_signo_d) + ((PUB_IMPORTE_HAB) * llave_rep01!com_signo_h)
       xl.Cells(F1, C1 + 2) = ((PUB_IMPORTE_DEB) * llave_rep01!com_signo_d) + ((PUB_IMPORTE_HAB) * llave_rep01!com_signo_h)
  
       llave_rep01.MoveNext
    Loop
    F1 = F1 + 1
    'xl.Cells(F1, C1 + 1) = "'            " + Trim(listacta.text) & "   = "
    xl.Cells(F1, C1 + 2) = Format(saldos, "0.00")
    xl.Cells(F1, C1 + 2).Borders.Item(xlEdgeTop).LineStyle = 1
sigue_cta:
Return

WEXCEL:
  Dim wsfile1
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblProceso.Caption = "Abriendo , Archivo BALANCE.xls . . . "
  DoEvents
  WPAS = ws_clave
  'If opnivel(0).Value = True Then
     xl.Workbooks.Open "C:\ADMIN\CONTABILIDAD\BALANCE.xls", 0, True, 4, WPAS, WPAS
  'Else
  '   xl.Workbooks.Open "C:\ADMIN\CONTABILIDAD\AUXILIAR.xls", 0, True, 4, WPAS, WPAS
  'End If


Return
Exit Sub

RESULTADO:
saldo = 0
total = 0
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
         MsgBox "Corregir tab_tipreg = 77...cuentas de Resultado Verificar "
      Else
       JALA_SALDO com_llave!com_cuenta, periodos.Value
       saldo = ((PUB_IMPORTE_DEB) * com_llave!com_signo_d) + ((PUB_IMPORTE_HAB) * com_llave!com_signo_h)
       total = total + (saldo * tab_mayor!TAB_CODART)
      End If
      
   End If
   tab_mayor.MoveNext
Loop

SQ_OPER = 1
PUB_CUENTA = CT_RESULTADO
LEER_COM_LLAVE
If com_llave.EOF Then
     MsgBox "Crear cuenta de Resultados.. ..."
     GoTo fin
End If
JALA_SALDO com_llave!com_cuenta, periodos.Value
saldo = ((PUB_IMPORTE_DEB) * com_llave!com_signo_d) + ((PUB_IMPORTE_HAB) * com_llave!com_signo_h)
saldo = saldo + total

Return
saldo = 0

total = 0
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
           MsgBox "Corregir tab_tipreg=77..."
      End If
      JALA_SALDO com_llave!com_cuenta, periodos.Value
      saldo = ((PUB_IMPORTE_DEB) * com_llave!com_signo_d) + ((PUB_IMPORTE_HAB) * com_llave!com_signo_h)
      'saldo = ((Val(com_llave!COM_DEB_ANO) + Val(com_llave!COM_DEB_MES)) * com_llave!com_SIGNO_D) + ((Val(com_llave!COM_HAB_ANO) + Val(com_llave!COM_HAB_MES)) * com_llave!com_SIGNO_H)
      total = total + (saldo * tab_mayor!TAB_CODART)
   End If
   tab_mayor.MoveNext
Loop

saldo = total
Return

CANCELA:
fin:
  FrmImpC1.pantalla.Enabled = True
  FrmImpC1.pantalla.Caption = "Por &Pantalla"
  FrmImpC1.lblProceso.Visible = False
  FrmImpC1.ProgBar.Visible = False
  pantalla.Enabled = True
  cerrar.Enabled = True
  Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
OJO:
If Err.Number = 70 Then
  MsgBox "Hoja de Calculo : " & wsfile1 & "  esta Abierta debe cerrar para Procesar Nuevamente ", 48, Pub_Titulo
  GoTo CANCELA
End If
Resume Next
Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 xl.Application.Visible = True
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImpC1
 
End Sub


Public Sub LLENA_COPARAN()
Dim XPSCOP_LLAVE As rdoQuery
Dim xcop_llave As rdoResultset
Dim i As Integer
Dim cade
cade = "SELECT * FROM COPARAM WHERE COP_CODCIA = ? "
Set XPSCOP_LLAVE = CN.CreateQuery("", cade)
XPSCOP_LLAVE(0) = 0
Set xcop_llave = XPSCOP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
If LK_EMP_PTO = "A" Then
 XPSCOP_LLAVE.rdoParameters(0) = "00"
Else
 XPSCOP_LLAVE.rdoParameters(0) = LK_CODCIA
End If
xcop_llave.Requery
If Not xcop_llave.EOF Then
  For i = 1 To 6
    If xcop_llave.rdoColumns(i) <> 0 Then
       wCOM_NIVEL(i) = xcop_llave.rdoColumns(i)
       NIVEL_MAX = i
    End If
  Next i
Else
  MsgBox "Definir parametros para el plan contable.", 48, Pub_Titulo
  Exit Sub
End If

End Sub
Public Function SON_FECHAS(wf1 As MaskEdBox, wf2 As MaskEdBox) As Boolean
SON_FECHAS = True
If Right(wf1.Text, 2) = "__" Then
  REP_FECHA1 = Left(wf1.Text, 8)
Else
  REP_FECHA1 = Trim(wf1.Text)
End If
If Not IsDate(REP_FECHA1) Then
    MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
    Azul2 wf1, wf1
    GoTo fin
End If
If Right(wf2.Text, 2) = "__" Then
  REP_FECHA2 = Left(wf2.Text, 8)
Else
  REP_FECHA2 = Trim(wf2.Text)
End If
If Not IsDate(REP_FECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 Azul2 wf2, wf2
 GoTo fin
End If
If CDate(REP_FECHA1) > CDate(REP_FECHA2) Then
 MsgBox "Fechas Invalidadas ..", 48, Pub_Titulo
 Azul2 wf1, wf1
 GoTo fin
End If

Exit Function
fin:
SON_FECHAS = False

End Function


Public Sub LLENA_CTA(wnivel As Integer)
 PSCTA1(0) = LK_CODCIA
 PSCTA1(1) = wnivel
 loc_cta1.Requery
 listacta.Visible = False
 listacta.Clear
 Do Until loc_cta1.EOF
   listacta.AddItem Trim(loc_cta1!com_cuenta) + "   " + Trim(loc_cta1!com_DESCRIPCION)
   loc_cta1.MoveNext
 Loop
 listacta.Visible = True
 
End Sub

Public Sub BAL_COMPRO()
'On Error GoTo FINTODO
Dim SALDO_898 As Currency
Dim SALDO_898_DEB As Currency
Dim SALDO_898_HAB As Currency
Dim CUENTA_898 As String
Dim DESCRIPCION_898 As String

Dim WCUENTA As String
Dim wtipcta
Dim CT_RESULTADO As String
Dim LETRAS(24) As String * 1
Dim WCAMBIA
Dim ws_clave As String
Dim WSFECHA As Date
Dim wcodven As Integer
Dim wvalor
Dim Wche As Integer
Dim wkSELECT As String
Dim wsfile As String
Dim F2 As Integer
Dim saldos As Currency
Dim SALDO_TOTAL As Currency
Dim Wflag As String * 1
Dim WCOL1 As Integer
Dim WCOL2 As Integer
Dim SALDO_COL1 As Currency
Dim SALDO_COL2 As Currency
Dim COL_SALDO As Currency
Dim saldos_D As Currency
Dim saldos_H  As Currency
Dim DEUDOR As Currency
Dim ACREEDOR  As Currency
Dim ACTIVO As Currency
Dim PASIVO  As Currency
Dim PGF_PERDIDAS As Currency
Dim PGF_GANANCIAS As Currency
Dim PGN_PERDIDAS As Currency
Dim PGN_GANANCIAS As Currency
If Val(txtnivel.Text) <= 0 Then
  MsgBox "Verificar su nivel que desea ver", 48, Pub_Titulo
  Azul txtnivel, txtnivel
  Exit Sub
End If
If Val(txtnivel.Text) > cop_llave!cop_nivel_max Then
  MsgBox "Nivel no procede...", 48, Pub_Titulo
  Azul txtnivel, txtnivel
  Exit Sub
End If

wsfile = ""
pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImpC1.lblProceso.Caption = "Activando Reporte... un Momento ."
DoEvents
pub_cadena = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND COM_NIVEL = " & Trim(txtnivel.Text) & "  ORDER BY COM_CODCIA, COM_CUENTA"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
DoEvents

'*** VERFICA SI HAY DATOS , O ESTAN CORRECTOS
PS_REP01(0) = LK_CODCIA
llave_rep01.Requery
If llave_rep01.EOF = True Then
   MsgBox "!!! NO EXISTEN Datos ...", 48, Pub_Titulo
   GoTo CANCELA
End If
ws_clave = PUB_CLAVE
FrmImpC1.ProgBar.Min = 0
FrmImpC1.ProgBar.Max = llave_rep01.RowCount
FrmImpC1.ProgBar.Value = 0
FrmImpC1.lblProceso.Visible = True
FrmImpC1.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL

GoSub LETRAS
'xlLineStyleNone
xl.Range("A4:L5").Borders.LineStyle = 3
'xl.Range("A4:L5").Borders.Item(xlEdgeTop).LineStyle = 3
'xl.Range("A4:L5").Borders.Item(xlEdgeTop).LineStyle = 3

FrmImpC1.ProgBar.Visible = True
DoEvents
xcuenta = 0
F2 = 5
F1 = 5  'Fila Inicial
C1 = 1
FrmImpC1.lblProceso.Caption = "Procesando . . .  un Momento ."
DoEvents
fila = 0
WCAMBIA = llave_rep01!com_tipo_cta
SALDO_TOTAL = 0
saldos = 0
Wflag = ""
CT_RESULTADO = "89"
saldos_D = 0
saldos_H = 0
ACREEDOR = 0
DEUDOR = 0
ACTIVO = 0
PASIVO = 0
PGF_PERDIDAS = 0
PGF_GANANCIAS = 0
PGN_PERDIDAS = 0
PGN_GANANCIAS = 0

Do Until llave_rep01.EOF
   WCUENTA = Trim(llave_rep01!com_cuenta)
'   If Left(WCUENTA, 2) = "79" Then Stop
   wtipcta = llave_rep01!com_tipo_cta
   FrmImpC1.ProgBar.Value = FrmImpC1.ProgBar.Value + 1
   JALA_SALDO llave_rep01!com_cuenta, periodos.Value
   If Abs(PUB_IMPORTE_DEB) = 0 And Abs(PUB_IMPORTE_HAB) = 0 Then GoTo otrito
   saldos_D = saldos_D + Abs(PUB_IMPORTE_DEB) '+ Abs(llave_rep01!COM_DEB_ANO)
   saldos_H = saldos_H + Abs(PUB_IMPORTE_HAB) '+ Abs(llave_rep01!COM_HAB_ANO)
   F1 = F1 + 1
   xl.Cells(F1, C1) = Trim(llave_rep01!com_cuenta)
   xl.Cells(F1, C1 + 1) = Left(llave_rep01!com_DESCRIPCION, 35)
   If Val(PUB_IMPORTE_DEB) <> 0 Then xl.Cells(F1, C1 + 2) = Abs(PUB_IMPORTE_DEB)
   If Val(PUB_IMPORTE_HAB) <> 0 Then xl.Cells(F1, C1 + 3) = Abs(PUB_IMPORTE_HAB)
   COL_SALDO = Abs(Val(PUB_IMPORTE_DEB)) - Abs(Val(PUB_IMPORTE_HAB))
SALTA_CALCULA:
   If COL_SALDO > 0 And COL_SALDO <> 0 Then
      DEUDOR = DEUDOR + Val(Abs(COL_SALDO))
      xl.Cells(F1, C1 + 4) = Val(Abs(COL_SALDO))
      If wtipcta = 4 Or wtipcta = 3 Or wtipcta = 5 Or wtipcta = 2 Or wtipcta = 1 Or wtipcta = 3 Then
        ACTIVO = ACTIVO + Val(Abs(COL_SALDO))
        xl.Cells(F1, C1 + 6) = Val(Abs(COL_SALDO))
      End If
      If wtipcta = 6 Or wtipcta = 7 Or Left(WCUENTA, 2) = "89" Then
         PGF_PERDIDAS = PGF_PERDIDAS + Val(Abs(COL_SALDO))
          xl.Cells(F1, C1 + 8) = Val(Abs(COL_SALDO))
      End If
      If wtipcta = 6 Or wtipcta = 10 Or Left(WCUENTA, 2) = "69" Or Left(WCUENTA, 2) = "66" Or Left(WCUENTA, 2) = "89" Then
          PGN_PERDIDAS = PGN_PERDIDAS + Val(Abs(COL_SALDO))
          xl.Cells(F1, C1 + 10) = Val(Abs(COL_SALDO))
      End If
   ElseIf COL_SALDO < 0 And COL_SALDO <> 0 Then
      ACREEDOR = ACREEDOR + Val(Abs(COL_SALDO))
      xl.Cells(F1, C1 + 5) = Val(Abs(COL_SALDO))
      If wtipcta = 4 Or wtipcta = 3 Or wtipcta = 5 Or wtipcta = 1 Or wtipcta = 2 Or wtipcta = 4 Then
        PASIVO = PASIVO + Val(Abs(COL_SALDO))
        xl.Cells(F1, C1 + 7) = Val(Abs(COL_SALDO))
      End If
      If wtipcta = 6 Or wtipcta = 7 Or Left(WCUENTA, 2) = "89" Then
         PGF_GANANCIAS = PGF_GANANCIAS + Val(Abs(COL_SALDO))
          xl.Cells(F1, C1 + 9) = Val(Abs(COL_SALDO))
      End If
      If wtipcta = 6 Or wtipcta = 10 Or Left(WCUENTA, 2) = "69" Or Left(WCUENTA, 2) = "66" Or Left(WCUENTA, 2) = "89" Then
          PGN_GANANCIAS = PGN_GANANCIAS + Val(Abs(COL_SALDO))
          xl.Cells(F1, C1 + 11) = Val(Abs(COL_SALDO))
      End If
   End If
otrito:
   llave_rep01.MoveNext
Loop



Dim CARAC As String
Dim saldo As Currency
Dim total As Currency
Dim wfi As Integer
saldo = 0
total = 0
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
         MsgBox "Corregir tab_tipreg = 77...cuentas de Resultado Verificar "
      Else
      JALA_SALDO com_llave!com_cuenta, periodos.Value
      saldo = ((PUB_IMPORTE_DEB) * com_llave!com_signo_d) + ((PUB_IMPORTE_HAB) * com_llave!com_signo_h)
      total = total + (saldo * tab_mayor!TAB_CODART)
      End If
   End If
   tab_mayor.MoveNext
Loop


' REVISAR SI ES NECESARIO
'SQ_OPER = 1
'PUB_CUENTA = CT_RESULTADO
'LEER_COM_LLAVE
'If com_llave.EOF Then
'     MsgBox "Crear cuenta de Resultados.. ..."
'     GoTo fin
'End If
'JALA_SALDO com_llave!com_cuenta, periodos.Value
'saldo = ((PUB_IMPORTE_DEB) * com_llave!com_signo_d) + ((PUB_IMPORTE_HAB) * com_llave!com_signo_h)
saldo = saldo + total
F1 = F1 + 1
wfi = F1
xl.Cells(F1, C1 + 2) = saldos_D
xl.Cells(F1, C1 + 3) = saldos_H
xl.Cells(F1, C1 + 4) = DEUDOR
xl.Cells(F1, C1 + 5) = ACREEDOR
xl.Cells(F1, C1 + 6) = ACTIVO
xl.Cells(F1, C1 + 7) = PASIVO
xl.Cells(F1, C1 + 8) = PGF_PERDIDAS
xl.Cells(F1, C1 + 9) = PGF_GANANCIAS
xl.Cells(F1, C1 + 10) = PGN_PERDIDAS
xl.Cells(F1, C1 + 11) = PGN_GANANCIAS
wranF = "A" & F1 & ":L" & F1
xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
F1 = F1 + 1
xl.Cells(F1, C1 + 1) = "" ' com_llave!com_DESCRIPCION

total = saldo


'If saldo > 0 Then
If ACTIVO > PASIVO Then
   'para activo y pasivo de inventario
   xl.Cells(F1, C1 + 7) = Format(Abs(total), "0.00")
   saldo = Val(xl.Cells(F1 - 1, C1 + 7)) + Val(xl.Cells(F1, C1 + 7))
   xl.Cells(F1 + 1, C1 + 7) = Format(saldo, "0.00")
   xl.Cells(F1 + 1, C1 + 7).Borders.Item(xlEdgeTop).LineStyle = 1
   saldo = Val(xl.Cells(F1 - 1, C1 + 6)) + Val(xl.Cells(F1, C1 + 6))
   xl.Cells(F1 + 1, C1 + 6) = Format(saldo, "0.00")
   xl.Cells(F1 + 1, C1 + 6).Borders.Item(xlEdgeTop).LineStyle = 1
   
Else

   'para activo y pasivo de inventario
   
   xl.Cells(F1, C1 + 6) = Format(Abs(total), "0.00")
   saldo = Val(xl.Cells(F1 - 1, C1 + 6)) + Val(xl.Cells(F1, C1 + 6))
   xl.Cells(F1 + 1, C1 + 6) = Format(saldo, "0.00")
   xl.Cells(F1 + 1, C1 + 6).Borders.Item(xlEdgeTop).LineStyle = 1
   saldo = Val(xl.Cells(F1 - 1, C1 + 7)) + Val(xl.Cells(F1, C1 + 7))
   xl.Cells(F1 + 1, C1 + 7) = Format(saldo, "0.00")
   xl.Cells(F1 + 1, C1 + 7).Borders.Item(xlEdgeTop).LineStyle = 1
End If
If PGF_PERDIDAS > PGF_GANANCIAS Then
   xl.Cells(F1, C1 + 9) = Format(Abs(total), "0.00")
   saldo = Val(xl.Cells(F1 - 1, C1 + 9)) + Val(xl.Cells(F1, C1 + 9))
   xl.Cells(F1 + 1, C1 + 9) = Format(saldo, "0.00")
   xl.Cells(F1 + 1, C1 + 9).Borders.Item(xlEdgeTop).LineStyle = 1
   saldo = Val(xl.Cells(F1 - 1, C1 + 8)) + Val(xl.Cells(F1, C1 + 8))
   xl.Cells(F1 + 1, C1 + 8) = Format(saldo, "0.00")
   xl.Cells(F1 + 1, C1 + 8).Borders.Item(xlEdgeTop).LineStyle = 1
   
Else
   xl.Cells(F1, C1 + 8) = Format(Abs(total), "0.00")
   saldo = Val(xl.Cells(F1 - 1, C1 + 8)) + Val(xl.Cells(F1, C1 + 8))
   xl.Cells(F1 + 1, C1 + 8) = Format(saldo, "0.00")
   xl.Cells(F1 + 1, C1 + 8).Borders.Item(xlEdgeTop).LineStyle = 1
   saldo = Val(xl.Cells(F1 - 1, C1 + 9)) + Val(xl.Cells(F1, C1 + 9))
   xl.Cells(F1 + 1, C1 + 9) = Format(saldo, "0.00")
   xl.Cells(F1 + 1, C1 + 9).Borders.Item(xlEdgeTop).LineStyle = 1
End If
If PGN_PERDIDAS > PGN_GANANCIAS Then
   xl.Cells(F1, C1 + 11) = Format(Abs(total), "0.00")
   saldo = Val(xl.Cells(F1 - 1, C1 + 11)) + Val(xl.Cells(F1, C1 + 11))
   xl.Cells(F1 + 1, C1 + 11) = Format(saldo, "0.00")
   xl.Cells(F1 + 1, C1 + 11).Borders.Item(xlEdgeTop).LineStyle = 1
   saldo = Val(xl.Cells(F1 - 1, C1 + 10)) + Val(xl.Cells(F1, C1 + 10))
   xl.Cells(F1 + 1, C1 + 10) = Format(saldo, "0.00")
   xl.Cells(F1 + 1, C1 + 10).Borders.Item(xlEdgeTop).LineStyle = 1
   
Else
   xl.Cells(F1, C1 + 10) = Format(Abs(total), "0.00")
   saldo = Val(xl.Cells(F1 - 1, C1 + 10)) + Val(xl.Cells(F1, C1 + 10))
   xl.Cells(F1 + 1, C1 + 10) = Format(saldo, "0.00")
   xl.Cells(F1 + 1, C1 + 10).Borders.Item(xlEdgeTop).LineStyle = 1
   saldo = Val(xl.Cells(F1 - 1, C1 + 11)) + Val(xl.Cells(F1, C1 + 11))
   xl.Cells(F1 + 1, C1 + 11) = Format(saldo, "0.00")
   xl.Cells(F1 + 1, C1 + 11).Borders.Item(xlEdgeTop).LineStyle = 1
End If


xl.Cells(1, 1) = Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption))
xl.Cells(2, 1) = "BALANCE DE COMPROBACION AL  " & Format(LOC_FECHA_ULT, "dd/mm/yyyy")
  FrmImpC1.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.Application.Visible = True
  FrmImpC1.lblProceso.Visible = False
  FrmImpC1.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImpC1.pantalla.Enabled = True
  FrmImpC1.cerrar.Enabled = True
  FrmImpC1.pantalla.Caption = "Por &Pantalla"
  FrmImpC1.lblProceso.Visible = False
Exit Sub

LETRAS:
LETRAS(1) = "A"
LETRAS(2) = "B"
LETRAS(3) = "C"
LETRAS(4) = "D"
LETRAS(5) = "E"
LETRAS(6) = "F"
LETRAS(7) = "G"
LETRAS(8) = "H"
LETRAS(9) = "I"
LETRAS(10) = "J"
LETRAS(11) = "K"
LETRAS(12) = "L"
LETRAS(13) = "M"
LETRAS(14) = "N"
LETRAS(15) = "O"
LETRAS(16) = "P"
LETRAS(17) = "Q"
LETRAS(18) = "R"
LETRAS(19) = "S"
LETRAS(20) = "T"
LETRAS(21) = "U"
LETRAS(22) = "V"
LETRAS(23) = "W"
LETRAS(24) = "X"
Return

WEXCEL:
  Dim wsfile1
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblProceso.Caption = "Abriendo , Archivo BAL_COMPRO.xls . . . "
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open "C:\ADMIN\CONTABILIDAD\BAL_COMPRO.xls", 0, True, 4, WPAS, WPAS

Return

Exit Sub
CANCELA:
  FrmImpC1.pantalla.Enabled = True
  FrmImpC1.pantalla.Caption = "Por &Pantalla"
  FrmImpC1.lblProceso.Visible = False
  FrmImpC1.ProgBar.Visible = False
  pantalla.Enabled = True
  cerrar.Enabled = True
  Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub


OJO:
If Err.Number = 70 Then
  MsgBox "Hoja de Calculo : " & wsfile1 & "  esta Abierta debe cerrar para Procesar Nuevamente ", 48, Pub_Titulo
  GoTo CANCELA
End If
Resume Next
Exit Sub
FINTODO:
fin:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 xl.Application.Visible = True
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImpC1
Exit Sub

CLASE_898:
MsgBox "Verificar procedimiento CLASE_898"
SALDO_898 = 0
SQ_OPER = 1
PUB_CUENTA = "988"
If periodos.Value = 1 Then
 PSCOH_LLAVE(0) = Left(Trim(FECHA1.Text), 10)
 PSCOH_LLAVE(1) = Right(Trim(FECHA1.Text), 10)
 PSCOH_LLAVE(2) = PUB_CODCIA
 PSCOH_LLAVE(3) = PUB_CUENTA
 coh_llave.Requery
 If coh_llave.EOF Then
     MsgBox "Crear cuenta de Resultados.. ..."
     GoTo fin
 End If
 SALDO_898 = ((Val(coh_llave!COH_DEB_ANO) + Val(coh_llave!COH_DEB_MES)) * coh_llave!COH_SIGNO_D) + ((Val(coh_llave!COH_HAB_ANO) + Val(coh_llave!COH_HAB_MES)) * coh_llave!COH_SIGNO_H)
 SALDO_898_DEB = Val(coh_llave!COH_DEB_MES)
 SALDO_898_HAB = Val(coh_llave!COH_HAB_MES)
Else
  LEER_COM_LLAVE
  If com_llave.EOF Then
     MsgBox "Crear cuenta de Resultados.. ..."
     GoTo fin
  End If
  SALDO_898 = ((Val(com_llave!COM_DEB_ANO) + Val(com_llave!COM_DEB_MES)) * com_llave!com_signo_d) + ((Val(com_llave!COM_HAB_ANO) + Val(com_llave!COM_HAB_MES)) * com_llave!com_signo_h)
  SALDO_898_DEB = Val(com_llave!COM_DEB_MES)
  SALDO_898_HAB = Val(com_llave!COM_HAB_MES)
End If

Return
End Sub


Public Sub ANALISIS()
'On Error GoTo FINTODO
Dim WCREDITO   As Currency
Dim wefectivo As Currency
Dim wRuta As String
Dim WMONTO As Currency
Dim wcodvend As Integer
Dim var_ACUTOT As Currency
Dim var_ACUATE As Currency
Dim var_ACUPED As Currency
Dim wnumfac As Currency
Dim ws_clave As String
Dim Wflag As String * 1
Dim wflag2 As String * 1
Dim wsFECHA1, wsFECHA2
Dim wranF, wran1, wran2, WPAS
Dim WS_VENTAS As Currency
Dim WS_COBRANZAS As Currency
Dim WS_CREDITOS As Currency
Dim WCHEQUE As Currency
Dim WS_GASTOS As Currency
Dim WS  As String
Dim wconcepto As String
Dim wgrupo As String * 1
Dim wnumoper2 As Integer


If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If

pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImpC1.lblProceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = PUB_CLAVE
FrmImpC1.ProgBar.Visible = True
DoEvents
FrmImpC1.ProgBar.Value = 0
FrmImpC1.lblProceso.Visible = True
FrmImpC1.lblProceso.Caption = "Activando Reporte. . . "
GoSub WEXCEL

' * *  CUADRO INGRESOS X VENTAS  * *
F1 = 7  'Fila Inicial
pub_cadena = "SELECT * FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_CODTRA = ? AND ALL_FECHA_DIA = ?  AND ALL_FLAG_EXT <> 'E'   ORDER BY ALL_CODCIA, ALL_CODVEN, ALL_NUMOPER"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = 0
PS_REP01(2) = LK_FECHA_DIA
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = 2401
PS_REP01(2) = wsFECHA1
llave_rep01.Requery
If llave_rep01.EOF Then
  GoTo COBRANZAS:
End If
F1 = 7  'Fila Inicial
Wflag = ""
FrmImpC1.ProgBar.Value = 0
FrmImpC1.ProgBar.Min = 0
FrmImpC1.ProgBar.Max = llave_rep01.RowCount
FrmImpC1.lblProceso.Caption = "Procesando . . . "
wcodvend = llave_rep01!ALL_CODVEN
WCREDITO = 0
wefectivo = 0
var_ACUTOT = 0
wgrupo = ""
'wnumoper2 = Nulo_Valor0(llave_rep01!ALL_numoper2)
wnumoper2 = Nulo_Valor0(llave_rep01!ALL_SIGNO_CAR)
Do Until llave_rep01.EOF
   FrmImpC1.ProgBar.Value = FrmImpC1.ProgBar.Value + 1
   If wcodvend <> llave_rep01!ALL_CODVEN Then
     F1 = F1 + 1
     xl.Cells(F1, 1) = wcodvend
     If wefectivo <> 0 Then xl.Cells(F1, 2) = Format(wefectivo, "0.00")
     If WCREDITO <> 0 Then xl.Cells(F1, 4) = Format(WCREDITO, "0.00")
     xl.Cells(F1, 6) = Format(WCREDITO + wefectivo, "0.00")
     var_ACUTOT = var_ACUTOT + (WCREDITO + wefectivo)
     wcodvend = llave_rep01!ALL_CODVEN
     wefectivo = 0
     WCREDITO = 0
     If llave_rep01!ALL_SIGNO_CAR = 0 Then
       wefectivo = wefectivo + llave_rep01!all_Importe_AMORT
     ElseIf llave_rep01!ALL_SIGNO_CAR <> 1 Then
       WCREDITO = WCREDITO + llave_rep01!all_Importe_AMORT
     End If
   Else
    If wnumoper2 <> Nulo_Valor0(llave_rep01!ALL_SIGNO_CAR) Then
       If llave_rep01!ALL_SIGNO_CAR = 0 Then
        wefectivo = wefectivo + llave_rep01!all_Importe_AMORT
       ElseIf llave_rep01!ALL_SIGNO_CAR <> 0 Then
        WCREDITO = WCREDITO + llave_rep01!all_Importe_AMORT
       End If
    Else
      llave_rep01.MoveNext
      If Not llave_rep01.EOF Then
        If Nulo_Valors(llave_rep01!all_TIPDOC) = "CH" Then
            WCHEQUE = WCHEQUE + llave_rep01!all_Importe_AMORT
            If wnumoper2 = 0 Then llave_rep01.MoveNext
            GoTo otrito
         Else
            llave_rep01.MovePrevious
        End If
       Else
            llave_rep01.MovePrevious
       End If
     If llave_rep01!ALL_SIGNO_CAR = 0 Then
        wefectivo = wefectivo + llave_rep01!all_Importe_AMORT
     ElseIf llave_rep01!ALL_SIGNO_CAR <> 0 Then
        WCREDITO = WCREDITO + llave_rep01!all_Importe_AMORT
     End If
otrito:
     wnumoper2 = Nulo_Valor0(llave_rep01!ALL_SIGNO_CAR)
    End If
   End If
   llave_rep01.MoveNext
Loop
F1 = F1 + 1
xl.Cells(F1, 1) = wcodvend
If wefectivo <> 0 Then xl.Cells(F1, 2) = Format(wefectivo, "0.00")
If WCHEQUE <> 0 Then xl.Cells(F1, 3) = Format(WCHEQUE, "0.00")
If WCREDITO <> 0 Then xl.Cells(F1, 4) = Format(WCREDITO, "0.00")

WS_CREDITOS = WCREDITO
xl.Cells(F1, 6) = Format(WCREDITO + wefectivo + WCHEQUE, "0.00")

var_ACUTOT = var_ACUTOT + (WCREDITO + wefectivo + WCHEQUE)
F1 = F1 + 1
xl.Cells(F1, 6) = Format(var_ACUTOT, "0.00")
xl.Cells(F1, 6).Borders.Item(xlEdgeTop).LineStyle = 1
WS_VENTAS = var_ACUTOT


COBRANZAS:
' * *  CUADRO INGRESOS X COBRANZAS  * *
F1 = F1 + 4 'Fila SIGUIENTE ES ?
xl.Cells(F1, 1) = "INGRESOS POR COBRANZAS"
F1 = F1 + 1
xl.Cells(F1, 1) = "Nº PLA"
xl.Cells(F1, 2) = "CHE/DEV"
xl.Cells(F1, 3) = "CREDIT"
xl.Cells(F1, 4) = "CTAS.CTES"
xl.Cells(F1, 5) = "EFECTIVO"
xl.Cells(F1, 6) = "CHEQUES"
xl.Cells(F1, 7) = "TOTAL"
'xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
wranF = "A" & F1 & ":G" & F1
xl.Range(wranF).Borders.LineStyle = 1

pub_cadena = "SELECT * FROM CARACU WHERE CAA_CODCIA = ? AND CAA_CP = ? AND CAA_FECHA  = ?   and CAA_ESTADO <> 'E'  AND CAA_SIGNO_CAJA = 1  ORDER BY CAA_CODCIA, CAA_NUMPLAN"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = 0
PS_REP01(2) = LK_FECHA_DIA
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = "C"
PS_REP01(2) = wsFECHA1
llave_rep01.Requery
If llave_rep01.EOF Then
  GoTo CHEQUESR
End If
Wflag = ""
FrmImpC1.ProgBar.Value = 0
FrmImpC1.ProgBar.Min = 0
FrmImpC1.ProgBar.Max = llave_rep01.RowCount
FrmImpC1.lblProceso.Caption = "Procesando . . . "
wcodvend = llave_rep01!CAA_NUMPLAN
WCREDITO = 0
wefectivo = 0
var_ACUTOT = 0
Do Until llave_rep01.EOF
   FrmImpC1.ProgBar.Value = FrmImpC1.ProgBar.Value + 1
   If wcodvend <> llave_rep01!CAA_NUMPLAN Then
     F1 = F1 + 1
     xl.Cells(F1, 1) = wcodvend
     If wefectivo <> 0 Then xl.Cells(F1, 5) = Format(wefectivo, "0.00")
     If WCREDITO <> 0 Then xl.Cells(F1, 6) = Format(WCREDITO, "0.00")
     xl.Cells(F1, 7) = Format(WCREDITO + wefectivo, "0.00")
     var_ACUTOT = var_ACUTOT + (WCREDITO + wefectivo)
     wcodvend = llave_rep01!CAA_NUMPLAN
     wefectivo = 0
     WCREDITO = 0
     If llave_rep01!CAA_TIPDOC = "FA" Then
       wefectivo = wefectivo + llave_rep01!CAA_IMPORTE * -1
     ElseIf llave_rep01!CAA_TIPDOC = "CH" Then
       WCREDITO = WCREDITO + llave_rep01!CAA_IMPORTE * -1
     End If
   Else
     If llave_rep01!CAA_TIPDOC = "FA" Then
       wefectivo = wefectivo + llave_rep01!CAA_IMPORTE * -1
     ElseIf llave_rep01!CAA_TIPDOC = "CH" Then
       WCREDITO = WCREDITO + Abs(llave_rep01!CAA_IMPORTE)
     End If
   End If
   llave_rep01.MoveNext
Loop
F1 = F1 + 1
xl.Cells(F1, 1) = wcodvend
If wefectivo <> 0 Then xl.Cells(F1, 5) = Format(wefectivo, "0.00")
If WCREDITO <> 0 Then xl.Cells(F1, 6) = Format(WCREDITO, "0.00")
xl.Cells(F1, 7) = Format(WCREDITO + wefectivo, "0.00")
var_ACUTOT = var_ACUTOT + (WCREDITO + wefectivo)
F1 = F1 + 1
xl.Cells(F1, 7) = Format(var_ACUTOT, "0.00")
xl.Cells(F1, 7).Borders.Item(xlEdgeTop).LineStyle = 1
WS_COBRANZAS = var_ACUTOT

CHEQUESR:
' * *  CUADRO INGRESOS X COBRANZAS  * *

F1 = F1 + 4 'Fila SIGUIENTE ES ?
xl.Cells(F1, 1) = "CHEQUES RECEPCIONADOS"
F1 = F1 + 1
xl.Cells(F1, 1) = "CHEQ.NRO."
xl.Cells(F1, 3) = "G I R A D O R"
xl.Cells(F1, 5) = "I M P O R T E"
'xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
wranF = "A" & F1 & ":F" & F1
xl.Range(wranF).Borders.LineStyle = 1

'pub_cadena = "SELECT * FROM CARTERA WHERE CAR_CODCIA = ? AND CAR_CP = ? AND CAR_FECHA_INGR = ?  AND  CAR_TIPDOC = 'CH'  ORDER BY CAR_CODCIA"
pub_cadena = "SELECT * FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_DIA = ?  AND ALL_TIPDOC = 'CH' AND ALL_FLAG_EXT <> 'E'   ORDER BY ALL_CODCIA,ALL_NUMOPER"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = LK_FECHA_DIA
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
PS_REP01(0) = LK_CODCIA
'PS_REP01(1) = "C"
PS_REP01(1) = wsFECHA1
llave_rep01.Requery
If llave_rep01.EOF Then
  GoTo BANCOS
End If
Wflag = ""
FrmImpC1.ProgBar.Value = 0
FrmImpC1.ProgBar.Min = 0
FrmImpC1.ProgBar.Max = llave_rep01.RowCount
FrmImpC1.lblProceso.Caption = "Procesando . . . "
'wcodvend = llave_rep01!CAA_NUMPLAN
WCREDITO = 0
wefectivo = 0
var_ACUTOT = 0
pu_cp = "C"
pu_codcia = LK_CODCIA
Do Until llave_rep01.EOF
   FrmImpC1.ProgBar.Value = FrmImpC1.ProgBar.Value + 1
   F1 = F1 + 1
   xl.Cells(F1, 1) = llave_rep01!ALL_CHENUM
   pu_codclie = llave_rep01!all_codclie
   SQ_OPER = 1
   LEER_CLI_LLAVE
   If Not cli_llave.EOF Then xl.Cells(F1, 3) = Trim(cli_llave!cli_nombre)
   xl.Cells(F1, 6) = Format(llave_rep01!all_Importe_AMORT, "0.00")
   wefectivo = wefectivo + (llave_rep01!all_Importe_AMORT)
   llave_rep01.MoveNext
Loop
F1 = F1 + 1
xl.Cells(F1, 6) = Format(wefectivo, "0.00")
xl.Cells(F1, 6).Borders.Item(xlEdgeTop).LineStyle = 1


BANCOS:
' * *  CUADRO DEPOSITOS A BANCOS  * *
F1 = F1 + 4 'Fila SIGUIENTE ES ?
xl.Cells(F1, 1) = "ENTREGAS BANCARIAS"
F1 = F1 + 1
xl.Cells(F1, 1) = "BANCO"
xl.Cells(F1, 4) = "DESCRIPCION"
xl.Cells(F1, 7) = "MONTO"
'xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
wranF = "A" & F1 & ":G" & F1
xl.Range(wranF).Borders.LineStyle = 1

pub_cadena = "SELECT * FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_CODTRA = ? AND ALL_FECHA_DIA = ?  AND ALL_FLAG_EXT <> 'E' ORDER BY ALL_CODCIA, ALL_CODBAN"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = 0
PS_REP01(2) = LK_FECHA_DIA
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = 5310
PS_REP01(2) = wsFECHA1
llave_rep01.Requery
If llave_rep01.EOF Then
  GoTo RESUMEN
End If
Wflag = ""
FrmImpC1.ProgBar.Value = 0
FrmImpC1.ProgBar.Min = 0
FrmImpC1.ProgBar.Max = llave_rep01.RowCount
FrmImpC1.lblProceso.Caption = "Procesando . . . "
wcodvend = llave_rep01!ALL_CODBAN
WCREDITO = 0
wefectivo = 0
var_ACUTOT = 0
pu_codcia = LK_CODCIA
Do Until llave_rep01.EOF
   FrmImpC1.ProgBar.Value = FrmImpC1.ProgBar.Value + 1
   If wcodvend <> llave_rep01!ALL_CODBAN Then
$ $  F1 = F1 + 1
     PUB_CODBAN = wcodvend
     LEER_CCM_LLAVE
     xl.Cells(F1, 1) = ccm_llave!CCM_DESCRIPCION
     xl.Cells(F1, 4) = llave_rep01!ALL_CONCEPTO
     If WCREDITO <> 0 Then xl.Cells(F1, 7) = Format(WCREDITO, "0.00")
     var_ACUTOT = var_ACUTOT + WCREDITO
     wcodvend = llave_rep01!ALL_CODBAN
     WCREDITO = 0
     WCREDITO = WCREDITO + llave_rep01!all_Importe_AMORT
     wconcepto = llave_rep01!ALL_CONCEPTO
   Else
    WCREDITO = WCREDITO + llave_rep01!ALL_IMPORTE
    wconcepto = llave_rep01!ALL_CONCEPTO
   End If
   llave_rep01.MoveNext
Loop
F1 = F1 + 1
PUB_CODBAN = wcodvend
LEER_CCM_LLAVE
xl.Cells(F1, 1) = Trim(ccm_llave!CCM_nombre)
xl.Cells(F1, 4) = wconcepto
If WCREDITO <> 0 Then xl.Cells(F1, 7) = Format(WCREDITO, "0.00")
var_ACUTOT = var_ACUTOT + WCREDITO
WCREDITO = 0
F1 = F1 + 1
xl.Cells(F1, 7) = Format(var_ACUTOT, "0.00")
xl.Cells(F1, 7).Borders.Item(xlEdgeTop).LineStyle = 1


RESUMEN:

F1 = F1 + 4 'Fila SIGUIENTE ES ?
xl.Cells(F1, 1) = "RESUMEN DE INGRESOS Y EGRESOS"
F1 = F1 + 1
xl.Cells(F1, 1) = "I N G R E S O S"
wranF = "A" & F1 & ":D" & F1
xl.Range(wranF).Borders.LineStyle = 1
F1 = F1 + 1
xl.Cells(F1, 1) = "- VENTAS PROPIAS "
xl.Cells(F1, 4) = Format(WS_VENTAS, "0.00")
F1 = F1 + 1
xl.Cells(F1, 1) = "- VENTAS COBRANZAS "
xl.Cells(F1, 4) = Format(WS_COBRANZAS, "0.00")
F1 = F1 + 1
xl.Cells(F1, 2) = "- TOTAL DE INGRESOS "
xl.Cells(F1, 4) = Format(WS_COBRANZAS + WS_VENTAS, "0.00")
xl.Cells(F1, 4).Borders.Item(xlEdgeTop).LineStyle = 1
F1 = F1 + 2
xl.Cells(F1, 1) = "E G R E S O S"
wranF = "A" & F1 & ":D" & F1
xl.Range(wranF).Borders.LineStyle = 1

F1 = F1 + 1
xl.Cells(F1, 1) = "- CREDITOS "
xl.Cells(F1, 4) = Format(WS_CREDITOS, "0.00")
F1 = F1 + 1
xl.Cells(F1, 1) = "- GASTOS VARIOS "
WS = wsFECHA1
WS_GASTOS = EGRE_CAJA(WS)
xl.Cells(F1, 4) = Format(WS_GASTOS, "0.00")
F1 = F1 + 1
xl.Cells(F1, 2) = "- TOTAL DE EGRESOS "
xl.Cells(F1, 4) = Format(WS_CREDITOS + WS_GASTOS, "0.00")
xl.Cells(F1, 4).Borders.Item(xlEdgeTop).LineStyle = 1

F1 = F1 + 5
xl.Cells(F1, 4) = "    C A J E R O    "
xl.Cells(F1, 4).Borders.Item(xlEdgeTop).LineStyle = 1
xl.Cells(F1, 5).Borders.Item(xlEdgeTop).LineStyle = 1
F1 = F1 + 3
xl.Cells(F1, 4) = "  C O N T A D O R "
xl.Cells(F1, 4).Borders.Item(xlEdgeTop).LineStyle = 1
xl.Cells(F1, 5).Borders.Item(xlEdgeTop).LineStyle = 1



  xl.Cells(1, 1) = Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption))
  FrmImpC1.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(2, 1) = "ANALISIS DE LA RECEPCION DE VENTAS Y COBRANZAS"
  xl.Cells(3, 6) = "FECHA :  " & LK_FECHA_DIA
  xl.Cells(4, 6) = "Nº    :  "
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.Application.Visible = True
  FrmImpC1.lblProceso.Visible = False
  FrmImpC1.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImpC1.pantalla.Enabled = True
  FrmImpC1.cerrar.Enabled = True
  FrmImpC1.pantalla.Caption = "Por &Pantalla"
  FrmImpC1.lblProceso.Visible = False
Exit Sub
WEXCEL:
  Dim wsfile1
  wsfile1 = "C:\ADMIN\VENUS\ANALISIS.xls"
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblProceso.Caption = "Abriendo , Archivo ANALISIS.xls . . . "
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open wsfile1, 0, True, 4, WPAS, WPAS

Return

Exit Sub
CANCELA:
  FrmImpC1.pantalla.Enabled = True
  FrmImpC1.pantalla.Caption = "Por &Pantalla"
  FrmImpC1.lblProceso.Visible = False
  FrmImpC1.ProgBar.Visible = False
  pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.Application.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
OJO:
If Err.Number = 70 Then
  MsgBox "Hoja de Calculo : " & wsfile1 & "  esta Abierta debe cerrar para Procesar Nuevamente ", 48, Pub_Titulo
  GoTo CANCELA
End If
Exit Sub
FINTODO:
 MsgBox Err.Description + " Reintente Nuevamente ..", 48, Pub_Titulo
 Resume Next
 GoTo CANCELA

End Sub

Public Function EGRE_CAJA(Optional WF) As Currency
Dim PS_1 As rdoQuery
Dim llave_1 As rdoResultset
Dim WS_SUMA As Currency
pub_cadena = "SELECT * FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_DIA = ?  AND ALL_FLAG_EXT <> 'E' AND ALL_FLAG_EXT <> 'X'  AND ALL_SIGNO_CAJA = -1 AND ALL_SIGNO_CCM = 0 ORDER BY ALL_CODCIA, ALL_CODVEN"
Set PS_1 = CN.CreateQuery("", pub_cadena)
PS_1(0) = 0
PS_1(0) = LK_FECHA_DIA
Set llave_1 = PS_1.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
PS_1(0) = LK_CODCIA
PS_1(1) = CDate(WF)
llave_1.Requery
If llave_1.EOF Then
  EGRE_CAJA = WS_SUMA
  Exit Function
End If
Do Until llave_1.EOF
  WS_SUMA = WS_SUMA + llave_1!ALL_IMPORTE
  llave_1.MoveNext
Loop
EGRE_CAJA = WS_SUMA
End Function
Public Sub LIBRO_DIARIO()
' *** REPORTES DE NUCLEOS
'On Error GoTo CANCELA
Dim wsglosita As String
Dim xF As Integer
Dim PSCOX_LLAVE As rdoQuery
Dim COX_LLAVE  As rdoResultset
Dim WS_NRO_MOV As Integer
Dim ws_nro_voucher As Integer
Dim WS_FECHA1 As Date
Dim WS_FECHA2 As Date
Dim WS_SAL_CUENTA As Currency
Dim WS_CUENTA As String * 12
Dim WS_TOT_IMPORTE_S As Currency
Dim WS_FLAG As String * 1
Dim WS_MAYOR As String
Dim XFF As Integer
Dim WS_SAL_CUENTA2 As Currency
Dim WS_SAL_DEB1 As Currency
Dim WS_SAL_DEB2 As Currency
Dim WS_SAL_HAB1 As Currency
Dim WS_SAL_HAB2 As Currency
Dim wdh As String * 1
Dim wfila_ult As Integer
Dim CTA_10101_D As Currency
Dim CTA_10101_H As Currency
Dim ws_asiento As Currency


'SON_FECHAS txtCampo1, txtCampo2
If periodos.Value = 1 Then
  REP_FECHA1 = Left(Trim(FECHA1.Text), 10)
  REP_FECHA2 = Right(Trim(FECHA1.Text), 10)
Else
  If Not SON_FECHAS(txtCampo1, txtCampo2) Then
   GoTo CANCELA
  End If
End If


pub_cadena = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND COV_FECHA_VOUCHER >= ? AND COV_FECHA_VOUCHER <= ?    ORDER BY COV_CODCIA, COV_FECHA_VOUCHER , COV_NRO_VOUCHER, COV_DH "
Set PSCOX_LLAVE = CN.CreateQuery("", pub_cadena)
PSCOX_LLAVE(0) = 0
PSCOX_LLAVE(1) = LK_FECHA_DIA
PSCOX_LLAVE(2) = LK_FECHA_DIA
Set COX_LLAVE = PSCOX_LLAVE.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
'If Not SON_FECHAS(txtCampo1, txtCampo2) Then
'  GoTo CANCELA
'End If

PSCOX_LLAVE(0) = LK_CODCIA
PSCOX_LLAVE(1) = REP_FECHA1
PSCOX_LLAVE(2) = REP_FECHA2

COX_LLAVE.Requery
If COX_LLAVE.EOF Then
  Screen.MousePointer = 0
  MsgBox "NO Existen datos para la Consulta ..", 48, Pub_Titulo
  Exit Sub
End If


FrmImpC1.ProgBar.Min = 0
FrmImpC1.ProgBar.Max = COX_LLAVE.RowCount
FrmImpC1.ProgBar.Value = 0
FrmImpC1.ProgBar.Visible = True
DoEvents
FrmImpC1.lblProceso.Visible = True
FrmImpC1.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
C1 = 1
xl.Worksheets(1).Activate
F1 = 4
'xl.Cells(F1, 8) = "Caja y Bancos "
xF = 4
wsglosita = ""
SQ_OPER = 1
XFF = 0
wdh = ""
WS_SAL_DEB1 = 0
WS_SAL_HAB1 = 0
CTA_10101_D = 0
CTA_10101_H = 0
FrmImpC1.lblProceso.Caption = "Procesando. . . "
DoEvents
F1 = F1 + 1
ws_asiento = COX_LLAVE!COV_NRO_VOUCHER
xl.Cells(F1, 3) = "'                       " & Format(ws_asiento, "0")
'xl.Application.Visible = True
Do Until COX_LLAVE.EOF ' loop 1
   FrmImpC1.ProgBar.Value = FrmImpC1.ProgBar.Value + 1
      If Trim(COX_LLAVE!COV_CODCTA) <> Trim(WS_CUENTA) Or wdh <> COX_LLAVE!COV_DH Then
       If WS_MAYOR <> Left(COX_LLAVE!COV_CODCTA, 2) Or wdh <> COX_LLAVE!COV_DH Then
          If WS_SAL_CUENTA <> 0 Then xl.Cells(xF, C1 + 4) = Format(WS_SAL_CUENTA, "0.00;(0.00)")
          WS_SAL_CUENTA = 0
          If XFF <> 0 Then
            If wdh = "H" Then
               xl.Cells(XFF, C1 + 5 + 1) = Format(WS_SAL_CUENTA2, "0.00;(0.00)")
            Else
               xl.Cells(XFF, C1 + 5) = Format(WS_SAL_CUENTA2, "0.00;(0.00)")
            End If
          End If
          If wdh = "H" Then
              WS_SAL_DEB2 = WS_SAL_DEB2 + WS_SAL_CUENTA2
          Else
              WS_SAL_HAB2 = WS_SAL_HAB2 + WS_SAL_CUENTA2
          End If
          WS_SAL_CUENTA2 = 0
          If COX_LLAVE!COV_DH = "D" And C1 = 1 Then
          '      wfila_ult = F1
          '      F1 = 4
          '      C1 = C1 + 7
          End If
          If ws_asiento <> COX_LLAVE!COV_NRO_VOUCHER Then
               F1 = F1 + 1
               xl.Cells(F1, C1 + 2) = wsglosita
               F1 = F1 + 2
               ws_asiento = COX_LLAVE!COV_NRO_VOUCHER
               xl.Cells(F1, 3) = "'                       " & Format(ws_asiento, "0")
          End If
          F1 = F1 + 1
 '              xl.Application.Visible = True
          xl.Cells(F1, C1) = "'" & Trim(Left(COX_LLAVE!COV_CODCTA, 2))
          PUB_CUENTA = Trim(Left(COX_LLAVE!COV_CODCTA, 2))
          LEER_COM_LLAVE
          xl.Cells(F1, C1 + 2) = Trim(com_llave!com_DESCRIPCION)
          XFF = F1
       End If
       If WS_SAL_CUENTA <> 0 Then xl.Cells(xF, C1 + 4) = Format(WS_SAL_CUENTA, "0.00;(0.00)")
       ' f1 = f1 + 1
        'PUB_CUENTA = COX_LLAVE!COV_CODCTA
        'LEER_COM_LLAVE
        'xl.Cells(f1, C1 + 2) = Trim(com_llave!com_DESCRIPCION)
                    'xl.Cells(F1, C1 + 1) = "'" & Trim(COX_LLAVE!COV_CODCTA)
        'xl.Cells(f1, C1 + 1) = "'" & Left(Trim(com_llave!com_cuenta), wCOM_NIVEL(NIVEL_MAX - 1))
        
        xF = F1
        WS_SAL_CUENTA = 0
     End If
     
     If COX_LLAVE!COV_DH = "D" And C1 = 1 Then
   '     wfila_ult = F1
   '     F1 = 4
   '     C1 = C1 + 7
     End If
     F1 = F1 + 1
     'xl.Cells(F1, C1 + 1) = "'" & Left(Trim(COX_LLAVE!COV_CODCTA), Len(wCOM_NIVEL(NIVEL_MAX - 1))) '"'" & Format(COX_LLAVE!cov_FECHA_VOUCHER, "dd.mm")
     xl.Cells(F1, C1 + 1) = "'" & Trim(COX_LLAVE!COV_CODCTA) '"'" & Format(COX_LLAVE!cov_FECHA_VOUCHER, "dd.mm")
     PUB_CUENTA = COX_LLAVE!COV_CODCTA
     LEER_COM_LLAVE
     If com_llave.EOF Then
       MsgBox " la Cuenta " & PUB_CUENTA & " NO EXISTE ", 48, Pub_Titulo
     End If
     xl.Cells(F1, C1 + 2) = Trim(com_llave!com_DESCRIPCION)
     wsglosita = Trim(COX_LLAVE!COV_glosa)
     xl.Cells(F1, C1 + 3) = Format(COX_LLAVE!COV_IMPORTE, "0.00;(0.00)")
     WS_SAL_CUENTA = WS_SAL_CUENTA + COX_LLAVE!COV_IMPORTE
     WS_SAL_CUENTA2 = WS_SAL_CUENTA2 + COX_LLAVE!COV_IMPORTE
     If COX_LLAVE!COV_DH = "H" Then
       WS_SAL_DEB1 = WS_SAL_DEB1 + COX_LLAVE!COV_IMPORTE
     Else
      WS_SAL_HAB1 = WS_SAL_HAB1 + COX_LLAVE!COV_IMPORTE
     End If
     WS_CUENTA = COX_LLAVE!COV_CODCTA
     WS_MAYOR = Left(COX_LLAVE!COV_CODCTA, 2)
     wdh = COX_LLAVE!COV_DH
OTRA_CTA:

    COX_LLAVE.MoveNext
Loop

If XFF <> 0 Then
 ' xl.Cells(XFF, C1 + 5) = Format(WS_SAL_CUENTA2, "0.00;(0.00)")
  If wdh = "H" Then
   xl.Cells(XFF, C1 + 5 + 1) = Format(WS_SAL_CUENTA2, "0.00;(0.00)")
  Else
   xl.Cells(XFF, C1 + 5) = Format(WS_SAL_CUENTA2, "0.00;(0.00)")
  End If
End If
xl.Cells(xF, C1 + 4) = Format(WS_SAL_CUENTA, "0.00;(0.00)")
If wdh = "H" Then
    WS_SAL_DEB2 = WS_SAL_DEB2 + WS_SAL_CUENTA2
Else
     WS_SAL_HAB2 = WS_SAL_HAB2 + WS_SAL_CUENTA2
End If

If WS_SAL_DEB1 <> WS_SAL_DEB2 Then
 MsgBox "Verificar Saldos  del Debe No Cuadra  !!! Diferencia = " & WS_SAL_DEB1 - WS_SAL_DEB2, 48, Pub_Titulo
End If
If WS_SAL_HAB1 <> WS_SAL_HAB2 Then
 MsgBox "Verificar Saldos  del Haber No Cuadra  !!! Diferencia = " & WS_SAL_DEB1 - WS_SAL_DEB2, 48, Pub_Titulo
End If

If wfila_ult >= F1 Then
  F1 = wfila_ult + 1
Else
  F1 = F1 + 1
End If
wranF = "F" & F1 + 1
wran1 = "F" & 5
wran2 = "F" & F1
xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
WS_SAL_DEB1 = Val(xl.Range(wranF).Text)
xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
wranF = "G" & F1 + 1
wran1 = "G" & 5
wran2 = "G" & F1
xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
WS_SAL_HAB1 = Val(xl.Range(wranF).Text)
If WS_SAL_DEB1 <> WS_SAL_HAB1 Then
  MsgBox "Diferencia No Cuadrada = " & Format(WS_SAL_DEB1 - WS_SAL_HAB1, "##,##0.00"), 48, Pub_Titulo
End If

'xl.Cells(F1, 4) = Format(WS_SAL_DEB1, "0.00;(0.00)")
'xl.Cells(F1, 4).Borders.Item(xlEdgeTop).LineStyle = 1
'xl.Cells(F1, 5) = Format(WS_SAL_DEB2, "0.00;(0.00)")
'xl.Cells(F1, 5).Borders.Item(xlEdgeTop).LineStyle = 1

'xl.Cells(F1, 11) = Format(WS_SAL_HAB1, "0.00;(0.00)")
'xl.Cells(F1, 11).Borders.Item(xlEdgeTop).LineStyle = 1
'xl.Cells(F1, 12) = Format(WS_SAL_HAB2, "0.00;(0.00)")
'xl.Cells(F1, 12).Borders.Item(xlEdgeTop).LineStyle = 1

xl.Cells(1, 1) = Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption))
xl.Cells(3, 1) = "DIARIO  - Del " & Format(REP_FECHA1, "dd/mm/yyyy") & " al " & Format(REP_FECHA2, "dd/mm/yyyy")

xl.DisplayAlerts = False
xl.Worksheets(1).Protect PUB_CLAVE
xl.Application.Visible = True

xcuenta = 0
Screen.MousePointer = 0
FrmImpC1.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
DoEvents
xl.Application.Visible = True
DoEvents
FrmImpC1.lblProceso.Visible = False
FrmImpC1.ProgBar.Visible = False
Set xl = Nothing
Screen.MousePointer = 0
FrmImpC1.pantalla.Enabled = True
FrmImpC1.pantalla.Caption = "Por &Pantalla"
FrmImpC1.lblProceso.Visible = False

Exit Sub
    


WEXCEL:
  Dim xlchart As Chart
  'Dim wranF, wran1, wran2, WPAS
  
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  FrmImpC1.lblProceso.Caption = "Abriendo , Archivo Ventas.xls . . . "
  DoEvents
  WPAS = "131296"
  xl.Workbooks.Open "C:\ADMIN\CONTABILIDAD\LIBRO_DIARIO.xls", 0, True, 4, WPAS, WPAS
Return



'*** RUTINAS PARA IMPRIMIR



WPROGRESO:

Return

Exit Sub
CANCELA:
  MsgBox "Verificar Datos ,e Intente Nuevamente..", 48, Pub_Titulo
  FrmImpC1.pantalla.Enabled = True
  FrmImpC1.pantalla.Caption = "Por &Pantalla"
  FrmImpC1.lblProceso.Visible = False
  xl.Application.Visible = True
  Set xl = Nothing
  Screen.MousePointer = 0

Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImpC1

End Sub


Public Sub LIBRO_DIARIO_en_Costruccion()
' *** REPORTES DE NUCLEOS
'On Error GoTo CANCELA
Dim XFF2 As Integer
Dim wsglosita As String
Dim xF As Integer
Dim PSCOX_LLAVE As rdoQuery
Dim COX_LLAVE  As rdoResultset
Dim WS_NRO_MOV As Integer
Dim ws_nro_voucher As Integer
Dim WS_FECHA1 As Date
Dim WS_FECHA2 As Date
Dim WS_SAL_CUENTA As Currency
Dim WS_CUENTA As String * 12
Dim WS_TOT_IMPORTE_S As Currency
Dim WS_FLAG As String * 1
Dim WS_MAYOR As String
Dim XFF As Integer
Dim WS_SAL_CUENTA2 As Currency
Dim WS_SAL_DEB1 As Currency
Dim WS_SAL_DEB2 As Currency
Dim WS_SAL_HAB1 As Currency
Dim WS_SAL_HAB2 As Currency
Dim wdh As String * 1
Dim wfila_ult As Integer
Dim CTA_10101_D As Currency
Dim CTA_10101_H As Currency
Dim ws_asiento As Currency


'SON_FECHAS txtCampo1, txtCampo2
If Not SON_FECHAS(txtCampo1, txtCampo2) Then
  GoTo CANCELA
End If

pub_cadena = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND COV_FECHA_VOUCHER >= ? AND COV_FECHA_VOUCHER <= ?    ORDER BY COV_nro_voucher, COV_DH ASC , COV_CODCTA ASC"
Set PSCOX_LLAVE = CN.CreateQuery("", pub_cadena)
Set COX_LLAVE = PSCOX_LLAVE.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
'If Not SON_FECHAS(txtCampo1, txtCampo2) Then
'  GoTo CANCELA
'End If

PSCOX_LLAVE(0) = LK_CODCIA
PSCOX_LLAVE(1) = REP_FECHA1
PSCOX_LLAVE(2) = REP_FECHA2

COX_LLAVE.Requery
If COX_LLAVE.EOF Then
  Screen.MousePointer = 0
  MsgBox "NO Existen datos para la Consulta ..", 48, Pub_Titulo
  Exit Sub
End If


FrmImpC1.ProgBar.Min = 0
FrmImpC1.ProgBar.Max = COX_LLAVE.RowCount
FrmImpC1.ProgBar.Value = 0
FrmImpC1.ProgBar.Visible = True
DoEvents
FrmImpC1.lblProceso.Visible = True
FrmImpC1.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
C1 = 1
xl.Worksheets(1).Activate
F1 = 4
'xl.Cells(F1, 8) = "Caja y Bancos "
xF = 4
wsglosita = ""
SQ_OPER = 1
XFF = 0
wdh = ""
WS_SAL_DEB1 = 0
WS_SAL_HAB1 = 0
CTA_10101_D = 0
CTA_10101_H = 0
WS_SAL_CUENTA = 0
FrmImpC1.lblProceso.Caption = "Procesando. . . "
DoEvents
F1 = F1 + 1
ws_asiento = COX_LLAVE!COV_NRO_VOUCHER
xl.Cells(F1, 4) = "'" & Format(ws_asiento, "0")
xl.Application.Visible = True
Do Until COX_LLAVE.EOF ' loop 1
   FrmImpC1.ProgBar.Value = FrmImpC1.ProgBar.Value + 1
      If Trim(COX_LLAVE!COV_CODCTA) <> Trim(WS_CUENTA) Or wdh <> COX_LLAVE!COV_DH Then
       If WS_MAYOR <> Left(COX_LLAVE!COV_CODCTA, 2) Or wdh <> COX_LLAVE!COV_DH Then
          If WS_SAL_CUENTA <> 0 Then
            '  xl.Cells(xF, C1 + 4) = Format(WS_SAL_CUENTA, "0.00;(0.00)")
              
              'xl.Cells(xF + 1, C1 + 4) = Format(WS_SAL_CUENTA, "0.00;(0.00)")
          End If
          WS_SAL_CUENTA = 0
          If XFF <> 0 Then
            If wdh = "H" Then
               xl.Cells(XFF, C1 + 5 + 1) = Format(WS_SAL_CUENTA2, "0.00;(0.00)")
            Else
               xl.Cells(XFF, C1 + 5) = Format(WS_SAL_CUENTA2, "0.00;(0.00)")
            End If
            PUB_CUENTA = Trim(Left(WS_CUENTA, 3))
            LEER_COM_LLAVE
            xl.Cells(XFF2, C1 + 1) = Trim(PUB_CUENTA)
            xl.Cells(XFF2, C1 + 2) = Trim(com_llave!com_DESCRIPCION)
            xl.Cells(XFF2, C1 + 4) = Format(WS_SAL_CUENTA2, "0.00;(0.00)")
          End If
          If wdh = "H" Then
              WS_SAL_DEB2 = WS_SAL_DEB2 + WS_SAL_CUENTA2
          Else
              WS_SAL_HAB2 = WS_SAL_HAB2 + WS_SAL_CUENTA2
          End If
          WS_SAL_CUENTA2 = 0
          If COX_LLAVE!COV_DH = "D" And C1 = 1 Then
          '      wfila_ult = F1
          '      F1 = 4
          '      C1 = C1 + 7
          End If
          If ws_asiento <> COX_LLAVE!COV_NRO_VOUCHER Then
               F1 = F1 + 1
               xl.Cells(F1, C1 + 2) = wsglosita
               F1 = F1 + 2
               ws_asiento = COX_LLAVE!COV_NRO_VOUCHER
               xl.Cells(F1, 4) = "'" & Format(ws_asiento, "0")
          End If
          F1 = F1 + 1
 '              xl.Application.Visible = True
          xl.Cells(F1, C1) = "'" & Trim(Left(COX_LLAVE!COV_CODCTA, 2))
          PUB_CUENTA = Trim(Left(COX_LLAVE!COV_CODCTA, 2))
          LEER_COM_LLAVE
          xl.Cells(F1, C1 + 2) = Trim(com_llave!com_DESCRIPCION)
          XFF = F1
          XFF2 = F1 + 1
          
       End If
       If WS_SAL_CUENTA <> 0 Then xl.Cells(xF, C1 + 4) = Format(WS_SAL_CUENTA, "0.00;(0.00)")
        'f1 = f1 + 1
        'PUB_CUENTA = COX_LLAVE!COV_CODCTA
        'LEER_COM_LLAVE
        'xl.Cells(f1, C1 + 2) = Trim(com_llave!com_DESCRIPCION)
        '            'xl.Cells(F1, C1 + 1) = "'" & Trim(COX_LLAVE!COV_CODCTA)
        'xl.Cells(f1, C1 + 1) = "'" & Left(Trim(com_llave!com_cuenta), wCOM_NIVEL(NIVEL_MAX - 1))
       '
        xF = F1
        WS_SAL_CUENTA = 0
       End If
     
     If COX_LLAVE!COV_DH = "D" And C1 = 1 Then
   '     wfila_ult = F1
   '     F1 = 4
   '     C1 = C1 + 7
     End If
     F1 = F1 + 2
     'xl.Cells(F1, C1 + 1) = "'" & Left(Trim(COX_LLAVE!COV_CODCTA), Len(wCOM_NIVEL(NIVEL_MAX - 1))) '"'" & Format(COX_LLAVE!cov_FECHA_VOUCHER, "dd.mm")
     xl.Cells(F1, C1 + 1) = "'" & Trim(COX_LLAVE!COV_CODCTA) '"'" & Format(COX_LLAVE!cov_FECHA_VOUCHER, "dd.mm")
     PUB_CUENTA = COX_LLAVE!COV_CODCTA
     LEER_COM_LLAVE
     If com_llave.EOF Then
       MsgBox " la Cuenta " & PUB_CUENTA & " NO EXISTE ", 48, Pub_Titulo
     End If
     xl.Cells(F1, C1 + 2) = Trim(com_llave!com_DESCRIPCION)
     wsglosita = Trim(COX_LLAVE!COV_glosa)
     xl.Cells(F1, C1 + 3) = Format(COX_LLAVE!COV_IMPORTE, "0.00;(0.00)")
     WS_SAL_CUENTA = WS_SAL_CUENTA + COX_LLAVE!COV_IMPORTE
     WS_SAL_CUENTA2 = WS_SAL_CUENTA2 + COX_LLAVE!COV_IMPORTE
     If COX_LLAVE!COV_DH = "H" Then
       WS_SAL_DEB1 = WS_SAL_DEB1 + COX_LLAVE!COV_IMPORTE
     Else
      WS_SAL_HAB1 = WS_SAL_HAB1 + COX_LLAVE!COV_IMPORTE
     End If
     WS_CUENTA = COX_LLAVE!COV_CODCTA
     WS_MAYOR = Left(COX_LLAVE!COV_CODCTA, 2)
     wdh = COX_LLAVE!COV_DH
OTRA_CTA:

    COX_LLAVE.MoveNext
Loop

If XFF <> 0 Then
 ' xl.Cells(XFF, C1 + 5) = Format(WS_SAL_CUENTA2, "0.00;(0.00)")
  If wdh = "H" Then
   xl.Cells(XFF, C1 + 5 + 1) = Format(WS_SAL_CUENTA2, "0.00;(0.00)")
  Else
   xl.Cells(XFF, C1 + 5) = Format(WS_SAL_CUENTA2, "0.00;(0.00)")
  End If
End If
xl.Cells(xF, C1 + 4) = Format(WS_SAL_CUENTA, "0.00;(0.00)")
If wdh = "H" Then
    WS_SAL_DEB2 = WS_SAL_DEB2 + WS_SAL_CUENTA2
Else
     WS_SAL_HAB2 = WS_SAL_HAB2 + WS_SAL_CUENTA2
End If

If WS_SAL_DEB1 <> WS_SAL_DEB2 Then
 MsgBox "Verificar Saldos  del Debe No Cuadra  !!! Diferencia = " & WS_SAL_DEB1 - WS_SAL_DEB2, 48, Pub_Titulo
End If
If WS_SAL_HAB1 <> WS_SAL_HAB2 Then
 MsgBox "Verificar Saldos  del Haber No Cuadra  !!! Diferencia = " & WS_SAL_DEB1 - WS_SAL_DEB2, 48, Pub_Titulo
End If

If wfila_ult >= F1 Then
  F1 = wfila_ult + 1
Else
  F1 = F1 + 1
End If
wranF = "F" & F1 + 1
wran1 = "F" & 5
wran2 = "F" & F1
xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
WS_SAL_DEB1 = Val(xl.Range(wranF).Text)
xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
wranF = "G" & F1 + 1
wran1 = "G" & 5
wran2 = "G" & F1
xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
xl.Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
WS_SAL_HAB1 = Val(xl.Range(wranF).Text)
If WS_SAL_DEB1 <> WS_SAL_HAB1 Then
  MsgBox "Diferencia No Cuadrada = " & Format(WS_SAL_DEB1 - WS_SAL_HAB1, "##,##0.00"), 48, Pub_Titulo
End If

'xl.Cells(F1, 4) = Format(WS_SAL_DEB1, "0.00;(0.00)")
'xl.Cells(F1, 4).Borders.Item(xlEdgeTop).LineStyle = 1
'xl.Cells(F1, 5) = Format(WS_SAL_DEB2, "0.00;(0.00)")
'xl.Cells(F1, 5).Borders.Item(xlEdgeTop).LineStyle = 1

'xl.Cells(F1, 11) = Format(WS_SAL_HAB1, "0.00;(0.00)")
'xl.Cells(F1, 11).Borders.Item(xlEdgeTop).LineStyle = 1
'xl.Cells(F1, 12) = Format(WS_SAL_HAB2, "0.00;(0.00)")
'xl.Cells(F1, 12).Borders.Item(xlEdgeTop).LineStyle = 1


xl.Cells(3, 1) = "DIARIO MES " & Format(LK_FECHA_COP2, "mmmm  yyyy")
xl.DisplayAlerts = False
xl.Worksheets(1).Protect PUB_CLAVE
xl.Application.Visible = True

xcuenta = 0
Screen.MousePointer = 0
FrmImpC1.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
DoEvents
xl.Application.Visible = True
DoEvents
FrmImpC1.lblProceso.Visible = False
FrmImpC1.ProgBar.Visible = False
Set xl = Nothing
Screen.MousePointer = 0
FrmImpC1.pantalla.Enabled = True
FrmImpC1.pantalla.Caption = "Por &Pantalla"
FrmImpC1.lblProceso.Visible = False

Exit Sub
    


WEXCEL:
  Dim xlchart As Chart
  'Dim wranF, wran1, wran2, WPAS
  
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  FrmImpC1.lblProceso.Caption = "Abriendo , Archivo Ventas.xls . . . "
  DoEvents
  WPAS = "131296"
  xl.Workbooks.Open "C:\ADMIN\CONTABILIDAD\LIBRO_DIARIO.xls", 0, True, 4, WPAS, WPAS
Return



'*** RUTINAS PARA IMPRIMIR



WPROGRESO:

Return

Exit Sub
CANCELA:
  MsgBox "Verificar Datos ,e Intente Nuevamente..", 48, Pub_Titulo
  FrmImpC1.pantalla.Enabled = True
  FrmImpC1.pantalla.Caption = "Por &Pantalla"
  FrmImpC1.lblProceso.Visible = False
  xl.Application.Visible = True
  Set xl = Nothing
  Screen.MousePointer = 0

Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImpC1

End Sub

Public Sub LIBRO_CAJA()
' *** REPORTES DE NUCLEOS
'On Error GoTo CANCELA
Dim ww_flag As String * 1
Dim xF As Integer
Dim PSCOX_LLAVE As rdoQuery
Dim COX_LLAVE  As rdoResultset
Dim WSGLOSA As String
Dim WS_NRO_MOV As Integer
Dim ws_nro_voucher As Integer
Dim WS_FECHA1 As Date
Dim WS_FECHA2 As Date
Dim WS_SAL_CUENTA As Currency
Dim WS_CUENTA As String * 12
Dim WS_TOT_IMPORTE_S As Currency
Dim WS_FLAG As String * 1
Dim WS_MAYOR As String
Dim WS_SAL_ANTERIOR As Currency
Dim XFF As Integer
Dim WS_SAL_CUENTA2 As Currency
Dim WS_SAL_DEB1 As Currency
Dim WS_SAL_DEB2 As Currency
Dim WS_SAL_HAB1 As Currency
Dim WS_SAL_HAB2 As Currency
Dim wdh As String * 1
Dim wfila_ult As Integer
Dim CTA_10101_D As Currency
Dim CTA_10101_H As Currency
Dim wscodcia  As String * 2
ww_flag = ""
WS_SAL_ANTERIOR = 0
'SON_FECHAS txtCampo1, txtCampo2
If Not SON_FECHAS(txtCampo1, txtCampo2) Then
  GoTo CANCELA
End If

If periodos.Value = 1 Then
  REP_FECHA1 = Left(Trim(FECHA1.Text), 10)
  REP_FECHA2 = Right(Trim(FECHA1.Text), 10)
Else
  If CDate(REP_FECHA1) <> LK_FECHA_COP1 Then che1.Value = 0
  If CDate(REP_FECHA2) <> LK_FECHA_COP2 Then che1.Value = 0
End If

pub_cadena = "SELECT * FROM COMOX WHERE COX_CODCIA = ? AND COX_FECHA_VOUCHER >= ? AND COX_FECHA_VOUCHER <= ?  AND COX_IDENTIFICADOR = ?   ORDER BY COX_IDENTIFICADOR ,COX_DH DESC , COX_CODCTA ASC"
Set PSCOX_LLAVE = CN.CreateQuery("", pub_cadena)
PSCOX_LLAVE(0) = 0
PSCOX_LLAVE(1) = LK_FECHA_DIA
PSCOX_LLAVE(2) = LK_FECHA_DIA
PSCOX_LLAVE(3) = 0
Set COX_LLAVE = PSCOX_LLAVE.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


wscodcia = LK_CODCIA
If LK_EMP_PTO = "A" Then
 wscodcia = "00"
End If

PSCOX_LLAVE(0) = wscodcia
PSCOX_LLAVE(1) = REP_FECHA1
PSCOX_LLAVE(2) = REP_FECHA2
PSCOX_LLAVE(3) = "C"
COX_LLAVE.Requery
If COX_LLAVE.EOF Then
  Screen.MousePointer = 0
  MsgBox "NO Existen movimientos para caja ..", 48, Pub_Titulo
  Exit Sub
  ww_flag = "A"
  GoTo SALTA_1
End If

PUB_CUENTA = "10101"
PUB_CODCIA = LK_CODCIA
If periodos.Value = 0 Then
  SQ_OPER = 1
  LEER_COM_LLAVE
  If com_llave.EOF Then
    MsgBox "Definir Cuenta Contable ...", 48, Pub_Titulo
    Exit Sub
  Else
    WS_SAL_ANTERIOR = Nulo_Valor0(com_llave!COM_DEB_ANO) - Nulo_Valor0(com_llave!COM_HAB_ANO)
  End If
Else
    PSCOH_LLAVE(0) = Left(Trim(FECHA1.Text), 10)
    PSCOH_LLAVE(1) = Right(Trim(FECHA1.Text), 10)
    PSCOH_LLAVE(2) = PUB_CODCIA
    PSCOH_LLAVE(3) = PUB_CUENTA
    coh_llave.Requery
    If coh_llave.EOF Then
      MsgBox "Definicion de Cuenta  NO Existe ...", 48, Pub_Titulo
    Else
     WS_SAL_ANTERIOR = ((Val(coh_llave!COH_DEB_ANO)) * coh_llave!COH_SIGNO_D) + ((Val(coh_llave!COH_HAB_ANO)) * coh_llave!COH_SIGNO_H)
    End If
End If

FrmImpC1.ProgBar.Min = 0
FrmImpC1.ProgBar.Max = COX_LLAVE.RowCount
FrmImpC1.ProgBar.Value = 0
FrmImpC1.ProgBar.Visible = True
DoEvents
FrmImpC1.lblProceso.Visible = True
FrmImpC1.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
C1 = 1
xl.Worksheets(1).Activate
F1 = 4
'xl.Cells(F1, 8) = "Caja y Bancos "
xF = 4
XFF = 0
wdh = ""
WS_SAL_DEB1 = 0
WS_SAL_HAB1 = 0
CTA_10101_D = 0
CTA_10101_H = 0
FrmImpC1.lblProceso.Caption = "Procesando. . . "
DoEvents
xl.Cells(3, 1) = "Del " & Format(REP_FECHA1, "dd/mm/yyyy") & " al " & Format(REP_FECHA2, "dd/mm/yyyy")
xl.Cells(xF, C1 + 3) = "Saldo Inicial  :"
xl.Cells(xF, C1 + 5) = Format(WS_SAL_ANTERIOR, "0.00;(0.00)")
Do Until COX_LLAVE.EOF ' loop 1
  FrmImpC1.ProgBar.Value = FrmImpC1.ProgBar.Value + 1
  If Trim(COX_LLAVE!cox_codcta) = "10101" Then
    If COX_LLAVE!cox_DH = "D" Then
        CTA_10101_D = CTA_10101_D + COX_LLAVE!cox_IMPORTE
    Else
        CTA_10101_H = CTA_10101_H + COX_LLAVE!cox_IMPORTE
    End If
    
    GoTo OTRA_CTA
  End If

      If Trim(COX_LLAVE!cox_codcta) <> Trim(WS_CUENTA) Or wdh <> COX_LLAVE!cox_DH Then
       If WS_MAYOR <> Left(COX_LLAVE!cox_codcta, 2) Or wdh <> COX_LLAVE!cox_DH Then
          If WS_SAL_CUENTA <> 0 Then xl.Cells(xF, C1 + 4) = Format(WS_SAL_CUENTA, "0.00;(0.00)")
          WS_SAL_CUENTA = 0
          If XFF <> 0 Then xl.Cells(XFF, C1 + 5) = Format(WS_SAL_CUENTA2, "0.00;(0.00)")
          If wdh = "H" Then
              WS_SAL_DEB2 = WS_SAL_DEB2 + WS_SAL_CUENTA2
          Else
              WS_SAL_HAB2 = WS_SAL_HAB2 + WS_SAL_CUENTA2
          End If
          WS_SAL_CUENTA2 = 0
          If COX_LLAVE!cox_DH = "D" And C1 = 1 Then
                wfila_ult = F1
                F1 = 4
                C1 = C1 + 7
          End If
           F1 = F1 + 1
           xl.Cells(F1, C1) = "'" & Trim(Left(COX_LLAVE!cox_codcta, 2))
           PUB_CUENTA = Trim(Left(COX_LLAVE!cox_codcta, 2))
           LEER_COM_LLAVE
           xl.Cells(F1, C1 + 2) = Trim(com_llave!com_DESCRIPCION)
           XFF = F1
       End If
        If WS_SAL_CUENTA <> 0 Then xl.Cells(xF, C1 + 4) = Format(WS_SAL_CUENTA, "0.00;(0.00)")
        F1 = F1 + 1
        PUB_CUENTA = COX_LLAVE!cox_codcta
        LEER_COM_LLAVE
        If com_llave.EOF Then
         MsgBox "Una cuenta  no Existe en el Plan Transaccion : " & COX_LLAVE!COX_NRO_VOUCHER & " " & COX_LLAVE!cox_codcta, 48, Pub_Titulo
         GoTo CANCELA
        Exit Sub
        End If
        xl.Cells(F1, C1 + 2) = Trim(com_llave!com_DESCRIPCION)
        xl.Cells(F1, C1 + 1) = "'" & Trim(COX_LLAVE!cox_codcta)
        'If COX_LLAVE!coX_DH = "D" Then
        '    WS_SAL_DEB1 = WS_SAL_DEB1 + WS_SAL_CUENTA
        'Else
        '    WS_SAL_HAB1 = WS_SAL_HAB1 + WS_SAL_CUENTA
        'End If
        xF = F1
        WS_SAL_CUENTA = 0
     End If
     
     
     If COX_LLAVE!cox_DH = "D" And C1 = 1 Then
        wfila_ult = F1
        F1 = 4
        C1 = C1 + 7
     End If
     F1 = F1 + 1
     xl.Cells(F1, C1 + 1) = "'" & Format(COX_LLAVE!coX_FECHA_VOUCHER, "dd.mm")
     xl.Cells(F1, C1 + 2) = Trim(COX_LLAVE!coX_GLOSA)
     xl.Cells(F1, C1 + 3) = Format(COX_LLAVE!cox_IMPORTE, "0.00;(0.00)")
     WS_SAL_CUENTA = WS_SAL_CUENTA + COX_LLAVE!cox_IMPORTE
     WS_SAL_CUENTA2 = WS_SAL_CUENTA2 + COX_LLAVE!cox_IMPORTE
     If COX_LLAVE!cox_DH = "H" Then
       WS_SAL_DEB1 = WS_SAL_DEB1 + COX_LLAVE!cox_IMPORTE
     Else
      WS_SAL_HAB1 = WS_SAL_HAB1 + COX_LLAVE!cox_IMPORTE
     End If
     WS_CUENTA = COX_LLAVE!cox_codcta
     WS_MAYOR = Left(COX_LLAVE!cox_codcta, 2)
     wdh = COX_LLAVE!cox_DH
OTRA_CTA:
     COX_LLAVE.MoveNext
Loop

If XFF <> 0 Then xl.Cells(XFF, C1 + 5) = Format(WS_SAL_CUENTA2, "0.00;(0.00)")
xl.Cells(xF, C1 + 4) = Format(WS_SAL_CUENTA, "0.00;(0.00)")
If wdh = "H" Then
    WS_SAL_DEB2 = WS_SAL_DEB2 + WS_SAL_CUENTA2
Else
     WS_SAL_HAB2 = WS_SAL_HAB2 + WS_SAL_CUENTA2
End If

If WS_SAL_DEB1 <> WS_SAL_DEB2 Then
 MsgBox "Verificar Saldos  del Debe No Cuadra  !!! Diferencia = " & WS_SAL_DEB1 - WS_SAL_DEB2, 48, Pub_Titulo
End If
If WS_SAL_HAB1 <> WS_SAL_HAB2 Then
 MsgBox "Verificar Saldos  del Haber No Cuadra  !!! Diferencia = " & WS_SAL_DEB1 - WS_SAL_DEB2, 48, Pub_Titulo
End If
Dim wsvalor As Currency
If wfila_ult >= F1 Then
  F1 = wfila_ult + 1
Else
  F1 = F1 + 1
End If
CTA_10101_D = CTA_10101_D + WS_SAL_ANTERIOR
WS_SAL_DEB2 = WS_SAL_DEB2 + WS_SAL_ANTERIOR
'xl.Visible = True
If (CTA_10101_D - CTA_10101_H) < 0 Then
  xl.Cells(F1, 6) = Format(WS_SAL_DEB2, "0.00;(0.00)")
  xl.Cells(F1, 6).Borders.Item(xlEdgeTop).LineStyle = 1
  F1 = F1 + 1
  xl.Cells(F1, 3) = "Saldo al " & Format(LK_FECHA_COP2, "dd/mm/yyyy")
  xl.Cells(F1, 6) = Format(Abs(CTA_10101_D - CTA_10101_H), "0.00;(0.00)")
  wsvalor = Val(xl.Cells(F1, 6))
  F1 = F1 + 1
  xl.Cells(F1, 6) = Format(WS_SAL_DEB2 + (Abs(CTA_10101_D - CTA_10101_H)), "0.00;(0.00)")
  xl.Cells(F1, 6).Borders.Item(xlEdgeTop).LineStyle = 1
  
  xl.Cells(F1, 13) = Format(WS_SAL_HAB2, "0.00;(0.00)")
  xl.Cells(F1, 13).Borders.Item(xlEdgeTop).LineStyle = 1
  
ElseIf (CTA_10101_D - CTA_10101_H) > 0 Then
  xl.Cells(F1, 13) = Format(WS_SAL_HAB2, "0.00;(0.00)")
  xl.Cells(F1, 13).Borders.Item(xlEdgeTop).LineStyle = 1
  F1 = F1 + 1
  xl.Cells(F1, 8) = "Saldo al " & Format(LK_FECHA_COP2, "dd/mm/yyyy")
  xl.Cells(F1, 13) = Format(Abs(CTA_10101_D - CTA_10101_H), "0.00;(0.00)")
  wsvalor = Val(xl.Cells(F1, 13))
  F1 = F1 + 1
  xl.Cells(F1, 13) = Format(WS_SAL_HAB2 + (Abs(CTA_10101_D - CTA_10101_H)), "0.00;(0.00)")
  xl.Cells(F1, 13).Borders.Item(xlEdgeTop).LineStyle = 1
  
  xl.Cells(F1, 6) = Format(WS_SAL_DEB2, "0.00;(0.00)")
  xl.Cells(F1, 6).Borders.Item(xlEdgeTop).LineStyle = 1
End If

'If WS_SAL_DEB1 <> WS_SAL_HAB1 Then
'  MsgBox "Saldo de Caja y Bancos = " & Format(WS_SAL_DEB1 - WS_SAL_HAB1, "##,##0.00"), 48, Pub_Titulo
  MsgBox "Saldo en Caja Y Bancos = " & Format(wsvalor, "0.00;(0.00)"), 48, Pub_Titulo
'End If

xl.DisplayAlerts = False
xl.Worksheets(1).Protect PUB_CLAVE

If che1.Value = 1 And periodos.Value = 0 Then
  FrmImpC1.lblProceso.Caption = "Procesando al Diario Contable . . "
  DoEvents
  GoSub PASA_CONTAB
  cop_llave.Requery
  cop_llave.Edit
  cop_llave!cop_FLAG_CAJA = "A"
  cop_llave.Update
End If
xcuenta = 0
Screen.MousePointer = 0
FrmImpC1.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
DoEvents
xl.Application.Visible = True
DoEvents
FrmImpC1.lblProceso.Visible = False
FrmImpC1.ProgBar.Visible = False
Set xl = Nothing
pasito:
Screen.MousePointer = 0
FrmImpC1.pantalla.Enabled = True
FrmImpC1.pantalla.Caption = "Por &Pantalla"
FrmImpC1.lblProceso.Visible = False

Exit Sub
    

PASA_CONTAB:

Dim wcta As String
Dim wflag1 As Integer
WSGLOSA = "Libro de Caja Bancos"
PSCOV_VOUCHER(0) = LK_CODCIA
PSCOV_VOUCHER(1) = LK_FECHA_COP1
PSCOV_VOUCHER(2) = LK_FECHA_COP2
cov_voucher.Requery
Do Until cov_voucher.EOF
    If cov_voucher!cov_flag_automatica = "1" Then cov_voucher.Delete
    cov_voucher.MoveNext
Loop

cov_voucher.Requery
If cov_voucher.EOF Then
 ws_nro_voucher = 0
Else
 cov_voucher.MoveLast
 ws_nro_voucher = cov_voucher!COV_NRO_VOUCHER
End If
ws_nro_voucher = ws_nro_voucher + 1
WS_NRO_MOV = 0
COX_LLAVE.Requery
wflag1 = 0
wcta = COX_LLAVE!cox_codcta
wdh = COX_LLAVE!cox_DH
WS_SAL_CUENTA = 0
Do Until COX_LLAVE.EOF
  If wdh <> COX_LLAVE!cox_DH Then
    ' GoSub GRABA
    ' wcta = COX_LLAVE!cox_codcta
    ' WS_SAL_CUENTA = 0
    ' WS_SAL_CUENTA = WS_SAL_CUENTA + COX_LLAVE!cox_importe
    ' wdh = COX_LLAVE!cox_DH
    ' GoTo OTRO_AS
  End If
  
  If wcta <> COX_LLAVE!cox_codcta Then
     GoSub GRABA
     wcta = COX_LLAVE!cox_codcta
     WS_SAL_CUENTA = 0
     WS_SAL_CUENTA = WS_SAL_CUENTA + COX_LLAVE!cox_IMPORTE
     wdh = COX_LLAVE!cox_DH
  Else
    WS_SAL_CUENTA = WS_SAL_CUENTA + COX_LLAVE!cox_IMPORTE
    wdh = COX_LLAVE!cox_DH
  End If
OTRO_AS:
  COX_LLAVE.MoveNext
Loop
GoSub GRABA
Dim WCUENTA

pub_cadena = "SELECT * FROM COMOX WHERE (COX_NRO_VOUCHER <> 2407 AND COX_NRO_VOUCHER <> 2409) AND COX_CODCIA = ? AND COX_FECHA_VOUCHER >= ? AND COX_FECHA_VOUCHER <= ?  AND COX_IDENTIFICADOR = ?   ORDER BY  COX_NRO_VOUCHER"
Set PSCOX_LLAVE = CN.CreateQuery("", pub_cadena)
PSCOX_LLAVE(0) = 0
PSCOX_LLAVE(1) = LK_FECHA_DIA
PSCOX_LLAVE(2) = LK_FECHA_DIA
PSCOX_LLAVE(3) = 0
Set COX_LLAVE = PSCOX_LLAVE.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

PSCOX_LLAVE(0) = LK_CODCIA
PSCOX_LLAVE(1) = REP_FECHA1
PSCOX_LLAVE(2) = REP_FECHA2
PSCOX_LLAVE(3) = "D"
COX_LLAVE.Requery
If COX_LLAVE.EOF Then
  GoTo SALTA_1
End If

WCUENTA = 0
WSGLOSA = "Libro Inventarios"
ws_nro_voucher = ws_nro_voucher + 1
Do Until COX_LLAVE.EOF
  WCUENTA = WCUENTA + 1
  If WCUENTA = 3 Then
    WCUENTA = 1
    ws_nro_voucher = ws_nro_voucher + 1
    WS_NRO_MOV = 0
  End If
  wcta = COX_LLAVE!cox_codcta
  wdh = COX_LLAVE!cox_DH
  WS_SAL_CUENTA = COX_LLAVE!cox_IMPORTE
  GoSub GRABA
  COX_LLAVE.MoveNext
Loop

SALTA_1:
pub_cadena = "SELECT * FROM COMOX WHERE COX_CODCIA = ? AND COX_FECHA_VOUCHER >= ? AND COX_FECHA_VOUCHER <= ?  AND COX_IDENTIFICADOR = ?   ORDER BY COX_IDENTIFICADOR, COX_CODCTA ASC"
Set PSCOX_LLAVE = CN.CreateQuery("", pub_cadena)
PSCOX_LLAVE(0) = 0
PSCOX_LLAVE(1) = LK_FECHA_DIA
PSCOX_LLAVE(2) = LK_FECHA_DIA
PSCOX_LLAVE(3) = 0
Set COX_LLAVE = PSCOX_LLAVE.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

PSCOX_LLAVE(0) = LK_CODCIA
PSCOX_LLAVE(1) = REP_FECHA1
PSCOX_LLAVE(2) = REP_FECHA2
PSCOX_LLAVE(3) = "A"
COX_LLAVE.Requery
If COX_LLAVE.EOF Then
  GoTo SALTA2
End If
WS_SAL_CUENTA = 0
WCUENTA = 0
WSGLOSA = "Fondo Fijo"
ws_nro_voucher = ws_nro_voucher + 1
WS_MAYOR = Trim(COX_LLAVE!cox_codcta)
WS_NRO_MOV = 0
WS_SAL_DEB1 = 0
WS_SAL_HAB1 = 0
Do Until COX_LLAVE.EOF
  If Trim(WS_MAYOR) = Trim(COX_LLAVE!cox_codcta) Then
     If COX_LLAVE!cox_DH = "D" Then
        WS_SAL_DEB1 = WS_SAL_DEB1 + COX_LLAVE!cox_IMPORTE
     Else
         WS_SAL_HAB1 = WS_SAL_HAB1 + COX_LLAVE!cox_IMPORTE
     End If
     WS_MAYOR = Trim(COX_LLAVE!cox_codcta)
  Else
    If WS_SAL_DEB1 <> 0 Then
      WS_NRO_MOV = WS_NRO_MOV + 1
      wcta = WS_MAYOR
      wdh = "D"
      WS_SAL_CUENTA = WS_SAL_DEB1
     GoSub GRABA
    End If
    If WS_SAL_HAB1 <> 0 Then
      WS_NRO_MOV = WS_NRO_MOV + 1
      wcta = WS_MAYOR
      wdh = "H"
      WS_SAL_CUENTA = WS_SAL_HAB1
      GoSub GRABA
    End If
    WS_SAL_DEB1 = 0
    WS_SAL_HAB1 = 0
    If COX_LLAVE!cox_DH = "D" Then
        WS_SAL_DEB1 = WS_SAL_DEB1 + COX_LLAVE!cox_IMPORTE
    Else
         WS_SAL_HAB1 = WS_SAL_HAB1 + COX_LLAVE!cox_IMPORTE
    End If
    WS_MAYOR = Trim(COX_LLAVE!cox_codcta)
  End If
  
  COX_LLAVE.MoveNext
Loop
If WS_SAL_DEB1 <> 0 Then
  WS_NRO_MOV = WS_NRO_MOV + 1
  wcta = WS_MAYOR
  wdh = "D"
  WS_SAL_CUENTA = WS_SAL_DEB1
 GoSub GRABA
End If
If WS_SAL_HAB1 <> 0 Then
  WS_NRO_MOV = WS_NRO_MOV + 1
  wcta = WS_MAYOR
  wdh = "H"
  WS_SAL_CUENTA = WS_SAL_DEB1
  GoSub GRABA
End If
   
SALTA2:
Dim wsdh As String
pub_cadena = "SELECT * FROM COMOX WHERE (COX_NRO_VOUCHER = 2407 OR COX_NRO_VOUCHER = 2409) AND COX_CODCIA = ? AND COX_FECHA_VOUCHER >= ? AND COX_FECHA_VOUCHER <= ?  AND COX_IDENTIFICADOR = ?   ORDER BY  COX_DH"
Set PSCOX_LLAVE = CN.CreateQuery("", pub_cadena)
PSCOX_LLAVE(0) = 0
PSCOX_LLAVE(1) = LK_FECHA_DIA
PSCOX_LLAVE(2) = LK_FECHA_DIA
PSCOX_LLAVE(3) = 0
Set COX_LLAVE = PSCOX_LLAVE.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

PSCOX_LLAVE(0) = LK_CODCIA
PSCOX_LLAVE(1) = REP_FECHA1
PSCOX_LLAVE(2) = REP_FECHA2
PSCOX_LLAVE(3) = "D"
COX_LLAVE.Requery
If COX_LLAVE.EOF Then
  Return
End If
WCUENTA = 0
WS_SAL_CUENTA = 0
WSGLOSA = "Libro Inventarios - Envio / Recepción "
WS_NRO_MOV = 0
ws_nro_voucher = ws_nro_voucher + 1
wsdh = COX_LLAVE!cox_DH
Do Until COX_LLAVE.EOF
  If COX_LLAVE!cox_DH <> wsdh Then
     GoSub GRABA
     WS_SAL_CUENTA = 0
     wcta = COX_LLAVE!cox_codcta
     wdh = COX_LLAVE!cox_DH
     WS_SAL_CUENTA = WS_SAL_CUENTA + COX_LLAVE!cox_IMPORTE
     WCUENTA = 0
     wsdh = COX_LLAVE!cox_DH
  Else
   wcta = COX_LLAVE!cox_codcta
   wdh = COX_LLAVE!cox_DH
   WS_SAL_CUENTA = WS_SAL_CUENTA + COX_LLAVE!cox_IMPORTE
   WCUENTA = 1
  End If
  COX_LLAVE.MoveNext
Loop
If WCUENTA = 1 Then
    GoSub GRABA
End If
If ww_flag = "A" Then
 GoTo pasito
End If
Return





GRABA:
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
    cov_voucher!COV_IMPORTE = WS_SAL_CUENTA
    cov_voucher!COV_ESTADO = "0"
    cov_voucher!COV_CODUSU = LK_CODUSU
    cov_voucher!cov_flag_automatica = "1"
    cov_voucher.Update
    WS_NRO_MOV = WS_NRO_MOV + 1
Return



WEXCEL:
  Dim xlchart As Chart
  'Dim wranF, wran1, wran2, WPAS
  
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  FrmImpC1.lblProceso.Caption = "Abriendo , Archivo Ventas.xls . . . "
  DoEvents
  WPAS = "131296"
  xl.Workbooks.Open "C:\ADMIN\CONTABILIDAD\LIBRO_CAJA.xls", 0, True, 4, WPAS, WPAS
Return



'*** RUTINAS PARA IMPRIMIR



WPROGRESO:

Return

Exit Sub
CANCELA:
  MsgBox "Verificar Datos ,e Intente Nuevamente..", 48, Pub_Titulo
  FrmImpC1.pantalla.Enabled = True
  FrmImpC1.pantalla.Caption = "Por &Pantalla"
  FrmImpC1.lblProceso.Visible = False
  xl.Application.Visible = True
  Set xl = Nothing
  Screen.MousePointer = 0

Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImpC1

End Sub

Public Sub LIBRO_RESTO()
' *** REPORTES DE NUCLEOS
'On Error GoTo CANCELA
Dim xF As Integer
Dim PSCOXD_LLAVE As rdoQuery
Dim COXD_LLAVE  As rdoResultset
Dim ww_descri
Dim WS_NRO_MOV As Integer
Dim ws_nro_voucher As Integer
Dim WS_FECHA1 As Date
Dim WS_FECHA2 As Date
Dim WS_SAL_CUENTA As Currency
Dim WS_CUENTA As String * 12
Dim WS_TOT_IMPORTE_S As Currency
Dim WS_FLAG As String * 1
Dim WS_MAYOR As String
Dim XFF As Integer
Dim WS_SAL_CUENTA2 As Currency
Dim WS_SAL_DEB1 As Currency
Dim WS_SAL_DEB2 As Currency
Dim WS_SAL_HAB1 As Currency
Dim WS_SAL_HAB2 As Currency
Dim wdh As String * 1
Dim wfila_ult As Integer
Dim CTA_10101_D As Currency
Dim CTA_10101_H As Currency

SON_FECHAS txtCampo1, txtCampo2

 REP_FECHA1 = LK_FECHA_COP1
REP_FECHA2 = LK_FECHA_COP2

pub_cadena = "SELECT * FROM COMOX WHERE COX_CODCIA = ? AND COX_FECHA_VOUCHER >= ? AND COX_FECHA_VOUCHER <= ?  AND COX_IDENTIFICADOR = 'D'  ORDER BY  COX_NRO_VOUCHER  ASC"
Set PSCOXD_LLAVE = CN.CreateQuery("", pub_cadena)
PSCOXD_LLAVE(0) = 0
PSCOXD_LLAVE(1) = LK_FECHA_DIA
PSCOXD_LLAVE(2) = LK_FECHA_DIA
Set COXD_LLAVE = PSCOXD_LLAVE.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

PSCOXD_LLAVE(0) = LK_CODCIA
PSCOXD_LLAVE(1) = REP_FECHA1
PSCOXD_LLAVE(2) = REP_FECHA2
COXD_LLAVE.Requery


FrmImpC1.ProgBar.Min = 0
FrmImpC1.ProgBar.Max = COXD_LLAVE.RowCount
FrmImpC1.ProgBar.Value = 0
FrmImpC1.ProgBar.Visible = True
DoEvents
FrmImpC1.lblProceso.Caption = "Procesando. . . "
DoEvents

GoSub PASA_CONTAB
MsgBox "Proceso Terminado"
Exit Sub

PASA_CONTAB:

Dim wcta As String
Dim wflag1 As Integer

PSCOV_VOUCHER(0) = LK_CODCIA
PSCOV_VOUCHER(1) = LK_FECHA_COP1
PSCOV_VOUCHER(2) = LK_FECHA_COP2
cov_voucher.Requery
Do Until cov_voucher.EOF
    If cov_voucher!cov_flag_automatica = "9" Then cov_voucher.Delete
    cov_voucher.MoveNext
Loop

cov_voucher.Requery
If cov_voucher.EOF Then
 ws_nro_voucher = 0
Else
 cov_voucher.MoveLast
 ws_nro_voucher = cov_voucher!COV_NRO_VOUCHER
End If
WS_NRO_MOV = 0
wflag1 = 0
wcta = ""
wdh = COXD_LLAVE!cox_DH
WS_SAL_CUENTA = 0
Do Until COXD_LLAVE.EOF
     If Val(wcta) <> Val(COXD_LLAVE!COX_NRO_VOUCHER) Then
        SQ_OPER = 1
        PUB_CODTRA = Val(COXD_LLAVE!COX_NRO_VOUCHER)
        ws_nro_voucher = ws_nro_voucher + 1
        LEER_TRA_LLAVE
        If tra_llave.EOF Then
           ww_descri = ".........."
        Else
           ww_descri = tra_llave(1)
        End If
     End If
     
     wcta = COXD_LLAVE!COX_NRO_VOUCHER
     GoSub GRABA
     COXD_LLAVE.MoveNext
Loop

Return
GRABA:
    cov_voucher.AddNew
    cov_voucher!COV_NRO_MOV = WS_NRO_MOV
    cov_voucher!COV_CODCIA = LK_CODCIA
    cov_voucher!COV_NRO_VOUCHER = ws_nro_voucher
    cov_voucher!COV_FECHA_VOUCHER = LK_FECHA_COP2
    cov_voucher!COV_glosa = ww_descri
    cov_voucher!COV_FECHA_doc = LK_FECHA_DIA
    cov_voucher!COV_CODCTA = COXD_LLAVE!cox_codcta
    cov_voucher!COV_DH = COXD_LLAVE!cox_DH
    cov_voucher!COV_IMPORTE = COXD_LLAVE!cox_IMPORTE
    cov_voucher!COV_ESTADO = "0"
    cov_voucher!COV_CODUSU = COXD_LLAVE!COX_CODUSU
    cov_voucher!cov_flag_automatica = "9"
    cov_voucher.Update
    WS_NRO_MOV = WS_NRO_MOV + 1
Return


End Sub
Public Sub LIBRO_MAYOR()
'On Error GoTo FINTODO
Dim WSALDO As Currency
 Dim WS_SALDO_INICIAL  As Currency
 
Dim CT_RESULTADO As String
Dim WCAMBIA
Dim ws_clave As String
Dim WSFECHA As Date
Dim wcodven As Integer
Dim wvalor
Dim Wche As Integer
Dim wkSELECT As String
Dim wsfile As String
Dim F2 As Integer
Dim saldos As Currency
Dim SALDO_TOTAL As Currency
Dim Wflag As String * 1
Dim WCOL1 As Integer
Dim WCOL2 As Integer
Dim SALDO_COL1 As Currency
Dim SALDO_COL2 As Currency
Dim wsaldo_resultado As Currency
Dim WS_SALDO_FINAL As Currency
Dim CARAC As String
Dim saldo As Currency
Dim total As Currency
Dim wfi As Integer


Dim wscta1  As String
Dim wscta2 As String
Dim ws_tot_debe   As Currency
Dim ws_tot_haber As Currency
Dim f_final_d  As Integer
Dim f_final_h As Integer
Dim i As Integer
pub_cadena = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND COV_FECHA_VOUCHER >=? AND COV_FECHA_VOUCHER <=? AND (COV_CODCTA>= ? AND COV_CODCTA < ? )  ORDER BY COV_CODCIA, COV_NRO_VOUCHER, COV_DH" ', COV_NRO_MOV"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = 0
PS_REP01(2) = 0
PS_REP01(3) = 0
PS_REP01(4) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

FrmImpC1.lblProceso.Visible = True
FrmImpC1.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
GoSub WEXCEL

ws_clave = PUB_CLAVE

If Not SON_FECHAS(txtCampo1, txtCampo2) Then
  GoTo CANCELA
End If

F1 = 5  'Fila Inicial
C1 = 1
For i = 0 To listacta.ListCount - 1
  listacta.ListIndex = i
  If listacta.Selected(i) Then
    wscta1 = Val(Left(listacta.Text, 8))
    wscta2 = Val(Left(listacta.Text, 8)) + 1
    If WCOL1 > NIVEL_MAX Then
      WCOL1 = NIVEL_MAX
    End If
    GoSub OTRA_CTA
    If Wflag = "A" Then
      F1 = F1 + 2
    End If
  End If
Next i

  xl.Cells(1, 1) = Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption))
  xl.Cells(2, 1) = "L I B R O   M A Y O R   "
 ' xl.Cells(3, 1) = "' DEL " & Format(REP_FECHA1, "dd/mm/yyyy") & " AL " & Format(REP_FECHA2, "dd/mm/yyyy")
  xl.Cells(3, 1) = "'PERIODO : " & UCase(Format(LK_FECHA_COP1, "mmmm")) & " (" & Format(LK_FECHA_COP1, "dd/mm/yyyy") & " al " & Format(LK_FECHA_COP2, "dd/mm/yyyy") & ")"
  FrmImpC1.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.Application.Visible = True
  FrmImpC1.lblProceso.Visible = False
  FrmImpC1.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImpC1.pantalla.Enabled = True
  FrmImpC1.cerrar.Enabled = True
  FrmImpC1.pantalla.Caption = "Por &Pantalla"
  FrmImpC1.lblProceso.Visible = False
Exit Sub

OTRA_CTA:

    PS_REP01(0) = LK_CODCIA
    PS_REP01(1) = REP_FECHA1
    PS_REP01(2) = REP_FECHA2
    If opnivel(1).Value Then
      PS_REP01(3) = Format(wscta1, "00")
      PS_REP01(4) = Format(wscta2, "00")
    ElseIf opnivel(2).Value Then
      PS_REP01(3) = Format(wscta1, "000")
      PS_REP01(4) = Format(wscta2, "000")
    ElseIf opnivel(3).Value Then
      PS_REP01(3) = Format(wscta1, "00000")
      PS_REP01(4) = Format(wscta2, "00000")
    End If
    llave_rep01.Requery
 FrmImpC1.ProgBar.Min = 0
 Wflag = "A"
 SQ_OPER = 1
 PUB_CUENTA = PS_REP01(3) ' Format(wscta1, "00") 'Left(Trim(listacta.Text), 2)
 PUB_CODCIA = LK_CODCIA
 LEER_COM_LLAVE
 WSALDO = 0
 WS_SALDO_INICIAL = 0
 If com_llave.EOF Then
     MsgBox "Verificar Cuenta Contable : " & PUB_CUENTA, 48, Pub_Titulo
     Exit Sub
 End If
If periodos.Value = 1 Then
  JALA_SALDO com_llave!com_cuenta, 3
Else
  JALA_SALDO com_llave!com_cuenta, 0
  PUB_IMPORTE_DEB = 0
  PUB_IMPORTE_HAB = 0
End If
If (PUB_IMPORTE_DEB + PUB_IMPORTE_HAB) = 0 Then
    If llave_rep01.EOF Then
'      FrmImpC1.ProgBar.Max = llave_rep01.RowCount
      Wflag = ""
      GoTo sigue_cta
    End If
End If
    FrmImpC1.ProgBar.Value = 0
    DoEvents
    FrmImpC1.ProgBar.Visible = True
    DoEvents
    xcuenta = 0
    FrmImpC1.lblProceso.Caption = "Procesando . . .  un Momento ."
    DoEvents
    ws_tot_debe = 0
    ws_tot_haber = 0
    Dim f_final As Integer
    f_final = 0
    f_final_d = 0
    f_final_h = 0
    SALDO_TOTAL = 0
    saldos = 0
    
    xl.Cells(F1, 1) = UCase(Trim(listacta.Text))
    F1 = F1 + 1
    xl.Cells(F1, 1) = "FECHA"
    xl.Cells(F1, 2) = "VOUCHER"
    xl.Cells(F1, 3) = "IMPORTE"
    xl.Cells(F1, 5) = "FECHA"
    xl.Cells(F1, 6) = "VOUCHER"
    xl.Cells(F1, 7) = "IMPORTE"
    
 
 WSALDO = (Val(PUB_IMPORTE_DEB) * com_llave!com_signo_d) + (Val(PUB_IMPORTE_HAB) * com_llave!com_signo_h)
 WSALDO = Abs(WSALDO)
 WS_SALDO_INICIAL = WSALDO ' (Val(com_llave!COM_deb_ANO) * com_llave!com_SIGNO_D) + (Val(com_llave!COM_hab_ANO) * com_llave!com_SIGNO_H)
 If LK_EMP = "PIU" And REP_FECHA1 = "01/07/1999" Then
     WS_SALDO_INICIAL = 0
 End If
   F1 = F1 + 1
   If (Val(PUB_IMPORTE_DEB) * com_llave!com_signo_d) > (Val(PUB_IMPORTE_HAB) * com_llave!com_signo_h) Then
     xl.Cells(F1, 1) = "Saldo Inicial: "
     xl.Cells(F1, 3) = Format(WS_SALDO_INICIAL, "0.00")
     ws_tot_debe = ws_tot_debe + WS_SALDO_INICIAL
   Else
     xl.Cells(F1, 5) = "Saldo Inicial: "
     xl.Cells(F1, 7) = Format(WS_SALDO_INICIAL, "0.00")
     ws_tot_haber = ws_tot_haber + WS_SALDO_INICIAL
   End If

    F1 = F1 + 1
    fila = F1
    xcuenta = F1
    Do Until llave_rep01.EOF
      FrmImpC1.ProgBar.Value = FrmImpC1.ProgBar.Value + 1
      If Left(llave_rep01!COV_CODCTA, 2) = Left(listacta.Text, 2) Then
      Else
          If Trim(llave_rep01!COV_CODCTA) <> Left(listacta.Text, 2) Then
             GoTo OTRO
          End If
      End If
       If llave_rep01!COV_DH = "D" Then
         xl.Cells(fila, 1) = "'" & Format(llave_rep01!COV_FECHA_VOUCHER, "dd/mm/yyyy")
         xl.Cells(fila, 2) = llave_rep01!COV_NRO_VOUCHER
         xl.Cells(fila, 3) = Format(llave_rep01!COV_IMPORTE, "0.00")
         ws_tot_debe = ws_tot_debe + llave_rep01!COV_IMPORTE
         f_final_d = fila
       Else
         f_final_h = xcuenta
         xl.Cells(xcuenta, 5) = "'" & Format(llave_rep01!COV_FECHA_VOUCHER, "dd/mm/yyyy")
         xl.Cells(xcuenta, 6) = llave_rep01!COV_NRO_VOUCHER
         xl.Cells(xcuenta, 7) = Format(llave_rep01!COV_IMPORTE, "0.00")
         ws_tot_haber = ws_tot_haber + llave_rep01!COV_IMPORTE
       End If
       If llave_rep01!COV_DH = "H" Then
             xcuenta = xcuenta + 1
        End If
        If llave_rep01!COV_DH = "D" Then
              fila = fila + 1
        End If
        Wflag = "A"
OTRO:
  llave_rep01.MoveNext
Loop
   If f_final_h > f_final_d Then
     xcuenta = xcuenta + 1
   Else
     fila = fila + 1
   End If
   If fila > xcuenta Then
    Else
      fila = xcuenta
    End If
    F1 = fila
    WS_SALDO_FINAL = 0
   If ws_tot_debe > ws_tot_haber Then
     WS_SALDO_FINAL = Abs(ws_tot_debe - ws_tot_haber)
     xl.Cells(F1, 5) = "SALDO  "
     'xl.Cells(F1, 7) = Format(WSALDO, "##,##0.00")
     xl.Cells(F1, 7) = Format(WS_SALDO_FINAL, "##,##0.00")
     ws_tot_haber = ws_tot_haber + WS_SALDO_FINAL 'WSALDO
    Else
     WS_SALDO_FINAL = Abs(ws_tot_haber - ws_tot_debe)
     xl.Cells(F1, 1) = "SALDO  "
     xl.Cells(F1, 3) = Format(WS_SALDO_FINAL, "##,##0.00")
     ws_tot_debe = ws_tot_debe + WS_SALDO_FINAL 'WSALDO
    End If
    xl.Cells(F1 + 1, 1) = "TOTALES "
    xl.Cells(F1 + 1, 2) = ""
    xl.Cells(F1 + 1, 3) = Format(ws_tot_debe, "##,##0.00")
    xl.Cells(F1 + 1, 7) = Format(ws_tot_haber, "##,##0.00")
sigue_cta:
Return

WEXCEL:
  Dim wsfile1
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblProceso.Caption = "Abriendo , Archivo LIBRO_MAYOR.XLS . . . "
  DoEvents
  WPAS = ws_clave
  'If opnivel(0).Value = True Then
     xl.Workbooks.Open "C:\ADMIN\CONTABILIDAD\LIBRO_MAYOR.XLS", 0, True, 4, WPAS, WPAS
  'Else
  '   xl.Workbooks.Open "C:\ADMIN\CONTABILIDAD\AUXILIAR.xls", 0, True, 4, WPAS, WPAS
  'End If


Return
Exit Sub

CANCELA:
fin:
  FrmImpC1.pantalla.Enabled = True
  FrmImpC1.pantalla.Caption = "Por &Pantalla"
  FrmImpC1.lblProceso.Visible = False
  FrmImpC1.ProgBar.Visible = False
  pantalla.Enabled = True
  cerrar.Enabled = True
  Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
OJO:
If Err.Number = 70 Then
  MsgBox "Hoja de Calculo : " & wsfile1 & "  esta Abierta debe cerrar para Procesar Nuevamente ", 48, Pub_Titulo
  GoTo CANCELA
End If
Resume Next
Exit Sub
FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 xl.Application.Visible = True
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImpC1
 
End Sub


Public Sub LLENA_PERIODOS()

pub_cadena = "SELECT * FROM COHMAEST WHERE (COH_FECHA_PROCESO >= ? AND COH_FECHA_PROCESO2 >= ?) AND COH_CODCIA  = ? AND COH_CUENTA = ? ORDER BY COH_CODCIA,COH_CUENTA"
Set PSCOH_LLAVE = CN.CreateQuery("", pub_cadena)
PSCOH_LLAVE(0) = LK_FECHA_DIA
PSCOH_LLAVE(1) = LK_FECHA_DIA
PSCOH_LLAVE(2) = 0
PSCOH_LLAVE(3) = 0
Set coh_llave = PSCOH_LLAVE.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


pub_cadena = "SELECT DISTINCT COH_FECHA_PROCESO, COH_FECHA_PROCESO2 FROM COHMAEST WHERE COH_CODCIA  = '" & LK_CODCIA & "' ORDER BY COH_FECHA_PROCESO"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
llave_rep01.Requery
FECHA1.Clear
Do Until llave_rep01.EOF
  FECHA1.AddItem Format(llave_rep01!COH_FECHA_PROCESO, "dd/mm/yyyy") & " - " & Format(llave_rep01!COH_FECHA_PROCESO2, "dd/mm/yyyy")
  llave_rep01.MoveNext
Loop

End Sub

Public Sub POWER_REPORT(WT_ESTADO As Integer)
Dim WR_PAG As Integer
Dim WR_FECHA As String
Dim WR_CIA As String
Dim NIVELES As String
Dim PW_VALOR1 As String
Dim PW_VALOR2 As String
Dim PW_CUENTA As String
Dim PW_NIVELES As String
Dim PW_GRUPO As Integer
Dim wTEMP1 As Integer
Dim wTEMP2 As Integer
Dim wTEMP3 As Integer
Dim wTEMP4 As Integer
Dim wTEMP5 As Integer
Dim wSUMGRUPO1 As Currency
Dim wSUMTOTAL1 As Currency
Dim wSUMTOTAL2 As Currency
Dim WMONTO As Currency
Dim CTA_RESTA_SOLES As Currency
Dim CTA_RESTA As String

Dim wp_SUMGRUPO1 As String * 13
Dim wp_SUMTOTAL1 As String * 13
Dim wp_SUMTOTAL2 As String * 13
Dim wp_CUENTA As String * 5
Dim wp_DESCRIPCION As String * 25
Dim wp_MONTO As String * 13
Dim cad
Dim PC_CUENTA As rdoQuery
Dim ps_cta As rdoResultset
Dim BAN_GRUPO As String * 1
Dim WTABULA2 As String * 4
Dim WTABULA As String * 4
Dim una_ves
Dim un_nivel
Dim spacio As Integer
Dim spacio2 As Integer
Dim sp_grupo As Integer
Dim unpoco
Dim RUTA
Dim ww3
Dim CTA_SIGNO As Integer

Dim PC_89 As rdoQuery
Dim ps_cta_89 As rdoResultset

unpoco = 0
SQ_OPER = 2
PUB_TIPREG = WT_ESTADO
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
If tab_mayor.EOF Then
    MsgBox "Definir parametros del estado (TABLAS = 78)", 48, Pub_Titulo
    GoTo fin
End If
cad = "SELECT * FROM COMAEST WHERE COM_CUENTA >= ? and COM_NIVEL <> 2  and COM_CODCIA = ? ORDER BY COM_CUENTA"
Set PC_CUENTA = CN.CreateQuery("", cad)
PC_CUENTA(0) = 0
PC_CUENTA(1) = 0
Set ps_cta = PC_CUENTA.OpenResultset(rdOpenKeyset, rdConcurValues)
SQ_OPER = 2
PUB_TIPREG = WT_ESTADO
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
If tab_mayor.EOF Then
  MsgBox "Definir parametros del estado (TABLAS = 77)", 48, Pub_Titulo
  GoTo fin
End If

WTABULA2 = String(1, " ")
WTABULA = String(1, " ")
FrmImpC1.ProgBar.Min = 0
FrmImpC1.ProgBar.Value = 0
FrmImpC1.ProgBar.Visible = True
DoEvents
FrmImpC1.lblProceso.Visible = True
FrmImpC1.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
F1 = 5
C1 = 1
FrmImpC1.lblProceso.Caption = "Procesando . . . "
DoEvents
wSUMGRUPO1 = 0
BAN_GRUPO = "N"
una_ves = ""
sp_grupo = 0
WR_PAG = 0
FrmImpC1.ProgBar.Max = tab_mayor.RowCount
Do Until tab_mayor.EOF ' LOOP 1
  FrmImpC1.ProgBar.Value = FrmImpC1.ProgBar.Value + 1
  PW_VALOR1 = Trim(tab_mayor!tab_nomlargo)
  PW_VALOR2 = Trim(tab_mayor!tab_nomcorto)
  CTA_SIGNO = Val(tab_mayor!TAB_CODART)
  CTA_RESTA = Trim(tab_mayor!TAB_CONTABLE2)
  GoSub JALA_PW
  If PW_NIVELES <> "X" And BAN_GRUPO <> "T" Then
     SQ_OPER = 1
     PUB_CUENTA = PW_CUENTA
     LEER_COM_LLAVE
     If com_llave.EOF Then
        MsgBox "Verificar la cuenta NO Existe  : " & PW_CUENTA, 48, Pub_Titulo
        GoTo fin
     End If
  End If
  If PW_VALOR2 = "S" Then
    wp_SUMTOTAL2 = NUM_NEGA(wSUMTOTAL2)
    wp_SUMTOTAL2 = wp_SUMTOTAL2
    wp_DESCRIPCION = PW_CUENTA
    If wSUMTOTAL2 < 0 Then
      unpoco = 1
    Else
      unpoco = 0
    End If
    xl.Cells(F1, C1 + 1) = WTABULA2 & wp_DESCRIPCION
    xl.Cells(F1, C1 + 3) = wp_SUMTOTAL2
    xl.Cells(F1, C1 + 3).Borders.Item(xlEdgeTop).LineStyle = 1
    F1 = F1 + 1
    GoTo OTRO
  End If
  If PW_VALOR2 = "T" Then
    xl.Cells(F1, C1) = WTABULA2 & String(2, " ") & PW_CUENTA
    F1 = F1 + 1
    GoTo OTRO
  End If
  If BAN_GRUPO = "T" Then
    wp_SUMGRUPO1 = NUM_NEGA(wSUMGRUPO1)
    wp_SUMGRUPO1 = BAN_LINE(wp_SUMGRUPO1)
    wp_DESCRIPCION = PW_CUENTA
    If wSUMGRUPO1 < 0 Then
     unpoco = 1
    Else
     unpoco = 0
    End If
    xl.Cells(F1, C1 + 1) = WTABULA2 & wp_DESCRIPCION
    xl.Cells(F1, C1 + 2) = wp_SUMGRUPO1
    F1 = F1 + 1
    BAN_GRUPO = "N"
    wSUMGRUPO1 = 0
     una_ves = ""
    GoTo OTRO
  ElseIf BAN_GRUPO = "S" And una_ves = "" Then
     una_ves = "x"
     'If LKCHEK Then Print #1, WTABULA2; ""
  End If
  
  PC_CUENTA(0) = PW_CUENTA
  PC_CUENTA(1) = LK_CODCIA
  ps_cta.Requery
  wSUMTOTAL1 = 0
  un_nivel = 0
  Do Until ps_cta.EOF
    NIVELES = Val(ps_cta!COM_NIVEL)
    If NIVELES = "1" Then
     un_nivel = un_nivel + 1
    End If
    If NIVELES > PW_NIVELES Then
      GoTo OTRACTA
    End If
    If un_nivel = 2 Then
      Exit Do
    End If
    wp_CUENTA = ps_cta!com_cuenta
    CTA_RESTA_SOLES = 0
    If Trim(CTA_RESTA) <> "" Then
         SQ_OPER = 1
         PUB_CUENTA = CTA_RESTA
         LEER_COM_LLAVE
         If Not com_llave.EOF Then
           JALA_SALDO com_llave!com_cuenta, periodos.Value
           CTA_RESTA_SOLES = ((PUB_IMPORTE_DEB) * com_llave!com_signo_d) + ((PUB_IMPORTE_HAB) * com_llave!com_signo_h)
           If CTA_RESTA_SOLES <> 0 Then CTA_RESTA_SOLES = CTA_RESTA_SOLES * -1
         End If
     End If
     JALA_SALDO ps_cta!com_cuenta, periodos.Value
     If CTA_SIGNO = 0 Then
          WMONTO = ((PUB_IMPORTE_DEB) * ps_cta!com_signo_h) + ((PUB_IMPORTE_HAB) * ps_cta!com_signo_d) + CTA_RESTA_SOLES
     Else
          WMONTO = (((PUB_IMPORTE_HAB) * ps_cta!com_signo_h) + ((PUB_IMPORTE_DEB) * ps_cta!com_signo_d)) * CTA_SIGNO + CTA_RESTA_SOLES
     End If

    If NIVELES = 1 Then
       wSUMTOTAL1 = wSUMTOTAL1 + WMONTO
       wSUMTOTAL2 = wSUMTOTAL2 + WMONTO
    End If
    wp_DESCRIPCION = ps_cta!com_DESCRIPCION
    wp_MONTO = Format(WMONTO, "##,###,###.00")
    If NIVELES = 1 Then
      wp_MONTO = WMONTO
      wp_DESCRIPCION = ps_cta!com_DESCRIPCION
      spacio = 0
      If BAN_GRUPO = "S" Then
       spacio2 = 8 + 6
      Else
       spacio2 = 22
      End If
    ElseIf NIVELES = 2 Then
      If WMONTO < 0 Then
        wp_MONTO = Format(WMONTO * -1, "##,###,###.00")
      End If
      spacio = 5
      spacio2 = 8
    ElseIf NIVELES = 3 Then
      If WMONTO < 0 Then
        wp_MONTO = Format(WMONTO * -1, "##,###,###.00")
      End If
      If BAN_GRUPO = "S" Then
        spacio2 = 3 - 8
      Else
        spacio2 = -4 '-5
      End If
      spacio = 10
      
    End If
    sp_grupo = 5
    'If BAN_GRUPO = "S" Then
    '   sp_grupo = -3
    'End If
    wp_MONTO = wp_MONTO
    If WMONTO < 0 Then
      unpoco = 1
    Else
     unpoco = 0
    End If
    If WMONTO <> 0 Then
       xl.Cells(F1, C1) = WTABULA2 & wp_CUENTA
       xl.Cells(F1, C1 + 1) = wp_DESCRIPCION
       If Len(Trim(wp_CUENTA)) = 2 Then
         xl.Cells(F1, C1 + 3) = wp_MONTO
       Else
         xl.Cells(F1, C1 + 2) = wp_MONTO
       End If
       F1 = F1 + 1
    End If
OTRACTA:
  ps_cta.MoveNext
  Loop
  If PW_GRUPO <> 0 Then
    If BAN_GRUPO = "S" Then
       wSUMGRUPO1 = wSUMGRUPO1 + wSUMTOTAL1
    End If
  End If
OTRO:
tab_mayor.MoveNext
Loop ' LOOP 1
xl.Cells(1, 1) = Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption))
If WT_ESTADO = 77 Then
  xl.Cells(2, 2) = "ESTADO DE GANANCIAS Y PERDIDAS  POR FUNCION"
Else
  xl.Cells(2, 2) = "ESTADO DE GANANCIAS Y PERDIDAS  POR NATURALEZA"
End If
xl.Cells(3, 2) = "AL " & Format(LOC_FECHA_ULT, "dd/mm/yyyy")
xl.DisplayAlerts = False
xl.Worksheets(1).Protect PUB_CLAVE
xl.Application.Visible = True
FrmImpC1.lblProceso.Visible = False
FrmImpC1.ProgBar.Visible = False
Set xl = Nothing
FrmImpC1.ProgBar.Visible = False


Exit Sub
JALA_PW:
Dim chk As String
If PW_VALOR2 = "S" Or PW_VALOR2 = "T" Then
  PW_CUENTA = PW_VALOR1
  PW_NIVELES = "X"
  BAN_GRUPO = "N"
ElseIf PW_VALOR2 <> "" Then
  chk = Mid(PW_VALOR1, 3, 1)
  BAN_GRUPO = "S"
  If chk = "," Then
     PW_CUENTA = Left(PW_VALOR1, 2)
     PW_NIVELES = Trim(Mid(PW_VALOR1, 4, 2))
     PW_GRUPO = Val(PW_VALOR2)
  Else
    PW_CUENTA = PW_VALOR1
    PW_GRUPO = Val(PW_VALOR2)
    BAN_GRUPO = "T"
  End If
  
Else
  PW_CUENTA = Left(PW_VALOR1, 2)
  PW_NIVELES = Trim(Mid(PW_VALOR1, 4, 2))
  PW_GRUPO = Val(PW_VALOR2)
  BAN_GRUPO = "N"
End If

Return

WEXCEL:
  Dim wsfile1
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblProceso.Caption = "Abriendo , Archivo ESTADOS.xls . . . "
  DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open "C:\ADMIN\CONTABILIDAD\ESTADOS.xls", 0, True, 4, WPAS, WPAS
Return

fin:
  If xl Is Nothing Then
  Else
    Set xl = Nothing
  End If


End Sub


Public Sub CTA_HISTORICO()
Dim CADENITA, wformula, wformula1, wformula2, wformula3, wformula4
Dim Modo, Modo1
Dim Wche, wkSELECT
Dim wfecha, wfiltra1
Dim wcodcia As String
Dim wscta1  As String
Dim wformula0

lblProceso.Visible = True
pantalla.Enabled = False
cerrar.Enabled = False
Reportes.ReportFileName = "c:\ADMIN\CONTABILIDAD\" & "ctah.rpt"
Reportes.Connect = PUB_ODBC
Reportes.WindowTitle = "Reporte :  " & Trim(tra_llave(1))
Reportes.Destination = crptToWindow
Reportes.WindowLeft = 2
Reportes.WindowTop = 70
Reportes.WindowWidth = 635
Reportes.WindowHeight = 390
DoEvents
ProgBar.Min = 0
ProgBar.Max = 7
ProgBar.Value = 0
ProgBar.Visible = True
ProgBar.Value = ProgBar.Value + 1
wfecha = Format(LK_FECHA_DIA, "dd/mm/yyyy")
CADENITA = "{COHMAEST.COH_CUENTA} in ["
Modo1 = ""
For fila = 0 To listacta.ListCount - 1
  listacta.ListIndex = fila
  If listacta.Selected(fila) Then
    wscta1 = Val(Left(listacta.Text, 6))
    Modo1 = Modo1 + "'" + wscta1 + "' ,"
 End If
Next fila
If Modo1 <> "" Then
 CADENITA = CADENITA + Left(Modo1, Len(Modo1) - 1) & "] AND "
Else
 CADENITA = ""
End If
CADENITA = CADENITA + "{COHMAEST.COH_CODCIA} = '" & LK_CODCIA & "' AND {COHMAEST.COH_NIVEL} = 1 "
pub_cadena = CADENITA
wformula0 = "CIA=  '" & Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)) & "'"
Reportes.Formulas(0) = wformula0
Reportes.SelectionFormula = pub_cadena
Reportes.WindowTitle = Reportes.WindowTitle & " [ " & Trim(Reportes.ReportFileName) & "]"
Reportes.Action = 1
ProgBar.Value = ProgBar.Value + 1
ProgBar.Value = ProgBar.Value + 1
lblProceso.Visible = False
pantalla.Enabled = True
cerrar.Enabled = True
ProgBar.Visible = False
Exit Sub
procancela:
MsgBox Err.Description, 48, Pub_Titulo
Unload FrmImpC1
Exit Sub
Cancel:
ProgBar.Visible = False
lblProceso.Visible = False
pantalla.Enabled = True
cerrar.Enabled = True

End Sub
Public Sub CTA_CTE()
' *** REPORTES DE NUCLEOS
'On Error GoTo CANCELA
Dim ww_flag As String * 1
Dim xF As Integer
Dim PSCOX_LLAVE As rdoQuery
Dim COX_LLAVE  As rdoResultset
Dim WSGLOSA As String
Dim WS_NRO_MOV As Integer
Dim ws_nro_voucher As Integer
Dim WS_FECHA1 As Date
Dim WS_FECHA2 As Date
Dim WS_SAL_CUENTA As Currency
Dim WS_CUENTA As String * 12
Dim WS_TOT_IMPORTE_S As Currency
Dim WS_FLAG As String * 1
Dim WS_MAYOR As String
Dim WS_SAL_ANTERIOR As Currency
Dim XFF As Integer
Dim WS_SAL_CUENTA2 As Currency
Dim WS_SAL_DEB1 As Currency
Dim WS_SAL_DEB2 As Currency
Dim WS_SAL_HAB1 As Currency
Dim WS_SAL_HAB2 As Currency
Dim wdh As String * 1
Dim wfila_ult As Integer
Dim CTA_10101_D As Currency
Dim CTA_10101_H As Currency
Dim wscodcia  As String * 2
ww_flag = ""
WS_SAL_ANTERIOR = 0
'SON_FECHAS txtCampo1, txtCampo2
If Not SON_FECHAS(txtCampo1, txtCampo2) Then
  GoTo CANCELA
End If

If periodos.Value = 1 Then
  REP_FECHA1 = Left(Trim(FECHA1.Text), 10)
  REP_FECHA2 = Right(Trim(FECHA1.Text), 10)
Else
  If CDate(REP_FECHA1) <> LK_FECHA_COP1 Then che1.Value = 0
  If CDate(REP_FECHA2) <> LK_FECHA_COP2 Then che1.Value = 0
End If

pub_cadena = "SELECT * FROM MOVICONT WHERE MOV_CODCIA = ? AND MOV_CP = ? AND MOV_CODCLIE = ?  ORDER BY MOV_CODCIA ,MOV_FECHA_EMI"
Set PSCOX_LLAVE = CN.CreateQuery("", pub_cadena)
PSCOX_LLAVE(0) = 0
PSCOX_LLAVE(1) = 0
PSCOX_LLAVE(2) = 0
Set COX_LLAVE = PSCOX_LLAVE.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


wscodcia = LK_CODCIA
PSCOX_LLAVE(0) = wscodcia
PSCOX_LLAVE(1) = PUB_CP
PSCOX_LLAVE(2) = PUB_CODCLIE
COX_LLAVE.Requery
If COX_LLAVE.EOF Then
  Screen.MousePointer = 0
  MsgBox "NO Existen movimientos para caja ..", 48, Pub_Titulo
  Exit Sub
  ww_flag = "A"
End If
FrmImpC1.ProgBar.Min = 0
FrmImpC1.ProgBar.Max = COX_LLAVE.RowCount
FrmImpC1.ProgBar.Value = 0
FrmImpC1.ProgBar.Visible = True
DoEvents
FrmImpC1.lblProceso.Visible = True
FrmImpC1.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
C1 = 1

F1 = 7

xF = 4
XFF = 0
wdh = ""
WS_SAL_DEB1 = 0
WS_SAL_HAB1 = 0
CTA_10101_D = 0
CTA_10101_H = 0
FrmImpC1.lblProceso.Caption = "Procesando. . . "
DoEvents
xl.Cells(3, 1) = "Del " & Format(REP_FECHA1, "dd/mm/yyyy") & " al " & Format(REP_FECHA2, "dd/mm/yyyy")
xl.Cells(F1, 3) = "Saldo Inicial  :"
xl.Cells(F1, 5) = Format(WS_SAL_ANTERIOR, "0.00;(0.00)")
Do Until COX_LLAVE.EOF ' loop 1
     FrmImpC1.ProgBar.Value = FrmImpC1.ProgBar.Value + 1
     F1 = F1 + 1
     xl.Cells(F1, 1) = "'" & Format(COX_LLAVE!MOV_fecha_EMI, "DD/MM/YY")
     xl.Cells(F1, 2) = Trim(COX_LLAVE!MOV_NRO_VOUCHER)
     xl.Cells(F1, 3) = Trim(COX_LLAVE!MOV_SUNAT)
     xl.Cells(F1, 4) = Trim(COX_LLAVE!MOV_SERIE)
     xl.Cells(F1, 5) = Trim(COX_LLAVE!MOV_numfac)
     xl.Cells(F1, 6) = Trim(COX_LLAVE!MOV_DETALLE)
     'WS_TIPO_CAMBIO = 1
     'If COX_LLAVE!MOV_DH = "D" Then
     '  xl.Cells(F1, 7) = Trim(COX_LLAVE!MOV_IMPORTE)
     '  cSaldo = cSaldo + Val(GridK.TextMatrix(fila, 9))
     'Else
     '  xl.Cells(F1, 8) = Trim(COX_LLAVE!MOV_IMPORTE)
     '  cSaldo = cSaldo - Val(GridK.TextMatrix(fila, 10))
     'End If
     COX_LLAVE.MoveNext
Loop

xcuenta = 0
Screen.MousePointer = 0
FrmImpC1.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
DoEvents
xl.Application.Visible = True
DoEvents
FrmImpC1.lblProceso.Visible = False
FrmImpC1.ProgBar.Visible = False
Set xl = Nothing
pasito:
Screen.MousePointer = 0
FrmImpC1.pantalla.Enabled = True
FrmImpC1.pantalla.Caption = "Por &Pantalla"
FrmImpC1.lblProceso.Visible = False
   
Exit Sub
CANCELA:
  MsgBox "Verificar Datos ,e Intente Nuevamente..", 48, Pub_Titulo
  FrmImpC1.pantalla.Enabled = True
  FrmImpC1.pantalla.Caption = "Por &Pantalla"
  FrmImpC1.lblProceso.Visible = False
  xl.Application.Visible = True
  Set xl = Nothing
  Screen.MousePointer = 0

Exit Sub
WEXCEL:
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  xl.Workbooks.Open "C:\ADMIN\CONTABILIDAD\ESTCTE.xls", 0, True, 4, PUB_CLAVE, PUB_CLAVE
Return

FINTODO:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImpC1

End Sub

Private Sub txtnivel_KeyPress(KeyAscii As Integer)
 SOLO_ENTERO KeyAscii
End Sub
Public Sub A_CUENTAS()
'On Error GoTo FINTODO
Dim WS_NRO_MES As Integer
Dim WNROMES As Date
Dim WCUENTA As String
Dim WCAMBIA
Dim ws_clave As String
Dim WSFECHA As Date
Dim F2 As Integer
Dim QFECHA As String
Dim Qvoucher As String
Dim Qdetalle As String
Dim QDEBE As Currency
Dim QHABER As Currency
Dim QSALDO As Currency
Dim QDEBE_SUM As Currency
Dim QHABER_SUM As Currency
 'If Val(a_cta1.Text) > Val(a_cta2.Text) Then
Dim QMES_DEB As Currency
Dim QMES_HAB As Currency

 ' MsgBox "NO Procede...", 48, Pub_Titulo
'  Azul a_cta1, a_cta1
'  Exit Sub
'End If
GoTo dale
 SQ_OPER = 1
 PUB_CUENTA = Trim(a_cta1.Text)
 PUB_CODCIA = LK_CODCIA
 LEER_COM_LLAVE
 If com_llave.EOF Then
     MsgBox "Cuenta NO Existe ", 48, Pub_Titulo
     Azul a_cta1, a_cta1
     Exit Sub
 End If
If Val(com_llave!COM_NIVEL) <> Val(cop_llave!cop_nivel_max) Then
    MsgBox "Cuenta no es analitica", 48, Pub_Titulo
    Azul a_cta1, a_cta1
    Exit Sub
End If

 SQ_OPER = 1
 PUB_CUENTA = Trim(a_cta2.Text)
 PUB_CODCIA = LK_CODCIA
 LEER_COM_LLAVE
 If com_llave.EOF Then
     MsgBox "Cuenta NO Existe ", 48, Pub_Titulo
     Azul a_cta2, a_cta2
     Exit Sub
 End If
If Val(com_llave!COM_NIVEL) <> Val(cop_llave!cop_nivel_max) Then
    MsgBox "Cuenta no es analitica", 48, Pub_Titulo
    Azul a_cta2, a_cta2
    Exit Sub
End If

dale:

        
pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImpC1.lblProceso.Caption = "Activando Reporte... un Momento ."
DoEvents
If periodos.Value = 1 Then
 pub_cadena = "SELECT MOV_NRO_MES, MOV_FECHA,MOV_TIPMOV, MOV_FECHA_EMI, MOV_NRO_VOUCHER, MOV_CODCTA, MOV_DETALLE, MOV_DH, MOV_IMPORTE FROM MOVICONT WHERE MOV_CODCIA = ? AND (MOV_CODCTA >= ?  AND MOV_CODCTA <= ?) AND MOV_NRO_MES > 0 ORDER BY MOV_CODCTA, MOV_NRO_MES , MOV_FECHA_EMI, MOV_TIPMOV"
Else
 pub_cadena = "SELECT MOV_NRO_MES, MOV_FECHA,MOV_TIPMOV, MOV_FECHA_EMI, MOV_NRO_VOUCHER, MOV_CODCTA, MOV_DETALLE, MOV_DH, MOV_IMPORTE FROM MOVICONT WHERE MOV_CODCIA = ? AND MOV_NRO_MES = ?  AND (MOV_CODCTA >= ?  AND MOV_CODCTA <= ?) ORDER BY MOV_CODCTA, MOV_TIPMOV"
End If
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
If periodos.Value = 1 Then
    PS_REP01(1) = 0
    PS_REP01(2) = 0
Else
    PS_REP01(1) = 0
    PS_REP01(2) = 0
    PS_REP01(3) = 0
End If
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
DoEvents

'*** VERFICA SI HAY DATOS , O ESTAN CORRECTOS
PS_REP01(0) = LK_CODCIA
If periodos.Value = 1 Then
   PS_REP01(1) = Trim(a_cta1.Text)
   PS_REP01(2) = Trim(a_cta2.Text)
Else
    PS_REP01(1) = LK_NRO_MES
    PS_REP01(2) = Trim(a_cta1.Text)
    PS_REP01(3) = Trim(a_cta2.Text)
End If


llave_rep01.Requery
If llave_rep01.EOF = True Then
   MsgBox "!!! NO EXISTEN Datos ...", 48, Pub_Titulo
   GoTo CANCELA
End If
ws_clave = PUB_CLAVE
FrmImpC1.ProgBar.Min = 0
FrmImpC1.ProgBar.Max = llave_rep01.RowCount
FrmImpC1.ProgBar.Value = 0
FrmImpC1.lblProceso.Visible = True
FrmImpC1.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL

'xlLineStyleNone
'xl.Range("A4:L5").Borders.LineStyle = 3
'xl.Range("A4:L5").Borders.Item(xlEdgeTop).LineStyle = 3
'xl.Range("A4:L5").Borders.Item(xlEdgeTop).LineStyle = 3

FrmImpC1.ProgBar.Visible = True
DoEvents
xcuenta = 0
F1 = 5
C1 = 1
FrmImpC1.lblProceso.Caption = "Procesando . . .  un Momento ."
DoEvents
fila = 0

QDEBE_SUM = 0
QHABER_SUM = 0
QDEBE = 0
QHABER = 0
QSALDO = 0
WCUENTA = Trim(llave_rep01!MOV_CODCTA)
SQ_OPER = 1
PUB_CUENTA = WCUENTA
PUB_CODCIA = LK_CODCIA
LEER_COM_LLAVE
If com_llave.EOF Then
End If
If periodos.Value = 0 Then
 JALA_SALDO WCUENTA, 3
Else
 JALA_SALDO WCUENTA, 0, 0
End If
QSALDO = (PUB_IMPORTE_DEB * Val(com_llave!com_signo_d)) + (PUB_IMPORTE_HAB * Val(com_llave!com_signo_h))
QMES_DEB = 0
QMES_HAB = 0
F1 = F1 + 1
xl.Cells(F1, 3) = Trim(com_llave!com_cuenta) & " " & Trim(com_llave!com_DESCRIPCION)
F1 = F1 + 1
xl.Cells(F1, 3) = "SALDO ANTERIOR"
xl.Cells(F1, 6) = QSALDO
F1 = F1 + 1
xl.Cells(F1, 3) = UCase(Format(llave_rep01!MOV_FECHA, "mmmm"))
WS_NRO_MES = Val(llave_rep01!MOV_NRO_MES)
Do Until llave_rep01.EOF
    FrmImpC1.ProgBar.Value = FrmImpC1.ProgBar.Value + 1
    If WS_NRO_MES <> Val(llave_rep01!MOV_NRO_MES) Then
        F1 = F1 + 1
        xl.Cells(F1, 3) = "            Sumas del Mes  = S/."
        
        xl.Cells(F1, 1) = ""
        xl.Cells(F1, 2) = ""
        xl.Cells(F1, 4) = QMES_DEB
        xl.Cells(F1, 5) = QMES_HAB
        xl.Cells(F1, 6) = ""
        QMES_DEB = 0
        QMES_HAB = 0
        F1 = F1 + 1
        xl.Cells(F1, 3) = UCase(Format(llave_rep01!MOV_FECHA, "mmmm"))
        xl.Cells(F1, 1) = ""
        xl.Cells(F1, 2) = ""
        xl.Cells(F1, 4) = ""
        xl.Cells(F1, 5) = ""
        xl.Cells(F1, 6) = ""
        WS_NRO_MES = Val(llave_rep01!MOV_NRO_MES)
    End If
    If WCUENTA <> Trim(llave_rep01!MOV_CODCTA) Then
        F1 = F1 + 1
        xl.Cells(F1, 1) = ""
        xl.Cells(F1, 2) = ""
        xl.Cells(F1, 3) = "            Suma de Cuenta = S/."
        xl.Cells(F1, 4) = QDEBE_SUM
        xl.Cells(F1, 5) = QHABER_SUM
        xl.Cells(F1, 6) = ""
        WCUENTA = Trim(llave_rep01!MOV_CODCTA)
        QDEBE_SUM = 0
        QHABER_SUM = 0
        QDEBE_SUM = 0
            QHABER_SUM = 0
            QDEBE = 0
            QHABER = 0
            QSALDO = 0
            WCUENTA = Trim(llave_rep01!MOV_CODCTA)
            SQ_OPER = 1
            PUB_CUENTA = WCUENTA
            PUB_CODCIA = LK_CODCIA
            LEER_COM_LLAVE
            If com_llave.EOF Then
            End If
            JALA_SALDO WCUENTA, 3
            QSALDO = (PUB_IMPORTE_DEB * Val(com_llave!com_signo_d)) + (PUB_IMPORTE_HAB * Val(com_llave!com_signo_h))
            F1 = F1 + 1
            xl.Cells(F1, 3) = Trim(com_llave!com_cuenta) & " " & Trim(com_llave!com_DESCRIPCION)
            F1 = F1 + 1
            xl.Cells(F1, 3) = "SALDO ANTERIOR"
            xl.Cells(F1, 6) = QSALDO
            F1 = F1 + 1
            xl.Cells(F1, 3) = UCase(Format(llave_rep01!MOV_FECHA, "mmmm"))
    End If
    F1 = F1 + 1
    xl.Cells(F1, 1) = Format(llave_rep01!MOV_fecha_EMI, "dd/mm/yy")
    If Val(llave_rep01!MOV_TIPMOV) = 1 Then
      xl.Cells(F1, 2) = "R.C.-" & Format(llave_rep01!MOV_NRO_VOUCHER, "00000")
    ElseIf Val(llave_rep01!MOV_TIPMOV) = 2 Then
      xl.Cells(F1, 2) = "R.V.-" & Format(llave_rep01!MOV_NRO_VOUCHER, "00000")
    ElseIf Val(llave_rep01!MOV_TIPMOV) = 3 Then
      xl.Cells(F1, 2) = "C.B.-" & Format(llave_rep01!MOV_NRO_VOUCHER, "00000")
    Else
      xl.Cells(F1, 2) = "OTR.-" & Format(llave_rep01!MOV_NRO_VOUCHER, "00000")
    End If
    xl.Cells(F1, 3) = Trim(llave_rep01!MOV_DETALLE)
    QDEBE = 0
    QHABER = 0
    If Trim(llave_rep01!MOV_DH) = "D" Then
      QDEBE = Val(llave_rep01!MOV_IMPORTE)
    Else
      QHABER = Val(llave_rep01!MOV_IMPORTE)
    End If
    xl.Cells(F1, 4) = QDEBE
    xl.Cells(F1, 5) = QHABER
    QSALDO = QSALDO + (QDEBE * Val(com_llave!com_signo_d)) + (QHABER * Val(com_llave!com_signo_h))
    xl.Cells(F1, 6) = QSALDO
    QDEBE_SUM = QDEBE_SUM + QDEBE
    QHABER_SUM = QHABER_SUM + QHABER
    QMES_DEB = QMES_DEB + QDEBE
    QMES_HAB = QMES_HAB + QHABER
otrito:
   llave_rep01.MoveNext
Loop
F1 = F1 + 1
xl.Cells(F1, 1) = ""
xl.Cells(F1, 2) = ""
xl.Cells(F1, 3) = "            Suma de Cuenta = S/."
xl.Cells(F1, 4) = QDEBE_SUM
xl.Cells(F1, 5) = QHABER_SUM
xl.Cells(F1, 6) = ""


 xl.Cells(1, 1) = Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption))
 xl.Cells(2, 1) = "ANALISIS DE CUENTAS AL  " & Format(LOC_FECHA_ULT, "dd/mm/yyyy")
  FrmImpC1.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.Application.Visible = True
  FrmImpC1.lblProceso.Visible = False
  FrmImpC1.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImpC1.pantalla.Enabled = True
  FrmImpC1.cerrar.Enabled = True
  FrmImpC1.pantalla.Caption = "Por &Pantalla"
  FrmImpC1.lblProceso.Visible = False
Exit Sub

WEXCEL:
  Dim wsfile1
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblProceso.Caption = "Abriendo , Archivo BAL_COMPRO.xls . . . "
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open "C:\ADMIN\CONTABILIDAD\A_CUENTAS.xls", 0, True, 4, WPAS, WPAS

Return

Exit Sub
CANCELA:
  FrmImpC1.pantalla.Enabled = True
  FrmImpC1.pantalla.Caption = "Por &Pantalla"
  FrmImpC1.lblProceso.Visible = False
  FrmImpC1.ProgBar.Visible = False
  pantalla.Enabled = True
  cerrar.Enabled = True
  Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub


OJO:
If Err.Number = 70 Then
  MsgBox "Hoja de Calculo : " & wsfile1 & "  esta Abierta debe cerrar para Procesar Nuevamente ", 48, Pub_Titulo
  GoTo CANCELA
End If
Resume Next
Exit Sub
FINTODO:
fin:
 MsgBox " Posible Error .. Reintente Nuevamente ..", 48, Pub_Titulo
 xl.Application.Visible = True
 Set xl = Nothing
 Screen.MousePointer = 0
 Unload FrmImpC1
Exit Sub

End Sub


