VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmimport 
   Caption         =   "Importar Datos"
   ClientHeight    =   2715
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   5415
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Pasar Reg, de Venta a Contabilidad"
      Height          =   2415
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4935
      Begin ComctlLib.ProgressBar pb 
         Height          =   255
         Left            =   840
         TabIndex        =   4
         Top             =   1440
         Visible         =   0   'False
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   0
      End
      Begin VB.CommandButton cmdExp 
         Caption         =   "&Exportar"
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   1920
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtCampo2 
         Height          =   285
         Left            =   2520
         TabIndex        =   1
         Top             =   720
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
         Left            =   1080
         TabIndex        =   0
         Top             =   720
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
      Begin VB.Label lblprocesos 
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Rango de Fechas:"
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   480
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Frmimport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xl As Object
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset
Dim PS_REP02 As rdoQuery
Dim llave_rep02 As rdoResultset
Dim PS_REP03 As rdoQuery
Dim llave_rep03 As rdoResultset
Dim PS_REP04 As rdoQuery
Dim llave_rep04 As rdoResultset

Private Sub cmdExp_Click()
Dim wsfile1
Dim QFBG As String * 1
Dim QNUMSER As Integer
Dim QNUMFAC As Currency
Dim QFECHA_COMPRA  As Date
Dim QBRUTO    As Currency
Dim QIMPTO As Currency
Dim QMONEDA As String * 1
Dim QCODCLIE As Currency
Dim QANULADO As String * 1

pub_cadena = "SELECT FAR_ESTADO, FAR_CODCLIE, FAR_MONEDA, FAR_NUMSER, FAR_NUMFAC, FAR_FBG, FAR_IMPTO, FAR_FECHA_COMPRA , FAR_BRUTO FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA_COMPRA >= ?  AND FAR_FECHA_COMPRA <= ?  AND  FAR_TIPMOV = 10 AND (FAR_FBG = 'F' OR FAR_FBG = 'B')   ORDER BY FAR_FBG, FAR_NUMSER, FAR_NUMFAC, FAR_NUMSEC"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = LK_FECHA_DIA
PS_REP01(2) = LK_FECHA_DIA
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = CDate(txtCampo1.Text)
PS_REP01(2) = CDate(txtCampo2.Text)
llave_rep01.Requery
If llave_rep01.EOF Then
  MsgBox "No existe registro para exportar", 48, Pub_Titulo
  Exit Sub
End If
pub_mensaje = "Exportar datos ¿Desea Continuar... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If
GoSub WEXCEL
cmdExp.Enabled = False
pb.Visible = True
DoEvents
pb.Min = 0
pb.max = llave_rep01.RowCount
pb.Value = 0
QFBG = llave_rep01!far_fbg
QNUMSER = llave_rep01!far_numser
QNUMFAC = llave_rep01!far_numfac
QFECHA_COMPRA = llave_rep01!FAR_fecha_compra
QBRUTO = llave_rep01!FAR_BRUTO
QIMPTO = llave_rep01!far_impto
QMONEDA = llave_rep01!FAR_MONEDA
QCODCLIE = llave_rep01!far_codclie
QANULADO = llave_rep01!far_estado
fila = 0
Do Until llave_rep01.EOF
    pb.Value = pb.Value + 1
    If llave_rep01!far_fbg <> QFBG Then
        fila = fila + 1
        GoSub AGREGA_OTRO
        QFBG = llave_rep01!far_fbg
        QNUMSER = llave_rep01!far_numser
        QNUMFAC = llave_rep01!far_numfac
    End If
    If llave_rep01!far_numser <> QNUMSER Then
        fila = fila + 1
        GoSub AGREGA_OTRO
        QFBG = llave_rep01!far_fbg
        QNUMSER = llave_rep01!far_numser
        QNUMFAC = llave_rep01!far_numfac
    End If
    If llave_rep01!far_numfac <> QNUMFAC Then
        fila = fila + 1
        GoSub AGREGA_OTRO
        QFBG = llave_rep01!far_fbg
        QNUMSER = llave_rep01!far_numser
        QNUMFAC = llave_rep01!far_numfac
    End If
        QFECHA_COMPRA = llave_rep01!FAR_fecha_compra
        QBRUTO = llave_rep01!FAR_BRUTO
        QIMPTO = llave_rep01!far_impto
        QMONEDA = llave_rep01!FAR_MONEDA
        QCODCLIE = llave_rep01!far_codclie
        QANULADO = llave_rep01!far_estado
    
llave_rep01.MoveNext
Loop
GoSub AGREGA_OTRO
pb.Visible = False
DoEvents
SALE:
  cmdExp.Enabled = True
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.Application.Visible = True
  xl.Workbooks(1).Save
  xl.Application.Visible = True
  DoEvents
  lblprocesos.Visible = False
  pb.Visible = False
  Set xl = Nothing
Exit Sub

AGREGA_OTRO:

xl.Cells(fila, 1) = QFECHA_COMPRA
If QFBG = "F" Then
    xl.Cells(fila, 2) = 1
    SQ_OPER = 1
    pu_cp = "C"
    pu_codclie = QCODCLIE
    pu_codcia = LK_CODCIA
    LEER_CLI_LLAVE
    If Trim(cli_llave!cli_ruc_esposo) = "" Then
 '     MsgBox "Ingrese el RUC. del Cliente : " & Trim(cli_llave!cli_nombre) & Chr(13) & "Intente nuevamente ", 48, Pub_Titulo
 '     GoTo SALE
    End If
    xl.Cells(fila, 9) = Trim(cli_llave!cli_ruc_esposo)
    xl.Cells(fila, 10) = Trim(cli_llave!CLI_NOMBRE)
End If
If QFBG = "B" Then
    xl.Cells(fila, 2) = 3
    SQ_OPER = 1
    pu_cp = "C"
    pu_codclie = QCODCLIE
    pu_codcia = LK_CODCIA
    LEER_CLI_LLAVE
    'If Trim(cli_llave!cli_ruc_esposo) = "" Then
    '  MsgBox "Ingrese el RUC. del Cliente : " & Trim(cli_llave!cli_nombre) & Chr(13) & "Intente nuevamente ", 48, Pub_Titulo
    '  GoTo SALE
    'End If
    xl.Cells(fila, 9) = QCODCLIE
    xl.Cells(fila, 10) = Trim(cli_llave!CLI_NOMBRE)
End If
xl.Cells(fila, 3) = QNUMSER
xl.Cells(fila, 4) = QNUMFAC
xl.Cells(fila, 5) = QMONEDA
If QANULADO <> "N" Then
QBRUTO = 0
QIMPTO = 0
End If
xl.Cells(fila, 6) = QBRUTO
xl.Cells(fila, 7) = QIMPTO
xl.Cells(fila, 8) = QBRUTO + QIMPTO
xl.Cells(fila, 11) = QANULADO

Return

Exit Sub


WEXCEL:
  wsfile1 = "RV" & LK_CODCIA & Format(txtCampo1.Text, "mm") & ".XLS"
  wsfile = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\OFFICE\" & wsfile1
  On Error GoTo OJO
  Kill wsfile
  On Error GoTo 0
  If xl Is Nothing Then
    Set xl = CreateObject("Excel.Application")
  End If
  lblprocesos.Caption = "Configurando Hoja de Calculo... un Momento ."
  DoEvents
  xl.SheetsInNewWorkbook = 1
  xl.Workbooks.Add
  xl.Worksheets(1).Name = "Ventas"
  xl.Windows(1).Caption = wsfile1 & " Registo de Venta - " & Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption))
  xl.Windows(1).WindowState = xlMaximized
  xl.Workbooks(1).SaveAs wsfile
  xl.ActiveWindow.Zoom = 83
  lblprocesos.Caption = "Procesando ..."
Return

OJO:
If Err.Number = 70 Then
  MsgBox "Archivo generado esta abierto .. ", 48, Pub_Titulo
  Exit Sub
End If
Resume Next
End Sub

Private Sub Form_Load()
CenterMe Frmimport
 txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 txtCampo2.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
' lblcampo1.Caption = "Fecha de Inicial : "
' lblcampo1.Visible = True
 txtCampo1.Mask = "##/##/####"
 txtCampo1.Visible = True
' lblcampo2.Caption = "Fecha de Final: "
' lblcampo2.Visible = True

End Sub

Private Sub txtCampo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If Not IsDate(txtCampo1.Text) Then
   MsgBox "Fecha invalidad", 48, Pub_Titulo
   Exit Sub
 End If
  Azul2 txtCampo2, txtCampo2
End If
End Sub

Private Sub txtcampo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If Not IsDate(txtCampo2.Text) Then
   MsgBox "Fecha invalidad", 48, Pub_Titulo
   Exit Sub
 End If
  cmdExp.SetFocus
End If

End Sub
