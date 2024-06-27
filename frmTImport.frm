VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmTImport 
   Caption         =   "Transfer. Export. (Envio de Información)"
   ClientHeight    =   6180
   ClientLeft      =   555
   ClientTop       =   1845
   ClientWidth     =   4950
   Icon            =   "frmTImport.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fracabe 
      Caption         =   "Opcion :"
      ForeColor       =   &H000000C0&
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4695
      Begin VB.OptionButton openvio 
         Caption         =   "Transfer. Ajustes de Mercaderia"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   2775
      End
      Begin VB.OptionButton openvio 
         Caption         =   "Transfer Envio de Mercaderia"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   2535
      End
      Begin VB.OptionButton openvio 
         Caption         =   "Transfer. de Ventas"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   2415
      End
   End
   Begin VB.Data transfer 
      Caption         =   "Transfer"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   1920
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4200
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.CommandButton ENVIO 
      Caption         =   "Exportar Transfer."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   360
      Picture         =   "frmTImport.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4680
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
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
      Height          =   750
      Left            =   3120
      Picture         =   "frmTImport.frx":0884
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Frame frainfo 
      Caption         =   "Imformación :"
      ForeColor       =   &H000000C0&
      Height          =   3855
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   4695
      Begin MSMask.MaskEdBox txtCampo2 
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   1320
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   14737632
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtCampo1 
         Height          =   285
         Left            =   1320
         TabIndex        =   13
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         BackColor       =   14737632
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PromptChar      =   "_"
      End
      Begin ComctlLib.ProgressBar barra 
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   2160
         Visible         =   0   'False
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lblarchi 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label lblpro 
         Caption         =   "Procesando Información ..."
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1920
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.Label lblinf 
         Caption         =   "Fecha Final :"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lblinf 
         Caption         =   "Fecha Inicial :"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   10
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblinf 
         Caption         =   "Archivo :"
         ForeColor       =   &H00000080&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Label lbltransfer 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1350
      Width           =   4455
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
      Left            =   -120
      TabIndex        =   3
      Top             =   5760
      Width           =   5295
   End
End
Attribute VB_Name = "frmTImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xl As Object

Private Sub Command1_Click()
Unload frmTImport
End Sub


Private Sub flopy_Click()

End Sub

Private Sub ENVIO_Click()
Dim REP_FECHA1  As String
Dim REP_FECHA2  As String
GEN.Requery
If Trim(Nulo_Valors(GEN!GEN_CIAEXP)) = "" Then
 MsgBox "Codigo de Compañia de Recepción de informacion no validad...", 48, Pub_Titulo
 Exit Sub
End If

If Right(txtCampo1.Text, 2) = "__" Then
  REP_FECHA1 = Left(txtCampo1.Text, 8)
Else
  REP_FECHA1 = Trim(txtCampo1.Text)
End If
If Not IsDate(REP_FECHA1) Then
    MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
    Azul2 txtCampo1, txtCampo1
    Exit Sub
End If
If Right(txtcampo2.Text, 2) = "__" Then
  REP_FECHA2 = Left(txtcampo2.Text, 8)
Else
  REP_FECHA2 = Trim(txtcampo2.Text)
End If
If Not IsDate(REP_FECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 Azul2 txtcampo2, txtcampo2
 Exit Sub
End If
If CDate(REP_FECHA1) > CDate(REP_FECHA2) Then
 MsgBox "Fechas Invalidadas ..", 48, Pub_Titulo
 Azul2 txtCampo1, txtCampo1
 Exit Sub
End If

pub_mensaje = "Procesar el Transfer del : " & REP_FECHA1 & " al " & REP_FECHA2 & " Continuar...? " & Chr(13) & "CiaExp : " & Trim(Nulo_Valors(GEN!GEN_CIAEXP))
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If

GoSub Enable_controles

Dim wd_banco As String * 20
Dim wd_chenum As String * 15
Dim wd_codigo  As String * 15
Dim RR
Dim PSFAR As rdoQuery
Dim far_r  As rdoResultset
Dim CONTADOR As Integer
Dim ArchOrigen, ArchDestino As String
Dim estilo, respuesta
Dim PSfar_menorx As rdoQuery
Dim far_menorx As rdoResultset

Dim WLUGAR As String
Dim WZONA As String
Dim WSUBZONA As String

'On Error GoTo SALE



pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCIA= ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND (FAR_TIPMOV = ? or FAR_TIPMOV = ?)  AND FAR_ESTADO <> 'E' ORDER BY FAR_FECHA, FAR_NUMOPER"
Set PSfar_menorx = CN.CreateQuery("", pub_cadena)
PSfar_menorx(0) = 0
PSfar_menorx(1) = LK_FECHA_DIA
PSfar_menorx(2) = LK_FECHA_DIA
PSfar_menorx(3) = 0
PSfar_menorx(4) = 0
Set far_menorx = PSfar_menorx.OpenResultset(rdOpenKeyset, rdConcurValues)

PSfar_menorx.rdoParameters(0) = LK_CODCIA
PSfar_menorx.rdoParameters(1) = txtCampo1.Text
PSfar_menorx.rdoParameters(2) = txtCampo1.Text
If openvio(0).Value Then
  PSfar_menorx.rdoParameters(3) = 100
  PSfar_menorx.rdoParameters(4) = -9999
ElseIf openvio(1).Value Then
  PSfar_menorx.rdoParameters(3) = 10
  PSfar_menorx.rdoParameters(4) = -9999
ElseIf openvio(2).Value Then
 PSfar_menorx.rdoParameters(3) = 5
 PSfar_menorx.rdoParameters(4) = 6
End If
far_menorx.Requery

Screen.MousePointer = 11

DoEvents
transfer.DatabaseName = lblarchi.Caption '  Left(Trim(PUB_RUTA_OTRO), 1) & ":\ADMIN\Standar\TAjustes.mdb"
transfer.RecordSource = "SELECT * FROM FACART"
On Error GoTo NOHAY
transfer.Refresh
On Error GoTo 0
Do Until transfer.Recordset.EOF
   transfer.Recordset.Edit
   transfer.Recordset.Delete
   transfer.Recordset.MoveNext
Loop

SALTA:
barra.Min = 0
barra.Value = 0
If far_menorx.RowCount > 0 Then barra.max = far_menorx.RowCount
CONTADOR = 0
If Not far_menorx.EOF Then
   PUB_FECHA = far_menorx!FAR_fecha
End If

CN.Execute "Begin Transaction", rdExecDirect
      
      Do Until far_menorx.EOF
         barra.Value = barra.Value + 1
         wd_chenum = " "
         wd_codigo = " "
         wd_banco = " "
         transfer.Recordset.AddNew
         transfer.Recordset!FAR_CODCIA = GEN!GEN_CIAEXP ' far_menorx!FAR_CODCIA
         transfer.Recordset!far_fbg = far_menorx!far_fbg
         transfer.Recordset!far_numser = far_menorx!far_numser
         transfer.Recordset!far_numfac = far_menorx!far_numfac
         transfer.Recordset!FAR_NUMSEC = far_menorx!FAR_NUMSEC
         transfer.Recordset!FAR_fecha = Format(far_menorx!FAR_fecha, "dd/mm/yy")
         transfer.Recordset!far_codclie = 0
         transfer.Recordset!far_codart = 0
         transfer.Recordset!FAR_PRECIO = far_menorx!FAR_PRECIO
         transfer.Recordset!FAR_pordescto1 = far_menorx!FAR_pordescto1
         transfer.Recordset!far_IMPTO = far_menorx!far_IMPTO
         transfer.Recordset!FAR_SUBTOTAL = far_menorx!FAR_SUBTOTAL
         transfer.Recordset!FAR_CODVEN = far_menorx!FAR_CODVEN
         transfer.Recordset!far_cantidad = far_menorx!far_cantidad
         transfer.Recordset!far_signo_car = far_menorx!far_signo_car
         transfer.Recordset!FAR_MONEDA = far_menorx!FAR_MONEDA
         transfer.Recordset!far_serguia = far_menorx!far_serguia
         transfer.Recordset!far_numguia = far_menorx!far_numguia
         transfer.Recordset!FAR_DIAS = far_menorx!FAR_DIAS
         transfer.Recordset!far_descri = far_menorx!far_descri
         transfer.Recordset!far_PESO = far_menorx!far_PESO
         transfer.Recordset!far_signo_arm = far_menorx!far_signo_arm
         transfer.Recordset!FAR_GASTOS = far_menorx!FAR_GASTOS
         transfer.Recordset!FAR_NUMSER_C = far_menorx!FAR_NUMSER_C
         transfer.Recordset!FAR_NUMFAC_C = far_menorx!FAR_NUMFAC_C
         transfer.Recordset!FAR_cp = far_menorx!FAR_cp
         transfer.Recordset!far_cod_sunat = far_menorx!far_cod_sunat
         transfer.Recordset!FAR_TIPDOC = far_menorx!FAR_TIPDOC
         transfer.Recordset!FAR_BRUTO = far_menorx!FAR_BRUTO
         transfer.Recordset!FAR_CHENUM = wd_chenum
         transfer.Recordset!FAR_CODBAN = wd_codigo
         transfer.Recordset!FAR_BANCO = wd_banco
         transfer.Recordset!FAR_SECUENCIA = far_menorx!FAR_NUM_LOTE
         SQ_OPER = 1
         pu_codclie = far_menorx!far_codclie
         pu_cp = "C"
         pu_codcia = LK_CODCIA
         LEER_CLI_LLAVE
         PUB_KEY = far_menorx!far_codart
         LEER_ART_LLAVE
         transfer.Recordset!far_nomart = Trim(art_LLAVE!art_nombre)
         If cli_llave.EOF Then
            transfer.Recordset!FAR_NOMCLI = 0
            transfer.Recordset!FAR_RUC = " "
            transfer.Recordset!FAR_DNI = " "
         Else
            transfer.Recordset!FAR_NOMCLI = cli_llave!CLI_NOMBRE
            If Trim(cli_llave!cli_ruc_esposo) <> "" Then
               transfer.Recordset!FAR_RUC = Trim(cli_llave!cli_ruc_esposo)
            Else
               transfer.Recordset!FAR_RUC = " "
            End If
            If Trim(cli_llave!cli_RUC_ESPOSA) <> "" Then
               transfer.Recordset!FAR_DNI = Trim(cli_llave!cli_RUC_ESPOSA)
            Else
               transfer.Recordset!FAR_DNI = " "
            End If
         End If
         transfer.Recordset!far_ALTERNO = art_LLAVE!ART_ALTERNO
         
       If Not cli_llave.EOF Then
        SQ_OPER = 1
        PUB_CODCIA = "00"
        PUB_NUMTAB = cli_llave!CLI_LUGAR_TRAB
        PUB_TIPREG = 25
        LEER_TAB_LLAVE
        WLUGAR = ""
        If Not tab_llave.EOF Then
        WLUGAR = Trim(tab_llave!tab_NOMLARGO)
        End If
        PUB_NUMTAB = cli_llave!cli_TRAB_ZONA
        PUB_TIPREG = 20
        LEER_TAB_LLAVE
        WZONA = ""
        If Not tab_llave.EOF Then
        WZONA = Trim(tab_llave!tab_NOMLARGO)
        End If
        PUB_NUMTAB = cli_llave!cli_TRAB_SUBZONA
        PUB_TIPREG = 35
        LEER_TAB_LLAVE
        WSUBZONA = ""
        If Not tab_llave.EOF Then
        WSUBZONA = Trim(tab_llave!tab_NOMLARGO)
        End If
        If val(cli_llave!CLI_TRAB_NUM) <> 0 Then
           transfer.Recordset!FAR_DIREC = Trim(WLUGAR) + " " + Trim(cli_llave!CLI_TRAB_DIREC) + " # " + Trim(cli_llave!CLI_TRAB_NUM) & "  " & WZONA & "  " & WSUBZONA
        Else
           transfer.Recordset!FAR_DIREC = Trim(WLUGAR) + " " + Trim(cli_llave!CLI_TRAB_DIREC) & "  " & WZONA & "  " & WSUBZONA
        End If
        transfer.Recordset!FAR_DIREC = Left(Trim(transfer.Recordset!FAR_DIREC), 150)
        If Trim(transfer.Recordset!FAR_DIREC) = "" Then transfer.Recordset!FAR_DIREC = " "
       Else
        transfer.Recordset!FAR_DIREC = " "
       End If
        PUB_FECHA_VCTO = far_menorx!FAR_fecha
        transfer.Recordset.Update
        far_menorx.MoveNext
        CONTADOR = CONTADOR + 1
      
        DoEvents
  Loop
  
CN.Execute "Commit Transaction", rdExecDirect

transfer.RecordSource = "SELECT * FROM FECHAS"
transfer.Refresh
transfer.Recordset.Edit
transfer.Recordset!fecha_final = PUB_FECHA_VCTO
transfer.Recordset.Update
On Error GoTo EMES
'RR = Shell("C:\WSADMIN\WST.BAT", 1)
GoSub Disable_controles
Screen.MousePointer = 0
MsgBox "Proceso de Terminado Satisfactoriamente " & Chr(13) & "Total de Registros : " & CONTADOR & Chr(13) & "DEL :" & txtCampo1.Text & "  AL " & txtcampo2.Text & Chr(13) & "Ubicación de Archivo :  " & Trim(transfer.DatabaseName), 48, Pub_Titulo



Exit Sub
SALE:
Screen.MousePointer = 0
MsgBox Err.Description, 48, Pub_Titulo
CN.Execute "Rollback Transaction", rdExecDirect
 Screen.MousePointer = 0
 GoSub Disable_controles:
 
Exit Sub
EMES:
  MsgBox "Proceso Terminado .. Pero EL proceso de Empaquetar hacerlo Manual", vbCritical, Pub_Titulo
  Screen.MousePointer = 0
  GoSub Disable_controles:
Exit Sub
NOHAY:
   MsgBox "No se encontro el Archivo : " & Trim(transfer.DatabaseName), vbCritical, Pub_Titulo
   Screen.MousePointer = 0
   GoSub Disable_controles:
Exit Sub
Enable_controles:
 lblpro.Visible = True
 DoEvents
 barra.Visible = True
 DoEvents
 ENVIO.Enabled = False
 Command1.Enabled = False
 frainfo.Enabled = False
 fracabe.Enabled = False
Return
Disable_controles:
 lblpro.Visible = False
 DoEvents
 barra.Visible = False
 DoEvents
 ENVIO.Enabled = True
 Command1.Enabled = True
 frainfo.Enabled = True
 fracabe.Enabled = True
Return

End Sub

Private Sub Form_Load()
LKCHEK = False
openvio(0).Value = True
txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
txtCampo1.Mask = "##/##/####"
txtcampo2.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
txtcampo2.Mask = "##/##/####"


End Sub


Private Sub openvio_Click(Index As Integer)
If Index = 0 Then
  lbltransfer.Caption = "Transfer. Ingreso de Mercaderia"
  lblarchi.Caption = Left(Trim(PUB_RUTA_OTRO), 1) & ":\ADMIN\STANDAR\TMERCA.MDB"
ElseIf Index = 1 Then
  lbltransfer.Caption = "Transfer. Ventas de Mercaderia"
  lblarchi.Caption = Left(Trim(PUB_RUTA_OTRO), 1) & ":\ADMIN\STANDAR\TRANSFER.MDB"
ElseIf Index = 2 Then
  lbltransfer.Caption = "Transfer. Ajustes de Mercaderia"
  lblarchi.Caption = Left(Trim(PUB_RUTA_OTRO), 1) & ":\ADMIN\STANDAR\TAJUSTES.MDB"
End If
Azul2 txtCampo1, txtCampo1
End Sub

Private Sub openvio_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
  Azul2 txtCampo1, txtCampo1
End If
End Sub

Private Sub txtCampo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Azul2 txtcampo2, txtcampo2
End Sub

Private Sub txtCampo2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If ENVIO.Enabled Then ENVIO.SetFocus
End If
End Sub
