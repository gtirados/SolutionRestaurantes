VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form FrmProce 
   BackColor       =   &H00FAEFDA&
   Caption         =   "Libros Contables"
   ClientHeight    =   3525
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6525
   Icon            =   "FrmProce.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox proce 
      ForeColor       =   &H00800000&
      Height          =   315
      ItemData        =   "FrmProce.frx":0442
      Left            =   1725
      List            =   "FrmProce.frx":0458
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   720
      Width           =   3210
   End
   Begin VB.Frame frmdes 
      BackColor       =   &H00FAEFDA&
      Caption         =   "Ctas destinos"
      Height          =   1935
      Left            =   5160
      TabIndex        =   12
      Top             =   480
      Visible         =   0   'False
      Width           =   1335
      Begin RichTextLib.RichTextBox TEXTOVAR 
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393217
         BackColor       =   16445402
         Enabled         =   -1  'True
         MultiLine       =   0   'False
         Appearance      =   0
         TextRTF         =   $"FrmProce.frx":04E8
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
         Height          =   1575
         Left            =   120
         TabIndex        =   13
         ToolTipText     =   "Tipreg = 56"
         Top             =   240
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   2778
         _Version        =   393216
         Cols            =   1
         FixedCols       =   0
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FAEFDA&
      Height          =   1455
      Left            =   120
      TabIndex        =   10
      Top             =   480
      Width           =   4935
      Begin VB.CommandButton cmdrepo 
         Caption         =   "Mostrar Reporte"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   0
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label lpro 
         AutoSize        =   -1  'True
         BackColor       =   &H00FAEFDA&
         Caption         =   "Libros Contables:"
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
         Left            =   45
         TabIndex        =   11
         Top             =   240
         Width           =   1440
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Procesos "
      Height          =   3495
      Left            =   6720
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   3015
      Begin VB.CommandButton autom 
         Caption         =   "Procesar Destinos de Ctas."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   840
         Width           =   2775
      End
      Begin VB.CommandButton cmdproce 
         Caption         =   "Pase para el Diario General "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   2040
         Width           =   2775
      End
      Begin VB.CommandButton cmdcli 
         Caption         =   "Actualizar Saldos de Clientes/Provee. del Mes"
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
         Left            =   120
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "2º Paso."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "1º Paso."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   855
      End
   End
   Begin ComctlLib.ProgressBar pb 
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   2160
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
   End
   Begin VB.CommandButton cmdcerrar 
      Caption         =   "Ce&rrar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4815
      TabIndex        =   4
      Top             =   2820
      Width           =   1395
   End
   Begin VB.Line Line1 
      X1              =   6600
      X2              =   6600
      Y1              =   0
      Y2              =   4200
   End
   Begin VB.Label lpb 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Procesando..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Label lperiodo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   1440
      TabIndex        =   6
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label lpro 
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "FrmProce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xl As Object
Dim PS_REP01 As rdoQuery
Dim llave_rep01 As rdoResultset
Dim pr_llave As rdoResultset
Dim PSPR_LLAVE As rdoQuery
Dim wran1 As String
Dim wran2  As String
Dim wranF As String
Dim LOC_PROCESO As Integer

Dim temporal As String
Dim PSTEMP_LLAVE As rdoQuery
Dim temp_llave As rdoResultset
Dim ws_nro_voucher As Integer
Dim ws_fecha_voucher As Date




Private Sub autom_Click()
' PRUEBA DE DESTINOS
 PRO_DESTINOS
 
Exit Sub
Print ""
Dim imp_total As Currency
Dim imp_des1 As Currency
Dim imp_des2 As Currency
Dim imp_des3 As Currency
Dim imp_des4 As Currency
Dim imp_des5 As Currency


Dim TEMPO_VAR
Dim PSCOV_CUENTA2 As rdoQuery
Dim cox_cuenta  As rdoResultset
Dim WS_CUENTA As String
Dim ws_parcial As Currency
Dim ws_dh_cov As String * 1
Dim tt_importe As Currency

Dim wcadena As String
Dim wvalor  As String * 1
' CHEQUE DE MES CERRADO 0 ABIERTO
wcadena = ""
SQ_OPER = 2
PUB_TIPREG = 155
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
If tab_mayor.EOF Then
  MsgBox "Crear tab_tipreg = 155 para seguridas", 48, Pub_Titulo
Else
    Do Until tab_mayor.EOF
      If tab_mayor!TAB_NUMTAB = Val(Format(LK_FECHA_COP1, "yyyy")) Then
        wcadena = Trim(tab_mayor!tab_nomlargo)
      End If
      tab_mayor.MoveNext
    Loop
    wvalor = Mid(wcadena, LK_NRO_MES + 1, 1)
    If wvalor = "1" Then
       MsgBox "<<<< Mes CERRADO Operaciones >>>>", vbCritical, Pub_Titulo
       Exit Sub
    End If
End If

cade = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND COV_FECHA_VOUCHER >= ? AND COV_FECHA_VOUCHER <= ?  AND COV_NRO_MES = " & LK_NRO_MES & " ORDER BY COV_CODCTA, COV_FECHA_VOUCHER, COV_NRO_MOV"
Set PSCOV_CUENTA2 = CN.CreateQuery("", cade)
PSCOV_CUENTA2(0) = 0
PSCOV_CUENTA2(1) = LK_FECHA_DIA
PSCOV_CUENTA2(2) = LK_FECHA_DIA
Set cox_cuenta = PSCOV_CUENTA2.OpenResultset(rdOpenKeyset, rdConcurValues)

cade = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND COV_FECHA_VOUCHER >= ? AND COV_FECHA_VOUCHER <= ?  AND COV_NRO_MES = " & LK_NRO_MES & " ORDER BY COV_CODCTA, COV_FECHA_VOUCHER, COV_NRO_MOV"
Set PSCOV_CUENTA2 = CN.CreateQuery("", cade)
PSCOV_CUENTA2(0) = 0
PSCOV_CUENTA2(1) = LK_FECHA_DIA
PSCOV_CUENTA2(2) = LK_FECHA_DIA
Set cox_cuenta = PSCOV_CUENTA2.OpenResultset(rdOpenKeyset, rdConcurValues)

ffecha1 = Format(LK_FECHA_COP1, "yyyy/mm/dd")
ffecha2 = Format(LK_FECHA_COP2, "yyyy/mm/dd")
wpub_cadena = "DELETE COMOV  WHERE ( COV_FLAG_AUTOMATICA = 'M'  ) AND COV_CODCIA = '" & LK_CODCIA & "' AND COV_FECHA_VOUCHER >=  '" & ffecha1 & "'  AND COV_FECHA_VOUCHER <=  '" & ffecha2 & "'"
CN.Execute wpub_cadena, rdExecDirect

pub_mensaje = "Procesoar los Destinos ¿Desea Continuar... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
    Exit Sub
End If

lpb.Caption = "Eliminando Movimientos Automaticos"
lpb.Caption = ""


PSCOV_CUENTA2.rdoParameters(0) = LK_CODCIA
PSCOV_CUENTA2.rdoParameters(1) = LK_FECHA_COP1
PSCOV_CUENTA2.rdoParameters(2) = LK_FECHA_COP2
cox_cuenta.Requery


' QUITE EN LIMA
'ws_fecha_proc = #1/1/00# ' por lo de piura
'If LK_FECHA_COP1 = ws_fecha_proc Then
'   Exit Sub
'End If

PSTEMP_LLAVE(0) = LK_CODCIA
PSTEMP_LLAVE(1) = LK_FECHA_COP1
PSTEMP_LLAVE(2) = #1/1/2020#
temp_llave.Requery
If temp_llave.EOF Then
  ws_nro_voucher = 0
  ws_fecha_voucher = LK_FECHA_COP1
Else
  temp_llave.MoveLast
  ws_nro_voucher = temp_llave!COV_NRO_VOUCHER
  ws_fecha_voucher = temp_llave!COV_FECHA_VOUCHER
End If


'barra.Max = cox_cuenta.RowCount
CONTADOR = 0
lpb.Caption = "Creando Movimientos Automaticos"
pb.Max = cox_cuenta.RowCount
pb.Min = 0
pb.Value = 0
pb.Visible = True

Do Until cox_cuenta.EOF
      pb.Value = pb.Value + 1
      
      SQ_OPER = 1
      PUB_CUENTA = cox_cuenta!COV_CODCTA
      PUB_CODCIA = LK_CODCIA
      LEER_COM_LLAVE
      If com_llave.EOF Then
         MsgBox "Revisar el diario ...Cuenta : " & PUB_CUENTA
         GoTo fin
      End If
      ww_ff = 0
      If com_llave!COM_POR_AUTOM_D <> 0 Then ww_ff = 1
      If com_llave!COM_POR_AUTOM_D2 <> 0 Then ww_ff = 2
      If com_llave!COM_POR_AUTOM_D3 <> 0 Then ww_ff = 3
      If com_llave!COM_POR_AUTOM_D4 <> 0 Then ww_ff = 4
      If com_llave!COM_POR_AUTOM_D5 <> 0 Then ww_ff = 5
      If Left(cox_cuenta!COV_CODCTA, 1) = "6" And TEMPO_VAR = 1 Then
'      cox_cuenta.MovePrevious
'        MsgBox ".."
      End If
      If Left(cox_cuenta!COV_CODCTA, 1) = "6" Then
        TEMPO_VAR = 1
      End If
      WS_NRO_MOV = 0
      
      WS_SUMA = Val(Nulo_Valors(com_llave!com_cuenta_AUTOM_D)) + Val(Nulo_Valors(com_llave!COM_CUENTA_AUTOM_D2)) + Val(Nulo_Valors(com_llave!COM_CUENTA_AUTOM_D3)) + Val(Nulo_Valors(com_llave!COM_CUENTA_AUTOM_D4)) + Val(Nulo_Valors(com_llave!COM_CUENTA_AUTOM_D5))
          
      If WS_SUMA <> 0 Then
         WS_SUMA = 0
         ws_glosa = cox_cuenta!COV_glosa
         ws_fecha_proc = cox_cuenta!COV_FECHA_VOUCHER
         wc_importe = cox_cuenta!COV_IMPORTE
         tt_importe = wc_importe
         ws_codusu = cox_cuenta!COV_CODUSU
         
         If cox_cuenta!COV_DH = "H" And (Left(cox_cuenta!COV_CODCTA, 1) = "6" Or Left(cox_cuenta!COV_CODCTA, 1) = "9") Then
            ws_dh = "H"
         Else
            ws_dh = "D"
         End If
         
         ws_nro_voucher = ws_nro_voucher + 1
         WS_CUENTA = com_llave!com_cuenta_AUTOM_D
         ws_dh_cov = cox_cuenta!COV_DH
         ' monto para las 5 destinos
         imp_des1 = 0
         imp_des2 = 0
         imp_des3 = 0
         imp_des4 = 0
         imp_des5 = 0
         If Trim(com_llave!com_cuenta_AUTOM_D) <> "" Then
           ws_por = com_llave!COM_POR_AUTOM_D
           imp_des1 = Format(wc_importe * ws_por / 100, "0.00")
         End If
         If Trim(com_llave!COM_CUENTA_AUTOM_D2) <> "" Then
           ws_por = com_llave!COM_POR_AUTOM_D2
           imp_des2 = Format(wc_importe * ws_por / 100, "0.00")
         End If
         If Trim(com_llave!COM_CUENTA_AUTOM_D3) <> "" Then
           ws_por = com_llave!COM_POR_AUTOM_D3
           imp_des3 = Format(wc_importe * ws_por / 100, "0.00")
         End If
         If Trim(com_llave!COM_CUENTA_AUTOM_D4) <> "" Then
           ws_por = com_llave!COM_POR_AUTOM_D4
           imp_des4 = Format(wc_importe * ws_por / 100, "0.00")
         End If
         If Trim(com_llave!COM_CUENTA_AUTOM_D5) <> "" Then
           ws_por = com_llave!COM_POR_AUTOM_D5
           imp_des5 = Format(wc_importe * ws_por / 100, "0.00")
         End If
         imp_total = imp_des1 + imp_des2 + imp_des3 + imp_des4 + imp_des5
         If Val(wc_importe) <> Val(imp_total) Then
          imp_des1 = Val(imp_des1) + (Val(wc_importe) - Val(imp_total))
          
         End If
         
         If Trim(com_llave!com_cuenta_AUTOM_D) <> "" Then
            'ws_por = com_llave!COM_POR_AUTOM_D
            ww_fff = 1
            ws_parcial = imp_des1
            GoSub graba_autom
         End If
         WS_CUENTA = com_llave!COM_CUENTA_AUTOM_D2
         If Trim(com_llave!COM_CUENTA_AUTOM_D2) <> "" Then
            'ws_por = com_llave!COM_POR_AUTOM_D2
            ww_fff = 2
            ws_parcial = imp_des2
            GoSub graba_autom
         End If
         WS_CUENTA = com_llave!COM_CUENTA_AUTOM_D3
         If Trim(com_llave!COM_CUENTA_AUTOM_D3) <> "" Then
            'ws_por = com_llave!COM_POR_AUTOM_D3
            ww_fff = 3
            ws_parcial = imp_des3
            GoSub graba_autom
         End If
         WS_CUENTA = com_llave!COM_CUENTA_AUTOM_D4
         If Trim(com_llave!COM_CUENTA_AUTOM_D4) <> "" Then
            'ws_por = com_llave!COM_POR_AUTOM_D4
            ww_fff = 4
            ws_parcial = imp_des4
            GoSub graba_autom
         End If
         WS_CUENTA = com_llave!COM_CUENTA_AUTOM_D5
         If Trim(com_llave!COM_CUENTA_AUTOM_D5) <> "" Then
            'ws_por = com_llave!COM_POR_AUTOM_D5
            ww_fff = 5
            ws_parcial = imp_des5
            GoSub graba_autom
         End If
         If Trim(com_llave!com_cuenta_AUTO_H) <> "" Then
            WS_CUENTA = com_llave!com_cuenta_AUTO_H
            WS_SUMA = 0
            ws_por = 100
            ws_parcial = wc_importe
            If cox_cuenta!COV_DH = "H" And (Left(cox_cuenta!COV_CODCTA, 1) = "6" Or Left(cox_cuenta!COV_CODCTA, 1) = "9") Then
               ws_dh = "D" ' Vws_dh = cox_cuenta!COV_DH
            Else
               ws_dh = "H"
            End If
            GoSub graba_autom
         End If
      End If
   CONTADOR = CONTADOR + 1
   DoEvents
'   barra.Value = CONTADOR
   cox_cuenta.MoveNext
Loop
 pb.Visible = False
 lpb.Caption = ""
 cop_llave.Requery
 cop_llave.Edit
 cop_llave!cop_FLAG_DES = "A"
 cop_llave.Update
 ACT_MESES (0)
 MsgBox "PROCESO TERMINADO...", 48, Pub_Titulo
Unload FrmProce
Exit Sub

graba_autom:
            If ws_por <> 100 Then
               ' MsgBox "Verificar Destinos de : " & com_llave!com_cuenta
            End If
            If ww_ff <> 0 Then
               If ws_por = 0 Then Return
            End If
            'ws_parcial = Format(wc_importe * ws_por / 100, "0.00")
'            If ww_ff = ww_fff Then ws_parcial = wc_importe - WS_SUMA
            If ws_parcial = 0 Then Return
            cov_voucher.AddNew
            cov_voucher!COV_FECHA_VOUCHER = ws_fecha_voucher 'ws_fecha_proc
            WS_NRO_MOV = WS_NRO_MOV + 1
            cov_voucher!COV_NRO_MOV = WS_NRO_MOV 'Nulo_Valor0(cox_cuenta!COV_NRO_MOV)
            cov_voucher!COV_CODCTA = WS_CUENTA
            cov_voucher!COV_NUMTAB = WS_NRO_MOV
            cov_voucher!COV_DH = ws_dh

            cov_voucher!COV_NRO_VOUCHER = ws_nro_voucher
            cov_voucher!COV_glosa = "Dest. de la Cta.: " & Trim(com_llave!com_cuenta) & " " & ws_glosa
            
            If ws_parcial = 0 Then
               cov_voucher!cov_flag_automatica = "0"
            Else
               cov_voucher!cov_flag_automatica = "M"
            End If
            'If ws_parcial = 0 Then Stop
                            
            WS_SUMA = WS_SUMA + ws_parcial
            cov_voucher!COV_IMPORTE = ws_parcial
            
            cov_voucher!COV_FECHA_doc = LK_FECHA_COP1
            cov_voucher!COV_CODUSU = ws_codusu
            cov_voucher!COV_CODCIA = LK_CODCIA
            cov_voucher!COV_ESTADO = WS_ESTADO
            cov_voucher!cov_nro_mes = LK_NRO_MES
'            If Val(tt_importe) <> Val(cov_voucher!COV_IMPORTE) Then Stop
            cov_voucher.Update
   TEMPO_VAR = 0
 Return
Exit Sub
fin:
End Sub

Private Sub cmdcerrar_Click()
Unload FrmProce
End Sub

Private Sub cmdcli_Click()
Dim wc_importe As Currency
Dim pr_llave9  As rdoResultset
Dim PSPR_LLAVE9 As rdoQuery

pub_mensaje = "Confirmar.. Proceso de Actualización de Saldos de Clientes."
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If
pub_cadena = "SELECT MOV_CP, MOV_MONEDA, MOV_CODCLIE, MOV_DH, MOV_NUMFAC, MOV_IMPORTE, MOV_FECHA_EMI FROM MOVICONT WHERE MOV_CODCIA = ? AND MOV_TIPMOV = ? AND (MOV_FECHA >= ? AND MOV_FECHA <= ?) AND MOV_NRO_MES = " & LK_NRO_MES & "  AND MOV_CODCLIE <> 0 ORDER BY MOV_CODCIA, MOV_CODCTA , MOV_DH"
Set PSPR_LLAVE9 = CN.CreateQuery("", pub_cadena)
PSPR_LLAVE9(0) = 0
PSPR_LLAVE9(1) = 0
PSPR_LLAVE9(2) = LK_FECHA_DIA
PSPR_LLAVE9(3) = LK_FECHA_DIA
Set pr_llave9 = PSPR_LLAVE9.OpenResultset(rdOpenKeyset, rdConcurValues)


cmdproce.Enabled = False
proce.Enabled = False
cmdrepo.Enabled = False
cmdcli.Enabled = False
If LK_NRO_MES = 1 Then
 wpub_cadena = "UPDATE CLISAL SET CLS_DEB01 = 0, CLS_HAB01 = 0  WHERE  CLS_CODCIA = '" & LK_CODCIA & "'"
ElseIf LK_NRO_MES = 2 Then
 wpub_cadena = "UPDATE CLISAL SET CLS_DEB02 = 0, CLS_HAB02 = 0  WHERE  CLS_CODCIA = '" & LK_CODCIA & "'"
ElseIf LK_NRO_MES = 3 Then
 wpub_cadena = "UPDATE CLISAL SET CLS_DEB03 = 0, CLS_HAB03 = 0  WHERE  CLS_CODCIA = '" & LK_CODCIA & "'"
ElseIf LK_NRO_MES = 4 Then
 wpub_cadena = "UPDATE CLISAL SET CLS_DEB04 = 0, CLS_HAB04 = 0  WHERE  CLS_CODCIA = '" & LK_CODCIA & "'"
ElseIf LK_NRO_MES = 5 Then
 wpub_cadena = "UPDATE CLISAL SET CLS_DEB05 = 0, CLS_HAB05 = 0  WHERE  CLS_CODCIA = '" & LK_CODCIA & "'"
ElseIf LK_NRO_MES = 6 Then
 wpub_cadena = "UPDATE CLISAL SET CLS_DEB06 = 0, CLS_HAB06 = 0  WHERE  CLS_CODCIA = '" & LK_CODCIA & "'"
ElseIf LK_NRO_MES = 7 Then
 wpub_cadena = "UPDATE CLISAL SET CLS_DEB07 = 0, CLS_HAB07 = 0  WHERE  CLS_CODCIA = '" & LK_CODCIA & "'"
ElseIf LK_NRO_MES = 8 Then
 wpub_cadena = "UPDATE CLISAL SET CLS_DEB08 = 0, CLS_HAB08 = 0  WHERE  CLS_CODCIA = '" & LK_CODCIA & "'"
ElseIf LK_NRO_MES = 9 Then
 wpub_cadena = "UPDATE CLISAL SET CLS_DEB09 = 0, CLS_HAB09 = 0  WHERE  CLS_CODCIA = '" & LK_CODCIA & "'"
ElseIf LK_NRO_MES = 10 Then
 wpub_cadena = "UPDATE CLISAL SET CLS_DEB10 = 0, CLS_HAB10 = 0  WHERE  CLS_CODCIA = '" & LK_CODCIA & "'"
ElseIf LK_NRO_MES = 11 Then
 wpub_cadena = "UPDATE CLISAL SET CLS_DEB11 = 0, CLS_HAB11 = 0  WHERE  CLS_CODCIA = '" & LK_CODCIA & "'"
ElseIf LK_NRO_MES = 12 Then
 wpub_cadena = "UPDATE CLISAL SET CLS_DEB12 = 0, CLS_HAB12 = 0  WHERE  CLS_CODCIA = '" & LK_CODCIA & "'"
End If
If Trim(wpub_cadena) <> "" Then
 CN.Execute wpub_cadena, rdExecDirect
End If
LOC_PROCESO = 0
otrito:
LOC_PROCESO = LOC_PROCESO + 1
PSPR_LLAVE9(0) = LK_CODCIA
PSPR_LLAVE9(1) = LOC_PROCESO
PSPR_LLAVE9(2) = LK_FECHA_COP1
PSPR_LLAVE9(3) = LK_FECHA_COP2
pr_llave9.Requery
If pr_llave9.EOF Then
  GoTo ava
End If
pb.Visible = True
DoEvents
pb.Max = pr_llave9.RowCount
pb.Min = 0
pb.Value = 0
'wc_codcta = pr_llave9!MOV_CODCTA
Do Until pr_llave9.EOF
  DoEvents
  pb.Value = pb.Value + 1
  wc_importe = Val(pr_llave9!mov_importe)
  If pr_llave9!MOV_MONEDA = "D" Then
     SQ_OPER = 1
     PUB_CAL_INI = pr_llave9!MOV_fecha_EMI
     PUB_CAL_FIN = pr_llave9!MOV_fecha_EMI
     PUB_CODCIA = LK_CODCIA
     LEER_CAL_LLAVE
     If Nulo_Valor0(cal_llave!cal_tipo_cambio) = 0 Then
       MsgBox "Falta tipo de cambio" & pr_llave9!MOV_fecha_EMI, 48, Pub_Titulo
       'Stop
       'pr_llave9.Edit
       'pr_llave9!MOV_FECHA_EMI = Format(pr_llave9!MOV_FECHA_EMI, "dd/mm") & "/2000"
       'pr_llave9.Update
       Exit Sub
     End If
     wc_importe = Format((wc_importe * Val(cal_llave!cal_tipo_cambio)), "0.00")
     If Val(pr_llave9!MOV_numfac) = 13341 And LK_NRO_MES = 2 And LOC_PROCESO = 1 Then
       wc_importe = Val(pr_llave9!mov_importe)
       wc_importe = Format((wc_importe * 3.484), "0.00")
     End If
     If Val(pr_llave9!MOV_numfac) = 8328 And LK_NRO_MES = 2 And LOC_PROCESO = 1 Then
       wc_importe = Val(pr_llave9!mov_importe)
       wc_importe = Format((wc_importe * 3.491), "0.00")
     End If
     If Val(pr_llave9!MOV_numfac) = 3050 And LK_NRO_MES = 2 And LOC_PROCESO = 2 Then
       wc_importe = Val(pr_llave9!mov_importe)
       wc_importe = Format((wc_importe * 3.46), "0.00")
     End If
     If Val(pr_llave9!MOV_numfac) = 17 And LK_NRO_MES = 2 And LOC_PROCESO = 2 Then
       wc_importe = Val(pr_llave9!mov_importe)
       wc_importe = Format((wc_importe * 3.472), "0.00")
     End If
     If Val(pr_llave9!MOV_numfac) = 86 And LK_NRO_MES = 4 And LOC_PROCESO = 2 Then
       wc_importe = Val(pr_llave9!mov_importe)
       wc_importe = Format((wc_importe * 3.476), "0.00")
     End If
     If Val(pr_llave9!MOV_numfac) = 3229 And LK_NRO_MES = 6 And LOC_PROCESO = 2 Then
       wc_importe = Val(pr_llave9!mov_importe)
       wc_importe = Format((wc_importe * 3.486), "0.00")
     End If
     If Val(pr_llave9!MOV_numfac) = 18487 And LK_NRO_MES = 6 And LOC_PROCESO = 1 Then
       wc_importe = Val(pr_llave9!mov_importe)
       wc_importe = Format((wc_importe * 3.486), "0.00")
     End If
     If Val(pr_llave9!MOV_numfac) = 230 And LK_NRO_MES = 6 And LOC_PROCESO = 1 Then
       wc_importe = Val(pr_llave9!mov_importe)
       wc_importe = Format((wc_importe * 3.486), "0.00")
     End If
     If Val(pr_llave9!MOV_numfac) = 8 And LK_NRO_MES = 7 And LOC_PROCESO = 2 Then
       wc_importe = Val(pr_llave9!mov_importe)
       wc_importe = Format((wc_importe * 3.485), "0.00")
     End If
  End If
  If Val(Nulo_Valor0(pr_llave9!MOV_codclie)) <> 0 Then
    If Trim(pr_llave9!MOV_CP) <> "H" Then
      If pr_llave9!MOV_DH = "D" Then
        ACT_SALDO_CLIS pr_llave9!MOV_codclie, pr_llave9!MOV_CP, wc_importe, 0
      Else
        ACT_SALDO_CLIS pr_llave9!MOV_codclie, pr_llave9!MOV_CP, 0, wc_importe
      End If
    End If
  End If

  pr_llave9.MoveNext
Loop
DoEvents
ava:
If LOC_PROCESO <> 4 Then
  GoTo otrito
End If
pb.Visible = False
cmdproce.Enabled = True
proce.Enabled = True
cmdrepo.Enabled = True
cmdcli.Enabled = True

MsgBox "Actualización ha Terminado.", 48, Pub_Titulo


End Sub

Private Sub cmdproce_Click()
Dim kl_voucher  As Integer
Dim WS_NRO_MOV  As Integer
Dim wc_importe As Currency
Dim wc_sum_debe  As Currency
Dim wc_sum_haber  As Currency
Dim WSUM_GEN_D As Currency
Dim WSUM_GEN_H As Currency
Dim Wflag As String * 1
Dim wfecha1  As String
Dim wfecha2  As String
If Trim(proce.Text) = "" Then
'  MsgBox "Seleccionar uno de la Lista ", 48, Pub_Titulo
'  proce.SetFocus
'  SendKeys "%{up}"
'  Exit Sub
End If
Dim wcadena As String
Dim wvalor  As String * 1
' CHEQUE DE MES CERRADO 0 ABIERTO
wcadena = ""
SQ_OPER = 2
PUB_TIPREG = 155
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
If tab_mayor.EOF Then
  MsgBox "Crear tab_tipreg = 155 para seguridas", 48, Pub_Titulo
Else
    Do Until tab_mayor.EOF
      If tab_mayor!TAB_NUMTAB = Val(Format(LK_FECHA_COP1, "yyyy")) Then
        wcadena = Trim(tab_mayor!tab_nomlargo)
      End If
      tab_mayor.MoveNext
    Loop
    wvalor = Mid(wcadena, LK_NRO_MES + 1, 1)
    If wvalor = "1" Then
       MsgBox "<<<< Mes CERRADO Operaciones >>>>", vbCritical, Pub_Titulo
       Exit Sub
    End If
End If

pub_mensaje = "Confirmar el Proceso de Pase al Diario General ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
   Exit Sub
End If

cmdproce.Enabled = False
proce.Enabled = False
cmdrepo.Enabled = False
cmdcli.Enabled = False
DoEvents
lpb.Caption = "Procesando ..."
DoEvents
CUENTA_PRO = 0
wpub_cadena = ""
If LK_EMP = "PIU" Then
 wfecha1 = Format(LK_FECHA_COP1, "dd/mm/yyyy")
 wfecha2 = Format(LK_FECHA_COP2, "dd/mm/yyyy")
Else
 wfecha1 = Format(LK_FECHA_COP1, "yyyy/mm/dd")
 wfecha2 = Format(LK_FECHA_COP2, "yyyy/mm/dd")
End If
'wpub_cadena = "DELETE COMOV  WHERE COV_FLAG_AUTOMATICA = '" & Trim(LOC_PROCESO) & "'  AND COV_CODCIA = '" & LK_CODCIA & "' AND COV_FECHA_VOUCHER >=  '" & wfecha1 & "'  AND COV_FECHA_VOUCHER <=  '" & wfecha2 & "' and COV_NRO_MES = " & LK_NRO_MES

wpub_cadena = "DELETE COMOV  WHERE COV_CODCIA = '" & LK_CODCIA & "' AND COV_FECHA_VOUCHER >=  '" & wfecha1 & "'  AND COV_FECHA_VOUCHER <=  '" & wfecha2 & "' and COV_NRO_MES = " & LK_NRO_MES
CN.Execute wpub_cadena, rdExecDirect


PROCESO_SIGUE:
DoEvents
If LOC_PROCESO = 1 Then
 lpb.Caption = "Procesando ...Registro de Compras"
ElseIf LOC_PROCESO = 2 Then
 lpb.Caption = "Procesando ...Registro de Venta"
ElseIf LOC_PROCESO = 3 Then
 lpb.Caption = "Procesando ...Libro de Ingresos de Fondos"
ElseIf LOC_PROCESO = 4 Then
 lpb.Caption = "Procesando ...Libro de Egresos de Fondos"
ElseIf LOC_PROCESO = 5 Then
 lpb.Caption = "Procesando ...Libro de Planillas"
ElseIf LOC_PROCESO = 6 Then
 lpb.Caption = "Procesando ...Libro de Otros."
Else
 lpb.Caption = "Procesando ..."
End If
DoEvents

WSUM_GEN_H = 0
WSUM_GEN_D = 0
wc_sum_debe = 0
wc_sum_haber = 0
wc_importe = 0
CUENTA_PRO = CUENTA_PRO + 1
LOC_PROCESO = CUENTA_PRO ' Val(Left(proce.Text, 3))
lpb.Visible = True
pb.Visible = True
PSPR_LLAVE(0) = LK_CODCIA
PSPR_LLAVE(1) = LOC_PROCESO
PSPR_LLAVE(2) = LK_FECHA_COP1
PSPR_LLAVE(3) = LK_FECHA_COP2
pr_llave.Requery
If pr_llave.EOF Then
  lpb.Visible = False
  pb.Visible = False
  cop_llave.Requery
  cop_llave.Edit
  If LOC_PROCESO = 1 Then
     cop_llave!cop_FLAG_REGC = "A"
  ElseIf LOC_PROCESO = 2 Then
     cop_llave!cop_FLAG_REGV = "A"
  ElseIf LOC_PROCESO = 3 Then
     cop_llave!cop_FLAG_CAJA = "A"
  End If
  cop_llave!cop_FLAG_DES = " "
  cop_llave!COP_FLAG_MAYORIZACION = " "
  cop_llave.Update
  'MsgBox "No existen movimientos ", 48, Pub_Titulo
GoTo ava
'  Exit Sub
End If

Wflag = ""
PSCOV_VOUCHER(0) = LK_CODCIA
PSCOV_VOUCHER(1) = LK_FECHA_COP1
PSCOV_VOUCHER(2) = LK_FECHA_COP2
'Debug.Print PSCOV_VOUCHER.SQL
cov_voucher.Requery
If cov_voucher.EOF Then
 ws_nro_voucher = 0
Else
 cov_voucher.MoveLast
 ws_nro_voucher = cov_voucher!COV_NRO_VOUCHER
End If
ws_nro_voucher = ws_nro_voucher + 1

pb.Max = pr_llave.RowCount
pb.Min = 0
pb.Value = 0
wc_codcta = pr_llave!MOV_CODCTA
If LOC_PROCESO = 1 Then
WSGLOSA = "Por las Compras del Mes : " & Format(LK_FECHA_COP2, "mmmm")
ElseIf LOC_PROCESO = 2 Then
WSGLOSA = "Por las Ventas del Mes : " & Format(LK_FECHA_COP2, "mmmm")
ElseIf LOC_PROCESO = 3 Then
WSGLOSA = "Por los Ingreso y Egresos de Cajs del Mes : " & Format(LK_FECHA_COP2, "mmmm")
Else
WSGLOSA = "Otros mes : " & Format(LK_FECHA_COP2, "mmmm")
End If
WS_NRO_MOV = 0
kl_voucher = 0
Do Until pr_llave.EOF
  pb.Value = pb.Value + 1
  If Trim(wc_codcta) <> Trim(pr_llave!MOV_CODCTA) Then
'  If Trim(wc_codcta) = "40110" Then Stop
'   If wc_sum_debe <> wc_sum_haber Then Stop
    wc_importe = wc_sum_debe
    wdh = "D"
    GoSub GRABA
    wdh = "H"
    wc_importe = wc_sum_haber
    GoSub GRABA
    wc_sum_debe = 0
    wc_importe = 0
    wc_sum_haber = 0
    wc_codcta = Trim(pr_llave!MOV_CODCTA)
  End If
  If kl_voucher = 0 And Trim(pr_llave!MOV_FLAG_DES) = "A" Then
    If WSUM_GEN_H <> WSUM_GEN_D Then
      If WSUM_GEN_D > WSUM_GEN_H Then
       wdh = "H"
       wc_importe = WSUM_GEN_D - WSUM_GEN_H
      Else
       wdh = "D"
       wc_importe = WSUM_GEN_H - WSUM_GEN_D
      End If
      wcta = "389004"
      GoSub GRABA
      'If wdh = "H" Then
      '  wdh = "D"
      '  wcta = "389004"
      '  GoSub GRABA
      '  wdh = "H"
      '  wcta = "75910"
      '  GoSub GRABA
      'Else
      '  wdh = "H"
      '  wcta = "389004"
      '  GoSub GRABA
      '  wdh = "D"
      '  wcta = "94599"
      '  GoSub GRABA
     'End If
     WSUM_GEN_D = 0
     WSUM_GEN_H = 0
    End If
    ws_nro_voucher = ws_nro_voucher + 1
    kl_voucher = 1
  End If
  wc_importe = Val(pr_llave!mov_importe)
  If pr_llave!MOV_MONEDA = "D" And Trim(pr_llave!MOV_FLAG_TC) = "" Then
     SQ_OPER = 1
     PUB_CAL_INI = pr_llave!MOV_fecha_EMI
     PUB_CAL_FIN = pr_llave!MOV_fecha_EMI
     PUB_CODCIA = LK_CODCIA
     LEER_CAL_LLAVE
     If Nulo_Valor0(cal_llave!cal_tipo_cambio) = 0 Then
       MsgBox "Falta tipo de cambio"
       Exit Sub
     End If
     wc_importe = Format((wc_importe * Val(cal_llave!cal_tipo_cambio)), "0.00")
  ElseIf Trim(pr_llave!MOV_MONEDA) = "D" And Trim(pr_llave!MOV_FLAG_TC) = "A" Then
      wc_importe = Format((wc_importe * Val(pr_llave!MOV_TIPO_CAMBIO)), "0.00")
  End If
  If pr_llave!MOV_DH = "D" Then
    wc_sum_debe = wc_sum_debe + wc_importe
    WSUM_GEN_D = WSUM_GEN_D + wc_importe
  Else
    wc_sum_haber = wc_sum_haber + wc_importe
    WSUM_GEN_H = WSUM_GEN_H + wc_importe
  End If
  wdh = pr_llave!MOV_DH
  wcta = pr_llave!MOV_CODCTA
  DoEvents
  pr_llave.MoveNext
  DoEvents
Loop
' siempre hay data el ultimo registro
wc_importe = wc_sum_debe
wdh = "D"
GoSub GRABA
wdh = "H"
wc_importe = wc_sum_haber
GoSub GRABA
If WSUM_GEN_H <> WSUM_GEN_D Then
   If WSUM_GEN_D > WSUM_GEN_H Then
      wdh = "H"
      wc_importe = WSUM_GEN_D - WSUM_GEN_H
   Else
      wdh = "D"
      wc_importe = WSUM_GEN_H - WSUM_GEN_D
   End If
   wcta = "389004"
   GoSub GRABA
   'If wdh = "H" Then
   '  wdh = "D"
   '  wcta = "389004"
   '  GoSub GRABA
   '  wdh = "H"
   '  wcta = "75910"
   '  GoSub GRABA
   'Else
   '  wdh = "H"
   '  wcta = "389004"
   '  GoSub GRABA
   '  wdh = "D"
   '  wcta = "94599"
   '  GoSub GRABA
   'End If
   
End If

 
wc_sum_debe = 0
wc_importe = 0
wc_sum_haber = 0
ava:
If CUENTA_PRO <> 6 Then
  'proce.ListIndex = CUENTA_PRO - 1
  GoTo PROCESO_SIGUE:
End If

' reporte a Medida
'cop_llave.Requery
'cop_llave.Edit
'cop_llave!cop_FLAG_DES = " "
'cop_llave!COP_FLAG_MAYORIZACION = " "
'Select Case Val(LOC_PROCESO)
'Case 1
'   REGISTROS 1
'   cop_llave!cop_FLAG_REGC = "A"
'Case 2
'   REGISTROS 2
 '  cop_llave!cop_FLAG_REGV = "A"
'Case 3
'   CAJABANCOS 3
'   cop_llave!cop_FLAG_CAJA = "A"
'Case 4

'End Select
'cop_llave.Update
pb.Visible = False
lpb.Visible = False
cmdproce.Enabled = True

proce.Enabled = True
cmdrepo.Enabled = True
cmdcli.Enabled = True
ACT_MESES (0)
MsgBox "Proceso Terminado", 48, Pub_Titulo
Unload FrmProce
Exit Sub
GRABA:
   If wc_importe = 0 Then GoTo PASACOV
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
    cov_voucher!COV_IMPORTE = wc_importe
    cov_voucher!COV_ESTADO = "0"
    cov_voucher!COV_CODUSU = LK_CODUSU
    cov_voucher!cov_flag_automatica = Trim(LOC_PROCESO)
    cov_voucher!cov_nro_mes = LK_NRO_MES
    If wc_importe <> cov_voucher!COV_IMPORTE Then Stop
    cov_voucher.Update
    
    WS_NRO_MOV = WS_NRO_MOV + 1
PASACOV:
Return





End Sub


Private Sub cmdrepo_Click()
LOC_PROCESO = Val(Left(proce.Text, 3))
If LOC_PROCESO = 0 Then
  MsgBox "Seleccionar Libro Contable.", 48, Pub_Titulo
  proce.SetFocus
  GoTo salea
End If

pb.Visible = True
cop_llave.Requery
cop_llave.Edit
cop_llave!cop_FLAG_DES = " "
cop_llave!COP_FLAG_MAYORIZACION = " "
Select Case Val(LOC_PROCESO)
Case 1
   REGISTROS 1
   cop_llave!cop_FLAG_REGC = "A"
Case 2
   REGISTROS 2
   cop_llave!cop_FLAG_REGV = "A"
Case 3
   CAJABANCOS 3
   cop_llave!cop_FLAG_CAJA = "A"
Case 4

End Select
cop_llave.Update
salea:
pb.Visible = False
lpb.Visible = False
Unload FrmProce

End Sub

Private Sub Form_Load()
'pub_cadena = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND (COV_FECHA_VOUCHER >= ? AND COV_FECHA_VOUCHER <= ?)  AND COV_NRO_MES = " & LK_NRO_MES & " ORDER BY COV_NRO_VOUCHER, COV_NRO_MOV "
'Set PSTEMP_LLAVE = CN.CreateQuery("", pub_cadena)
'PSTEMP_LLAVE(0) = 0
'PSTEMP_LLAVE(1) = LK_FECHA_DIA
'PSTEMP_LLAVE(2) = LK_FECHA_DIA
'Set temp_llave = PSTEMP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

'pub_cadena = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND (COV_FECHA_VOUCHER >= ? AND COV_FECHA_VOUCHER <= ?)  AND COV_NRO_MES = " & LK_NRO_MES & " ORDER BY COV_NRO_VOUCHER, COV_NRO_MOV "

pub_cadena = "SELECT *  FROM MOVICONT WHERE MOV_CODCIA = ? AND (MOV_FECHA >= ? AND MOV_FECHA <=?) AND MOV_NRO_MES = ?  AND MOV_TIPMOV = ?  ORDER BY MOV_NRO_VOUCHER DESC "
Set PSTEMP_LLAVE = CN.CreateQuery("", pub_cadena)
PSTEMP_LLAVE(0) = 0
PSTEMP_LLAVE(1) = LK_FECHA_DIA
PSTEMP_LLAVE(2) = LK_FECHA_DIA
PSTEMP_LLAVE(3) = 0
PSTEMP_LLAVE(4) = 0
Set temp_llave = PSTEMP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

LOC_PROCESO = 0
CenterMe FrmProce
'COX_DH DESC , COX_CODCTA ASC
pub_cadena = "SELECT MOV_TIPO_CAMBIO, MOV_FLAG_TC, MOV_FLAG_DES,MOV_CP, MOV_MONEDA, MOV_CODCLIE, MOV_DH, MOV_NUMFAC, MOV_IMPORTE, MOV_FECHA_EMI, MOV_CODCTA  FROM MOVICONT WHERE MOV_CODCIA = ? AND MOV_TIPMOV = ? AND (MOV_FECHA >= ? AND MOV_FECHA <= ?) AND MOV_NRO_MES = " & LK_NRO_MES & "  ORDER BY MOV_FLAG_DES, MOV_CODCTA , MOV_DH"
Set PSPR_LLAVE = CN.CreateQuery("", pub_cadena)
PSPR_LLAVE(0) = 0
PSPR_LLAVE(1) = 0
PSPR_LLAVE(2) = LK_FECHA_DIA
PSPR_LLAVE(3) = LK_FECHA_DIA
Set pr_llave = PSPR_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)
lperiodo.Caption = "Del " & Format(LK_FECHA_COP1, "dd/mm/yyyy") & " al  " & Format(LK_FECHA_COP2, "dd/mm/yyyy")

gridigv.TextMatrix(0, 0) = "Cta."
gridigv.ColWidth(0) = 450
SQ_OPER = 2
PUB_TIPREG = 56
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
fila = 1
gridigv.Rows = 1
Do Until tab_mayor.EOF
  gridigv.Rows = gridigv.Rows + 1
   gridigv.TextMatrix(fila, 0) = Trim(tab_mayor!tab_nomlargo)
   fila = fila + 1
   tab_mayor.MoveNext
Loop

proce.Clear
PUB_TIPREG = 150
PUB_CODCIA = "00"
SQ_OPER = 2
LEER_TAB_LLAVE
Do Until tab_mayor.EOF
    proce.AddItem "0" & tab_mayor!TAB_NUMTAB & ".-" & Trim(tab_mayor!tab_nomlargo) & String(80, " ") & Trim(tab_mayor!TAB_CONTABLE2)
    proce.ItemData(proce.NewIndex) = tab_mayor!TAB_NUMTAB
    tab_mayor.MoveNext
Loop


End Sub

Private Sub MSFlexGrid1_Click()

End Sub

Private Sub proce_Click()
If Val(Left(proce.Text, 2)) = 3 Then
  frmdes.Visible = True
Else
  frmdes.Visible = False
End If

End Sub

Private Sub proce_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  cmdrepo.SetFocus
End If
End Sub

Public Sub REGISTROS(WTIPMOV)
Dim wc_detalle As String
Dim wc_desTc As String * 20
Dim wflag_vocucher As String * 1
Dim wc_voucher  As Integer
Dim wc_fecha As String
Dim wc_cuenta As String
Dim wc_sunat As String
Dim wc_serie As String
Dim wc_numfac As String
Dim wc_bruto As Currency
Dim wc_inaf As Currency
Dim wc_afect As Currency
Dim wc_total As Currency
Dim wc_fila_I As Integer
Dim wc_fila_F As Integer
Dim wc_cp As String * 1
Dim wc_codclie As Currency
Dim wc_importe As Currency
Dim WS_VOU2 As Currency
Dim vbRespContinua As Integer
    cmdproce.Enabled = False
    proce.Enabled = False
    cmdrepo.Enabled = False
    DoEvents
    lpb.Caption = "Procesando Registros..."
    DoEvents
      
 pub_cadena = "SELECT * FROM MOVICONT WHERE MOV_CODCIA = ? AND MOV_TIPMOV = ? AND (MOV_FECHA >= ? AND MOV_FECHA <= ?) AND MOV_NRO_MES = " & LK_NRO_MES & " AND MOV_MARCA='X' ORDER BY MOV_NRO_VOUCHER"
 Set PS_REP01 = CN.CreateQuery("", pub_cadena)
 PS_REP01(0) = 0
 PS_REP01(1) = 0
 PS_REP01(2) = LK_FECHA_DIA
 PS_REP01(3) = LK_FECHA_DIA
 Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
 PS_REP01(0) = LK_CODCIA
 PS_REP01(1) = WTIPMOV
 PS_REP01(2) = LK_FECHA_COP1
 PS_REP01(3) = LK_FECHA_COP2
 llave_rep01.Requery
 If llave_rep01.EOF Then
    cmdproce.Enabled = True
    proce.Enabled = True
    cmdrepo.Enabled = True
    DoEvents
    lpb.Visible = False
    DoEvents
   MsgBox "No existen Movimientos.", 48, Pub_Titulo
   Exit Sub
 End If
 wflag_vocucher = ""
 pub_mensaje = "Mostrar el Registro en Detalle ...?"
 Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
 If Pub_Respuesta = vbYes Then
     wflag_vocucher = "A"
 End If
 GoSub WEXCEL
 If WTIPMOV = 1 Then
    xl.Cells(4, 7) = "PROVEEDOR"
    xl.Cells(5, 8) = "COMPRAS"
    xl.Cells(5, 10) = "COMPRAS"
    xl.Cells(5, 11) = "COMPRA"
 Else
    xl.Cells(4, 7) = "CLIENTE"
    xl.Cells(4, 8) = "VALOR"
    xl.Cells(4, 9) = "EXON"
    xl.Cells(5, 11) = "VENTA"
 End If
 If wflag_vocucher = "A" Then
'    xl.Cells(4, 8) = "Nro."
'    xl.Cells(4, 9) = ""
'    xl.Cells(5, 8) = "Voucher"
'    xl.Cells(5, 9) = "Detalle"
 End If
 F1 = 5
 wc_total = 0
 WC_HABER = 0
 WC_DEBE = 0
 wc_bruto = 0
 wc_inaf = 0
 wc_afect = 0
 wc_fila_I = F1 + 1
 wc_voucher = 0 'llave_rep01!MOV_NRO_VOUCHER
 pb.Max = llave_rep01.RowCount
 pb.Min = 0
 pb.Value = 0
 wc_importe = 0
 Do Until llave_rep01.EOF
    If Val(llave_rep01!MOV_NRO_VOUCHER) = 43 Then MsgBox "HHH"
   pb.Value = pb.Value + 1
'   If llave_rep01!MOV_NRO_VOUCHER = 73 Then Stop
   If Val(wc_voucher) <> Val(llave_rep01!MOV_NRO_VOUCHER) Then
     If WC_HABER = 0 And llave_rep01!MOV_TIPMOV = 1 Then GoTo DALE1
      F1 = F1 + 1
      xl.Cells(F1, 1) = wc_fecha
      
      If wc_codclie = 0 Then GoTo SALE
      SQ_OPER = 10
      pu_codcia = LK_CODCIA
      pu_cp = wc_cp
      pu_codclie = wc_codclie
      LEER_CLI_LLAVE
      If cli_llave10.EOF Then
         vbRespContinua = MsgBox("Verificar Codigo de Cliente " & wc_codclie & vbCrLf & "Desea Continuar....???", vbYesNo + vbQuestion, Pub_Titulo)
         If vbRespContinua = vbYes Then
            GoTo SALE
         Else
         'bloqueado por mic
            GoTo CANCELA
            cmdproce.Enabled = True
            proce.Enabled = True
            cmdrepo.Enabled = True
            DoEvents
            lpb.Visible = False
            DoEvents
            Exit Sub
        End If
      End If
'      If F1 = 199 Then Stop
      xl.Cells(F1, 7) = cli_llave10!cli_nombre ' com_llave!com_descripcion
      xl.Cells(F1, 6) = Trim(cli_llave10!cli_ruc_esposo)
SALE:
      xl.Cells(F1, 3) = wc_sunat
      xl.Cells(F1, 4) = Format(wc_serie, "0000")
      xl.Cells(F1, 5) = Format(wc_numfac, "0000")
      wc_sunat = ""
      wc_serie = ""
      wc_numfac = ""
  
      wc_total = WC_HABER
      wc_bruto = Val(wc_total - wc_afect)
      If wc_afect = 0 Then
        xl.Cells(F1, 8) = wc_bruto
        xl.Cells(F1, 12) = ""
      Else
        'xl.Cells(F1, 9) = wc_bruto
        xl.Cells(F1, 8) = wc_bruto
      End If
      If wflag_vocucher = "A" Then
        xl.Cells(F1, 2) = "'" & Format(WS_VOU2, "########")
        'xl.Cells(F1, 9) = Trim(wc_detalle)
      End If
      If wc_afect <> 0 Then xl.Cells(F1, 10) = wc_afect Else xl.Cells(F1, 10) = ""
      xl.Cells(F1, 11) = wc_total
      'xl.Cells(F1, 16) = WS_VOU2
      WS_VOU2 = llave_rep01!MOV_NRO_VOUCHER
DALE1:
      wc_voucher = llave_rep01!MOV_NRO_VOUCHER
      wc_total = 0
      WC_HABER = 0
      WC_DEBE = 0
      wc_bruto = 0
      wc_inaf = 0
      wc_afect = 0
   End If
   wc_importe = Val(llave_rep01!mov_importe)
   wc_desTc = ""
   wc_codclie = llave_rep01!MOV_codclie
   wc_cp = llave_rep01!MOV_CP
   If Trim(llave_rep01!MOV_MONEDA) = "D" And Trim(llave_rep01!MOV_FLAG_TC) = "" Then
'     MsgBox pr_llave!MOV_DETALLE
     SQ_OPER = 1
     PUB_CAL_INI = llave_rep01!MOV_fecha_EMI
     PUB_CAL_FIN = llave_rep01!MOV_fecha_EMI
     PUB_CODCIA = LK_CODCIA
     LEER_CAL_LLAVE
     If Nulo_Valor0(cal_llave!cal_tipo_cambio) = 0 Then
       MsgBox "Falta tipo de cambio, Fecha: " & llave_rep01!MOV_fecha_EMI, 48, Pub_Titulo
       GoTo CANCELA
       Exit Sub
     End If
     wc_importe = Format((wc_importe * Val(cal_llave!cal_tipo_cambio)), "0.00")
     wc_desTc = " T.C. en S/. " & Format(cal_llave!cal_tipo_cambio, "0.000")
   ElseIf Trim(llave_rep01!MOV_MONEDA) = "D" And Trim(llave_rep01!MOV_FLAG_TC) = "A" Then
      wc_importe = Format((wc_importe * Val(llave_rep01!MOV_TIPO_CAMBIO)), "0.00")
      wc_desTc = " T.C. en S/. " & Format(llave_rep01!MOV_TIPO_CAMBIO, "0.000")
   End If
  
   If llave_rep01!MOV_DH = "D" Then
     WC_DEBE = WC_DEBE + wc_importe
   Else
     WC_HABER = WC_HABER + Val(llave_rep01!mov_importe)
   End If
   If WTIPMOV = 1 Then
'     Print llave_rep01!MOV_GLOSA
     If Left(Trim(llave_rep01!MOV_CODCTA), 2) = "42" Then
       WC_HABER = WC_HABER + wc_importe
       wc_cp = llave_rep01!MOV_CP
       wc_codclie = llave_rep01!MOV_codclie
     End If
   Else
     If Left(Trim(llave_rep01!MOV_CODCTA), 2) = "12" Then
      WC_HABER = WC_HABER + wc_importe
      wc_cp = llave_rep01!MOV_CP
      wc_codclie = llave_rep01!MOV_codclie
     End If
   End If
   If Left(Trim(llave_rep01!MOV_CODCTA), 3) = "401" Then
     wc_afect = wc_afect + wc_importe
     wc_inaf = 0
   End If
   
   wc_fecha = "'" & Format(llave_rep01!MOV_fecha_EMI, "dd/mm/yy")
   WS_VOU2 = Nulo_Valor0(llave_rep01!MOV_NRO_VOUCHER)
   wc_detalle = llave_rep01!MOV_DETALLE
  

   If WTIPMOV = 1 Then
     If Left(Trim(llave_rep01!MOV_CODCTA), 2) = "42" Or Left(Trim(llave_rep01!MOV_CODCTA), 2) = "46" Then
          wc_cuenta = llave_rep01!MOV_CODCTA
          wc_numfac = llave_rep01!MOV_numfac
          wc_sunat = Val(llave_rep01!MOV_SUNAT)
          wc_serie = llave_rep01!MOV_serie
     End If
   Else
     If Left(Trim(llave_rep01!MOV_CODCTA), 2) = "12" Then
         wc_cuenta = llave_rep01!MOV_CODCTA
         wc_numfac = llave_rep01!MOV_numfac
         wc_sunat = Val(llave_rep01!MOV_SUNAT)
         wc_serie = llave_rep01!MOV_serie
     End If
   End If
   llave_rep01.MoveNext
 Loop
 ' siempre hay datos al final
If WC_HABER = 0 And WTIPMOV = 1 Then GoTo termina
F1 = F1 + 1
SQ_OPER = 1
If wc_codclie = 0 Then GoTo SALE33
pu_codcia = LK_CODCIA
pu_cp = wc_cp
pu_codclie = wc_codclie
LEER_CLI_LLAVE
If cli_llave.EOF Then
    vbRespContinua = MsgBox("Verificar Codigo de Cliente " & wc_codclie & vbCrLf & "Desea Continuar....???", vbYesNo + vbQuestion, Pub_Titulo)
    If vbRespContinua = vbYes Then
       GoTo CONTINUA
    Else
       GoTo CANCELA
    End If
   Exit Sub
End If
CONTINUA:
xl.Cells(F1, 7) = cli_llave!cli_nombre ' com_llave!com_descripcion
xl.Cells(F1, 6) = Trim(cli_llave!cli_ruc_esposo)
SALE33:
xl.Cells(F1, 1) = wc_fecha
xl.Cells(F1, 3) = wc_sunat
xl.Cells(F1, 4) = Format(wc_serie, "000")
xl.Cells(F1, 5) = Val(wc_numfac)
wc_total = WC_HABER
wc_bruto = Val(wc_total - wc_afect)
If wc_afect = 0 Then
xl.Cells(F1, 8) = wc_bruto
xl.Cells(F1, 12) = ""
Else
'xl.Cells(F1, 9) = wc_bruto
xl.Cells(F1, 8) = wc_bruto
End If
If wflag_vocucher = "A" Then
    wranF = "G"
    'xl.Columns(wranF).ColumnWidth = 7
    wranF = "I"
    'xl.Columns(wranF).ColumnWidth = 0
    wranF = "J"
    'xl.Columns(wranF).ColumnWidth = 0
   xl.Cells(F1, 2) = "'" & Format(WS_VOU2, "########")
   'xl.Cells(F1, 9) = wc_detalle
End If
xl.Cells(F1, 10) = wc_afect
xl.Cells(F1, 11) = wc_total
'xl.Cells(F1, 16) = WS_VOU2
termina:
wc_total = 0
WC_HABER = 0
WC_DEBE = 0
wc_bruto = 0
wc_inaf = 0
wc_afect = 0

wc_fila_F = F1
'If WTIPMOV = 2 Then
'  wranF = "A" & 6 & ":O" & wc_fila_F
'  xl.Application.Range(wranF).Select
'  xl.Application.Range(wranF).Sort Key1:=Range("E1"), Order1:=xlAscending, Key2:=Range("A1"), Order2:=xlAscending, Key3:=Range("F1"), Order3:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
'Else
'  wranF = "A1"
'  xl.Application.Range(wranF).Select
'  wranF = "A" & 6 & ":O" & wc_fila_F
'  xl.Application.Range(wranF).Select
'  xl.Application.Range(wranF).Sort Key1:=Range("O1"), Order1:=xlAscending, Header:=xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom
'End If
'xl.Application.Visible = True
If WTIPMOV = 2 Then
  wranF = "A" & 6 & ":P" & wc_fila_F
  'xl.Application.Worksheets("Hoja1").Range(wranF).Sort Key1:=xl.Application.Worksheets("Hoja1").Range("B1") ' , Key2:=xl.Application.Worksheets("Hoja1").Range("F1"), Key1:=xl.Application.Worksheets("Hoja1").Range("G1")
  'xl.Application.Worksheets("Hoja1").Range(wranF).Sort Key1:=xl.Application.Worksheets("Hoja1").Range("E1"), Key2:=xl.Application.Worksheets("Hoja1").Range("F1"), Key3:=xl.Application.Worksheets("Hoja1").Range("G1")
Else
  wranF = "A" & 6 & ":P" & wc_fila_F
  xl.Application.Worksheets("Hoja1").Range(wranF).Sort Key1:=xl.Application.Worksheets("Hoja1").Range("O1")
End If


  FILAX = 0
  For fila = wc_fila_I To wc_fila_F
     FILAX = FILAX + 1
     'xl.Cells(fila, 4) = FILAX
  Next fila
If wflag_vocucher = "A" Then
  wranF = "H"
  'xl.Columns(wranF).ColumnWidth = 6.5
  wranF = "I"
  'xl.Columns(wranF).ColumnWidth = 38
  wranF = "K"
  'xl.Columns(wranF).ColumnWidth = 0
  wranF = "J"
  'xl.Columns(wranF).ColumnWidth = 0
End If
'Else
'   wranF = "D"
'   xl.Columns(wranF).ColumnWidth = 0
'End If

wran1 = "H" & wc_fila_I
wran2 = "H" & wc_fila_F
wranF = "H" & wc_fila_F + 1
xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
wran1 = "I" & wc_fila_I
wran2 = "I" & wc_fila_F
wranF = "I" & wc_fila_F + 1
xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
wran1 = "J" & wc_fila_I
wran2 = "J" & wc_fila_F
wranF = "J" & wc_fila_F + 1
xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
wran1 = "K" & wc_fila_I
wran2 = "K" & wc_fila_F
wranF = "K" & wc_fila_F + 1
xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
wran1 = "L" & wc_fila_I
wran2 = "L" & wc_fila_F
wranF = "L" & wc_fila_F + 1
xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
wran1 = "M" & wc_fila_I
wran2 = "M" & wc_fila_F
wranF = "M" & wc_fila_F + 1
xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
wran1 = "N" & wc_fila_I
wran2 = "N" & wc_fila_F
wranF = "N" & wc_fila_F + 1
xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"

xl.Cells(wc_fila_F + 1, 3) = "TOTAL GENERAL = S/."

wranF = "G" & wc_fila_F + 1 & ":N" & wc_fila_F + 1
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
xl.Worksheets(1).Rows(wc_fila_F + 1).RowHeight = 17

  xl.Cells(1, 1) = Trim(Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)))
  xl.Cells(2, 1) = Trim(Mid(proce.Text, 5, Len(proce.Text)))
  xl.Cells(3, 1) = "'PERIODO : " & UCase(Format(LK_FECHA_COP1, "mmmm")) & " (" & Format(LK_FECHA_COP1, "dd/mm/yyyy") & " al " & Format(LK_FECHA_COP2, "dd/mm/yyyy") & ")"
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.Application.Visible = True
  DoEvents
  Set xl = Nothing
  Screen.MousePointer = 0
 cmdproce.Enabled = True
    proce.Enabled = True
    cmdrepo.Enabled = True
    DoEvents
    lpb.Visible = False
    DoEvents
Exit Sub

WEXCEL:
'  lblProceso.Caption = "Abriendo , Archivo REGISTROS.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open CONS_ADMIN & "CONTABILIDAD\REGISTROS.xls", 0, True, 4, WPAS, WPAS
Return

CANCELA:
  If xl Is Nothing Then
  Else
   xl.Application.Visible = True
   Set xl = Nothing
  End If
  Screen.MousePointer = 0

End Sub

Public Sub CAJABANCOS(WTIPMOV)
Dim wc_importe As Currency
Dim filad As Integer
Dim as_f_fecha(500) As String
Dim as_f_cuenta(500) As String
Dim as_f_descri(500) As String
Dim as_f_vou(500) As String
Dim as_f_sunat(500) As Integer
Dim as_f_docu(500) As String
Dim as_f_importe(500) As String
Dim acu_fijos As Integer
Dim wc_voucher  As Integer
Dim wc_fecha As String
Dim wc_cuenta As String
Dim wc_sunat As String
Dim wc_serie As String
Dim wc_numfac As String
Dim wc_fila_I As Integer
Dim wc_fila_F As Integer
Dim wc_dh As String * 1
Dim ws_tipmov As Integer
Dim wc_saldo_debe As Currency
Dim wc_saldo_haber As Currency
Dim wc_saldo_final As Currency
Dim as_ctades(500) As String * 2



Dim WS_FLAG As String * 1
 pub_cadena = "SELECT * FROM MOVICONT WHERE MOV_CODCIA = ? AND MOV_TIPMOV = ? AND (MOV_FECHA >= ? AND MOV_FECHA <= ?)  AND MOV_NRO_MES = " & LK_NRO_MES & " ORDER BY MOV_DH DESC , MOV_FECHA_EMI, MOV_NRO_VOUCHER"
 Set PS_REP01 = CN.CreateQuery("", pub_cadena)
 PS_REP01(0) = 0
 PS_REP01(1) = 0
 PS_REP01(2) = LK_FECHA_DIA
 PS_REP01(3) = LK_FECHA_DIA
 Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
 PS_REP01(0) = LK_CODCIA
 PS_REP01(1) = WTIPMOV
 PS_REP01(2) = LK_FECHA_COP1
 PS_REP01(3) = LK_FECHA_COP2
 llave_rep01.Requery
 If llave_rep01.EOF Then
   MsgBox "No existen Movimientos.", 48, Pub_Titulo
   Exit Sub
 End If
wc_importe = 0
SQ_OPER = 2
PUB_TIPREG = 56
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
fila = 1
Do Until tab_mayor.EOF
   as_ctades(fila) = Trim(tab_mayor!tab_nomlargo)
   fila = fila + 1
   tab_mayor.MoveNext
Loop


 wc_saldo_debe = 0
 wc_saldo_haber = 0
 wc_saldo_final = 0
 GoSub WEXCEL
 
    PUB_CUENTA = "10"
    PUB_CODCIA = LK_CODCIA
    SQ_OPER = 1
    LEER_COM_LLAVE
    If com_llave.EOF Then
      MsgBox "Definir Cuenta Contable ...", 48, Pub_Titulo
      Exit Sub
    Else
      WS_SAL_ANTERIOR = Nulo_Valor0(com_llave!COM_DEB_ANO) - Nulo_Valor0(com_llave!COM_HAB_ANO)
    End If
    
    If LK_NRO_MES = 0 Then
      WS_SAL_ANTERIOR = 0
    Else
      JALA_SALDO PUB_CUENTA, 3
      WS_SAL_ANTERIOR = PUB_IMPORTE_DEB - PUB_IMPORTE_HAB
    End If
    
    wc_saldo_debe = WS_SAL_ANTERIOR
  ' INGRESOS A CAJA Y BANCOS
  wc_fila_F = 0
 wc_fila_I = 3
 xl.Cells(wc_fila_I + 2, 1) = "SALDO   :"
 xl.Cells(wc_fila_I + 2, 2) = Format(WS_SAL_ANTERIOR, "0.00;(0.00)")
 xl.Cells(wc_fila_I + 3, 1) = "MOVIM. :"
 xl.Cells(wc_fila_I + 3, 2) = "INGRESOS:"
 xl.Cells(wc_fila_I + 4, 1) = "FECHA"
 xl.Cells(wc_fila_I + 4, 2) = "CUENTA"
 xl.Cells(wc_fila_I + 4, 3) = "CONCEPTO"
 xl.Cells(wc_fila_I + 3, 4) = "COD."
 xl.Cells(wc_fila_I + 4, 4) = "SUNAT"
 xl.Cells(wc_fila_I + 3, 5) = "NRO."
 xl.Cells(wc_fila_I + 4, 5) = "VOUC."
 xl.Cells(wc_fila_I + 4, 6) = "DOC."
 xl.Cells(wc_fila_I + 4, 7) = "DEBE."
 F1 = wc_fila_I + 4

 WC_HABER = 0
 WC_DEBE = 0
 wc_dh = Trim(llave_rep01!MOV_DH)
 pb.Max = llave_rep01.RowCount
 pb.Min = 0
 pb.Value = 0
 acu_fijos = 0
 WS_FLAG = ""
 Do Until llave_rep01.EOF
   pb.Value = pb.Value + 1
 '  Print Val(llave_rep01!MOV_IMPORTE)
'      If Trim(llave_rep01!MOV_NRO_VOUCHER) = 1 Then Stop
   filad = 1
   Do Until Val(as_ctades(filad)) = 0
     If Trim(Left(Trim(llave_rep01!MOV_CODCTA), 2)) = Trim(as_ctades(filad)) Then
         GoTo SALE_CAJA
     End If
     filad = filad + 1
   Loop
   
   'If Trim(llave_rep01!MOV_CODCTA) = "10101" Then
   If Left(Trim(llave_rep01!MOV_CODCTA), 2) = "10" Then
      If Trim(llave_rep01!MOV_PLANTILLA) = 100 Then
        If Trim(llave_rep01!MOV_DH) = "D" Then
         GoTo SALE_CAJA
        End If
'        Stop
      End If
      If Trim(llave_rep01!MOV_PLANTILLA) = 126 Then
        If Trim(llave_rep01!MOV_DH) = "H" Then
        GoTo SALE_CAJA
        End If
 '       Stop
      End If
   End If
   'Print Val(llave_rep01!MOV_IMPORTE)
   If wc_dh <> Trim(llave_rep01!MOV_DH) Then
     wran1 = "G" & wc_fila_I + 5
     wran2 = "G" & F1
     wranF = "G" & F1 + 1
     xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
     F1 = F1 + 1
     wc_saldo_debe = wc_saldo_debe + Val(xl.Range(wranF))
     F1 = F1 + 1
     xl.Cells(F1, 5) = "SUMA DEL DEBE = "
     xl.Cells(F1, 7) = wc_saldo_debe
   'xl.Application.Visible = True
     'wranF = "G" & wc_fila_F + 1 & ":J" & F1 + 1
     'xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
     xl.Worksheets(1).Rows(F1 + 1).RowHeight = 17
     wc_fila_I = F1
     xl.Cells(wc_fila_I + 3, 1) = "MOVIM. :"
     xl.Cells(wc_fila_I + 3, 2) = "EGRESOS:"
     xl.Cells(wc_fila_I + 4, 1) = "FECHA"
     xl.Cells(wc_fila_I + 4, 2) = "CUENTA"
     xl.Cells(wc_fila_I + 4, 3) = "CONCEPTO"
     xl.Cells(wc_fila_I + 3, 4) = "COD."
     xl.Cells(wc_fila_I + 4, 4) = "SUNAT"
     xl.Cells(wc_fila_I + 3, 5) = "NRO."
     xl.Cells(wc_fila_I + 4, 5) = "VOUC."
     xl.Cells(wc_fila_I + 4, 6) = "DOC."
     xl.Cells(wc_fila_I + 4, 7) = "HABER"
     F1 = wc_fila_I + 4
     wc_dh = llave_rep01!MOV_DH
     WS_FLAG = "A"
   End If
   wc_fecha = Format(llave_rep01!MOV_fecha_EMI, "dd/mm/yy")
   wc_sunat = Val(llave_rep01!MOV_SUNAT)
   wc_cuenta = llave_rep01!MOV_CODCTA
   wc_voucher = llave_rep01!MOV_NRO_VOUCHER
   ws_tipmov = Val(llave_rep01!MOV_TIPMOV)
   If llave_rep01!MOV_serie <> 0 Then
     If Val(llave_rep01!MOV_numfac) <> 0 Then
       wc_numfac = Str(llave_rep01!MOV_serie) & "-" & Str(Format(llave_rep01!MOV_numfac, "###############"))
     Else
       wc_numfac = Str(llave_rep01!MOV_serie)
     End If
   Else
     wc_numfac = Format(llave_rep01!MOV_numfac, "###########")
   End If
   '¡SQ_OPER = 1
   'PUB_CUENTA = wc_cuenta
   'PUB_CODCIA = LK_CODCIA
   'LEER_COM_LLAVE
   If Left(Trim(llave_rep01!MOV_CODCTA), 3) = "422" And Trim(llave_rep01!MOV_DH) = "H" Then
    acu_fijos = acu_fijos + 1
    as_f_fecha(acu_fijos) = wc_fecha
    as_f_cuenta(acu_fijos) = wc_cuenta ' PUB_CUENTA
    as_f_descri(acu_fijos) = Trim(llave_rep01!MOV_DETALLE) ' com_llave!com_DESCRIPCION
    If ws_tipmov = 1 Then
       as_f_vou(acu_fijos) = "R.V " & Format(wc_voucher, "000")
    ElseIf ws_tipmov = 2 Then
       as_f_vou(acu_fijos) = "R.C " & Format(wc_voucher, "000")
    ElseIf ws_tipmov = 3 Then
       as_f_vou(acu_fijos) = "C.B " & Format(wc_voucher, "000")
    ElseIf ws_tipmov = 4 Then
       as_f_vou(acu_fijos) = "O.T " & Format(wc_voucher, "000")
    End If
    as_f_sunat(acu_fijos) = wc_sunat
    as_f_docu(acu_fijos) = wc_numfac
    wc_importe = Val(llave_rep01!mov_importe)
    If Trim(llave_rep01!MOV_MONEDA) = "D" And Trim(llave_rep01!MOV_FLAG_TC) = "" Then
          SQ_OPER = 1
          PUB_CAL_INI = llave_rep01!MOV_fecha_EMI
          PUB_CAL_FIN = llave_rep01!MOV_fecha_EMI
          PUB_CODCIA = LK_CODCIA
          LEER_CAL_LLAVE
          If Nulo_Valor0(cal_llave!cal_tipo_cambio) = 0 Then
            MsgBox "Falta tipo de cambio, Fecha: " & llave_rep01!MOV_fecha_EMI, 48, Pub_Titulo
            GoTo CANCELA
            Exit Sub
          End If
          wc_importe = Format((wc_importe * Val(cal_llave!cal_tipo_cambio)), "0.00")
          'wc_desTc = " T.C. en S/. " & Format(cal_llave!cal_tipo_cambio, "0.000")
    ElseIf Trim(llave_rep01!MOV_MONEDA) = "D" And Trim(llave_rep01!MOV_FLAG_TC) = "A" Then
           wc_importe = Format((wc_importe * Val(llave_rep01!MOV_TIPO_CAMBIO)), "0.00")
           'wc_desTc = " T.C. en S/. " & Format(llave_rep01!MOV_TIPO_CAMBIO, "0.000")
    End If
    as_f_importe(acu_fijos) = wc_importe * -1
    GoTo SALE_CAJA
   End If
   If WS_FLAG = "A" Then
     XFILA = 1
     Do Until Val(Nulo_Valor0(as_f_importe(XFILA))) = 0
     F1 = F1 + 1
        xl.Cells(F1, 1) = "'" & as_f_fecha(XFILA)
        xl.Cells(F1, 2) = as_f_cuenta(XFILA)
        xl.Cells(F1, 3) = as_f_descri(XFILA)
        xl.Cells(F1, 4) = as_f_sunat(XFILA)
        xl.Cells(F1, 5) = as_f_vou(XFILA)
        xl.Cells(F1, 6) = as_f_docu(XFILA)
        xl.Cells(F1, 7) = Val(as_f_importe(XFILA))
     XFILA = XFILA + 1
    Loop
    WS_FLAG = ""
   End If
   
   F1 = F1 + 1
   xl.Cells(F1, 1) = "'" & wc_fecha
   xl.Cells(F1, 2) = wc_cuenta
   xl.Cells(F1, 3) = llave_rep01!MOV_DETALLE ' com_llave!com_DESCRIPCION
   xl.Cells(F1, 4) = " "
   If Val(wc_sunat) <> 0 Then xl.Cells(F1, 4) = wc_sunat
   If ws_tipmov = 1 Then
       xl.Cells(F1, 5) = "R.V " & Format(wc_voucher, "000")
    ElseIf ws_tipmov = 2 Then
       xl.Cells(F1, 5) = "R.C " & Format(wc_voucher, "000")
    ElseIf ws_tipmov = 3 Then
       xl.Cells(F1, 5) = "C.B " & Format(wc_voucher, "000")
    ElseIf ws_tipmov = 4 Then
       xl.Cells(F1, 5) = "O.T " & Format(wc_voucher, "000")
    End If
    xl.Cells(F1, 6) = wc_numfac
   'Print llave_rep01!MOV_DH
    wc_importe = Val(llave_rep01!mov_importe)
    If Trim(llave_rep01!MOV_MONEDA) = "D" And Trim(llave_rep01!MOV_FLAG_TC) = "" Then
          SQ_OPER = 1
          PUB_CAL_INI = llave_rep01!MOV_fecha_EMI
          PUB_CAL_FIN = llave_rep01!MOV_fecha_EMI
          PUB_CODCIA = LK_CODCIA
          LEER_CAL_LLAVE
          If Nulo_Valor0(cal_llave!cal_tipo_cambio) = 0 Then
            MsgBox "Falta tipo de cambio, Fecha: " & llave_rep01!MOV_fecha_EMI, 48, Pub_Titulo
            GoTo CANCELA
            Exit Sub
          End If
          wc_importe = Format((wc_importe * Val(cal_llave!cal_tipo_cambio)), "0.00")
          'wc_desTc = " T.C. en S/. " & Format(cal_llave!cal_tipo_cambio, "0.000")
     ElseIf Trim(llave_rep01!MOV_MONEDA) = "D" And Trim(llave_rep01!MOV_FLAG_TC) = "A" Then
           wc_importe = Format((wc_importe * Val(llave_rep01!MOV_TIPO_CAMBIO)), "0.00")
           'wc_desTc = " T.C. en S/. " & Format(llave_rep01!MOV_TIPO_CAMBIO, "0.000")
     End If
   xl.Cells(F1, 7) = wc_importe
SALE_CAJA:
   llave_rep01.MoveNext
 Loop
 xl.Application.Visible = True
 ' siempre hay datos al final
wran1 = "G" & wc_fila_I + 5
wran2 = "G" & F1
wranF = "G" & F1 + 1
xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
wc_saldo_haber = wc_saldo_haber + Val(xl.Range(wranF))
'wranF = "G" & wc_fila_F + 1 & ":J" & F1 + 1
'xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
xl.Worksheets(1).Rows(F1 + 1).RowHeight = 17
wc_saldo_final = wc_saldo_debe - wc_saldo_haber


'If LK_NRO_MES = 0 Then
'   WS_SAL_ANTERIOR = 0
'Else
'   JALA_SALDO PUB_CUENTA, 1
'   'JALA_SALDO PUB_CUENTA, 3
'   wc_saldo_final = PUB_IMPORTE_DEB - PUB_IMPORTE_HAB
'End If
 
 

F1 = F1 + 1
F1 = F1 + 1
xl.Cells(F1, 6) = "SALDO = "
xl.Cells(F1, 7) = wc_saldo_final
F1 = F1 + 1
xl.Cells(F1, 5) = "SUMA DEL HABER = "
xl.Cells(F1, 7) = wc_saldo_final + wc_saldo_haber



  xl.Cells(1, 1) = Trim(Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)))
  xl.Cells(2, 1) = Trim(Mid(proce.Text, 6, Len(proce.Text)))
  xl.Cells(3, 1) = "'PERIODO : " & UCase(Format(LK_FECHA_COP1, "mmmm")) & " (" & Format(LK_FECHA_COP1, "dd/mm/yyyy") & " al " & Format(LK_FECHA_COP2, "dd/mm/yyyy") & ")"
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.Application.Visible = True
  DoEvents
  Set xl = Nothing
  Screen.MousePointer = 0
 
Exit Sub

WEXCEL:
'  lblProceso.Caption = "Abriendo , Archivo REGISTROS.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open CONS_ADMIN & "CONTABILIDAD\CAJABANCOS.xls", 0, True, 4, WPAS, WPAS
Return
Exit Sub
CANCELA:

  xl.Cells(1, 1) = Trim(Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)))
  xl.Cells(2, 1) = Trim(Mid(proce.Text, 6, Len(proce.Text)))
  xl.Cells(3, 1) = "'PERIODO : " & UCase(Format(LK_FECHA_COP1, "mmmm")) & " (" & Format(LK_FECHA_COP1, "dd/mm/yyyy") & " al " & Format(LK_FECHA_COP2, "dd/mm/yyyy") & ")"
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.Application.Visible = True
  DoEvents
  Set xl = Nothing
  Screen.MousePointer = 0

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
Dim tf, t, tc
Dim SALE As Boolean
Dim Wsec

'If WMODO = "C" Then Exit Sub

'If cop_llave!COP_FLAG_MAYORIZACION = "M" Then
 'MsgBox "Ojo estaba Mayorizado..."
'End If

If KeyCode = 46 Then
  If gridigv.Rows <= 2 Then
    gridigv.Rows = 1
  Else
   gridigv.RemoveItem gridigv.Row
   Exit Sub
  End If
End If
If KeyCode = 45 Then
    Wsec = Wsec + 1
  '  If Trim(gridigv.TextMatrix(gridigv.Row + 1, 11)) = "8" Then
  '       Exit Sub
  '  Else
  '    If Trim(gridigv.TextMatrix(gridigv.Row + 1, 0)) = "T" Then Exit Sub
  '  End If
    'If Val(gridigv.TextMatrix(gridigv.Row, 4)) = 0 And Val(gridigv.TextMatrix(gridigv.Row, 5)) = 0 Then Exit Sub
    gridigv.Rows = gridigv.Rows + 1
    'gridigv.AddItem "", gridigv.Row + 1
    gridigv.TextMatrix(gridigv.Rows - 1, 0) = ""
    gridigv.Row = gridigv.Rows - 1
    gridigv.Col = 0
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
gridigv.Text = TEXTOVAR.Text
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
'If gridigv.Col = 1 Then Consistencias gridigv, TEXTOVAR, KeyAscii
'If gridigv.Col = 4 Then Consistencias gridigv, TEXTOVAR, KeyAscii
If KeyAscii <> 13 Then
   GoTo fin
End If
If gridigv.Col = 1 Or gridigv.Col = 4 Then
  If Val(TEXTOVAR.Text) > 99 Then
    Azul3 TEXTOVAR, TEXTOVAR
    Exit Sub
  End If
End If
' grabar
SQ_OPER = 1
PUB_CUENTA = Trim(TEXTOVAR.Text)
PUB_CODCIA = LK_CODCIA
LEER_COM_LLAVE
If com_llave.EOF Then
  MsgBox "Cuenta no Existe", 48, Pub_Titulo
  Azul3 TEXTOVAR, TEXTOVAR
  Exit Sub
End If
If com_llave!com_nivel <> 1 Then
  MsgBox "Cuenta debe ser principal", 48, Pub_Titulo
  Azul3 TEXTOVAR, TEXTOVAR
  Exit Sub
End If

pub_cadena = "DELETE TABLAS  WHERE TAB_CODCIA = '" & LK_CODCIA & "' AND TAB_TIPREG =  56 "
CN.Execute pub_cadena, rdExecDirect
SQ_OPER = 2
PUB_TIPREG = 56
LEER_TAB_LLAVE
For fila = 1 To gridigv.Rows - 1
    tab_mayor.AddNew
    tab_mayor!TAB_CODCIA = LK_CODCIA
    tab_mayor!TAB_TIPREG = 56
    tab_mayor!TAB_NUMTAB = fila
    tab_mayor!tab_nomlargo = gridigv.TextMatrix(fila, 0)
    tab_mayor!tab_nomcorto = ""
    tab_mayor!TAB_CODART = 0
    tab_mayor!TAB_CODCLIE = 0
    tab_mayor.Update
Next fila

If gridigv.Row >= gridigv.Rows - 1 Then
Else
  gridigv.Row = gridigv.Row + 1
End If
gridigv.SetFocus
TEXTOVAR.Visible = False

fin:

End Sub

Private Sub DESTINO_ANTERIOR()
Dim PSCOV_CUENTA2 As rdoQuery
Dim cox_cuenta  As rdoResultset
Dim WS_CUENTA As String
Dim ws_parcial As Currency
Dim ws_dh_cov As String * 1

Dim wcadena As String
Dim wvalor  As String * 1
' CHEQUE DE MES CERRADO 0 ABIERTO
wcadena = ""
SQ_OPER = 2
PUB_TIPREG = 155
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
If tab_mayor.EOF Then
  MsgBox "Crear tab_tipreg = 155 para seguridas", 48, Pub_Titulo
Else
    Do Until tab_mayor.EOF
      If tab_mayor!TAB_NUMTAB = Val(Format(LK_FECHA_COP1, "yyyy")) Then
        wcadena = Trim(tab_mayor!tab_nomlargo)
      End If
      tab_mayor.MoveNext
    Loop
    wvalor = Mid(wcadena, LK_NRO_MES + 1, 1)
    If wvalor = "1" Then
       MsgBox "<<<< Mes CERRADO Operaciones >>>>", vbCritical, Pub_Titulo
       Exit Sub
    End If
End If

cade = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND COV_FECHA_VOUCHER >= ? AND COV_FECHA_VOUCHER <= ?  AND COV_NRO_MES = " & LK_NRO_MES & " ORDER BY COV_CODCTA, COV_FECHA_VOUCHER, COV_NRO_MOV"
Set PSCOV_CUENTA2 = CN.CreateQuery("", cade)
PSCOV_CUENTA2(0) = 0
PSCOV_CUENTA2(1) = LK_FECHA_DIA
PSCOV_CUENTA2(2) = LK_FECHA_DIA
Set cox_cuenta = PSCOV_CUENTA2.OpenResultset(rdOpenKeyset, rdConcurValues)

cade = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND COV_FECHA_VOUCHER >= ? AND COV_FECHA_VOUCHER <= ?  AND COV_NRO_MES = " & LK_NRO_MES & " ORDER BY COV_CODCTA, COV_FECHA_VOUCHER, COV_NRO_MOV"
Set PSCOV_CUENTA2 = CN.CreateQuery("", cade)
PSCOV_CUENTA2(0) = 0
PSCOV_CUENTA2(1) = LK_FECHA_DIA
PSCOV_CUENTA2(2) = LK_FECHA_DIA
Set cox_cuenta = PSCOV_CUENTA2.OpenResultset(rdOpenKeyset, rdConcurValues)

lpb.Caption = "Eliminando Movimientos Automaticos"
ffecha1 = Format(LK_FECHA_COP1, "yyyy/mm/dd")
ffecha2 = Format(LK_FECHA_COP2, "yyyy/mm/dd")
wpub_cadena = "DELETE COMOV  WHERE ( COV_FLAG_AUTOMATICA = 'M'  ) AND COV_CODCIA = '" & LK_CODCIA & "' AND COV_FECHA_VOUCHER >=  '" & ffecha1 & "'  AND COV_FECHA_VOUCHER <=  '" & ffecha2 & "'"
CN.Execute wpub_cadena, rdExecDirect
lpb.Caption = ""
pub_mensaje = "Procesoar los Destinos ¿Desea Continuar... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
    Exit Sub
End If


PSCOV_CUENTA2.rdoParameters(0) = LK_CODCIA
PSCOV_CUENTA2.rdoParameters(1) = LK_FECHA_COP1
PSCOV_CUENTA2.rdoParameters(2) = LK_FECHA_COP2
cox_cuenta.Requery


' QUITE EN LIMA
'ws_fecha_proc = #1/1/00# ' por lo de piura
'If LK_FECHA_COP1 = ws_fecha_proc Then
'   Exit Sub
'End If

PSTEMP_LLAVE(0) = LK_CODCIA
PSTEMP_LLAVE(1) = LK_FECHA_COP1
PSTEMP_LLAVE(2) = #1/1/2020#
temp_llave.Requery
If temp_llave.EOF Then
  ws_nro_voucher = 0
  ws_fecha_voucher = LK_FECHA_COP1
Else
  temp_llave.MoveLast
  ws_nro_voucher = temp_llave!COV_NRO_VOUCHER
  ws_fecha_voucher = temp_llave!COV_FECHA_VOUCHER
End If


'barra.Max = cox_cuenta.RowCount
CONTADOR = 0
lpb.Caption = "Creando Movimientos Automaticos"
pb.Max = cox_cuenta.RowCount
pb.Min = 0
pb.Value = 0
pb.Visible = True

Do Until cox_cuenta.EOF
      pb.Value = pb.Value + 1
      
      SQ_OPER = 1
      PUB_CUENTA = cox_cuenta!COV_CODCTA
'      If Left(cox_cuenta!COV_CODCTA, 1) = "9" Then Stop
      PUB_CODCIA = LK_CODCIA
      LEER_COM_LLAVE
      If com_llave.EOF Then
         MsgBox "Revisar el diario ...Cuenta : " & PUB_CUENTA
         GoTo fin
      End If
      ww_ff = 0
      If com_llave!COM_POR_AUTOM_D <> 0 Then ww_ff = 1
      If com_llave!COM_POR_AUTOM_D2 <> 0 Then ww_ff = 2
      If com_llave!COM_POR_AUTOM_D3 <> 0 Then ww_ff = 3
      If com_llave!COM_POR_AUTOM_D4 <> 0 Then ww_ff = 4
      If com_llave!COM_POR_AUTOM_D5 <> 0 Then ww_ff = 5
      WS_NRO_MOV = 0
      
      WS_SUMA = Val(Nulo_Valors(com_llave!com_cuenta_AUTOM_D)) + Val(Nulo_Valors(com_llave!COM_CUENTA_AUTOM_D2)) + Val(Nulo_Valors(com_llave!COM_CUENTA_AUTOM_D3)) + Val(Nulo_Valors(com_llave!COM_CUENTA_AUTOM_D4)) + Val(Nulo_Valors(com_llave!COM_CUENTA_AUTOM_D5))
          
      If WS_SUMA <> 0 Then
         WS_SUMA = 0
         'ws_nro_voucher = cox_cuenta!COV_NRO_VOUCHER
         ws_glosa = cox_cuenta!COV_glosa
         ws_fecha_proc = cox_cuenta!COV_FECHA_VOUCHER
         WS_IMPORTE = cox_cuenta!COV_IMPORTE
'         If WS_IMPORTE = 8 Then Stop
         
         ws_codusu = cox_cuenta!COV_CODUSU
         ' PRUEBA
         
         If cox_cuenta!COV_DH = "H" And Left(cox_cuenta!COV_CODCTA, 1) = "6" Then
            ws_dh = "H"
         Else
            ws_dh = "D"
         End If
         ws_nro_voucher = ws_nro_voucher + 1
         WS_CUENTA = com_llave!com_cuenta_AUTOM_D
         ws_dh_cov = cox_cuenta!COV_DH
          If Trim(com_llave!com_cuenta_AUTOM_D) <> "" Then
            ws_por = com_llave!COM_POR_AUTOM_D
            ww_fff = 1
            GoSub graba_autom
         End If
         WS_CUENTA = com_llave!COM_CUENTA_AUTOM_D2
         If Trim(com_llave!COM_CUENTA_AUTOM_D2) <> "" Then
            ws_por = com_llave!COM_POR_AUTOM_D2
            ww_fff = 2
            GoSub graba_autom
         End If
         WS_CUENTA = com_llave!COM_CUENTA_AUTOM_D3
         If Trim(com_llave!COM_CUENTA_AUTOM_D3) <> "" Then
            ws_por = com_llave!COM_POR_AUTOM_D3
            ww_fff = 3
            GoSub graba_autom
         End If
         WS_CUENTA = com_llave!COM_CUENTA_AUTOM_D4
         If Trim(com_llave!COM_CUENTA_AUTOM_D4) <> "" Then
            ws_por = com_llave!COM_POR_AUTOM_D4
            ww_fff = 4
            GoSub graba_autom
         End If
         WS_CUENTA = com_llave!COM_CUENTA_AUTOM_D5
         If Trim(com_llave!COM_CUENTA_AUTOM_D5) <> "" Then
            ws_por = com_llave!COM_POR_AUTOM_D5
            ww_fff = 5
            GoSub graba_autom
         End If
         If Trim(com_llave!com_cuenta_AUTO_H) <> "" Then
            WS_CUENTA = com_llave!com_cuenta_AUTO_H
            WS_SUMA = 0
            ws_por = 100
            If cox_cuenta!COV_DH = "H" And Left(cox_cuenta!COV_CODCTA, 1) = "6" Then
               ws_dh = "D" ' Vws_dh = cox_cuenta!COV_DH
            Else
               ws_dh = "H"
            End If
            GoSub graba_autom
         End If
      End If
   CONTADOR = CONTADOR + 1
   DoEvents
'   barra.Value = CONTADOR
   cox_cuenta.MoveNext
Loop
 pb.Visible = False
 lpb.Caption = ""
 cop_llave.Requery
 cop_llave.Edit
 cop_llave!cop_FLAG_DES = "A"
 cop_llave.Update
 MsgBox "PROCESO TERMINADO...", 48, Pub_Titulo

Exit Sub

graba_autom:
            If ww_ff <> 0 Then
               If ws_por = 0 Then Return
            End If
            
            cov_voucher.AddNew
            cov_voucher!COV_FECHA_VOUCHER = ws_fecha_voucher 'ws_fecha_proc
            WS_NRO_MOV = WS_NRO_MOV + 1
            cov_voucher!COV_NRO_MOV = WS_NRO_MOV 'Nulo_Valor0(cox_cuenta!COV_NRO_MOV)
            cov_voucher!COV_CODCTA = WS_CUENTA
            cov_voucher!COV_NUMTAB = WS_NRO_MOV
            cov_voucher!COV_DH = ws_dh

            cov_voucher!COV_NRO_VOUCHER = ws_nro_voucher
            cov_voucher!COV_glosa = ws_glosa
            ws_parcial = Format(WS_IMPORTE * ws_por / 100, "0.00")
            If ww_ff = ww_fff Then ws_parcial = WS_IMPORTE - WS_SUMA
            If ws_parcial = 0 Then
               cov_voucher!cov_flag_automatica = "0"
            Else
               cov_voucher!cov_flag_automatica = "M"
            End If
                            
            WS_SUMA = WS_SUMA + ws_parcial
            cov_voucher!COV_IMPORTE = ws_parcial
            
            cov_voucher!COV_FECHA_doc = LK_FECHA_COP1
            cov_voucher!COV_CODUSU = ws_codusu
            cov_voucher!COV_CODCIA = LK_CODCIA
            cov_voucher!COV_ESTADO = WS_ESTADO
            cov_voucher!cov_nro_mes = LK_NRO_MES
            cov_voucher.Update

 Return
Exit Sub
fin:
End Sub


Public Sub PRO_DESTINOS()
Dim WS_NRO As Currency
Dim TEMPO_VAR
Dim PSCOV_CUENTA2 As rdoQuery
Dim cox_cuenta  As rdoResultset
Dim WS_CUENTA As String
Dim ws_parcial As Currency
Dim ws_dh_cov As String * 1
Dim tt_importe As Currency
Dim tempo_tipmov As Integer

Dim wcadena As String
Dim wvalor  As String * 1

Dim imp_total As Currency
Dim imp_des1 As Currency
Dim imp_des2 As Currency
Dim imp_des3 As Currency
Dim imp_des4 As Currency
Dim imp_des5 As Currency

' CHEQUE DE MES CERRADO 0 ABIERTO
WS_NRO = 0
wcadena = ""
SQ_OPER = 2
PUB_TIPREG = 155
PUB_CODCIA = LK_CODCIA
LEER_TAB_LLAVE
If tab_mayor.EOF Then
  MsgBox "Crear tab_tipreg = 155 para seguridas", 48, Pub_Titulo
Else
    Do Until tab_mayor.EOF
      If tab_mayor!TAB_NUMTAB = Val(Format(LK_FECHA_COP1, "yyyy")) Then
        wcadena = Trim(tab_mayor!tab_nomlargo)
      End If
      tab_mayor.MoveNext
    Loop
    wvalor = Mid(wcadena, LK_NRO_MES + 1, 1)
    If wvalor = "1" Then
       MsgBox "<<<< Mes CERRADO Operaciones >>>>", vbCritical, Pub_Titulo
       Exit Sub
    End If
End If

'cade = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND COV_FECHA_VOUCHER >= ? AND COV_FECHA_VOUCHER <= ?  AND COV_NRO_MES = " & LK_NRO_MES & " ORDER BY COV_CODCTA, COV_FECHA_VOUCHER, COV_NRO_MOV"
'Set PSCOV_CUENTA2 = CN.CreateQuery("", cade)
'PSCOV_CUENTA2(0) = 0
'PSCOV_CUENTA2(1) = LK_FECHA_DIA
'PSCOV_CUENTA2(2) = LK_FECHA_DIA
'Set cox_cuenta = PSCOV_CUENTA2.OpenResultset(rdOpenKeyset, rdConcurValues)

cade = "SELECT * FROM MOVICONT WHERE MOV_CODCIA = ? AND MOV_FECHA >= ? AND MOV_FECHA <= ?  AND MOV_NRO_MES = " & LK_NRO_MES & " ORDER BY MOV_TIPMOV, MOV_DH, MOV_CODCTA, MOV_FECHA, MOV_NRO_MOV"
Set PSCOV_CUENTA2 = CN.CreateQuery("", cade)
PSCOV_CUENTA2(0) = 0
PSCOV_CUENTA2(1) = LK_FECHA_DIA
PSCOV_CUENTA2(2) = LK_FECHA_DIA
Set cox_cuenta = PSCOV_CUENTA2.OpenResultset(rdOpenKeyset, rdConcurValues)

ffecha1 = Format(LK_FECHA_COP1, "yyyy/mm/dd")
ffecha2 = Format(LK_FECHA_COP2, "yyyy/mm/dd")
wpub_cadena = "DELETE MOVICONT  WHERE (MOV_FLAG_DES = 'A') AND MOV_CODCIA = '" & LK_CODCIA & "' AND MOV_FECHA >=  '" & ffecha1 & "'  AND MOV_FECHA <=  '" & ffecha2 & "'"
CN.Execute wpub_cadena, rdExecDirect

pub_mensaje = "Procesoar los Destinos ¿Desea Continuar... ?"
Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
If Pub_Respuesta = vbNo Then
    Exit Sub
End If

lpb.Caption = "Eliminando Movimientos Automaticos"
lpb.Caption = ""


PSCOV_CUENTA2.rdoParameters(0) = LK_CODCIA
PSCOV_CUENTA2.rdoParameters(1) = LK_FECHA_COP1
PSCOV_CUENTA2.rdoParameters(2) = LK_FECHA_COP2
cox_cuenta.Requery


PSTEMP_LLAVE(0) = LK_CODCIA
PSTEMP_LLAVE(1) = LK_FECHA_COP1
PSTEMP_LLAVE(2) = LK_FECHA_COP2
PSTEMP_LLAVE(3) = LK_NRO_MES

'barra.Max = cox_cuenta.RowCount
CONTADOR = 0
lpb.Caption = "Creando Movimientos Automaticos"
If Not cox_cuenta.EOF Then
 pb.Max = cox_cuenta.RowCount
End If
pb.Min = 0
pb.Value = 0
pb.Visible = True
tempo_tipmov = -1
ws_vou = -99 'cox_cuenta!MOV_nro_voucher
Do Until cox_cuenta.EOF
'     If cox_cuenta!MOV_TIPMOV = 3 Then Stop
      pb.Value = pb.Value + 1
      SQ_OPER = 1
      PUB_CUENTA = cox_cuenta!MOV_CODCTA
      PUB_CODCIA = LK_CODCIA
      LEER_COM_LLAVE
      If com_llave.EOF Then
         MsgBox "Revisar el diario ...Cuenta : " & PUB_CUENTA
         GoTo fin
      End If
      ww_ff = 0
      If com_llave!COM_POR_AUTOM_D <> 0 Then ww_ff = 1
      If com_llave!COM_POR_AUTOM_D2 <> 0 Then ww_ff = 2
      If com_llave!COM_POR_AUTOM_D3 <> 0 Then ww_ff = 3
      If com_llave!COM_POR_AUTOM_D4 <> 0 Then ww_ff = 4
      If com_llave!COM_POR_AUTOM_D5 <> 0 Then ww_ff = 5
      If Left(cox_cuenta!MOV_CODCTA, 1) = "9" And TEMPO_VAR = 1 Then
'        MsgBox ".."
      End If
      If Left(cox_cuenta!MOV_CODCTA, 1) = "9" Then
        TEMPO_VAR = 1
      End If
'      WS_NRO_MOV = 0
      
      WS_SUMA = Val(Nulo_Valors(com_llave!com_cuenta_AUTOM_D)) + Val(Nulo_Valors(com_llave!COM_CUENTA_AUTOM_D2)) + Val(Nulo_Valors(com_llave!COM_CUENTA_AUTOM_D3)) + Val(Nulo_Valors(com_llave!COM_CUENTA_AUTOM_D4)) + Val(Nulo_Valors(com_llave!COM_CUENTA_AUTOM_D5))
          
      If WS_SUMA <> 0 Then
         WS_SUMA = 0
         'ws_nro_voucher = cox_cuenta!COV_NRO_VOUCHER
         WS_NRO = cox_cuenta!MOV_NRO_VOUCHER
         ws_glosa = cox_cuenta!MOV_GLOSA
         ws_fecha_proc = cox_cuenta!MOV_FECHA
         wc_importe = cox_cuenta!mov_importe
         tt_importe = wc_importe
         ws_codusu = cox_cuenta!MOV_CODUSU
         
         If cox_cuenta!MOV_DH = "H" And (Left(cox_cuenta!MOV_CODCTA, 1) = "6" Or Left(cox_cuenta!MOV_CODCTA, 1) = "9") Then
            ws_dh = "H"
         Else
            ws_dh = "D"
         End If
         imp_des1 = 0
         imp_des2 = 0
         imp_des3 = 0
         imp_des4 = 0
         imp_des5 = 0
         WS_CUENTA = com_llave!com_cuenta_AUTOM_D
         ws_dh_cov = cox_cuenta!MOV_DH
         If Trim(com_llave!com_cuenta_AUTOM_D) <> "" Then
            ws_por = com_llave!COM_POR_AUTOM_D
            imp_des1 = Format(wc_importe * ws_por / 100, "0.00")
         End If

         If Trim(com_llave!COM_CUENTA_AUTOM_D2) <> "" Then
            ws_por = com_llave!COM_POR_AUTOM_D2
            imp_des2 = Format(wc_importe * ws_por / 100, "0.00")
         End If
         If Trim(com_llave!COM_CUENTA_AUTOM_D3) <> "" Then
            ws_por = com_llave!COM_POR_AUTOM_D3
            imp_des3 = Format(wc_importe * ws_por / 100, "0.00")
         End If
         If Trim(com_llave!COM_CUENTA_AUTOM_D4) <> "" Then
            ws_por = com_llave!COM_POR_AUTOM_D4
            imp_des4 = Format(wc_importe * ws_por / 100, "0.00")
         End If
         If Trim(com_llave!COM_CUENTA_AUTOM_D5) <> "" Then
            ws_por = com_llave!COM_POR_AUTOM_D5
            imp_des5 = Format(wc_importe * ws_por / 100, "0.00")
         End If
         
         imp_total = imp_des1 + imp_des2 + imp_des3 + imp_des4 + imp_des5
         If Val(wc_importe) <> Val(imp_total) Then
          imp_des1 = Val(imp_des1) + (Val(wc_importe) - Val(imp_total))
         End If
         
         WS_CUENTA = com_llave!com_cuenta_AUTOM_D
         If Trim(com_llave!com_cuenta_AUTOM_D) <> "" Then
            ww_fff = 1
            ws_parcial = imp_des1
            GoSub graba_autom
         End If
         WS_CUENTA = com_llave!COM_CUENTA_AUTOM_D2
         If Trim(com_llave!COM_CUENTA_AUTOM_D2) <> "" Then
            ww_fff = 2
            ws_parcial = imp_des2
            GoSub graba_autom
         End If
         WS_CUENTA = com_llave!COM_CUENTA_AUTOM_D3
         If Trim(com_llave!COM_CUENTA_AUTOM_D3) <> "" Then
            ww_fff = 3
            ws_parcial = imp_des3
            GoSub graba_autom
         End If
         WS_CUENTA = com_llave!COM_CUENTA_AUTOM_D4
         If Trim(com_llave!COM_CUENTA_AUTOM_D4) <> "" Then
            ww_fff = 4
            ws_parcial = imp_des4
            GoSub graba_autom
         End If
         WS_CUENTA = com_llave!COM_CUENTA_AUTOM_D5
         If Trim(com_llave!COM_CUENTA_AUTOM_D5) <> "" Then
            ww_fff = 5
            ws_parcial = imp_des5
            GoSub graba_autom
         End If
         
         
         If Trim(com_llave!com_cuenta_AUTO_H) <> "" Then
            WS_CUENTA = com_llave!com_cuenta_AUTO_H
            WS_SUMA = 0
            ws_por = 100
            ws_parcial = wc_importe
            If cox_cuenta!MOV_DH = "H" And (Left(cox_cuenta!MOV_CODCTA, 1) = "6" Or Left(cox_cuenta!MOV_CODCTA, 1) = "9") Then
               ws_dh = "D" ' Vws_dh = cox_cuenta!COV_DH
            Else
               ws_dh = "H"
            End If
            GoSub graba_autom
         End If
      End If
   CONTADOR = CONTADOR + 1
   DoEvents
'   barra.Value = CONTADOR
   cox_cuenta.MoveNext
Loop
 pb.Visible = False
 lpb.Caption = ""
 cop_llave.Requery
 cop_llave.Edit
 cop_llave!cop_FLAG_DES = "A"
 cop_llave.Update
 ACT_MESES (0)
 MsgBox "PROCESO TERMINADO...", 48, Pub_Titulo
 
 
 

Exit Sub

graba_autom:
           If ww_ff <> 0 Then
               If ws_por = 0 Then Return
           End If
           'ws_parcial = Format(wc_importe * ws_por / 100, "0.00")
           If ws_parcial = 0 Then Return
            
           If tempo_tipmov <> cox_cuenta!MOV_TIPMOV Then
             PSTEMP_LLAVE(4) = cox_cuenta!MOV_TIPMOV
             temp_llave.Requery
             If temp_llave.EOF Then
                ws_nro_voucher = 0
                ws_fecha_voucher = LK_FECHA_COP1
             Else
           '     temp_llave.MoveLast
                ws_nro_voucher = temp_llave!MOV_NRO_VOUCHER
                ws_fecha_voucher = temp_llave!MOV_FECHA
            End If
            tempo_tipmov = cox_cuenta!MOV_TIPMOV
            WS_NRO_MOV = 0
            ws_nro_voucher = ws_nro_voucher + 1
            ws_vou = cox_cuenta!MOV_NRO_VOUCHER
          Else
           If cox_cuenta!MOV_NRO_VOUCHER <> ws_vou Then
                ws_nro_voucher = ws_nro_voucher + 1
                ws_vou = cox_cuenta!MOV_NRO_VOUCHER
            End If
          End If
          
      
            
            
            temp_llave.AddNew
            temp_llave!MOV_NRO_VOUCHER = Val(WS_NRO) + 0.1 ' ws_nro_voucher
            temp_llave!MOV_FECHA = ws_fecha_voucher 'ws_fecha_proc
            WS_NRO_MOV = WS_NRO_MOV + 1
            temp_llave!MOV_NRO_MOV = WS_NRO_MOV
            temp_llave!MOV_TIPMOV = cox_cuenta!MOV_TIPMOV
            If tempo_tipmov = 1 Then
                temp_llave!MOV_GLOSA = "Destino por las Compras "
            ElseIf tempo_tipmov = 2 Then
                temp_llave!MOV_GLOSA = "Destino por las Ventas "
            ElseIf tempo_tipmov = 3 Then
                If cox_cuenta!MOV_PLANTILLA = 100 Then
                  temp_llave!MOV_GLOSA = "Destino por los Ingresos de Caja"
                Else
                    temp_llave!MOV_GLOSA = "Destino por los Egresos de Caja"
                End If
            Else
                temp_llave!MOV_GLOSA = "Destino por " & Trim(ws_glosa)
            End If
            temp_llave!MOV_MONEDA = "S"
            temp_llave!MOV_SUNAT = "00"
            temp_llave!MOV_serie = 0
            temp_llave!MOV_numfac = 0
            temp_llave!MOV_codclie = 0
            temp_llave!MOV_CP = ""
            temp_llave!MOV_FBG = " "
            temp_llave!MOV_MARCA = ""
            temp_llave!MOV_fecha_EMI = ws_fecha_voucher 'ws_fecha_proc
            temp_llave!MOV_serie_c = 0
            temp_llave!MOV_numfac_c = 0
            temp_llave!MOV_FBG_C = " "
            temp_llave!MOV_PLANTILLA = cox_cuenta!MOV_PLANTILLA
            temp_llave!MOV_FLAG_TC = " "
            temp_llave!MOV_TIPO_CAMBIO = 0
            temp_llave!MOV_CODUSU = LK_CODUSU
            temp_llave!MOV_CODCTA = WS_CUENTA
             'If Trim(WS_CUENTA) = "" Then Stop
            temp_llave!MOV_DH = ws_dh
            
            temp_llave!MOV_DETALLE = Trim(temp_llave!MOV_GLOSA) ' "Destino Cta.: " & Trim(com_llave!com_cuenta) & " " & ws_glosa
            temp_llave!MOV_FLAG_DES = "A"
            WS_SUMA = WS_SUMA + ws_parcial
            temp_llave!mov_importe = ws_parcial
            temp_llave!MOV_CODUSU = ws_codusu
            temp_llave!MOV_CODCIA = LK_CODCIA
            temp_llave!MOV_nro_MES = LK_NRO_MES
            temp_llave.Update
   TEMPO_VAR = 0
 Return
Exit Sub
fin:


End Sub
