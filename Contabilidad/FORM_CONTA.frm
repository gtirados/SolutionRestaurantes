VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Begin VB.Form FORM_CONTA 
   BackColor       =   &H00FAEFDA&
   Caption         =   "Ingreso de Vouchers"
   ClientHeight    =   4890
   ClientLeft      =   1500
   ClientTop       =   1140
   ClientWidth     =   6600
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "FORM_CONTA.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   Tag             =   "55"
   WindowState     =   2  'Maximized
   Begin RichTextLib.RichTextBox TEXTOVAR 
      Height          =   375
      Left            =   1110
      TabIndex        =   31
      Top             =   2670
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393217
      BackColor       =   16445402
      MultiLine       =   0   'False
      Appearance      =   0
      TextRTF         =   $"FORM_CONTA.frx":0442
   End
   Begin MSFlexGridLib.MSFlexGrid grid1 
      Height          =   4230
      Left            =   750
      TabIndex        =   32
      Tag             =   "9999"
      Top             =   1755
      Visible         =   0   'False
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   7461
      _Version        =   393216
      Rows            =   3
      FixedRows       =   2
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      GridColorFixed  =   16445402
      FocusRect       =   2
      HighLight       =   2
      GridLines       =   2
      GridLinesFixed  =   1
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin MSFlexGridLib.MSFlexGrid grid_fac 
      Height          =   4215
      Left            =   750
      TabIndex        =   33
      Tag             =   "9999"
      Top             =   1755
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   7435
      _Version        =   393216
      Rows            =   3
      FixedRows       =   2
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      GridColorFixed  =   16445402
      FocusRect       =   2
      HighLight       =   2
      GridLines       =   2
      GridLinesFixed  =   1
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FAEFDA&
      Height          =   810
      Left            =   210
      TabIndex        =   25
      Top             =   6075
      Width           =   11430
      Begin VB.CommandButton cancelar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   4245
         Style           =   1  'Graphical
         TabIndex        =   29
         TabStop         =   0   'False
         Tag             =   "9999"
         Top             =   225
         Width           =   1290
      End
      Begin VB.CommandButton SALIR 
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
         Height          =   420
         Left            =   9765
         Style           =   1  'Graphical
         TabIndex        =   28
         TabStop         =   0   'False
         Tag             =   "9999"
         Top             =   225
         Width           =   1290
      End
      Begin VB.CommandButton cmdIngreso 
         Caption         =   "&Ingreso"
         Height          =   420
         Left            =   585
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   225
         Width           =   1290
      End
      Begin VB.CommandButton cmdConsultar 
         Appearance      =   0  'Flat
         Caption         =   "&Mostrar"
         Height          =   420
         Left            =   2415
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   225
         Width           =   1290
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FAEFDA&
      Height          =   1110
      Left            =   210
      TabIndex        =   14
      Top             =   15
      Width           =   11430
      Begin VB.CheckBox mayor 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Ver como Libro Mayor"
         Height          =   255
         Left            =   4605
         TabIndex        =   23
         Top             =   660
         Width           =   2295
      End
      Begin VB.ComboBox i_fecha2 
         BackColor       =   &H00FFFFFF&
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
         Height          =   315
         Left            =   1395
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   210
         Width           =   2295
      End
      Begin VB.TextBox i_tipdoc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1410
         TabIndex        =   17
         Text            =   "i_tipdoc"
         Top             =   660
         Width           =   2280
      End
      Begin VB.TextBox i_cuenta 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7620
         MaxLength       =   10
         TabIndex        =   16
         TabStop         =   0   'False
         Tag             =   "9999"
         Top             =   210
         Width           =   1905
      End
      Begin VB.TextBox i_voucher 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4635
         MaxLength       =   10
         TabIndex        =   15
         TabStop         =   0   'False
         Tag             =   "9999"
         Top             =   210
         Width           =   1095
      End
      Begin VB.Label lblperiodo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PERIODO : ASIENTOS DE AJUSTES"
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
         Left            =   8025
         TabIndex        =   24
         Top             =   720
         Width           =   2745
      End
      Begin VB.Label L1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo Doc:"
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
         Left            =   495
         TabIndex        =   22
         Top             =   705
         Width           =   765
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
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
         Left            =   495
         TabIndex        =   21
         Top             =   210
         Width           =   570
      End
      Begin VB.Label LCODART 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cta. para Filtrar:"
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
         Left            =   6045
         TabIndex        =   20
         Tag             =   "9999"
         Top             =   210
         Width           =   1365
      End
      Begin VB.Label LCODART 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voucher"
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
         Index           =   3
         Left            =   3825
         TabIndex        =   19
         Tag             =   "9999"
         Top             =   210
         Width           =   690
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FAEFDA&
      Height          =   615
      Left            =   600
      TabIndex        =   8
      Top             =   7425
      Visible         =   0   'False
      Width           =   7935
      Begin VB.ComboBox i_d_h 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "FORM_CONTA.frx":04C3
         Left            =   3240
         List            =   "FORM_CONTA.frx":04CD
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   120
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.TextBox i_importe 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6120
         TabIndex        =   9
         TabStop         =   0   'False
         Tag             =   "9999"
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label LCODART 
         BackStyle       =   0  'Transparent
         Caption         =   "Importe:"
         Height          =   255
         Index           =   1
         Left            =   5280
         TabIndex        =   12
         Tag             =   "9999"
         Top             =   240
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label LCODART 
         BackStyle       =   0  'Transparent
         Caption         =   "D/H"
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   11
         Tag             =   "9999"
         Top             =   240
         Visible         =   0   'False
         Width           =   405
      End
   End
   Begin ComctlLib.ProgressBar Barra 
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Tag             =   "0"
      Top             =   4080
      Visible         =   0   'False
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   0
      Min             =   77
      Max             =   91
   End
   Begin VB.Frame ESTADO 
      BackColor       =   &H00FAEFDA&
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
      Height          =   4950
      Left            =   210
      TabIndex        =   1
      Tag             =   "100"
      Top             =   1110
      Width           =   11430
      Begin VB.CommandButton cmdcambiar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cam&biar Glosa"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9540
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Tag             =   "9999"
         Top             =   240
         Width           =   1350
      End
      Begin ComctlLib.ProgressBar PB 
         Height          =   255
         Left            =   2040
         TabIndex        =   6
         Top             =   1560
         Visible         =   0   'False
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   450
         _Version        =   327682
         Appearance      =   0
      End
      Begin VB.TextBox i_glosa 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1185
         TabIndex        =   0
         TabStop         =   0   'False
         Tag             =   "9999"
         Text            =   "i_glosa"
         Top             =   240
         Width           =   8025
      End
      Begin VB.CommandButton cmdcorta 
         Caption         =   "Detener"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6240
         TabIndex        =   7
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Lcuenta 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   4560
         Width           =   3015
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Glosa:"
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
         Left            =   585
         TabIndex        =   2
         Tag             =   "9999"
         Top             =   270
         Width           =   510
      End
      Begin VB.Label momen 
         Caption         =   "Un Momento ..."
         Height          =   375
         Left            =   3360
         TabIndex        =   4
         Top             =   1560
         Width           =   1575
      End
   End
   Begin VB.Label lblbarraos 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Solution - Gestión Contable"
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
      Height          =   330
      Left            =   15
      TabIndex        =   30
      Top             =   6915
      Width           =   11955
   End
End
Attribute VB_Name = "FORM_CONTA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WPASA As Boolean
Dim WSELE As String * 1
Dim llave1
Dim fila As Integer
Dim ws_bruto_d, ws_bruto_h As Currency
Dim SUM_D As Currency
Dim SUM_H As Currency
Dim PSTEMP_LLAVE As rdoQuery
Dim temp_llave As rdoResultset
Dim WMODO As String * 1
Dim LOC_ITEM As Integer
Dim cop_llave As rdoResultset
Dim PSCOP_LLAVE As rdoQuery
Dim LOC_CANCELA As Integer
Dim PSTEMP_MAYOR As rdoQuery
Dim temp_mayor As rdoResultset

Dim temporal
Dim wfila_act As Integer
Dim loc_ini As Integer
Dim loc_fin  As Integer
Dim Wsec As Integer
Option Explicit
Private Sub cancelar_Click()
grid1.Visible = False
mayor.Enabled = False
i_voucher.Text = ""
WMODO = ""
i_voucher.Enabled = False
i_voucher.BackColor = QBColor(7)

cmdIngreso.Caption = "&Ingreso"
cmdcambiar.Visible = False
cmdConsultar.Caption = "&Mostrar"
ESTADO.Caption = "Estado : "
fila = 0
SUM_D = 0
SUM_H = 0
LIMPIA_DATOS
CABE_ING
FORM_CONTA.lcuenta.Caption = ""
grid_fac.SetFocus
SendKeys "^{HOME}", True
WMODO = ""
End Sub

Private Sub cmdcambiar_Click()
On Error GoTo SALE
Dim WF As Integer
Dim WFIN As Integer
Dim WINI As Integer
Dim fx As Integer
If WMODO = "M" Then Exit Sub
If Val(grid_fac.TextMatrix(grid_fac.Row, 3)) = 0 Then
   MsgBox "Seleccione un voucher de ll lista .", 48, Pub_Titulo
   grid_fac.SetFocus
   Exit Sub
End If
'pub_mensaje = "Desea cambiar Glosa al Voucher Nro. : " & grid_fac.TextMatrix(grid_fac.Row, 3)
'Pub_Respuesta = MsgBox(pub_mensaje, Pub_Estilo, Pub_Titulo)
'If Pub_Respuesta = vbNo Then
'   grid_fac.SetFocus
'   Exit Sub
'End If
fx = grid_fac.Row
fx = grid_fac.Row
WF = 0
fila = 0
Do Until WF = 1
   fx = fx - 1
   If Left(grid_fac.TextMatrix(fx, 0), 1) = "T" Or fx = 1 Then
     WINI = fx + 1
     WF = 1
   End If
Loop
fx = grid_fac.Row
WF = 0
Do Until WF = 1
   fx = fx + 1
   If Left(grid_fac.TextMatrix(fx, 0), 1) = "T" Or fx = 1 Then
     WFIN = fx - 1
     WF = 1
   End If
Loop
loc_ini = WINI
loc_fin = WFIN
i_glosa.Enabled = True
i_glosa.Locked = False
Azul i_glosa, i_glosa
Exit Sub
SALE:
MsgBox "Verficar Importe.", 48, Pub_Titulo
If TEXTOVAR.Visible Then Azul3 TEXTOVAR, TEXTOVAR

End Sub

Private Sub cmdConsultar_Click()
Dim flag_grabar As String * 1
Dim w_dh As String
Dim IMPORTE_DEB As Currency
Dim IMPORTE_HAB As Currency
Dim Wflag  As String * 1
If Left(cmdConsultar.Caption, 2) = "&G" Then
  flag_grabar = ""
  Wflag = ""
  For fila = 2 To grid_fac.Rows - 1
    If Trim(grid_fac.TextMatrix(fila, 9)) = "9" Or Trim(grid_fac.TextMatrix(fila, 9)) = "-1" Or Trim(grid_fac.TextMatrix(fila, 11)) = "8" Then
    Else
      GoTo pasa
    End If
    If Trim(grid_fac.TextMatrix(fila, 1)) = "" Then GoTo pasa
    
    PUB_CUENTA = Trim(grid_fac.TextMatrix(fila, 1))
    If Val(grid_fac.TextMatrix(fila, 4)) <> 0 Then
         w_dh = "D"
         PUB_IMPORTE = Val(grid_fac.TextMatrix(fila, 4))
    ElseIf Val(grid_fac.TextMatrix(fila, 5)) <> 0 Then
         w_dh = "H"
         PUB_IMPORTE = Val(grid_fac.TextMatrix(fila, 5))
    End If
    ' grabo todo
    If Trim(grid_fac.TextMatrix(fila, 9)) = "-1" Then
        PSCOV_LLAVE(0) = LK_CODCIA
        PSCOV_LLAVE(1) = grid_fac.TextMatrix(fila, 10)
        PSCOV_LLAVE(2) = Val(grid_fac.TextMatrix(fila, 3))
        PSCOV_LLAVE(3) = Val(grid_fac.TextMatrix(fila, 6))
        cov_llave.Requery
        If cov_llave.EOF Then
          flag_grabar = "B"
        Else
          cov_llave.Delete
          flag_grabar = "B"
        End If
    ElseIf Trim(grid_fac.TextMatrix(fila, 11)) = "8" Then
        Wflag = "A"
        cov_llave.AddNew
        cov_llave!COV_CODCTA = Trim(grid_fac.TextMatrix(fila, 1))
        cov_llave!COV_NRO_MOV = grid_fac.TextMatrix(fila, 6)
        cov_llave!COV_CODCIA = LK_CODCIA
        cov_llave!COV_NRO_VOUCHER = grid_fac.TextMatrix(fila, 3)
        cov_llave!COV_FECHA_VOUCHER = grid_fac.TextMatrix(fila, 0) 'LK_FECHA_COP2
        cov_llave!COV_glosa = grid_fac.TextMatrix(fila, 7)
        cov_llave!COV_FECHA_doc = LK_FECHA_DIA
        If Val(grid_fac.TextMatrix(fila, 4)) <> 0 Then
          cov_llave!COV_DH = "D"
          cov_llave!COV_IMPORTE = Val(grid_fac.TextMatrix(fila, 4))
        Else
          cov_llave!COV_DH = "H"
          cov_llave!COV_IMPORTE = Val(grid_fac.TextMatrix(fila, 5))
        End If
        cov_llave!COV_ESTADO = "0"
        cov_llave!COV_CODUSU = LK_CODUSU
        cov_llave!cov_flag_automatica = grid_fac.TextMatrix(fila, 8)
        cov_llave!cov_nro_mes = LK_NRO_MES
        cov_llave.Update
    Else
        PSCOV_LLAVE(0) = LK_CODCIA
        PSCOV_LLAVE(1) = grid_fac.TextMatrix(fila, 10)
        PSCOV_LLAVE(2) = Val(grid_fac.TextMatrix(fila, 3))
        PSCOV_LLAVE(3) = Val(grid_fac.TextMatrix(fila, 6))
        cov_llave.Requery
        If cov_llave.EOF Then
           MsgBox "Verificar Voucher.", 48, Pub_Titulo
           GoTo fin
        End If
       cov_llave.Edit
       cov_llave!COV_CODCTA = PUB_CUENTA
       cov_llave!COV_FECHA_VOUCHER = grid_fac.TextMatrix(fila, 0) 'LK_FECHA_COP2
       cov_llave!COV_DH = w_dh
       cov_llave!COV_IMPORTE = PUB_IMPORTE
       cov_llave!COV_glosa = Trim(grid_fac.TextMatrix(fila, 7))
       cov_llave!COV_CODUSU = LK_CODUSU
       cov_llave.Update
       flag_grabar = "A"
    End If
    
pasa:
Next fila
If Wflag = "A" Then
  cancelar_Click
  MsgBox "Diario Actualizado. debe Mayorizar Nuevamente.", 48, Pub_Titulo
  cmdConsultar.Caption = "&Mostrar"
  GoTo listo
End If

If flag_grabar = "A" Then
  cop_llave.Requery
  cop_llave.Edit
  cop_llave!COP_FLAG_MAYORIZACION = " "
  cop_llave.Update
  MsgBox "Diario Actualizado. debe Mayorizar Nuevamente.", 48, Pub_Titulo
ElseIf flag_grabar = "B" Then
  cancelar_Click
  MsgBox "Diario Actualizado. debe Mayorizar Nuevamente.", 48, Pub_Titulo
  cmdConsultar.Caption = "&Mostrar"
End If
listo:
If Trim(flag_grabar) = "" And Wflag <> "A" Then
  MsgBox "Existen un dato incorrecto.  Verificar", 48, Pub_Titulo
End If
Exit Sub
End If
PSTEMP_MAYOR(0) = LK_CODCIA
PSTEMP_MAYOR(1) = LK_FECHA_COP1
PSTEMP_MAYOR(2) = LK_FECHA_COP2
temp_mayor.Requery
If temp_mayor.EOF Then
   Wsec = 0
Else
   temp_mayor.MoveLast
   Wsec = temp_mayor!COV_NRO_MOV
End If

ESTADO.Caption = "Estado :   < CONSULTA, MODIFICA, ELIMINA >"
cmdConsultar.Caption = "&Grabar"
WMODO = "C"
CABE_MAN
cmdcambiar.Visible = True
i_cuenta.Enabled = True
i_cuenta.BackColor = QBColor(15)
i_cuenta.Text = ""
i_glosa.Text = ""
i_importe.Text = ""
i_d_h.ListIndex = -1
i_tipdoc.Locked = True
i_tipdoc.BackColor = QBColor(7)
i_glosa.Locked = True
i_voucher.Locked = False
i_voucher.Enabled = True
i_voucher.BackColor = QBColor(15)

FORM_CONTA.i_d_h.Enabled = True
FORM_CONTA.i_tipdoc.Enabled = False
FORM_CONTA.i_cuenta.Enabled = True
FORM_CONTA.i_importe.Enabled = True
mayor.Enabled = True

cmdIngreso.Enabled = False
cmdConsultar.Enabled = True
WPASA = False
i_fecha2.SetFocus
fin:
End Sub

Private Sub cmdcorta_Click()
LOC_CANCELA = 2
End Sub

Private Sub cmdIngreso_Click()
Dim ws_tot_debe, ws_tot_haber As Currency
Dim er As rdoError
Dim pub_mensaje As String
Const ingre = 2
Const modif = 1
Dim N As Integer
Dim LOC_SALDO_CAR As Currency
Dim FLAG As Boolean
Dim pub_mensaje_err As String
Dim WS_NRO_MOV, ws_nro_voucher As Long
Dim w_dh  As String

If Left(cmdIngreso.Caption, 2) = "&G" Then

FLAG = False
ws_tot_debe = Val(grid_fac.TextMatrix(1, 4))
ws_tot_haber = Val(grid_fac.TextMatrix(1, 5))

If ws_tot_debe = 0 And ws_tot_haber = 0 Then
  MsgBox "Ingrese los Vouchers ..", 48, Pub_Titulo
  If grid_fac.Rows > 2 Then
     grid_fac.Col = 1
     grid_fac.Row = 2
     grid_fac.SetFocus
   End If
  Barra.Visible = False
  GoTo fin
End If
For fila = 2 To grid_fac.Rows - 1
If grid_fac.TextMatrix(fila, 1) <> "" Then
  If Val(grid_fac.TextMatrix(fila, 4)) = 0 And Val(grid_fac.TextMatrix(fila, 5)) = 0 Then
    MsgBox "Verificar hay alguna Cuenta sin Importe.", 48, Pub_Titulo
    grid_fac.SetFocus
    GoTo fin
  End If
End If
Next fila
If ws_tot_debe <> ws_tot_haber Then
   pub_mensaje = MsgBox("Voucher no cuadra Desea Registrar ... ?", 36)
   If pub_mensaje = vbNo Then GoTo fin
End If

Screen.MousePointer = 11
DoEvents
Barra.Visible = True
DoEvents
Barra.Min = 0
Barra.Max = fila
Barra.Value = 0

exito = True
Barra.Value = 1

GoSub ACT1
fila = 1
SUM_D = 0
SUM_H = 0

FORM_CONTA.i_d_h.ListIndex = -1
FORM_CONTA.i_tipdoc = ""
FORM_CONTA.grid_fac.Clear
fila = 0
FORM_CONTA.i_cuenta = ""
FORM_CONTA.i_glosa = ""
FORM_CONTA.i_importe = ""
WPASA = True
CABE_ING
SendKeys "^{HOME}", True
cancelar.SetFocus
If exito = False Then
   Barra.Visible = False
   MsgBox pub_mensaje_err
    GoTo errorr
End If
i_importe.Text = ""
Barra.Visible = False
cmdIngreso.Caption = "&Ingreso"
cmdConsultar.Enabled = True
GoTo fin

ACT1:

PSTEMP_LLAVE(0) = LK_CODCIA
PSTEMP_LLAVE(1) = LK_FECHA_COP1
PSTEMP_LLAVE(2) = #1/1/2020#
temp_llave.Requery
If temp_llave.EOF Then
 ws_nro_voucher = 0
Else
 temp_llave.MoveLast
 ws_nro_voucher = temp_llave!COV_NRO_VOUCHER
End If
fila = 1
ws_nro_voucher = ws_nro_voucher + 1
PUB_FECHA = i_fecha2.Text
PUB_CONCEPTO = Trim(i_glosa.Text)
FLAG = False
WS_NRO_MOV = 0
fila = 2
Do While FLAG = False
   If Trim(grid_fac.TextMatrix(fila, 1)) = "" Then GoTo pasa
   PUB_FECHA = i_fecha2.Text
   PUB_CUENTA = Trim(grid_fac.TextMatrix(fila, 1))
   If Trim(grid_fac.TextMatrix(fila, 4)) <> "" Then
        w_dh = "D"
        PUB_IMPORTE = grid_fac.TextMatrix(fila, 4)
   ElseIf Trim(grid_fac.TextMatrix(fila, 5)) <> "" Then
        w_dh = "H"
        PUB_IMPORTE = grid_fac.TextMatrix(fila, 5)
   End If
SIGUE_MAS:
    ' grabo todo
   cov_llave.AddNew
   cov_llave!COV_NRO_MOV = WS_NRO_MOV
   cov_llave!COV_NUMTAB = 0
   cov_llave!COV_CODCIA = LK_CODCIA
   cov_llave!COV_NRO_VOUCHER = ws_nro_voucher
   cov_llave!COV_FECHA_VOUCHER = PUB_FECHA
   cov_llave!COV_glosa = PUB_CONCEPTO
   cov_llave!COV_FECHA_doc = LK_FECHA_DIA
   cov_llave!COV_CODCTA = PUB_CUENTA
   cov_llave!COV_DH = w_dh
   cov_llave!COV_IMPORTE = PUB_IMPORTE
   cov_llave!COV_ESTADO = "0"
   cov_llave!COV_CODUSU = LK_CODUSU
   cov_llave!cov_flag_automatica = "0"
   cov_llave!cov_nro_mes = LK_NRO_MES
   cov_llave.Update
pasa:
   fila = fila + 1
   WS_NRO_MOV = WS_NRO_MOV + 1
   If fila >= FORM_CONTA.grid_fac.Rows Then
      FLAG = True
   End If
  
Loop
cop_llave.Requery
cop_llave.Edit
cop_llave!COP_FLAG_MAYORIZACION = " "
cop_llave.Update
Return

Screen.MousePointer = 1


Exit Sub
End If
cmdIngreso.Caption = "&Grabar"
i_cuenta.Enabled = False
i_cuenta.BackColor = QBColor(7)
ESTADO.Caption = "Estado :   < Ingreso de VOUCHER >"
WMODO = "I"
i_tipdoc.Locked = False
i_tipdoc.BackColor = QBColor(15)
FORM_CONTA.i_d_h.Enabled = True
FORM_CONTA.i_tipdoc.Enabled = True
FORM_CONTA.i_cuenta.Enabled = True
FORM_CONTA.i_glosa.Enabled = True
FORM_CONTA.i_importe.Enabled = True

cmdConsultar.Enabled = False
WPASA = True
CABE_ING


fila = 0
SUM_D = 0
SUM_H = 0
i_fecha2.SetFocus
Exit Sub

Error_fatal:
    pub_mensaje = "Se ha producido un error " & "al abrir la conexión:" & Err & " - " & Error & vbCr
    For Each er In rdoErrors
        pub_mensaje = pub_mensaje & er.Description & ":" & er.Number & vbCr
        MsgBox pub_mensaje
    Next er
    CN.Execute "Rollback Transaction", rdExecDirect
'    Resume AbandonCn
Exit Sub

errorr:
 MsgBox pub_mensaje_err, 48, Pub_Titulo
fin:
Screen.MousePointer = 0
Exit Sub
SALE:
If Err.Number = 6 Then
  MsgBox "Verficar Importe.", 48, Pub_Titulo
  If TEXTOVAR.Visible Then Azul3 TEXTOVAR, TEXTOVAR
  FORM_CONTA.Barra.Visible = False
  Screen.MousePointer = 0
  grid_fac.SetFocus
Else
  MsgBox Err.Description, 48, Pub_Titulo
End If

End Sub



Private Sub Form_Load()
'On Error GoTo SALE

Wsec = 0
LOC_CANCELA = 0
fila = 0
wfila_act = 0
WSELE = ""
Dim ws_indice As Integer
Dim cade
WMODO = ""
pub_cadena = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND COV_FECHA_VOUCHER >= ? and COV_FECHA_VOUCHER <= ? AND COV_NRO_MES = " & LK_NRO_MES & " ORDER BY  COV_NRO_MOV"
Set PSTEMP_MAYOR = CN.CreateQuery("", pub_cadena)
PSTEMP_MAYOR(0) = 0
PSTEMP_MAYOR(1) = LK_FECHA_DIA
PSTEMP_MAYOR(2) = LK_FECHA_DIA
Set temp_mayor = PSTEMP_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND COV_FECHA_VOUCHER>=? AND COV_FECHA_VOUCHER <=?  AND COV_NRO_MES = " & LK_NRO_MES & " ORDER BY COV_NRO_VOUCHER, COV_NRO_MOV"
Set PSTEMP_LLAVE = CN.CreateQuery("", pub_cadena)
PSTEMP_LLAVE(0) = 0
PSTEMP_LLAVE(1) = LK_FECHA_DIA
PSTEMP_LLAVE(2) = LK_FECHA_DIA
Set temp_llave = PSTEMP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND COV_FECHA_VOUCHER >=? AND COV_FECHA_VOUCHER <=?  AND COV_NRO_MES = " & LK_NRO_MES & "  ORDER BY COV_CODCIA, COV_FECHA_VOUCHER , COV_NRO_VOUCHER, COV_DH" ', COV_NRO_MOV"
Set PSCOV_MAYOR = CN.CreateQuery("", pub_cadena)
PSCOV_MAYOR(0) = 0
PSCOV_MAYOR(1) = LK_FECHA_DIA
PSCOV_MAYOR(2) = LK_FECHA_DIA
Set cov_mayor = PSCOV_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM COMOV WHERE COV_CODCIA = ?  AND COV_FECHA_VOUCHER=? AND COV_NRO_VOUCHER = ? AND COV_NRO_MOV = ? AND COV_NRO_MES = " & LK_NRO_MES & " ORDER BY COV_CODCIA, COV_FECHA_VOUCHER, COV_NRO_VOUCHER, COV_NRO_MOV"
Set PSCOV_LLAVE = CN.CreateQuery("", pub_cadena)
PSCOV_LLAVE(0) = 0
PSCOV_LLAVE(1) = 0
PSCOV_LLAVE(2) = 0
PSCOV_LLAVE(3) = 0
Set cov_llave = PSCOV_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM COPARAM WHERE COP_CODCIA = ?"
Set PSCOP_LLAVE = CN.CreateQuery("", pub_cadena)
PSCOP_LLAVE(0) = 0
Set cop_llave = PSCOP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

PSCOP_LLAVE.rdoParameters(0) = LK_CODCIA
cop_llave.Requery
If cop_llave.EOF Then
 Screen.MousePointer = 0
 MsgBox "Hay que definir parametros de contabilidad.", 48, Pub_Titulo
 Unload FORM_CONTA
 Exit Sub
End If
If DatePart("m", LK_FECHA_COP1) = 1 And DatePart("d", LK_FECHA_COP1) = 1 And DatePart("m", LK_FECHA_COP2) = 1 And DatePart("d", LK_FECHA_COP2) = 1 Then
   lblperiodo.Caption = "PERIODO : ASIENTOS DE AJUSTES"
Else
   lblperiodo.Caption = "PERIODO : " & Format(LK_FECHA_COP1, "dd/mm/yyyy") & " al " & Format(LK_FECHA_COP2, "dd/mm/yyyy")
End If


fila = 0
DoEvents
LIMPIA_DATOS
Dim ws_fecha As Date
i_fecha2.Clear
ws_fecha = LK_FECHA_COP1
Do Until fila = 999
   i_fecha2.AddItem Format(ws_fecha, "dd/mmm/YYYY")
   ws_fecha = DateAdd("d", 1, ws_fecha)
   If ws_fecha > LK_FECHA_COP2 Then fila = 999
Loop
i_fecha2.ListIndex = 0
fila = 0
i_tipdoc.Enabled = False
i_tipdoc.BackColor = QBColor(7)
i_voucher.Enabled = False
i_voucher.BackColor = QBColor(7)
mayor.Enabled = False
CABE_MAN

Exit Sub
SALE:
MsgBox "Depurar: " & Err.Description, 48, Pub_Titulo
Resume Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
If LOC_CANCELA = 1 Then
  cmdcorta_Click
  Exit Sub
End If
End Sub

Private Sub grid_fac_EnterCell()
If i_glosa.Locked = True Then i_glosa.Text = grid_fac.TextMatrix(grid_fac.Row, 7)
End Sub

Private Sub grid_fac_GotFocus()
If Val(grid_fac.TextMatrix(grid_fac.Row, 1)) = 0 Then
  'grid_fac.TextMatrix(grid_fac.Row, 0) = ""
End If
If i_glosa.Locked = True Then i_glosa.Text = grid_fac.TextMatrix(grid_fac.Row, 7)
End Sub

Private Sub grid_fac_KeyPress(KeyAscii As Integer)
Dim a As Integer
Dim t, WC
Static CONS
If KeyAscii <> 13 Then Exit Sub

If WMODO = "M" Then Exit Sub

If WMODO = "C" And (grid_fac.Col = 2 Or grid_fac.Col = 3) Then

GoTo leer
End If
If WMODO = "C" Then
  If Trim(grid_fac.TextMatrix(grid_fac.Row, 9)) <> "8" Then
    If Trim(grid_fac.TextMatrix(grid_fac.Row, 0)) = "" Then Exit Sub
    If Trim(grid_fac.TextMatrix(grid_fac.Row, 1)) <> "" And grid_fac.Col = 2 Or grid_fac.Col = 3 Then GoTo leer
    'If Trim(grid_fac.TextMatrix(grid_fac.Row, 8)) <> "0" Then Exit Sub
    If Left(grid_fac.TextMatrix(grid_fac.Row, 0), 1) = "T" Then Exit Sub
  End If
End If


If grid_fac.Col = 1 And WMODO = "I" Then
   a = Val(grid_fac.TextMatrix(grid_fac.Row - 1, 0))
   a = a + 1
  grid_fac.TextMatrix(grid_fac.Row, 0) = a
End If
If grid_fac.Col = 4 Then
  If Val(grid_fac.TextMatrix(grid_fac.Row, 0)) = 0 And WMODO = "I" Then Exit Sub
  If Val(grid_fac.TextMatrix(grid_fac.Row, 1)) = 0 And WMODO = "C" Then Exit Sub
  If Val(grid_fac.TextMatrix(grid_fac.Row, 4)) = 0 And Val(grid_fac.TextMatrix(grid_fac.Row, 5)) = 0 Then
  Else
    If Val(grid_fac.TextMatrix(grid_fac.Row, 4)) = 0 Then Exit Sub
  End If
End If

If grid_fac.Col = 5 Then
  If Val(grid_fac.TextMatrix(grid_fac.Row, 0)) = 0 And WMODO = "I" Then Exit Sub
  If Val(grid_fac.TextMatrix(grid_fac.Row, 1)) = 0 And WMODO = "C" Then Exit Sub
  If Val(grid_fac.TextMatrix(grid_fac.Row, 4)) = 0 And Val(grid_fac.TextMatrix(grid_fac.Row, 5)) = 0 Then
  Else
     If Val(grid_fac.TextMatrix(grid_fac.Row, 5)) = 0 Then Exit Sub
  End If
End If
If WMODO = "I" Or WMODO = "C" Then
    TEXTOVAR.Left = grid_fac.Left + grid_fac.CellLeft '+ESTADO.Left
    TEXTOVAR.Width = grid_fac.CellWidth
    TEXTOVAR.Height = grid_fac.CellHeight + 30
    TEXTOVAR.Top = grid_fac.Top + grid_fac.CellTop - 30 '+ ESTADO.Top '- 840  ' +480
    TEXTOVAR.Text = grid_fac.TextMatrix(grid_fac.Row, grid_fac.Col)
    wfila_act = grid_fac.Row
    TEXTOVAR.Visible = True
    Azul3 TEXTOVAR, TEXTOVAR
    TEXTOVAR.SetFocus
End If
leer:
If grid_fac.TextMatrix(grid_fac.Row, 1) <> "" Then
 SQ_OPER = 1
 PUB_CUENTA = grid_fac.TextMatrix(grid_fac.Row, 1)
 PUB_CODCIA = LK_CODCIA
 LEER_COM_LLAVE
 If Not com_llave.EOF Then
   grid_fac.TextMatrix(grid_fac.Row, 2) = com_llave!com_DESCRIPCION
 End If
End If


End Sub

Private Sub grid_fac_KeyUp(KeyCode As Integer, Shift As Integer)
Dim WC
Dim a, WF As Integer
Dim tf, t, tc
Dim SALE As Boolean

'If WMODO = "C" Then Exit Sub

If cop_llave!COP_FLAG_MAYORIZACION = "M" Then
 'MsgBox "Ojo estaba Mayorizado..."
End If


'If Left(grid_fac.TextMatrix(grid_fac.Row, 0), 2) <> "MA" Then Exit Sub
 If KeyCode = 32 Then
  If WMODO <> "C" Then Exit Sub
  If Left(grid_fac.TextMatrix(grid_fac.Row, 0), 1) = "T" Then Exit Sub
  tc = grid_fac.Col
  For fila = 0 To grid_fac.Cols - 1
      grid_fac.Col = fila
      If grid_fac.CellBackColor = QBColor(12) Then
         grid_fac.CellBackColor = QBColor(15)
         grid_fac.TextMatrix(grid_fac.Row, 9) = "9"
      Else
         grid_fac.CellBackColor = QBColor(12)
         grid_fac.TextMatrix(grid_fac.Row, 9) = "-1"
      End If
  Next fila
  grid_fac.Col = tc
  grid_fac.SetFocus
  Exit Sub
End If
If KeyCode = 45 Then
    If grid_fac.Row >= grid_fac.Rows - 1 Then Exit Sub
    Wsec = Wsec + 1
    If Trim(grid_fac.TextMatrix(grid_fac.Row + 1, 11)) = "8" Then
         Exit Sub
    Else
      If Trim(grid_fac.TextMatrix(grid_fac.Row + 1, 0)) = "T" Then Exit Sub
    End If
    If Val(grid_fac.TextMatrix(grid_fac.Row, 4)) = 0 And Val(grid_fac.TextMatrix(grid_fac.Row, 5)) = 0 Then Exit Sub
    grid_fac.AddItem "", grid_fac.Row + 1
    'grid_fac.TextMatrix(grid_fac.Row + 1, 0) = "MAN. " & Format(grid_fac.TextMatrix(grid_fac.Row, 10), "dd/mm/yyyy")
    grid_fac.TextMatrix(grid_fac.Row + 1, 0) = Format(grid_fac.TextMatrix(grid_fac.Row, 10), "dd/mm/yyyy")
    grid_fac.TextMatrix(grid_fac.Row + 1, 6) = Wsec
    grid_fac.TextMatrix(grid_fac.Row + 1, 8) = grid_fac.TextMatrix(grid_fac.Row, 8)
    grid_fac.TextMatrix(grid_fac.Row + 1, 3) = grid_fac.TextMatrix(grid_fac.Row, 3)
    grid_fac.TextMatrix(grid_fac.Row + 1, 7) = grid_fac.TextMatrix(grid_fac.Row, 7)
    grid_fac.TextMatrix(grid_fac.Row + 1, 10) = grid_fac.TextMatrix(grid_fac.Row, 10)
    grid_fac.TextMatrix(grid_fac.Row + 1, 11) = "8"
    grid_fac.Row = grid_fac.Row + 1
    grid_fac.Col = 1
    grid_fac.SetFocus
End If
'Exit Sub
If KeyCode = 46 And Left(cmdIngreso.Caption, 2) = "&G" Then
If grid_fac.Rows <= 3 Then
Else
   pub_mensaje = MsgBox("Desea Quitar el Item de la Cuenta : " & Trim(grid_fac.TextMatrix(grid_fac.Row, 1)), vbYesNo + vbExclamation + vbDefaultButton2, Pub_Titulo)
   If pub_mensaje = vbNo Then
     grid_fac.SetFocus
     Exit Sub
   Else
   grid_fac.RemoveItem (grid_fac.Row)
   grid_fac.Refresh
   suma_grid
   grid_fac.SetFocus
   End If
End If
End If
'grid_fac.SetFocus
Exit Sub



End Sub


Private Sub i_cuenta_GotFocus()
Azul i_cuenta, i_cuenta
lcuenta.Caption = ""
End Sub

Private Sub i_d_h_KeyPress(KeyAscii As Integer)
Azul i_importe, i_importe
End Sub

Private Sub i_fecha2_GotFocus()
If grid_fac.Visible = False Then Exit Sub
If WMODO = "I" Then Exit Sub
WMODO = "C"
grid1.Visible = False
mayor.Value = 0
CABE_MAN

End Sub

Private Sub i_fecha2_KeyPress(KeyAscii As Integer)
'On Error GoTo pasa
Dim Wflag As String * 1
Dim wsFECHA1, WS_FECHA2
If KeyAscii <> 13 Then
  GoTo fin
End If

If Trim(WMODO) = "" Then Exit Sub
If Left(cmdConsultar.Caption, 2) = "&M" And Left(cmdIngreso.Caption, 2) = "&I" Then Exit Sub

If WMODO = "C" Then
 GoTo SIGUE
End If
If WMODO = "I" Then
i_tipdoc.SetFocus
End If
 Exit Sub

SIGUE:
' ' OTRO CASO
Dim WS_SALDO As Currency
Dim Tit As String
Dim a As Integer
Dim i As Integer
Dim success%
Dim con_cuenta As String * 1
If KeyAscii <> 13 Then
  GoTo fin
End If
con_cuenta = "N"
If Trim(i_cuenta.Text) <> "" Then
 SQ_OPER = 1
 PUB_CUENTA = i_cuenta.Text
 PUB_CODCIA = LK_CODCIA
 LEER_COM_LLAVE
 If com_llave.EOF Then
   MsgBox "Su Codigo de Cuenta NO Existe ...", 48, Pub_Titulo
   i_cuenta.SetFocus
 End If
 con_cuenta = "S"
End If

Screen.MousePointer = 11
grid_fac.Visible = False
DoEvents
'success% = SetWindowPos(FrmProcesa.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
DoEvents
Dim ws_tot_debe, ws_tot_haber As Currency
Dim ws_fecha  As String
Dim ws_voucher As Currency



PSCOV_MAYOR.rdoParameters(0) = LK_CODCIA
PSCOV_MAYOR.rdoParameters(1) = i_fecha2.Text
PSCOV_MAYOR.rdoParameters(2) = #1/1/2020#

cov_mayor.Requery
fila = 1
cov_mayor.MoveFirst
If cov_mayor.EOF = True Then
   Screen.MousePointer = 0
   MsgBox "No hay Registros....", 48, Pub_Titulo
   i_fecha2.SetFocus

   grid_fac.Visible = True
   GoTo fin
End If
LOC_CANCELA = 1
cancelar.Enabled = False
cmdConsultar.Enabled = False
Frame1.Enabled = False
cmdcorta.Enabled = True
PB.Visible = True
DoEvents
PB.Min = 0
PB.Max = cov_mayor.RowCount
PB.Value = 0
ws_tot_debe = 0
ws_tot_haber = 0
Tit = grid_fac.FormatString
grid_fac.FormatString = Tit
ws_fecha = cov_mayor!COV_FECHA_VOUCHER
ws_voucher = cov_mayor!COV_NRO_VOUCHER
cmdcorta.SetFocus
Wflag = "X"
Do Until cov_mayor.EOF
   PB.Value = PB.Value + 1
    If Trim(i_voucher.Text) <> "" Then
      If cov_mayor!COV_NRO_VOUCHER <> Val(i_voucher.Text) Then GoTo OTRO
      If Wflag = "X" Then
        Wflag = ""
        ws_voucher = cov_mayor!COV_NRO_VOUCHER
      End If
    End If
       If con_cuenta = "N" Then
        If ws_voucher <> cov_mayor!COV_NRO_VOUCHER Then
          fila = fila + 1
          grid_fac.Rows = fila + 1
          grid_fac.TextMatrix(fila, 0) = "Totales "
          grid_fac.TextMatrix(fila, 1) = ""
          grid_fac.TextMatrix(fila, 2) = ""
          grid_fac.TextMatrix(fila, 3) = ""
          grid_fac.TextMatrix(fila, 4) = Format(ws_tot_debe, "##,##0.00")
          grid_fac.TextMatrix(fila, 5) = Format(ws_tot_haber, "##,##0.00")
          grid_fac.TextMatrix(fila, 7) = "1"
          For i = 0 To 7
            grid_fac.Col = i
            grid_fac.Row = fila
            grid_fac.CellBackColor = &HDCE6BB
          Next i
          ws_tot_debe = 0
          ws_tot_haber = 0
        End If
       End If
       If Left(cov_mayor!COV_CODCTA, Len(i_cuenta.Text)) = Trim(i_cuenta.Text) Then ' And Len(i_cuenta.text) = 2 Then
       Else
          If Trim(cov_mayor!COV_CODCTA) <> Trim(i_cuenta.Text) And con_cuenta = "S" Then
             GoTo OTRO
          End If
       End If
       fila = fila + 1
       grid_fac.Rows = fila + 1
       grid_fac.TextMatrix(fila, 0) = Format(cov_mayor!COV_FECHA_VOUCHER, "dd/mm/yyyy")
       If cov_mayor!cov_flag_automatica <> "0" Then
       '    grid_fac.TextMatrix(fila, 0) = "LIB. " & Format(cov_mayor!cov_FECHA_VOUCHER, "dd/mm/yyyy")
       End If
       If cov_mayor!cov_flag_automatica = "0" Then
       '    grid_fac.TextMatrix(fila, 0) = "MAN. " & Format(cov_mayor!cov_FECHA_VOUCHER, "dd/mm/yyyy")
       End If
       If cov_mayor!cov_flag_automatica = "A" Then
       '   grid_fac.TextMatrix(fila, 0) = "AUT. " & Format(cov_mayor!cov_FECHA_VOUCHER, "dd/mm/yyyy")
       End If
       grid_fac.TextMatrix(fila, 8) = cov_mayor!cov_flag_automatica
       grid_fac.TextMatrix(fila, 1) = Trim(cov_mayor!COV_CODCTA)
       grid_fac.TextMatrix(fila, 3) = cov_mayor!COV_NRO_VOUCHER
       grid_fac.TextMatrix(fila, 7) = Trim(cov_mayor!COV_glosa)
       grid_fac.TextMatrix(fila, 10) = cov_mayor!COV_FECHA_VOUCHER
       If cov_mayor!COV_DH = "D" Then
          grid_fac.TextMatrix(fila, 4) = Nulo_Valor0(cov_mayor!COV_IMPORTE)
          grid_fac.TextMatrix(fila, 5) = ""
          ws_tot_debe = ws_tot_debe + Nulo_Valor0(cov_mayor!COV_IMPORTE)
       ElseIf cov_mayor!COV_DH = "H" Then
             grid_fac.TextMatrix(fila, 4) = ""
             grid_fac.TextMatrix(fila, 5) = Nulo_Valor0(cov_mayor!COV_IMPORTE)
             ws_tot_haber = ws_tot_haber + Nulo_Valor0(cov_mayor!COV_IMPORTE)
       End If
       grid_fac.TextMatrix(fila, 6) = cov_mayor!COV_NRO_MOV
'       If cov_mayor!cov_flag_automatica = "A" Then
'        grid_fac.Row = fila
'           For a = 1 To grid_fac.Cols - 1
'               grid_fac.Col = a
'               grid_fac.CellBackColor = vbCyan
'           Next a
'       End If
OTRO:
        ws_fecha = cov_mayor!COV_FECHA_VOUCHER
        ws_voucher = cov_mayor!COV_NRO_VOUCHER
        cov_mayor.MoveNext
        DoEvents
        If LOC_CANCELA = 2 Then LOC_CANCELA = 0: Exit Do
Loop
          fila = fila + 1
          grid_fac.Rows = fila + 1
          grid_fac.TextMatrix(fila, 0) = "Totales "
          grid_fac.TextMatrix(fila, 1) = ""
          grid_fac.TextMatrix(fila, 2) = ""
          grid_fac.TextMatrix(fila, 3) = ""
          grid_fac.TextMatrix(fila, 4) = Format(ws_tot_debe, "##,##0.00")
          grid_fac.TextMatrix(fila, 5) = Format(ws_tot_haber, "##,##0.00")
          grid_fac.TextMatrix(fila, 7) = "1"
          For i = 0 To 7
            grid_fac.Col = i
            grid_fac.Row = fila
            grid_fac.CellBackColor = &HDCE6BB
          Next i

pasa:
 LOC_CANCELA = 0
 cancelar.Enabled = True
 PB.Visible = False
 cmdcorta.Enabled = False
 DoEvents
 suma_grid
 ws_tot_debe = 0
 ws_tot_haber = 0
 grid_fac.Visible = True
 Frame1.Enabled = True
 cmdConsultar.Enabled = True
 grid_fac.Col = 1
 grid_fac.Row = 2
 If grid_fac.Enabled And grid_fac.Visible Then grid_fac.SetFocus
 Screen.MousePointer = 0

fin:

End Sub

Private Sub i_glosa_GotFocus()
temporal = Trim(i_glosa.Text)
End Sub

Private Sub i_glosa_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 And WMODO = "C" Then
  i_glosa.Text = temporal
  i_glosa.Enabled = False
  i_glosa.Locked = True
  grid_fac.SetFocus
End If
If KeyAscii = 13 Then
  If WMODO = "C" Then
    i_glosa.Enabled = False
    i_glosa.Locked = True
    For fila = loc_ini To loc_fin
      grid_fac.TextMatrix(fila, 7) = i_glosa.Text
      grid_fac.TextMatrix(fila, 9) = "9"
    Next
    grid_fac.SetFocus
    Exit Sub
  End If
   If grid_fac.Rows > 2 Then
     grid_fac.Col = 1
     grid_fac.Row = 2
     grid_fac.SetFocus
   End If
End If

End Sub

Private Sub i_glosa_LostFocus()
If WMODO = "C" Then
  i_glosa.Text = temporal
  i_glosa.Enabled = False
  i_glosa.Locked = True
  grid_fac.SetFocus
End If

End Sub

Private Sub i_importe_KeyPress(KeyAscii As Integer)
Dim valor As Currency
Dim subtotal As Currency
Dim tf As Integer
SOLO_DECIMAL i_importe, KeyAscii

If KeyAscii <> 13 Then
   GoTo fin
End If
If WMODO = "C" Then
' Exit Sub
End If
If com_llave.EOF = True Then
   MsgBox "Cuenta No Existe... ", 48, Pub_Titulo
   Azul i_cuenta, i_cuenta
   GoTo fin
Else
   If com_llave!com_cuenta <> Val(i_cuenta.Text) Then
   MsgBox "Primero seleccione Articulo.. ", 48, Pub_Titulo
   GoTo fin
   End If
End If
If i_importe.Text = "" Then
   i_importe.SetFocus
   GoTo fin
End If
If i_d_h.Text = "" Then
   i_d_h.SetFocus
   SendKeys "%{UP}"
   GoTo fin
End If
If grid_fac.TextMatrix(2, 1) = "" Then fila = 1

salta:
   
   If fila = grid_fac.Rows - 2 Then grid_fac.Rows = grid_fac.Rows + 1

   fila = fila + 1
   grid_fac.Row = fila
   grid_fac.Col = 0
   grid_fac.CellAlignment = 7
   grid_fac.Text = fila
   grid_fac.Col = 1
   grid_fac.CellAlignment = 1
   grid_fac.Text = com_llave!com_cuenta
   grid_fac.Col = 2
   grid_fac.Text = com_llave!com_DESCRIPCION
   grid_fac.Col = 3
   grid_fac.Text = i_glosa.Text
   If Trim(i_d_h.Text) = "D" Then
     grid_fac.Col = 5
     grid_fac.Text = ""
     grid_fac.Col = 4
     grid_fac.Text = ""
     grid_fac.CellAlignment = 7
     grid_fac.Text = i_importe.Text
   ElseIf Trim(i_d_h.Text) = "H" Then
      grid_fac.Col = 4
      grid_fac.Text = ""
      grid_fac.Col = 5
      grid_fac.Text = ""
      grid_fac.CellAlignment = 7
      grid_fac.Text = i_importe.Text
   End If
   suma_grid
   If fila > 10 Then
      grid_fac.SetFocus
      SendKeys "{HOME}", True
      SendKeys "{DOWN}", True
      SendKeys "{UP 6}", True
  End If
 i_cuenta.SetFocus
fin:


End Sub


Private Sub i_tipdoc_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
   GoTo fin
End If
i_glosa.SetFocus

fin:

End Sub

Private Sub i_cuenta_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer
Dim ws_tot_debe  As Currency
Dim f_final_h As Integer
Dim f_final_d As Integer
Dim ws_tot_haber  As Currency
Dim xcuenta
Dim i As Integer
If KeyAscii = 27 Then
  i_cuenta.Text = ""
  Exit Sub
End If


If KeyAscii <> 13 Then
   GoTo fin
End If
If Left(i_cuenta.Text, 1) = "*" Then
  BUSCAR_CTA 1
  Exit Sub
End If
 If i_cuenta.Text = "" Then
  i_voucher.SetFocus
  llave1 = ""
  Exit Sub
 Else
    SQ_OPER = 1
    PUB_CUENTA = i_cuenta.Text
    PUB_CODCIA = LK_CODCIA
    LEER_COM_LLAVE
    If com_llave.EOF Then
       MsgBox "CUENTA NO EXISTE ...", 48, Pub_Titulo
       Azul i_cuenta, i_cuenta
       GoTo fin
    Else
       lcuenta.Caption = Trim(com_llave!com_DESCRIPCION)
    End If
 End If
If mayor.Value = 0 Then
 i_voucher.SetFocus
 Exit Sub
End If
WMODO = "M"
grid1.Visible = True
PSCOV_MAYOR.rdoParameters(0) = LK_CODCIA
PSCOV_MAYOR.rdoParameters(1) = i_fecha2.Text
PSCOV_MAYOR.rdoParameters(2) = #1/1/2020#
cov_mayor.Requery
If cov_mayor.EOF = True Then
   MsgBox "No hay Registros....", 48, Pub_Titulo
   Screen.MousePointer = 0
   grid1.Visible = True
   GoTo fin
End If
cancelar.Enabled = False
cmdConsultar.Enabled = False
Frame1.Enabled = False
cmdcorta.Enabled = True
PB.Visible = True
DoEvents
PB.Min = 0
PB.Max = cov_mayor.RowCount
PB.Value = 0
ws_tot_debe = 0
ws_tot_haber = 0
Dim f_final As Integer
Dim WS_SALDO_INICIAL As Currency
Dim WSALDO As Currency
fila = 3
cabe_mayor
xcuenta = 3
f_final = 0
grid1.Rows = 4
f_final_d = 0
f_final_h = 0

SQ_OPER = 1
PUB_CUENTA = Left(Trim(i_cuenta.Text), 2)
PUB_CODCIA = LK_CODCIA
LEER_COM_LLAVE
WSALDO = 0
WS_SALDO_INICIAL = 0
If Not com_llave.EOF Then
  WSALDO = ((Val(com_llave!COM_DEB_ANO) + Val(com_llave!COM_DEB_MES)) * com_llave!com_signo_d) + ((Val(com_llave!COM_HAB_ANO) + Val(com_llave!COM_HAB_MES)) * com_llave!com_signo_h)
  WSALDO = Abs(WSALDO)
  WS_SALDO_INICIAL = (Val(com_llave!COM_DEB_ANO) * com_llave!com_signo_d) + (Val(com_llave!COM_HAB_ANO) * com_llave!com_signo_h)
  If (Val(com_llave!COM_DEB_ANO) * com_llave!com_signo_d) > (Val(com_llave!COM_HAB_ANO) * com_llave!com_signo_h) Then
  'If WS_SALDO_INICIAL > 0 Then
     grid1.TextMatrix(2, 0) = "          Saldo"
     grid1.TextMatrix(2, 1) = "Inicial:"
     grid1.TextMatrix(2, 2) = Format(WS_SALDO_INICIAL, "0.00")
     ws_tot_debe = ws_tot_debe + WS_SALDO_INICIAL
     'fila = fila + 2
  Else ' If WS_SALDO_INICIAL < 0 Then
     grid1.TextMatrix(2, 4) = "          Saldo"
     grid1.TextMatrix(2, 5) = "Inicial:"
     grid1.TextMatrix(2, 6) = Format(WS_SALDO_INICIAL, "0.00")
     ws_tot_haber = ws_tot_haber + WS_SALDO_INICIAL
     'xcuenta = xcuenta + 2
  'Else
  End If
End If

 
Do Until cov_mayor.EOF
   'PB.Value = PB.Value + 1
       If Left(cov_mayor!COV_CODCTA, Len(i_cuenta.Text)) = Trim(i_cuenta.Text) Then ' And Len(i_cuenta.text) = 2 Then
       Else
          If Trim(cov_mayor!COV_CODCTA) <> Trim(i_cuenta.Text) Then
             GoTo OTRO
          End If
       End If
       If cov_mayor!COV_DH = "D" Then
         grid1.TextMatrix(fila, 0) = Format(cov_mayor!COV_FECHA_VOUCHER, "dd/mm/yyyy")
         grid1.TextMatrix(fila, 1) = cov_mayor!COV_NRO_VOUCHER
         grid1.TextMatrix(fila, 2) = Nulo_Valor0(cov_mayor!COV_IMPORTE)
         ws_tot_debe = ws_tot_debe + cov_mayor!COV_IMPORTE
         f_final_d = fila
       Else
         f_final_h = xcuenta
         grid1.TextMatrix(xcuenta, 4) = Format(cov_mayor!COV_FECHA_VOUCHER, "dd/mm/yyyy")
         grid1.TextMatrix(xcuenta, 5) = cov_mayor!COV_NRO_VOUCHER
         grid1.TextMatrix(xcuenta, 6) = Nulo_Valor0(cov_mayor!COV_IMPORTE)
         ws_tot_haber = ws_tot_haber + cov_mayor!COV_IMPORTE
       End If
       If cov_mayor!COV_DH = "H" Then
             xcuenta = xcuenta + 1
             If xcuenta > fila Then
                grid1.Rows = grid1.Rows + 1
              End If
        End If
        If cov_mayor!COV_DH = "D" Then
              fila = fila + 1
              If fila > xcuenta Then
                grid1.Rows = grid1.Rows + 1
              End If
        End If
        
OTRO:
        cov_mayor.MoveNext
Loop

 grid1.Visible = True
 If f_final_h > f_final_d Then
     grid1.Rows = xcuenta + 1
     xcuenta = xcuenta + 1
 Else
     grid1.Rows = fila + 1
     fila = fila + 1
 End If
 
 If ws_tot_debe > ws_tot_haber Then
     grid1.Rows = grid1.Rows + 2
     grid1.RowHeight(grid1.Rows - 1) = 300
     grid1.TextMatrix(grid1.Rows - 2, 4) = "Saldo  "
     grid1.TextMatrix(grid1.Rows - 2, 6) = Format(WSALDO, "##,##0.00")
     
     'grid1.TextMatrix(fila - 1, 4) = "Saldo  "
     'grid1.TextMatrix(fila - 1, 6) = Format(WSALDO, "##,##0.00")
     ws_tot_haber = ws_tot_haber + WSALDO
    Else
     grid1.Rows = grid1.Rows + 2
     grid1.RowHeight(grid1.Rows - 1) = 300
     grid1.TextMatrix(grid1.Rows - 2, 0) = "Saldo  "
     grid1.TextMatrix(grid1.Rows - 2, 2) = Format(WSALDO, "##,##0.00")

     'grid1.TextMatrix(xcuenta - 1, 0) = "Saldo  "
     'grid1.TextMatrix(xcuenta - 1, 2) = Format(WSALDO, "##,##0.00")
     ws_tot_debe = ws_tot_debe + WSALDO
    End If
    'fila = fila + 1
    'grid1.Rows = fila + 1
    grid1.RowHeight(grid1.Rows - 1) = 300
    grid1.TextMatrix(grid1.Rows - 1, 0) = "Totales "
    grid1.TextMatrix(grid1.Rows - 1, 1) = ""
    grid1.TextMatrix(grid1.Rows - 1, 2) = ""
    grid1.TextMatrix(grid1.Rows - 1, 3) = ""
    grid1.TextMatrix(grid1.Rows - 1, 2) = Format(ws_tot_debe, "##,##0.00")
    grid1.TextMatrix(grid1.Rows - 1, 6) = Format(ws_tot_haber, "##,##0.00")
    grid1.SetFocus
    For i = 0 To 6
        grid1.Col = i
        grid1.Row = fila
        grid1.CellBackColor = &HDCE6BB
    Next i
cancelar.Enabled = True
Frame1.Enabled = True
cmdcorta.Enabled = False
PB.Visible = False

fin:
End Sub

Private Sub i_voucher_GotFocus()
grid1.Visible = False
mayor.Value = 0
CABE_MAN
End Sub

Private Sub i_voucher_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
  i_fecha2_KeyPress 13
  Exit Sub
End If
End Sub

Private Sub mayor_Click()
If mayor.Value = 1 Then
  i_cuenta.SetFocus
End If
End Sub

Private Sub salir_Click()
If LOC_CANCELA = 1 Then
  cmdcorta_Click
  Exit Sub
End If
Unload FORM_CONTA
End Sub


Public Sub LIMPIA_DATOS()
i_tipdoc.BackColor = QBColor(15)
'FORM_CONTA.i_fecha = ""
FORM_CONTA.i_d_h.ListIndex = -1
FORM_CONTA.i_d_h.Enabled = False
FORM_CONTA.i_tipdoc = ""
FORM_CONTA.i_tipdoc.Enabled = False
FORM_CONTA.grid_fac.Clear
fila = 1
FORM_CONTA.i_glosa.Locked = False
FORM_CONTA.i_cuenta = ""
FORM_CONTA.i_cuenta.Enabled = False
FORM_CONTA.i_glosa = ""
FORM_CONTA.i_glosa.Enabled = False
FORM_CONTA.i_importe = ""
FORM_CONTA.i_importe.Enabled = False
cmdIngreso.Enabled = True
cmdConsultar.Enabled = True
WPASA = False
End Sub

Public Sub CABE_MAN()
grid_fac.Cols = 12
grid_fac.Rows = 2
grid_fac.Clear
grid_fac.MergeCells = 0

fila = 0
grid_fac.ColWidth(0) = 1650
grid_fac.ColWidth(1) = 1200
grid_fac.ColWidth(2) = 2500
grid_fac.ColWidth(3) = 1000
grid_fac.ColWidth(4) = 1500
grid_fac.ColWidth(5) = 1500
grid_fac.ColWidth(6) = 0
grid_fac.ColWidth(7) = 0
grid_fac.ColWidth(8) = 0
grid_fac.ColWidth(9) = 0
grid_fac.ColWidth(10) = 0
grid_fac.ColWidth(11) = 0


grid_fac.TextMatrix(0, 0) = "Fecha"
grid_fac.TextMatrix(0, 1) = "Cuenta"
grid_fac.TextMatrix(0, 2) = "Descripcion"
grid_fac.TextMatrix(0, 3) = "Voucher"
grid_fac.TextMatrix(0, 4) = "Debe"
grid_fac.TextMatrix(0, 5) = "Haber"

End Sub
Public Sub suma_grid()
On Error GoTo SALE
Dim WF As Integer
WF = 2
Dim fx As Integer
fx = 1
SUM_H = 0
SUM_D = 0
Do While fx = 1
    If Left(grid_fac.TextMatrix(WF, 0), 1) <> "T" Then
      SUM_D = SUM_D + Val(grid_fac.TextMatrix(WF, 4))
      SUM_H = SUM_H + Val(grid_fac.TextMatrix(WF, 5))
    End If
    WF = WF + 1
    If WF = grid_fac.Rows Then
        fx = 0
    Else
        If Trim(grid_fac.TextMatrix(WF, 0)) = "" Then fx = 0
    End If
Loop
   fila = WF - 1
   grid_fac.TextMatrix(1, 4) = Format(SUM_D, "###,##0.00")
   grid_fac.TextMatrix(1, 5) = Format(SUM_H, "###,##0.00")
Exit Sub
SALE:
cancelar_Click
'MsgBox "Verficar Importe.", 48, Pub_Titulo
'Resume Next
'If TEXTOVAR.Visible Then Azul3 TEXTOVAR, TEXTOVAR
End Sub
Public Sub suma_subtotal()
If WMODO = "I" Then Exit Sub
On Error GoTo SALE
Dim WF As Integer
Dim WFIN As Integer
Dim WINI As Integer

Dim fx As Integer
fx = grid_fac.Row

fx = grid_fac.Row
WF = 0
fila = 0
Do Until WF = 1
   fx = fx - 1
   If Left(grid_fac.TextMatrix(fx, 0), 1) = "T" Or fx = 1 Then
     WINI = fx + 1
     WF = 1
   End If
Loop
fx = grid_fac.Row
WF = 0
Do Until WF = 1
   fx = fx + 1
   If Left(grid_fac.TextMatrix(fx, 0), 1) = "T" Or fx = 1 Then
     WFIN = fx - 1
     WF = 1
   End If
Loop

WF = 2
fx = 1
SUM_H = 0
SUM_D = 0
For fila = WINI To WFIN
    SUM_D = SUM_D + Val(grid_fac.TextMatrix(fila, 4))
    SUM_H = SUM_H + Val(grid_fac.TextMatrix(fila, 5))
Next fila
'fila = WF - 1
grid_fac.TextMatrix(WFIN + 1, 4) = Format(SUM_D, "###,##0.00")
grid_fac.TextMatrix(WFIN + 1, 5) = Format(SUM_H, "###,##0.00")
Exit Sub
SALE:
MsgBox "Verficar Importe.", 48, Pub_Titulo
If TEXTOVAR.Visible Then Azul3 TEXTOVAR, TEXTOVAR
End Sub

Private Sub Consistencias(wsGrid As MSFlexGrid, wsTexto As RichTextBox, wsKeyAscii As Integer)
  Static valor
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

Public Sub CABE_ING()
grid_fac.Cols = 6
grid_fac.Rows = 3
grid_fac.Clear
grid_fac.MergeCells = 4
grid_fac.MergeCol(0) = True
grid_fac.MergeCol(1) = True
grid_fac.MergeCol(2) = True
grid_fac.MergeCol(3) = True
grid_fac.MergeCol(4) = False
grid_fac.MergeCol(5) = False
grid_fac.MergeRow(2) = False
grid_fac.RowHeight(0) = 285
grid_fac.RowHeight(1) = 285
grid_fac.RowHeight(2) = 285

fila = 0
grid_fac.ColWidth(0) = 700
grid_fac.ColWidth(1) = 1400
grid_fac.ColWidth(2) = 3500
grid_fac.ColWidth(3) = 0
grid_fac.ColWidth(4) = 1500
grid_fac.ColWidth(5) = 1500

grid_fac.TextMatrix(0, 0) = "Item"
grid_fac.TextMatrix(0, 1) = "Cuenta"
grid_fac.TextMatrix(0, 2) = "Descripcion"
grid_fac.TextMatrix(0, 3) = "Glosa"
grid_fac.TextMatrix(0, 4) = "Debe"
grid_fac.TextMatrix(0, 5) = "Haber"
grid_fac.TextMatrix(1, 0) = "Item"
grid_fac.TextMatrix(1, 1) = "Cuenta"
grid_fac.TextMatrix(1, 2) = "Descripcion"
grid_fac.TextMatrix(1, 3) = "Glosa"

'grid_fac.MergeCol
'grid_fac.MergeRow(2) = True



End Sub

Private Sub textovar_Change()
If grid_fac.Col = 0 Then
   If IsDate(TEXTOVAR.Text) Then
      grid_fac.Text = Format(TEXTOVAR.Text, "dd/mm/yyyy")
      Exit Sub
   End If
End If
If grid_fac.Col = 1 Then
Else
 grid_fac.Text = Format(TEXTOVAR.Text, "0.00")
 suma_grid
 suma_subtotal
End If
End Sub

Private Sub TEXTOVAR_GotFocus()
temporal = grid_fac.TextMatrix(grid_fac.Row, grid_fac.Col)
End Sub

Private Sub textovar_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
  TEXTOVAR.Text = temporal
  TEXTOVAR.Visible = False
  grid_fac.SetFocus
  Exit Sub
End If
If grid_fac.Col = 4 Or grid_fac.Col = 5 Then Consistencias grid_fac, TEXTOVAR, KeyAscii
If KeyAscii <> 13 Then
   GoTo fin
End If
If grid_fac.Col = 0 Then
 If WMODO = "C" Then
   If Not IsDate(TEXTOVAR.Text) Then
      TEXTOVAR.SetFocus
      Exit Sub
   End If
 End If
End If

If grid_fac.Col <> 1 Then
  TEXTOVAR.Visible = False
  If WMODO = "C" Then
   grid_fac.TextMatrix(grid_fac.Row, 9) = "9"
  End If
  If Trim(grid_fac.TextMatrix(grid_fac.Rows - 1, 1)) = "" Then
     grid_fac.SetFocus
     Exit Sub
  End If
  grid_fac.Rows = grid_fac.Rows + 1
  grid_fac.RowHeight(grid_fac.Rows - 1) = 285
  grid_fac.MergeRow(grid_fac.Rows - 1) = False
  If grid_fac.Rows > 12 Then
      SendKeys "{HOME}", True
      SendKeys "{DOWN}", True
      SendKeys "{UP 6}", True
  End If
  grid_fac.Col = 1
  grid_fac.Row = grid_fac.Rows - 1
  grid_fac.SetFocus
  Exit Sub
End If
If Left(TEXTOVAR.Text, 1) = "*" Then
  BUSCAR_CTA 0
  Exit Sub
End If
If WMODO = "I" Then
     If Val(TEXTOVAR.Text) = 0 Then
       If grid_fac.Col = 4 Then
         grid_fac.Col = 5
       End If
       TEXTOVAR.Visible = False
       Exit Sub
     End If
 End If
If TEXTOVAR.Text = "" Then
  llave1 = ""
  Exit Sub
Else
    SQ_OPER = 1
    PUB_CUENTA = TEXTOVAR.Text
    PUB_CODCIA = LK_CODCIA
    LEER_COM_LLAVE
    If com_llave.EOF Then
         MsgBox "CUENTA NO EXISTE ...", 48, Pub_Titulo
         Azul3 TEXTOVAR, TEXTOVAR
         GoTo fin
    Else
     If Val(com_llave!com_nivel) <> Val(cop_llave!cop_nivel_max) Then
         MsgBox "Cuenta es validad, pero no es Analitica ...", 48, Pub_Titulo
         Azul3 TEXTOVAR, TEXTOVAR
         GoTo fin
     End If
     grid_fac.TextMatrix(grid_fac.Row, 2) = Trim(com_llave!com_DESCRIPCION)
     grid_fac.TextMatrix(grid_fac.Row, 1) = Trim(com_llave!com_cuenta)
    End If
    
 End If
 
TEXTOVAR.Visible = False
If WMODO = "C" Then
   grid_fac.TextMatrix(grid_fac.Row, 9) = "9"
End If
grid_fac.Col = 4
'grid_fac.Row = grid_fac.Row + 1
grid_fac.SetFocus

fin:

End Sub

Private Sub textovar_LostFocus()
'TEXTOVAR.Visible = False
If TEXTOVAR.Visible Then
 '  TEXTOVAR.Visible = False
   grid_fac.Row = wfila_act
'   grid_fac.SetFocus
   Exit Sub
End If

End Sub

Public Sub BUSCAR_CTA(WTIPO As Integer)
Dim WCUENTA As TextBox
Dim wgrupo As String
Dim wq_cuenta As String

LK_TABLA = "BUSCAR"
If WTIPO = 1 Then
 If i_cuenta.Text = "*" Then
  wgrupo = "" 'Trim(i_cuenta.text)
  archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND COM_CUENTA >= '" & wgrupo & "'  ORDER BY COM_CUENTA"
 Else
 i_cuenta.Text = Mid(i_cuenta.Text, 2, Len(i_cuenta.Text))
 wgrupo = Trim(i_cuenta.Text)
 If Val(wgrupo) = 0 Then Exit Sub
 archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND COM_CUENTA >= '" & wgrupo & "' AND COM_CUENTA < '" & Trim(Str(Val(wgrupo) + 1)) & "'  ORDER BY COM_CUENTA"
 End If
Else
 If TEXTOVAR.Text = "*" Then
  wgrupo = "" 'Trim(i_cuenta.text)
  archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND COM_CUENTA >= '" & wgrupo & "'  ORDER BY COM_CUENTA"
 Else
 TEXTOVAR.Text = Mid(TEXTOVAR.Text, 2, Len(TEXTOVAR.Text))
 wgrupo = Trim(TEXTOVAR.Text)
 If Val(wgrupo) = 0 Then Exit Sub
 archi = "SELECT * FROM COMAEST WHERE COM_CODCIA = ? AND COM_CUENTA >= '" & wgrupo & "' AND COM_CUENTA < '" & Trim(Str(Val(wgrupo) + 1)) & "'  ORDER BY COM_CUENTA"
 End If
End If
Load frmBuscacta
frmBuscacta.lbltabla.Caption = LK_TABLA
frmBuscacta.Show 1
wq_cuenta = Trim(frmBuscacta.tcuenta)
If wq_cuenta <> "" Then
  If WTIPO = 1 Then
    i_cuenta.Text = Trim(frmBuscacta.tcuenta)
    lcuenta.Caption = Trim(frmBuscacta.tnombre.Text)
  Else
  TEXTOVAR.Text = Trim(frmBuscacta.tcuenta)
  End If
End If
Unload frmBuscacta
If wq_cuenta <> "" Then
   If WTIPO = 1 Then
     i_cuenta_KeyPress 13
   Else
     textovar_KeyPress 13
   End If
ElseIf wq_cuenta <> "" Then
  If WTIPO = 1 Then
     i_cuenta.SetFocus
  Else
     textovar_KeyPress 13
  End If
Else
  Azul3 TEXTOVAR, TEXTOVAR
End If


End Sub

Public Sub cabe_mayor()
 grid1.Cols = 7
 grid1.Rows = 2
 grid1.Clear
 grid1.ColWidth(0) = 1200
 grid1.ColWidth(1) = 1000
 grid1.ColWidth(2) = 1200
 grid1.ColWidth(3) = 300
 grid1.ColWidth(4) = 1200
 grid1.ColWidth(5) = 1000
 grid1.ColWidth(6) = 1200
 grid1.TextMatrix(0, 0) = "Fecha"
 grid1.TextMatrix(1, 0) = "Debe"
 grid1.TextMatrix(0, 1) = "Voucher"
 grid1.TextMatrix(0, 2) = "Importe"
 grid1.TextMatrix(0, 3) = " "
 grid1.TextMatrix(0, 4) = "Fecha"
 grid1.TextMatrix(0, 5) = "Voucher"
 grid1.TextMatrix(0, 6) = "Importe"
 grid1.TextMatrix(1, 6) = "Haber"
 
End Sub
