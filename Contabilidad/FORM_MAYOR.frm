VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.2#0"; "Comctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form FORM_MAYOR 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Libro Mayor"
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
   Icon            =   "FORM_MAYOR.frx":0000
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4890
   ScaleWidth      =   6600
   Tag             =   "55"
   WindowState     =   2  'Maximized
   Begin VB.TextBox i_voucher 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   385
      Left            =   4800
      MaxLength       =   10
      TabIndex        =   27
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   240
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   360
      TabIndex        =   21
      Top             =   6120
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
         ItemData        =   "FORM_MAYOR.frx":0442
         Left            =   3240
         List            =   "FORM_MAYOR.frx":044C
         Style           =   2  'Dropdown List
         TabIndex        =   23
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
         TabIndex        =   22
         TabStop         =   0   'False
         Tag             =   "9999"
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label LCODART 
         Caption         =   "Importe:"
         Height          =   255
         Index           =   1
         Left            =   5280
         TabIndex        =   25
         Tag             =   "9999"
         Top             =   240
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label LCODART 
         Caption         =   "D/H"
         Height          =   255
         Index           =   2
         Left            =   2880
         TabIndex        =   24
         Tag             =   "9999"
         Top             =   240
         Visible         =   0   'False
         Width           =   405
      End
   End
   Begin VB.TextBox i_cuenta 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3000
      MaxLength       =   10
      TabIndex        =   19
      TabStop         =   0   'False
      Tag             =   "9999"
      Top             =   240
      Width           =   1455
   End
   Begin VB.TextBox i_tipdoc 
      Height          =   375
      Left            =   7200
      TabIndex        =   13
      Text            =   "i_tipdoc"
      Top             =   240
      Width           =   1935
   End
   Begin VB.ComboBox i_fecha2 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   240
      Width           =   2535
   End
   Begin ComctlLib.ProgressBar Barra 
      Height          =   255
      Left            =   240
      TabIndex        =   4
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
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H00000000&
      Height          =   5295
      Left            =   0
      TabIndex        =   2
      Tag             =   "100"
      Top             =   600
      Width           =   9495
      Begin VB.CommandButton cmdcambiar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cam&biar Glosa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6960
         Style           =   1  'Graphical
         TabIndex        =   26
         TabStop         =   0   'False
         Tag             =   "9999"
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton SALIR 
         Caption         =   "Ce&rrar"
         Height          =   495
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   18
         TabStop         =   0   'False
         Tag             =   "9999"
         Top             =   4440
         Width           =   1095
      End
      Begin VB.CommandButton cancelar 
         BackColor       =   &H00C0C0C0&
         Caption         =   "&Cancelar"
         Height          =   495
         Left            =   8160
         Style           =   1  'Graphical
         TabIndex        =   17
         TabStop         =   0   'False
         Tag             =   "9999"
         Top             =   3720
         Width           =   1095
      End
      Begin RichTextLib.RichTextBox TEXTOVAR 
         Height          =   375
         Left            =   6360
         TabIndex        =   16
         Top             =   1080
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         _Version        =   327680
         BackColor       =   16776960
         BorderStyle     =   0
         MultiLine       =   0   'False
         TextRTF         =   $"FORM_MAYOR.frx":0456
      End
      Begin ComctlLib.ProgressBar PB 
         Height          =   255
         Left            =   2040
         TabIndex        =   10
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
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   720
         TabIndex        =   1
         TabStop         =   0   'False
         Tag             =   "9999"
         Text            =   "i_glosa"
         Top             =   240
         Width           =   6135
      End
      Begin VB.Frame Frame1 
         Height          =   3495
         Left            =   8160
         TabIndex        =   7
         Top             =   120
         Width           =   1215
         Begin VB.CommandButton cmdConsultar 
            Appearance      =   0  'Flat
            Caption         =   "&Mostrar"
            Height          =   855
            Left            =   120
            Picture         =   "FORM_MAYOR.frx":0506
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   2040
            Width           =   975
         End
         Begin VB.CommandButton cmdIngreso 
            Caption         =   "&Ingreso"
            Height          =   855
            Left            =   120
            Picture         =   "FORM_MAYOR.frx":0650
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   720
            Width           =   975
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grid_fac 
         Height          =   4455
         Left            =   120
         TabIndex        =   0
         Tag             =   "9999"
         Top             =   720
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   7858
         _Version        =   327680
         Rows            =   3
         FixedRows       =   2
         FocusRect       =   2
         HighLight       =   2
         GridLines       =   2
         AllowUserResizing=   3
      End
      Begin VB.CommandButton cmdcorta 
         Caption         =   "Detener"
         Enabled         =   0   'False
         Height          =   285
         Left            =   6240
         TabIndex        =   11
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Lcuenta 
         BackColor       =   &H00C0C0C0&
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
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   4560
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "Glosa:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Tag             =   "9999"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label momen 
         Caption         =   "Un Momento ..."
         Height          =   375
         Left            =   3360
         TabIndex        =   8
         Top             =   1560
         Width           =   1575
      End
   End
   Begin VB.Label LCODART 
      Caption         =   "Voucher"
      Height          =   255
      Index           =   3
      Left            =   4920
      TabIndex        =   28
      Tag             =   "9999"
      Top             =   0
      Width           =   885
   End
   Begin VB.Label LCODART 
      Caption         =   "Cta. para Filtrar."
      Height          =   255
      Index           =   0
      Left            =   3000
      TabIndex        =   20
      Tag             =   "9999"
      Top             =   0
      Width           =   1485
   End
   Begin VB.Label Label5 
      Caption         =   "Desde:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   855
   End
   Begin VB.Label L1 
      Caption         =   "Tipo Doc:"
      Height          =   255
      Left            =   7080
      TabIndex        =   5
      Top             =   0
      Width           =   975
   End
End
Attribute VB_Name = "FORM_MAYOR"
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
Dim WSEC As Integer
Option Explicit

Private Sub cancelar_Click()
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
FORM_MAYOR.Lcuenta.Caption = ""
grid_fac.SetFocus
SendKeys "^{HOME}", True
End Sub

Private Sub cmdcambiar_Click()
On Error GoTo SALE
Dim WF As Integer
Dim WFIN As Integer
Dim WINI As Integer
Dim fx As Integer

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
        cov_llave!COV_FECHA_VOUCHER = cop_llave!COP_FECHA_PROCESO2
        cov_llave!COV_glosa = grid_fac.TextMatrix(fila, 7)
        cov_llave!COV_FECHA_doc = LK_FECHA_DIA
        If Val(grid_fac.TextMatrix(fila, 4)) <> 0 Then
          cov_llave!coV_dh = "D"
          cov_llave!COV_IMPORTE = Val(grid_fac.TextMatrix(fila, 4))
        Else
          cov_llave!coV_dh = "H"
          cov_llave!COV_IMPORTE = Val(grid_fac.TextMatrix(fila, 5))
        End If
        cov_llave!COV_ESTADO = "0"
        cov_llave!COV_CODUSU = LK_CODUSU
        cov_llave!cov_flag_automatica = grid_fac.TextMatrix(fila, 8)
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
       cov_llave!coV_dh = w_dh
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

Exit Sub
End If
PSTEMP_MAYOR(0) = LK_CODCIA
PSTEMP_MAYOR(1) = cop_llave!COP_FECHA_PROCESO2
temp_mayor.Requery
If temp_mayor.EOF Then
   WSEC = 0
Else
   temp_mayor.MoveLast
   WSEC = temp_mayor!COV_NRO_MOV
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

FORM_MAYOR.i_d_h.Enabled = True
FORM_MAYOR.i_tipdoc.Enabled = False
FORM_MAYOR.i_cuenta.Enabled = True
FORM_MAYOR.i_importe.Enabled = True

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

FORM_MAYOR.i_d_h.ListIndex = -1
FORM_MAYOR.i_tipdoc = ""
FORM_MAYOR.grid_fac.Clear
fila = 0
FORM_MAYOR.i_cuenta = ""
FORM_MAYOR.i_glosa = ""
FORM_MAYOR.i_importe = ""
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
PSTEMP_LLAVE(1) = i_fecha2.Text
PSTEMP_LLAVE(2) = #1/1/20#
temp_llave.Requery
If temp_llave.EOF Then
 ws_nro_voucher = 0
Else
 temp_llave.MoveLast
 ws_nro_voucher = temp_llave!COV_NRO_VOUCHER
End If
fila = 1
PUB_CONCEPTO = Trim(i_glosa.Text)
ws_nro_voucher = ws_nro_voucher + 1
PUB_FECHA = i_fecha2.Text
FLAG = False
WS_NRO_MOV = 0
fila = 2
Do While FLAG = False
   If Trim(grid_fac.TextMatrix(fila, 1)) = "" Then GoTo pasa
   PUB_FECHA = i_fecha2.Text
   PUB_CONCEPTO = Trim(grid_fac.TextMatrix(fila, 3))
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
   cov_llave!COV_CODCIA = LK_CODCIA
   cov_llave!COV_NRO_VOUCHER = ws_nro_voucher
   cov_llave!COV_FECHA_VOUCHER = PUB_FECHA
   cov_llave!COV_glosa = PUB_CONCEPTO
   cov_llave!COV_FECHA_doc = LK_FECHA_DIA
   cov_llave!COV_CODCTA = PUB_CUENTA
   cov_llave!coV_dh = w_dh
   cov_llave!COV_IMPORTE = PUB_IMPORTE
   cov_llave!COV_ESTADO = "0"
   cov_llave!COV_CODUSU = LK_CODUSU
   cov_llave!cov_flag_automatica = "0"
   cov_llave.Update
pasa:
   fila = fila + 1
   WS_NRO_MOV = WS_NRO_MOV + 1
   If fila >= FORM_MAYOR.grid_fac.Rows Then
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
FORM_MAYOR.i_d_h.Enabled = True
FORM_MAYOR.i_tipdoc.Enabled = True
FORM_MAYOR.i_cuenta.Enabled = True
FORM_MAYOR.i_glosa.Enabled = True
FORM_MAYOR.i_importe.Enabled = True

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
  FORM_MAYOR.Barra.Visible = False
  Screen.MousePointer = 0
  grid_fac.SetFocus
Else
  MsgBox Err.Description, 48, Pub_Titulo
End If

End Sub



Private Sub Form_Load()
'On Error GoTo SALE
WSEC = 0
LOC_CANCELA = 0
fila = 0
wfila_act = 0
WSELE = ""
Dim ws_indice As Integer
Dim cade
WMODO = ""
pub_cadena = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND COV_FECHA_VOUCHER = ?   ORDER BY  COV_NRO_MOV"
Set PSTEMP_MAYOR = CN.CreateQuery("", pub_cadena)
Set temp_mayor = PSTEMP_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND COV_FECHA_VOUCHER>=? AND COV_FECHA_VOUCHER <=?  ORDER BY COV_NRO_VOUCHER, COV_NRO_MOV"
Set PSTEMP_LLAVE = CN.CreateQuery("", pub_cadena)
Set temp_llave = PSTEMP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM COMOV WHERE COV_CODCIA = ? AND COV_FECHA_VOUCHER >=? AND COV_FECHA_VOUCHER <=?    ORDER BY COV_CODCIA, COV_NRO_VOUCHER, COV_DH, COV_NRO_MOV"
Set PSCOV_MAYOR = CN.CreateQuery("", pub_cadena)
Set cov_mayor = PSCOV_MAYOR.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM COMOV WHERE COV_CODCIA = ?  AND COV_FECHA_VOUCHER=? AND COV_NRO_VOUCHER = ? AND COV_NRO_MOV = ? ORDER BY COV_CODCIA, COV_FECHA_VOUCHER, COV_NRO_VOUCHER, COV_NRO_MOV"
Set PSCOV_LLAVE = CN.CreateQuery("", pub_cadena)
Set cov_llave = PSCOV_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

pub_cadena = "SELECT * FROM COPARAM WHERE COP_CODCIA = ?"
Set PSCOP_LLAVE = CN.CreateQuery("", pub_cadena)
Set cop_llave = PSCOP_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)

PSCOP_LLAVE.rdoParameters(0) = LK_CODCIA
cop_llave.Requery

fila = 0
DoEvents
LIMPIA_DATOS
Dim ws_fecha As Date
i_fecha2.Clear
ws_fecha = cop_llave!cop_fecha_proceso
Do Until fila = 999
   i_fecha2.AddItem Format(ws_fecha, "dd/mmm/YYYY")
   ws_fecha = DateAdd("d", 1, ws_fecha)
   If ws_fecha > cop_llave!COP_FECHA_PROCESO2 Then fila = 999
Loop
i_fecha2.ListIndex = 0
fila = 0
i_tipdoc.Enabled = False
i_tipdoc.BackColor = QBColor(7)
i_voucher.Enabled = False
i_voucher.BackColor = QBColor(7)

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
Dim t, wc
Static CONS
If KeyAscii <> 13 Then Exit Sub

If WMODO = "C" And (grid_fac.Col = 2 Or grid_fac.Col = 3) Then

GoTo leer
End If
If WMODO = "C" Then
  If Trim(grid_fac.TextMatrix(grid_fac.Row, 9)) <> "8" Then
    If Trim(grid_fac.TextMatrix(grid_fac.Row, 0)) = "" Then Exit Sub
    If Trim(grid_fac.TextMatrix(grid_fac.Row, 1)) <> "" And grid_fac.Col = 2 Or grid_fac.Col = 3 Then GoTo leer
    If Trim(grid_fac.TextMatrix(grid_fac.Row, 8)) <> "0" Then Exit Sub
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
    TEXTOVAR.Left = ESTADO.Left + grid_fac.Left + grid_fac.CellLeft
    TEXTOVAR.Width = grid_fac.CellWidth
    TEXTOVAR.Height = grid_fac.CellHeight
    TEXTOVAR.Top = ESTADO.Top + grid_fac.Top + grid_fac.CellTop - 600 '480
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
Dim wc
Dim a, WF As Integer
Dim tf, t, tC
Dim SALE As Boolean

'If WMODO = "C" Then Exit Sub

If cop_llave!COP_FLAG_MAYORIZACION = "M" Then
 'MsgBox "Ojo estaba Mayorizado..."
End If


If Left(grid_fac.TextMatrix(grid_fac.Row, 0), 2) <> "MA" Then Exit Sub
 If KeyCode = 32 Then
  If WMODO <> "C" Then Exit Sub
  tC = grid_fac.Col
  For fila = 1 To grid_fac.Cols - 1
      grid_fac.Col = fila
      If grid_fac.CellBackColor = QBColor(12) Then
         grid_fac.CellBackColor = QBColor(15)
         grid_fac.TextMatrix(grid_fac.Row, 9) = "9"
      Else
         grid_fac.CellBackColor = QBColor(12)
         grid_fac.TextMatrix(grid_fac.Row, 9) = "-1"
      End If
  Next fila
  grid_fac.Col = tC
  grid_fac.SetFocus
  Exit Sub
End If
If KeyCode = 45 Then
    WSEC = WSEC + 1
    If Trim(grid_fac.TextMatrix(grid_fac.Row + 1, 11)) = "8" Then
         Exit Sub
    Else
      If Trim(grid_fac.TextMatrix(grid_fac.Row + 1, 0)) = "T" Then Exit Sub
    End If
    If Val(grid_fac.TextMatrix(grid_fac.Row, 4)) = 0 And Val(grid_fac.TextMatrix(grid_fac.Row, 5)) = 0 Then Exit Sub
    grid_fac.AddItem "", grid_fac.Row + 1
    grid_fac.TextMatrix(grid_fac.Row + 1, 0) = "MAN. " & Format(grid_fac.TextMatrix(grid_fac.Row, 10), "dd/mm/yyyy")
    grid_fac.TextMatrix(grid_fac.Row + 1, 6) = WSEC
    grid_fac.TextMatrix(grid_fac.Row + 1, 8) = grid_fac.TextMatrix(grid_fac.Row, 8)
    grid_fac.TextMatrix(grid_fac.Row + 1, 3) = grid_fac.TextMatrix(grid_fac.Row, 3)
    grid_fac.TextMatrix(grid_fac.Row + 1, 7) = grid_fac.TextMatrix(grid_fac.Row, 7)
    grid_fac.TextMatrix(grid_fac.Row + 1, 10) = grid_fac.TextMatrix(grid_fac.Row, 10)
    grid_fac.TextMatrix(grid_fac.Row + 1, 11) = "8"
    grid_fac.Row = grid_fac.Row + 1
    grid_fac.Col = 1
    grid_fac.SetFocus
End If
Exit Sub
If KeyCode = 46 Then
If grid_fac.Rows <= 3 Then
Else
   pub_mensaje = MsgBox("Desea Quitar el Item de la Cuenta : " & Trim(grid_fac.TextMatrix(grid_fac.Row, 1)), vbYesNo + vbExclamation + vbDefaultButton2, Pub_Titulo)
   If pub_mensaje = vbNo Then
     grid_fac.SetFocus
     Exit Sub
   Else
     grid_fac.RowHeight(grid_fac.Row) = 1
     grid_fac.Row = grid_fac.Row + 1
    
   'grid_fac.RemoveItem (grid_fac.Row)
   'grid_fac.Refresh
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
Lcuenta.Caption = ""
End Sub

Private Sub i_d_h_KeyPress(KeyAscii As Integer)
Azul i_importe, i_importe
End Sub

Private Sub i_fecha2_KeyPress(KeyAscii As Integer)
On Error GoTo pasa
Dim Wflag As String * 1
Dim wsFECHA1, WS_FECHA2
If KeyAscii <> 13 Then
  GoTo fin
End If

If Trim(WMODO) = "" Then Exit Sub
If WMODO = "C" Then
 GoTo sigue
End If
i_tipdoc.SetFocus
 Exit Sub

sigue:
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
PSCOV_MAYOR.rdoParameters(2) = #1/1/20#

cov_mayor.Requery
fila = 1
cov_mayor.MoveFirst
If cov_mayor.EOF = True Then
   MsgBox "No hay Registros....", 48, Pub_Titulo
   i_fecha2.SetFocus
   Screen.MousePointer = 0
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
       grid_fac.TextMatrix(fila, 0) = Format(cov_mayor!COV_FECHA_VOUCHER, "dd/mm/yy")
       
       If cov_mayor!cov_flag_automatica <> "0" Then
           grid_fac.TextMatrix(fila, 0) = "LIB. " & Format(cov_mayor!COV_FECHA_VOUCHER, "dd/mm/yyyy")
       End If
       If cov_mayor!cov_flag_automatica = "0" Then
           grid_fac.TextMatrix(fila, 0) = "MAN. " & Format(cov_mayor!COV_FECHA_VOUCHER, "dd/mm/yyyy")
       End If
       If cov_mayor!cov_flag_automatica = "A" Then
          grid_fac.TextMatrix(fila, 0) = "AUT. " & Format(cov_mayor!COV_FECHA_VOUCHER, "dd/mm/yyyy")
       End If
       grid_fac.TextMatrix(fila, 8) = cov_mayor!cov_flag_automatica
       grid_fac.TextMatrix(fila, 1) = Trim(cov_mayor!COV_CODCTA)
       grid_fac.TextMatrix(fila, 3) = cov_mayor!COV_NRO_VOUCHER
       grid_fac.TextMatrix(fila, 7) = Trim(cov_mayor!COV_glosa)
       grid_fac.TextMatrix(fila, 10) = cov_mayor!COV_FECHA_VOUCHER
       
       If cov_mayor!coV_dh = "D" Then
          grid_fac.TextMatrix(fila, 4) = Nulo_Valor0(cov_mayor!COV_IMPORTE)
          grid_fac.TextMatrix(fila, 5) = ""
          ws_tot_debe = ws_tot_debe + Nulo_Valor0(cov_mayor!COV_IMPORTE)
       ElseIf cov_mayor!coV_dh = "H" Then
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
       Lcuenta.Caption = Trim(com_llave!com_DESCRIPCION)
    End If
    
 End If
i_voucher.SetFocus
fin:
End Sub

Private Sub i_voucher_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
  i_fecha2_KeyPress 13
  Exit Sub
End If
End Sub

Private Sub salir_Click()
If LOC_CANCELA = 1 Then
  cmdcorta_Click
  Exit Sub
End If
Unload FORM_MAYOR
End Sub


Public Sub LIMPIA_DATOS()
i_tipdoc.BackColor = QBColor(15)
'FORM_MAYOR.i_fecha = ""
FORM_MAYOR.i_d_h.ListIndex = -1
FORM_MAYOR.i_d_h.Enabled = False
FORM_MAYOR.i_tipdoc = ""
FORM_MAYOR.i_tipdoc.Enabled = False
FORM_MAYOR.grid_fac.Clear
fila = 1
FORM_MAYOR.i_glosa.Locked = False
FORM_MAYOR.i_cuenta = ""
FORM_MAYOR.i_cuenta.Enabled = False
FORM_MAYOR.i_glosa = ""
FORM_MAYOR.i_glosa.Enabled = False
FORM_MAYOR.i_importe = ""
FORM_MAYOR.i_importe.Enabled = False
cmdIngreso.Enabled = True
cmdConsultar.Enabled = True
WPASA = False
End Sub

Public Sub CABE_MAN()
grid_fac.Cols = 9
grid_fac.Rows = 2
grid_fac.Clear
grid_fac.MergeCells = 0

fila = 0
grid_fac.ColWidth(0) = 1650
grid_fac.ColWidth(1) = 1200
grid_fac.ColWidth(2) = 1000
grid_fac.ColWidth(3) = 600
grid_fac.ColWidth(4) = 1500
grid_fac.ColWidth(5) = 1500
grid_fac.ColWidth(6) = 0
grid_fac.ColWidth(7) = 0
grid_fac.ColWidth(8) = 0


grid_fac.TextMatrix(0, 0) = "Fecha"
grid_fac.TextMatrix(0, 1) = "Cuenta"
grid_fac.TextMatrix(0, 2) = "Descrip."
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
MsgBox "Verficar Importe.", 48, Pub_Titulo
If TEXTOVAR.Visible Then Azul3 TEXTOVAR, TEXTOVAR
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
grid_fac.ColWidth(0) = 400
grid_fac.ColWidth(1) = 1400
grid_fac.ColWidth(2) = 2500
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
  Exit Sub
End If
If grid_fac.Col = 4 Or grid_fac.Col = 5 Then Consistencias grid_fac, TEXTOVAR, KeyAscii
If KeyAscii <> 13 Then
   GoTo fin
End If
If grid_fac.Col = 4 Or grid_fac.Col = 5 Then
 If WMODO = "C" Then
'   If Val(textovar.text) = 0 Then
'       Exit Sub
'   End If
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
 
End If
If grid_fac.Col <> 1 Then
  TEXTOVAR.Visible = False
  If WMODO = "C" Then
   grid_fac.TextMatrix(grid_fac.Row, 9) = "9"
  End If
  If Trim(grid_fac.TextMatrix(grid_fac.Rows - 1, 1)) = "" Then
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
   If frmBuscacta.Visible Then
   Else
      TEXTOVAR.SetFocus
   End If
End If

End Sub

Public Sub BUSCAR_CTA(WTIPO As Integer)
Dim wcuenta As TextBox
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
    Lcuenta.Caption = Trim(frmBuscacta.tnombre.Text)
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
