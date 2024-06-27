VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FrmImp3 
   Caption         =   "Reportes"
   ClientHeight    =   4470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8715
   Icon            =   "FrmImp3.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   8715
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmbmoneda 
      Height          =   315
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   1200
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.ComboBox CP 
      Height          =   315
      ItemData        =   "FrmImp3.frx":0442
      Left            =   360
      List            =   "FrmImp3.frx":044C
      Style           =   2  'Dropdown List
      TabIndex        =   29
      Top             =   600
      Visible         =   0   'False
      Width           =   2655
   End
   Begin ComctlLib.ListView ListView2 
      Height          =   375
      Left            =   5760
      TabIndex        =   28
      Top             =   3960
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame frmop 
      Caption         =   "Clientes :"
      ForeColor       =   &H00800000&
      Height          =   1815
      Left            =   240
      TabIndex        =   21
      Top             =   1560
      Visible         =   0   'False
      Width           =   8535
      Begin VB.OptionButton op1 
         Caption         =   "Todas"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   3000
         TabIndex        =   27
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton op1 
         Caption         =   "Entregadas"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   26
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton op1 
         Caption         =   "Pendientes"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   25
         Top             =   1080
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.ListBox listop 
         Height          =   1410
         Left            =   4200
         Style           =   1  'Checkbox
         TabIndex        =   24
         Top             =   120
         Width           =   4215
      End
      Begin VB.TextBox txt_cli 
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblcliente 
         Height          =   735
         Left            =   1800
         TabIndex        =   23
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.ListBox lisD 
      Height          =   2085
      Left            =   3720
      Style           =   1  'Checkbox
      TabIndex        =   19
      Top             =   840
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Frame frmdocu 
      Caption         =   "Documentos"
      Height          =   855
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Visible         =   0   'False
      Width           =   4215
      Begin VB.TextBox numfin 
         Height          =   285
         Left            =   3120
         TabIndex        =   14
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox numini 
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox serie 
         Height          =   285
         Left            =   960
         TabIndex        =   12
         Top             =   480
         Width           =   735
      End
      Begin VB.ComboBox fbg 
         Height          =   315
         ItemData        =   "FrmImp3.frx":046F
         Left            =   120
         List            =   "FrmImp3.frx":0479
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
      Begin VB.Label docu 
         Caption         =   "Doc."
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   735
      End
      Begin VB.Label docuX 
         Caption         =   "Nª Final"
         Height          =   255
         Left            =   3120
         TabIndex        =   17
         Top             =   240
         Width           =   855
      End
      Begin VB.Label docu 
         Caption         =   "Nª Inicial"
         Height          =   255
         Index           =   1
         Left            =   1920
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.Label docu 
         Caption         =   "Serie"
         Height          =   255
         Index           =   0
         Left            =   960
         TabIndex        =   15
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   1800
      TabIndex        =   6
      Top             =   -120
      Width           =   4455
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
         Width           =   4140
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
      Top             =   3840
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
      Top             =   3840
      Width           =   1095
   End
   Begin ComctlLib.ProgressBar ProgBar 
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   3360
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
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin MSMask.MaskEdBox txtCampo2 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   1200
      Visible         =   0   'False
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
   Begin MSMask.MaskEdBox txtCampo1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
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
   Begin VB.Label lblmoneda 
      Caption         =   "Moneda:"
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2760
      TabIndex        =   31
      Top             =   960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label LUSUARIO 
      Height          =   255
      Left            =   3600
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label lblcampo2 
      AutoSize        =   -1  'True
      Caption         =   "Campo1"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1560
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label lblcampo1 
      AutoSize        =   -1  'True
      Caption         =   "Campo1"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label lblProceso 
      Alignment       =   2  'Center
      Caption         =   "Procesando ..."
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   3600
      Visible         =   0   'False
      Width           =   3375
   End
End
Attribute VB_Name = "FrmImp3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim loc_key As Integer
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
Dim c1 As Integer
Dim f1 As Integer
Dim xcuenta As Integer
Dim i As Integer
Dim Mensaje, Titulo, valorpred As String
Dim Wfile  As String
Dim WFORM  As String
Dim REP_FECHA1
Dim REP_FECHA2
Dim PSPRO1 As rdoQuery
Dim pro1_llave As rdoResultset



Private Sub cerrar_Click()
Unload FrmImp3
End Sub


Private Sub cmdMostrar_Click()

End Sub

Private Sub cmbmoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And Wfile = "CONSO_FAC" Then
  REP_FECHA1 = txtCampo1.Text
  REP_FECHA2 = txtcampo2.Text
  PROC
  Exit Sub
End If

End Sub

Private Sub fbg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Azul serie, serie
End If

End Sub

Private Sub Form_Load()
Screen.MousePointer = 11
If retra_llave.EOF Then
   Screen.MousePointer = 0
   Exit Sub
End If
Screen.MousePointer = 0
Wfile = Trim(retra_llave(3))
WFORM = Trim(retra_llave(7))
lblreporte.Caption = Trim(retra_llave(1))
If Wfile = "CAJA_GRIFO" Or Wfile = "REG_COMPRA" Or Wfile = "REG_COMPRA_COM" Then
 txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 txtcampo2.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 If Wfile = "REG_COMPRA" Or Wfile = "REG_COMPRA_COM" Then
  lblcampo1.Caption = "Fec. de Proceso Contable: "
  lblcampo2.Caption = ""
 Else
  lblcampo1.Caption = "Fecha de Inicial : "
  lblcampo2.Caption = "Fecha de Final: "
 End If
 lblcampo1.Visible = True
 txtCampo1.Mask = "##/##/####"
 txtCampo1.Visible = True
 
 lblcampo2.Visible = True
 txtcampo2.Mask = "##/##/####"
 txtcampo2.Visible = True
End If
If Wfile = "PROGRA" Or Wfile = "PROGRA_TEJ" Then
 txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 txtcampo2.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 lblcampo1.Caption = "Fecha de Inicial : "
 lblcampo1.Visible = True
 txtCampo1.Mask = "##/##/####"
 txtCampo1.Visible = True
 lblcampo2.Caption = "Fecha de Final: "
 lblcampo2.Visible = True
 txtcampo2.Mask = "##/##/####"
 txtcampo2.Visible = True
 frmop.Caption = "Proveedor de Servicio :"
 frmop.Visible = True
 listop.Visible = False
 op1(0).Visible = False
 op1(1).Visible = False
 op1(2).Visible = False

End If
If Wfile = "CONTROL_MOVI" Then
 txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 lblcampo1.Caption = "Fecha de Inicial : "
 lblcampo1.Visible = True
 txtCampo1.Mask = "##/##/####"
 txtCampo1.Visible = True
End If
If Wfile = "SALDO_CTA" Then
 txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 lblcampo1.Caption = "Saldos al : "
 lblcampo1.Visible = True
 txtCampo1.Mask = "##/##/####"
 txtCampo1.Visible = True
 CP.Visible = True
 CP.ListIndex = 0
End If

If Wfile = "CONTROL_STOCK" Then
 txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 lblcampo1.Caption = "Fecha de Desde : "
 lblcampo1.Visible = True
 txtCampo1.Mask = "##/##/####"
 txtCampo1.Visible = True
End If
If Wfile = "NUMFAC_FAL" Then
  frmdocu.Visible = True
  fbg.TabIndex = 0
  fbg.ListIndex = 0
  numfin.Enabled = False
'  numini.Enabled = False
End If
If Wfile = "Wfile = " Then
    frmdocu.Visible = True
    fbg.Clear
    PUB_TIPREG = 50
    PUB_CODCIA = "00"
    SQ_OPER = 2
    LEER_TAB_LLAVE
    Do Until tab_mayor.EOF
        fbg.AddItem Format(tab_mayor!TAB_NUMTAB, "00") & " - " & Trim(tab_mayor!tab_NOMLARGO)
        tab_mayor.MoveNext
    Loop
    fbg.TabIndex = 0
End If


 
If Wfile = "VENTAS_ARTICULO" Then

End If
If Wfile = "CONSO_FAC" Then
 lblMoneda.Visible = True
 cmbmoneda.Clear
 cmbmoneda.AddItem "Todas"
 cmbmoneda.AddItem "S = Soles"
 cmbmoneda.AddItem "D = Dolares"
 cmbmoneda.Visible = True
 cmbmoneda.ListIndex = 0

 txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 txtcampo2.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 lblcampo1.Caption = "Fecha de Inicial : "
 lblcampo1.Visible = True
 txtCampo1.Mask = "##/##/####"
 txtCampo1.Visible = True
 lblcampo2.Caption = "Fecha de Final: "
 lblcampo2.Visible = True
 txtcampo2.Mask = "##/##/####"
 txtcampo2.Visible = True
 lisD.Visible = True
 LUSUARIO.Visible = True
 LUSUARIO.Caption = "Documentos del Usuario :" & LK_CODUSU
End If
If Wfile = "DETALLE_OP" Then
 pub_cadena = "SELECT PED_FECHA, PED_NUMSER, PED_NUMFAC  FROM PEDIDOS WHERE (PED_FECHA >= ? AND PED_FECHA <= ?) AND PED_CODCIA = ? AND PED_TIPMOV = ? AND PED_CIERRE <> ? AND PED_CODCLIE = ?  GROUP BY PED_FECHA, PED_NUMSER, PED_NUMFAC "
 Set PS_REP01 = CN.CreateQuery("", pub_cadena)
 PS_REP01(0) = LK_FECHA_DIA
 PS_REP01(1) = LK_FECHA_DIA
 PS_REP01(2) = 0
 PS_REP01(3) = 0
 PS_REP01(4) = 0
 Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

 txtCampo1.Text = "01/01/" & Format(LK_FECHA_DIA, "yyyy")
 txtcampo2.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 lblcampo1.Caption = "Fec.Emision Ini.: "
 lblcampo1.Visible = True
 txtCampo1.Mask = "##/##/####"
 txtCampo1.Visible = True
 lblcampo2.Caption = "Fec.Emision Fin. : "
 lblcampo2.Visible = True
 txtcampo2.Mask = "##/##/####"
 txtcampo2.Visible = True
 frmop.Visible = True
End If


End Sub


Private Sub lisD_KeyDown(KeyCode As Integer, Shift As Integer)
Dim webusca
If KeyCode = 45 Then

 webusca = InputBox("Buscar Nro. Documento: ")
 If webusca = "" Then Exit Sub
 lisD.Visible = False
 For fila = 0 To lisD.ListCount - 1
   lisD.ListIndex = fila
   If val(webusca) = val(Mid(lisD.Text, 10, 16)) Then
     lisD.Selected(fila) = True
     lisD.Visible = True
     lisD.SetFocus
     Exit Sub
    End If
 Next fila
 lisD.Visible = True
 lisD.SetFocus
 Exit Sub
End If
If KeyCode = 113 Or KeyCode = 114 Then
 For fila = 0 To lisD.ListCount - 1
   If KeyCode = 113 Then
    lisD.Selected(fila) = True
   Else
    lisD.Selected(fila) = False
   End If
 Next fila
End If
If KeyCode = 46 And lisD.ListIndex <> -1 Then
  lisD.RemoveItem lisD.ListIndex
End If
End Sub

Private Sub numfin_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
  If Pantalla.Enabled And Pantalla.Visible Then
     Pantalla.SetFocus
  End If
End If

End Sub

Private Sub numini_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 And Wfile = "CONSO_FAC" Then
  lisD.AddItem Trim(fbg.Text) + "-" + Format(Trim(serie.Text), "000") + " - " + Trim(numini.Text)
  fbg.SetFocus
  'Azul serie, serie
  Exit Sub
End If
If KeyAscii = 13 Then
  Azul numfin, numfin
End If

End Sub

Private Sub op1_Click(Index As Integer)
If val(txt_cli.Text) = 0 Then Exit Sub
 'PED_FECHA = ? AND PED_FECHA = ?) AND PED_CODCIA = ? AND PED_TIPMOV = ? AND PED_CIERRE <> ? AND PED_CODCLIE = ?  ODER BY PED_CODCIA "
PS_REP01(0) = txtCampo1.Text
PS_REP01(1) = txtcampo2.Text
PS_REP01(2) = LK_CODCIA
PS_REP01(3) = 177
If op1(0).Value Then  ' pendientes
 PS_REP01(4) = " "
ElseIf op1(1).Value Then  ' entregadas
 PS_REP01(4) = "X"
ElseIf op1(2).Value Then  ' todas
 PS_REP01(4) = "A"
End If
PS_REP01(5) = val(txt_cli.Text)
llave_rep01.Requery
listop.Clear
If Not llave_rep01.EOF Then
 ProgBar.Min = 0
 ProgBar.max = llave_rep01.RowCount
 ProgBar.Value = 0
End If
ProgBar.Visible = True
DoEvents
Do Until llave_rep01.EOF
 ProgBar.Value = ProgBar.Value + 1
 listop.AddItem Format(llave_rep01!PED_FECHA, "dd/mm/yyyy") & " O/P. " & Format(llave_rep01!PED_numser, "000") & " - " & Format(llave_rep01!PED_NUMFAC, "0000000")
 llave_rep01.MoveNext
Loop
listop.SetFocus
ProgBar.Visible = False
DoEvents

End Sub

Private Sub Pantalla_Click()
Dim wsFECHA1
Dim wsFECHA2
'On Error GoTo SALE
If Wfile = "CAJA_GRIFO" Then
  Call CAJA_GRIFO
End If
If Wfile = "NUMFAC_FAL" Then
 Call NUMFAC_FAL
End If
If Wfile = "CONTROL_STOCK" Then
 Call CONTROL_STOCK
End If
If Wfile = "VENTAS_ARTICULO" Then
 Call VENTAS_ARTICULO
End If
If Wfile = "CONSO_FAC" Then
 Call CONSO_FAC
End If
If Wfile = "DETALLE_OP" Then
 Call DETALLE_OP
End If
If Wfile = "CONTROL_MOVI" Then
 Call CONTROL_MOVI
End If
If Wfile = "PROGRA" Then
 Call PROGRA
End If
If Wfile = "PROGRA_TEJ" Then
 Call PROGRA_TEJ
End If
If Wfile = "SALDO_CTA" Then
 Call SALDO_CTA
End If
If Wfile = "REG_COMPRA" Then
 Call REG_COMPRA
End If

Exit Sub
SALE:
ProgBar.Visible = False
lblproceso.Visible = False
Pantalla.Enabled = True
cerrar.Enabled = True
MsgBox Err.Description + "Intente Nuevamente.", 48, Pub_Titulo
End Sub

Private Sub serie_KeyPress(KeyAscii As Integer)
SOLO_ENTERO KeyAscii
If KeyAscii = 13 Then
  Azul numini, numini
End If

End Sub

Private Sub txtCampo1_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
 Exit Sub
End If

If txtcampo2.Visible Then
 If Not IsDate(txtcampo2) Then
   txtcampo2.Text = Format(txtCampo1.Text, "dd/mm/yyyy")
 End If
 Azul2 txtcampo2, txtcampo2
Else
Pantalla.SetFocus
End If
 

End Sub

Private Sub txtcampo2_GotFocus()
'Azul txtCampo2, txtCampo2
End Sub

Private Sub txtCampo2_KeyPress(KeyAscii As Integer)
If KeyAscii <> 13 Then
 Exit Sub
End If
If KeyAscii = 13 And Wfile = "CONSO_FAC" Then
  cmbmoneda.SetFocus
  Exit Sub
End If
If txt_cli.Visible Then
 txt_cli.SetFocus
 Exit Sub
End If
If Pantalla.Enabled Then
   Pantalla.SetFocus
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
        cont.AddItem tab_mayor!tab_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
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
        cont.AddItem tab_mayor!tab_NOMLARGO & String(60, " ") & tab_mayor!TAB_NUMTAB
        tab_mayor.MoveNext
    Loop
End Sub

Public Sub CAJA_GRIFO()
'On Error GoTo FINTODO
Dim CAPTURA_FILA As Integer
Dim wsFECHA1
Dim wsFECHA2
Dim ws_clave

Dim suma_v_contado As Currency
Dim suma_c_contado As Currency
Dim suma_v_credito As Currency
Dim suma_c_credito As Currency
Dim suma_v_auto As Currency
Dim suma_c_auto As Currency
Dim suma_v_total As Currency
Dim suma_c_total As Currency

Dim TT_VOUCHER As Currency

Dim wvalor As Currency
Dim wcanti As Currency

Dim TT_TOTAL_VENTAS As Currency
Dim TT_TOTAL_CREDITO As Currency
Dim TT_TOTAL_AUTOCONSUMO As Currency
Dim TT_TOTAL_CONTADO As Currency

Dim WS_VENTAS As Currency
Dim WS_CREDITO As Currency
Dim WS_AUTOCONSUMO As Currency
Dim WS_TOTAL_CONTADO As Currency
Dim WS_COBRANZA As Currency
Dim WS_TARJETA As Currency
Dim WS_GASTOS As Currency
Dim WS_EFECTIVO As Currency
Dim TT_COSTO_VENTAS As Currency
Dim TT_TOTAL_BRUTO As Currency
Dim TT_TOTAL_GASTOS As Currency
Dim TT_TOTAL_OTROS As Currency
Dim TT_COBRANZA_CLI As Currency

TT_COBRANZA_CLI = 0
TT_TOTAL_BRUTO = 0
TT_TOTAL_GASTOS = 0
TT_TOTAL_OTROS = 0
    
TT_TOTAL_VENTAS = 0
TT_TOTAL_CREDITO = 0
TT_TOTAL_AUTOCONSUMO = 0
    
Pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtcampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtcampo2.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If Not IsDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
GoSub WEXCEL
pub_cadena = ""
xcuenta = 0

Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp3.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = PUB_CLAVE
f1 = 6  'Fila Inicial

'WCONTROL = WCONTROL + 1
pub_cadena = "SELECT FAR_SUBTOTAL, FAR_CANTIDAD, FAR_EQUIV, FAR_PRECIO, FAR_COSPRO, FAR_CODART, FAR_SIGNO_CAR, FAR_TIPDOC FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND  FAR_CODART = ? AND FAR_TIPMOV = 10 AND FAR_ESTADO = 'N'  ORDER BY FAR_TIPMOV, FAR_FBG DESC , FAR_FECHA, FAR_NUMSER,FAR_NUMFAC"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


' el PS_REP1(0) ESTA MAS ABAJO
DoEvents
FrmImp3.lblproceso.Visible = True
FrmImp3.ProgBar.Visible = True
FrmImp3.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
PUB_KEY = 0
SQ_OPER = 2
pu_codcia = LK_CODCIA
If LK_EMP_PTO = "A" Then
  pu_codcia = "00"
End If
LEER_ART_LLAVE
If art_mayor.EOF Then
 MsgBox "No Exieten Productos ..", 48, Pub_Titulo
 Exit Sub
End If
FrmImp3.lblproceso.Caption = "Procesando . . . "
DoEvents
FrmImp3.ProgBar.Min = 0
FrmImp3.ProgBar.Value = 0
FrmImp3.ProgBar.max = art_mayor.RowCount
TT_COSTO_VENTAS = 0
TT_VOUCHER = 0
Do Until art_mayor.EOF
 If art_mayor!ART_KEY <> 0 Then GoTo OTRO_ARTI
 If art_mayor!art_familia <> 1 Then GoTo OTRO_ARTI
 FrmImp3.ProgBar.Value = FrmImp3.ProgBar.Value + 1
 PS_REP02(0) = LK_CODCIA
 PS_REP02(1) = wsFECHA1
 PS_REP02(2) = wsFECHA2
 PS_REP02(3) = art_mayor!ART_KEY
 PUB_KEY = art_mayor!ART_KEY
 llave_rep02.Requery
 If llave_rep02.EOF Then
   GoTo OTRO_ARTI
 End If
 GoSub PROFACART
 GoSub IMPFACART
OTRO_ARTI:
 art_mayor.MoveNext
Loop
TT_TOTAL_VENTAS = 0
TT_TOTAL_CREDITO = 0
TT_TOTAL_AUTOCONSUMO = 0
TT_TOTAL_CONTADO = 0

WS_VENTAS = 0
WS_CREDITO = 0
WS_AUTOCONSUMO = 0
WS_TOTAL_CONTADO = 0
WS_COBRANZA = 0
WS_TARJETA = 0
WS_GASTOS = 0
WS_EFECTIVO = 0

 wran1 = "C" & 6
 wran2 = "C" & f1
 wranF = "C" & f1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 wran1 = "D" & 6
 wran2 = "D" & f1
 wranF = "D" & f1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 TT_TOTAL_VENTAS = TT_TOTAL_VENTAS + val(xl.Range(wranF))
 TT_TOTAL_CONTADO = TT_TOTAL_CONTADO + val(xl.Range(wranF))

 wran1 = "E" & 6
 wran2 = "E" & f1
 wranF = "E" & f1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 
 wran1 = "F" & 6
 wran2 = "F" & f1
 wranF = "F" & f1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 TT_TOTAL_VENTAS = TT_TOTAL_VENTAS + val(xl.Range(wranF))
 TT_TOTAL_CREDITO = TT_TOTAL_CREDITO + val(xl.Range(wranF))
 
 wran1 = "G" & 6
 wran2 = "G" & f1
 wranF = "G" & f1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 wran1 = "H" & 6
 wran2 = "H" & f1
 wranF = "H" & f1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 TT_TOTAL_VENTAS = TT_TOTAL_VENTAS + val(xl.Range(wranF))
 TT_TOTAL_AUTOCONSUMO = TT_TOTAL_AUTOCONSUMO + val(xl.Range(wranF))
 
 wran1 = "I" & 6
 wran2 = "I" & f1
 wranF = "I" & f1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 wran1 = "J" & 6
 wran2 = "J" & f1
 wranF = "J" & f1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
   
 wranF = "A" & f1 + 1 & ":J" & f1 + 1
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
 wranF = "A" & f1 + 2 & ":J" & f1 + 2
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 6
 
pub_cadena = "SELECT ALL_IMPORTE, ALL_CODCLIE FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_DIA  >= ? AND ALL_FECHA_DIA <= ? AND  ALL_SIGNO_CAJA = 1 AND ALL_FLAG_EXT <> 'E' AND (ALL_CODTRA = 2770 OR ALL_CODTRA = 2725)  ORDER BY ALL_CODCLIE "
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

GoSub COBRANZA_CLI
 
 'Reporte de Resumen
 '------------------
    

    f1 = f1 + 4
    CAPTURA_FILA = f1
    xl.Cells(f1, 2) = "RESUMEN"
    f1 = f1 + 1
    xl.Cells(f1, 3) = "S/."
    f1 = f1 + 1
    xl.Cells(f1, 2) = "TOTAL VENTAS"
    xl.Cells(f1, 3) = TT_TOTAL_VENTAS
    f1 = f1 + 1
    TT_TOTAL_CREDITO = TT_TOTAL_CREDITO - TT_VOUCHER
    xl.Cells(f1, 2) = "'(-) VENTAS CREDITO"
    xl.Cells(f1, 3) = TT_TOTAL_CREDITO
    f1 = f1 + 1
    xl.Cells(f1, 2) = "'(-) TARJ. DE CREDITO"
    xl.Cells(f1, 3) = TT_VOUCHER
    f1 = f1 + 1
    xl.Cells(f1, 2) = "'(-) AUTO CONSUMO"
    xl.Cells(f1, 3) = TT_TOTAL_AUTOCONSUMO
    f1 = f1 + 2
    TT_TOTAL_CONTADO = TT_TOTAL_VENTAS - TT_TOTAL_CREDITO - TT_TOTAL_AUTOCONSUMO - TT_VOUCHER
    xl.Cells(f1, 2) = "'TOTAL VTAS. CONTADO"
    xl.Cells(f1, 3) = TT_TOTAL_CONTADO
    f1 = f1 + 3
    xl.Cells(f1, 2) = "LIQUIDACION DE EFECTIVO"
    f1 = f1 + 2
    xl.Cells(f1, 2) = "TOT.VTAS.CONTADO"
    xl.Cells(f1, 3) = TT_TOTAL_CONTADO
    WS_EFECTIVO = WS_EFECTIVO + TT_TOTAL_CONTADO
    f1 = f1 + 1
    xl.Cells(f1, 2) = "COBRANZA CLIENTE"
    xl.Cells(f1, 3) = TT_COBRANZA_CLI
    WS_EFECTIVO = WS_EFECTIVO + TT_COBRANZA_CLI
    f1 = f1 + 1
    xl.Cells(f1, 2) = "'(-) GASTOS VARIOS"
    xl.Cells(f1, 3) = WS_GASTOS
    WS_EFECTIVO = WS_EFECTIVO - WS_GASTOS
    f1 = f1 + 3
    xl.Cells(f1, 2) = "'TOTAL EFECTIVO"
    xl.Cells(f1, 3) = WS_EFECTIVO
 
 ' ESTADO DE GANANCIAS Y PERIDAS
    f1 = CAPTURA_FILA
    xl.Cells(f1, 7) = "ESTADO DE GANANCIAS Y PERDIDAS"
    f1 = f1 + 1
    xl.Cells(f1, 9) = "S/."
    f1 = f1 + 1
    xl.Cells(f1, 7) = "VENTAS NETAS"
    xl.Cells(f1, 9) = TT_TOTAL_VENTAS
    f1 = f1 + 1
    xl.Cells(f1, 7) = "'(-) CTO. DE VENTAS"
    xl.Cells(f1, 9) = TT_COSTO_VENTAS
    TT_TOTAL_BRUTO = TT_TOTAL_VENTAS - TT_COSTO_VENTAS
    f1 = f1 + 1
    xl.Cells(f1, 7) = "'UTILIDAD BRUTA"
    xl.Cells(f1, 9) = TT_TOTAL_BRUTO
    
    f1 = f1 + 2
    xl.Cells(f1, 7) = "'GASTOS"
    xl.Cells(f1, 9) = TT_TOTAL_GASTOS
    f1 = f1 + 3
    xl.Cells(f1, 7) = "UTILIDAD DE OPERACION"
    xl.Cells(f1, 9) = TT_TOTAL_BRUTO - TT_TOTAL_GASTOS
    f1 = f1 + 1
    xl.Cells(f1, 7) = "OTROS INGRESOS"
    xl.Cells(f1, 9) = TT_TOTAL_OTROS
    f1 = f1 + 1
    xl.Cells(f1, 9) = "S/."
    f1 = f1 + 1
    xl.Cells(f1, 7) = "UTLIDAD NETA"
    xl.Cells(f1, 9) = (TT_TOTAL_BRUTO - TT_TOTAL_GASTOS) - TT_TOTAL_OTROS
 
  FrmImp3.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
  xl.Cells(2, 1) = Trim(retra_llave!TRA_DESCRIPCION)
  xl.Cells(3, 1) = "'" & Format(wsFECHA1, "dd/mm/yyyy") & " al " & Format(wsFECHA2, "dd/mm/yyyy")
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.APPLICATION.Visible = True
  DoEvents
  FrmImp3.lblproceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp3.Pantalla.Enabled = True
  FrmImp3.cerrar.Enabled = True
  FrmImp3.Pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblproceso.Visible = False

Exit Sub

PROFACART:
suma_v_total = 0
suma_c_total = 0
suma_v_contado = 0
suma_c_contado = 0
suma_v_credito = 0
suma_c_credito = 0
suma_v_auto = 0
suma_c_auto = 0

Do Until llave_rep02.EOF
    If val(llave_rep02!FAR_equiv) = 0 Then
      MsgBox "Equivalencia 0 Reset "
      GoTo CANCELA
    End If
    wvalor = val(llave_rep02!FAR_SUBTOTAL)
    wcanti = redondea((val(llave_rep02!far_cantidad)) / val(llave_rep02!FAR_equiv))
    If llave_rep02!far_signo_car = 0 And llave_rep02!FAR_TIPDOC <> "AU" Then
      suma_v_contado = suma_v_contado + wvalor
      suma_c_contado = suma_c_contado + wcanti
    End If
    If llave_rep02!far_signo_car = 1 And llave_rep02!FAR_TIPDOC <> "AU" Then   ' Or llave_rep02!FAR_TIPDOC = "VO" Then
      suma_v_credito = suma_v_credito + wvalor
      suma_c_credito = suma_c_credito + wcanti
    End If
    If llave_rep02!far_signo_car = 1 And llave_rep02!FAR_TIPDOC = "AU" Then
      suma_v_auto = suma_v_auto + wvalor
      suma_c_auto = suma_c_auto + wcanti
    End If
    If llave_rep02!far_signo_car = 1 And llave_rep02!FAR_TIPDOC = "VO" Then
      TT_VOUCHER = TT_VOUCHER + wvalor
    End If
    suma_v_total = suma_v_total + wvalor ' (suma_v_contado + suma_v_credito + suma_v_auto)
    suma_c_total = suma_c_total + wcanti '(suma_c_contado + suma_c_credito + suma_c_auto)
    TT_COSTO_VENTAS = TT_COSTO_VENTAS + redondea((val(llave_rep02!far_cantidad) * val(llave_rep02!FAR_COSPRO)))
  llave_rep02.MoveNext
 Loop
 
 
Return
  
 
Exit Sub

IMPFACART:
       f1 = f1 + 1
       PUB_KEY = PUB_KEY
       SQ_OPER = 1
       pu_codcia = LK_CODCIA
       LEER_ART_LLAVE
       xl.Cells(f1, 2) = art_LLAVE!art_nombre
       xl.Cells(f1, 3) = suma_c_contado
       xl.Cells(f1, 4) = suma_v_contado
       xl.Cells(f1, 5) = suma_c_credito
       xl.Cells(f1, 6) = suma_v_credito
       xl.Cells(f1, 7) = suma_c_auto
       xl.Cells(f1, 8) = suma_v_auto
       xl.Cells(f1, 9) = suma_c_total
       xl.Cells(f1, 10) = suma_v_total
Return

COBRANZA_CLI:
 PS_REP02(0) = LK_CODCIA
 PS_REP02(1) = wsFECHA1
 PS_REP02(2) = wsFECHA2
 llave_rep02.Requery
 Do Until llave_rep02.EOF
    TT_COBRANZA_CLI = TT_COBRANZA_CLI + val(llave_rep02!ALL_IMPORTE)
  llave_rep02.MoveNext
 Loop


Return

CANCELA:
  FrmImp3.Pantalla.Enabled = True
  FrmImp3.Pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblproceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  lblproceso.Caption = "Abriendo , Archivo REGVENTA.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\GRIFOS\CAJAGEN.xls", 0, True, 4, WPAS, WPAS
Return



FINTODO:
 MsgBox Err.Description & " .-  Reintente Nuevamente ..", 48, Pub_Titulo
  Resume Next
 GoTo CANCELA
 Resume Next
End Sub

Public Sub CONTROL_MOVI()
'On Error GoTo FINTODO
Dim CAPTURA_FILA As Integer
Dim wsFECHA1
Dim wsFECHA2
Dim ws_clave
Dim TT_VOUCHER As Currency
Dim wvalor As Currency
Dim wcanti As Currency
Dim suma_total As Currency
If Not IsDate(txtCampo1.Text) Then
   MsgBox "Fecha Invalida", 48, Pub_Titulo
   Azul2 txtCampo1, txtCampo1
   Exit Sub
End If
Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
pub_cadena = "SELECT FAR_CODART FROM FACART  WHERE FAR_CODCIA = ? AND (FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ?) AND  FAR_CODART = ? AND FAR_TIPMOV = 10 AND FAR_ESTADO <> 'E'"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02.MaxRows = 1
PS_REP02(0) = ""
PS_REP02(1) = LK_FECHA_DIA
PS_REP02(2) = LK_FECHA_DIA
PS_REP02(3) = 0
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
GoSub WEXCEL

SQ_OPER = 2
PUB_KEY = 0
pu_codcia = LK_CODCIA
LEER_ART_LLAVE
FrmImp3.ProgBar.Visible = True
FrmImp3.ProgBar.Min = 0
FrmImp3.ProgBar.Value = 0
FrmImp3.ProgBar.max = art_mayor.RowCount
suma_total = 0
PS_REP02(0) = LK_CODCIA
PS_REP02(1) = txtCampo1.Text
PS_REP02(2) = LK_FECHA_DIA
f1 = 5
Do Until art_mayor.EOF
  FrmImp3.ProgBar.Value = FrmImp3.ProgBar.Value + 1
  DoEvents
  PS_REP02(3) = val(art_mayor!ART_KEY)
  llave_rep02.Requery
  If llave_rep02.EOF Then
    PUB_CODART = art_mayor!ART_KEY
    PUB_CODCIA = LK_CODCIA
    SQ_OPER = 1
    LEER_ARM_LLAVE
    If val(arm_llave!arm_stock) <> 0 Then
      f1 = f1 + 1
      xl.Cells(f1, 1) = Trim(art_mayor!ART_ALTERNO)
      xl.Cells(f1, 2) = Trim(art_mayor!art_nombre)
      xl.Cells(f1, 3) = val(arm_llave!arm_stock)
      suma_total = suma_total + val(arm_llave!arm_stock)
    End If
  End If
art_mayor.MoveNext
Loop
f1 = f1 + 1
  xl.Cells(f1, 2) = "Total Kg.="
  xl.Cells(f1, 3) = suma_total
  
  xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
  xl.Cells(4, 1) = "'DESDE EL  " & Format(txtCampo1.Text, "dd/mm/yyyy") & " al " & Format(LK_FECHA_DIA, "dd/mm/yyyy")
    Pantalla.Enabled = True
    cerrar.Enabled = True
    DoEvents

  xl.DisplayAlerts = False
  FrmImp3.ProgBar.Visible = False
  xl.Worksheets(1).Protect PUB_CLAVE
  xl.APPLICATION.Visible = True
  Set xl = Nothing

Exit Sub
WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  lblproceso.Caption = "Abriendo , Archivo CONTROL.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open PUB_RUTA_OTRO + "\CONTROL.xls", 0, True, 4, WPAS, WPAS
Return

FINTODO:
 MsgBox Err.Description & " .-  Reintente Nuevamente ..", 48, Pub_Titulo

End Sub




Public Sub NUMFAC_FAL()
'On Error GoTo FINTODO
Dim ws_clave
Dim wvalor As Currency
Dim wcanti As Currency
Dim WW_CORRELA As Currency
Dim tfechaini
Dim tfechafin
Dim ip_numfac
Dim ip_fecha
Dim wwnumfac
Dim WFLAG As String * 1
If Trim(fbg.Text) = "" Then
  MsgBox "Ingrese Tipo de Documento.", 48, Pub_Titulo
  Exit Sub
End If
If val(serie.Text) <= 0 Then
   MsgBox "Ingrese Serie del Documento.", 48, Pub_Titulo
   Azul serie, serie
   Exit Sub
End If
If val(numini.Text) <= 0 Then
   MsgBox "Ingrese N° inicial del Documento.", 48, Pub_Titulo
   Azul numini, numini
   Exit Sub
End If
'If Val(numfin.Text) <= 0 Then
'   MsgBox "Ingrese N° Final del Documento.", 48, Pub_Titulo
'   Azul numfin, numfin
'   Exit Sub
'End If

    
Pantalla.Enabled = False
cerrar.Enabled = False
pub_cadena = ""
xcuenta = 0
Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp3.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = PUB_CLAVE
f1 = 9  'Fila Inicial
WFLAG = ""

'WCONTROL = WCONTROL + 1
pub_cadena = "SELECT Distinct FAR_NUMFAC, FAR_FECHA_COMPRA FROM FACART WHERE FAR_CODCIA = ? AND FAR_FBG = ? AND FAR_NUMSER = ? AND FAR_NUMFAC >= ? AND FAR_TIPMOV = 10  ORDER BY  FAR_NUMFAC"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = 0
PS_REP02(1) = 0
PS_REP02(2) = 0
PS_REP02(3) = 0
'PS_REP02(4) = 0
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
PS_REP02(0) = LK_CODCIA
PS_REP02(1) = fbg.Text
PS_REP02(2) = serie.Text
PS_REP02(3) = numini.Text
'PS_REP02(4) = numfin.Text
llave_rep02.Requery
If llave_rep02.EOF Then
   MsgBox "En este rango no Existe documentos ", 48, Pub_Titulo
   GoTo CANCELA
End If

' el PS_REP1(0) ESTA MAS ABAJO
DoEvents
FrmImp3.lblproceso.Visible = True
FrmImp3.ProgBar.Visible = True
FrmImp3.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
FrmImp3.lblproceso.Caption = "Procesando . . . "
DoEvents
FrmImp3.ProgBar.Min = 0
FrmImp3.ProgBar.Value = 0

 
 FrmImp3.ProgBar.max = llave_rep02.RowCount
 If Trim(fbg.Text) = "F" Then
   xl.Cells(5, 2) = "FACTURAS"
 Else
   xl.Cells(5, 2) = "BOLETAS"
 End If
 WW_CORRELA = llave_rep02!far_numfac ' Val(numini.Text)
 ip_numfac = llave_rep02!far_numfac 'Val(numini.Text)
 ip_fecha = Format(llave_rep02!FAR_fecha_compra, "dd/mm/yyyy")
 tfechaini = Format(llave_rep02!FAR_fecha_compra, "dd/mm/yyyy")
 wwnumfac = llave_rep02!far_numfac
 xl.Cells(6, 2) = serie.Text
 xl.Cells(7, 2) = ip_numfac 'numini.Text
 xl.Cells(8, 2) = tfechaini
 
 Do Until llave_rep02.EOF
    FrmImp3.ProgBar.Value = FrmImp3.ProgBar.Value + 1
    DoEvents
    ip_fecha = Format(llave_rep02!FAR_fecha_compra, "dd/mm/yyyy")
    If WW_CORRELA < llave_rep02!far_numfac Then
      For fila = 1 To (llave_rep02!far_numfac - WW_CORRELA)
      ip_numfac = WW_CORRELA + fila - 1
      GoSub IMP_NUMFAC
      Next fila
      WFLAG = "A"
    End If
    WW_CORRELA = llave_rep02!far_numfac
    WW_CORRELA = WW_CORRELA + 1
    numfin.Text = llave_rep02!far_numfac
    tfechafin = Format(llave_rep02!FAR_fecha_compra, "dd/mm/yyyy")
    llave_rep02.MoveNext
 Loop
 wranF = "A" & f1 + 1 & ":D" & f1 + 1
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
 f1 = f1 + 2
 xl.Cells(f1, 1) = "FIN DE LISTADO "
 If WFLAG <> "A" Then
   f1 = f1 + 1
   xl.Cells(f1, 1) = "*** LOS DOC. ESTAN CORRECTOS *** "
 End If
 f1 = f1 + 1
 xl.Cells(f1, 1) = "N° FINAL:"
 xl.Cells(f1, 2) = numfin.Text
 f1 = f1 + 1
 xl.Cells(f1, 1) = "FEC.FINAL:"
 xl.Cells(f1, 2) = "'" & tfechafin
 wranF = "A" & f1 + 1 & ":D" & f1 + 1
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
 
  
 
 'wranF = "A" & F1 + 2 & ":C" & F1 + 2
 'xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 6
 
 'Do Until llave_rep02.EOF
 '   ip_fecha = Format(llave_rep02!FAR_fecha_COMPRA, "dd/mm/yyyy")
 '   ip_numfac = WW_CORRELA
 '   If WW_CORRELA <> llave_rep02!FAR_numfac Then
 '     GoSub IMP_NUMFAC
 '     WW_CORRELA = WW_CORRELA + 1
 '   End If
 '   WW_CORRELA = WW_CORRELA + 1
 '   wwnumfac = llave_rep02!FAR_numfac
 '   tfechafin = Format(llave_rep02!FAR_fecha_COMPRA, "dd/mm/yyyy")
 '   llave_rep02.MoveNext
 'Loop
 
  FrmImp3.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  

  xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
  xl.Cells(2, 1) = "'" & Format(LK_FECHA_DIA, "dd/mm/yyyy")
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.APPLICATION.Visible = True
  DoEvents
  FrmImp3.lblproceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp3.Pantalla.Enabled = True
  FrmImp3.cerrar.Enabled = True
  FrmImp3.Pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblproceso.Visible = False

Exit Sub

IMP_NUMFAC:
       f1 = f1 + 1
       If Trim(fbg.Text) = "F" Then
           xl.Cells(f1, 1) = "FACTURAS"
       Else
           xl.Cells(f1, 1) = "BOLETAS"
       End If
       xl.Cells(f1, 2) = ip_numfac
       xl.Cells(f1, 3) = "FECHA PROBABLE:"
       xl.Cells(f1, 4) = "'" + ip_fecha
Return


CANCELA:
  FrmImp3.Pantalla.Enabled = True
  FrmImp3.Pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblproceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  lblproceso.Caption = "Abriendo , Archivo NUMDOC.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\NUMDOC.xls", 0, True, 4, WPAS, WPAS
Return



FINTODO:
 MsgBox Err.Description & " .-  Reintente Nuevamente ..", 48, Pub_Titulo
  Resume Next
 GoTo CANCELA
 Resume Next
End Sub


Public Sub CONTROL_STOCK()
'On Error GoTo FINTODO
Dim WDIFE As Currency
Dim WSTCKR As Currency
Dim stock_llave As rdoResultset
Dim PSST_LLAVE As rdoQuery
Dim LETRAS(50) As String * 2
Dim ACU_INGRESOS As Currency
Dim TOT_VENTA As Currency
Dim sum_turnos As Currency
Dim wturno As Integer
Dim d_numfac
Dim ff_flag
Dim WDOCU
Dim WSALDO As Currency
Dim ACU_SALDO_ARTI As Currency
Dim CAPTURA_FILA As Integer
Dim wsFECHA1
Dim wsFECHA2
Dim ws_clave
Dim flag_saldo_ini As String * 1
Dim SALDO_INI As Currency
Dim F001 As Integer
Dim max_islas As Integer
Dim fila_islas As Integer
Dim WISLA
Dim wconta As Integer
Pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
pub_cadena = ""
xcuenta = 0

Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp3.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = PUB_CLAVE

'WCONTROL = WCONTROL + 1
pub_cadena = "SELECT VEM_CODVEN FROM VEMAEST WHERE VEM_CODCIA = ? "
Set PS_REP04 = CN.CreateQuery("", pub_cadena)
PS_REP04(0) = 0
Set llave_rep04 = PS_REP04.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT far_NUMFAC_C, far_NUMser_C, far_numguia, FAR_NUMFAC, FAR_CANTIDAD, FAR_EQUIV, FAR_STOCK, FAR_SIGNO_ARM, FAR_CODART, FAR_SIGNO_CAR, FAR_TIPDOC FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA >= ? AND  FAR_CODART = ? AND (FAR_TIPMOV = ? OR FAR_TIPMOV = ?) AND FAR_ESTADO = 'N'  ORDER BY FAR_FECHA, FAR_NUMOPER, FAR_NUMSEC"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT FAR_ISLA, FAR_NUMFAC, FAR_CANTIDAD, FAR_EQUIV, FAR_STOCK, FAR_SIGNO_ARM, FAR_CODART, FAR_SIGNO_CAR, FAR_TIPDOC, FAR_TURNO  FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA >= ? AND  FAR_CODART = ? AND (FAR_TIPMOV = ? OR FAR_TIPMOV = ?) AND FAR_TURNO = ? AND FAR_ESTADO = 'N'   ORDER BY FAR_TURNO, FAR_ISLA"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


pub_cadena = "SELECT FAR_STOCK, FAR_EQUIV, FAR_SIGNO_ARM, FAR_CANTIDAD  FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA >= ? AND  FAR_CODART = ? AND FAR_ESTADO <> 'M'  ORDER BY FAR_FECHA, FAR_NUMOPER, FAR_NUMSEC "
Set PS_REP03 = CN.CreateQuery("", pub_cadena)
PS_REP03(0) = 0
PS_REP03(1) = LK_FECHA_DIA
PS_REP03(2) = 0
PS_REP03.MaxRows = 1
Set llave_rep03 = PS_REP03.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT * FROM TABLAS WHERE TAB_CODCIA = ? AND TAB_TIPREG = ?  AND TAB_NOMCORTO = ? AND TAB_CODCLIE = ? AND TAB_CODART = ?  ORDER BY TAB_NUMTAB "
Set PSST_LLAVE = CN.CreateQuery("", pub_cadena)
PSST_LLAVE(0) = 0
PSST_LLAVE(1) = 0
PSST_LLAVE(2) = 0
PSST_LLAVE(3) = 0
PSST_LLAVE(4) = 0
Set stock_llave = PSST_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)


'el PS_REP1(0) ESTA MAS ABAJO
GoSub WEXCEL

DoEvents
FrmImp3.lblproceso.Visible = True
FrmImp3.ProgBar.Visible = True
FrmImp3.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
PUB_KEY = 0
SQ_OPER = 2
pu_codcia = LK_CODCIA
LEER_ART_LLAVE

If art_mayor.EOF Then
 MsgBox "No Exieten Productos ..", 48, Pub_Titulo
 Exit Sub
End If
PS_REP04(0) = LK_CODCIA
llave_rep04.Requery
If llave_rep04.EOF Then
 MsgBox "Islas No Existen ", 48, Pub_Titulo
 GoTo CANCELA
End If
max_islas = llave_rep04.RowCount
FrmImp3.lblproceso.Caption = "Procesando . . . "
DoEvents
FrmImp3.ProgBar.Min = 0
FrmImp3.ProgBar.Value = 0
FrmImp3.ProgBar.max = art_mayor.RowCount
GoSub LETRAS
c1 = 1
F001 = 0
CAPTURA_FILA = 0
xl.Cells(7, 1) = "SALDO INICIAL "
F001 = 0
TOT_VENTA = 0
Do Until art_mayor.EOF
 If art_mayor!ART_KEY = 0 Then GoTo OTRO_ARTI
 If art_mayor!art_familia <> 1 Then GoTo OTRO_ARTI
 c1 = c1 + 2
 f1 = 6  'Fila Inicial
 flag_saldo_ini = ""
 ACU_SALDO_ARTI = 0
 FrmImp3.ProgBar.Value = FrmImp3.ProgBar.Value + 1
 PS_REP03(0) = LK_CODCIA
 PS_REP03(1) = wsFECHA1
 PS_REP03(2) = art_mayor!ART_KEY
 llave_rep03.Requery
 SALDO_INI = 0
 ACU_INGRESOS = 0
 
 If Not llave_rep03.EOF Then
    SALDO_INI = (llave_rep03!FAR_STOCK / llave_rep03!FAR_equiv) + (llave_rep03!far_signo_arm * (llave_rep03!far_cantidad / llave_rep03!FAR_equiv) * -1)
 Else
    SQ_OPER = 1
    PUB_CODART = art_mayor!ART_KEY
    pu_codcia = LK_CODCIA
    LEER_ARM_LLAVE
    SALDO_INI = arm_llave!arm_stock
 End If
 xl.Cells(f1, c1) = art_mayor!art_nombre
 xl.Cells(f1, c1 - 1) = "ISLA"
 
 f1 = f1 + 1
 xl.Cells(f1, c1) = SALDO_INI
 ACU_SALDO_ARTI = ACU_SALDO_ARTI + SALDO_INI
 wranF = "A" & f1 & ":" & Trim(LETRAS(c1)) & f1
 xl.Range(wranF).Borders.LineStyle = 1
 wranF = "A" & f1 - 1 & ":" & Trim(LETRAS(c1)) & f1 - 1
 xl.Range(wranF).Borders.LineStyle = 1
 
 xl.Cells(f1, c1).Font.Bold = True
 
 PS_REP02(0) = LK_CODCIA
 PS_REP02(1) = wsFECHA1
 PS_REP02(2) = art_mayor!ART_KEY
 PS_REP02(3) = 20
 PS_REP02(4) = 20
 llave_rep02.Requery
 If Not llave_rep02.EOF Then
   f1 = f1 + 1
   GoSub PROC_INGRESOS
 End If
 PS_REP01(0) = LK_CODCIA
 PS_REP01(1) = wsFECHA1
 PS_REP01(2) = art_mayor!ART_KEY
 PS_REP01(3) = 10
 PS_REP01(4) = 10
 SQ_OPER = 2
 PUB_TIPREG = 2102
 PUB_CODCIA = LK_CODCIA
 LEER_TAB_LLAVE
 f1 = f1 + 1
 If F001 = 0 Then F001 = f1
 xl.Cells(F001 + 1, c1) = ACU_INGRESOS
 xl.Cells(F001 + 1, 1) = "TOTAL INGRESOS "
 xl.Cells(F001 + 1, 1).Font.Bold = True
 xl.Cells(F001 + 1, c1).Font.Bold = True
 wranF = "A" & F001 + 1 & ":" & Trim(LETRAS(c1)) & F001 + 1
 xl.Range(wranF).Borders.LineStyle = 1
 xl.Worksheets(1).Rows(F001 + 1).RowHeight = 18
 f1 = F001
 TOT_VENTA = 0
 Do Until tab_mayor.EOF
    max_islas = llave_rep04.RowCount + f1 + 2
    WDOCU = "TURNO: " & tab_mayor!TAB_NUMTAB
    f1 = f1 + 1
    GoSub PROC_TURNOS
    wturno = tab_mayor!TAB_NUMTAB
    tab_mayor.MoveNext
 Loop
 xl.Cells(max_islas + 1, c1) = TOT_VENTA
 xl.Cells(max_islas + 1, c1).Font.Bold = True
 xl.Cells(max_islas + 1, 1) = "TOTAL VENTA "
 xl.Cells(max_islas + 1, 1).Font.Bold = True
 wranF = "A" & max_islas + 1 & ":" & Trim(LETRAS(c1)) & max_islas + 1
 xl.Range(wranF).Borders.LineStyle = 1
 
 xl.Worksheets(1).Rows(max_islas + 1).RowHeight = 18
 xl.Cells(max_islas + 2, c1) = ACU_SALDO_ARTI
 xl.Cells(max_islas + 2, c1).Font.Bold = True
 xl.Cells(max_islas + 2, 1) = "STOCK ACTUAL "
 xl.Cells(max_islas + 2, 1).Font.Bold = True
 'STOCK POR REGLA
 PSST_LLAVE(0) = LK_CODCIA
 PSST_LLAVE(1) = 2110
 PSST_LLAVE(2) = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 PSST_LLAVE(3) = wturno
 PSST_LLAVE(4) = art_mayor!ART_KEY
 stock_llave.Requery
 If Not stock_llave.EOF Then
  WSTCKR = stock_llave!tab_NOMLARGO
 Else
  WSTCKR = 0
 End If
 WDIFE = ACU_SALDO_ARTI - WSTCKR
 art_mayor.Edit
 art_mayor!art_margen = val(WDIFE)
 art_mayor.Update
 xl.Cells(max_islas + 3, c1) = WSTCKR
 xl.Cells(max_islas + 3, c1).Font.Bold = True
 xl.Cells(max_islas + 3, 1) = "STOCK POR REGLA "
 xl.Cells(max_islas + 3, 1).Font.Bold = True
 xl.Cells(max_islas + 4, c1) = WDIFE
 xl.Cells(max_islas + 4, c1).Font.Bold = True
 xl.Cells(max_islas + 4, 1) = "DIFERENCIA: "
 xl.Cells(max_islas + 4, 1).Font.Bold = True
 
 
 wranF = "A" & max_islas + 2 & ":" & Trim(LETRAS(c1)) & max_islas + 2
 xl.Range(wranF).Borders.LineStyle = 1
 
 xl.Worksheets(1).Rows(max_islas + 3).RowHeight = 18
 
 If f1 > CAPTURA_FILA Then CAPTURA_FILA = f1
OTRO_ARTI:
 art_mayor.MoveNext
Loop
  wranF = "A6:" & Trim(LETRAS(c1)) & 6
  xl.Range(wranF).Borders.Item(3).LineStyle = 1
  wranF = "A6:A" & f1 + 3
  xl.Range(wranF).Borders.Item(7).LineStyle = 1
  wranF = Trim(LETRAS(c1)) & 6 & ":" & Trim(LETRAS(c1)) & f1 + 3
  xl.Range(wranF).Borders.Item(10).LineStyle = 1
 

  FrmImp3.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
  xl.Cells(2, 1) = Trim(retra_llave!TRA_DESCRIPCION)
  xl.Cells(3, 1) = "'DEL " & Format(wsFECHA1, "dd/mm/yyyy") & " al " & Format(LK_FECHA_DIA, "dd/mm/yyyy")
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.APPLICATION.Visible = True
  DoEvents
  FrmImp3.lblproceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp3.Pantalla.Enabled = True
  FrmImp3.cerrar.Enabled = True
  FrmImp3.Pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblproceso.Visible = False

Exit Sub

PROC_INGRESOS:
WSALDO = 0
d_numfac = llave_rep02!far_numfac
ff_flag = ""
WDOCU = ""
Do Until llave_rep02.EOF
    If d_numfac <> llave_rep02!far_numfac Then
       f1 = f1 + 1
       xl.Cells(f1, 1) = WDOCU
       xl.Cells(f1, c1) = WSALDO
       d_numfac = llave_rep02!far_numfac
       ACU_INGRESOS = ACU_INGRESOS + WSALDO
       ACU_SALDO_ARTI = ACU_SALDO_ARTI + WSALDO
       WSALDO = 0
    End If
    WSALDO = WSALDO + (val(llave_rep02!far_cantidad) / val(llave_rep02!FAR_equiv))
    If val(llave_rep02!FAR_NUMFAC_C) <> 0 Then
      WDOCU = "FAC.: " & llave_rep02!FAR_NUMSER_C & " - " & llave_rep02!FAR_NUMFAC_C
    Else
     WDOCU = "GUIA.: " & llave_rep02!far_numguia
    End If
    ff_flag = "A"
    llave_rep02.MoveNext
Loop

If ff_flag = "A" Then
   f1 = f1 + 1
   xl.Cells(f1, 1) = WDOCU
   xl.Cells(f1, c1) = WSALDO
   ACU_INGRESOS = ACU_INGRESOS + WSALDO
   ACU_SALDO_ARTI = ACU_SALDO_ARTI + WSALDO
   WSALDO = 0
End If

 
Return

PROC_TURNOS:
sum_turnos = 0
wconta = 0
ff_flag = ""
WSALDO = 0
If fila_islas = 0 Then fila_islas = f1
PS_REP01(5) = tab_mayor!TAB_NUMTAB
llave_rep01.Requery
If Not llave_rep01.EOF Then
 d_numfac = llave_rep01!far_ISLA
End If
Do Until llave_rep01.EOF
   If (d_numfac <> llave_rep01!far_ISLA) Then
       f1 = f1 + 1
       xl.Cells(f1, c1 - 1) = WISLA
       xl.Cells(f1, c1) = WSALDO
       d_numfac = llave_rep01!far_numfac
       sum_turnos = sum_turnos + WSALDO
       WSALDO = 0
       wconta = wconta + 1
   End If
   WISLA = llave_rep01!far_ISLA
   WSALDO = WSALDO + (val(llave_rep01!far_cantidad) / val(llave_rep01!FAR_equiv))
   ff_flag = "A"
   llave_rep01.MoveNext
Loop
If ff_flag = "A" Then
   f1 = f1 + 1
   xl.Cells(f1, c1 - 1) = WISLA
   xl.Cells(f1, c1) = WSALDO
   sum_turnos = sum_turnos + WSALDO
   wconta = wconta + 1
End If
xl.Cells(max_islas, 1) = WDOCU
xl.Cells(max_islas, c1) = sum_turnos
wranF = "A" & max_islas & ":" & Trim(LETRAS(c1)) & max_islas
xl.Range(wranF).Borders.LineStyle = 1

TOT_VENTA = TOT_VENTA + sum_turnos
ACU_SALDO_ARTI = ACU_SALDO_ARTI - sum_turnos
If wconta <> llave_rep04.RowCount Then
    f1 = f1 + (llave_rep04.RowCount - wconta)
End If

WSALDO = 0


Return
 
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
LETRAS(25) = "Y"
LETRAS(26) = "Z"
LETRAS(27) = "AA"
LETRAS(28) = "AB"
LETRAS(29) = "AC"
LETRAS(30) = "AD"
LETRAS(31) = "AE"
LETRAS(32) = "AF"
LETRAS(33) = "AG"
LETRAS(34) = "AH"
LETRAS(35) = "AI"
LETRAS(36) = "AJ"
LETRAS(37) = "AK"
LETRAS(38) = "AL"
LETRAS(39) = "AM"
LETRAS(40) = "AN"
LETRAS(41) = "AO"
LETRAS(42) = "AP"
LETRAS(43) = "AQ"
LETRAS(44) = "AR"
LETRAS(45) = "AS"
LETRAS(46) = "AT"
LETRAS(47) = "AU"
LETRAS(48) = "AV"
LETRAS(49) = "AW"
LETRAS(50) = "AX"
Return
CANCELA:
  FrmImp3.Pantalla.Enabled = True
  FrmImp3.Pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblproceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  lblproceso.Caption = "Abriendo , Archivo REGVENTA.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\GRIFOS\STOCK.xls", 0, True, 4, WPAS, WPAS
Return



FINTODO:
 MsgBox Err.Description & " .-  Reintente Nuevamente ..", 48, Pub_Titulo
  Resume Next
 GoTo CANCELA
 Resume Next
End Sub
Public Sub VENTAS_ARTICULO()
Dim pag_llave As rdoResultset
Dim ww_cant As Currency
Dim PSPAG_LLAVE As rdoQuery
'On Error GoTo FINTODO
Dim WQ_cANTIDAD As Currency
Dim WS_FILA As Integer
Dim ws_ultimo As Integer
Dim WS_FILA_ULTIMA As Integer
Dim wq_fecha As Date
Dim wq_codart As Long
Dim J As Integer
Dim ww As String
Dim LETRAS(24) As String * 1

pub_cadena = "SELECT * FROM PARGRI WHERE PAG_CODCIA = ?   AND PAG_FECHA = ?  AND PAG_CODART = ? "
Set PSPAG_LLAVE = CN.CreateQuery("", pub_cadena)

PSPAG_LLAVE(0) = 0
PSPAG_LLAVE(1) = 0
PSPAG_LLAVE(2) = 0

Set pag_llave = PSPAG_LLAVE.OpenResultset(rdOpenKeyset, rdConcurValues)


GoSub WEXCEL
GoSub LETRAS
Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp3.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
f1 = 5  'Fila Inicial

pub_cadena = "SELECT * FROM FACART WHERE FAR_TIPMOV = 10 AND FAR_CODCIA=?   AND FAR_FECHA >= ? AND FAR_FECHA<= ?  AND FAR_ESTADO <> 'E' ORDER BY FAR_FECHA, FAR_CODART "
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)



PUB_KEY = 0
pu_codcia = LK_CODCIA
SQ_OPER = 2
LEER_ART_LLAVE
If art_mayor.EOF = True Then
   MsgBox "NO HAY DATOS... "
   Exit Sub
End If
WS_FILA = 1
Do Until art_mayor.EOF
   If art_mayor!art_familia = 1 Then
      WS_FILA = WS_FILA + 1
      xl.Cells(1, WS_FILA) = art_mayor!ART_KEY
      xl.Cells(4, WS_FILA) = Trim(art_mayor!art_nombre)
   End If
   art_mayor.MoveNext
Loop
    wranF = "A4:" & LETRAS(WS_FILA) & "4"
    xl.Range(wranF).Borders.LineStyle = 9
    xl.APPLICATION.Range(wranF).Select
    
'    With Selection.Interior
'        .ColorIndex = 19
   '      .Pattern = xlSolid
     '      .PatternColorIndex = xlAutomatic
'    End With
ws_ultimo = WS_FILA

PS_REP02(0) = LK_CODCIA

mas:
ww = InputBox("Fecha Inicial")
If ww = "" Then GoTo mas
If IsDate(ww) = False Then GoTo mas
PS_REP02(1) = ww
ww = InputBox("Fecha Final")
If ww = "" Then GoTo mas
If IsDate(ww) = False Then GoTo mas
PS_REP02(2) = ww
llave_rep02.Requery
If llave_rep02.EOF Then GoTo SALIR

FrmImp3.ProgBar.max = llave_rep02.RowCount
FrmImp3.ProgBar.Visible = True
FrmImp3.ProgBar.Min = 0
DoEvents
WS_FILA = 5
wq_fecha = Format(llave_rep02!FAR_fecha, "dd/mmm/yy")
xl.Cells(WS_FILA, 1) = "'" & wq_fecha
wq_codart = llave_rep02!far_codart

Do Until llave_rep02.EOF
  
  If wq_fecha = llave_rep02!FAR_fecha And wq_codart = llave_rep02!far_codart Then
  Else
     GoSub escribe
     wq_fecha = llave_rep02!FAR_fecha
     wq_codart = llave_rep02!far_codart
  End If
  
  WQ_cANTIDAD = val(llave_rep02!far_cantidad) + WQ_cANTIDAD
'xl.Visible = True
  DoEvents
  FrmImp3.ProgBar.Value = FrmImp3.ProgBar.Value + 1
  llave_rep02.MoveNext
Loop
  GoSub escribe
WS_FILA_ULTIMA = WS_FILA
i = 5
J = 2
Do Until i > WS_FILA_ULTIMA
   Do Until J > ws_ultimo
   PSPAG_LLAVE(0) = LK_CODCIA
   PSPAG_LLAVE(1) = xl.Cells(i, 1)
   PSPAG_LLAVE(2) = xl.Cells(1, J)
   pag_llave.Requery
   If pag_llave.EOF Then MsgBox "ERROR ..."
   ww_cant = 0
   Do Until pag_llave.EOF
      ww_cant = pag_llave!PAG_LEC_CIERRE - pag_llave!PAG_LEC_INICIO + ww_cant
      pag_llave.MoveNext
   Loop
   xl.Cells(i, J) = ww_cant - val(xl.Cells(i, J))

   J = J + 1
   Loop
   i = i + 1
Loop
   
  
SALIR:
J = 2
Do Until J > ws_ultimo
   xl.Cells(1, J) = ""
   J = J + 1
Loop


  FrmImp3.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(2, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
  xl.Cells(2, 5) = Format(LK_FECHA_DIA, "dddd,dd-mmm-yy")
  xl.DisplayAlerts = False
  xl.APPLICATION.Visible = True
  DoEvents
  FrmImp3.lblproceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp3.Pantalla.Enabled = True
  FrmImp3.cerrar.Enabled = True
  FrmImp3.Pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblproceso.Visible = False
Exit Sub

 

CANCELA:
  FrmImp3.Pantalla.Enabled = True
  FrmImp3.Pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblproceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  lblproceso.Caption = "Abriendo , Archivo REGVENTA.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = ""
  xl.Workbooks.Open "c:\ADMIN\HERTISA\POR_FACTURAR.xls", 0, True
Return





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


Return
escribe:
J = 2
Do Until J > ws_ultimo
   If wq_codart = xl.Cells(1, J) Then
      xl.Cells(WS_FILA, J) = WQ_cANTIDAD
   End If
   J = J + 1
Loop

WQ_cANTIDAD = 0

If Not llave_rep02.EOF Then
If wq_fecha <> llave_rep02!FAR_fecha Then
   WS_FILA = WS_FILA + 1
   xl.Cells(WS_FILA, 1) = "'" & Format(llave_rep02!FAR_fecha, "dd/mmm/yy")
End If
End If
Return


FINTODO:
 MsgBox Err.Description & " .-  Reintente Nuevamente ..", 48, Pub_Titulo
 GoTo CANCELA
 Resume Next
End Sub

Public Sub CRED_CONT()
'On Error GoTo FINTODO
Dim CAPTURA_FILA As Integer
Dim wsFECHA1
Dim wsFECHA2
Dim ws_clave

Dim suma_v_contado As Currency
Dim suma_c_contado As Currency
Dim suma_v_credito As Currency
Dim suma_c_credito As Currency
Dim suma_v_auto As Currency
Dim suma_c_auto As Currency
Dim suma_v_total As Currency
Dim suma_c_total As Currency

Dim TT_VOUCHER As Currency

Dim wvalor As Currency
Dim wcanti As Currency

Dim TT_TOTAL_VENTAS As Currency
Dim TT_TOTAL_CREDITO As Currency
Dim TT_TOTAL_AUTOCONSUMO As Currency
Dim TT_TOTAL_CONTADO As Currency

Dim WS_VENTAS As Currency
Dim WS_CREDITO As Currency
Dim WS_AUTOCONSUMO As Currency
Dim WS_TOTAL_CONTADO As Currency
Dim WS_COBRANZA As Currency
Dim WS_TARJETA As Currency
Dim WS_GASTOS As Currency
Dim WS_EFECTIVO As Currency
Dim TT_COSTO_VENTAS As Currency
Dim TT_TOTAL_BRUTO As Currency
Dim TT_TOTAL_GASTOS As Currency
Dim TT_TOTAL_OTROS As Currency
Dim TT_COBRANZA_CLI As Currency

TT_COBRANZA_CLI = 0
TT_TOTAL_BRUTO = 0
TT_TOTAL_GASTOS = 0
TT_TOTAL_OTROS = 0
    
TT_TOTAL_VENTAS = 0
TT_TOTAL_CREDITO = 0
TT_TOTAL_AUTOCONSUMO = 0
    
Pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtcampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtcampo2.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If Not IsDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
GoSub WEXCEL
pub_cadena = ""
xcuenta = 0

Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp3.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = PUB_CLAVE
f1 = 6  'Fila Inicial

'WCONTROL = WCONTROL + 1
pub_cadena = "SELECT FAR_CANTIDAD, FAR_EQUIV, FAR_PRECIO, FAR_COSPRO, FAR_CODART, FAR_SIGNO_CAR, FAR_TIPDOC FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND  FAR_CODART = ? AND FAR_TIPMOV = 10 AND FAR_ESTADO = 'N'  ORDER BY FAR_TIPMOV, FAR_FBG DESC , FAR_FECHA, FAR_NUMSER,FAR_NUMFAC"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


' el PS_REP1(0) ESTA MAS ABAJO
DoEvents
FrmImp3.lblproceso.Visible = True
FrmImp3.ProgBar.Visible = True
FrmImp3.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
PUB_KEY = 0
SQ_OPER = 2
pu_codcia = LK_CODCIA
If LK_EMP_PTO = "A" Then
  pu_codcia = "00"
End If
LEER_ART_LLAVE
If art_mayor.EOF Then
 MsgBox "No Exieten Productos ..", 48, Pub_Titulo
 Exit Sub
End If
FrmImp3.lblproceso.Caption = "Procesando . . . "
DoEvents
FrmImp3.ProgBar.Min = 0
FrmImp3.ProgBar.Value = 0
FrmImp3.ProgBar.max = art_mayor.RowCount
TT_COSTO_VENTAS = 0
TT_VOUCHER = 0
Do Until art_mayor.EOF
 FrmImp3.ProgBar.Value = FrmImp3.ProgBar.Value + 1
 PS_REP02(0) = LK_CODCIA
 PS_REP02(1) = wsFECHA1
 PS_REP02(2) = wsFECHA2
 PS_REP02(3) = art_mayor!ART_KEY
 PUB_KEY = art_mayor!ART_KEY
 llave_rep02.Requery
 If llave_rep02.EOF Then
   GoTo OTRO_ARTI
 End If
 GoSub PROFACART
 GoSub IMPFACART
OTRO_ARTI:
 art_mayor.MoveNext
Loop
TT_TOTAL_VENTAS = 0
TT_TOTAL_CREDITO = 0
TT_TOTAL_AUTOCONSUMO = 0
TT_TOTAL_CONTADO = 0

WS_VENTAS = 0
WS_CREDITO = 0
WS_AUTOCONSUMO = 0
WS_TOTAL_CONTADO = 0
WS_COBRANZA = 0
WS_TARJETA = 0
WS_GASTOS = 0
WS_EFECTIVO = 0

 wran1 = "C" & 6
 wran2 = "C" & f1
 wranF = "C" & f1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 wran1 = "D" & 6
 wran2 = "D" & f1
 wranF = "D" & f1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 TT_TOTAL_VENTAS = TT_TOTAL_VENTAS + val(xl.Range(wranF))
 TT_TOTAL_CONTADO = TT_TOTAL_CONTADO + val(xl.Range(wranF))

 wran1 = "E" & 6
 wran2 = "E" & f1
 wranF = "E" & f1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 
 wran1 = "F" & 6
 wran2 = "F" & f1
 wranF = "F" & f1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 TT_TOTAL_VENTAS = TT_TOTAL_VENTAS + val(xl.Range(wranF))
 TT_TOTAL_CREDITO = TT_TOTAL_CREDITO + val(xl.Range(wranF))
 
 wran1 = "G" & 6
 wran2 = "G" & f1
 wranF = "G" & f1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 wran1 = "H" & 6
 wran2 = "H" & f1
 wranF = "H" & f1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 TT_TOTAL_VENTAS = TT_TOTAL_VENTAS + val(xl.Range(wranF))
 TT_TOTAL_AUTOCONSUMO = TT_TOTAL_AUTOCONSUMO + val(xl.Range(wranF))
 
 wran1 = "I" & 6
 wran2 = "I" & f1
 wranF = "I" & f1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 wran1 = "J" & 6
 wran2 = "J" & f1
 wranF = "J" & f1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
   
 wranF = "A" & f1 + 1 & ":J" & f1 + 1
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
 wranF = "A" & f1 + 2 & ":J" & f1 + 2
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 6
 
  FrmImp3.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
  xl.Cells(2, 1) = Trim(retra_llave!TRA_DESCRIPCION)
  xl.Cells(3, 1) = "'" & Format(wsFECHA1, "dd/mm/yyyy") & " al " & Format(wsFECHA2, "dd/mm/yyyy")
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.APPLICATION.Visible = True
  DoEvents
  FrmImp3.lblproceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp3.Pantalla.Enabled = True
  FrmImp3.cerrar.Enabled = True
  FrmImp3.Pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblproceso.Visible = False

Exit Sub

PROFACART:
suma_v_total = 0
suma_c_total = 0
suma_v_contado = 0
suma_c_contado = 0
suma_v_credito = 0
suma_c_credito = 0
suma_v_auto = 0
suma_c_auto = 0

Do Until llave_rep02.EOF
    If val(llave_rep02!FAR_equiv) = 0 Then
      MsgBox "Equivalencia 0 Reset "
      GoTo CANCELA
    End If
    wvalor = val(llave_rep02!FAR_PRECIO) * (val(llave_rep02!far_cantidad) / val(llave_rep02!FAR_equiv))
    wcanti = (val(llave_rep02!far_cantidad)) / val(llave_rep02!FAR_equiv)
    If llave_rep02!far_signo_car = 0 Then
      suma_v_contado = suma_v_contado + wvalor
      suma_c_contado = suma_c_contado + wcanti
    End If
    If llave_rep02!far_signo_car = 1 Then
      suma_v_credito = suma_v_credito + wvalor
      suma_c_credito = suma_c_credito + wcanti
    End If
    suma_v_total = suma_v_total + wvalor
    suma_c_total = suma_c_total + wcanti
    'TT_COSTO_VENTAS = TT_COSTO_VENTAS + (Val(llave_rep02!FAR_CANTIDAD) * Val(llave_rep02!far_cospro))
    
  llave_rep02.MoveNext
 Loop
 
 
Return
  
 
Exit Sub

IMPFACART:
       f1 = f1 + 1
       PUB_KEY = PUB_KEY
       SQ_OPER = 1
       pu_codcia = LK_CODCIA
       LEER_ART_LLAVE
       xl.Cells(f1, 2) = art_LLAVE!art_nombre
       xl.Cells(f1, 3) = suma_c_contado
       xl.Cells(f1, 4) = suma_v_contado
       xl.Cells(f1, 5) = suma_c_credito
       xl.Cells(f1, 6) = suma_v_credito
       xl.Cells(f1, 7) = suma_c_auto
       xl.Cells(f1, 8) = suma_v_auto
       xl.Cells(f1, 9) = suma_c_total
       xl.Cells(f1, 10) = suma_v_total
Return

COBRANZA_CLI:
 PS_REP02(0) = LK_CODCIA
 PS_REP02(1) = wsFECHA1
 PS_REP02(2) = wsFECHA2
 llave_rep02.Requery
 Do Until llave_rep02.EOF
    TT_COBRANZA_CLI = TT_COBRANZA_CLI + val(llave_rep02!ALL_IMPORTE)
  llave_rep02.MoveNext
 Loop


Return

CANCELA:
  FrmImp3.Pantalla.Enabled = True
  FrmImp3.Pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblproceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  lblproceso.Caption = "Abriendo , Archivo REGVENTA.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\GRIFOS\CAJAGEN.xls", 0, True, 4, WPAS, WPAS
Return



FINTODO:
 MsgBox Err.Description & " .-  Reintente Nuevamente ..", 48, Pub_Titulo
  Resume Next
 GoTo CANCELA
 Resume Next
End Sub


Public Sub CONSO_FAC()
Dim cade_fbg As String
Dim cade_serie As String
Dim cade_docu As String
Dim CADENITA
Dim wkSELECT
Dim Wche
Dim WNUMSER
  Dim wnumfac
  Dim WFBG
  Dim WDOC As String
Pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
      REP_FECHA1 = Left(txtCampo1.Text, 8)
Else
     REP_FECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtcampo2.Text, 2) = "__" Then
     REP_FECHA2 = Left(txtcampo2.Text, 8)
Else
     REP_FECHA2 = Trim(txtcampo2.Text)
End If
If Not IsDate(REP_FECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If Not IsDate(REP_FECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(REP_FECHA1) > CDate(REP_FECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If

Reportes.Connect = PUB_ODBC
Reportes.WindowTitle = "Reporte :  " & Trim(retra_llave(1))
Reportes.Destination = crptToWindow
Reportes.WindowLeft = 2
Reportes.WindowTop = 70
Reportes.WindowWidth = 635
Reportes.WindowHeight = 390

CADENITA = ""
wkSELECT = ""
Wche = 0
WDOC = ""
cade_fbg = ""
cade_serie = ""
cade_docu = ""
For fila = 0 To lisD.ListCount - 1
  lisD.ListIndex = fila
  WNUMSER = str(val(Mid(lisD.Text, 4, 3)))
  wnumfac = str(val(Mid(lisD.Text, 10, 16)))
  WFBG = Left(lisD.Text, 1)
  If lisD.Selected(fila) And Trim(lisD.Text) <> "" Then
      cade_fbg = cade_fbg + "'" + Trim(WFBG) + "',"
      cade_serie = cade_serie + "'" + Trim(WNUMSER) + "',"
      cade_docu = cade_docu + wnumfac + ","
  End If
  If lisD.Selected(fila) And Trim(lisD.Text) <> "" Then
       WDOC = WDOC + Right(Trim(lisD.Text), 6)
  '  If WPLA = "A" Then
  '    If Wche = 0 Then
  '     wkSELECT = "({FACART.FAR_FBG} = '" & Trim(WFBG) & "' AND {FACART.FAR_NUMSER} = " & Trim(WNUMSER) & " AND {FACART.FAR_NUMFAC} = " & wnumfac & ") "
  '    Else
  '     wkSELECT = wkSELECT + " OR ({FACART.FAR_FBG} = '" & Trim(WFBG) & "' AND {FACART.FAR_NUMSER} = " & Trim(WNUMSER) & " AND {FACART.FAR_NUMFAC} = " & wnumfac & ") "
  '    End If
  '   Else
      If Wche = 0 Then
       wkSELECT = "({FACART.FAR_FBG} = '" & Trim(WFBG) & "' AND {FACART.FAR_NUMSER} = '" & Trim(WNUMSER) & "' AND {FACART.FAR_NUMFAC} = " & wnumfac & ") "
      Else
       wkSELECT = wkSELECT + " OR ({FACART.FAR_FBG} = '" & Trim(WFBG) & "' AND {FACART.FAR_NUMSER} = '" & Trim(WNUMSER) & "' AND {FACART.FAR_NUMFAC} = " & wnumfac & ") "
      End If
  '   End If
    Wche = 1
  End If
  lisD.ListIndex = fila
Next fila
If Wche = 0 Then
     wkSELECT = "( {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_TIPMOV} = 10 AND {FACART.FAR_ESTADO} <> 'E')"
Else
    cade_fbg = "{FACART.FAR_FBG} in [" & Mid(cade_fbg, 1, Len(cade_fbg) - 1) & "]"
    cade_serie = "{FACART.FAR_NUMSER} in [" & Mid(cade_serie, 1, Len(cade_serie) - 1) & "]"
    cade_docu = "{FACART.FAR_NUMFAC} in [" & Mid(cade_docu, 1, Len(cade_docu) - 1) & "]"
    wkSELECT = cade_fbg + " AND " & " " + cade_serie + " AND " & " " + cade_docu + " AND ( {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_TIPMOV} = 10 AND {FACART.FAR_ESTADO} <> 'E')"
     'Debug.Print wkSELECT
    
    'wkSELECT = "(" + wkSELECT + ") AND ( {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_TIPMOV} = 10 AND {FACART.FAR_ESTADO} <> 'E')"
End If

CADENITA = wkSELECT
Reportes.Formulas(0) = ""
Reportes.Formulas(1) = ""
Reportes.Formulas(2) = ""
Reportes.Formulas(3) = ""

Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "HRESU.rpt"


Reportes.Formulas(0) = "CIA='" & MDIForm1.lblciagen & "'"
Reportes.Formulas(1) = "FECHA='" & LK_FECHA_DIA & "'"

Reportes.SelectionFormula = CADENITA
If Left(cmbmoneda.Text, 1) <> "T" Then
 If Trim(CADENITA) = "" Then
   CADENITA = CADENITA + "{FACART.FAR_MONEDA} = '" & Left(cmbmoneda.Text, 1) & "'"
 Else
   CADENITA = CADENITA + " AND {FACART.FAR_MONEDA} = '" & Left(cmbmoneda.Text, 1) & "'"
 End If
End If
 'Debug.Print CADENITA
 'CADENITA = "{FACART.FAR_FBG} in ['F'] AND  {FACART.FAR_NUMSER} in ['20'] AND  {FACART.FAR_NUMFAC} in [10769]  AND ( {FACART.FAR_CODCIA} = '01' AND {FACART.FAR_TIPMOV} = 10 AND {FACART.FAR_ESTADO} <> 'E')"
 'CADENITA = "{FACART.FAR_FBG} in ['F','F'] AND  {FACART.FAR_NUMSER} in ['20','20'] AND  {FACART.FAR_NUMFAC} in [ 10769, 10770]  AND ( {FACART.FAR_CODCIA} = '01' AND {FACART.FAR_TIPMOV} = 10 AND {FACART.FAR_ESTADO} <> 'E')"
 'Reportes.SelectionFormula = CADENITA
Reportes.Action = 1




Reportes.ReportFileName = Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\" & "hRESU3.rpt"
Reportes.Formulas(0) = "TITULO='----> CONSOLIDADO DE DOCUMENTOS '"
Reportes.Formulas(1) = "CIA='" & MDIForm1.lblciagen & "'"
Reportes.Formulas(2) = "FECHA='" & LK_FECHA_DIA & "'"

Reportes.Action = 1
Pantalla.Enabled = True
cerrar.Enabled = True



Exit Sub
CANCELA:
 ' REPORTES
 
End Sub

Public Sub PROC()

Dim WFBG
Dim WNUMSER
Dim wnumfac

pub_cadena = "SELECT far_moneda, far_numfac, far_numser, far_fbg FROM facart WHERE FAR_CODCIA = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND FAR_ESTADO <> 'E' AND FAR_TIPMOV = 10  ORDER BY  FAR_FECHA, FAR_NUMFAC"
Set PSPRO1 = CN.CreateQuery("", pub_cadena)
PSPRO1(0) = 0
PSPRO1(1) = 0
PSPRO1(2) = 0
Set pro1_llave = PSPRO1.OpenResultset(rdOpenKeyset, rdConcurValues)
PSPRO1(0) = LK_CODCIA
PSPRO1(1) = REP_FECHA1
PSPRO1(2) = REP_FECHA2
pro1_llave.Requery
If pro1_llave.EOF Then
  Exit Sub
End If
WFBG = pro1_llave!far_fbg
WNUMSER = pro1_llave!far_numser
wnumfac = pro1_llave!far_numfac
lisD.Clear
If Left(cmbmoneda.Text, 1) <> "T" Then
   If Trim(pro1_llave!FAR_MONEDA) <> Left(cmbmoneda.Text, 1) Then GoTo saleart1
End If

lisD.AddItem Trim(pro1_llave!far_fbg) & "/ " & Format(pro1_llave!far_numser, "000") & " - " & Format(pro1_llave!far_numfac, "0000000")
saleart1:
fila = 0
Do Until pro1_llave.EOF
  If Left(cmbmoneda.Text, 1) <> "T" Then
     If Trim(pro1_llave!FAR_MONEDA) <> Left(cmbmoneda.Text, 1) Then GoTo saleart
  End If
  If Trim(WFBG) = Trim(pro1_llave!far_fbg) And Trim(WNUMSER) = Trim(pro1_llave!far_numser) And val(wnumfac) = val(pro1_llave!far_numfac) Then
  Else
    fila = fila + 1
    lisD.AddItem Trim(pro1_llave!far_fbg) & "/ " & Format(pro1_llave!far_numser, "000") & " - " & Format(pro1_llave!far_numfac, "0000000")

'    lisD.Selected(fila) = True
  End If
  
  
  WFBG = pro1_llave!far_fbg
  WNUMSER = pro1_llave!far_numser
  wnumfac = pro1_llave!far_numfac
saleart:
  pro1_llave.MoveNext
Loop

lisD.Visible = True
lisD.SetFocus

End Sub

Public Sub PROC_EXCEL()
Dim WFBG
Dim WNUMSER
Dim wnumfac
Dim WMAX_FILA  As Integer
Dim WCOL_MAX As Integer

GoSub WEXCEL
f1 = 6
WMAX_FILA = 26
WCOL_MAX = 4
c1 = 4
fila = 0
For fila = 0 To lisD.ListCount - 1
 lisD.ListIndex = fila
 xl.APPLICATION.Visible = True
 If lisD.Selected(fila) Then
   WNUMSER = val(Mid(lisD.Text, 4, 3))
   wnumfac = val(Mid(lisD.Text, 10, 16))
   WFBG = Left(lisD.Text, 1)
   If f1 >= WMAX_FILA + 6 Then
     f1 = 6
     c1 = c1 + 1
   End If
   f1 = f1 + 1
   xl.Cells(f1, c1) = WFBG + "." + str(WNUMSER) + "-" + str(wnumfac)
 End If
Next fila
xl.DisplayAlerts = False
xl.Worksheets(1).Protect PUB_CLAVE
xl.APPLICATION.Visible = True
Set xl = Nothing
Screen.MousePointer = 0
Exit Sub

WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  lblproceso.Caption = "Abriendo , Archivo DOCU.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\DOCU.xls", 0, True, 4, WPAS, WPAS
Return

End Sub

Private Sub txt_cli_GotFocus()
'Azul txt_cli, txt_cli
lblCliente.Caption = ""
End Sub
Private Sub txt_cli_KeyDown(KeyCode As Integer, Shift As Integer)
Dim strFindMe As String
Dim itmFound As ListItem    ' Variable FoundItem.
If Not ListView2.Visible Then
 Exit Sub
End If
If KeyCode <> 40 And KeyCode <> 38 And KeyCode <> 34 And KeyCode <> 33 And txt_cli.Text = "" Then
  loc_key = 1
  Set ListView2.SelectedItem = ListView2.ListItems(loc_key)
  ListView2.ListItems.Item(loc_key).Selected = True
  ListView2.ListItems.Item(loc_key).EnsureVisible
  GoTo fin
End If

If KeyCode = 40 Then  ' flecha abajo
  loc_key = loc_key + 1
  If loc_key > ListView2.ListItems.count Then loc_key = ListView2.ListItems.count
  GoTo posicion
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo posicion
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > ListView2.ListItems.count Then loc_key = ListView2.ListItems.count
 GoTo posicion
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo posicion
End If
GoTo fin
posicion:
'  KeyCode = 0
  ListView2.ListItems.Item(loc_key).Selected = True
  ListView2.ListItems.Item(loc_key).EnsureVisible
  txt_cli.Text = Trim(ListView2.ListItems.Item(loc_key).Text) & " "
  DoEvents
  txt_cli.SelStart = Len(txt_cli.Text)
  DoEvents
fin:

End Sub
Private Sub txt_cli_KeyPress(KeyAscii As Integer)
Dim valor As String
Dim tf As Integer
Dim i
Dim itmFound As ListItem    ' Variable FoundItem.
If KeyAscii = 27 Then
 ListView2.Visible = False
 txt_cli.Text = ""
 Exit Sub
End If
If KeyAscii <> 13 Then
   GoTo fin
End If
On Error GoTo ERROR_CODIGO
pu_codclie = val(txt_cli.Text)
On Error GoTo 0
If Len(txt_cli.Text) = 0 Then
   Exit Sub
End If

If pu_codclie <> 0 And IsNumeric(txt_cli.Text) = True Then
   SQ_OPER = 1
   pu_cp = "C"
   If Wfile = "PROGRA" Or Wfile = "PROGRA_TEJ" Then pu_cp = "P"
   pu_codcia = LK_CODCIA
   LEER_CLI_LLAVE
   If cli_llave.EOF Then
     lblCliente.Caption = ""
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     Azul txt_cli, txt_cli
     GoTo fin
   Else
     lblCliente.Caption = Trim(cli_llave!CLI_NOMBRE)
   End If
   If Wfile = "PROGRA" Or Wfile = "PROGRA_TEJ" Then
    If Pantalla.Visible Then Pantalla.SetFocus
   Else
   op1_Click 0
   op1(0).SetFocus
   End If
   'If pantalla.Visible And pantalla.Enabled Then
   '  pantalla.SetFocus
   'End If
Else
   If loc_key > ListView2.ListItems.count Or loc_key = 0 Then
     Exit Sub
   End If
   valor = UCase(ListView2.ListItems.Item(loc_key).Text)
   If Trim(UCase(txt_cli.Text)) = Left(valor, Len(Trim(txt_cli.Text))) Then
   Else
      Exit Sub
   End If
   lblCliente.Caption = Trim(ListView2.ListItems.Item(loc_key).Text)
   txt_cli.Text = Trim(ListView2.ListItems.Item(loc_key).SubItems(1))
   If Wfile = "PROGRA" Or Wfile = "PROGRA_TEJ" Then
      If Pantalla.Visible Then Pantalla.SetFocus
   Else
    op1_Click 0
    op1(0).SetFocus
   End If
   'If pantalla.Visible And pantalla.Enabled Then
   '  pantalla.SetFocus
   'End If
End If

dale:
ListView2.Visible = False
fin:
Exit Sub
ERROR_CODIGO:
MsgBox "Codigo NO Valido .... ", 48, Pub_Titulo
Azul txt_cli, txt_cli

End Sub

Private Sub txt_cli_KeyUp(KeyCode As Integer, Shift As Integer)
Dim VAR
If Len(txt_cli.Text) = 0 Or IsNumeric(txt_cli.Text) = True Then
   ListView2.Visible = False
   Exit Sub
End If
If ListView2.Visible = False And KeyCode <> 13 Then
    VAR = Asc(txt_cli.Text)
    VAR = VAR + 1
    If VAR = 33 Or VAR = 91 Then
       VAR = "ZZZZZZZZ"
    ElseIf VAR = 58 Then
       VAR = "A"
    Else
       VAR = Chr(VAR)
    End If
    numarchi = 1
    PUB_CP = "C"
    If Wfile = "PROGRA" Or Wfile = "PROGRA_TEJ" Then
      PUB_CP = "P"
    End If
    'archi = "SELECT CLI_CODCLIE, CLI_CODCIA, CLI_CP, CLI_NOMBRE,CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM  FROM CLIENTES WHERE  CLI_CP = '" & PUB_CP & "' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & txt_cli.Text & "' AND  '" & VAR & "' ORDER BY CLI_NOMBRE"
    archi = "SELECT CLI_CODCLIE , CLI_CODCIA, CLI_CP, CLI_NOMBRE, CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM, TAB_NOMLARGO  FROM CLIENTES,TABLAS WHERE (TAB_CODCIA = '00') AND (TAB_TIPREG = 35) AND (TAB_NUMTAB = CLI_ZONA_NEW) AND CLI_CP = '" & PUB_CP & "' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & txt_cli.Text & "' AND  '" & VAR & "' ORDER BY CLI_NOMBRE"
'    If Trim(txt_cli.text) <> "" And ListView1.ListItems.count = 0 Then
'    Else
     PROC_LISVIEW ListView2
     loc_key = 0
     If ListView2.Visible Then
      loc_key = 1
     End If
 '   End If
    Exit Sub
End If

If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
 Exit Sub
End If
Dim itmFound As ListItem    ' Variable FoundItem.
If ListView2.Visible Then
  Set itmFound = ListView2.FindItem(LTrim(txt_cli.Text), lvwText, , lvwPartial)
  If itmFound Is Nothing Then
  Else
   itmFound.EnsureVisible
   itmFound.Selected = True
   loc_key = itmFound.Tag
   If loc_key + 8 > ListView2.ListItems.count Then
      ListView2.ListItems.Item(ListView2.ListItems.count).EnsureVisible
   Else
     ListView2.ListItems.Item(loc_key + 8).EnsureVisible
   End If
   DoEvents
  End If
  Exit Sub
End If


End Sub

Public Sub DETALLE_OP()
Dim DIA
Dim DIA1
Dim MES
Dim MES1
Dim ANO
Dim ANO1
Dim WFLAG As String * 1
Dim CADENITA
Dim wkSELECT
Dim Wche
Dim WNUMSER
  Dim wnumfac
  Dim WFBG
Pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
      REP_FECHA1 = Left(txtCampo1.Text, 8)
Else
     REP_FECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtcampo2.Text, 2) = "__" Then
     REP_FECHA2 = Left(txtcampo2.Text, 8)
Else
     REP_FECHA2 = Trim(txtcampo2.Text)
End If
If Not IsDate(REP_FECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If Not IsDate(REP_FECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(REP_FECHA1) > CDate(REP_FECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If

Reportes.Connect = PUB_ODBC
Reportes.WindowTitle = "Reporte :  " & Trim(retra_llave(1))
Reportes.Destination = crptToWindow
Reportes.WindowLeft = 2
Reportes.WindowTop = 70
Reportes.WindowWidth = 635
Reportes.WindowHeight = 390
DIA = Day(REP_FECHA1)
MES = Month(REP_FECHA1)
ANO = Year(REP_FECHA1)
DIA1 = Day(REP_FECHA2)
MES1 = Month(REP_FECHA2)
ANO1 = Year(REP_FECHA2)

 
If op1(0).Value Then
    WFLAG = " "
ElseIf op1(1).Value Then
    WFLAG = "X"
ElseIf op1(1).Value Then
    WFLAG = "A"
End If

CADENITA = ""
wkSELECT = ""
Wche = 0
For fila = 0 To listop.ListCount - 1
  listop.ListIndex = fila
  WNUMSER = str(val(Mid(listop.Text, 17, 3)))
  wnumfac = str(val(Mid(listop.Text, 23, 30)))
  WFBG = Left(listop.Text, 1)
  If listop.Selected(fila) And Trim(listop.Text) <> "" Then
  '{PEDIDOS.PED_FECHA} >= Date ( " & ANO & "," & MES & "," & DIA & ") AND {PEDIDOS.PED_FECHA} <= Date ( " & ANO1 & "," & MES1 & "," & DIA1 & ")"
      If Wche = 0 Then
       wkSELECT = "( {PEDIDOS.PED_NUMSER} = '" & Trim(WNUMSER) & "' AND {PEDIDOS.PED_NUMFAC} = " & wnumfac & ") "
      Else
       wkSELECT = wkSELECT + " OR ( {PEDIDOS.PED_NUMSER} = '" & Trim(WNUMSER) & "' AND {PEDIDOS.PED_NUMFAC} = " & wnumfac & ") "
      End If
      Wche = 1
  End If
  listop.ListIndex = fila
Next fila
If Wche = 0 Then
     wkSELECT = "({PEDIDOS.PED_CODCIA} = '" & LK_CODCIA & "' AND {PEDIDOS.PED_TIPMOV} = 177 )"
Else
    wkSELECT = "(" + wkSELECT + ") AND ( {PEDIDOS.PED_CODCIA} = '" & LK_CODCIA & "' AND {PEDIDOS.PED_TIPMOV} = 177 )"
End If

CADENITA = wkSELECT
Reportes.Formulas(0) = "CIA=  '" & Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))) & "'"
If op1(0).Value Then
 Reportes.Formulas(1) = "TITULO= 'LISTADO DE PENDIENTES DE ORD. DE PEDIDOS POR CLIENTE '"
ElseIf op1(1).Value Then
 Reportes.Formulas(1) = "TITULO= 'LISTADO DE ORD. DE PEDIDOS ENTREGADOS POR CLIENTE ' "
ElseIf op1(2).Value Then
 Reportes.Formulas(1) = "TITULO= 'LISTADO DE ORD. DE PEDIDOS POR CLIENTE' "
End If

Reportes.Formulas(2) = ""
Reportes.Formulas(3) = ""
Reportes.ReportFileName = PUB_RUTA_OTRO & "estop.rpt"
Reportes.SelectionFormula = CADENITA
'Debug.Print CADENITA
Reportes.Action = 1
Pantalla.Enabled = True
cerrar.Enabled = True
Exit Sub
CANCELA:
 ' REPORTES
 

End Sub
Public Sub MOVICONT_FAL()
'On Error GoTo FINTODO
Dim ws_clave
Dim wvalor As Currency
Dim wcanti As Currency
Dim WW_CORRELA As Currency
Dim tfechaini
Dim tfechafin
Dim ip_numfac
Dim ip_fecha
Dim wwnumfac
Dim WFLAG As String * 1
If Trim(fbg.Text) = "" Then
  MsgBox "Ingrese Tipo de Documento.", 48, Pub_Titulo
  Exit Sub
End If
If val(serie.Text) <= 0 Then
   MsgBox "Ingrese Serie del Documento.", 48, Pub_Titulo
   Azul serie, serie
   Exit Sub
End If
If val(numini.Text) <= 0 Then
   MsgBox "Ingrese N° inicial del Documento.", 48, Pub_Titulo
   Azul numini, numini
   Exit Sub
End If
If val(numfin.Text) <= 0 Then
   MsgBox "Ingrese N° Final del Documento.", 48, Pub_Titulo
   Azul numfin, numfin
   Exit Sub
End If

    
Pantalla.Enabled = False
cerrar.Enabled = False
pub_cadena = ""
xcuenta = 0
Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp3.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = PUB_CLAVE
f1 = 9  'Fila Inicial
WFLAG = ""

'WCONTROL = WCONTROL + 1
pub_cadena = "SELECT Distinct MOV_NUMFAC, MOV_FECHA_EMI FROM MOVICONT WHERE MOV_CODCIA = ? AND MOV_FBG = ? AND MOV_SERIE = ? AND MOV_NUMFAC >= ? AND MOV_NUMFAC <= ? AND MOV_TIPMOV = ?  ORDER BY  MOV_NUMFAC"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = 0
PS_REP02(1) = 0
PS_REP02(2) = 0
PS_REP02(3) = 0
PS_REP02(4) = 0
PS_REP02(5) = 0
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
PS_REP02(0) = LK_CODCIA
PS_REP02(1) = Left(fbg.Text, 2)
PS_REP02(2) = serie.Text
PS_REP02(3) = numini.Text
PS_REP02(4) = numfin.Text
PS_REP02(5) = "01"
llave_rep02.Requery
If llave_rep02.EOF Then
   MsgBox "En este rango no Existe documentos ", 48, Pub_Titulo
   GoTo CANCELA
End If

' el PS_REP1(0) ESTA MAS ABAJO
DoEvents
FrmImp3.lblproceso.Visible = True
FrmImp3.ProgBar.Visible = True
FrmImp3.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
FrmImp3.lblproceso.Caption = "Procesando . . . "
DoEvents
FrmImp3.ProgBar.Min = 0
FrmImp3.ProgBar.Value = 0

 
 FrmImp3.ProgBar.max = llave_rep02.RowCount
 If Trim(fbg.Text) = "F" Then
   xl.Cells(5, 2) = "FACTURAS"
 Else
   xl.Cells(5, 2) = "BOLETAS"
 End If
 WW_CORRELA = val(numini.Text)
 ip_numfac = val(numini.Text)
 ip_fecha = Format(llave_rep02!MOV_fecha_EMI, "dd/mm/yyyy")
 tfechaini = Format(llave_rep02!MOV_fecha_EMI, "dd/mm/yyyy")
 wwnumfac = llave_rep02!MOV_numfac
 xl.Cells(6, 2) = serie.Text
 xl.Cells(7, 2) = numini.Text
 xl.Cells(8, 2) = tfechaini
 
 Do Until llave_rep02.EOF
    FrmImp3.ProgBar.Value = FrmImp3.ProgBar.Value + 1
    DoEvents
    ip_fecha = Format(llave_rep02!MOV_fecha_EMI, "dd/mm/yyyy")
    If WW_CORRELA < llave_rep02!MOV_numfac Then
      For fila = 1 To (llave_rep02!MOV_numfac - WW_CORRELA)
      ip_numfac = WW_CORRELA + fila - 1
      GoSub IMP_NUMFAC
      Next fila
      WFLAG = "A"
    End If
    WW_CORRELA = llave_rep02!MOV_numfac
    WW_CORRELA = WW_CORRELA + 1
    numfin.Text = llave_rep02!MOV_numfac
    tfechafin = Format(llave_rep02!MOV_fecha_EMI, "dd/mm/yyyy")
    llave_rep02.MoveNext
 Loop
 wranF = "A" & f1 + 1 & ":D" & f1 + 1
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
 f1 = f1 + 2
 xl.Cells(f1, 1) = "FIN DE LISTADO "
 If WFLAG <> "A" Then
   f1 = f1 + 1
   xl.Cells(f1, 1) = "*** LOS DOC. ESTAN CORRECTOS *** "
 End If
 f1 = f1 + 1
 xl.Cells(f1, 1) = "N° FINAL:"
 xl.Cells(f1, 2) = numfin.Text
 f1 = f1 + 1
 xl.Cells(f1, 1) = "FEC.FINAL:"
 xl.Cells(f1, 2) = "'" & tfechafin
 wranF = "A" & f1 + 1 & ":D" & f1 + 1
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
 
  
 
 'wranF = "A" & F1 + 2 & ":C" & F1 + 2
 'xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 6
 
 'Do Until llave_rep02.EOF
 '   ip_fecha = Format(llave_rep02!FAR_fecha_COMPRA, "dd/mm/yyyy")
 '   ip_numfac = WW_CORRELA
 '   If WW_CORRELA <> llave_rep02!FAR_numfac Then
 '     GoSub IMP_NUMFAC
 '     WW_CORRELA = WW_CORRELA + 1
 '   End If
 '   WW_CORRELA = WW_CORRELA + 1
 '   wwnumfac = llave_rep02!FAR_numfac
 '   tfechafin = Format(llave_rep02!FAR_fecha_COMPRA, "dd/mm/yyyy")
 '   llave_rep02.MoveNext
 'Loop
 
  FrmImp3.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  

  xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
  xl.Cells(2, 1) = "'" & Format(LK_FECHA_DIA, "dd/mm/yyyy")
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.APPLICATION.Visible = True
  DoEvents
  FrmImp3.lblproceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp3.Pantalla.Enabled = True
  FrmImp3.cerrar.Enabled = True
  FrmImp3.Pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblproceso.Visible = False

Exit Sub

IMP_NUMFAC:
       f1 = f1 + 1
       If Trim(fbg.Text) = "F" Then
           xl.Cells(f1, 1) = "FACTURAS"
       Else
           xl.Cells(f1, 1) = "BOLETAS"
       End If
       xl.Cells(f1, 2) = ip_numfac
       xl.Cells(f1, 3) = "FECHA PROBABLE:"
       xl.Cells(f1, 4) = "'" + ip_fecha
Return


CANCELA:
  FrmImp3.Pantalla.Enabled = True
  FrmImp3.Pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblproceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  lblproceso.Caption = "Abriendo , Archivo NUMDOC.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\STANDAR\NUMDOC.xls", 0, True, 4, WPAS, WPAS
Return



FINTODO:
 MsgBox Err.Description & " .-  Reintente Nuevamente ..", 48, Pub_Titulo
  Resume Next
 GoTo CANCELA
 Resume Next
End Sub

Public Sub PROGRA()
'On Error GoTo FINTODO
Dim wcar As Currency
Dim wcodclie As Currency
Dim CONS_CARDADO As Currency
Dim WPRECIO As Currency
Dim CAPTURA_FILA As Integer
Dim wsFECHA1
Dim wsFECHA2
Dim ws_clave

CONS_CARDADO = 0
Pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtcampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtcampo2.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If Not IsDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If


pub_cadena = ""
xcuenta = 0

Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp3.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = PUB_CLAVE
f1 = 6  'Fila Inicial

pub_cadena = "SELECT FAR_BRUTO, FAR_CODCIA , FAR_PEDFAC, FAR_PEDSER, FAR_TIPMOV , FAR_CODART, FAR_SERGUIA, FAR_NUMGUIA, FAR_FECHA_COMPRA, FAR_CODCLIE , FAR_JABAS, FAR_CANTIDAD FROM FACART WHERE FAR_CODCIA = ? AND FAR_TIPMOV = ? AND (FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ?) AND FAR_CODCLIE = ? AND FAR_ESTADO <> 'E' AND FAR_SIGNO_ARM = 1  ORDER BY FAR_CODCLIE,  FAR_FECHA_COMPRA, FAR_NUMOPER2 "
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = 0
PS_REP01(2) = LK_FECHA_DIA
PS_REP01(3) = LK_FECHA_DIA
PS_REP01(4) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT * FROM TABLAS WHERE TAB_TIPREG = 232 AND TAB_CODCIA = ? AND TAB_CODART = ? AND TAB_CONTABLE2 = ? "
Set PS_REP03 = CN.CreateQuery("", pub_cadena)
PS_REP03(0) = 0
PS_REP03(1) = 0
PS_REP03(2) = 0
Set llave_rep03 = PS_REP03.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


pub_cadena = "SELECT PED_SERGUIA, PED_NUMGUIA, PED_FECHA_EMISION, PED_CANTIDAD, PED_ROLLOS, PED_CONTACTO, ART_MARCA , ART_NOMBRE FROM PEDIDOS, ARTI WHERE (PED_CODCIA = ART_CODCIA) AND (PED_CODART = ART_KEY) AND PED_CODCIA = ? AND PED_TIPMOV = ? AND PED_NUMSER = ? AND PED_NUMFAC = ? AND PED_CODART = ? "
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = 0
PS_REP02(1) = 0
PS_REP02(2) = 0
PS_REP02(3) = 0
PS_REP02(4) = 0
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


'WCONTROL = WCONTROL + 1


DoEvents
FrmImp3.lblproceso.Visible = True
FrmImp3.ProgBar.Visible = True
FrmImp3.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
'PUB_KEY = 0
'SQ_OPER = 2
'pu_codcia = LK_CODCIA
'If LK_EMP_PTO = "A" Then
'  pu_codcia = "00"
'End If
'LEER_ART_LLAVE
'If art_mayor.EOF Then
' MsgBox "No Exieten Productos ..", 48, Pub_Titulo
' Exit Sub
'End If
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = 180
PS_REP01(2) = wsFECHA1
PS_REP01(3) = wsFECHA2
PS_REP01(4) = val(txt_cli.Text)
llave_rep01.Requery
If llave_rep01.EOF Then
  Pantalla.Enabled = True
  cerrar.Enabled = True
  FrmImp3.lblproceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  MsgBox "No Existe Datos en este Rango..", 48, Pub_Titulo
  Exit Sub
End If
GoSub WEXCEL

FrmImp3.lblproceso.Caption = "Procesando . . . "
DoEvents
FrmImp3.ProgBar.Min = 0
FrmImp3.ProgBar.Value = 0
FrmImp3.ProgBar.max = llave_rep01.RowCount
wcodclie = llave_rep01!far_codclie
PS_REP02(0) = LK_CODCIA
PS_REP02(1) = 180

SQ_OPER = 1
pu_cp = "P"
pu_codcia = LK_CODCIA
pu_codclie = llave_rep01!far_codclie
LEER_CLI_LLAVE
xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
xl.Cells(2, 1) = "LISTADO DE lIQUIDACION DE SERVICIO DEL " & Format(wsFECHA1, "dd/mm/yy") & " AL " & Format(wsFECHA2, "dd/mm/yy")
xl.Cells(3, 1) = "Prov. : " & Trim(cli_llave!CLI_NOMBRE)
xl.Cells(4, 1) = "Prcios del Cardado :"
xl.Cells(4, 4) = val(cli_llave!CLI_NOMBRE_EMPRESA)
f1 = 5

Do Until llave_rep01.EOF
    FrmImp3.ProgBar.Value = FrmImp3.ProgBar.Value + 1
    PS_REP02(2) = llave_rep01!FAR_PEDSER
    PS_REP02(3) = llave_rep01!FAR_PEDFAC
    PS_REP02(4) = llave_rep01!far_codart
    llave_rep02.Requery
    f1 = f1 + 1
    xl.Cells(f1, 1) = "'" & llave_rep02!PED_NUMGUIA
    xl.Cells(f1, 2) = llave_rep02!PED_FECHA_EMISION
    xl.Cells(f1, 3) = llave_rep01!far_JABAS
    xl.Cells(f1, 4) = Format(llave_rep01!FAR_BRUTO, "0.000")
    xl.Cells(f1, 5) = Trim(llave_rep02!art_nombre)
    xl.Cells(f1, 6) = llave_rep01!FAR_fecha_compra
    xl.Cells(f1, 7) = "'" & Format(llave_rep01!far_serguia, "000") & "-" & llave_rep01!far_numguia
    PS_REP03(0) = "00"
    PS_REP03(1) = llave_rep02!art_marca
    
    SQ_OPER = 1
    pu_cp = "P"
    pu_codcia = LK_CODCIA
    pu_codclie = llave_rep01!far_codclie
    LEER_CLI_LLAVE
    PS_REP03(2) = Trim(cli_llave!cli_ruc_esposo)
    llave_rep03.Requery
    WPRECIO = 0
    CONS_CARDADO = val(cli_llave!CLI_NOMBRE_EMPRESA)
    If Not llave_rep03.EOF Then
       WPRECIO = val(llave_rep03!tab_NOMLARGO)
    End If
    xl.Cells(f1, 8) = Format(WPRECIO, "0.000")
    xl.Cells(f1, 9) = Format(llave_rep01!FAR_BRUTO * WPRECIO, "0.000")
    If Trim(llave_rep02!PED_CONTACTO) = "OK" Then
       xl.Cells(f1, 10) = Format(llave_rep01!FAR_BRUTO * CONS_CARDADO, "0.000")
    End If
 llave_rep01.MoveNext
Loop
  wran1 = "I" & 6
  wran2 = "I" & f1
  wranF = "I" & f1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wran1 = "J" & 6
  wran2 = "J" & f1
  wranF = "J" & f1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  
  xl.Cells(f1 + 1, 6) = "Total US$ ="
  wran1 = "I" & f1 + 1
  wran2 = "J" & f1 + 1
  wranF = "I" & f1 + 2
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  xl.Cells(f1 + 2, 6) = "Total General US$ ="
  wranF = "A" & f1 + 2 & ":J" & f1 + 2
  xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
  FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  DoEvents
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ""
  xl.APPLICATION.Visible = True
  DoEvents
  lblproceso.Visible = False
  ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  
  Set xl = Nothing
Exit Sub



CANCELA:
  FrmImp3.Pantalla.Enabled = True
  FrmImp3.Pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblproceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  lblproceso.Caption = "Abriendo , Archivo REGVENTA.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open PUB_RUTA_OTRO & "\SERVICIOS.xls", 0, True, 4, WPAS, WPAS
Return



FINTODO:
 MsgBox Err.Description & " .-  Reintente Nuevamente ..", 48, Pub_Titulo
  Resume Next
 GoTo CANCELA
 Resume Next
End Sub

Public Sub PROGRA_TEJ()
'On Error GoTo FINTODO
Dim wcodclie As Currency
Dim CONS_CARDADO As Currency
Dim WPRECIO As Currency
Dim CAPTURA_FILA As Integer
Dim wsFECHA1
Dim wsFECHA2
Dim ws_clave

CONS_CARDADO = 0
Pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtcampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtcampo2.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If Not IsDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If


pub_cadena = ""
xcuenta = 0

Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp3.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = PUB_CLAVE
f1 = 6  'Fila Inicial

'pub_cadena = "SELECT FAR_BRUTO, FAR_CODCIA , FAR_PEDFAC, FAR_PEDSER, FAR_TIPMOV , FAR_CODART, FAR_SERGUIA, FAR_NUMGUIA, FAR_FECHA_COMPRA, FAR_CODCLIE , FAR_JABAS, FAR_CANTIDAD FROM FACART WHERE FAR_CODCIA = ? AND FAR_TIPMOV = ? AND (FAR_FECHA_COMPRA >= ? AND FAR_FECHA_COMPRA <= ?) AND FAR_CODCLIE = ? AND FAR_ESTADO <> 'E' AND FAR_SIGNO_ARM = 1  ORDER BY FAR_CODCLIE,  FAR_FECHA_COMPRA, FAR_NUMOPER2 "
pub_cadena = "SELECT PED_codclie, PED_SERGUIA, PED_NUMGUIA, PED_FECHA_EMISION, PED_CANTIDAD, PED_ROLLOS, PED_CONTACTO, ART_MARCA , ART_NOMBRE, ART_NUMERO FROM PEDIDOS, ARTI WHERE (PED_CODCIA = ART_CODCIA) AND (PED_CODART = ART_KEY) AND PED_CODCIA = ? AND PED_TIPMOV = ? AND (PED_FECHA_EMISION >= ? AND PED_FECHA_EMISION <= ?) AND PED_CODCLIE = ? ORDER BY PED_FECHA_EMISION " ' PED_NUMSER = ? AND PED_NUMFAC = ? AND PED_CODART = ? "
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = 0
PS_REP01(2) = LK_FECHA_DIA
PS_REP01(3) = LK_FECHA_DIA
PS_REP01(4) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT * FROM TABLAS WHERE TAB_TIPREG = 230 AND TAB_CODCIA = ? AND TAB_CODART = ? AND TAB_CONTABLE2 = ? "
Set PS_REP03 = CN.CreateQuery("", pub_cadena)
PS_REP03(0) = 0
PS_REP03(1) = 0
PS_REP03(2) = 0
Set llave_rep03 = PS_REP03.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


pub_cadena = "SELECT PED_SERGUIA, PED_NUMGUIA, PED_FECHA_EMISION, PED_CANTIDAD, PED_ROLLOS, PED_CONTACTO, ART_NUMERO   FROM PEDIDOS, ARTI WHERE (PED_CODCIA = ART_CODCIA) AND (PED_CODART = ART_KEY) AND PED_CODCIA = ? AND PED_TIPMOV = ? AND PED_NUMSER = ? AND PED_NUMFAC = ? AND PED_CODART = ? "
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = 0
PS_REP02(1) = 0
PS_REP02(2) = 0
PS_REP02(3) = 0
PS_REP02(4) = 0
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


'WCONTROL = WCONTROL + 1


DoEvents
FrmImp3.lblproceso.Visible = True
FrmImp3.ProgBar.Visible = True
FrmImp3.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
'PUB_KEY = 0
'SQ_OPER = 2
'pu_codcia = LK_CODCIA
'If LK_EMP_PTO = "A" Then
'  pu_codcia = "00"
'End If
'LEER_ART_LLAVE
'If art_mayor.EOF Then
' MsgBox "No Exieten Productos ..", 48, Pub_Titulo
' Exit Sub
'End If
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = 180
PS_REP01(2) = wsFECHA1
PS_REP01(3) = wsFECHA2
PS_REP01(4) = val(txt_cli.Text)
llave_rep01.Requery
If llave_rep01.EOF Then
  Pantalla.Enabled = True
  cerrar.Enabled = True
  FrmImp3.lblproceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  MsgBox "No Existe Datos en este Rango..", 48, Pub_Titulo
  Exit Sub
End If
GoSub WEXCEL

FrmImp3.lblproceso.Caption = "Procesando . . . "
DoEvents
FrmImp3.ProgBar.Min = 0
FrmImp3.ProgBar.Value = 0
FrmImp3.ProgBar.max = llave_rep01.RowCount
wcodclie = llave_rep01!PED_CODCLIE
PS_REP02(0) = LK_CODCIA
PS_REP02(1) = 180

SQ_OPER = 1
pu_cp = "P"
pu_codcia = LK_CODCIA
pu_codclie = llave_rep01!PED_CODCLIE
LEER_CLI_LLAVE
xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
xl.Cells(2, 1) = "LISTADO DE lIQUIDACION DE SERVICIO DE TEJEDO DEL " & Format(wsFECHA1, "dd/mm/yy") & " AL " & Format(wsFECHA2, "dd/mm/yy")
xl.Cells(3, 1) = "Prov. : " & Trim(cli_llave!CLI_NOMBRE)
'xl.Cells(4, 1) = "Prcios del Cardado :"
'xl.Cells(4, 4) = Val(cli_llave!CLI_NOMBRE_EMPRESA)
f1 = 5

Do Until llave_rep01.EOF
    FrmImp3.ProgBar.Value = FrmImp3.ProgBar.Value + 1
    'PS_REP02(2) = llave_rep01!FAR_PEDSER
    'PS_REP02(3) = llave_rep01!FAR_PEDFAC
    'PS_REP02(4) = llave_rep01!far_codart
    'llave_rep02.Requery
    f1 = f1 + 1
    xl.Cells(f1, 1) = "'" & llave_rep01!PED_NUMGUIA
    xl.Cells(f1, 2) = llave_rep01!PED_FECHA_EMISION
    xl.Cells(f1, 3) = val(llave_rep01!PED_ROLLOS)
    xl.Cells(f1, 4) = Format(llave_rep01!PED_CANTIDAD, "0.000")
    xl.Cells(f1, 5) = Trim(llave_rep01!art_nombre)
    'xl.Cells(f1, 6) = llave_rep01!FAR_fecha_compra
    'xl.Cells(f1, 7) = "'" & Format(llave_rep01!far_serguia, "000") & "-" & llave_rep01!far_numguia
    PS_REP03(0) = "00"
    PS_REP03(1) = llave_rep01!art_numero
    pu_cp = "P"
    pu_codcia = LK_CODCIA
    pu_codclie = llave_rep01!PED_CODCLIE
    LEER_CLI_LLAVE
    PS_REP03(2) = Trim(cli_llave!cli_ruc_esposo)
    llave_rep03.Requery
    WPRECIO = 0
    CONS_CARDADO = val(cli_llave!CLI_NOMBRE_EMPRESA)
    If Not llave_rep03.EOF Then
       WPRECIO = val(llave_rep03!tab_NOMLARGO)
    End If
    xl.Cells(f1, 8) = Format(WPRECIO, "0.000")
    xl.Cells(f1, 9) = Format(llave_rep01!PED_CANTIDAD * WPRECIO, "0.000")
    'If Trim(llave_rep02!PED_CONTACTO) = "OK" Then
    '   xl.Cells(f1, 10) = CONS_CARDADO
    'End If
 llave_rep01.MoveNext
Loop
  wran1 = "I" & 6
  wran2 = "I" & f1
  wranF = "I" & f1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wran1 = "J" & 6
  wran2 = "J" & f1
  wranF = "J" & f1 + 1
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wran1 = "C" & 6
  wran2 = "C" & f1
  wranF = "C" & f1 + 2
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  wran1 = "D" & 6
  wran2 = "D" & f1
  wranF = "D" & f1 + 2
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  
  xl.Cells(f1 + 1, 6) = "Total US$ ="
  wran1 = "I" & f1 + 1
  wran2 = "J" & f1 + 1
  wranF = "I" & f1 + 2
  xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
  xl.Cells(f1 + 2, 5) = "    Total General US$ ="
  wranF = "A" & f1 + 2 & ":J" & f1 + 2
  xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
  FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  DoEvents
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ""
  xl.APPLICATION.Visible = True
  DoEvents
  lblproceso.Visible = False
  ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  
  Set xl = Nothing
Exit Sub



CANCELA:
  FrmImp3.Pantalla.Enabled = True
  FrmImp3.Pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblproceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  lblproceso.Caption = "Abriendo , Archivo REGVENTA.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open PUB_RUTA_OTRO & "\SERVICIOS2.xls", 0, True, 4, WPAS, WPAS
Return



FINTODO:
 MsgBox Err.Description & " .-  Reintente Nuevamente ..", 48, Pub_Titulo
  Resume Next
 GoTo CANCELA
 Resume Next
End Sub

Public Sub SALDO_CTA()
'On Error GoTo FINTODO
Dim ws_codclie As Currency
Dim WS_MONTO As Currency
Dim wcar As Currency
Dim wcodclie As Currency
Dim CONS_CARDADO As Currency
Dim WPRECIO As Currency
Dim CAPTURA_FILA As Integer
Dim wsFECHA1
Dim wsFECHA2
Dim ws_clave
Dim wsuma_ini As Currency
Dim wsuma_sal As Currency
Dim wsuma_ini_DOL As Currency
Dim wsuma_sal_DOL As Currency
Dim total_ini_DOL As Currency
Dim total_sal_DOL As Currency
Dim total_ini_SOL As Currency
Dim total_sal_SOL As Currency

Dim vdocum As String

CONS_CARDADO = 0
Pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
PUB_CP = Left(CP.Text, 1)
pub_cadena = ""
xcuenta = 0

Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp3.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = PUB_CLAVE
f1 = 6  'Fila Inicial

pub_cadena = "SELECT * FROM CARTERA WHERE CAR_CODCIA = ? AND CAR_CP = ?  AND CAR_FECHA_SUNAT <= ?  ORDER BY  CAR_CODCLIE,  CAR_FECHA_SUNAT"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = 0
PS_REP01(2) = 0
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT CAA_IMPORTE FROM CARACU WHERE CAA_CODCIA = ? AND CAA_CP = ? AND CAA_CODCLIE = ? AND CAA_TIPDOC = ? AND CAA_SERDOC = ? AND CAA_NUMDOC = ? AND CAA_FECHA_COBRO < ? AND CAA_ESTADO <> 'E' ORDER BY CAA_FECHA_COBRO "
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = 0
PS_REP02(1) = 0
PS_REP02(2) = 0
PS_REP02(3) = 0
PS_REP02(4) = 0
PS_REP02(5) = 0
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
DoEvents
FrmImp3.lblproceso.Visible = True
FrmImp3.ProgBar.Visible = True
FrmImp3.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = PUB_CP
PS_REP01(2) = wsFECHA1
llave_rep01.Requery
If llave_rep01.EOF Then
  Pantalla.Enabled = True
  cerrar.Enabled = True
  FrmImp3.lblproceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  MsgBox "No Existe Datos en este Rango..", 48, Pub_Titulo
  Exit Sub
End If
GoSub WEXCEL

FrmImp3.lblproceso.Caption = "Procesando . . . "
DoEvents
FrmImp3.ProgBar.Min = 0
FrmImp3.ProgBar.Value = 0
FrmImp3.ProgBar.max = llave_rep01.RowCount
'wcodclie = llave_rep01!far_codclie
PS_REP02(0) = LK_CODCIA
PS_REP02(1) = 180
xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
If PUB_CP = "C" Then
  xl.Cells(2, 1) = "CUANTAS POR COBRAR AL " & Format(wsFECHA1, "dd/mm/yy")
  xl.Cells(5, 2) = "C L I E N T E S"
Else
  xl.Cells(2, 1) = "CUANTAS POR PÀGAR AL " & Format(wsFECHA1, "dd/mm/yy")
  xl.Cells(5, 2) = "P R O V E E D O R"
End If
f1 = 5
ws_codclie = llave_rep01!CAR_CODCLIE
wsuma_sal = 0
wsuma_ini = 0
wsuma_sal_DOL = 0
wsuma_ini_DOL = 0
total_sal_DOL = 0
total_ini_DOL = 0
Do Until llave_rep01.EOF
    FrmImp3.ProgBar.Value = FrmImp3.ProgBar.Value + 1
    PS_REP02(0) = LK_CODCIA
    PS_REP02(1) = PUB_CP
    PS_REP02(2) = llave_rep01!CAR_CODCLIE
    PS_REP02(3) = llave_rep01!car_TIPDOC
    PS_REP02(4) = llave_rep01!car_SERDOC
    PS_REP02(5) = llave_rep01!car_NUMDOC
    PS_REP02(6) = wsFECHA1
    llave_rep02.Requery
    WS_MONTO = 0
    Do Until llave_rep02.EOF
      WS_MONTO = WS_MONTO + val(llave_rep02!CAA_IMPORTE)
      llave_rep02.MoveNext
    Loop
    If WS_MONTO = 0 Then GoTo dale
    If f1 = 5 Then ws_codclie = llave_rep01!CAR_CODCLIE
    If ws_codclie <> llave_rep01!CAR_CODCLIE Then
       f1 = f1 + 1
       xl.Cells(f1, 4) = "     S/.="
       xl.Cells(f1, 7) = wsuma_ini
       xl.Cells(f1, 8) = wsuma_sal
       wranF = "A" & f1 + 2 & ":H" & f1 + 2
       xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
       f1 = f1 + 1
       xl.Worksheets(1).Rows(f1).RowHeight = 15
       xl.Cells(f1, 4) = "     US$.="
       xl.Cells(f1, 7) = wsuma_ini_DOL
       xl.Cells(f1, 8) = wsuma_sal_DOL
       total_sal_DOL = total_sal_DOL + wsuma_ini_DOL
       total_ini_DOL = total_ini_DOL + wsuma_sal_DOL
       total_sal_SOL = total_sal_SOL + wsuma_ini
       total_ini_SOL = total_ini_SOL + wsuma_sal
       wsuma_sal = 0
       wsuma_ini = 0
       wsuma_sal_DOL = 0
       wsuma_ini_DOL = 0
       ws_codclie = llave_rep01!CAR_CODCLIE
    End If
    
    SQ_OPER = 1
    pu_cp = PUB_CP
    pu_codcia = LK_CODCIA
    pu_codclie = llave_rep01!CAR_CODCLIE
    LEER_CLI_LLAVE
    f1 = f1 + 1
    xl.Cells(f1, 1) = "'" & Format(llave_rep01!CAR_FECHA_SUNAT, "dd/mm/yy")
    xl.Cells(f1, 2) = Left(cli_llave!CLI_NOMBRE, 30)
    xl.Cells(f1, 3) = llave_rep01!car_TIPDOC
    vdocum = ""
    If llave_rep01!CAR_cp = "C" Then
     If Trim(llave_rep01!car_FBG) = "F" Then
        vdocum = "Fact. " & Trim(llave_rep01!car_NUMSER) & " - " & llave_rep01!car_NUMFAC
     ElseIf Trim(llave_rep01!car_FBG) = "B" Then
        vdocum = "Bolet. " & Trim(llave_rep01!car_NUMSER) & " - " & llave_rep01!car_NUMFAC
     ElseIf Trim(llave_rep01!car_FBG) = "G" Then
        vdocum = "Guia . " & Trim(llave_rep01!car_NUMSER) & " - " & Trim(llave_rep01!car_NUMFAC)
     ElseIf Trim(llave_rep01!car_FBG) = "N" Then
        vdocum = "N.Cred. " & llave_rep01!car_NUMSER & "-" & llave_rep01!car_NUMFAC
     ElseIf Trim(llave_rep01!car_FBG) = "D" Then
        vdocum = "N.Deb. " & llave_rep01!car_NUMSER & "-" & llave_rep01!car_NUMFAC
     Else
        vdocum = " "
     End If
     End If
     
     If llave_rep01!CAR_TIPMOV = 97 Then
        vdocum = "N.-Cred. " & llave_rep01!car_NUMSER & "-" & llave_rep01!car_NUMFAC
     End If
     If llave_rep01!CAR_TIPMOV = 98 Then
        vdocum = "N.-Deb. " & llave_rep01!car_NUMSER & "-" & llave_rep01!car_NUMFAC
     End If
    
    If llave_rep01!CAR_cp = "P" Then
       If llave_rep01!CAR_TIPMOV = 98 Then
          vdocum = "N.-Deb. " & llave_rep01!CAR_NUMSER_C & "-" & llave_rep01!CAR_NUMFAC_C
       End If
       If llave_rep01!CAR_TIPMOV = 97 Then
          vdocum = "N.-Cred. " & llave_rep01!CAR_NUMSER_C & "-" & llave_rep01!CAR_NUMFAC_C
       End If
       If llave_rep01!CAR_TIPMOV = 20 Then

          If val(Nulo_Valor0(llave_rep01!CAR_NUMFAC_C)) > 0 Then
             vdocum = "Fact. " & llave_rep01!CAR_NUMSER_C & "-" & llave_rep01!CAR_NUMFAC_C
          Else
             vdocum = "Guia . " & llave_rep01!car_numguia
          End If
        End If
        If llave_rep01!CAR_TIPMOV = 99 Then
           If val(Nulo_Valor0(llave_rep01!CAR_NUMFAC_C)) > 0 Then
                vdocum = "Fact. " & llave_rep01!CAR_NUMSER_C & "-" & llave_rep01!CAR_NUMFAC_C
            Else
               vdocum = "Guia . " & Trim(llave_rep01!car_numguia)
            End If
        End If
        
     Else
        If llave_rep01!CAR_TIPMOV = 10 Then
          If val(Nulo_Valor0(llave_rep01!CAR_NUMFAC_C)) > 0 Then
             vdocum = "Fact. " & llave_rep01!CAR_NUMSER_C & "-" & llave_rep01!CAR_NUMFAC_C
          ElseIf Trim(llave_rep01!car_FBG) <> "F" And Trim(llave_rep01!car_FBG) <> "B" Then
             vdocum = "Guia . " & llave_rep01!car_NUMSER & " - " & llave_rep01!car_NUMFAC
          End If
        End If
     End If
     If llave_rep01!car_TIPDOC = "LE" Then
        vdocum = " " & llave_rep01!car_NUMSER & " - " & llave_rep01!car_NUMFAC
     End If
    
     If llave_rep01!CAR_TIPMOV = 0 Or llave_rep01!car_TIPDOC = "CD" Then
        If val(llave_rep01!car_NUM_CHEQUE) <> 0 Then
          vdocum = "Nº. " & llave_rep01!car_NUM_CHEQUE
         End If
        If llave_rep01!car_TIPDOC = "FA" And llave_rep01!car_FBG <> " " Then
          vdocum = llave_rep01!car_FBG & " / " & llave_rep01!CAR_NUMSER_C & "-" & llave_rep01!CAR_NUMFAC_C
        End If
     End If
    xl.Cells(f1, 4) = vdocum
    xl.Cells(f1, 5) = "'" & Format(llave_rep01!car_fecha_vcto, "dd/mm/yy")
    xl.Cells(f1, 7) = val(llave_rep01!CAR_IMP_INI)
    xl.Cells(f1, 8) = WS_MONTO
    If llave_rep01!CAR_MONEDA = "S" Then
      xl.Cells(f1, 6) = "S/."
      wsuma_ini = wsuma_ini + val(llave_rep01!CAR_IMP_INI)
      wsuma_sal = wsuma_sal + WS_MONTO
    Else
      xl.Cells(f1, 6) = "US$."
      wsuma_ini_DOL = wsuma_ini_DOL + val(llave_rep01!CAR_IMP_INI)
      wsuma_sal_DOL = wsuma_sal_DOL + WS_MONTO
    End If
dale:
 llave_rep01.MoveNext
Loop
f1 = f1 + 1
xl.Cells(f1, 4) = "     S/.="
xl.Cells(f1, 7) = wsuma_ini
xl.Cells(f1, 8) = wsuma_sal
wranF = "A" & f1 + 2 & ":H" & f1 + 2
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
f1 = f1 + 1
xl.Worksheets(1).Rows(f1).RowHeight = 15
xl.Cells(f1, 4) = "     US$.="
xl.Cells(f1, 7) = wsuma_ini_DOL
xl.Cells(f1, 8) = wsuma_sal_DOL
total_sal_DOL = total_sal_DOL + wsuma_ini_DOL
total_ini_DOL = total_ini_DOL + wsuma_sal_DOL
total_sal_SOL = total_sal_SOL + wsuma_ini
total_ini_SOL = total_ini_SOL + wsuma_sal
f1 = f1 + 1
xl.Cells(f1, 4) = "TOTAL GENERAL S/.="
xl.Cells(f1, 7) = total_ini_SOL
xl.Cells(f1, 8) = total_sal_SOL
wranF = "A" & f1 + 2 & ":H" & f1 + 2
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
xl.Worksheets(1).Rows(f1).RowHeight = 15
f1 = f1 + 1
xl.Cells(f1, 4) = "TOTAL GENERAL US$.="
xl.Cells(f1, 7) = total_ini_DOL
xl.Cells(f1, 8) = total_sal_DOL
wranF = "A" & f1 + 2 & ":H" & f1 + 2
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
xl.Worksheets(1).Rows(f1).RowHeight = 15

wsuma_sal = 0
wsuma_ini = 0
wsuma_sal_DOL = 0
wsuma_ini_DOL = 0
     
    
  xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
  FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  DoEvents
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ""
  xl.APPLICATION.Visible = True
  DoEvents
  lblproceso.Visible = False
  ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  
  Set xl = Nothing
Exit Sub



CANCELA:
  FrmImp3.Pantalla.Enabled = True
  FrmImp3.Pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblproceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  lblproceso.Caption = "Abriendo , Archivo SALDO_CTA.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open PUB_RUTA_OTRO & "\SALDO_CTA.xls", 0, True, 4, WPAS, WPAS
Return



FINTODO:
 MsgBox Err.Description & " .-  Reintente Nuevamente ..", 48, Pub_Titulo
  Resume Next
 GoTo CANCELA
 Resume Next
End Sub


Public Sub MERMA()
'On Error GoTo FINTODO
Dim ws_codclie As Currency
Dim WS_MONTO As Currency
Dim wcar As Currency
Dim wcodclie As Currency
Dim CONS_CARDADO As Currency
Dim WPRECIO As Currency
Dim CAPTURA_FILA As Integer
Dim wsFECHA1
Dim wsFECHA2
Dim ws_clave
Dim wsuma_ini As Currency
Dim wsuma_sal As Currency
Dim wsuma_ini_DOL As Currency
Dim wsuma_sal_DOL As Currency
Dim total_ini_DOL As Currency
Dim total_sal_DOL As Currency
Dim total_ini_SOL As Currency
Dim total_sal_SOL As Currency

Dim vdocum As String

CONS_CARDADO = 0
Pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA1 = Left(txtcampo2.Text, 8)
Else
     wsFECHA1 = Trim(txtcampo2.Text)
End If
If Not IsDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If


PUB_CP = Left(CP.Text, 1)
pub_cadena = ""
xcuenta = 0

Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp3.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = PUB_CLAVE
f1 = 6  'Fila Inicial
'**************
pub_cadena = "SELECT * FROM FACART WHERE FAR_CODCIA = ? AND FAR_TIPMOV = ?  AND (FAR_FECHA_COMPRA <= ? AND FAR_FECHA_COMPRA <= ?)   ORDER BY  FAR_CODART"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = 0
PS_REP01(2) = LK_FECHA_DIA
PS_REP01(3) = LK_FECHA_DIA
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

pub_cadena = "SELECT CAA_IMPORTE FROM CARACU WHERE CAA_CODCIA = ? AND CAA_CP = ? AND CAA_CODCLIE = ? AND CAA_TIPDOC = ? AND CAA_SERDOC = ? AND CAA_NUMDOC = ? AND CAA_FECHA_COBRO < ? AND CAA_ESTADO <> 'E' ORDER BY CAA_FECHA_COBRO "
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = 0
PS_REP02(1) = 0
PS_REP02(2) = 0
PS_REP02(3) = 0
PS_REP02(4) = 0
PS_REP02(5) = 0
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
DoEvents
FrmImp3.lblproceso.Visible = True
FrmImp3.ProgBar.Visible = True
FrmImp3.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = PUB_CP
PS_REP01(2) = wsFECHA1
llave_rep01.Requery
If llave_rep01.EOF Then
  Pantalla.Enabled = True
  cerrar.Enabled = True
  FrmImp3.lblproceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  MsgBox "No Existe Datos en este Rango..", 48, Pub_Titulo
  Exit Sub
End If
GoSub WEXCEL

FrmImp3.lblproceso.Caption = "Procesando . . . "
DoEvents
FrmImp3.ProgBar.Min = 0
FrmImp3.ProgBar.Value = 0
FrmImp3.ProgBar.max = llave_rep01.RowCount
'wcodclie = llave_rep01!far_codclie
PS_REP02(0) = LK_CODCIA
PS_REP02(1) = 180
xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
If PUB_CP = "C" Then
  xl.Cells(2, 1) = "CUANTAS POR COBRAR AL " & Format(wsFECHA1, "dd/mm/yy")
  xl.Cells(5, 2) = "C L I E N T E S"
Else
  xl.Cells(2, 1) = "CUANTAS POR PÀGAR AL " & Format(wsFECHA1, "dd/mm/yy")
  xl.Cells(5, 2) = "P R O V E E D O R"
End If
f1 = 5
ws_codclie = llave_rep01!CAR_CODCLIE
wsuma_sal = 0
wsuma_ini = 0
wsuma_sal_DOL = 0
wsuma_ini_DOL = 0
total_sal_DOL = 0
total_ini_DOL = 0
Do Until llave_rep01.EOF
    FrmImp3.ProgBar.Value = FrmImp3.ProgBar.Value + 1
    PS_REP02(0) = LK_CODCIA
    PS_REP02(1) = PUB_CP
    PS_REP02(2) = llave_rep01!CAR_CODCLIE
    PS_REP02(3) = llave_rep01!car_TIPDOC
    PS_REP02(4) = llave_rep01!car_SERDOC
    PS_REP02(5) = llave_rep01!car_NUMDOC
    PS_REP02(6) = wsFECHA1
    llave_rep02.Requery
    WS_MONTO = 0
    Do Until llave_rep02.EOF
      WS_MONTO = WS_MONTO + val(llave_rep02!CAA_IMPORTE)
      llave_rep02.MoveNext
    Loop
    If WS_MONTO = 0 Then GoTo dale
    If f1 = 5 Then ws_codclie = llave_rep01!CAR_CODCLIE
    If ws_codclie <> llave_rep01!CAR_CODCLIE Then
       f1 = f1 + 1
       xl.Cells(f1, 4) = "     S/.="
       xl.Cells(f1, 7) = wsuma_ini
       xl.Cells(f1, 8) = wsuma_sal
       wranF = "A" & f1 + 2 & ":H" & f1 + 2
       xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
       f1 = f1 + 1
       xl.Worksheets(1).Rows(f1).RowHeight = 15
       xl.Cells(f1, 4) = "     US$.="
       xl.Cells(f1, 7) = wsuma_ini_DOL
       xl.Cells(f1, 8) = wsuma_sal_DOL
       total_sal_DOL = total_sal_DOL + wsuma_ini_DOL
       total_ini_DOL = total_ini_DOL + wsuma_sal_DOL
       total_sal_SOL = total_sal_SOL + wsuma_ini
       total_ini_SOL = total_ini_SOL + wsuma_sal
       wsuma_sal = 0
       wsuma_ini = 0
       wsuma_sal_DOL = 0
       wsuma_ini_DOL = 0
       ws_codclie = llave_rep01!CAR_CODCLIE
    End If
    
    SQ_OPER = 1
    pu_cp = PUB_CP
    pu_codcia = LK_CODCIA
    pu_codclie = llave_rep01!CAR_CODCLIE
    LEER_CLI_LLAVE
    f1 = f1 + 1
    xl.Cells(f1, 1) = "'" & Format(llave_rep01!CAR_FECHA_SUNAT, "dd/mm/yy")
    xl.Cells(f1, 2) = Left(cli_llave!CLI_NOMBRE, 30)
    xl.Cells(f1, 3) = llave_rep01!car_TIPDOC
    vdocum = ""
    If llave_rep01!CAR_cp = "C" Then
     If Trim(llave_rep01!car_FBG) = "F" Then
        vdocum = "Fact. " & Trim(llave_rep01!car_NUMSER) & " - " & llave_rep01!car_NUMFAC
     ElseIf Trim(llave_rep01!car_FBG) = "B" Then
        vdocum = "Bolet. " & Trim(llave_rep01!car_NUMSER) & " - " & llave_rep01!car_NUMFAC
     ElseIf Trim(llave_rep01!car_FBG) = "G" Then
        vdocum = "Guia . " & Trim(llave_rep01!car_NUMSER) & " - " & Trim(llave_rep01!car_NUMFAC)
     ElseIf Trim(llave_rep01!car_FBG) = "N" Then
        vdocum = "N.Cred. " & llave_rep01!car_NUMSER & "-" & llave_rep01!car_NUMFAC
     ElseIf Trim(llave_rep01!car_FBG) = "D" Then
        vdocum = "N.Deb. " & llave_rep01!car_NUMSER & "-" & llave_rep01!car_NUMFAC
     Else
        vdocum = " "
     End If
     End If
     
     If llave_rep01!CAR_TIPMOV = 97 Then
        vdocum = "N.-Cred. " & llave_rep01!car_NUMSER & "-" & llave_rep01!car_NUMFAC
     End If
     If llave_rep01!CAR_TIPMOV = 98 Then
        vdocum = "N.-Deb. " & llave_rep01!car_NUMSER & "-" & llave_rep01!car_NUMFAC
     End If
    
    If llave_rep01!CAR_cp = "P" Then
       If llave_rep01!CAR_TIPMOV = 98 Then
          vdocum = "N.-Deb. " & llave_rep01!CAR_NUMSER_C & "-" & llave_rep01!CAR_NUMFAC_C
       End If
       If llave_rep01!CAR_TIPMOV = 97 Then
          vdocum = "N.-Cred. " & llave_rep01!CAR_NUMSER_C & "-" & llave_rep01!CAR_NUMFAC_C
       End If
       If llave_rep01!CAR_TIPMOV = 20 Then

          If val(Nulo_Valor0(llave_rep01!CAR_NUMFAC_C)) > 0 Then
             vdocum = "Fact. " & llave_rep01!CAR_NUMSER_C & "-" & llave_rep01!CAR_NUMFAC_C
          Else
             vdocum = "Guia . " & llave_rep01!car_numguia
          End If
        End If
        If llave_rep01!CAR_TIPMOV = 99 Then
           If val(Nulo_Valor0(llave_rep01!CAR_NUMFAC_C)) > 0 Then
                vdocum = "Fact. " & llave_rep01!CAR_NUMSER_C & "-" & llave_rep01!CAR_NUMFAC_C
            Else
               vdocum = "Guia . " & Trim(llave_rep01!car_numguia)
            End If
        End If
        
     Else
        If llave_rep01!CAR_TIPMOV = 10 Then
          If val(Nulo_Valor0(llave_rep01!CAR_NUMFAC_C)) > 0 Then
             vdocum = "Fact. " & llave_rep01!CAR_NUMSER_C & "-" & llave_rep01!CAR_NUMFAC_C
          ElseIf Trim(llave_rep01!car_FBG) <> "F" And Trim(llave_rep01!car_FBG) <> "B" Then
             vdocum = "Guia . " & llave_rep01!car_NUMSER & " - " & llave_rep01!car_NUMFAC
          End If
        End If
     End If
     If llave_rep01!car_TIPDOC = "LE" Then
        vdocum = " " & llave_rep01!car_NUMSER & " - " & llave_rep01!car_NUMFAC
     End If
    
     If llave_rep01!CAR_TIPMOV = 0 Or llave_rep01!car_TIPDOC = "CD" Then
        If val(llave_rep01!car_NUM_CHEQUE) <> 0 Then
          vdocum = "Nº. " & llave_rep01!car_NUM_CHEQUE
         End If
        If llave_rep01!car_TIPDOC = "FA" And llave_rep01!car_FBG <> " " Then
          vdocum = llave_rep01!car_FBG & " / " & llave_rep01!CAR_NUMSER_C & "-" & llave_rep01!CAR_NUMFAC_C
        End If
     End If
    xl.Cells(f1, 4) = vdocum
    xl.Cells(f1, 5) = "'" & Format(llave_rep01!car_fecha_vcto, "dd/mm/yy")
    xl.Cells(f1, 7) = val(llave_rep01!CAR_IMP_INI)
    xl.Cells(f1, 8) = WS_MONTO
    If llave_rep01!CAR_MONEDA = "S" Then
      xl.Cells(f1, 6) = "S/."
      wsuma_ini = wsuma_ini + val(llave_rep01!CAR_IMP_INI)
      wsuma_sal = wsuma_sal + WS_MONTO
    Else
      xl.Cells(f1, 6) = "US$."
      wsuma_ini_DOL = wsuma_ini_DOL + val(llave_rep01!CAR_IMP_INI)
      wsuma_sal_DOL = wsuma_sal_DOL + WS_MONTO
    End If
dale:
 llave_rep01.MoveNext
Loop
f1 = f1 + 1
xl.Cells(f1, 4) = "     S/.="
xl.Cells(f1, 7) = wsuma_ini
xl.Cells(f1, 8) = wsuma_sal
wranF = "A" & f1 + 2 & ":H" & f1 + 2
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
f1 = f1 + 1
xl.Worksheets(1).Rows(f1).RowHeight = 15
xl.Cells(f1, 4) = "     US$.="
xl.Cells(f1, 7) = wsuma_ini_DOL
xl.Cells(f1, 8) = wsuma_sal_DOL
total_sal_DOL = total_sal_DOL + wsuma_ini_DOL
total_ini_DOL = total_ini_DOL + wsuma_sal_DOL
total_sal_SOL = total_sal_SOL + wsuma_ini
total_ini_SOL = total_ini_SOL + wsuma_sal
f1 = f1 + 1
xl.Cells(f1, 4) = "TOTAL GENERAL S/.="
xl.Cells(f1, 7) = total_ini_SOL
xl.Cells(f1, 8) = total_sal_SOL
wranF = "A" & f1 + 2 & ":H" & f1 + 2
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
xl.Worksheets(1).Rows(f1).RowHeight = 15
f1 = f1 + 1
xl.Cells(f1, 4) = "TOTAL GENERAL US$.="
xl.Cells(f1, 7) = total_ini_DOL
xl.Cells(f1, 8) = total_sal_DOL
wranF = "A" & f1 + 2 & ":H" & f1 + 2
xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
xl.Worksheets(1).Rows(f1).RowHeight = 15

wsuma_sal = 0
wsuma_ini = 0
wsuma_sal_DOL = 0
wsuma_ini_DOL = 0
     
    
  xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
  FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  DoEvents
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ""
  xl.APPLICATION.Visible = True
  DoEvents
  lblproceso.Visible = False
  ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  
  Set xl = Nothing
Exit Sub



CANCELA:
  FrmImp3.Pantalla.Enabled = True
  FrmImp3.Pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblproceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  lblproceso.Caption = "Abriendo , Archivo SALDO_CTA.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = ws_clave
  xl.Workbooks.Open PUB_RUTA_OTRO & "\MERMA.xls", 0, True, 4, WPAS, WPAS
Return



FINTODO:
 MsgBox Err.Description & " .-  Reintente Nuevamente ..", 48, Pub_Titulo
  Resume Next
 GoTo CANCELA
 Resume Next
End Sub


Public Sub REG_COMPRA()
'On Error GoTo FINTODO
Dim wsFECHA1
Dim wsFECHA2
Dim WS_BRUTO As Currency
Dim ws_igv As Currency
Dim ws_total As Currency
Dim ws_bruto_d As Currency
Dim ws_igv_d As Currency
Dim ws_total_d As Currency

WS_BRUTO = 0
ws_igv = 0
ws_total = 0
ws_bruto_d = 0
ws_igv_d = 1
ws_total_d = 0

wsFECHA1 = Trim(txtCampo1.Text)
wsFECHA2 = Trim(txtcampo2.Text)

Pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Not IsDate(wsFECHA1) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If Right(txtcampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtcampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtcampo2.Text)
End If
If Not IsDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If
If CDate(wsFECHA1) > CDate(wsFECHA2) Then
 MsgBox "Fecha Invalidad ..", 48, Pub_Titulo
 GoTo CANCELA
End If


Pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp3.lblproceso.Caption = "Activando Reporte... un Momento ."
DoEvents

f1 = 6  'Fila Inicial
'**************
pub_cadena = "SELECT * FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_TIPMOV = ?  AND (ALL_FECHA_PRO >= ? AND ALL_FECHA_PRO <= ?) AND ALL_FLAG_EXT <> 'E'   ORDER BY  ALL_FECHA_SUNAT"
Set PS_REP01 = CN.CreateQuery("", pub_cadena)
PS_REP01(0) = 0
PS_REP01(1) = 0
PS_REP01(2) = LK_FECHA_DIA
PS_REP01(3) = LK_FECHA_DIA
Set llave_rep01 = PS_REP01.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

DoEvents
FrmImp3.lblproceso.Visible = True
FrmImp3.ProgBar.Visible = True
FrmImp3.lblproceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
PS_REP01(0) = LK_CODCIA
PS_REP01(1) = 99
PS_REP01(2) = wsFECHA1
PS_REP01(3) = wsFECHA2
llave_rep01.Requery
If llave_rep01.EOF Then
  Pantalla.Enabled = True
  cerrar.Enabled = True
  FrmImp3.lblproceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  MsgBox "No Existe Datos en este Rango..", 48, Pub_Titulo
  Exit Sub
End If
GoSub WEXCEL

FrmImp3.lblproceso.Caption = "Procesando . . . "
DoEvents
FrmImp3.ProgBar.Min = 0
FrmImp3.ProgBar.Value = 0
FrmImp3.ProgBar.max = llave_rep01.RowCount
'wcodclie = llave_rep01!far_codclie
'PS_REP02(0) = LK_CODCIA
'PS_REP02(1) = 180
xl.Cells(1, 1) = Trim(Mid(MDIForm1.stb_EB.Panels("cia"), 4, Len(MDIForm1.stb_EB.Panels("cia"))))
xl.Cells(2, 1) = "REGISTRO DE COMPRA DEL " & Format(wsFECHA1, "dd/mm/yy") & "AL " & Format(wsFECHA2, "dd/mm/yy")
f1 = 5
Do Until llave_rep01.EOF
    FrmImp3.ProgBar.Value = FrmImp3.ProgBar.Value + 1
    f1 = f1 + 1
    xl.Cells(f1, 1) = "'" & Format(llave_rep01!ALL_FECHA_SUNAT, "dd/mm")
    xl.Cells(f1, 2) = llave_rep01!ALL_CODSUNAT
    xl.Cells(f1, 3) = Format(llave_rep01!all_numser_c, "000") & "-" & llave_rep01!all_numfac_c
    SQ_OPER = 1
    pu_cp = llave_rep01!ALL_CP
    pu_codcia = LK_CODCIA
    pu_codclie = llave_rep01!ALL_CODCLIE
    LEER_CLI_LLAVE
    xl.Cells(f1, 4) = Left(cli_llave!CLI_NOMBRE, 25)
    xl.Cells(f1, 5) = Trim(cli_llave!cli_ruc_esposo)
    If Trim(llave_rep01!ALL_MONEDA_CLI) = "S" Then
       WS_BRUTO = WS_BRUTO + val(llave_rep01!ALL_BRUTO)
       ws_igv = ws_igv + val(llave_rep01!ALL_IMPTO)
       ws_total = ws_total + val(llave_rep01!ALL_IMPORTE)
       xl.Cells(f1, 6) = "S"
    Else
       ws_bruto_d = ws_bruto_d + val(llave_rep01!ALL_BRUTO)
       ws_igv_d = ws_igv_d + val(llave_rep01!ALL_IMPTO)
       ws_total_d = ws_total_d + val(llave_rep01!ALL_IMPORTE)
       xl.Cells(f1, 6) = "D"
    End If
    xl.Cells(f1, 7) = llave_rep01!ALL_BRUTO
    xl.Cells(f1, 8) = llave_rep01!ALL_IMPTO
    xl.Cells(f1, 9) = llave_rep01!ALL_IMPORTE
    
    xl.Cells(f1, 10) = Trim(cli_llave!CLI_CUENTA_CONTAB)
    xl.Cells(f1, 11) = Trim(llave_rep01!ALL_CTAG1)
    xl.Cells(f1, 12) = llave_rep01!ALL_IMPG1
    xl.Cells(f1, 13) = Trim(llave_rep01!ALL_CTAG2)
    xl.Cells(f1, 14) = llave_rep01!ALL_IMPG2
    xl.Cells(f1, 15) = Format(llave_rep01!all_numfac, "0000")
 
 llave_rep01.MoveNext
Loop
    
    f1 = f1 + 1
    xl.Cells(f1, 6) = "S/."
    xl.Cells(f1, 7) = WS_BRUTO
    xl.Cells(f1, 8) = ws_igv
    xl.Cells(f1, 9) = ws_total
    f1 = f1 + 1
    xl.Cells(f1, 6) = "US$."
    xl.Cells(f1, 7) = ws_bruto_d
    xl.Cells(f1, 8) = ws_igv_d
    xl.Cells(f1, 9) = ws_total_d
    
 ' xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
  FrmImp2.lblproceso.Caption = "Mostrando Hoja de Calculo  . . . "
  DoEvents
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ""
  xl.APPLICATION.Visible = True
  DoEvents
  lblproceso.Visible = False
  ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  
  Set xl = Nothing
Exit Sub



CANCELA:
  FrmImp3.Pantalla.Enabled = True
  FrmImp3.Pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblproceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  Pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.APPLICATION.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim DD As Excel.APPLICATION
  Dim wsfile1
  lblproceso.Caption = "Abriendo , Archivo SALDO_CTA.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  xl.Workbooks.Open PUB_RUTA_OTRO & "\REG_COMPRA.xls", 0, True, 4, PUB_CLAVE, PUB_CLAVE
Return



FINTODO:
 MsgBox Err.Description & " .-  Reintente Nuevamente ..", 48, Pub_Titulo
  Resume Next
 GoTo CANCELA
 Resume Next
End Sub

Public Function JALAR(wfecha As Date) As Currency
PUB_CAL_INI = wfecha
PUB_CAL_FIN = wfecha
pu_codcia = LK_CODCIA
PUB_CODCIA = LK_CODCIA
SQ_OPER = 1
LEER_CAL_LLAVE
If cal_llave.EOF Then
  JALAR = 0
  Exit Function
End If
If IsNull(cal_llave!cal_tipo_cambio) Then
  JALAR = 0
  Exit Function
End If
JALAR = cal_llave!cal_tipo_cambio

End Function

