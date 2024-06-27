VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FrmImp3 
   BackColor       =   &H00FAEFDA&
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
   StartUpPosition =   3  'Windows Default
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
      BackColor       =   &H00FAEFDA&
      Caption         =   "Clientes :"
      Height          =   1815
      Left            =   120
      TabIndex        =   21
      Top             =   1560
      Visible         =   0   'False
      Width           =   8535
      Begin VB.OptionButton op1 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Todas"
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
         Index           =   2
         Left            =   3000
         TabIndex        =   27
         Top             =   1080
         Width           =   975
      End
      Begin VB.OptionButton op1 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Entregadas"
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
         Index           =   1
         Left            =   1560
         TabIndex        =   26
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton op1 
         BackColor       =   &H00FAEFDA&
         Caption         =   "Pendientes"
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
         BackStyle       =   0  'Transparent
         Height          =   735
         Left            =   1800
         TabIndex        =   23
         Top             =   240
         Width           =   3495
      End
   End
   Begin VB.ListBox lisD 
      Height          =   2085
      Left            =   3480
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
         ItemData        =   "FrmImp3.frx":0442
         Left            =   120
         List            =   "FrmImp3.frx":044C
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
      BackColor       =   &H00FAEFDA&
      Height          =   615
      Left            =   420
      TabIndex        =   6
      Top             =   -45
      Width           =   4455
      Begin VB.Label lblreporte 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Name            =   "Tahoma"
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
         Name            =   "Tahoma"
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
      Left            =   1800
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
      Left            =   360
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
   Begin VB.Label LUSUARIO 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   3600
      TabIndex        =   20
      Top             =   600
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.Label lblcampo2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Campo1"
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
      Height          =   195
      Left            =   1920
      TabIndex        =   9
      Top             =   960
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label lblcampo1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Campo1"
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
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   960
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label lblProceso 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
Dim C1 As Integer
Dim F1 As Integer
Dim xcuenta As Integer
Dim i As Integer
Dim Mensaje, titulo, valorpred As String
Dim Wfile  As String
Dim WFORM  As String
Dim REP_FECHA1
Dim REP_FECHA2
Dim PSPRO1 As rdoQuery
Dim pro1_llave As rdoResultset



Private Sub cerrar_Click()
Unload FrmImp3
End Sub



Private Sub fbg_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  Azul serie, serie
End If

End Sub

Private Sub Form_Load()
CenterMe FrmImp3
Screen.MousePointer = 11
If tra_llave.EOF Then
   Screen.MousePointer = 0
   Exit Sub
End If
Screen.MousePointer = 0
Wfile = Trim(tra_llave(3))
WFORM = Trim(tra_llave(7))
lblreporte.Caption = Trim(tra_llave(1))
If Wfile = "CAJA_GRIFO" Then
 txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 txtCampo2.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 lblcampo1.Caption = "Fecha de Inicial : "
 lblcampo1.Visible = True
 txtCampo1.Mask = "##/##/####"
 txtCampo1.Visible = True
 lblcampo2.Caption = "Fecha de Final: "
 lblcampo2.Visible = True
 txtCampo2.Mask = "##/##/####"
 txtCampo2.Visible = True
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
End If
If Wfile = "MOVICONT_FAL" Then
    frmdocu.Visible = True
    fbg.Clear
    PUB_TIPREG = 50
    PUB_CODCIA = "00"
    SQ_OPER = 2
    LEER_TAB_LLAVE
    Do Until tab_mayor.EOF
        fbg.AddItem Format(tab_mayor!TAB_NUMTAB, "00") & " - " & Trim(tab_mayor!tab_nomlargo)
        tab_mayor.MoveNext
    Loop
    fbg.TabIndex = 0
End If


 
If Wfile = "VENTAS_ARTICULO" Then

End If
If Wfile = "CONSO_FAC" Then
 txtCampo1.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 txtCampo2.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 lblcampo1.Caption = "Fecha de Inicial : "
 lblcampo1.Visible = True
 txtCampo1.Mask = "##/##/####"
 txtCampo1.Visible = True
 lblcampo2.Caption = "Fecha de Final: "
 lblcampo2.Visible = True
 txtCampo2.Mask = "##/##/####"
 txtCampo2.Visible = True
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
 txtCampo2.Text = Format(LK_FECHA_DIA, "dd/mm/yyyy")
 lblcampo1.Caption = "Fec.Emision Ini.: "
 lblcampo1.Visible = True
 txtCampo1.Mask = "##/##/####"
 txtCampo1.Visible = True
 lblcampo2.Caption = "Fec.Emision Fin. : "
 lblcampo2.Visible = True
 txtCampo2.Mask = "##/##/####"
 txtCampo2.Visible = True
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
   If Val(webusca) = Val(Mid(lisD.Text, 10, 16)) Then
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
  If pantalla.Enabled And pantalla.Visible Then
     pantalla.SetFocus
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
If Val(txt_cli.Text) = 0 Then Exit Sub
 'PED_FECHA = ? AND PED_FECHA = ?) AND PED_CODCIA = ? AND PED_TIPMOV = ? AND PED_CIERRE <> ? AND PED_CODCLIE = ?  ODER BY PED_CODCIA "
PS_REP01(0) = txtCampo1.Text
PS_REP01(1) = txtCampo2.Text
PS_REP01(2) = LK_CODCIA
PS_REP01(3) = 177
If op1(0).Value Then  ' pendientes
 PS_REP01(4) = " "
ElseIf op1(1).Value Then  ' entregadas
 PS_REP01(4) = "X"
ElseIf op1(2).Value Then  ' todas
 PS_REP01(4) = "A"
End If
PS_REP01(5) = Val(txt_cli.Text)
llave_rep01.Requery
listop.Clear
If Not llave_rep01.EOF Then
 ProgBar.Min = 0
 ProgBar.Max = llave_rep01.RowCount
 ProgBar.Value = 0
End If
ProgBar.Visible = True
DoEvents
Do Until llave_rep01.EOF
 ProgBar.Value = ProgBar.Value + 1
 listop.AddItem Format(llave_rep01!PED_FECHA, "dd/mm/yyyy") & " O/P. " & Format(llave_rep01!PED_NUMSER, "000") & " - " & Format(llave_rep01!PED_NUMFAC, "0000000")
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
If Wfile = "MOVICONT_FAL" Then
 Call MOVICONT_FAL
End If

Exit Sub
SALE:
ProgBar.Visible = False
lblProceso.Visible = False
pantalla.Enabled = True
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
If KeyAscii = 13 And Wfile = "CONSO_FAC" Then
  REP_FECHA1 = txtCampo1.Text
  REP_FECHA2 = txtCampo2.Text
  PROC
  Exit Sub
End If
If txt_cli.Visible Then
 txt_cli.SetFocus
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
    
pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtCampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtCampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtCampo2.Text)
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

pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp3.lblProceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = PUB_CLAVE
F1 = 6  'Fila Inicial

'WCONTROL = WCONTROL + 1
pub_cadena = "SELECT FAR_SUBTOTAL, FAR_CANTIDAD, FAR_EQUIV, FAR_PRECIO, FAR_COSPRO, FAR_CODART, FAR_SIGNO_CAR, FAR_TIPDOC FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND  FAR_CODART = ? AND FAR_TIPMOV = 10 AND FAR_ESTADO = 'N'  ORDER BY FAR_TIPMOV, FAR_FBG DESC , FAR_FECHA, FAR_NUMSER,FAR_NUMFAC"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


' el PS_REP1(0) ESTA MAS ABAJO
DoEvents
FrmImp3.lblProceso.Visible = True
FrmImp3.ProgBar.Visible = True
FrmImp3.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
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
FrmImp3.lblProceso.Caption = "Procesando . . . "
DoEvents
FrmImp3.ProgBar.Min = 0
FrmImp3.ProgBar.Value = 0
FrmImp3.ProgBar.Max = art_mayor.RowCount
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
 wran2 = "C" & F1
 wranF = "C" & F1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 wran1 = "D" & 6
 wran2 = "D" & F1
 wranF = "D" & F1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 TT_TOTAL_VENTAS = TT_TOTAL_VENTAS + Val(xl.Range(wranF))
 TT_TOTAL_CONTADO = TT_TOTAL_CONTADO + Val(xl.Range(wranF))

 wran1 = "E" & 6
 wran2 = "E" & F1
 wranF = "E" & F1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 
 wran1 = "F" & 6
 wran2 = "F" & F1
 wranF = "F" & F1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 TT_TOTAL_VENTAS = TT_TOTAL_VENTAS + Val(xl.Range(wranF))
 TT_TOTAL_CREDITO = TT_TOTAL_CREDITO + Val(xl.Range(wranF))
 
 wran1 = "G" & 6
 wran2 = "G" & F1
 wranF = "G" & F1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 wran1 = "H" & 6
 wran2 = "H" & F1
 wranF = "H" & F1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 TT_TOTAL_VENTAS = TT_TOTAL_VENTAS + Val(xl.Range(wranF))
 TT_TOTAL_AUTOCONSUMO = TT_TOTAL_AUTOCONSUMO + Val(xl.Range(wranF))
 
 wran1 = "I" & 6
 wran2 = "I" & F1
 wranF = "I" & F1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 wran1 = "J" & 6
 wran2 = "J" & F1
 wranF = "J" & F1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
   
 wranF = "A" & F1 + 1 & ":J" & F1 + 1
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
 wranF = "A" & F1 + 2 & ":J" & F1 + 2
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 6
 
pub_cadena = "SELECT ALL_IMPORTE, ALL_CODCLIE FROM ALLOG WHERE ALL_CODCIA = ? AND ALL_FECHA_DIA  >= ? AND ALL_FECHA_DIA <= ? AND  ALL_SIGNO_CAJA = 1 AND ALL_FLAG_EXT <> 'E' AND (ALL_CODTRA = 2770 OR ALL_CODTRA = 2725)  ORDER BY ALL_CODCLIE "
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)

GoSub COBRANZA_CLI
 
 'Reporte de Resumen
 '------------------
    

    F1 = F1 + 4
    CAPTURA_FILA = F1
    xl.Cells(F1, 2) = "RESUMEN"
    F1 = F1 + 1
    xl.Cells(F1, 3) = "S/."
    F1 = F1 + 1
    xl.Cells(F1, 2) = "TOTAL VENTAS"
    xl.Cells(F1, 3) = TT_TOTAL_VENTAS
    F1 = F1 + 1
    TT_TOTAL_CREDITO = TT_TOTAL_CREDITO - TT_VOUCHER
    xl.Cells(F1, 2) = "'(-) VENTAS CREDITO"
    xl.Cells(F1, 3) = TT_TOTAL_CREDITO
    F1 = F1 + 1
    xl.Cells(F1, 2) = "'(-) TARJ. DE CREDITO"
    xl.Cells(F1, 3) = TT_VOUCHER
    F1 = F1 + 1
    xl.Cells(F1, 2) = "'(-) AUTO CONSUMO"
    xl.Cells(F1, 3) = TT_TOTAL_AUTOCONSUMO
    F1 = F1 + 2
    TT_TOTAL_CONTADO = TT_TOTAL_VENTAS - TT_TOTAL_CREDITO - TT_TOTAL_AUTOCONSUMO - TT_VOUCHER
    xl.Cells(F1, 2) = "'TOTAL VTAS. CONTADO"
    xl.Cells(F1, 3) = TT_TOTAL_CONTADO
    F1 = F1 + 3
    xl.Cells(F1, 2) = "LIQUIDACION DE EFECTIVO"
    F1 = F1 + 2
    xl.Cells(F1, 2) = "TOT.VTAS.CONTADO"
    xl.Cells(F1, 3) = TT_TOTAL_CONTADO
    WS_EFECTIVO = WS_EFECTIVO + TT_TOTAL_CONTADO
    F1 = F1 + 1
    xl.Cells(F1, 2) = "COBRANZA CLIENTE"
    xl.Cells(F1, 3) = TT_COBRANZA_CLI
    WS_EFECTIVO = WS_EFECTIVO + TT_COBRANZA_CLI
    F1 = F1 + 1
    xl.Cells(F1, 2) = "'(-) GASTOS VARIOS"
    xl.Cells(F1, 3) = WS_GASTOS
    WS_EFECTIVO = WS_EFECTIVO - WS_GASTOS
    F1 = F1 + 3
    xl.Cells(F1, 2) = "'TOTAL EFECTIVO"
    xl.Cells(F1, 3) = WS_EFECTIVO
 
 ' ESTADO DE GANANCIAS Y PERIDAS
    F1 = CAPTURA_FILA
    xl.Cells(F1, 7) = "ESTADO DE GANANCIAS Y PERDIDAS"
    F1 = F1 + 1
    xl.Cells(F1, 9) = "S/."
    F1 = F1 + 1
    xl.Cells(F1, 7) = "VENTAS NETAS"
    xl.Cells(F1, 9) = TT_TOTAL_VENTAS
    F1 = F1 + 1
    xl.Cells(F1, 7) = "'(-) CTO. DE VENTAS"
    xl.Cells(F1, 9) = TT_COSTO_VENTAS
    TT_TOTAL_BRUTO = TT_TOTAL_VENTAS - TT_COSTO_VENTAS
    F1 = F1 + 1
    xl.Cells(F1, 7) = "'UTILIDAD BRUTA"
    xl.Cells(F1, 9) = TT_TOTAL_BRUTO
    
    F1 = F1 + 2
    xl.Cells(F1, 7) = "'GASTOS"
    xl.Cells(F1, 9) = TT_TOTAL_GASTOS
    F1 = F1 + 3
    xl.Cells(F1, 7) = "UTILIDAD DE OPERACION"
    xl.Cells(F1, 9) = TT_TOTAL_BRUTO - TT_TOTAL_GASTOS
    F1 = F1 + 1
    xl.Cells(F1, 7) = "OTROS INGRESOS"
    xl.Cells(F1, 9) = TT_TOTAL_OTROS
    F1 = F1 + 1
    xl.Cells(F1, 9) = "S/."
    F1 = F1 + 1
    xl.Cells(F1, 7) = "UTLIDAD NETA"
    xl.Cells(F1, 9) = (TT_TOTAL_BRUTO - TT_TOTAL_GASTOS) - TT_TOTAL_OTROS
 
  FrmImp3.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(1, 1) = Trim(Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)))
  xl.Cells(2, 1) = Trim(tra_llave!tra_descripcion)
  xl.Cells(3, 1) = "'" & Format(wsFECHA1, "dd/mm/yyyy") & " al " & Format(wsFECHA2, "dd/mm/yyyy")
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.Application.Visible = True
  DoEvents
  FrmImp3.lblProceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp3.pantalla.Enabled = True
  FrmImp3.cerrar.Enabled = True
  FrmImp3.pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblProceso.Visible = False

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
    If Val(llave_rep02!FAR_equiv) = 0 Then
      MsgBox "Equivalencia 0 Reset "
      GoTo CANCELA
    End If
    wvalor = Val(llave_rep02!FAR_SUBTOTAL)
    wcanti = redondea((Val(llave_rep02!FAR_CANTIDAD)) / Val(llave_rep02!FAR_equiv))
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
    TT_COSTO_VENTAS = TT_COSTO_VENTAS + redondea((Val(llave_rep02!FAR_CANTIDAD) * Val(llave_rep02!FAR_COSPRO)))
  llave_rep02.MoveNext
 Loop
 
 
Return
  
 
Exit Sub

IMPFACART:
       F1 = F1 + 1
       PUB_KEY = PUB_KEY
       SQ_OPER = 1
       pu_codcia = LK_CODCIA
       LEER_ART_LLAVE
       xl.Cells(F1, 2) = art_LLAVE!ART_NOMBRE
       xl.Cells(F1, 3) = suma_c_contado
       xl.Cells(F1, 4) = suma_v_contado
       xl.Cells(F1, 5) = suma_c_credito
       xl.Cells(F1, 6) = suma_v_credito
       xl.Cells(F1, 7) = suma_c_auto
       xl.Cells(F1, 8) = suma_v_auto
       xl.Cells(F1, 9) = suma_c_total
       xl.Cells(F1, 10) = suma_v_total
Return

COBRANZA_CLI:
 PS_REP02(0) = LK_CODCIA
 PS_REP02(1) = wsFECHA1
 PS_REP02(2) = wsFECHA2
 llave_rep02.Requery
 Do Until llave_rep02.EOF
    TT_COBRANZA_CLI = TT_COBRANZA_CLI + Val(llave_rep02!ALL_IMPORTE)
  llave_rep02.MoveNext
 Loop


Return

CANCELA:
  FrmImp3.pantalla.Enabled = True
  FrmImp3.pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblProceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.Application.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim dd As Excel.Application
  Dim wsfile1
  lblProceso.Caption = "Abriendo , Archivo REGVENTA.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\GRIFOS\CAJAGEN.xls", 0, True, 4, WPAS, WPAS
Return



FINTODO:
 MsgBox Err.Description & " .-  Reintente Nuevamente ..", 48, Pub_Titulo
  Resume Next
 GoTo CANCELA
 Resume Next
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
Dim Wflag As String * 1
If Trim(fbg.Text) = "" Then
  MsgBox "Ingrese Tipo de Documento.", 48, Pub_Titulo
  Exit Sub
End If
If Val(serie.Text) <= 0 Then
   MsgBox "Ingrese Serie del Documento.", 48, Pub_Titulo
   Azul serie, serie
   Exit Sub
End If
If Val(numini.Text) <= 0 Then
   MsgBox "Ingrese N° inicial del Documento.", 48, Pub_Titulo
   Azul numini, numini
   Exit Sub
End If
If Val(numfin.Text) <= 0 Then
   MsgBox "Ingrese N° Final del Documento.", 48, Pub_Titulo
   Azul numfin, numfin
   Exit Sub
End If

    
pantalla.Enabled = False
cerrar.Enabled = False
pub_cadena = ""
xcuenta = 0
pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp3.lblProceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = PUB_CLAVE
F1 = 9  'Fila Inicial
Wflag = ""

'WCONTROL = WCONTROL + 1
pub_cadena = "SELECT Distinct FAR_NUMFAC, FAR_FECHA_COMPRA FROM FACART WHERE FAR_CODCIA = ? AND FAR_FBG = ? AND FAR_NUMSER = ? AND FAR_NUMFAC >= ? AND FAR_NUMFAC <= ? AND FAR_TIPMOV = 10  ORDER BY  FAR_NUMFAC"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
PS_REP02(0) = 0
PS_REP02(1) = 0
PS_REP02(2) = 0
PS_REP02(3) = 0
PS_REP02(4) = 0
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)
PS_REP02(0) = LK_CODCIA
PS_REP02(1) = fbg.Text
PS_REP02(2) = serie.Text
PS_REP02(3) = numini.Text
PS_REP02(4) = numfin.Text
llave_rep02.Requery
If llave_rep02.EOF Then
   MsgBox "En este rango no Existe documentos ", 48, Pub_Titulo
   GoTo CANCELA
End If

' el PS_REP1(0) ESTA MAS ABAJO
DoEvents
FrmImp3.lblProceso.Visible = True
FrmImp3.ProgBar.Visible = True
FrmImp3.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
FrmImp3.lblProceso.Caption = "Procesando . . . "
DoEvents
FrmImp3.ProgBar.Min = 0
FrmImp3.ProgBar.Value = 0

 
 FrmImp3.ProgBar.Max = llave_rep02.RowCount
 If Trim(fbg.Text) = "F" Then
   xl.Cells(5, 2) = "FACTURAS"
 Else
   xl.Cells(5, 2) = "BOLETAS"
 End If
 WW_CORRELA = Val(numini.Text)
 ip_numfac = Val(numini.Text)
 ip_fecha = Format(llave_rep02!FAR_fecha_compra, "dd/mm/yyyy")
 tfechaini = Format(llave_rep02!FAR_fecha_compra, "dd/mm/yyyy")
 wwnumfac = llave_rep02!far_numfac
 xl.Cells(6, 2) = serie.Text
 xl.Cells(7, 2) = numini.Text
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
      Wflag = "A"
    End If
    WW_CORRELA = llave_rep02!far_numfac
    WW_CORRELA = WW_CORRELA + 1
    numfin.Text = llave_rep02!far_numfac
    tfechafin = Format(llave_rep02!FAR_fecha_compra, "dd/mm/yyyy")
    llave_rep02.MoveNext
 Loop
 wranF = "A" & F1 + 1 & ":D" & F1 + 1
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
 F1 = F1 + 2
 xl.Cells(F1, 1) = "FIN DE LISTADO "
 If Wflag <> "A" Then
   F1 = F1 + 1
   xl.Cells(F1, 1) = "*** LOS DOC. ESTAN CORRECTOS *** "
 End If
 F1 = F1 + 1
 xl.Cells(F1, 1) = "N° FINAL:"
 xl.Cells(F1, 2) = numfin.Text
 F1 = F1 + 1
 xl.Cells(F1, 1) = "FEC.FINAL:"
 xl.Cells(F1, 2) = "'" & tfechafin
 wranF = "A" & F1 + 1 & ":D" & F1 + 1
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
 
  FrmImp3.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
  

  xl.Cells(1, 1) = Trim(Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)))
  xl.Cells(2, 1) = "'" & Format(LK_FECHA_DIA, "dd/mm/yyyy")
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.Application.Visible = True
  DoEvents
  FrmImp3.lblProceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp3.pantalla.Enabled = True
  FrmImp3.cerrar.Enabled = True
  FrmImp3.pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblProceso.Visible = False

Exit Sub

IMP_NUMFAC:
       F1 = F1 + 1
       If Trim(fbg.Text) = "F" Then
           xl.Cells(F1, 1) = "FACTURAS"
       Else
           xl.Cells(F1, 1) = "BOLETAS"
       End If
       xl.Cells(F1, 2) = ip_numfac
       xl.Cells(F1, 3) = "FECHA PROBABLE:"
       xl.Cells(F1, 4) = "'" + ip_fecha
Return


CANCELA:
  FrmImp3.pantalla.Enabled = True
  FrmImp3.pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblProceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.Application.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim dd As Excel.Application
  Dim wsfile1
  lblProceso.Caption = "Abriendo , Archivo NUMDOC.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open PUB_RUTA_REPORTE + "NUMDOC.xls", 0, True, 4, WPAS, WPAS
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
Dim wdocu
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
pantalla.Enabled = False
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

pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp3.lblProceso.Caption = "Activando Reporte... un Momento ."
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
FrmImp3.lblProceso.Visible = True
FrmImp3.ProgBar.Visible = True
FrmImp3.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
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
FrmImp3.lblProceso.Caption = "Procesando . . . "
DoEvents
FrmImp3.ProgBar.Min = 0
FrmImp3.ProgBar.Value = 0
FrmImp3.ProgBar.Max = art_mayor.RowCount
GoSub LETRAS
C1 = 1
F001 = 0
CAPTURA_FILA = 0
xl.Cells(7, 1) = "SALDO INICIAL "
F001 = 0
TOT_VENTA = 0
Do Until art_mayor.EOF
 If art_mayor!ART_KEY = 0 Then GoTo OTRO_ARTI
 If art_mayor!art_familia <> 1 Then GoTo OTRO_ARTI
 C1 = C1 + 2
 F1 = 6  'Fila Inicial
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
    SALDO_INI = (llave_rep03!FAR_STOCK / llave_rep03!FAR_equiv) + (llave_rep03!far_SIGNO_aRM * (llave_rep03!FAR_CANTIDAD / llave_rep03!FAR_equiv) * -1)
 Else
    SQ_OPER = 1
    PUB_CODART = art_mayor!ART_KEY
    pu_codcia = LK_CODCIA
    LEER_ARM_LLAVE
    SALDO_INI = arm_llave!arm_stock
 End If
 xl.Cells(F1, C1) = art_mayor!ART_NOMBRE
 xl.Cells(F1, C1 - 1) = "ISLA"
 
 F1 = F1 + 1
 xl.Cells(F1, C1) = SALDO_INI
 ACU_SALDO_ARTI = ACU_SALDO_ARTI + SALDO_INI
 wranF = "A" & F1 & ":" & Trim(LETRAS(C1)) & F1
 xl.Range(wranF).Borders.LineStyle = 1
 wranF = "A" & F1 - 1 & ":" & Trim(LETRAS(C1)) & F1 - 1
 xl.Range(wranF).Borders.LineStyle = 1
 
 xl.Cells(F1, C1).Font.Bold = True
 
 PS_REP02(0) = LK_CODCIA
 PS_REP02(1) = wsFECHA1
 PS_REP02(2) = art_mayor!ART_KEY
 PS_REP02(3) = 20
 PS_REP02(4) = 20
 llave_rep02.Requery
 If Not llave_rep02.EOF Then
   F1 = F1 + 1
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
 F1 = F1 + 1
 If F001 = 0 Then F001 = F1
 xl.Cells(F001 + 1, C1) = ACU_INGRESOS
 xl.Cells(F001 + 1, 1) = "TOTAL INGRESOS "
 xl.Cells(F001 + 1, 1).Font.Bold = True
 xl.Cells(F001 + 1, C1).Font.Bold = True
 wranF = "A" & F001 + 1 & ":" & Trim(LETRAS(C1)) & F001 + 1
 xl.Range(wranF).Borders.LineStyle = 1
 xl.Worksheets(1).Rows(F001 + 1).RowHeight = 18
 F1 = F001
 TOT_VENTA = 0
 Do Until tab_mayor.EOF
    max_islas = llave_rep04.RowCount + F1 + 2
    wdocu = "TURNO: " & tab_mayor!TAB_NUMTAB
    F1 = F1 + 1
    GoSub PROC_TURNOS
    wturno = tab_mayor!TAB_NUMTAB
    tab_mayor.MoveNext
 Loop
 xl.Cells(max_islas + 1, C1) = TOT_VENTA
 xl.Cells(max_islas + 1, C1).Font.Bold = True
 xl.Cells(max_islas + 1, 1) = "TOTAL VENTA "
 xl.Cells(max_islas + 1, 1).Font.Bold = True
 wranF = "A" & max_islas + 1 & ":" & Trim(LETRAS(C1)) & max_islas + 1
 xl.Range(wranF).Borders.LineStyle = 1
 
 xl.Worksheets(1).Rows(max_islas + 1).RowHeight = 18
 xl.Cells(max_islas + 2, C1) = ACU_SALDO_ARTI
 xl.Cells(max_islas + 2, C1).Font.Bold = True
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
  WSTCKR = stock_llave!tab_nomlargo
 Else
  WSTCKR = 0
 End If
 WDIFE = ACU_SALDO_ARTI - WSTCKR
 art_mayor.Edit
 art_mayor!art_margen = Val(WDIFE)
 art_mayor.Update
 xl.Cells(max_islas + 3, C1) = WSTCKR
 xl.Cells(max_islas + 3, C1).Font.Bold = True
 xl.Cells(max_islas + 3, 1) = "STOCK POR REGLA "
 xl.Cells(max_islas + 3, 1).Font.Bold = True
 xl.Cells(max_islas + 4, C1) = WDIFE
 xl.Cells(max_islas + 4, C1).Font.Bold = True
 xl.Cells(max_islas + 4, 1) = "DIFERENCIA: "
 xl.Cells(max_islas + 4, 1).Font.Bold = True
 
 
 wranF = "A" & max_islas + 2 & ":" & Trim(LETRAS(C1)) & max_islas + 2
 xl.Range(wranF).Borders.LineStyle = 1
 
 xl.Worksheets(1).Rows(max_islas + 3).RowHeight = 18
 
 If F1 > CAPTURA_FILA Then CAPTURA_FILA = F1
OTRO_ARTI:
 art_mayor.MoveNext
Loop
  wranF = "A6:" & Trim(LETRAS(C1)) & 6
  xl.Range(wranF).Borders.Item(3).LineStyle = 1
  wranF = "A6:A" & F1 + 3
  xl.Range(wranF).Borders.Item(7).LineStyle = 1
  wranF = Trim(LETRAS(C1)) & 6 & ":" & Trim(LETRAS(C1)) & F1 + 3
  xl.Range(wranF).Borders.Item(10).LineStyle = 1
 

  FrmImp3.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(1, 1) = Trim(Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)))
  xl.Cells(2, 1) = Trim(tra_llave!tra_descripcion)
  xl.Cells(3, 1) = "'DEL " & Format(wsFECHA1, "dd/mm/yyyy") & " al " & Format(LK_FECHA_DIA, "dd/mm/yyyy")
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.Application.Visible = True
  DoEvents
  FrmImp3.lblProceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp3.pantalla.Enabled = True
  FrmImp3.cerrar.Enabled = True
  FrmImp3.pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblProceso.Visible = False

Exit Sub

PROC_INGRESOS:
WSALDO = 0
d_numfac = llave_rep02!far_numfac
ff_flag = ""
wdocu = ""
Do Until llave_rep02.EOF
    If d_numfac <> llave_rep02!far_numfac Then
       F1 = F1 + 1
       xl.Cells(F1, 1) = wdocu
       xl.Cells(F1, C1) = WSALDO
       d_numfac = llave_rep02!far_numfac
       ACU_INGRESOS = ACU_INGRESOS + WSALDO
       ACU_SALDO_ARTI = ACU_SALDO_ARTI + WSALDO
       WSALDO = 0
    End If
    WSALDO = WSALDO + (Val(llave_rep02!FAR_CANTIDAD) / Val(llave_rep02!FAR_equiv))
    If Val(llave_rep02!far_numfac_c) <> 0 Then
      wdocu = "FAC.: " & llave_rep02!far_numser_c & " - " & llave_rep02!far_numfac_c
    Else
     wdocu = "GUIA.: " & llave_rep02!far_NUMGUIA
    End If
    ff_flag = "A"
    llave_rep02.MoveNext
Loop

If ff_flag = "A" Then
   F1 = F1 + 1
   xl.Cells(F1, 1) = wdocu
   xl.Cells(F1, C1) = WSALDO
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
If fila_islas = 0 Then fila_islas = F1
PS_REP01(5) = tab_mayor!TAB_NUMTAB
llave_rep01.Requery
If Not llave_rep01.EOF Then
 d_numfac = llave_rep01!far_ISLA
End If
Do Until llave_rep01.EOF
   If (d_numfac <> llave_rep01!far_ISLA) Then
       F1 = F1 + 1
       xl.Cells(F1, C1 - 1) = WISLA
       xl.Cells(F1, C1) = WSALDO
       d_numfac = llave_rep01!far_numfac
       sum_turnos = sum_turnos + WSALDO
       WSALDO = 0
       wconta = wconta + 1
   End If
   WISLA = llave_rep01!far_ISLA
   WSALDO = WSALDO + (Val(llave_rep01!FAR_CANTIDAD) / Val(llave_rep01!FAR_equiv))
   ff_flag = "A"
   llave_rep01.MoveNext
Loop
If ff_flag = "A" Then
   F1 = F1 + 1
   xl.Cells(F1, C1 - 1) = WISLA
   xl.Cells(F1, C1) = WSALDO
   sum_turnos = sum_turnos + WSALDO
   wconta = wconta + 1
End If
xl.Cells(max_islas, 1) = wdocu
xl.Cells(max_islas, C1) = sum_turnos
wranF = "A" & max_islas & ":" & Trim(LETRAS(C1)) & max_islas
xl.Range(wranF).Borders.LineStyle = 1

TOT_VENTA = TOT_VENTA + sum_turnos
ACU_SALDO_ARTI = ACU_SALDO_ARTI - sum_turnos
If wconta <> llave_rep04.RowCount Then
    F1 = F1 + (llave_rep04.RowCount - wconta)
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
  FrmImp3.pantalla.Enabled = True
  FrmImp3.pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblProceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.Application.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim dd As Excel.Application
  Dim wsfile1
  lblProceso.Caption = "Abriendo , Archivo REGVENTA.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = PUB_CLAVE
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
pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp3.lblProceso.Caption = "Activando Reporte... un Momento ."
DoEvents
F1 = 5  'Fila Inicial

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
      xl.Cells(4, WS_FILA) = Trim(art_mayor!ART_NOMBRE)
   End If
   art_mayor.MoveNext
Loop
    wranF = "A4:" & LETRAS(WS_FILA) & "4"
    xl.Range(wranF).Borders.LineStyle = 9
    xl.Application.Range(wranF).Select
    
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

FrmImp3.ProgBar.Max = llave_rep02.RowCount
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
  
  WQ_cANTIDAD = Val(llave_rep02!FAR_CANTIDAD) + WQ_cANTIDAD
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
   xl.Cells(i, J) = ww_cant - Val(xl.Cells(i, J))

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


  FrmImp3.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(2, 1) = Trim(Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)))
  xl.Cells(2, 5) = Format(LK_FECHA_DIA, "dddd,dd-mmm-yy")
  xl.DisplayAlerts = False
  xl.Application.Visible = True
  DoEvents
  FrmImp3.lblProceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp3.pantalla.Enabled = True
  FrmImp3.cerrar.Enabled = True
  FrmImp3.pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblProceso.Visible = False
Exit Sub

 

CANCELA:
  FrmImp3.pantalla.Enabled = True
  FrmImp3.pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblProceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.Application.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim dd As Excel.Application
  Dim wsfile1
  lblProceso.Caption = "Abriendo , Archivo REGVENTA.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = ""
  xl.Workbooks.Open CONS_ADMIN & "HERTISA\POR_FACTURAR.xls", 0, True
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
    
pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
     wsFECHA1 = Left(txtCampo1.Text, 8)
Else
     wsFECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtCampo2.Text, 2) = "__" Then
     wsFECHA2 = Left(txtCampo2.Text, 8)
Else
     wsFECHA2 = Trim(txtCampo2.Text)
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

pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp3.lblProceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = PUB_CLAVE
F1 = 6  'Fila Inicial

'WCONTROL = WCONTROL + 1
pub_cadena = "SELECT FAR_CANTIDAD, FAR_EQUIV, FAR_PRECIO, FAR_COSPRO, FAR_CODART, FAR_SIGNO_CAR, FAR_TIPDOC FROM FACART WHERE FAR_CODCIA = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND  FAR_CODART = ? AND FAR_TIPMOV = 10 AND FAR_ESTADO = 'N'  ORDER BY FAR_TIPMOV, FAR_FBG DESC , FAR_FECHA, FAR_NUMSER,FAR_NUMFAC"
Set PS_REP02 = CN.CreateQuery("", pub_cadena)
Set llave_rep02 = PS_REP02.OpenResultset(rdOpenKeyset, rdConcurReadOnly)


' el PS_REP1(0) ESTA MAS ABAJO
DoEvents
FrmImp3.lblProceso.Visible = True
FrmImp3.ProgBar.Visible = True
FrmImp3.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
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
FrmImp3.lblProceso.Caption = "Procesando . . . "
DoEvents
FrmImp3.ProgBar.Min = 0
FrmImp3.ProgBar.Value = 0
FrmImp3.ProgBar.Max = art_mayor.RowCount
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
 wran2 = "C" & F1
 wranF = "C" & F1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 wran1 = "D" & 6
 wran2 = "D" & F1
 wranF = "D" & F1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 TT_TOTAL_VENTAS = TT_TOTAL_VENTAS + Val(xl.Range(wranF))
 TT_TOTAL_CONTADO = TT_TOTAL_CONTADO + Val(xl.Range(wranF))

 wran1 = "E" & 6
 wran2 = "E" & F1
 wranF = "E" & F1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 
 wran1 = "F" & 6
 wran2 = "F" & F1
 wranF = "F" & F1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 TT_TOTAL_VENTAS = TT_TOTAL_VENTAS + Val(xl.Range(wranF))
 TT_TOTAL_CREDITO = TT_TOTAL_CREDITO + Val(xl.Range(wranF))
 
 wran1 = "G" & 6
 wran2 = "G" & F1
 wranF = "G" & F1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 wran1 = "H" & 6
 wran2 = "H" & F1
 wranF = "H" & F1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 TT_TOTAL_VENTAS = TT_TOTAL_VENTAS + Val(xl.Range(wranF))
 TT_TOTAL_AUTOCONSUMO = TT_TOTAL_AUTOCONSUMO + Val(xl.Range(wranF))
 
 wran1 = "I" & 6
 wran2 = "I" & F1
 wranF = "I" & F1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
 wran1 = "J" & 6
 wran2 = "J" & F1
 wranF = "J" & F1 + 1
 xl.Range(wranF).Formula = "=SUM(" & wran1 & ":" & wran2 & ")"
   
 wranF = "A" & F1 + 1 & ":J" & F1 + 1
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
 wranF = "A" & F1 + 2 & ":J" & F1 + 2
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 6
 
  FrmImp3.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
  xl.Cells(1, 1) = Trim(Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)))
  xl.Cells(2, 1) = Trim(tra_llave!tra_descripcion)
  xl.Cells(3, 1) = "'" & Format(wsFECHA1, "dd/mm/yyyy") & " al " & Format(wsFECHA2, "dd/mm/yyyy")
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.Application.Visible = True
  DoEvents
  FrmImp3.lblProceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp3.pantalla.Enabled = True
  FrmImp3.cerrar.Enabled = True
  FrmImp3.pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblProceso.Visible = False

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
    If Val(llave_rep02!FAR_equiv) = 0 Then
      MsgBox "Equivalencia 0 Reset "
      GoTo CANCELA
    End If
    wvalor = Val(llave_rep02!far_PRECIO) * (Val(llave_rep02!FAR_CANTIDAD) / Val(llave_rep02!FAR_equiv))
    wcanti = (Val(llave_rep02!FAR_CANTIDAD)) / Val(llave_rep02!FAR_equiv)
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
       F1 = F1 + 1
       PUB_KEY = PUB_KEY
       SQ_OPER = 1
       pu_codcia = LK_CODCIA
       LEER_ART_LLAVE
       xl.Cells(F1, 2) = art_LLAVE!ART_NOMBRE
       xl.Cells(F1, 3) = suma_c_contado
       xl.Cells(F1, 4) = suma_v_contado
       xl.Cells(F1, 5) = suma_c_credito
       xl.Cells(F1, 6) = suma_v_credito
       xl.Cells(F1, 7) = suma_c_auto
       xl.Cells(F1, 8) = suma_v_auto
       xl.Cells(F1, 9) = suma_c_total
       xl.Cells(F1, 10) = suma_v_total
Return

COBRANZA_CLI:
 PS_REP02(0) = LK_CODCIA
 PS_REP02(1) = wsFECHA1
 PS_REP02(2) = wsFECHA2
 llave_rep02.Requery
 Do Until llave_rep02.EOF
    TT_COBRANZA_CLI = TT_COBRANZA_CLI + Val(llave_rep02!ALL_IMPORTE)
  llave_rep02.MoveNext
 Loop


Return

CANCELA:
  FrmImp3.pantalla.Enabled = True
  FrmImp3.pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblProceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.Application.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim dd As Excel.Application
  Dim wsfile1
  lblProceso.Caption = "Abriendo , Archivo REGVENTA.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open Left(PUB_RUTA_OTRO, 2) + "\ADMIN\GRIFOS\CAJAGEN.xls", 0, True, 4, WPAS, WPAS
Return



FINTODO:
 MsgBox Err.Description & " .-  Reintente Nuevamente ..", 48, Pub_Titulo
  Resume Next
 GoTo CANCELA
 Resume Next
End Sub


Public Sub CONSO_FAC()
Dim CADENITA
Dim wkSELECT
Dim Wche
Dim WNUMSER
  Dim wnumfac
  Dim WFBG
pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
      REP_FECHA1 = Left(txtCampo1.Text, 8)
Else
     REP_FECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtCampo2.Text, 2) = "__" Then
     REP_FECHA2 = Left(txtCampo2.Text, 8)
Else
     REP_FECHA2 = Trim(txtCampo2.Text)
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
Reportes.WindowTitle = "Reporte :  " & Trim(tra_llave(1))
Reportes.Destination = crptToWindow
Reportes.WindowLeft = 2
Reportes.WindowTop = 70
Reportes.WindowWidth = 635
Reportes.WindowHeight = 390

CADENITA = ""
wkSELECT = ""
Wche = 0
For fila = 0 To lisD.ListCount - 1
  lisD.ListIndex = fila
  WNUMSER = Str(Val(Mid(lisD.Text, 4, 3)))
  wnumfac = Str(Val(Mid(lisD.Text, 10, 16)))
  WFBG = Left(lisD.Text, 1)
  If lisD.Selected(fila) And Trim(lisD.Text) <> "" Then
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
    wkSELECT = "(" + wkSELECT + ") AND ( {FACART.FAR_CODCIA} = '" & LK_CODCIA & "' AND {FACART.FAR_TIPMOV} = 10 AND {FACART.FAR_ESTADO} <> 'E')"
End If

CADENITA = wkSELECT
Reportes.Formulas(0) = ""
Reportes.Formulas(1) = ""
Reportes.Formulas(2) = ""
Reportes.Formulas(3) = ""
Reportes.ReportFileName = PUB_RUTA_REPORTE + "hguia.rpt"
Reportes.SelectionFormula = CADENITA
Reportes.Action = 1
Reportes.ReportFileName = PUB_RUTA_REPORTE + "hguia1.rpt"
Reportes.Action = 1
Reportes.ReportFileName = PUB_RUTA_REPORTE + "hguia2.rpt"
Reportes.Action = 1
PROC_EXCEL

pantalla.Enabled = True
cerrar.Enabled = True



Exit Sub
CANCELA:
 ' REPORTES
 
End Sub

Public Sub PROC()
Dim WFBG
Dim WNUMSER
Dim wnumfac
pub_cadena = "SELECT far_numfac, far_numser, far_fbg FROM facart WHERE FAR_CODCIA = ? AND FAR_FECHA >= ? AND FAR_FECHA <= ? AND FAR_ESTADO <> 'E' AND FAR_TIPMOV = 10 AND FAR_CODUSU = ? ORDER BY  FAR_FECHA, FAR_NUMOPER"
Set PSPRO1 = CN.CreateQuery("", pub_cadena)
PSPRO1(0) = 0
PSPRO1(1) = 0
PSPRO1(2) = 0
PSPRO1(3) = 0
Set pro1_llave = PSPRO1.OpenResultset(rdOpenKeyset, rdConcurValues)
PSPRO1(0) = LK_CODCIA
PSPRO1(1) = REP_FECHA1
PSPRO1(2) = REP_FECHA2
PSPRO1(3) = LK_CODUSU
pro1_llave.Requery
If pro1_llave.EOF Then
  Exit Sub
End If
WFBG = pro1_llave!far_fbg
WNUMSER = pro1_llave!far_numser
wnumfac = pro1_llave!far_numfac
lisD.Clear
lisD.AddItem Trim(pro1_llave!far_fbg) & "/ " & Format(pro1_llave!far_numser, "000") & " - " & Format(pro1_llave!far_numfac, "0000000")
fila = 0
Do Until pro1_llave.EOF
  If Trim(WFBG) = Trim(pro1_llave!far_fbg) And Trim(WNUMSER) = Trim(pro1_llave!far_numser) And Val(wnumfac) = Val(pro1_llave!far_numfac) Then
  Else
    fila = fila + 1
    lisD.AddItem Trim(pro1_llave!far_fbg) & "/ " & Format(pro1_llave!far_numser, "000") & " - " & Format(pro1_llave!far_numfac, "0000000")
'    lisD.Selected(fila) = True
  End If
  
  
  WFBG = pro1_llave!far_fbg
  WNUMSER = pro1_llave!far_numser
  wnumfac = pro1_llave!far_numfac
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
F1 = 6
WMAX_FILA = 26
WCOL_MAX = 4
C1 = 4
fila = 0
For fila = 0 To lisD.ListCount - 1
 lisD.ListIndex = fila
 xl.Application.Visible = True
 If lisD.Selected(fila) Then
   WNUMSER = Val(Mid(lisD.Text, 4, 3))
   wnumfac = Val(Mid(lisD.Text, 10, 16))
   WFBG = Left(lisD.Text, 1)
   If F1 >= WMAX_FILA + 6 Then
     F1 = 6
     C1 = C1 + 1
   End If
   F1 = F1 + 1
   xl.Cells(F1, C1) = WFBG + "." + Str(WNUMSER) + "-" + Str(wnumfac)
 End If
Next fila
xl.DisplayAlerts = False
xl.Worksheets(1).Protect PUB_CLAVE
xl.Application.Visible = True
Set xl = Nothing
Screen.MousePointer = 0
Exit Sub

WEXCEL:
  Dim dd As Excel.Application
  Dim wsfile1
  lblProceso.Caption = "Abriendo , Archivo DOCU.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open PUB_RUTA_REPORTE + "DOCU.xls", 0, True, 4, WPAS, WPAS
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
  If loc_key > ListView2.ListItems.Count Then loc_key = ListView2.ListItems.Count
  GoTo POSICION
End If
If KeyCode = 38 Then
  loc_key = loc_key - 1
  If loc_key < 1 Then loc_key = 1
  GoTo POSICION
End If
If KeyCode = 34 Then
 loc_key = loc_key + 17
 If loc_key > ListView2.ListItems.Count Then loc_key = ListView2.ListItems.Count
 GoTo POSICION
End If
If KeyCode = 33 Then
 loc_key = loc_key - 17
 If loc_key < 1 Then loc_key = 1
 GoTo POSICION
End If
GoTo fin
POSICION:
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
pu_codclie = Val(txt_cli.Text)
On Error GoTo 0
If Len(txt_cli.Text) = 0 Then
   Exit Sub
End If

If pu_codclie <> 0 And IsNumeric(txt_cli.Text) = True Then
   SQ_OPER = 1
   pu_cp = "C"
   pu_codcia = LK_CODCIA
   LEER_CLI_LLAVE
   If cli_llave.EOF Then
     lblCliente.Caption = ""
     MsgBox "REGISTRO NO EXISTE ...", 48, Pub_Titulo
     Azul txt_cli, txt_cli
     GoTo fin
   Else
     lblCliente.Caption = Trim(cli_llave!cli_nombre)
   End If
   op1_Click 0
   op1(0).SetFocus
   'If pantalla.Visible And pantalla.Enabled Then
   '  pantalla.SetFocus
   'End If
Else
   If loc_key > ListView2.ListItems.Count Or loc_key = 0 Then
     Exit Sub
   End If
   valor = UCase(ListView2.ListItems.Item(loc_key).Text)
   If Trim(UCase(txt_cli.Text)) = Left(valor, Len(Trim(txt_cli.Text))) Then
   Else
      Exit Sub
   End If
   lblCliente.Caption = Trim(ListView2.ListItems.Item(loc_key).Text)
   txt_cli.Text = Trim(ListView2.ListItems.Item(loc_key).SubItems(1))
   op1_Click 0
   op1(0).SetFocus
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
Dim var
If Len(txt_cli.Text) = 0 Or IsNumeric(txt_cli.Text) = True Then
   ListView2.Visible = False
   Exit Sub
End If
If ListView2.Visible = False And KeyCode <> 13 Then
    var = Asc(txt_cli.Text)
    var = var + 1
    If var = 33 Or var = 91 Then
       var = "ZZZZZZZZ"
    ElseIf var = 58 Then
       var = "A"
    Else
       var = Chr(var)
    End If
    numarchi = 1
    archi = "SELECT CLI_CODCLIE, CLI_CODCIA, CLI_CP, CLI_NOMBRE,CLI_CASA_DIREC,CLI_ZONA_NEW, CLI_CASA_NUM  FROM CLIENTES WHERE  CLI_CP = 'C' AND CLI_CODCIA = '" & LK_CODCIA & "' AND CLI_NOMBRE BETWEEN '" & txt_cli.Text & "' AND  '" & var & "' ORDER BY CLI_NOMBRE"
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
   If loc_key + 8 > ListView2.ListItems.Count Then
      ListView2.ListItems.Item(ListView2.ListItems.Count).EnsureVisible
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
Dim ano
Dim ANO1
Dim Wflag As String * 1
Dim CADENITA
Dim wkSELECT
Dim Wche
Dim WNUMSER
  Dim wnumfac
  Dim WFBG
pantalla.Enabled = False
cerrar.Enabled = False
If Right(txtCampo1.Text, 2) = "__" Then
      REP_FECHA1 = Left(txtCampo1.Text, 8)
Else
     REP_FECHA1 = Trim(txtCampo1.Text)
End If
If Right(txtCampo2.Text, 2) = "__" Then
     REP_FECHA2 = Left(txtCampo2.Text, 8)
Else
     REP_FECHA2 = Trim(txtCampo2.Text)
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
Reportes.WindowTitle = "Reporte :  " & Trim(tra_llave(1))
Reportes.Destination = crptToWindow
Reportes.WindowLeft = 2
Reportes.WindowTop = 70
Reportes.WindowWidth = 635
Reportes.WindowHeight = 390
DIA = Day(REP_FECHA1)
MES = Month(REP_FECHA1)
ano = Year(REP_FECHA1)
DIA1 = Day(REP_FECHA2)
MES1 = Month(REP_FECHA2)
ANO1 = Year(REP_FECHA2)

 
If op1(0).Value Then
    Wflag = " "
ElseIf op1(1).Value Then
    Wflag = "X"
ElseIf op1(1).Value Then
    Wflag = "A"
End If

CADENITA = ""
wkSELECT = ""
Wche = 0
For fila = 0 To listop.ListCount - 1
  listop.ListIndex = fila
  WNUMSER = Str(Val(Mid(listop.Text, 17, 3)))
  wnumfac = Str(Val(Mid(listop.Text, 23, 30)))
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
Reportes.Formulas(0) = "CIA=  '" & Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)) & "'"
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
pantalla.Enabled = True
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
Dim Wflag As String * 1
If Trim(fbg.Text) = "" Then
  MsgBox "Ingrese Tipo de Documento.", 48, Pub_Titulo
  Exit Sub
End If
If Val(serie.Text) <= 0 Then
   MsgBox "Ingrese Serie del Documento.", 48, Pub_Titulo
   Azul serie, serie
   Exit Sub
End If
If Val(numini.Text) <= 0 Then
   MsgBox "Ingrese N° inicial del Documento.", 48, Pub_Titulo
   Azul numini, numini
   Exit Sub
End If
If Val(numfin.Text) <= 0 Then
   MsgBox "Ingrese N° Final del Documento.", 48, Pub_Titulo
   Azul numfin, numfin
   Exit Sub
End If

    
pantalla.Enabled = False
cerrar.Enabled = False
pub_cadena = ""
xcuenta = 0
pantalla.Enabled = False
cerrar.Enabled = False
DoEvents
FrmImp3.lblProceso.Caption = "Activando Reporte... un Momento ."
DoEvents
ws_clave = PUB_CLAVE
F1 = 9  'Fila Inicial
Wflag = ""

'WCONTROL = WCONTROL + 1
pub_cadena = "SELECT Distinct MOV_NUMFAC, MOV_FECHA_EMI FROM MOVICONT WHERE MOV_CODCIA = ? AND MOV_FBG = ? AND MOV_SERIE = ? AND MOV_NUMFAC >= ? AND MOV_NUMFAC <= ? AND MOV_TIPMOV = ? AND MOV_NRO_MES = " & LK_NRO_MES & " ORDER BY  MOV_NUMFAC"
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
PS_REP02(5) = "02"
llave_rep02.Requery
If llave_rep02.EOF Then
   MsgBox "En este rango no Existe documentos ", 48, Pub_Titulo
   GoTo CANCELA
End If

' el PS_REP1(0) ESTA MAS ABAJO
DoEvents
FrmImp3.lblProceso.Visible = True
FrmImp3.ProgBar.Visible = True
FrmImp3.lblProceso.Caption = "Abriendo Microsoft Excel . . . "
DoEvents
GoSub WEXCEL
FrmImp3.lblProceso.Caption = "Procesando . . . "
DoEvents
FrmImp3.ProgBar.Min = 0
FrmImp3.ProgBar.Value = 0

 
 FrmImp3.ProgBar.Max = llave_rep02.RowCount
 If Left(fbg.Text, 2) = "01" Then
   xl.Cells(5, 2) = "FACTURAS"
 Else
   xl.Cells(5, 2) = "BOLETAS"
 End If
 WW_CORRELA = Val(numini.Text)
 ip_numfac = Val(numini.Text)
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
      Wflag = "A"
    End If
    WW_CORRELA = llave_rep02!MOV_numfac
    WW_CORRELA = WW_CORRELA + 1
    numfin.Text = llave_rep02!MOV_numfac
    tfechafin = Format(llave_rep02!MOV_fecha_EMI, "dd/mm/yyyy")
    llave_rep02.MoveNext
 Loop
 wranF = "A" & F1 + 1 & ":D" & F1 + 1
 xl.Worksheets("Hoja1").Range(wranF).Borders.Item(xlEdgeTop).LineStyle = 1
 F1 = F1 + 2
 xl.Cells(F1, 1) = "FIN DE LISTADO "
 If Wflag <> "A" Then
   F1 = F1 + 1
   xl.Cells(F1, 1) = "*** LOS DOC. ESTAN CORRECTOS *** "
 End If
 F1 = F1 + 1
 xl.Cells(F1, 1) = "N° FINAL:"
 xl.Cells(F1, 2) = numfin.Text
 F1 = F1 + 1
 xl.Cells(F1, 1) = "FEC.FINAL:"
 xl.Cells(F1, 2) = "'" & tfechafin
 wranF = "A" & F1 + 1 & ":D" & F1 + 1
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
 
  FrmImp3.lblProceso.Caption = "Mostrando Hoja de Calculo  . . . "
  

  xl.Cells(1, 1) = Trim(Mid(MDIForm1.TXTCIA.Caption, 4, Len(MDIForm1.TXTCIA.Caption)))
  xl.Cells(2, 1) = "'" & Format(LK_FECHA_DIA, "dd/mm/yyyy")
  xl.DisplayAlerts = False
  xl.Worksheets(1).Protect ws_clave
  xl.Application.Visible = True
  DoEvents
  FrmImp3.lblProceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  Set xl = Nothing
  Screen.MousePointer = 0
  FrmImp3.pantalla.Enabled = True
  FrmImp3.cerrar.Enabled = True
  FrmImp3.pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblProceso.Visible = False

Exit Sub

IMP_NUMFAC:
       F1 = F1 + 1
       If Trim(fbg.Text) = "F" Then
           xl.Cells(F1, 1) = "FACTURAS"
       Else
           xl.Cells(F1, 1) = "BOLETAS"
       End If
       xl.Cells(F1, 2) = ip_numfac
       xl.Cells(F1, 3) = "FECHA PROBABLE:"
       xl.Cells(F1, 4) = "'" + ip_fecha
Return


CANCELA:
  FrmImp3.pantalla.Enabled = True
  FrmImp3.pantalla.Caption = "Por &Pantalla"
  FrmImp3.lblProceso.Visible = False
  FrmImp3.ProgBar.Visible = False
  pantalla.Enabled = True
  cerrar.Enabled = True
  If xl Is Nothing Then
  Else
   xl.Application.Visible = True
  End If
   Set xl = Nothing
  Screen.MousePointer = 0
Exit Sub
WEXCEL:
  Dim dd As Excel.Application
  Dim wsfile1
  lblProceso.Caption = "Abriendo , Archivo NUMDOC.xls . . . "
  If xl Is Nothing Then
     Set xl = CreateObject("Excel.Application")
  End If
  DoEvents
  WPAS = PUB_CLAVE
  xl.Workbooks.Open PUB_RUTA_REPORTE + "NUMDOC.xls", 0, True, 4, WPAS, WPAS
Return



FINTODO:
 MsgBox Err.Description & " .-  Reintente Nuevamente ..", 48, Pub_Titulo
  Resume Next
 GoTo CANCELA
 Resume Next
End Sub

